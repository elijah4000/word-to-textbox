/**
 * Text check: proves the converted HTML contains exactly the words of the
 * original document.
 *
 * Both sides are normalized with the same function, which only collapses runs
 * of whitespace (spaces, tabs, line breaks, non-breaking spaces) into a single
 * space. Nothing else is folded: quotes, dashes, case, accents and punctuation
 * must match character for character.
 */

export function normalizeText(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

// ---------------------------------------------------------------------------
// Extracting text from the final HTML
// ---------------------------------------------------------------------------

const BLOCK_TAGS = new Set([
  "P", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "UL", "OL", "TABLE", "THEAD",
  "TBODY", "TR", "TD", "TH", "DIV", "BLOCKQUOTE", "PRE", "DL", "DT", "DD", "HR",
]);

export interface HtmlText {
  /** Main document text (footnote markers and the notes list excluded). */
  body: string;
  /** Text of each footnote/endnote, keyed by the note's id (e.g. "footnote-1"). */
  notes: Map<string, string>;
}

export interface ExtractOptions {
  /**
   * How a <br> is represented. Mammoth's raw text emits nothing for a line
   * break, so the text check mirrors that ("") to compare like with like.
   * The plain-text clipboard copy uses "\n".
   */
  lineBreak?: string;
  /** Separator written around block elements. */
  blockSeparator?: string;
}

function isNoteRef(el: Element): boolean {
  return el.tagName === "SUP" && el.classList.contains("note-ref");
}

function isNotesList(el: Element): boolean {
  return el.tagName === "OL" && el.classList.contains("notes");
}

export function textOfNode(node: Node, options: ExtractOptions = {}): string {
  const lineBreak = options.lineBreak ?? "";
  const sep = options.blockSeparator ?? "\n";
  let out = "";
  const walk = (n: Node) => {
    if (n.nodeType === 3 /* text */) {
      out += n.nodeValue ?? "";
      return;
    }
    if (n.nodeType !== 1 /* element */) return;
    const el = n as Element;
    if (isNoteRef(el) || isNotesList(el)) return;
    if (el.tagName === "BR") {
      out += lineBreak;
      return;
    }
    const block = BLOCK_TAGS.has(el.tagName);
    if (block) out += sep;
    el.childNodes.forEach(walk);
    if (block) out += sep;
  };
  node.childNodes.forEach(walk);
  return out;
}

export function extractHtmlText(html: string, options: ExtractOptions = {}): HtmlText {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  const body = textOfNode(doc.body, options);
  const notes = new Map<string, string>();
  doc.querySelectorAll("ol.notes > li").forEach((li) => {
    notes.set(li.id, textOfNode(li, options));
  });
  return { body, notes };
}

// ---------------------------------------------------------------------------
// Word-level diff
// ---------------------------------------------------------------------------

export type DiffOp = { type: "equal" | "removed" | "added"; words: string[] };

/** Above this many cells the LCS table is skipped and the changed middle is
 * reported as one removed block plus one added block. */
const MAX_LCS_CELLS = 4_000_000;

function tokens(normalized: string): string[] {
  return normalized === "" ? [] : normalized.split(" ");
}

function push(ops: DiffOp[], type: DiffOp["type"], words: string[]) {
  if (words.length === 0) return;
  const last = ops[ops.length - 1];
  if (last && last.type === type) last.words.push(...words);
  else ops.push({ type, words: [...words] });
}

export function diffWords(a: string[], b: string[]): DiffOp[] {
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) start++;
  let endA = a.length;
  let endB = b.length;
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
    endA--;
    endB--;
  }

  const ops: DiffOp[] = [];
  push(ops, "equal", a.slice(0, start));

  const midA = a.slice(start, endA);
  const midB = b.slice(start, endB);
  const n = midA.length;
  const m = midB.length;

  if (n > 0 || m > 0) {
    if ((n + 1) * (m + 1) > MAX_LCS_CELLS) {
      push(ops, "removed", midA);
      push(ops, "added", midB);
    } else {
      // lcs[i][j] = LCS length of midA[i:] and midB[j:]
      const w = m + 1;
      const lcs = new Uint32Array((n + 1) * w);
      for (let i = n - 1; i >= 0; i--) {
        for (let j = m - 1; j >= 0; j--) {
          lcs[i * w + j] =
            midA[i] === midB[j]
              ? lcs[(i + 1) * w + j + 1] + 1
              : Math.max(lcs[(i + 1) * w + j], lcs[i * w + j + 1]);
        }
      }
      let i = 0;
      let j = 0;
      while (i < n && j < m) {
        if (midA[i] === midB[j]) {
          push(ops, "equal", [midA[i]]);
          i++;
          j++;
        } else if (lcs[(i + 1) * w + j] >= lcs[i * w + j + 1]) {
          push(ops, "removed", [midA[i]]);
          i++;
        } else {
          push(ops, "added", [midB[j]]);
          j++;
        }
      }
      push(ops, "removed", midA.slice(i));
      push(ops, "added", midB.slice(j));
    }
  }

  push(ops, "equal", a.slice(endA));
  return ops;
}

// ---------------------------------------------------------------------------
// Comparison
// ---------------------------------------------------------------------------

export interface TextComparison {
  identical: boolean;
  wordCount: number;
  ops: DiffOp[];
}

export function compareTexts(original: string, converted: string): TextComparison {
  const a = normalizeText(original);
  const b = normalizeText(converted);
  const wa = tokens(a);
  if (a === b) {
    return { identical: true, wordCount: wa.length, ops: wa.length ? [{ type: "equal", words: wa }] : [] };
  }
  return { identical: false, wordCount: wa.length, ops: diffWords(wa, tokens(b)) };
}

export interface NoteComparison {
  id: string;
  comparison: TextComparison;
}

export interface TextCheckResult {
  identical: boolean;
  body: TextComparison;
  notes: NoteComparison[];
  /** Notes present in the original but missing from the HTML, or vice versa. */
  missingNotes: string[];
  extraNotes: string[];
}

/**
 * @param originalBody  mammoth raw text of the document body
 * @param originalNotes raw text of each footnote/endnote, keyed like "footnote-1"
 * @param finalHtml     the exact HTML that will be copied
 */
export function runTextCheck(
  originalBody: string,
  originalNotes: Map<string, string>,
  finalHtml: string,
): TextCheckResult {
  const extracted = extractHtmlText(finalHtml);
  const body = compareTexts(originalBody, extracted.body);
  const notes: NoteComparison[] = [];
  const missingNotes: string[] = [];
  for (const [id, text] of originalNotes) {
    const converted = extracted.notes.get(id);
    if (converted === undefined) missingNotes.push(id);
    else notes.push({ id, comparison: compareTexts(text, converted) });
  }
  const extraNotes = [...extracted.notes.keys()].filter((id) => !originalNotes.has(id));
  const identical =
    body.identical &&
    missingNotes.length === 0 &&
    extraNotes.length === 0 &&
    notes.every((n) => n.comparison.identical);
  return { identical, body, notes, missingNotes, extraNotes };
}
