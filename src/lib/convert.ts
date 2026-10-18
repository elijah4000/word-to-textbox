import mammoth from "mammoth";

/**
 * Converts a .docx with mammoth. The wording is never touched here: mammoth
 * copies text runs verbatim and HTML-escapes them. Formatting options are
 * applied later (see format.ts) so toggling them never re-reads the file.
 */

export const HANGING_STYLE = "__docx2bs_hanging";
export const HANGING_CLASS = "docx2bs-hanging";

export interface ConversionResult {
  /** HTML straight from mammoth (images embedded as data URIs). */
  html: string;
  /** mammoth.extractRawText of the document body: the "original" side of the text check. */
  rawText: string;
  /** Raw text of each referenced footnote/endnote, keyed like the HTML ids ("footnote-1"). */
  notesText: Map<string, string>;
  /** Paragraphs that have a hanging indent in the Word file. */
  hangingParagraphs: number;
  /** Warnings and errors reported by mammoth. */
  messages: string[];
}

// Mirrors mammoth's own raw-text conversion (lib/raw-text.js) so notes are
// extracted the same way the body is.
interface DocElement {
  type: string;
  value?: string;
  children?: DocElement[];
  [key: string]: unknown;
}

function rawTextOf(element: DocElement): string {
  if (element.type === "text") return element.value ?? "";
  if (element.type === "tab") return "\t";
  const tail = element.type === "paragraph" ? "\n\n" : "";
  return (element.children ?? []).map(rawTextOf).join("") + tail;
}

function walk(element: DocElement, visit: (el: DocElement) => void) {
  visit(element);
  for (const child of element.children ?? []) walk(child, visit);
}

function toTwips(value: unknown): number {
  const n = typeof value === "string" ? parseInt(value, 10) : NaN;
  return Number.isFinite(n) ? n : 0;
}

function input(data: ArrayBuffer) {
  // The browser build of mammoth reads `arrayBuffer`; the Node build (used by
  // the tests) reads `buffer`. Both accept the same bytes.
  return { arrayBuffer: data, buffer: data } as unknown as { arrayBuffer: ArrayBuffer };
}

export async function convertDocx(data: ArrayBuffer): Promise<ConversionResult> {
  const notesText = new Map<string, string>();
  let hangingParagraphs = 0;

  const transformDocument = (document: DocElement & { notes: { resolve(ref: unknown): DocElement | null } }) => {
    walk(document, (el) => {
      if (el.type === "paragraph") {
        const indent = el.indent as { hanging: string | null } | undefined;
        // List items and headings keep their own mapping; only plain paragraphs are marked.
        const isHeading = /^(heading|title)/i.test(String(el.styleName ?? ""));
        if (indent && toTwips(indent.hanging) > 0 && !el.numbering && !isHeading) {
          el.styleName = HANGING_STYLE;
          hangingParagraphs++;
        }
      }
      if (el.type === "noteReference") {
        const note = document.notes.resolve(el);
        if (note) {
          notesText.set(`${String(note.noteType)}-${String(note.noteId)}`, (note.body as unknown as DocElement[]).map(rawTextOf).join(""));
        }
      }
    });
    return document;
  };

  const html = await mammoth.convertToHtml(input(data), {
    styleMap: [`p[style-name='${HANGING_STYLE}'] => p.${HANGING_CLASS}:fresh`],
    transformDocument: transformDocument as (element: unknown) => unknown,
    convertImage: mammoth.images.dataUri,
    externalFileAccess: false,
  });
  const raw = await mammoth.extractRawText(input(data));

  const messages = [...html.messages, ...raw.messages].map((m) => `${m.type}: ${m.message}`);
  return {
    html: html.value,
    rawText: raw.value,
    notesText,
    hangingParagraphs,
    messages: [...new Set(messages)],
  };
}
