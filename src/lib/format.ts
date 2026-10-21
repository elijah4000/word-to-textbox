import { LAYOUT_CLASS_PREFIX } from "./convert";
import { normalizeText } from "./textCheck";

/**
 * Turns mammoth's HTML into paste-ready HTML with inline styles (Brightspace's
 * editor drops stylesheets and classes, so every style must be inline).
 *
 * Words are never edited. The only text this step changes is:
 * - text mammoth itself added: footnote markers ("[1]" → "1") and back-links ("↑");
 * - tab characters at the start of a paragraph, which become a 0.5in text-indent
 *   (HTML would otherwise collapse them to nothing).
 * The text check runs on this function's output.
 */

export interface FormatOptions {
  doubleSpacing: boolean;
  hangingIndent: boolean;
  /** CSS font-family value, or "" to leave unset. */
  fontFamily: string;
  /** CSS font-size value such as "12pt", or "" to leave unset. */
  fontSize: string;
  embedImages: boolean;
}

export const DEFAULT_OPTIONS: FormatOptions = {
  doubleSpacing: false,
  hangingIndent: true,
  fontFamily: "",
  fontSize: "",
  embedImages: true,
};

export interface FormatResult {
  html: string;
  images: number;
  /** Approximate decoded size of embedded images, in bytes. */
  imageBytes: number;
  /** Paragraphs with a hanging indent in the output (from Word or added for Works Cited). */
  hangingParagraphs: number;
  /** Works Cited / References paragraphs that got a hanging indent added. */
  addedHanging: number;
}

const BIBLIOGRAPHY_TITLES =
  /^(works cited|works consulted|references|reference list|bibliography|sources cited|sources)$/i;

const SAFE_HREF = /^(https?:|mailto:|#)/i;

const TEXT_BLOCKS = "p, li, td, th, h1, h2, h3, h4, h5, h6, blockquote";

/** Word's default tab stop and the indent added for Works Cited entries. */
const HALF_INCH_PT = 36;

function setStyles(el: HTMLElement, styles: Record<string, string>) {
  for (const [prop, value] of Object.entries(styles)) el.style.setProperty(prop, value);
}

function isHeading(el: Element): boolean {
  return /^H[1-6]$/.test(el.tagName);
}

function fixNotes(root: HTMLElement) {
  // Footnote/endnote reference markers: <sup><a href="#footnote-1" id="footnote-ref-1">[1]</a></sup>
  root.querySelectorAll<HTMLAnchorElement>('sup > a[href^="#footnote-"], sup > a[href^="#endnote-"]').forEach((a) => {
    const sup = a.parentElement!;
    sup.classList.add("note-ref");
    sup.textContent = (a.textContent ?? "").replace(/^\[(\d+)\]$/, "$1");
  });

  // Back-links mammoth appends to each note: " <a href="#footnote-ref-1">↑</a>"
  root.querySelectorAll('a[href^="#footnote-ref-"], a[href^="#endnote-ref-"]').forEach((a) => {
    const prev = a.previousSibling;
    if (prev && prev.nodeType === 3 && prev.nodeValue?.endsWith(" ")) {
      prev.nodeValue = prev.nodeValue.slice(0, -1);
    }
    a.remove();
  });

  // Mammoth always emits the notes <ol> and comments <dl> after the body, even when empty.
  root.querySelectorAll(":scope > ol, :scope > dl").forEach((list) => {
    const isNotes = list.tagName === "OL" && list.querySelector(':scope > li[id^="footnote-"], :scope > li[id^="endnote-"]');
    if (isNotes) {
      list.classList.add("notes");
      root.insertBefore(root.ownerDocument.createElement("hr"), list);
    } else if (list.children.length === 0 && normalizeText(list.textContent ?? "") === "") {
      list.remove();
    }
  });
}

function sanitize(root: HTMLElement) {
  root.querySelectorAll("script, style, iframe, object, embed").forEach((el) => el.remove());
  root.querySelectorAll("*").forEach((el) => {
    for (const attr of [...el.attributes]) {
      if (attr.name.startsWith("on")) el.removeAttribute(attr.name);
    }
  });
  root.querySelectorAll("a[href]").forEach((a) => {
    if (!SAFE_HREF.test(a.getAttribute("href") ?? "")) a.removeAttribute("href");
  });
  root.querySelectorAll("img").forEach((img) => {
    if (!/^data:image\//i.test(img.getAttribute("src") ?? "")) img.remove();
  });
}

function applyLayouts(root: HTMLElement, layoutStyles: Map<string, Record<string, string>>) {
  root.querySelectorAll<HTMLElement>(`[class*="${LAYOUT_CLASS_PREFIX}"]`).forEach((el) => {
    for (const cls of [...el.classList]) {
      if (!cls.startsWith(LAYOUT_CLASS_PREFIX)) continue;
      setStyles(el, layoutStyles.get(cls) ?? {});
      el.classList.remove(cls);
    }
    if (el.classList.length === 0) el.removeAttribute("class");
  });
}

function ptValue(css: string): number {
  return css.endsWith("pt") ? parseFloat(css) || 0 : 0;
}

/** A block with no other text blocks inside it (a td holding a p is not a leaf). */
function isLeafBlock(el: Element): boolean {
  return el.querySelector(TEXT_BLOCKS) === null;
}

function firstTextNode(el: Node): Text | null {
  for (const child of el.childNodes) {
    if (child.nodeType === 3) {
      if ((child.nodeValue ?? "") !== "") return child as Text;
    } else if (child.nodeType === 1) {
      const found = firstTextNode(child);
      if (found) return found;
    }
  }
  return null;
}

/**
 * Keeps the spacing people type: Tab at the start of a paragraph, tabs and runs
 * of spaces inside it, and blank paragraphs. Browsers collapse all of these.
 */
function preserveWhitespace(root: HTMLElement) {
  root.querySelectorAll<HTMLElement>(TEXT_BLOCKS).forEach((block) => {
    if (!isLeafBlock(block)) return;

    const first = firstTextNode(block);
    const leadingTabs = first ? /^\t+/.exec(first.nodeValue ?? "") : null;
    if (first && leadingTabs) {
      first.nodeValue = (first.nodeValue ?? "").slice(leadingTabs[0].length);
      const indent = ptValue(block.style.textIndent) + leadingTabs[0].length * HALF_INCH_PT;
      setStyles(block, { "text-indent": `${indent}pt` });
    }

    const text = block.textContent ?? "";
    if (/\t| {2}|^ /.test(text)) setStyles(block, { "white-space": "pre-wrap" });

    // An empty paragraph is a blank line in Word; an empty <p> has no height.
    if (text === "" && !block.querySelector("img, br") && block.tagName === "P") {
      block.append(root.ownerDocument.createElement("br"));
    }
  });
}

function hangingTargets(root: HTMLElement): HTMLElement[] {
  // Every paragraph after a "Works Cited"/"References"/... title, up to the next heading.
  const targets = new Set<HTMLElement>();
  const children = [...root.children] as HTMLElement[];
  for (let i = 0; i < children.length; i++) {
    const el = children[i];
    if (!(isHeading(el) || el.tagName === "P")) continue;
    if (!BIBLIOGRAPHY_TITLES.test(normalizeText(el.textContent ?? "").replace(/[:.]$/, ""))) continue;
    for (let j = i + 1; j < children.length; j++) {
      const next = children[j];
      if (isHeading(next) || next.tagName === "HR" || next.classList.contains("notes")) break;
      const blank = normalizeText(next.textContent ?? "") === "";
      // Paragraphs that already have their own first-line or hanging indent are left alone.
      if (next.tagName === "P" && !blank && !next.style.textIndent) targets.add(next);
    }
  }
  return [...targets];
}

export function formatHtml(
  mammothHtml: string,
  options: FormatOptions,
  layoutStyles: Map<string, Record<string, string>> = new Map(),
): FormatResult {
  const doc = new DOMParser().parseFromString(`<body><div id="root">${mammothHtml}</div></body>`, "text/html");
  const root = doc.getElementById("root") as HTMLElement;

  sanitize(root);
  fixNotes(root);
  applyLayouts(root, layoutStyles);
  preserveWhitespace(root);

  // Images
  let images = 0;
  let imageBytes = 0;
  root.querySelectorAll<HTMLImageElement>("img").forEach((img) => {
    if (!options.embedImages) {
      img.remove();
      return;
    }
    images++;
    const src = img.getAttribute("src") ?? "";
    const base64 = src.slice(src.indexOf(",") + 1);
    imageBytes += Math.floor((base64.length * 3) / 4);
    setStyles(img, { "max-width": "100%", height: "auto" });
  });

  // Hanging indent for Works Cited entries that don't have one in Word
  let addedHanging = 0;
  if (options.hangingIndent) {
    for (const p of hangingTargets(root)) {
      setStyles(p, { "padding-left": `${HALF_INCH_PT}pt`, "text-indent": `-${HALF_INCH_PT}pt` });
      addedHanging++;
    }
  }
  const hangingParagraphs = [...root.querySelectorAll<HTMLElement>(TEXT_BLOCKS)].filter((el) =>
    el.style.textIndent.startsWith("-"),
  ).length;

  // Tables need inline borders or they are invisible once pasted.
  root.querySelectorAll<HTMLElement>("table").forEach((t) => setStyles(t, { "border-collapse": "collapse" }));
  root.querySelectorAll<HTMLElement>("td, th").forEach((c) =>
    setStyles(c, { border: "1px solid #000000", padding: "4px 8px", "vertical-align": "top" }),
  );

  root.querySelectorAll<HTMLElement>(TEXT_BLOCKS).forEach((el) => {
    if (options.doubleSpacing) setStyles(el, { "line-height": "2" });
    if (options.fontFamily) setStyles(el, { "font-family": options.fontFamily });
    if (options.fontSize && !isHeading(el)) setStyles(el, { "font-size": options.fontSize });
  });
  root.querySelectorAll<HTMLElement>("ol.notes > li").forEach((li) => {
    if (options.fontFamily) setStyles(li, { "font-family": options.fontFamily });
  });

  return { html: root.innerHTML, images, imageBytes, hangingParagraphs, addedHanging };
}
