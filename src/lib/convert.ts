import JSZip from "jszip";
import mammoth from "mammoth";
import { layoutToCss, parseStyleSheet, readIndent, resolveLayout } from "./paragraphStyles";

/**
 * Converts a .docx with mammoth. The wording is never touched here: mammoth
 * copies text runs verbatim and HTML-escapes them. Formatting options are
 * applied later (see format.ts) so toggling them never re-reads the file.
 */

/** Class prefix that carries a paragraph's Word layout from here to format.ts. */
export const LAYOUT_CLASS_PREFIX = "d2b-layout-";

export interface ConversionResult {
  /** HTML straight from mammoth (images embedded as data URIs). */
  html: string;
  /** mammoth.extractRawText of the document body: the "original" side of the text check. */
  rawText: string;
  /** Raw text of each referenced footnote/endnote, keyed like the HTML ids ("footnote-1"). */
  notesText: Map<string, string>;
  /** Inline styles (indent, alignment) for each layout class used in `html`. */
  layoutStyles: Map<string, Record<string, string>>;
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

/** The HTML tag mammoth's default style map gives a paragraph style. */
function defaultTag(styleId: unknown, styleName: unknown): string {
  const name = String(styleName ?? "");
  const id = String(styleId ?? "");
  const m = /^heading ?([1-6])$/i.exec(name) ?? /^Heading([1-6])$/.exec(id);
  if (m) return `h${m[1]}`;
  if (/^heading$/i.test(name) || id === "Heading") return "h1";
  return "p";
}

function input(data: ArrayBuffer) {
  // The browser build of mammoth reads `arrayBuffer`; the Node build (used by
  // the tests) reads `buffer`. Both accept the same bytes.
  return { arrayBuffer: data, buffer: data } as unknown as { arrayBuffer: ArrayBuffer };
}

export async function convertDocx(data: ArrayBuffer): Promise<ConversionResult> {
  const zip = await JSZip.loadAsync(data);
  const sheet = parseStyleSheet((await zip.file("word/styles.xml")?.async("string")) ?? "");

  const notesText = new Map<string, string>();
  const layoutStyles = new Map<string, Record<string, string>>();
  const classBySignature = new Map<string, string>();
  const styleMap: string[] = [];

  const layoutClass = (tag: string, css: Record<string, string>): string => {
    const signature = `${tag}|${JSON.stringify(css)}`;
    let cls = classBySignature.get(signature);
    if (!cls) {
      cls = `${LAYOUT_CLASS_PREFIX}${classBySignature.size + 1}`;
      classBySignature.set(signature, cls);
      layoutStyles.set(cls, css);
      // A unique style name per layout lets the style map add the class while
      // keeping the tag (p or h1–h6) mammoth would have used anyway.
      styleMap.push(`p[style-name='${cls}'] => ${tag}.${cls}:fresh`);
    }
    return cls;
  };

  const transformDocument = (document: DocElement & { notes: { resolve(ref: unknown): DocElement | null } }) => {
    walk(document, (el) => {
      if (el.type === "paragraph" && !el.numbering) {
        // List items keep mammoth's list mapping; their indent comes from the HTML list.
        const ind = (el.indent ?? {}) as { start?: string; end?: string; firstLine?: string; hanging?: string };
        const layout = resolveLayout(sheet, el.styleId as string | null, {
          indent: readIndent(ind),
          alignment: (el.alignment as string | null) ?? undefined,
        });
        const css = layoutToCss(layout);
        if (Object.keys(css).length > 0) {
          el.styleName = layoutClass(defaultTag(el.styleId, el.styleName), css);
        }
      }
      if (el.type === "noteReference") {
        const note = document.notes.resolve(el);
        if (note) {
          notesText.set(
            `${String(note.noteType)}-${String(note.noteId)}`,
            (note.body as unknown as DocElement[]).map(rawTextOf).join(""),
          );
        }
      }
    });
    return document;
  };

  // `styleMap` is filled in by transformDocument. Mammoth 1.13 keeps a reference
  // to this array and reads it only after the transform, which is why the
  // version is pinned; the indent tests fail if that ever changes.
  const html = await mammoth.convertToHtml(input(data), {
    styleMap,
    transformDocument: transformDocument as (element: unknown) => unknown,
    convertImage: mammoth.images.dataUri,
    externalFileAccess: false,
    ignoreEmptyParagraphs: false,
  });
  const raw = await mammoth.extractRawText(input(data));

  const messages = [...html.messages, ...raw.messages].map((m) => `${m.type}: ${m.message}`);
  return {
    html: html.value,
    rawText: raw.value,
    notesText,
    layoutStyles,
    messages: [...new Set(messages)],
  };
}
