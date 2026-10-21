/**
 * Resolves the indentation and alignment Word would actually apply to a
 * paragraph: document defaults, then the paragraph style and every style it is
 * based on, then the paragraph's own settings. Mammoth only reports the last of
 * these, but most essays get their first-line indent from the "Normal" style.
 */

/** All values in twips (1/1440 inch). `first` > 0 is a first-line indent, < 0 a hanging indent. */
export interface Indent {
  left?: number;
  right?: number;
  first?: number;
}

export interface ParagraphLayout {
  indent: Indent;
  /** Word's w:jc value, e.g. "center", "right", "both". */
  alignment?: string;
}

interface StyleDef extends ParagraphLayout {
  basedOn?: string;
}

export interface StyleSheet {
  defaults: ParagraphLayout;
  defaultStyleId?: string;
  styles: Map<string, StyleDef>;
}

function twips(value: string | null | undefined): number | undefined {
  if (value == null || value === "") return undefined;
  const n = parseInt(value, 10);
  return Number.isFinite(n) ? n : undefined;
}

/** Reads a w:ind element (or mammoth's copy of its attributes). */
export function readIndent(attrs: {
  start?: string | null;
  end?: string | null;
  firstLine?: string | null;
  hanging?: string | null;
}): Indent {
  const indent: Indent = {};
  const left = twips(attrs.start);
  const right = twips(attrs.end);
  const hanging = twips(attrs.hanging);
  const firstLine = twips(attrs.firstLine);
  if (left !== undefined) indent.left = left;
  if (right !== undefined) indent.right = right;
  // Word ignores firstLine when hanging is present.
  if (hanging !== undefined) indent.first = -hanging;
  else if (firstLine !== undefined) indent.first = firstLine;
  return indent;
}

function readPPr(pPr: Element | null | undefined): ParagraphLayout {
  if (!pPr) return { indent: {} };
  const child = (name: string) => [...pPr.children].find((c) => c.tagName === name);
  const ind = child("w:ind");
  const jc = child("w:jc");
  return {
    indent: ind
      ? readIndent({
          start: ind.getAttribute("w:start") ?? ind.getAttribute("w:left"),
          end: ind.getAttribute("w:end") ?? ind.getAttribute("w:right"),
          firstLine: ind.getAttribute("w:firstLine"),
          hanging: ind.getAttribute("w:hanging"),
        })
      : {},
    alignment: jc?.getAttribute("w:val") ?? undefined,
  };
}

export function parseStyleSheet(stylesXml: string): StyleSheet {
  const sheet: StyleSheet = { defaults: { indent: {} }, styles: new Map() };
  if (!stylesXml) return sheet;
  const doc = new DOMParser().parseFromString(stylesXml, "application/xml");

  const pPrDefault = doc.getElementsByTagName("w:pPrDefault")[0];
  if (pPrDefault) sheet.defaults = readPPr(pPrDefault.getElementsByTagName("w:pPr")[0]);

  for (const style of doc.getElementsByTagName("w:style")) {
    if (style.getAttribute("w:type") !== "paragraph") continue;
    const id = style.getAttribute("w:styleId");
    if (!id) continue;
    const children = [...style.children];
    const basedOn = children.find((c) => c.tagName === "w:basedOn")?.getAttribute("w:val") ?? undefined;
    const layout = readPPr(children.find((c) => c.tagName === "w:pPr"));
    sheet.styles.set(id, { ...layout, basedOn });
    if (style.getAttribute("w:default") === "1" || style.getAttribute("w:default") === "true") {
      sheet.defaultStyleId = id;
    }
  }
  return sheet;
}

function merge(base: ParagraphLayout, over: ParagraphLayout): ParagraphLayout {
  return {
    indent: { ...base.indent, ...over.indent },
    alignment: over.alignment ?? base.alignment,
  };
}

/** Layout of a style including everything it inherits. */
export function styleLayout(sheet: StyleSheet, styleId: string | null | undefined): ParagraphLayout {
  const chain: StyleDef[] = [];
  const seen = new Set<string>();
  let id = styleId ?? sheet.defaultStyleId;
  while (id && !seen.has(id)) {
    seen.add(id);
    const def = sheet.styles.get(id);
    if (!def) break;
    chain.unshift(def);
    id = def.basedOn;
  }
  return chain.reduce(merge, sheet.defaults);
}

export function resolveLayout(
  sheet: StyleSheet,
  styleId: string | null | undefined,
  direct: ParagraphLayout,
): ParagraphLayout {
  return merge(styleLayout(sheet, styleId), direct);
}

const pt = (tw: number) => `${Math.round((tw / 20) * 100) / 100}pt`;

/** Inline CSS for a resolved layout. Zero values are omitted. */
export function layoutToCss(layout: ParagraphLayout): Record<string, string> {
  const css: Record<string, string> = {};
  const { left = 0, right = 0, first = 0 } = layout.indent;
  if (left) css["padding-left"] = pt(left);
  if (right) css["padding-right"] = pt(right);
  if (first) css["text-indent"] = pt(first);
  switch (layout.alignment) {
    case "center":
      css["text-align"] = "center";
      break;
    case "right":
    case "end":
      css["text-align"] = "right";
      break;
    case "both":
    case "distribute":
      css["text-align"] = "justify";
      break;
  }
  return css;
}
