import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS } from "../src/lib/format";
import { layoutToCss, parseStyleSheet, readIndent, resolveLayout, styleLayout } from "../src/lib/paragraphStyles";
import { loadDocx, produceOutput } from "../src/lib/pipeline";

const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:pPrDefault><w:pPr><w:ind w:right="144"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
    <w:pPr><w:ind w:firstLine="720"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="1440" w:firstLine="0"/><w:jc w:val="both"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Bib"><w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:start="720" w:hanging="720"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Loop"><w:basedOn w:val="Loop"/></w:style>
  <w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;

describe("paragraph style resolution", () => {
  const sheet = parseStyleSheet(STYLES);

  it("uses the default paragraph style when a paragraph has none", () => {
    expect(styleLayout(sheet, null).indent).toEqual({ right: 144, first: 720 });
  });

  it("inherits through basedOn and lets the child override", () => {
    const quote = styleLayout(sheet, "Quote");
    expect(quote.indent).toEqual({ right: 144, left: 1440, first: 0 });
    expect(quote.alignment).toBe("both");
  });

  it("reads w:start and w:hanging", () => {
    expect(styleLayout(sheet, "Bib").indent).toEqual({ right: 144, left: 720, first: -720 });
  });

  it("lets direct paragraph settings win", () => {
    const layout = resolveLayout(sheet, "Quote", { indent: readIndent({ firstLine: "360" }), alignment: "center" });
    expect(layout.indent.first).toBe(360);
    expect(layout.alignment).toBe("center");
  });

  it("survives a style based on itself", () => {
    // Not based on Normal, so it only gets the document defaults.
    expect(styleLayout(sheet, "Loop").indent).toEqual({ right: 144 });
  });

  it("handles a missing styles.xml", () => {
    expect(styleLayout(parseStyleSheet(""), "Anything")).toEqual({ indent: {} });
  });

  it("prefers hanging over firstLine like Word does", () => {
    expect(readIndent({ firstLine: "720", hanging: "360" })).toEqual({ first: -360 });
  });

  it("converts to inline CSS in points", () => {
    expect(layoutToCss({ indent: { left: 720, right: 0, first: -720 }, alignment: "both" })).toEqual({
      "padding-left": "36pt",
      "text-indent": "-36pt",
      "text-align": "justify",
    });
    expect(layoutToCss({ indent: { first: 0 }, alignment: "left" })).toEqual({});
  });
});

describe("indents.docx", async () => {
  const buf = readFileSync(resolve(process.cwd(), "samples", "indents.docx"));
  const doc = await loadDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
  const out = produceOutput(doc, DEFAULT_OPTIONS);
  const page = new DOMParser().parseFromString(out.formatted.html, "text/html");
  const para = (startsWith: string) => {
    const el = [...page.querySelectorAll<HTMLElement>("p, h1, h2")].find((p) => (p.textContent ?? "").startsWith(startsWith));
    if (!el) throw new Error(`No paragraph starting with "${startsWith}"`);
    return el.style;
  };

  it("passes the text check", () => {
    expect(out.textCheck.identical).toBe(true);
  });

  it("keeps the first-line indent inherited from the Normal style", () => {
    expect(para("This body paragraph").textIndent).toBe("36pt");
    expect(para("So does this").textIndent).toBe("36pt");
  });

  it("keeps a custom style's left indent", () => {
    const s = para("This quoted passage");
    expect(s.paddingLeft).toBe("72pt");
    expect(s.textIndent).toBe("");
  });

  it("combines a direct left indent with the inherited first-line indent", () => {
    const s = para("This paragraph follows");
    expect(s.paddingLeft).toBe("36pt");
    expect(s.textIndent).toBe("36pt");
  });

  it("does not indent paragraphs that turn the indent off", () => {
    expect(para("Jane Student").textIndent).toBe("");
  });

  it("keeps alignment from styles and paragraphs", () => {
    expect(para("Indentation Test Essay").textAlign).toBe("center");
    expect(para("20 October 2026").textAlign).toBe("right");
  });

  it("turns a leading Tab into a first-line indent and removes the tab", () => {
    const el = [...page.querySelectorAll("p")].find((p) => p.textContent?.includes("pressing Tab"))!;
    expect(el.style.textIndent).toBe("36pt");
    expect(el.textContent?.startsWith("This paragraph was indented")).toBe(true);
  });

  it("preserves tabs and runs of spaces inside a line", () => {
    const s = para("Name:");
    expect(s.whiteSpace).toBe("pre-wrap");
    expect(out.formatted.html).toContain("Name:\ta tab in the middle of a line, and   three spaces.");
  });

  it("keeps the blank line", () => {
    expect(out.formatted.html).toContain("<p><br></p>");
  });

  it("keeps Word's hanging indent and adds one where Works Cited lacks it", () => {
    const fromWord = para("Alvarez");
    expect(fromWord.paddingLeft).toBe("36pt");
    expect(fromWord.textIndent).toBe("-36pt");
    const added = para("Chen");
    expect(added.textIndent).toBe("-36pt");
    expect(out.formatted.addedHanging).toBe(1);
    expect(out.formatted.hangingParagraphs).toBe(2);
  });

  it("leaves Word's indents in place when the Works Cited option is off", () => {
    const off = produceOutput(doc, { ...DEFAULT_OPTIONS, hangingIndent: false });
    expect(off.formatted.addedHanging).toBe(0);
    expect(off.formatted.hangingParagraphs).toBe(1);
    expect(off.formatted.html).toContain("text-indent: 36pt");
  });

  it("leaves no internal classes in the output", () => {
    expect(out.formatted.html).not.toContain("d2b-layout");
  });
});
