import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { DEFAULT_OPTIONS, type FormatOptions } from "../src/lib/format";
import { loadDocx, produceOutput } from "../src/lib/pipeline";
import { normalizeText } from "../src/lib/textCheck";

function load(name: string): ArrayBuffer {
  const buf = readFileSync(resolve(process.cwd(), "samples", name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

const ALL_ON: FormatOptions = {
  doubleSpacing: true,
  hangingIndent: true,
  embedImages: true,
  fontFamily: "'Times New Roman', Times, serif",
  fontSize: "12pt",
};

describe("sample.docx", async () => {
  const doc = await loadDocx(load("sample.docx"));

  it("passes the text check with default and with all options", () => {
    for (const options of [DEFAULT_OPTIONS, ALL_ON, { ...ALL_ON, embedImages: false, hangingIndent: false }]) {
      const out = produceOutput(doc, options);
      expect(out.textCheck.identical, JSON.stringify(out.textCheck.body.ops.filter((o) => o.type !== "equal"))).toBe(true);
    }
  });

  it("keeps special characters byte for byte", () => {
    const out = produceOutput(doc, DEFAULT_OPTIONS);
    for (const s of ["“smart quotes,”", "‘single quotes’", "doesn’t", "—", "1990–2000", "café", "naïve", "pp. 12–14."]) {
      expect(out.formatted.html).toContain(s);
    }
  });

  it("checks both footnotes", () => {
    const out = produceOutput(doc, DEFAULT_OPTIONS);
    expect(out.textCheck.notes.map((n) => n.id)).toEqual(["footnote-1", "footnote-2"]);
    expect(out.formatted.html).toContain('<sup class="note-ref">1</sup>');
    expect(out.formatted.html).not.toContain("↑");
  });

  it("converts headings, lists and tables with inline borders", () => {
    const html = produceOutput(doc, DEFAULT_OPTIONS).formatted.html;
    expect(html).toMatch(/<h1>Seasonal Patterns in Trail Use<\/h1>/);
    expect(html).toMatch(/<ul>/);
    expect(html).toMatch(/<ol>/);
    expect(html).toMatch(/<td style="border: 1px solid/);
  });

  it("keeps Word's hanging indent on the three Works Cited entries", () => {
    for (const hangingIndent of [true, false]) {
      const out = produceOutput(doc, { ...DEFAULT_OPTIONS, hangingIndent });
      expect(out.formatted.hangingParagraphs).toBe(3);
      expect(out.formatted.addedHanging).toBe(0);
      expect(out.formatted.html.match(/padding-left: 36pt; text-indent: -36pt/g)).toHaveLength(3);
      expect(out.formatted.html).not.toContain("d2b-layout");
    }
  });

  it("applies spacing and font inline", () => {
    const html = produceOutput(doc, ALL_ON).formatted.html;
    expect(html).toContain("line-height: 2");
    expect(html).toContain("font-size: 12pt");
    expect(html).toContain("Times New Roman");
  });

  it("embeds or omits the image and always reports it", () => {
    const on = produceOutput(doc, DEFAULT_OPTIONS);
    expect(on.formatted.images).toBe(1);
    expect(on.formatted.html).toMatch(/<img [^>]*src="data:image\/png;base64,/);
    expect(on.issues.find((i) => i.id === "images")?.severity).toBe("info");

    const off = produceOutput(doc, { ...DEFAULT_OPTIONS, embedImages: false });
    expect(off.formatted.html).not.toContain("<img");
    expect(off.issues.find((i) => i.id === "images")).toMatchObject({ severity: "warning", count: 1 });
  });

  it("lists footnotes in the warning panel", () => {
    const out = produceOutput(doc, DEFAULT_OPTIONS);
    expect(out.issues.find((i) => i.id === "footnotes")?.count).toBe(2);
    expect(out.issues.some((i) => i.severity === "error")).toBe(false);
  });

  it("produces a readable plain-text alternative", () => {
    const text = produceOutput(doc, DEFAULT_OPTIONS).plainText;
    expect(text.startsWith("Seasonal Patterns in Trail Use\n\n")).toBe(true);
    expect(text).toContain("1. Environment Canada");
    expect(normalizeText(text)).toContain("This sentence starts after a line break.");
  });
});

describe("problems.docx", async () => {
  const doc = await loadDocx(load("problems.docx"));
  const out = produceOutput(doc, DEFAULT_OPTIONS);
  const byId = Object.fromEntries(out.issues.map((i) => [i.id, i]));

  it("flags every unsupported feature with counts", () => {
    expect(byId.equations).toMatchObject({ severity: "error", count: 1 });
    expect(byId.tracked).toMatchObject({ severity: "error", count: 2 });
    expect(byId.comments).toMatchObject({ severity: "warning", count: 1 });
    expect(byId.headers).toMatchObject({ severity: "warning", count: 1 });
    expect(byId.footers).toMatchObject({ severity: "warning", count: 1 });
    expect(byId.textboxes).toMatchObject({ severity: "warning", count: 1 });
  });

  it("still reports which words reach the output", () => {
    // Mammoth accepts insertions and drops deletions on both sides of the check,
    // which is why tracked changes are flagged as errors above.
    const html = out.formatted.html;
    expect(html).toContain("inserted words");
    expect(html).not.toContain("deleted words");
    expect(html).toContain("Text inside a text box.");
  });
});

describe("invalid input", () => {
  it("rejects a file that is not a .docx", async () => {
    const bytes = new TextEncoder().encode("not a zip");
    await expect(loadDocx(bytes.buffer as ArrayBuffer)).rejects.toThrow();
  });
});
