import { describe, expect, it } from "vitest";
import { compareTexts, diffWords, extractHtmlText, normalizeText, runTextCheck } from "../src/lib/textCheck";

describe("normalizeText", () => {
  it("collapses whitespace runs, tabs, newlines and non-breaking spaces", () => {
    expect(normalizeText("  a\t b\n\nc  d  ")).toBe("a b c d");
  });

  it("does not fold quotes, dashes, case or accents", () => {
    const s = "“Café” — it’s ‘Naïve’ – 1990–2000";
    expect(normalizeText(s)).toBe(s);
  });
});

describe("compareTexts", () => {
  it("reports identical text", () => {
    const r = compareTexts("Hello world.\n\nSecond paragraph.", "Hello world. Second paragraph.");
    expect(r.identical).toBe(true);
    expect(r.wordCount).toBe(4);
  });

  it("treats empty input as identical with zero words", () => {
    expect(compareTexts("", "  \n ")).toEqual({ identical: true, wordCount: 0, ops: [] });
  });

  it("detects a changed word", () => {
    const r = compareTexts("the quick brown fox", "the quick red fox");
    expect(r.identical).toBe(false);
    expect(r.ops).toEqual([
      { type: "equal", words: ["the", "quick"] },
      { type: "removed", words: ["brown"] },
      { type: "added", words: ["red"] },
      { type: "equal", words: ["fox"] },
    ]);
  });

  it("detects smart quotes changed to straight quotes", () => {
    const r = compareTexts("it’s “fine”", "it's \"fine\"");
    expect(r.identical).toBe(false);
  });

  it("detects an em dash changed to a hyphen", () => {
    expect(compareTexts("yes — no", "yes - no").identical).toBe(false);
  });

  it("detects a case change", () => {
    expect(compareTexts("Brightspace", "brightspace").identical).toBe(false);
  });

  it("detects a missing word", () => {
    const r = compareTexts("one two three", "one three");
    expect(r.ops).toContainEqual({ type: "removed", words: ["two"] });
  });

  it("detects an added word", () => {
    const r = compareTexts("one three", "one two three");
    expect(r.ops).toContainEqual({ type: "added", words: ["two"] });
  });

  it("detects two words joined together", () => {
    expect(compareTexts("foo bar", "foobar").identical).toBe(false);
  });
});

describe("diffWords", () => {
  it("returns a single equal op for identical input", () => {
    expect(diffWords(["a", "b"], ["a", "b"])).toEqual([{ type: "equal", words: ["a", "b"] }]);
  });

  it("handles multiple separated changes", () => {
    const ops = diffWords("a b c d e f".split(" "), "a X c d Y f".split(" "));
    expect(ops.filter((o) => o.type !== "equal")).toEqual([
      { type: "removed", words: ["b"] },
      { type: "added", words: ["X"] },
      { type: "removed", words: ["e"] },
      { type: "added", words: ["Y"] },
    ]);
  });

  it("rebuilds both sides exactly", () => {
    const a = "the cat sat on the mat today".split(" ");
    const b = "a cat sat under the mat".split(" ");
    const ops = diffWords(a, b);
    const left = ops.filter((o) => o.type !== "added").flatMap((o) => o.words);
    const right = ops.filter((o) => o.type !== "removed").flatMap((o) => o.words);
    expect(left).toEqual(a);
    expect(right).toEqual(b);
  });

  it("falls back to a block diff for very large changed regions", () => {
    const a = Array.from({ length: 3000 }, (_, i) => `a${i}`);
    const b = Array.from({ length: 3000 }, (_, i) => `b${i}`);
    const ops = diffWords(["same", ...a], ["same", ...b]);
    expect(ops).toEqual([
      { type: "equal", words: ["same"] },
      { type: "removed", words: a },
      { type: "added", words: b },
    ]);
  });
});

describe("extractHtmlText", () => {
  it("separates block elements and decodes entities", () => {
    const { body } = extractHtmlText("<h1>Title</h1><p>A &amp; B</p><ul><li>one</li><li>two</li></ul>");
    expect(normalizeText(body)).toBe("Title A & B one two");
  });

  it("separates table cells", () => {
    const { body } = extractHtmlText("<table><tr><td><p>x</p></td><td><p>y</p></td></tr></table>");
    expect(normalizeText(body)).toBe("x y");
  });

  it("mirrors mammoth raw text for <br> (no separator)", () => {
    expect(extractHtmlText("<p>a<br>b</p>").body).toBe("\nab\n");
  });

  it("excludes note markers from the body and reads notes separately", () => {
    const html =
      '<p>Claim<sup class="note-ref">1</sup> here.</p><hr>' +
      '<ol class="notes"><li id="footnote-1"><p>Source text.</p></li></ol>';
    const { body, notes } = extractHtmlText(html);
    expect(normalizeText(body)).toBe("Claim here.");
    expect([...notes]).toEqual([["footnote-1", "\nSource text.\n"]]);
  });
});

describe("runTextCheck", () => {
  const notes = new Map([["footnote-1", "Source text.\n\n"]]);
  const html =
    '<p>Claim<sup class="note-ref">1</sup> here.</p><hr>' +
    '<ol class="notes"><li id="footnote-1"><p>Source text.</p></li></ol>';

  it("passes when body and notes match", () => {
    const r = runTextCheck("Claim here.\n\n", notes, html);
    expect(r.identical).toBe(true);
  });

  it("fails when a note's text differs", () => {
    const r = runTextCheck("Claim here.\n\n", new Map([["footnote-1", "Other text.\n\n"]]), html);
    expect(r.identical).toBe(false);
    expect(r.notes[0].comparison.identical).toBe(false);
  });

  it("fails when a note is missing from the HTML", () => {
    const r = runTextCheck("Claim here.\n\n", new Map([...notes, ["footnote-2", "Another."]]), html);
    expect(r.identical).toBe(false);
    expect(r.missingNotes).toEqual(["footnote-2"]);
  });

  it("fails when the HTML has a note the original does not", () => {
    const r = runTextCheck("Claim here.\n\n", new Map(), html);
    expect(r.extraNotes).toEqual(["footnote-1"]);
    expect(r.identical).toBe(false);
  });

  it("fails when body wording changed", () => {
    const r = runTextCheck("Claim there.\n\n", notes, html);
    expect(r.identical).toBe(false);
    expect(r.body.identical).toBe(false);
  });
});
