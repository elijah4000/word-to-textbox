// Generates the test documents in samples/. Run with `npm run make-sample`.
//
// samples/sample.docx   – a typical essay: headings, lists, a table, footnotes,
//                         a hanging-indent Works Cited list and an image.
// samples/problems.docx – content that cannot be converted cleanly: an equation,
//                         a text box, tracked changes, a comment and a header/footer.
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync, crc32 } from "node:zlib";
import {
  AlignmentType, CommentRangeEnd, CommentRangeStart, CommentReference, DeletedTextRun, Document,
  ExternalHyperlink, Footer, FootnoteReferenceRun, Header, HeadingLevel, ImageRun, InsertedTextRun,
  LevelFormat, Math as OMath, MathFraction, MathRun, Packer, Paragraph, Tab, Table, TableCell, TableRow,
  TextRun, Textbox, WidthType,
} from "docx";

// --- a tiny PNG (two-tone bar chart look), built by hand -------------------
function png(width, height, pixel) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td) >>> 0);
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr.set([8, 2, 0, 0, 0], 8); // 8-bit RGB
  const rows = [];
  for (let y = 0; y < height; y++) {
    rows.push(0); // filter: none
    for (let x = 0; x < width; x++) rows.push(...pixel(x, y));
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(Buffer.from(rows))),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}
const chart = png(120, 60, (x, y) => {
  const bar = Math.floor(x / 20);
  const heights = [20, 35, 50, 30, 45, 15];
  return 60 - y <= heights[bar] && x % 20 > 3 ? [31, 95, 153] : [240, 240, 235];
});

const numbering = {
  config: [
    {
      reference: "numbers",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    },
  ],
};

const bib = (text, italic) =>
  new Paragraph({
    indent: { left: 720, hanging: 720 },
    children: italic
      ? [new TextRun(text[0]), new TextRun({ text: text[1], italics: true }), new TextRun(text[2])]
      : [new TextRun(text)],
  });

const cell = (text, bold = false) =>
  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold })] })] });

const sample = new Document({
  creator: "DocxToBrightspace test generator",
  title: "Sample essay",
  numbering,
  footnotes: {
    1: { children: [new Paragraph("Environment Canada, “Historical Climate Data,” accessed 3 March 2026.")] },
    2: { children: [new Paragraph({ children: [new TextRun("See also the "), new TextRun({ text: "Field Guide", italics: true }), new TextRun(", pp. 12–14.")] })] },
  },
  sections: [
    {
      children: [
        new Paragraph({ text: "Seasonal Patterns in Trail Use", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          children: [
            new TextRun("This essay examines why visitors prefer early-season hikes — and why the data doesn’t always agree."),
            new FootnoteReferenceRun(1),
            new TextRun(" It uses “smart quotes,” ‘single quotes’, an em dash —, an en dash 1990–2000, café, naïve, and two  spaces."),
          ],
        }),
        new Paragraph({ text: "Background", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({
          children: [
            new TextRun({ text: "Bold text", bold: true }),
            new TextRun(", "),
            new TextRun({ text: "italic text", italics: true }),
            new TextRun(", "),
            new TextRun({ text: "underlined text", underline: {} }),
            new TextRun(", and a "),
            new ExternalHyperlink({ link: "https://example.org/trails", children: [new TextRun({ text: "hyperlink", style: "Hyperlink" })] }),
            new TextRun(". A tab follows:"),
            new TextRun({ children: [new Tab(), "after the tab."] }),
            new TextRun({ text: "This sentence starts after a line break.", break: 1 }),
            new FootnoteReferenceRun(2),
          ],
        }),
        new Paragraph({ text: "Key factors", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: "Snowpack and trail conditions", bullet: { level: 0 } }),
        new Paragraph({ text: "Daylight hours", bullet: { level: 0 } }),
        new Paragraph({ text: "Sunrise times in late spring", bullet: { level: 1 } }),
        new Paragraph({ text: "Collect counts at each trailhead.", numbering: { reference: "numbers", level: 0 } }),
        new Paragraph({ text: "Compare with weather records.", numbering: { reference: "numbers", level: 0 } }),
        new Paragraph({ text: "Report the differences.", numbering: { reference: "numbers", level: 0 } }),
        new Paragraph({ text: "Results", heading: HeadingLevel.HEADING_2 }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({ tableHeader: true, children: [cell("Month", true), cell("Visitors", true), cell("Avg. temp (°C)", true)] }),
            new TableRow({ children: [cell("April"), cell("1,240"), cell("8.5")] }),
            new TableRow({ children: [cell("May"), cell("3,980"), cell("14.2")] }),
            new TableRow({ children: [cell("June"), cell("5,105"), cell("19.0")] }),
          ],
        }),
        new Paragraph("Figure 1 shows monthly visitors."),
        new Paragraph({
          children: [new ImageRun({ type: "png", data: chart, transformation: { width: 240, height: 120 }, altText: { title: "Figure 1", description: "Bar chart of monthly visitors", name: "chart" } })],
        }),
        new Paragraph({ text: "Works Cited", heading: HeadingLevel.HEADING_1 }),
        bib(["Alvarez, Maria. ", "Mountain Seasons: A Study of Alpine Visitors", ". Northern Press, 2021."], true),
        bib(["Chen, David, and Priya Shah. “Weather and Weekend Hiking.” ", "Journal of Outdoor Studies", ", vol. 14, no. 2, 2023, pp. 45–67, https://doi.org/10.0000/example.2023."], true),
        bib("Parks Canada. “Visitor Statistics 2024.” Government of Canada, 2025, www.example.gc.ca/stats. Accessed 1 Mar. 2026."),
      ],
    },
  ],
});

const problems = new Document({
  creator: "DocxToBrightspace test generator",
  title: "Problem content",
  comments: {
    children: [{ id: 0, author: "Instructor", date: new Date("2026-01-01"), children: [new Paragraph("Please cite this.")] }],
  },
  sections: [
    {
      headers: { default: new Header({ children: [new Paragraph("Student Name — Running Head")] }) },
      footers: { default: new Footer({ children: [new Paragraph("Footer text")] }) },
      children: [
        new Paragraph({ text: "Problem content", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          children: [
            new TextRun("The formula is "),
            new OMath({ children: [new MathFraction({ numerator: [new MathRun("a")], denominator: [new MathRun("b")] })] }),
            new TextRun(" as shown."),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun("Tracked: "),
            new InsertedTextRun({ text: "inserted words", id: 1, author: "Student", date: "2026-01-01T00:00:00Z" }),
            new TextRun(" and "),
            new DeletedTextRun({ text: "deleted words", id: 2, author: "Student", date: "2026-01-01T00:00:00Z" }),
            new TextRun("."),
          ],
        }),
        new Paragraph({
          children: [
            new CommentRangeStart(0),
            new TextRun("A claim that needs a source."),
            new CommentRangeEnd(0),
            new TextRun({ children: [new CommentReference(0)] }),
          ],
        }),
        new Textbox({
          alignment: AlignmentType.LEFT,
          style: { width: "200pt", height: "50pt" },
          children: [new Paragraph("Text inside a text box.")],
        }),
      ],
    },
  ],
});

mkdirSync(new URL("../samples/", import.meta.url), { recursive: true });
writeFileSync(new URL("../samples/sample.docx", import.meta.url), await Packer.toBuffer(sample));
writeFileSync(new URL("../samples/problems.docx", import.meta.url), await Packer.toBuffer(problems));
console.log("Wrote samples/sample.docx and samples/problems.docx");
