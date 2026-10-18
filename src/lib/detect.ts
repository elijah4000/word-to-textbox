import JSZip from "jszip";

/**
 * Scans the raw Word XML for content that mammoth cannot convert faithfully,
 * so nothing is ever dropped without the user being told.
 */

export type Severity = "error" | "warning" | "info";

export interface Issue {
  id: string;
  severity: Severity;
  label: string;
  count: number;
  detail: string;
}

export interface Scan {
  equations: number;
  textBoxes: number;
  insertions: number;
  deletions: number;
  moves: number;
  formatChanges: number;
  comments: number;
  headers: number;
  footers: number;
  images: number;
  charts: number;
  smartArt: number;
  embeddedObjects: number;
  footnotes: number;
  endnotes: number;
}

function count(xml: string, tag: string): number {
  // `<w:ins ` or `<w:ins>` but not `<w:instrText`
  const re = new RegExp(`<${tag}[\\s>/]`, "g");
  return xml.match(re)?.length ?? 0;
}

/** Word stores many drawings twice (mc:Choice + mc:Fallback); count each once. */
function withoutFallbacks(xml: string): string {
  return xml.replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, "");
}

function countNotes(xml: string, tag: "w:footnote" | "w:endnote"): number {
  const re = new RegExp(`<${tag}\\b([^>]*)>`, "g");
  let n = 0;
  for (const match of xml.matchAll(re)) {
    // Separator "notes" are page furniture, not content.
    if (!/w:type="(separator|continuationSeparator|continuationNotice)"/.test(match[1])) n++;
  }
  return n;
}

function hasText(xml: string): boolean {
  for (const match of xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)) {
    if (match[1].trim() !== "") return true;
  }
  // Page-number fields etc. have no w:t until rendered.
  return /<w:fldSimple\b|<w:instrText\b/.test(xml);
}

export async function scanDocx(data: ArrayBuffer | Uint8Array): Promise<Scan> {
  const zip = await JSZip.loadAsync(data);
  const read = async (path: string) => (await zip.file(path)?.async("string")) ?? "";

  const documentXml = await read("word/document.xml");
  if (!documentXml) throw new Error("This file has no word/document.xml. Is it really a .docx?");
  const footnotesXml = await read("word/footnotes.xml");
  const endnotesXml = await read("word/endnotes.xml");
  const commentsXml = await read("word/comments.xml");

  // The parts whose text ends up in the output.
  const content = withoutFallbacks([documentXml, footnotesXml, endnotesXml].join("\n"));

  let headers = 0;
  let footers = 0;
  for (const path of Object.keys(zip.files)) {
    const m = /^word\/(header|footer)\d*\.xml$/.exec(path);
    if (!m) continue;
    if (!hasText(await read(path))) continue;
    if (m[1] === "header") headers++;
    else footers++;
  }

  return {
    equations: count(content, "m:oMath"),
    textBoxes: count(content, "w:txbxContent"),
    insertions: count(content, "w:ins"),
    deletions: count(content, "w:del"),
    moves: count(content, "w:moveFrom") + count(content, "w:moveTo"),
    formatChanges:
      count(content, "w:rPrChange") + count(content, "w:pPrChange") + count(content, "w:tblPrChange") +
      count(content, "w:tcPrChange") + count(content, "w:trPrChange") + count(content, "w:sectPrChange"),
    comments: count(commentsXml, "w:comment"),
    headers,
    footers,
    images: count(content, "pic:pic") + count(content, "v:imagedata"),
    charts: count(content, "c:chart"),
    smartArt: count(content, "dgm:relIds"),
    embeddedObjects: count(content, "o:OLEObject"),
    footnotes: countNotes(footnotesXml, "w:footnote"),
    endnotes: countNotes(endnotesXml, "w:endnote"),
  };
}

export interface IssueContext {
  embedImages: boolean;
}

export function issuesFromScan(scan: Scan, ctx: IssueContext): Issue[] {
  const issues: Issue[] = [];
  const add = (id: string, severity: Severity, label: string, n: number, detail: string) => {
    if (n > 0) issues.push({ id, severity, label, count: n, detail });
  };

  add("equations", "error", "Equations", scan.equations,
    "Not converted. Equation content is MISSING from the output. Retype it in Brightspace's equation editor or attach the original file if allowed.");
  add("charts", "error", "Charts", scan.charts,
    "Not converted. Charts are MISSING from the output.");
  add("smartart", "error", "SmartArt diagrams", scan.smartArt,
    "Not converted. SmartArt is MISSING from the output (its text may also be missing).");
  add("objects", "error", "Embedded objects", scan.embeddedObjects,
    "Embedded objects (e.g. spreadsheets, OLE equations) are not converted and are MISSING from the output.");

  const tracked = scan.insertions + scan.deletions + scan.moves + scan.formatChanges;
  add("tracked", "error", "Tracked changes", tracked,
    `${scan.insertions} insertion(s), ${scan.deletions} deletion(s), ${scan.moves} move(s), ${scan.formatChanges} formatting change(s). ` +
      "The output shows the document as if every change were accepted: inserted text is kept, deleted text is removed. " +
      "Accept or reject the changes in Word (Review → Accept/Reject) and load the file again so you decide what is submitted.");

  add("textboxes", "warning", "Text boxes", scan.textBoxes,
    "Text from text boxes is placed inline where the box is anchored; the box's position, border and shading are lost. Check it appears where you expect.");
  add("comments", "warning", "Comments", scan.comments,
    "Comments are not included in the output.");
  add("headers", "warning", "Headers", scan.headers,
    "Page headers (e.g. running heads, your name/page number) are not included. Add them manually if your instructor requires them.");
  add("footers", "warning", "Footers", scan.footers,
    "Page footers are not included. Add them manually if required.");

  if (ctx.embedImages) {
    add("images", "info", "Images", scan.images,
      "Embedded in the copied HTML. Brightspace may not keep pasted images; if one is missing after pasting, insert it with Brightspace's image button.");
  } else {
    add("images", "warning", "Images", scan.images,
      "Image embedding is OFF, so images are NOT included in the output. Turn it on or insert them in Brightspace.");
  }

  add("footnotes", "info", "Footnotes", scan.footnotes,
    "Converted. Markers become superscript numbers and the notes are listed at the end of the text.");
  add("endnotes", "info", "Endnotes", scan.endnotes,
    "Converted. Listed at the end of the text, numbered in the order they are referenced.");

  return issues;
}
