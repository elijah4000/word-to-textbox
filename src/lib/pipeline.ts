import { convertDocx, type ConversionResult } from "./convert";
import { issuesFromScan, scanDocx, type Issue, type Scan } from "./detect";
import { formatHtml, type FormatOptions, type FormatResult } from "./format";
import { extractHtmlText, runTextCheck, type TextCheckResult } from "./textCheck";

/** Everything read from the file once. Formatting is re-applied from this on every toggle. */
export interface LoadedDocument {
  conversion: ConversionResult;
  scan: Scan;
}

export interface Output {
  formatted: FormatResult;
  plainText: string;
  textCheck: TextCheckResult;
  issues: Issue[];
}

export async function loadDocx(data: ArrayBuffer): Promise<LoadedDocument> {
  // Scan first: it fails fast with a clear message if the file is not a .docx.
  const scan = await scanDocx(data);
  const conversion = await convertDocx(data);
  return { conversion, scan };
}

export function produceOutput(doc: LoadedDocument, options: FormatOptions): Output {
  const formatted = formatHtml(doc.conversion.html, options);
  const textCheck = runTextCheck(doc.conversion.rawText, doc.conversion.notesText, formatted.html);
  const plainText = toPlainText(formatted.html);
  const issues = issuesFromScan(doc.scan, { embedImages: options.embedImages });
  return { formatted, plainText, textCheck, issues };
}

/** Plain-text alternative for the clipboard: paragraphs separated by blank lines. */
export function toPlainText(html: string): string {
  const { body, notes } = extractHtmlText(html, { lineBreak: "\n", blockSeparator: "\n\n" });
  const tidy = (s: string) => s.replace(/\n{3,}/g, "\n\n").trim();
  const noteLines = [...notes.values()].map((text, i) => `${i + 1}. ${tidy(text)}`);
  return [tidy(body), ...(noteLines.length ? ["", ...noteLines] : [])].join("\n");
}
