import "./styles.css";
import { copyRichText } from "./lib/clipboard";
import { DEFAULT_OPTIONS, type FormatOptions } from "./lib/format";
import { loadDocx, produceOutput, type LoadedDocument, type Output } from "./lib/pipeline";
import type { DiffOp, TextComparison } from "./lib/textCheck";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const fileInput = $<HTMLInputElement>("file");
const drop = $("drop");
const fileName = $("file-name");
const loadError = $("load-error");
const optionsSet = $<HTMLFieldSetElement>("options");
const optDouble = $<HTMLInputElement>("opt-double");
const optHanging = $<HTMLInputElement>("opt-hanging");
const optImages = $<HTMLInputElement>("opt-images");
const optFont = $<HTMLSelectElement>("opt-font");
const optSize = $<HTMLSelectElement>("opt-size");
const hangingCount = $("hanging-count");
const imageSize = $("image-size");
const copyBtn = $<HTMLButtonElement>("copy");
const copyStatus = $("copy-status");
const ackRow = $("ack-row");
const ack = $<HTMLInputElement>("ack");
const warningsPanel = $("warnings-panel");
const issuesList = $("issues");
const messagesBox = $("messages-box");
const messagesCount = $("messages-count");
const messagesList = $("messages");
const checkPanel = $("check-panel");
const checkSummary = $("check-summary");
const checkDiff = $("check-diff");
const previewPanel = $("preview-panel");
const preview = $("preview");
const empty = $("empty");

let loaded: LoadedDocument | null = null;
let output: Output | null = null;

optDouble.checked = DEFAULT_OPTIONS.doubleSpacing;
optHanging.checked = DEFAULT_OPTIONS.hangingIndent;
optImages.checked = DEFAULT_OPTIONS.embedImages;

function readOptions(): FormatOptions {
  return {
    doubleSpacing: optDouble.checked,
    hangingIndent: optHanging.checked,
    embedImages: optImages.checked,
    fontFamily: optFont.value,
    fontSize: optSize.value,
  };
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, text?: string, className?: string) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

const LARGE_IMAGES = 1024 * 1024;
const CONTEXT_WORDS = 8;

async function handleFile(file: File) {
  loaded = null;
  output = null;
  loadError.hidden = true;
  fileName.textContent = file.name;
  copyStatus.textContent = "";
  if (!/\.docx$/i.test(file.name)) {
    showLoadError(
      /\.doc$/i.test(file.name)
        ? "This is an old .doc file. Open it in Word and use File → Save As → Word Document (.docx), then load that file."
        : "Please choose a Word .docx file.",
    );
    return;
  }
  try {
    loaded = await loadDocx(await file.arrayBuffer());
  } catch (err) {
    showLoadError(`Could not read this file: ${err instanceof Error ? err.message : String(err)}`);
    return;
  }
  optionsSet.disabled = false;
  render();
}

function showLoadError(message: string) {
  loadError.textContent = message;
  loadError.hidden = false;
  optionsSet.disabled = true;
  copyBtn.disabled = true;
  for (const p of [warningsPanel, checkPanel, previewPanel]) p.hidden = true;
  empty.hidden = false;
}

function render() {
  if (!loaded) return;
  output = produceOutput(loaded, readOptions());
  empty.hidden = true;

  // Options hints
  const h = output.formatted.hangingParagraphs;
  hangingCount.textContent = optHanging.checked ? `(${h} paragraph${h === 1 ? "" : "s"})` : "";
  if (loaded.scan.images === 0) {
    imageSize.textContent = "(no images)";
  } else if (optImages.checked) {
    const big = output.formatted.imageBytes > LARGE_IMAGES;
    imageSize.textContent = `(${output.formatted.images} image${output.formatted.images === 1 ? "" : "s"}, ${formatBytes(output.formatted.imageBytes)}${big ? ", large: Brightspace may reject or strip them" : ""})`;
  } else {
    imageSize.textContent = `(${loaded.scan.images} image${loaded.scan.images === 1 ? "" : "s"} left out)`;
  }

  renderIssues(output);
  renderTextCheck(output);

  preview.innerHTML = output.formatted.html;
  previewPanel.hidden = false;

  const needsAck = !output.textCheck.identical || output.issues.some((i) => i.severity === "error");
  ackRow.hidden = !needsAck;
  if (!needsAck) ack.checked = false;
  copyBtn.disabled = needsAck && !ack.checked;
}

function renderIssues(out: Output) {
  issuesList.replaceChildren();
  const extra: { severity: "error" | "warning" | "info"; label: string; count: number; detail: string }[] = [];
  if (out.formatted.images > 0 && out.formatted.imageBytes > LARGE_IMAGES) {
    extra.push({
      severity: "warning",
      label: "Large images",
      count: out.formatted.images,
      detail: `Embedded images add about ${formatBytes(out.formatted.imageBytes)} to the paste. Brightspace may refuse or strip large pasted content. Consider turning off image embedding and inserting images with Brightspace's image button.`,
    });
  }
  const all = [...out.issues, ...extra];
  for (const issue of all) {
    const li = el("li", undefined, `issue ${issue.severity}`);
    li.append(
      el("span", issue.severity === "error" ? "Not converted" : issue.severity === "warning" ? "Check" : "Note", "sev"),
      el("b", issue.label),
      el("span", `× ${issue.count}`, "count"),
      el("p", issue.detail),
    );
    issuesList.append(li);
  }
  if (all.length === 0) {
    const li = el("li", undefined, "issue info");
    li.append(el("p", "No equations, text boxes, tracked changes, comments, headers/footers, images or notes were found."));
    issuesList.append(li);
  }

  const messages = loaded?.conversion.messages ?? [];
  messagesBox.hidden = messages.length === 0;
  messagesCount.textContent = String(messages.length);
  messagesList.replaceChildren(...messages.map((m) => el("li", m)));
  warningsPanel.hidden = false;
}

function renderOps(ops: DiffOp[]): HTMLElement {
  const box = el("div", undefined, "diff");
  ops.forEach((op, i) => {
    if (op.type === "equal") {
      const first = i === 0;
      const last = i === ops.length - 1;
      const words = op.words;
      let text: string;
      if (words.length > CONTEXT_WORDS * 2 + 1 || (first && words.length > CONTEXT_WORDS) || (last && words.length > CONTEXT_WORDS)) {
        const head = first ? [] : words.slice(0, CONTEXT_WORDS);
        const tail = last ? [] : words.slice(-CONTEXT_WORDS);
        box.append(head.length ? head.join(" ") + " " : "");
        box.append(el("span", `… ${words.length - head.length - tail.length} matching words …`, "gap"));
        text = tail.length ? " " + tail.join(" ") + " " : " ";
      } else {
        text = words.join(" ") + " ";
      }
      box.append(text);
    } else {
      box.append(el(op.type === "removed" ? "del" : "ins", op.words.join(" ")), " ");
    }
  });
  return box;
}

function renderComparison(title: string, comparison: TextComparison) {
  if (comparison.identical) return;
  checkDiff.append(el("p", title, "diff-title"), renderOps(comparison.ops));
}

function renderTextCheck(out: Output) {
  const tc = out.textCheck;
  checkDiff.replaceChildren();
  const noteCount = tc.notes.length;
  const scope = `${tc.body.wordCount} words${noteCount ? ` + ${noteCount} note${noteCount === 1 ? "" : "s"}` : ""}`;
  if (tc.identical) {
    checkSummary.textContent = `Text identical ✓ (${scope} compared)`;
    checkSummary.className = "check ok";
    if (out.issues.some((i) => i.severity === "error")) {
      checkDiff.append(
        el("p", "The check can only compare text the converter can read. Items marked “Not converted” above are missing from both sides, so this check does not cover them.", "diff-title"),
      );
    }
  } else {
    checkSummary.textContent = "Text differs ✗ — review before copying";
    checkSummary.className = "check bad";
    const legend = el("p", undefined, "hint");
    legend.append("Red struck-through = in the Word file but not in the copy. Green = in the copy but not in the Word file.");
    checkDiff.append(legend);
    renderComparison("Main text", tc.body);
    for (const note of tc.notes) renderComparison(`Note ${note.id}`, note.comparison);
    if (tc.missingNotes.length) checkDiff.append(el("p", `Notes missing from the copy: ${tc.missingNotes.join(", ")}`, "diff-title"));
    if (tc.extraNotes.length) checkDiff.append(el("p", `Unexpected notes in the copy: ${tc.extraNotes.join(", ")}`, "diff-title"));
  }
  checkPanel.hidden = false;
}

// --- events ---------------------------------------------------------------

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) void handleFile(file);
});

for (const type of ["dragenter", "dragover"]) {
  drop.addEventListener(type, (e) => {
    e.preventDefault();
    drop.classList.add("over");
  });
}
for (const type of ["dragleave", "drop"]) {
  drop.addEventListener(type, () => drop.classList.remove("over"));
}
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  const file = (e as DragEvent).dataTransfer?.files?.[0];
  if (file) void handleFile(file);
});
// Dropping a file anywhere else should not navigate away from the page.
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

for (const input of [optDouble, optHanging, optImages, optFont, optSize]) {
  input.addEventListener("change", () => {
    copyStatus.textContent = "";
    render();
  });
}

ack.addEventListener("change", () => {
  copyBtn.disabled = !ack.checked;
});

copyBtn.addEventListener("click", async () => {
  if (!output) return;
  copyStatus.className = "status";
  try {
    const method = await copyRichText(output.formatted.html, output.plainText, preview);
    copyStatus.textContent =
      method === "clipboard-api"
        ? "Copied ✓ Now paste into the Brightspace text box."
        : "Copied ✓ (using the browser's copy command). Now paste into Brightspace.";
    copyStatus.classList.add("ok");
  } catch (err) {
    copyStatus.textContent = err instanceof Error ? err.message : "Copy failed.";
    copyStatus.classList.add("bad");
  }
});
