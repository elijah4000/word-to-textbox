import "./styles.css";
import { copyRichText } from "./lib/clipboard";
import { DEFAULT_OPTIONS, type FormatOptions } from "./lib/format";
import { loadDocx, produceOutput, type LoadedDocument, type Output } from "./lib/pipeline";
import type { DiffOp, TextComparison } from "./lib/textCheck";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const fileInput = $<HTMLInputElement>("file");
const drop = $("drop");
const dropTitle = $("drop-title");
const loadError = $("load-error");
const result = $("result");
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
const attention = $<HTMLDetailsElement>("attention");
const attentionSummary = $("attention-summary");
const issuesList = $("issues");
const messagesBox = $("messages-box");
const messagesCount = $("messages-count");
const messagesList = $("messages");
const checkPanel = $("check-panel");
const checkDiff = $("check-diff");
const preview = $("preview");

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

const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

const LARGE_IMAGES = 1024 * 1024;
const CONTEXT_WORDS = 8;

async function handleFile(file: File) {
  loaded = null;
  output = null;
  loadError.hidden = true;
  dropTitle.textContent = file.name;
  drop.classList.add("loaded");
  copyStatus.textContent = "";
  ack.checked = false;
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
  render();
}

function showLoadError(message: string) {
  loadError.textContent = message;
  loadError.hidden = false;
  result.hidden = true;
}

function render() {
  if (!loaded) return;
  output = produceOutput(loaded, readOptions());
  result.hidden = false;

  // Option hints
  const h = output.formatted.hangingParagraphs;
  hangingCount.textContent = optHanging.checked && h > 0 ? `(${plural(h, "paragraph")})` : "";
  if (loaded.scan.images === 0) {
    imageSize.textContent = "";
  } else if (optImages.checked) {
    imageSize.textContent = `(${plural(output.formatted.images, "image")}, ${formatBytes(output.formatted.imageBytes)})`;
  } else {
    imageSize.textContent = `(${plural(loaded.scan.images, "image")} left out)`;
  }

  renderIssues(output);
  renderTextCheck(output);

  preview.innerHTML = output.formatted.html;

  const needsAck = !output.textCheck.identical || output.issues.some((i) => i.severity === "error");
  ackRow.hidden = !needsAck;
  copyBtn.disabled = needsAck && !ack.checked;
}

function renderIssues(out: Output) {
  const all = [...out.issues];
  if (out.formatted.images > 0 && out.formatted.imageBytes > LARGE_IMAGES) {
    all.push({
      id: "large-images",
      severity: "warning",
      label: "Large images",
      count: out.formatted.images,
      detail: `Images add about ${formatBytes(out.formatted.imageBytes)} to the paste. Some text boxes refuse or strip large pasted content; if so, turn off "Include images" and insert them with the editor's image button.`,
    });
  }

  issuesList.replaceChildren(
    ...all.map((issue) => {
      const li = el("li", undefined, `issue ${issue.severity}`);
      li.append(el("b", `${issue.label} × ${issue.count}`), el("p", issue.detail));
      return li;
    }),
  );

  const messages = loaded?.conversion.messages ?? [];
  messagesBox.hidden = messages.length === 0;
  messagesCount.textContent = String(messages.length);
  messagesList.replaceChildren(...messages.map((m) => el("li", m)));

  // Hidden when there is nothing to say; open only when something needs action.
  const serious = all.filter((i) => i.severity !== "info");
  attention.hidden = all.length === 0;
  attention.classList.toggle("has-issues", serious.length > 0);
  attention.open = serious.length > 0;
  attentionSummary.textContent =
    serious.length > 0
      ? `Needs attention: ${serious.map((i) => i.label.toLowerCase()).join(", ")}`
      : `Notes: ${all.map((i) => `${i.label.toLowerCase()} ×${i.count}`).join(", ")}`;
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

// The text check always runs; it is only shown when the text differs.
function renderTextCheck(out: Output) {
  const tc = out.textCheck;
  checkPanel.hidden = tc.identical;
  checkDiff.replaceChildren();
  if (tc.identical) return;
  renderComparison("Main text", tc.body);
  for (const note of tc.notes) renderComparison(`Note ${note.id}`, note.comparison);
  if (tc.missingNotes.length) checkDiff.append(el("p", `Notes missing from the copy: ${tc.missingNotes.join(", ")}`, "diff-title"));
  if (tc.extraNotes.length) checkDiff.append(el("p", `Unexpected notes in the copy: ${tc.extraNotes.join(", ")}`, "diff-title"));
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
    await copyRichText(output.formatted.html, output.plainText, preview);
    copyStatus.textContent = "Copied ✓ Now paste into the text box.";
    copyStatus.classList.add("ok");
  } catch (err) {
    copyStatus.textContent = err instanceof Error ? err.message : "Copy failed.";
    copyStatus.classList.add("bad");
  }
});
