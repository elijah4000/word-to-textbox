export type CopyMethod = "clipboard-api" | "exec-command";

/**
 * Puts the HTML (with a plain-text alternative) on the clipboard.
 * Falls back to selecting the rendered preview and using the browser's copy
 * command when the async Clipboard API or ClipboardItem is unavailable.
 */
export async function copyRichText(html: string, text: string, previewEl: HTMLElement): Promise<CopyMethod> {
  if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
      return "clipboard-api";
    } catch {
      // Permission denied or unsupported type: fall through to the selection copy.
    }
  }

  const selection = window.getSelection();
  if (!selection) throw new Error("Copy failed: this browser does not allow text selection.");
  const range = document.createRange();
  range.selectNodeContents(previewEl);
  selection.removeAllRanges();
  selection.addRange(range);
  try {
    if (!document.execCommand("copy")) {
      throw new Error("Copy failed. Select the preview and press Ctrl+C (⌘C on Mac).");
    }
  } finally {
    selection.removeAllRanges();
  }
  return "exec-command";
}
