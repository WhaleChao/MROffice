/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const COVER_SELECTOR = [
  'canvas[data-pdf-native-text-editing-surface="true"]',
  'canvas[data-pdf-native-text-candidate-preview="true"]',
  'canvas[data-pdf-native-text-committed-preview="true"]',
  "canvas[data-pdf-native-text-source-handoff]",
].join(", ");

/** Whether an inline edit's canvas still stands over the page. */
export function pdfPageCanvasCovered(host: ParentNode | null | undefined): boolean {
  if (!host) return false;
  return Array.from(host.querySelectorAll<HTMLCanvasElement>(COVER_SELECTOR)).some(
    (cover) => cover.isConnected && cover.width > 0 && cover.height > 0 && cover.style.visibility !== "hidden",
  );
}

/** Show a page canvas an inline edit left hidden when nothing covers it any more. */
export function revealUncoveredPdfPageCanvas(canvas: HTMLCanvasElement): boolean {
  if (canvas.style.visibility !== "hidden" || pdfPageCanvasCovered(canvas.parentElement)) return false;
  canvas.style.visibility = "";
  return true;
}
