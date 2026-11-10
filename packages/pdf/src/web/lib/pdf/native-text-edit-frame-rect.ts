/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Where an Edit Content box's frame is drawn, as page fractions
 * ("x,y,width,height"), set by synchronizePdfNativeTextEditFrame. */
export const PDF_NATIVE_TEXT_EDIT_FRAME_RECT_ATTR = "data-pdf-native-text-edit-frame-rect";

/**
 * The box an Edit Content frame is drawn on, in client px: the stored page
 * fractions on the span's page, or the span's own box when it has none.
 * Presses and the pointer are matched against this, never against where the
 * span itself happens to sit: the frame is what the reader sees, and the
 * frame follows the page's own layout even when a span does not.
 */
export function pdfNativeTextEditFrameClientRect(span: HTMLElement): Readonly<{
  left: number; top: number; right: number; bottom: number;
}> {
  const stored = span.getAttribute(PDF_NATIVE_TEXT_EDIT_FRAME_RECT_ATTR)?.split(",").map(Number);
  const page = span.closest<HTMLElement>(".pdf-viewer__page");
  if (page && stored?.length === 4 && stored.every(Number.isFinite)) {
    const pageRect = page.getBoundingClientRect();
    const [x, y, width, height] = stored;
    return {
      left: pageRect.left + x * pageRect.width,
      top: pageRect.top + y * pageRect.height,
      right: pageRect.left + (x + width) * pageRect.width,
      bottom: pageRect.top + (y + height) * pageRect.height,
    };
  }
  const rect = span.getBoundingClientRect();
  return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
}
