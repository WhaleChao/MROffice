/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The rotation the canvas's bitmap was painted at, as its painter noted it. */
export const PDF_PAINTED_ROTATION_ATTR = "data-pdf-painted-rotation";
/** Set while the bitmap is shown turned ahead of its new drawing. */
export const PDF_TURNED_AHEAD_ATTR = "data-pdf-turned-ahead";

/** Note the rotation a page canvas's bitmap has just been painted at: its
 * box is the page's, upright for that rotation. */
export function notePdfCanvasPainted(canvas: HTMLCanvasElement, rotation: number,
  box: Readonly<{ width: number; height: number }>): void {
  canvas.setAttribute(PDF_PAINTED_ROTATION_ATTR, String(normalized(rotation)));
  clearTurn(canvas);
  canvas.style.width = `${box.width}px`;
  canvas.style.height = `${box.height}px`;
}

/**
 * Fit a page canvas's painted bitmap to the page's box (CSS pixels) at the
 * rotation the page now has: its own size where it was painted at that
 * rotation (a zoom stretches it, as before), else turned by the difference
 * into the box. Returns the quarter turns shown ahead (0: none).
 */
export function fitPdfCanvasToPageBox(canvas: HTMLCanvasElement, box: Readonly<{ width: number; height: number }>,
  rotation: number): number {
  const painted = Number(canvas.getAttribute(PDF_PAINTED_ROTATION_ATTR));
  const delta = Number.isFinite(painted) && canvas.hasAttribute(PDF_PAINTED_ROTATION_ATTR)
    ? normalized(rotation - painted) : 0;
  if (delta === 0) {
    clearTurn(canvas);
    canvas.style.width = `${box.width}px`;
    canvas.style.height = `${box.height}px`;
    return 0;
  }
  // The bitmap's box before it is turned: the page's, quarter-turned back.
  const [width, height] = delta === 180 ? [box.width, box.height] : [box.height, box.width];
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  canvas.style.transformOrigin = "0 0";
  // Turned about its corner, then moved back into the page's box.
  canvas.style.transform = delta === 90 ? `translate(${box.width}px, 0px) rotate(90deg)`
    : delta === 180 ? `translate(${box.width}px, ${box.height}px) rotate(180deg)`
      : `translate(0px, ${box.height}px) rotate(270deg)`;
  canvas.setAttribute(PDF_TURNED_AHEAD_ATTR, String(delta));
  return delta / 90;
}

function clearTurn(canvas: HTMLCanvasElement): void {
  if (!canvas.hasAttribute(PDF_TURNED_AHEAD_ATTR)) return;
  canvas.removeAttribute(PDF_TURNED_AHEAD_ATTR);
  canvas.style.transform = "";
  canvas.style.transformOrigin = "";
}

function normalized(degrees: number): number {
  return ((Math.round(degrees / 90) * 90) % 360 + 360) % 360;
}
