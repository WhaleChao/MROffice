/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Crop-rect math for the "AI transcribe formula" affordance: map a formula
 * text box (page fractions) onto the rendered pdf.js canvas (device pixels),
 * with padding so ascenders/descenders and radical bars survive the crop.
 * Pure math, unit-testable without a DOM.
 */

export interface FormulaCropRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export function formulaTranscribeCropRect({
  box,
  canvasWidth,
  canvasHeight,
}: {
  /** Page-fraction geometry of the formula box (0..1). */
  box: { x: number; y: number; width: number; height: number; fontSize: number };
  canvasWidth: number;
  canvasHeight: number;
}): FormulaCropRect | null {
  if (canvasWidth <= 0 || canvasHeight <= 0) return null;
  const fontPx = Math.max(4, box.fontSize * canvasHeight);
  const padX = fontPx * 0.6;
  const padY = fontPx * 0.6;
  const left = box.x * canvasWidth - padX;
  const top = box.y * canvasHeight - padY;
  const right = (box.x + box.width) * canvasWidth + padX;
  const bottom = (box.y + box.height) * canvasHeight + padY;
  const x = Math.max(0, Math.floor(left));
  const y = Math.max(0, Math.floor(top));
  const width = Math.min(canvasWidth, Math.ceil(right)) - x;
  const height = Math.min(canvasHeight, Math.ceil(bottom)) - y;
  if (width < 4 || height < 4) return null;
  return { x, y, width, height };
}

/** Strip the data-URL prefix from `canvas.toDataURL()` output. */
export function dataUrlToBase64(dataUrl: string): string | null {
  const comma = dataUrl.indexOf(",");
  if (comma < 0 || !dataUrl.slice(0, comma).includes(";base64")) return null;
  const body = dataUrl.slice(comma + 1);
  return body.length > 0 ? body : null;
}
