/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The point size a PDF run is drawn at on paper.
 *
 * A show's glyphs are `Tf` units of text space tall, and text space reaches
 * the page through the text matrix (`Tm`) and the current transformation
 * matrix (`cm`): the size on paper is `Tf` times how long the text-space y
 * unit is once both have mapped it (ISO 32000-1 9.4.4). Word for Mac writes
 * every run as `1 Tf` under a `150 0 0 150 Tm` inside `0.24 0 0 0.24 cm`:
 * 1 x 150 x 0.24 = 36 pt. Reading `Tf` alone gave 1, `Tf x Tm` 150.
 */

type Linear = readonly number[];

/** How far the text-space y unit reaches on the page (Tm, then the CTM). */
export function pdfTextVerticalScale(textMatrix: Linear, graphicsMatrix?: Linear | null): number {
  const c = Number(textMatrix[2]);
  const d = Number(textMatrix[3]);
  if (!Number.isFinite(c) || !Number.isFinite(d)) return 1;
  let x = c;
  let y = d;
  if (graphicsMatrix && graphicsMatrix.length >= 4) {
    const [a, b, cc, dd] = graphicsMatrix.map(Number);
    if ([a, b, cc, dd].every(Number.isFinite)) {
      x = c * a + d * cc;
      y = c * b + d * dd;
    }
  }
  const scale = Math.hypot(x, y);
  return Number.isFinite(scale) && scale > 1e-9 ? scale : 1;
}

/** A run's size on paper in points, to the hundredth. */
export function pdfTextPointSize(fontSize: number, textMatrix: Linear, graphicsMatrix?: Linear | null): number {
  return Math.round(fontSize * pdfTextVerticalScale(textMatrix, graphicsMatrix) * 100) / 100;
}

/**
 * The one size the visible runs of a selection share, or null when they
 * differ. A run of white space only (Word ends each line with a space in the
 * paragraph mark's own font and size) draws nothing and does not count.
 */
export function pdfUniformTextPointSize(
  runs: readonly Readonly<{ text: string; fontSize: number; textMatrix: Linear; graphicsMatrix?: Linear | null }>[],
): number | null {
  let size: number | null = null;
  for (const run of runs) {
    if (!run.text.trim()) continue;
    const own = pdfTextPointSize(run.fontSize, run.textMatrix, run.graphicsMatrix);
    if (!(own > 0)) return null;
    if (size === null) size = own;
    else if (Math.abs(size - own) > 0.05) return null;
  }
  return size;
}





export function pdfTextPointSizeOver(
  runs: readonly Readonly<{ text: string; fontSize: number; textMatrix: Linear; graphicsMatrix?: Linear | null }>[],
  range: Readonly<{ start: number; end: number }>,
): number | null {
  if (range.end <= range.start) return null;
  let offset = 0;
  let size: number | null = null;
  for (const run of runs) {
    const start = offset, end = offset + run.text.length;
    offset = end;
    if (end <= range.start || start >= range.end) continue;
    const part = run.text.slice(Math.max(0, range.start - start), Math.min(run.text.length, range.end - start));
    if (!part.trim()) continue;
    const own = pdfTextPointSize(run.fontSize, run.textMatrix, run.graphicsMatrix);
    if (!(own > 0) || (size !== null && Math.abs(size - own) > 0.05)) return null;
    size = own;
  }
  return size === null ? null : Number(size.toFixed(2));
}
