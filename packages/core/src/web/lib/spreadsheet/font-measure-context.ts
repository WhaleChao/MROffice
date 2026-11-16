/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// One canvas context per font a sheet measures in (shared by the grid's
// glyph advances, native-typography, and the canvas text layer's widths and
// metrics, grid-canvas-text). A small module of its own so the canvas text
// layer does not pull in native-typography's dependencies.

/** A measuring context for each font a sheet draws in, kept with that font
 * set. A measurement just after a canvas font changes resolves the face
 * again (a family list with CJK fallbacks: about 30 µs), and cells in
 * different fonts or sizes interleave on every row, so one shared context
 * switching font paid that for each new text (a cold Japanese workbook:
 * about a second of its first renders). Fonts beyond the bound go oldest
 * first. */
const measureContexts = new Map<string, CanvasRenderingContext2D>();
const MAX_MEASURE_CONTEXTS = 48;
export function fontMeasureContext(font: string, kerning: CanvasFontKerning = "auto"): CanvasRenderingContext2D | null {
  const key = `${kerning}\u0000${font}`;
  const known = measureContexts.get(key);
  if (known) return known;
  if (typeof document === "undefined") return null;
  const context = document.createElement("canvas").getContext("2d");
  if (!context) return null;
  context.font = font;
  context.fontKerning = kerning;
  if (measureContexts.size >= MAX_MEASURE_CONTEXTS) {
    const oldest = measureContexts.keys().next().value;
    if (oldest !== undefined) measureContexts.delete(oldest);
  }
  measureContexts.set(key, context);
  return context;
}
