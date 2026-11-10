/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { gridCanvasFontMetrics, gridCanvasLeadingAbove, gridCanvasNormalLineHeight, gridCanvasTextZoom } from "./grid-canvas-text";

/**
 * Text whose glyphs all come from its stack's first face that draws Latin:
 * Latin letters, general punctuation and currency signs. A line holding
 * anything else may take a fallback face's taller ascent or descent, which
 * the canvas cannot see.
 */
const PRIMARY_FACE_TEXT = /^[ -ɏ -⁯₠-₿]*$/u;

/**
 * An inline box's share of a `line-height: normal` line: its ascent and
 * descent with the font's line gap shared about them as layout shares it
 * (the upper share floored to a whole pixel of the zoomed sheet: 264 of 264
 * faces and sizes in Chromium and WebKit, where an even split was a pixel
 * off for every odd gap).
 */
function normalShare(font: string): { top: number; bottom: number } | null {
  const metrics = gridCanvasFontMetrics(font);
  const normal = gridCanvasNormalLineHeight(font);
  if (!metrics || normal === null) return null;
  const zoom = gridCanvasTextZoom();
  const gap = Math.max(0, Math.round((normal - metrics.ascent - metrics.descent) * zoom));
  const above = gridCanvasLeadingAbove(gap);
  return { top: metrics.ascent + above / zoom, bottom: metrics.descent + (gap - above) / zoom };
}

/**
 * Each line box of a paragraph laid out in `line-height: normal`, as
 * spreadsheetEngineLineBoxes gives it (its top to the baseline and the
 * baseline to its bottom, px), from the faces the browser itself draws in:
 * for a font the text engine has no file of (an installed Calibri, a system
 * fallback). The strut and every run put their ascent and descent, each with
 * half their line gap, about one baseline. Undefined for text a fallback
 * face may draw part of, or a font the canvas cannot measure.
 */
export function canvasLineBoxes(
  lines: readonly (readonly { readonly text: string; readonly fontCss: string }[])[],
  strutCss: string,
): { top: number; bottom: number }[] | undefined {
  const strut = normalShare(strutCss);
  if (!strut) return undefined;
  const boxes: { top: number; bottom: number }[] = [];
  for (const line of lines) {
    let { top, bottom } = strut;
    for (const run of line) {
      if (!PRIMARY_FACE_TEXT.test(run.text)) return undefined;
      const share = normalShare(run.fontCss);
      if (!share) return undefined;
      top = Math.max(top, share.top);
      bottom = Math.max(bottom, share.bottom);
    }
    boxes.push({ top, bottom });
  }
  return boxes;
}
