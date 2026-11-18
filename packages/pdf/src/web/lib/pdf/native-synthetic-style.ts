/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bold and italic for a run whose font has no such face, the way Word and
 * WPS fake them in place: bold by stroking the glyph outlines in the fill
 * colour (`2 Tr` under a stroke width), italic by slanting the run's text
 * matrix. Pure arithmetic, read by the editor's ribbon and by the planner.
 */

/** The slant of a synthetic italic: tan 12°, the angle of the standard
 * oblique faces (Helvetica-Oblique, Courier-Oblique). */
export const PDF_NATIVE_SYNTHETIC_ITALIC_SKEW = 0.21256;

/** The stroke width of a synthetic bold, as a fraction of the font size:
 * a stem grows by this much, which takes a regular face about to the weight
 * of its bold. */
export const PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM = 0.035;

type Linear = readonly [number, number, number, number];

const round6 = (value: number): number => {
  const rounded = Number(value.toFixed(6));
  return Object.is(rounded, -0) ? 0 : rounded;
};

/** How far the matrix's second axis leans along its first: 0 when the two
 * are perpendicular (an upright run, turned or not), the skew of a slanted
 * one. */
export function pdfNativeTextMatrixSkew(linear: Linear): number {
  const [a, b, c, d] = linear;
  const first = a * a + b * b;
  if (!(first > 0)) return 0;
  return (a * c + b * d) / first;
}

/** Whether the run is drawn under the synthetic italic slant. */
export function pdfNativeTextMatrixIsSlanted(linear: Linear): boolean {
  return Math.abs(pdfNativeTextMatrixSkew(linear) - PDF_NATIVE_SYNTHETIC_ITALIC_SKEW) <= 0.02;
}









export function pdfNativeDrawnRunIsSlanted(
  run: Readonly<{ textMatrix: readonly number[]; graphicsMatrix: readonly number[] }>,
): boolean {
  const [ta, tb, tc, td] = run.textMatrix, [ga, gb, gc, gd] = run.graphicsMatrix;
  if (![ta, tb, tc, td, ga, gb, gc, gd].every(Number.isFinite)) return false;
  return pdfNativeTextMatrixIsSlanted([ta * ga + tb * gc, ta * gb + tb * gd, tc * ga + td * gc, tc * gb + td * gd]);
}

/** The run's matrix with the synthetic slant put on or taken off: the
 * second axis leans along the first by the skew, the first axis and the
 * origin stay, so the baseline and the run's start do not move. */
export function pdfNativeSlantedTextMatrixLinear(linear: Linear, slanted: boolean): Linear {
  const [a, b, c, d] = linear;
  const current = pdfNativeTextMatrixSkew(linear);
  const wanted = slanted ? PDF_NATIVE_SYNTHETIC_ITALIC_SKEW : 0;
  const delta = wanted - current;
  return Object.freeze([a, b, round6(c + delta * a), round6(d + delta * b)]) as unknown as Linear;
}

/** Whether `replacement` is `original` with only its slant changed: the same
 * first axis, the second moved along it. A move of nothing but the slant
 * keeps the run's origin, where a turn pivots about its centre. */
export function pdfNativeTextMatrixSlantOnly(original: Linear, replacement: Linear): boolean {
  const [a, b, c, d] = original;
  const [ra, rb, rc, rd] = replacement;
  if (Math.abs(a - ra) > 1e-9 || Math.abs(b - rb) > 1e-9) return false;
  if (Math.abs(c - rc) < 1e-9 && Math.abs(d - rd) < 1e-9) return false;
  // (rc - c, rd - d) is parallel to (a, b): their cross product vanishes.
  return Math.abs((rc - c) * b - (rd - d) * a) <= 1e-6 * Math.max(1, Math.hypot(a, b));
}

/** The `w` a synthetic bold strokes its glyphs under, in the user space the
 * text matrix maps the run into (the stroke is laid down after the matrix,
 * so the width scales with it as the glyphs do). */
export function pdfNativeSyntheticBoldStrokeWidth(fontSize: number, linear: Linear): number {
  const [a, b] = linear;
  return round6(PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM * fontSize * Math.hypot(a, b));
}

/** Whether a bold can be drawn in place on a run in this state: the stroke
 * width in force is known, so the run after it is stroked under the width it
 * was, and so is the stroke colour in force (or the run already strokes its
 * glyphs in its own colour), so the paths drawn after it keep their colour. */
export function pdfNativeSyntheticBoldDrawable(
  state: Readonly<{ lineWidth?: number; renderingMode?: number; strokeColor?: unknown }>,
): boolean {
  return state.lineWidth !== undefined && (state.renderingMode === 2 || state.strokeColor !== undefined);
}
