/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Line spacing / paragraph spacing model for presentation text blocks
 * (PowerPoint / WPS "Line Spacing" + "Space Before / After" parity).
 *
 * A text block historically carried only the visual block style (font / colour /
 * align). This module adds an OPTIONAL, migration-safe spacing model — a line
 * spacing MULTIPLE (1.0 / 1.5 / 2.0 …) plus space-before / space-after in POINTS
 * — and the pure conversion that turns it into the `lineHeight` + `marginTop` /
 * `marginBottom` CSS the render edge applies. When a block carries none of these
 * fields it renders exactly as before (no `lineHeight` / margin overrides), so a
 * deck that never set spacing stays byte-for-byte identical.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and shared by the schema, the editor inspector, and the render
 * edge alike.
 */

import { CSS_PX_PER_POINT } from "./presentation/text-units";
import { LINE_SPACING_LADDER } from "../components/office-ribbon/officeRibbonValues";

/**
 * Bounds for the line-spacing MULTIPLE (unitless, e.g. 1.0 = single, 2.0 =
 * double). PowerPoint clamps line spacing well below 10×; we mirror a sane
 * 0.5..10 range so a hand-edited value can never inject a degenerate height.
 */
export const MIN_LINE_SPACING = 0.5;
export const MAX_LINE_SPACING = 10;








export const MIN_STORED_LINE_SPACING = 0.01;

/**
 * Bounds for space-before / space-after, in POINTS (the PowerPoint unit). 0..432
 * pt (== 0..6 inches) matches PowerPoint's accepted range and keeps the derived
 * margin well under any geometry limit.
 */
export const MIN_PARAGRAPH_SPACING_PT = 0;
export const MAX_PARAGRAPH_SPACING_PT = 432;

/** CSS px per point (text-units.ts). */
export const PX_PER_POINT = CSS_PX_PER_POINT;

/** The quick-pick line-spacing multiples the inspector offers (PowerPoint set). */
// Single-sourced from the neutral ribbon vocabulary (see officeRibbonValues);
// numbers here because the spacing engine clamps numerically.
export const LINE_SPACING_PRESETS = LINE_SPACING_LADDER.map(Number);

/**
 * Bounds for LETTER SPACING (character / tracking spacing), in POINTS — the
 * PowerPoint "Character Spacing" unit. PowerPoint allows condensed (negative)
 * through expanded (positive) tracking; ±200pt is far past any real use and keeps
 * the derived `letter-spacing` px well within sane bounds. A value of exactly 0 is
 * "normal" (no override) and is dropped to absent so a never-tracked block stays
 * byte-clean.
 */
export const MIN_LETTER_SPACING_PT = -200;
export const MAX_LETTER_SPACING_PT = 200;

/** The quick-pick letter-spacing presets (points) the inspector offers. */
export const LETTER_SPACING_PRESETS = [-1, -0.5, 0, 0.5, 1, 2, 4] as const;

/**
 * The supported text WRITING-MODE (text direction) values. `horizontal-tb` is
 * the default (left-to-right rows, the historical render); `vertical-rl` is the
 * CJK-common top-to-bottom, right-to-left columns ("竖排"). Only the non-default
 * `vertical-rl` is ever persisted so an existing deck stays byte-for-byte.
 */
export const WRITING_MODES = [
  "horizontal-tb",
  "vertical-rl",
  "rot90",
  "rot270",
  "stacked",
] as const;
export type WritingMode = (typeof WRITING_MODES)[number];
const WRITING_MODE_SET: ReadonlySet<string> = new Set(WRITING_MODES);

/** The spacing fields a text block may carry. All optional + additive. */
export interface ParagraphSpacing {
  /** Line spacing multiple (unitless). Absent → inherit / single. */
  lineSpacing?: number;
  /** Space ABOVE the block, in points. Absent → none. */
  spaceBefore?: number;
  /** Space BELOW the block, in points. Absent → none. */
  spaceAfter?: number;
  /**
   * LETTER (character / tracking) spacing, in POINTS. Absent (or 0) → normal.
   * Projected to CSS `letter-spacing` (px) at the render edge.
   */
  letterSpacing?: number;
  /**
   * Text WRITING-MODE (direction). Absent → `horizontal-tb` (default). Only
   * `vertical-rl` (竖排) is persisted; projected to CSS `writing-mode`.
   */
  writingMode?: WritingMode;
}

/**
 * Round + clamp a line-spacing multiple into the supported range. Non-finite →
 * `undefined` (so the control never persists a degenerate value → inherit).
 * Rounded to 2 decimals so "1.5" stays exactly 1.5 (no float drift). Pure.
 */
export function clampLineSpacing(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  const clamped = Math.min(MAX_LINE_SPACING, Math.max(MIN_LINE_SPACING, value));
  return Math.round(clamped * 100) / 100;
}

/**
 * Round + clamp a paragraph spacing (points) into the supported range, to a
 * hundredth of a point, matching DrawingML spcPts. A non-finite value →
 * `undefined`; an exact 0 is preserved (a
 * caller that wants "remove" passes `undefined`, not 0 — but 0 is still a valid
 * explicit "no gap" the model can hold). Pure.
 */
export function clampParagraphSpacing(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  const clamped = Math.min(
    MAX_PARAGRAPH_SPACING_PT,
    Math.max(MIN_PARAGRAPH_SPACING_PT, value),
  );
  return Math.round(clamped * 100) / 100;
}

/**
 * Round + clamp a letter-spacing (points) into the supported range, to 2
 * decimals. A non-finite value → `undefined` (no override); an exact 0 → `0` is
 * a valid "normal" the model can hold, but {@link normalizeParagraphSpacing}
 * DROPS a 0 so a never-tracked block stays byte-clean. Pure.
 */
export function clampLetterSpacing(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  const clamped = Math.min(
    MAX_LETTER_SPACING_PT,
    Math.max(MIN_LETTER_SPACING_PT, value),
  );
  return Math.round(clamped * 100) / 100;
}

/**
 * Normalize a persisted writing-mode into a known {@link WritingMode}, or
 * `undefined` for the default / unknown (so the field is dropped and the block
 * renders left-to-right, byte-stable). Only `vertical-rl` survives as a real
 * value. Pure.
 */
export function normalizeWritingMode(value: unknown): WritingMode | undefined {
  if (typeof value !== "string" || !WRITING_MODE_SET.has(value)) return undefined;
  // The default `horizontal-tb` is the historical render → drop it to absent.
  return value === "horizontal-tb" ? undefined : (value as WritingMode);
}

/**
 * Project a writing mode to CSS. PPT's five direction modes map onto CSS
 * writing-mode + text-orientation:
 *  - vertical-rl (eaVert: CJK upright, Latin sideways) → vertical-rl + mixed
 *  - rot90  (所有文字旋转90°)                     → vertical-rl + sideways
 *  - rot270 (所有文字旋转270°)                    → sideways-lr (fallback:
 *    vertical-lr + sideways, close enough where sideways-lr is unsupported)
 *  - stacked (堆积: upright letters, one column)  → vertical-lr + upright
 */
export function writingModeCss(
  mode: WritingMode | undefined,
): {
  writingMode?: "vertical-rl" | "vertical-lr" | "sideways-lr";
  textOrientation?: "mixed" | "upright" | "sideways";
} {
  switch (mode) {
    case "vertical-rl":
      return { writingMode: "vertical-rl", textOrientation: "mixed" };
    case "rot90":
      return { writingMode: "vertical-rl", textOrientation: "sideways" };
    case "rot270":
      return { writingMode: "sideways-lr" };
    case "stacked":
      return { writingMode: "vertical-lr", textOrientation: "upright" };
    default:
      return {};
  }
}

/** Convert a points value to integer-ish px (rounded to 2 decimals). Pure. */
export function pointsToPx(points: number): number {
  return Math.round(points * PX_PER_POINT * 100) / 100;
}

/**
 * PowerPoint's Line Spacing > Exactly: "a point value (between 0 and 1584)"
 * (Microsoft Support, "Change text alignment, indentation, and spacing in
 * PowerPoint"); DrawingML `a:lnSpc/a:spcPts`. Mirrors the server's
 * `MAX_PRESENTATION_EXACT_LINE_SPACING_PT`.
 */
export const MAX_EXACT_LINE_SPACING_PT = 1584;

/**
 * A row's exact line height in points, rounded to the hundredth of a point
 * `a:spcPts` states, or `undefined` for anything that is not a finite number
 * in 0..1584 (the row then keeps its multiple). Pure.
 */
export function normalizeExactLineSpacingPoints(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  if (value < 0 || value > MAX_EXACT_LINE_SPACING_PT) return undefined;
  return Math.round(value * 100) / 100;
}












export function exactLineSpacingCss(
  points: number,
  typeSizePt: number | undefined,
): { lineHeight: string; position?: "relative"; top?: string } {
  const lineHeight = `${pointsToPx(points)}px`;
  if (typeSizePt === undefined || !Number.isFinite(typeSizePt) || typeSizePt <= 0) {
    return { lineHeight };
  }
  const shift = pointsToPx(0.25 * points - 0.35 * typeSizePt);
  return shift === 0 ? { lineHeight } : { lineHeight, position: "relative", top: `${shift}px` };
}

/**
 * Normalize an arbitrary persisted value into a clean {@link ParagraphSpacing},
 * or `undefined` when it carries no usable spacing. Each field is independently
 * clamped; a field that clamps to `undefined` (non-finite) is dropped. A result
 * with no surviving field collapses to `undefined` so a block without real
 * spacing stays byte-stable (no phantom keys). Pure.
 */
/**
 * A STORED multiple as the page holds it: rounded like {@link clampLineSpacing}
 * but bounded by the stored span, so loading a page never raises an imported
 * 0.06 to 0.5 -- which the next save would then write into the package.
 */
export function normalizeStoredLineSpacing(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  const clamped = Math.min(MAX_LINE_SPACING, Math.max(MIN_STORED_LINE_SPACING, value));
  return Math.round(clamped * 100) / 100;
}

export function normalizeParagraphSpacing(value: unknown): ParagraphSpacing | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const next: ParagraphSpacing = {};
  const ls = normalizeStoredLineSpacing(record.lineSpacing);
  if (ls !== undefined) next.lineSpacing = ls;
  const sb = clampParagraphSpacing(record.spaceBefore);
  if (sb !== undefined) next.spaceBefore = sb;
  const sa = clampParagraphSpacing(record.spaceAfter);
  if (sa !== undefined) next.spaceAfter = sa;
  // Letter spacing: drop an exact 0 (== normal) so a never-tracked block stays
  // byte-clean — only a real condensed/expanded value is carried.
  const lsp = clampLetterSpacing(record.letterSpacing);
  if (lsp !== undefined && lsp !== 0) next.letterSpacing = lsp;
  const wm = normalizeWritingMode(record.writingMode);
  if (wm !== undefined) next.writingMode = wm;
  return Object.keys(next).length > 0 ? next : undefined;
}

/**
 * Project a (possibly absent) spacing into the CSS the render edge layers onto
 * the text block: `lineHeight` (a unitless multiple) + `marginTop` /
 * `marginBottom` (px). Only keys the block actually sets are returned, so a
 * block without spacing yields an empty object (byte-stable render). Pure;
 * shared by every text-render path so they never drift.
 */
export function spacingToCss(spacing: ParagraphSpacing | undefined): {
  lineHeight?: number;
  marginTop?: number;
  marginBottom?: number;
  letterSpacing?: string;
  writingMode?: "vertical-rl" | "vertical-lr" | "sideways-lr";
  textOrientation?: "mixed" | "upright" | "sideways";
} {
  const normalized = normalizeParagraphSpacing(spacing);
  if (!normalized) return {};
  const css: {
    lineHeight?: number;
    marginTop?: number;
    marginBottom?: number;
    letterSpacing?: string;
    writingMode?: "vertical-rl" | "vertical-lr" | "sideways-lr";
    textOrientation?: "mixed" | "upright" | "sideways";
  } = {};
  if (normalized.lineSpacing !== undefined) css.lineHeight = normalized.lineSpacing;
  if (normalized.spaceBefore !== undefined) css.marginTop = pointsToPx(normalized.spaceBefore);
  if (normalized.spaceAfter !== undefined) css.marginBottom = pointsToPx(normalized.spaceAfter);
  // Letter spacing → CSS `letter-spacing` in px (a unit string so 0-less values
  // like "0.67px" are unambiguous; absent when normal).
  if (normalized.letterSpacing !== undefined) {
    css.letterSpacing = `${pointsToPx(normalized.letterSpacing)}px`;
  }
  if (normalized.writingMode !== undefined) {
    Object.assign(css, writingModeCss(normalized.writingMode));
  }
  return css;
}

/**
 * Read a spacing field off a text-block STYLE object (where the editor stores
 * `lineSpacing` / `spaceBefore` / `spaceAfter` alongside the visual style),
 * returning the normalized value or `undefined`. Pure helper for the inspector
 * to reflect the active value back into its controls.
 */
export function readSpacingFromStyle(
  style: Record<string, unknown> | undefined,
): ParagraphSpacing | undefined {
  if (!style) return undefined;
  return normalizeParagraphSpacing({
    lineSpacing: style.lineSpacing,
    spaceBefore: style.spaceBefore,
    spaceAfter: style.spaceAfter,
    letterSpacing: style.letterSpacing,
    writingMode: style.writingMode,
  });
}
