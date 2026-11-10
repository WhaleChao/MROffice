/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Underline and strikethrough on part of a text run (WPS: two characters of
 * a line underlined change only those two). A run carries at most one MRAI
 * rule of each style, drawn beside its text object (text-decoration.ts), and
 * that rule may run under the whole run or under one stretch of it. A
 * stretch is kept as the share of the run's advance before its start and up
 * to its end (`from`, `to`, 0..1), so it rides along when an edit moves,
 * resizes or reflows the run (list-text-fidelity.ts, textRulesBesideRun).
 *
 * Pure geometry, browser-safe: the dialog, the window and the planner all
 * read it.
 */
import type { PdfTextSourceAppearance } from "./source-text-appearance";

/** Where the writer draws a rule, off the baseline, and how thick, in ems of
 * the run's size (text-rule-block); a draft shows it there too. */
export const PDF_NATIVE_UNDERLINE_OFFSET_EM = -0.12;
export const PDF_NATIVE_STRIKETHROUGH_OFFSET_EM = 0.26;
export const PDF_NATIVE_RULE_THICKNESS_EM = 0.06;

export type PdfNativeTextRuleRange = Readonly<{ from: number; to: number }>;

/** What one press of Underline or Strikethrough asks for: the characters
 * selected (`part`, offsets into the run's text; the whole run without), or
 * the run's rule taken away whatever stretch it runs under (`remove`, as
 * 清除格式 does). */
export type PdfNativeTextDecorationToggleOptions = Readonly<{
  part?: Readonly<{ start: number; end: number }>;
  remove?: boolean;
  

  quiet?: boolean;
}>;

/** Shares closer than this are one place (a rule's numbers carry six
 * decimals; one character of a long line is still far wider). */
export const PDF_NATIVE_TEXT_RULE_RANGE_EPSILON = 1e-3;

/** The whole run. */
export const PDF_NATIVE_TEXT_RULE_WHOLE: PdfNativeTextRuleRange = Object.freeze({ from: 0, to: 1 });
const WHOLE = PDF_NATIVE_TEXT_RULE_WHOLE;

/**
 * The rules a run carries, as the editor shows them: `underline` and
 * `strikethrough` for a rule under the whole run (the buttons pressed), and
 * the stretch of a rule under part of it only.
 */
export type PdfNativeTextRuleDecorations = Readonly<{
  underline: boolean;
  strikethrough: boolean;
  underlineRange?: PdfNativeTextRuleRange;
  strikethroughRange?: PdfNativeTextRuleRange;
}>;

export function pdfNativeTextRuleDecorations(rules: Readonly<{
  underline: PdfNativeTextRuleRange | null;
  strikethrough: PdfNativeTextRuleRange | null;
}>): PdfNativeTextRuleDecorations {
  const whole = (range: PdfNativeTextRuleRange | null) => Boolean(range && pdfNativeTextRuleRangeIsWhole(range));
  return Object.freeze({
    underline: whole(rules.underline),
    strikethrough: whole(rules.strikethrough),
    ...(rules.underline && !whole(rules.underline) ? { underlineRange: rules.underline } : {}),
    ...(rules.strikethrough && !whole(rules.strikethrough) ? { strikethroughRange: rules.strikethrough } : {}),
  });
}

/** The stretch the run's rule of `style` runs under, or null with none. */
export function pdfNativeTextRuleRangeOf(
  decorations: PdfNativeTextRuleDecorations | null | undefined,
  style: "underline" | "strikethrough",
): PdfNativeTextRuleRange | null {
  if (!decorations) return null;
  if (decorations[style]) return WHOLE;
  return (style === "underline" ? decorations.underlineRange : decorations.strikethroughRange) ?? null;
}

export function pdfNativeTextRuleRangeIsWhole(range: PdfNativeTextRuleRange): boolean {
  return range.from <= PDF_NATIVE_TEXT_RULE_RANGE_EPSILON &&
    range.to >= 1 - PDF_NATIVE_TEXT_RULE_RANGE_EPSILON;
}

/**
 * The stretch of the run the characters `[start, end)` of `text` take, read
 * from where the reader draws them: the share of the run's drawn length
 * before the first and up to the last (min and max along the baseline, so
 * a right-to-left stretch reads the same). Null when the reader's characters
 * are not those of `text` or leave part of the stretch unread.
 */
export function pdfNativeTextRuleRangeOfText(
  appearance: PdfTextSourceAppearance | null | undefined,
  text: string,
  range: Readonly<{ start: number; end: number }>,
): PdfNativeTextRuleRange | null {
  if (!appearance || appearance.text !== text || !(range.end > range.start)) return null;
  const characters = appearance.characters;
  const direction = characters[0]?.direction;
  if (!direction || direction.length !== 2) return null;
  const [dx, dy] = direction;
  if (!(Math.hypot(dx, dy) > 0)) return null;
  const along = (x: number, y: number) => x * dx + y * dy;
  let runLow = Infinity, runHigh = -Infinity, low = Infinity, high = -Infinity, covered = 0;
  for (const character of characters) {
    const quad = character.quad;
    if (quad.length !== 8 || quad.some((value) => !Number.isFinite(value))) return null;
    // Upper-left and upper-right: where the glyph's advance starts and ends.
    const from = along(quad[0], quad[1]);
    const to = along(quad[2], quad[3]);
    runLow = Math.min(runLow, from, to);
    runHigh = Math.max(runHigh, from, to);
    if (character.start >= range.start && character.end <= range.end) {
      low = Math.min(low, from, to);
      high = Math.max(high, from, to);
      covered += character.end - character.start;
    }
  }
  const length = runHigh - runLow;
  if (!(length > 0) || !(high > low) || covered !== range.end - range.start) return null;
  const share = (value: number) => Math.min(1, Math.max(0, (value - runLow) / length));
  const result = { from: share(low), to: share(high) };
  return pdfNativeTextRuleRangeIsWhole(result) ? WHOLE : Object.freeze(result);
}

/** Whether the rule standing on the run already runs under all of `target`. */
export function pdfNativeTextRuleRangeCovers(
  standing: PdfNativeTextRuleRange | null | undefined,
  target: PdfNativeTextRuleRange,
): boolean {
  const epsilon = PDF_NATIVE_TEXT_RULE_RANGE_EPSILON;
  return Boolean(standing && standing.from <= target.from + epsilon && target.to <= standing.to + epsilon);
}

/**
 * What one press of Underline (or Strikethrough) on `target` leaves on the
 * run, as Word and WPS toggle a character format: a stretch the rule already
 * covers loses it, any other stretch gains it (joined to the rule standing).
 * `"none"` takes the rule away; null is a result one rule cannot draw (two
 * stretches apart), which the caller refuses with its reason.
 */
export function pdfNativeTextRuleToggle(
  standing: PdfNativeTextRuleRange | null | undefined,
  target: PdfNativeTextRuleRange,
): PdfNativeTextRuleRange | "none" | null {
  const epsilon = PDF_NATIVE_TEXT_RULE_RANGE_EPSILON;
  const settle = (range: PdfNativeTextRuleRange): PdfNativeTextRuleRange =>
    pdfNativeTextRuleRangeIsWhole(range) ? WHOLE : Object.freeze({ from: range.from, to: range.to });
  if (standing && pdfNativeTextRuleRangeCovers(standing, target)) {
    const pieces = [
      { from: standing.from, to: target.from },
      { from: target.to, to: standing.to },
    ].filter((piece) => piece.to - piece.from > epsilon);
    if (pieces.length === 0) return "none";
    return pieces.length === 1 ? settle(pieces[0]) : null;
  }
  if (!standing) return settle(target);
  // Joined to the rule standing when the two touch; two stretches apart need
  // two rules, which a run does not carry.
  if (standing.from <= target.to + epsilon && target.from <= standing.to + epsilon) {
    return settle({ from: Math.min(standing.from, target.from), to: Math.max(standing.to, target.to) });
  }
  return null;
}

/**
 * The geometry of a rule under `range` of a run whose whole rule `geometry`
 * draws: the run's own matrix moved along the run to the stretch's start,
 * and the stretch's share of the advance. The whole run's geometry itself
 * for the whole run, so its rule keeps its exact bytes.
 */
export function pdfNativeTextRuleGeometryOfRange<
  G extends Readonly<{ textMatrix: readonly [number, number, number, number, number, number]; advance: number }>,
>(geometry: G, range: PdfNativeTextRuleRange): G {
  if (pdfNativeTextRuleRangeIsWhole(range)) return geometry;
  const [a, b, c, d, e, f] = geometry.textMatrix;
  const offset = geometry.advance * range.from;
  return Object.freeze({
    ...geometry,
    textMatrix: Object.freeze([a, b, c, d, e + a * offset, f + b * offset]),
    advance: geometry.advance * (range.to - range.from),
  }) as unknown as G;
}
