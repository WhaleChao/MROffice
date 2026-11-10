/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One text size for a whole SmartArt graphic, as Office sizes it.
 *
 * Microsoft Support ("Learn more about SmartArt graphics"): "When the font size
 * in one shape shrinks because you add more text to the shape, all of the
 * other text in the remaining shapes of the SmartArt graphic will also shrink
 * to the same size to keep the SmartArt graphic looking consistent". A layout
 * states a largest size for its text and Office's autofit takes it down until
 * every shape's words fit; a shape never breaks a word to make room, it
 * shrinks instead (the layout definitions' `primFontSz` constraint, shared by
 * every text shape with `op="equ"`, with the shrink rule of their `ruleLst`).
 *
 * The three hosts draw a label in different boxes (a slide's label block, a
 * sheet shape's padded rectangle, a Docs member's inset), so each one hands
 * its own boxes to {@link smartArtTextFitScale}; the wrap model, the word rule
 * and the one-scale-for-all rule live here once. Pure and deterministic: the
 * glyph advance is an estimate, as in `lib/presentation-autofit` (whose line
 * height and floor this reuses), so a test and a render agree to the pixel.
 */

import {
  DEFAULT_LINE_HEIGHT,
  MIN_AUTOFIT_SCALE,
} from "../presentation-autofit";

/** One label: its words, the box they are laid out in, and its nominal size. */
export interface SmartArtTextFitLabel {
  readonly text: string;
  /** Inner width available to the words (the host's own text box). */
  readonly width: number;
  /** Inner height available to the words. */
  readonly height: number;
  /** The label's nominal size, in the same unit as the box. */
  readonly fontSize: number;
  /**
   * The line-height multiple THIS label is drawn with, when it is not the
   * graphic's ({@link SmartArtTextFitOptions.lineHeight}): an item whose words
   * the author spaced (`a:lnSpc`) stands its lines at that spacing, and the
   * one size the graphic shares is the one those taller lines still hold.
   */
  readonly lineHeight?: number;
}

export interface SmartArtTextFitOptions {
  /** Line-height multiple the host draws the label with. */
  readonly lineHeight?: number;
  /** Whether the host draws the label bold (a wider advance). */
  readonly bold?: boolean;
  /** Smallest scale the fit may reach; below it the text stays and clips. */
  readonly minScale?: number;
}

const CJK =
  /[ᄀ-ᇿ⺀-〿぀-ヿ㄀-ㇿ㐀-䶿一-鿿ꥠ-꥿가-힯豈-﫿︰-﹏＀-￯]/u;
const NARROW = /[ijlI.,:;'!|`]/;

/** One glyph's advance, in em. */
function advanceEm(char: string): number {
  if (char === " " || char === "\t") return 0.28;
  if (CJK.test(char)) return 1;
  if (NARROW.test(char)) return 0.3;
  if (/[A-Z]/.test(char)) return 0.66;
  if (/[mwMW]/.test(char)) return 0.82;
  if (/[0-9]/.test(char)) return 0.56;
  if (/[a-z]/.test(char)) return 0.53;
  return 0.6;
}

/**
 * The pieces a line may break between: every CJK character on its own (East
 * Asian text breaks between any two), a run of anything else up to and with
 * its trailing spaces (a Latin word, which never breaks).
 */
function breakUnits(paragraph: string): { width: number; trailing: number }[] {
  const units: { width: number; trailing: number }[] = [];
  let word = 0;
  let spaces = 0;
  const flush = () => {
    if (word > 0 || spaces > 0) units.push({ width: word, trailing: spaces });
    word = 0;
    spaces = 0;
  };
  for (const char of Array.from(paragraph)) {
    if (char === " " || char === "\t") {
      spaces += advanceEm(char);
      continue;
    }
    if (spaces > 0) flush();
    if (CJK.test(char)) {
      flush();
      units.push({ width: 1, trailing: 0 });
      continue;
    }
    word += advanceEm(char);
  }
  flush();
  return units;
}

/**
 * How many lines `text` takes at `fontSize` in a box `width` wide, or `null`
 * when some word is wider than the box by itself (Office shrinks rather than
 * breaking it).
 */
export function smartArtWrappedLineCount(
  text: string,
  fontSize: number,
  width: number,
  bold = false,
): number | null {
  const scale = fontSize * (bold ? 1.06 : 1);
  let lines = 0;
  for (const paragraph of text.split("\n")) {
    let used = 0;
    let pending = 0;
    let started = false;
    lines += 1;
    for (const unit of breakUnits(paragraph)) {
      const unitWidth = unit.width * scale;
      if (unitWidth > width + 0.01) return null;
      if (started && used + pending + unitWidth > width + 0.01) {
        lines += 1;
        used = 0;
        pending = 0;
      }
      used += pending + unitWidth;
      pending = unit.trailing * scale;
      started = true;
    }
  }
  return Math.max(1, lines);
}

function labelFits(
  label: SmartArtTextFitLabel,
  scale: number,
  lineHeight: number,
  bold: boolean,
): boolean {
  const size = label.fontSize * scale;
  const lines = smartArtWrappedLineCount(label.text, size, label.width, bold);
  const own = label.lineHeight !== undefined && Number.isFinite(label.lineHeight) && label.lineHeight > 0 ? label.lineHeight : lineHeight;
  return lines !== null && lines * size * own <= label.height + 0.01;
}

/**
 * The ONE scale every label of a graphic is drawn at: the largest in
 * `[minScale, 1]` at which each label's words fit its box without breaking a
 * word, or `1` when they already fit. Labels with no words, or no box, do
 * not vote.
 */
export function smartArtTextFitScale(
  labels: readonly SmartArtTextFitLabel[],
  options: SmartArtTextFitOptions = {},
): number {
  const lineHeight =
    options.lineHeight && options.lineHeight > 0
      ? options.lineHeight
      : DEFAULT_LINE_HEIGHT;
  const bold = options.bold === true;
  const floor = Math.min(
    1,
    Math.max(0.01, options.minScale ?? MIN_AUTOFIT_SCALE),
  );
  const voting = labels.filter(
    (label) =>
      label.text.trim().length > 0 &&
      Number.isFinite(label.width) &&
      Number.isFinite(label.height) &&
      Number.isFinite(label.fontSize) &&
      label.width > 0 &&
      label.height > 0 &&
      label.fontSize > 0,
  );
  const fitsAt = (scale: number) =>
    voting.every((label) => labelFits(label, scale, lineHeight, bold));
  if (voting.length === 0 || fitsAt(1)) return 1;
  if (!fitsAt(floor)) return floor;
  let fit = floor;
  let over = 1;
  for (let step = 0; step < 20; step += 1) {
    const mid = (fit + over) / 2;
    if (fitsAt(mid)) fit = mid;
    else over = mid;
  }
  return Math.floor(fit * 1000) / 1000;
}
