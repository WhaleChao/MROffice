/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text spill (Excel / WPS 文字溢出).
 *
 * A cell whose text is wider than its column does not get cut off in a real
 * spreadsheet: as long as the neighbouring cells are empty, the text runs over
 * them. It is the single most visible thing about how a spreadsheet renders —
 * a timetable whose "Statistics of Measurement Lecture" reads "Statistics of"
 * looks broken, and every imported report is full of exactly that shape.
 *
 * The rule has an edge people rely on: the moment a neighbour holds anything,
 * the spill stops dead at that boundary rather than drawing over it. That is
 * how a sheet stays readable when a column fills up.
 *
 * Direction follows alignment, as it does in Excel: left-aligned text runs
 * right, right-aligned runs left, centred runs both ways. Wrapped text never
 * spills — it grew the row instead.
 */

export type SpillAlign = "left" | "right" | "center";

export interface SpillInput {
  row: number;
  col: number;
  /** Horizontal alignment as resolved for the cell (not the raw style). */
  align: SpillAlign;
  columnCount: number;
  /** Whether a cell blocks a spill: it holds content, or is covered by a merge. */
  isOccupied: (row: number, col: number) => boolean;
  /**
   * How far a single cell may run. Excel has no explicit limit, but an
   * unbounded scan across a 1,000-column grid is a per-cell cost paid on every
   * render, so the walk stops well before it could matter visually.
   */
  maxColumns?: number;
}

export interface SpillRange {
  /** First column the text may paint over, inclusive. */
  startCol: number;
  /** Last column the text may paint over, inclusive. */
  endCol: number;
}

const DEFAULT_MAX_COLUMNS = 40;

/**
 * The inclusive column range a cell's text may occupy.
 *
 * Always contains the cell's own column, so a caller can use the result
 * unconditionally; `startCol === endCol` means "no room to spill".
 */
export function spillRange(input: SpillInput): SpillRange {
  const { row, col, align, columnCount, isOccupied } = input;
  const budget = Math.max(0, input.maxColumns ?? DEFAULT_MAX_COLUMNS);
  let startCol = col;
  let endCol = col;

  const canRunRight = align === "left" || align === "center";
  const canRunLeft = align === "right" || align === "center";

  if (canRunRight) {
    for (let next = col + 1; next < columnCount && next - col <= budget; next += 1) {
      if (isOccupied(row, next)) break;
      endCol = next;
    }
  }
  if (canRunLeft) {
    for (let next = col - 1; next >= 0 && col - next <= budget; next -= 1) {
      if (isOccupied(row, next)) break;
      startCol = next;
    }
  }
  return { startCol, endCol };
}

export interface SpillDecisionInput extends SpillInput {
  /** The text as displayed. Nothing to spill when it is empty. */
  text: string;
  /** Wrapped text grew its row instead; it never runs over a neighbour. */
  wrapText: boolean;
  /** Shrink to fit draws the text smaller inside its own cell (Excel: the
   * size is reduced until the text fits the column); it never runs over a
   * neighbour. Spilled, it was drawn full size across the next cell and
   * shrank only while the cell was open for editing. */
  shrinkToFit?: boolean;
  /** A cell being edited shows its own editor, never a spill overlay. */
  active: boolean;
  /** A merged cell already spans its region; it does not also spill. */
  merged: boolean;
  /** The text's rendered width, px (measured, or estimated where no canvas); a
   * function is called only when the text could run over a neighbour at all. */
  textWidthPx: number | (() => number);
  /** The cell's usable content width, px (spanned width minus padding). */
  cellWidthPx: number;
}

/**
 * The range to paint, or null when this cell must not spill at all.
 *
 * Returning null for "fits anyway" is deliberate: the caller renders an overlay
 * only when there is something to show, so an ordinary sheet pays nothing.
 */
export function spillDecision(input: SpillDecisionInput): SpillRange | null {
  if (input.active || input.wrapText || input.shrinkToFit || input.merged) return null;
  if (!input.text || input.text.includes("\n")) return null;
  // The fits check. This is the line the doc-comment always promised: text
  // that fits its own cell NEVER spills. Without it, every centred cell with
  // an empty neighbour was painted into a multi-column overlay and centred in
  // THAT — which is why a timetable's "2-3" sat on its column boundary
  // instead of in its box.
  // A cell whose neighbours on the sides it may run to are taken does not spill,
  // however wide its text: asked first, as measuring the text was most of a
  // workbook's first render (300 ms of 0.47 MB), cell by cell, for nothing.
  if ((input.maxColumns ?? DEFAULT_MAX_COLUMNS) <= 0) return null;
  const mayRunRight = input.align !== "right" && input.col + 1 < input.columnCount && !input.isOccupied(input.row, input.col + 1);
  const mayRunLeft = input.align !== "left" && input.col > 0 && !input.isOccupied(input.row, input.col - 1);
  if (!mayRunRight && !mayRunLeft) return null;
  const textWidthPx = typeof input.textWidthPx === "function" ? input.textWidthPx() : input.textWidthPx;
  if (textWidthPx <= input.cellWidthPx) return null;
  const range = spillRange(input);
  if (range.startCol === input.col && range.endCol === input.col) return null;
  return range;
}

/**
 * A deterministic width estimate for environments without canvas text
 * measurement (jsdom). Fullwidth/CJK code points advance ~1em, ASCII ~0.55em,
 * spaces ~0.33em — coarse, but the decision only needs "clearly fits" vs
 * "clearly overflows", and the renderer uses REAL canvas metrics wherever
 * they exist.
 */
export function estimateTextWidthPx(text: string, fontSizePx: number): number {
  let width = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code === 0x20) width += 0.33;
    else if (
      (code >= 0x1100 && code <= 0x115f) ||
      (code >= 0x2e80 && code <= 0xa4cf) ||
      (code >= 0xac00 && code <= 0xd7a3) ||
      (code >= 0xf900 && code <= 0xfaff) ||
      (code >= 0xfe30 && code <= 0xfe4f) ||
      (code >= 0xff00 && code <= 0xff60) ||
      (code >= 0xffe0 && code <= 0xffe6)
    )
      width += 1;
    else width += 0.55;
  }
  return width * fontSizePx;
}
