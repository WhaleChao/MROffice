/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeFaceLineEm } from "./axis-spacing";

const PT = 4 / 3;

/** The face's line (px) at `fontPx`. */
export function officeLegendLinePx(fontPx: number, family: unknown): number {
  return officeFaceLineEm(family) * fontPx;
}

/**
 * The slot each entry of a side legend stands in when its entries are set
 * in more than one size (`format.legendEntryFonts`): as tall as the tallest
 * entry's text and 5.75 pt, the same for every entry (px; `tallestLinePx` is
 * that entry's text, all its lines). Windows PowerPoint, Calibri entries
 * (office-truth/probes/c30legentry, COM LegendEntry tops):
 *
 *   tallest entry    text (pt)   slot measured   slot here   error
 *   12 pt (p7)        14.65        20.50          20.40       -0.10
 *   18 pt (p5)        21.97        27.77          27.72       -0.05
 *   20 pt (p1)        24.41        30.20          30.16       -0.04
 *   24 pt (p4)        29.30        35.00          35.05       +0.05
 *   36 pt in two lines (p11)
 *                     87.89        93.00          93.64       +0.64
 *
 * The entry of two lines stands 0.64 pt closer than one line's gap would
 * have it (its own lines 2 x 1.2207 em); kept, as the corpus' entries are
 * one line.
 */
export const OFFICE_LEGEND_ENTRY_SLOT_GAP_PT = 5.75;

export function officeLegendEntrySlotPx(tallestLinePx: number): number {
  return tallestLinePx + OFFICE_LEGEND_ENTRY_SLOT_GAP_PT * PT;
}

/** On a legend whose keys `withOfficeLegendKeys` sized: "lined" or "filled". */
export const OFFICE_LEGEND_KEY_SIZED = "officeLegendKeySized";
/** On a legend whose names `withOfficeKeyToName` moved to Office's gap. */
export const OFFICE_LEGEND_NAME_SHIFTED = "officeLegendNameShifted";
/** The engine's space between a legend key and its name (px). */
export const OFFICE_LEGEND_KEY_TO_TEXT_PX = 5;

/** A filled key's side (px). */
export function officeLegendKeyPx(fontPx: number, family: unknown): number {
  return 0.45 * officeLegendLinePx(fontPx, family);
}

/** A key in a legend of lines: 19.2 pt long at any size. */
export const OFFICE_LEGEND_LINE_KEY_PX = 19.2 * PT;

/** From a key to its name (px): half a filled key less 0.37 pt; 2.03 pt after a line's. */
export function officeLegendKeyToNamePx(keyPx: number, lined: boolean): number {
  return lined ? 2.03 * PT : keyPx / 2 - 0.37 * PT;
}

/** The pitch (px) of rows of entries, or of entries one under another. */
export function officeLegendRowPitchPx(fontPx: number, family: unknown): number {
  return officeLegendLinePx(fontPx, family) + 5.8 * PT;
}

/** How far right of the chart's middle a row or block of entries stands (px). */
export function officeLegendCentreShiftPx(fontPx: number, family: unknown): number {
  return 0.0875 * officeLegendLinePx(fontPx, family) + PT;
}

export type OfficeLegendRow =
  /** One row, its names `gap` px short of the next key. */
  | { kind: "row"; gap: number }
  /** Columns `columnGap` px apart, each `cellPx` wide (key, gap, widest name). */
  | { kind: "columns"; columns: number; rows: number; cellPx: number; columnGap: number };

/**
 * How Office sets out a top or bottom legend's entries: `entries` each
 * entry's width (px: key, key-to-name gap and name), `chartPx` the chart's
 * width, `linePx` the face's line at the legend's size.
 */
export function officeLegendRow(entries: readonly number[], chartPx: number, linePx: number): OfficeLegendRow {
  const count = entries.length;
  const taken = entries.reduce((sum, width) => sum + width, 0);
  const natural = 0.28 * linePx + (0.2 * taken) / (count + 1);
  const room = 0.895 * chartPx + 6.5 * PT;
  if (taken + (count + 1) * natural <= room) return { kind: "row", gap: natural };
  const pressed = (room - taken) / (count + 1);
  if (pressed >= 0.29 * linePx || count < 2) return { kind: "row", gap: Math.max(0, pressed) };
  const cellPx = Math.max(...entries);
  const columnGap = 0.25 * linePx;
  let columns = 1;
  while (columns < count && (columns + 1) * cellPx + columns * columnGap <= room) columns += 1;
  const rows = Math.ceil(count / columns);
  return { kind: "columns", columns: Math.ceil(count / rows), rows, cellPx, columnGap };
}
