/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Automate tab's script gallery, as Excel's sample Office Scripts run it
 * (learn.microsoft.com/office/dev/scripts/resources/samples):
 *
 * - Unhide All Rows and Columns: the used range's hidden rows and columns are
 *   shown (`getUsedRange().setRowHidden(false)` / `setColumnHidden(false)`).
 * - Freeze Selected Cells: `getFreezePanes().freezeAt(selection)` -- the
 *   selection is the frozen top-left pane, so the rows down to its last row
 *   and the columns up to its last column stay put. A selection of whole
 *   columns freezes columns only, one of whole rows rows only.
 * - Remove Hyperlinks from Sheet: every hyperlink in the sheet's used range is
 *   cleared, its value kept, its text no longer underlined and black.
 * - Count Empty Rows: the rows of the used range (by values) in which no cell
 *   has a value; a row with formatting only is empty, a cell of spaces is not.
 */
import { storedUsedRangeRect, type GridRect } from "@mrnote/core/web/lib/spreadsheet/grid-nav.ts";

type GalleryCell = {
  value?: unknown;
  hyperlink?: string;
  hyperlinkTooltip?: string;
  hyperlinkDisplay?: string;
  style?: object;
};

/** The used range by values (`getUsedRange(true)`), or null on an empty sheet. */
export function galleryValueUsedRange(
  cells: Readonly<Record<string, GalleryCell | null | undefined>>,
  rowCount: number,
  columnCount: number,
): GridRect | null {
  return storedUsedRangeRect(cells as Record<string, { value?: unknown }>, rowCount, columnCount);
}

/** The used range with formatted cells (`getUsedRange()`), or null. */
export function galleryUsedRange(
  cells: Readonly<Record<string, GalleryCell | null | undefined>>,
  rowCount: number,
  columnCount: number,
): GridRect | null {
  return storedUsedRangeRect(cells as Record<string, { value?: unknown }>, rowCount, columnCount, undefined, true);
}

/**
 * Count Empty Rows: `valueAt` gives a cell's value as `getValues()` would (a
 * formula's result); a row counts when every value's text is empty.
 */
export function galleryCountEmptyRows(
  range: GridRect,
  valueAt: (row: number, col: number) => unknown,
): number {
  let empty = 0;
  for (let row = range.r1; row <= range.r2; row += 1) {
    let blank = true;
    for (let col = range.c1; col <= range.c2 && blank; col += 1) {
      const value = valueAt(row, col);
      if (value !== null && value !== undefined && String(value).length > 0) blank = false;
    }
    if (blank) empty += 1;
  }
  return empty;
}

/**
 * Freeze Selected Cells: the cell Freeze Panes would be given so that the
 * selection is the frozen pane (rows and columns before it are frozen).
 * Null when the selection is the whole sheet (nothing would scroll).
 */
export function galleryFreezeAtCell(
  selection: GridRect,
  rowCount: number,
  columnCount: number,
): { row: number; col: number } | null {
  const wholeColumns = selection.r1 === 0 && selection.r2 >= rowCount - 1;
  const wholeRows = selection.c1 === 0 && selection.c2 >= columnCount - 1;
  const row = wholeColumns ? 0 : selection.r2 + 1;
  const col = wholeRows ? 0 : selection.c2 + 1;
  return row === 0 && col === 0 ? null : { row, col };
}

/**
 * Remove Hyperlinks from Sheet: the cells after the sample's
 * `clear(hyperlinks)`, `setUnderline(none)` and `setColor("Black")`, and the
 * keys of the cells it cleared.
 */
export function galleryCellsWithoutHyperlinks<C extends GalleryCell>(
  cells: Readonly<Record<string, C>>,
): { cells: Record<string, C>; cleared: string[] } {
  const cleared: string[] = [];
  let next: Record<string, C> | null = null;
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (!cell?.hyperlink) continue;
    next ??= { ...cells };
    const updated = { ...cell } as C;
    delete updated.hyperlink;
    delete updated.hyperlinkTooltip;
    delete updated.hyperlinkDisplay;
    const style: Record<string, unknown> = { ...(cell.style ?? {}) };
    delete style.underline;
    delete style.underlineStyle;
    style.color = "#000000";
    const themeColors = style.themeColors as Record<string, unknown> | undefined;
    if (themeColors && "color" in themeColors) {
      const rest = { ...themeColors };
      delete rest.color;
      if (Object.keys(rest).length) style.themeColors = rest;
      else delete style.themeColors;
    }
    (updated as GalleryCell).style = style;
    next[key] = updated;
    cleared.push(key);
  }
  return { cells: next ?? (cells as Record<string, C>), cleared };
}
