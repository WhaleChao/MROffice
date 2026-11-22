/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Enter / Tab inside a selected range (Excel parity).
 *
 * Microsoft Support, "Keyboard shortcuts in Excel" (Other useful shortcut
 * keys / Move within a selected range): Enter moves from top to bottom within
 * the selection, Shift+Enter from bottom to top; Tab moves forward through a
 * row (down through a single-column selection), Shift+Tab back. The selection
 * itself stays, so a block is typed without reaching for the mouse:
 *
 *  - Enter walks the range column by column: down the column, then to the top
 *    of the next column, and from the last cell back to the first.
 *  - Tab walks it row by row: across the row, then to the start of the next
 *    row, wrapping the same way.
 *  - Shift reverses either walk.
 *  - Hidden rows and columns are passed over.
 *  - A merged cell is one stop: its top-left cell (inside the range) is the
 *    stop, and the cells it covers are passed over.
 *
 * The grid's selection is anchor..active, which always puts the active cell on
 * a corner. A walk leaves the active cell anywhere in the range, so the range
 * is remembered in a `SelectionEntryWalk` for as long as the active cell and
 * anchor are the ones the walk wrote; any other selection write ends it.
 */

import type { GridRect } from "./grid-nav";
import type { MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";

export interface EntryWalkCell {
  row: number;
  col: number;
}

export interface EntryWalkOptions {
  isRowHidden?: (row: number) => boolean;
  isColHidden?: (col: number) => boolean;
  merges?: readonly MergeSpan[];
}

/** Upper bound on cells inspected per step (a guard, not a rule). */
const MAX_WALK_STEPS = 2_000_000;

function normalizeRect(rect: GridRect): GridRect {
  return {
    r1: Math.min(rect.r1, rect.r2),
    r2: Math.max(rect.r1, rect.r2),
    c1: Math.min(rect.c1, rect.c2),
    c2: Math.max(rect.c1, rect.c2),
  };
}

function anyVisible(
  from: number,
  to: number,
  hidden: ((line: number) => boolean) | undefined,
): boolean {
  if (!hidden) return true;
  for (let line = from; line <= to; line += 1) {
    if (!hidden(line)) return true;
  }
  return false;
}

/**
 * The cell Enter (Tab) moves the active cell to inside `rect`, or null when
 * `from` is outside the range or the range has no visible cell. A range with
 * a single stop returns that stop.
 */
export function nextCellInSelection(
  rect: GridRect,
  from: EntryWalkCell,
  key: "Enter" | "Tab",
  reverse: boolean,
  options: EntryWalkOptions = {},
): EntryWalkCell | null {
  const { r1, r2, c1, c2 } = normalizeRect(rect);
  if (from.row < r1 || from.row > r2 || from.col < c1 || from.col > c2) {
    return null;
  }
  const { isRowHidden, isColHidden, merges = [] } = options;
  if (!anyVisible(r1, r2, isRowHidden) || !anyVisible(c1, c2, isColHidden)) {
    return null;
  }
  // Enter walks columns (major) cell by cell down their rows (minor); Tab
  // walks rows (major) across their columns (minor).
  const byColumn = key === "Enter";
  const majorLo = byColumn ? c1 : r1;
  const majorHi = byColumn ? c2 : r2;
  const minorLo = byColumn ? r1 : c1;
  const minorHi = byColumn ? r2 : c2;
  const majorHidden = byColumn ? isColHidden : isRowHidden;
  const minorHidden = byColumn ? isRowHidden : isColHidden;
  const step = reverse ? -1 : 1;
  const cellAt = (major: number, minor: number): EntryWalkCell =>
    byColumn ? { row: minor, col: major } : { row: major, col: minor };
  const isStop = (cell: EntryWalkCell) => {
    const span = merges.find(
      (candidate) =>
        cell.row >= candidate.row &&
        cell.row < candidate.row + candidate.rowSpan &&
        cell.col >= candidate.col &&
        cell.col < candidate.col + candidate.colSpan,
    );
    if (!span) return true;
    // The merge's first cell inside the range is its one stop; that is its
    // anchor whenever the anchor lies in the range.
    return (
      cell.row === Math.max(span.row, r1) && cell.col === Math.max(span.col, c1)
    );
  };

  let major = byColumn ? from.col : from.row;
  let minor = byColumn ? from.row : from.col;
  for (let guard = 0; guard < MAX_WALK_STEPS; guard += 1) {
    minor += step;
    if (minor < minorLo || minor > minorHi) {
      // Off the end of the line: the next line, wrapping past the last one.
      major += step;
      if (major < majorLo) major = majorHi;
      if (major > majorHi) major = majorLo;
      minor = reverse ? minorHi : minorLo;
    }
    if (majorHidden?.(major)) {
      // Pass a hidden line in one move: land beyond its last cell.
      minor = reverse ? minorLo : minorHi;
      continue;
    }
    const cell = cellAt(major, minor);
    if (cell.row === from.row && cell.col === from.col) return cell;
    if (minorHidden?.(minor)) continue;
    if (isStop(cell)) return cell;
  }
  return null;
}

/**
 * The same step where the screen shows the rows in another order than they
 * are stored in (a sorted Sheet View). `rect` and `from` name stored rows, as
 * the selection does; the walk goes through the range's rows in the order
 * they are displayed (`displayPosition`: stored row to shown row) and answers
 * with a stored row again. Merges are not walked as one stop here: a sorted
 * view's sort range holds none.
 */
export function nextCellInReorderedSelection(
  rect: GridRect,
  from: EntryWalkCell,
  key: "Enter" | "Tab",
  reverse: boolean,
  displayPosition: (row: number) => number,
  options: Omit<EntryWalkOptions, "merges"> = {},
): EntryWalkCell | null {
  const { r1, r2, c1, c2 } = normalizeRect(rect);
  if (from.row < r1 || from.row > r2 || from.col < c1 || from.col > c2) {
    return null;
  }
  const rows: number[] = [];
  for (let row = r1; row <= r2; row += 1) rows.push(row);
  rows.sort((a, b) => displayPosition(a) - displayPosition(b));
  const { isRowHidden, isColHidden } = options;
  const next = nextCellInSelection(
    { r1: 0, r2: rows.length - 1, c1, c2 },
    { row: rows.indexOf(from.row), col: from.col },
    key,
    reverse,
    {
      isRowHidden: isRowHidden ? (index) => isRowHidden(rows[index]) : undefined,
      isColHidden,
    },
  );
  return next ? { row: rows[next.row], col: next.col } : null;
}

/**
 * An anchor that keeps the grid's anchor..active rectangle inside `rect` for
 * a walked active cell: the far corner. It reproduces `rect` exactly whenever
 * the active cell is on a corner, and it never equals the active cell of a
 * multi-cell range, so a click on the active cell ends the walk.
 */
export function entryWalkAnchor(
  rect: GridRect,
  active: EntryWalkCell,
): EntryWalkCell {
  const { r1, r2, c1, c2 } = normalizeRect(rect);
  return {
    row: active.row - r1 >= r2 - active.row ? r1 : r2,
    col: active.col - c1 >= c2 - active.col ? c1 : c2,
  };
}

/** A range an Enter / Tab walk is moving through. */
export interface SelectionEntryWalk {
  sheetId: string;
  rect: GridRect;
  active: EntryWalkCell;
  anchor: EntryWalkCell;
}

/**
 * True while the selection is still the one the walk wrote: same sheet, same
 * active cell, same anchor. Any other selection write (a click, an arrow key,
 * a Shift extension, the Name Box) changes one of them and ends the walk.
 */
export function selectionEntryWalkHolds(
  walk: SelectionEntryWalk | null,
  sheetId: string,
  active: EntryWalkCell,
  anchor: EntryWalkCell | null,
): walk is SelectionEntryWalk {
  return (
    walk !== null &&
    walk.sheetId === sheetId &&
    walk.active.row === active.row &&
    walk.active.col === active.col &&
    anchor !== null &&
    walk.anchor.row === anchor.row &&
    walk.anchor.col === anchor.col
  );
}
