/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nearestKnownCellMap } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

interface ExtentSheet {
  rowCount: number;
  columnCount: number;
  cells: Readonly<Record<string, unknown>>;
  merges?: ReadonlyArray<{ row: number; col: number; rowSpan: number; colSpan: number }> | unknown;
  objects?: unknown;
  /** Charts made on the sheet: their cell anchors (`drawingAnchor`) count too. */
  charts?: unknown;
}

/** The anchor cells (0-based) of a drawing object placed against the grid. */
function objectAnchorCells(object: unknown): Array<{ row: number; col: number }> {
  const anchor = (object as { anchor?: unknown } | null)?.anchor as
    | { kind?: unknown; from?: unknown; to?: unknown }
    | undefined;
  if (!anchor || typeof anchor !== "object") return [];
  // A SmartArt member's graphic frame is anchored to cells too.
  const frame = (anchor as { smartArtFrame?: unknown }).smartArtFrame;
  const frameCells = frame && typeof frame === "object" ? objectAnchorCells({ anchor: frame }) : [];
  // An absolute anchor is measured from the sheet's origin, not a cell; an
  // absent kind is a one-cell anchor (Insert Shape writes just `from`).
  const kind = anchor.kind ?? "oneCell";
  if (kind !== "oneCell" && kind !== "twoCell") return frameCells;
  const points = kind === "twoCell" ? [anchor.from, anchor.to] : [anchor.from];
  const cells: Array<{ row: number; col: number }> = [...frameCells];
  for (const point of points) {
    const at = point as { row?: unknown; column?: unknown; col?: unknown } | undefined;
    const row = Number(at?.row);
    const col = Number(at?.column ?? at?.col);
    if (Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0) cells.push({ row, col });
  }
  return cells;
}

/** How far a map's cells reach, by the map. */
const cellsReachOfMap = new WeakMap<object, { rows: number; cols: number }>();

function reachOfKey(key: string, reach: { rows: number; cols: number }): void {
  const split = key.indexOf(":");
  if (split <= 0) return;
  const row = Number(key.slice(0, split));
  const col = Number(key.slice(split + 1));
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) return;
  if (row >= reach.rows) reach.rows = row + 1;
  if (col >= reach.cols) reach.cols = col + 1;
}

/**
 * Walking every key at each commit was about 10 ms per entry on 30,000 cells. A
 * cell cleared since the map it follows from keeps the reach it gave: the extent
 * this feeds only ever grows, and a sheet that shrinks is one whose lines moved,
 * whose cells are a new map with no such chain.
 */
function cellsReach(cells: Readonly<Record<string, unknown>>): { rows: number; cols: number } {
  const kept = cellsReachOfMap.get(cells);
  if (kept) return kept;
  const near = nearestKnownCellMap(cells as never, (map) => cellsReachOfMap.get(map as object));
  const reach = near ? { ...near.value } : { rows: 0, cols: 0 };
  if (near) for (const key of near.keys) { if (key in cells) reachOfKey(key, reach); }
  else for (const key in cells) reachOfKey(key, reach);
  cellsReachOfMap.set(cells, reach);
  return reach;
}

/**
 * The sheet with its stored extent grown to hold every cell, merge and
 * drawing-object anchor it carries, or the same object when it already does.
 *
 * `rowCount` / `columnCount` are the used range, and the page is only valid
 * when every stored cell lies inside it. Most commands write where data
 * already is; the ones that can reach a blank cell past the used range
 * (Merge & Center, a format, a comment, a paste) each had to remember to grow
 * it, and one that forgot produced a workbook the server refused, so nothing
 * after it was saved. Growing it where every sheet edit is committed makes
 * that a property of the model instead of a duty of each command. The server
 * holds a picture, shape or chart's anchor cells to the same extent, so a
 * picture pasted at H250 of a sheet used to row 218 is held the same way.
 */
export function sheetCoveringItsCells<S extends ExtentSheet>(sheet: S): S {
  const reach = cellsReach(sheet.cells);
  let rows = Math.max(sheet.rowCount, reach.rows);
  let cols = Math.max(sheet.columnCount, reach.cols);
  if (Array.isArray(sheet.merges)) {
    for (const merge of sheet.merges as ReadonlyArray<Partial<{ row: number; col: number; rowSpan: number; colSpan: number }>>) {
      const lastRow = Number(merge?.row) + Number(merge?.rowSpan);
      const lastCol = Number(merge?.col) + Number(merge?.colSpan);
      if (Number.isInteger(lastRow) && lastRow > rows) rows = lastRow;
      if (Number.isInteger(lastCol) && lastCol > cols) cols = lastCol;
    }
  }
  if (Array.isArray(sheet.objects)) {
    for (const object of sheet.objects) {
      for (const cell of objectAnchorCells(object)) {
        if (cell.row >= rows) rows = cell.row + 1;
        if (cell.col >= cols) cols = cell.col + 1;
      }
    }
  }
  if (Array.isArray(sheet.charts)) {
    for (const chart of sheet.charts) {
      for (const cell of objectAnchorCells({ anchor: (chart as { drawingAnchor?: unknown } | null)?.drawingAnchor })) {
        if (cell.row >= rows) rows = cell.row + 1;
        if (cell.col >= cols) cols = cell.col + 1;
      }
    }
  }
  rows = Math.min(rows, MAX_GRID_ROWS);
  cols = Math.min(cols, MAX_GRID_COLUMNS);
  return rows === sheet.rowCount && cols === sheet.columnCount ? sheet : { ...sheet, rowCount: rows, columnCount: cols };
}

/**
 * The sheet with its stored extent grown to hold every range (0-based,
 * inclusive), or the same object when it already does. For sheet metadata
 * addressed by range rather than by stored cell — a conditional-format rule's
 * Applies to — which the loader clamps to the stored extent, so a rule set
 * ahead on rows not yet filled in would otherwise come back cut short.
 */
export function sheetCoveringRanges<S extends { rowCount: number; columnCount: number }>(
  sheet: S,
  ranges: ReadonlyArray<{ r2: number; c2: number }>,
): S {
  let rows = sheet.rowCount;
  let cols = sheet.columnCount;
  for (const range of ranges) {
    if (Number.isInteger(range.r2) && range.r2 >= rows) rows = range.r2 + 1;
    if (Number.isInteger(range.c2) && range.c2 >= cols) cols = range.c2 + 1;
  }
  rows = Math.min(rows, MAX_GRID_ROWS);
  cols = Math.min(cols, MAX_GRID_COLUMNS);
  return rows === sheet.rowCount && cols === sheet.columnCount ? sheet : { ...sheet, rowCount: rows, columnCount: cols };
}
