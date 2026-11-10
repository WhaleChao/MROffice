/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A split window through a structural edit (G1b-split-structure-SPEC.md).
 *
 * Excel's View > Split is window state: the bars sit at a physical distance
 * from the headings (`xSplit`/`ySplit`, twentieths of a point) and each pane
 * keeps its own first visible row and column. Inserting or deleting rows,
 * columns or cells changes the worksheet under the window, not the window:
 * the bars stay where they are and every pane keeps its top-left cell.
 *
 * MRAI stores a bar as the line boundary whose distance from row 1 / column A
 * is the bar's distance (`splitRow`/`splitCol`, the projection the XLSX import
 * and export use), so an edit that changes the sizes of the lines before the
 * bar re-projects that same distance onto the new lines. The sheet's stored
 * extent must hold every reference the split makes, so an edit that shrinks
 * it (a delete) leaves it large enough for the split view.
 */
import {
  normalizeSplitPaneViewState,
  readFreezeConfig,
  hasFreeze,
  readSplitConfig,
  hasSplit,
  type SplitConfig,
  type SplitPaneViewState,
} from "./freeze-panes";
import {
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
} from "./grid-sizing";

type Axis = "row" | "col";

/** The sheet fields a split depends on. */
export interface SplitStructureSheet {
  rowCount: number;
  columnCount: number;
  splitRow?: number;
  splitCol?: number;
  splitPaneState?: SplitPaneViewState;
  freezeRows?: number;
  freezeCols?: number;
  rowHeights?: Record<string, number>;
  columnWidths?: Record<string, number>;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  manuallyHiddenRows?: number[];
  manuallyHiddenCols?: number[];
}

function positive(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

/**
 * The stored size of each line on an axis, as the window draws an unfiltered
 * sheet: an explicit height/width, else the sheet default; a manually hidden
 * line takes no room.
 */
function storedLineSize(sheet: SplitStructureSheet, axis: Axis): (line: number) => number {
  const sizes = (axis === "row" ? sheet.rowHeights : sheet.columnWidths) ?? {};
  const fallback =
    axis === "row"
      ? positive(sheet.defaultRowHeight, DEFAULT_SPREADSHEET_ROW_HEIGHT_PX)
      : positive(sheet.defaultColumnWidth, DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX);
  const hidden = new Set(
    (axis === "row" ? sheet.manuallyHiddenRows : sheet.manuallyHiddenCols) ?? [],
  );
  return (line) => (hidden.has(line) ? 0 : positive(sizes[String(line)], fallback));
}

function distanceOfBoundary(boundary: number, size: (line: number) => number): number {
  let total = 0;
  for (let line = 0; line < boundary; line += 1) total += size(line);
  return total;
}

/**
 * The boundary (>= 1) whose distance from the axis start is nearest
 * `distance` under `size`; ties go to the earlier boundary, as the import's
 * projection does. Searches no further than `limit` lines.
 */
function boundaryNearest(
  distance: number,
  size: (line: number) => number,
  limit: number,
): number {
  let offset = 0;
  let best = 1;
  let bestGap = Number.POSITIVE_INFINITY;
  for (let boundary = 1; boundary <= limit; boundary += 1) {
    offset += size(boundary - 1);
    const gap = Math.abs(offset - distance);
    if (gap < bestGap) {
      best = boundary;
      bestGap = gap;
    }
    if (offset >= distance) break;
  }
  return best;
}

function parseRef(ref: string): { row: number; col: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(ref);
  if (!match) return null;
  let col = 0;
  for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, col: col - 1 };
}

/** The furthest row and column a split view references (0-based). */
function splitViewReach(
  split: SplitConfig,
  state: SplitPaneViewState | undefined,
): { row: number; col: number } {
  let row = split.splitRow;
  let col = split.splitCol;
  const refs: string[] = [];
  if (state) {
    refs.push(state.topLeftCell);
    if (state.leadingTopLeftCell) refs.push(state.leadingTopLeftCell);
  }
  for (const ref of refs) {
    const cell = parseRef(ref);
    if (!cell) continue;
    row = Math.max(row, cell.row);
    col = Math.max(col, cell.col);
  }
  return { row, col };
}

/** True when every cell a selection names lies inside the grid. */
function selectionFits(
  selection: SplitPaneViewState["selections"][number],
  dims: { rowCount: number; columnCount: number },
): boolean {
  const refs = [selection.activeCell, ...selection.sqref.split(/[\s:]+/)];
  return refs.every((ref) => {
    const cell = parseRef(ref);
    return Boolean(cell && cell.row < dims.rowCount && cell.col < dims.columnCount);
  });
}

/**
 * The split's pane authority normalized onto a grid. A remembered pane
 * selection that no longer fits (its cells were deleted off the end of a
 * shrunken sheet) returns to that pane's default; the others are kept.
 */
function fittedPaneState(
  state: SplitPaneViewState | undefined,
  split: SplitConfig,
  dims: { rowCount: number; columnCount: number },
): SplitPaneViewState {
  const kept = state
    ? { ...state, selections: state.selections.filter((selection) => selectionFits(selection, dims)) }
    : undefined;
  return normalizeSplitPaneViewState(kept, split, dims);
}

/**
 * The split of `after` (the sheet a structural edit produced from `before`)
 * as Excel leaves it: the bars at the distance they had in `before`, every
 * pane at its top-left cell, the active pane and the panes' remembered
 * selections unchanged. The extent grows to hold the split view when the edit
 * shrank it. Sheets without a split (or with a freeze) come back unchanged.
 */
export function splitAfterStructuralEdit<S extends SplitStructureSheet>(
  before: SplitStructureSheet,
  after: S,
  axis: Axis,
): S {
  const beforeFreeze = readFreezeConfig({
    freezeRows: before.freezeRows,
    freezeCols: before.freezeCols,
    rowCount: before.rowCount,
    columnCount: before.columnCount,
  });
  if (hasFreeze(beforeFreeze)) return after;
  const beforeSplit = readSplitConfig(before);
  if (!hasSplit(beforeSplit)) return after;

  const line = axis === "row" ? "splitRow" : "splitCol";
  const beforeLine = beforeSplit[line];
  const nextSplit: SplitConfig = { ...beforeSplit };
  if (beforeLine > 0) {
    const beforeSize = storedLineSize(before, axis);
    const afterSize = storedLineSize(after, axis);
    const distance = distanceOfBoundary(beforeLine, beforeSize);
    const limit = Math.max(
      1,
      (axis === "row" ? after.rowCount : after.columnCount) + beforeLine,
    );
    nextSplit[line] =
      distance === distanceOfBoundary(beforeLine, afterSize)
        ? beforeLine
        : boundaryNearest(distance, afterSize, limit);
  }

  const reach = splitViewReach(nextSplit, before.splitPaneState);
  const rowCount = Math.max(after.rowCount, reach.row + 1);
  const columnCount = Math.max(after.columnCount, reach.col + 1);
  const splitPaneState = fittedPaneState(before.splitPaneState, nextSplit, {
    rowCount,
    columnCount,
  });
  const next: S = { ...after, rowCount, columnCount, splitPaneState };
  if (nextSplit.splitRow > 0) next.splitRow = nextSplit.splitRow;
  else delete next.splitRow;
  if (nextSplit.splitCol > 0) next.splitCol = nextSplit.splitCol;
  else delete next.splitCol;
  return next;
}

/**
 * Carry the live split onto a sheet Undo or Redo restores. The split is
 * window state, which Excel's Undo does not touch, so the restored sheet
 * takes the split the window shows now. Its extent grows to hold that view
 * when the restored grid is smaller. Returns the restored sheet unchanged
 * when the live sheet has no split.
 */
export function splitCarriedOntoRestoredSheet<S extends SplitStructureSheet>(
  restored: S,
  live: SplitStructureSheet,
): S {
  const next = { ...restored };
  delete next.splitRow;
  delete next.splitCol;
  delete next.splitPaneState;
  const liveFreeze = readFreezeConfig({
    freezeRows: live.freezeRows,
    freezeCols: live.freezeCols,
    rowCount: live.rowCount,
    columnCount: live.columnCount,
  });
  const liveSplit = hasFreeze(liveFreeze) ? { splitRow: 0, splitCol: 0 } : readSplitConfig(live);
  if (!hasSplit(liveSplit)) {
    // Nothing was taken off: the restored record itself, so a sheet the
    // step did not touch keeps its identity (and every cache keyed on it).
    return Object.hasOwn(restored, "splitRow") ||
      Object.hasOwn(restored, "splitCol") ||
      Object.hasOwn(restored, "splitPaneState")
      ? next
      : restored;
  }
  const liveState = normalizeSplitPaneViewState(live.splitPaneState, liveSplit, {
    rowCount: live.rowCount,
    columnCount: live.columnCount,
  });
  const reach = splitViewReach(liveSplit, liveState);
  next.rowCount = Math.max(restored.rowCount, reach.row + 1);
  next.columnCount = Math.max(restored.columnCount, reach.col + 1);
  if (liveSplit.splitRow > 0) next.splitRow = liveSplit.splitRow;
  if (liveSplit.splitCol > 0) next.splitCol = liveSplit.splitCol;
  next.splitPaneState = fittedPaneState(liveState, liveSplit, {
    rowCount: next.rowCount,
    columnCount: next.columnCount,
  });
  // A restored freeze and the live split cannot coexist; the live window wins.
  delete next.freezeRows;
  delete next.freezeCols;
  return next;
}
