/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pane geometry for Excel's View > Split (G1-split-SPEC.md).
 *
 * A split window shows the same worksheet in two or four panes. Each pane is
 * an ordinary view of the sheet clipped to its rectangle, so this module is
 * only the arithmetic that view needs: which panes share a scroll axis, which
 * headings a pane carries, where a new split lands, which lines a pane shows,
 * and which cell a point in a pane hits. The renderer draws; nothing here
 * touches the DOM.
 *
 * Coordinates are worksheet content pixels measured from the leading edge of
 * row 1 / column A (no headings), the same origin as `offsetOfLine`.
 */
import { pxToPt } from "@mrnote/core/web/lib/office-units.ts";
import {
  lineIndexAtOffset,
  nextVisibleAxisLine,
  offsetOfLine,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { splitPaneNames, type SplitConfig, type SplitPaneName } from "@mrnote/core/web/lib/spreadsheet/freeze-panes.ts";

type LineSize = (line: number) => number;

export type SplitPaneRowGroup = "top" | "bottom";
export type SplitPaneColumnGroup = "left" | "right";

/** The vertical scroll a pane shares with its row partner. */
export function splitPaneRowGroup(
  pane: SplitPaneName,
  config: SplitConfig,
): SplitPaneRowGroup {
  return config.splitRow > 0 && pane.startsWith("bottom") ? "bottom" : "top";
}

/** The horizontal scroll a pane shares with its column partner. */
export function splitPaneColumnGroup(
  pane: SplitPaneName,
  config: SplitConfig,
): SplitPaneColumnGroup {
  return config.splitCol > 0 && pane.endsWith("Right") ? "right" : "left";
}

/**
 * One scroll per pane row and one per pane column (Excel's linked scroll
 * bars): top/bottom are vertical scroll offsets, left/right horizontal ones.
 * Values are content pixels (zoom-independent; a pane's scroll container
 * reads them as `scrollTop / zoom` and `|scrollLeft| / zoom`, since a pane
 * either pins its headings or does not draw them).
 */
export interface SplitPaneScrollState {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

export const EMPTY_SPLIT_PANE_SCROLL: SplitPaneScrollState = {
  top: 0,
  bottom: 0,
  left: 0,
  right: 0,
};

/** The scroll a pane shows under a linked-scroll state. */
export function splitPaneScrollFor(
  state: SplitPaneScrollState,
  pane: SplitPaneName,
  config: SplitConfig,
): { left: number; top: number } {
  return {
    top: state[splitPaneRowGroup(pane, config)],
    left: state[splitPaneColumnGroup(pane, config)],
  };
}

function finiteScroll(value: number | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.max(0, value)
    : fallback;
}

/**
 * Record one pane's scroll. The pane's row partner takes the vertical value
 * and its column partner the horizontal one, because they read the same
 * group. Returns the SAME object when nothing moved, so callers can skip work.
 */
export function withSplitPaneScroll(
  state: SplitPaneScrollState,
  pane: SplitPaneName,
  config: SplitConfig,
  scroll: { left?: number; top?: number },
): SplitPaneScrollState {
  const rowGroup = splitPaneRowGroup(pane, config);
  const columnGroup = splitPaneColumnGroup(pane, config);
  const top = finiteScroll(scroll.top, state[rowGroup]);
  const left = finiteScroll(scroll.left, state[columnGroup]);
  if (top === state[rowGroup] && left === state[columnGroup]) return state;
  return { ...state, [rowGroup]: top, [columnGroup]: left };
}

/**
 * The other panes that must follow `pane` and on which axis: the pane in the
 * same row follows its vertical scroll, the pane in the same column its
 * horizontal scroll. The diagonal pane shares nothing.
 */
export function splitPaneScrollPartners(
  pane: SplitPaneName,
  config: SplitConfig,
): Array<{ pane: SplitPaneName; axis: "top" | "left" }> {
  const rowGroup = splitPaneRowGroup(pane, config);
  const columnGroup = splitPaneColumnGroup(pane, config);
  const partners: Array<{ pane: SplitPaneName; axis: "top" | "left" }> = [];
  for (const candidate of splitPaneNames(config)) {
    if (candidate === pane) continue;
    if (splitPaneRowGroup(candidate, config) === rowGroup) {
      partners.push({ pane: candidate, axis: "top" });
    } else if (splitPaneColumnGroup(candidate, config) === columnGroup) {
      partners.push({ pane: candidate, axis: "left" });
    }
  }
  return partners;
}

/**
 * Headings belong to the window edge: column letters (and the column outline
 * bar) sit above the top panes only, row numbers (and the row outline bar)
 * left of the left panes only; the select-all corner and the level buttons
 * only at the top-left. View > Headings off removes all of them.
 */
export function splitPaneHeadings(
  pane: SplitPaneName,
  config: SplitConfig,
  headingsVisible: boolean,
): { columnHeadings: boolean; rowHeadings: boolean; corner: boolean } {
  const columnHeadings =
    headingsVisible && splitPaneRowGroup(pane, config) === "top";
  const rowHeadings =
    headingsVisible && splitPaneColumnGroup(pane, config) === "left";
  return { columnHeadings, rowHeadings, corner: columnHeadings && rowHeadings };
}

/**
 * Which scroll bars a pane carries. Excel draws them at the window edge: one
 * vertical bar per pane row on the trailing side, one horizontal bar per pane
 * column at the bottom. A pane without the bar still scrolls on that axis,
 * through its partner that has it.
 */
export function splitPaneScrollBars(
  pane: SplitPaneName,
  config: SplitConfig,
): { vertical: boolean; horizontal: boolean } {
  return {
    vertical: config.splitCol === 0 || splitPaneColumnGroup(pane, config) === "right",
    horizontal: config.splitRow === 0 || splitPaneRowGroup(pane, config) === "bottom",
  };
}

/**
 * The internal boundary (1..count-1) whose distance from the axis start is
 * nearest `distance`, or 0 when the distance is not inside the axis (no split
 * on this axis). Ties go to the earlier boundary so the bar never lands below
 * the row it was asked for.
 */
export function nearestSplitBoundary(
  distance: number,
  count: number,
  size: LineSize,
): number {
  const total = Math.max(0, Math.floor(count));
  if (total < 2 || !Number.isFinite(distance) || distance <= 0) return 0;
  const axisEnd = offsetOfLine(total, total, size);
  if (distance >= axisEnd) return 0;
  const line = lineIndexAtOffset(distance, total, size);
  const before = offsetOfLine(line, total, size);
  const after = offsetOfLine(line + 1, total, size);
  const candidate = distance - before <= after - distance ? line : line + 1;
  return Math.min(Math.max(1, candidate), total - 1);
}

/**
 * Excel's heading bands inside `xSplit` / `ySplit` (G1b-split-structure-SPEC
 * §4). Excel saves a split bar as its distance from the window's top / left
 * edge in twips at the current zoom, so the value holds the column-letter band
 * (in `ySplit`) or the row-number band (in `xSplit`) when headings are shown.
 * Measured on Excel for Mac 16 (Calibri 9/11/16/24 pt workbooks, 100%):
 *
 * - the letter band is the default row height, never under 22 pt;
 * - the number band fits the largest row number shown, never under three
 *   digits: about three digit advances plus 6.5 pt, never under 22 pt, and one
 *   digit advance wider for each further digit.
 *
 * Returns twips at 100% zoom.
 */
export const EXCEL_SPLIT_MIN_HEADING_PT = 22;

export function excelSplitHeadingTwips(
  axis: "row" | "col",
  metrics: {
    /** The sheet's default row height, CSS px (1 px = 0.75 pt). */
    readonly defaultRowHeightPx: number;
    /** The grid font's max digit advance, CSS px. */
    readonly digitWidthPx: number;
    /** The largest row number (1-based) the window shows. */
    readonly lastRowNumber: number;
  },
): number {
  const rowPx = Number.isFinite(metrics.defaultRowHeightPx) ? metrics.defaultRowHeightPx : 20;
  if (axis === "row") {
    return Math.round(Math.max(EXCEL_SPLIT_MIN_HEADING_PT, pxToPt(rowPx)) * 20);
  }
  const digitPx = Number.isFinite(metrics.digitWidthPx) && metrics.digitWidthPx > 0
    ? metrics.digitWidthPx
    : 7;
  const lastRow = Number.isFinite(metrics.lastRowNumber)
    ? Math.max(1, Math.floor(metrics.lastRowNumber))
    : 1;
  const digits = Math.max(3, String(lastRow).length);
  return Math.round(
    Math.max(EXCEL_SPLIT_MIN_HEADING_PT, digits * pxToPt(digitPx) + 6.5) * 20,
  );
}

/**
 * The largest row number a split window shows, from its panes' first rows
 * (0-based): Excel sizes the row-number band for it. A window shows about 50
 * rows below a pane's first row.
 */
export function excelSplitLastRowNumber(firstRows: readonly number[]): number {
  let first = 0;
  for (const row of firstRows) {
    if (Number.isFinite(row) && row > first) first = Math.floor(row);
  }
  return first + 1 + 50;
}

export interface SplitAtActiveCellInput {
  readonly active: { row: number; col: number };
  /** Current scroll of the grid, content pixels from row 1 / column A. */
  readonly scroll: { top: number; left: number };
  /** The data area on screen (viewport minus headings), content pixels. */
  readonly viewport: { width: number; height: number };
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowSize: LineSize;
  readonly colSize: LineSize;
}

export interface SplitAtActiveCellResult {
  readonly split: SplitConfig;
  /** Top-left visible cell of the trailing (bottom/right) pane. */
  readonly topLeftCell: { row: number; col: number };
  /** Linked scroll (content pixels) that keeps every cell where it was. */
  readonly scroll: SplitPaneScrollState;
}

/**
 * Excel's View > Split at the active cell (SPEC §1): bars above and left of
 * the active cell, at its on-screen edges; a cell in the first visible row or
 * column splits only the other axis; the top-left visible cell splits the
 * window into four equal panes. Nothing on screen moves: the leading panes
 * keep the current scroll and the trailing panes continue where the leading
 * ones stop.
 *
 * MRAI stores the bar as the boundary nearest the pane's size measured from
 * row 1 / column A (`splitRow`/`splitCol`), the projection the XLSX import uses
 * for `ySplit`/`xSplit`. Returns null when no bar can be placed.
 */
export function splitAtActiveCell(
  input: SplitAtActiveCellInput,
): SplitAtActiveCellResult | null {
  const rowCount = Math.max(0, Math.floor(input.rowCount));
  const columnCount = Math.max(0, Math.floor(input.columnCount));
  if (rowCount < 1 || columnCount < 1) return null;
  const scrollTop = finiteScroll(input.scroll.top, 0);
  const scrollLeft = finiteScroll(input.scroll.left, 0);
  const activeRow = Math.min(Math.max(0, Math.floor(input.active.row)), rowCount - 1);
  const activeCol = Math.min(Math.max(0, Math.floor(input.active.col)), columnCount - 1);
  const firstRow = lineIndexAtOffset(scrollTop, rowCount, input.rowSize);
  const firstCol = lineIndexAtOffset(scrollLeft, columnCount, input.colSize);
  const inFirstRow = activeRow <= firstRow;
  const inFirstCol = activeCol <= firstCol;
  const height = finiteScroll(input.viewport.height, 0);
  const width = finiteScroll(input.viewport.width, 0);
  const rowDistance = inFirstRow
    ? inFirstCol
      ? height / 2
      : 0
    : offsetOfLine(activeRow, rowCount, input.rowSize) - scrollTop;
  const colDistance = inFirstCol
    ? inFirstRow
      ? width / 2
      : 0
    : offsetOfLine(activeCol, columnCount, input.colSize) - scrollLeft;
  const splitRow = nearestSplitBoundary(rowDistance, rowCount, input.rowSize);
  const splitCol = nearestSplitBoundary(colDistance, columnCount, input.colSize);
  if (splitRow === 0 && splitCol === 0) return null;
  const topHeight = offsetOfLine(splitRow, rowCount, input.rowSize);
  const leftWidth = offsetOfLine(splitCol, columnCount, input.colSize);
  // The trailing pane starts at the active cell (Excel: "selecting row 13
  // splits the worksheet below row 12"); a centred split continues the
  // leading pane exactly.
  const bottomTop =
    splitRow === 0
      ? scrollTop
      : inFirstRow
        ? scrollTop + topHeight
        : offsetOfLine(activeRow, rowCount, input.rowSize);
  const rightLeft =
    splitCol === 0
      ? scrollLeft
      : inFirstCol
        ? scrollLeft + leftWidth
        : offsetOfLine(activeCol, columnCount, input.colSize);
  return {
    split: { splitRow, splitCol },
    topLeftCell: {
      row: lineIndexAtOffset(bottomTop, rowCount, input.rowSize),
      col: lineIndexAtOffset(rightLeft, columnCount, input.colSize),
    },
    scroll: {
      top: scrollTop,
      bottom: splitRow > 0 ? bottomTop : scrollTop,
      left: scrollLeft,
      right: splitCol > 0 ? rightLeft : scrollLeft,
    },
  };
}

/**
 * Linked scroll for a split restored from saved authority: the trailing panes
 * start at `topLeftCell` (`pane/@topLeftCell`), the leading panes at
 * `leadingTopLeftCell` (`sheetView/@topLeftCell`), row 1 / column A when it is
 * absent.
 */
export function initialSplitPaneScroll(
  config: SplitConfig,
  topLeftCell: { row: number; col: number } | null,
  rowOffset: (line: number) => number,
  colOffset: (line: number) => number,
  leadingTopLeftCell: { row: number; col: number } | null = null,
): SplitPaneScrollState {
  const row = topLeftCell?.row ?? config.splitRow;
  const col = topLeftCell?.col ?? config.splitCol;
  // With one bar the trailing pane shares the other axis with the leading
  // pane, so topLeftCell's other coordinate is the shared scroll.
  const bottom = Math.max(0, rowOffset(row));
  const right = Math.max(0, colOffset(col));
  const top = leadingTopLeftCell ? Math.max(0, rowOffset(leadingTopLeftCell.row)) : 0;
  const left = leadingTopLeftCell ? Math.max(0, colOffset(leadingTopLeftCell.col)) : 0;
  return {
    top: config.splitRow > 0 ? top : bottom,
    bottom,
    left: config.splitCol > 0 ? left : right,
    right,
  };
}

export interface SplitPaneWindowInput {
  /** Scroll of the pane in content pixels (headings excluded). */
  readonly top: number;
  readonly left: number;
  /** Visible data size in content pixels. */
  readonly height: number;
  readonly width: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowSize: LineSize;
  readonly colSize: LineSize;
  readonly overscan?: { rows: number; cols: number };
  /** Hard cap per axis so a pathological size cannot render unbounded. */
  readonly budget?: { rows: number; cols: number };
}

export interface SplitPaneWindow {
  /** Visible (non-hidden) visual rows, ascending. */
  readonly rows: readonly number[];
  /** Visible (non-hidden) columns, ascending. */
  readonly cols: readonly number[];
  /** Stable identity of the line ranges, for memoization. */
  readonly key: string;
}

function visibleLines(
  start: number,
  end: number,
  count: number,
  size: LineSize,
  budget: number,
): number[] {
  const lines: number[] = [];
  let line = nextVisibleAxisLine(Math.max(0, start), count, size, 1);
  while (line >= 0 && line < count && line < end && lines.length < budget) {
    if (size(line) > 0) lines.push(line);
    line = nextVisibleAxisLine(line + 1, count, size, 1);
  }
  return lines;
}

/**
 * The lines a pane paints: every visible line overlapping its viewport, plus
 * a small overscan, with hidden (zero-size) lines skipped in O(runs).
 */
export function splitPaneVisibleWindow(input: SplitPaneWindowInput): SplitPaneWindow {
  const rowCount = Math.max(0, Math.floor(input.rowCount));
  const columnCount = Math.max(0, Math.floor(input.columnCount));
  const overscan = input.overscan ?? { rows: 2, cols: 1 };
  const budget = input.budget ?? { rows: 400, cols: 160 };
  const top = finiteScroll(input.top, 0);
  const left = finiteScroll(input.left, 0);
  const height = finiteScroll(input.height, 0);
  const width = finiteScroll(input.width, 0);
  if (rowCount === 0 || columnCount === 0) return { rows: [], cols: [], key: "empty" };
  const rowStart = Math.max(0, lineIndexAtOffset(top, rowCount, input.rowSize) - overscan.rows);
  const rowEnd = Math.min(
    rowCount,
    lineIndexAtOffset(top + height, rowCount, input.rowSize) + overscan.rows + 1,
  );
  const colStart = Math.max(0, lineIndexAtOffset(left, columnCount, input.colSize) - overscan.cols);
  const colEnd = Math.min(
    columnCount,
    lineIndexAtOffset(left + width, columnCount, input.colSize) + overscan.cols + 1,
  );
  const rows = visibleLines(rowStart, rowEnd, rowCount, input.rowSize, budget.rows);
  const cols = visibleLines(colStart, colEnd, columnCount, input.colSize, budget.cols);
  return {
    rows,
    cols,
    key: `${rows[0] ?? -1}:${rows.at(-1) ?? -1}:${rows.length}|${cols[0] ?? -1}:${cols.at(-1) ?? -1}:${cols.length}`,
  };
}

export interface SplitPaneMerge {
  readonly row: number;
  readonly col: number;
  readonly rowSpan: number;
  readonly colSpan: number;
}

/** The merge covering a cell, if any. */
export function splitPaneMergeAt(
  merges: readonly SplitPaneMerge[],
  row: number,
  col: number,
): SplitPaneMerge | null {
  for (const merge of merges) {
    if (
      row >= merge.row &&
      row < merge.row + merge.rowSpan &&
      col >= merge.col &&
      col < merge.col + merge.colSpan
    ) {
      return merge;
    }
  }
  return null;
}

/**
 * The rectangle a merge occupies in content pixels. A merge across a bar is
 * drawn whole in every pane and clipped by the pane, which is how each pane
 * shows its part. Hidden lines inside the merge take no room. Null when the
 * merge is entirely hidden.
 */
export function splitPaneMergeRect(
  merge: SplitPaneMerge,
  rowCount: number,
  columnCount: number,
  rowSize: LineSize,
  colSize: LineSize,
  visualRow: (sourceRow: number) => number = (row) => row,
): { left: number; top: number; width: number; height: number } | null {
  const firstRow = visualRow(merge.row);
  const top = offsetOfLine(firstRow, rowCount, rowSize);
  const height = offsetOfLine(firstRow + merge.rowSpan, rowCount, rowSize) - top;
  const left = offsetOfLine(merge.col, columnCount, colSize);
  const width = offsetOfLine(merge.col + merge.colSpan, columnCount, colSize) - left;
  if (height <= 0 || width <= 0) return null;
  return { left, top, width, height };
}

/**
 * The cell under a point of a pane (content pixels, headings excluded): the
 * line whose extent holds the point, never a hidden line, and the anchor of
 * the merge the point falls in. Null outside the sheet.
 */
export function splitPaneCellAt(input: {
  readonly x: number;
  readonly y: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowSize: LineSize;
  readonly colSize: LineSize;
  readonly merges?: readonly SplitPaneMerge[];
  readonly sourceRow?: (visualRow: number) => number;
}): { row: number; col: number; visualRow: number } | null {
  const rowCount = Math.max(0, Math.floor(input.rowCount));
  const columnCount = Math.max(0, Math.floor(input.columnCount));
  if (
    rowCount === 0 ||
    columnCount === 0 ||
    !Number.isFinite(input.x) ||
    !Number.isFinite(input.y) ||
    input.x < 0 ||
    input.y < 0 ||
    input.y >= offsetOfLine(rowCount, rowCount, input.rowSize) ||
    input.x >= offsetOfLine(columnCount, columnCount, input.colSize)
  ) {
    return null;
  }
  const visualRow = lineIndexAtOffset(input.y, rowCount, input.rowSize);
  const col = lineIndexAtOffset(input.x, columnCount, input.colSize);
  const row = input.sourceRow ? input.sourceRow(visualRow) : visualRow;
  const merge = input.merges ? splitPaneMergeAt(input.merges, row, col) : null;
  return merge
    ? { row: merge.row, col: merge.col, visualRow }
    : { row, col, visualRow };
}

/**
 * F6 / Shift+F6 between split panes. Excel documents only that F6 "includes
 * the split panes"; MRAI walks them in the Panes collection order (top-left,
 * top-right, bottom-left, bottom-right — `Panes(3)` is the lower-left pane).
 */
export function nextSplitPane(
  config: SplitConfig,
  current: SplitPaneName,
  direction: 1 | -1,
): SplitPaneName {
  const panes = splitPaneNames(config);
  const index = panes.indexOf(current);
  if (index < 0) return panes[0] ?? "topLeft";
  return panes[(index + direction + panes.length) % panes.length];
}

/**
 * Where a dragged bar settles. Dragging it onto the heading edge (or past the
 * far edge of the window) removes that split (0), as dragging the old split
 * box back did; otherwise it snaps to the nearest cell boundary, keeping one
 * line on each side.
 *
 * @param pointer distance from the data edge (headings excluded), content px
 * @param extent the data area's size on this axis, content px
 */
export function splitLineForBarDrag(
  pointer: number,
  extent: number,
  count: number,
  size: LineSize,
): number {
  const total = Math.max(0, Math.floor(count));
  if (total < 2 || !Number.isFinite(pointer) || pointer <= 0) return 0;
  if (Number.isFinite(extent) && extent > 0 && pointer >= extent) return 0;
  const line = nearestSplitBoundary(pointer, total, size);
  return line > 0 ? line : Math.min(Math.max(1, lineIndexAtOffset(pointer, total, size)), total - 1);
}
