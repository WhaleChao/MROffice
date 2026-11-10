/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet freeze panes (Excel parity).
 *
 * Freeze panes pin the first N rows and/or first M columns so they stay visible
 * while the rest of the grid scrolls. The sheet model stores the freeze counts
 * as `freezeRows` / `freezeCols`; this module keeps all the freeze math out of
 * the React component so it can be unit-tested in isolation, and so the toolbar
 * (which resolves the menu actions to freeze counts) and the renderer (which
 * decides which rows/cols are pinned and where) agree on the model.
 *
 * Rendering model: the grid is a virtualized table. The frozen band(s) are the
 * leading rows/cols (0..freezeRows-1 and 0..freezeCols-1). Frozen cells are
 * pinned with `position: sticky` at the header offset plus the cumulative size
 * of the earlier frozen cells, so the pinned band always aligns with the
 * scrolling body (same fixed column width / row height). The corner where a
 * frozen row and a frozen column overlap stays pinned on both axes.
 */

/** The freeze configuration persisted on a sheet (both default to 0 = none). */
export interface FreezeConfig {
  freezeRows: number;
  freezeCols: number;
}

/**
 * The toolbar menu actions, mirroring Excel for the web's Freeze Panes
 * dropdown. That menu lists FIVE commands in three groups:
 *
 *   Freeze at selection            → freezePanes
 *   ---
 *   Top row                        → freezeTopRow
 *   Up to row N                    → freezeRowsToActive
 *   ---
 *   First column                   → freezeFirstColumn
 *   Up to column X                 → freezeColumnsToActive
 *
 * "Top row" / "First column" are FIXED one-line freezes; the "Up to …" entries
 * are separate commands that read the active cell. Collapsing each pair into a
 * single active-cell-driven command loses Excel's fixed one-line freeze, which
 * is the one people reach for most.
 */
export type FreezeAction =
  | "freezeTopRow"
  | "freezeFirstColumn"
  | "freezeRowsToActive"
  | "freezeColumnsToActive"
  | "freezePanes"
  | "unfreeze";

/**
 * Coerce a persisted/raw freeze count into a safe non-negative integer, clamped
 * to the sheet's row/col total (you cannot freeze more lines than exist, and a
 * freeze that covers the whole sheet would leave nothing to scroll). `undefined`
 * / non-finite / negative values normalize to 0 (unfrozen).
 */
export function normalizeFreezeCount(value: unknown, total: number): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return 0;
  const max = Math.max(0, Math.floor(total) - 1);
  return Math.min(Math.floor(n), max);
}

/**
 * Read a normalized freeze config off a raw sheet-like object, clamped to the
 * sheet's dimensions. Stored fields are optional, so missing values mean "no
 * freeze".
 */
export function readFreezeConfig(sheet: {
  freezeRows?: unknown;
  freezeCols?: unknown;
  rowCount: number;
  columnCount: number;
}): FreezeConfig {
  return {
    freezeRows: normalizeFreezeCount(sheet.freezeRows, sheet.rowCount),
    freezeCols: normalizeFreezeCount(sheet.freezeCols, sheet.columnCount),
  };
}

/**
 * Resolve a toolbar menu action to the next freeze config.
 *
 * Excel for the web's menu, in menu order:
 * - freezePanes ("Freeze at selection"): pin every row above and every column
 *   left of the active cell, i.e. freezeRows = activeRow, freezeCols =
 *   activeCol. When the active cell is A1 this freezes nothing, matching Excel.
 * - freezeTopRow ("Top row"): pin row 0 ONLY (rows=1), whatever the active cell
 *   is, and no columns. This is a fixed freeze, not an active-cell one.
 * - freezeRowsToActive ("Up to row N"): pin rows 1..N where N is the active
 *   cell's 1-based row, i.e. rows = activeRow + 1, and no columns. A7 pins
 *   rows 1-7; A1 pins one row, which is where it coincides with "Top row".
 * - freezeFirstColumn ("First column"): pin column 0 ONLY (cols=1), no rows.
 * - freezeColumnsToActive ("Up to column X"): pin columns A..X where X is the
 *   active cell's column, i.e. cols = activeCol + 1, and no rows. T7 pins A-T.
 * - unfreeze: clear both.
 *
 * Each command sets ONE axis except freezePanes, which sets both; that is
 * Excel's behaviour — freezing the top row clears any frozen columns.
 *
 * Counts are clamped to the sheet so a freeze never swallows the whole grid.
 */
export function freezeConfigForAction(
  action: FreezeAction,
  active: { row: number; col: number },
  dims: { rowCount: number; columnCount: number },
): FreezeConfig {
  switch (action) {
    case "freezeTopRow":
      return {
        freezeRows: normalizeFreezeCount(1, dims.rowCount),
        freezeCols: 0,
      };
    case "freezeFirstColumn":
      return {
        freezeRows: 0,
        freezeCols: normalizeFreezeCount(1, dims.columnCount),
      };
    case "freezeRowsToActive":
      return {
        freezeRows: normalizeFreezeCount(
          Math.max(0, Math.floor(active.row)) + 1,
          dims.rowCount,
        ),
        freezeCols: 0,
      };
    case "freezeColumnsToActive":
      return {
        freezeRows: 0,
        freezeCols: normalizeFreezeCount(
          Math.max(0, Math.floor(active.col)) + 1,
          dims.columnCount,
        ),
      };
    case "freezePanes":
      return {
        freezeRows: normalizeFreezeCount(active.row, dims.rowCount),
        freezeCols: normalizeFreezeCount(active.col, dims.columnCount),
      };
    case "unfreeze":
    default:
      return { freezeRows: 0, freezeCols: 0 };
  }
}

/** True when the sheet has any frozen rows or columns. */
export function hasFreeze(config: FreezeConfig): boolean {
  return config.freezeRows > 0 || config.freezeCols > 0;
}

/** A single line (row or column) index tagged with whether it is frozen. */
export interface FreezeLine {
  /** Grid index (0-based row or column). */
  index: number;
  /** True when this line sits inside the frozen band. */
  frozen: boolean;
}

/**
 * Pixel offset, measured from the inside edge of the header, at which a frozen
 * line is pinned. The first frozen line sits flush against the header; each
 * later frozen line is offset by the cumulative size of the frozen lines before
 * it. Returns the header offset (sticky anchor in the scroll container) when you
 * add `headerSizePx`.
 */
export function frozenLineOffset(
  frozenOrdinal: number,
  lineSizePx: number,
  headerSizePx: number,
): number {
  return headerSizePx + frozenOrdinal * lineSizePx;
}

// The band's total thickness is never a count times one line size: the grid
// measures it with `offsetOfLine`, which honours resized rows and columns
// and collapses hidden ones, so no fixed-size helper is offered here.

/**
 * Split a virtualized render window into the frozen lines (always rendered, in
 * order 0..frozenCount-1) and the scrolling lines from the window that fall
 * outside the frozen band. Frozen lines are de-duplicated out of the scrolling
 * set so a line never renders twice when the window overlaps the frozen band
 * (e.g. when scrolled to the very top).
 *
 * @param windowLines the virtualized window's line indices (already overscanned)
 * @param frozenCount how many leading lines are frozen
 * @param totalLines the sheet's total row/column count (bounds the frozen band)
 */
export function splitFrozenLines(
  windowLines: readonly number[],
  frozenCount: number,
  totalLines: number,
): { frozen: number[]; scrolling: number[] } {
  const safeFrozen = Math.max(
    0,
    Math.min(Math.floor(frozenCount), Math.max(0, totalLines)),
  );
  const frozen: number[] = [];
  for (let i = 0; i < safeFrozen; i += 1) frozen.push(i);
  const scrolling = windowLines.filter((index) => index >= safeFrozen);
  return { frozen, scrolling };
}

/** Resolved sticky placement for one frozen line in the rendered grid. */
export interface FrozenLinePlacement {
  index: number;
  /** Ordinal within the frozen band (0-based). */
  ordinal: number;
  /** Pixel offset from the scroll-container edge where the line is pinned. */
  offsetPx: number;
}

/**
 * Resolve the sticky placement of every frozen line: its grid index, its ordinal
 * inside the band, and the pixel offset (header + earlier frozen lines) at which
 * it pins. Returned in band order so the renderer can stack them.
 */
export function resolveFrozenPlacements(
  frozenCount: number,
  lineSizePx: number,
  headerSizePx: number,
  totalLines: number,
): FrozenLinePlacement[] {
  const safeFrozen = Math.max(
    0,
    Math.min(Math.floor(frozenCount), Math.max(0, totalLines)),
  );
  const placements: FrozenLinePlacement[] = [];
  for (let ordinal = 0; ordinal < safeFrozen; ordinal += 1) {
    placements.push({
      index: ordinal,
      ordinal,
      offsetPx: frozenLineOffset(ordinal, lineSizePx, headerSizePx),
    });
  }
  return placements;
}

/**
 * `next` unless it holds exactly the offsets `previous` holds, in which case
 * `previous`. The frozen offset maps are recomputed whenever the grid's line
 * count grows (a scroll toward the end of the session viewport adds rows), and
 * every grid row keys its paint on their identity: an equal map under a new
 * identity repainted every mounted cell twice in the middle of a scroll.
 */
export function keepEqualFrozenOffsets(
  previous: ReadonlyMap<number, number> | null,
  next: Map<number, number>,
): Map<number, number> {
  if (!previous || previous.size !== next.size) return next;
  for (const [index, offset] of next) {
    if (previous.get(index) !== offset) return next;
  }
  return previous as Map<number, number>;
}

/** Whether a given grid line index falls inside the frozen band. */
export function isLineFrozen(index: number, frozenCount: number): boolean {
  return index >= 0 && index < Math.max(0, Math.floor(frozenCount));
}

/* -------------------------------------------------------------------------- */
/* Split panes (Excel/WPS "View > Split").                                    */
/*                                                                            */
/* SPLIT is distinct from FREEZE. Freeze PINS a leading band so it never      */
/* scrolls. Split places a MOVABLE DIVIDER that carves the grid into          */
/* independently-scrolling regions — the two sides show the SAME cells but    */
/* can be scrolled separately. Like freeze, split is purely a VIEW concern:   */
/* it changes only scroll/layout, never cell data or the formula evaluator.   */
/*                                                                            */
/* The sheet model stores the split as `splitRow` / `splitCol` — the grid     */
/* index of the FIRST line in the bottom/right region (so a split "at B3"     */
/* draws the horizontal divider above row 2 and the vertical divider left of  */
/* column 1, i.e. the divider sits on the leading edge of the split line,     */
/* matching Excel where View>Split splits above-and-left of the active cell). */
/* 0 / absent on an axis means "no split on that axis".                       */
/*                                                                            */
/* SPLIT-vs-FREEZE INTERACTION: they are MUTUALLY EXCLUSIVE. A sheet has one  */
/* or the other, never both. Applying a split clears any freeze; applying a   */
/* freeze clears any split (see splitConfigForToggle / the renderer wiring).  */
/* This mirrors Excel, where toggling Split off-and-on does not stack with a  */
/* frozen band, and keeps the offset math unambiguous (a single divider per   */
/* axis). The component enforces this when it commits either edit.            */
/* -------------------------------------------------------------------------- */

/** The split configuration persisted on a sheet (both default to 0 = none). */
export interface SplitConfig {
  /** First row index of the bottom region (0 = no horizontal split). */
  splitRow: number;
  /** First column index of the right region (0 = no vertical split). */
  splitCol: number;
}

/** Native OOXML pane identifiers, also used for the desktop pane focus order. */
export const SPLIT_PANE_NAMES = [
  "topLeft",
  "topRight",
  "bottomLeft",
  "bottomRight",
] as const;

export type SplitPaneName = (typeof SPLIT_PANE_NAMES)[number];

/** One sibling `<selection pane=...>` entry from the active sheet view. */
export interface SplitPaneSelection {
  pane: SplitPaneName;
  activeCell: string;
  sqref: string;
  activeCellId: number;
}

/**
 * Strict native split-view authority. OOXML persists one trailing-pane
 * `topLeftCell` on `<pane>` and one selection per pane; extra browser scroll
 * offsets intentionally remain editor-session state instead of being invented
 * as native attributes.
 */
export interface SplitPaneViewState {
  activePane: SplitPaneName;
  /** `pane/@topLeftCell`: the top-left visible cell of the trailing
   * (bottom-right) pane. */
  topLeftCell: string;
  /** `sheetView/@topLeftCell`: the top-left visible cell of the top-left pane,
   * which fixes where every leading pane is scrolled. Absent means A1, the
   * value Excel leaves unwritten. */
  leadingTopLeftCell?: string;
  selections: SplitPaneSelection[];
}

function splitColumnLabel(column: number): string {
  let value = Math.max(0, Math.floor(column)) + 1;
  let label = "";
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function splitCellRef(row: number, column: number): string {
  return `${splitColumnLabel(column)}${Math.max(0, Math.floor(row)) + 1}`;
}

function normalizeSplitCellRef(
  value: unknown,
  dims: { rowCount: number; columnCount: number },
): string | null {
  if (typeof value !== "string") return null;
  const match = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(value.trim());
  if (!match) return null;
  let column = 0;
  for (const char of match[1].toUpperCase()) {
    column = column * 26 + char.charCodeAt(0) - 64;
  }
  const row = Number(match[2]);
  if (
    !Number.isSafeInteger(row) ||
    row < 1 ||
    row > Math.max(1, Math.floor(dims.rowCount)) ||
    column < 1 ||
    column > Math.max(1, Math.floor(dims.columnCount))
  ) {
    return null;
  }
  return `${match[1].toUpperCase()}${row}`;
}

function splitCellCoordinates(value: string): { row: number; col: number } {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(value);
  if (!match) return { row: 0, col: 0 };
  let col = 0;
  for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, col: col - 1 };
}

function normalizeSplitSqref(
  value: unknown,
  dims: { rowCount: number; columnCount: number },
  activeCell: string,
  activeCellId: number,
): string | null {
  if (typeof value !== "string") return null;
  const ranges = value.trim().split(/\s+/).filter(Boolean);
  if (ranges.length === 0 || ranges.length > 64) return null;
  const normalized: string[] = [];
  const active = splitCellCoordinates(activeCell);
  if (
    !Number.isSafeInteger(activeCellId) ||
    activeCellId < 0 ||
    activeCellId >= ranges.length
  ) {
    return null;
  }
  let indexedRangeContainsActive = false;
  for (const [index, range] of ranges.entries()) {
    const [startRaw, endRaw, extra] = range.split(":");
    if (extra !== undefined) return null;
    const start = normalizeSplitCellRef(startRaw, dims);
    const end = endRaw ? normalizeSplitCellRef(endRaw, dims) : null;
    if (!start || (endRaw && !end)) return null;
    const first = splitCellCoordinates(start);
    const last = splitCellCoordinates(end ?? start);
    if (index === activeCellId) {
      indexedRangeContainsActive =
        active.row >= Math.min(first.row, last.row) &&
        active.row <= Math.max(first.row, last.row) &&
        active.col >= Math.min(first.col, last.col) &&
        active.col <= Math.max(first.col, last.col);
    }
    normalized.push(end ? `${start}:${end}` : start);
  }
  return indexedRangeContainsActive ? normalized.join(" ") : null;
}

/** Pane identifiers that physically exist for a normalized split. */
export function splitPaneNames(config: SplitConfig): SplitPaneName[] {
  if (config.splitRow > 0 && config.splitCol > 0) {
    return [...SPLIT_PANE_NAMES];
  }
  if (config.splitRow > 0) return ["topLeft", "bottomLeft"];
  if (config.splitCol > 0) return ["topLeft", "topRight"];
  return ["topLeft"];
}

/**
 * Normalize imported/editor split authority and fail closed to deterministic
 * defaults when pane names, cell refs, duplicate selections, or bounds are bad.
 */
export function normalizeSplitPaneViewState(
  value: unknown,
  config: SplitConfig,
  dims: { rowCount: number; columnCount: number },
  fallbackCell: { row: number; col: number } = { row: 0, col: 0 },
): SplitPaneViewState {
  const panes = splitPaneNames(config);
  const trailingPane: SplitPaneName =
    config.splitRow > 0 && config.splitCol > 0
      ? "bottomRight"
      : config.splitRow > 0
        ? "bottomLeft"
        : config.splitCol > 0
          ? "topRight"
          : "topLeft";
  const fallbackActive = splitCellRef(
    Math.min(Math.max(0, fallbackCell.row), Math.max(0, dims.rowCount - 1)),
    Math.min(Math.max(0, fallbackCell.col), Math.max(0, dims.columnCount - 1)),
  );
  const fallbackTopLeft = splitCellRef(
    config.splitRow > 0 ? config.splitRow : 0,
    config.splitCol > 0 ? config.splitCol : 0,
  );
  const record =
    value && typeof value === "object"
      ? (value as Record<string, unknown>)
      : null;
  const requestedActive = record?.activePane;
  const activePane =
    typeof requestedActive === "string" &&
    panes.includes(requestedActive as SplitPaneName)
      ? (requestedActive as SplitPaneName)
      : trailingPane;
  const requestedTopLeft = normalizeSplitCellRef(record?.topLeftCell, dims);
  // In a true split, the trailing pane may scroll back across the divider's
  // original cell boundary. Native topLeftCell therefore needs only a valid,
  // in-bounds A1 ref; the boundary-derived value is creation fallback only.
  const topLeftCell = requestedTopLeft ?? fallbackTopLeft;
  // The leading panes' scroll; A1 is the unwritten default.
  const requestedLeading = normalizeSplitCellRef(record?.leadingTopLeftCell, dims);
  const leadingTopLeftCell =
    requestedLeading !== null && requestedLeading !== "A1"
      ? requestedLeading
      : null;
  const selectionsByPane = new Map<SplitPaneName, SplitPaneSelection>();
  const rawSelections = Array.isArray(record?.selections)
    ? record.selections
    : [];
  let invalidSelectionSet = false;
  for (const raw of rawSelections) {
    if (!raw || typeof raw !== "object") {
      invalidSelectionSet = true;
      break;
    }
    const selection = raw as Record<string, unknown>;
    const pane = selection.pane;
    if (
      typeof pane !== "string" ||
      !panes.includes(pane as SplitPaneName) ||
      selectionsByPane.has(pane as SplitPaneName)
    ) {
      invalidSelectionSet = true;
      break;
    }
    const activeCell = normalizeSplitCellRef(selection.activeCell, dims);
    const activeCellId =
      selection.activeCellId === undefined ? 0 : Number(selection.activeCellId);
    const sqref = activeCell
      ? normalizeSplitSqref(selection.sqref, dims, activeCell, activeCellId)
      : null;
    if (!activeCell || !sqref || !Number.isSafeInteger(activeCellId)) {
      invalidSelectionSet = true;
      break;
    }
    selectionsByPane.set(pane as SplitPaneName, {
      pane: pane as SplitPaneName,
      activeCell,
      sqref,
      activeCellId,
    });
  }
  if (invalidSelectionSet) selectionsByPane.clear();
  const selections = panes.map((pane) => {
    const existing = selectionsByPane.get(pane);
    if (existing) return existing;
    const paneCell =
      pane === activePane
        ? fallbackActive
        : splitCellRef(
            pane.startsWith("bottom") && config.splitRow > 0
              ? config.splitRow
              : 0,
            pane.endsWith("Right") && config.splitCol > 0
              ? config.splitCol
              : 0,
          );
    return { pane, activeCell: paneCell, sqref: paneCell, activeCellId: 0 };
  });
  return leadingTopLeftCell
    ? { activePane, topLeftCell, leadingTopLeftCell, selections }
    : { activePane, topLeftCell, selections };
}

/**
 * Coerce a persisted/raw split count into a safe split index, clamped to the
 * sheet's row/col total. A split at line 0 is a no-op (nothing above/left of it
 * to scroll independently) → normalizes to 0. A split must leave at least one
 * line on the trailing side, so the max is `total - 1`. `undefined` /
 * non-finite / negative values normalize to 0 (no split). This mirrors
 * `normalizeFreezeCount` so split and freeze clamp identically.
 */
export function normalizeSplitLine(value: unknown, total: number): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return 0;
  const max = Math.max(0, Math.floor(total) - 1);
  return Math.min(Math.floor(n), max);
}

/**
 * Read a normalized split config off a raw sheet-like object, clamped to the
 * sheet's dimensions. Stored fields are optional, so missing values mean "no
 * split".
 */
export function readSplitConfig(sheet: {
  splitRow?: unknown;
  splitCol?: unknown;
  rowCount: number;
  columnCount: number;
}): SplitConfig {
  return {
    splitRow: normalizeSplitLine(sheet.splitRow, sheet.rowCount),
    splitCol: normalizeSplitLine(sheet.splitCol, sheet.columnCount),
  };
}

/** True when the sheet has a split on either axis. */
export function hasSplit(config: SplitConfig): boolean {
  return config.splitRow > 0 || config.splitCol > 0;
}

/**
 * Resolve a "Split" toggle at the active cell into the next split config.
 *
 * Excel's View>Split toggles: if the sheet is already split, the toggle REMOVES
 * the split (restoring a single scroll region). Otherwise it splits above-and-
 * left of the active cell, i.e. splitRow = activeRow, splitCol = activeCol
 * (clamped to the sheet). Toggling at A1 with no existing split is a no-op
 * (nothing to split above/left of row/col 0), matching Excel.
 */
export function splitConfigForToggle(
  current: SplitConfig,
  active: { row: number; col: number },
  dims: { rowCount: number; columnCount: number },
): SplitConfig {
  if (hasSplit(current)) {
    // Toggle off: a second invocation clears the split entirely.
    return { splitRow: 0, splitCol: 0 };
  }
  return {
    splitRow: normalizeSplitLine(active.row, dims.rowCount),
    splitCol: normalizeSplitLine(active.col, dims.columnCount),
  };
}

/**
 * Pixel position of a split divider on one axis, measured from the inside edge
 * of the header (the same origin `frozenLineOffset` uses). The divider sits on
 * the LEADING edge of the split line, so its distance from the header edge is
 * the cumulative size of every line BEFORE the split line. `offsetOfLine`
 * supplies that cumulative size (honouring per-line widths/heights); add
 * `headerSizePx` to get the position inside the scroll container.
 *
 * Returns `headerSizePx` (the header edge, i.e. no visible divider) when the
 * split line is 0 / out of range — never NaN. The split line is clamped to
 * [0, total] before measuring so a corrupt value cannot read past the axis.
 *
 * @param splitLine first line of the trailing region (0 = no split)
 * @param totalLines the axis's total row/column count
 * @param headerSizePx the leading header thickness (row-header width / col-header height)
 * @param cumulativeOffset offsetOfLine-style resolver: size summed for lines [0, line)
 */
export function splitDividerOffsetPx(
  splitLine: number,
  totalLines: number,
  headerSizePx: number,
  cumulativeOffset: (line: number, total: number) => number,
): number {
  const safeHeader = Number.isFinite(headerSizePx) ? headerSizePx : 0;
  const safeTotal = Math.max(0, Math.floor(totalLines));
  const line = Math.min(Math.max(0, Math.floor(splitLine)), safeTotal);
  if (line <= 0) return safeHeader;
  const before = cumulativeOffset(line, safeTotal);
  const safeBefore = Number.isFinite(before) ? before : 0;
  return safeHeader + safeBefore;
}

/**
 * Convert a pointer position (measured from the header edge, i.e. the scroll-
 * container content origin minus the header) into the split line it should snap
 * to: the first line whose LEADING edge is at-or-after the pointer, so dragging
 * the divider lands it on a whole-cell boundary (Excel snaps the split bar to
 * cell edges). The result is clamped to [1, total - 1] so a dragged split always
 * keeps at least one line on each side; a drag to the very top/left clamps to 1
 * (the minimal split) rather than 0 (which would mean "remove the split" — that
 * is the toggle's job, not a drag's). Returns 0 only for a degenerate sheet
 * (total < 2) where no split can exist.
 *
 * @param pointerPx distance from the header's inside edge to the pointer
 * @param totalLines the axis's total row/column count
 * @param lineAtOffset lineIndexAtOffset-style resolver: line index at a pixel offset
 */
export function splitLineForPointer(
  pointerPx: number,
  totalLines: number,
  lineAtOffset: (offset: number, total: number) => number,
): number {
  const safeTotal = Math.max(0, Math.floor(totalLines));
  // A sheet with fewer than 2 lines cannot be split (no trailing line remains).
  if (safeTotal < 2) return 0;
  const px = Number.isFinite(pointerPx) ? Math.max(0, pointerPx) : 0;
  const raw = lineAtOffset(px, safeTotal);
  const line = Number.isFinite(raw) ? Math.floor(raw) : 1;
  // Clamp to [1, total - 1]: a dragged split never collapses to 0 (remove) and
  // never swallows the whole axis (the trailing region keeps ≥1 line).
  return Math.min(Math.max(1, line), safeTotal - 1);
}
