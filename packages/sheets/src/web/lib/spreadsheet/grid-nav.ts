/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure keyboard-navigation target math for the spreadsheet grid (Excel/WPS
 * table-stakes). Kept OUT of the React component so the jump semantics can be
 * unit-tested in isolation and so every shortcut shares one implementation.
 *
 * Cell PRESENCE (does a cell hold a value?) is threaded in as a predicate so
 * this module never imports the sheet model or the React state. "Present" means
 * the cell has a non-empty value; a styled-but-empty cell is treated as empty
 * for navigation (Excel jumps over blank-valued cells regardless of formatting).
 *
 * Coordinates are 0-based. `rowCount`/`colCount` are the grid extent (so the
 * last addressable cell is `rowCount-1`/`colCount-1`). Every target is clamped
 * to `[0, count-1]` so a caller can pass it straight to the move handler.
 *
 * Shortcuts modeled (Excel semantics):
 *  - Ctrl/Cmd+Arrow  → `dataRegionEdge`: jump to the edge of the current data
 *    region in the arrow's direction. From a filled cell next to a filled cell,
 *    jump to the LAST contiguous filled cell. From a filled cell next to a blank
 *    (or from a blank), jump to the NEXT filled cell (or the grid edge if none).
 *  - Home            → `rowStart`: column 0 of the current row.
 *  - Ctrl/Cmd+Home   → `{ row: 0, col: 0 }` (A1) — a plain constant, no helper.
 *  - Ctrl/Cmd+End    → `lastUsedCell`: the bottom-right corner of the used range
 *    (max filled row × max filled col), or A1 when the sheet is empty.
 *  - PageUp/PageDown → `pageRow`: move the active row by one viewport of rows.
 *  - Ctrl/Cmd+A      → `usedRangeRect`: the bounding rectangle of all filled
 *    cells (so a first press selects the data island; the component widens to
 *    the whole grid on a second press / when already covering the used range).
 */

export type CellPresent = (row: number, col: number) => boolean;

export type NavDirection = "up" | "down" | "left" | "right";

export interface GridRect {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

function clamp(value: number, max: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value < 0) return 0;
  if (value > max) return max;
  return Math.floor(value);
}

/** The (row, col) step for each arrow direction. */
function stepFor(dir: NavDirection): { dRow: number; dCol: number } {
  switch (dir) {
    case "up":
      return { dRow: -1, dCol: 0 };
    case "down":
      return { dRow: 1, dCol: 0 };
    case "left":
      return { dRow: 0, dCol: -1 };
    case "right":
      return { dRow: 0, dCol: 1 };
  }
}

/**
 * The logical direction an arrow key travels in. The arrow keys move the way
 * they point on screen ("Move right one cell: Right arrow key", Microsoft
 * Support, "Keyboard shortcuts in Excel"), and a sheet laid out right to left
 * (sheetView rightToLeft) puts column A on the right, so there Right steps to
 * the previous column and Left to the next one. Up and Down are unchanged.
 * Tab and Shift+Tab are not arrows: they go to the next / previous cell in
 * the row whatever the sheet's direction.
 */
export function sheetArrowDirection(
  dir: NavDirection,
  rightToLeft: boolean,
): NavDirection {
  if (!rightToLeft) return dir;
  if (dir === "left") return "right";
  if (dir === "right") return "left";
  return dir;
}

/** The column step of a horizontal arrow key on a sheet of this direction. */
export function sheetArrowColumnStep(
  key: "ArrowLeft" | "ArrowRight",
  rightToLeft: boolean,
): 1 | -1 {
  return sheetArrowDirection(key === "ArrowRight" ? "right" : "left", rightToLeft) === "right"
    ? 1
    : -1;
}

/**
 * Excel's Ctrl+Arrow "jump to data-region edge" target.
 *
 * Rules (matching Excel exactly):
 *  - If the next cell in the direction is FILLED, jump to the last cell in the
 *    contiguous filled run (stop at the last filled cell before a blank/edge).
 *  - If the next cell is BLANK (or the start cell is blank), skip blanks and
 *    land on the FIRST filled cell found; if none exists before the edge, land
 *    on the grid edge in that direction.
 *  - Hidden rows and columns (filtered, grouped or hidden by hand) are not
 *    there: the walk steps over them and never lands on one, as an arrow key
 *    does, so the cell it lands on is one the reader can see and type into
 *    (live audit S-134: in a filtered region Cmd+Down landed on the hidden
 *    last record, whose cell has no input, and the typing went to the old
 *    cell).
 *
 * Always returns an in-bounds coordinate. A 1-row/1-col grid returns the start.
 */
export function dataRegionEdge(
  present: CellPresent,
  row: number,
  col: number,
  dir: NavDirection,
  rowCount: number,
  colCount: number,
  rowHidden: (row: number) => boolean = () => false,
  colHidden: (col: number) => boolean = () => false,
): { row: number; col: number } {
  const maxRow = Math.max(0, rowCount - 1);
  const maxCol = Math.max(0, colCount - 1);
  const startRow = clamp(row, maxRow);
  const startCol = clamp(col, maxCol);
  const { dRow, dCol } = stepFor(dir);

  const inBounds = (r: number, c: number): boolean =>
    r >= 0 && r <= maxRow && c >= 0 && c <= maxCol;
  // The next VISIBLE cell in the direction, or null past the grid's edge.
  const next = (r: number, c: number): { row: number; col: number } | null => {
    let nr = r + dRow;
    let nc = c + dCol;
    while (inBounds(nr, nc) && (rowHidden(nr) || colHidden(nc))) {
      nr += dRow;
      nc += dCol;
    }
    return inBounds(nr, nc) ? { row: nr, col: nc } : null;
  };

  const first = next(startRow, startCol);
  // No room to move in this direction → stay put.
  if (!first) return { row: startRow, col: startCol };

  const startFilled = present(startRow, startCol);
  let at = { row: startRow, col: startCol };

  if (startFilled && present(first.row, first.col)) {
    // Inside/at the start of a run → walk to the last filled cell of the run.
    for (let step = next(at.row, at.col); step && present(step.row, step.col); step = next(at.row, at.col)) {
      at = step;
    }
    return at;
  }

  // Start is blank, or sits at the trailing edge of a run (next is blank) →
  // skip blanks and land on the first filled cell; else the grid edge (its
  // last visible cell).
  at = first;
  while (!present(at.row, at.col)) {
    const step = next(at.row, at.col);
    if (!step) return at;
    at = step;
  }
  return at;
}

/**
 * Ctrl+Home: "the beginning of a worksheet" (Microsoft Support, "Keyboard
 * shortcuts in Excel"). With frozen panes that is the top-left cell of the
 * SCROLLABLE pane, the first cell past the frozen rows and columns, not A1
 * inside the band; without them it is A1. Hidden lines are stepped over, so
 * the cell is one the reader can see. Rows are display positions.
 */
export function worksheetHomeCell(
  freezeRows: number,
  freezeCols: number,
  rowCount: number,
  colCount: number,
  rowHidden: (row: number) => boolean = () => false,
  colHidden: (col: number) => boolean = () => false,
): { row: number; col: number } {
  const firstVisible = (from: number, count: number, hidden: (line: number) => boolean) => {
    const last = Math.max(0, count - 1);
    const start = clamp(from, last);
    for (let line = start; line <= last; line += 1) if (!hidden(line)) return line;
    return start;
  };
  return { row: firstVisible(freezeRows, rowCount, rowHidden), col: firstVisible(freezeCols, colCount, colHidden) };
}










export function steppedVisibleLine(
  target: number,
  from: number,
  limit: number,
  hidden: (line: number) => boolean,
): number {
  const direction = Math.sign(target - from);
  if (direction === 0) return target;
  let line = target;
  while (line >= 0 && line < limit && hidden(line)) line += direction;
  return line >= 0 && line < limit ? line : from;
}








export function rowsShownLandingPast(row: number, shown: number, max: number, margin: number): number {
  return Math.max(shown, Math.min(max, row + 1 + Math.max(0, margin)));
}

/** Home: column 0 of the given row (clamped). */
export function rowStart(row: number, rowCount: number): { row: number; col: number } {
  return { row: clamp(row, Math.max(0, rowCount - 1)), col: 0 };
}

/**
 * The bounding rectangle of every filled cell (the "used range"). Returns null
 * when the sheet has no filled cells. Scans the whole grid extent, so callers
 * should keep `rowCount`/`colCount` to the addressable grid (not an absolute
 * cap) — which the component already does.
 */
export function usedRangeRect(
  present: CellPresent,
  rowCount: number,
  colCount: number,
): GridRect | null {
  let minR = Infinity;
  let minC = Infinity;
  let maxR = -Infinity;
  let maxC = -Infinity;
  for (let r = 0; r < rowCount; r += 1) {
    for (let c = 0; c < colCount; c += 1) {
      if (present(r, c)) {
        if (r < minR) minR = r;
        if (r > maxR) maxR = r;
        if (c < minC) minC = c;
        if (c > maxC) maxC = c;
      }
    }
  }
  if (maxR < 0) return null;
  return { r1: minR, r2: maxR, c1: minC, c2: maxC };
}

/** A stored cell that holds formatting (a style, a formula or a comment)
 * without a value: Excel's used range and Ctrl+End reach it (live audit B140). */
function storedCellIsFormatted(
  cell: { value?: unknown; style?: unknown; formula?: unknown; comment?: unknown } | null | undefined,
): boolean {
  if (!cell) return false;
  if (typeof cell.formula === "string" && cell.formula !== "") return true;
  if (cell.comment !== undefined && cell.comment !== null) return true;
  const style = cell.style;
  return style !== null && typeof style === "object" && Object.keys(style as object).length > 0;
}

/** Sparse worksheet equivalent of usedRangeRect. Keyboard navigation counts
 * only nonblank stored string values, not formatting or computed overlays;
 * `includeFormatted` (Ctrl+End) also counts cells that only carry a format.
 * A cell of spaces is nonblank text (ISBLANK is FALSE), so it counts too. */
export function storedUsedRangeRect(
  cells: Readonly<Record<string, { value?: unknown } | null | undefined>>,
  rowCount: number,
  colCount: number,
  rowPositions?: ReadonlyMap<number, number>,
  includeFormatted = false,
): GridRect | null {
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
  for (const key in cells) {
    const separator = key.indexOf(":");
    if (separator < 1) continue;
    const row = Number(key.slice(0, separator)), col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 ||
        col >= colCount || key !== `${row}:${col}`) continue;
    const visualRow = rowPositions ? rowPositions.get(row) : row;
    if (visualRow === undefined || visualRow < 0 || visualRow >= rowCount) continue;
    const value = cells[key]?.value;
    const hasValue = typeof value === "string" && value !== "";
    if (!hasValue && !(includeFormatted && storedCellIsFormatted(cells[key]))) continue;
    minR = Math.min(minR, visualRow); minC = Math.min(minC, col);
    maxR = Math.max(maxR, visualRow); maxC = Math.max(maxC, col);
  }
  return maxR < 0 ? null : { r1: minR, r2: maxR, c1: minC, c2: maxC };
}

/**
 * Ctrl/Cmd+End target: the bottom-right corner of the used range (max filled
 * row × max filled col — Excel reports these independently, so the corner can
 * be a blank cell). Returns A1 when the sheet is empty.
 */
export function lastUsedCell(
  present: CellPresent,
  rowCount: number,
  colCount: number,
): { row: number; col: number } {
  const rect = usedRangeRect(present, rowCount, colCount);
  if (!rect) return { row: 0, col: 0 };
  return { row: rect.r2, col: rect.c2 };
}

/** Sparse Ctrl+End, preserving the last VISUAL occurrence of a source row.
 * Sheet View supplies its already filtered order; ordinary sheets include all
 * stored rows just as the dense presence callback does. Empty views retain the
 * old first-visual-row fallback. The independent max column may be blank there. */
export function storedLastUsedCell(
  cells: Parameters<typeof storedUsedRangeRect>[0],
  rowCount: number,
  colCount: number,
  displayOrder: readonly number[] | null = null,
): { row: number; col: number } {
  const positions = displayOrder === null ? undefined : new Map<number, number>();
  if (displayOrder) for (let visual = 0; visual < displayOrder.length; visual++) {
    positions!.set(displayOrder[visual] ?? visual, visual);
  }
  const bounds = storedUsedRangeRect(cells, displayOrder?.length ?? rowCount, colCount, positions, true);
  const visualRow = bounds?.r2 ?? 0;
  return { row: displayOrder?.[visualRow] ?? visualRow, col: bounds?.c2 ?? 0 };
}

/**
 * PageUp/PageDown target row: shift the active row by `pageRows` (one viewport
 * of rows) up (`-1`) or down (`+1`), clamped to the grid. `pageRows` is the
 * number of fully-visible data rows; callers pass at least 1 so the key always
 * moves at least one row even in a tiny viewport.
 */
export function pageRow(
  row: number,
  sign: -1 | 1,
  pageRows: number,
  rowCount: number,
): number {
  const step = Math.max(1, Math.floor(pageRows));
  return clamp(row + sign * step, Math.max(0, rowCount - 1));
}
