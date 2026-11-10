/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation table cell MERGE / SPLIT (PowerPoint / WPS "Merge Cells" /
 * "Split Cells" parity).
 *
 * The table schema already carries `colspan` / `rowspan` on each cell (set by
 * the .pptx importer). These pure helpers let the editor MERGE a rectangular
 * block of cells into one spanning cell, and SPLIT a previously-merged cell back
 * into 1×1 cells. The model matches HTML table semantics exactly: a cell with
 * `colspan` / `rowspan > 1` COVERS the grid positions to its right / below, and
 * those covered positions carry NO `<td>` of their own (they are absent from the
 * row arrays). So a merge REMOVES the covered cells from the rows; a split
 * RE-INSERTS empty cells for the positions the span used to cover.
 *
 * Keep this module free of React / DOM so it can be unit tested in isolation
 * (its only import is the pure cell rich-text helper, so a merge can carry the
 * cells' character runs). Every helper is total: a request that cannot be
 * satisfied (out-of-range anchor, a selection that straddles an existing span,
 * a 1×1 target) returns the rows UNCHANGED so the caller can no-op safely.
 */

import {
  cellCharacterRuns,
  cellDirectCharacterStyle,
  persistedCellRuns,
} from "./presentation-cell-rich-text";
import { pxToPtRounded } from "./office-units";
import type { TextRun } from "./presentation-runs";

/** The minimal cell shape these helpers read / write. Extra fields pass through. */
export interface MergeableCell {
  text: string;
  colspan?: number;
  rowspan?: number;
  [key: string]: unknown;
}

type Grid = MergeableCell[][];











export const MAX_TABLE_SPLIT_ROWS = 4_096;
export const MAX_TABLE_SPLIT_COLUMNS = 1_024;
export const MAX_PRESENTATION_TABLE_CELLS = 100_000;

/** Clamp a persisted span to a positive integer ≥ 1 (1 ⇒ no span). */
function spanOf(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 1;
  const n = Math.floor(value);
  return n >= 1 ? n : 1;
}

/**
 * One occupancy cell of the LOGICAL grid: which row/col array index the covering
 * cell lives at (its anchor), and whether THIS logical position is the anchor
 * itself (vs. a covered continuation position).
 */
interface Occupant {
  /** Row index (in the source `rows` arrays) of the covering cell's anchor. */
  anchorRow: number;
  /** Position of the covering cell WITHIN its row array. */
  cellIndex: number;
  /** True when this logical (row,col) is the anchor's own top-left position. */
  isAnchor: boolean;
}

/**
 * Build the LOGICAL occupancy grid: for every (row, col) logical position, which
 * source cell covers it. Spans are honoured — a 2×2 span at one anchor occupies
 * four logical positions, only one of which `isAnchor`. The grid width is the
 * max logical width across rows; ragged rows leave trailing positions `null`.
 *
 * Returns `null` when the table is structurally inconsistent in a way that makes
 * a safe merge/split impossible (a span overlaps another span). Callers treat
 * `null` as "leave the table unchanged".
 */
export function buildOccupancyGrid(rows: Grid): (Occupant | null)[][] | null {
  const rowCount = rows.length;
  if (rowCount === 0) return null;
  // Width upper bound: a row reaches no further than its own colspans plus the
  // columns that rowspans from the rows above still hold in it (cells are
  // placed left to right, skipping only those). The bound used to be the
  // colspans of ALL rows added up: a 185-row, 19-column table built a
  // 185 x 1000 grid each time it was asked (several times a keystroke,
  
  // `heldChange[r]`: how the columns held from above change at row r.
  const heldChange = new Array<number>(rowCount + 1).fill(0);
  let held = 0, widthGuess = 1;
  for (let r = 0; r < rowCount; r += 1) {
    held += heldChange[r];
    let reach = held;
    for (const c of rows[r]) {
      const cspan = spanOf(c.colspan);
      reach += cspan;
      const until = Math.min(rowCount, r + spanOf(c.rowspan));
      if (until > r + 1) { heldChange[r + 1] += cspan; heldChange[until] -= cspan; }
    }
    if (reach > widthGuess) widthGuess = reach;
  }
  // A generous, bounded grid; we trim trailing empty columns at the end.
  const maxCols = Math.min(widthGuess, 1000);
  const grid: (Occupant | null)[][] = Array.from({ length: rowCount }, () =>
    Array.from({ length: maxCols }, () => null),
  );
  for (let r = 0; r < rowCount; r += 1) {
    const row = rows[r];
    let logicalCol = 0;
    for (let cellIndex = 0; cellIndex < row.length; cellIndex += 1) {
      const cell = row[cellIndex];
      // Advance past any logical columns already occupied (by a rowspan from a
      // row above) before placing this cell.
      while (logicalCol < maxCols && grid[r][logicalCol] !== null) {
        logicalCol += 1;
      }
      const cspan = spanOf(cell.colspan);
      const rspan = spanOf(cell.rowspan);
      for (let dr = 0; dr < rspan; dr += 1) {
        for (let dc = 0; dc < cspan; dc += 1) {
          const rr = r + dr;
          const cc = logicalCol + dc;
          if (rr >= rowCount || cc >= maxCols) continue;
          if (grid[rr][cc] !== null) {
            // Two spans overlap → inconsistent table; bail.
            return null;
          }
          grid[rr][cc] = {
            anchorRow: r,
            cellIndex,
            isAnchor: dr === 0 && dc === 0,
          };
        }
      }
      logicalCol += cspan;
    }
  }
  return grid;
}

/** The logical width of the occupancy grid after trimming trailing empties. */
function logicalWidth(grid: (Occupant | null)[][]): number {
  let width = 0;
  for (const row of grid) {
    for (let c = row.length - 1; c >= 0; c -= 1) {
      if (row[c] !== null) {
        width = Math.max(width, c + 1);
        break;
      }
    }
  }
  return width;
}

export interface MergeSelection {
  /** Top logical row (inclusive). */
  row1: number;
  /** Left logical column (inclusive). */
  col1: number;
  /** Bottom logical row (inclusive). */
  row2: number;
  /** Right logical column (inclusive). */
  col2: number;
}

/**
 * Normalize a two-corner selection into an ordered, in-bounds rectangle, or
 * `null` when it collapses to a single logical cell (nothing to merge) or falls
 * outside the grid.
 */
function normalizeSelection(
  grid: (Occupant | null)[][],
  selection: MergeSelection,
): MergeSelection | null {
  const rowCount = grid.length;
  const width = logicalWidth(grid);
  const row1 = Math.min(selection.row1, selection.row2);
  const row2 = Math.max(selection.row1, selection.row2);
  const col1 = Math.min(selection.col1, selection.col2);
  const col2 = Math.max(selection.col1, selection.col2);
  if (row1 < 0 || col1 < 0 || row2 >= rowCount || col2 >= width) return null;
  if (row1 === row2 && col1 === col2) return null;
  return { row1, col1, row2, col2 };
}

/**
 * Can the logical rectangle be merged? It must be FULLY covered by anchors whose
 * spans lie ENTIRELY inside the rectangle (no span pokes out of the selection),
 * and every logical position inside must be occupied (no holes from a ragged
 * row). Returns the list of source cell coordinates `{row, cellIndex}` inside
 * the rectangle (anchors only) on success, else `null`.
 */
function collectMergeAnchors(
  grid: (Occupant | null)[][],
  rect: MergeSelection,
): { row: number; cellIndex: number }[] | null {
  const anchors: { row: number; cellIndex: number }[] = [];
  const seen = new Set<string>();
  for (let r = rect.row1; r <= rect.row2; r += 1) {
    for (let c = rect.col1; c <= rect.col2; c += 1) {
      const occ = grid[r][c];
      if (!occ) return null; // hole in selection (ragged row)
      // The covering anchor must itself sit inside the rectangle, else a span
      // straddles the selection boundary and the merge is illegal.
      const key = `${occ.anchorRow}:${occ.cellIndex}`;
      if (!seen.has(key)) {
        seen.add(key);
        anchors.push({ row: occ.anchorRow, cellIndex: occ.cellIndex });
      }
    }
  }
  // Verify no anchor's full span extends OUTSIDE the rectangle: re-walk the grid
  // for each collected anchor's covered positions.
  for (let r = 0; r < grid.length; r += 1) {
    for (let c = 0; c < grid[r].length; c += 1) {
      const occ = grid[r][c];
      if (!occ) continue;
      const inAnchors = anchors.some(
        (a) => a.row === occ.anchorRow && a.cellIndex === occ.cellIndex,
      );
      if (!inAnchors) continue;
      const inside =
        r >= rect.row1 && r <= rect.row2 && c >= rect.col1 && c <= rect.col2;
      if (!inside) return null; // an in-selection anchor reaches outside
    }
  }
  return anchors;
}

/**
 * MERGE the rectangular logical selection into a single spanning cell. The
 * surviving cell is the rectangle's TOP-LEFT anchor; its `colspan` / `rowspan`
 * grow to the rectangle's size, its text becomes the non-empty texts of the
 * merged cells joined by a space (PowerPoint concatenates content), and every
 * other covered source cell is REMOVED from its row array.
 *
 * Returns the rows UNCHANGED when the selection is degenerate (single cell), out
 * of range, straddles an existing span, or the table is inconsistent — so the
 * caller can always apply the result safely.
 */
export function mergeCells(rows: Grid, selection: MergeSelection): Grid {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return rows;
  const rect = normalizeSelection(grid, selection);
  if (!rect) return rows;
  const anchors = collectMergeAnchors(grid, rect);
  if (!anchors || anchors.length === 0) return rows;

  // The top-left source cell that will SURVIVE as the merged cell.
  const topLeftOcc = grid[rect.row1][rect.col1];
  if (!topLeftOcc) return rows;
  const survivorRow = topLeftOcc.anchorRow;
  const survivorIndex = topLeftOcc.cellIndex;

  // Collect the merged text (anchor order: by row then cell index) skipping the
  // survivor's own text position so it leads.
  const ordered = [...anchors].sort(
    (a, b) => a.row - b.row || a.cellIndex - b.cellIndex,
  );
  const texts: string[] = [];
  // The character runs follow the text: each contributing cell's runs are
  // concatenated in the same order, joined by the same plain-space separator,
  // so the merged cell's runs still concatenate exactly to its text (the save
  // contract) and no cell keeps runs that described a text it no longer has.
  const runs: TextRun[] = [];
  for (const a of ordered) {
    const cell = rows[a.row]?.[a.cellIndex];
    const t = cell?.text ?? "";
    if (typeof t !== "string" || t.length === 0) continue;
    if (texts.length > 0) runs.push({ text: " " });
    texts.push(t);
    const contributed = cellCharacterRuns(t, cell?.runs);
    // A covered cell's own direct character format (the highlight, colour,
    // size or flag written on the cell itself, as the importer stores a
    // uniformly formatted cell) travels with its words: the survivor keeps
    // only its own cell fields, so the words would otherwise lose it when
    // an empty cell leads the merge (PPT-B351). A run's own style still wins.
    const survivor = a.row === survivorRow && a.cellIndex === survivorIndex;
    const direct = !survivor && cell ? cellDirectCharacterStyle(cell) : undefined;
    runs.push(
      ...(direct
        ? contributed.map((run) => ({ text: run.text, style: { ...direct, ...run.style } }))
        : contributed),
    );
  }
  const mergedText = texts.join(" ");
  const mergedRuns = persistedCellRuns(mergedText, runs);

  const newColspan = rect.col2 - rect.col1 + 1;
  const newRowspan = rect.row2 - rect.row1 + 1;

  // Build a removal set of every source cell EXCEPT the survivor.
  const removal = new Set<string>();
  for (const a of anchors) {
    if (a.row === survivorRow && a.cellIndex === survivorIndex) continue;
    removal.add(`${a.row}:${a.cellIndex}`);
  }

  return rows.map((row, r) => {
    // Drop every covered (non-survivor) cell from this row.
    const kept = row.filter(
      (_, cellIndex) => !removal.has(`${r}:${cellIndex}`),
    );
    if (r !== survivorRow) return kept;
    // Re-derive the survivor's NEW index after removals earlier in its row, then
    // stamp the merged span + concatenated text onto it.
    let removedBefore = 0;
    for (let i = 0; i < survivorIndex; i += 1) {
      if (removal.has(`${r}:${i}`)) removedBefore += 1;
    }
    const newIndex = survivorIndex - removedBefore;
    return kept.map((cell, cellIndex) => {
      if (cellIndex !== newIndex) return cell;
      const next: MergeableCell = { ...cell, text: mergedText };
      if (mergedRuns) next.runs = mergedRuns;
      else delete next.runs;
      if (newColspan > 1) next.colspan = newColspan;
      else delete next.colspan;
      if (newRowspan > 1) next.rowspan = newRowspan;
      else delete next.rowspan;
      return next;
    });
  });
}

/**
 * SPLIT the spanning cell at logical position `(row, col)` back into 1×1 cells.
 * The anchor cell loses its `colspan` / `rowspan` (keeping its text), and an
 * EMPTY cell is inserted for every grid position the span used to cover —
 * restoring a full rectangular grid for those rows.
 *
 * Returns the rows UNCHANGED when the position is out of range, is not a merged
 * (spanning) anchor, or the table is inconsistent.
 */
/** A cell that looks like `source` but holds no text of its own. */
function splitAppearanceOf(source: MergeableCell): MergeableCell {
  const cell: MergeableCell = { ...source, text: "" };
  delete cell.rowspan;
  delete cell.colspan;
  delete cell.runs;
  delete cell.paragraphs;
  delete cell.href;
  delete cell.hyperlink;
  return cell;
}

export function splitCell(rows: Grid, row: number, col: number): Grid {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return rows;
  if (row < 0 || row >= grid.length) return rows;
  if (col < 0 || col >= (grid[row]?.length ?? 0)) return rows;
  const occ = grid[row][col];
  if (!occ) return rows;
  const anchorRow = occ.anchorRow;
  const anchorIndex = occ.cellIndex;
  const anchorCell = rows[anchorRow]?.[anchorIndex];
  if (!anchorCell) return rows;
  const cspan = spanOf(anchorCell.colspan);
  const rspan = spanOf(anchorCell.rowspan);
  if (cspan <= 1 && rspan <= 1) return rows; // nothing merged here

  // Logical column where this anchor starts (scan its anchor row in the grid).
  let startCol = -1;
  for (let c = 0; c < grid[anchorRow].length; c += 1) {
    const g = grid[anchorRow][c];
    if (
      g &&
      g.isAnchor &&
      g.anchorRow === anchorRow &&
      g.cellIndex === anchorIndex
    ) {
      startCol = c;
      break;
    }
  }
  if (startCol < 0) return rows;

  // Strip the span off the anchor (keep its text + other fields).
  const result: Grid = rows.map((r, rowIndex) =>
    rowIndex === anchorRow
      ? r.map((cell, cellIndex) => {
          if (cellIndex !== anchorIndex) return cell;
          const next: MergeableCell = { ...cell };
          delete next.colspan;
          delete next.rowspan;
          return next;
        })
      : r,
  );

  // Re-insert empty cells for every covered position, EXCEPT the anchor's own
  // top-left. We compute, per affected row, how many empties to add and at which
  // array index, by mapping logical columns → array insert positions.
  for (let dr = 0; dr < rspan; dr += 1) {
    const targetRow = anchorRow + dr;
    if (targetRow >= result.length) break;
    // Number of new empty cells this row gains from the split: the covered
    // logical columns in [startCol, startCol+cspan) for this row, minus the
    // anchor's own top-left position (only on dr===0, dc===0).
    const insertCols: number[] = [];
    for (let dc = 0; dc < cspan; dc += 1) {
      if (dr === 0 && dc === 0) continue; // the surviving anchor keeps this slot
      insertCols.push(startCol + dc);
    }
    if (insertCols.length === 0) continue;
    // Find the array index in `result[targetRow]` corresponding to logical
    // column `startCol` (for dr>0) or just after the anchor (for dr===0). We
    // re-walk the (already span-stripped) row's logical layout.
    const arrayIndex = arrayIndexForLogicalCol(
      result,
      grid,
      targetRow,
      anchorRow,
      anchorIndex,
      startCol,
      dr,
    );
    // The new cells wear the anchor's appearance (shading, borders, text
    // style), as PowerPoint's Split Cells leaves them; only its text and its
    // text model stay with the anchor (PPT-B405).
    const empties: MergeableCell[] = insertCols.map(() => splitAppearanceOf(anchorCell));
    result[targetRow] = [
      ...result[targetRow].slice(0, arrayIndex),
      ...empties,
      ...result[targetRow].slice(arrayIndex),
    ];
  }
  return result;
}

/**
 * SPLIT an ordinary 1×1 logical cell into `rowParts × columnParts` cells.
 *
 * DrawingML tables always have one deck-wide row/column grid, so splitting one
 * cell cannot create a private nested grid. PowerPoint/WPS insert the requested
 * grid lines across the whole table, then extend every non-target cell that
 * crosses those lines with `rowspan` / `colspan`. This preserves the visual
 * rectangle of every neighbouring cell while the target becomes independently
 * editable sub-cells.
 *
 * The target's top-left child keeps its content and formatting. Other children
 * inherit its visual fields but start empty and do not inherit rich text or
 * hyperlink payloads. Existing merges elsewhere expand across the inserted
 * lines. Invalid requests, merged targets, inconsistent tables, or edits beyond
 * native patchback limits are a safe no-op.
 */
export function splitCellInto(
  rows: Grid,
  row: number,
  col: number,
  rowParts: number,
  columnParts: number,
): Grid {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return rows;
  const width = logicalWidth(grid);
  const splitRows = Math.floor(rowParts);
  const splitColumns = Math.floor(columnParts);
  if (
    !Number.isFinite(rowParts) ||
    !Number.isFinite(columnParts) ||
    splitRows < 1 ||
    splitColumns < 1 ||
    (splitRows === 1 && splitColumns === 1) ||
    row < 0 ||
    row >= rows.length ||
    col < 0 ||
    col >= width ||
    rows.length + splitRows - 1 > MAX_TABLE_SPLIT_ROWS ||
    width + splitColumns - 1 > MAX_TABLE_SPLIT_COLUMNS
  ) {
    return rows;
  }
  const targetOccupant = grid[row]?.[col];
  if (!targetOccupant?.isAnchor) return rows;
  const target = rows[targetOccupant.anchorRow]?.[targetOccupant.cellIndex];
  if (!target) return rows;
  if (spanOf(target.rowspan) !== 1 || spanOf(target.colspan) !== 1) {
    return rows;
  }

  type Placement = {
    row: number;
    col: number;
    cell: MergeableCell;
  };
  const placements: Placement[] = [];
  const targetKey = `${targetOccupant.anchorRow}:${targetOccupant.cellIndex}`;
  const rowDelta = splitRows - 1;
  const columnDelta = splitColumns - 1;

  for (let logicalRow = 0; logicalRow < grid.length; logicalRow += 1) {
    for (let logicalCol = 0; logicalCol < width; logicalCol += 1) {
      const occupant = grid[logicalRow]?.[logicalCol];
      if (!occupant?.isAnchor) continue;
      const key = `${occupant.anchorRow}:${occupant.cellIndex}`;
      if (key === targetKey) continue;
      const source = rows[occupant.anchorRow]?.[occupant.cellIndex];
      if (!source) return rows;
      const sourceRows = spanOf(source.rowspan);
      const sourceColumns = spanOf(source.colspan);
      const crossesSplitRow =
        logicalRow <= row && row < logicalRow + sourceRows;
      const crossesSplitColumn =
        logicalCol <= col && col < logicalCol + sourceColumns;
      const nextRow = logicalRow > row ? logicalRow + rowDelta : logicalRow;
      const nextCol = logicalCol > col ? logicalCol + columnDelta : logicalCol;
      const nextRowspan = sourceRows + (crossesSplitRow ? rowDelta : 0);
      const nextColspan =
        sourceColumns + (crossesSplitColumn ? columnDelta : 0);
      const next: MergeableCell = { ...source };
      if (nextRowspan > 1) next.rowspan = nextRowspan;
      else delete next.rowspan;
      if (nextColspan > 1) next.colspan = nextColspan;
      else delete next.colspan;
      placements.push({ row: nextRow, col: nextCol, cell: next });
    }
  }

  const targetBase: MergeableCell = { ...target };
  delete targetBase.rowspan;
  delete targetBase.colspan;
  for (let rowOffset = 0; rowOffset < splitRows; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < splitColumns; columnOffset += 1) {
      const child: MergeableCell = {
        ...targetBase,
        text: rowOffset === 0 && columnOffset === 0 ? targetBase.text : "",
      };
      if (rowOffset !== 0 || columnOffset !== 0) {
        // New native cells inherit appearance, not another cell's text model or
        // relationship payload.
        delete child.runs;
        delete child.paragraphs;
        delete child.href;
        delete child.hyperlink;
      }
      placements.push({
        row: row + rowOffset,
        col: col + columnOffset,
        cell: child,
      });
    }
  }

  const nextRows: Grid = Array.from(
    { length: rows.length + rowDelta },
    () => [],
  );
  placements
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .forEach((placement) => {
      nextRows[placement.row].push(placement.cell);
    });
  const nextGrid = buildOccupancyGrid(nextRows);
  if (
    !nextGrid ||
    logicalWidth(nextGrid) !== width + columnDelta ||
    nextRows.some((nextRow) => nextRow.length === 0)
  ) {
    return rows;
  }
  return nextRows;
}

/**
 * Compute the array insert index within `result[targetRow]` for the empty cells
 * a split adds at logical `startCol`. For the anchor's own row (dr === 0) the
 * empties go immediately AFTER the anchor cell; for rows below, they go where
 * logical `startCol` maps in that row's (post-split) layout, derived from the
 * ORIGINAL occupancy grid so we honour any other spans before it.
 */
function arrayIndexForLogicalCol(
  result: Grid,
  grid: (Occupant | null)[][],
  targetRow: number,
  anchorRow: number,
  anchorIndex: number,
  startCol: number,
  dr: number,
): number {
  if (dr === 0) {
    // Empties follow the anchor cell directly.
    return anchorIndex + 1;
  }
  // For a lower row, count how many of that row's ORIGINAL cells have a logical
  // start strictly LESS than startCol; the empties insert right after them.
  const gridRow = grid[targetRow];
  let arrayIndex = 0;
  let countedAnchors = 0;
  const seenAnchors = new Set<string>();
  for (let c = 0; c < startCol && c < gridRow.length; c += 1) {
    const g = gridRow[c];
    if (!g || !g.isAnchor) continue;
    // Skip the split anchor itself if (impossibly) seen; count distinct anchors
    // whose own cell lives in THIS target row.
    if (g.anchorRow !== targetRow) continue;
    const key = `${g.anchorRow}:${g.cellIndex}`;
    if (seenAnchors.has(key)) continue;
    seenAnchors.add(key);
    countedAnchors += 1;
  }
  arrayIndex = countedAnchors;
  // Clamp to the row's current length.
  void result;
  void anchorRow;
  void anchorIndex;
  return Math.min(arrayIndex, result[targetRow].length);
}

/**
 * Whether the logical selection CAN be merged (≥ 2 logical cells, fully covered,
 * no straddling span). Drives the editor's "Merge" button disabled state.
 */
export function canMergeSelection(
  rows: Grid,
  selection: MergeSelection,
): boolean {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return false;
  const rect = normalizeSelection(grid, selection);
  if (!rect) return false;
  const anchors = collectMergeAnchors(grid, rect);
  return anchors !== null && anchors.length >= 1;
}

/**
 * Whether the cell at logical `(row, col)` is a MERGED (spanning) cell that can
 * be split. Drives the editor's "Split" button disabled state.
 */
export function canSplitCell(rows: Grid, row: number, col: number): boolean {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return false;
  if (row < 0 || row >= grid.length) return false;
  if (col < 0 || col >= (grid[row]?.length ?? 0)) return false;
  const occ = grid[row][col];
  if (!occ) return false;
  const cell = rows[occ.anchorRow]?.[occ.cellIndex];
  if (!cell) return false;
  return spanOf(cell.colspan) > 1 || spanOf(cell.rowspan) > 1;
}

/**
 * Map a source-array cell position `{row, cellIndex}` to its LOGICAL top-left
 * `(row, col)` in the occupancy grid (the inverse of what the editor needs when
 * a user clicks a `<td>` and we must address it logically). Returns `null` when
 * the table is inconsistent or the position is out of range.
 */
export function logicalPositionOfCell(
  rows: Grid,
  row: number,
  cellIndex: number,
): { row: number; col: number } | null {
  return logicalCellPositions(rows)[row]?.[cellIndex] ?? null;
}

/**
 * The rectangle two clicked cells select, each taken whole: a merged cell
 * reaches to the end of its span. Two side-by-side cells merged down two
 * rows select both rows (PPT-B063), where their top-left corners alone
 * named only the first row and Merge refused it as crossing a merged cell.
 */
export function mergeSelectionOfCells(
  rows: Grid,
  first: { row: number; cellIndex: number },
  second: { row: number; cellIndex: number },
): MergeSelection | null {
  const positions = logicalCellPositions(rows);
  const a = positions[first.row]?.[first.cellIndex];
  const b = positions[second.row]?.[second.cellIndex];
  if (!a || !b) return null;
  const aCell = rows[first.row]?.[first.cellIndex];
  const bCell = rows[second.row]?.[second.cellIndex];
  return {
    row1: Math.min(a.row, b.row),
    col1: Math.min(a.col, b.col),
    row2: Math.max(a.row + spanOf(aCell?.rowspan) - 1, b.row + spanOf(bCell?.rowspan) - 1),
    col2: Math.max(a.col + spanOf(aCell?.colspan) - 1, b.col + spanOf(bCell?.colspan) - 1),
  };
}









export function mergeSelectionOfScope(
  rows: Grid,
  scope: "cells" | "row" | "column" | "table",
  first: { row: number; cellIndex: number },
  second: { row: number; cellIndex: number } = first,
): MergeSelection | null {
  const corners = mergeSelectionOfCells(rows, first, second);
  if (!corners) return null;
  if (scope === "cells") return corners;
  const lastRow = rows.length - 1;
  const lastCol = tableLogicalWidth(rows) - 1;
  if (lastCol < 0) return null;
  const wholeRows = scope === "column" || scope === "table";
  const wholeCols = scope === "row" || scope === "table";
  return {
    row1: wholeRows ? 0 : corners.row1,
    col1: wholeCols ? 0 : corners.col1,
    row2: wholeRows ? lastRow : corners.row2,
    col2: wholeCols ? lastCol : corners.col2,
  };
}

type LogicalCellPositions = Array<Array<{ row: number; col: number } | null>>;

/** What the occupancy grid says of a table's shape: where each source cell
 * sits, and the logical width (null when the spans overlap). */
interface OccupancyFacts {
  positions: LogicalCellPositions;
  width: number | null;
}










const recentShapes: { shape: Int32Array; facts: OccupancyFacts }[] = [];
const RECENT_SHAPES = 4;

function shapeOf(rows: Grid): Int32Array {
  let size = 1 + rows.length;
  for (const row of rows) size += row.length * 2;
  const shape = new Int32Array(size);
  let at = 0;
  shape[at++] = rows.length;
  for (const row of rows) {
    shape[at++] = row.length;
    for (const cell of row) {
      shape[at++] = spanOf(cell.colspan);
      shape[at++] = spanOf(cell.rowspan);
    }
  }
  return shape;
}

function hasShape(rows: Grid, shape: Int32Array): boolean {
  if (shape[0] !== rows.length) return false;
  let at = 1;
  for (const row of rows) {
    if (shape[at++] !== row.length) return false;
    for (const cell of row) {
      if (shape[at++] !== spanOf(cell.colspan) || shape[at++] !== spanOf(cell.rowspan)) return false;
    }
  }
  return at === shape.length;
}

function occupancyFacts(rows: Grid): OccupancyFacts {
  const index = recentShapes.findIndex((entry) => hasShape(rows, entry.shape));
  if (index >= 0) {
    const [entry] = index > 0 ? recentShapes.splice(index, 1) : [recentShapes[0]];
    if (index > 0) recentShapes.unshift(entry);
    return entry.facts;
  }
  const grid = buildOccupancyGrid(rows);
  const positions: LogicalCellPositions = rows.map((row) => row.map(() => null));
  if (grid) {
    for (let r = 0; r < grid.length; r += 1) {
      for (let c = 0; c < grid[r].length; c += 1) {
        const occ = grid[r][c];
        if (occ?.isAnchor && positions[occ.anchorRow]?.[occ.cellIndex] === null) {
          positions[occ.anchorRow][occ.cellIndex] = { row: r, col: c };
        }
      }
    }
  }
  const facts: OccupancyFacts = { positions, width: grid ? logicalWidth(grid) : null };
  recentShapes.unshift({ shape: shapeOf(rows), facts });
  if (recentShapes.length > RECENT_SHAPES) recentShapes.length = RECENT_SHAPES;
  return facts;
}

/**
 * Resolve every source-array cell to its logical grid position in one pass.
 * Editors use this when a rectangular selection needs to test many cells; it
 * avoids rebuilding the occupancy grid once per cell on large tables. The
 * result is shared by every table of the same shape: read it, never change it.
 */
export function logicalCellPositions(
  rows: Grid,
): Array<Array<{ row: number; col: number } | null>> {
  return occupancyFacts(rows).positions;
}

/**
 * The LOGICAL column width of a table — the occupancy-grid width, which accounts
 * for both colspan AND the reach of rowspans from above (so it can exceed the
 * widest source row's array length). Falls back to the widest source-row length
 * when the table is structurally inconsistent, so callers (e.g. column banding /
 * first-last emphasis) always get a usable count. A table with no merges returns
 * exactly the array width, matching the un-merged fast path.
 */
export function tableLogicalWidth(rows: Grid): number {
  const { width } = occupancyFacts(rows);
  if (width === null) return Math.max(0, ...rows.map((row) => row.length));
  return width;
}

interface PositionedCell {
  row: number;
  col: number;
  cell: MergeableCell;
}

function positionedCells(rows: Grid): PositionedCell[] | null {
  const grid = buildOccupancyGrid(rows);
  if (!grid) return null;
  const result: PositionedCell[] = [];
  for (let row = 0; row < grid.length; row += 1) {
    for (let col = 0; col < grid[row].length; col += 1) {
      const occupant = grid[row][col];
      if (!occupant?.isAnchor) continue;
      const cell = rows[occupant.anchorRow]?.[occupant.cellIndex];
      if (!cell) return null;
      result.push({ row, col, cell });
    }
  }
  return result;
}

function rowsFromPositionedCells(
  rowCount: number,
  cells: PositionedCell[],
): Grid {
  const result: Grid = Array.from({ length: rowCount }, () => []);
  for (const item of [...cells].sort(
    (a, b) => a.row - b.row || a.col - b.col,
  )) {
    if (item.row < 0 || item.row >= rowCount) continue;
    result[item.row].push(item.cell);
  }
  return result;
}

function cellWithSpan(
  cell: MergeableCell,
  axis: "colspan" | "rowspan",
  span: number,
): MergeableCell {
  const next: MergeableCell = { ...cell };
  if (span > 1) next[axis] = span;
  else delete next[axis];
  return next;
}

/** What a cell of a new row or column copies from the cell beside it: its
 * formatting (fill, borders, margins, alignment, font), never its words, its
 * spans, its measured layout or its file receipts (`native*`, read against
 * the original cell's own XML). */
const CELL_KEYS_NOT_COPIED = new Set([
  "text",
  "runs",
  "paragraphs",
  "colspan",
  "rowspan",
  "textLayout",
  "id",
  "source",
]);

export function cellFormattingOf(cell: MergeableCell | undefined): MergeableCell {
  const next: MergeableCell = { text: "" };
  if (!cell) return next;
  for (const [key, value] of Object.entries(cell)) {
    if (CELL_KEYS_NOT_COPIED.has(key) || key.startsWith("native")) continue;
    next[key] = value;
  }
  // PowerPoint types a new row's or column's cell as its source cell's LAST
  // paragraph end mark (`a:endParaRPr`, without `dirty`), whatever its words
  
  // of 14 pt ending on a 16 pt mark gave 16 pt; words of 14 and 20 ending on
  // 20 gave 20; a mark stating no size gave a cell stating none). An imported
  // cell states that mark (`nativeEndMark`, reading row 478): the new cell's
  // character format is exactly it.
  // A format the person gave the whole cell since (its value no longer the
  // one read beside the mark, `nativeEndMark.cell`) is the end mark's too, as
  // PowerPoint applies a cell's format to its end marks.
  const endMark = cellEndMark(cell);
  if (endMark) {
    const read = endMarkCellRead(cell);
    for (const key of END_MARK_KEYS) {
      if (!sameValue(cell[key], read[key])) continue;
      delete next[key];
      if (key === "underline") delete next.underlineStyle;
      if (key === "strike") delete next.strikeStyle;
      if (endMark[key] !== undefined) next[key] = endMark[key];
    }
    // A mark stating no size draws at the size the import read for it, the
    // master's table text (`nativeEndMark.inheritedFontSize`, readings row
    
    
    const inherited = endMarkInheritedFontSize(cell);
    if (next.fontSize === undefined && inherited !== null && sameValue(cell.fontSize, read.fontSize)) {
      next.fontSize = inherited;
    }
    return next;
  }
  // A cell read before that mark was: where its size or face sits on its
  // characters rather than on the cell (a cell of several rows reads its `sz`
  // per run), the new cell is typed in the size and face every one of those
  // characters shares, so the save writes them on its end mark. Characters
  // that disagree name no one size: the new cell inherits, as the deck's
  
  // its new neighbours no size, drawn at the master's 13.5 pt).
  const shared = sharedRunCharacters(cell);
  if (next.fontSize === undefined && shared.fontSize !== undefined) next.fontSize = shared.fontSize;
  if (next.fontFamily === undefined && shared.fontFamily !== undefined) next.fontFamily = shared.fontFamily;
  return next;
}

/** The character formats a cell's end mark can state (`nativeEndMark`). */
const END_MARK_KEYS = ["fontSize", "fontFamily", "bold", "italic", "underline", "strike", "color"] as const;

/** A cell's last paragraph end mark as the import read it, or null. */
function cellEndMark(cell: MergeableCell): Partial<Record<(typeof END_MARK_KEYS)[number], unknown>> | null {
  const mark = (cell as { nativeEndMark?: unknown }).nativeEndMark;
  if (!mark || typeof mark !== "object" || Array.isArray(mark)) return null;
  const read: Partial<Record<(typeof END_MARK_KEYS)[number], unknown>> = {};
  const record = mark as Record<string, unknown>;
  if (typeof record.fontSize === "number" && Number.isFinite(record.fontSize) && record.fontSize > 0) {
    read.fontSize = record.fontSize;
  }
  if (typeof record.fontFamily === "string" && record.fontFamily.trim()) read.fontFamily = record.fontFamily;
  for (const flag of ["bold", "italic", "underline", "strike"] as const) {
    if (record[flag] === true) read[flag] = true;
  }
  if (typeof record.color === "string" && /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(record.color)) {
    read.color = record.color;
  }
  return read;
}

/** The size (points) a cell's end mark stating none draws at, as read, or null. */
function endMarkInheritedFontSize(cell: MergeableCell): number | null {
  const size = ((cell as { nativeEndMark?: { inheritedFontSize?: unknown } }).nativeEndMark ?? {}).inheritedFontSize;
  return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : null;
}

/** The cell's own character formats as read beside its end mark. */
function endMarkCellRead(cell: MergeableCell): Record<string, unknown> {
  const read = ((cell as { nativeEndMark?: { cell?: unknown } }).nativeEndMark ?? {}).cell;
  return read && typeof read === "object" && !Array.isArray(read) ? (read as Record<string, unknown>) : {};
}

function sameValue(a: unknown, b: unknown): boolean {
  return a === b || (typeof a === "string" && typeof b === "string" && a.toLowerCase() === b.toLowerCase());
}

/** The size (points) and face every worded run of a cell states alike. */
function sharedRunCharacters(cell: MergeableCell): { fontSize?: number; fontFamily?: string } {
  const runs = cellCharacterRuns(typeof cell.text === "string" ? cell.text : "", cell.runs).filter(
    (run) => run.text.replace(/\n/g, "") !== "",
  );
  if (runs.length === 0) return {};
  const sizes = new Set(runs.map((run) => run.style?.fontSize));
  const faces = new Set(runs.map((run) => run.style?.fontFamily));
  const [size] = sizes;
  const [face] = faces;
  return {
    ...(sizes.size === 1 && typeof size === "number" && Number.isFinite(size) && size > 0
      ? { fontSize: pxToPtRounded(size, 2) }
      : {}),
    ...(faces.size === 1 && typeof face === "string" && face ? { fontFamily: face } : {}),
  };
}

/** The cell anchoring logical position (row, col), or undefined. */
function cellAtPosition(rows: Grid, row: number, col: number): MergeableCell | undefined {
  const grid = buildOccupancyGrid(rows);
  const occupant = grid?.[row]?.[col];
  return occupant ? rows[occupant.anchorRow]?.[occupant.cellIndex] : undefined;
}

/**
 * Insert a logical row before `atRow` (0..rowCount). A vertical merge that
 * crosses the insertion boundary grows to cover the new row; uncovered
 * positions receive native empty cells.
 *
 * `formatFrom` (a logical row of `rows`): the new cells take the formatting
 * of that row's cell in their column, as PowerPoint's Insert Above / Below
 * gives the new row the formatting of the row it was inserted from. Without
 * it the new cells are bare, and a row added to a table whose cells state
 * their own fill and font drew in the table style's defaults, and was saved
 * into the deck as bare `a:tc` (AI70H H41 C12: 数据中台建设 above 合计).
 */
export function insertTableRow(rows: Grid, atRow: number, formatFrom?: number): Grid {
  if (rows.length === 0) return [[{ text: "" }]];
  if (!Number.isInteger(atRow) || atRow < 0 || atRow > rows.length) {
    return rows;
  }
  const positioned = positionedCells(rows);
  if (!positioned) return rows;
  const width = Math.max(tableLogicalWidth(rows), 1);
  const transformed: PositionedCell[] = positioned.map((item) => {
    const rowspan = spanOf(item.cell.rowspan);
    if (item.row < atRow && item.row + rowspan > atRow) {
      return {
        ...item,
        cell: cellWithSpan(item.cell, "rowspan", rowspan + 1),
      };
    }
    return item.row >= atRow ? { ...item, row: item.row + 1 } : item;
  });

  const partial = rowsFromPositionedCells(rows.length + 1, transformed);
  const partialGrid = buildOccupancyGrid(partial);
  if (!partialGrid) return rows;
  const template =
    formatFrom !== undefined && Number.isInteger(formatFrom) && formatFrom >= 0 && formatFrom < rows.length
      ? formatFrom
      : null;
  for (let col = 0; col < width; col += 1) {
    if (partialGrid[atRow]?.[col] === null) {
      transformed.push({
        row: atRow,
        col,
        cell: template === null ? { text: "" } : cellFormattingOf(cellAtPosition(rows, template, col)),
      });
    }
  }
  const result = rowsFromPositionedCells(rows.length + 1, transformed);
  return buildOccupancyGrid(result) ? result : rows;
}

/**
 * Insert a logical column before `atCol` (0..logicalWidth). Horizontal merges
 * crossing the boundary grow by one; every uncovered row gains an empty cell
 * anchored at the new logical column.
 *
 * `formatFrom` (a logical column of `rows`): each new cell takes the
 * formatting of that column's cell in its row, as Insert Left / Right does
 * (AI70H H41 C13: 首销价 beside 价格).
 */
export function insertTableColumn(rows: Grid, atCol: number, formatFrom?: number): Grid {
  const width = tableLogicalWidth(rows);
  if (rows.length === 0 || width === 0) return [[{ text: "" }]];
  if (!Number.isInteger(atCol) || atCol < 0 || atCol > width) return rows;
  const positioned = positionedCells(rows);
  if (!positioned) return rows;
  const transformed: PositionedCell[] = positioned.map((item) => {
    const colspan = spanOf(item.cell.colspan);
    if (item.col < atCol && item.col + colspan > atCol) {
      return {
        ...item,
        cell: cellWithSpan(item.cell, "colspan", colspan + 1),
      };
    }
    return item.col >= atCol ? { ...item, col: item.col + 1 } : item;
  });

  for (let row = 0; row < rows.length; row += 1) {
    const covered = transformed.some((item) => {
      const rowspan = spanOf(item.cell.rowspan);
      const colspan = spanOf(item.cell.colspan);
      return (
        item.row <= row &&
        row < item.row + rowspan &&
        item.col <= atCol &&
        atCol < item.col + colspan
      );
    });
    if (!covered) {
      transformed.push({
        row,
        col: atCol,
        cell:
          formatFrom !== undefined && Number.isInteger(formatFrom) && formatFrom >= 0 && formatFrom < width
            ? cellFormattingOf(cellAtPosition(rows, row, formatFrom))
            : { text: "" },
      });
    }
  }
  const result = rowsFromPositionedCells(rows.length, transformed);
  return buildOccupancyGrid(result) ? result : rows;
}

/**
 * Delete one logical row. A merge anchored above shrinks; a merge anchored in
 * the deleted row is re-anchored to the next surviving row so its editable
 * content and direct formatting are not discarded.
 */
export function deleteTableRow(rows: Grid, targetRow: number): Grid {
  if (
    rows.length <= 1 ||
    !Number.isInteger(targetRow) ||
    targetRow < 0 ||
    targetRow >= rows.length
  ) {
    return rows;
  }
  const positioned = positionedCells(rows);
  if (!positioned) return rows;
  const transformed: PositionedCell[] = [];
  for (const item of positioned) {
    const rowspan = spanOf(item.cell.rowspan);
    if (item.row < targetRow && item.row + rowspan > targetRow) {
      transformed.push({
        ...item,
        cell: cellWithSpan(item.cell, "rowspan", rowspan - 1),
      });
      continue;
    }
    if (item.row === targetRow) {
      if (rowspan > 1) {
        transformed.push({
          ...item,
          cell: cellWithSpan(item.cell, "rowspan", rowspan - 1),
        });
      }
      continue;
    }
    transformed.push(
      item.row > targetRow ? { ...item, row: item.row - 1 } : item,
    );
  }
  const result = rowsFromPositionedCells(rows.length - 1, transformed);
  return buildOccupancyGrid(result) ? result : rows;
}

/**
 * Delete one logical column. A merge covering it shrinks; an anchor that starts
 * in the deleted column survives when it spans farther right and otherwise is
 * removed. Cells to the right shift left by one logical column.
 */
export function deleteTableColumn(rows: Grid, targetCol: number): Grid {
  const width = tableLogicalWidth(rows);
  if (
    width <= 1 ||
    !Number.isInteger(targetCol) ||
    targetCol < 0 ||
    targetCol >= width
  ) {
    return rows;
  }
  const positioned = positionedCells(rows);
  if (!positioned) return rows;
  const transformed: PositionedCell[] = [];
  for (const item of positioned) {
    const colspan = spanOf(item.cell.colspan);
    if (item.col < targetCol && item.col + colspan > targetCol) {
      transformed.push({
        ...item,
        cell: cellWithSpan(item.cell, "colspan", colspan - 1),
      });
      continue;
    }
    if (item.col === targetCol) {
      if (colspan > 1) {
        transformed.push({
          ...item,
          cell: cellWithSpan(item.cell, "colspan", colspan - 1),
        });
      }
      continue;
    }
    transformed.push(
      item.col > targetCol ? { ...item, col: item.col - 1 } : item,
    );
  }
  const result = rowsFromPositionedCells(rows.length, transformed);
  return buildOccupancyGrid(result) ? result : rows;
}

/**
 * Append one full logical row. Colspans and rowspans make source row-array
 * lengths unreliable, so the new row is sized from the occupancy grid.
 */
export function appendTableRow(rows: Grid): Grid {
  return insertTableRow(rows, rows.length);
}

/**
 * Append one logical column at the table's right edge. Adding one source cell
 * to each row is correct even when a rowspan from above occupies part of that
 * row: the occupancy builder skips covered positions before placing the new
 * right-edge cell.
 */
export function appendTableColumn(rows: Grid): Grid {
  return insertTableColumn(rows, tableLogicalWidth(rows));
}

/**
 * Delete the bottom logical row without leaving a rowspan pointing beyond the
 * table. Anchors above the deleted row shrink by one; anchors that begin on the
 * deleted row disappear with it.
 */
export function deleteLastTableRow(rows: Grid): Grid {
  return deleteTableRow(rows, rows.length - 1);
}

/**
 * Delete the right-most LOGICAL column. Every unique anchor occupying that
 * column is handled once: a horizontally spanning anchor shrinks by one,
 * otherwise its source cell is removed. This preserves native colspan/rowspan
 * semantics across merged tables instead of blindly slicing row arrays.
 */
export function deleteLastTableColumn(rows: Grid): Grid {
  return deleteTableColumn(rows, tableLogicalWidth(rows) - 1);
}
