/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The live What-If Data Table a sheet keeps (Data > What-If Analysis > Data
 * Table): Excel's `{=TABLE(row_input, column_input)}` over the table's result
 * cells. The sheet persists only the recipe (`whatIfTables`); the result cells
 * keep their last values as ordinary numbers, so a reader that does not know
 * tables still reads them, and every calculation recomputes them by typing
 * each header value into the input cell (formula-evaluator, TABLE). The pure
 * one-shot computation stays in ./what-if-table.
 *
 * Semantics reproduced (Microsoft Support, "Calculate multiple results by
 * using a data table"; ECMA-376 CT_CellFormula t="dataTable"; MS-XLS Table):
 * - Column-oriented (column input only): the values run down the column left
 *   of the results, the formulas along the row above them; each result is the
 *   formula above it with its row's value typed into the column input cell.
 * - Row-oriented (row input only): the values run along the row above the
 *   results, the formulas down the column left of them.
 * - Two variables (both): the formula is the cell above-left of the results;
 *   the row above feeds the row input cell, the column left the column input.
 * - "Whenever a worksheet recalculates, any data tables will also
 *   recalculate", even when nothing they read changed.
 * - The input cells lie outside the table (MS-XLS Table: not inside the
 *   results nor the row and column of values).
 * - The results are one block: one cell cannot change ("Cannot change part of
 *   a data table."), the whole block can be cleared, which removes the table.
 * A definition that does not fit these rules is dropped and its cells keep
 * the values they show (fail closed).
 */

import { columnToLetters, parseCellRef } from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";

/** One table as the sheet stores it (`sheet.whatIfTables[]`). */
export interface WhatIfTableDefinition {
  /** The result cells, A1 without "$" ("D6:D9"): the selection the table was
   *  made from without its top row and left column. */
  readonly ref: string;
  /** Row input cell (A1): each value of the row above the results goes here. */
  readonly rowInput?: string;
  /** Column input cell (A1): each value of the column left of the results. */
  readonly colInput?: string;
}

type Cell = { readonly row: number; readonly col: number };
type Rect = { readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };

/** One table read from its definition: the result block and its inputs. */
export interface WhatIfTableBlock extends Rect {
  readonly rowInput: Cell | null;
  readonly colInput: Cell | null;
}

/** Most tables one sheet keeps. */
export const MAX_WHAT_IF_TABLES = 64;
/** Most result cells one table keeps (each is a recalculation of its formula;
 *  the Data Table command's own bound). */
export const MAX_WHAT_IF_TABLE_CELLS = 20_000;

const a1 = (cell: Cell): string => `${columnToLetters(cell.col)}${cell.row + 1}`;

function parseA1Cell(value: unknown): Cell | null {
  if (typeof value !== "string" || value.length === 0 || value.length > 16) return null;
  const cell = parseCellRef(value.replace(/\$/g, ""));
  return cell && cell.row < MAX_GRID_ROWS && cell.col < MAX_GRID_COLUMNS ? cell : null;
}

function parseA1Area(value: unknown): Rect | null {
  if (typeof value !== "string" || value.length === 0 || value.length > 40) return null;
  const parts = value.split(":");
  if (parts.length > 2) return null;
  const from = parseA1Cell(parts[0]);
  const to = parts.length === 2 ? parseA1Cell(parts[1]) : from;
  if (!from || !to) return null;
  return {
    r1: Math.min(from.row, to.row),
    c1: Math.min(from.col, to.col),
    r2: Math.max(from.row, to.row),
    c2: Math.max(from.col, to.col),
  };
}

const intersects = (a: Rect, b: Rect): boolean =>
  a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
const covers = (outer: Rect, inner: Rect): boolean =>
  outer.r1 <= inner.r1 && outer.r2 >= inner.r2 && outer.c1 <= inner.c1 && outer.c2 >= inner.c2;
const contains = (area: Rect, cell: Cell): boolean =>
  cell.row >= area.r1 && cell.row <= area.r2 && cell.col >= area.c1 && cell.col <= area.c2;

/** The result block in A1 ("D6:D9", or "D6" for one cell). */
export function whatIfTableRef(area: Rect): string {
  const from = a1({ row: area.r1, col: area.c1 });
  return area.r1 === area.r2 && area.c1 === area.c2 ? from : `${from}:${a1({ row: area.r2, col: area.c2 })}`;
}

/** The whole table: its results plus the row of values and formulas above
 *  and the column left of them (the selection it was made from). */
export function whatIfTableFootprint(block: Rect): Rect {
  return { r1: block.r1 - 1, c1: block.c1 - 1, r2: block.r2, c2: block.c2 };
}

/**
 * The sheet's tables, each checked: a result block with a row above and a
 * column left of it, inside the sheet and within the cell bound, at least one
 * input, every input outside the table, and no block overlapping an earlier
 * one. Whatever fails is left out (its cells keep their values).
 */
export function whatIfTableBlocks(
  raw: unknown,
  bounds?: { readonly rowCount: number; readonly columnCount: number },
): WhatIfTableBlock[] {
  if (!Array.isArray(raw) || raw.length === 0) return [];
  const out: WhatIfTableBlock[] = [];
  for (const entry of raw.slice(0, MAX_WHAT_IF_TABLES)) {
    if (!entry || typeof entry !== "object") continue;
    const record = entry as Record<string, unknown>;
    const area = parseA1Area(record.ref);
    if (!area || area.r1 < 1 || area.c1 < 1) continue;
    if ((area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) > MAX_WHAT_IF_TABLE_CELLS) continue;
    if (bounds && (area.r2 >= bounds.rowCount || area.c2 >= bounds.columnCount)) continue;
    const rowInput = record.rowInput === undefined ? null : parseA1Cell(record.rowInput);
    const colInput = record.colInput === undefined ? null : parseA1Cell(record.colInput);
    if ((record.rowInput !== undefined && !rowInput) || (record.colInput !== undefined && !colInput)) continue;
    if (!rowInput && !colInput) continue;
    const footprint = whatIfTableFootprint(area);
    if ((rowInput && contains(footprint, rowInput)) || (colInput && contains(footprint, colInput))) continue;
    if (out.some((block) => intersects(block, area))) continue;
    out.push({ ...area, rowInput, colInput });
  }
  return out;
}

/** A block as the sheet stores it. */
export function whatIfTableDefinition(block: WhatIfTableBlock): WhatIfTableDefinition {
  return {
    ref: whatIfTableRef(block),
    ...(block.rowInput ? { rowInput: a1(block.rowInput) } : {}),
    ...(block.colInput ? { colInput: a1(block.colInput) } : {}),
  };
}

/** The checked definitions of a sheet (see {@link whatIfTableBlocks}). */
export function normalizeWhatIfTables(
  raw: unknown,
  bounds?: { readonly rowCount: number; readonly columnCount: number },
): WhatIfTableDefinition[] {
  return whatIfTableBlocks(raw, bounds).map(whatIfTableDefinition);
}

/** The table whose results hold (row, col), or null. */
export function whatIfTableBlockAt(
  blocks: readonly WhatIfTableBlock[],
  row: number,
  col: number,
): WhatIfTableBlock | null {
  for (const block of blocks) if (contains(block, { row, col })) return block;
  return null;
}

/** Excel's text for the results: `=TABLE(B1,)`, `=TABLE(,B1)`,
 *  `=TABLE(B1,B2)` (row input first; a missing one is left blank). The
 *  formula bar shows it in braces, as an array formula. */
export function whatIfTableFormula(block: WhatIfTableBlock): string {
  return `=TABLE(${block.rowInput ? a1(block.rowInput) : ""},${block.colInput ? a1(block.colInput) : ""})`;
}

/** Whether a raw cell reads as a data table's result (`=TABLE(...)`). */
export function isWhatIfTableFormula(raw: string | null | undefined): boolean {
  return typeof raw === "string" && /^\s*=\s*TABLE\s*\(/i.test(raw);
}

/** What one result cell is: `formula` recalculated with each substitution's
 *  value cell typed into its input cell. A later substitution wins where two
 *  name one cell (the row input, as the dialog reads its boxes). */
export interface WhatIfTableProbe {
  readonly formula: Cell;
  readonly substitutions: ReadonlyArray<{ readonly input: Cell; readonly value: Cell }>;
}

export function whatIfTableProbe(block: WhatIfTableBlock, row: number, col: number): WhatIfTableProbe {
  const top = block.r1 - 1;
  const left = block.c1 - 1;
  if (block.rowInput && block.colInput) {
    return {
      formula: { row: top, col: left },
      substitutions: [
        { input: block.colInput, value: { row, col: left } },
        { input: block.rowInput, value: { row: top, col } },
      ],
    };
  }
  if (block.rowInput) {
    return { formula: { row, col: left }, substitutions: [{ input: block.rowInput, value: { row: top, col } }] };
  }
  return { formula: { row: top, col }, substitutions: [{ input: block.colInput!, value: { row, col: left } }] };
}

/** Every "row:col" key of a block's results. */
export function whatIfTableResultKeys(block: Rect): string[] {
  const keys: string[] = [];
  for (let row = block.r1; row <= block.r2; row += 1) {
    for (let col = block.c1; col <= block.c2; col += 1) keys.push(`${row}:${col}`);
  }
  return keys;
}

/** "Cannot change part of a data table.": a change over `rect` meets a
 *  table's results without covering all of them. */
export function whatIfTablePartRefused(raw: unknown, rect: Rect): boolean {
  return whatIfTableBlocks(raw).some((block) => intersects(block, rect) && !covers(rect, block));
}

/** The definitions left once the cells of `rects` are written or cleared: a
 *  table whose results they cover goes (Excel clears a whole table); null
 *  when none goes. Partial overlaps are the caller's to refuse first. */
export function whatIfTablesAfterOverwrite(
  raw: unknown,
  rects: readonly Rect[],
): WhatIfTableDefinition[] | null {
  if (raw === undefined || raw === null) return null;
  const blocks = whatIfTableBlocks(raw);
  const kept = blocks.filter((block) => !rects.some((rect) => covers(rect, block)));
  // An unchanged list stays as it is; a malformed entry goes with the change.
  return kept.length === blocks.length && Array.isArray(raw) && raw.length === blocks.length
    ? null
    : kept.map(whatIfTableDefinition);
}

/**
 * Whether inserting `count` lines at `at` (op "insert") or deleting the lines
 * `[at, at + count - 1]` (op "delete") would change part of a table: a line
 * inserted inside it (below its row of values, up to its last result) splits
 * it, and deleted lines that cut it without removing it whole shrink it.
 */
export function whatIfTableLineRefused(
  raw: unknown,
  axis: "row" | "col",
  at: number,
  count: number,
  op: "insert" | "delete",
): boolean {
  const n = Math.trunc(count);
  if (n < 1 || !Number.isInteger(at) || at < 0) return false;
  return whatIfTableBlocks(raw).some((block) => {
    const footprint = whatIfTableFootprint(block);
    const start = axis === "row" ? footprint.r1 : footprint.c1;
    const end = axis === "row" ? footprint.r2 : footprint.c2;
    if (op === "insert") return at > start && at <= end;
    const last = at + n - 1;
    return at <= end && last >= start && !(at <= start && last >= end);
  });
}

/**
 * A sheet's raw-cell reader with its tables' results read as their
 * `=TABLE(...)`: every calculation input of the sheet (the grid's, a save's,
 * the session snapshot's) reads through it, so the results are formulas to
 * the engine, the session's coordinate index and the grid's value store
 * alike, while the stored cells keep their values. Null without a table.
 * `whatIfTables` is the checked definitions the engine reads them with.
 */
export function whatIfTableReader(
  raw: unknown,
  bounds: { readonly rowCount: number; readonly columnCount: number },
  readRaw: (row: number, col: number) => string | null,
): {
  readonly readRaw: (row: number, col: number) => string | null;
  readonly whatIfTables: WhatIfTableDefinition[];
  readonly isResult: (row: number, col: number) => boolean;
} | null {
  const blocks = whatIfTableBlocks(raw, bounds);
  if (blocks.length === 0) return null;
  const formulas = blocks.map(whatIfTableFormula);
  const at = (row: number, col: number): number => {
    for (let index = 0; index < blocks.length; index += 1) {
      const block = blocks[index];
      if (row >= block.r1 && row <= block.r2 && col >= block.c1 && col <= block.c2) return index;
    }
    return -1;
  };
  return {
    readRaw: (row, col) => {
      const index = at(row, col);
      return index < 0 ? readRaw(row, col) : formulas[index];
    },
    whatIfTables: blocks.map(whatIfTableDefinition),
    isResult: (row, col) => at(row, col) >= 0,
  };
}

/**
 * Whether Insert / Delete Cells with a shift on `range` would move any of a
 * table: its results, its row or column of values, or an input cell. Only a
 * table the shift leaves where it is lets the shift run (MRAI does not move
 * a table with its cells; Excel refuses a shift through one).
 */
export function whatIfTableCellShiftRefused(
  raw: unknown,
  range: Rect,
  direction: "up" | "down" | "left" | "right",
): boolean {
  const vertical = direction === "up" || direction === "down";
  const moved: Rect = vertical
    ? { r1: range.r1, c1: range.c1, r2: MAX_GRID_ROWS - 1, c2: range.c2 }
    : { r1: range.r1, c1: range.c1, r2: range.r2, c2: MAX_GRID_COLUMNS - 1 };
  return whatIfTableBlocks(raw).some(
    (block) =>
      intersects(whatIfTableFootprint(block), moved) ||
      (block.rowInput !== null && contains(moved, block.rowInput)) ||
      (block.colInput !== null && contains(moved, block.colInput)),
  );
}
