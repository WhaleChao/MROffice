/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure grid-operation helpers for the spreadsheet (Excel/WPS table-stakes).
 *
 * These cover five daily-frequency interactions whose math is kept OUT of the
 * React component so each can be unit-tested in isolation and so the keyboard
 * shortcut, the toolbar button and the picker all share one source of truth:
 *
 *  - Fill Right (Ctrl+R): the horizontal mirror of Fill Down — the selection's
 *    LEFT column (or the cell to the left of a single active cell) is copied
 *    rightward, with relative formula references shifted by the column delta.
 *  - Selection aggregates: Excel's status bar offers Average / Count /
 *    Numerical Count / Minimum / Maximum / Sum over the selection;
 *    `aggregateSelection` walks a list of computed cell values and returns all
 *    six (which of them show is `status-bar-statistics.ts`).
 *  - Name Box navigation: typing a cell ("B50"), a range ("A1:C9") or a defined
 *    name into the address box jumps/selects there; `resolveNameBoxTarget`
 *    parses the text to a normalized rectangle (single cell ⇒ a 1×1 rect).
 *  - AutoSum (Σ / Alt+=): detect the contiguous numeric run directly ABOVE the
 *    active cell (else to its LEFT) and produce `=SUM(range)` plus the cell the
 *    formula should be written into; with a multi-cell selection the sum is
 *    written into the cell just below / to the right of the selection.
 *
 * Every helper is side-effect free. Cell access is threaded in via callbacks so
 * this module never imports the React state or the sheet model; A1 parsing /
 * formatting reuse the formula engine's shared helpers so the column-letter and
 * reference grammar agree everywhere.
 */

import {
  columnToLetters,
  isFormulaError,
  lettersToColumn,
  parseCellRef,
  shiftFormula,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { splitSheetQualifier } from "@mrnote/core/web/lib/spreadsheet/name-syntax.ts";

// ---------------------------------------------------------------------------
// Fill Right (Ctrl+R) — horizontal mirror of Fill Down
// ---------------------------------------------------------------------------

/**
 * The value to write into (row, targetCol) when filling right from the cell at
 * (row, srcCol). A formula's relative references are shifted by the column
 * delta (Excel parity); a plain value is copied verbatim. `getCell` returns the
 * source cell's stored { value, formula } (or undefined for an empty cell).
 */
/**
 * The minimal cell shape Fill Right reads: a stored display/formula `value` and
 * a truthy `formula` marker when that value is a formula. Matches the host
 * SpreadsheetCell (whose `formula` is the optional formula string — truthy ⇒ the
 * value is a formula), so the host can pass its cells through directly.
 */
export interface FillCell {
  value?: string;
  formula?: string | boolean;
}

export function fillRightValue(
  row: number,
  srcCol: number,
  targetCol: number,
  getCell: (row: number, col: number) => FillCell | undefined,
): string {
  const src = getCell(row, srcCol);
  const srcVal = src?.value ?? "";
  return src?.formula ? shiftFormula(srcVal, 0, targetCol - srcCol) : srcVal;
}

/** The four Home > Fill directions (Ctrl+D is down, Ctrl+R right). */
export type EdgeFillDirection = "down" | "up" | "right" | "left";

/** One filled cell: what it receives and the source cell it copies. */
export interface EdgeFillEntry {
  /** The value written: a formula moved by the displacement, else verbatim. */
  value: string;
  /** Whether the source cell is a formula (its identity, not its text). */
  formula: boolean;
  sourceRow: number;
  sourceCol: number;
}

/**
 * Fill Down / Up / Right / Left over a selection (Excel Home > Fill): the
 * selection's edge row or column on the side the fill comes from is copied
 * into the rest of the selection. Each entry names its source cell so the
 * caller can copy the source's formats with it (Excel: "the contents and
 * format"). A formula's relative references move by the displacement; any
 * other value, a text reading "=A2" included, is copied as it is. The grid is
 * anchored at (row, col); null when the selection has no cell to fill.
 */
export function computeEdgeFill(
  sel: { r1: number; c1: number; r2: number; c2: number },
  direction: EdgeFillDirection,
  getCell: (row: number, col: number) => FillCell | undefined,
): { row: number; col: number; grid: EdgeFillEntry[][] } | null {
  const vertical = direction === "down" || direction === "up";
  if (vertical ? sel.r2 <= sel.r1 : sel.c2 <= sel.c1) return null;
  const row = direction === "down" ? sel.r1 + 1 : sel.r1;
  const col = direction === "right" ? sel.c1 + 1 : sel.c1;
  const lastRow = direction === "up" ? sel.r2 - 1 : sel.r2;
  const lastCol = direction === "left" ? sel.c2 - 1 : sel.c2;
  const grid: EdgeFillEntry[][] = [];
  for (let r = row; r <= lastRow; r += 1) {
    const line: EdgeFillEntry[] = [];
    for (let c = col; c <= lastCol; c += 1) {
      const sourceRow =
        direction === "down" ? sel.r1 : direction === "up" ? sel.r2 : r;
      const sourceCol =
        direction === "right" ? sel.c1 : direction === "left" ? sel.c2 : c;
      const source = getCell(sourceRow, sourceCol);
      const formula = Boolean(source?.formula);
      const raw = source?.value ?? "";
      line.push({
        value: formula
          ? shiftFormula(raw, r - sourceRow, c - sourceCol)
          : raw,
        formula,
        sourceRow,
        sourceCol,
      });
    }
    grid.push(line);
  }
  return { row, col, grid };
}

/**
 * Build the grid (rows × cols) that Fill Right writes to the right of the
 * selection's LEFT column. The returned grid is anchored at (selr1, selc1 + 1)
 * and spans every row of the selection and every column from selc1+1..selc2.
 * Returns an empty array when the selection is a single column (nothing to
 * fill). The values of {@link computeEdgeFill}.
 */
export function computeFillRightGrid(
  sel: { r1: number; c1: number; r2: number; c2: number },
  getCell: (row: number, col: number) => FillCell | undefined,
): string[][] {
  const fill = computeEdgeFill(sel, "right", getCell);
  return fill ? fill.grid.map((line) => line.map((entry) => entry.value)) : [];
}

/**
 * Build the grid that Fill Left writes — the horizontal mirror of
 * {@link computeFillRightGrid}. The selection's RIGHT column (selc2) is the
 * source, filled leftward into every column from selc1..selc2-1; the grid is
 * anchored at (selr1, selc1). A formula's relative refs shift by the (negative)
 * column delta. Returns an empty array for a single-column selection (nothing
 * to fill).
 */
export function computeFillLeftGrid(
  sel: { r1: number; c1: number; r2: number; c2: number },
  getCell: (row: number, col: number) => FillCell | undefined,
): string[][] {
  const fill = computeEdgeFill(sel, "left", getCell);
  return fill ? fill.grid.map((line) => line.map((entry) => entry.value)) : [];
}

// ---------------------------------------------------------------------------
// Selection aggregates — Average / Count / Numerical Count / Min / Max / Sum
// ---------------------------------------------------------------------------

/** The six aggregates Excel's status bar reports over a selection. */
export interface SelectionAggregates {
  /** Total of every numeric cell. */
  sum: number;
  /** Mean of the numeric cells (0 when none). */
  average: number;
  /** Count of NON-EMPTY cells (numbers, text, TRUE/FALSE and errors), Excel's
   * "Count". */
  count: number;
  /** Count of numeric cells only, Excel's "Numerical Count". */
  countNumbers: number;
  /** Count of error values (#N/A, #DIV/0!, ...). Excel still counts them, but
   * an error anywhere in the selection leaves Average, Min, Max and Sum off the
   * status bar, the same way =SUM over it would be an error. */
  errors: number;
  /** Minimum numeric value, or null when no numeric cell is present. */
  min: number | null;
  /** Maximum numeric value, or null when no numeric cell is present. */
  max: number | null;
}

/**
 * Aggregate a flat list of computed cell values (one entry per selected cell)
 * into the status-bar figures. A value is "numeric" when it is a finite number;
 * it is "non-empty" (counted by `count`) when it is a finite number, any
 * non-empty string (a space too, and a formula's "" when the caller says it is
 * one), a logical or an error — matching
 * Excel, where Count tallies what COUNTA does and Numerical Count only the
 * numeric ones. An empty cell (null/undefined/"") and non-finite numbers
 * contribute to neither count.
 */
/** One ordered accumulator for synchronous and yielded read-only summaries.
 * Do not combine partial sums: preserving the original addition order avoids
 * changing floating-point results when a large selection is scheduled. */
export function createSelectionAggregateAccumulator() {
  let sum = 0;
  let count = 0;
  let countNumbers = 0;
  let errors = 0;
  let min: number | null = null;
  let max: number | null = null;
  return {
    /** `formulaResult`: the value is a formula's (or its spill's) result, so
     * an empty string is text the formula returned, not an empty cell. */
    add(value: unknown, formulaResult = false) {
      if (typeof value === "number" && Number.isFinite(value)) {
        sum += value;
        countNumbers += 1;
        count += 1;
        if (min === null || value < min) min = value;
        if (max === null || value > max) max = value;
      } else if (typeof value === "string" && (value !== "" || formulaResult)) {
        // Any text counts, a single space and a formula's "" included: Excel's
        // Count tallies what COUNTA does (audit P-STATUSBAR-EMPTY-SPACE-COUNT).
        // A bare "" from no formula is a blank, as an empty cell is.
        count += 1;
      } else if (typeof value === "boolean") {
        count += 1;
      } else if (isFormulaError(value)) {
        count += 1;
        errors += 1;
      }
    },
    result(): SelectionAggregates {
      return { sum, average: countNumbers ? sum / countNumbers : 0, count, countNumbers, errors, min, max };
    },
  };
}

export function aggregateSelection(values: ReadonlyArray<unknown>): SelectionAggregates {
  const accumulator = createSelectionAggregateAccumulator();
  for (const value of values) accumulator.add(value);
  return accumulator.result();
}

// ---------------------------------------------------------------------------
// Name Box navigation — parse "B50" / "A1:C9" / a defined name to a rectangle
// ---------------------------------------------------------------------------

/** A normalized inclusive rectangle plus the cell to make active. */
export interface NameBoxTarget {
  /** Active cell after the jump (top-left of the resolved rectangle). */
  row: number;
  col: number;
  /** Inclusive selection rectangle (a single cell ⇒ r1==r2, c1==c2). */
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** `C:C` / `3:5`: the selection spans the whole line (the caller sizes it
   *  to the grid it shows rather than a million cells; B500). */
  wholeColumns?: true;
  wholeRows?: true;
  /**
   * The sheet named before `!` — a typed `回款!C4` or a defined name stored
   * the way Excel stores one (`回款!$C$4:$C$10`). Absent ⇒ the active sheet.
   */
  sheet?: string;
}

// A range "A1:C9" (each side optionally $-absolute), matching the formula
// engine / named-ranges grammar so the Name Box accepts exactly what a formula
// reference would.
const RANGE_REF_RE =
  /^(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}):(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})$/;

/** Clamp a 0-based (row, col) into the addressable grid. */
function clampCell(row: number, col: number): { row: number; col: number } {
  return {
    row: Math.max(0, Math.min(row, MAX_GRID_ROWS - 1)),
    col: Math.max(0, Math.min(col, MAX_GRID_COLUMNS - 1)),
  };
}

/** Build a normalized NameBoxTarget from a parsed corner pair (any order). */
function targetFromCorners(
  a: { row: number; col: number },
  b: { row: number; col: number },
): NameBoxTarget {
  const lo = clampCell(Math.min(a.row, b.row), Math.min(a.col, b.col));
  const hi = clampCell(Math.max(a.row, b.row), Math.max(a.col, b.col));
  return {
    row: lo.row,
    col: lo.col,
    r1: lo.row,
    c1: lo.col,
    r2: hi.row,
    c2: hi.col,
  };
}

/**
 * Resolve a Name Box entry to a selection rectangle, or null when it is neither
 * a cell reference, a range reference, nor a defined name. Resolution order
 * (Excel parity): a literal cell/range — optionally sheet-qualified, the target
 * then carrying `sheet` — is parsed directly; otherwise the text
 * is looked up as a defined name (case-insensitive) and its stored reference is
 * parsed. A defined name whose ref no longer parses yields null. Coordinates
 * are clamped to the grid so an out-of-range "Z99999" still lands on a real
 * cell rather than off the sheet.
 */
export function resolveNameBoxTarget(
  text: string,
  resolveName: (name: string) => string | null,
): NameBoxTarget | null {
  if (typeof text !== "string") return null;
  const trimmed = text.trim();
  if (trimmed === "") return null;

  const fromRef = parseQualifiedRefToTarget(trimmed);
  if (fromRef) return fromRef;

  // Not a literal ref → try a defined name, then parse what it expands to
  // (a stored refers-to may keep Excel's leading "=").
  const ref = resolveName(trimmed);
  if (ref) return parseQualifiedRefToTarget(ref.trim().replace(/^=\s*/, ""));
  return null;
}

/** A cell / range, optionally sheet-qualified (`Sheet2!A1`, `'Q 1'!$A$1:$B$2`). */
function parseQualifiedRefToTarget(ref: string): NameBoxTarget | null {
  const { sheet, ref: body } = splitSheetQualifier(ref.trim());
  const target = parseRefToTarget(body);
  if (!target || sheet === undefined) return target;
  return sheet === "" ? null : { ...target, sheet };
}

/** Parse a literal cell ("B50") or range ("A1:C9") into a target, else null. */
const COLUMN_RANGE_REF_RE = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/;
const ROW_RANGE_REF_RE = /^\$?([0-9]{1,7}):\$?([0-9]{1,7})$/;

function parseRefToTarget(ref: string): NameBoxTarget | null {
  const trimmed = ref.trim();
  // A whole column ($C:$C) or row (3:5), as a defined name of a column reads
  // and as the Name Box accepts a formula reference (B500).
  const columns = COLUMN_RANGE_REF_RE.exec(trimmed);
  if (columns) {
    const from = parseCellRef(`${columns[1]}1`);
    const to = parseCellRef(`${columns[2]}1`);
    if (!from || !to) return null;
    const target = targetFromCorners(
      { row: 0, col: from.col },
      { row: MAX_GRID_ROWS - 1, col: to.col },
    );
    return { ...target, wholeColumns: true };
  }
  const rows = ROW_RANGE_REF_RE.exec(trimmed);
  if (rows) {
    const from = Number(rows[1]) - 1;
    const to = Number(rows[2]) - 1;
    if (from < 0 || to < 0) return null;
    const target = targetFromCorners(
      { row: from, col: 0 },
      { row: to, col: MAX_GRID_COLUMNS - 1 },
    );
    return { ...target, wholeRows: true };
  }
  const rangeMatch = RANGE_REF_RE.exec(trimmed);
  if (rangeMatch) {
    const from = parseCellRef(rangeMatch[1]);
    const to = parseCellRef(rangeMatch[2]);
    if (!from || !to) return null;
    return targetFromCorners(from, to);
  }
  const cell = parseCellRef(trimmed);
  if (cell) return targetFromCorners(cell, cell);
  return null;
}

// ---------------------------------------------------------------------------
// AutoSum (Σ / Alt+=) — contiguous-numeric-region detection
// ---------------------------------------------------------------------------

/** Where to write an AutoSum formula and the formula text itself. */
export interface AutoSumTarget {
  /** Cell the =SUM(...) formula is written into. */
  writeRow: number;
  writeCol: number;
  /** The formula string, e.g. "=SUM(A1:A5)". */
  formula: string;
}

/**
 * The aggregate functions Excel's AutoSum dropdown offers (the Σ ▾ menu): Sum,
 * Average, Count Numbers, Max, Min. Each maps 1:1 to an engine function name, so
 * AutoSum just swaps the wrapper around the SAME detected/selected range.
 * "count" emits COUNT (count of NUMBERS, Excel's "Count Numbers" — matching the
 * numeric-run detection); a text count would use COUNTA, which AutoSum does not.
 */
export type AutoSumFn = "sum" | "average" | "count" | "max" | "min";

/** Map an AutoSum function id to its engine function NAME (the formula wrapper). */
export function autoSumFnName(fn: AutoSumFn): string {
  switch (fn) {
    case "average":
      return "AVERAGE";
    case "count":
      return "COUNT";
    case "max":
      return "MAX";
    case "min":
      return "MIN";
    case "sum":
      return "SUM";
  }
}

/** Format a range's two corners as an A1 reference ("A1:A5" or "A1" for 1×1). */
function rangeRef(r1: number, c1: number, r2: number, c2: number): string {
  const a = `${columnToLetters(c1)}${r1 + 1}`;
  if (r1 === r2 && c1 === c2) return a;
  const b = `${columnToLetters(c2)}${r2 + 1}`;
  return `${a}:${b}`;
}

/**
 * Walk UP from (startRow-1, col) while the cell is numeric; return the topmost
 * numeric row, or null when the cell directly above is not numeric. `isNumeric`
 * reports whether a cell currently holds a finite number (computed value).
 */
function numericRunAbove(
  startRow: number,
  col: number,
  isNumeric: (row: number, col: number) => boolean,
): number | null {
  let r = startRow - 1;
  if (r < 0 || !isNumeric(r, col)) return null;
  while (r - 1 >= 0 && isNumeric(r - 1, col)) r -= 1;
  return r;
}

/**
 * Walk LEFT from (row, startCol-1) while the cell is numeric; return the
 * leftmost numeric column, or null when the cell directly to the left is not
 * numeric.
 */
function numericRunLeft(
  row: number,
  startCol: number,
  isNumeric: (row: number, col: number) => boolean,
): number | null {
  let c = startCol - 1;
  if (c < 0 || !isNumeric(row, c)) return null;
  while (c - 1 >= 0 && isNumeric(row, c - 1)) c -= 1;
  return c;
}

/**
 * Compute the AutoSum action for the current selection/active cell, or null
 * when there is nothing to sum.
 *
 * Single active cell: prefer the contiguous numeric run directly ABOVE the cell
 * (Excel's first choice) and write =SUM(<that column range>) into the active
 * cell; otherwise the run to its LEFT. With neither, return null.
 *
 * Multi-cell selection: sum the WHOLE selection into the cell just BELOW the
 * selection (a single-column / general selection) or just to the RIGHT (a
 * single-row selection). {@link autoSumTargets} is the caller's planner: it
 * totals a block column by column and fills a blank total row or column
 * inside the selection, as Excel does; this single target is its fallback
 * for a selection of one line. The write target is clamped to the grid; the
 * formula always references the original selection.
 */
export function autoSumTarget(
  active: { row: number; col: number },
  sel: { r1: number; c1: number; r2: number; c2: number },
  hasMultiSelection: boolean,
  isNumeric: (row: number, col: number) => boolean,
  // Which aggregate to wrap the detected/selected range in (Σ dropdown). Defaults
  // to "sum" so existing callers keep producing =SUM(...) unchanged.
  fn: AutoSumFn = "sum",
): AutoSumTarget | null {
  const name = autoSumFnName(fn);
  if (hasMultiSelection) {
    const singleRow = sel.r1 === sel.r2;
    // A single ROW of cells totals to the right; anything taller totals below.
    const writeRow = singleRow ? sel.r1 : sel.r2 + 1;
    const writeCol = singleRow ? sel.c2 + 1 : sel.c1;
    const target = clampCell(writeRow, writeCol);
    return {
      writeRow: target.row,
      writeCol: target.col,
      formula: `=${name}(${rangeRef(sel.r1, sel.c1, sel.r2, sel.c2)})`,
    };
  }

  const top = numericRunAbove(active.row, active.col, isNumeric);
  if (top !== null) {
    return {
      writeRow: active.row,
      writeCol: active.col,
      formula: `=${name}(${rangeRef(top, active.col, active.row - 1, active.col)})`,
    };
  }
  const left = numericRunLeft(active.row, active.col, isNumeric);
  if (left !== null) {
    return {
      writeRow: active.row,
      writeCol: active.col,
      formula: `=${name}(${rangeRef(active.row, left, active.row, active.col - 1)})`,
    };
  }
  return null;
}

/**
 * Plan the complete AutoSum write set for the current selection.
 *
 * Excel lets a user select several EMPTY result cells (for example B7:C7) and
 * invoke AutoSum once to total both columns. Each destination owns an
 * independent above-then-left inference; treating that selection as source data
 * would instead write one incorrect formula to D7. A blank destination batch is
 * all-or-nothing: if any selected cell has no numeric run, no target is
 * authorized. Non-blank selections retain the existing source-range behavior.
 */
export function autoSumTargets(
  active: { row: number; col: number },
  sel: { r1: number; c1: number; r2: number; c2: number },
  hasMultiSelection: boolean,
  isNumeric: (row: number, col: number) => boolean,
  isEmpty: (row: number, col: number) => boolean,
  fn: AutoSumFn = "sum",
  /** A row a filter hides. Over a filtered range Excel's AutoSum writes
   * SUBTOTAL(9, ...) (1 average, 2 count, 4 max, 5 min), so the total is of
   * what the filter shows (audit P-AUTOSUM-FILTERED-RANGE). */
  isRowFiltered?: (row: number) => boolean,
): AutoSumTarget[] {
  const targets = plainAutoSumTargets(active, sel, hasMultiSelection, isNumeric, isEmpty, fn);
  return isRowFiltered ? targets.map((target) => subtotalOverFilteredRows(target, fn, isRowFiltered)) : targets;
}

/** The SUBTOTAL function number of each AutoSum aggregate. */
const AUTO_SUM_SUBTOTAL_CODE: Record<AutoSumFn, number> = { average: 1, count: 2, max: 4, min: 5, sum: 9 };

function subtotalOverFilteredRows(
  target: AutoSumTarget,
  fn: AutoSumFn,
  isRowFiltered: (row: number) => boolean,
): AutoSumTarget {
  const match = /^=[A-Z]+\(([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?\)$/.exec(target.formula);
  if (!match) return target;
  const first = Number(match[2]) - 1;
  const last = match[4] === undefined ? first : Number(match[4]) - 1;
  for (let row = Math.min(first, last); row <= Math.max(first, last); row += 1) {
    if (isRowFiltered(row)) {
      const ref = target.formula.slice(target.formula.indexOf("(") + 1, -1);
      return { ...target, formula: `=SUBTOTAL(${AUTO_SUM_SUBTOTAL_CODE[fn]},${ref})` };
    }
  }
  return target;
}

function plainAutoSumTargets(
  active: { row: number; col: number },
  sel: { r1: number; c1: number; r2: number; c2: number },
  hasMultiSelection: boolean,
  isNumeric: (row: number, col: number) => boolean,
  isEmpty: (row: number, col: number) => boolean,
  fn: AutoSumFn,
): AutoSumTarget[] {
  if (hasMultiSelection) {
    let allDestinationsEmpty = true;
    for (let row = sel.r1; row <= sel.r2; row += 1) {
      for (let col = sel.c1; col <= sel.c2; col += 1) {
        if (!isEmpty(row, col)) {
          allDestinationsEmpty = false;
          break;
        }
      }
      if (!allDestinationsEmpty) break;
    }

    if (allDestinationsEmpty) {
      const name = autoSumFnName(fn);
      const targets: AutoSumTarget[] = [];
      for (let row = sel.r1; row <= sel.r2; row += 1) {
        for (let col = sel.c1; col <= sel.c2; col += 1) {
          const top = numericRunAbove(row, col, isNumeric);
          if (top !== null) {
            targets.push({
              writeRow: row,
              writeCol: col,
              formula: `=${name}(${rangeRef(top, col, row - 1, col)})`,
            });
            continue;
          }
          const left = numericRunLeft(row, col, isNumeric);
          if (left === null) return [];
          targets.push({
            writeRow: row,
            writeCol: col,
            formula: `=${name}(${rangeRef(row, left, row, col - 1)})`,
          });
        }
      }
      return targets;
    }
  }

  if (hasMultiSelection && sel.r1 !== sel.r2 && sel.c1 !== sel.c2) {
    // A block of data: Excel totals it column by column. A blank last row
    // (or column) inside the selection is the total row (column) and gets
    // the sums; otherwise they land in the row below. Both blank edges give
    // the corner the total of the column totals.
    const name = autoSumFnName(fn);
    const rowEmpty = (row: number) => {
      for (let col = sel.c1; col <= sel.c2; col += 1) if (!isEmpty(row, col)) return false;
      return true;
    };
    const colEmpty = (col: number) => {
      for (let row = sel.r1; row <= sel.r2; row += 1) if (!isEmpty(row, col)) return false;
      return true;
    };
    const bottomBlank = rowEmpty(sel.r2);
    const rightBlank = colEmpty(sel.c2);
    const lastDataRow = bottomBlank ? sel.r2 - 1 : sel.r2;
    const lastDataCol = rightBlank ? sel.c2 - 1 : sel.c2;
    const targets: AutoSumTarget[] = [];
    if (bottomBlank || !rightBlank) {
      const writeRow = bottomBlank ? sel.r2 : sel.r2 + 1;
      for (let col = sel.c1; col <= lastDataCol; col += 1) {
        const target = clampCell(writeRow, col);
        targets.push({
          writeRow: target.row,
          writeCol: target.col,
          formula: `=${name}(${rangeRef(sel.r1, col, lastDataRow, col)})`,
        });
      }
    }
    if (rightBlank) {
      for (let row = sel.r1; row <= lastDataRow; row += 1) {
        targets.push({
          writeRow: row,
          writeCol: sel.c2,
          formula: `=${name}(${rangeRef(row, sel.c1, row, lastDataCol)})`,
        });
      }
      if (bottomBlank) {
        targets.push({
          writeRow: sel.r2,
          writeCol: sel.c2,
          formula: `=${name}(${rangeRef(sel.r2, sel.c1, sel.r2, lastDataCol)})`,
        });
      }
    }
    return targets;
  }
  if (hasMultiSelection && sel.c1 === sel.c2 && sel.r1 !== sel.r2 && isEmpty(sel.r2, sel.c1)) {
    // One column ending in a blank cell: that cell is the total.
    return [{
      writeRow: sel.r2,
      writeCol: sel.c1,
      formula: `=${autoSumFnName(fn)}(${rangeRef(sel.r1, sel.c1, sel.r2 - 1, sel.c1)})`,
    }];
  }
  if (hasMultiSelection && sel.r1 === sel.r2 && sel.c1 !== sel.c2 && isEmpty(sel.r1, sel.c2)) {
    // One row ending in a blank cell: that cell is the total.
    return [{
      writeRow: sel.r1,
      writeCol: sel.c2,
      formula: `=${autoSumFnName(fn)}(${rangeRef(sel.r1, sel.c1, sel.r1, sel.c2 - 1)})`,
    }];
  }
  const target = autoSumTarget(
    active,
    sel,
    hasMultiSelection,
    isNumeric,
    fn,
  );
  return target ? [target] : [];
}

// ---------------------------------------------------------------------------
// Re-exports used by callers that already depend on this module's grammar
// ---------------------------------------------------------------------------

export { lettersToColumn, columnToLetters };
