/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Go To Special (Excel F5 ▸ Special, WPS 定位条件).
 *
 * The power feature behind half the spreadsheet tricks people learn: select
 * every blank in a range so you can fill them in one go, every formula so you
 * can protect them, every constant so you can clear them without touching the
 * arithmetic. Pure predicates over a cell map, so the dialog stays a renderer
 * and every rule is testable on its own.
 *
 * "Visible cells only" is the one with teeth: copying a filtered list without
 * it silently copies the rows the filter is hiding, which is how people paste
 * data they never saw.
 *
 * The options follow Excel's dialog (Microsoft Support, "Find and select
 * cells that meet specific conditions"). Precedents and Dependents
 * (XLS-A062) read a formula's same-sheet references the way Trace
 * Precedents does: the engine's exact reference set when it has one, else
 * the A1 cells and ranges the formula names. Defined names and structured
 * references are not followed, which is the trace arrows' grammar too.
 * Current array and Objects are answered by the window, which owns the spill
 * map and the drawing selection.
 */

import { formulaA1ToR1C1 } from "./r1c1";
import {
  dependencyRangesOf,
  isExcelNumericText,
  isFormulaError,
} from "./formula-engine";
import { extractPrecedents } from "./sheet-interactions";
import {
  surroundingDataRegion,
  type DataRegionRange,
  regionCellHoldsData,
} from "./data-region";
import { sameValidationSettingsAs, type ValidationRule } from "./data-validation";
import { spreadsheetCellIsText } from "./entry-parse";
import { cellMapChangesSince } from "./cell-map-lineage";

export type GoToSpecialKind =
  | "blanks"
  | "constants"
  | "formulas"
  | "errors"
  | "comments"
  | "visible"
  | "validation"
  | "conditionalFormats"
  | "rowDifferences"
  | "columnDifferences";

/**
 * Every choice the dialog offers: the cell scans above, plus the two options
 * that name a rectangle outright instead of testing cells.
 */
export type GoToSpecialOption =
  | GoToSpecialKind
  | "currentRegion"
  | "lastCell"
  // XLS-A062: Excel's four further choices.
  | "currentArray"
  | "objects"
  | "precedents"
  | "dependents";

/**
 * What a Constants or Formulas search keeps, by value type: Excel's Numbers,
 * Text, Logicals and Errors check boxes under Formulas, which apply to
 * Constants too.
 */
export interface GoToSpecialValueTypes {
  numbers: boolean;
  text: boolean;
  logicals: boolean;
  errors: boolean;
}

export const ALL_GO_TO_SPECIAL_VALUE_TYPES: Readonly<GoToSpecialValueTypes> =
  Object.freeze({ numbers: true, text: true, logicals: true, errors: true });

export interface GoToSpecialRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface GoToSpecialCell {
  value?: unknown;
  formula?: unknown;
  display?: unknown;
  comment?: unknown;
  style?: Record<string, unknown>;
  /** The entry's stored kind: "number", or "text" for everything else. */
  kind?: unknown;
  /** The workbook file typed the value as a string (live audit B312). */
  stringValue?: unknown;
}

export interface GoToSpecialInput {
  cells: Record<string, GoToSpecialCell | undefined>;
  range: GoToSpecialRange;
  kind: GoToSpecialKind;
  /** Rows the filter or a collapsed group is hiding. */
  hiddenRows?: ReadonlySet<number>;
  /** Columns currently hidden. */
  hiddenCols?: ReadonlySet<number>;
  /** Keys a dynamic array has spilled a value into: they show a result and
   *  are not blanks to fill, though no cell is stored there (live audit B278). */
  spilledKeys?: { has(key: string): boolean };
  /** Keys carrying a validation rule, when the caller tracks them per cell. */
  validationKeys?: ReadonlySet<string>;
  /** Keys a conditional-format rule currently covers. */
  conditionalKeys?: ReadonlySet<string>;
  /** Constants/Formulas: the value types to keep (all when omitted). */
  types?: GoToSpecialValueTypes;
  /**
   * The engine's value for a cell: a formula's result, a literal's typed
   * value. Classifies Numbers/Text/Logicals/Errors the way the rest of the
   * editor (the status bar's Count Numbers) reads a cell. Without it the
   * stored text is classified on its own.
   */
  computedAt?: (row: number, col: number) => unknown;
  /**
   * Row/Column differences: the active cell, whose column (rows) or row
   * (columns) holds each line's comparison cell.
   */
  active?: { row: number; col: number };
}

/**
 * Hard cap on what one scan returns. A caller that hands in a whole declared
 * grid (10,000 x 1,000) would otherwise get ten million keys back -- five
 * seconds and half a gigabyte of strings, which in a browser is a hang. Excel
 * bounds Go To Special to the used range; this bounds it regardless of what the
 * caller passes, so a bad range can never take the tab down.
 */
export const MAX_SPECIAL_CELLS = 50_000;

/** An Excel error literal, which counts as an error cell however it arrived. */
const ERROR_RE = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CIRCULAR!)$/;

function text(cell: GoToSpecialCell | undefined): string {
  const raw = cell?.value;
  return raw === undefined || raw === null ? "" : String(raw);
}

/**
 * Whether a literal is TEXT whatever its characters read as: typed after an
 * apostrophe, into a Text (@) cell, or typed so by the file -- the engine's
 * own test (spreadsheetCellIsText). The stored kind is not that test: a
 * plain TRUE, #N/A or "=..." entry is stored with kind "text" as well, and
 * reads as a logical, an error or a formula (live audit S-041).
 */
function typedAsText(cell: GoToSpecialCell | undefined): boolean {
  return spreadsheetCellIsText(cell as Parameters<typeof spreadsheetCellIsText>[0]);
}

function isFormula(cell: GoToSpecialCell | undefined): boolean {
  if (typeof cell?.formula === "string" && cell.formula.startsWith("=")) {
    return true;
  }
  // An entry stored as Text ('=1+1) is text, whatever it starts with.
  if (typedAsText(cell)) return false;
  return text(cell).startsWith("=");
}

function isError(cell: GoToSpecialCell | undefined): boolean {
  // A formula's error shows in its computed display; a literal error can also
  // be typed straight into a cell.
  const display = cell?.display;
  if (typeof display === "string" && ERROR_RE.test(display)) return true;
  return ERROR_RE.test(text(cell));
}

type ValueType = keyof GoToSpecialValueTypes;

function literalType(raw: string): ValueType {
  const trimmed = raw.trim();
  if (ERROR_RE.test(trimmed)) return "errors";
  const upper = trimmed.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") return "logicals";
  return isExcelNumericText(trimmed) ? "numbers" : "text";
}

/**
 * Excel's type of a non-empty cell. A formula is typed by its result (an
 * empty-string result is Text, an empty reference reads as 0); a constant
 * by its value, where a typed error literal such as #N/A is an error.
 */
function valueType(
  cell: GoToSpecialCell | undefined,
  computed: unknown,
  hasComputed: boolean,
): ValueType {
  const raw = text(cell);
  if (!isFormula(cell)) {
    // A Text entry that spells an error ('#N/A) or a logical ('TRUE) is
    // text; a plain TRUE or #N/A is the logical or error it reads as.
    if (typedAsText(cell)) return "text";
    if (ERROR_RE.test(raw.trim())) return "errors";
    if (!hasComputed) return literalType(raw);
    if (typeof computed === "boolean") return "logicals";
    if (typeof computed === "number") return "numbers";
    if (isFormulaError(computed)) return "errors";
    return "text";
  }
  if (hasComputed) {
    if (isFormulaError(computed)) return "errors";
    if (typeof computed === "boolean") return "logicals";
    if (typeof computed === "number" || computed === null) return "numbers";
    // A formula's string result is text even when it spells an error
    // (="#N/A" is text; only a FormulaError is an error).
    if (typeof computed === "string") return "text";
  }
  const display = typeof cell?.display === "string" ? cell.display : "";
  return display === "" ? "text" : literalType(display);
}

/**
 * What Row/Column differences compares. A formula is compared in its
 * relative (R1C1) form, so a formula filled along the line is the same
 * formula; a constant by its value, numbers numerically and text without
 * regard to case, as Excel's comparison does.
 */
function differenceSignature(
  cell: GoToSpecialCell | undefined,
  row: number,
  col: number,
): string {
  if (isFormula(cell)) {
    const formula =
      typeof cell?.formula === "string" && cell.formula.startsWith("=")
        ? cell.formula
        : text(cell);
    return `f:${formulaA1ToR1C1(formula, { row, col }).toUpperCase()}`;
  }
  const raw = text(cell);
  if (raw === "") return "b:";
  const trimmed = raw.trim();
  if (isExcelNumericText(trimmed)) return `n:${Number(trimmed)}`;
  return `v:${raw.toLowerCase()}`;
}

function hasComment(cell: GoToSpecialCell | undefined): boolean {
  const comment = cell?.comment;
  if (!comment) return false;
  if (typeof comment === "string") return comment.trim() !== "";
  if (typeof comment === "object") {
    const record = comment as Record<string, unknown>;
    const body = record.text;
    if (typeof body === "string" && body.trim() !== "") return true;
    return Array.isArray(record.messages) && record.messages.length > 0;
  }
  return false;
}

/**
 * The cell keys in `range` matching `kind`, in row-major order.
 *
 * A cell that does not exist is blank, so "blanks" walks the range rather than
 * the map — the whole point is to find the holes.
 */
export function findSpecialCells(input: GoToSpecialInput): string[] {
  const { cells, range, kind } = input;
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  const hiddenRows = input.hiddenRows ?? new Set<number>();
  const hiddenCols = input.hiddenCols ?? new Set<number>();
  const types = input.types ?? ALL_GO_TO_SPECIAL_VALUE_TYPES;
  const allTypes = types.numbers && types.text && types.logicals && types.errors;
  const typed = (cell: GoToSpecialCell | undefined, row: number, col: number) =>
    allTypes ||
    types[
      valueType(
        cell,
        input.computedAt?.(row, col),
        input.computedAt !== undefined,
      )
    ];
  const active = input.active;
  const found: string[] = [];

  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      const key = `${row}:${col}`;
      const cell = cells[key];
      let match = false;
      switch (kind) {
        case "blanks":
          match = text(cell) === "" && !input.spilledKeys?.has(key);
          break;
        case "constants":
          // A constant is a value that is not the result of a formula.
          match = text(cell) !== "" && !isFormula(cell) && typed(cell, row, col);
          break;
        case "formulas":
          match = isFormula(cell) && typed(cell, row, col);
          break;
        case "rowDifferences":
          // Each row against its own cell in the active cell's column.
          match =
            active !== undefined &&
            col !== active.col &&
            differenceSignature(cell, row, col) !==
              differenceSignature(cells[`${row}:${active.col}`], row, active.col);
          break;
        case "columnDifferences":
          match =
            active !== undefined &&
            row !== active.row &&
            differenceSignature(cell, row, col) !==
              differenceSignature(cells[`${active.row}:${col}`], active.row, col);
          break;
        case "errors":
          match = isError(cell);
          break;
        case "comments":
          match = hasComment(cell);
          break;
        case "visible":
          match = !hiddenRows.has(row) && !hiddenCols.has(col);
          break;
        case "validation":
          match = input.validationKeys?.has(key) === true;
          break;
        case "conditionalFormats":
          match = input.conditionalKeys?.has(key) === true;
          break;
        default:
          match = false;
      }
      if (!match) continue;
      found.push(key);
      if (found.length >= MAX_SPECIAL_CELLS) return found;
    }
  }
  return found;
}

export interface GoToSpecialSelection {
  keys: string[];
  /** The bounding rectangle of the matches, or null when there were none. */
  bounds: GoToSpecialRange | null;
  /**
   * True only when the matches fill `bounds` exactly, so selecting `bounds`
   * selects the matches and nothing else. Excel answers every other shape
   * with a multi-area selection; a caller that can only select one rectangle
   * must not select `bounds` then, or a later Ctrl+Enter / Delete / format
   * reaches the cells in between that did not match.
   */
  rectangle: boolean;
  /** True when the scan stopped at {@link MAX_SPECIAL_CELLS}. */
  truncated: boolean;
}

/**
 * The matches plus the rectangle that encloses them, which is what a caller
 * needs to move the selection. Returns a null rectangle rather than an empty
 * one when nothing matched, so the caller reports "no cells found" instead of
 * silently selecting the first cell.
 */
export function goToSpecial(input: GoToSpecialInput): GoToSpecialSelection {
  return goToSpecialSelectionOf(findSpecialCells(input));
}

/** The selection a set of distinct matching keys makes (see goToSpecial). */
export function goToSpecialSelectionOf(keys: string[]): GoToSpecialSelection {
  const truncated = keys.length >= MAX_SPECIAL_CELLS;
  if (keys.length === 0) {
    return { keys, bounds: null, rectangle: false, truncated };
  }
  let r1 = Number.POSITIVE_INFINITY;
  let c1 = Number.POSITIVE_INFINITY;
  let r2 = Number.NEGATIVE_INFINITY;
  let c2 = Number.NEGATIVE_INFINITY;
  for (const key of keys) {
    const [rowText, colText] = key.split(":");
    const row = Number(rowText);
    const col = Number(colText);
    r1 = Math.min(r1, row);
    c1 = Math.min(c1, col);
    r2 = Math.max(r2, row);
    c2 = Math.max(c2, col);
  }
  // The scan visits each coordinate once, so the keys are distinct and the
  // matches fill their bounds exactly when the counts agree.
  const rectangle =
    !truncated && keys.length === (r2 - r1 + 1) * (c2 - c1 + 1);
  return { keys, bounds: { r1, c1, r2, c2 }, rectangle, truncated };
}

/** A formula cell's text, or null for a literal (a Text '=1+1 included). */
function formulaText(cell: GoToSpecialCell | undefined): string | null {
  if (!isFormula(cell)) return null;
  if (typeof cell?.formula === "string" && cell.formula.startsWith("=")) {
    return cell.formula;
  }
  return text(cell);
}

/**
 * The same-sheet rectangles a formula reads (XLS-A062): the engine's exact
 * dependency set when it can prove one, else every A1 cell and range the
 * formula names (Trace Precedents' grammar), so a formula that also calls
 * NOW() or names another sheet still reports its local references.
 */
export function formulaReferenceRanges(formula: string): GoToSpecialRange[] {
  const exact = dependencyRangesOf(formula);
  if (exact) return exact;
  return extractPrecedents(formula, MAX_SPECIAL_CELLS).map((cell) => ({
    r1: cell.row,
    c1: cell.col,
    r2: cell.row,
    c2: cell.col,
  }));
}

export interface GoToSpecialReferenceInput {
  cells: Record<string, GoToSpecialCell | undefined>;
  /** The selection whose formulas (precedents) or cells (dependents) start the search. */
  range: GoToSpecialRange;
  /** Excel's "Direct only" (false) or "All levels" (true). */
  allLevels: boolean;
  /** The sheet's extent: whole-column references stop at its edge. */
  rowCount: number;
  columnCount: number;
  /**
   * XLS-A062: the same-sheet areas a formula reaches through defined names
   * and table (structured) references, which the A1 grammar does not see.
   * `row`/`col` is the formula's cell (a this-row reference reads it).
   */
  indirectRanges?: (formula: string, row: number, col: number) => GoToSpecialRange[];
}

/** Every area a formula at (row, col) refers to: A1, names and tables. */
function referenceAreas(
  input: GoToSpecialReferenceInput,
  formula: string,
  key: string,
): GoToSpecialRange[] {
  const direct = formulaReferenceRanges(formula);
  if (!input.indirectRanges) return direct;
  const [row, col] = key.split(":").map(Number);
  return [...direct, ...input.indirectRanges(formula, row, col)];
}

function sortedKeys(keys: Iterable<string>): string[] {
  return [...keys]
    .map((key) => {
      const [row, col] = key.split(":").map(Number);
      return { key, row, col };
    })
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .map((entry) => entry.key);
}

/**
 * Go To Special > Precedents (XLS-A062): every cell the formulas in `range`
 * refer to, blank ones included, as Excel selects them. With All levels the
 * formulas among those cells are followed in turn, breadth first, and a
 * cycle is visited once. Same sheet only, as in Excel.
 */
export function findPrecedentCells(input: GoToSpecialReferenceInput): string[] {
  const { cells, range } = input;
  const maxRow = Math.max(0, input.rowCount - 1);
  const maxCol = Math.max(0, input.columnCount - 1);
  const found = new Set<string>();
  const visited = new Set<string>();
  let frontier: string[] = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    for (let col = range.c1; col <= range.c2; col += 1) {
      const key = `${row}:${col}`;
      if (formulaText(cells[key]) !== null) frontier.push(key);
    }
  }
  while (frontier.length > 0 && found.size < MAX_SPECIAL_CELLS) {
    const next: string[] = [];
    for (const key of frontier) {
      if (visited.has(key)) continue;
      visited.add(key);
      const formula = formulaText(cells[key]);
      if (formula === null) continue;
      for (const area of referenceAreas(input, formula, key)) {
        const r2 = Math.min(area.r2, maxRow);
        const c2 = Math.min(area.c2, maxCol);
        for (let row = area.r1; row <= r2; row += 1) {
          for (let col = area.c1; col <= c2; col += 1) {
            const target = `${row}:${col}`;
            if (found.has(target)) continue;
            found.add(target);
            if (found.size >= MAX_SPECIAL_CELLS) return sortedKeys(found);
            if (input.allLevels && formulaText(cells[target]) !== null) {
              next.push(target);
            }
          }
        }
      }
    }
    frontier = next;
  }
  return sortedKeys(found);
}

/**
 * Go To Special > Dependents (XLS-A062): every formula cell on the sheet that
 * refers to a cell of `range`. With All levels, the formulas referring to
 * those dependents are added in turn until nothing new turns up.
 */
export function findDependentCells(input: GoToSpecialReferenceInput): string[] {
  const entries: Array<{ key: string; areas: GoToSpecialRange[] }> = [];
  for (const key in input.cells) {
    const formula = formulaText(input.cells[key]);
    if (formula === null) continue;
    entries.push({ key, areas: referenceAreas(input, formula, key) });
  }
  const found = new Set<string>();
  // Level one reads the selection as a rectangle; deeper levels read the
  // dependents just found, one cell each.
  let targets: GoToSpecialRange[] = [input.range];
  while (targets.length > 0 && found.size < MAX_SPECIAL_CELLS) {
    const next: GoToSpecialRange[] = [];
    for (const entry of entries) {
      if (found.has(entry.key)) continue;
      const hit = entry.areas.some((area) =>
        targets.some(
          (target) =>
            area.r1 <= target.r2 &&
            area.r2 >= target.r1 &&
            area.c1 <= target.c2 &&
            area.c2 >= target.c1,
        ),
      );
      if (!hit) continue;
      found.add(entry.key);
      if (found.size >= MAX_SPECIAL_CELLS) break;
      const [row, col] = entry.key.split(":").map(Number);
      next.push({ r1: row, c1: col, r2: row, c2: col });
    }
    targets = input.allLevels ? next : [];
  }
  return sortedKeys(found);
}

/**
 * The next marked cell after `from`, for Enter and Tab stepping through a
 * match set the grid cannot hold as one selection. Excel moves through a
 * selection the same way: Enter goes down a column and on to the next one,
 * Tab goes along a row and on to the next one, Shift reverses, and the end
 * wraps to the start. Returns null when `from` is not one of the marks.
 */
export function stepSpecialMark(
  keys: readonly string[],
  from: { row: number; col: number },
  key: "Enter" | "Tab",
  reverse: boolean,
): { row: number; col: number } | null {
  const cells = keys.map((entry) => {
    const [rowText, colText] = entry.split(":");
    return { row: Number(rowText), col: Number(colText) };
  });
  const order =
    key === "Enter"
      ? (a: { row: number; col: number }, b: { row: number; col: number }) =>
          a.col - b.col || a.row - b.row
      : (a: { row: number; col: number }, b: { row: number; col: number }) =>
          a.row - b.row || a.col - b.col;
  cells.sort(order);
  const index = cells.findIndex(
    (cell) => cell.row === from.row && cell.col === from.col,
  );
  if (index < 0 || cells.length === 0) return null;
  const next = (index + (reverse ? -1 : 1) + cells.length) % cells.length;
  return cells[next];
}

/**
 * Excel's Current region: the block of data around the active cell, bounded
 * by blank rows and columns. From a blank cell that touches data (the row
 * under a list, say) the region still takes in that data; an isolated blank
 * cell is its own region.
 */
export function currentRegionAt(
  cells: Parameters<typeof surroundingDataRegion>[0],
  row: number,
  col: number,
  rowCount: number,
  columnCount: number,
  merges: Parameters<typeof surroundingDataRegion>[6] = [],
  /** Past this the region is not read: what runs on every selection names it. */
  maxCells?: number,
): DataRegionRange {
  const own = surroundingDataRegion(
    cells,
    row,
    col,
    rowCount,
    columnCount,
    maxCells,
    merges,
  );
  if (own) return own;
  let region: DataRegionRange = { r1: row, c1: col, r2: row, c2: col };
  for (let r = Math.max(0, row - 1); r <= Math.min(rowCount - 1, row + 1); r += 1) {
    for (
      let c = Math.max(0, col - 1);
      c <= Math.min(columnCount - 1, col + 1);
      c += 1
    ) {
      if (r === row && c === col) continue;
      const touching = surroundingDataRegion(
        cells,
        r,
        c,
        rowCount,
        columnCount,
        maxCells,
        merges,
      );
      if (!touching) continue;
      region = {
        r1: Math.min(region.r1, touching.r1),
        c1: Math.min(region.c1, touching.c1),
        r2: Math.max(region.r2, touching.r2),
        c2: Math.max(region.c2, touching.c2),
      };
    }
  }
  return region;
}

/**
 * The list a list command (Sort, Filter, Remove Duplicates, Subtotal, Create
 * Table, Forecast...) takes from one selected cell: Excel's CurrentRegion,
 * counted from a blank cell too -- the blank A9 beside B9 sorts $A$8:$B$14,
 * measured in Excel 16 for Mac (xlsx-rounds evidence/xprobe6). Null for an
 * isolated blank, whose region is itself and which Excel refuses.
 */
export function listRegionAt(
  cells: Parameters<typeof surroundingDataRegion>[0],
  row: number,
  col: number,
  rowCount: number,
  columnCount: number,
  merges: Parameters<typeof surroundingDataRegion>[6] = [],
  maxCells?: number,
): DataRegionRange | null {
  const region = currentRegionAt(cells, row, col, rowCount, columnCount, merges, maxCells);
  return region.r1 === region.r2 && region.c1 === region.c2 &&
    !surroundingDataRegion(cells, row, col, rowCount, columnCount, maxCells, merges)
    ? null
    : region;
}

export type ListRegionMemo = {
  cells: object; r1: number; c1: number; rowCount: number; columnCount: number; merges: unknown; maxCells: number | undefined;
  region: DataRegionRange | null;
};

/**
 * {@link listRegionAt}, kept in `slot`: the same answer for the same cell
 * while the cells' data stays where it was. Their lineage names what changed
 * (cell-map-lineage), and a region reads of a cell only whether it holds
 * data, so a new formula result or format moves none. Every list reader of
 * the Data tab re-read a sorted 728,000-cell list each time its save stamped
 * the formulas' results, 360 ms a time.
 */
export function listRegionAtReusing(
  slot: { current: ListRegionMemo | null },
  cells: Parameters<typeof surroundingDataRegion>[0],
  row: number,
  col: number,
  rowCount: number,
  columnCount: number,
  merges?: Parameters<typeof surroundingDataRegion>[6],
  maxCells?: number,
): DataRegionRange | null {
  const previous = slot.current;
  if (previous && previous.r1 === row && previous.c1 === col && previous.rowCount === rowCount &&
      previous.columnCount === columnCount && previous.merges === merges && previous.maxCells === maxCells) {
    if (previous.cells === cells) return previous.region;
    const before = previous.cells as typeof cells;
    const changed = cellMapChangesSince(cells, before);
    if (changed && changed.every((key) => regionCellHoldsData(before[key]) === regionCellHoldsData(cells[key]))) {
      slot.current = { ...previous, cells };
      return previous.region;
    }
  }
  const region = listRegionAt(cells, row, col, rowCount, columnCount, merges ?? [], maxCells);
  slot.current = { cells, r1: row, c1: col, rowCount, columnCount, merges, maxCells, region };
  return region;
}

/**
 * Excel's Last cell: the last row and the last column that hold data or
 * formatting, taken independently, so the corner itself may be blank. An
 * empty sheet's last cell is A1.
 */
export function lastCellOf(
  cells: Readonly<Record<string, GoToSpecialCell | undefined>>,
  rowCount: number,
  columnCount: number,
): { row: number; col: number } {
  let lastRow = 0;
  let lastCol = 0;
  for (const key in cells) {
    const cell = cells[key];
    if (!cell) continue;
    const used =
      text(cell) !== "" ||
      isFormula(cell) ||
      hasComment(cell) ||
      (cell.style !== undefined && Object.keys(cell.style).length > 0);
    if (!used) continue;
    const separator = key.indexOf(":");
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row < 0 ||
      col < 0 ||
      row >= rowCount ||
      col >= columnCount
    ) {
      continue;
    }
    if (row > lastRow) lastRow = row;
    if (col > lastCol) lastCol = col;
  }
  return { row: lastRow, col: lastCol };
}

/**
 * Data validation "Same": two rules are the same validation when every
 * setting but the range they cover agrees, references read from each rule's
 * own cells (sameValidationSettingsAs, shared with the dialog's "Apply these
 * changes to all other cells with the same settings"). An imported rule that
 * covered several areas arrives as one rule per area, so identity alone would
 * miss its other areas.
 */
export function sameValidationSettings(
  left: { range: unknown },
  right: { range: unknown },
): boolean {
  return sameValidationSettingsAs(left as ValidationRule)(right as ValidationRule);
}
