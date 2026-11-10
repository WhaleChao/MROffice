/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { matchingChartValueCache } from "./chart-value-cache";
import { worksheetFormulaResultKey } from "./calculation-mode";
import { shouldCalculateSpreadsheetInBackground, type BackgroundCalculationHostBinding } from "./background-calculation-host";
import { readLiteralCellValue, type EngineValue, type WorkbookSheetInput } from "./formula-engine";
import type { WorksheetFormulaCoordinate } from "./calculation-mode";
import { evaluateArithmeticFormula, type ArithmeticValue } from "./arithmetic-evaluator";
import { isWhatIfTableFormula } from "./what-if-table-definition";
import { cellMapRowPermutation, nearestKnownCellMap, permutedCellKeys, permutedSpanOf } from "./cell-map-lineage";
import { coveredCells, normalizeMerges } from "./merge-cells";

const ERROR_CODES = new Set([
  "#DIV/0!", "#VALUE!", "#NAME?", "#REF!", "#N/A", "#NUM!", "#SPILL!", "#CALC!", "#NULL!", "#BUSY!", "#CONNECT!", "#BLOCKED!",
]);
const NUMBER_LIKE = /^[\s(+\-−]*[$€£¥¥₩]?\s*[\d.,]/u;

const cache = new WeakMap<readonly BackgroundCalculationHostBinding[], ReadonlyMap<string, EngineValue>>();

export function storedFormulaResult(raw: string | null, cell: { display?: unknown; chartValueCache?: unknown; value?: unknown }): EngineValue | undefined {
  if (!raw || !raw.trimStart().startsWith("=")) return undefined;
  // A What-If Data Table's result keeps its last value as the cell's own
  // (what-if-table-definition): that is the result until the session answers.
  if (isWhatIfTableFormula(raw)) {
    return typeof cell.value === "string" && cell.value !== "" ? readLiteralCellValue(cell.value) : undefined;
  }
  if (matchingChartValueCache(cell.chartValueCache, raw)) return cell.chartValueCache.value;
  const display = cell.display;
  if (typeof display !== "string") return undefined;
  if (display === "TRUE") return true;
  if (display === "FALSE") return false;
  if (ERROR_CODES.has(display)) return { __formulaError: true, code: display };
  if (NUMBER_LIKE.test(display)) return undefined;
  return display;
}

/** Whether stored text can spell a formula: `=` after any leading space. */
function mayBeFormula(value: unknown): boolean {
  if (typeof value !== "string" || value === "") return false;
  const first = value.charCodeAt(0);
  return first === 61 || first <= 32 || first === 160 || first > 127;
}

/** Bound on the cells a pending frame scans for arithmetic hints. */
const MAX_HINT_CELLS = 50_000;

const isFormulaError = (value: unknown): value is { __formulaError: true; code: string } =>
  typeof value === "object" && value !== null && (value as { __formulaError?: unknown }).__formulaError === true;

/** A literal cell as the engine reads it (a typed "50%" is 0.5, "TRUE" a logical). */
function literal(raw: string | null): ArithmeticValue {
  if (raw === null || raw === "") return null;
  const value = readLiteralCellValue(raw);
  return typeof value === "number" || typeof value === "boolean" || typeof value === "string" || isFormulaError(value) ? value : raw;
}

/**
 * `known` plus a provisional value for every edited formula the bounded
 * arithmetic evaluator can read (`=A1+B1` typed a moment ago, also when it
 * reads another formula edited with it): what the pending frame shows until
 * the calculation session answers (C2). Formulas not edited keep their value
 * in `known` (their rows need not redraw while pending), and so does one the
 * evaluator cannot read (a function call, another sheet, a cycle). `known`
 * itself when there is nothing to add (no copy).
 */
export function withArithmeticHints(
  bindings: readonly BackgroundCalculationHostBinding[],
  known: ReadonlyMap<string, EngineValue>,
  /** The cell keys of binding `index` edited since `known` (the dirty
   * journal's), or null when unknown: a render never scans a whole sheet. */
  editedKeys: (index: number) => readonly string[] | null,
): ReadonlyMap<string, EngineValue> {
  const edited = bindings.map((_, index) => editedKeys(index));
  if (edited.reduce((total, keys) => total + (keys?.length ?? 0), 0) > MAX_HINT_CELLS) return known;
  let hints: Map<string, EngineValue> | null = null;
  for (const [index, binding] of bindings.entries()) {
    const keys = edited[index];
    // The evaluator reads its own sheet only: no edit here, nothing reached.
    if (!keys?.length) continue;
    const name = binding.input.name;
    const editedSet = new Set(keys);
    const candidates = editedSet;
    type Read = { value: ArithmeticValue | undefined; reached: boolean };
    const done = new Map<string, Read>();
    const visiting = new Set<string>();
    const knownValue = (row: number, col: number): ArithmeticValue | undefined => {
      const value = known.get(worksheetFormulaResultKey(name, row, col));
      return value === undefined || (typeof value === "object" && value !== null && !("__formulaError" in value))
        ? undefined : value as ArithmeticValue;
    };
    const cell = (row: number, col: number): Read => {
      const local = `${row}:${col}`;
      const raw = binding.input.getRawCell(row, col);
      if (!raw || !raw.trimStart().startsWith("=")) return { value: literal(raw), reached: editedSet.has(local) };
      const memo = done.get(local);
      if (memo) return memo;
      if (!candidates.has(local) || visiting.has(local)) return { value: knownValue(row, col), reached: false };
      visiting.add(local);
      let reached = editedSet.has(local);
      const value = evaluateArithmeticFormula(raw, (r, c) => {
        const read = cell(r, c);
        if (read.reached) reached = true;
        return read.value;
      });
      visiting.delete(local);
      const read: Read = reached ? { value, reached } : { value: knownValue(row, col), reached };
      done.set(local, read);
      return read;
    };
    for (const key of candidates) {
      const [row, col] = key.split(":").map(Number);
      if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col)) continue;
      const read = cell(row, col);
      if (!read.reached || read.value === undefined) continue;
      const raw = binding.input.getRawCell(row, col);
      if (!raw || !raw.trimStart().startsWith("=")) continue;
      (hints ??= new Map()).set(worksheetFormulaResultKey(name, row, col), read.value);
    }
  }
  if (!hints) return known;
  return new Map([...known, ...hints]);
}

/**
 * One sheet's stamped results and unstamped formula keys, for one reading of
 * its raw cells. A sheet's raw cells read as its stored values, masked by its
 * merges; its What-If Data Tables read too, and those (and their bounds) keep
 * a sheet from being carried across an edit at all.
 */
interface SheetStored {
  readonly name: string;
  readonly merges: unknown;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly results: ReadonlyMap<string, EngineValue>;
  /** Formula keys with no stamp, in the order they were found. */
  readonly unstamped: ReadonlySet<string>;
}
const sheetStored = new WeakMap<object, SheetStored[]>();
type StoredBinding = Pick<BackgroundCalculationHostBinding, "sheet" | "input">;
function storedEntryFor(binding: StoredBinding): SheetStored {
  const sheet = binding.sheet as StoredBinding["sheet"] & { readonly merges?: unknown; readonly whatIfTables?: unknown };
  const cells = sheet.cells;
  const name = binding.input.name;
  const matches = (entry: SheetStored) => entry.name === name && entry.merges === sheet.merges &&
    entry.rowCount === sheet.rowCount && entry.columnCount === sheet.columnCount;
  const own = sheetStored.get(cells)?.find(matches);
  if (own) return own;
  // Only a formula holds a result, and a What-If Data Table's reader is the
  // one that answers a formula for a cell not spelled as one. Every other
  // constant leaves here on its own text, not through the key and the reader.
  const everyCellRead = Boolean(sheet.whatIfTables);
  const read = (key: string, cell: unknown, results: Map<string, EngineValue>, unstamped: Set<string>) => {
    if (!everyCellRead && !mayBeFormula((cell as { value?: unknown } | null)?.value)) return;
    const [row, col] = key.split(":").map(Number);
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col)) return;
    const raw = binding.input.getRawCell(row, col);
    const value = storedFormulaResult(typeof raw === "string" ? raw : null, cell as { display?: unknown; chartValueCache?: unknown; value?: unknown });
    if (value !== undefined) results.set(worksheetFormulaResultKey(name, row, col), value);
    else if (raw && raw.trimStart().startsWith("=")) unstamped.add(key);
  };
  // A row sort of a map read here: the results go with their rows, but those
  // of the formulas it rewrote (they hold none of their own now): the frame
  // after a 730,000-cell sort read every cell again, 0.46 s.
  const sorted = sheet.whatIfTables ? null : cellMapRowPermutation(cells);
  const sortedFrom = sorted?.before ? sheetStored.get(sorted.before)?.find(matches) : undefined;
  // An edit changes the cells it names and nothing else a stored result is
  // read from (What-If Data Tables read other cells: those sheets walk).
  const prior = sheet.whatIfTables || sortedFrom ? null : nearestKnownCellMap(cells, (before) => sheetStored.get(before)?.find(matches));
  let entry: SheetStored;
  if (sorted && sortedFrom) {
    const span = permutedSpanOf(sorted.permutation);
    const { dataR1, c1, c2 } = sorted.permutation;
    const prefix = worksheetFormulaResultKey(name, 0, 0).slice(0, -"0:0".length);
    const results = new Map<string, EngineValue>();
    const unstamped = new Set<string>();
    const rewritten = new Set(sorted.permutation.relocated);
    for (const [resultKey, value] of sortedFrom.results) {
      const key = resultKey.slice(prefix.length);
      const at = key.indexOf(":");
      const row = Number(key.slice(0, at)), col = Number(key.slice(at + 1));
      if (row < dataR1 || row > span.r2 || col < c1 || col > c2) {
        results.set(resultKey, value);
        continue;
      }
      const movedKey = `${span.destination[row - dataR1]}:${col}`;
      if (!rewritten.has(movedKey)) results.set(prefix + movedKey, value);
    }
    for (const key of permutedCellKeys([...sortedFrom.unstamped], span, cells, () => false)) unstamped.add(key);
    // Asked again: the formulas the sort rewrote.
    for (const key of rewritten) read(key, cells[key], results, unstamped);
    entry = { name, merges: sheet.merges, rowCount: sheet.rowCount, columnCount: sheet.columnCount, results, unstamped };
  } else if (prior) {
    const results = new Map(prior.value.results);
    const unstamped = new Set(prior.value.unstamped);
    for (const key of prior.keys) {
      const [row, col] = key.split(":").map(Number);
      if (Number.isSafeInteger(row) && Number.isSafeInteger(col)) results.delete(worksheetFormulaResultKey(name, row, col));
      unstamped.delete(key);
      if (Object.prototype.hasOwnProperty.call(cells, key)) read(key, cells[key], results, unstamped);
    }
    entry = { name, merges: sheet.merges, rowCount: sheet.rowCount, columnCount: sheet.columnCount, results, unstamped };
  } else {
    const results = new Map<string, EngineValue>();
    const unstamped = new Set<string>();
    // Walked in place: Object.entries made a pair for each of a sorted
    
    for (const key in cells) {
      if (Object.prototype.hasOwnProperty.call(cells, key)) read(key, cells[key], results, unstamped);
    }
    entry = { name, merges: sheet.merges, rowCount: sheet.rowCount, columnCount: sheet.columnCount, results, unstamped };
  }
  const entries = sheetStored.get(cells) ?? [];
  sheetStored.set(cells, [entry, ...entries].slice(0, 2));
  return entry;
}

/** Every stamped formula result of the bindings' sheets, by result key. */
export function storedFormulaResults(bindings: readonly BackgroundCalculationHostBinding[]): ReadonlyMap<string, EngineValue> {
  const known = cache.get(bindings);
  if (known) return known;
  // Each sheet's share is kept with its cells: an edit reads its own sheet
  // again (from the map before it, at the keys it changed), never the others.
  const shares = bindings.map((binding) => storedEntryFor(binding).results);
  const results = shares.length === 1 ? new Map(shares[0]) : new Map<string, EngineValue>();
  if (shares.length !== 1) for (const share of shares) for (const [key, value] of share) results.set(key, value);
  cache.set(bindings, results);
  return results;
}

/** Bound on the unstamped formulas a first paint evaluates arithmetically. */
const MAX_FIRST_PAINT_HINTS = 2_048;
const firstPaintCache = new WeakMap<readonly BackgroundCalculationHostBinding[], ReadonlyMap<string, EngineValue>>();

/**
 * What the grid shows on open before the session answers: the stamped
 * results ({@link storedFormulaResults}) plus, for a formula carrying no
 * stamp (a workbook never saved by MRAI, a number shown only as formatted
 * text), the bounded arithmetic evaluator's value -- so `=A1*2` paints at
 * once, as it did when the browser evaluated small arithmetic sheets itself.
 * Past MAX_FIRST_PAINT_HINTS unstamped formulas, only the stamps.
 */
export function firstPaintFormulaResults(bindings: readonly BackgroundCalculationHostBinding[]): ReadonlyMap<string, EngineValue> {
  const known = firstPaintCache.get(bindings);
  if (known) return known;
  const stored = storedFormulaResults(bindings);
  let total = 0;
  const unstamped = bindings.map((binding) => {
    const keys: string[] = [];
    if (total > MAX_FIRST_PAINT_HINTS) return keys;
    // A formula with no stamp of its own sheet (another sheet of the same
    // name never shares a result key with it: names are unique).
    for (const key of storedEntryFor(binding).unstamped) {
      const [row, col] = key.split(":").map(Number);
      if (stored.has(worksheetFormulaResultKey(binding.input.name, row, col))) continue;
      keys.push(key);
      if (++total > MAX_FIRST_PAINT_HINTS) break;
    }
    return keys;
  });
  const results = total > MAX_FIRST_PAINT_HINTS ? stored : withArithmeticHints(bindings, stored, (index) => unstamped[index]);
  firstPaintCache.set(bindings, results);
  return results;
}

/** Excel's grid: a reference past it is #REF! in the engine; decline it here. */
const GRID_ROWS = 1_048_576;
const GRID_COLUMNS = 16_384;

/**
 * Every requested formula's value from the bounded arithmetic evaluator, or
 * null when any one needs the server. This is the authority for the small
 * arithmetic workbooks the browser used to evaluate itself (fewer than 256
 * formulas, all of them arithmetic over same-sheet references: the old
 * main-thread bypass): over numbers, logicals, blanks and errors the evaluator
 * answers exactly what the engine answers (randomized parity,
 * __tests__/arithmetic-evaluator.test.ts), so such a workbook settles at once
 * without a round trip. Anything else declines: a text operand (Excel's text
 * to number coercion -- dates, times, "1,000" -- is the server's), a text or
 * picture cell, a cycle, a reference past the grid.
 */
export function localArithmeticResults(
  bindings: readonly BackgroundCalculationHostBinding[],
  coordinates: readonly WorksheetFormulaCoordinate[],
): ReadonlyMap<string, EngineValue> | null {
  // A legacy array formula (Ctrl+Shift+Enter, `arrayRef`) lays its result
  // over its whole range, and a formula reading a cell of that range reads
  // the element shown there: the engine's (formula-evaluator
  // fitLegacyArrayResult), never this scalar walk. `{=1+2}` over A1:B2 shows
  // 3 in all four cells (K16, X-009).
  if (bindings.some((binding) => binding.input.options?.legacyArrayExtent)) return null;
  const byId = new Map(bindings.map((binding) => [binding.sheet.id, binding]));
  const results = new Map<string, EngineValue>();
  const visiting = new Set<string>();
  function evaluate(binding: BackgroundCalculationHostBinding, row: number, col: number, raw: string): ArithmeticValue | undefined {
    const key = worksheetFormulaResultKey(binding.input.name, row, col);
    if (results.has(key)) return results.get(key) as ArithmeticValue;
    if (visiting.has(key)) return undefined;
    visiting.add(key);
    const value = evaluateArithmeticFormula(raw, (r, c) => read(binding, r, c));
    visiting.delete(key);
    if (value !== undefined) results.set(key, value);
    return value;
  }
  function read(binding: BackgroundCalculationHostBinding, row: number, col: number): ArithmeticValue | undefined {
    if (row >= GRID_ROWS || col >= GRID_COLUMNS) return undefined;
    const raw = binding.input.getRawCell(row, col);
    if (raw === null || raw === "") return null;
    const options = binding.input.options;
    if (options?.isTextCell?.(row, col) || options?.cellPicture?.(row, col)) return undefined;
    const value = raw.trimStart().startsWith("=") ? evaluate(binding, row, col, raw) : literal(raw);
    return typeof value === "string" ? undefined : value;
  }
  for (const cell of coordinates) {
    const binding = byId.get(cell.sheetId);
    const raw = binding?.input.getRawCell(cell.row, cell.col);
    if (!binding || !raw || binding.input.options?.isTextCell?.(cell.row, cell.col)) return null;
    if (evaluate(binding, cell.row, cell.col, raw) === undefined) return null;
  }
  return results;
}

/**
 * Whether the sheets hold `limit` cells or more whose own text is a formula,
 * outside merged areas. The walk over the engine inputs below counts every
 * such cell (an input reads a cell's own text, and a What-If table only adds
 * formulas), so reaching `limit` here means it would: the inputs need not be
 * built to learn it. Built for a sorted 730,000-cell workbook (hidden rows,
 * text cells, pictures, ...) they were a second of its save, which ended in
 * "2,122 formulas: not here".
 */
/** The keys of a map's cells whose text reads as a formula, by the map. */
const formulaCellsOfMap = new WeakMap<object, readonly string[]>();

/**
 * Which of a map's cells read as formulas. Each save walked every cell of the
 * workbook to find out (30,000 cells: about 20 ms per entry typed, on the main
 * thread): from the nearest map this is known for, only the keys changed since
 * are looked at again.
 */
function cellsReadingAsFormulas(
  cells: Readonly<Record<string, unknown>>,
  input: { getRawCell(row: number, col: number): string | null | undefined },
): readonly string[] {
  const kept = formulaCellsOfMap.get(cells);
  if (kept) return kept;
  const reads = (key: string): boolean => {
    const separator = key.indexOf(":");
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col)) return false;
    const raw = input.getRawCell(row, col);
    return Boolean(raw && raw.trimStart().startsWith("="));
  };
  const near = nearestKnownCellMap(cells as never, (map) => formulaCellsOfMap.get(map as object));
  let keys: readonly string[];
  if (near) {
    const changed = new Set(near.keys);
    keys = [...near.value.filter((key) => !changed.has(key)), ...near.keys.filter((key) => key in cells && reads(key))];
  } else {
    keys = Object.keys(cells).filter(reads);
  }
  formulaCellsOfMap.set(cells, keys);
  return keys;
}

function formulaTextsReach(
  sheets: readonly { cells: Readonly<Record<string, unknown>>; merges?: unknown }[],
  limit: number,
  formulaKeysOf?: (cells: Readonly<Record<string, unknown>>) => readonly string[] | undefined,
): boolean {
  let count = 0;
  for (const sheet of sheets) {
    let covered: Set<string> | null | undefined;
    // The window's cell index names every cell whose text begins "=" (and
    // any with a formula field): walked instead of every key, 128 ms of a
    // sorted 728,000-cell sheet's save. It may leave out a formula written
    // after spaces, which only counts fewer: the inputs then decide.
    for (const key of formulaKeysOf?.(sheet.cells) ?? Object.keys(sheet.cells)) {
      const value = (sheet.cells[key] as { value?: unknown } | undefined)?.value;
      if (typeof value !== "string" || !value.trimStart().startsWith("=")) continue;
      const separator = key.indexOf(":");
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (separator < 0 || !Number.isSafeInteger(row) || !Number.isSafeInteger(col) || `${row}:${col}` !== key) continue;
      covered ??= sheet.merges ? coveredCells(normalizeMerges(sheet.merges)) : null;
      if (covered?.has(key)) continue;
      count += 1;
      if (count >= limit) return true;
    }
  }
  return false;
}

/**
 * {@link localArithmeticResults} for a workbook about to be saved (its sheets
 * and their engine inputs, not yet a calculated source): the exact results a
 * save stamps without a round trip, or null when any formula needs the server
 * or there are 256 formulas or more.
 */
export function localArithmeticWorkbookResults(
  sheets: readonly { id: string; name?: unknown; rowCount?: unknown; columnCount?: unknown; merges?: unknown; cells: Readonly<Record<string, unknown>> }[],
  /** The sheets' engine inputs, or how to build one: a workbook already known
   * to hold 256 formulas or more (its stored-result index says so) builds none. */
  givenInputs: readonly WorkbookSheetInput[] | ((index: number) => WorkbookSheetInput),
  /** The cells a sheet's cell index names as formulas (a superset of those
   * whose own text begins "="), when the caller keeps one. */
  formulaKeysOf?: (cells: Readonly<Record<string, unknown>>) => readonly string[] | undefined,
): ReadonlyMap<string, EngineValue> | null {
  // The stored-result index of each sheet (kept with its cells, see
  // storedEntryFor) counts its formulas exactly as the walk below does.
  let known = 0;
  for (const sheet of sheets) {
    const name = typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id;
    const entry = sheetStored.get(sheet.cells)?.find((candidate) => candidate.name === name && candidate.merges === sheet.merges &&
      candidate.rowCount === sheet.rowCount && candidate.columnCount === sheet.columnCount);
    if (!entry) { known = -1; break; }
    known += entry.results.size + entry.unstamped.size;
  }
  if (known >= 256) return null;
  if (known < 0 && typeof givenInputs === "function" && formulaTextsReach(sheets, 256, formulaKeysOf)) return null;
  const inputs = typeof givenInputs === "function" ? sheets.map((_, index) => givenInputs(index)) : givenInputs;
  const bindings = sheets.map((sheet, index) => ({ sheet, input: inputs[index] })) as unknown as BackgroundCalculationHostBinding[];
  const coordinates: WorksheetFormulaCoordinate[] = [];
  for (const binding of bindings) {
    for (const key of cellsReadingAsFormulas(binding.sheet.cells, binding.input)) {
      const separator = key.indexOf(":");
      coordinates.push({ sheetId: binding.sheet.id, sheetName: binding.input.name, row: Number(key.slice(0, separator)), col: Number(key.slice(separator + 1)) });
      if (coordinates.length >= 256) return null;
    }
  }
  if (coordinates.length === 0) return new Map();
  return shouldCalculateSpreadsheetInBackground(bindings, coordinates) ? null : localArithmeticResults(bindings, coordinates);
}
