/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The file's own last value of a formula MRAI cannot evaluate: one that reads
 * external data -- an OLAP cube (CUBEMEMBER, CUBEVALUE, ...), a real-time data
 * server (RTD) or another workbook (`[1]Sheet1!A1`, `'[Book.xlsx]Sheet'!A1`).
 *
 * Excel keeps showing the value it last retrieved for such a formula until the
 * data is refreshed: it does not recalculate cube functions or external links
 * on open, and a workbook opened without its connection shows the cached
 * results (O19_Pivot3, K08_CubeFromRange-Input; X-079). The importer stamps
 * that value on the cell (`keptResult`, with the formula it belongs to), and
 * the formula engine answers it for the cell while the cell still holds that
 * formula -- the grid, every formula reading the cell, a save's stamps and an
 * export all see it. Typing another formula retires it.
 *
 * A formula calling a function this engine does not have -- a VBA user
 * defined function, an add-in's -- is held the same way, with what it read
 * (`inputs`: each area its references name and the values the file stores in
 * it). Excel opens such a workbook showing the cached results, and with its
 * macros off (VBA never runs here) a cell it recalculates is #NAME?; it
 * recalculates a cell when a cell it reads changes. The engine answers the
 * held value while every input still holds the file's value, and calculates
 * the cell (#NAME?) once one does not (`keptInputsHold`).
 *
 * The API importer keeps the same recogniser
 * (`app/services/file_adapters/xlsx/kept_formula_results.py`).
 */
import { matchingChartValueCache } from "./chart-value-cache";
import { coerceLiteral, ERRORS, isArrayValue, isFormulaError, isKnownFunction, type CellValue, type EngineValue } from "./formula-engine";

export type KeptFormulaResultKind = "number" | "text" | "boolean" | "error";

/** One area a held formula reads: zero-based [top, left, bottom, right] on
 *  `sheet`, and the values the file stores in it ("row:col"; a blank left out). */
export type KeptFormulaInput = Readonly<{
  sheet: string;
  area: readonly [number, number, number, number];
  cells: Readonly<Record<string, readonly [KeptFormulaResultKind, number | string | boolean]>>;
}>;

/** A cell's `keptResult`: the file's cached `<v>` of the formula it names. */
export type KeptFormulaResult = Readonly<{
  /** The cell's formula ("=..."), exactly as its `value` holds it. */
  formula: string;
  kind: KeptFormulaResultKind;
  value: number | string | boolean;
  /** A held result's inputs (kept_formula_results.py `held_result`). */
  inputs?: readonly KeptFormulaInput[];
}>;

/** The most cells a held result's inputs span (the importer's MAX_INPUT_CELLS). */
export const MAX_KEPT_INPUT_CELLS = 4096;
const MAX_KEPT_INPUT_AREAS = 64;

const STRING_LITERAL = /"(?:[^"]|"")*"/g;
// `(?:^|[^…])` instead of a lookbehind (Safari < 16.4 cannot compile one);
// both are only tested, so consuming the character before is the same.
const EXTERNAL_DATA_FUNCTION =
  /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.)?(?:CUBEKPIMEMBER|CUBEMEMBERPROPERTY|CUBEMEMBER|CUBERANKEDMEMBER|CUBESETCOUNT|CUBESET|CUBEVALUE|RTD)\s*\(/i;
// [1]Sheet1!A1, [1]!Name, '[1]Sheet 1'!A1, [Book.xlsx]Sheet!A1,
// '[Book.xlsx]Sheet'!A1 -- never a table's [Column] or [@Column], nor
// [0]!Name (this workbook's own name, P16_56737).
const EXTERNAL_WORKBOOK_REFERENCE =
  /(?:^|[^A-Za-z0-9_.\]])\[[1-9]\d*\][^[\]!(),]*!|'\[[^\]]+\][^']*'!|(?:^|[^A-Za-z0-9_.\]])\[[^[\]]+\.xl[a-z]{1,2}\][^[\]!(),]*!/i;

/** Whether a formula reads data MRAI cannot reach (cube, RTD, another workbook). */
export function formulaReadsExternalData(formula: string): boolean {
  if (typeof formula !== "string" || !formula.trimStart().startsWith("=")) return false;
  const code = formula.replace(STRING_LITERAL, '""');
  return EXTERNAL_DATA_FUNCTION.test(code) || EXTERNAL_WORKBOOK_REFERENCE.test(code);
}

const BINDS_NAMES = /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i;
const QUOTED_SHEET = /'(?:[^']|'')*'/g;
const BRACKET_TEXT = /\[[^[\]]*\]/g;

/**
 * Whether a formula may hold the file's value as a held result: it calls a
 * function this engine does not have (never IMAGE, which this engine draws),
 * binds no names of its own (LET, LAMBDA) and reads no external data. The
 * importer judges it with the workbook's defined names too (a LAMBDA name is
 * a function); this is the check that the cell still holds such a formula.
 */
export function formulaHoldsCachedResult(formula: string): boolean {
  if (typeof formula !== "string" || !formula.trimStart().startsWith("=") || formulaReadsExternalData(formula)) return false;
  const code = formula.replace(STRING_LITERAL, '""').replace(QUOTED_SHEET, "''").replace(BRACKET_TEXT, "[]");
  if (BINDS_NAMES.test(code)) return false;
  let unknown = false;
  for (const match of code.matchAll(/(^|[^\w.!$'\]])([A-Za-z_][\w.]*)\s*\(/g)) {
    const name = match[2]!.replace(/^(?:_xlfn\.|_xlws\.)+/i, "").toUpperCase();
    if (name === "IMAGE") return false;
    if (!isKnownFunction(name)) unknown = true;
  }
  return unknown;
}

function validKeptValue(kind: unknown, value: unknown): boolean {
  if (kind === "number") return typeof value === "number" && Number.isFinite(value);
  if (kind === "boolean") return typeof value === "boolean";
  return (kind === "text" || kind === "error") && typeof value === "string";
}

/** A held result's `inputs` as stored, or null when malformed or too large. */
function readKeptInputs(raw: unknown): KeptFormulaInput[] | null {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_KEPT_INPUT_AREAS) return null;
  const out: KeptFormulaInput[] = [];
  let span = 0;
  for (const entry of raw) {
    if (typeof entry !== "object" || entry === null) return null;
    const { sheet, area, cells } = entry as Record<string, unknown>;
    if (typeof sheet !== "string" || !sheet || !Array.isArray(area) || area.length !== 4) return null;
    if (!area.every((n) => Number.isSafeInteger(n) && (n as number) >= 0)) return null;
    const [top, left, bottom, right] = area as number[];
    if (top! > bottom! || left! > right!) return null;
    span += (bottom! - top! + 1) * (right! - left! + 1);
    if (span > MAX_KEPT_INPUT_CELLS || typeof cells !== "object" || cells === null || Array.isArray(cells)) return null;
    const read: Record<string, readonly [KeptFormulaResultKind, number | string | boolean]> = {};
    for (const [key, value] of Object.entries(cells as Record<string, unknown>)) {
      const match = /^(\d+):(\d+)$/.exec(key);
      if (!match || !Array.isArray(value) || value.length !== 2 || !validKeptValue(value[0], value[1])) return null;
      const row = Number(match[1]);
      const col = Number(match[2]);
      if (row < top! || row > bottom! || col < left! || col > right!) return null;
      read[key] = [value[0] as KeptFormulaResultKind, value[1] as number | string | boolean];
    }
    out.push({ sheet, area: [top!, left!, bottom!, right!], cells: read });
  }
  return out;
}

/** Whether `value`, the engine's value of an input cell now, is the one the
 *  file stored there (`stored`; undefined: a blank). A number calculated here
 *  may differ from Excel's in its last bits. */
export function keptInputHolds(
  value: EngineValue,
  stored: readonly [KeptFormulaResultKind, number | string | boolean] | undefined,
): boolean {
  const scalar = isArrayValue(value) ? (value.grid[0]?.[0] ?? null) : value;
  if (stored === undefined) return scalar === null;
  const [kind, expected] = stored;
  if (kind === "number") {
    if (typeof scalar !== "number") return false;
    const a = scalar;
    const b = expected as number;
    return a === b || Math.abs(a - b) <= 1e-12 * Math.max(Math.abs(a), Math.abs(b));
  }
  if (kind === "error") return isFormulaError(scalar) && scalar.code === String(expected).toUpperCase();
  return scalar === expected;
}

/** A cell's stored `keptResult`, or null when it states none (or a malformed one). */
export function readKeptFormulaResult(raw: unknown): KeptFormulaResult | null {
  if (typeof raw !== "object" || raw === null) return null;
  const { formula, kind, value, inputs } = raw as Record<string, unknown>;
  if (typeof formula !== "string" || !formula.startsWith("=")) return null;
  if (!validKeptValue(kind, value)) return null;
  if (inputs === undefined) return { formula, kind: kind as KeptFormulaResultKind, value: value as number | string | boolean };
  const read = readKeptInputs(inputs);
  return read ? { formula, kind: kind as KeptFormulaResultKind, value: value as number | string | boolean, inputs: read } : null;
}

/** Whether a stored result still belongs to its formula as one this engine
 *  cannot calculate: over external data, or held (`formulaHoldsCachedResult`). */
function keptFormulaApplies(formula: string): boolean {
  return formulaReadsExternalData(formula) || formulaHoldsCachedResult(formula);
}

/** The value the engine answers for a kept result. */
export function keptFormulaResultValue(kept: KeptFormulaResult): CellValue {
  if (kept.kind === "error") {
    const error = coerceLiteral(String(kept.value).toUpperCase());
    return typeof error === "object" && error !== null ? error : ERRORS.NA;
  }
  return kept.value;
}

/** One serialisable entry: "row:col", formula, kind, value as text; "pending"
 *  marks the file's last value of a formula that reads a sheet not filled in
 *  yet; a held result carries its inputs after a null marker. */
export type KeptFormulaResultEntry =
  | readonly [string, string, KeptFormulaResultKind, string]
  | readonly [string, string, KeptFormulaResultKind, string, "pending"]
  | readonly [string, string, KeptFormulaResultKind, string, null, readonly KeptFormulaInput[]];

// A sheet qualifier: 'Quoted name'! or a bare name (or First:Last) before "!".
const SHEET_QUALIFIER = /'((?:[^']|'')+)'!|([^\s'!"(),=+\-*/^&<>:;{}[\]]+(?::[^\s'!"(),=+\-*/^&<>:;{}[\]]+)?)!/g;

/**
 * Whether a formula names one of `sheets` (lowercased names) in a reference:
 * `Name!A1`, `'Na me'!A1`, or either end of a 3-D `First:Last!A1`.
 */
export function formulaReadsSheets(formula: string, sheets: ReadonlySet<string>): boolean {
  if (sheets.size === 0 || typeof formula !== "string" || !formula.includes("!")) return false;
  const code = formula.replace(STRING_LITERAL, '""');
  for (const match of code.matchAll(SHEET_QUALIFIER)) {
    const quoted = match[1] !== undefined;
    const name = (quoted ? match[1].replace(/''/g, "'") : match[2]).toLowerCase();
    // [1]Sheet!A1 and '[Book.xlsx]Sheet'!A1 are another workbook's sheets.
    if (name.startsWith("[") || (!quoted && code[match.index - 1] === "]")) continue;
    if (sheets.has(name)) return true;
    const colon = name.indexOf(":");
    if (colon > 0 && (sheets.has(name.slice(0, colon)) || sheets.has(name.slice(colon + 1)))) return true;
  }
  return false;
}

const KEPT_ERROR_CODES = new Set([
  "#DIV/0!", "#VALUE!", "#NAME?", "#REF!", "#N/A", "#NUM!", "#SPILL!", "#CALC!", "#NULL!",
]);
const NUMBER_LIKE = /^[\s(+\-−]*[$€£¥¥₩]?\s*[\d.,]/u;

/** The result a formula cell's stamps state (`chartValueCache` bound to the
 *  formula, else `display`), as the first paint reads it (stored-formula-results). */
function stampedFormulaResult(
  formula: string,
  cell: { display?: unknown; chartValueCache?: unknown },
): readonly [KeptFormulaResultKind, string] | null {
  if (matchingChartValueCache(cell.chartValueCache, formula)) return ["number", String(cell.chartValueCache.value)];
  const display = cell.display;
  if (typeof display !== "string") return null;
  if (display === "TRUE" || display === "FALSE") return ["boolean", display === "TRUE" ? "true" : "false"];
  if (KEPT_ERROR_CODES.has(display)) return ["error", display];
  if (display.startsWith("#") || NUMBER_LIKE.test(display)) return null;
  return ["text", display];
}

/**
 * The entries a sheet's cells carry, for the background calculation's context.
 *
 * `pending`: the workbook's sheets still placeholders (`sheetPending`,
 * lib/spreadsheet/deferred-sheets; lowercased names) and every cell of this
 * sheet. A formula that names one of them would read an empty sheet (0, "")
 * until the page open fills it in, and a save made meanwhile would stamp that
 * and write it to the file as the formula's value (xlsx-rounds R283-X12): it
 * answers the file's last value, which its import stamped, until the sheet is
 * in. Then `pending` is empty and the formula is calculated.
 */
export function keptFormulaResultEntries(
  cells: Readonly<Record<string, { value?: unknown; keptResult?: unknown } | undefined>>,
  covered: ReadonlySet<string>,
  pending?: {
    readonly sheets: ReadonlySet<string>;
    readonly cells: Readonly<Record<string, { value?: unknown; display?: unknown; chartValueCache?: unknown } | undefined>>;
  },
): KeptFormulaResultEntry[] {
  const entries: KeptFormulaResultEntry[] = [];
  if (pending && pending.sheets.size > 0) {
    for (const key in pending.cells) {
      const cell = pending.cells[key];
      const formula = cell?.value;
      if (typeof formula !== "string" || formula.charCodeAt(0) !== 61 || covered.has(key)) continue;
      if (!formulaReadsSheets(formula, pending.sheets)) continue;
      const stamped = stampedFormulaResult(formula, cell!);
      if (stamped) entries.push([key, formula, stamped[0], stamped[1], "pending"]);
    }
  }
  const held = entries.length ? new Set(entries.map((entry) => entry[0])) : null;
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (!cell?.keptResult || covered.has(key)) continue;
    const kept = readKeptFormulaResult(cell.keptResult);
    if (!kept || cell.value !== kept.formula || !keptFormulaApplies(kept.formula) || held?.has(key)) continue;
    entries.push(kept.inputs
      ? [key, kept.formula, kept.kind, String(kept.value), null, kept.inputs]
      : [key, kept.formula, kept.kind, String(kept.value)]);
  }
  return entries;
}

/**
 * The engine's `keptFormulaResult` option over `entries`: the kept value of
 * the cell at (row, col) while the sheet's raw cell still holds its formula.
 */
export function keptFormulaResultOption(
  entries: readonly KeptFormulaResultEntry[] | undefined,
  readRaw: (row: number, col: number) => string | null,
): {
  keptFormulaResult?: (row: number, col: number) => CellValue | undefined;
  keptFormulaInputs?: (row: number, col: number) => readonly KeptFormulaInput[] | undefined;
} {
  if (!entries?.length) return {};
  const byKey = new Map<string, CellValue>();
  const formulas = new Map<string, string>();
  const inputsByKey = new Map<string, readonly KeptFormulaInput[]>();
  for (const [key, formula, kind, text, pending, inputs] of entries) {
    const kept = readKeptFormulaResult({
      formula,
      kind,
      value: kind === "number" ? Number(text) : kind === "boolean" ? text === "true" : text,
      ...(inputs ? { inputs } : {}),
    });
    if (!kept || (pending !== "pending" && !keptFormulaApplies(formula))) continue;
    byKey.set(key, keptFormulaResultValue(kept));
    formulas.set(key, formula);
    if (kept.inputs) inputsByKey.set(key, kept.inputs);
  }
  if (!byKey.size) return {};
  const holds = (row: number, col: number) => {
    const key = `${row}:${col}`;
    const formula = formulas.get(key);
    return formula !== undefined && readRaw(row, col) === formula ? key : null;
  };
  return {
    keptFormulaResult: (row: number, col: number) => {
      const key = holds(row, col);
      return key === null ? undefined : byKey.get(key);
    },
    ...(inputsByKey.size
      ? {
          keptFormulaInputs: (row: number, col: number) => {
            const key = holds(row, col);
            return key === null ? undefined : inputsByKey.get(key);
          },
        }
      : {}),
  };
}
