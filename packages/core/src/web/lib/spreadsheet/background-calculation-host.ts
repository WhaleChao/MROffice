/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cellMapChangesSince } from "./cell-map-lineage";
import type { WorksheetCalculationToken, WorksheetFormulaCoordinate } from "./calculation-mode";
import { worksheetFormulaResultKey } from "./calculation-mode";
import type { SheetEvaluatorOptions, WorkbookSheetInput } from "./formula-engine";
import type { BackgroundCalculationSnapshot, BackgroundCalculationDelta, BackgroundCalculationSheet, BackgroundCalculationSheetContext } from "./background-calculation-protocol";
import { whatIfTableBlocks, whatIfTableResultKeys } from "./what-if-table-definition";
import { excelDateSystemOf, isExcelDate1904, type ExcelDateSystemArg } from "./excel-1900-calendar";
import { isMaterializedLineStyleBlank } from "./line-styles";

/** Structural subset of SpreadsheetWindow's SheetModel/CellMap; no UI imports. */
export interface BackgroundCalculationHostSheet {
  readonly id: string;
  readonly name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly cells: Readonly<Record<string, { readonly value?: string }>>;
  readonly pivot?: unknown;
  /** What-If Data Tables: their results are formulas the session holds
   *  even where no cell is stored (what-if-table-definition). */
  readonly whatIfTables?: unknown;
}
export type BackgroundCalculationCellMap = BackgroundCalculationHostSheet["cells"];
export interface BackgroundCalculationHostBinding {
  readonly sheet: BackgroundCalculationHostSheet;
  /** Actual input used by createWorkbook, including its merge-masked reader. */
  readonly input: WorkbookSheetInput;
  /** Stable identity of everything affecting raw reads EXCEPT cells (merge mask). */
  readonly rawSemanticsKey: unknown;
  /** Host constructs metadata from the same inputs as these exact callbacks. */
  readonly resolverContext: {
    readonly options: SheetEvaluatorOptions | undefined;
    readonly serialized: BackgroundCalculationSheetContext;
  };
}
/** Formula count alone misses a single SUM/SUBTOTAL over a large range.
 * Only bounded arithmetic and direct cell references keep the synchronous fast
 * path. Calls, ranges, names, array/spill expressions and uncertain syntax use
 * the existing worker. Inspect at most 255 indexed formulas, never all cells.
 * This is scheduling only; the same evaluator retains formula semantics.
 */
export function shouldCalculateSpreadsheetInBackground(
  bindings: readonly BackgroundCalculationHostBinding[],
  coordinates: readonly WorksheetFormulaCoordinate[],
): boolean {
  if (coordinates.length === 0) return false;
  if (coordinates.length >= 256) return true;
  const byId = new Map(bindings.map(binding => [binding.sheet.id, binding]));
  const arithmeticToken = /(?:\s+|\$?[A-Z]{1,3}\$?[1-9]\d{0,6}(?![\w.])|(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?|[+*/^%=<>()-])/iy;
  return coordinates.some(cell => {
    const raw = byId.get(cell.sheetId)?.input.getRawCell(cell.row, cell.col)?.trim();
    if (!raw || raw.length > 512) return true;
    // Sticky tokenization has no repeated-pattern backtracking on malformed
    // imported formulas. Unknown tokens conservatively take the worker path.
    let offset = 0;
    while (offset < raw.length) {
      arithmeticToken.lastIndex = offset;
      if (!arithmeticToken.test(raw)) return true;
      offset = arithmeticToken.lastIndex;
    }
    return false;
  });
}

export type BackgroundCalculationHostFailure = { readonly ok: false; readonly reason:
  "unsupported-resolver" | "missing-context" | "invalid-sheet" | "invalid-cell" |
  "reset-required" | "stale-token" | "not-initialized" };
interface AcceptedEdit {
  readonly sheetId: string;
  readonly beforeCells: BackgroundCalculationCellMap;
  readonly afterCells: BackgroundCalculationCellMap;
  /** Complete affected keys, including deletions. Omit for unknown/bulk edits. */
  readonly knownKeys?: readonly string[];
}

/** Session-only hints from accepted immutable cell-map mutations, never authority. */
export class BackgroundCalculationDirtyJournal {
  #edges = new Map<string, AcceptedEdit[]>();
  record(edit: AcceptedEdit): void {
    if (edit.beforeCells === edit.afterCells) return;
    const edges = this.#edges.get(edit.sheetId) ?? [];
    if (edges.length >= 128) edges.length = 0; // Missing chain => safe one-time diff.
    edges.push({ ...edit, knownKeys: edit.knownKeys ? [...edit.knownKeys] : undefined });
    this.#edges.set(edit.sheetId, edges);
  }
  keys(sheetId: string, before: BackgroundCalculationCellMap, after: BackgroundCalculationCellMap): readonly string[] | null {
    if (before === after) return [];
    const edges = this.#edges.get(sheetId) ?? [];
    const keys = new Set<string>();
    let cursor = before;
    for (const edge of edges) {
      if (edge.beforeCells !== cursor) continue;
      if (!edge.knownKeys) return null;
      for (const key of edge.knownKeys) keys.add(key);
      cursor = edge.afterCells;
      if (cursor === after) return [...keys];
    }
    return null;
  }
  clear(): void { this.#edges.clear(); }
}

interface PreparedBinding {
  readonly binding: BackgroundCalculationHostBinding;
  readonly descriptor: Omit<BackgroundCalculationSheet, "cells">;
  readonly contextFingerprint: string;
}
function prepare(binding: BackgroundCalculationHostBinding): PreparedBinding | BackgroundCalculationHostFailure {
  const { sheet, input, resolverContext } = binding;
  if (!sheet.id || input.name !== (sheet.name || sheet.id) || !Number.isSafeInteger(sheet.rowCount) || !Number.isSafeInteger(sheet.columnCount) || sheet.rowCount < 1 || sheet.columnCount < 1) return { ok: false, reason: "invalid-sheet" };
  const options = input.options;
  const context = resolverContext.serialized;
  // A PivotTable's report is not an input: the session lays it over the
  // pre-report cells itself (calculation-pivots), so only the overlaid
  // reader of an already overlaid input is refused.
  if (options?.resolvePivotData) return { ok: false, reason: "unsupported-resolver" };
  if (resolverContext.options !== options) return { ok: false, reason: "missing-context" };
  // spillAnchors and isSpillTargetOccupied need no data of their own: the
  // worker derives both from the cells it holds and the covered cells the
  // context names.
  const supported = new Set(["usedRange", "resolveName", "resolveTableRef", "isRowHiddenByFilter", "isRowManuallyHidden", "isTextCell", "cellPicture", "cellPhonetic", "keptFormulaResult", "keptFormulaInputs", "isFormulaHidden", "legacyArrayExtent", "spillAnchors", "isSpillTargetOccupied", "whatIfTables", "columnWidth", "columnWidthIsDefault"]);
  if (Object.keys(options ?? {}).some((key) => !supported.has(key))) return { ok: false, reason: "unsupported-resolver" };
  const pairs = [
    [options?.resolveName, context.namedRanges], [options?.resolveTableRef, context.tables],
    [options?.isRowHiddenByFilter, context.filterHiddenRows], [options?.isRowManuallyHidden, context.manuallyHiddenRows],
    [options?.isTextCell, context.textCells],
    [options?.cellPicture, context.pictureCells],
    [options?.cellPhonetic, context.phoneticCells],
    [options?.keptFormulaResult, context.keptResults],
    [options?.isFormulaHidden, context.formulaHiddenCells],
    [options?.legacyArrayExtent, context.legacyArrays],
    [options?.whatIfTables, context.whatIfTables],
  ] as const;
  if (pairs.some(([callback, data]) => callback ? !Array.isArray(data) : Boolean(data?.length))) return { ok: false, reason: "missing-context" };
  // A held result's inputs travel in its kept-result entry.
  if (options?.keptFormulaInputs && !options.keptFormulaResult) return { ok: false, reason: "missing-context" };
  // Both width readers come from the one column width reading (columnWidthOption).
  if (Boolean(options?.columnWidth) !== Array.isArray(context.columnWidths?.columns)) return { ok: false, reason: "missing-context" };
  if (Boolean(options?.columnWidthIsDefault) !== Boolean(options?.columnWidth)) return { ok: false, reason: "missing-context" };
  try {
    const serialized = structuredClone(context);
    const usedRange = options?.usedRange ? { ...options.usedRange } : null;
    return { binding, descriptor: { id: sheet.id, name: input.name, rowCount: sheet.rowCount, columnCount: sheet.columnCount, usedRange, context: serialized }, contextFingerprint: JSON.stringify(serialized) };
  } catch { return { ok: false, reason: "unsupported-resolver" }; }
}
function coordinate(key: string, sheet: BackgroundCalculationHostSheet): { row: number; col: number } | null {
  // Read digit by digit: a snapshot passes every cell of a workbook through
  // here, where a regular expression and a split were 0.1 s per 250,000.
  let row = 0, col = 0, part = 0, digits = 0;
  for (let index = 0; index < key.length; index += 1) {
    const code = key.charCodeAt(index);
    if (code === 58) {
      if (part === 1 || digits === 0) return null;
      part = 1;
      digits = 0;
    } else if (code >= 48 && code <= 57) {
      digits += 1;
      if (part === 0) row = row * 10 + code - 48;
      else col = col * 10 + code - 48;
    } else return null;
  }
  if (part === 0 || digits === 0) return null;
  return Number.isSafeInteger(row) && Number.isSafeInteger(col) && row >= 0 && col >= 0 && row < sheet.rowCount && col < sheet.columnCount && `${row}:${col}` === key ? { row, col } : null;
}
function fallbackChangedKeys(before: BackgroundCalculationCellMap, after: BackgroundCalculationCellMap): string[] {
  const keys: string[] = [];
  // Walked in place, not through Object.keys copies (a sorted 730,000-cell
  
  const own = Object.prototype.hasOwnProperty;
  for (const key in after) if (own.call(after, key) && before[key] !== after[key]) keys.push(key);
  for (const key in before) if (own.call(before, key) && !own.call(after, key)) keys.push(key);
  return keys;
}

/**
 * Adapter for existing host evaluators. Does not rebuild filter/pivot/name/table
 * semantics. Its private index tracks formula coordinates, never calculated values.
 */
export class BackgroundCalculationHost {
  #token: WorksheetCalculationToken | null = null;
  #bindings: PreparedBinding[] = [];
  #formulas = new Map<string, WorksheetFormulaCoordinate>();
  /** The coordinate lists handed out since the formula index last changed,
   * by target sheet: an edit that leaves the formulas where they are (a
   * typed value) hands the very same frozen list out again, so the request
   * built from it is not copied and checked whole again per edit. */
  #coordinateLists = new Map<string | null, readonly WorksheetFormulaCoordinate[]>();
  reset(token: WorksheetCalculationToken, bindings: readonly BackgroundCalculationHostBinding[], workbook?: string | { readonly name?: string; readonly folder?: string }, date1904?: ExcelDateSystemArg,
    iteration?: { readonly maxIterations: number; readonly maxChange: number; readonly seeds?: ReadonlyArray<readonly [string, number, number, number]> }):
    { ok: true; snapshot: BackgroundCalculationSnapshot; coordinates: readonly WorksheetFormulaCoordinate[] } | BackgroundCalculationHostFailure {
    const prepared: PreparedBinding[] = [];
    const sheets: BackgroundCalculationSheet[] = [];
    const formulas = new Map<string, WorksheetFormulaCoordinate>();
    const ids = new Set<string>(); const names = new Set<string>();
    const workbookName = typeof workbook === "string" ? workbook : workbook?.name;
    const workbookFolder = typeof workbook === "string" ? undefined : workbook?.folder;
    for (const binding of bindings) {
      const entry = prepare(binding);
      if ("ok" in entry) return entry;
      if (ids.has(binding.sheet.id) || names.has(binding.input.name.toLowerCase())) return { ok: false, reason: "invalid-sheet" };
      ids.add(binding.sheet.id); names.add(binding.input.name.toLowerCase());
      const cells: Array<readonly [string, string]> = [];
      // A What-If Data Table's result is a formula wherever it stands, also
      // where no cell is stored (a blank result): the session holds it too.
      const tableKeys = whatIfTableBlocks(binding.input.options?.whatIfTables, binding.sheet)
        .flatMap(whatIfTableResultKeys)
        .filter((key) => !Object.prototype.hasOwnProperty.call(binding.sheet.cells, key));
      for (const key of [...Object.keys(binding.sheet.cells), ...tableKeys]) {
        // A blank a row or column format wrote out (lib/spreadsheet/line-styles)
        // reads "" and is no text: nothing of it is held (some 50,000 of the
        // 1.6 MB benchmark's, each parsed and read here on every open).
        if (isMaterializedLineStyleBlank(binding.sheet.cells[key])) continue;
        const cell = coordinate(key, binding.sheet);
        if (!cell) return { ok: false, reason: "invalid-cell" };
        const raw = binding.input.getRawCell(cell.row, cell.col);
        if (raw !== null && typeof raw !== "string") return { ok: false, reason: "invalid-cell" };
        // A cell that is only formatted reads "" and calculates as the blank
        // it is: leaving it out keeps a sheet of formatted blanks inside the
        // session's budget (S14_tdf131536: 200,000 of them past the 100,000-
        // cell snapshot ceiling, so every formula stayed blank; X-096). A
        // stored empty TEXT cell ("" as a value) is kept: it is not blank.
        if (raw !== null && (raw !== "" || binding.input.options?.isTextCell?.(cell.row, cell.col))) cells.push([key, raw]);
        if (raw?.trim().startsWith("=")) formulas.set(worksheetFormulaResultKey(binding.input.name, cell.row, cell.col), { sheetId: binding.sheet.id, sheetName: binding.input.name, ...cell });
      }
      prepared.push(entry); sheets.push({ ...entry.descriptor, cells });
    }
    if (sheets.length === 0) return { ok: false, reason: "invalid-sheet" };
    this.#token = { ...token }; this.#bindings = prepared; this.#formulas = formulas; this.#coordinateLists.clear();
    return { ok: true, snapshot: { token: { ...token }, sheets, workbookName, ...(workbookName && workbookFolder ? { workbookFolder } : {}), ...(isExcelDate1904(date1904) ? { date1904: true } : excelDateSystemOf(date1904) === "1900-noncompat" ? { dateSystem: "1900-noncompat" as const } : {}), ...(iteration ? { iteration: { maxIterations: iteration.maxIterations, maxChange: iteration.maxChange, ...(iteration.seeds?.length ? { seeds: iteration.seeds.map((seed) => [...seed] as const) } : {}) } } : {}) }, coordinates: this.coordinates() };
  }

  advance(token: WorksheetCalculationToken, bindings: readonly BackgroundCalculationHostBinding[], journal?: BackgroundCalculationDirtyJournal):
    { ok: true; delta: BackgroundCalculationDelta; coordinates: readonly WorksheetFormulaCoordinate[]; scannedSheets: readonly string[] } | BackgroundCalculationHostFailure {
    const beforeToken = this.#token;
    if (!beforeToken) return { ok: false, reason: "not-initialized" };
    if (token.pageId !== beforeToken.pageId || token.loadEpoch !== beforeToken.loadEpoch || token.topologyRevision !== beforeToken.topologyRevision || bindings.length !== this.#bindings.length) return { ok: false, reason: "reset-required" };
    if (token.generation <= beforeToken.generation || token.workbookRevision < beforeToken.workbookRevision) return { ok: false, reason: "stale-token" };
    const prepared: PreparedBinding[] = [];
    const changed: Array<BackgroundCalculationDelta["cells"][number]> = [];
    const contexts: Array<NonNullable<BackgroundCalculationDelta["contexts"]>[number]> = [];
    const formulaUpdates = new Map<string, WorksheetFormulaCoordinate | null>();
    const scannedSheets: string[] = [];
    for (let index = 0; index < bindings.length; index += 1) {
      const binding = bindings[index]; const previous = this.#bindings[index];
      // A binding the editor kept (its sheet unchanged) was prepared already:
      // its context is not cloned and serialized again on every edit.
      const entry = binding === previous.binding ? previous : prepare(binding);
      if ("ok" in entry) return entry;
      const old = previous.descriptor; const next = entry.descriptor;
      if (old.id !== next.id || old.name !== next.name || old.rowCount !== next.rowCount || old.columnCount !== next.columnCount || !Object.is(previous.binding.rawSemanticsKey, binding.rawSemanticsKey) || JSON.stringify(old.usedRange) !== JSON.stringify(next.usedRange)) return { ok: false, reason: "reset-required" };
      if (entry.contextFingerprint !== previous.contextFingerprint) contexts.push({ sheetId: next.id, context: next.context! });
      let keys = journal?.keys(next.id, previous.binding.sheet.cells, binding.sheet.cells) ?? null;
      if (previous.binding.sheet.cells === binding.sheet.cells) keys = [];
      keys ??= cellMapChangesSince(binding.sheet.cells, previous.binding.sheet.cells);
      if (keys === null) { keys = fallbackChangedKeys(previous.binding.sheet.cells, binding.sheet.cells); scannedSheets.push(next.id); }
      for (const key of keys) {
        const cell = coordinate(key, binding.sheet);
        if (!cell) return { ok: false, reason: "invalid-cell" };
        const raw = binding.input.getRawCell(cell.row, cell.col);
        if (raw !== null && typeof raw !== "string") return { ok: false, reason: "invalid-cell" };
        if (raw === previous.binding.input.getRawCell(cell.row, cell.col)) continue;
        changed.push({ sheetId: next.id, ...cell, raw });
        formulaUpdates.set(worksheetFormulaResultKey(next.name, cell.row, cell.col), raw?.trim().startsWith("=") ? { sheetId: next.id, sheetName: next.name, ...cell } : null);
      }
      prepared.push(entry);
    }
    for (const [key, cell] of formulaUpdates) {
      const before = this.#formulas.get(key);
      if (cell) {
        if (before && before.sheetId === cell.sheetId && before.sheetName === cell.sheetName && before.row === cell.row && before.col === cell.col) continue;
        this.#formulas.set(key, cell);
      } else if (before) this.#formulas.delete(key);
      else continue;
      this.#coordinateLists.clear();
    }
    this.#token = { ...token }; this.#bindings = prepared; journal?.clear();
    return { ok: true, delta: { baseToken: beforeToken, token: { ...token }, cells: changed, ...(contexts.length ? { contexts } : {}) }, coordinates: this.coordinates(), scannedSheets };
  }
  coordinates(targetSheetId: string | null = null): readonly WorksheetFormulaCoordinate[] {
    const known = this.#coordinateLists.get(targetSheetId);
    if (known) return known;
    const list = Object.freeze([...this.#formulas.values()].filter((cell) => targetSheetId === null || cell.sheetId === targetSheetId).map((cell) => Object.freeze({ ...cell })));
    this.#coordinateLists.set(targetSheetId, list);
    return list;
  }
  clear(): void { this.#token = null; this.#bindings = []; this.#formulas.clear(); this.#coordinateLists.clear(); }
}
