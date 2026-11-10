/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TranslationFn } from "../errors/user-message";

/**
 * Calculation mode (公式 ▸ 计算选项:自动 / 手动,F9).
 *
 * Excel and WPS let a workbook stop recalculating on every edit. In manual
 * mode a formula keeps showing the value from the LAST calculation and the
 * workbook is marked dirty; F9 recalculates everything, Shift+F9 only the
 * active sheet. This module is the state machine for that — which mode is in
 * force, whether an edit leaves work outstanding, and what a given key press
 * should recalculate — kept pure so the host's hot display path only has to ask
 * "am I allowed to recompute?".
 */

export type CalculationMode = "auto" | "manual";

export interface CalculationState {
  mode: CalculationMode;
  /** True when edits have landed that manual mode has not recalculated yet. */
  dirty: boolean;
}

export type CalculationScope = "workbook" | "sheet" | "none";

export const MAX_WORKSHEET_CALCULATION_FORMULAS = 20_000;
export const MAX_WORKSHEET_CALCULATION_RESULTS = 200_000;

export interface WorksheetCalculationSheetDescriptor {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
}

export interface WorksheetCalculationSheetAuthority extends WorksheetCalculationSheetDescriptor {
  /** Any sheet-model edit, including chart configuration, advances this. */
  sourceRevision: number;
  /** Insert/delete/shift/sort-style coordinate changes advance this. */
  coordinateRevision: number;
}

/**
 * Session-only authority. It is intentionally not serialized: reopening a page
 * creates a new load epoch, so cached formula results can never be treated as a
 * durable workbook truth merely because content_json survived.
 */
export interface WorksheetCalculationAuthority {
  pageId: string;
  loadEpoch: number;
  generation: number;
  workbookRevision: number;
  topologyRevision: number;
  sheets: readonly WorksheetCalculationSheetAuthority[];
}

export interface WorksheetFormulaCoordinate {
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
}

export interface WorksheetCalculationLedger<T> {
  pageId: string;
  loadEpoch: number;
  createdGeneration: number;
  createdWorkbookRevision: number;
  topologyRevision: number;
  sheets: readonly (WorksheetCalculationSheetAuthority & {
    calculatedSourceRevision: number;
    formulaKeys: readonly string[];
  })[];
  results: ReadonlyMap<string, T>;
}

export interface WorksheetCalculationToken {
  pageId: string;
  loadEpoch: number;
  generation: number;
  workbookRevision: number;
  topologyRevision: number;
  targetSheetId: string | null;
}

export type WorksheetCalculationFailureReason =
  | "invalid-authority"
  | "invalid-target"
  | "formula-limit"
  | "result-limit"
  | "invalid-coordinate"
  | "stale-page"
  | "stale-generation"
  | "topology-changed"
  | "coordinate-changed"
  | "missing-formula-seed";

export type WorksheetCalculationCheck =
  { ok: true } | { ok: false; reason: WorksheetCalculationFailureReason };

const MAX_WORKSHEET_CALCULATION_SHEETS = 1_024;

function normalizeSheetDescriptors(
  sheets: readonly WorksheetCalculationSheetDescriptor[],
): readonly WorksheetCalculationSheetDescriptor[] | null {
  if (
    !Array.isArray(sheets) ||
    sheets.length < 1 ||
    sheets.length > MAX_WORKSHEET_CALCULATION_SHEETS
  ) {
    return null;
  }
  const ids = new Set<string>();
  const names = new Set<string>();
  const normalized: WorksheetCalculationSheetDescriptor[] = [];
  for (let index = 0; index < sheets.length; index += 1) {
    if (!(index in sheets)) return null;
    const sheet = sheets[index];
    if (!sheet || typeof sheet !== "object") return null;
    const id = typeof sheet.id === "string" ? sheet.id.trim() : "";
    const name = typeof sheet.name === "string" ? sheet.name.trim() : "";
    const rowCount = sheet.rowCount;
    const columnCount = sheet.columnCount;
    if (
      id === "" ||
      name === "" ||
      ids.has(id) ||
      names.has(name.toLowerCase()) ||
      !Number.isSafeInteger(rowCount) ||
      rowCount < 1 ||
      !Number.isSafeInteger(columnCount) ||
      columnCount < 1
    ) {
      return null;
    }
    ids.add(id);
    names.add(name.toLowerCase());
    normalized.push(Object.freeze({ id, name, rowCount, columnCount }));
  }
  return Object.freeze(normalized);
}

export function worksheetFormulaResultKey(
  sheetName: string,
  row: number,
  col: number,
): string {
  return `${sheetName.toLowerCase()}\u0000${row}:${col}`;
}

export function resetWorksheetCalculationAuthority(
  pageId: string,
  sheets: readonly WorksheetCalculationSheetDescriptor[],
  previous?: WorksheetCalculationAuthority | null,
): WorksheetCalculationAuthority | null {
  const normalized = normalizeSheetDescriptors(sheets);
  if (typeof pageId !== "string" || pageId.trim() === "" || !normalized) {
    return null;
  }
  return Object.freeze({
    pageId,
    loadEpoch: (previous?.loadEpoch ?? 0) + 1,
    generation: (previous?.generation ?? 0) + 1,
    workbookRevision: 0,
    topologyRevision: (previous?.topologyRevision ?? 0) + 1,
    sheets: Object.freeze(
      normalized.map((sheet) =>
        Object.freeze({
          ...sheet,
          sourceRevision: 0,
          coordinateRevision: 0,
        }),
      ),
    ),
  });
}

export function advanceWorksheetCalculationAuthority(
  current: WorksheetCalculationAuthority,
  sheets: readonly WorksheetCalculationSheetDescriptor[],
  mutation:
    | { kind: "content"; sheetId: string; coordinateChanged?: boolean }
    | { kind: "topology" },
): WorksheetCalculationAuthority | null {
  const normalized = normalizeSheetDescriptors(sheets);
  if (!normalized) return null;
  if (mutation.kind === "content") {
    if (
      normalized.length !== current.sheets.length ||
      !current.sheets.some((sheet) => sheet.id === mutation.sheetId)
    ) {
      return null;
    }
    for (let index = 0; index < normalized.length; index += 1) {
      const before = current.sheets[index];
      const after = normalized[index];
      if (!before || before.id !== after.id || before.name !== after.name) {
        return null;
      }
    }
    return Object.freeze({
      ...current,
      generation: current.generation + 1,
      workbookRevision: current.workbookRevision + 1,
      sheets: Object.freeze(
        normalized.map((sheet, index) => {
          const before = current.sheets[index];
          const changed = sheet.id === mutation.sheetId;
          // A smaller extent can drop formula cells; a larger one (typing past
          // the stored edge) moves nothing, so it is a content edit.
          const dimensionsChanged =
            sheet.rowCount < before.rowCount ||
            sheet.columnCount < before.columnCount;
          return Object.freeze({
            ...sheet,
            sourceRevision: before.sourceRevision + (changed ? 1 : 0),
            coordinateRevision:
              before.coordinateRevision +
              (changed && (mutation.coordinateChanged || dimensionsChanged)
                ? 1
                : 0),
          });
        }),
      ),
    });
  }

  const beforeById = new Map(current.sheets.map((sheet) => [sheet.id, sheet]));
  return Object.freeze({
    ...current,
    generation: current.generation + 1,
    workbookRevision: current.workbookRevision + 1,
    topologyRevision: current.topologyRevision + 1,
    sheets: Object.freeze(
      normalized.map((sheet) => {
        const before = beforeById.get(sheet.id);
        return Object.freeze({
          ...sheet,
          sourceRevision: (before?.sourceRevision ?? 0) + 1,
          coordinateRevision: (before?.coordinateRevision ?? 0) + 1,
        });
      }),
    ),
  });
}

export function beginWorksheetCalculation(
  current: WorksheetCalculationAuthority,
  targetSheetId: string | null,
): {
  authority: WorksheetCalculationAuthority;
  token: WorksheetCalculationToken;
} {
  const authority = Object.freeze({
    ...current,
    generation: current.generation + 1,
  });
  return {
    authority,
    token: Object.freeze({
      pageId: authority.pageId,
      loadEpoch: authority.loadEpoch,
      generation: authority.generation,
      workbookRevision: authority.workbookRevision,
      topologyRevision: authority.topologyRevision,
      targetSheetId,
    }),
  };
}

export function worksheetCalculationTokenIsCurrent(
  authority: WorksheetCalculationAuthority,
  token: WorksheetCalculationToken,
): boolean {
  return (
    authority.pageId === token.pageId &&
    authority.loadEpoch === token.loadEpoch &&
    authority.generation === token.generation &&
    authority.workbookRevision === token.workbookRevision &&
    authority.topologyRevision === token.topologyRevision
  );
}

function formulaCoordinateKey(
  authority: WorksheetCalculationAuthority,
  formula: WorksheetFormulaCoordinate,
): { key: string; sheet: WorksheetCalculationSheetAuthority } | null {
  const sheet = authority.sheets.find(
    (candidate) => candidate.id === formula.sheetId,
  );
  if (
    !sheet ||
    sheet.name !== formula.sheetName ||
    !Number.isSafeInteger(formula.row) ||
    !Number.isSafeInteger(formula.col) ||
    formula.row < 0 ||
    formula.col < 0 ||
    formula.row >= sheet.rowCount ||
    formula.col >= sheet.columnCount
  ) {
    return null;
  }
  return {
    key: worksheetFormulaResultKey(sheet.name, formula.row, formula.col),
    sheet,
  };
}

export function createWorksheetCalculationLedger<T>(params: {
  authority: WorksheetCalculationAuthority;
  formulas: readonly WorksheetFormulaCoordinate[];
  results: ReadonlyMap<string, T>;
}): WorksheetCalculationLedger<T> | null {
  const { authority, formulas, results } = params;
  if (
    !Array.isArray(formulas) ||
    formulas.length > MAX_WORKSHEET_CALCULATION_FORMULAS ||
    !(results instanceof Map) ||
    results.size > MAX_WORKSHEET_CALCULATION_RESULTS
  ) {
    return null;
  }
  const formulaKeysBySheet = new Map<string, string[]>();
  const copiedResults = new Map<string, T>();
  const seen = new Set<string>();
  for (let index = 0; index < formulas.length; index += 1) {
    if (!(index in formulas)) return null;
    const resolved = formulaCoordinateKey(authority, formulas[index]);
    if (!resolved || seen.has(resolved.key) || !results.has(resolved.key)) {
      return null;
    }
    seen.add(resolved.key);
    copiedResults.set(resolved.key, results.get(resolved.key)!);
    const bucket = formulaKeysBySheet.get(resolved.sheet.id);
    if (bucket) bucket.push(resolved.key);
    else formulaKeysBySheet.set(resolved.sheet.id, [resolved.key]);
  }
  return Object.freeze({
    pageId: authority.pageId,
    loadEpoch: authority.loadEpoch,
    createdGeneration: authority.generation,
    createdWorkbookRevision: authority.workbookRevision,
    topologyRevision: authority.topologyRevision,
    sheets: Object.freeze(
      authority.sheets.map((sheet) =>
        Object.freeze({
          ...sheet,
          calculatedSourceRevision: sheet.sourceRevision,
          formulaKeys: Object.freeze(formulaKeysBySheet.get(sheet.id) ?? []),
        }),
      ),
    ),
    results: copiedResults,
  });
}

export function validateWorksheetCalculationRequest<T>(params: {
  authority: WorksheetCalculationAuthority;
  ledger: WorksheetCalculationLedger<T>;
  token: WorksheetCalculationToken;
  targetSheetId: string;
  formulas: readonly WorksheetFormulaCoordinate[];
}): WorksheetCalculationCheck {
  const { authority, ledger, token, targetSheetId, formulas } = params;
  if (!worksheetCalculationTokenIsCurrent(authority, token)) {
    return { ok: false, reason: "stale-generation" };
  }
  if (
    ledger.pageId !== authority.pageId ||
    ledger.loadEpoch !== authority.loadEpoch
  ) {
    return { ok: false, reason: "stale-page" };
  }
  if (
    ledger.topologyRevision !== authority.topologyRevision ||
    ledger.sheets.length !== authority.sheets.length
  ) {
    return { ok: false, reason: "topology-changed" };
  }
  if (!authority.sheets.some((sheet) => sheet.id === targetSheetId)) {
    return { ok: false, reason: "invalid-target" };
  }
  // A PivotTable is no reason to refuse: its output is derived from the
  // workbook it reads, never a formula the ledger freezes, and it changes on
  // its own Refresh as Excel's does in any calculation mode.
  for (let index = 0; index < authority.sheets.length; index += 1) {
    const current = authority.sheets[index];
    const recorded = ledger.sheets[index];
    if (
      !recorded ||
      recorded.id !== current.id ||
      recorded.name !== current.name ||
      current.rowCount < recorded.rowCount ||
      current.columnCount < recorded.columnCount
    ) {
      return { ok: false, reason: "topology-changed" };
    }
    if (recorded.coordinateRevision !== current.coordinateRevision) {
      return { ok: false, reason: "coordinate-changed" };
    }
  }
  if (
    !Array.isArray(formulas) ||
    formulas.length > MAX_WORKSHEET_CALCULATION_FORMULAS
  ) {
    return { ok: false, reason: "formula-limit" };
  }
  const seen = new Set<string>();
  for (let index = 0; index < formulas.length; index += 1) {
    if (!(index in formulas)) {
      return { ok: false, reason: "invalid-coordinate" };
    }
    const resolved = formulaCoordinateKey(authority, formulas[index]);
    if (!resolved || seen.has(resolved.key)) {
      return { ok: false, reason: "invalid-coordinate" };
    }
    seen.add(resolved.key);
    if (
      resolved.sheet.id !== targetSheetId &&
      !ledger.results.has(resolved.key)
    ) {
      return { ok: false, reason: "missing-formula-seed" };
    }
  }
  return { ok: true };
}

/**
 * Authenticate an automatic-mode worksheet refresh without requiring a manual
 * calculation ledger. The engine still receives frozen non-target formula
 * results separately; this gate proves that the live target and every formula
 * coordinate belong to the exact generation captured by Shift+F9.
 */
export function validateAutomaticWorksheetCalculationRequest(params: {
  authority: WorksheetCalculationAuthority;
  token: WorksheetCalculationToken;
  targetSheetId: string;
  formulas: readonly WorksheetFormulaCoordinate[];
}): WorksheetCalculationCheck {
  const { authority, token, targetSheetId, formulas } = params;
  if (
    token.targetSheetId !== targetSheetId ||
    !worksheetCalculationTokenIsCurrent(authority, token)
  ) {
    return { ok: false, reason: "stale-generation" };
  }
  if (!authority.sheets.some((sheet) => sheet.id === targetSheetId)) {
    return { ok: false, reason: "invalid-target" };
  }
  if (
    !Array.isArray(formulas) ||
    formulas.length > MAX_WORKSHEET_CALCULATION_FORMULAS
  ) {
    return { ok: false, reason: "formula-limit" };
  }
  const seen = new Set<string>();
  for (let index = 0; index < formulas.length; index += 1) {
    if (!(index in formulas)) {
      return { ok: false, reason: "invalid-coordinate" };
    }
    const resolved = formulaCoordinateKey(authority, formulas[index]);
    if (!resolved || seen.has(resolved.key)) {
      return { ok: false, reason: "invalid-coordinate" };
    }
    seen.add(resolved.key);
  }
  return { ok: true };
}

export function publishWorksheetCalculationResults<T>(params: {
  authority: WorksheetCalculationAuthority;
  ledger: WorksheetCalculationLedger<T>;
  token: WorksheetCalculationToken;
  targetSheetId: string;
  formulas: readonly WorksheetFormulaCoordinate[];
  targetResults: ReadonlyMap<string, T>;
}):
  | { ok: true; ledger: WorksheetCalculationLedger<T> }
  | { ok: false; reason: WorksheetCalculationFailureReason } {
  const check = validateWorksheetCalculationRequest(params);
  if (!check.ok) return check;
  if (
    !(params.targetResults instanceof Map) ||
    params.targetResults.size > MAX_WORKSHEET_CALCULATION_RESULTS
  ) {
    return { ok: false, reason: "result-limit" };
  }
  const nextResults = new Map<string, T>();
  const formulaKeysBySheet = new Map<string, string[]>();
  for (const formula of params.formulas) {
    const resolved = formulaCoordinateKey(params.authority, formula);
    if (!resolved) return { ok: false, reason: "invalid-coordinate" };
    const source =
      resolved.sheet.id === params.targetSheetId
        ? params.targetResults
        : params.ledger.results;
    if (!source.has(resolved.key)) {
      return { ok: false, reason: "missing-formula-seed" };
    }
    nextResults.set(resolved.key, source.get(resolved.key)!);
    const bucket = formulaKeysBySheet.get(resolved.sheet.id);
    if (bucket) bucket.push(resolved.key);
    else formulaKeysBySheet.set(resolved.sheet.id, [resolved.key]);
  }
  if (nextResults.size > MAX_WORKSHEET_CALCULATION_RESULTS) {
    return { ok: false, reason: "result-limit" };
  }
  const previousById = new Map(
    params.ledger.sheets.map((sheet) => [sheet.id, sheet]),
  );
  return {
    ok: true,
    ledger: Object.freeze({
      pageId: params.authority.pageId,
      loadEpoch: params.authority.loadEpoch,
      createdGeneration: params.token.generation,
      createdWorkbookRevision: params.authority.workbookRevision,
      topologyRevision: params.authority.topologyRevision,
      sheets: Object.freeze(
        params.authority.sheets.map((sheet) => {
          const previous = previousById.get(sheet.id)!;
          return Object.freeze({
            ...sheet,
            calculatedSourceRevision:
              sheet.id === params.targetSheetId
                ? sheet.sourceRevision
                : previous.calculatedSourceRevision,
            formulaKeys: Object.freeze(formulaKeysBySheet.get(sheet.id) ?? []),
          });
        }),
      ),
      results: nextResults,
    }),
  };
}

/**
 * Admit only results whose formula coordinates still have the same identity.
 * Content-only edits deliberately keep the previous result, while any sheet
 * topology or coordinate drift drops the affected seed instead of attaching a
 * stale value to a different formula after insert/delete/shift/sort.
 */
export function worksheetCalculationSeedSnapshot<T>(
  authority: WorksheetCalculationAuthority,
  ledger: WorksheetCalculationLedger<T>,
): ReadonlyMap<string, T> {
  if (
    ledger.pageId !== authority.pageId ||
    ledger.loadEpoch !== authority.loadEpoch ||
    ledger.topologyRevision !== authority.topologyRevision ||
    ledger.sheets.length !== authority.sheets.length
  ) {
    return new Map();
  }
  const denied = new Set<string>();
  for (let index = 0; index < authority.sheets.length; index += 1) {
    const current = authority.sheets[index];
    const recorded = ledger.sheets[index];
    if (
      !recorded ||
      recorded.id !== current.id ||
      recorded.name !== current.name ||
      // Typing past the stored extent grows it without moving a cell: a
      // content edit, so the sheet's results stay (Excel keeps calculated
      // values in Manual mode). A smaller extent, or any move, drops them.
      current.rowCount < recorded.rowCount ||
      current.columnCount < recorded.columnCount ||
      recorded.coordinateRevision !== current.coordinateRevision
    ) {
      for (const key of recorded?.formulaKeys ?? []) denied.add(key);
    }
  }
  if (denied.size === 0) return ledger.results;
  const admitted = new Map<string, T>();
  for (const [key, value] of ledger.results) {
    if (!denied.has(key)) admitted.set(key, value);
  }
  return admitted;
}

export function worksheetCalculationHasOutstanding<T>(
  authority: WorksheetCalculationAuthority,
  ledger: WorksheetCalculationLedger<T> | null,
): boolean {
  if (
    !ledger ||
    ledger.pageId !== authority.pageId ||
    ledger.loadEpoch !== authority.loadEpoch ||
    ledger.topologyRevision !== authority.topologyRevision ||
    ledger.sheets.length !== authority.sheets.length
  ) {
    return true;
  }
  return authority.sheets.some((sheet, index) => {
    const calculated = ledger.sheets[index];
    return (
      !calculated ||
      calculated.id !== sheet.id ||
      calculated.coordinateRevision !== sheet.coordinateRevision ||
      calculated.calculatedSourceRevision !== sheet.sourceRevision
    );
  });
}

export function withCalculationDirty(
  state: CalculationState,
  dirty: boolean,
): CalculationState {
  return state.dirty === dirty ? state : { ...state, dirty };
}

export const DEFAULT_CALCULATION_STATE: CalculationState = {
  mode: "auto",
  dirty: false,
};

/** Read a persisted mode, defaulting to Excel's own default. */
export function normalizeCalculationMode(raw: unknown): CalculationMode {
  return raw === "manual" ? "manual" : "auto";
}

/**
 * Whether the grid may evaluate formulas from the CURRENT cells. In manual mode
 * with outstanding edits it may not: the last calculated values stand until the
 * user asks for a recalculation, which is the whole point of the mode.
 */
export function shouldEvaluateLive(state: CalculationState): boolean {
  return state.mode === "auto" || !state.dirty;
}

/** Apply an edit: auto stays clean, manual accumulates outstanding work. */
export function afterEdit(state: CalculationState): CalculationState {
  if (state.mode === "auto")
    return state.dirty ? { ...state, dirty: false } : state;
  return state.dirty ? state : { ...state, dirty: true };
}

/** Switching to auto settles the outstanding work; switching to manual keeps it. */
export function withMode(
  state: CalculationState,
  mode: CalculationMode,
): CalculationState {
  if (state.mode === mode) return state;
  return mode === "auto"
    ? { mode, dirty: false }
    : { mode, dirty: state.dirty };
}

/**
 * Settle the outstanding work — but only a WORKBOOK recalculation may do that.
 * Shift+F9 recalculates the active sheet, so the rest of the workbook can still
 * be stale; clearing the flag there told the user everything was calculated
 * when it was not.
 */
export function afterRecalculation(
  state: CalculationState,
  scope: CalculationScope = "workbook",
): CalculationState {
  if (scope !== "workbook") return state;
  return state.dirty ? { ...state, dirty: false } : state;
}

/**
 * What a key press should recalculate. F9 is the workbook, Shift+F9 the active
 * sheet; anything else is not a recalculation request. Excel accepts both in
 * auto mode too (a forced refresh), so the mode does not gate this.
 */
export function recalculationScopeForKey(event: {
  key: string;
  shiftKey?: boolean;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
}): CalculationScope {
  if (event.key !== "F9" || event.ctrlKey || event.metaKey || event.altKey) {
    return "none";
  }
  return event.shiftKey ? "sheet" : "workbook";
}

/** Status text for the ribbon/status bar, in the user's language. */
export function calculationStatusText(
  state: CalculationState,
  translate: TranslationFn,
): string | null {
  if (!state.dirty) return null;
  return translate("spreadsheet.helperCopy.calculation.pending");
}

/**
 * The ledger after whole rows or columns were inserted into or deleted from
 * worksheets it holds results for, the results moved with their formulas.
 * Under Manual calculation Excel keeps showing a moved formula's last result
 * (measured in Excel for Mac 16: =SEQUENCE(5) and =B1:B5*10 keep their five
 * values after a row is inserted or deleted inside them, until F9). Without
 * it the coordinate change dropped every result of the sheet and the grid
 * went blank until the next calculation.
 *
 * `mapKey` takes a sheet's "row:col" to where the edit put it (null: deleted).
 * Each edited sheet's recorded coordinates are brought to `after`'s, as if
 * the last calculation had been made on the moved grid; its calculated source
 * revision stays, so the workbook still has calculation outstanding. A ledger
 * that was not current for a sheet before the edit is left as it was (the
 * edit drops it, as before).
 */
export function worksheetCalculationLedgerThroughLineEdit<T>(params: {
  before: WorksheetCalculationAuthority;
  after: WorksheetCalculationAuthority;
  ledger: WorksheetCalculationLedger<T>;
  edits: ReadonlyMap<string, (key: string) => string | null>;
}): WorksheetCalculationLedger<T> {
  return worksheetCalculationLedgerLineEdit(params).ledger;
}

/**
 * The last results a whole-line edit took off the ledger with the cells it
 * deleted, by sheet: what an Undo of that edit brings back with the cells
 * (Excel's Undo restores the deleted cells as they were, their last result
 * included, so a restored formula under Manual calculation shows it until F9,
 * not an empty cell). Valid only for the ledger it was taken from (its page,
 * load and topology).
 */
export interface WorksheetCalculationLinesDropped<T> {
  pageId: string;
  loadEpoch: number;
  topologyRevision: number;
  sheets: readonly {
    id: string;
    name: string;
    results: readonly (readonly [string, T])[];
    formulaKeys: readonly string[];
  }[];
}

/**
 * ``worksheetCalculationLedgerThroughLineEdit``, with what the edit dropped,
 * and with ``restore``: results an earlier, inverse edit dropped (an Undo or a
 * Redo of a deletion) put back at their keys once the cells are back. A
 * restore taken from another ledger (another load, a topology change since,
 * a renamed sheet) is ignored.
 */
export function worksheetCalculationLedgerLineEdit<T>(params: {
  before: WorksheetCalculationAuthority;
  after: WorksheetCalculationAuthority;
  ledger: WorksheetCalculationLedger<T>;
  edits: ReadonlyMap<string, (key: string) => string | null>;
  restore?: WorksheetCalculationLinesDropped<T> | null;
}): { ledger: WorksheetCalculationLedger<T>; dropped: WorksheetCalculationLinesDropped<T> | null } {
  const { before, after, ledger, edits, restore } = params;
  if (
    ledger.pageId !== before.pageId ||
    ledger.loadEpoch !== before.loadEpoch ||
    ledger.topologyRevision !== before.topologyRevision ||
    after.topologyRevision !== before.topologyRevision ||
    ledger.sheets.length !== before.sheets.length ||
    after.sheets.length !== before.sheets.length
  ) {
    return { ledger, dropped: null };
  }
  const moved = new Map<string, { prefix: string; mapKey: (key: string) => string | null }>();
  for (let index = 0; index < ledger.sheets.length; index += 1) {
    const recorded = ledger.sheets[index];
    const mapKey = edits.get(recorded.id);
    if (!mapKey) continue;
    const was = before.sheets[index];
    const now = after.sheets[index];
    if (
      !was ||
      !now ||
      was.id !== recorded.id ||
      now.id !== recorded.id ||
      was.name !== recorded.name ||
      now.name !== recorded.name ||
      was.coordinateRevision !== recorded.coordinateRevision ||
      was.rowCount < recorded.rowCount ||
      was.columnCount < recorded.columnCount
    ) {
      continue;
    }
    moved.set(recorded.id, { prefix: `${recorded.name.toLowerCase()}\u0000`, mapKey });
  }
  if (moved.size === 0) return { ledger, dropped: null };
  const prefixes = [...moved.entries()];
  const sheetOf = (key: string): string | null => {
    for (const [id, { prefix }] of prefixes) if (key.startsWith(prefix)) return id;
    return null;
  };
  const remap = (key: string): string | null => {
    for (const [, { prefix, mapKey }] of prefixes) {
      if (!key.startsWith(prefix)) continue;
      const cell = mapKey(key.slice(prefix.length));
      return cell === null ? null : prefix + cell;
    }
    return key;
  };
  const droppedResults = new Map<string, (readonly [string, T])[]>();
  const droppedFormulas = new Map<string, string[]>();
  const results = new Map<string, T>();
  for (const [key, value] of ledger.results) {
    const next = remap(key);
    if (next !== null) {
      results.set(next, value);
      continue;
    }
    const id = sheetOf(key);
    if (id !== null) droppedResults.set(id, [...(droppedResults.get(id) ?? []), [key, value] as const]);
  }
  // Results an inverse edit took with the cells this edit brings back.
  const restoring = new Map<string, WorksheetCalculationLinesDropped<T>["sheets"][number]>();
  if (
    restore &&
    restore.pageId === ledger.pageId &&
    restore.loadEpoch === ledger.loadEpoch &&
    restore.topologyRevision === ledger.topologyRevision
  ) {
    for (const entry of restore.sheets) {
      const recorded = ledger.sheets.find((candidate) => candidate.id === entry.id);
      if (!recorded || recorded.name !== entry.name || !moved.has(entry.id)) continue;
      restoring.set(entry.id, entry);
      for (const [key, value] of entry.results) if (!results.has(key)) results.set(key, value);
    }
  }
  const sheets = ledger.sheets.map((recorded, index) => {
    if (!moved.has(recorded.id)) return recorded;
    const now = after.sheets[index];
    const kept: string[] = [];
    for (const key of recorded.formulaKeys) {
      const next = remap(key);
      if (next !== null) kept.push(next);
      else droppedFormulas.set(recorded.id, [...(droppedFormulas.get(recorded.id) ?? []), key]);
    }
    const restored = restoring.get(recorded.id);
    if (restored) {
      const present = new Set(kept);
      for (const key of restored.formulaKeys) if (!present.has(key)) kept.push(key);
    }
    return Object.freeze({
      ...recorded,
      rowCount: now.rowCount,
      columnCount: now.columnCount,
      coordinateRevision: now.coordinateRevision,
      formulaKeys: Object.freeze(kept),
    });
  });
  const droppedSheets = ledger.sheets
    .filter((recorded) => droppedResults.has(recorded.id) || droppedFormulas.has(recorded.id))
    .map((recorded) =>
      Object.freeze({
        id: recorded.id,
        name: recorded.name,
        results: Object.freeze(droppedResults.get(recorded.id) ?? []),
        formulaKeys: Object.freeze(droppedFormulas.get(recorded.id) ?? []),
      }),
    );
  return {
    ledger: Object.freeze({ ...ledger, sheets: Object.freeze(sheets), results }),
    dropped:
      droppedSheets.length > 0
        ? Object.freeze({
            pageId: ledger.pageId,
            loadEpoch: ledger.loadEpoch,
            topologyRevision: ledger.topologyRevision,
            sheets: Object.freeze(droppedSheets),
          })
        : null,
  };
}
