/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { worksheetFormulaResultKey } from "./calculation-mode";
import { checkCalculationPivots, MEMBER_PIVOT_LIMITS, type CalculationPivotLimits, type CalculationPivots } from "./calculation-pivots";
import { sameCalculationToken, type BackgroundCalculationDelta, type BackgroundCalculationSheet, type BackgroundCalculationSnapshot } from "./background-calculation-protocol";

function checkedCoordinate(row: number, col: number, sheet: BackgroundCalculationSheet): void {
  if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col) || row < 0 || col < 0 || row >= sheet.rowCount || col >= sheet.columnCount)
    throw new Error("Invalid calculation coordinate.");
}

export interface OwnedCalculationSnapshot {
  /** The snapshot without its cells (they are in `cells`) and PivotTables. */
  readonly meta: BackgroundCalculationSnapshot;
  /** Raw cells by sheet id, "row:col" keys. */
  readonly cells: Map<string, Map<string, string>>;
  /** Formula cells by result key. */
  readonly formulas: Map<string, { sheet: string; row: number; col: number; formula: string }>;
  readonly pivots: CalculationPivots | null;
}

/** Validates and copies a snapshot (throws on any malformed part). */
/** A cell key's `[row, col]` when it is `/^\d+:\d+$/`, else null; read by
 * hand (a regular expression and a split a cell were a fifth of loading a
 * 400,000-formula snapshot). */
function cellKeyParts(key: string): [number, number] | null {
  const colon = key.indexOf(":");
  if (colon <= 0 || colon === key.length - 1) return null;
  let row = 0;
  for (let index = 0; index < colon; index += 1) {
    const digit = key.charCodeAt(index) - 48;
    if (digit < 0 || digit > 9) return null;
    row = row * 10 + digit;
  }
  let col = 0;
  for (let index = colon + 1; index < key.length; index += 1) {
    const digit = key.charCodeAt(index) - 48;
    if (digit < 0 || digit > 9) return null;
    col = col * 10 + digit;
  }
  return [row, col];
}

export function ownCalculationSnapshot(snapshot: BackgroundCalculationSnapshot, pivotLimits: CalculationPivotLimits = MEMBER_PIVOT_LIMITS): OwnedCalculationSnapshot {
  // The cells are copied into maps below (strings, which nothing can change):
  // cloning them first copied every cell twice (a third of the load).
  const sources = snapshot.sheets;
  const { pivots: snapshotPivots, ...owned }: BackgroundCalculationSnapshot = structuredClone({
    ...snapshot, sheets: sources.map(({ cells: _cells, ...sheet }) => ({ ...sheet, cells: [] as BackgroundCalculationSheet["cells"] })),
  });
  const cells = new Map<string, Map<string, string>>();
  const formulas = new Map<string, { sheet: string; row: number; col: number; formula: string }>();
  const names = new Set<string>();
  owned.sheets.forEach((sheet, index) => {
    if (cells.has(sheet.id) || names.has(sheet.name.toLowerCase()) || !sheet.name || !Number.isSafeInteger(sheet.rowCount) || !Number.isSafeInteger(sheet.columnCount) || sheet.rowCount < 1 || sheet.columnCount < 1) throw new Error("Invalid calculation sheet.");
    names.add(sheet.name.toLowerCase());
    const raw = new Map<string, string>();
    for (const [key, value] of sources[index].cells) {
      const parts = typeof key === "string" ? cellKeyParts(key) : null;
      if (!parts || typeof value !== "string") throw new Error("Invalid calculation cell.");
      const [row, col] = parts;
      checkedCoordinate(row, col, sheet);
      // Spelled as the number it is (no leading zero): within the sheet, the
      // only other spelling a parse would accept.
      if ((key.charCodeAt(0) === 48 && key.charCodeAt(1) !== 58) || (key.charCodeAt(key.indexOf(":") + 1) === 48 && key.length - key.indexOf(":") > 2) || raw.has(key)) {
        throw new Error("Invalid calculation cell key.");
      }
      raw.set(key, value);
      if (value.trim().startsWith("=")) formulas.set(worksheetFormulaResultKey(sheet.name, row, col), { sheet: sheet.name, row, col, formula: value });
    }
    cells.set(sheet.id, raw);
  });
  const pivots = snapshotPivots ? checkCalculationPivots(snapshotPivots, owned.sheets, pivotLimits) : null;
  // Cells are kept in one map, not duplicated in the metadata snapshot.
  return { meta: { ...owned, sheets: owned.sheets.map((sheet) => ({ ...sheet, cells: [] })) }, cells, formulas, pivots };
}

/**
 * Checks a delta against the snapshot it must follow (throws: reset
 * required) and returns its owned context replacements by sheet id.
 */
export function checkCalculationDelta(snapshot: BackgroundCalculationSnapshot | null, delta: BackgroundCalculationDelta): Map<string, NonNullable<BackgroundCalculationDelta["contexts"]>[number]["context"]> {
  if (!snapshot || !sameCalculationToken(snapshot.token, delta.baseToken) || delta.token.pageId !== snapshot.token.pageId || delta.token.loadEpoch !== snapshot.token.loadEpoch || delta.token.topologyRevision !== snapshot.token.topologyRevision || delta.token.generation <= snapshot.token.generation || delta.token.workbookRevision < snapshot.token.workbookRevision) throw new Error("Stale or structural calculation delta; reset required.");
  const byId = new Map(snapshot.sheets.map((sheet) => [sheet.id, sheet]));
  for (const cell of delta.cells) {
    const sheet = byId.get(cell.sheetId);
    if (!sheet || (cell.raw !== null && typeof cell.raw !== "string")) throw new Error("Invalid calculation delta.");
    checkedCoordinate(cell.row, cell.col, sheet);
  }
  for (const replacement of delta.contexts ?? []) {
    if (!byId.has(replacement.sheetId)) throw new Error("Unsupported calculation context.");
  }
  return new Map(structuredClone(delta.contexts ?? []).map((entry) => [entry.sheetId, entry.context]));
}

/** The metadata snapshot after a checked delta (token and contexts). */
export function advanceCalculationSnapshot(snapshot: BackgroundCalculationSnapshot, delta: BackgroundCalculationDelta, contexts: ReturnType<typeof checkCalculationDelta>): BackgroundCalculationSnapshot {
  return { ...snapshot, token: { ...delta.token }, sheets: snapshot.sheets.map((sheet) => contexts.has(sheet.id) ? { ...sheet, context: contexts.get(sheet.id) } : sheet) };
}

export class BackgroundCalculationMirror {
  #snapshot: BackgroundCalculationSnapshot | null = null;
  #cells = new Map<string, Map<string, string>>();
  #pivots: CalculationPivots | null = null;

  reset(snapshot: BackgroundCalculationSnapshot): void {
    const owned = ownCalculationSnapshot(snapshot);
    this.#snapshot = owned.meta;
    this.#cells = owned.cells;
    this.#pivots = owned.pivots;
  }

  applyDelta(delta: BackgroundCalculationDelta): void {
    const snapshot = this.#snapshot;
    const contexts = checkCalculationDelta(snapshot, delta);
    for (const cell of delta.cells) {
      const raw = this.#cells.get(cell.sheetId)!;
      const key = `${cell.row}:${cell.col}`;
      if (cell.raw === null) raw.delete(key); else raw.set(key, cell.raw);
    }
    this.#snapshot = advanceCalculationSnapshot(snapshot!, delta, contexts);
  }

  /** Holds these PivotTables instead (null: none), validated against the limits. */
  replacePivots(value: CalculationPivots | null): void {
    const snapshot = this.#snapshot;
    if (!snapshot) throw new Error("Calculation snapshot not initialized.");
    this.#pivots = value === null ? null : checkCalculationPivots(structuredClone(value), snapshot.sheets, MEMBER_PIVOT_LIMITS);
  }

  /** The `id` of the PivotTables held. */
  get pivotsId(): string | undefined { return this.#pivots?.id; }

  exportSnapshot(): BackgroundCalculationSnapshot {
    if (!this.#snapshot) throw new Error("Calculation snapshot not initialized.");
    return { ...this.#snapshot, sheets: this.#snapshot.sheets.map((sheet) => ({ ...sheet, cells: [...this.#cells.get(sheet.id)!] })),
      ...(this.#pivots ? { pivots: this.#pivots } : {}) };
  }

  clear(): void { this.#snapshot = null; this.#cells.clear(); this.#pivots = null; }
}
