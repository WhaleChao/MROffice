/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { KeptFormulaResultEntry } from "./kept-formula-results";
import type { BackgroundCalculationPlanRequest, BackgroundCalculationPlan } from "./background-calculation-plans";
import type { WorksheetCalculationToken, WorksheetFormulaCoordinate } from "./calculation-mode";
import type { ColumnWidthReading, EngineValue, TableDefinition, WorkbookFormulaResultSeed } from "./formula-engine";
import type { NamedRange } from "./named-ranges";
import type { CalculationPivots } from "./calculation-pivots";
import type { WhatIfTableDefinition } from "./what-if-table-definition";
import type { CalculationExternalDatum } from "./external-data";

/** Session-only read model. Never persist this protocol or accept worker model writes. */
export interface BackgroundCalculationSheet {
  readonly id: string;
  readonly name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  /** null preserves an explicitly absent host extent; omitted is the legacy sheet extent. */
  readonly usedRange?: { readonly rowCount: number; readonly columnCount: number } | null;
  /** Sparse, merge-masked canonical raw values; never formatted display caches. */
  readonly cells: ReadonlyArray<readonly [string, string]>;
  readonly context?: BackgroundCalculationSheetContext;
}
export interface BackgroundCalculationSheetContext {
  readonly namedRanges?: readonly NamedRange[];
  readonly tables?: readonly TableDefinition[];
  /** These must be derived from the SAME source revision as cells. */
  readonly filterHiddenRows?: readonly number[];
  readonly manuallyHiddenRows?: readonly number[];
  /**
   * "row:col" keys of the cells the host types as TEXT (an apostrophe-prefixed
   * entry, or one typed into a Text cell), so the worker reads them as strings
   * exactly as the main thread does. Omitted when the sheet has none.
   */
  readonly textCells?: readonly string[];
  /**
   * "row:col" keys a merge covers. The worker's cells are merge-masked, so
   * these are the blanks a dynamic array may NOT spill into, exactly as the
   * grid's overlay refuses them; omitted when the sheet has no merges.
   */
  readonly coveredCells?: readonly string[];
  /**
   * The literal cells holding a picture placed in the cell, as "row:col" key
   * and alternative text, so the worker reads each one as the picture value
   * (TYPE 128, not an error) exactly as the main thread does. The picture's
   * bytes never cross; omitted when the sheet has none.
   */
  readonly pictureCells?: ReadonlyArray<readonly [string, string]>;
  /**
   * The text cells holding furigana, as "row:col" key and the reading
   * PHONETIC returns for them; omitted when the sheet has none.
   */
  readonly phoneticCells?: ReadonlyArray<readonly [string, string]>;
  /**
   * The formula cells over external data that keep the file's last value
   * (kept-formula-results): "row:col", the formula it belongs to, its kind
   * and value as text; omitted when the sheet has none.
   */
  readonly keptResults?: ReadonlyArray<KeptFormulaResultEntry>;
  /**
   * The "row:col" keys of the formula cells the sheet's protection hides
   * (formatted Hidden on a protected sheet), whose FORMULATEXT is #N/A;
   * omitted when there are none.
   */
  readonly formulaHiddenCells?: readonly string[];
  /**
   * The legacy array formulas (Ctrl+Shift+Enter, `arrayRef`) as anchor
   * "row:col" key, range height and width, so the worker lays each result
   * over its own range as the main thread does; omitted when there are none.
   */
  readonly legacyArrays?: ReadonlyArray<readonly [string, number, number]>;
  /**
   * The sheet's What-If Data Tables (what-if-table-definition): their
   * results are `=TABLE(...)` cells of the snapshot, recalculated from these
   * definitions; omitted when the sheet has none.
   */
  readonly whatIfTables?: readonly WhatIfTableDefinition[];
  /** The sheet's column widths in characters, for CELL("width"); omitted
   * when the host gives none. */
  readonly columnWidths?: ColumnWidthReading;
}
export interface BackgroundCalculationSnapshot {
  readonly token: WorksheetCalculationToken;
  readonly sheets: readonly BackgroundCalculationSheet[];
  readonly workbookName?: string;
  /** The folder of the workbook's location, for CELL("filename") (X-194). */
  readonly workbookFolder?: string;
  /** The workbook reads serials in the 1904 date system. A change of system
   * is a topology change: the host resets with a new snapshot. */
  readonly date1904?: boolean;
  /** The non-compatible 1900 date base (Strict `dateCompatibility="0"`,
   * X-236), beside `date1904` so an engine host that predates it reads the
   * 1900 system rather than mistaking it for 1904. */
  readonly dateSystem?: "1900-noncompat";
  /** Enable iterative calculation (calcPr@iterate) and its limits. A change
   * is a topology change: the host resets with a new snapshot. */
  readonly iteration?: {
    readonly maxIterations: number;
    readonly maxChange: number;
    /** The formula cells' current numbers ([sheet name, row, col, value]):
     * the first calculation's iterations start from them (the file's
     * cached values, or the last save's), as Excel's do. */
    readonly seeds?: ReadonlyArray<readonly [string, number, number, number]>;
  };
  /** The PivotTables the runtime lays over the sheets (calculation-pivots):
   * a session re-opened from the client's mirror keeps them. */
  readonly pivots?: CalculationPivots;
}
export interface BackgroundCalculationDelta {
  readonly baseToken: WorksheetCalculationToken;
  readonly token: WorksheetCalculationToken;
  readonly cells: ReadonlyArray<{ sheetId: string; row: number; col: number; raw: string | null }>;
  /** Complete replacement context for affected sheets. Invalidates warm results. */
  readonly contexts?: ReadonlyArray<{ sheetId: string; context: BackgroundCalculationSheetContext }>;
}
/**
 * A conditional-format formula (a formula rule, a formula stop) evaluated for
 * every cell of its range relative to its anchor, as the rule paints (C2):
 * the session answers it with the calculation, in the same round trip.
 */
export interface CalculationRuleRequest {
  readonly sheetId: string;
  readonly formula: string;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly range: { readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };
}
/** Most rule cells one calculation evaluates (larger rules are left unanswered). */
export const MAX_CALCULATION_RULE_CELLS = 20_000;
/** Most rule formulas one calculation carries (the session refuses more:
 * engine-host/engines/sheet-calc.ts MAX_RULES). */
export const MAX_CALCULATION_RULES = 256;
/** A rule's identity among a calculation's rules. */
export const calculationRuleKey = (rule: Pick<CalculationRuleRequest, "sheetId" | "formula" | "anchorRow" | "anchorCol">): string =>
  `${rule.sheetId}\u0000${rule.formula}\u0000${rule.anchorRow}:${rule.anchorCol}`;
export interface BackgroundCalculationRequest {
  readonly token: WorksheetCalculationToken;
  readonly coordinates: readonly WorksheetFormulaCoordinate[];
  readonly formulaResultSeed?: WorkbookFormulaResultSeed;
  readonly maxComputedCells?: number;
  readonly rules?: readonly CalculationRuleRequest[];
  /** Replace the PivotTables held (null: hold none; omitted: keep them),
   * applied before calculating; the session drops a replacement whose
   * `basis` its workbook moved past. */
  readonly pivots?: CalculationPivots | null;
  /** The IMAGE sources the browser was refused (it is the one that fetches
   * pictures), each with Excel's error; replaces the ones held, applied
   * before calculating. Omitted: keep them. A source not listed answers its
   * picture value. */
  readonly webImages?: readonly CalculationWebImage[];
  /** What WEBSERVICE / IMPORTRANGE read, as the browser fetched it
   * (external-data); replaces the data held, applied before calculating.
   * Omitted: keep it. A call whose data is not held reads #BUSY! and is
   * named in the answer's `externalRequests`. */
  readonly externalData?: readonly CalculationExternalDatum[];
}
/** One refused IMAGE source and the error it reads as. */
export type CalculationWebImage = readonly [source: string, status: "#CONNECT!" | "#BLOCKED!" | "#VALUE!"];
export const MAX_CALCULATION_WEB_IMAGES = 256;
export const MAX_CALCULATION_WEB_IMAGE_SOURCE = 2_048;
export interface BackgroundCalculationResult {
  readonly token: WorksheetCalculationToken;
  readonly results: ReadonlyMap<string, EngineValue>;
  readonly execution: "worker" | "cooperative";
  /** Per rule key: the value at each "row:col" of its range that is not FALSE
   * (absent cell = FALSE). Only from the server's session. */
  readonly ruleValues?: ReadonlyMap<string, ReadonlyMap<string, EngineValue>>;
  /** The server session's answer number (a pivot replacement's basis). */
  readonly seq?: number;
  /** The `id` of the PivotTables this answer was calculated with. */
  readonly pivotsId?: string;
  /** WEBSERVICE / IMPORTRANGE calls read without their data (keys of
   * external-data): the browser fetches them and calculates again. */
  readonly externalRequests?: readonly string[];
}
export type BackgroundCalculationMessage =
  | { kind: "reset"; snapshot: BackgroundCalculationSnapshot }
  | { kind: "delta"; delta: BackgroundCalculationDelta }
  | { kind: "calculate"; requestId: number; request: BackgroundCalculationRequest }
  | { kind: "plan"; requestId: number; request: BackgroundCalculationPlanRequest }
  | { kind: "cancel"; requestId: number };
export type BackgroundCalculationResponse =
  | { kind: "plan-result"; requestId: number; token: WorksheetCalculationToken; plan: BackgroundCalculationPlan }
  | { kind: "result"; requestId: number; token: WorksheetCalculationToken; results: ReadonlyMap<string, EngineValue>;
      ruleValues?: ReadonlyMap<string, ReadonlyMap<string, EngineValue>>; seq?: number; pivotsId?: string; externalRequests?: readonly string[] }
  | { kind: "error"; requestId: number | null; message: string; reason?: CalculationUnavailableReason };

export function sameCalculationToken(a: WorksheetCalculationToken, b: WorksheetCalculationToken): boolean {
  return a.pageId === b.pageId && a.loadEpoch === b.loadEpoch && a.generation === b.generation &&
    a.workbookRevision === b.workbookRevision && a.topologyRevision === b.topologyRevision &&
    a.targetSheetId === b.targetSheetId;
}

export class BackgroundCalculationCancelledError extends Error {
  constructor() { super("Calculation cancelled or superseded."); this.name = "BackgroundCalculationCancelledError"; }
}
export async function yieldCalculationTask(): Promise<void> {
  const scheduler = (globalThis as typeof globalThis & { scheduler?: { yield?: () => Promise<void> } }).scheduler;
  if (typeof scheduler?.yield === "function") await scheduler.yield();
  else await new Promise<void>((resolve) => setTimeout(resolve, 0));
}






export class BackgroundCalculationUnavailableError extends Error {
  readonly reason: CalculationUnavailableReason;
  constructor(reason: CalculationUnavailableReason = "unavailable", message = "Calculation unavailable.") {
    super(message); this.name = "BackgroundCalculationUnavailableError"; this.reason = reason;
  }
}

/**
 * Why the calculation could not answer (C4), as the status bar names it:
 * no connection; the calculation pool busy or this account's rate used up
 * (it comes back by itself); the workbook over the calculation's size or
 * computation budget (an edit that shrinks it, or a smaller recalculation,
 * is needed); any other failure of the service.
 */
export type CalculationUnavailableReason = "offline" | "busy" | "budget" | "unavailable";
