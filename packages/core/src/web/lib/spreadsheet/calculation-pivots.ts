/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BackgroundCalculationRequest, BackgroundCalculationResult } from "./background-calculation-protocol";
import type { EngineValue, WorkbookSheetInput } from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { coveredCells, normalizeMerges } from "./merge-cells";
import type { PivotConfig, PivotRange } from "./pivot";
import { pivotRangesContain, type PivotOutputLabels, type PivotOutputPlan } from "./pivot-output";
import {
  pivotPlanReady,
  pivotPlanRequests,
  planPivotRequest,
  type PivotOverlaySheet,
  type PivotPlanRequest,
} from "./pivot-overlay";
import { createValueStoreWorkbook } from "./value-store";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

/** A PivotTable's labels as text: a template's `\u0001` (and `\u0002`) are
 * its arguments, so the session writes the viewer's language. */
export interface CalculationPivotLabels {
  readonly grandTotal: string;
  readonly subtotal: string;
  readonly valueTotal: string;
  readonly itemValue: string;
  readonly rowLabels: string;
  readonly columnLabels: string;
  readonly values: string;
  readonly allItems: string;
  readonly multipleItems: string;
  readonly blank: string;
  readonly total: string;
}

/** Every placed PivotTable of a workbook, as the session holds them. */
export interface CalculationPivots {
  /** Identity of the content (sheets, labels, collapsed), never of `basis`. */
  readonly id: string;
  /** Per sheet the PivotTables on it, in slot order (the first one's
   * placement answers GETPIVOTDATA, as on the grid). */
  readonly sheets: ReadonlyArray<{
    readonly sheetId: string;
    readonly requests: readonly PivotPlanRequest[];
    /** "row:col" of the cells a merge covers that still hold a value or a
     * formula: occupied for a report, as on the grid, though the session's
     * merge-masked cells do not show them. Omitted when there are none. */
    readonly maskedOccupied?: readonly string[];
  }>;
  readonly labels: CalculationPivotLabels;
  /** Drill paths shown collapsed (view state). */
  readonly collapsed: readonly string[];
  /** The answer a replacement was derived from; the session drops a
   * replacement whose workbook revision or answer moved on. */
  readonly basis?: { readonly workbookRevision: number; readonly answerSeq: number };
}

export interface CalculationPivotLimits {
  /** PivotTables per workbook. */
  readonly pivotsMax: number;
  /** One PivotTable's computation rows x columns. */
  readonly pivotMatrixCells: number;
  /** All PivotTables' computation cells together. */
  readonly pivotsMatrixCells: number;
  /** The whole replacement as JSON text. */
  readonly pivotsChars: number;
}
/** Members (the server's ceiling, whatever a caller is stamped with). */
export const MEMBER_PIVOT_LIMITS: CalculationPivotLimits = {
  pivotsMax: 64, pivotMatrixCells: 100_000, pivotsMatrixCells: 200_000, pivotsChars: 8_000_000,
};
/** Guests (and a caller the route stamped nothing for). */
export const GUEST_PIVOT_LIMITS: CalculationPivotLimits = {
  pivotsMax: 8, pivotMatrixCells: 10_000, pivotsMatrixCells: 20_000, pivotsChars: 1_000_000,
};

/** The limits a route stamped (`callerLimits`), within the members' ceiling;
 * guests' without a stamp. */
export function calculationPivotLimitsOf(callerLimits: unknown): CalculationPivotLimits {
  const record = callerLimits && typeof callerLimits === "object" ? callerLimits as Record<string, unknown> : null;
  if (!record || !("pivotsMax" in record)) return GUEST_PIVOT_LIMITS;
  const pick = (key: keyof CalculationPivotLimits) => {
    const asked = record[key];
    return typeof asked === "number" && Number.isSafeInteger(asked) && asked > 0
      ? Math.min(asked, MEMBER_PIVOT_LIMITS[key]) : GUEST_PIVOT_LIMITS[key];
  };
  return { pivotsMax: pick("pivotsMax"), pivotMatrixCells: pick("pivotMatrixCells"),
    pivotsMatrixCells: pick("pivotsMatrixCells"), pivotsChars: pick("pivotsChars") };
}

export class CalculationPivotsError extends Error {
  constructor(readonly code: "invalid_pivots" | "pivots_too_large", readonly details: Record<string, unknown> = {}) {
    super(code);
    this.name = "CalculationPivotsError";
  }
}

const ITEM = "\u0001";
const SECOND = "\u0002";
const fill = (template: string, ...args: string[]) =>
  args.reduce((text, arg, index) => text.split(index === 0 ? ITEM : SECOND).join(arg), template);

export function serializePivotLabels(labels: PivotOutputLabels): CalculationPivotLabels {
  return {
    grandTotal: labels.grandTotal, subtotal: labels.subtotal(ITEM), valueTotal: labels.valueTotal(ITEM),
    itemValue: labels.itemValue(ITEM, SECOND), rowLabels: labels.rowLabels, columnLabels: labels.columnLabels,
    values: labels.values, allItems: labels.allItems, multipleItems: labels.multipleItems, blank: labels.blank,
    total: labels.total,
  };
}

/** The view a PivotTable report is laid out in (labels, collapsed items), as
 * a key: results read over the reports depend on it. */
export function pivotViewKey(labels: PivotOutputLabels, collapsed: ReadonlySet<string>): string {
  return contentId(JSON.stringify([serializePivotLabels(labels), [...collapsed].sort()]));
}

export function pivotLabelsOf(labels: CalculationPivotLabels): PivotOutputLabels {
  return {
    ...labels,
    subtotal: (item) => fill(labels.subtotal, item),
    valueTotal: (caption) => fill(labels.valueTotal, caption),
    itemValue: (item, caption) => fill(labels.itemValue, item, caption),
  };
}

/** A 64-bit FNV-1a of the text, hex (identity, not security). */
function contentId(text: string): string {
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ 0x5bd1e995;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    a = Math.imul(a ^ code, 0x01000193);
    b = Math.imul(b ^ code, 0x5bd1e995) ^ (b >>> 15);
  }
  return (a >>> 0).toString(16).padStart(8, "0") + (b >>> 0).toString(16).padStart(8, "0");
}

/**
 * The workbook's placed PivotTables as the session should hold them, read
 * from a calculation answer (the source cells' values as the grid shows
 * them): null when the workbook has none.
 */
export function calculationPivotsFromResults(
  sheets: readonly PivotOverlaySheet[],
  inputs: readonly WorkbookSheetInput[],
  results: ReadonlyMap<string, EngineValue>,
  options: { date1904?: ExcelDateSystemArg; labels: PivotOutputLabels; collapsed: ReadonlySet<string> },
): CalculationPivots | null {
  const bootstrap = createValueStoreWorkbook(inputs, { results, date1904: options.date1904 ?? false });
  const { requests } = pivotPlanRequests(sheets, bootstrap.evaluatorFor);
  const bySheet = new Map<string, PivotPlanRequest[]>();
  for (const request of requests) {
    if (request.chartOnly) continue;
    const list = bySheet.get(request.sheetId) ?? [];
    list.push(request);
    bySheet.set(request.sheetId, list);
  }
  if (bySheet.size === 0) return null;
  const byId = new Map(sheets.map((sheet) => [sheet.id, sheet]));
  const body = {
    sheets: [...bySheet].map(([sheetId, list]) => {
      const maskedOccupied = maskedOccupiedCells(byId.get(sheetId)!);
      return { sheetId, requests: list, ...(maskedOccupied.length ? { maskedOccupied } : {}) };
    }),
    labels: serializePivotLabels(options.labels),
    collapsed: [...options.collapsed].sort(),
  };
  return { id: contentId(JSON.stringify(body)), ...body };
}

/** The merge-covered cells of a sheet that hold a value or a formula (the
 * grid's occupancy: pivot-overlay's sheetCellOccupied). */
function maskedOccupiedCells(sheet: PivotOverlaySheet): string[] {
  const out: string[] = [];
  for (const key of coveredCells(normalizeMerges(sheet.merges))) {
    const cell = sheet.cells[key];
    if (cell && ((cell.value ?? "") !== "" || (cell.formula ?? "") !== "")) out.push(key);
  }
  return out.sort();
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const gridInt = (value: unknown, max: number) => Number.isSafeInteger(value) && (value as number) >= 0 && (value as number) < max;

function checkedRange(value: unknown): PivotRange {
  if (!isRecord(value) || !gridInt(value.r1, MAX_GRID_ROWS) || !gridInt(value.r2, MAX_GRID_ROWS) ||
      !gridInt(value.c1, MAX_GRID_COLUMNS) || !gridInt(value.c2, MAX_GRID_COLUMNS) ||
      (value.r1 as number) > (value.r2 as number) || (value.c1 as number) > (value.c2 as number)) {
    throw new CalculationPivotsError("invalid_pivots", { field: "footprint" });
  }
  return { r1: value.r1 as number, c1: value.c1 as number, r2: value.r2 as number, c2: value.c2 as number };
}

const rangesOverlap = (a: PivotRange, b: PivotRange) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

/**
 * A replacement as the session may hold it, or a closed refusal: every field
 * typed, every PivotTable on a sheet the session holds, footprints inside
 * the grid and apart from each other on a sheet, sizes within `limits`.
 */
export function checkCalculationPivots(
  value: unknown,
  sheets: ReadonlyArray<{ readonly id: string }>,
  limits: CalculationPivotLimits,
): CalculationPivots {
  if (!isRecord(value) || typeof value.id !== "string" || value.id.length > 64 || !Array.isArray(value.sheets) ||
      !isRecord(value.labels) || !Array.isArray(value.collapsed)) {
    throw new CalculationPivotsError("invalid_pivots", { field: "pivots" });
  }
  const chars = JSON.stringify(value).length;
  if (chars > limits.pivotsChars) throw new CalculationPivotsError("pivots_too_large", { maxChars: limits.pivotsChars });
  const labelKeys: Array<keyof CalculationPivotLabels> = ["grandTotal", "subtotal", "valueTotal", "itemValue", "rowLabels",
    "columnLabels", "values", "allItems", "multipleItems", "blank", "total"];
  const labels = value.labels;
  if (!labelKeys.every((key) => typeof labels[key] === "string" && (labels[key] as string).length <= 256)) {
    throw new CalculationPivotsError("invalid_pivots", { field: "labels" });
  }
  if (!value.collapsed.every((path) => typeof path === "string")) throw new CalculationPivotsError("invalid_pivots", { field: "collapsed" });
  const basis = value.basis;
  if (basis !== undefined && (!isRecord(basis) || !Number.isSafeInteger(basis.workbookRevision) || !Number.isSafeInteger(basis.answerSeq))) {
    throw new CalculationPivotsError("invalid_pivots", { field: "basis" });
  }
  const known = new Set(sheets.map((sheet) => sheet.id));
  const seenSheets = new Set<string>();
  const planKeys = new Set<string>();
  let pivots = 0;
  let totalCells = 0;
  const outSheets: Array<{ sheetId: string; requests: PivotPlanRequest[]; maskedOccupied?: string[] }> = [];
  for (const entry of value.sheets) {
    if (!isRecord(entry) || typeof entry.sheetId !== "string" || !known.has(entry.sheetId) || seenSheets.has(entry.sheetId) ||
        !Array.isArray(entry.requests) || entry.requests.length === 0) {
      throw new CalculationPivotsError("invalid_pivots", { field: "sheets" });
    }
    seenSheets.add(entry.sheetId);
    const masked = entry.maskedOccupied;
    if (masked !== undefined && (!Array.isArray(masked) || masked.length > 100_000 ||
        !masked.every((key) => typeof key === "string" && /^\d{1,7}:\d{1,5}$/.test(key)))) {
      throw new CalculationPivotsError("invalid_pivots", { field: "maskedOccupied" });
    }
    const requests: PivotPlanRequest[] = [];
    const footprints: PivotRange[] = [];
    for (const request of entry.requests) {
      pivots += 1;
      if (pivots > limits.pivotsMax) throw new CalculationPivotsError("pivots_too_large", { maxPivots: limits.pivotsMax });
      if (!isRecord(request) || typeof request.planKey !== "string" || request.planKey.length > 512 || planKeys.has(request.planKey) ||
          request.sheetId !== entry.sheetId || request.chartOnly !== false || !isRecord(request.config) ||
          !Array.isArray(request.footprint) || request.footprint.length > 2 || !Array.isArray(request.matrix) ||
          !Number.isSafeInteger(request.rowCount) || (request.rowCount as number) < 1 || (request.rowCount as number) > MAX_GRID_ROWS ||
          !Number.isSafeInteger(request.columnCount) || (request.columnCount as number) < 1 || (request.columnCount as number) > MAX_GRID_COLUMNS) {
        throw new CalculationPivotsError("invalid_pivots", { field: "request" });
      }
      const config = request.config;
      if (!Array.isArray(config.rows) || !Array.isArray(config.columns) || !Array.isArray(config.values) || !Array.isArray(config.filters)) {
        throw new CalculationPivotsError("invalid_pivots", { field: "config" });
      }
      planKeys.add(request.planKey);
      let width = 0;
      for (const row of request.matrix) {
        if (!Array.isArray(row) || !row.every((cell) => typeof cell === "string")) throw new CalculationPivotsError("invalid_pivots", { field: "matrix" });
        width = Math.max(width, row.length);
      }
      const cells = request.matrix.length * width;
      if (cells > limits.pivotMatrixCells) throw new CalculationPivotsError("pivots_too_large", { maxMatrixCells: limits.pivotMatrixCells });
      totalCells += cells;
      if (totalCells > limits.pivotsMatrixCells) throw new CalculationPivotsError("pivots_too_large", { maxTotalMatrixCells: limits.pivotsMatrixCells });
      const footprint = request.footprint.map(checkedRange);
      // An imported PivotTable's report never overlaps another's on a sheet.
      if (footprint.some((range) => footprints.some((other) => rangesOverlap(range, other)))) {
        throw new CalculationPivotsError("invalid_pivots", { field: "overlap" });
      }
      footprints.push(...footprint);
      requests.push({ planKey: request.planKey, sheetId: entry.sheetId, config: config as unknown as PivotConfig,
        matrix: request.matrix as string[][], footprint, rowCount: request.rowCount as number,
        columnCount: request.columnCount as number, chartOnly: false });
    }
    outSheets.push({ sheetId: entry.sheetId, requests, ...(masked?.length ? { maskedOccupied: masked as string[] } : {}) });
  }
  return { id: value.id, sheets: outSheets, labels: labels as unknown as CalculationPivotLabels,
    collapsed: value.collapsed as string[], ...(basis ? { basis: basis as unknown as CalculationPivots["basis"] } : {}) };
}

/** A PivotTable report the session lays over its sheet: placed and free. */
export interface CalculationPivotPlan {
  readonly plan: PivotOutputPlan;
  readonly footprint: readonly PivotRange[];
}

/**
 * Each sheet's placed reports, planned exactly as the grid plans them;
 * `occupied(sheetId)` reads the session's raw cells. A request the planner
 * cannot read is simply not placed.
 */
export function planCalculationPivots(
  pivots: CalculationPivots,
  occupied: (sheetId: string) => (row: number, col: number) => boolean,
): Map<string, CalculationPivotPlan[]> {
  const labels = pivotLabelsOf(pivots.labels);
  const collapsed = new Set(pivots.collapsed);
  const out = new Map<string, CalculationPivotPlan[]>();
  for (const { sheetId, requests, maskedOccupied } of pivots.sheets) {
    const raw = occupied(sheetId);
    const masked = new Set(maskedOccupied);
    const isOccupied = masked.size ? (row: number, col: number) => raw(row, col) || masked.has(`${row}:${col}`) : raw;
    const ready: CalculationPivotPlan[] = [];
    for (const request of requests) {
      let plan: PivotOutputPlan;
      try {
        plan = planPivotRequest(request, { labels, collapsed, occupied: isOccupied });
      } catch {
        continue;
      }
      if (pivotPlanReady(plan)) ready.push({ plan, footprint: request.footprint });
    }
    if (ready.length > 0) out.set(sheetId, ready);
  }
  return out;
}

/** What the overlay reads at a cell (undefined: the sheet's own cell). */
export function calculationPivotRead(plans: readonly CalculationPivotPlan[], row: number, col: number): string | null | undefined {
  for (const { plan } of plans) {
    const cell = plan.cells.get(`${row}:${col}`);
    if (cell) return cell.value === null || cell.value === "" ? null : String(cell.value);
  }
  return plans.some(({ footprint }) => pivotRangesContain(footprint, row, col)) ? null : undefined;
}


/** Pivot replacements one calculation may send before it stops as stale
 * (pivot option C, rule 2): a report whose source reads another report's
 * output can change every round. */
export const MAX_PIVOT_REPLACEMENTS = 2;

/**
 * Settles a calculation answer with the PivotTables read from it: while the
 * reports read from the answer are not the ones it was calculated with
 * (`pivotsId`), hand the session those, naming the answer they were read
 * from, and calculate again -- at most MAX_PIVOT_REPLACEMENTS times; a third
 * change answers `pivotStale`. Null from `isCurrent()`: the caller moved on.
 */
export async function settleCalculationPivots(
  client: { calculate(request: BackgroundCalculationRequest): Promise<BackgroundCalculationResult> },
  request: BackgroundCalculationRequest,
  first: BackgroundCalculationResult,
  derive: (results: ReadonlyMap<string, EngineValue>) => CalculationPivots | null,
  isCurrent: () => boolean = () => true,
): Promise<{ answer: BackgroundCalculationResult; pivotStale: boolean } | null> {
  let answer = first;
  for (let round = 0; ; round += 1) {
    if (!isCurrent()) return null;
    const pivots = derive(answer.results);
    if ((pivots?.id ?? null) === (answer.pivotsId ?? null)) return { answer, pivotStale: false };
    if (round === MAX_PIVOT_REPLACEMENTS) return { answer, pivotStale: true };
    answer = await client.calculate({ ...request, pivots: pivots
      ? { ...pivots, basis: { workbookRevision: request.token.workbookRevision, answerSeq: answer.seq ?? -1 } } : null });
  }
}
