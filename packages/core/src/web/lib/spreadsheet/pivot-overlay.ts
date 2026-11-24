/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { buildSpillOverlay, resolveCellDisplay } from "./display-cache";
import type { CellMap } from "./cell-buttons";
import type { SpreadsheetPivotChartSource } from "./chart";
import type { SheetEvaluator, WorkbookSheetInput } from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { coveredCells, normalizeMerges } from "./merge-cells";
import {
  normalizePivotConfig,
  normalizePivotRange,
  pivotComputationMatrix,
  pivotIsScenarioReport,
  pivotLiveSourceRange,
  pivotSourceRowsRead,
  pivotPlanKey,
  pivotReadsLiveCells,
  sheetPivots,
  type PivotConfig,
  type PivotMatrix,
  type PivotRange,
} from "./pivot";
import { resolvePivotChartPivot } from "./pivot-chart";
import {
  nativePivotPageBlock,
  pivotRangesContain,
  planPivotOutput,
  type PivotChartReport,
  type PivotOutputLabels,
  type PivotOutputPlan,
} from "./pivot-output";

/** What a sheet must carry for its PivotTables to be planned. */
export interface PivotOverlaySheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: CellMap;
  merges?: unknown;
  dataTables?: Parameters<typeof pivotLiveSourceRange>[1];
  charts?: readonly unknown[];
  objects?: readonly unknown[];
  pivot?: PivotConfig;
  morePivots?: readonly PivotConfig[];
}

/** Whether every cell of rows `range` of a sheet is in the model. A large
 * tiled sheet holds only the tiles around the screen (lib/spreadsheet/
 * tiled-sheets), and a sheet that is not shown often none at all. */
export type PivotSourceReadable = (sheetId: string, range: PivotRange) => boolean;

export interface PivotPlanReadOptions {
  /** Omitted: every sheet is whole. A report computed from its source's live
   * cells whose source rows are not all in the model is HELD: not planned,
   * not drawn and not read by formulas, rather than totalled over whatever
   * rows happen to be in (an imported report's saved cells then show as they
   * were saved). A report computed from its cache snapshot is not held. */
  readonly sourceReadable?: PivotSourceReadable;
}

/** A PivotTable report that reads its source's live cells, and the rows it
 * reads: what must be in the model before it can be computed. */
export interface PivotLiveSourceRead {
  readonly planKey: string;
  /** The sheet the PivotTable sits on. */
  readonly sheetId: string;
  readonly sourceSheetId: string;
  readonly range: PivotRange;
}

export interface WorkbookPivotPlans {
  /** Each PivotTable's report, by `pivotPlanKey(sheetId, slot)`. */
  readonly plans: Map<string, PivotOutputPlan>;
  /** The cells an imported PivotTable left behind (its report, and the report
   * filter rows above it while they hold only filter captions): the derived
   * output replaces them, so they neither collide with it nor feed formulas. */
  readonly nativeFootprints: Map<string, PivotRange[]>;
  /** What each PivotChart plots, by the plan key of its PivotTable. */
  readonly chartReports: Map<string, PivotChartReport>;
  /** The rows each such report is computed from (a field button's list). */
  readonly chartMatrices: Map<string, PivotMatrix>;
  /** The reports held until their source's rows are read (`sourceReadable`). */
  readonly held: readonly PivotLiveSourceRead[];
}

/**
 * One PivotTable's report as it is to be planned: its recipe, the rows it is
 * computed from (its source as the sheet shows it), what an imported report
 * left behind and the extent the report may grow into. Plain data: the
 * calculation session plans the same report from it (pivot option C).
 */
export interface PivotPlanRequest {
  readonly planKey: string;
  readonly sheetId: string;
  readonly config: PivotConfig;
  readonly matrix: PivotMatrix;
  /** The cells an imported report left behind (none for an authored one). */
  readonly footprint: readonly PivotRange[];
  readonly rowCount: number;
  readonly columnCount: number;
  /** Planned only for the PivotChart bound to it (no on-grid destination). */
  readonly chartOnly: boolean;
}

/**
 * One PivotTable's report from its request; `occupied` says whether a sheet
 * cell holds a value or a formula (the report never covers one; an imported
 * report's own leftovers do not count).
 */
export function planPivotRequest(
  request: PivotPlanRequest,
  options: {
    labels: PivotOutputLabels;
    collapsed: ReadonlySet<string>;
    occupied: (row: number, col: number) => boolean;
  },
): PivotOutputPlan {
  const { config, matrix, footprint } = request;
  return planPivotOutput(config, matrix, {
    rowCount: request.rowCount,
    columnCount: request.columnCount,
    ...(request.chartOnly
      ? {}
      : {
          isOccupied: (row: number, col: number) =>
            !pivotRangesContain(footprint, row, col) && options.occupied(row, col),
        }),
    labels: options.labels,
    collapsed: options.collapsed,
  });
}

/** Whether a sheet cell holds a value or a formula. */
function sheetCellOccupied(sheet: PivotOverlaySheet, row: number, col: number): boolean {
  const existing = sheet.cells[`${row}:${col}`];
  return Boolean(existing && ((existing.value ?? "") !== "" || (existing.formula ?? "") !== ""));
}

export function planWorkbookPivots(
  sheets: readonly PivotOverlaySheet[],
  /** The workbook evaluated without the PivotTables' reports. */
  evaluatorFor: (sheetName: string) => SheetEvaluator,
  options: { labels: PivotOutputLabels; collapsed: ReadonlySet<string> } & PivotPlanReadOptions,
): WorkbookPivotPlans {
  const { requests, chartMatrices, held } = pivotPlanRequests(sheets, evaluatorFor, options);
  const plans = new Map<string, PivotOutputPlan>();
  const nativeFootprints = new Map<string, PivotRange[]>();
  // What each PivotChart plots, by the plan key of its PivotTable: every
  // drawn report's reading, and the reading of a report no chart could see
  // otherwise (a side-panel PivotTable without an on-grid destination).
  const chartReports = new Map<string, PivotChartReport>();
  const byId = new Map(sheets.map((sheet) => [sheet.id, sheet]));
  for (const request of requests) {
    const sheet = byId.get(request.sheetId)!;
    const plan = planPivotRequest(request, {
      ...options,
      occupied: (row, col) => sheetCellOccupied(sheet, row, col),
    });
    if (!request.chartOnly) {
      nativeFootprints.set(request.planKey, [...request.footprint]);
      plans.set(request.planKey, plan);
    }
    if (plan.chart) chartReports.set(request.planKey, plan.chart);
  }
  return { plans, nativeFootprints, chartReports, chartMatrices, held };
}

/** The plan keys of the PivotTables a PivotChart is bound to. */
function boundPivotPlanKeysOf(sheets: readonly PivotOverlaySheet[]): Set<string> {
  const bound = new Set<string>();
  for (const dataSheet of sheets) {
    for (const chart of [...(dataSheet.charts ?? []), ...(dataSheet.objects ?? [])]) {
      const source = (chart as { pivotSource?: SpreadsheetPivotChartSource }).pivotSource;
      const pivot = source ? resolvePivotChartPivot(sheets, source) : null;
      if (pivot) bound.add(pivotPlanKey(pivot.sheet.id, pivot.slot));
    }
  }
  return bound;
}

/** Every PivotTable that gets a derived report, with the sheet and the range
 * it is computed from. */
function* plannedPivots(sheets: readonly PivotOverlaySheet[], boundPivotPlanKeys: ReadonlySet<string>) {
  for (const { candidate, slot, rawPivot } of sheets.flatMap(
    (dataSheet) =>
      sheetPivots(dataSheet).map((pivot, pivotSlot) => ({
        candidate: dataSheet,
        slot: pivotSlot,
        rawPivot: pivot,
      })),
  )) {
    const planKey = pivotPlanKey(candidate.id, slot);
    // An imported PivotTable the file only preserves keeps the report it was
    // saved with, even when its authority is past what the window models
    // and does not normalize (R480-X04: 136 cache fields, no snapshot, its
    // source on 'Analysis-2012 and 2015'). Dropped, it read as a PivotTable
    // of its own sheet: a report of five rows of its own cells, and its
    // PivotChart printed "no visible chart data" where Excel plots the cache.
    const rawNative = (rawPivot as { native?: unknown }).native;
    if (rawNative && typeof rawNative === "object" &&
      (rawNative as { disposition?: unknown }).disposition === "preserve-only") continue;
    const sourceRange = normalizePivotRange(rawPivot.source);
    const maxCol = sourceRange
      ? sourceRange.c2 - sourceRange.c1
      : candidate.columnCount - 1;
    const config = normalizePivotConfig(rawPivot, maxCol);
    const chartOnly = !config.output && boundPivotPlanKeys.has(planKey);
    // A layout with any field is drawn, a label list or a PivotTable being
    // built included; a fully cleared one keeps its placeholder.
    if (
      (!config.output && !chartOnly) ||
      !sourceRange ||
      (config.rows.length === 0 &&
        config.columns.length === 0 &&
        config.values.length === 0) ||
      config.native?.disposition === "preserve-only"
    ) {
      continue;
    }
    // An authored PivotTable placed on another worksheet names its source
    // sheet by id (`sourceSheetId`); an imported one names it by its cache's
    // sheet name; one without either reads the sheet it sits on.
    const sourceCandidate = config.sourceSheetId
      ? sheets.find(
          (sheetCandidate) => sheetCandidate.id === config.sourceSheetId,
        )
      : config.native
        ? sheets.find(
            (sheetCandidate) =>
              sheetCandidate.name === config.native?.sourceSheetName,
          )
        : candidate;
    // A Scenario PivotTable report reads only its own cache, so it stays
    // whole when the sheet its scenarios were on is gone.
    const readSheet =
      sourceCandidate ?? (pivotIsScenarioReport(config) ? candidate : undefined);
    if (!readSheet) continue;
    // A table source is read where the table is now.
    const liveRange =
      pivotLiveSourceRange(config, readSheet.dataTables) ?? sourceRange;
    yield { candidate, slot, planKey, config, chartOnly, readSheet, liveRange };
  }
}

/**
 * The reports computed from their source's live cells (no cache snapshot),
 * each with the rows it reads: a window holding only part of a large tiled
 * sheet reads those rows before such a report is computed.
 */
export function pivotLiveSourceReads(sheets: readonly PivotOverlaySheet[]): PivotLiveSourceRead[] {
  const out: PivotLiveSourceRead[] = [];
  for (const { candidate, planKey, config, readSheet, liveRange } of plannedPivots(sheets, boundPivotPlanKeysOf(sheets))) {
    if (pivotReadsLiveCells(config)) {
      out.push({ planKey, sheetId: candidate.id, sourceSheetId: readSheet.id, range: liveRange });
    }
  }
  return out;
}

/** Every PivotTable's request, read from the workbook evaluated without the
 * reports, and the rows each PivotChart's PivotTable is computed from. A
 * report whose live source rows are not all in the model is held instead
 * (`sourceReadable`). */
export function pivotPlanRequests(
  sheets: readonly PivotOverlaySheet[],
  evaluatorFor: (sheetName: string) => SheetEvaluator,
  options: PivotPlanReadOptions = {},
): { requests: PivotPlanRequest[]; chartMatrices: Map<string, PivotMatrix>; held: PivotLiveSourceRead[] } {
  const requests: PivotPlanRequest[] = [];
  const held: PivotLiveSourceRead[] = [];
  // The rows each such report is computed from (a field button's list).
  const pivotChartMatrices = new Map<string, PivotMatrix>();
  const boundPivotPlanKeys = boundPivotPlanKeysOf(sheets);
  // Every PivotTable of every sheet gets its own derived report.
  for (const { candidate, planKey, config, chartOnly, readSheet, liveRange } of plannedPivots(sheets, boundPivotPlanKeys)) {
    // Totals of the rows around the screen are no report of the source: one
    // read from live cells waits until every row it reads is in.
    if (
      options.sourceReadable &&
      pivotReadsLiveCells(config) &&
      !options.sourceReadable(readSheet.id, liveRange)
    ) {
      held.push({ planKey, sheetId: candidate.id, sourceSheetId: readSheet.id, range: liveRange });
      continue;
    }
    const sourceName =
      typeof readSheet.name === "string" && readSheet.name
        ? readSheet.name
        : readSheet.id;
    const evaluator = evaluatorFor(sourceName);
    const sourceMerges = normalizeMerges(readSheet.merges);
    const sourceCovered = coveredCells(sourceMerges);
    const sourceSpill = buildSpillOverlay(
      readSheet.cells,
      sourceCovered,
      evaluator,
    );
    const liveMatrix: string[][] = [];
    const readRange = pivotSourceRowsRead(liveRange, readSheet, sourceSpill);
    for (let row = readRange.r1; row <= readRange.r2; row += 1) {
      const values: string[] = [];
      for (let col = readRange.c1; col <= readRange.c2; col += 1) {
        values.push(
          resolveCellDisplay(
            readSheet.cells[`${row}:${col}`],
            row,
            col,
            evaluator,
            sourceSpill,
          ),
        );
      }
      liveMatrix.push(values);
    }
    const matrix = pivotComputationMatrix(config, liveMatrix);
    if (boundPivotPlanKeys.has(planKey)) pivotChartMatrices.set(planKey, matrix);
    if (chartOnly) {
      requests.push({ planKey, sheetId: candidate.id, config, matrix, footprint: [],
        rowCount: candidate.rowCount, columnCount: candidate.columnCount, chartOnly: true });
      continue;
    }
    // The cells an imported PivotTable left behind (its report, and the report
    // filter rows above it while they hold only filter captions): the derived
    // output replaces them, so they neither collide with it nor feed formulas.
    const nativeOutput = config.native
      ? normalizePivotRange(config.native.outputRef)
      : null;
    const nativeFootprint = nativeOutput ? [nativeOutput] : [];
    const nativePageBlock =
      nativeOutput && config.native
        ? nativePivotPageBlock(
            nativeOutput,
            config.filters.length,
            new Set(config.native.fields.map((field) => field.name)),
            (row, col) => candidate.cells[`${row}:${col}`],
          )
        : null;
    if (nativePageBlock) nativeFootprint.push(nativePageBlock);
    requests.push({
      planKey,
      sheetId: candidate.id,
      config,
      matrix,
      footprint: nativeFootprint,
      // An imported sheet is only as large as its used range, which the
      // PivotTable's saved output often fills. Excel's grid is not bounded
      // by it: a filter, slicer or timeline that admits more items grows
      // the output into the empty cells beyond.
      rowCount: config.native ? MAX_GRID_ROWS : candidate.rowCount,
      columnCount: config.native ? MAX_GRID_COLUMNS : candidate.columnCount,
      chartOnly: false,
    });
  }
  return { requests, chartMatrices: pivotChartMatrices, held };
}

/** Whether a report is placed and free: its cells join the read model. */
export function pivotPlanReady(plan: PivotOutputPlan): boolean {
  return plan.status === "ready" && plan.collisions.length === 0 && Boolean(plan.placement);
}

/**
 * One sheet's input with its placed reports (its first PivotTable's first)
 * as virtual cells, an imported report's leftovers blank, and the first
 * report's placement answering GETPIVOTDATA. The grid and the calculation
 * session lay the same overlay.
 */
export function overlayPivotInput(
  input: WorkbookSheetInput,
  readyPlans: ReadonlyArray<{ readonly plan: PivotOutputPlan; readonly footprint: readonly PivotRange[] }>,
): WorkbookSheetInput {
  if (readyPlans.length === 0) return input;
  const getRawCell = input.getRawCell;
  return {
    ...input,
    getRawCell(row, col) {
      for (const { plan } of readyPlans) {
        const derivedCell = plan.cells.get(`${row}:${col}`);
        if (derivedCell) {
          return derivedCell.value === null || derivedCell.value === ""
            ? null
            : String(derivedCell.value);
        }
      }
      if (readyPlans.some(({ footprint }) => pivotRangesContain(footprint, row, col))) {
        return null;
      }
      return getRawCell(row, col);
    },
    options: {
      ...(input.options ?? {}),
      resolvePivotData: readyPlans[0].plan.resolvePivotData,
    },
  };
}

/** The workbook's sheet inputs with the placed reports as virtual cells and
 * GETPIVOTDATA's placement resolver. */
export function withPivotOverlay(
  sheets: readonly PivotOverlaySheet[],
  inputs: readonly WorkbookSheetInput[],
  overlay: Pick<WorkbookPivotPlans, "plans" | "nativeFootprints">,
): WorkbookSheetInput[] {
  // A valid, collision-free plan becomes part of the workbook's read model:
  // ordinary A1 references can read its derived cells, and GETPIVOTDATA gets
  // the placement resolver. The values remain virtual — sheet.cells and the
  // persisted JSON are untouched.
  return inputs.map((input, index): WorkbookSheetInput => {
    const candidate = sheets[index];
    // The drawn reports of the sheet, its first PivotTable's first.
    const readyPlans = candidate
      ? sheetPivots(candidate)
          .map((_, slot) => {
            const planKey = pivotPlanKey(candidate.id, slot);
            const plan = overlay.plans.get(planKey);
            return plan && pivotPlanReady(plan)
              ? { plan, footprint: overlay.nativeFootprints.get(planKey) ?? [] }
              : null;
          })
          .filter((item): item is NonNullable<typeof item> => item !== null)
      : [];
    return overlayPivotInput(input, readyPlans);
  });
}
