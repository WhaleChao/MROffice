/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PivotTables as pure workbook transactions: reading a source range into the
 * snapshot Create PivotTable stands on, placing a report safely on a sheet or
 * on a new worksheet, committing a recipe change (with its PivotCharts and the
 * column fit Excel applies on update) and the PivotChart record. Moved out of
 * SpreadsheetWindow so the window and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps the dialogs, receipts,
 * protection notices, localized captions and selection. The report itself is
 * never stored, only its recipe (lib/spreadsheet/pivot).
 */

import {
  hasCellMetadata,
  withBestFitColumns,
  type CellButtonsSheet,
  type CellMap,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  normalizeSpreadsheetPivotChartSource,
  type SpreadsheetChart,
  type SpreadsheetChartType,
  type SpreadsheetPivotChartSource,
} from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { spreadsheetChartBornWithModel } from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";
import {
  rangesOverlap as conditionalRangesOverlap,
  ruleOverlaps as conditionalRuleOverlaps,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import type { DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { validationRuleAreas, type ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { buildSpillOverlay, resolveCellDisplay } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import {
  EXCEL_1900_FICTITIOUS_LEAP_SERIAL,
  excelMaxDateSerial,
  excelSystemSerialToUtcMs,
} from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";
import {
  formatCellValue,
  isFormulaError,
  type SheetEvaluator,
  type SpillOverlay,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  autoFitColumnWidth,
  clampColumnWidth,
  DEFAULT_SPREADSHEET_COLUMN_COUNT,
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
  DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_COUNT,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  offsetOfLine,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { coveredCells, normalizeMerges, rangeHasMerge, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { applyNumberFormat, numberFormatFamily } from "@mrnote/core/web/lib/spreadsheet/number-format.ts";
import {
  headerLabel,
  isPercentShowValuesAs,
  markPivotTextCells,
  normalizePivotConfig,
  normalizePivotRange,
  pivotComputationMatrix,
  pivotIsScenarioReport,
  pivotLiveSourceRange,
  pivotSourceRowsRead,
  pivotNativeScalarDisplay,
  pivotRangeToA1,
  profilePivotFieldsFromRange,
  rebasePivotNativeSnapshot,
  refreshPivotSnapshot,
  sheetPivots,
  sheetWithPivotAt,
  type PivotCacheSnapshot,
  type PivotConfig,
  type PivotFieldProfile,
  type PivotMatrix,
  type PivotNativeScalar,
  type PivotNativeScalarMatrix,
  type PivotRange,
  type PivotRefreshRefusal,
} from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import {
  nextPivotTableName,
  pivotChartReportCategoryLevels,
  pivotChartReportToRows,
  pivotChartWithoutRefreshLosses,
  pivotTableName,
  resolvePivotChartPivot,
} from "@mrnote/core/web/lib/spreadsheet/pivot-chart.ts";
import { staticChartFromPivotChart } from "./pivot-chart-static";
import {
  planPivotOutput,
  suggestPivotOutputAnchor,
  type PivotChartReport,
  type PivotOutputLabels,
  type PivotOutputPlan,
} from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";
import { defaultSheetName, nextDefaultSheetNumber } from "./sheet-name-rules";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";
import { excelSystemHasFictitiousLeapDay } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

/** A drawing object a PivotChart may be (the chart objects of `objects`). */
export interface PivotChartObjectLike {
  id: string;
  type: string;
  pivotSource?: unknown;
  [key: string]: unknown;
}

/** The parts of a worksheet the PivotTable transactions read and write. */
export interface PivotButtonSheet extends CellButtonsSheet {
  id: string;
  name: string;
  pivot?: PivotConfig;
  morePivots?: PivotConfig[];
  charts?: SpreadsheetChart[];
  objects?: PivotChartObjectLike[];
  merges?: MergeSpan[];
  dataTables?: DataTable[];
  conditionalFormats?: ConditionalFormatRule[];
  validations?: ValidationRule[];
  rowHeights?: Record<string, number>;
  defaultRowHeight?: number;
}

/** The cells one PivotTable source may hold (Create PivotTable's bound). */
export const MAX_PIVOT_SOURCE_CELLS = 100_000;

/**
 * The table a PivotTable source box names ("Table1" / "表1"), wherever in the
 * workbook it is: Excel table names are workbook-wide and case-insensitive. A
 * table without its header row has no field names to offer, so it is not a
 * source. Null when no table, or more than one, answers to the text.
 */
export function pivotSourceTableNamed<S extends { dataTables?: readonly DataTable[] }>(
  workbook: readonly S[],
  rawName: string,
): { sheet: S; table: DataTable } | null {
  const wanted = rawName.trim().toLocaleLowerCase();
  if (!wanted) return null;
  const matches = workbook.flatMap((candidate) =>
    (candidate.dataTables ?? []).flatMap((table) =>
      table.headerRow !== false &&
      typeof table.name === "string" &&
      table.name.trim().toLocaleLowerCase() === wanted
        ? [{ sheet: candidate, table }]
        : [],
    ),
  );
  return matches.length === 1 ? matches[0] : null;
}

/** Excel places a PivotTable made on a new worksheet at A3, leaving the rows
 * above it for report filters. */
export const PIVOT_NEW_SHEET_ANCHOR = "A3";

/**
 * The id and tab name of the worksheet a PivotTable command inserts, by the
 * default-name rule every inserted sheet follows. Insert PivotTable and Move
 * PivotTable > New Worksheet name their sheet the same way.
 */
export function nextPivotWorksheetIdentity(workbook: readonly { id: string; name: string }[]): {
  id: string;
  name: string;
} {
  const existingIds = new Set(workbook.map((candidate) => candidate.id));
  const existingNames = new Set(
    workbook.map((candidate) => candidate.name.toLowerCase()),
  );
  let suffix = nextDefaultSheetNumber(existingNames, workbook.length + 1);
  while (
    existingIds.has(`sheet-${suffix}`) ||
    existingNames.has(defaultSheetName(suffix).toLowerCase())
  ) {
    suffix += 1;
  }
  return { id: `sheet-${suffix}`, name: defaultSheetName(suffix) };
}

/**
 * The recipe of a PivotTable created on a NEW worksheet, and the size that
 * worksheet needs. The source stays on its own sheet, named by id so a rename
 * cannot detach it. Null when the source is unreadable or the report would not
 * fit any sheet.
 */
export function pivotConfigForNewSheet(
  base: PivotConfig,
  matrix: PivotMatrix,
  sourceSheetId: string,
): { config: PivotConfig; rowCount: number; columnCount: number; top: PivotRange } | null {
  const sourceRange = normalizePivotRange(base.source);
  if (!sourceRange) return null;
  const config: PivotConfig = {
    ...base,
    source: pivotRangeToA1(sourceRange),
    sourceSheetId,
    output: PIVOT_NEW_SHEET_ANCHOR,
  };
  const plan = planPivotOutput(config, matrix, {
    rowCount: MAX_GRID_ROWS,
    columnCount: MAX_GRID_COLUMNS,
  });
  if (plan.status !== "ready" || !plan.output) return null;
  return {
    config,
    // A default-sized sheet, grown when the report is larger than one.
    rowCount: Math.max(DEFAULT_SPREADSHEET_ROW_COUNT, plan.output.r2 + 2),
    columnCount: Math.max(
      DEFAULT_SPREADSHEET_COLUMN_COUNT,
      Math.max(plan.output.c2, plan.pageRange?.c2 ?? 0) + 2,
    ),
    top: plan.output,
  };
}

/**
 * What Excel gives a PivotTable it CREATES: compact form - the row fields
 * share one column under "Row Labels" / "行标签" - with each group's subtotal on
 * the group's own row. An existing or imported PivotTable keeps its layout.
 */
export const NEW_PIVOT_LAYOUT_DEFAULTS = {
  reportLayout: "compact",
  subtotalPosition: "top",
} as const satisfies Partial<PivotConfig>;

/**
 * Excel's PivotTable option "Autofit column widths on update" is on for every
 * PivotTable it creates (PivotTable Options > Layout & Format; the file's
 * `useAutoFormatting`), so a report's columns fit their captions and numbers
 * each time it is created, refreshed or laid out again.
 *
 * MRAI's report is derived, not stored, so the fit is written as the host
 * sheet's column widths in the SAME commit as the PivotTable change. It only
 * ever widens: a width the user set is never reduced, and a column that was
 * free or already fitted stays best-fit. An imported sheet keeps the file's own
 * grid metrics, and an imported PivotTable keeps the widths it came with.
 */
export function sheetWithPivotColumnsFitted<S extends PivotButtonSheet & { nativeGrid?: unknown }>(
  host: S,
  plan: PivotOutputPlan,
  config: PivotConfig,
): S {
  if (host.nativeGrid || config.native || plan.status !== "ready") return host;
  const defaultWidth =
    host.defaultColumnWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX;
  const wanted = new Map<number, number>();
  for (const cell of plan.cells.values()) {
    if (cell.value === null || cell.value === "") continue;
    const field =
      cell.valueIndex === undefined ? undefined : config.values[cell.valueIndex];
    const text =
      typeof cell.value !== "number"
        ? String(cell.value)
        : field?.numberFormat
          ? applyNumberFormat(cell.value, field.numberFormat)
          : field?.showAs && isPercentShowValuesAs(field.showAs.mode)
            ? `${Math.round(cell.value * 10_000) / 100}%`
            : formatCellValue(cell.value);
    // A field caption and a report-filter cell carry a dropdown button, an
    // outer item its expand box, and compact form indents inner items.
    const chrome =
      (cell.field !== undefined ? 20 : 0) +
      (cell.drill ? 14 : 0) +
      (cell.indent ?? 0) * 12;
    const width = clampColumnWidth(
      autoFitColumnWidth(text, defaultWidth) + chrome,
      defaultWidth,
    );
    if (width > (wanted.get(cell.col) ?? 0)) wanted.set(cell.col, width);
  }
  let next = host;
  const fitted: number[] = [];
  for (const [col, width] of wanted) {
    const stored = next.columnWidths?.[col];
    if (width <= (stored ?? defaultWidth)) continue;
    if (stored === undefined || next.bestFitColumns?.[col] === true) {
      fitted.push(col);
    }
    next = { ...next, columnWidths: { ...next.columnWidths, [col]: width } };
  }
  return fitted.length > 0 ? withBestFitColumns(next, fitted, true) : next;
}

/** Exact typed source scalar for an explicit native Pivot Refresh. */
export function pivotNativeDateScalar(
  value: number,
  numberFormat: unknown,
  date1904: ExcelDateSystemArg = false,
): PivotNativeScalar | null {
  if (
    !Number.isFinite(value) ||
    value < 0 ||
    value > excelMaxDateSerial(date1904) + 0.999_999 ||
    numberFormatFamily(numberFormat) !== "date"
  ) {
    return null;
  }
  // Excel's 1900 calendar; its fictitious 1900-02-29 has no ISO day to cache.
  // The 1904 system reads the serial four years and one day later.
  if (excelSystemHasFictitiousLeapDay(date1904) && Math.floor(value) === EXCEL_1900_FICTITIOUS_LEAP_SERIAL) return null;
  const instant = new Date(Math.round(excelSystemSerialToUtcMs(value, date1904)));
  if (!Number.isFinite(instant.getTime())) return null;
  const iso = instant.toISOString();
  return {
    kind: "d",
    raw:
      instant.getUTCMilliseconds() === 0
        ? iso.slice(0, 19)
        : iso.slice(0, -1),
  };
}

export function pivotNativeScalarFromLiveCell(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay?: SpillOverlay,
): PivotNativeScalar {
  if (cell?.formula) {
    const computed = evaluator.getComputed(row, col);
    if (computed === null || computed === undefined || computed === "") {
      return { kind: "m", raw: "" };
    }
    if (typeof computed === "boolean") {
      return { kind: "b", raw: computed ? "1" : "0" };
    }
    if (typeof computed === "number" && Number.isFinite(computed)) {
      return (
        pivotNativeDateScalar(computed, cell.style?.numberFormat, evaluator.date1904) ?? {
          kind: "n",
          raw: String(computed),
        }
      );
    }
    if (isFormulaError(computed)) {
      return { kind: "e", raw: evaluator.getDisplay(row, col) };
    }
    return { kind: "s", raw: String(computed) };
  }
  const raw = String(cell?.value ?? "");
  if (raw === "") {
    // A blank cell inside an array formula's spill shows the spilled element;
    // the pivot's source is what the grid shows, typed, not the blank.
    const spill = overlay?.spilled.get(`${row}:${col}`);
    if (spill === undefined || spill === null || spill === "") {
      return { kind: "m", raw: "" };
    }
    if (typeof spill === "boolean") {
      return { kind: "b", raw: spill ? "1" : "0" };
    }
    if (typeof spill === "number" && Number.isFinite(spill)) {
      return (
        pivotNativeDateScalar(spill, cell?.style?.numberFormat, evaluator.date1904) ?? {
          kind: "n",
          raw: String(spill),
        }
      );
    }
    if (isFormulaError(spill)) {
      return { kind: "e", raw: formatCellValue(spill) };
    }
    return { kind: "s", raw: String(spill) };
  }
  if (cell?.kind === "number") {
    const numeric = Number(raw);
    return (
      pivotNativeDateScalar(numeric, cell.style?.numberFormat, evaluator.date1904) ?? {
        kind: "n",
        raw,
      }
    );
  }
  return { kind: "s", raw };
}

/**
 * Each source column's own number format, read the way Excel gives a new
 * value field its format: from the column's first non-blank data cell (the
 * row under the header). A column whose first data cell is unformatted, or
 * carries General, has none; undefined keeps the value field's default.
 */
export function pivotSourceColumnNumberFormats(
  cells: CellMap,
  range: { r1: number; c1: number; r2: number; c2: number },
): (string | undefined)[] {
  const formats: (string | undefined)[] = [];
  for (let col = range.c1; col <= range.c2; col += 1) {
    let format: string | undefined;
    for (let row = range.r1 + 1; row <= range.r2; row += 1) {
      const cell = cells[`${row}:${col}`];
      if (!cell || (cell.value === "" && !cell.formula)) continue;
      const numberFormat = cell.style?.numberFormat?.trim() ?? "";
      if (numberFormat !== "" && numberFormat !== "General") {
        format = numberFormat;
      }
      break;
    }
    formats.push(format);
  }
  return formats;
}

/**
 * The report a PivotChart's PivotTable gives at save time, read from the
 * saved workbook's own evaluator (the same reading the grid draws).
 */
export function pivotChartReportForSave(
  sheets: readonly PivotButtonSheet[],
  source: SpreadsheetPivotChartSource,
  evaluatorFor: (sheetName: string) => SheetEvaluator,
  labels: Partial<PivotOutputLabels> | undefined,
): PivotChartReport | null {
  const bound = resolvePivotChartPivot(sheets, source);
  if (!bound) return null;
  const host = bound.sheet;
  const sourceRange = normalizePivotRange(bound.pivot.source);
  if (!sourceRange) return null;
  const config = normalizePivotConfig(bound.pivot, sourceRange.c2 - sourceRange.c1);
  if (config.native?.disposition === "preserve-only") return null;
  const sourceSheet = config.sourceSheetId
    ? sheets.find((candidate) => candidate.id === config.sourceSheetId)
    : config.native
      ? sheets.find((candidate) => candidate.name === config.native?.sourceSheetName)
      : host;
  if (!sourceSheet) return null;
  try {
    const evaluator = evaluatorFor(sourceSheet.name || sourceSheet.id);
    const spill = buildSpillOverlay(
      sourceSheet.cells,
      coveredCells(normalizeMerges(sourceSheet.merges)),
      evaluator,
    );
    const range = pivotSourceRowsRead(pivotLiveSourceRange(config, sourceSheet.dataTables) ?? sourceRange, sourceSheet, spill);
    const live: string[][] = [];
    for (let row = range.r1; row <= range.r2; row += 1) {
      const values: string[] = [];
      for (let col = range.c1; col <= range.c2; col += 1) {
        values.push(resolveCellDisplay(sourceSheet.cells[`${row}:${col}`], row, col, evaluator, spill));
      }
      live.push(values);
    }
    return (
      planPivotOutput(config, pivotComputationMatrix(config, live), {
        rowCount: host.rowCount,
        columnCount: host.columnCount,
        labels,
      }).chart ?? null
    );
  } catch {
    return null;
  }
}


/**
 * A new PivotChart record for one PivotTable: beside its report, Excel's
 * 5 x 3 inch chart with the legend on the right, titled by its one series.
 * The grid draws a chart that has never been moved at one fixed spot, level
 * with the report's first row.
 */
export function newPivotChartRecord(
  host: PivotButtonSheet & { columnWidths?: Record<string, number>; defaultColumnWidth?: number },
  source: SpreadsheetPivotChartSource,
  report: PivotChartReport | null | undefined,
  output: PivotRange,
  chartType: SpreadsheetChartType,
  id = `pivot-chart-${Date.now().toString(36)}-${host.charts?.length ?? 0}`,
  chartSubtype?: string,
): SpreadsheetChart {
  const rows = report ? pivotChartReportToRows(report) : [];
  const seriesCount = report?.series.length ?? 0;
  const columnWidth = (col: number) =>
    host.columnWidths?.[String(col)] ?? host.defaultColumnWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX;
  const rowHeightOf = (row: number) =>
    host.rowHeights?.[String(row)] ?? host.defaultRowHeight ?? DEFAULT_SPREADSHEET_ROW_HEIGHT_PX;
  const chart: SpreadsheetChart = {
    id,
    type: chartType,
    range: { ...output },
    anchor: {
      row: output.r1,
      col: Math.min(host.columnCount - 1, output.c2 + 1),
    },
    x:
      DEFAULT_SPREADSHEET_HEADER_WIDTH_PX +
      offsetOfLine(output.c2 + 1, host.columnCount, columnWidth) +
      24,
    y:
      DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX +
      offsetOfLine(output.r1, host.rowCount, rowHeightOf),
    width: 480,
    height: 288,
    title: seriesCount === 1 ? (report?.series[0]?.name ?? "") : "",
    legendPosition: "right",
    pivotSource: { ...source },
    ...(chartSubtype ? { chartSubtype: chartSubtype as SpreadsheetChart["chartSubtype"] } : {}),
  };
  const categoryLevels = report ? pivotChartReportCategoryLevels(report) : null;
  return spreadsheetChartBornWithModel(
    chart,
    rows.length ? { rows, headerRow: true, ...(categoryLevels ? { categoryLevels } : {}) } : null,
  );
}

/** What Create PivotTable reads of its source range. */
export interface PivotSourceSnapshot {
  sheetId: string;
  /** Plain A1 range the recipe persists. */
  source: string;
  /** Sheet-qualified absolute reference the dialog's Table/Range box shows. */
  sourceRef: string;
  matrix: string[][];
  scalarMatrix: PivotNativeScalarMatrix;
  fields: PivotFieldProfile[];
  dataRowCount: number;
}

/** A source range as the dialog's Table/Range box writes it: sheet-qualified
 * and absolute (Sheet1!$A$1:$D$10). */
export function pivotSourceReference(
  sheetName: string,
  range: PivotRange,
  sheetRefPrefix: (name: string) => string,
  columnLabel: (index: number) => string,
): string {
  return `${sheetRefPrefix(sheetName)}$${columnLabel(range.c1)}$${range.r1 + 1}:$${columnLabel(range.c2)}$${range.r2 + 1}`;
}

/**
 * Why Create PivotTable cannot summarise a source range, or null when it can:
 * - "partial": the sheet is a large tiled one holding only the rows around the
 *   screen (`TiledSheets.isWhollyResident` is false). Its totals would be of
 *   the rows read so far, not of the range, so nothing is summarised.
 * - "tooLarge": more than `MAX_PIVOT_SOURCE_CELLS` cells.
 * A range that is simply unusable (outside the sheet, one row high) is not a
 * refusal here: `pivotSourceSnapshot` answers null for it.
 */
export type PivotSourceRefusal = "partial" | "tooLarge";

export function pivotSourceRefusal(
  range: PivotRange,
  whollyResident: boolean,
): PivotSourceRefusal | null {
  if (!whollyResident) return "partial";
  const cells = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  return cells > MAX_PIVOT_SOURCE_CELLS ? "tooLarge" : null;
}

/**
 * The source of a PivotTable as Create PivotTable reads it: the header row as
 * the grid shows it, every data cell as its typed value (a formula's result, a
 * dynamic array's spilled element), and the fields inferred from them. Null
 * for a range outside the sheet, one row high, past the source bound, or on a
 * sheet whose cells are not all in the model (`whollyResident: false`, a large
 * tiled sheet read around the screen): a snapshot of the rows read so far
 * would total part of the source as if it were all of it.
 */
export function pivotSourceSnapshot(
  sheet: Pick<PivotButtonSheet, "id" | "name" | "cells" | "rowCount" | "columnCount">,
  range: PivotRange,
  read: {
    display: (cell: SpreadsheetCell | undefined, row: number, col: number) => string;
    evaluator: SheetEvaluator;
    overlay: SpillOverlay;
    sheetRefPrefix: (name: string) => string;
    columnLabel: (index: number) => string;
    /** False while the sheet holds only some of its rows (default true). */
    whollyResident?: boolean;
  },
): PivotSourceSnapshot | null {
  if (read.whollyResident === false) return null;
  const coords = [range.r1, range.c1, range.r2, range.c2];
  if (
    !coords.every(Number.isInteger) ||
    range.r1 < 0 ||
    range.c1 < 0 ||
    range.r2 < range.r1 ||
    range.c2 < range.c1 ||
    range.r2 >= sheet.rowCount ||
    range.c2 >= sheet.columnCount
  ) {
    return null;
  }
  const height = range.r2 - range.r1 + 1;
  const width = range.c2 - range.c1 + 1;
  if (height < 2 || height * width > MAX_PIVOT_SOURCE_CELLS) return null;
  const matrix: string[][] = [];
  const scalarMatrix: PivotNativeScalarMatrix = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    const values: string[] = [];
    const scalarValues: PivotNativeScalar[] = [];
    for (let col = range.c1; col <= range.c2; col += 1) {
      const cell = sheet.cells[`${row}:${col}`];
      const scalar =
        row === range.r1
          ? { kind: "s" as const, raw: read.display(cell, row, col) }
          : pivotNativeScalarFromLiveCell(cell, row, col, read.evaluator, read.overlay);
      scalarValues.push(scalar);
      values.push(pivotNativeScalarDisplay(scalar));
    }
    matrix.push(values);
    scalarMatrix.push(scalarValues);
  }
  const source = pivotRangeToA1(range);
  // What the dialog's Table/Range box shows. Excel writes the proposed
  // range the way it writes every range reference: sheet-qualified and
  // absolute (Sheet1!$A$1:$D$10). The stored config keeps the plain A1
  // form the pivot model persists; this is the editable reference only,
  // and normalizePivotRange reads the prefix and the $ back off.
  const sourceRef = pivotSourceReference(sheet.name, range, read.sheetRefPrefix, read.columnLabel);
  const inference = profilePivotFieldsFromRange(
    matrix,
    scalarMatrix,
    pivotSourceColumnNumberFormats(sheet.cells, range),
  );
  return {
    sheetId: sheet.id,
    source,
    sourceRef,
    matrix,
    scalarMatrix,
    fields: inference.fields,
    dataRowCount: inference.dataRowCount,
  };
}

/**
 * The recipe with its report placed where it covers nothing the sheet owns:
 * a requested top-left cell, or Excel's suggestion beside / below the source.
 * The derived output may not cover a filled cell, a locked cell, a merge, a
 * conditional format, a validation rule or a table; the exact footprint (report
 * filters included) is also the protection authority for the write. A layout
 * without any field is returned as it is. Null when no safe place exists.
 */
export function pivotConfigWithSafeOutput<S extends PivotButtonSheet>(
  sheet: S,
  base: PivotConfig,
  matrix: string[][],
  requestedOutput: string | undefined,
  context: {
    rangeHasLockedCell: (sheet: S, range: PivotRange) => boolean;
    merges?: readonly MergeSpan[];
    conditionalFormats?: readonly ConditionalFormatRule[];
    validations?: readonly ValidationRule[];
    dataTables?: readonly DataTable[];
    /** The location was typed by the user (not the dialog's proposal): it
     *  may be any blank cell of the grid, the sheet growing to hold it. */
    growToFit?: boolean;
  },
): PivotConfig | null {
  // Any layout with a field is drawn on the grid (a label list or a
  // PivotTable being built included), so its footprint is checked too.
  if (
    base.rows.length === 0 &&
    base.columns.length === 0 &&
    base.values.length === 0
  ) {
    return base;
  }
  const mergeSpans = context.merges ?? normalizeMerges(sheet.merges);
  const conditionalFormats = context.conditionalFormats ?? sheet.conditionalFormats ?? [];
  const validations = context.validations ?? sheet.validations ?? [];
  const dataTables = context.dataTables ?? sheet.dataTables ?? [];
  // A location the user names may be any blank cell on the grid, as typing
  // there would be: the sheet grows to hold the report when it is committed
  // (pivotCommitPlan). Bounded by the sheet's stored extent, J160 was refused
  // until something was typed below it (user audit S-020). The location the
  // dialog proposes itself still fits the sheet as it stands.
  const bounds = requestedOutput !== undefined && context.growToFit
    ? { rowCount: MAX_GRID_ROWS, columnCount: MAX_GRID_COLUMNS }
    : { rowCount: sheet.rowCount, columnCount: sheet.columnCount };
  let output: string | null = null;
  if (requestedOutput !== undefined) {
    const requestedRange = normalizePivotRange(requestedOutput);
    if (!requestedRange) return null;
    output = pivotRangeToA1({
      r1: requestedRange.r1,
      c1: requestedRange.c1,
      r2: requestedRange.r1,
      c2: requestedRange.c1,
    });
  } else {
    const initialOutput = suggestPivotOutputAnchor(base.source, bounds);
    if (!initialOutput) return null;
    const initialPlan = planPivotOutput(
      { ...base, output: initialOutput },
      matrix,
      bounds,
    );
    // The one-cell probe anchor may be too close to a sheet edge for the
    // computed PivotTable. Its out-of-bounds plan still carries the exact
    // footprint (report filters included) needed to retry beside or below
    // the source safely.
    if (!initialPlan.output) return null;
    output = suggestPivotOutputAnchor(base.source, bounds, {
      height: initialPlan.output.r2 - initialPlan.output.r1 + 1,
      width:
        Math.max(initialPlan.output.c2, initialPlan.pageRange?.c2 ?? 0) -
        initialPlan.output.c1 +
        1,
      above: initialPlan.pageRange
        ? initialPlan.output.r1 - initialPlan.pageRange.r1
        : 0,
    });
  }
  if (!output) return null;
  const config = { ...base, output };
  const finalPlan = planPivotOutput(config, matrix, {
    ...bounds,
    isOccupied: (row, col) => {
      const cell = sheet.cells[`${row}:${col}`];
      return Boolean(
        cell &&
        ((cell.value ?? "") !== "" ||
          (cell.formula ?? "") !== "" ||
          hasCellMetadata(cell)),
      );
    },
  });
  const outputRange = finalPlan.output;
  if (
    finalPlan.status !== "ready" ||
    !outputRange ||
    finalPlan.collisions.length > 0 ||
    [outputRange, finalPlan.pageRange].some(
      (range) =>
        range !== null &&
        (context.rangeHasLockedCell(sheet, range) ||
          rangeHasMerge([...mergeSpans], range) ||
          conditionalFormats.some((rule) =>
            conditionalRuleOverlaps(rule, range),
          ) ||
          validations.some((rule) =>
            validationRuleAreas(rule).some((area) => conditionalRangesOverlap(area, range)),
          ) ||
          dataTables.some((table) =>
            conditionalRangesOverlap(range, {
              ...table.range,
              r2: table.range.r2 + (table.totalsRow ? 1 : 0),
            }),
          )),
    )
  ) {
    return null;
  }
  return config;
}

/** What a recipe commit changes: nothing, the host sheet alone, or several
 *  sheets (PivotCharts elsewhere, a shared cache refreshed together). */
export type PivotCommitPlan<S> =
  | { kind: "unchanged" }
  | { kind: "sheet"; sheet: S }
  | { kind: "workbook"; sheets: S[] };

export interface PivotCommitInput<S extends PivotButtonSheet> {
  sheets: readonly S[];
  /** The PivotTable's host sheet (one of `sheets`). */
  host: S;
  /** The PivotTable's slot on the host (0 = `pivot`). */
  slot: number;
  /** The recipe the slot holds now. */
  active: PivotConfig | undefined;
  /** The new recipe, or null to remove the PivotTable. */
  rawConfig: PivotConfig | null;
  /** The live source rows (header first), for captions and the column fit. */
  liveMatrix: PivotMatrix | null;
  labels?: Partial<PivotOutputLabels>;
  /** Value fields new to the recipe take the viewer's caption (identity headless). */
  localizeCaptions?: (next: PivotConfig, fieldName: (colIndex: number) => string, previous?: PivotConfig) => PivotConfig;
  /** What a PivotChart of this PivotTable shows now, as rows. */
  chartRowsFor: (source: SpreadsheetPivotChartSource | undefined) => string[][];
  /** A sheet's column letter for an untitled field. */
  columnLabel: (index: number) => string;
  /** The id of a PivotChart created with the PivotTable. */
  newChartId?: string;
  options?: {
    resetAssociatedPivotCharts?: boolean;
    /** Refresh: the PivotTable's charts lose their trendlines and error
     * bars, which Excel does not keep through a refresh. */
    refreshAssociatedPivotCharts?: boolean;
    /** The source rows of a PivotTable being CREATED (no live source
     * matrix exists for it yet), for the column fit. */
    sourceMatrix?: PivotMatrix;
    /** Create PivotChart: the new PivotTable gets its PivotChart in the
     * same step. */
    withPivotChart?: boolean;
  };
}

/**
 * One PivotTable recipe change as a workbook transaction: the slot's recipe
 * normalized (new value fields captioned), the PivotTable's PivotCharts
 * following it (turned static when it is removed, reset by Clear All, losing
 * trendlines and error bars on Refresh), a PivotChart created with it, the
 * report's columns fitted ("Autofit column widths on update"), and a refreshed
 * shared cache reaching every PivotTable that names it. Protection is the
 * caller's (Use PivotTable & PivotChart).
 */
export function pivotCommitPlan<S extends PivotButtonSheet>(input: PivotCommitInput<S>): PivotCommitPlan<S> {
  const { sheets, host: sheet, slot: activePivotSlot, active: activePivot, rawConfig } = input;
  const options = input.options ?? {};
  // The edit lands on the active PivotTable's slot; the others stay.
  let next: S = { ...sheet };
  let committedPivot: PivotConfig | undefined;
  // The PivotCharts of THIS PivotTable, on any sheet: their PivotTable
  // leaves, or its layout is cleared, with it.
  const ownChart = (source: SpreadsheetPivotChartSource | undefined) => {
    const bound = resolvePivotChartPivot(sheets, source);
    return bound !== null && bound.sheet.id === sheet.id && bound.slot === activePivotSlot;
  };
  const chartEdit: { workbook: S[] | null } = { workbook: null };
  const mapOwnCharts = (
    mapChart: (chart: SpreadsheetChart) => SpreadsheetChart,
    mapObject: <O extends PivotChartObjectLike>(object: O) => O,
  ) => {
    let changed = false;
    chartEdit.workbook = sheets.map((candidate) => {
      const base = candidate === sheet ? next : candidate;
      // Only a chart the edit actually changes touches its sheet: a Refresh
      // that takes nothing off its charts writes nothing.
      let touched = false;
      const charts = base.charts?.map((chart) => {
        if (!ownChart(chart.pivotSource)) return chart;
        const mapped = mapChart(chart);
        if (mapped !== chart) touched = true;
        return mapped;
      });
      const objects = base.objects?.map((object) => {
        if (object.type !== "chart" || !ownChart(normalizeSpreadsheetPivotChartSource(object.pivotSource))) return object;
        const mapped = mapObject(object);
        if (mapped !== object) touched = true;
        return mapped;
      });
      if (!touched) return base;
      changed = true;
      return {
        ...base,
        ...(charts ? { charts } : {}),
        ...(objects ? { objects } : {}),
      };
    });
    if (!changed) chartEdit.workbook = null;
  };
  const fitSource = options.sourceMatrix ?? input.liveMatrix;
  if (rawConfig === null) {
    // The PivotTables after a removed one move up (`sheetWithPivotAt`).
    next = sheetWithPivotAt(next, activePivotSlot, undefined);
    // Its PivotCharts become standard charts of what they last showed
    // (Microsoft Support: deleting the PivotTable turns its PivotChart
    // into a standard chart that can no longer be pivoted or updated).
    mapOwnCharts(
      (chart) => staticChartFromPivotChart(chart, input.chartRowsFor(chart.pivotSource)),
      (object) =>
        staticChartFromPivotChart(
          object,
          input.chartRowsFor(normalizeSpreadsheetPivotChartSource(object.pivotSource)),
        ),
    );
  } else {
    const range = normalizePivotRange(rawConfig.source);
    const maxCol = range ? range.c2 - range.c1 : sheet.columnCount - 1;
    const captionSource = fitSource;
    const normalized = normalizePivotConfig(rawConfig, maxCol);
    const localized = input.localizeCaptions
      ? input.localizeCaptions(
          normalized,
          (colIndex) =>
            (captionSource ? headerLabel(captionSource, colIndex).trim() : "") ||
            input.columnLabel(colIndex),
          activePivot,
        )
      : normalized;
    const hasRecipe =
      localized.rows.length > 0 ||
      localized.columns.length > 0 ||
      localized.values.length > 0 ||
      localized.source !== "";
    committedPivot = hasRecipe ? localized : undefined;
    if (options.withPivotChart && committedPivot && !pivotTableName(committedPivot)) {
      committedPivot = { ...committedPivot, name: nextPivotTableName(sheets) };
    }
    next = sheetWithPivotAt(next, activePivotSlot, committedPivot);
    // A report placed past the sheet's extent grows the sheet to hold it
    // (its report filters included), as entering a value there would.
    if (committedPivot?.output && fitSource) {
      const reach = planPivotOutput(
        committedPivot,
        pivotComputationMatrix(committedPivot, fitSource),
        { rowCount: MAX_GRID_ROWS, columnCount: MAX_GRID_COLUMNS, labels: input.labels },
      );
      const bottom = Math.max(reach.output?.r2 ?? -1, reach.pageRange?.r2 ?? -1);
      const right = Math.max(reach.output?.c2 ?? -1, reach.pageRange?.c2 ?? -1);
      if (bottom >= next.rowCount || right >= next.columnCount) {
        next = {
          ...next,
          rowCount: Math.max(next.rowCount, bottom + 1),
          columnCount: Math.max(next.columnCount, right + 1),
        };
      }
    }
    if (options.withPivotChart && committedPivot?.output && fitSource) {
      const plan = planPivotOutput(
        committedPivot,
        pivotComputationMatrix(committedPivot, fitSource),
        { rowCount: next.rowCount, columnCount: next.columnCount, labels: input.labels },
      );
      if (plan.output) {
        next = {
          ...next,
          charts: [
            ...(next.charts ?? []),
            newPivotChartRecord(
              next,
              { sheetId: sheet.id, pivotName: pivotTableName(committedPivot) },
              plan.chart,
              plan.output,
              "column",
              input.newChartId,
            ),
          ],
        };
      }
    }
    if (options.refreshAssociatedPivotCharts) {
      mapOwnCharts(
        (chart) => pivotChartWithoutRefreshLosses(chart),
        (object) => pivotChartWithoutRefreshLosses(object),
      );
    }
    if (options.resetAssociatedPivotCharts) {
      // Clear All: the PivotChart keeps its place and binding, and loses
      // its fields, formatting and chart elements with the PivotTable's.
      const output = normalizePivotRange(localized.output ?? "");
      const resetRange = output
        ? { r1: output.r1, r2: output.r1, c1: output.c1, c2: output.c1 }
        : null;
      if (resetRange) {
        mapOwnCharts(
          (chart) =>
            ({
              id: chart.id,
              ...(chart.name ? { name: chart.name } : {}),
              ...(chart.hidden === true ? { hidden: true } : {}),
              type: "column",
              range: resetRange,
              anchor: { ...chart.anchor },
              width: chart.width,
              height: chart.height,
              ...(chart.x !== undefined ? { x: chart.x } : {}),
              ...(chart.y !== undefined ? { y: chart.y } : {}),
              ...(chart.pivotSource ? { pivotSource: { ...chart.pivotSource } } : {}),
              ...(chart.pivotOptions ? { pivotOptions: { ...chart.pivotOptions } } : {}),
            }) satisfies SpreadsheetChart,
          (object) => object,
        );
      }
    }
  }
  const chartWorkbook = chartEdit.workbook;
  if (chartWorkbook) {
    const rebound = chartWorkbook[sheets.indexOf(sheet)];
    if (rebound) next = rebound;
  }
  if (
    JSON.stringify(committedPivot ?? null) ===
      JSON.stringify(activePivot ?? null) &&
    JSON.stringify(next.charts ?? []) === JSON.stringify(sheet.charts ?? []) &&
    !chartWorkbook
  ) {
    return { kind: "unchanged" };
  }
  // "Autofit column widths on update": the report's columns are fitted in
  // the same commit as the change that redraws it.
  if (committedPivot?.output && fitSource) {
    next = sheetWithPivotColumnsFitted(
      next,
      planPivotOutput(
        committedPivot,
        pivotComputationMatrix(committedPivot, fitSource),
        {
          rowCount: next.rowCount,
          columnCount: next.columnCount,
          labels: input.labels,
        },
      ),
      committedPivot,
    );
  }
  // PivotTables on one cache refresh together, as Excel refreshes them:
  // the new snapshot reaches every report that names the cache, in one
  // undoable step.
  const sharedKey = committedPivot?.native?.cacheKey;
  const sharedSnapshot = committedPivot?.native?.snapshot;
  if (
    sharedKey &&
    sharedSnapshot &&
    sharedSnapshot.fingerprint !== activePivot?.native?.snapshot.fingerprint
  ) {
    const sharedFields = committedPivot?.native?.fields;
    const follow = (pivot: PivotConfig): PivotConfig => {
      if (
        pivot.native?.cacheKey !== sharedKey ||
        pivot.native.snapshot === sharedSnapshot
      ) {
        return pivot;
      }
      // A Refresh that gained or lost source columns rebuilt the cache's
      // fields: every report on it finds its own fields again by name.
      if (
        sharedFields &&
        sharedFields.length !== pivot.native.fields.length
      ) {
        return rebasePivotNativeSnapshot(pivot, sharedSnapshot, sharedFields) ?? pivot;
      }
      return { ...pivot, native: { ...pivot.native, snapshot: sharedSnapshot } };
    };
    return {
      kind: "workbook",
      sheets: sheets.map((candidate, index) => {
        const base = candidate === sheet ? next : (chartWorkbook?.[index] ?? candidate);
        if (!sheetPivots(base).some((pivot) => pivot.native?.cacheKey === sharedKey)) {
          return base;
        }
        return {
          ...base,
          ...(base.pivot ? { pivot: follow(base.pivot) } : {}),
          ...(base.morePivots ? { morePivots: base.morePivots.map(follow) } : {}),
        };
      }),
    };
  }
  if (chartWorkbook) {
    // Charts of this PivotTable on other sheets change in the same step.
    const ownIndex = sheets.indexOf(sheet);
    return {
      kind: "workbook",
      sheets: chartWorkbook.map((candidate, index) => (index === ownIndex ? next : candidate)),
    };
  }
  return { kind: "sheet", sheet: next };
}

/** The worksheet a PivotTable reads: its host, the sheet its recipe names by
 *  id, or (an imported one) by name; null when that is not exactly one sheet. */
export function pivotSourceSheetOf<S extends { id: string; name: string }>(
  config: PivotConfig,
  host: S,
  sheets: readonly S[],
): S | null {
  if (!config.native && !config.sourceSheetId) return host;
  const matches = sheets.filter((candidate) =>
    config.sourceSheetId
      ? candidate.id === config.sourceSheetId
      : candidate.name === config.native?.sourceSheetName,
  );
  return matches.length === 1 ? matches[0] : null;
}

/** A PivotTable's source as it reads now: shown rows (text cells marked),
 *  typed values, and each column's own number format. */
export interface PivotLiveSource {
  matrix: PivotMatrix;
  scalarMatrix: PivotNativeScalarMatrix;
  columnNumberFormats: (string | undefined)[];
}

/**
 * Live source display values, kept separate from the computation matrix: an
 * imported native PivotTable computes from its explicit cache snapshot, and
 * live cells only drive the dirty indicator until Refresh replaces it. A table
 * source follows its table, so rows typed under it are read at once. A
 * Scenario PivotTable report has no worksheet source: its fields and data are
 * its own cache.
 */
export function pivotLiveSource(
  config: PivotConfig | undefined,
  sourceSheet: Pick<PivotButtonSheet, "id" | "name" | "cells" | "merges" | "dataTables"> | null,
  evaluatorFor: (sheetName: string) => SheetEvaluator,
): PivotLiveSource | null {
  if (config?.cacheSnapshot && pivotIsScenarioReport(config)) {
    const cache = config.cacheSnapshot;
    const matrix = cache.matrix.map((row) => [...row]);
    return {
      matrix: markPivotTextCells(matrix, cache.scalarMatrix),
      scalarMatrix: cache.scalarMatrix,
      columnNumberFormats: cache.fieldIds.map(() => undefined),
    };
  }
  if (!config || !sourceSheet) return null;
  const range = pivotLiveSourceRange(config, sourceSheet.dataTables);
  if (!range) return null;
  const evaluator = evaluatorFor(sourceSheet.name || sourceSheet.id);
  const sourceSpill = buildSpillOverlay(
    sourceSheet.cells,
    coveredCells(normalizeMerges(sourceSheet.merges)),
    evaluator,
  );
  const matrix: string[][] = [];
  const scalarMatrix: PivotNativeScalarMatrix = [];
  for (let r = range.r1; r <= range.r2; r += 1) {
    const row: string[] = [];
    const scalarRow: PivotNativeScalar[] = [];
    for (let c = range.c1; c <= range.c2; c += 1) {
      const scalar =
        r === range.r1
          ? {
              kind: "s" as const,
              raw: resolveCellDisplay(sourceSheet.cells[`${r}:${c}`], r, c, evaluator, sourceSpill),
            }
          : pivotNativeScalarFromLiveCell(sourceSheet.cells[`${r}:${c}`], r, c, evaluator, sourceSpill);
      scalarRow.push(scalar);
      row.push(pivotNativeScalarDisplay(scalar));
    }
    matrix.push(row);
    scalarMatrix.push(scalarRow);
  }
  return {
    matrix: markPivotTextCells(matrix, scalarMatrix),
    scalarMatrix,
    columnNumberFormats: pivotSourceColumnNumberFormats(sourceSheet.cells, range),
  };
}

/** What PivotTable Analyze > Refresh does to one PivotTable. */
export type PivotRefreshPlan =
  /** It reads live cells (no cache snapshot): only its charts are refreshed. */
  | { kind: "live" }
  /** The recipe names a snapshot it no longer carries: nothing to do. */
  | { kind: "none" }
  | { kind: "refused"; refusal: Exclude<PivotRefreshRefusal, "no-snapshot">; cachedWidth: number }
  | { kind: "refreshed"; config: PivotConfig; fieldsChanged: boolean };

/**
 * Refresh of a PivotTable that computes from a cache snapshot: the snapshot is
 * rebuilt over the source as it is now (a column added to the source becomes
 * a new, unplaced field; placed fields stay on their columns), and a table
 * source's stored range catches up with the table. A stored cache of another
 * width than the source range is now (a column inserted inside, or deleted
 * from, a plain-range source) is left out of the normalized recipe, which then
 * reads live cells; Refresh is where that cache is rebuilt. An imported
 * report finds each placed field by name, wherever its column went.
 */
export function pivotRefreshPlan(
  config: PivotConfig,
  rawActive: unknown,
  live: Pick<PivotLiveSource, "matrix" | "scalarMatrix"> | null,
  sourceSheet: Pick<PivotButtonSheet, "dataTables"> | null,
): PivotRefreshPlan {
  const storedSnapshot =
    !config.native && !config.cacheSnapshot
      ? (normalizePivotConfig(rawActive ?? {}).cacheSnapshot ?? null)
      : null;
  if (!config.native && !config.cacheSnapshot && !storedSnapshot) return { kind: "live" };
  const refreshBase: PivotConfig = storedSnapshot
    ? { ...config, cacheSnapshot: storedSnapshot }
    : config;
  const cachedWidth =
    (refreshBase.native?.snapshot ?? refreshBase.cacheSnapshot)?.fieldIds.length ?? 0;
  if (!live) return { kind: "refused", refusal: "source-unreadable", cachedWidth };
  const outcome = refreshPivotSnapshot(refreshBase, live.matrix, live.scalarMatrix);
  if (!outcome.config) {
    return outcome.refusal === "no-snapshot"
      ? { kind: "none" }
      : { kind: "refused", refusal: outcome.refusal, cachedWidth };
  }
  const refreshed = outcome.config;
  const refreshedSnapshot = refreshed.native?.snapshot ?? refreshed.cacheSnapshot;
  const fieldsChanged =
    refreshedSnapshot !== undefined &&
    refreshedSnapshot.fieldIds.length !== cachedWidth;
  const tableRange =
    config.sourceTableId && sourceSheet
      ? pivotLiveSourceRange(config, sourceSheet.dataTables)
      : null;
  return {
    kind: "refreshed",
    config: tableRange ? { ...refreshed, source: pivotRangeToA1(tableRange) } : refreshed,
    fieldsChanged,
  };
}

/**
 * Create PivotTable > New Worksheet: the worksheet holding the new report at
 * A3 (its source named by id, so a rename cannot detach it), sized for the
 * report, its columns fitted, with the report's PivotChart when asked. The
 * caller inserts it in front of the active sheet. Null when the report would
 * not fit any sheet.
 */
export function pivotNewWorksheet<S extends PivotButtonSheet>(
  workbook: readonly S[],
  config: PivotConfig,
  snapshot: { sheetId: string; matrix: PivotMatrix },
  cacheSnapshot: PivotCacheSnapshot,
  sourceRange: PivotRange,
  options: {
    withChart: boolean;
    labels?: Partial<PivotOutputLabels>;
    localizeCaptions?: (next: PivotConfig, fieldName: (colIndex: number) => string) => PivotConfig;
    columnLabel: (index: number) => string;
    /** The new sheet's workbook-level mirrors (defined names and the like). */
    mirrors: Partial<S>;
  },
): { sheet: S; top: PivotRange } | null {
  const planned = pivotConfigForNewSheet(config, snapshot.matrix, snapshot.sheetId);
  if (!planned) return null;
  const { id: targetId, name: targetName } = nextPivotWorksheetIdentity(workbook);
  const sourceWidth = sourceRange.c2 - sourceRange.c1;
  const normalized = normalizePivotConfig({ ...planned.config, cacheSnapshot }, sourceWidth);
  const localizedPivot = options.localizeCaptions
    ? options.localizeCaptions(
        normalized,
        (colIndex) => headerLabel(snapshot.matrix, colIndex).trim() || options.columnLabel(colIndex),
      )
    : normalized;
  // A PivotChart names its PivotTable (Excel's next PivotTableN).
  const targetPivot = options.withChart && !pivotTableName(localizedPivot)
    ? { ...localizedPivot, name: nextPivotTableName(workbook) }
    : localizedPivot;
  const bounds = { rowCount: planned.rowCount, columnCount: planned.columnCount, labels: options.labels };
  const blank = {
    ...options.mirrors,
    id: targetId,
    name: targetName,
    rowCount: planned.rowCount,
    columnCount: planned.columnCount,
    cells: {},
    pivot: targetPivot,
  } as S;
  const fittedSheet = sheetWithPivotColumnsFitted(
    blank,
    planPivotOutput(targetPivot, pivotComputationMatrix(targetPivot, snapshot.matrix), bounds),
    targetPivot,
  );
  const targetPlan = options.withChart
    ? planPivotOutput(targetPivot, pivotComputationMatrix(targetPivot, snapshot.matrix), bounds)
    : null;
  const sheet: S =
    options.withChart && targetPlan?.output
      ? {
          ...fittedSheet,
          charts: [
            newPivotChartRecord(
              fittedSheet,
              { sheetId: targetId, pivotName: pivotTableName(targetPivot) },
              targetPlan.chart,
              targetPlan.output,
              "column",
            ),
          ],
        }
      : fittedSheet;
  return { sheet, top: planned.top };
}

/**
 * PivotTable Analyze > Clear > Clear All: every field leaves the layout; the
 * PivotTable keeps its source, name, place, cache and native identity. Null
 * when it has no field to clear.
 */
export function pivotClearedAll(config: PivotConfig): PivotConfig | null {
  if (
    !(
      config.rows.length ||
      config.columns.length ||
      config.values.length ||
      config.filters.length ||
      config.grouping.length ||
      config.calculatedFields.length
    )
  ) {
    return null;
  }
  return {
    source: config.source,
    ...(config.name ? { name: config.name } : {}),
    ...(config.output ? { output: config.output } : {}),
    rows: [],
    columns: [],
    values: [],
    filters: [],
    grouping: [],
    calculatedFields: [],
    ...(config.cacheSnapshot ? { cacheSnapshot: config.cacheSnapshot } : {}),
    ...(config.native ? { native: config.native } : {}),
  };
}
