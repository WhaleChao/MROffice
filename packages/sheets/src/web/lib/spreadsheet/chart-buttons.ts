/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Charts and the Chart Design / Format tabs' buttons as pure
 * functions over a sheet's chart records, moved out of SpreadsheetWindow so
 * the window's buttons and the spreadsheet editor commands
 * (lib/editor-commands/spreadsheet) run the same code; behaviour is unchanged.
 *
 * A chart is a record born here (`sheet.charts[]`, "native") or an imported
 * chart object (`sheet.objects[]` of type chart, "drawing"). Every Chart
 * Design button computes the chart's next model (`ChartBlockAttrsV2`) from
 * its current one with the shared chart commands Docs and the deck use; the
 * caller stores it (`sheetWithChartModel`). What the buttons read of the
 * grid -- the chart's live data, displayed values -- comes in as
 * `SheetChartData`, so the window passes what it draws and a headless caller
 * what its recalculation stamped. Receipts, protection, undo, notices and the
 * contextual tab stay with the caller.
 */

import {
  chartTakesColorPalette,
  chartFormatFollowingDataEdit,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  chartSupportsAxisTitle,
  clampChartText,
  chartHasAxis,
  isChartDisplayBlanksAs,
  MAX_CHART_CATEGORY_LEVELS,
  chartWithHistogramBins,
  normalizeChartHistogramBins,
  officeInsertedChartLegendAndLabels,
  withOfficeInsertedChartFormat,
  type ChartFamily,
  type ChartSeriesMarker,
  type ChartSubtype,
  type ComboRenderAs,
} from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { pointsFromRows } from "@mrnote/core/web/components/notebook/contents/MiniChart.tsx";
import { chartWithColorPalette } from "@mrnote/core/web/lib/office-charts/chart-color-palettes.ts";
import { officeChartDateAxisTimes } from "@mrnote/core/web/lib/office-charts/axis-options.ts";
import {
  chartWithAxis,
  chartWithDataLabels,
  chartWithDataTable,
  chartWithElementIntent,
  chartWithErrorBars,
  chartWithGridlines,
  chartWithSeriesMarker,
  chartWithLine,
  chartWithoutLines,
  chartWithSeriesTrendline,
  chartWithTrendline,
  chartWithUpDownBars,
  type ChartDataLabelOverridePatch,
  type ChartDataLabelPlacement,
  type ChartDataLabelTarget,
} from "@mrnote/core/web/lib/office-charts/chart-elements.ts";
import { CHART_STYLE_PRESET_FOR_C_STYLE } from "@mrnote/core/web/lib/office-charts/chart-style-presets.generated.ts";
import {
  OFFICE_CHART_EX_FAMILIES,
  officeChartCombinationRefusal,
  officeChartCombinationRows,
  officeChartWithType,
} from "@mrnote/core/web/lib/office-charts/chart-type.ts";
import { officeChartWithQuickLayout } from "@mrnote/core/web/lib/office-charts/quick-layout.ts";
import {
  buildChartSeries,
  buildChartSeriesFromRefs,
  nativeChartHeaderRow,
  nativeChartRangeLayout,
  chartCategoryLevelsForRows,
  nativeChartOuterCategoryCells,
  isSpreadsheetChartLegendMenuPosition,
  MAX_CHART_TITLE_LENGTH,
  normalizeChartSeriesRefs,
  normalizeErrorBars,
  normalizeSecondaryAxisSeries,
  normalizeSpreadsheetChartColorPalette,
  normalizeSpreadsheetChartType,
  normalizeSpreadsheetChartV2,
  normalizeTrendline,
  SPREADSHEET_CHART_STYLE_GALLERY,
  SPREADSHEET_LINE_MARKER_SYMBOLS,
  spreadsheetTrendlineContract,
  type SpreadsheetChart,
  type SpreadsheetChartColorPaletteId,
  type SpreadsheetChartSeriesRef,
  type SpreadsheetChartType,
  type SpreadsheetPivotChartSource,
} from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { chartRangeToA1, normalizeChartRange, type ChartRange } from "@mrnote/core/web/lib/spreadsheet/chart-config.ts";
import {
  drawingChartDefaultSeriesInColumns,
  drawingChartRangesOf,
  drawingChartSeriesRefsForRange,
  drawingChartSheetOf,
  drawingChartSheetPrefix,
  drawingChartSourceOf,
  drawingChartTakesDataRange,
} from "@mrnote/core/web/lib/spreadsheet/drawing-chart-series.ts";
import { spreadsheetChartLegacyDisplay, type SpreadsheetChartFormat } from "@mrnote/core/web/lib/spreadsheet/chart-format.ts";
import { CHART_QUICK_LAYOUT_IDS as NATIVE_CHART_QUICK_LAYOUT_IDS, type ChartQuickLayoutId as NativeChartQuickLayoutId } from "@mrnote/core/web/lib/spreadsheet/chart-quick-layout.ts";
import { chartMatrixSupportsType as classicChartMatrixSupportsType, type ChartMatrixProfile } from "./chart-recommendations";
import { spreadsheetChartTitle } from "@mrnote/core/web/lib/spreadsheet/chart-title.ts";
import {
  spreadsheetChartFamily,
  spreadsheetChartLeafOf,
  spreadsheetChartTypeForLeaf,
  spreadsheetChartTypeSupportsGridlines,
  spreadsheetChartV2Drawn,
} from "@mrnote/core/web/lib/spreadsheet/chart-v2-binding.ts";
import {
  reconcileSpreadsheetChartV2,
  sheetChartErrorBarsOfSpec,
  sheetChartRetargeted,
  spreadsheetChartBindingOf,
  spreadsheetChartBornWithModel,
  spreadsheetChartEditRefusal,
  spreadsheetChartV2Reading,
  withSpreadsheetChartV2,
  type SpreadsheetChartBinding,
  type SpreadsheetChartLike,
  type SpreadsheetChartV2Live,
} from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";
import {
  DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
  DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
  offsetOfLine,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { isSpreadsheetPivotChartType } from "@mrnote/core/web/lib/spreadsheet/pivot-chart.ts";

/**
 * Whether an imported chart's object carries the projection's own reading of
 * one analysis element (`nativeChartAnalysis`, stated by the server when the
 * part holds nothing this editor cannot describe). Only then can the ribbon
 * offer the element: the save tells a removal from an older page's silence by
 * that list, and a part with a polynomial trendline or custom error amounts
 * keeps them untouched.
 */
export function importedChartStatesAnalysis(
  chart: { nativeChartAnalysis?: unknown },
  key: "trendline" | "errorBars" | "secondaryAxisSeries",
): boolean {
  return (
    Array.isArray(chart.nativeChartAnalysis) &&
    chart.nativeChartAnalysis.includes(key)
  );
}

// Direct chart authoring must never offer a kind that the current native XLSX
// serializer omits or reopens as a different chart. Column and Bar are distinct
// end-to-end; XY/Bubble and Combo consume the shared bounded series-ref wire.
// Histogram and chartEx kinds remain preserve-only.
export const NATIVE_XLSX_CHART_AUTHORING_TYPES = new Set<SpreadsheetChartType>([
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "radar",
  "scatter",
  "bubble",
  "combo",
  "surface",
  "stock", "histogram", "pareto", "box-and-whisker", "waterfall", "treemap", "sunburst", "funnel", "map",
]);

export const EXTENDED_XLSX_CHART_TYPES = new Set<SpreadsheetChartType>(["histogram","pareto","box-and-whisker","waterfall","treemap","sunburst","funnel","map"]);

export function chartMatrixSupportsType(matrix:readonly (readonly string[])[],type:SpreadsheetChartType,seriesInColumns=true,chartSubtype?:string,profile?:ChartMatrixProfile):boolean {
  if(classicChartMatrixSupportsType(matrix,type,seriesInColumns,profile))return true;
  if(!EXTENDED_XLSX_CHART_TYPES.has(type)&&type!=="stock")return false;
  const rows=seriesInColumns?matrix.map(row=>[...row]):Array.from({length:Math.max(0,...matrix.map(row=>row.length))},(_,col)=>matrix.map(row=>row[col]??""));
  let width=rows[0]?.length??0;
  // A treemap or a sunburst over several text columns and a value column:
  // the columns before the one beside the values are its outer branches
  // (insertChartSelectionColumns keeps them as levels); the leaf column and
  // the values are what must plot.
  if((type==="treemap"||type==="sunburst")&&width>2){
    const numericAt=(col:number)=>rows.slice(1).some(row=>{const raw=(row[col]??"").replace(/[$,%\s]/g,"");return raw!==""&&Number.isFinite(Number(raw));});
    const first=Array.from({length:width},(_,col)=>col).find(col=>col>0&&numericAt(col));
    if(first===undefined)return false;
    const leaf=rows.map(row=>[row[first-1]??"",row[first]??""]);
    return spreadsheetChartV2Drawn({type,rows:leaf,...(chartSubtype?{chartSubtype}:{})})!==null;
  }
  // A histogram bins one column of numbers on its own, as a box and whisker
  // chart does (Excel inserts both over a bare column). A waterfall or a
  // funnel over a bare column stays out until the workbook writer exports
  // it: its one-column branch knows only the binned kinds, and the chart
  // would be left out of the .xlsx.
  if(type==="histogram"&&width===1)width=2;
  if(type!=="stock"&&type!=="box-and-whisker"&&width!==2)return false;
  // A subtype can need MORE columns than its family's default (a stock chart's
  // OHLC vs VOHLC series order), so the gate validates the exact leaf.
  return spreadsheetChartV2Drawn({type,rows,...(chartSubtype?{chartSubtype}:{})})!==null;
}

export function canAuthorNativeXlsxChart(type: SpreadsheetChartType): boolean {
  return NATIVE_XLSX_CHART_AUTHORING_TYPES.has(type);
}

/**
 * A filled map is drawn on geography, not on axes: Excel's own Chart Design
 * tab offers a map chart a title, data labels, a legend and its map series
 * options, and no axes or gridlines (Microsoft Support, "Create a Map chart in
 * Excel"). The parts Excel writes agree — a filled map's format states no axis
 * or gridline key at all — and so does the shared chart contract's
 * `CHART_AXIS_VISIBILITY_FAMILIES`.
 */
export function spreadsheetChartTypeHasAxes(type: SpreadsheetChartType): boolean {
  return ![
    "pie",
    "doughnut",
    "treemap",
    "sunburst",
    "map",
  ].includes(type);
}

/** Gridlines follow the shared chart contract, not the axes alone: a funnel has
 * its category axis and Excel's Chart Elements menu offers it no gridlines. */
export function spreadsheetChartTypeHasGridlines(type: SpreadsheetChartType): boolean {
  return (
    spreadsheetChartTypeHasAxes(type) &&
    spreadsheetChartTypeSupportsGridlines(type)
  );
}

/** Excel offers no data labels on a surface chart: its points are a coloured
 * band across a matrix, not marks of their own. */
export function spreadsheetChartTypeSupportsDataLabels(
  type: SpreadsheetChartType,
): boolean {
  return type !== "surface";
}

export function spreadsheetChartTypeSupportsDataTable(
  type: SpreadsheetChartType,
): boolean {
  return ["column", "stacked-column", "bar", "line", "area", "combo"].includes(
    type,
  );
}

export type ChartQuickLayoutId = NativeChartQuickLayoutId;

export const CHART_QUICK_LAYOUT_IDS = NATIVE_CHART_QUICK_LAYOUT_IDS;

/**
 * What a chart's element switches are actually showing right now.
 *
 * An imported chart's real gridline, axis and data-table state lives in
 * `chartFormat`, read from its own part; the plain `showGridlines` /
 * `showAxisLabels` / `showDataTable` keys are written only by an edit made
 * here. Comparing a menu pick with the bare key alone made every switch
 * one-way on an imported chart — Office draws no gridlines, the key says
 * nothing, so "Gridlines on" looked like a no-op and was dropped.
 */
export function chartElementDisplayState(
  chart: Readonly<{
    chartFormat?: SpreadsheetChartFormat;
    showAxisLabels?: boolean;
    showGridlines?: boolean;
    showDataTable?: boolean;
  }>,
): { showAxisLabels: boolean; showGridlines: boolean; showDataTable: boolean } {
  const actual = { ...chart, ...spreadsheetChartLegacyDisplay(chart.chartFormat) };
  return {
    showAxisLabels: actual.showAxisLabels !== false,
    showGridlines: actual.showGridlines !== false,
    showDataTable: actual.showDataTable === true,
  };
}

/** "x:true" / "y:false": one axis of Add Chart Element's Axes or Gridlines. */
export function spreadsheetChartAxisPick(
  parameter: string,
): { axis: "x" | "y"; on: boolean } | null {
  const match = /^(x|y):(true|false)$/.exec(parameter);
  return match
    ? { axis: match[1] as "x" | "y", on: match[2] === "true" }
    : null;
}

export function chartStyleIdFromGallery(
  galleryId: string,
): number | null {
  const styleId = Number(galleryId);
  return SPREADSHEET_CHART_STYLE_GALLERY.some(
    (preset) => preset.id === styleId,
  )
    ? styleId
    : null;
}

export function chartColorPaletteIdFromGallery(
  galleryId: string,
): SpreadsheetChartColorPaletteId | null {
  return normalizeSpreadsheetChartColorPalette(galleryId) ?? null;
}

/**
 * Change Colors on a sheet chart's model: the palette repaints every series'
 * fill -- a gradient fill included -- and keeps its outline and effect. The
 * commit and the menu's hover preview both run it.
 */
export function sheetChartWithChangedColors(attrs: ChartBlockAttrsV2, paletteId: SpreadsheetChartColorPaletteId): ChartBlockAttrsV2 | null {
  const format = { ...(attrs.format ?? {}) };
  for (const key of ["seriesStyles", "sliceStyles"] as const) {
    const entries = format[key] as readonly Record<string, unknown>[] | undefined;
    if (!entries) continue;
    (format as Record<string, unknown>)[key] = entries.map((entry) => {
      const paint = entry.shapePaint as Record<string, unknown> | undefined;
      if (!paint?.fillGradient) return entry;
      const { fillGradient: _gradient, ...rest } = paint;
      return { ...entry, shapePaint: rest };
    });
  }
  return chartWithColorPalette({ ...attrs, format }, paletteId);
}


type Json = Record<string, unknown>;

/** The stored fields of a chart record the buttons read. */
interface SheetChartFields {
  type?: SpreadsheetChartType;
  chartType?: unknown;
  range?: ChartRange;
  sourceSheetId?: string;
  seriesInColumns?: boolean;
  noCategories?: true;
  noSeriesNames?: true;
  seriesRefs?: unknown;
  sourceRefError?: boolean;
  pivotSource?: SpreadsheetPivotChartSource;
  showDataLabels?: boolean;
  showLegend?: boolean;
  axisTitleX?: string;
  axisTitleY?: string;
  chartFormat?: SpreadsheetChartFormat;
  showAxisLabels?: boolean;
  showGridlines?: boolean;
  showDataTable?: boolean;
  nativeChartAnalysis?: unknown;
  nativeChartStyleBlocker?: string;
  nativeChartLineBlocker?: string;
  nativeChartColorBlocker?: string;
  nativeChartColorEditBlocker?: string;
  nativeChartLegendPositionBlocker?: string;
}

/** One chart on its sheet: the sheet, the list holding the record, the record. */
export interface SheetChartTarget {
  readonly kind: "native" | "drawing";
  readonly sheet: object;
  readonly chartCollection: readonly object[];
  readonly chart: object;
  readonly chartId: string;
}

/** What the chart buttons read of the grid beside the record. */
export interface SheetChartData {
  /** The live data of a chart record (retargeted or not), as it draws. */
  liveOf(kind: "native" | "drawing", chart: object): SpreadsheetChartV2Live | null;
  sheetIdByName(name: string): string | undefined;
  /** The name of the chart's own sheet. */
  readonly sheetName: string;
  /** The displayed values of a rectangle of a chart's source sheet. */
  matrixForRange(range: ChartRange, sourceSheetId?: string): string[][];
  /** An imported chart's rows as it draws them. */
  importedRows(chart: object): string[][];
  /** A PivotChart's report rows. */
  pivotRows(source: SpreadsheetPivotChartSource): string[][];
  /** A cell of the chart's own sheet as displayed, and as a chart value. */
  cellDisplay(row: number, col: number): string;
  cellValue(row: number, col: number): string;
  /** The placeholder words Quick Layout writes (Office's English by default). */
  words?: { chartTitle: string; axisTitle: string };
}

/** A chart's next model, its binding, and record flags the write settles. */
/** Record fields a write removes: a #REF! source repaired, a label line the
 * chart's range now has. */
export type SheetChartSettledField = "sourceRefError" | "noCategories" | "noSeriesNames";

export interface SheetChartWrite {
  attrs: ChartBlockAttrsV2;
  binding?: SpreadsheetChartBinding;
  settled?: readonly SheetChartSettledField[];
  /** Record fields the write sets beside the model (XLS-A072: the chart's
   *  `chartVisibility`, Show data in hidden rows and columns). */
  record?: Readonly<Record<string, unknown>>;
}

/** A Chart Design button: the chart, the ribbon's parameter, the grid data. */
export type SheetChartButton = (target: SheetChartTarget, parameter: string, data: SheetChartData) => SheetChartWrite | null;

function fieldsOf(target: SheetChartTarget): SheetChartFields {
  return target.chart as SheetChartFields;
}

function chartTypeOf(target: SheetChartTarget): SpreadsheetChartType {
  const chart = fieldsOf(target);
  return target.kind === "drawing" ? normalizeSpreadsheetChartType(chart.chartType) : (chart.type as SpreadsheetChartType);
}

/** A chart whose data the chart itself states (an imported XY chart over
 * text X values, literal bubble points): its data is not re-read from cells. */
export function sheetChartDataLiteral(binding?: SpreadsheetChartBinding): boolean {
  return (binding as { literal?: unknown } | undefined)?.literal === true;
}

/** The chart's current model and binding: its stored model (an imported
 * chart's reading of its part included) with the data cache refreshed from
 * the cells, or, for a record not yet carrying one, its reading. */
export function sheetChartModelOf(
  target: SheetChartTarget,
  data: SheetChartData,
): { attrs: ChartBlockAttrsV2; binding?: SpreadsheetChartBinding } | null {
  try {
    const like = target.chart as unknown as SpreadsheetChartLike;
    const live = data.liveOf(target.kind, target.chart);
    const current = reconcileSpreadsheetChartV2(like, live, { sheetIdByName: (name) => data.sheetIdByName(name) });
    const refreshed = current.chartV2 as ChartBlockAttrsV2 | undefined;
    const stored = normalizeSpreadsheetChartV2(like.chartV2);
    const binding = spreadsheetChartBindingOf(current, { sheetIdByName: (name) => data.sheetIdByName(name) });
    // An older imported object bound to no cells (its data is its own
    // cache) is read as it draws, with no binding.
    const attrs = refreshed && validateChartBlockAttrsV2(refreshed).valid ? refreshed : stored ?? (live && !binding ? spreadsheetChartV2Reading(like, live) : null);
    return attrs ? { attrs, binding } : null;
  } catch {
    return null;
  }
}

/** The sheet's reading of the chart record with `patch` applied (another
 * kind, another source or orientation), its data re-read for it. */
export function sheetChartReadingOf(
  target: SheetChartTarget,
  data: SheetChartData,
  patch: Record<string, unknown>,
): ChartBlockAttrsV2 | null {
  try {
    const retargeted = { ...target.chart, ...patch };
    const live = data.liveOf(target.kind, retargeted);
    return live ? spreadsheetChartV2Reading(retargeted as unknown as SpreadsheetChartLike, live) : null;
  } catch {
    return null;
  }
}

/** The sheet with the chart storing `attrs` (and the flat fields derived
 * from it); null when the chart's content is not the editor's to change or
 * the model is not valid. */
export function sheetWithChartModel<S extends object>(
  target: SheetChartTarget & { sheet: S },
  attrs: ChartBlockAttrsV2,
  binding?: SpreadsheetChartBinding,
  // Record flags the write settles (a #REF! source repaired by Select Data).
  settled: readonly SheetChartSettledField[] = [],
  recordFields: Readonly<Record<string, unknown>> = {},
): S | null {
  if (spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike) || !validateChartBlockAttrsV2(attrs).valid) return null;
  const field = target.kind === "native" ? "charts" : "objects";
  const record = { ...target.chart, ...recordFields } as Json;
  for (const key of settled) delete record[key];
  const next = withSpreadsheetChartV2(record as SpreadsheetChartLike, attrs, binding);
  return { ...target.sheet, [field]: target.chartCollection.map((item) => (item === target.chart ? next : item)) } as S;
}

/** A command on the chart's current model (the one shared contract): its
 * next model and binding, or null when it changes nothing. */
export function sheetChartModelEdited(
  target: SheetChartTarget,
  data: SheetChartData,
  command: (attrs: ChartBlockAttrsV2, binding?: SpreadsheetChartBinding) => ChartBlockAttrsV2 | null,
  options: { binding?: SpreadsheetChartBinding } = {},
): SheetChartWrite | null {
  if (spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike)) return null;
  const model = sheetChartModelOf(target, data);
  if (!model) return null;
  const next = command(model.attrs, model.binding);
  if (!next || (JSON.stringify(next) === JSON.stringify(model.attrs) && !options.binding)) return null;
  return { attrs: next, binding: options.binding ?? model.binding };
}

/** Whether Change Chart Type may turn the chart into `targetType` over the
 * data it reads. */
export function chartDesignTypeSupportsSource(
  kind: "native" | "drawing",
  chart: object,
  targetType: SpreadsheetChartType,
  data: SheetChartData,
): boolean {
  const record = chart as SheetChartFields;
  const currentType =
    kind === "drawing"
      ? normalizeSpreadsheetChartType(record.chartType)
      : (record.type as SpreadsheetChartType);
  if (targetType === currentType) return true;
  if (kind === "drawing") {
    if(EXTENDED_XLSX_CHART_TYPES.has(currentType)||EXTENDED_XLSX_CHART_TYPES.has(targetType)||targetType==="stock"||targetType==="combo"){
      const refs=normalizeChartSeriesRefs(record.seriesRefs);
      if(!refs||refs.some(ref=>ref.xValues||ref.implicitX||ref.bubbleSizes||ref.axis==="secondary"))return false;
      if(EXTENDED_XLSX_CHART_TYPES.has(targetType)&&targetType!=="box-and-whisker"&&refs.length!==1)return false;
      const rows=data.importedRows(chart);
      return chartMatrixSupportsType(rows,targetType);
    }
    // Column <-> bar remains a metadata-only edit and therefore does not
    // need the richer regeneration receipt below.
    if (
      (currentType === "column" || currentType === "bar") &&
      (targetType === "column" || targetType === "bar")
    ) {
      return true;
    }
    // The classic kinds the original-package writer changes an imported
    // chart between, keeping its series, title and legend (the plot is
    // rebuilt over the same cells; API `chart_desktop`).
    const safeClassicTypes = new Set<SpreadsheetChartType>([
      "column",
      "bar",
      "stacked-column",
      "stacked-bar",
      "line",
      "area",
      "pie",
      "doughnut",
      "radar",
      "scatter",
    ]);
    if (
      !safeClassicTypes.has(currentType) ||
      !safeClassicTypes.has(targetType)
    ) {
      return false;
    }
    const refs = normalizeChartSeriesRefs(record.seriesRefs);
    if (
      !refs ||
      refs.some(
        (recipe) =>
          recipe.implicitX === true ||
          Boolean(recipe.bubbleSizes) ||
          recipe.axis === "secondary" ||
          (!recipe.categories && !recipe.xValues),
      ) ||
      (targetType === "pie" && refs.length !== 1)
    ) {
      return false;
    }
    const rows = buildChartSeriesFromRefs(
      refs,
      currentType,
      data.sheetName,
      (row, col) => data.cellDisplay(row, col),
      (row, col) => data.cellValue(row, col),
    );
    if (rows.length === 0 || pointsFromRows(rows).length === 0) {
      return false;
    }
    if (targetType === "scatter") {
      return rows.every((row) => Number.isFinite(Number(row[0])));
    }
    return true;
  }
  if (record.pivotSource) {
    return chartMatrixSupportsType(data.pivotRows(record.pivotSource), targetType, record.seriesInColumns !== false);
  }
  const gate = nativeChartGateMatrix(data.matrixForRange(record.range as ChartRange, record.sourceSheetId), record);
  return chartMatrixSupportsType(gate.matrix, targetType, gate.seriesInColumns);
}

/**
 * The table a range-bound chart's type gates read over `matrix` (its range's
 * cells as shown, row 0 column 0 its top-left): a range that states its
 * label lines is laid by them, a label column first and a row of names on
 * top, series in columns; any other is read as it stands, its orientation
 * the chart's.
 */
export function nativeChartGateMatrix(
  matrix: readonly (readonly string[])[],
  fields: { seriesInColumns?: boolean; noCategories?: true; noSeriesNames?: true },
): { matrix: readonly (readonly string[])[]; seriesInColumns: boolean } {
  const width = Math.max(0, ...matrix.map((row) => row.length));
  if (!(fields.noCategories || fields.noSeriesNames) || !matrix.length || !width) {
    return { matrix, seriesInColumns: fields.seriesInColumns !== false };
  }
  const cell = (row: number, col: number) => matrix[row]?.[col] ?? "";
  const range = { r1: 0, c1: 0, r2: matrix.length - 1, c2: width - 1 };
  const chart = {
    range,
    ...(fields.seriesInColumns === false ? { seriesInColumns: false } : {}),
    ...(fields.noCategories ? { noCategories: true as const } : {}),
    ...(fields.noSeriesNames ? { noSeriesNames: true as const } : {}),
  };
  const rows = buildChartSeries(chart, cell, cell);
  const header = Array.from({ length: nativeChartRangeLayout(chart).seriesLines.length + 1 }, () => "");
  return { matrix: fields.noSeriesNames ? [header, ...rows] : rows, seriesInColumns: true };
}

/**
 * Insert > Charts on a selection: the columns that chart plots. A text
 * column after the first is skipped up to the first numeric one, and
 * trailing text columns are left out (XY charts take the selection as it
 * is).
 */
/** The kinds whose category axis Excel draws in levels (`c:multiLvlStrRef`). */
const MULTI_LEVEL_CATEGORY_TYPES = new Set<SpreadsheetChartType>([
  "column", "stacked-column", "bar", "stacked-bar", "line", "area", "radar", "combo",
  // A treemap's and a sunburst's text columns are its branches, outermost
  // first (Region, Country, Sales), as Excel reads them.
  "treemap", "sunburst",
]);

export function insertChartSelectionColumns(
  rect: ChartRange,
  type: SpreadsheetChartType,
  cellValue: (row: number, col: number) => string,
): { c1: number; c2: number; outer: number } {
  let chartC1 = rect.c1;
  let chartC2 = rect.c2;
  let outer = 0;
  if (type !== "scatter" && type !== "bubble" && rect.c2 > rect.c1) {
    const numericColumn = (col: number): boolean => {
      for (let row = rect.r1; row <= rect.r2; row += 1) {
        const raw = cellValue(row, col);
        if (raw !== "" && Number.isFinite(Number(raw))) return true;
      }
      return false;
    };
    let firstNumeric = -1;
    for (let col = rect.c1 + 1; col <= rect.c2; col += 1) {
      if (numericColumn(col)) {
        firstNumeric = col;
        break;
      }
    }
    if (firstNumeric > rect.c1 + 1) {
      chartC1 = firstNumeric - 1;
      // XLS-A107: the text columns before the one next to the numbers are
      // the outer levels of a multi-level category axis (A=city, B=owner,
      // C=amount: Beijing / Wang Wei under each column), as Excel inserts it,
      // on the kinds whose category axis takes levels.
      if (MULTI_LEVEL_CATEGORY_TYPES.has(type)) {
        outer = Math.min(chartC1 - rect.c1, MAX_CHART_CATEGORY_LEVELS - 1);
      }
    }
    if (firstNumeric >= 0) {
      while (chartC2 > firstNumeric && !numericColumn(chartC2)) chartC2 -= 1;
    }
    // A pie keeps every series of the selection, as Excel's Insert does: it
    // plots the first and holds the rest (another chart type shows them; a
    
    // three value columns has three series.
  }
  return { c1: chartC1, c2: chartC2, outer };
}

/** The kinds Insert Chart lays over a selection by Excel's label rule
 * (`excelChartRangeLabelLines`): a category axis, or a sample of values.
 * An XY chart takes its X column as it is, a stock chart its fixed columns,
 * a hierarchy its text columns as branches, a map its region names. */
const LABEL_RULE_TYPES = new Set<SpreadsheetChartType>([
  "column", "stacked-column", "bar", "stacked-bar", "line", "area", "pie", "doughnut", "radar", "combo", "surface",
  "histogram", "pareto", "box-and-whisker", "waterfall", "funnel",
]);

/** Whether Insert Chart and Select Data lay a range of `type` by Excel's
 * label rule. */
export function insertChartTypeTakesLabelRule(type: SpreadsheetChartType): boolean {
  return LABEL_RULE_TYPES.has(type);
}




const ONE_SERIES_TYPES = new Set<SpreadsheetChartType>(["histogram", "pareto", "waterfall", "funnel"]);

/** A cell text that reads as a number (a separator, a currency or a percent
 * sign allowed), as the chart readings take one. */
function chartNumericText(text: string): boolean {
  const stripped = text.trim().replace(/[$,%\s]/g, "");
  return stripped !== "" && Number.isFinite(Number(stripped));
}


















export function excelChartRangeLabelLines(
  rect: ChartRange,
  value: (row: number, col: number) => string,
  display: (row: number, col: number) => string = value,
): { firstRow: boolean; firstColumn: boolean } {
  const label = (row: number, col: number): boolean => {
    const raw = String(value(row, col) ?? "").trim();
    const shown = String(display(row, col) ?? "").trim();
    if (raw === "" && shown === "") return false;
    if (!chartNumericText(raw === "" ? shown : raw)) return true;
    
    return !chartNumericText(shown) && officeChartDateAxisTimes([shown]) !== null;
  };
  const blank = (row: number, col: number) =>
    String(value(row, col) ?? "").trim() === "" && String(display(row, col) ?? "").trim() === "";
  const corner = blank(rect.r1, rect.c1);
  // The last row read from the left, the last column from the top: whether
  // the first filled cell is a label.
  const leadsWithLabel = (cells: readonly [number, number][]) => {
    const first = cells.find(([row, col]) => !blank(row, col));
    return first !== undefined && label(first[0], first[1]);
  };
  const lastRow = Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, index): [number, number] => [rect.r2, rect.c1 + index]);
  const lastColumn = Array.from({ length: rect.r2 - rect.r1 + 1 }, (_, index): [number, number] => [rect.r1 + index, rect.c2]);
  return {
    firstRow: rect.r2 > rect.r1 && (corner || leadsWithLabel(lastColumn)),
    firstColumn: rect.c2 > rect.c1 && (corner || leadsWithLabel(lastRow)),
  };
}

/**
 * The layout fields a chart over `range` states for the label lines Excel
 * reads there (`excelChartRangeLabelLines`), laid by columns or by rows.
 */
export function chartRangeLabelFields(
  lines: { firstRow: boolean; firstColumn: boolean },
  seriesInColumns: boolean,
): { noCategories?: true; noSeriesNames?: true } {
  const categories = seriesInColumns ? lines.firstColumn : lines.firstRow;
  const names = seriesInColumns ? lines.firstRow : lines.firstColumn;
  return { ...(categories ? {} : { noCategories: true as const }), ...(names ? {} : { noSeriesNames: true as const }) };
}

/** Where Insert Chart lays a selection: the range the chart binds, which
 * way its series run and which of its edge lines are labels, and the text
 * columns left of it that are the outer levels of its category axis. */
export interface InsertChartSelectionLayout {
  readonly range: ChartRange;
  readonly seriesInColumns: boolean;
  readonly noCategories?: true;
  readonly noSeriesNames?: true;
  readonly outer: number;
}

/**
 * How Insert Chart lays a selection as `type`, as Excel does: the label
 * lines by Excel's rule, the series in columns when the data (the selection
 * without its label lines) has more rows than columns and in rows otherwise
 * (a square block too), and -- for the kinds that plot one series -- the
 * selection's first series alone. Text columns between the first column and
 * the numbers are outer category levels (XLS-A107); trailing text columns are
 * left out. An XY, stock, hierarchy or map chart takes the selection's
 * columns as it always did, its first column its labels.
 */
export function insertChartSelectionLayout(
  rect: ChartRange,
  type: SpreadsheetChartType,
  value: (row: number, col: number) => string,
  display: (row: number, col: number) => string = value,
): InsertChartSelectionLayout {
  if (type === "scatter") {
    // An X Y chart's series run as a classic chart's do (Excel for Mac,
    
    // when the data -- the selection without its label lines -- has more
    // rows than columns, else in rows, the first line (a label line, or the
    // first line of numbers) being X. x | y1 | y2 over three rows plots each
    // row against x, y1, y2 read as 1, 2, 3.
    const lines = excelChartRangeLabelLines(rect, value, display);
    const dataRows = rect.r2 - rect.r1 + 1 - (lines.firstRow ? 1 : 0);
    const dataColumns = rect.c2 - rect.c1 + 1 - (lines.firstColumn ? 1 : 0);
    return { range: { ...rect }, seriesInColumns: dataRows > dataColumns, outer: 0 };
  }
  if (!LABEL_RULE_TYPES.has(type)) {
    const { c1, c2, outer } = insertChartSelectionColumns(rect, type, value);
    const range = { r1: rect.r1, c1, r2: rect.r2, c2 };
    // A tree map, a sunburst and a stock chart read their first row as the
    // series' names by Excel's label rule: "A 店" / 50 or 门店 / 2024 on top
    // is a point, every row plotted, the series named by its place (Excel
    
    // (source outlines, Select Data) and the file say the same.
    if (
      (type === "treemap" || type === "sunburst" || type === "stock") &&
      rect.r2 > rect.r1 &&
      !excelChartRangeLabelLines(rect, value, display).firstRow
    ) {
      return { range, seriesInColumns: true, noSeriesNames: true, outer };
    }
    return { range, seriesInColumns: true, outer };
  }
  const lines = excelChartRangeLabelLines(rect, value, display);
  let range: ChartRange = { ...rect };
  let outer = 0;
  if (lines.firstColumn) {
    const columns = insertChartSelectionColumns(rect, type, value);
    range = { ...range, c1: columns.c1, c2: columns.c2 };
    outer = columns.outer;
  }
  const dataRows = range.r2 - range.r1 + 1 - (lines.firstRow ? 1 : 0);
  const dataColumns = range.c2 - range.c1 + 1 - (lines.firstColumn ? 1 : 0);
  const seriesInColumns = dataRows > dataColumns;
  const fields = chartRangeLabelFields(lines, seriesInColumns);
  if (ONE_SERIES_TYPES.has(type)) {
    // The first series alone: the label line, then the first value line.
    const step = fields.noCategories ? 0 : 1;
    range = seriesInColumns
      ? { ...range, c2: Math.min(range.c2, range.c1 + step) }
      : { ...range, r2: Math.min(range.r2, range.r1 + step) };
  }
  // The outer levels sit beside a category column, under a row of names.
  if (!seriesInColumns || fields.noCategories || fields.noSeriesNames) outer = 0;
  return { range, seriesInColumns, ...fields, outer };
}

/**
 * The table Insert Chart's type gates read for a selection laid as `type`
 * (`insertChartSelectionLayout`): a label column first and a row of names
 * on top (blank where the range has none), one column per series. A kind
 * outside Excel's label rule reads its columns as they are.
 */
export function insertChartLaidMatrix(
  rect: ChartRange,
  type: SpreadsheetChartType,
  value: (row: number, col: number) => string,
  display: (row: number, col: number) => string = value,
): { laid: InsertChartSelectionLayout; matrix: string[][] } {
  const laid = insertChartSelectionLayout(rect, type, value, display);
  const range = laid.range;
  if (!LABEL_RULE_TYPES.has(type)) {
    return {
      laid,
      matrix: Array.from({ length: range.r2 - range.r1 + 1 }, (_, row) =>
        Array.from({ length: range.c2 - range.c1 + 1 }, (__, col) => display(range.r1 + row, range.c1 + col)),
      ),
    };
  }
  const fields = {
    ...(laid.seriesInColumns ? {} : { seriesInColumns: false as const }),
    ...(laid.noCategories ? { noCategories: true as const } : {}),
    ...(laid.noSeriesNames ? { noSeriesNames: true as const } : {}),
  };
  const rows = buildChartSeries({ range, type, ...fields }, display, value);
  const width = nativeChartRangeLayout({ range, ...fields }).seriesLines.length + 1;
  return { laid, matrix: laid.noSeriesNames ? [Array.from({ length: width }, () => ""), ...rows] : rows };
}

/** The laid table of a selection's matrix (its cells as shown), per kind:
 * the ribbon asks each gallery kind of the same selection. */
const laidMatrixCache = new WeakMap<readonly (readonly string[])[], Map<SpreadsheetChartType, string[][]>>();

/** `insertChartLaidMatrix` over a selection's matrix (row 0, column 0 its
 * top-left cell), as the Insert ribbon's gates read it. */
export function insertChartGateMatrix(
  matrix: readonly (readonly string[])[],
  type: SpreadsheetChartType,
): string[][] {
  let byType = laidMatrixCache.get(matrix);
  if (!byType) {
    byType = new Map();
    laidMatrixCache.set(matrix, byType);
  }
  const cached = byType.get(type);
  if (cached) return cached;
  const width = Math.max(0, ...matrix.map((row) => row.length));
  const cell = (row: number, col: number) => matrix[row]?.[col] ?? "";
  const laid =
    matrix.length && width
      ? insertChartLaidMatrix({ r1: 0, c1: 0, r2: matrix.length - 1, c2: width - 1 }, type, cell).matrix
      : [];
  byType.set(type, laid);
  return laid;
}

/** What Insert Chart reads to plan a chart over a selection. */
export interface PlannedSheetChartInput {
  readonly rect: ChartRange;
  readonly type: SpreadsheetChartType;
  /** The shared catalog subtype, written only when the leaf is not the
   * kind's own default (see spreadsheetChartTypeForLeaf). */
  readonly chartSubtype?: string;
  /** The combo leaf that is expressed as a secondary-axis series instead. */
  readonly secondaryAxisLastSeries?: boolean;
  /** Custom Combination: each series' chart type and the secondary ones. */
  readonly combination?: { types: readonly ComboRenderAs[]; secondary: readonly number[] };
  readonly id: string;
  /** The placeholder title ("Chart Title") in the page's language. */
  readonly placeholderTitle: string;
  readonly rowCount: number;
  readonly columnCount: number;
  cellDisplay(row: number, col: number): string;
  cellValue(row: number, col: number): string;
  /** The grid's column widths and row heights, as the grid resolves them. */
  readonly colWidth: (col: number) => number;
  readonly rowHeight: (row: number) => number;
}

/**
 * The chart Insert Chart makes of the selection -- its record, born with its
 * model over the cells it binds -- without inserting it. The one
 * construction behind every insert and behind the Recommended Charts
 * preview, so what that dialog shows is exactly what its OK inserts. Null
 * when the selection cannot be plotted as `type`.
 */
/** The families whose category axis may be a date axis. */
const DATE_AXIS_FAMILIES: ReadonlySet<string> = new Set(["column", "bar", "line", "area", "stock", "combo"]);

/** The families whose legend lists the points of one series. */
const LEGEND_LISTS_POINTS = new Set(["pie", "hierarchy", "funnel", "map"]);

export function plannedSheetChart(input: PlannedSheetChartInput): { chart: SpreadsheetChart; rows: string[][] } | null {
  const { rect, type, chartSubtype, secondaryAxisLastSeries, combination } = input;
  if (!canAuthorNativeXlsxChart(type)) return null;
  // The table the chart plots, a label column and a row of names first
  // (blank where the range has none), as the type gates read one.
  const { laid, matrix } = insertChartLaidMatrix(
    rect,
    type,
    (row, col) => input.cellValue(row, col),
    (row, col) => input.cellDisplay(row, col),
  );
  const { outer } = laid;
  const chartRange = laid.range;
  const labelFields = {
    ...(laid.seriesInColumns ? {} : { seriesInColumns: false as const }),
    ...(laid.noCategories ? { noCategories: true as const } : {}),
    ...(laid.noSeriesNames ? { noSeriesNames: true as const } : {}),
  };
  const layout = nativeChartRangeLayout({ range: chartRange, ...labelFields });
  if (!chartMatrixSupportsType(matrix, type, true, chartSubtype)) return null;
  // Excel puts exactly the last value series on the secondary axis for
  // "Clustered Column - Line on Secondary Axis".
  // An X Y chart in rows plots every row after the first (its X line).
  const xyRows = type === "scatter" && !laid.seriesInColumns;
  const seriesCount = LABEL_RULE_TYPES.has(type)
    ? layout.seriesLines.length
    : Math.max(1, xyRows ? chartRange.r2 - chartRange.r1 : chartRange.c2 - chartRange.c1);
  const lastValueSeries = Math.max(0, seriesCount - 1);
  // Excel inserts a chart with its automatic title on: a chart that plots
  // one series is titled with that series' name -- a pie plots only its
  // first, whatever else the selection holds -- and any other reads "Chart
  // Title" (the placeholder Docs and the deck insert too), as does every
  // chart of the newer kinds (waterfall, funnel, histogram, Pareto, box and
  // whisker, treemap, sunburst, map), one series or not (Excel for Mac,
  
  const seriesNames = (
    xyRows ? matrix.slice(1).map((row) => row[0] ?? "") : (matrix[0] ?? []).slice(1)
  ).map((name) => name.trim()).slice(0, seriesCount);
  const plottedNames = type === "pie" ? seriesNames.slice(0, 1) : seriesNames;
  const namedByHeader =
    matrix.length > 1 &&
    plottedNames.length > 0 &&
    plottedNames.every((name) => name !== "" && !Number.isFinite(Number(name)));
  const title =
    plottedNames.length === 1 && namedByHeader && !EXTENDED_XLSX_CHART_TYPES.has(type)
      ? plottedNames[0].slice(0, MAX_CHART_TITLE_LENGTH)
      : input.placeholderTitle;
  const leaf = spreadsheetChartLeafOf({ type, chartSubtype });
  const kindDefaults = leaf ? officeInsertedChartLegendAndLabels(leaf.family, leaf.subtype) : {};
  const chart: SpreadsheetChart = {
    id: input.id,
    type,
    title,
    range: { ...chartRange },
    ...labelFields,
    anchor: { row: rect.r1, col: rect.c1 },
    // Beside the data, level with its first row, not on top of it. The
    // grid draws a chart that has never been moved at one fixed spot near
    // the sheet's top-left, which is where a table usually is, so a new
    // chart covered the numbers it was made from. An explicit position is
    // what a drag writes too.
    x:
      DEFAULT_SPREADSHEET_HEADER_WIDTH_PX +
      offsetOfLine(rect.c2 + 1, input.columnCount, input.colWidth) +
      24,
    y:
      DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX +
      offsetOfLine(rect.r1, input.rowCount, input.rowHeight),
    // Excel's inserted chart: 5 x 3 inches (12.7 x 7.62 cm), which is what
    // gives its category axis room for every label (XLS-A107).
    width: 480,
    height: 288,
    // Excel's inserted chart puts its legend under the plot area, as the
    // deck's and Docs' inserts do; stated, so the file says so too. A radar,
    // a waterfall and a treemap put it at the top.
    legendPosition: kindDefaults.legend === "top" ? "top" : "bottom",
    // ...and a chart of ONE series has none: its title already names the
    // series (XLS-A107). A chart whose legend lists its points (a pie, a
    // doughnut, a treemap...) or its kinds of point (a waterfall) keeps it;
    // a funnel, a sunburst and the statistical charts have none at all.
    ...(kindDefaults.legend === "none" ||
    (kindDefaults.legend === undefined &&
      seriesNames.length === 1 &&
      !LEGEND_LISTS_POINTS.has(String(spreadsheetChartFamily(type))))
      ? { showLegend: false }
      : {}),
    // A waterfall and a funnel are inserted with their values labelled.
    ...(kindDefaults.dataLabels ? { showDataLabels: true } : {}),
    // Born with Office's Change Colors pick, Colorful Palette 1 (the
    // workbook theme's accents), as createChartBlockAttrsV2 gives a new
    // Docs or deck chart; the file states it as the chart's colour part.
    ...(chartTakesColorPalette(spreadsheetChartFamily(type)) ? { chartColorPalette: "colorful-1" as const } : {}),
    ...(chartSubtype ? { chartSubtype } : {}),
    ...(outer > 0 ? { outerCategoryColumns: outer } : {}),
    ...(secondaryAxisLastSeries
      ? { secondaryAxisSeries: [lastValueSeries] }
      : {}),
    ...(combination && combination.types.length === lastValueSeries + 1
      ? {
          comboSeriesTypes: [...combination.types],
          ...(combination.secondary.length
            ? { secondaryAxisSeries: [...combination.secondary] }
            : {}),
        }
      : {}),
  };
  // Born with its model: the shared factory over the cells it binds.
  const rows = buildChartSeries(
    chart,
    (row, col) => input.cellDisplay(row, col),
    (row, col) => input.cellValue(row, col),
  );
  // XLS-A107: the outer text columns are the axis's outer levels.
  const categoryLevels = chartCategoryLevelsForRows(
    rows,
    nativeChartOuterCategoryCells(chart, (row, col) => input.cellDisplay(row, col)),
  );
  const headerRow = nativeChartHeaderRow(chart);
  const born = spreadsheetChartBornWithModel(chart, {
    rows,
    ...(headerRow !== undefined ? { headerRow } : {}),
    ...(categoryLevels ? { categoryLevels } : {}),
  });
  // Excel 2016+ inserts a chart with Show #N/A as an empty cell on (XLS-A072)
  // -- a classic chart's setting: an Office 2016 chart's part has no such
  // option, and a model stating it was one its file could not hold (AI70H
  // H127 B12/B13: a histogram and a box plot kept on the page only) --
  // and states what a person's inserted chart looks like: a column or bar
  
  const model = born.chartV2 as ChartBlockAttrsV2 | undefined;
  // Excel's category axis is Automatic: over dates -- numbers the sheet
  // shows as dates -- it is a date axis, the dates at their places in time
  // (a stock chart's weekend is a gap), as Excel for Mac's inserted charts
  
  const datedCategories = (() => {
    if (!model || !DATE_AXIS_FAMILIES.has(model.chartType)) return false;
    const line = LABEL_RULE_TYPES.has(type) ? layout.categoryLine : chartRange.c1;
    if (line === null) return false;
    // The category cells, by columns or by rows.
    const at = (point: number): [number, number] => (layout.byColumns ? [point, line] : [line, point]);
    const [first, last] = LABEL_RULE_TYPES.has(type) ? layout.points : [chartRange.r1 + 1, chartRange.r2];
    let dates = 0;
    for (let point = first; point <= last; point += 1) {
      const [row, col] = at(point);
      const raw = input.cellValue(row, col).trim();
      const shown = input.cellDisplay(row, col).trim();
      if (raw === "" && shown === "") continue;
      const serial = Number(raw);
      if (!Number.isFinite(serial) || Number.isFinite(Number(shown.replace(/[,\s]/g, ""))) || !officeChartDateAxisTimes([shown])) {
        return false;
      }
      dates += 1;
    }
    return dates > 0;
  })();
  return {
    chart: model
      ? {
          ...born,
          chartV2: withOfficeInsertedChartFormat({
            ...model,
            format: {
              ...(model.format ?? {}),
              ...(OFFICE_CHART_EX_FAMILIES.has(model.chartType) ? {} : { displayNaAsBlank: true }),
              ...(datedCategories
                ? {
                    axisOptions: {
                      ...(model.format?.axisOptions ?? {}),
                      category: { ...(model.format?.axisOptions?.category ?? {}), axisType: "date" as const },
                    },
                  }
                : {}),
            },
          }),
        }
      : born,
    rows,
  };
}

// ---------------------------------------------------------------------------
// Chart Design and Format buttons. `parameter` is what the ribbon control
// hands its writer (a type id, "true"/"false", a JSON pick, ...).

/** Whether the chart plots its range's rows as series (Switch Row/Column's
 * current side): a native chart's `seriesInColumns`, an imported chart's
 * series layout over its range. */
export function sheetChartPlotsByRows(target: Pick<SheetChartTarget, "kind" | "chart">): boolean {
  const chart = target.chart as SheetChartFields;
  if (target.kind === "native") return chart.seriesInColumns === false;
  const source = drawingChartSourceOf(
    normalizeChartSeriesRefs(chart.seriesRefs),
    normalizeSpreadsheetChartType(chart.chartType),
  );
  return source?.seriesInColumns === false;
}

/** An imported chart over another layout of its data: the model re-read
 * over the new series references, which the save writes into the chart's
 * own series (each keeping its formatting). */
function drawingChartWithSeriesRefs(
  target: SheetChartTarget,
  data: SheetChartData,
  refs: SpreadsheetChartSeriesRef[],
  options: { switched?: boolean; seriesOrigins?: readonly (number | null)[] } = {},
): SheetChartWrite | null {
  const model = sheetChartModelOf(target, data);
  const reading = sheetChartReadingOf(target, data, { seriesRefs: refs, ranges: drawingChartRangesOf(refs) });
  if (!model || !reading || sheetChartDataLiteral(model.binding)) return null;
  // Legend Entries removed or moved one by one: the format follows its
  // series to their new places (Excel's formatting stays with the series),
  // and what belonged to a removed series goes with it (XLS-A072).
  const followed = options.seriesOrigins
    ? chartFormatFollowingDataEdit(model.attrs.format, options.seriesOrigins, null)
    : model.attrs.format;
  const base: ChartBlockAttrsV2 = followed === model.attrs.format
    ? model.attrs
    : (() => {
        const moved = { ...model.attrs, format: followed } as ChartBlockAttrsV2;
        if (!followed || !Object.keys(followed).length) delete (moved as { format?: unknown }).format;
        return moved;
      })();
  const next = sheetChartRetargeted(base, reading, options);
  if (!next) return null;
  const binding = { ...(model.binding ?? {}), seriesRefs: refs } as Json;
  const rect = drawingChartSourceOf(refs, chartTypeOf(target))?.range;
  if (rect) binding.range = rect;
  else delete binding.range;
  return { attrs: next, binding: binding as unknown as SpreadsheetChartBinding };
}

/**
 * A chart born here as one that states its series one by one: what Select
 * Data's Legend Entries make of it once the edited series are no longer one
 * range (a series removed from the middle, moved, or read from elsewhere --
 * XLS-A072). The record is the chart drawing object the same chart is after
 * an XLSX round trip: its series references, the one chart model re-read
 * over them with the format following its series, and the frame the host
 * gives it (`anchor`). Null when the chart is not one the series writer
 * lays, or the model does not take the new data.
 */
export function nativeChartAsSeriesBoundObject(
  target: SheetChartTarget,
  data: SheetChartData,
  refs: SpreadsheetChartSeriesRef[],
  seriesOrigins: readonly (number | null)[],
  anchor: unknown,
): Record<string, unknown> | null {
  if (target.kind !== "native") return null;
  const chart = fieldsOf(target);
  const type = chartTypeOf(target);
  if (chart.pivotSource || !drawingChartTakesDataRange(type) || refs.length === 0) return null;
  const model = sheetChartModelOf(target, data);
  if (!model || sheetChartDataLiteral(model.binding)) return null;
  const record = target.chart as unknown as Record<string, unknown>;
  const shell: Record<string, unknown> = {
    id: record.id,
    type: "chart",
    chartType: type,
    title: typeof record.title === "string" ? record.title : "",
    ...(typeof record.chartSubtype === "string" ? { chartSubtype: record.chartSubtype } : {}),
    ...(typeof record.name === "string" ? { name: record.name } : {}),
    ...(record.hidden === true ? { hidden: true } : {}),
    ...(typeof record.groupId === "string" ? { groupId: record.groupId } : {}),
    seriesRefs: refs,
    ranges: drawingChartRangesOf(refs),
    chartVisibility: { state: "known", plotVisOnly: true },
    anchor,
  };
  const reading = sheetChartReadingOf({ ...target, kind: "drawing", chart: shell } as unknown as SheetChartTarget, data, {});
  if (!reading) return null;
  const followed = chartFormatFollowingDataEdit(model.attrs.format, seriesOrigins, null);
  const base = { ...model.attrs, format: followed } as ChartBlockAttrsV2;
  if (!followed || !Object.keys(followed).length) delete (base as { format?: unknown }).format;
  const next = sheetChartRetargeted(base, reading);
  if (!next) return null;
  const binding = { ...(model.binding ?? {}), seriesRefs: refs } as Json;
  delete binding.plotBy;
  const rect = drawingChartSourceOf(refs, type)?.range;
  if (rect) binding.range = rect;
  else delete binding.range;
  return withSpreadsheetChartV2(
    shell as unknown as SpreadsheetChartLike,
    next,
    binding as unknown as SpreadsheetChartBinding,
  ) as unknown as Record<string, unknown>;
}

/** Design > Switch Row/Column (a range-bound chart born here). */
export const chartSwitchRowColumnButton: SheetChartButton = (target, parameter, data) => {
  const chart = fieldsOf(target);
  if (target.kind === "drawing") {
    // Excel lays the chart's one data range the other way; a chart whose
    // series are not one range ("too complex") has no other way to lay.
    const type = chartTypeOf(target);
    const source = drawingChartSourceOf(normalizeChartSeriesRefs(chart.seriesRefs), type);
    if (
      !source ||
      chart.pivotSource ||
      spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike) ||
      parameter !== String(!source.seriesInColumns)
    ) {
      return null;
    }
    const refs = drawingChartSeriesRefsForRange(source.prefix, source.range, !source.seriesInColumns, type);
    return refs ? drawingChartWithSeriesRefs(target, data, refs, { switched: true }) : null;
  }
  if (target.kind !== "native" || chart.pivotSource || parameter !== String(chart.seriesInColumns === false)) return null;
  // Laid the other way, the line that named the series labels the categories.
  const switchedGate = nativeChartGateMatrix(data.matrixForRange(chart.range as ChartRange, chart.sourceSheetId), {
    seriesInColumns: parameter === "true",
    ...(chart.noSeriesNames ? { noCategories: true as const } : {}),
    ...(chart.noCategories ? { noSeriesNames: true as const } : {}),
  });
  if (!chartMatrixSupportsType(switchedGate.matrix, chart.type as SpreadsheetChartType, switchedGate.seriesInColumns)) {
    return null;
  }
  // The binding plots the other way; the data is re-read for it and
  // Chart Filters' hidden series become hidden categories and back. The
  // line that named the series labels the categories and back: a range
  // without a category column (North and South side by side) switched has
  // no row of names, its categories North and South (Excel's Switch).
  const byRows = chart.seriesInColumns !== false;
  const swapped = {
    noCategories: chart.noSeriesNames === true ? (true as const) : undefined,
    noSeriesNames: chart.noCategories === true ? (true as const) : undefined,
  };
  const model = sheetChartModelOf(target, data);
  const reading = sheetChartReadingOf(target, data, { seriesInColumns: byRows ? false : undefined, ...swapped });
  if (!model || !reading) return null;
  const binding: SpreadsheetChartBinding = { ...model.binding } as SpreadsheetChartBinding;
  if (byRows) (binding as Json).plotBy = "rows";
  else delete (binding as Json).plotBy;
  const next = sheetChartRetargeted(model.attrs, reading, { switched: true });
  return next ? { attrs: next, binding, ...sheetChartLabelFieldsWrite(swapped) } : null;
};

/** A write's record fields for a range-bound chart's label lines: the ones
 * it states set, the others removed. */
function sheetChartLabelFieldsWrite(fields: { noCategories?: true; noSeriesNames?: true }): Pick<SheetChartWrite, "record" | "settled"> {
  const record: Record<string, unknown> = {};
  const settled: SheetChartSettledField[] = [];
  for (const key of ["noCategories", "noSeriesNames"] as const) {
    if (fields[key]) record[key] = true;
    else settled.push(key);
  }
  return { ...(Object.keys(record).length ? { record } : {}), settled };
}

/**
 * The label lines Select Data's Chart data range has, laid by columns or by
 * rows, as Excel reads them (`excelChartRangeLabelLines`): read on the
 * chart's own sheet from its cells, on another from what they show.
 */
export function selectDataRangeLabelFields(
  data: Pick<SheetChartData, "cellValue" | "cellDisplay" | "matrixForRange">,
  range: ChartRange,
  sourceSheetId: string | undefined,
  byRows: boolean,
): { noCategories?: true; noSeriesNames?: true } {
  if (sourceSheetId) {
    const matrix = data.matrixForRange(range, sourceSheetId);
    const shown = (row: number, col: number) => matrix[row - range.r1]?.[col - range.c1] ?? "";
    return chartRangeLabelFields(excelChartRangeLabelLines(range, shown), !byRows);
  }
  return chartRangeLabelFields(
    excelChartRangeLabelLines(range, (row, col) => data.cellValue(row, col), (row, col) => data.cellDisplay(row, col)),
    !byRows,
  );
}

/** Design > Quick Layout. */
export const chartQuickLayoutButton: SheetChartButton = (target, parameter, data) => {
  const layoutId = parameter as ChartQuickLayoutId;
  if (!CHART_QUICK_LAYOUT_IDS.includes(layoutId)) return null;
  return sheetChartModelEdited(target, data, (attrs) =>
    officeChartWithQuickLayout(attrs, layoutId, data.words),
  );
};

/** Design > Chart Styles: a `c:style` pick is the model's `stylePreset`. */
export const chartStyleButton: SheetChartButton = (target, parameter, data) => {
  const styleId = chartStyleIdFromGallery(parameter);
  const preset = styleId === null ? undefined : CHART_STYLE_PRESET_FOR_C_STYLE[String(styleId)];
  // An imported chart's old style is replaced, as Excel replaces it; the
  // save writes the pick through the chart's model (or declines that pick
  // alone when the part cannot take it).
  if (!preset) return null;
  return sheetChartModelEdited(target, data, (attrs) =>
    attrs.stylePreset === preset ? null : { ...attrs, stylePreset: preset },
  );
};

/** Design > Change Colors. */
export const chartColorsButton: SheetChartButton = (target, parameter, data) => {
  const paletteId = chartColorPaletteIdFromGallery(parameter);
  // An imported paint (a gradient, a style the flat fields had no words
  // for) is repainted from the palette, as Excel's Change Colors does.
  if (paletteId === null) return null;
  // Change Colors repaints every series' fill from the palette -- a
  // gradient fill included -- and keeps its outline and effect.
  return sheetChartModelEdited(target, data, (attrs) => sheetChartWithChangedColors(attrs, paletteId));
};

/** Design > Change Chart Type. */
export const chartTypeButton: SheetChartButton = (target, parameter, data) => {
  const type = parameter as SpreadsheetChartType;
  if (
    !canAuthorNativeXlsxChart(type) ||
    (target.kind === "native" &&
      fieldsOf(target).pivotSource &&
      !isSpreadsheetPivotChartType(type)) ||
    !chartDesignTypeSupportsSource(target.kind, target.chart, type, data)
  ) {
    return null;
  }
  const currentType = chartTypeOf(target);
  if (type === currentType) return null;
  // The kind's own leaf, read over the chart's source: the shared type
  // command keeps the model's titles and format where the data takes
  // them.
  const reading = sheetChartReadingOf(target, data, {
    [target.kind === "native" ? "type" : "chartType"]: type,
    chartSubtype: undefined,
    secondaryAxisSeries: undefined,
    comboSeriesTypes: undefined,
  });
  if (!reading) return null;
  return sheetChartModelEdited(
    target,
    data,
    (attrs, binding) =>
      officeChartWithType(attrs, reading.chartType, reading.chartSubtype) ??
      (sheetChartDataLiteral(binding) ? null : sheetChartRetargeted(attrs, reading)),
  );
};

/**
 * Design > Change Chart Type with the gallery's tile picked (`family.subtype`,
 * e.g. "column.clustered3d", "pie.pie3d"): the dialog's own route
 * (SpreadsheetWindow chooseDesktopChartType) -- the shared type command on
 * the chart's model, else the leaf's kind re-read from the cells. A 3-D tile
 * of the chart's own kind is a subtype change, not a kind change. Combo
 * tiles go through Custom Combination, not here.
 */
export const chartTypeLeafButton: SheetChartButton = (target, parameter, data) => {
  const [family, subtype] = parameter.split(".") as [ChartFamily, ChartSubtype];
  if (!family || !subtype || family === "combo") return null;
  const leaf = spreadsheetChartTypeForLeaf(family, subtype);
  if (!leaf) return null;
  const type = leaf.type;
  if (
    !canAuthorNativeXlsxChart(type) ||
    (target.kind === "native" && fieldsOf(target).pivotSource && !isSpreadsheetPivotChartType(type)) ||
    !chartDesignTypeSupportsSource(target.kind, target.chart, type, data)
  ) {
    return null;
  }
  const reading = sheetChartReadingOf(target, data, {
    [target.kind === "native" ? "type" : "chartType"]: type,
    chartSubtype: leaf.chartSubtype,
    secondaryAxisSeries: undefined,
    comboSeriesTypes: undefined,
  });
  return sheetChartModelEdited(target, data, (attrs, binding) => {
    if (attrs.chartType === family && attrs.chartSubtype === subtype) return null;
    const typed = officeChartWithType(attrs, family, subtype);
    if (typed) return typed;
    return reading && reading.chartType === family && !sheetChartDataLiteral(binding) ? sheetChartRetargeted(attrs, reading) : null;
  });
};

/** Kinds whose series Plot Series On moves one by one: a 2-D line, column
 * or stacked column chart (a custom combination of its one kind once a
 * series is on the secondary axis). */
const SECONDARY_AXIS_KINDS: ReadonlySet<string> = new Set(["line", "column", "stacked-column"]);

/**
 * Plot Series On for an imported chart: the series its part states one by
 * one (`seriesRefs`), each keeping its references. A chart of one kind is
 * read with the listed series on the secondary axis (`secondaryAxisSeries`,
 * as the import states such a chart and a chart born here keeps it); a
 * combination chart becomes Combo > Custom Combination, each series' type
 * kept and its Secondary Axis box the pick -- Excel's own two routes
 * (Microsoft Support, "Add or remove a secondary axis in a chart in Excel").
 * The series count is the references', the model must take the pick (the
 * reading plots exactly those series on the secondary axis), and the save
 * writes the second value axis into the chart's own part.
 */
function drawingChartWithSecondaryAxis(
  target: SheetChartTarget,
  data: SheetChartData,
  secondary: number[] | undefined,
): SheetChartWrite | null {
  const chart = fieldsOf(target);
  if (chart.pivotSource || spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike)) return null;
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  const model = sheetChartModelOf(target, data);
  if (!refs || !model || sheetChartDataLiteral(model.binding)) return null;
  const { attrs } = model;
  if (attrs.data.kind !== "category" || attrs.data.series.length !== refs.length || /3d/i.test(attrs.chartSubtype)) {
    return null;
  }
  const format = attrs.format ?? {};
  const wanted = new Set(secondary ?? []);
  if (
    [...wanted].some((index) => index >= refs.length) ||
    (wanted.size > 0 && (format.upDownBars === true || format.dropLines === true || format.highLowLines === true || format.seriesLines === true))
  ) {
    return null;
  }
  // The record's own kind decides the route: a chart of one kind with a
  // series already on the secondary axis is that kind + its list (its model
  // reads as a custom combination of that kind), a combination chart is
  // `combo` with each series' type in its references.
  const type = chartTypeOf(target);
  const combination = type === "combo";
  if (!combination && (!SECONDARY_AXIS_KINDS.has(type) || !["line", "column", "combo"].includes(attrs.chartType))) return null;
  const rows = officeChartCombinationRows(attrs);
  if (!rows || rows.length !== refs.length) return null;
  const nextRows = rows.map((row, index) => ({ ...row, secondary: wanted.has(index) }));
  if (nextRows.every((row, index) => row.secondary === rows[index].secondary)) return null;
  if (wanted.size > 0 && officeChartCombinationRefusal(nextRows) !== null) return null;
  const nextRefs: SpreadsheetChartSeriesRef[] = refs.map((ref, index) => {
    const rest: SpreadsheetChartSeriesRef = { ...ref };
    delete rest.axis;
    if (combination) return { ...rest, renderType: nextRows[index].renderAs, axis: nextRows[index].secondary ? "secondary" : "primary" };
    return wanted.has(index) ? { ...rest, axis: "secondary" } : rest;
  });
  const reading = sheetChartReadingOf(
    target,
    data,
    combination
      ? { chartSubtype: "custom", seriesRefs: nextRefs, secondaryAxisSeries: undefined, comboSeriesTypes: undefined }
      : { seriesRefs: nextRefs, secondaryAxisSeries: secondary },
  );
  if (
    !reading ||
    reading.data.kind !== "category" ||
    reading.data.series.length !== refs.length ||
    reading.data.series.some((series, index) => (series.axis === "secondary") !== wanted.has(index))
  ) {
    return null;
  }
  const changed = sheetChartRetargeted(attrs, reading);
  if (!changed) return null;
  const binding = { ...(model.binding ?? {}), seriesRefs: nextRefs } as Json;
  return { attrs: changed, binding: binding as unknown as SpreadsheetChartBinding };
}

/**
 * Format Data Series > Plot Series On > Secondary Axis: `parameter` is the
 * JSON list of the series (0-based) drawn on the secondary axis ("[]" puts
 * them all back on the primary one). The chart is read with those series
 * on the secondary axis -- a custom combination of its one kind -- its
 * titles and format kept. A line, column or stacked column chart (or, for
 * an imported chart, a combination chart), at least one series left on the
 * primary axis, and no up/down bars or drop, high-low or series lines.
 */
export const chartSecondaryAxisButton: SheetChartButton = (target, parameter, data) => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  if (!Array.isArray(parsed)) return null;
  if (target.kind === "drawing") return drawingChartWithSecondaryAxis(target, data, normalizeSecondaryAxisSeries(parsed));
  if (target.kind !== "native") return null;
  const chart = fieldsOf(target) as SheetChartFields & {
    secondaryAxisSeries?: number[];
    upDownBars?: boolean;
    dropLines?: boolean;
    highLowLines?: boolean;
    seriesLines?: boolean;
  };
  const next = normalizeSecondaryAxisSeries(parsed);
  const rows = buildChartSeries(target.chart as never, (row, col) => data.cellDisplay(row, col));
  const seriesCount = Math.max(0, (rows[0]?.length ?? 1) - 1);
  if (
    (next?.some((index) => index >= seriesCount) ?? false) ||
    (next !== undefined && next.length >= seriesCount) ||
    (next !== undefined && chart.upDownBars === true) ||
    (next !== undefined && (chart.dropLines === true || chart.highLowLines === true || chart.seriesLines === true)) ||
    (next && !["line", "column", "stacked-column"].includes(String(chart.type)))
  ) {
    return null;
  }
  if (JSON.stringify(next) === JSON.stringify(chart.secondaryAxisSeries)) return null;
  const model = sheetChartModelOf(target, data);
  const reading = sheetChartReadingOf(target, data, { secondaryAxisSeries: next });
  if (!model || !reading) return null;
  const changed = sheetChartRetargeted(model.attrs, reading);
  return changed ? { attrs: changed, binding: model.binding } : null;
};

/** Add Chart Element > Chart Title; "" hides it. */
export const chartTitleButton: SheetChartButton = (target, parameter, data) => {
  const nextTitle = parameter
    .trim()
    .slice(0, MAX_CHART_TITLE_LENGTH);
  const currentTitle = spreadsheetChartTitle(target.chart as Parameters<typeof spreadsheetChartTitle>[0]);
  if (nextTitle === currentTitle) return null;
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithElementIntent(attrs, {
      element: "chartTitle",
      on: Boolean(nextTitle),
      text: nextTitle,
    }),
  );
};

/** The Data Labels switch ("true" / "false", the state it turns to). */
export const chartDataLabelsButton: SheetChartButton = (target, parameter, data) => {
  if (!spreadsheetChartTypeSupportsDataLabels(chartTypeOf(target))) return null;
  const currentlyShown = fieldsOf(target).showDataLabels === true;
  if (
    parameter !==
    String(!currentlyShown)
  ) {
    return null;
  }
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithDataLabels(attrs, parameter === "true"),
  );
};

/** Add Chart Element > Axis Titles, and the pane's axis title fields. */
export function chartAxisTitleButton(axis: "x" | "y"): SheetChartButton {
  return (target, parameter, data) => {
    if (!spreadsheetChartTypeHasAxes(chartTypeOf(target))) return null;
    const key = axis === "x" ? "axisTitleX" : "axisTitleY";
    const nextTitle = parameter.trim();
    if (nextTitle === (fieldsOf(target)[key] ?? "")) return null;
    return sheetChartModelEdited(target, data, (attrs) => {
      if (!chartSupportsAxisTitle(attrs.chartType, axis)) return null;
      const next: ChartBlockAttrsV2 = { ...attrs };
      if (nextTitle) next[key] = clampChartText(nextTitle);
      else delete next[key];
      return next;
    });
  };
}

/** Select Data's OK: the chart's data range ("A1:C9", canonical). */
export const chartSelectDataButton: SheetChartButton = (target, parameter, data) => {
  const chart = fieldsOf(target);
  if (chart.pivotSource) {
    return null;
  }
  const normalized = normalizeChartRange(parameter);
  if (!normalized) return null;
  const canonical = chartRangeToA1(normalized);
  if (canonical !== parameter) return null;
  if (target.kind === "drawing") {
    // The range is on the sheet the chart's series read, laid the way they
    // are laid now (Excel's Select Data keeps the chart's orientation).
    // A chart whose references name no one sheet takes the range on its own.
    const type = chartTypeOf(target);
    const current = normalizeChartSeriesRefs(chart.seriesRefs);
    const sheet = (current ? drawingChartSheetOf(current) : null) ?? {
      prefix: drawingChartSheetPrefix(data.sheetName),
      sheetName: data.sheetName,
    };
    if (spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike)) return null;
    const source = drawingChartSourceOf(current, type);
    const seriesInColumns = source ? source.seriesInColumns : drawingChartDefaultSeriesInColumns(current, normalized);
    const refs = drawingChartSeriesRefsForRange(sheet.prefix, normalized, seriesInColumns, type);
    if (!refs || JSON.stringify(refs) === JSON.stringify(current)) return null;
    return drawingChartWithSeriesRefs(target, data, refs);
  }
  // The new range's label lines, as Excel reads a Chart data range; the
  // chart keeps its orientation.
  const byRows = chart.seriesInColumns === false;
  const labels = LABEL_RULE_TYPES.has(chart.type as SpreadsheetChartType)
    ? selectDataRangeLabelFields(data, normalized, chart.sourceSheetId, byRows)
    : {};
  const gate = nativeChartGateMatrix(data.matrixForRange(normalized, chart.sourceSheetId), {
    seriesInColumns: !byRows,
    ...labels,
  });
  if (!chartMatrixSupportsType(gate.matrix, chart.type as SpreadsheetChartType, gate.seriesInColumns)) {
    return null;
  }
  const range = chart.range as ChartRange;
  if (
    range.r1 === normalized.r1 &&
    range.r2 === normalized.r2 &&
    range.c1 === normalized.c1 &&
    range.c2 === normalized.c2 &&
    (labels.noCategories === true) === (chart.noCategories === true) &&
    (labels.noSeriesNames === true) === (chart.noSeriesNames === true) &&
    !chart.seriesRefs &&
    !chart.sourceRefError
  ) {
    return null;
  }
  // The binding takes the new rectangle (per-series references give
  // way to it) and the data is re-read from it. Picking data repairs
  // a #REF! source left by deleted lines.
  const model = sheetChartModelOf(target, data);
  const reading = sheetChartReadingOf(target, data, {
    range: normalized,
    seriesRefs: undefined,
    sourceRefError: undefined,
    noCategories: labels.noCategories,
    noSeriesNames: labels.noSeriesNames,
  });
  if (!model || !reading) return null;
  const next = sheetChartRetargeted(model.attrs, reading);
  const binding = { ...(model.binding ?? {}) } as Json;
  delete binding.seriesRefs;
  const labelWrite = sheetChartLabelFieldsWrite(labels);
  return next
    ? {
        attrs: next,
        binding: { ...binding, range: normalized } as unknown as SpreadsheetChartBinding,
        ...(labelWrite.record ? { record: labelWrite.record } : {}),
        settled: ["sourceRefError", ...(labelWrite.settled ?? [])],
      }
    : null;
};

/**
 * Select Data Source's OK (XLS-A072): everything the dialog changed, in one
 * write -- a JSON `{ range?, byRows?, seriesRefs?, displayBlanksAs? }`:
 * - `seriesRefs`: an imported chart's Legend Entries as edited (each series'
 *   name and values, the shared axis labels), which the save writes into the
 *   chart's own series in place;
 * - else `range` ("A1:C9", canonical) laid by rows or by columns (`byRows`,
 *   the dialog's Switch Row/Column), every series replaced by what it lays;
 *   on another worksheet when the range named one: `sheetName` (an imported
 *   chart's series references) or `sourceSheetId` (a chart born here; null
 *   for its own sheet);
 * - `displayBlanksAs`: Hidden and Empty Cells > Show empty cells as.
 * Null when the dialog changed nothing or the chart does not take it.
 */
export const chartSelectDataSourceButton: SheetChartButton = (target, parameter, data) => {
  let intent: Json;
  try {
    intent = JSON.parse(parameter) as Json;
  } catch {
    return null;
  }
  if (!intent || typeof intent !== "object" || Array.isArray(intent)) return null;
  const chart = fieldsOf(target);
  if (chart.pivotSource || spreadsheetChartEditRefusal(target.chart as unknown as SpreadsheetChartLike)) return null;
  const type = chartTypeOf(target);
  const byRows = intent.byRows === true;
  let write: SheetChartWrite | null = null;
  if (intent.seriesRefs !== undefined) {
    // Legend Entries edited one by one: only an imported chart states its
    // series as references.
    const refs = normalizeChartSeriesRefs(intent.seriesRefs);
    if (target.kind !== "drawing" || !refs || !drawingChartTakesDataRange(type)) return null;
    if (JSON.stringify(refs) !== JSON.stringify(normalizeChartSeriesRefs(chart.seriesRefs))) {
      // Where each series was before the edit (null: one the dialog added).
      const origins = Array.isArray(intent.seriesOrigins) && intent.seriesOrigins.length === refs.length &&
        intent.seriesOrigins.every((origin) => origin === null || (Number.isInteger(origin) && (origin as number) >= 0))
        ? (intent.seriesOrigins as (number | null)[])
        : undefined;
      const moved = origins?.some((origin, index) => origin !== index) ?? false;
      write = drawingChartWithSeriesRefs(target, data, refs, moved ? { seriesOrigins: origins } : {});
      if (!write) return null;
    }
  } else if (typeof intent.range === "string") {
    const normalized = normalizeChartRange(intent.range);
    if (!normalized || chartRangeToA1(normalized) !== intent.range) return null;
    if (target.kind === "drawing") {
      const current = normalizeChartSeriesRefs(chart.seriesRefs);
      if (intent.sheetName !== undefined && (typeof intent.sheetName !== "string" || !intent.sheetName)) return null;
      const sheet = typeof intent.sheetName === "string"
        ? { prefix: drawingChartSheetPrefix(intent.sheetName), sheetName: intent.sheetName }
        : (current ? drawingChartSheetOf(current) : null) ?? {
            prefix: drawingChartSheetPrefix(data.sheetName),
            sheetName: data.sheetName,
          };
      const source = drawingChartSourceOf(current, type);
      const refs = drawingChartSeriesRefsForRange(sheet.prefix, normalized, !byRows, type);
      if (!refs) return null;
      if (JSON.stringify(refs) !== JSON.stringify(current)) {
        // The same range laid the other way is Switch Row/Column: the
        // format's per-series and per-point styles change places with it.
        const switched = Boolean(
          source &&
            source.seriesInColumns === byRows &&
            JSON.stringify(source.range) === JSON.stringify(normalized),
        );
        write = drawingChartWithSeriesRefs(target, data, refs, { switched });
        if (!write) return null;
      }
    } else {
      if (intent.sourceSheetId !== undefined && intent.sourceSheetId !== null &&
        (typeof intent.sourceSheetId !== "string" || !intent.sourceSheetId)) return null;
      const sourceSheetId = intent.sourceSheetId === undefined
        ? chart.sourceSheetId
        : (intent.sourceSheetId as string | null) ?? undefined;
      const sourceMatrix = data.matrixForRange(normalized, sourceSheetId);
      // A sheet that is not there reads nothing: refused, not drawn from here.
      if (sourceSheetId && sourceMatrix.length === 0) return null;
      // The range's label lines, as Excel reads a Chart data range.
      const labels = LABEL_RULE_TYPES.has(type) ? selectDataRangeLabelFields(data, normalized, sourceSheetId, byRows) : {};
      const gate = nativeChartGateMatrix(sourceMatrix, { seriesInColumns: !byRows, ...labels });
      if (!chartMatrixSupportsType(gate.matrix, type, gate.seriesInColumns)) return null;
      const range = chart.range as ChartRange;
      const wasByRows = chart.seriesInColumns === false;
      const sameRange =
        range.r1 === normalized.r1 && range.r2 === normalized.r2 && range.c1 === normalized.c1 && range.c2 === normalized.c2;
      const sameSource = (sourceSheetId ?? null) === (chart.sourceSheetId ?? null);
      const sameLabels =
        (labels.noCategories === true) === (chart.noCategories === true) &&
        (labels.noSeriesNames === true) === (chart.noSeriesNames === true);
      if (!sameRange || !sameSource || wasByRows !== byRows || !sameLabels || chart.seriesRefs || chart.sourceRefError) {
        const model = sheetChartModelOf(target, data);
        const reading = sheetChartReadingOf(target, data, {
          range: normalized,
          sourceSheetId,
          seriesInColumns: byRows ? false : undefined,
          seriesRefs: undefined,
          sourceRefError: undefined,
          noCategories: labels.noCategories,
          noSeriesNames: labels.noSeriesNames,
        });
        if (!model || !reading) return null;
        const next = sheetChartRetargeted(model.attrs, reading, { switched: wasByRows !== byRows });
        if (!next) return null;
        const binding = { ...(model.binding ?? {}), range: normalized } as Json;
        delete binding.seriesRefs;
        if (byRows) binding.plotBy = "rows";
        else delete binding.plotBy;
        // The sheet the range is read on (the chart's sourceSheetId derives
        // from it); none is the chart's own sheet.
        if (sourceSheetId) binding.sheetId = sourceSheetId;
        else delete binding.sheetId;
        const labelWrite = sheetChartLabelFieldsWrite(labels);
        write = {
          attrs: next,
          binding: binding as unknown as SpreadsheetChartBinding,
          ...(labelWrite.record ? { record: labelWrite.record } : {}),
          settled: ["sourceRefError", ...(labelWrite.settled ?? [])],
        };
      }
    }
  }
  if (intent.showHidden !== undefined) {
    // Hidden and Empty Cells > Show data in hidden rows and columns: the
    // chart's `c:plotVisOnly`, stated (Excel writes it on every chart).
    if (typeof intent.showHidden !== "boolean") return null;
    const base = write ?? sheetChartModelOf(target, data);
    if (!base) return null;
    write = {
      ...(write ?? {}),
      attrs: base.attrs,
      binding: base.binding,
      record: { ...(write?.record ?? {}), chartVisibility: { state: "known", plotVisOnly: !intent.showHidden } },
    };
  }
  if (intent.displayNaAsBlank !== undefined) {
    // Hidden and Empty Cells > Show #N/A as an empty cell, on the model.
    if (typeof intent.displayNaAsBlank !== "boolean") return null;
    const base = write ?? sheetChartModelOf(target, data);
    if (!base) return null;
    const format = { ...(base.attrs.format ?? {}) } as Json;
    if (intent.displayNaAsBlank) format.displayNaAsBlank = true;
    else delete format.displayNaAsBlank;
    write = {
      ...(write ?? {}),
      attrs: { ...base.attrs, format: format as ChartBlockAttrsV2["format"] },
      binding: write?.binding ?? base.binding,
    };
  }
  if (intent.displayBlanksAs !== undefined) {
    // Hidden and Empty Cells > Show empty cells as, on the same model.
    const value = intent.displayBlanksAs;
    if (!isChartDisplayBlanksAs(value)) return null;
    const base = write ?? sheetChartModelOf(target, data);
    if (!base) return null;
    const next = chartWithElementIntent(base.attrs, { element: "displayBlanksAs", value });
    if (next && JSON.stringify(next) !== JSON.stringify(base.attrs)) {
      write = { ...(write ?? {}), attrs: next, binding: write?.binding ?? base.binding };
    }
  }
  return write;
};

/** The Legend switch ("true" / "false"): on shows it at the right. */
export const chartLegendButton: SheetChartButton = (target, parameter, data) => {
  const chart = fieldsOf(target);
  const next = parameter === "true";
  const current = chart.showLegend !== false;
  if (next === current) return null;
  // Legend on: Office's Chart Elements check box shows it at the right.
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithElementIntent(attrs, {
      element: "legend",
      position: next ? "right" : "none",
    }),
  );
};

/** Add Chart Element > Legend > Right / Top / Left / Bottom. */
export const chartLegendPositionButton: SheetChartButton = (target, parameter, data) => {
  const position = parameter;
  // Format Legend ▸ Legend Position also offers Top Right.
  if (!isSpreadsheetChartLegendMenuPosition(position) && position !== "topRight") {
    return null;
  }
  // Office's Add Chart Element > Legend > Right / Top / Left / Bottom
  // shows the legend at that edge: a hidden legend comes back there
  // in the same one Undo step.
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithElementIntent(attrs, { element: "legend", position }),
  );
};

/**
 * Format Axis > Bins of a histogram or pareto: `{binWidth | binCount,
 * overflow, underflow}` as JSON, or `null` for Automatic bins with no
 * overflow / underflow bin (`format.histogramBins`).
 */
export const chartHistogramBinsButton: SheetChartButton = (target, parameter, data) => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  const bins = parsed === null ? null : normalizeChartHistogramBins(parsed);
  if (bins === undefined) return null;
  return sheetChartModelEdited(target, data, (attrs) => chartWithHistogramBins(attrs, bins));
};

/** Format Data Series > Marker > Marker Options size bounds (points). */
export const CHART_SERIES_MARKER_MIN_SIZE = 2;
export const CHART_SERIES_MARKER_MAX_SIZE = 72;

/**
 * Format Data Series > Marker > Marker Options on one series of a line chart
 * born on the sheet: `{seriesIndex, symbol?, size?}` as JSON. A new symbol
 * keeps the series' size (6 when it has none; "none" takes no size); a size
 * alone needs a built-in marker the series already states, as the pane's
 * Size box does.
 */
export const chartSeriesMarkerButton: SheetChartButton = (target, parameter, data) => {
  let parsed: { seriesIndex?: unknown; symbol?: unknown; size?: unknown };
  try {
    parsed = JSON.parse(parameter) as typeof parsed;
  } catch {
    return null;
  }
  const { seriesIndex, symbol, size } = parsed;
  if (!Number.isInteger(seriesIndex) || (seriesIndex as number) < 0) return null;
  if (symbol !== undefined && !(SPREADSHEET_LINE_MARKER_SYMBOLS as readonly unknown[]).includes(symbol)) return null;
  if (
    size !== undefined &&
    (!Number.isInteger(size) || (size as number) < CHART_SERIES_MARKER_MIN_SIZE || (size as number) > CHART_SERIES_MARKER_MAX_SIZE)
  ) {
    return null;
  }
  if (symbol === undefined && size === undefined) return null;
  if (chartTypeOf(target) !== "line") return null;
  return sheetChartModelEdited(target, data, (attrs) => {
    const prior = (attrs.format?.seriesStyles as readonly Record<string, unknown>[] | undefined)
      ?.find((entry) => entry.seriesIndex === seriesIndex)?.marker as ChartSeriesMarker | undefined;
    const marker: ChartSeriesMarker | null =
      symbol === "none"
        ? ({ symbol: "none" } as ChartSeriesMarker)
        : symbol !== undefined
          ? ({ symbol, size: (size as number | undefined) ?? prior?.size ?? 6 } as ChartSeriesMarker)
          : prior && prior.symbol !== "none"
            ? ({ symbol: prior.symbol, size } as ChartSeriesMarker)
            : null;
    return marker ? chartWithSeriesMarker(attrs, seriesIndex as number, marker) : null;
  });
};

/**
 * Format Data Labels on one series' labels or on one point's label (select
 * the series' labels, click one label): `{target: {seriesIndex, pointIndex?},
 * patch}` as JSON, `patch` null to make the target follow the chart again.
 * The pane's dataLabelOverride intent, through the one chart editor.
 */
export const chartDataLabelOverrideButton: SheetChartButton = (target, parameter, data) => {
  let parsed: { target?: unknown; patch?: unknown };
  try {
    parsed = JSON.parse(parameter) as typeof parsed;
  } catch {
    return null;
  }
  const label = parsed.target as ChartDataLabelTarget | undefined;
  if (!label || typeof label !== "object" || !Number.isInteger(label.seriesIndex) || label.seriesIndex < 0) return null;
  if (label.pointIndex !== undefined && (!Number.isInteger(label.pointIndex) || label.pointIndex < 0)) return null;
  const patch = parsed.patch;
  if (patch !== null && (typeof patch !== "object" || Array.isArray(patch))) return null;
  // A field sent as null is one the pane took off.
  const restored = patch === null
    ? null
    : (Object.fromEntries(
        Object.entries(patch as Record<string, unknown>).map(([key, value]) => [key, value === null ? undefined : value]),
      ) as ChartDataLabelOverridePatch);
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithElementIntent(attrs, { element: "dataLabelOverride", target: label, patch: restored }),
  );
};

/** Hidden and Empty Cells > Show empty cells as (gap, zero, span). */
export const chartDisplayBlanksButton: SheetChartButton = (target, parameter, data) => {
  // Hidden and Empty Cells ▸ Show empty cells as: the chart's
  // `c:dispBlanksAs`, through the one chart editor.
  const value = parameter;
  if (!isChartDisplayBlanksAs(value)) return null;
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithElementIntent(attrs, { element: "displayBlanksAs", value }),
  );
};

/** Data Labels > a position ("automatic", "center", ...) or "callout". */
export const chartDataLabelPositionButton: SheetChartButton = (target, parameter, data) => {
  // Excel's Data Labels > position pick: labels on, at `parameter`
  // ("automatic" is where Office puts them). Which key the pick lands
  // on -- `dataLabelPosition` on the families that write `c:dLbls` /
  // `cx:dataLabels`, `dataLabelRegionLayout` on a filled map, whose
  // labels are Office's Map labels tri-state -- is decided once for
  // the three editors by `chart-elements.chartWithDataLabels`.
  // "callout" is Data Labels > Data Callout (Excel's bubble labels).
  const callout = parameter === "callout";
  const placement =
    parameter === "automatic" || callout
      ? null
      : (parameter as ChartDataLabelPlacement);
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithDataLabels(attrs, true, placement, callout),
  );
};

/** Add Chart Element > Axes: "x:true" / "y:false" one axis, "true" / "false" both. */
export const chartAxesButton: SheetChartButton = (target, parameter, data) => {
  if (!spreadsheetChartTypeHasAxes(chartTypeOf(target))) return null;
  // Add Chart Element > Axes > Primary Horizontal / Primary Vertical
  // ("x:true", "y:false"): one axis; "true" / "false": both.
  // Axes > Secondary Vertical ("y2:false") and Depth Axis ("z:false"):
  // the model's own switches.
  const secondary = /^(y2|z):(true|false)$/.exec(parameter);
  if (secondary) {
    return sheetChartModelEdited(target, data, (attrs) =>
      chartWithAxis(attrs, secondary[1] as "y2" | "z", secondary[2] === "true"),
    );
  }
  const axisPick = spreadsheetChartAxisPick(parameter);
  if (axisPick) {
    return sheetChartModelEdited(target, data, (attrs) =>
      chartWithAxis(attrs, axisPick.axis, axisPick.on),
    );
  }
  const next = parameter === "true";
  if (next === chartElementDisplayState(fieldsOf(target)).showAxisLabels) {
    return null;
  }
  return sheetChartModelEdited(target, data, (attrs) => {
    let changed = attrs;
    for (const axis of ["x", "y"] as const) {
      if (!chartHasAxis(changed.chartType, axis)) continue;
      changed = chartWithAxis(changed, axis, next) ?? changed;
    }
    return changed === attrs ? null : changed;
  });
};

/** Add Chart Element > Gridlines: "x:true", "minor-y:false", or the switch "true" / "false". */
export const chartGridlinesButton: SheetChartButton = (target, parameter, data) => {
  if (!spreadsheetChartTypeHasGridlines(chartTypeOf(target))) return null;
  // Gridlines > Primary Minor Horizontal / Vertical ("minor-x:true").
  const minorPick = parameter.startsWith("minor-")
    ? spreadsheetChartAxisPick(parameter.slice(6))
    : null;
  if (minorPick) {
    return sheetChartModelEdited(target, data, (attrs) =>
      chartWithElementIntent(attrs, {
        element: "gridlines",
        axis: minorPick.axis,
        on: minorPick.on,
        minor: true,
      }),
    );
  }
  // Gridlines > Primary Major Horizontal / Vertical, one axis each.
  const gridPick = spreadsheetChartAxisPick(parameter);
  if (gridPick) {
    return sheetChartModelEdited(target, data, (attrs) =>
      chartWithGridlines(attrs, gridPick.axis, gridPick.on),
    );
  }
  // The Gridlines switch: on is the value axis' major gridlines (the
  // horizontal ones, a bar chart's vertical ones); off is every one.
  const next = parameter === "true";
  if (next === chartElementDisplayState(fieldsOf(target)).showGridlines) {
    return null;
  }
  return sheetChartModelEdited(target, data, (attrs) => {
    if (next) return chartWithGridlines(attrs, attrs.chartType === "bar" ? "x" : "y", true);
    let changed = attrs;
    for (const axis of ["x", "y"] as const) {
      changed = chartWithGridlines(changed, axis, false) ?? changed;
    }
    return changed === attrs ? null : changed;
  });
};

/** Add Chart Element > Data Table: "true:keys", "true:no-keys", or the switch. */
export const chartDataTableButton: SheetChartButton = (target, parameter, data) => {
  if (!spreadsheetChartTypeSupportsDataTable(chartTypeOf(target))) return null;
  // Data Table > With Legend Keys ("true:keys") / No Legend Keys
  // ("true:no-keys"), or the switch ("true" / "false").
  if (parameter === "true:no-keys" || parameter === "true:keys") {
    return sheetChartModelEdited(target, data, (attrs) =>
      chartWithDataTable(attrs, true, parameter === "true:keys"),
    );
  }
  const next = parameter === "true";
  if (next === chartElementDisplayState(fieldsOf(target)).showDataTable) {
    return null;
  }
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithDataTable(attrs, next, attrs.format?.dataTableLegendKeys !== false),
  );
};

/**
 * Add Chart Element > Trendline: the trendline as JSON, or null to remove it
 * (the first series'). Format Data Series > Trendline names its series:
 * `{"seriesIndex": n, "trendline": {...} | null}`.
 */
export const chartTrendlineButton: SheetChartButton = (target, parameter, data) => {
  // An imported chart takes the element too: its part writer writes
  // it, and the projection's own reading of the part
  // (`nativeChartAnalysis`) is what lets the save tell a removal from
  // an older page's silence.
  const chart = fieldsOf(target);
  const imported = target.kind === "drawing";
  if (
    imported && !chart.pivotSource
      ? !importedChartStatesAnalysis(chart, "trendline")
      : false
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  // The series the trendline is based on, when the caller names one.
  let seriesIndex: number | null = null;
  if (
    parsed &&
    typeof parsed === "object" &&
    !Array.isArray(parsed) &&
    "seriesIndex" in parsed
  ) {
    const named = parsed as { seriesIndex?: unknown; trendline?: unknown };
    if (!Number.isInteger(named.seriesIndex)) return null;
    seriesIndex = named.seriesIndex as number;
    parsed = named.trendline ?? null;
  }
  const next = normalizeTrendline(parsed);
  if (parsed !== null && next === undefined) return null;
  const trendline = next ? spreadsheetTrendlineContract(next) : null;
  if (next && !trendline) return null;
  return sheetChartModelEdited(target, data, (attrs) =>
    seriesIndex === null
      ? chartWithTrendline(attrs, trendline ?? null)
      : chartWithSeriesTrendline(attrs, seriesIndex, trendline ?? null),
  );
};

/** Add Chart Element > Error Bars: the error bars as JSON, or null to remove them. */
export const chartErrorBarsButton: SheetChartButton = (target, parameter, data) => {
  // An imported chart takes the element too: its part writer writes
  // it, and the projection's own reading of the part
  // (`nativeChartAnalysis`) is what lets the save tell a removal from
  // an older page's silence.
  const chart = fieldsOf(target);
  const imported = target.kind === "drawing";
  if (
    imported && !chart.pivotSource
      ? !importedChartStatesAnalysis(chart, "errorBars")
      : false
  ) {
    return null;
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  const next = normalizeErrorBars(parsed);
  if (parsed !== null && next === undefined) return null;
  // Format Error Bars > Error Amount, on the model with its amount.
  const amount = next ? sheetChartErrorBarsOfSpec(next) : null;
  if (next && !amount) return null;
  return sheetChartModelEdited(
    target,
    data,
    (attrs) => chartWithErrorBars(attrs, amount ?? null),
  );
};

/** Add Chart Element > Up/Down Bars ("true" / "false"). */
export const chartUpDownBarsButton: SheetChartButton = (target, parameter, data) => {
  // Which charts take up/down bars (a 2-D line chart of two or more
  // series on one axis) is the shared element rule's.
  const next = parameter === "true";
  return sheetChartModelEdited(target, data, (attrs) =>
    chartWithUpDownBars(attrs, next),
  );
};

/** Add Chart Element > Lines: {"dropLines":true,...} as JSON, or null for None. */
export const chartLinesButton: SheetChartButton = (target, parameter, data) => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  const requested =
    parsed === null
      ? {}
      : parsed && typeof parsed === "object" && !Array.isArray(parsed)
        ? (parsed as Record<string, unknown>)
        : null;
  if (
    !requested ||
    Object.keys(requested).some(
      (key) =>
        !["dropLines", "highLowLines", "seriesLines"].includes(key),
    ) ||
    Object.values(requested).some((value) => value !== true)
  ) {
    return null;
  }
  // Each line through the shared command (which charts take which
  // line is the shared element rule's), or Lines > None.
  return sheetChartModelEdited(target, data, (attrs) => {
    if (!Object.keys(requested).length) return chartWithoutLines(attrs);
    let changed = attrs;
    for (const [line, key] of [["drop", "dropLines"], ["highLow", "highLowLines"], ["series", "seriesLines"]] as const) {
      const on = requested[key] === true;
      if ((changed.format?.[key] === true) === on) continue;
      const next = chartWithLine(changed, line, on);
      if (!next) return null;
      changed = next;
    }
    return changed === attrs ? null : changed;
  });
};

/** Add Chart Element > Axis Titles > Primary Horizontal. */
export const chartAxisTitleXButton = chartAxisTitleButton("x");
/** Add Chart Element > Axis Titles > Primary Vertical. */
export const chartAxisTitleYButton = chartAxisTitleButton("y");
/** Add Chart Element > Axis Titles > Depth: a 3-D chart's depth (series)
 * axis title (`format.seriesAxisTitle`, `c:serAx/c:title`), the words to
 * show it with, or "" to take it off. Which charts have a depth axis, and
 * that a title taken off takes its hand placement with it, is the shared
 * element rule's (Docs and the deck turn the same row through it). */
export const chartAxisTitleDepthButton: SheetChartButton = (target, parameter, data) => {
  const words = parameter.trim();
  return sheetChartModelEdited(target, data, (attrs) => {
    const shown = attrs.format?.seriesAxisTitle;
    if (words && shown) {
      // A title already shown takes the new words where it is.
      const text = clampChartText(words);
      if (text === shown) return null;
      const next: ChartBlockAttrsV2 = { ...attrs, format: { ...attrs.format, seriesAxisTitle: text } };
      return validateChartBlockAttrsV2(next).valid ? next : null;
    }
    return chartWithElementIntent(attrs, { element: "axisTitle", axis: "z", on: Boolean(words), text: words });
  });
};

/** Delete a chart born on the sheet: the sheet without it (the key goes
 * with the last one); null when it is not among `charts`. */
export function sheetWithoutChart<S extends object>(
  sheet: S,
  charts: readonly { id: string }[],
  id: string,
): S | null {
  const next = charts.filter((c) => c.id !== id);
  if (next.length === charts.length) return null;
  const nextSheet = { ...sheet } as S & { charts?: readonly { id: string }[] };
  if (next.length > 0) nextSheet.charts = next;
  else delete nextSheet.charts;
  return nextSheet;
}
