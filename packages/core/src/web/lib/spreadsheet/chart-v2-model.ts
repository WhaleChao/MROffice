/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet chart on the ONE chart model every editor shares.
 *
 * A sheet chart -- authored over a range or imported from a workbook -- is
 * `chartV2: ChartBlockAttrsV2` (the kind, titles, format and a cache of the
 * plotted data) plus a `binding` (where its data lives: a rectangle of one
 * sheet plotted by rows or columns, or the per-series references an imported
 * chart states). See `SpreadsheetChart` in ./chart.
 *
 * This module is the bridge between that model and the flat fields the record
 * carried before it (`type`, `showLegend`, `chartFormat`, ...), which the XLSX
 * writers and the not-yet-migrated call sites still read:
 *
 * - {@link spreadsheetChartV2Reading} is the MIGRATION RULE: a record without
 *   `chartV2` is read as the attrs the canvas has been drawing all along
 *   (`buildSpreadsheetChartV2Plan(...).attrs`), so nothing a chart drew changes.
 *   It runs in memory on load; the record gains `chartV2` on its next write.
 * - {@link legacySpreadsheetChartFields} derives the flat fields from `chartV2`
 *   + `binding`, and {@link withSpreadsheetChartV2} is the one write path that
 *   stores the model and the derived fields together, so they never disagree.
 * - Every sheet surface draws the STORED model ({@link spreadsheetChartStoredModel},
 *   {@link spreadsheetChartModelInput}): the planner is handed `chartV2` itself,
 *   its data read live from the bound cells, as Docs and the deck hand theirs.
 *   The derived flat fields remain for the XLSX writers, the sheet's chrome
 *   around the plot (header, axis-title strips, the frame's colours) and the
 *   sheet's display patches over the plan.
 *
 * Every sheet chart command edits the model and commits it through
 * {@link withSpreadsheetChartV2}; the flat fields are only ever derived. The
 * sheet's own settings are the model's keys too: `chartStyle` (c:style) is
 * `stylePreset` through the one preset table, markers and smoothing are
 * `format.seriesStyles[i].marker` / `.smooth`, a line's width
 * `format.seriesStyles[i].shapePaint.outlineWidthPt`, a series gradient its
 * `shapePaint.fillGradient`, an error amount `format.errorBars`, and the value
 * axis' number format `format.axisOptions.value.numberFormat` (absent = Linked
 * to source). The shared schema states every one of them; the shared
 * commands (`chartWithSeriesMarker`, `chartWithErrorBars`, ...) edit them.
 */

import {
  chartSeriesRenderAsOf,
  chartSeriesTakesMarker,
  chartSeriesTakesSmooth,
  chartSupportsAxisTitle,
  chartSupportsAxisVisibility,
  chartSupportsDataLabels,
  chartDataTakesErrorBars,
  chartSupportsGridlines,
  chartSupportsTrendline,
  createChartBlockAttrsV2,
  normalizeChartErrorBars,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartBlockFormat,
  type ChartErrorBars,
  type WaterfallPointKind,
} from "../../components/console/editor/extensions/chartBlockShared";
import { withOfficeChartAxisOptions } from "../office-charts/axis-options";
import { officeChartFormatWithRowsAndColumnsSwitched } from "../office-charts/chart-filters";
import { officeChartWithType } from "../office-charts/chart-type";
import {
  CHART_STYLE_PRESETS,
  DEFAULT_CHART_STYLE_PRESET_ID,
} from "../office-charts/chart-style-presets.generated";
import type {
  PresentationChartSeriesStyle,
  PresentationChartSliceStyle,
} from "../presentation-chart-area-style";
import {
  buildChartSeries,
  nativeChartHeaderRow,
  buildChartSeriesFromRefs,
  buildChartXYSeriesFromRefs,
  normalizeChartLineSeriesStyles,
  normalizeChartSeriesGradientStyles,
  normalizeChartSeriesRefs,
  normalizeChartValueAxisNumberFormat,
  normalizeErrorBars,
  normalizeSpreadsheetChartBinding,
  normalizeSpreadsheetChartType,
  normalizeTrendline,
  parseChartRefAreas,
  parseChartRefRange,
  spreadsheetChartStylePresetId,
  spreadsheetTrendlineContract,
  type ChartErrorBarSpec,
  type ChartSourceRange,
  type ChartTrendlineSpec,
  type MiniChartXYSeries,
  type SpreadsheetChart,
  type SpreadsheetChartBinding,
  type SpreadsheetChartLineSeriesStyle,
  type SpreadsheetChartSeriesGradientStyle,
  type SpreadsheetChartSeriesRef,
  type SpreadsheetChartValueAxisNumberFormat,
} from "./chart";
import {
  normalizeSpreadsheetChartFormat,
  spreadsheetChartLegacyDisplay,
  type SpreadsheetChartFormat,
} from "./chart-format";
import { excelBuiltInChartTextLook, isImportedClassicChartPart } from "./chart-builtin-text";
import { spreadsheetChartTitle } from "./chart-title";
import { drawingChartRangesOf } from "./drawing-chart-series";
import {
  comboSeriesXValuesOver,
  spreadsheetChartKindForV2,
  spreadsheetChartLeafOf,
  spreadsheetChartModelOverData,
  spreadsheetChartV2Drawn,
  type SpreadsheetChartMapRegionCode,
  type SpreadsheetChartV2Input,
} from "./chart-v2-binding";

export type {
  SpreadsheetChartBinding,
} from "./chart";
export {
  normalizeSpreadsheetChartBinding,
  normalizeSpreadsheetChartV2,
} from "./chart";

/**
 * Whether the chart's category axis is a date axis (Format Axis > Axis Type,
 * or Office's automatic choice of one): its categories are read as dates.
 */
export function spreadsheetChartHasDateAxis(chart: SpreadsheetChartLike): boolean {
  const typeOf = (format: unknown) =>
    (format as { axisOptions?: { category?: { axisType?: unknown } } } | undefined)?.axisOptions?.category?.axisType;
  return typeOf((chart.chartV2 as { format?: unknown } | undefined)?.format) === "date" || typeOf(chart.chartFormat) === "date";
}

/** Whether `value` is a chart the one model states (the shared validator). */
const validChartModel = (value: unknown): value is ChartBlockAttrsV2 =>
  validateChartBlockAttrsV2(value).valid;

/** A chart record of either collection, read loosely: the flat fields are `unknown`. */
export type SpreadsheetChartLike = { id: string; [key: string]: unknown };

/** What the reading needs beside the record: the live data of its bound range. */
export interface SpreadsheetChartV2Live {
  rows?: string[][];
  /** Whether rows[0] is the series-title row (rows from series references always are). */
  headerRow?: boolean;
  /**
   * A multi-level category axis over `rows` (a PivotChart with two or more
   * row fields): the levels above each row's leaf label, outermost first, one
   * entry per data row (after the header row), a group's label on its first
   * row and "" after it. The reading's `data.categoryLevels`; levels that do
   * not align with the rows are left out (`SpreadsheetChartV2Input`).
   */
  categoryLevels?: readonly (readonly string[])[];
  xySeries?: MiniChartXYSeries[];
  themeAccents?: readonly string[];
  /** The resolved format code of a value axis Linked to source. */
  valueAxisNumberFormat?: string;
  locale?: string;
  mapRegionCodes?: readonly SpreadsheetChartMapRegionCode[];
}

export type SpreadsheetChartCollectionKind = "native" | "drawing";

/** Which collection a record belongs to: an imported object says `type: "chart"`. */
export function spreadsheetChartCollectionKind(chart: SpreadsheetChartLike): SpreadsheetChartCollectionKind {
  return chart.type === "chart" ? "drawing" : "native";
}

/** The flat kind a record states, whichever collection it is in. */
export function spreadsheetChartFlatKind(chart: SpreadsheetChartLike): unknown {
  return spreadsheetChartCollectionKind(chart) === "drawing" ? chart.chartType : chart.type;
}

/**
 * The binding a record states through its flat fields: a range-bound chart's
 * rectangle (and orientation, and source sheet), an imported chart's series
 * references -- with the one rectangle they bound when every reference is on
 * one sheet, so the same chart can be re-read the way an authored one is.
 */
export function spreadsheetChartBindingOf(
  chart: SpreadsheetChartLike,
  options: { sheetIdByName?: (name: string) => string | undefined } = {},
): SpreadsheetChartBinding | undefined {
  const stored = normalizeSpreadsheetChartBinding(chart.binding);
  const flat = flatSpreadsheetChartBinding(chart, options);
  if (!stored) return flat;
  if (!flat || (stored as { literal?: unknown }).literal === true) return stored;
  // A structural edit (rows or columns inserted or deleted, cells shifted, a
  // chart moved to another sheet) moves the flat source -- the range, the
  // series references -- and not yet the binding: the binding follows it, the
  // last write winning as for every flat field.
  const native = spreadsheetChartCollectionKind(chart) === "native";
  const moved =
    !same(stored.seriesRefs, flat.seriesRefs) ||
    (native &&
      (!same(stored.range, flat.range) ||
        stored.sheetId !== flat.sheetId ||
        (stored.plotBy ?? "columns") !== (flat.plotBy ?? "columns")));
  if (!moved) return stored;
  const next: SpreadsheetChartBinding = { ...stored, ...flat };
  if (native && !flat.plotBy) delete next.plotBy;
  if (native && !flat.sheetId) delete next.sheetId;
  if (!flat.seriesRefs) delete next.seriesRefs;
  return next;
}

/** The binding a record states through its flat fields alone. */
function flatSpreadsheetChartBinding(
  chart: SpreadsheetChartLike,
  options: { sheetIdByName?: (name: string) => string | undefined },
): SpreadsheetChartBinding | undefined {
  const seriesRefs = normalizeChartSeriesRefs(chart.seriesRefs);
  if (spreadsheetChartCollectionKind(chart) === "native") {
    const range = normalizeSpreadsheetChartBinding({ range: chart.range })?.range;
    if (!range) return undefined;
    return {
      ...(typeof chart.sourceSheetId === "string" && chart.sourceSheetId.trim()
        ? { sheetId: chart.sourceSheetId.trim() }
        : {}),
      range,
      ...(chart.seriesInColumns === false ? { plotBy: "rows" as const } : {}),
      ...(seriesRefs ? { seriesRefs } : {}),
    };
  }
  if (!seriesRefs) return undefined;
  const bounded = referencesRectangle(seriesRefs);
  const sheetId = bounded ? options.sheetIdByName?.(bounded.sheetName) : undefined;
  return {
    ...(sheetId ? { sheetId } : {}),
    ...(bounded ? { range: bounded.range } : {}),
    seriesRefs,
  };
}

function referencesRectangle(
  refs: readonly SpreadsheetChartSeriesRef[],
): { sheetName: string; range: ChartSourceRange } | null {
  let sheetName: string | undefined;
  let rect: ChartSourceRange | undefined;
  for (const ref of refs) {
    for (const formula of [ref.title, ref.categories, ref.xValues, ref.values, ref.bubbleSizes]) {
      if (!formula) continue;
      const parsed = parseChartRefRange(formula);
      if (!parsed?.sheetName) return null;
      if (sheetName !== undefined && sheetName.toLowerCase() !== parsed.sheetName.toLowerCase()) return null;
      sheetName = parsed.sheetName;
      const { r1, c1, r2, c2 } = parsed.range;
      rect = rect
        ? {
            r1: Math.min(rect.r1, r1),
            c1: Math.min(rect.c1, c1),
            r2: Math.max(rect.r2, r2),
            c2: Math.max(rect.c2, c2),
          }
        : { r1, c1, r2, c2 };
    }
  }
  return sheetName !== undefined && rect ? { sheetName, range: rect } : null;
}

/**
 * What a record states through its flat fields, as the planner's input: the
 * reading the model migration makes ({@link spreadsheetChartV2Reading}), and
 * what a record not on the model yet is drawn from. Never the stored model.
 */
export function spreadsheetChartRecordInput(chart: SpreadsheetChartLike): SpreadsheetChartV2Input {
  const kind = spreadsheetChartFlatKind(chart);
  const record: Record<string, unknown> = { ...chart };
  delete record.storedModel;
  delete record.storedModelOwnData;
  return {
    ...(record as SpreadsheetChartV2Input),
    type: typeof kind === "string" ? kind : undefined,
    chartFormat: normalizeSpreadsheetChartFormat(chart.chartFormat),
    title: spreadsheetChartTitle({ title: chart.title, caption: chart.caption, nativeTitlePresent: chart.nativeTitlePresent }),
  };
}

/**
 * The record as the shared model: what `buildSpreadsheetChartV2Plan` reads
 * from the flat fields and the live range -- the reading the canvas draws --
 * with the chart's own title and axis titles, and every trendline the sheet
 * states (the adapter keeps the moving average out of the shared renderer's
 * plan; the model still carries it). Null when the data cannot be plotted.
 */
export function spreadsheetChartV2Reading(
  chart: SpreadsheetChartLike,
  live: SpreadsheetChartV2Live,
): ChartBlockAttrsV2 | null {
  const recorded = spreadsheetChartRecordInput(chart);
  const chartFormat = recorded.chartFormat;
  const title = recorded.title ?? "";
  const plan = spreadsheetChartV2Drawn({
    ...recorded,
    rows: live.rows,
    headerRow: live.headerRow,
    categoryLevels: live.categoryLevels,
    xySeries: live.xySeries,
    themeAccents: live.themeAccents,
    valueAxisNumberFormat: live.valueAxisNumberFormat,
    locale: live.locale,
    mapRegionCodes: live.mapRegionCodes ?? (chart.mapRegionCodes as SpreadsheetChartMapRegionCode[] | undefined),
    chartFormat,
    title,
  });
  if (!plan) return null;
  const base: ChartBlockAttrsV2 = { ...plan.attrs, title };
  const axisTitle = (value: unknown, axis: "x" | "y") =>
    typeof value === "string" && value.trim() && chartSupportsAxisTitle(base.chartType, axis)
      ? value
      : undefined;
  const axisTitleX = axisTitle(chart.axisTitleX, "x");
  const axisTitleY = axisTitle(chart.axisTitleY, "y");
  const format: ChartBlockFormat = { ...base.format };
  if (chartFormat?.layout) format.layout = chartFormat.layout;
  // A title the part states with no words of its own is Office's automatic
  // title (an Office 2016 chart's bare `cx:title`: 图表标题 over
  // C01_pareto-gapwidth's unnamed series).
  if (!title.trim() && chart.nativeTitlePresent === true) format.automaticTitle = true;
  const trendline = normalizeTrendline(chart.trendline);
  if (trendline && chartSupportsTrendline(base.chartType, base.chartSubtype)) {
    const contract = spreadsheetTrendlineContract(trendline);
    if (contract) format.trendline = contract;
  }
  const attrs: ChartBlockAttrsV2 = {
    ...base,
    ...(axisTitleX ? { axisTitleX } : {}),
    ...(axisTitleY ? { axisTitleY } : {}),
    ...(Object.keys(format).length ? { format } : {}),
  };
  const extended = withSheetChartExtensionsRead(chart, attrs);
  if (validChartModel(extended)) return extended;
  if (validChartModel(attrs)) return attrs;
  // A title or trendline the shared contract refuses on this family is left
  // out rather than losing the chart; the plan's own attrs are always valid.
  return validChartModel(base) ? base : null;
}

/**
 * A model read off a chart's part (an importer's shared reading), with what
 * the sheet's own reading of the same part states beyond it laid on (see
 * below): a series' marker or width the part reading left out is not lost
 * when the record's flat fields are derived from the model. The model as
 * given when the result would not validate.
 */
export function spreadsheetChartModelWithFlatReading(
  chart: SpreadsheetChartLike,
  attrs: ChartBlockAttrsV2,
): ChartBlockAttrsV2 {
  const next = withSheetChartExtensionsRead(chart, attrs);
  return next !== attrs && validChartModel(next) ? next : attrs;
}

/**
 * The sheet's own flat settings read onto the model's keys: each series'
 * marker, smoothing and line width (`lineSeriesStyles`), a column chart's
 * series gradients (`seriesGradientStyles`), an error amount outside the
 * shared recipes (`errorBars`) and a value axis number format that is not
 * Linked to source (`valueAxisNumberFormat`). What the chart's shared format
 * already states for a series or an axis wins, as it does on the canvas.
 */
function withSheetChartExtensionsRead(chart: SpreadsheetChartLike, attrs: ChartBlockAttrsV2): ChartBlockAttrsV2 {
  const format = { ...(attrs.format ?? {}) } as Record<string, unknown>;
  const count = "series" in attrs.data && Array.isArray(attrs.data.series) ? attrs.data.series.length : 0;
  const styles = new Map<number, Record<string, unknown>>(
    ((format.seriesStyles ?? []) as Record<string, unknown>[]).map((entry) => [entry.seriesIndex as number, { ...entry }]),
  );
  let touched = false;
  for (const style of normalizeChartLineSeriesStyles(chart.lineSeriesStyles) ?? []) {
    if (style.seriesIndex >= count) continue;
    const entry = styles.get(style.seriesIndex) ?? { seriesIndex: style.seriesIndex };
    const renderAs = chartSeriesRenderAsOf(attrs, style.seriesIndex);
    if (chartSeriesTakesMarker(attrs.chartType, attrs.chartSubtype, renderAs) && style.markerSymbol && entry.marker === undefined) {
      entry.marker = { symbol: style.markerSymbol, ...(style.markerSize !== undefined ? { size: style.markerSize } : {}) };
    }
    if (chartSeriesTakesSmooth(attrs.chartType, attrs.chartSubtype, renderAs) && style.smooth !== undefined && entry.smooth === undefined) {
      entry.smooth = style.smooth;
    }
    if (style.lineWidthPt !== undefined && entry.shapePaint === undefined) {
      entry.shapePaint = { outlineWidthPt: style.lineWidthPt };
    }
    if (Object.keys(entry).length > 1) {
      styles.set(style.seriesIndex, entry);
      touched = true;
    }
  }
  const gradients = normalizeChartSeriesGradientStyles(chart.seriesGradientStyles);
  if (
    gradients &&
    attrs.chartType === "column" &&
    gradients.length === count &&
    gradients.every((style, index) => style.seriesIndex === index)
  ) {
    for (const gradient of gradients) {
      const entry = styles.get(gradient.seriesIndex) ?? { seriesIndex: gradient.seriesIndex };
      // A flat fill chosen for the series beats the imported gradient.
      if (entry.fill !== undefined || entry.shapePaint !== undefined) continue;
      entry.fill = gradient.from;
      entry.outline = entry.outline ?? "none";
      entry.shapePaint = {
        outlineWidthPt: 0,
        fillGradient: {
          angle: 0,
          stops: [
            { position: 0, color: gradient.from },
            { position: 1, color: gradient.to },
          ],
        },
      };
      styles.set(gradient.seriesIndex, entry);
      touched = true;
    }
  }
  if (touched) format.seriesStyles = [...styles.values()].sort((a, b) => (a.seriesIndex as number) - (b.seriesIndex as number));
  const errorBars = normalizeErrorBars(chart.errorBars);
  if (errorBars && format.errorBars === undefined && chartDataTakesErrorBars(attrs.chartType, attrs.chartSubtype, attrs.data)) {
    const amount = sheetChartErrorBarsOfSpec(errorBars);
    if (amount) {
      format.errorBars = amount;
      touched = true;
    }
  }
  let next: ChartBlockAttrsV2 = touched ? { ...attrs, format: format as ChartBlockFormat } : attrs;
  const numberFormat = normalizeChartValueAxisNumberFormat(chart.valueAxisNumberFormat);
  if (numberFormat && !numberFormat.sourceLinked && !next.format?.axisOptions?.value?.numberFormat) {
    next = withOfficeChartAxisOptions(next, "value", {
      ...(next.format?.axisOptions?.value ?? {}),
      numberFormat: numberFormat.formatCode,
    });
  }
  return next;
}

/** The sheet's error-bar spec as the model states it. */
export function sheetChartErrorBarsOfSpec(spec: ChartErrorBarSpec): ChartErrorBars | undefined {
  // Direction and End Style ride with every amount.
  const sides = {
    ...(spec.include ? { include: spec.include } : {}),
    ...(spec.noEndCap ? { noEndCap: true as const } : {}),
  };
  switch (spec.type) {
    case "stderr":
      return normalizeChartErrorBars({ type: "standardError", ...sides });
    case "percent":
      return normalizeChartErrorBars({ type: "percentage", value: spec.value, ...sides });
    case "stddev":
      return normalizeChartErrorBars({ type: "standardDeviation", value: spec.value, ...sides });
    case "fixed":
      return normalizeChartErrorBars({ type: "fixed", value: spec.value, ...sides });
  }
}

/** The model's error bars as the sheet's flat spec. */
function sheetErrorBarsSpec(value: unknown): ChartErrorBarSpec | undefined {
  const amount = normalizeChartErrorBars(value);
  if (!amount) return undefined;
  const sides = {
    ...(amount.include ? { include: amount.include } : {}),
    ...(amount.noEndCap ? { noEndCap: true as const } : {}),
  };
  if (amount.type === "standardError") return { type: "stderr", value: 0, ...sides };
  // Custom ▸ Specify Value has no single amount: the shared model alone
  // keeps it, and the workbook writers write it from the model.
  if (amount.type === "custom") return undefined;
  return {
    type: amount.type === "percentage" ? "percent" : amount.type === "standardDeviation" ? "stddev" : "fixed",
    value: amount.value,
    ...sides,
  };
}

/**
 * A combination chart's error bars stated by its record alone: the record
 * kept them beside its model while the shared contract refused a
 * combination's (before chart unification B8), and a record saved then still
 * does. They are the model's own now: drawn from the model (the reconcile on
 * load reads them into it, `withSheetChartExtensionsRead`), never as the
 * sheet's overlay beside it. The model with them, else as given.
 */
function comboModelWithRecordErrorBars(chart: SpreadsheetChartLike, model: ChartBlockAttrsV2): ChartBlockAttrsV2 {
  if (model.chartType !== "combo" || model.format?.errorBars !== undefined) return model;
  const spec = normalizeErrorBars(chart.errorBars);
  const amount = spec ? sheetChartErrorBarsOfSpec(spec) : undefined;
  if (!amount || !chartDataTakesErrorBars(model.chartType, model.chartSubtype, model.data)) return model;
  const next: ChartBlockAttrsV2 = { ...model, format: { ...model.format, errorBars: amount } as ChartBlockFormat };
  return validChartModel(next) ? next : model;
}

/** The keys of `chartFormat` (the sheet's envelope of the shared format). */
const SPREADSHEET_FORMAT_KEYS = new Set<string>([
  "chartAreaStyle", "plotAreaStyle", "chartTitleStyle", "legendStyle",
  "backWallStyle", "sideWallStyle", "floorStyle",
  "gridlineStyle", "dataLabelStyle", "dataLabelLeaderLineStyle", "dataTableStyle",
  "gridlineStyleX", "gridlineStyleY", "minorGridlineStyleX", "minorGridlineStyleY",
  "trendlineStyle", "errorBarStyle", "upBarStyle", "downBarStyle",
  "dropLineStyle", "highLowLineStyle", "seriesLineStyle",
  "dataLabels", "dataLabelValue", "dataLabelLegendKey", "dataLabelCategoryName",
  "dataLabelSeriesName", "dataLabelPercentage", "dataLabelBubbleSize", "dataLabelLeaderLines",
  "dataLabelCallout",
  "axisX", "axisY", "gridlinesX", "gridlinesY", "minorGridlinesX", "minorGridlinesY",
  "dataTable", "dataTableLegendKeys",
  "dataTableHorizontalBorder", "dataTableVerticalBorder", "dataTableOutline",
  "seriesStyles", "sliceStyles", "axisStyles", "dataLabelPosition", "dataLabelSeparator",
  "dataLabelRegionLayout", "binIntervalClosed", "parentLabelLayout", "histogramBins", "boxWhisker", "gapWidth", "displayBlanksAs", "legend", "axisOptions",
  "hiddenSeries", "hiddenCategories", "text", "layout",
]);

/** The flat fields {@link legacySpreadsheetChartFields} owns on a record. */
export const LEGACY_SPREADSHEET_CHART_FIELDS = [
  "type", "chartType", "chartSubtype", "title", "axisTitleX", "axisTitleY", "chartFormat",
  "showLegend", "legendPosition", "showAxisLabels", "showGridlines", "showDataLabels", "showDataTable",
  "trendline", "errorBars", "upDownBars", "dropLines", "highLowLines", "seriesLines",
  "seriesStyles", "sliceStyles", "secondaryAxisSeries", "comboSeriesTypes",
  "range", "seriesInColumns", "sourceSheetId", "seriesRefs", "chartColorPalette",
  "chartStyle", "lineSeriesStyles", "seriesGradientStyles", "valueAxisNumberFormat", "desktopStyleId",
] as const;

type LegacyFields = Partial<Record<(typeof LEGACY_SPREADSHEET_CHART_FIELDS)[number], unknown>> & {
  errorBars?: ChartErrorBarSpec;
  lineSeriesStyles?: readonly SpreadsheetChartLineSeriesStyle[];
  seriesGradientStyles?: readonly SpreadsheetChartSeriesGradientStyle[];
  valueAxisNumberFormat?: SpreadsheetChartValueAxisNumberFormat;
  nativeTitlePresent?: boolean;
};

const mergeByIndex = <T extends Record<string, unknown>>(
  base: readonly T[] | undefined,
  over: readonly T[] | undefined,
  key: string,
): T[] => {
  const merged = new Map<unknown, T>();
  for (const item of base ?? []) merged.set(item[key], item);
  for (const item of over ?? []) merged.set(item[key], { ...merged.get(item[key]), ...item });
  return [...merged.values()].sort((a, b) => Number(a[key]) - Number(b[key]));
};

// Equal as values: key order is a spelling (a normalizer may reorder keys).
const stable = (value: unknown): string =>
  JSON.stringify(value ?? null, (_key, item) =>
    item && typeof item === "object" && !Array.isArray(item)
      ? Object.fromEntries(Object.entries(item as Record<string, unknown>).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
      : item,
  );
const same = (left: unknown, right: unknown) => stable(left) === stable(right);

/**
 * The flat fields a record carries for the writers and readers that do not
 * read `chartV2` yet, derived from the shared model and its binding. A key is
 * present only when the model states it, so a caller replacing the owned
 * fields (see {@link withSpreadsheetChartV2}) drops what the model no longer
 * has. `previous` keeps a split the record already had (its explicit series
 * paint beside the effect-only entries of `chartFormat`) when the merged
 * paint is unchanged, so an untouched chart does not change on disk.
 */
export function legacySpreadsheetChartFields(
  attrs: ChartBlockAttrsV2,
  binding: SpreadsheetChartBinding | undefined,
  kind: SpreadsheetChartCollectionKind,
  previous?: SpreadsheetChartLike,
): LegacyFields {
  const fields: LegacyFields = {};
  // A series' marker, smoothing and a line width alone are the record's
  // `lineSeriesStyles`, not its `chartFormat` envelope.
  const model = attrs;
  const { shared, lineStyles } = sheetLineSeriesStylesOf(model);
  attrs = shared;
  const leaf = spreadsheetChartKindForV2(attrs.chartType, attrs.chartSubtype);
  const format = attrs.format ?? {};
  const series = attrs.data.kind === "category" ? attrs.data.series : [];
  // A combo the adapter derived from a secondary axis on one plain kind reads
  // back as that kind + `secondaryAxisSeries`, the way it was authored.
  const renderKinds = new Set(series.map((entry) => entry.renderAs ?? "column"));
  const secondary = series.flatMap((entry, index) => (entry.axis === "secondary" ? [index] : []));
  if (attrs.chartType === "combo" && attrs.chartSubtype === "custom" && renderKinds.size === 1 && secondary.length) {
    const [only] = [...renderKinds];
    fields[kind === "native" ? "type" : "chartType"] = only;
    fields.secondaryAxisSeries = secondary;
  } else if (leaf) {
    // The record's own spelling of the same leaf stays (a `stacked-column`
    // is the column family's stacked leaf as much as `column` + "stacked").
    const previousKind = previous ? normalizeSpreadsheetChartType(spreadsheetChartFlatKind(previous)) : undefined;
    const previousSubtype = typeof previous?.chartSubtype === "string" ? previous.chartSubtype : undefined;
    const previousLeaf = previousKind && previousKind !== "camera"
      ? spreadsheetChartLeafOf({ type: previousKind, chartSubtype: previousSubtype })
      : null;
    if (previousLeaf && previousLeaf.family === attrs.chartType && previousLeaf.subtype === attrs.chartSubtype) {
      fields[kind === "native" ? "type" : "chartType"] = spreadsheetChartFlatKind(previous!);
      if (previousSubtype) fields.chartSubtype = previousSubtype;
    } else {
      fields[kind === "native" ? "type" : "chartType"] = leaf.type;
      if (leaf.chartSubtype) fields.chartSubtype = leaf.chartSubtype;
    }
    if (leaf.secondaryAxisLastSeries && series.length) fields.secondaryAxisSeries = [series.length - 1];
    if (attrs.chartType === "combo" && attrs.chartSubtype === "custom") {
      if (secondary.length) fields.secondaryAxisSeries = secondary;
      // An imported chart states each series' type in its references instead.
      if (!binding?.seriesRefs && series.length >= 2) fields.comboSeriesTypes = series.map((entry) => entry.renderAs ?? "column");
    }
  }
  if (attrs.title) fields.title = attrs.title;
  if (kind === "drawing") {
    fields.title = attrs.title;
    if (previous?.type === "chart" || typeof previous?.nativeTitlePresent === "boolean") {
      fields.nativeTitlePresent = Boolean(attrs.title.trim());
    }
  }
  if (attrs.axisTitleX) fields.axisTitleX = attrs.axisTitleX;
  if (attrs.axisTitleY) fields.axisTitleY = attrs.axisTitleY;

  // The sheet's envelope of the shared format, and the switches it projects.
  const picked: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(format)) {
    if (SPREADSHEET_FORMAT_KEYS.has(key) && value !== undefined) picked[key] = value;
  }
  // The switches whose flat fields default otherwise than the model does (a
  // flat chart without them draws both gridlines, labels off, the legend on
  // top): stated as the model draws them, so the record reads back the same.
  const look = CHART_STYLE_PRESETS[(model.stylePreset ?? DEFAULT_CHART_STYLE_PRESET_ID) as keyof typeof CHART_STYLE_PRESETS]
    ?? CHART_STYLE_PRESETS[DEFAULT_CHART_STYLE_PRESET_ID];
  if (chartSupportsAxisVisibility(attrs.chartType)) {
    picked.axisX ??= true;
    picked.axisY ??= true;
  }
  if (chartSupportsGridlines(attrs.chartType)) {
    picked.gridlinesX ??= false;
    picked.gridlinesY ??= false;
  }
  if (attrs.chartType !== "map") picked.dataLabels ??= chartSupportsDataLabels(attrs.chartType) && look.dataLabels;
  picked.legend ??= look.legendPosition;
  // Series and slice paint: keep the record's own split when nothing changed.
  const previousFormat = normalizeSpreadsheetChartFormat(previous?.chartFormat);
  const previousSeries = mergeByIndex(
    previous?.seriesStyles as readonly PresentationChartSeriesStyle[] | undefined,
    previousFormat?.seriesStyles,
    "seriesIndex",
  );
  const previousSlices = mergeByIndex(
    previous?.sliceStyles as readonly PresentationChartSliceStyle[] | undefined,
    previousFormat?.sliceStyles,
    "pointIndex",
  );
  if (previous && same(previousSeries, format.seriesStyles ?? [])) {
    if (previous.seriesStyles !== undefined) fields.seriesStyles = previous.seriesStyles;
    if (previousFormat?.seriesStyles) picked.seriesStyles = previousFormat.seriesStyles;
    else delete picked.seriesStyles;
  }
  if (previous && same(previousSlices, format.sliceStyles ?? [])) {
    if (previous.sliceStyles !== undefined) fields.sliceStyles = previous.sliceStyles;
    if (previousFormat?.sliceStyles) picked.sliceStyles = previousFormat.sliceStyles;
    else delete picked.sliceStyles;
  }
  const chartFormat = Object.keys(picked).length ? normalizeSpreadsheetChartFormat(picked) : undefined;
  if (chartFormat) fields.chartFormat = chartFormat;
  Object.assign(fields, spreadsheetChartLegacyDisplay(chartFormat as SpreadsheetChartFormat | undefined));
  // Spelled as the page normalizers keep them: labels and a data table are
  // stated only when on, so a saved and reopened record still derives equal.
  if (fields.showDataLabels === false) delete fields.showDataLabels;
  if (fields.showDataTable === false) delete fields.showDataTable;

  if (format.trendline) {
    const { type, ...options } = format.trendline as { type: ChartTrendlineSpec["type"] } & Record<string, unknown>;
    fields.trendline = { type, ...options } as ChartTrendlineSpec;
  }
  const errorBars = sheetErrorBarsSpec(format.errorBars);
  if (errorBars) fields.errorBars = errorBars;
  for (const key of ["upDownBars", "dropLines", "highLowLines", "seriesLines"] as const) {
    if (format[key] === true) fields[key] = true;
  }
  // Change Colors: the model's palette is the sheet's pick.
  if (format.colorPalette) fields.chartColorPalette = format.colorPalette;
  // Chart Styles: the model's preset, spelled as the record's own `c:style`
  // while that still reads as the same preset (an imported part's style 26
  // stays 26; an Office 2016 chart without one keeps its native look).
  const preset = model.stylePreset ?? DEFAULT_CHART_STYLE_PRESET_ID;
  if (previous && spreadsheetChartStylePresetId(previous.chartStyle) === preset) {
    if (previous.chartStyle !== undefined) fields.chartStyle = previous.chartStyle;
  } else if (preset in CHART_STYLE_PRESETS) {
    fields.chartStyle = CHART_STYLE_PRESETS[preset as keyof typeof CHART_STYLE_PRESETS].cStyle;
  }
  // Markers, smoothing and a line's width alone, per series.
  const lineSeriesStyles = normalizeChartLineSeriesStyles(lineStyles);
  if (lineSeriesStyles) fields.lineSeriesStyles = lineSeriesStyles;
  // A series gradient of two stops across the bar.
  const gradients = (format.seriesStyles ?? []).flatMap((entry) => {
    const gradient = entry.shapePaint?.fillGradient;
    if (!gradient || gradient.angle !== 0 || gradient.stops.length !== 2) return [];
    const [from, to] = gradient.stops;
    if (from.position !== 0 || to.position !== 1 || from.alpha !== undefined || to.alpha !== undefined) return [];
    return [{ seriesIndex: entry.seriesIndex, from: from.color, to: to.color }];
  });
  const seriesGradientStyles = normalizeChartSeriesGradientStyles(gradients);
  if (seriesGradientStyles) fields.seriesGradientStyles = seriesGradientStyles;
  // The value axis' number format: a code, or Linked to source (the record
  // keeps the resolved linked code it already carries).
  const numberFormat = format.axisOptions?.value?.numberFormat;
  const previousNumberFormat = normalizeChartValueAxisNumberFormat(previous?.valueAxisNumberFormat);
  if (numberFormat) {
    const custom = normalizeChartValueAxisNumberFormat({ formatCode: numberFormat, sourceLinked: false });
    if (custom) fields.valueAxisNumberFormat = custom;
  } else if (previousNumberFormat?.sourceLinked) {
    fields.valueAxisNumberFormat = previous!.valueAxisNumberFormat as SpreadsheetChartValueAxisNumberFormat;
  }

  if (binding) {
    if (kind === "native") {
      if (binding.range) fields.range = binding.range;
      if (binding.plotBy === "rows") fields.seriesInColumns = false;
      if (binding.sheetId) fields.sourceSheetId = binding.sheetId;
    }
    if (binding.seriesRefs) fields.seriesRefs = binding.seriesRefs;
  }
  return fields;
}

/**
 * The model's per-series line options as the record's flat
 * `lineSeriesStyles` (marker type and size, Smoothed line, a width on a
 * series that states no line colour), and the model without them for the
 * record's `chartFormat` envelope. A model whose series state none comes back
 * as itself.
 */
function sheetLineSeriesStylesOf(attrs: ChartBlockAttrsV2): {
  shared: ChartBlockAttrsV2;
  lineStyles: SpreadsheetChartLineSeriesStyle[];
} {
  const format = attrs.format as Record<string, unknown> | undefined;
  const styles = format?.seriesStyles;
  if (!Array.isArray(styles)) return { shared: attrs, lineStyles: [] };
  const lineStyles: SpreadsheetChartLineSeriesStyle[] = [];
  const kept: Record<string, unknown>[] = [];
  let changed = false;
  for (const raw of styles as Record<string, unknown>[]) {
    if (!raw || typeof raw !== "object" || !Number.isInteger(raw.seriesIndex)) {
      kept.push(raw);
      continue;
    }
    const entry: Record<string, unknown> = { ...raw };
    const style: SpreadsheetChartLineSeriesStyle = { seriesIndex: entry.seriesIndex as number };
    if ("marker" in entry) {
      const marker = entry.marker as { symbol?: unknown; size?: unknown } | undefined;
      if (marker && typeof marker.symbol === "string") {
        style.markerSymbol = marker.symbol as SpreadsheetChartLineSeriesStyle["markerSymbol"];
        if (typeof marker.size === "number" && marker.symbol !== "none") style.markerSize = marker.size;
      }
      delete entry.marker;
    }
    if ("smooth" in entry) {
      if (typeof entry.smooth === "boolean") style.smooth = entry.smooth;
      delete entry.smooth;
    }
    const paint = entry.shapePaint as Record<string, unknown> | undefined;
    if (
      paint && entry.outline === undefined && Object.keys(paint).length === 1 &&
      typeof paint.outlineWidthPt === "number" && paint.outlineWidthPt > 0 && paint.outlineWidthPt <= 24
    ) {
      style.lineWidthPt = paint.outlineWidthPt;
      delete entry.shapePaint;
    }
    if (Object.keys(style).length > 1) {
      changed = true;
      lineStyles.push(style);
    }
    if (Object.keys(entry).some((key) => key !== "seriesIndex")) kept.push(entry);
    else changed = true;
  }
  if (!changed) return { shared: attrs, lineStyles };
  const nextFormat: Record<string, unknown> = { ...format };
  if (kept.length) nextFormat.seriesStyles = kept;
  else delete nextFormat.seriesStyles;
  const shared = { ...attrs } as ChartBlockAttrsV2;
  if (Object.keys(nextFormat).length) shared.format = nextFormat as ChartBlockFormat;
  else delete shared.format;
  return { shared, lineStyles };
}

/**
 * The record with `chartV2` (and its binding) as the one source of truth and
 * the flat fields derived from it, written together. Geometry, identity and
 * the native extensions ride along untouched.
 */
export function withSpreadsheetChartV2<T extends SpreadsheetChartLike>(
  chart: T,
  attrs: ChartBlockAttrsV2,
  binding?: SpreadsheetChartBinding,
): T & { chartV2: ChartBlockAttrsV2; binding?: SpreadsheetChartBinding } {
  const kind = spreadsheetChartCollectionKind(chart);
  const bound = binding ?? spreadsheetChartBindingOf(chart);
  const fields = legacySpreadsheetChartFields(attrs, bound, kind, chart);
  const next = { ...chart } as Record<string, unknown>;
  for (const key of LEGACY_SPREADSHEET_CHART_FIELDS) {
    // An imported object's `range` is the legacy editor's A1 text, not the
    // rectangle, and its `type` is the object kind ("chart"): neither is this
    // model's to rewrite.
    if (kind === "drawing" && (key === "range" || key === "type")) continue;
    delete next[key];
  }
  // A native chart always has a rectangle; a binding without one keeps it.
  if (kind === "native" && fields.range === undefined) next.range = chart.range;
  // A trendline or error bars the shared contract refuses on this family (a
  // combo, a stacked kind) are the sheet's own overlay: they stay beside the
  // model.
  if (!chartSupportsTrendline(attrs.chartType, attrs.chartSubtype) && chart.trendline !== undefined) {
    next.trendline = chart.trendline;
  }
  if (!chartDataTakesErrorBars(attrs.chartType, attrs.chartSubtype, attrs.data) && chart.errorBars !== undefined) {
    next.errorBars = chart.errorBars;
  }
  // A field whose derived value is what the record already holds keeps its
  // own object, so a write changes exactly what it changed.
  for (const [key, value] of Object.entries(fields)) {
    next[key] = same(chart[key], value) ? chart[key] : value;
  }
  // An imported chart's `ranges` are the formulas its series state; a write
  // that lays other series (Select Data, Switch Row/Column) restates them, so
  // a later row insert moves the references the chart now reads.
  if (kind === "drawing" && Array.isArray(next.seriesRefs) && !same(chart.seriesRefs, next.seriesRefs)) {
    next.ranges = drawingChartRangesOf(next.seriesRefs as SpreadsheetChartSeriesRef[]);
  }
  next.chartV2 = attrs;
  if (bound) next.binding = same(chart.binding, bound) ? chart.binding : bound;
  else delete next.binding;
  return next as T & { chartV2: ChartBlockAttrsV2; binding?: SpreadsheetChartBinding };
}

/** A stored cell as the page holds it: the value, and the caches the save stamps. */
export interface SpreadsheetChartStoredCell {
  value?: unknown;
  formula?: unknown;
  /** The on-screen string the last save stamped (formula result, number format applied). */
  display?: unknown;
  /** The unformatted numeric formula result the last save stamped. */
  chartValueCache?: { value?: unknown } | undefined;
}
export interface SpreadsheetChartStoredSheet {
  id: string;
  name: string;
  cells: Record<string, SpreadsheetChartStoredCell | undefined>;
}

/**
 * The live data of a chart read from STORED cells -- what a page carries when
 * it is opened, before an evaluator exists: a literal's value, a formula's
 * stamped result (`chartValueCache`, else its stamped display). The reading
 * the save makes stamps those caches first, so at save time it is exact; at
 * load time it is the last save's, and the next save refreshes it.
 */
export function spreadsheetChartLiveFromStoredCells(
  chart: SpreadsheetChartLike,
  sheet: SpreadsheetChartStoredSheet,
  sheets: readonly SpreadsheetChartStoredSheet[],
): SpreadsheetChartV2Live | null {
  const kind = spreadsheetChartCollectionKind(chart);
  const type = normalizeSpreadsheetChartType(spreadsheetChartFlatKind(chart));
  if (type === "camera" || chart.pivotSource) return null;
  const cellOf = (owner: SpreadsheetChartStoredSheet, row: number, col: number) => owner.cells[`${row}:${col}`];
  const display = (owner: SpreadsheetChartStoredSheet) => (row: number, col: number): string => {
    const cell = cellOf(owner, row, col);
    if (!cell) return "";
    if (typeof cell.display === "string") return cell.display;
    if (cell.formula) {
      // A formula's text is not what it shows: with no stamped display, its
      // stamped result (R454-X07: category labels =+C8+B10-B9 read as their
      // text on load, and an untouched save read the charts as changed).
      const cached = cell.chartValueCache?.value;
      return typeof cached === "number" && Number.isFinite(cached) ? String(cached) : "";
    }
    return String(cell.value ?? "");
  };
  const numeric = (owner: SpreadsheetChartStoredSheet) => (row: number, col: number): string => {
    const cell = cellOf(owner, row, col);
    if (!cell) return "";
    const cached = cell.chartValueCache?.value;
    if (typeof cached === "number" && Number.isFinite(cached)) return String(cached);
    if (cell.formula) return typeof cell.display === "string" ? cell.display : "";
    return String(cell.value ?? "");
  };
  if (kind === "native") {
    const source = typeof chart.sourceSheetId === "string"
      ? sheets.find((item) => item.id === chart.sourceSheetId)
      : sheet;
    if (!source) return null;
    const record = chart as unknown as SpreadsheetChart;
    const xySeries = record.seriesRefs && !record.sourceRefError
      ? buildChartXYSeriesFromRefs(record.seriesRefs, type, source.name, display(source), numeric(source))
      : [];
    const rows = record.range ? buildChartSeries({ ...record, type }, display(source), numeric(source)) : [];
    const headerRow = nativeChartHeaderRow(record);
    return { rows, xySeries, ...(headerRow !== undefined ? { headerRow } : {}) };
  }
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  if (!refs) {
    const data = chart.data as { rows?: unknown } | undefined;
    const rows = Array.isArray(data?.rows) ? (data!.rows as string[][]) : [];
    return rows.length ? { rows } : null;
  }
  const sourceName = parseChartRefAreas(refs[0]?.values)?.[0]?.sheetName ?? sheet.name;
  const source = sheets.find((item) => item.name === sourceName) ?? sheet;
  const owner = (name: string) => {
    const matches = sheets.filter((item) => item.name.toLowerCase() === name.toLowerCase());
    return matches.length === 1 ? matches[0] : undefined;
  };
  const workbook = {
    display: (name: string, row: number, col: number) => {
      const found = owner(name);
      return found ? display(found)(row, col) : undefined;
    },
    value: (name: string, row: number, col: number) => {
      const found = owner(name);
      return found ? numeric(found)(row, col) : undefined;
    },
  };
  const rows = buildChartSeriesFromRefs(refs, type, source.name, display(source), numeric(source), workbook, {
    dateCategories: spreadsheetChartHasDateAxis(chart),
  });
  const xySeries = buildChartXYSeriesFromRefs(refs, type, source.name, display(source), numeric(source), workbook);
  return { rows, xySeries, headerRow: rows.length > 0 ? true : undefined };
}

/**
 * Whether a record's flat fields still say what its `chartV2` says: false
 * when a writer that does not know the model wrote a flat field after it.
 * Both are compared as the flat fields they derive to, so a record migrated
 * without a rewrite (its switches implicit, the model's explicit) agrees.
 */
export function spreadsheetChartFlatFieldsAgree(
  chart: SpreadsheetChartLike,
  reading: ChartBlockAttrsV2,
): boolean {
  const attrs = chart.chartV2 as ChartBlockAttrsV2 | undefined;
  if (!attrs) return false;
  const kind = spreadsheetChartCollectionKind(chart);
  const binding = spreadsheetChartBindingOf(chart);
  return same(
    legacySpreadsheetChartFields(attrs, binding, kind, chart),
    legacySpreadsheetChartFields(reading, binding, kind, chart),
  );
}

/**
 * The record with its `chartV2` current: a record without one gains it (the
 * migration, in memory on load and on disk at save), a record whose flat
 * fields were written since by a writer that does not know the model is
 * re-read from them (the last write wins), and a record in agreement keeps
 * its model -- including what only the model states -- with its data cache
 * refreshed from the live range. A chart whose data cannot be read is left
 * as it is.
 *
 * An imported chart's model is the reading of its part (the server's
 * projection at import, the reading the deck and Docs store), which the
 * sheet's reading of the flat fields need not spell the same way: it is kept
 * as it is (its data cache refreshed). Every chart command edits the model
 * itself and derives the flat fields from it ({@link withSpreadsheetChartV2}),
 * so nothing here translates a flat write back into the model.
 */
/**
 * The live reading's data with what only the model states about each series
 * kept: a combination's per-series chart type and axis (`renderAs`, `axis`).
 * The flat fields name a combination by its catalog leaf, whose reading lays
 * the roles out by position (the last series a line on the secondary axis),
 * so refreshing the data from them moved Excel's "columns on the secondary
 * axis, lines on the primary" onto the leaf's own layout (X-050,
 * C14_testCombinedChartAxis).
 */
/** A waterfall's point kinds over refreshed `categories`: point by point
 * while the points are as many as before; after rows were added or removed,
 * each category takes the kind of the one of its name it had (in order, a
 * repeated name its next one), and a new category is a step. */
export function waterfallPointKindsOver(
  model: { categories: readonly string[]; pointKinds: readonly WaterfallPointKind[] },
  categories: readonly string[],
): WaterfallPointKind[] {
  if (model.pointKinds.length === categories.length) return [...model.pointKinds];
  const byName = new Map<string, WaterfallPointKind[]>();
  model.categories.forEach((name, index) => {
    const kind = model.pointKinds[index];
    if (kind) byName.set(name, [...(byName.get(name) ?? []), kind]);
  });
  return categories.map((name) => byName.get(name)?.shift() ?? "value");
}

/** A series whose name the file does not state (`nameUnstated`, the default
 *  "Series n" the reader gave it) keeps saying so while the name the cells
 *  give it is still that one. */
function seriesNameUnstatedOver(
  series: { name?: unknown; nameUnstated?: true },
  own: { name?: unknown; nameUnstated?: true },
): { nameUnstated: true } | null {
  return own.nameUnstated === true && series.nameUnstated === undefined && series.name === own.name
    ? { nameUnstated: true }
    : null;
}

function withModelSeriesRoles(model: ChartBlockAttrsV2, data: ChartBlockAttrsV2["data"]): ChartBlockAttrsV2["data"] {
  // A waterfall's totals (Set as Total, `cx:subtotals`) are the model's,
  // point by point: the cells hold the amounts only. Refreshed from them, an
  // imported waterfall's totals became steps on the page while its file
  // kept them (AI70H H133 A08), as the canvas already kept them
  // (`storedModelOverLiveData`).
  if (model.data.kind === "waterfall" && data.kind === "waterfall") {
    return { ...data, pointKinds: waterfallPointKindsOver(model.data, data.categories) };
  }
  if (model.data.kind === "xy" && data.kind === "xy") {
    // A scatter series on the secondary axis (Format Data Series > Plot
    // Series On) and its values' source format are the model's: the cells
    // hold the points only. Refreshed from them by an edit on another sheet,
    // the series fell to the primary axis and the save declined both charts
    // (xlsx-rounds R209-X01 E1).
    const stated = model.data.series;
    if (stated.length !== data.series.length) return data;
    return {
      ...data,
      series: data.series.map((series, index) => {
        const own = stated[index];
        const axis = own.axis !== undefined ? { axis: own.axis } : {};
        const numberFormat =
          own.numberFormat !== undefined && series.numberFormat === undefined ? { numberFormat: own.numberFormat } : {};
        const unstated = seriesNameUnstatedOver(series, own);
        return own.axis === undefined && !("numberFormat" in numberFormat) && !unstated
          ? series
          : { ...series, ...axis, ...numberFormat, ...unstated };
      }),
    };
  }
  if (model.data.kind === "sample" && data.kind === "sample") {
    // A statistical chart (box & whisker, histogram, Pareto) names each
    // series by a `_xlchart.v1.N` workbook name its part states, which the
    // cells cannot resolve: re-read, the name fell to "Series n" and every
    // save of a workbook with such a chart read as a change to it (window
    // load audit, xlsx-rounds R453-X15 "proportion_alive"). The model's name
    // stands where the cells give only the default.
    const stated = model.data.series;
    if (stated.length !== data.series.length) return data;
    let kept = false;
    const series = data.series.map((entry, index) => {
      const own = stated[index];
      if (entry.name !== `Series ${index + 1}` || !own || own.name === entry.name) return entry;
      kept = true;
      return { ...entry, name: own.name };
    });
    return kept ? { ...data, series } : data;
  }
  if (model.data.kind !== "category" || data.kind !== "category") return data;
  const stated = model.data.series;
  if (stated.length !== data.series.length) return data;
  return {
    ...data,
    series: data.series.map((series, index) => {
      const own = stated[index];
      const { xValues: _liveX, ...rest } = series;
      // A combination's scatter X values (`comboSeriesXValuesOver`).
      const xValues = comboSeriesXValuesOver(series, own);
      // The values' source format and a name the file does not state are the
      // model's too: the cells re-read neither. Refreshed away, every save of
      // a workbook with such a chart (R450-X03: 23 pie series, "0%") read as
      
      const numberFormat =
        own.numberFormat !== undefined && series.numberFormat === undefined ? { numberFormat: own.numberFormat } : {};
      const unstated = seriesNameUnstatedOver(series, own);
      return own.renderAs === undefined && own.axis === undefined && !xValues && !_liveX
        && !("numberFormat" in numberFormat) && !unstated
        ? series
        : {
            ...rest,
            ...(own.renderAs !== undefined ? { renderAs: own.renderAs } : {}),
            ...(own.axis !== undefined ? { axis: own.axis } : {}),
            ...(xValues ? { xValues } : {}),
            ...numberFormat,
            ...unstated,
          };
    }),
  };
}

/**
 * The flat fields that state a chart's look ({@link legacySpreadsheetChartFields}).
 * A record carrying none of them was written with its model alone -- a chart
 * the server authored (the AI workbook studio states `chartV2` from its
 * recipe: legend at the bottom, Colorful Palette 1) -- and no writer that
 * does not know the model wrote a look after it.
 */
const FLAT_LOOK_KEYS = [
  "chartFormat", "showLegend", "legendPosition", "showGridlines", "showAxisLabels", "showDataLabels",
  "showDataTable", "chartColorPalette", "chartStyle", "trendline", "errorBars", "upDownBars", "dropLines",
  "highLowLines", "seriesLines", "lineSeriesStyles", "seriesStyles", "sliceStyles", "seriesGradientStyles",
  "axisTitleX", "axisTitleY", "valueAxisNumberFormat",
] as const;

function spreadsheetChartStatesFlatLook(chart: SpreadsheetChartLike): boolean {
  const record = chart as Record<string, unknown>;
  return FLAT_LOOK_KEYS.some((key) => record[key] !== undefined);
}

export function reconcileSpreadsheetChartV2<T extends SpreadsheetChartLike>(
  chart: T,
  live: SpreadsheetChartV2Live | null,
  options: { sheetIdByName?: (name: string) => string | undefined } = {},
): T {
  if (!live) return chart;
  const current = chart.chartV2 as ChartBlockAttrsV2 | undefined;
  // A chart whose data is its own (`binding.literal`): the model IS the
  // chart, and there are no cells to re-read it from.
  if (current && normalizeSpreadsheetChartBinding(chart.binding)?.literal) return chart;
  const reading = spreadsheetChartV2Reading(chart, live);
  if (!reading) return chart;
  const binding = spreadsheetChartBindingOf(chart, options);
  if (!binding) return chart;
  // A kept model states the sheet's own flat settings it does not state yet
  // (an imported chart's markers, say, which the part's reading left out).
  const extended = (model: ChartBlockAttrsV2) => {
    const next = withSheetChartExtensionsRead(chart, model);
    return next !== model && validChartModel(next) ? next : model;
  };
  if (current && spreadsheetChartCollectionKind(chart) === "drawing" && !spreadsheetChartFlatFieldsAgree(chart, reading)) {
    const next = extended(
      reading.data.kind === current.data.kind
        ? { ...current, data: withModelSeriesRoles(current, reading.data) }
        : current,
    );
    if (same(next, current) && same(chart.binding, binding)) return chart;
    return { ...chart, chartV2: next, binding };
  }
  // A model written with no flat look beside it (a chart the server
  // authored) is the chart's look: re-reading the absent flat fields drew
  // every AI chart with the legend on top and without its palette, and
  
  if (
    (current && spreadsheetChartFlatFieldsAgree(chart, reading))
    || (current && reading.chartType === current.chartType && !spreadsheetChartStatesFlatLook(chart))
  ) {
    const refreshed = extended({ ...current, data: withModelSeriesRoles(current, reading.data) });
    if (same(refreshed, current) && same(chart.binding, binding)) return chart;
    return { ...chart, chartV2: refreshed, binding };
  }
  if (current && same(current, reading) && same(chart.binding, binding)) return chart;
  return { ...chart, chartV2: reading, binding };
}

/**
 * A chart that plots no series at all (a part with an empty plot group, no
 * series references, no PivotTable): Excel draws its chart area alone --
 * no automatic title, no axes, no words (C13_empty_chart, X-224). The sheet
 * and its print draw the empty frame.
 */
export function spreadsheetChartHasNoSeries(chart: SpreadsheetChartLike): boolean {
  const model = chart.chartV2;
  if (!model || typeof model !== "object" || Array.isArray(model)) return false;
  const data = (model as { data?: unknown }).data;
  if (!data || typeof data !== "object" || (data as { kind?: unknown }).kind !== "category") return false;
  const series = (data as { series?: unknown }).series;
  const refs = chart.seriesRefs;
  return (
    Array.isArray(series) &&
    series.length === 0 &&
    !(Array.isArray(refs) && refs.length > 0) &&
    (chart as { pivotSource?: unknown }).pivotSource === undefined
  );
}

/**
 * The record's stored model, when the shared validator admits it: what every
 * sheet surface draws (whatever its flat fields say -- they are only ever its
 * derivation, and the load's reconciliation has already read a flat write into
 * it) and what every chart command edits. Null for a record not on the model
 * yet (its data could not be read), which is drawn from its flat fields.
 */
export function spreadsheetChartStoredModel(chart: SpreadsheetChartLike): ChartBlockAttrsV2 | null {
  const attrs = chart.chartV2;
  return attrs && typeof attrs === "object" && !Array.isArray(attrs) && validChartModel(attrs) ? attrs : null;
}

/**
 * The planner's input for a record drawn from its model: the stored model
 * itself (`storedModel`, what the planner draws), with the model's derivation
 * of the flat fields the sheet's own display reads around the plan, and the
 * record's native extensions the model does not hold (a trendline or error
 * bars the shared contract refuses on this family, which the record keeps
 * beside its model). The live data and the host's resolutions (theme accents,
 * the value axis' linked number format, an imported chart's look) are the
 * caller's to add.
 */
export function spreadsheetChartModelInput(
  chart: SpreadsheetChartLike & { chartV2: ChartBlockAttrsV2 },
): SpreadsheetChartV2Input {
  const model = comboModelWithRecordErrorBars(chart, chart.chartV2);
  const kind = spreadsheetChartCollectionKind(chart);
  const binding = spreadsheetChartBindingOf(chart);
  const fields = legacySpreadsheetChartFields(model, binding, kind, chart);
  // The value axis' number format is the caller's to resolve (a Linked one
  // reads the source cells).
  const { type, chartType, range: _range, sourceSheetId: _source, seriesInColumns: _plotBy, nativeTitlePresent: _present,
    valueAxisNumberFormat: _numberFormat, desktopStyleId: _desktop, ...rest } = fields;
  const trendline = !chartSupportsTrendline(model.chartType, model.chartSubtype) ? normalizeTrendline(chart.trendline) : undefined;
  const errorBars = !chartDataTakesErrorBars(model.chartType, model.chartSubtype, model.data) ? normalizeErrorBars(chart.errorBars) : undefined;
  // A classic part imported from a workbook that states no text draws in
  // Excel's built-in chart text (18pt bold title, 10pt text in black), not
  // the inserted-chart style's 14pt grey.
  const builtInText =
    !(chart as { nativeLook?: unknown }).nativeLook && isImportedClassicChartPart(chart.source)
      ? excelBuiltInChartTextLook(
          model.format as Record<string, unknown> | undefined,
          spreadsheetChartBuiltInStyleLook(chart),
        )
      : null;
  return {
    ...(rest as Partial<SpreadsheetChartV2Input>),
    ...(trendline ? { trendline } : {}),
    ...(errorBars ? { errorBars } : {}),
    ...(builtInText ? { nativeLook: builtInText } : {}),
    type: typeof (type ?? chartType) === "string" ? String(type ?? chartType) : undefined,
    title: model.title,
    storedModel: model,
    ...((binding as { literal?: unknown } | undefined)?.literal === true ? { storedModelOwnData: true } : {}),
  };
}

type FormatRecord = Record<string, unknown>;

/**
 * The built-in chart style paint the server read at import
 * (`builtInStyleLook`), while the chart still wears that style: a gallery
 * pick (`chartStyle`) replaces it.
 */
function spreadsheetChartBuiltInStyleLook(chart: unknown): Record<string, unknown> | undefined {
  const record = chart as { builtInStyleLook?: unknown; chartStyle?: unknown };
  const look = record.builtInStyleLook;
  if (typeof look !== "object" || look === null || Array.isArray(look)) return undefined;
  const style = (look as { style?: unknown }).style;
  return record.chartStyle === undefined || record.chartStyle === style ? (look as Record<string, unknown>) : undefined;
}

/**
 * Format Axis > Number on the value axis: a format code, or null for Linked
 * to source (Excel's default, the axis reading its source cells' format).
 */
export function sheetChartWithValueAxisNumberFormat(
  attrs: ChartBlockAttrsV2,
  formatCode: string | null,
): ChartBlockAttrsV2 | null {
  const current = attrs.format?.axisOptions?.value;
  if ((current?.numberFormat ?? null) === formatCode) return null;
  const options = { ...(current ?? {}) } as FormatRecord;
  if (formatCode === null) delete options.numberFormat;
  else options.numberFormat = formatCode;
  const next = withOfficeChartAxisOptions(attrs, "value", options);
  if ((next.format?.axisOptions?.value?.numberFormat ?? null) !== formatCode) return null;
  const tidy = { ...next } as ChartBlockAttrsV2;
  if (tidy.format && !Object.keys(tidy.format).length) delete tidy.format;
  return validChartModel(tidy) ? tidy : null;
}

/**
 * The chart retargeted -- another kind (Change Chart Type), another layout of
 * its source (Switch Row/Column), another source (Select Data), a series on
 * the secondary axis: `reading` is the sheet's reading of the retargeted
 * record, its data re-read from the binding for the new kind. The model's own
 * title, axis titles and format ride through the shared type command wherever
 * the new data takes them; where it does not, the reading is the chart (with
 * the model's titles), as the record would read after the change.
 */
export function sheetChartRetargeted(
  model: ChartBlockAttrsV2,
  reading: ChartBlockAttrsV2,
  options: { switched?: boolean } = {},
): ChartBlockAttrsV2 | null {
  const kept = ((): ChartBlockAttrsV2 | null => {
    let base = model;
    if (options.switched) {
      const format = officeChartFormatWithRowsAndColumnsSwitched(base.format);
      base = { ...base, ...(format ? { format } : {}) };
    }
    const moved = spreadsheetChartModelOverData(base, reading.data);
    if (moved.chartType === reading.chartType && moved.chartSubtype === reading.chartSubtype) {
      return validateChartBlockAttrsV2(moved).valid ? moved : null;
    }
    return officeChartWithType(moved, reading.chartType, reading.chartSubtype);
  })();
  if (kept) return kept;
  const titled: ChartBlockAttrsV2 = { ...reading, title: model.title };
  for (const [key, axis] of [["axisTitleX", "x"], ["axisTitleY", "y"]] as const) {
    if (model[key] && chartSupportsAxisTitle(reading.chartType, axis)) titled[key] = model[key];
  }
  if (validChartModel(titled)) return titled;
  return validChartModel(reading) ? reading : null;
}

/**
 * Why a chart record's content is not the editor's to change -- a
 * preserve-only imported chart, whose part states what the shared model
 * cannot -- as its first reason code; null when every chart command may edit
 * it. The ONE gate every sheet chart content command and control asks (the
 * deck's `presentationChartV2CanEdit` counterpart): the import receipt
 * `source.nativeChartV2Authority` with `editable: false` and its
 * `preserveReasons`, or an older page's `chartV2Preserve {reasons}` marker.
 * The chart stays selectable, movable and resizable.
 */
export function spreadsheetChartEditRefusal(chart: SpreadsheetChartLike): string | null {
  const reasons = spreadsheetChartPreserveReasons(chart);
  return reasons ? (reasons[0] ?? "preserve-only") : null;
}

/** Every reason a preserve-only chart's receipt (or older marker) lists, in
 * order; null for a chart the editor may change. */
export function spreadsheetChartPreserveReasons(chart: SpreadsheetChartLike): string[] | null {
  const reasons = (value: unknown) =>
    Array.isArray(value) ? value.filter((reason): reason is string => typeof reason === "string" && reason.trim() !== "") : [];
  const source = chart.source && typeof chart.source === "object" ? (chart.source as Record<string, unknown>) : null;
  const receipt = source?.nativeChartV2Authority && typeof source.nativeChartV2Authority === "object"
    ? (source.nativeChartV2Authority as Record<string, unknown>)
    : null;
  if (receipt?.editable === false) return reasons(receipt.preserveReasons);
  const preserve = chart.chartV2Preserve;
  if (preserve && typeof preserve === "object" && !Array.isArray(preserve)) {
    return reasons((preserve as { reasons?: unknown }).reasons);
  }
  return null;
}

/**
 * A new sheet chart born with its model, the way `createChartBlockAttrsV2`
 * gives a new Docs or deck chart its own: the shared factory (Office's
 * default look -- Colorful Palette 1, the value axis' major gridlines) over
 * the data its binding reads, with the kind, titles and legend the insert
 * states. A chart whose data cannot be plotted stays as it is (it gains its
 * model at its first write).
 */
export function spreadsheetChartBornWithModel<T extends SpreadsheetChartLike>(
  chart: T,
  live: SpreadsheetChartV2Live | null,
  options: { sheetIdByName?: (name: string) => string | undefined } = {},
): T {
  if (!live) return chart;
  const reading = spreadsheetChartV2Reading(chart, live);
  if (!reading) return chart;
  let attrs: ChartBlockAttrsV2 = reading;
  try {
    const born = createChartBlockAttrsV2(reading.chartType, reading.chartSubtype, {
      title: reading.title,
      data: reading.data,
      ...(reading.axisTitleX ? { axisTitleX: reading.axisTitleX } : {}),
      ...(reading.axisTitleY ? { axisTitleY: reading.axisTitleY } : {}),
      ...(reading.format?.legend ? { format: { legend: reading.format.legend } } : {}),
    });
    if (validChartModel(born)) attrs = born;
  } catch {
    // A family the factory refuses keeps the reading.
  }
  return withSpreadsheetChartV2(chart, attrs, spreadsheetChartBindingOf(chart, options));
}

/**
 * A patch a writer that does not know the model made to a chart record's
 * flat fields (Copilot's `update_sheet_object` on an imported chart: its
 * `title`, its kind), stated on the record's model as well -- the save reads
 * only the model. A kind whose data is of another shape is left to the
 * model's own commands (the record's flat kind changes, its model does not).
 * A record without a model takes the patch as it always did.
 */
export function withSpreadsheetChartFlatPatch<T extends SpreadsheetChartLike>(
  chart: T,
  patch: Record<string, unknown>,
): T {
  const patched = { ...chart, ...patch } as T;
  const model = chart.chartV2 as ChartBlockAttrsV2 | undefined;
  if (!model || typeof model !== "object") return patched;
  let next: ChartBlockAttrsV2 = model;
  if (typeof patch.title === "string") next = { ...next, title: patch.title.slice(0, 200) };
  const kind = spreadsheetChartCollectionKind(chart) === "drawing" ? patch.chartType : patch.type;
  if (typeof kind === "string" && kind.trim()) {
    const leaf = spreadsheetChartLeafOf({
      type: normalizeSpreadsheetChartType(kind),
      chartSubtype: typeof patch.chartSubtype === "string" ? patch.chartSubtype : undefined,
    });
    if (leaf && (leaf.family !== next.chartType || leaf.subtype !== next.chartSubtype)) {
      const typed = officeChartWithType(next, leaf.family, leaf.subtype as ChartBlockAttrsV2["chartSubtype"]);
      if (typed) next = typed;
    }
  }
  if (next === model || !validateChartBlockAttrsV2(next).valid) return patched;
  return withSpreadsheetChartV2(patched, next);
}
