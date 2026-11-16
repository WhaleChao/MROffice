/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A worksheet chart's binding: the live range and the chart record read into
 * the shared chart model (`ChartBlockAttrsV2`), the accessible data table and
 * the sheet's chart-area appearance. Model knowledge, no drawing: it runs in
 * the browser (the editors' data-only questions, the canvas's frame and table)
 * and under the server's planner (`chart-v2-adapter.ts`, the chart engine),
 * which draws from exactly this reading.
 */
import type { ChartV2RenderPlan } from "@/components/console/editor/extensions/chartV2Echarts";
import { isComboRenderAs } from "../../components/console/editor/extensions/chartBlockShared";
import type {
  ChartPointValue,
  CategoryChartData,
  ChartBlockAttrsV2,
  ChartData,
  ChartFamily,
  ChartSubtype,
  ComboRenderAs,
  HierarchyChartData,
  HierarchyChartPoint,
  SampleChartData,
  StockChartData,
  WaterfallChartData,
  XYChartData,
} from "../../components/console/editor/extensions/chartBlockShared";
import { officeChartHostSupport } from "../office-charts/catalog";
import { DEFAULT_CHART_STYLE_PRESET_ID } from "../office-charts/chart-style-presets.generated";
import { MAX_CHART_POINTS, MAX_CHART_POINTS_PER_SERIES, MAX_CHART_SERIES } from "../office-charts/chart-limits.generated";
import { nativeLookPalette, nativeLookRgba, type ChartNativeLook } from "../office-chart-native-look";
import { excelBuiltInChartStyleColors, excelBuiltInStyleSeriesColors } from "../office-charts/chart-color-palettes";
import { spreadsheetChartRenderFormat, type SpreadsheetChartFormat } from "./chart-format";
import { CHART_SUBTYPES_BY_FAMILY, chartCategoryLevelsPicked, normalizeChartCategoryLevels, chartTakesColorPalette, chartTakesParentLabelLayout, chartDataLabelPositions, chartMapLabelLayouts, chartSupportsAxisVisibility, chartSupportsGridlines, chartSupportsMinorGridlines, chartSupportsSeriesFilter, chartSupportsCategoryFilter, chartSupportsDataLabels, chartSupportsDataLabelCallout, chartSupportsDataTable, chartSupportsErrorBars, chartSupportsTrendline, validateChartBlockAttrsV2, type ChartBlockFormat } from "../../components/console/editor/extensions/chartBlockShared";
import {
  normalizePresentationChartSeriesStyles,
  normalizePresentationChartSliceStyles,
  type PresentationChartSeriesStyle,
  type PresentationChartSliceStyle,
} from "../presentation-chart-area-style";

import {
  CHART_HIERARCHY_PATH_SEPARATOR,
  normalizeChartLineSeriesStyles,
  normalizeChartSeriesRefs,
  normalizeSpreadsheetChartType,
  spreadsheetChartStyleAppearance,
  spreadsheetChartColorPaletteColors,
  normalizeSpreadsheetChartColorPalette,
  spreadsheetTrendlineContract,
  type MiniChartXYSeries,
  type SpreadsheetChartLineSeriesStyle,
  type SpreadsheetChartSeriesGradientStyle,
  type SpreadsheetChartType,
  type ChartTrendlineSpec,
  type ChartErrorBarSpec,
} from "./chart";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

// One chart's size is the one limits table every editor shares
// (lib/office-charts/chart-limits.generated): Office's 255 series, 4,000
// points per series, 50,000 values per chart. A worksheet chart plots the
// first rows and columns of its range within them, as it always kept the
// first 4,000 categories (never a silent 50).
export const MAX_RENDER_CATEGORIES = MAX_CHART_POINTS_PER_SERIES;
const MAX_RENDER_SERIES = MAX_CHART_SERIES;
/** The categories a chart of `seriesCount` series keeps within the per-chart bound. */
function renderCategoryBound(seriesCount: number): number {
  return Math.max(1, Math.min(MAX_RENDER_CATEGORIES, Math.floor(MAX_CHART_POINTS / Math.max(1, seriesCount))));
}
// Native writer's Natural Earth 1:50m ISO index, locked by the adapter test.
// Geometry stays in the existing lazy chartV2Map/world-atlas frontend loader;
// never import API files across the apps/web Turbopack root.
const MAP_COUNTRY_CODES = new Set(
  "004 008 010 012 016 020 024 028 031 032 036 040 044 048 050 051 052 056 060 064 068 070 072 076 084 086 090 092 096 100 104 108 112 116 120 124 132 136 140 144 148 152 156 158 170 174 178 180 184 188 191 192 196 203 204 208 212 214 218 222 226 231 232 233 234 238 239 242 246 248 250 258 260 262 266 268 270 275 276 288 296 300 304 308 316 320 324 328 332 334 336 340 344 348 352 356 360 364 368 372 376 380 384 388 392 398 400 404 408 410 414 417 418 422 426 428 430 434 438 440 442 446 450 454 458 462 466 470 478 480 484 492 496 498 499 500 504 508 512 516 520 524 528 531 533 534 540 548 554 558 562 566 570 574 578 580 583 584 585 586 591 598 600 604 608 612 616 620 624 626 630 634 642 643 646 652 654 659 660 662 663 666 670 674 678 682 686 688 690 694 702 703 704 705 706 710 716 724 728 729 732 740 748 752 756 760 762 764 768 776 780 784 788 792 795 796 800 804 807 818 826 831 832 833 834 840 850 854 858 860 862 876 882 887 894".split(" "),
);

/** One label an imported filled map's own geocoding cache bound to a country. */
export interface SpreadsheetChartMapRegionCode {
  label: string;
  /** ISO 3166-1 numeric code, the key the world map draws each country by. */
  id: string;
}

/** Keeps a well-formed reading; the server reads it from the part's cache. */
export function normalizeSpreadsheetChartMapRegionCodes(
  raw: unknown,
): SpreadsheetChartMapRegionCode[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > 100) return undefined;
  const codes: SpreadsheetChartMapRegionCode[] = [];
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") return undefined;
    const { label, id } = entry as Record<string, unknown>;
    if (typeof label !== "string" || !label || label.length > 200) return undefined;
    if (typeof id !== "string" || !/^\d{3}$/.test(id)) return undefined;
    codes.push({ label, id });
  }
  return codes;
}

type EChartsRecord = Record<string, unknown>;

interface SpreadsheetChartDefinition {
  family: ChartFamily;
  subtype: ChartSubtype;
}

const SPREADSHEET_CHART_V2_DEFINITIONS: Record<
  Exclude<SpreadsheetChartType, "camera">,
  SpreadsheetChartDefinition
> = {
  column: { family: "column", subtype: "clustered" },
  bar: { family: "bar", subtype: "clustered" },
  line: { family: "line", subtype: "standardMarkers" },
  pie: { family: "pie", subtype: "pie2d" },
  area: { family: "area", subtype: "standard" },
  scatter: { family: "xy", subtype: "markers" },
  "stacked-bar": { family: "bar", subtype: "stacked" },
  "stacked-column": { family: "column", subtype: "stacked" },
  combo: { family: "combo", subtype: "columnLine" },
  doughnut: { family: "pie", subtype: "doughnut" },
  radar: { family: "radar", subtype: "standard" },
  bubble: { family: "xy", subtype: "bubble" },
  histogram: { family: "statistical", subtype: "histogram" },
  waterfall: { family: "waterfall", subtype: "waterfall" },
  "box-and-whisker": { family: "statistical", subtype: "boxWhisker" },
  pareto: { family: "statistical", subtype: "pareto" },
  treemap: { family: "hierarchy", subtype: "treemap" },
  sunburst: { family: "hierarchy", subtype: "sunburst" },
  stock: { family: "stock", subtype: "openHighLowClose" },
  funnel: { family: "funnel", subtype: "funnel" },
  map: { family: "map", subtype: "filledMap" },
  surface: { family: "surface", subtype: "surface3d" },
};

/**
 * What inserting one catalog leaf means to the worksheet: the flat native kind
 * plus, only when the leaf is not that kind's own default, the shared catalog
 * subtype. See {@link spreadsheetChartTypeForLeaf}.
 */
export interface SpreadsheetChartLeafInsert {
  type: SpreadsheetChartType;
  /** Omitted when the leaf IS the kind's default (byte-stable persistence). */
  chartSubtype?: string;
  /**
   * The leaf is Office's "clustered column + line on secondary axis" combo,
   * which the XLSX writer expresses as the plain combo plus a secondary-axis
   * series index. Only the host knows how many value columns the selection
   * has, so it resolves the index (the LAST value series) at insert time.
   */
  secondaryAxisLastSeries?: true;
}

/**
 * Which flat `SpreadsheetChartType` (and subtype) one shared-catalog leaf
 * inserts as — the inverse of {@link SPREADSHEET_CHART_V2_DEFINITIONS}, and the
 * ONE place that inversion lives (the ribbon galleries and Chart Design ▸
 * Change Chart Type both read it, so a menu and a retype cannot disagree).
 *
 * Rules, all of them forced by what the XLSX writer round-trips:
 * - `type` is the family's base kind. The two stacked kinds
 *   (`stacked-column` / `stacked-bar`) are NEVER produced here: they are not in
 *   `NATIVE_XLSX_CHART_AUTHORING_TYPES`, so a stacked leaf ships as the base
 *   kind carrying the desktop subtype, which is the route the writer already
 *   round-trips.
 * - `chartSubtype` is written only when the leaf differs from the kind's
 *   default subtype, because the persistence convention is to omit an optional
 *   field whose value is the default.
 * - `combo`: Stacked Area - Clustered Column is its `chartSubtype`, which the
 *   XLSX writer states, and so is Custom Combination (with the per-series
 *   types the insert asks for); `columnLineSecondary` becomes a secondary-axis
 *   series index.
 * - `null` means this host cannot author the leaf; callers must not fall back
 *   to a neighbouring chart.
 */
export function spreadsheetChartTypeForLeaf(
  family: ChartFamily,
  subtype: ChartSubtype,
): SpreadsheetChartLeafInsert | null {
  if (!officeChartHostSupport("spreadsheet", family, subtype)) return null;
  if (family === "combo") {
    return subtype === "columnLineSecondary"
      ? { type: "combo", secondaryAxisLastSeries: true }
      : subtype === "stackedAreaColumn" || subtype === "custom"
        ? { type: "combo", chartSubtype: subtype }
        : { type: "combo" };
  }
  const type = SPREADSHEET_CHART_LEAF_KINDS[family]?.(subtype);
  if (!type) return null;
  const definition = SPREADSHEET_CHART_V2_DEFINITIONS[type];
  if (definition.family !== family) return null;
  return definition.subtype === subtype
    ? { type }
    : { type, chartSubtype: subtype };
}

/** The shared family a worksheet chart kind is drawn and written as. */
export function spreadsheetChartFamily(type: unknown): ChartFamily | undefined {
  const kind = normalizeSpreadsheetChartType(type);
  return kind === "camera" ? undefined : SPREADSHEET_CHART_V2_DEFINITIONS[kind]?.family;
}

/**
 * The worksheet kind a shared (family, subtype) is stored as, host support
 * aside: the inverse of {@link SPREADSHEET_CHART_V2_DEFINITIONS}, for a chart
 * that already IS a V2 record (an import may carry a subtype the sheet cannot
 * author; it is still that family's kind). Combo keeps its shared subtype.
 */
export function spreadsheetChartKindForV2(
  family: ChartFamily,
  subtype: ChartSubtype,
): SpreadsheetChartLeafInsert | null {
  if (family === "combo") {
    return subtype === "columnLine" ? { type: "combo" } : { type: "combo", chartSubtype: subtype };
  }
  const type = SPREADSHEET_CHART_LEAF_KINDS[family]?.(subtype);
  if (!type) return null;
  const definition = SPREADSHEET_CHART_V2_DEFINITIONS[type];
  return definition.subtype === subtype ? { type } : { type, chartSubtype: subtype };
}

/**
 * Whether the shared chart contract offers gridlines on this sheet chart kind:
 * `chartSupportsGridlines` on the family the kind draws as. A funnel has axes
 * and no gridlines — Excel's Chart Elements menu lists none for it.
 */
export function spreadsheetChartTypeSupportsGridlines(
  type: SpreadsheetChartType,
): boolean {
  if (type === "camera") return false;
  const definition = SPREADSHEET_CHART_V2_DEFINITIONS[type];
  return definition !== undefined && chartSupportsGridlines(definition.family);
}

/**
 * Whether Office draws a trendline, or error bars, on this sheet chart — the
 * shared contract's rule on the family and subtype the kind draws as.
 *
 * The ribbon used to carry two lists of its own (a trendline on `line`,
 * `column`, `bar` and their stacked kinds; error bars on `line` and `column`),
 * which offered a trendline on a stacked chart Excel gives none, and greyed
 * both out on an area, a scatter, a bubble and a stock chart Excel offers them
 * on. A chart's own `chartSubtype` wins over the kind's default, so a stacked
 * or a 3-D line is not mistaken for a plain one.
 */
function spreadsheetChartLeaf(chart: {
  type: SpreadsheetChartType;
  chartSubtype?: string;
}): { family: ChartFamily; subtype: string } | null {
  if (chart.type === "camera") return null;
  const definition = SPREADSHEET_CHART_V2_DEFINITIONS[chart.type];
  if (!definition) return null;
  return {
    family: definition.family,
    subtype: chart.chartSubtype || definition.subtype,
  };
}

/** The shared (family, subtype) a flat sheet kind (and its subtype) draws as. */
export function spreadsheetChartLeafOf(chart: {
  type: SpreadsheetChartType;
  chartSubtype?: string;
}): { family: ChartFamily; subtype: string } | null {
  return spreadsheetChartLeaf(chart);
}

export function spreadsheetChartSupportsTrendline(chart: {
  type: SpreadsheetChartType;
  chartSubtype?: string;
}): boolean {
  const leaf = spreadsheetChartLeaf(chart);
  return leaf !== null && chartSupportsTrendline(leaf.family, leaf.subtype);
}

/** Whether the chart states primary minor gridlines as format keys (a
 * classic chart with gridlines that takes no calibrated Quick Layout). */
export function spreadsheetChartSupportsMinorGridlines(chart: {
  type: SpreadsheetChartType;
  chartSubtype?: string;
}): boolean {
  const leaf = spreadsheetChartLeaf(chart);
  return leaf !== null && chartSupportsMinorGridlines(leaf.family, leaf.subtype);
}

export function spreadsheetChartSupportsErrorBars(chart: {
  type: SpreadsheetChartType;
  chartSubtype?: string;
}): boolean {
  const leaf = spreadsheetChartLeaf(chart);
  return leaf !== null && chartSupportsErrorBars(leaf.family, leaf.subtype);
}

/**
 * Which native kind OWNS each leaf of a family. Only the families whose leaves
 * split across several flat kinds need a rule; the rest are the family's single
 * kind. `bubble3d` belongs to `bubble` (not `scatter`), which is why this is a
 * table and not a lookup of the first matching definition.
 */
const SPREADSHEET_CHART_LEAF_KINDS: Partial<
  Record<
    ChartFamily,
    (subtype: ChartSubtype) => Exclude<SpreadsheetChartType, "camera"> | null
  >
> = {
  column: () => "column",
  bar: () => "bar",
  line: () => "line",
  area: () => "area",
  radar: () => "radar",
  stock: () => "stock",
  waterfall: () => "waterfall",
  funnel: () => "funnel",
  map: () => "map",
  surface: () => "surface",
  pie: (subtype) => (subtype === "doughnut" ? "doughnut" : "pie"),
  xy: (subtype) => (subtype.startsWith("bubble") ? "bubble" : "scatter"),
  hierarchy: (subtype) => (subtype === "sunburst" ? "sunburst" : "treemap"),
  statistical: (subtype) =>
    subtype === "pareto"
      ? "pareto"
      : subtype === "boxWhisker"
        ? "box-and-whisker"
        : "histogram",
};

export interface SpreadsheetChartV2Input {
  type?: string;
  /**
   * The picture draws the chart's title and axis titles itself (a chart on
   * the model, in an editor that types them in place there, as Excel's chart
   * area holds them); else the sheet draws them around the plot.
   */
  drawTitles?: boolean;
  rows?: string[][];
  /** Whether rows[0] is the series-title row. Set by a host that knows (rows
   *  built from series references always are); otherwise sniffed from the
   *  first row's values. */
  headerRow?: boolean;
  /**
   * A multi-level category axis over `rows` (a PivotChart with two or more
   * row fields): the levels above the leaf labels in each row's first cell,
   * outermost first, one entry per DATA row (the rows after a header row),
   * a group's label on its first row and "" on the rows after it --
   * `CategoryChartData.categoryLevels` aligned to the rows. Only the empty
   * rows past the range's last entry are dropped; levels that do not align
   * with the rows are not drawn.
   */
  categoryLevels?: readonly (readonly string[])[];
  xySeries?: MiniChartXYSeries[];
  title?: string;
  chartStyle?: number;
  chartColorPalette?: string;
  themeAccents?: readonly string[];
  /** The face the chart's text is drawn in when the chart states none: the
   * workbook theme's body face, as Excel draws chart text (`+mn-lt`). */
  chartFontFamily?: string;
  chartFormat?: SpreadsheetChartFormat;
  chartSubtype?: string;
  seriesRefs?: unknown;
  sourcePointIndices?: readonly number[];
  seriesStyles?: readonly PresentationChartSeriesStyle[];
  seriesGradientStyles?: readonly SpreadsheetChartSeriesGradientStyle[];
  lineSeriesStyles?: readonly SpreadsheetChartLineSeriesStyle[];
  sliceStyles?: readonly PresentationChartSliceStyle[];
  showLegend?: boolean;
  legendPosition?: "top" | "bottom" | "left" | "right" | "topRight";
  showAxisLabels?: boolean;
  showGridlines?: boolean;
  showDataLabels?: boolean;
  secondaryAxisSeries?: number[];
  /** A range-bound Custom Combination's per-series chart types. */
  comboSeriesTypes?: readonly ComboRenderAs[];
  valueAxisNumberFormat?: string;
  /** Each series' source format, by series index (its value cells' one
   *  number format, `resolveLinkedSeriesNumberFormats`): a data label Linked
   *  to source shows its series' (Excel: "17,422", not 17422). */
  seriesNumberFormats?: readonly (string | undefined)[];
  /** The workbook reads serials in the 1904 date system: a value axis or data
   *  label in a date format reads its numbers as 1904 dates. */
  date1904?: ExcelDateSystemArg;
  trendline?: ChartTrendlineSpec;
  errorBars?: ChartErrorBarSpec;
  upDownBars?: boolean;
  dropLines?: boolean;
  highLowLines?: boolean;
  seriesLines?: boolean;
  /** The look an imported Office 2016 chart's style parts give it; a gallery pick replaces it. */
  nativeLook?: ChartNativeLook;
  /** The countries an imported filled map's cells name, as its cache resolved them. */
  mapRegionCodes?: readonly SpreadsheetChartMapRegionCode[];
  /** The reader's language, for names Office leaves to it. */
  locale?: string;
  /** The catalogue word for a Pie of Pie's combined slice ("Other"). */
  otherSliceLabel?: string;
  /** The number format the category cells are written in (Linked to
   *  source): a date axis that states none of its own writes its labels in it. */
  categoryNumberFormat?: string;
  /** The catalogue words of Office's automatic title ("Chart Title"), for a
   *  model whose title is automatic (`format.automaticTitle`). */
  chartTitleWord?: string;
  /** The catalogue words of Office's automatic axis title ("Axis Title",
   *  坐标轴标题), for a model whose axis title is the reading's automatic one. */
  axisTitleWord?: string;
  /** The catalogue's name for a series that names nothing ("Series{index}",
   *  "系列{index}"), for a chart with such a series. */
  defaultSeriesName?: string;
  /**
   * The chart's stored model (the record's `chartV2`, lib/spreadsheet/
   * chart-v2-model). Given, it IS what the shared planner draws -- handed to
   * it as Docs and the deck hand theirs (`documentChartRenderPlan`), and as
   * the worksheet PDF export draws the same chart -- its data read live from
   * `rows` / `xySeries` in the model's own shape. The flat fields above are
   * then the model's derivation (`spreadsheetChartModelInput`): they read the
   * live data and steer the sheet's own display around the plan (legend and
   * plot margins, axis ink, markers), and restate nothing of the chart. A
   * model the live data cannot refresh is not drawn; the flat reading is.
   */
  storedModel?: ChartBlockAttrsV2;
  /** The stored model's data is its own (`binding.literal`): drawn as stored, never re-read. */
  storedModelOwnData?: boolean;
}

export type SpreadsheetChartAccessibleHeaderRole = "category" | "series" | "size" | "literal";

export interface SpreadsheetChartAccessibleTable {
  headers: string[];
  /** Display provenance only; literal headers include all user names and X/Y. */
  headerRoles: SpreadsheetChartAccessibleHeaderRole[];
  rows: string[][];
  /** Each row's index among the rows read from the range, when rows the
   * chart does not draw were left out (Chart Filters' hidden categories). */
  rowSources?: number[];
}

/**
 * The table and counts of what the chart draws: a series or category Chart
 * Filters hide is left out (live audit S-054: the summary said "2 series, 4
 * points" and the data table listed the hidden series and quarter of a chart
 * drawing one series over three), and a pie, which draws its first series
 * only, lists that one (a doughnut draws them all, as rings).
 */
function drawnChartTable(
  table: SpreadsheetChartAccessibleTable,
  data: ChartData,
  family: string,
  hiddenSeries: readonly number[] | undefined,
  hiddenCategories: readonly number[] | undefined,
  firstSeriesOnly: boolean,
): { table: SpreadsheetChartAccessibleTable; series: number; categories: number } | null {
  const seriesHidden = new Set(chartSupportsSeriesFilter(family) ? hiddenSeries ?? [] : []);
  const categoriesHidden = new Set(chartSupportsCategoryFilter(family) ? hiddenCategories ?? [] : []);
  if (table.headerRoles[0] === "category") {
    const seriesTotal = table.headers.length - 1;
    const keepSeries = (index: number) =>
      !seriesHidden.has(index) && (!firstSeriesOnly || index === 0);
    const columns = [0, ...Array.from({ length: seriesTotal }, (_, index) => index).filter(keepSeries).map((index) => index + 1)];
    const rowSources = table.rows.map((_, index) => index).filter((index) => !categoriesHidden.has(index));
    if (columns.length === table.headers.length && rowSources.length === table.rows.length) return null;
    return {
      table: {
        headers: columns.map((column) => table.headers[column]!),
        headerRoles: columns.map((column) => table.headerRoles[column]!),
        rows: rowSources.map((row) => columns.map((column) => table.rows[row]![column] ?? "")),
        ...(rowSources.length !== table.rows.length ? { rowSources } : {}),
      },
      series: columns.length - 1,
      categories: rowSources.length,
    };
  }
  // One row per point, series after series (an XY chart's own series).
  if (table.headerRoles[0] === "series" && data.kind === "xy" && seriesHidden.size > 0) {
    const rowSources: number[] = [];
    let row = 0;
    data.series.forEach((series, index) => {
      for (let point = 0; point < series.xValues.length; point += 1, row += 1) {
        if (!seriesHidden.has(index)) rowSources.push(row);
      }
    });
    const kept = data.series.filter((_, index) => !seriesHidden.has(index));
    return {
      table: { ...table, rows: rowSources.map((index) => table.rows[index]!), rowSources },
      series: kept.length,
      categories: Math.max(0, ...kept.map((series) => series.xValues.length)),
    };
  }
  return null;
}

export interface SpreadsheetChartAppearance {
  background: string;
  plotBackground: string;
  text: string;
  /** Data label and legend text an imported chart's own look colours apart from its axes. */
  labelText?: string;
  legendText?: string;
  /** The axis line an imported chart's own look draws, apart from its border. */
  axisLine?: string;
  gridline: string;
  border: string;
  palette: readonly string[];
}

export interface SpreadsheetChartV2Plan extends ChartV2RenderPlan {
  /** What was drawn: the stored model over its live data, or the flat reading. */
  attrs: ChartBlockAttrsV2;
  /** The planner drew the stored model (`storedModel`), not the flat reading. */
  drawnFromModel?: true;
  accessibleTable: SpreadsheetChartAccessibleTable;
  appearance: SpreadsheetChartAppearance;
  summary: string;
  categoryCount: number;
  seriesCount: number;
}

export interface ParsedCategoryTable {
  data: CategoryChartData;
  table: SpreadsheetChartAccessibleTable;
}

export function finiteNumber(value: unknown): number | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" && value.trim() === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

export function boundedText(value: unknown, fallback: string): string {
  const text = String(value ?? "").trim();
  return (text || fallback).slice(0, 200);
}

/** A worksheet error value: a chart leaves its point out, as it does a blank. */
const CELL_ERROR_VALUE =
  /^#(?:N\/A|NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)$/i;







function chartCellNumber(value: unknown, textAsZero: boolean, errorsAsZero = false): number | null {
  const number = finiteNumber(value);
  if (number !== null || !textAsZero || typeof value !== "string") return number;
  const text = value.trim();
  if (!text) return null;
  if (!CELL_ERROR_VALUE.test(text)) return 0;
  // A classic chart plots an error value as 0 too, all but #N/A, which is
  // no point (Show #N/A as an empty cell decides how it is drawn).
  return errorsAsZero && text.toUpperCase() !== "#N/A" ? 0 : null;
}

/**
 * A value cell that holds text with no characters -- a formula's "" result
 * (`=IF(x, y, "")`) -- spelled for the chart's value matrix. A matrix cell is
 * its text, so "" would read as an empty cell, which Show empty cells as
 * draws; Excel plots it as text, 0 (Excel for Mac, rtfix_x210: ="" -> 0 in
 * the chart's own value cache, the line drops to the axis; X-210). The
 * mark is a zero-width space: text, never shown.
 */
export const CHART_EMPTY_TEXT_POINT = "\u200b";

/** ``value`` (a value cell as the chart reads it) with a formula's "" result
 * marked as text (`CHART_EMPTY_TEXT_POINT`); any other value as it is. */
export function chartPointCellValue(value: string, cell: { formula?: unknown } | undefined): string {
  return value === "" && typeof cell?.formula === "string" && cell.formula !== "" ? CHART_EMPTY_TEXT_POINT : value;
}

function parseCategoryTable(
  rows: string[][] | undefined,
  headerRow?: boolean,
  categoryLevels?: readonly (readonly string[])[],
  textAsZero = false,
  /** A date axis: a blank category cell is a point with no date, not an
   * ordinal (X-130). */
  blankCategoriesStayBlank = false,
  errorsAsZero = false,
): ParsedCategoryTable | null {
  if (!Array.isArray(rows) || rows.length === 0) return null;
  const boundedRows = rows
    .slice(0, MAX_RENDER_CATEGORIES + 1)
    .filter((row): row is string[] => Array.isArray(row));
  if (boundedRows.length === 0) return null;

  const maxColumns = Math.max(
    1,
    ...boundedRows.map((row) => Math.min(row.length, MAX_RENDER_SERIES + 1)),
  );
  const valueColumnCount = Math.max(1, maxColumns - 1);
  const firstValueCells =
    maxColumns === 1
      ? [boundedRows[0]?.[0]]
      : boundedRows[0].slice(1, valueColumnCount + 1);
  const hasHeader =
    boundedRows.length > 1 &&
    (headerRow ?? firstValueCells.some((value) => finiteNumber(value) === null));
  const headers = Array.from({ length: valueColumnCount }, (_, index) =>
    hasHeader
      ? boundedText(
          maxColumns === 1 ? boundedRows[0]?.[0] : boundedRows[0]?.[index + 1],
          `Series ${index + 1}`,
        )
      : `Series ${index + 1}`,
  );
  const categoryBound = renderCategoryBound(valueColumnCount);
  const categories: string[] = [];
  const values = Array.from({ length: valueColumnCount }, () => [] as number[]);
  // XLS-A072: the points whose cell holds #N/A (Show #N/A as an empty cell).
  const naPoints = Array.from({ length: valueColumnCount }, () => [] as number[]);
  const tableRows: string[][] = [];
  // The data row each category was read from, for its levels.
  const sourceRows: number[] = [];
  const dataRowCount = rows.length - (hasHeader ? 1 : 0);
  const rowLevels = normalizeChartCategoryLevels(categoryLevels, dataRowCount);
  const dataRows = boundedRows.slice(hasHeader ? 1 : 0);
  // Every row of the range is a category, as Excel plots it: a month whose
  // value is blank keeps its place and is drawn by Show empty cells as (a
  // gap, zero, or a line across it). Only the empty rows the range runs on
  // past its last entry are not plotted (user audit S-015 lost February).
  let lastFilledRow = -1;
  dataRows.forEach((row, index) => {
    if (row.some((cell) => String(cell ?? "").trim() !== "")) lastFilledRow = index;
  });

  for (const [sourceIndex, row] of dataRows.entries()) {
    if (sourceIndex > lastFilledRow) break;
    const rawValues =
      maxColumns === 1
        ? [row[0]]
        : Array.from(
            { length: valueColumnCount },
            (_, index) => row[index + 1],
          );
    const numericValues = rawValues.map((value) => chartCellNumber(value, textAsZero, errorsAsZero));
    // A category with a missing value (#N/A, a blank) keeps its place on the
    // axis: only that series has a gap there (NaN, which ECharts draws as an
    // empty point), the other series keep their values -- every series blank
    // included.
    // A row whose outer level names it (a PivotChart's collapsed item) keeps
    // an empty innermost label; only an unnamed row reads as its ordinal.
    const namedByLevel = rowLevels?.some((level) => (level[sourceIndex] ?? "") !== "") === true;
    const category =
      maxColumns === 1
        ? String(sourceIndex + 1)
        : (namedByLevel || blankCategoriesStayBlank) && String(row[0] ?? "").trim() === ""
          ? ""
          : boundedText(row[0], String(sourceIndex + 1));
    categories.push(category);
    sourceRows.push(sourceIndex);
    const tableValues: string[] = [];
    numericValues.forEach((value, seriesIndex) => {
      if (value === null && String(rawValues[seriesIndex] ?? "").trim().toUpperCase() === "#N/A") {
        naPoints[seriesIndex].push(values[seriesIndex].length);
      }
      values[seriesIndex].push(value ?? Number.NaN);
      tableValues.push(value === null ? "" : String(value));
    });
    tableRows.push([category, ...tableValues]);
    if (categories.length >= categoryBound) break;
  }

  if (categories.length === 0 || values.every((series) => series.every(Number.isNaN))) return null;
  const levels = rowLevels;
  const categoryLevelsPlotted = levels
    ? chartCategoryLevelsPicked(
        { categories: Array.from({ length: dataRowCount }, () => ""), categoryLevels: levels },
        sourceRows,
      )
    : undefined;
  return {
    data: {
      kind: "category",
      categories,
      series: values.map((seriesValues, index) => ({
        name: headers[index],
        values: seriesValues,
        ...(naPoints[index].length ? { naPoints: naPoints[index] } : {}),
      })),
      ...(categoryLevelsPlotted ? { categoryLevels: categoryLevelsPlotted } : {}),
    },
    table: {
      headers: ["Category", ...headers],
      headerRoles: ["category", ...headers.map(() => "literal" as const)],
      rows: tableRows,
    },
  };
}

function xyDataFromNativeSeries(
  xySeries: MiniChartXYSeries[] | undefined,
  bubble: boolean,
): { data: XYChartData; table: SpreadsheetChartAccessibleTable } | null {
  if (!Array.isArray(xySeries) || xySeries.length === 0) return null;
  const pointBound = renderCategoryBound(Math.min(xySeries.length, MAX_RENDER_SERIES));
  const normalized = xySeries
    .slice(0, MAX_RENDER_SERIES)
    .map((series, seriesIndex) => {
      const points = series.points
        .slice(0, pointBound)
        .filter(
          (point) =>
            Number.isFinite(point.x) &&
            Number.isFinite(point.y) &&
            (!bubble || Number.isFinite(point.size)),
        );
      if (points.length === 0) {
        // A series with no point (its cells deleted, `Sheet!#REF!`) is still
        // the chart's series, as Excel keeps it: one blank point holds its
        // place, so the series after it keep their indices -- the styles,
        // hidden series and labels the model keys by index (R326-X06 E7: a
        // column delete dropped series 3 of 6 and its styles fell on the
        // wrong series, the export declined).
        return {
          name: boundedText(series.name, `Series ${seriesIndex + 1}`),
          xValues: [null],
          yValues: [null],
          ...(bubble ? { bubbleSizes: [0] } : {}),
        };
      }
      return {
        name: boundedText(series.name, `Series ${seriesIndex + 1}`),
        xValues: points.map((point) => point.x),
        yValues: points.map((point) => point.y),
        ...(bubble
          ? { bubbleSizes: points.map((point) => point.size as number) }
          : {}),
      };
    });
  if (normalized.every((series) => series.xValues[0] === null)) return null;
  return {
    data: { kind: "xy", series: normalized },
    table: {
      headers: bubble ? ["Series", "X", "Y", "Size"] : ["Series", "X", "Y"],
      headerRoles: bubble ? ["series", "literal", "literal", "size"] : ["series", "literal", "literal"],
      rows: normalized.flatMap((series) =>
        series.xValues[0] === null ? [] : series.xValues.map((x, index) => [
          series.name,
          String(x),
          String(series.yValues[index]),
          ...(bubble ? [String(series.bubbleSizes?.[index] ?? "")] : []),
        ]),
      ),
    },
  };
}

/** The XY data of a chart of its own data (`binding.literal`, the stored
 * model's), and its table; null for any other chart. */
function ownXYData(
  input: SpreadsheetChartV2Input,
): { data: XYChartData; table: SpreadsheetChartAccessibleTable } | null {
  const model = input.storedModel;
  if (input.storedModelOwnData !== true || !model || model.data.kind !== "xy") return null;
  const data = model.data as XYChartData;
  const bubble = data.series.some((series) => Array.isArray(series.bubbleSizes));
  const text = (value: number | null | undefined) => (typeof value === "number" && Number.isFinite(value) ? String(value) : "");
  return {
    data,
    table: {
      headers: bubble ? ["Series", "X", "Y", "Size"] : ["Series", "X", "Y"],
      headerRoles: bubble ? ["series", "literal", "literal", "size"] : ["series", "literal", "literal"],
      rows: data.series.flatMap((series) =>
        series.xValues.map((x, index) => [
          series.name,
          text(x),
          text(series.yValues[index]),
          ...(bubble ? [text(series.bubbleSizes?.[index])] : []),
        ]),
      ),
    },
  };
}

function xyDataFromCategoryTable(
  parsed: ParsedCategoryTable,
  bubble: boolean,
): XYChartData | null {
  const numericCategories = parsed.data.categories.map(finiteNumber);
  const xValues = numericCategories.every((value) => value !== null)
    ? (numericCategories as number[])
    : parsed.data.categories.map((_, index) => index + 1);
  // Rows of a range (buildChartSeries) carry the X vector twice, as the
  // point label and as the first value column, so a two-column X/Y
  // selection is executable. That copy is not a series: Excel plots
  // Month/A/B or X/A/B as two series, not three, and a bubble's Y and
  // size are the columns after it.
  const allSeries = parsed.data.series;
  const xCopy =
    allSeries.length > 1 &&
    allSeries[0].values.length === xValues.length &&
    allSeries[0].values.every((value, index) => value === xValues[index]);
  // A series that names nothing is named by its place among the Y series
  // (Series 1, which the picture says as Excel does: 系列1), not by its
  
  // numbers with no heading plots 系列1).
  const series = xCopy
    ? allSeries.slice(1).map((entry, index) =>
        isReadingDefaultSeriesName(entry.name, index + 1) ? { ...entry, name: `Series ${index + 1}` } : entry,
      )
    : allSeries;
  if (bubble) {
    if (series.length < 2) return null;
    return {
      kind: "xy",
      series: [
        {
          name: series[0].name,
          xValues,
          yValues: series[0].values,
          bubbleSizes: series[1].values.map((value) => value ?? Number.NaN),
        },
      ],
    };
  }
  return {
    kind: "xy",
    series: series.map((series) => ({
      name: series.name,
      xValues,
      yValues: series.values,
    })),
  };
}

function stockData(parsed: ParsedCategoryTable, requested?: string): {
  data: StockChartData;
  subtype: ChartSubtype;
} | null {
  const series = parsed.data.series;
  const recipes={highLowClose:["high","low","close"],openHighLowClose:["open","high","low","close"],volumeHighLowClose:["volume","high","low","close"],volumeOpenHighLowClose:["volume","open","high","low","close"]} as const;
  const selected=requested??(series.length===5?"volumeOpenHighLowClose":series.length===4?"openHighLowClose":"highLowClose");
  const selectedRoles=recipes[selected as keyof typeof recipes];
  if(!selectedRoles||series.length!==selectedRoles.length)return null;
  // Multi-level labels over the dates (H28_32readwriteStockChart2: Qtr 1..4
  // under 2008..2010) stay the stock chart's, as a column chart's do.
  const levels=parsed.data.categoryLevels?.length?{categoryLevels:parsed.data.categoryLevels.map(level=>[...level])}:{};
  return {subtype:selected as ChartSubtype,data:{kind:"stock",categories:parsed.data.categories,...levels,series:selectedRoles.map((role,index)=>({name:series[index].name,role,values:series[index].values}))}};
}



/** The reading the planner draws a worksheet chart from (`buildSpreadsheetChartV2Binding`). */
export interface SpreadsheetChartV2Binding {
  /** The flat reading: the chart as the sheet's flat fields state it. */
  attrs: ChartBlockAttrsV2;
  /** The stored model over the live data, when the record carries a valid one. */
  stored: ChartBlockAttrsV2 | null;
  nativeLook?: ChartNativeLook;
  parsed: ParsedCategoryTable;
  table: SpreadsheetChartAccessibleTable;
  appearance: SpreadsheetChartAppearance;
  summary: string;
  categoryCount: number;
  seriesCount: number;
}

/** What a worksheet chart draws as: `SpreadsheetChartV2Plan` without its drawing. */
export interface SpreadsheetChartV2Drawn {
  attrs: ChartBlockAttrsV2;
  drawnFromModel?: true;
  accessibleTable: SpreadsheetChartAccessibleTable;
  appearance: SpreadsheetChartAppearance;
  summary: string;
  categoryCount: number;
  seriesCount: number;
}

/** A series' plotted values: its blank points (`null`, a sheet's NaN) left out. */
export const plottedValues=(values:readonly ChartPointValue[]):number[]=>values.filter((value):value is number=>value!==null&&Number.isFinite(value));


export function comboData(
  parsed: ParsedCategoryTable,
  input: SpreadsheetChartV2Input,
  normalizedType: SpreadsheetChartType,
): CategoryChartData {
  const secondary = new Set(
    (input.secondaryAxisSeries ?? []).filter(
      (index) => Number.isInteger(index) && index >= 0,
    ),
  );
  const refs=normalizeChartSeriesRefs(input.seriesRefs);
  const subtype=input.chartSubtype;
  const render=(index:number):ComboRenderAs=>{
    if(subtype==="custom"&&refs){const value=refs[index]?.renderType;return value!=="bubble"&&isComboRenderAs(value)?value:"column";}
    if(subtype==="custom"&&input.comboSeriesTypes)return input.comboSeriesTypes[index]??"column";
    if(subtype==="stackedAreaColumn")return index<parsed.data.series.length-1?"area":"column";
    return normalizedType==="line"?"line":normalizedType==="area"?"area":normalizedType==="combo"&&index===parsed.data.series.length-1?"line":"column";
  };
  // A scatter series whose X cells are the chart's category cells (Excel
  // writes a combination's X Y series so; X03_chart_combined07) is plotted
  // at those values taken as slot numbers: read from the categories as they
  // are now, text X as 1, 2, 3, ... (Excel's rule for text X values).
  const sameCells=(a:string|undefined,b:string|undefined)=>Boolean(a&&b)&&a!.replace(/\$/g,"").toLowerCase()===b!.replace(/\$/g,"").toLowerCase();
  const categoryX=(()=>{
    const numbers=parsed.data.categories.map(label=>label.trim()===""?Number.NaN:Number(label));
    return numbers.every(Number.isFinite)?numbers:parsed.data.categories.map((_label,index)=>index+1);
  })();
  return {
    ...parsed.data,
    series: parsed.data.series.map((series, index) => {
      const renderAs=render(index);
      const ref=subtype==="custom"&&refs?refs[index]:undefined;
      const xValues=renderAs==="scatter"&&ref?.xValues&&sameCells(ref.xValues,refs?.[0]?.categories)&&categoryX.length===series.values.length?categoryX:undefined;
      return {
        ...series,
        renderAs,
        axis: subtype==="custom"&&refs?refs[index]?.axis??"primary":secondary.has(index)||(subtype==="columnLineSecondary"&&index===parsed.data.series.length-1)?"secondary":"primary",
        ...(xValues?{xValues}:{}),
      };
    }),
  };
}

/**
 * A combination series' X values over new data: the live reading's (an X of
 * 1, 2, 3, ... is each point over its own category, which the model does not
 * state), else the model's own while it still fits the series (X cells the
 * sheet's reading does not read, kept as the part stated them).
 */
export function comboSeriesXValuesOver(
  live: { renderAs?: ComboRenderAs; values: readonly unknown[]; xValues?: readonly number[] },
  own: { renderAs?: ComboRenderAs; xValues?: readonly number[] } | undefined,
): number[] | undefined {
  if ((own?.renderAs ?? live.renderAs) !== "scatter") return undefined;
  if (live.xValues) {
    return live.xValues.every((value, index) => value === index + 1) ? undefined : [...live.xValues];
  }
  return own?.xValues && own.xValues.length === live.values.length ? [...own.xValues] : undefined;
}

/**
 * Each row's branch path, outermost first, from a multi-level category's
 * levels (a group's label stands on its first row and "" on the rows after
 * it, as `CategoryChartData.categoryLevels` holds them): Asia / China, Asia /
 * Japan, Europe / France...
 */
function hierarchyPathLabels(
  levels: readonly (readonly string[])[],
  leaves: readonly string[],
): string[] {
  const carried = levels.map(() => "");
  return leaves.map((leaf, row) => {
    levels.forEach((level, depth) => {
      const entry = level[row] ?? "";
      if (entry === "") return;
      carried[depth] = entry;
      for (let inner = depth + 1; inner < levels.length; inner += 1) carried[inner] = "";
    });
    return [...carried.filter((part) => part !== ""), leaf].join(CHART_HIERARCHY_PATH_SEPARATOR);
  });
}








function hierarchyDataFromLabels(
  labels: readonly string[],
  values: readonly number[],
): HierarchyChartData {
  if (!labels.some((label) => label.includes(CHART_HIERARCHY_PATH_SEPARATOR))) {
    return {
      kind: "hierarchy",
      points: labels.flatMap((label, index) => values[index] < 0 ? [] : [{
        id: `point-${index + 1}`,
        label,
        value: values[index],
      }]),
    };
  }
  const points = new Map<string, HierarchyChartPoint>();
  labels.forEach((label, index) => {
    if (values[index] < 0) return;
    const path = label.split(CHART_HIERARCHY_PATH_SEPARATOR);
    let parentId: string | undefined;
    path.forEach((part, depth) => {
      const key = path.slice(0, depth + 1).join(CHART_HIERARCHY_PATH_SEPARATOR);
      const point = points.get(key) ?? {
        id: `point-${points.size + 1}`,
        ...(parentId ? { parentId } : {}),
        label: part,
        value: 0,
      };
      point.value += values[index];
      points.set(key, point);
      parentId = point.id;
    });
  });
  return { kind: "hierarchy", points: [...points.values()] };
}

function dataCounts(data: ChartData): { series: number; categories: number } {
  const record = data as { series?: unknown[]; categories?: unknown[] };
  return {
    series: Array.isArray(record.series) ? record.series.length : 0,
    categories: Array.isArray(record.categories) ? record.categories.length : 0,
  };
}

/**
 * The model over other data (its binding re-read, a live refresh, a retarget):
 * what the format addresses by series, point or category is kept where the new
 * data still has it. Not validated -- the caller's command or planner is.
 */
export function spreadsheetChartModelOverData(attrs: ChartBlockAttrsV2, data: ChartData): ChartBlockAttrsV2 {
  const { series, categories } = dataCounts(data);
  const format = { ...(attrs.format ?? {}) } as EChartsRecord;
  const within = (key: string, index: string, bound: number) => {
    if (!Array.isArray(format[key])) return;
    const kept = (format[key] as EChartsRecord[]).filter((entry) => (entry[index] as number) < bound);
    if (kept.length) format[key] = kept;
    else delete format[key];
  };
  within("seriesStyles", "seriesIndex", series);
  within("sliceStyles", "pointIndex", categories);
  within("seriesTrendlines", "seriesIndex", series);
  for (const [key, bound] of [["hiddenSeries", series], ["hiddenCategories", categories]] as const) {
    if (!Array.isArray(format[key])) continue;
    const kept = (format[key] as number[]).filter((index) => index < bound);
    if (kept.length && kept.length < bound) format[key] = kept;
    else delete format[key];
  }
  // A series' or a point's own labels, and a Quick Layout's, go with them.
  const pointed = (entry: { seriesIndex: number; pointIndex?: number }) =>
    entry.seriesIndex < series && (entry.pointIndex === undefined || entry.pointIndex < categories);
  if (Array.isArray(format.dataLabelOverrides)) {
    const kept = (format.dataLabelOverrides as { seriesIndex: number; pointIndex?: number }[]).filter(pointed);
    if (kept.length) format.dataLabelOverrides = kept;
    else delete format.dataLabelOverrides;
  }
  const layout = format.layout as { labels?: Array<{ seriesIndex: number; pointIndex?: number }> } | undefined;
  if (layout?.labels?.length) format.layout = { ...layout, labels: layout.labels.filter(pointed) };
  const next = { ...attrs, data } as ChartBlockAttrsV2;
  if (Object.keys(format).length) next.format = format as ChartBlockFormat;
  else delete next.format;
  return next;
}

/**
 * The stored model over the chart's live data: what the sheet draws, and what
 * the next reconciliation stores (lib/spreadsheet/chart-v2-model). The cells'
 * values, names and categories as they are now, read in the model's own data
 * shape, with what only the model states about them kept (a combination's
 * per-series chart type and axis, a waterfall's totals). A model whose data is
 * its own (`binding.literal`), or whose kind the cells are not read as, keeps
 * its data as stored. Null when the result is not a chart the shared contract
 * states.
 */
function storedModelOverLiveData(
  model: ChartBlockAttrsV2,
  live: ChartData,
  ownData: boolean,
): ChartBlockAttrsV2 | null {
  let next = model;
  if (!ownData && live.kind === model.data.kind) {
    let data: ChartData = live;
    if (live.kind === "category" && model.data.kind === "category") {
      const stated = model.data.series;
      data = {
        ...live,
        series: live.series.map((series, index) => {
          const own = stated[index];
          if (!own) return series;
          const { xValues: _liveX, ...rest } = series;
          const xValues = comboSeriesXValuesOver(series, own);
          return {
            ...rest,
            ...(own.renderAs !== undefined ? { renderAs: own.renderAs } : {}),
            ...(own.axis !== undefined ? { axis: own.axis } : {}),
            // Its values' source format (Linked to source): the cells read
            // as text here, so a secondary-axis combination's axes, which
            // take their formats from the series, kept #,##0 / 0.00% only
            // while the model's own data was drawn (XP003_out, X-291).
            ...(own.numberFormat !== undefined && rest.numberFormat === undefined
              ? { numberFormat: own.numberFormat }
              : {}),
            ...(xValues ? { xValues } : {}),
          };
        }),
      };
    } else if (live.kind === "xy" && model.data.kind === "xy") {
      // A scatter series' secondary axis and its values' source format are
      // the model's (`withModelSeriesRoles`, R209-X01).
      const stated = model.data.series;
      data = {
        ...live,
        series: live.series.map((series, index) => {
          const own = stated[index];
          if (!own) return series;
          return {
            ...series,
            ...(own.axis !== undefined ? { axis: own.axis } : {}),
            ...(own.numberFormat !== undefined && series.numberFormat === undefined
              ? { numberFormat: own.numberFormat }
              : {}),
          };
        }),
      };
    } else if (
      live.kind === "waterfall" &&
      model.data.kind === "waterfall" &&
      model.data.pointKinds.length === live.categories.length
    ) {
      data = { ...live, pointKinds: [...model.data.pointKinds] };
    }
    const before = dataCounts(model.data);
    const after = dataCounts(data);
    next =
      before.series === after.series && before.categories === after.categories
        ? { ...model, data }
        : spreadsheetChartModelOverData(model, data);
  }
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/**
 * The model as the sheet's plot draws it. A sheet chart's title and its axis
 * titles are drawn by the sheet around the plot -- the chart's header and its
 * axis strips, where they are edited in place -- so the plot keeps no room
 * for them: the chart title is left out by the planner's own switch
 * (`documentChartRenderPlan`'s `includeTitle`, as a Docs host that draws no
 * title uses it), the axis titles, which the planner draws as axis names, are
 * left out of this view. The model itself keeps them all.
 */
export function sheetPlotAttrs(attrs: ChartBlockAttrsV2): ChartBlockAttrsV2 {
  if (attrs.axisTitleX === undefined && attrs.axisTitleY === undefined) return attrs;
  const plot: ChartBlockAttrsV2 = { ...attrs };
  delete plot.axisTitleX;
  delete plot.axisTitleY;
  return plot;
}

/**
 * Chart Filters name a category by its place in the chart's whole source
 * range, as Excel keeps them in the part; a chart that plots visible cells
 * only draws fewer points, so the hidden categories are re-counted among the
 * points drawn (`sourcePointIndices`). Without it, hiding another source row
 * moved the filter onto the next category: only 丙 picked, 甲 hidden on the
 * sheet, and the chart showed 丁 (live audit S-105).
 */
export function hiddenCategoriesAmongDrawn(
  hidden: readonly number[] | undefined,
  sourcePointIndices: readonly number[] | undefined,
): number[] | undefined {
  if (!hidden || !sourcePointIndices) return hidden ? [...hidden] : hidden;
  const sources = new Set(hidden);
  const drawn: number[] = [];
  sourcePointIndices.forEach((source, index) => {
    if (sources.has(source)) drawn.push(index);
  });
  return drawn;
}

/** The inverse: Chart Filters picked among the drawn points, stated over the
 * whole source range; a filter on a source row not drawn now is kept. */
export function hiddenCategoriesAmongSource(
  drawnHidden: readonly number[] | undefined,
  previousSourceHidden: readonly number[] | undefined,
  sourcePointIndices: readonly number[] | undefined,
): number[] | undefined {
  if (!sourcePointIndices) return drawnHidden ? [...drawnHidden] : drawnHidden;
  const drawnSources = new Set(sourcePointIndices);
  const kept = (previousSourceHidden ?? []).filter((source) => !drawnSources.has(source));
  const picked = (drawnHidden ?? [])
    .map((index) => sourcePointIndices[index])
    .filter((source): source is number => source !== undefined);
  const all = [...new Set([...kept, ...picked])].sort((a, b) => a - b);
  return all.length ? all : undefined;
}

function withDrawnHiddenCategories(input: SpreadsheetChartV2Input): SpreadsheetChartV2Input {
  const points = input.sourcePointIndices;
  if (!points) return input;
  const identity = points.every((source, index) => source === index);
  const formatHidden = input.chartFormat?.hiddenCategories;
  const storedHidden = input.storedModel?.format?.hiddenCategories;
  if (identity && (!formatHidden || formatHidden.every((index) => index < points.length)) &&
    (!storedHidden || storedHidden.every((index) => index < points.length))) return input;
  const next: SpreadsheetChartV2Input = { ...input };
  if (formatHidden) {
    const drawn = hiddenCategoriesAmongDrawn(formatHidden, points)!;
    const format = { ...input.chartFormat! };
    if (drawn.length) format.hiddenCategories = drawn;
    else delete format.hiddenCategories;
    next.chartFormat = format;
  }
  if (storedHidden && input.storedModel) {
    const drawn = hiddenCategoriesAmongDrawn(storedHidden, points)!;
    const format = { ...(input.storedModel.format ?? {}) };
    if (drawn.length) format.hiddenCategories = drawn;
    else delete format.hiddenCategories;
    next.storedModel = { ...input.storedModel, format } as ChartBlockAttrsV2;
  }
  return next;
}

/**
 * A series that names nothing (its reference states no `c:tx`) is named by
 * Excel in the reader's language and by its place, with no space: 系列1 in
 * Chinese Excel, Series1 in English (X19_chart_blank06). The sheet's reading
 * spells such a name "Series n"; the picture draws Excel's words when the
 * host hands them (`defaultSeriesName`, the catalogue's "{index}" template).
 * Display only: the stored model keeps the reading's own.
 */
/** The name the reading gives a series that names nothing: "Series n" by its place. */
export function isReadingDefaultSeriesName(name: unknown, index: number): boolean {
  return name === `Series ${index + 1}`;
}

function withDefaultSeriesNames(data: ChartData, input: SpreadsheetChartV2Input): ChartData {
  const template = input.defaultSeriesName;
  // A stock chart's series too (X14_chart_crossing07's three unnamed
  // high, low and close series read 系列1..3 in Excel).
  if (
    !template ||
    !template.includes("{index}") ||
    (data.kind !== "category" && data.kind !== "xy" && data.kind !== "stock" && data.kind !== "sample")
  )
    return data;
  // An Office 2016 chart (a histogram's, a Pareto's or a box plot's sample)
  // names it with a space before the place: 系列 1 (Excel for Mac,
  // C01_pareto-gapwidth's legend), where a classic chart says 系列1.
  const name = data.kind === "sample" ? template.replace(/\s*\{index\}/, " {index}") : template;
  const refs = normalizeChartSeriesRefs(input.seriesRefs);
  // A chart of its own data (no references, C07_axis_title_default_rotation's
  // six XY series without c:tx): the reading spells an unnamed series
  // "Series n" by its place, and that is the only name it gives it.
  const unnamed = (entry: { name: string }, index: number) =>
    refs ? !refs[index]?.title && !refs[index]?.titleText && Boolean(refs[index]) : isReadingDefaultSeriesName(entry.name, index);
  let changed = false;
  const series = (data.series as readonly { name: string }[]).map((entry, index) => {
    if (!unnamed(entry, index)) return entry;
    changed = true;
    return { ...entry, name: name.replace("{index}", String(index + 1)).slice(0, 255) };
  });
  return changed ? ({ ...data, series } as ChartData) : data;
}

export function buildSpreadsheetChartV2Binding(
  rawInput: SpreadsheetChartV2Input,
): SpreadsheetChartV2Binding | null {
  const input = withDrawnHiddenCategories(rawInput);
  const normalizedType = normalizeSpreadsheetChartType(input.type);
  if (normalizedType === "camera") return null;
  let definition = SPREADSHEET_CHART_V2_DEFINITIONS[normalizedType];
  if (input.chartSubtype) {
    if (!(CHART_SUBTYPES_BY_FAMILY[definition.family] as readonly string[]).includes(input.chartSubtype)) return null;
    definition = { ...definition, subtype: input.chartSubtype as ChartSubtype };
  }
  const lineSeriesStyles = normalizeChartLineSeriesStyles(
    input.lineSeriesStyles,
  );
  // Text among a chart's values plots as 0 in Excel, every family: a
  // classic chart's as much as a box plot's (Excel for Mac, rtfix_x210: "",
  // "abc" and #DIV/0! cache as 0, #N/A as #N/A, an empty cell as no point;
  // X-210). Only a classic chart takes an error value as 0 as well.
  const officeTwentySixteen = definition.family === "statistical" || definition.family === "waterfall";
  let parsed = parseCategoryTable(
    input.rows,
    input.headerRow,
    input.categoryLevels,
    true,
    (input.storedModel?.format?.axisOptions as { category?: { axisType?: unknown } } | undefined)?.category?.axisType === "date" ||
      (input.chartFormat?.axisOptions as { category?: { axisType?: unknown } } | undefined)?.category?.axisType === "date",
    !officeTwentySixteen,
  );
  if (!parsed && definition.family === "xy") {
    // A scatter whose series have lengths of their own shares no row matrix
    // (P22_DataTableCities: 12, 12, 1 and 1 points from two sheets): its
    // native series are the whole of its data, and the chart is drawn from
    // them rather than left to the fallback sketch without axes (X-121).
    const nativeXY = xyDataFromNativeSeries(input.xySeries, normalizedType === "bubble") ?? ownXYData(input);
    if (nativeXY) parsed = { data: { kind: "category", categories: [], series: [] }, table: nativeXY.table };
  }
  if (!parsed) return null;
  if (normalizedType === "line" && lineSeriesStyles && !input.chartSubtype) {
    // A line without markers only when every series it draws says "none": a
    // series the styles leave out keeps its automatic marker (S-126: setting
    // the second series to none took the first series' circles away too
    // once the workbook was reopened).
    const unmarked = new Set(
      lineSeriesStyles.filter((style) => style.markerSymbol === "none").map((style) => style.seriesIndex),
    );
    definition = {
      ...definition,
      subtype: parsed.data.series.every((_series, index) => unmarked.has(index))
        ? "standard"
        : "standardMarkers",
    };
  }
  // A funnel draws a negative stage as an empty one (Excel for Mac 16 measured
  
  // negative point out and plots the rest, its line reaching 100% over them
  
  // 甲 戊 乙 丁).
  const paretoWithoutNegatives = normalizedType === "pareto";
  // Families that size, stack or nest by value have no empty point: a gap is 0.
  const gapless=(values:readonly ChartPointValue[])=>values.map(value=>value!==null&&Number.isFinite(value)?value:0);

  let data: ChartData = parsed.data;
  let table = parsed.table;
  if (definition.family === "xy") {
    const nativeXY = xyDataFromNativeSeries(
      input.xySeries,
      normalizedType === "bubble",
    );
    // A chart of its own data whose points have no Y at all (a bubble
    // series stating only its sizes, C24_invertIfNeg_bubble; X-127) leaves
    // the sheet no point to read: its model's data is what it plots.
    const own = nativeXY ? null : ownXYData(input);
    const xyData =
      nativeXY?.data ??
      xyDataFromCategoryTable(parsed, normalizedType === "bubble") ??
      own?.data;
    if (!xyData) return null;
    data = xyData;
    table = nativeXY?.table ?? (xyData === own?.data ? own.table : parsed.table);
  } else if (definition.family === "pie") {
    // A pie draws its first series only; a doughnut draws every series as a
    
    // not draw, as Excel keeps them in the part and the model reads them
    // (the planner draws series 0, the summary and table say so): cutting
    // them here left an imported pie's model naming paint of a series it no
    // longer had, which the validator refused, so the sheet dropped the
    // model and drew the chart from its flat fields (C17_strict_chart).
    data = {
      ...parsed.data,
      series: parsed.data.series.map((series) => ({ ...series, values: gapless(series.values) })),
    };
  } else if (definition.family === "stock") {
    const stock = stockData(parsed,input.chartSubtype);
    if (!stock) return null;
    data = stock.data;
    definition = { ...definition, subtype: stock.subtype };
  } else if (definition.family === "map") {
    // A workbook's cells name each country as its author typed it; an imported
    // map's own geocoding cache says which country each such name is.
    const bound=new Map((input.mapRegionCodes??[]).map(entry=>[entry.label,entry.id]));
    const codes=parsed.data.categories.map(code=>/^\d{3}$/.test(code)?code:bound.get(code)??"");
    if(codes.some(code=>!MAP_COUNTRY_CODES.has(code))||new Set(codes).size!==codes.length||parsed.data.series.length!==1)return null;
    data={kind:"map",regions:codes.map((id,index)=>({id,label:parsed.data.categories[index],value:gapless(parsed.data.series[0].values)[index]}))};
  } else if (definition.family === "waterfall") {
    const waterfall: WaterfallChartData = {
      kind: "waterfall",
      categories: parsed.data.categories,
      values: gapless(parsed.data.series[0].values),
      pointKinds: parsed.data.categories.map(() => "value" as const),
    };
    data = waterfall;
  } else if (definition.family === "hierarchy") {
    // The outer text columns (category levels) are the outer branches.
    const levels = parsed.data.categoryLevels;
    data = hierarchyDataFromLabels(
      levels?.length ? hierarchyPathLabels(levels, parsed.data.categories) : parsed.data.categories,
      gapless(parsed.data.series[0].values),
    );
    table = {
      ...parsed.table,
      rows: parsed.table.rows.map(([category, ...cells]) => [
        category.split(CHART_HIERARCHY_PATH_SEPARATOR).join(" / "),
        ...cells,
      ]),
    };
  } else if (definition.family === "statistical") {
    const sample: SampleChartData = {
      kind: "sample",
      series: parsed.data.series.map((series) => ({
        name: series.name,
        values: paretoWithoutNegatives
          ? plottedValues(series.values).filter((value) => value >= 0)
          : plottedValues(series.values),
      })),
    };
    data = sample;
  } else if (definition.family === "surface") {
    // Excel plots a surface over a matrix: a row per category, a column per
    // series, each cell one height.
    // A contour names a category over several levels as Excel's flat label
    // does: the leaf, and under it each outer level that starts a group there
    // (PH006: "Qtr 1" over "2008", then "Qtr 2", "Qtr 3"...; X-255).
    const levels = parsed.data.categoryLevels;
    const contour = definition.subtype === "contour" || definition.subtype === "wireframeContour";
    data = {
      kind: "matrix",
      xCategories: contour && levels?.length
        ? parsed.data.categories.map((leaf, index) =>
          [leaf, ...[...levels].reverse().map((level) => level[index] ?? "")].filter((label) => label !== "").join("\n"))
        : parsed.data.categories,
      yCategories: parsed.data.series.map((series) => series.name),
      values: parsed.data.series.map((series) => gapless(series.values)),
    };
  } else if (
    definition.family === "combo" ||
    (input.secondaryAxisSeries?.length ?? 0) > 0
  ) {
    if (parsed.data.series.length < 2) return null;
    definition = {
      family: "combo",
      subtype: definition.family === "combo" ? definition.subtype : "custom",
    };
    data = comboData(parsed, input, normalizedType);
  }
  data = withDefaultSeriesNames(data, input);

  // The one preset table (`c:style` <-> `stylePreset`) the Docs and deck
  // galleries, the canvas and every writer share.
  const appearancePreset = spreadsheetChartStyleAppearance(input.chartStyle);
  // An imported Office 2016 chart keeps the look its own style parts give it
  // until a gallery pick, baked into the part and read back as `chartStyle`,
  // replaces it; the palette control still wins over both.
  // Excel's built-in text (`chart-builtin-text`) is no style part's look:
  // a classic chart's `c:style` leaves it.
  const nativeLook = input.chartStyle === undefined || input.nativeLook?.builtInText ? input.nativeLook : undefined;
  const lookPaint = (value: string | undefined) => (value && value !== "none" ? value : undefined);
  // A look colour keeps the opacity its style gives it, so a faint line on a
  // dark chart area stays faint there.
  const lookFill = (value: { fill?: string; fillOpacity?: number } | undefined) => {
    const paint = lookPaint(value?.fill);
    return paint === undefined ? undefined : nativeLookRgba(paint, value?.fillOpacity);
  };
  const lookLine = (value: { outline?: string; outlineOpacity?: number } | undefined) => {
    const paint = lookPaint(value?.outline);
    return paint === undefined ? undefined : nativeLookRgba(paint, value?.outlineOpacity);
  };
  const axisLine = lookLine(nativeLook?.axis?.line);
  // How many things the chart's colours paint: a pie's points (and a Pie of
  // Pie's "Other"), a waterfall's three kinds, a treemap's top branches,
  // else its series.
  const paintedCount =
      definition.family === "pie"
        ? // A pie of pie's "Other" slice takes the colour after its points'.
          parsed.data.categories.length +
          (definition.subtype === "pieOfPie" || definition.subtype === "barOfPie" ? 1 : 0)
        : // A waterfall paints its increases, decreases and totals apart.
          definition.family === "waterfall"
          ? 3
          : // A treemap or sunburst paints each top-level branch apart.
            data.kind === "hierarchy"
            ? Math.max(1, data.points.filter((point) => !point.parentId).length)
            : "series" in data && Array.isArray(data.series)
          ? data.series.length
          : parsed.data.series.length;
  const appearance: SpreadsheetChartAppearance = {
    background: lookFill(nativeLook?.chartArea) ?? appearancePreset.background,
    plotBackground:
      lookFill(nativeLook?.plotArea) ??
      lookFill(nativeLook?.chartArea) ??
      appearancePreset.plotBackground,
    text: nativeLook?.axis?.color ?? nativeLook?.title?.color ?? appearancePreset.text,
    ...(nativeLook?.dataLabel?.color ? { labelText: nativeLook.dataLabel.color } : {}),
    ...(nativeLook?.legend?.color ? { legendText: nativeLook.legend.color } : {}),
    ...(axisLine ? { axisLine } : {}),
    gridline: lookLine(nativeLook?.gridline) ?? appearancePreset.gridline,
    border: lookLine(nativeLook?.chartArea) ?? appearancePreset.border,
    // A Monochromatic pick spreads its shades over the series (a pie's points).
    palette: spreadsheetChartColorPaletteColors(
      input.chartColorPalette,
      input.themeAccents ?? appearancePreset.palette,
      paintedCount,
    ) ??
      nativeLookPalette(nativeLook) ??
      // Office's default style paints the workbook theme's accents; Excel's
      // built-in look shades and tints them round by round past the sixth
      // series (X-046).
      // An imported part in Excel's built-in look paints its c:style's own
      // colours over the workbook theme (X-071: style 11 is accent 1 dark and
      // light, not the gallery preset the number maps to here).
      // Office 2010 keeps the style in an mc:AlternateContent the writer
      // leaves alone, so no `chartStyle` is read; the number the server read
      // for the built-in look paints it all the same (U01_aaa's scatter,
      // c14:style 123 = style 23: accent 5, X-196).
      (nativeLook?.builtInText &&
      (input.chartStyle ?? nativeLook.builtInStyle?.style) !== undefined &&
      input.themeAccents?.length
        ? excelBuiltInChartStyleColors(
            (input.chartStyle ?? nativeLook.builtInStyle?.style)!,
            input.themeAccents,
            paintedCount,
          )
        : null) ??
      (appearancePreset.stylePreset===DEFAULT_CHART_STYLE_PRESET_ID&&input.themeAccents?.length
        ? nativeLook?.builtInText
          ? excelBuiltInStyleSeriesColors(input.themeAccents, paintedCount)
          : input.themeAccents
        : appearancePreset.palette),
  };
  const seriesCount =
    data.kind === "xy" || data.kind === "stock" || data.kind === "sample"
      ? data.series.length
      : parsed.data.series.length;
  const categoryCount =
    data.kind === "xy"
      ? Math.max(0, ...data.series.map((series) => series.xValues.length))
      : parsed.data.categories.length;
  const summary = `${normalizedType} chart with ${seriesCount} series and ${categoryCount} plotted ${categoryCount === 1 ? "point" : "points"}.`;
  const attrs: ChartBlockAttrsV2 = {
    schemaVersion: 2,
    chartType: definition.family,
    chartSubtype: definition.subtype,
    title: boundedText(input.title, "Spreadsheet chart"),
    data,
    stylePreset: appearancePreset.stylePreset,
    format: spreadsheetChartRenderFormat({
      // Every kind but the moving average is drawn by the shared Chart V2
      // renderer, through the fit the three editors share; the moving average
      // keeps the sheet's own overlay below (its window is free, 2–1000).
      ...(input.trendline && input.trendline.type !== "movingAverage"
        ? (() => {
            // With its options: the forecast, the pinned intercept and the
            // equation / R² label are the shared renderer's to draw.
            const trendline = spreadsheetTrendlineContract(input.trendline);
            return trendline ? { trendline } : {};
          })()
        : {}),
      ...(input.errorBars?.type === "percent" && input.errorBars.value===5 ? {errorBars:{type:"percentage" as const,value:5 as const}} : input.errorBars?.type === "stddev" && input.errorBars.value===1 ? {errorBars:{type:"standardDeviation" as const,value:1 as const}} : input.errorBars?.type === "stderr" ? {errorBars:{type:"standardError" as const}}:{}),
      ...(input.upDownBars ? {upDownBars:true}:{}),
      ...(input.dropLines ? {dropLines:true}:{}),
      ...(input.highLowLines ? {highLowLines:true}:{}),
      ...(input.seriesLines ? {seriesLines:true}:{}),
      // A filled map's labels are Office's Map labels tri-state, stated in the
      // part (`cx:regionLabelLayout`) and carried in the native envelope. The
      // block's boolean defaults to false for every family, and manufacturing
      // that false here would report "labels off" for an imported map whose
      // part says `bestFitOnly` — the reverse of the drawing, which reads the
      // part. When nobody has said anything, say nothing.
      ...(definition.family === "map" &&
      input.showDataLabels === undefined &&
      input.chartFormat?.dataLabels === undefined
        ? {}
        : { dataLabels: input.showDataLabels === true }),
      ...(chartSupportsAxisVisibility(definition.family) ? { axisX: input.showAxisLabels !== false, axisY: input.showAxisLabels !== false } : {}),
      ...(chartSupportsGridlines(definition.family) ? { gridlinesX: input.showGridlines !== false, gridlinesY: input.showGridlines !== false } : {}),
      // Format Legend ▸ Top Right is the model's `topRight` as well.
      legend:
        input.showLegend === false
          ? "none"
          : (input.legendPosition ?? "top"),
      ...(normalizePresentationChartSeriesStyles(input.seriesStyles)
        ? {
            seriesStyles: normalizePresentationChartSeriesStyles(
              input.seriesStyles,
            ),
          }
        : {}),
      ...(definition.family === "pie" &&
      normalizePresentationChartSliceStyles(input.sliceStyles)
        ? {
            sliceStyles: normalizePresentationChartSliceStyles(
              input.sliceStyles,
            ),
          }
        : {}),
    }, input.chartFormat),
  };

  // The chart's stored model over the live data, when the host hands it: what
  // the planner draws (its refusal falls back to the flat reading).
  // A chart of its own data draws its model's data: its unnamed series take
  // Excel's localized names there too (C07: 系列1..6; X-049).
  const stored = input.storedModel
    ? storedModelOverLiveData(
        input.storedModelOwnData === true
          ? { ...input.storedModel, data: withDefaultSeriesNames(input.storedModel.data, input) }
          : input.storedModel,
        data,
        input.storedModelOwnData === true,
      )
    : null;
  {
    // Worksheet-native ChartEx axes/labels are projected after the shared
    // family planner. Keep their persisted atoms; don't feed unsupported Docs
    // authoring flags into its deliberately narrower node validator.
    const format={...attrs.format};
    // Change Colors: the sheet's pick is the model's, the one palette key every
    // host states (and every writer turns into the chart's colour part).
    const palette=normalizeSpreadsheetChartColorPalette(input.chartColorPalette);
    if(palette&&chartTakesColorPalette(definition.family))format.colorPalette=palette;
    if(!chartSupportsAxisVisibility(definition.family)){delete format.axisX;delete format.axisY;delete format.axisStyles;}
    if(!chartSupportsGridlines(definition.family)){delete format.gridlinesX;delete format.gridlinesY;delete format.gridlineStyle;}
    if(!chartSupportsMinorGridlines(definition.family,definition.subtype)){delete format.minorGridlinesX;delete format.minorGridlinesY;}
    // Chart Filters: a filter the chart's type cannot state, or that names a
    // series / category the bound range no longer has (or would leave none
    // plotted), is not drawn; the sheet's chart still is.
    {
      const seriesTotal=data.kind==="category"||data.kind==="xy"?data.series.length:0;
      const categoryTotal=data.kind==="category"?data.categories.length:0;
      const fits=(hidden:readonly number[]|undefined,total:number,keep:number)=>
        !hidden||(hidden.length>0&&hidden.length<=total-keep&&hidden.every(index=>index<total));
      if(!chartSupportsSeriesFilter(definition.family)||!fits(format.hiddenSeries,seriesTotal,definition.family==="combo"?2:1))delete format.hiddenSeries;
      if(!chartSupportsCategoryFilter(definition.family)||!fits(format.hiddenCategories,categoryTotal,1))delete format.hiddenCategories;
    }
    if(!chartSupportsDataTable(definition.family)){delete format.dataTable;delete format.dataTableStyle;delete format.dataTableLegendKeys;}
    // A treemap's Label Options are a treemap's: the sheet's chart made a
    // sunburst draws (and writes) none.
    if(!chartTakesParentLabelLayout(definition.family,definition.subtype))delete format.parentLabelLayout;
    if(!chartSupportsDataLabels(definition.family)){
      for(const key of Object.keys(format))if(key.startsWith("dataLabel"))delete (format as Record<string,unknown>)[key];
    }
    // A family may carry labels and still state no POSITION for them (ECMA-376
    // names a histogram, a Pareto and a radar): keep the labels the file
    // carries and drop the position, rather than drawing nothing at all.
    else{
      // Data Callout draws only on a 2-D classic chart whose labels show.
      if(format.dataLabelCallout!==undefined&&(format.dataLabels!==true||!chartSupportsDataLabelCallout(definition.family,definition.subtype)))delete format.dataLabelCallout;
      if(format.dataLabelPosition!==undefined&&!chartDataLabelPositions(definition.family,definition.subtype).includes(format.dataLabelPosition)){
        delete format.dataLabelPosition;
      }
      // The same rule for the one family whose labels are a Map labels layout
      // rather than a position: keep the labels, drop a pick this family has
      // no place for.
      if(format.dataLabelRegionLayout!==undefined&&!chartMapLabelLayouts(definition.family).includes(format.dataLabelRegionLayout)){
        delete format.dataLabelRegionLayout;
      }
    }
    if(data.kind!=="category"&&data.kind!=="xy"&&data.kind!=="stock")delete format.seriesStyles;
    attrs.format=format;
  }
  // The summary and the data table describe what is drawn.
  const drawnFormat = (stored ?? attrs).format;
  const drawn = drawnChartTable(
    table,
    data,
    definition.family,
    drawnFormat?.hiddenSeries,
    drawnFormat?.hiddenCategories,
    definition.family === "pie" && definition.subtype !== "doughnut",
  );
  if (drawn) {
    return {
      attrs, stored, nativeLook, parsed, table: drawn.table, appearance,
      summary: `${normalizedType} chart with ${drawn.series} series and ${drawn.categories} plotted ${drawn.categories === 1 ? "point" : "points"}.`,
      categoryCount: drawn.categories, seriesCount: drawn.series,
    };
  }
  return { attrs, stored, nativeLook, parsed, table, appearance, summary, categoryCount, seriesCount };
}

/**
 * What a worksheet chart draws as, without drawing it: the chart's stored
 * model over its live data when the shared contract states it, else the flat
 * reading when that is a valid model; null when neither is (the sheet keeps
 * its legacy picture). The editors' data-only questions ask this (the Format
 * pane's model, Change Chart Type's support check, the model reconciliation);
 * the picture itself is the server chart engine's.
 */
export function spreadsheetChartV2Drawn(input: SpreadsheetChartV2Input): SpreadsheetChartV2Drawn | null {
  const binding = buildSpreadsheetChartV2Binding(input);
  if (!binding) return null;
  const { attrs, stored, table, appearance, summary, categoryCount, seriesCount } = binding;
  const shared = { accessibleTable: table, appearance, summary, categoryCount, seriesCount };
  if (stored && validateChartBlockAttrsV2(sheetPlotAttrs(stored)).valid) return { ...shared, attrs: stored, drawnFromModel: true };
  return validateChartBlockAttrsV2(attrs).valid ? { ...shared, attrs } : null;
}
