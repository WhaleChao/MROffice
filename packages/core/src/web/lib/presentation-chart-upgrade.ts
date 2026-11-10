/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";
import {
  chartSupportsAxisVisibility,
  chartSupportsDataTable,
  chartSupportsGridlines,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
} from "../components/console/editor/extensions/chartBlockShared";
import { officeChartLeafForLegacyKind } from "./office-charts/catalog";
import { officeChartWithType } from "./office-charts/chart-type";
import { seedPresentationChartWorkbook } from "./presentation-chart-desktop";
import {
  chartDataToGrid,
  chartGridToCategoriesAndSeries,
  chartValueFromCellText,
  normalizeChartType,
  presentationChartHasNativeSource,
} from "./presentation-chart-data";
import { legacyXYSeriesFromRows } from "./presentation-legacy-chart-grid";
import {
  resolveChartDataLabels,
  resolveChartDataTable,
  resolveChartGridlines,
  resolveChartLegendPosition,
  resolveChartToggles,
} from "./presentation-chart-toggles";
import { readPresentationChartRenderAttrs } from "./presentation-native-chart-view";
import { presentationBlockIsPreservedOnly } from "./presentation-preserved-object";
import {
  presentationChartV2AwaitsReceipt,
  presentationChartV2Kind,
  readPresentationChartV2,
  withPresentationChartV2,
} from "./presentation-chart-v2";
import {
  DEFAULT_SPREADSHEET_CHART_STYLE_ID,
  normalizeSpreadsheetChartStyle,
  spreadsheetChartStylePreset,
  spreadsheetChartStylePresetId,
} from "./spreadsheet/chart";
import {
  normalizePresentationChartAreaStyle,
  normalizePresentationChartAxisStyles,
  normalizePresentationChartSeriesStyles,
  normalizePresentationChartSliceStyles,
  type PresentationChartSeriesStyle,
} from "./presentation-chart-area-style";
import { parsePresentationChartDisplay } from "./presentation-chart-display";

/** A list with entries, or nothing: an emptied paint list states no paint. */
function presentOnly<T>(list: readonly T[] | undefined): readonly T[] | undefined {
  return list?.length ? list : undefined;
}

/**
 * The series paints a block states, or the ones the part's own appearance
 * (`source.chartDisplay`, an import stored before the native view reading)
 * gave its series: a line's colour and width, a bar's fill. A paint that
 * addresses a series the chart does not have is dropped, not the reading.
 */
function legacySeriesStyles(
  block: PresentationChartBlock,
  count: number,
): readonly PresentationChartSeriesStyle[] | undefined {
  const stated = presentOnly(
    normalizePresentationChartSeriesStyles(block.seriesStyles)?.filter(
      (style) => style.seriesIndex < count,
    ),
  );
  if (stated) return stated;
  const display = parsePresentationChartDisplay(block.source?.chartDisplay);
  const styles = (display?.series ?? []).slice(0, count).flatMap((series, seriesIndex) => {
    const fill = series.color ?? series.lineColor;
    const outline = series.lineColor;
    if (!fill || fill === "none") return [];
    return [
      {
        seriesIndex,
        fill,
        ...(outline ? { outline } : {}),
        ...(series.lineWidth ? { shapePaint: { outlineWidthPt: series.lineWidth } } : {}),
      },
    ];
  });
  return normalizePresentationChartSeriesStyles(styles);
}

/**
 * A series name the grid states, or the one the part gave it when the grid
 * only numbers the series ("Series 1"), as the view drew it.
 */
function legacySeriesName(block: PresentationChartBlock, index: number, name: string): string {
  const trimmed = name.trim();
  if (trimmed && !/^Series \d+$/.test(trimmed)) return trimmed;
  const display = parsePresentationChartDisplay(block.source?.chartDisplay);
  return display?.series[index]?.name || trimmed || `Series ${index + 1}`;
}

/**
 * The keys a chart block carried before the one chart model (`chartV2`):
 * its type name, data grid, switches, titles, style pick and element paints.
 * A block on the one model states all of that in its attrs; these keys are
 * dropped when it is upgraded, so no second reading of the chart remains.
 */
export const LEGACY_PRESENTATION_CHART_KEYS = [
  "data",
  "chartStyle",
  "axisTitleX",
  "axisTitleY",
  "showLegend",
  "legendPosition",
  "trendline",
  "showAxisLabels",
  "showDataLabels",
  "showGridlines",
  "showDataTable",
  "chartAreaStyle",
  "plotAreaStyle",
  "chartTitleStyle",
  "legendStyle",
  "seriesStyles",
  "sliceStyles",
  "axisStyles",
  "axisTitleStyles",
  "gridlineStyle",
  "trendlineStyle",
  "dataTableStyle",
  "dataLabelStyle",
  "chartElements",
] as const;

/**
 * The attrs a pre-V2 category chart's own fields read as: the grid as
 * category data, the block's switches as the format keys they stand for, and
 * its element paints copied over. `bakeStyle` paints the block's style-gallery
 * pick (`chartStyle`) onto every series or slice and the chart and plot areas,
 * so the upgraded chart keeps the colours the pick gave it; a block with no
 * pick (or Office's default) states nothing and draws the default look.
 */
function legacyCategoryChartAttrs(
  block: PresentationChartBlock,
  bakeStyle: boolean,
  asColumn = false,
): ChartBlockAttrsV2 | null {
  const kind = normalizeChartType(block.chartType);
  // A kind a string grid cannot state (a waterfall, a combo, a stock chart,
  // a scatter grid with no Y column) reads as a clustered column chart of
  // the grid, as the PPTX export has always written it (and reports as
  // `chart_kind_degraded:<kind>`, `chart_block_reading.py`).
  const definition = asColumn
    ? officeChartLeafForLegacyKind("column")
    : officeChartLeafForLegacyKind(kind);
  if (!definition || definition.family === "xy") return null;
  // The page's normalizer has already cut a stored grid to the editor's
  // grid; a grid at those edges may be a cut one, and a cut chart is never
  // read as the chart (the PPTX importer's complete subset: 50 categories,
  // 12 series).
  const grid = chartDataToGrid(block.data);
  if (grid.rows.length > 50 || grid.headers.length > 13) return null;
  const read = chartGridToCategoriesAndSeries(grid);
  const data = {
    ...read,
    series: read.series.map((series, index) => ({
      ...series,
      name: legacySeriesName(block, index, series.name),
    })),
  };
  const preset = spreadsheetChartStylePreset(block.chartStyle),
    toggles = resolveChartToggles(block),
    display = parsePresentationChartDisplay(block.source?.chartDisplay);
  const paint = (index: number) => preset.palette[index % preset.palette.length];
  const seriesCount = definition.family === "pie" ? Math.min(1, data.series.length) : data.series.length;
  const area = (fill: string | undefined) =>
    fill ? normalizePresentationChartAreaStyle({ fill }) : undefined;
  const attrs: ChartBlockAttrsV2 = {
    schemaVersion: 2,
    chartType: definition.family,
    chartSubtype: definition.subtype,
    title: block.title ?? "",
    stylePreset: "classic",
    data: {
      kind: "category",
      categories: data.categories,
      series: definition.family === "pie" ? data.series.slice(0, 1) : data.series,
    },
    axisTitleX: block.axisTitleX,
    axisTitleY: block.axisTitleY,
    format: {
      legend: toggles.showLegend ? resolveChartLegendPosition(block) : "none",
      // A pie or doughnut has no axes, gridlines or data table; stating them
      // makes the reading invalid.
      ...(chartSupportsAxisVisibility(definition.family)
        ? { axisX: toggles.showAxisLabels, axisY: toggles.showAxisLabels }
        : {}),
      ...(chartSupportsGridlines(definition.family)
        ? { gridlinesY: resolveChartGridlines(block) }
        : {}),
      dataLabels: resolveChartDataLabels(block),
      ...(chartSupportsDataTable(definition.family)
        ? { dataTable: resolveChartDataTable(block) }
        : {}),
      chartAreaStyle:
        normalizePresentationChartAreaStyle(block.chartAreaStyle) ??
        (bakeStyle ? { fill: preset.background, outline: "none" } : area(display?.background)),
      plotAreaStyle:
        normalizePresentationChartAreaStyle(block.plotAreaStyle) ??
        (bakeStyle ? { fill: preset.plotBackground, outline: "none" } : area(display?.plotBackground)),
      chartTitleStyle: normalizePresentationChartAreaStyle(block.chartTitleStyle),
      legendStyle: normalizePresentationChartAreaStyle(block.legendStyle),
      seriesStyles:
        legacySeriesStyles(block, seriesCount) ??
        (bakeStyle && definition.family !== "pie"
          ? data.series.map((_, seriesIndex) => ({ seriesIndex, fill: paint(seriesIndex) }))
          : undefined),
      sliceStyles:
        presentOnly(
          normalizePresentationChartSliceStyles(block.sliceStyles)?.filter(
            (style) => style.pointIndex < data.categories.length,
          ),
        ) ??
        (bakeStyle && definition.family === "pie"
          ? data.categories.slice(0, 50).map((_, pointIndex) => ({ pointIndex, fill: paint(pointIndex) }))
          : undefined),
      axisStyles: normalizePresentationChartAxisStyles(block.axisStyles),
      dataLabelStyle: normalizePresentationChartAreaStyle(block.dataLabelStyle),
      gridlineStyle: normalizePresentationChartAreaStyle(block.gridlineStyle),
      dataTableStyle: normalizePresentationChartAreaStyle(block.dataTableStyle),
      trendline: block.trendline,
      trendlineStyle: normalizePresentationChartAreaStyle(block.trendlineStyle),
    },
  };
  return validateChartBlockAttrsV2(attrs).valid ? attrs : null;
}

/**
 * The attrs a pre-V2 scatter or bubble chart reads as: its [label, x, y(,
 * size)] grid (`presentation-legacy-chart-grid`) as one XY series, the way
 * the preview drew it and the PPTX writer exported it.
 */
function legacyXYChartAttrs(block: PresentationChartBlock): ChartBlockAttrsV2 | null {
  const kind = normalizeChartType(block.chartType);
  if (kind !== "scatter" && kind !== "bubble") return null;
  const series = legacyXYSeriesFromRows(block.data?.rows, kind);
  if (!series || series.points.length === 0) return null;
  const definition = officeChartLeafForLegacyKind(series.kind)!;
  const name = block.data?.headers?.[2]?.trim() || "Series 1";
  const toggles = resolveChartToggles(block);
  const attrs: ChartBlockAttrsV2 = {
    schemaVersion: 2,
    chartType: definition.family,
    chartSubtype: definition.subtype,
    title: block.title ?? "",
    stylePreset: "classic",
    data: {
      kind: "xy",
      series: [
        {
          name,
          xValues: series.points.map((point) => point.x),
          yValues: series.points.map((point) => point.y),
          ...(series.kind === "bubble"
            ? { bubbleSizes: series.points.map((point) => point.size ?? 0) }
            : {}),
        },
      ],
    },
    axisTitleX: block.axisTitleX,
    axisTitleY: block.axisTitleY,
    format: {
      legend: toggles.showLegend ? resolveChartLegendPosition(block) : "none",
      axisX: toggles.showAxisLabels,
      axisY: toggles.showAxisLabels,
      gridlinesY: resolveChartGridlines(block),
      dataLabels: resolveChartDataLabels(block),
      chartAreaStyle: normalizePresentationChartAreaStyle(block.chartAreaStyle),
      plotAreaStyle: normalizePresentationChartAreaStyle(block.plotAreaStyle),
      chartTitleStyle: normalizePresentationChartAreaStyle(block.chartTitleStyle),
      legendStyle: normalizePresentationChartAreaStyle(block.legendStyle),
      seriesStyles: legacySeriesStyles(block, 1),
      axisStyles: normalizePresentationChartAxisStyles(block.axisStyles),
      trendline: block.trendline,
    },
  };
  return validateChartBlockAttrsV2(attrs).valid ? attrs : null;
}

/**
 * The chart a block made before the one chart model shows, as Chart V2
 * attrs: an imported chart's native view with the block's changes composed
 * onto it (what the canvas drew), else the reading of the block's own grid
 * and switches. Null for a block with `chartV2` already, and for one no
 * reading takes (a grid past the editor's envelope).
 */
export function legacyChartReading(
  block: PresentationChartBlock,
): ChartBlockAttrsV2 | null {
  if (block.chartV2 !== undefined) return null;
  const view = readPresentationChartRenderAttrs(block);
  if (view) return view;
  const style = normalizeSpreadsheetChartStyle(block.chartStyle);
  const bakeStyle = style !== undefined && style !== DEFAULT_SPREADSHEET_CHART_STYLE_ID;
  return (
    legacyCategoryChartAttrs(block, bakeStyle) ??
    legacyXYChartAttrs(block) ??
    // A grid with no rows holds no chart, of any kind.
    (block.data?.rows?.length ? legacyCategoryChartAttrs(block, bakeStyle, true) : null)
  );
}

/**
 * The block on the one chart model, as the deck loads it.
 *
 * A block this editor authored before Chart V2 becomes a V2 block: its
 * reading is its `chartV2`, and the fields that reading was made from are
 * dropped, so the block saves as V2 and no chart has two readings. An
 * imported chart stored before the server's V2 reading (`nativeChartView`
 * and no `nativeChartV2Authority`) carries the view's composed attrs as
 * `chartV2` for the canvas and the one ribbon, and keeps its stored fields
 * beside them: the receipt that lets the writer edit the part is the
 * server's to grant, so until the stored reading is upgraded the chart is
 * read-only (`presentationChartV2CanEdit`) and the save sends the stored
 * reading back untouched (`presentationChartBlockForSave`).
 */
export function upgradePresentationChartBlock(
  block: PresentationChartBlock,
): PresentationChartBlock {
  if (block.chartV2 !== undefined) return block;
  const attrs = legacyChartReading(block);
  if (!attrs) return block;
  if (presentationChartHasNativeSource(block)) return { ...block, chartV2: attrs };
  const stripped = { ...block } as Record<string, unknown>;
  for (const key of LEGACY_PRESENTATION_CHART_KEYS) delete stripped[key];
  return withPresentationChartV2(stripped as unknown as PresentationChartBlock, attrs);
}

/**
 * Whether the block holds a reading of its chart beside `chartV2`: the
 * native view of an import stored before the server's V2 reading, or the
 * fields a pre-V2 block was stored with.
 */
function presentationChartHasStoredReading(block: PresentationChartBlock): boolean {
  const source = block.source;
  const view =
    source && typeof source === "object" && !Array.isArray(source)
      ? (source as Record<string, unknown>).nativeChartView
      : undefined;
  return (
    (view !== null && typeof view === "object" && !Array.isArray(view)) ||
    LEGACY_PRESENTATION_CHART_KEYS.some((key) => block[key] !== undefined)
  );
}

/**
 * A chart block new to the deck (an AI's append, a pasted block) on the one
 * chart model: a pre-V2 block becomes its V2 reading, and a block that states
 * `chartV2` keeps no pre-V2 field beside it. An imported chart keeps its
 * stored reading (`upgradePresentationChartBlock`).
 */
export function presentationChartBlockOnV2(block: PresentationChartBlock): PresentationChartBlock {
  const upgraded = upgradePresentationChartBlock(block);
  if (upgraded.chartV2 === undefined || presentationChartHasNativeSource(upgraded)) return upgraded;
  if (!LEGACY_PRESENTATION_CHART_KEYS.some((key) => upgraded[key] !== undefined)) return upgraded;
  const stripped = { ...upgraded } as Record<string, unknown>;
  for (const key of LEGACY_PRESENTATION_CHART_KEYS) delete stripped[key];
  return stripped as unknown as PresentationChartBlock;
}

/**
 * The block as the save sends it: a chart awaiting its receipt goes back as
 * its stored reading (the attrs were the page's own view of it), so a page
 * stored before the server's V2 reading is saved as it was stored and the
 * patchback sees no change it cannot write. A chart whose one reading is its
 * `chartV2` (the browser deck import states nothing else) is saved as it is.
 */
export function presentationChartBlockForSave(
  block: PresentationChartBlock,
): PresentationChartBlock {
  // A chart the file keeps and this page cannot draw (`nativeChartBlocker`)
  // was stored with no chart reading at all: the V2 attrs the editor composed
  // on loading it stand in for nothing and go, as for a chart whose stored
  
  // stored as `chartType: "chart"` came back with a default clustered bar,
  // and every save of the deck was refused, "chartType must match the V2
  // chart").
  const storedWithoutV2 =
    presentationChartHasStoredReading(block) ||
    presentationBlockIsPreservedOnly(block as unknown as Record<string, unknown>);
  if (!presentationChartV2AwaitsReceipt(block) || !storedWithoutV2) {
    // The V2 title is the chart's title; the server refuses a block whose
    // outer `title` says otherwise. A chart authored here states both the
    // same way, and a chart that drifted is sent as the V2 reading says
    // (PPT-B399).
    const v2Title = block.chartV2?.title;
    if (typeof v2Title === "string" && (block.title ?? "") !== v2Title) {
      return { ...block, title: v2Title };
    }
    return block;
  }
  const rest = { ...block };
  delete rest.chartV2;
  delete rest.chartWorkbook;
  return rest;
}

/**
 * The keys of a chart block patch that name the chart the way a block made
 * before the one model did (`LEGACY_PRESENTATION_CHART_KEYS`) plus its two
 * labels, `title` and `chartType`. An AI command, a Copilot rewrite or an
 * old tool may still send them; they are read into `chartV2` on the way in.
 */
const LIFTED_CHART_PATCH_KEYS: ReadonlySet<string> = new Set([
  ...LEGACY_PRESENTATION_CHART_KEYS,
  "title",
  "chartType",
]);

/** The pre-V2 element paints whose record shape the V2 format shares. */
const LIFTED_CHART_STYLE_KEYS = [
  "chartAreaStyle",
  "plotAreaStyle",
  "chartTitleStyle",
  "legendStyle",
  "seriesStyles",
  "sliceStyles",
  "axisStyles",
  "gridlineStyle",
  "dataTableStyle",
  "dataLabelStyle",
  "trendlineStyle",
] as const;

type LegacyGrid = { headers?: readonly string[]; rows?: readonly (readonly string[])[] };

/** A pre-V2 grid as the data of the chart it is patched onto, or null. */
function liftedChartData(attrs: ChartBlockAttrsV2, value: unknown): ChartBlockAttrsV2["data"] | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const grid = value as LegacyGrid;
  if (attrs.data.kind === "xy") {
    const kind = attrs.chartSubtype === "bubble" ? "bubble" : "scatter";
    const series = legacyXYSeriesFromRows(grid.rows, kind);
    if (!series || series.points.length === 0 || series.kind !== kind) return null;
    return {
      kind: "xy",
      series: [
        {
          name: grid.headers?.[2]?.trim() || attrs.data.series[0]?.name || "Series 1",
          xValues: series.points.map((point) => point.x),
          yValues: series.points.map((point) => point.y),
          ...(kind === "bubble"
            ? { bubbleSizes: series.points.map((point) => Math.abs(point.size ?? 0)) }
            : {}),
        },
      ],
    };
  }
  if (attrs.data.kind !== "category") return null;
  if (!Array.isArray(grid.rows) || grid.rows.length === 0) return null;
  // Read the way Copilot's table-to-chart reads a table (`apply_slide_chart`):
  // the first column names the categories, each later column is a series
  // when it holds a number; a cell that is not a number is a missing value.
  // A grid with no number at all is no chart data: refused, not stored as
  // zeros.
  const grid2 = chartDataToGrid(grid);
  const categories = grid2.rows.map((row) => row[0] || "");
  const series = grid2.headers
    .slice(1)
    .map((name, index) => ({
      name: name || `Series ${index + 1}`,
      values: grid2.rows.map((row) => chartValueFromCellText(String(row[index + 1] ?? ""))),
    }))
    .filter((entry) => entry.values.some((value) => value !== null))
    .map((entry) => ({ name: entry.name, values: entry.values.map((value) => value ?? 0) }));
  if (series.length === 0) return null;
  return {
    kind: "category",
    categories,
    series: attrs.chartType === "pie" ? series.slice(0, 1) : series,
  };
}

/**
 * A chart block patch on the one chart model: every key that names the chart
 * the pre-V2 way is read into the block's `chartV2` (a grid into its data, a
 * kind into its type, a switch into the format key it stands for, a paint
 * record into the same record), the labels follow the attrs, and none of
 * those keys is kept. A patch without such keys comes back unchanged; one
 * whose keys the model cannot take (a gallery style index, a grid the chart's
 * kind cannot read, a result the validator refuses) is null. When the data
 * changes, a workbook that no longer reads as the data is seeded anew from
 * it (the save refuses a workbook that disagrees with the chart's cache).
 */
export function presentationChartPatchOnV2(
  block: PresentationChartBlock,
  patch: Record<string, unknown>,
): Record<string, unknown> | null {
  const lifted = Object.keys(patch).filter((key) => LIFTED_CHART_PATCH_KEYS.has(key));
  if (lifted.length === 0) return patch;
  const rest = Object.fromEntries(
    Object.entries(patch).filter(([key]) => !LIFTED_CHART_PATCH_KEYS.has(key)),
  );
  if (patch.chartV2 !== undefined) return rest;
  const current = readPresentationChartV2(block.chartV2) ?? legacyChartReading(block);
  if (!current) return null;
  let next: ChartBlockAttrsV2 = {
    ...current,
    format: current.format ? { ...current.format } : undefined,
  };
  if (typeof patch.chartType === "string" && patch.chartType !== presentationChartV2Kind(current)) {
    const leaf = officeChartLeafForLegacyKind(patch.chartType);
    if (!leaf) return null;
    const data =
      leaf.family === "pie" && next.data.kind === "category"
        ? { ...next.data, series: next.data.series.slice(0, 1) }
        : next.data;
    const retyped = officeChartWithType({ ...next, data }, leaf.family, leaf.subtype);
    // A kind the chart's data cannot take (category data asked to be a
    // scatter) leaves the chart as it is, as the kind switch always did.
    if (retyped) next = { ...retyped, format: retyped.format ? { ...retyped.format } : undefined };
  }
  if ("data" in patch) {
    const data = liftedChartData(next, patch.data);
    if (!data) return null;
    next = { ...next, data };
  }
  if ("title" in patch) {
    if (typeof patch.title !== "string") return null;
    next.title = patch.title;
  }
  for (const key of ["axisTitleX", "axisTitleY"] as const) {
    if (!(key in patch)) continue;
    const value = patch[key];
    if (value === null || value === "" || value === undefined) delete next[key];
    else if (typeof value === "string") next[key] = value;
    else return null;
  }
  const format: Record<string, unknown> = { ...(next.format ?? {}) };
  const family = next.chartType;
  if ("showLegend" in patch || "legendPosition" in patch) {
    const position =
      typeof patch.legendPosition === "string"
        ? patch.legendPosition
        : typeof format.legend === "string" && format.legend !== "none"
          ? format.legend
          : "bottom";
    format.legend = patch.showLegend === false ? "none" : position;
  }
  if (typeof patch.showDataLabels === "boolean") format.dataLabels = patch.showDataLabels;
  if (typeof patch.showGridlines === "boolean" && chartSupportsGridlines(family))
    format.gridlinesY = patch.showGridlines;
  if (typeof patch.showAxisLabels === "boolean" && chartSupportsAxisVisibility(family)) {
    format.axisX = patch.showAxisLabels;
    format.axisY = patch.showAxisLabels;
  }
  if (typeof patch.showDataTable === "boolean" && chartSupportsDataTable(family))
    format.dataTable = patch.showDataTable;
  if ("trendline" in patch) {
    if (patch.trendline === null || patch.trendline === undefined) delete format.trendline;
    else format.trendline = patch.trendline;
  }
  for (const key of LIFTED_CHART_STYLE_KEYS) {
    if (!(key in patch)) continue;
    if (patch[key] === null || patch[key] === undefined) delete format[key];
    else format[key] = patch[key];
  }
  // A gallery pick (`chartStyle`, a `c:style` index) reads through the one
  // preset table, as an imported part's `c:style` does; an index the table
  // has no preset for is no pick the model can state.
  if ("chartStyle" in patch) {
    const style = normalizeSpreadsheetChartStyle(patch.chartStyle);
    if (style === undefined) return null;
    next.stylePreset = spreadsheetChartStylePresetId(style);
  }
  // A whole-chart element list and the axis title paints have no reading on
  // the one model: the patch is refused rather than stored beside the chart,
  // where nothing would draw it.
  if ("chartElements" in patch || "axisTitleStyles" in patch) return null;
  next = { ...next, format: format as ChartBlockAttrsV2["format"] };
  if (!validateChartBlockAttrsV2(next).valid) return null;
  const out: Record<string, unknown> = {
    ...rest,
    chartV2: next,
    title: next.title,
    chartType: presentationChartV2Kind(next),
  };
  if (JSON.stringify(next.data) !== JSON.stringify(current.data) && block.chartWorkbook !== undefined) {
    const seeded = seedPresentationChartWorkbook(next);
    out.chartWorkbook = seeded ?? undefined;
  }
  return out;
}
