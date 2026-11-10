/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Design ▸ Quick Layout for a Chart V2 block, in every editor that holds
 * one (Docs, and the deck's owned V2 chart), from the ONE catalogue
 * (`lib/office-charts/chart-layout.ts`).
 *
 * Microsoft documents Quick Layout as one command of the shared Chart Design
 * tab in Word, PowerPoint and Excel ("Change the layout or style of a chart":
 * Chart Design ▸ Quick Layout, then pick a layout). A 2-D column or bar chart
 * takes Excel's eleven calibrated layouts, the ones the worksheet offers: each
 * switches the elements (title, axis titles, legend edge, gridlines, data
 * labels, data table) and carries the rest as `format.layout` (minor
 * gridlines, tick marks, gap width, overlap, per-point labels), which every
 * classic writer states the way Excel does. Any other family takes the first
 * five as element switches it can show; layouts 6-11 are column / bar ones.
 *
 * Applying a layout never discards a title or axis title the author typed (the
 * worksheet's rule): a layout that shows one where the chart has none adds
 * Office's placeholder text.
 */
import {
  chartHasAxis,
  chartSupportsAxisTitle,
  chartSupportsAxisVisibility,
  chartSupportsDataLabels,
  chartSupportsDataTable,
  chartSupportsGridlines,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
} from "../../components/console/editor/extensions/chartBlockShared";
import {
  OFFICE_CHART_QUICK_LAYOUT_IDS,
  officeChartQuickLayoutRecipe,
  officeChartTakesCalibratedLayouts,
  type OfficeChartQuickLayoutId,
} from "./chart-layout";

export { OFFICE_CHART_QUICK_LAYOUT_IDS, type OfficeChartQuickLayoutId };

/** What one of the five generic layouts switches on or off. */
export interface OfficeChartQuickLayoutPatch {
  readonly legend: "none" | "top" | "bottom";
  readonly axes: boolean;
  readonly gridlines: boolean;
  readonly dataLabels: boolean;
  readonly dataTable: boolean;
}

type GenericLayoutId = "layout-1" | "layout-2" | "layout-3" | "layout-4" | "layout-5";

const PATCHES: Readonly<Record<GenericLayoutId, OfficeChartQuickLayoutPatch>> =
  Object.freeze({
    "layout-1": Object.freeze({
      legend: "bottom",
      axes: true,
      gridlines: true,
      dataLabels: false,
      dataTable: false,
    }),
    "layout-2": Object.freeze({
      legend: "top",
      axes: true,
      gridlines: false,
      dataLabels: true,
      dataTable: false,
    }),
    "layout-3": Object.freeze({
      legend: "none",
      axes: true,
      gridlines: true,
      dataLabels: true,
      dataTable: false,
    }),
    "layout-4": Object.freeze({
      legend: "none",
      axes: true,
      gridlines: true,
      dataLabels: false,
      dataTable: true,
    }),
    "layout-5": Object.freeze({
      legend: "none",
      axes: true,
      gridlines: false,
      dataLabels: false,
      dataTable: false,
    }),
  });

/** The generic element switches of layouts 1-5 (for a non-column chart). */
export function officeChartQuickLayoutPatch(
  id: OfficeChartQuickLayoutId,
): OfficeChartQuickLayoutPatch | null {
  return (PATCHES as Record<string, OfficeChartQuickLayoutPatch>)[id] ?? null;
}

/** Whether this chart takes Excel's calibrated column / bar layouts. */
export function officeChartQuickLayoutCalibrated(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "chartSubtype">,
): boolean {
  return officeChartTakesCalibratedLayouts(attrs.chartType, attrs.chartSubtype);
}

/** Does this chart take this layout at all? */
export function officeChartQuickLayoutAvailable(
  attrs: Pick<ChartBlockAttrsV2, "chartType" | "chartSubtype">,
  id: OfficeChartQuickLayoutId,
): boolean {
  if (officeChartQuickLayoutCalibrated(attrs)) return true;
  const patch = officeChartQuickLayoutPatch(id);
  return Boolean(patch) && (!patch!.dataTable || chartSupportsDataTable(attrs.chartType));
}

/** The words a layout writes where it shows a title the chart does not have. */
export interface OfficeChartQuickLayoutWords {
  chartTitle: string;
  axisTitle: string;
}

const DEFAULT_WORDS: OfficeChartQuickLayoutWords = {
  chartTitle: "Chart Title",
  axisTitle: "Axis Title",
};

function withoutDataLabelFields(format: NonNullable<ChartBlockAttrsV2["format"]>) {
  // Whatever requires visible labels goes off with them; an imported chart
  // carries its part's own fields, which would otherwise refuse the layout.
  delete format.dataLabelPosition;
  delete format.dataLabelValue;
  delete format.dataLabelSeriesName;
  delete format.dataLabelCategoryName;
  delete format.dataLabelPercentage;
  delete format.dataLabelBubbleSize;
  delete format.dataLabelLegendKey;
  delete format.dataLabelSeparator;
}

/** The labels' own paint, which a chart keeps only while they show (a Chart
 * Style's label paint goes with labels a layout turns off). */
function withoutDataLabelPaint(format: NonNullable<ChartBlockAttrsV2["format"]>) {
  delete format.dataLabelStyle;
  delete format.dataLabelLeaderLineStyle;
}

function withCalibratedLayout(
  attrs: ChartBlockAttrsV2,
  id: OfficeChartQuickLayoutId,
  words: OfficeChartQuickLayoutWords,
): ChartBlockAttrsV2 | null {
  const categoryCount =
    attrs.data.kind === "category" ? attrs.data.categories.length : 1;
  const seriesCount =
    attrs.data.kind === "category" ? attrs.data.series.length : 1;
  const recipe = officeChartQuickLayoutRecipe(id, categoryCount, seriesCount);
  const next: ChartBlockAttrsV2 = { ...attrs };
  // A layout that shows a title adds Office's placeholder; one that hides it
  // keeps what the author typed.
  if (recipe.layout.title && !next.title) next.title = words.chartTitle;
  if (recipe.layout.axisTitleX && !next.axisTitleX) next.axisTitleX = words.axisTitle;
  if (recipe.layout.axisTitleY && !next.axisTitleY) next.axisTitleY = words.axisTitle;
  const format: NonNullable<ChartBlockAttrsV2["format"]> = {
    ...(attrs.format ?? {}),
    legend: recipe.legend,
    axisX: recipe.axisX,
    axisY: recipe.axisY,
    gridlinesX: recipe.gridlinesX,
    gridlinesY: recipe.gridlinesY,
    dataLabels: recipe.dataLabels,
    dataTable: recipe.dataTable,
    layout: {
      ...recipe.layout,
      // A stacked chart's series share one bar: Excel keeps its overlap at
      // 100 whatever the layout, or the stack would come apart.
      ...(attrs.chartSubtype === "stacked" ? { overlap: 100 } : {}),
      title: Boolean(next.title),
      axisTitleX: Boolean(next.axisTitleX),
      axisTitleY: Boolean(next.axisTitleY),
    },
  };
  withoutDataLabelFields(format);
  if (recipe.dataLabels) format.dataLabelValue = true;
  else withoutDataLabelPaint(format);
  delete format.dataTableLegendKeys;
  if (recipe.dataTable && !recipe.dataTableLegendKeys) format.dataTableLegendKeys = false;
  next.format = format;
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/**
 * The chart with the layout applied, or `null` when the chart cannot take it
 * or the result is not a valid chart. Only what the layout owns changes: an
 * element the chart type cannot show is neither written nor left behind.
 */
export function officeChartWithQuickLayout(
  placed: ChartBlockAttrsV2,
  id: OfficeChartQuickLayoutId,
  words: OfficeChartQuickLayoutWords = DEFAULT_WORDS,
): ChartBlockAttrsV2 | null {
  // A layout lays every element out again: what was dragged by hand goes
  // back to the chart's own places.
  const attrs = placed.format?.elementLayout
    ? (() => {
        const { elementLayout: _placed, ...format } = placed.format!;
        return { ...placed, format };
      })()
    : placed;
  if (!officeChartQuickLayoutAvailable(attrs, id)) return null;
  if (officeChartQuickLayoutCalibrated(attrs)) {
    return withCalibratedLayout(attrs, id, words);
  }
  const patch = officeChartQuickLayoutPatch(id)!;
  const format = { ...(attrs.format ?? {}), legend: patch.legend };
  delete format.layout;
  // One axis at a time: a funnel has its categories and no value axis, and a
  // layout that wrote `axisY` onto it refused the whole save.
  for (const axis of ["x", "y"] as const) {
    const key = axis === "x" ? "axisX" : "axisY";
    if (chartHasAxis(attrs.chartType, axis)) format[key] = patch.axes;
    else if (!chartSupportsAxisVisibility(attrs.chartType)) {
      delete format[key];
    }
  }
  if (chartSupportsGridlines(attrs.chartType)) {
    format.gridlinesX = false;
    format.gridlinesY = patch.gridlines;
  } else if (!chartSupportsAxisVisibility(attrs.chartType)) {
    // A funnel keeps the gridline keys its own part states; it is offered
    // none, and a layout does not rewrite what it cannot show.
    delete format.gridlinesX;
    delete format.gridlinesY;
  }
  if (chartSupportsDataLabels(attrs.chartType)) {
    format.dataLabels = patch.dataLabels;
  } else {
    delete format.dataLabels;
  }
  if (format.dataLabels !== true) {
    withoutDataLabelFields(format);
    withoutDataLabelPaint(format);
  }
  if (chartSupportsDataTable(attrs.chartType)) {
    format.dataTable = patch.dataTable;
  } else {
    delete format.dataTable;
  }
  const next: ChartBlockAttrsV2 = { ...attrs, format };
  // A generic layout keeps the axis titles a family can show.
  if (!chartSupportsAxisTitle(attrs.chartType, "x")) delete next.axisTitleX;
  if (!chartSupportsAxisTitle(attrs.chartType, "y")) delete next.axisTitleY;
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/** The layout the chart currently shows, if it is exactly one of the catalogue. */
export function officeChartActiveQuickLayout(
  attrs: ChartBlockAttrsV2,
): OfficeChartQuickLayoutId | null {
  const current = JSON.stringify(attrs.format ?? {});
  return (
    OFFICE_CHART_QUICK_LAYOUT_IDS.find((id) => {
      const next = officeChartWithQuickLayout(attrs, id);
      return (
        next !== null &&
        next.title === attrs.title &&
        JSON.stringify(next.format ?? {}) === current
      );
    }) ?? null
  );
}
