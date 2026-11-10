/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Design > Type > Change Chart Type on a Chart V2 block, for every
 * editor that holds one (Docs and the deck's owned V2 chart).
 *
 * Word, Excel and PowerPoint offer one Change Chart Type command listing every
 * chart type and its subtypes. The deck changed family AND subtype and retired
 * the layout atoms the new family cannot show; Docs changed the family only,
 * always to its first subtype, and kept every atom whether or not the new
 * family can show it, leaving the validator to refuse the result. One rule
 * now.
 *
 * The data is never replaced: a family whose data kind differs (category data
 * into a hierarchy, say) is refused rather than filled with a sample.
 */
import {
  chartDataLabelPositions,
  chartLabelOverridesForType,
  chartHasAxis,
  chartSupportsAxisTitle,
  chartSupportsAxisVisibility,
  chartHasDepthAxis,
  chartSupportsDataTable,
  chartSupportsGridlines,
  chartSupportsMinorGridlines,
  chartSupportsSeriesFilter,
  chartSupportsCategoryFilter,
  chartSupportsDataLabelCallout,
  chartSeriesStylesForType,
  defaultChartData,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartFamily,
  type ChartSubtype,
  chartTakesView3D,
  chartTakesBarSpacing,
  chartTakes3DGapWidth,
  chartTakesBarShape,
  chartTakesGapDepth,
  chartTakesAreaStyles,
  chartTakesSeriesStyles,
  chartHasSecondaryValueAxis,
  chartTakesBinIntervalClosed,
  chartTakesBoxWhisker,
  chartTakesExtendedGapWidth,
  chartTakesParentLabelLayout,
  COMBO_RING_RENDER_AS,
  comboBarAxisRefusal,
  type ComboRenderAs,
} from "../../components/console/editor/extensions/chartBlockShared";
import { officeChartTakesCalibratedLayouts } from "./chart-layout";

/** The element styles only a classic (`c:chartSpace`) chart's part states. */
const CLASSIC_ELEMENT_STYLE_KEYS = [
  "axisStyles",
  // An axis title's own paint and text (`ChartAttrsV2` takes it only on a
  // classic chart that draws axes): a workbook chart reads it from what its
  // part states (S-179), and left behind it refused every Office 2016 type.
  "axisTitleStyles",
  "gridlineStyle",
  "gridlineStyleX",
  "gridlineStyleY",
  "minorGridlineStyleX",
  "minorGridlineStyleY",
  "dataLabelStyle",
  "dataLabelLeaderLineStyle",
  "dataTableStyle",
  "trendlineStyle",
  "errorBarStyle",
  "errorBarXStyle",
  "upBarStyle",
  "downBarStyle",
  "dropLineStyle",
  "highLowLineStyle",
  "seriesLineStyle",
] as const;

/** The chart as that type, or `null` when its data cannot be that type. */
export function officeChartWithType(
  attrs: ChartBlockAttrsV2,
  family: ChartFamily,
  subtype: ChartSubtype,
): ChartBlockAttrsV2 | null {
  const sample = defaultChartData(family, subtype);
  if (sample.kind !== attrs.data.kind) return null;
  let data = attrs.data;
  // A pie draws one series; a doughnut draws each series as a ring, so a
  // two-series doughnut turned into columns can come back (PPT-B305).
  if (
    family === "pie" &&
    subtype !== "doughnut" &&
    data.kind === "category" &&
    data.series.length !== 1
  )
    return null;
  // A combo PRESET states how its series are drawn: Stacked Area - Clustered
  // Column is an area and columns, whatever the chart drew before. Each series
  // takes the preset's mapping (the first series the first one, the rest the
  // second), as the preset's own insert does; the kept column-and-line mapping
  // drew the old chart under the new name (PPT-B119). Custom keeps each
  // series' own choice.
  if (
    family === "combo" &&
    subtype !== "custom" &&
    data.kind === "category" &&
    sample.kind === "category"
  ) {
    const pattern = sample.series;
    data = {
      ...data,
      series: data.series.map((series, index) => {
        const preset = pattern[Math.min(index, pattern.length - 1)];
        const { renderAs: _renderAs, axis: _axis, ...rest } = series;
        return {
          ...rest,
          ...(preset?.renderAs !== undefined ? { renderAs: preset.renderAs } : {}),
          ...(preset?.axis !== undefined ? { axis: preset.axis } : {}),
        };
      }),
    };
  }
  // How a combination draws each series (`renderAs`, `axis`) is a combo's
  // alone: Change Chart Type to one type draws every series as that type, and
  // a column chart whose series still said "line, secondary" was refused by
  // the save for the whole page (`ChartAttrsV2`: "renderAs/axis are valid only
  // for combo charts").
  if (
    family !== "combo" &&
    data.kind === "category" &&
    data.series.some((series) => series.renderAs !== undefined || series.axis !== undefined)
  ) {
    data = {
      ...data,
      series: data.series.map(({ renderAs: _renderAs, axis: _axis, ...rest }) => rest),
    };
  }
  // A bubble's sizes are its third value; a scatter has nowhere to keep them
  // (`c:bubbleSize` is CT_BubbleSer's), and the save refused a scatter that
  // still carried them ("bubbleSizes are valid only for bubble charts").
  if (
    data.kind === "xy" &&
    subtype !== "bubble" &&
    subtype !== "bubble3d" &&
    data.series.some((series) => series.bubbleSizes !== undefined)
  ) {
    data = {
      ...data,
      series: data.series.map(({ bubbleSizes: _bubbleSizes, ...rest }) => rest),
    };
  }
  const next = { ...attrs, chartType: family, chartSubtype: subtype, data };
  // Layout atoms that the destination does not own are retired together.
  const format = { ...next.format };
  // An Office 2016 (chartEx) part keeps only its frame's paint (chart area,
  // plot area, title, legend); the element styles are a classic chart's, and
  // a series' own paint reaches a funnel at most. Office keeps what the new
  // type can carry and drops the rest; leaving them refused the switch here,
  // and a save of what got through refused the page.
  if (!chartTakesAreaStyles(family)) {
    for (const key of CLASSIC_ELEMENT_STYLE_KEYS) delete format[key];
  }
  if (!chartTakesSeriesStyles(family)) delete format.seriesStyles;
  // A classic axis' options (tick marks, crossing, units) have no place in
  // the chartEx part any host writes (`CHART_EX_PART_FORMAT_KEYS`); a workbook
  // chart reads the tick marks its part states (S-179), and every switch to
  // an Office 2016 type then declined them.
  if (officeChartPartKind(family) === "chartEx") delete format.axisOptions;
  // A depth axis's options belong to a 3-D chart that still has one
  // (`chartHasDepthAxis`): a clustered 3-D column has no series axis.
  if (format.axisOptions?.series && !chartHasDepthAxis(family, subtype)) {
    const { series: _series, ...axes } = format.axisOptions;
    if (Object.keys(axes).length) format.axisOptions = axes;
    else delete format.axisOptions;
  }
  // Error bars along the X axis are an X Y series' alone (Excel drops them
  // when the series becomes a line or a column).
  if (family !== "xy") {
    delete format.errorBarsX;
    delete format.errorBarXStyle;
  }
  // A secondary value axis' gridlines and title belong to a chart that still
  // draws one (`chartHasSecondaryValueAxis`).
  if (!chartHasSecondaryValueAxis(family, subtype, data)) {
    delete format.secondaryValueGridlines;
    delete format.secondaryValueAxisTitle;
    delete format.secondaryValueAxis;
  }
  // A depth axis title belongs to a chart that still has a depth axis
  // (`chartHasDepthAxis`: 3-D Column, Line, Area and the surfaces).
  if (!chartHasDepthAxis(family, subtype)) {
    delete format.seriesAxisTitle;
    if (format.elementLayout?.axisTitleSeries) {
      const { axisTitleSeries: _gone, ...rest } = format.elementLayout;
      void _gone;
      if (Object.keys(rest).length) format.elementLayout = rest;
      else delete format.elementLayout;
    }
  }
  if (!chartSupportsAxisVisibility(family)) {
    delete format.axisX;
    delete format.axisY;
    delete format.axisStyles;
    delete format.axisTitleStyles;
  } else if (!chartHasAxis(family, "y")) {
    delete format.axisY;
  }
  // Axis titles retire on their own table, one axis at a time: a radar has both
  // axes and no title Office will draw, and a funnel has one axis and one
  // title. Retiring them with the axes alone left an invalid chart behind.
  if (!chartSupportsAxisTitle(family, "x")) delete next.axisTitleX;
  if (!chartSupportsAxisTitle(family, "y")) delete next.axisTitleY;
  if (!chartSupportsGridlines(family)) {
    delete format.gridlinesX;
    delete format.gridlinesY;
    delete format.gridlineStyle;
    delete format.gridlineStyleX;
    delete format.gridlineStyleY;
  }
  if (!chartSupportsMinorGridlines(family, subtype)) {
    delete format.minorGridlinesX;
    delete format.minorGridlinesY;
    delete format.minorGridlineStyleX;
    delete format.minorGridlineStyleY;
  }
  // A type whose series are not independent plots shows every series again.
  if (!chartSupportsSeriesFilter(family)) delete format.hiddenSeries;
  if (!chartSupportsCategoryFilter(family)) delete format.hiddenCategories;
  if (!chartSupportsDataLabelCallout(family, subtype)) delete format.dataLabelCallout;
  // A 3-D pie's 3-D Rotation is its own; another type keeps Office's view.
  if (!chartTakesView3D(family, subtype)) {
    delete format.view3D;
    // So are its walls and floor (Format Back Wall / Side Wall / Floor).
    delete format.backWallStyle;
    delete format.sideWallStyle;
    delete format.floorStyle;
    delete format.wallThickness;
  }
  // A series' own marker and curve belong to a series that draws them: one
  // that becomes a column or an area loses them, as it does in Office.
  if (format.seriesStyles) {
    const kept = chartSeriesStylesForType(format.seriesStyles, {
      chartType: family,
      chartSubtype: subtype,
      data,
    });
    if (kept) format.seriesStyles = kept;
    else delete format.seriesStyles;
  }
  if (!chartSupportsDataTable(family)) {
    delete format.dataTable;
    delete format.dataTableStyle;
    // ... and what only a data table states: a No Legend Keys switch left
    // behind made the converted chart invalid, and the type was greyed as
    // one the data could not become (PPT-B296).
    delete format.dataTableLegendKeys;
  }
  // A calibrated Quick Layout is a column / bar chart's; another type keeps
  // its elements and drops the layout's bar geometry.
  if (format.layout !== undefined) {
    if (!officeChartTakesCalibratedLayouts(family, subtype)) delete format.layout;
    else if (subtype === "stacked") format.layout = { ...format.layout, overlap: 100 };
  }
  // A 100% stacked chart's own spacing is that chart's alone; another type
  // draws at its own (Excel drops it with the 100% stacking).
  if (!chartTakesBarSpacing(family, subtype)) {
    // A 3-D column or bar chart keeps its Gap Width (`c:bar3DChart` holds
    // one, and no Series Overlap).
    if (!chartTakes3DGapWidth(family, subtype)) delete format.barGapWidth;
    delete format.barOverlap;
  }
  // A 3-D chart's Gap Depth and Column shapes are that 3-D family's: a flat
  // chart has no depth to space and no column shape to draw.
  if (!chartTakesGapDepth(family, subtype)) delete format.gapDepth;
  if (!chartTakesBarShape(family, subtype)) {
    delete format.barShape;
    if (format.seriesStyles?.some((style) => style.barShape !== undefined)) {
      const kept = format.seriesStyles
        .map(({ barShape: _shape, ...rest }) => rest)
        .filter((style) => Object.keys(style).some((key) => key !== "seriesIndex"));
      if (kept.length) format.seriesStyles = kept;
      else delete format.seriesStyles;
    }
  }
  // A combination's stacked column group is that combination's alone.
  if (family !== "combo") delete format.comboColumnGrouping;
  // An Office 2016 family's Series Options are its own part's: a treemap's
  // Label Options (a sunburst names its branches on its rings), a box plot's
  // Quartile Calculation and points, a histogram's or pareto's bins, and the
  // Gap Width of the families whose columns have one. Left behind, the new
  // type was refused as one the data could not become.
  if (!chartTakesParentLabelLayout(family, subtype)) delete format.parentLabelLayout;
  if (!chartTakesBoxWhisker(family, subtype)) delete format.boxWhisker;
  if (!chartTakesBinIntervalClosed(family, subtype)) {
    delete format.binIntervalClosed;
    delete format.histogramBins;
  }
  if (!chartTakesExtendedGapWidth(family, subtype)) delete format.gapWidth;
  // A label position the new type does not offer goes back to where Office
  // puts labels itself (a column's Outside End means nothing on a funnel).
  if (
    format.dataLabelPosition !== undefined &&
    !chartDataLabelPositions(family, subtype).includes(format.dataLabelPosition)
  ) {
    delete format.dataLabelPosition;
  }
  // A series' or point's own label keeps its words and its dragged place
  // and gives up only what the new type cannot show.
  if (format.dataLabelOverrides !== undefined) {
    const labels = chartLabelOverridesForType(format.dataLabelOverrides, family, subtype);
    if (labels) format.dataLabelOverrides = labels;
    else delete format.dataLabelOverrides;
  }
  const changed: ChartBlockAttrsV2 = { ...next, format };
  return validateChartBlockAttrsV2(changed).valid ? changed : null;
}

/**
 * The Office 2016 chart families. In a file each is a `cx:chart` part (with a
 * classic `c:chart` twin in `mc:Fallback`); every other family is a classic
 * `c:chart` part (ECMA-376 21.2; [MS-ODRAWXML] 2.24 chartex).
 */
export const OFFICE_CHART_EX_FAMILIES: ReadonlySet<ChartFamily> = new Set([
  "hierarchy",
  "statistical",
  "waterfall",
  "funnel",
  "map",
]);

export type OfficeChartPartKind = "classic" | "chartEx";

export function officeChartPartKind(family: ChartFamily): OfficeChartPartKind {
  return OFFICE_CHART_EX_FAMILIES.has(family) ? "chartEx" : "classic";
}

/**
 * Why a host's writer cannot make the chart this type, or `null` when it can.
 *
 * Word, Excel and PowerPoint change a chart between a classic and an Office
 * 2016 type from one Change Chart Type list. A chart that is not yet a part of
 * a file (a Docs chart, a deck chart not yet saved) is written from its attrs
 * by the exporters, which write either kind. A chart that already IS a part of
 * the file is rewritten by that part's writer; Docs' and the deck's savers
 * replace a part of one kind with the other (`chart_kind_conversion.py`), so
 * they pass `null`. A host whose writer cannot turn one kind of part into the
 * other passes the kind of the part the chart is, and those types stay listed
 * and say why.
 */
export function officeChartTypeRefusal(
  attrs: ChartBlockAttrsV2,
  family: ChartFamily,
  subtype: ChartSubtype,
  filePartKind: OfficeChartPartKind | null,
): "data" | "partKind" | null {
  if (filePartKind !== null && officeChartPartKind(family) !== filePartKind) {
    return "partKind";
  }
  return officeChartWithType(attrs, family, subtype) ? null : "data";
}

/** One row of Change Chart Type > Combo > Custom Combination. */
export interface OfficeChartCombinationRow {
  name: string;
  renderAs: ComboRenderAs;
  secondary: boolean;
}

/**
 * The Custom Combination rows a chart opens with: a combo's own mapping, else
 * every series as the chart draws it now (a line chart's as lines, an area
 * chart's as areas, anything else as clustered columns) on its axis. Null when
 * the data is not category series or holds fewer than two: Office combines
 * 2-D category series only ("Create a combination chart": "you must use a
 * 2-D chart ... To combine charts, you should have more than one data series").
 */
export function officeChartCombinationRows(
  attrs: ChartBlockAttrsV2,
): OfficeChartCombinationRow[] | null {
  if (attrs.data.kind !== "category" || attrs.data.series.length < 2) return null;
  const drawn: OfficeChartCombinationRow["renderAs"] =
    attrs.chartType === "line" ? "line" : attrs.chartType === "area" ? "area" : "column";
  const count = attrs.data.series.length;
  return attrs.data.series.map((series, index) => {
    if (attrs.chartType === "combo") {
      const fallback =
        attrs.chartSubtype === "stackedAreaColumn"
          ? index < count - 1
            ? "area"
            : "column"
          : index === count - 1
            ? "line"
            : "column";
      return {
        name: series.name,
        renderAs: series.renderAs ?? fallback,
        secondary:
          series.axis === "secondary" ||
          (series.axis === undefined &&
            attrs.chartSubtype === "columnLineSecondary" &&
            index === count - 1),
      };
    }
    return { name: series.name, renderAs: drawn, secondary: series.axis === "secondary" };
  });
}

/**
 * Why a Custom Combination cannot be applied, or null. A combo plots at least
 * two series and keeps a primary axis: with every series on the secondary
 * axis there is nothing to scale the primary one.
 */
export function officeChartCombinationRefusal(
  rows: readonly OfficeChartCombinationRow[],
): "series" | "primary" | "rings" | "bars" | null {
  if (rows.length < 2) return "series";
  // A pie or doughnut ring has no axes to share: Excel combines the two
  // ring kinds only with each other, and such a combination has no primary
  // axis to keep.
  const rings = rows.filter((row) => COMBO_RING_RENDER_AS.has(row.renderAs)).length;
  if (rings) return rings === rows.length ? null : "rings";
  if (!rows.some((row) => !row.secondary)) return "primary";
  // Bars run their categories up the side: Excel combines them only with
  // series on the other axis group (`comboBarAxisRefusal`).
  return comboBarAxisRefusal(
    rows.map((row) => ({ renderAs: row.renderAs, axis: row.secondary ? "secondary" : "primary" })),
  );
}

/** The chart as Combo > Custom Combination with these rows, or null. */
export function officeChartWithCombination(
  attrs: ChartBlockAttrsV2,
  rows: readonly OfficeChartCombinationRow[],
): ChartBlockAttrsV2 | null {
  if (
    attrs.data.kind !== "category" ||
    rows.length !== attrs.data.series.length ||
    officeChartCombinationRefusal(rows) !== null
  )
    return null;
  const data = {
    ...attrs.data,
    series: attrs.data.series.map((series, index) => ({
      ...series,
      renderAs: rows[index].renderAs,
      axis: rows[index].secondary ? ("secondary" as const) : ("primary" as const),
    })),
  };
  return officeChartWithType({ ...attrs, data }, "combo", "custom");
}

/**
 * Format Data Series ▸ Series Options ▸ Plot Series On for one series: the
 * axis it is plotted on, or null when the chart offers no choice (a 3-D
 * chart, a pie, a chart Office cannot turn into a combination). A scatter's
 * and a bar chart's series name their axis themselves; a 2-D column, line
 * or area chart's become a Custom Combination, as Excel turns them into one
 * when a series moves to the secondary axis.
 */
export function chartSeriesPlottedOn(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
): "primary" | "secondary" | null {
  if (/3d/i.test(attrs.chartSubtype)) return null;
  if (attrs.data.kind === "xy" || (attrs.chartType === "bar" && attrs.data.kind === "category")) {
    const series = attrs.data.series[seriesIndex];
    return series ? (series.axis === "secondary" ? "secondary" : "primary") : null;
  }
  if (!["column", "line", "area", "combo"].includes(attrs.chartType)) return null;
  const rows = officeChartCombinationRows(attrs);
  const row = rows?.[seriesIndex];
  if (!row || COMBO_RING_RENDER_AS.has(row.renderAs)) return null;
  return row.secondary ? "secondary" : "primary";
}

/** The chart with one series plotted on `axis`, or null when it offers no
 * choice, nothing changes, or the result is not one Office draws (every
 * series on the secondary axis). */
export function chartWithSeriesAxis(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  axis: "primary" | "secondary",
): ChartBlockAttrsV2 | null {
  const current = chartSeriesPlottedOn(attrs, seriesIndex);
  if (current === null || current === axis) return null;
  if (attrs.data.kind === "xy" || attrs.chartType === "bar") {
    if (attrs.data.kind !== "xy" && attrs.data.kind !== "category") return null;
    const series = attrs.data.series.map((entry, index) =>
      index === seriesIndex ? { ...entry, axis } : entry,
    );
    if (!series.some((entry) => entry.axis !== "secondary")) return null;
    const next = { ...attrs, data: { ...attrs.data, series } } as ChartBlockAttrsV2;
    return validateChartBlockAttrsV2(next).valid ? next : null;
  }
  const rows = officeChartCombinationRows(attrs);
  if (!rows) return null;
  return officeChartWithCombination(
    attrs,
    rows.map((row, index) => (index === seriesIndex ? { ...row, secondary: axis === "secondary" } : row)),
  );
}
