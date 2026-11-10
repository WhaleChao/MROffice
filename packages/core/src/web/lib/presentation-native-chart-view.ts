/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";
import {
  chartHasAxis,
  chartSupportsGridlines,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartData,
} from "../components/console/editor/extensions/chartBlockShared";
import { readPresentationChartV2 } from "./presentation-chart-v2";
import {
  normalizeSpreadsheetChartStyle,
  spreadsheetChartStylePreset,
} from "./spreadsheet/chart";

const record = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
const same = (left: unknown, right: unknown) =>
  JSON.stringify(left) === JSON.stringify(right);

function updateData(
  data: ChartData,
  block: PresentationChartBlock,
): ChartData | null {
  const rows = block.data?.rows;
  const headers = block.data?.headers;
  if (!rows?.length || !headers || rows.length > 50 || headers.length < 2)
    return null;
  const names = headers.slice(1);
  if (rows.some((row) => row.length !== names.length + 1)) return null;
  const categories = rows.map((row) => String(row[0]));
  const values = names.map((_, index) =>
    rows.map((row) => Number(row[index + 1])),
  );
  if (values.some((series) => series.some((value) => !Number.isFinite(value))))
    return null;
  // An Office 2016 chart's grid, read back exactly as the server reads it
  // (`native_chart_view.chart_ex_data_from_grid`): a statistical chart's
  // samples are one column per series, a waterfall's and a funnel's rows are
  // its categories and values, and a map's regions keep the ids the part
  // holds, because only Office's geocoder can name them.
  if (data.kind === "sample") {
    if (
      values.length !== data.series.length ||
      data.series.some((series, index) => series.values.length !== values[index].length)
    )
      return null;
    return {
      ...data,
      series: data.series.map((series, index) => ({
        ...series,
        values: values[index],
      })),
    };
  }
  if (data.kind === "waterfall") {
    if (values.length !== 1 || values[0].length !== data.values.length) return null;
    return { ...data, categories, values: values[0] };
  }
  if (data.kind === "hierarchy") {
    // Only the leaf values are read back: the path column names rows the writer
    // will not let move. An interior value is stored nowhere and is derived as
    // the sum of the rows beneath it, the rule the grid calculators keep too.
    const points = data.points.map((point) => ({ ...point }));
    const parents = new Set(
      points.flatMap((point) => (point.parentId ? [point.parentId] : [])),
    );
    const leaves = points.filter((point) => !parents.has(point.id));
    if (values.length !== 1 || values[0].length !== leaves.length) return null;
    leaves.forEach((leaf, index) => {
      leaf.value = values[0][index];
    });
    const children = new Map<string, typeof points>();
    for (const point of points) {
      const key = point.parentId ?? "";
      children.set(key, [...(children.get(key) ?? []), point]);
    }
    const total = (point: (typeof points)[number]): number => {
      const kids = children.get(point.id) ?? [];
      if (kids.length)
        point.value = kids.reduce((sum, kid) => sum + total(kid), 0);
      return point.value;
    };
    for (const point of points) if (!point.parentId) total(point);
    return { ...data, points };
  }
  if (data.kind === "map") {
    if (values.length !== 1 || values[0].length !== data.regions.length) return null;
    return {
      ...data,
      regions: data.regions.map((region, index) => ({
        ...region,
        label: categories[index],
        value: values[0][index],
      })),
    };
  }
  if (data.kind === "category" || data.kind === "stock") {
    if (data.series.length !== names.length) return null;
    return {
      ...data,
      categories,
      series: data.series.map((series, i) => ({
        ...series,
        name: names[i],
        values: values[i],
      })),
    } as ChartData;
  }
  if (data.kind === "matrix") {
    if (data.yCategories.length !== names.length) return null;
    return { ...data, xCategories: categories, yCategories: names, values };
  }
  // The legacy grid exposes only Y for XY charts. Do not reinterpret row
  // labels as X, or discard native bubble sizes, when updating those values.
  if (
    data.kind === "xy" &&
    data.series.length === names.length &&
    data.series.every((series) => series.yValues.length === rows.length)
  ) {
    return {
      ...data,
      series: data.series.map((series, i) => ({
        ...series,
        name: names[i],
        yValues: values[i],
      })),
    };
  }
  return null;
}

/**
 * The controls this composition can give up, the most consequential first.
 *
 * The twin of `native_chart_view.VIEW_EDIT_CONTROLS` on the server: when a set
 * of changes cannot be composed, the one given up is sought from the end, so
 * formatting yields before the switch it depends on.
 *
 * The data is deliberately not on this list. The server may give it up because
 * it reports what it gave up, and the status bar names it; the canvas has no
 * such surface, so drawing the part's own numbers under a grid the reader has
 * already changed would be a silent divergence. A data change this composition
 * cannot carry still refuses the whole composition.
 */
const CHART_VIEW_EDIT_CONTROLS = [
  "title",
  "axisTitleX",
  "axisTitleY",
  "legend",
  "showAxisLabels",
  "showGridlines",
  "showDataLabels",
  "showDataTable",
  "trendline",
  "chartStyle",
  "chartAreaStyle",
  "plotAreaStyle",
  "chartTitleStyle",
  "legendStyle",
  "seriesStyles",
  "sliceStyles",
  "axisStyles",
  "gridlineStyle",
  "dataLabelStyle",
  "dataTableStyle",
  "trendlineStyle",
] as const;

/**
 * The finer Add Chart Element picks an imported chart carries beside the
 * block's coarse switches — a partial Chart V2 format under `chartElements`,
 * `null` clearing a key the part states. The twin of
 * `native_chart_view.VIEW_ELEMENT_KEYS`; each key is its own control, applied
 * last and given up one at a time like the rest.
 */
export const CHART_VIEW_ELEMENT_KEYS = [
  "axisX",
  "axisY",
  "gridlinesX",
  "gridlinesY",
  // Primary minor gridlines on a chart that takes no calibrated layout.
  "minorGridlinesX",
  "minorGridlinesY",
  "dataLabels",
  "dataLabelPosition",
  // A filled map's Map labels pick: the tri-state the map has instead of a
  // position, written as `cx:regionLabelLayout`.
  "dataLabelRegionLayout",
  "dataTable",
  "dataTableLegendKeys",
  "trendline",
  "errorBars",
  "upDownBars",
  "dropLines",
  "highLowLines",
  "seriesLines",
  // Change Colors on an Office 2016 chart: a palette of the shared table,
  // written as the chart's Chart Colors part.
  "colorPalette",
  // Format Axis > Axis Options: the shared record, written on the part's own
  // axes by the shared classic editor's layout patcher.
  "axisOptions",
  // Chart Filters: the hidden series and categories, stated in the part as
  // Office states them by the shared classic editor.
  "hiddenSeries",
  "hiddenCategories",
  // Data Labels > Data Callout, the labels' callout geometry.
  "dataLabelCallout",
  // The chart's own text (Home ▸ Font on the selected chart), written as
  // `c:chartSpace/c:txPr` by the shared classic editor.
  "text",
] as const;
export type ChartViewElementKey = (typeof CHART_VIEW_ELEMENT_KEYS)[number];
type ChartViewElementControl = `element:${ChartViewElementKey}`;

type ChartViewEditControl =
  | (typeof CHART_VIEW_EDIT_CONTROLS)[number]
  | ChartViewElementControl;

function chartElementValue(
  source: Record<string, unknown> | null | undefined,
  key: ChartViewElementKey,
): unknown {
  const elements = record(source?.chartElements);
  // `undefined` is "the block does not state this key"; `null` clears it.
  return elements && key in elements ? elements[key] : undefined;
}

/** The block keys each control reads; every other key is outside this composition. */
function chartViewEditControlKeys(
  control: (typeof CHART_VIEW_EDIT_CONTROLS)[number],
): readonly string[] {
  return control === "legend" ? ["showLegend", "legendPosition"] : [control];
}

/** The style keys the composition copies straight from the block. */
const CHART_VIEW_EDIT_STYLE_CONTROLS = [
  "chartAreaStyle",
  "plotAreaStyle",
  "chartTitleStyle",
  "legendStyle",
  "seriesStyles",
  "sliceStyles",
  "axisStyles",
  "gridlineStyle",
  "dataLabelStyle",
  "dataTableStyle",
  "trendline",
  "trendlineStyle",
] as const satisfies readonly ChartViewEditControl[];

function composeChartRenderAttrs(
  block: PresentationChartBlock,
  controls: ReadonlySet<ChartViewEditControl> | null,
): ChartBlockAttrsV2 | null {
  const composes = (control: ChartViewEditControl) =>
    controls === null || controls.has(control);
  const source = record(block.source),
    view = record(source?.nativeChartView),
    baseline = record(view?.baseline);
  if (
    !view ||
    !baseline ||
    view.schemaVersion !== "pptx-native-chart-view-v1" ||
    !/^ppt\/(?:[A-Za-z0-9_-]+\/)*charts\/[A-Za-z0-9_-]+\.xml$/.test(
      String(view.chartPart),
    ) ||
    !/^[a-f0-9]{64}$/.test(String(view.chartFingerprint)) ||
    !validateChartBlockAttrsV2(view.attrs).valid ||
    block.chartType !== baseline.chartType
  )
    return null;
  const native = view.attrs as ChartBlockAttrsV2;
  const data = same(block.data, baseline.data)
    ? native.data
    : updateData(native.data, block);
  if (!data) return null;
  const next: ChartBlockAttrsV2 = {
    ...native,
    data,
    format: { ...native.format },
  };
  for (const key of ["title", "axisTitleX", "axisTitleY"] as const) {
    if (composes(key) && !same(block[key], baseline[key]))
      next[key] = block[key] ?? "";
  }
  const format = next.format as Record<string, unknown>;
  const current = block as unknown as Record<string, unknown>;
  for (const key of CHART_VIEW_EDIT_STYLE_CONTROLS) {
    if (composes(key) && !same(current[key], baseline[key])) {
      if (current[key] === undefined) delete format[key];
      else format[key] = current[key];
    }
  }
  if (
    composes("legend") &&
    (!same(block.showLegend, baseline.showLegend) ||
      !same(block.legendPosition, baseline.legendPosition))
  ) {
    // A position cleared, or a legend the part lacks, sits at the right, where
    // the ribbon's Legend Position shows it; otherwise the part's own stays.
    const placed = native.format?.legend;
    format.legend =
      block.showLegend === false
        ? "none"
        : (block.legendPosition ??
          (typeof baseline.legendPosition === "string" || placed === "none"
            ? "right"
            : (placed ?? "bottom")));
  }
  // The twin of `native_chart_view_edited_attrs`: the block's one Axes switch
  // reaches only the axes the family has (a funnel has no value axis), and its
  // one Gridlines switch only a family that takes gridlines. Composing either
  // onto a chart without them made one the writer refuses at write time.
  if (
    composes("showAxisLabels") &&
    !same(block.showAxisLabels, baseline.showAxisLabels)
  ) {
    const axes = (["x", "y"] as const).filter((axis) =>
      chartHasAxis(native.chartType, axis),
    );
    if (axes.length === 0) return null;
    for (const axis of axes)
      format[axis === "x" ? "axisX" : "axisY"] = block.showAxisLabels !== false;
  }
  if (
    composes("showGridlines") &&
    !same(block.showGridlines, baseline.showGridlines)
  ) {
    if (!chartSupportsGridlines(native.chartType)) return null;
    format.gridlinesY = block.showGridlines !== false;
  }
  if (
    composes("showDataLabels") &&
    !same(block.showDataLabels, baseline.showDataLabels)
  ) {
    format.dataLabels = block.showDataLabels === true;
    if (!format.dataLabels) {
      // The twin of `native_chart_view_edited_attrs`: a position and the
      // content fields each require visible labels, so switching the labels
      // off clears them. Left in place the edited chart is unrepresentable and
      // the edit is dropped — which reads as a dead switch.
      delete format.dataLabelPosition;
      delete format.dataLabelValue;
      delete format.dataLabelSeriesName;
      delete format.dataLabelCategoryName;
      delete format.dataLabelSeparator;
      delete format.dataLabelRegionLayout;
    }
  }
  if (
    composes("showDataTable") &&
    !same(block.showDataTable, baseline.showDataTable)
  )
    format.dataTable = block.showDataTable === true;
  if (current.chartElements !== undefined) {
    const elements = record(current.chartElements);
    if (
      !elements ||
      Object.keys(elements).some(
        (key) => !(CHART_VIEW_ELEMENT_KEYS as readonly string[]).includes(key),
      )
    )
      return null;
    for (const key of CHART_VIEW_ELEMENT_KEYS) {
      const value = chartElementValue(current, key);
      if (
        value === undefined ||
        !composes(`element:${key}`) ||
        same(value, chartElementValue(baseline, key))
      )
        continue;
      if (
        (key === "axisX" || key === "axisY") &&
        !chartHasAxis(native.chartType, key === "axisX" ? "x" : "y")
      )
        return null;
      if (
        (key === "gridlinesX" || key === "gridlinesY") &&
        !chartSupportsGridlines(native.chartType)
      )
        return null;
      if (value === null) delete format[key];
      else format[key] = value;
    }
    if (format.dataLabels !== true) {
      delete format.dataLabelPosition;
      delete format.dataLabelRegionLayout;
    }
    if (format.dataTable !== true) delete format.dataTableLegendKeys;
  }
  // Keep the existing style gallery's palette preview for inherited paints.
  // Explicit native fills/gradients continue to override the numeric style,
  // just as they do after the native writer updates only c:style.
  if (
    composes("chartStyle") &&
    !same(block.chartStyle, baseline.chartStyle) &&
    normalizeSpreadsheetChartStyle(block.chartStyle)
  ) {
    const palette = spreadsheetChartStylePreset(block.chartStyle).palette;
    if (native.chartType === "pie" && data.kind === "category") {
      next.format!.sliceStyles = data.categories.map((_, pointIndex) => {
        const style = next.format?.sliceStyles?.find(
          (item) => item.pointIndex === pointIndex,
        );
        return {
          pointIndex,
          ...(style?.fill || style?.shapePaint?.fillGradient
            ? {}
            : { fill: palette[pointIndex % palette.length] }),
          ...style,
        };
      });
    } else if ("series" in data) {
      next.format!.seriesStyles = data.series.map((_, seriesIndex) => {
        const style = next.format?.seriesStyles?.find(
          (item) => item.seriesIndex === seriesIndex,
        );
        return {
          seriesIndex,
          ...(style?.fill || style?.shapePaint?.fillGradient
            ? {}
            : { fill: palette[seriesIndex % palette.length] }),
          ...style,
        };
      });
    }
  }
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/**
 * Render-only adapter. This never grants native Chart V2 writer ownership.
 *
 * Every change is composed together first; only when that is refused is a
 * control given up, the one whose removal lets the rest compose, sought from
 * the end of `CHART_VIEW_EDIT_CONTROLS`. One control the chart cannot take --
 * data labels on a family that has none, gridlines on a pie -- used to leave
 * no chart at all: the canvas drew nothing and every control went dead until
 * the pick was undone. The server gave up one control instead of the whole
 * chart in `native_chart_view_edit`; this is its twin, so what the canvas
 * draws is what the save writes.
 */
export function readPresentationChartRenderAttrs(
  block: PresentationChartBlock,
): ChartBlockAttrsV2 | null {
  if (block.chartV2 !== undefined)
    return readPresentationChartV2(block.chartV2);
  if (composeChartRenderAttrs(block, new Set()) === null) return null;
  const baseline = record(record(record(block.source)?.nativeChartView)?.baseline);
  const current = block as unknown as Record<string, unknown>;
  const applied: ChartViewEditControl[] = [
    ...CHART_VIEW_EDIT_CONTROLS.filter((control) =>
      chartViewEditControlKeys(control).some(
        (key) => !same(current[key], baseline?.[key]),
      ),
    ),
    ...CHART_VIEW_ELEMENT_KEYS.filter(
      (key) =>
        !same(chartElementValue(current, key), chartElementValue(baseline, key)),
    ).map((key) => `element:${key}` as const),
  ];
  let attrs = composeChartRenderAttrs(block, new Set(applied));
  while (attrs === null && applied.length > 0) {
    const reversed = [...applied].reverse();
    const culprit =
      reversed.find(
        (control) =>
          composeChartRenderAttrs(
            block,
            new Set(applied.filter((item) => item !== control)),
          ) !== null,
      ) ??
      reversed.find(
        (control) => composeChartRenderAttrs(block, new Set([control])) === null,
      ) ??
      applied[applied.length - 1];
    applied.splice(applied.indexOf(culprit), 1);
    attrs = composeChartRenderAttrs(block, new Set(applied));
  }
  return attrs;
}
