/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The name a chart tip gives the element under the pointer, as Office's
 * chart tips do ("Show chart element names on hover", Excel Options >
 * Advanced; `Application.ShowChartTipNames`, on by default): the chart area,
 * the plot area, the chart title, an axis or its title, the legend or one of
 * its entries, a series' data labels, a series, and on a plotted point the
 * point's name and value ("Show data point values on hover",
 * `ShowChartTipValues`). Office names the parts the way its Current
 * Selection list does: Series "Sales" Point "Q2" Value: 18,
 * Vertical (Value) Axis, Horizontal (Category) Axis Title, ...
 *
 * Pure: the picture says what the pointer is over (`OfficeChartTipTarget`),
 * this module says which name that is, and the host's catalogue words it.
 */
import { officeChartAxisRole, type OfficeChartPicturePart } from "./chart-hit";

/** What the pointer is over, as the picture's hit test finds it. */
export interface OfficeChartTipTarget {
  /** A marked part of the picture. */
  readonly part?: OfficeChartPicturePart | "chartArea";
  /** On the legend: the entry's words. */
  readonly legendEntry?: string;
  /** On a data label: its series' name. */
  readonly labelSeries?: string;
  /** On a plotted point: its series' name, the point's name and value. */
  readonly point?: { readonly series: string; readonly name?: string; readonly value?: string };
}

/** Office's four axis names (an XY chart's two axes are both value axes). */
export type OfficeChartAxisName = "horizontalCategory" | "horizontalValue" | "verticalCategory" | "verticalValue";

export type OfficeChartElementTip =
  | { readonly key: "chartArea" | "plotArea" | "chartTitle" | "legend" }
  | { readonly key: "legendEntry"; readonly name: string; readonly byPoint: boolean }
  | { readonly key: "axis" | "axisTitle"; readonly axis: OfficeChartAxisName }
  | { readonly key: "dataLabels"; readonly series: string }
  | { readonly key: "series"; readonly series: string }
  | { readonly key: "point"; readonly series: string; readonly point: string; readonly value?: string };

/** The families whose legend lists their points (a pie's slices), not series. */
const LEGEND_BY_POINT = new Set(["pie", "funnel", "treemap", "sunburst", "hierarchy"]);
/** Whether the legend of a chart of this family lists its points (a pie's
 * slices) rather than its series. */
export function officeChartLegendListsPoints(chartType: string): boolean {
  return LEGEND_BY_POINT.has(chartType);
}

/** The families both of whose axes are value axes. */
const VALUE_BY_VALUE = new Set(["xy", "bubble"]);

/** The planner's mark on a repeated series name (`withDistinctSeriesNames`). */
const SAME_NAME_MARK = "​";
const shown = (name: string) => name.split(SAME_NAME_MARK).join("");

/** Office's name of the axis drawn along the bottom (X) or up the side (Y). */
export function officeChartAxisName(chartType: string, physical: "X" | "Y"): OfficeChartAxisName {
  const role = VALUE_BY_VALUE.has(chartType) ? "value" : officeChartAxisRole({ chartType }, physical);
  if (physical === "X") return role === "category" ? "horizontalCategory" : "horizontalValue";
  return role === "category" ? "verticalCategory" : "verticalValue";
}

/** The tip for what the pointer is over, or null where Office shows none. */
export function officeChartElementTip(chartType: string, target: OfficeChartTipTarget): OfficeChartElementTip | null {
  if (target.point) {
    const series = shown(target.point.series);
    if (target.point.name === undefined) return { key: "series", series };
    return {
      key: "point",
      series,
      point: shown(target.point.name),
      ...(target.point.value !== undefined ? { value: target.point.value } : {}),
    };
  }
  switch (target.part) {
    case "chartArea":
      return { key: "chartArea" };
    case "plot":
      return { key: "plotArea" };
    case "title":
      return { key: "chartTitle" };
    case "legend":
      return target.legendEntry !== undefined
        ? { key: "legendEntry", name: shown(target.legendEntry), byPoint: LEGEND_BY_POINT.has(chartType) }
        : { key: "legend" };
    case "axisX":
    case "axisY":
      return { key: "axisTitle", axis: officeChartAxisName(chartType, target.part === "axisX" ? "X" : "Y") };
    case "axisLabelX":
    case "axisLabelY":
      return { key: "axis", axis: officeChartAxisName(chartType, target.part === "axisLabelX" ? "X" : "Y") };
    case "dataLabels":
      return target.labelSeries !== undefined ? { key: "dataLabels", series: shown(target.labelSeries) } : null;
    default:
      return null;
  }
}

/** A stable identity of a tip (the tip stays put while it names the same element). */
export function officeChartElementTipKey(tip: OfficeChartElementTip | null): string {
  return tip ? JSON.stringify(tip) : "";
}
