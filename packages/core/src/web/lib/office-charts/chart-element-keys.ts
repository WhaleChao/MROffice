/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The keys a chart with an element selected takes, as Office's do.
 *
 * Office's arrow keys never nudge a chart element (a title, the legend, the
 * plot area moves only by its mouse drag, or by its Format pane): with an
 * element of a chart selected they select another element. Up and Down step
 * through the chart's element groups, Left and Right through the members of
 * a group -- one series to the next, one point of a series to the next
 * (Excel's chart keys: "Select the next / previous group of elements",
 * "Select the next / previous element within a group"). Esc lets the element
 * go: the chart itself stays selected (its chart area), and the next Esc is
 * the host's (the chart let go).
 *
 * Hosts call `officeChartAreaAfterKey` from their own key handling while one
 * of their charts has an element selected (Current Selection is not the chart
 * area) and nothing is being typed, and keep their own meaning of a key it
 * leaves alone (`undefined`).
 */
import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";

import { chartDataLabelsShown, chartElementAvailability, chartLegendShown } from "./chart-elements";
import { officeChartAxisRole, officeChartPointArea } from "./chart-hit";
import { officeChartSelectedArea, setOfficeChartSelectedArea } from "./chart-selection-store";

/** The families whose points the pane addresses one by one (`chart-hit`). */
const POINT_FAMILIES = new Set(["column", "bar", "line", "area", "radar", "combo", "xy"]);

function seriesCount(attrs: ChartBlockAttrsV2): number {
  const data = attrs.data as { series?: readonly unknown[] };
  return Array.isArray(data.series) ? data.series.length : 0;
}

function pointCount(attrs: ChartBlockAttrsV2, seriesIndex: number): number {
  if (attrs.data.kind === "category") return attrs.data.categories.length;
  if (attrs.data.kind === "xy") return attrs.data.series[seriesIndex]?.xValues.length ?? 0;
  return 0;
}

/**
 * The chart's element groups in the order the arrow keys step through them:
 * the chart area (null), the title, the plot area, the value axis and its
 * title, the category axis and its title, the legend, the data labels, then
 * the series. Only the elements this chart draws are listed.
 */
export function officeChartKeyboardAreas(attrs: ChartBlockAttrsV2): (string | null)[] {
  const areas: (string | null)[] = [null];
  if (attrs.title) areas.push("chartTitleStyle");
  const available = chartElementAvailability(attrs);
  const hasAxes = available.axisX || available.axisY;
  if (hasAxes) areas.push("plotAreaStyle");
  const xRole = officeChartAxisRole(attrs, "X");
  for (const role of ["value", "category"] as const) {
    const physical = xRole === role ? "X" : "Y";
    const shown = physical === "X" ? available.axisX && attrs.format?.axisX !== false : available.axisY && attrs.format?.axisY !== false;
    if (shown) areas.push(`axis:${role}`);
    if (physical === "X" ? attrs.axisTitleX : attrs.axisTitleY) areas.push(`axisTitle:${role}`);
  }
  if (chartLegendShown(attrs) !== "none") areas.push("legendStyle");
  if (available.dataLabels && chartDataLabelsShown(attrs)) areas.push("dataLabelStyle");
  const hidden = new Set(attrs.format?.hiddenSeries ?? []);
  for (let index = 0; index < seriesCount(attrs); index += 1) if (!hidden.has(index)) areas.push(`series:${index}`);
  return areas;
}

/**
 * Current Selection after a key: the area to select (null: the chart area,
 * the element let go), or undefined for a key this leaves to the host.
 * `area` is the selected element's (never the chart area: a chart selected
 * as a whole keeps the host's meaning of every key).
 */
export function officeChartAreaAfterKey(
  attrs: ChartBlockAttrsV2,
  area: string | null,
  key: string,
): string | null | undefined {
  if (!area || area === "chartAreaStyle") return undefined;
  if (key === "Escape") return null;
  if (key !== "ArrowUp" && key !== "ArrowDown" && key !== "ArrowLeft" && key !== "ArrowRight") return undefined;
  const forward = key === "ArrowUp" || key === "ArrowRight";
  const step = (index: number, count: number) => (index + (forward ? 1 : -1) + count) % count;
  const point = officeChartPointArea(area);
  if (point && (key === "ArrowLeft" || key === "ArrowRight") && POINT_FAMILIES.has(attrs.chartType)) {
    // One point of a series selected: its neighbour in the series.
    const count = pointCount(attrs, point.seriesIndex);
    if (count > 0) return `point:${point.seriesIndex}:${step(point.pointIndex, count)}`;
  }
  const series = /^series:(\d+)$/.exec(area);
  const areas = officeChartKeyboardAreas(attrs);
  if ((series || point) && (key === "ArrowLeft" || key === "ArrowRight")) {
    // A series: the next series of the chart.
    const listed = areas.filter((entry): entry is string => entry !== null && entry.startsWith("series:"));
    const current = `series:${series ? series[1] : point!.seriesIndex}`;
    const at = listed.indexOf(current);
    if (listed.length) return listed[at < 0 ? 0 : step(at, listed.length)];
  }
  // A group: the next group (a series or a point stands in the series group).
  const groups = areas.filter((entry) => entry === null || !entry.startsWith("series:"));
  const hasSeries = areas.length > groups.length;
  if (hasSeries) groups.push("series:");
  const group = series || point ? "series:" : area;
  const at = groups.indexOf(group);
  const next = groups[at < 0 ? 0 : step(at, groups.length)];
  if (next === "series:") return areas.find((entry) => entry?.startsWith("series:")) ?? null;
  return next;
}

/**
 * What the picture marks for an area selected by a key: a series' points or
 * one point, in the picture's own indices (the plotted series and points,
 * hidden ones left out), as a click on them marks them.
 */
export function officeChartDetailOfArea(
  attrs: ChartBlockAttrsV2,
  area: string | null,
): { points: { series: number; point?: number } } | null {
  if (!area) return null;
  const point = officeChartPointArea(area);
  const series = point ? point.seriesIndex : Number(/^series:(\d+)$/.exec(area)?.[1] ?? Number.NaN);
  if (!Number.isInteger(series)) return null;
  const hiddenSeries = new Set(attrs.format?.hiddenSeries ?? []);
  if (hiddenSeries.has(series)) return null;
  const drawnSeries = Array.from({ length: series }, (_, index) => index).filter((index) => !hiddenSeries.has(index)).length;
  if (!point) return { points: { series: drawnSeries } };
  const hiddenPoints = new Set(attrs.data.kind === "category" ? (attrs.format?.hiddenCategories ?? []) : []);
  if (hiddenPoints.has(point.pointIndex)) return null;
  const drawnPoint = Array.from({ length: point.pointIndex }, (_, index) => index).filter((index) => !hiddenPoints.has(index)).length;
  return { points: { series: drawnSeries, point: drawnPoint } };
}

/**
 * A host's key handling for its chart with an element selected: the key
 * moves Current Selection (`officeChartAreaAfterKey`) and true is returned
 * (the host consumes the key); false for a key that is the host's own.
 */
export function selectOfficeChartAreaByKey(
  selectionKey: string | null | undefined,
  attrs: ChartBlockAttrsV2,
  key: string,
): boolean {
  if (!selectionKey) return false;
  const next = officeChartAreaAfterKey(attrs, officeChartSelectedArea(selectionKey), key);
  if (next === undefined) return false;
  setOfficeChartSelectedArea(selectionKey, next, officeChartDetailOfArea(attrs, next));
  return true;
}
