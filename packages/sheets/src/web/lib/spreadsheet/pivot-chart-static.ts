/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PivotChart that loses its PivotTable becomes a standard chart (Microsoft
 * Support, "Delete a PivotTable": the PivotChart turns into a standard chart
 * you can no longer pivot or update). Excel keeps what the chart last showed
 * as the chart's own values; so does this: the chart's model carries the
 * report's last reading and its binding says the data is its own.
 */

import type { ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { normalizeSpreadsheetChartBinding, type MiniChartXYSeries } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import {
  spreadsheetChartV2Reading,
  withSpreadsheetChartV2,
  type SpreadsheetChartLike,
} from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";

/**
 * The standard chart a PivotChart becomes: no PivotTable binding or field
 * buttons, its data the rows it last plotted (`rows`, header row first, as
 * `pivotChartReportToRows` gives them), its kind, titles and format kept.
 */
export function staticChartFromPivotChart<T extends SpreadsheetChartLike>(
  chart: T,
  rows: readonly string[][],
): T {
  const rest: Record<string, unknown> = { ...chart };
  delete rest.pivotSource;
  delete rest.pivotOptions;
  // An imported part's cell references were the report's; a standard chart
  // of its own values names none.
  delete rest.ranges;
  delete rest.seriesRefs;
  const unbound = rest as T;
  const current = chart.chartV2 as ChartBlockAttrsV2 | undefined;
  const reading =
    rows.length > 1
      ? spreadsheetChartV2Reading(unbound, { rows: rows.map((row) => [...row]), headerRow: true })
      : null;
  const attrs =
    current && reading && reading.data.kind === current.data.kind
      ? { ...current, data: reading.data }
      : (reading ?? current);
  if (!attrs) return unbound;
  return withSpreadsheetChartV2(unbound, attrs, { literal: true }) as unknown as T;
}

/**
 * The rows a chart whose data is its own plots (`binding.literal`), header
 * row first; null for a chart bound to cells.
 */
export function spreadsheetChartLiteralRows(
  chart: SpreadsheetChartLike,
): string[][] | null {
  if (!normalizeSpreadsheetChartBinding(chart.binding)?.literal) return null;
  return spreadsheetChartCachedRows(chart);
}

/**
 * The rows the chart's model last read, header row first, whatever its
 * binding: what a PivotChart whose PivotTable draws no report plots (a
 * preserve-only one, an OLAP cube's), as Excel shows the values the file
 * cached (X-235). R480-X04: an imported PivotChart bound to its report's
 * cells (no `binding.literal`) on a preserve-only PivotTable plotted
 * nothing and printed "no visible chart data". Null for a model with no
 * category data.
 */
export function spreadsheetChartCachedRows(
  chart: SpreadsheetChartLike,
): string[][] | null {
  const data = (chart.chartV2 as ChartBlockAttrsV2 | undefined)?.data;
  if (!data || data.kind !== "category" || data.series.length === 0) return null;
  return [
    ["", ...data.series.map((series) => series.name)],
    ...data.categories.map((category, index) => [
      category,
      ...data.series.map((series) => {
        const value = series.values[index];
        return typeof value === "number" && Number.isFinite(value) ? String(value) : "";
      }),
    ]),
  ];
}

/**
 * The multi-level category axis the chart's model last read (outer levels,
 * outermost first, one entry per category), for the rows
 * `spreadsheetChartCachedRows` gives; undefined for an axis of one level.
 */
export function spreadsheetChartCachedCategoryLevels(
  chart: SpreadsheetChartLike,
): string[][] | undefined {
  const data = (chart.chartV2 as ChartBlockAttrsV2 | undefined)?.data;
  if (!data || data.kind !== "category" || !data.categoryLevels?.length) return undefined;
  return data.categoryLevels.map((level) => [...level]);
}

/**
 * The X Y series of a chart whose data is its own (`binding.literal`), as
 * the sheet hands live X Y series to the planner: an imported scatter or
 * bubble chart the sheet binds no references for (C28_tdf126033, whose X
 * values are text and so plotted at 1, 2, ...). Null for any other chart.
 */
export function spreadsheetChartLiteralXYSeries(
  chart: SpreadsheetChartLike,
): MiniChartXYSeries[] | null {
  if (!normalizeSpreadsheetChartBinding(chart.binding)?.literal) return null;
  const data = (chart.chartV2 as ChartBlockAttrsV2 | undefined)?.data;
  if (!data || data.kind !== "xy" || !data.series.length) return null;
  return data.series.map((series) => ({
    name: series.name,
    points: series.xValues.flatMap((x, index) => {
      const y = series.yValues[index];
      const size = series.bubbleSizes?.[index];
      return typeof x === "number" && Number.isFinite(x) && typeof y === "number" && Number.isFinite(y)
        ? [{ x, y, ...(typeof size === "number" ? { size } : {}) }]
        : [];
    }),
  }));
}
