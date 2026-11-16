/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Filters > Values > Series, once for every editor that holds a Chart V2
 * chart (Microsoft Support, "Change the data series in a chart": clear a
 * series' check box and Apply; the series stays in the chart's data and is not
 * plotted).
 *
 * Office keeps a filtered series in its chart part: it moves out of the chart
 * group's plotted `c:ser` list into `c15:filtered<Group>Series` inside that
 * group's `c:extLst` (MS-ODRAWXML CT_FilteredBarSer and its line / area / pie
 * / scatter / radar / bubble / surface twins), with its references, caches,
 * formatting and index intact, so clearing the filter plots it again with its
 * own colour. The model states it as `format.hiddenSeries`; the server's chart
 * writer (`office_chart_filters.py`) moves those series, and the readers move
 * them back before reading.
 *
 * Values > Categories is kept the other way round: every series' category and
 * value references are narrowed to the categories still plotted (`c:f` the
 * union of the visible cells, the caches only their points) and the full
 * reference stays beside it as `c15:fullRef/c15:sqref` (MS-ODRAWXML fullRef,
 * "the non-filtered data source reference for filtered out chart series
 * data"). The model states it as `format.hiddenCategories`
 * (`office_chart_filters.with_filtered_categories` writes it).
 */
import {
  chartCategoryFullNames,
  chartSupportsCategoryFilter,
  chartSupportsSeriesFilter,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
} from "../../components/console/editor/extensions/chartBlockShared";

/** The chart with one series shown or hidden, or null (no change / refused). */
export function chartWithSeriesShown(
  attrs: ChartBlockAttrsV2,
  index: number,
  shown: boolean,
): ChartBlockAttrsV2 | null {
  if (!chartSupportsSeriesFilter(attrs.chartType) || !("series" in attrs.data)) return null;
  const hidden = new Set(attrs.format?.hiddenSeries ?? []);
  if (hidden.has(index) === !shown) return null;
  if (shown) hidden.delete(index);
  else hidden.add(index);
  const format: NonNullable<ChartBlockAttrsV2["format"]> = { ...attrs.format };
  if (hidden.size) format.hiddenSeries = [...hidden].sort((a, b) => a - b);
  else delete format.hiddenSeries;
  const next = { ...attrs, format };
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/**
 * Chart Design > Switch Row/Column moves the series to the categories and the
 * categories to the series, so a hidden series becomes a hidden category and
 * the other way round. A format with neither is returned as it is.
 */
export function officeChartFormatWithRowsAndColumnsSwitched<
  F extends { hiddenSeries?: number[]; hiddenCategories?: number[] },
>(format: F | undefined): F | undefined {
  if (!format || (!format.hiddenSeries && !format.hiddenCategories)) return format;
  const next = { ...format };
  delete next.hiddenSeries;
  delete next.hiddenCategories;
  if (format.hiddenCategories?.length) next.hiddenSeries = [...format.hiddenCategories];
  if (format.hiddenSeries?.length) next.hiddenCategories = [...format.hiddenSeries];
  return next;
}

/**
 * The format with every series, point and label statement past the switched
 * data's series and point counts dropped: a fourth series' colour on a
 * chart switched to two series is nobody's, and validation refused the
 * whole switch for it (PPT-B295).
 */
export function officeChartFormatBoundedToData(
  format: ChartBlockAttrsV2["format"],
  data: ChartBlockAttrsV2["data"],
): ChartBlockAttrsV2["format"] {
  if (!format || data.kind !== "category") return format;
  const seriesCount = data.series.length;
  const pointCount = data.categories.length;
  const next = { ...format };
  const bounded = <T extends { seriesIndex?: number; pointIndex?: number }>(
    list: readonly T[] | undefined,
  ): T[] | undefined => {
    if (!list) return undefined;
    const kept = list.filter(
      (entry) =>
        (entry.seriesIndex ?? 0) < seriesCount &&
        (entry.pointIndex === undefined || entry.pointIndex < pointCount),
    );
    return kept.length ? kept : undefined;
  };
  for (const key of ["seriesStyles", "sliceStyles", "dataLabelOverrides"] as const) {
    if (next[key] === undefined) continue;
    const kept = bounded(next[key] as readonly { seriesIndex?: number; pointIndex?: number }[]);
    if (kept) (next as Record<string, unknown>)[key] = kept;
    else delete next[key];
  }
  return next;
}

/**
 * Switch Row/Column on a Chart V2 chart whose data the host already switched
 * (or, without `data`, on its own category data): the data transposed, the
 * filters moved with it (dropped where the new layout cannot keep them: a
 * type that hides no series, or nothing left plotted). Null when the chart's
 * data is not category data.
 */
export function officeChartWithRowsAndColumnsSwitched(
  attrs: ChartBlockAttrsV2,
  data?: ChartBlockAttrsV2["data"],
): ChartBlockAttrsV2 | null {
  const source = attrs.data;
  if (source.kind !== "category") return null;
  // A multi-level category becomes a series named by all its levels
  // ("East Q1"), the name Excel gives it; the new categories are one level.
  const names = chartCategoryFullNames(source);
  const switched: ChartBlockAttrsV2["data"] = data ?? {
    kind: "category",
    // A series the part names nothing ("Series n" by its place) gives its
    // category no words: Excel's axis draws its place, "1", and its part
    
    categories: source.series.map((series, index) =>
      series.nameUnstated && series.name === `Series ${index + 1}` ? String(index + 1) : series.name,
    ),
    series: source.categories.map((_category, categoryIndex) => ({
      name: names[categoryIndex],
      values: source.series.map((series) => series.values[categoryIndex]),
    })),
  };
  let next: ChartBlockAttrsV2 = {
    ...attrs,
    data: switched,
    format: officeChartFormatBoundedToData(
      officeChartFormatWithRowsAndColumnsSwitched(attrs.format),
      switched,
    ),
  };
  for (const key of ["hiddenSeries", "hiddenCategories"] as const) {
    if (next.format?.[key] === undefined) continue;
    const trial = { ...next, format: { ...next.format } };
    if (validateChartBlockAttrsV2(trial).valid) break;
    delete trial.format[key];
    next = trial;
  }
  return next;
}

/** The chart with one category shown or hidden, or null (no change / refused). */
export function chartWithCategoryShown(
  attrs: ChartBlockAttrsV2,
  index: number,
  shown: boolean,
): ChartBlockAttrsV2 | null {
  if (!chartSupportsCategoryFilter(attrs.chartType) || attrs.data.kind !== "category") return null;
  const hidden = new Set(attrs.format?.hiddenCategories ?? []);
  if (hidden.has(index) === !shown) return null;
  if (shown) hidden.delete(index);
  else hidden.add(index);
  const format: NonNullable<ChartBlockAttrsV2["format"]> = { ...attrs.format };
  if (hidden.size) format.hiddenCategories = [...hidden].sort((a, b) => a - b);
  else delete format.hiddenCategories;
  const next = { ...attrs, format };
  return validateChartBlockAttrsV2(next).valid ? next : null;
}

/**
 * The Chart Filters controls every host hands the quick buttons: the hidden
 * series and categories, a toggle for each where the chart's type and the
 * host's writers state that filter (`commit` null: the chart cannot be
 * changed here), and why a greyed check box is greyed.
 */
export function officeChartFilterControls(
  attrs: ChartBlockAttrsV2 | null | undefined,
  commit: ((next: ChartBlockAttrsV2) => void) | null,
  readOnlyReason?: string,
): {
  hiddenSeries?: readonly number[];
  hiddenCategories?: readonly number[];
  onToggleSeries?: (index: number, shown: boolean) => void;
  onToggleCategory?: (index: number, shown: boolean) => void;
  unavailableReason: string;
  /** The catalogue key of the reason, where it is this module's own. */
  unavailableReasonKey?: string;
} {
  const series = Boolean(attrs && commit && chartSupportsSeriesFilter(attrs.chartType));
  const categories = Boolean(
    attrs && commit && chartSupportsCategoryFilter(attrs.chartType) && attrs.data.kind === "category",
  );
  const reasonKey = !commit
    ? readOnlyReason === undefined
      ? "office.chartFilters.cannotChange"
      : undefined
    : !series && !categories
      ? "office.chartFilters.notPlottedSeparately"
      : !series
        ? "office.chartFilters.oneSeries"
        : "office.chartFilters.notPerCategory";
  return {
    hiddenSeries: attrs?.format?.hiddenSeries,
    hiddenCategories: attrs?.format?.hiddenCategories,
    onToggleSeries:
      series && attrs && commit
        ? (index, shown) => {
            const next = chartWithSeriesShown(attrs, index, shown);
            if (next) commit(next);
          }
        : undefined,
    onToggleCategory:
      categories && attrs && commit
        ? (index, shown) => {
            const next = chartWithCategoryShown(attrs, index, shown);
            if (next) commit(next);
          }
        : undefined,
    unavailableReason: reasonKey ? "" : (readOnlyReason ?? ""),
    unavailableReasonKey: reasonKey,
  };
}
