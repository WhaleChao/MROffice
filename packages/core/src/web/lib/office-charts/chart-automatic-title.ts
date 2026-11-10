/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's automatic chart title (`format.automaticTitle`).
 *
 * A chart part with an empty `c:title` that `c:autoTitleDeleted` leaves
 * standing has a title Office makes up: over a chart of one named series the
 * series' name (a pie's first series, the one it plots), otherwise "Chart
 * Title" in the reader's own language (Word's
 * pages of the round-trip corpus: R29-D05, R30-D12, R34-D07 draw "Chart
 * Title", R34-D09 and R19-D04 their one series' name, R19-D23's nameless
 * series "Chart Title"). The words are never the chart's own: the model keeps
 * no title and the file keeps the empty `c:title`, so the drawing follows the
 * series name and the language. Only drawing reads them.
 */

import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";

/** The words Office draws for the chart's automatic title, or "" for none. */
export function chartAutomaticTitleText(
  attrs: ChartBlockAttrsV2,
  chartTitleWord: string,
  /** Whether the chart's series `index` states a name (`c:tx`); a series
   * named only by its place (Excel's 系列1, "Series 1") titles nothing:
   * Excel reads "Chart Title" (C09_tdf133190_tdf133191). Absent, every
   * series' name counts. */
  seriesNamed?: (index: number) => boolean,
): string {
  if (attrs.title || attrs.format?.automaticTitle !== true) return "";
  const series = (attrs.data as { series?: readonly { name?: unknown }[] } | undefined)?.series;
  // A pie (a Pie of Pie and a Bar of Pie too, not a doughnut) plots its first
  // series alone, whatever else the part keeps: Excel titles it with that
  // series' name (C17_strict_chart: two series, titled "col1").
  const plotsFirstOnly = attrs.chartType === "pie" && attrs.chartSubtype !== "doughnut";
  if (Array.isArray(series) && (series.length === 1 || (plotsFirstOnly && series.length > 1))) {
    const name = series[0]?.name;
    if (typeof name === "string" && name.trim() && seriesNamed?.(0) !== false) return name;
  }
  return chartTitleWord;
}

/** The chart as it is drawn: its automatic title, if it has one, as its title. */
export function chartWithAutomaticTitle(
  attrs: ChartBlockAttrsV2,
  chartTitleWord: string,
  seriesNamed?: (index: number) => boolean,
): ChartBlockAttrsV2 {
  const text = chartAutomaticTitleText(attrs, chartTitleWord, seriesNamed);
  return text ? { ...attrs, title: text } : attrs;
}

/**
 * The words a part reads for an automatic axis title (`c:title` with no
 * `c:tx`): Office draws its placeholder in the reader's language ("Axis
 * Title", 坐标轴标题; C08_axis_title_rotated, X-156). The
 * reading stores the English placeholder; only drawing reads the language.
 */
export const AUTOMATIC_AXIS_TITLE_READING = "Axis Title";

/** The chart as it is drawn: an automatic axis title in the reader's words. */
export function chartWithAutomaticAxisTitles(attrs: ChartBlockAttrsV2, axisTitleWord: string | undefined): ChartBlockAttrsV2 {
  if (!axisTitleWord || axisTitleWord === AUTOMATIC_AXIS_TITLE_READING) return attrs;
  const x = attrs.axisTitleX === AUTOMATIC_AXIS_TITLE_READING;
  const y = attrs.axisTitleY === AUTOMATIC_AXIS_TITLE_READING;
  return x || y
    ? { ...attrs, ...(x ? { axisTitleX: axisTitleWord } : {}), ...(y ? { axisTitleY: axisTitleWord } : {}) }
    : attrs;
}
