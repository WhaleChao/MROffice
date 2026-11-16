/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Design > Quick Layout: Excel's eleven column / bar layouts, the ONE
 * recipe set of the three editors.
 *
 * Microsoft documents Quick Layout as one command of the shared Chart Design
 * tab ("Change the layout or style of a chart": Chart Design > Quick Layout,
 * then pick a layout). A layout switches the chart's elements AND lays them
 * out: title and axis titles, legend edge, major gridlines, minor gridlines,
 * major tick marks, the bars' gap width and overlap, and data labels on
 * whole series or on single points. These eleven were calibrated against
 * Excel's own package (the worksheet's native sample) and were the worksheet's
 * alone; Docs and the deck drew five generic ones. They are the one catalogue
 * now: a column or bar chart takes all eleven in every editor, carried as
 * `format.layout` (`OfficeChartLayout`) that every classic writer states the
 * way Excel does (`c:gapWidth`, `c:overlap`, `c:minorGridlines`,
 * `c:majorTickMark`, per-series `c:dLbls` and per-point `c:dLbl`). Any other
 * family takes the first five, whose elements it can show
 * (`lib/office-charts/quick-layout.ts`).
 */
export const OFFICE_CHART_QUICK_LAYOUT_IDS = [
  "layout-1",
  "layout-2",
  "layout-3",
  "layout-4",
  "layout-5",
  "layout-6",
  "layout-7",
  "layout-8",
  "layout-9",
  "layout-10",
  "layout-11",
] as const;
export type OfficeChartQuickLayoutId = (typeof OFFICE_CHART_QUICK_LAYOUT_IDS)[number];
/** A data-label override on one series, or one point of it. */
export interface OfficeChartLabelOverride {
  seriesIndex: number;
  pointIndex?: number;
  value?: boolean;
  seriesName?: boolean;
  position?: "outEnd" | "inEnd" | "ctr" | "inBase" | "bestFit";
  /**
   * A series' own labels stand where their chart type stands them: its own
   * `c:dLbls` states no `c:dLblPos` where its group's does. A series', never
   * beside a position of its own (the server's `normalize_layout`).
   */
  positionAutomatic?: true;
}
/** What a calibrated layout lays out beyond the element switches. */
export interface OfficeChartLayout {
  title: boolean;
  axisTitleX: boolean;
  axisTitleY: boolean;
  minorGridlinesX: boolean;
  minorGridlinesY: boolean;
  majorTickX: "none" | "out" | "in" | "cross";
  majorTickY: "none" | "out" | "in" | "cross";
  gapWidth: number;
  overlap: number;
  labels: OfficeChartLabelOverride[];
}
const flags = [
  "title",
  "axisTitleX",
  "axisTitleY",
  "minorGridlinesX",
  "minorGridlinesY",
] as const;
const positions = new Set(["outEnd", "inEnd", "ctr", "inBase", "bestFit"]);
const record = (v: unknown): v is Record<string, unknown> =>
  Boolean(v) && typeof v === "object" && !Array.isArray(v);
export function normalizeOfficeChartLayout(
  raw: unknown,
): OfficeChartLayout | undefined {
  if (
    !record(raw) ||
    Object.keys(raw).sort().join() !==
      [...flags, "majorTickX", "majorTickY", "gapWidth", "overlap", "labels"]
        .sort()
        .join()
  )
    return;
  if (
    flags.some((k) => typeof raw[k] !== "boolean") ||
    ["majorTickX", "majorTickY"].some(
      (k) => !["none", "out", "in", "cross"].includes(String(raw[k])),
    ) ||
    !Number.isInteger(raw.gapWidth) ||
    Number(raw.gapWidth) < 0 ||
    Number(raw.gapWidth) > 500 ||
    !Number.isInteger(raw.overlap) ||
    Number(raw.overlap) < -100 ||
    Number(raw.overlap) > 100 ||
    !Array.isArray(raw.labels) ||
    raw.labels.length > 128
  )
    return;
  const seen = new Set<string>();
  for (const item of raw.labels) {
    if (
      !record(item) ||
      Object.keys(item).some(
        (k) =>
          ![
            "seriesIndex",
            "pointIndex",
            "value",
            "seriesName",
            "position",
            "positionAutomatic",
          ].includes(k),
      ) ||
      !Number.isInteger(item.seriesIndex) ||
      Number(item.seriesIndex) < 0 ||
      Number(item.seriesIndex) > 11 ||
      (item.pointIndex !== undefined &&
        (!Number.isInteger(item.pointIndex) ||
          Number(item.pointIndex) < 0 ||
          Number(item.pointIndex) > 10000)) ||
      ["value", "seriesName"].some(
        (k) => item[k] !== undefined && typeof item[k] !== "boolean",
      ) ||
      (item.position !== undefined && !positions.has(String(item.position))) ||
      (item.positionAutomatic !== undefined &&
        (item.positionAutomatic !== true || item.pointIndex !== undefined || item.position !== undefined)) ||
      !["value", "seriesName", "position", "positionAutomatic"].some((k) => item[k] !== undefined)
    )
      return;
    const key = `${item.seriesIndex}:${item.pointIndex ?? "series"}`;
    if (seen.has(key)) return;
    seen.add(key);
  }
  return raw as unknown as OfficeChartLayout;
}












export function officeChartLayoutWithAxisTicks<F extends { readonly layout?: unknown; readonly axisOptions?: unknown }>(
  chartType: unknown,
  format: F,
): F {
  const layout = format.layout;
  const options = format.axisOptions;
  if (!record(layout) || !record(options)) return format;
  const horizontal = chartType === "bar";
  const stated: Record<string, string> = {};
  for (const [suffix, axis] of [
    ["X", horizontal ? "value" : "category"],
    ["Y", horizontal ? "category" : "value"],
  ] as const) {
    const own = options[axis];
    const mark = record(own) ? own.majorTickMark : undefined;
    if (
      typeof mark === "string" &&
      ["none", "out", "in", "cross"].includes(mark) &&
      layout[`majorTick${suffix}`] !== mark
    ) {
      stated[`majorTick${suffix}`] = mark;
    }
  }
  return Object.keys(stated).length ? { ...format, layout: { ...layout, ...stated } } : format;
}









export function officeChartLayoutWithBarSpacing<F extends { readonly layout?: unknown; readonly barGapWidth?: unknown; readonly barOverlap?: unknown }>(
  format: F,
): F {
  const layout = format.layout;
  if (!record(layout)) return format;
  const stated: Record<string, number> = {};
  for (const [own, key] of [["barGapWidth", "gapWidth"], ["barOverlap", "overlap"]] as const) {
    const value = format[own];
    if (typeof value === "number" && Number.isInteger(value) && layout[key] !== value) stated[key] = value;
  }
  return Object.keys(stated).length ? { ...format, layout: { ...layout, ...stated } } : format;
}

/** Format Data Labels' positions a calibrated layout states (`c:dLblPos`). */
const LABEL_POSITION_TOKENS: Readonly<Record<string, OfficeChartLabelOverride["position"]>> = {
  bestFit: "bestFit",
  center: "ctr",
  insideBase: "inBase",
  insideEnd: "inEnd",
  outsideEnd: "outEnd",
};










export function officeChartLayoutWithLabelOverrides<F extends { readonly layout?: unknown; readonly dataLabelOverrides?: unknown; readonly dataLabelPosition?: unknown }>(
  format: F,
): F {
  const layout = format.layout;
  const overrides = format.dataLabelOverrides;
  if (!record(layout) || !Array.isArray(overrides) || !overrides.length) return format;
  const labels = Array.isArray(layout.labels) ? (layout.labels as OfficeChartLabelOverride[]).map((item) => ({ ...item })) : [];
  let changed = false;
  for (const entry of overrides) {
    if (!record(entry) || typeof entry.seriesIndex !== "number") continue;
    const fields: { value?: boolean; seriesName?: boolean; position?: OfficeChartLabelOverride["position"]; positionAutomatic?: true } = {};
    if (entry.dataLabels === false) {
      fields.value = false;
      fields.seriesName = false;
    } else {
      if (typeof entry.dataLabelValue === "boolean") fields.value = entry.dataLabelValue;
      if (typeof entry.dataLabelSeriesName === "boolean") fields.seriesName = entry.dataLabelSeriesName;
      const position = typeof entry.dataLabelPosition === "string" ? LABEL_POSITION_TOKENS[entry.dataLabelPosition] : undefined;
      if (position) fields.position = position;
      else if (
        entry.dataLabelPositionAutomatic === true &&
        typeof entry.pointIndex !== "number" &&
        typeof format.dataLabelPosition === "string" &&
        LABEL_POSITION_TOKENS[format.dataLabelPosition]
      ) {
        
        // 1): the layout's series stands at the chart type's place.
        fields.positionAutomatic = true;
      }
    }
    if (!Object.keys(fields).length) continue;
    const pointIndex = typeof entry.pointIndex === "number" ? entry.pointIndex : undefined;
    const at = labels.findIndex((item) => item.seriesIndex === entry.seriesIndex && item.pointIndex === pointIndex);
    const merged: OfficeChartLabelOverride = { ...(at >= 0 ? labels[at] : { seriesIndex: entry.seriesIndex, ...(pointIndex !== undefined ? { pointIndex } : {}) }), ...fields };
    if (fields.position) delete merged.positionAutomatic;
    else if (fields.positionAutomatic) delete merged.position;
    if (at >= 0 && JSON.stringify(labels[at]) === JSON.stringify(merged)) continue;
    if (at >= 0) labels[at] = merged;
    else labels.push(merged);
    changed = true;
  }
  return changed ? { ...format, layout: { ...layout, labels } } : format;
}

/** The calibrated layout a chart draws and is written with: its stated
 * layout, with the ticks its axes, the spacing its Series Options and the
 * label switches Format Data Labels state (`officeChartLayoutWithAxisTicks`,
 * `officeChartLayoutWithBarSpacing`, `officeChartLayoutWithLabelOverrides`). */
export function officeChartLayoutAsStated(
  chartType: unknown,
  format:
    | { readonly layout?: unknown; readonly axisOptions?: unknown; readonly barGapWidth?: unknown; readonly barOverlap?: unknown; readonly dataLabelOverrides?: unknown; readonly dataLabelPosition?: unknown }
    | undefined,
): OfficeChartLayout | undefined {
  if (!format || !record(format.layout)) return undefined;
  return officeChartLayoutWithLabelOverrides(officeChartLayoutWithBarSpacing(officeChartLayoutWithAxisTicks(chartType, format)))
    .layout as OfficeChartLayout;
}

/**
 * A chart format with one record of its layout's atoms: where it has a
 * layout, the layout takes what the chart's own keys state (its axes' ticks,
 * its Series Options' spacing) and the spacing keys go, so an edit of the
 * layout is not answered by a stale own key. The one form every command that
 * writes a chart's format commits.
 */
export function officeChartFormatWithOneLayoutRecord<F extends { readonly layout?: unknown; readonly axisOptions?: unknown; readonly barGapWidth?: unknown; readonly barOverlap?: unknown; readonly dataLabelOverrides?: unknown; readonly dataLabelPosition?: unknown }>(
  chartType: unknown,
  format: F,
): F {
  if (!record(format.layout)) return format;
  const { barGapWidth: _gap, barOverlap: _overlap, ...rest } = format;
  return { ...rest, layout: officeChartLayoutAsStated(chartType, format) } as unknown as F;
}

export function officeChartQuickLayoutRecipe(
  id: OfficeChartQuickLayoutId,
  categoryCount = 3,
  seriesCount = 2,
) {
  const index = OFFICE_CHART_QUICK_LAYOUT_IDS.indexOf(id) + 1;
  const layout: OfficeChartLayout = {
    title: ![4, 7, 11].includes(index),
    axisTitleX: [7, 8, 9].includes(index),
    axisTitleY: [5, 6, 7, 8, 9].includes(index),
    minorGridlinesX: false,
    minorGridlinesY: index === 7,
    majorTickX: index === 11 ? "out" : "none",
    majorTickY: [7, 8, 9, 11].includes(index) ? "out" : "none",
    gapWidth:
      index === 7
        ? 300
        : index === 8
          ? 0
          : [3, 4, 10].includes(index)
            ? 75
            : 150,
    overlap: [2, 3].includes(index) ? -25 : index === 10 ? 40 : 0,
    labels:
      index === 6
        ? Array.from(
            { length: Math.min(12, seriesCount) },
            (_, seriesIndex) => ({
              seriesIndex,
              pointIndex: Math.max(0, categoryCount - 1),
              value: true,
              seriesName: true,
            }),
          )
        : index === 10
          ? [
              {
                seriesIndex: Math.max(0, Math.min(12, seriesCount) - 1),
                value: true,
                position: "inEnd",
              },
            ]
          : [],
  };
  const legend = [5, 6, 8].includes(index)
    ? "none"
    : index === 2
      ? "top"
      : [3, 4].includes(index)
        ? "bottom"
        : "right";
  return {
    layout,
    legend,
    axisX: true,
    axisY: index !== 2,
    gridlinesX: false,
    gridlinesY: ![2, 4, 8].includes(index),
    dataLabels: [2, 4].includes(index),
    dataLabelValue: [2, 4].includes(index),
    dataLabelSeriesName: false,
    dataLabelLegendKey: false,
    dataLabelCategoryName: false,
    dataLabelPercentage: false,
    dataLabelBubbleSize: false,
    dataTable: index === 5,
    dataTableLegendKeys: index === 5,
  } as const;
}


/**
 * The charts Excel's calibrated layouts are for: a 2-D clustered or stacked
 * column or bar chart (one `c:barChart` plotted on a category and a value
 * axis). The worksheet has only ever offered them there.
 */
export function officeChartTakesCalibratedLayouts(
  family: unknown,
  subtype: unknown,
): boolean {
  return (
    (family === "column" || family === "bar") &&
    (subtype === "clustered" || subtype === "stacked")
  );
}

/**
 * The layout a column / bar chart is drawn with: its own calibrated layout,
 * or, for a chart no Quick Layout has touched, the one its writers state
 * today (Excel's default: no tick marks, gap width 150, the bars side by side
 * or, stacked, fully overlapped; no minor gridlines, no single labels). An
 * element pick that only a layout carries (minor gridlines) starts from it,
 * so nothing else on the chart moves. Null for a chart Excel's layouts are
 * not for.
 */
export function officeChartLayoutOf(chart: {
  chartType?: unknown;
  chartSubtype?: unknown;
  title?: unknown;
  axisTitleX?: unknown;
  axisTitleY?: unknown;
  layout?: unknown;
  /** The chart's format: a layout made for it takes the tick marks its axes
   *  already state (`axisOptions.<axis>.majorTickMark`) and the spacing its
   *  Series Options state (`barGapWidth` / `barOverlap`). */
  format?: unknown;
}): OfficeChartLayout | null {
  if (!officeChartTakesCalibratedLayouts(chart.chartType, chart.chartSubtype))
    return null;
  const format = record(chart.format) ? chart.format : {};
  const stated = normalizeOfficeChartLayout(chart.layout);
  // A stored layout as the chart states it (a page saved while the two
  // records disagreed): what it draws, and what a command edits.
  if (stated) return officeChartLayoutAsStated(chart.chartType, { ...format, layout: stated }) ?? stated;
  // A layout made for a chart that had none (minor gridlines turned on, a
  // Chart Style's spacing) keeps the tick marks the chart's axes state and
  // the spacing it states: made from Office's defaults, it wrote an imported
  
  
  const made: OfficeChartLayout = {
    title: Boolean(chart.title),
    axisTitleX: Boolean(chart.axisTitleX),
    axisTitleY: Boolean(chart.axisTitleY),
    minorGridlinesX: false,
    minorGridlinesY: false,
    majorTickX: "none",
    majorTickY: "none",
    gapWidth: 150,
    // A stacked or 100% stacked chart's bars lie on one another.
    overlap: chart.chartSubtype === "stacked" || chart.chartSubtype === "percentStacked" ? 100 : 0,
    labels: [],
  };
  return officeChartLayoutAsStated(chart.chartType, { ...format, layout: made }) ?? made;
}









export function officeInsertedChartLayout(chart: {
  chartType?: unknown;
  chartSubtype?: unknown;
  title?: unknown;
}): OfficeChartLayout | null {
  const base = officeChartLayoutOf({ ...chart, layout: undefined });
  if (!base) return null;
  if (chart.chartSubtype !== "clustered") return base;
  return chart.chartType === "bar"
    ? { ...base, gapWidth: 182, overlap: 0 }
    : { ...base, gapWidth: 219, overlap: -27 };
}
