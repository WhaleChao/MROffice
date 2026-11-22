/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a chart's elements stand when they were dragged by hand: Office's
 * `c:layout/c:manualLayout` of the title, the axis titles, the legend and
 * the plot area (Microsoft Support, "Change the placement of titles in a
 * chart": drag a title by its border to where you want it; a legend and the
 * plot area move and size the same way). An element with no place of its
 * own is laid out by the chart; choosing a place from the ribbon (a legend
 * position, a quick layout) gives it back to the chart.
 *
 * Every number is a fraction of the chart space: `x`, `y` the element box's
 * top-left corner (`c:xMode`/`c:yMode` "edge"), `w`, `h` its size (factor).
 */

/** A text element's place: its box's top-left corner. */
export interface ChartElementPlace {
  readonly x: number;
  readonly y: number;
}

/** A box element's place and size. */
export interface ChartElementBox extends ChartElementPlace {
  readonly w: number;
  readonly h: number;
}

export interface ChartElementLayout {
  readonly title?: ChartElementPlace;
  readonly axisTitleCategory?: ChartElementPlace;
  readonly axisTitleValue?: ChartElementPlace;
  /** A 3-D chart's depth (series) axis title (`format.seriesAxisTitle`). */
  readonly axisTitleSeries?: ChartElementPlace;
  

  readonly legend?: ChartElementBox | ChartElementPlace;
  /** The INNER plot area (the plot without its axes' labels). */
  readonly plotArea?: ChartElementBox;
}

export type ChartElementLayoutKey = keyof ChartElementLayout;

export const CHART_ELEMENT_PLACE_KEYS = ["title", "axisTitleCategory", "axisTitleValue", "axisTitleSeries"] as const;
export const CHART_ELEMENT_BOX_KEYS = ["legend", "plotArea"] as const;

const EPSILON = 1e-6;
const round6 = (value: number) => Math.round(value * 1e6) / 1e6;
const fraction = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 1;

function normalizePlace(value: unknown): ChartElementPlace | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if (Object.keys(record).some((key) => key !== "x" && key !== "y")) return undefined;
  if (!fraction(record.x) || !fraction(record.y)) return undefined;
  return { x: round6(record.x), y: round6(record.y) };
}

/** `value` as an element box (x, y, w, h fractions inside the chart), or undefined. */
export function normalizeChartElementBox(value: unknown): ChartElementBox | undefined {
  return normalizeBox(value);
}

function normalizeBox(value: unknown): ChartElementBox | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if (Object.keys(record).some((key) => !["x", "y", "w", "h"].includes(key))) return undefined;
  const { x, y, w, h } = record;
  if (!fraction(x) || !fraction(y) || !fraction(w) || !fraction(h) || w <= 0 || h <= 0) return undefined;
  if (x + w > 1 + EPSILON || y + h > 1 + EPSILON) return undefined;
  return { x: round6(x), y: round6(y), w: round6(w), h: round6(h) };
}

/** `value` as an element layout, or undefined when it is not one (or names nothing). */
export function normalizeChartElementLayout(value: unknown): ChartElementLayout | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (!keys.length) return undefined;
  const layout: Record<string, ChartElementPlace | ChartElementBox> = {};
  for (const key of keys) {
    if ((CHART_ELEMENT_PLACE_KEYS as readonly string[]).includes(key)) {
      const place = normalizePlace(record[key]);
      if (!place) return undefined;
      layout[key] = place;
    } else if ((CHART_ELEMENT_BOX_KEYS as readonly string[]).includes(key)) {
      // A legend may be placed by its corner alone (no size of its own).
      const box = normalizeBox(record[key]) ?? (key === "legend" ? normalizePlace(record[key]) : undefined);
      if (!box) return undefined;
      layout[key] = box;
    } else {
      return undefined;
    }
  }
  return layout as ChartElementLayout;
}

/** Whether a placed element states its size (a legend placed by its corner alone does not). */
export function chartElementPlaceIsBox(place: ChartElementPlace | ChartElementBox | undefined): place is ChartElementBox {
  return place !== undefined && typeof (place as Partial<ChartElementBox>).w === "number" && typeof (place as Partial<ChartElementBox>).h === "number";
}

/** The layout's place for an element, kept inside the chart. */
export function chartElementPlaceWithin(place: ChartElementPlace): ChartElementPlace {
  return { x: round6(Math.min(1, Math.max(0, place.x))), y: round6(Math.min(1, Math.max(0, place.y))) };
}

/** The layout's box for an element, kept inside the chart (at least `min` in size). */
export function chartElementBoxWithin(box: ChartElementBox, min = 0.02): ChartElementBox {
  const w = Math.min(1, Math.max(min, box.w));
  const h = Math.min(1, Math.max(min, box.h));
  const x = Math.min(1 - w, Math.max(0, box.x));
  const y = Math.min(1 - h, Math.max(0, box.y));
  return { x: round6(x), y: round6(y), w: round6(w), h: round6(h) };
}

/** The layout with an element's place set (`null`: given back to the chart). */
export function withChartElementPlace(
  layout: ChartElementLayout | undefined,
  key: ChartElementLayoutKey,
  place: ChartElementPlace | ChartElementBox | null,
): ChartElementLayout | undefined {
  const next: Record<string, unknown> = { ...(layout ?? {}) };
  if (place === null) delete next[key];
  else next[key] = (CHART_ELEMENT_BOX_KEYS as readonly string[]).includes(key) && chartElementPlaceIsBox(place)
    ? chartElementBoxWithin(place)
    : chartElementPlaceWithin(place);
  return normalizeChartElementLayout(next);
}

/** The Office 2016 families (`cx:` charts): their parts lay elements out by
 * position only (no `c:manualLayout`, no `c:legendEntry`), so nothing is
 * placed by hand on them. */
const PLACED_BY_POSITION_ONLY: ReadonlySet<string> = new Set(["hierarchy", "statistical", "waterfall", "funnel", "map"]);

/** Whether a chart's elements may be placed by hand (and its legend's entries deleted). */
export function chartTakesElementLayout(chartType: string): boolean {
  return !PLACED_BY_POSITION_ONLY.has(chartType);
}

/**
 * The directions a dragged data label's offset (`c:dLbl/c:layout/
 * c:manualLayout`, `dataLabelOverrides[].offset`) runs in the picture: +1
 * as stated, -1 mirrored. Office states the offset along the value axis'
 * own direction, so on a bar or column chart whose values run in reverse
 * order (Format Axis > Values in reverse order, `c:orientation` maxMin) the
 * value-direction offset runs the other way on the page: CDC p34's two
 * charts state the same -0.1175 x offset for one Outside End label, and
 * PowerPoint draws both just inside their bars' ends -- the males' bars
 * hanging from the right, the females' from the left. Read mirrored as
 * stated, the males' label stood off the chart.
 */
export function chartDataLabelOffsetSigns(attrs: {
  chartType?: unknown;
  format?: { axisOptions?: { value?: { reverse?: unknown } } } | null;
}): { x: 1 | -1; y: 1 | -1 } {
  const reversed = attrs.format?.axisOptions?.value?.reverse === true;
  if (!reversed) return { x: 1, y: 1 };
  if (attrs.chartType === "bar") return { x: -1, y: 1 };
  if (attrs.chartType === "column") return { x: 1, y: -1 };
  return { x: 1, y: 1 };
}

/** ECharts' place of a fraction of the chart (a percentage of its box). */
export const chartPercent = (value: number) => `${round6(value * 100)}%`;
