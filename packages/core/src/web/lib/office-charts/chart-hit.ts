/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A click on a drawn chart, as the Format pane's Current Selection names it,
 * in every editor alike.
 *
 * Office selects a chart's parts by clicking: one click on a series selects
 * the series (every point and label of it), a second click on one of its
 * points selects that point alone, and Format then formats what is selected
 * -- a series' labels, one point's label (Microsoft Support, "Add or remove
 * data labels in a chart" and "Move data labels": one click selects the data
 * labels of the whole series, a second click the one label). A host reports
 * what its drawing was hit on -- the drawn series'
 * name and the point's index among the drawn points -- and the pane turns it
 * into `series:N` or `point:N:P` of the chart's own series and points, whose
 * drawing may leave some out (Chart Filters) or add some (a trendline, a
 * helper series).
 */
import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
// A type only: this module also serves the chart engine, which has no React.
import type { OfficeChartSelectionDetail } from "./chart-selection-store";

/** What the drawing was hit on: the drawn series' name (as the plan names
 * it) and the point's index among that series' drawn points. */
export interface OfficeChartHit {
  readonly seriesName: string;
  readonly dataIndex: number | null;
  /** A click on a marked part of the picture (the title, an axis title, the
   * legend, an axis's labels, the data labels) rather than on a series:
   * Office selects that element (`seriesName` is then ""). */
  readonly part?: OfficeChartPicturePart;
  /** On the legend: the entry the click landed on (its index in the
   * legend's whole list). */
  readonly legendEntry?: number;
  /** On a data label: whose it is (the chart's own series and point). */
  readonly label?: { readonly series: number; readonly point: number };
  /** On a plotted point: its place in the PICTURE (the drawn series and
   * point, hidden ones left out), which the picture marks as selected. */
  readonly plotted?: { readonly series: number; readonly point: number | null };
}

/** The planner's mark on a repeated series name (`withDistinctSeriesNames`). */
const SAME_NAME_MARK = "​";

/** The families whose points are addressed one by one on the pane (a pie's
 * points are its `slice:N` entries, listed there already). */
const POINT_FAMILIES = new Set(["column", "bar", "line", "area", "radar", "combo", "xy"]);

const POINT_AREA = /^point:(\d+):(\d+)$/;

export function officeChartPointArea(area: string): { seriesIndex: number; pointIndex: number } | null {
  const match = POINT_AREA.exec(area);
  return match ? { seriesIndex: Number(match[1]), pointIndex: Number(match[2]) } : null;
}

/** The chart's series index a drawn series name stands for, or null for a
 * series the chart does not own (a trendline's, a helper's). */
export function officeChartHitSeries(attrs: ChartBlockAttrsV2, seriesName: string): number | null {
  if (!("series" in attrs.data)) return null;
  const series = attrs.data.series as readonly { name: string }[];
  const hidden = new Set(attrs.format?.hiddenSeries ?? []);
  const plotted = series.map((_, index) => index).filter((index) => !hidden.has(index));
  const base = seriesName.split(SAME_NAME_MARK).join("");
  const repeat = seriesName.length - base.length;
  const named = plotted.filter((index) => series[index].name === base);
  return named[repeat] ?? null;
}

/** The chart's point index a drawn point stands for (Chart Filters leave
 * hidden categories out of the drawing). */
export function officeChartHitPoint(attrs: ChartBlockAttrsV2, dataIndex: number): number | null {
  if (attrs.data.kind === "category") {
    const hidden = new Set(attrs.format?.hiddenCategories ?? []);
    const plotted = attrs.data.categories
      .map((_, index) => index)
      .filter((index) => !hidden.has(index));
    return plotted[dataIndex] ?? null;
  }
  if (attrs.data.kind === "xy") return dataIndex;
  return null;
}

/**
 * A pie's slice names in its plotted points' order, as the planner names
 * them: ECharts keys a pie's slices by name, so a repeated category carries
 * one invisible mark per earlier namesake and stays a slice of its own.
 */
export function officeChartPieSliceNames(categories: readonly string[]): string[] {
  const seen = new Map<string, number>();
  return categories.map((category) => {
    const repeat = seen.get(category) ?? 0;
    seen.set(category, repeat + 1);
    return category + SAME_NAME_MARK.repeat(repeat);
  });
}

/**
 * The plotted point a drawn pie item stands for, found by its slice name
 * (`officeChartPieSliceNames`) rather than its place: a pie of pie and a bar
 * of pie draw their points sorted by size, and their "Other" slice stands for
 * no one point.
 */
export function officeChartPieItemPoint(item: unknown, sliceNames: readonly string[]): number | null {
  const name = item !== null && typeof item === "object" ? (item as { name?: unknown }).name : undefined;
  if (typeof name !== "string") return null;
  const index = sliceNames.indexOf(name);
  return index < 0 ? null : index;
}

/**
 * The Current Selection a click makes: the series on the first click, the
 * point on a click on a point of the series already selected (or of one of
 * its points), null for what the pane does not select by a click.
 */
export function officeChartHitArea(
  attrs: ChartBlockAttrsV2,
  hit: OfficeChartHit,
  current: string | null,
): string | null {
  if (hit.part) return officeChartPartArea(attrs, hit.part);
  if (attrs.chartType === "pie") {
    // A pie's points are its slices, one click each (the doughnut's rings
    // and a 3-D pie's slab have no slice entries).
    if (attrs.chartSubtype !== "pie2d" || hit.dataIndex === null) return null;
    const point = officeChartHitPoint(attrs, hit.dataIndex);
    return point === null ? null : `slice:${point}`;
  }
  const seriesIndex = officeChartHitSeries(attrs, hit.seriesName);
  if (seriesIndex === null) return null;
  const series = `series:${seriesIndex}`;
  const selectedPoint = current ? officeChartPointArea(current) : null;
  const again = current === series || selectedPoint?.seriesIndex === seriesIndex;
  if (!again || hit.dataIndex === null || !POINT_FAMILIES.has(attrs.chartType)) return series;
  const point = officeChartHitPoint(attrs, hit.dataIndex);
  return point === null ? series : `point:${seriesIndex}:${point}`;
}

/**
 * The selection a click makes, with what inside the element it landed on
 * (Office's clicks): the legend, then one of its entries on a second click;
 * a data label's series' labels, then that one label on a second click;
 * everything else as `officeChartHitArea` names it.
 */
export function officeChartHitSelection(
  attrs: ChartBlockAttrsV2,
  hit: OfficeChartHit,
  current: string | null,
  currentDetail: OfficeChartSelectionDetail | null,
): { area: string | null; detail: OfficeChartSelectionDetail | null } {
  if (hit.part === "legend") {
    const again = current === "legendStyle";
    return {
      area: "legendStyle",
      detail: again && hit.legendEntry !== undefined ? { legendEntry: hit.legendEntry } : null,
    };
  }
  if (hit.part === "dataLabels" && hit.label) {
    const { series, point } = hit.label;
    const seriesLabels = currentDetail?.labels?.series === series;
    // The label's own series: a doughnut's rings are its series, each with
    
    // `slice:N` on the pane; another series' (a doughnut's outer ring) is
    // its `point:S:P`.
    if (!seriesLabels) return { area: `series:${series}`, detail: { labels: { series } } };
    return {
      area: attrs.chartType === "pie" && series === 0 ? `slice:${point}` : `point:${series}:${point}`,
      detail: { labels: { series, point } },
    };
  }
  const area = officeChartHitArea(attrs, hit, current);
  // A series or one of its points: the picture marks what is selected
  // (Office's handles on every point of the series, or on the one point).
  const one = area !== null && /^(point|slice):/.test(area);
  return {
    area,
    detail:
      area !== null && hit.plotted
        ? {
            points: {
              series: hit.plotted.series,
              ...(one && hit.plotted.point !== null ? { point: hit.plotted.point } : {}),
            },
          }
        : null,
  };
}

/** The point's name on the pane: its series' and its category's (or X's). */
export function officeChartPointName(
  attrs: ChartBlockAttrsV2,
  seriesIndex: number,
  pointIndex: number,
): string | null {
  if (attrs.data.kind === "category") {
    const series = attrs.data.series[seriesIndex];
    const category = attrs.data.categories[pointIndex];
    return series && category !== undefined ? `${series.name} · ${category}` : null;
  }
  if (attrs.data.kind === "xy") {
    const series = attrs.data.series[seriesIndex];
    const x = series?.xValues[pointIndex];
    return series && pointIndex < series.xValues.length ? `${series.name} · ${String(x ?? "")}` : null;
  }
  return null;
}

/** The DOM event a drawn chart raises when one of its series is clicked. */
export const OFFICE_CHART_HIT_EVENT = "mrai-office-chart-hit";

export interface OfficeChartHitEventDetail {
  readonly hit: OfficeChartHit;
  /** A double click: Office opens the Format pane on what it selects. */
  readonly open: boolean;
  /** A right click: Office selects what it lands on and opens that
   * element's own shortcut menu (a host that asked for these). */
  readonly context?: boolean;
  /** Where the click landed (a title's edit puts its caret there). */
  readonly at?: { clientX: number; clientY: number };
  /** The click landed on the element already selected: Office's second
   * click, which puts a text element (a title) into editing. */
  readonly again?: boolean;
}

/** Report a click on a drawn series from the chart's own element; it
 * bubbles to the host that owns the Format pane. */
export function dispatchOfficeChartHit(
  target: Element,
  hit: OfficeChartHit,
  open: boolean,
  context = false,
  where?: { at?: { clientX: number; clientY: number }; again?: boolean },
): void {
  target.dispatchEvent(
    new CustomEvent<OfficeChartHitEventDetail>(OFFICE_CHART_HIT_EVENT, {
      bubbles: true,
      detail: {
        hit,
        open,
        ...(context ? { context } : {}),
        ...(where?.at ? { at: where.at } : {}),
        ...(where?.again ? { again: true } : {}),
      },
    }),
  );
}

/** The DOM event a drawn chart raises when its selected element was dragged
 * (or sized) by hand: the host keeps the place (`format.elementLayout`). */
export const OFFICE_CHART_PLACE_EVENT = "mrai-office-chart-place";

export interface OfficeChartPlaceEventDetail {
  readonly part: OfficeChartPicturePart;
  /** The element's new box, fractions of the chart. */
  readonly box: { x: number; y: number; w: number; h: number };
  /** One data label dragged: whose it is, and where it was (fractions). */
  readonly label?: { readonly series: number; readonly point: number };
  readonly from?: { x: number; y: number; w: number; h: number };
}

/** Report an element dragged or sized on a chart's picture; it bubbles to the host. */
export function dispatchOfficeChartPlace(
  target: Element,
  part: OfficeChartPicturePart,
  box: { x: number; y: number; w: number; h: number },
  extra: { label?: { series: number; point: number }; from?: { x: number; y: number; w: number; h: number } } = {},
): void {
  target.dispatchEvent(
    new CustomEvent<OfficeChartPlaceEventDetail>(OFFICE_CHART_PLACE_EVENT, {
      bubbles: true,
      detail: { part, box, ...extra },
    }),
  );
}

/**
 * The chart text a click begins to type in, as Office's clicks do: a
 * double-click on a title or an axis title, or a click on the one already
 * selected. Null for any other click.
 */
export function officeChartTextEditOfHit(
  detail: OfficeChartHitEventDetail,
): "title" | "axisX" | "axisY" | `label:${number}:${number}` | null {
  const part = detail.hit.part;
  if (detail.context) return null;
  // One data label selected, clicked again: its words are typed in place.
  if (part === "dataLabels" && detail.hit.label && detail.again)
    return `label:${detail.hit.label.series}:${detail.hit.label.point}`;
  if (part !== "title" && part !== "axisX" && part !== "axisY") return null;
  return detail.open || detail.again ? part : null;
}

/**
 * Report a click on a chart part the host lays out itself, around the
 * picture (a worksheet chart's title and axis titles): the same hit a marked
 * part of the picture reports, so Current Selection -- and a Format pane
 * open on another element -- follows it, as a click on the drawn part does.
 */
export function dispatchOfficeChartPartHit(
  target: Element,
  part: OfficeChartPicturePart,
  open: boolean,
): void {
  dispatchOfficeChartHit(target, { seriesName: "", dataIndex: null, part }, open);
}

/** A part of a drawn chart the picture marks (`ecmeta_part`): the chart
 * title, the horizontal / vertical axis title, the legend, the labels along
 * the horizontal / vertical axis, the series' data labels. */
export type OfficeChartPicturePart =
  | "title"
  | "axisX"
  | "axisY"
  | "legend"
  | "axisLabelX"
  | "axisLabelY"
  | "dataLabels"
  /** The plot area (a click inside it on no series). */
  | "plot";

/** The parts a click selects, the smaller before the larger (an axis title
 * sits beside its labels, a data label over its series). */
const CLICKABLE_PARTS: readonly OfficeChartPicturePart[] = [
  "title",
  "axisX",
  "axisY",
  "dataLabels",
  "axisLabelX",
  "axisLabelY",
  "legend",
];

/**
 * The model axis a drawn axis stands for: the category axis runs along the
 * bottom except on a bar chart, whose bars run sideways (the display plan's
 * `horizontal`); a scatter's two value axes are the model's category (X) and
 * value (Y) records.
 */
export function officeChartAxisRole(
  attrs: { readonly chartType: string },
  physical: "X" | "Y",
): "category" | "value" {
  const horizontal = attrs.chartType === "bar";
  return (physical === "X") !== horizontal ? "category" : "value";
}

/**
 * The Current Selection a click on a marked part makes, as the Format pane
 * names it: the chart title, the legend, an axis (its labels), an axis
 * title, the data labels. Null for a part this chart has no entry for.
 */
export function officeChartPartArea(
  attrs: { readonly chartType: string },
  part: OfficeChartPicturePart,
): string | null {
  switch (part) {
    case "title":
      return "chartTitleStyle";
    case "legend":
      return "legendStyle";
    case "dataLabels":
      return "dataLabelStyle";
    case "plot":
      return "plotAreaStyle";
    case "axisLabelX":
      return `axis:${officeChartAxisRole(attrs, "X")}`;
    case "axisLabelY":
      return `axis:${officeChartAxisRole(attrs, "Y")}`;
    case "axisX":
      return `axisTitle:${officeChartAxisRole(attrs, "X")}`;
    case "axisY":
      return `axisTitle:${officeChartAxisRole(attrs, "Y")}`;
    default:
      return null;
  }
}

/** The drawn part a Current Selection names, for the picture's selection
 * frame (null for a selection the picture draws no one box for). */
export function officeChartAreaPart(
  attrs: { readonly chartType: string },
  area: string | null | undefined,
): OfficeChartPicturePart | null {
  if (!area) return null;
  if (area === "chartTitleStyle") return "title";
  if (area === "legendStyle") return "legend";
  if (area === "dataLabelStyle") return "dataLabels";
  if (area === "plotAreaStyle") return "plot";
  const axis = /^axis:(category|value)$/.exec(area);
  if (axis) return officeChartAxisRole(attrs, "X") === axis[1] ? "axisLabelX" : "axisLabelY";
  const title = /^axisTitle:(category|value)$/.exec(area);
  if (title) return officeChartAxisRole(attrs, "X") === title[1] ? "axisX" : "axisY";
  return null;
}

/**
 * The marked part of a chart picture under a point of the screen: the
 * legend, the chart title or an axis title, found by the boxes of the shapes
 * that draw it (the picture takes no pointer events of its own). Null for
 * the rest of the chart (its plot, axes, gridlines and area).
 */
export function officeChartPicturePartAt(
  root: Element,
  clientX: number,
  clientY: number,
): OfficeChartPicturePart | null {
  for (const part of CLICKABLE_PARTS) {
    for (const element of Array.from(root.querySelectorAll(`svg [ecmeta_part="${part}"]`))) {
      const box = element.getBoundingClientRect();
      if (
        (box.width > 0 || box.height > 0) &&
        clientX >= box.left &&
        clientX <= box.right &&
        clientY >= box.top &&
        clientY <= box.bottom
      )
        return part;
    }
  }
  return null;
}

/** An ECharts click's hit on one of the drawn series, or null. */
export function officeChartHitOfEvent(params: unknown): OfficeChartHit | null {
  if (!params || typeof params !== "object") return null;
  const event = params as { componentType?: unknown; seriesName?: unknown; dataIndex?: unknown };
  if (event.componentType !== "series" || typeof event.seriesName !== "string") return null;
  return {
    seriesName: event.seriesName,
    dataIndex: typeof event.dataIndex === "number" ? event.dataIndex : null,
  };
}

/**
 * What a point of the screen falls on in a chart drawn as ECharts' SVG (the
 * deck's static pictures): the topmost shape that carries ECharts' series
 * and data marks, found by geometry (the picture takes no pointer events of
 * its own), as `{seriesIndex, dataIndex}` of the drawn series.
 */
export function officeChartSvgHitAt(
  svg: SVGSVGElement,
  clientX: number,
  clientY: number,
  /** Only the plotted points: a legend entry (`ecmeta_part`) carries its
   * series' mark too, which a click uses and a tooltip must not. */
  options: { plotOnly?: boolean } = {},
): { seriesIndex: number; dataIndex: number | null } | null {
  const marked = Array.from(svg.querySelectorAll("[ecmeta_series_index]")).reverse();
  // The point in the picture's own viewport px. WebKit's getScreenCTM leaves
  // out the CSS transforms of the page around the picture (a slide is a
  // scaled box), so in Safari a click on a series on a slide landed on
  // another series or none: the viewport point is taken from the picture's
  // laid-out box instead, and each shape's own transform (getCTM, which
  // stays inside the picture) maps it into the shape.
  const box = svg.getBoundingClientRect();
  const viewport =
    box.width > 0 && box.height > 0 && svg.clientWidth > 0 && svg.clientHeight > 0
      ? new DOMPoint(
          ((clientX - box.left) * svg.clientWidth) / box.width,
          ((clientY - box.top) * svg.clientHeight) / box.height,
        )
      : null;
  for (const element of marked) {
    if (!(element instanceof SVGGeometryElement)) continue;
    if (options.plotOnly && element.hasAttribute("ecmeta_part")) continue;
    const own = viewport ? element.getCTM() : null;
    const matrix = own ?? element.getScreenCTM();
    if (!matrix) continue;
    // A shape mid-animation (a bar growing from zero height) has a scale of
    // zero, which has no inverse; it takes no click (regression sweep O-5).
    let inverse: DOMMatrix;
    try {
      inverse = matrix.inverse();
    } catch {
      continue;
    }
    const local = (own && viewport ? viewport : new DOMPoint(clientX, clientY)).matrixTransform(inverse);
    const point = svg.createSVGPoint();
    point.x = local.x;
    point.y = local.y;
    let inside = false;
    try {
      inside = element.isPointInFill(point) || element.isPointInStroke(point);
    } catch {
      inside = false;
    }
    if (!inside) continue;
    const seriesIndex = Number(element.getAttribute("ecmeta_series_index"));
    const dataIndex = element.getAttribute("ecmeta_data_index");
    if (!Number.isInteger(seriesIndex)) continue;
    return {
      seriesIndex,
      dataIndex: dataIndex !== null && Number.isInteger(Number(dataIndex)) ? Number(dataIndex) : null,
    };
  }
  return null;
}
