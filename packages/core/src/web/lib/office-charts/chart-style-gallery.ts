/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Chart Styles galleries, as the three editors apply them.
 *
 * Excel gives each chart type a gallery of numbered styles (Chart Design ▸
 * Chart Styles: Style 1, Style 2, ...; a clustered column's are styles
 * 201-215 and 340, a clustered bar's start at 216). Picking one writes
 * explicit, theme-relative formatting into the chart: the chart and plot
 * areas, the title, the legend, the axes and gridlines, each series and its
 * data labels -- and it turns data labels, the legend and gridlines on or off
 * as the style lays them out. The recipes here were read from workbooks
 * Excel saved with each style applied (`chart-style-gallery.generated.json`,
 * `scripts/generate-chart-style-gallery.py`); a recipe's colours name theme
 * slots and the series' own colour, and are resolved here against the
 * document's theme and the chart's own colours (a Change Colors pick), so a
 * style never changes a chart's hues: Excel's do not either.
 *
 * The result is the chart's ordinary format (`ChartBlockFormat`), so every
 * host draws, saves and exports a styled chart the way it does any formatted
 * one; which style a chart wears is read back by applying each and comparing.
 */
import {
  chartDataLabelPositions,
  chartSupportsDataLabels,
  chartTakesAreaStyles,
  chartTakesSeriesStyles,
  type ChartBlockAttrsV2,
  type ChartBlockFormat,
  type ChartDataLabelPosition,
} from "../../components/console/editor/extensions/chartBlockShared";
import type {
  PresentationChartAreaStyle,
  PresentationChartAxisStyle,
  PresentationChartSeriesStyle,
  PresentationChartSliceStyle,
} from "../presentation-chart-area-style";
import {
  OFFICE_CHART_FILL_PATTERNS,
  OFFICE_CHART_LINE_DASHES,
  officeChartDefaultLineCap,
  type OfficeChartShapePaint,
} from "../office-chart-shape-paint";
import type { OfficeShapeGradient } from "../office-shape-paint";
import { officeStopOpacity } from "../office-shape-paint";
import { officeTransformedHex } from "../office-color-transform";
import type { OfficeChartText } from "./chart-text";
import { officeChartLayoutOf } from "./chart-layout";
import TABLE from "./chart-style-gallery.generated.json";

type Transform = readonly [string, number | null];
/** ["scheme", "tx1", ...] | ["series", null, ...] | ["srgb", "#rrggbb", ...] | ["prst", "black", ...] */
type ColourRef = readonly [string, string | null, readonly Transform[]];
interface GradientRecipe {
  kind?: string;
  angle?: number;
  stops: readonly (readonly [number, ColourRef])[];
  fillToRect?: Record<string, number>;
}
interface FillRecipe {
  solid?: ColourRef;
  none?: true;
  gradient?: GradientRecipe;
  pattern?: { preset: string; fg: ColourRef; bg: ColourRef };
  picture?: true;
}
interface LineRecipe {
  widthPt?: number;
  paint?: FillRecipe;
  dash?: string;
  cap?: string;
  compound?: string;
  join?: string;
}
interface EffectRecipe {
  kind: string;
  blurRad?: number;
  dist?: number;
  dir?: number;
  rad?: number;
  colour?: ColourRef;
}
interface TextRecipe {
  sizePt?: number;
  bold?: boolean;
  italic?: boolean;
  caps?: string;
  spacing?: number;
  rotation?: number;
  fill?: FillRecipe;
  effects?: readonly EffectRecipe[];
}
interface ElementRecipe {
  fill?: FillRecipe;
  line?: LineRecipe;
  effects?: readonly EffectRecipe[];
  text?: TextRecipe;
}
interface LabelRecipe extends ElementRecipe {
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showLegendKey?: boolean;
  showBubbleSize?: boolean;
  position?: string;
  leaderLines?: ElementRecipe;
}
interface SeriesRecipe extends ElementRecipe {
  marker?: ElementRecipe & { symbol?: string | null; size?: number | null };
  dataLabels?: LabelRecipe;
  point?: ElementRecipe;
  explosion?: number;
  smooth?: boolean;
}
interface GroupRecipe {
  type: string;
  gapWidth?: number;
  overlap?: number;
  holeSize?: number;
  firstSliceAng?: number;
  series?: SeriesRecipe;
  hiLowLines?: ElementRecipe;
  dropLines?: ElementRecipe;
  serLines?: ElementRecipe;
  upDownBars?: { gapWidth?: string; up?: ElementRecipe; down?: ElementRecipe };
}
interface AxisRecipe extends ElementRecipe {
  deleted?: boolean;
  axPos?: string;
  majorTickMark?: string;
  minorTickMark?: string;
  majorGridlines?: ElementRecipe;
  minorGridlines?: ElementRecipe;
  title?: ElementRecipe;
}
interface StyleRecipe {
  chartArea?: ElementRecipe | null;
  plotArea?: ElementRecipe | null;
  title?: ElementRecipe | null;
  legend?: (ElementRecipe & { position?: string; overlay?: boolean }) | null;
  axes?: Record<string, readonly AxisRecipe[]>;
  groups?: readonly GroupRecipe[];
  roundedCorners?: boolean;
  styleId?: string;
}

export type ChartStyleGalleryKey =
  | "column"
  | "bar"
  | "stacked"
  | "column3d"
  | "bar3d"
  | "line"
  | "line3d"
  | "area"
  | "area3d"
  | "pie"
  | "pie3d"
  | "ofPie"
  | "doughnut"
  | "scatter"
  | "bubble"
  | "bubble3d"
  | "radar"
  | "stock"
  | "surface"
  | "surfaceTop";

const RAW = TABLE as unknown as {
  galleries: Record<string, readonly number[]>;
  parts: readonly unknown[];
  recipes: Record<string, Record<string, unknown>>;
};

const decoded = new Map<string, StyleRecipe>();
function resolveParts(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(resolveParts);
  if (node && typeof node === "object") {
    const record = node as Record<string, unknown>;
    if (typeof record["@"] === "number" && Object.keys(record).length === 1) {
      return resolveParts(RAW.parts[record["@"] as number]);
    }
    return Object.fromEntries(Object.entries(record).map(([key, value]) => [key, resolveParts(value)]));
  }
  return node;
}

function recipeOf(gallery: ChartStyleGalleryKey, styleId: number): StyleRecipe | null {
  const key = `${gallery}:${styleId}`;
  const hit = decoded.get(key);
  if (hit) return hit;
  const raw = RAW.recipes[gallery]?.[String(styleId)];
  if (raw === undefined) return null;
  const recipe = resolveParts(raw) as StyleRecipe;
  decoded.set(key, recipe);
  return recipe;
}

/** The gallery Excel shows for a chart of this type, or null (no Excel gallery here yet). */
export function chartStyleGalleryKey(chartType: unknown, chartSubtype: unknown): ChartStyleGalleryKey | null {
  const subtype = String(chartSubtype ?? "");
  const threeD = /3d$/i.test(subtype) || subtype === "column3d";
  let key: ChartStyleGalleryKey | null = null;
  switch (chartType) {
    case "column":
      key = threeD ? "column3d" : subtype === "clustered" ? "column" : "stacked";
      break;
    case "bar":
      key = threeD ? "bar3d" : subtype === "clustered" ? "bar" : "stacked";
      break;
    case "line":
      key = subtype === "line3d" ? "line3d" : "line";
      break;
    case "area":
      key = threeD ? "area3d" : "area";
      break;
    case "pie":
      key =
        subtype === "pie3d" ? "pie3d" : subtype === "doughnut" ? "doughnut" : subtype === "pie2d" ? "pie" : "ofPie";
      break;
    case "xy":
      key = subtype === "bubble" ? "bubble" : subtype === "bubble3d" ? "bubble3d" : "scatter";
      break;
    case "radar":
      key = "radar";
      break;
    case "stock":
      key = "stock";
      break;
    case "surface":
      key = subtype === "contour" || subtype === "wireframeContour" ? "surfaceTop" : "surface";
      break;
    default:
      key = null;
  }
  return key && RAW.galleries[key] ? key : null;
}

/** The style ids of a gallery, Style 1 first. */
export function chartStyleGalleryIds(gallery: ChartStyleGalleryKey): readonly number[] {
  return RAW.galleries[gallery] ?? [];
}

/** The most styles any gallery holds (a clustered column's sixteen). */
export const CHART_STYLE_GALLERY_MAX = Math.max(...Object.values(RAW.galleries).map((ids) => ids.length));

/** What a style's colours resolve against. */
export interface ChartStyleColours {
  /** The document theme's accent1..accent6. */
  accents: readonly string[];
  /** The chart's own colour for series (a pie's point) `index`. */
  own: (index: number) => string;
  /** Dark 1 / Light 1 / Dark 2 / Light 2; Office's theme when absent. */
  dk1?: string;
  lt1?: string;
  dk2?: string;
  lt2?: string;
}

const PRESET_COLOURS: Readonly<Record<string, string>> = { black: "#000000", white: "#ffffff" };

function base(ref: ColourRef, colours: ChartStyleColours, own: string): string {
  const [kind, value] = ref;
  if (kind === "series") return own;
  if (kind === "srgb") return value ?? "#000000";
  if (kind === "prst") return PRESET_COLOURS[value ?? ""] ?? "#000000";
  switch (value) {
    case "tx1":
    case "dk1":
      return colours.dk1 ?? "#000000";
    case "bg1":
    case "lt1":
      return colours.lt1 ?? "#ffffff";
    case "tx2":
    case "dk2":
      return colours.dk2 ?? "#0e2841";
    case "bg2":
    case "lt2":
      return colours.lt2 ?? "#e8e8e8";
    default: {
      const slot = /^accent([1-6])$/.exec(value ?? "");
      return slot ? (colours.accents[Number(slot[1]) - 1] ?? "#000000") : "#000000";
    }
  }
}

const ALPHA = new Set(["alpha", "alphaMod", "alphaOff"]);
interface Resolved {
  color: `#${string}`;
  opacity: number;
}
function resolve(ref: ColourRef, colours: ChartStyleColours, own: string): Resolved {
  const transforms = ref[2].filter(([kind]) => !ALPHA.has(kind)).map(([kind, value]) => [kind, value] as [string, number | null]);
  const hex = (officeTransformedHex(base(ref, colours, own), transforms) ?? base(ref, colours, own)).toLowerCase();
  const opacity = officeStopOpacity(ref[2].map(([kind, value]) => [kind, value] as [string, number | null]));
  return { color: hex as `#${string}`, opacity };
}

const opacityOf = (value: number) => Math.round(value * 100_000) / 100_000;

function gradient(recipe: GradientRecipe, colours: ChartStyleColours, own: string): OfficeShapeGradient | null {
  const stops = [...recipe.stops]
    .map(([position, ref]) => {
      const { color, opacity } = resolve(ref, colours, own);
      return { position: Math.min(1, Math.max(0, position)), color, ...(opacity < 1 ? { alpha: opacityOf(opacity) } : {}) };
    })
    .sort((a, b) => a.position - b.position);
  if (stops.length < 2) return null;
  const path = recipe.kind && recipe.kind !== "linear" ? recipe.kind : null;
  return {
    // The model's angle is CSS's (0 = upward), DrawingML's `a:lin@ang` plus a
    // quarter turn. A path gradient (`a:path`) states no `a:lin` and has no
    
    // 0, read back without one, and was held as never landed.
    ...(path ? {} : { angle: ((((recipe.angle ?? 0) + 90) % 360) + 360) % 360 }),
    stops,
    ...(path === "circle" || path === "rect" || path === "shape"
      ? {
          path: {
            kind: path,
            ...(recipe.fillToRect
              ? {
                  fillToRect: Object.fromEntries(
                    Object.entries(recipe.fillToRect).map(([side, value]) => [side, Math.round(value * 100_000)]),
                  ),
                }
              : {}),
          },
        }
      : {}),
  } as OfficeShapeGradient;
}

const DASHES = new Set<string>(OFFICE_CHART_LINE_DASHES);
const PATTERNS = new Set<string>(OFFICE_CHART_FILL_PATTERNS);
const COMPOUNDS: Readonly<Record<string, OfficeChartShapePaint["lineCompound"]>> = {
  dbl: "double",
  thickThin: "thickThin",
  thinThick: "thinThick",
  tri: "triple",
};
const CAPS: Readonly<Record<string, OfficeChartShapePaint["lineCap"]>> = { rnd: "round", sq: "square" };
const EMU_PX = 9_525;

interface Painted {
  style: PresentationChartAreaStyle & { fill?: `#${string}` | "none"; outline?: `#${string}` | "none" };
}

/**
 * One element's fill, line and effects as the chart format states them.
 * ``defaultCap``: the cap the element's line takes when the format states none
 * (`officeChartDefaultLineCap`): a style's cap that is that default is left
 * unstated, as the format pane and the file's reader leave it (a line series'
 * `cap="rnd"` stated as `lineCap: "round"` read back without it, and the
 * restyled chart was declined, 04f77645).
 */
function areaPaint(
  element: ElementRecipe | null | undefined,
  colours: ChartStyleColours,
  own: string,
  { fill = true, line = true, defaultCap = "flat" }: { fill?: boolean; line?: boolean; defaultCap?: string } = {},
): Painted["style"] | undefined {
  if (!element) return undefined;
  const style: Record<string, unknown> = {};
  const shape: OfficeChartShapePaint = {};
  if (fill && element.fill) {
    if (element.fill.none) style.fill = "none";
    else if (element.fill.solid) {
      const { color, opacity } = resolve(element.fill.solid, colours, own);
      style.fill = color;
      if (opacity < 1) shape.fillAlpha = opacityOf(opacity);
    } else if (element.fill.gradient) {
      const g = gradient(element.fill.gradient, colours, own);
      if (g) {
        style.fill = g.stops[0].color;
        shape.fillGradient = g;
      }
    } else if (element.fill.pattern && PATTERNS.has(element.fill.pattern.preset)) {
      const fg = resolve(element.fill.pattern.fg, colours, own).color;
      const bg = resolve(element.fill.pattern.bg, colours, own).color;
      style.fill = fg;
      shape.fillPattern = {
        preset: element.fill.pattern.preset as (typeof OFFICE_CHART_FILL_PATTERNS)[number],
        foreground: fg,
        background: bg,
      };
    }
  }
  const ln = element.line;
  if (line && ln) {
    const paint = ln.paint;
    if (paint?.none) style.outline = "none";
    else if (paint?.solid) {
      const { color, opacity } = resolve(paint.solid, colours, own);
      style.outline = color;
      if (opacity < 1) shape.lineAlpha = opacityOf(opacity);
    } else if (paint?.gradient) {
      const g = gradient(paint.gradient, colours, own);
      if (g) {
        style.outline = g.stops[0].color;
        shape.lineGradient = g;
      }
    }
    if (style.outline !== undefined && style.outline !== "none") {
      if (ln.widthPt !== undefined) shape.outlineWidthPt = Math.min(1584, Math.max(0, ln.widthPt));
      if (ln.dash && ln.dash !== "solid" && DASHES.has(ln.dash)) shape.lineDash = ln.dash as OfficeChartShapePaint["lineDash"];
      if (ln.compound && COMPOUNDS[ln.compound]) shape.lineCompound = COMPOUNDS[ln.compound];
      if (ln.cap && CAPS[ln.cap] && CAPS[ln.cap] !== defaultCap) shape.lineCap = CAPS[ln.cap];
      if (ln.join === "bevel" || ln.join === "miter") shape.lineJoin = ln.join;
    }
  }
  for (const effect of element.effects ?? []) {
    if (effect.kind === "outerShdw" && effect.colour && !shape.shadow) {
      const { color, opacity } = resolve(effect.colour, colours, own);
      shape.shadow = {
        color,
        blur: Math.min(64, (effect.blurRad ?? 0) / EMU_PX),
        distance: Math.min(64, (effect.dist ?? 0) / EMU_PX),
        opacity: Math.min(1, Math.max(0, opacityOf(opacity))),
        ...(effect.dir !== undefined ? { direction: (effect.dir / 60_000) % 360 } : {}),
      };
    } else if (effect.kind === "glow" && style.effect === undefined) {
      style.effect = "glow";
    } else if (effect.kind === "softEdge" && style.effect === undefined) {
      style.effect = "softEdge";
    }
  }
  if (Object.keys(shape).length) style.shapePaint = shape;
  return Object.keys(style).length ? (style as Painted["style"]) : undefined;
}

/** An element's text (size, weight, caps, spacing, colour). */
function textOf(
  text: TextRecipe | undefined,
  colours: ChartStyleColours,
  own: string,
): { text?: OfficeChartText; color?: `#${string}`; shadow?: boolean } {
  if (!text) return {};
  const out: { -readonly [K in keyof OfficeChartText]: OfficeChartText[K] } = {};
  if (text.sizePt !== undefined) out.fontSizePt = text.sizePt;
  if (text.bold) out.bold = true;
  if (text.italic) out.italic = true;
  if (text.caps === "all" || text.caps === "small") out.caps = text.caps;
  if (text.spacing !== undefined && text.spacing !== 0) out.spacingPt = text.spacing;
  let color: `#${string}` | undefined;
  if (text.fill?.solid) color = resolve(text.fill.solid, colours, own).color;
  if (color) out.color = color;
  return {
    ...(Object.keys(out).length ? { text: out } : {}),
    ...(color ? { color } : {}),
    ...(text.effects?.some((effect) => effect.kind === "outerShdw") ? { shadow: true } : {}),
  };
}

const LABEL_POSITIONS: Readonly<Record<string, ChartDataLabelPosition>> = {
  outEnd: "outsideEnd",
  inEnd: "insideEnd",
  ctr: "center",
  inBase: "insideBase",
  bestFit: "bestFit",
  t: "top",
  b: "bottom",
  l: "left",
  r: "right",
};
const LEGEND_POSITIONS: Readonly<Record<string, NonNullable<ChartBlockFormat["legend"]>>> = {
  b: "bottom",
  t: "top",
  r: "right",
  l: "left",
  tr: "right",
};

/** Every format key a Chart Style owns: applying one restates them all. */
export const CHART_STYLE_OWNED_FORMAT_KEYS = [
  "chartAreaStyle",
  "plotAreaStyle",
  "chartTitleStyle",
  "legendStyle",
  "legend",
  "seriesStyles",
  "sliceStyles",
  "axisStyles",
  "axisTitleStyles",
  "gridlineStyle",
  "gridlinesX",
  "gridlinesY",
  "minorGridlinesX",
  "minorGridlinesY",
  "axisX",
  "axisY",
  "dataLabels",
  "dataLabelStyle",
  "dataLabelLeaderLineStyle",
  "dataLabelPosition",
  "dataLabelValue",
  "dataLabelCategoryName",
  "dataLabelSeriesName",
  "dataLabelPercentage",
  "dataLabelBubbleSize",
  "dataLabelLegendKey",
  "dataLabelSeparator",
  "dataLabelLeaderLines",
  "roundedCorners",
  "upBarStyle",
  "downBarStyle",
  "highLowLineStyle",
  "dropLineStyle",
  "seriesLineStyle",
] as const satisfies readonly (keyof ChartBlockFormat)[];

/** Whether an axis lies along the bottom or top (its gridlines run up and down). */
function horizontal(axis: AxisRecipe): boolean {
  return axis.axPos === "b" || axis.axPos === "t";
}

/**
 * The format a chart takes in Excel's style `styleId` of its gallery: what
 * the style owns restated (`CHART_STYLE_OWNED_FORMAT_KEYS`), everything else
 * kept. Null when the chart has no such style.
 */
export function chartStyleFormat(
  attrs: ChartBlockAttrsV2,
  styleId: number,
  colours: ChartStyleColours,
): ChartBlockFormat | null {
  const gallery = chartStyleGalleryKey(attrs.chartType, attrs.chartSubtype);
  if (!gallery || !chartStyleGalleryIds(gallery).includes(styleId)) return null;
  const recipe = recipeOf(gallery, styleId);
  if (!recipe) return null;
  const format: Record<string, unknown> = { ...(attrs.format ?? {}) };
  for (const key of CHART_STYLE_OWNED_FORMAT_KEYS) delete format[key];
  const first = colours.own(0);
  const elements = chartTakesAreaStyles(attrs.chartType);
  if (elements) {
    const chartArea = areaPaint(recipe.chartArea, colours, first);
    if (chartArea) format.chartAreaStyle = chartArea;
    const plotArea = areaPaint(recipe.plotArea, colours, first);
    if (plotArea) format.plotAreaStyle = plotArea;
    if (recipe.roundedCorners) format.roundedCorners = true;
    if (attrs.title && recipe.title) {
      const box = areaPaint(recipe.title, colours, first) ?? {};
      const { text, color } = textOf(recipe.title.text, colours, first);
      const title: Record<string, unknown> = { ...box };
      if (color) title.textFill = color;
      // The title's colour is its `textFill` (what the canvas reads first).
      const rest = text ? Object.fromEntries(Object.entries(text).filter(([key]) => key !== "color")) : {};
      if (Object.keys(rest).length) title.text = rest;
      if (Object.keys(title).length) format.chartTitleStyle = title;
    }
  }
  // The legend: where the style puts it, or none.
  if (recipe.legend === null) format.legend = "none";
  else if (recipe.legend) {
    format.legend = LEGEND_POSITIONS[recipe.legend.position ?? "b"] ?? "bottom";
    if (elements) {
      const box = areaPaint(recipe.legend, colours, first) ?? {};
      const { text } = textOf(recipe.legend.text, colours, first);
      const legend = { ...box, ...(text ? { text } : {}) };
      if (Object.keys(legend).length) format.legendStyle = legend;
    }
  }
  // Axes: their lines and labels, their gridlines, and whether they show.
  const axes = Object.entries(recipe.axes ?? {}).flatMap(([kind, list]) => list.map((axis) => ({ kind, axis })));
  const xy = attrs.chartType === "xy";
  const axisStyles: PresentationChartAxisStyle[] = [];
  let gridlineStyle: Record<string, unknown> | undefined;
  for (const { kind, axis } of axes) {
    if (kind === "serAx") continue;
    const along = horizontal(axis) ? "X" : "Y";
    const model: "category" | "value" = kind === "valAx" && !(xy && horizontal(axis)) ? "value" : "category";
    if (axis.majorGridlines !== undefined) {
      format[`gridlines${along}`] = true;
      gridlineStyle ??= areaPaint(axis.majorGridlines, colours, first, { fill: false });
    } else format[`gridlines${along}`] = false;
    if (axis.minorGridlines !== undefined) format[`minorGridlines${along}`] = true;
    format[`axis${along}`] = !axis.deleted;
    if (elements && !axis.deleted) {
      const line = areaPaint(axis, colours, first, { fill: false });
      const { text } = textOf(axis.text, colours, first);
      const entry: Record<string, unknown> = { axis: model };
      if (line?.outline !== undefined) entry.outline = line.outline;
      if (line?.shapePaint) entry.shapePaint = line.shapePaint;
      if (text) entry.text = text;
      if (Object.keys(entry).length > 1 && !axisStyles.some((item) => item.axis === model)) {
        axisStyles.push(entry as PresentationChartAxisStyle);
      }
    }
  }
  if (elements && axisStyles.length) format.axisStyles = axisStyles;
  if (elements && gridlineStyle) format.gridlineStyle = gridlineStyle;
  // Series: each in its own colour under the style's paint; a pie's points.
  const group = recipe.groups?.[0];
  const series = group?.series;
  const seriesCount = "series" in attrs.data && Array.isArray(attrs.data.series) ? attrs.data.series.length : 0;
  if (series && attrs.chartType === "pie" && attrs.data.kind === "category") {
    const point = series.point ?? series;
    if (attrs.chartSubtype === "pie2d" || attrs.chartSubtype === "doughnut") {
      const rings = attrs.chartSubtype === "doughnut" ? Math.max(1, seriesCount) : 1;
      const slices: PresentationChartSliceStyle[] = [];
      for (let ring = 0; ring < rings; ring += 1) {
        attrs.data.categories.forEach((_, pointIndex) => {
          const paint = areaPaint(point, colours, colours.own(pointIndex));
          if (paint)
            slices.push({ ...(rings > 1 ? { seriesIndex: ring } : {}), pointIndex, ...paint } as PresentationChartSliceStyle);
        });
      }
      if (slices.length) format.sliceStyles = slices;
    }
    if (attrs.chartSubtype === "doughnut" && group?.holeSize !== undefined) format.doughnutHoleSize = group.holeSize;
  } else if (series && chartTakesSeriesStyles(attrs.chartType) && seriesCount > 0) {
    const styles: PresentationChartSeriesStyle[] = [];
    // A line's colour IS its line: a markers-only scatter's "no line" is the
    // subtype's, and stating it would take the series' colour off its markers.
    const lined = attrs.chartType === "line" || attrs.chartType === "xy" || attrs.chartType === "radar" || attrs.chartType === "stock";
    const unlined = lined && series.line?.paint?.none;
    for (let index = 0; index < seriesCount; index += 1) {
      const own = colours.own(index);
      const paint =
        areaPaint(series, colours, own, {
          line: !unlined,
          defaultCap: officeChartDefaultLineCap(`series:${index}`, attrs.chartType),
        }) ?? {};
      const entry: Record<string, unknown> = { seriesIndex: index, ...paint };
      if (series.marker && series.marker.symbol && series.marker.symbol !== "auto") {
        const symbol = series.marker.symbol;
        entry.marker = symbol === "none" ? { symbol: "none" } : { symbol, ...(series.marker.size ? { size: series.marker.size } : {}) };
      }
      if (Object.keys(entry).length > 1) styles.push(entry as PresentationChartSeriesStyle);
    }
    if (styles.length) format.seriesStyles = styles;
  }
  // Data labels: on where the style shows them, in its paint and place.
  if (chartSupportsDataLabels(attrs.chartType)) {
    const labels = series?.dataLabels;
    const shown = Boolean(labels && (labels.showVal || labels.showCatName || labels.showSerName || labels.showPercent));
    format.dataLabels = shown;
    if (labels && shown) {
      if (labels.showVal === false) format.dataLabelValue = false;
      if (labels.showCatName) format.dataLabelCategoryName = true;
      if (labels.showSerName) format.dataLabelSeriesName = true;
      if (labels.showPercent) format.dataLabelPercentage = true;
      if (labels.showLegendKey) format.dataLabelLegendKey = true;
      const position = labels.position ? LABEL_POSITIONS[labels.position] : undefined;
      if (position && chartDataLabelPositions(attrs.chartType, attrs.chartSubtype).includes(position)) {
        format.dataLabelPosition = position;
      }
      if (elements) {
        const box = areaPaint(labels, colours, first) ?? {};
        const { text, shadow } = textOf(labels.text, colours, first);
        const style: Record<string, unknown> = { ...box, ...(text ? { text } : {}) };
        if (shadow && style.textEffect === undefined) style.textEffect = "shadow";
        if (Object.keys(style).length) format.dataLabelStyle = style;
        const leader = areaPaint(labels.leaderLines, colours, first, { fill: false });
        if (leader && attrs.chartType === "pie") format.dataLabelLeaderLineStyle = leader;
      }
    }
  }
  // Column and bar spacing (Excel states it with the style).
  if (group && (group.gapWidth !== undefined || group.overlap !== undefined)) {
    const layout = officeChartLayoutOf({ ...attrs, layout: attrs.format?.layout });
    if (layout) {
      format.layout = {
        ...layout,
        ...(group.gapWidth !== undefined ? { gapWidth: group.gapWidth } : {}),
        overlap: group.overlap ?? (attrs.chartSubtype === "clustered" ? 0 : layout.overlap),
      };
      // The style states the spacing: the chart's own keys go (one record).
      delete format.barGapWidth;
      delete format.barOverlap;
    }
  }
  // Stock and line extras: high-low lines, up/down bars, drop lines.
  if (elements && group) {
    const hiLow = areaPaint(group.hiLowLines, colours, first, { fill: false });
    if (hiLow && (attrs.chartType === "stock" || attrs.format?.highLowLines)) format.highLowLineStyle = hiLow;
    if (group.upDownBars && attrs.chartType === "stock") {
      const up = areaPaint(group.upDownBars.up, colours, first);
      const down = areaPaint(group.upDownBars.down, colours, first);
      if (up) format.upBarStyle = up;
      if (down) format.downBarStyle = down;
    }
    const drop = areaPaint(group.dropLines, colours, first, { fill: false });
    if (drop && attrs.format?.dropLines) format.dropLineStyle = drop;
  }
  return format as ChartBlockFormat;
}
