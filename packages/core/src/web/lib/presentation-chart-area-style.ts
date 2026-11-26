/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeChartText, officeChartTextCss, type OfficeChartText } from "./office-charts/chart-text";
import type { CSSProperties } from "react";
import { normalizeOfficeTextPaint, officeTextPaintCss, patchOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";
import { normalizeOfficeChartShapePaint, officeChartBorderStyleCss, officeChartColorWithOpacity, officeChartShapePaintCss, OFFICE_CHART_LINE_DASH_PATTERN, type OfficeChartShapePaint } from "./office-chart-shape-paint";
import { officeDashLengthsWithCaps } from "./office-stroke-dash";
import type { OfficeStrokeCompound } from "./office-stroke-detail";
import { MAX_CHART_POINTS, MAX_CHART_POINTS_PER_SERIES, MAX_CHART_SERIES } from "./office-charts/chart-limits.generated";

const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

export type PresentationChartAreaTarget = "chart-area" | "plot-area";
export type PresentationChartFormatTarget =
  | PresentationChartAreaTarget
  | "chart-title"
  | "legend"
  | "axis:category"
  | "axis:value"
  | "axis-title:category"
  | "axis-title:value"
  | "gridlines"
  | "trendline"
  | "data-table"
  | "data-labels"
  | `slice:${number}`
  | `series:${number}`;
export type PresentationChartAreaPaint = "none" | `#${string}`;
export const PRESENTATION_CHART_EFFECT_PRESETS = [
  "shadow",
  "reflection",
  "glow",
  "softEdge",
  "bevel",
  "threeD",
] as const;
export type PresentationChartEffectPreset =
  (typeof PRESENTATION_CHART_EFFECT_PRESETS)[number];
export const PRESENTATION_CHART_TEXT_EFFECT_PRESETS = [
  "shadow",
  "reflection",
  "glow",
  "softEdge",
] as const;
export type PresentationChartTextEffectPreset =
  (typeof PRESENTATION_CHART_TEXT_EFFECT_PRESETS)[number];

export type PresentationChartAreaStyle = Readonly<{
  shapePaint?: OfficeChartShapePaint;
  textPaint?: OfficeTextPaint;
  fill?: PresentationChartAreaPaint;
  outline?: PresentationChartAreaPaint;
  effect?: PresentationChartEffectPreset;
  textFill?: PresentationChartAreaPaint;
  textOutline?: PresentationChartAreaPaint;
  textEffect?: PresentationChartTextEffectPreset;
  /**
   * The element's own text (face, size, bold, italic, underline, colour):
   * the title's `c:title/c:txPr`, the legend's `c:legend/c:txPr`, the data
   * labels' `c:dLbls/c:txPr`. It overrides the chart's `format.text` for
   * that element alone, as Office formats the one element selected (Home ▸
   * Font on a selected chart title sizes the title, not the axes).
   */
  text?: OfficeChartText;
}>;
/**
 * Format Data Series ▸ Marker ▸ Marker Options ▸ Built-in ▸ Type, as ECMA-376
 * `ST_MarkerStyle` spells it. `auto` is the ABSENT marker (the subtype's own)
 * and `picture` a fill the model does not state, so neither is a value. The
 * twin of `document_charts.CHART_MARKER_SYMBOLS` and of the worksheet's
 * `SPREADSHEET_LINE_MARKER_SYMBOLS`.
 */
export const CHART_MARKER_SYMBOLS = [
  "none",
  "circle",
  "dash",
  "diamond",
  "dot",
  "plus",
  "square",
  "star",
  "triangle",
  "x",
] as const;
export type ChartMarkerSymbol = (typeof CHART_MARKER_SYMBOLS)[number];
/** `ST_MarkerSize`: an unsigned byte from 2 to 72, in points. */
export const MIN_CHART_MARKER_SIZE = 2;
export const MAX_CHART_MARKER_SIZE = 72;

/**
 * One series' own `c:marker`: its symbol and, for a visible one, its size in
 * points. Absent is the automatic marker its subtype draws (Line with Markers:
 * a marker; Line: none). The twin of `document_charts.ChartSeriesMarker`.
 */
export type ChartSeriesMarker = Readonly<{
  symbol: ChartMarkerSymbol;
  size?: number;
  /**
   * Marker > Fill and Marker > Border (`c:marker/c:spPr`): a hex colour or
   * "none" for a visible marker. Absent is the series' own paint.
   */
  fill?: string;
  outline?: string;
  





  shapePaint?: ChartMarkerShapePaint;
}>;

/**
 * A marker's own shape paint: its 3-D Format and its Border > Width
 * (`a:ln@w`, `outlineWidthPt`). Without a width the marker's border is
 * Office's 0.75 pt, never its series' line width
 * (`officeChartMarkerBorderWidthPx`).
 */
export type ChartMarkerShapePaint = Pick<OfficeChartShapePaint, "bevel" | "format3d" | "outlineWidthPt">;

const MARKER_PAINT = /^(#[0-9a-f]{6}|none)$/;
/** What a marker's own shape paint may state (`ChartSeriesMarker.shapePaint`). */
const MARKER_SHAPE_PAINT_KEYS: ReadonlySet<string> = new Set(["bevel", "format3d", "outlineWidthPt"]);

function normalizeChartMarkerShapePaint(value: unknown): ChartMarkerShapePaint | undefined | null {
  if (value === undefined) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  if (Object.keys(value).some((key) => !MARKER_SHAPE_PAINT_KEYS.has(key))) return null;
  try {
    const paint = normalizeOfficeChartShapePaint(value);
    return paint
      ? {
          ...(paint.bevel ? { bevel: paint.bevel } : {}),
          ...(paint.format3d ? { format3d: paint.format3d } : {}),
          ...(paint.outlineWidthPt !== undefined ? { outlineWidthPt: paint.outlineWidthPt } : {}),
        }
      : null;
  } catch {
    return null;
  }
}

export function normalizeChartSeriesMarker(
  value: unknown,
): ChartSeriesMarker | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  if (Object.keys(record).some((key) => key !== "symbol" && key !== "size" && key !== "fill" && key !== "outline" && key !== "shapePaint")) {
    return undefined;
  }
  const symbol = record.symbol;
  if (!(CHART_MARKER_SYMBOLS as readonly unknown[]).includes(symbol)) return undefined;
  const paint: { fill?: string; outline?: string; shapePaint?: ChartMarkerShapePaint } = {};
  for (const key of ["fill", "outline"] as const) {
    const raw = record[key];
    if (raw === undefined) continue;
    if (symbol === "none" || typeof raw !== "string" || !MARKER_PAINT.test(raw.toLowerCase())) return undefined;
    paint[key] = raw.toLowerCase();
  }
  const threeD = normalizeChartMarkerShapePaint(record.shapePaint);
  if (threeD === null || (threeD && symbol === "none")) return undefined;
  if (threeD) paint.shapePaint = threeD;
  if (record.size === undefined) return { symbol: symbol as ChartMarkerSymbol, ...paint };
  const size = record.size;
  if (
    symbol === "none" ||
    typeof size !== "number" ||
    !Number.isInteger(size) ||
    size < MIN_CHART_MARKER_SIZE ||
    size > MAX_CHART_MARKER_SIZE
  ) {
    return undefined;
  }
  return { symbol: symbol as ChartMarkerSymbol, size, ...paint };
}

export type PresentationChartSeriesStyle = Readonly<{
  shapePaint?: OfficeChartShapePaint;
  seriesIndex: number;
  /** "none" is Fill ▸ No fill: a hollow bar, column or area. */
  fill?: PresentationChartAreaPaint;
  outline?: PresentationChartAreaPaint;
  effect?: PresentationChartEffectPreset;
  /** Format Data Series ▸ Marker: this series' own `c:marker`, on a series
   * that draws one (`chartSeriesTakesMarker`). */
  marker?: ChartSeriesMarker;
  /**
   * ▸ Line ▸ Smoothed line: this series' own `c:smooth`. It OVERRIDES the
   * chart-wide `format.curve` (and a scatter's smooth subtype) for this
   * series alone, on a series that takes one (`chartSeriesTakesSmooth`).
   * A line or scatter series' width is `shapePaint.outlineWidthPt` and its
   * colour `outline`: its line IS its `c:spPr/a:ln`.
   */
  smooth?: boolean;
  /** ▸ Series Options ▸ Column shape of this series alone (its `c:shape`), on
   * a 3-D column or bar chart: it OVERRIDES the chart-wide `format.barShape`. */
  barShape?: ChartBarShape;
}>;

/** Format Data Series ▸ Series Options ▸ Column shape of a 3-D column or bar
 * chart (ECMA-376 ST_Shape; VBA `XlBarShape`): Box, Cylinder, Full Cone
 * (`cone`, its point at the value), Partial Cone (`coneToMax`, a cone whose
 * point is the value axis' maximum, cut off at the value), Full and Partial
 * Pyramid. The twin of `document_charts.CHART_BAR_SHAPES`. */
export const CHART_BAR_SHAPES = ["box", "cylinder", "cone", "coneToMax", "pyramid", "pyramidToMax"] as const;
export type ChartBarShape = (typeof CHART_BAR_SHAPES)[number];
export function isChartBarShape(value: unknown): value is ChartBarShape {
  return typeof value === "string" && (CHART_BAR_SHAPES as readonly string[]).includes(value);
}
export type PresentationChartAxisStyle = Readonly<{
  shapePaint?: OfficeChartShapePaint;
  

  axis: "category" | "value" | "secondaryValue";
  outline?: PresentationChartAreaPaint;
  /** Format Axis ▸ Fill: a solid colour behind each of the axis' labels
   * (never "none": an axis without one states nothing). */
  fill?: `#${string}`;
  effect?: PresentationChartEffectPreset;
  /** The axis labels' own text (the axis's `c:txPr`), over the chart's. */
  text?: OfficeChartText;
}>;

/**
 * The part of an element's style an axis keeps (Format Axis ▸ its line, its
 * labels' fill, effect and text): a Format pane or ribbon edit made on the
 * axis is written as the axis' own style, never an area style's other keys,
 * and a No Fill states no fill.
 */
export function presentationChartAxisStylePart(
  style: PresentationChartAreaStyle,
): Omit<PresentationChartAxisStyle, "axis"> {
  return {
    ...(style.shapePaint ? { shapePaint: style.shapePaint } : {}),
    ...(style.outline === undefined ? {} : { outline: style.outline }),
    ...(style.fill === undefined || style.fill === "none" ? {} : { fill: style.fill }),
    ...(style.effect === undefined ? {} : { effect: style.effect }),
    ...(style.text ? { text: style.text } : {}),
  };
}
/** An axis title's own style: its box paint, its text paint and its text
 * (the axis's `c:title`), by the model axis it titles. */
export type PresentationChartAxisTitleStyle = Readonly<{
  axis: "category" | "value";
  /** Never stated on a title (its box takes plain paint); typed for the
   * hosts that read one style record for every selection. */
  shapePaint?: OfficeChartShapePaint;
  fill?: PresentationChartAreaPaint;
  outline?: PresentationChartAreaPaint;
  effect?: PresentationChartEffectPreset;
  textFill?: PresentationChartAreaPaint;
  textOutline?: PresentationChartAreaPaint;
  textEffect?: PresentationChartTextEffectPreset;
  textPaint?: OfficeTextPaint;
  text?: OfficeChartText;
}>;
export type PresentationChartSliceStyle = Readonly<{
  shapePaint?: OfficeChartShapePaint;
  /**
   * The series whose point this is (a doughnut's rings are its series); the
   * first series when absent. The twin of `ChartSliceStyle.series_index`.
   */
  seriesIndex?: number;
  pointIndex: number;
  fill?: PresentationChartAreaPaint;
  outline?: PresentationChartAreaPaint;
  effect?: PresentationChartEffectPreset;
  /** Format Data Point ▸ Point Explosion, percent of the radius (1–400). */
  explosion?: number;
  /** Format Data Point ▸ Marker on a line or X Y point (`c:dPt/c:marker`). */
  marker?: ChartSeriesMarker;
}>;
/** Every point of every ring of a doughnut: at most one style per value of
 * the chart (the shared limits table's `MAX_CHART_POINTS`). */
export const MAX_PRESENTATION_CHART_SLICE_STYLES = MAX_CHART_POINTS;
export const MAX_PRESENTATION_CHART_POINT_EXPLOSION = 400;

export function normalizePresentationChartEffectPreset(
  value: unknown,
): PresentationChartEffectPreset | undefined {
  return (PRESENTATION_CHART_EFFECT_PRESETS as readonly unknown[]).includes(
    value,
  )
    ? (value as PresentationChartEffectPreset)
    : undefined;
}

export function normalizePresentationChartTextEffectPreset(
  value: unknown,
): PresentationChartTextEffectPreset | undefined {
  return (
    PRESENTATION_CHART_TEXT_EFFECT_PRESETS as readonly unknown[]
  ).includes(value)
    ? (value as PresentationChartTextEffectPreset)
    : undefined;
}

export function normalizePresentationChartAreaPaint(
  value: unknown,
): PresentationChartAreaPaint | undefined {
  if (value === "none") return "none";
  return typeof value === "string" && HEX_COLOR_RE.test(value)
    ? (value.toLowerCase() as PresentationChartAreaPaint)
    : undefined;
}

export function normalizePresentationChartAreaStyle(
  value: unknown,
): PresentationChartAreaStyle | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  let shapePaint: OfficeChartShapePaint | undefined, textPaint: OfficeTextPaint | undefined;
  try {
    // A paint whose last extra an edit took away (a gradient recoloured
    // solid, say) states nothing more: it is no paint.
    const rawPaint = record.shapePaint;
    shapePaint =
      rawPaint && typeof rawPaint === "object" && !Array.isArray(rawPaint) && !Object.keys(rawPaint).length
        ? undefined
        : normalizeOfficeChartShapePaint(rawPaint);
    textPaint = normalizeOfficeTextPaint(record.textPaint);
  } catch { return undefined; }
  const fill = normalizePresentationChartAreaPaint(record.fill);
  const outline = normalizePresentationChartAreaPaint(record.outline);
  const effect = normalizePresentationChartEffectPreset(record.effect);
  const textFill = normalizePresentationChartAreaPaint(record.textFill);
  const textOutline = normalizePresentationChartAreaPaint(record.textOutline);
  const textEffect = normalizePresentationChartTextEffectPreset(
    record.textEffect,
  );
  // A stated text that is not one (an unknown key, a size out of range)
  // makes the style no style: the validator then refuses it by name.
  const text = record.text === undefined ? undefined : normalizeOfficeChartText(record.text);
  if (record.text !== undefined && !text) return undefined;
  if (shapePaint && ((shapePaint.lineGradient && outline === undefined) ||
      // A transparency is the solid colour's, so it rides on one.
      (shapePaint.fillAlpha !== undefined && (fill === undefined || fill === "none")) ||
      (shapePaint.lineAlpha !== undefined && (outline === undefined || outline === "none")) ||
      (shapePaint.fillGradient && fill !== shapePaint.fillGradient.stops[0].color) ||
      (shapePaint.fillPattern && fill !== shapePaint.fillPattern.foreground) ||
      // A picture fill is the element's whole fill: no solid colour beside it.
      (shapePaint.fillPicture && fill !== undefined) ||
      (shapePaint.lineGradient && outline !== shapePaint.lineGradient.stops[0].color) ||
      ((shapePaint.shadow || shapePaint.glow || shapePaint.softEdge) && effect !== undefined))) return undefined;
  if (textPaint && (textFill !== undefined || textOutline !== undefined || textEffect !== undefined)) return undefined;
  return fill === undefined &&
    outline === undefined &&
    effect === undefined &&
    textFill === undefined &&
    textOutline === undefined &&
    textEffect === undefined && shapePaint === undefined && textPaint === undefined &&
    text === undefined
    ? undefined
    : {
        ...(fill === undefined ? {} : { fill }),
        ...(outline === undefined ? {} : { outline }),
        ...(effect === undefined ? {} : { effect }),
        ...(textFill === undefined ? {} : { textFill }),
        ...(textOutline === undefined ? {} : { textOutline }),
        ...(textEffect === undefined ? {} : { textEffect }),
        ...(shapePaint ? { shapePaint } : {}),
        ...(textPaint ? { textPaint } : {}),
        ...(text ? { text } : {}),
      };
}

export function presentationChartTextStyleCss(
  value: unknown,
): CSSProperties {
  const style = normalizePresentationChartAreaStyle(value);
  if (!style) return {};
  // The element's own font (size, face, weight, slant) under its paint.
  const font = officeChartTextCss(style.text);
  if (style.textPaint) return { ...font, ...officeTextPaintCss(style.textPaint) };
  const textEffect = style.textEffect;
  return {
    ...font,
    ...(style.textFill === undefined
      ? {}
      : { color: style.textFill === "none" ? "transparent" : style.textFill }),
    ...(style.textOutline === undefined
      ? {}
      : style.textOutline === "none"
        ? { WebkitTextStroke: "0 transparent" }
        : { WebkitTextStroke: `1px ${style.textOutline}` }),
    ...(textEffect === "shadow"
      ? { textShadow: "2px 2px 3px rgb(15 23 42 / 38%)" }
      : textEffect === "reflection"
        ? {
            WebkitBoxReflect:
              "below 1px linear-gradient(transparent 55%, rgb(15 23 42 / 18%))",
          }
        : textEffect === "glow"
          ? { textShadow: "0 0 6px rgb(59 130 246 / 78%)" }
          : textEffect === "softEdge"
            ? { textShadow: "0 0 4px rgb(255 255 255 / 92%)" }
            : {}),
  } as CSSProperties;
}

export function presentationChartEffectCss(
  value: unknown,
): CSSProperties {
  const effect = normalizePresentationChartEffectPreset(value);
  if (!effect) return {};
  if (effect === "shadow") {
    return { filter: "drop-shadow(3px 3px 4px rgb(15 23 42 / 34%))" };
  }
  if (effect === "reflection") {
    return {
      WebkitBoxReflect:
        "below 2px linear-gradient(transparent 58%, rgb(15 23 42 / 16%))",
    } as CSSProperties;
  }
  if (effect === "glow") {
    return { filter: "drop-shadow(0 0 5px rgb(59 130 246 / 72%))" };
  }
  if (effect === "softEdge") {
    return { filter: "drop-shadow(0 0 4px rgb(255 255 255 / 90%))" };
  }
  if (effect === "bevel") {
    return {
      boxShadow:
        "inset 2px 2px 2px rgb(255 255 255 / 68%), inset -2px -2px 2px rgb(15 23 42 / 28%)",
    };
  }
  return {
    boxShadow:
      "inset 2px 2px 2px rgb(255 255 255 / 58%), 3px 4px 5px rgb(15 23 42 / 26%)",
  };
}

/**
 * Whether a chart element can draw an effect. A reflection is a mirrored
 * copy of a box of its own, which only the chart area has (it is drawn on
 * the chart's frame); a series, the plot area and the other parts are marks
 * inside one drawing and cannot be mirrored one by one. Office greys the
 * effects an element does not take ("Preset, reflection, and bevel effects
 * are not available for all chart elements"), and so does the Effects menu.
 */
export function presentationChartAreaDrawsEffect(
  area: string,
  effect: PresentationChartEffectPreset,
): boolean {
  return (
    effect !== "reflection" || area === "chartAreaStyle" || area === "chart-area"
  );
}

export function presentationChartEffectEChartsStyle(
  value: unknown,
): Readonly<Record<string, unknown>> {
  const effect = normalizePresentationChartEffectPreset(value);
  if (!effect) return {};
  if (effect === "shadow") {
    return {
      shadowBlur: 8,
      shadowColor: "rgba(15, 23, 42, 0.34)",
      shadowOffsetX: 3,
      shadowOffsetY: 3,
    };
  }
  // A mark inside the drawing cannot be mirrored; a grey shadow under it is
  // not a reflection, so a stored one draws nothing (PPT-B340).
  if (effect === "reflection") return {};
  if (effect === "glow") {
    return { shadowBlur: 12, shadowColor: "rgba(59, 130, 246, 0.72)" };
  }
  if (effect === "softEdge") {
    return { shadowBlur: 7, shadowColor: "rgba(255, 255, 255, 0.9)" };
  }
  if (effect === "bevel") {
    return {
      shadowBlur: 2,
      shadowColor: "rgba(15, 23, 42, 0.3)",
      shadowOffsetX: 1,
      shadowOffsetY: 2,
    };
  }
  return {
    shadowBlur: 5,
    shadowColor: "rgba(15, 23, 42, 0.32)",
    shadowOffsetX: 3,
    shadowOffsetY: 4,
  };
}

/**
 * The paint a chart title's words are drawn in as WordArt: its WordArt style
 * (`textPaint`), or -- the deck's picture draws those as one paint too -- its
 * text fill, outline and effect. Undefined for a title in plain text. One
 * reading for the picture and for every box the title is typed in, so the
 * words do not lose their WordArt the moment editing begins.
 */
export function presentationChartTitleTextPaint(value: unknown): OfficeTextPaint | undefined {
  const style = normalizePresentationChartAreaStyle(value);
  if (!style) return undefined;
  if (style.textPaint) return style.textPaint;
  if (style.textFill === undefined && style.textOutline === undefined && style.textEffect === undefined)
    return undefined;
  const paint: OfficeTextPaint = {
    fill: style.textFill ?? "#000000",
    outline: style.textOutline ?? "none",
    outlineWidthPt: style.textOutline && style.textOutline !== "none" ? 0.75 : 0,
  };
  return style.textEffect && style.textEffect !== "softEdge"
    ? patchOfficeTextPaint(paint, { effect: style.textEffect })
    : paint;
}

/**
 * A chart title's words drawn as HTML (a host's own title row, the stand-in
 * while the picture is drawn) in the title's WordArt. Its box format is
 * `presentationChartAreaStyleCss`'s, on the element around the words: a
 * WordArt gradient is clipped to the words, and would clip the box's fill
 * with it.
 */
export function presentationChartTitleWordsCss(value: unknown): CSSProperties {
  const paint = presentationChartTitleTextPaint(value);
  return paint ? officeTextPaintCss(paint) : {};
}








export interface PresentationChartAreaOutline {
  readonly color: string;
  /** The line's width in CSS px (a stated 0 or sub-pixel line: 1 px). */
  readonly width: number;
  /** Ink bands as [from, to] fractions of the width, outside in. */
  readonly bands: readonly (readonly [number, number])[];
  readonly dashArray?: string;
  readonly cap: "butt" | "round" | "square";
  readonly join: "round" | "bevel" | "miter";
}

const COMPOUND_BANDS: Record<OfficeStrokeCompound, readonly (readonly [number, number])[]> = {
  // `officeCompoundStrokeLayers`' measured layers, as bands from the outside.
  double: [[0, 1 / 3], [2 / 3, 1]],
  thinThick: [[0, 0.2], [0.4, 1]],
  thickThin: [[0, 0.6], [0.8, 1]],
  triple: [[0, 1 / 6], [1 / 3, 2 / 3], [5 / 6, 1]],
};

export function presentationChartAreaOutline(value: unknown): PresentationChartAreaOutline | undefined {
  const style = normalizePresentationChartAreaStyle(value);
  if (!style || style.outline === undefined || style.outline === "none") return undefined;
  const paint = style.shapePaint;
  const widthPt = paint?.outlineWidthPt ?? 0.75;
  const width = Math.max(1, (widthPt * 4) / 3);
  const cap = paint?.lineCap === "round" || paint?.lineCap === "square" ? paint.lineCap : "butt";
  // Office's default join for a frame is a round one.
  const join = paint?.lineJoin ?? "round";
  const dash =
    paint?.lineDash && widthPt > 0
      ? officeDashLengthsWithCaps(OFFICE_CHART_LINE_DASH_PATTERN[paint.lineDash], width, cap !== "butt")
      : undefined;
  return {
    color: officeChartColorWithOpacity(style.outline, paint?.lineAlpha),
    width,
    bands: paint?.lineCompound ? COMPOUND_BANDS[paint.lineCompound] : [[0, 1]],
    ...(dash ? { dashArray: dash.map((n) => Math.round(n * 1000) / 1000).join(" ") } : {}),
    cap,
    join,
  };
}

/**
 * `presentationChartAreaStyleCss` without the frame's line, for a host that
 * draws it as `presentationChartAreaOutline` (`OfficeChartAreaOutline`).
 */
export function presentationChartAreaStyleCssWithoutLine(
  value: unknown,
  options: { readonly shadowBlurDeviation?: number } = {},
): CSSProperties {
  const css = { ...presentationChartAreaStyleCss(value, options) } as Record<string, unknown>;
  for (const key of Object.keys(css)) if (/^border(?!Radius)/.test(key)) delete css[key];
  return css as CSSProperties;
}

export function presentationChartAreaStyleCss(
  value: unknown,
  /** The Gaussian deviation its outer shadow's blur radius draws, per unit of it (a slide's chart: `POWERPOINT_SHADOW_BLUR_DEVIATION`). */
  options: { readonly shadowBlurDeviation?: number } = {},
): CSSProperties {
  const style = normalizePresentationChartAreaStyle(value);
  if (!style) return {};
  const paint = style.shapePaint;
  const borderStyle = officeChartBorderStyleCss(paint);
  // A soft-edged solid box is drawn by the chart's picture, faded in from
  // its edge (`chart-render-server`): the frame paints no fill or line of
  // its own over it.
  if (
    paint?.softEdge &&
    style.fill !== undefined &&
    style.fill !== "none" &&
    !paint.fillGradient &&
    !paint.fillPattern &&
    !paint.fillPicture
  ) {
    const { fill: _fill, outline: _outline, shapePaint, ...rest } = style;
    // Its effects stay the frame's (a shadow, a glow); its fill and line are the picture's.
    const effects = Object.fromEntries(
      Object.entries(shapePaint ?? {}).filter(([key]) => ["shadow", "glow", "innerShadow", "bevel", "format3d"].includes(key)),
    );
    return {
      ...presentationChartAreaStyleCss({ ...rest, ...(Object.keys(effects).length ? { shapePaint: effects } : {}) }, options),
      backgroundColor: "transparent",
      border: "none",
    };
  }
  return {
    boxSizing: "border-box",
    ...(style.fill === undefined
      ? {}
      : {
          backgroundColor:
            style.fill === "none" ? "transparent" : officeChartColorWithOpacity(style.fill, paint?.fillAlpha),
        }),
    ...(style.outline === undefined
      ? {}
      : style.outline === "none"
        ? { border: "none" }
        : {
            borderColor: officeChartColorWithOpacity(style.outline, paint?.lineAlpha),
            borderStyle: "solid",
            borderWidth: 1,
          }),
    ...presentationChartEffectCss(style.effect),
    ...presentationChartTextStyleCss({ ...style, textPaint: undefined }),
    ...(paint ? officeChartShapePaintCss(paint, options) : {}),
    // A compound or dashed frame, once the frame is drawn at all.
    ...(borderStyle && style.outline !== "none" ? { borderStyle } : {}),
  };
}

export function normalizePresentationChartSeriesStyles(
  value: unknown,
): readonly PresentationChartSeriesStyle[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_CHART_SERIES) {
    return undefined;
  }
  const seen = new Set<number>();
  const styles: PresentationChartSeriesStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return undefined;
    const record = entry as Record<string, unknown>;
    const seriesIndex = record.seriesIndex;
    const shapePaint = normalizePresentationChartAreaStyle(record)?.shapePaint;
    if (record.shapePaint !== undefined && !shapePaint) return undefined;
    const fill = normalizePresentationChartAreaPaint(record.fill);
    const outline = normalizePresentationChartAreaPaint(record.outline);
    const effect = normalizePresentationChartEffectPreset(record.effect);
    const marker =
      record.marker === undefined ? undefined : normalizeChartSeriesMarker(record.marker);
    if (record.marker !== undefined && !marker) return undefined;
    if (record.smooth !== undefined && typeof record.smooth !== "boolean") return undefined;
    const smooth = record.smooth as boolean | undefined;
    if (record.barShape !== undefined && !isChartBarShape(record.barShape)) return undefined;
    const barShape = record.barShape as ChartBarShape | undefined;
    if (
      !Number.isInteger(seriesIndex) ||
      (seriesIndex as number) < 0 ||
      (seriesIndex as number) >= MAX_CHART_SERIES ||
      (fill === undefined &&
        outline === undefined &&
        effect === undefined &&
        marker === undefined &&
        smooth === undefined &&
        barShape === undefined &&
        !shapePaint) ||
      seen.has(seriesIndex as number)
    ) {
      return undefined;
    }
    seen.add(seriesIndex as number);
    styles.push({
      seriesIndex: seriesIndex as number,
      ...(shapePaint ? { shapePaint } : {}),
      ...(fill === undefined ? {} : { fill }),
      ...(outline === undefined ? {} : { outline }),
      ...(effect === undefined ? {} : { effect }),
      ...(marker === undefined ? {} : { marker }),
      ...(smooth === undefined ? {} : { smooth }),
      ...(barShape === undefined ? {} : { barShape }),
    });
  }
  return styles.sort((left, right) => left.seriesIndex - right.seriesIndex);
}

export function presentationChartSeriesColorCss(
  value: unknown,
): CSSProperties {
  const styles = normalizePresentationChartSeriesStyles(value);
  if (!styles) return {};
  const css: Record<string, string> = {};
  for (const style of styles) {
    if (style.fill) {
      css[`--mrai-chart-series-${style.seriesIndex + 1}`] =
        style.fill === "none" ? "transparent" : style.fill;
    }
    if (style.outline) {
      css[`--mrai-chart-series-outline-${style.seriesIndex + 1}`] =
        style.outline === "none" ? "transparent" : style.outline;
    }
  }
  return css as CSSProperties;
}

export function presentationChartFormatSeriesIndex(
  value: PresentationChartFormatTarget,
): number | null {
  const match = /^series:([0-9]+)$/.exec(value);
  if (!match) return null;
  const index = Number(match[1]);
  return Number.isInteger(index) && index >= 0 && index < MAX_CHART_SERIES ? index : null;
}

export function presentationChartFormatSliceIndex(
  value: PresentationChartFormatTarget,
): number | null {
  const match = /^slice:([0-9]+)$/.exec(value);
  if (!match) return null;
  const index = Number(match[1]);
  return Number.isInteger(index) && index >= 0 && index < MAX_CHART_POINTS_PER_SERIES ? index : null;
}

export function normalizePresentationChartSliceStyles(
  value: unknown,
): readonly PresentationChartSliceStyle[] | undefined {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > MAX_PRESENTATION_CHART_SLICE_STYLES
  ) {
    return undefined;
  }
  const seen = new Set<string>();
  const styles: PresentationChartSliceStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return undefined;
    const record = entry as Record<string, unknown>;
    const pointIndex = record.pointIndex;
    const seriesIndex = record.seriesIndex ?? 0;
    const explosion = record.explosion;
    const shapePaint = normalizePresentationChartAreaStyle(record)?.shapePaint;
    if (record.shapePaint !== undefined && !shapePaint) return undefined;
    const fill = normalizePresentationChartAreaPaint(record.fill);
    const outline = normalizePresentationChartAreaPaint(record.outline);
    const effect = normalizePresentationChartEffectPreset(record.effect);
    const marker = record.marker === undefined ? undefined : normalizeChartSeriesMarker(record.marker);
    if (record.marker !== undefined && !marker) return undefined;
    const key = `${String(seriesIndex)}:${String(pointIndex)}`;
    if (
      !Number.isInteger(pointIndex) ||
      (pointIndex as number) < 0 ||
      (pointIndex as number) >= MAX_CHART_POINTS_PER_SERIES ||
      !Number.isInteger(seriesIndex) ||
      (seriesIndex as number) < 0 ||
      (seriesIndex as number) >= MAX_CHART_SERIES ||
      (record.seriesIndex !== undefined && record.seriesIndex === 0) ||
      (explosion !== undefined &&
        (!Number.isInteger(explosion) ||
          (explosion as number) < 1 ||
          (explosion as number) > MAX_PRESENTATION_CHART_POINT_EXPLOSION)) ||
      (fill === undefined &&
        outline === undefined &&
        effect === undefined &&
        explosion === undefined &&
        marker === undefined &&
        // A point's own effects alone (a funnel stage's glow or inner
        // shadow, X-051) are a paint of its own.
        !shapePaint?.glow &&
        !shapePaint?.innerShadow &&
        !shapePaint?.softEdge &&
        !shapePaint?.bevel &&
        !shapePaint?.format3d) ||
      seen.has(key)
    ) {
      return undefined;
    }
    seen.add(key);
    styles.push({
      ...(seriesIndex ? { seriesIndex: seriesIndex as number } : {}),
      pointIndex: pointIndex as number,
      ...(shapePaint ? { shapePaint } : {}),
      ...(fill === undefined ? {} : { fill }),
      ...(outline === undefined ? {} : { outline }),
      ...(effect === undefined ? {} : { effect }),
      ...(explosion === undefined ? {} : { explosion: explosion as number }),
      ...(marker === undefined ? {} : { marker }),
    });
  }
  return styles.sort(
    (left, right) =>
      (left.seriesIndex ?? 0) - (right.seriesIndex ?? 0) ||
      left.pointIndex - right.pointIndex,
  );
}

export function presentationChartSliceColorCss(value: unknown): CSSProperties {
  const styles = normalizePresentationChartSliceStyles(value);
  if (!styles) return {};
  const css: Record<string, string> = {};
  for (const style of styles) {
    if (style.seriesIndex) continue;
    if (style.fill) {
      css[`--mrai-chart-slice-fill-${style.pointIndex + 1}`] =
        style.fill === "none" ? "transparent" : style.fill;
    }
    if (style.outline) {
      css[`--mrai-chart-slice-outline-${style.pointIndex + 1}`] =
        style.outline === "none" ? "transparent" : style.outline;
    }
  }
  return css as CSSProperties;
}

export function presentationChartFormatAxis(
  value: PresentationChartFormatTarget,
): "category" | "value" | null {
  return value === "axis:category"
    ? "category"
    : value === "axis:value"
      ? "value"
      : null;
}

export function presentationChartFormatAxisTitle(
  value: PresentationChartFormatTarget,
): "category" | "value" | null {
  return value === "axis-title:category"
    ? "category"
    : value === "axis-title:value"
      ? "value"
      : null;
}

export function normalizePresentationChartAxisTitleStyles(
  value: unknown,
): readonly PresentationChartAxisTitleStyle[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > 2) {
    return undefined;
  }
  const seen = new Set<string>();
  const styles: PresentationChartAxisTitleStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return undefined;
    const record = entry as Record<string, unknown>;
    const axis = record.axis;
    // The title's paint and text are an area style's (its box, its words).
    const { axis: _axis, ...rest } = record;
    const style = Object.keys(rest).length ? normalizePresentationChartAreaStyle(rest) : undefined;
    if (
      (axis !== "category" && axis !== "value") ||
      !style ||
      style.shapePaint !== undefined ||
      seen.has(axis)
    ) {
      return undefined;
    }
    seen.add(axis);
    styles.push({ axis, ...style });
  }
  const rank = { category: 0, value: 1 } as const;
  return styles.sort((left, right) => rank[left.axis] - rank[right.axis]);
}

export function presentationChartAxisTitleStyleCss(
  value: unknown,
  axis: "category" | "value",
): CSSProperties {
  const style = normalizePresentationChartAxisTitleStyles(value)?.find(
    (entry) => entry.axis === axis,
  );
  return presentationChartAreaStyleCss(style);
}

export function presentationChartGridlineStyleCss(value: unknown): CSSProperties {
  const outline = normalizePresentationChartAreaStyle(value)?.outline;
  return outline === undefined
    ? {}
    : ({
        "--mrai-chart-gridline":
          outline === "none" ? "transparent" : outline,
      } as CSSProperties);
}

export function presentationChartTrendlineStyleCss(value: unknown): CSSProperties {
  const outline = normalizePresentationChartAreaStyle(value)?.outline;
  return outline === undefined
    ? {}
    : ({
        "--mrai-chart-trendline":
          outline === "none" ? "transparent" : outline,
      } as CSSProperties);
}

export function presentationChartDataLabelStyleCss(value: unknown): CSSProperties {
  const style = normalizePresentationChartAreaStyle(value);
  return {
    "--mrai-chart-data-label-fill":
      style?.fill === undefined || style.fill === "none"
        ? "transparent"
        : style.fill,
    "--mrai-chart-data-label-outline":
      style?.outline === undefined || style.outline === "none"
        ? "transparent"
        : style.outline,
  } as CSSProperties;
}

export function normalizePresentationChartAxisStyles(
  value: unknown,
): readonly PresentationChartAxisStyle[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > 3) {
    return undefined;
  }
  const seen = new Set<string>();
  const styles: PresentationChartAxisStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return undefined;
    const record = entry as Record<string, unknown>;
    const axis = record.axis;
    const shapePaint = normalizePresentationChartAreaStyle(record)?.shapePaint;
    if (record.shapePaint !== undefined && !shapePaint) return undefined;
    const outline = normalizePresentationChartAreaPaint(record.outline);
    const fill = normalizePresentationChartAreaPaint(record.fill);
    if (record.fill !== undefined && (fill === undefined || fill === "none")) return undefined;
    const effect = normalizePresentationChartEffectPreset(record.effect);
    // The axis labels' own text (the axis's `c:txPr`).
    const text = record.text === undefined ? undefined : normalizeOfficeChartText(record.text);
    if (record.text !== undefined && !text) return undefined;
    if (
      (axis !== "category" && axis !== "value" && axis !== "secondaryValue") ||
      (outline === undefined && fill === undefined && effect === undefined && !shapePaint && !text) ||
      (axis === "secondaryValue" && (!text || outline !== undefined || fill !== undefined || effect !== undefined || shapePaint)) ||
      seen.has(axis)
    ) {
      return undefined;
    }
    seen.add(axis);
    styles.push({
      axis,
      ...(shapePaint ? { shapePaint } : {}),
      ...(outline === undefined ? {} : { outline }),
      ...(fill === undefined || fill === "none" ? {} : { fill }),
      ...(effect === undefined ? {} : { effect }),
      ...(text ? { text } : {}),
    });
  }
  return styles.sort(comparePresentationChartAxisStyles);
}

/** The order the axes' styles are kept in, as the API writes and reads
 * them: the category axis, the primary value axis, then the secondary value
 * axis (whose own labels' text the shared reader appends last). */
const PRESENTATION_CHART_AXIS_STYLE_RANK = { category: 0, value: 1, secondaryValue: 2 } as const;

export function comparePresentationChartAxisStyles(
  left: Pick<PresentationChartAxisStyle, "axis">,
  right: Pick<PresentationChartAxisStyle, "axis">,
): number {
  return PRESENTATION_CHART_AXIS_STYLE_RANK[left.axis] - PRESENTATION_CHART_AXIS_STYLE_RANK[right.axis];
}

export function presentationChartAxisStyleCss(
  value: unknown,
  horizontalBars = false,
): CSSProperties {
  const styles = normalizePresentationChartAxisStyles(value);
  if (!styles) return {};
  const category = styles.find((style) => style.axis === "category")?.outline;
  const numeric = styles.find((style) => style.axis === "value")?.outline;
  const horizontal = horizontalBars ? numeric : category;
  const vertical = horizontalBars ? category : numeric;
  return {
    boxSizing: "border-box",
    ...(horizontal === undefined
      ? {}
      : {
          borderBottomColor: horizontal === "none" ? "transparent" : horizontal,
          borderBottomStyle: "solid",
          borderBottomWidth: 1,
        }),
    ...(vertical === undefined
      ? {}
      : {
          borderLeftColor: vertical === "none" ? "transparent" : vertical,
          borderLeftStyle: "solid",
          borderLeftWidth: 1,
        }),
  };
}

/**
 * What a series' or a point's row holds beside its paint -- a series' marker
 * and Smoothed line, a point's explosion and series -- as a paint write
 * leaves it: the style's own when it states one (a key stated as undefined
 * takes it off), the row's otherwise. A recolour no longer drops a series'
 * marker or a slice's explosion.
 */
export function chartStyleRowExtras(
  key: string,
  existing: Readonly<Record<string, unknown>> | undefined,
  style: Readonly<Record<string, unknown>> | null,
): Record<string, unknown> {
  const names =
    key === "seriesStyles" ? ["marker", "smooth", "barShape"] : key === "sliceStyles" ? ["seriesIndex", "explosion", "marker"] : [];
  const extras: Record<string, unknown> = {};
  for (const name of names) {
    const value = style && Object.hasOwn(style, name) ? style[name] : existing?.[name];
    if (value !== undefined) extras[name] = value;
  }
  return extras;
}

/**
 * Format Data Point (`point:S:P`, a second click on one column, bar or funnel
 * stage): series S's point P's own row in `sliceStyles` -- its `c:dPt` in
 * that series -- or undefined. The first series' rows state no series.
 */
export function chartPointStyleRow<Row extends { readonly pointIndex: number; readonly seriesIndex?: number }>(
  rows: readonly Row[] | undefined,
  seriesIndex: number,
  pointIndex: number,
): Row | undefined {
  return rows?.find((row) => (row.seriesIndex ?? 0) === seriesIndex && row.pointIndex === pointIndex);
}

/**
 * `sliceStyles` with series S's point P formatted `style` (null: Reset to
 * Match Style, the point's row gone). A paint write keeps what the row holds
 * beside its paint (`chartStyleRowExtras`); the rows stay in series, then
 * point, order. Undefined when no row is left.
 */
export function chartRowsWithPointStyle<Row extends { readonly pointIndex: number; readonly seriesIndex?: number }>(
  rows: readonly Row[] | undefined,
  seriesIndex: number,
  pointIndex: number,
  style: PresentationChartAreaStyle | null,
): Row[] | undefined {
  const normalized = style === null ? null : normalizePresentationChartAreaStyle(style) ?? null;
  const existing = chartPointStyleRow(rows, seriesIndex, pointIndex);
  const extras = style === null ? {} : chartStyleRowExtras("sliceStyles", existing, style);
  delete extras.seriesIndex;
  const next: Row[] = (rows ?? []).filter((row) => row !== existing);
  if (normalized || Object.keys(extras).length) {
    next.push({ ...(seriesIndex ? { seriesIndex } : {}), pointIndex, ...extras, ...(normalized ?? {}) } as unknown as Row);
  }
  next.sort((left, right) => (left.seriesIndex ?? 0) - (right.seriesIndex ?? 0) || left.pointIndex - right.pointIndex);
  return next.length ? next : undefined;
}

/**
 * What Format Data Point shows and edits for series S's point P: its own row,
 * or, before it has one, its series' paint -- what Office copies into a
 * point's first own formatting, so one change keeps the rest of what the
 * point showed (its series' colour under a new glow).
 */
export function chartPointEffectiveStyle(
  format: { readonly sliceStyles?: unknown; readonly seriesStyles?: unknown } | undefined,
  seriesIndex: number,
  pointIndex: number,
): PresentationChartAreaStyle | undefined {
  const own = chartPointStyleRow(normalizePresentationChartSliceStyles(format?.sliceStyles), seriesIndex, pointIndex);
  if (own) {
    const { seriesIndex: _series, pointIndex: _point, explosion: _explosion, marker: _marker, ...paint } = own as Record<string, unknown>;
    return normalizePresentationChartAreaStyle(paint);
  }
  const series = normalizePresentationChartSeriesStyles(format?.seriesStyles)?.find((entry) => entry.seriesIndex === seriesIndex);
  if (!series) return undefined;
  const { seriesIndex: _series, marker: _marker, smooth: _smooth, barShape: _shape, ...paint } = series as Record<string, unknown>;
  return normalizePresentationChartAreaStyle(paint);
}
