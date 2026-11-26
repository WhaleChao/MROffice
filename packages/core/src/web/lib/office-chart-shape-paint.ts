/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import { officeDashLengthsWithCaps } from "./office-stroke-dash";
import type { CSSProperties } from "react";
import {
  normalizeOfficeShapeGradient,
  normalizeOfficeShapeShadow,
  officeGradientCss,
  officeShadowOffset,
  officeShadowCss,
  type OfficeShapeGradient,
  type OfficeGradientRect,
  type OfficeShapeShadow, officeGradientAngle, officeGradientDrawnStops } from "./office-shape-paint";
import type { OfficeShapeStyleRecipe } from "./office-shape-style-catalog";
import {
  OFFICE_STROKE_CAPS,
  OFFICE_STROKE_COMPOUNDS,
  OFFICE_STROKE_JOINS,
  type OfficeStrokeCap,
  type OfficeStrokeCompound,
  type OfficeStrokeJoin,
} from "./office-stroke-detail";
import { smartArtPictureSrc } from "./office-smartart/picture";
import { OFFICE_SHADOW_ALIGNMENTS, type OfficeShadowAlignment } from "./office-shadow-effects";
import { OFFICE_CHART_BEVEL, type OfficeChartBevel } from "./office-charts/chart-bevel";
import {
  normalizeOfficePictureEffects,
  type OfficeBevelEffect,
  type OfficeFormat3dEffect,
} from "./office-picture-effects";
import {
  normalizeOfficePictureFillOptions,
  officePictureFillOptionsStated,
  officePictureNaturalSize,
  type OfficePictureFillOptions,
} from "./office-picture-fill-options";
import { chartShadowBlur } from "./office-charts/shadow-deviation";

/** `a:headEnd@type` / `a:tailEnd@type` but "none" (ECMA-376 ST_LineEndType). */
export const OFFICE_CHART_LINE_END_TYPES = ["triangle", "stealth", "diamond", "oval", "arrow"] as const;
export type OfficeChartLineEndType = (typeof OFFICE_CHART_LINE_END_TYPES)[number];
/** ST_LineEndWidth / ST_LineEndLength. */
export const OFFICE_CHART_LINE_END_SIZES = ["sm", "med", "lg"] as const;
export type OfficeChartLineEndSize = (typeof OFFICE_CHART_LINE_END_SIZES)[number];

/**
 * Format ▸ Line ▸ Begin / End Arrow type and size: the line's `a:headEnd` and
 * `a:tailEnd`. Only a decorated end is stated (an end of type "none" is no
 * end); a size states its width and length, as Office's size grid writes
 * them, and an unstated size is the medium one Office draws. The field names
 * are the shapes' (`OfficeLineEndsState`).
 */
export interface OfficeChartLineEnds {
  headEnd?: OfficeChartLineEndType;
  headEndWidth?: OfficeChartLineEndSize;
  headEndLength?: OfficeChartLineEndSize;
  tailEnd?: OfficeChartLineEndType;
  tailEndWidth?: OfficeChartLineEndSize;
  tailEndLength?: OfficeChartLineEndSize;
}

/**
 * Extra paint only. Existing fill/outline fields remain the sole color owners.
 * Each part is optional -- Office's own chart styles paint a gradient or a
 * shadow with no line width stated -- but a paint states at least one. The
 * twin of `office_chart_shape_paint.chart_shape_paint`.
 *
 * The line's detail (compound, cap, join, ends) states only what differs from
 * the line Office draws for the element by default -- one line, round joins,
 * flat caps (round on a line, scatter or radar series), no arrows -- so a
 * chart Office wrote with its defaults reads as it always did.
 */
export interface OfficeChartShapePaint {
  outlineWidthPt?: number;
  fillGradient?: OfficeShapeGradient;
  /** A gradient line (`a:ln/a:gradFill`); `outline` is its first stop. */
  lineGradient?: OfficeShapeGradient;
  /** Format ▸ Line ▸ Dash type: the line's preset dash (absent = solid). */
  lineDash?: OfficeChartLineDash;
  /** Format ▸ Fill ▸ Pattern fill; `fill` is its foreground colour. */
  fillPattern?: OfficeChartFillPattern;
  /**
   * Format ▸ Fill ▸ Picture or texture fill (`a:blipFill`): the picture as
   * the one inline spelling every host keeps a picture fill in (a PNG or
   * JPEG `data:` URL, `officePictureFillSrc`). The element states no solid
   * `fill` beside it.
   */
  fillPicture?: string;
  /** The picture fill's options (tile / stretch offsets / transparency), the shapes' model. */
  fillPictureOptions?: OfficePictureFillOptions;
  shadow?: OfficeShapeShadow;
  /**
   * The outer shadow's scale and skew (`a:outerShdw@sx/@sy/@kx/@ky`, reading
   * 1047) and the corner of the element's box they hold (`algn`): the shadow
   * is the element's whole shape scaled about that corner (a series': all its
   * marks', `chart-render-server`). Only beside `shadow`.
   */
  shadowScale?: OfficeChartShadowScale;
  /**
   * Format ▸ Fill ▸ Solid fill ▸ Transparency, as the solid `fill` colour's
   * opacity (`a:solidFill/a:srgbClr/a:alpha`, 0 clear .. below 1); absent is
   * opaque. Only beside a solid `fill` colour.
   */
  fillAlpha?: number;
  /** Format ▸ Line ▸ Solid line ▸ Transparency: the solid `outline`'s opacity. */
  lineAlpha?: number;
  /** Format ▸ Line ▸ Compound type (`a:ln@cmpd`); absent = a single line. */
  lineCompound?: OfficeStrokeCompound;
  /** Format ▸ Line ▸ Cap type (`a:ln@cap`); absent = the element's default. */
  lineCap?: OfficeStrokeCap;
  /** Format ▸ Line ▸ Join type (`a:bevel` / `a:miter`); absent = round. */
  lineJoin?: OfficeStrokeJoin;
  /** Format ▸ Line ▸ Begin / End Arrow type and size. */
  lineEnds?: OfficeChartLineEnds;
  /** Format ▸ Effects ▸ Glow (`a:effectLst/a:glow`): its colour, size and opacity. */
  glow?: OfficeChartGlow;
  /** Format ▸ Effects ▸ Shadow, an Inner one (`a:effectLst/a:innerShdw`). */
  innerShadow?: OfficeChartInnerShadow;
  /** Format ▸ Effects ▸ Soft Edges (`a:effectLst/a:softEdge`): its size. */
  softEdge?: OfficeChartSoftEdge;
  /** Format ▸ Effects ▸ 3-D Format ▸ Top bevel (`a:sp3d/a:bevelT`), the shapes' model. */
  bevel?: OfficeBevelEffect;
  /**
   * Format ▸ Effects ▸ 3-D Format beyond the top bevel: Bottom bevel, Depth,
   * Contour, Material and Lighting (`a:sp3d`, `a:scene3d/a:lightRig`).
   */
  format3d?: OfficeFormat3dEffect;
}

/**
 * A chart element's Inner shadow: colour, opacity (absent = opaque), blur and
 * distance in points, angle in whole degrees clockwise from the right.
 */
export interface OfficeChartInnerShadow {
  readonly color: `#${string}`;
  readonly opacity?: number;
  readonly blurPt: number;
  readonly distancePt: number;
  readonly direction: number;
}

/** Office's Inner shadow Blur and Distance bound, points (`MAX_CHART_INNER_SHADOW_PT`). */
const OFFICE_CHART_MAX_INNER_SHADOW_PT = 100;

/** The twin of `office_chart_shape_paint.chart_inner_shadow`. */
export function normalizeOfficeChartInnerShadow(value: unknown): OfficeChartInnerShadow | undefined {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw Error("invalid_chart_inner_shadow");
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["color", "opacity", "blurPt", "distancePt", "direction"].includes(key)))
    throw Error("invalid_chart_inner_shadow");
  const opacity = raw.opacity === undefined ? 1 : raw.opacity;
  const within = (number: unknown) =>
    typeof number === "number" && Number.isFinite(number) && number >= 0 && number <= OFFICE_CHART_MAX_INNER_SHADOW_PT;
  if (
    typeof raw.color !== "string" ||
    !/^#[0-9a-f]{6}$/i.test(raw.color) ||
    !within(raw.blurPt) ||
    !within(raw.distancePt) ||
    typeof raw.direction !== "number" ||
    !Number.isInteger(raw.direction) ||
    raw.direction < 0 ||
    raw.direction >= 360 ||
    typeof opacity !== "number" ||
    !Number.isFinite(opacity) ||
    opacity <= 0 ||
    opacity > 1
  )
    throw Error("invalid_chart_inner_shadow");
  return {
    color: raw.color.toLowerCase() as `#${string}`,
    ...(opacity < 1 ? { opacity: Math.round(opacity * 10_000) / 10_000 } : {}),
    blurPt: Math.round((raw.blurPt as number) * 100) / 100,
    distancePt: Math.round((raw.distancePt as number) * 100) / 100,
    direction: raw.direction,
  };
}

/**
 * The key an ECharts item style carries a point's Inner shadow under: the
 * engine draws none, so the picture renderer shades the mark's inside itself
 * (`chart-render-server`, X-051).
 */
export const OFFICE_CHART_INNER_SHADOW = "officeInnerShadow";
/** The itemStyle key a mark's Soft Edges rides on to the picture renderer. */
export const OFFICE_CHART_SOFT_EDGE = "officeSoftEdge";

/** `OfficeChartShapePaint.shadowScale` (the server's `chart_shadow_scale`). */
export interface OfficeChartShadowScale {
  readonly algn: OfficeShadowAlignment;
  /** Scale, 1/1000 %; absent = 100 %. */
  readonly sx?: number;
  readonly sy?: number;
  /** Skew, 1/60000 degree; absent = 0. */
  readonly kx?: number;
  readonly ky?: number;
}

/** The scale bound (200 %) and skew bound the shared shadow value takes (`office_shadow_effects`). */
const MAX_CHART_SHADOW_SCALE = 200_000;
const MAX_CHART_SHADOW_SKEW = 5_399_999;

export function normalizeOfficeChartShadowScale(value: unknown): OfficeChartShadowScale | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["algn", "sx", "sy", "kx", "ky"].includes(key))) return null;
  if (!(OFFICE_SHADOW_ALIGNMENTS as readonly unknown[]).includes(raw.algn)) return null;
  const result: { -readonly [K in keyof OfficeChartShadowScale]: OfficeChartShadowScale[K] } = {
    algn: raw.algn as OfficeShadowAlignment,
  };
  for (const [key, neutral, bound] of [
    ["sx", 100_000, MAX_CHART_SHADOW_SCALE],
    ["sy", 100_000, MAX_CHART_SHADOW_SCALE],
    ["kx", 0, MAX_CHART_SHADOW_SKEW],
    ["ky", 0, MAX_CHART_SHADOW_SKEW],
  ] as const) {
    const number = raw[key];
    if (number === undefined) continue;
    if (typeof number !== "number" || !Number.isInteger(number) || Math.abs(number) > bound) return null;
    if (number !== neutral) result[key] = number;
  }
  return Object.keys(result).length > 1 ? result : null;
}

/**
 * The key an ECharts item style carries a point's scaled or skewed outer
 * shadow under (`OfficeChartScaledShadow`): the engine draws none for it,
 * the picture renderer draws the series' whole shape as its shadow
 * (`chart-render-server`, `scaledShadowsDrawn`).
 */
export const OFFICE_CHART_SCALED_SHADOW = "officeScaledShadow";

export interface OfficeChartScaledShadow {
  readonly shadow: OfficeShapeShadow;
  readonly scale: OfficeChartShadowScale;
}

/**
 * The width an arrow end is sized from, px: the line's, but never under
 * 0.7 mm (about 2pt). DrawingML sizes an end 2, 3 or 5 line widths (small,
 * medium, large); Excel for Mac draws a 1pt gridline's medium triangle 6pt
 * long and 6pt across (X-223 oracle, H26's value-axis gridlines), three
 * times 2pt, as LibreOffice's DrawingML import also floors it.
 */
export function officeLineEndUnitPx(lineWidthPx: number): number {
  return Math.max(lineWidthPx, ptToPx(1.984));
}

/** A chart element's Soft Edges: how far in from its edge it fades, in points. */
export interface OfficeChartSoftEdge {
  readonly radiusPt: number;
}

/** Office's Soft Edges Size spinner's bound, points (`office_chart_shape_paint.MAX_CHART_SOFT_EDGE_PT`). */
export const OFFICE_CHART_MAX_SOFT_EDGE_PT = 100;

/** The twin of `office_chart_shape_paint.chart_soft_edge`. */
export function normalizeOfficeChartSoftEdge(value: unknown): OfficeChartSoftEdge | undefined {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw Error("invalid_chart_soft_edge");
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).length !== 1 || !("radiusPt" in raw)) throw Error("invalid_chart_soft_edge");
  const radius = raw.radiusPt;
  if (typeof radius !== "number" || !Number.isFinite(radius) || radius <= 0 || radius > OFFICE_CHART_MAX_SOFT_EDGE_PT)
    throw Error("invalid_chart_soft_edge");
  return { radiusPt: Math.round(radius * 100) / 100 };
}

/** A chart element's Glow: colour, size (`rad`) in points, opacity (absent = opaque). */
export interface OfficeChartGlow {
  readonly color: `#${string}`;
  readonly radiusPt: number;
  readonly opacity?: number;
}

/** Office's Glow size spinner's bound, points (`office_chart_text.MAX_GLOW_RADIUS_PT`). */
const OFFICE_CHART_MAX_GLOW_RADIUS_PT = 150;

/** The twin of `office_chart_shape_paint.chart_glow`. */
export function normalizeOfficeChartGlow(value: unknown): OfficeChartGlow | undefined {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw Error("invalid_chart_glow");
  const raw = value as Record<string, unknown>;
  if (Object.keys(raw).some((key) => !["color", "radiusPt", "opacity"].includes(key))) throw Error("invalid_chart_glow");
  const opacity = raw.opacity === undefined ? 1 : raw.opacity;
  if (
    typeof raw.color !== "string" ||
    !/^#[0-9a-f]{6}$/i.test(raw.color) ||
    typeof raw.radiusPt !== "number" ||
    !Number.isFinite(raw.radiusPt) ||
    raw.radiusPt <= 0 ||
    raw.radiusPt > OFFICE_CHART_MAX_GLOW_RADIUS_PT ||
    typeof opacity !== "number" ||
    !Number.isFinite(opacity) ||
    opacity <= 0 ||
    opacity > 1
  )
    throw Error("invalid_chart_glow");
  return {
    color: raw.color.toLowerCase() as `#${string}`,
    radiusPt: Math.round(raw.radiusPt * 100) / 100,
    ...(opacity < 1 ? { opacity: Math.round(opacity * 10_000) / 10_000 } : {}),
  };
}

/**
 * A glow as the engine's one soft shadow round the mark (no offset). Office
 * spreads the glow's colour over its whole size and fades it at the rim
 * (H26_32readwriteLineChart4, Excel's print: the 40 % red at full strength
 * for a third of the 5pt, gone by its end); a blur can only fade from the
 * mark's edge, so the colour is drawn twice as strong over a blur of the
 * glow's size and a quarter.
 */
export function officeChartGlowShadow(
  glow: OfficeChartGlow,
  /** A filled mark's glow (a bar, a funnel stage): Office's halo round an
   * area fades sooner the larger it is (X-051 oracle, C18 zeta's 18pt glow:
   * full for 3pt, half by 6pt, gone by 11pt; 5pt gone by its size), so its
   * blur runs to about half the size and 2.7pt. */
  area = false,
): { shadowBlur: number; shadowColor: string } {
  return {
    shadowBlur: Math.round(ptToPx(area ? Math.min(glow.radiusPt * 1.25, glow.radiusPt * 0.46 + 2.7) : glow.radiusPt * 1.25) * 100) / 100,
    shadowColor: officeChartColorWithOpacity(glow.color, Math.min(1, 2 * (glow.opacity ?? 1))),
  };
}

/** An opacity on the file's grid (`a:alpha@val` is in 1/1000 of a percent). */
const opacityOnGrid = (value: number) => Math.round(value * 100_000) / 100_000;

function normalizeChartOpacity(value: unknown): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isFinite(value)) throw Error("invalid_chart_opacity");
  const opacity = opacityOnGrid(value);
  // Fully opaque is no transparency: the paint states none.
  if (opacity < 0 || opacity >= 1) throw Error("invalid_chart_opacity");
  return opacity;
}

export function normalizeOfficeChartLineEnds(value: unknown): OfficeChartLineEnds | undefined {
  if (value == null) return undefined;
  const raw = value as Record<string, unknown>;
  const keys = ["headEnd", "headEndWidth", "headEndLength", "tailEnd", "tailEndWidth", "tailEndLength"];
  if (
    typeof value !== "object" ||
    Array.isArray(value) ||
    !Object.keys(raw).length ||
    Object.keys(raw).some((key) => !keys.includes(key))
  )
    throw Error("invalid_chart_line_ends");
  const ends: OfficeChartLineEnds = {};
  for (const side of ["head", "tail"] as const) {
    const type = raw[`${side}End`];
    const width = raw[`${side}EndWidth`];
    const length = raw[`${side}EndLength`];
    if (type === undefined) {
      if (width !== undefined || length !== undefined) throw Error("invalid_chart_line_ends");
      continue;
    }
    if (!(OFFICE_CHART_LINE_END_TYPES as readonly unknown[]).includes(type)) throw Error("invalid_chart_line_ends");
    for (const size of [width, length])
      if (size !== undefined && !(OFFICE_CHART_LINE_END_SIZES as readonly unknown[]).includes(size))
        throw Error("invalid_chart_line_ends");
    ends[`${side}End`] = type as OfficeChartLineEndType;
    if (width !== undefined) ends[`${side}EndWidth`] = width as OfficeChartLineEndSize;
    if (length !== undefined) ends[`${side}EndLength`] = length as OfficeChartLineEndSize;
  }
  return ends;
}

/**
 * The cap Office draws a chart element's line with when its part states the
 * line's cap as Office writes it: round on the line of a line, scatter or
 * radar series and on a trendline (`cap="rnd"`), flat everywhere else
 * (`cap="flat"`). The pane shows it for an unstated cap; the twin of
 * `office_chart_shape_paint.chart_default_line_cap`.
 */
export function officeChartDefaultLineCap(
  area: string,
  chartType: string | undefined,
): OfficeStrokeCap {
  return area === "trendlineStyle" ||
    (area.startsWith("series:") && (chartType === "line" || chartType === "xy" || chartType === "radar"))
    ? "round"
    : "flat";
}

/** `#rrggbb` at an opacity, as CSS and ECharts take a colour. */
export function officeChartColorWithOpacity(color: string, opacity: number | undefined): string {
  if (opacity === undefined || !/^#[0-9a-fA-F]{6}$/.test(color)) return color;
  const channel = (at: number) => Number.parseInt(color.slice(at, at + 2), 16);
  return `rgba(${channel(1)}, ${channel(3)}, ${channel(5)}, ${opacity})`;
}
/**
 * ST_PresetPatternVal (ECMA-376 §20.1.10.51). The twin of
 * `office_chart_shape_paint.CHART_FILL_PATTERNS`.
 */
export const OFFICE_CHART_FILL_PATTERNS = [
  "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70",
  "pct75", "pct80", "pct90", "horz", "vert", "ltHorz", "ltVert", "dkHorz", "dkVert",
  "narHorz", "narVert", "dashHorz", "dashVert", "cross", "dnDiag", "upDiag",
  "ltDnDiag", "ltUpDiag", "dkDnDiag", "dkUpDiag", "wdDnDiag", "wdUpDiag",
  "dashDnDiag", "dashUpDiag", "diagCross", "smCheck", "lgCheck", "smGrid", "lgGrid",
  "dotGrid", "smConfetti", "lgConfetti", "horzBrick", "diagBrick", "solidDmnd",
  "openDmnd", "dotDmnd", "plaid", "sphere", "weave", "divot", "shingle", "wave",
  "trellis", "zigZag",
] as const;
export type OfficeChartFillPattern = Readonly<{
  preset: (typeof OFFICE_CHART_FILL_PATTERNS)[number];
  foreground: string;
  background: string;
  

  foregroundOpacity?: number;
  backgroundOpacity?: number;
}>;
const PATTERN_OPACITY_KEYS = ["foregroundOpacity", "backgroundOpacity"] as const;
/** The pattern's stated opacities (each below 1), for a normalised copy. */
function patternOpacities(raw: Record<string, unknown>): Pick<OfficeChartFillPattern, "foregroundOpacity" | "backgroundOpacity"> {
  const out: { foregroundOpacity?: number; backgroundOpacity?: number } = {};
  for (const key of PATTERN_OPACITY_KEYS) {
    const value = raw[key];
    if (typeof value === "number" && Number.isFinite(value) && value >= 0 && value < 1) out[key] = value;
  }
  return out;
}
const PATTERN_HEX = /^#[0-9a-fA-F]{6}$/;
export function normalizeOfficeChartFillPattern(
  value: unknown,
): OfficeChartFillPattern | undefined {
  if (value == null) return undefined;
  const raw = value as Record<string, unknown>;
  if (
    typeof value !== "object" ||
    Array.isArray(value) ||
    !Object.keys(raw).every((key) => key === "preset" || key === "foreground" || key === "background" || (PATTERN_OPACITY_KEYS as readonly string[]).includes(key)) ||
    PATTERN_OPACITY_KEYS.some((key) => raw[key] !== undefined && !(typeof raw[key] === "number" && Number.isFinite(raw[key]) && (raw[key] as number) >= 0 && (raw[key] as number) <= 1)) ||
    !(OFFICE_CHART_FILL_PATTERNS as readonly unknown[]).includes(raw.preset) ||
    typeof raw.foreground !== "string" ||
    !PATTERN_HEX.test(raw.foreground) ||
    typeof raw.background !== "string" ||
    !PATTERN_HEX.test(raw.background)
  )
    throw Error("invalid_chart_fill_pattern");
  return {
    preset: raw.preset as OfficeChartFillPattern["preset"],
    foreground: raw.foreground.toLowerCase(),
    background: raw.background.toLowerCase(),
    ...patternOpacities(raw),
  };
}
/**
 * A pattern as the chart canvas draws it: an ECharts decal whose symbol is
 * Office's 8 x 8 bitmap of it ({@link officeFillPatternTile}) as one path,
 * one symbol a cell, a point a bit (the leading moves pin the path's box to
 * the whole cell, so a sparse bitmap is not stretched over it).
 */
export function officeChartPatternDecal(
  pattern: OfficeChartFillPattern,
): Record<string, unknown> {
  const { size, marks } = officeFillPatternTile(pattern);
  
  // period; 3 pt at a pixel a bit). The decal's cell is a whole number of pixels,
  // so three bitmaps stand in one 32 px cell.
  const repeats = 3;
  const whole = size * repeats;
  let path = `M0 0M${whole} ${whole}`;
  for (let row = 0; row < repeats; row += 1) {
    for (let column = 0; column < repeats; column += 1) {
      for (const mark of marks) {
        if (mark.tag === "rect") {
          path += `M${mark.x + column * size} ${mark.y + row * size}h${mark.width}v${mark.height}h${-mark.width}Z`;
        }
      }
    }
  }
  const cell = (whole * 4) / 3;
  return {
    color: pattern.foreground,
    backgroundColor: pattern.background,
    symbol: `path://${path}`,
    symbolSize: 1,
    symbolKeepAspect: false,
    dashArrayX: [[cell, 0]],
    dashArrayY: [cell, 0],
    rotation: 0,
  };
}
/**
 * A pattern as an area's fill in the chart picture (a data label box, the
 * legend, the plot area): zrender's vector pattern, Office's 8 x 8 bitmap as
 * plain rectangles inside an SVG `<pattern>`. ECharts takes an image fill
 * for a label box only as one stretched picture (a `data:` reference the
 * picture sanitizer refuses), so the hatch is drawn as geometry instead: no
 * image and no external or `data:` reference anywhere. The vector form is the
 * one the chart engine draws (it renders SVG only).
 */
export function officeChartPatternAreaFill(pattern: OfficeChartFillPattern): Record<string, unknown> {
  const { size, marks } = officeFillPatternTile(pattern);
  const rect = (key: string, attrs: Record<string, string | number>) => ({ tag: "rect", key, attrs, children: undefined });
  return {
    type: "pattern",
    repeat: "repeat",
    svgWidth: size,
    svgHeight: size,
    svgElement: {
      tag: "g",
      key: "mrai-pattern",
      attrs: { "shape-rendering": "crispEdges" },
      children: [
        rect("bg", { x: 0, y: 0, width: size, height: size, fill: pattern.background }),
        ...marks.flatMap((mark, index) =>
          mark.tag === "rect"
            ? [rect("m" + index, { x: mark.x, y: mark.y, width: mark.width, height: mark.height, fill: pattern.foreground })]
            : [],
        ),
      ],
    },
  };
}
/**
 * The key a picture fill's pattern carries its placement under (read by the
 * chart render server once the element is laid out).
 */
export const OFFICE_CHART_PICTURE_PLACE = "officePicturePlace";
export type OfficeChartPicturePlace =
  | { readonly column: number; readonly row: number; readonly dx: number; readonly dy: number }
  | { readonly stretch: { readonly l: number; readonly t: number; readonly r: number; readonly b: number } };
/**
 * A picture fill in the chart picture: zrender's vector pattern holding the
 * picture -- an inline PNG / JPEG, the one reference the chart picture
 * admits -- at its own size (its resolution, a tile's Scale X / Y). The
 * element's box is not known when the chart is planned, so the pattern
 * carries where it goes (`OFFICE_CHART_PICTURE_PLACE`): "Tile picture as
 * texture" from the box's corner by its Alignment and Offset, a stretched
 * picture over the box. Undefined when the picture's size cannot be read.
 */
export function officeChartPictureFill(
  src: string,
  options: OfficePictureFillOptions | undefined,
): Record<string, unknown> | undefined {
  const opacity = options?.transparency ? Math.round((1 - options.transparency / 100) * 1000) / 1000 : undefined;
  const natural = officePictureNaturalSize(src);
  if (!natural) return undefined;
  const width = Math.max(1, (natural.width * Math.abs(options?.tile?.sx ?? 100_000)) / 100_000);
  const height = Math.max(1, (natural.height * Math.abs(options?.tile?.sy ?? 100_000)) / 100_000);
  const tile = options?.tile;
  const algn = tile?.algn ?? "tl";
  return {
    type: "pattern",
    repeat: "repeat",
    svgWidth: width,
    svgHeight: height,
    // Where the engine places the tile once the element's box is known
    // (`OFFICE_CHART_PICTURE_PLACE`, the chart render server's
    // `pictureFillsOnTheirBoxes`): a tile from the box's corner by its
    // Alignment and Offset, a stretch over the box (inside its offsets).
    [OFFICE_CHART_PICTURE_PLACE]: tile
      ? {
          column: /r$/.test(algn) ? 2 : algn === "t" || algn === "ctr" || algn === "b" ? 1 : 0,
          row: /^b/.test(algn) ? 2 : algn === "l" || algn === "ctr" || algn === "r" ? 1 : 0,
          dx: (tile.tx ?? 0) / 9525,
          dy: (tile.ty ?? 0) / 9525,
        }
      : {
          stretch: {
            l: (options?.fillRect?.l ?? 0) / 100_000,
            t: (options?.fillRect?.t ?? 0) / 100_000,
            r: (options?.fillRect?.r ?? 0) / 100_000,
            b: (options?.fillRect?.b ?? 0) / 100_000,
          },
        },
    svgElement: {
      tag: "image",
      key: "mrai-picture",
      attrs: {
        href: src,
        x: 0,
        y: 0,
        width,
        height,
        preserveAspectRatio: "none",
        ...(opacity === undefined ? {} : { opacity }),
      },
      children: undefined,
    },
  };
}
/**
 * A picture fill as a DOM box's background (the chart area's frame): tiled
 * at the picture's own size from the box's corner, or stretched over it.
 * (A stretch's offsets and a transparency are the chart picture's alone.)
 */
function officeChartPictureFillCss(src: string, options: OfficePictureFillOptions | undefined): CSSProperties {
  const natural = options?.tile ? officePictureNaturalSize(src) : null;
  return {
    backgroundImage: `url("${src}")`,
    ...(options?.tile && natural
      ? {
          backgroundRepeat: "repeat",
          backgroundSize: `${(natural.width * Math.abs(options.tile.sx ?? 100_000)) / 100_000}px ${(natural.height * Math.abs(options.tile.sy ?? 100_000)) / 100_000}px`,
          backgroundPosition: `${(options.tile.tx ?? 0) / 9525}px ${(options.tile.ty ?? 0) / 9525}px`,
        }
      : { backgroundRepeat: "no-repeat", backgroundSize: "100% 100%" }),
  };
}
/**
 * The same pattern as a CSS `background` (a DOM host: a table cell's Shading,
 * a shape): Office's 8 x 8 bitmap of it as an SVG tile, one pixel a bit, in
 * the foreground over the background -- the same bits the chart's decal and
 * the shapes' SVG tile draw.
 */
export function officeFillPatternCss(pattern: OfficeChartFillPattern): string {
  const { backgroundImage, backgroundSize, backgroundPosition, backgroundColor } = officeFillPatternBackground(pattern);
  return `${backgroundImage} ${backgroundPosition} / ${backgroundSize} repeat, ${backgroundColor}`;
}
/**
 * {@link officeFillPatternCss} as background LONGHANDS, for a host that also
 * states `backgroundClip` (WordArt clipped to the glyphs): React refuses to
 * mix the `background` shorthand with its longhands in one style object
 * (it warns, and a paint change can leave a stale layer behind).
 */
export function officeFillPatternBackground(pattern: OfficeChartFillPattern): {
  backgroundImage: string;
  backgroundColor: string;
  backgroundSize: string;
  backgroundPosition: string;
} {
  const { marks, paperInMarks } = officeFillPatternTile(pattern);
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8" shape-rendering="crispEdges">` +
    marks
      .map((mark) =>
        mark.tag === "rect"
          ? `<rect x="${mark.x}" y="${mark.y}" width="${mark.width}" height="${mark.height}" fill="${mark.fill ?? pattern.foreground}"${mark.opacity !== undefined ? ` fill-opacity="${mark.opacity}"` : ""}/>`
          : "",
      )
      .join("") +
    "</svg>";
  return {
    backgroundImage: `url("data:image/svg+xml,${encodeURIComponent(svg)}")`,
    // A see-through pattern's paper is in its tile.
    backgroundColor: paperInMarks ? "transparent" : pattern.background,
    // One bit a CSS pixel on screen, the tiles on the slide's (the page's)
    // grid: the host's pattern anchor sets these per element
    // (lib/office-pattern-anchor); outside one, the element's own 8 px tile
    // from its corner.
    backgroundSize: `var(--office-pattern-size, ${OFFICE_PATTERN_TILE}px) var(--office-pattern-size, ${OFFICE_PATTERN_TILE}px)`,
    backgroundPosition: "var(--office-pattern-x, 0px) var(--office-pattern-y, 0px)",
  };
}
/** One mark of a pattern tile, in tile units. */
export type OfficeFillPatternMark =
  | {
      readonly tag: "rect";
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
      /** A half-tone cell's colour (half way between the two); absent: the foreground. */
      readonly fill?: string;
      /** The mark's opacity below 1 (a see-through pattern colour); absent: opaque. */
      readonly opacity?: number;
    }
  | { readonly tag: "circle"; readonly cx: number; readonly cy: number; readonly r: number };

/** Office draws a pattern fill as an 8 x 8 bitmap, one screen point (one CSS pixel) a bit at any zoom, the tiles on the slide's grid (lib/office-pattern-anchor). */
const OFFICE_PATTERN_TILE = 8;

/**
 * Office's 54 pattern bitmaps (ST_PresetPatternVal), eight rows of eight
 * bits each (hex, the leftmost pixel the high bit), foreground where a bit is
 * set. Measured from PowerPoint for Mac's own drawing of every preset (black
 * on white, PDF export, nine tiles of each identical).
 */
const OFFICE_PATTERN_BITMAPS: Readonly<Record<OfficeChartFillPattern["preset"], string>> = {
  pct5: "8000000008000000",
  pct10: "8000080080000800",
  pct20: "8800220088002200",
  pct25: "8822882288228822",
  pct30: "aa44aa11aa44aa11",
  pct40: "aa55aa51aa55aa15",
  pct50: "aa55aa55aa55aa55",
  pct60: "ee55bb55ee55bb55",
  pct70: "77dd77dd77dd77dd",
  pct75: "77ffddff77ffddff",
  pct80: "effffeffeffffeff",
  pct90: "fffffff7ffffff7f",
  horz: "ff00000000000000",
  vert: "8080808080808080",
  ltHorz: "ff000000ff000000",
  ltVert: "8888888888888888",
  dkHorz: "ffff0000ffff0000",
  dkVert: "cccccccccccccccc",
  narHorz: "ff00ff00ff00ff00",
  narVert: "5555555555555555",
  dashHorz: "f00000000f000000",
  dashVert: "8080808008080808",
  cross: "ff80808080808080",
  dnDiag: "8040201008040201",
  upDiag: "0102040810204080",
  ltDnDiag: "8844221188442211",
  ltUpDiag: "1122448811224488",
  dkDnDiag: "cc663399cc663399",
  dkUpDiag: "3366cc993366cc99",
  wdDnDiag: "c1e070381c0e0783",
  wdUpDiag: "83070e1c3870e0c1",
  dashDnDiag: "0000884422110000",
  dashUpDiag: "0000112244880000",
  diagCross: "8142241818244281",
  smCheck: "9966669999666699",
  lgCheck: "f0f0f0f00f0f0f0f",
  smGrid: "ff888888ff888888",
  lgGrid: "ff80808080808080",
  dotGrid: "aa00800080008000",
  smConfetti: "8008400210012004",
  lgConfetti: "b130031bd8c00c8d",
  horzBrick: "ff808080ff080808",
  diagBrick: "0102040818244281",
  solidDmnd: "10387cfe7c381000",
  openDmnd: "8244281028448201",
  dotDmnd: "8000220008002200",
  plaid: "aa55aa55f0f0f0f0",
  sphere: "77898f8f7798f8f8",
  weave: "8854224588142251",
  divot: "0010081000800180",
  shingle: "038448300c020101",
  wave: "001825c0001825c0",
  trellis: "ff66ff99ff66ff99",
  zigZag: "8142241881422418",
};

/**
 * The cells Office draws half way between foreground and background beside
 * the one-pixel diagonals' ink (PowerPoint for Mac, pattern-presets-probe:
 * dnDiag, upDiag and diagCross carry a 50% grey cell on each side of the
 * line, black on white; every other preset is two-tone) -- the mirror of
 * ``OFFICE_PATTERN_HALF_BITS`` in ``office_chart_shape_paint.py``.
 */
const OFFICE_PATTERN_HALF_BITS: Readonly<Partial<Record<OfficeChartFillPattern["preset"], string>>> = {
  dnDiag: "41a05028140a0582",
  upDiag: "82050a142850a041",
  diagCross: "42a55a24245aa542",
};

/** The colour half way between two `#rrggbb` colours, or null. */
function officePatternHalfTone(foreground: string, background: string): string | null {
  const parse = (value: string) => (/^#[0-9a-f]{6}$/i.test(value) ? [1, 3, 5].map((at) => Number.parseInt(value.slice(at, at + 2), 16)) : null);
  const ink = parse(foreground);
  const paper = parse(background);
  if (!ink || !paper) return null;
  return `#${ink.map((channel, index) => Math.round((channel + paper[index]!) / 2).toString(16).padStart(2, "0")).join("")}`;
}

/** The runs of set bits of an 8 x 8 bitmap, as tile-unit rectangles. */
function officePatternRuns(bitmap: string, fill?: string, opacity?: number): OfficeFillPatternMark[] {
  const marks: OfficeFillPatternMark[] = [];
  for (let y = 0; y < OFFICE_PATTERN_TILE; y += 1) {
    const row = Number.parseInt(bitmap.slice(y * 2, y * 2 + 2), 16);
    let x = 0;
    while (x < OFFICE_PATTERN_TILE) {
      if (!(row & (0x80 >> x))) {
        x += 1;
        continue;
      }
      const start = x;
      while (x < OFFICE_PATTERN_TILE && row & (0x80 >> x)) x += 1;
      marks.push({ tag: "rect", x: start, y, width: x - start, height: 1, ...(fill ? { fill } : {}), ...(opacity !== undefined && opacity < 1 ? { opacity } : {}) });
    }
  }
  return marks;
}

/**
 * A pattern as an SVG tile: Office's 8 x 8 bitmap of it, the set bits of
 * each row drawn as runs in the foreground over the background, and the
 * one-pixel diagonals' half-tone cells in the colour half way between
 * (each mark's `fill`, when it states one; else the foreground).
 */
export function officeFillPatternTile(pattern: OfficeChartFillPattern): {
  readonly size: number;
  readonly angle: number;
  readonly marks: readonly OfficeFillPatternMark[];
  /** The background is among the marks (a see-through pattern: each bit its
   * own colour over what is under the shape, as PowerPoint's tile is): the
   * host lays no background under them. */
  readonly paperInMarks: boolean;
} {
  const bitmap = OFFICE_PATTERN_BITMAPS[pattern.preset] ?? OFFICE_PATTERN_BITMAPS.pct50;
  const ink = pattern.foregroundOpacity;
  const paper = pattern.backgroundOpacity;
  const seeThrough = (ink !== undefined && ink < 1) || (paper !== undefined && paper < 1);
  const marks = officePatternRuns(bitmap, undefined, ink);
  const halfBits = OFFICE_PATTERN_BITMAPS[pattern.preset] ? OFFICE_PATTERN_HALF_BITS[pattern.preset] : undefined;
  const halfTone = halfBits ? officePatternHalfTone(pattern.foreground, pattern.background) : null;
  if (halfBits && halfTone) marks.push(...officePatternRuns(halfBits, halfTone, seeThrough ? ((ink ?? 1) + (paper ?? 1)) / 2 : undefined));
  if (seeThrough) {
    // The paper where neither the ink nor a half-tone cell is.
    const taken = (bits: string | undefined, row: number) => (bits ? Number.parseInt(bits.slice(row * 2, row * 2 + 2), 16) : 0);
    const free = Array.from({ length: OFFICE_PATTERN_TILE }, (_, row) =>
      ((~(taken(bitmap, row) | taken(halfTone ? halfBits : undefined, row))) & 0xff).toString(16).padStart(2, "0"),
    ).join("");
    marks.unshift(...officePatternRuns(free, pattern.background, paper));
  }
  return { size: OFFICE_PATTERN_TILE, angle: 0, marks, paperInMarks: seeThrough };
}

/**
 * The same pattern as an SVG `<pattern>` element (a DOM-built SVG host: a
 * Docs shape's fill). `id` names the pattern for `url(#id)`.
 */
export function officeFillPatternSvg(pattern: OfficeChartFillPattern, id: string): SVGPatternElement {
  const ns = "http://www.w3.org/2000/svg";
  const make = <K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number>) => {
    const element = document.createElementNS(ns, tag);
    for (const [name, value] of Object.entries(attrs)) element.setAttribute(name, String(value));
    return element;
  };
  const { size, angle, marks, paperInMarks } = officeFillPatternTile(pattern);
  const tile = make("pattern", {
    id,
    // Laid on the page's grid, one bit a CSS pixel, by the page's pattern
    // anchor (lib/office-pattern-anchor).
    "data-office-pattern": "",
    patternUnits: "userSpaceOnUse",
    width: size,
    height: size,
    ...(angle ? { patternTransform: `rotate(${angle})` } : {}),
  }) as SVGPatternElement;
  tile.append(
    ...(paperInMarks ? [] : [make("rect", { x: 0, y: 0, width: size, height: size, fill: pattern.background })]),
    ...marks.map(({ tag, ...attrs }) => {
      const { opacity, ...rest } = attrs as typeof attrs & { opacity?: number };
      return make(tag, { ...rest, fill: ("fill" in rest && rest.fill) || pattern.foreground, ...(opacity !== undefined ? { "fill-opacity": opacity } : {}) });
    }),
  );
  return tile;
}

/** The slide background's four patterns, as the presets the server writes for them. */
const LEGACY_PATTERN_PRESETS: Readonly<Record<string, OfficeChartFillPattern["preset"]>> = {
  dots: "dotGrid",
  grid: "smGrid",
  diagonal: "ltUpDiag",
  checker: "smCheck",
};

/**
 * A shape's `fillPattern` in either form the page keeps: Office's preset
 * (`{preset, foreground, background}`), or the slide background's four
 * patterns (`{pattern, color, background}`), read as the preset the server
 * writes for it. Null when it is neither.
 */
export function officeFillPatternOf(value: unknown): OfficeChartFillPattern | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const hex = (color: unknown, fallback: string | null) =>
    typeof color === "string" && PATTERN_HEX.test(color) ? color.toLowerCase() : fallback;
  if ("preset" in raw) {
    const foreground = hex(raw.foreground, null);
    const background = hex(raw.background, null);
    return (OFFICE_CHART_FILL_PATTERNS as readonly unknown[]).includes(raw.preset) && foreground && background
      ? { preset: raw.preset as OfficeChartFillPattern["preset"], foreground, background, ...patternOpacities(raw) }
      : null;
  }
  const preset = LEGACY_PATTERN_PRESETS[typeof raw.pattern === "string" ? raw.pattern : "dots"];
  if (!preset) return null;
  return { preset, foreground: hex(raw.color, "#cbd5e1")!, background: hex(raw.background, "#ffffff")! };
}
/**
 * ST_PresetLineDashVal (ECMA-376 §20.1.10.48) but "solid". The twin of
 * `office_chart_shape_paint.CHART_LINE_DASHES`.
 */
export const OFFICE_CHART_LINE_DASHES = [
  "dot",
  "dash",
  "lgDash",
  "dashDot",
  "lgDashDot",
  "lgDashDotDot",
  "sysDash",
  "sysDot",
  "sysDashDot",
  "sysDashDotDot",
] as const;
export type OfficeChartLineDash = (typeof OFFICE_CHART_LINE_DASHES)[number];
/**
 * Each preset's dash and gap lengths in line widths (DrawingML's preset
 * dash definitions): what the canvas draws for it.
 */
export const OFFICE_CHART_LINE_DASH_PATTERN: Readonly<
  Record<OfficeChartLineDash, readonly number[]>
> = {
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
/**
 * The widest outline (a line series' line included) the model states, in
 * points: the worksheet's own line-width bound (`normalizeSpreadsheetLineWidthPt`),
 * so every width a sheet states fits. The twin of
 * `office_chart_shape_paint.MAX_CHART_OUTLINE_WIDTH_PT`.
 */
export const MAX_CHART_OUTLINE_WIDTH_PT = 24;

export function normalizeOfficeChartShapePaint(
  value: unknown,
): OfficeChartShapePaint | undefined {
  if (value == null) return undefined;
  if (
    typeof value !== "object" ||
    Array.isArray(value) ||
    Object.keys(value).length === 0 ||
    Object.keys(value).some(
      (key) =>
        ![
          "outlineWidthPt",
          "fillGradient",
          "fillPattern",
          "fillPicture",
          "fillPictureOptions",
          "lineGradient",
          "lineDash",
          "shadow",
          "shadowScale",
          "fillAlpha",
          "lineAlpha",
          "lineCompound",
          "lineCap",
          "lineJoin",
          "lineEnds",
          "glow",
          "innerShadow",
          "softEdge",
          "bevel",
          "format3d",
        ].includes(key),
    )
  )
    throw Error("invalid_chart_shape_paint");
  const raw = value as OfficeChartShapePaint;
  const fillAlpha = normalizeChartOpacity(raw.fillAlpha);
  const lineAlpha = normalizeChartOpacity(raw.lineAlpha);
  const lineEnds = normalizeOfficeChartLineEnds(raw.lineEnds);
  for (const [word, words] of [
    [raw.lineCompound, OFFICE_STROKE_COMPOUNDS],
    [raw.lineCap, OFFICE_STROKE_CAPS],
    [raw.lineJoin, OFFICE_STROKE_JOINS],
  ] as const)
    if (word !== undefined && !(words as readonly unknown[]).includes(word))
      throw Error("invalid_chart_line_detail");
  // A transparency is a solid colour's: a gradient states its own per stop,
  // and a pattern takes none.
  if (fillAlpha !== undefined && (raw.fillGradient || raw.fillPattern))
    throw Error("invalid_chart_opacity");
  if (lineAlpha !== undefined && raw.lineGradient) throw Error("invalid_chart_opacity");
  if (
    raw.outlineWidthPt !== undefined &&
    (!Number.isFinite(raw.outlineWidthPt) ||
      raw.outlineWidthPt < 0 ||
      raw.outlineWidthPt > MAX_CHART_OUTLINE_WIDTH_PT)
  )
    throw Error("invalid_chart_outline_width");
  const fillGradient = normalizeOfficeShapeGradient(raw.fillGradient);
  const fillPattern = normalizeOfficeChartFillPattern(raw.fillPattern);
  if (fillGradient && fillPattern) throw Error("invalid_chart_shape_paint");
  // A picture fill is one fill of its own: no gradient, pattern or solid
  // transparency beside it, and its options only with it.
  const fillPicture = raw.fillPicture === undefined ? undefined : smartArtPictureSrc(raw.fillPicture);
  if (raw.fillPicture !== undefined && !fillPicture) throw Error("invalid_chart_fill_picture");
  const fillPictureOptions =
    raw.fillPictureOptions === undefined ? undefined : normalizeOfficePictureFillOptions(raw.fillPictureOptions);
  if (
    fillPictureOptions === null ||
    (fillPictureOptions !== undefined && !fillPicture) ||
    (fillPicture && (fillGradient || fillPattern || fillAlpha !== undefined))
  )
    throw Error("invalid_chart_fill_picture");
  const lineGradient = normalizeOfficeShapeGradient(raw.lineGradient);
  const shadow = normalizeOfficeShapeShadow(raw.shadow);
  const shadowScale = raw.shadowScale === undefined ? undefined : normalizeOfficeChartShadowScale(raw.shadowScale);
  if (shadowScale === null || (shadowScale && !shadow)) throw Error("invalid_chart_shadow_scale");
  const glow = normalizeOfficeChartGlow(raw.glow);
  const innerShadow = normalizeOfficeChartInnerShadow(raw.innerShadow);
  const softEdge = normalizeOfficeChartSoftEdge(raw.softEdge);
  // Format ▸ Effects ▸ 3-D Format, in the shapes' own model (the server's
  // `office_chart_shape_paint.chart_3d_format`).
  const threeD =
    raw.bevel === undefined && raw.format3d === undefined
      ? {}
      : normalizeOfficePictureEffects({
          ...(raw.bevel === undefined ? {} : { bevel: raw.bevel }),
          ...(raw.format3d === undefined ? {} : { format3d: raw.format3d }),
        });
  if (!threeD || threeD.bevel === "none" || threeD.format3d === "none") throw Error("invalid_chart_3d_format");
  const bevel = threeD.bevel as OfficeBevelEffect | undefined;
  const format3d =
    threeD.format3d && Object.keys(threeD.format3d).length ? (threeD.format3d as OfficeFormat3dEffect) : undefined;
  if (
    raw.lineDash !== undefined &&
    !(OFFICE_CHART_LINE_DASHES as readonly unknown[]).includes(raw.lineDash)
  )
    throw Error("invalid_chart_line_dash");
  if (
    raw.outlineWidthPt === undefined &&
    !fillGradient &&
    !fillPattern &&
    !fillPicture &&
    !lineGradient &&
    raw.lineDash === undefined &&
    !shadow &&
    fillAlpha === undefined &&
    lineAlpha === undefined &&
    raw.lineCompound === undefined &&
    raw.lineCap === undefined &&
    raw.lineJoin === undefined &&
    !lineEnds &&
    !glow &&
    !innerShadow &&
    !softEdge &&
    !bevel &&
    !format3d
  )
    throw Error("invalid_chart_shape_paint");
  return {
    ...(raw.outlineWidthPt === undefined ? {} : { outlineWidthPt: raw.outlineWidthPt }),
    ...(fillGradient ? { fillGradient } : {}),
    ...(fillPattern ? { fillPattern } : {}),
    ...(fillPicture ? { fillPicture } : {}),
    ...(fillPictureOptions && officePictureFillOptionsStated(fillPictureOptions) ? { fillPictureOptions } : {}),
    ...(lineGradient ? { lineGradient } : {}),
    ...(raw.lineDash === undefined ? {} : { lineDash: raw.lineDash }),
    ...(shadow ? { shadow } : {}),
    ...(shadowScale ? { shadowScale } : {}),
    ...(fillAlpha === undefined ? {} : { fillAlpha }),
    ...(lineAlpha === undefined ? {} : { lineAlpha }),
    ...(raw.lineCompound === undefined ? {} : { lineCompound: raw.lineCompound }),
    ...(raw.lineCap === undefined ? {} : { lineCap: raw.lineCap }),
    ...(raw.lineJoin === undefined ? {} : { lineJoin: raw.lineJoin }),
    ...(lineEnds ? { lineEnds } : {}),
    ...(glow ? { glow } : {}),
    ...(innerShadow ? { innerShadow } : {}),
    ...(softEdge ? { softEdge } : {}),
    ...(bevel ? { bevel } : {}),
    ...(format3d ? { format3d } : {}),
  };
}
export function officeChartShapeRecipePaint(
  recipe: OfficeShapeStyleRecipe,
  filled: boolean,
): OfficeChartShapePaint {
  return normalizeOfficeChartShapePaint({
    outlineWidthPt: Math.round(recipe.strokeWidth * 750) / 1000,
    ...(filled && recipe.fillGradient
      ? { fillGradient: recipe.fillGradient }
      : {}),
    ...(recipe.shapeShadow ? { shadow: recipe.shapeShadow } : {}),
  })!;
}
export function officeChartShapePaintCss(
  value: OfficeChartShapePaint,
  /** The Gaussian deviation its outer shadow's blur radius draws, per unit of it (a slide's: `POWERPOINT_SHADOW_BLUR_DEVIATION`). */
  options: { readonly shadowBlurDeviation?: number } = {},
): CSSProperties {
  const p = normalizeOfficeChartShapePaint(value)!;
  return {
    ...(p.fillGradient
      ? { backgroundImage: officeGradientCss(p.fillGradient) }
      : p.fillPattern
        ? officeFillPatternBackground(p.fillPattern)
        : p.fillPicture
          ? officeChartPictureFillCss(p.fillPicture, p.fillPictureOptions)
          : {}),
    ...(p.outlineWidthPt === undefined ? {} : { borderWidth: p.outlineWidthPt + "pt" }),
    ...(p.shadow
      ? { boxShadow: officeShadowCss(p.shadow, options.shadowBlurDeviation === undefined ? {} : { blurDeviation: options.shadowBlurDeviation }) }
      : p.glow
        ? { boxShadow: `0 0 ${officeChartGlowShadow(p.glow).shadowBlur}px ${officeChartGlowShadow(p.glow).shadowColor}` }
        : {}),
  };
}

/**
 * A box's drawn border (the chart area's frame) as CSS draws one: a double
 * line for every compound type (the nearest CSS has), a dotted line for the
 * round and square dots, a dashed one for the other dashes. Undefined for a
 * single solid line.
 */
export function officeChartBorderStyleCss(value: OfficeChartShapePaint | undefined): "double" | "dotted" | "dashed" | undefined {
  if (!value) return undefined;
  if (value.lineCompound) return "double";
  if (value.lineDash === "dot" || value.lineDash === "sysDot") return "dotted";
  return value.lineDash ? "dashed" : undefined;
}
/**
 * A path gradient (`a:gradFill/a:path`) on a chart's element, as the picture
 * draws it once the chart is laid out (`chart-render-server`): over each
 * mark's own box (each bar, slice, doughnut segment, area; the legend key,
 * the plot and chart areas), rectangular contours for a rect (or shape) path
 * (`officeRectPathGradient`), each ray to a circle about the box's middle for
 * a circle path (`officeCirclePathGradient`). It rides on the colour the
 * engine is given (`OFFICE_CHART_PATH_GRADIENT`); the engine's own radial
 * stands in until then. Windows PowerPoint's print of MRAI-corpus/chart23
 * probes/chart23-pathgrad-probe(2).pptx: every case within 0.008 along the
 * stops, each bar and slice over its own box.
 */
export const OFFICE_CHART_PATH_GRADIENT = "officePathGradient";
export interface OfficeChartPathGradient {
  readonly kind: "rect" | "circle";
  readonly fillToRect?: OfficeGradientRect;
  readonly tileRect?: OfficeGradientRect;
  /** As drawn (`officeGradientDrawnStops`). */
  readonly stops: readonly { readonly position: number; readonly color: string; readonly alpha?: number }[];
}
export function officeChartGradientColor(value: OfficeShapeGradient) {
  // Each stop at its own transparency (Gradient stops ▸ Transparency), the
  // stops drawn as PowerPoint draws them (`officeGradientDrawnStops`).
  const drawn = officeGradientDrawnStops(value.stops);
  const colorStops = drawn.map((stop) => ({
    offset: stop.position,
    color: officeChartColorWithOpacity(stop.color, stop.alpha),
  }));
  if (value.path) {
    // The engine's stand-in: a radial from the focus rectangle to the
    // farthest corner. The picture draws the path gradient itself over the
    // element's own box (`OFFICE_CHART_PATH_GRADIENT`).
    const rect = value.path.fillToRect;
    const x = Math.min(1, Math.max(0, (rect?.l ?? 50_000) / 100_000));
    const y = Math.min(1, Math.max(0, (rect?.t ?? 50_000) / 100_000));
    const r = Math.max(Math.hypot(x, y), Math.hypot(1 - x, y), Math.hypot(x, 1 - y), Math.hypot(1 - x, 1 - y));
    const path: OfficeChartPathGradient = {
      // A shape path over a box is its rectangle's contours (a bar's, exactly).
      kind: value.path.kind === "circle" ? "circle" : "rect",
      ...(value.path.fillToRect ? { fillToRect: value.path.fillToRect } : {}),
      ...(value.tileRect ? { tileRect: value.tileRect } : {}),
      stops: drawn.map((stop) => ({ position: stop.position, color: stop.color, ...(stop.alpha !== undefined && stop.alpha < 1 ? { alpha: stop.alpha } : {}) })),
    };
    return {
      type: "radial",
      x,
      y,
      r: value.path.kind === "circle" ? r : r / Math.SQRT2,
      colorStops,
      [OFFICE_CHART_PATH_GRADIENT]: path,
    };
  }
  const angle = ((officeGradientAngle(value) - 90) * Math.PI) / 180;
  const dx = Math.cos(angle),
    dy = Math.sin(angle);
  return {
    type: "linear",
    x: (1 - dx) / 2,
    y: (1 - dy) / 2,
    x2: (1 + dx) / 2,
    y2: (1 + dy) / 2,
    colorStops,
  };
}
/**
 * The paint as ECharts takes it for an area (a legend, a label box, the plot
 * area), an item (a bar, a slice, a marker) or a line. `colors` are the
 * element's own solid `fill` and `outline`, which a transparency is drawn
 * through (the colour itself stays the element's).
 */
export function officeChartShapeEChartsStyle(
  value: OfficeChartShapePaint,
  target: "area" | "item" | "line" = "area",
  colors?: { readonly fill?: string; readonly outline?: string },
): Record<string, unknown> {
  const p = normalizeOfficeChartShapePaint(value)!;
  const fillSeen =
    p.fillAlpha !== undefined && colors?.fill && colors.fill !== "none"
      ? officeChartColorWithOpacity(colors.fill, p.fillAlpha)
      : undefined;
  const lineSeen =
    p.lineAlpha !== undefined && colors?.outline && colors.outline !== "none"
      ? officeChartColorWithOpacity(colors.outline, p.lineAlpha)
      : undefined;
  const cap = p.lineCap === "flat" ? "butt" : p.lineCap;
  const color = p.fillGradient
    ? officeChartGradientColor(p.fillGradient)
    : undefined;
  const lineColor = p.lineGradient
    ? officeChartGradientColor(p.lineGradient)
    : undefined;
  // A stated line thinner than a pixel is drawn as Office draws it, a
  // one-pixel hairline: the 0.25pt black legend box of
  // H18_32readwriteScatterChart2 drew as a third of a pixel, a faint grey.
  // A line stated with width 0 (`a:ln w="0"`, LibreOffice's thin lines) in a
  // colour of its own is the thinnest line Excel draws, a hairline: the
  // plot-area frame, legend box and leader lines of LibreOffice re-saves
  // vanished (chart matrix, L_P01/L_P04/L_P06). A line of "none" stays none.
  const hairline = typeof colors?.outline === "string" && /^#[0-9a-f]{6}$/i.test(colors.outline);
  const width =
    p.outlineWidthPt === undefined
      ? undefined
      : p.outlineWidthPt > 0
        ? Math.max(1, ptToPx(p.outlineWidthPt))
        : hairline
          ? 1
          : 0;
  // A preset dash, in line widths, drawn at the line's width (1 px at least).
  // A hairline's dash is made of zero-width steps, and Excel draws it solid
  // (L_P01: LibreOffice's dashed w="0" plot-area frame is a solid gold line).
  // A round or square cap is counted inside each dash, as Office counts it:
  // drawn as stated, a round System Dot gridline ran solid (X-223).
  const dash = p.lineDash && p.outlineWidthPt !== 0
    ? officeDashLengthsWithCaps(
        OFFICE_CHART_LINE_DASH_PATTERN[p.lineDash],
        Math.max(1, width ?? 1),
        cap === "round" || cap === "square",
      )
    : undefined;
  return {
    ...(target === "line"
      ? {
          ...(width === undefined ? {} : { width }),
          ...(lineColor ? { color: lineColor } : lineSeen ? { color: lineSeen } : {}),
          ...(dash ? { type: dash } : {}),
          ...(cap ? { cap } : {}),
          ...(p.lineJoin ? { join: p.lineJoin } : {}),
        }
      : {
          ...(width === undefined ? {} : { borderWidth: width }),
          ...(lineColor ? { borderColor: lineColor } : lineSeen ? { borderColor: lineSeen } : {}),
          ...(dash ? { borderType: dash } : {}),
          ...(cap ? { borderCap: cap } : {}),
          ...(p.lineJoin ? { borderJoin: p.lineJoin } : {}),
          ...(fillSeen ? { [target === "item" ? "color" : "backgroundColor"]: fillSeen } : {}),
          ...(p.fillPattern
            ? target === "item"
              ? { color: p.fillPattern.background, decal: officeChartPatternDecal(p.fillPattern) }
              : { backgroundColor: officeChartPatternAreaFill(p.fillPattern) }
            : {}),
          ...(p.fillPicture
            ? (() => {
                const picture = officeChartPictureFill(p.fillPicture, p.fillPictureOptions);
                return picture ? { [target === "item" ? "color" : "backgroundColor"]: picture } : {};
              })()
            : {}),
          ...(color
            ? { [target === "item" ? "color" : "backgroundColor"]: color }
            : {}),
        }),
    ...(p.shadow && p.shadowScale && target === "item"
      ? { [OFFICE_CHART_SCALED_SHADOW]: { shadow: p.shadow, scale: p.shadowScale } satisfies OfficeChartScaledShadow }
      : p.shadow
      ? {
          // A slide's chart: PowerPoint's deviation (`chartShadowBlur`, scoped by the picture server).
          shadowBlur: chartShadowBlur(p.shadow.blur),
          shadowColor:
            p.shadow.color +
            Math.round(p.shadow.opacity * 255)
              .toString(16)
              .padStart(2, "0"),
          // Its offset is along its direction (straight down when unstated).
          shadowOffsetX: officeShadowOffset(p.shadow).dx,
          shadowOffsetY: officeShadowOffset(p.shadow).dy,
        }
      : p.glow
        ? { ...officeChartGlowShadow(p.glow, target === "item"), shadowOffsetX: 0, shadowOffsetY: 0 }
        : {}),
    ...(p.innerShadow && target === "item" ? { [OFFICE_CHART_INNER_SHADOW]: p.innerShadow } : {}),
    // The engine draws no soft edge: the picture renderer fades the mark in
    // from its edge itself (`chart-render-server`).
    ...(p.softEdge && target === "item" ? { [OFFICE_CHART_SOFT_EDGE]: p.softEdge } : {}),
    // The engine draws no bevel: the picture renderer lights the mark's
    // rims itself (`chart-render-server`, `office-charts/chart-bevel`).
    ...((p.bevel || p.format3d) && target === "item"
      ? {
          [OFFICE_CHART_BEVEL]: {
            ...(p.bevel ? { bevel: p.bevel } : {}),
            ...(p.format3d ? { format3d: p.format3d } : {}),
            ...(colors?.outline && colors.outline !== "none" ? { automaticContour: colors.outline } : {}),
          } satisfies OfficeChartBevel,
        }
      : {}),
  };
}
