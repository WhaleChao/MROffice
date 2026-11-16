/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The look an Office 2016 chart's style and colour parts give it.
 *
 * An imported chartEx part carries little paint of its own; its sibling Chart
 * Style and Chart Colors parts hold what Excel and Word draw. The server reads
 * those against the host theme into this bounded, display-only shape, and the
 * canvas draws it under any explicit paint the chart states. Nothing here is
 * editable or written back.
 */
import { PX_PER_PT } from "./office-units";
import {
  normalizeOfficeShapeGradient,
  type OfficeShapeGradient,
} from "./office-shape-paint";
import { officeChartGradientColor } from "./office-chart-shape-paint";
import type { ChartMarkerSymbol } from "./presentation-chart-area-style";

const isHex = (value: unknown): value is string =>
  typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value);
const isPaint = (value: unknown): value is string =>
  value === "none" || isHex(value);
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const finite = (value: unknown, lo: number, hi: number): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= lo && value <= hi;

export interface ChartNativeLookPattern {
  preset: string;
  fg: string;
  bg: string;
}
export interface ChartNativeLookArea {
  fill?: string;
  /** A flat fill's own opacity, kept apart from its colour; absent means opaque. */
  fillOpacity?: number;
  outline?: string;
  outlineOpacity?: number;
  outlineWidthPt?: number;
  dash?: string;
  fillGradient?: OfficeShapeGradient;
  radialGradient?: { stops: OfficeShapeGradient["stops"] };
  pattern?: ChartNativeLookPattern;
}
export interface ChartNativeLookText {
  color?: string;
  sizePt?: number;
  bold?: boolean;
  caps?: "all" | "small";
}
export interface ChartNativeLookLine {
  outline: string;
  outlineOpacity?: number;
  outlineWidthPt?: number;
  dash?: string;
}
/** Where a treemap writes its top-level branch names. */
export type ChartNativeLookParentLabelLayout = "overlapping" | "banner" | "none";

/** How a filled map names its regions: not at all, where a name fits, or always. */
export type ChartNativeMapLabelLayout = "none" | "bestFitOnly" | "showAll";

/**
 * The view Office draws an imported filled map in, read from the part's own
 * geocoding cache: the viewed continent's bounds, every country of it as
 * compressed GeoJSON, and which outline each ISO numeric code names.
 */
export interface ChartNativeLookMap {
  labelLayout: ChartNativeMapLabelLayout;
  /** Where the colour scale sits along its side; Office centres it when the part says nothing. */
  legendAlign?: "min" | "ctr" | "max";
  projection: "native-africa-albers-auto-v1";
  parentBounds: [number, number, number, number];
  featuresBase64: string;
  featureBytes: number;
  bindings: { id: string; entityId: string }[];
  mapName: string;
  attribution: string;
  copyrights: string[];
  fontFamily?: string;
  /** The series name the part states, painted over the colour scale. */
  seriesName?: string;
  /** The colour scale's ends when a host already resolved them; otherwise they come from the series colour. */
  scale?: { low: string; high: string };
}

/** A funnel's category axis as its part states it: stages are bands on it. */
export interface ChartNativeLookFunnel {
  /** The gap between stages, as a fraction of each stage's band. */
  gapFraction: number;
  showCategoryLabels?: boolean;
  majorTicks?: "none" | "in" | "out";
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  /** The axis' number format, which Office also shows the stage values in. */
  numberFormat?: string;
}

/** One colour transform on a series' automatic colour (`a:shade`, `a:satMod` ...). */
export type ChartNativeLookTransform = [string, number];

/**
 * The paint Office's built-in chart style 9-48 (`c:style`, the Office
 * 2007/2010 gallery) gives an unpainted bar, column, area or pie series from
 * the workbook theme's style matrix (server `office_chart_builtin_style`):
 * its fill as gradient stops of transforms on the series' own automatic
 * colour, its line as transforms on that colour or as a fixed theme colour,
 * its shadow in pixels. `own` lists what the part paints itself (by series
 * place, and the points of a series), which keeps its own paint.
 */
export interface ChartNativeLookBuiltInStyle {
  style: number;
  /** `angle`: DrawingML's `a:lin@ang` in degrees (90 = top to bottom); a
   * path gradient (the theme's third fill `a:path`, reading row 822) states
   * `path` instead, its focus `fillToRect` in 1/1000 %. */
  fill?: {
    angle?: number;
    path?: { kind: "circle" | "rect" | "shape"; fillToRect?: { l?: number; t?: number; r?: number; b?: number } };
    stops: { position: number; transforms: ChartNativeLookTransform[] }[];
  };
  line?: { widthPt: number; color?: string; transforms?: ChartNativeLookTransform[] };
  shadow?: { color: string; opacity: number; blur: number; distance: number; direction: number };
  own?: { series: number; fill?: true; line?: true; effect?: true; unfilled?: true; points?: number[] }[];
  /** Each series' colour slot (`c:idx`) by its place, when one differs. */
  slots?: number[];
  


  ink?: string;
  /** The theme's first line style's transforms (`a:lnStyleLst`[0] on
   * phClr) an unpainted line, X Y or radar series draws its automatic
   * colour through, in any classic style 1-48 (X-257). */
  seriesLine?: ChartNativeLookTransform[];
}

export interface ChartNativeLook {
  


  labelRoom?: boolean;
  /** The text insets (pt) the chart's data labels' box states
   * (`c:dLbls/c:txPr/a:bodyPr`, the import's `source.dataLabelInsets`,
   * reading row 821); absent is DrawingML's 7.2 / 3.6pt. Display only. */
  dataLabelInsets?: { l: number; t: number; r: number; b: number };
  /** Office's built-in chart style 9-48 paint for unpainted filled series. */
  builtInStyle?: ChartNativeLookBuiltInStyle;
  styleId?: string;
  colorMethod?: string;
  /** A treemap part's own parent label layout; Office overlaps when it says nothing. */
  parentLabelLayout?: ChartNativeLookParentLabelLayout;
  /** A filled map's own view; without one the map is drawn on the world. */
  map?: ChartNativeLookMap;
  funnel?: ChartNativeLookFunnel;
  chartArea?: ChartNativeLookArea;
  plotArea?: ChartNativeLookArea;
  legend?: ChartNativeLookArea & ChartNativeLookText;
  title?: ChartNativeLookText;
  axisTitle?: ChartNativeLookText;
  /**
   * The look is Excel's built-in chart text (a classic part that states no
   * text of its own: `excelBuiltInChartTextLook`), not a style part's: a
   * gallery pick (`chartStyle`) does not replace it.
   */
  builtInText?: true;
  /**
   * The host draws `builtInStyle`'s series paint without Excel's built-in
   * text: PowerPoint's classic charts in style 9-48
   * (`presentationBuiltInStyleLook`).
   */
  builtInStylePaint?: true;
  dataLabel?: ChartNativeLookText;
  axis?: ChartNativeLookText & { line?: ChartNativeLookLine };
  gridline?: ChartNativeLookLine;
  /** Minor gridlines that state no paint of their own. */
  minorGridline?: ChartNativeLookLine;
  /** A trendline that states no line of its own. */
  trendline?: ChartNativeLookLine;
  /**
   * Drop lines, high-low lines, series lines and error bars that state no
   * line of their own (Excel's built-in look: text-colour hairlines).
   */
  helperLine?: ChartNativeLookLine;
  /**
   * Automatic markers: the symbol each marked series takes by its index
   * (cycling) and their size, for a series that states no marker.
   */
  markers?: { symbols: readonly ChartMarkerSymbol[]; sizePt: number };
  /** A waterfall's connector lines; the part hides them only by saying so. */
  connector?: ChartNativeLookLine;
  connectorLines?: false;
  points?: ChartNativeLookArea[];
  /**
   * The host document's accent1..accent6, resolved from its theme by the same
   * reading the points came from. Change Colors names theme slots, so a pick
   * made on the page is drawn against these.
   */
  themeAccents?: string[];
  /**
   * The Change Colors palette of the shared table the chart's Chart Colors
   * part spells, when it spells one exactly: the pick these points were
   * resolved with. A page pick that differs repaints them
   * (`chartLookWithColorPalette`).
   */
  colorPalette?: string;
}

function stops(value: unknown): OfficeShapeGradient["stops"] | undefined {
  if (!Array.isArray(value) || value.length < 2 || value.length > 8) return undefined;
  const result: OfficeShapeGradient["stops"] = [];
  for (const stop of value) {
    if (!isRecord(stop) || !finite(stop.position, 0, 1) || !isHex(stop.color)) return undefined;
    result.push({ position: stop.position, color: stop.color.toLowerCase() });
  }
  return result;
}

function area(value: unknown): ChartNativeLookArea | undefined {
  if (!isRecord(value)) return undefined;
  const result: ChartNativeLookArea = {};
  if (isPaint(value.fill)) result.fill = value.fill.toLowerCase();
  if (result.fill && finite(value.fillOpacity, 0, 1)) result.fillOpacity = value.fillOpacity;
  if (isPaint(value.outline)) result.outline = value.outline.toLowerCase();
  if (result.outline && finite(value.outlineOpacity, 0, 1)) result.outlineOpacity = value.outlineOpacity;
  if (finite(value.outlineWidthPt, 0, 48)) result.outlineWidthPt = value.outlineWidthPt;
  if (typeof value.dash === "string" && /^[a-zA-Z]{1,24}$/.test(value.dash)) result.dash = value.dash;
  const gradient = normalizeOfficeShapeGradient(value.fillGradient);
  if (gradient) result.fillGradient = gradient;
  if (isRecord(value.radialGradient)) {
    const radial = stops(value.radialGradient.stops);
    if (radial) result.radialGradient = { stops: radial };
  }
  if (
    isRecord(value.pattern) &&
    typeof value.pattern.preset === "string" &&
    /^[a-zA-Z0-9]{1,24}$/.test(value.pattern.preset) &&
    isHex(value.pattern.fg) &&
    isHex(value.pattern.bg)
  ) {
    result.pattern = {
      preset: value.pattern.preset,
      fg: value.pattern.fg.toLowerCase(),
      bg: value.pattern.bg.toLowerCase(),
    };
  }
  return Object.keys(result).length ? result : undefined;
}

function text(value: unknown): ChartNativeLookText | undefined {
  if (!isRecord(value)) return undefined;
  const result: ChartNativeLookText = {};
  if (isHex(value.color)) result.color = value.color.toLowerCase();
  if (finite(value.sizePt, 1, 200)) result.sizePt = value.sizePt;
  if (value.bold === true) result.bold = true;
  if (value.caps === "all" || value.caps === "small") result.caps = value.caps;
  return Object.keys(result).length ? result : undefined;
}

function line(value: unknown): ChartNativeLookLine | undefined {
  if (!isRecord(value) || !isPaint(value.outline)) return undefined;
  const result: ChartNativeLookLine = { outline: value.outline.toLowerCase() };
  if (finite(value.outlineOpacity, 0, 1)) result.outlineOpacity = value.outlineOpacity;
  if (finite(value.outlineWidthPt, 0, 48)) result.outlineWidthPt = value.outlineWidthPt;
  if (typeof value.dash === "string" && /^[a-zA-Z]{1,24}$/.test(value.dash)) result.dash = value.dash;
  return result;
}

function nativeMap(value: unknown): ChartNativeLookMap | undefined {
  if (!isRecord(value)) return undefined;
  const bounds = value.parentBounds;
  const bindings = value.bindings;
  const copyrights = value.copyrights;
  if (
    (value.labelLayout !== "none" && value.labelLayout !== "bestFitOnly" && value.labelLayout !== "showAll") ||
    value.projection !== "native-africa-albers-auto-v1" ||
    !Array.isArray(bounds) ||
    bounds.length !== 4 ||
    !bounds.every((entry, index) => finite(entry, index < 2 ? -180 : -90, index < 2 ? 180 : 90)) ||
    typeof value.featuresBase64 !== "string" ||
    // The server keeps it under a page's per-string limit.
    value.featuresBase64.length > 1_000_000 ||
    !Number.isInteger(value.featureBytes) ||
    !finite(value.featureBytes, 1, 4_000_000) ||
    !Array.isArray(bindings) ||
    bindings.length < 1 ||
    bindings.length > 100 ||
    !bindings.every(
      (binding) =>
        isRecord(binding) &&
        typeof binding.id === "string" &&
        /^\d{3}$/.test(binding.id) &&
        typeof binding.entityId === "string" &&
        /^\d{1,20}$/.test(binding.entityId),
    ) ||
    typeof value.mapName !== "string" ||
    !/^mrai-native-map-[a-f0-9]{64}$/.test(value.mapName) ||
    typeof value.attribution !== "string" ||
    value.attribution.length > 300 ||
    !Array.isArray(copyrights) ||
    copyrights.length > 100 ||
    !copyrights.every((entry) => typeof entry === "string" && entry.length >= 1 && entry.length <= 200)
  ) {
    return undefined;
  }
  const result: ChartNativeLookMap = {
    labelLayout: value.labelLayout,
    projection: value.projection,
    parentBounds: bounds as [number, number, number, number],
    featuresBase64: value.featuresBase64,
    featureBytes: value.featureBytes as number,
    bindings: bindings.map((binding) => ({ id: binding.id as string, entityId: binding.entityId as string })),
    mapName: value.mapName,
    attribution: value.attribution,
    copyrights: copyrights as string[],
  };
  if (typeof value.fontFamily === "string" && /^[^"'\\;{}<>\u0000-\u001f]{1,64}$/.test(value.fontFamily)) {
    result.fontFamily = value.fontFamily;
  }
  if (value.legendAlign === "min" || value.legendAlign === "ctr" || value.legendAlign === "max") {
    result.legendAlign = value.legendAlign;
  }
  if (typeof value.seriesName === "string" && value.seriesName && value.seriesName.length <= 300) {
    result.seriesName = value.seriesName;
  }
  if (isRecord(value.scale) && isHex(value.scale.low) && isHex(value.scale.high)) {
    result.scale = { low: value.scale.low.toLowerCase(), high: value.scale.high.toLowerCase() };
  }
  return result;
}

function nativeFunnel(value: unknown): ChartNativeLookFunnel | undefined {
  if (!isRecord(value) || !finite(value.gapFraction, 0, 0.999999)) return undefined;
  const result: ChartNativeLookFunnel = { gapFraction: value.gapFraction };
  if (typeof value.showCategoryLabels === "boolean") result.showCategoryLabels = value.showCategoryLabels;
  if (value.majorTicks === "none" || value.majorTicks === "in" || value.majorTicks === "out") {
    result.majorTicks = value.majorTicks;
  }
  if (finite(value.fontSize, 1, 200)) result.fontSize = value.fontSize;
  if (typeof value.fontFamily === "string" && /^[^"'\\;{}<>\u0000-\u001f]{1,64}$/.test(value.fontFamily)) {
    result.fontFamily = value.fontFamily;
  }
  if (isHex(value.fontColor)) result.fontColor = value.fontColor.toLowerCase();
  if (typeof value.numberFormat === "string" && value.numberFormat.length <= 100) {
    result.numberFormat = value.numberFormat;
  }
  return result;
}

function transforms(value: unknown): ChartNativeLookTransform[] | undefined {
  if (!Array.isArray(value) || value.length > 8) return undefined;
  const result: ChartNativeLookTransform[] = [];
  for (const entry of value) {
    if (
      !Array.isArray(entry) ||
      entry.length !== 2 ||
      typeof entry[0] !== "string" ||
      !/^[a-zA-Z]{2,12}$/.test(entry[0]) ||
      !finite(entry[1], -10_000_000, 10_000_000)
    )
      return undefined;
    result.push([entry[0], entry[1]]);
  }
  return result;
}

/** The server's built-in style paint (`ChartNativeLookBuiltInStyle`), bounded. */
export function normalizeChartNativeLookBuiltInStyle(raw: unknown): ChartNativeLookBuiltInStyle | undefined {
  if (!isRecord(raw) || !Number.isInteger(raw.style) || !finite(raw.style, 1, 48)) return undefined;
  const result: ChartNativeLookBuiltInStyle = { style: raw.style as number };
  if (isHex(raw.textColor)) result.ink = raw.textColor.toLowerCase();
  const seriesLine = transforms(raw.seriesLine);
  if (seriesLine?.length) result.seriesLine = seriesLine;
  // Styles 1-8 paint nothing past the series' colours: only their line
  // series' first-line-style transforms are stated (X-257).
  const gallery = (raw.style as number) >= 9;
  const fillPath = isRecord(raw.fill) && isRecord(raw.fill.path) && ["circle", "rect", "shape"].includes(String(raw.fill.path.kind))
    ? raw.fill.path
    : undefined;
  if (gallery && isRecord(raw.fill) && (finite(raw.fill.angle, 0, 360) || fillPath) && Array.isArray(raw.fill.stops)) {
    const fillStops = raw.fill.stops.slice(0, 8).map((stop) =>
      isRecord(stop) && finite(stop.position, 0, 1) ? { position: stop.position, transforms: transforms(stop.transforms) } : undefined,
    );
    if (fillStops.length >= 2 && fillStops.every((stop) => stop?.transforms)) {
      const stops = fillStops as { position: number; transforms: ChartNativeLookTransform[] }[];
      if (finite(raw.fill.angle, 0, 360)) {
        result.fill = { angle: raw.fill.angle as number, stops };
      } else if (fillPath) {
        const focus = isRecord(fillPath.fillToRect) ? fillPath.fillToRect : {};
        const rect: { l?: number; t?: number; r?: number; b?: number } = {};
        for (const side of ["l", "t", "r", "b"] as const) {
          if (finite(focus[side], -1_000_000, 1_000_000)) rect[side] = focus[side] as number;
        }
        result.fill = {
          path: { kind: fillPath.kind as "circle" | "rect" | "shape", ...(Object.keys(rect).length ? { fillToRect: rect } : {}) },
          stops,
        };
      }
    }
  }
  if (gallery && isRecord(raw.line) && finite(raw.line.widthPt, 0, 48)) {
    const lineTransforms = transforms(raw.line.transforms);
    if (isHex(raw.line.color)) result.line = { widthPt: raw.line.widthPt, color: raw.line.color.toLowerCase() };
    else if (lineTransforms) result.line = { widthPt: raw.line.widthPt, transforms: lineTransforms };
  }
  const shadow = raw.shadow;
  if (
    gallery &&
    isRecord(shadow) &&
    isHex(shadow.color) &&
    finite(shadow.opacity, 0, 1) &&
    finite(shadow.blur, 0, 100) &&
    finite(shadow.distance, 0, 100) &&
    finite(shadow.direction, 0, 360)
  ) {
    result.shadow = {
      color: shadow.color.toLowerCase(),
      opacity: shadow.opacity,
      blur: shadow.blur,
      distance: shadow.distance,
      direction: shadow.direction,
    };
  }
  if (Array.isArray(raw.own)) {
    const own: NonNullable<ChartNativeLookBuiltInStyle["own"]> = [];
    for (const entry of raw.own.slice(0, 255)) {
      if (!isRecord(entry) || !Number.isInteger(entry.series) || !finite(entry.series, 0, 254)) continue;
      const item: NonNullable<ChartNativeLookBuiltInStyle["own"]>[number] = { series: entry.series as number };
      for (const key of ["fill", "line", "effect", "unfilled"] as const) if (entry[key] === true) item[key] = true;
      if (Array.isArray(entry.points)) {
        const points = entry.points.filter((point): point is number => Number.isInteger(point) && finite(point, 0, 100_000));
        if (points.length) item.points = points.slice(0, 1_000);
      }
      own.push(item);
    }
    if (own.length) result.own = own;
  }
  if (Array.isArray(raw.slots) && raw.slots.length <= 255 && raw.slots.every((slot) => Number.isInteger(slot) && finite(slot, 0, 254))) {
    result.slots = (raw.slots as number[]).slice();
  }
  return result.fill || result.line || result.shadow || result.seriesLine ? result : undefined;
}

/** A chart's data labels' text insets (pt, each 0..100), or undefined. */
export function chartLabelInsets(raw: unknown): ChartNativeLook["dataLabelInsets"] {
  if (!isRecord(raw)) return undefined;
  const sides = ["l", "t", "r", "b"] as const;
  if (!sides.every((side) => finite(raw[side], 0, 100))) return undefined;
  return { l: raw.l as number, t: raw.t as number, r: raw.r as number, b: raw.b as number };
}

/** A bounded reading of the server's look; anything malformed is simply left out. */
export function normalizeChartNativeLook(raw: unknown): ChartNativeLook | undefined {
  if (!isRecord(raw)) return undefined;
  const result: ChartNativeLook = {};
  if (raw.labelRoom === true) result.labelRoom = true;
  const insets = chartLabelInsets(raw.dataLabelInsets);
  if (insets) result.dataLabelInsets = insets;
  const builtInStyle = normalizeChartNativeLookBuiltInStyle(raw.builtInStyle);
  if (builtInStyle) result.builtInStyle = builtInStyle;
  if (typeof raw.styleId === "string" && /^\d{1,6}$/.test(raw.styleId)) result.styleId = raw.styleId;
  if (typeof raw.colorMethod === "string" && /^[a-zA-Z]{1,32}$/.test(raw.colorMethod)) {
    result.colorMethod = raw.colorMethod;
  }
  const chartArea = area(raw.chartArea);
  if (chartArea) result.chartArea = chartArea;
  const plotArea = area(raw.plotArea);
  if (plotArea) result.plotArea = plotArea;
  const legend = { ...area(raw.legend), ...text(raw.legend) };
  if (Object.keys(legend).length) result.legend = legend;
  for (const key of ["title", "axisTitle", "dataLabel"] as const) {
    const value = text(raw[key]);
    if (value) result[key] = value;
  }
  if (isRecord(raw.axis)) {
    const axis: ChartNativeLook["axis"] = { ...text(raw.axis) };
    const axisLine = line(raw.axis.line);
    if (axisLine) axis.line = axisLine;
    if (Object.keys(axis).length) result.axis = axis;
  }
  const gridline = line(raw.gridline);
  if (gridline) result.gridline = gridline;
  const minorGridline = line(raw.minorGridline);
  if (minorGridline) result.minorGridline = minorGridline;
  const connector = line(raw.connector);
  if (connector) result.connector = connector;
  if (raw.connectorLines === false) result.connectorLines = false;
  if (
    raw.parentLabelLayout === "overlapping" ||
    raw.parentLabelLayout === "banner" ||
    raw.parentLabelLayout === "none"
  ) {
    result.parentLabelLayout = raw.parentLabelLayout;
  }
  const map = nativeMap(raw.map);
  if (map) result.map = map;
  const funnel = nativeFunnel(raw.funnel);
  if (funnel) result.funnel = funnel;
  if (Array.isArray(raw.points)) {
    const points = raw.points.slice(0, 48).map((point) => area(point) ?? {});
    if (points.some((point) => point.fill || point.outline)) result.points = points;
  }
  if (
    Array.isArray(raw.themeAccents) &&
    raw.themeAccents.length === 6 &&
    raw.themeAccents.every(isHex)
  ) {
    result.themeAccents = raw.themeAccents.map((accent) => accent.toLowerCase());
  }
  if (typeof raw.colorPalette === "string" && /^[a-z]+-[1-9]$/.test(raw.colorPalette)) {
    result.colorPalette = raw.colorPalette;
  }
  return Object.keys(result).length ? result : undefined;
}

const PT_TO_PX = PX_PER_PT;

function relativeLuminance(hex: string): number {
  const channel = (offset: number) => {
    const value = Number.parseInt(hex.slice(offset, offset + 2), 16) / 255;
    return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5);
}

function contrastRatio(left: string, right: string): number {
  const [light, dark] = [relativeLuminance(left), relativeLuminance(right)].sort((a, b) => b - a);
  return (light + 0.05) / (dark + 0.05);
}

/** Every flat colour an area paints with: its fill and each gradient stop. */
export function nativeLookAreaFills(area: ChartNativeLookArea | undefined): string[] {
  if (!area) return [];
  return [
    ...(area.fill && area.fill !== "none" ? [area.fill] : []),
    ...(area.fillGradient?.stops ?? area.radialGradient?.stops ?? []).map((stop) => stop.color),
  ].filter(isHex);
}

/**
 * The ink Office prints a series' inside data labels in. The style's label
 * colour stays while it reads (a 4.5 contrast ratio) against every fill the
 * labels sit on; otherwise the whole series takes the opposite automatic ink,
 * #FEFEFE for a dark style colour and #424242 for a light one. Excel and Word
 * print a pareto's columns, a waterfall's steps and a funnel's stages so.
 */
export function nativeLookInsideLabelInk(stated: string, fills: readonly string[]): string {
  if (!isHex(stated)) return stated;
  if (fills.filter(isHex).every((fill) => contrastRatio(stated, fill) >= 4.5)) return stated;
  return relativeLuminance(stated) < 0.179 ? "#fefefe" : "#424242";
}

/** A look colour at the opacity its style gives it; an opaque colour stays as it is. */
export function nativeLookRgba(color: string, opacity: number | undefined): string {
  if (opacity === undefined || opacity >= 1 || !isHex(color)) return color;
  const channel = (offset: number) => Number.parseInt(color.slice(offset, offset + 2), 16);
  return `rgba(${channel(1)}, ${channel(3)}, ${channel(5)}, ${Math.max(0, opacity)})`;
}

/** Points to CSS pixels, the way the shared chart paint already converts. */
export function nativeLookPx(pt: number | undefined, fallback: number): number {
  return pt === undefined ? fallback : Math.max(0, pt * PT_TO_PX);
}

/**
 * A DrawingML preset pattern as a tile the canvas can repeat. Only the presets
 * Office's chart styles use are drawn; any other reads as its background.
 */
export function nativeLookPatternTile(pattern: ChartNativeLookPattern): string | undefined {
  const { preset, fg, bg } = pattern;
  const size = 8;
  const diagonal = (up: boolean, width: number) =>
    up
      ? `<path d="M-2 ${size + 2} L${size + 2} -2 M-2 2 L2 -2 M${size - 2} ${size + 2} L${size + 2} ${size - 2}" stroke="${fg}" stroke-width="${width}"/>`
      : `<path d="M-2 -2 L${size + 2} ${size + 2} M${size - 2} -2 L${size + 2} 2 M-2 ${size - 2} L2 ${size + 2}" stroke="${fg}" stroke-width="${width}"/>`;
  const body: Record<string, string> = {
    ltDnDiag: diagonal(false, 1),
    dnDiag: diagonal(false, 1.5),
    dkDnDiag: diagonal(false, 2.5),
    wdDnDiag: diagonal(false, 3),
    ltUpDiag: diagonal(true, 1),
    upDiag: diagonal(true, 1.5),
    dkUpDiag: diagonal(true, 2.5),
    wdUpDiag: diagonal(true, 3),
    ltHorz: `<rect x="0" y="3.5" width="${size}" height="1" fill="${fg}"/>`,
    horz: `<rect x="0" y="3" width="${size}" height="2" fill="${fg}"/>`,
    dkHorz: `<rect x="0" y="2" width="${size}" height="4" fill="${fg}"/>`,
    narHorz: `<rect x="0" y="1.5" width="${size}" height="1" fill="${fg}"/><rect x="0" y="5.5" width="${size}" height="1" fill="${fg}"/>`,
    ltVert: `<rect x="3.5" y="0" width="1" height="${size}" fill="${fg}"/>`,
    vert: `<rect x="3" y="0" width="2" height="${size}" fill="${fg}"/>`,
    dkVert: `<rect x="2" y="0" width="4" height="${size}" fill="${fg}"/>`,
    narVert: `<rect x="1.5" y="0" width="1" height="${size}" fill="${fg}"/><rect x="5.5" y="0" width="1" height="${size}" fill="${fg}"/>`,
    cross: `<rect x="0" y="3.5" width="${size}" height="1" fill="${fg}"/><rect x="3.5" y="0" width="1" height="${size}" fill="${fg}"/>`,
    smGrid: `<rect x="0" y="3.5" width="${size}" height="1" fill="${fg}"/><rect x="3.5" y="0" width="1" height="${size}" fill="${fg}"/>`,
    lgGrid: `<rect x="0" y="0" width="${size}" height="1" fill="${fg}"/><rect x="0" y="0" width="1" height="${size}" fill="${fg}"/>`,
    dotGrid: `<rect x="3.5" y="3.5" width="1" height="1" fill="${fg}"/>`,
    diagCross: diagonal(false, 1) + diagonal(true, 1),
    pct5: `<circle cx="4" cy="4" r="0.6" fill="${fg}"/>`,
    pct10: `<circle cx="4" cy="4" r="0.9" fill="${fg}"/>`,
    pct20: `<circle cx="2" cy="2" r="0.9" fill="${fg}"/><circle cx="6" cy="6" r="0.9" fill="${fg}"/>`,
    pct25: `<circle cx="2" cy="2" r="1" fill="${fg}"/><circle cx="6" cy="6" r="1" fill="${fg}"/>`,
    pct30: `<circle cx="2" cy="2" r="1.1" fill="${fg}"/><circle cx="6" cy="6" r="1.1" fill="${fg}"/>`,
    pct40: `<circle cx="2" cy="2" r="1.3" fill="${fg}"/><circle cx="6" cy="6" r="1.3" fill="${fg}"/>`,
    pct50: `<circle cx="2" cy="2" r="1.4" fill="${fg}"/><circle cx="6" cy="6" r="1.4" fill="${fg}"/>`,
  };
  const marks = body[preset];
  if (!marks) return undefined;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">` +
    `<rect width="${size}" height="${size}" fill="${bg}"/>${marks}</svg>`;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

/** The ECharts colour for an area: flat, linear, radial, or a repeated pattern tile. */
export function nativeLookAreaColor(value: ChartNativeLookArea | undefined): unknown {
  if (!value) return undefined;
  if (value.fillGradient) return officeChartGradientColor(value.fillGradient);
  if (value.radialGradient) {
    return {
      type: "radial",
      x: 0.5,
      y: 0.5,
      r: 0.85,
      colorStops: value.radialGradient.stops.map((stop) => ({
        offset: stop.position,
        color: stop.color,
      })),
    };
  }
  if (value.pattern) {
    const tile = nativeLookPatternTile(value.pattern);
    if (tile) return { image: tile, repeat: "repeat" };
    return value.pattern.bg;
  }
  if (value.fill === "none") return "transparent";
  return value.fill === undefined ? undefined : nativeLookRgba(value.fill, value.fillOpacity);
}

/** The ECharts lineStyle a look line draws. */
export function nativeLookLineStyle(value: ChartNativeLookLine | undefined): Record<string, unknown> {
  if (!value) return {};
  return {
    color: value.outline === "none" ? "transparent" : nativeLookRgba(value.outline, value.outlineOpacity),
    ...(value.outlineWidthPt === undefined ? {} : { width: nativeLookPx(value.outlineWidthPt, 1) }),
    ...(value.dash
      ? { type: /dot/i.test(value.dash) && !/dash/i.test(value.dash) ? "dotted" : "dashed" }
      : {}),
  };
}

/** The ECharts textStyle a look text carries; caps are the caller's to apply to the string. */
export function nativeLookTextStyle(value: ChartNativeLookText | undefined): Record<string, unknown> {
  if (!value) return {};
  return {
    ...(value.color ? { color: value.color } : {}),
    ...(value.sizePt === undefined ? {} : { fontSize: nativeLookPx(value.sizePt, 12) }),
    ...(value.bold ? { fontWeight: "bold" } : {}),
  };
}

/** The item paint (fill plus border) for one data point of the look. */
export function nativeLookItemStyle(point: ChartNativeLookArea | undefined): Record<string, unknown> {
  if (!point) return {};
  const color = nativeLookAreaColor(point);
  return {
    ...(color === undefined ? {} : { color }),
    ...(point.outline === undefined
      ? {}
      : point.outline === "none"
        ? { borderWidth: 0 }
        : {
            borderColor: nativeLookRgba(point.outline, point.outlineOpacity),
            borderWidth: nativeLookPx(point.outlineWidthPt, 1),
          }),
  };
}

/** The flat colours of the look's points, for a palette: the first stop of a gradient stands for it. */
export function nativeLookPalette(value: ChartNativeLook | undefined): string[] | undefined {
  const colours = (value?.points ?? [])
    .map((point) => (point.fill && point.fill !== "none" ? point.fill : point.outline))
    .filter((colour): colour is string => Boolean(colour) && colour !== "none");
  return colours.length ? colours : undefined;
}
