/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF *shape* annotations (Acrobat "Drawing Markups" parity:
 * rectangle / line / arrow / ellipse).
 *
 * A shape annotation is a single vector primitive drawn on a page with the
 * pointer, captured from a drag's two corners in page-relative *percentage*
 * coordinates [0..100] with a TOP-LEFT origin (resolution-independent, matching
 * how highlight rects / sticky notes / ink / redaction boxes are stored). Unlike
 * a redaction (an OPAQUE black blackout box), a shape is NON-DESTRUCTIVE: it is
 * a coloured STROKE (rectangle outline, straight line, arrow, ellipse outline)
 * with no fill, drawn over the page without obscuring the content beneath.
 *
 * This module mirrors `ink.ts` / `redaction.ts` (box/polyline annotations): it
 * holds everything with no DOM / React / pdf-lib dependency so it can be
 * unit-tested in isolation and reused by both the viewer and the flatten-export
 * planner:
 *
 *   - the persisted annotation shape + its sanitizer ({@link sanitizeShapePayload}),
 *   - point clamping (keep the two corners on-page) + a drag→shape normalizer
 *     ({@link normalizeShapeDrag}) that rejects a stray click / micro-drag,
 *   - stroke-width clamping + an allow-listed color set (so a crafted payload
 *     can't smuggle arbitrary CSS into the inline SVG / an unbounded stroke),
 *   - `buildShapePayload` to produce exactly what we POST (persisted bytes match
 *     what we render),
 *   - geometry helpers shared by BOTH the SVG overlay and the pdf-lib bake:
 *     {@link shapeBox} (the axis-aligned bounding box for rect/ellipse),
 *     {@link arrowHeadPoints} (the two barb endpoints for an arrow),
 *     {@link ellipsePath} (an SVG path `d` for an ellipse outline).
 *
 * The persisted annotation type is `pdf_shapes` and is additive to the existing
 * `pdf_highlight` / `pdf_note` / `pdf_ink` / `pdf_redactions` / `pdf_signatures`
 * annotations — old PDFs with none of these keep loading unchanged, and shapes
 * are entirely optional.
 */

import { isSafeHexColor } from "./safe-color";
import { normalizeXfdfSourceId } from "./xfdf";

/** The four supported shape primitives. */
export type ShapeKind =
  | "rect"
  | "line"
  | "arrow"
  | "ellipse"
  | "cloud"
  | "polygon"
  | "polyline";

/** Set form of the supported kinds, for validation. */
export const SHAPE_KINDS: readonly ShapeKind[] = [
  "rect",
  "line",
  "arrow",
  "ellipse",
  "cloud",
  "polygon",
  "polyline",
];
const SHAPE_KIND_SET = new Set<string>(SHAPE_KINDS);

/** True when a value is one of the supported shape kinds. */
export function isShapeKind(value: unknown): value is ShapeKind {
  return typeof value === "string" && SHAPE_KIND_SET.has(value);
}

/**
 * Allowed stroke colors. An allow-list (like the ink colors) so a crafted
 * payload can't inject arbitrary CSS into the inline SVG `stroke`. Solid opaque
 * markup colors.
 */
export const PDF_SHAPE_COLORS: readonly string[] = [
  "#ef4444", // red
  "#3b82f6", // blue
  "#22c55e", // green
  "#eab308", // yellow
  "#a855f7", // purple
  "#111827", // near-black
];
const PDF_SHAPE_COLOR_SET = new Set(PDF_SHAPE_COLORS);

/** Default stroke color used when a payload omits / supplies an unknown color. */
export const DEFAULT_SHAPE_COLOR = "#ef4444";

/** Longest comment a drawing carries (mirrors the server's bound). */
export const MAX_SHAPE_COMMENT = 20_000;

/** Stroke width bounds, in page-percentage units (resolution-independent). */
export const PDF_SHAPE_MIN_WIDTH = 0.1;
export const PDF_SHAPE_MAX_WIDTH = 12;
export const DEFAULT_SHAPE_WIDTH = 0.6;

export const SHAPE_LINE_STYLES = ["solid", "dashed", "dotted"] as const;
export type ShapeLineStyle = (typeof SHAPE_LINE_STYLES)[number];
export const SHAPE_LINE_ENDINGS = [
  "none",
  "openArrow",
  "closedArrow",
  "circle",
  "square",
] as const;
export type ShapeLineEnding = (typeof SHAPE_LINE_ENDINGS)[number];

export type ShapeLineCap = "butt" | "round";
export interface ShapeLineDashPattern {
  dashArray: [number, number];
  lineCap: ShapeLineCap;
}

// Each arrow ending is its own discriminated member so consumers that branch
// on "openArrow" and then "closedArrow" narrow the remainder to circle/square.
export type ShapeLineMarkerGeometry =
  | { kind: "openArrow"; tip: ShapePoint; left: ShapePoint; right: ShapePoint }
  | {
      kind: "closedArrow";
      tip: ShapePoint;
      left: ShapePoint;
      right: ShapePoint;
    }
  | { kind: "circle"; cx: number; cy: number; radius: number }
  | { kind: "square"; x: number; y: number; width: number; height: number };

const SHAPE_LINE_MAX_DASH = 1_000;
const SHAPE_LINE_MAX_MARKER_SIZE = 1_000;
const SHAPE_LINE_MAX_COORDINATE = 1_000_000;

function isShapeLineStyle(value: unknown): value is ShapeLineStyle {
  return (SHAPE_LINE_STYLES as readonly unknown[]).includes(value);
}

function isShapeLineEnding(value: unknown): value is ShapeLineEnding {
  return (SHAPE_LINE_ENDINGS as readonly unknown[]).includes(value);
}

export function shapeKindSupportsLineEndings(kind: ShapeKind): boolean {
  return kind === "line" || kind === "arrow" || kind === "polyline";
}

/** Canonical PDF dash pattern shared by derived-output consumers. */
export function shapeLineDashPattern(
  style: ShapeLineStyle,
  thickness: number,
): ShapeLineDashPattern | null {
  if (style === "solid") return null;
  if (
    !isShapeLineStyle(style) ||
    !Number.isFinite(thickness) ||
    thickness <= 0
  ) {
    return null;
  }
  const bounded = (value: number): number =>
    Math.min(SHAPE_LINE_MAX_DASH, Math.max(0.1, value));
  return style === "dashed"
    ? {
        dashArray: [
          bounded(Math.max(2, thickness * 4)),
          bounded(Math.max(1, thickness * 2)),
        ],
        lineCap: "butt",
      }
    : {
        dashArray: [bounded(thickness), bounded(Math.max(1, thickness * 2))],
        lineCap: "round",
      };
}

/** Marker scale used by every derived-output line-ending renderer. */
export function shapeLineMarkerSize(thickness: number): number | null {
  if (!Number.isFinite(thickness) || thickness <= 0) return null;
  return Math.min(SHAPE_LINE_MAX_MARKER_SIZE, Math.max(4, thickness * 4));
}

function isBoundedShapeLinePoint(point: ShapePoint): boolean {
  return (
    Number.isFinite(point.x) &&
    Number.isFinite(point.y) &&
    Math.abs(point.x) <= SHAPE_LINE_MAX_COORDINATE &&
    Math.abs(point.y) <= SHAPE_LINE_MAX_COORDINATE
  );
}

/** Canonical geometry for a semantic line-ending marker. */
export function shapeLineMarkerGeometry(
  ending: ShapeLineEnding,
  input: { tip: ShapePoint; toward: ShapePoint; size: number },
): ShapeLineMarkerGeometry | null {
  if (
    ending === "none" ||
    !isShapeLineEnding(ending) ||
    !isBoundedShapeLinePoint(input.tip) ||
    !isBoundedShapeLinePoint(input.toward) ||
    !Number.isFinite(input.size) ||
    input.size <= 0 ||
    input.size > SHAPE_LINE_MAX_MARKER_SIZE
  ) {
    return null;
  }
  const dx = input.toward.x - input.tip.x;
  const dy = input.toward.y - input.tip.y;
  const length = Math.hypot(dx, dy);
  if (!Number.isFinite(length) || length <= 0.000001) return null;
  const ux = dx / length;
  const uy = dy / length;
  const px = -uy;
  const py = ux;
  const left = {
    x: input.tip.x + ux * input.size + px * input.size * 0.55,
    y: input.tip.y + uy * input.size + py * input.size * 0.55,
  };
  const right = {
    x: input.tip.x + ux * input.size - px * input.size * 0.55,
    y: input.tip.y + uy * input.size - py * input.size * 0.55,
  };
  if (!isBoundedShapeLinePoint(left) || !isBoundedShapeLinePoint(right)) {
    return null;
  }
  if (ending === "openArrow" || ending === "closedArrow") {
    return { kind: ending, tip: { ...input.tip }, left, right };
  }
  const radius = input.size * 0.42;
  if (
    Math.abs(input.tip.x) + radius > SHAPE_LINE_MAX_COORDINATE ||
    Math.abs(input.tip.y) + radius > SHAPE_LINE_MAX_COORDINATE
  ) {
    return null;
  }
  return ending === "circle"
    ? { kind: "circle", cx: input.tip.x, cy: input.tip.y, radius }
    : {
        kind: "square",
        x: input.tip.x - radius,
        y: input.tip.y - radius,
        width: radius * 2,
        height: radius * 2,
      };
}

/**
 * Fill opacity bounds. A shape's interior fill is OPTIONAL (rect / ellipse only)
 * and translucent so the page content stays partly visible beneath it (matching
 * how Acrobat / WPS draw a semi-transparent shape fill). `0` (or no fill color)
 * means "no fill" — a pure outline, the prior behavior. The default applied when
 * a fill color is chosen without an explicit opacity is a light tint.
 */
export const PDF_SHAPE_MIN_FILL_OPACITY = 0;
export const PDF_SHAPE_MAX_FILL_OPACITY = 1;
export const DEFAULT_SHAPE_FILL_OPACITY = 0.25;
export const DEFAULT_SHAPE_OPACITY = 1;
export const DEFAULT_SHAPE_CORNER_RADIUS = 0;
export const PDF_SHAPE_MAX_CORNER_RADIUS = 50;

/** Max number of shapes kept per page (oldest dropped first when exceeded). */
export const PDF_SHAPE_MAX_PER_PAGE = 500;

/**
 * A visual mask changes only the page appearance. It is deliberately distinct
 * from a `pdf_redactions` mark and never implies that covered source content was
 * removed. Keeping this purpose on the persisted shape lets every reader flow
 * preserve that safety distinction after reload, transform and paste.
 */
export const PDF_SHAPE_VISUAL_MASK_PURPOSE = "visual_mask" as const;
export type ShapePurpose = typeof PDF_SHAPE_VISUAL_MASK_PURPOSE;

function isShapePurpose(value: unknown): value is ShapePurpose {
  return value === PDF_SHAPE_VISUAL_MASK_PURPOSE;
}

/**
 * Minimum drag extent (page-percentage units) for a drag to count as a shape.
 * Below this in BOTH dimensions the drag is a stray click and is rejected (so a
 * tap doesn't leave a zero-size shape). A line/arrow only needs ONE dimension to
 * exceed it (a horizontal/vertical line has ~0 extent on the other axis).
 */
export const SHAPE_MIN_DRAG = 0.5;

export interface ShapeAnnotation {
  id: string;
  /** Stable XFDF `name` retained across server-generated reader ids. */
  xfdfSourceId?: string;
  /** 1-based page number. */
  page: number;
  kind: ShapeKind;
  /** Start corner / endpoint in page-relative percent [0..100], top-left origin. */
  x1: number;
  y1: number;
  /** End corner / endpoint in page-relative percent [0..100], top-left origin. */
  x2: number;
  y2: number;
  /** Exact polygon vertices in page-percent space; absent for box/line kinds. */
  points?: ShapePoint[];
  

  comment?: string;
  /** Stroke color; always one of {@link PDF_SHAPE_COLORS}. */
  color: string;
  /** Stroke width in page-percentage units; clamped to the width bounds. */
  width: number;
  lineStyle?: ShapeLineStyle;
  /** Endpoint styles apply only to line/arrow/polyline shapes. */
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  /** Whole-object opacity in [0,1]; absent means fully opaque. */
  opacity?: number;
  /** Rounded-corner percentage [0,50], relative to the rect's shorter side. */
  cornerRadius?: number;
  /** Explicitly non-secure visual-cover identity; absent for ordinary shapes. */
  purpose?: ShapePurpose;
  /**
   * OPTIONAL interior fill color (rect / ellipse only). An inert hex literal (or
   * an allow-listed shape color); absent / null means NO fill (a pure outline,
   * the original behavior). line / arrow never carry a fill.
   */
  fillColor?: string | null;
  /**
   * Fill opacity in [{@link PDF_SHAPE_MIN_FILL_OPACITY},
   * {@link PDF_SHAPE_MAX_FILL_OPACITY}]. Only meaningful when {@link fillColor}
   * is set; `0` is treated as "no fill" regardless of color.
   */
  fillOpacity?: number;
}

export interface ShapePoint {
  x: number;
  y: number;
}

export function resolveShapeLineStyle(shape: {
  lineStyle?: unknown;
}): ShapeLineStyle {
  return isShapeLineStyle(shape.lineStyle) ? shape.lineStyle : "solid";
}

export function resolveShapeLineEndings(shape: {
  kind: ShapeKind;
  lineStart?: unknown;
  lineEnd?: unknown;
}): { start: ShapeLineEnding; end: ShapeLineEnding } {
  if (!shapeKindSupportsLineEndings(shape.kind)) {
    return { start: "none", end: "none" };
  }
  return {
    start: isShapeLineEnding(shape.lineStart) ? shape.lineStart : "none",
    end: isShapeLineEnding(shape.lineEnd)
      ? shape.lineEnd
      : shape.kind === "arrow"
        ? "openArrow"
        : "none",
  };
}

export const PDF_SHAPE_MAX_POLYGON_POINTS = 200;
const PDF_SHAPE_MIN_POLYGON_AREA = 0.25;

export function normalizePolygonPoints(value: unknown): ShapePoint[] | null {
  if (!Array.isArray(value)) return null;
  const points: ShapePoint[] = [];
  for (const raw of value) {
    if (!raw || typeof raw !== "object") return null;
    const x = Number((raw as { x?: unknown }).x);
    const y = Number((raw as { y?: unknown }).y);
    if (!isFinitePercent(x) || !isFinitePercent(y)) return null;
    const previous = points.at(-1);
    if (!previous || previous.x !== x || previous.y !== y) points.push({ x, y });
    if (points.length > PDF_SHAPE_MAX_POLYGON_POINTS) return null;
  }
  if (
    points.length > 3 &&
    points[0].x === points.at(-1)!.x &&
    points[0].y === points.at(-1)!.y
  ) {
    points.pop();
  }
  if (points.length < 3) return null;
  let twiceArea = 0;
  for (let index = 0; index < points.length; index += 1) {
    const current = points[index];
    const next = points[(index + 1) % points.length];
    twiceArea += current.x * next.y - next.x * current.y;
  }
  return Math.abs(twiceArea) / 2 >= PDF_SHAPE_MIN_POLYGON_AREA
    ? points
    : null;
}

/** Normalize an OPEN polyline without closing its first and last vertices. */
export function normalizePolylinePoints(value: unknown): ShapePoint[] | null {
  if (!Array.isArray(value)) return null;
  const points: ShapePoint[] = [];
  let pathLength = 0;
  for (const raw of value) {
    if (!raw || typeof raw !== "object") return null;
    const x = Number((raw as { x?: unknown }).x);
    const y = Number((raw as { y?: unknown }).y);
    if (!isFinitePercent(x) || !isFinitePercent(y)) return null;
    const previous = points.at(-1);
    if (!previous || previous.x !== x || previous.y !== y) {
      if (previous) pathLength += Math.hypot(x - previous.x, y - previous.y);
      points.push({ x, y });
    }
    if (points.length > PDF_SHAPE_MAX_POLYGON_POINTS) return null;
  }
  return points.length >= 2 && pathLength >= SHAPE_MIN_DRAG ? points : null;
}

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

/** Clamp a single coordinate value into the on-page percentage range. */
function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value < 0) return 0;
  if (value > 100) return 100;
  return value;
}

/**
 * True when the value is a usable stroke color: a member of the quick-pick
 * allow-list OR a strict hex color (`#rrggbb`, what an `<input type=color>` free
 * swatch emits). Accepting strict hex lets the user pick any pen color without
 * re-opening the CSS-injection hole the allow-list closed — `isSafeHexColor`
 * rejects everything that isn't an inert hex literal. Translucent rgba is NOT
 * allowed here: shape strokes are opaque, and the export bakes a solid line.
 */
export function isShapeColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (PDF_SHAPE_COLOR_SET.has(value) || isSafeHexColor(value))
  );
}







export function shapeWidthToPoints(
  width: number,
  pageMinPoints: number,
): number {
  return (width / 100) * pageMinPoints;
}

/** The inverse of {@link shapeWidthToPoints}, clamped to the stroke range. */
export function shapeWidthFromPoints(
  points: number,
  pageMinPoints: number,
): number {
  return clampShapeWidth((points / pageMinPoints) * 100);
}

/**
 * Clamp a stroke width into [{@link PDF_SHAPE_MIN_WIDTH},
 * {@link PDF_SHAPE_MAX_WIDTH}]. Non-finite / missing → {@link DEFAULT_SHAPE_WIDTH}.
 */
export function clampShapeWidth(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SHAPE_WIDTH;
  }
  if (value < PDF_SHAPE_MIN_WIDTH) return PDF_SHAPE_MIN_WIDTH;
  if (value > PDF_SHAPE_MAX_WIDTH) return PDF_SHAPE_MAX_WIDTH;
  return value;
}

/**
 * Whether a shape kind can carry an interior fill. Only the two CLOSED shapes
 * (rect / ellipse) do; an open line / arrow has no interior, so a fill on those
 * is meaningless and is always dropped.
 */
export function kindSupportsFill(kind: unknown): boolean {
  return kind === "rect" || kind === "ellipse";
}

/**
 * Deterministic scalloped cloud outline shared by SVG preview and PDF export.
 * Points follow a gently rippled ellipse inside the drag box; the closed path
 * is resolution-independent and stable across readers.
 */
export function cloudOutlinePoints(
  shape: { x1: number; y1: number; x2: number; y2: number },
  segments = 36,
): Array<{ x: number; y: number }> {
  const box = shapeBox(shape);
  if (box.width <= 0 || box.height <= 0) return [];
  const count = Math.max(12, Math.min(72, Math.trunc(segments)));
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const rx = box.width / 2;
  const ry = box.height / 2;
  return Array.from({ length: count }, (_, index) => {
    const angle = (index / count) * Math.PI * 2;
    const ripple = 0.9 + 0.1 * Math.cos(angle * 9);
    return {
      x: cx + rx * ripple * Math.cos(angle),
      y: cy + ry * ripple * Math.sin(angle),
    };
  });
}

/**
 * Clamp a fill opacity into [{@link PDF_SHAPE_MIN_FILL_OPACITY},
 * {@link PDF_SHAPE_MAX_FILL_OPACITY}]. Non-finite / missing →
 * {@link DEFAULT_SHAPE_FILL_OPACITY} (a light tint) so a chosen fill color is
 * visible by default.
 */
export function clampShapeFillOpacity(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SHAPE_FILL_OPACITY;
  }
  if (value < PDF_SHAPE_MIN_FILL_OPACITY) return PDF_SHAPE_MIN_FILL_OPACITY;
  if (value > PDF_SHAPE_MAX_FILL_OPACITY) return PDF_SHAPE_MAX_FILL_OPACITY;
  return value;
}

/** Resolve whole-object opacity; invalid/missing values preserve legacy opaque output. */
export function resolveShapeOpacity(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SHAPE_OPACITY;
  }
  return Math.min(1, Math.max(0, value));
}

/** Resolve rect-only corner rounding; invalid/missing values preserve square corners. */
export function resolveShapeCornerRadius(
  kind: ShapeKind,
  value: unknown,
): number {
  if (kind !== "rect" || typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SHAPE_CORNER_RADIUS;
  }
  return Math.min(PDF_SHAPE_MAX_CORNER_RADIUS, Math.max(0, value));
}

/**
 * Resolve the EFFECTIVE interior fill of a shape, or null when there is none.
 * A fill applies only to a closed kind (rect / ellipse) with a usable fill color
 * AND a strictly-positive opacity; a line/arrow, a missing/invalid color, or a
 * zero opacity all collapse to null (a pure outline). The returned opacity is
 * clamped, so callers (SVG render + flatten bake) get the same answer.
 */
export function resolveShapeFill(shape: {
  kind: ShapeKind;
  fillColor?: string | null;
  fillOpacity?: number;
}): { color: string; opacity: number } | null {
  if (!kindSupportsFill(shape.kind)) return null;
  if (!isShapeColor(shape.fillColor)) return null;
  const opacity = clampShapeFillOpacity(shape.fillOpacity);
  if (opacity <= 0) return null;
  return { color: shape.fillColor, opacity };
}

/**
 * The axis-aligned bounding box (top-left origin, percent) of a shape's two
 * corners. Used to render rect / ellipse outlines (which are anchored to the
 * box, not the raw drag direction) and to bound-check.
 */
export function shapeBox(shape: {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  points?: readonly ShapePoint[];
}): { x: number; y: number; width: number; height: number } {
  if (shape.points && shape.points.length >= 2) {
    const xs = shape.points.map((point) => point.x);
    const ys = shape.points.map((point) => point.y);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return {
      x,
      y,
      width: Math.max(...xs) - x,
      height: Math.max(...ys) - y,
    };
  }
  const x = Math.min(shape.x1, shape.x2);
  const y = Math.min(shape.y1, shape.y2);
  const width = Math.abs(shape.x2 - shape.x1);
  const height = Math.abs(shape.y2 - shape.y1);
  return { x, y, width, height };
}

export type ShapeGeometryTransform =
  "rotate-left" | "rotate-right" | "flip-horizontal" | "flip-vertical";

/**
 * Transform one shape in rendered-page coordinates, then translate its whole
 * AABB on-page. Points are never individually clamped or scaled.
 */
export function transformShapeGeometry(
  shape: {
    kind: ShapeKind;
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    points?: readonly ShapePoint[];
  },
  page: { width: number; height: number },
  action: ShapeGeometryTransform,
): {
  kind: ShapeKind;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  points?: ShapePoint[];
} | null {
  if (
    !Number.isFinite(page.width) ||
    !Number.isFinite(page.height) ||
    page.width <= 0 ||
    page.height <= 0 ||
    !isShapeKind(shape.kind)
  ) {
    return null;
  }
  const source = shape.points?.length
    ? shape.points
    : [
        { x: shape.x1, y: shape.y1 },
        { x: shape.x2, y: shape.y2 },
      ];
  if (
    source.length < 2 ||
    source.some(
      (point) =>
        !Number.isFinite(point.x) ||
        !Number.isFinite(point.y) ||
        point.x < 0 ||
        point.x > 100 ||
        point.y < 0 ||
        point.y > 100,
    )
  ) {
    return null;
  }
  const pixels = source.map((point) => ({
    x: (point.x / 100) * page.width,
    y: (point.y / 100) * page.height,
  }));
  const sourceXs = pixels.map((point) => point.x);
  const sourceYs = pixels.map((point) => point.y);
  const sourceWidth = Math.max(...sourceXs) - Math.min(...sourceXs);
  const sourceHeight = Math.max(...sourceYs) - Math.min(...sourceYs);
  const openKind = shapeKindSupportsLineEndings(shape.kind);
  if (
    (openKind && sourceWidth <= 0.000001 && sourceHeight <= 0.000001) ||
    (!openKind && (sourceWidth <= 0.000001 || sourceHeight <= 0.000001))
  ) {
    return null;
  }
  const cx = (Math.min(...sourceXs) + Math.max(...sourceXs)) / 2;
  const cy = (Math.min(...sourceYs) + Math.max(...sourceYs)) / 2;
  const transformed = pixels.map((point) => {
    const dx = point.x - cx;
    const dy = point.y - cy;
    if (action === "flip-horizontal") return { x: cx - dx, y: point.y };
    if (action === "flip-vertical") return { x: point.x, y: cy - dy };
    return action === "rotate-right"
      ? { x: cx - dy, y: cy + dx }
      : { x: cx + dy, y: cy - dx };
  });
  const xs = transformed.map((point) => point.x);
  const ys = transformed.map((point) => point.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  if (
    ![minX, maxX, minY, maxY].every(Number.isFinite) ||
    maxX - minX > page.width + 0.000001 ||
    maxY - minY > page.height + 0.000001
  ) {
    return null;
  }
  const shiftX = minX < 0 ? -minX : maxX > page.width ? page.width - maxX : 0;
  const shiftY = minY < 0 ? -minY : maxY > page.height ? page.height - maxY : 0;
  const result = transformed.map((point) => ({
    x: ((point.x + shiftX) / page.width) * 100,
    y: ((point.y + shiftY) / page.height) * 100,
  }));
  if (shape.points?.length) {
    const resultXs = result.map((point) => point.x);
    const resultYs = result.map((point) => point.y);
    return {
      kind: shape.kind,
      x1: Math.min(...resultXs),
      y1: Math.min(...resultYs),
      x2: Math.max(...resultXs),
      y2: Math.max(...resultYs),
      points: result,
    };
  }
  return {
    kind: shape.kind,
    x1: result[0].x,
    y1: result[0].y,
    x2: result[1].x,
    y2: result[1].y,
  };
}

/**
 * Normalize a raw drag (two corners) into a clean shape geometry for the given
 * kind, or return null when the drag is too small to be a usable shape.
 *
 * For rect / ellipse the geometry is the bounding box's two opposite corners
 * (so a backwards drag still yields a top-left box); both dimensions must exceed
 * {@link SHAPE_MIN_DRAG}. For line / arrow the two endpoints are kept AS-IS (the
 * arrow points from start→end, so direction matters); the drag passes when
 * EITHER dimension exceeds {@link SHAPE_MIN_DRAG} (so an axis-aligned line is
 * valid). Inputs are clamped onto the page first.
 */
export function normalizeShapeDrag(
  kind: ShapeKind,
  start: { x: number; y: number },
  end: { x: number; y: number },
): { x1: number; y1: number; x2: number; y2: number } | null {
  if (
    !isShapeKind(kind) ||
    !Number.isFinite(start.x) ||
    !Number.isFinite(start.y) ||
    !Number.isFinite(end.x) ||
    !Number.isFinite(end.y)
  ) {
    return null;
  }
  const sx = clampPercent(start.x);
  const sy = clampPercent(start.y);
  const ex = clampPercent(end.x);
  const ey = clampPercent(end.y);
  const dx = Math.abs(ex - sx);
  const dy = Math.abs(ey - sy);

  if (kind === "line" || kind === "arrow") {
    // A directional primitive: keep the endpoints as drawn, accept when either
    // axis has enough extent (so a perfectly horizontal/vertical line is valid).
    if (dx < SHAPE_MIN_DRAG && dy < SHAPE_MIN_DRAG) return null;
    return { x1: sx, y1: sy, x2: ex, y2: ey };
  }
  // rect / ellipse: a box — require BOTH dimensions, normalize to top-left
  // origin so the persisted corners are (min, min)→(max, max).
  if (dx < SHAPE_MIN_DRAG || dy < SHAPE_MIN_DRAG) return null;
  return {
    x1: Math.min(sx, ex),
    y1: Math.min(sy, ey),
    x2: Math.max(sx, ex),
    y2: Math.max(sy, ey),
  };
}

/**
 * Validate + normalize a persisted `pdf_shapes` payload into a usable
 * annotation, or return null when it is structurally invalid (mirrors
 * `sanitizeInkPayload` / `sanitizeRedactionPayload`). Out-of-range coords are
 * rejected (returns null) rather than silently snapped; width is clamped; an
 * unknown color falls back to the default. A degenerate shape (both extents
 * below {@link SHAPE_MIN_DRAG} for line/arrow, OR either extent below it for
 * rect/ellipse) is rejected so a zero-size shape is never rendered/baked.
 */
export function sanitizeShapePayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): ShapeAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const record = payload as Record<string, unknown>;
  const page = Number(record.page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const kind = record.kind;
  if (!isShapeKind(kind)) return null;
  const pathPoints =
    kind === "polygon"
      ? normalizePolygonPoints(record.points)
      : kind === "polyline"
        ? normalizePolylinePoints(record.points)
        : null;
  if ((kind === "polygon" || kind === "polyline") && !pathPoints) return null;
  const pathBox = pathPoints
    ? shapeBox({ x1: 0, y1: 0, x2: 0, y2: 0, points: pathPoints })
    : null;
  const x1 = pathBox?.x ?? record.x1;
  const y1 = pathBox?.y ?? record.y1;
  const x2 = pathBox ? pathBox.x + pathBox.width : record.x2;
  const y2 = pathBox ? pathBox.y + pathBox.height : record.y2;
  if (
    !isFinitePercent(x1) ||
    !isFinitePercent(y1) ||
    !isFinitePercent(x2) ||
    !isFinitePercent(y2)
  ) {
    return null;
  }
  const dx = Math.abs((x2 as number) - (x1 as number));
  const dy = Math.abs((y2 as number) - (y1 as number));
  if (kind === "line" || kind === "arrow" || kind === "polyline") {
    if (dx < SHAPE_MIN_DRAG && dy < SHAPE_MIN_DRAG) return null;
  } else if (dx < SHAPE_MIN_DRAG || dy < SHAPE_MIN_DRAG) {
    return null;
  }
  const colorRaw = record.color;
  const xfdfSourceId = normalizeXfdfSourceId(record.xfdfSourceId);
  const hasPurpose = Object.prototype.hasOwnProperty.call(record, "purpose");
  if (hasPurpose && !isShapePurpose(record.purpose)) return null;
  // Optional interior fill (rect / ellipse only). An unusable / absent fill
  // color, an unsupported kind, or a zero opacity all mean "no fill" — the
  // fields are then omitted so a legacy outline-only shape round-trips byte-for-
  // byte. A finite fill color carried on a line/arrow payload is ignored.
  const fillColorRaw = record.fillColor;
  const hasFill =
    kindSupportsFill(kind) &&
    isShapeColor(fillColorRaw) &&
    clampShapeFillOpacity(record.fillOpacity) > 0;
  const lineStyle = isShapeLineStyle(record.lineStyle)
    ? record.lineStyle
    : "solid";
  const supportsEndings = shapeKindSupportsLineEndings(kind);
  const lineStart =
    supportsEndings && isShapeLineEnding(record.lineStart)
      ? record.lineStart
      : "none";
  const lineEnd =
    supportsEndings && isShapeLineEnding(record.lineEnd)
      ? record.lineEnd
      : undefined;
  const opacity = resolveShapeOpacity(record.opacity);
  const cornerRadius = resolveShapeCornerRadius(kind, record.cornerRadius);
  return {
    id,
    ...(xfdfSourceId ? { xfdfSourceId } : null),
    page,
    kind,
    x1: x1 as number,
    y1: y1 as number,
    x2: x2 as number,
    y2: y2 as number,
    ...(pathPoints ? { points: pathPoints } : null),
    color: isShapeColor(colorRaw) ? colorRaw : DEFAULT_SHAPE_COLOR,
    width: clampShapeWidth(record.width),
    ...(lineStyle !== "solid" ? { lineStyle } : null),
    ...(lineStart !== "none" ? { lineStart } : null),
    ...(lineEnd !== undefined && (kind === "arrow" || lineEnd !== "none")
      ? { lineEnd }
      : null),
    ...(opacity !== DEFAULT_SHAPE_OPACITY ? { opacity } : null),
    ...(cornerRadius !== DEFAULT_SHAPE_CORNER_RADIUS ? { cornerRadius } : null),
    ...(hasPurpose ? { purpose: record.purpose as ShapePurpose } : null),
    ...(typeof record.comment === "string" && record.comment.trim()
      ? {
          comment: record.comment
            .replace(/\r\n?/gu, "\n")
            .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/gu, "")
            .slice(0, MAX_SHAPE_COMMENT),
        }
      : null),
    ...(hasFill
      ? {
          fillColor: fillColorRaw as string,
          fillOpacity: clampShapeFillOpacity(record.fillOpacity),
        }
      : null),
  };
}

/**
 * Build the payload to POST for a new shape. Geometry is normalized (corners
 * clamped + degenerate-rejected via {@link normalizeShapeDrag}), width clamped,
 * color forced onto the allow-list — so what we persist matches what we render.
 * Returns null when the geometry is degenerate.
 */
export function buildShapePayload(input: {
  page: number;
  kind: ShapeKind;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  points?: readonly ShapePoint[];
  xfdfSourceId?: string;
  color?: string;
  width?: number;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  opacity?: number;
  cornerRadius?: number;
  /** Optional non-secure visual-cover identity; no other purpose is accepted. */
  purpose?: ShapePurpose;
  /** Optional interior fill color (rect / ellipse only); null/absent → no fill. */
  fillColor?: string | null;
  /** Fill opacity; only persisted when a usable fill color is present. */
  fillOpacity?: number;
}): {
  page: number;
  kind: ShapeKind;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  points?: ShapePoint[];
  xfdfSourceId?: string;
  color: string;
  width: number;
  lineStyle?: ShapeLineStyle;
  lineStart?: ShapeLineEnding;
  lineEnd?: ShapeLineEnding;
  opacity?: number;
  cornerRadius?: number;
  purpose?: ShapePurpose;
  fillColor?: string;
  fillOpacity?: number;
} | null {
  if (!isShapeKind(input.kind)) return null;
  if (input.purpose !== undefined && !isShapePurpose(input.purpose)) return null;
  const pathPoints =
    input.kind === "polygon"
      ? normalizePolygonPoints(input.points)
      : input.kind === "polyline"
        ? normalizePolylinePoints(input.points)
        : null;
  if ((input.kind === "polygon" || input.kind === "polyline") && !pathPoints) {
    return null;
  }
  const pathBox = pathPoints
    ? shapeBox({ x1: 0, y1: 0, x2: 0, y2: 0, points: pathPoints })
    : null;
  const geometry = pathBox
    ? {
        x1: pathBox.x,
        y1: pathBox.y,
        x2: pathBox.x + pathBox.width,
        y2: pathBox.y + pathBox.height,
      }
    : normalizeShapeDrag(
        input.kind,
        { x: input.x1, y: input.y1 },
        { x: input.x2, y: input.y2 },
      );
  if (!geometry) return null;
  // Persist a fill only for a closed kind with a usable color + positive opacity,
  // so an outline-only shape posts exactly the prior payload (no fill keys).
  const hasFill =
    kindSupportsFill(input.kind) &&
    isShapeColor(input.fillColor) &&
    clampShapeFillOpacity(input.fillOpacity) > 0;
  const xfdfSourceId = normalizeXfdfSourceId(input.xfdfSourceId);
  const lineStyle = isShapeLineStyle(input.lineStyle)
    ? input.lineStyle
    : "solid";
  const supportsEndings = shapeKindSupportsLineEndings(input.kind);
  const lineStart =
    supportsEndings && isShapeLineEnding(input.lineStart)
      ? input.lineStart
      : "none";
  const lineEnd =
    supportsEndings && isShapeLineEnding(input.lineEnd)
      ? input.lineEnd
      : undefined;
  const opacity = resolveShapeOpacity(input.opacity);
  const cornerRadius = resolveShapeCornerRadius(input.kind, input.cornerRadius);
  return {
    page: input.page,
    kind: input.kind,
    ...(xfdfSourceId ? { xfdfSourceId } : null),
    x1: geometry.x1,
    y1: geometry.y1,
    x2: geometry.x2,
    y2: geometry.y2,
    ...(pathPoints ? { points: pathPoints } : null),
    color: isShapeColor(input.color) ? input.color : DEFAULT_SHAPE_COLOR,
    width: clampShapeWidth(input.width),
    ...(lineStyle !== "solid" ? { lineStyle } : null),
    ...(lineStart !== "none" ? { lineStart } : null),
    ...(lineEnd !== undefined && (input.kind === "arrow" || lineEnd !== "none")
      ? { lineEnd }
      : null),
    ...(opacity !== DEFAULT_SHAPE_OPACITY ? { opacity } : null),
    ...(cornerRadius !== DEFAULT_SHAPE_CORNER_RADIUS ? { cornerRadius } : null),
    ...(input.purpose ? { purpose: input.purpose } : null),
    ...(hasFill
      ? {
          fillColor: input.fillColor as string,
          fillOpacity: clampShapeFillOpacity(input.fillOpacity),
        }
      : null),
  };
}

/**
 * Default tint for an AREA HIGHLIGHT — a translucent rectangle the reader drags
 * over a figure / scanned region / chart to call it out (vs. the text-markup
 * highlight, which snaps to glyph rects). A warm yellow, the classic highlighter
 * colour; it is an allow-listed {@link PDF_SHAPE_COLORS} member so it survives
 * the fill-color guard. The stroke is the SAME color so the box reads as a solid
 * wash with a faint border rather than a hollow outline.
 */
export const DEFAULT_AREA_HIGHLIGHT_COLOR = "#eab308";
/** Default fill opacity for an area highlight (a touch lighter than a shape fill). */
export const DEFAULT_AREA_HIGHLIGHT_OPACITY = 0.3;

/**
 * Build the payload to POST for an AREA HIGHLIGHT: a {@link ShapeKind} `rect`
 * shape whose interior is a translucent fill, so it highlights an arbitrary
 * rectangular REGION (a figure, a scanned block, a chart) rather than text. It
 * is a thin, intention-revealing wrapper over {@link buildShapePayload} that
 * (a) forces the kind to `rect`, (b) always carries a fill (defaulting to the
 * highlighter tint), and (c) matches the stroke colour to the fill so the result
 * is a clean wash. Returns null for a degenerate drag (same rule as any rect
 * shape: both extents must exceed {@link SHAPE_MIN_DRAG}).
 *
 * The persisted shape is an ordinary `pdf_shapes` rect with a fill — it loads,
 * renders (translucent SVG fill) and BAKES (translucent `RectOp` under the
 * outline) through the exact same paths as a manually-filled rectangle, so no
 * new annotation type / flatten op is introduced.
 */
export function buildAreaHighlightPayload(input: {
  page: number;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  /** Highlight colour (fill + matching stroke); defaults to the yellow tint. */
  color?: string;
  /** Fill opacity; defaults to {@link DEFAULT_AREA_HIGHLIGHT_OPACITY}. */
  opacity?: number;
}): {
  page: number;
  kind: ShapeKind;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  width: number;
  fillColor?: string;
  fillOpacity?: number;
} | null {
  const color = isShapeColor(input.color)
    ? (input.color as string)
    : DEFAULT_AREA_HIGHLIGHT_COLOR;
  const fillOpacity =
    typeof input.opacity === "number" && Number.isFinite(input.opacity)
      ? clampShapeFillOpacity(input.opacity)
      : DEFAULT_AREA_HIGHLIGHT_OPACITY;
  return buildShapePayload({
    page: input.page,
    kind: "rect",
    x1: input.x1,
    y1: input.y1,
    x2: input.x2,
    y2: input.y2,
    color,
    // A thin border in the same hue so the wash has a crisp edge.
    width: DEFAULT_SHAPE_WIDTH,
    fillColor: color,
    fillOpacity,
  });
}

/**
 * Cap the number of shapes kept for a single page to
 * {@link PDF_SHAPE_MAX_PER_PAGE}, dropping the OLDEST first (the tail — most
 * recent — is what the user just added). Returns a new array; input untouched.
 * Mirrors `capStrokesPerPage` / `capRedactionsPerPage`.
 */
export function capShapesPerPage<T>(
  shapes: ReadonlyArray<T>,
  max = PDF_SHAPE_MAX_PER_PAGE,
): T[] {
  if (!Array.isArray(shapes)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : PDF_SHAPE_MAX_PER_PAGE;
  if (shapes.length <= cap) return shapes.slice();
  return shapes.slice(shapes.length - cap);
}

/**
 * Compute the two barb endpoints of an arrowhead at the LINE's end point
 * (`x2,y2`), pointing back toward the start (`x1,y1`). All coordinates are in
 * the SAME units as the input (callers pass percent for the SVG overlay and PDF
 * points for the bake). The barbs sit at `headLength` back from the tip, splayed
 * by `headAngle` (radians) on either side of the line direction.
 *
 * Returns null for a zero-length line (no direction → no arrowhead). The head
 * length passed by callers is already in the caller's unit; this function does
 * NOT clamp it (the caller sizes it relative to the line / page).
 */
export function arrowHeadPoints(
  line: { x1: number; y1: number; x2: number; y2: number },
  headLength: number,
  headAngle = Math.PI / 7,
): { left: { x: number; y: number }; right: { x: number; y: number } } | null {
  const dx = line.x2 - line.x1;
  const dy = line.y2 - line.y1;
  const len = Math.hypot(dx, dy);
  if (!Number.isFinite(len) || len <= 0) return null;
  if (!Number.isFinite(headLength) || headLength <= 0) return null;
  // Angle of the line, then back-rotate by ±headAngle to get the two barbs.
  const angle = Math.atan2(dy, dx);
  const leftAngle = angle - headAngle;
  const rightAngle = angle + headAngle;
  return {
    left: {
      x: line.x2 - headLength * Math.cos(leftAngle),
      y: line.y2 - headLength * Math.sin(leftAngle),
    },
    right: {
      x: line.x2 - headLength * Math.cos(rightAngle),
      y: line.y2 - headLength * Math.sin(rightAngle),
    },
  };
}

/**
 * SVG path `d` for an ellipse outline inscribed in a shape's bounding box. Meant
 * for an SVG overlay with `viewBox="0 0 100 100"` (percent coords). Uses two arc
 * segments. Returns "" for a zero-size box. Coordinates rounded to 3 decimals to
 * keep the serialized path compact (mirrors `inkPathD`).
 */
export function ellipsePath(shape: {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}): string {
  const box = shapeBox(shape);
  if (box.width <= 0 || box.height <= 0) return "";
  const rx = box.width / 2;
  const ry = box.height / 2;
  const cx = box.x + rx;
  const cy = box.y + ry;
  const fmt = (n: number) => String(Math.round(n * 1000) / 1000);
  // Start at the left vertex, sweep to the right vertex, then back — two arcs.
  return (
    `M ${fmt(cx - rx)} ${fmt(cy)} ` +
    `A ${fmt(rx)} ${fmt(ry)} 0 1 1 ${fmt(cx + rx)} ${fmt(cy)} ` +
    `A ${fmt(rx)} ${fmt(ry)} 0 1 1 ${fmt(cx - rx)} ${fmt(cy)} Z`
  );
}
