/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF freehand ink / drawing annotations.
 *
 * An ink annotation is a single polyline ("stroke") drawn on a page with the
 * pointer, captured as a sequence of points in page-relative *percentage*
 * coordinates [0..100] (resolution-independent, matching how highlight rects and
 * sticky notes are stored) carrying a solid stroke color + width. This module
 * holds everything with no DOM / React dependency so it can be unit-tested in
 * isolation and reused by both the viewer and (potentially) a backend mirror:
 *
 *   - the persisted annotation shape + its sanitizer (mirrors
 *     `sanitizeNotePayload` in sticky-notes.ts so a malformed/legacy payload
 *     never crashes the reader),
 *   - point clamping + per-stroke point caps (drop runaway captures onto the
 *     page, keep the persisted polyline bounded),
 *   - width clamping + an allow-listed color set (so a crafted payload can't
 *     smuggle arbitrary CSS or an unbounded stroke into the inline SVG),
 *   - `buildInkPayload` to produce exactly what we POST (so persisted bytes
 *     match what we render),
 *   - `inkPathD` to turn a normalized stroke into an SVG path `d` (percentages
 *     map directly onto a `viewBox="0 0 100 100"` overlay).
 *
 * The persisted annotation type is `pdf_ink` and is additive to the existing
 * `pdf_highlight` / `pdf_note` types: old PDFs keep loading unchanged.
 *
 * NOTE: this is distinct from the page-snapshot `pdfInkStrokeSchema`
 * (native-page-schemas.ts), which stores flattened 0..1 fraction points for the
 * page-ops snapshot. The runtime annotation here mirrors `pdf_note` (percent
 * coords + allow-listed color) so it round-trips through the datasets
 * annotations endpoint exactly like sticky notes do.
 */

import { inkStrokePathData } from "../office-ink-smooth";
import { isSafeHexColor } from "./safe-color";
import { normalizeXfdfSourceId } from "./xfdf-import";

/** Default stroke color used when a payload omits / supplies an unknown color. */
export const DEFAULT_INK_COLOR = "#ef4444";

/**
 * Allowed stroke colors. Kept as an allow-list (like the note marker colors) so
 * a crafted payload can't inject arbitrary CSS into the inline SVG `stroke`.
 * These are solid (opaque) pen colors, unlike the translucent highlight colors.
 */
export const PDF_INK_COLORS: readonly string[] = [
  "#ef4444", // red
  "#3b82f6", // blue
  "#22c55e", // green
  "#eab308", // yellow
  "#a855f7", // purple
  "#111827", // near-black ink
];

const PDF_INK_COLOR_SET = new Set(PDF_INK_COLORS);

/** Stroke width bounds, in page-percentage units (resolution-independent). */
export const PDF_INK_MIN_WIDTH = 0.1;
export const PDF_INK_MAX_WIDTH = 12;
export const DEFAULT_INK_WIDTH = 0.6;











export function inkWidthPagePercent(
  width: number,
  page: { width: number; height: number } | undefined,
): number {
  const pageWidth = page?.width ?? 0;
  const pageHeight = page?.height ?? 0;
  if (!Number.isFinite(width) || !(pageWidth > 0) || !(pageHeight > 0)) {
    return clampInkWidth(width);
  }
  const shorter = Math.min(pageWidth, pageHeight);
  const longer = Math.max(pageWidth, pageHeight);
  const thickness = Math.max(0.25, Math.min(100, (width * shorter) / longer));
  return (thickness / shorter) * 100;
}

/** Max number of points kept per stroke (excess is decimated to fit). */
export const PDF_INK_MAX_POINTS = 1000;
/** Minimum points for a usable stroke (a lone point is not a polyline). */
export const PDF_INK_MIN_POINTS = 2;
/** Max number of ink strokes kept per page. */
export const PDF_INK_MAX_STROKES_PER_PAGE = 500;

/** A single point in page-relative percent [0..100]. */
export interface InkPoint {
  x: number;
  y: number;
  




  p?: number;
}

/** A pressure as a point keeps it: a finite number in [0, 1]. */
export function inkPointPressure(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 1
    ? value
    : undefined;
}

/** Whether a stroke's pressures vary enough to draw at a varying width. */
export function inkStrokeHasPressure(points: ReadonlyArray<InkPoint>): boolean {
  let min = Infinity;
  let max = -Infinity;
  let seen = 0;
  for (const point of points) {
    if (point.p === undefined) continue;
    seen += 1;
    min = Math.min(min, point.p);
    max = Math.max(max, point.p);
  }
  return seen >= 2 && max - min > 0.05;
}

export interface InkAnnotation {
  id: string;
  /** 1-based page number. */
  page: number;
  /** Stroke color; always one of {@link PDF_INK_COLORS}. */
  color: string;
  /** Stroke width in page-percentage units; clamped to the width bounds. */
  width: number;
  /** Polyline points in page-relative percent [0..100]. */
  points: InkPoint[];
  xfdfSourceId?: string;
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
 * rejects everything that isn't an inert hex literal.
 */
export function isInkColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (PDF_INK_COLOR_SET.has(value) || isSafeHexColor(value))
  );
}

/**
 * Clamp a stroke width into [{@link PDF_INK_MIN_WIDTH}, {@link PDF_INK_MAX_WIDTH}].
 * Non-finite / missing values fall back to {@link DEFAULT_INK_WIDTH}.
 */
export function clampInkWidth(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_INK_WIDTH;
  }
  if (value < PDF_INK_MIN_WIDTH) return PDF_INK_MIN_WIDTH;
  if (value > PDF_INK_MAX_WIDTH) return PDF_INK_MAX_WIDTH;
  return value;
}

/**
 * Reduce a list of points to at most {@link PDF_INK_MAX_POINTS}, evenly sampling
 * across the stroke (keeping the first + last) so the overall shape survives
 * decimation. Non-finite coordinates are clamped onto the page. Points whose
 * coordinates are not numbers are dropped. The result preserves draw order.
 */
export function clampPoints(
  points: ReadonlyArray<{ x: unknown; y: unknown; p?: unknown }>,
  max = PDF_INK_MAX_POINTS,
): InkPoint[] {
  if (!Array.isArray(points) || points.length === 0) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : PDF_INK_MAX_POINTS;
  // Keep only structurally valid points first (numeric x/y); a pen's
  // pressure rides along when it is one.
  const valid: InkPoint[] = [];
  for (const p of points) {
    if (!p || typeof p !== "object") continue;
    const x = (p as { x: unknown }).x;
    const y = (p as { y: unknown }).y;
    if (typeof x !== "number" || typeof y !== "number") continue;
    const pressure = inkPointPressure((p as { p?: unknown }).p);
    valid.push(
      pressure === undefined
        ? { x: clampPercent(x), y: clampPercent(y) }
        : { x: clampPercent(x), y: clampPercent(y), p: pressure },
    );
  }
  if (valid.length <= cap) return valid;
  // Even-stride decimation that always keeps first + last point.
  const result: InkPoint[] = [];
  const stride = (valid.length - 1) / (cap - 1);
  for (let i = 0; i < cap; i += 1) {
    const idx = Math.round(i * stride);
    result.push(valid[Math.min(idx, valid.length - 1)]);
  }
  return result;
}

/**
 * Normalize a raw captured stroke into a clean {@link InkAnnotation} polyline,
 * or return null when it is too short to be a usable stroke (after dropping
 * invalid points). Color is forced onto the allow-list; width clamped; points
 * clamped + capped. Does NOT assign an id — that is the caller's concern (the
 * server assigns the persisted id; the client may use a temporary one).
 */
export function normalizeStroke(input: {
  points: ReadonlyArray<{ x: unknown; y: unknown; p?: unknown }>;
  color?: unknown;
  width?: unknown;
}): { color: string; width: number; points: InkPoint[] } | null {
  const points = clampPoints(input.points);
  if (points.length < PDF_INK_MIN_POINTS) return null;
  return {
    color: isInkColor(input.color) ? input.color : DEFAULT_INK_COLOR,
    width: clampInkWidth(input.width),
    points,
  };
}

/**
 * Validate + normalize a persisted `pdf_ink` payload into a usable annotation,
 * or return null when it is structurally invalid (mirrors
 * `sanitizeNotePayload`). Unknown/missing color falls back to the default; width
 * is clamped; out-of-range / non-numeric points are dropped; a stroke that ends
 * up shorter than {@link PDF_INK_MIN_POINTS} is rejected (returns null) rather
 * than rendered as a degenerate dot.
 */
export function sanitizeInkPayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): InkAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const page = Number((payload as Record<string, unknown>).page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const rawPoints = (payload as Record<string, unknown>).points;
  if (!Array.isArray(rawPoints)) return null;
  // Reject grossly oversized blobs outright (defense before clamping), mirroring
  // the note sanitizer's oversized-text guard.
  if (rawPoints.length > PDF_INK_MAX_POINTS * 4) return null;
  // Only keep points that are well-formed AND already in range; out-of-range
  // points are dropped (not snapped) so corrupt data shrinks rather than warps.
  const inRange = rawPoints.filter(
    (p): p is { x: number; y: number; p?: unknown } =>
      Boolean(p) &&
      typeof p === "object" &&
      isFinitePercent((p as { x: unknown }).x) &&
      isFinitePercent((p as { y: unknown }).y),
  );
  const points = clampPoints(inRange);
  if (points.length < PDF_INK_MIN_POINTS) return null;
  const colorRaw = (payload as Record<string, unknown>).color;
  const color = isInkColor(colorRaw) ? colorRaw : DEFAULT_INK_COLOR;
  const width = clampInkWidth((payload as Record<string, unknown>).width);
  const xfdfSourceId = normalizeXfdfSourceId(
    (payload as Record<string, unknown>).xfdfSourceId,
  );
  return {
    id,
    page,
    color,
    width,
    points,
    ...(xfdfSourceId ? { xfdfSourceId } : {}),
  };
}

/**
 * Build the payload to POST for a new ink stroke. Points are clamped + capped,
 * width clamped, color forced onto the allow-list — so what we persist matches
 * what we render.
 */
export function buildInkPayload(input: {
  page: number;
  color?: string;
  width?: number;
  points: ReadonlyArray<{ x: number; y: number; p?: number }>;
  xfdfSourceId?: string;
}): {
  page: number;
  color: string;
  width: number;
  points: InkPoint[];
  xfdfSourceId?: string;
} {
  const xfdfSourceId = normalizeXfdfSourceId(input.xfdfSourceId);
  return {
    page: input.page,
    color: isInkColor(input.color) ? input.color : DEFAULT_INK_COLOR,
    width: clampInkWidth(input.width),
    points: clampPoints(input.points),
    ...(xfdfSourceId ? { xfdfSourceId } : {}),
  };
}

/**
 * Turn a normalized stroke's points into an SVG path `d` string. The points are
 * percentages, so the path is meant to be rendered inside an SVG overlay with
 * `viewBox="0 0 100 100"` (preserveAspectRatio="none") sized to the page. Each
 * coordinate is rounded to 3 decimals to keep the serialized path compact.
 * Returns "" for an empty / single-point stroke (nothing to draw).
 */
export function inkPathD(points: ReadonlyArray<InkPoint>): string {
  if (!Array.isArray(points) || points.length < PDF_INK_MIN_POINTS) return "";
  const fmt = (n: number) => {
    const rounded = Math.round(clampPercent(n) * 1000) / 1000;
    return String(rounded);
  };
  let d = `M ${fmt(points[0].x)} ${fmt(points[0].y)}`;
  for (let i = 1; i < points.length; i += 1) {
    d += ` L ${fmt(points[i].x)} ${fmt(points[i].y)}`;
  }
  return d;
}

/**
 * The same stroke as a pen draws it: one curve through its samples
 * (lib/office-ink-smooth), for the page overlay. {@link inkPathD} is the
 * polyline the stroke states; a quick loop drawn with a pen is six or eight
 * samples, and joined by straight lines it read as a polygon. `page` is the
 * page's size, so a sharp turn is judged as it looks on the page and not in
 * the stretched 0..100 box. Returns "" where `inkPathD` does.
 */
export function inkCurvePathD(
  points: ReadonlyArray<InkPoint>,
  page?: { width: number; height: number },
): string {
  if (!Array.isArray(points) || points.length < PDF_INK_MIN_POINTS) return "";
  const unit =
    page && page.width > 0 && page.height > 0 ? { x: page.width / 100, y: page.height / 100 } : undefined;
  return inkStrokePathData(
    points.map((point) => ({ x: clampPercent(point.x), y: clampPercent(point.y) })),
    { unit, precision: 3 },
  );
}

/**
 * Cap the number of strokes kept for a single page to
 * {@link PDF_INK_MAX_STROKES_PER_PAGE}, dropping the OLDEST strokes first (the
 * tail — most recent — is what a user is actively working on). Returns a new
 * array; the input is left untouched.
 */
export function capStrokesPerPage<T>(
  strokes: ReadonlyArray<T>,
  max = PDF_INK_MAX_STROKES_PER_PAGE,
): T[] {
  if (!Array.isArray(strokes)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : PDF_INK_MAX_STROKES_PER_PAGE;
  if (strokes.length <= cap) return strokes.slice();
  return strokes.slice(strokes.length - cap);
}

// ---------------------------------------------------------------------------
// Capture-time stroke conditioning (smoothing / simplification / pressure) and
// the highlighter pen. Pure, unit-AGNOSTIC geometry over `{x, y}` points: the
// same functions work on the reader's percent [0..100] coords and the native
// editor's fraction [0..1] coords, as long as `tolerance` is expressed in the
// same units as the points. Applied at CAPTURE time only — already-persisted
// strokes are rendered verbatim, so smoothing never rewrites old documents.
// ---------------------------------------------------------------------------

/** The two supported ink pen types. Absent/unknown on a stroke means "pen". */
export type InkPenType = "pen" | "highlighter";

/**
 * Stroke opacity used to RENDER and to EXPORT a highlighter stroke (a pen
 * stroke stays fully opaque). Shared by the SVG overlays and the flatten
 * planner so the download matches the screen.
 */
export const PDF_HIGHLIGHTER_OPACITY = 0.35;

/**
 * Normalize a persisted/captured pen-type value: only the literal
 * `"highlighter"` opts into the translucent pen; anything else (including
 * legacy strokes with no penType at all) is a plain pen.
 */
export function normalizeInkPenType(value: unknown): InkPenType {
  return value === "highlighter" ? "highlighter" : "pen";
}

/** Default Catmull-Rom subdivisions inserted between captured points. */
export const INK_SMOOTH_SUBDIVISIONS = 4;
/** Hard cap on subdivisions so a crafted option can't explode point counts. */
const INK_SMOOTH_MAX_SUBDIVISIONS = 8;

/**
 * Ramer-Douglas-Peucker point simplification: drop points that deviate from
 * the surrounding chord by less than `tolerance` (same units as the points).
 * Keeps the first + last point and preserves draw order. A non-positive /
 * non-finite tolerance or a segment-or-shorter input returns a plain copy.
 * Non-finite coordinates are treated as zero-distance so they never dominate
 * the split choice (the capture path clamps them before this runs anyway).
 */
export function simplifyInkPoints<P extends InkPoint>(
  points: ReadonlyArray<P>,
  tolerance: number,
): P[] {
  if (!Array.isArray(points)) return [];
  if (
    points.length < 3 ||
    typeof tolerance !== "number" ||
    !Number.isFinite(tolerance) ||
    tolerance <= 0
  ) {
    return points.slice();
  }
  const keep = new Array<boolean>(points.length).fill(false);
  keep[0] = true;
  keep[points.length - 1] = true;
  // Iterative RDP (explicit stack) so a pathological 10k-point stroke can't
  // blow the call stack.
  const stack: Array<[number, number]> = [[0, points.length - 1]];
  while (stack.length > 0) {
    const [start, end] = stack.pop()!;
    if (end - start < 2) continue;
    const sx = points[start].x;
    const sy = points[start].y;
    const dx = points[end].x - sx;
    const dy = points[end].y - sy;
    const chordLength = Math.hypot(dx, dy);
    let maxDistance = -1;
    let maxIndex = -1;
    for (let i = start + 1; i < end; i += 1) {
      // Perpendicular distance to the chord; for a degenerate (closed) chord
      // fall back to the distance from the shared endpoint.
      const distance =
        chordLength > 0
          ? Math.abs(dy * (points[i].x - sx) - dx * (points[i].y - sy)) /
            chordLength
          : Math.hypot(points[i].x - sx, points[i].y - sy);
      if (Number.isFinite(distance) && distance > maxDistance) {
        maxDistance = distance;
        maxIndex = i;
      }
    }
    if (maxIndex > 0 && maxDistance > tolerance) {
      keep[maxIndex] = true;
      stack.push([start, maxIndex], [maxIndex, end]);
    }
  }
  const result: P[] = [];
  for (let i = 0; i < points.length; i += 1) {
    if (keep[i]) result.push(points[i]);
  }
  return result;
}

/**
 * Catmull-Rom smoothing: interpolate `subdivisions` extra points into every
 * segment of the polyline using a uniform Catmull-Rom spline (endpoints
 * duplicated as phantom controls). The curve passes through EVERY input point,
 * so the stroke's shape is preserved while jagged pointer sampling is rounded
 * off. Returns a plain copy for inputs shorter than 3 points (a straight
 * segment needs no smoothing). Output points are plain `{x, y}`.
 */
export function smoothInkPoints(
  points: ReadonlyArray<InkPoint>,
  subdivisions: number = INK_SMOOTH_SUBDIVISIONS,
): InkPoint[] {
  if (!Array.isArray(points)) return [];
  const steps =
    Number.isInteger(subdivisions) && subdivisions > 0
      ? Math.min(subdivisions, INK_SMOOTH_MAX_SUBDIVISIONS)
      : 0;
  if (points.length < 3 || steps === 0) {
    return points.map((p) => ({ x: p.x, y: p.y }));
  }
  const at = (index: number): InkPoint =>
    points[Math.min(points.length - 1, Math.max(0, index))];
  const result: InkPoint[] = [{ x: points[0].x, y: points[0].y }];
  for (let i = 0; i < points.length - 1; i += 1) {
    const p0 = at(i - 1);
    const p1 = at(i);
    const p2 = at(i + 1);
    const p3 = at(i + 2);
    for (let s = 1; s <= steps; s += 1) {
      const t = s / (steps + 1);
      const t2 = t * t;
      const t3 = t2 * t;
      result.push({
        x:
          0.5 *
          (2 * p1.x +
            (-p0.x + p2.x) * t +
            (2 * p0.x - 5 * p1.x + 4 * p2.x - p3.x) * t2 +
            (-p0.x + 3 * p1.x - 3 * p2.x + p3.x) * t3),
        y:
          0.5 *
          (2 * p1.y +
            (-p0.y + p2.y) * t +
            (2 * p0.y - 5 * p1.y + 4 * p2.y - p3.y) * t2 +
            (-p0.y + 3 * p1.y - 3 * p2.y + p3.y) * t3),
      });
    }
    result.push({ x: p2.x, y: p2.y });
  }
  return result;
}

/**
 * Full capture-time conditioning pipeline: simplify the raw pointer samples
 * (dropping sensor jitter), then Catmull-Rom-smooth the survivors, then
 * decimate back under `maxPoints` (even stride, first + last kept) so the
 * persisted stroke stays bounded. Unit-agnostic — `tolerance` must be in the
 * same units as the points.
 */
export function conditionCapturedInkPoints(
  points: ReadonlyArray<InkPoint>,
  options: {
    tolerance: number;
    subdivisions?: number;
    maxPoints?: number;
  },
): InkPoint[] {
  if (!Array.isArray(points)) return [];
  const simplified = simplifyInkPoints(points, options.tolerance);
  const smoothed = smoothInkPoints(simplified, options.subdivisions);
  const cap =
    Number.isInteger(options.maxPoints) && (options.maxPoints as number) > 1
      ? (options.maxPoints as number)
      : PDF_INK_MAX_POINTS;
  if (smoothed.length <= cap) return smoothed;
  const result: InkPoint[] = [];
  const stride = (smoothed.length - 1) / (cap - 1);
  for (let i = 0; i < cap; i += 1) {
    const idx = Math.round(i * stride);
    result.push(smoothed[Math.min(idx, smoothed.length - 1)]);
  }
  return result;
}

/** Pressure-width factor bounds: a feather touch halves the pen, hard press ×1.5. */
export const INK_PRESSURE_MIN_FACTOR = 0.5;
export const INK_PRESSURE_MAX_FACTOR = 1.5;

/**
 * Turn the pressure samples collected across a stroke into a single width
 * factor (PER-STROKE AVERAGE pressure — not per-point width). PointerEvents
 * report pressure in [0,1] with 0.5 as the no-sensor default, so 0.5 maps to
 * factor 1 (no change). Samples outside (0,1] (mice report 0 when a button
 * state is unknown) are ignored; no usable samples -> factor 1.
 */
export function pressureWidthFactor(
  samples: ReadonlyArray<number>,
): number {
  if (!Array.isArray(samples) || samples.length === 0) return 1;
  let sum = 0;
  let count = 0;
  for (const sample of samples) {
    if (
      typeof sample !== "number" ||
      !Number.isFinite(sample) ||
      sample <= 0 ||
      sample > 1
    ) {
      continue;
    }
    sum += sample;
    count += 1;
  }
  if (count === 0) return 1;
  const factor = (sum / count) / 0.5;
  if (factor < INK_PRESSURE_MIN_FACTOR) return INK_PRESSURE_MIN_FACTOR;
  if (factor > INK_PRESSURE_MAX_FACTOR) return INK_PRESSURE_MAX_FACTOR;
  return factor;
}
