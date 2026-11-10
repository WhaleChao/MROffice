/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry for the PDF viewer's MEASURE tool (Acrobat "Measure ▸
 * Distance / Perimeter / Area", WPS 测量). The viewer captures a sequence of
 * click points in page-relative percent [0..100] (top-left origin, the same
 * coordinate space highlights / shapes / ink use) and a {@link MeasurePageSize}
 * (the pdf.js scale-1.0 viewport, i.e. PDF user-space points). This module turns
 * those into a real-world DISTANCE, PERIMETER, or AREA under a user-chosen SCALE
 * (e.g. "1 in on the page = 10 ft in reality").
 *
 * Everything here is DOM-free / pdf.js-free so it unit-tests in isolation and can
 * back both the live overlay readout and a persisted measurement annotation. It
 * mirrors the style of `shapes.ts` / `crop-box.ts`: small, defensive, total
 * functions that never throw on malformed input.
 *
 * Coordinate model:
 *   - A point's percent x/y is mapped to PDF points via `pt = pct/100 * dim`.
 *   - PDF points convert to inches at 72 pt/in (the PDF unit), to other display
 *     units via {@link UNIT_PER_INCH}.
 *   - The SCALE multiplies the on-page physical length by `outputPerPage` to give
 *     the real-world length in `outputUnit` (Acrobat's calibration ratio).
 */

/** Display units the measure readout can use (page side + real-world side). */
export type MeasureUnit = "pt" | "in" | "cm" | "mm" | "ft" | "m";

export const MEASURE_UNITS: readonly MeasureUnit[] = [
  "pt",
  "in",
  "cm",
  "mm",
  "ft",
  "m",
] as const;

/**
 * How many of each unit fit in ONE inch. PDF user space is 72 points / inch;
 * 1 in = 2.54 cm = 25.4 mm; 1 ft = 1/12 in; 1 m = 1/0.0254 in. Used to convert a
 * length measured in PDF points into any display unit and back.
 */
export const UNIT_PER_INCH: Readonly<Record<MeasureUnit, number>> = {
  pt: 72,
  in: 1,
  cm: 2.54,
  mm: 25.4,
  ft: 1 / 12,
  m: 0.0254,
};

/** Short label suffix shown after a measurement value, per unit. */
export const UNIT_LABEL: Readonly<Record<MeasureUnit, string>> = {
  pt: "pt",
  in: "in",
  cm: "cm",
  mm: "mm",
  ft: "ft",
  m: "m",
};

/** The three measurement modes Acrobat's Measure tool offers. */
export type MeasureMode = "distance" | "perimeter" | "area";

export const MEASURE_MODES: readonly MeasureMode[] = [
  "distance",
  "perimeter",
  "area",
] as const;

/** Narrow an arbitrary value to a measure mode (defaults to "distance"). */
export function asMeasureMode(value: unknown): MeasureMode {
  return value === "perimeter" || value === "area" ? value : "distance";
}

/** Narrow an arbitrary value to a display unit (defaults to "in"). */
export function asMeasureUnit(value: unknown): MeasureUnit {
  return typeof value === "string" &&
    (MEASURE_UNITS as readonly string[]).includes(value)
    ? (value as MeasureUnit)
    : "in";
}

/** The pdf.js scale-1.0 page size in PDF user-space points (= PageRecord). */
export interface MeasurePageSize {
  /** Page width in PDF points (scale-1.0 viewport width). */
  width: number;
  /** Page height in PDF points (scale-1.0 viewport height). */
  height: number;
}

/** A click point in page-relative percent [0..100], top-left origin. */
export interface MeasurePoint {
  x: number;
  y: number;
}

/**
 * A measurement scale (Acrobat's calibration). `pageUnit` length on the PAGE
 * maps to `outputUnit` length in REALITY by the ratio `outputPerPage`. Example:
 * "1 in = 10 ft" ⇒ `{ pageUnit: "in", outputUnit: "ft", outputPerPage: 10 }`.
 * A 1:1 default (`{ pageUnit: "in", outputUnit: "in", outputPerPage: 1 }`) means
 * the readout is the physical on-page size.
 */
export interface MeasureScale {
  pageUnit: MeasureUnit;
  outputUnit: MeasureUnit;
  /** Real-world units per one page unit; must be a finite positive number. */
  outputPerPage: number;
}

/** A sensible identity scale: report the physical on-page size in inches. */
export const DEFAULT_MEASURE_SCALE: MeasureScale = {
  pageUnit: "in",
  outputUnit: "in",
  outputPerPage: 1,
};

export const MEASURE_MIN_RATIO = 1e-6;
export const MEASURE_MAX_RATIO = 1e9;

/** Clamp / default a raw scale ratio into a usable positive number. */
export function clampMeasureRatio(value: unknown): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n) || n <= 0) return 1;
  if (n < MEASURE_MIN_RATIO) return MEASURE_MIN_RATIO;
  if (n > MEASURE_MAX_RATIO) return MEASURE_MAX_RATIO;
  return n;
}

/** Validate + normalize a raw scale into a usable {@link MeasureScale}. */
export function sanitizeMeasureScale(value: unknown): MeasureScale {
  const record =
    value && typeof value === "object" ? (value as Record<string, unknown>) : {};
  return {
    pageUnit: asMeasureUnit(record.pageUnit),
    outputUnit: asMeasureUnit(record.outputUnit),
    outputPerPage: clampMeasureRatio(record.outputPerPage),
  };
}

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

/** Map a percent point onto PDF user-space points for the given page size. */
export function percentPointToPdf(
  point: MeasurePoint,
  page: MeasurePageSize,
): { x: number; y: number } | null {
  if (
    !isFinitePercent(point.x) ||
    !isFinitePercent(point.y) ||
    !Number.isFinite(page.width) ||
    !Number.isFinite(page.height) ||
    page.width <= 0 ||
    page.height <= 0
  ) {
    return null;
  }
  return {
    x: (point.x / 100) * page.width,
    y: (point.y / 100) * page.height,
  };
}

/**
 * The straight-line distance between two percent points expressed in PDF points
 * (the page's physical user-space length, BEFORE the calibration scale). Returns
 * null for malformed input.
 */
export function segmentLengthPt(
  a: MeasurePoint,
  b: MeasurePoint,
  page: MeasurePageSize,
): number | null {
  const pa = percentPointToPdf(a, page);
  const pb = percentPointToPdf(b, page);
  if (!pa || !pb) return null;
  return Math.hypot(pb.x - pa.x, pb.y - pa.y);
}

/**
 * Sum the straight-line lengths of a polyline (`points` in order) in PDF points.
 * When `closed`, the segment from the last point back to the first is included
 * (perimeter). Fewer than two points ⇒ 0. A single malformed point aborts to
 * null (the whole measurement is untrustworthy).
 */
export function polylineLengthPt(
  points: ReadonlyArray<MeasurePoint>,
  page: MeasurePageSize,
  closed = false,
): number | null {
  if (!Array.isArray(points) || points.length < 2) return 0;
  let total = 0;
  for (let i = 0; i < points.length - 1; i += 1) {
    const seg = segmentLengthPt(points[i], points[i + 1], page);
    if (seg === null) return null;
    total += seg;
  }
  if (closed) {
    const back = segmentLengthPt(points[points.length - 1], points[0], page);
    if (back === null) return null;
    total += back;
  }
  return total;
}

/**
 * Signed polygon area (PDF points²) via the shoelace formula over percent points
 * mapped to PDF space; the absolute value is the unsigned area. Fewer than three
 * points ⇒ 0. A malformed point aborts to null.
 */
export function polygonAreaPt2(
  points: ReadonlyArray<MeasurePoint>,
  page: MeasurePageSize,
): number | null {
  if (!Array.isArray(points) || points.length < 3) return 0;
  const mapped: Array<{ x: number; y: number }> = [];
  for (const p of points) {
    const pt = percentPointToPdf(p, page);
    if (!pt) return null;
    mapped.push(pt);
  }
  let twiceArea = 0;
  for (let i = 0; i < mapped.length; i += 1) {
    const cur = mapped[i];
    const next = mapped[(i + 1) % mapped.length];
    twiceArea += cur.x * next.y - next.x * cur.y;
  }
  return Math.abs(twiceArea) / 2;
}

/** Convert a length in PDF points into a display unit. */
export function pointsToUnit(points: number, unit: MeasureUnit): number {
  if (!Number.isFinite(points)) return 0;
  // points -> inches (÷72) -> unit (× per-inch).
  return (points / 72) * UNIT_PER_INCH[unit];
}

/**
 * Apply a calibration scale to a PHYSICAL on-page LENGTH (in PDF points) to get
 * the real-world length in `scale.outputUnit`. The page length is first expressed
 * in `scale.pageUnit`, then multiplied by `outputPerPage`.
 */
export function scaleLength(lengthPt: number, scale: MeasureScale): number {
  if (!Number.isFinite(lengthPt)) return 0;
  const inPageUnits = pointsToUnit(lengthPt, scale.pageUnit);
  return inPageUnits * scale.outputPerPage;
}

/**
 * Apply a calibration scale to a PHYSICAL on-page AREA (in PDF points²) to get
 * the real-world area in `scale.outputUnit`². The linear ratio (output units per
 * page unit) is SQUARED for area, matching Acrobat's area measurement.
 */
export function scaleArea(areaPt2: number, scale: MeasureScale): number {
  if (!Number.isFinite(areaPt2)) return 0;
  // points² -> pageUnit² : (pt→pageUnit linear factor)².
  const linearPerPt = pointsToUnit(1, scale.pageUnit);
  const inPageUnits2 = areaPt2 * linearPerPt * linearPerPt;
  return inPageUnits2 * scale.outputPerPage * scale.outputPerPage;
}

/** Round to at most `decimals` places, trimming trailing zeros. */
function roundTrim(value: number, decimals: number): string {
  if (!Number.isFinite(value)) return "0";
  const factor = 10 ** Math.max(0, Math.min(8, Math.trunc(decimals)));
  const rounded = Math.round(value * factor) / factor;
  // toString already trims trailing zeros; guard -0 → 0.
  return (rounded === 0 ? 0 : rounded).toString();
}

export const MEASURE_DEFAULT_DECIMALS = 2;

/**
 * Format a real-world LENGTH for display, e.g. `"12.5 ft"`. `decimals` defaults
 * to {@link MEASURE_DEFAULT_DECIMALS}. Pure text — safe to render into a text
 * node.
 */
export function formatLength(
  value: number,
  unit: MeasureUnit,
  decimals = MEASURE_DEFAULT_DECIMALS,
): string {
  return `${roundTrim(value, decimals)} ${UNIT_LABEL[unit]}`;
}

/**
 * Format a real-world AREA for display, e.g. `"30 ft²"`. The unit is suffixed
 * with a superscript-two (`²`), Acrobat's area convention.
 */
export function formatArea(
  value: number,
  unit: MeasureUnit,
  decimals = MEASURE_DEFAULT_DECIMALS,
): string {
  return `${roundTrim(value, decimals)} ${UNIT_LABEL[unit]}²`;
}

/** Result of measuring a point set under a scale. */
export interface MeasureResult {
  mode: MeasureMode;
  /** The raw physical on-page value (points for length modes, points² for area). */
  rawPt: number;
  /** The scaled real-world value in {@link MeasureScale.outputUnit}. */
  value: number;
  /** A ready-to-render label, e.g. `"12.5 ft"` or `"30 ft²"`. */
  label: string;
}

/** Map a complete measurement point set into one new percent-space box. */
export function transformMeasurePointsToBox(
  points: ReadonlyArray<MeasurePoint>,
  target: Readonly<{ x: number; y: number; width: number; height: number }>,
): MeasurePoint[] | null {
  if (
    !Array.isArray(points) ||
    points.length < 2 ||
    [target.x, target.y, target.width, target.height].some(
      (value) => !Number.isFinite(value),
    ) ||
    target.x < 0 ||
    target.y < 0 ||
    target.width < 0 ||
    target.height < 0 ||
    target.x + target.width > 100.000001 ||
    target.y + target.height > 100.000001
  ) {
    return null;
  }
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  if ([...xs, ...ys].some((value) => !Number.isFinite(value))) return null;
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const width = Math.max(...xs) - left;
  const height = Math.max(...ys) - top;
  const scaleX = width <= 0.000001 ? 1 : target.width / width;
  const scaleY = height <= 0.000001 ? 1 : target.height / height;
  if (
    (width <= 0.000001 && target.width > 0.000001) ||
    (height <= 0.000001 && target.height > 0.000001) ||
    scaleX < 0.01 ||
    scaleY < 0.01 ||
    scaleX > 100 ||
    scaleY > 100
  ) {
    return null;
  }
  return points.map((point) => ({
    x: target.x + (point.x - left) * scaleX,
    y: target.y + (point.y - top) * scaleY,
  }));
}

/** Rotate measurement vertices clockwise in displayed-page physical space. */
export function rotateMeasurePoints(
  points: ReadonlyArray<MeasurePoint>,
  page: MeasurePageSize,
  rotationDegrees: number,
): MeasurePoint[] | null {
  if (
    !Array.isArray(points) ||
    points.length < 2 ||
    !Number.isFinite(page.width) ||
    !Number.isFinite(page.height) ||
    page.width <= 0 ||
    page.height <= 0 ||
    !Number.isFinite(rotationDegrees)
  ) {
    return null;
  }
  const normalized = ((rotationDegrees % 360) + 360) % 360;
  if (normalized <= 0.000001 || 360 - normalized <= 0.000001) return null;
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  if ([...xs, ...ys].some((value) => !Number.isFinite(value))) return null;
  const center = {
    x: ((Math.min(...xs) + Math.max(...xs)) / 200) * page.width,
    y: ((Math.min(...ys) + Math.max(...ys)) / 200) * page.height,
  };
  const radians = (normalized * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  const rotated = points.map((point) => {
    const x = (point.x / 100) * page.width - center.x;
    const y = (point.y / 100) * page.height - center.y;
    return {
      x: ((center.x + cosine * x - sine * y) / page.width) * 100,
      y: ((center.y + sine * x + cosine * y) / page.height) * 100,
    };
  });
  return rotated.some(
    (point) =>
      point.x < -0.000001 ||
      point.y < -0.000001 ||
      point.x > 100.000001 ||
      point.y > 100.000001,
  )
    ? null
    : rotated;
}

/**
 * Measure a set of percent points under a mode + scale, returning the physical
 * value, the scaled real-world value, and a formatted label. Returns null when
 * the geometry is malformed (a NaN point) or there aren't enough points for the
 * mode (≥2 for distance/perimeter, ≥3 for area). The caller decides whether to
 * show a partial readout while the user is still placing points.
 *
 *   - distance  : straight line between the FIRST and LAST point.
 *   - perimeter : sum of consecutive segments + closing segment (closed loop).
 *   - area      : shoelace polygon area (closed loop).
 */
export function measurePoints(
  mode: MeasureMode,
  points: ReadonlyArray<MeasurePoint>,
  page: MeasurePageSize,
  scale: MeasureScale = DEFAULT_MEASURE_SCALE,
  decimals = MEASURE_DEFAULT_DECIMALS,
): MeasureResult | null {
  if (!Array.isArray(points)) return null;
  const safeScale = sanitizeMeasureScale(scale);
  if (mode === "area") {
    if (points.length < 3) return null;
    const rawPt = polygonAreaPt2(points, page);
    if (rawPt === null) return null;
    const value = scaleArea(rawPt, safeScale);
    return {
      mode,
      rawPt,
      value,
      label: formatArea(value, safeScale.outputUnit, decimals),
    };
  }
  if (points.length < 2) return null;
  // Distance uses just the endpoints; perimeter sums (and closes) the path.
  const rawPt =
    mode === "distance"
      ? segmentLengthPt(points[0], points[points.length - 1], page)
      : polylineLengthPt(points, page, true);
  if (rawPt === null) return null;
  const value = scaleLength(rawPt, safeScale);
  return {
    mode,
    rawPt,
    value,
    label: formatLength(value, safeScale.outputUnit, decimals),
  };
}
