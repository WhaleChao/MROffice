/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeInkPenType, type InkPenType } from "@mrnote/core/web/lib/pdf/ink.ts";

/**
 * Native `/Ink` annotations discovered through pdf.js.
 *
 * pdf.js exposes the indirect object reference plus every `/InkList` path.
 * Single-path annotations can therefore be deleted by exact source object.
 * Multi-path annotations remain visible and are reported as explicit losses,
 * because rewriting an arbitrary third-party appearance stream is not safe.
 */

export const NATIVE_PDF_INK_SOURCE_KIND = "pdf_ink_annotation";
export const MAX_NATIVE_PDF_INK_ANNOTATIONS_PER_PAGE = 1_024;
export const MAX_NATIVE_PDF_INK_PATHS_PER_ANNOTATION = 256;
export const MAX_NATIVE_PDF_INK_PATHS_PER_PAGE = 2_048;
export const MAX_NATIVE_PDF_INK_POINT_VALUES = 4_000;
export const MAX_NATIVE_PDF_INK_POINT_VALUES_PER_PAGE = 40_000;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/;
const COORDINATE_EPSILON = 0.000_001;

export interface NativePdfInkAnnotationSource {
  kind: typeof NATIVE_PDF_INK_SOURCE_KIND;
  objectRef: string;
  /** Original zero-based position inside the annotation's `/InkList`. */
  pathIndex: number;
  /** Original number of paths in that `/InkList`. */
  pathCount: number;
  /** Immutable displayed-page fractions used to prove translation-only edits. */
  originalPoints?: number[];
  /** The colour and width (the stroke's own units) the file draws it in: a
   * stroke that differs from them is restyled (/C, /BS /W, a new /AP). */
  originalColor?: string;
  originalWidth?: number;
}

export interface NativePdfInkStroke {
  id: string;
  /** Zero-based source PDF page index, matching PdfInkStroke. */
  page: number;
  color: string;
  width: number;
  points: number[];
  penType: InkPenType;
  nativeSource: NativePdfInkAnnotationSource;
  




  managedId?: string;
}

export type NativePdfInkImportLossReason =
  | "invalid_object_ref"
  | "invalid_page_geometry"
  | "invalid_path_geometry"
  | "multi_path_annotation"
  | "path_limit"
  | "point_limit"
  | "annotation_limit";

export interface NativePdfInkImportLoss {
  objectRef?: string;
  reason: NativePdfInkImportLossReason;
}

export interface NativePdfInkImportResult {
  strokes: NativePdfInkStroke[];
  losses: NativePdfInkImportLoss[];
}

export interface NativePdfInkViewport {
  width: number;
  height: number;
  convertToViewportPoint: (x: number, y: number) => readonly [number, number];
}

interface RawNativePdfInkAnnotation {
  subtype?: unknown;
  id?: unknown;
  inkLists?: unknown;
  color?: unknown;
  opacity?: unknown;
  it?: unknown;
  borderStyle?: {
    width?: unknown;
    rawWidth?: unknown;
  };
}

interface NativeInkMergeCandidate {
  id: string;
  page: number;
  color: string;
  width: number;
  points: number[];
  penType?: InkPenType;
  nativeSource?: NativePdfInkAnnotationSource;
}

function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

function finiteNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function numericArray(value: unknown): number[] | null {
  if (!Array.isArray(value) && !ArrayBuffer.isView(value)) return null;
  const raw = Array.from(value as ArrayLike<unknown>);
  // One x/y pair is a dot, a stroke like any other: requiring two points
  
  if (
    raw.length < 2 ||
    raw.length > MAX_NATIVE_PDF_INK_POINT_VALUES ||
    raw.length % 2 !== 0
  ) {
    return null;
  }
  const numbers: number[] = [];
  for (const item of raw) {
    const number = finiteNumber(item);
    if (number === null) return null;
    numbers.push(number);
  }
  return numbers;
}








function colorComponents(values: readonly unknown[]): number[] | null {
  const numbers = values.map(finiteNumber);
  if (numbers.some((number) => number === null || number < 0 || number > 255)) {
    return null;
  }
  const finite = numbers as number[];
  if (finite.every(Number.isInteger)) return finite;
  if (finite.some((number) => number > 1)) {
    return finite.map((number) => Math.round(number));
  }
  return finite.map((number) => Math.round(number * 255));
}

function nativeInkColor(value: unknown): string {
  if (!Array.isArray(value) && !ArrayBuffer.isView(value)) return "#000000";
  const components = Array.from(value as ArrayLike<unknown>);
  if (components.length === 1) {
    const gray = colorComponents(components)?.[0] ?? null;
    if (gray === null) return "#000000";
    const token = gray.toString(16).padStart(2, "0");
    return `#${token}${token}${token}`.toUpperCase();
  }
  if (components.length < 3) return "#000000";
  const rgb = colorComponents(components.slice(0, 3));
  if (!rgb) return "#000000";
  return `#${rgb
    .map((component) => component.toString(16).padStart(2, "0"))
    .join("")}`.toUpperCase();
}

function nativeInkPenType(annotation: RawNativePdfInkAnnotation): InkPenType {
  const opacity = finiteNumber(annotation.opacity);
  return annotation.it === "InkHighlight" ||
    (opacity !== null && opacity > 0 && opacity < 0.75)
    ? "highlighter"
    : "pen";
}

function nativeInkWidth(
  annotation: RawNativePdfInkAnnotation,
  pageView: readonly number[],
): number {
  const pageWidth = Math.abs(pageView[2] - pageView[0]);
  const pageHeight = Math.abs(pageView[3] - pageView[1]);
  const sourceWidth =
    finiteNumber(annotation.borderStyle?.width) ??
    finiteNumber(annotation.borderStyle?.rawWidth) ??
    1;
  const normalized =
    sourceWidth *
    (Math.max(pageWidth, pageHeight) /
      Math.max(1, Math.min(pageWidth, pageHeight)));
  return Math.min(100, Math.max(0.1, normalized));
}

function normalizePath(
  value: unknown,
  viewport: NativePdfInkViewport,
): number[] | null {
  const source = numericArray(value);
  if (!source) return null;
  const points: number[] = [];
  for (let index = 0; index < source.length; index += 2) {
    const converted = viewport.convertToViewportPoint(
      source[index],
      source[index + 1],
    );
    const x = finiteNumber(converted?.[0]);
    const y = finiteNumber(converted?.[1]);
    if (x === null || y === null) return null;
    const xFraction = x / viewport.width;
    const yFraction = y / viewport.height;
    if (
      xFraction < -COORDINATE_EPSILON ||
      xFraction > 1 + COORDINATE_EPSILON ||
      yFraction < -COORDINATE_EPSILON ||
      yFraction > 1 + COORDINATE_EPSILON
    ) {
      return null;
    }
    points.push(
      Math.min(1, Math.max(0, xFraction)),
      Math.min(1, Math.max(0, yFraction)),
    );
  }
  return points;
}

export function sanitizeNativePdfInkAnnotationSource(
  value: unknown,
): NativePdfInkAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const pathIndex = Number(source.pathIndex ?? source.path_index);
  const pathCount = Number(source.pathCount ?? source.path_count);
  const originalPointsValue = source.originalPoints ?? source.original_points;
  const originalPoints =
    originalPointsValue === undefined ? undefined : numericArray(originalPointsValue);
  if (
    source.kind !== NATIVE_PDF_INK_SOURCE_KIND ||
    !isPdfObjectRef(source.objectRef ?? source.object_ref) ||
    !Number.isInteger(pathIndex) ||
    pathIndex < 0 ||
    !Number.isInteger(pathCount) ||
    pathCount < 1 ||
    pathCount > MAX_NATIVE_PDF_INK_PATHS_PER_ANNOTATION ||
    pathIndex >= pathCount ||
    (originalPointsValue !== undefined &&
      (!originalPoints ||
        originalPoints.some((coordinate) => coordinate < 0 || coordinate > 1)))
  ) {
    return undefined;
  }
  const originalColor = source.originalColor ?? source.original_color;
  const originalWidth = Number(source.originalWidth ?? source.original_width);
  return {
    kind: NATIVE_PDF_INK_SOURCE_KIND,
    objectRef: String(source.objectRef ?? source.object_ref),
    pathIndex,
    pathCount,
    ...(originalPoints ? { originalPoints } : {}),
    ...(typeof originalColor === "string" && /^#[0-9A-Fa-f]{6}$/.test(originalColor)
      ? { originalColor: originalColor.toUpperCase() }
      : {}),
    ...(Number.isFinite(originalWidth) && originalWidth > 0 && originalWidth <= 100
      ? { originalWidth }
      : {}),
  };
}

export function isNativePdfInkStroke(
  stroke: Pick<NativeInkMergeCandidate, "nativeSource">,
): stroke is NativeInkMergeCandidate & {
  nativeSource: NativePdfInkAnnotationSource;
} {
  return stroke.nativeSource?.kind === NATIVE_PDF_INK_SOURCE_KIND;
}

/**
 * Normalize every editable path from pdf.js' annotation payload.
 *
 * Unsupported objects remain painted by pdf.js and are returned as explicit
 * losses; callers must never interpret an empty stroke list as lossless.
 */
export function normalizeNativePdfInkAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  pageView: readonly number[] | undefined,
  viewport: NativePdfInkViewport,
): NativePdfInkImportResult {
  const strokes: NativePdfInkStroke[] = [];
  const losses: NativePdfInkImportLoss[] = [];
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1 ||
    !Array.isArray(pageView) ||
    pageView.length < 4 ||
    pageView.slice(0, 4).some((value) => !Number.isFinite(value)) ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return {
      strokes,
      losses: [{ reason: "invalid_page_geometry" }],
    };
  }

  let inkAnnotationCount = 0;
  let totalPathCount = 0;
  let totalPointValues = 0;
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfInkAnnotation;
    if (annotation.subtype !== "Ink") continue;
    inkAnnotationCount += 1;
    const objectRef = isPdfObjectRef(annotation.id) ? annotation.id : undefined;
    if (inkAnnotationCount > MAX_NATIVE_PDF_INK_ANNOTATIONS_PER_PAGE) {
      losses.push({ objectRef, reason: "annotation_limit" });
      continue;
    }
    if (!objectRef) {
      losses.push({ reason: "invalid_object_ref" });
      continue;
    }
    if (!Array.isArray(annotation.inkLists)) {
      losses.push({ objectRef, reason: "invalid_path_geometry" });
      continue;
    }
    const pathCount = annotation.inkLists.length;
    if (
      pathCount < 1 ||
      pathCount > MAX_NATIVE_PDF_INK_PATHS_PER_ANNOTATION ||
      totalPathCount + pathCount > MAX_NATIVE_PDF_INK_PATHS_PER_PAGE
    ) {
      losses.push({ objectRef, reason: "path_limit" });
      continue;
    }
    const normalizedPaths: number[][] = [];
    let invalidReason: NativePdfInkImportLossReason | null = null;
    for (const path of annotation.inkLists) {
      const pathValues =
        Array.isArray(path) || ArrayBuffer.isView(path)
          ? Array.from(path as ArrayLike<unknown>).length
          : 0;
      if (
        pathValues > MAX_NATIVE_PDF_INK_POINT_VALUES ||
        totalPointValues +
          normalizedPaths.reduce((sum, item) => sum + item.length, 0) +
          pathValues >
          MAX_NATIVE_PDF_INK_POINT_VALUES_PER_PAGE
      ) {
        invalidReason = "point_limit";
        break;
      }
      const points = normalizePath(path, viewport);
      if (!points) {
        invalidReason = "invalid_path_geometry";
        break;
      }
      normalizedPaths.push(points);
    }
    if (invalidReason) {
      losses.push({ objectRef, reason: invalidReason });
      continue;
    }

    const color = nativeInkColor(annotation.color);
    const width = nativeInkWidth(annotation, pageView);
    const penType = nativeInkPenType(annotation);
    normalizedPaths.forEach((points, pathIndex) => {
      strokes.push({
        id: `pdf-native-ink:${pageNumber}:${objectRef}:${pathIndex}`,
        page: pageNumber - 1,
        color,
        width,
        points,
        penType,
        nativeSource: {
          kind: NATIVE_PDF_INK_SOURCE_KIND,
          objectRef,
          pathIndex,
          pathCount,
          originalPoints: [...points],
          originalColor: color.toUpperCase(),
          originalWidth: width,
        },
      });
    });
    totalPathCount += normalizedPaths.length;
    totalPointValues += normalizedPaths.reduce(
      (sum, points) => sum + points.length,
      0,
    );
  }
  return { strokes, losses };
}

function samePoints(
  left: readonly number[],
  right: readonly number[],
): boolean {
  return (
    left.length === right.length &&
    left.every(
      (coordinate, index) =>
        Math.abs(coordinate - (right[index] ?? Number.NaN)) <= 0.000_01,
    )
  );
}

function sameAppearance(
  left: NativeInkMergeCandidate,
  right: NativeInkMergeCandidate,
): boolean {
  return (
    left.page === right.page &&
    left.color.toUpperCase() === right.color.toUpperCase() &&
    normalizeInkPenType(left.penType) === normalizeInkPenType(right.penType) &&
    Math.abs(left.width - right.width) <= 0.01 &&
    samePoints(left.points, right.points)
  );
}











export function mergeNativePdfInkAnnotations<T extends NativeInkMergeCandidate>(
  current: readonly T[],
  incoming: readonly NativePdfInkStroke[],
  pageNumber?: number,
): T[] {
  const pageIndex =
    pageNumber !== undefined && Number.isInteger(pageNumber) ? pageNumber - 1 : null;
  if (incoming.length === 0 && pageIndex === null) return current as T[];
  const next = [...current];
  let changed = false;
  const claimedIndexes = new Set<number>();

  for (const discovered of incoming) {
    if (discovered.managedId !== undefined) {
      const managedIndex = next.findIndex(
        (stroke, index) =>
          !claimedIndexes.has(index) &&
          !stroke.nativeSource &&
          stroke.id === discovered.managedId,
      );
      if (managedIndex >= 0) {
        // The patched preview paints this managed stroke; its own record is
        // the one to show, move and save. A source twin left from before is
        // unclaimed and goes with the page snapshot below.
        claimedIndexes.add(managedIndex);
        continue;
      }
    }
    let matchIndex = next.findIndex(
      (stroke, index) =>
        !claimedIndexes.has(index) &&
        stroke.nativeSource?.kind === NATIVE_PDF_INK_SOURCE_KIND &&
        stroke.nativeSource.objectRef === discovered.nativeSource.objectRef &&
        stroke.nativeSource.pathIndex === discovered.nativeSource.pathIndex,
    );
    if (matchIndex < 0) {
      matchIndex = next.findIndex(
        (stroke, index) =>
          !claimedIndexes.has(index) && sameAppearance(stroke, discovered),
      );
    }
    if (matchIndex >= 0) {
      claimedIndexes.add(matchIndex);
      const existing = next[matchIndex];
      // A geometry match against an MRNote-managed stroke means the patched
      // preview is painting the same object. Do not add a source-backed twin.
      if (!existing.nativeSource) continue;
      if (
        existing.page !== discovered.page ||
        existing.color !== discovered.color ||
        Math.abs(existing.width - discovered.width) > 0.01 ||
        normalizeInkPenType(existing.penType) !== discovered.penType ||
        !samePoints(existing.points, discovered.points) ||
        !samePoints(
          existing.nativeSource.originalPoints ?? [],
          discovered.points,
        ) ||
        // A restyle the reopened revision now draws is its new base.
        (existing.nativeSource.originalColor ?? discovered.color).toUpperCase() !==
          discovered.color.toUpperCase() ||
        Math.abs((existing.nativeSource.originalWidth ?? discovered.width) - discovered.width) > 0.01
      ) {
        next[matchIndex] = {
          ...existing,
          page: discovered.page,
          color: discovered.color,
          width: discovered.width,
          points: discovered.points,
          penType: discovered.penType,
          // Preserve the stable imported identity while rebasing geometry to
          // the exact PDF revision that just reopened.
          nativeSource: {
            ...existing.nativeSource,
            pathIndex: discovered.nativeSource.pathIndex,
            pathCount: discovered.nativeSource.pathCount,
            originalPoints: [...discovered.points],
            originalColor: discovered.color.toUpperCase(),
            originalWidth: discovered.width,
          },
        };
        changed = true;
      }
      continue;
    }
    const { managedId: _managedId, ...sourceStroke } = discovered;
    void _managedId;
    next.push(sourceStroke as T);
    claimedIndexes.add(next.length - 1);
    changed = true;
  }
  if (pageIndex !== null) {
    const kept = next.filter(
      (stroke, index) =>
        claimedIndexes.has(index) ||
        stroke.page !== pageIndex ||
        stroke.nativeSource?.kind !== NATIVE_PDF_INK_SOURCE_KIND,
    );
    if (kept.length !== next.length) return kept;
  }
  return changed ? next : (current as T[]);
}
