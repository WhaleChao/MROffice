/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Editable third-party `/Stamp` annotations discovered through pdf.js.
 *
 * The source PDF canvas remains the visual authority. These records only add
 * a transparent selection target plus the exact indirect-object identity
 * required to transform or delete the original annotation without rebuilding
 * (and therefore potentially degrading) its `/AP` appearance stream.
 */

export const NATIVE_PDF_STAMP_SOURCE_KIND = "pdf_stamp_annotation";
export const MAX_NATIVE_PDF_STAMP_ANNOTATIONS_PER_PAGE = 512;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/;
const COORDINATE_EPSILON = 0.000_001;

export interface NativePdfStampAnnotationSource {
  kind: typeof NATIVE_PDF_STAMP_SOURCE_KIND;
  objectRef: string;
  subtype: "Stamp";
  /** Optional source `/Name`, retained only for diagnostics and joins. */
  name?: string;
}

export interface NativePdfStampAnnotation {
  id: string;
  /** One-based original PDF page number. */
  page: number;
  /** Top-left page-relative percentage box. */
  x: number;
  y: number;
  width: number;
  height: number;
  nativeSource: NativePdfStampAnnotationSource;
  /** Absent for untouched source objects. */
  sourceAction?: "transform";
}

export interface NativePdfStampTombstone {
  id: string;
  page: number;
  deleted: true;
  nativeSource: NativePdfStampAnnotationSource;
}

export type NativePdfStampRecord =
  (NativePdfStampAnnotation & { deleted?: false }) | NativePdfStampTombstone;

export type NativePdfStampImportLossReason =
  | "invalid_object_ref"
  | "duplicate_object_ref"
  | "invalid_page_geometry"
  | "invalid_stamp_geometry"
  | "annotation_limit";

export interface NativePdfStampImportLoss {
  objectRef?: string;
  reason: NativePdfStampImportLossReason;
}

export interface NativePdfStampImportResult {
  stamps: NativePdfStampAnnotation[];
  losses: NativePdfStampImportLoss[];
}

export interface NativePdfStampViewport {
  width: number;
  height: number;
  convertToViewportRectangle: (
    rect: readonly number[],
  ) => readonly [number, number, number, number];
}

interface RawNativePdfStampAnnotation {
  subtype?: unknown;
  id?: unknown;
  rect?: unknown;
  name?: unknown;
}

function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

function finiteNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function numericRect(value: unknown): [number, number, number, number] | null {
  if (!Array.isArray(value) && !ArrayBuffer.isView(value)) return null;
  const values = Array.from(value as ArrayLike<unknown>).map(finiteNumber);
  if (values.length !== 4 || values.some((item) => item === null)) return null;
  return values as [number, number, number, number];
}

function sourceName(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed && trimmed.length <= 80 ? trimmed : undefined;
}

function normalizedViewportBox(
  sourceRect: [number, number, number, number],
  viewport: NativePdfStampViewport,
): Pick<NativePdfStampAnnotation, "x" | "y" | "width" | "height"> | null {
  let converted: readonly number[];
  try {
    converted = viewport.convertToViewportRectangle(sourceRect);
  } catch {
    return null;
  }
  const rect = numericRect(converted);
  if (!rect) return null;
  const left = Math.min(rect[0], rect[2]) / viewport.width;
  const right = Math.max(rect[0], rect[2]) / viewport.width;
  const top = Math.min(rect[1], rect[3]) / viewport.height;
  const bottom = Math.max(rect[1], rect[3]) / viewport.height;
  if (right - left <= COORDINATE_EPSILON || bottom - top <= COORDINATE_EPSILON) {
    return null;
  }
  // A stamp reaching past the paper is still the author's whole stamp: its
  // box keeps its true, signed extent (the part on the paper is what takes
  
  // stamp with nothing on the paper is out of reach.
  const visibleWidth = Math.min(1, right) - Math.max(0, left);
  const visibleHeight = Math.min(1, bottom) - Math.max(0, top);
  if (visibleWidth <= COORDINATE_EPSILON || visibleHeight <= COORDINATE_EPSILON) {
    return null;
  }
  return {
    x: left * 100,
    y: top * 100,
    width: (right - left) * 100,
    height: (bottom - top) * 100,
  };
}

export function sanitizeNativePdfStampAnnotationSource(
  value: unknown,
): NativePdfStampAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  if (
    source.kind !== NATIVE_PDF_STAMP_SOURCE_KIND ||
    !isPdfObjectRef(objectRef) ||
    source.subtype !== "Stamp"
  ) {
    return undefined;
  }
  const name = sourceName(source.name);
  return {
    kind: NATIVE_PDF_STAMP_SOURCE_KIND,
    objectRef,
    subtype: "Stamp",
    ...(name ? { name } : {}),
  };
}

export function normalizeNativePdfStampAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  viewport: NativePdfStampViewport,
): NativePdfStampImportResult {
  const stamps: NativePdfStampAnnotation[] = [];
  const losses: NativePdfStampImportLoss[] = [];
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1 ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return { stamps, losses: [{ reason: "invalid_page_geometry" }] };
  }

  const objectRefCounts = new Map<string, number>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfStampAnnotation;
    if (annotation.subtype === "Stamp" && isPdfObjectRef(annotation.id)) {
      objectRefCounts.set(
        annotation.id,
        (objectRefCounts.get(annotation.id) ?? 0) + 1,
      );
    }
  }

  const reportedDuplicateRefs = new Set<string>();
  let stampCount = 0;
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfStampAnnotation;
    if (annotation.subtype !== "Stamp") continue;
    stampCount += 1;
    const objectRef = isPdfObjectRef(annotation.id) ? annotation.id : undefined;
    if (stampCount > MAX_NATIVE_PDF_STAMP_ANNOTATIONS_PER_PAGE) {
      losses.push({ objectRef, reason: "annotation_limit" });
      continue;
    }
    if (!objectRef) {
      losses.push({ reason: "invalid_object_ref" });
      continue;
    }
    if ((objectRefCounts.get(objectRef) ?? 0) > 1) {
      if (!reportedDuplicateRefs.has(objectRef)) {
        losses.push({ objectRef, reason: "duplicate_object_ref" });
        reportedDuplicateRefs.add(objectRef);
      }
      continue;
    }
    const sourceRect = numericRect(annotation.rect);
    const geometry = sourceRect
      ? normalizedViewportBox(sourceRect, viewport)
      : null;
    if (!geometry) {
      losses.push({ objectRef, reason: "invalid_stamp_geometry" });
      continue;
    }
    const name = sourceName(annotation.name);
    stamps.push({
      id: `pdf-native-stamp:${pageNumber}:${objectRef}`,
      page: pageNumber,
      ...geometry,
      nativeSource: {
        kind: NATIVE_PDF_STAMP_SOURCE_KIND,
        objectRef,
        subtype: "Stamp",
        ...(name ? { name } : {}),
      },
    });
  }
  return { stamps, losses };
}

/**
 * Merge one freshly rendered page without overwriting a pending source
 * transform or deletion tombstone that still targets the immutable input PDF.
 */
export function mergeNativePdfStampAnnotations(
  current: readonly NativePdfStampRecord[],
  incoming: readonly NativePdfStampAnnotation[],
  pageNumber: number,
): NativePdfStampRecord[] {
  const currentPage = current.filter((record) => record.page === pageNumber);
  const protectedByRef = new Map(
    currentPage
      .filter((record) => record.deleted || record.sourceAction === "transform")
      .map((record) => [record.nativeSource.objectRef, record]),
  );
  const nextPage: NativePdfStampRecord[] = [...protectedByRef.values()];
  for (const discovered of incoming) {
    if (protectedByRef.has(discovered.nativeSource.objectRef)) continue;
    const matchesProtectedTransform = [...protectedByRef.values()].some(
      (record) =>
        !record.deleted &&
        record.sourceAction === "transform" &&
        Math.abs(record.x - discovered.x) <= 0.0001 &&
        Math.abs(record.y - discovered.y) <= 0.0001 &&
        Math.abs(record.width - discovered.width) <= 0.0001 &&
        Math.abs(record.height - discovered.height) <= 0.0001,
    );
    if (matchesProtectedTransform) continue;
    nextPage.push(discovered);
  }

  const unchanged =
    currentPage.length === nextPage.length &&
    nextPage.every((record, index) => record === currentPage[index]);
  if (unchanged) return current as NativePdfStampRecord[];

  const merged: NativePdfStampRecord[] = [];
  let insertedPage = false;
  for (const record of current) {
    if (record.page === pageNumber) {
      if (!insertedPage) {
        merged.push(...nextPage);
        insertedPage = true;
      }
      continue;
    }
    merged.push(record);
  }
  if (!insertedPage) merged.push(...nextPage);
  return merged;
}
