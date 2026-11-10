/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  buildLinkOverlayPayload,
  isPdfNamedPageAction,
  linkContentsLabel,
  normalizeLinkQuadPoints,
  normalizeRect,
  pdfAnnotationFlagsHideOnScreen,
  sanitizeLinkUrl,
  viewportRectToPercent,
  type PdfLinkOverlay,
} from "@mrnote/core/web/lib/pdf/links.ts";

/**
 * Editable third-party `/Link` annotations discovered through pdf.js.
 *
 * The source page remains the visual authority. These records retain the exact
 * indirect-object identity required to update or delete the original
 * annotation without replacing its rectangle, border, appearance, or metadata.
 */

export const NATIVE_PDF_LINK_SOURCE_KIND = "pdf_link_annotation";
export const MAX_NATIVE_PDF_LINK_ANNOTATIONS_PER_PAGE = 1_024;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/u;

export interface NativePdfLinkAnnotationSource {
  kind: typeof NATIVE_PDF_LINK_SOURCE_KIND;
  objectRef: string;
  subtype: "Link";
}

export interface NativePdfLinkAnnotation extends PdfLinkOverlay {
  nativeSource: NativePdfLinkAnnotationSource;
  /**
   * Persisted only after the destination or exact widget rectangle changes.
   * V1 shadows normalize either upsert back to `destination` after reopen.
   */
  sourceAction?: "destination" | "geometry";
   
  label?: string;
  


  quads?: Array<{ x: number; y: number; width: number; height: number }>;
}

export interface NativePdfLinkTombstone {
  id: string;
  /** One-based original PDF page number. */
  page: number;
  deleted: true;
  nativeSource: NativePdfLinkAnnotationSource;
}

export type NativePdfLinkRecord =
  (NativePdfLinkAnnotation & { deleted?: false }) | NativePdfLinkTombstone;

export type NativePdfLinkImportLossReason =
  | "invalid_object_ref"
  | "duplicate_object_ref"
  | "invalid_page_geometry"
  | "invalid_link_geometry"
  | "unsupported_action"
  

  | "read_only_page_action"
  | "unresolved_destination"
   
  | "hidden_annotation"
  | "annotation_limit";

export interface NativePdfLinkImportLoss {
  objectRef?: string;
  reason: NativePdfLinkImportLossReason;
}

export interface NativePdfLinkImportResult {
  links: NativePdfLinkAnnotation[];
  losses: NativePdfLinkImportLoss[];
}

export interface NativePdfLinkViewport {
  width: number;
  height: number;
  convertToViewportRectangle: (
    rect: readonly number[],
  ) => readonly [number, number, number, number];
}

interface RawNativePdfLinkAnnotation {
  subtype?: unknown;
  id?: unknown;
  rect?: unknown;
  dest?: unknown;
  url?: unknown;
  annotationFlags?: unknown;
  quadPoints?: unknown;
  contentsObj?: unknown;
}

function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

function isInternalDestination(value: unknown): value is string | unknown[] {
  return (
    (typeof value === "string" && value.length > 0) ||
    (Array.isArray(value) && value.length > 0)
  );
}

export function sanitizeNativePdfLinkAnnotationSource(
  value: unknown,
): NativePdfLinkAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  if (
    source.kind !== NATIVE_PDF_LINK_SOURCE_KIND ||
    !isPdfObjectRef(objectRef) ||
    source.subtype !== "Link"
  ) {
    return undefined;
  }
  return {
    kind: NATIVE_PDF_LINK_SOURCE_KIND,
    objectRef,
    subtype: "Link",
  };
}

export function isNativePdfLinkRecord(value: {
  nativeSource?: unknown;
}): value is NativePdfLinkRecord {
  return Boolean(sanitizeNativePdfLinkAnnotationSource(value.nativeSource));
}

/**
 * Normalize safe, standard source links into rotation-aware page percentages.
 *
 * pdf.js intentionally does not expose dangerous action dictionaries. A link
 * with neither a safe http(s) URL nor a resolvable in-document destination is
 * therefore reported as read-only instead of being approximated. The backend
 * repeats the exact action check against the source PDF before any mutation.
 */
export async function normalizeNativePdfLinkAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  viewport: NativePdfLinkViewport,
  resolveInternalDestination: (
    destination: string | unknown[],
  ) => Promise<number | null>,
): Promise<NativePdfLinkImportResult> {
  const links: NativePdfLinkAnnotation[] = [];
  const losses: NativePdfLinkImportLoss[] = [];
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1 ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return { links, losses: [{ reason: "invalid_page_geometry" }] };
  }

  const objectRefCounts = new Map<string, number>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfLinkAnnotation;
    if (annotation.subtype === "Link" && isPdfObjectRef(annotation.id)) {
      objectRefCounts.set(
        annotation.id,
        (objectRefCounts.get(annotation.id) ?? 0) + 1,
      );
    }
  }

  const reportedDuplicateRefs = new Set<string>();
  let linkCount = 0;
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfLinkAnnotation;
    if (annotation.subtype !== "Link") continue;
    linkCount += 1;
    const objectRef = isPdfObjectRef(annotation.id) ? annotation.id : undefined;
    if (linkCount > MAX_NATIVE_PDF_LINK_ANNOTATIONS_PER_PAGE) {
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

    
    // is no editable hotspot either.
    if (pdfAnnotationFlagsHideOnScreen(annotation.annotationFlags)) {
      losses.push({ objectRef, reason: "hidden_annotation" });
      continue;
    }

    const sourceRect = normalizeRect(annotation.rect);
    let percent: ReturnType<typeof viewportRectToPercent> = null;
    if (sourceRect) {
      try {
        percent = viewportRectToPercent(
          viewport.convertToViewportRectangle(sourceRect),
          viewport.width,
          viewport.height,
        );
      } catch {
        percent = null;
      }
    }
    if (!percent) {
      losses.push({ objectRef, reason: "invalid_link_geometry" });
      continue;
    }

    const destination = isInternalDestination(annotation.dest)
      ? annotation.dest
      : null;
    const url = destination == null ? sanitizeLinkUrl(annotation.url) : null;
    let goToPage: number | null = null;
    if (destination != null) {
      try {
        goToPage = await resolveInternalDestination(destination);
      } catch {
        goToPage = null;
      }
      if (!Number.isInteger(goToPage) || (goToPage ?? 0) < 1) {
        losses.push({ objectRef, reason: "unresolved_destination" });
        continue;
      }
    } else if (!url) {
      losses.push({
        objectRef,
        reason: isPdfNamedPageAction((annotation as { action?: unknown }).action)
          ? "read_only_page_action"
          : "unsupported_action",
      });
      continue;
    }

    const built = buildLinkOverlayPayload({
      page: pageNumber,
      ...percent,
      url,
      goToPage,
    });
    if (!built) {
      losses.push({ objectRef, reason: "unsupported_action" });
      continue;
    }
    const label = linkContentsLabel(annotation.contentsObj);
    const quads = sourceRect
      ? normalizeLinkQuadPoints(annotation.quadPoints, sourceRect)
          ?.map((quad) => {
            try {
              return viewportRectToPercent(
                viewport.convertToViewportRectangle(quad),
                viewport.width,
                viewport.height,
              );
            } catch {
              return null;
            }
          })
          .filter((quad): quad is NonNullable<typeof quad> => quad != null)
      : undefined;
    links.push({
      id: `pdf-native-link:${pageNumber}:${objectRef}`,
      ...built,
      ...(label ? { label } : {}),
      ...(quads && quads.length > 0 ? { quads } : {}),
      nativeSource: {
        kind: NATIVE_PDF_LINK_SOURCE_KIND,
        objectRef,
        subtype: "Link",
      },
    });
  }
  return { links, losses };
}

/**
 * Merge one freshly discovered page without overwriting a pending destination
 * edit or deletion tombstone that still targets the immutable source object.
 */
export function mergeNativePdfLinkAnnotations(
  current: readonly NativePdfLinkRecord[],
  incoming: readonly NativePdfLinkAnnotation[],
  pageNumber: number,
): NativePdfLinkRecord[] {
  const currentPage = current.filter((record) => record.page === pageNumber);
  const protectedByRef = new Map(
    currentPage
      .filter(
        (record) =>
          record.deleted ||
          record.sourceAction === "destination" ||
          record.sourceAction === "geometry",
      )
      .map((record) => [record.nativeSource.objectRef, record]),
  );
  const nextPage: NativePdfLinkRecord[] = [...protectedByRef.values()];
  for (const discovered of incoming) {
    if (protectedByRef.has(discovered.nativeSource.objectRef)) continue;
    nextPage.push(discovered);
  }

  const unchanged =
    currentPage.length === nextPage.length &&
    nextPage.every((record, index) => record === currentPage[index]);
  if (unchanged) return current as NativePdfLinkRecord[];

  const merged: NativePdfLinkRecord[] = [];
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
