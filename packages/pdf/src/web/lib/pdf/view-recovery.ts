/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-document PDF view recovery.
 *
 * This is local viewer state, never PDF/page content. Snapshots are keyed only
 * by a stable data-item id, versioned, size-bounded, and treated as untrusted.
 * Corrupt, stale-shaped, oversized, or unavailable storage is a silent no-op.
 */

import {
  PDF_DEFAULT_READ_MODE,
  isPdfReadMode,
  type PdfReadMode,
} from "./invert-filter";
import {
  PDF_DEFAULT_VIEW_MODE,
  PDF_VIEW_MODES,
  clampScale,
  type PdfViewMode,
} from "./view-modes";

export interface PdfViewRecoverySnapshot {
  version: 1;
  /** Current original/source page, 1-based. */
  page: number;
  /** Current slot in the final visible sequence, including inserted pages. */
  visibleOrdinal: number;
  /** Absolute container position; reused only for the exact same structure. */
  scrollTop: number;
  scale: number;
  viewMode: PdfViewMode;
  readMode: PdfReadMode;
  /** Source-page count, excluding locally inserted pages. */
  pageCount: number;
  /** Exact bounded structure identity, or null when too large to retain. */
  structureSignature: string | null;
}

export interface ResolvedPdfViewRecovery {
  page: number;
  /** Present only when the saved and live page structure match exactly. */
  visibleOrdinal: number | null;
  /** Present only when the saved and live page structure match exactly. */
  scrollTop: number | null;
  scale: number;
  viewMode: PdfViewMode;
  readMode: PdfReadMode;
}

export interface PdfPageIntentState {
  readonly page: number;
  readonly generation: number;
}

export interface PdfViewRecoverySourcePageLayout {
  readonly page: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
}

export interface PdfViewRecoveryInsertedPageLayout {
  readonly id: string;
  readonly afterPage: number;
  readonly width: number;
  readonly height: number;
  readonly copyOfPage?: number;
}

const STORAGE_PREFIX = "mrnote.pdf.view-recovery.v1:";
const MAX_SCOPE_LENGTH = 512;
const MAX_RAW_LENGTH = 16_384;
const MAX_PAGE_COUNT = 10_000;
const MAX_SCROLL_TOP = 1_000_000_000;
const MAX_STRUCTURE_SIGNATURE_LENGTH = 8_192;
const MAX_LAYOUT_ID_LENGTH = 512;
const MAX_LAYOUT_DIMENSION = 10_000_000;

const VIEW_MODE_SET = new Set<string>(PDF_VIEW_MODES);

/**
 * Advance host page intent only for a real page transition. In particular, an
 * eager viewer echo of the neutral page 1 must not manufacture authority that
 * can suppress the first local reading-position recovery.
 */
export function advancePdfPageIntent(
  current: PdfPageIntentState,
  requestedPage: number,
): PdfPageIntentState {
  if (!Number.isFinite(requestedPage)) return current;
  const page = Math.max(1, Math.trunc(requestedPage));
  if (page === current.page) return current;
  const generation =
    Number.isSafeInteger(current.generation) &&
    current.generation > 0 &&
    current.generation < Number.MAX_SAFE_INTEGER
      ? current.generation + 1
      : 1;
  return Object.freeze({ page, generation });
}

function positiveInteger(value: unknown, max: number): number | null {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < 1 ||
    value > max
  ) {
    return null;
  }
  return value;
}

function normalizedScrollTop(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    return null;
  }
  return Math.min(MAX_SCROLL_TOP, value);
}

function normalizedStructureSignature(
  value: unknown,
): string | null | undefined {
  if (value === null) return null;
  if (typeof value !== "string") return undefined;
  if (value.length > MAX_STRUCTURE_SIGNATURE_LENGTH) return null;
  return value;
}

/** A stable localStorage key, or null when no safe stable scope was supplied. */
export function pdfViewRecoveryStorageKey(scope: unknown): string | null {
  if (typeof scope !== "string") return null;
  const trimmed = scope.trim();
  if (!trimmed || trimmed.length > MAX_SCOPE_LENGTH) return null;
  return `${STORAGE_PREFIX}${encodeURIComponent(trimmed)}`;
}

/**
 * Keep the exact structure signature when bounded. A pathological signature is
 * represented as null, which deliberately disables absolute-scroll recovery.
 */
export function boundedPdfStructureSignature(value: string): string | null {
  return value.length <= MAX_STRUCTURE_SIGNATURE_LENGTH ? value : null;
}

/**
 * Exact, bounded identity for the geometry that gives an absolute scrollTop
 * meaning. It deliberately returns null instead of hashing/truncating when the
 * full source generation, edit structure, page geometry, or effective rotation
 * cannot be represented exactly within the bound.
 */
export function pdfViewRecoveryStructureSignature(input: {
  sourceGeneration: unknown;
  structuralSignature: unknown;
  sourcePages: readonly PdfViewRecoverySourcePageLayout[];
  insertedPages: readonly PdfViewRecoveryInsertedPageLayout[];
}): string | null {
  if (
    typeof input.sourceGeneration !== "string" ||
    input.sourceGeneration.length === 0 ||
    input.sourceGeneration.length > MAX_STRUCTURE_SIGNATURE_LENGTH ||
    typeof input.structuralSignature !== "string" ||
    input.structuralSignature.length > MAX_STRUCTURE_SIGNATURE_LENGTH ||
    input.sourcePages.length < 1 ||
    input.sourcePages.length > MAX_PAGE_COUNT ||
    input.insertedPages.length > MAX_PAGE_COUNT
  ) {
    return null;
  }

  let signature = JSON.stringify([
    "pdf-view-layout-v1",
    input.sourceGeneration,
    input.structuralSignature,
  ]);
  if (boundedPdfStructureSignature(signature) === null) return null;

  const append = (value: readonly unknown[]): boolean => {
    signature += `|${JSON.stringify(value)}`;
    return boundedPdfStructureSignature(signature) !== null;
  };
  for (const page of input.sourcePages) {
    if (
      !Number.isSafeInteger(page.page) ||
      page.page < 1 ||
      page.page > MAX_PAGE_COUNT ||
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      page.width > MAX_LAYOUT_DIMENSION ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      page.height > MAX_LAYOUT_DIMENSION ||
      ![0, 90, 180, 270].includes(page.rotation) ||
      !append(["source", page.page, page.width, page.height, page.rotation])
    ) {
      return null;
    }
  }
  for (const page of input.insertedPages) {
    const copyOfPage = page.copyOfPage ?? null;
    if (
      typeof page.id !== "string" ||
      page.id.length === 0 ||
      page.id.length > MAX_LAYOUT_ID_LENGTH ||
      !Number.isSafeInteger(page.afterPage) ||
      page.afterPage < 0 ||
      page.afterPage > MAX_PAGE_COUNT ||
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      page.width > MAX_LAYOUT_DIMENSION ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      page.height > MAX_LAYOUT_DIMENSION ||
      (copyOfPage !== null &&
        (!Number.isSafeInteger(copyOfPage) ||
          copyOfPage < 1 ||
          copyOfPage > MAX_PAGE_COUNT)) ||
      !append([
        "inserted",
        page.id,
        page.afterPage,
        page.width,
        page.height,
        copyOfPage,
      ])
    ) {
      return null;
    }
  }
  return signature;
}

/** Validate a stored or newly-built snapshot without touching storage. */
export function normalizePdfViewRecoverySnapshot(
  candidate: unknown,
): PdfViewRecoverySnapshot | null {
  if (!candidate || typeof candidate !== "object") return null;
  const record = candidate as Record<string, unknown>;
  if (record.version !== 1) return null;

  const page = positiveInteger(record.page, MAX_PAGE_COUNT);
  const visibleOrdinal = positiveInteger(
    record.visibleOrdinal,
    MAX_PAGE_COUNT * 2,
  );
  const pageCount = positiveInteger(record.pageCount, MAX_PAGE_COUNT);
  const scrollTop = normalizedScrollTop(record.scrollTop);
  const structureSignature = normalizedStructureSignature(
    record.structureSignature,
  );
  if (
    page === null ||
    visibleOrdinal === null ||
    pageCount === null ||
    scrollTop === null ||
    structureSignature === undefined ||
    typeof record.viewMode !== "string" ||
    !VIEW_MODE_SET.has(record.viewMode) ||
    !isPdfReadMode(record.readMode) ||
    typeof record.scale !== "number" ||
    !Number.isFinite(record.scale) ||
    record.scale <= 0
  ) {
    return null;
  }

  return {
    version: 1,
    page,
    visibleOrdinal,
    scrollTop,
    scale: clampScale(record.scale),
    viewMode: record.viewMode as PdfViewMode,
    readMode: record.readMode,
    pageCount,
    structureSignature,
  };
}

/** Load one current, bounded snapshot. Storage failures are a safe miss. */
export function loadPdfViewRecovery(
  scope: unknown,
): PdfViewRecoverySnapshot | null {
  const key = pdfViewRecoveryStorageKey(scope);
  if (!key) return null;
  try {
    const raw = globalThis.localStorage?.getItem(key);
    if (!raw || raw.length > MAX_RAW_LENGTH) return null;
    return normalizePdfViewRecoverySnapshot(JSON.parse(raw));
  } catch {
    return null;
  }
}

/** Persist one bounded snapshot, returning the value actually written. */
export function savePdfViewRecovery(
  scope: unknown,
  candidate: PdfViewRecoverySnapshot,
): PdfViewRecoverySnapshot | null {
  const key = pdfViewRecoveryStorageKey(scope);
  const snapshot = normalizePdfViewRecoverySnapshot(candidate);
  if (!key || !snapshot) return null;
  try {
    const raw = JSON.stringify(snapshot);
    if (raw.length > MAX_RAW_LENGTH) return null;
    globalThis.localStorage?.setItem(key, raw);
    return snapshot;
  } catch {
    return null;
  }
}

/**
 * Reconcile a valid snapshot with the currently loaded document. Page and
 * scale remain bounded. Absolute scroll/visible-slot recovery is allowed only
 * when source page count and the exact bounded structure signature both match.
 */
export function resolvePdfViewRecovery(
  snapshot: PdfViewRecoverySnapshot,
  current: {
    pageCount: number;
    visibleCount: number;
    structureSignature: string | null;
  },
): ResolvedPdfViewRecovery {
  const pageCount = Math.max(
    1,
    Math.min(MAX_PAGE_COUNT, Math.trunc(current.pageCount)),
  );
  const visibleCount = Math.max(
    1,
    Math.min(MAX_PAGE_COUNT * 2, Math.trunc(current.visibleCount)),
  );
  const sameStructure =
    snapshot.pageCount === pageCount &&
    snapshot.structureSignature !== null &&
    current.structureSignature !== null &&
    snapshot.structureSignature === current.structureSignature;

  return {
    page: Math.min(snapshot.page, pageCount),
    visibleOrdinal: sameStructure
      ? Math.min(snapshot.visibleOrdinal, visibleCount)
      : null,
    scrollTop: sameStructure ? snapshot.scrollTop : null,
    scale: clampScale(snapshot.scale),
    viewMode: VIEW_MODE_SET.has(snapshot.viewMode)
      ? snapshot.viewMode
      : PDF_DEFAULT_VIEW_MODE,
    readMode: isPdfReadMode(snapshot.readMode)
      ? snapshot.readMode
      : PDF_DEFAULT_READ_MODE,
  };
}
