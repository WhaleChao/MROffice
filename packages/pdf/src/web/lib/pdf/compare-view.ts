/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded, read-only PDF side-by-side comparison helpers.
 *
 * The comparison document is deliberately outside the editable document
 * model. These helpers contain only local-file admission, proportional scroll
 * projection, and terminal resource cleanup; they never expose save/export
 * primitives or mutate the source PDF.
 */

import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
export const PDF_COMPARE_MAX_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_COMPARE_MAX_PAGES = 500;

export type PdfCompareFileRejection = "empty" | "not_pdf" | "too_large";

export interface PdfCompareFileLike {
  readonly name: string;
  readonly size: number;
  readonly type?: string;
}

/** Cheap admission before pdf.js performs the authoritative parse. */
export function rejectPdfCompareFile(
  file: PdfCompareFileLike,
): PdfCompareFileRejection | null {
  if (!Number.isFinite(file.size) || file.size <= 0) return "empty";
  if (file.size > PDF_COMPARE_MAX_BYTES) return "too_large";

  const type = String(file.type ?? "")
    .trim()
    .toLowerCase();
  const name = String(file.name ?? "")
    .trim()
    .toLowerCase();
  if (type && type !== "application/pdf" && !name.endsWith(".pdf")) {
    return "not_pdf";
  }
  if (!type && !name.endsWith(".pdf")) return "not_pdf";
  return null;
}

interface ScrollGeometry {
  readonly scrollTop: number;
  readonly scrollHeight: number;
  readonly clientHeight: number;
}

/** Position in the scrollable range, clamped to [0, 1]. */
export function pdfCompareScrollProgress(source: ScrollGeometry): number {
  const range = Math.max(0, source.scrollHeight - source.clientHeight);
  if (!(range > 0)) return 0;
  return Math.min(1, Math.max(0, source.scrollTop / range));
}

/** Project a document-relative position into the other column. */
export function pdfCompareProjectedScrollTop(
  source: ScrollGeometry,
  target: ScrollGeometry,
): number {
  const range = Math.max(0, target.scrollHeight - target.clientHeight);
  return pdfCompareScrollProgress(source) * range;
}

export function synchronizePdfCompareScroll(
  source: ScrollGeometry,
  target: ScrollGeometry & { scrollTop: number },
): void {
  target.scrollTop = pdfCompareProjectedScrollTop(source, target);
}

export interface PdfCompareDestroyable {
  destroy(): void | Promise<void>;
}

export interface PdfCompareResources {
  readonly loadingTask?: PdfCompareDestroyable | null;
  readonly document?: PdfCompareDestroyable | null;
  readonly objectUrl?: string | null;
  readonly revokeObjectUrl?: (url: string) => void;
}

/**
 * Terminal cleanup is best-effort and complete: one broken pdf.js destructor
 * must not retain the other resource or the browser object URL.
 */
export async function destroyPdfCompareResources({
  loadingTask,
  document,
  objectUrl,
  revokeObjectUrl = (url) => URL.revokeObjectURL(url),
}: PdfCompareResources): Promise<void> {
  const owner = document ?? loadingTask;
  await Promise.allSettled([Promise.resolve().then(() => owner?.destroy())]);
  if (objectUrl) revokeObjectUrl(objectUrl);
}
