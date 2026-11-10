/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Page-thumbnail rendering math for the PDF viewer sidebar.
 *
 * Thumbnails reuse the same pdf.js document as the main view but render at a
 * fixed CSS target width (~120px) so a tall sidebar can hold many pages. The
 * helpers here are pure so they can be unit-tested without pdf.js/jsdom.
 */

/** Default CSS width (in px) a single thumbnail canvas should occupy. */
export const THUMBNAIL_TARGET_WIDTH = 120;

/**
 * Documents with at most this many pages are small enough to render every
 * thumbnail eagerly when the sidebar opens; larger docs render lazily as
 * thumbnails scroll into view.
 */
export const THUMBNAIL_EAGER_PAGE_LIMIT = 40;

/**
 * Scale factor that maps a full-size page width (pdf.js viewport at scale 1.0)
 * to the desired thumbnail CSS width. Clamped to a sane range so a degenerate
 * page width (0 / NaN / negative) can never produce an invalid render scale.
 *
 * @param pageWidth   Full-size page width in CSS px (pdf.js scale 1.0).
 * @param targetWidth Desired thumbnail width in CSS px.
 */
export function thumbnailScale(
  pageWidth: number,
  targetWidth: number = THUMBNAIL_TARGET_WIDTH,
): number {
  const width = Number(pageWidth);
  const target = Number(targetWidth);
  if (!Number.isFinite(width) || width <= 0) return 0.1;
  if (!Number.isFinite(target) || target <= 0) return 0.1;
  const scale = target / width;
  // Keep within bounds: never below 0.02 (tiny, avoids 0-area canvas) and never
  // above 1.0 (a thumbnail should never be larger than the source page).
  return Math.min(1, Math.max(0.02, scale));
}

/**
 * Whether a document is small enough to render every thumbnail eagerly when the
 * sidebar first opens (vs. lazy on-scroll rendering for large docs).
 */
export function shouldEagerRenderThumbnails(
  pageCount: number,
  limit: number = THUMBNAIL_EAGER_PAGE_LIMIT,
): boolean {
  const count = Number(pageCount);
  if (!Number.isFinite(count) || count <= 0) return false;
  return count <= limit;
}
