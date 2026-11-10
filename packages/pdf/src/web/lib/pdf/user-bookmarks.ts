/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for USER-created PDF bookmarks (Acrobat "Add Bookmark" / WPS).
 *
 * The outline panel renders the document's SOURCE-embedded bookmark tree
 * (pdf.js `getOutline()`, see lib/pdf/outline.ts) — read-only, supplied by the
 * file. This module is the complement: bookmarks the READER creates, naming a
 * spot in the document at the current page so they can jump back later. Unlike
 * the embedded outline, these are persisted per data-item via the same
 * annotations endpoint as `pdf_note` / `pdf_ink` / `pdf_shapes` (additive — a
 * PDF with neither embedded nor user bookmarks keeps loading unchanged).
 *
 * A user bookmark is just a 1-based page number plus a short label. It holds
 * everything with no DOM / React / pdf-lib dependency so it can be unit-tested
 * in isolation and reused by the viewer:
 *
 *   - the persisted annotation shape + its sanitizer ({@link sanitizeBookmarkPayload}),
 *   - label normalization (collapse whitespace, cap length — ALWAYS rendered as
 *     plain text, so this is normalization, not HTML escaping),
 *   - {@link buildBookmarkPayload} to produce exactly what we POST,
 *   - {@link defaultBookmarkLabel} for a sensible "Page N" fallback name,
 *   - {@link sortBookmarks} (page order, then label) + {@link capBookmarks}.
 *
 * The persisted annotation type is `pdf_bookmark`.
 */

/** A persisted user-created bookmark. */
export interface UserBookmark {
  id: string;
  /** 1-based page number the bookmark points at. */
  page: number;
  /** User-authored label, already normalized; render as PLAIN TEXT only. */
  label: string;
}

/** Max length of a bookmark label after normalization. */
export const PDF_BOOKMARK_MAX_LABEL = 200;

/** Max number of user bookmarks kept per document (oldest dropped when exceeded). */
export const PDF_BOOKMARK_MAX = 500;

/** Upper bound on a 1-based page number (mirrors the other annotation sanitizers). */
const MAX_PAGE = 10000;

/**
 * Normalize a user bookmark label: collapse ALL whitespace + control chars into
 * single spaces, trim, and hard-cap the length. The label is rendered as a text
 * node downstream, so this is normalization (keeps a garbage label from blowing
 * out the panel), NOT HTML escaping. A non-string input yields "".
 */
export function normalizeBookmarkLabel(
  value: unknown,
  maxLength: number = PDF_BOOKMARK_MAX_LABEL,
): string {
  if (typeof value !== "string") return "";
  // Replace whitespace + C0/DEL/C1 control ranges (which \s misses) with a single
  // space, then collapse runs and trim — same approach as sanitizeOutlineTitle.
  const collapsed = value
    .replace(/[\s\u0000-\u001f\u007f-\u009f]+/g, " ")
    .trim();
  const cap =
    Number.isFinite(maxLength) && maxLength > 0
      ? Math.trunc(maxLength)
      : PDF_BOOKMARK_MAX_LABEL;
  if (collapsed.length <= cap) return collapsed;
  return `${collapsed.slice(0, Math.max(0, cap - 1)).trimEnd()}…`;
}

/** A sensible fallback label ("Page N") when the user names a bookmark nothing. */
export function defaultBookmarkLabel(page: number): string {
  const n = Number.isInteger(page) && page >= 1 ? page : 1;
  return `Page ${n}`;
}

/**
 * Validate + normalize a persisted `pdf_bookmark` payload into a usable
 * bookmark, or return null when it is structurally invalid (mirrors
 * `sanitizeNotePayload`). An out-of-range page is rejected; an empty/blank label
 * falls back to "Page N" so a bookmark always has a clickable, readable name.
 */
export function sanitizeBookmarkPayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): UserBookmark | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const record = payload as Record<string, unknown>;
  const page = Number(record.page);
  if (!Number.isInteger(page) || page < 1 || page > MAX_PAGE) return null;
  const rawLabel = record.label;
  if (typeof rawLabel === "string" && rawLabel.length > PDF_BOOKMARK_MAX_LABEL * 8) {
    // Reject grossly oversized blobs outright (defense before normalization).
    return null;
  }
  const label = normalizeBookmarkLabel(rawLabel) || defaultBookmarkLabel(page);
  return { id, page, label };
}

/**
 * Build the payload to POST for a new bookmark. The page is bounded and the
 * label normalized (with a "Page N" fallback) so what we persist matches what we
 * render. Returns null when the page number is out of range.
 */
export function buildBookmarkPayload(input: {
  page: number;
  label?: string;
}): { page: number; label: string } | null {
  const page = Number(input.page);
  if (!Number.isInteger(page) || page < 1 || page > MAX_PAGE) return null;
  const label = normalizeBookmarkLabel(input.label) || defaultBookmarkLabel(page);
  return { page, label };
}

/**
 * Sort bookmarks for display: ascending page number, then case-insensitive
 * label, then id (stable tiebreak). Returns a NEW array; input untouched. So the
 * panel reads top-to-bottom in document order regardless of insertion order.
 */
export function sortBookmarks<T extends { page: number; label: string; id: string }>(
  bookmarks: ReadonlyArray<T>,
): T[] {
  if (!Array.isArray(bookmarks)) return [];
  return bookmarks.slice().sort((a, b) => {
    if (a.page !== b.page) return a.page - b.page;
    const byLabel = a.label.localeCompare(b.label, undefined, {
      sensitivity: "base",
    });
    if (byLabel !== 0) return byLabel;
    return a.id < b.id ? -1 : a.id > b.id ? 1 : 0;
  });
}

/**
 * Cap the number of bookmarks kept to {@link PDF_BOOKMARK_MAX}, dropping the
 * OLDEST first (the tail — most recent — is what the user just added). Returns a
 * new array preserving INSERTION order; input untouched. Mirrors
 * `capShapesPerPage` / `capStrokesPerPage`.
 */
export function capBookmarks<T>(
  bookmarks: ReadonlyArray<T>,
  max = PDF_BOOKMARK_MAX,
): T[] {
  if (!Array.isArray(bookmarks)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : PDF_BOOKMARK_MAX;
  if (bookmarks.length <= cap) return bookmarks.slice();
  return bookmarks.slice(bookmarks.length - cap);
}
