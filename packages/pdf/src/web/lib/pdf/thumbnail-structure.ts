/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Stable structural-edit signature for the PDF viewer's thumbnail cache.
 *
 * The thumbnail render cache (a `Set<pageNumber>` of already-painted slots) and
 * the in-flight thumbnail render tasks must be invalidated whenever a native
 * PDF edit command changes the structure of the document — pages deleted,
 * reordered, or blank pages inserted. Without that, deleting/reordering/
 * inserting pages leaves stale or blank thumbnails because the cache still
 * marks the (re-mounted) canvas slots as "rendered".
 *
 * React effect dependency arrays compare by reference. The structural-edit
 * props are fresh arrays/objects on every render, so depending on them
 * directly would clear the cache on *every* render (a perf regression and a
 * flicker). These helpers fold the structural edits into a single primitive
 * string that only changes when the structure genuinely changes, giving the
 * effect a stable, value-based dependency.
 *
 * Pure (no pdf.js / jsdom) so it can be unit-tested in isolation.
 */

/** Original 1-based page numbers removed by native PDF edit commands. */
export type DeletedPagesInput = number[] | undefined | null;

/** Original 1-based page numbers moved after another original page. */
export interface PageMoveInput {
  page: number;
  afterPage: number;
}

/** Blank pages inserted after original 1-based page numbers. */
export interface InsertedPageInput {
  id: string;
  afterPage: number;
  width?: number;
  height?: number;
}

/**
 * Fold the deleted-page list into a deterministic, order-independent string.
 * Deleting pages {3, 1} and {1, 3} are the same structural edit, so the
 * signature sorts before serializing.
 */
export function deletedPagesSignature(deleted: DeletedPagesInput): string {
  if (!Array.isArray(deleted) || deleted.length === 0) return "";
  return [...deleted]
    .filter((page) => Number.isFinite(page))
    .map((page) => Math.trunc(page))
    .sort((a, b) => a - b)
    .join(",");
}

/**
 * Fold the page-move list into a deterministic string. Moves are ordered, so
 * the signature preserves the array order (a different application order can
 * produce a different visible page sequence).
 */
export function pageMovesSignature(
  moves: PageMoveInput[] | undefined | null,
): string {
  if (!Array.isArray(moves) || moves.length === 0) return "";
  return moves
    .filter((move) => move && Number.isFinite(move.page) && Number.isFinite(move.afterPage))
    .map((move) => `${Math.trunc(move.page)}>${Math.trunc(move.afterPage)}`)
    .join("|");
}

/**
 * Fold the inserted-page list into a deterministic string. Insertions are keyed
 * by their stable `id` plus the page they follow and (best-effort) dimensions,
 * so re-inserting after a different page or resizing a blank page invalidates
 * the cached thumbnail.
 */
export function insertedPagesSignature(
  inserted: InsertedPageInput[] | undefined | null,
): string {
  if (!Array.isArray(inserted) || inserted.length === 0) return "";
  return inserted
    .filter((page) => page && Number.isFinite(page.afterPage))
    .map(
      (page) =>
        `${page.id ?? ""}@${Math.trunc(page.afterPage)}:${
          Number.isFinite(page.width) ? page.width : ""
        }x${Number.isFinite(page.height) ? page.height : ""}`,
    )
    .join("|");
}

/**
 * Combine all three structural-edit signatures into a single primitive that is
 * stable across renders and only changes when the document structure changes.
 * Suitable as a value-based React effect dependency.
 */
export function structuralEditSignature(input: {
  deletedPages?: DeletedPagesInput;
  pageMoves?: PageMoveInput[] | null;
  insertedPages?: InsertedPageInput[] | null;
}): string {
  return [
    deletedPagesSignature(input.deletedPages),
    pageMovesSignature(input.pageMoves),
    insertedPagesSignature(input.insertedPages),
  ].join("~");
}
