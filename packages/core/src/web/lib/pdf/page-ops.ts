/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure edit-state logic for PDF reader page operations
 * (rotate / delete + restore / reorder / insert blank).
 *
 * The PDFViewer already CONSUMES four structural-edit props and renders the
 * transformed result:
 *
 *   - `pageRotations: Record<number, 0|90|180|270>` — absolute rotation per
 *     ORIGINAL 1-based page,
 *   - `deletedPages: number[]` — original 1-based page numbers removed,
 *   - `pageMoves: Array<{ page, afterPage }>` — original page moved after
 *     another original page (`afterPage === 0` ⇒ move to the front),
 *   - `insertedPages: Array<{ id, afterPage, width?, height? }>` — blank pages
 *     inserted after an original page (`afterPage === 0` ⇒ front).
 *
 * This module is the single source of truth for PRODUCING those edits from
 * user gestures and for RESOLVING them into a final ordered+rotated+filtered
 * page list. It has zero DOM / pdf.js / React dependency so it can be unit
 * tested in isolation and reused by both the viewer and (potentially) the
 * native editor's command pipeline.
 *
 * The resolver intentionally mirrors PDFViewer's existing `visiblePages` /
 * `visiblePageItems` derivation (filter deleted → apply ordered moves → splice
 * inserted blanks) so the produced state renders identically.
 *
 * Persistence note: every field of {@link PageEditState} maps 1:1 onto the
 * additive, OPTIONAL fields already present in the native PDF page schema
 * (`apps/web/lib/native-page-schemas.ts`) and onto the props PDFViewer
 * consumes, so the state round-trips without any new persisted shape.
 */

export type PdfRotation = 0 | 90 | 180 | 270;

/**
 * A page inserted after an original 1-based page (0 ⇒ front). By default this is
 * a BLANK page sized `width`×`height`. When {@link copyOfPage} is set it is a
 * DUPLICATE of that original 1-based page instead: the viewer renders the source
 * page in the slot and the flatten export copies the source page's content (so
 * "duplicate page N" rides the exact same inserted-pages model + persistence as
 * blank inserts — no new structural prop). `width`/`height` are kept for layout
 * fallback but are ignored when a copy renders the real source page.
 */
export interface InsertedBlankPage {
  id: string;
  afterPage: number;
  width: number;
  height: number;
  /** Original 1-based page this slot duplicates; absent ⇒ a blank insert. */
  copyOfPage?: number;
}

/** Original page moved after another original page (afterPage 0 ⇒ front). */
export interface PageMove {
  page: number;
  afterPage: number;
}

/**
 * The complete, normalized page-edit state. Mirrors the four PDFViewer props.
 * `rotations` is keyed by ORIGINAL 1-based page number and stores an ABSOLUTE
 * PDF `/Rotate` value. A missing entry means "use the source page's intrinsic
 * rotation"; an explicit 0 is therefore meaningful for a source page that
 * arrived rotated. The state is always kept normalized by the operations below
 * (no NaN keys, deleted set deduped + sorted, etc.).
 */
export interface PageEditState {
  rotations: Record<number, PdfRotation>;
  deletedPages: number[];
  pageMoves: PageMove[];
  insertedPages: InsertedBlankPage[];
}

/** Default dimensions (US Letter, in PDF points) for a fresh blank page. */
export const DEFAULT_BLANK_PAGE_WIDTH = 612;
export const DEFAULT_BLANK_PAGE_HEIGHT = 792;

/**
 * Defensive upper bounds applied when DESERIALIZING untrusted page-edit state
 * (e.g. a corrupt or maliciously crafted localStorage value). They cap how much
 * work / memory a single normalize pass can consume so a giant array or object
 * cannot bloat memory or hang the tab. They are far above any realistic
 * document, so legitimate edits are never affected.
 */
/** Max number of inserted blank pages accepted from untrusted input. */
export const MAX_INSERTED_PAGES = 1_000;
/** Max number of original pages we ever address (rotations / deletes / moves). */
export const MAX_PAGE_ENTRIES = 10_000;
/** Max length of an inserted-page id; longer ids are regenerated. */
export const MAX_INSERTED_ID_LENGTH = 256;
/** Safe charset for an inserted-page id (no control chars / unbounded growth). */
const INSERTED_ID_PATTERN = /^[A-Za-z0-9_-]{1,256}$/;

/** True when an id is safe to keep as-is (bounded length + safe charset). */
export function isSafeInsertedId(value: unknown): value is string {
  return typeof value === "string" && INSERTED_ID_PATTERN.test(value);
}

/** An empty, fully-normalized edit state. */
export function emptyPageEditState(): PageEditState {
  return { rotations: {}, deletedPages: [], pageMoves: [], insertedPages: [] };
}

function isPositiveInt(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 1;
}

/** Normalize an arbitrary degree value to one of {0, 90, 180, 270}. */
export function normalizeRotation(value: unknown): PdfRotation {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  const normalized = ((Math.trunc(numeric) % 360) + 360) % 360;
  return ([0, 90, 180, 270] as const).includes(normalized as PdfRotation)
    ? (normalized as PdfRotation)
    : 0;
}

/**
 * Coerce any partial / legacy / fresh-from-props shape into a normalized
 * {@link PageEditState}. Accepts both the PDFViewer prop shape
 * (`pageRotations` record, `deletedPages`, `pageMoves`, `insertedPages`) and an
 * already-{@link PageEditState}. Invalid entries are dropped, never thrown.
 */
export function normalizePageEditState(input?: {
  rotations?: Record<number, number> | null;
  pageRotations?: Record<number, number> | null;
  deletedPages?: number[] | null;
  pageMoves?: Array<{ page: number; afterPage: number }> | null;
  insertedPages?: Array<{
    id?: string;
    afterPage?: number;
    width?: number;
    height?: number;
    copyOfPage?: number;
  }> | null;
} | null): PageEditState {
  const state = emptyPageEditState();
  if (!input) return state;

  const rawRotations = input.rotations ?? input.pageRotations ?? null;
  if (rawRotations && typeof rawRotations === "object") {
    // Cap the number of rotation entries kept so a giant object cannot exhaust
    // memory. We stop once MAX_PAGE_ENTRIES valid entries have been accepted.
    for (const [key, value] of Object.entries(rawRotations)) {
      if (Object.keys(state.rotations).length >= MAX_PAGE_ENTRIES) break;
      const page = Number(key);
      if (!isPositiveInt(page) || page > MAX_PAGE_ENTRIES) continue;
      const rotation = normalizeRotation(value);
      // Preserve a canonical explicit zero: it is the only way to represent
      // "remove the source page's intrinsic /Rotate" after reopening. Invalid
      // values that merely normalize to zero (45, NaN, booleans, etc.) remain
      // rejected instead of becoming accidental overrides.
      if (
        rotation !== 0 ||
        (typeof value === "number" && Number.isFinite(value) && value === 0)
      ) {
        state.rotations[page] = rotation;
      }
    }
  }

  if (Array.isArray(input.deletedPages)) {
    const seen = new Set<number>();
    for (const page of input.deletedPages) {
      if (seen.size >= MAX_PAGE_ENTRIES) break;
      if (!isPositiveInt(page) || page > MAX_PAGE_ENTRIES || seen.has(page))
        continue;
      seen.add(page);
    }
    state.deletedPages = [...seen].sort((a, b) => a - b);
  }

  if (Array.isArray(input.pageMoves)) {
    const movesByPage = new Map<number, PageMove>();
    for (const move of input.pageMoves) {
      if (movesByPage.size >= MAX_PAGE_ENTRIES && !movesByPage.has(move?.page ?? -1))
        break;
      if (
        !move ||
        !isPositiveInt(move.page) ||
        move.page > MAX_PAGE_ENTRIES ||
        !Number.isInteger(move.afterPage) ||
        move.afterPage < 0 ||
        move.afterPage > MAX_PAGE_ENTRIES ||
        move.afterPage === move.page
      ) {
        continue;
      }
      movesByPage.delete(move.page);
      movesByPage.set(move.page, {
        page: move.page,
        afterPage: move.afterPage,
      });
    }
    state.pageMoves = [...movesByPage.values()];
  }

  if (Array.isArray(input.insertedPages)) {
    let autoIndex = 0;
    const insertedIds = new Set<string>();
    for (const inserted of input.insertedPages) {
      // Stop once the cap is reached so a 100k-entry array cannot bloat memory.
      if (state.insertedPages.length >= MAX_INSERTED_PAGES) break;
      if (
        !inserted ||
        !Number.isInteger(inserted.afterPage) ||
        (inserted.afterPage ?? -1) < 0 ||
        (inserted.afterPage ?? -1) > MAX_PAGE_ENTRIES
      ) {
        continue;
      }
      // Keep the id only when it is bounded + a safe charset; otherwise mint a
      // deterministic fallback so a malicious id cannot grow memory / break UI.
      // A `copyOfPage` is kept only when it is a valid in-range page number; an
      // out-of-range / non-page value drops the field (the slot stays a blank).
      const copyOfPage =
        isPositiveInt(inserted.copyOfPage) &&
        (inserted.copyOfPage as number) <= MAX_PAGE_ENTRIES
          ? (inserted.copyOfPage as number)
          : undefined;
      let insertedId = isSafeInsertedId(inserted.id)
        ? inserted.id
        : `inserted-${autoIndex}`;
      while (insertedIds.has(insertedId)) {
        autoIndex += 1;
        insertedId = `inserted-${autoIndex}`;
      }
      insertedIds.add(insertedId);
      state.insertedPages.push({
        id: insertedId,
        afterPage: inserted.afterPage as number,
        width:
          typeof inserted.width === "number" && inserted.width > 0
            ? inserted.width
            : DEFAULT_BLANK_PAGE_WIDTH,
        height:
          typeof inserted.height === "number" && inserted.height > 0
            ? inserted.height
            : DEFAULT_BLANK_PAGE_HEIGHT,
        ...(copyOfPage !== undefined ? { copyOfPage } : {}),
      });
      autoIndex += 1;
    }
  }

  return state;
}

/** Read the current absolute rotation for an original 1-based page. */
export function rotationFor(state: PageEditState, page: number): PdfRotation {
  return normalizeRotation(state.rotations[page] ?? 0);
}

/**
 * Resolve a page's displayed absolute rotation against the source PDF.
 * A persisted override wins even when it is explicitly 0; otherwise the
 * source page's intrinsic `/Rotate` value is used.
 */
export function effectiveRotationFor(
  state: PageEditState,
  page: number,
  sourceRotation: unknown = 0,
): PdfRotation {
  return Object.prototype.hasOwnProperty.call(state.rotations, page)
    ? normalizeRotation(state.rotations[page])
    : normalizeRotation(sourceRotation);
}

/**
 * Rotate an original page by `steps` * 90deg (default +1 step = +90deg). The
 * stored rotation is ABSOLUTE so it composes with the page's intrinsic
 * orientation exactly like PDFViewer's render path expects. Returning to the
 * source rotation removes the key; rotating a source-rotated page to 0 keeps an
 * explicit zero override. No-op for non-page inputs.
 */
export function rotatePage(
  state: PageEditState,
  page: number,
  steps = 1,
  sourceRotation: unknown = 0,
): PageEditState {
  if (!isPositiveInt(page)) return state;
  const safeSteps = Number.isFinite(steps) ? Math.trunc(steps) : 1;
  const source = normalizeRotation(sourceRotation);
  const current = effectiveRotationFor(state, page, source);
  const next = normalizeRotation(current + safeSteps * 90);
  const rotations = { ...state.rotations };
  if (next === source) {
    delete rotations[page];
  } else {
    rotations[page] = next;
  }
  return { ...state, rotations };
}

/**
 * Rotate EVERY original page by `steps` * 90deg (default +1). Applies the same
 * absolute composition as {@link rotatePage} to each page in `originalPages`,
 * collapsing any net-zero result to a removed key so the state stays minimal.
 * Deleted pages are rotated too (so a later restore keeps the document-wide
 * orientation). Non-page entries in `originalPages` are ignored. Returns the
 * state unchanged when there is nothing to rotate (empty doc / steps ≡ 0 mod 4).
 */
export function rotateAllPages(
  state: PageEditState,
  originalPages: number[],
  steps = 1,
  sourceRotations: Readonly<Record<number, number>> = {},
): PageEditState {
  const safeSteps = Number.isFinite(steps) ? Math.trunc(steps) : 1;
  // A whole number of full turns is a no-op; skip the work + churn.
  if (((safeSteps % 4) + 4) % 4 === 0) return state;
  const rotations = { ...state.rotations };
  let changed = false;
  for (const page of originalPages) {
    if (!isPositiveInt(page)) continue;
    const source = normalizeRotation(sourceRotations[page]);
    const current = Object.prototype.hasOwnProperty.call(rotations, page)
      ? normalizeRotation(rotations[page])
      : source;
    const next = normalizeRotation(current + safeSteps * 90);
    if (next === source) {
      if (page in rotations) {
        delete rotations[page];
        changed = true;
      }
    } else if (rotations[page] !== next) {
      rotations[page] = next;
      changed = true;
    }
  }
  return changed ? { ...state, rotations } : state;
}

/** True when an original page is currently in the deleted set. */
export function isPageDeleted(state: PageEditState, page: number): boolean {
  return state.deletedPages.includes(page);
}

/**
 * Toggle an original page's deleted status (delete ↔ restore). Deletion keeps
 * the page in the deleted SET — restoring simply removes it, so no annotations
 * / rotations are lost. The set is kept sorted + deduped.
 */
export function toggleDeletePage(
  state: PageEditState,
  page: number,
): PageEditState {
  if (!isPositiveInt(page)) return state;
  if (state.deletedPages.includes(page)) {
    return {
      ...state,
      deletedPages: state.deletedPages.filter((p) => p !== page),
    };
  }
  return {
    ...state,
    deletedPages: [...state.deletedPages, page].sort((a, b) => a - b),
  };
}

/** Explicitly restore (un-delete) an original page; no-op if not deleted. */
export function restorePage(state: PageEditState, page: number): PageEditState {
  if (!state.deletedPages.includes(page)) return state;
  return {
    ...state,
    deletedPages: state.deletedPages.filter((p) => p !== page),
  };
}

/**
 * The moves that turn `originalOrder` into `desiredOrder` when replayed (here
 * and by the server's `_pdf_apply_page_moves`, the same replay): one entry a
 * page at most, in the order they must run.
 *
 * Pages that already stand in the desired relative order (a longest run of
 * them) stay; every other page is placed after the page it follows in the
 * desired order, left to right. After each such move the pages placed so far
 * and the ones that stay are in the desired relative order, so the last move
 * leaves exactly `desiredOrder`.
 *
 * `desiredOrder` must be a permutation of `originalOrder`; anything else
 * yields no moves.
 */
export function pageMovesForOrder(
  originalOrder: readonly number[],
  desiredOrder: readonly number[],
): PageMove[] {
  if (
    originalOrder.length !== desiredOrder.length ||
    new Set(desiredOrder).size !== desiredOrder.length ||
    desiredOrder.some((page) => !originalOrder.includes(page))
  ) {
    return [];
  }
  const rank = new Map(originalOrder.map((page, index) => [page, index]));
  // Longest increasing run of original ranks along the desired order
  // (patience sorting), with back-pointers to recover its members.
  const tails: number[] = [];
  const previous: number[] = new Array(desiredOrder.length).fill(-1);
  for (let index = 0; index < desiredOrder.length; index += 1) {
    const value = rank.get(desiredOrder[index]!)!;
    let low = 0;
    let high = tails.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (rank.get(desiredOrder[tails[middle]!]!)! < value) low = middle + 1;
      else high = middle;
    }
    if (low > 0) previous[index] = tails[low - 1]!;
    tails[low] = index;
  }
  const stays = new Set<number>();
  for (let at = tails.length > 0 ? tails[tails.length - 1]! : -1; at >= 0; at = previous[at]!) {
    stays.add(desiredOrder[at]!);
  }
  const moves: PageMove[] = [];
  desiredOrder.forEach((page, index) => {
    if (!stays.has(page)) moves.push({ page, afterPage: index === 0 ? 0 : desiredOrder[index - 1]! });
  });
  return moves;
}














export function movePageAfter(
  state: PageEditState,
  page: number,
  afterPage: number,
  originalPages?: readonly number[],
): PageEditState {
  if (
    !isPositiveInt(page) ||
    !Number.isInteger(afterPage) ||
    afterPage < 0 ||
    afterPage === page
  ) {
    return state;
  }
  if (originalPages) {
    const current = resolveOriginalOrder(state, [...originalPages]);
    if (!current.includes(page) || (afterPage !== 0 && !current.includes(afterPage))) return state;
    const desired = current.filter((candidate) => candidate !== page);
    desired.splice(afterPage === 0 ? 0 : desired.indexOf(afterPage) + 1, 0, page);
    const deleted = new Set(state.deletedPages);
    const source = originalPages.filter(
      (candidate) => isPositiveInt(candidate) && !deleted.has(candidate),
    );
    return { ...state, pageMoves: pageMovesForOrder(source, desired) };
  }
  const withoutPage = state.pageMoves.filter((m) => m.page !== page);
  return { ...state, pageMoves: [...withoutPage, { page, afterPage }] };
}

/**
 * Move a page up or down by one slot in the CURRENT resolved (visible) order of
 * ORIGINAL pages. Inserted blanks are ignored for the purpose of ordering
 * originals (they ride along after their anchor). `originalPages` is the
 * intrinsic 1-based page list of the document (e.g. [1, 2, 3, 4]). Returns the
 * state unchanged when the move would fall off either end.
 */
export function movePageByOffset(
  state: PageEditState,
  originalPages: number[],
  page: number,
  offset: number,
): PageEditState {
  if (!isPositiveInt(page) || offset === 0) return state;
  const order = resolveOriginalOrder(state, originalPages);
  const index = order.indexOf(page);
  if (index < 0) return state;
  const targetIndex = index + offset;
  if (targetIndex < 0 || targetIndex >= order.length) return state;
  // Compute the original page the moved page should sit AFTER at its new slot.
  // Build the order WITHOUT the moving page, then read its left neighbor.
  const without = order.filter((p) => p !== page);
  const insertBefore =
    offset < 0
      ? // moving up: land before the page currently at targetIndex
        order[targetIndex]
      : // moving down: land after the page currently at targetIndex
        order[targetIndex];
  if (offset < 0) {
    const neighborIndex = without.indexOf(insertBefore);
    const afterPage = neighborIndex <= 0 ? 0 : without[neighborIndex - 1];
    return movePageAfter(state, page, afterPage, originalPages);
  }
  // moving down: place after the page at targetIndex
  return movePageAfter(state, page, insertBefore, originalPages);
}

/**
 * Reorder for a DRAG-AND-DROP gesture in the thumbnail rail: move `source`
 * (an original 1-based page) to sit relative to `target` (another original
 * page), with `dropAfter` choosing whether it lands immediately AFTER the
 * target (the drop fell on the target's lower half) or immediately BEFORE it
 * (upper half). The new slot is encoded as a {@link movePageAfter} against the
 * CURRENT resolved (visible) order of originals — so the reorder rides the
 * existing `pageMoves` model + persistence exactly like the up/down buttons.
 *
 * Pure index remap (no DOM): for a BEFORE drop the anchor is the target's left
 * neighbour in the order WITHOUT the source (front ⇒ afterPage 0); for an AFTER
 * drop the anchor is the target itself. Returns the state unchanged when either
 * page is not in the visible order (e.g. a deleted page) or `source === target`.
 */
export function reorderPageByDrop(
  state: PageEditState,
  originalPages: number[],
  source: number,
  target: number,
  dropAfter: boolean,
): PageEditState {
  if (!isPositiveInt(source) || !isPositiveInt(target) || source === target) {
    return state;
  }
  const order = resolveOriginalOrder(state, originalPages);
  if (!order.includes(source) || !order.includes(target)) return state;
  if (dropAfter) {
    return movePageAfter(state, source, target, originalPages);
  }
  // Insert BEFORE the target: anchor to the target's left neighbour in the
  // order WITHOUT the source (so dropping just above its own position is a
  // clean no-op), front (0) when the target is first there.
  const without = order.filter((p) => p !== source);
  const targetIndex = without.indexOf(target);
  const afterPage = targetIndex > 0 ? without[targetIndex - 1] : 0;
  return movePageAfter(state, source, afterPage, originalPages);
}

/**
 * Insert a blank page after an original 1-based page (`afterPage === 0` ⇒
 * front). Returns the new state plus the generated insertion id so the caller
 * can focus / scroll to it. `idFactory` defaults to a time+random id but is
 * injectable for deterministic tests.
 */
export function insertBlankPage(
  state: PageEditState,
  afterPage: number,
  options?: {
    width?: number;
    height?: number;
    idFactory?: () => string;
  },
): { state: PageEditState; id: string } {
  const safeAfter =
    Number.isInteger(afterPage) && afterPage >= 0 && afterPage <= 10_000
      ? afterPage
      : 0;
  const id =
    options?.idFactory?.() ??
    `inserted-${Date.now().toString(36)}-${Math.random()
      .toString(36)
      .slice(2, 8)}`;
  const inserted: InsertedBlankPage = {
    id,
    afterPage: safeAfter,
    width:
      options?.width && options.width > 0
        ? options.width
        : DEFAULT_BLANK_PAGE_WIDTH,
    height:
      options?.height && options.height > 0
        ? options.height
        : DEFAULT_BLANK_PAGE_HEIGHT,
  };
  return {
    state: { ...state, insertedPages: [...state.insertedPages, inserted] },
    id,
  };
}

/**
 * Insert a DUPLICATE of an original 1-based `sourcePage` immediately after a
 * given `afterPage` (`afterPage === 0` ⇒ front; defaults to right after the
 * source itself, the Acrobat "Duplicate Pages" behaviour). The duplicate rides
 * the inserted-pages model with `copyOfPage` set — the viewer renders the source
 * page in the slot and the flatten export copies its content. `width`/`height`
 * default to the source page's size when supplied so a blank-fallback (older
 * renderer) still lays out sensibly. Returns the new state + the generated id.
 * No-op (returns the same state + an empty id) for a non-page source.
 */
export function duplicatePage(
  state: PageEditState,
  sourcePage: number,
  options?: {
    afterPage?: number;
    width?: number;
    height?: number;
    idFactory?: () => string;
  },
): { state: PageEditState; id: string } {
  if (!isPositiveInt(sourcePage) || sourcePage > MAX_PAGE_ENTRIES) {
    return { state, id: "" };
  }
  const afterRaw = options?.afterPage ?? sourcePage;
  const safeAfter =
    Number.isInteger(afterRaw) && afterRaw >= 0 && afterRaw <= MAX_PAGE_ENTRIES
      ? afterRaw
      : sourcePage;
  const id =
    options?.idFactory?.() ??
    `inserted-${Date.now().toString(36)}-${Math.random()
      .toString(36)
      .slice(2, 8)}`;
  const inserted: InsertedBlankPage = {
    id,
    afterPage: safeAfter,
    width:
      options?.width && options.width > 0
        ? options.width
        : DEFAULT_BLANK_PAGE_WIDTH,
    height:
      options?.height && options.height > 0
        ? options.height
        : DEFAULT_BLANK_PAGE_HEIGHT,
    copyOfPage: sourcePage,
  };
  return {
    state: { ...state, insertedPages: [...state.insertedPages, inserted] },
    id,
  };
}

/** Remove a previously-inserted blank page by id; no-op if unknown. */
export function removeInsertedPage(
  state: PageEditState,
  id: string,
): PageEditState {
  if (!state.insertedPages.some((p) => p.id === id)) return state;
  return {
    ...state,
    insertedPages: state.insertedPages.filter((p) => p.id !== id),
  };
}

/**
 * Resolve the ordered list of VISIBLE original pages after applying deletions
 * and moves. Mirrors PDFViewer's `visiblePages` derivation exactly: filter the
 * deleted set, then apply each move in array order (afterPage 0 ⇒ front; a move
 * whose anchor is missing is skipped). `originalPages` is the intrinsic 1-based
 * page list (e.g. [1, 2, 3]).
 */
export function resolveOriginalOrder(
  state: PageEditState,
  originalPages: number[],
): number[] {
  const deleted = new Set(state.deletedPages);
  const ordered = originalPages.filter(
    (page) => isPositiveInt(page) && !deleted.has(page),
  );
  for (const move of state.pageMoves) {
    const movingIndex = ordered.indexOf(move.page);
    if (movingIndex < 0) continue;
    if (move.afterPage !== 0 && !ordered.includes(move.afterPage)) {
      continue;
    }
    ordered.splice(movingIndex, 1);
    if (move.afterPage === 0) {
      ordered.unshift(move.page);
      continue;
    }
    const targetIndex = ordered.indexOf(move.afterPage);
    ordered.splice(targetIndex + 1, 0, move.page);
  }
  return ordered;
}

/** A single resolved slot in the final page sequence. */
export type ResolvedPageItem =
  | { kind: "original"; page: number; rotation: PdfRotation }
  | { kind: "inserted"; page: InsertedBlankPage };

/**
 * Resolve the FINAL ordered page sequence the viewer renders: visible originals
 * (rotation-annotated) interleaved with inserted blanks after their anchor
 * (afterPage 0 ⇒ front). Mirrors PDFViewer's `visiblePageItems`: front-inserts
 * first, then for each visible original the page followed by any blanks anchored
 * to it. Inserted blanks are sorted by `afterPage` then insertion order, like
 * the viewer's `normalizedInsertedPages`.
 */
export function resolvePageItems(
  state: PageEditState,
  originalPages: number[],
): ResolvedPageItem[] {
  const order = resolveOriginalOrder(state, originalPages);
  const visibleSet = new Set(order);
  // A blank anchored after a page since deleted stays where it was: after
  // the nearest earlier original still shown (or at the front). Dropped from
  // the sequence, it vanished from the reader while the saved file still
  
  // showed 21 pages for a 22-page file).
  const anchorFor = (afterPage: number): number => {
    let anchor = afterPage;
    while (anchor > 0 && !visibleSet.has(anchor)) anchor -= 1;
    return anchor;
  };
  const sortedInserts = state.insertedPages
    .map((page, index) => ({ ...page, afterPage: anchorFor(page.afterPage), order: index }))
    .sort((a, b) => a.afterPage - b.afterPage || a.order - b.order);
  const insertsFor = (anchor: number) =>
    sortedInserts
      .filter((p) => p.afterPage === anchor)
      .map<ResolvedPageItem>((p) => ({
        kind: "inserted",
        page: {
          id: p.id,
          afterPage: p.afterPage,
          width: p.width,
          height: p.height,
          ...(p.copyOfPage !== undefined ? { copyOfPage: p.copyOfPage } : {}),
        },
      }));

  const items: ResolvedPageItem[] = [];
  // Front-inserted blanks (afterPage 0).
  items.push(...insertsFor(0));
  for (const page of order) {
    items.push({ kind: "original", page, rotation: rotationFor(state, page) });
    items.push(...insertsFor(page));
  }
  return items;
}

/** One slot of the page sequence the reader sees: an original page, or a
 * page inserted into the document (by id). */
export type PdfPageSlot =
  | { kind: "original"; page: number }
  | { kind: "inserted"; id: string };

/** The page sequence the reader sees, as slots. */
export function pageSlots(
  state: PageEditState,
  originalPages: number[],
): PdfPageSlot[] {
  return resolvePageItems(state, originalPages).map((item) =>
    item.kind === "original"
      ? { kind: "original", page: item.page }
      : { kind: "inserted", id: item.page.id },
  );
}

/**
 * The sequence with the slot at `index` swapped with its neighbour `offset`
 * (+1 down, -1 up) away: "move this page down" moves the page shown past the
 * next page shown, an inserted blank or an original alike (user audit PDF
 * 0008: on an inserted blank it moved the original page behind it, and an
 * original moved past every blank anchored to its neighbour). Null at either
 * end.
 */
export function swapPageSlot(
  slots: readonly PdfPageSlot[],
  index: number,
  offset: -1 | 1,
): PdfPageSlot[] | null {
  const target = index + offset;
  if (index < 0 || index >= slots.length || target < 0 || target >= slots.length) {
    return null;
  }
  const next = [...slots];
  next[index] = slots[target]!;
  next[target] = slots[index]!;
  return next;
}

/**
 * The page edit state that shows `sequence`, the pages the reader sees in a
 * new order: the originals' order becomes the canonical moves
 * ({@link pageMovesForOrder}), and each inserted page is anchored after the
 * original before it (0 at the front), anchored pages keeping the sequence's
 * order. An inserted page the sequence does not hold (anchored to a deleted
 * page) is kept as it is. Null when the sequence is not exactly the pages on
 * screen.
 */
export function pageEditStateForSequence(
  state: PageEditState,
  originalPages: readonly number[],
  sequence: readonly PdfPageSlot[],
): PageEditState | null {
  const current = pageSlots(state, [...originalPages]);
  const key = (slot: PdfPageSlot) =>
    slot.kind === "original" ? `p:${slot.page}` : `i:${slot.id}`;
  const currentKeys = current.map(key).sort();
  const sequenceKeys = sequence.map(key).sort();
  if (
    currentKeys.length !== sequenceKeys.length ||
    currentKeys.some((value, index) => value !== sequenceKeys[index])
  ) {
    return null;
  }
  const deleted = new Set(state.deletedPages);
  const source = originalPages.filter(
    (page) => isPositiveInt(page) && !deleted.has(page),
  );
  const desired = sequence.flatMap((slot) =>
    slot.kind === "original" ? [slot.page] : [],
  );
  const byId = new Map(state.insertedPages.map((page) => [page.id, page]));
  const placed: InsertedBlankPage[] = [];
  let anchor = 0;
  for (const slot of sequence) {
    if (slot.kind === "original") {
      anchor = slot.page;
      continue;
    }
    const page = byId.get(slot.id);
    if (page) placed.push({ ...page, afterPage: anchor });
  }
  const placedIds = new Set(placed.map((page) => page.id));
  return {
    ...state,
    pageMoves: pageMovesForOrder(source, desired),
    insertedPages: [
      ...placed,
      ...state.insertedPages.filter((page) => !placedIds.has(page.id)),
    ],
  };
}

/**
 * Convert state → the `pageRotations` record prop PDFViewer consumes.
 * Explicit zero values are retained because they clear a source page's
 * intrinsic rotation.
 */
export function toPageRotationsProp(
  state: PageEditState,
): Record<number, number> {
  const out: Record<number, number> = {};
  for (const [key, value] of Object.entries(state.rotations)) {
    const page = Number(key);
    if (isPositiveInt(page)) out[page] = normalizeRotation(value);
  }
  return out;
}

/** Serialize state into the additive persisted shape (matches the schema). */
export function serializePageEditState(state: PageEditState): {
  pageRotations: Record<number, number>;
  deletedPages: number[];
  pageMoves: PageMove[];
  insertedPages: InsertedBlankPage[];
} {
  return {
    pageRotations: toPageRotationsProp(state),
    deletedPages: [...state.deletedPages],
    pageMoves: state.pageMoves.map((m) => ({ ...m })),
    insertedPages: state.insertedPages.map((p) => ({ ...p })),
  };
}

/**
 * Resolve the page DUPLICATIONS to bake into a flatten export from the inserted
 * pages that carry a `copyOfPage`. Each becomes `{ sourcePage, afterPage }` where
 * `sourcePage` is the original page being copied and `afterPage` is its anchor
 * (0 ⇒ front) — exactly the shape `flattenPdf`'s `duplicatePages` option expects.
 * Blank inserts (no `copyOfPage`) are skipped. Order follows insertion order.
 */
export function resolveExportDuplications(
  state: PageEditState,
): Array<{ sourcePage: number; afterPage: number }> {
  const out: Array<{ sourcePage: number; afterPage: number }> = [];
  for (const inserted of state.insertedPages) {
    if (isPositiveInt(inserted.copyOfPage)) {
      out.push({ sourcePage: inserted.copyOfPage, afterPage: inserted.afterPage });
    }
  }
  return out;
}

/**
 * The exported page ORDER as 1-based SOURCE page numbers in output sequence:
 * one entry per resolved output slot that renders an original page's content —
 * surviving originals in their final order, plus each duplicate slot's source
 * page at the copy's position. Blank inserts carry no source content and are
 * SKIPPED (so e.g. a running page-number stamp numbers only content pages).
 * This is exactly the `FlattenAnnotations.pageOrder` shape the stamp planner
 * consumes.
 */
export function exportPageOrder(
  state: PageEditState,
  originalPages: number[],
): number[] {
  const order: number[] = [];
  for (const item of resolvePageItems(state, originalPages)) {
    if (item.kind === "original") order.push(item.page);
    else if (isPositiveInt(item.page.copyOfPage)) order.push(item.page.copyOfPage);
  }
  return order;
}

/**
 * Map each ORIGINAL 1-based page to its OUTPUT 1-based page number(s) after the
 * structural bake (delete / reorder / insert / duplicate). A deleted page maps
 * to nothing (absent key); a duplicated page maps to multiple outputs (its
 * surviving self plus every copy slot). Derived from {@link resolvePageItems}
 * so it stays byte-for-byte consistent with what the viewer renders and what
 * the flatten export reconstructs. Used to remap page-keyed post-processing
 * (e.g. which exported pages must be rasterized for a redaction).
 */
export function mapOriginalPagesToOutput(
  state: PageEditState,
  originalPages: number[],
): Map<number, number[]> {
  const out = new Map<number, number[]>();
  resolvePageItems(state, originalPages).forEach((item, index) => {
    const source =
      item.kind === "original" ? item.page : item.page.copyOfPage;
    if (!isPositiveInt(source)) return;
    const list = out.get(source);
    if (list) list.push(index + 1);
    else out.set(source, [index + 1]);
  });
  return out;
}

/** True when the state carries no edits at all (used to skip persistence). */
export function isEmptyPageEditState(state: PageEditState): boolean {
  return (
    Object.keys(state.rotations).length === 0 &&
    state.deletedPages.length === 0 &&
    state.pageMoves.length === 0 &&
    state.insertedPages.length === 0
  );
}
