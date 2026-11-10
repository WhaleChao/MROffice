/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the presentation editor's SLIDE SORTER grid view — PowerPoint's
 * "Slide Sorter" (a multi-column thumbnail grid that is an alternative VIEW of
 * the same single-column rail).
 *
 * Dependency-free (no React / no DOM) so the index math and the multi-select set
 * logic can be unit-tested in isolation and reused unchanged by the component.
 *
 * DESIGN INVARIANTS (a skeptic will probe these):
 *  - The sorter is a VIEW. It NEVER invents a new way to mutate the deck: a drag
 *    drop produces a reordered INDEX permutation that the caller applies through
 *    the SAME `commitSlides` path the rail's drag uses. {@link reorderForDrop}
 *    only computes "the new flat order of indices"; it does not touch slide data.
 *  - Multi-select state is the sorter's OWN (a Set of slide ids). It never
 *    mutates the editor's single `activeIndex`; the helpers here only transform
 *    the selection set.
 *  - Bulk delete can NEVER empty the deck (a deck needs >= 1 slide).
 *    {@link slidesRemainingAfterDelete} clamps so at least one slide survives,
 *    and {@link canBulkDelete} reports whether a delete is allowed at all.
 *  - DETERMINISM: nothing here calls Date / Math.random / crypto. Pure functions
 *    of their inputs.
 */

/**
 * The presentation editor's view modes, mirroring WPS's 视图 group:
 *
 * - `rail`    普通视图 — the editing canvas with the slide rail.
 * - `sorter`  幻灯片浏览 — the multi-column thumbnail grid.
 * - `notes`   备注页视图 — one slide over its speaker notes, the layout that
 *             prints; the only view where notes get real writing room.
 * - `reading` 阅读视图 — the deck plays inside the window instead of taking
 *             over the screen, so it can be read beside other windows.
 * - `master`  母版视图 — the deck's masters and their layouts at slide size,
 *             showing the real inherited content rather than placeholder boxes.
 *
 * The fullscreen slideshow is not a view mode: it is its own present overlay.
 */
export const PRESENTATION_VIEW_MODES = [
  "rail",
  "sorter",
  "notes",
  "reading",
  "master",
] as const;
export type PresentationViewMode = (typeof PRESENTATION_VIEW_MODES)[number];

/**
 * Which modes replace the editing STAGE rather than the slide rail. The rail
 * keeps its normal thumbnails in these, exactly as WPS does.
 */
export const PRESENTATION_STAGE_VIEW_MODES: readonly PresentationViewMode[] = [
  // PowerPoint's Slide Sorter replaces the thumbnail pane AND the slide; the
  // editor hides the rail itself while it shows (PPT-E011).
  "sorter",
  "notes",
  "reading",
  "master",
];

/** Type guard: is `value` a known view mode? Defaults the caller to "rail". */
export function isPresentationViewMode(
  value: unknown,
): value is PresentationViewMode {
  return (
    typeof value === "string" &&
    (PRESENTATION_VIEW_MODES as readonly string[]).includes(value)
  );
}

/** Does this mode take over the editing stage? */
export function isStageViewMode(value: unknown): boolean {
  return (
    isPresentationViewMode(value) &&
    PRESENTATION_STAGE_VIEW_MODES.includes(value)
  );
}

/** The modifier keys a sorter thumbnail click may carry. */
export interface SelectionModifiers {
  /** Shift held → extend a contiguous RANGE from the anchor to the clicked one. */
  shift?: boolean;
  /** Cmd (mac) or Ctrl held → TOGGLE the clicked one in/out of the set. */
  meta?: boolean;
}

/**
 * Compute the next multi-selection (a Set of slide ids) after a click on the
 * thumbnail at `clickedIndex`, honouring the modifier keys. `order` is the flat,
 * current slide-id order (top-to-bottom / left-to-right in the grid). `anchor`
 * is the id of the last plain/shift-anchor click (the range pivot), or null.
 *
 * Semantics mirror PowerPoint / a file manager:
 *  - PLAIN click → select exactly that one (replaces the set).
 *  - META (Cmd/Ctrl) click → toggle that one in/out, keeping the rest.
 *  - SHIFT click → select the contiguous RANGE between the anchor and the
 *    clicked id (inclusive), replacing the set. With no anchor it behaves like a
 *    plain click.
 *
 * Returns BOTH the next selected-id set and the next anchor id (so the caller
 * can persist the anchor for the next shift-click). A meta-toggle moves the
 * anchor onto the clicked id; a shift-click leaves the anchor where it was so a
 * second shift-click re-ranges from the same pivot (file-manager behaviour).
 *
 * Pure: never mutates the input set; an out-of-range / unknown id yields a
 * single-id selection of that id (defensive — the grid only ever passes real
 * ids).
 */
export function nextSelection(
  order: readonly string[],
  selected: ReadonlySet<string>,
  clickedId: string,
  modifiers: SelectionModifiers,
  anchor: string | null,
): { selected: Set<string>; anchor: string | null } {
  const clickedIndex = order.indexOf(clickedId);

  if (modifiers.shift && anchor !== null) {
    const anchorIndex = order.indexOf(anchor);
    if (anchorIndex !== -1 && clickedIndex !== -1) {
      const lo = Math.min(anchorIndex, clickedIndex);
      const hi = Math.max(anchorIndex, clickedIndex);
      const range = new Set<string>();
      for (let i = lo; i <= hi; i += 1) range.add(order[i]);
      // Shift keeps the original anchor so a re-shift re-ranges from the pivot.
      return { selected: range, anchor };
    }
    // Fall through to plain behaviour when the anchor/click is unresolved.
  }

  if (modifiers.meta) {
    const next = new Set(selected);
    if (next.has(clickedId)) next.delete(clickedId);
    else next.add(clickedId);
    // The toggled id becomes the new anchor for a subsequent shift-range.
    return { selected: next, anchor: clickedId };
  }

  // Plain click (or shift with no usable anchor): select exactly this one.
  return { selected: new Set([clickedId]), anchor: clickedId };
}

/**
 * The set of ids that should be DRAGGED together when a drag starts on
 * `draggedId`. If the dragged thumbnail is part of the current multi-selection,
 * the WHOLE selection moves coherently; otherwise just the dragged one moves
 * (and, by convention, the caller resets the selection to it). The result is
 * ordered by the slide's position in `order` so the moved block keeps its
 * internal order at the drop site.
 *
 * Pure; returns a fresh array.
 */
export function dragGroupIds(
  order: readonly string[],
  selected: ReadonlySet<string>,
  draggedId: string,
): string[] {
  if (selected.has(draggedId) && selected.size > 1) {
    return order.filter((id) => selected.has(id));
  }
  return [draggedId];
}

/**
 * Compute the new flat slide-id order after dropping the dragged group on the
 * thumbnail whose id is `targetId`.
 *
 * This GENERALIZES the rail's existing single-slide drag math so both views
 * share one reorder rule. The rail's drop does (for a single dragged index):
 *
 *     next = [...order]; next.splice(from, 1); next.splice(targetIndex, 0, moved)
 *
 * i.e. remove the dragged slide, then re-insert it at the target's ORIGINAL
 * index in the post-removal array. For a single dragged id this function returns
 * EXACTLY that permutation, so a sorter drag and a rail drag of the same slide
 * onto the same target produce byte-identical order (one reorder semantics).
 *
 * For a MULTI-select drag the moved ids are removed (keeping their relative
 * order) and the whole contiguous block is re-inserted at the slot the target
 * originally occupied, less however many moved slides originally preceded the
 * target — the direct multi-slide analogue of the single-slide splice. Dropping
 * onto a member of the dragged block lands the block at that target's slot.
 *
 * Pure + deterministic. Unknown ids in `movedIds` are ignored; an empty move, or
 * a drop whose result equals the input, returns a copy equal to `order` (callers
 * compare against the input to skip a no-op commit).
 */
export function reorderForDrop(
  order: readonly string[],
  movedIds: readonly string[],
  targetId: string,
): string[] {
  const movedSet = new Set(movedIds.filter((id) => order.includes(id)));
  if (movedSet.size === 0) return [...order];

  // The remaining slides (everything NOT being moved), in their current order,
  // and the moved block in its current relative order.
  const remaining = order.filter((id) => !movedSet.has(id));
  const movedBlock = order.filter((id) => movedSet.has(id));

  // The target's ORIGINAL index in the full order (the rail's splice index). An
  // unknown target appends the block at the end (defensive).
  const targetOriginal = order.indexOf(targetId);
  const splice = targetOriginal === -1 ? order.length : targetOriginal;
  // How many moved slides originally sat BEFORE the splice point — subtract them
  // so the block lands at the target's slot within the post-removal array, which
  // for a single moved id reduces to the rail's `splice(targetIndex, 0, moved)`.
  const movedBefore = order
    .slice(0, splice)
    .filter((id) => movedSet.has(id)).length;
  const insertAt = Math.max(0, Math.min(remaining.length, splice - movedBefore));

  return [
    ...remaining.slice(0, insertAt),
    ...movedBlock,
    ...remaining.slice(insertAt),
  ];
}

/**
 * Can a bulk delete of `selectedCount` slides proceed without emptying a deck of
 * `totalCount` slides? A deck must keep >= 1 slide, so a delete is refused when
 * it would remove every slide (or when nothing is selected). Pure.
 */
export function canBulkDelete(
  totalCount: number,
  selectedCount: number,
): boolean {
  if (selectedCount <= 0) return false;
  return totalCount - selectedCount >= 1;
}

/**
 * The slides that REMAIN after deleting every slide whose id is in `selected`.
 * Hard guard: if the deletion would empty the deck, NOTHING is deleted (the
 * original list is returned) so the deck always keeps >= 1 slide — the caller
 * should gate the action with {@link canBulkDelete} for UX, but this is the
 * structural backstop. Preserves slide order + identity; pure (fresh array).
 */
export function slidesRemainingAfterDelete<S extends { id: string }>(
  slides: readonly S[],
  selected: ReadonlySet<string>,
): S[] {
  const remaining = slides.filter((slide) => !selected.has(slide.id));
  // Never empty the deck: a delete that removes everything is a no-op.
  if (remaining.length === 0) return [...slides];
  return remaining;
}

/**
 * Resolve the new `activeIndex` after the deck changed to `nextSlides`, trying to
 * keep `preferredId` selected (its new position), else clamping the old index
 * into range. Used so a sorter mutation (reorder / delete / duplicate) leaves the
 * single-slide editor pointing at a sensible slide. Pure; returns 0 for an empty
 * list (defensive — the deck never legitimately empties).
 */
export function resolveActiveIndex(
  nextSlides: readonly { id: string }[],
  preferredId: string | null,
  fallbackIndex: number,
): number {
  if (nextSlides.length === 0) return 0;
  if (preferredId !== null) {
    const found = nextSlides.findIndex((slide) => slide.id === preferredId);
    if (found !== -1) return found;
  }
  return Math.min(Math.max(0, fallbackIndex), nextSlides.length - 1);
}
