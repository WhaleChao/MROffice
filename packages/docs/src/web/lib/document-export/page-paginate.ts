/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure pagination for the WYSIWYG page-view preview (Word "Print Layout" /
 * Google Docs paged view).
 *
 * Given the measured pixel heights of the document's top-level blocks (in
 * document order) and the usable content height of ONE printed page
 * (`pageContentHeightPx` from native-document/pageSetup.ts), pack the blocks
 * into pages and return the block indices that belong on each page.
 *
 * Kept free of any DOM/React imports so the packing math is unit-testable in
 * isolation and reused by the PagePreview component (which supplies the measured
 * heights) and its tests (which supply fixed heights).
 *
 * v1 packing rules (intentionally tractable — no mid-block splitting):
 *   - WHOLE BLOCKS ONLY. A block never splits across a page boundary. This is
 *     Word's "keep paragraph together" applied unconditionally; it sidesteps the
 *     hard problem of slicing a paragraph's lines across pages (that is v2).
 *   - KEEP-TOGETHER OVERFLOW: a block that would overflow the remaining space on
 *     the current (non-empty) page starts a new page instead.
 *   - TALLER-THAN-PAGE OVERFLOW: a single block taller than a whole page cannot
 *     fit anywhere, so it sits alone on its own page and simply overflows that
 *     page's bottom edge (documented, accepted in v1).
 *   - EMPTY DOC: an empty document yields exactly ONE empty page (`[[]]`) so the
 *     preview always shows at least one paper card.
 *   - HARD BOUNDARY (page / section break): a caller may force a block to begin
 *     on the next, next-even, or next-odd page regardless of remaining space.
 *     Even/odd boundaries can insert an empty parity page, matching the part of
 *     Word's section-break contract that a whole-block preview can model.
 */

/** Page-starting break kinds supported by the preview paginator. */
export type PaginationBreakKind = "nextPage" | "evenPage" | "oddPage";

/**
 * Break directives keyed by the block index that must start the new page.
 * A key equal to `blockHeights.length` represents a terminal break and creates
 * a trailing blank page. A Set is retained as a backwards-compatible shorthand
 * for next-page breaks.
 */
export type PaginationBreaks =
  | ReadonlyMap<number, PaginationBreakKind>
  | ReadonlySet<number>;

/**
 * Pagination result with the physical height each block actually consumes.
 *
 * `effectiveBlockHeights` normally equals the sanitized input height. When a
 * natural overflow moves a block to a fresh sheet, its collapsible leading
 * flow spacing is discarded at the fragmentainer boundary (the same rule Word
 * applies to paragraph spacing before at a naturally-created page top).
 * Explicit page/section breaks deliberately retain that spacing.
 */
export interface PaginatedBlockLayout {
  pages: number[][];
  effectiveBlockHeights: number[];
  /** Block indices whose leading flow spacing was discarded at a natural top. */
  naturalPageStarts: number[];
}

function breakKindAt(
  breaks: PaginationBreaks | undefined,
  index: number,
): PaginationBreakKind | null {
  if (!breaks) return null;
  const maybeMap = breaks as ReadonlyMap<number, PaginationBreakKind>;
  if (typeof maybeMap.get === "function") {
    const value = maybeMap.get(index);
    return value === "evenPage" || value === "oddPage"
      ? value
      : value
        ? "nextPage"
        : null;
  }
  return (breaks as ReadonlySet<number>).has(index) ? "nextPage" : null;
}

/** Insert parity-only blank pages before the next content page when required. */
function appendParityBlankPage(
  pages: number[][],
  kind: PaginationBreakKind,
): void {
  const nextPageNumber = pages.length + 1;
  if (
    (kind === "evenPage" && nextPageNumber % 2 !== 0) ||
    (kind === "oddPage" && nextPageNumber % 2 !== 1)
  ) {
    pages.push([]);
  }
}

/**
 * Pack block heights into pages.
 *
 * @param blockHeights Per-block measured heights in CSS px, in document order.
 *   Non-finite or negative entries are coerced to 0 (a zero-height block always
 *   fits on the current page).
 * @param pageContentHeightPx Usable content height of one printed page in CSS px
 *   (page height minus both vertical margins). Non-positive or non-finite values
 *   are treated as a single unbounded page so the preview degrades to one card
 *   rather than dividing by an invalid height or looping forever.
 * @param breaksBefore Optional break directives for block indices that MUST
 *   begin a new page (for example, the block following a page/section break).
 *   A Set means next-page for compatibility. A Map can request next/even/odd.
 *   A directive on the first block is a no-op; real document break nodes map to
 *   the following index, so a leading break node still produces a blank page.
 * @param flowSpacingBefore Collapsible leading flow already included in each
 *   block height. It is discarded only when a natural overflow creates a new
 *   page; first-page, same-page, and explicit hard-break spacing is preserved.
 * @returns Page membership plus the effective per-block heights used for fill
 *   calculations and the natural page-start indices. Always at least one page.
 */
export function paginateBlockLayout(
  blockHeights: number[],
  pageContentHeightPx: number,
  breaksBefore?: PaginationBreaks,
  flowSpacingBefore: readonly number[] = [],
): PaginatedBlockLayout {
  const effectiveBlockHeights = blockHeights.map((raw) =>
    Number.isFinite(raw) && raw > 0 ? raw : 0,
  );

  // Empty document → one empty page so the preview always renders a paper card.
  if (blockHeights.length === 0) {
    return {
      pages: [[]],
      effectiveBlockHeights,
      naturalPageStarts: [],
    };
  }

  // A non-positive / non-finite page height has no meaningful boundary; keep all
  // blocks on a single page rather than looping or splitting on bad geometry —
  // EXCEPT for hard section-break boundaries, which always split.
  const hasForcedBreaks = breaksBefore !== undefined && breaksBefore.size > 0;
  if (
    (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) &&
    !hasForcedBreaks
  ) {
    return {
      pages: [blockHeights.map((_, index) => index)],
      effectiveBlockHeights,
      naturalPageStarts: [],
    };
  }
  const boundedHeight =
    Number.isFinite(pageContentHeightPx) && pageContentHeightPx > 0
      ? pageContentHeightPx
      : Number.POSITIVE_INFINITY;

  const pages: number[][] = [];
  const naturalPageStarts: number[] = [];
  let current: number[] = [];
  let used = 0;

  for (let index = 0; index < blockHeights.length; index += 1) {
    let height = effectiveBlockHeights[index] ?? 0;

    // A hard boundary (section break) forces this block onto a new page even
    // when it would otherwise fit; a forced break on a block that is already at
    // the top of a page is a no-op (no leading empty card).
    const forcedKind = hasForcedBreaks
      ? breakKindAt(breaksBefore, index)
      : null;
    if (forcedKind && current.length > 0) {
      pages.push(current);
      current = [];
      used = 0;
      appendParityBlankPage(pages, forcedKind);
    }
    // Start a new page when this block would overflow the remaining space on a
    // page that already holds at least one block. A block on an empty page stays
    // put even when it is taller than the page (taller-than-page overflow): it
    // cannot fit anywhere, so it gets its own page and overflows the bottom.
    // Zero-height control blocks (page/section break markers) always stay with
    // the preceding content, including after an oversized block; otherwise the
    // marker itself would manufacture an extra blank page before its directive.
    if (current.length > 0 && height > 0 && used + height > boundedHeight) {
      pages.push(current);
      current = [];
      used = 0;
      const rawFlowSpacing = flowSpacingBefore[index];
      const collapsibleFlowSpacing =
        Number.isFinite(rawFlowSpacing) && (rawFlowSpacing ?? 0) > 0
          ? Math.min(height, rawFlowSpacing ?? 0)
          : 0;
      height = Math.max(0, height - collapsibleFlowSpacing);
      effectiveBlockHeights[index] = height;
      naturalPageStarts.push(index);
    }

    current.push(index);
    used += height;
  }

  pages.push(current);

  // A terminal page/section break has no following block to trigger the normal
  // loop boundary. Preserve its semantics by creating the requested trailing
  // page (and, for even/odd, an intervening parity page when necessary).
  const terminalKind = breakKindAt(breaksBefore, blockHeights.length);
  if (terminalKind) {
    appendParityBlankPage(pages, terminalKind);
    pages.push([]);
  }
  return { pages, effectiveBlockHeights, naturalPageStarts };
}

export function paginateBlocks(
  blockHeights: number[],
  pageContentHeightPx: number,
  breaksBefore?: PaginationBreaks,
): number[][] {
  return paginateBlockLayout(blockHeights, pageContentHeightPx, breaksBefore)
    .pages;
}
