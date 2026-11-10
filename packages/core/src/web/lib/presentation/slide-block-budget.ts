/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How many blocks one SLIDE may carry: a share of the page, not a count.
 *
 * The cap used to be `MAX_PRESENTATION_BLOCKS_PER_SLIDE = 200`, written out
 * four times — the page schema, the page validator, the server projection and
 * this editor's own insert gate. It was the same mistake the workbook's 256
 * drawing objects were (docs/spreadsheet-page-limits.md, "Drawing objects: a
 * share of the page, not a count"), and the measurement says so just as
 * plainly: on the local 132-package corpus a slide block costs from ~250 bytes
 * (a connector) to 283 KB (a picture), mean 6.1 KB, median 1.8 KB. 200 blocks
 * is therefore half a percent of a presentation page as connectors and several
 * times the whole page as pictures, and one number cannot be right for both.
 * On that corpus the count cut six slides of one deck in silence, the worst of
 * them from 519 shapes to 200.
 *
 * What binds now is the bytes a slide's blocks put on the page. The count
 * survives only so a hand-written or runaway body cannot drive an unbounded
 * per-block loop, and it is the same 10,000 the workbook uses. The whole-page
 * envelope (400,000 JSON nodes / 10 MiB for a presentation) still has the last
 * word on save, and the server projection's own ceiling sits deliberately
 * ABOVE this one so a slide past it is REPORTED by the page rather than cut in
 * silence by the projection.
 *
 * The rule and the measurement are shared with the workbook in
 * `lib/page-object-budget`; the server twins are
 * `app/services/resources/object_budget.py` and `_MAX_PPTX_SLIDE_BLOCK_BYTES`.
 */

import {
  pageObjectBudgetSpent,
  pageObjectBytes,
  pageObjectsWithinBudget,
  type PageObjectBudgetLimits,
} from "../page-object-budget";

/** A bound on work, not the bound that binds. Mirrors the server validator. */
export const MAX_PRESENTATION_SLIDE_BLOCKS = 10_000;
/**
 * A slide's share of the page's 10 MiB. The heaviest real slide measured is
 * 1.03 MB, so every slide in the corpus arrives whole with room to spare,
 * while one synthetic slide still cannot eat the whole deck's page by itself.
 */
export const PRESENTATION_SLIDE_BLOCK_BYTE_BUDGET = 4 * 1024 * 1024;

const PRESENTATION_SLIDE_BLOCK_LIMITS: PageObjectBudgetLimits = {
  maxObjects: MAX_PRESENTATION_SLIDE_BLOCKS,
  byteBudget: PRESENTATION_SLIDE_BLOCK_BYTE_BUDGET,
};

/** Bytes the slide's blocks put on the page, measured once per block. */
export function presentationSlideBlockBytes(
  blocks: readonly unknown[],
): number {
  return pageObjectBytes(blocks, PRESENTATION_SLIDE_BLOCK_LIMITS);
}

/**
 * Whether this slide has used its share: no further block is accepted.
 *
 * Asked BEFORE a block is added, so a slide never holds a block it could not
 * afford. Every editor gate that used to read `blocks.length >= 200` asks this
 * instead.
 */
export function presentationSlideBlockBudgetSpent(
  blocks: readonly unknown[],
): boolean {
  return pageObjectBudgetSpent(blocks, PRESENTATION_SLIDE_BLOCK_LIMITS);
}

/** The longest prefix of `blocks` that fits one slide's share. */
export function presentationSlideBlocksWithinBudget<T>(
  blocks: readonly T[],
): T[] {
  return pageObjectsWithinBudget(blocks, PRESENTATION_SLIDE_BLOCK_LIMITS);
}
