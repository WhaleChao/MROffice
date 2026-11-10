/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whole-document find & replace over the editable PDF text-box model.
 *
 * `text-box-replace` already replaces across a list of boxes; what it cannot do
 * is reach a page the editor has never rendered. A textbook opens on page 1 and
 * "replace all" quietly means "replace all on this page", which is the shape of
 * under-delivery the roadmap flagged: the hard part (masking, origin rects,
 * formula latex) was built, and the scope was left at one page.
 *
 * These helpers walk the document one page at a time through a caller-supplied
 * `boxesForPage`, so the editor decides how a page's boxes are obtained — the
 * mounted page's live boxes when it has them, a headless extraction when it does
 * not — and this module stays free of React and pdf.js.
 *
 * Pages are visited in order and awaited one at a time. That is deliberate: a
 * 600-page book extracted concurrently would hold every page's text content in
 * memory at once, and the user is waiting on a progress count either way.
 */

import {
  applyReplaceAll,
  findBoxMatches,
  type ReplaceBox,
  type ReplaceBoxWithPage,
  type ReplaceEdit,
} from "./text-box-replace";
import type { PdfTextSearchOptions } from "@mrnote/core/web/lib/pdf/text-search.ts";

/** How a page's boxes are obtained. Returning `null` skips the page — the
 *  editor uses that for a page whose extraction failed, so one unreadable page
 *  cannot abort a whole-document replace. */
export type BoxesForPage<B extends ReplaceBox> = (
  page: number,
) => Promise<readonly B[] | null> | readonly B[] | null;

export interface DocumentScanParams<B extends ReplaceBox> {
  /** How many pages to visit. */
  pageCount: number;
  /**
   * Index of the first page. PDF pages are conventionally 1-based, but the
   * inline editor keys its boxes and edits 0-based, and silently converting in
   * the caller is how an off-by-one gets shipped — so the base is explicit and
   * whatever the caller passes is what `boxesForPage` and `editFor` receive.
   */
  firstPage?: number;
  boxesForPage: BoxesForPage<B>;
  query: string;
  options?: PdfTextSearchOptions;
  /** Called after each page is visited, for a progress indicator. */
  /**
   * The text a box currently shows when it differs from the extraction --
   * its saved or pending edit -- so a scan counts what the reader sees.
   */
  textFor?: (box: B, page: number) => string | undefined;
  onPageVisited?: (page: number, pageCount: number) => void;
  /** Polled between pages so a closing find bar can stop a long scan. */
  isCancelled?: () => boolean;
}

export interface DocumentMatchCount {
  /** Total occurrences across every visited page. */
  matches: number;
  /** How many pages hold at least one occurrence. */
  pages: number;
  /** Pages actually visited — lower than `pageCount` when cancelled. */
  visited: number;
  /** True when a cancel request ended the walk early. */
  cancelled: boolean;
}

/** Count occurrences of `query` across the whole document. */
export async function countMatchesAcrossDocument<B extends ReplaceBox>(
  params: DocumentScanParams<B>,
): Promise<DocumentMatchCount> {
  const {
    pageCount,
    firstPage = 1,
    boxesForPage,
    query,
    options,
    onPageVisited,
    isCancelled,
  } = params;
  let matches = 0;
  let pages = 0;
  let visited = 0;
  if (!query.trim()) return { matches: 0, pages: 0, visited: 0, cancelled: false };
  const lastPage = firstPage + pageCount - 1;
  for (let page = firstPage; page <= lastPage; page += 1) {
    if (isCancelled?.()) return { matches, pages, visited, cancelled: true };
    const boxes = await boxesForPage(page);
    visited += 1;
    onPageVisited?.(page, pageCount);
    if (!boxes || boxes.length === 0) continue;
    const found = findBoxMatches(
      withPage(boxes, page, params.textFor),
      query,
      options,
    );
    if (found.length > 0) {
      matches += found.length;
      pages += 1;
    }
  }
  return { matches, pages, visited, cancelled: false };
}

export interface DocumentReplaceResult<E extends ReplaceEdit> {
  /** One entry per box that actually changed, across every page. */
  changedEdits: E[];
  /** Total occurrences replaced. */
  replaced: number;
  /** How many pages had at least one replacement. */
  pages: number;
  /** Occurrences found in boxes `canReplace` refused: counted, never
   *  replaced, so the result line can say what was left alone. */
  skipped: number;
  /** Occurrences inside a formula its LaTeX cannot follow: left, the
   *  formula whole (`applyReplaceAll`). */
  formulaSkipped: number;
  visited: number;
  cancelled: boolean;
}

export interface DocumentReplaceParams<B extends ReplaceBox, E extends ReplaceEdit>
  extends DocumentScanParams<B> {
  replacement: string;
  /**
   * The existing edit for a box, or a freshly seeded one carrying that box's
   * geometry. It receives the BOX (not just its id) because a page the editor
   * has not rendered has no entry in its by-id map — the geometry has to come
   * from the scan itself, or the replacement would have no origin rect and the
   * original glyphs would show through on export.
   */
  editFor: (box: B, page: number) => E | undefined;
  





  canReplace?: (box: B, page: number) => boolean;
}

/** Replace every occurrence of `query` across the whole document. */
export async function replaceAcrossDocument<
  B extends ReplaceBox,
  E extends ReplaceEdit,
>(params: DocumentReplaceParams<B, E>): Promise<DocumentReplaceResult<E>> {
  const {
    pageCount,
    firstPage = 1,
    boxesForPage,
    query,
    replacement,
    options,
    editFor,
    canReplace,
    onPageVisited,
    isCancelled,
  } = params;
  const changedEdits: E[] = [];
  let replaced = 0;
  let pages = 0;
  let skipped = 0;
  let formulaSkipped = 0;
  let visited = 0;
  if (!query.trim()) {
    return { changedEdits, replaced: 0, pages: 0, skipped: 0, formulaSkipped: 0, visited: 0, cancelled: false };
  }
  const lastPage = firstPage + pageCount - 1;
  for (let page = firstPage; page <= lastPage; page += 1) {
    if (isCancelled?.()) {
      return { changedEdits, replaced, pages, skipped, formulaSkipped, visited, cancelled: true };
    }
    const scanned = await boxesForPage(page);
    visited += 1;
    onPageVisited?.(page, pageCount);
    if (!scanned || scanned.length === 0) continue;
    let boxes: readonly B[] = scanned;
    if (canReplace) {
      const kept = scanned.filter((box) => canReplace(box, page));
      const refused = scanned.filter((box) => !kept.includes(box));
      if (refused.length > 0) {
        skipped += findBoxMatches(withPage(refused, page), query, options).length;
      }
      boxes = kept;
      if (boxes.length === 0) continue;
    }
    const byId = new Map(boxes.map((box) => [box.id, box]));
    const result = applyReplaceAll<E>({
      boxes: withPage(boxes, page),
      query,
      replacement,
      options,
      editFor: (boxId) => {
        const box = byId.get(boxId);
        return box ? editFor(box, page) : undefined;
      },
    });
    formulaSkipped += result.formulaSkipped;
    if (result.replaced === 0) continue;
    replaced += result.replaced;
    pages += 1;
    changedEdits.push(...result.changedEdits);
  }
  return { changedEdits, replaced, pages, skipped, formulaSkipped, visited, cancelled: false };
}

function withPage<B extends ReplaceBox>(
  boxes: readonly B[],
  page: number,
  textFor?: (box: B, page: number) => string | undefined,
): ReplaceBoxWithPage[] {
  return boxes.map((box) => {
    const edited = textFor?.(box, page);
    return {
      id: box.id,
      text: typeof edited === "string" ? edited : box.text,
      kind: box.kind,
      latex: box.latex,
      page,
    };
  });
}
