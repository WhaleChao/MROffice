/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { resolvePageItems, type PageMove } from "./page-ops";

/**
 * A page inserted in MR Office (a blank page, a copy) has no number in the
 * upload, which every page-keyed record names its page by. An annotation or a
 * comment made on it names it by the inserted page's own id
 * (`insertedPageId`), beside `page`, which then holds the page it was
 * inserted after (at least 1): a reader that does not know the field shows it
 * on the page before, never loses it. The writer (pdf_row_plan, the page
 * save) puts it on the inserted page wherever that page now is in the file;
 * deleted with its page, it goes with it -- as Acrobat annotates a page added
 * through Organize Pages at once, and keeps the markup on that page.
 */

/** The reader's own number for an inserted page: above every upload page
 * (the reader keys its layers by number). Never stored. */
export const PDF_INSERTED_READER_PAGE_BASE = 1_000_000;

const INSERTED_PAGE_ID_PATTERN = /^[A-Za-z0-9._:-]{1,200}$/;

/** The inserted page a stored record names, or null. */
export function sanitizeInsertedPageId(value: unknown): string | null {
  return typeof value === "string" && INSERTED_PAGE_ID_PATTERN.test(value) ? value : null;
}

/** Whether `page` is a reader number of an inserted page. */
export function isInsertedReaderPage(page: number): boolean {
  return Number.isInteger(page) && page >= PDF_INSERTED_READER_PAGE_BASE;
}

/** The page a record on the inserted page `afterPage` follows keeps in its
 * `page` (1-based): that page, or the first for a page leading the file. */
export function insertedPageAnchorPage(afterPage: number): number {
  return Number.isInteger(afterPage) && afterPage >= 1 ? afterPage : 1;
}

export interface PdfInsertedReaderPages {
  /** The reader number of each inserted page shown, by its id. */
  readonly readerPageOf: ReadonlyMap<string, number>;
  /** What each reader number stands for. */
  readonly insertedAt: ReadonlyMap<number, Readonly<{ id: string; afterPage: number }>>;
}

/** Reader numbers for the inserted pages, in their order (held imports,
 * which have no lasting id before Save, are not given one). */
export function pdfInsertedReaderPages(
  inserted: readonly Readonly<{ id: string; afterPage: number }>[],
): PdfInsertedReaderPages {
  const readerPageOf = new Map<string, number>();
  const insertedAt = new Map<number, { id: string; afterPage: number }>();
  for (const page of inserted) {
    if (!sanitizeInsertedPageId(page.id) || page.id.startsWith("held-import-") || readerPageOf.has(page.id)) continue;
    const number = PDF_INSERTED_READER_PAGE_BASE + readerPageOf.size;
    readerPageOf.set(page.id, number);
    insertedAt.set(number, { id: page.id, afterPage: page.afterPage });
  }
  return { readerPageOf, insertedAt };
}

/** The reader number a record is drawn under: its inserted page's, its own
 * page otherwise; null when it names an inserted page no longer there. */
export function readerPageOfRecord(
  record: Readonly<{ page: number; insertedPageId?: string }>,
  pages: PdfInsertedReaderPages,
): number | null {
  if (!record.insertedPageId) return record.page;
  return pages.readerPageOf.get(record.insertedPageId) ?? null;
}

/** A reader number as stored: an inserted page's as the page before it and
 * its id; any other as itself. Null for a reader number no page has. */
export function storedPageOfReaderPage(
  page: number,
  pages: PdfInsertedReaderPages,
): { page: number; insertedPageId?: string } | null {
  if (!isInsertedReaderPage(page)) return { page };
  const inserted = pages.insertedAt.get(page);
  return inserted ? { page: insertedPageAnchorPage(inserted.afterPage), insertedPageId: inserted.id } : null;
}

/** The 1-based place in the reading order of an upload page or an inserted
 * page under the page's structure (deleted, moved, added), as the reader
 * shows it: what "Page N" means to the person reading. */
export function pdfVisiblePageOrdinals(
  structure: Readonly<{
    deletedPages: readonly number[];
    pageMoves: readonly PageMove[];
    insertedPages: readonly Readonly<{ id: string; afterPage: number; width?: number; height?: number }>[];
  }>,
  originalPageCount: number,
): { ofOriginal: ReadonlyMap<number, number>; ofInserted: ReadonlyMap<string, number> } {
  const ofOriginal = new Map<number, number>();
  const ofInserted = new Map<string, number>();
  if (!Number.isInteger(originalPageCount) || originalPageCount < 1) return { ofOriginal, ofInserted };
  const items = resolvePageItems(
    {
      rotations: {},
      deletedPages: [...structure.deletedPages],
      pageMoves: [...structure.pageMoves],
      insertedPages: structure.insertedPages.map((page) => ({
        id: page.id, afterPage: page.afterPage, width: page.width ?? 612, height: page.height ?? 792,
      })),
    },
    Array.from({ length: originalPageCount }, (_, index) => index + 1),
  );
  items.forEach((item, index) => {
    if (item.kind === "original") ofOriginal.set(item.page, index + 1);
    else if (item.page.id) ofInserted.set(item.page.id, index + 1);
  });
  return { ofOriginal, ofInserted };
}

/** Strokes stored by 0-based page, each one on an inserted page put under
 * that page's reader number (0-based too); one whose page is gone is left
 * out (it went with the page). */
export function onInsertedReaderPages<T extends { page: number; insertedPageId?: unknown }>(
  strokes: readonly T[] | undefined,
  pages: PdfInsertedReaderPages,
): T[] | undefined {
  if (!strokes) return undefined;
  return strokes.flatMap((stroke) => {
    if (typeof stroke.insertedPageId !== "string") return [stroke];
    const reader = pages.readerPageOf.get(stroke.insertedPageId);
    return reader === undefined ? [] : [{ ...stroke, page: reader - 1 }];
  });
}
