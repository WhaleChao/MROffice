/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { resolvePageItems, type PageMove } from "@mrnote/core/web/lib/pdf/page-ops.ts";








export interface PdfDocumentPageStructure {
  /** The file this structure is written in (the saved export's digest). */
  readonly sourceSha256: string;
  readonly deletedPages: readonly number[];
  readonly pageMoves: readonly PageMove[];
  /** Pages added, as the page keeps them (their size is not needed here). */
  readonly insertedPages: readonly Readonly<{ id: string; afterPage: number }>[];
}

export interface PdfDocumentPageMap {
  readonly sourceSha256: string;
  /** For each page of the file, in order: the upload's page it is (its
   * number there), or null for a page added to it. */
  readonly documentPages: readonly (number | null)[];
  readonly originalPageCount: number;
}

/** The map of a file holding `structure`, whose page count is
 * `documentPageCount`; null when the structure does not account for that
 * many pages (the file is not the one the structure was written in). */
export function pdfDocumentPageMap(
  structure: PdfDocumentPageStructure,
  documentPageCount: number,
): PdfDocumentPageMap | null {
  const deleted = new Set(structure.deletedPages.filter((page) => Number.isInteger(page) && page >= 1));
  const originalPageCount = documentPageCount + deleted.size - structure.insertedPages.length;
  if (!Number.isInteger(documentPageCount) || documentPageCount < 1 || originalPageCount < 1 ||
    [...deleted].some((page) => page > originalPageCount)) return null;
  const items = resolvePageItems(
    { rotations: {}, deletedPages: [...deleted], pageMoves: [...structure.pageMoves],
      insertedPages: structure.insertedPages.map(({ id, afterPage }) => ({ id, afterPage, width: 0, height: 0 })) },
    Array.from({ length: originalPageCount }, (_, index) => index + 1),
  );
  if (items.length !== documentPageCount) return null;
  return Object.freeze({
    sourceSha256: structure.sourceSha256,
    documentPages: Object.freeze(items.map((item) => (item.kind === "original" ? item.page : null))),
    originalPageCount,
  });
}

/** The file's 1-based page showing the upload's page `original`; null for a
 * page deleted from it. */
export function documentPageForOriginal(map: PdfDocumentPageMap, original: number): number | null {
  const index = map.documentPages.indexOf(original);
  return index < 0 ? null : index + 1;
}

/** The upload's page the file's 1-based page `documentPage` is; null for an
 * added page. */
export function originalForDocumentPage(map: PdfDocumentPageMap, documentPage: number): number | null {
  return map.documentPages[documentPage - 1] ?? null;
}

/** Whether the page's structure now (`edits`) is the one the file holds:
 * a page operation not saved yet is drawn over the file but not in it. */
export function pageStructureIsHeld(
  structure: Pick<PdfDocumentPageStructure, "deletedPages" | "pageMoves" | "insertedPages">,
  edits: Readonly<{
    deletedPages: readonly number[];
    pageMoves: readonly PageMove[];
    insertedPages: readonly Readonly<{ id: string; afterPage: number }>[];
  }>,
): boolean {
  const deleted = (pages: readonly number[]) => [...pages].sort((a, b) => a - b).join();
  const moves = (list: readonly PageMove[]) => list.map((move) => `${move.page}>${move.afterPage}`).join();
  const inserted = (list: readonly Readonly<{ id: string; afterPage: number }>[]) =>
    list.map((page) => `${page.id}@${page.afterPage}`).join();
  return deleted(structure.deletedPages) === deleted(edits.deletedPages) &&
    moves(structure.pageMoves) === moves(edits.pageMoves) &&
    inserted(structure.insertedPages) === inserted(edits.insertedPages);
}

/** Asked for a page the file no longer holds (deleted from it). */
export class PdfStructuredDeletedPageError extends Error {
  constructor(readonly originalPage: number) {
    super(`Page ${originalPage} is deleted from the shown file.`);
    this.name = "PdfStructuredDeletedPageError";
  }
}

interface PdfDocumentLike {
  numPages: number;
  getPage(pageNumber: number): Promise<unknown>;
  getPageIndex?(ref: unknown): Promise<number>;
  getPageLabels?(): Promise<string[] | null>;
}

const STRUCTURED_MAP = Symbol("mrai.pdf.structuredMap");

/** The map a wrapped document was made with; null for a plain one. */
export function structuredDocumentMap(document: unknown): PdfDocumentPageMap | null {
  return (document && typeof document === "object"
    ? (document as Record<symbol, PdfDocumentPageMap | undefined>)[STRUCTURED_MAP] : undefined) ?? null;
}

/**
 * The file as the reader reads the upload: its pages numbered as uploaded
 * (`numPages` the upload's count, `getPage(n)` the file's page showing the
 * upload's page n, with that number; a deleted page throws
 * PdfStructuredDeletedPageError), destinations and labels mapped back. Every
 * other member is the file's own.
 */
export function wrapStructuredPdfDocument<T extends PdfDocumentLike>(document: T, map: PdfDocumentPageMap): T {
  const pages = new Map<number, Promise<unknown>>();
  const pageProxy = (page: unknown, original: number) =>
    page && typeof page === "object"
      ? new Proxy(page as object, {
          get(target, key, receiver) {
            if (key === "pageNumber") return original;
            const value = Reflect.get(target, key, receiver);
            return typeof value === "function" ? value.bind(target) : value;
          },
        })
      : page;
  return new Proxy(document as object, {
    get(target, key, receiver) {
      if (key === STRUCTURED_MAP) return map;
      if (key === "numPages") return map.originalPageCount;
      if (key === "getPage") {
        return (pageNumber: number) => {
          const documentPage = documentPageForOriginal(map, pageNumber);
          if (documentPage === null) return Promise.reject(new PdfStructuredDeletedPageError(pageNumber));
          let page = pages.get(pageNumber);
          if (!page) {
            page = (target as T).getPage(documentPage).then((value) => pageProxy(value, pageNumber));
            page.catch(() => pages.delete(pageNumber));
            pages.set(pageNumber, page);
          }
          return page;
        };
      }
      if (key === "getPageIndex" && typeof (target as T).getPageIndex === "function") {
        return async (ref: unknown) => {
          const index = await (target as T).getPageIndex!(ref);
          const original = originalForDocumentPage(map, index + 1);
          // A destination on a page added to the file names no upload page.
          if (original === null) throw new Error("The destination is on an added page.");
          return original - 1;
        };
      }
      if (key === "getPageLabels" && typeof (target as T).getPageLabels === "function") {
        return async () => {
          const labels = await (target as T).getPageLabels!();
          if (!labels) return labels;
          return Array.from({ length: map.originalPageCount }, (_, index) => {
            const documentPage = documentPageForOriginal(map, index + 1);
            return documentPage === null ? "" : labels[documentPage - 1] ?? "";
          });
        };
      }
      const value = Reflect.get(target, key, receiver);
      return typeof value === "function" ? value.bind(target) : value;
    },
  }) as T;
}

const onFilePages = new WeakMap<PdfDocumentPageMap, WeakMap<object, object>>();
const shownOf = new WeakMap<object, object>();

function filePagesOf(map: PdfDocumentPageMap): WeakMap<object, object> {
  let converted = onFilePages.get(map);
  if (!converted) onFilePages.set(map, (converted = new WeakMap()));
  return converted;
}

/**
 * A selection the reader made (its `pageNumber` the upload's) as the page's
 * edits address it: on the shown file's page, which is where the engine and
 * every byte read find it. The same selection gives the same object back;
 * one already on the file's pages (or handed to the reader from them by
 * selectionOnShownPages) gives its file selection. Null for a page the file
 * no longer holds. Without a map (the upload shown) it is unchanged.
 */
export function selectionOnFilePages<T extends Readonly<{ pageNumber: number }>>(
  map: PdfDocumentPageMap | null,
  selection: T,
): T | null {
  if (!map || shownOf.has(selection)) return selection;
  const converted = filePagesOf(map);
  const known = converted.get(selection);
  if (known) return known as T;
  const page = documentPageForOriginal(map, selection.pageNumber);
  if (page === null) return null;
  const onFile = Object.freeze({ ...selection, pageNumber: page }) as T;
  converted.set(selection, onFile);
  shownOf.set(onFile, selection);
  return onFile;
}

/** A selection on the file's pages (every selection the page's edits hold)
 * as handed to the reader: on the upload's page, the reader's numbering. The
 * reader's own selection comes back as it was; another is mapped, and given
 * back to selectionOnFilePages it returns this one. */
export function selectionOnShownPages<T extends Readonly<{ pageNumber: number }>>(
  map: PdfDocumentPageMap | null,
  selection: T,
): T {
  if (!map) return selection;
  const known = shownOf.get(selection) as T | undefined;
  if (known) return known;
  const page = originalForDocumentPage(map, selection.pageNumber);
  if (page === null) return selection;
  const shown = Object.freeze({ ...selection, pageNumber: page }) as T;
  filePagesOf(map).set(shown, selection);
  shownOf.set(selection, shown);
  return shown;
}

/** Maps every `pageNumber` of a plain value (objects and arrays, nested a
 * few levels; bytes and other instances are left as they are). Null when a
 * page has no counterpart. */
function mapPageNumbers<T>(value: T, page: (pageNumber: number) => number | null, depth = 0): T | null {
  if (depth > 6 || value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) {
    const mapped: unknown[] = [];
    for (const entry of value) {
      const next = mapPageNumbers(entry, page, depth + 1);
      if (next === null && entry !== null) return null;
      mapped.push(next);
    }
    return Object.freeze(mapped) as T;
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return value;
  const mapped: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
    if (key === "pageNumber" && typeof entry === "number") {
      const next = page(entry);
      if (next === null) return null;
      mapped[key] = next;
    } else if (typeof entry === "object" && entry !== null) {
      const next = mapPageNumbers(entry, page, depth + 1);
      if (next === null) return null;
      mapped[key] = next;
    } else {
      mapped[key] = entry;
    }
  }
  return (Object.isFrozen(value) ? Object.freeze(mapped) : mapped) as T;
}

/** A request the reader made (every `pageNumber` the upload's) on the shown
 * file's pages; null when it names a page deleted from the file. Unchanged
 * without a map. */
export function requestOnFilePages<T>(map: PdfDocumentPageMap | null, request: T): T | null {
  return map ? mapPageNumbers(request, (page) => documentPageForOriginal(map, page)) : request;
}

/** What the page's edits hand the reader (every `pageNumber` the file's) on
 * the reader's pages; a page added to the file keeps its number. */
export function outcomeOnShownPages<T>(map: PdfDocumentPageMap | null, outcome: T): T {
  return map ? mapPageNumbers(outcome, (page) => originalForDocumentPage(map, page) ?? page) ?? outcome : outcome;
}
