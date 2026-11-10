/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Page texts joined the way the pages follow each other; empty pages add nothing. */
export function joinDocumentPageTexts(pages: readonly string[]): string {
  return pages.filter((text) => text.trim() !== "").join("\n");
}

/**
 * The whole document's text when every page has already been read, else
 * `null`: a copy event can only be answered synchronously with text that is
 * already there.
 */
export function documentTextIfReady(
  pageNumbers: readonly number[],
  cached: (pageNumber: number) => string | undefined,
): string | null {
  const texts: string[] = [];
  for (const pageNumber of pageNumbers) {
    const text = cached(pageNumber);
    if (text === undefined) return null;
    texts.push(text);
  }
  return joinDocumentPageTexts(texts);
}

/**
 * Reads every page and resolves with the whole document's text. A page that
 * cannot be read rejects: part of a document is never delivered as all of it.
 */
export async function readDocumentText(
  pageNumbers: readonly number[],
  readPage: (pageNumber: number) => Promise<string | null>,
): Promise<string> {
  const texts: string[] = [];
  for (const pageNumber of pageNumbers) {
    const text = await readPage(pageNumber);
    if (text === null) {
      throw new Error(`pdf_document_text_page_unreadable:${pageNumber}`);
    }
    texts.push(text);
  }
  return joinDocumentPageTexts(texts);
}
