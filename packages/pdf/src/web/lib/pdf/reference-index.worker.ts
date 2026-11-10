/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Builds a PDF's reference index (lib/pdf/reference-index.ts) off the main
// thread: the whole document's text is folded and scanned there, a quarter of
// a second of the reader's thread on a 300-page book while it opened.
import { buildReferenceIndex, type PdfReferenceIndex, type PdfReferencePageInput } from "./reference-index";

type Request =
  | { id: number; page: PdfReferencePageInput | null }
  | { id: number; build: true };
type Response = { ready: true } | { id: number; index?: PdfReferenceIndex; error?: string };

const scope = self as unknown as { postMessage(value: Response): void; onmessage: ((event: MessageEvent<Request>) => void) | null };
const pagesById = new Map<number, Array<PdfReferencePageInput | null>>();

scope.postMessage({ ready: true });
scope.onmessage = (event) => {
  const request = event.data;
  if ("page" in request) {
    let pages = pagesById.get(request.id);
    if (!pages) pagesById.set(request.id, (pages = []));
    pages.push(request.page);
    return;
  }
  const pages = pagesById.get(request.id) ?? [];
  pagesById.delete(request.id);
  try {
    const index = buildReferenceIndex(pages);
    // The page needs a page's searchable text only where it draws mentions:
    // the rest stays here instead of crossing to the reader's thread.
    const entries = index.entries.map((entry, pageIndex) => (index.mentionsByPage.has(pageIndex) ? entry : null));
    scope.postMessage({ id: request.id, index: { ...index, entries } });
  } catch (error) {
    scope.postMessage({ id: request.id, error: error instanceof Error ? error.message : String(error) });
  }
};
