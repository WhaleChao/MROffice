/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function createPdfStagedTextReads<Content>() {
  const reads = new WeakMap<object, Map<number, Promise<Content>>>();
  return {
    /** Starts the read of each page not read ahead yet. */
    start(document: object, pageNumbers: readonly number[], read: (pageNumber: number) => Promise<Content>): void {
      let ofDocument = reads.get(document);
      if (!ofDocument) { ofDocument = new Map(); reads.set(document, ofDocument); }
      for (const pageNumber of pageNumbers) {
        if (!Number.isSafeInteger(pageNumber) || pageNumber < 1 || ofDocument.has(pageNumber)) continue;
        const pending = Promise.resolve().then(() => read(pageNumber));
        // A revision given up before its read ended (the reader typed on).
        pending.catch(() => undefined);
        ofDocument.set(pageNumber, pending);
      }
    },
    /** The page's read, when one was started and nothing took it yet. */
    take(document: object, pageNumber: number): Promise<Content> | null {
      const ofDocument = reads.get(document);
      const pending = ofDocument?.get(pageNumber) ?? null;
      if (pending) ofDocument!.delete(pageNumber);
      return pending;
    },
  };
}
