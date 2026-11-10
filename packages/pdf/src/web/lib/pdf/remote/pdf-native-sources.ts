/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which page and stored version each working head the editor holds belongs
 * to, by its SHA-256 and then its page. A planner call
 * (`pdf-native-planners.ts`) finds its source among its byte arguments by
 * digest here, asking with the page its planner set is bound to
 * (`createPdfNativePlanners(pageId)`), so callers keep the planners'
 * signatures: whoever loads a head for editing registers it (the pinned
 * source authority, a working session's next head).
 */

import type { PdfNativeEngineSourceRef } from "../pdf-native-remote";

/** A resource's original upload, when the editor does not know its version
 * id (an unpatched page): the server resolves it to the version itself. */
export const PDF_NATIVE_ORIGINAL_VERSION = "original";

const LIMIT = 64;






const heads = new Map<string, Map<string, PdfNativeEngineSourceRef>>();

function headCount(): number {
  let count = 0;
  for (const pages of heads.values()) count += pages.size;
  return count;
}

export function registerPdfNativeWorkingHead(sha256: string, ref: PdfNativeEngineSourceRef): void {
  const key = sha256.toLowerCase();
  const pages = heads.get(key) ?? new Map<string, PdfNativeEngineSourceRef>();
  heads.delete(key);
  pages.delete(ref.pageId);
  pages.set(ref.pageId, Object.freeze({ pageId: ref.pageId, resourceId: ref.resourceId, resourceVersionId: ref.resourceVersionId }));
  heads.set(key, pages);
  while (headCount() > LIMIT) {
    const [oldestKey, oldest] = heads.entries().next().value as [string, Map<string, PdfNativeEngineSourceRef>];
    oldest.delete(oldest.keys().next().value as string);
    if (!oldest.size) heads.delete(oldestKey);
  }
}

/** The pages that registered these bytes, most recent last. */
export function pdfNativeWorkingHeadPages(sha256: string): readonly string[] {
  return [...(heads.get(sha256.toLowerCase())?.keys() ?? [])];
}

/** The page and version these bytes are held under. A window asks with its
 * own page (its planner set, a writer source, a line edit's engine source, a
 * moved head, a hold): it is answered only with its own registration, never
 * another page's that holds the same bytes.
 *
 * Asked without a page, only bytes one page holds are answered: bytes two
 * pages registered (one PDF uploaded twice, both open in one tab) are
 * refused (null), as no caller can say which of the two it meant -- it
 * once took the last one registered, and a write of one window could name
 * the other's resource. `pdfNativeWorkingHeadPages` tells the two apart. */
export function pdfNativeWorkingHead(sha256: string, pageId?: string | null): PdfNativeEngineSourceRef | null {
  const pages = heads.get(sha256.toLowerCase());
  if (!pages) return null;
  if (pageId) return pages.get(pageId) ?? null;
  if (pages.size !== 1) return null;
  return pages.values().next().value ?? null;
}

/** Heads assembled here that the server was never sent: a text edit's
 * candidate comes back from `prepare`, which keeps no candidate, so the first
 * call naming the next head was always answered `blob_unknown` and asked
 * again with the bytes (a wasted round trip after every saved line, 0.5-0.8
 * s). Such a head goes with its first call; once sent, the server keeps it. */
const unheld = new Map<string, PdfNativeHeadBase | null>();

/** The head a new head is an incremental update of: its digest and length
 * (the new head is those bytes with the update appended). */
export type PdfNativeHeadBase = Readonly<{ sha256: string; length: number }>;

/** `base`: what the head appends to, when it does. Its first call then
 * sends only the appended bytes (the server usually holds the base: the
 * saved version, or the head before, which a call already sent); a whole
 * head went with every saved line before, the size of the file. */
export function notePdfNativeHeadUnheld(sha256: string, base: PdfNativeHeadBase | null = null): void {
  const key = sha256.toLowerCase();
  unheld.delete(key);
  unheld.set(key, base && Object.freeze({ sha256: base.sha256.toLowerCase(), length: base.length }));
  while (unheld.size > LIMIT) unheld.delete(unheld.keys().next().value as string);
}

export function notePdfNativeHeadHeld(sha256: string): void {
  unheld.delete(sha256.toLowerCase());
}

export function pdfNativeHeadBase(sha256: string): PdfNativeHeadBase | null {
  return unheld.get(sha256.toLowerCase()) ?? null;
}

export function pdfNativeHeadUnheld(sha256: string): boolean {
  return unheld.has(sha256.toLowerCase());
}

/** A PDF shown outside an editable page (a reference document, a
 * projection's source column): the server authorizes it as its plain
 * download and runs only read-only inspections on it (`READER_FUNCTIONS`). */
export const PDF_NATIVE_READER_PAGE = "reader";

/** Name a reader's shown head by its resource, under the reader's own page
 * (`createPdfNativePlanners(PDF_NATIVE_READER_PAGE)` asks with it): an
 * editing page holding the same bytes keeps its own name beside it. */
export function registerPdfNativeReaderHead(sha256: string, resourceId: string): void {
  if (pdfNativeWorkingHead(sha256, PDF_NATIVE_READER_PAGE)?.resourceId === resourceId) return;
  registerPdfNativeWorkingHead(sha256, {
    pageId: PDF_NATIVE_READER_PAGE, resourceId, resourceVersionId: PDF_NATIVE_ORIGINAL_VERSION,
  });
}

/** Tests. */
export function clearPdfNativeWorkingHeads(): void {
  heads.clear();
  unheld.clear();
}
