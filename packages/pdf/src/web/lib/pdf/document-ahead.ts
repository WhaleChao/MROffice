"use client";



















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PDFDocumentLoadingTask, PDFDocumentProxy, PDFWorker } from "pdfjs-dist";

// (Safari has no async iteration of a ReadableStream, which pdf.js's
// getTextContent needs; the reader installs the same before it reads text.)
import "@mrnote/core/web/lib/polyfills/readable-stream-async-iterator.ts";

import { bindPdfWorkerToTask, takePdfReaderWorker } from "./pdfjs-reader-worker";
import { loadCompatiblePdfJs, PDF_JS_DOCUMENT_ASSETS } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";

/** How long a document opened ahead waits for its reader. */
const TAKE_WITHIN_MS = 10_000;
/** Larger files are opened by their reader alone: opening ahead holds two
 * more copies of the bytes for a moment. */
export const PDF_DOCUMENT_AHEAD_MAX_BYTES = 32 * 1024 * 1024;

export interface PdfDocumentAhead {
  readonly task: PDFDocumentLoadingTask;
  readonly worker: PDFWorker;
}

interface Entry {
  /** The bytes the document was opened from (pdf.js was handed a copy). */
  readonly bytes: Uint8Array;
  readonly opened: Promise<PdfDocumentAhead | null>;
  timer: ReturnType<typeof setTimeout> | null;
}

const ahead = new Map<string, Entry>();

function destroy(opened: Promise<PdfDocumentAhead | null>): void {
  void opened.then((document) => document?.task.destroy()).catch(() => undefined);
}

/** Zero the bytes an entry kept, once pdf.js has been handed its copy of
 * them (a reader may take the entry while its worker is still starting). */
function forget(entry: Entry): void {
  void entry.opened.finally(() => entry.bytes.fill(0));
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  let index = 0;
  // Four bytes at a step where both are word-aligned (as every copy made by
  // `slice` is): a 30 MB file is compared in a few milliseconds.
  if (left.byteOffset % 4 === 0 && right.byteOffset % 4 === 0) {
    const words = left.byteLength >>> 2;
    const leftWords = new Uint32Array(left.buffer, left.byteOffset, words);
    const rightWords = new Uint32Array(right.buffer, right.byteOffset, words);
    for (let word = 0; word < words; word += 1) {
      if (leftWords[word] !== rightWords[word]) return false;
    }
    index = words * 4;
  }
  for (; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

/**
 * The first page's fonts and operators, read on the document's worker while
 * the reader is still mounting.
 *
 * The worker had the document a quarter of a second before the reader asked
 * it for anything, and then read the first page's fonts with the reader
 * waiting to draw it. A page that holds text is asked for its operator list
 * here, exactly as the reader asks for it once it has the page's text (the
 * glyph advances its editor needs, PDFViewer's text layer): the fonts are
 * read and handed to the page's thread before the reader draws, and the
 * reader's own request for that list is this one. A page with no text (a
 * scan) is not asked: nothing reads its operator list but the drawing, and
 * its pictures would be decoded twice.
 */
function readFirstPageAhead(document: PDFDocumentProxy): void {
  if ((globalThis as { __pdfFirstPageAheadOff?: boolean }).__pdfFirstPageAheadOff === true) return;
  void (async () => {
    const page = await document.getPage(1);
    const text = await page.getTextContent();
    if (text.items.length > 0) await page.getOperatorList();
  })().catch(() => undefined);
}

/**
 * Open the document of the file at `url` from `bytes` (left untouched), for
 * the reader that will show it. Nothing when one is already opened for the
 * URL, or where no worker can be started for it (pdf.js would then start one
 * itself and parse on it no sooner than the reader does).
 */
export function openPdfDocumentAhead(url: string, bytes: Uint8Array): void {
  if (!url || bytes.byteLength === 0 || bytes.byteLength > PDF_DOCUMENT_AHEAD_MAX_BYTES || ahead.has(url)) return;
  const kept = bytes.slice();
  const opened = (async (): Promise<PdfDocumentAhead | null> => {
    // (The worker first: where none can be started, pdf.js is not loaded
    // for a document nobody will open here.)
    const worker = await takePdfReaderWorker();
    if (!worker) return null;
    try {
      const pdfjs = await loadCompatiblePdfJs();
      const task = pdfjs.getDocument({
        data: kept.slice(),
        worker,
        ...PDF_JS_DOCUMENT_ASSETS,
        // As the reader opens it (components/notebook/contents/PDFViewer).
        fontExtraProperties: true,
      }) as unknown as PDFDocumentLoadingTask;
      bindPdfWorkerToTask(task, worker);
      // A file that does not open (a password, a broken file) is the
      // reader's to report: it meets the same refusal on this task.
      task.promise.then(readFirstPageAhead, () => undefined);
      return { task, worker };
    } catch {
      try { worker.destroy(); } catch { /* an ended worker stays ended */ }
      return null;
    }
  })().catch(() => null);
  const entry: Entry = { bytes: kept, opened, timer: null };
  entry.timer = setTimeout(() => {
    if (ahead.get(url) !== entry) return;
    ahead.delete(url);
    destroy(opened);
    forget(entry);
  }, TAKE_WITHIN_MS);
  (entry.timer as { unref?: () => void }).unref?.();
  ahead.set(url, entry);
}

/**
 * The document opened ahead for the file at `url`, for a reader that holds
 * `bytes`: its loading task (already bound to its worker), or null when none
 * was opened for exactly these bytes. Taken once; the reader owns it then,
 * and ends it as it ends a task of its own.
 */
export async function takePdfDocumentAhead(url: string, bytes: Uint8Array): Promise<PdfDocumentAhead | null> {
  const entry = ahead.get(url);
  if (!entry) return null;
  ahead.delete(url);
  if (entry.timer) clearTimeout(entry.timer);
  const same = sameBytes(entry.bytes, bytes);
  forget(entry);
  if (!same) {
    destroy(entry.opened);
    return null;
  }
  return entry.opened;
}

/** Whether a document is opened (or being opened) ahead for `url`. */
export function hasPdfDocumentAhead(url: string): boolean {
  return ahead.has(url);
}

/** Tests only. */
export function resetPdfDocumentsAheadForTests(): void {
  for (const entry of ahead.values()) {
    if (entry.timer) clearTimeout(entry.timer);
    destroy(entry.opened);
    forget(entry);
  }
  ahead.clear();
}
