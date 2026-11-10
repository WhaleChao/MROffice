"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF page's file record and its bytes, read while its window's code is
 * still loading.
 *
 * Opening a PDF ran in sequence: the page read named the file, the window's
 * code loaded and mounted (over a second on a first open), the window asked
 * for the file record (`GET /data-items/{id}`: its size and digest), the
 * viewer mounted, looked in the bytes kept by digest (lib/pdf/source-bytes-
 * cache) and only then downloaded the file. The record depends on the page
 * alone and the download on the record alone, so both go out as soon as the
 * page read (or the copy of the page this browser kept) names the file
 * (lib/page-open-prefetch), and the window and its viewer take what is
 * already on its way.
 *
 * Only the file the viewer will read whole is fetched: a page with a saved
 * revision of its own (`editable_page_patchback`, read from another URL), a
 * file read in ranges, or a file whose bytes are kept here starts no
 * download. What was read ahead is dropped a short while after it arrived
 * (and its bytes zeroed); a failed read is read again by the window itself. The download
 * is the viewer's own request (`fetchPdfSource`, same URL), taken only for the
 * digest the record names, and treated as a download the viewer made (its
 * inspection's digest decides whether the bytes are kept).
 */

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import { getApiHttpBaseUrl } from "@mrnote/core/web/lib/env.ts";
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import { fetchPdfSource } from "./fetch-source";
import { getDataItemId } from "./page-content-model";
import { shouldReadPdfInRanges } from "./range-source";
import { pdfSourceBytesPrincipal, readCachedPdfSourceBytes } from "./source-bytes-cache";
import { warmPdfReaderWorker } from "./pdfjs-reader-worker";
import { getSafeExternalUrl } from "@mrnote/core/web/lib/security.ts";

interface PdfFileRecord {
  sha256?: string | null;
  size_bytes?: number | null;
}

/** How long a record or bytes read ahead wait to be taken once they arrived. */
const TAKE_WITHIN_MS = 15_000;

interface RecordRead { read: Promise<unknown>; settledAt?: number; taken?: boolean }
const recordReads = new Map<string, RecordRead>();
/** Bytes read ahead for a URL, or (`read` null) a viewer that already asked
 * for that URL and found nothing: nothing is then read ahead for it. */
interface SourceRead {
  sha256: string;
  read: Promise<PrefetchedPdfSource | null> | null;
  settledAt?: number;
  /** The login principal the bytes were read under (shared only with it). */
  principal?: string | null;
}
const sourceReads = new Map<string, SourceRead>();

export interface PrefetchedPdfSource {
  readonly bytes: Uint8Array;
  /** From the bytes kept by digest (the viewer checks the digest again) or
   * downloaded now (checked by the viewer's inspection, then kept). */
  readonly from: "cache" | "network";
}

/** The URL the viewer reads a page's file at when it shows the file itself. */
export function pdfDataItemPreviewUrl(dataItemId: string): string {
  return getSafeExternalUrl(`${getApiHttpBaseUrl()}/api/v1/data-items/${dataItemId}/preview`) || "";
}

function expireAfterSettling<T extends { settledAt?: number }>(map: Map<string, T>, key: string, entry: T,
  read: Promise<unknown>, drop?: () => void, takenAlready?: () => boolean): void {
  const retire = () => {
    if (map.get(key) !== entry) return;
    map.delete(key);
    drop?.();
  };
  read.then(() => {
    entry.settledAt = Date.now();
    if (takenAlready?.()) { retire(); return; }
    const timer = setTimeout(retire, TAKE_WITHIN_MS);
    (timer as { unref?: () => void }).unref?.();
  }, () => { if (map.get(key) === entry) map.delete(key); });
}

/**
 * The file record of `dataItemId`: the read already on its way (or read
 * ahead a moment ago) when there is one, else a new one. The window and the
 * read ahead share it, whichever asks first (a window restored with the
 * notebook asks before the page read that names the file has answered).
 * `take`: the window's own read; once it has answered it is not offered
 * again (the next window reads the record afresh, as it always did).
 */
export function readPdfFileRecord<T>(dataItemId: string, options: { take?: boolean } = {}): Promise<T> {
  // The worker its document will open on starts with it (pdfjs-reader-worker).
  warmPdfReaderWorker();
  const current = recordReads.get(dataItemId);
  if (current && !(current.taken && current.settledAt !== undefined) &&
      (current.settledAt === undefined || Date.now() - current.settledAt <= TAKE_WITHIN_MS)) {
    if (options.take) {
      current.taken = true;
      if (current.settledAt !== undefined) recordReads.delete(dataItemId);
    }
    return current.read as Promise<T>;
  }
  const read = apiGet<T>(`/api/v1/data-items/${dataItemId}`);
  const entry: RecordRead = { read, ...(options.take ? { taken: true } : {}) };
  recordReads.set(dataItemId, entry);
  expireAfterSettling(recordReads, dataItemId, entry, read, undefined, () => entry.taken === true);
  return read;
}

/** Start reading the file record of a PDF page and, from it, its bytes. */
export function prefetchPdfPageSource(page: PageItem | null | undefined): void {
  if (typeof window === "undefined" || !page || page.page_type !== "pdf") return;
  const dataItemId = getDataItemId(page.content_json);
  if (!dataItemId) return;
  const read = readPdfFileRecord<PdfFileRecord>(dataItemId);
  // A page with a saved revision of its own is read from that revision.
  if (page.editable_page_patchback != null) return;
  void read.then((record) => prefetchPdfSourceBytes(dataItemId, record), () => undefined);
}

function prefetchPdfSourceBytes(dataItemId: string, record: PdfFileRecord): void {
  const sha256 = typeof record?.sha256 === "string" ? record.sha256.trim().toLowerCase() : "";
  const url = pdfDataItemPreviewUrl(dataItemId);
  if (!/^[0-9a-f]{64}$/u.test(sha256) || !url || sourceReads.has(url) ||
      shouldReadPdfInRanges(record.size_bytes)) return;
  // (A viewer that already asked for this URL marked it: it downloads itself.)
  const principal = pdfSourceBytesPrincipal();
  // The same file read for another window of this login (the same bytes
  // uploaded twice): its bytes, checked against this record by the viewer.
  const twin = sameFileReadElsewhere(url, sha256, principal);
  const read = twin ? twin.read!.then((source) => source && { bytes: source.bytes.slice(), from: "cache" as const })
    : (async (): Promise<PrefetchedPdfSource | null> => {
    const cached = await readCachedPdfSourceBytes(sha256).catch(() => null);
    if (cached) return { bytes: cached, from: "cache" };
    const response = await fetchPdfSource(url);
    if (!response.ok) return null;
    return { bytes: new Uint8Array(await response.arrayBuffer()), from: "network" };
  })().catch(() => null);
  const entry: SourceRead = { sha256, read, principal };
  sourceReads.set(url, entry);
  expireAfterSettling(sourceReads, url, entry, read, () => {
    void read.then((source) => source?.bytes.fill(0));
  });
}

/** The bytes read ahead for the file at `url` with digest `sha256` (each
 * taker gets its own copy: pdf.js may take the buffer it is handed, and an
 * effect that is cancelled and run again takes them a second time); null
 * when none were read for exactly that file. */
export function takePrefetchedPdfSource(url: string, sha256: string | null | undefined,
  signal?: AbortSignal): Promise<PrefetchedPdfSource | null> | null {
  const entry = sourceReads.get(url);
  const read = entry?.read ?? null;
  if (!entry || !read) {
    // The viewer reads this file itself: nothing is read ahead for it now.
    const claimed = { sha256: "", read: null };
    sourceReads.set(url, claimed);
    const timer = setTimeout(() => { if (sourceReads.get(url) === claimed) sourceReads.delete(url); }, TAKE_WITHIN_MS);
    (timer as { unref?: () => void }).unref?.();
    return null;
  }
  if (typeof sha256 !== "string" || entry.sha256 !== sha256.trim().toLowerCase()) return null;
  // Another viewer's download (downloadPdfSourceOnce): this one waits for it too.
  if (signal) shareable.get(entry)?.join(signal);
  return read.then((source) => source && { from: source.from, bytes: source.bytes.slice() });
}

/** How long a viewer's download outlives the last viewer waiting for it: a
 * load that is run again (its inputs settled a moment later, StrictMode's
 * second run) asks for the same file a few awaits later and takes it. */
const ABANDONED_DOWNLOAD_GRACE_MS = 2_000;


























export function downloadPdfSourceOnce(
  url: string,
  sha256: string | null | undefined,
  signal: AbortSignal,
  download: (signal: AbortSignal) => Promise<Uint8Array>,
): Promise<PrefetchedPdfSource> {
  const digest = typeof sha256 === "string" ? sha256.trim().toLowerCase() : "";
  const principal = pdfSourceBytesPrincipal();
  const fresh = (entry: SourceRead) =>
    entry.settledAt === undefined || Date.now() - entry.settledAt <= TAKE_WITHIN_MS;
  let joined: SourceRead | null = null;
  let joinedFrom: PrefetchedPdfSource["from"] = "network";
  const own = sourceReads.get(url);
  if (own?.read && digest && own.sha256 === digest && fresh(own)) {
    joined = own;
  } else {
    joined = sameFileReadElsewhere(url, digest, principal);
    if (joined) joinedFrom = "cache";
  }
  const copy = (source: PrefetchedPdfSource | null, from: PrefetchedPdfSource["from"]) =>
    source ? { from, bytes: source.bytes.slice() } : null;
  if (joined) {
    const shared = shareable.get(joined);
    shared?.join(signal);
    return waitOrAbort(joined.read!.then((source) => copy(source, joinedFrom === "cache" ? "cache" : source?.from ?? "network")), signal)
      .then((source) => source ?? downloadPdfSourceAlone(signal, download));
  }
  const controller = new AbortController();
  const read: Promise<PrefetchedPdfSource | null> = download(controller.signal)
    .then((bytes) => ({ bytes, from: "network" as const }));
  // Taken by the viewers that wait for it; a refused or failed download is
  // not offered again (the next viewer asks for itself).
  const offered = read.catch(() => null);
  const entry: SourceRead = { sha256: /^[0-9a-f]{64}$/u.test(digest) ? digest : "", read: offered, principal };
  // A URL read ahead for another digest keeps its own entry; this download
  // is then this viewer's alone.
  const registered = !(own?.read && own.sha256 && own.sha256 !== entry.sha256) && entry.sha256 !== "";
  if (registered) {
    sourceReads.set(url, entry);
    expireAfterSettling(sourceReads, url, entry, offered, () => {
      void offered.then((source) => source?.bytes.fill(0));
    });
  }
  const waiters = createWaiters(controller, read);
  if (registered) shareable.set(entry, waiters);
  waiters.join(signal);
  return waitOrAbort(read.then((source) => copy(source, "network")!), signal);
}

/** A read of the same file (digest) for this login under another URL, on
 * its way or arrived a moment ago. */
function sameFileReadElsewhere(url: string, digest: string, principal: string | null): SourceRead | null {
  if (!/^[0-9a-f]{64}$/u.test(digest) || !principal) return null;
  for (const [otherUrl, entry] of sourceReads) {
    if (otherUrl === url || !entry.read || entry.sha256 !== digest || entry.principal !== principal) continue;
    if (entry.settledAt !== undefined && Date.now() - entry.settledAt > TAKE_WITHIN_MS) continue;
    return entry;
  }
  return null;
}

function downloadPdfSourceAlone(
  signal: AbortSignal,
  download: (signal: AbortSignal) => Promise<Uint8Array>,
): Promise<PrefetchedPdfSource> {
  return download(signal).then((bytes) => ({ bytes, from: "network" as const }));
}

interface DownloadWaiters {
  /** One more viewer waits for the download until `signal` aborts. */
  join(signal: AbortSignal): void;
}

/** The viewers waiting for a viewer's shared download. */
const shareable = new WeakMap<SourceRead, DownloadWaiters>();

function createWaiters(controller: AbortController, read: Promise<unknown>): DownloadWaiters {
  let waiting = 0;
  let settled = false;
  let graceTimer: ReturnType<typeof setTimeout> | null = null;
  const settle = () => {
    settled = true;
    if (graceTimer) clearTimeout(graceTimer);
    graceTimer = null;
  };
  read.then(settle, settle);
  const leave = () => {
    waiting -= 1;
    if (waiting > 0 || settled) return;
    graceTimer = setTimeout(() => {
      graceTimer = null;
      if (waiting === 0 && !settled) controller.abort();
    }, ABANDONED_DOWNLOAD_GRACE_MS);
    (graceTimer as { unref?: () => void }).unref?.();
  };
  return {
    join(signal) {
      if (signal.aborted || settled) return;
      waiting += 1;
      if (graceTimer) clearTimeout(graceTimer);
      graceTimer = null;
      signal.addEventListener("abort", leave, { once: true });
    },
  };
}

/** `read`, or an AbortError as soon as `signal` aborts (the shared read
 * itself goes on for the others). */
function waitOrAbort<T>(read: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) return Promise.reject(new DOMException("Aborted", "AbortError"));
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(new DOMException("Aborted", "AbortError"));
    signal.addEventListener("abort", onAbort, { once: true });
    read.then(
      (value) => { signal.removeEventListener("abort", onAbort); resolve(value); },
      (error) => { signal.removeEventListener("abort", onAbort); reject(error); },
    );
  });
}

/** Tests only. */
export function resetPdfSourcePrefetch(): void {
  recordReads.clear();
  sourceReads.clear();
}
