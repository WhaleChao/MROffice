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
 * The page itself names its file's size, so a file read whole is asked for
 * together with its record, not after it: one round trip less before the
 * bytes arrive, and no wait for this thread (busy mounting the window) to
 * get to the record's answer before the download can start. That request
 * names the digests of the files kept in this browser (`If-None-Match`);
 * for one of them the server answers 304 and nothing is sent
 * (`startConditionalDownload`).
 *
 * Only the file the viewer will read whole is fetched: a file read in ranges
 * or a file whose bytes are kept here starts no download. A page with a saved
 * revision of its own (`editable_page_patchback`) shows that revision, read
 * from its own URL: its ticket is asked for here and its bytes read from it,
 * and the window shares that request (`readPdfVersionTicketAhead`) instead
 * of asking for one once it has mounted. What was read ahead is dropped a
 * short while after it arrived (and its bytes zeroed); a failed read is read
 * again by the window itself. The download is the viewer's own request
 * (`fetchPdfSource`, same URL), taken only for the digest the record names,
 * and treated as a download the viewer made (its inspection's digest decides
 * whether the bytes are kept).
 */

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import { getApiHttpBaseUrl } from "@mrnote/core/web/lib/env.ts";
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import { fetchPdfSource } from "./fetch-source";
import { getDataItemId } from "./page-content-model";
import { shouldReadPdfInRanges } from "./range-source";
import {
  listCachedPdfSourceDigests,
  pdfSourceBytesPrincipal,
  readCachedPdfSourceBytes,
} from "./source-bytes-cache";
import { openPdfDocumentAhead } from "./document-ahead";
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
  if (current && (current.settledAt === undefined || Date.now() - current.settledAt <= TAKE_WITHIN_MS)) {
    // Answered and taken by its window: not offered to another window. A
    // read ahead for the same opening (the page's prefetch runs once from
    // the press and again as the window and its editor mount) reads it
    // where it is: each such run after the window had taken its record
    
    const spent = current.taken === true && current.settledAt !== undefined;
    if (!options.take) return current.read as Promise<T>;
    if (!spent) {
      current.taken = true;
      return current.read as Promise<T>;
    }
  }
  const read = apiGet<T>(`/api/v1/data-items/${dataItemId}`);
  const entry: RecordRead = { read, ...(options.take ? { taken: true } : {}) };
  recordReads.set(dataItemId, entry);
  expireAfterSettling(recordReads, dataItemId, entry, read);
  return read;
}

/** The size a PDF page's content names for its file, or null. */
function pageSourceSizeBytes(content: unknown): number | null {
  const source = (content as { source?: unknown } | null | undefined)?.source;
  const size = source && typeof source === "object" ? (source as { sizeBytes?: unknown }).sizeBytes : null;
  return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : null;
}

/** A file asked for before its record has answered. */
interface EarlyDownload {
  /** The file's bytes; or the digest (one of those kept in this browser)
   * the server says the file has; or null: ask as before. */
  readonly answer: Promise<{ bytes: Uint8Array; kept?: undefined } | { kept: string; bytes?: undefined } | null>;
  /** Nobody will take it: stop the download and zero what arrived. */
  abandon(): void;
}
/** By URL, until the file's record has answered. */
const earlyDownloads = new Map<string, EarlyDownload>();

const ENTITY_TAG = /^(?:W\/)?"sha256-([0-9a-f]{64})"$/u;

/**
 * Ask for the file at `url` now, naming the digests of the files kept in
 * this browser: the server sends the bytes, or says which kept file it is.
 * (Without their names it would send a file that is already here; when they
 * cannot be listed in time nothing is asked, and the record decides as it
 * always did.)
 */
function startConditionalDownload(url: string): EarlyDownload {
  const controller = new AbortController();
  const answer = (async () => {
    const kept = await listCachedPdfSourceDigests();
    if (kept === null || controller.signal.aborted) return null;
    const response = await fetchPdfSource(url, {
      signal: controller.signal,
      ...(kept.length > 0
        ? { headers: { "If-None-Match": kept.map((digest) => `"sha256-${digest}"`).join(", ") } }
        : {}),
    });
    if (response.status === 304) {
      const digest = ENTITY_TAG.exec(response.headers.get("etag")?.trim() ?? "")?.[1];
      return digest && kept.includes(digest) ? { kept: digest } : null;
    }
    if (!response.ok) return null;
    return { bytes: new Uint8Array(await response.arrayBuffer()) };
  })().catch(() => null);
  return {
    answer,
    abandon: () => {
      controller.abort();
      void answer.then((got) => got?.bytes?.fill(0));
    },
  };
}

/** Start reading the file record of a PDF page and its bytes. */
export function prefetchPdfPageSource(page: PageItem | null | undefined): void {
  if (typeof window === "undefined" || !page || page.page_type !== "pdf") return;
  const dataItemId = getDataItemId(page.content_json);
  if (!dataItemId) return;
  const read = readPdfFileRecord<PdfFileRecord>(dataItemId);
  // A page with a saved revision of its own is read from that revision.
  if (page.editable_page_patchback != null) {
    prefetchPdfSavedRevisionSource(page);
    return;
  }
  const url = pdfDataItemPreviewUrl(dataItemId);
  // The page names its file's size: a file read whole is asked for with its
  // record, not after it.
  const sizeBytes = pageSourceSizeBytes(page.content_json);
  let early: EarlyDownload | null = null;
  if (url && sizeBytes !== null && !shouldReadPdfInRanges(sizeBytes) &&
      !sourceReads.has(url) && !earlyDownloads.has(url) && !uploadedSources.has(url)) {
    early = startConditionalDownload(url);
    earlyDownloads.set(url, early);
  }
  const release = () => {
    if (early && earlyDownloads.get(url) === early) earlyDownloads.delete(url);
  };
  void read.then((record) => {
    release();
    if (shouldReadPdfInRanges(record?.size_bytes)) { early?.abandon(); return; }
    prefetchPdfSourceBytes(url, record?.sha256, undefined, early ?? undefined);
  }, () => { release(); early?.abandon(); });
}

/** Files this tab has just uploaded whose record is on its way: their bytes
 * are the picked file's own, so nothing is asked of the server for them. */
const uploadedSources = new Set<string>();















export function prefetchUploadedPdfSource(dataItemId: string, file: Blob): void {
  if (typeof window === "undefined" || !dataItemId || !(file.size > 0) || shouldReadPdfInRanges(file.size)) return;
  const url = pdfDataItemPreviewUrl(dataItemId);
  if (!url || sourceReads.has(url) || earlyDownloads.has(url) || uploadedSources.has(url)) return;
  uploadedSources.add(url);
  const own = file.arrayBuffer().then((buffer) => new Uint8Array(buffer), () => null);
  const drop = () => { void own.then((bytes) => bytes?.fill(0)); };
  void readPdfFileRecord<PdfFileRecord>(dataItemId).then((record) => {
    uploadedSources.delete(url);
    if (shouldReadPdfInRanges(record?.size_bytes)) { drop(); return; }
    const same = record?.size_bytes === file.size;
    if (!same) drop();
    prefetchPdfSourceBytes(url, record?.sha256, same ? own : undefined);
  }, () => { uploadedSources.delete(url); drop(); });
}

/** The ticket of a saved revision (`resourceId:versionId`), asked ahead of
 * the window that shows it. */
interface VersionTicketRead { read: Promise<PdfVersionTicket | null>; settledAt?: number; answer?: PdfVersionTicket | null }
const versionTickets = new Map<string, VersionTicketRead>();

/** What `GET /resources/{id}/versions/{id}/download` answers. */
export interface PdfVersionTicket { download_url?: string | null }


















function prefetchPdfSavedRevisionSource(page: PageItem): void {
  const receipt = page.editable_page_patchback as {
    status?: unknown; adapter?: unknown; resource_id?: unknown; resource_version_id?: unknown; output_sha256?: unknown;
  } | null | undefined;
  if (!receipt || typeof receipt !== "object" || receipt.status !== "succeeded" || receipt.adapter !== "pdf") return;
  void readPdfVersionTicketAhead(
    typeof receipt.resource_id === "string" ? receipt.resource_id : "",
    typeof receipt.resource_version_id === "string" ? receipt.resource_version_id : "",
    typeof receipt.output_sha256 === "string" ? receipt.output_sha256 : "",
  );
}

/**
 * The ticket of one saved revision and, from its answer, the revision's
 * bytes: the request already on its way (or answered a moment ago), else a
 * new one. Whoever asks first -- the page read's prefetch or the window that
 * shows the revision -- makes the one request, and the bytes are read from
 * its answer at once, without waiting for the reader to mount and ask for
 * them. (A window that got to its page before the prefetch had asked for a
 * ticket of its own: two tickets, and the bytes read ahead under a URL the
 * reader was never given.) A promise of null: the request failed, or the
 * revision is not named well enough to ask for; the caller then asks as it
 * always did.
 */
export function readPdfVersionTicketAhead(resourceId: string, versionId: string,
  digest: string | null | undefined): Promise<PdfVersionTicket | null> {
  const sha256 = typeof digest === "string" ? digest.trim().toLowerCase() : "";
  if (typeof window === "undefined" || !/^[\w-]{1,64}$/u.test(resourceId) || !/^[\w-]{1,64}$/u.test(versionId) ||
      !/^[0-9a-f]{64}$/u.test(sha256)) return Promise.resolve(null);
  const key = `${resourceId}:${versionId}`;
  const current = versionTickets.get(key);
  if (current && (current.settledAt === undefined || Date.now() - current.settledAt <= TAKE_WITHIN_MS)) return current.read;
  const read = apiGet<PdfVersionTicket>(`/api/v1/resources/${resourceId}/versions/${versionId}/download`)
    .then((ticket) => (ticket && typeof ticket === "object" ? ticket : null), () => null);
  // The revision's bytes kept in this browser are looked for beside the
  // ticket, not once it has answered: the digest names them, the ticket only
  // the URL the reader will ask for them under.
  const kept = readCachedPdfSourceBytes(sha256).catch(() => null);
  const entry: VersionTicketRead = { read };
  versionTickets.set(key, entry);
  // (The answer is kept before the entry counts as settled: `peek` reads both.)
  const settled = read.then((ticket) => {
    entry.answer = ticket;
    const url = getSafeExternalUrl(ticket?.download_url || "");
    if (url && versionTickets.get(key) === entry) prefetchPdfSourceBytes(url, sha256, kept);
    else void kept.then((bytes) => bytes?.fill(0));
  });
  expireAfterSettling(versionTickets, key, entry, settled);
  return read;
}

/**
 * The ticket asked ahead for one saved revision when it has already
 * answered: what a window mounting now shows from its first render, where
 * it otherwise showed the file as uploaded until its own effect had taken
 * the ticket. Null while the request is on its way, when it failed, or when
 * none was asked (`readPdfVersionTicketAhead` then decides).
 */
export function peekPrefetchedPdfVersionTicket(resourceId: string, versionId: string): PdfVersionTicket | null {
  const entry = versionTickets.get(`${resourceId}:${versionId}`);
  if (!entry || entry.settledAt === undefined || Date.now() - entry.settledAt > TAKE_WITHIN_MS) return null;
  return entry.answer ?? null;
}

function prefetchPdfSourceBytes(url: string, digest: string | null | undefined,
  // The read of the bytes kept by this digest, when it is already on its way.
  kept?: Promise<Uint8Array | null>,
  // The same URL asked for before its record answered (`startConditionalDownload`).
  early?: EarlyDownload): void {
  const sha256 = typeof digest === "string" ? digest.trim().toLowerCase() : "";
  if (!/^[0-9a-f]{64}$/u.test(sha256) || !url || sourceReads.has(url)) {
    void kept?.then((bytes) => bytes?.fill(0));
    early?.abandon();
    return;
  }
  // (A viewer that already asked for this URL marked it: it downloads itself.)
  const principal = pdfSourceBytesPrincipal();
  // The same file read for another window of this login (the same bytes
  // uploaded twice): its bytes, checked against this record by the viewer.
  const twin = sameFileReadElsewhere(url, sha256, principal);
  const read = twin ? twin.read!.then((source) => source && { bytes: source.bytes.slice(), from: "cache" as const })
    : (async (): Promise<PrefetchedPdfSource | null> => {
    // Asked for with its record: the bytes are the URL's own download, as
    // the one below is. An answer that names a kept file is read from the
    // bytes kept by the record's digest, as any kept file is; no answer, or
    // a kept file that is no longer here, is downloaded as before.
    const first = early ? await early.answer : null;
    if (first?.bytes) return { bytes: first.bytes, from: "network" };
    const cached = await (kept ?? readCachedPdfSourceBytes(sha256).catch(() => null));
    if (cached) return { bytes: cached, from: "cache" };
    const response = await fetchPdfSource(url);
    if (!response.ok) return null;
    return { bytes: new Uint8Array(await response.arrayBuffer()), from: "network" };
  })().catch(() => null);
  if (twin) {
    void kept?.then((bytes) => bytes?.fill(0));
    early?.abandon();
  }
  const entry: SourceRead = { sha256, read, principal };
  sourceReads.set(url, entry);
  expireAfterSettling(sourceReads, url, entry, read, () => {
    void read.then((source) => source?.bytes.fill(0));
  });
  // The document is opened from them at once, while the reader that will
  // show it is still mounting (lib/pdf/document-ahead); the reader adopts it
  // only for the very bytes it ends up holding (kept bytes that do not hash
  // to the record are never shown: what was opened from them is ended
  // unread). pdf.js runs no document script and follows no action, so
  // nothing in the file acts by being parsed.
  void read.then((source) => {
    // A/B switch for timing probes: the reader then opens its document itself.
    if ((window as { __pdfDocumentAheadOff?: boolean }).__pdfDocumentAheadOff === true) return;
    if (source && sourceReads.get(url) === entry) openPdfDocumentAhead(url, source.bytes);
    // Its source inspection likewise (lib/pdf/source-risk-ahead): nothing
    // that writes the file is offered before it has landed. (A copy: these
    // bytes are zeroed when the read expires.)
    if (source && sourceReads.get(url) === entry && source.bytes.byteLength <= 32 * 1024 * 1024) {
      const held = source.bytes.slice();
      void import("./source-risk-ahead")
        .then(({ inspectPdfSourceRiskAhead }) => inspectPdfSourceRiskAhead(held))
        .catch(() => undefined)
        .finally(() => held.fill(0));
    }
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
    const claim = () => {
      // The viewer reads this file itself: nothing is read ahead for it now.
      const claimed = { sha256: "", read: null };
      sourceReads.set(url, claimed);
      const timer = setTimeout(() => { if (sourceReads.get(url) === claimed) sourceReads.delete(url); }, TAKE_WITHIN_MS);
      (timer as { unref?: () => void }).unref?.();
    };
    // Asked for with its record, which has not answered yet (the viewer
    // mounted first): the bytes on their way are this URL's own download,
    // whatever digest the record will name. The viewer started a second
    // download here, and a third when the record gave it the digest.
    const early = entry ? undefined : earlyDownloads.get(url);
    if (early) {
      return early.answer.then((got) => {
        if (got?.bytes) return { from: "network" as const, bytes: got.bytes.slice() };
        // No bytes came of it (a file kept here, a refusal): the viewer
        // reads the file itself, unless the record has since had it read.
        if (!sourceReads.has(url)) claim();
        return null;
      });
    }
    claim();
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
  for (const early of earlyDownloads.values()) early.abandon();
  earlyDownloads.clear();
  uploadedSources.clear();
  recordReads.clear();
  sourceReads.clear();
  versionTickets.clear();
}
