/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A large file uploaded in parts: a source file becoming a page, made into its
 * page on the server (`app/services/source_imports.py`), or an editor's file
 * (a stored file, or a page's attachment) past what one request body carries.
 *
 * The file is cut into the parts the server asks for and six are sent at a
 * time. Where the server allows it (`direct_parts`) each part goes straight to
 * object storage, to an address signed for that part and its exact length, so
 * no server carries the bytes; otherwise, and for any part that could not go
 * that way, it is a PUT of the upload proxy route on the file host
 * (`lib/private-file-routes.json`). Both write the same upload. A part that
 * fails in transit is sent again; the upload fails only when one part has
 * failed four times, or the server refuses it for a reason a retry cannot
 * change. The import then starts and is followed until its page exists.
 *
 * What a file becomes is the server's decision: an editable page, a
 * read-only preview of its leading part with the original to download, or a
 * page offering the original alone.
 */

import {
  ApiRequestError,
  apiDelete,
  apiFetchRaw,
  apiGet,
  apiPost,
  isApiRequestError,
} from "../api";
import { reportDirectUploadOutcome, type DirectUploadFallback } from "./direct-upload-report";
import { lazyFileReader } from "./lazy-file";
import { desktopLocalFileToken, linkDesktopOriginal } from "../desktop-local-files";
import { FILE_ACCESS_REQUIRED_CODE } from "../file-access-repair";
import { beginTransfer } from "../transfer-store";
import { fetchWithUploadProgress, reportUploadTo } from "../upload-progress-fetch";

/**
 * The largest file sent as one upload request: anything larger goes up in
 * parts, each retried, and straight to object storage where that is open.
 */
export const SINGLE_REQUEST_UPLOAD_MAX_BYTES = 8 * 1024 * 1024;

/** The file kinds the server keeps (and may convert), by its own declaration. */
export interface SourceUploadSession {
  upload_id: string;
  data_item_id: string;
  size_bytes: number;
  part_size: number;
  part_count: number;
  expires_at: string | null;
  /** Whether this page may send the parts straight to object storage. */
  direct_parts?: boolean;
}

export type SourceImportOutcome = "pdf" | "editable" | "preview" | "original";

export interface SourceImportJob {
  job_id: string;
  upload_id: string | null;
  status: "queued" | "running" | "succeeded" | "failed" | "cancelled";
  /** `streaming_parts`: a deck's page is made from its first slides and the
   * import is still writing its later parts (INTERFACE-deck-parts.md 13.3). */
  phase: "queued" | "reading" | "creating_page" | "streaming_parts" | "done";
  /** While the import still runs: its page exists and opens (a deck held as
   * parts whose later parts are still being written), `page_id` naming it. */
  page_ready?: boolean;
  revision: number;
  page_id: string | null;
  page_type: string | null;
  outcome: SourceImportOutcome | null;
  outcome_reason: "too_large" | "not_convertible" | "unreadable" | `encrypted:${string}` | null;
  title: string | null;
  folder_id: string | null;
  filename: string | null;
  size_bytes: number | null;
  source_kind: string | null;
  /** While a long deck is imported: its first slides, read beside the import
   * and kept for the import preview window (lib/document-import/pptx-first-slides). */
  first_slides?: { kept_slides: number; total_slides: number } | null;
  error: {
    code: string;
    message: string;
    status: number;
    details?: Record<string, unknown> | null;
  } | null;
}

export interface SourceImportProgress {
  /** ``uploading`` while parts are sent, then ``converting`` until the page exists. */
  readonly phase: "uploading" | "converting";
  /** Bytes the server has accepted (uploading). */
  readonly sentBytes: number;
  readonly totalBytes: number;
  /** While converting: the step the server reports for the import
   * (``SourceImportJob.phase``: waiting its turn, reading the file, writing
   * the page). Absent until the server has answered once. */
  readonly serverPhase?: SourceImportJob["phase"];
  /** While converting a long deck: the import's job, once its first slides
   * are kept for the preview window (`GET .../source-imports/{job}/first-slides`). */
  readonly firstSlides?: { readonly jobId: string; readonly keptSlides: number; readonly totalSlides: number };
}

/** A refusal or failure of one source import, with the file it was for. */
export class SourceImportError extends Error {
  readonly code: string;
  readonly status: number;
  readonly details: Record<string, unknown>;
  readonly fileName: string;
  readonly fileSize: number;

  constructor(
    message: string,
    options: {
      code: string;
      status: number;
      details?: Record<string, unknown> | null;
      fileName: string;
      fileSize: number;
      cause?: unknown;
    },
  ) {
    super(message, { cause: options.cause });
    this.name = "SourceImportError";
    this.code = options.code;
    this.status = options.status;
    this.details = options.details ?? {};
    this.fileName = options.fileName;
    this.fileSize = options.fileSize;
  }
}

// Six parts in flight: the file host speaks HTTP/1.1, so each part gets its
// own TCP connection and they really run side by side (browsers open six
// per host). Over the ~90 ms path one connection alone reached ~5 MiB/s.
const PARALLEL_PARTS = 6;
const PART_ATTEMPTS = 4;
const RETRY_DELAYS_MS = [1000, 3000, 9000];

function delay(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(signal.reason ?? new DOMException("Aborted", "AbortError"));
      return;
    }
    const timer = window.setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    const onAbort = () => {
      window.clearTimeout(timer);
      reject(signal?.reason ?? new DOMException("Aborted", "AbortError"));
    };
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

function isAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

/** Whether sending the same part again can succeed. */
function retryable(error: unknown): boolean {
  if (!isApiRequestError(error)) return true; // the connection itself failed
  // The file host's own session, already reconnected once by the request
  // itself (lib/file-access-repair): asking again only repeats the refusal.
  if (error.code === FILE_ACCESS_REQUIRED_CODE) return false;
  if (error.status === 0 || error.status >= 500) return true;
  return error.status === 408 || error.status === 429;
}

function asImportError(error: unknown, file: SourceFileLike): unknown {
  if (isAbort(error) || error instanceof SourceImportError) return error;
  if (isApiRequestError(error)) {
    return new SourceImportError(error.message, {
      code: error.code || (error.status === 0 ? "network" : "request_failed"),
      status: error.status,
      details: error.details,
      fileName: file.name,
      fileSize: file.size,
      cause: error,
    });
  }
  return new SourceImportError(
    error instanceof Error ? error.message : "Upload failed",
    {
      code: "network",
      status: 0,
      fileName: file.name,
      fileSize: file.size,
      cause: error,
    },
  );
}

/**
 * The bytes of a file the uploader reads part by part. A browser `File` is
 * one; the desktop app hands over files it reads from disk on request.
 */
export interface SourceFileLike {
  readonly name: string;
  readonly size: number;
  readonly type: string;
  readPart(start: number, end: number): Promise<Blob>;
  /** The file itself, when one is behind this: MRNote Classic reads a file
   * on this computer where it is (lib/desktop-local-files). */
  readonly localFile?: File;
}

export function sourceFileFromFile(file: File): SourceFileLike {
  const lazy = lazyFileReader(file);
  return {
    name: file.name,
    size: file.size,
    type: file.type,
    readPart: lazy ?? (async (start, end) => file.slice(start, end)),
    localFile: file,
  };
}

/** MRNote Classic: every part of the upload read by the app's API from the
 * file on this computer. False when there is no such file (or it could not
 * be read): the parts are then sent from the page as anywhere else. */
async function fillPartsFromLocalFile(
  session: SourceUploadSession,
  file: SourceFileLike,
  signal?: AbortSignal,
): Promise<boolean> {
  const token = file.localFile ? await desktopLocalFileToken(file.localFile) : null;
  if (!token) return false;
  try {
    await apiPost(
      `/api/v1/desktop/source-uploads/${encodeURIComponent(session.upload_id)}/from-local`,
      { token },
      { signal },
    );
    return true;
  } catch (error) {
    if (isAbort(error)) throw error;
    return false;
  }
}

/** How many addresses one request asks for, more than are sent at a time. */
const PART_URL_BATCH = PARALLEL_PARTS * 2;
const DIRECT_FAILURES_ALLOWED = 3;
const SIGNATURE_READ_BYTES = 8192;
const DIRECT_UNAVAILABLE_CODE = "direct_upload_unavailable";

/**
 * This browser could not reach object storage at all (a network that blocks
 * it): later uploads of this page do not try again part by part.
 */
let storageUnreachable = false;

export function forgetDirectUploadOutcomeForTests(): void {
  storageUnreachable = false;
}

async function firstBytesBase64(file: SourceFileLike): Promise<string> {
  const blob = await file.readPart(0, Math.min(file.size, SIGNATURE_READ_BYTES));
  let text = "";
  for (const byte of new Uint8Array(await blob.arrayBuffer())) text += String.fromCharCode(byte);
  return btoa(text);
}

/**
 * Sends parts straight to object storage. It answers `false` whenever a part
 * should take the upload route instead, so a blocked or failing road costs one
 * attempt and never the upload; only a refusal of the file itself is thrown.
 */
class DirectParts {
  private readonly urls = new Map<number, string>();
  private asking: Promise<void> | null = null;
  private usable: boolean;
  private sent = 0;
  private proxied = 0;
  private failures = 0;
  /** Why the first part that could have gone straight to storage did not. */
  private fallback: DirectUploadFallback | null = null;

  constructor(
    private readonly session: SourceUploadSession,
    private readonly file: SourceFileLike,
    private readonly signal?: AbortSignal,
  ) {
    this.usable = Boolean(session.direct_parts) && !storageUnreachable;
    if (session.direct_parts && storageUnreachable) this.fallback = "storage_unreachable_earlier";
  }

  private fellBack(reason: DirectUploadFallback): false {
    this.fallback ??= reason;
    return false;
  }

  /** A part this road did not send went through the upload route. */
  noteProxied(): void {
    this.proxied += 1;
  }

  /** Once every part is in: how the road went, where the server offered it. */
  reportOutcome(): void {
    if (!this.session.direct_parts) return;
    reportDirectUploadOutcome({ fallback: this.proxied > 0 ? this.fallback ?? "other" : null, direct: this.sent, proxied: this.proxied });
  }

  private async ask(part: number): Promise<void> {
    const parts: number[] = [];
    const last = Math.min(this.session.part_count, part + PART_URL_BATCH - 1);
    for (let number = part; number <= last; number += 1) {
      if (!this.urls.has(number)) parts.push(number);
    }
    const answer = await apiPost<{ urls: Record<string, string> }>(
      `/api/v1/source-uploads/${encodeURIComponent(this.session.upload_id)}/part-urls`,
      {
        parts,
        ...(parts.includes(1) ? { first_bytes: await firstBytesBase64(this.file) } : {}),
      },
      { signal: this.signal },
    );
    for (const [number, url] of Object.entries(answer.urls)) this.urls.set(Number(number), url);
  }

  private async urlFor(part: number): Promise<string | null> {
    while (!this.urls.has(part)) {
      if (!this.asking) {
        this.asking = this.ask(part).finally(() => {
          this.asking = null;
        });
        await this.asking;
        if (!this.urls.has(part)) return null;
      } else {
        // Another part's request may carry this one too; if not, ask next.
        await this.asking.catch(() => undefined);
      }
    }
    return this.urls.get(part) ?? null;
  }

  async send(part: number, body: Blob): Promise<boolean> {
    if (!this.usable) return false;
    let url: string | null;
    try {
      url = await this.urlFor(part);
    } catch (error) {
      if (isAbort(error)) throw error;
      if (isApiRequestError(error) && !retryable(error) && error.code !== DIRECT_UNAVAILABLE_CODE) {
        throw error;
      }
      this.usable = false;
      const code = isApiRequestError(error) ? error.code || String(error.status) : "failed";
      return this.fellBack(`part_urls_${String(code).toLowerCase().replace(/[^a-z0-9_]/g, "_").slice(0, 48)}`);
    }
    if (!url) return this.fellBack("part_urls_missing");
    let reachedStorage = false;
    let status = 0;
    try {
      const response = await fetchWithUploadProgress(url, {
        method: "PUT",
        body,
        signal: this.signal,
        credentials: "omit",
        redirect: "error",
      });
      reachedStorage = true;
      status = response.status;
      if (response.ok) {
        this.sent += 1;
        return true;
      }
    } catch (error) {
      if (isAbort(error)) throw error;
    }
    // An address may have expired; a part sent again asks for a new one.
    this.urls.delete(part);
    this.failures += 1;
    if (this.sent === 0 || this.failures >= DIRECT_FAILURES_ALLOWED) this.usable = false;
    if (this.sent === 0 && !reachedStorage) storageUnreachable = true;
    return this.fellBack(reachedStorage ? `storage_${status}` : "storage_unreachable");
  }
}

async function sendPart(
  session: SourceUploadSession,
  file: SourceFileLike,
  part: number,
  direct: DirectParts,
  signal?: AbortSignal,
  onPartSent?: (sentBytes: number) => void,
): Promise<number> {
  const start = (part - 1) * session.part_size;
  const end = Math.min(file.size, start + session.part_size);
  let lastError: unknown = null;
  for (let attempt = 0; attempt < PART_ATTEMPTS; attempt += 1) {
    if (attempt > 0) await delay(RETRY_DELAYS_MS[attempt - 1] ?? 9000, signal);
    try {
      const read = await file.readPart(start, end);
      // the part's bytes as they go out, so a large part moves the line
      // while it is sent rather than when it lands
      onPartSent?.(0);
      const body = onPartSent && read instanceof Blob ? reportUploadTo(read, onPartSent) : read;
      if (await direct.send(part, body)) return end - start;
      onPartSent?.(0);
      await apiFetchRaw(
        // The upload proxy route, which the file host already carries.
        `/api/v1/uploads/proxy/${encodeURIComponent(session.upload_id)}?part=${part}`,
        { method: "PUT", body, signal },
        { contentType: "application/octet-stream" },
      );
      direct.noteProxied();
      return end - start;
    } catch (error) {
      if (isAbort(error) || !retryable(error)) throw error;
      lastError = error;
    }
  }
  throw lastError;
}

/**
 * Send every part, six at a time, reporting the bytes accepted so far (to the
 * caller and to the window's transfer line).
 */
async function sendParts(
  session: SourceUploadSession,
  file: SourceFileLike,
  onProgress: ((sent: number) => void) | undefined,
  signal?: AbortSignal,
): Promise<void> {
  let next = 1;
  let sent = 0;
  const transfer = beginTransfer({ direction: "upload", name: file.name, totalBytes: file.size });
  const direct = new DirectParts(session, file, signal);
  // bytes of the parts still going out, by part
  const inFlight = new Map<number, number>();
  const report = () => {
    let total = sent;
    for (const bytes of inFlight.values()) total += bytes;
    transfer.update(total);
    onProgress?.(total);
  };
  const worker = async () => {
    while (next <= session.part_count) {
      const part = next;
      next += 1;
      const partSent = await sendPart(session, file, part, direct, signal, (bytes) => {
        inFlight.set(part, bytes);
        report();
      });
      inFlight.delete(part);
      sent += partSent;
      report();
    }
  };
  try {
    await Promise.all(
      Array.from({ length: Math.min(PARALLEL_PARTS, session.part_count) }, worker),
    );
    direct.reportOutcome();
  } finally {
    transfer.end();
  }
}

/** Wait for an import already started (one listed after a reload);
 * ``onJob`` hears each state the server answers with, the first included. */
export async function followSourceImport(
  notebookId: string,
  job: SourceImportJob,
  signal?: AbortSignal,
  onJob?: (job: SourceImportJob) => void,
): Promise<SourceImportJob> {
  return followImport(notebookId, job, signal, onJob);
}

/** The caller's imports into this notebook whose pages are still being made. */
export async function listActiveSourceImports(notebookId: string): Promise<SourceImportJob[]> {
  const response = await apiGet<{ items?: SourceImportJob[] }>(
    `/api/v1/notebooks/${encodeURIComponent(notebookId)}/source-imports`,
  );
  return Array.isArray(response.items) ? response.items : [];
}

/** Whether the import's page opens already, before the import ends: a
 * deck whose page is made from its first slides while the rest is still
 * written part by part (the editor reads those parts as they come). */
export function sourceImportPageReady(job: SourceImportJob): boolean {
  return (
    job.status === "running" &&
    job.page_ready === true &&
    typeof job.page_id === "string" &&
    job.page_id.length > 0
  );
}

async function followImport(
  notebookId: string,
  job: SourceImportJob,
  signal?: AbortSignal,
  onJob?: (job: SourceImportJob) => void,
): Promise<SourceImportJob> {
  let current = job;
  onJob?.(current);
  while ((current.status === "queued" || current.status === "running") && !sourceImportPageReady(current)) {
    if (signal?.aborted) {
      throw signal.reason ?? new DOMException("Aborted", "AbortError");
    }
    try {
      current = await apiGet<SourceImportJob>(
        `/api/v1/notebooks/${encodeURIComponent(notebookId)}/source-imports/${encodeURIComponent(
          current.job_id,
        )}?wait_ms=5000&since=${current.revision}`,
        { signal },
      );
      onJob?.(current);
    } catch (error) {
      // The page is being made on the server whatever this connection does;
      // keep asking after it.
      if (isAbort(error) || (isApiRequestError(error) && error.status === 404)) throw error;
      await delay(3000, signal);
    }
  }
  return current;
}

/** What an upload in parts becomes (the server admits it under that kind's rules). */
export type SourceUploadPurpose =
  | { readonly kind: "page" }
  | { readonly kind: "file" }
  | { readonly kind: "attachment"; readonly pageId: string };

export interface UploadInPartsOptions {
  readonly signal?: AbortSignal;
  /** Bytes the server has accepted so far. */
  readonly onProgress?: (sentBytes: number, totalBytes: number) => void;
}

/**
 * Start an upload of ``file`` in parts and send every part. Resolves with
 * the session once all parts have arrived; the caller completes it into
 * what ``purpose`` names. On a failure or abort the reserved storage is
 * given back.
 */
async function uploadInParts(
  notebookId: string,
  file: SourceFileLike,
  purpose: SourceUploadPurpose,
  options: UploadInPartsOptions,
): Promise<SourceUploadSession> {
  let session: SourceUploadSession | null = null;
  try {
    session = await apiPost<SourceUploadSession>(
      `/api/v1/notebooks/${encodeURIComponent(notebookId)}/source-uploads`,
      {
        filename: file.name,
        media_type: file.type || "application/octet-stream",
        size_bytes: file.size,
        ...(purpose.kind === "page" ? {} : { purpose: purpose.kind }),
        ...(purpose.kind === "attachment" ? { page_id: purpose.pageId } : {}),
      },
      { signal: options.signal },
    );
    options.onProgress?.(0, file.size);
    if (await fillPartsFromLocalFile(session, file, options.signal)) {
      options.onProgress?.(file.size, file.size);
    } else {
      await sendParts(session, file, (sent) => options.onProgress?.(sent, file.size), options.signal);
    }
    return session;
  } catch (error) {
    if (session) cancelSourceUpload(session.upload_id);
    throw asImportError(error, file);
  }
}

function cancelSourceUpload(uploadId: string): void {
  // Give the reserved storage back; the server reclaims the parts.
  void apiDelete(`/api/v1/source-uploads/${encodeURIComponent(uploadId)}`).catch(
    () => undefined,
  );
}

/** A stored file of a notebook (a DataItem), as an editor inserts one. */
export interface StoredEditorFile {
  readonly data_item_id: string;
  readonly filename: string;
  readonly media_type: string;
  readonly size_bytes: number;
}

/**
 * Upload ``file`` in parts as a stored file of the notebook (no page): a
 * slide's video or picture, a PDF's attachment. Past one request body, which
 * the server bounds, this is how an editor's file arrives at all.
 */
export async function uploadEditorFileInParts(
  notebookId: string,
  file: SourceFileLike,
  options: UploadInPartsOptions = {},
): Promise<StoredEditorFile> {
  const session = await uploadInParts(notebookId, file, { kind: "file" }, options);
  try {
    return await apiPost<StoredEditorFile>(
      `/api/v1/source-uploads/${encodeURIComponent(session.upload_id)}/complete`,
      {},
      { signal: options.signal },
    );
  } catch (error) {
    if (isAbort(error) || options.signal?.aborted) cancelSourceUpload(session.upload_id);
    throw asImportError(error, file);
  }
}

/** A page attachment, as `POST /pages/{id}/attachments/upload` answers it. */
export interface UploadedPageAttachment {
  readonly attachment_id: string;
  readonly filename: string;
  readonly mime_type: string;
  readonly size_bytes: number;
  readonly attachment_type: string;
}

/** Upload ``file`` in parts as an attachment of ``pageId``. */
export async function uploadPageAttachmentInParts(
  notebookId: string,
  pageId: string,
  file: SourceFileLike,
  options: UploadInPartsOptions & { readonly title?: string } = {},
): Promise<UploadedPageAttachment> {
  const session = await uploadInParts(notebookId, file, { kind: "attachment", pageId }, options);
  try {
    return await apiPost<UploadedPageAttachment>(
      `/api/v1/source-uploads/${encodeURIComponent(session.upload_id)}/attach`,
      { page_id: pageId, ...(options.title ? { title: options.title } : {}) },
      { signal: options.signal },
    );
  } catch (error) {
    if (isAbort(error) || options.signal?.aborted) cancelSourceUpload(session.upload_id);
    throw asImportError(error, file);
  }
}

export interface ImportSourceFileOptions {
  readonly title: string;
  readonly folderId?: string | null;
  readonly signal?: AbortSignal;
  readonly onProgress?: (progress: SourceImportProgress) => void;
  /** A delimited file's Automatic Data Conversion switches, "1111" in
   *  Excel's order (lib/spreadsheet/data-conversion-options). */
  readonly csvConversions?: string;
}

/**
 * Upload ``file`` in parts and make it a page. Resolves with the finished
 * import (its page, and what the file became); rejects with a
 * {@link SourceImportError} naming the file.
 */
export async function importSourceFile(
  notebookId: string,
  file: SourceFileLike,
  options: ImportSourceFileOptions,
): Promise<SourceImportJob> {
  let session: SourceUploadSession | null = null;
  try {
    session = await uploadInParts(notebookId, file, { kind: "page" }, {
      signal: options.signal,
      onProgress: (sent, total) =>
        options.onProgress?.({ phase: "uploading", sentBytes: sent, totalBytes: total }),
    });
    const started = session;
    options.onProgress?.({ phase: "converting", sentBytes: file.size, totalBytes: file.size });
    const job = await apiPost<SourceImportJob>(
      `/api/v1/notebooks/${encodeURIComponent(notebookId)}/source-imports`,
      {
        upload_id: started.upload_id,
        title: options.title,
        ...(options.folderId ? { folder_id: options.folderId } : {}),
        ...(options.csvConversions ? { csv_conversions: options.csvConversions } : {}),
      },
      { signal: options.signal },
    );
    const finished = await followImport(notebookId, job, options.signal, (current) => {
      if (current.status !== "queued" && current.status !== "running") return;
      options.onProgress?.({
        phase: "converting",
        sentBytes: file.size,
        totalBytes: file.size,
        serverPhase: current.phase,
        ...(current.first_slides
          ? {
              firstSlides: {
                jobId: current.job_id,
                keptSlides: current.first_slides.kept_slides,
                totalSlides: current.first_slides.total_slides,
              },
            }
          : {}),
      });
    });
    // A deck held as parts opens as soon as its page exists; the import
    // goes on writing the later parts, which the editor reads as they come.
    const opened = sourceImportPageReady(finished);
    if ((finished.status !== "succeeded" && !opened) || !finished.page_id) {
      const error = finished.error;
      throw new SourceImportError(error?.message || "Import failed", {
        code: error?.code || "source_import_failed",
        status: error?.status || 500,
        details: error?.details,
        fileName: file.name,
        fileSize: file.size,
      });
    }
    if (file.localFile) linkDesktopOriginal(finished.page_id, file.localFile, notebookId);
    // Only a deck's import opens its page early: it is that deck, editable.
    return opened
      ? { ...finished, page_type: finished.page_type ?? "presentation", outcome: finished.outcome ?? "editable" }
      : finished;
  } catch (error) {
    if (session && (isAbort(error) || options.signal?.aborted)) {
      cancelSourceUpload(session.upload_id);
    }
    throw asImportError(error, file);
  }
}

export function isSourceImportError(error: unknown): error is SourceImportError {
  return error instanceof SourceImportError;
}

/** Kept for callers that need to tell a plain API refusal apart. */
export { ApiRequestError };
