/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PDFDocument as PdfLibDocument } from "pdf-lib";

import { pdfPageWindowMutationGuardV2Enabled } from "./pdf-flags";
import { waitForDecryptedPdfBytes } from "./pdf-password";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export interface PdfMutationFetchResponse {
  ok: boolean;
  status: number;
  arrayBuffer: () => Promise<ArrayBuffer>;
  headers?: { get: (name: string) => string | null };
  body?: {
    getReader: () => {
      read: () => Promise<{
        done: boolean;
        value?: ArrayBufferView;
      }>;
      cancel: (reason?: unknown) => Promise<void>;
      releaseLock?: () => void;
    };
  } | null;
}

export type PdfMutationFetcher = (
  url: string,
  init: { credentials: "include"; signal?: AbortSignal },
) => Promise<PdfMutationFetchResponse>;

async function fetchPdfBytes(
  previewUrl: string,
  fetcher: PdfMutationFetcher,
): Promise<ArrayBuffer> {
  const response = await fetcher(previewUrl, { credentials: "include" });
  if (!response.ok) throw new Error(`preview ${response.status}`);
  return response.arrayBuffer();
}

export const MAX_PDF_PAGE_WINDOW_MUTATION_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_PAGE_WINDOW_MUTATION_FETCH_TIMEOUT_MS = 8_000;

/** Preserve HTTP identity so an authenticated version reader can renew an
 * expired transport lease without parsing error text or relaxing byte checks. */
export class PdfMutationSourceHttpError extends Error {
  constructor(readonly status: number) {
    super(`preview ${status}`);
    this.name = "PdfMutationSourceHttpError";
  }
}

class PdfMutationSourceTimeoutError extends Error {
  constructor() {
    super("PDF source fetch timed out");
    this.name = "PdfMutationSourceTimeoutError";
  }
}

class PdfMutationSourceTransportError extends Error {
  constructor(error: unknown) {
    super(
      error instanceof Error
        ? error.message
        : "PDF source fetch failed before a response was received",
    );
    this.name = "PdfMutationSourceTransportError";
  }
}

function isAbortError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "name" in error &&
    error.name === "AbortError"
  );
}

function exactContentLength(response: PdfMutationFetchResponse): number | null {
  const raw = response.headers?.get("content-length")?.trim();
  if (!raw) return null;
  if (!/^[0-9]+$/.test(raw)) {
    throw new Error("Invalid PDF preview Content-Length");
  }
  const length = Number(raw);
  if (!Number.isSafeInteger(length) || length < 1) {
    throw new Error("Invalid PDF preview Content-Length");
  }
  return length;
}

async function readBoundedPdfResponse(
  response: PdfMutationFetchResponse,
  maxBytes = MAX_PDF_PAGE_WINDOW_MUTATION_BYTES,
): Promise<Uint8Array> {
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 1) {
    throw new Error("Invalid PDF preview byte limit");
  }
  const declaredLength = exactContentLength(response);
  if (declaredLength !== null && declaredLength > maxBytes) {
    throw new Error("PDF preview exceeds the editable PDF mutation limit");
  }

  const reader = response.body?.getReader();
  if (!reader) {
    // A body without a stream can only be admitted when its allocation was
    // bounded by an exact response length first. Browser Responses expose a
    // stream; this fallback exists only for narrowly modelled test/bridge
    // responses and never turns a missing length into allocation authority.
    if (declaredLength === null) {
      throw new Error("PDF preview is not bounded-stream readable");
    }
    const buffer = await response.arrayBuffer();
    if (
      buffer.byteLength !== declaredLength ||
      buffer.byteLength < 1 ||
      buffer.byteLength > maxBytes
    ) {
      throw new Error("PDF preview byte length mismatch");
    }
    return new Uint8Array(buffer.slice(0));
  }

  const chunks: Uint8Array[] = [];
  let total = 0;
  let completed = false;
  try {
    while (true) {
      const next = await reader.read();
      if (next.done) break;
      const view = next.value;
      if (!view || !ArrayBuffer.isView(view)) {
        throw new Error("Invalid PDF preview stream chunk");
      }
      if (view.byteLength === 0) continue;
      const chunk = new Uint8Array(
        view.buffer,
        view.byteOffset,
        view.byteLength,
      );
      total += view.byteLength;
      if (total > maxBytes || total > Number.MAX_SAFE_INTEGER) {
        await reader.cancel("PDF preview byte limit exceeded").catch(() => {});
        throw new Error("PDF preview exceeds the editable PDF mutation limit");
      }
      const owned = new Uint8Array(view.byteLength);
      owned.set(chunk);
      chunks.push(owned);
    }
    if (total < 1 || (declaredLength !== null && total !== declaredLength)) {
      throw new Error("PDF preview byte length mismatch");
    }
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.byteLength;
      chunk.fill(0);
    }
    completed = true;
    return bytes;
  } finally {
    if (!completed) {
      for (const chunk of chunks) chunk.fill(0);
    }
    reader.releaseLock?.();
  }
}

async function fetchBoundedPdfBytesAttempt(
  previewUrl: string,
  fetcher: PdfMutationFetcher,
): Promise<Uint8Array> {
  const controller = new AbortController();
  const timeoutError = new PdfMutationSourceTimeoutError();
  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  const timeout = new Promise<never>((_resolve, reject) => {
    timeoutId = setTimeout(() => {
      controller.abort();
      reject(timeoutError);
    }, PDF_PAGE_WINDOW_MUTATION_FETCH_TIMEOUT_MS);
  });
  const loading = (async () => {
    let response: PdfMutationFetchResponse;
    try {
      response = await fetcher(previewUrl, {
        credentials: "include",
        signal: controller.signal,
      });
    } catch (error) {
      if (controller.signal.aborted || isAbortError(error)) throw timeoutError;
      throw new PdfMutationSourceTransportError(error);
    }
    if (!response.ok) throw new PdfMutationSourceHttpError(response.status);
    try {
      return await readBoundedPdfResponse(response);
    } catch (error) {
      if (controller.signal.aborted || isAbortError(error)) throw timeoutError;
      if (error instanceof TypeError) {
        throw new PdfMutationSourceTransportError(error);
      }
      throw error;
    }
  })();
  try {
    return await Promise.race([loading, timeout]);
  } finally {
    if (timeoutId !== null) clearTimeout(timeoutId);
  }
}

async function fetchBoundedPdfBytes(
  previewUrl: string,
  fetcher: PdfMutationFetcher,
): Promise<Uint8Array> {
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return await fetchBoundedPdfBytesAttempt(previewUrl, fetcher);
    } catch (error) {
      const retryable =
        error instanceof PdfMutationSourceTimeoutError ||
        error instanceof PdfMutationSourceTransportError;
      if (!retryable || attempt === 1) throw error;
    }
  }
  throw new Error("Unreachable PDF source fetch state");
}

export interface PdfPageWindowMutationSource {
  readonly bytes: Uint8Array;
  readonly document: PdfLibDocument;
}

/**
 * Strict source admission for derivative PDF writers. Unlike the legacy
 * loader, no unbounded `arrayBuffer()` is allowed: raw responses stream under
 * the editable PDF cap and cached decrypted bytes are checked before cloning.
 */
export async function loadBoundedPdfForPageWindowMutation(
  previewUrl: string,
  fetcher: PdfMutationFetcher = fetch,
): Promise<PdfPageWindowMutationSource> {
  const decrypted = await waitForDecryptedPdfBytes(previewUrl);
  let bytes: Uint8Array;
  if (decrypted) {
    if (
      decrypted.byteLength < 1 ||
      decrypted.byteLength > MAX_PDF_PAGE_WINDOW_MUTATION_BYTES
    ) {
      throw new Error("PDF preview exceeds the editable PDF mutation limit");
    }
    bytes = new Uint8Array(decrypted.byteLength);
    bytes.set(new Uint8Array(decrypted));
  } else {
    bytes = await fetchBoundedPdfBytes(previewUrl, fetcher);
  }
  const { loadPdfForMutation } = await import("./encrypted-guard");
  const document = await loadPdfForMutation(bytes, { updateMetadata: false });
  return Object.freeze({ bytes, document });
}

/**
 * Load the current PageWindow source for a pdf-lib mutation.
 *
 * The default-off branch preserves the legacy request and load path exactly.
 * With V2 enabled, an unlock already accepted by PDF.js is joined to the
 * server-decrypted bytes for the same preview URL (including an in-flight
 * request), and raw fallback bytes pass through the shared encrypted guard.
 */
export async function loadPdfForPageWindowMutation(
  previewUrl: string,
  fetcher: PdfMutationFetcher = fetch,
): Promise<PdfLibDocument> {
  if (!pdfPageWindowMutationGuardV2Enabled()) {
    const { PDFDocument } = await import("pdf-lib");
    return PDFDocument.load(await fetchPdfBytes(previewUrl, fetcher));
  }

  const decrypted = await waitForDecryptedPdfBytes(previewUrl);
  const bytes = decrypted
    ? decrypted.slice(0)
    : await fetchPdfBytes(previewUrl, fetcher);
  const { loadPdfForMutation } = await import("./encrypted-guard");
  return loadPdfForMutation(bytes);
}
