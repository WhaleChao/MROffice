/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import { EngineUnavailableError } from "@mrnote/core/web/lib/engine-host/client.ts";
import { PdfNativeTransactionError } from "./engine/native/transaction-error";

/** The longest a retry waits for a busy server's Retry-After. */
export const MAX_PDF_PREPARATION_RETRY_DELAY_MS = 1_500;
/** The longest a retry after any other failure waits (it doubles from 0.5 s). */
export const MAX_PDF_PREPARATION_BACKOFF_MS = 4_000;
/** How long an open line keeps asking an engine that cannot answer. */
export const PDF_PREPARATION_PATIENCE_MS = 120_000;

/** Before retry `attempt` (0 for the first): a busy answer's Retry-After
 * (capped); the first retry of another failure at once, later ones after
 * 0.5 s doubling to 4 s. Resolves early (without throwing) when `signal`
 * aborts; the caller checks it. */
export function pdfPreparationRetryDelay(error: unknown, signal: AbortSignal, attempt = 0): Promise<void> {
  const wait = error instanceof EngineUnavailableError && (error.reason === "busy" || error.reason === "rate_limited")
    ? Math.min(MAX_PDF_PREPARATION_RETRY_DELAY_MS, Math.max(0, error.retryAfterMs ?? 1_000))
    : attempt === 0 ? 0 : Math.min(MAX_PDF_PREPARATION_BACKOFF_MS, 500 * 2 ** (attempt - 1));
  if (!wait || signal.aborted) return Promise.resolve();
  return new Promise((resolve) => {
    const timer = setTimeout(done, wait);
    function done() { clearTimeout(timer); signal.removeEventListener("abort", done); resolve(); }
    signal.addEventListener("abort", done, { once: true });
  });
}

/** The source's resource and version reads a line opens with, when the
 * server did not answer them in time or the connection dropped
 * (components/notebook/contents/pdf/window/nativeSourceAuthority.ts). */
const SOURCE_METADATA_UNANSWERED = new Set([
  "PdfNativeSourceMetadataTimeoutError",
  "PdfNativeSourceMetadataTransportError",
]);

export function isTransientPdfPreparationFailure(error: unknown): boolean {
  if (error instanceof DOMException && error.name === "AbortError") return false;
  if (error instanceof PdfNativeTransactionError) return false;
  if (error instanceof EngineUnavailableError) return error.reason !== "signed_out";
  if (error instanceof Error && SOURCE_METADATA_UNANSWERED.has(error.name)) return true;
  // A server too busy to answer (a full connection pool, a restart) is no
  // verdict on the line.
  if (isApiRequestError(error)) return error.status === 0 || error.status === 429 || error.status >= 502;
  return error instanceof Error && error.message === "PDF text preparation failed.";
}
