/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Password-protected PDF support shared by the reader (PDFViewer) and the
 * native text editor (PdfTextEditor).
 *
 * pdf.js rejects `getDocument` for an encrypted document with a
 * `PasswordException` whose `code` distinguishes "no password was given"
 * (NEED_PASSWORD = 1) from "the given password is wrong"
 * (INCORRECT_PASSWORD = 2). The viewer retries `getDocument({ data,
 * password })` with the user-typed password, so pure viewing never needs a
 * server round-trip.
 *
 * Editing and export, however, run on pdf-lib — which cannot decrypt. After
 * a successful unlock the client asks the server to strip the encryption
 * (`POST /api/v1/ai/notebook/pdf-decrypt`, pikepdf) and caches the decrypted
 * bytes here PER SESSION, keyed by the same preview URL every byte consumer
 * already uses. The stored original is never overwritten.
 */

import { apiFetchRaw, ApiRequestError } from "@mrnote/core/web/lib/api.ts";

/** Mirrors pdf.js `PasswordResponses.NEED_PASSWORD`. */
export const PDF_PASSWORD_NEEDED = 1;
/** Mirrors pdf.js `PasswordResponses.INCORRECT_PASSWORD`. */
export const PDF_PASSWORD_INCORRECT = 2;

/** True when a pdf.js load rejection means "this document needs a password".
 * Matched by name (not instanceof) because pdfjs-dist is lazy-imported and
 * the exception can cross the worker boundary as a re-hydrated clone. */
export function isPdfPasswordException(err: unknown): boolean {
  return (
    typeof err === "object" &&
    err !== null &&
    (err as { name?: unknown }).name === "PasswordException"
  );
}

/** True when the rejection specifically means the SUPPLIED password was
 * wrong (vs. no password having been supplied at all). */
export function isIncorrectPdfPassword(err: unknown): boolean {
  return (
    isPdfPasswordException(err) &&
    (err as { code?: unknown }).code === PDF_PASSWORD_INCORRECT
  );
}

/** Error code the decrypt endpoint returns (403) for a wrong password. */
export const PDF_DECRYPT_WRONG_PASSWORD_CODE = "invalid_pdf_password";

/** True when a server decrypt failure means the password was wrong (so the
 * prompt should re-open) rather than an infrastructure problem. */
export function isServerWrongPdfPassword(err: unknown): boolean {
  return (
    err instanceof ApiRequestError &&
    err.code === PDF_DECRYPT_WRONG_PASSWORD_CODE
  );
}

/** Extract the data-item id from the app's canonical preview URL
 * (`…/api/v1/data-items/{id}/preview`). Returns null for any other shape —
 * callers then simply skip the server decrypt. */
export function dataItemIdFromPreviewUrl(url: string): string | null {
  const match = /\/api\/v1\/data-items\/([^/?#]+)\/preview(?:[?#]|$)/.exec(
    url,
  );
  return match ? match[1] : null;
}

// Per-session decrypted-bytes cache. Bounded like PdfTextEditor's
// pdfBytesCache: sessions touch a handful of documents, not hundreds.
const DECRYPTED_CACHE_LIMIT = 8;
const decryptedPdfBytes = new Map<string, ArrayBuffer>();
const pendingPdfDecrypts = new Map<string, Promise<ArrayBuffer>>();

/** Decrypted bytes for a preview URL unlocked earlier this session, if any. */
export function getDecryptedPdfBytes(url: string): ArrayBuffer | null {
  return decryptedPdfBytes.get(url) ?? null;
}

export function seedDecryptedPdfBytes(url: string, bytes: ArrayBuffer): void {
  decryptedPdfBytes.set(url, bytes);
  while (decryptedPdfBytes.size > DECRYPTED_CACHE_LIMIT) {
    const oldest = decryptedPdfBytes.keys().next().value;
    if (!oldest) break;
    decryptedPdfBytes.delete(oldest);
  }
}

/**
 * Return plaintext already cached for this preview, or wait for the latest
 * server decrypt that is still in flight. A failed decrypt yields `null` so a
 * caller can inspect its normal source and let the encrypted-byte guard make
 * the final decision; the password itself is never retained here.
 */
export async function waitForDecryptedPdfBytes(
  url: string,
): Promise<ArrayBuffer | null> {
  const cached = getDecryptedPdfBytes(url);
  if (cached) return cached;
  const pending = pendingPdfDecrypts.get(url);
  if (!pending) return null;
  try {
    return await pending;
  } catch {
    return null;
  }
}

/** Test seam: reset the session cache between vitest cases. */
export function clearDecryptedPdfBytesForTest(): void {
  decryptedPdfBytes.clear();
  pendingPdfDecrypts.clear();
}

/**
 * Ask the server to decrypt the stored PDF with the user's password and
 * cache the plaintext bytes for this session, so pdf-lib edit/export paths
 * (which cannot decrypt) and subsequent loads work on the unlocked bytes.
 *
 * Throws the `apiFetchRaw` error on failure (403 `invalid_pdf_password` for
 * a wrong password); callers decide whether that is fatal — after a
 * successful CLIENT-side unlock it is best-effort.
 */
export async function requestServerPdfDecrypt(
  dataItemId: string,
  password: string,
  previewUrl: string,
): Promise<ArrayBuffer> {
  const request = (async () => {
    const res = await apiFetchRaw(
      "/api/v1/ai/notebook/pdf-decrypt",
      {
        method: "POST",
        body: JSON.stringify({ data_item_id: dataItemId, password }),
      },
      { contentType: "application/json" },
    );
    const bytes = await res.arrayBuffer();
    seedDecryptedPdfBytes(previewUrl, bytes);
    return bytes;
  })();
  // Keep only the latest request for a URL. This avoids storing a password in
  // the registry key and lets a later successful retry supersede an older
  // failing request without the older finally-handler deleting the new one.
  pendingPdfDecrypts.set(previewUrl, request);
  try {
    return await request;
  } finally {
    if (pendingPdfDecrypts.get(previewUrl) === request) {
      pendingPdfDecrypts.delete(previewUrl);
    }
  }
}
