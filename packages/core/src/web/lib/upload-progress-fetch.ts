/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * `fetch` for a request that sends a file, with its upload on the transfer
 * line (lib/transfer-progress).
 *
 * `fetch` reports nothing while a request body goes out, so a page import,
 * an attachment or a picture showed no progress at all until the server had
 * answered. A request whose body carries a file (a Blob, a File, or a
 * FormData holding one) of at least UPLOAD_PROGRESS_MIN_BYTES goes through
 * XMLHttpRequest instead, whose `upload` events say how much has been sent;
 * everything else is the platform `fetch`, unchanged.
 *
 * The answer is a real `Response` built from the request's status, headers
 * and body, so callers read it exactly as they read fetch's. What fetch's
 * init asks for is kept: credentials, headers, the abort signal, and a
 * `redirect: "error"` request fails like fetch's when it was redirected.
 */

import { beginTransfer } from "./transfer-store";

/** A body smaller than this is on its way before a line could show. */
export const UPLOAD_PROGRESS_MIN_BYTES = 256 * 1024;

/** Bodies whose sent bytes a caller reports itself (one transfer for many
 * requests, such as the parts of a multipart upload). */
const observedBodies = new WeakMap<object, (sentBytes: number) => void>();

/**
 * Report this body's sent bytes to `onSent` instead of giving it a line of
 * its own. For a caller that already shows one transfer for several
 * requests and wants the bytes inside each request, not only at its end.
 */
export function reportUploadTo<T extends Blob | FormData>(body: T, onSent: (sentBytes: number) => void): T {
  observedBodies.set(body, onSent);
  return body;
}

/** Statuses whose response may not carry a body (the Fetch standard). */
const NULL_BODY_STATUSES = new Set([101, 103, 204, 205, 304]);

interface UploadBody {
  bytes: number;
  name: string;
}

function uploadBody(body: BodyInit | null | undefined): UploadBody | null {
  if (!body || typeof Blob === "undefined") return null;
  if (body instanceof Blob) {
    const name = typeof File !== "undefined" && body instanceof File ? body.name : "";
    return { bytes: body.size, name };
  }
  if (typeof FormData !== "undefined" && body instanceof FormData) {
    let bytes = 0;
    const names: string[] = [];
    let carriesFile = false;
    body.forEach((value) => {
      if (value instanceof Blob) {
        carriesFile = true;
        bytes += value.size;
        if (typeof File !== "undefined" && value instanceof File && value.name) names.push(value.name);
      } else {
        bytes += value.length;
      }
    });
    if (!carriesFile) return null;
    return { bytes, name: names.length > 1 ? `${names[0]} +${names.length - 1}` : (names[0] ?? "") };
  }
  return null;
}

function parseHeaders(raw: string): Headers {
  const headers = new Headers();
  for (const line of raw.trim().split(/[\r\n]+/)) {
    const at = line.indexOf(":");
    if (at <= 0) continue;
    try {
      headers.append(line.slice(0, at).trim(), line.slice(at + 1).trim());
    } catch {
      // a header the platform will not hold is left out, as fetch hides it
    }
  }
  return headers;
}

function abortError(signal: AbortSignal): unknown {
  return signal.reason ?? new DOMException("The operation was aborted", "AbortError");
}

function sendWithProgress(
  url: string,
  init: RequestInit,
  body: UploadBody,
  onSent: ((sentBytes: number) => void) | undefined,
): Promise<Response> {
  const signal = init.signal ?? null;
  if (signal?.aborted) return Promise.reject(abortError(signal));
  return new Promise<Response>((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    const transfer = onSent
      ? { update: onSent, end: () => undefined }
      : beginTransfer({
          direction: "upload",
          name: body.name || "upload",
          totalBytes: body.bytes,
        });
    let settled = false;
    const finish = (outcome: () => void) => {
      if (settled) return;
      settled = true;
      transfer.end();
      signal?.removeEventListener("abort", onAbort);
      outcome();
    };
    const onAbort = () => {
      xhr.abort();
      finish(() => reject(abortError(signal as AbortSignal)));
    };
    xhr.open((init.method || "GET").toUpperCase(), url, true);
    xhr.responseType = "blob";
    xhr.withCredentials = init.credentials === "include";
    new Headers(init.headers ?? undefined).forEach((value, key) => {
      xhr.setRequestHeader(key, value);
    });
    xhr.upload.onprogress = (event) => {
      transfer.update(event.lengthComputable ? event.loaded : Math.min(event.loaded, body.bytes));
    };
    xhr.upload.onload = () => transfer.update(body.bytes);
    xhr.onload = () => {
      if (init.redirect === "error" && xhr.responseURL && xhr.responseURL !== new URL(url, location.href).href) {
        finish(() => reject(new TypeError("Failed to fetch: the request was redirected")));
        return;
      }
      const status = xhr.status;
      const responseBody = NULL_BODY_STATUSES.has(status) ? null : (xhr.response as Blob | null);
      let response: Response;
      try {
        response = new Response(responseBody, {
          status,
          statusText: xhr.statusText,
          headers: parseHeaders(xhr.getAllResponseHeaders()),
        });
      } catch (error) {
        finish(() => reject(error));
        return;
      }
      finish(() => resolve(response));
    };
    xhr.onerror = () => finish(() => reject(new TypeError("Failed to fetch")));
    xhr.ontimeout = () => finish(() => reject(new TypeError("Failed to fetch")));
    xhr.onabort = () => {
      if (signal?.aborted) finish(() => reject(abortError(signal)));
      else finish(() => reject(new TypeError("Failed to fetch")));
    };
    signal?.addEventListener("abort", onAbort, { once: true });
    xhr.send(init.body as XMLHttpRequestBodyInit);
  });
}

/** `fetch`, or for a file upload of a useful size, the same request with
 * its progress on the transfer line. */
export function fetchWithUploadProgress(url: string, init: RequestInit = {}): Promise<Response> {
  const body = uploadBody(init.body);
  const onSent =
    init.body && typeof init.body === "object" ? observedBodies.get(init.body as object) : undefined;
  if (
    !body ||
    (body.bytes < UPLOAD_PROGRESS_MIN_BYTES && !onSent) ||
    typeof XMLHttpRequest === "undefined" ||
    init.keepalive
  ) {
    return fetch(url, init);
  }
  return sendWithProgress(url, init, body, onSent);
}
