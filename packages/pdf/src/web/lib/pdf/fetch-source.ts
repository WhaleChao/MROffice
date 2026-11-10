/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { fetchWorkspaceScoped } from "@mrnote/core/web/lib/api.ts";
import { getApiHttpBaseUrl } from "@mrnote/core/web/lib/env.ts";
import { isPrivateFileUrl, fetchPrivateFile } from "@mrnote/core/web/lib/private-file-transfer.ts";
import { copyPdfWorkingDisplayBytes, type PdfWorkingDisplaySource } from "./working-display-source";










const workingSources = new Map<string, PdfWorkingDisplaySource>();

/** Answer `source.displayKey` with its bytes while the viewer shows it. */
export function registerPdfWorkingDisplaySource(source: PdfWorkingDisplaySource): () => void {
  workingSources.set(source.displayKey, source);
  return () => { if (workingSources.get(source.displayKey) === source) workingSources.delete(source.displayKey); };
}

async function fetchWorkingSource(key: string, init: RequestInit): Promise<Response> {
  if (init.signal?.aborted) throw new DOMException("Aborted", "AbortError");
  const source = workingSources.get(key);
  if (!source) return new Response(null, { status: 410, statusText: "Working revision is gone" });
  const bytes = await copyPdfWorkingDisplayBytes(source, source.dataItemId);
  return new Response(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    { status: 200, headers: { "content-type": "application/pdf", "content-length": String(bytes.byteLength) } });
}

/**
 * Fetch PDF source bytes without confusing application authorization with a
 * storage bearer URL.
 *
 * MRAI/API URLs need cookies plus the active mounted-workspace header, so they
 * stay on `fetchWorkspaceScoped`. External object-storage URLs authenticate by
 * their query signature; sending those credentials or headers makes Safari
 * issue a CORS preflight that OSS correctly refuses. Keep ordinary request
 * options such as AbortSignal/cache, but force a simple credential-free fetch.
 */
export function fetchPdfSource(
  url: string,
  init: RequestInit = {},
): Promise<Response> {
  if (url.startsWith("mrai-pdf-working:")) return fetchWorkingSource(url, init);
  if (typeof window === "undefined") {
    return fetchWorkspaceScoped(url, init);
  }
  try {
    const target = new URL(url, window.location.origin);
    if (isPrivateFileUrl(target.href, init.method || "GET")) return fetchPrivateFile(target.href, init);
    const apiOrigin = new URL(getApiHttpBaseUrl(), window.location.origin)
      .origin;
    if (
      target.origin === window.location.origin ||
      target.origin === apiOrigin
    ) {
      return fetchWorkspaceScoped(target.toString(), init);
    }
    const externalInit: RequestInit = { ...init };
    delete externalInit.credentials;
    delete externalInit.headers;
    return fetch(target.toString(), {
      ...externalInit,
      credentials: "omit",
    });
  } catch {
    return Promise.reject(new Error("Invalid PDF source URL"));
  }
}
