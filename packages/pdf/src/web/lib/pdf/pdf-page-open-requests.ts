"use client";













/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useState } from "react";

export interface PdfPageOpenRequest {
  readonly page: number;
  readonly requestId: number;
}

/** Highest physical page a request may name (the reader's own page bound). */
const MAX_PDF_PAGE = 10_000;
const EVENT_NAME = "mrai:pdf-page-open-request";

const pending = new Map<string, PdfPageOpenRequest>();
let sequence = 0;

/** A strict 1-based page number from a query value, else null. */
export function parsePdfPageNumber(value: unknown): number | null {
  if (typeof value !== "string" || !/^[1-9]\d{0,4}$/.test(value.trim())) {
    return null;
  }
  const page = Number(value.trim());
  return page <= MAX_PDF_PAGE ? page : null;
}

/** Ask the window showing `pageId` to go to physical page `page`. */
export function requestPdfPageOpenAt(pageId: string, page: number): void {
  if (!pageId || !Number.isSafeInteger(page) || page < 1 || page > MAX_PDF_PAGE) {
    return;
  }
  sequence += 1;
  const request: PdfPageOpenRequest = { page, requestId: sequence };
  pending.set(pageId, request);
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(EVENT_NAME, { detail: { pageId, request } }),
    );
  }
}

/** Take (and forget) the request waiting for `pageId`, if any. */
export function takePendingPdfPageOpenRequest(
  pageId: string,
): PdfPageOpenRequest | null {
  const request = pending.get(pageId) ?? null;
  pending.delete(pageId);
  return request;
}

/** The latest open-at-page request for the PDF page `pageId`. */
export function usePdfPageOpenRequest(
  pageId: string,
): PdfPageOpenRequest | null {
  // A request made before this window mounted is waiting for it. Read it
  // without consuming (an initializer may run twice); the effect forgets it.
  const [request, setRequest] = useState<PdfPageOpenRequest | null>(
    () => pending.get(pageId) ?? null,
  );
  useEffect(() => {
    pending.delete(pageId);
    if (typeof window === "undefined") return undefined;
    const onRequest = (event: Event) => {
      const detail = (event as CustomEvent<{
        pageId?: unknown;
        request?: PdfPageOpenRequest;
      }>).detail;
      if (detail?.pageId !== pageId || !detail.request) return;
      pending.delete(pageId);
      setRequest(detail.request);
    };
    window.addEventListener(EVENT_NAME, onRequest);
    return () => window.removeEventListener(EVENT_NAME, onRequest);
  }, [pageId]);
  return request;
}
