"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word files are written by the MRAI server, never in the browser.
 *
 * A saved page exports through the authenticated page export. A document that
 * is not a saved page -- the guest editor's, an unsaved draft's, each record
 * of a mail merge, the PDF viewer's annotation summary -- sends its editor
 * JSON (the canonical, storage-shaped body) to
 * `POST /api/v1/office/docx/document` and receives the file the saved page's
 * writer makes of it. Mail merge sends its records in batches to
 * `POST /api/v1/office/docx/documents` and receives a ZIP of their files.
 *
 * A signed-in caller asks with its session and CSRF token. The marketing
 * guest editor has no session (the CSRF bootstrap refuses it), so it asks
 * again without one: the endpoint serves an anonymous caller from an allowed
 * origin under a per-IP limit. Without a connection there is no Word file:
 * the caller says so (`isDocxExportNeedsConnectionError`) instead of failing
 * silently.
 */

import { apiFetchRaw, isApiRequestError } from "../api";
import type { RevisionExportPolicy } from "./revision-policy";

export const DOCX_DOCUMENT_PATH = "/api/v1/office/docx/document";
export const DOCX_DOCUMENTS_PATH = "/api/v1/office/docx/documents";

/** The batch route's per-request bound (`MAX_BATCH_DOCUMENTS`). */
export const MAX_DOCX_BATCH_DOCUMENTS = 25;
/**
 * A batch is packed until its JSON reaches this size; one larger document
 * still goes on its own (the single-document bound is the page envelope).
 */
export const MAX_DOCX_BATCH_BODY_CHARS = 8 * 1024 * 1024;

/** Why no Word file could be had from the server. */
export type DocxExportUnavailableReason =
  | "offline"
  | "signed_out"
  | "busy"
  | "failed";

/** The server could not be asked (or could not answer): nothing was refused. */
export class DocxExportUnavailableError extends Error {
  readonly reason: DocxExportUnavailableReason;

  constructor(reason: DocxExportUnavailableReason, detail?: string) {
    super(`Word export unavailable (${reason})${detail ? `: ${detail}` : ""}`);
    this.name = "DocxExportUnavailableError";
    this.reason = reason;
  }
}

/** The server's typed refusal of this document (the saved page's answer too). */
export class DocxExportRefusedError extends Error {
  readonly status: number;
  readonly code: string | null;
  readonly details: Record<string, unknown>;

  constructor(status: number, code: string | null, details: Record<string, unknown>, message: string) {
    super(message);
    this.name = "DocxExportRefusedError";
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

/** No connection to the server: reconnect, then export again. */
export function isDocxExportNeedsConnectionError(error: unknown): boolean {
  return error instanceof DocxExportUnavailableError && error.reason === "offline";
}

/** The document is larger than a page body may be. */
export function isDocxExportTooLargeError(error: unknown): boolean {
  return (
    error instanceof DocxExportRefusedError &&
    (error.status === 413 ||
      error.code === "payload_too_large" ||
      (error.code === "invalid_page_content" &&
        typeof error.details.reason === "string" &&
        error.details.reason.includes("budget")))
  );
}

function isAbort(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

function translate(error: unknown): Error {
  if (isAbort(error)) return error as Error;
  if (isApiRequestError(error)) {
    if (error.status === 0 || error.code === "network_unreachable") {
      return new DocxExportUnavailableError("offline");
    }
    if (error.status === 401 || error.status === 403) {
      return new DocxExportUnavailableError("signed_out");
    }
    if (error.status === 429 || error.status === 503) {
      return new DocxExportUnavailableError("busy", `${error.status} ${error.code ?? ""}`.trim());
    }
    if (error.status >= 400 && error.status < 500) {
      return new DocxExportRefusedError(error.status, error.code ?? null, error.details ?? {}, error.message);
    }
    return new DocxExportUnavailableError("failed", `${error.status} ${error.code ?? ""}`.trim());
  }
  return new DocxExportUnavailableError("failed", error instanceof Error ? error.message : undefined);
}

/** The CSRF bootstrap refused (no session), or the session itself did. */
function noSession(error: unknown): boolean {
  if (isAbort(error)) return false;
  if (isApiRequestError(error)) return error.status === 401;
  return error instanceof Error;
}

async function post(path: string, body: string, signal?: AbortSignal): Promise<Blob> {
  if (typeof navigator !== "undefined" && navigator.onLine === false) {
    throw new DocxExportUnavailableError("offline");
  }
  const ask = (signedIn: boolean) =>
    apiFetchRaw(
      path,
      { method: "POST", body, signal },
      {
        contentType: "application/json",
        // An export never bounces the page to the login screen.
        suppressUnauthorizedHandling: true,
        ...(signedIn ? {} : { requireCsrf: false }),
      },
    );
  try {
    let response: Response;
    try {
      response = await ask(true);
    } catch (error) {
      if (!noSession(error)) throw error;
      response = await ask(false);
    }
    return await response.blob();
  } catch (error) {
    throw translate(error);
  }
}

export interface DocxServerExportRequest {
  /** The canonical, storage-shaped editor JSON (no page wrappers). */
  document: unknown;
  title?: string;
  revisionPolicy?: RevisionExportPolicy;
  signal?: AbortSignal;
}

/** One document -> its .docx, written by the server. */
export async function requestServerDocx({
  document,
  title,
  revisionPolicy = "accept",
  signal,
}: DocxServerExportRequest): Promise<Blob> {
  const blob = await post(
    DOCX_DOCUMENT_PATH,
    JSON.stringify({
      document,
      title: (title ?? "").slice(0, 500),
      revision_policy: revisionPolicy === "preserve" ? "preserve" : "accept",
    }),
    signal,
  );
  return blob.type
    ? blob
    : new Blob([blob], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
}

/**
 * Many documents (mail merge) -> their .docx files, in order. Documents are
 * packed into requests of at most `MAX_DOCX_BATCH_DOCUMENTS` and about
 * `MAX_DOCX_BATCH_BODY_CHARS` of JSON; the requests go one after another.
 */
export async function requestServerDocxBatch(
  documents: ReadonlyArray<{ document: unknown; title?: string }>,
  options: { signal?: AbortSignal; onProgress?: (done: number) => void } = {},
): Promise<Blob[]> {
  const encoded = documents.map((item) =>
    JSON.stringify({ document: item.document, title: (item.title ?? "").slice(0, 500) }),
  );
  const chunks: string[][] = [];
  let current: string[] = [];
  let size = 0;
  for (const item of encoded) {
    if (
      current.length > 0 &&
      (current.length >= MAX_DOCX_BATCH_DOCUMENTS || size + item.length > MAX_DOCX_BATCH_BODY_CHARS)
    ) {
      chunks.push(current);
      current = [];
      size = 0;
    }
    current.push(item);
    size += item.length;
  }
  if (current.length > 0) chunks.push(current);

  const { default: JSZip } = await import("jszip");
  const files: Blob[] = [];
  for (const chunk of chunks) {
    const archive = await post(DOCX_DOCUMENTS_PATH, `{"documents":[${chunk.join(",")}]}`, options.signal);
    let zip: Awaited<ReturnType<typeof JSZip.loadAsync>>;
    try {
      zip = await JSZip.loadAsync(await archive.arrayBuffer());
    } catch (error) {
      throw new DocxExportUnavailableError("failed", error instanceof Error ? error.message : "malformed batch");
    }
    const names = Object.keys(zip.files)
      .filter((name) => /^\d+\.docx$/u.test(name))
      .sort((left, right) => Number.parseInt(left, 10) - Number.parseInt(right, 10));
    if (names.length !== chunk.length) {
      throw new DocxExportUnavailableError("failed", "the batch answer does not match the request");
    }
    for (const name of names) {
      files.push(
        new Blob([await zip.file(name)!.async("arraybuffer")], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
      );
    }
    options.onProgress?.(files.length);
  }
  return files;
}

export const CHART_WORKBOOK_XLSX_PATH = "/api/v1/office/xlsx/chart-workbook";

/**
 * A Docs chart's embedded workbook -> its .xlsx, written by the server with
 * the chart's own writer (the same package a Word export embeds for the
 * chart: formulas with their evaluated cache, and the chart binding). Same
 * callers, errors and guest fallback as the Word routes above.
 */
export async function requestServerChartWorkbookXlsx(
  workbook: unknown,
  options: { title?: string; signal?: AbortSignal } = {},
): Promise<Blob> {
  const blob = await post(
    CHART_WORKBOOK_XLSX_PATH,
    JSON.stringify({ workbook, title: (options.title ?? "ChartData").slice(0, 500) }),
    options.signal,
  );
  return blob.type
    ? blob
    : new Blob([blob], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
}

/**
 * A safe `.docx` filename from a document title: characters illegal in
 * filenames removed, whitespace collapsed, 120 characters at most.
 */
export function docxFilename(title: string | undefined, fallback = "document"): string {
  const base = String(title ?? "")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/gu, "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
  return `${base || fallback}.docx`;
}
