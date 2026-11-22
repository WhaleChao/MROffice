"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { embedExportWatermark, getEmbedAccess } from "./embed/session-access";
import { markServerStamped } from "./embed/watermark-pdf";
import { SpreadsheetPdfViewError, assertSpreadsheetPdfViewBudget, type SpreadsheetPdfExportView } from "./spreadsheet/pdf-export-view";

import { apiFetchRaw, isApiRequestError } from "./api";
import { cachedDocsLayoutReceipt } from "./docs-layout-receipt-cache";
import { CONTENT_LOCALE_QUERY, requestContentLocale } from "./i18n/content-locale";
import { downloadBlobFile, downloadUrlFile, startDownloadWhenAllowed } from "./download-file";
import { beginTransfer, readResponseWithProgress } from "./transfer-store";
import { fetchPrivateFile } from "./private-file-transfer";

export type PageExportFormat =
  | "markdown"
  | "md"
  | "text"
  | "txt"
  | "html"
  | "pdf"
  | "docx"
  | "csv"
  | "xlsx"
  | "pptx"
  // OpenDocument, written by the server from the page's own xlsx/docx/pptx
  // export (`file_adapters/office_conversion`).
  | "ods"
  | "odt"
  | "odp"
  // Word's other Save As types, written by the server from the document's
  // own docx export: Word 97-2003, Rich Text and Word Template.
  | "doc"
  | "rtf"
  | "dotx"
  // Word's Word XML Document (a flat package, downloaded as .xml).
  | "wordxml"
  // Excel 97-2003 workbook, written by the server from the page's xlsx export.
  | "xls"
  // Excel Template (a macro workbook's: .xltm) and OpenDocument Spreadsheet
  // Flat XML, written by the server from the page's xlsx export.
  | "xltx"
  | "fods"
  // Excel's XML Spreadsheet 2003 (one SpreadsheetML file, downloaded as .xml).
  | "excelxml"
  // Excel's SYLK and DIF: the active sheet, in the reader's ANSI code page.
  | "slk"
  | "dif"
  // PowerPoint's other Save As types, written by the server from the deck's
  // own pptx export: PowerPoint 97-2003 Presentation, PowerPoint Design
  // Template and PowerPoint Show.
  | "ppt"
  | "potx"
  | "ppsx";

/**
 * A PDF the editor could not print, drawn the older server way instead
 * (`X-MRAI-PDF-Layout: fallback; reason=<codes>`, pdf_layout_receipt on the
 * API): the file may differ from the page, and the shell says so beside it
 * (components/console/PdfLayoutNotice).
 */
export const PDF_LAYOUT_FALLBACK_EVENT = "mrai:pdf-layout-fallback";

export interface PdfLayoutFallbackDetail {
  readonly pageId: string;
  /** Short reason codes (`print_failed`, `browser_unavailable`, ...). */
  readonly reasons: readonly string[];
}

/** The reason codes a `X-MRAI-PDF-Layout` value names, or null when the
 * editor printed the PDF (or the server did not say). */
export function pdfLayoutFallbackReasons(header: string | null): string[] | null {
  const value = (header ?? "").trim();
  if (!value.toLowerCase().startsWith("fallback")) return null;
  const named = /reason=([^;]*)/i.exec(value)?.[1] ?? "";
  const reasons = named
    .split(",")
    .map((reason) => reason.trim())
    .filter((reason) => /^[a-z0-9_.-]{1,64}$/i.test(reason));
  return reasons.length ? [...new Set(reasons)] : ["unknown"];
}

function announcePdfLayout(response: Response, pageId: string): void {
  const reasons = pdfLayoutFallbackReasons(response.headers.get("x-mrai-pdf-layout"));
  if (!reasons || typeof window === "undefined") return;
  const detail: PdfLayoutFallbackDetail = { pageId, reasons };
  window.dispatchEvent(new CustomEvent(PDF_LAYOUT_FALLBACK_EVENT, { detail }));
}

export type PageExportRevisionPolicy = "accept" | "reject" | "preserve";
/** The review display a document PDF showing markup prints its tracked
 * changes in (Word's Display for Review / Show Markup: the screen's). */
export type PageExportMarkupView =
  | "simple"
  | "all"
  | "insertions"
  | "deletions"
  | "formatting";
export type PdfEditableProjectionTarget =
  | "document"
  | "spreadsheet"
  | "presentation";

export function pageExportFormatForPdfProjection(
  target: PdfEditableProjectionTarget,
): "docx" | "xlsx" | "pptx" {
  if (target === "spreadsheet") return "xlsx";
  if (target === "presentation") return "pptx";
  return "docx";
}

export interface PageExportRequest {
  pageId: string;
  format: PageExportFormat;
  layout?: string;
  revisionPolicy?: PageExportRevisionPolicy;
  /** A document PDF with `revisionPolicy: "preserve"` (Word's "Document
   * showing markup"): the review display its markup prints in. */
  markupView?: PageExportMarkupView;
  /** Exact persisted page generation the caller authorized for this export. */
  expectedContentVersion?: number;
  /** Ephemeral source-row view; never replaces persisted cells. PDF only. */
  spreadsheetView?: SpreadsheetPdfExportView;
  /**
   * Review > Notes > Show All Notes is on in this session. It is Excel's
   * display setting, not workbook data, so a spreadsheet PDF whose sheets
   * print notes "As displayed on sheet" learns it from the request. PDF only.
   */
  showAllNotes?: boolean;
  /**
   * The workbook PDF is the editor's Print: made as Windows Excel's File >
   * Print makes it (every spilled text's tail carried onto a later page
   * prints there), not as its Save as PDF. Kept apart from the export's PDF
   * on the server. PDF only.
   */
  printRoute?: boolean;
  /**
   * A deck PDF's hidden slides (`hidden_slides`). Absent, the export's
   * default: left out, as PowerPoint's Save as PDF / Export > Create PDF/XPS
   * leaves them out (Options > Include hidden slides, off). The editor's
   * Print says `include` when its print falls back to this PDF (Print Hidden
   * Slides is on by default in PowerPoint). PDF only.
   */
  hiddenSlides?: "include" | "exclude";
  /** Explicit user consent for a server-declared lossy retry of this format. */
  lossyConsent?: boolean;
  /** Cancels request setup/fetch without becoming part of the export URL. */
  signal?: AbortSignal;
  /**
   * Encrypt with Password (DOCX, XLSX): the downloaded package is
   * encrypted with this password (lib/office/office-encryption) before it is
   * saved. Never part of the export request: only the encryption request's
   * body carries it.
   */
  openPassword?: string;
  /**
   * Save as a fillable PDF form (format pdf): the server prints the fields'
   * places (`layout=fillable`) and the file is made a form here
   * (lib/document-export/fillable-pdf) before it is saved.
   */
  fillableForm?: boolean;
  /**
   * The file is printed, not saved (an embedded session's `print`
   * permission admits it where `download` does not): `purpose=print`.
   */
  purpose?: "print";
}

export const PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE =
  "pptx_export_loss_confirmation_required";
export const DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE =
  "document_docx_final_copy_loss_confirmation_required";

/**
 * Each format's loss gate raises exactly one refusal code. Matching the PAIR
 * rather than the code alone is what keeps a refusal raised for one format
 * from ever being answered with consent recorded for another.
 */
/** An imported workbook whose original package cannot carry the current
 * version (a worksheet the package has no part for): the rebuilt workbook
 * exports once the user accepts what the package alone held. */
export const XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE =
  "xlsx_export_loss_confirmation_required";

const EXPORT_LOSS_CONFIRMATION_CODE: Partial<
  Record<PageExportFormat, string>
> = {
  pptx: PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  docx: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  xlsx: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  // OpenDocument is the Office export converted: its loss gate is that one's.
  ods: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  odt: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  odp: PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  doc: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  rtf: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  dotx: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  wordxml: DOCX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  xls: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  xltx: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  fods: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  excelxml: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  slk: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  dif: XLSX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  ppt: PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  potx: PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
  ppsx: PPTX_EXPORT_LOSS_CONFIRMATION_REQUIRED_CODE,
};

const MAX_EXPORT_LOSS_CODES = 40;
const MAX_EXPORT_LOSS_COUNT = 200;
const SAFE_EXPORT_LOSS_CODE = /^[A-Za-z0-9][A-Za-z0-9_.:-]{0,95}$/;

/**
 * A stored file the person can still have when an export cannot be built
 * from the current version (Office's previous version / the file as last
 * saved): `GET /api/v1/resources/{resource}/versions/{version}/download`
 * answers its short-lived URL.
 */
export interface PageExportLastGoodCopy {
  readonly kind: "last_export" | "original";
  readonly resourceId: string;
  readonly versionId: string;
  /** ISO time it was saved, or null when the server did not say. */
  readonly createdAt: string | null;
  readonly sizeBytes: number | null;
}

const SAFE_ID = /^[A-Za-z0-9-]{1,64}$/;
const MAX_LAST_GOOD_COPIES = 8;

/** The `details.last_good_copies` of an export refusal, validated. */
export function readLastGoodCopies(details: unknown): PageExportLastGoodCopy[] {
  const raw = details && typeof details === "object" ? (details as { last_good_copies?: unknown }).last_good_copies : null;
  if (!Array.isArray(raw)) return [];
  const copies: PageExportLastGoodCopy[] = [];
  for (const entry of raw.slice(0, MAX_LAST_GOOD_COPIES)) {
    if (!entry || typeof entry !== "object") continue;
    const value = entry as Record<string, unknown>;
    const kind = value.kind === "original" ? "original" : value.kind === "last_export" ? "last_export" : null;
    if (!kind || typeof value.resource_id !== "string" || typeof value.version_id !== "string" ||
        !SAFE_ID.test(value.resource_id) || !SAFE_ID.test(value.version_id)) continue;
    const createdAt = typeof value.created_at === "string" && !Number.isNaN(Date.parse(value.created_at))
      ? value.created_at : null;
    const sizeBytes = typeof value.size_bytes === "number" && Number.isSafeInteger(value.size_bytes) ? value.size_bytes : null;
    copies.push({ kind, resourceId: value.resource_id, versionId: value.version_id, createdAt, sizeBytes });
  }
  return copies;
}

export class PageExportLossConfirmationError extends Error {
  readonly lossCodes: readonly string[];
  readonly lastGoodCopies: readonly PageExportLastGoodCopy[];

  constructor(
    readonly code: string,
    lossCodes: readonly string[],
    readonly lossCount: number,
    lastGoodCopies: readonly PageExportLastGoodCopy[] = [],
  ) {
    super(code);
    this.name = "PageExportLossConfirmationError";
    this.lossCodes = Object.freeze([...lossCodes]);
    this.lastGoodCopies = Object.freeze([...lastGoodCopies]);
  }
}

export function isPageExportLossConfirmationError(
  error: unknown,
): error is PageExportLossConfirmationError {
  return error instanceof PageExportLossConfirmationError;
}

function toPageExportLossConfirmationError(
  error: unknown,
  format: PageExportFormat,
): PageExportLossConfirmationError | null {
  const expectedCode = EXPORT_LOSS_CONFIRMATION_CODE[format];
  if (
    expectedCode === undefined ||
    !isApiRequestError(error) ||
    error.status !== 409 ||
    error.code !== expectedCode
  ) {
    return null;
  }
  const details = error.details;
  
  // person can have instead; nothing else.
  const keys = details ? Object.keys(details).filter((key) => key !== "last_good_copies") : [];
  if (!details || keys.sort().join(",") !== "loss_codes,loss_count") {
    return null;
  }
  const lossCodes = details.loss_codes;
  const lossCount = details.loss_count;
  if (
    !Array.isArray(lossCodes) ||
    lossCodes.length === 0 ||
    lossCodes.length > MAX_EXPORT_LOSS_CODES ||
    typeof lossCount !== "number" ||
    !Number.isSafeInteger(lossCount) ||
    lossCount < lossCodes.length ||
    lossCount > MAX_EXPORT_LOSS_COUNT
  ) {
    return null;
  }
  const uniqueCodes = new Set<string>();
  for (const code of lossCodes) {
    if (
      typeof code !== "string" ||
      !SAFE_EXPORT_LOSS_CODE.test(code) ||
      uniqueCodes.has(code)
    ) {
      return null;
    }
    uniqueCodes.add(code);
  }
  return new PageExportLossConfirmationError(
    expectedCode,
    [...uniqueCodes],
    lossCount,
    readLastGoodCopies(details),
  );
}

const FORMAT_EXTENSION: Record<PageExportFormat, string> = {
  markdown: "md",
  md: "md",
  text: "txt",
  txt: "txt",
  html: "html",
  pdf: "pdf",
  docx: "docx",
  csv: "csv",
  xlsx: "xlsx",
  pptx: "pptx",
  ods: "ods",
  odt: "odt",
  odp: "odp",
  doc: "doc",
  rtf: "rtf",
  dotx: "dotx",
  wordxml: "xml",
  xls: "xls",
  xltx: "xltx",
  fods: "fods",
  excelxml: "xml",
  slk: "slk",
  dif: "dif",
  ppt: "ppt",
  potx: "potx",
  ppsx: "ppsx",
};

const MAX_DOWNLOAD_FILENAME_LENGTH = 255;
const UNSAFE_FILENAME_FORMATTING =
  /[\u0000-\u001f\u007f-\u009f\u200e\u200f\u202a-\u202e\u2066-\u2069]/g;

/**
 * The other extension the server may name a format's file with: a deck
 * that holds macros downloads as a .pptm (PowerPoint's rule: a .pptx is
 * macro-free). Only this alternate is kept; any other name still gets the
 * format's own extension.
 */
const SERVER_ALTERNATE_EXTENSIONS: Readonly<Record<string, readonly string[]>> = {
  pptx: ["pptm"],
  // A Design Template or Show of a deck holding macros is its macro-enabled twin.
  potx: ["potm"],
  ppsx: ["ppsm"],
  // An Excel Template of a workbook holding macros is Excel's macro-enabled template.
  xltx: ["xltm"],
};

function safeDownloadFilename(
  value: string,
  fallback: string,
  extension: string,
  alternates: readonly string[] = [],
): string {
  let normalized = value
    .normalize("NFC")
    .replace(UNSAFE_FILENAME_FORMATTING, "")
    .replace(/[\\/]/g, "_")
    .trim()
    .replace(/[. ]+$/g, "");
  if (!normalized || normalized === "." || normalized === "..") {
    return fallback;
  }
  const lowered = normalized.toLowerCase();
  const alternate = alternates.find((candidate) => lowered.endsWith(`.${candidate}`));
  const suffix = `.${(alternate ?? extension).toLowerCase()}`;
  if (!normalized.toLowerCase().endsWith(suffix)) normalized += suffix;
  const maxStemLength = MAX_DOWNLOAD_FILENAME_LENGTH - suffix.length;
  const stem = normalized.slice(0, -suffix.length).slice(0, maxStemLength);
  return `${stem || "page-export"}${suffix}`;
}

/**
 * The download name of a file the server named (its Content-Disposition,
 * a download ticket): the format's extension, or the one alternate the
 * server may give it ({@link SERVER_ALTERNATE_EXTENSIONS}): a deck holding
 * macros is saved as `deck.pptm`, never `deck.pptm.pptx`.
 */
export function serverDownloadFilename(value: string, fallback: string, extension: string): string {
  return safeDownloadFilename(value, fallback, extension, SERVER_ALTERNATE_EXTENSIONS[extension.toLowerCase()] ?? []);
}

/** The download name an export the browser writes itself gives a page. */
export function pageExportFilename(title: string, format: PageExportFormat): string {
  const extension = FORMAT_EXTENSION[format];
  return safeDownloadFilename(title, `page-export.${extension}`, extension);
}

function contentDispositionFilename(
  contentDisposition: string | null,
  fallback: string,
  extension: string,
): string {
  if (!contentDisposition) return fallback;

  const encoded = /(?:^|;)\s*filename\*\s*=\s*UTF-8''([^;]+)/i.exec(
    contentDisposition,
  )?.[1];
  if (encoded) {
    try {
      return serverDownloadFilename(
        decodeURIComponent(encoded.trim()),
        fallback,
        extension,
      );
    } catch {
      // A malformed extended filename must not abort an otherwise valid export;
      // continue to the ASCII fallback emitted by the API.
    }
  }

  const quoted = /(?:^|;)\s*filename\s*=\s*"((?:\\.|[^"])*)"/i.exec(
    contentDisposition,
  )?.[1];
  if (quoted) {
    return serverDownloadFilename(
      quoted.replace(/\\([\\"])/g, "$1"),
      fallback,
      extension,
    );
  }
  const token = /(?:^|;)\s*filename\s*=\s*([^;]+)/i.exec(
    contentDisposition,
  )?.[1];
  return token
    ? serverDownloadFilename(token.trim(), fallback, extension)
    : fallback;
}

interface SignedDownloadTicket {
  download_url: string;
  filename: string;
}

function signedDownloadTicket(value: unknown): SignedDownloadTicket | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  return typeof record.download_url === "string" &&
    typeof record.filename === "string"
    ? { download_url: record.download_url, filename: record.filename }
    : null;
}

function isJsonResponse(response: Pick<Response, "headers">): boolean {
  const mediaType = response.headers
    .get("content-type")
    ?.split(";", 1)[0]
    ?.trim()
    .toLowerCase();
  return (
    mediaType === "application/json" || Boolean(mediaType?.endsWith("+json"))
  );
}

/**
 * Export a native page through the authenticated mutation boundary.
 *
 * apiFetchRaw supplies credentials, workspace scoping, and the session-bound
 * CSRF header for POST. In-memory exports arrive as bytes. Existing private
 * objects arrive as a short-lived signed-URL ticket and are downloaded with a
 * plain anchor navigation, keeping API credentials/custom headers off the
 * object-store request and removing any bucket-CORS dependency.
 */
/** Content-Type is CORS-visible; refuse older services that ignored the view body. */
function hasSpreadsheetPdfViewReceipt(header: string | null): boolean {
  let remaining = header?.trim() ?? "";
  const media = /^application\/pdf(?=\s*(?:;|$))/i.exec(remaining);
  if (!media) return false;
  remaining = remaining.slice(media[0].length).trimStart();
  const values: string[] = [];
  while (remaining) {
    const parameter = /^;\s*([!#$%&'*+.^_`|~\w-]+)\s*=\s*(?:"((?:[^"\\]|\\.)*)"|([!#$%&'*+.^_`|~\w-]+))\s*(?=;|$)/.exec(remaining);
    if (!parameter) return false;
    if (parameter[1].toLowerCase() === "mrnote-spreadsheet-view") {
      values.push(parameter[2] !== undefined ? parameter[2].replace(/\\(.)/g, "$1") : parameter[3]);
    }
    remaining = remaining.slice(parameter[0].length);
  }
  return values.length === 1 && values[0] === "1";
}

/** How long the export request itself may wait before the server answers
 * with a job (it must stay under the edge's response timeout). */
const PAGE_EXPORT_JOB_WAIT_SECONDS = 20;
const PAGE_EXPORT_JOB_POLL_MS = [1000, 2000, 3000, 5000];
/** Transient answers while collecting a job (the edge or a restarting
 * process): asked again, never reported as the export failing. */
const PAGE_EXPORT_JOB_TRANSIENT_STATUSES = new Set([0, 408, 425, 429, 500, 502, 503, 504, 520, 521, 522, 523, 524]);
const PAGE_EXPORT_JOB_MAX_TRANSIENT = 30;
/** A job the poll's process does not hold yet (the job store is per process
 * without Redis, and the poll may reach another one): asked again a few
 * times before the 404 is believed. */
const PAGE_EXPORT_JOB_MAX_NOT_FOUND = 4;
/** The Word file the saves write is still being written past the server's
 * wait: the export is asked again, the transfer line still preparing, never
 * written from the page with losses to confirm. */
const PAGE_EXPORT_PACKAGE_STILL_WRITING = "docx_package_still_writing";
/** No worker takes the Word file's background writer at all: the server's
 * clear refusal, shown at once (a 503, but never asked again as transient). */
const PAGE_EXPORT_PACKAGE_WRITER_UNAVAILABLE = "docx_package_writer_unavailable";
const PAGE_EXPORT_PACKAGE_RETRY_MS = 10_000;
const PAGE_EXPORT_PACKAGE_MAX_RETRIES = 30;

function packageStillWritingRetryMs(error: unknown): number | null {
  if (!isApiRequestError(error) || error.code !== PAGE_EXPORT_PACKAGE_STILL_WRITING) return null;
  const seconds = Number(error.details?.retry_after);
  return Number.isFinite(seconds) && seconds > 0 ? Math.min(seconds, 60) * 1000 : PAGE_EXPORT_PACKAGE_RETRY_MS;
}

/**
 * The page's newer saved version an export was refused for
 * (`page_version_conflict`: a save landed after the version the export
 * pinned, or while the server was writing the file), or null. The server
 * names it `server_content_version` before rendering and
 * `current_content_version` once the file was written.
 */
export function pageExportNewerContentVersion(error: unknown): number | null {
  if (error instanceof SpreadsheetPdfViewError) return error.newerContentVersion ?? null;
  if (!isApiRequestError(error) || error.code !== "page_version_conflict") return null;
  const version = Number(error.details?.server_content_version ?? error.details?.current_content_version);
  return Number.isSafeInteger(version) && version >= 1 ? version : null;
}

interface PageExportJobTicket {
  job_id: string;
  poll_url: string;
  filename?: string | null;
}

function pageExportJobTicket(value: unknown): PageExportJobTicket | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (
    typeof record.job_id !== "string" ||
    !/^[0-9a-f]{32}$/.test(record.job_id) ||
    typeof record.poll_url !== "string" ||
    !record.poll_url.startsWith("/api/v1/pages/")
  ) {
    return null;
  }
  return {
    job_id: record.job_id,
    poll_url: record.poll_url,
    filename: typeof record.filename === "string" ? record.filename : null,
  };
}

function waitForPoll(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(signal.reason ?? new DOMException("Aborted", "AbortError"));
      return;
    }
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    const onAbort = () => {
      clearTimeout(timer);
      reject(signal?.reason ?? new DOMException("Aborted", "AbortError"));
    };
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

/**
 * The response of an export the server is still writing (202 + job): asked
 * for until it is the export's own answer -- the file, the download ticket,
 * or its refusal, raised exactly as the request itself would have raised it.
 * While it runs the transfer line shows it preparing; a dropped or timed-out
 * poll is asked again, so a slow export is never reported failed while the
 * server is still producing it. Only the server saying the job stopped
 * (``page_export_interrupted``) ends it as a failure.
 */
async function collectPageExportJob(
  accepted: Response,
  signal?: AbortSignal,
  acceptLink = false,
): Promise<Response> {
  const ticket = pageExportJobTicket(await accepted.json().catch(() => null));
  if (!ticket) throw new Error("The export job ticket is invalid.");
  // A caller that can take a download link has a large finished file handed
  // to it as one (the server no longer reads it back into memory to proxy it).
  const pollUrl = acceptLink
    ? `${ticket.poll_url}${ticket.poll_url.includes("?") ? "&" : "?"}delivery=link`
    : ticket.poll_url;
  const transfer = beginTransfer({
    direction: "download",
    name: ticket.filename || "export",
    totalBytes: 0,
    preparing: true,
  });
  const started = Date.now();
  let attempt = 0;
  let transient = 0;
  let notFound = 0;
  try {
    for (;;) {
      await waitForPoll(PAGE_EXPORT_JOB_POLL_MS[Math.min(attempt, PAGE_EXPORT_JOB_POLL_MS.length - 1)], signal);
      attempt += 1;
      transfer.preparing((Date.now() - started) / 1000);
      let response: Response;
      try {
        response = await apiFetchRaw(pollUrl, { method: "GET", ...(signal ? { signal } : {}) });
      } catch (error) {
        if (signal?.aborted) throw error;
        const status = isApiRequestError(error) ? (error.status ?? 0) : 0;
        const code = isApiRequestError(error) ? error.code : undefined;
        if (
          code !== "page_export_interrupted" &&
          code !== PAGE_EXPORT_PACKAGE_STILL_WRITING &&
          code !== PAGE_EXPORT_PACKAGE_WRITER_UNAVAILABLE &&
          PAGE_EXPORT_JOB_TRANSIENT_STATUSES.has(status) &&
          transient < PAGE_EXPORT_JOB_MAX_TRANSIENT
        ) {
          transient += 1;
          continue;
        }
        if (status === 404 && notFound < PAGE_EXPORT_JOB_MAX_NOT_FOUND) {
          notFound += 1;
          continue;
        }
        throw error;
      }
      transient = 0;
      notFound = 0;
      if (response.status !== 202) return response;
    }
  } finally {
    transfer.end();
  }
}

/** The receipt's bound a stored page holds it to (the API refuses more). */
const LAYOUT_RECEIPT_MAX_BYTES = 2 * 1024 * 1024;

/** A kept layout receipt as the export body carries it: its four parts, within the API's bound. */
export function exportableLayoutReceipt(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object") return null;
  const { version, canonicalHash, pages, geometry } = value as Record<string, unknown>;
  if (version !== 1 || typeof canonicalHash !== "string" || !Array.isArray(pages) || !Array.isArray(geometry)) return null;
  const receipt = { version, canonicalHash, pages, geometry };
  try {
    return new TextEncoder().encode(JSON.stringify(receipt)).length <= LAYOUT_RECEIPT_MAX_BYTES ? receipt : null;
  } catch {
    return null;
  }
}

/** The export's URL and request (query, body, headers) as every caller sends it. */
function pageExportRequest(
  request: PageExportRequest,
  delivery: "link" | "bytes",
): { path: string; init: RequestInit } {
  const pageId = request.pageId.trim();
  if (!pageId) throw new Error("A page id is required for export.");
  const expectedContentVersion = request.expectedContentVersion;
  if (
    expectedContentVersion !== undefined &&
    (!Number.isFinite(expectedContentVersion) ||
      !Number.isSafeInteger(expectedContentVersion) ||
      expectedContentVersion < 1)
  ) {
    throw new Error(
      "The expected content version must be a positive safe integer.",
    );
  }

  const query = new URLSearchParams({ format: request.format });
  // A document PDF showing markup names its review display as what the PDF
  // publishes (Word's Save as PDF "Publish what", a deck's layout).
  const layout = request.markupView && request.format === "pdf" && request.revisionPolicy === "preserve"
    ? `markup-${request.markupView}`
    : request.layout;
  if (layout) query.set("layout", layout);
  if (request.revisionPolicy) {
    query.set("revision_policy", request.revisionPolicy);
  }
  if (expectedContentVersion !== undefined) {
    query.set("expected_content_version", String(expectedContentVersion));
  }
  if (request.lossyConsent === true) {
    query.set("lossy_consent", "true");
  }
  if (request.showAllNotes === true && request.format === "pdf") {
    query.set("show_all_notes", "true");
  }
  // An embedded session tells a print from a download (its permissions
  // differ); the workspace's requests are unchanged.
  if (request.purpose === "print" && request.format === "pdf" && getEmbedAccess() !== null) {
    query.set("purpose", "print");
  }
  if (request.printRoute === true && request.format === "pdf") {
    query.set("print_route", "true");
  }
  if (request.hiddenSlides && request.format === "pdf") {
    query.set("hidden_slides", request.hiddenSlides);
  }
  if (delivery === "bytes") query.set("delivery", "bytes");
  // The caller's word that this file is encrypted before it is saved.
  if (request.openPassword && (request.format === "docx" || request.format === "xlsx")) {
    query.set("client_encrypts", "true");
  }
  // The words the export writes into the file (an empty deck's "Untitled",
  // a chart's "Series 1") are the reader's. A query parameter, not the
  // header the other API requests carry: the export can go through the
  // file-transfer host, whose preflight allows a fixed header list.
  const contentLocale = requestContentLocale();
  if (contentLocale) query.set(CONTENT_LOCALE_QUERY, contentLocale);
  const path = `/api/v1/pages/${encodeURIComponent(pageId)}/export?${query.toString()}`;
  const requestInit: RequestInit = { method: "POST" };
  if (request.spreadsheetView) {
    if (request.format !== "pdf" || expectedContentVersion === undefined) throw new SpreadsheetPdfViewError("source_changed");
    assertSpreadsheetPdfViewBudget(request.spreadsheetView);
    requestInit.headers = { "Content-Type": "application/json" };
    requestInit.body = JSON.stringify({ spreadsheetView: request.spreadsheetView });
  } else if (request.format === "pdf") {
    // A document's pages as this editor laid them out: the print restores
    // them onto the content they name instead of laying a long book out again
    // in the server's browser (minutes for a 271-page book).
    const documentLayoutReceipt = exportableLayoutReceipt(cachedDocsLayoutReceipt(pageId));
    if (documentLayoutReceipt) {
      requestInit.headers = { "Content-Type": "application/json" };
      requestInit.body = JSON.stringify({ documentLayoutReceipt });
    }
  }
  if (request.signal) requestInit.signal = request.signal;
  return { path, init: requestInit };
}

/** POST the export and check the spreadsheet view receipt; the caller reads the body. */
async function requestPageExport(
  request: PageExportRequest,
  /**
   * "link": the caller saves the file by navigating, so a stored file (an
   * imported original, a large finished job) may come as a signed download
   * ticket. "bytes": the caller reads the file in the page (an archive, a
   * print preview), which a private link cannot serve (the file gate sends
   * no CORS headers and wants cookies a fetch without credentials does not
   * carry): the API streams the stored file itself.
   */
  delivery: "link" | "bytes",
): Promise<Response> {
  const { path, init: requestInit } = pageExportRequest(request, delivery);
  const pageId = request.pageId.trim();
  // A long export (a book's Word file) outlasts the edge's response timeout:
  // the server answers 202 with a job instead, and the export's own response
  // is collected from it (``collectPageExportJob``). A quick one answers as
  // before.
  requestInit.headers = {
    ...(requestInit.headers as Record<string, string> | undefined),
    Prefer: `respond-async, wait=${PAGE_EXPORT_JOB_WAIT_SECONDS}`,
  };
  let response: Response;
  try {
    for (let retries = 0; ; retries += 1) {
      try {
        response = await apiFetchRaw(path, requestInit);
        if (response.status === 202) {
          response = await collectPageExportJob(response, request.signal, delivery === "link");
        }
        break;
      } catch (error) {
        const retryMs = packageStillWritingRetryMs(error);
        if (retryMs === null || retries >= PAGE_EXPORT_PACKAGE_MAX_RETRIES || request.signal?.aborted) throw error;
        await waitForPoll(retryMs, request.signal);
      }
    }
  } catch (error) {
    if (request.spreadsheetView && isApiRequestError(error) && ["page_version_conflict", "spreadsheet_export_view_invalid"].includes(error.code ?? "")) {
      throw new SpreadsheetPdfViewError("source_changed", pageExportNewerContentVersion(error) ?? undefined);
    }
    if (request.spreadsheetView && isApiRequestError(error) && error.code === "spreadsheet_pdf_view_too_large") throw new SpreadsheetPdfViewError("too_complex");
    const lossConfirmation = toPageExportLossConfirmationError(
      error,
      request.format,
    );
    if (lossConfirmation) throw lossConfirmation;
    throw error;
  }
  if (request.spreadsheetView && !hasSpreadsheetPdfViewReceipt(response.headers.get("content-type"))) throw new SpreadsheetPdfViewError("view_unavailable");
  // Every PDF export or print (download, print preview, archive) says when
  // the file was not printed by the editor.
  if (request.format === "pdf") announcePdfLayout(response, pageId);
  return response;
}

/**
 * Print a page's PDF ahead of the click into the server's kept exports
 * (``prepare=1``: the same request as the export, so the export that follows
 * finds it; no quota, no audit, no bytes back). A print dialog opening, or a
 * pointer on Print, asks for it; whatever it answers, the print itself runs
 * as before. Answers the server's word on it (``X-MRAI-Export-Prepare``:
 * kept, kept_already, running, busy, not_kept, ...), or null.
 */
export async function preparePageExport(request: PageExportRequest): Promise<string | null> {
  try {
    const { path, init } = pageExportRequest(request, "bytes");
    init.headers = { ...(init.headers as Record<string, string> | undefined), Prefer: "respond-async, wait=0" };
    const response = await apiFetchRaw(`${path}&prepare=1`, init);
    return response.headers.get("x-mrai-export-prepare") ?? (response.status === 202 ? "running" : null);
  } catch {
    // A prepare that fails changes nothing: the print renders as before.
    return null;
  }
}

/**
 * The exported file itself, for a caller that shows or prints it rather than
 * saving it (a spreadsheet's File > Print). It is asked for as bytes; a
 * signed-URL ticket (an older API) is refused: its object cannot be read
 * back from the page.
 */
export async function fetchPageExportBlob(
  request: PageExportRequest,
): Promise<Blob> {
  // A PDF read to be shown or printed: a print, to an embedded session.
  const asked = request.format === "pdf" && request.purpose === undefined ? { ...request, purpose: "print" as const } : request;
  return withExportPassword(asked, async (secured) => {
    const response = await requestPageExport(secured, "bytes");
    if (
      response.headers.get("x-mrnote-export-download") === "signed-url" ||
      isJsonResponse(response)
    ) {
      throw new Error("The export arrived as a download link, not a file.");
    }
    return serverStampedInEmbed(await encryptedExportBlob(await response.blob(), secured), secured.format);
  });
}

/** A PDF the export answered an embedded session whose watermark is stamped
 * into exports: the server stamped it (lib/embed/watermark-pdf). */
function serverStampedInEmbed(blob: Blob, format: PageExportFormat): Blob {
  return format === "pdf" && embedExportWatermark() ? markServerStamped(blob) : blob;
}

/**
 * Run an export under the page's password: every file of a page opened with
 * one, or marked Encrypt with Password, leaves encrypted (or not at all),
 * whichever window asked for it. The server's refusal of a plain Word file
 * covers a mark this tab did not see.
 */
async function withExportPassword<T>(
  request: PageExportRequest,
  run: (secured: PageExportRequest) => Promise<T>,
): Promise<T> {
  const { askPageExportPassword, resolvePageExportPassword } = await import("./office/office-encryption");
  const password = request.openPassword ?? (await resolvePageExportPassword(request.pageId.trim(), request.format));
  const secured = password ? { ...request, openPassword: password } : request;
  try {
    return await run(secured);
  } catch (error) {
    if (secured.openPassword || !isApiRequestError(error) || error.code !== "office_encryption_required") throw error;
    return run({ ...request, openPassword: await askPageExportPassword(request.pageId.trim()) });
  }
}

async function encryptedExportBlob(blob: Blob, request: PageExportRequest): Promise<Blob> {
  if (!request.openPassword) return blob;
  const { encryptWordPackage } = await import("./office/office-encryption");
  return encryptWordPackage(blob, request.openPassword, request.signal);
}

/**
 * The exported file's bytes and name, for a caller that packs it (the
 * notebook archive). It is asked for as bytes (`delivery=bytes`): a stored
 * file (an imported original, a PDF page's saved version) came as a signed
 * link the page could not read in production (the private file gate), and
 * every such page was left out of the archive. A ticket from an older API
 * is still read from its link; a link that is not plain http(s), or a store
 * that refuses the read, fails the call rather than packing nothing.
 */
export async function fetchPageExportFile(
  request: PageExportRequest,
): Promise<{ blob: Blob; filename: string }> {
  return withExportPassword(request, async (secured) => {
    const file = await fetchPageExportFileFrom(await requestPageExport(secured, "bytes"), secured);
    return { ...file, blob: await encryptedExportBlob(file.blob, secured) };
  });
}

async function fetchPageExportFileFrom(
  response: Response,
  request: PageExportRequest,
): Promise<{ blob: Blob; filename: string }> {
  const extension = FORMAT_EXTENSION[request.format];
  const fallback = `page-export.${extension}`;
  if (
    response.headers.get("x-mrnote-export-download") === "signed-url" ||
    isJsonResponse(response)
  ) {
    const ticket = signedDownloadTicket(await response.json());
    if (!ticket) throw new Error("The export download ticket is invalid.");
    let url: URL;
    try {
      url = new URL(ticket.download_url, globalThis.location?.href);
    } catch {
      throw new Error("The export download URL is invalid.");
    }
    if (
      (url.protocol !== "https:" && url.protocol !== "http:") ||
      url.username !== "" ||
      url.password !== ""
    ) {
      throw new Error("The export download URL is invalid.");
    }
    // The private file host answers the app origin with credentialed CORS
    // and wants the gate and account cookies: a file link on it is fetched
    // with them (fetchPrivateFile), any other link without.
    const file = await fetchPrivateFile(url.href, {
      referrerPolicy: "no-referrer",
    });
    if (!file.ok) {
      throw new Error(`The export download failed (${file.status}).`);
    }
    return {
      blob: await file.blob(),
      filename: serverDownloadFilename(ticket.filename, fallback, extension),
    };
  }
  return {
    blob: await response.blob(),
    filename: contentDispositionFilename(
      response.headers.get("content-disposition"),
      fallback,
      extension,
    ),
  };
}

/**
 * Export a page and hand the file to the browser. From the click to the
 * file the line at the top of the window shows it (the server writing it,
 * then its bytes), whichever window asked.
 */
export async function downloadPageExport(
  request: PageExportRequest,
): Promise<PageExportDownloadLosses> {
  const operation = beginTransfer({
    direction: "download",
    name: FORMAT_EXTENSION[request.format].toUpperCase(),
    totalBytes: 0,
    umbrella: "secondary",
  });
  const started = Date.now();
  const timer = setInterval(() => operation.preparing((Date.now() - started) / 1000), 1000);
  try {
    return await downloadPageExportFile(request);
  } finally {
    clearInterval(timer);
    operation.end();
  }
}

/**
 * What a downloaded file says it could not carry: the server names it in
 * `X-MRAI-Export-Losses` / `X-MRAI-Export-Loss-Count` on a file it wrote
 * without asking (an imported deck rebuilt from its page, or with slides
 * written anew). Codes only, bounded and checked as the loss gate's are;
 * none when the file carries everything.
 */
/**
 * Not a loss: the server names it on a deck it handed out as a .pptm because
 * the deck holds macros (kept as they are, never run).
 */
export const PPTX_SAVED_AS_PPTM_NOTICE_CODE = "pptx_saved_as_pptm_with_macros";

export interface PageExportDownloadLosses {
  readonly lossCodes: readonly string[];
  readonly lossCount: number;
}

export function pageExportDownloadLosses(headers: Headers): PageExportDownloadLosses {
  const codes: string[] = [];
  for (const raw of (headers.get("x-mrai-export-losses") ?? "").split(",")) {
    const code = raw.trim();
    if (!code || !SAFE_EXPORT_LOSS_CODE.test(code) || codes.includes(code)) continue;
    codes.push(code);
    if (codes.length >= MAX_EXPORT_LOSS_CODES) break;
  }
  const counted = Number(headers.get("x-mrai-export-loss-count"));
  const lossCount =
    Number.isSafeInteger(counted) && counted >= codes.length && counted <= MAX_EXPORT_LOSS_COUNT
      ? counted
      : codes.length;
  return { lossCodes: codes, lossCount: codes.length ? lossCount : 0 };
}

async function downloadEncryptedPageExport(
  request: PageExportRequest & { openPassword: string },
): Promise<PageExportDownloadLosses> {
  const response = await requestPageExport(request, "bytes");
  const losses = pageExportDownloadLosses(response.headers);
  const { blob, filename } = await fetchPageExportFileFrom(response, request);
  const encrypted = await encryptedExportBlob(blob, request);
  if (!startDownloadWhenAllowed(filename, () => downloadBlobFile(filename, encrypted))) {
    throw new Error("The export download could not be started.");
  }
  return losses;
}

async function downloadFillablePdfExport(request: PageExportRequest): Promise<PageExportDownloadLosses> {
  const exportRequest: PageExportRequest = { ...request, fillableForm: false, format: "pdf", layout: "fillable" };
  const response = await requestPageExport(exportRequest, "bytes");
  const losses = pageExportDownloadLosses(response.headers);
  const { blob, filename } = await fetchPageExportFileFrom(response, exportRequest);
  const { fillablePdfFromFieldLinks } = await import("./document-export/fillable-pdf");
  const form = await fillablePdfFromFieldLinks(await blob.arrayBuffer());
  const file = new Blob([form as BlobPart], { type: "application/pdf" });
  if (!startDownloadWhenAllowed(filename, () => downloadBlobFile(filename, file))) {
    throw new Error("The export download could not be started.");
  }
  return losses;
}

async function downloadPageExportFile(given: PageExportRequest): Promise<PageExportDownloadLosses> {
  return withExportPassword(given, (request) =>
    request.openPassword
      ? downloadEncryptedPageExport({ ...request, openPassword: request.openPassword })
      : downloadPlainPageExport(request),
  );
}

async function downloadPlainPageExport(request: PageExportRequest): Promise<PageExportDownloadLosses> {
  if (request.fillableForm && request.format === "pdf") return downloadFillablePdfExport(request);
  // An embedded editor reads the file itself: a link it navigated to would,
  // refused or expired, put the storage provider's error page in place of
  // the editor in the partner's page, and the host may take the file as bytes
  // (lib/embed/frame-downloads).
  const embedded = (globalThis as Record<string, unknown>).__mroEmbedSessionSignedIn === true;
  const response = await requestPageExport(request, embedded ? "bytes" : "link");
  const losses = pageExportDownloadLosses(response.headers);
  const extension = FORMAT_EXTENSION[request.format];
  const fallback = `page-export.${extension}`;

  // `Content-Type` is a CORS-safelisted response header, whereas the optional
  // X-MRNote marker is hidden unless every deployment exposes it explicitly.
  // Treat any successful JSON export response as a ticket so ticket JSON can
  // never be silently downloaded with a .docx/.pdf extension.
  if (
    response.headers.get("x-mrnote-export-download") === "signed-url" ||
    isJsonResponse(response)
  ) {
    const ticket = signedDownloadTicket(await response.json());
    if (!ticket) throw new Error("The export download ticket is invalid.");
    const filename = serverDownloadFilename(ticket.filename, fallback, extension);
    if (
      !startDownloadWhenAllowed(filename, () =>
        downloadUrlFile(filename, ticket.download_url),
      )
    ) {
      throw new Error("The export download URL is invalid.");
    }
    return losses;
  }

  const filename = contentDispositionFilename(
    response.headers.get("content-disposition"),
    fallback,
    extension,
  );
  const blob = serverStampedInEmbed(await readResponseWithProgress(response, filename), request.format);
  // A long export finishes after its click stopped counting as one: the
  // browser may then block the download silently (startDownloadWhenAllowed).
  if (!startDownloadWhenAllowed(filename, () => downloadBlobFile(filename, blob))) {
    throw new Error("The export download could not be started.");
  }
  return losses;
}
