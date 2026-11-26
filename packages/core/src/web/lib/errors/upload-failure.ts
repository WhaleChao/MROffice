/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One sentence for why a file did not upload or import, and what to do next.
 *
 * Every upload path (page import, a picture or a video inserted into an
 * editor, an attachment) reaches the same refusals: the storage is full, the
 * file is past a size limit, its type is not accepted, the connection broke,
 * the server could not store it. The person needs the file's name, the sizes
 * that decided it, and the step that helps; a status code or an internal
 * failure name helps nobody. This module reads the stable backend code and
 * details and returns that sentence from the `error` catalog, in the page's
 * language, without a translation hook.
 */

import { isApiRequestError } from "../api";
import { currentErrorCatalog } from "./error-catalog";
import { desktopLibreOffice } from "../desktop-libreoffice";

function currentLocale(): string {
  if (typeof document !== "undefined" && document.documentElement.lang) {
    return document.documentElement.lang;
  }
  return "en";
}

/** A size as people read it: "350 MB", "1.2 GB", "820 KB". */
export function formatFileSize(bytes: number, locale = currentLocale()): string {
  const value = Number(bytes);
  if (!Number.isFinite(value) || value < 0) return "";
  const units: [number, string][] = [
    [1024 ** 4, "terabyte"],
    [1024 ** 3, "gigabyte"],
    [1024 ** 2, "megabyte"],
    [1024, "kilobyte"],
  ];
  for (const [size, unit] of units) {
    if (value >= size) {
      const amount = value / size;
      try {
        return new Intl.NumberFormat(locale, {
          style: "unit",
          unit,
          unitDisplay: "short",
          maximumFractionDigits: amount < 10 ? 1 : 0,
        }).format(amount);
      } catch {
        return `${amount.toFixed(amount < 10 ? 1 : 0)} ${unit[0].toUpperCase()}B`;
      }
    }
  }
  try {
    return new Intl.NumberFormat(locale, { style: "unit", unit: "byte", unitDisplay: "short" }).format(value);
  } catch {
    return `${value} B`;
  }
}

function fill(template: string, values: Record<string, string | number>): string {
  return template.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in values ? String(values[key]) : match,
  );
}

function catalogText(key: string, values: Record<string, string | number> = {}): string | null {
  const template = currentErrorCatalog()[key];
  return typeof template === "string" && template ? fill(template, values) : null;
}

function numberDetail(details: Record<string, unknown> | undefined, key: string): number | null {
  const value = details?.[key];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function extensionOf(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot > 0 ? name.slice(dot).toLowerCase() : "";
}

export interface UploadFailureFile {
  readonly name: string;
  readonly size: number;
}

interface ErrorShape {
  readonly status: number;
  readonly code: string;
  readonly details: Record<string, unknown>;
}

function errorShape(error: unknown): ErrorShape | null {
  let current: unknown = error;
  for (let depth = 0; depth < 5 && current; depth += 1) {
    if (isApiRequestError(current)) {
      return {
        status: current.status,
        code: current.code || "",
        details: current.details ?? {},
      };
    }
    const candidate = current as { code?: unknown; status?: unknown; details?: unknown; cause?: unknown };
    if (typeof candidate.code === "string" && typeof candidate.status === "number") {
      return {
        status: candidate.status,
        code: candidate.code,
        details:
          candidate.details && typeof candidate.details === "object"
            ? (candidate.details as Record<string, unknown>)
            : {},
      };
    }
    current = candidate.cause;
  }
  return null;
}

function isConnectionFailure(error: unknown): boolean {
  if (error instanceof TypeError) return true; // fetch: the request never got an answer
  const shape = errorShape(error);
  return shape !== null && (shape.status === 0 || shape.code === "network");
}

/**
 * Why ``file`` did not upload or import, as one sentence with what to do.
 * Returns ``null`` for an error this module cannot place, so the caller keeps
 * its own wording for it.
 */
export function uploadFailureReason(error: unknown, file: UploadFailureFile): string | null {
  const size = formatFileSize(file.size);
  if (isConnectionFailure(error)) {
    return catalogText("upload.reason.network");
  }
  const shape = errorShape(error);
  if (!shape) return null;
  const { status, code, details } = shape;
  if (code === "storage_quota_exceeded" || details.key === "storage.bytes.max") {
    const quota = numberDetail(details, "quota_bytes");
    const used = numberDetail(details, "used_bytes");
    const incoming = numberDetail(details, "incoming_bytes") ?? file.size;
    if (quota !== null && used !== null) {
      return catalogText("upload.reason.storageFull", {
        size: formatFileSize(incoming),
        free: formatFileSize(Math.max(0, quota - used)),
        quota: formatFileSize(quota),
      });
    }
    return catalogText("upload.reason.storageFullPlain");
  }
  if (code === "plan_limit_reached" && details.key === "pages.max") {
    const limit = numberDetail(details, "limit");
    return limit !== null
      ? catalogText("upload.reason.pagesLimit", { limit })
      : catalogText("upload.reason.pagesLimitPlain");
  }
  if (code === "plan_limit_reached" && details.key === "resource.uploads.monthly") {
    const current = numberDetail(details, "current");
    const limit = numberDetail(details, "limit");
    return current !== null && limit !== null
      ? catalogText("upload.monthlyLimitReachedWithUsage", { current, limit })
      : catalogText("upload.monthlyLimitReached");
  }
  if (code === "concurrency_limited") {
    const limit = numberDetail(details, "limit");
    const running = numberDetail(details, "running");
    const key =
      details.kind === "imports" ? "upload.reason.importsAtOnce"
      : details.kind === "files" ? "upload.reason.filesAtOnce"
      : null;
    return key && limit !== null && running !== null
      ? catalogText(key, { limit, running })
      : catalogText("code.concurrency_limited");
  }
  if (code === "file_too_large" || code === "payload_too_large" || status === 413) {
    const limit =
      numberDetail(details, "max_bytes") ?? numberDetail(details, "hard_reject_bytes");
    return limit !== null
      ? catalogText("upload.reason.fileTooLarge", { size, limit: formatFileSize(limit) })
      : catalogText("upload.reason.fileTooLargePlain", { size });
  }
  if (code === "file_infected") {
    const signature = typeof details.signature === "string" ? details.signature : "";
    return catalogText("upload.reason.malwareFound", { signature });
  }
  if (code === "unsupported_media_type" || code === "invalid_import_source_kind" || status === 415) {
    const ext = extensionOf(file.name);
    return ext
      ? catalogText("upload.reason.unsupportedType", { ext })
      : catalogText("upload.reason.unsupportedTypePlain");
  }
  if (code === "libreoffice_required") {
    // MRNote Classic: .doc/.xls/.ppt/.odt... convert with an installed LibreOffice.
    const format = typeof details.format === "string" ? `.${details.format}` : extensionOf(file.name);
    // The app can fetch it (the prompt offers "Download and enable").
    const key = desktopLibreOffice() ? "upload.reason.needsLibreOfficeDownload" : "upload.reason.needsLibreOffice";
    return catalogText(key, { ext: format || file.name }) ?? catalogText("upload.reason.needsLibreOffice", { ext: format || file.name });
  }
  if (code.startsWith("office_")) {
    // A file encrypted with a password (lib/office/office-encryption).
    const reason = typeof details.reason === "string" ? details.reason : "";
    if (code === "office_encryption_unsupported" && reason) {
      const scheme = catalogText(`upload.reason.encryption.${reason}`) ?? reason;
      return catalogText("upload.reason.encryptionUnsupported", { scheme, reason });
    }
    const key =
      code === "office_encryption_cancelled" ? "upload.reason.encryptedCancelled"
      : code === "office_password_incorrect" ? "upload.reason.encryptedIncorrect"
      : code === "office_integrity_failed" ? "upload.reason.encryptedIntegrity"
      : code === "office_file_damaged" ? "upload.reason.encryptedDamaged"
      : null;
    if (key) return catalogText(key);
  }
  if (code === "upload_mismatch" || code === "content_type_mismatch") {
    return catalogText("upload.reason.mismatch");
  }
  if (code === "upload_incomplete") {
    const missing = numberDetail(details, "missing_count");
    return missing !== null
      ? catalogText("upload.reason.incomplete", { count: missing })
      : catalogText("upload.reason.incompletePlain");
  }
  if (code === "upload_expired" || code === "upload_not_found") {
    return catalogText("upload.reason.expired");
  }
  if (code === "empty_body") {
    return catalogText("upload.reason.empty");
  }
  if (code === "zip_empty") {
    return catalogText("upload.reason.zipEmpty");
  }
  if (code === "invalid_filename") {
    return catalogText("upload.reason.badName");
  }
  if (status === 403 || code === "forbidden" || code === "page_read_only") {
    return catalogText("upload.reason.noPermission");
  }
  if (status === 404 || code === "not_found") {
    return catalogText("upload.reason.notebookGone");
  }
  if (status === 429 || code === "rate_limited") {
    return catalogText("upload.reason.tooFast");
  }
  if (code === "source_reading_interrupted") {
    // The server ran out of room or time reading the file (503, retryable;
    // the import already tried again once): never that the file is broken.
    return catalogText("code.source_reading_interrupted") ?? catalogText("upload.reason.serverBusy");
  }
  if (
    status >= 500 ||
    code === "storage_error" ||
    code === "upload_storage_unavailable" ||
    code === "upload_lifecycle_timeout"
  ) {
    return catalogText("upload.reason.serverBusy");
  }
  return null;
}

/** {@link uploadFailureReason}, or the catalog's generic upload sentence. */
export function uploadFailureText(error: unknown, file: UploadFailureFile): string {
  return (
    uploadFailureReason(error, file) ??
    catalogText("upload.reason.generic") ??
    "The upload did not finish. Try again."
  );
}
