/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Refusals of a whole import selection, kept out of `page-file-import` so the
 * sidebar can recognize them without loading the importer.
 */
import { isApiRequestError } from "./api";

/**
 * Whether the import was stopped by the server's moment, not the file: the
 * API's 503 ``source_reading_interrupted`` with ``details.retryable`` (memory
 * or a wall clock run out under load, a reader process gone). The file is
 * not at fault and reads whole a moment later (corpus R492-D12): the import
 * is tried again once by itself, then the person is told the server ran out
 * of room or time and offered Retry -- never that the file is broken.
 */
export function isRetryableReadingInterruption(error: unknown): boolean {
  if (!isApiRequestError(error)) return false;
  if (error.status !== 503 || error.code !== "source_reading_interrupted") return false;
  return error.details?.retryable === true;
}

/** A refusal of a whole selection, with what decided it. */
export class PageImportSelectionError extends Error {
  readonly code: "no_files" | "too_many_files" | "batch_too_large" | "invalid_size";
  readonly details: Record<string, number | string>;

  constructor(
    code: PageImportSelectionError["code"],
    message: string,
    details: Record<string, number | string> = {},
  ) {
    super(message);
    this.name = "PageImportSelectionError";
    this.code = code;
    this.details = details;
  }
}
