/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Stable code emitted by the loss-manifest gate of an import projection. */
export const COMPLEX_IMPORT_LOSSY_CODE = "complex_import_lossy" as const;

/** The server refused a PDF it cannot open without its password. */
export function isPdfPasswordRequiredImportError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    (error as { code?: unknown }).code === "pdf_password_required"
  );
}

interface ErrorWithCode {
  readonly code?: unknown;
  readonly cause?: unknown;
}

export function isLossyImportRejection(error: unknown): boolean {
  let current = error;
  const seen = new WeakSet<object>();
  while (typeof current === "object" && current !== null) {
    if (seen.has(current)) return false;
    seen.add(current);
    const candidate = current as ErrorWithCode;
    if (candidate.code === COMPLEX_IMPORT_LOSSY_CODE) return true;
    current = candidate.cause;
  }
  return false;
}

/**
 * User-facing wrapper that preserves the compiler's machine-readable code and
 * original error as `cause`. Callers must never convert this into an implicit
 * lossy retry.
 */
export class LosslessImportRequiredError extends Error {
  readonly code = COMPLEX_IMPORT_LOSSY_CODE;

  constructor(
    message: string,
    readonly cause: unknown,
  ) {
    super(message);
    this.name = "LosslessImportRequiredError";
  }
}

export function mapLossyImportRejection(
  error: unknown,
  message: string,
): unknown {
  if (!isLossyImportRejection(error)) return error;
  if (error instanceof LosslessImportRequiredError) return error;
  return new LosslessImportRequiredError(message, error);
}
