/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF write refused before anything was sent, carrying the reason the
 * reader should see (the save policy's, already translated). A dialog shows
 * its message instead of its generic 「未能保存…」: the write did not fail,
 * it was not allowed, and the reader is told why.
 */
export class PdfWriteRefusedError extends Error {
  readonly refused = true as const;
  constructor(reason: string) {
    super(reason);
    this.name = "PdfWriteRefusedError";
  }
}

export function pdfWriteRefusalReason(error: unknown): string | null {
  return error instanceof PdfWriteRefusedError && error.message ? error.message : null;
}
