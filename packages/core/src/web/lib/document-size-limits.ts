/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Each editor's document total: the one bound that replaced per-item caps.
 *
 * Word, PowerPoint and Excel embed what is inserted (a picture, a clip, a
 * sound, an attached file) whole, with no size limit of its own: the file
 * grows, and only the file's size can stop it. The editors follow that: an
 * item too large for the page is stored beside it (a presentation clip as a
 * DataItem named by `assetId`, a document picture as a page attachment named
 * by `attachment_id`), and what bounds it is the total of its document.
 *
 * The API holds the same numbers (app/services/document_size_limits.py) and
 * refuses past them with `document_total_exceeded`; the editors check here
 * first, before anything is uploaded, and say the same thing.
 */

const MIB = 1024 * 1024;

/**
 * A presentation's local media -- every video, audio clip and transition
 * sound stored beside the deck, together. The PPTX export holds them in
 * memory as it writes the package; one video also stays within PowerPoint
 * for the web's documented 256 MB per video.
 */
export const PRESENTATION_MEDIA_MAX_BYTES = 256 * MIB;














export const PDF_EDITABLE_MAX_BYTES = 128 * MIB;

export type DocumentTotalEditor = "presentation" | "pdf";

export class DocumentTotalExceededError extends Error {
  readonly editor: DocumentTotalEditor;
  readonly limitBytes: number;
  readonly totalBytes: number;

  constructor(editor: DocumentTotalEditor, limitBytes: number, totalBytes: number) {
    super(`document_total_exceeded:${editor}`);
    this.name = "DocumentTotalExceededError";
    this.editor = editor;
    this.limitBytes = limitBytes;
    this.totalBytes = totalBytes;
  }
}

const LIMIT_BYTES: Readonly<Record<DocumentTotalEditor, number>> = {
  presentation: PRESENTATION_MEDIA_MAX_BYTES,
  pdf: PDF_EDITABLE_MAX_BYTES,
};

/** The catalog key (console-notebooks) of each editor's refusal; `{limit}`. */
export const DOCUMENT_TOTAL_MESSAGE_KEYS = {
  presentation: "documentLimits.presentationMediaExceeded",
  pdf: "documentLimits.pdfExceeded",
} as const satisfies Readonly<Record<DocumentTotalEditor, string>>;

/** The editor's total, in bytes. */
export function documentTotalLimitBytes(editor: DocumentTotalEditor): number {
  return LIMIT_BYTES[editor];
}

/**
 * Refuse an insert that would take the document past its total: `current`
 * is what the document already holds, `adding` what the insert brings.
 */
export function requireWithinDocumentTotal(
  editor: DocumentTotalEditor,
  currentBytes: number,
  addingBytes: number,
): void {
  const limit = LIMIT_BYTES[editor];
  const total = Math.max(0, currentBytes) + Math.max(0, addingBytes);
  if (total > limit) throw new DocumentTotalExceededError(editor, limit, total);
}

/** "256 MB": the limit as the messages state it (binary megabytes, as Office). */
export function formatDocumentLimit(bytes: number): string {
  const megabytes = bytes / MIB;
  return `${Number.isInteger(megabytes) ? megabytes : megabytes.toFixed(1)} MB`;
}

/** The API's refusal (`document_total_exceeded`) as the same error, or null. */
export function documentTotalExceededFromApi(
  error: unknown,
  depth = 0,
): DocumentTotalExceededError | null {
  if (error instanceof DocumentTotalExceededError) return error;
  // The API's error (lib/api ApiRequestError), read by its fields; an
  // editor command that wraps it keeps it as its cause.
  const apiError =
    error && typeof error === "object"
      ? (error as { code?: unknown; details?: unknown; cause?: unknown })
      : null;
  if (apiError?.code !== "document_total_exceeded") {
    return apiError?.cause !== undefined && depth < 4
      ? documentTotalExceededFromApi(apiError.cause, depth + 1)
      : null;
  }
  const details =
    apiError.details && typeof apiError.details === "object"
      ? (apiError.details as Record<string, unknown>)
      : {};
  const editor = details.editor;
  if (editor !== "presentation" && editor !== "pdf") return null;
  const limit = Number(details.limit_bytes);
  const total = Number(details.total_bytes);
  return new DocumentTotalExceededError(
    editor,
    Number.isFinite(limit) && limit > 0 ? limit : LIMIT_BYTES[editor],
    Number.isFinite(total) && total > 0 ? total : 0,
  );
}
