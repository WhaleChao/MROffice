/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which files can become an editable document, and how large they may be.
 *
 * Only the file's NAME, media type and size are read here. Every Word, Excel
 * and PowerPoint file is read on the server
 * (`POST /api/v1/office/import-projection`, or the import batch's
 * `create_source_page`); the browser keeps no OOXML reader. A PDF's text is
 * still read in the browser by pdf.js (`pdfTextImport.ts`).
 */

import { WORD_SOURCE_MAX_BYTES } from "./word-source-limit";

export type EditableDocumentExtension = ".docx" | ".xlsx" | ".pptx" | ".pdf";
export type EditableDocumentFileDescriptor = Pick<File, "name"> &
  Partial<Pick<File, "size" | "type">>;

/**
 * The server's own cap for a workbook or deck that becomes an EDITABLE page
 * (`editable_page._MAX_OFFICE_BYTES`). A Word file's is
 * {@link WORD_SOURCE_MAX_BYTES}, the server's `MAX_DOCX_SOURCE_BYTES`.
 */
export const OFFICE_IMPORT_MAX_BYTES = 32 * 1024 * 1024;



export const PDF_IMPORT_MAX_BYTES = 64 * 1024 * 1024;

const EDITABLE_DOCUMENT_MEDIA_TYPES: ReadonlyMap<string, EditableDocumentExtension> =
  new Map([
    ["application/pdf", ".pdf"],
    [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ".docx",
    ],
    ["application/vnd.ms-word.document.macroenabled.12", ".docx"],
    [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ".xlsx",
    ],
    [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      ".pptx",
    ],
    ["application/vnd.ms-powerpoint.presentation.macroenabled.12", ".pptx"],
    ["application/vnd.ms-powerpoint.template.macroenabled.12", ".pptx"],
    ["application/vnd.ms-powerpoint.slideshow.macroenabled.12", ".pptx"],
    ["application/vnd.openxmlformats-officedocument.presentationml.template", ".pptx"],
    ["application/vnd.openxmlformats-officedocument.presentationml.slideshow", ".pptx"],
  ] as const);

function ext(name: string): string {
  const index = name.toLowerCase().lastIndexOf(".");
  return index >= 0 ? name.toLowerCase().slice(index) : "";
}

export function editableDocumentExtension(
  file: EditableDocumentFileDescriptor,
): EditableDocumentExtension | null {
  const extension = ext(file.name);
  if (extension === ".docm") return ".docx";
  // A macro-enabled presentation, a Design Template or a Show is a deck.
  if ([".pptm", ".potm", ".ppsm", ".potx", ".ppsx"].includes(extension)) return ".pptx";
  if ([".docx", ".xlsx", ".pptx", ".pdf"].includes(extension)) {
    return extension as EditableDocumentExtension;
  }
  const mediaType = (file.type || "").split(";", 1)[0].toLowerCase();
  return EDITABLE_DOCUMENT_MEDIA_TYPES.get(mediaType) || null;
}

export function canCompileFileAsEditableDocument(
  file: EditableDocumentFileDescriptor,
): boolean {
  return Boolean(editableDocumentExtension(file));
}

export function isComplexImportTooLarge(
  file: EditableDocumentFileDescriptor,
): boolean {
  const size = file.size || 0;
  const extension = editableDocumentExtension(file);
  const maxBytes =
    extension === ".pdf"
      ? PDF_IMPORT_MAX_BYTES
      : extension === ".docx"
        ? WORD_SOURCE_MAX_BYTES
        : OFFICE_IMPORT_MAX_BYTES;
  return size > maxBytes;
}

/** One entry of an import's loss record: what it could not carry, how many. */
export interface ImportLossEntry {
  code: string;
  limit: number;
  omitted?: number;
  /**
   * A gap of a PREVIEW that the persisted import does not share; reported
   * without rejecting the file.
   */
  degraded?: boolean;
}

export interface ImportLossManifest {
  status: "lossless" | "lossy";
  entries: ImportLossEntry[];
}

/** What an editable-document import hands its caller. */
export interface CompiledDocumentImport {
  nodes: Array<Record<string, unknown>>;
  lossManifest: ImportLossManifest;
  title?: string;
  /** The Word file's document-level attrs, in the page's own spelling. */
  docAttrs?: Record<string, unknown>;
  /** The loss record in the page's receipt spelling (`importLosses`). */
  importLosses?: { code: string; omitted: number }[];
}

export class LossyImportRejectedError extends Error {
  readonly code = "complex_import_lossy";

  constructor(readonly lossManifest: ImportLossManifest) {
    super("complex_import_lossy");
    this.name = "LossyImportRejectedError";
  }
}







export class PdfPasswordRequiredImportError extends Error {
  readonly code = "pdf_password_required";

  constructor() {
    super("pdf_password_required");
    this.name = "PdfPasswordRequiredImportError";
  }
}
