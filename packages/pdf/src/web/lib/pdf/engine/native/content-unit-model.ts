/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-safe part of the page-content layer move
 * (vector-content-transaction.ts: the Selection pane's objects): the rows'
 * letters and the refusals the window says in words.
 */

export type PdfNativeContentUnitKind =
  | "path"
  | "text"
  | "image"
  | "inline-image"
  | "shading"
  | "form"
  | "group"
  | "other";

/** The kinds a content-unit layer move writes (paths and pictures have their
 * own planners, which hold them to their proofs). */
export type PdfNativeMovableContentUnitKind = "text" | "shading" | "form" | "group";

export interface PdfNativeContentUnitIdentity {
  readonly kind: PdfNativeMovableContentUnitKind;
  /** Its place among the page's rows of its kind, back first. */
  readonly order: number;
  /** SHA-256 of the row's bytes as the page draws them. */
  readonly digest: string;
}

/** The letter each row kind is written as in a page's `kinds`. */
export const PDF_NATIVE_CONTENT_UNIT_KIND_LETTER: Readonly<Record<PdfNativeContentUnitKind, string>> =
  Object.freeze({
    path: "p",
    text: "t",
    image: "i",
    "inline-image": "n",
    shading: "s",
    form: "f",
    group: "g",
    other: "o",
  });

/** A page's rows as one letter each. */
export function pdfNativeContentUnitKinds(rows: readonly Readonly<{ kind: PdfNativeContentUnitKind }>[]): string {
  return rows.map((row) => PDF_NATIVE_CONTENT_UNIT_KIND_LETTER[row.kind]).join("");
}

export const MRAI_NATIVE_CONTENT_UNIT_ALREADY_FRONT = "PDF page object is already in front of the page content.";
export const MRAI_NATIVE_CONTENT_UNIT_ALREADY_BACK = "PDF page object is already behind the page content.";
/** The page's content reads as other objects than the reader showed: the
 * move is refused rather than made on another object. */
export const MRAI_NATIVE_CONTENT_UNIT_ROWS_DIFFER = "PDF page content reads differently from the page shown.";
