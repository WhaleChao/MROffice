/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The one sentence a text box the editor cannot edit at all shows: it is
 * drawn in a Type3 font (WPS refuses these too; docs/pdf/edit-gates.md A2).
 * The English sentence is also the refusal's message across the worker. */
export const PDF_FALLBACK_TYPE3_MESSAGE =
  "This text is drawn in a Type3 font, whose glyphs are pictures rather than characters; it cannot be edited as text.";

export function pdfFallbackType3Copy(t: (key: string) => string): string {
  return t("pdf.pageWindow.type3FontNotEditable");
}



export const PDF_FALLBACK_NOT_FOUND_MESSAGE = "The page content draws no glyph where this text stands.";

/** A line that paints nothing because a text box of the page draws its
 * words: the line is that box's hidden source, not text of its own. Said
 * when the box could not be read as a box (its record is damaged or was
 * written by something else), so the line is not offered as plain text. */
export const PDF_FALLBACK_TEXT_BOX_SOURCE_MESSAGE =
  "This text is drawn by a text box of the page whose record cannot be read; the line under it is not edited on its own.";
