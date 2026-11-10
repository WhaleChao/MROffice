/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfActionButtonFieldKind =
  | "resetButton"
  | "printButton"
  | "navigationButton"
  | "viewButton"
  | "saveButton"
  | "goToPageButton"
  | "webLinkButton";

export function isPdfActionButtonFieldKind(
  value: unknown,
): value is PdfActionButtonFieldKind {
  return (
    value === "resetButton" ||
    value === "printButton" ||
    value === "navigationButton" ||
    value === "viewButton" ||
    value === "saveButton" ||
    value === "goToPageButton" ||
    value === "webLinkButton"
  );
}

export const PDF_CREATE_FORM_FIELDS_MAX_CANDIDATES = 200;
