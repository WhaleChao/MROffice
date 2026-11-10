/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The `content_json.type` of each native page. Kept apart from
 * `native-page-schemas` (which re-exports them) so a reader that only needs
 * to tell a page's kind does not load every page schema with it.
 */
export const DOCUMENT_CONTENT_TYPE = "doc";
export const SPREADSHEET_CONTENT_TYPE = "mrnote.spreadsheet.v1";
export const PRESENTATION_CONTENT_TYPE = "mrnote.presentation.v1";
export const PDF_CONTENT_TYPE = "mrnote.pdf.v1";
export const CODE_CONTENT_TYPE = "mrnote.code.v1";
