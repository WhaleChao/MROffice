/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The largest Word file read whole, shared by the import size gate, the Word
 * zip loader and the sparse-table worker. It mirrors the server's
 * `MAX_DOCX_SOURCE_BYTES`: Word stores each chart's embedded workbook
 * uncompressed, so a document of many native charts is large on disk while
 * inflating little. Kept free of imports so the worker bundle does not pull in
 * the zip loader.
 *
 * A larger Word file is uploaded in parts (`needsSourceImport`) and the server
 * reads it without its pictures: they stay in the stored file, named by digest
 * (`apps/api/.../docx/picture_references.py`), so a file of up to
 * {@link WORD_EDITABLE_MAX_BYTES} still opens as an editable document; what
 * bounds it is its text, tables and styles, not its pictures.
 */
export const WORD_SOURCE_MAX_BYTES = 50 * 1024 * 1024;

/** The largest Word file made into an editable page (the server's
 * `EDITABLE_SOURCE_MAX_BYTES["docx"]`); a larger one opens as a preview. */
export const WORD_EDITABLE_MAX_BYTES = 2 * 1024 * 1024 * 1024;
