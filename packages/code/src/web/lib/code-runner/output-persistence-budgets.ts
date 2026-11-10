/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Shared persistence limits for standalone Code execution output rendering. */

export const MAX_PERSISTED_CELL_OUTPUT_JSON_CHARS = 2_000_000;
export const MAX_PERSISTED_ATTACHMENT_CHARS = 1_000_000;

// SVG is active XML/markup and is deliberately excluded from data-URI image
// rendering. Keep this tuple aligned with output-persistence.ts.
export const ALLOWED_PERSISTED_RASTER_IMAGE_MIMES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
  "image/bmp",
] as const;
