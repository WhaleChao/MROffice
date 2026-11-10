/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Change Picture: the images a picture may be changed to.
 *
 * Excel gives a changed picture an image part of its own — of the new
 * image's own type — whenever its old part is shared with another picture
 * (Excel saves copies of one picture over one image) or is of another type.
 * The original-package writer does the same (`picture_media.py`), so neither
 * a shared part nor the old part's extension limits the choice: any bitmap
 * Excel stores as a picture's image is one. A page saved from scratch keeps
 * its wider list (its exporter re-encodes every picture it writes).
 */

export type SpreadsheetPictureWriteRoute = "patchback" | "from-scratch";

/** What the original-package writer stores as a picture's media part
 * (`picture_media.MEDIA_EXTENSIONS`). */
export const PATCHBACK_PICTURE_MEDIA_TYPES = [
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/gif",
  "image/bmp",
] as const;

export const FROM_SCRATCH_PICTURE_MEDIA_TYPES = [
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
  "image/bmp",
] as const;

export function spreadsheetPictureReplacementMediaTypes(
  route: SpreadsheetPictureWriteRoute,
): readonly string[] {
  return route === "patchback"
    ? PATCHBACK_PICTURE_MEDIA_TYPES
    : FROM_SCRATCH_PICTURE_MEDIA_TYPES;
}
