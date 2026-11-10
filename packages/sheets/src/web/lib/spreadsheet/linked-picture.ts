/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pictureObjectDisplaySrc } from "@mrnote/core/web/lib/page-picture-objects.ts";
import { MAX_CALCULATION_WEB_IMAGE_SOURCE } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";
import { requestSpreadsheetWebImage, spreadsheetWebImagePicture } from "./web-image-store";

/** The https address a picture links, when the sheet may fetch it. */
export function linkedPictureFetchSource(object: { readonly linkedPicture?: unknown }): string | null {
  const link = object.linkedPicture;
  if (typeof link !== "string" || link.length > MAX_CALCULATION_WEB_IMAGE_SOURCE) return null;
  try {
    return new URL(link).protocol === "https:" ? link : null;
  } catch {
    return null;
  }
}

/**
 * What a sheet picture draws: a linked picture's fetched picture once this
 * tab has it (asking for it the first time), else the picture the page holds
 * (Excel's frame, for a linked one), a stored picture by its route.
 */
export function sheetPictureDisplaySrc(object: { readonly src?: string; readonly linkedPicture?: unknown }): string {
  const link = linkedPictureFetchSource(object);
  if (link) {
    const fetched = spreadsheetWebImagePicture(link);
    if (fetched) return fetched.src;
    requestSpreadsheetWebImage(link);
  }
  return pictureObjectDisplaySrc(object.src ?? "");
}
