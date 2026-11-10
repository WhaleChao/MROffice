/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 图片字段 (Acrobat's Image Field, ONLYOFFICE's picture form): the picture a
 * person filling the form places on the field. It is prepared as any Office
 * picture is (presentation-image-file: PNG and JPEG kept, HEIC, WebP, GIF,
 * BMP, TIFF and AVIF converted, the whole fitted to the picture envelope) and
 * sent to the native form writer, which makes it the field's icon and face.
 */
import {
  PresentationImageFileError,
  preparePresentationImageFile,
  type PresentationImageFileFailureCode,
} from "@mrnote/core/web/lib/presentation-image-file.ts";

export interface PdfImageFieldPicture {
  readonly mime: "image/png" | "image/jpeg";
  /** The picture's bytes, base64. */
  readonly data: string;
}

export type PdfImageFieldPictureFailure = PresentationImageFileFailureCode;

/** The file types the picker offers (what the preparation decodes). */
export const PDF_IMAGE_FIELD_ACCEPT =
  "image/png,image/jpeg,image/webp,image/gif,image/bmp,image/tiff,image/heic,image/heif,image/avif";

/** The picture to place, or the reason it cannot be (a vector picture has
 * no place on a button's icon: Acrobat takes raster pictures only). */
export async function preparePdfImageFieldPicture(
  file: File,
): Promise<PdfImageFieldPicture | { readonly failure: PdfImageFieldPictureFailure }> {
  try {
    const prepared = await preparePresentationImageFile(file);
    if (prepared.mimeType !== "image/png" && prepared.mimeType !== "image/jpeg") {
      return { failure: "unsupported-type" };
    }
    const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/]+={0,2})$/.exec(prepared.src);
    if (!match || match[1] !== prepared.mimeType) return { failure: "cache-failed" };
    return { mime: prepared.mimeType, data: match[2] };
  } catch (error) {
    return { failure: error instanceof PresentationImageFileError ? error.code : "decode-failed" };
  }
}
