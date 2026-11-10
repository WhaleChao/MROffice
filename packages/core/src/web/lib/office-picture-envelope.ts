/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fit an inline picture into the page's single-string bound.
 *
 * Office embeds a picture whole; a page cannot — the API's page envelope
 * refuses any single JSON string longer than OFFICE_PICTURE_SRC_MAX_LENGTH
 * (MAX_JSON_STRING_CHARS), not as a picture it trims but as a malformed
 * page, so a data: URL past it makes every save of the workbook or deck
 * fail. A PNG, JPEG or WebP that overruns it is re-encoded through the same
 * decoder/encoder the document editor's Compress Pictures uses, at a lower
 * quality first (a photo usually fits at its own size) and then at smaller
 * sizes, until its URL fits; a picture that cannot be made to fit, or one of
 * a kind the encoder does not write (GIF, BMP, SVG), is refused where it is
 * inserted instead of at every save. Shared by the spreadsheet and the
 * presentation editors.
 */

import {
  blobDataUrl,
  browserPictureCompressionAdapter,
  pictureBlobFromDataUrl,
  type DocumentPictureCompressionAdapter,
} from "../components/console/editor/native-document/documentPictureCompression";

/** The page envelope's longest single string (the API's MAX_JSON_STRING_CHARS). */
export const OFFICE_PICTURE_SRC_MAX_LENGTH = 1_048_576;

/**
 * The largest file whose base64 data: URL fits OFFICE_PICTURE_SRC_MAX_LENGTH;
 * 32 characters cover "data:image/jpeg;base64," and every longer prefix. The
 * ceiling for a file the fit step cannot re-encode.
 */
export const OFFICE_PICTURE_FILE_MAX_BYTES = Math.floor((OFFICE_PICTURE_SRC_MAX_LENGTH - 32) / 4) * 3;

/** The kinds the fit step can re-encode (what the compression helper writes). */
export const FITTABLE_PICTURE_MIME_TYPES = Object.freeze([
  "image/png",
  "image/jpeg",
  "image/webp",
] as const);

/**
 * The largest PNG, JPEG or WebP file an insert accepts: it is decoded straight
 * from the file ({@link FittedPictureReader}) and re-encoded down to the
 * page's bound, never held whole as a data: URL. The deck's own file limit.
 */
export const MAX_FITTABLE_PICTURE_FILE_BYTES = 50 * 1024 * 1024;

/** Whether a local file of `mimeType` can be re-encoded down to the envelope. */
export function isFittablePictureMimeType(mimeType: string): boolean {
  const normalized = mimeType.trim().toLowerCase();
  return (FITTABLE_PICTURE_MIME_TYPES as readonly string[]).includes(
    normalized === "image/jpg" ? "image/jpeg" : normalized,
  );
}

/** The largest local file of `mimeType` an insert accepts before fitting. */
export function officePictureFileByteCeiling(mimeType: string): number {
  return isFittablePictureMimeType(mimeType)
    ? MAX_FITTABLE_PICTURE_FILE_BYTES
    : OFFICE_PICTURE_FILE_MAX_BYTES;
}

// Edge scales tried in turn; each step drops the pixel count by roughly a
// third, so a 40-megapixel photo reaches a 786 KB JPEG within the list.
const FIT_SCALES = [1, 0.85, 0.7, 0.58, 0.48, 0.4, 0.33, 0.27, 0.22, 0.18, 0.15, 0.125, 0.1, 0.08, 0.065, 0.05];

/**
 * The picture's data: URL at a length the page can hold: `source` itself when
 * it already fits, a re-encoded one otherwise, or null when nothing fits or
 * `source` is not a decodable data: URL.
 */
/**
 * The page's cache of a picture file, fitted to ``maxLength`` without first
 * holding the whole file as a data URL: a photo of tens of megabytes is
 * decoded once and re-encoded at the first scale that fits.
 */
export async function fitOfficePictureBlob(
  blob: Blob,
  mimeType: "image/png" | "image/jpeg" | "image/webp",
  adapter?: DocumentPictureCompressionAdapter,
  maxLength = OFFICE_PICTURE_SRC_MAX_LENGTH,
): Promise<string | null> {
  const maxBytes = Math.floor((maxLength - 32) / 4) * 3;
  if (blob.size <= maxBytes) {
    const direct = await blobDataUrl(blob);
    if (direct.length <= maxLength) return direct;
  }
  const activeAdapter = adapter ?? (await browserPictureCompressionAdapter());
  const raster = await activeAdapter.decode(blob);
  try {
    if (!(raster.width > 0) || !(raster.height > 0)) return null;
    const qualities = mimeType === "image/png" ? [1] : [0.82, 0.72];
    for (const scale of FIT_SCALES) {
      const width = Math.max(1, Math.round(raster.width * scale));
      const height = Math.max(1, Math.round(raster.height * scale));
      for (const quality of qualities) {
        const encoded = await raster.encode(width, height, mimeType, quality);
        if (!encoded || encoded.type !== mimeType || encoded.size <= 0 || encoded.size > maxBytes) continue;
        const dataUrl = await blobDataUrl(encoded);
        if (dataUrl.length <= maxLength) return dataUrl;
      }
    }
    return null;
  } finally {
    raster.dispose();
  }
}

export async function fitOfficePictureSource(
  source: unknown,
  adapter?: DocumentPictureCompressionAdapter,
  maxLength = OFFICE_PICTURE_SRC_MAX_LENGTH,
): Promise<string | null> {
  if (typeof source !== "string" || !source) return null;
  if (source.length <= maxLength) return source;
  const decoded = await pictureBlobFromDataUrl(source);
  if (!decoded) return null;
  const activeAdapter = adapter ?? (await browserPictureCompressionAdapter());
  const raster = await activeAdapter.decode(decoded.blob);
  try {
    if (!(raster.width > 0) || !(raster.height > 0)) return null;
    // PNG is lossless: only a smaller size makes it smaller. The lossy kinds
    // first give up quality at their own size, then size.
    const qualities = decoded.mimeType === "image/png" ? [1] : [0.82, 0.72];
    // The base64 URL is 4/3 of the bytes plus its prefix.
    const maxBytes = Math.floor((maxLength - 32) / 4) * 3;
    for (const scale of FIT_SCALES) {
      const width = Math.max(1, Math.round(raster.width * scale));
      const height = Math.max(1, Math.round(raster.height * scale));
      for (const quality of qualities) {
        const blob = await raster.encode(width, height, decoded.mimeType, quality);
        if (!blob || blob.type !== decoded.mimeType || blob.size <= 0 || blob.size > maxBytes) continue;
        const dataUrl = await blobDataUrl(blob);
        if (dataUrl.length <= maxLength) return dataUrl;
      }
    }
    return null;
  } finally {
    raster.dispose();
  }
}

/**
 * The page's inline copy of a picture file: a PNG, JPEG or WebP decoded from
 * the file and fitted ({@link fitOfficePictureBlob}); any other kind as its
 * own data: URL when that fits. Null when it cannot be made to fit.
 */
export async function readFittedPicture(file: Blob): Promise<string | null> {
  const type = file.type.trim().toLowerCase();
  const normalized = type === "image/jpg" ? "image/jpeg" : type;
  if (normalized === "image/png" || normalized === "image/jpeg" || normalized === "image/webp") {
    return fitOfficePictureBlob(file, normalized);
  }
  const direct = await blobDataUrl(file);
  return direct.length <= OFFICE_PICTURE_SRC_MAX_LENGTH ? direct : null;
}

/**
 * A FileReader for a picture a page holds inline: ``readAsDataURL`` answers
 * the picture already fitted to the page's bound ({@link readFittedPicture}),
 * so a photo of tens of megabytes is decoded once instead of first becoming
 * a data: URL several times its size. ``result`` is null when it cannot fit.
 */
export class FittedPictureReader {
  result: string | null = null;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;

  readAsDataURL(file: Blob): void {
    void readFittedPicture(file).then(
      (fitted) => {
        this.result = fitted;
        this.onload?.();
      },
      () => this.onerror?.(),
    );
  }
}
