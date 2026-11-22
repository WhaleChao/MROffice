/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture file as the PDF window places it: the steps Insert > Image (and
 * the stamp picker's upload) take from a picked file to the data URL a page
 * can hold, in one place so the picker and a programmatic insert (the live
 * editor's `insertPicture`, an integration's insertImage) check it the same.
 *
 *  1. the file's kind and size as a signature/stamp picture takes them; a
 *     picture past the stamp's 1 MB is fitted to it (decoded and encoded
 *     again smaller), as a sheet's or a slide's inline picture is;
 *  2. the data URL as sanitized for a stamp;
 *  3. its natural size, read by drawing it (a file the browser cannot draw
 *     is refused here);
 *  4. the picture as the PDF's writer takes it (stamp-picture-check.ts).
 *
 * Resolves null when any step refuses the file, or as soon as `current()`
 * says a later pick replaced this one (no further step runs for it).
 */
import {
  SIGNATURE_MAX_IMAGE_BYTES,
  sanitizeSignatureImage,
  validateSignatureImageFile,
} from "@mrnote/core/web/lib/pdf/signatures.ts";
import { fitOfficePictureBlob } from "@mrnote/core/web/lib/office-picture-envelope.ts";
import { pdfStampPictureForWriter } from "./stamp-picture-check";

export interface PdfStampPictureFile {
  /** The picture's data URL, as the page's writer takes it. */
  readonly image: string;
  /** natural width / natural height. */
  readonly aspectRatio: number;
}

function readDataUrl(file: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : null);
    reader.onerror = () => resolve(null);
    reader.readAsDataURL(file);
  });
}

function naturalSize(src: string): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const preview = new Image();
    preview.onload = () => resolve({ width: preview.naturalWidth, height: preview.naturalHeight });
    preview.onerror = () => resolve(null);
    preview.src = src;
  });
}

export async function pdfStampPictureFromFile(
  file: File,
  current: () => boolean = () => true,
): Promise<PdfStampPictureFile | null> {
  const check = validateSignatureImageFile(file);
  let dataUrl: string | null;
  if (check.ok) {
    dataUrl = await readDataUrl(file);
  } else if (check.reason === "size") {
    dataUrl = await fitOfficePictureBlob(
      file,
      file.type.toLowerCase() === "image/png" ? "image/png" : "image/jpeg",
      undefined,
      Math.floor(SIGNATURE_MAX_IMAGE_BYTES / 3) * 4 + 32,
    ).catch(() => null);
  } else {
    return null;
  }
  if (!current()) return null;
  const image = sanitizeSignatureImage(dataUrl);
  if (!image) return null;
  const size = await naturalSize(image);
  if (!current() || !size) return null;
  const aspectRatio = size.width / size.height;
  if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) return null;
  const taken = await pdfStampPictureForWriter({
    dataUrl: image,
    width: size.width,
    height: size.height,
    maxBytes: SIGNATURE_MAX_IMAGE_BYTES,
  }).catch(() => null);
  if (!current()) return null;
  const picture = taken ? sanitizeSignatureImage(taken) : null;
  return picture ? { image: picture, aspectRatio } : null;
}
