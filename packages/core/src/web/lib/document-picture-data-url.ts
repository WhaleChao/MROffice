/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A document picture's `data:` URL, decoded and checked against its bytes.
 *
 * Pure: no package format is read or written here. Used where the editor must
 * know what a picture really is before re-encoding it (Draw on Picture keeps
 * the picture's own browser-encodable format).
 */

export interface DocumentPictureFormat {
  ext: "png" | "jpeg" | "gif" | "bmp" | "webp";
  contentType: string;
}

export interface DecodedDocumentPicture {
  format: DocumentPictureFormat;
  bytes: Uint8Array;
  /** The normalized (whitespace-free) base64 payload. */
  base64: string;
}

function pictureFormatForMime(mime: string): DocumentPictureFormat | null {
  switch (mime.toLowerCase()) {
    case "image/png":
      return { ext: "png", contentType: "image/png" };
    case "image/jpeg":
    case "image/jpg":
      return { ext: "jpeg", contentType: "image/jpeg" };
    case "image/gif":
      return { ext: "gif", contentType: "image/gif" };
    case "image/bmp":
      return { ext: "bmp", contentType: "image/bmp" };
    case "image/webp":
      return { ext: "webp", contentType: "image/webp" };
    default:
      return null;
  }
}

function base64ToBytes(b64: string): Uint8Array | null {
  const decode = (globalThis as { atob?: (s: string) => string }).atob;
  if (typeof decode !== "function") return null;
  let binary: string;
  try {
    binary = decode(b64);
  } catch {
    return null;
  }
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/**
 * Parse an image `src` into its format and bytes, or null when it is not a
 * base64 data URL of a supported raster format whose bytes carry that
 * format's signature. Remote and non-base64 sources return null.
 */
export function decodeDocumentPictureDataUrl(src: unknown): DecodedDocumentPicture | null {
  if (typeof src !== "string") return null;
  const match = src
    .trim()
    .match(/^data:([a-z]+\/[a-z0-9.+-]+)((?:;[a-z0-9-]+=[^;,]+)*);base64,([\s\S]*)$/i);
  if (!match) return null;
  const format = pictureFormatForMime(match[1]);
  if (!format) return null;
  const b64 = match[3].replace(/\s+/g, "");
  if (!b64) return null;
  const bytes = base64ToBytes(b64);
  if (!bytes || bytes.length === 0) return null;
  const head = String.fromCharCode(...bytes.slice(0, 12));
  const validMagic =
    (format.ext === "png" &&
      bytes.length >= 8 &&
      [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every((value, index) => bytes[index] === value)) ||
    (format.ext === "jpeg" && bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) ||
    (format.ext === "gif" && bytes.length >= 6 && (head.startsWith("GIF87a") || head.startsWith("GIF89a"))) ||
    (format.ext === "bmp" && bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) ||
    (format.ext === "webp" && bytes.length >= 12 && head.slice(0, 4) === "RIFF" && head.slice(8, 12) === "WEBP");
  if (!validMagic) return null;
  return { format, bytes, base64: b64 };
}
