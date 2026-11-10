/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture sizing and Crop > Fill / Fit, shared by Docs, the deck and the
 * worksheet.
 *
 * Office semantics (Microsoft Support, "Crop a picture to fit in a shape" and
 * "Crop a picture"): Fill keeps the picture's frame and crops the source to
 * the frame's aspect ratio about its centre, so the whole frame is covered;
 * Fit keeps the frame and shows the whole source inside it, letterboxed --
 * DrawingML states both as the picture's `a:srcRect` (Fit with negative,
 * outset edges). An inserted picture takes its natural size at 96 DPI,
 * scaled down together on both axes only when it would not fit.
 *
 * The crop math is the deck's (`pictureFillCropRect` / `pictureFitCropRect`
 * in `presentation-crop-group`), reused rather than restated.
 */

import {
  pictureFillCropRect,
  pictureFitCropRect,
  type CropRect,
} from "./presentation-crop-group";

export interface OfficePixelSize {
  readonly width: number;
  readonly height: number;
}

/**
 * The picture's own pixel size, or null when it cannot be decoded here. A
 * realm that never decodes (a headless DOM) answers null after `timeoutMs`
 * rather than holding the caller's command open.
 */
export function loadOfficeImageNaturalSize(
  src: unknown,
  timeoutMs = 500,
): Promise<OfficePixelSize | null> {
  if (typeof src !== "string" || !src || typeof Image === "undefined") return Promise.resolve(null);
  return new Promise((resolve) => {
    const image = new Image();
    const timer = setTimeout(() => resolve(null), timeoutMs);
    image.onload = () => {
      clearTimeout(timer);
      resolve(
        image.naturalWidth > 0 && image.naturalHeight > 0
          ? { width: image.naturalWidth, height: image.naturalHeight }
          : null,
      );
    };
    image.onerror = () => {
      clearTimeout(timer);
      resolve(null);
    };
    image.src = src;
  });
}

/** EMU per CSS pixel at Office's 96 DPI baseline (914400 EMU per inch / 96). */
export const EMU_PER_96_DPI_PIXEL = 9_525;







export function officeInsertedPictureSizeEmu(
  natural: OfficePixelSize,
  boundsEmu: OfficePixelSize,
): OfficePixelSize {
  return officeInsertedPictureSize(
    {
      width: natural.width * EMU_PER_96_DPI_PIXEL,
      height: natural.height * EMU_PER_96_DPI_PIXEL,
    },
    boundsEmu,
  );
}

/** An inserted picture's size: natural at 96 DPI, scaled down to fit `bounds`. */
export function officeInsertedPictureSize(
  natural: OfficePixelSize,
  bounds: OfficePixelSize,
): OfficePixelSize {
  const scale = Math.min(
    1,
    Math.max(1, bounds.width) / Math.max(1, natural.width),
    Math.max(1, bounds.height) / Math.max(1, natural.height),
  );
  return {
    width: Math.max(1, Math.round(natural.width * scale)),
    height: Math.max(1, Math.round(natural.height * scale)),
  };
}

/** Crop > Fill / Fit as the kept source rectangle (`a:srcRect`), or null. */
export function officePictureFitCrop(
  mode: "fill" | "fit",
  natural: OfficePixelSize,
  frame: OfficePixelSize,
): CropRect | null {
  return (mode === "fill" ? pictureFillCropRect : pictureFitCropRect)(
    natural.width,
    natural.height,
    frame.width,
    frame.height,
  );
}

/**
 * The same rectangle as the four edge percentages Word's picture crop stores
 * (`pictureCropTop/Right/Bottom/Left`, 0..100 of the source). Null when an
 * edge would be negative, which a percentage crop cannot state (Fit on a
 * picture whose aspect differs from its frame).
 */
export function officeCropRectPercents(
  rect: CropRect | null,
): { top: number; right: number; bottom: number; left: number } | null {
  if (!rect) return null;
  const round = (value: number) => Math.round(value * 1000) / 10;
  const edges = {
    top: round(rect.y),
    right: round(1 - rect.x - rect.w),
    bottom: round(1 - rect.y - rect.h),
    left: round(rect.x),
  };
  return Object.values(edges).every((value) => value >= 0) ? edges : null;
}

/** A JPEG APP1 Exif segment's orientation tag (1-8), else 1. */
function jpegExifOrientation(bytes: Uint8Array, start: number, end: number): number {
  const at = start;
  if (end - at < 14 || String.fromCharCode(...bytes.subarray(at, at + 6)) !== "Exif\0\0") return 1;
  const tiff = at + 6;
  const little = bytes[tiff] === 0x49 && bytes[tiff + 1] === 0x49;
  if (!little && !(bytes[tiff] === 0x4d && bytes[tiff + 1] === 0x4d)) return 1;
  const u16 = (offset: number) =>
    little ? bytes[offset] | (bytes[offset + 1] << 8) : (bytes[offset] << 8) | bytes[offset + 1];
  const u32 = (offset: number) =>
    little
      ? (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
      : ((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]) >>> 0;
  if (u16(tiff + 2) !== 42) return 1;
  const ifd = tiff + u32(tiff + 4);
  if (ifd + 2 > end) return 1;
  const entries = u16(ifd);
  if (entries > 512 || ifd + 2 + entries * 12 > end) return 1;
  for (let index = 0; index < entries; index += 1) {
    const entry = ifd + 2 + index * 12;
    if (u16(entry) !== 0x0112) continue;
    const value = u16(entry + 8);
    return u16(entry + 2) === 3 && value >= 1 && value <= 8 ? value : 1;
  }
  return 1;
}

/**
 * A picture's displayed pixel size read from its own header bytes, for the
 * formats a document stores: PNG (IHDR), JPEG (SOFn; an Exif orientation of
 * 5-8 swaps the axes, as the browser draws it -- `image-orientation:
 * from-image` -- and Compress Pictures decodes it), GIF (logical screen),
 * WebP (VP8 / VP8L / VP8X) and BMP (BITMAPINFOHEADER and the OS/2 core
 * header; a top-down BMP's negative height counts as its size). Null when the
 * header is none of those or is cut short. Pure: nothing is decoded.
 */
export function imagePixelSizeFromBytes(bytes: Uint8Array): OfficePixelSize | null {
  const size = (width: number, height: number) =>
    Number.isInteger(width) && Number.isInteger(height) && width > 0 && height > 0 ? { width, height } : null;
  const be32 = (at: number) => ((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0;
  const le16 = (at: number) => bytes[at] | (bytes[at + 1] << 8);
  const le24 = (at: number) => bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16);
  const le32 = (at: number) => (bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16) | (bytes[at + 3] << 24)) | 0;
  const ascii = (at: number, length: number) => String.fromCharCode(...bytes.subarray(at, at + length));
  if (bytes.length >= 24 && bytes[0] === 0x89 && ascii(1, 3) === "PNG" && ascii(12, 4) === "IHDR") {
    return size(be32(16), be32(20));
  }
  if (bytes.length >= 10 && ascii(0, 4) === "GIF8") return size(le16(6), le16(8));
  if (bytes.length >= 30 && ascii(0, 4) === "RIFF" && ascii(8, 4) === "WEBP") {
    const chunk = ascii(12, 4);
    if (chunk === "VP8 " && bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a) {
      return size(le16(26) & 0x3fff, le16(28) & 0x3fff);
    }
    if (chunk === "VP8L" && bytes[20] === 0x2f) {
      const bits = (bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24)) >>> 0;
      return size((bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1);
    }
    if (chunk === "VP8X") return size(le24(24) + 1, le24(27) + 1);
    return null;
  }
  if (bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d) {
    const header = le32(14);
    if (header === 12) return size(le16(18), le16(20));
    if (header >= 40 && bytes.length >= 26) return size(le32(18), Math.abs(le32(22)));
    return null;
  }
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let orientation = 1;
    let at = 2;
    while (at + 4 <= bytes.length) {
      if (bytes[at] !== 0xff) return null;
      const marker = bytes[at + 1];
      if (marker === 0xff) {
        at += 1;
        continue;
      }
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        at += 2;
        continue;
      }
      if (marker === 0xd9 || marker === 0xda) return null;
      const length = (bytes[at + 2] << 8) | bytes[at + 3];
      if (length < 2) return null;
      if (marker === 0xe1) orientation = jpegExifOrientation(bytes, at + 4, Math.min(bytes.length, at + 2 + length));
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        if (at + 9 > bytes.length) return null;
        const height = (bytes[at + 5] << 8) | bytes[at + 6];
        const width = (bytes[at + 7] << 8) | bytes[at + 8];
        return orientation >= 5 ? size(height, width) : size(width, height);
      }
      at += 2 + length;
    }
  }
  return null;
}

function base64Bytes(data: string): Uint8Array | null {
  if (typeof atob !== "function") return null;
  try {
    const binary = atob(data.slice(0, data.length - (data.length % 4)));
    return Uint8Array.from(binary, (character) => character.charCodeAt(0));
  } catch {
    return null;
  }
}

/** Base64 characters read first: enough for every header but a JPEG whose
 *  Exif thumbnail and ICC segments come before its frame header. */
const HEADER_PREFIX_CHARS = 262_144;
/** The most of a data URL ever decoded for its header (the documents' own
 *  inline picture limit is 2,000,000 characters). */
const MAX_HEADER_SCAN_CHARS = 16_000_000;

/**
 * A picture's pixel size read from its data URL's own header (see
 * `imagePixelSizeFromBytes`), synchronously; null when it is not a base64
 * image data URL of a known format. A paste needs the size before it
 * commits, and a decode never completes in a headless DOM (PPT-B242), so the
 * document's headless Crop > Fill / Fit / Aspect Ratio read it here too.
 */
export function imagePixelSizeFromDataUrl(dataUrl: unknown): OfficePixelSize | null {
  if (typeof dataUrl !== "string") return null;
  const comma = dataUrl.indexOf(",");
  if (!dataUrl.startsWith("data:image/") || comma < 0 || !/;base64,/i.test(dataUrl.slice(0, comma + 1))) return null;
  const data = dataUrl.slice(comma + 1);
  const head = base64Bytes(data.slice(0, HEADER_PREFIX_CHARS));
  if (!head) return null;
  const found = imagePixelSizeFromBytes(head);
  // A JPEG whose frame header lies past the prefix: read on, within bounds.
  if (found || data.length <= HEADER_PREFIX_CHARS || head[0] !== 0xff || head[1] !== 0xd8) return found;
  const whole = base64Bytes(data.slice(0, MAX_HEADER_SCAN_CHARS));
  return whole ? imagePixelSizeFromBytes(whole) : null;
}
