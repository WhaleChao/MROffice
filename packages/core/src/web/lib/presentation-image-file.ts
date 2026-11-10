/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { DocumentPictureCompressionAdapter } from "../components/console/editor/native-document/documentPictureCompression";
import { isValidInlineImageSource } from "./inline-image-source";
import { fitOfficePictureBlob } from "./office-picture-envelope";

/**
 * The largest local file. It is uploaded whole as the authoritative DataItem
 * (one upload request carries it); the renderer cache written into the page
 * is fitted to the page envelope's single-string bound separately
 * (lib/office-picture-envelope), the way PowerPoint keeps the original and
 * shows a compressed copy.
 */
export const MAX_PRESENTATION_IMAGE_FILE_BYTES = 50 * 1024 * 1024;
export const MAX_PRESENTATION_IMAGE_DIMENSION = 10_000;
export const MAX_PRESENTATION_IMAGE_PIXELS = 40_000_000;

export type PresentationRasterImageMimeType = "image/png" | "image/jpeg";
export type PresentationImageMimeType =
  | PresentationRasterImageMimeType
  | "image/svg+xml";
export type PresentationImageFileFailureCode =
  | "empty"
  | "unsupported-type"
  | "too-large"
  | "bad-magic"
  | "bad-dimensions"
  | "decode-failed"
  /** An SVG with nothing the picture boundary keeps to draw (svg_sanitizer). */
  | "svg-not-drawable"
  | "cache-failed";

export class PresentationImageFileError extends Error {
  readonly code: PresentationImageFileFailureCode;

  constructor(code: PresentationImageFileFailureCode) {
    super(`presentation_image_file_${code}`);
    this.name = "PresentationImageFileError";
    this.code = code;
  }
}

export interface PreparedPresentationImageFile {
  /**
   * The file to upload when it is not the one chosen: a HEIC, WebP, GIF,
   * BMP, TIFF or AVIF picture is inserted as the PNG or JPEG it was turned
   * into, as PowerPoint converts what it cannot embed as is.
   */
  readonly uploadFile?: File;
  readonly src: string;
  readonly mimeType: PresentationImageMimeType;
  /** Display dimensions after applying EXIF orientation. */
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  readonly sizeBytes: number;
  readonly hasAlpha: boolean;
  readonly exifOrientation: number;
  /** Native Office Graphic identity; absent for ordinary raster pictures. */
  readonly officeGraphicKind?: "svg";
  /** Required compatibility preview paired with the native `asvg:svgBlip`. */
  readonly graphicFallbackSrc?: string;
}

interface ParsedImageMetadata {
  readonly storedWidth: number;
  readonly storedHeight: number;
  readonly hasAlpha: boolean;
  readonly exifOrientation: number;
}

function fail(code: PresentationImageFileFailureCode): never {
  throw new PresentationImageFileError(code);
}

function canonicalMime(value: string): PresentationRasterImageMimeType | null {
  const normalized = value.trim().toLowerCase();
  if (normalized === "image/png") return "image/png";
  if (normalized === "image/jpeg" || normalized === "image/jpg") {
    return "image/jpeg";
  }
  return null;
}

function boundedDimensions(width: number, height: number): void {
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width <= 0 ||
    height <= 0 ||
    width > MAX_PRESENTATION_IMAGE_DIMENSION ||
    height > MAX_PRESENTATION_IMAGE_DIMENSION ||
    width * height > MAX_PRESENTATION_IMAGE_PIXELS
  ) {
    fail("bad-dimensions");
  }
}

function parsePng(bytes: Uint8Array): ParsedImageMetadata {
  const signature = [137, 80, 78, 71, 13, 10, 26, 10] as const;
  if (
    bytes.length < 33 ||
    signature.some((value, index) => bytes[index] !== value)
  ) {
    fail("bad-magic");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    view.getUint32(8, false) !== 13 ||
    String.fromCharCode(...bytes.subarray(12, 16)) !== "IHDR"
  ) {
    fail("decode-failed");
  }
  const storedWidth = view.getUint32(16, false);
  const storedHeight = view.getUint32(20, false);
  boundedDimensions(storedWidth, storedHeight);
  const colorType = bytes[25];
  if (![0, 2, 3, 4, 6].includes(colorType)) fail("decode-failed");

  let offset = 8;
  let hasTransparencyChunk = false;
  let sawEnd = false;
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset, false);
    if (length > bytes.length - offset - 12) fail("decode-failed");
    const kind = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8));
    if (kind === "tRNS") hasTransparencyChunk = true;
    offset += 12 + length;
    if (kind === "IEND") {
      sawEnd = true;
      break;
    }
  }
  if (!sawEnd) fail("decode-failed");
  return {
    storedWidth,
    storedHeight,
    hasAlpha: colorType === 4 || colorType === 6 || hasTransparencyChunk,
    exifOrientation: 1,
  };
}

function readExifOrientation(segment: Uint8Array): number {
  if (
    segment.length < 14 ||
    String.fromCharCode(...segment.subarray(0, 6)) !== "Exif\0\0"
  ) {
    return 1;
  }
  const tiffOffset = 6;
  const byteOrder = String.fromCharCode(
    segment[tiffOffset],
    segment[tiffOffset + 1],
  );
  const littleEndian = byteOrder === "II";
  if (!littleEndian && byteOrder !== "MM") return 1;
  const view = new DataView(
    segment.buffer,
    segment.byteOffset,
    segment.byteLength,
  );
  if (view.getUint16(tiffOffset + 2, littleEndian) !== 42) return 1;
  const firstIfd = view.getUint32(tiffOffset + 4, littleEndian);
  const ifdOffset = tiffOffset + firstIfd;
  if (ifdOffset + 2 > segment.length) return 1;
  const entries = view.getUint16(ifdOffset, littleEndian);
  if (entries > 512 || ifdOffset + 2 + entries * 12 > segment.length) return 1;
  for (let index = 0; index < entries; index += 1) {
    const entryOffset = ifdOffset + 2 + index * 12;
    if (view.getUint16(entryOffset, littleEndian) !== 0x0112) continue;
    if (
      view.getUint16(entryOffset + 2, littleEndian) !== 3 ||
      view.getUint32(entryOffset + 4, littleEndian) !== 1
    ) {
      return 1;
    }
    const orientation = view.getUint16(entryOffset + 8, littleEndian);
    return orientation >= 1 && orientation <= 8 ? orientation : 1;
  }
  return 1;
}

function parseJpeg(bytes: Uint8Array): ParsedImageMetadata {
  // Only the start is checked: phones and cameras append data after the
  // end-of-image marker (a second image, an HDR gain map, a thumbnail), and
  // every decoder ignores it.
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8 || bytes[2] !== 0xff) {
    fail("bad-magic");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const startOfFrameMarkers = new Set([
    0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce,
    0xcf,
  ]);
  let storedWidth = 0;
  let storedHeight = 0;
  let exifOrientation = 1;
  let offset = 2;
  while (offset + 1 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1;
    if (offset >= bytes.length) break;
    const marker = bytes[offset++];
    if (marker === 0xd9 || marker === 0xda) break;
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 2 > bytes.length) fail("decode-failed");
    const segmentLength = view.getUint16(offset, false);
    if (segmentLength < 2 || offset + segmentLength > bytes.length) {
      fail("decode-failed");
    }
    const payloadStart = offset + 2;
    const payloadEnd = offset + segmentLength;
    if (marker === 0xe1) {
      exifOrientation = readExifOrientation(
        bytes.subarray(payloadStart, payloadEnd),
      );
    }
    if (startOfFrameMarkers.has(marker)) {
      if (segmentLength < 7) fail("decode-failed");
      storedHeight = view.getUint16(payloadStart + 1, false);
      storedWidth = view.getUint16(payloadStart + 3, false);
    }
    offset += segmentLength;
  }
  boundedDimensions(storedWidth, storedHeight);
  return {
    storedWidth,
    storedHeight,
    hasAlpha: false,
    exifOrientation,
  };
}

async function verifyBrowserDecode(
  file: File,
  expectedWidth: number,
  expectedHeight: number,
): Promise<void> {
  const decode = globalThis.createImageBitmap;
  if (typeof decode !== "function") return;
  let bitmap: ImageBitmap | null = null;
  try {
    bitmap = await decode(file, { imageOrientation: "from-image" });
    boundedDimensions(bitmap.width, bitmap.height);
    if (bitmap.width !== expectedWidth || bitmap.height !== expectedHeight) {
      fail("decode-failed");
    }
  } catch {
    fail("decode-failed");
  } finally {
    bitmap?.close();
  }
}

async function readFileBytes(file: File): Promise<Uint8Array> {
  if (typeof file.arrayBuffer === "function") {
    return new Uint8Array(await file.arrayBuffer());
  }
  if (typeof FileReader === "undefined") fail("decode-failed");
  return new Promise<Uint8Array>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (reader.result instanceof ArrayBuffer) {
        resolve(new Uint8Array(reader.result));
      } else {
        reject(new PresentationImageFileError("decode-failed"));
      }
    };
    reader.onerror = () =>
      reject(new PresentationImageFileError("decode-failed"));
    reader.readAsArrayBuffer(file);
  });
}

/** Pictures a browser decodes that a deck stores as PNG or JPEG. */
const CONVERTIBLE_IMAGE_TYPES: ReadonlyMap<string, string> = new Map([
  ["image/heic", "heic"],
  ["image/heif", "heif"],
  ["image/webp", "webp"],
  ["image/gif", "gif"],
  ["image/bmp", "bmp"],
  ["image/x-ms-bmp", "bmp"],
  ["image/tiff", "tiff"],
  ["image/avif", "avif"],
]);
const CONVERTIBLE_EXTENSIONS: ReadonlyMap<string, string> = new Map([
  [".heic", "image/heic"],
  [".heif", "image/heif"],
  [".webp", "image/webp"],
  [".gif", "image/gif"],
  [".bmp", "image/bmp"],
  [".tif", "image/tiff"],
  [".tiff", "image/tiff"],
  [".avif", "image/avif"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
]);

function declaredImageType(file: File): string {
  const declared = file.type.trim().toLowerCase();
  if (declared && declared !== "application/octet-stream") return declared;
  const dot = file.name.lastIndexOf(".");
  return CONVERTIBLE_EXTENSIONS.get(dot >= 0 ? file.name.slice(dot).toLowerCase() : "") ?? declared;
}

/** The picture as a PNG (when any pixel is transparent) or a JPEG. */
async function convertedPicture(file: File): Promise<File> {
  const decode = globalThis.createImageBitmap;
  if (typeof decode !== "function" || typeof document === "undefined") {
    fail("unsupported-type");
  }
  let bitmap: ImageBitmap;
  try {
    bitmap = await decode(file, { imageOrientation: "from-image" });
  } catch {
    // A HEIC photo outside Safari, or a file whose name claims a picture.
    fail("unsupported-type");
  }
  try {
    boundedDimensions(bitmap.width, bitmap.height);
    const canvas = document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    const context = canvas.getContext("2d");
    if (!context) fail("decode-failed");
    context.drawImage(bitmap, 0, 0);
    const sample = document.createElement("canvas");
    sample.width = Math.min(64, bitmap.width);
    sample.height = Math.min(64, bitmap.height);
    const sampleContext = sample.getContext("2d");
    let transparent = false;
    if (sampleContext) {
      sampleContext.drawImage(bitmap, 0, 0, sample.width, sample.height);
      const pixels = sampleContext.getImageData(0, 0, sample.width, sample.height).data;
      for (let index = 3; index < pixels.length; index += 4) {
        if (pixels[index] < 255) {
          transparent = true;
          break;
        }
      }
    }
    const mimeType = transparent ? "image/png" : "image/jpeg";
    const blob = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob(resolve, mimeType, 0.92),
    );
    if (!blob || blob.type !== mimeType) fail("decode-failed");
    const stem = file.name.replace(/\.[^.]+$/, "") || "picture";
    return new File([blob], `${stem}.${transparent ? "png" : "jpg"}`, {
      type: mimeType,
      lastModified: file.lastModified,
    });
  } finally {
    bitmap.close();
  }
}

/**
 * Validate a single local file completely before upload or deck mutation.
 * `fitAdapter` stands in for the browser decoder/encoder that fits an
 * oversize renderer cache (tests).
 *
 * PNG and JPEG are inserted as they are; HEIC, WebP, GIF, BMP, TIFF and AVIF
 * are converted (where this browser decodes them) and the converted file is
 * what is uploaded.
 */
export async function preparePresentationImageFile(
  file: File,
  options: { readonly fitAdapter?: DocumentPictureCompressionAdapter } = {},
): Promise<PreparedPresentationImageFile> {
  if (file.size <= 0) fail("empty");
  if (file.size > MAX_PRESENTATION_IMAGE_FILE_BYTES) fail("too-large");
  const declared = declaredImageType(file);
  let source = file;
  let converted: File | undefined;
  if (!canonicalMime(declared)) {
    if (!CONVERTIBLE_IMAGE_TYPES.has(declared)) fail("unsupported-type");
    converted = await convertedPicture(file);
    if (converted.size > MAX_PRESENTATION_IMAGE_FILE_BYTES) fail("too-large");
    source = converted;
  }
  const mimeType = canonicalMime(source.type || declared);
  if (!mimeType) fail("unsupported-type");

  let bytes: Uint8Array;
  try {
    bytes = await readFileBytes(source);
  } catch {
    fail("decode-failed");
  }
  if (bytes.byteLength !== source.size) fail("decode-failed");
  const metadata =
    mimeType === "image/png" ? parsePng(bytes) : parseJpeg(bytes);

  const rotates =
    metadata.exifOrientation >= 5 && metadata.exifOrientation <= 8;
  const pixelWidth = rotates ? metadata.storedHeight : metadata.storedWidth;
  const pixelHeight = rotates ? metadata.storedWidth : metadata.storedHeight;
  boundedDimensions(pixelWidth, pixelHeight);
  await verifyBrowserDecode(source, pixelWidth, pixelHeight);
  // The cache the page holds, at a length the page envelope admits; the
  // block keeps the original's size and dimensions (the DataItem is whole).
  let src: string | null;
  try {
    src = await fitOfficePictureBlob(
      new Blob([source], { type: mimeType }),
      mimeType,
      options.fitAdapter,
    );
  } catch {
    src = null;
  }
  if (src === null || !isValidInlineImageSource(src)) fail("cache-failed");
  return Object.freeze({
    ...(converted ? { uploadFile: converted } : {}),
    src,
    mimeType,
    pixelWidth,
    pixelHeight,
    sizeBytes: bytes.byteLength,
    hasAlpha: metadata.hasAlpha,
    exifOrientation: metadata.exifOrientation,
  });
}
