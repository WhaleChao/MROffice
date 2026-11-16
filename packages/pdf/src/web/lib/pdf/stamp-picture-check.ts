/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Unzlib } from "fflate";

import {
  blobDataUrl,
  browserPictureCompressionAdapter,
  type DocumentPictureCompressionAdapter,
} from "@mrnote/core/web/components/console/editor/native-document/documentPictureCompression.ts";

export const PDF_STAMP_PICTURE_MAX_SIDE = 4096;
export const PDF_STAMP_PICTURE_MAX_PIXELS = 16_000_000;

export type PdfStampPictureIssue = "format" | "structure" | "animated" | "dimensions";

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

let crcTable: Uint32Array | null = null;
function crc32(bytes: Uint8Array, start: number, end: number): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
      crcTable[index] = value >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let index = start; index < end; index += 1) crc = crcTable[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** The most pixel data read to check a PNG (a picture within the writer's
 * bounds at 8 bits a sample is under 68 MB; more is drawn again instead). */
const PNG_DATA_CHECK_MAX_BYTES = 72 * 1024 * 1024;

/** Whether the PNG's pixel data is whole as zlib reads it: it inflates to
 * the rows its header names and its own checksum (Adler-32) holds. Never
 * more than the header's rows is inflated, however much the data claims. */
function pngDataIsWhole(parts: readonly Uint8Array[], expected: number): boolean {
  if (expected > PNG_DATA_CHECK_MAX_BYTES) return false;
  let got = 0;
  let a = 1, b = 0;
  let over = false;
  const inflate = new Unzlib((chunk) => {
    got += chunk.length;
    if (got > expected) { over = true; return; }
    for (let index = 0; index < chunk.length;) {
      // (5552 bytes at most before the sums are reduced, as zlib does.)
      const stop = Math.min(chunk.length, index + 5552);
      for (; index < stop; index += 1) { a += chunk[index]; b += a; }
      a %= 65521; b %= 65521;
    }
  });
  try {
    for (let index = 0; index < parts.length && !over; index += 1) inflate.push(parts[index], index === parts.length - 1);
  } catch {
    return false;
  }
  if (over || got !== expected) return false;
  // The stream's last four bytes are the checksum of what it inflates to.
  const tail: number[] = [];
  for (let index = parts.length - 1; index >= 0 && tail.length < 4; index -= 1) {
    for (let at = parts[index].length - 1; at >= 0 && tail.length < 4; at -= 1) tail.unshift(parts[index][at]);
  }
  return tail.length === 4 && (((b << 16) | a) >>> 0) === (((tail[0] << 24) | (tail[1] << 16) | (tail[2] << 8) | tail[3]) >>> 0);
}

function pngIssue(bytes: Uint8Array): PdfStampPictureIssue | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = PNG_SIGNATURE.length;
  let rows: number | null = null;
  const data: Uint8Array[] = [];
  for (;;) {
    if (offset + 12 > bytes.length) return "structure";
    const length = view.getUint32(offset);
    const end = offset + 12 + length;
    if (end > bytes.length) return "structure";
    const type = String.fromCharCode(bytes[offset + 4], bytes[offset + 5], bytes[offset + 6], bytes[offset + 7]);
    if (rows === null) {
      if (type !== "IHDR" || length !== 13) return "structure";
      const width = view.getUint32(offset + 8), height = view.getUint32(offset + 12);
      const depth = bytes[offset + 16], channels = [1, 0, 3, 1, 2, 0, 4][bytes[offset + 17]] ?? 0;
      if (!width || !height || !channels || ![1, 2, 4, 8, 16].includes(depth)) return "structure";
      // (An interlaced picture is laid out in passes: drawn again by the
      // browser rather than measured here.)
      if (bytes[offset + 20]) return "structure";
      rows = height * (1 + Math.ceil((width * channels * depth) / 8));
    }
    // The checksums the writer's reader holds a file to (its header, its
    // palette, its end); the pixel data it holds to inflating whole.
    if ((type === "IHDR" || type === "PLTE" || type === "tRNS" || type === "IEND") &&
      crc32(bytes, offset + 4, offset + 8 + length) !== view.getUint32(offset + 8 + length)) return "structure";
    // (An animated PNG is several frames to the writer.)
    if (type === "acTL") return "animated";
    if (type === "IDAT") data.push(bytes.subarray(offset + 8, offset + 8 + length));
    if (type === "IEND") return data.length && pngDataIsWhole(data, rows) ? null : "structure";
    offset = end;
  }
}

function jpegIssue(bytes: Uint8Array): PdfStampPictureIssue | null {
  // Cut short, a JPEG has no end-of-image marker (nothing but padding may
  // follow it).
  let end = bytes.length;
  while (end > 2 && (bytes[end - 1] === 0x00 || bytes[end - 1] === 0x0a || bytes[end - 1] === 0x0d)) end -= 1;
  return end >= 4 && bytes[end - 2] === 0xff && bytes[end - 1] === 0xd9 ? null : "structure";
}

/** The picture's size as its own header names it (a PNG's IHDR, a JPEG's
 * frame header), or null. */
export function pdfPictureHeaderSize(bytes: Uint8Array): { width: number; height: number } | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (PNG_SIGNATURE.every((value, index) => bytes[index] === value)) {
    return bytes.length >= 24 ? { width: view.getUint32(16), height: view.getUint32(20) } : null;
  }
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  for (let offset = 2; offset + 9 < bytes.length;) {
    if (bytes[offset] !== 0xff) return null;
    const marker = bytes[offset + 1];
    if (marker === 0xff) { offset += 1; continue; }
    // (Start-of-frame markers, bar the tables that share their range.)
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: view.getUint16(offset + 5), width: view.getUint16(offset + 7) };
    }
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) { offset += 2; continue; }
    offset += 2 + view.getUint16(offset + 2);
  }
  return null;
}

/** What the PDF's writer will refuse of this picture, or null. */
export function pdfStampPictureIssue(
  bytes: Uint8Array,
  mimeType: string,
  size: Readonly<{ width: number; height: number }>,
): PdfStampPictureIssue | null {
  const mime = mimeType.trim().toLowerCase();
  const png = PNG_SIGNATURE.every((value, index) => bytes[index] === value);
  const jpeg = bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;
  if (mime === "image/png" ? !png : mime === "image/jpeg" || mime === "image/jpg" ? !jpeg : true) return "format";
  const structure = png ? pngIssue(bytes) : jpegIssue(bytes);
  if (structure) return structure;
  const { width, height } = size;
  if (!(width > 0) || !(height > 0) || width > PDF_STAMP_PICTURE_MAX_SIDE || height > PDF_STAMP_PICTURE_MAX_SIDE ||
    width * height > PDF_STAMP_PICTURE_MAX_PIXELS) return "dimensions";
  return null;
}

function dataUrlBytes(dataUrl: string): { mimeType: string; bytes: Uint8Array } | null {
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([A-Za-z0-9+/]+={0,2})$/i.exec(dataUrl.trim());
  if (!match) return null;
  try {
    const binary = atob(match[2]);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return { mimeType: match[1].toLowerCase(), bytes };
  } catch {
    return null;
  }
}

/**
 * The picture as the writer takes it: `dataUrl` itself when the writer has
 * nothing against it; else the picture the browser draws of it, encoded
 * again within the writer's bounds and `maxBytes`; null when it cannot be
 * drawn (or does not fit).
 */
export async function pdfStampPictureForWriter(input: {
  dataUrl: string;
  /** The picture's size as the browser decoded it, where the caller has
   * decoded it; else its header is read. */
  width?: number;
  height?: number;
  maxBytes: number;
  adapter?: DocumentPictureCompressionAdapter;
}): Promise<string | null> {
  const read = dataUrlBytes(input.dataUrl);
  if (!read) return null;
  const size = input.width !== undefined && input.height !== undefined
    ? { width: input.width, height: input.height }
    : pdfPictureHeaderSize(read.bytes) ?? { width: 0, height: 0 };
  if (pdfStampPictureIssue(read.bytes, read.mimeType, size) === null) return input.dataUrl;
  try {
    const adapter = input.adapter ?? (await browserPictureCompressionAdapter());
    const raster = await adapter.decode(new Blob([read.bytes as BlobPart], { type: read.mimeType }));
    try {
      if (!(raster.width > 0) || !(raster.height > 0)) return null;
      const fit = Math.min(1, PDF_STAMP_PICTURE_MAX_SIDE / Math.max(raster.width, raster.height),
        Math.sqrt(PDF_STAMP_PICTURE_MAX_PIXELS / (raster.width * raster.height)));
      // A PNG stays a PNG (its transparency); past the byte limit it is
      // made smaller, as an oversized picture is.
      const mimeType = read.mimeType === "image/png" ? "image/png" : "image/jpeg";
      for (const scale of [1, 0.7, 0.5, 0.35, 0.25, 0.18, 0.125]) {
        const width = Math.max(1, Math.floor(raster.width * fit * scale));
        const height = Math.max(1, Math.floor(raster.height * fit * scale));
        const encoded = await raster.encode(width, height, mimeType, 0.9);
        if (!encoded || encoded.type !== mimeType || encoded.size <= 0 || encoded.size > input.maxBytes) continue;
        const dataUrl = await blobDataUrl(encoded);
        const again = dataUrlBytes(dataUrl);
        if (again && pdfStampPictureIssue(again.bytes, again.mimeType, { width, height }) === null) return dataUrl;
      }
      return null;
    } finally {
      raster.dispose();
    }
  } catch {
    return null;
  }
}
