/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The resolution a PNG or JPEG states, written into the encoded bytes: a
 * browser canvas encodes neither (BMP and TIFF state theirs in
 * `raster-encoders`). PowerPoint's slide pictures state their dots per inch
 * (truth export-all-pictures: a PNG's pHYs 3,779 pixels a metre -- 96 dpi, the
 * metre count taken down to a whole number -- and a JPEG's JFIF density
 * 96 x 96 dots per inch).
 */

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/** Pixels a metre for `dpi`, as PowerPoint writes it (whole, rounded down). */
export function pixelsPerMetre(dpi: number): number {
  return Math.floor(dpi / 0.0254);
}

let crcTable: Uint32Array | null = null;
function crc32(bytes: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function u32(value: number): number[] {
  return [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];
}

/**
 * The PNG with one pHYs chunk stating `dpi` (unit: the metre), right after
 * IHDR; a pHYs it already had is replaced. Bytes that are not a PNG are
 * returned as they are.
 */
export function pngWithDpi(png: Uint8Array, dpi: number): Uint8Array {
  if (png.length < 33 || PNG_SIGNATURE.some((byte, index) => png[index] !== byte)) return png;
  const ppm = pixelsPerMetre(dpi);
  const body = new Uint8Array([0x70, 0x48, 0x59, 0x73, ...u32(ppm), ...u32(ppm), 1]);
  const chunk = new Uint8Array([...u32(9), ...body, ...u32(crc32(body))]);
  const kept: Uint8Array[] = [png.subarray(0, 8)];
  let offset = 8;
  let placed = false;
  while (offset + 12 <= png.length) {
    const length = ((png[offset]! << 24) | (png[offset + 1]! << 16) | (png[offset + 2]! << 8) | png[offset + 3]!) >>> 0;
    const end = offset + 12 + length;
    if (end > png.length) return png;
    const type = String.fromCharCode(png[offset + 4]!, png[offset + 5]!, png[offset + 6]!, png[offset + 7]!);
    if (type !== "pHYs") kept.push(png.subarray(offset, end));
    if (type === "IHDR" && !placed) {
      kept.push(chunk);
      placed = true;
    }
    offset = end;
  }
  if (!placed) return png;
  const out = new Uint8Array(kept.reduce((sum, part) => sum + part.length, 0));
  let at = 0;
  for (const part of kept) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/**
 * The JPEG with its JFIF density stating `dpi` x `dpi` dots per inch: the
 * APP0 JFIF segment right after the start of image is rewritten, or one is
 * put there. Bytes that are not a JPEG are returned as they are.
 */
export function jpegWithDpi(jpeg: Uint8Array, dpi: number): Uint8Array {
  if (jpeg.length < 4 || jpeg[0] !== 0xff || jpeg[1] !== 0xd8) return jpeg;
  const density = Math.max(1, Math.min(0xffff, Math.round(dpi)));
  const isJfif =
    jpeg[2] === 0xff &&
    jpeg[3] === 0xe0 &&
    jpeg.length >= 20 &&
    jpeg[6] === 0x4a && jpeg[7] === 0x46 && jpeg[8] === 0x49 && jpeg[9] === 0x46 && jpeg[10] === 0x00;
  if (isJfif) {
    const out = jpeg.slice();
    out[13] = 1;
    out[14] = density >> 8;
    out[15] = density & 0xff;
    out[16] = density >> 8;
    out[17] = density & 0xff;
    return out;
  }
  const app0 = new Uint8Array([
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 1,
    density >> 8, density & 0xff, density >> 8, density & 0xff, 0x00, 0x00,
  ]);
  const out = new Uint8Array(jpeg.length + app0.length);
  out.set(jpeg.subarray(0, 2), 0);
  out.set(app0, 2);
  out.set(jpeg.subarray(2), 2 + app0.length);
  return out;
}

/** The picture (`format` png or jpeg) stating `dpi`; any other kept as it is. */
export async function pictureWithDpi(picture: Blob, format: string, dpi: number): Promise<Blob> {
  if (format !== "png" && format !== "jpeg") return picture;
  const bytes = new Uint8Array(await picture.arrayBuffer());
  const stamped = format === "png" ? pngWithDpi(bytes, dpi) : jpegWithDpi(bytes, dpi);
  return stamped === bytes ? picture : new Blob([stamped as Uint8Array<ArrayBuffer>], { type: picture.type });
}
