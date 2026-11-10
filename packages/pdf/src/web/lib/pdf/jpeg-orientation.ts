/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function jpegExifOrientation(bytes: Uint8Array): number {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return 1;
  let offset = 2;
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return 1;
    const marker = bytes[offset + 1];
    // Fill bytes before a marker.
    if (marker === 0xff) {
      offset += 1;
      continue;
    }
    // Start of scan or end of image: the header segments are over.
    if (marker === 0xda || marker === 0xd9) return 1;
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3];
    if (length < 2 || offset + 2 + length > bytes.length) return 1;
    if (marker === 0xe1) {
      const orientation = exifSegmentOrientation(
        bytes.subarray(offset + 4, offset + 2 + length),
      );
      if (orientation !== null) return orientation;
    }
    offset += 2 + length;
  }
  return 1;
}

/** Orientation from one APP1 payload; `null` when it is not an EXIF block. */
function exifSegmentOrientation(segment: Uint8Array): number | null {
  // "Exif\0\0", then a TIFF header.
  if (
    segment.length < 14 ||
    segment[0] !== 0x45 ||
    segment[1] !== 0x78 ||
    segment[2] !== 0x69 ||
    segment[3] !== 0x66 ||
    segment[4] !== 0 ||
    segment[5] !== 0
  ) {
    return null;
  }
  const tiff = segment.subarray(6);
  const little = tiff[0] === 0x49 && tiff[1] === 0x49;
  const big = tiff[0] === 0x4d && tiff[1] === 0x4d;
  if (!little && !big) return 1;
  const u16 = (at: number) =>
    little ? tiff[at] | (tiff[at + 1] << 8) : (tiff[at] << 8) | tiff[at + 1];
  const u32 = (at: number) =>
    little
      ? (tiff[at] | (tiff[at + 1] << 8) | (tiff[at + 2] << 16) | (tiff[at + 3] << 24)) >>> 0
      : ((tiff[at] << 24) | (tiff[at + 1] << 16) | (tiff[at + 2] << 8) | tiff[at + 3]) >>> 0;
  if (u16(2) !== 0x002a) return 1;
  const directory = u32(4);
  if (directory + 2 > tiff.length) return 1;
  const entries = u16(directory);
  for (let index = 0; index < entries; index += 1) {
    const entry = directory + 2 + index * 12;
    if (entry + 12 > tiff.length) return 1;
    if (u16(entry) !== 0x0112) continue;
    // SHORT, one value, stored in the entry itself.
    if (u16(entry + 2) !== 3 || u32(entry + 4) !== 1) return 1;
    const value = u16(entry + 8);
    return value >= 1 && value <= 8 ? value : 1;
  }
  return 1;
}
