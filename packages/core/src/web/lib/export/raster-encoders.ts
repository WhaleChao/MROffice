/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bitmap files written from a canvas's pixels, for the picture formats a
 * browser canvas does not encode itself: PowerPoint's Save As / Export of a
 * slide offers Device Independent Bitmap (.bmp) and Tag Image File Format
 * (.tif) beside PNG and JPEG (Microsoft Support, "File formats that are
 * supported in PowerPoint").
 *
 * - BMP: a BITMAPFILEHEADER and a BITMAPINFOHEADER, then 24-bit BGR rows
 *   bottom-up, each padded to four bytes, uncompressed (BI_RGB), the form
 *   every Windows and macOS reader opens.
 * - TIFF: a baseline TIFF 6.0 RGB image, little-endian ("II*"), one strip of
 *   8-bit RGB samples, uncompressed, with its resolution stated.
 *
 * Both are written opaque: transparent pixels are composited over the
 * background the caller drew (the slide's own), as PowerPoint writes a
 * slide picture.
 */

export interface RasterPixels {
  readonly width: number;
  readonly height: number;
  /** RGBA, row-major, top-down: a canvas `ImageData.data`. */
  readonly data: Uint8ClampedArray;
}

const INCH_PER_METRE = 39.3700787;

function opaque(value: number, alpha: number): number {
  // Composited over white: what an unpainted canvas pixel shows.
  return Math.round((value * alpha + 255 * (255 - alpha)) / 255);
}

function assertPixels(pixels: RasterPixels): void {
  const { width, height, data } = pixels;
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new RangeError("raster size");
  }
  if (data.length !== width * height * 4) throw new RangeError("raster data");
}

/** A 24-bit uncompressed Windows bitmap of `pixels`, `dpi` dots per inch. */
export function encodeBmp(pixels: RasterPixels, dpi = 96): Uint8Array {
  assertPixels(pixels);
  const { width, height, data } = pixels;
  const rowBytes = Math.ceil((width * 3) / 4) * 4;
  const imageBytes = rowBytes * height;
  const out = new Uint8Array(54 + imageBytes);
  const view = new DataView(out.buffer);
  // BITMAPFILEHEADER
  out[0] = 0x42; // B
  out[1] = 0x4d; // M
  view.setUint32(2, out.length, true);
  view.setUint32(10, 54, true);
  // BITMAPINFOHEADER
  const ppm = Math.round(dpi * INCH_PER_METRE);
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true); // positive: rows bottom-up
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(30, 0, true); // BI_RGB
  view.setUint32(34, imageBytes, true);
  view.setInt32(38, ppm, true);
  view.setInt32(42, ppm, true);
  for (let y = 0; y < height; y += 1) {
    const source = (height - 1 - y) * width * 4;
    let target = 54 + y * rowBytes;
    for (let x = 0; x < width; x += 1) {
      const at = source + x * 4;
      const alpha = data[at + 3];
      out[target] = opaque(data[at + 2], alpha);
      out[target + 1] = opaque(data[at + 1], alpha);
      out[target + 2] = opaque(data[at], alpha);
      target += 3;
    }
  }
  return out;
}

const TIFF_SHORT = 3;
const TIFF_LONG = 4;
const TIFF_RATIONAL = 5;

/** A baseline TIFF 6.0 RGB image of `pixels` (uncompressed), `dpi` dots per inch. */
export function encodeTiff(pixels: RasterPixels, dpi = 96): Uint8Array {
  assertPixels(pixels);
  const { width, height, data } = pixels;
  const imageBytes = width * height * 3;
  // Header (8), then the image strip, then the out-of-line values, then the IFD.
  const stripOffset = 8;
  const bitsOffset = stripOffset + imageBytes;
  const xResolutionOffset = bitsOffset + 6;
  const yResolutionOffset = xResolutionOffset + 8;
  let ifdOffset = yResolutionOffset + 8;
  if (ifdOffset % 2) ifdOffset += 1; // an IFD begins on a word boundary
  const entries: Array<[tag: number, type: number, count: number, value: number]> = [
    [256, TIFF_LONG, 1, width], // ImageWidth
    [257, TIFF_LONG, 1, height], // ImageLength
    [258, TIFF_SHORT, 3, bitsOffset], // BitsPerSample 8,8,8
    [259, TIFF_SHORT, 1, 1], // Compression: none
    [262, TIFF_SHORT, 1, 2], // PhotometricInterpretation: RGB
    [273, TIFF_LONG, 1, stripOffset], // StripOffsets
    [277, TIFF_SHORT, 1, 3], // SamplesPerPixel
    [278, TIFF_LONG, 1, height], // RowsPerStrip
    [279, TIFF_LONG, 1, imageBytes], // StripByteCounts
    [282, TIFF_RATIONAL, 1, xResolutionOffset], // XResolution
    [283, TIFF_RATIONAL, 1, yResolutionOffset], // YResolution
    [284, TIFF_SHORT, 1, 1], // PlanarConfiguration: chunky
    [296, TIFF_SHORT, 1, 2], // ResolutionUnit: inch
  ];
  const out = new Uint8Array(ifdOffset + 2 + entries.length * 12 + 4);
  const view = new DataView(out.buffer);
  out[0] = 0x49; // I
  out[1] = 0x49; // I
  view.setUint16(2, 42, true);
  view.setUint32(4, ifdOffset, true);
  let target = stripOffset;
  for (let at = 0; at < data.length; at += 4) {
    const alpha = data[at + 3];
    out[target] = opaque(data[at], alpha);
    out[target + 1] = opaque(data[at + 1], alpha);
    out[target + 2] = opaque(data[at + 2], alpha);
    target += 3;
  }
  view.setUint16(bitsOffset, 8, true);
  view.setUint16(bitsOffset + 2, 8, true);
  view.setUint16(bitsOffset + 4, 8, true);
  const resolution = Math.max(1, Math.round(dpi));
  for (const offset of [xResolutionOffset, yResolutionOffset]) {
    view.setUint32(offset, resolution, true);
    view.setUint32(offset + 4, 1, true);
  }
  view.setUint16(ifdOffset, entries.length, true);
  entries.forEach(([tag, type, count, value], index) => {
    const at = ifdOffset + 2 + index * 12;
    view.setUint16(at, tag, true);
    view.setUint16(at + 2, type, true);
    view.setUint32(at + 4, count, true);
    if (type === TIFF_SHORT && count === 1) view.setUint16(at + 8, value, true);
    else view.setUint32(at + 8, value, true);
  });
  view.setUint32(ifdOffset + 2 + entries.length * 12, 0, true); // no next IFD
  return out;
}
