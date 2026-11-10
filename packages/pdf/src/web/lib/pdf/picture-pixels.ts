/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pixels of a picture a PDF page paints, at the picture's own
 * resolution, for the reader's picture commands that work on pixels: 清晰化
 * (written back as a native replacement), a copy pasted as page content,
 * 保存图片 and 图片另存为PDF when the PDF does not hold the picture as a
 * file of its own.
 *
 * pdf.js decodes every picture of a page it reads into the page's object
 * store (a picture shared by several pages into the document's), keyed by
 * the id its paint operator carries -- the `name` the image walker reports.
 * Reading the operator list again re-requests what a page cleanup released.
 */

/** pdf.js `ImageKind`. */
const IMAGE_KIND_GRAYSCALE_1BPP = 1;
const IMAGE_KIND_RGB_24BPP = 2;
const IMAGE_KIND_RGBA_32BPP = 3;

/** A decoded pdf.js image object (`page.objs` / `page.commonObjs` entry). */
export interface PdfDecodedImageLike {
  readonly width?: number;
  readonly height?: number;
  readonly kind?: number;
  readonly data?: ArrayLike<number>;
  readonly bitmap?: CanvasImageSource & { readonly width: number; readonly height: number };
}

interface PdfObjectStoreLike {
  get(objId: string, callback: (data: unknown) => void): unknown;
}

/** The parts of a pdf.js page proxy the reader uses. */
export interface PdfPicturePageLike {
  readonly objs?: PdfObjectStoreLike;
  readonly commonObjs?: PdfObjectStoreLike;
  getOperatorList?: () => Promise<unknown>;
}

/** Largest picture read for a command (the native writer's own pixel cap). */
export const PICTURE_PIXELS_MAX_AREA = 40_000_000;

/** The picture's pixels as a canvas, or null when pdf.js cannot hand them
 * over (an inline image, a picture it never decoded, a timeout). */
export async function readPdfPagePicturePixels(
  page: PdfPicturePageLike,
  objId: string,
  timeoutMs = 4_000,
): Promise<HTMLCanvasElement | null> {
  if (!objId) return null;
  try {
    await page.getOperatorList?.();
  } catch {
    return null;
  }
  const store = objId.startsWith("g_") ? page.commonObjs : page.objs;
  if (!store) return null;
  const image = await new Promise<unknown>((resolve) => {
    const timer = setTimeout(() => resolve(null), timeoutMs);
    try {
      store.get(objId, (data) => {
        clearTimeout(timer);
        resolve(data);
      });
    } catch {
      clearTimeout(timer);
      resolve(null);
    }
  });
  return image && typeof image === "object"
    ? pdfDecodedImageToCanvas(image as PdfDecodedImageLike)
    : null;
}

/**
 * A decoded pdf.js image's pixels as RGBA, upright (its first row at the
 * top), without a canvas: the reader draws them, a served write (sharpen)
 * reworks them where no canvas exists. Null for a bitmap (drawn only) or a
 * kind or size the picture tools do not read.
 */
export function pdfDecodedImageRgba(
  image: PdfDecodedImageLike,
): { width: number; height: number; data: Uint8ClampedArray } | null {
  const width = image.width ?? 0;
  const height = image.height ?? 0;
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width < 1 ||
    height < 1 ||
    width * height > PICTURE_PIXELS_MAX_AREA
  ) {
    return null;
  }
  const source = image.data;
  if (!source) return null;
  const out = new Uint8ClampedArray(width * height * 4);
  if (image.kind === IMAGE_KIND_RGBA_32BPP) {
    if (source.length < width * height * 4) return null;
    for (let index = 0; index < width * height * 4; index += 1) out[index] = source[index];
  } else if (image.kind === IMAGE_KIND_RGB_24BPP) {
    if (source.length < width * height * 3) return null;
    for (let pixel = 0; pixel < width * height; pixel += 1) {
      out[pixel * 4] = source[pixel * 3];
      out[pixel * 4 + 1] = source[pixel * 3 + 1];
      out[pixel * 4 + 2] = source[pixel * 3 + 2];
      out[pixel * 4 + 3] = 255;
    }
  } else if (image.kind === IMAGE_KIND_GRAYSCALE_1BPP) {
    // One bit per pixel, rows padded to a byte; 0 is black, 1 white.
    const rowBytes = (width + 7) >> 3;
    if (source.length < rowBytes * height) return null;
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const bit = (source[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1;
        const offset = (y * width + x) * 4;
        const value = bit ? 255 : 0;
        out[offset] = value;
        out[offset + 1] = value;
        out[offset + 2] = value;
        out[offset + 3] = 255;
      }
    }
  } else {
    return null;
  }
  return { width, height, data: out };
}

/** A decoded pdf.js image drawn upright (its first row at the top). */
export function pdfDecodedImageToCanvas(
  image: PdfDecodedImageLike,
  doc: Document = document,
): HTMLCanvasElement | null {
  const bitmap = image.bitmap;
  const width = bitmap?.width ?? image.width ?? 0;
  const height = bitmap?.height ?? image.height ?? 0;
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width < 1 ||
    height < 1 ||
    width * height > PICTURE_PIXELS_MAX_AREA
  ) {
    return null;
  }
  const canvas = doc.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return null;
  if (bitmap) {
    context.drawImage(bitmap, 0, 0);
    return canvas;
  }
  const rgba = pdfDecodedImageRgba(image);
  if (!rgba) return null;
  const pixels = context.createImageData(width, height);
  pixels.data.set(rgba.data);
  context.putImageData(pixels, 0, 0);
  return canvas;
}

/**
 * The part of a picture its unit-square rect covers (the paint's clip, the
 * visible part): x runs right, y UP from the picture's bottom edge, as in
 * PDF; the canvas holds the picture upright.
 */
export function cropPictureCanvas(
  canvas: HTMLCanvasElement,
  unit: Readonly<{ x: number; y: number; width: number; height: number }>,
  doc: Document = document,
): HTMLCanvasElement {
  if (unit.x <= 1e-6 && unit.y <= 1e-6 && unit.width >= 1 - 1e-6 && unit.height >= 1 - 1e-6) {
    return canvas;
  }
  const left = Math.max(0, Math.floor(unit.x * canvas.width));
  const right = Math.min(canvas.width, Math.ceil((unit.x + unit.width) * canvas.width));
  const top = Math.max(0, Math.floor((1 - unit.y - unit.height) * canvas.height));
  const bottom = Math.min(canvas.height, Math.ceil((1 - unit.y) * canvas.height));
  const width = Math.max(1, right - left);
  const height = Math.max(1, bottom - top);
  const out = doc.createElement("canvas");
  out.width = width;
  out.height = height;
  out.getContext("2d")?.drawImage(canvas, left, top, width, height, 0, 0, width, height);
  return out;
}

/** Whether any pixel of the canvas lets the page through. */
export function canvasHasTransparency(canvas: HTMLCanvasElement): boolean {
  const context = canvas.getContext("2d");
  if (!context) return false;
  const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
  for (let index = 3; index < data.length; index += 4) {
    if (data[index] < 255) return true;
  }
  return false;
}

function canvasBlob(
  canvas: HTMLCanvasElement,
  type: string,
  quality?: number,
): Promise<Blob | null> {
  return new Promise((resolve) => {
    try {
      canvas.toBlob((blob) => resolve(blob), type, quality);
    } catch {
      resolve(null);
    }
  });
}

async function blobBytes(blob: Blob): Promise<Uint8Array> {
  return withoutColorProfile(new Uint8Array(await blob.arrayBuffer()));
}

/**
 * The encoded picture without the colour profile the browser tags it with
 * (a JPEG's APP2 ICC_PROFILE, a PNG's iCCP): the native writers take an
 * untagged picture only (「只支持不带色彩配置文件的 PNG 或 JPEG」, every
 * JPEG a canvas made on a colour-managed screen), and a canvas's pixels are
 * already in the page's RGB. Anything not read as one is returned as it is.
 */
export function withoutColorProfile(bytes: Uint8Array): Uint8Array {
  if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    const kept: Uint8Array[] = [bytes.subarray(0, 2)];
    let cursor = 2;
    let dropped = false;
    while (cursor + 4 <= bytes.length && bytes[cursor] === 0xff) {
      const marker = bytes[cursor + 1];
      // Start of scan: the rest is the picture's data.
      if (marker === 0xda || marker === 0xd9) break;
      const length = (bytes[cursor + 2] << 8) | bytes[cursor + 3];
      if (length < 2 || cursor + 2 + length > bytes.length) return bytes;
      const end = cursor + 2 + length;
      const icc = marker === 0xe2 && length >= 14 &&
        String.fromCharCode(...bytes.subarray(cursor + 4, cursor + 16)) === "ICC_PROFILE\u0000";
      if (icc) dropped = true;
      else kept.push(bytes.subarray(cursor, end));
      cursor = end;
    }
    if (!dropped) return bytes;
    kept.push(bytes.subarray(cursor));
    return concatBytes(kept);
  }
  const PNG = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (bytes.length > 8 && PNG.every((value, index) => bytes[index] === value)) {
    const kept: Uint8Array[] = [bytes.subarray(0, 8)];
    let cursor = 8;
    let dropped = false;
    while (cursor + 12 <= bytes.length) {
      const length = ((bytes[cursor] << 24) >>> 0) + (bytes[cursor + 1] << 16) + (bytes[cursor + 2] << 8) + bytes[cursor + 3];
      const end = cursor + 12 + length;
      if (end > bytes.length) return bytes;
      const type = String.fromCharCode(...bytes.subarray(cursor + 4, cursor + 8));
      if (type === "iCCP") dropped = true;
      else kept.push(bytes.subarray(cursor, end));
      cursor = end;
      if (type === "IEND") break;
    }
    return dropped ? concatBytes(kept) : bytes;
  }
  return bytes;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function scaledCanvas(canvas: HTMLCanvasElement, factor: number): HTMLCanvasElement {
  const out = canvas.ownerDocument.createElement("canvas");
  out.width = Math.max(1, Math.round(canvas.width * factor));
  out.height = Math.max(1, Math.round(canvas.height * factor));
  const context = out.getContext("2d");
  if (context) {
    context.imageSmoothingQuality = "high";
    context.drawImage(canvas, 0, 0, out.width, out.height);
  }
  return out;
}

/**
 * The canvas as a file the native writers accept (PNG or JPEG without a
 * colour profile) within `maxBytes`: a picture that lets the page through
 * stays PNG; an opaque one is PNG when that is reasonably small (line art,
 * screenshots), else JPEG at falling quality. Only a picture no encoding
 * fits is reduced in size, and `downscaled` says so. Null when nothing fits.
 */
export async function encodePictureCanvas(
  canvas: HTMLCanvasElement,
  options: { readonly maxBytes: number; readonly allowDownscale?: boolean },
): Promise<{ bytes: Uint8Array; mimeType: "image/png" | "image/jpeg"; downscaled: boolean } | null> {
  const transparent = canvasHasTransparency(canvas);
  let current = canvas;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const png = await canvasBlob(current, "image/png");
    if (transparent) {
      if (png && png.size <= options.maxBytes) {
        return { bytes: await blobBytes(png), mimeType: "image/png", downscaled: attempt > 0 };
      }
    } else {
      let jpeg: Blob | null = null;
      for (const quality of [0.92, 0.85, 0.75, 0.6]) {
        jpeg = await canvasBlob(current, "image/jpeg", quality);
        if (jpeg && jpeg.size <= options.maxBytes) break;
      }
      if (png && png.size <= options.maxBytes && (!jpeg || png.size <= jpeg.size * 2)) {
        return { bytes: await blobBytes(png), mimeType: "image/png", downscaled: attempt > 0 };
      }
      if (jpeg && jpeg.size <= options.maxBytes) {
        return { bytes: await blobBytes(jpeg), mimeType: "image/jpeg", downscaled: attempt > 0 };
      }
    }
    if (!options.allowDownscale) return null;
    current = scaledCanvas(current, 0.7);
  }
  return null;
}
