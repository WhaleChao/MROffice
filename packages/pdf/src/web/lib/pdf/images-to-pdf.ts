/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client-side "images → PDF" builder for the notebook's create-page surface.
 *
 * A user picks one or more image files; we build a PDF with ONE PAGE PER
 * IMAGE, each page sized to the image's pixel dimensions (1 px == 1 pt, the
 * 72-dpi convention used by Preview / Acrobat "create from images"), then the
 * resulting File goes through the EXISTING `createPdfPageFromFile` upload
 * path — so the produced page behaves exactly like an uploaded PDF.
 *
 * EXIF orientation: a photo straight off a phone often stores its rotation
 * only in EXIF metadata. `PDFDocument.embedJpg` embeds the RAW bytes (no EXIF
 * applied), so we decode each file through
 * `createImageBitmap(file, { imageOrientation: "from-image" })` — which bakes
 * the orientation into the pixels — and re-encode via a canvas before
 * embedding. JPEG sources re-encode as JPEG (on an opaque white backing;
 * JPEG has no alpha), everything else (PNG / WebP / GIF first frame / BMP …)
 * re-encodes as lossless PNG. When `createImageBitmap` is unavailable the
 * decode fails with a clear error rather than silently embedding sideways
 * photos.
 *
 * Split so each half is testable in isolation:
 *   - {@link decodeImageFileForPdf} — the DOM boundary (createImageBitmap +
 *     canvas), with injectable deps so tests stub exactly that boundary;
 *   - {@link buildPdfFromImages} — pure pdf-lib assembly from decoded images;
 *   - {@link imagesToPdfFile} / {@link blankPdfFile} — the composed helpers
 *     the create-page menu calls.
 */

import { PDFDocument } from "pdf-lib";

/** US-Letter page size in PDF points (the blank-PDF starter page). */
export const US_LETTER_WIDTH = 612;
export const US_LETTER_HEIGHT = 792;
export const BLANK_PDF_MAX_DIMENSION = 14_400;

export interface BlankPdfPageSize {
  width: number;
  height: number;
}

/** Hard cap on images per conversion (a menu pick, not a batch pipeline). */
export const IMAGES_TO_PDF_MAX_FILES = 50;
/** Per-image source-file byte cap (matches a generous photo upload). */
export const IMAGES_TO_PDF_MAX_BYTES = 30 * 1024 * 1024;
/** Per-side pixel cap so a malformed header can't allocate a giant canvas. */
export const IMAGES_TO_PDF_MAX_DIMENSION = 12_000;

/** JPEG re-encode quality for EXIF-corrected photos. */
const JPEG_REENCODE_QUALITY = 0.92;

/** One decoded, orientation-corrected image ready to embed. */
export interface DecodedImageForPdf {
  /** Encoded image bytes (JPEG or PNG) with orientation baked in. */
  bytes: Uint8Array;
  format: "jpg" | "png";
  /** Pixel dimensions AFTER orientation correction (== the PDF page size). */
  width: number;
  height: number;
}

/** The subset of ImageBitmap the decoder needs (stubable in tests). */
export interface BitmapLike {
  width: number;
  height: number;
  close?: () => void;
}

/** The subset of a 2D canvas the decoder needs (stubable in tests). */
export interface CanvasLike {
  width: number;
  height: number;
  getContext: (id: "2d") => {
    fillStyle: string | CanvasGradient | CanvasPattern;
    fillRect: (x: number, y: number, w: number, h: number) => void;
    drawImage: (image: unknown, dx: number, dy: number) => void;
  } | null;
  toBlob: (
    callback: (blob: Blob | null) => void,
    type?: string,
    quality?: number,
  ) => void;
}

/** Injectable DOM boundary for {@link decodeImageFileForPdf}. */
export interface ImageDecodeDeps {
  createBitmap?: (file: Blob) => Promise<BitmapLike>;
  createCanvas?: (width: number, height: number) => CanvasLike;
}

const IMAGE_EXTENSIONS = new Set([
  ".jpg",
  ".jpeg",
  ".png",
  ".webp",
  ".gif",
  ".bmp",
  ".avif",
]);

/** True for files the images→PDF path accepts (by MIME, else extension). */
export function isImportableImageFile(file: {
  name: string;
  type: string;
}): boolean {
  const mediaType = (file.type || "").split(";", 1)[0].toLowerCase();
  if (mediaType.startsWith("image/")) {
    // SVG is a document, not a raster photo: createImageBitmap support is
    // inconsistent and intrinsic dimensions can be absent — reject it.
    return mediaType !== "image/svg+xml";
  }
  const name = (file.name || "").toLowerCase();
  const index = name.lastIndexOf(".");
  const extension = index >= 0 ? name.slice(index) : "";
  return IMAGE_EXTENSIONS.has(extension);
}

function defaultCreateBitmap(file: Blob): Promise<BitmapLike> {
  if (typeof createImageBitmap !== "function") {
    return Promise.reject(
      new Error("Image decoding is not supported in this browser."),
    );
  }
  // "from-image" applies the EXIF orientation while decoding, so the bitmap's
  // pixels (and width/height) are already upright.
  return createImageBitmap(file, { imageOrientation: "from-image" });
}

function defaultCreateCanvas(width: number, height: number): CanvasLike {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return canvas as unknown as CanvasLike;
}

/** Blob → bytes with a FileReader fallback (older WebKit lacks arrayBuffer). */
function blobToBytes(blob: Blob): Promise<Uint8Array> {
  if (typeof blob.arrayBuffer === "function") {
    return blob.arrayBuffer().then((buffer) => new Uint8Array(buffer));
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
    reader.onerror = () =>
      reject(reader.error || new Error("Image read failed."));
    reader.readAsArrayBuffer(blob);
  });
}

function canvasToBlob(
  canvas: CanvasLike,
  type: string,
  quality?: number,
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (blob) resolve(blob);
        else reject(new Error("Image re-encoding failed."));
      },
      type,
      quality,
    );
  });
}

/**
 * Decode one image file into orientation-corrected encoded bytes + dimensions.
 * Throws on unreadable / degenerate / oversized images so a bad file surfaces
 * a clear error instead of an empty PDF page.
 */
export async function decodeImageFileForPdf(
  file: File,
  deps: ImageDecodeDeps = {},
): Promise<DecodedImageForPdf> {
  if (file.size > IMAGES_TO_PDF_MAX_BYTES) {
    throw new Error("Image file is too large to convert to PDF.");
  }
  const createBitmap = deps.createBitmap ?? defaultCreateBitmap;
  const createCanvas = deps.createCanvas ?? defaultCreateCanvas;
  const bitmap = await createBitmap(file);
  try {
    const width = Math.floor(bitmap.width);
    const height = Math.floor(bitmap.height);
    if (
      !Number.isFinite(width) ||
      !Number.isFinite(height) ||
      width <= 0 ||
      height <= 0
    ) {
      throw new Error("Image has no usable dimensions.");
    }
    if (
      width > IMAGES_TO_PDF_MAX_DIMENSION ||
      height > IMAGES_TO_PDF_MAX_DIMENSION
    ) {
      throw new Error("Image dimensions are too large to convert to PDF.");
    }
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      throw new Error("2d canvas context unavailable.");
    }
    const isJpeg = (file.type || "").toLowerCase() === "image/jpeg";
    if (isJpeg) {
      // JPEG has no alpha channel: back the canvas with opaque white so any
      // transparent edge pixels don't encode as black.
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, width, height);
    }
    ctx.drawImage(bitmap, 0, 0);
    const blob = await canvasToBlob(
      canvas,
      isJpeg ? "image/jpeg" : "image/png",
      isJpeg ? JPEG_REENCODE_QUALITY : undefined,
    );
    return {
      bytes: await blobToBytes(blob),
      format: isJpeg ? "jpg" : "png",
      width,
      height,
    };
  } finally {
    bitmap.close?.();
  }
}

/**
 * Assemble a PDF from decoded images: one page per image, page size == image
 * pixel size (1 px == 1 pt), image drawn to fill the page. Order preserved.
 */
export async function buildPdfFromImages(
  images: ReadonlyArray<DecodedImageForPdf>,
): Promise<Uint8Array> {
  if (!Array.isArray(images) || images.length === 0) {
    throw new Error("No images to convert.");
  }
  if (images.length > IMAGES_TO_PDF_MAX_FILES) {
    throw new Error("Too many images to convert in one PDF.");
  }
  const doc = await PDFDocument.create();
  for (const image of images) {
    const embedded =
      image.format === "jpg"
        ? await doc.embedJpg(image.bytes)
        : await doc.embedPng(image.bytes);
    const page = doc.addPage([image.width, image.height]);
    page.drawImage(embedded, {
      x: 0,
      y: 0,
      width: image.width,
      height: image.height,
    });
  }
  return doc.save();
}

/** Build one blank page, defaulting to US Letter for the create-page starter. */
export async function buildBlankPdfBytes(
  size?: BlankPdfPageSize,
): Promise<Uint8Array> {
  const width = size?.width ?? US_LETTER_WIDTH;
  const height = size?.height ?? US_LETTER_HEIGHT;
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= 0 ||
    height <= 0 ||
    width > BLANK_PDF_MAX_DIMENSION ||
    height > BLANK_PDF_MAX_DIMENSION
  ) {
    throw new Error("Blank PDF page dimensions are invalid.");
  }
  const doc = await PDFDocument.create();
  doc.addPage([width, height]);
  return doc.save();
}

/** Derive "photos.pdf" from the first image's filename (fallback provided). */
export function pdfFileNameForImages(
  files: ReadonlyArray<{ name: string }>,
  fallback: string,
): string {
  const first = files[0]?.name || "";
  const base = first.replace(/\.[^.]+$/, "").trim();
  return `${base || fallback}.pdf`;
}

function bytesToPdfFile(bytes: Uint8Array, name: string): File {
  // Copy into a fresh ArrayBuffer-backed chunk so the File never aliases a
  // larger buffer (pdf-lib returns an exact-size view, but be defensive).
  const copy = new Uint8Array(bytes);
  return new File([copy], name, { type: "application/pdf" });
}

/**
 * Convert picked image files into a single PDF File ready for the existing
 * `createPdfPageFromFile` upload path. Rejects non-image files up front so a
 * stray pick fails with a clear error instead of a corrupt page.
 */
export async function imagesToPdfFile(
  files: ReadonlyArray<File>,
  options: { fallbackName: string; deps?: ImageDecodeDeps },
): Promise<File> {
  if (files.length === 0) {
    throw new Error("No images to convert.");
  }
  if (files.length > IMAGES_TO_PDF_MAX_FILES) {
    throw new Error("Too many images to convert in one PDF.");
  }
  for (const file of files) {
    if (!isImportableImageFile(file)) {
      throw new Error(`Not an image file: ${file.name}`);
    }
  }
  const decoded: DecodedImageForPdf[] = [];
  for (const file of files) {
    decoded.push(await decodeImageFileForPdf(file, options.deps));
  }
  const bytes = await buildPdfFromImages(decoded);
  return bytesToPdfFile(bytes, pdfFileNameForImages(files, options.fallbackName));
}

/** Build the blank-PDF starter as a File for the same upload path. */
export async function blankPdfFile(
  baseName: string,
  size?: BlankPdfPageSize,
): Promise<File> {
  const bytes = await buildBlankPdfBytes(size);
  const name = `${(baseName || "").trim() || "blank"}.pdf`;
  return bytesToPdfFile(bytes, name);
}
