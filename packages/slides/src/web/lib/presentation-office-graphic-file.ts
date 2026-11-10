/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ApiRequestError, apiFetchRaw } from "@mrnote/core/web/lib/api.ts";
import { rasterDataUrlToPngBlob } from "@mrnote/core/web/lib/export/svg-to-png.ts";
import {
  fitOfficePictureSource,
  OFFICE_PICTURE_FILE_MAX_BYTES,
} from "@mrnote/core/web/lib/office-picture-envelope.ts";
import {
  isSafeInlineSvgPayload,
  isValidInlineImageSource,
} from "@mrnote/core/web/lib/inline-image-source.ts";
import {
  MAX_PRESENTATION_IMAGE_DIMENSION,
  MAX_PRESENTATION_IMAGE_FILE_BYTES,
  MAX_PRESENTATION_IMAGE_PIXELS,
  PresentationImageFileError,
  type PreparedPresentationImageFile,
} from "@mrnote/core/web/lib/presentation-image-file.ts";

const SVG_MIME = "image/svg+xml" as const;

function fail(): never {
  throw new PresentationImageFileError("decode-failed");
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof btoa !== "function") fail();
  const chunks: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    chunks.push(
      String.fromCharCode(...bytes.subarray(offset, offset + 0x8000)),
    );
  }
  return btoa(chunks.join(""));
}

async function readFileBytes(file: File): Promise<Uint8Array> {
  if (typeof file.arrayBuffer === "function") {
    return new Uint8Array(await file.arrayBuffer());
  }
  if (typeof FileReader === "undefined") fail();
  return new Promise<Uint8Array>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () =>
      reader.result instanceof ArrayBuffer
        ? resolve(new Uint8Array(reader.result))
        : reject(new PresentationImageFileError("decode-failed"));
    reader.onerror = () =>
      reject(new PresentationImageFileError("decode-failed"));
    reader.readAsArrayBuffer(file);
  });
}

function positiveSvgLength(value: string | null): number | null {
  const normalized = (value ?? "").trim();
  if (!/^[0-9]*\.?[0-9]+(?:px)?$/i.test(normalized)) return null;
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

function svgDimensions(markup: string): { width: number; height: number } {
  if (typeof DOMParser === "undefined") fail();
  const document = new DOMParser().parseFromString(markup, SVG_MIME);
  const root = document.documentElement;
  if (root.localName.toLowerCase() !== "svg") fail();
  const viewBox = (root.getAttribute("viewBox") ?? "")
    .split(/[\s,]+/)
    .map(Number)
    .filter(Number.isFinite);
  const width =
    positiveSvgLength(root.getAttribute("width")) ??
    (viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : null);
  const height =
    positiveSvgLength(root.getAttribute("height")) ??
    (viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : null);
  if (width === null || height === null) fail();
  const pixelWidth = Math.round(width);
  const pixelHeight = Math.round(height);
  if (
    pixelWidth <= 0 ||
    pixelHeight <= 0 ||
    pixelWidth > MAX_PRESENTATION_IMAGE_DIMENSION ||
    pixelHeight > MAX_PRESENTATION_IMAGE_DIMENSION ||
    pixelWidth * pixelHeight > MAX_PRESENTATION_IMAGE_PIXELS
  ) {
    fail();
  }
  return { width: pixelWidth, height: pixelHeight };
}

/** The passive SVG the server makes of `bytes` (`POST /attachments/svg/sanitize`). */
async function passiveSvgMarkup(bytes: Uint8Array): Promise<string> {
  let response: Response;
  try {
    response = await apiFetchRaw(
      "/api/v1/attachments/svg/sanitize",
      // A copy is backed by a plain ArrayBuffer, which a Blob part requires.
      { method: "POST", body: new Blob([bytes.slice()], { type: SVG_MIME }) },
      { contentType: SVG_MIME },
    );
  } catch (error) {
    if (error instanceof ApiRequestError && error.code === "svg_not_drawable") {
      throw new PresentationImageFileError("svg-not-drawable");
    }
    throw error;
  }
  const body = (await response.json()) as { svg?: unknown };
  if (typeof body.svg !== "string" || !body.svg) fail();
  return body.svg;
}

async function pngBlobToDataUrl(blob: Blob): Promise<string> {
  if (
    blob.type !== "image/png" ||
    blob.size <= 0 ||
    blob.size > MAX_PRESENTATION_IMAGE_FILE_BYTES
  ) {
    fail();
  }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const src = `data:image/png;base64,${bytesToBase64(bytes)}`;
  if (!isValidInlineImageSource(src)) fail();
  // The preview is held inline: fitted to the page envelope's single-string
  // bound (lib/office-picture-envelope), else it would fail the deck's save.
  const fitted = await fitOfficePictureSource(src);
  if (fitted === null) fail();
  return fitted;
}

/** One local SVG as the passive picture a page keeps inline. */
export interface PassiveSvgPicture {
  /** `data:image/svg+xml;base64,` of the passive SVG (svg_sanitizer). */
  readonly src: string;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  readonly sizeBytes: number;
}

/**
 * Validate one local SVG and make it the passive picture an editor keeps in
 * its page (the server's svg_sanitizer when it carries anything active). The
 * Word and Excel writers draw its PNG compatibility picture on the server;
 * the deck pairs it with one drawn here
 * (`preparePresentationOfficeGraphicFile`).
 */
export async function preparePassiveSvgPicture(file: File): Promise<PassiveSvgPicture> {
  if (file.size <= 0) throw new PresentationImageFileError("empty");
  // The SVG markup itself is held inline (no raster to re-encode), so the
  // file must fit the page envelope's single-string bound as it is.
  if (
    file.size > MAX_PRESENTATION_IMAGE_FILE_BYTES ||
    file.size > OFFICE_PICTURE_FILE_MAX_BYTES
  ) {
    throw new PresentationImageFileError("too-large");
  }
  if (file.type.trim().toLowerCase() !== SVG_MIME) {
    throw new PresentationImageFileError("unsupported-type");
  }
  const bytes = await readFileBytes(file);
  if (bytes.byteLength !== file.size) fail();
  let markup: string;
  try {
    markup = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    fail();
  }
  let passiveBytes = bytes;
  if (!isSafeInlineSvgPayload(markup)) {
    // An SVG as drawing tools write it (styles, links, metadata, a script):
    // the server keeps its drawing and drops the rest, the same boundary a
    // stored SVG crosses (svg_sanitizer). Word and PowerPoint insert such a
    // file as a picture; only one with nothing to draw is refused.
    markup = await passiveSvgMarkup(bytes);
    if (!isSafeInlineSvgPayload(markup)) fail();
    passiveBytes = new TextEncoder().encode(markup);
    if (
      passiveBytes.byteLength > MAX_PRESENTATION_IMAGE_FILE_BYTES ||
      passiveBytes.byteLength > OFFICE_PICTURE_FILE_MAX_BYTES
    ) {
      throw new PresentationImageFileError("too-large");
    }
  }
  const { width: pixelWidth, height: pixelHeight } = svgDimensions(markup);
  const src = `data:${SVG_MIME};base64,${bytesToBase64(passiveBytes)}`;
  if (!isValidInlineImageSource(src)) fail();
  return Object.freeze({ src, pixelWidth, pixelHeight, sizeBytes: passiveBytes.byteLength });
}

/** Validate, keep and pair one local SVG with its Office compatibility preview. */
export async function preparePresentationOfficeGraphicFile(
  file: File,
): Promise<PreparedPresentationImageFile> {
  const { src, pixelWidth, pixelHeight, sizeBytes } = await preparePassiveSvgPicture(file);
  const fallback = await rasterDataUrlToPngBlob(src, pixelWidth, pixelHeight, {
    scale: 1,
    background: null,
  });
  if (!fallback) throw new PresentationImageFileError("cache-failed");
  return Object.freeze({
    src,
    mimeType: SVG_MIME,
    pixelWidth,
    pixelHeight,
    sizeBytes,
    hasAlpha: true,
    exifOrientation: 1,
    officeGraphicKind: "svg",
    graphicFallbackSrc: await pngBlobToDataUrl(fallback),
  });
}

/** The pixels a rasterized SVG is drawn at, per CSS pixel of its own size:
 * enough for a small icon enlarged on a slide, bounded by the picture budget. */
function officeGraphicRasterScale(width: number, height: number): number {
  const longest = Math.max(width, height);
  const byDimension = Math.floor(Math.min(2048, MAX_PRESENTATION_IMAGE_DIMENSION) / longest);
  const byPixels = Math.floor(Math.sqrt(MAX_PRESENTATION_IMAGE_PIXELS / (width * height)));
  return Math.max(1, Math.min(4, byDimension, byPixels));
}

/**
 * An SVG (a built-in icon, a local .svg) as the PNG picture a slide can
 * store. The upload channel takes no SVG and the page model states picture
 * custody for PNG / JPEG only, so the file itself could never be inserted
 * (PPT-B353): the same validated markup is drawn once, at up to four device
 * pixels per CSS pixel, and inserted as an ordinary picture at the SVG's own
 * size (`width` / `height`, CSS pixels). The vector identity is not kept.
 */
export async function rasterizePresentationOfficeGraphicFile(
  file: File,
): Promise<{ file: File; width: number; height: number }> {
  const prepared = await preparePresentationOfficeGraphicFile(file);
  const scale = officeGraphicRasterScale(prepared.pixelWidth, prepared.pixelHeight);
  const png = await rasterDataUrlToPngBlob(prepared.src, prepared.pixelWidth, prepared.pixelHeight, {
    scale,
    background: null,
  });
  if (!png || png.type !== "image/png" || png.size <= 0) {
    throw new PresentationImageFileError("cache-failed");
  }
  if (png.size > MAX_PRESENTATION_IMAGE_FILE_BYTES) {
    throw new PresentationImageFileError("too-large");
  }
  const base = file.name.replace(/\.svgz?$/i, "").trim() || "graphic";
  return {
    file: new File([png], `${base}.png`, { type: "image/png" }),
    width: prepared.pixelWidth,
    height: prepared.pixelHeight,
  };
}
