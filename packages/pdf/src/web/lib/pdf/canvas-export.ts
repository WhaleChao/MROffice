/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for "Export page as image" (PNG / JPG) in the PDF viewer.
 *
 * The component grabs the already-rendered pdf.js page <canvas>, calls
 * `canvas.toDataURL(mime, quality)` (same pattern as SignatureComposer), and
 * hands the resulting data URL here to turn into a downloadable Blob. Keeping
 * the data-URL → Blob conversion + filename + format mapping here (DOM-free)
 * makes the logic unit-testable without a real canvas.
 *
 * For the "export ALL pages / a page range as images" action (Acrobat/WPS), the
 * component renders EACH selected page to an offscreen canvas and downloads it;
 * the pure page-LIST math (current page vs all vs a range spec) lives here in
 * {@link imageExportPageList} so it stays unit-testable.
 */

import { parseRangeSpec } from "./range-spec";

/** The image formats the export action offers. */
export type ImageExportFormat = "png" | "jpg";

/** Which pages an image export covers. */
export type ImageExportScope = "current" | "all" | "range";

/**
 * Resolve the ORDERED list of 1-based page numbers an image export should
 * produce, given the scope, the document's `totalPages`, the `currentPage`
 * (used by the "current" scope) and an Acrobat-style `rangeSpec` (used by the
 * "range" scope, parsed/clamped/sorted/de-duped by the shared range parser):
 *
 *   - "current": just `[currentPage]` (clamped to `[1, totalPages]`).
 *   - "all":     `[1 .. totalPages]`.
 *   - "range":   the pages the spec selects (e.g. `"1-3, 5, 8-"`); `[]` when the
 *                spec resolves to nothing (empty / pure garbage).
 *
 * Returns `[]` when `totalPages < 1`. Pure — no DOM / pdf.js. Defends against a
 * runaway "all" by capping at the range parser's {@link MAX_SPLIT_PAGES} ceiling.
 */
export function imageExportPageList(
  scope: ImageExportScope,
  totalPages: number,
  currentPage: number,
  rangeSpec: string,
): number[] {
  const total =
    Number.isFinite(totalPages) && totalPages >= 1
      ? Math.min(Math.trunc(totalPages), 10_000)
      : 0;
  if (total < 1) return [];
  if (scope === "range") {
    return parseRangeSpec(
      typeof rangeSpec === "string" ? rangeSpec : "",
      total,
    );
  }
  if (scope === "all") {
    const out: number[] = [];
    for (let p = 1; p <= total; p += 1) out.push(p);
    return out;
  }
  // "current": clamp to a valid single page.
  const cur =
    Number.isFinite(currentPage) && currentPage >= 1
      ? Math.min(Math.trunc(currentPage), total)
      : 1;
  return [cur];
}







export const IMAGE_EXPORT_RESOLUTIONS = [72, 96, 150, 300, 600] as const;
export type ImageExportResolution = (typeof IMAGE_EXPORT_RESOLUTIONS)[number];
export const DEFAULT_IMAGE_EXPORT_RESOLUTION: ImageExportResolution = 150;

export function isImageExportResolution(
  value: unknown,
): value is ImageExportResolution {
  return (IMAGE_EXPORT_RESOLUTIONS as readonly unknown[]).includes(value);
}

export interface ImageExportRenderSize {
  /** pdf.js render scale (1 = 72 ppi on the page's size in points). */
  readonly scale: number;
  readonly width: number;
  readonly height: number;
  /** True when the pixel budget lowered the resolution for this page. */
  readonly capped: boolean;
}

/**
 * The render scale and pixel size of one page exported at `ppi`, for a page
 * `widthPoints` x `heightPoints` as displayed (rotation applied). A page too
 * large for `maxPixels` at that resolution is scaled down uniformly and
 * reported as capped.
 */
export function imageExportRenderSize(
  widthPoints: number,
  heightPoints: number,
  ppi: number,
  maxPixels: number,
): ImageExportRenderSize | null {
  if (
    !Number.isFinite(widthPoints) ||
    !Number.isFinite(heightPoints) ||
    widthPoints <= 0 ||
    heightPoints <= 0 ||
    !Number.isFinite(ppi) ||
    ppi <= 0
  ) {
    return null;
  }
  let scale = ppi / 72;
  let capped = false;
  const pixels = widthPoints * scale * heightPoints * scale;
  if (Number.isFinite(maxPixels) && maxPixels > 0 && pixels > maxPixels) {
    scale *= Math.sqrt(maxPixels / pixels);
    capped = true;
  }
  return {
    scale,
    width: Math.max(1, Math.floor(widthPoints * scale)),
    height: Math.max(1, Math.floor(heightPoints * scale)),
    capped,
  };
}

export const IMAGE_EXPORT_FORMATS: readonly ImageExportFormat[] = [
  "png",
  "jpg",
];

export const LONG_IMAGE_MAX_DIMENSION = 32_767;
export const LONG_IMAGE_MAX_PIXELS = 64 * 1024 * 1024;
export const LONG_IMAGE_PAGE_GAP = 12;

export interface LongImagePageSize {
  readonly pageNumber: number;
  readonly width: number;
  readonly height: number;
}

export interface LongImagePagePlacement extends LongImagePageSize {
  readonly x: number;
  readonly y: number;
}

export interface LongImageLayout {
  readonly width: number;
  readonly height: number;
  readonly scale: number;
  readonly placements: readonly LongImagePagePlacement[];
}

/**
 * Plan one bounded vertical image without allocating a canvas. Every page is
 * scaled uniformly and centered, preserving mixed page sizes and order. The
 * 5% floor prevents an apparently successful export whose text is unreadable.
 */
export function planLongImageLayout(
  pages: readonly LongImagePageSize[],
): LongImageLayout | null {
  if (pages.length < 1 || pages.length > 200) return null;
  const normalized = pages.map((page) => ({
    pageNumber: Math.trunc(page.pageNumber),
    width: Math.ceil(page.width),
    height: Math.ceil(page.height),
  }));
  if (
    normalized.some(
      (page) =>
        !Number.isSafeInteger(page.pageNumber) ||
        page.pageNumber < 1 ||
        !Number.isSafeInteger(page.width) ||
        !Number.isSafeInteger(page.height) ||
        page.width < 1 ||
        page.height < 1 ||
        page.width > LONG_IMAGE_MAX_DIMENSION ||
        page.height > LONG_IMAGE_MAX_DIMENSION,
    )
  ) {
    return null;
  }
  const widest = Math.max(...normalized.map((page) => page.width));
  const unscaledHeight =
    normalized.reduce((total, page) => total + page.height, 0) +
    LONG_IMAGE_PAGE_GAP * (normalized.length - 1);
  const scale = Math.min(
    1,
    LONG_IMAGE_MAX_DIMENSION / widest,
    LONG_IMAGE_MAX_DIMENSION / unscaledHeight,
    Math.sqrt(LONG_IMAGE_MAX_PIXELS / (widest * unscaledHeight)),
  );
  if (!Number.isFinite(scale) || scale < 0.05) return null;
  const width = Math.max(1, Math.floor(widest * scale));
  const gap = Math.max(1, Math.floor(LONG_IMAGE_PAGE_GAP * scale));
  let y = 0;
  const placements = normalized.map((page, index) => {
    const pageWidth = Math.max(1, Math.floor(page.width * scale));
    const pageHeight = Math.max(1, Math.floor(page.height * scale));
    const placement = Object.freeze({
      pageNumber: page.pageNumber,
      width: pageWidth,
      height: pageHeight,
      x: Math.floor((width - pageWidth) / 2),
      y,
    });
    y += pageHeight + (index + 1 < normalized.length ? gap : 0);
    return placement;
  });
  if (
    y < 1 ||
    y > LONG_IMAGE_MAX_DIMENSION ||
    width * y > LONG_IMAGE_MAX_PIXELS
  ) {
    return null;
  }
  return Object.freeze({
    width,
    height: y,
    scale,
    placements: Object.freeze(placements),
  });
}

/** Narrow an arbitrary value to a supported format (defaults to "png"). */
export function asImageExportFormat(value: unknown): ImageExportFormat {
  return value === "jpg" ? "jpg" : "png";
}

/**
 * The `canvas.toDataURL` MIME type for a format. JPG maps to `image/jpeg`
 * (the canonical canvas MIME); PNG maps to `image/png`.
 */
export function canvasMimeForFormat(format: ImageExportFormat): string {
  return format === "jpg" ? "image/jpeg" : "image/png";
}

/**
 * Quality argument passed to `toDataURL` for lossy formats. PNG ignores it
 * (returns undefined so the call omits the argument); JPG uses a high-quality
 * default so exported pages stay crisp.
 */
export function canvasQualityForFormat(
  format: ImageExportFormat,
): number | undefined {
  return format === "jpg" ? 0.92 : undefined;
}

/** Strip a trailing `.pdf` (case-insensitive) and collapse to a safe stem. */
function exportStem(baseName: string | null | undefined): string {
  const raw = (typeof baseName === "string" ? baseName : "").trim();
  const noExt = raw.replace(/\.pdf$/i, "").trim();
  // Replace path separators / control chars so the download name is safe.
  const safe = noExt.replace(/[\\/\u0000-\u001f]+/g, " ").trim();
  return safe || "page";
}

/**
 * Build the download filename for a page image, e.g. `report-p3.png`. `page`
 * is 1-based; a non-finite / non-positive page falls back to `p1`.
 */
export function imageExportFilename(
  baseName: string | null | undefined,
  page: number,
  format: ImageExportFormat,
): string {
  const stem = exportStem(baseName);
  const n = Number.isFinite(page) && page >= 1 ? Math.trunc(page) : 1;
  const ext = format === "jpg" ? "jpg" : "png";
  return `${stem}-p${n}.${ext}`;
}

export function longImageExportFilename(
  baseName: string | null | undefined,
  pages: readonly number[],
  format: ImageExportFormat,
): string {
  const stem = exportStem(baseName);
  const normalized = pages.filter(
    (page) => Number.isSafeInteger(page) && page >= 1,
  );
  const first = normalized[0] ?? 1;
  const last = normalized.at(-1) ?? first;
  const range = first === last ? `p${first}` : `p${first}-${last}`;
  return `${stem}-${range}-long.${format === "jpg" ? "jpg" : "png"}`;
}

/** Result of decoding a canvas data URL into raw bytes. */
export interface DecodedImageData {
  bytes: Uint8Array;
  mime: string;
}

/**
 * Decode a `data:<mime>;base64,<payload>` URL (what `canvas.toDataURL` emits)
 * into its raw bytes + MIME type. Returns null for anything that is not a
 * base64 image data URL we can safely turn into a Blob — including a non-image
 * MIME, a missing/empty payload, or invalid base64.
 *
 * Uses `atob` (available in the browser + jsdom) and is otherwise pure.
 */
export function decodeImageDataUrl(dataUrl: unknown): DecodedImageData | null {
  if (typeof dataUrl !== "string") return null;
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([\s\S]*)$/i.exec(dataUrl);
  if (!match) return null;
  const mime = match[1].toLowerCase();
  const payload = match[2].replace(/\s/g, "");
  if (!payload) return null;
  let binary: string;
  try {
    binary = atob(payload);
  } catch {
    return null;
  }
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i) & 0xff;
  }
  return { bytes, mime };
}







export const IMAGE_EXPORT_MAX_PAGES = 200;

/**
 * Split a page selection into page-range specs (`"1-200"`, `"201-205, 300"`)
 * of at most `max` pages each, in selection order, for exporting in batches.
 */
export function imageExportBatchRanges(
  pages: ReadonlyArray<number>,
  max: number = IMAGE_EXPORT_MAX_PAGES,
): string[] {
  const size = Number.isInteger(max) && max > 0 ? max : IMAGE_EXPORT_MAX_PAGES;
  const batches: string[] = [];
  for (let start = 0; start < pages.length; start += size) {
    const chunk = pages.slice(start, start + size);
    const runs: string[] = [];
    let runStart = chunk[0];
    let previous = chunk[0];
    for (let index = 1; index <= chunk.length; index += 1) {
      const page = chunk[index];
      if (page !== undefined && page === previous + 1) {
        previous = page;
        continue;
      }
      runs.push(runStart === previous ? `${runStart}` : `${runStart}-${previous}`);
      runStart = page;
      previous = page;
    }
    batches.push(runs.join(", "));
  }
  return batches;
}
