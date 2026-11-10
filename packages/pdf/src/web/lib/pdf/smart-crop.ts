/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  cropBoxFromMargins,
  marginsForCropBox,
  type CropMargins,
} from "./crop-box";

export interface PdfSmartCropRaster {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8ClampedArray;
}

/** A page region in page percent (top-left origin). */
export interface PdfSmartCropKeepBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type PdfSmartCropResult = (
  | Readonly<{ status: "detected"; margins: CropMargins }>
  | Readonly<{ status: "no-uniform-margin" | "unavailable" }>
) & {
  /** Visible content the page canvas does not hold (a note icon the reader
   * draws itself): what the crop must keep and the preview must show. */
  readonly keepBoxes?: readonly PdfSmartCropKeepBox[];
};






export function keepBoxesInsideSmartCrop(
  result: PdfSmartCropResult,
  boxes: readonly PdfSmartCropKeepBox[],
): PdfSmartCropResult {
  const valid = boxes.filter(
    (box) =>
      [box.x, box.y, box.width, box.height].every(Number.isFinite) &&
      box.width > 0 &&
      box.height > 0,
  );
  if (valid.length === 0) return result;
  const withBoxes = { ...result, keepBoxes: valid };
  if (result.status !== "detected") return Object.freeze(withBoxes);
  const padding = 1;
  const left = Math.min(...valid.map((box) => box.x)) - padding;
  const top = Math.min(...valid.map((box) => box.y)) - padding;
  const right = 100 - Math.max(...valid.map((box) => box.x + box.width)) - padding;
  const bottom = 100 - Math.max(...valid.map((box) => box.y + box.height)) - padding;
  const margins = {
    left: roundedMargin(Math.max(0, Math.min(result.margins.left, left))),
    top: roundedMargin(Math.max(0, Math.min(result.margins.top, top))),
    right: roundedMargin(Math.max(0, Math.min(result.margins.right, right))),
    bottom: roundedMargin(Math.max(0, Math.min(result.margins.bottom, bottom))),
  };
  if (Math.max(...Object.values(margins)) < 0.5) {
    return Object.freeze({ status: "no-uniform-margin", keepBoxes: valid });
  }
  return Object.freeze({ status: "detected", margins: Object.freeze(margins), keepBoxes: valid });
}

const MAX_SAMPLE_DIMENSION = 512;
const BACKGROUND_DISTANCE = 18;
const MIN_BACKGROUND_LUMINANCE = 235;
const MAX_BACKGROUND_CHANNEL_SPREAD = 24;
const MIN_BORDER_CONSENSUS = 0.9;

function median(values: readonly number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}

function pixelOffset(width: number, x: number, y: number): number {
  return (y * width + x) * 4;
}

function colorDistance(
  data: Uint8ClampedArray,
  offset: number,
  background: readonly [number, number, number],
): number {
  return Math.max(
    Math.abs((data[offset] ?? 0) - background[0]),
    Math.abs((data[offset + 1] ?? 0) - background[1]),
    Math.abs((data[offset + 2] ?? 0) - background[2]),
  );
}

function roundedMargin(value: number): number {
  return Math.round(Math.max(0, Math.min(49, value)) * 10) / 10;
}

/**
 * Detect one uniform, paper-coloured border around rendered PDF pixels.
 *
 * This is intentionally conservative: coloured/full-bleed or noisy borders
 * return `no-uniform-margin`, and a detected rectangle is only a draft for the
 * crop dialog. Applying still goes through the existing exact PDF receipt.
 */
export function detectPdfSmartCropMargins(
  raster: PdfSmartCropRaster,
): PdfSmartCropResult {
  const { width, height, data } = raster;
  if (
    !Number.isSafeInteger(width) ||
    !Number.isSafeInteger(height) ||
    width < 8 ||
    height < 8 ||
    data.length !== width * height * 4
  ) {
    return Object.freeze({ status: "unavailable" });
  }

  const perimeter: Array<readonly [number, number]> = [];
  const horizontalStep = Math.max(1, Math.floor(width / 256));
  const verticalStep = Math.max(1, Math.floor(height / 256));
  for (let x = 0; x < width; x += horizontalStep) {
    perimeter.push([x, 0], [x, height - 1]);
  }
  for (let y = 1; y < height - 1; y += verticalStep) {
    perimeter.push([0, y], [width - 1, y]);
  }
  const channels = [[], [], []] as [number[], number[], number[]];
  for (const [x, y] of perimeter) {
    const offset = pixelOffset(width, x, y);
    if ((data[offset + 3] ?? 0) < 240) {
      return Object.freeze({ status: "no-uniform-margin" });
    }
    channels[0].push(data[offset] ?? 0);
    channels[1].push(data[offset + 1] ?? 0);
    channels[2].push(data[offset + 2] ?? 0);
  }
  const background = Object.freeze([
    median(channels[0]),
    median(channels[1]),
    median(channels[2]),
  ]) as readonly [number, number, number];
  const luminance =
    background[0] * 0.2126 +
    background[1] * 0.7152 +
    background[2] * 0.0722;
  const channelSpread =
    Math.max(...background) - Math.min(...background);
  const consensus =
    perimeter.filter(([x, y]) =>
      colorDistance(data, pixelOffset(width, x, y), background) <=
      BACKGROUND_DISTANCE,
    ).length / perimeter.length;
  if (
    luminance < MIN_BACKGROUND_LUMINANCE ||
    channelSpread > MAX_BACKGROUND_CHANNEL_SPREAD ||
    consensus < MIN_BORDER_CONSENSUS
  ) {
    return Object.freeze({ status: "no-uniform-margin" });
  }

  const rowCounts = new Uint32Array(height);
  const columnCounts = new Uint32Array(width);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = pixelOffset(width, x, y);
      if (
        (data[offset + 3] ?? 0) >= 16 &&
        colorDistance(data, offset, background) > BACKGROUND_DISTANCE
      ) {
        rowCounts[y] += 1;
        columnCounts[x] += 1;
      }
    }
  }
  const minimumRowPixels = Math.max(2, Math.round(width * 0.002));
  const minimumColumnPixels = Math.max(2, Math.round(height * 0.002));
  const firstRow = rowCounts.findIndex((count) => count >= minimumRowPixels);
  const firstColumn = columnCounts.findIndex(
    (count) => count >= minimumColumnPixels,
  );
  let lastRow = -1;
  let lastColumn = -1;
  for (let y = height - 1; y >= 0; y -= 1) {
    if (rowCounts[y] >= minimumRowPixels) {
      lastRow = y;
      break;
    }
  }
  for (let x = width - 1; x >= 0; x -= 1) {
    if (columnCounts[x] >= minimumColumnPixels) {
      lastColumn = x;
      break;
    }
  }
  if (
    firstRow < 0 ||
    firstColumn < 0 ||
    lastRow < firstRow ||
    lastColumn < firstColumn
  ) {
    return Object.freeze({ status: "no-uniform-margin" });
  }

  const paddingX = Math.max(2, Math.round(width * 0.01));
  const paddingY = Math.max(2, Math.round(height * 0.01));
  const left = Math.max(0, firstColumn - paddingX);
  const top = Math.max(0, firstRow - paddingY);
  const right = Math.min(width - 1, lastColumn + paddingX);
  const bottom = Math.min(height - 1, lastRow + paddingY);
  const margins = {
    left: roundedMargin((left / width) * 100),
    top: roundedMargin((top / height) * 100),
    right: roundedMargin(((width - 1 - right) / width) * 100),
    bottom: roundedMargin(((height - 1 - bottom) / height) * 100),
  };
  if (Math.max(...Object.values(margins)) < 0.5) {
    return Object.freeze({ status: "no-uniform-margin" });
  }
  // The keep-rect round trip reintroduces binary float noise (7.3 comes back
  // as 7.299999999999997); the detected margins are one-decimal values and
  
  const settled = marginsForCropBox(cropBoxFromMargins(margins));
  return Object.freeze({
    status: "detected",
    margins: Object.freeze({
      left: roundedMargin(settled.left),
      top: roundedMargin(settled.top),
      right: roundedMargin(settled.right),
      bottom: roundedMargin(settled.bottom),
    }),
  });
}

/** Downsample the rendered page before reading pixels; never mutate it. */
export function detectPdfSmartCropMarginsFromCanvas(
  source: HTMLCanvasElement | null | undefined,
  options: { readonly createCanvas?: () => HTMLCanvasElement } = {},
): PdfSmartCropResult {
  if (!source || source.width < 1 || source.height < 1) {
    return Object.freeze({ status: "unavailable" });
  }
  const scale = Math.min(
    1,
    MAX_SAMPLE_DIMENSION / Math.max(source.width, source.height),
  );
  const width = Math.max(1, Math.round(source.width * scale));
  const height = Math.max(1, Math.round(source.height * scale));
  const target =
    options.createCanvas?.() ?? document.createElement("canvas");
  target.width = width;
  target.height = height;
  const context = target.getContext("2d", { willReadFrequently: true });
  if (!context) return Object.freeze({ status: "unavailable" });
  try {
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    context.drawImage(source, 0, 0, width, height);
    return detectPdfSmartCropMargins(
      Object.freeze({
        width,
        height,
        data: context.getImageData(0, 0, width, height).data,
      }),
    );
  } catch {
    return Object.freeze({ status: "unavailable" });
  } finally {
    target.width = 0;
    target.height = 0;
  }
}
