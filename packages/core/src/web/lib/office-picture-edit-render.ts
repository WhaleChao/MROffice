/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser half of the Office picture edits that change a picture's pixels:
 * Remove Background (Keep Changes) and Edit Picture ▸ Erase. The pure math
 * lives in `office-picture-background-removal`; this module decodes the
 * picture, runs it at a bounded working size, and encodes the result as a
 * PNG data URL (PNG: the removed background is transparency).
 */

import {
  applyBackgroundRemovalMask,
  computeBackgroundRemovalMask,
  erasePictureRegion,
  scaleMask,
  type BackgroundRemovalSegment,
  type BackgroundRemovalSpec,
  type RgbaPixels,
} from "./office-picture-background-removal";

/** The longest side the estimate is computed at (the mask is scaled up). */
export const BACKGROUND_REMOVAL_WORKING_SIZE = 640;
/** The longest side a result keeps (a larger picture is reduced to this). */
export const PICTURE_EDIT_MAX_SIZE = 4096;

export async function loadPictureImage(src: string): Promise<HTMLImageElement> {
  const image = new Image();
  image.decoding = "async";
  if (!src.startsWith("data:") && !src.startsWith("blob:")) image.crossOrigin = "anonymous";
  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error("picture could not be decoded"));
    image.src = src;
  });
  return image;
}

/** The picture's pixels at most `maxSide` on its longest side. */
export function picturePixels(image: HTMLImageElement, maxSide: number): RgbaPixels {
  const natural = Math.max(1, image.naturalWidth || image.width);
  const naturalHeight = Math.max(1, image.naturalHeight || image.height);
  const scale = Math.min(1, maxSide / Math.max(natural, naturalHeight));
  const width = Math.max(1, Math.round(natural * scale));
  const height = Math.max(1, Math.round(naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) throw new Error("no 2d canvas");
  context.drawImage(image, 0, 0, width, height);
  const data = context.getImageData(0, 0, width, height).data;
  return { width, height, data };
}

function encodePng(pixels: RgbaPixels): string {
  const canvas = document.createElement("canvas");
  canvas.width = pixels.width;
  canvas.height = pixels.height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("no 2d canvas");
  const image = context.createImageData(pixels.width, pixels.height);
  image.data.set(pixels.data);
  context.putImageData(image, 0, 0);
  return canvas.toDataURL("image/png");
}

/** Keep Changes: the picture with its background removed, as a PNG data URL. */
export async function renderBackgroundRemoval(
  src: string,
  spec: BackgroundRemovalSpec,
): Promise<string> {
  const image = await loadPictureImage(src);
  const working = picturePixels(image, BACKGROUND_REMOVAL_WORKING_SIZE);
  const mask = computeBackgroundRemovalMask(working, spec);
  const full = picturePixels(image, PICTURE_EDIT_MAX_SIZE);
  const fullMask = scaleMask(mask, working.width, working.height, full.width, full.height);
  return encodePng({
    width: full.width,
    height: full.height,
    data: applyBackgroundRemovalMask(full, fullMask),
  });
}

/**
 * Erase: the brushed strokes filled from their surroundings, worked on the
 * strokes' bounding box only so a large picture stays quick.
 */
export async function renderPictureErase(
  src: string,
  segments: readonly BackgroundRemovalSegment[],
  radiusFraction: number,
): Promise<string> {
  const image = await loadPictureImage(src);
  const full = picturePixels(image, PICTURE_EDIT_MAX_SIZE);
  if (segments.length === 0) return encodePng(full);
  const radius = Math.max(1, Math.round(Math.min(full.width, full.height) * radiusFraction));
  const xs = segments.flatMap((s) => [s.x1, s.x2]).map((x) => x * (full.width - 1));
  const ys = segments.flatMap((s) => [s.y1, s.y2]).map((y) => y * (full.height - 1));
  const pad = radius * 3;
  const left = Math.max(0, Math.floor(Math.min(...xs) - pad));
  const top = Math.max(0, Math.floor(Math.min(...ys) - pad));
  const right = Math.min(full.width, Math.ceil(Math.max(...xs) + pad));
  const bottom = Math.min(full.height, Math.ceil(Math.max(...ys) + pad));
  const width = Math.max(1, right - left);
  const height = Math.max(1, bottom - top);
  const crop = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    const from = ((top + y) * full.width + left) * 4;
    crop.set(full.data.subarray(from, from + width * 4), y * width * 4);
  }
  const local = segments.map((s) => ({
    x1: (s.x1 * (full.width - 1) - left) / Math.max(1, width - 1),
    y1: (s.y1 * (full.height - 1) - top) / Math.max(1, height - 1),
    x2: (s.x2 * (full.width - 1) - left) / Math.max(1, width - 1),
    y2: (s.y2 * (full.height - 1) - top) / Math.max(1, height - 1),
  }));
  const filled = erasePictureRegion(
    { width, height, data: crop },
    local,
    radius / Math.max(1, Math.min(width, height)),
  );
  const out = new Uint8ClampedArray(full.data);
  for (let y = 0; y < height; y += 1) {
    out.set(filled.subarray(y * width * 4, (y + 1) * width * 4), ((top + y) * full.width + left) * 4);
  }
  return encodePng({ width: full.width, height: full.height, data: out });
}
