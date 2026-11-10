/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A photo's drawn region on the page, in fractions of the page as shown. */
export interface PdfReadModePhotoRegion {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** The drawn parallelogram (four corners, fractions of the region) when
   * it is not the region itself: a turned or sheared image. */
  readonly polygon: readonly (readonly [number, number])[] | null;
}

const MAX_REGIONS = 400;
/** A region smaller than this (fraction of the page's side) is a rule or a
 * dot drawn as an image: copying it is not worth a canvas. */
const MIN_SIDE = 0.004;

const turnQuarter = ([x, y]: readonly [number, number]): [number, number] => [1 - y, x];

/**
 * The regions of the images pdf.js recorded (`imageCoordinates`, six numbers
 * an image) for a page drawn at `recordedRotation`, shown now at `rotation`
 * (both 0/90/180/270). Out-of-page parts are cut away; tiny ones are left
 * to the filter.
 */
export function pdfReadModePhotoRegions(
  coordinates: ArrayLike<number> | null | undefined,
  recordedRotation: number,
  rotation: number,
): PdfReadModePhotoRegion[] {
  if (!coordinates || coordinates.length < 6) return [];
  const turns = ((((rotation - recordedRotation) / 90) % 4) + 4) % 4;
  if (!Number.isInteger(turns)) return [];
  const regions: PdfReadModePhotoRegion[] = [];
  for (let at = 0; at + 6 <= coordinates.length && regions.length < MAX_REGIONS; at += 6) {
    const origin: [number, number] = [coordinates[at], coordinates[at + 1]];
    const up: [number, number] = [coordinates[at + 2], coordinates[at + 3]];
    const along: [number, number] = [coordinates[at + 4], coordinates[at + 5]];
    let corners: [number, number][] = [origin, up, [up[0] + along[0] - origin[0], up[1] + along[1] - origin[1]], along];
    if (!corners.every(([x, y]) => Number.isFinite(x) && Number.isFinite(y))) continue;
    for (let turn = 0; turn < turns; turn += 1) corners = corners.map(turnQuarter);
    const xs = corners.map(([x]) => x), ys = corners.map(([, y]) => y);
    const left = Math.max(0, Math.min(...xs)), right = Math.min(1, Math.max(...xs));
    const top = Math.max(0, Math.min(...ys)), bottom = Math.min(1, Math.max(...ys));
    const width = right - left, height = bottom - top;
    if (!(width >= MIN_SIDE && height >= MIN_SIDE)) continue;
    // Upright (each edge along an axis): the region is the image.
    const upright = corners.every(([x, y], index) => {
      const [nx, ny] = corners[(index + 1) % 4];
      return Math.abs(nx - x) < 1e-4 || Math.abs(ny - y) < 1e-4;
    });
    regions.push({
      x: left, y: top, width, height,
      polygon: upright ? null : corners.map(([x, y]) => [(x - left) / width, (y - top) / height] as const),
    });
  }
  return regions;
}

// pdf.js operator codes (pdfjs-dist OPS, 5.x).
const OP = {
  save: 10, restore: 11, clip: 29, eoClip: 30, setTextRenderingMode: 38,
  shadingFill: 62, beginGroup: 76, endGroup: 77, constructPath: 91, rawFillPath: 94,
} as const;
/** A path's paint (constructPath's first argument); endPath only clips. */
const PATH_PAINTS = new Set([20, 21, 22, 23, 24, 25, 26, 27]);
const TEXT_SHOWS = new Set([44, 45, 46, 47]);
/** Images drawn in their own colours (an image mask is a one-colour
 * stencil, recoloured with the text). */
const IMAGES = new Set([85, 86, 87, 88]);
const OTHER_PAINTS = new Set([OP.shadingFill, 83, 84, 89, 90, OP.rawFillPath]);

/** Which operations of a page's operator list each extra drawing runs. */
export interface PdfReadModeLayerFilters {
  /** The images alone: every image, no other paint. */
  readonly images: (index: number) => boolean;
  /** What the page paints after its first image, no image. */
  readonly over: (index: number) => boolean;
}

/**
 * The two drawings' operation filters for a page's operator list (`fnArray`,
 * `argsArray` of pdf.js's getOperatorList, the same intent and annotation
 * mode as the page's own drawing), or null when it draws no image. Only a
 * paint is ever left out: the state, the transforms and the clips run in
 * both, and so does a paint that sets a clip (a path painted and clipped at
 * once, text in a clipping render mode) and everything inside a soft mask's
 * group (the mask is made of them).
 */
export function pdfReadModeLayerFilters(
  fnArray: ArrayLike<number>,
  argsArray: ArrayLike<unknown>,
): PdfReadModeLayerFilters | null {
  // 0: state (both run it); 1: an image; 2: another paint.
  const kinds = new Uint8Array(fnArray.length);
  let firstImage = -1;
  let pendingClip = false;
  let textMode = 0;
  const textModes: number[] = [];
  const groups: boolean[] = [];
  for (let index = 0; index < fnArray.length; index += 1) {
    const fn = fnArray[index];
    const args = argsArray[index] as unknown[] | null | undefined;
    const inMask = groups.includes(true);
    if (fn === OP.save) textModes.push(textMode);
    else if (fn === OP.restore) textMode = textModes.pop() ?? textMode;
    else if (fn === OP.setTextRenderingMode) textMode = Number(args?.[0]) || 0;
    else if (fn === OP.clip || fn === OP.eoClip) pendingClip = true;
    else if (fn === OP.beginGroup) groups.push(Boolean((args?.[0] as { smask?: unknown } | undefined)?.smask));
    else if (fn === OP.endGroup) groups.pop();
    if (inMask) continue;
    if (IMAGES.has(fn)) {
      kinds[index] = 1;
      if (firstImage < 0) firstImage = index;
    } else if (fn === OP.constructPath) {
      if (PATH_PAINTS.has(Number(args?.[0])) && !pendingClip) kinds[index] = 2;
      pendingClip = false;
    } else if (TEXT_SHOWS.has(fn)) {
      if (textMode < 4) kinds[index] = 2;
    } else if (OTHER_PAINTS.has(fn)) {
      kinds[index] = 2;
    }
  }
  if (firstImage < 0) return null;
  return {
    images: (index) => kinds[index] !== 2,
    over: (index) => kinds[index] === 0 || (kinds[index] === 2 && index > firstImage),
  };
}

/** The page drawn again for the photos: the images alone, and what is drawn
 * over them (both the page canvas's size, transparent elsewhere). */
export interface PdfReadModePhotoDrawings {
  readonly images: HTMLCanvasElement;
  readonly over: HTMLCanvasElement;
  /** The dark mode's filter, for what is drawn over the photos. */
  readonly overFilter: string;
}

/**
 * Lays the regions over the page, each from the two drawings (the images in
 * their colours, under what the page draws over them in the dark mode's),
 * in `layer` (an element over the page canvas, under the page's overlays).
 * An empty list clears the layer.
 */
export function paintPdfReadModePhotos(
  layer: HTMLElement,
  drawings: PdfReadModePhotoDrawings,
  regions: readonly PdfReadModePhotoRegion[],
): void {
  const document = layer.ownerDocument;
  const pieces: HTMLElement[] = [];
  const sourceWidth = drawings.images.width, sourceHeight = drawings.images.height;
  const crop = (source: HTMLCanvasElement, sx: number, sy: number, sw: number, sh: number): HTMLCanvasElement | null => {
    const canvas = document.createElement("canvas");
    canvas.width = sw;
    canvas.height = sh;
    const context = canvas.getContext("2d");
    if (!context) return null;
    context.drawImage(source, sx, sy, sw, sh, 0, 0, sw, sh);
    canvas.dataset.pdfReadModePhoto = "true";
    return canvas;
  };
  if (sourceWidth > 0 && sourceHeight > 0 && drawings.over.width === sourceWidth && drawings.over.height === sourceHeight) {
    for (const region of regions) {
      const sx = Math.floor(region.x * sourceWidth), sy = Math.floor(region.y * sourceHeight);
      const sw = Math.min(sourceWidth - sx, Math.ceil((region.x + region.width) * sourceWidth) - sx);
      const sh = Math.min(sourceHeight - sy, Math.ceil((region.y + region.height) * sourceHeight) - sy);
      if (sw < 1 || sh < 1) continue;
      const photo = crop(drawings.images, sx, sy, sw, sh);
      const over = crop(drawings.over, sx, sy, sw, sh);
      if (!photo || !over) continue;
      if (drawings.overFilter !== "none") over.style.filter = drawings.overFilter;
      const piece = document.createElement("span");
      piece.dataset.pdfReadModePhotoRegion = "true";
      piece.style.left = `${(sx / sourceWidth) * 100}%`;
      piece.style.top = `${(sy / sourceHeight) * 100}%`;
      piece.style.width = `${(sw / sourceWidth) * 100}%`;
      piece.style.height = `${(sh / sourceHeight) * 100}%`;
      if (region.polygon) {
        piece.style.clipPath = `polygon(${region.polygon.map(([x, y]) => `${(x * 100).toFixed(3)}% ${(y * 100).toFixed(3)}%`).join(", ")})`;
      }
      piece.append(photo, over);
      pieces.push(piece);
    }
  }
  clearPdfReadModePhotos(layer);
  layer.append(...pieces);
}

/** Takes the photos' copies off the page (their backing stores freed). */
export function clearPdfReadModePhotos(layer: HTMLElement): void {
  for (const old of layer.querySelectorAll<HTMLCanvasElement>("canvas")) {
    old.width = 0;
    old.height = 0;
  }
  layer.replaceChildren();
}
