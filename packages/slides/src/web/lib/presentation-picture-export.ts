/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Saving slides as pictures (GAP #3): PowerPoint's File > Save As > PNG / JPEG
 * asks "Which slides do you want to export?" -- All Slides saves each slide as
 * a separate picture, Just This One the selected slide (Microsoft Support,
 * "Save a slide as an image or as a separate presentation file").
 *
 * Size: PowerPoint's default is 96 dots per inch -- a 13.333 x 7.5 in slide is
 * 1280 x 720 pixels, a 10 x 7.5 in one 960 x 720 -- and a picture holds at most
 * 100,000,000 pixels (Microsoft Learn, "How to export high-resolution (high-dpi)
 * slides from PowerPoint"). One slide and every slide are made at the same
 * size; a higher resolution is the user's choice, within that bound (and the
 * browser's canvas, `rasterElementPicture`).
 *
 * Every slide goes into one .zip, made as the slides are drawn (each picture
 * leaves memory once it is in the archive), with progress, cancel, and every
 * slide that could not be drawn named. The folder and file names are
 * PowerPoint's (truth export-all-pictures, PowerPoint 16.0 in English: a folder
 * named as the file, Slide1.PNG ... SlideN.PNG, .JPG for JPEG; 1280 x 720 for
 * 16:9 and 1040 x 720 for a 9,906,000 x 6,858,000 EMU slide, both at 96 dpi),
 * the file name in the interface's language.
 */

import { Zip, ZipPassThrough } from "fflate";

/** The resolutions offered, dots per inch; the first is PowerPoint's default. */
export const SLIDE_PICTURE_RESOLUTIONS = [96, 150, 200, 300] as const;
export type SlidePictureResolution = (typeof SLIDE_PICTURE_RESOLUTIONS)[number];
export const SLIDE_PICTURE_DEFAULT_DPI: SlidePictureResolution = 96;

/**
 * Whether a hidden slide is saved too. Microsoft's article does not say;
 * PowerPoint does (truth export-all-pictures: a 4-slide deck whose slide 3 is
 * hidden saves Slide1-Slide4).
 */
export const SLIDE_PICTURES_INCLUDE_HIDDEN = true;

/** PowerPoint's ceiling for one exported picture, pixels. */
export const SLIDE_PICTURE_MAX_PIXELS = 100_000_000;

/** CSS pixels are 96 to the inch: the slide's design size is its 96-dpi picture. */
const DESIGN_DPI = 96;

/**
 * The multiplier from a slide's design size (`widthPx` x `heightPx`, CSS
 * pixels) to its picture at `dpi`, held to {@link SLIDE_PICTURE_MAX_PIXELS}.
 */
export function slidePictureScale(dpi: number, widthPx: number, heightPx: number): number {
  const wanted = Number.isFinite(dpi) && dpi > 0 ? dpi / DESIGN_DPI : 1;
  const area = widthPx * heightPx;
  if (!(area > 0)) return wanted;
  return Math.min(wanted, Math.sqrt(SLIDE_PICTURE_MAX_PIXELS / area));
}

/** The pixel size of a slide's picture at `dpi`, as it is rastered. */
export function slidePictureSize(dpi: number, widthPx: number, heightPx: number): { width: number; height: number } {
  const scale = slidePictureScale(dpi, widthPx, heightPx);
  return { width: Math.max(1, Math.round(widthPx * scale)), height: Math.max(1, Math.round(heightPx * scale)) };
}

/** EMU to the inch (DrawingML). */
const EMU_PER_INCH = 914_400;
/** The standard 16:9 slide, for a slide that does not state its size. */
const DEFAULT_SLIDE_EMU = { width: 12_192_000, height: 6_858_000 } as const;

/**
 * The pixel size of a slide's picture at `dpi` from the slide's own size
 * (EMU), as PowerPoint sizes it: 96 dpi makes a 12,192,000 x 6,858,000 EMU
 * slide 1280 x 720.
 */
export function slidePicturePixels(
  dpi: number,
  widthEmu: number | null | undefined,
  heightEmu: number | null | undefined,
): { width: number; height: number } {
  const width = widthEmu && widthEmu > 0 ? widthEmu : DEFAULT_SLIDE_EMU.width;
  const height = heightEmu && heightEmu > 0 ? heightEmu : DEFAULT_SLIDE_EMU.height;
  return slidePictureSize(dpi, (width / EMU_PER_INCH) * DESIGN_DPI, (height / EMU_PER_INCH) * DESIGN_DPI);
}

/** A name safe as a folder or file name on every system (no path separators,
 * reserved characters or trailing dots and spaces). */
export function safePictureName(name: string, fallback: string): string {
  const cleaned = name
    .replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/, "");
  return cleaned || fallback;
}

export interface SlidePicturesArchive {
  /** The archive's parts, in order (`new Blob(parts, { type: "application/zip" })`). */
  readonly parts: readonly Blob[];
  /** The slides (0-based, of those asked for) whose picture could not be made. */
  readonly failed: readonly number[];
  /** How many pictures are in the archive. */
  readonly written: number;
}

export class SlidePicturesCancelled extends Error {
  constructor() {
    super("cancelled");
    this.name = "SlidePicturesCancelled";
  }
}

/**
 * Draws `count` pictures one at a time (`draw(i)`), each written into a .zip
 * under `folder/` as `fileName(i)` as soon as it is made: no picture waits in
 * memory for the others. A picture that fails is left out and named in
 * `failed`; an abort stops before the next picture (`SlidePicturesCancelled`).
 * Pictures are stored, not compressed again (PNG and JPEG already are).
 */
export async function zipSlidePictures(options: {
  readonly count: number;
  readonly folder: string;
  readonly fileName: (index: number) => string;
  readonly draw: (index: number) => Promise<Blob>;
  readonly onProgress?: (done: number, count: number) => void;
  readonly signal?: AbortSignal;
}): Promise<SlidePicturesArchive> {
  const { count, folder, fileName, draw, onProgress, signal } = options;
  const parts: Blob[] = [];
  const failed: number[] = [];
  let written = 0;
  let zipError: unknown = null;
  let finished!: () => void;
  const done = new Promise<void>((resolve) => {
    finished = resolve;
  });
  const zip = new Zip((error, chunk, final) => {
    if (error) {
      zipError = error;
      finished();
      return;
    }
    // Each piece becomes its own Blob, which the browser may keep out of the
    // page's memory; nothing here holds the pictures themselves.
    if (chunk.length) parts.push(new Blob([chunk as Uint8Array<ArrayBuffer>]));
    if (final) finished();
  });
  for (let index = 0; index < count; index += 1) {
    if (signal?.aborted) {
      zip.terminate();
      throw new SlidePicturesCancelled();
    }
    let picture: Blob | null = null;
    try {
      picture = await draw(index);
    } catch {
      picture = null;
    }
    if (signal?.aborted) {
      zip.terminate();
      throw new SlidePicturesCancelled();
    }
    if (picture) {
      const entry = new ZipPassThrough(`${folder}/${fileName(index)}`);
      zip.add(entry);
      entry.push(new Uint8Array(await picture.arrayBuffer()), true);
      written += 1;
    } else {
      failed.push(index);
    }
    onProgress?.(index + 1, count);
  }
  zip.end();
  await done;
  if (zipError) throw zipError;
  return { parts, failed, written };
}
