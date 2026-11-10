/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pictures in a worksheet's header or footer (Excel's Page Setup > Header/
 * Footer > Custom Header > Insert Picture / Format Picture; `&G`, shown as
 * `&[Picture]`).
 *
 * Pure geometry and colour helpers shared by every place that draws one (the
 * Page Setup preview, Page Layout view) and by the Format Picture dialog.
 * The PDF renderer on the server follows the same rules
 * (apps/api/app/services/file_adapters/xlsx/header_footer_pictures.py).
 *
 * Semantics (Microsoft Learn, VBA `Graphic` object; ECMA-376 VML
 * `v:imagedata`):
 * - Excel inserts a picture at its original size; Height and Width are the
 *   printed size in points.
 * - Crop is measured against the ORIGINAL picture ("Cropping is calculated
 *   relative to the original size of the picture"); the model keeps each
 *   side as the share of the original it removes, as VML's cropleft/... do.
 * - Image control: Color Automatic / Grayscale / Black and White / Washout,
 *   Brightness and Contrast 0..1 with 0.5 as "no change". VML writes them as
 *   `grayscale`, `bilevel`, `blacklevel` (brightness − 0.5) and `gain`
 *   (contrast); Washout is gain 0.3 with blacklevel 0.35.
 */

import { PX_PER_PT } from "@mrnote/core/web/lib/office-units.ts";
import type { PrintHeaderFooterPicture } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { MAX_HEADER_FOOTER_PICTURE_POINTS } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

/** Points per CSS pixel at the 96 DPI Excel assumes for a picture's pixels. */
export const HEADER_FOOTER_PICTURE_POINTS_PER_PIXEL = 1 / PX_PER_PT;

/** Washout as VML writes it (gain 19661f, blacklevel 22938f). */
export const HEADER_FOOTER_WASHOUT = Object.freeze({ gain: 0.3, blacklevel: 0.35 });

export interface HeaderFooterPictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** The four crop shares of a picture, 0 for the sides it does not crop. */
export function headerFooterPictureCrop(
  picture: Pick<PrintHeaderFooterPicture, "crop">,
): HeaderFooterPictureCrop {
  return {
    left: picture.crop?.left ?? 0,
    top: picture.crop?.top ?? 0,
    right: picture.crop?.right ?? 0,
    bottom: picture.crop?.bottom ?? 0,
  };
}

/**
 * The picture Insert Picture puts in a section: the original size, scaled
 * down (aspect kept) only past the largest size a page can carry.
 */
export function insertedHeaderFooterPicture(
  src: string,
  natural: { width: number; height: number } | null,
): PrintHeaderFooterPicture {
  const width = Math.max(1, natural?.width ?? 96) * HEADER_FOOTER_PICTURE_POINTS_PER_PIXEL;
  const height = Math.max(1, natural?.height ?? 96) * HEADER_FOOTER_PICTURE_POINTS_PER_PIXEL;
  const scale = Math.min(
    1,
    MAX_HEADER_FOOTER_PICTURE_POINTS / width,
    MAX_HEADER_FOOTER_PICTURE_POINTS / height,
  );
  const round = (value: number) => Math.max(0.01, Math.round(value * scale * 100) / 100);
  return { src, width: round(width), height: round(height) };
}

/**
 * How a picture is drawn in a box `width`×`height` points scaled by
 * `pointToPx`: the whole original image sized so the kept part fills the
 * box, shifted by the cropped share (the box clips the rest).
 */
export function headerFooterPictureLayout(
  picture: Pick<PrintHeaderFooterPicture, "width" | "height" | "crop">,
  pointToPx: (points: number) => number,
): {
  boxWidth: number;
  boxHeight: number;
  imageWidth: number;
  imageHeight: number;
  imageLeft: number;
  imageTop: number;
} {
  const crop = headerFooterPictureCrop(picture);
  const keptWidth = Math.max(0.0001, 1 - crop.left - crop.right);
  const keptHeight = Math.max(0.0001, 1 - crop.top - crop.bottom);
  const boxWidth = pointToPx(picture.width);
  const boxHeight = pointToPx(picture.height);
  const imageWidth = boxWidth / keptWidth;
  const imageHeight = boxHeight / keptHeight;
  return {
    boxWidth,
    boxHeight,
    imageWidth,
    imageHeight,
    imageLeft: -crop.left * imageWidth,
    imageTop: -crop.top * imageHeight,
  };
}

/** VML `gain` for a contrast of 0..1 (0.5 = 1, the identity). */
export function headerFooterContrastGain(contrast: number): number {
  const value = Math.max(0, Math.min(1, contrast));
  if (value <= 0.5) return value * 2;
  // Approaching 1 the gain grows without bound; Office caps the stored value.
  return Math.min(0x7fff, 0.5 / Math.max(1e-4, 1 - value));
}

/** Contrast 0..1 for a VML `gain` (inverse of {@link headerFooterContrastGain}). */
export function headerFooterContrastFromGain(gain: number): number {
  if (!Number.isFinite(gain) || gain <= 0) return 0;
  if (gain <= 1) return gain / 2;
  return 1 - 0.5 / gain;
}

/**
 * The colour transfer the picture's Image Control applies, per channel
 * value v in 0..1: grey first when `grayscale`, then
 * `out = slope * v + intercept`, then a 50% threshold when `threshold`
 * (Black and White). Null when the picture is drawn as stored.
 */
export function headerFooterPictureTransfer(
  picture: Pick<PrintHeaderFooterPicture, "colorType" | "brightness" | "contrast">,
): { grayscale: boolean; threshold: boolean; slope: number; intercept: number } | null {
  let gain = headerFooterContrastGain(picture.contrast ?? 0.5);
  let blacklevel = (picture.brightness ?? 0.5) - 0.5;
  if (picture.colorType === "washout") {
    gain = HEADER_FOOTER_WASHOUT.gain;
    blacklevel = HEADER_FOOTER_WASHOUT.blacklevel;
  }
  const grayscale = picture.colorType === "grayscale" || picture.colorType === "blackAndWhite";
  const threshold = picture.colorType === "blackAndWhite";
  const identity = Math.abs(gain - 1) < 1e-6 && Math.abs(blacklevel) < 1e-6;
  if (identity && !grayscale) return null;
  return {
    grayscale,
    threshold,
    slope: gain,
    intercept: 0.5 - 0.5 * gain + blacklevel,
  };
}

/** The original (uncropped) size of a picture in points, from its displayed size. */
export function headerFooterPictureOriginalFromDisplayed(
  picture: Pick<PrintHeaderFooterPicture, "width" | "height" | "crop">,
  scale: { width: number; height: number },
): { width: number; height: number } {
  const crop = headerFooterPictureCrop(picture);
  return {
    width: picture.width / Math.max(1e-6, scale.width) / Math.max(1e-6, 1 - crop.left - crop.right),
    height: picture.height / Math.max(1e-6, scale.height) / Math.max(1e-6, 1 - crop.top - crop.bottom),
  };
}

/**
 * Format Picture's state for one picture: the original size (points, from
 * the picture's pixels at 96 DPI), the crop in points of that original, and
 * the scale of the displayed size against the kept (cropped) original.
 */
export interface HeaderFooterPictureFormat {
  width: number;
  height: number;
  cropLeft: number;
  cropTop: number;
  cropRight: number;
  cropBottom: number;
  lockAspectRatio: boolean;
  colorType: "automatic" | "grayscale" | "blackAndWhite" | "washout";
  brightness: number;
  contrast: number;
  alt: string;
}

export function headerFooterPictureFormat(
  picture: PrintHeaderFooterPicture,
  original: { width: number; height: number },
): HeaderFooterPictureFormat {
  const crop = headerFooterPictureCrop(picture);
  return {
    width: picture.width,
    height: picture.height,
    cropLeft: crop.left * original.width,
    cropTop: crop.top * original.height,
    cropRight: crop.right * original.width,
    cropBottom: crop.bottom * original.height,
    lockAspectRatio: picture.lockAspectRatio !== false,
    colorType: picture.colorType ?? "automatic",
    brightness: picture.brightness ?? 0.5,
    contrast: picture.contrast ?? 0.5,
    alt: picture.alt ?? "",
  };
}

/**
 * The picture a Format Picture result describes. Crop in points becomes the
 * share of the original; a crop that would leave nothing of an axis is
 * refused (null), as Excel refuses it.
 */
export function headerFooterPictureFromFormat(
  base: PrintHeaderFooterPicture,
  format: HeaderFooterPictureFormat,
  original: { width: number; height: number },
): PrintHeaderFooterPicture | null {
  const share = (points: number, of: number) =>
    of > 0 ? Math.max(0, Math.min(1, points / of)) : 0;
  const crop = {
    left: share(format.cropLeft, original.width),
    top: share(format.cropTop, original.height),
    right: share(format.cropRight, original.width),
    bottom: share(format.cropBottom, original.height),
  };
  if (crop.left + crop.right >= 0.999 || crop.top + crop.bottom >= 0.999) return null;
  if (!(format.width > 0) || !(format.height > 0)) return null;
  const clamp = (value: number) =>
    Math.max(0.01, Math.min(MAX_HEADER_FOOTER_PICTURE_POINTS, Math.round(value * 100) / 100));
  const next: PrintHeaderFooterPicture = {
    src: base.src,
    width: clamp(format.width),
    height: clamp(format.height),
  };
  const kept = Object.fromEntries(
    Object.entries(crop)
      .filter(([, value]) => value > 1e-6)
      .map(([side, value]) => [side, Math.round(value * 1e5) / 1e5]),
  );
  if (Object.keys(kept).length > 0) next.crop = kept;
  if (format.colorType !== "automatic") next.colorType = format.colorType;
  if (Math.abs(format.brightness - 0.5) > 1e-6) {
    next.brightness = Math.round(Math.max(0, Math.min(1, format.brightness)) * 1000) / 1000;
  }
  if (Math.abs(format.contrast - 0.5) > 1e-6) {
    next.contrast = Math.round(Math.max(0, Math.min(1, format.contrast)) * 1000) / 1000;
  }
  if (!format.lockAspectRatio) next.lockAspectRatio = false;
  const alt = format.alt.trim() === "" ? "" : format.alt;
  if (alt) next.alt = alt;
  // The same picture formatted: still the rendering of the package's own
  
  if (base.metafileRendered === true && base.sourcePart) {
    next.metafileRendered = true;
    next.sourcePart = base.sourcePart;
  }
  return next;
}

/**
 * Height and width after one of them is edited: with Lock aspect ratio the
 * other follows in proportion (Excel's Size tab).
 */
export function headerFooterPictureResized(
  format: Pick<HeaderFooterPictureFormat, "width" | "height" | "lockAspectRatio">,
  edit: { width?: number; height?: number },
): { width: number; height: number } {
  if (edit.width !== undefined) {
    const width = Math.max(0.01, edit.width);
    return {
      width,
      height: format.lockAspectRatio && format.width > 0 ? (format.height * width) / format.width : format.height,
    };
  }
  if (edit.height !== undefined) {
    const height = Math.max(0.01, edit.height);
    return {
      height,
      width: format.lockAspectRatio && format.height > 0 ? (format.width * height) / format.height : format.width,
    };
  }
  return { width: format.width, height: format.height };
}

/** The kept (cropped) original size in points, against which Scale is measured. */
export function headerFooterPictureKeptOriginal(
  format: Pick<HeaderFooterPictureFormat, "cropLeft" | "cropTop" | "cropRight" | "cropBottom">,
  original: { width: number; height: number },
): { width: number; height: number } {
  return {
    width: Math.max(0.01, original.width - format.cropLeft - format.cropRight),
    height: Math.max(0.01, original.height - format.cropTop - format.cropBottom),
  };
}
