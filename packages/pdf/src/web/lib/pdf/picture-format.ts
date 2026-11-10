/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  OFFICE_ARTISTIC_EFFECT_FILTERS,
  normalizeOfficeArtisticEffect,
  type OfficeArtisticEffectId,
} from "@mrnote/core/web/lib/office-artistic-effects.ts";
import {
  isOfficeRecolorPreset,
  officeAdjustFilterSteps,
  type OfficeAdjustFilterStep,
  type OfficeRecolorPresetId,
} from "@mrnote/core/web/lib/office-recolor-presets.ts";
import {
  officePictureStyle,
  type OfficePictureStyleId,
} from "@mrnote/core/web/lib/office-picture-styles/catalog.ts";

export type PdfPictureBorderDash = "solid" | "dashed" | "dotted";

export type PdfPictureBorder = Readonly<{
  /** `#rrggbb` */
  color: string;
  widthPt: number;
  dash: PdfPictureBorderDash;
}>;

export interface PdfPictureFormat {
  /** Corrections, Office's whole signed percentages (-100..100). */
  readonly brightness: number;
  readonly contrast: number;
  /** Color Saturation, percent (0..400; 100 leaves the colour). */
  readonly saturation: number;
  /** Color Tone, Kelvin (4700..11200; 6500 leaves the colour). */
  readonly tone: number;
  /** Sharpen/Soften, Office's -100..100 (0 leaves the picture). */
  readonly sharpness: number;
  readonly recolor: OfficeRecolorPresetId | null;
  readonly artisticEffect: Exclude<OfficeArtisticEffectId, "none"> | null;
  readonly style: OfficePictureStyleId | null;
  /**
   * The Picture Border: a line of its own over the style's, "none" for No
   * Outline, null for the style's own line (none without a style).
   */
  readonly border: PdfPictureBorder | "none" | null;
  /** Crop to Shape: an Office shape preset, null for the picture's rectangle. */
  readonly shape: string | null;
}

/** The colour temperature a picture is taken to have: Color Tone at it is none. */
export const PDF_PICTURE_NEUTRAL_TONE = 6500;
export const PDF_PICTURE_TONE_RANGE = Object.freeze({ min: 1500, max: 11500 });

export const PDF_PICTURE_FORMAT_NONE: PdfPictureFormat = Object.freeze({
  brightness: 0,
  contrast: 0,
  saturation: 100,
  tone: PDF_PICTURE_NEUTRAL_TONE,
  sharpness: 0,
  recolor: null,
  artisticEffect: null,
  style: null,
  border: null,
  shape: null,
});

/** Saturation's allowed range (Office's Color Saturation field). */
export const PDF_PICTURE_SATURATION_RANGE = Object.freeze({ min: 0, max: 400 });

const clamp = (value: number, min: number, max: number) =>
  Math.max(min, Math.min(max, value));

/** One format with every field in range (anything unknown taken as unset). */
export function normalizePdfPictureFormat(format: Partial<PdfPictureFormat> | null | undefined): PdfPictureFormat {
  const source = format ?? {};
  const whole = (value: unknown, fallback: number, min: number, max: number) =>
    typeof value === "number" && Number.isFinite(value) ? clamp(Math.round(value), min, max) : fallback;
  const border = source.border;
  return Object.freeze({
    brightness: whole(source.brightness, 0, -100, 100),
    contrast: whole(source.contrast, 0, -100, 100),
    saturation: whole(source.saturation, 100, PDF_PICTURE_SATURATION_RANGE.min, PDF_PICTURE_SATURATION_RANGE.max),
    tone: whole(source.tone, PDF_PICTURE_NEUTRAL_TONE, PDF_PICTURE_TONE_RANGE.min, PDF_PICTURE_TONE_RANGE.max),
    sharpness: whole(source.sharpness, 0, -100, 100),
    recolor: isOfficeRecolorPreset(source.recolor) ? source.recolor : null,
    artisticEffect: normalizeOfficeArtisticEffect(source.artisticEffect),
    style: source.style && officePictureStyle(source.style) ? source.style : null,
    border:
      border === "none"
        ? "none"
        : border && typeof border === "object" && /^#[0-9a-fA-F]{6}$/.test(border.color) &&
            Number.isFinite(border.widthPt) && border.widthPt > 0
          ? Object.freeze({
              color: border.color.toLowerCase(),
              widthPt: clamp(border.widthPt, 0.25, 1584),
              dash: border.dash === "dashed" || border.dash === "dotted" ? border.dash : "solid",
            })
          : null,
    shape: typeof source.shape === "string" && source.shape.trim() ? source.shape.trim() : null,
  });
}

/** Whether the picture's pixels are drawn through a filter. */
export function pdfPictureFormatAdjusts(format: PdfPictureFormat): boolean {
  return (
    format.brightness !== 0 ||
    format.contrast !== 0 ||
    format.saturation !== 100 ||
    format.tone !== PDF_PICTURE_NEUTRAL_TONE ||
    format.sharpness !== 0 ||
    format.recolor !== null ||
    (format.artisticEffect !== null && Boolean(OFFICE_ARTISTIC_EFFECT_FILTERS[format.artisticEffect]))
  );
}

/** Whether a frame is drawn around or over the picture (a style, a border, a shape). */
export function pdfPictureFormatFrames(format: PdfPictureFormat): boolean {
  return format.style !== null || (format.border !== null && format.border !== "none") || format.shape !== null;
}

export function pdfPictureFormatIsPlain(format: PdfPictureFormat): boolean {
  return !pdfPictureFormatAdjusts(format) && !pdfPictureFormatFrames(format);
}

export function samePdfPictureFormat(left: PdfPictureFormat, right: PdfPictureFormat): boolean {
  const border = (value: PdfPictureFormat["border"]) =>
    value === null || value === "none" ? String(value) : `${value.color}/${value.widthPt}/${value.dash}`;
  return (
    left.brightness === right.brightness &&
    left.contrast === right.contrast &&
    left.saturation === right.saturation &&
    left.tone === right.tone &&
    left.sharpness === right.sharpness &&
    left.recolor === right.recolor &&
    left.artisticEffect === right.artisticEffect &&
    left.style === right.style &&
    border(left.border) === border(right.border) &&
    left.shape === right.shape
  );
}

const number = (value: number) => String(Math.round(value * 10_000) / 10_000);

function stepMarkup(step: OfficeAdjustFilterStep): string {
  if (step.type === "matrix") {
    return `<feColorMatrix type="matrix" values="${step.values.map(number).join(" ")}"/>`;
  }
  const funcs = (attributes: string) =>
    ["R", "G", "B"].map((channel) => `<feFunc${channel} ${attributes}/>`).join("");
  if (step.type === "discrete") {
    return `<feComponentTransfer>${funcs(`type="discrete" tableValues="${step.table.map(number).join(" ")}"`)}</feComponentTransfer>`;
  }
  return `<feComponentTransfer>${funcs(`type="linear" slope="${number(step.slope)}" intercept="${number(step.intercept)}"`)}</feComponentTransfer>`;
}

/** A light's colour at `kelvin` (Tanner Helland's fit of the blackbody
 * curve), 0..1 per channel. */
function kelvinWhite(kelvin: number): [number, number, number] {
  const k = kelvin / 100;
  const red = k <= 66 ? 255 : 329.698727446 * Math.pow(k - 60, -0.1332047592);
  const green = k <= 66 ? 99.4708025861 * Math.log(k) - 161.1195681661 : 288.1221695283 * Math.pow(k - 60, -0.0755148492);
  const blue = k >= 66 ? 255 : k <= 19 ? 0 : 138.5177312231 * Math.log(k - 10) - 305.0447927307;
  return [red, green, blue].map((value) => clamp(value, 0, 255) / 255) as [number, number, number];
}

/**
 * Color Tone at `kelvin` as per-channel gains: the picture taken as lit at
 * 6500 K and shown as if balanced for `kelvin` -- a lower temperature cools
 * it, a higher one warms it, as Office's tiles run -- its brightness kept.
 */
export function pdfPictureToneGains(kelvin: number): [number, number, number] {
  const neutral = kelvinWhite(PDF_PICTURE_NEUTRAL_TONE);
  const light = kelvinWhite(clamp(kelvin, PDF_PICTURE_TONE_RANGE.min, PDF_PICTURE_TONE_RANGE.max));
  const gains = neutral.map((value, channel) => value / Math.max(0.05, light[channel])) as [number, number, number];
  const luma = 0.2126 * gains[0] + 0.7152 * gains[1] + 0.0722 * gains[2];
  return gains.map((gain) => gain / luma) as [number, number, number];
}

/**
 * The SVG filter bodies that draw the picture's adjustments: the artistic
 * effect (Office's recipe) and then Color Saturation, Color Tone, Recolor,
 * Corrections and Sharpen/Soften. Null where the format leaves that part
 * alone.
 */
export function pdfPictureAdjustFilters(format: PdfPictureFormat): Readonly<{
  artistic: string | null;
  color: string | null;
}> {
  const artistic = format.artisticEffect ? OFFICE_ARTISTIC_EFFECT_FILTERS[format.artisticEffect] ?? null : null;
  const steps: string[] = [];
  if (format.saturation !== 100) {
    steps.push(`<feColorMatrix type="saturate" values="${number(format.saturation / 100)}"/>`);
  }
  if (format.tone !== PDF_PICTURE_NEUTRAL_TONE) {
    const [r, g, b] = pdfPictureToneGains(format.tone);
    steps.push(`<feColorMatrix type="matrix" values="${[r, 0, 0, 0, 0, 0, g, 0, 0, 0, 0, 0, b, 0, 0, 0, 0, 0, 1, 0].map(number).join(" ")}"/>`);
  }
  for (const step of officeAdjustFilterSteps({
    recolor: format.recolor,
    brightness: format.brightness,
    contrast: format.contrast,
  })) {
    steps.push(stepMarkup(step));
  }
  if (format.sharpness > 0) {
    // Sharpen: an unsharp kernel, +100 % the strongest Office draws.
    const k = (format.sharpness / 100) * 1.2;
    steps.push(`<feConvolveMatrix order="3" kernelMatrix="0 ${number(-k)} 0 ${number(-k)} ${number(1 + 4 * k)} ${number(-k)} 0 ${number(-k)} 0" preserveAlpha="true" edgeMode="duplicate"/>`);
  } else if (format.sharpness < 0) {
    // Soften: a blur, -100 % about three page pixels.
    steps.push(`<feGaussianBlur stdDeviation="${number((-format.sharpness / 100) * 3)}" edgeMode="duplicate"/>`);
  }
  return Object.freeze({ artistic, color: steps.length ? steps.join("") : null });
}

const escapeAttribute = (value: string) =>
  value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/**
 * The picture through its adjustments as one SVG document: `href` (a data
 * URL) laid over `viewWidth` x `viewHeight` -- the picture's size on the page,
 * so the effects' blurs and textures measure as they do on screen -- drawn at
 * `width` x `height` pixels. Null when the format adjusts nothing.
 */
export function pdfPictureAdjustSvg(input: Readonly<{
  href: string;
  width: number;
  height: number;
  viewWidth: number;
  viewHeight: number;
  format: PdfPictureFormat;
}>): string | null {
  const { artistic, color } = pdfPictureAdjustFilters(input.format);
  if (!artistic && !color) return null;
  const filter = (id: string, body: string) =>
    `<filter id="${id}" x="0" y="0" width="100%" height="100%" filterUnits="objectBoundingBox" color-interpolation-filters="sRGB">${body}</filter>`;
  const view = `${number(input.viewWidth)} ${number(input.viewHeight)}`;
  const image = `<image href="${escapeAttribute(input.href)}" x="0" y="0" width="${number(input.viewWidth)}" height="${number(input.viewHeight)}" preserveAspectRatio="none"${artistic ? ' filter="url(#a)"' : ""}/>`;
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.round(input.width)}" height="${Math.round(input.height)}" viewBox="0 0 ${view}" preserveAspectRatio="none">` +
    `<defs>${artistic ? filter("a", artistic) : ""}${color ? filter("c", color) : ""}</defs>` +
    (color ? `<g filter="url(#c)">${image}</g>` : image) +
    "</svg>"
  );
}

/** Word's Compress Pictures resolutions (pixels per inch), Office's order. */
export const PDF_PICTURE_COMPRESS_PPI = [330, 220, 150, 96] as const;
export type PdfPictureCompressPpi = (typeof PDF_PICTURE_COMPRESS_PPI)[number];

/**
 * The pixel size a picture of `pixels` shown at `points` is kept at for
 * `ppi`: never larger than it is (Word only ever lowers the resolution).
 */
export function pdfPictureCompressedSize(
  pixels: Readonly<{ width: number; height: number }>,
  points: Readonly<{ width: number; height: number }>,
  ppi: number,
): Readonly<{ width: number; height: number }> {
  const scale = Math.min(
    1,
    (Math.abs(points.width) / 72) * ppi / Math.max(1, pixels.width),
    (Math.abs(points.height) / 72) * ppi / Math.max(1, pixels.height),
  );
  return Object.freeze({
    width: Math.max(1, Math.round(pixels.width * scale)),
    height: Math.max(1, Math.round(pixels.height * scale)),
  });
}
