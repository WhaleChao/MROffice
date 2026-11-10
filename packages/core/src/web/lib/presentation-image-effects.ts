/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation image block visual effects (PowerPoint / WPS parity subset):
 * a soft drop shadow, a coloured frame (border colour + width), and rounded
 * corners.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the present-mode overlay so the effect bounds + default
 * resolution + the render-ready CSS live in exactly one place. Keep this module
 * dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the zod schema.
 *
 * Mirrors `presentation-shapes.ts`: a six-digit-hex colour validator, clamped
 * numeric bounds, a normalizer that collapses unknown/stale/hand-edited values
 * to a safe default, and a CSS resolver. An image with NO effects resolves to a
 * descriptor whose `hasEffects` is false so the caller can keep the historical
 * plain-`<img>` render path byte-for-byte.
 */

import { CSS_PX_PER_POINT } from "./presentation/text-units";
import type { CSSProperties } from "react";

import {
  DEFAULT_SHAPE_BLEND_MODE,
  normalizeShapeBlendMode,
  type ShapeBlendMode,
} from "./presentation-shapes";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";
import { officeLuminanceMap } from "./office-recolor-presets";

/** Six-digit hex colour, e.g. `#1a2b3c`. Matches the shape colour regex. */
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

/** Default frame colour when a border is enabled without an explicit colour. */
export const DEFAULT_IMAGE_BORDER_COLOR = "#1f2937";
/** Inclusive bounds (px) the frame weight is clamped to. 0 ⇒ no frame. */
export const MAX_IMAGE_BORDER_WIDTH = 24;
/** DrawingML's visible default when `a:ln@w` is omitted (one point). */
export const DEFAULT_IMAGE_BORDER_WIDTH = CSS_PX_PER_POINT;
/** Inclusive bounds (px) the corner radius is clamped to. 0 ⇒ square. */
export const MAX_IMAGE_CORNER_RADIUS = 64;

/** A soft drop shadow used when `shadow` is set (matches the shape shadow). */
export const IMAGE_SHADOW_CSS = "0 6px 16px rgba(15, 23, 42, 0.30)";

/** Persisted/rendered recolor modes shipped before native Picture Color. */
export const PERSISTED_IMAGE_RECOLOR_MODES = [
  "none",
  "grayscale",
  "sepia",
  "washout",
  "invert",
  "duotone",
  "vivid",
] as const;
export type ImageRecolorMode = (typeof PERSISTED_IMAGE_RECOLOR_MODES)[number];
const IMAGE_RECOLOR_SET: ReadonlySet<string> = new Set(
  PERSISTED_IMAGE_RECOLOR_MODES,
);

/** The lossless native Picture Color command subset. */
export const IMAGE_RECOLOR_MODES = ["none", "grayscale"] as const;
export const NATIVE_IMAGE_RECOLOR_MODES = IMAGE_RECOLOR_MODES;
export type NativeImageRecolorMode =
  (typeof NATIVE_IMAGE_RECOLOR_MODES)[number];
const NATIVE_IMAGE_RECOLOR_SET: ReadonlySet<string> = new Set(
  NATIVE_IMAGE_RECOLOR_MODES,
);

/**
 * Inclusive bounds for the brightness / contrast / opacity adjustments, as a
 * PERCENTAGE where 100 is the unmodified image. PowerPoint's brightness /
 * contrast sliders span roughly ±100%; we clamp to [0, 200] (0 = black / flat,
 * 200 = double). Opacity is [0, 100] (% visible).
 */
export const IMAGE_ADJUST_MIN = 0;
export const IMAGE_ADJUST_MAX = 200;
export const DEFAULT_IMAGE_ADJUST = 100;
export const IMAGE_SATURATION_MIN = 0;
export const IMAGE_SATURATION_MAX = 200;
export const DEFAULT_IMAGE_SATURATION = 100;
export const IMAGE_SATURATION_PRESETS = [0, 33, 66, 100, 150, 200] as const;
export const PICTURE_CORRECTIONS_STORED_UNIT = "stored-percent-v1" as const;
export const PICTURE_CORRECTIONS_AMBIGUOUS_LEGACY_FLAG =
  "pictureCorrectionsLegacyAmbiguous" as const;
/** Whole signed deltas exposed by the Picture Corrections preset gallery. */
export const IMAGE_CORRECTION_PRESET_DELTAS = [-40, -20, 0, 20, 40] as const;
export type ImageCorrectionPresetDelta =
  (typeof IMAGE_CORRECTION_PRESET_DELTAS)[number];
export type PictureCorrectionPreset = Readonly<{
  id: string;
  brightnessDelta: ImageCorrectionPresetDelta;
  contrastDelta: ImageCorrectionPresetDelta;
}>;
const IMAGE_CORRECTION_PRESET_DELTA_SET: ReadonlySet<number> = new Set(
  IMAGE_CORRECTION_PRESET_DELTAS,
);

export function pictureCorrectionPresetId(
  brightnessDelta: ImageCorrectionPresetDelta,
  contrastDelta: ImageCorrectionPresetDelta,
): string {
  return `brightness-${brightnessDelta}-contrast-${contrastDelta}`;
}

/** PowerPoint's bounded 5×5 Brightness/Contrast gallery identities. */
export const PICTURE_CORRECTION_PRESETS: readonly PictureCorrectionPreset[] =
  Object.freeze(
    IMAGE_CORRECTION_PRESET_DELTAS.flatMap((brightnessDelta) =>
      IMAGE_CORRECTION_PRESET_DELTAS.map((contrastDelta) =>
        Object.freeze({
          id: pictureCorrectionPresetId(brightnessDelta, contrastDelta),
          brightnessDelta,
          contrastDelta,
        }),
      ),
    ),
  );

export function resolvePictureCorrectionPreset(
  id: unknown,
): PictureCorrectionPreset | null {
  if (typeof id !== "string") return null;
  const match = /^brightness-(-?[0-9]+)-contrast-(-?[0-9]+)$/.exec(id);
  if (!match) return null;
  const brightnessDelta = Number(match[1]);
  const contrastDelta = Number(match[2]);
  if (
    !IMAGE_CORRECTION_PRESET_DELTA_SET.has(brightnessDelta) ||
    !IMAGE_CORRECTION_PRESET_DELTA_SET.has(contrastDelta)
  ) {
    return null;
  }
  return {
    id,
    brightnessDelta: brightnessDelta as ImageCorrectionPresetDelta,
    contrastDelta: contrastDelta as ImageCorrectionPresetDelta,
  };
}
export const IMAGE_OPACITY_MIN = 0;
export const IMAGE_OPACITY_MAX = 100;
export const DEFAULT_IMAGE_OPACITY = 100;
/** PowerPoint 16.86 Picture Transparency gallery, in displayed percentages. */
export const IMAGE_TRANSPARENCY_PRESETS = [0, 15, 30, 50, 65, 80, 95] as const;
export type ImageTransparencyPreset =
  (typeof IMAGE_TRANSPARENCY_PRESETS)[number];

/** Normalize a persisted recolor mode, defaulting to `none`. */
export function normalizeImageRecolor(value: unknown): ImageRecolorMode {
  return typeof value === "string" && IMAGE_RECOLOR_SET.has(value)
    ? (value as ImageRecolorMode)
    : "none";
}

/** True for every persisted/render-compatible recolor mode. */
export function isImageRecolorMode(value: unknown): value is ImageRecolorMode {
  return typeof value === "string" && IMAGE_RECOLOR_SET.has(value);
}

/** True only for the typed, lossless native Picture Color command domain. */
export function isNativeImageRecolorMode(
  value: unknown,
): value is NativeImageRecolorMode {
  return typeof value === "string" && NATIVE_IMAGE_RECOLOR_SET.has(value);
}

/** True only when native colour grammar is proven editable. */
export function hasEditablePictureColor(source: unknown): boolean {
  return !(
    typeof source === "object" &&
    source !== null &&
    !Array.isArray(source) &&
    (source as Record<string, unknown>).pictureColorDisposition ===
      "preserve-only"
  );
}

/**
 * Clamp a persisted brightness / contrast percentage to [IMAGE_ADJUST_MIN,
 * IMAGE_ADJUST_MAX] (rounded), defaulting to {@link DEFAULT_IMAGE_ADJUST} (100,
 * unmodified) for a missing / non-finite value.
 */
export function normalizeImageAdjust(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_IMAGE_ADJUST;
  }
  const rounded = Math.round(value);
  if (rounded < IMAGE_ADJUST_MIN) return IMAGE_ADJUST_MIN;
  return rounded > IMAGE_ADJUST_MAX ? IMAGE_ADJUST_MAX : rounded;
}

/** Native Picture Color saturation percentage; 100 is the unmodified image. */
export function normalizeImageSaturation(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_IMAGE_SATURATION;
  }
  const rounded = Math.round(value);
  if (rounded < IMAGE_SATURATION_MIN) return IMAGE_SATURATION_MIN;
  return rounded > IMAGE_SATURATION_MAX ? IMAGE_SATURATION_MAX : rounded;
}

/** True only for the typed native Picture Corrections command domain. */
export function isImageCorrectionDelta(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    Number.isInteger(value) &&
    value >= -100 &&
    value <= 100
  );
}

type PictureCorrectionsStorageMode =
  "stored" | "signed" | "ambiguous" | "invalid";

function correctionInteger(value: unknown): number | null {
  return typeof value === "number" &&
    Number.isFinite(value) &&
    Number.isInteger(value)
    ? value
    : null;
}

/** Classify an unmarked native pair without treating identity as provenance. */
function pictureCorrectionsStorageMode(
  brightness: unknown,
  contrast: unknown,
  source: unknown,
): PictureCorrectionsStorageMode {
  if (!hasNativePictureSource(source)) return "stored";
  const record = source as Record<string, unknown>;
  if (record.pictureCorrectionsUnit === PICTURE_CORRECTIONS_STORED_UNIT) {
    return "stored";
  }
  if (record.pictureCorrectionsUnit !== undefined) return "invalid";
  const values = [brightness, contrast]
    .filter((value) => value !== undefined)
    .map(correctionInteger);
  if (values.some((value) => value === null)) return "invalid";
  const integers = values as number[];
  if (integers.length === 0) return "stored";
  if (integers.some((value) => value < 0)) {
    return integers.every((value) => value >= -100 && value <= 100)
      ? "signed"
      : "invalid";
  }
  if (integers.some((value) => value > 100)) {
    return integers.every((value) => value >= 0 && value <= 200)
      ? "stored"
      : "invalid";
  }
  return "ambiguous";
}

/** Resolve the pair together so one negative authenticates both signed values. */
function resolvedPictureCorrectionsStored(
  brightness: unknown,
  contrast: unknown,
  source: unknown,
): { brightness: number; contrast: number } {
  const mode = pictureCorrectionsStorageMode(brightness, contrast, source);
  if (mode === "signed") {
    return {
      brightness:
        brightness === undefined
          ? DEFAULT_IMAGE_ADJUST
          : (brightness as number) + DEFAULT_IMAGE_ADJUST,
      contrast:
        contrast === undefined
          ? DEFAULT_IMAGE_ADJUST
          : (contrast as number) + DEFAULT_IMAGE_ADJUST,
    };
  }
  return {
    brightness: normalizeImageAdjust(brightness),
    contrast: normalizeImageAdjust(contrast),
  };
}

/** Adapt one already-canonical persisted Web percentage. */
export function imageCorrectionStoredFromPersisted(value: unknown): number {
  return normalizeImageAdjust(value);
}

/** Whether the persisted unit marker is absent (legacy) or the current v1. */
export function hasSupportedPictureCorrectionsStoredUnit(
  source: unknown,
): boolean {
  if (typeof source !== "object" || source === null || Array.isArray(source)) {
    return true;
  }
  const unit = (source as Record<string, unknown>).pictureCorrectionsUnit;
  return unit === undefined || unit === PICTURE_CORRECTIONS_STORED_UNIT;
}

/** True only when a live command can prove the current correction domain. */
export function hasEditablePictureCorrections(
  brightness: unknown,
  contrast: unknown,
  source: unknown,
): boolean {
  if (typeof source === "object" && source !== null && !Array.isArray(source)) {
    if (
      (source as Record<string, unknown>).pictureCorrectionsDisposition ===
      "preserve-only"
    ) {
      return false;
    }
  }
  const mode = pictureCorrectionsStorageMode(brightness, contrast, source);
  return mode === "stored" || mode === "signed";
}

/** Stamp the current unit only for an authenticated native picture lineage. */
export function pictureCorrectionsSourceWithStoredUnit(
  source: unknown,
): unknown {
  if (!hasNativePictureSource(source)) return source;
  const record = source as Record<string, unknown>;
  if (record.pictureCorrectionsUnit !== undefined) return source;
  return {
    ...record,
    pictureCorrectionsUnit: PICTURE_CORRECTIONS_STORED_UNIT,
  };
}

/**
 * Add a durable unit only when the pair proves its representation. Ambiguous
 * positive 0..100 values remain byte-identical and become preserve-only, so an
 * old imported picture later edited to canonical 20 cannot silently become
 * 120 merely because it still has slide/shape identity.
 */
export function pictureCorrectionsSourceForLoad(
  source: unknown,
  brightness: unknown,
  contrast: unknown,
): unknown {
  if (!hasNativePictureSource(source)) return source;
  const record = source as Record<string, unknown>;
  if (record.pictureCorrectionsUnit !== undefined) return source;
  const mode = pictureCorrectionsStorageMode(brightness, contrast, source);
  if (mode === "ambiguous") {
    if (
      record.pictureCorrectionsDisposition === "preserve-only" &&
      record[PICTURE_CORRECTIONS_AMBIGUOUS_LEGACY_FLAG] === true
    ) {
      return source;
    }
    return {
      ...record,
      pictureCorrectionsDisposition: "preserve-only",
      [PICTURE_CORRECTIONS_AMBIGUOUS_LEGACY_FLAG]: true,
    };
  }
  if (mode === "stored" || mode === "signed") {
    return {
      ...record,
      pictureCorrectionsUnit: PICTURE_CORRECTIONS_STORED_UNIT,
    };
  }
  return source;
}

/** Adapt persisted Web storage (neutral 100) to a signed native delta. */
export function imageCorrectionDeltaFromStored(value: unknown): number {
  return imageCorrectionStoredFromPersisted(value) - DEFAULT_IMAGE_ADJUST;
}

/** Resolve both fields before adapting so a negative can authenticate the pair. */
export function imageCorrectionDeltasFromPersisted(
  brightness: unknown,
  contrast: unknown,
  source: unknown,
): { brightnessDelta: number; contrastDelta: number } {
  const stored = resolvedPictureCorrectionsStored(brightness, contrast, source);
  return {
    brightnessDelta: stored.brightness - DEFAULT_IMAGE_ADJUST,
    contrastDelta: stored.contrast - DEFAULT_IMAGE_ADJUST,
  };
}

/** Adapt a validated signed native delta back to persisted Web storage. */
export function imageCorrectionStoredFromDelta(value: number): number {
  if (!isImageCorrectionDelta(value)) {
    throw new RangeError(
      "Picture correction delta must be a whole -100..100 value",
    );
  }
  return value + DEFAULT_IMAGE_ADJUST;
}

/**
 * Clamp a persisted opacity percentage to [IMAGE_OPACITY_MIN, IMAGE_OPACITY_MAX]
 * (rounded), defaulting to {@link DEFAULT_IMAGE_OPACITY} (100, fully opaque) for
 * a missing / non-finite value.
 */
export function normalizeImageOpacity(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_IMAGE_OPACITY;
  }
  const rounded = Math.round(value);
  if (rounded < IMAGE_OPACITY_MIN) return IMAGE_OPACITY_MIN;
  return rounded > IMAGE_OPACITY_MAX ? IMAGE_OPACITY_MAX : rounded;
}

function hasNativePictureSource(value: unknown): boolean {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const source = value as Record<string, unknown>;
  return [source.slideId, source.shapeId].every(
    (part) => typeof part === "string" && part.trim().length > 0,
  );
}

/**
 * Repair the one legacy imported representation that stored visible opacity as
 * a 0..1 fraction.  A fraction is ambiguous without the native slide + shape
 * identity, so authored/unaddressed pictures continue through the ordinary
 * percentage normalizer and the API remains their strict validation boundary.
 */
function isLegacyImportedImageOpacity(
  value: unknown,
  source: unknown,
): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value > 0 &&
    value < 1 &&
    !Number.isInteger(value) &&
    hasNativePictureSource(source)
  );
}

function normalizePersistedImageOpacity(
  value: unknown,
  source: unknown,
): number {
  return isLegacyImportedImageOpacity(value, source)
    ? Math.round(value * 100)
    : normalizeImageOpacity(value);
}

/** Clamp a displayed Transparency percentage; missing means 0% transparent. */
export function normalizeImageTransparency(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 0;
  const rounded = Math.round(value);
  if (rounded < 0) return 0;
  return rounded > 100 ? 100 : rounded;
}

/** Convert PowerPoint's displayed Transparency to persisted visible opacity. */
export function imageOpacityFromTransparency(value: unknown): number {
  return DEFAULT_IMAGE_OPACITY - normalizeImageTransparency(value);
}

/** Convert persisted visible opacity to PowerPoint's displayed Transparency. */
export function imageTransparencyFromOpacity(value: unknown): number {
  return DEFAULT_IMAGE_OPACITY - normalizeImageOpacity(value);
}

/**
 * Normalize an arbitrary colour value into a six-digit hex string, or return
 * `fallback` when the value is missing / malformed. Always lower-cases so a
 * persisted deck stays byte-stable. Identical contract to the shape colour
 * validator so a stale or hand-edited deck can never inject arbitrary CSS.
 */
export function normalizeImageBorderColor<F extends string | null>(
  value: unknown,
  fallback: F,
): string | F {
  if (typeof value === "string" && HEX_COLOR_RE.test(value.trim())) {
    return value.trim().toLowerCase();
  }
  return fallback;
}

/**
 * Clamp a persisted px length to `[0, max]` (rounded), else 0. Used for lengths
 * whose native contract is integral, such as corner radius; picture-border
 * weight intentionally has its own fractional-preserving normalizer below.
 */
function clampPxLength(value: unknown, max: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 0;
  const rounded = Math.round(value);
  if (rounded <= 0) return 0;
  return rounded > max ? max : rounded;
}

/** Clamp a persisted frame weight to `[0, MAX_IMAGE_BORDER_WIDTH]` px. */
export function normalizeImageBorderWidth(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return 0;
  }
  return Math.min(value, MAX_IMAGE_BORDER_WIDTH);
}

export const IMAGE_BORDER_STYLES = [
  "none",
  "solid",
  "dashed",
  "dotted",
] as const;
export type ImageBorderStyle = (typeof IMAGE_BORDER_STYLES)[number];
const IMAGE_BORDER_STYLE_SET: ReadonlySet<string> = new Set(
  IMAGE_BORDER_STYLES,
);

export function normalizeImageBorderStyle(
  value: unknown,
  fallback: ImageBorderStyle = "none",
): ImageBorderStyle {
  return typeof value === "string" && IMAGE_BORDER_STYLE_SET.has(value)
    ? (value as ImageBorderStyle)
    : fallback;
}

/** Clamp a persisted corner radius to `[0, MAX_IMAGE_CORNER_RADIUS]` px. */
export function normalizeImageCornerRadius(value: unknown): number {
  return clampPxLength(value, MAX_IMAGE_CORNER_RADIUS);
}








function persistedImageCornerRadius(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 0;
  const clamped = Math.min(MAX_IMAGE_CORNER_RADIUS, Math.max(0, value));
  return Math.round(clamped * 100) / 100;
}

/** Native roundRect keeps a relative adjustment when its frame is resized.
 * Derive its visible radius without rewriting the protected original scalar.
 * Source-less user-authored radii remain literal design pixels.
 */
export function effectivePresentationImageCornerRadius(block: {
  cornerRadius?: unknown;
  width?: unknown;
  height?: unknown;
  source?: unknown;
}): number {
  const radius = typeof block.cornerRadius === "number" && Number.isFinite(block.cornerRadius)
    ? Math.min(MAX_IMAGE_CORNER_RADIUS, Math.max(0, block.cornerRadius)) : 0;
  const source = block.source && typeof block.source === "object"
    ? block.source as Record<string, unknown> : null;
  const sizes = [block.width, block.height, source?.pictureSizeOriginalWidthEmu, source?.pictureSizeOriginalHeightEmu];
  if (!radius || source?.pictureSizeSchemaVersion !== "pptx-picture-size-v1" ||
      !sizes.every(value => typeof value === "number" && Number.isSafeInteger(value) && value > 0 && value <= PRESENTATION_GEOMETRY_MAX_EMU)) return radius;
  const [width, height, originalWidth, originalHeight] = sizes as number[];
  return Math.round(Math.min(MAX_IMAGE_CORNER_RADIUS,
    radius * Math.min(width, height) / Math.min(originalWidth, originalHeight)) * 100) / 100;
}

/**
 * The raw, optional image-effect fields a block may persist. All optional /
 * additive — an image without any of them resolves to "no effects".
 */
export interface ImageEffectsInput {
  shadow?: unknown;
  borderColor?: unknown;
  borderWidth?: unknown;
  borderStyle?: unknown;
  cornerRadius?: unknown;
  /** Picture ADJUSTMENTS (PowerPoint "Color" / "Corrections" parity). */
  opacity?: unknown;
  /** Native identity used only to repair legacy fractional picture opacity. */
  source?: unknown;
  recolor?: unknown;
  saturation?: unknown;
  brightness?: unknown;
  contrast?: unknown;
  /** Compositing: a CSS `mix-blend-mode` (shared whitelist with shapes). */
  blendMode?: unknown;
}

/** A resolved, render-ready image-effects descriptor. */
export interface ResolvedImageEffects {
  /** Whether a soft drop shadow is drawn under the image. */
  shadow: boolean;
  /** Frame weight in px (0 ⇒ no frame). */
  borderWidth: number;
  /** Frame colour (only meaningful when `borderWidth > 0`). */
  borderColor: string;
  /** CSS/native line style; `none` also represents an explicit no-outline. */
  borderStyle: ImageBorderStyle;
  /** Corner radius in px (0 ⇒ square). */
  cornerRadius: number;
  /** Opacity percentage [0, 100] (100 ⇒ fully opaque, the default). */
  opacity: number;
  /** Recolor mode (`none` ⇒ full colour, the default). */
  recolor: ImageRecolorMode;
  /** Color saturation percentage [0, 200] (100 ⇒ unmodified). */
  saturation: number;
  /** Brightness percentage [0, 200] (100 ⇒ unmodified, the default). */
  brightness: number;
  /** Contrast percentage [0, 200] (100 ⇒ unmodified, the default). */
  contrast: number;
  /** CSS `mix-blend-mode` (`normal` ⇒ ordinary stacking, the default). */
  blendMode: ShapeBlendMode;
  /** True iff at least one effect is active (so the caller can branch render). */
  hasEffects: boolean;
}

/**
 * Resolve a block's persisted effect fields into a normalized descriptor. A
 * border colour is only retained when a positive width makes it visible; the
 * descriptor is `hasEffects: false` exactly when no shadow, no frame, and no
 * radius apply — letting the renderer keep the plain-`<img>` path unchanged.
 */
export function resolveImageEffects(
  input: ImageEffectsInput | undefined,
): ResolvedImageEffects {
  const shadow = input?.shadow === true;
  const persistedWidth = normalizeImageBorderWidth(input?.borderWidth);
  const explicitStyle = normalizeImageBorderStyle(input?.borderStyle, "none");
  const borderStyle =
    input?.borderStyle !== undefined
      ? explicitStyle
      : persistedWidth > 0
        ? "solid"
        : "none";
  const borderWidth =
    borderStyle === "none"
      ? 0
      : persistedWidth > 0
        ? persistedWidth
        : DEFAULT_IMAGE_BORDER_WIDTH;
  const borderColor =
    borderWidth > 0
      ? normalizeImageBorderColor(
          input?.borderColor,
          DEFAULT_IMAGE_BORDER_COLOR,
        )
      : DEFAULT_IMAGE_BORDER_COLOR;
  const cornerRadius = normalizeImageCornerRadius(input?.cornerRadius);
  const opacity = normalizePersistedImageOpacity(input?.opacity, input?.source);
  const recolor = normalizeImageRecolor(input?.recolor);
  const saturation = normalizeImageSaturation(input?.saturation);
  const { brightness, contrast } = resolvedPictureCorrectionsStored(
    input?.brightness,
    input?.contrast,
    input?.source,
  );
  const blendMode = normalizeShapeBlendMode(input?.blendMode);
  const hasAdjust =
    opacity !== DEFAULT_IMAGE_OPACITY ||
    recolor !== "none" ||
    saturation !== DEFAULT_IMAGE_SATURATION ||
    brightness !== DEFAULT_IMAGE_ADJUST ||
    contrast !== DEFAULT_IMAGE_ADJUST ||
    blendMode !== DEFAULT_SHAPE_BLEND_MODE;
  return {
    shadow,
    borderWidth,
    borderColor,
    borderStyle,
    cornerRadius,
    opacity,
    recolor,
    saturation,
    brightness,
    contrast,
    blendMode,
    hasEffects: shadow || borderWidth > 0 || cornerRadius > 0 || hasAdjust,
  };
}

/**
 * The CSS `filter` value composing the resolved recolor + brightness + contrast
 * adjustments, or `undefined` when all are at their defaults (so an un-adjusted
 * image keeps no `filter` and renders byte-for-byte). Recolor is layered FIRST
 * so brightness / contrast apply on top of the tint, matching PowerPoint's
 * "Color" → "Corrections" order.
 */
export function imageFilterCss(
  effects: ResolvedImageEffects,
): string | undefined {
  const parts: string[] = [];
  switch (effects.recolor) {
    case "grayscale":
      parts.push("grayscale(1)");
      break;
    case "sepia":
      parts.push("sepia(0.65)");
      break;
    case "washout":
      // PowerPoint's Washout tile IS a luminance change (a:lum bright 70 %,
      // contrast -70 %), drawn by the same map as the corrections.
      parts.push(...officeCorrectionFilters(170, 30));
      break;
    case "invert":
      parts.push("invert(1)");
      break;
    case "duotone":
      parts.push("hue-rotate(320deg)", "saturate(0.65)");
      break;
    case "vivid":
      parts.push("contrast(1.4)", "saturate(1.5)");
      break;
    case "none":
    default:
      break;
  }
  if (effects.saturation !== DEFAULT_IMAGE_SATURATION) {
    parts.push(`saturate(${effects.saturation}%)`);
  }
  parts.push(...officeCorrectionFilters(effects.brightness, effects.contrast));
  return parts.length > 0 ? parts.join(" ") : undefined;
}















export function officeCorrectionFilters(
  brightnessStored: number,
  contrastStored: number,
): string[] {
  const delta = (stored: number) =>
    Math.max(-1, Math.min(1, (stored - DEFAULT_IMAGE_ADJUST) / 100));
  const brightness = delta(brightnessStored);
  const contrast = delta(contrastStored);
  if (brightness === 0 && contrast === 0) return [];
  const { slope, intercept } = officeLuminanceMap(brightness, contrast);
  const round = (value: number) => Number(value.toFixed(6));
  const term = (name: string, value: number) =>
    round(value) === 1 ? [] : [`${name}(${round(value)})`];
  if (intercept >= 0) {
    const scale = slope + 2 * intercept;
    if (scale <= 0) return ["brightness(0)"];
    return [...term("contrast", slope / scale), ...term("brightness", scale)];
  }
  const k = 1 - 2 * intercept;
  return [...term("brightness", slope / k), ...term("contrast", k)];
}












export function brightnessOffsetFilters(stored: number): string[] {
  const delta = Math.max(-1, Math.min(1, (stored - DEFAULT_IMAGE_ADJUST) / 100));
  if (delta === 0) return [];
  const round = (value: number) => Number(value.toFixed(6));
  if (delta > 0) {
    const scale = 1 + 2 * delta;
    return [`contrast(${round(1 / scale)})`, `brightness(${round(scale)})`];
  }
  const scale = 1 - 2 * delta;
  return [`brightness(${round(1 / scale)})`, `contrast(${round(scale)})`];
}

/**
 * The CSS to layer onto the visible image element for the resolved effects.
 * Returns an empty object when no effect is active (so spreading it onto the
 * plain `<img>` is a no-op and the historical render stays byte-stable). The
 * frame uses `border` (drawn outside the image content via `box-sizing`),
 * `borderRadius` clips the corners, `boxShadow` is the soft drop shadow, and the
 * picture adjustments (recolor / brightness / contrast → `filter`; transparency
 * → `opacity`) tint / fade the picture.
 */
export function imageEffectsCss(effects: ResolvedImageEffects): CSSProperties {
  if (!effects.hasEffects) return {};
  const css: CSSProperties = {};
  if (effects.borderWidth > 0) {
    css.border = `${effects.borderWidth}px ${effects.borderStyle} ${effects.borderColor}`;
  }
  if (effects.cornerRadius > 0) {
    css.borderRadius = effects.cornerRadius;
  }
  if (effects.shadow) css.boxShadow = IMAGE_SHADOW_CSS;
  const filter = imageFilterCss(effects);
  if (filter) css.filter = filter;
  if (effects.opacity !== DEFAULT_IMAGE_OPACITY) {
    css.opacity = effects.opacity / 100;
  }
  // Compositing: only emit `mix-blend-mode` for a non-default mode so an
  // ordinary (normal) picture keeps no blend property and renders byte-for-byte.
  if (effects.blendMode !== DEFAULT_SHAPE_BLEND_MODE) {
    css.mixBlendMode = effects.blendMode;
  }
  return css;
}

/** The persisted image-effect field names handled by the normalizers. */
export const IMAGE_EFFECT_KEYS = [
  "shadow",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "cornerRadius",
  "opacity",
  "recolor",
  "brightness",
  "contrast",
  "blendMode",
] as const;

/**
 * The effect fields to PERSIST for a block, omitting every field at its default
 * (no shadow / zero width / zero radius / colour that matches the default while
 * no frame is shown) so an image without meaningful effects serializes
 * byte-for-byte as before. Returns an empty object in that case.
 *
 * `borderColor` is persisted only when a frame is actually drawn. An explicit
 * style with omitted width uses DrawingML's one-point visual default while
 * keeping `borderWidth` absent, so native width omission survives roundtrip.
 */
export function persistedImageEffects(
  input: ImageEffectsInput | undefined,
): Partial<{
  shadow: true;
  borderColor: string;
  borderWidth: number;
  borderStyle: ImageBorderStyle;
  cornerRadius: number;
  opacity: number;
  recolor: ImageRecolorMode;
  saturation: number;
  brightness: number;
  contrast: number;
  blendMode: ShapeBlendMode;
}> {
  const effects = resolveImageEffects(input);
  const out: ReturnType<typeof persistedImageEffects> = {};
  if (effects.shadow) out.shadow = true;
  const persistedWidth = normalizeImageBorderWidth(input?.borderWidth);
  const hasExplicitStyle = input?.borderStyle !== undefined;
  if (hasExplicitStyle && effects.borderStyle === "none") {
    out.borderStyle = "none";
    if (persistedWidth > 0) out.borderWidth = persistedWidth;
  } else if (effects.borderWidth > 0) {
    if (persistedWidth > 0) out.borderWidth = persistedWidth;
    out.borderColor = effects.borderColor;
    if (hasExplicitStyle) out.borderStyle = effects.borderStyle;
  }
  const persistedRadius = persistedImageCornerRadius(input?.cornerRadius);
  if (persistedRadius > 0) out.cornerRadius = persistedRadius;
  // Picture adjustments: persist ONLY non-default values so an un-adjusted image
  // stays byte-clean (no lingering `opacity: 100` / `recolor: "none"` keys).
  const rawOpacity = input?.opacity;
  if (
    typeof rawOpacity === "number" &&
    Number.isFinite(rawOpacity) &&
    rawOpacity >= IMAGE_OPACITY_MIN &&
    rawOpacity <= IMAGE_OPACITY_MAX &&
    !Number.isInteger(rawOpacity) &&
    !isLegacyImportedImageOpacity(rawOpacity, input?.source)
  ) {
    // Keep an unaddressed/authored fractional percentage invalid through the
    // save boundary so the typed API rejects it instead of silently rounding.
    out.opacity = rawOpacity;
  } else if (effects.opacity !== DEFAULT_IMAGE_OPACITY) {
    out.opacity = effects.opacity;
  }
  if (effects.recolor !== "none") out.recolor = effects.recolor;
  if (effects.saturation !== DEFAULT_IMAGE_SATURATION)
    out.saturation = effects.saturation;
  if (effects.brightness !== DEFAULT_IMAGE_ADJUST)
    out.brightness = effects.brightness;
  if (effects.contrast !== DEFAULT_IMAGE_ADJUST)
    out.contrast = effects.contrast;
  if (effects.blendMode !== DEFAULT_SHAPE_BLEND_MODE)
    out.blendMode = effects.blendMode;
  return out;
}
