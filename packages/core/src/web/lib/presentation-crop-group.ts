/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for two PowerPoint-parity presentation features:
 *  - IMAGE CROP: an optional crop rectangle (`cropRect`) on an image block,
 *    expressed as signed fractions of the SOURCE image. Free-crop gestures stay
 *    inside 0..1, while PowerPoint Fill/Fit can persist an outset kept region
 *    (negative x/y and w/h greater than 1) without flattening that native
 *    `a:srcRect` geometry.
 *  - PERSISTENT GROUPING: a shared `groupId` on blocks so selecting one selects
 *    the whole group and move/align operate on every member; ungroup clears it.
 *
 * Dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * reused by both the editor component and tests. Mirrors the conventions of the
 * sibling pure modules (`presentation-drag.ts`, `presentation-align.ts`): a
 * type guard / normalizer, integer-clean output where geometry-adjacent, small
 * composable functions, no mutation of inputs.
 *
 * Invariants:
 *  - {@link clampCropRect} always returns a rect whose x/y/w/h are finite, with
 *    x,y in [0,1), w,h in (0,1], and x+w ≤ 1, y+h ≤ 1 (a real visible region).
 *  - {@link normalizeSignedCropRect} accepts only finite, bounded kept regions
 *    that overlap the source. {@link cropToCss} renders those signed values
 *    directly and falls back to identity for malformed input.
 *  - {@link normalizeGroupId} yields a trimmed non-empty string or `undefined`,
 *    never the empty string, so an absent group is represented uniformly.
 */

/**
 * A crop rectangle in SOURCE-image fractions. `x`/`y` are the top-left origin
 * of the kept region and `w`/`h` are its width/height. Free crop uses the 0..1
 * subset; native Fit may use negative origins and dimensions greater than 1.
 * Identity is `{ x: 0, y: 0, w: 1, h: 1 }`.
 */
export interface CropRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface PresentationNativeGroupAuthority {
  schemaVersion: "pptx-native-group-v1";
  slidePart: string;
  groupShapeId: string;
  groupFingerprint: string;
  members: Array<{ shapeId: string; fingerprint: string }>;
  offX: number;
  offY: number;
  extX: number;
  extY: number;
  childOffX: number;
  childOffY: number;
  childExtX: number;
  childExtY: number;
  rotation: number;
  ungroupSafe: boolean;
}

const SHA256_RE = /^[0-9a-f]{64}$/;
const SHAPE_ID_RE = /^shp[1-9][0-9]*$/;
const NORMALIZED_PPTX_PART_RE = /^(?!\/)(?!.*(?:^|\/)\.\.?\/)[^\\\u0000]+$/;

/** Strictly parse the server-authenticated identity of one direct p:grpSp. */
export function parsePresentationNativeGroupAuthority(
  value: unknown,
): PresentationNativeGroupAuthority | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const candidate = value as Record<string, unknown>;
  const keys = [
    "schemaVersion",
    "slidePart",
    "groupShapeId",
    "groupFingerprint",
    "members",
    "offX",
    "offY",
    "extX",
    "extY",
    "childOffX",
    "childOffY",
    "childExtX",
    "childExtY",
    "rotation",
    "ungroupSafe",
  ];
  if (
    Object.keys(candidate).length !== keys.length ||
    keys.some((key) => !Object.hasOwn(candidate, key)) ||
    candidate.schemaVersion !== "pptx-native-group-v1" ||
    typeof candidate.slidePart !== "string" ||
    !NORMALIZED_PPTX_PART_RE.test(candidate.slidePart) ||
    typeof candidate.groupShapeId !== "string" ||
    !SHAPE_ID_RE.test(candidate.groupShapeId) ||
    typeof candidate.groupFingerprint !== "string" ||
    !SHA256_RE.test(candidate.groupFingerprint) ||
    typeof candidate.ungroupSafe !== "boolean" ||
    !Array.isArray(candidate.members) ||
    candidate.members.length < 2 ||
    candidate.members.length > 64
  ) {
    return null;
  }
  const numericKeys = [
    "offX",
    "offY",
    "extX",
    "extY",
    "childOffX",
    "childOffY",
    "childExtX",
    "childExtY",
    "rotation",
  ] as const;
  if (
    numericKeys.some(
      (key) =>
        typeof candidate[key] !== "number" ||
        !Number.isSafeInteger(candidate[key]),
    ) ||
    Number(candidate.extX) <= 0 ||
    Number(candidate.extY) <= 0 ||
    Number(candidate.childExtX) <= 0 ||
    Number(candidate.childExtY) <= 0 ||
    Math.abs(Number(candidate.rotation)) > 21_600_000
  ) {
    return null;
  }
  const members = candidate.members as unknown[];
  const memberIds = new Set<string>();
  for (const member of members) {
    if (!member || typeof member !== "object" || Array.isArray(member)) return null;
    const entry = member as Record<string, unknown>;
    if (
      Object.keys(entry).length !== 2 ||
      typeof entry.shapeId !== "string" ||
      !SHAPE_ID_RE.test(entry.shapeId) ||
      typeof entry.fingerprint !== "string" ||
      !SHA256_RE.test(entry.fingerprint) ||
      memberIds.has(entry.shapeId)
    ) {
      return null;
    }
    memberIds.add(entry.shapeId);
  }
  return candidate as unknown as PresentationNativeGroupAuthority;
}

export function blockNativeGroupAuthority(
  block: Record<string, unknown>,
): PresentationNativeGroupAuthority | null {
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return null;
  return parsePresentationNativeGroupAuthority(
    (source as Record<string, unknown>).nativeGroupAuthority,
  );
}








export interface PresentationNativeGroupMoveAuthority {
  schemaVersion: "pptx-native-group-move-v1";
  slidePart: string;
  groupShapeId: string;
  groupFingerprint: string;
  offX: number;
  offY: number;
}

const NATIVE_GROUP_MOVE_KEYS = [
  "schemaVersion",
  "slidePart",
  "groupShapeId",
  "groupFingerprint",
  "offX",
  "offY",
] as const;

/**
 * Strictly parse a block's `source.nativeGroupMove`, exactly as the save
 * planner reads it (`parse_presentation_native_group_move_authority`).
 */
export function blockNativeGroupMoveAuthority(
  block: Record<string, unknown>,
): PresentationNativeGroupMoveAuthority | null {
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return null;
  const value = (source as Record<string, unknown>).nativeGroupMove;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const candidate = value as Record<string, unknown>;
  if (
    Object.keys(candidate).length !== NATIVE_GROUP_MOVE_KEYS.length ||
    NATIVE_GROUP_MOVE_KEYS.some((key) => !Object.hasOwn(candidate, key)) ||
    candidate.schemaVersion !== "pptx-native-group-move-v1" ||
    typeof candidate.slidePart !== "string" ||
    !NORMALIZED_PPTX_PART_RE.test(candidate.slidePart) ||
    typeof candidate.groupShapeId !== "string" ||
    !SHAPE_ID_RE.test(candidate.groupShapeId) ||
    typeof candidate.groupFingerprint !== "string" ||
    !SHA256_RE.test(candidate.groupFingerprint) ||
    !Number.isSafeInteger(candidate.offX) ||
    !Number.isSafeInteger(candidate.offY)
  ) {
    return null;
  }
  return candidate as unknown as PresentationNativeGroupMoveAuthority;
}

/** Whether two move authorities name the same group at the same offset. */
export function sameNativeGroupMoveAuthority(
  a: PresentationNativeGroupMoveAuthority,
  b: PresentationNativeGroupMoveAuthority,
): boolean {
  return NATIVE_GROUP_MOVE_KEYS.every((key) => a[key] === b[key]);
}

/** The full-image identity crop — keeps the whole source image. */
export const IDENTITY_CROP: Readonly<CropRect> = Object.freeze({
  x: 0,
  y: 0,
  w: 1,
  h: 1,
});

/**
 * Smallest crop dimension (fraction). Keeps a cropped region grabbable and
 * avoids a divide-by-~0 blow-up in {@link cropToCss}'s scale math. ~1% of the
 * source on each axis.
 */
export const MIN_CROP_SIZE = 0.01;

/**
 * Signed persisted crop bounds. The 10,000:1 span covers the image preflight's
 * maximum useful aspect-ratio mismatch while keeping CSS/OOXML math bounded.
 * DrawingML percentages are quantized to 1/100000, so a smaller positive kept
 * span cannot be represented faithfully.
 */
export const MIN_SIGNED_CROP_SIZE = 0.00001;
export const MAX_SIGNED_CROP_MAGNITUDE = 10_000;
const DRAWINGML_CROP_PERCENT_SCALE = 100_000;

/** Office picture-crop aspect ratios (Word and PowerPoint desktop list the
 * same eleven), grouped as rendered in the gallery. */
export const PICTURE_CROP_ASPECT_RATIOS = Object.freeze([
  { id: "1:1", group: "square", width: 1, height: 1 },
  { id: "2:3", group: "portrait", width: 2, height: 3 },
  { id: "3:4", group: "portrait", width: 3, height: 4 },
  { id: "3:5", group: "portrait", width: 3, height: 5 },
  { id: "4:5", group: "portrait", width: 4, height: 5 },
  { id: "3:2", group: "landscape", width: 3, height: 2 },
  { id: "4:3", group: "landscape", width: 4, height: 3 },
  { id: "5:3", group: "landscape", width: 5, height: 3 },
  { id: "5:4", group: "landscape", width: 5, height: 4 },
  { id: "16:9", group: "landscape", width: 16, height: 9 },
  { id: "16:10", group: "landscape", width: 16, height: 10 },
] as const);

type PictureCropAspectRatioPreset = (typeof PICTURE_CROP_ASPECT_RATIOS)[number];
export type PictureCropAspectRatioId = PictureCropAspectRatioPreset["id"];

/** Resolve only an exact catalog id; arbitrary or malformed ratios fail shut. */
function pictureCropAspectRatioPreset(
  value: unknown,
): PictureCropAspectRatioPreset | null {
  if (typeof value !== "string") return null;
  return (
    PICTURE_CROP_ASPECT_RATIOS.find((preset) => preset.id === value) ?? null
  );
}

/** Round to a clean fraction (6 dp) guarding NaN/Infinity → fallback. Collapses
 * a signed-zero (`-0`) to `0` so stringified CSS is stable. */
function round6(value: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback;
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return rounded === 0 ? 0 : rounded;
}

/** Clamp a number into `[min, max]`. Returns `min` when `min > max`. */
function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(max, Math.max(min, value));
}

/** Type guard: is `value` a crop-rect-shaped object with finite numeric sides? */
export function isCropRect(value: unknown): value is CropRect {
  if (!value || typeof value !== "object") return false;
  const r = value as Record<string, unknown>;
  return (
    typeof r.x === "number" &&
    typeof r.y === "number" &&
    typeof r.w === "number" &&
    typeof r.h === "number"
  );
}

/**
 * Validate + clamp a raw crop-rect-ish value into a guaranteed-visible region.
 *
 * Rules (applied in order):
 *  - non-finite / missing fields fall back to the identity crop's value,
 *  - width/height are floored at {@link MIN_CROP_SIZE} and capped at 1,
 *  - origin x/y are clamped to `[0, 1 - size]` so the region never spills past
 *    the source edge (x + w ≤ 1, y + h ≤ 1),
 *  - output is rounded to 6 dp so persisted crops stay clean.
 *
 * A `null`/`undefined`/garbage input returns the {@link IDENTITY_CROP} clone.
 */
export function clampCropRect(rect: unknown): CropRect {
  const r = (rect ?? {}) as Record<string, unknown>;
  const w = clamp(
    round6(typeof r.w === "number" ? r.w : Number.NaN, 1),
    MIN_CROP_SIZE,
    1,
  );
  const h = clamp(
    round6(typeof r.h === "number" ? r.h : Number.NaN, 1),
    MIN_CROP_SIZE,
    1,
  );
  const x = clamp(
    round6(typeof r.x === "number" ? r.x : Number.NaN, 0),
    0,
    1 - w,
  );
  const y = clamp(
    round6(typeof r.y === "number" ? r.y : Number.NaN, 0),
    0,
    1 - h,
  );
  return {
    x: round6(x, 0),
    y: round6(y, 0),
    w: round6(w, 1),
    h: round6(h, 1),
  };
}

/**
 * Validate a persisted signed keep rectangle without clamping away native Fit
 * outsets. A valid rectangle must be finite, bounded, positive-sized, and
 * overlap the source on both axes. Returns `null` for malformed/extreme input.
 */
export function normalizeSignedCropRect(rect: unknown): CropRect | null {
  if (!isCropRect(rect)) return null;
  const x = round6(rect.x, Number.NaN);
  const y = round6(rect.y, Number.NaN);
  const w = round6(rect.w, Number.NaN);
  const h = round6(rect.h, Number.NaN);
  if (![x, y, w, h].every(Number.isFinite)) return null;
  if (
    Math.abs(x) > MAX_SIGNED_CROP_MAGNITUDE ||
    Math.abs(y) > MAX_SIGNED_CROP_MAGNITUDE ||
    w < MIN_SIGNED_CROP_SIZE ||
    h < MIN_SIGNED_CROP_SIZE ||
    w > MAX_SIGNED_CROP_MAGNITUDE ||
    h > MAX_SIGNED_CROP_MAGNITUDE
  ) {
    return null;
  }
  // The kept rectangle must intersect the source [0,1] on both axes.
  if (x >= 1 || y >= 1 || x + w <= 0 || y + h <= 0) return null;
  return { x, y, w, h };
}

/** Is `rect` (after signed validation) the full-image identity crop (a no-op)? */
export function isIdentityCrop(rect: unknown): boolean {
  const c = normalizeSignedCropRect(rect);
  if (!c) return true;
  return c.x === 0 && c.y === 0 && c.w === 1 && c.h === 1;
}

/** Compare two optional persisted crops, treating omission as identity. */
export function sameCropRect(left: unknown, right: unknown): boolean {
  const a = normalizeSignedCropRect(left) ?? IDENTITY_CROP;
  const b = normalizeSignedCropRect(right) ?? IDENTITY_CROP;
  return a.x === b.x && a.y === b.y && a.w === b.w && a.h === b.h;
}

function positiveDimension(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : null;
}

/**
 * Canonicalize a signed keep rectangle onto DrawingML's integer-percentage
 * lattice. PowerPoint persists the four inset edges (`l/t/r/b`), not `w/h`, so
 * deriving width and height back from those quantized edges makes the editor's
 * immediate state identical to a native PPTX save/reopen.
 */
export function canonicalizeDrawingMlCropRect(
  rect: unknown,
): CropRect | null {
  const normalized = normalizeSignedCropRect(rect);
  if (!normalized) return null;
  const left = Math.round(normalized.x * DRAWINGML_CROP_PERCENT_SCALE);
  const top = Math.round(normalized.y * DRAWINGML_CROP_PERCENT_SCALE);
  const right = Math.round(
    (1 - normalized.x - normalized.w) * DRAWINGML_CROP_PERCENT_SCALE,
  );
  const bottom = Math.round(
    (1 - normalized.y - normalized.h) * DRAWINGML_CROP_PERCENT_SCALE,
  );
  return normalizeSignedCropRect({
    x: left / DRAWINGML_CROP_PERCENT_SCALE,
    y: top / DRAWINGML_CROP_PERCENT_SCALE,
    w: (DRAWINGML_CROP_PERCENT_SCALE - left - right) /
      DRAWINGML_CROP_PERCENT_SCALE,
    h: (DRAWINGML_CROP_PERCENT_SCALE - top - bottom) /
      DRAWINGML_CROP_PERCENT_SCALE,
  });
}

/**
 * PowerPoint Picture Crop → Fill. The kept source sub-region has the frame's
 * aspect ratio and is centred on the source.
 */
export function pictureFillCropRect(
  sourceWidth: unknown,
  sourceHeight: unknown,
  frameWidth: unknown,
  frameHeight: unknown,
): CropRect | null {
  const sw = positiveDimension(sourceWidth);
  const sh = positiveDimension(sourceHeight);
  const fw = positiveDimension(frameWidth);
  const fh = positiveDimension(frameHeight);
  if (sw === null || sh === null || fw === null || fh === null) return null;
  const sourceAspect = sw / sh;
  const frameAspect = fw / fh;
  const raw =
    sourceAspect > frameAspect
      ? {
          x: (1 - frameAspect / sourceAspect) / 2,
          y: 0,
          w: frameAspect / sourceAspect,
          h: 1,
        }
      : sourceAspect < frameAspect
        ? {
            x: 0,
            y: (1 - sourceAspect / frameAspect) / 2,
            w: 1,
            h: sourceAspect / frameAspect,
          }
        : IDENTITY_CROP;
  return normalizeSignedCropRect(raw);
}

/**
 * PowerPoint Picture Crop → Fit. The full source is centred inside the frame,
 * represented as an outset kept rectangle. This deliberately preserves
 * negative x/y and w/h greater than 1 for native `a:srcRect`.
 */
export function pictureFitCropRect(
  sourceWidth: unknown,
  sourceHeight: unknown,
  frameWidth: unknown,
  frameHeight: unknown,
): CropRect | null {
  const sw = positiveDimension(sourceWidth);
  const sh = positiveDimension(sourceHeight);
  const fw = positiveDimension(frameWidth);
  const fh = positiveDimension(frameHeight);
  if (sw === null || sh === null || fw === null || fh === null) return null;
  const sourceAspect = sw / sh;
  const frameAspect = fw / fh;
  const raw =
    sourceAspect > frameAspect
      ? {
          x: 0,
          y: (1 - sourceAspect / frameAspect) / 2,
          w: 1,
          h: sourceAspect / frameAspect,
        }
      : sourceAspect < frameAspect
        ? {
            x: (1 - frameAspect / sourceAspect) / 2,
            y: 0,
            w: frameAspect / sourceAspect,
            h: 1,
          }
        : IDENTITY_CROP;
  return normalizeSignedCropRect(raw);
}

/**
 * Crop the current visible source region to a catalog aspect ratio while
 * preserving its centre. Only the source keep rectangle changes: the picture
 * frame geometry remains untouched. The calculation deliberately starts from
 * the current signed crop, so authenticated native Fit/outset rectangles can
 * move directly into an Aspect Ratio preset without first being flattened.
 */
export function pictureAspectRatioCropRect(
  currentCrop: unknown,
  sourceWidth: unknown,
  sourceHeight: unknown,
  ratioId: unknown,
): CropRect | null {
  const current =
    currentCrop === null || currentCrop === undefined
      ? IDENTITY_CROP
      : normalizeSignedCropRect(currentCrop);
  const sw = positiveDimension(sourceWidth);
  const sh = positiveDimension(sourceHeight);
  const preset = pictureCropAspectRatioPreset(ratioId);
  if (!current || sw === null || sh === null || !preset) return null;

  const canonicalCurrent = canonicalizeDrawingMlCropRect(current);
  if (!canonicalCurrent) return null;

  const targetAspect = preset.width / preset.height;
  const currentAspect = (canonicalCurrent.w * sw) / (canonicalCurrent.h * sh);
  const centerX = canonicalCurrent.x + canonicalCurrent.w / 2;
  const centerY = canonicalCurrent.y + canonicalCurrent.h / 2;
  const heightPinnedWidth = (canonicalCurrent.h * sh * targetAspect) / sw;
  const heightPinnedCandidate = canonicalizeDrawingMlCropRect({
    x: centerX - heightPinnedWidth / 2,
    y: canonicalCurrent.y,
    w: heightPinnedWidth,
    h: canonicalCurrent.h,
  });
  const widthPinnedHeight = (canonicalCurrent.w * sw) / (targetAspect * sh);
  const widthPinnedCandidate = canonicalizeDrawingMlCropRect({
    x: canonicalCurrent.x,
    y: centerY - widthPinnedHeight / 2,
    w: canonicalCurrent.w,
    h: widthPinnedHeight,
  });
  if (!heightPinnedCandidate || !widthPinnedCandidate) return null;

  // A target aspect generally cannot be represented exactly on DrawingML's
  // 1/100000 edge lattice. Recognize either axis-pinned projection as the same
  // canonical fixed point before choosing an axis again; otherwise a rounded
  // result can straddle the ideal ratio and shrink the opposite axis on the
  // second click, creating a spurious save/Undo step.
  if (
    sameCropRect(canonicalCurrent, heightPinnedCandidate) ||
    sameCropRect(canonicalCurrent, widthPinnedCandidate)
  ) {
    return canonicalCurrent;
  }
  return currentAspect > targetAspect
    ? heightPinnedCandidate
    : widthPinnedCandidate;
}

/** The picture frame and crop that Crop → Aspect Ratio commits together. */
export interface PictureAspectRatioCropGeometry {
  cropRect: CropRect;
  left?: number;
  top?: number;
  width: number;
  height: number;
}

/**
 * PowerPoint Picture Crop → Aspect Ratio, applied to the picture AS DRAWN. The
 * crop frame takes the chosen ratio around the current frame's centre while
 * the picture underneath keeps its drawn size on both axes (its own
 * proportions, and any stretch it already carried), so nothing is distorted:
 * the kept source span shrinks on one axis and the frame follows it.
 *
 * The crop is computed in display space, where the full source as currently
 * drawn spans `width / w` by `height / h`; the source's pixel size is not
 * needed. Frame values are rounded to whole units (EMU) so the native
 * geometry patch can carry them, and the frame's centre is kept by moving
 * `left` / `top` by half the change. A picture already at the ratio (on the
 * DrawingML edge lattice) returns its geometry unchanged.
 */
export function pictureAspectRatioCropGeometry(
  frame: {
    readonly cropRect?: unknown;
    readonly left?: unknown;
    readonly top?: unknown;
    readonly width?: unknown;
    readonly height?: unknown;
  },
  ratioId: unknown,
): PictureAspectRatioCropGeometry | null {
  const current =
    frame.cropRect === null || frame.cropRect === undefined
      ? IDENTITY_CROP
      : normalizeSignedCropRect(frame.cropRect);
  const fw = positiveDimension(frame.width);
  const fh = positiveDimension(frame.height);
  if (!current || fw === null || fh === null) return null;
  const crop = pictureAspectRatioCropRect(
    current,
    fw / current.w,
    fh / current.h,
    ratioId,
  );
  if (!crop) return null;
  const offset = (value: unknown, delta: number): number | undefined =>
    typeof value === "number" && Number.isFinite(value)
      ? Math.round(value + delta)
      : undefined;
  if (sameCropRect(current, crop)) {
    return {
      cropRect: crop,
      left: offset(frame.left, 0),
      top: offset(frame.top, 0),
      width: fw,
      height: fh,
    };
  }
  // The drawn scale on each axis is unchanged, so the frame scales with the
  // kept span on that axis.
  const width = Math.round((fw * crop.w) / current.w);
  const height = Math.round((fh * crop.h) / current.h);
  if (width <= 0 || height <= 0) return null;
  return {
    cropRect: crop,
    left: offset(frame.left, (fw - width) / 2),
    top: offset(frame.top, (fh - height) / 2),
    width,
    height,
  };
}

/**
 * A minimal subset of the inline-style props {@link cropToCss} emits. Kept
 * framework-agnostic (plain string CSS values) so it can be spread onto a React
 * `style` object or serialized in a test without importing React types.
 */
export interface CropCss {
  /** The kept region scaled to fill the block box. */
  width: string;
  height: string;
  /** Offset so the kept region's top-left lands at the block's top-left. */
  transform: string;
  transformOrigin: "top left";
  /** Stretch the source to the block box before the crop scale is applied. */
  objectFit: "fill";
  /** Always hide the spilled-over (non-kept) part of the scaled image. */
  position: "absolute";
  top: 0;
  left: 0;
}

/**
 * Render a crop deterministically as inline `<img>` styles, to be used inside a
 * wrapper with `overflow: hidden`. The kept sub-region (after {@link
 * normalizeSignedCropRect}) is scaled to the wrapper and translated so its
 * top-left corner aligns with the wrapper's top-left, with the rest clipped by
 * the wrapper's overflow.
 *
 * Math: to show a fraction `w` of the source across the full wrapper width, the
 * image is scaled to `100 / w` percent. CSS transform percentages are relative
 * to the transformed image itself, so shifting it by `-x * 100%` moves source
 * coordinate `x` to the wrapper's left edge. Dividing x by w again would apply
 * the scale twice. Symmetric on the vertical axis.
 *
 * Malformed input falls back to identity. Valid signed Fit values are not
 * clamped, so source padding survives reload and rendering.
 */
export function cropToCss(rect: unknown): CropCss {
  const c = normalizeSignedCropRect(rect) ?? IDENTITY_CROP;
  const scaleX = 100 / c.w;
  const scaleY = 100 / c.h;
  // Transform percentages are relative to the scaled image element.
  const offsetX = -c.x * 100;
  const offsetY = -c.y * 100;
  return {
    width: `${round6(scaleX, 100)}%`,
    height: `${round6(scaleY, 100)}%`,
    transform: `translate(${round6(offsetX, 0)}%, ${round6(offsetY, 0)}%)`,
    transformOrigin: "top left",
    objectFit: "fill",
    position: "absolute",
    top: 0,
    left: 0,
  };
}

/**
 * Adjust a crop by dragging one inset edge handle inward/outward by a fraction
 * delta, keeping the OPPOSITE edge pinned (PowerPoint crop-handle behaviour).
 * `edge` is one of the four sides; `delta` is a signed fraction of the source on
 * that handle's axis.
 *
 * The crop is the same SIGNED keep rectangle Fill / Fit and native `a:srcRect`
 * use: a handle dragged outward past the source adds blank margin (x < 0, or
 * x + w > 1), exactly as PowerPoint's crop handles do, and a Fit outset is not
 * flattened by touching one handle. Only the touched edge moves; the other
 * three keep their exact values, so a zero delta is a no-op. The moving edge
 * is bounded so the kept region stays at least {@link MIN_CROP_SIZE} wide /
 * tall, keeps overlapping the source, and stays inside the persisted bounds.
 *
 * - `left`  moves the left edge:  x += delta, w -= delta (right edge pinned).
 * - `right` moves the right edge: w += delta (left edge pinned).
 * - `top`   moves the top edge:   y += delta, h -= delta (bottom edge pinned).
 * - `bottom` moves the bottom edge: h += delta (top edge pinned).
 */
export const CROP_EDGES = ["left", "right", "top", "bottom"] as const;

export type CropEdge = (typeof CROP_EDGES)[number];

const CROP_EDGE_SET: ReadonlySet<string> = new Set(CROP_EDGES);

/** Type guard: is `value` exactly one of the four crop edges? */
export function isCropEdge(value: unknown): value is CropEdge {
  return typeof value === "string" && CROP_EDGE_SET.has(value);
}

export function adjustCropEdge(
  rect: unknown,
  edge: CropEdge,
  delta: number,
): CropRect {
  // A valid signed crop (a Fit outset included) is adjusted as it is; only a
  // malformed one falls back to the clamped identity-ish region.
  const c = normalizeSignedCropRect(rect) ?? clampCropRect(rect);
  const d = Number.isFinite(delta) ? delta : 0;
  if (!isCropEdge(edge) || d === 0) return c;
  // Derive the moving edge from the PINNED opposite edge. Passing only a size
  // delta and re-clamping x/y silently translated the pinned edge when a
  // handle was dragged OUTWARD past the source (x-clamp pulled the origin),
  // corrupting the crop region on save. Compute the moving boundary against
  // the fixed one so the opposite edge never moves, and never rewrite the
  // other axis.
  const MAX = MAX_SIGNED_CROP_MAGNITUDE;
  let next: CropRect;
  switch (edge) {
    case "left": {
      const right = c.x + c.w; // pinned
      // The kept region must still reach into the source (x < 1) and keep its
      // minimum span against the pinned right edge.
      const x = clamp(
        c.x + d,
        Math.max(-MAX, right - MAX),
        Math.min(right - MIN_CROP_SIZE, 1 - MIN_CROP_SIZE),
      );
      next = { ...c, x, w: right - x };
      break;
    }
    case "right": {
      // Pin the left edge; the right boundary must stay past the source's left
      // edge (x + w > 0) and the minimum span past the pinned left edge.
      const right = clamp(
        c.x + c.w + d,
        Math.max(c.x + MIN_CROP_SIZE, MIN_CROP_SIZE),
        c.x + MAX,
      );
      next = { ...c, w: right - c.x };
      break;
    }
    case "top": {
      const bottom = c.y + c.h; // pinned
      const y = clamp(
        c.y + d,
        Math.max(-MAX, bottom - MAX),
        Math.min(bottom - MIN_CROP_SIZE, 1 - MIN_CROP_SIZE),
      );
      next = { ...c, y, h: bottom - y };
      break;
    }
    case "bottom": {
      // Pin the top edge.
      const bottom = clamp(
        c.y + c.h + d,
        Math.max(c.y + MIN_CROP_SIZE, MIN_CROP_SIZE),
        c.y + MAX,
      );
      next = { ...c, h: bottom - c.y };
      break;
    }
    default:
      return c;
  }
  return normalizeSignedCropRect(next) ?? c;
}

/**
 * PowerPoint's crop has a handle at each corner as well as each side ("drag a
 * corner cropping handle to crop two adjacent sides at once"); a corner is
 * its two edges, each pinned against its own opposite edge.
 */
export const CROP_CORNERS = ["top-left", "top-right", "bottom-left", "bottom-right"] as const;
export type CropCorner = (typeof CROP_CORNERS)[number];

export function cropCornerEdges(corner: CropCorner): readonly [CropEdge, CropEdge] {
  const [vertical, horizontal] = corner.split("-") as [CropEdge, CropEdge];
  return [horizontal, vertical];
}

/** Move a corner: `dx` on its side edge, `dy` on its top/bottom edge. */
export function adjustCropCorner(
  rect: unknown,
  corner: CropCorner,
  dx: number,
  dy: number,
): CropRect {
  const [horizontal, vertical] = cropCornerEdges(corner);
  return adjustCropEdge(adjustCropEdge(rect, horizontal, dx), vertical, dy);
}

/**
 * Drag the picture inside the crop frame (PowerPoint: "to change the crop
 * position, drag the picture"): the frame stays, the picture moves under it,
 * so the kept region slides the other way across the source by the same
 * fraction. Size never changes; the region stays within the source (or, for
 * a Fit outset larger than the source, keeps covering it).
 */
export function panCropRect(rect: unknown, dx: number, dy: number): CropRect {
  const c = normalizeSignedCropRect(rect) ?? clampCropRect(rect);
  const slide = (start: number, span: number, delta: number) => {
    if (!Number.isFinite(delta) || delta === 0) return start;
    const low = Math.min(0, 1 - span, start);
    const high = Math.max(0, 1 - span, start);
    return clamp(start - delta, low, high);
  };
  const next = { ...c, x: slide(c.x, c.w, dx), y: slide(c.y, c.h, dy) };
  return normalizeSignedCropRect(next) ?? c;
}

// ---------------------------------------------------------------------------
// Grouping
// ---------------------------------------------------------------------------

/**
 * A block-ish value carrying an optional id, its group and the groups nested
 * inside that one.
 *
 * `groupId` is the block's OUTERMOST group: the unit a click selects, a drag
 * moves and the Selection Pane lists, exactly as before groups could nest.
 * `groupPath` names the groups inside it that the block also belongs to, from
 * the one just inside the outermost down to the innermost (DrawingML nests a
 * `p:grpSp` in a `p:grpSp`, ECMA-376 19.3.1.22). A block in a single group
 * has no `groupPath`, so every deck stored before nesting reads as it did.
 */
export interface GroupableBlock {
  id: string;
  groupId?: unknown;
  groupPath?: unknown;
  [key: string]: unknown;
}

/** How deep groups nest inside one another (the path holds one fewer). */
export const MAX_GROUP_NESTING = 8;

/**
 * Normalize a raw `groupId` into a trimmed non-empty string, or `undefined`
 * when it is absent / blank / non-string. This keeps "no group" represented
 * uniformly as `undefined` (never an empty string), so an absent group never
 * accidentally matches another absent group.
 */
export function normalizeGroupId(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/** The normalized groupId of a block, or `undefined` when ungrouped. */
export function blockGroupId(block: GroupableBlock | null | undefined): string | undefined {
  return normalizeGroupId(block?.groupId);
}

/**
 * The groups nested inside a block's outermost group that the block belongs
 * to, outermost first: trimmed, distinct, never the outermost id itself, and
 * bounded. Anything malformed reads as no nesting (the block is then simply a
 * member of its outermost group), so a bad path can thin a hierarchy but never
 * break a group.
 */
export function blockGroupPath(block: GroupableBlock | null | undefined): string[] {
  const outer = blockGroupId(block);
  const raw = block?.groupPath;
  if (!outer || !Array.isArray(raw)) return [];
  const path: string[] = [];
  for (const entry of raw) {
    const id = normalizeGroupId(entry);
    if (!id || id === outer || path.includes(id)) return [];
    path.push(id);
  }
  return path.length < MAX_GROUP_NESTING ? path : [];
}

/** `block` with this group and nesting, the `groupPath` key absent when empty. */
function withGroup<T extends GroupableBlock>(
  block: T,
  groupId: string | undefined,
  path: readonly string[],
): T {
  const next = { ...block } as Record<string, unknown>;
  if (groupId) next.groupId = groupId;
  else delete next.groupId;
  if (groupId && path.length) next.groupPath = [...path];
  else delete next.groupPath;
  return next as T;
}

/**
 * Every block sharing the given (normalized) groupId, in input order. A blank /
 * undefined `groupId` argument matches NOTHING (returns []), so "no group" can
 * never select every ungrouped block. Order is preserved.
 */
export function blocksInGroup<T extends GroupableBlock>(
  blocks: ReadonlyArray<T>,
  groupId: unknown,
): T[] {
  const target = normalizeGroupId(groupId);
  if (!target) return [];
  return blocks.filter((block) => normalizeGroupId(block.groupId) === target);
}

/**
 * The set of block ids that should be selected when `blockId` is clicked: if
 * that block belongs to a group, the whole group's ids; otherwise just the
 * clicked id. Unknown ids return an empty set. Used so selecting one grouped
 * block selects the group (move/align then operate on all members).
 */
export function expandSelectionForGroup<T extends GroupableBlock>(
  blocks: ReadonlyArray<T>,
  blockId: string,
): Set<string> {
  const block = blocks.find((b) => b.id === blockId);
  if (!block) return new Set();
  const groupId = normalizeGroupId(block.groupId);
  if (!groupId) return new Set([blockId]);
  return new Set(blocksInGroup(blocks, groupId).map((b) => b.id));
}

/**
 * Generate a fresh, collision-free group id. Prefers `crypto.randomUUID` when
 * available (browser / modern Node), falling back to a timestamp+random token
 * so the helper stays usable in any environment (and unit tests). The `grp-`
 * prefix keeps grouped blocks recognizable in persisted JSON.
 */
export function newGroupId(): string {
  const g = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
  if (g?.randomUUID) return `grp-${g.randomUUID()}`;
  return `grp-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Assign a shared `groupId` to the blocks whose ids are in `ids`, returning a
 * NEW array (inputs untouched). Blocks not in `ids` are returned unchanged.
 * When `groupId` is omitted a fresh id is generated; an explicit blank groupId
 * is treated as ungroup (the field is cleared) for the targeted blocks.
 *
 * Idempotent-friendly: re-assigning a block already in the group is a no-op for
 * that block's groupId value.
 */
export function assignGroup<T extends GroupableBlock>(
  blocks: ReadonlyArray<T>,
  ids: ReadonlyArray<string> | ReadonlySet<string>,
  groupId?: string,
): T[] {
  const idSet = ids instanceof Set ? ids : new Set(ids);
  const normalized = groupId === undefined ? newGroupId() : normalizeGroupId(groupId);
  if (!normalized) {
    // Explicit blank groupId → clear the field (ungroup this block outright).
    return blocks.map((block) => {
      if (!idSet.has(block.id)) return block;
      if (normalizeGroupId(block.groupId) === undefined) return block;
      return withGroup(block, undefined, []);
    });
  }
  // A group taken WHOLE into the new group stays a group inside it
  // (PowerPoint nests it; Ungroup then removes the new group alone,
  // PPT-B204). A group only partly taken is not a group any more for the
  // members that leave it: they join the new group as plain members.
  const wholeGroups = new Set<string>();
  for (const block of blocks) {
    const current = blockGroupId(block);
    if (!current || current === normalized || !idSet.has(block.id)) continue;
    if (blocksInGroup(blocks, current).every((member) => idSet.has(member.id))) {
      wholeGroups.add(current);
    }
  }
  const planned = blocks.map((block) => {
    if (!idSet.has(block.id)) return block;
    const current = blockGroupId(block);
    if (current === normalized) return block;
    const path = current && wholeGroups.has(current) ? [current, ...blockGroupPath(block)] : [];
    return withGroup(block, normalized, path);
  });
  // Deeper than groups may nest: refused whole, the blocks as they were.
  if (planned.some((block) => Array.isArray(block.groupPath) && block.groupPath.length >= MAX_GROUP_NESTING)) {
    return [...blocks];
  }
  return planned;
}

/**
 * Stack the blocks in `ids` together where the frontmost of them was, keeping
 * their own order, and every other block in its order. PowerPoint's Group
 * makes one object of the selection at the frontmost member's layer, so a
 * shape that sat between two grouped shapes ends up behind the group. Grouping
 * A and C around B used to leave the group split around B, which the
 * Selection Pane cannot list. Returns the input array when already contiguous.
 */
export function stackGroupAtFrontmostMember<T extends GroupableBlock>(
  blocks: ReadonlyArray<T>,
  ids: ReadonlyArray<string> | ReadonlySet<string>,
): T[] {
  const idSet = ids instanceof Set ? ids : new Set(ids);
  const members = blocks.filter((block) => idSet.has(block.id));
  let last = -1;
  blocks.forEach((block, index) => {
    if (idSet.has(block.id)) last = index;
  });
  if (members.length < 2 || last - members.length + 1 === blocks.indexOf(members[0])) {
    return [...blocks];
  }
  const next: T[] = [];
  blocks.forEach((block, index) => {
    if (index === last) next.push(...members);
    else if (!idSet.has(block.id)) next.push(block);
  });
  return next;
}


















export function ungroup<T extends GroupableBlock>(
  blocks: ReadonlyArray<T>,
  ids: ReadonlyArray<string> | ReadonlySet<string>,
): T[] {
  const idSet = ids instanceof Set ? ids : new Set(ids);
  return blocks.map((block) => {
    if (!idSet.has(block.id)) return block;
    if (normalizeGroupId(block.groupId) === undefined) return block;
    const [inner, ...rest] = blockGroupPath(block);
    const next = withGroup(block, inner, rest) as Record<string, unknown>;
    delete next.groupAlt;
    delete next.groupDecorative;
    return next as T;
  });
}

/**
 * The set of distinct (normalized) group ids present across `blocks`, in first-
 * seen order. Useful for diagnostics / "select all groups" affordances and for
 * tests asserting grouping integrity.
 */
export function groupIds(blocks: ReadonlyArray<GroupableBlock>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const block of blocks) {
    const id = normalizeGroupId(block.groupId);
    if (id && !seen.has(id)) {
      seen.add(id);
      out.push(id);
    }
  }
  return out;
}

/**
 * The group a selection is, whole: the one groupId every selected block
 * carries, when the selection holds every member of that group. Null for an
 * ungrouped or mixed selection, or part of a group.
 */
export function wholeGroupSelection<T extends GroupableBlock & { id: string }>(
  slideBlocks: ReadonlyArray<T>,
  selected: ReadonlyArray<T>,
): string | null {
  if (selected.length < 2) return null;
  const groupId = blockGroupId(selected[0]);
  if (!groupId || selected.some((block) => blockGroupId(block) !== groupId)) return null;
  const members = blocksInGroup(slideBlocks, groupId);
  const chosen = new Set(selected.map((block) => block.id));
  return members.length === selected.length && members.every((block) => chosen.has(block.id))
    ? groupId
    : null;
}

export type GroupMemberBox = Readonly<{
  id: string;
  left: number;
  top: number;
  width: number;
  height: number;
}>;

/** A group's frame: the box around its members. */
export function groupFrame(members: ReadonlyArray<GroupMemberBox>): Omit<GroupMemberBox, "id"> | null {
  if (members.length === 0) return null;
  const left = Math.min(...members.map((box) => box.left));
  const top = Math.min(...members.map((box) => box.top));
  const right = Math.max(...members.map((box) => box.left + box.width));
  const bottom = Math.max(...members.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Size > Width / Height on a group: the group's frame takes the value and every
 * member is scaled with it about the frame's top-left corner, so the members
 * keep their places and the gaps between them (PowerPoint sizes a group as one
 * object). With `keepRatio` the other axis follows. Null when the frame has no
 * extent on that axis or a member would fall below 1 EMU.
 */
export function resizedGroupMembers(
  members: ReadonlyArray<GroupMemberBox>,
  axis: "width" | "height",
  target: number,
  keepRatio: boolean,
): GroupMemberBox[] | null {
  const frame = groupFrame(members);
  if (!frame || frame[axis] <= 0 || !(target > 0)) return null;
  const factor = target / frame[axis];
  const fx = axis === "width" || keepRatio ? factor : 1;
  const fy = axis === "height" || keepRatio ? factor : 1;
  const scaled = members.map((box) => ({
    id: box.id,
    left: Math.round(frame.left + (box.left - frame.left) * fx),
    top: Math.round(frame.top + (box.top - frame.top) * fy),
    width: Math.round(box.width * fx),
    height: Math.round(box.height * fy),
  }));
  return scaled.every((box) => box.width >= 1 && box.height >= 1) ? scaled : null;
}
