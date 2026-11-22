/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Effects > Shadow: the page value, PowerPoint's gallery and the canvas
 * drawing, shared by Picture / Video / Graphics Effects
 * (`pictureEffects.shadow`, `lib/office-picture-effects`) and a table's
 * Effects (`tableEffects.shadow`, `lib/presentation-table-effects`).
 *
 * PowerPoint's Shadow submenu is "No Shadow" over three galleries: Outer
 * (nine Offset variations), Inner (nine Inside variations) and Perspective
 * (five). In the file a shadow is one `a:outerShdw` (ECMA-376 §20.1.8.45:
 * blur, distance, direction, scale `sx`/`sy`, skew `kx`/`ky` about the
 * alignment corner `algn`) or `a:innerShdw` (§20.1.8.40: blur, distance,
 * direction); a perspective shadow is a scaled and skewed outer one. The
 * reader, writer and in-place patch are the server's
 * `app/services/office_shadow_effects.py`.
 *
 * The page value: `blurRad`, `dist` (EMU), `dir` (1/60000 degree, the
 * direction the shadow is offset), `color` (`#rrggbb`), `alpha` (1/1000 %
 * opacity); an outer shadow adds `algn` and optional `sx` `sy` (1/1000 %) and
 * `kx` `ky` (1/60000 degree); an inner shadow is marked `inner: true`.
 *
 * Every gallery value is PowerPoint's own: PowerPoint for Mac 16.86 opened a
 * deck carrying candidate parameters and classified each one through its
 * scripting model as `msoShadow21`..`msoShadow43` -- numbered in the
 * gallery's order (Microsoft Learn, "MsoShadowType enumeration") -- while
 * off-by-one controls classified as none
 * (`MRAI-corpus/ppt-audit/L15a/fx/oracle/probe1_out.txt`). Names and order
 * are PowerPoint's string table (`mso_intl/strman_f401` 0x23-0x3c).
 */

import type { CSSProperties } from "react";

export const OFFICE_SHADOW_ALIGNMENTS = ["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"] as const;
export type OfficeShadowAlignment = (typeof OFFICE_SHADOW_ALIGNMENTS)[number];

export interface OfficeShadowEffect {
  /** `a:innerShdw`; absent = `a:outerShdw`. */
  readonly inner?: true;
  /** EMU. */
  readonly blurRad: number;
  /** EMU. */
  readonly dist: number;
  /** 1/60000 degree, clockwise from the x axis: where the shadow is offset. */
  readonly dir: number;
  /** Outer only (required there): the corner the scale and skew hold. */
  readonly algn?: OfficeShadowAlignment;
  /** `#rrggbb`. */
  readonly color: string;
  /** Opacity, DrawingML `a:alpha` (100000 = opaque). */
  readonly alpha: number;
  /** Outer only: scale, 1/1000 % (negative flips); absent = 100 %. */
  readonly sx?: number;
  readonly sy?: number;
  /** Outer only: skew, 1/60000 degree; absent = 0. */
  readonly kx?: number;
  readonly ky?: number;
}

export type OfficeShadowGroup = "outer" | "inner" | "perspective";

const EMU_PER_PT = 12_700;
const EMU_PER_PX = 9_525;
const MAX_BLUR = 100 * EMU_PER_PT;
const MAX_DISTANCE = 200 * EMU_PER_PT;
const ANGLE_LIMIT = 21_600_000;
const MAX_SCALE = 200_000;
const MAX_SKEW = 5_399_999;

const BLACK = "#000000";
const outer = (dir: number, algn: OfficeShadowAlignment): OfficeShadowEffect => ({
  blurRad: 50_800,
  dist: 38_100,
  dir,
  algn,
  color: BLACK,
  alpha: 40_000,
});
const inner = (dir: number): OfficeShadowEffect => ({
  inner: true,
  blurRad: 63_500,
  dist: 50_800,
  dir,
  color: BLACK,
  alpha: 50_000,
});

/** The Shadow galleries, PowerPoint's order (three rows of three; one row of five). */
export const OFFICE_SHADOW_GALLERY: readonly {
  readonly group: OfficeShadowGroup;
  readonly columns: number;
  readonly variations: readonly { readonly id: string; readonly value: OfficeShadowEffect }[];
}[] = [
  {
    group: "outer",
    columns: 3,
    variations: [
      { id: "bottomRight", value: outer(2_700_000, "tl") },
      { id: "bottom", value: outer(5_400_000, "t") },
      { id: "bottomLeft", value: outer(8_100_000, "tr") },
      { id: "right", value: outer(0, "l") },
      {
        id: "center",
        value: { blurRad: 63_500, dist: 0, dir: 0, algn: "ctr", color: BLACK, alpha: 40_000, sx: 102_000, sy: 102_000 },
      },
      { id: "left", value: outer(10_800_000, "r") },
      { id: "topRight", value: outer(18_900_000, "bl") },
      { id: "top", value: outer(16_200_000, "b") },
      { id: "topLeft", value: outer(13_500_000, "br") },
    ],
  },
  {
    group: "inner",
    columns: 3,
    variations: [
      { id: "topLeft", value: inner(13_500_000) },
      { id: "top", value: inner(16_200_000) },
      { id: "topRight", value: inner(18_900_000) },
      { id: "left", value: inner(10_800_000) },
      { id: "center", value: { inner: true, blurRad: 114_300, dist: 0, dir: 0, color: BLACK, alpha: 100_000 } },
      { id: "right", value: inner(0) },
      { id: "bottomLeft", value: inner(8_100_000) },
      { id: "bottom", value: inner(5_400_000) },
      { id: "bottomRight", value: inner(2_700_000) },
    ],
  },
  {
    group: "perspective",
    columns: 5,
    variations: [
      {
        id: "upperLeft",
        value: { blurRad: 76_200, dist: 0, dir: 13_500_000, algn: "br", color: BLACK, alpha: 20_000, sy: 23_000, kx: 1_200_000 },
      },
      {
        id: "upperRight",
        value: { blurRad: 76_200, dist: 0, dir: 18_900_000, algn: "bl", color: BLACK, alpha: 20_000, sy: 23_000, kx: -1_200_000 },
      },
      {
        id: "below",
        value: { blurRad: 152_400, dist: 317_500, dir: 5_400_000, algn: "b", color: BLACK, alpha: 15_000, sx: 90_000, sy: -19_000 },
      },
      {
        id: "lowerLeft",
        value: { blurRad: 76_200, dist: 12_700, dir: 8_100_000, algn: "br", color: BLACK, alpha: 20_000, sy: -23_000, kx: 800_400 },
      },
      {
        id: "lowerRight",
        value: { blurRad: 76_200, dist: 12_700, dir: 2_700_000, algn: "bl", color: BLACK, alpha: 20_000, sy: -23_000, kx: -800_400 },
      },
    ],
  },
];

function int(value: unknown, low: number, high: number): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= low && value <= high ? value : null;
}

/** The page's shadow value, or null when it is not one (the server's rule). */
export function normalizeOfficeShadowEffect(value: unknown): OfficeShadowEffect | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if ("inner" in raw && raw.inner !== true) return null;
  const isInner = raw.inner === true;
  const allowed = new Set(
    isInner
      ? ["inner", "blurRad", "dist", "dir", "color", "alpha"]
      : ["blurRad", "dist", "dir", "algn", "color", "alpha", "sx", "sy", "kx", "ky"],
  );
  if (Object.keys(raw).some((key) => !allowed.has(key))) return null;
  const blurRad = int(raw.blurRad, 0, MAX_BLUR);
  const dist = int(raw.dist, 0, MAX_DISTANCE);
  const dir = int(raw.dir, 0, ANGLE_LIMIT - 1);
  const alpha = int(raw.alpha, 0, 100_000);
  const color =
    typeof raw.color === "string" && /^#[0-9a-fA-F]{6}$/.test(raw.color) ? raw.color.toLowerCase() : null;
  if (blurRad === null || dist === null || dir === null || alpha === null || color === null) return null;
  if (isInner) return { inner: true, blurRad, dist, dir, color, alpha };
  if (!(OFFICE_SHADOW_ALIGNMENTS as readonly unknown[]).includes(raw.algn)) return null;
  const shadow: { -readonly [K in keyof OfficeShadowEffect]: OfficeShadowEffect[K] } = {
    blurRad,
    dist,
    dir,
    algn: raw.algn as OfficeShadowAlignment,
    color,
    alpha,
  };
  for (const key of ["sx", "sy"] as const) {
    if (!(key in raw)) continue;
    const scale = int(raw[key], -MAX_SCALE, MAX_SCALE);
    if (scale === null) return null;
    shadow[key] = scale;
  }
  for (const key of ["kx", "ky"] as const) {
    if (!(key in raw)) continue;
    const skew = int(raw[key], -MAX_SKEW, MAX_SKEW);
    if (skew === null) return null;
    shadow[key] = skew;
  }
  return shadow;
}

/** Whether two shadows are the same pick (a gallery tile's checked state). */
export function sameOfficeShadowEffect(left: unknown, right: unknown): boolean {
  const a = normalizeOfficeShadowEffect(left);
  const b = normalizeOfficeShadowEffect(right);
  if (!a || !b) return false;
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  return [...keys].every(
    (key) => (a as unknown as Record<string, unknown>)[key] === (b as unknown as Record<string, unknown>)[key],
  );
}

// --- drawing ------------------------------------------------------------------

function px(emu: number): number {
  return Math.round((emu / EMU_PER_PX) * 100) / 100;
}

function rgba(hex: string, alpha: number): string {
  const value = Number.parseInt(hex.slice(1), 16);
  return `rgb(${(value >> 16) & 255} ${(value >> 8) & 255} ${value & 255} / ${Math.round(alpha / 1000) / 100})`;
}

const ORIGIN: Record<OfficeShadowAlignment, string> = {
  tl: "0% 0%",
  t: "50% 0%",
  tr: "100% 0%",
  l: "0% 50%",
  ctr: "50% 50%",
  r: "100% 50%",
  bl: "0% 100%",
  b: "50% 100%",
  br: "100% 100%",
};

/**
 * How a shadow is drawn on the canvas, approximated with CSS:
 *
 * - `boxShadow`: an outer shadow at an even scale, on the object's own box
 *   (offset along `dir`, its blur, a scale grown as spread);
 * - `under`: a layer under the object for a scaled or skewed (perspective)
 *   outer shadow -- the object's box in the shadow colour, blurred, moved
 *   along `dir` and scaled / skewed about its alignment corner;
 * - `over`: a layer over the object for an inner shadow (an inset shadow
 *   darkening the side `dir` points to), so cell or picture paint does not
 *   hide it.
 *
 * `size` is the object's box in px.
 */
export interface OfficeShadowCss {
  readonly boxShadow?: string;
  readonly under?: CSSProperties;
  readonly over?: CSSProperties;
}












export const POWERPOINT_SHADOW_BLUR_DEVIATION = 1 / (2 * Math.SQRT2);

export interface OfficeShadowCssOptions {
  /** The Gaussian deviation an outer shadow's blur radius draws, per unit of it (default 1/2: CSS's own reading of a blur radius). */
  readonly blurDeviation?: number;
  /**
   * An evenly scaled shadow (sx = sy, no skew) is a layer under the box
   * scaled about its alignment point, as an uneven one is (a slide's
   * picture: PowerPoint grows a 150 % shadow about `tl` to the right and
   * down, probe lf13-shscale). Unstated, it is a box shadow whose spread
   * grows every side by the scale of the box's shorter side, as before.
   */
  readonly evenScaleUnder?: boolean;
}

export function officeShadowCss(
  value: unknown,
  size: { readonly width: number; readonly height: number },
  options: OfficeShadowCssOptions = {},
): OfficeShadowCss | null {
  const shadow = normalizeOfficeShadowEffect(value);
  if (!shadow) return null;
  const radians = (shadow.dir / 60_000) * (Math.PI / 180);
  const distance = px(shadow.dist);
  const x = Math.round(Math.cos(radians) * distance * 100) / 100 + 0;
  const y = Math.round(Math.sin(radians) * distance * 100) / 100 + 0;
  const blur = px(shadow.blurRad);
  // An outer shadow's blur, as the CSS blur radius (twice the deviation).
  const outerBlur = Math.round(blur * 2 * (options.blurDeviation ?? 0.5) * 100) / 100;
  const color = rgba(shadow.color, shadow.alpha);
  const layer: CSSProperties = { position: "absolute", inset: 0, pointerEvents: "none" };
  if (shadow.inner) {
    // PowerPoint darkens the side the inner shadow is offset TOWARD (its PDF
    // of Inside: Top Left, `dir` up-left, is dark along the top and left
    // edges); a CSS inset shadow darkens the side opposite its offset.
    return { over: { ...layer, boxShadow: `inset ${-x + 0}px ${-y + 0}px ${blur}px ${color}` } };
  }
  const sx = (shadow.sx ?? 100_000) / 100_000;
  const sy = (shadow.sy ?? 100_000) / 100_000;
  const kx = (shadow.kx ?? 0) / 60_000;
  const ky = (shadow.ky ?? 0) / 60_000;
  if (sx === sy && sx > 0 && !kx && !ky && !(options.evenScaleUnder && sx !== 1)) {
    const spread = Math.round((((sx - 1) * Math.min(size.width, size.height)) / 2) * 100) / 100;
    return { boxShadow: `${x}px ${y}px ${outerBlur}px ${spread}px ${color}` };
  }
  const silhouette = officeShadowSilhouette(shadow, options)!;
  return {
    under: {
      ...layer,
      zIndex: -1,
      background: color,
      filter: silhouette.filter || undefined,
      transformOrigin: silhouette.transformOrigin,
      transform: silhouette.transform,
    },
  };
}

/**
 * An outer shadow drawn as a silhouette: whatever casts it, painted in the
 * shadow colour (`color`, its opacity in it), blurred (`filter`, CSS
 * `blur()` or ""), moved along `dir` and scaled / skewed about its alignment
 * corner (`transform` about `transformOrigin`, a box-relative origin). The
 * scaled or skewed shadow's layer under a shape or picture (`officeShadowCss`'s
 * `under`), and the silhouette of what a table's cells paint
 * (`lib/presentation-table-effects`). Null for an inner or unreadable shadow.
 */
export interface OfficeShadowSilhouette {
  readonly color: string;
  readonly filter: string;
  readonly transform: string;
  readonly transformOrigin: string;
}

export function officeShadowSilhouette(value: unknown, options: OfficeShadowCssOptions = {}): OfficeShadowSilhouette | null {
  const shadow = normalizeOfficeShadowEffect(value);
  if (!shadow || shadow.inner) return null;
  const radians = (shadow.dir / 60_000) * (Math.PI / 180);
  const distance = px(shadow.dist);
  const x = Math.round(Math.cos(radians) * distance * 100) / 100 + 0;
  const y = Math.round(Math.sin(radians) * distance * 100) / 100 + 0;
  const outerBlur = Math.round(px(shadow.blurRad) * 2 * (options.blurDeviation ?? 0.5) * 100) / 100;
  const sx = (shadow.sx ?? 100_000) / 100_000;
  const sy = (shadow.sy ?? 100_000) / 100_000;
  const kx = (shadow.kx ?? 0) / 60_000;
  const ky = (shadow.ky ?? 0) / 60_000;
  return {
    color: rgba(shadow.color, shadow.alpha),
    filter: outerBlur ? `blur(${Math.round((outerBlur / 2) * 100) / 100}px)` : "",
    transformOrigin: ORIGIN[shadow.algn ?? "b"],
    // Skewed about the corner, then scaled (a negative scale flips the
    // shadow below it): Perspective: Upper Left leans up and left,
    // Lower Left down and left.
    transform: `translate(${x}px, ${y}px) scale(${sx}, ${sy}) skew(${kx}deg, ${ky}deg)`,
  };
}
