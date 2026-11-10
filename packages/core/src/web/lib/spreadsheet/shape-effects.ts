/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shadow effects for floating objects (Excel/WPS 形状效果 → 阴影).
 *
 * WPS lets any drawn object — a shape, a picture, a chart — carry a shadow,
 * chosen from a grid of directions and then adjusted for colour, distance,
 * blur and transparency. Without it every inserted object sits flat on the
 * sheet, which is the difference between a diagram that reads as layered and
 * one that reads as clip art.
 *
 * Rendered as a CSS `drop-shadow` filter rather than `box-shadow`, because a
 * shadow has to follow the object's actual silhouette: an ellipse, a triangle
 * and a transparent PNG must not all cast a rectangle.
 *
 * Dependency-free and deterministic, so the geometry is unit-testable without
 * the editor.
 */

/**
 * The directions WPS offers, named for where the shadow falls rather than for
 * a position in the gallery, so a stored value survives a redesign.
 *
 * `outer-center` casts no offset at all — a halo, which is what the centre
 * cell of the gallery does. The two `perspective` entries throw further and
 * softer, the way a light placed above or below the object would.
 */
export const SPREADSHEET_SHADOW_PRESETS = [
  "outer-bottom-right",
  "outer-bottom",
  "outer-bottom-left",
  "outer-right",
  "outer-center",
  "outer-left",
  "outer-top-right",
  "outer-top",
  "outer-top-left",
  "perspective-below",
  "perspective-above",
] as const;
export type SpreadsheetShadowPreset =
  (typeof SPREADSHEET_SHADOW_PRESETS)[number];
const PRESET_SET: ReadonlySet<string> = new Set(SPREADSHEET_SHADOW_PRESETS);

export const DEFAULT_SHADOW_PRESET: SpreadsheetShadowPreset =
  "outer-bottom-right";
export const DEFAULT_SHADOW_COLOR = "#000000";
export const DEFAULT_SHADOW_OPACITY = 0.35;
export const MAX_SHADOW_DISTANCE = 64;
export const MAX_SHADOW_BLUR = 64;

/** Where a preset throws its shadow, and how hard, before any override. */
interface ShadowGeometry {
  /** Unit direction; both axes are offset by the full distance on a diagonal,
   *  as Office does, so a 45° shadow reads as further than a straight one. */
  dx: -1 | 0 | 1;
  dy: -1 | 0 | 1;
  distance: number;
  blur: number;
  opacity: number;
}

const GEOMETRY: Record<SpreadsheetShadowPreset, ShadowGeometry> = {
  "outer-bottom-right": { dx: 1, dy: 1, distance: 4, blur: 6, opacity: 0.35 },
  "outer-bottom": { dx: 0, dy: 1, distance: 4, blur: 6, opacity: 0.35 },
  "outer-bottom-left": { dx: -1, dy: 1, distance: 4, blur: 6, opacity: 0.35 },
  "outer-right": { dx: 1, dy: 0, distance: 4, blur: 6, opacity: 0.35 },
  "outer-center": { dx: 0, dy: 0, distance: 0, blur: 8, opacity: 0.4 },
  "outer-left": { dx: -1, dy: 0, distance: 4, blur: 6, opacity: 0.35 },
  "outer-top-right": { dx: 1, dy: -1, distance: 4, blur: 6, opacity: 0.35 },
  "outer-top": { dx: 0, dy: -1, distance: 4, blur: 6, opacity: 0.35 },
  "outer-top-left": { dx: -1, dy: -1, distance: 4, blur: 6, opacity: 0.35 },
  "perspective-below": { dx: 0, dy: 1, distance: 10, blur: 14, opacity: 0.28 },
  "perspective-above": { dx: 0, dy: -1, distance: 10, blur: 14, opacity: 0.28 },
};

/** A stored shadow. Absent on the object means no shadow at all. */
export interface SpreadsheetShadow {
  preset: SpreadsheetShadowPreset;
  /** #rrggbb. Absent uses black, as the gallery presets do. */
  color?: string;
  /** How far the shadow is thrown, in px. Absent uses the preset's own. */
  distance?: number;
  /** Softness in px. Absent uses the preset's own. */
  blur?: number;
  /** 0 (invisible) … 1 (solid). Absent uses the preset's own. */
  opacity?: number;
}

const HEX_RE = /^#[0-9a-fA-F]{6}$/;

/** Coerce to a known preset, or the default for anything unrecognised. */
export function normalizeShadowPreset(value: unknown): SpreadsheetShadowPreset {
  return typeof value === "string" && PRESET_SET.has(value)
    ? (value as SpreadsheetShadowPreset)
    : DEFAULT_SHADOW_PRESET;
}

function clampNumber(
  value: unknown,
  min: number,
  max: number,
): number | undefined {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return undefined;
  return Math.max(min, Math.min(max, n));
}

/**
 * Validate a stored shadow.
 *
 * Returns undefined for "no shadow" — including for the string `"none"` and
 * for junk — so an object without one persists nothing and stays byte-for-byte
 * what it was before shadows existed.
 */
export function normalizeShadow(value: unknown): SpreadsheetShadow | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const raw = value as Record<string, unknown>;
  if (raw.preset === "none" || raw.preset === undefined) return undefined;

  const shadow: SpreadsheetShadow = { preset: normalizeShadowPreset(raw.preset) };
  if (typeof raw.color === "string" && HEX_RE.test(raw.color)) {
    shadow.color = raw.color.toLowerCase();
  }
  const distance = clampNumber(raw.distance, 0, MAX_SHADOW_DISTANCE);
  if (distance !== undefined) shadow.distance = Math.round(distance);
  const blur = clampNumber(raw.blur, 0, MAX_SHADOW_BLUR);
  if (blur !== undefined) shadow.blur = Math.round(blur);
  const opacity = clampNumber(raw.opacity, 0, 1);
  if (opacity !== undefined) shadow.opacity = Math.round(opacity * 100) / 100;
  return shadow;
}

function rgba(hex: string, opacity: number): string {
  const r = Number.parseInt(hex.slice(1, 3), 16);
  const g = Number.parseInt(hex.slice(3, 5), 16);
  const b = Number.parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

/**
 * The CSS `filter` value for a shadow, or undefined when there is none.
 *
 * A `drop-shadow` follows the alpha silhouette, so an ellipse casts an ellipse
 * and a transparent picture casts its subject rather than its bounding box.
 */
export function shadowFilterCss(shadow: unknown): string | undefined {
  const normalized = normalizeShadow(shadow);
  if (!normalized) return undefined;
  const base = GEOMETRY[normalized.preset];
  const distance = normalized.distance ?? base.distance;
  const blur = normalized.blur ?? base.blur;
  const opacity = normalized.opacity ?? base.opacity;
  const color = rgba(normalized.color ?? DEFAULT_SHADOW_COLOR, opacity);
  const offsetX = base.dx * distance;
  const offsetY = base.dy * distance;
  return `drop-shadow(${offsetX}px ${offsetY}px ${blur}px ${color})`;
}

/**
 * Room a shadow needs around the object so the filter is not clipped by the
 * object's own box. Callers pad the container by this much.
 */
export function shadowPaddingPx(shadow: unknown): number {
  const normalized = normalizeShadow(shadow);
  if (!normalized) return 0;
  const base = GEOMETRY[normalized.preset];
  const distance = normalized.distance ?? base.distance;
  const blur = normalized.blur ?? base.blur;
  return Math.ceil(distance + blur);
}

/** A stored shadow with its preset's own values filled in (direction as a
 * unit offset, distance and blur in px, opacity 0..1); undefined for none. */
export function spreadsheetShadowResolved(
  shadow: unknown,
): { dx: number; dy: number; distance: number; blur: number; opacity: number; color: string } | undefined {
  const normalized = normalizeShadow(shadow);
  if (!normalized) return undefined;
  const base = GEOMETRY[normalized.preset];
  return {
    dx: base.dx,
    dy: base.dy,
    distance: normalized.distance ?? base.distance,
    blur: normalized.blur ?? base.blur,
    opacity: normalized.opacity ?? base.opacity,
    color: normalized.color ?? DEFAULT_SHADOW_COLOR,
  };
}
