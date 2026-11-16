/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officePictureEffectPresetEffects, isOfficePictureEffectPreset } from "./office-picture-effect-presets";
import { normalizeOfficePictureEffects } from "./office-picture-effects";
import { OFFICE_STROKE_COMPOUNDS, type OfficeStrokeCompound } from "./office-stroke-detail";

export type OfficePictureLineJoin = "round" | "bevel" | "miter";

/** `a:miter@lim` PowerPoint writes (8 line widths), as an SVG miter limit. */
export const OFFICE_PICTURE_LINE_MITER_LIMIT = 8;

/** A page's `borderCompound`, when it is one. */
export function officePictureLineCompound(value: unknown): OfficeStrokeCompound | undefined {
  return typeof value === "string" && (OFFICE_STROKE_COMPOUNDS as readonly string[]).includes(value)
    ? (value as OfficeStrokeCompound)
    : undefined;
}

/** A page's `borderJoin`; a frame stating none draws round corners. */
export function officePictureLineJoin(value: unknown): OfficePictureLineJoin {
  return value === "miter" || value === "bevel" ? value : "round";
}

/**
 * Whether PowerPoint draws the picture in 3-D (a top or bottom bevel, or
 * depth: Picture Effects > Bevel / 3-D Format, or a Preset with a bevel),
 * which centres its line on its frame. A 3-D scene alone, or a 3-D shape
 * stating none of them, draws a flat picture.
 */
export function officePictureDrawnIn3d(pictureEffects: unknown): boolean {
  const effects = normalizeOfficePictureEffects(pictureEffects);
  if (!effects) return false;
  const bevel = effects.bevel;
  if (bevel && bevel !== "none" && (bevel.width > 0 || bevel.height > 0)) return true;
  const format = effects.format3d;
  if (format && format !== "none") {
    const bottom = format.bottom;
    if (bottom && (bottom.width > 0 || bottom.height > 0)) return true;
    if (format.depth && format.depth.height > 0) return true;
  }
  return (
    typeof effects.preset === "number" &&
    isOfficePictureEffectPreset(effects.preset) &&
    bevel !== "none" &&
    officePictureEffectPresetEffects(effects.preset).bevel > 0
  );
}

/**
 * How far outside its frame a picture's line is stroked, design px: half its
 * width for a flat picture, none for one drawn in 3-D.
 */
export function officePictureLineOutset(lineWidth: number, pictureEffects: unknown): number {
  return lineWidth > 0 && !officePictureDrawnIn3d(pictureEffects) ? lineWidth / 2 : 0;
}

/**
 * The path a rectangular picture's line is stroked along, in its frame's
 * design px: the frame (its corners rounded by `radius`) pushed out by
 * `outset`. A rounded corner keeps its centre, its radius grown by the outset;
 * a square one stays square (the join draws the corner).
 */
export function officePictureLineRectPath(width: number, height: number, radius: number, outset: number): string {
  const x = -outset;
  const y = -outset;
  const right = width + outset;
  const bottom = height + outset;
  const r = radius > 0 ? Math.min(radius + outset, (right - x) / 2, (bottom - y) / 2) : 0;
  const n = (value: number) => Math.round(value * 1000) / 1000;
  if (r <= 0) return `M${n(x)} ${n(y)}H${n(right)}V${n(bottom)}H${n(x)}Z`;
  return (
    `M${n(x + r)} ${n(y)}H${n(right - r)}A${n(r)} ${n(r)} 0 0 1 ${n(right)} ${n(y + r)}` +
    `V${n(bottom - r)}A${n(r)} ${n(r)} 0 0 1 ${n(right - r)} ${n(bottom)}` +
    `H${n(x + r)}A${n(r)} ${n(r)} 0 0 1 ${n(x)} ${n(bottom - r)}` +
    `V${n(y + r)}A${n(r)} ${n(r)} 0 0 1 ${n(x + r)} ${n(y)}Z`
  );
}
