/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Shape Effects presets, drawn the same way in the three
 * editors (Docs' node view, the deck's shape block, the worksheet's shape
 * object). The file side is `app/services/office_shape_effects.py` on the
 * server and Docs' browser DOCX writer; this module is the screen side.
 *
 * MRAI's menu applies one preset at a time (`OFFICE_SHAPE_EFFECTS` in the
 * shared `OfficeShapeFormatTab`); each is an approximation of Office's
 * rendering with CSS filters, so a glow is a coloured halo, a soft edge a
 * slight blur, bevel a light / dark rim and 3-D an offset solid.
 */

import type { OfficeShapeEffect } from "../components/office-ribbon/OfficeShapeFormatTab";

/** The presets a host stores beside its own shadow model. */
export type OfficeShapeEffectPreset = Exclude<OfficeShapeEffect, "none" | "shadow">;

export const OFFICE_SHAPE_EFFECT_PRESETS: readonly OfficeShapeEffectPreset[] = [
  "reflection",
  "glow",
  "softEdge",
  "bevel",
  "threeD",
];

export function normalizeOfficeShapeEffectPreset(
  value: unknown,
): OfficeShapeEffectPreset | null {
  return typeof value === "string" &&
    (OFFICE_SHAPE_EFFECT_PRESETS as readonly string[]).includes(value)
    ? (value as OfficeShapeEffectPreset)
    : null;
}

/** The preset's CSS `filter` (empty for none and reflection). */
export function officeShapeEffectFilterCss(effect: unknown): string {
  switch (effect) {
    case "shadow":
      return "drop-shadow(4px 4px 4px rgb(15 23 42 / 34%))";
    case "glow":
      return "drop-shadow(0 0 6px rgb(59 130 246 / 72%))";
    case "softEdge":
      return "blur(0.7px)";
    case "bevel":
      return "drop-shadow(1px 1px 0 rgb(255 255 255 / 70%)) drop-shadow(-1px -1px 0 rgb(15 23 42 / 30%))";
    case "threeD":
      return "drop-shadow(5px 5px 2px rgb(15 23 42 / 28%))";
    default:
      return "";
  }
}

/** The Reflection preset's `-webkit-box-reflect` (empty for the others). */
export function officeShapeEffectReflectionCss(effect: unknown): string {
  return effect === "reflection"
    ? "below -8px linear-gradient(transparent, rgb(15 23 42 / 22%))"
    : "";
}

/**
 * Reflection for a layer drawn inside the host, such as a shape's text
 * rectangle: Office reflects the shape WITH its text, about one line. `top` and
 * `bottom` are the layer's insets as fractions of the host's height. A layer
 * whose bottom sits `i` above the host's reflects `o + 2i` below itself; the
 * `2i` is stated as a percentage of the layer's own height, so it holds at any
 * zoom.
 */
export function officeShapeEffectInsetLayerReflectionCss(
  effect: unknown,
  top: number,
  bottom: number,
): string {
  const reflect = officeShapeEffectReflectionCss(effect);
  const layerHeight = 1 - top - bottom;
  if (!reflect || !(bottom > 0) || !(layerHeight > 0)) return reflect;
  const lift = Math.round(((2 * bottom) / layerHeight) * 100_000) / 1000;
  return reflect.replace(
    /^below (-?\d+(?:\.\d+)?px)/,
    (_match, offset: string) => `below calc(${offset} + ${lift}%)`,
  );
}

/**
 * The preset's CSS for a host element, after the host's own `filter` (its
 * shadow): `filter` and, for Reflection, `-webkit-box-reflect`.
 */
export function officeShapeEffectStyle(
  effect: unknown,
  hostFilter?: string,
): { filter?: string; WebkitBoxReflect?: string } {
  const preset = normalizeOfficeShapeEffectPreset(effect);
  const filter = [hostFilter ?? "", preset ? officeShapeEffectFilterCss(preset) : ""]
    .filter(Boolean)
    .join(" ");
  const reflect = officeShapeEffectReflectionCss(preset);
  return {
    ...(filter ? { filter } : {}),
    ...(reflect ? { WebkitBoxReflect: reflect } : {}),
  };
}

/**
 * The shared menu's value for a host that keeps its shadow apart: the preset
 * when one is stored, else Shadow / None from the host's shadow flag.
 */
export function officeShapeEffectMenuValue(
  preset: unknown,
  shadow: boolean,
): OfficeShapeEffect {
  return normalizeOfficeShapeEffectPreset(preset) ?? (shadow ? "shadow" : "none");
}
