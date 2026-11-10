/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text Effects (WordArt) on a shape's words, as Office's galleries: the
 * shared text paint (lib/office-text-paint) read as the Office effects the
 * Shape Effects galleries pick (lib/office-picture-effects), and one pick
 * written back onto the paint. Shadow (Outer / Inner / Perspective),
 * Reflection, Glow, Bevel and 3-D Rotation; Transform is the shape's
 * text warp (lib/office-text-warp).
 *
 * The paint keeps CSS pixels; the galleries EMU, 60 000ths of a degree and
 * 1 000ths of a percent. A reflection's size is a whole percent on the paint
 * (Office's Tight 4 pt, 38.5 %, is written 39 %).
 */

import {
  officeBevelVariation,
  OFFICE_BEVEL_PRESETS,
  OFFICE_REFLECTION_VARIATIONS,
  type OfficeBevelEffect,
  type OfficeBevelPreset,
  type OfficeGlowEffect,
  type OfficePictureEffectKind,
  type OfficePictureEffects,
  type OfficePictureEffectValues,
  type OfficeReflectionEffect,
} from "./office-picture-effects";
import { OFFICE_SHADOW_GALLERY, type OfficeShadowEffect } from "./office-shadow-effects";
import {
  normalizeOfficeTextPaint,
  patchOfficeTextPaintParts,
  type OfficeTextPaint,
  type OfficeTextShadow,
} from "./office-text-paint";

const EMU_PER_PX = 9_525;
const ANGLE = 60_000;
const round = (value: number, digits = 3) => Math.round(value * 10 ** digits) / 10 ** digits;

/** The kinds Office's Text Effects menu lists (Transform aside). */
export const OFFICE_TEXT_EFFECT_KINDS = ["shadow", "reflection", "glow", "bevel", "rotation3d"] as const satisfies readonly OfficePictureEffectKind[];

function textShadowOf(shadow: OfficeShadowEffect): OfficeTextShadow {
  const radians = (shadow.dir / ANGLE) * (Math.PI / 180);
  const distance = shadow.dist / EMU_PER_PX;
  return {
    color: shadow.color,
    blur: Math.min(64, round(shadow.blurRad / EMU_PER_PX)),
    offsetX: Math.max(-64, Math.min(64, round(Math.cos(radians) * distance))),
    offsetY: Math.max(-64, Math.min(64, round(Math.sin(radians) * distance))),
    opacity: round(shadow.alpha / 100_000, 4),
    ...(shadow.inner
      ? {}
      : {
          ...(shadow.sx !== undefined && shadow.sx !== 100_000 ? { scale: round(shadow.sx / 1000) } : {}),
          ...(shadow.sy !== undefined && shadow.sy !== (shadow.sx ?? 100_000) ? { scaleY: round(shadow.sy / 1000) } : {}),
          ...(shadow.kx ? { skewX: round(shadow.kx / ANGLE) } : {}),
          ...(shadow.ky ? { skewY: round(shadow.ky / ANGLE) } : {}),
          ...(shadow.algn && shadow.algn !== "ctr" ? { align: shadow.algn as OfficeTextShadow["align"] } : {}),
        }),
  };
}

const sameText = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

function galleryShadowOf(text: OfficeTextShadow, inner: boolean): OfficeShadowEffect | undefined {
  for (const group of OFFICE_SHADOW_GALLERY) {
    for (const variation of group.variations) {
      if (Boolean(variation.value.inner) !== inner) continue;
      if (sameText(textShadowOf(variation.value), text)) return variation.value;
    }
  }
  return undefined;
}

function textReflectionOf(reflection: OfficeReflectionEffect): NonNullable<OfficeTextPaint["reflection"]> {
  return {
    opacity: round(reflection.stA / 100_000, 4),
    blur: Math.min(32, round(reflection.blurRad / EMU_PER_PX)),
    distance: Math.min(32, round(reflection.dist / EMU_PER_PX)),
    size: Math.max(0, Math.min(100, Math.round(reflection.endPos / 1000))),
  };
}

function textGlowOf(glow: OfficeGlowEffect): NonNullable<OfficeTextPaint["glow"]> {
  return {
    color: glow.color,
    radius: Math.min(32, round(glow.radius / EMU_PER_PX)),
    opacity: round(glow.alpha / 100_000, 4),
  };
}

/**
 * The paint's effects as Office's galleries name them, a gallery tile's own
 * value where the paint is one (so the menu checks it).
 */
export function officeTextPaintOfficeEffects(
  paint: OfficeTextPaint | null | undefined,
  glowVariations: readonly OfficeGlowEffect[] = [],
): OfficePictureEffects | undefined {
  if (!paint) return undefined;
  const effects: Partial<Record<OfficePictureEffectKind, unknown>> = {};
  const shadow = paint.innerShadow ?? paint.shadow;
  if (shadow) {
    const inner = Boolean(paint.innerShadow);
    effects.shadow = galleryShadowOf(shadow, inner) ?? {
      ...(inner ? { inner: true as const } : { algn: shadow.align ?? "ctr" }),
      blurRad: Math.round(shadow.blur * EMU_PER_PX),
      dist: Math.round(Math.hypot(shadow.offsetX, shadow.offsetY) * EMU_PER_PX),
      dir: Math.round(((((Math.atan2(shadow.offsetY, shadow.offsetX) * 180) / Math.PI) % 360) + 360) % 360 * ANGLE),
      color: shadow.color,
      alpha: Math.round(shadow.opacity * 100_000),
    };
  }
  if (paint.reflection) {
    const tile = OFFICE_REFLECTION_VARIATIONS.find((variation) =>
      sameText(textReflectionOf(variation.value), { ...paint.reflection, size: paint.reflection!.size ?? 40 }),
    );
    effects.reflection = tile?.value ?? {
      blurRad: Math.round(paint.reflection.blur * EMU_PER_PX),
      stA: Math.round(paint.reflection.opacity * 100_000),
      endA: 300,
      endPos: (paint.reflection.size ?? 40) * 1000,
      dist: Math.round(paint.reflection.distance * EMU_PER_PX),
    };
  }
  if (paint.glow) {
    effects.glow =
      glowVariations.find((variation) => sameText(textGlowOf(variation), paint.glow)) ??
      { radius: Math.round(paint.glow.radius * EMU_PER_PX), color: paint.glow.color, alpha: Math.round(paint.glow.opacity * 100_000) };
  }
  if (paint.bevel) {
    const preset = (OFFICE_BEVEL_PRESETS as readonly string[]).includes(paint.bevel.preset ?? "")
      ? (paint.bevel.preset as OfficeBevelPreset)
      : "circle";
    const tile = officeBevelVariation(preset);
    const width = Math.round(paint.bevel.width * EMU_PER_PX);
    const height = Math.round(paint.bevel.height * EMU_PER_PX);
    effects.bevel =
      Math.abs(width - tile.width) < EMU_PER_PX / 2 && Math.abs(height - tile.height) < EMU_PER_PX / 2
        ? tile
        : ({ preset, width, height } satisfies OfficeBevelEffect);
  }
  if (paint.rotation3d) effects.rotation3d = { camera: paint.rotation3d.camera };
  return Object.keys(effects).length ? (effects as OfficePictureEffects) : undefined;
}

/**
 * The paint after one Text Effects pick (`value` null: "No <effect>"); the
 * others are kept, as Office's effects combine. An edit is the user's own
 * paint, so Office's statement is dropped.
 */
export function withOfficeTextEffect<K extends OfficePictureEffectKind>(
  paint: OfficeTextPaint,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
): OfficeTextPaint {
  switch (kind) {
    case "shadow": {
      const shadow = value as OfficeShadowEffect | null;
      if (!shadow) return patchOfficeTextPaintParts(paint, { shadow: null, innerShadow: null });
      return shadow.inner
        ? patchOfficeTextPaintParts(paint, { shadow: null, innerShadow: textShadowOf(shadow) })
        : patchOfficeTextPaintParts(paint, { shadow: textShadowOf(shadow), innerShadow: null });
    }
    case "reflection": {
      const reflection = value as OfficeReflectionEffect | null;
      return patchOfficeTextPaintParts(paint, { reflection: reflection ? textReflectionOf(reflection) : null });
    }
    case "glow": {
      const glow = value as OfficeGlowEffect | null;
      return patchOfficeTextPaintParts(paint, { glow: glow ? textGlowOf(glow) : null });
    }
    case "bevel": {
      const bevel = value as OfficeBevelEffect | null;
      return patchOfficeTextPaintParts(paint, {
        bevel: bevel
          ? {
              width: Math.min(16, round(bevel.width / EMU_PER_PX)),
              height: Math.min(16, round(bevel.height / EMU_PER_PX)),
              depth: paint.bevel?.depth ?? 0,
              ...(bevel.preset !== "circle" ? { preset: bevel.preset } : {}),
            }
          : null,
      });
    }
    case "rotation3d": {
      const rotation = value as { camera: string } | null;
      const next = { ...normalizeOfficeTextPaint(paint)! } as OfficeTextPaint & Record<string, unknown>;
      delete next.stated;
      if (rotation && rotation.camera !== "orthographicFront") next.rotation3d = { camera: rotation.camera };
      else delete next.rotation3d;
      return normalizeOfficeTextPaint(next)!;
    }
    default:
      return paint;
  }
}
