/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a shape saved before Office's combinable effects (`shapeEffects`)
 * carries, read as the Office effect each one drew: the older shadow switch,
 * a custom `shapeShadow`, and the one-of-seven Shape Effects preset name. The
 * deck, Docs and the worksheet read their older shapes through this one table.
 */

import {
  officeBevelVariation,
  officeGlowVariation,
  officeSoftEdgeVariation,
  OFFICE_REFLECTION_VARIATIONS,
  type OfficePictureEffectKind,
} from "./office-picture-effects";
import {
  OFFICE_SHADOW_ALIGNMENTS,
  type OfficeShadowAlignment,
  type OfficeShadowEffect,
} from "./office-shadow-effects";
import { OFFICE_SHADOW_DEFAULT_DIRECTION, type OfficeShapeShadow } from "./office-shape-paint";

const EMU_PER_PX = 9_525;

/** The shadow an older shadow switch alone draws (the deck's default). */
export const OFFICE_SWITCH_SHADOW: OfficeShadowEffect = {
  blurRad: 6 * EMU_PER_PX,
  dist: 3 * EMU_PER_PX,
  dir: 5_400_000,
  algn: "ctr",
  color: "#0f172a",
  alpha: 28_000,
};

export function officeShadowFromShapeShadow(shadow: OfficeShapeShadow): OfficeShadowEffect {
  const algn = (OFFICE_SHADOW_ALIGNMENTS as readonly string[]).includes(shadow.align ?? "")
    ? (shadow.align as OfficeShadowAlignment)
    : "ctr";
  return {
    blurRad: Math.round(Math.max(0, shadow.blur) * EMU_PER_PX),
    dist: Math.round(Math.max(0, shadow.distance) * EMU_PER_PX),
    dir: Math.round((shadow.direction ?? OFFICE_SHADOW_DEFAULT_DIRECTION) * 60_000),
    algn,
    color: shadow.color,
    alpha: Math.round(Math.min(1, Math.max(0, shadow.opacity)) * 100_000),
  };
}

/**
 * The Office effect an older one-of-seven Shape Effects preset name drew
 * (`effect`, on a slide's or a sheet's shape): the first variation of its
 * gallery, Glow in Accent 1.
 */
export function officeEffectsOfLegacyPreset(
  effect: unknown,
  accent1 = "#4472c4",
): Partial<Record<OfficePictureEffectKind, unknown>> {
  const legacy: Partial<Record<OfficePictureEffectKind, unknown>> = {};
  switch (effect) {
    case "reflection":
      legacy.reflection = OFFICE_REFLECTION_VARIATIONS[0]!.value;
      break;
    case "glow":
      legacy.glow = officeGlowVariation(5, "accent1", accent1);
      break;
    case "softEdge":
      legacy.softEdge = officeSoftEdgeVariation(2.5);
      break;
    case "bevel":
      legacy.bevel = officeBevelVariation("circle");
      break;
    case "threeD":
      legacy.rotation3d = { camera: "isometricOffAxis1Right" };
      break;
    default:
      break;
  }
  return legacy;
}
