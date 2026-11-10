/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A sheet shape's Office effects (`object.shapeEffects`, the model slides,
 * pictures and Docs shapes keep) and the Shape Format writes on them.
 */

import {
  normalizeOfficePictureEffects,
  OFFICE_PICTURE_EFFECT_KINDS,
  withOfficePictureEffect,
  type OfficePictureEffectKind,
  type OfficePictureEffects,
  type OfficePictureEffectValues,
} from "@mrnote/core/web/lib/office-picture-effects.ts";
import { officeEffectsOfLegacyPreset } from "@mrnote/core/web/lib/office-shape-legacy-effects.ts";
import { spreadsheetShadowResolved } from "@mrnote/core/web/lib/spreadsheet/shape-effects.ts";

const EMU_PER_PX = 9_525;

/**
 * A sheet shape's effects as Office states them: its own `shapeEffects`, else
 * the older shadow preset and one-of-seven effect name it carries, read as
 * the Office effect each one drew (so the first pick in the full Shape
 * Effects menu adds to what the shape already shows). `accent1` is the
 * workbook theme's Accent 1 (an older Glow drew in it).
 */
export function spreadsheetShapeOfficeEffects(
  object: { readonly shapeEffects?: unknown; readonly shadow?: unknown; readonly effect?: unknown },
  accent1?: string,
): OfficePictureEffects | undefined {
  const own = normalizeOfficePictureEffects(object.shapeEffects);
  if (own) return own;
  const legacy: Partial<Record<OfficePictureEffectKind, unknown>> = {
    ...officeEffectsOfLegacyPreset(object.effect, accent1),
  };
  const shadow = spreadsheetShadowResolved(object.shadow);
  if (shadow) {
    const distance = Math.hypot(shadow.dx, shadow.dy) * shadow.distance;
    const degrees = (Math.atan2(shadow.dy, shadow.dx) * 180) / Math.PI;
    legacy.shadow = {
      // The CSS blur radius this shadow drew with is half Office's `blurRad`.
      blurRad: Math.round(shadow.blur * 2 * EMU_PER_PX),
      dist: Math.round(distance * EMU_PER_PX),
      dir: distance ? Math.round((((degrees % 360) + 360) % 360) * 60_000) : 0,
      algn: "ctr",
      color: shadow.color,
      alpha: Math.round(shadow.opacity * 100_000),
    };
  }
  return normalizeOfficePictureEffects(legacy) ?? undefined;
}

/**
 * Shape Format > Shape Effects on a sheet shape: `parameter` is the pick as
 * JSON (`{kind, value, accent1?}`, `value` null for "No <effect>"). The
 * object's whole `shapeEffects`, with the older shadow and preset name
 * cleared since it now states them; null when nothing changes.
 */
export function shapeOfficeEffectsPatch(
  object: { readonly shapeEffects?: unknown; readonly shadow?: unknown; readonly effect?: unknown },
  parameter: string,
): Record<string, unknown> | null {
  let picked: { kind?: unknown; value?: unknown; accent1?: unknown };
  try {
    picked = JSON.parse(parameter) as typeof picked;
  } catch {
    return null;
  }
  if (!picked || typeof picked !== "object") return null;
  const kind = picked.kind;
  if (!(OFFICE_PICTURE_EFFECT_KINDS as readonly unknown[]).includes(kind)) return null;
  const accent1 = typeof picked.accent1 === "string" ? picked.accent1 : undefined;
  const next = withOfficePictureEffect(
    spreadsheetShapeOfficeEffects(object, accent1),
    kind as OfficePictureEffectKind,
    (picked.value ?? null) as OfficePictureEffectValues[OfficePictureEffectKind] | null,
    false,
  );
  if (
    JSON.stringify(next ?? null) === JSON.stringify(normalizeOfficePictureEffects(object.shapeEffects) ?? null) &&
    object.shadow === undefined &&
    object.effect === undefined
  ) {
    return null;
  }
  return { shapeEffects: next, shadow: undefined, effect: undefined };
}
