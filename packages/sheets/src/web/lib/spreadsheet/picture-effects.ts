/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Picture Effects on a worksheet picture beyond its Shadow:
 * Reflection, Glow, Soft Edges, Bevel, 3-D Format and 3-D Rotation, as the
 * picture's `pictureEffects` -- the model the deck's pictures and every
 * editor's shapes keep (lib/office-picture-effects; the server reads and
 * writes it on `xdr:pic/xdr:spPr`, reading row 543 `xlsxPictureEffects`).
 * The Shadow stays the sheet's own shadow preset (`object.shadow`, the
 * Shadow presets Excel's pane names), so the picture has one shadow -- except
 * on a picture whose reading put its Shadow and Preset in its Picture
 * Effects (readings row 570, `spreadsheetPictureShadowInEffects`): Excel's
 * gallery shadows (Outer, Inner, Perspective) and Presets 1-12 there.
 */

import {
  normalizeOfficePictureEffects,
  withOfficePictureEffect,
  type OfficePictureEffectKind,
  type OfficePictureEffects,
  type OfficePictureEffectValues,
} from "@mrnote/core/web/lib/office-picture-effects.ts";
import { officePictureStyleEffectKinds } from "@mrnote/core/web/lib/office-picture-styles/native.ts";

/** The Picture Effects menu's categories on a worksheet picture (its Shadow is the sheet's own). */
export const SPREADSHEET_PICTURE_EFFECT_MENU_KINDS = [
  "reflection",
  "glow",
  "softEdge",
  "bevel",
  "rotation3d",
] as const satisfies readonly OfficePictureEffectKind[];

/** The effects a worksheet picture keeps in `pictureEffects` (the menu's and 3-D Format). */
export const SPREADSHEET_PICTURE_EFFECT_KINDS: readonly OfficePictureEffectKind[] = [
  ...SPREADSHEET_PICTURE_EFFECT_MENU_KINDS,
  "format3d",
];

/**
 * The Picture Effects menu's categories on a picture that keeps its Shadow
 * and Preset in its Picture Effects (readings row 570): Excel's whole menu,
 * Preset 1-12 and Shadow (Outer / Inner / Perspective) first.
 */
export const SPREADSHEET_PICTURE_EFFECT_MENU_KINDS_WITH_SHADOW = [
  "preset",
  "shadow",
  ...SPREADSHEET_PICTURE_EFFECT_MENU_KINDS,
] as const satisfies readonly OfficePictureEffectKind[];

const KINDS_WITH_SHADOW: readonly OfficePictureEffectKind[] = [...SPREADSHEET_PICTURE_EFFECT_KINDS, "preset", "shadow"];

/**
 * Whether the picture's Shadow and Preset are its Picture Effects' (the
 * server read them so, `source.pictureShadowInEffects`, readings row 570)
 * rather than the sheet's own shadow preset (`object.shadow`): one shadow.
 */
export function spreadsheetPictureShadowInEffects(object: object): boolean {
  const fields = object as { source?: unknown; shadow?: unknown };
  const source = fields.source && typeof fields.source === "object" ? (fields.source as Record<string, unknown>) : null;
  return source?.pictureShadowInEffects === true && !fields.shadow;
}

/** The effects the picture keeps in `pictureEffects`. */
export function spreadsheetPictureEffectKinds(object: object): readonly OfficePictureEffectKind[] {
  return spreadsheetPictureShadowInEffects(object) ? KINDS_WITH_SHADOW : SPREADSHEET_PICTURE_EFFECT_KINDS;
}

/**
 * A worksheet picture's `pictureEffects`, validated; undefined for none.
 * `object` (the picture) keeps its Shadow and Preset where they are its
 * Picture Effects'.
 */
export function normalizeSpreadsheetPictureEffects(value: unknown, object?: object): OfficePictureEffects | undefined {
  const effects = normalizeOfficePictureEffects(value);
  if (!effects) return undefined;
  const kinds = object ? spreadsheetPictureEffectKinds(object) : SPREADSHEET_PICTURE_EFFECT_KINDS;
  const kept = Object.fromEntries(
    Object.entries(effects).filter(([kind]) => kinds.includes(kind as OfficePictureEffectKind)),
  );
  return Object.keys(kept).length ? (kept as OfficePictureEffects) : undefined;
}

/**
 * Why the picture's effects cannot be edited, or null: effects the file
 * states that the page cannot (`pictureEffectsDisposition`), which stay the
 * package's.
 */
export function spreadsheetPictureEffectsBlocker(object: object): "preserved" | null {
  const source = (object as { source?: unknown }).source;
  return source && typeof source === "object" && (source as { pictureEffectsDisposition?: unknown }).pictureEffectsDisposition === "preserve-only"
    ? "preserved"
    : null;
}

/**
 * One Picture Effects pick (a menu tile, a pane field; null is "No
 * <effect>") as the picture's patch, or null when it changes nothing or the
 * picture's effects are not editable. A Picture Style that draws the effect
 * is answered with "none", so the save removes the style's own.
 */
export function spreadsheetPictureEffectPatch<K extends OfficePictureEffectKind>(
  object: object,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
): { pictureEffects: OfficePictureEffects | undefined } | null {
  if (!spreadsheetPictureEffectKinds(object).includes(kind)) return null;
  if (spreadsheetPictureEffectsBlocker(object)) return null;
  const fields = object as { pictureEffects?: unknown; pictureStyle?: unknown };
  const current = normalizeSpreadsheetPictureEffects(fields.pictureEffects, object);
  const next = normalizeSpreadsheetPictureEffects(
    withOfficePictureEffect(current, kind, value, officePictureStyleEffectKinds(fields.pictureStyle).has(kind)),
    object,
  );
  if (JSON.stringify(next ?? null) === JSON.stringify(current ?? null)) return null;
  return { pictureEffects: next };
}

/** The pick a Picture Effects control carries as its string parameter. */
export function spreadsheetPictureEffectParameter(kind: OfficePictureEffectKind, value: unknown): string {
  return JSON.stringify({ kind, value: value ?? null });
}

/** A control parameter back as its pick, or null when it is not one. */
export function spreadsheetPictureEffectFromParameter(
  parameter: string,
): { kind: OfficePictureEffectKind; value: unknown } | null {
  try {
    const parsed = JSON.parse(parameter) as { kind?: unknown; value?: unknown };
    // Which kinds the picture takes is the patch's to decide (a Shadow or a
    // Preset only where they are its Picture Effects').
    if (typeof parsed?.kind !== "string" || !KINDS_WITH_SHADOW.includes(parsed.kind as OfficePictureEffectKind)) {
      return null;
    }
    if (parsed.value !== null && !normalizeOfficePictureEffects({ [parsed.kind]: parsed.value })) return null;
    return { kind: parsed.kind as OfficePictureEffectKind, value: parsed.value };
  } catch {
    return null;
  }
}
