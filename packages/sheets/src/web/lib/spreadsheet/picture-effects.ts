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
 * Shadow presets Excel's pane names), so the picture has one shadow.
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

/** A worksheet picture's `pictureEffects`, validated; undefined for none. */
export function normalizeSpreadsheetPictureEffects(value: unknown): OfficePictureEffects | undefined {
  const effects = normalizeOfficePictureEffects(value);
  if (!effects) return undefined;
  const kept = Object.fromEntries(
    Object.entries(effects).filter(([kind]) => SPREADSHEET_PICTURE_EFFECT_KINDS.includes(kind as OfficePictureEffectKind)),
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
  if (!SPREADSHEET_PICTURE_EFFECT_KINDS.includes(kind)) return null;
  if (spreadsheetPictureEffectsBlocker(object)) return null;
  const fields = object as { pictureEffects?: unknown; pictureStyle?: unknown };
  const current = normalizeSpreadsheetPictureEffects(fields.pictureEffects);
  const next = normalizeSpreadsheetPictureEffects(
    withOfficePictureEffect(current, kind, value, officePictureStyleEffectKinds(fields.pictureStyle).has(kind)),
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
    if (typeof parsed?.kind !== "string" || !SPREADSHEET_PICTURE_EFFECT_KINDS.includes(parsed.kind as OfficePictureEffectKind)) {
      return null;
    }
    if (parsed.value !== null && !normalizeOfficePictureEffects({ [parsed.kind]: parsed.value })) return null;
    return { kind: parsed.kind as OfficePictureEffectKind, value: parsed.value };
  } catch {
    return null;
  }
}
