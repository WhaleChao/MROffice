/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Shape Effects on a deck shape: Office's combinable effects
 * (`shapeEffects`, the model pictures and Docs shapes keep). A shape saved
 * before them carries a shadow switch, a custom shadow or one preset name;
 * those read here as the Office effect each one draws, so the first pick in
 * the full menu adds to what the shape already shows instead of replacing it.
 */

import type {
  PresentationBlock,
  PresentationShapeBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  normalizeOfficePictureEffects,
  withOfficePictureEffect,
  type OfficePictureEffectKind,
  type OfficePictureEffects,
  type OfficePictureEffectValues,
} from "@mrnote/core/web/lib/office-picture-effects.ts";
import {
  OFFICE_SWITCH_SHADOW,
  officeEffectsOfLegacyPreset,
  officeShadowFromShapeShadow,
} from "@mrnote/core/web/lib/office-shape-legacy-effects.ts";
import { shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";

/**
 * The shape's effects as Office states them. `accent1` is the deck's Accent 1
 * (a legacy Glow preset drew in it).
 */
export function presentationShapeEffects(
  block: Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect">,
  accent1 = "#4472c4",
): OfficePictureEffects | undefined {
  const own = normalizeOfficePictureEffects(block.shapeEffects);
  if (own) return own;
  const legacy: Partial<Record<OfficePictureEffectKind, unknown>> = {
    ...officeEffectsOfLegacyPreset(block.effect, accent1),
  };
  if (block.shapeShadow) legacy.shadow = officeShadowFromShapeShadow(block.shapeShadow);
  else if (block.shadow === true) legacy.shadow = OFFICE_SWITCH_SHADOW;
  return normalizeOfficePictureEffects(legacy) ?? undefined;
}

/**
 * The fields one Shape Effects pick writes on a shape (`value` null is "No
 * <effect>"): the whole `shapeEffects`, and the legacy switch, custom shadow
 * and preset name cleared, since `shapeEffects` now states them all.
 */
export function presentationShapeEffectPatch<K extends OfficePictureEffectKind>(
  block: Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect">,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
  accent1?: string,
): Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect"> {
  const next = withOfficePictureEffect(presentationShapeEffects(block, accent1), kind, value, false);
  return { shapeEffects: next, shadow: undefined, shapeShadow: undefined, effect: undefined };
}

/**
 * Every targeted shape after one Shape Effects pick, each from its own
 * effects (a multi-selection adds the pick to what each shape already wears,
 * as Office does). Null when nothing changes. Text frames are left alone:
 * their effect list is their shadow.
 */
export function blocksWithShapeEffect<K extends OfficePictureEffectKind>(
  blocks: readonly PresentationBlock[],
  targets: readonly PresentationBlock[],
  kind: K,
  value: OfficePictureEffectValues[K] | null,
  accent1?: string,
): PresentationBlock[] | null {
  const targetIds = new Set(targets.map((block) => block.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (block.type !== "shape" || !targetIds.has(block.id)) return block;
    const patch = presentationShapeEffectPatch(block, kind, value, accent1);
    if (
      JSON.stringify(patch.shapeEffects ?? null) === JSON.stringify(block.shapeEffects ?? null) &&
      block.shadow === undefined &&
      block.shapeShadow === undefined &&
      block.effect === undefined
    ) {
      return block;
    }
    changed = true;
    return shapeBlockWithPatch(block, patch);
  });
  return changed ? next : null;
}

/** The effect kinds a one-name Shape Effects preset (`effect`) draws. */
const LEGACY_PRESET_EFFECT_KINDS = ["reflection", "glow", "softEdge", "bevel", "rotation3d"] as const;













export function presentationShapeEffectsAfterLegacyPick(
  block: Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect">,
  pick: { shadow?: boolean; effect?: string | null },
  accent1?: string,
): Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect"> {
  let effects: OfficePictureEffects | undefined = presentationShapeEffects(block, accent1);
  if (pick.effect !== undefined) {
    if (pick.effect === null) {
      for (const kind of LEGACY_PRESET_EFFECT_KINDS) effects = withOfficePictureEffect(effects, kind, null, false);
    } else {
      const added = officeEffectsOfLegacyPreset(pick.effect, accent1);
      for (const kind of Object.keys(added) as OfficePictureEffectKind[]) {
        effects = withOfficePictureEffect(
          effects,
          kind,
          added[kind] as OfficePictureEffectValues[typeof kind],
          false,
        );
      }
    }
  }
  if (pick.shadow === false) effects = withOfficePictureEffect(effects, "shadow", null, false);
  else if (pick.shadow === true && !effects?.shadow) {
    effects = withOfficePictureEffect(effects, "shadow", OFFICE_SWITCH_SHADOW, false);
  }
  return { shapeEffects: effects, shadow: undefined, shapeShadow: undefined, effect: undefined };
}
