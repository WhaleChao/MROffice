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
  PresentationTextBlock,
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
import { presentationTextFrameHasShapeAuthority, shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { presentationBlockWithMaterializedPlaceholder } from "@mrnote/core/web/lib/presentation-shape-style-adapter.ts";

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
 * `styleHasEffect`: a style the shape follows draws that effect (a SmartArt
 * Style's bevel), so "No <effect>" states `"none"` to take the style's off
 * too, as Office's own empty statement does.
 */
export function presentationShapeEffectPatch<K extends OfficePictureEffectKind>(
  block: Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect">,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
  accent1?: string,
  styleHasEffect = false,
): Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect"> {
  const next = withOfficePictureEffect(presentationShapeEffects(block, accent1), kind, value, styleHasEffect);
  return { shapeEffects: next, shadow: undefined, shapeShadow: undefined, effect: undefined };
}

/**
 * A text box after a Format Shape > Effects edit: a text box is a shape
 * (`p:sp txBox="1"`), and PowerPoint gives it the shape's whole Effects page,
 * held in the shape's own fields (`shapeEffects`, `keepTextFlat`; the older
 * soft-shadow switch `shadow` is cleared by an effects edit, which states
 * it). An undefined value takes the field off. Null when nothing changes or
 * the frame holds no shape authority; an inherited placeholder is
 * materialised, as every Shape Format edit on one is.
 */
export function presentationTextFrameWithEffectsPatch(
  block: PresentationTextBlock,
  patch: Partial<Pick<PresentationTextBlock, "shapeEffects" | "shadow" | "keepTextFlat">>,
): PresentationTextBlock | null {
  if (!presentationTextFrameHasShapeAuthority(block)) return null;
  const next = { ...block } as Record<string, unknown>;
  for (const key of ["shapeEffects", "shadow", "keepTextFlat"] as const) {
    if (!(key in patch)) continue;
    if (patch[key] === undefined || patch[key] === null) delete next[key];
    else next[key] = patch[key];
  }
  const record = block as Record<string, unknown>;
  if (
    ["shapeEffects", "shadow", "keepTextFlat"].every(
      (key) => JSON.stringify(record[key] ?? null) === JSON.stringify(next[key] ?? null),
    )
  ) {
    return null;
  }
  return presentationBlockWithMaterializedPlaceholder(next as unknown as PresentationTextBlock);
}

/**
 * A shape or a text box with 3-D Rotation > Keep text flat checked or not
 * (`keepTextFlat`, the body's `a:flatTx`): the one field Shape Options'
 * Effects and Text Options' Text Effects both set. Null when nothing changes.
 */
export function presentationTextFrameWithKeepTextFlat(
  block: PresentationShapeBlock | PresentationTextBlock,
  checked: boolean,
): PresentationShapeBlock | PresentationTextBlock | null {
  if ((block.keepTextFlat === true) === checked) return null;
  if (block.type === "text") return presentationTextFrameWithEffectsPatch(block, { keepTextFlat: checked ? true : undefined });
  const next = shapeBlockWithPatch(block, { keepTextFlat: checked ? true : undefined });
  return next === block ? null : next;
}

/**
 * Every targeted shape and text box after one Shape Effects pick, each from
 * its own effects (a multi-selection adds the pick to what each one already
 * wears, as Office does). Null when nothing changes.
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
    if (!targetIds.has(block.id)) return block;
    if (block.type === "text") {
      // A text box wears the shape's Effects in the same fields.
      const { shapeEffects, shadow } = presentationShapeEffectPatch(
        block as unknown as Pick<PresentationShapeBlock, "shapeEffects" | "shadow" | "shapeShadow" | "effect">,
        kind,
        value,
        accent1,
      );
      const patched = presentationTextFrameWithEffectsPatch(block, { shapeEffects, shadow });
      if (!patched) return block;
      changed = true;
      return patched;
    }
    if (block.type !== "shape") return block;
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
