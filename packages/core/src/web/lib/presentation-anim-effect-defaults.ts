/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An animation effect's own Effect Options defaults (Smooth start / end it
 * has of its own, whether it already plays text letter by letter, whether
 * it can rewind), read from the timeline catalog and the recipe catalog.
 * Kept apart from `presentation-anim-effect-settings` so the page schema
 * does not load the recipe catalog.
 */

import {
  PRESENTATION_EFFECTS_WITHOUT_REWIND,
  MAX_EFFECT_SHARE,
  type EffectSettingsDefaults,
} from "./presentation-anim-effect-settings";
import { presentationAnimationEffect } from "./presentation-anim-timeline";
import { effectRecipe } from "./presentation-effect-recipes";

/**
 * An effect's defaults: a hand-drawn effect's from the timeline catalog
 * (Stretchy's 100000 smooth start), a recipe-drawn effect's from its recipe
 * (`ac` / `dc` / `it`).
 */
export function effectSettingsDefaults(
  animation: string,
  effectOption: string | undefined,
): EffectSettingsDefaults {
  const recipe = animation === "none" ? null : effectRecipe(animation, effectOption);
  const nativeEffectAccel =
    animation === "none" ? null : (presentationAnimationEffect(animation)?.nativeEffectAccel ?? null);
  return {
    accel: recipe ? Math.round((recipe.ac ?? 0) * MAX_EFFECT_SHARE) : (nativeEffectAccel ?? 0),
    decel: recipe ? Math.round((recipe.dc ?? 0) * MAX_EFFECT_SHARE) : 0,
    iterates: Boolean(recipe?.it),
    rewindable: !PRESENTATION_EFFECTS_WITHOUT_REWIND.has(animation),
  };
}
