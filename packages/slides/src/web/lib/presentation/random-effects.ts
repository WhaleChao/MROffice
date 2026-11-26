/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { resolvePresentationAnimationEffectOption, type TimelineStep } from "@mrnote/core/web/lib/presentation-anim-timeline.ts";









export const RANDOM_EFFECT_PICKS = [
  { animation: "plus-in", effectOption: "in", durationMs: 2000 },
  { animation: "wipe-in", effectOption: "from-bottom", durationMs: 500 },
  { animation: "wheel-in", effectOption: "2-spokes", durationMs: 2000 },
  { animation: "blinds-in", effectOption: "horizontal", durationMs: 500 },
  { animation: "fade", effectOption: undefined, durationMs: 500 },
] as const satisfies readonly {
  animation: TimelineStep["animation"];
  effectOption: TimelineStep["effectOption"];
  durationMs: number;
}[];

/** A new show's Random Effects seed: set once when the show starts. */
export function newRandomEffectSeed(): number {
  return Math.floor(Math.random() * 0x7fffffff);
}

/** The seed for one slide of a show: the show's seed mixed with the slide. */
export function slideRandomEffectSeed(showSeed: number, slideId: string): number {
  let hash = showSeed >>> 0;
  for (let index = 0; index < slideId.length; index += 1) {
    hash = Math.imul(hash ^ slideId.charCodeAt(index), 0x01000193) >>> 0;
  }
  return hash;
}

/**
 * The steps as one show plays them: each Random Effects step becomes one of
 * the picks, chosen by `seed` and its place. The show's seed is set once
 * when the show starts and shared with every window of it (the presenter
 * sync snapshot), so they all play the same pick; the next show picks anew,
 * as PowerPoint's does.
 */
export function resolveRandomEffectSteps<T extends TimelineStep>(steps: readonly T[], seed: number): T[] {
  return steps.map((step, index) => {
    if (step.animation !== "random-effects") return step;
    const mixed = Math.imul((seed ^ Math.imul(index + 1, 0x9e3779b1)) >>> 0, 0x85ebca6b) >>> 0;
    const pick = RANDOM_EFFECT_PICKS[mixed % RANDOM_EFFECT_PICKS.length]!;
    const played = { ...step, animation: pick.animation, durationMs: pick.durationMs } as T;
    // The option as every other step carries it: through the catalog's own
    // resolution, which states no option for the effect's default (Wipe From
    // Bottom, Blinds Horizontal) -- the keyframes play by that rule.
    const option = pick.effectOption === undefined
      ? undefined
      : resolvePresentationAnimationEffectOption(pick.animation, pick.effectOption)?.effectOption;
    if (option === undefined) delete (played as { effectOption?: unknown }).effectOption;
    else (played as { effectOption?: unknown }).effectOption = option;
    return played;
  });
}
