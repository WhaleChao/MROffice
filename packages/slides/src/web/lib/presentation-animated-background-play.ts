/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Animate as Background in the show (PPT-A134): PowerPoint
 * for Mac's timing played through the effect engine
 * (`presentation-effect-recipes`) -- pan tracks along each stage's path and
 * two scale tracks, 30 s each with PowerPoint's 50 % accelerate /
 * decelerate, a 5 s pause before every later stage and a 5 s hold before it
 * all repeats.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  effectRecipeKeyframeBody,
  type EffectRecipeBox,
  type EffectRecipeVariant,
} from "@mrnote/core/web/lib/presentation-effect-recipes.ts";
import {
  normalizePresentationAnimatedBackground,
  presentationAnimatedBackgroundPan,
  type PresentationAnimatedBackground,
} from "@mrnote/core/web/lib/presentation-animated-background.ts";

const STAGE_MS = 30_000;
const GAP_MS = 5_000;
const HOLD_MS = 5_000;
const ZOOM = 1.5;

/** The whole loop's length: the stages, their pauses and the hold. */
export function presentationAnimatedBackgroundDurationMs(points: number): number {
  return STAGE_MS + (points - 1) * (GAP_MS + STAGE_MS) + GAP_MS + STAGE_MS + HOLD_MS;
}

const path = (from: [number, number], to: [number, number]) =>
  `M ${from[0]} ${from[1]} L ${to[0]} ${to[1]}`;

/** The loop as the effect engine's recipe (one object, the picture). */
export function presentationAnimatedBackgroundRecipe(
  background: PresentationAnimatedBackground,
  slide: { width: number; height: number },
): EffectRecipeVariant {
  const pans = background.focalPoints.map(
    (point): [number, number] => [
      presentationAnimatedBackgroundPan(point.x, slide.width),
      presentationAnimatedBackgroundPan(point.y, slide.height),
    ],
  );
  const total = presentationAnimatedBackgroundDurationMs(pans.length);
  const at = (ms: number) => ms / total;
  const stage = { d: at(STAGE_MS), ac: 0.5, dc: 0.5, hold: 1 as const };
  const tracks: EffectRecipeVariant["tracks"] = [
    { k: "scale", from: [1, 1], to: [ZOOM, ZOOM], b: 0, ...stage },
    { k: "motion", path: path([0, 0], pans[0]), b: 0, ...stage },
  ];
  let start = STAGE_MS;
  for (let index = 1; index < pans.length; index += 1) {
    tracks.push({ k: "motion", path: path(pans[index - 1], pans[index]), b: at(start + GAP_MS), ...stage });
    start += GAP_MS + STAGE_MS;
  }
  tracks.push(
    { k: "motion", path: path(pans[pans.length - 1], [0, 0]), b: at(start + GAP_MS), ...stage },
    { k: "scale", from: [ZOOM, ZOOM], to: [1, 1], b: at(start + GAP_MS), ...stage },
  );
  return { d: total, tracks };
}

/** The keyframe name of a picture's loop (stable for the same points and slide). */
export function presentationAnimatedBackgroundKeyframe(
  background: PresentationAnimatedBackground,
  slide: { width: number; height: number },
): string {
  const key = background.focalPoints
    .map((point) => `${Math.round(point.x * 1e4)}_${Math.round(point.y * 1e4)}`)
    .join("-");
  return `mrai-ppt-animbg-${slide.width}x${slide.height}-${key}`.replace(/[^A-Za-z0-9_-]/g, "_");
}

export interface PresentationAnimatedBackgroundPlay {
  readonly blockId: string;
  readonly keyframe: string;
  readonly durationMs: number;
  /** The `@keyframes` rule. */
  readonly css: string;
}

/**
 * The slide's animated background in the show: the first picture that
 * carries one (PowerPoint keeps one per slide), its keyframes and length.
 */
export function presentationAnimatedBackgroundPlay(
  blocks: readonly PresentationBlock[],
  slide: { width: number; height: number },
  box: (block: PresentationBlock) => EffectRecipeBox,
): PresentationAnimatedBackgroundPlay | null {
  for (const block of blocks) {
    if (block.type !== "image") continue;
    const background = normalizePresentationAnimatedBackground(
      (block as { animatedBackground?: unknown }).animatedBackground,
    );
    if (!background) continue;
    const recipe = presentationAnimatedBackgroundRecipe(background, slide);
    const keyframe = presentationAnimatedBackgroundKeyframe(background, slide);
    // Enough samples that each 30 s stage's ease is drawn smoothly.
    const steps = Math.max(40, Math.ceil(recipe.d / 1000));
    return {
      blockId: block.id,
      keyframe,
      durationMs: recipe.d,
      css: `@keyframes ${keyframe}{${effectRecipeKeyframeBody(recipe, "emph", box(block), { steps })}}`,
    };
  }
  return null;
}

