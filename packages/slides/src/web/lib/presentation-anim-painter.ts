/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the ANIMATION PAINTER (PowerPoint "Animation Painter" — the
 * sibling of Format Painter, which copies FORMATTING): copy the animation step(s)
 * of one block and apply them to another. Operates on the slide's flat
 * `animationSequence` (each step carries the `blockId` it animates). Pure +
 * dependency-light so the copy / retarget logic is unit-testable apart from the
 * editor's arm / click plumbing.
 */

import type { TimelineStep } from "@mrnote/core/web/lib/presentation-anim-timeline.ts";

/**
 * The steps animating `sourceBlockId`, copied (a fresh step object per match, with
 * the motion path cloned) so later edits to the source never mutate the captured
 * set. Empty when the source block has no animation.
 */
export function copyAnimationSteps(
  sequence: readonly TimelineStep[],
  sourceBlockId: string,
): TimelineStep[] {
  return sequence
    .filter((step) => step.blockId === sourceBlockId)
    .map((step) => ({
      ...step,
      motionPath: step.motionPath
        ? { ...step.motionPath, waypoints: [...step.motionPath.waypoints] }
        : undefined,
    }));
}

/**
 * Apply `copied` animation steps to `targetBlockId`: drop the target's OWN
 * existing steps (the paint replaces, not stacks), then append the copies
 * retargeted — `blockId` set to the target, a fresh `id`, and a `triggerBlockId`
 * that would now self-reference the target dropped (a self-triggering step is
 * nonsensical; the validation layer drops other stale triggers). The copies keep
 * their relative order + the rest of the sequence is untouched. Returns a new
 * array. A no-op (cloned input) when there is nothing to copy.
 */
export function applyAnimationSteps(
  sequence: readonly TimelineStep[],
  targetBlockId: string,
  copied: readonly TimelineStep[],
  makeId: (index: number) => string,
): TimelineStep[] {
  if (copied.length === 0) return sequence.slice();
  const withoutTarget = sequence.filter((step) => step.blockId !== targetBlockId);
  const retargeted = copied.map((step, index) => {
    const next: TimelineStep = {
      ...step,
      blockId: targetBlockId,
      id: makeId(index),
      motionPath: step.motionPath
        ? { ...step.motionPath, waypoints: [...step.motionPath.waypoints] }
        : undefined,
    };
    if (next.triggerBlockId === targetBlockId) delete next.triggerBlockId;
    return next;
  });
  return [...withoutTarget, ...retargeted];
}
