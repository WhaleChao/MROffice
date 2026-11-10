/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  partitionTimelineSteps,
  type TimelineStep,
} from "@mrnote/core/web/lib/presentation-anim-timeline.ts";

/**
 * One of the tags PowerPoint draws beside an animated object in Normal view
 * while the Animations tab or the Animation Pane is open: the number of the
 * click that plays the effect, so the order of a slide's builds reads off the
 * slide itself. Tags never print and never show in the slide show.
 */
export interface PresentationAnimationTag {
  /**
   * The click that plays it. 0 is the run that plays by itself as the slide
   * opens (a sequence that starts With / After Previous); `null` is an effect
   * a trigger plays (PowerPoint's lightning-bolt tag), outside the clicks.
   */
  click: number | null;
  /**
   * The sequence positions of the object's effects this tag stands for, in
   * order. More than one is PowerPoint's numbered stack: effects of the same
   * object that play on the same click.
   */
  steps: number[];
}

/**
 * Every animated object's tags, in sequence order. Click numbers follow the
 * show: an On Click effect opens the next click; With / After Previous
 * effects take the number of the click they run in; when the first click's
 * run does not start On Click it plays on entering the slide and is 0.
 */
export function presentationAnimationTags(
  sequence: readonly TimelineStep[],
): Map<string, PresentationAnimationTag[]> {
  const tags = new Map<string, PresentationAnimationTag[]>();
  if (sequence.length === 0) return tags;
  const clickOf = new Map<number, number | null>();
  const indexOf = new Map<TimelineStep, number>();
  sequence.forEach((step, index) => indexOf.set(step, index));
  const { autoGroups, triggeredIndices } = partitionTimelineSteps(sequence);
  const first = autoGroups[0]?.steps[0];
  const offset = first && first.trigger !== "on-click" ? 0 : 1;
  autoGroups.forEach((group, groupIndex) => {
    for (const step of group.steps) {
      const index = indexOf.get(step);
      if (index !== undefined) clickOf.set(index, groupIndex + offset);
    }
  });
  for (const index of triggeredIndices) clickOf.set(index, null);
  sequence.forEach((step, index) => {
    if (!clickOf.has(index)) return;
    const click = clickOf.get(index) ?? null;
    const own = tags.get(step.blockId) ?? [];
    const same = own.find((tag) => tag.click === click);
    if (same) same.steps.push(index);
    else own.push({ click, steps: [index] });
    tags.set(step.blockId, own);
  });
  return tags;
}
