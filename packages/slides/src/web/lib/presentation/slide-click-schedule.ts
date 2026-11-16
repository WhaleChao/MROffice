/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { computeStepStartOffsets, resolveTimelineGroups, type TimelineStep } from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import { slideNativeTimingTree } from "../presentation-native-timing-player";

type ScheduledSlide = {
  readonly animationSequence?: readonly TimelineStep[];
  readonly nativeTimingTree?: unknown;
  readonly rehearsedClickTimingsMs?: readonly number[];
};

/**
 * How long each of a slide's groups (its first, automatic one included)
 * builds: a timing tree's run lengths, or each group's latest end. A
 * repeating effect counts one pass.
 */
export function slideGroupBuildsMs(
  slide: ScheduledSlide,
  playDurationMs: (step: TimelineStep) => number = (step) => step.durationMs,
): number[] {
  const tree = slideNativeTimingTree(slide as Parameters<typeof slideNativeTimingTree>[0]);
  if (tree) return tree.clicks.map((click) => (Number.isFinite(click.length) ? Math.max(0, click.length) : 0));
  return resolveTimelineGroups(slide.animationSequence ?? []).map((group) => {
    const offsets = computeStepStartOffsets(group, playDurationMs);
    let end = 0;
    group.steps.forEach((step, index) => {
      end = Math.max(end, (offsets[index] ?? 0) + Math.max(0, Number(playDurationMs(step)) || 0));
    });
    return end;
  });
}

/** A rehearsal that times every one of `clicks` clicks, or null. */
export function rehearsedClickDelaysMs(slide: ScheduledSlide, clicks: number): number[] | null {
  const timings = slide.rehearsedClickTimingsMs;
  if (!Array.isArray(timings) || clicks <= 0 || timings.length !== clicks) return null;
  if (!timings.every((value) => typeof value === "number" && Number.isFinite(value) && value >= 0)) return null;
  return timings.slice();
}

export type TimedClickPlan = {
  /** Each click's delay after the click before it (the first: after the slide's timeline starts), ms. */
  readonly delaysMs: readonly number[];
  /** How long the slide stays after its last click, ms. */
  readonly afterLastMs: number;
  readonly rehearsed: boolean;
};

/**
 * When the slide's `clicks` clicks come by themselves, for a slide that is up
 * `slideMs` (its After, or Seconds spent on each slide): rehearsed when
 * `rehearsed` (the slide's rehearsal, by {@link rehearsedClickDelaysMs}),
 * else shared. `groupBuildsMs` is {@link slideGroupBuildsMs}; `initialAuto`
 * whether its first group plays by itself (then click k reveals group k).
 */
export function timedClickPlan({
  slideMs,
  clicks,
  groupBuildsMs,
  initialAuto,
  rehearsed,
}: {
  slideMs: number;
  clicks: number;
  groupBuildsMs: readonly number[];
  initialAuto: boolean;
  rehearsed: readonly number[] | null;
}): TimedClickPlan {
  const build = (group: number) => {
    const value = groupBuildsMs[group];
    return typeof value === "number" && Number.isFinite(value) ? Math.max(0, value) : 0;
  };
  // The group a click reveals, and the one shown before it.
  const revealed = (click: number) => (initialAuto ? click : click - 1);
  const lastBuild = clicks > 0 ? build(revealed(clicks)) : initialAuto ? build(0) : 0;
  const asked = Math.max(0, Number.isFinite(slideMs) ? slideMs : 0);
  if (rehearsed && rehearsed.length === clicks) {
    const total = rehearsed.reduce((sum, value) => sum + value, 0);
    return { delaysMs: rehearsed.slice(), afterLastMs: Math.max(lastBuild, asked - total), rehearsed: true };
  }
  let builds = 0;
  for (let group = 0; group < groupBuildsMs.length; group += 1) builds += build(group);
  const share = Math.max(0, asked - builds) / (Math.max(0, clicks) + 1);
  const delaysMs: number[] = [];
  for (let click = 1; click <= clicks; click += 1) {
    const before = revealed(click) - 1;
    delaysMs.push((before >= 0 ? build(before) : 0) + share);
  }
  return { delaysMs, afterLastMs: lastBuild + share, rehearsed: false };
}
