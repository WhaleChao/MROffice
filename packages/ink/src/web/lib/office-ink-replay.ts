/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink Replay (XLS-A032): one timeline for the ink on a surface.
 *
 * Microsoft Support, "Replay your ink strokes in Office": the app removes the
 * ink and redraws it in the order it was drawn; the replay can be paused and
 * started again from any point. A stroke that recorded when each of its points
 * was drawn plays back at that pace; a stroke drawn before timing was kept has
 * none, so it is drawn at one constant speed. Strokes follow one another with a
 * short pause, whatever time passed between them while drawing.
 *
 * Host-neutral: strokes are point lists in any one pixel space.
 */

export interface InkReplayStrokeInput {
  readonly id: string;
  /** The stroke's points in drawing order (pixels). */
  readonly points: readonly { readonly x: number; readonly y: number }[];
  /** When each point was drawn (ms from the first), or null for none. */
  readonly times?: readonly number[] | null;
}

export interface InkReplayStroke {
  readonly id: string;
  /** When this stroke starts on the timeline (ms). */
  readonly start: number;
  /** When each point is reached, on the timeline (ms). */
  readonly at: readonly number[];
}

export interface InkReplayTimeline {
  readonly strokes: readonly InkReplayStroke[];
  /** The whole replay's length (ms). */
  readonly duration: number;
}

/** The constant pace of an untimed stroke, pixels per millisecond (480 px/s). */
export const INK_REPLAY_PX_PER_MS = 0.48;
/** The pause between one stroke's end and the next one's start. */
export const INK_REPLAY_STROKE_GAP_MS = 160;
/** A recorded pause inside one stroke never holds the replay longer than this. */
const MAX_RECORDED_STEP_MS = 1_000;

export function inkReplayTimeline(strokes: readonly InkReplayStrokeInput[]): InkReplayTimeline {
  const out: InkReplayStroke[] = [];
  let clock = 0;
  for (const stroke of strokes) {
    if (stroke.points.length === 0) continue;
    const timed =
      stroke.times && stroke.times.length === stroke.points.length ? stroke.times : null;
    const at: number[] = [clock];
    for (let index = 1; index < stroke.points.length; index += 1) {
      const step = timed
        ? Math.min(MAX_RECORDED_STEP_MS, Math.max(0, timed[index]! - timed[index - 1]!))
        : Math.hypot(
            stroke.points[index]!.x - stroke.points[index - 1]!.x,
            stroke.points[index]!.y - stroke.points[index - 1]!.y,
          ) / INK_REPLAY_PX_PER_MS;
      at.push(at[index - 1]! + step);
    }
    out.push({ id: stroke.id, start: clock, at });
    clock = at[at.length - 1]! + INK_REPLAY_STROKE_GAP_MS;
  }
  return { strokes: out, duration: out.length > 0 ? clock - INK_REPLAY_STROKE_GAP_MS : 0 };
}

/**
 * How many of each stroke's points are drawn at `position`: 0 for a stroke
 * not yet begun, all of them for a finished one. A stroke under way shows at
 * least its first two points, so its pen tip is visible from the start.
 */
export function inkReplayFrame(
  timeline: InkReplayTimeline,
  position: number,
): ReadonlyMap<string, number> {
  const frame = new Map<string, number>();
  for (const stroke of timeline.strokes) {
    if (position < stroke.start) {
      frame.set(stroke.id, 0);
      continue;
    }
    let count = 0;
    while (count < stroke.at.length && stroke.at[count]! <= position) count += 1;
    frame.set(stroke.id, Math.min(stroke.at.length, Math.max(2, count)));
  }
  return frame;
}

/** The start of the stroke before / after `position` (the Rewind and Forward buttons). */
export function inkReplayStep(
  timeline: InkReplayTimeline,
  position: number,
  direction: -1 | 1,
): number {
  const starts = timeline.strokes.map((stroke) => stroke.start);
  if (direction > 0) {
    return starts.find((start) => start > position) ?? timeline.duration;
  }
  // Back from inside a stroke goes to that stroke's start; from (just after)
  // a start, to the one before it.
  const before = starts.filter((start) => start < position - 120);
  return before.length > 0 ? before[before.length - 1]! : 0;
}
