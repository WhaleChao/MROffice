/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Discrete 90° "Rotate Right / Rotate Left" quantization for presentation blocks
 * (PowerPoint / WPS "Arrange → Rotate → Rotate Right 90° / Rotate Left 90°").
 *
 * Dependency-free (no React / no DOM) so it can be unit-tested in isolation and
 * shared by the editor component + tests. It layers on top of the existing
 * rotation model in `presentation-drag.ts` (`normalizeRotationDegrees` keeps
 * every angle in `[0, 360)`); a block's `rotation` is the same degree field the
 * free-drag rotate handle writes, so a quarter-turn composes cleanly with a
 * hand-rotated block.
 *
 * The two operations match PowerPoint exactly:
 *  - Rotate RIGHT (clockwise) adds 90°.
 *  - Rotate LEFT (counter-clockwise) subtracts 90°.
 *
 * A block that is currently at an OFF-grid angle (e.g. 37°, set by the free
 * rotate handle) is first SNAPPED to the nearest 90° multiple, then the ±90° step
 * is applied — so the first click lands the block on a clean quarter and the next
 * click steps a full quarter from there (PowerPoint / WPS behaviour). The result
 * is always one of `0 / 90 / 180 / 270`.
 */

import { normalizeRotationDegrees } from "./presentation-drag";

/** The discrete rotation step, in degrees. */
export const QUARTER_TURN_DEGREES = 90;

/** Direction of a discrete quarter-turn. `right` = clockwise (+90°). */
export type QuarterTurnDirection = "left" | "right";

/** The four on-grid rotation values a quarter-turn ever produces. */
export const QUARTER_ANGLES = [0, 90, 180, 270] as const;

export type QuarterAngle = (typeof QUARTER_ANGLES)[number];

/**
 * Snap any degree value to the NEAREST multiple of 90, normalized into the four
 * canonical quarter angles `{0, 90, 180, 270}`. A value exactly halfway between
 * two quarters rounds up (e.g. 45° → 90°, matching `Math.round`'s half-up rule).
 * Non-finite input collapses to 0. Pure + deterministic.
 *
 * 360° (and any exact multiple of 360) normalizes to 0 — never returns 360.
 */
export function snapToQuarter(degrees: number): QuarterAngle {
  if (!Number.isFinite(degrees)) return 0;
  const snapped = Math.round(degrees / QUARTER_TURN_DEGREES) * QUARTER_TURN_DEGREES;
  // normalizeRotationDegrees keeps it in [0, 360); a snapped multiple of 90 is
  // therefore always exactly one of 0 / 90 / 180 / 270.
  return normalizeRotationDegrees(snapped) as QuarterAngle;
}

/**
 * Rotate `current` by one quarter-turn in `direction`, as PowerPoint's Rotate
 * Right 90° / Rotate Left 90° do: 90° is ADDED to the current angle (a
 * hand-rotated 30° shape becomes 120°, then 30° again), never snapped to the
 * nearest right angle first (PPT-B229). The result is normalized to [0, 360).
 *
 * Pure + deterministic; never mutates input; non-finite `current` is treated as 0.
 */
export function quarterTurn(
  current: number | null | undefined,
  direction: QuarterTurnDirection,
): number {
  const base = Number.isFinite(current as number) ? (current as number) : 0;
  const delta = direction === "right" ? QUARTER_TURN_DEGREES : -QUARTER_TURN_DEGREES;
  return normalizeRotationDegrees(base + delta);
}
