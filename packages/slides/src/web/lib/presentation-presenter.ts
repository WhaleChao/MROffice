/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presenter view (speaker console) pure helpers.
 *
 * Side-effect-free, dependency-free (no React / no DOM) logic shared by the
 * present-mode overlay's presenter layout and its unit tests. Keeping the
 * elapsed-timer formatting, next-slide derivation, and the slide-counter copy
 * here means the overlay component stays a thin render of already-computed
 * values, and the tricky bits (clock formatting, clamping) can be exercised in
 * isolation.
 */

const MS_PER_SECOND = 1000;
const SECONDS_PER_MINUTE = 60;
const MINUTES_PER_HOUR = 60;

/** Two-digit zero-pad for the minutes / seconds segments of the clock. */
function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/**
 * Format an elapsed duration (in milliseconds) as a wall-clock string.
 *
 * Under one hour it renders `MM:SS` (e.g. `07:05`); at or beyond one hour it
 * renders `H:MM:SS` (e.g. `1:02:09`), matching how PowerPoint's presenter timer
 * grows a leading hours segment only when needed. Negative or non-finite inputs
 * (a clock that ran backwards, NaN from a bad subtraction) clamp to `00:00` so
 * the console never shows a garbage timer.
 */
export function formatElapsed(ms: number): string {
  if (!Number.isFinite(ms) || ms <= 0) return "00:00";
  const totalSeconds = Math.floor(ms / MS_PER_SECOND);
  const seconds = totalSeconds % SECONDS_PER_MINUTE;
  const totalMinutes = Math.floor(totalSeconds / SECONDS_PER_MINUTE);
  const minutes = totalMinutes % MINUTES_PER_HOUR;
  const hours = Math.floor(totalMinutes / MINUTES_PER_HOUR);
  if (hours > 0) {
    return `${hours}:${pad2(minutes)}:${pad2(seconds)}`;
  }
  return `${pad2(minutes)}:${pad2(seconds)}`;
}

/**
 * The index of the slide that follows `current` in a deck of `total` slides, or
 * `null` when `current` is already the last slide (so the presenter console can
 * render an "end of deck" placeholder instead of a next-slide preview). Inputs
 * are clamped defensively: a non-positive `total`, or a `current` outside
 * `[0, total)`, yields `null` rather than an out-of-range index.
 */
export function nextSlideIndex(current: number, total: number): number | null {
  if (!Number.isFinite(current) || !Number.isFinite(total)) return null;
  if (total <= 0) return null;
  const clamped = Math.max(0, Math.min(Math.floor(current), total - 1));
  const next = clamped + 1;
  return next < total ? next : null;
}

/**
 * The 1-based "current of total" pair the counter label is built from. Mirrors
 * the audience counter's data attributes so both surfaces report the same
 * numbers. `current` is clamped into `[0, total)` (then shown 1-based) and
 * `total` is floored at zero, so a transiently-empty deck reports `0/0` rather
 * than negative or fractional values.
 */
export function presenterCounterLabel(
  current: number,
  total: number,
): { current: number; total: number } {
  const safeTotal = Number.isFinite(total) ? Math.max(0, Math.floor(total)) : 0;
  if (safeTotal === 0) return { current: 0, total: 0 };
  const clamped = Math.max(
    0,
    Math.min(Number.isFinite(current) ? Math.floor(current) : 0, safeTotal - 1),
  );
  return { current: clamped + 1, total: safeTotal };
}
