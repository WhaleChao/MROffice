/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-slide HIDE / SKIP support (PowerPoint / WPS "Hide Slide" parity).
 *
 * A slide may carry an optional `hidden: true` flag. A hidden slide is kept in
 * the deck (it still edits + serializes) but is SKIPPED during a slideshow:
 * present-mode advance / retreat / jump land on the next VISIBLE slide instead.
 *
 * Pure, side-effect-free, dependency-free (no React / no DOM) so it can be unit
 * tested in isolation and shared by the editor, the present overlay and the
 * schema. All navigation here operates on INDICES into the full slide array (the
 * editor keeps a single canonical `slides[]` and `presentIndex` indexes into it;
 * hiding only changes which indices a slideshow stops on, never the array order
 * or the slide ids — so ink / triggers / `slide://` links keyed by id are
 * unaffected).
 */

/** The minimal slide shape this module needs: just the optional hide flag. */
export interface HideableSlide {
  hidden?: unknown;
  [key: string]: unknown;
}

/**
 * Normalize an arbitrary persisted `hidden` value to a strict boolean flag, or
 * `undefined` when the slide is visible. Only an explicit `true` hides a slide;
 * every other value (absent / `false` / non-boolean) yields `undefined` so a
 * visible slide serializes byte-clean (no `hidden` key at all). Mirrors the
 * "drop a false flag" convention used by the other additive slide fields.
 */
export function normalizeHiddenFlag(value: unknown): true | undefined {
  return value === true ? true : undefined;
}

/** True when a slide is explicitly hidden (skipped in a slideshow). */
export function isSlideHidden(slide: HideableSlide | null | undefined): boolean {
  return !!slide && slide.hidden === true;
}

/**
 * The index of the first VISIBLE slide at or after `from` (inclusive), or `-1`
 * when none exists from there on. `from` is clamped to `[0, length]`; a negative
 * `from` starts at 0. Used to resolve a jump target onto a presentable slide.
 */
export function firstVisibleFrom(
  slides: readonly HideableSlide[],
  from: number,
): number {
  const start = Math.max(0, Math.floor(Number.isFinite(from) ? from : 0));
  for (let i = start; i < slides.length; i += 1) {
    if (!isSlideHidden(slides[i])) return i;
  }
  return -1;
}

/**
 * The index of the last VISIBLE slide at or before `from` (inclusive), or `-1`
 * when none exists up to there. `from` is clamped to `[-1, length-1]`. Used to
 * resolve a backward jump onto a presentable slide.
 */
export function lastVisibleAtOrBefore(
  slides: readonly HideableSlide[],
  from: number,
): number {
  const start = Math.min(
    slides.length - 1,
    Math.floor(Number.isFinite(from) ? from : slides.length - 1),
  );
  for (let i = start; i >= 0; i -= 1) {
    if (!isSlideHidden(slides[i])) return i;
  }
  return -1;
}

/**
 * The index of the next VISIBLE slide STRICTLY after `current`, or `-1` when the
 * current slide is the last visible one (advance is then a no-op — the show
 * holds on the final visible slide, matching PowerPoint). Hidden slides between
 * `current` and the next visible one are skipped.
 */
export function nextVisibleIndex(
  slides: readonly HideableSlide[],
  current: number,
): number {
  return firstVisibleFrom(slides, Math.floor(current) + 1);
}

/**
 * The index of the previous VISIBLE slide STRICTLY before `current`, or `-1`
 * when there is no earlier visible slide (retreat is then a no-op). Hidden
 * slides between the previous visible one and `current` are skipped.
 */
export function prevVisibleIndex(
  slides: readonly HideableSlide[],
  current: number,
): number {
  return lastVisibleAtOrBefore(slides, Math.floor(current) - 1);
}

/**
 * The starting present index for a slideshow: the first visible slide AT or
 * AFTER `desired` (typically the editor's active index), else searching from the
 * deck start, else `desired` clamped — so a slideshow always has a slide to show
 * even in the (schema-prevented) degenerate case of an all-hidden deck. Pure.
 */
export function presentStartIndex(
  slides: readonly HideableSlide[],
  desired: number,
): number {
  if (slides.length === 0) return 0;
  const clampedDesired = Math.min(
    Math.max(0, Math.floor(Number.isFinite(desired) ? desired : 0)),
    slides.length - 1,
  );
  const ahead = firstVisibleFrom(slides, clampedDesired);
  if (ahead !== -1) return ahead;
  const fromStart = firstVisibleFrom(slides, 0);
  if (fromStart !== -1) return fromStart;
  // Degenerate all-hidden deck (the editor/schema keep ≥1 visible, so this is a
  // belt-and-braces fallback): show the desired slide anyway.
  return clampedDesired;
}

/** The number of VISIBLE (presentable) slides in the deck. */
export function visibleSlideCount(slides: readonly HideableSlide[]): number {
  let count = 0;
  for (const slide of slides) if (!isSlideHidden(slide)) count += 1;
  return count;
}
