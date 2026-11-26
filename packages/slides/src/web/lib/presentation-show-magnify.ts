/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface ShowMagnify {
  /** Magnification, 1 < scale <= SHOW_MAGNIFY_MAX. */
  readonly scale: number;
  /** The slide point at the middle of the view, as fractions of the slide's width and height. */
  readonly cx: number;
  readonly cy: number;
}

/** A click with the glass magnifies to 200 % (Support). */
export const SHOW_MAGNIFY_FIRST = 2;
/** + magnifies up to 400 % (Support). */
export const SHOW_MAGNIFY_MAX = 4;
/** One press of + or - multiplies or divides by this (show-magnify scenario A): 2^(1/3). */
export const SHOW_MAGNIFY_STEP_FACTOR = 2 ** (1 / 3);

const clamp = (value: number, low: number, high: number) => Math.min(high, Math.max(low, value));

/** The view kept on the slide: its middle no nearer an edge than half the view. */
export function clampShowMagnify(state: ShowMagnify): ShowMagnify {
  const scale = clamp(state.scale, 1, SHOW_MAGNIFY_MAX);
  const half = 1 / (2 * scale);
  return { scale, cx: clamp(state.cx, half, 1 - half), cy: clamp(state.cy, half, 1 - half) };
}

/** The glass's click at a slide point: magnified there to the first level. */
export function magnifyAt(cx: number, cy: number, scale: number = SHOW_MAGNIFY_FIRST): ShowMagnify {
  return clampShowMagnify({ scale, cx, cy });
}

/**
 * One press of + (`direction` 1) or - (-1), about the view's middle: the next rung of
 * 2^(k/3) (k = 1..6, 126 % to 400 %) up or down; with nothing magnified, + magnifies the slide's
 * centre to the first rung. Below the first rung the slide is whole again: null.
 */
export function stepShowMagnify(state: ShowMagnify | null, direction: 1 | -1): ShowMagnify | null {
  if (!state) return direction > 0 ? magnifyAt(0.5, 0.5, SHOW_MAGNIFY_STEP_FACTOR) : null;
  // The rung the scale is on (a pinch may leave it between two): + goes to the next one above, - below.
  const rung = Math.log(state.scale) / Math.log(SHOW_MAGNIFY_STEP_FACTOR);
  const next = direction > 0 ? Math.floor(rung + 1e-6) + 1 : Math.ceil(rung - 1e-6) - 1;
  if (next < 1) return null;
  const scale = Math.min(SHOW_MAGNIFY_MAX, SHOW_MAGNIFY_STEP_FACTOR ** next);
  return clampShowMagnify({ ...state, scale });
}

/**
 * A drag of (dx, dy) screen pixels over a slide drawn `slideWidthPx` x `slideHeightPx` on screen
 * (its magnified box): the slide follows the pointer, so the view's middle moves the other way.
 */
export function panShowMagnify(
  state: ShowMagnify,
  dx: number,
  dy: number,
  slideWidthPx: number,
  slideHeightPx: number,
): ShowMagnify {
  if (!(slideWidthPx > 0) || !(slideHeightPx > 0)) return state;
  return clampShowMagnify({ ...state, cx: state.cx - dx / slideWidthPx, cy: state.cy - dy / slideHeightPx });
}

/** One press of an arrow key while magnified pans the view by this share of itself (v3, scenario E2). */
export const SHOW_MAGNIFY_ARROW_PAN = 0.1;

/**
 * An arrow key while magnified: the view moves `dx` / `dy` (-1, 0 or 1) tenths of itself across the
 * slide (Right: the view moves right, the slide left), kept on the slide.
 */
export function panShowMagnifyByArrow(state: ShowMagnify, dx: number, dy: number): ShowMagnify {
  const step = SHOW_MAGNIFY_ARROW_PAN / state.scale;
  return clampShowMagnify({ ...state, cx: state.cx + dx * step, cy: state.cy + dy * step });
}

/**
 * The CSS for the projected slide (centred in its stage, origin at its centre): the slide point
 * (cx, cy) brought to the stage's middle and the slide scaled about it.
 */
export function showMagnifyStyle(state: ShowMagnify | null): { translate: string; scale: string } | null {
  if (!state) return null;
  const tx = -state.scale * (state.cx - 0.5) * 100;
  const ty = -state.scale * (state.cy - 0.5) * 100;
  return { translate: `${round(tx)}% ${round(ty)}%`, scale: `${round(state.scale)}` };
}

const round = (value: number) => Math.round(value * 1000) / 1000;

/** The same state, from a message across windows: anything malformed is no magnification. */
export function parseShowMagnify(value: unknown): ShowMagnify | null {
  if (!value || typeof value !== "object") return null;
  const { scale, cx, cy } = value as Record<string, unknown>;
  if (typeof scale !== "number" || typeof cx !== "number" || typeof cy !== "number") return null;
  if (![scale, cx, cy].every(Number.isFinite) || scale <= 1) return null;
  return clampShowMagnify({ scale, cx, cy });
}
