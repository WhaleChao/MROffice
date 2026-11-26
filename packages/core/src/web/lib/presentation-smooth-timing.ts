/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Smooth start / Smooth end (Effect Options > Timing; ECMA-376
 * `p:cTn/@accel`, `@decel`, which take SMIL 2.0's accelerate / decelerate):
 * one progress mapping for every player of the deck.
 *
 * SMIL 2.0 Time Manipulations: `accelerate` and `decelerate` are shares of the
 * simple duration (their sum at most 1). Over the accelerate share the speed
 * rises evenly from 0 to the run rate r = 1 / (1 - a/2 - d/2), holds r in the
 * middle, and falls evenly back to 0 over the decelerate share -- the area
 * under the speed is 1, so the effect still ends on time. VBA's
 * Timing.Accelerate: "a value of 0.9 means that an acceleration should start
 * slower than the default speed for 90% of the total animation time".
 * PowerPoint shortens the two when their sum would pass the duration; so does
 * this (the decelerate share is cut to what is left).
 *
 * The mapping is of the WHOLE effect: progress p = E(t) over its duration.
 * A CSS `animation-timing-function` (or a Web Animations keyframe easing)
 * applies between two keyframes, so an effect drawn with more than two
 * keyframes -- a motion path, a blinds or checkerboard reveal -- eased that
 * way slows down and speeds up again at every keyframe and runs as a whole at
 * an even pace. Such an effect bakes E in: its frames are laid at even times
 * at the place E puts them (`smoothFramesAtEvenTimes`), or its keyframes are
 * moved to the times E reaches them, each carrying the piece of E up to the
 * next (`smoothFrameTimes`, `smoothKeyframeBody`), and it runs linearly.
 */

let linearEasingSupport: boolean | null = null;

/**
 * Whether this browser takes the CSS `linear()` easing (Chrome 113, Safari
 * 17.2, Firefox 112 on). Where it does not, a `linear()` timing function is an
 * invalid value: an `animation` shorthand carrying it is dropped whole, and a
 * Web Animation given it as its easing throws. Without a CSS engine to ask
 * (the server's render, a test) it is taken as supported, as every browser
 * the product ships to answers.
 */
export function supportsLinearEasing(): boolean {
  if (linearEasingSupport === null) {
    try {
      const css = (globalThis as { CSS?: { supports?: (property: string, value: string) => boolean } }).CSS;
      linearEasingSupport = typeof css?.supports === "function" ? css.supports("animation-timing-function", "linear(0, 1)") : true;
    } catch {
      linearEasingSupport = true;
    }
  }
  return linearEasingSupport;
}

/** Tests: answer `supportsLinearEasing` with `value` (null asks the browser again). */
export function setLinearEasingSupportForTests(value: boolean | null): void {
  linearEasingSupport = value;
}

/**
 * Where `linear()` is not taken: the cubic Bézier nearest SMIL's curve of
 * the simple kind (its control points two thirds into each ramp; at most
 * 0.0625 off the exact progress over accel / decel from 0 to 1).
 */
function smoothEasingCubic(accel: number, decel: number): string {
  const { accel: a, decel: d } = smoothShares(accel, decel);
  return `cubic-bezier(${round4((2 * a) / 3)}, 0, ${round4(1 - (2 * d) / 3)}, 1)`;
}

/** The two shares as SMIL takes them: each within 0..1, the sum at most 1. */
export function smoothShares(accel: number, decel: number): { accel: number; decel: number } {
  const a = Number.isFinite(accel) ? Math.max(0, Math.min(1, accel)) : 0;
  const d = Number.isFinite(decel) ? Math.max(0, Math.min(1 - a, decel)) : 0;
  return { accel: a, decel: d };
}

/** Whether the shares change anything (both 0 is an even pace). */
export function isSmooth(accel: number, decel: number): boolean {
  const shares = smoothShares(accel, decel);
  return shares.accel > 0 || shares.decel > 0;
}

/** E(t): the progress at time `t` (both 0..1) under Smooth start / end. */
export function smoothProgress(t: number, accel: number, decel: number): number {
  const { accel: a, decel: d } = smoothShares(accel, decel);
  const time = Math.max(0, Math.min(1, t));
  if (a === 0 && d === 0) return time;
  const rate = 1 / (1 - a / 2 - d / 2);
  if (time < a) return (rate * time * time) / (2 * a);
  if (time > 1 - d) {
    const rest = 1 - time;
    return 1 - (rate * rest * rest) / (2 * d);
  }
  return rate * (time - a / 2);
}

/** E⁻¹(p): the time (0..1) at which the effect reaches progress `p`. */
export function smoothTimeAt(progress: number, accel: number, decel: number): number {
  const p = Math.max(0, Math.min(1, progress));
  if (!isSmooth(accel, decel) || p === 0 || p === 1) return p;
  let low = 0;
  let high = 1;
  for (let step = 0; step < 48; step += 1) {
    const middle = (low + high) / 2;
    if (smoothProgress(middle, accel, decel) < p) low = middle;
    else high = middle;
  }
  return (low + high) / 2;
}

const round4 = (value: number) => {
  const rounded = Math.round(value * 10_000) / 10_000;
  return Object.is(rounded, -0) ? 0 : rounded;
};

/**
 * A `linear()` easing that follows `progress` (0..1 -> 0..1) at `steps + 1`
 * even times.
 */
function linearEasingOf(progress: (t: number) => number, steps: number): string {
  const samples: string[] = [];
  for (let index = 0; index <= steps; index += 1) {
    const t = index / steps;
    samples.push(`${round4(Math.min(1, Math.max(0, progress(t))))} ${Math.round(t * 1000) / 10}%`);
  }
  return `linear(${samples.join(", ")})`;
}

/**
 * The CSS easing of Smooth start / end for an effect drawn by two keyframes
 * (its start and its end): E sampled into a `linear()` -- or, where the
 * browser has no `linear()`, the nearest cubic Bézier. Null for an even pace.
 */
export function smoothEasingCss(accel: number, decel: number): string | null {
  if (!isSmooth(accel, decel)) return null;
  if (!supportsLinearEasing()) return smoothEasingCubic(accel, decel);
  return linearEasingOf((t) => smoothProgress(t, accel, decel), 20);
}

/**
 * `count + 1` frames at even times 0, 1/count, ... 1, each at the progress E
 * gives that time (`at` turns a progress into the frame's value): played
 * linearly, they run the whole effect under Smooth start / end.
 */
export function smoothFramesAtEvenTimes<T>(
  count: number,
  accel: number,
  decel: number,
  at: (progress: number) => T,
): Array<{ offset: number; value: T }> {
  const frames: Array<{ offset: number; value: T }> = [];
  const n = Math.max(1, Math.floor(count));
  for (let index = 0; index <= n; index += 1) {
    const offset = index / n;
    frames.push({ offset, value: at(smoothProgress(offset, accel, decel)) });
  }
  return frames;
}

/**
 * Keyframes laid at progress `offsets` (ascending, 0..1), moved to the times
 * E reaches them, each with the easing that runs E from it to the next one:
 * played so (and the effect itself linearly), they follow Smooth start / end
 * over the whole effect whatever lies between them. An even pace leaves them
 * where they are, linear.
 */
export function smoothFrameTimes(
  offsets: readonly number[],
  accel: number,
  decel: number,
): Array<{ time: number; easing: string }> {
  const smooth = isSmooth(accel, decel);
  const times = offsets.map((offset) => (smooth ? smoothTimeAt(offset, accel, decel) : Math.max(0, Math.min(1, offset))));
  return offsets.map((offset, index) => {
    const time = times[index]!;
    const nextTime = times[index + 1];
    const nextOffset = offsets[index + 1];
    if (!smooth || nextTime === undefined || nextOffset === undefined || nextTime - time < 1e-6 || nextOffset - offset < 1e-9) {
      return { time, easing: "linear" };
    }
    if (!supportsLinearEasing()) {
      // No `linear()`: the whole effect between its two ends takes the
      // nearest cubic Bézier; a piece of it runs evenly between its stops,
      // which E has already put at their times.
      return { time, easing: offset === 0 && nextOffset === 1 ? smoothEasingCubic(accel, decel) : "linear" };
    }
    const span = nextTime - time;
    const rise = nextOffset - offset;
    return {
      time,
      easing: linearEasingOf((s) => (smoothProgress(time + s * span, accel, decel) - offset) / rise, 8),
    };
  });
}

/**
 * A CSS `@keyframes` body (`0%{...}50%{...}100%{...}`, `from` / `to`, a stop
 * list sharing one block) with Smooth start / end baked in: each stop moved
 * to the time E reaches it and given the piece of E up to the next stop as its
 * `animation-timing-function`. Null when the body has no stop between its
 * start and end (two keyframes take the plain easing) or cannot be read.
 */
export function smoothKeyframeBody(body: string, accel: number, decel: number): string | null {
  if (!isSmooth(accel, decel)) return null;
  const blocks: Array<{ stops: number[]; declarations: string }> = [];
  let at = 0;
  while (at < body.length) {
    const open = body.indexOf("{", at);
    if (open < 0) {
      if (body.slice(at).trim()) return null;
      break;
    }
    const close = body.indexOf("}", open);
    if (close < 0 || body.slice(open + 1, close).includes("{")) return null;
    const selector = body.slice(at, open).trim();
    const stops: number[] = [];
    for (const raw of selector.split(",")) {
      const stop = raw.trim().toLowerCase();
      if (stop === "from") stops.push(0);
      else if (stop === "to") stops.push(100);
      else if (/^\d+(?:\.\d+)?%$/.test(stop)) stops.push(Number(stop.slice(0, -1)));
      else return null;
    }
    blocks.push({ stops, declarations: body.slice(open + 1, close).trim() });
    at = close + 1;
  }
  const frames = blocks
    .flatMap((block, order) => block.stops.map((stop) => ({ stop, order, declarations: block.declarations })))
    .sort((left, right) => left.stop - right.stop || left.order - right.order);
  if (!frames.some((frame) => frame.stop > 0 && frame.stop < 100)) return null;
  const timed = smoothFrameTimes(
    frames.map((frame) => frame.stop / 100),
    accel,
    decel,
  );
  return frames
    .map((frame, index) => {
      const { time, easing } = timed[index]!;
      const declarations = frame.declarations.replace(/;?\s*$/, "");
      const separator = declarations ? ";" : "";
      return `${round4(time * 100)}%{${declarations}${separator}animation-timing-function:${easing}}`;
    })
    .join("");
}
