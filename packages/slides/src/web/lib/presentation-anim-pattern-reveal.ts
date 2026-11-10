/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Blinds, Checkerboard and Random Bars reveal an object in many pieces at once
 * (PowerPoint's `blinds(horizontal|vertical)`, `checkerboard(across|down)` and
 * `randombar(horizontal|vertical)` filters). A single `inset()` wipe drew all
 * three as one rectangle (PPT-B138).
 *
 * Each is drawn as ONE `clip-path` polygon: its pieces are rectangles, joined
 * along the object's left edge (the top edge when vertical) by zero-area
 * back-and-forth segments, so the pieces fill and the joins do not. The
 * polygon keeps the same vertices at every keyframe, so the browser
 * interpolates it vertex by vertex, and every keyframe sits where a piece
 * starts or stops moving, so the interpolation is exact.
 */

export type PatternRevealDirection = "horizontal" | "vertical";

/** `[x0, y0, x1, y1]` in percent of the object, laid out along rows. */
type Piece = readonly [number, number, number, number];

/** A pattern's pieces at progress `p` (0 hidden, 1 revealed). */
type Pattern = (p: number) => readonly Piece[];

const BLINDS_SLATS = 6;
const CHECKERBOARD_ROWS = 6;
const CHECKERBOARD_COLUMNS = 8;
const RANDOM_BARS = 24;
const RANDOM_BARS_PER_STEP = 2;

/** Horizontal slats, each opening from its top edge down. */
function blinds(p: number): readonly Piece[] {
  const slat = 100 / BLINDS_SLATS;
  return Array.from({ length: BLINDS_SLATS }, (_, index) => {
    const top = index * slat;
    return [0, top, 100, top + p * slat] as const;
  });
}

/**
 * Squares opening across, every other row half a square behind, so halfway
 * through the object is a checkerboard. The first and last squares of an
 * offset row are clipped at the object's edges; both clips turn on at
 * `p = 0.5`, which is a keyframe.
 */
function checkerboard(p: number): readonly Piece[] {
  const width = 100 / CHECKERBOARD_COLUMNS;
  const height = 100 / CHECKERBOARD_ROWS;
  const clamp = (value: number) => Math.min(100, Math.max(0, value));
  const pieces: Piece[] = [];
  for (let row = 0; row < CHECKERBOARD_ROWS; row += 1) {
    const offset = row % 2 === 0 ? 0 : width / 2;
    const count = offset ? CHECKERBOARD_COLUMNS + 1 : CHECKERBOARD_COLUMNS;
    for (let column = 0; column < count; column += 1) {
      const start = column * width - offset;
      pieces.push([
        clamp(start),
        row * height,
        clamp(start + p * width),
        (row + 1) * height,
      ]);
    }
  }
  return pieces;
}

/**
 * The order the bars appear in: one fixed shuffle, so every play looks the
 * same, and no bar follows its neighbour.
 */
const RANDOM_BAR_ORDER: readonly number[] = [
  13, 4, 19, 9, 22, 1, 16, 7, 11, 20, 2, 15, 5, 23, 10, 0, 18, 8, 14, 3, 21, 12, 6, 17,
];

const RANDOM_BAR_STEPS = Math.ceil(RANDOM_BARS / RANDOM_BARS_PER_STEP);

/** Thin full-width bars, a few at a time in a scattered order. */
function randomBars(p: number): readonly Piece[] {
  const bar = 100 / RANDOM_BARS;
  const reached = p * RANDOM_BAR_STEPS;
  return Array.from({ length: RANDOM_BARS }, (_, index) => {
    const step = Math.floor(RANDOM_BAR_ORDER.indexOf(index) / RANDOM_BARS_PER_STEP);
    const grown = Math.min(1, Math.max(0, reached - step));
    return [0, index * bar, 100, index * bar + grown * bar] as const;
  });
}

function coordinate(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  return rounded === 0 ? "0" : `${rounded}%`;
}

/**
 * The pieces as one polygon. A piece that starts at the left edge is entered
 * and left along that edge; one that starts further in is reached along its
 * row's top line and left the same way back. Which joins are written depends
 * only on where each piece starts, never on `p`, so every keyframe of a
 * pattern has the same vertices.
 */
function piecesPolygon(
  pieces: readonly Piece[],
  direction: PatternRevealDirection,
): string {
  const points: Array<readonly [number, number]> = [[0, 0]];
  let previous: Piece | null = null;
  for (const piece of pieces) {
    const [x0, y0, x1, y1] = piece;
    const onEdgeAlready = previous !== null && previous[1] === y0 && previous[0] !== 0;
    if (!onEdgeAlready) points.push([0, y0]);
    if (x0 !== 0) points.push([x0, y0]);
    points.push([x1, y0], [x1, y1], [x0, y1]);
    if (x0 !== 0) points.push([x0, y0], [0, y0]);
    previous = piece;
  }
  return `polygon(${points
    .map(([x, y]) =>
      direction === "horizontal"
        ? `${coordinate(x)} ${coordinate(y)}`
        : `${coordinate(y)} ${coordinate(x)}`,
    )
    .join(",")})`;
}

const PATTERNS: Readonly<
  Record<"blinds" | "checkerboard" | "random-bars", {
    pattern: Pattern;
    /** Progress values that need a keyframe. */
    stops: readonly number[];
    /**
     * Each segment's own timing, where the element's easing would otherwise
     * repeat once per segment.
     */
    segmentTiming?: (entering: boolean) => string;
  }>
> = {
  blinds: { pattern: blinds, stops: [0, 1] },
  checkerboard: { pattern: checkerboard, stops: [0, 0.5, 1], segmentTiming: () => "linear" },
  "random-bars": {
    pattern: randomBars,
    stops: Array.from({ length: RANDOM_BAR_STEPS + 1 }, (_, index) => index / RANDOM_BAR_STEPS),
    // A bar appears (or leaves) whole, as PowerPoint draws it, rather than
    // growing: an entrance shows each step's bars as the step starts, an exit
    // removes them as it ends, the one journey run backwards.
    segmentTiming: (entering) => (entering ? "step-start" : "step-end"),
  },
};

export type PatternRevealKind = keyof typeof PATTERNS;

export function patternRevealKind(animation: string): PatternRevealKind | null {
  if (animation.startsWith("blinds-")) return "blinds";
  if (animation.startsWith("checkerboard-")) return "checkerboard";
  if (animation.startsWith("random-bars-")) return "random-bars";
  return null;
}

/** The pieces at progress `p`, for tests and anything that samples them. */
export function patternRevealPieces(kind: PatternRevealKind, p: number): readonly Piece[] {
  return PATTERNS[kind].pattern(p);
}

/**
 * The `@keyframes` body: an entrance runs from hidden to revealed, an exit the
 * same journey backwards.
 */
export function patternRevealKeyframeBody(
  kind: PatternRevealKind,
  direction: PatternRevealDirection,
  entering: boolean,
): string {
  const { pattern, stops, segmentTiming } = PATTERNS[kind];
  return stops
    .map((stop, index) => {
      const offset = Math.round(stop * 100000) / 1000;
      const progress = entering ? stop : 1 - stop;
      const timing =
        segmentTiming && index < stops.length - 1
          ? `;animation-timing-function:${segmentTiming(entering)}`
          : "";
      return `${offset}%{clip-path:${piecesPolygon(pattern(progress), direction)}${timing}}`;
    })
    .join("");
}

/**
 * The same journey as Web Animations keyframes (the timing-tree player's
 * `blinds(...)`, `checkerboard(...)` and `randombar(...)` filters).
 */
export function patternRevealKeyframes(
  kind: PatternRevealKind,
  direction: PatternRevealDirection,
  entering: boolean,
): Keyframe[] {
  const { pattern, stops, segmentTiming } = PATTERNS[kind];
  return stops.map((stop, index) => ({
    offset: stop,
    clipPath: piecesPolygon(pattern(entering ? stop : 1 - stop), direction),
    ...(segmentTiming && index < stops.length - 1 ? { easing: segmentTiming(entering) } : {}),
  }));
}
