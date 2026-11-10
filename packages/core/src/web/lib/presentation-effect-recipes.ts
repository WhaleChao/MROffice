/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Playback recipes for the PowerPoint animation effects the editor does not
 * draw with a hand-written keyframe.
 *
 * A recipe is a neutral description of what an effect does to an object over
 * its duration: tracks that move, size, turn, fade or reveal it. The recipes in
 * `presentation-effect-recipes.json` are generated (never edited by hand) from
 * effects measured in PowerPoint itself, one track per behaviour PowerPoint
 * plays. This module turns a recipe into CSS keyframes for one object: the
 * values depend on where the object sits on the slide (an effect that enters
 * from the slide's edge travels further for an object near the other edge), so
 * keyframes are built per object box.
 *
 * Units. `X`/`Y` are the object's centre and `W`/`H` its size, as fractions of
 * the slide; an expression may use them, the live values `x`/`y`/`w`/`h`, and
 * `$`, the interpolated key value of a formula key.
 */
import recipesJson from "./presentation-effect-recipes.json";
import { RUN_FONT_FAMILIES, runFontFamilyName } from "./presentation-runs";
import { cssFontFamilyList } from "./css-font-family";

export interface EffectRecipeBox {
  /** Centre and size as fractions of the slide. */
  x: number;
  y: number;
  w: number;
  h: number;
}

interface TrackTiming {
  /** Begin and duration as fractions of the effect's duration. */
  b: number;
  d: number;
  ac?: number;
  dc?: number;
  ar?: 1;
  tf?: number[][];
  hold?: 1;
}

type RecipeKey = [number, string] | [number, string, string];

type RecipeTrack =
  | { k: "vis"; at: number; v: string }
  | ({
      k: "num";
      p: "x" | "y" | "w" | "h" | "rot" | "shear" | "opacity";
      keys?: RecipeKey[];
      from?: string;
      to?: string;
      by?: string;
      add?: string;
    } & TrackTiming)
  | ({ k: "set"; p: "x" | "y" | "w" | "h" | "rot" | "shear" | "opacity"; v: string } & TrackTiming)
  | ({ k: "scale"; from?: number[]; to?: number[]; by?: number[] } & TrackTiming)
  | ({ k: "rot"; by: number } & TrackTiming)
  | ({ k: "motion"; path: string } & TrackTiming)
  | ({ k: "fx"; f: string; t: "in" | "out"; arg?: string; pr?: string } & TrackTiming)
  | ({ k: "clr"; p: string; by?: Record<string, number | string>; to?: Record<string, string> } & TrackTiming)
  | ({ k: "prop"; p: string; v?: string; to?: string; keys?: RecipeKey[] } & TrackTiming);

export interface EffectRecipeVariant {
  /** PowerPoint's own duration for the effect, in milliseconds. */
  d: number;
  ac?: number;
  dc?: number;
  ar?: 1;
  /**
   * Text iteration: the effect plays on the text letter by letter, each letter
   * starting a share of the duration (`pct`) or a time (`abs`, ms) after the
   * one before.
   */
  it?: { type: string; pct?: number; abs?: number };
  tracks: RecipeTrack[];
}

interface EffectRecipe {
  cls: "entr" | "exit" | "emph";
  group: string;
  /**
   * Options of an effect the editor draws by hand: only these options play
   * from a recipe, the effect's own default stays hand-drawn.
   */
  ext?: 1;
  /** The option the effect plays when none is named. */
  def?: string;
  variants: Record<string, EffectRecipeVariant>;
}

const RECIPES = recipesJson as unknown as Record<string, EffectRecipe>;

/** The recipe of an effect and option, or `null` for an effect drawn another way. */
export function effectRecipe(
  effect: string | undefined,
  option?: string | null,
): EffectRecipeVariant | null {
  if (!effect) return null;
  const recipe = Object.prototype.hasOwnProperty.call(RECIPES, effect)
    ? RECIPES[effect]
    : undefined;
  if (!recipe) return null;
  const variants = recipe.variants;
  if (effect === "font") {
    // The one value the recipe leaves open is the family the option names.
    const named = option?.startsWith("font:") ? option.slice(5) : "";
    const base = variants[""];
    return {
      ...base,
      tracks: base.tracks.map((track) =>
        track.k === "prop" && track.p === "style.fontFamily"
          ? { ...track, v: named || "Arial" }
          : track,
      ),
    };
  }
  if (option && variants[option]) return variants[option];
  if (recipe.ext) return null;
  return variants[recipe.def ?? ""] ?? Object.values(variants)[0] ?? null;
}

/** PowerPoint's own duration of an effect (ms), or `null` for a non-recipe effect. */
export function effectRecipeDefaultDurationMs(
  effect: string | undefined,
  option?: string | null,
): number | null {
  const recipe = effectRecipe(effect, option);
  return recipe ? recipe.d : null;
}

// ── expressions ────────────────────────────────────────────────────────────

type Scope = Record<string, number>;

const FUNCTIONS: Record<string, (...args: number[]) => number> = {
  sin: Math.sin,
  cos: Math.cos,
  tan: Math.tan,
  abs: Math.abs,
  sqrt: Math.sqrt,
  min: Math.min,
  max: Math.max,
  exp: Math.exp,
  ln: Math.log,
};

/**
 * Evaluate a recipe expression: numbers, the scope's names, `pi`, `e`,
 * `+ - * / ^`, parentheses and a few functions. Anything else is `NaN`, which
 * the caller treats as "no change".
 */
export function evaluateRecipeExpression(source: string, scope: Scope): number {
  let index = 0;
  const text = source;
  const peek = () => text[index];
  const skip = () => {
    while (index < text.length && /\s/.test(text[index])) index += 1;
  };
  const primary = (): number => {
    skip();
    const char = peek();
    if (char === undefined) return Number.NaN;
    if (char === "(") {
      index += 1;
      const value = sum();
      skip();
      if (peek() === ")") index += 1;
      return value;
    }
    if (char === "-") {
      index += 1;
      return -power();
    }
    if (char === "+") {
      index += 1;
      return power();
    }
    const number = /^(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/.exec(text.slice(index));
    if (number) {
      index += number[0].length;
      return Number.parseFloat(number[0]);
    }
    if (char === "$") {
      index += 1;
      return scope.$ ?? Number.NaN;
    }
    const name = /^[A-Za-z_][A-Za-z_0-9]*/.exec(text.slice(index));
    if (!name) {
      index = text.length;
      return Number.NaN;
    }
    index += name[0].length;
    skip();
    if (peek() === "(") {
      index += 1;
      const args: number[] = [];
      skip();
      if (peek() !== ")") {
        args.push(sum());
        skip();
        while (peek() === ",") {
          index += 1;
          args.push(sum());
          skip();
        }
      }
      if (peek() === ")") index += 1;
      const fn = FUNCTIONS[name[0].toLowerCase()];
      return fn ? fn(...args) : Number.NaN;
    }
    if (name[0] === "pi") return Math.PI;
    if (name[0] === "e") return Math.E;
    return Object.prototype.hasOwnProperty.call(scope, name[0])
      ? scope[name[0]]
      : Number.NaN;
  };
  const power = (): number => {
    const base = primary();
    skip();
    if (peek() === "^") {
      index += 1;
      return Math.pow(base, power());
    }
    return base;
  };
  const product = (): number => {
    let value = power();
    for (;;) {
      skip();
      const char = peek();
      if (char === "*") {
        index += 1;
        value *= power();
      } else if (char === "/") {
        index += 1;
        value /= power();
      } else return value;
    }
  };
  const sum = (): number => {
    let value = product();
    for (;;) {
      skip();
      const char = peek();
      if (char === "+") {
        index += 1;
        value += product();
      } else if (char === "-") {
        index += 1;
        value -= product();
      } else return value;
    }
  };
  const value = sum();
  skip();
  return index < text.length ? Number.NaN : value;
}

// ── timing ─────────────────────────────────────────────────────────────────

/** PowerPoint's accelerate / decelerate: a share of the time speeding up, a share slowing down. */
function ease(t: number, accel = 0, decel = 0): number {
  const a = Math.max(0, Math.min(1, accel));
  const d = Math.max(0, Math.min(1 - a, decel));
  if (a === 0 && d === 0) return t;
  const speed = 1 / (1 - a / 2 - d / 2);
  if (t < a) return (speed * t * t) / (2 * a);
  if (t > 1 - d) {
    const rest = 1 - t;
    return 1 - (speed * rest * rest) / (2 * d);
  }
  return speed * (t - a / 2);
}

function timeFilter(t: number, points: number[][] | undefined): number {
  if (!points || points.length < 2) return t;
  for (let i = 1; i < points.length; i += 1) {
    const [t0, v0] = points[i - 1];
    const [t1, v1] = points[i];
    if (t <= t1) {
      return t1 === t0 ? v1 : v0 + ((v1 - v0) * (t - t0)) / (t1 - t0);
    }
  }
  return points[points.length - 1][1];
}

/**
 * A track's local progress at effect time `u` (both 0..1): `null` while the
 * track is not in force (before it begins, or after it ends without holding).
 */
function trackProgress(track: TrackTiming, u: number): number | null {
  const begin = track.b ?? 0;
  if (u < begin - 1e-9) return null;
  if (!track.d) return 1;
  let local = (u - begin) / track.d;
  if (track.ar) {
    if (local > 2) {
      if (!track.hold) return null;
      local = 0;
    } else if (local > 1) local = 2 - local;
  } else if (local > 1) {
    if (!track.hold) return null;
    local = 1;
  }
  return timeFilter(ease(local, track.ac, track.dc), track.tf);
}

// ── reveal shapes ──────────────────────────────────────────────────────────

const pct = (value: number) => `${Math.round(value * 100) / 100}%`;
const polygon = (points: Array<[number, number]>, rule?: "evenodd") =>
  `polygon(${rule ? "evenodd," : ""}${points.map(([x, y]) => `${pct(x)} ${pct(y)}`).join(",")})`;

const FRAME: Array<[number, number]> = [
  [0, 0],
  [100, 0],
  [100, 100],
  [0, 100],
  [0, 0],
];

function ring(shape: Array<[number, number]>, hole: boolean): string {
  // A shape alone shows its inside; with the frame around it (even-odd) it
  // shows everything but its inside.
  return hole
    ? polygon([...FRAME, ...shape, shape[0], [0, 0]], "evenodd")
    : polygon(shape);
}

function circlePoints(radius: number): Array<[number, number]> {
  const points: Array<[number, number]> = [];
  for (let i = 0; i < 32; i += 1) {
    const angle = (i / 32) * 2 * Math.PI;
    points.push([50 + radius * Math.cos(angle), 50 + radius * Math.sin(angle)]);
  }
  return points;
}

function sector(from: number, to: number): Array<[number, number]> {
  const points: Array<[number, number]> = [[50, 50]];
  for (let i = 0; i <= 12; i += 1) {
    const angle = from + ((to - from) * i) / 12;
    points.push([50 + 150 * Math.sin(angle), 50 - 150 * Math.cos(angle)]);
  }
  points.push([50, 50]);
  return points;
}

const DISSOLVE_COLUMNS = 24;
const DISSOLVE_ROWS = 14;
/** One fixed order in which the dissolve's cells appear (a small LCG, so every run draws the same). */
const DISSOLVE_ORDER: number[] = (() => {
  const count = DISSOLVE_COLUMNS * DISSOLVE_ROWS;
  const order = Array.from({ length: count }, (_, i) => i);
  let seed = 20261004;
  for (let i = count - 1; i > 0; i -= 1) {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    const j = seed % (i + 1);
    [order[i], order[j]] = [order[j], order[i]];
  }
  return order;
})();

function dissolveMask(shown: number): string {
  const count = Math.round(shown * DISSOLVE_ORDER.length);
  if (count <= 0) return "linear-gradient(transparent,transparent)";
  if (count >= DISSOLVE_ORDER.length) return "none";
  let path = "";
  for (let i = 0; i < count; i += 1) {
    const cell = DISSOLVE_ORDER[i];
    path += `M${cell % DISSOLVE_COLUMNS} ${Math.floor(cell / DISSOLVE_COLUMNS)}h1v1h-1z`;
  }
  const svg =
    `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 ${DISSOLVE_COLUMNS} ${DISSOLVE_ROWS}' ` +
    `preserveAspectRatio='none' shape-rendering='crispEdges'><path d='${path}'/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

/** The part of the box a diagonal Strips reveal shows at `shown` (always six
 * points, so successive frames interpolate). */
function stripsPoints(arg: string | undefined, shown: number): Array<[number, number]> {
  const direction = (arg ?? "downright").toLowerCase();
  const fromRight = direction.endsWith("left");
  const fromBottom = direction.startsWith("up");
  const reach = shown * 200;
  // In the reveal's own corner frame: x + y <= reach.
  let points: Array<[number, number]> = [[0, 0], [100, 0], [100, 100], [0, 100]];
  const inside = ([x, y]: [number, number]) => x + y <= reach;
  const clipped: Array<[number, number]> = [];
  for (let i = 0; i < points.length; i += 1) {
    const a = points[i]!;
    const b = points[(i + 1) % points.length]!;
    if (inside(a)) clipped.push(a);
    if (inside(a) !== inside(b)) {
      const t = (reach - a[0] - a[1]) / (b[0] + b[1] - a[0] - a[1]);
      clipped.push([a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1])]);
    }
  }
  points = clipped.length > 0 ? clipped : [[0, 0]];
  while (points.length < 6) points.push(points[points.length - 1]!);
  return points.map(([x, y]) => [fromRight ? 100 - x : x, fromBottom ? 100 - y : y]);
}

/**
 * The `clip-path` a shape reveal (`p:animEffect@filter`: circle, box,
 * diamond, plus, wedge, wheel, strips, wipe) shows `shown` of the object
 * with (0 nothing, 1 all), as the editor's effect recipes draw it; `arg` is
 * the filter's option (`in` / `out`, the spoke count, the direction).
 */
export function effectRevealClip(filter: string, arg: string | undefined, shown: number): string | undefined {
  return reveal(filter, arg, shown).clip;
}

/** How much of the object a reveal shows: `clip-path`, or a mask for the dissolve. */
function reveal(
  filter: string,
  arg: string | undefined,
  shown: number,
): { clip?: string; mask?: string } {
  const s = Math.max(0, Math.min(1, shown));
  switch (filter) {
    case "circle": {
      // `out`: grows from the centre. `in`: closes in from the edges.
      const radius = arg === "out" ? s * 75 : (1 - s) * 75;
      return { clip: ring(circlePoints(radius), arg !== "out") };
    }
    case "diamond": {
      const reach = arg === "out" ? s * 100 : (1 - s) * 100;
      return {
        clip: ring(
          [
            [50, 50 - reach],
            [50 + reach, 50],
            [50, 50 + reach],
            [50 - reach, 50],
          ],
          arg !== "out",
        ),
      };
    }
    case "plus": {
      const half = arg === "out" ? s * 50 : (1 - s) * 50;
      const a = 50 - half;
      const b = 50 + half;
      return {
        clip: ring(
          [
            [a, 0],
            [b, 0],
            [b, a],
            [100, a],
            [100, b],
            [b, b],
            [b, 100],
            [a, 100],
            [a, b],
            [0, b],
            [0, a],
            [a, a],
          ],
          arg !== "out",
        ),
      };
    }
    case "box": {
      // `out`: grows from the centre. `in`: closes in from the edges.
      const reach = arg === "out" ? s * 50 : (1 - s) * 50;
      return {
        clip: ring(
          [
            [50 - reach, 50 - reach],
            [50 + reach, 50 - reach],
            [50 + reach, 50 + reach],
            [50 - reach, 50 + reach],
          ],
          arg !== "out",
        ),
      };
    }
    case "strips":
      // The edge travels diagonally in the named direction (downRight: from
      // the top-left corner to the bottom-right one).
      return { clip: polygon(stripsPoints(arg, s)) };
    case "wedge":
      return { clip: polygon(sector(-s * Math.PI, s * Math.PI)) };
    case "wheel": {
      const spokes = Math.max(1, Number.parseInt(arg ?? "1", 10) || 1);
      const points: Array<[number, number]> = [];
      for (let i = 0; i < spokes; i += 1) {
        const start = (i / spokes) * 2 * Math.PI;
        points.push(...sector(start, start + (s * 2 * Math.PI) / spokes));
      }
      return { clip: polygon(points) };
    }
    case "wipe": {
      const rest = pct((1 - s) * 100);
      // The edge travels in the named direction.
      const inset =
        arg === "up"
          ? `inset(${rest} 0 0 0)`
          : arg === "down"
            ? `inset(0 0 ${rest} 0)`
            : arg === "left"
              ? `inset(0 0 0 ${rest})`
              : `inset(0 ${rest} 0 0)`;
      return { clip: inset };
    }
    case "dissolve":
      return { mask: dissolveMask(s) };
    default:
      return {};
  }
}

// ── motion path ────────────────────────────────────────────────────────────

/** The points of a motion path (slide fractions from the object's place), curves flattened. */
function pathPoints(path: string): Array<[number, number]> {
  const tokens = path.match(/[A-Za-z]|-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  const points: Array<[number, number]> = [];
  let cursor: [number, number] = [0, 0];
  let start: [number, number] = [0, 0];
  let i = 0;
  const number = () => Number.parseFloat(tokens[i++] ?? "0");
  // PowerPoint writes absolute commands in capitals and relative ones in
  // lower case, as SVG does; a command's letter may be left off its repeats.
  let command = "";
  while (i < tokens.length) {
    if (/^[A-Za-z]$/.test(tokens[i])) command = tokens[i++];
    else if (!command) {
      i += 1;
      continue;
    }
    const relative = command === command.toLowerCase();
    const at = (x: number, y: number): [number, number] =>
      relative ? [cursor[0] + x, cursor[1] + y] : [x, y];
    const kind = command.toUpperCase();
    if (kind === "M" || kind === "L") {
      cursor = at(number(), number());
      if (kind === "M") start = cursor;
      points.push(cursor);
      if (kind === "M") command = relative ? "l" : "L";
    } else if (kind === "C") {
      const p0 = cursor;
      const p1 = at(number(), number());
      const p2 = at(number(), number());
      const p3 = at(number(), number());
      for (let step = 1; step <= 16; step += 1) {
        const t = step / 16;
        const m = 1 - t;
        points.push([
          m * m * m * p0[0] + 3 * m * m * t * p1[0] + 3 * m * t * t * p2[0] + t * t * t * p3[0],
          m * m * m * p0[1] + 3 * m * m * t * p1[1] + 3 * m * t * t * p2[1] + t * t * t * p3[1],
        ]);
      }
      cursor = p3;
    } else if (kind === "Z") {
      if (cursor[0] !== start[0] || cursor[1] !== start[1]) points.push(start);
      cursor = start;
      command = "";
    } else {
      // `E` ends the path; anything unknown ends the reading.
      break;
    }
  }
  return points;
}

function pathPoint(path: string, progress: number): [number, number] {
  const points = pathPoints(path);
  if (points.length === 0) return [0, 0];
  const lengths = [0];
  for (let k = 1; k < points.length; k += 1) {
    lengths.push(
      lengths[k - 1] +
        Math.hypot(points[k][0] - points[k - 1][0], points[k][1] - points[k - 1][1]),
    );
  }
  const target = Math.max(0, Math.min(1, progress)) * lengths[lengths.length - 1];
  for (let k = 1; k < points.length; k += 1) {
    if (target <= lengths[k]) {
      const span = lengths[k] - lengths[k - 1];
      const f = span > 0 ? (target - lengths[k - 1]) / span : 0;
      return [
        points[k - 1][0] + (points[k][0] - points[k - 1][0]) * f,
        points[k - 1][1] + (points[k][1] - points[k - 1][1]) * f,
      ];
    }
  }
  return points[points.length - 1];
}

/**
 * The line a motion-path effect travels, as points in slide fractions from
 * the object's place; `null` for an effect that follows no path. For drawing
 * the path (the gallery's glyph), not for playing it.
 */
export function effectRecipeMotionPoints(
  effect: string | undefined,
  option?: string | null,
): Array<[number, number]> | null {
  const recipe = effectRecipe(effect, option);
  const track = recipe?.tracks.find((candidate) => candidate.k === "motion");
  return track && track.k === "motion" ? pathPoints(track.path) : null;
}

// ── sampling ───────────────────────────────────────────────────────────────

/**
 * The CSS `font-family` of a family an effect names. One the editor offers
 * uses its own fallback stack; any other is quoted, and a name that could not
 * be written into a style sheet as a plain string is not drawn at all.
 */
function cssFontFamily(name: string): string {
  const known = RUN_FONT_FAMILIES.find(
    (family) => family.id !== "system" && runFontFamilyName(family) === name,
  );
  if (known) return known.css;
  return !name || /["\\<>{};\n\r]/.test(name) ? "" : `"${name}", sans-serif`;
}

export interface EffectRecipeFrame {
  /** Position in the effect, 0..1. */
  at: number;
  visible: boolean;
  /** Offset of the centre, in slide fractions. */
  dx: number;
  dy: number;
  scaleX: number;
  scaleY: number;
  rotate: number;
  shear: number;
  opacity: number;
  clip?: string;
  mask?: string;
  /** A colour shift of the whole object (hue turn in degrees, saturation and lightness deltas, -1..1). */
  hsl?: { h: number; s: number; l: number };
  /** Text style the effect switches on. */
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** A CSS `font-family` the text changes to; "" where it does not apply. */
  fontFamily?: string;
  /** The theme colour the text takes while the effect states one. */
  color?: string;
}

function keyValue(
  keys: RecipeKey[],
  progress: number,
  scope: Scope,
): number {
  if (keys.length === 0) return Number.NaN;
  let i = 0;
  while (i < keys.length - 2 && progress > keys[i + 1][0]) i += 1;
  const from = keys[i];
  const to = keys[Math.min(i + 1, keys.length - 1)];
  const span = to[0] - from[0];
  const f = span > 0 ? Math.max(0, Math.min(1, (progress - from[0]) / span)) : 1;
  if (from.length === 3) {
    const a = evaluateRecipeExpression(from[1], scope);
    const b = evaluateRecipeExpression(to[1], scope);
    return evaluateRecipeExpression(from[2] as string, { ...scope, $: a + (b - a) * f });
  }
  const a = evaluateRecipeExpression(from[1], scope);
  const b = evaluateRecipeExpression(to[1], scope);
  return a + (b - a) * f;
}

/** The object's state at effect time `u` (0..1). */
export function sampleEffectRecipe(
  recipe: EffectRecipeVariant,
  cls: "entr" | "exit" | "emph",
  box: EffectRecipeBox,
  at: number,
): EffectRecipeFrame {
  const u = ease(Math.max(0, Math.min(1, at)), recipe.ac, recipe.dc);
  const state = { x: box.x, y: box.y, w: box.w, h: box.h, rot: 0, shear: 0, opacity: 1 };
  let scaleX = 1;
  let scaleY = 1;
  let visible = cls !== "entr";
  let fade = 1;
  let clip: string | undefined;
  let mask: string | undefined;
  let hsl: { h: number; s: number; l: number } | undefined;
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underline: boolean | undefined;
  let fontFamily: string | undefined;
  let color: string | undefined;
  let hasColor = false;
  const scope = (): Scope => ({
    X: box.x,
    Y: box.y,
    W: box.w,
    H: box.h,
    x: state.x,
    y: state.y,
    w: state.w,
    h: state.h,
  });
  for (const track of recipe.tracks) {
    if (track.k === "vis") {
      if (u >= track.at - 1e-9) visible = track.v !== "hidden";
      continue;
    }
    const progress = trackProgress(track, u);
    if (track.k === "fx") {
      // Before a reveal begins nothing of an entrance shows yet; an exit is whole.
      const p = progress ?? (u < track.b ? 0 : 1);
      const shown = track.t === "out" ? 1 - p : p;
      if (track.f === "fade") fade *= shown;
      else {
        const shape = reveal(track.f, track.arg, shown);
        clip = shape.clip ?? clip;
        mask = shape.mask ?? mask;
      }
      continue;
    }
    if (progress === null) continue;
    // A value that is set and then auto-reversed goes back when the reverse
    // ends, whatever the track holds afterwards.
    const reverted =
      Boolean(track.ar) && track.d > 0 && u >= track.b + 2 * track.d - 1e-9;
    if (track.k === "num") {
      let value: number;
      if (track.keys) value = keyValue(track.keys, progress, scope());
      else {
        const base = state[track.p];
        const from = track.from !== undefined ? evaluateRecipeExpression(track.from, scope()) : base;
        const to =
          track.to !== undefined
            ? evaluateRecipeExpression(track.to, scope())
            : track.by !== undefined
              ? from + evaluateRecipeExpression(track.by, scope())
              : base;
        value = from + (to - from) * progress;
      }
      if (Number.isFinite(value)) state[track.p] = value;
    } else if (track.k === "set") {
      if (reverted) continue;
      const value = evaluateRecipeExpression(track.v, scope());
      if (Number.isFinite(value)) state[track.p] = value;
    } else if (track.k === "scale") {
      const from = track.from ?? [scaleX, scaleY];
      const to = track.to ?? (track.by ? [from[0] * track.by[0], from[1] * track.by[1]] : from);
      scaleX = from[0] + (to[0] - from[0]) * progress;
      scaleY = from[1] + (to[1] - from[1]) * progress;
    } else if (track.k === "rot") {
      state.rot += track.by * progress;
    } else if (track.k === "motion") {
      const [px, py] = pathPoint(track.path, progress);
      state.x = box.x + px;
      state.y = box.y + py;
    } else if (track.k === "clr") {
      // A shift in hue / saturation / lightness moves fill, line and text
      // together; a change TO a colour is drawn by the caller's colour layers
      // (see `effectRecipeColorLayer`).
      if (track.by && track.by.sp === "hsl") {
        hsl = {
          h: (Number(track.by.h) / 60000) * progress,
          s: (Number(track.by.s) / 100000) * progress,
          l: (Number(track.by.l) / 100000) * progress,
        };
      }
    } else if (track.k === "prop") {
      let value = track.v;
      if (track.keys && track.keys.length > 0) {
        // Discrete keys: the value of the last key reached.
        let index = 0;
        while (index < track.keys.length - 1 && progress >= track.keys[index + 1][0]) index += 1;
        value = track.keys[index][1];
      }
      if (track.p === "style.fontWeight") bold = !reverted && value === "bold";
      else if (track.p === "style.fontStyle") italic = !reverted && value === "italic";
      else if (track.p === "style.textDecorationUnderline") underline = !reverted && value === "true";
      else if (track.p === "style.fontFamily") fontFamily = reverted ? "" : cssFontFamily(value ?? "");
      else if (track.p === "style.color") {
        hasColor = true;
        color = reverted ? undefined : value;
      } else if (track.p === "style.fontSize" && track.to !== undefined) {
        // A font size the effect grows to: drawn as a size change of the text.
        const target = Number.parseFloat(track.to);
        if (Number.isFinite(target)) {
          scaleX *= 1 + (target - 1) * progress;
          scaleY *= 1 + (target - 1) * progress;
        }
      }
    }
  }
  for (const track of recipe.tracks) {
    if (track.k === "prop" && track.p === "style.color") hasColor = true;
  }
  return {
    at,
    visible,
    dx: state.x - box.x,
    dy: state.y - box.y,
    scaleX: (box.w ? state.w / box.w : 1) * scaleX,
    scaleY: (box.h ? state.h / box.h : 1) * scaleY,
    rotate: state.rot,
    shear: state.shear,
    opacity: Math.max(0, Math.min(1, state.opacity * fade)),
    clip,
    mask,
    hsl,
    bold,
    italic,
    underline,
    ...(fontFamily !== undefined ? { fontFamily } : {}),
    ...(hasColor ? { color: color ?? "" } : {}),
  };
}

/** The colours an effect may name, as custom properties the show sets from the theme. */
const SCHEME_FALLBACK: Readonly<Record<string, string>> = {
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
  hlink: "#0563c1",
  folHlink: "#954f72",
  bg1: "#ffffff",
  tx1: "#000000",
};

/** The CSS colour of a theme colour an effect names. */
export function effectRecipeSchemeColor(scheme: string): string {
  const fallback = SCHEME_FALLBACK[scheme] ?? "currentColor";
  return /^[A-Za-z0-9]+$/.test(scheme)
    ? `var(--mrai-ppt-fx-${scheme},${fallback})`
    : fallback;
}

/**
 * A hue / saturation / lightness shift as a CSS filter. The hue turn is the
 * filter's own; saturation is relative (-0.7 leaves 30%); lightness moves a
 * mid-tone by the same amount (a quarter lighter is half again as bright).
 */
function hslFilter(shift: { h: number; s: number; l: number }): string {
  const parts: string[] = [];
  if (Math.abs(shift.h) > 0.01) parts.push(`hue-rotate(${round(shift.h, 2)}deg)`);
  if (Math.abs(shift.s) > 0.0005) parts.push(`saturate(${round(Math.max(0, 1 + shift.s))})`);
  if (Math.abs(shift.l) > 0.0005) {
    parts.push(`brightness(${round(Math.max(0.1, Math.min(2.5, 1 + 2 * shift.l)))})`);
  }
  return parts.length > 0 ? parts.join(" ") : "none";
}

interface FrameChannels {
  clip: boolean;
  mask: boolean;
  filter: boolean;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  fontFamily: boolean;
  color: boolean;
}

function frameChannels(frames: readonly EffectRecipeFrame[]): FrameChannels {
  return {
    clip: frames.some((frame) => frame.clip !== undefined),
    mask: frames.some((frame) => frame.mask !== undefined),
    filter: frames.some((frame) => frame.hsl !== undefined),
    bold: frames.some((frame) => frame.bold !== undefined),
    italic: frames.some((frame) => frame.italic !== undefined),
    underline: frames.some((frame) => frame.underline !== undefined),
    fontFamily: frames.some((frame) => frame.fontFamily !== undefined),
    color: frames.some((frame) => frame.color !== undefined),
  };
}

const round = (value: number, digits = 3) => {
  const factor = 10 ** digits;
  const rounded = Math.round(value * factor) / factor;
  return Object.is(rounded, -0) ? 0 : rounded;
};

 
function turnCss(degrees: number): string {
  return degrees === 0 ? "rotate(0deg)" : `rotate(calc(${degrees}deg * var(--mrai-ppt-turn,1)))`;
}

function frameCss(frame: EffectRecipeFrame, channels: FrameChannels): string {
  const transform =
    `translate(${round(frame.dx * 100)}cqw,${round(frame.dy * 100)}cqh) ` +
    `var(--mrai-ppt-base-transform,) ${turnCss(round(frame.rotate))} ` +
    `skewX(${round((Math.atan(frame.shear) * 180) / Math.PI)}deg) ` +
    `scale(${round(frame.scaleX, 4)},${round(frame.scaleY, 4)})`;
  const parts = [
    `transform:${transform}`,
    `opacity:${round(frame.opacity)}`,
    `visibility:${frame.visible ? "visible" : "hidden"}`,
  ];
  if (channels.clip) parts.push(`clip-path:${frame.clip ?? "none"}`);
  if (channels.mask) {
    const mask = frame.mask ?? "none";
    parts.push(`-webkit-mask-image:${mask}`, `mask-image:${mask}`);
  }
  if (channels.filter) parts.push(`filter:${frame.hsl ? hslFilter(frame.hsl) : "none"}`);
  // Bold is drawn as a stroke in the text's own colour: it thickens every run,
  // whatever weight the run states for itself.
  if (channels.bold) {
    parts.push(`-webkit-text-stroke:${frame.bold ? "0.035em currentColor" : "0 currentColor"}`);
  }
  if (channels.italic) parts.push(`font-style:${frame.italic ? "italic" : "inherit"}`);
  if (channels.underline) {
    parts.push(`text-decoration-line:${frame.underline ? "underline" : "none"}`);
  }
  if (channels.fontFamily) parts.push(`font-family:${frame.fontFamily ? cssFontFamilyList(frame.fontFamily) : "inherit"}`);
  if (channels.color) {
    parts.push(`color:${frame.color ? effectRecipeSchemeColor(frame.color) : "inherit"}`);
  }
  return parts.join(";");
}

/** The sample times of a recipe: an even grid plus every moment a track starts, turns or ends. */
function sampleTimes(recipe: EffectRecipeVariant, steps = 40): number[] {
  const times = new Set<number>();
  for (let i = 0; i <= steps; i += 1) times.add(i / steps);
  const mark = (value: number) => {
    if (value > 0 && value < 1) {
      times.add(round(value, 5));
      times.add(round(Math.max(0, value - 0.0005), 5));
    }
  };
  for (const track of recipe.tracks) {
    if (track.k === "vis") {
      mark(track.at);
      continue;
    }
    mark(track.b);
    mark(track.b + track.d);
    if (track.ar) mark(track.b + 2 * track.d);
    if (track.k === "prop" && track.keys) {
      for (const key of track.keys) mark(track.b + key[0] * track.d);
    }
  }
  return [...times].sort((a, b) => a - b);
}

/**
 * The `@keyframes` body of a recipe for one object box. Timing is baked into
 * the samples, so the rule plays with a `linear` timing function.
 */
export function effectRecipeKeyframeBody(
  recipe: EffectRecipeVariant,
  cls: "entr" | "exit" | "emph",
  box: EffectRecipeBox,
  /** The even grid's step count; a long recipe (an animated background's
   * minutes-long loop, PPT-A134) asks for more so each stage stays smooth. */
  options: { steps?: number } = {},
): string {
  const frames = sampleTimes(recipe, options.steps).map((at) => sampleEffectRecipe(recipe, cls, box, at));
  const channels = frameChannels(frames);
  let previous = "";
  const rules: string[] = [];
  frames.forEach((frame, index) => {
    const css = frameCss(frame, channels);
    // A run of identical frames keeps its first and last only.
    const next = index + 1 < frames.length ? frameCss(frames[index + 1], channels) : "";
    if (css === previous && css === next) return;
    previous = css;
    rules.push(`${round(frame.at * 100, 3)}%{${css}}`);
  });
  return rules.join("");
}

/** A stable keyframe name for (effect, option, box). */
export function effectRecipeKeyframeName(
  effect: string,
  option: string | null | undefined,
  box: EffectRecipeBox,
): string {
  const key = [box.x, box.y, box.w, box.h].map((value) => Math.round(value * 10000)).join("_");
  return `mrai-ppt-rx-${effect}${option ? `-${option}` : ""}-${key}`.replace(/[^A-Za-z0-9_-]/g, "_");
}

/** The class of a recipe effect (`entr` / `exit` / `emph`), or `null`. */
export function effectRecipeClass(effect: string | undefined): "entr" | "exit" | "emph" | null {
  if (!effect || !Object.prototype.hasOwnProperty.call(RECIPES, effect)) return null;
  return RECIPES[effect].cls;
}

/** Every effect id that has a recipe. */
export const EFFECT_RECIPE_IDS: readonly string[] = Object.freeze(
  Object.keys(RECIPES).filter((id) => !RECIPES[id].ext),
);

/**
 * What a recipe leaves on its object once it has played and is held: the last
 * frame's declarations (`null` when the object is simply back at rest). Used
 * to keep a held emphasis in force while later effects play.
 */
export function effectRecipeSettledStyle(
  recipe: EffectRecipeVariant,
  cls: "entr" | "exit" | "emph",
  box: EffectRecipeBox,
): Record<string, string> | null {
  const frame = sampleEffectRecipe(recipe, cls, box, 1);
  const style: Record<string, string> = {};
  if (frame.hsl) {
    const filter = hslFilter(frame.hsl);
    if (filter !== "none") style.filter = filter;
  }
  if (frame.bold) style.WebkitTextStroke = "0.035em currentColor";
  if (frame.italic) style.fontStyle = "italic";
  if (frame.underline) style.textDecorationLine = "underline";
  if (frame.fontFamily) style.fontFamily = frame.fontFamily;
  if (frame.color) style.color = effectRecipeSchemeColor(frame.color);
  // A held size or turn (Blast ends larger): the caller composes it with the
  // object's own transform.
  const moved =
    Math.abs(frame.scaleX - 1) > 1e-4 ||
    Math.abs(frame.scaleY - 1) > 1e-4 ||
    Math.abs(frame.rotate % 360) > 1e-4;
  // A motion path that does not close leaves the object where it ends.
  const away = Math.abs(frame.dx) > 1e-5 || Math.abs(frame.dy) > 1e-5;
  if (moved || away) {
    style.transform =
      (away ? `translate(${round(frame.dx * 100)}cqw,${round(frame.dy * 100)}cqh) ` : "") +
      `rotate(${round(frame.rotate)}deg) scale(${round(frame.scaleX, 4)},${round(frame.scaleY, 4)})`;
  }
  return Object.keys(style).length > 0 ? style : null;
}

export interface EffectRecipeColorLayer {
  /** The theme colour the effect changes the object TO. */
  scheme: string;
  fill: boolean;
  line: boolean;
  text: boolean;
  /** The colour stays once the effect has played. */
  holds: boolean;
  /** `@keyframes` body of the colour layer's opacity (0 = the object's own colour). */
  keyframeBody: string;
}

/**
 * The colour an effect changes its object to, and how the change runs: the
 * caller draws it as a layer in that colour over the fill / line / text, with
 * this opacity. `null` for an effect that changes to no colour.
 */
export function effectRecipeColorLayer(
  recipe: EffectRecipeVariant,
): EffectRecipeColorLayer | null {
  let scheme: string | null = null;
  let fill = false;
  let line = false;
  let text = false;
  let timing: TrackTiming | null = null;
  for (const track of recipe.tracks) {
    if (track.k === "clr" && track.to && track.to.sp === "schemeClr" && track.to.val) {
      scheme = scheme ?? track.to.val;
      timing = timing ?? track;
      const target = track.p.toLowerCase();
      if (target === "fillcolor") fill = true;
      else if (target === "stroke.color") line = true;
      else if (target === "style.color") text = true;
    } else if (track.k === "prop" && track.p === "style.color" && track.v) {
      scheme = scheme ?? track.v;
      timing = timing ?? track;
      text = true;
    }
  }
  if (!scheme || !timing) return null;
  const times = new Set<number>([0, 1]);
  for (let i = 1; i < 40; i += 1) times.add(i / 40);
  times.add(timing.b);
  times.add(Math.min(1, timing.b + timing.d));
  if (timing.ar) times.add(Math.min(1, timing.b + 2 * timing.d));
  const span = timing;
  const body = [...times]
    .filter((at) => at >= 0 && at <= 1)
    .sort((a, b) => a - b)
    .map((at) => {
      const u = ease(at, recipe.ac, recipe.dc);
      const progress = trackProgress(span, u);
      return `${round(at * 100, 3)}%{opacity:${round(progress ?? 0)}}`;
    })
    .join("");
  return { scheme, fill, line, text, holds: Boolean(timing.hold) && !timing.ar, keyframeBody: body };
}

/**
 * How a recipe iterates over text: the time between two letters' starts for an
 * effect of `durationMs`, or `null` for an effect that plays on the object as
 * a whole.
 */







export function effectPlayDurationMs(
  step: {
    animation?: string;
    effectOption?: string | null;
    durationMs: number;
    paragraphs?: unknown;
  },
  block: { type?: unknown; text?: unknown } | undefined,
): number {
  const recipe = effectRecipe(step.animation, step.effectOption);
  const stepMs = recipe ? effectRecipeLetterStepMs(recipe, step.durationMs) : null;
  if (stepMs === null || step.paragraphs) return step.durationMs;
  const text =
    block && (block.type === "text" || block.type === "shape") && typeof block.text === "string"
      ? block.text
      : "";
  let letters = 0;
  for (const character of text) if (character !== "\n") letters += 1;
  if (letters === 0) return step.durationMs;
  return step.durationMs + stepMs * (letters - 1);
}

export function effectRecipeLetterStepMs(
  recipe: EffectRecipeVariant,
  durationMs: number,
): number | null {
  if (!recipe.it) return null;
  if (typeof recipe.it.abs === "number") return Math.max(0, recipe.it.abs);
  if (typeof recipe.it.pct === "number") return Math.max(0, recipe.it.pct * durationMs);
  return null;
}
