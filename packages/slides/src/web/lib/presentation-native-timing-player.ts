/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Plays a slide's animations as its timing tree states them (render only).
 *
 * The server lays the main sequence of an imported slide out per click when
 * none of the editable readers recognises it
 * (`apps/api/.../pptx/animation_timing_tree.py`, `slide.nativeTimingTree`).
 * Each behaviour becomes a Web Animation on the object it names (its
 * `data-presentation-block-id` element, or a paragraph of it), starting at
 * its own time in the click's run (ECMA-376 §19.5, as PowerPoint plays it):
 *
 * - `set` and `anim` on the properties a show uses -- `ppt_x` / `ppt_y` (the
 *   centre, a share of the slide) as a translation, `ppt_w` / `ppt_h` as a
 *   scale, rotation, opacity, visibility, text colour -- with their keyframes'
 *   formulas evaluated (`#ppt_x`, `#ppt_w`, `sin`, `$` ...);
 * - `animEffect` filters (wipe, circle, box, barn, diamond, plus, slide,
 *   fade, dissolve ...) as clip paths or fades, in or out;
 * - `animMotion` paths, `animScale`, `animRot`, `animClr`.
 *
 * Animations on one element compose (`composite: "add"` on the independent
 * `translate` / `scale` / `rotate` properties, which leave the element's own
 * `transform` alone). Being Web Animations, a recording sets their time like
 * any other (`PresentationVideoExporter`).
 */

import { patternRevealKeyframes } from "./presentation-anim-pattern-reveal";
import { effectRevealClip } from "@mrnote/core/web/lib/presentation-effect-recipes.ts";

export type NativeTimingBehaviour = {
  kind: "set" | "anim" | "animEffect" | "animMotion" | "animScale" | "animRot" | "animClr" | "cmd" | "sound";
  /** An effect's sound (`kind` "sound"): its digest in the tree's ``sounds``, and its volume 0..1. */
  sound?: string;
  volume?: number;
  /** A media call's command (`playFrom(0.0)`, `togglePause`, `pause`,
   * `resume`, `stop`) on the clip `target` names. */
  cmd?: string;
  cmdType?: string;
  target: string;
  paragraphs?: [number, number];
  start: number;
  dur: number;
  repeat?: number | "indefinite";
  autoRev?: boolean;
  accel?: number;
  decel?: number;
  fill?: string;
  attrs?: string[];
  additive?: string;
  presetClass?: string;
  to?: unknown;
  from?: unknown;
  by?: unknown;
  calcmode?: string;
  valueType?: string;
  keys?: Array<{ tm: number | null; val: unknown; fmla?: string }>;
  transition?: string;
  filter?: string;
  path?: string;
  origin?: string;
  /** The node runs backwards (`spd` < 0). */
  reverse?: boolean;
  /** The time container (the effect) the behaviour plays in, when that
   * container repeats, reverses, eases, runs at another speed, removes what
   * it did at its end (Rewind when done playing) or ends at the next click. */
  wrap?: NativeTimingWrap;
};

export type NativeTimingWrap = {
  /** When the container begins in the click's run, ms. */
  begin: number;
  /** One pass of the container, ms (its speed applied). */
  span: number;
  repeat?: number | "indefinite";
  autoRev?: boolean;
  accel?: number;
  decel?: number;
  reverse?: boolean;
  remove?: boolean;
  untilNext?: boolean;
};

export type NativeTimingRun = { length: number; behaviours: NativeTimingBehaviour[] };

export type NativeTimingTree = {
  schema: "pptx-timing-tree-v1";
  clicks: Array<{ auto: boolean } & NativeTimingRun>;
  hiddenUntil: Record<string, number>;
  /** "Trigger: On Click of": a click on `block` plays its next run. */
  triggers?: Array<{ block: string; runs: NativeTimingRun[] }>;
};

export function readNativeTimingTree(value: unknown): NativeTimingTree | null {
  if (typeof value !== "object" || value === null) return null;
  const tree = value as Partial<NativeTimingTree>;
  if (tree.schema !== "pptx-timing-tree-v1" || !Array.isArray(tree.clicks)) return null;
  return {
    schema: tree.schema,
    clicks: tree.clicks.filter(
      (click) => typeof click === "object" && click !== null && Array.isArray((click as { behaviours?: unknown }).behaviours),
    ),
    hiddenUntil: typeof tree.hiddenUntil === "object" && tree.hiddenUntil !== null ? tree.hiddenUntil : {},
    ...(Array.isArray(tree.triggers)
      ? {
          triggers: tree.triggers.filter(
            (trigger) =>
              typeof trigger === "object" &&
              trigger !== null &&
              typeof trigger.block === "string" &&
              Array.isArray(trigger.runs) &&
              trigger.runs.every((run) => typeof run === "object" && run !== null && Array.isArray(run.behaviours)),
          ),
        }
      : {}),
  };
}

/** The objects whose click plays one of the tree's interactive sequences. */
export function nativeTimingTriggerBlocks(tree: NativeTimingTree | null): ReadonlySet<string> {
  return new Set((tree?.triggers ?? []).map((trigger) => trigger.block));
}

/** The run the `count`-th click (1-based) on `block` plays: each click the
 * next group of each of its sequences, the first again after the last. */
export function nativeTimingTriggerRuns(tree: NativeTimingTree, block: string, count: number): NativeTimingRun[] {
  if (count <= 0) return [];
  return (tree.triggers ?? [])
    .filter((trigger) => trigger.block === block && trigger.runs.length > 0)
    .map((trigger) => trigger.runs[(count - 1) % trigger.runs.length]!);
}

/** The tree a slide plays: only when it has no sequence of its own. */
export function slideNativeTimingTree(slide: { animationSequence?: readonly unknown[]; nativeTimingTree?: unknown }): NativeTimingTree | null {
  if (slide.animationSequence && slide.animationSequence.length > 0) return null;
  return readNativeTimingTree(slide.nativeTimingTree);
}

/** Click groups the show steps through on the slide (its first may play by itself). */
export function nativeTimingClickCount(tree: NativeTimingTree): number {
  return tree.clicks.length;
}

/** The step the slide opens at: 1 when its first run plays by itself. */
export function nativeTimingInitialStep(tree: NativeTimingTree): number {
  return tree.clicks[0]?.auto ? 1 : 0;
}

// ---------------------------------------------------------------------------
// Formulas (ECMA-376 §19.5.86 fmla): + - * / ^, unary minus, parentheses,
// numbers, #ppt_x #ppt_y #ppt_w #ppt_h, $ (the time), and abs acos asin atan
// ceil cos deg exp floor ln max min rad rand sin sqrt tan.
// ---------------------------------------------------------------------------

type Variables = Record<string, number>;

export function evaluateTimingFormula(text: string, variables: Variables): number {
  const tokens = text.match(/#?[A-Za-z_][A-Za-z0-9_]*|\$|\d*\.?\d+(?:[eE][-+]?\d+)?|[-+*/^(),]/g) ?? [];
  let at = 0;
  const peek = () => tokens[at];
  const take = () => tokens[at++];
  const functions: Record<string, (...args: number[]) => number> = {
    abs: Math.abs, acos: Math.acos, asin: Math.asin, atan: Math.atan, ceil: Math.ceil, cos: Math.cos,
    deg: (x) => (x * 180) / Math.PI, exp: Math.exp, floor: Math.floor, ln: Math.log, max: Math.max, min: Math.min,
    rad: (x) => (x * Math.PI) / 180, rand: (x) => Math.random() * x, sin: Math.sin, sqrt: Math.sqrt, tan: Math.tan,
  };
  const primary = (): number => {
    const token = take();
    if (token === undefined) return 0;
    if (token === "(") {
      const value = expression();
      if (peek() === ")") take();
      return value;
    }
    if (token === "-") return -primary();
    if (token === "+") return primary();
    if (token === "$") return variables.$ ?? 0;
    if (token.startsWith("#")) return variables[token.slice(1)] ?? 0;
    if (/^[A-Za-z_]/.test(token)) {
      const fn = functions[token.toLowerCase()];
      if (fn && peek() === "(") {
        take();
        const args: number[] = [];
        if (peek() !== ")") {
          args.push(expression());
          while (peek() === ",") {
            take();
            args.push(expression());
          }
        }
        if (peek() === ")") take();
        return fn(...args);
      }
      return variables[token] ?? 0;
    }
    return Number(token);
  };
  const power = (): number => {
    let base = primary();
    while (peek() === "^") {
      take();
      base = Math.pow(base, primary());
    }
    return base;
  };
  const term = (): number => {
    let value = power();
    while (peek() === "*" || peek() === "/") {
      const op = take();
      const right = power();
      value = op === "*" ? value * right : right === 0 ? 0 : value / right;
    }
    return value;
  };
  function expression(): number {
    let value = term();
    while (peek() === "+" || peek() === "-") {
      const op = take();
      const right = term();
      value = op === "+" ? value + right : value - right;
    }
    return value;
  }
  const result = expression();
  return Number.isFinite(result) ? result : 0;
}

// ---------------------------------------------------------------------------
// Playing.
// ---------------------------------------------------------------------------

type Frame = { left: number; top: number; width: number; height: number };

export type NativeTimingStage = {
  /** The slide's design size, px. */
  readonly width: number;
  readonly height: number;
  /** Each object's frame on the slide, design px. */
  readonly frames: ReadonlyMap<string, Frame>;
  /** Where an effect's sound (its digest in the tree's ``sounds``) plays from; absent, effects are silent. */
  readonly soundUrl?: (digest: string) => string | null;
};

/** The slide's own drawing of an object (not a Morph layer's copy of it). */
function blockElement(root: ParentNode, id: string): HTMLElement | null {
  for (const element of Array.from(root.querySelectorAll<HTMLElement>(`[data-presentation-block-id="${CSS.escape(id)}"]`))) {
    if (!element.closest(".presentation-present__morph")) return element;
  }
  return null;
}

function targetsOf(root: ParentNode, behaviour: NativeTimingBehaviour): HTMLElement[] {
  const element = blockElement(root, behaviour.target);
  if (!element) return [];
  if (!behaviour.paragraphs) return [element];
  const rows = Array.from(
    element.querySelectorAll<HTMLElement>('[data-testid="presentation-paragraph"], [data-testid="presentation-text-row"]'),
  );
  const [first, last] = behaviour.paragraphs;
  const picked = rows.slice(first, last + 1);
  return picked.length > 0 ? picked : [element];
}

function easingOf(accel: number, decel: number): string {
  return accel > 0 || decel > 0 ? `cubic-bezier(${Math.min(1, accel)}, 0, ${Math.max(0, 1 - decel)}, 1)` : "linear";
}

function timingOf(behaviour: NativeTimingBehaviour): KeyframeAnimationOptions {
  const accel = behaviour.accel ?? 0;
  const decel = behaviour.decel ?? 0;
  const iterations = behaviour.repeat === "indefinite" ? Infinity : typeof behaviour.repeat === "number" ? behaviour.repeat : 1;
  const wrap = behaviour.wrap;
  if (wrap) {
    // The behaviour's own pass is folded into its container's keyframes
    // (`wrapKeyframes`); the animation is the container's.
    const passes = wrap.repeat === "indefinite" ? Infinity : typeof wrap.repeat === "number" ? wrap.repeat : 1;
    const removes = wrap.remove === true || behaviour.fill === "remove";
    return {
      delay: wrap.begin,
      duration: Math.max(1, wrap.span),
      iterations: wrap.autoRev ? passes * 2 : passes,
      direction: wrap.autoRev ? (wrap.reverse ? "alternate-reverse" : "alternate") : wrap.reverse ? "reverse" : "normal",
      fill: removes ? (behaviour.kind === "set" ? "none" : "backwards") : behaviour.kind === "set" ? "forwards" : "both",
      easing: easingOf(wrap.accel ?? 0, wrap.decel ?? 0),
    };
  }
  return {
    delay: behaviour.start,
    duration: Math.max(1, behaviour.dur),
    iterations: behaviour.autoRev ? iterations * 2 : iterations,
    direction: behaviour.autoRev
      ? behaviour.reverse ? "alternate-reverse" : "alternate"
      : behaviour.reverse ? "reverse" : "normal",
    // A run is held where it ends unless it says to remove itself. A `set`
    // takes effect at its own time only: held before it, an entrance's
    // visibility would show the object while its delay runs.
    fill: behaviour.fill === "remove" ? (behaviour.kind === "set" ? "none" : "backwards") : behaviour.kind === "set" ? "forwards" : "both",
    easing: easingOf(accel, decel),
  };
}

/** At most this many of a behaviour's own passes are drawn into its
 * container's keyframes; past it the last pass holds. */
const MAX_FOLDED_PASSES = 40;

/**
 * A behaviour's keyframes over one pass of the container it plays in: its own
 * passes (repeat, auto-reverse, reversed, eased) laid at its own time inside
 * the container's, the value before it begins and after it ends held.
 */
export function wrapKeyframes(behaviour: NativeTimingBehaviour, frames: Keyframe[]): Keyframe[] {
  const wrap = behaviour.wrap;
  if (!wrap || frames.length === 0) return frames;
  const span = Math.max(1, wrap.span);
  const own = Math.max(1, behaviour.dur);
  const rawPasses = behaviour.repeat === "indefinite" ? 1 : typeof behaviour.repeat === "number" ? behaviour.repeat : 1;
  const passes = Math.min(MAX_FOLDED_PASSES, Math.max(1, Math.ceil(rawPasses)));
  const legs = behaviour.autoRev ? passes * 2 : passes;
  const local = Math.max(0, behaviour.start - wrap.begin);
  const plain = frames.map((frame, index) => ({
    ...frame,
    offset: typeof frame.offset === "number" ? frame.offset : frames.length === 1 ? 0 : index / (frames.length - 1),
  }));
  const ease = easingOf(behaviour.accel ?? 0, behaviour.decel ?? 0);
  const leg = (backwards: boolean): Keyframe[] => {
    const ordered = backwards ? [...plain].reverse().map((frame) => ({ ...frame, offset: 1 - (frame.offset as number) })) : plain;
    return ordered;
  };
  const out: Keyframe[] = [];
  const at = (time: number) => Math.min(1, Math.max(0, time / span));
  // Before its own time a behaviour shows its first value -- except a `set`,
  // which takes effect at its time only: held from the container's start, an
  // exit's "hidden" at 999 ms hid the shape for the whole of its fade-out
  // (probe slide 5: a faded-out, rewound shape was gone from the first frame,
  // PowerPoint shows it fading), and an entrance's "visible" showed the shape
  // before its delay. Its start is left implicit: the shape's own value.
  if (local > 0 && behaviour.kind !== "set") out.push({ ...plain[0]!, offset: 0, easing: "linear" });
  for (let index = 0; index < legs; index += 1) {
    const backwards = (behaviour.reverse === true) !== (behaviour.autoRev === true && index % 2 === 1);
    const legStart = local + index * own;
    for (const frame of leg(backwards)) {
      const { easing: frameEasing, ...rest } = frame;
      out.push({
        ...rest,
        offset: at(legStart + (frame.offset as number) * own),
        easing: ease !== "linear" ? ease : frameEasing ?? "linear",
      });
    }
  }
  const last = out[out.length - 1]!;
  if ((last.offset as number) < 1) out.push({ ...last, offset: 1, easing: "linear" });
  // Offsets must not decrease (a pass cut at the container's end piles at 1).
  let floor = 0;
  for (const frame of out) {
    floor = Math.max(floor, frame.offset as number);
    frame.offset = floor;
  }
  return out;
}

/** Animations that end at the next click (Repeat: Until Next Click). */
const endsAtNextClick = new WeakSet<Animation>();

/** Does `animation` stop at the next click of the show? */
export function nativeTimingEndsAtNextClick(animation: Animation): boolean {
  return endsAtNextClick.has(animation);
}

/**
 * Bring a run to its end, as the show's next click does (`nextAc="seek"`):
 * what has an end jumps there; what repeats Until Next Click stops -- where it
 * is while it plays, at the end of its first pass when the run is stepped
 * past unseen; what repeats until the end of the slide plays on.
 */
export function endNativeTimingRun(animations: readonly Animation[], unseen: boolean): void {
  for (const animation of animations) {
    try {
      if (unseen && mediaCalls.has(animation)) {
        // A clip call in a run the show stepped past unseen is not made.
        animation.cancel();
      } else if (endsAtNextClick.has(animation)) {
        if (unseen) {
          const timing = animation.effect?.getTiming();
          const delay = typeof timing?.delay === "number" ? timing.delay : 0;
          const duration = typeof timing?.duration === "number" ? timing.duration : 0;
          animation.currentTime = delay + duration;
        }
        animation.pause();
      } else {
        animation.finish();
      }
    } catch {
      // An endless one plays on.
    }
  }
}

const PROPERTY_OF: Record<string, string> = {
  "style.opacity": "opacity",
  "style.visibility": "visibility",
  "style.rotation": "rotate",
  r: "rotate",
  "style.color": "color",
  ppt_x: "x",
  ppt_y: "y",
  ppt_w: "w",
  ppt_h: "h",
};

function keyValues(behaviour: NativeTimingBehaviour, variables: Variables): Array<{ offset: number; value: number | string }> {
  const keys = behaviour.keys ?? [];
  if (keys.length > 0) {
    const stated = keys.map((key, index) => key.tm ?? (keys.length === 1 ? 0 : index / (keys.length - 1)));
    return keys.map((key, index) => {
      const local = { ...variables, $: stated[index] ?? 0 };
      let value: number | string;
      if (typeof key.val === "number") value = key.val;
      else if (typeof key.val === "string") value = /[#$a-z(]/i.test(key.val) ? evaluateTimingFormula(key.val, local) : Number(key.val);
      else value = 0;
      if (key.fmla) value = evaluateTimingFormula(key.fmla.replace(/\$/g, `(${String(value)})`), local);
      return { offset: Math.min(1, Math.max(0, stated[index] ?? 0)), value };
    });
  }
  const parse = (raw: unknown, fallback: number) =>
    typeof raw === "number" ? raw : typeof raw === "string" ? evaluateTimingFormula(raw, variables) : fallback;
  const from = parse(behaviour.from, variables.current ?? 0);
  const to = behaviour.by !== undefined ? from + parse(behaviour.by, 0) : parse(behaviour.to, from);
  return [
    { offset: 0, value: from },
    { offset: 1, value: to },
  ];
}

/** The keyframes of an `animEffect` filter, entering or leaving. */
export function nativeTimingFilterKeyframes(filter: string, entering: boolean): Keyframe[] {
  const [name, rawOption] = filter.split(/[()]/);
  const option = (rawOption ?? "").toLowerCase();
  let shown: Keyframe;
  let hidden: Keyframe;
  switch ((name ?? "").trim().toLowerCase()) {
    case "wipe": {
      const start =
        option === "up" ? "inset(100% 0 0 0)" : option === "left" ? "inset(0 0 0 100%)" : option === "right" ? "inset(0 100% 0 0)" : "inset(0 0 100% 0)";
      hidden = { clipPath: start };
      shown = { clipPath: "inset(0 0 0 0)" };
      break;
    }
    case "circle":
    case "box":
    case "diamond":
    case "plus":
    case "wedge":
    case "wheel":
    case "strips": {
      // The editor's own reveal shapes (`effectRevealClip`), with the
      // filter's direction: `out` grows from the centre, `in` closes in from
      // the edges; a wheel's spokes; a strip's corner. Sampled along the run
      // (a sweep's edge moves on an arc, not a straight line).
      const kind = name!.trim().toLowerCase();
      const samples = kind === "wheel" || kind === "wedge" ? 24 : 12;
      return Array.from({ length: samples + 1 }, (_, index) => {
        const progress = index / samples;
        return {
          offset: progress,
          clipPath: effectRevealClip(kind, rawOption?.trim(), entering ? progress : 1 - progress) ?? "none",
        };
      });
    }
    case "barn":
      hidden = { clipPath: option.includes("horizontal") ? "inset(50% 0 50% 0)" : "inset(0 50% 0 50%)" };
      shown = { clipPath: "inset(0 0 0 0)" };
      break;
    case "blinds":
    case "checkerboard":
    case "randombar": {
      // The editor's own piece-by-piece reveals (`presentation-anim-pattern-reveal`):
      // `horizontal` / `across` lay the pieces along rows, `vertical` / `down`
      // along columns.
      const kind = name!.trim().toLowerCase() === "randombar" ? "random-bars" : (name!.trim().toLowerCase() as "blinds" | "checkerboard");
      const direction = option === "vertical" || option === "down" ? "vertical" : "horizontal";
      return patternRevealKeyframes(kind, direction, entering);
    }
    case "slide": {
      const offset =
        option === "fromtop" ? "0 -100%" : option === "frombottom" ? "0 100%" : option === "fromleft" ? "-100% 0" : "100% 0";
      hidden = { translate: offset, opacity: 0 };
      shown = { translate: "0 0", opacity: 1 };
      break;
    }
    default:
      // fade, dissolve, image and the patterned reveals: a fade.
      hidden = { opacity: 0 };
      shown = { opacity: 1 };
  }
  return entering ? [hidden, shown] : [shown, hidden];
}

/** Points along a motion path (`M x y L x y C ... E`, shares of the slide). */
function motionPoints(path: string, stage: NativeTimingStage): Array<[number, number]> {
  const tokens = path.match(/[MLCZEmlcze]|-?\d*\.?\d+(?:[eE][-+]?\d+)?/g) ?? [];
  const points: Array<[number, number]> = [];
  let at = 0;
  let command = "M";
  let current: [number, number] = [0, 0];
  const number = () => Number(tokens[at++] ?? 0);
  const point = (): [number, number] => [number() * stage.width, number() * stage.height];
  while (at < tokens.length) {
    if (/[A-Za-z]/.test(tokens[at]!)) command = tokens[at++]!.toUpperCase();
    if (command === "M" || command === "L") {
      current = point();
      points.push(current);
    } else if (command === "C") {
      const c1 = point();
      const c2 = point();
      const end = point();
      for (let step = 1; step <= 12; step += 1) {
        const t = step / 12;
        const u = 1 - t;
        points.push([
          u * u * u * current[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * end[0],
          u * u * u * current[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * end[1],
        ]);
      }
      current = end;
    } else {
      // Z / E: close or end.
      at += 1;
    }
  }
  return points.length > 0 ? points : [[0, 0]];
}

function evenlyAlong(points: Array<[number, number]>, count: number): Array<[number, number]> {
  const lengths = [0];
  for (let index = 1; index < points.length; index += 1) {
    const [ax, ay] = points[index - 1]!;
    const [bx, by] = points[index]!;
    lengths.push(lengths[index - 1]! + Math.hypot(bx - ax, by - ay));
  }
  const total = lengths[lengths.length - 1]!;
  if (total === 0) return [points[0]!];
  const out: Array<[number, number]> = [];
  let segment = 1;
  for (let step = 0; step <= count; step += 1) {
    const goal = (step / count) * total;
    while (segment < lengths.length - 1 && lengths[segment]! < goal) segment += 1;
    const span = lengths[segment]! - lengths[segment - 1]!;
    const t = span > 0 ? (goal - lengths[segment - 1]!) / span : 0;
    const [ax, ay] = points[segment - 1]!;
    const [bx, by] = points[segment]!;
    out.push([ax + (bx - ax) * t, ay + (by - ay) * t]);
  }
  return out;
}

function keyframesOf(behaviour: NativeTimingBehaviour, stage: NativeTimingStage): { frames: Keyframe[]; composite?: CompositeOperation } | null {
  const frame = stage.frames.get(behaviour.target);
  const variables: Variables = frame
    ? {
        ppt_x: (frame.left + frame.width / 2) / stage.width,
        ppt_y: (frame.top + frame.height / 2) / stage.height,
        ppt_w: frame.width / stage.width,
        ppt_h: frame.height / stage.height,
      }
    : { ppt_x: 0.5, ppt_y: 0.5, ppt_w: 0, ppt_h: 0 };
  switch (behaviour.kind) {
    case "set": {
      const property = PROPERTY_OF[behaviour.attrs?.[0] ?? ""];
      if (!property || property === "x" || property === "y" || property === "w" || property === "h") return null;
      const value = property === "rotate" ? `${Number(behaviour.to) || 0}deg` : String(behaviour.to ?? "");
      return { frames: [{ [property]: value }, { [property]: value }] };
    }
    case "anim": {
      const property = PROPERTY_OF[behaviour.attrs?.[0] ?? ""];
      if (!property) return null;
      const values = keyValues(behaviour, { ...variables, current: property === "x" ? variables.ppt_x! : property === "y" ? variables.ppt_y! : property === "w" ? variables.ppt_w! : property === "h" ? variables.ppt_h! : 0 });
      const discrete = behaviour.calcmode === "discrete";
      const frames: Keyframe[] = values.map(({ offset, value }) => {
        const n = typeof value === "number" ? value : Number(value);
        const keyframe: Keyframe = { offset };
        if (property === "x") keyframe.translate = `${(n - variables.ppt_x!) * stage.width}px 0px`;
        else if (property === "y") keyframe.translate = `0px ${(n - variables.ppt_y!) * stage.height}px`;
        else if (property === "w") keyframe.scale = `${variables.ppt_w ? n / variables.ppt_w : 1} 1`;
        else if (property === "h") keyframe.scale = `1 ${variables.ppt_h ? n / variables.ppt_h : 1}`;
        else if (property === "rotate") keyframe.rotate = `${n}deg`;
        else if (property === "opacity") keyframe.opacity = Math.min(1, Math.max(0, n));
        else keyframe[property] = String(value);
        if (discrete) keyframe.easing = "steps(1, end)";
        return keyframe;
      });
      return { frames, composite: property === "x" || property === "y" || property === "w" || property === "h" || property === "rotate" ? "add" : "replace" };
    }
    case "animEffect": {
      if (!behaviour.filter) return null;
      return { frames: nativeTimingFilterKeyframes(behaviour.filter, behaviour.transition !== "out") };
    }
    case "animMotion": {
      if (!behaviour.path) return null;
      const points = evenlyAlong(motionPoints(behaviour.path, stage), 40);
      const [ox, oy] = points[0]!;
      return {
        frames: points.map(([x, y], index) => ({ offset: index / (points.length - 1 || 1), translate: `${x - ox}px ${y - oy}px` })),
        composite: "add",
      };
    }
    case "animScale": {
      const by = Array.isArray(behaviour.by) ? (behaviour.by as number[]) : null;
      const from = Array.isArray(behaviour.from) ? (behaviour.from as number[]) : [1, 1];
      const to = Array.isArray(behaviour.to) ? (behaviour.to as number[]) : by ?? [1, 1];
      return { frames: [{ scale: `${from[0]} ${from[1]}` }, { scale: `${to[0]} ${to[1]}` }], composite: "add" };
    }
    case "animRot": {
      const from = typeof behaviour.from === "number" ? behaviour.from : 0;
      const to = typeof behaviour.to === "number" ? behaviour.to : from + (typeof behaviour.by === "number" ? behaviour.by : 0);
      return { frames: [{ rotate: `${from}deg` }, { rotate: `${to}deg` }], composite: "add" };
    }
    case "animClr": {
      if (typeof behaviour.to !== "string" || !behaviour.attrs?.some((attr) => attr === "style.color")) return null;
      return { frames: [{}, { color: behaviour.to }] };
    }
    default:
      return null;
  }
}

/** Hide what an entrance has yet to bring in, given how many runs played. */
export function applyNativeTimingVisibility(root: ParentNode, tree: NativeTimingTree, played: number): void {
  for (const [target, click] of Object.entries(tree.hiddenUntil)) {
    const element = blockElement(root, target);
    if (!element) continue;
    element.style.visibility = click >= played ? "hidden" : "";
  }
}

/**
 * Play run `index` of the tree on the slide drawn under `root`. `leadMs` holds
 * the run back that long: a run the slide starts by itself waits for the
 * slide's transition, as PowerPoint plays the transition first (the slide's
 * timeline starts when it ends).
 */
export function playNativeTimingClick(
  root: ParentNode,
  tree: NativeTimingTree,
  index: number,
  stage: NativeTimingStage,
  leadMs = 0,
): Animation[] {
  const run = tree.clicks[index];
  return run ? playNativeTimingRun(root, run, stage, leadMs) : [];
}

/** Media calls waiting for their time: stepping past them unseen drops them. */
const mediaCalls = new WeakSet<Animation>();

/** Carry out a media call on a clip (ECMA-376 19.5.17 cmd, PowerPoint's
 * media verbs). */
export function runNativeMediaCommand(media: HTMLMediaElement, command: string): void {
  const from = /^playFrom\(\s*(-?[0-9.]+)\s*\)$/.exec(command.trim());
  const play = () => {
    void media.play()?.catch(() => {
      // Autoplay refused (no gesture yet): the clip's own controls remain.
    });
  };
  if (from) {
    const at = Number(from[1]);
    if (Number.isFinite(at) && at >= 0) {
      try {
        media.currentTime = at;
      } catch {
        // Not seekable yet: plays from where it is.
      }
    }
    play();
  } else if (command === "togglePause") {
    if (media.paused) play();
    else media.pause();
  } else if (command === "pause") {
    media.pause();
  } else if (command === "resume") {
    play();
  } else if (command === "stop") {
    media.pause();
    try {
      media.currentTime = 0;
    } catch {
      // Nothing to rewind.
    }
  }
}

function scheduleMediaCall(root: ParentNode, behaviour: NativeTimingBehaviour, leadMs: number): Animation | null {
  const element = blockElement(root, behaviour.target);
  const media = element?.querySelector<HTMLMediaElement>("video, audio");
  if (!element || !media || !behaviour.cmd) return null;
  // A keyframe-less animation is the call's clock, so the show's seek
  // (`endNativeTimingRun`) and a recording's time govern it like the rest.
  const clock = element.animate([], { delay: behaviour.start + leadMs, duration: 1, fill: "none" });
  mediaCalls.add(clock);
  const command = behaviour.cmd;
  clock.finished.then(
    () => runNativeMediaCommand(media, command),
    () => {
      // Cancelled (the show went back): the call never happens.
    },
  );
  return clock;
}

/**
 * An effect's sound (Effect Options > Sound), started with the effect: a
 * hidden player under the slide, so it stops with the slide and a recording
 * hears it (`export-media`). On the run's own clock, like a media call:
 * stepped past unseen, it is not played.
 */
function scheduleEffectSound(root: ParentNode, behaviour: NativeTimingBehaviour, stage: NativeTimingStage, leadMs: number): Animation | null {
  const url = behaviour.sound && stage.soundUrl ? stage.soundUrl(behaviour.sound) : null;
  const host = root instanceof Element ? root : (root as Document).body ?? null;
  if (!url || !host) return null;
  const clock = host.animate([], { delay: behaviour.start + leadMs, duration: 1, fill: "none" });
  mediaCalls.add(clock);
  clock.finished.then(
    () => {
      const audio = host.ownerDocument.createElement("audio");
      audio.src = url;
      audio.hidden = true;
      audio.preload = "auto";
      audio.volume = Math.max(0, Math.min(1, typeof behaviour.volume === "number" ? behaviour.volume : 1));
      audio.dataset.nativeEffectSound = behaviour.sound!;
      // It plays its length and holds no slide (a recording does not wait for it).
      audio.dataset.playbackAcross = "true";
      audio.addEventListener("ended", () => audio.remove(), { once: true });
      audio.autoplay = true;
      host.append(audio);
      try {
        void audio.play()?.catch(() => undefined);
      } catch {
        // Autoplay alone then (or a page that cannot play it).
      }
    },
    () => {
      // Cancelled (the show went back): never heard.
    },
  );
  return clock;
}

/** Play one run (a click's, or a trigger's) on the slide drawn under `root`. */
export function playNativeTimingRun(
  root: ParentNode,
  run: NativeTimingRun,
  stage: NativeTimingStage,
  leadMs = 0,
): Animation[] {
  const animations: Animation[] = [];
  const lead = Number.isFinite(leadMs) && leadMs > 0 ? leadMs : 0;
  for (const behaviour of run.behaviours) {
    if (behaviour.kind !== "cmd" && behaviour.kind !== "sound") continue;
    try {
      const clock =
        behaviour.kind === "sound" ? scheduleEffectSound(root, behaviour, stage, lead) : scheduleMediaCall(root, behaviour, lead);
      if (clock) animations.push(clock);
    } catch {
      // No clip to call, no sound to play: the rest still play.
    }
  }
  for (const behaviour of run.behaviours) {
    const keyframes = keyframesOf(behaviour, stage);
    if (!keyframes) continue;
    for (const element of targetsOf(root, behaviour)) {
      // An entrance's `set visibility` overrides the inline hiding while it
      // holds (animations win over inline style), from its own time on.
      try {
        const timing = timingOf(behaviour);
        const animation = element.animate(wrapKeyframes(behaviour, keyframes.frames), {
          ...timing,
          ...(lead > 0 ? { delay: (typeof timing.delay === "number" ? timing.delay : 0) + lead } : {}),
          composite: keyframes.composite ?? "replace",
        });
        if (behaviour.wrap?.untilNext) endsAtNextClick.add(animation);
        animations.push(animation);
      } catch {
        // A keyframe this browser cannot animate: the rest still play.
      }
    }
  }
  return animations;
}
