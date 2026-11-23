/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Animation playback on the editor canvas and in the slideshow: per-block and
 * per-paragraph play state, the effect recipes and colour layers, the fly
 * offsets, and the keyframe bodies and CSS of every timeline effect.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import {
  patternRevealKeyframeBody,
  patternRevealKind,
} from "../presentation-anim-pattern-reveal";
import { createContext, type CSSProperties } from "react";
import { normalizeParagraphEndLatinFace, paragraphCarriedFields, paragraphsFromText } from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import type { EffectRecipeBox } from "@mrnote/core/web/lib/presentation-effect-recipes.ts";
import { blockAnimationKeyframeName, isBlockAnimation } from "@mrnote/core/web/lib/presentation-block-fx.ts";
import {
  type SlideParagraphWithRuns,
  withFlatRunsSplitAcrossParagraphs,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import {
  PRESENTATION_WEB_ANIMATION_EFFECTS,
  DEFAULT_STEP_DURATION_MS,
  isPresentationAnimationEffectId,
  presentationAnimationEffect,
  resolvePresentationAnimationEffectOption,
  resolveStepPlayStates,
  type PresentationAnimationEffectId,
  type PresentationHandDrawnAnimationEffectId,
  type PresentationAnimationEffectOption,
  type TimelineAnimation,
  type TimelineStep,
} from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import type {
  PresentationBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { LetterPlay } from "../../components/notebook/contents/presentation/PresentationRunsText";
import {
  effectSmoothEasing,
  type PresentationAnimationEffectSettings,
} from "@mrnote/core/web/lib/presentation-anim-effect-settings.ts";
import { effectSettingsDefaults } from "@mrnote/core/web/lib/presentation-anim-effect-defaults.ts";
import { blockGeometryStyle } from "./selection-geometry";

/**
 * The resolved present-mode animation-timeline play state of one block (a
 * collapse of every timeline STEP that references it). When absent the block is
 * not on the slide's timeline and renders exactly as before.
 */
export interface BlockPlayState {
  /** The block is referenced by at least one timeline step. */
  timelineControlled: boolean;
  /** Whether the block should be shown yet (its group has been revealed). */
  visible: boolean;
  /** Whether the block's entrance keyframes should play on this mount. */
  animate: boolean;
  /** The animation to play when `animate` (the currently-animating step's). */
  animation?: TimelineAnimation;
  /** Directional Effect Options value for the currently-animating step. */
  effectOption?: PresentationAnimationEffectOption;
  /** Fill Color target for the currently-animating emphasis step. */
  effectColor?: string;
  /** Most recent completed Fill Color target held by an earlier group. */
  settledFillColor?: string;
  /** Most recent completed Line Color target held by an earlier group. */
  settledStrokeColor?: string;
  /** Most recent completed Font Color target held by an earlier group. */
  settledTextColor?: string;
  /**
   * The end transform / opacity of completed emphasis effects (Grow / Shrink,
   * Transparency) held by earlier groups: the object stays grown or faded
   * while later groups play (PPT-B024).
   */
  settledTransform?: string;
  settledOpacity?: string;
  /** Cumulative start offset (ms) within the group → CSS animation-delay. */
  delayMs?: number;
  /** Animation duration (ms) → CSS animation-duration. */
  durationMs?: number;
  /**
   * The motion-path keyframe NAME to play (when the animating step carries a
   * motion path). The `@keyframes` is injected once per slide into a present-mode
   * `<style>` tag; the block references it via an inline `animation` shorthand so
   * it composes with (never overwrites) an entrance/emphasis/exit keyframe.
   */
  motionKeyframe?: string;
  /** Final slide-relative offset held after a completed motion-path group. */
  settledMotionTranslate?: string;
  /**
   * EVERY occurrence of this block that plays on the current mount, in sequence
   * order (a fade on click followed by a spin after-previous is two). The
   * single-valued fields above describe the LAST one; the renderer composes all
   * of them into one animation list so each plays at its own offset.
   */
  effects?: readonly BlockPlayEffect[];
  /**
   * The identity of this block's latest play: it changes whenever a NEW
   * occurrence of this block starts (a later group reveal, a fired trigger) and
   * stays the same while other objects animate. The renderer keys the block on
   * it, so a repeated identical effect restarts on the object alone, without
   * remounting the slide (which would restart its media).
   */
  playKey?: string;
  /**
   * The keyframe of the currently-animating step when its effect is drawn
   * from a recipe (built for this block's place on the slide).
   */
  recipeKeyframe?: string;
  /**
   * What held recipe emphases leave on the block (a colour shift, bold,
   * underline): kept in force while later groups play.
   */
  settledRecipeStyle?: CSSProperties;
  /** The letter-by-letter play of the block's text on this mount. */
  letterPlay?: LetterPlay;
  /** Effect Options timing of the last animating occurrence. */
  timing?: EffectCssTiming;
  /**
   * After animation: when (ms into this mount) the object is taken away --
   * Hide After Animation, or an entrance that rewinds when done playing.
   */
  afterHideAtMs?: number;
  /**
   * An imported By Paragraph build: the play state of each paragraph of this
   * block's text that a step names (zero-based). The block itself stays on
   * stage; its rows are revealed and animated one by one (PPT-B370).
   */
  paragraphPlay?: ReadonlyMap<number, ParagraphPlayRow>;
  /**
   * Picture Format > Animate as Background (PPT-A134): the picture's
   * endless zoom-and-pan loop in the show (its keyframe and length).
   */
  backgroundLoop?: { readonly keyframe: string; readonly durationMs: number };
}

/** The play state of one paragraph of a By Paragraph build. */
interface ParagraphPlayRow {
  visible: boolean;
  animate: boolean;
  animation?: TimelineAnimation;
  /** The recipe keyframe of the row's effect, when it has one. */
  recipeKeyframe?: string;
  effectOption?: PresentationAnimationEffectOption;
  delayMs?: number;
  durationMs?: number;
}

/** Rows one build step can address; a longer range plays on the first ones. */
export const MAX_PARAGRAPH_PLAY_ROWS = 500;

/**
 * Whether an object (or one paragraph of a By Paragraph build) is on stage
 * after a revealed step of its own: an entrance brings it in, an exit takes it
 * away once played (it stays while it plays), anything else leaves it as it
 * was. The one rule the show, the Morph layer and an outgoing slide read.
 */
export function timelineVisibleAfterStep(
  category: string | undefined,
  state: { visible: boolean; animate: boolean; afterHidden?: boolean; afterShown?: boolean },
  priorVisible: boolean,
): boolean {
  // Effect Options once the step has played: After animation hides the
  // object, a rewound entrance leaves it out, a rewound exit brings it back.
  if (state.afterHidden) return false;
  if (state.afterShown) return true;
  if (category === "entrance") return state.visible || priorVisible;
  if (category === "exit") return !state.visible ? false : state.animate ? true : priorVisible;
  return priorVisible;
}

/**
 * Where each object a whole-object step names starts: hidden when its first
 * such step is an entrance, on stage when it is an exit. An object no step
 * names is not in the map (on stage).
 */
export function initialTimelineVisibility(
  sequence: readonly TimelineStep[],
): Map<string, boolean> {
  const initial = new Map<string, boolean>();
  for (const step of sequence) {
    // A paragraph build reveals rows; the block itself is on stage.
    if (step.paragraphs) continue;
    if (initial.has(step.blockId)) continue;
    const category = presentationAnimationEffect(step.animation)?.category;
    if (category === "entrance") initial.set(step.blockId, false);
    else if (category === "exit") initial.set(step.blockId, true);
  }
  return initial;
}

/**
 * What of a slide is on stage once `revealedGroups` of its build have played
 * and settled, nothing animating: the objects and By Paragraph rows a build
 * has not brought in yet, or has taken away, are hidden. A slide the show is
 * leaving or arriving on during a transition is drawn this way, so a
 * transition never shows an object the slide itself does not (an exited
 * picture reappearing, an entrance's object already there).
 */
export function presentationTimelineVisibility(
  sequence: readonly TimelineStep[],
  revealedGroups: number,
  firedTriggers?: ReadonlySet<string>,
): {
  blocks: Map<string, boolean>;
  rows: Map<string, Map<number, ParagraphPlayRow>>;
} {
  const blocks = new Map<string, boolean>();
  const rows = new Map<string, Map<number, ParagraphPlayRow>>();
  if (sequence.length === 0) return { blocks, rows };
  const states = resolveStepPlayStates(sequence, revealedGroups, firedTriggers, true);
  const initial = initialTimelineVisibility(sequence);
  for (let i = 0; i < sequence.length; i += 1) {
    const step = sequence[i];
    const state = states.get(i);
    if (!state) continue;
    const category = presentationAnimationEffect(step.animation)?.category;
    const priorBlock = blocks.get(step.blockId) ?? initial.get(step.blockId) ?? true;
    if (step.paragraphs) {
      blocks.set(step.blockId, priorBlock);
      const blockRows = rows.get(step.blockId) ?? new Map<number, ParagraphPlayRow>();
      const last = Math.min(step.paragraphs.end, step.paragraphs.start + MAX_PARAGRAPH_PLAY_ROWS - 1);
      for (let row = step.paragraphs.start; row <= last; row += 1) {
        const priorRow = blockRows.get(row)?.visible ?? category !== "entrance";
        blockRows.set(row, { visible: timelineVisibleAfterStep(category, state, priorRow), animate: false });
      }
      rows.set(step.blockId, blockRows);
      continue;
    }
    blocks.set(step.blockId, timelineVisibleAfterStep(category, state, priorBlock));
  }
  return { blocks, rows };
}

interface ParagraphPlayContextValue {
  rows: ReadonlyMap<number, ParagraphPlayRow>;
  /** Fly distances from the block's place on the slide, shared by its rows. */
  flyVars: Record<string, string> | null;
}

/**
 * The paragraph play of the text/shape block being drawn in the show; `null`
 * everywhere else (the canvas, thumbnails, a block with no paragraph build).
 */
export const PresentationParagraphPlayContext =
  createContext<ParagraphPlayContextValue | null>(null);

/**
 * What a paragraph row adds for its build step: nothing when no step names it,
 * hidden-in-layout until its click, and its effect on the mount it plays.
 */
export function paragraphPlayRowProps(
  value: ParagraphPlayContextValue | null,
  index: number,
): { className?: string; style?: CSSProperties; state?: "pending" | "playing" | "shown" } {
  const row = value?.rows.get(index);
  if (!value || !row) return {};
  if (!row.visible && !row.animate) {
    return { state: "pending", style: { visibility: "hidden" } };
  }
  if (!row.animate) return { state: "shown" };
  const effectName = playKeyframeName(row.animation, row.effectOption, row.recipeKeyframe);
  if (!effectName) return { state: "shown" };
  // An entrance ends on the row's own resting style, so it only has to hold its
  // first frame through the delay (`backwards`); holding the last frame too
  // would keep `transform: none` over the row's baseline offset. An exit or an
  // emphasis holds where it ends.
  const entrance =
    presentationAnimationEffect(row.animation)?.category === "entrance";
  return {
    state: "playing",
    // The reduced-motion reset for inline timeline animations.
    className: "presentation-present__block--motion",
    style: {
      ...(value.flyVars as CSSProperties | null),
      animationName: effectName,
      animationTimingFunction: playTimingFunction(row.recipeKeyframe),
      animationFillMode: entrance ? "backwards" : "both",
      animationDelay: `${row.delayMs ?? 0}ms`,
      animationDuration: `${row.durationMs ?? DEFAULT_STEP_DURATION_MS}ms`,
    },
  };
}

/**
 * The rows a By Paragraph build plays on when the body is stored as flat text
 * (rows that state nothing of their own are not persisted as `paragraphs`):
 * one row per line, each with its share of the block's runs.
 */
export function paragraphBuildRows(
  text: string,
  runs: unknown,
  /** The flat body's mark Latin face (`endLatinFace`, readings row 899): its last row's. */
  endLatinFace?: unknown,
  /**
   * The flat body's own kerning and East Asian line rules (its
   * `nativeKerning` / `nativeEastAsianLines`, readings row 1009): every row's,
   * as the one paragraph's rows -- and one Enter adds -- all take its `a:pPr`.
   */
  lineReceipts?: unknown,
): SlideParagraphWithRuns[] {
  const split = withFlatRunsSplitAcrossParagraphs(
    paragraphsFromText(text) as SlideParagraphWithRuns[],
    runs,
    text,
  );
  const carried = lineReceipts && typeof lineReceipts === "object"
    ? (({ nativeKerning, nativeEastAsianLines }) => ({ nativeKerning, nativeEastAsianLines }))(paragraphCarriedFields(lineReceipts))
    : {};
  const ownReceipts = Object.fromEntries(Object.entries(carried).filter(([, value]) => value !== undefined));
  const rows = Object.keys(ownReceipts).length ? split.map((row) => ({ ...ownReceipts, ...row })) : split;
  const face = normalizeParagraphEndLatinFace(endLatinFace);
  if (face === undefined || !rows.length || rows[rows.length - 1]!.endLatinFace !== undefined) return rows;
  // A body saying its mark's face is one paragraph (readings row 899): its rows
  // before the last end at line breaks (`a:br`), which hold no mark.
  return [
    ...rows.slice(0, -1).map((row) => ({ ...row, endsInLineBreak: true as const })),
    { ...rows[rows.length - 1]!, endLatinFace: face },
  ];
}

/**
 * The fly distances of a block's paragraphs: from the block's place on the
 * slide, in slide units, so every row starts past the same edge the block
 * would (a row is not as tall as its block, so the block's own percentages
 * would stop it short).
 */
export function presentationParagraphFlyOffsetVars(
  geometry: CSSProperties | undefined,
): Record<string, string> | null {
  const percent = (value: unknown) =>
    typeof value === "string" && value.endsWith("%") ? Number.parseFloat(value) : Number.NaN;
  const left = percent(geometry?.left);
  const top = percent(geometry?.top);
  const width = percent(geometry?.width);
  const height = percent(geometry?.height);
  if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    return null;
  }
  const round = (value: number) => Math.round(value * 1000) / 1000;
  return {
    "--mrai-ppt-fly-left": `${round(-(left + width))}cqw`,
    "--mrai-ppt-fly-right": `${round(100 - left)}cqw`,
    "--mrai-ppt-fly-top": `${round(-(top + height))}cqh`,
    "--mrai-ppt-fly-bottom": `${round(100 - top)}cqh`,
  };
}

/** One animating occurrence of a block on the current mount. */
/**
 * How an occurrence's keyframe runs under its Effect Options: Smooth start /
 * end as an easing, Repeat and Auto-reverse as iterations and direction,
 * Rewind when done playing as a fill that does not hold the end.
 */
export interface EffectCssTiming {
  /** Replaces the effect's own easing (Smooth start / end were changed). */
  timingFunction?: string;
  iterationCount: string;
  direction: "normal" | "alternate";
  /** Rewind when done playing: the end frame is not held. */
  rewind: boolean;
}

/** The CSS timing of a step's Effect Options, or undefined for the defaults. */
export function effectCssTiming(step: {
  animation: TimelineAnimation;
  effectOption?: PresentationAnimationEffectOption;
  effectSettings?: PresentationAnimationEffectSettings;
}): EffectCssTiming | undefined {
  const settings = step.effectSettings;
  if (!settings) return undefined;
  const { accel, decel, repeat, autoReverse, rewind } = settings;
  if (accel === undefined && decel === undefined && repeat === undefined && !autoReverse && !rewind) {
    return undefined;
  }
  const defaults = effectSettingsDefaults(step.animation, step.effectOption);
  const smoothChanged = accel !== undefined || decel !== undefined;
  const easing = smoothChanged
    ? (effectSmoothEasing(accel ?? defaults.accel, decel ?? defaults.decel) ?? "linear")
    : undefined;
  const plays =
    typeof repeat === "string" ? Infinity : (repeat ?? 1000) / 1000;
  const iterations = plays * (autoReverse ? 2 : 1);
  return {
    ...(easing ? { timingFunction: easing } : {}),
    iterationCount: Number.isFinite(iterations) ? String(Math.round(iterations * 1000) / 1000) : "infinite",
    direction: autoReverse ? "alternate" : "normal",
    rewind: Boolean(rewind),
  };
}

/** The keyframe that takes an object away when its effect ends (After
 * animation: Hide After Animation; an entrance that rewinds). */
export const PRESENTATION_AFTER_HIDE_KEYFRAME = "mrai-ppt-after-hide";
export const PRESENTATION_AFTER_HIDE_KEYFRAME_CSS = `@keyframes ${PRESENTATION_AFTER_HIDE_KEYFRAME}{from,to{visibility:hidden}}`;

/**
 * The colour a Dim names, on this slide's theme: a hex as it is, a theme
 * colour through the default colour map (bg1 = lt1, tx1 = dk1, bg2 = lt2,
 * tx2 = dk2), Office grey when the theme does not state it.
 */
export function resolveEffectDimColor(
  dimColor: string | undefined,
  slots: Partial<Record<string, string>> | undefined,
  accents: readonly string[],
): string | undefined {
  if (!dimColor) return undefined;
  if (dimColor.startsWith("#")) return dimColor;
  const slot =
    ({ bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" } as Record<string, string>)[dimColor] ?? dimColor;
  const accent = /^accent([1-6])$/.exec(slot);
  return slots?.[slot] ?? (accent ? accents[Number(accent[1]) - 1] : undefined) ?? "#7f7f7f";
}

interface BlockPlayEffect {
  animation: TimelineAnimation;
  effectOption?: PresentationAnimationEffectOption;
  delayMs: number;
  durationMs: number;
  /** Effect Options timing of this occurrence (Repeat, Smooth start …). */
  timing?: EffectCssTiming;
  /** The occurrence plays on the block's words / letters, not the block. */
  letters?: true;
  motionKeyframe?: string;
  /** The occurrence's recipe keyframe, when its effect is drawn from one. */
  recipeKeyframe?: string;
  /** A Fill / Line / Font Color occurrence's target colour. */
  effectColor?: string;
  /**
   * The colour a recipe effect changes the object to (Object Color, Color
   * Pulse, Blast): which of fill / line / text take it, and the keyframe of
   * the colour layer's opacity.
   */
  colorLayer?: RecipeColorLayer;
}

export interface RecipeColorLayer {
  color: string;
  fill: boolean;
  line: boolean;
  text: boolean;
  keyframe: string;
}

/** Where a block sits on the slide, for a recipe (centre and size as slide fractions). */
export function effectRecipeBoxOf(
  block: PresentationBlock | undefined,
  bounds: { width: number; height: number },
): EffectRecipeBox {
  const style = block ? blockGeometryStyle(block, bounds) : undefined;
  const percent = (value: unknown) =>
    typeof value === "string" && value.endsWith("%") ? Number.parseFloat(value) / 100 : Number.NaN;
  const left = percent(style?.left);
  const top = percent(style?.top);
  const width = percent(style?.width);
  const height = percent(style?.height);
  if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    // A block the slide lays out itself: the middle of the slide is as good a
    // place as any for the distances the recipe needs.
    return { x: 0.5, y: 0.5, w: 0.5, h: 0.2 };
  }
  return { x: left + width / 2, y: top + height / 2, w: width, h: height };
}

/** The keyframe a play occurrence runs: its recipe's, else the effect's own. */
export function playKeyframeName(
  animation: TimelineAnimation | undefined,
  effectOption: PresentationAnimationEffectOption | undefined,
  recipeKeyframe: string | undefined,
): string | null {
  return recipeKeyframe ?? timelineAnimationKeyframeName(animation, effectOption);
}

/** A recipe bakes its own acceleration into its samples. */
export function playTimingFunction(recipeKeyframe: string | undefined): "linear" | "ease-out" {
  return recipeKeyframe ? "linear" : "ease-out";
}

/**
 * The colour change of kind `animation` playing on this mount, with its own
 * start and length. It is looked up among EVERY occurrence of the block, not
 * only the last one: a Fill Color followed by a Spin "After Previous" in one
 * click plays both, and the colour layer used to be drawn only when the colour
 * was the last effect (PPT-B049), so it appeared a click late.
 */
export function activeBlockColorEffect(
  play: BlockPlayState | undefined,
  animation: "fill-color" | "line-color" | "font-color",
): { color: string; delayMs: number; durationMs: number; keyframe?: string } | undefined {
  if (!play?.animate) return undefined;
  const effects = play.effects ?? [];
  const channel =
    animation === "fill-color" ? "fill" : animation === "line-color" ? "line" : "text";
  for (let index = effects.length - 1; index >= 0; index -= 1) {
    const effect = effects[index];
    // A recipe effect that changes the object to a colour draws the same
    // layers, with its own opacity keyframe.
    if (effect.colorLayer?.[channel]) {
      return {
        color: effect.colorLayer.color,
        delayMs: effect.delayMs,
        durationMs: effect.durationMs,
        keyframe: effect.colorLayer.keyframe,
      };
    }
    if (effect.animation === animation && effect.effectColor) {
      return {
        color: effect.effectColor,
        delayMs: effect.delayMs,
        durationMs: effect.durationMs,
      };
    }
  }
  return play.animation === animation && play.effectColor
    ? {
        color: play.effectColor,
        delayMs: play.delayMs ?? 0,
        durationMs: play.durationMs ?? DEFAULT_STEP_DURATION_MS,
      }
    : undefined;
}

/**
 * Fly In starts, and Fly Out ends, with the object just past the slide's edge
 * on the chosen side: PowerPoint writes `#ppt_x` from `0-#ppt_w/2` (From
 * Left), `1+#ppt_w/2` (From Right), and the same on `#ppt_y`. The distance
 * depends on where the object sits, so each object publishes it (see
 * {@link presentationFlyOffsetVars}) in percent of its own box, which is what
 * `translate` percentages measure. Without it, one slide's size is enough.
 */
const PRESENTATION_FLY_OFFSETS = {
  "from-bottom": ["translateY", "--mrai-ppt-fly-bottom", "100cqh"],
  "from-left": ["translateX", "--mrai-ppt-fly-left", "-100cqw"],
  "from-top": ["translateY", "--mrai-ppt-fly-top", "-100cqh"],
  "from-right": ["translateX", "--mrai-ppt-fly-right", "100cqw"],
} as const;

/** A corner is its two sides together. */
const PRESENTATION_FLY_CORNERS: Readonly<
  Record<string, readonly (keyof typeof PRESENTATION_FLY_OFFSETS)[]>
> = {
  "from-bottom-left": ["from-left", "from-bottom"],
  "from-bottom-right": ["from-right", "from-bottom"],
  "from-top-left": ["from-left", "from-top"],
  "from-top-right": ["from-right", "from-top"],
};

function presentationFlyOffsetTransform(
  option: PresentationAnimationEffectOption,
): string | null {
  const sides = PRESENTATION_FLY_CORNERS[option] ?? [
    option as keyof typeof PRESENTATION_FLY_OFFSETS,
  ];
  const parts = sides.map((side) => {
    const offset = PRESENTATION_FLY_OFFSETS[side];
    return offset ? `${offset[0]}(var(${offset[1]},${offset[2]}))` : null;
  });
  return parts.every((part): part is string => part !== null) ? parts.join(" ") : null;
}

/** The fly distances for an object at `geometry` (percent of the slide). */
export function presentationFlyOffsetVars(
  geometry: CSSProperties | undefined,
): Record<string, string> | null {
  const percent = (value: unknown) =>
    typeof value === "string" && value.endsWith("%") ? Number.parseFloat(value) : Number.NaN;
  const left = percent(geometry?.left);
  const top = percent(geometry?.top);
  const width = percent(geometry?.width);
  const height = percent(geometry?.height);
  if (![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
    return null;
  }
  const own = (distance: number, size: number) =>
    `${Math.round((distance / size) * 100 * 1000) / 1000}%`;
  return {
    "--mrai-ppt-fly-left": own(-(left + width), width),
    "--mrai-ppt-fly-right": own(100 - left, width),
    "--mrai-ppt-fly-top": own(-(top + height), height),
    "--mrai-ppt-fly-bottom": own(100 - top, height),
  };
}

const PRESENTATION_TIMELINE_EFFECT_KEYFRAME_BODIES = Object.freeze({
  appear: "0%{opacity:0}0.1%,100%{opacity:1}",
  fade: "0%{opacity:0}100%{opacity:1}",
  "fly-in": `0%{transform:${presentationFlyOffsetTransform("from-bottom")}}100%{transform:none}`,
  "split-in": "0%{clip-path:inset(50% 0)}100%{clip-path:inset(0)}",
  "wipe-in": "0%{clip-path:inset(0 100% 0 0)}100%{clip-path:inset(0)}",
  "random-bars-in": patternRevealKeyframeBody("random-bars", "horizontal", true),
  "grow-and-turn": "0%{opacity:0;transform:scale(0) rotate(90deg)}100%{opacity:1;transform:none}",
  zoom: "0%{opacity:0;transform:scale(0)}100%{opacity:1;transform:none}",
  "basic-zoom-in": "0%{transform:scale(0)}100%{transform:none}",
  "blinds-in": patternRevealKeyframeBody("blinds", "horizontal", true),
  "box-in": "0%{clip-path:inset(50%)}100%{clip-path:inset(0)}",
  "checkerboard-in": patternRevealKeyframeBody("checkerboard", "horizontal", true),
  expand: "0%{opacity:0;transform:scaleX(.7)}100%{opacity:1;transform:none}",
  "peek-in": presentationPeekKeyframeBody("from-bottom", true)!,
  "rise-up": "0%{opacity:0;transform:translateY(100vh)}90%{opacity:1;transform:translateY(-3vh)}100%{opacity:1;transform:none}",
  "stretch-in": "0%{transform:scaleX(0)}100%{transform:none}",
  "strips-in": "0%{clip-path:polygon(0 0,0 0,0 100%,0 100%)}100%{clip-path:polygon(0 0,100% 0,100% 100%,0 100%)}",
  spin: "0%{transform:rotate(0deg)}100%{transform:rotate(calc(360deg * var(--mrai-ppt-turn,1)))}",
  "grow-shrink": "0%{transform:none}100%{transform:scale(1.5)}",
  transparency: "0%{opacity:1}100%{opacity:.5}",
  "line-color": "0%,100%{filter:none}",
  "fill-color": "0%,100%{filter:none}",
  "font-color": "0%,100%{filter:none}",
  disappear: "0%,99.9%{opacity:1}100%{opacity:0}",
  "fade-out": "0%{opacity:1}100%{opacity:0}",
  "fly-out": `0%{transform:none}100%{transform:${presentationFlyOffsetTransform("from-bottom")}}`,
  "split-out": "0%{clip-path:inset(0)}100%{clip-path:inset(50% 0)}",
  "wipe-out": "0%{clip-path:inset(0)}100%{clip-path:inset(0 0 0 100%)}",
  "random-bars-out": patternRevealKeyframeBody("random-bars", "horizontal", false),
  "zoom-out": "0%{opacity:1;transform:none}100%{opacity:0;transform:scale(0)}",
  "basic-zoom-out": "0%{transform:none}100%{transform:scale(0)}",
  "blinds-out": patternRevealKeyframeBody("blinds", "horizontal", false),
  "box-out": "0%{clip-path:inset(0)}100%{clip-path:inset(50%)}",
  "checkerboard-out": patternRevealKeyframeBody("checkerboard", "horizontal", false),
  "peek-out": presentationPeekKeyframeBody("from-bottom", false)!,
  "stretchy-out": "0%{opacity:1;transform:none}100%{opacity:0;transform:scaleX(1.3)}",
  "stretch-out": "0%{transform:none}100%{transform:scaleX(0)}",
  "strips-out": "0%{clip-path:polygon(0 0,100% 0,100% 100%,0 100%)}100%{clip-path:polygon(100% 0,100% 0,100% 100%,100% 100%)}",
  // A media call plays or pauses its clip; nothing about the clip moves.
  "media-play": "0%,100%{}",
  "media-pause": "0%,100%{}",
  // Blink: hidden for the first half, shown for the second.
  blink: "0%,49.99%{visibility:hidden}50%,100%{visibility:visible}",
} satisfies Record<PresentationHandDrawnAnimationEffectId, string>);

/** The hand-written keyframe of an effect; `undefined` for one played from a recipe. */
function handDrawnKeyframeBody(effect: string): string | undefined {
  return Object.prototype.hasOwnProperty.call(
    PRESENTATION_TIMELINE_EFFECT_KEYFRAME_BODIES,
    effect,
  )
    ? PRESENTATION_TIMELINE_EFFECT_KEYFRAME_BODIES[
        effect as PresentationHandDrawnAnimationEffectId
      ]
    : undefined;
}

/**
 * Peek In / Out, as PowerPoint writes it: the object travels 1.125 of its own
 * size along one axis (`#ppt_y+#ppt_h*1.125000` for From Bottom) while a wipe
 * in the same direction uncovers it, so it rises from behind the edge of the
 * box it ends in and is never seen outside that box. The clip is in the
 * object's own coordinates, so as the object is moved by 112.5% of a side the
 * same share is cut off its leading side, which keeps what shows inside the
 * final box at every instant. No fade. It used to slide 35% and fade in with
 * no clip, showing the whole object outside its place (PPT-B182). An exit
 * runs the journey backwards.
 */
function presentationPeekKeyframeBody(
  option: PresentationAnimationEffectOption,
  entering: boolean,
): string | null {
  const away: Partial<
    Record<PresentationAnimationEffectOption, readonly [string, string]>
  > = {
    "from-bottom": ["translateY(112.5%)", "inset(0 0 112.5% 0)"],
    "from-left": ["translateX(-112.5%)", "inset(0 0 0 112.5%)"],
    "from-top": ["translateY(-112.5%)", "inset(112.5% 0 0 0)"],
    "from-right": ["translateX(112.5%)", "inset(0 112.5% 0 0)"],
  };
  const hidden = away[option];
  if (!hidden) return null;
  const outside = `transform:${hidden[0]};clip-path:${hidden[1]}`;
  const inPlace = "transform:none;clip-path:inset(0)";
  return entering
    ? `0%{${outside}}100%{${inPlace}}`
    : `0%{${inPlace}}100%{${outside}}`;
}

/**
 * The object with a rectangular hole cut out of it, in percent of its box:
 * one even-odd polygon, so a keyframe can move the hole's edges and the
 * browser interpolates the clip vertex by vertex.
 */
function clipWithCentredHole(
  left: number,
  top: number,
  right: number,
  bottom: number,
): string {
  const [l, t, r, b] = [left, top, right, bottom].map((edge) => `${edge}%`);
  return `polygon(evenodd,0% 0%,100% 0%,100% 100%,0% 100%,0% 0%,${l} ${t},${l} ${b},${r} ${b},${r} ${t},${l} ${t})`;
}

/**
 * Split In's two doors: the object minus a centred band across it (Horizontal)
 * or down it (Vertical), `edge` percent in from each side. At 0 the band is
 * the whole object and nothing shows; at 50 it has closed to a line.
 */
function splitDoorsClip(axis: "horizontal" | "vertical", edge: number): string {
  return axis === "horizontal"
    ? clipWithCentredHole(0, edge, 100, 100 - edge)
    : clipWithCentredHole(edge, 0, 100 - edge, 100);
}

export function animationEffectOptionKeyframeBody(
  animation: PresentationAnimationEffectId,
  option: PresentationAnimationEffectOption,
): string | null {
  if (animation === "transparency") {
    const visibleOpacity = resolvePresentationAnimationEffectOption(animation, option)?.nativeOpacity;
    return visibleOpacity === undefined || visibleOpacity === null
      ? null
      : `0%{opacity:1}100%{opacity:${visibleOpacity}}`;
  }
  const entering = presentationAnimationEffect(animation)?.category === "entrance";
  if (animation.startsWith("wipe-")) {
    const hiddenByOption: Partial<
      Record<PresentationAnimationEffectOption, string>
    > = {
      "from-bottom": "inset(100% 0 0)",
      "from-left": "inset(0 100% 0 0)",
      "from-top": "inset(0 0 100%)",
      "from-right": "inset(0 0 0 100%)",
    };
    const hidden = hiddenByOption[option];
    return hidden
      ? entering
        ? `0%{clip-path:${hidden}}100%{clip-path:inset(0)}`
        : `0%{clip-path:inset(0)}100%{clip-path:${hidden}}`
      : null;
  }
  if (animation.startsWith("peek-")) {
    return presentationPeekKeyframeBody(option, entering);
  }
  if (animation.startsWith("fly-")) {
    const offset = presentationFlyOffsetTransform(option);
    // Fly In / Out travels to or from just past the slide's edge and does not
    // fade (PPT-B139).
    return offset
      ? entering
        ? `0%{transform:${offset}}100%{transform:none}`
        : `0%{transform:none}100%{transform:${offset}}`
      : null;
  }
  if (animation === "stretch-in" || animation === "stretch-out") {
    const stretchByOption: Partial<
      Record<PresentationAnimationEffectOption, readonly [string, string]>
    > = {
      across: ["center", "scaleX(0)"],
      "from-bottom": ["center bottom", "scaleY(0)"],
      "from-left": ["left center", "scaleX(0)"],
      "from-right": ["right center", "scaleX(0)"],
      "from-top": ["center top", "scaleY(0)"],
    };
    const stretch = stretchByOption[option];
    if (!stretch) return null;
    // The exit is the same journey run backwards, as PowerPoint writes it.
    return entering
      ? `0%{transform-origin:${stretch[0]};transform:${stretch[1]}}100%{transform-origin:${stretch[0]};transform:none}`
      : `0%{transform-origin:${stretch[0]};transform:none}100%{transform-origin:${stretch[0]};transform:${stretch[1]}}`;
  }
  if (animation.startsWith("split-")) {
    // Split is DrawingML's barn door (`barn(outVertical)` and so on, the
    // SMIL barnDoorWipe): a Vertical door is a band as tall as the object
    // centred on its vertical axis, a Horizontal one the same turned a
    // quarter. Out opens the band from the centre line; In closes the object
    // in from its two edges, a centred hole narrowing to nothing. Horizontal
    // In and Vertical Out used to play the same horizontal band, and the
    // two Outs had their axes swapped (PPT-B181). An exit runs the same
    // journey backwards, as Box does.
    const axis = option.startsWith("horizontal-")
      ? "horizontal"
      : option.startsWith("vertical-")
        ? "vertical"
        : null;
    if (!axis || (!option.endsWith("-in") && !option.endsWith("-out"))) {
      return null;
    }
    const [hidden, shown] = option.endsWith("-out")
      ? [axis === "horizontal" ? "inset(50% 0)" : "inset(0 50%)", "inset(0)"]
      : [splitDoorsClip(axis, 0), splitDoorsClip(axis, 50)];
    return entering
      ? `0%{clip-path:${hidden}}100%{clip-path:${shown}}`
      : `0%{clip-path:${shown}}100%{clip-path:${hidden}}`;
  }
  if (animation.startsWith("box-")) {
    if (option !== "in" && option !== "out") return null;
    // Out opens a box from the centre; In closes one from the edges, so the
    // object shows as a frame that thickens inward: the whole object minus a
    // centred hole that shrinks to a point. `inset(0)` as In's start was the
    // object already whole, which left only a fade (PPT-B160). An exit runs
    // the same journey backwards; neither fades, as PowerPoint draws them.
    const frame = (hole: number) =>
      clipWithCentredHole(hole, hole, 100 - hole, 100 - hole);
    const [hidden, shown] =
      option === "out" ? ["inset(50%)", "inset(0)"] : [frame(0), frame(50)];
    return entering
      ? `0%{clip-path:${hidden}}100%{clip-path:${shown}}`
      : `0%{clip-path:${shown}}100%{clip-path:${hidden}}`;
  }
  const pattern = patternRevealKind(animation);
  if (pattern) {
    // Slats, squares or scattered bars, each drawn as its own pieces
    // (PPT-B138); no fade, as PowerPoint draws them.
    if (option !== "horizontal" && option !== "vertical") return null;
    return patternRevealKeyframeBody(pattern, option, entering);
  }
  if (animation.startsWith("strips-")) {
    const hiddenByOption: Partial<
      Record<PresentationAnimationEffectOption, string>
    > = {
      "down-left": "polygon(100% 0,100% 0,0 100%,0 100%)",
      "up-left": "polygon(100% 100%,100% 100%,0 0,0 0)",
      "up-right": "polygon(0 100%,0 100%,100% 0,100% 0)",
      "down-right": "polygon(0 0,0 0,100% 100%,100% 100%)",
    };
    const hidden = hiddenByOption[option];
    const full = "polygon(0 0,100% 0,100% 100%,0 100%)";
    return hidden
      ? entering
        ? `0%{clip-path:${hidden}}100%{clip-path:${full}}`
        : `0%{clip-path:${full}}100%{clip-path:${hidden}}`
      : null;
  }
  return null;
}

/**
 * The object's own rotate / flip, exposed to its effect keyframes. An effect
 * animates `transform` on the same node that carries the object's geometry
 * transform, so a keyframe that wrote `transform: none` (or a bare
 * `scale()`) erased a 90-degree rotation for the whole effect and after it
 * (PPT-B023). Every effect transform is composed with this variable instead:
 * a slide-space translation in front of it, an object-space scale / turn
 * after it, `none` as the variable itself.
 */
export const PRESENTATION_BLOCK_BASE_TRANSFORM_VAR = "--mrai-ppt-base-transform";

function withBlockBaseTransform(css: string): string {
  return css.replace(/transform:([^;}]+)/g, (_match, raw: string) => {
    const value = raw.trim();
    if (value === "none") {
      return `transform:var(${PRESENTATION_BLOCK_BASE_TRANSFORM_VAR},none)`;
    }
    return value.startsWith("translate")
      ? `transform:${value} var(${PRESENTATION_BLOCK_BASE_TRANSFORM_VAR},)`
      : `transform:var(${PRESENTATION_BLOCK_BASE_TRANSFORM_VAR},) ${value}`;
  });
}

export const PRESENTATION_TIMELINE_EFFECT_KEYFRAMES_CSS = withBlockBaseTransform(
  // A recipe-drawn effect has no static keyframe: its keyframe is built for
  // the object's place on the slide when it plays.
  PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (effect) => effect.nativeBehavior !== "recipe",
  ).flatMap((effect) => {
    const defaultBody = effect.defaultOption
      ? animationEffectOptionKeyframeBody(effect.id, effect.defaultOption)
      : null;
    return [
      `@keyframes mrai-ppt-effect-${effect.id}{${defaultBody ?? handDrawnKeyframeBody(effect.id) ?? ""}}`,
      ...effect.options
        .filter((option) => option !== effect.defaultOption)
        .map((option) => {
          const body = animationEffectOptionKeyframeBody(effect.id, option);
          return body
            ? `@keyframes mrai-ppt-effect-${effect.id}-${option}{${body}}`
            : "";
        }),
    ];
  })
    .filter(Boolean)
    .join("\n") +
  "\n@keyframes mrai-ppt-effect-fill-color-overlay{0%{opacity:0}100%{opacity:1}}" +
  "\n@keyframes mrai-ppt-text-fill-color{0%{background-color:var(--mrai-ppt-color-from)}100%{background-color:var(--mrai-ppt-color-to)}}" +
  "\n@keyframes mrai-ppt-text-line-color{0%{border-color:var(--mrai-ppt-color-from)}100%{border-color:var(--mrai-ppt-color-to)}}",
);

/**
 * The end state an emphasis effect leaves on its object: its keyframes' final
 * `transform` / `opacity` (the option's variant when it has one). Null when it
 * leaves neither.
 */
export function settledEmphasisFrame(
  animation: TimelineAnimation,
  effectOption?: PresentationAnimationEffectOption,
): { transform?: string; opacity?: string } | null {
  const effect = presentationAnimationEffect(animation);
  if (!effect) return null;
  const option = effectOption ?? effect.defaultOption;
  const body =
    (option ? animationEffectOptionKeyframeBody(effect.id, option) : null) ??
    handDrawnKeyframeBody(effect.id);
  if (!body) return null;
  const last = /(?:^|\})[^{}]*100%\{([^}]*)\}\s*$/.exec(body)?.[1];
  if (!last) return null;
  const frame: { transform?: string; opacity?: string } = {};
  for (const declaration of last.split(";")) {
    const [key, ...rest] = declaration.split(":");
    const value = rest.join(":").trim();
    if (key.trim() === "transform" && value && value !== "none") frame.transform = value;
    if (key.trim() === "opacity" && value && value !== "1") frame.opacity = value;
  }
  return frame.transform || frame.opacity ? frame : null;
}

export function timelineAnimationKeyframeName(
  animation: TimelineAnimation | undefined,
  effectOption?: PresentationAnimationEffectOption,
): string | null {
  const effect = presentationAnimationEffect(animation);
  if (effect?.defaultOption || effect?.nativeBehavior === "zoom" ||
      effect?.id === "spin" || effect?.id === "grow-shrink") {
    return `mrai-ppt-effect-${effect.id}${effectOption ? `-${effectOption.replace(".", "_")}` : ""}`;
  }
  return isBlockAnimation(animation)
    ? blockAnimationKeyframeName(animation)
    : isPresentationAnimationEffectId(animation)
      ? `mrai-ppt-effect-${animation}${effectOption ? `-${effectOption}` : ""}`
      : null;
}
