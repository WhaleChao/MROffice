/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-slide present-mode ANIMATION TIMELINE (PowerPoint animation-pane parity).
 *
 * Per-block entrance animations (see `presentation-block-fx.ts`) answer "what
 * does this block do when its slide appears". The timeline answers the richer
 * question PowerPoint's animation pane does: an ORDERED sequence of steps, each
 * naming a block + animation + a TRIGGER (start `on click`, `with previous`, or
 * `after previous`) + a delay + a duration. In present mode the steps play in
 * order: each `on-click` step opens a new group the audience advances to, and
 * `with-previous` / `after-previous` steps chain off the group's start via
 * cumulative offsets.
 *
 * This module is the single source of truth for the allowed trigger set, the
 * clamps, and the grouping + timing maths. Keep it dependency-free (no React /
 * no DOM / no zod) so it can be unit tested in isolation and imported from both
 * the zod schema and the present-mode renderer. It mirrors the conventions of
 * the sibling presentation libs: frozen allowed lists, a lenient normalizer that
 * collapses stale / hand-edited values to a safe default, and pure resolvers.
 */

import {
  BLOCK_ANIMATIONS,
  DEFAULT_BLOCK_ANIMATION,
} from "./presentation-block-fx";
import {
  PRESENTATION_EFFECT_GALLERY_GROUPS,
  PRESENTATION_RECIPE_EFFECT_EXTENSIONS,
  PRESENTATION_RECIPE_EFFECT_IDENTITIES,
} from "./presentation-effect-recipe-identities";
import { RUN_FONT_FAMILIES, runFontFamilyName } from "./presentation-runs";
import {
  effectSettingsEqual,
  normalizeEffectSettings,
  type PresentationAnimationEffectSettings,
} from "./presentation-anim-effect-settings";
import {
  normalizeMotionPath,
  presetMotionPath,
  type MotionPath,
  type MotionPathPreset,
} from "./presentation-motion-path";

/**
 * How a timeline step STARTS relative to the rest of the sequence:
 *  - `on-click`: waits for the presenter to advance; begins a NEW play group.
 *  - `with-previous`: starts at the SAME time as the previous step in its group.
 *  - `after-previous`: starts when the previous step in its group has finished.
 *
 * `on-click` is the implicit default — a fresh step is its own click group, the
 * least surprising behaviour and what PowerPoint defaults a first effect to.
 */
export const TIMELINE_TRIGGERS = [
  "on-click",
  "with-previous",
  "after-previous",
] as const;

export type TimelineTrigger = (typeof TIMELINE_TRIGGERS)[number];

/** The trigger applied when a step has no explicit (or invalid) value. */
export const DEFAULT_TIMELINE_TRIGGER: TimelineTrigger = "on-click";

/**
 * MRAI exposes the bounded desktop-compatible "On Click of" object trigger.
 * Media-bookmark triggers remain outside this authority until bookmark XML is
 * modeled. The persisted grammar and present-mode playback share this flag.
 */
export const PRESENTATION_INTERACTIVE_TRIGGER_AUTHORING_SUPPORTED = true;

const TIMELINE_TRIGGER_SET: ReadonlySet<string> = new Set(TIMELINE_TRIGGERS);

/**
 * Timing bounds (milliseconds). Delay and duration are clamped into these so a
 * hand-edited or stale deck can never schedule a negative / NaN / absurdly long
 * animation that would freeze present mode. `DEFAULT_STEP_DURATION_MS` mirrors
 * the 380ms CSS keyframe duration in `styles/native-page-window.css` so a step
 * with no explicit duration matches the per-block entrance animation feel.
 */
export const MIN_STEP_DELAY_MS = 0;
export const MAX_STEP_DELAY_MS = 60_000;
export const MIN_STEP_DURATION_MS = 0;
export const MAX_STEP_DURATION_MS = 60_000;
export const DEFAULT_STEP_DELAY_MS = 0;
export const DEFAULT_STEP_DURATION_MS = 380;
export const ANIMATION_TIMING_STEP_MS = 50;

/**
 * Hard cap on the number of steps in a single slide's sequence. Bounds the
 * persisted payload and the present-mode work; extra steps beyond the cap are
 * dropped by {@link normalizeAnimationSequence}. Generous — far more than any
 * real slide needs.
 */
export const MAX_TIMELINE_STEPS = 200;

export const PRESENTATION_ANIMATION_EFFECT_CATEGORIES = [
  "entrance",
  "emphasis",
  "exit",
] as const;

export type PresentationAnimationEffectCategory =
  (typeof PRESENTATION_ANIMATION_EFFECT_CATEGORIES)[number];

/**
 * PowerPoint for the web's current Add Animation gallery.
 *
 * IDs are class-qualified whenever Office exposes the same display name in
 * more than one category (`split-in` / `split-out`, for example). That keeps a
 * saved step and its native presetClass unambiguous without a second handwritten
 * category table. Each `authoring` receipt is one standards-valid default
 * variant shared by Web, API, scratch export and patchback. Entries without one
 * remain visible/playable but preserve-only.
 */
// PowerPoint pairs every single-object build with a shape build entry: the
// effect p:cTn carries a grpId and p:bldLst carries the matching p:bldP. The
// API's catalogue does the same for these three families; leaving them off
// here made a deck dragged into the editor lose every Wipe, Split, Blinds,
// Box, Checkerboard, Peek, Strips, Random Bars, Spin and Grow/Shrink.
const nativeEffect = (subtype: number, filter: string | null) => ({
  subtype,
  filter,
  behavior: "filter" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

/**
 * An effect measured in PowerPoint and played from its recipe
 * (`presentation-effect-recipes`); the server holds the matching template it
 * writes and reads. Nothing about its behaviours lives in this catalog.
 */
const nativeRecipe = (subtype: number) => ({
  subtype,
  filter: null,
  behavior: "recipe" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeRotation = (rotationBy: number) => ({
  subtype: 0,
  filter: null,
  behavior: "rotation" as const,
  rotationBy,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeScale = (scaleX: number, scaleY: number) => ({
  subtype: 0,
  filter: null,
  behavior: "scale" as const,
  scaleX,
  scaleY,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativePosition = (subtype: number) => ({
  subtype,
  filter: null,
  behavior: "position" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeDimension = (
  widthFrom: string,
  widthTo: string,
  heightFrom: string,
  heightTo: string,
) => ({
  subtype: 0,
  filter: "fade",
  behavior: "dimension" as const,
  widthFrom,
  widthTo,
  heightFrom,
  heightTo,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeZoom = (subtype: 16 | 32, withFade: boolean) => ({
  subtype,
  filter: withFade ? "fade" : null,
  behavior: "zoom" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeGrowTurn = () => ({
  subtype: 0,
  filter: "fade",
  behavior: "grow-turn" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeStretch = () => ({
  subtype: 10,
  filter: null,
  behavior: "stretch" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeStretchy = () => ({
  subtype: 0,
  filter: "fade",
  behavior: "stretchy" as const,
  effectGroupId: 0,
  effectAccel: 100_000,
  shapeBuild: true,
});

const nativeRiseUp = () => ({
  subtype: 0,
  filter: "fade",
  behavior: "rise-up" as const,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeOpacity = (opacity: number) => ({
  subtype: 0,
  filter: "image",
  behavior: "opacity" as const,
  opacity,
  effectGroupId: 0,
  shapeBuild: true,
});

const nativeColor = (
  subtype: number,
  attribute: "fillcolor" | "stroke.color" | "style.color",
  defaultScheme: "accent2",
  companionSets: readonly (readonly [string, string])[],
  options: Readonly<{
    behaviorOverride?: "childStyle";
    effectGroupId?: number;
    paragraphBuild?: boolean;
  }> = {},
) => ({
  subtype,
  filter: null,
  behavior: "color" as const,
  colorAttribute: attribute,
  defaultColorScheme: defaultScheme,
  colorCompanionSets: companionSets,
  colorBehaviorOverride: options.behaviorOverride ?? null,
  effectGroupId: options.effectGroupId ?? null,
  paragraphBuild: options.paragraphBuild ?? false,
});

const PRESENTATION_HAND_DRAWN_EFFECT_IDENTITIES = [
  {
    id: "appear",
    category: "entrance",
    nativePresetId: 1,
    authoring: nativeEffect(0, null),
  },
  {
    id: "fade",
    category: "entrance",
    nativePresetId: 10,
    authoring: nativeEffect(0, "fade"),
  },
  {
    id: "fly-in",
    category: "entrance",
    nativePresetId: 2,
    authoring: nativePosition(4),
  },
  {
    id: "split-in",
    category: "entrance",
    nativePresetId: 16,
    authoring: nativeEffect(42, "barn(outHorizontal)"),
  },
  {
    id: "wipe-in",
    category: "entrance",
    nativePresetId: 22,
    authoring: nativeEffect(1, "wipe(up)"),
  },
  {
    id: "random-bars-in",
    category: "entrance",
    nativePresetId: 14,
    authoring: nativeEffect(10, "randombar(horizontal)"),
  },
  {
    id: "grow-and-turn",
    category: "entrance",
    nativePresetId: 31,
    authoring: nativeGrowTurn(),
  },
  {
    id: "zoom",
    category: "entrance",
    nativePresetId: 53,
    authoring: nativeZoom(16, true),
  },
  {
    id: "basic-zoom-in",
    category: "entrance",
    nativePresetId: 23,
    authoring: nativeZoom(16, false),
  },
  {
    id: "blinds-in",
    category: "entrance",
    nativePresetId: 3,
    authoring: nativeEffect(10, "blinds(horizontal)"),
  },
  {
    id: "box-in",
    category: "entrance",
    nativePresetId: 4,
    authoring: nativeEffect(32, "box(out)"),
  },
  {
    id: "checkerboard-in",
    category: "entrance",
    nativePresetId: 5,
    authoring: nativeEffect(10, "checkerboard(across)"),
  },
  {
    id: "expand",
    category: "entrance",
    nativePresetId: 55,
    authoring: nativeDimension(
      "#ppt_w*0.70",
      "#ppt_w",
      "#ppt_h",
      "#ppt_h",
    ),
  },
  {
    id: "peek-in",
    category: "entrance",
    nativePresetId: 12,
    authoring: nativeEffect(4, "slide(fromBottom)"),
  },
  {
    id: "rise-up",
    category: "entrance",
    nativePresetId: 37,
    authoring: nativeRiseUp(),
  },
  {
    id: "stretch-in",
    category: "entrance",
    nativePresetId: 17,
    authoring: nativeStretch(),
  },
  {
    id: "strips-in",
    category: "entrance",
    nativePresetId: 18,
    authoring: nativeEffect(12, "strips(downLeft)"),
  },

  {
    id: "spin",
    category: "emphasis",
    nativePresetId: 8,
    authoring: nativeRotation(21_600_000),
  },
  {
    id: "grow-shrink",
    category: "emphasis",
    nativePresetId: 6,
    authoring: nativeScale(150_000, 150_000),
  },
  {
    id: "transparency",
    category: "emphasis",
    nativePresetId: 9,
    authoring: nativeOpacity(0.5),
  },
  {
    id: "line-color",
    category: "emphasis",
    nativePresetId: 7,
    authoring: nativeColor(2, "stroke.color", "accent2", [
      ["stroke.on", "true"],
    ]),
  },
  {
    id: "fill-color",
    category: "emphasis",
    nativePresetId: 1,
    authoring: nativeColor(2, "fillcolor", "accent2", [
      ["fill.type", "solid"],
      ["fill.on", "true"],
    ]),
  },
  {
    id: "font-color",
    category: "emphasis",
    nativePresetId: 3,
    authoring: nativeColor(2, "style.color", "accent2", [], {
      behaviorOverride: "childStyle",
      effectGroupId: 0,
      paragraphBuild: true,
    }),
  },

  {
    id: "disappear",
    category: "exit",
    nativePresetId: 1,
    authoring: nativeEffect(0, null),
  },
  {
    id: "fade-out",
    category: "exit",
    nativePresetId: 10,
    authoring: nativeEffect(0, "fade"),
  },
  {
    id: "fly-out",
    category: "exit",
    nativePresetId: 2,
    authoring: nativePosition(4),
  },
  {
    id: "split-out",
    category: "exit",
    nativePresetId: 16,
    authoring: nativeEffect(42, "barn(outHorizontal)"),
  },
  {
    id: "wipe-out",
    category: "exit",
    nativePresetId: 22,
    authoring: nativeEffect(1, "wipe(up)"),
  },
  {
    id: "random-bars-out",
    category: "exit",
    nativePresetId: 14,
    authoring: nativeEffect(10, "randombar(horizontal)"),
  },
  {
    id: "zoom-out",
    category: "exit",
    nativePresetId: 53,
    authoring: nativeZoom(32, true),
  },
  {
    id: "basic-zoom-out",
    category: "exit",
    nativePresetId: 23,
    authoring: nativeZoom(32, false),
  },
  {
    id: "blinds-out",
    category: "exit",
    nativePresetId: 3,
    authoring: nativeEffect(10, "blinds(horizontal)"),
  },
  {
    id: "box-out",
    category: "exit",
    nativePresetId: 4,
    authoring: nativeEffect(32, "box(out)"),
  },
  {
    id: "checkerboard-out",
    category: "exit",
    nativePresetId: 5,
    authoring: nativeEffect(10, "checkerboard(across)"),
  },
  {
    id: "peek-out",
    category: "exit",
    nativePresetId: 12,
    authoring: nativeEffect(4, "slide(fromBottom)"),
  },
  {
    id: "stretchy-out",
    category: "exit",
    nativePresetId: 50,
    authoring: nativeStretchy(),
  },
  {
    // Stretch run as an exit is preset 17, the entrance's own id; Stretchy
    // above is 50 and a different effect. PowerPoint's exit gallery offers
    // only Stretchy, so this one is read, played and written but never
    // offered as a new choice.
    id: "stretch-out",
    category: "exit",
    nativePresetId: 17,
    authoring: nativeStretch(),
    galleryHidden: true,
  },
  {
    id: "strips-out",
    category: "exit",
    nativePresetId: 18,
    authoring: nativeEffect(12, "strips(downLeft)"),
  },
] as const satisfies readonly {
  readonly id: string;
  readonly category: PresentationAnimationEffectCategory;
  readonly nativePresetId: number;
  /** Read, played and written, but not offered in the effect gallery. */
  readonly galleryHidden?: boolean;
  readonly authoring: {
    readonly subtype: number;
    readonly filter: string | null;
    readonly behavior:
      | "filter"
      | "rotation"
      | "scale"
      | "position"
      | "dimension"
      | "zoom"
      | "grow-turn"
      | "stretch"
      | "stretchy"
      | "rise-up"
      | "opacity"
      | "color"
      | "recipe";
    readonly rotationBy?: number;
    readonly scaleX?: number;
    readonly scaleY?: number;
    readonly opacity?: number;
    readonly colorAttribute?: "fillcolor" | "stroke.color" | "style.color";
    readonly defaultColorScheme?: "accent2";
    readonly colorCompanionSets?: readonly (readonly [string, string])[];
    readonly colorBehaviorOverride?: "childStyle" | null;
    readonly effectGroupId?: number | null;
    readonly paragraphBuild?: boolean;
    readonly shapeBuild?: boolean;
    readonly widthFrom?: string;
    readonly widthTo?: string;
    readonly heightFrom?: string;
    readonly heightTo?: string;
    readonly effectAccel?: number;
  } | null;
}[];

const PRESENTATION_ANIMATION_EFFECT_IDENTITIES = [
  ...PRESENTATION_HAND_DRAWN_EFFECT_IDENTITIES,
  ...PRESENTATION_RECIPE_EFFECT_IDENTITIES.map((entry) => ({
    id: entry.id,
    category: entry.category,
    nativePresetId: entry.nativePresetId,
    authoring: nativeRecipe(entry.nativePresetSubtype),
    // A motion path is PowerPoint's fourth effect class. In the timeline it
    // behaves as an emphasis does (the object is on stage, and stays where
    // the path ends), so that is its category here.
    motionPath: entry.motionPath,
  })),
];

export type PresentationAnimationEffectId =
  (typeof PRESENTATION_ANIMATION_EFFECT_IDENTITIES)[number]["id"];

/** The effects drawn with a hand-written keyframe (the rest play from a recipe). */
export type PresentationHandDrawnAnimationEffectId =
  (typeof PRESENTATION_HAND_DRAWN_EFFECT_IDENTITIES)[number]["id"];

/**
 * PowerPoint's own Duration for a newly added hand-drawn effect, measured in
 * PowerPoint 16.86 (each effect added with its defaults, saved, and the
 * longest behaviour read back from the timing tree). Most are 0.5 s; Box,
 * the colour emphases, Spin and Grow/Shrink are 2 s; Grow & Turn, Expand,
 * Rise Up and Stretchy are 1 s. Appear and Disappear are instant (0), and
 * Transparency lasts "until end of slide", which the model writes as 0.
 */
export const HAND_DRAWN_EFFECT_POWERPOINT_DURATION_MS: Readonly<
  Record<PresentationHandDrawnAnimationEffectId, number>
> = {
  appear: 0,
  fade: 500,
  "fly-in": 500,
  "split-in": 500,
  "wipe-in": 500,
  "random-bars-in": 500,
  "grow-and-turn": 1_000,
  zoom: 500,
  "basic-zoom-in": 500,
  "blinds-in": 500,
  "box-in": 2_000,
  "checkerboard-in": 500,
  expand: 1_000,
  "peek-in": 500,
  "rise-up": 1_000,
  "stretch-in": 500,
  "strips-in": 500,
  spin: 2_000,
  "grow-shrink": 2_000,
  transparency: 0,
  "line-color": 2_000,
  "fill-color": 2_000,
  "font-color": 2_000,
  disappear: 0,
  "fade-out": 500,
  "fly-out": 500,
  "split-out": 500,
  "wipe-out": 500,
  "random-bars-out": 500,
  "zoom-out": 500,
  "basic-zoom-out": 500,
  "blinds-out": 500,
  "box-out": 2_000,
  "checkerboard-out": 500,
  "peek-out": 500,
  "stretchy-out": 1_000,
  "stretch-out": 500,
  "strips-out": 500,
};

/** A new step's Duration when PowerPoint has nothing more specific: 0.5 s. */
export const NEW_STEP_POWERPOINT_DURATION_MS = 500;

export const PRESENTATION_ANIMATION_EFFECT_OPTIONS = [
  // Font Style's faces and Font Size's larger amounts.
  "bold",
  "italic",
  "underline",
  "150-percent",
  "400-percent",
  // Zoom's options (PowerPoint's Effect Options of Basic Zoom and Zoom).
  "in-slightly",
  "out-slightly",
  "in-from-screen-center",
  "out-from-screen-bottom",
  "out-to-screen-center",
  "in-to-screen-bottom",
  "object-center",
  "slide-center",
  "1-spoke",
  "2-spokes",
  "3-spokes",
  "4-spokes",
  "8-spokes",
  "from-bottom-left",
  "from-bottom-right",
  "from-top-left",
  "from-top-right",
  "horizontal-in",
  "horizontal-out",
  "vertical-in",
  "vertical-out",
  "from-bottom",
  "from-left",
  "from-top",
  "from-right",
  "across",
  "horizontal",
  "vertical",
  "in",
  "out",
  "down-left",
  "up-left",
  "up-right",
  "down-right",
  "25-percent",
  "50-percent",
  "75-percent",
  "100-percent",
] as const;

export type PresentationAnimationEffectOption =
  | (typeof PRESENTATION_ANIMATION_EFFECT_OPTIONS)[number]
  | `${number}-percent`
  // Font names the family it changes the text to: "font:Arial".
  | `font:${string}`;

/** The family Font changes text to when its option names none. */
export const PRESENTATION_ANIMATION_DEFAULT_FONT = "Arial";

/** A font family an effect may name: one line of printable text. */
export function presentationAnimationFontOfOption(value: unknown): string | null {
  if (typeof value !== "string" || !value.startsWith("font:")) return null;
  const font = value.slice(5);
  if (
    !font ||
    font.length > 64 ||
    font !== font.trim() ||
    /[\u0000-\u001f\u007f]/.test(font)
  ) {
    return null;
  }
  return font;
}

/**
 * The families Font's Effect Options offer; a file may name any other. An
 * option names the family itself -- the first of the editor's stack for it --
 * not the gallery's label, which may carry a second name ("PingFang SC (苹方)")
 * that no font answers to.
 */
const FONT_EFFECT_OPTIONS: readonly PresentationAnimationEffectOption[] =
  RUN_FONT_FAMILIES.filter((family) => family.id !== "system").map(
    (family) => `font:${runFontFamilyName(family)}` as const,
  );

export interface PresentationAnimationEffectOptionVariant {
  readonly option: PresentationAnimationEffectOption;
  readonly nativePresetSubtype: number;
  readonly nativeFilter: string | null;
  readonly nativeOpacity?: number;
  /**
   * Peek's other rendering. PowerPoint does not write one slide() filter: it
   * nudges the shape a little over its own size along this axis by this offset
   * and wipes it in behind that. Measured through PowerPoint's object model.
   */
  readonly nativePeekAxis?: "ppt_x" | "ppt_y";
  readonly nativePeekOffset?: string;
  readonly nativePeekWipeIn?: string;
  readonly nativePeekWipeOut?: string;
}

interface PresentationAnimationEffectOptionSpec {
  readonly defaultOption: PresentationAnimationEffectOption;
  readonly variants: readonly PresentationAnimationEffectOptionVariant[];
}

const SPLIT_EFFECT_OPTIONS = [
  {
    option: "horizontal-in",
    nativePresetSubtype: 26,
    nativeFilter: "barn(inHorizontal)",
  },
  {
    option: "horizontal-out",
    nativePresetSubtype: 42,
    nativeFilter: "barn(outHorizontal)",
  },
  {
    option: "vertical-in",
    nativePresetSubtype: 21,
    nativeFilter: "barn(inVertical)",
  },
  {
    option: "vertical-out",
    nativePresetSubtype: 37,
    nativeFilter: "barn(outVertical)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const WIPE_EFFECT_OPTIONS = [
  { option: "from-bottom", nativePresetSubtype: 1, nativeFilter: "wipe(up)" },
  { option: "from-left", nativePresetSubtype: 2, nativeFilter: "wipe(right)" },
  { option: "from-top", nativePresetSubtype: 4, nativeFilter: "wipe(down)" },
  { option: "from-right", nativePresetSubtype: 8, nativeFilter: "wipe(left)" },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const FLY_EFFECT_OPTIONS = [
  { option: "from-bottom", nativePresetSubtype: 4, nativeFilter: null },
  { option: "from-left", nativePresetSubtype: 8, nativeFilter: null },
  { option: "from-top", nativePresetSubtype: 1, nativeFilter: null },
  { option: "from-right", nativePresetSubtype: 2, nativeFilter: null },
  // The four corners PowerPoint offers beside the sides.
  { option: "from-bottom-left", nativePresetSubtype: 12, nativeFilter: null },
  { option: "from-bottom-right", nativePresetSubtype: 6, nativeFilter: null },
  { option: "from-top-left", nativePresetSubtype: 9, nativeFilter: null },
  { option: "from-top-right", nativePresetSubtype: 3, nativeFilter: null },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const STRETCH_EFFECT_OPTIONS = [
  { option: "across", nativePresetSubtype: 10, nativeFilter: null },
  { option: "from-bottom", nativePresetSubtype: 4, nativeFilter: null },
  { option: "from-left", nativePresetSubtype: 8, nativeFilter: null },
  { option: "from-right", nativePresetSubtype: 2, nativeFilter: null },
  { option: "from-top", nativePresetSubtype: 1, nativeFilter: null },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const RANDOM_BARS_EFFECT_OPTIONS = [
  {
    option: "horizontal",
    nativePresetSubtype: 10,
    nativeFilter: "randombar(horizontal)",
  },
  {
    option: "vertical",
    nativePresetSubtype: 5,
    nativeFilter: "randombar(vertical)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const BLINDS_EFFECT_OPTIONS = [
  {
    option: "horizontal",
    nativePresetSubtype: 10,
    nativeFilter: "blinds(horizontal)",
  },
  {
    option: "vertical",
    nativePresetSubtype: 5,
    nativeFilter: "blinds(vertical)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

// Measured by driving PowerPoint itself: the inward Box build is
// presetSubtype 16 with filter box(in) and the outward one is 32 with
// box(out), for the entrance and the exit alike.
const BOX_EFFECT_OPTIONS = [
  { option: "in", nativePresetSubtype: 16, nativeFilter: "box(in)" },
  { option: "out", nativePresetSubtype: 32, nativeFilter: "box(out)" },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const CHECKERBOARD_EFFECT_OPTIONS = [
  {
    option: "horizontal",
    nativePresetSubtype: 10,
    nativeFilter: "checkerboard(across)",
  },
  {
    option: "vertical",
    nativePresetSubtype: 5,
    nativeFilter: "checkerboard(down)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const PEEK_EFFECT_OPTIONS = [
  {
    option: "from-bottom",
    nativePresetSubtype: 4,
    nativeFilter: "slide(fromBottom)",
    nativePeekAxis: "ppt_y",
    nativePeekOffset: "+#ppt_h*1.125000",
    nativePeekWipeIn: "wipe(up)",
    nativePeekWipeOut: "wipe(down)",
  },
  {
    option: "from-left",
    nativePresetSubtype: 8,
    nativeFilter: "slide(fromLeft)",
    nativePeekAxis: "ppt_x",
    nativePeekOffset: "-#ppt_w*1.125000",
    nativePeekWipeIn: "wipe(right)",
    nativePeekWipeOut: "wipe(left)",
  },
  {
    option: "from-top",
    nativePresetSubtype: 1,
    nativeFilter: "slide(fromTop)",
    nativePeekAxis: "ppt_y",
    nativePeekOffset: "-#ppt_h*1.125000",
    nativePeekWipeIn: "wipe(down)",
    nativePeekWipeOut: "wipe(up)",
  },
  {
    option: "from-right",
    nativePresetSubtype: 2,
    nativeFilter: "slide(fromRight)",
    nativePeekAxis: "ppt_x",
    nativePeekOffset: "+#ppt_w*1.125000",
    nativePeekWipeIn: "wipe(left)",
    nativePeekWipeOut: "wipe(right)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

const STRIPS_EFFECT_OPTIONS = [
  {
    option: "down-left",
    nativePresetSubtype: 12,
    nativeFilter: "strips(downLeft)",
  },
  {
    option: "up-left",
    nativePresetSubtype: 9,
    nativeFilter: "strips(upLeft)",
  },
  {
    option: "up-right",
    nativePresetSubtype: 3,
    nativeFilter: "strips(upRight)",
  },
  {
    option: "down-right",
    nativePresetSubtype: 6,
    nativeFilter: "strips(downRight)",
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

/**
 * PowerPoint's Transparency emphasis exposes four AMOUNT choices. OOXML stores
 * the resulting visible opacity (so 75% transparent is `0.25`) in both the
 * `style.opacity` set behavior and the image-filter property list.
 */
const TRANSPARENCY_EFFECT_OPTIONS = [
  {
    option: "25-percent",
    nativePresetSubtype: 0,
    nativeFilter: "image",
    nativeOpacity: 0.75,
  },
  {
    option: "50-percent",
    nativePresetSubtype: 0,
    nativeFilter: "image",
    nativeOpacity: 0.5,
  },
  {
    option: "75-percent",
    nativePresetSubtype: 0,
    nativeFilter: "image",
    nativeOpacity: 0.25,
  },
  {
    option: "100-percent",
    nativePresetSubtype: 0,
    nativeFilter: "image",
    nativeOpacity: 0,
  },
] as const satisfies readonly PresentationAnimationEffectOptionVariant[];

/**
 * The name of the option an effect drawn by hand plays without one, for the
 * effects PowerPoint's measured recipes add options to. Basic Zoom zooms in
 * to enter and out to leave; Zoom grows from the object's own centre.
 */
const HAND_DRAWN_DEFAULT_OPTION = {
  "basic-zoom-in": "in",
  "basic-zoom-out": "out",
  zoom: "object-center",
  "zoom-out": "object-center",
} as const satisfies Record<
  (typeof PRESENTATION_RECIPE_EFFECT_EXTENSIONS)[number]["id"],
  PresentationAnimationEffectOption
>;

const PRESENTATION_ANIMATION_EFFECT_OPTION_SPECS = new Map<
  PresentationAnimationEffectId,
  PresentationAnimationEffectOptionSpec
>([
  ...PRESENTATION_RECIPE_EFFECT_EXTENSIONS.map((entry) => {
    const drawn = PRESENTATION_HAND_DRAWN_EFFECT_IDENTITIES.find(
      (candidate) => candidate.id === entry.id,
    )!;
    return [
      entry.id,
      {
        defaultOption: HAND_DRAWN_DEFAULT_OPTION[entry.id],
        variants: [
          {
            option: HAND_DRAWN_DEFAULT_OPTION[entry.id],
            nativePresetSubtype: drawn.authoring.subtype,
            nativeFilter: drawn.authoring.filter,
          },
          ...entry.options.map(([option, subtype]) => ({
            option: option as PresentationAnimationEffectOption,
            nativePresetSubtype: subtype,
            nativeFilter: drawn.authoring.filter,
          })),
        ],
      },
    ] as const;
  }),
  ...PRESENTATION_RECIPE_EFFECT_IDENTITIES.filter(
    (entry) => entry.options.length > 0 && entry.defaultOption !== null,
  ).map(
    (entry) =>
      [
        entry.id,
        {
          defaultOption: entry.defaultOption as PresentationAnimationEffectOption,
          variants: entry.options.map(([option, subtype]) => ({
            option: option as PresentationAnimationEffectOption,
            nativePresetSubtype: subtype,
            nativeFilter: null,
          })),
        },
      ] as const,
  ),
  ...(["split-in", "split-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "horizontal-out", variants: SPLIT_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["wipe-in", "wipe-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "from-bottom", variants: WIPE_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["fly-in", "fly-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "from-bottom", variants: FLY_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["stretch-in", "stretch-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "across", variants: STRETCH_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["random-bars-in", "random-bars-out"] as const).map(
    (kind) =>
      [
        kind,
        {
          defaultOption: "horizontal",
          variants: RANDOM_BARS_EFFECT_OPTIONS,
        },
      ] as const,
  ),
  ...(["blinds-in", "blinds-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "horizontal", variants: BLINDS_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["box-in", "box-out"] as const).map(
    (kind) =>
      [kind, { defaultOption: "out", variants: BOX_EFFECT_OPTIONS }] as const,
  ),
  ...(["checkerboard-in", "checkerboard-out"] as const).map(
    (kind) =>
      [
        kind,
        {
          defaultOption: "horizontal",
          variants: CHECKERBOARD_EFFECT_OPTIONS,
        },
      ] as const,
  ),
  ...(["peek-in", "peek-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "from-bottom", variants: PEEK_EFFECT_OPTIONS },
      ] as const,
  ),
  ...(["strips-in", "strips-out"] as const).map(
    (kind) =>
      [
        kind,
        { defaultOption: "down-left", variants: STRIPS_EFFECT_OPTIONS },
      ] as const,
  ),
  [
    "transparency",
    {
      defaultOption: "50-percent",
      variants: TRANSPARENCY_EFFECT_OPTIONS,
    },
  ],
]);

export interface PresentationAnimationEffectCatalogEntry {
  readonly id: PresentationAnimationEffectId;
  readonly category: PresentationAnimationEffectCategory;
  readonly nativePresetClass: "entr" | "emph" | "exit" | "path";
  /** One of PowerPoint's motion paths (offered in their own gallery section). */
  readonly motionPath: boolean;
  readonly nativePresetId: number;
  readonly nativePresetSubtype: number;
  readonly nativeFilter: string | null;
  readonly nativeTransition: "in" | "out" | null;
  readonly nativeBehavior:
    | "filter"
    | "rotation"
    | "scale"
    | "position"
    | "dimension"
    | "zoom"
    | "grow-turn"
    | "stretch"
    | "stretchy"
    | "rise-up"
    | "opacity"
    | "color"
    | "recipe"
    | null;
  readonly nativeRotationBy: number | null;
  readonly nativeScaleX: number | null;
  readonly nativeScaleY: number | null;
  readonly nativeOpacity: number | null;
  readonly nativeColorAttribute: "fillcolor" | "stroke.color" | "style.color" | null;
  readonly nativeDefaultColorScheme: "accent2" | null;
  readonly nativeColorCompanionSets: readonly (readonly [string, string])[];
  readonly nativeColorBehaviorOverride: "childStyle" | null;
  readonly nativeEffectGroupId: number | null;
  readonly nativeParagraphBuild: boolean;
  readonly nativeShapeBuild: boolean;
  readonly nativeEffectAccel: number | null;
  readonly nativeDimensionWidthFrom: string | null;
  readonly nativeDimensionWidthTo: string | null;
  readonly nativeDimensionHeightFrom: string | null;
  readonly nativeDimensionHeightTo: string | null;
  readonly options: readonly PresentationAnimationEffectOption[];
  readonly defaultOption: PresentationAnimationEffectOption | null;
  readonly authoring: "native" | "preserve-only";
  /** Read, played and written, but not offered in the effect gallery. */
  readonly galleryHidden: boolean;
  /** PowerPoint's gallery group (Basic / Subtle / Moderate / Exciting). */
  readonly group: PresentationAnimationEffectGroup;
  /** PowerPoint's own duration for a recipe-drawn effect; `null` for the rest. */
  readonly recipeDurationMs: number | null;
  readonly controlId: `PPTCTL.RIBBON.ANIMATIONS.${string}`;
}

export const PRESENTATION_ANIMATION_EFFECT_GROUPS = [
  "basic",
  "subtle",
  "moderate",
  "exciting",
  // Motion paths: Basic shapes, Lines & Curves, Special.
  "lines",
  "special",
] as const;
export type PresentationAnimationEffectGroup =
  (typeof PRESENTATION_ANIMATION_EFFECT_GROUPS)[number];

const RECIPE_DURATION_BY_ID: ReadonlyMap<string, number> = new Map(
  PRESENTATION_RECIPE_EFFECT_IDENTITIES.map((entry) => [
    entry.id,
    entry.defaultDurationMs,
  ]),
);

const ANIMATION_PRESET_CLASS_BY_CATEGORY = Object.freeze({
  entrance: "entr",
  emphasis: "emph",
  exit: "exit",
} satisfies Record<
  PresentationAnimationEffectCategory,
  "entr" | "emph" | "exit"
>);

function effectControlToken(value: string): string {
  return value.replaceAll("-", "_").toUpperCase();
}

/** Expanded immutable catalog consumed by UI, commands and native I/O. */
export const PRESENTATION_WEB_ANIMATION_EFFECTS = Object.freeze(
  PRESENTATION_ANIMATION_EFFECT_IDENTITIES.map(
    (entry): PresentationAnimationEffectCatalogEntry => {
      const optionSpec = PRESENTATION_ANIMATION_EFFECT_OPTION_SPECS.get(
        entry.id,
      );
      return {
        ...entry,
        galleryHidden: "galleryHidden" in entry && entry.galleryHidden === true,
        motionPath: "motionPath" in entry && entry.motionPath === true,
        group: PRESENTATION_EFFECT_GALLERY_GROUPS[entry.id] ?? "basic",
        recipeDurationMs: RECIPE_DURATION_BY_ID.get(entry.id) ?? null,
        nativePresetClass:
          "motionPath" in entry && entry.motionPath === true
            ? "path"
            : ANIMATION_PRESET_CLASS_BY_CATEGORY[entry.category],
        nativePresetSubtype: entry.authoring?.subtype ?? 0,
        nativeFilter: entry.authoring?.filter ?? null,
        nativeBehavior: entry.authoring?.behavior ?? null,
        nativeRotationBy:
          entry.authoring?.behavior === "rotation"
            ? entry.authoring.rotationBy
            : null,
        nativeScaleX:
          entry.authoring?.behavior === "scale" ? entry.authoring.scaleX : null,
        nativeScaleY:
          entry.authoring?.behavior === "scale" ? entry.authoring.scaleY : null,
        nativeOpacity:
          entry.authoring?.behavior === "opacity" ? entry.authoring.opacity : null,
        nativeColorAttribute:
          entry.authoring?.behavior === "color"
            ? entry.authoring.colorAttribute
            : null,
        nativeDefaultColorScheme:
          entry.authoring?.behavior === "color"
            ? entry.authoring.defaultColorScheme
            : null,
        nativeColorCompanionSets:
          entry.authoring?.behavior === "color"
            ? entry.authoring.colorCompanionSets
            : [],
        nativeColorBehaviorOverride:
          entry.authoring?.behavior === "color"
            ? entry.authoring.colorBehaviorOverride
            : null,
        nativeEffectGroupId:
          entry.authoring && "effectGroupId" in entry.authoring
            ? entry.authoring.effectGroupId
            : null,
        nativeParagraphBuild:
          entry.authoring && "paragraphBuild" in entry.authoring
            ? entry.authoring.paragraphBuild
            : false,
        nativeShapeBuild:
          entry.authoring && "shapeBuild" in entry.authoring
            ? entry.authoring.shapeBuild === true
            : false,
        nativeEffectAccel:
          entry.authoring && "effectAccel" in entry.authoring
            ? entry.authoring.effectAccel
            : null,
        nativeDimensionWidthFrom:
          entry.authoring?.behavior === "dimension"
            ? entry.authoring.widthFrom
            : null,
        nativeDimensionWidthTo:
          entry.authoring?.behavior === "dimension"
            ? entry.authoring.widthTo
            : null,
        nativeDimensionHeightFrom:
          entry.authoring?.behavior === "dimension"
            ? entry.authoring.heightFrom
            : null,
        nativeDimensionHeightTo:
          entry.authoring?.behavior === "dimension"
            ? entry.authoring.heightTo
            : null,
        nativeTransition:
          entry.authoring?.behavior !== "filter" ||
          entry.authoring.filter === null
            ? null
            : entry.category === "exit"
              ? "out"
              : "in",
        options:
          entry.id === "font"
            ? FONT_EFFECT_OPTIONS
            : (optionSpec?.variants.map((variant) => variant.option) ?? []),
        defaultOption:
          entry.id === "font"
            ? (`font:${PRESENTATION_ANIMATION_DEFAULT_FONT}` as const)
            : (optionSpec?.defaultOption ?? null),
        authoring: entry.authoring ? "native" : "preserve-only",
        controlId: `PPTCTL.RIBBON.ANIMATIONS.${entry.category.toUpperCase()}.${effectControlToken(entry.id)}`,
      };
    },
  ),
);

/** Exact set currently accepted by every native writer/patchback seam. */
export const PRESENTATION_NATIVE_AUTHORABLE_ANIMATION_EFFECTS = Object.freeze(
  PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (entry) => entry.authoring === "native",
  ).map((entry) => entry.id),
);

const PRESENTATION_ANIMATION_EFFECT_BY_ID = new Map(
  PRESENTATION_WEB_ANIMATION_EFFECTS.map((entry) => [entry.id, entry]),
);

export function presentationAnimationEffect(
  value: unknown,
): PresentationAnimationEffectCatalogEntry | null {
  return typeof value === "string"
    ? (PRESENTATION_ANIMATION_EFFECT_BY_ID.get(
        value as PresentationAnimationEffectId,
      ) ?? null)
    : null;
}

export function isPresentationAnimationEffectId(
  value: unknown,
): value is PresentationAnimationEffectId {
  return presentationAnimationEffect(value) !== null;
}

export interface PresentationAnimationEffectOptionResolution {
  /** Canonical persisted value; the default variant is represented by omission. */
  readonly effectOption?: PresentationAnimationEffectOption;
  /** Effective radio selection, including the default variant. */
  readonly selectedOption: PresentationAnimationEffectOption | null;
  readonly nativePresetSubtype: number;
  readonly nativeFilter: string | null;
  readonly nativeOpacity: number | null;
}

/** Resolve an exact option without trimming, coercion, or nearest-value fallback. */
export function resolvePresentationAnimationEffectOption(
  kind: unknown,
  value: unknown,
): PresentationAnimationEffectOptionResolution | null {
  const effect = presentationAnimationEffect(kind);
  if (!effect || effect.authoring !== "native") return null;
  // Native files may store an exact amount beyond the four gallery choices.
  // Accept canonical hundredths of one percent without expanding that menu.
  if (kind === "transparency" && typeof value === "string" &&
      /^(?:0|[1-9]\d?|100)(?:\.\d{1,2})?-percent$/.test(value)) {
    const percent = Number(value.slice(0, -8));
    if (percent <= 100 && `${percent}-percent` === value) {
      const option = value as PresentationAnimationEffectOption;
      return {
        ...(percent === 50 ? {} : { effectOption: option }),
        selectedOption: option,
        nativePresetSubtype: 0,
        nativeFilter: "image",
        nativeOpacity: (10000 - Math.round(percent * 100)) / 10000,
      };
    }
  }
  if (effect.id === "font") {
    // Any family is an option: the gallery lists some, a file may name others.
    const font =
      value === undefined || value === null
        ? PRESENTATION_ANIMATION_DEFAULT_FONT
        : presentationAnimationFontOfOption(value);
    if (font === null) return null;
    const selectedOption = `font:${font}` as const;
    return {
      ...(font === PRESENTATION_ANIMATION_DEFAULT_FONT ? {} : { effectOption: selectedOption }),
      selectedOption,
      nativePresetSubtype: effect.nativePresetSubtype,
      nativeFilter: effect.nativeFilter,
      nativeOpacity: effect.nativeOpacity,
    };
  }
  const spec = PRESENTATION_ANIMATION_EFFECT_OPTION_SPECS.get(effect.id);
  if (!spec) {
    return value === undefined || value === null
      ? {
          selectedOption: null,
          nativePresetSubtype: effect.nativePresetSubtype,
          nativeFilter: effect.nativeFilter,
          nativeOpacity: effect.nativeOpacity,
        }
      : null;
  }
  const selectedOption =
    value === undefined || value === null
      ? spec.defaultOption
      : typeof value === "string" &&
          spec.variants.some((variant) => variant.option === value)
        ? (value as PresentationAnimationEffectOption)
        : null;
  if (selectedOption === null) return null;
  const variant = spec.variants.find(
    (candidate) => candidate.option === selectedOption,
  );
  if (!variant) return null;
  return {
    ...(selectedOption === spec.defaultOption ? {} : { effectOption: selectedOption }),
    selectedOption,
    nativePresetSubtype: variant.nativePresetSubtype,
    nativeFilter: variant.nativeFilter,
    nativeOpacity: variant.nativeOpacity ?? effect.nativeOpacity,
  };
}

export function presentationAnimationEffectOptionVariants(
  kind: PresentationAnimationEffectId,
): readonly PresentationAnimationEffectOptionVariant[] {
  return PRESENTATION_ANIMATION_EFFECT_OPTION_SPECS.get(kind)?.variants ?? [];
}

export interface PresentationAnimationEffectColorResolution {
  /** Canonical persisted RGB override; omission keeps PowerPoint's theme accent2. */
  readonly effectColor?: string;
  readonly nativeColor:
    | Readonly<{ kind: "scheme"; value: "accent2" }>
    | Readonly<{ kind: "srgb"; value: string }>;
}

/**
 * Resolve Fill Color's exact native target without aliases or nearest-colour
 * fallback. PowerPoint's default is theme accent2 and remains byte-clean by
 * omission; an explicit picker choice is a canonical lower-case #rrggbb value.
 */
export function resolvePresentationAnimationEffectColor(
  kind: unknown,
  value: unknown,
): PresentationAnimationEffectColorResolution | null {
  const effect = presentationAnimationEffect(kind);
  if (
    !effect ||
    effect.authoring !== "native" ||
    effect.nativeBehavior !== "color" ||
    (effect.nativeColorAttribute !== "fillcolor" &&
      effect.nativeColorAttribute !== "stroke.color" &&
      effect.nativeColorAttribute !== "style.color") ||
    effect.nativeDefaultColorScheme !== "accent2"
  ) {
    return null;
  }
  if (value === undefined || value === null) {
    return { nativeColor: { kind: "scheme", value: "accent2" } };
  }
  if (typeof value !== "string" || !/^#[0-9a-f]{6}$/i.test(value)) {
    return null;
  }
  const effectColor = value.toLowerCase();
  return {
    effectColor,
    nativeColor: { kind: "srgb", value: effectColor.slice(1).toUpperCase() },
  };
}

export function presentationAnimationEffectsByCategory(
  category: PresentationAnimationEffectCategory,
): readonly PresentationAnimationEffectCatalogEntry[] {
  return PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (entry) => entry.category === category && !entry.galleryHidden,
  );
}

/** ANI01's original native subset, retained as a migration-safe alias. */
export const BASIC_ANIMATION_EFFECTS = ["appear", "fade", "fly-in"] as const;
export type BasicAnimationEffect = (typeof BASIC_ANIMATION_EFFECTS)[number];

/**
 * Web effects whose visual + OOXML meaning is exact in the from-scratch writer
 * but which are intentionally outside the authenticated imported-PPTX grammar.
 * Keeping this list separate from ANI01's basic effects prevents a local-only
 * export capability from being mistaken for original-package patchback authority.
 */
export const SCRATCH_ANIMATION_EFFECTS = ["fade-out", "fly-out"] as const;
export type ScratchAnimationEffect = (typeof SCRATCH_ANIMATION_EFFECTS)[number];















export const SCRATCH_ANIMATION_MOTION_PATHS = [
  "line-right",
  "line-left",
  "line-down",
  "line-up",
] as const satisfies readonly MotionPathPreset[];

/** Choices that may be newly authored from the Web animation pane. */
export const WEB_AUTHORABLE_TIMELINE_ANIMATIONS = [
  "none",
  ...PRESENTATION_NATIVE_AUTHORABLE_ANIMATION_EFFECTS,
] as readonly TimelineAnimation[];

export type TimelineAnimationAuthoringRequest = Readonly<{
  slideSource: unknown;
  blockSource: unknown;
  animation: TimelineAnimation;
  motionPathPreset?: MotionPathPreset;
}>;

export type TimelineAnimationAuthoringCapability = Readonly<
  | {
      enabled: true;
      authority: "imported-patchback" | "from-scratch";
      reason: null;
    }
  | {
      enabled: false;
      authority: "unavailable";
      reason:
        | "preserve-only-source"
        | "unaddressable-imported-target"
        | "imported-patchback-limited"
        | "unsupported-effect";
    }
>;

/**
 * Timeline effects include exact Appear without widening the legacy per-block
 * animation vocabulary. Appear is rendered by the timeline reveal gate itself:
 * the block becomes visible with no CSS keyframe, so it never masquerades as a
 * fade.
 */
/**
 * Persisted timeline kind. `pulse` and `bounce` remain readable for old decks,
 * but are intentionally absent from the Web catalog because they are not in the
 * current PowerPoint for the web effect list.
 */
export type TimelineAnimation =
  "none" | PresentationAnimationEffectId | "pulse" | "bounce";
export const TIMELINE_ANIMATIONS = [
  "none",
  ...PRESENTATION_WEB_ANIMATION_EFFECTS.map((entry) => entry.id),
  "pulse",
  "bounce",
] as const satisfies readonly TimelineAnimation[];

const TIMELINE_ANIMATION_SET: ReadonlySet<string> = new Set(
  TIMELINE_ANIMATIONS,
);
const BASIC_ANIMATION_EFFECT_SET: ReadonlySet<string> = new Set(
  BASIC_ANIMATION_EFFECTS,
);

export function isBasicAnimationEffect(
  value: unknown,
): value is BasicAnimationEffect {
  return typeof value === "string" && BASIC_ANIMATION_EFFECT_SET.has(value);
}

export function isTimelineAnimation(
  value: unknown,
): value is TimelineAnimation {
  return typeof value === "string" && TIMELINE_ANIMATION_SET.has(value);
}

export function normalizeTimelineAnimation(value: unknown): TimelineAnimation {
  if (typeof value !== "string") return DEFAULT_BLOCK_ANIMATION;
  const candidate = value.trim().toLowerCase();
  return isTimelineAnimation(candidate) ? candidate : DEFAULT_BLOCK_ANIMATION;
}

export const PRESENTATION_SLIDE_ANIMATION_SCHEMA_VERSION =
  "pptx-slide-animation-v2" as const;

const PRESENTATION_SLIDE_ANIMATION_RECEIPT_KEYS: ReadonlySet<string> = new Set([
  "slideAnimationSchemaVersion",
  "slideAnimationDisposition",
  "slideAnimationContainerPresent",
  "slideAnimationOriginalEffectCount",
  "slideAnimationOriginalSequence",
  "slideAnimationFingerprint",
  "slideAnimationRejectReason",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/** True only for a slide that has no imported native-source custody. */
export function isLocallyAuthoredSlideAnimationSource(
  source: unknown,
): boolean {
  return (
    source === undefined ||
    source === null ||
    (isRecord(source) && Object.keys(source).length === 0)
  );
}

/**
 * Imported slides are authorable only when their immutable timing receipt is a
 * complete native v2 receipt. Locally-created slides have no source evidence
 * (or an empty source) and remain editable; old/opaque imported slides fail
 * closed even when an older browser import did not stamp a native slide id.
 */
export function hasEditableSlideAnimationSource(source: unknown): boolean {
  if (isLocallyAuthoredSlideAnimationSource(source)) return true;
  if (!isRecord(source)) return false;
  const receiptKeys = Object.keys(source).filter((key) =>
    key.startsWith("slideAnimation"),
  );
  if (receiptKeys.length === 0) return false;
  if (
    receiptKeys.some(
      (key) => !PRESENTATION_SLIDE_ANIMATION_RECEIPT_KEYS.has(key),
    ) ||
    source.slideAnimationSchemaVersion !==
      PRESENTATION_SLIDE_ANIMATION_SCHEMA_VERSION ||
    source.slideAnimationDisposition !== "native" ||
    typeof source.slideAnimationContainerPresent !== "boolean" ||
    typeof source.slideAnimationOriginalEffectCount !== "number" ||
    !Number.isInteger(source.slideAnimationOriginalEffectCount) ||
    source.slideAnimationOriginalEffectCount < 0 ||
    source.slideAnimationOriginalEffectCount > MAX_TIMELINE_STEPS ||
    !Array.isArray(source.slideAnimationOriginalSequence) ||
    source.slideAnimationOriginalSequence.length !==
      source.slideAnimationOriginalEffectCount ||
    typeof source.slideAnimationFingerprint !== "string" ||
    !/^[0-9a-f]{64}$/.test(source.slideAnimationFingerprint) ||
    Object.hasOwn(source, "slideAnimationRejectReason")
  ) {
    return false;
  }
  // Effects need a container; a container may hold none (an empty
  // `<p:timing>` PowerPoint 2007-2010 keep on a slide whose effects are gone).
  if (
    !source.slideAnimationContainerPresent &&
    source.slideAnimationOriginalEffectCount > 0
  ) {
    return false;
  }
  return source.slideAnimationOriginalSequence.every((entry) => {
    if (!isRecord(entry)) return false;
    const entryKeys = Object.keys(entry);
    const expectedKeys = new Set([
      "kind",
      "nativeShapeId",
      "trigger",
      "delayMs",
      "durationMs",
      ...(entry.effectOption === undefined ? [] : ["effectOption"]),
      ...(entry.effectColor === undefined ? [] : ["effectColor"]),
      ...(entry.effectSettings === undefined ? [] : ["effectSettings"]),
      ...(entry.paragraphs === undefined ? [] : ["paragraphs"]),
      ...(entry.triggerNativeShapeId === undefined
        ? []
        : ["triggerNativeShapeId"]),
    ]);
    if (
      entry.paragraphs !== undefined &&
      (() => {
        const range = normalizeStepParagraphs(entry.paragraphs);
        const raw = entry.paragraphs as { start?: unknown; end?: unknown } | null;
        return !range || Object.keys(raw ?? {}).length !== 2 || range.start !== raw?.start || range.end !== raw?.end;
      })()
    ) {
      return false;
    }
    if (
      entry.effectSettings !== undefined &&
      !effectSettingsEqual(
        normalizeEffectSettings(entry.effectSettings),
        entry.effectSettings as PresentationAnimationEffectSettings,
      )
    ) {
      return false;
    }
    const option = resolvePresentationAnimationEffectOption(
      entry.kind,
      entry.effectOption,
    );
    const color = resolvePresentationAnimationEffectColor(
      entry.kind,
      entry.effectColor,
    );
    const effect = presentationAnimationEffect(entry.kind);
    return (
      entryKeys.length === expectedKeys.size &&
      entryKeys.every((key) => expectedKeys.has(key)) &&
      effect?.authoring === "native" &&
      option !== null &&
      option.effectOption === entry.effectOption &&
      (effect.nativeBehavior === "color"
        ? color !== null && color.effectColor === entry.effectColor
        : entry.effectColor === undefined) &&
      typeof entry.nativeShapeId === "string" &&
      /^shp[1-9][0-9]*$/.test(entry.nativeShapeId) &&
      (entry.triggerNativeShapeId === undefined ||
        (typeof entry.triggerNativeShapeId === "string" &&
          /^shp[1-9][0-9]*$/.test(entry.triggerNativeShapeId) &&
          entry.triggerNativeShapeId !== entry.nativeShapeId)) &&
      isTimelineTrigger(entry.trigger) &&
      isEditableAnimationTimingMs(entry.delayMs) &&
      (isEditableAnimationDurationMs(entry.durationMs) ||
        // PowerPoint's own duration for a recipe effect (80 ms a letter).
        (typeof entry.durationMs === "number" && entry.durationMs === effect.recipeDurationMs))
    );
  });
}

function animationSourceValue(
  source: Record<string, unknown>,
  camelKey: string,
  snakeKey: string,
): string | null {
  const value = source[camelKey] ?? source[snakeKey];
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

/**
 * The single Web authoring gate for animation menu state and execution.
 *
 * Imported native timing may append only catalogued exact effects, and only to a
 * block whose immutable slide/shape identity can be addressed by patchback.
 * Locally-created decks use the same exact native-effect catalog. Approximate
 * legacy effects and motion paths stay renderable when already present, but
 * cannot be newly authored from Web UI.
 */
export function timelineAnimationAuthoringCapability(
  request: TimelineAnimationAuthoringRequest,
): TimelineAnimationAuthoringCapability {
  const local = isLocallyAuthoredSlideAnimationSource(request.slideSource);
  if (!local && !hasEditableSlideAnimationSource(request.slideSource)) {
    return {
      enabled: false,
      authority: "unavailable",
      reason: "preserve-only-source",
    };
  }

  if (!local) {
    if (!isRecord(request.slideSource) || !isRecord(request.blockSource)) {
      return {
        enabled: false,
        authority: "unavailable",
        reason: "unaddressable-imported-target",
      };
    }
    const slideId = animationSourceValue(
      request.slideSource,
      "slideId",
      "slide_id",
    );
    const blockSlideId = animationSourceValue(
      request.blockSource,
      "slideId",
      "slide_id",
    );
    const shapeId = animationSourceValue(
      request.blockSource,
      "shapeId",
      "shape_id",
    );
    if (
      !slideId ||
      blockSlideId !== slideId ||
      !shapeId ||
      !/^shp[1-9][0-9]*$/.test(shapeId)
    ) {
      return {
        enabled: false,
        authority: "unavailable",
        reason: "unaddressable-imported-target",
      };
    }
  }

  if (request.motionPathPreset !== undefined) {
    // Unconditional, and said once. This used to branch on whether the preset
    // was one of SCRATCH_ANIMATION_MOTION_PATHS and whether the effect beside
    // it was authorable, and then answer `unsupported-effect` either way -- two
    // identical returns reading as a live distinction. There is none: no motion
    // path may be newly authored here at all, because PowerPoint for the web's
    // Add Animation has no Motion Paths category (see
    // SCRATCH_ANIMATION_MOTION_PATHS above for why this is no longer a question
    // about the package writer). A path already on a step stays playable, and
    // removable through the animation pane.
    return {
      enabled: false,
      authority: "unavailable",
      reason: "unsupported-effect",
    };
  }

  const effect = presentationAnimationEffect(request.animation);
  if (effect?.authoring === "native") {
    return {
      enabled: true,
      authority: local ? "from-scratch" : "imported-patchback",
      reason: null,
    };
  }
  return {
    enabled: false,
    authority: "unavailable",
    reason: "unsupported-effect",
  };
}

/** Max length of a step's `blockId` reference; bounds a hand-edited value. */
const MAX_BLOCK_ID_LENGTH = 200;

/**
 * Max length of a step's optional `triggerBlockId` reference; mirrors the
 * `blockId` bound (both name a block, so they share the same shape constraint).
 */
const MAX_TRIGGER_BLOCK_ID_LENGTH = MAX_BLOCK_ID_LENGTH;

/** Max length of a step's optional `id`; bounds a hand-edited value. */
const MAX_STEP_ID_LENGTH = 200;

/**
 * A single, normalized timeline step. `id` is an optional stable identity for
 * the editor (drag / reorder keys); when absent the renderer keys by index.
 * `animation` reuses the per-block entrance set so the timeline and the
 * single-block control share one vocabulary and one set of CSS keyframes.
 */
export interface TimelineStep {
  /** Optional stable id for editor keys; absent steps key by index. */
  id?: string;
  /** The block this step animates. Bounded, non-empty string. */
  blockId: string;
  /** The entrance animation to play (shared with per-block fx). */
  animation: TimelineAnimation;
  /** Non-default directional Effect Options value; omission means Office default. */
  effectOption?: PresentationAnimationEffectOption;
  /** Explicit native color-emphasis target; omission keeps theme accent2. */
  effectColor?: string;
  /**
   * PowerPoint's Effect Options beyond the effect, its direction and timing
   * (Smooth start / end, Auto-reverse, Repeat, Rewind, After animation,
   * Animate text By word / By letter): `presentation-anim-effect-settings`.
   * Absent → the effect as the gallery adds it, byte-for-byte as before.
   */
  effectSettings?: PresentationAnimationEffectSettings;
  /** How this step starts relative to its neighbours. */
  trigger: TimelineTrigger;
  /** Delay (ms) before this step starts, measured from its anchor. */
  delayMs: number;
  /** How long (ms) the step's animation runs. */
  durationMs: number;
  /**
   * Optional MOTION PATH (PowerPoint's fourth animation class). When present the
   * step moves its block along the path's waypoints over `durationMs`. ADDITIVE
   * to (and independent of) `animation`: a step may carry both an entrance
   * keyframe AND a motion path, but in practice motion paths are authored on
   * their own step. Absent → the step is a plain entrance/emphasis/exit step,
   * byte-for-byte as before. Geometry + clamps live in
   * `lib/presentation-motion-path`; a malformed value is dropped at validation.
   */
  motionPath?: MotionPath;
  /**
   * Optional INTERACTIVE TRIGGER (PowerPoint "Trigger → On Click of"). When set,
   * this step is NOT part of the normal click-advance sequence; instead it fires
   * when the audience clicks the block whose id this names IN PRESENT MODE. A step
   * chained `with-previous` / `after-previous` AFTER a triggered step inherits its
   * trigger group (it plays when the same trigger block is clicked). ADDITIVE:
   * absent → the step plays in the normal click-advance flow exactly as before,
   * byte-for-byte. Invariants enforced at validation:
   *  - A non-empty, length-bounded, trimmed string (else dropped → falls back to
   *    the normal on-click sequence).
   *  - It can NEVER equal the step's own `blockId` (a self-triggering step is
   *    nonsensical — the block would have to appear to be clicked) → dropped.
   *  - When the valid block-id set is known (the slide's live blocks), a
   *    `triggerBlockId` naming a block that no longer exists is dropped (it
   *    falls back to the normal on-click sequence) so a deleted trigger block
   *    never leaves the step unreachable.
   */
  triggerBlockId?: string;
  /**
   * A By Paragraph build read from an imported package: the step plays on
   * these paragraphs of the block's text only (PowerPoint's `p:txEl/p:pRg`,
   * zero-based and inclusive). Read-only: MRAI plays it in the show and never
   * authors it, so nothing in the editor sets or changes this field. Absent →
   * the step plays on the whole block, byte-for-byte as before. A malformed
   * value is dropped at validation (the step then plays on the whole block).
   */
  paragraphs?: TimelineStepParagraphs;
}

/** Zero-based, inclusive paragraph range of a By Paragraph build step. */
export interface TimelineStepParagraphs {
  start: number;
  end: number;
}

/** Highest paragraph index a build step may name (matches the importer). */
export const MAX_STEP_PARAGRAPH_INDEX = 9999;

/** A well-formed paragraph range, or `null` (never a partially valid one). */
export function normalizeStepParagraphs(
  raw: unknown,
): TimelineStepParagraphs | null {
  const record = asRecord(raw);
  if (!record) return null;
  const { start, end } = record;
  if (
    typeof start !== "number" ||
    typeof end !== "number" ||
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > MAX_STEP_PARAGRAPH_INDEX
  ) {
    return null;
  }
  return { start, end };
}

/**
 * Type guard: is `value` exactly one of the allowed trigger identifiers? Exact,
 * case-sensitive — use {@link normalizeTimelineTrigger} for lenient input.
 */
export function isTimelineTrigger(value: unknown): value is TimelineTrigger {
  return typeof value === "string" && TIMELINE_TRIGGER_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed trigger. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_TIMELINE_TRIGGER}. The
 * comparison trims surrounding whitespace and is case-insensitive; everything
 * else falls back to the default.
 */
export function normalizeTimelineTrigger(value: unknown): TimelineTrigger {
  if (typeof value !== "string") return DEFAULT_TIMELINE_TRIGGER;
  const candidate = value.trim().toLowerCase();
  return isTimelineTrigger(candidate) ? candidate : DEFAULT_TIMELINE_TRIGGER;
}

/**
 * Clamp a numeric ms value into `[min, max]`. Non-finite / non-number inputs
 * (including NaN and string numbers) collapse to `fallback`, then the result is
 * floored to a whole millisecond (no sub-ms scheduling) and clamped. Negative
 * values are pulled up to `min`.
 */
function clampMs(
  value: unknown,
  min: number,
  max: number,
  fallback: number,
): number {
  const raw =
    typeof value === "number" && Number.isFinite(value) ? value : fallback;
  const floored = Math.floor(raw);
  if (floored < min) return min;
  if (floored > max) return max;
  return floored;
}

/** Clamp a delay (ms) into the allowed range, defaulting empty/invalid input. */
export function clampDelayMs(value: unknown): number {
  return clampMs(
    value,
    MIN_STEP_DELAY_MS,
    MAX_STEP_DELAY_MS,
    DEFAULT_STEP_DELAY_MS,
  );
}

/** Clamp a duration (ms) into the allowed range, defaulting empty/invalid input. */
export function clampDurationMs(value: unknown): number {
  return clampMs(
    value,
    MIN_STEP_DURATION_MS,
    MAX_STEP_DURATION_MS,
    DEFAULT_STEP_DURATION_MS,
  );
}

/** 50ms-aligned delay values ANI03 may author without silently clamping. */
export function isEditableAnimationTimingMs(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 0 &&
    value <= MAX_STEP_DURATION_MS &&
    value % ANIMATION_TIMING_STEP_MS === 0
  );
}

/** Durations also retain ANI01's canonical 380ms default. */
export function isEditableAnimationDurationMs(value: unknown): value is number {
  return (
    isEditableAnimationTimingMs(value) || value === DEFAULT_STEP_DURATION_MS
  );
}

/** Parse the PowerPoint-style seconds fields into an exact millisecond value. */
export function parseAnimationTimingSeconds(
  value: unknown,
  field: "delayMs" | "durationMs" = "delayMs",
): number | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const text = String(value).trim();
  if (!/^\d+(?:\.\d{0,2})?$/.test(text)) return undefined;
  const milliseconds = Math.round(Number(text) * 1000);
  return (
    field === "durationMs"
      ? isEditableAnimationDurationMs(milliseconds)
      : isEditableAnimationTimingMs(milliseconds)
  )
    ? milliseconds
    : undefined;
}

/** Format persisted milliseconds for the two-decimal timing fields. */
export function formatAnimationTimingSeconds(value: unknown): string {
  const milliseconds =
    typeof value === "number" && Number.isFinite(value)
      ? Math.max(0, Math.min(MAX_STEP_DURATION_MS, Math.round(value)))
      : 0;
  // PowerPoint's Duration and Delay fields read "00.50": two whole digits.
  const [whole, fraction] = (milliseconds / 1000).toFixed(2).split(".");
  return `${whole.padStart(2, "0")}.${fraction}`;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * Validate + normalize a single RAW step into a {@link TimelineStep}, or return
 * `null` when it cannot be salvaged into a renderable step.
 *
 * The ONLY hard requirement is a non-empty `blockId` (a step that names no block
 * is meaningless and is dropped). Everything else degrades gracefully:
 *  - `animation` collapses to the per-block default for unknown/stale values.
 *  - `trigger` collapses to {@link DEFAULT_TIMELINE_TRIGGER}.
 *  - `delayMs` / `durationMs` are clamped into their bounds (NaN/∞ → default).
 *  - `id` is preserved only when it is a short non-empty string.
 *  - The `blockId` is trimmed and length-bounded.
 *  - `triggerBlockId` (interactive "on click of") is preserved only when it is a
 *    short non-empty string that does NOT equal the step's own `blockId` (a
 *    self-triggering step is rejected) AND — when `validBlockIds` is supplied —
 *    names a block that EXISTS. A dropped value simply omits the field so the
 *    step falls back to the normal on-click sequence.
 *
 * `validBlockIds`, when supplied, is the set of block ids on the step's slide; it
 * gates the existence check for `triggerBlockId`. Omit it (the default) for
 * existence-agnostic round-trips (e.g. serialize/deserialize where the slide's
 * blocks travel separately) — then only the shape + self-reference rules apply.
 *
 * Pure: never mutates its input.
 */
export function validateStep(
  raw: unknown,
  validBlockIds?: ReadonlySet<string>,
): TimelineStep | null {
  const record = asRecord(raw);
  if (!record) return null;
  const blockIdRaw = record.blockId;
  if (typeof blockIdRaw !== "string") return null;
  const blockId = blockIdRaw.trim();
  if (!blockId || blockId.length > MAX_BLOCK_ID_LENGTH) return null;

  const step: TimelineStep = {
    blockId,
    animation: normalizeTimelineAnimation(record.animation),
    trigger: normalizeTimelineTrigger(record.trigger),
    delayMs: clampDelayMs(record.delayMs),
    durationMs: clampDurationMs(record.durationMs),
  };

  const effectOption = resolvePresentationAnimationEffectOption(
    step.animation,
    record.effectOption,
  );
  if (effectOption?.effectOption !== undefined) {
    step.effectOption = effectOption.effectOption;
  }
  const effectColor = resolvePresentationAnimationEffectColor(
    step.animation,
    record.effectColor,
  );
  if (effectColor?.effectColor !== undefined) {
    step.effectColor = effectColor.effectColor;
  }
  // Effect Options: a malformed key is dropped, never the step.
  const effectSettings = normalizeEffectSettings(record.effectSettings);
  if (effectSettings) step.effectSettings = effectSettings;

  if (
    typeof record.id === "string" &&
    record.id.trim().length > 0 &&
    record.id.length <= MAX_STEP_ID_LENGTH
  ) {
    step.id = record.id;
  }

  // Optional motion path. A malformed value (bad preset name, NaN/huge waypoint,
  // empty `custom` list) normalizes to `null` and is simply OMITTED — the step
  // is still a valid entrance/emphasis/exit step, never dropped. A valid path is
  // clamped into the EMU window by `normalizeMotionPath`.
  if (record.motionPath !== undefined) {
    const motionPath = normalizeMotionPath(record.motionPath);
    if (motionPath) step.motionPath = motionPath;
  }

  // Optional paragraph range of an imported By Paragraph build. Kept exactly
  // as stated (the saved step must equal the imported one for the package to
  // stay the export authority); a malformed value is omitted.
  if (record.paragraphs !== undefined) {
    const paragraphs = normalizeStepParagraphs(record.paragraphs);
    if (paragraphs) step.paragraphs = paragraphs;
  }

  // Optional interactive trigger ("start on click of <block>"). Preserved only
  // when it is a sane reference that:
  //  - is a non-empty, length-bounded string (a hand-edited / stale value can
  //    never bloat the payload or crash present mode);
  //  - is NOT the step's own block (a step can never trigger itself — the block
  //    would have to look clicked to reveal itself, which is nonsensical);
  //  - names a block that EXISTS when the live block-id set is known (a deleted
  //    trigger block would otherwise leave the step permanently unreachable).
  // Any failure simply OMITS the field, so the step falls back to the normal
  // on-click sequence — never dropping the step itself.
  const triggerRaw = record.triggerBlockId;
  if (typeof triggerRaw === "string") {
    const triggerBlockId = triggerRaw.trim();
    if (
      triggerBlockId &&
      triggerBlockId.length <= MAX_TRIGGER_BLOCK_ID_LENGTH &&
      triggerBlockId !== blockId &&
      (validBlockIds === undefined || validBlockIds.has(triggerBlockId))
    ) {
      step.triggerBlockId = triggerBlockId;
    }
  }

  return step;
}

/**
 * Normalize a RAW persisted `animationSequence` into an ordered, render-safe
 * array of {@link TimelineStep}. Non-array input → `[]`. Each element is run
 * through {@link validateStep}; invalid / unsalvageable steps are DROPPED (not
 * defaulted to a phantom step) so the order of the surviving steps is preserved.
 * The array is capped at {@link MAX_TIMELINE_STEPS}.
 *
 * `validBlockIds`, when supplied, is forwarded to {@link validateStep} so a step
 * whose `triggerBlockId` names a now-deleted block has that dangling reference
 * dropped (the step survives, falling back to the normal on-click sequence).
 * Omit it for existence-agnostic round-trips.
 *
 * Pure: returns a fresh array of fresh step objects, never mutating the input.
 */
export function normalizeAnimationSequence(
  raw: unknown,
  validBlockIds?: ReadonlySet<string>,
): TimelineStep[] {
  if (!Array.isArray(raw)) return [];
  const steps: TimelineStep[] = [];
  for (const entry of raw) {
    if (steps.length >= MAX_TIMELINE_STEPS) break;
    const step = validateStep(entry, validBlockIds);
    if (step) steps.push(step);
  }
  return steps;
}

/**
 * One PLAY GROUP: the run of steps the audience reveals with a single advance
 * (a click). The FIRST step in a group is the one whose trigger opened it (an
 * `on-click` step, or — defensively — the very first step regardless of its
 * trigger, since a sequence cannot begin mid-air with `with`/`after-previous`).
 * Subsequent `with-previous` / `after-previous` steps attach to that group and
 * play automatically once it is revealed.
 */
export interface TimelineGroup {
  /** The steps in this group, in sequence order. Always non-empty. */
  steps: TimelineStep[];
}

/**
 * Is `value` a valid interactive trigger reference on a step? True only when the
 * step carries a non-empty `triggerBlockId` string — the single predicate that
 * decides whether a step is TRIGGERED (excluded from the normal click-advance
 * sequence) versus a plain auto-sequence step. Used by every partitioner so the
 * "triggered" test lives in exactly one place.
 */
function isTriggeredStep(step: TimelineStep): boolean {
  return (
    typeof step.triggerBlockId === "string" && step.triggerBlockId.length > 0
  );
}

/**
 * One interactive TRIGGER GROUP: the run of steps that fire together when the
 * audience clicks the block named by `triggerBlockId` in present mode. The group
 * is opened by a step carrying that `triggerBlockId`, and any immediately
 * following `with-previous` / `after-previous` steps chain into it (they play on
 * the SAME click of the trigger block, at their cumulative offsets). `indices`
 * are the flat positions of the group's steps in the ORIGINAL sequence, so a
 * caller can map a group step back to its absolute position.
 */
export interface TriggerGroup {
  /** The block id whose present-mode click fires this group. */
  triggerBlockId: string;
  /** The steps in this group, in sequence order. Always non-empty. */
  steps: TimelineStep[];
  /** The flat index of each step in the original sequence (parallel to steps). */
  indices: number[];
}

/**
 * The full partition of a sequence into the NORMAL click-advance flow versus the
 * interactive TRIGGER groups.
 *
 *  - `autoGroups`: the ordered click-advance groups (each revealed by a stage
 *    click), EXACTLY as {@link resolveTimelineGroups} returns — but with every
 *    triggered step (and its chained followers) removed first. This is the heart
 *    of the highest-risk invariant: a triggered step is NEVER part of a click
 *    group, so a background advance click can never fire it.
 *  - `triggerGroups`: the interactive groups, each keyed by the block whose click
 *    fires it. Multiple groups can share a `triggerBlockId` (two separate
 *    triggered runs on the same block) — clicking that block fires them all.
 *  - `autoIndices` / `triggeredIndices`: the flat positions of every auto-flow vs
 *    triggered step, so the play-state resolver can reason in absolute indices.
 *
 * Pure: references the same step objects; never mutates the input.
 */
export interface TimelinePartition {
  autoGroups: TimelineGroup[];
  triggerGroups: TriggerGroup[];
  autoIndices: number[];
  triggeredIndices: number[];
}

/**
 * Partition a normalized sequence into its auto-flow groups and its interactive
 * trigger groups. The single source of truth for which steps belong to the
 * click-advance sequence and which fire on a trigger block's click.
 *
 * Walk semantics (PowerPoint "Trigger → On Click of" parity):
 *  - A step carrying a `triggerBlockId` OPENS a new trigger group keyed by that
 *    id and becomes the "current" group, REGARDLESS of its own `trigger` field
 *    (a triggered effect is the opener of its own interactive run).
 *  - A non-triggered `with-previous` / `after-previous` step CHAINS into the
 *    current group — which may be an auto group OR a trigger group — so a step
 *    chained after a triggered step inherits the trigger (it fires on the same
 *    block click). At sequence start (no current group) such a step instead
 *    opens the leading auto group (a sequence cannot begin chained to nothing).
 *  - A non-triggered `on-click` step opens a new AUTO group and resets "current"
 *    to the auto flow.
 *
 * Pure: returns fresh groups referencing the SAME step objects; never mutates.
 */
export function partitionTimelineSteps(
  sequence: readonly TimelineStep[],
): TimelinePartition {
  const autoGroups: TimelineGroup[] = [];
  const triggerGroups: TriggerGroup[] = [];
  const autoIndices: number[] = [];
  const triggeredIndices: number[] = [];
  // The group the NEXT chained (with/after-previous) step attaches to. `null`
  // until the first group opens, then either an auto group or a trigger group.
  let current: TimelineGroup | TriggerGroup | null = null;
  let currentIsTrigger = false;

  for (let index = 0; index < sequence.length; index += 1) {
    const step = sequence[index];
    if (isTriggeredStep(step)) {
      // A With/After Previous step on the same trigger goes on in the run its
      // click started, as in PowerPoint's interactive sequence: After Previous
      // waits for the effect before it. Opening a group of its own, it played
      // at its delay from the click, so an exit after a 1.25 s entrance began
      
      if (
        step.trigger !== "on-click" &&
        currentIsTrigger &&
        (current as TriggerGroup).triggerBlockId === step.triggerBlockId
      ) {
        const group = current as TriggerGroup;
        group.steps.push(step);
        group.indices.push(index);
        triggeredIndices.push(index);
        continue;
      }
      // A triggered step otherwise opens its own trigger group (keyed by the
      // block whose click fires it) and becomes the chain anchor for any
      // followers.
      const group: TriggerGroup = {
        triggerBlockId: step.triggerBlockId as string,
        steps: [step],
        indices: [index],
      };
      triggerGroups.push(group);
      triggeredIndices.push(index);
      current = group;
      currentIsTrigger = true;
      continue;
    }
    // Non-triggered step. It chains into the current group only when it is a
    // with/after-previous step AND a current group exists; otherwise it opens a
    // fresh AUTO group (this also covers the leading step, whose with/after is
    // chained to nothing → it must open the first auto group).
    const chains = step.trigger !== "on-click" && current !== null;
    if (chains && currentIsTrigger) {
      const group = current as TriggerGroup;
      group.steps.push(step);
      group.indices.push(index);
      triggeredIndices.push(index);
      continue;
    }
    if (chains) {
      (current as TimelineGroup).steps.push(step);
      autoIndices.push(index);
      continue;
    }
    // Opens a new auto group.
    const group: TimelineGroup = { steps: [step] };
    autoGroups.push(group);
    autoIndices.push(index);
    current = group;
    currentIsTrigger = false;
  }

  return { autoGroups, triggerGroups, autoIndices, triggeredIndices };
}

/**
 * Partition a normalized sequence into ordered {@link TimelineGroup}s by their
 * triggers, modelling PowerPoint's animation-pane semantics:
 *
 *  - An `on-click` step STARTS a new group (the audience must advance to it).
 *  - The VERY FIRST step always starts the first group, EVEN IF its persisted
 *    trigger is `with-previous` / `after-previous` — a sequence cannot begin
 *    chained to a non-existent predecessor, so the leading step is treated as
 *    the group opener. (PowerPoint shows the first effect on the first click.)
 *  - `with-previous` / `after-previous` steps ATTACH to the current group.
 *  - A step carrying an interactive `triggerBlockId` (and anything chained after
 *    it) is EXCLUDED entirely — it is not part of the click-advance flow, it
 *    fires on its trigger block's click. So the click groups count only the
 *    advances the audience makes on the STAGE, never the interactive runs.
 *
 * The number of groups equals the number of click-advances needed to reveal the
 * whole slide's auto-flow. An empty (or fully-triggered) sequence yields no
 * groups.
 *
 * Pure: returns fresh groups referencing the SAME step objects (no copy of the
 * steps themselves); never mutates the input.
 */
export function resolveTimelineGroups(
  sequence: readonly TimelineStep[],
): TimelineGroup[] {
  return partitionTimelineSteps(sequence).autoGroups;
}

























export type StepPlayDurationMs = (step: TimelineStep) => number;
const stepOwnDurationMs: StepPlayDurationMs = (step) => step.durationMs;

export function computeStepStartOffsets(
  group: TimelineGroup,
  /** How long a step plays in full: its own duration unless the host knows
   * better (an effect over a block's letters, `effectPlayDurationMs`). */
  playDurationMs: StepPlayDurationMs = stepOwnDurationMs,
): number[] {
  const offsets: number[] = [];
  let runStart = 0;
  let latestEnd = 0;
  for (let index = 0; index < group.steps.length; index += 1) {
    const step = group.steps[index];
    if (index > 0 && step.trigger !== "with-previous") {
      // After Previous (and, defensively, an on-click that ended up
      // mid-group): a new run once everything before it has played.
      runStart = latestEnd;
    }
    const start = Math.max(0, runStart + step.delayMs);
    offsets.push(start);
    latestEnd = Math.max(latestEnd, start + Math.max(0, playDurationMs(step)));
  }
  return offsets;
}






export function legacyStepStartOffsets(
  group: TimelineGroup,
  playDurationMs: StepPlayDurationMs = stepOwnDurationMs,
): number[] {
  const offsets: number[] = [];
  let prevStart = 0;
  let prevDuration = 0;
  for (let index = 0; index < group.steps.length; index += 1) {
    const step = group.steps[index];
    let start: number;
    if (index === 0) {
      start = step.delayMs;
    } else if (step.trigger === "with-previous") {
      start = prevStart + step.delayMs;
    } else {
      start = prevStart + prevDuration + step.delayMs;
    }
    if (start < 0) start = 0;
    offsets.push(start);
    prevStart = start;
    prevDuration = playDurationMs(step);
  }
  return offsets;
}

/**
 * Triggers and delays that start every step of `group` at `starts` (ms from
 * the group's reveal) in PowerPoint's timing ({@link computeStepStartOffsets}):
 * a With step's delay is from its run's start; an After step whose start
 * comes before everything ahead of it has finished (a delay PowerPoint cannot
 * state, below zero) becomes a With step of the current run.
 */
export function stepsStartingAt<T extends TimelineStep>(
  steps: readonly T[],
  starts: readonly number[],
  playDurationMs: StepPlayDurationMs = stepOwnDurationMs,
): T[] {
  const out: T[] = [];
  let runStart = 0;
  let latestEnd = 0;
  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index];
    const start = Math.max(0, starts[index] ?? 0);
    let trigger = step.trigger;
    let delayMs: number;
    if (index === 0) {
      runStart = 0;
      delayMs = start;
    } else if (
      start < latestEnd ||
      // A With delay past the longest a step states, when After can say it.
      (trigger === "with-previous" && start - runStart <= MAX_STEP_DELAY_MS)
    ) {
      trigger = "with-previous";
      delayMs = start - runStart;
    } else {
      trigger = "after-previous";
      runStart = latestEnd;
      delayMs = start - runStart;
    }
    out.push(
      trigger === step.trigger && delayMs === step.delayMs
        ? step
        : { ...step, trigger, delayMs: Math.max(0, delayMs) },
    );
    latestEnd = Math.max(latestEnd, start + Math.max(0, playDurationMs(step)));
  }
  return out;
}

 
export const POWERPOINT_ANIMATION_TIMING = "powerpoint" as const;







export function migrateLegacyAnimationSequence<T extends TimelineStep>(sequence: readonly T[]): T[] {
  const { autoGroups, triggerGroups, autoIndices } = partitionTimelineSteps(sequence);
  const out = sequence.slice();
  let cursor = 0;
  for (const group of autoGroups) {
    const indices = autoIndices.slice(cursor, cursor + group.steps.length);
    cursor += group.steps.length;
    const steps = indices.map((index) => sequence[index]);
    const moved = stepsStartingAt(steps, legacyStepStartOffsets({ steps }));
    indices.forEach((index, position) => {
      out[index] = moved[position];
    });
  }
  for (const group of triggerGroups) {
    const steps = group.indices.map((index) => sequence[index]);
    const moved = stepsStartingAt(steps, legacyStepStartOffsets({ steps }));
    group.indices.forEach((index, position) => {
      out[index] = moved[position];
    });
  }
  return out;
}

/**
 * One frame in a full-slide animation preview. `revealedGroups` is the count
 * passed to the shared slide renderer; `durationMs` is how long the newest
 * group needs to finish after accounting for with/after-previous offsets.
 */
export interface TimelinePreviewFrame {
  revealedGroups: number;
  durationMs: number;
}

/**
 * Build the auto-flow playback plan used by Ribbon Preview and Pane Play All.
 * Interactive trigger groups are intentionally absent because
 * {@link resolveTimelineGroups} exposes only the stage-advance flow. Groups
 * without a visible effect are skipped without disturbing the reveal count of
 * later groups, so a later frame still makes every preceding group static.
 */
export function resolveTimelinePreviewFrames(
  sequence: readonly TimelineStep[],
): TimelinePreviewFrame[] {
  const frames: TimelinePreviewFrame[] = [];
  for (const [groupIndex, group] of resolveTimelineGroups(sequence).entries()) {
    const offsets = computeStepStartOffsets(group);
    let durationMs = 0;
    let playable = false;
    for (let index = 0; index < group.steps.length; index += 1) {
      const step = group.steps[index];
      if (step.animation === "none" && step.motionPath === undefined) continue;
      playable = true;
      durationMs = Math.max(durationMs, offsets[index] + step.durationMs);
    }
    if (playable) {
      frames.push({
        revealedGroups: groupIndex + 1,
        durationMs,
      });
    }
  }
  return frames;
}

/** Build the same preview plan beginning at one selected timeline step. */
export function resolveTimelinePreviewFramesFrom(
  sequence: readonly TimelineStep[],
  startIndex: number,
): TimelinePreviewFrame[] {
  const safeStart = Number.isInteger(startIndex)
    ? Math.min(sequence.length, Math.max(0, startIndex))
    : 0;
  return resolveTimelinePreviewFrames(sequence.slice(safeStart));
}

/**
 * Build a fresh, fully-defaulted step for a block. Used by the editor when the
 * author adds the block to the timeline. The animation defaults to a visible
 * `fade` (a no-op `none` step would be pointless on the timeline) unless an
 * explicit animation is supplied; trigger/delay/duration take their defaults.
 */
export function createStep(
  blockId: string,
  animation: TimelineAnimation = "fade",
): TimelineStep {
  const safeAnimation = normalizeTimelineAnimation(animation);
  return {
    // Every new occurrence is born with its own identity so the collaboration
    // layer merges a slide's animation list per step: two people editing the
    // timing of two different objects at once both keep their change, instead
    // of the whole list being one last-writer-wins value.
    id: newTimelineStepId(),
    blockId,
    animation:
      safeAnimation === DEFAULT_BLOCK_ANIMATION ? "fade" : safeAnimation,
    trigger: DEFAULT_TIMELINE_TRIGGER,
    delayMs: DEFAULT_STEP_DELAY_MS,
    // A new step starts at PowerPoint's own duration for its effect (Circle
    // 2 s, Dissolve In 0.5 s, Crawl In 5 s, Fade 0.5 s, Spin 2 s). The
    // 380 ms DEFAULT_STEP_DURATION_MS stays only as the reading of a stored
    // step that carries no duration, so old decks play as they always did.
    durationMs: newStepDurationMs(safeAnimation),
  };
}

function newStepDurationMs(animation: TimelineAnimation): number {
  const recipe = presentationAnimationEffect(animation)?.recipeDurationMs;
  if (recipe !== null && recipe !== undefined) return recipe;
  return (
    (HAND_DRAWN_EFFECT_POWERPOINT_DURATION_MS as Readonly<Record<string, number>>)[
      animation
    ] ?? NEW_STEP_POWERPOINT_DURATION_MS
  );
}

/** A fresh, unique step identity (a UUID where the runtime offers one). */
export function newTimelineStepId(): string {
  const runtimeCrypto = (globalThis as { crypto?: { randomUUID?: () => string } })
    .crypto;
  if (typeof runtimeCrypto?.randomUUID === "function") {
    return runtimeCrypto.randomUUID();
  }
  return `step-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 12)}`;
}

/**
 * Give every step that has no identity yet (decks authored before steps were
 * born with an id) a DETERMINISTIC one, and re-key a duplicate id so the list's
 * identities are unique. The id is derived only from persisted facts (`scope`,
 * normally the slide id, the step's block and its position), so two
 * collaborators loading the same stored deck derive exactly the same ids and
 * the collaboration layer can merge the list per step for both of them. The
 * ids then persist through the normal save. A list that already carries
 * unique ids everywhere comes back as an unchanged copy.
 */
export function assignTimelineStepIds(
  sequence: readonly TimelineStep[],
  scope: string,
): TimelineStep[] {
  const used = new Set<string>();
  let changed = false;
  const next = sequence.map((step, index) => {
    if (step.id && !used.has(step.id)) {
      used.add(step.id);
      return step;
    }
    const base = `${scope}:${step.blockId}:${index}`.slice(
      0,
      MAX_STEP_ID_LENGTH - 8,
    );
    let id = base;
    let attempt = 1;
    while (used.has(id)) {
      id = `${base}~${attempt}`;
      attempt += 1;
    }
    used.add(id);
    changed = true;
    return { ...step, id };
  });
  return changed ? next : [...sequence];
}

/** Build one menu-authored step only after the shared capability gate passes. */
export function createAuthorableTimelineStep(
  blockId: string,
  request: TimelineAnimationAuthoringRequest,
): TimelineStep | null {
  if (!timelineAnimationAuthoringCapability(request).enabled) return null;
  if (request.motionPathPreset !== undefined) {
    const motionPath = presetMotionPath(request.motionPathPreset);
    if (!motionPath) return null;
    // Add Animation > Motion Paths is its own behavior. It must not smuggle an
    // Appear entrance into the same click group merely to satisfy createStep's
    // non-noop default.
    const step = createStep(blockId, "appear");
    step.animation = "none";
    step.motionPath = motionPath;
    return step;
  }
  return createStep(blockId, request.animation);
}

/**
 * Move the step at `from` to `to` within a sequence, returning a NEW array (the
 * input is never mutated). Out-of-range indices, a no-op move, or a non-array
 * input return a shallow copy unchanged so the caller can treat the result
 * uniformly. Used by the timeline panel's reorder controls.
 */
export function reorderSteps<T>(
  steps: readonly T[],
  from: number,
  to: number,
): T[] {
  const next = [...steps];
  if (
    from < 0 ||
    from >= next.length ||
    to < 0 ||
    to >= next.length ||
    from === to
  ) {
    return next;
  }
  const [moved] = next.splice(from, 1);
  next.splice(to, 0, moved);
  return next;
}

/** ANI02's exact one-occurrence edits exposed by the Animation Pane. */
export type AnimationSequenceEditOperation =
  "move-earlier" | "move-later" | "remove";

export interface AnimationTimingValue {
  readonly trigger: TimelineTrigger;
  readonly delayMs: number;
  readonly durationMs: number;
}

/**
 * Make detached evidence for one live timeline occurrence. Command payloads
 * carry this copy so an async registry invocation can prove that the row at the
 * requested index is still the same complete step before it mutates anything.
 */
export function copyTimelineStep(step: TimelineStep): TimelineStep {
  return {
    ...step,
    ...(step.effectSettings ? { effectSettings: { ...step.effectSettings } } : {}),
    ...(step.motionPath
      ? {
          motionPath: {
            preset: step.motionPath.preset,
            waypoints: step.motionPath.waypoints.map((point) => ({ ...point })),
          },
        }
      : {}),
  };
}

/** Exact semantic equality for detached occurrence evidence. */
export function timelineStepsEqual(
  left: TimelineStep,
  right: TimelineStep,
): boolean {
  if (
    left.id !== right.id ||
    left.blockId !== right.blockId ||
    left.animation !== right.animation ||
    left.effectOption !== right.effectOption ||
    left.effectColor !== right.effectColor ||
    !effectSettingsEqual(left.effectSettings, right.effectSettings) ||
    left.trigger !== right.trigger ||
    left.delayMs !== right.delayMs ||
    left.durationMs !== right.durationMs ||
    left.triggerBlockId !== right.triggerBlockId ||
    left.paragraphs?.start !== right.paragraphs?.start ||
    left.paragraphs?.end !== right.paragraphs?.end ||
    left.motionPath?.preset !== right.motionPath?.preset
  ) {
    return false;
  }
  const leftWaypoints = left.motionPath?.waypoints ?? [];
  const rightWaypoints = right.motionPath?.waypoints ?? [];
  return (
    leftWaypoints.length === rightWaypoints.length &&
    leftWaypoints.every(
      (point, index) =>
        point.x === rightWaypoints[index]?.x &&
        point.y === rightWaypoints[index]?.y,
    )
  );
}

/** Replace one exact selected occurrence's directional option, or return null. */
export function editTimelineStepEffectOption(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  option: PresentationAnimationEffectOption,
): TimelineStep[] | null {
  const current = steps[index];
  const resolution = resolvePresentationAnimationEffectOption(
    current?.animation,
    option,
  );
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep) ||
    !resolution ||
    resolution.selectedOption === null ||
    current.effectOption === resolution.effectOption
  ) {
    return null;
  }
  // A text built a paragraph at a time takes the direction as one effect,
  // every paragraph's step together (PowerPoint's Effect Options).
  const build = timelineTextBuildRange(steps, index) ?? { first: index, last: index };
  return steps.map((step, stepIndex) => {
    const copy = copyTimelineStep(step);
    if (stepIndex < build.first || stepIndex > build.last) return copy;
    if (resolution.effectOption === undefined) {
      const { effectOption: _effectOption, ...withoutOption } = copy;
      return withoutOption;
    }
    return { ...copy, effectOption: resolution.effectOption };
  });
}

/**
 * Replace the Effect Options (the Effect Options dialog's Effect and Timing
 * tabs) of one exact selected occurrence, or return null when nothing would
 * change. `settings` must already be canonical for the step's effect
 * (`canonicalEffectSettings`); `undefined` returns the effect to the
 * gallery's defaults. A text built a paragraph at a time takes them as one
 * effect, every paragraph's step together, as PowerPoint's dialog applies to
 * the whole build.
 */
export function editTimelineStepEffectSettings(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  settings: PresentationAnimationEffectSettings | undefined,
): TimelineStep[] | null {
  const current = steps[index];
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep) ||
    effectSettingsEqual(current.effectSettings, settings)
  ) {
    return null;
  }
  const build = timelineTextBuildRange(steps, index) ?? { first: index, last: index };
  return steps.map((step, stepIndex) => {
    const copy = copyTimelineStep(step);
    if (stepIndex < build.first || stepIndex > build.last) return copy;
    const { effectSettings: _previous, ...withoutSettings } = copy;
    return settings ? { ...withoutSettings, effectSettings: { ...settings } } : withoutSettings;
  });
}

/**
 * PowerPoint's Effect Options > Sequence for an effect on a text box: the
 * whole box at once (As One Object), every paragraph together (All at Once)
 * or one paragraph per click (By Paragraph).
 */
export type TimelineTextSequence = "as-one-object" | "all-at-once" | "by-paragraph";

/**
 * The steps of the text build the step at `index` belongs to: the adjacent
 * steps of the same effect on the same block that each play on paragraphs.
 * A step on the whole block is a build of one.
 */
export function timelineTextBuildRange(
  steps: readonly TimelineStep[],
  index: number,
): { first: number; last: number } | null {
  const current = steps[index];
  if (!current) return null;
  if (!current.paragraphs) return { first: index, last: index };
  const sameBuild = (step: TimelineStep | undefined) =>
    step !== undefined &&
    step.blockId === current.blockId &&
    step.animation === current.animation &&
    step.paragraphs !== undefined;
  let first = index;
  while (first > 0 && sameBuild(steps[first - 1])) first -= 1;
  let last = index;
  while (last + 1 < steps.length && sameBuild(steps[last + 1])) last += 1;
  return { first, last };
}

/** Which Sequence the step at `index` is played with, or null. */
export function timelineTextSequenceOf(
  steps: readonly TimelineStep[],
  index: number,
): TimelineTextSequence | null {
  const build = timelineTextBuildRange(steps, index);
  if (!build) return null;
  if (!steps[build.first]?.paragraphs) return "as-one-object";
  for (let position = build.first + 1; position <= build.last; position += 1) {
    if (steps[position]?.trigger !== "with-previous") return "by-paragraph";
  }
  return build.last > build.first ? "all-at-once" : "by-paragraph";
}

/**
 * Set the Sequence of the effect at `index` on its text box, or null. By
 * Paragraph and All at Once play the effect once per paragraph that holds
 * words (`paragraphIndexes`), the first in the effect's own place and
 * timing, the rest on a click each or with the first; As One Object plays it
 * once on the whole box. A build with a motion path or an object trigger
 * stays whole.
 */
export function editTimelineStepTextSequence(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  sequence: TimelineTextSequence,
  paragraphIndexes: readonly number[],
): TimelineStep[] | null {
  const current = steps[index];
  const build = timelineTextBuildRange(steps, index);
  if (
    !current ||
    !build ||
    !timelineStepsEqual(current, expectedStep) ||
    timelineTextSequenceOf(steps, index) === sequence
  ) {
    return null;
  }
  const first = copyTimelineStep(steps[build.first]);
  const { paragraphs: _paragraphs, ...whole } = first;
  let replacement: TimelineStep[];
  if (sequence === "as-one-object") {
    replacement = [whole];
  } else {
    if (
      paragraphIndexes.length < 2 ||
      whole.motionPath !== undefined ||
      whole.triggerBlockId !== undefined
    ) {
      return null;
    }
    replacement = paragraphIndexes.map((paragraph, position) => ({
      ...whole,
      ...(position === 0
        ? {}
        : {
            id: newTimelineStepId(),
            trigger:
              sequence === "by-paragraph"
                ? ("on-click" as const)
                : ("with-previous" as const),
            delayMs: 0,
          }),
      paragraphs: { start: paragraph, end: paragraph },
    }));
  }
  const next = [
    ...steps.slice(0, build.first).map(copyTimelineStep),
    ...replacement,
    ...steps.slice(build.last + 1).map(copyTimelineStep),
  ];
  return next.length > MAX_TIMELINE_STEPS ? null : next;
}

/** The paragraphs of a text box that hold words, by index. */
export function timelineTextParagraphIndexes(text: unknown): number[] {
  if (typeof text !== "string") return [];
  return text
    .split("\n")
    .flatMap((line, paragraph) => (line.trim() ? [paragraph] : []));
}

/** Replace one exact Fill Color occurrence's target, or return null. */
export function editTimelineStepEffectColor(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  color: string | null,
): TimelineStep[] | null {
  const current = steps[index];
  const resolution = resolvePresentationAnimationEffectColor(
    current?.animation,
    color,
  );
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep) ||
    !resolution ||
    current.effectColor === resolution.effectColor
  ) {
    return null;
  }
  return steps.map((step, stepIndex) => {
    const copy = copyTimelineStep(step);
    if (stepIndex !== index) return copy;
    if (resolution.effectColor === undefined) {
      const { effectColor: _effectColor, ...withoutColor } = copy;
      return withoutColor;
    }
    return { ...copy, effectColor: resolution.effectColor };
  });
}

/** Bind or clear one exact occurrence's On Click of object trigger. */
export function editTimelineStepObjectTrigger(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  triggerBlockId: string | null,
  validBlockIds: ReadonlySet<string>,
): TimelineStep[] | null {
  const current = steps[index];
  const trigger = triggerBlockId?.trim() || null;
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep) ||
    (trigger !== null &&
      (trigger === current.blockId ||
        trigger.length > MAX_TRIGGER_BLOCK_ID_LENGTH ||
        !validBlockIds.has(trigger))) ||
    (current.triggerBlockId ?? null) === trigger
  ) {
    return null;
  }
  return steps.map((step, stepIndex) => {
    const copy = copyTimelineStep(step);
    if (stepIndex !== index) return copy;
    if (trigger === null) {
      const { triggerBlockId: _triggerBlockId, ...withoutTrigger } = copy;
      return withoutTrigger;
    }
    return { ...copy, triggerBlockId: trigger };
  });
}

/**
 * Apply one Pane operation only when `index` still names the exact occurrence
 * captured at event time. Invalid, stale, boundary and no-op requests return
 * `null`; a valid result is always a detached array of detached steps.
 */
export function editTimelineSequence(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  operation: AnimationSequenceEditOperation,
): TimelineStep[] | null {
  const current = steps[index];
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep)
  ) {
    return null;
  }
  if (operation === "remove") {
    return steps
      .filter((_, stepIndex) => stepIndex !== index)
      .map(copyTimelineStep);
  }
  const target = operation === "move-earlier" ? index - 1 : index + 1;
  if (
    (operation !== "move-earlier" && operation !== "move-later") ||
    target < 0 ||
    target >= steps.length
  ) {
    return null;
  }
  return reorderSteps(steps.map(copyTimelineStep), index, target);
}

/**
 * Replace only ANI03 timing fields on one still-identical live occurrence.
 * Invalid, stale and semantic no-op requests return `null`.
 */
export function editTimelineStepTiming(
  steps: readonly TimelineStep[],
  index: number,
  expectedStep: TimelineStep,
  value: AnimationTimingValue,
): TimelineStep[] | null {
  const current = steps[index];
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    index >= steps.length ||
    !current ||
    !timelineStepsEqual(current, expectedStep) ||
    !isTimelineTrigger(value.trigger) ||
    !isEditableAnimationTimingMs(value.delayMs) ||
    !isEditableAnimationDurationMs(value.durationMs) ||
    (current.trigger === value.trigger &&
      current.delayMs === value.delayMs &&
      current.durationMs === value.durationMs)
  ) {
    return null;
  }
  return steps.map((step, stepIndex) =>
    stepIndex === index
      ? {
          ...copyTimelineStep(step),
          trigger: value.trigger,
          delayMs: value.delayMs,
          durationMs: value.durationMs,
        }
      : copyTimelineStep(step),
  );
}

/**
 * Drop every step that references an animated block id NOT present in
 * `validBlockIds`, AND strip a dangling `triggerBlockId` from any surviving step
 * whose trigger block no longer exists (so the step falls back to the normal
 * on-click sequence rather than becoming permanently unreachable). Called by the
 * editor after a block is deleted so the persisted timeline never dangles.
 *
 * Pure: returns a fresh array; a surviving step keeps its identity unless its
 * dangling trigger had to be stripped (then a fresh copy without it is returned).
 */
export function pruneSequence(
  sequence: readonly TimelineStep[],
  validBlockIds: ReadonlySet<string>,
): TimelineStep[] {
  const next: TimelineStep[] = [];
  for (const step of sequence) {
    // A step whose ANIMATED block is gone is dropped entirely.
    if (!validBlockIds.has(step.blockId)) continue;
    // A surviving step whose TRIGGER block is gone keeps playing but loses the
    // dangling reference (it reverts to the normal on-click flow). A trigger that
    // still points at the step's own block was never valid; strip it too.
    if (
      step.triggerBlockId !== undefined &&
      (!validBlockIds.has(step.triggerBlockId) ||
        step.triggerBlockId === step.blockId)
    ) {
      const copy = { ...step };
      delete copy.triggerBlockId;
      next.push(copy);
      continue;
    }
    next.push(step);
  }
  return next;
}

/**
 * Total number of click-advances a sequence needs in present mode: one per play
 * group (each `on-click` group, plus the implicit leading group). Used to decide
 * whether an audience "advance" reveals the next group on the current slide or
 * moves on to the next slide. An empty sequence needs 0 advances.
 */
export function timelineClickCount(sequence: readonly TimelineStep[]): number {
  return resolveTimelineGroups(sequence).length;
}

/**
 * The resolved present-mode play state of ONE step at a given revealed-group
 * index. `revealedGroups` is how many click groups the audience has advanced
 * through on the current slide (1 = only the first/auto group is showing).
 *
 *  - `visible`: whether the step's block should be shown at all yet. A step in a
 *    group the audience has NOT reached is hidden (PowerPoint keeps un-triggered
 *    effects off-stage). A TRIGGERED step is hidden until its trigger block has
 *    been clicked (its `triggerBlockId` is in `firedTriggers`).
 *  - `animate`: whether the step's keyframes should play now. True only on the
 *    FIRST frame its group becomes revealed (the caller keys the slide on
 *    `revealedGroups` so a newly revealed group remounts + plays once). A step
 *    with NO visible entrance/emphasis/exit (`animation === "none"`) still
 *    animates when it carries a {@link TimelineStep.motionPath} — a pure
 *    motion-path step (the common PowerPoint authoring) must move its block.
 *  - `delayMs`: the cumulative start offset within its group (CSS animation-delay
 *    for `with-previous` / `after-previous` chaining).
 *  - `durationMs`: the step's animation duration.
 *  - `motionPath`: the step's motion path when set (so the renderer can play the
 *    translate keyframes); absent for a plain entrance/emphasis/exit step.
 */
export interface StepPlayState {
  visible: boolean;
  animate: boolean;
  delayMs: number;
  durationMs: number;
  motionPath?: MotionPath;
  /** True after an auto-flow motion-path group has completed. */
  motionPathSettled: boolean;
  /** True after an auto-flow Fill Color emphasis has reached and held its target. */
  colorEffectSettled: boolean;
  /**
   * True after an auto-flow emphasis that changes the object itself (Grow /
   * Shrink's size, Transparency's opacity, Spin's turn) has completed: the
   * object keeps that end state while later groups play (PPT-B024).
   */
  emphasisSettled: boolean;
  /**
   * Effect Options, once the step's group has played and a later click has
   * come: the object is taken away (After animation: Hide After Animation /
   * Hide on Next Mouse Click; an entrance that rewinds), brought back (an
   * exit that rewinds), or dimmed (After animation: a colour).
   */
  afterHidden?: boolean;
  afterShown?: boolean;
  dimmed?: boolean;
}

/** What a step's Effect Options leave once its group has played (auto flow). */
function settledAfterEffect(
  step: TimelineStep,
  settled: boolean,
): Pick<StepPlayState, "afterHidden" | "afterShown" | "dimmed"> {
  const settings = step.effectSettings;
  if (!settings || !settled) return {};
  const category = presentationAnimationEffect(step.animation)?.category;
  const after = settings.afterEffect;
  const rewound = settings.rewind === true;
  return {
    ...(after === "hide" || after === "hide-on-next-click" || (rewound && category === "entrance")
      ? { afterHidden: true }
      : {}),
    ...(rewound && category === "exit" && !after ? { afterShown: true } : {}),
    ...(after === "dim" ? { dimmed: true } : {}),
  };
}

function timelineStepVisibleAtGroup(
  step: TimelineStep,
  groupVisible: boolean,
  groupAnimates: boolean,
): boolean {
  const category = presentationAnimationEffect(step.animation)?.category;
  if (category === "entrance") return groupVisible;
  if (category === "emphasis") return true;
  if (category === "exit") return !groupVisible || groupAnimates;
  return groupVisible;
}

/** Shared empty set so the no-trigger fast path allocates nothing per call. */
const EMPTY_FIRED_TRIGGERS: ReadonlySet<string> = new Set<string>();

/**
 * Resolve, for every step in a sequence, its {@link StepPlayState} at the given
 * `revealedGroups` count. Returns a Map keyed by the step's array INDEX in the
 * sequence (stable + unique even when blockIds repeat across steps). A block can
 * appear in multiple steps; the caller decides how to combine (typically: a
 * block is visible if ANY of its steps is visible).
 *
 * The sequence is first PARTITIONED (see {@link partitionTimelineSteps}) into the
 * normal click-advance AUTO groups and the interactive TRIGGER groups:
 *
 *  - AUTO groups are revealed by `revealedGroups`, clamped to `[0, autoCount]`.
 *    At 0 nothing in the auto flow is visible; at `autoCount` every auto group is
 *    revealed (the fully-played auto flow). The "newest" revealed auto group
 *    (`groupIndex === revealedGroups - 1`) is the only one flagged to `animate`;
 *    the renderer keys the slide subtree on `revealedGroups`, so a group only
 *    replays its keyframes on the frame it first appears.
 *  - TRIGGER groups are revealed by `firedTriggers`: a group is visible + animates
 *    iff its `triggerBlockId` is in the set (the audience has clicked that block).
 *    A triggered step is therefore NEVER affected by `revealedGroups` — a stage
 *    advance can never reveal it, which is the core interactive-trigger invariant.
 *    `firedTriggers` is OPTIONAL and defaults to empty, so a sequence with no
 *    triggers (the common case) behaves byte-for-byte as before.
 *
 * Pure: reads only its inputs.
 */
export function resolveStepPlayStates(
  sequence: readonly TimelineStep[],
  revealedGroups: number,
  firedTriggers?: ReadonlySet<string>,
  /**
   * The revealed groups are an end state being RETURNED to (the show stepped
   * back onto a slide it had fully played): nothing animates, every revealed
   * group is settled, so an exit stays gone instead of replaying (PPT-B036).
   */
  settled = false,
  /** See {@link computeStepStartOffsets}: how long each step plays in full. */
  playDurationMs: StepPlayDurationMs = stepOwnDurationMs,
): Map<number, StepPlayState> {
  const { autoGroups, triggerGroups, autoIndices } =
    partitionTimelineSteps(sequence);
  const states = new Map<number, StepPlayState>();

  // ── AUTO flow ──────────────────────────────────────────────────────────────
  // `autoIndices` is the flat, in-order list of every auto step's absolute
  // position. Auto groups preserve sequence order, so consuming it sequentially
  // as we iterate the groups maps each group step back to its absolute index
  // (rather than carrying the index on TimelineGroup, whose public shape stays a
  // plain `{ steps }` kept stable for callers/tests).
  const revealed = Math.max(0, Math.min(revealedGroups, autoGroups.length));
  let autoCursor = 0;
  for (let groupIndex = 0; groupIndex < autoGroups.length; groupIndex += 1) {
    const group = autoGroups[groupIndex];
    const offsets = computeStepStartOffsets(group, playDurationMs);
    const groupVisible = groupIndex < revealed;
    const groupAnimates = !settled && groupIndex === revealed - 1;
    for (let i = 0; i < group.steps.length; i += 1) {
      const step = group.steps[i];
      const flatIndex = autoIndices[autoCursor];
      autoCursor += 1;
      const hasEffect =
        step.animation !== "none" || step.motionPath !== undefined;
      // Rewind when done playing: nothing of the effect is held afterwards.
      const holds = step.effectSettings?.rewind !== true;
      states.set(flatIndex, {
        visible: timelineStepVisibleAtGroup(step, groupVisible, groupAnimates),
        animate: groupVisible && groupAnimates && hasEffect,
        delayMs: offsets[i],
        durationMs: step.durationMs,
        motionPath: step.motionPath,
        motionPathSettled:
          holds && groupVisible && !groupAnimates && step.motionPath !== undefined,
        colorEffectSettled:
          holds &&
          groupVisible &&
          !groupAnimates &&
          presentationAnimationEffect(step.animation)?.nativeBehavior === "color",
        emphasisSettled:
          holds &&
          groupVisible &&
          !groupAnimates &&
          presentationAnimationEffect(step.animation)?.category === "emphasis" &&
          presentationAnimationEffect(step.animation)?.nativeBehavior !== "color",
        ...settledAfterEffect(step, groupVisible && !groupAnimates),
      });
    }
  }

  // ── TRIGGER groups ─────────────────────────────────────────────────────────
  // Each group is gated solely on whether its trigger block has fired; it is
  // entirely independent of the stage advance count.
  const fired = firedTriggers ?? EMPTY_FIRED_TRIGGERS;
  for (const group of triggerGroups) {
    const groupVisible = fired.has(group.triggerBlockId);
    // The group always (re)animates on the mount where its trigger is fired; the
    // caller keys the present subtree so a fired group plays once. A group whose
    // trigger has not fired is fully hidden.
    const offsets = computeStepStartOffsets({ steps: group.steps }, playDurationMs);
    for (let i = 0; i < group.steps.length; i += 1) {
      const step = group.steps[i];
      const flatIndex = group.indices[i];
      const hasEffect =
        step.animation !== "none" || step.motionPath !== undefined;
      states.set(flatIndex, {
        visible: timelineStepVisibleAtGroup(step, groupVisible, groupVisible),
        animate: groupVisible && hasEffect,
        delayMs: offsets[i],
        durationMs: step.durationMs,
        motionPath: step.motionPath,
        motionPathSettled: false,
        colorEffectSettled: false,
        emphasisSettled: false,
      });
    }
  }

  return states;
}

/** The allowed entrance animations a timeline step may use (re-export). */
export { BLOCK_ANIMATIONS };
