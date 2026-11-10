/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-slide present-mode transitions (PowerPoint-parity subset).
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * control, and the present-mode overlay so the allowed set lives in exactly one
 * place. Keep this module dependency-free (no React / no DOM) so it can be unit
 * tested in isolation and imported from the zod schema.
 */

/**
 * The supported transitions. `none` is the implicit default (no animation).
 *
 * The first four (`none`, `fade`, `slide-left`, `slide-up`) are the original
 * set and MUST stay in this order for back-compat. Everything after them is a
 * purely additive PowerPoint-parity expansion (directional slides, `zoom`,
 * `dissolve`, the `push-*` family, and native orientation effects). Each
 * non-`none` value has a matching
 * `.presentation-present__slide--enter-<value>` rule + `@keyframes` in
 * `styles/native-page-window.css`; adding a value here without that rule would
 * render as a no-op, so keep the two in lock-step.
 *
 * `morph`, `cut`, and `random` are the EXCEPTIONS to the one-class model. Morph is not
 * a slide-level entrance keyframe but a per-block FLIP tween driven in present
 * mode (see
 * `lib/presentation-morph.ts` + the morph layer in `PresentationWindow`). It has
 * no `.presentation-present__slide--enter-morph` rule on purpose — when a morph
 * cannot be played (backward nav, no previous slide, etc.) the slide simply
 * renders statically with NO entrance animation. Cut is an instantaneous slide
 * replacement and likewise has no fabricated keyframe. Random selects one of
 * the bounded existing renderers for each playback without persisting that
 * runtime choice. {@link slideTransitionClassName} therefore returns `null`
 * only for None, Morph, and Cut.
 */
export const SLIDE_TRANSITIONS = [
  "none",
  "fade",
  "slide-left",
  "slide-up",
  "slide-right",
  "slide-down",
  "zoom",
  "dissolve",
  "push-left",
  "push-right",
  "push-up",
  "push-down",
  "wipe-right",
  "wipe-left",
  "wipe-down",
  "wipe-up",
  "morph",
  // Exact PowerPoint 16.86 basic-gallery identities. These are additive: the
  // historical directional values above remain valid for persisted decks, but
  // a native Push/Wipe selection must no longer collapse back to slide-up /
  // wipe-right on import.
  "push",
  "wipe",
  "cut",
  "shape",
  "split",
  "random-bars",
  "checkerboard",
  "blinds",
  "comb",
  "fall-over",
  "drape",
  "curtains",
  "wind",
  "prestige",
  "fracture",
  "crush",
  "peel-off",
  "page-curl",
  "airplane",
  "origami",
  "switch",
  "flip",
  "gallery",
  "ferris-wheel",
  "conveyor",
  "uncover",
  "cover",
  "doors",
  "window",
  "vortex",
  "pan",
  "flash",
  "honeycomb",
  "reveal",
  "ripple",
  "glitter",
  "shred",
  "fly-through",
  "clock",
  "random",
  "cube",
  "box",
  "rotate",
  "orbit",
] as const;

export type SlideTransition = (typeof SLIDE_TRANSITIONS)[number];

/** The largest sound held inline in the page (a built-in sound, one a deck
 * brought); a chosen file is stored beside the deck instead, with no size of
 * its own past the deck's media total (lib/document-size-limits). */
export const MAX_PRESENTATION_TRANSITION_SOUND_BYTES = 1_450_000;
export const MAX_PRESENTATION_TRANSITION_SOUND_NAME_CHARS = 160;

/** A WAV held inline in the page as a data: URL. */
export interface InlinePresentationTransitionSound {
  readonly dataUrl: string;
  readonly mimeType: "audio/wav";
  readonly name: string;
  readonly sizeBytes: number;
}

/** A WAV stored beside the deck: a DataItem of its notebook. */
export interface StoredPresentationTransitionSound {
  readonly sourceKind: "asset";
  readonly assetId: string;
  readonly mimeType: "audio/wav";
  readonly name: string;
  readonly sizeBytes: number;
}

/**
 * An imported deck's WAV, kept where PowerPoint keeps it -- a part of the
 * page's own package -- and named by its digest (readings row 560): played
 * through `presentation-native-media/snd-<fingerprint>`, never copied into
 * the page. The part, the slide relationship that points at it and its
 * size are the receipt the server checks before it serves a byte.
 */
export interface PackagePresentationTransitionSound {
  readonly sourceKind: "package";
  readonly fingerprint: string;
  readonly part: string;
  readonly relationshipsPart: string;
  readonly relationshipId: string;
  readonly mimeType: "audio/wav";
  readonly name: string;
  readonly sizeBytes: number;
}

export type PresentationTransitionSound =
  | InlinePresentationTransitionSound
  | StoredPresentationTransitionSound
  | PackagePresentationTransitionSound;

const PACKAGE_SOUND_KEYS = [
  "sourceKind",
  "fingerprint",
  "part",
  "relationshipsPart",
  "relationshipId",
  "mimeType",
  "name",
  "sizeBytes",
];

function packageTransitionSound(
  record: Record<string, unknown>,
): PackagePresentationTransitionSound | null {
  const name = transitionSoundName(record.name);
  const { fingerprint, part, relationshipsPart, relationshipId, sizeBytes } = record;
  if (
    Object.keys(record).some((key) => !PACKAGE_SOUND_KEYS.includes(key)) ||
    record.mimeType !== "audio/wav" ||
    !name ||
    typeof fingerprint !== "string" ||
    !/^[0-9a-f]{64}$/.test(fingerprint) ||
    typeof part !== "string" ||
    !/^ppt\/media\/[A-Za-z0-9][A-Za-z0-9._-]{0,199}$/.test(part) ||
    typeof relationshipsPart !== "string" ||
    !/^ppt\/slides\/_rels\/slide[1-9][0-9]*\.xml\.rels$/.test(relationshipsPart) ||
    typeof relationshipId !== "string" ||
    !/^rId[1-9][0-9]*$/.test(relationshipId) ||
    !Number.isSafeInteger(sizeBytes) ||
    (sizeBytes as number) < 12 ||
    (sizeBytes as number) > MAX_PRESENTATION_TRANSITION_SOUND_BYTES
  ) {
    return null;
  }
  return Object.freeze({
    sourceKind: "package",
    fingerprint,
    part,
    relationshipsPart,
    relationshipId,
    mimeType: "audio/wav",
    name,
    sizeBytes: sizeBytes as number,
  });
}

/** Where a package sound plays from: the page's own package part, by digest. */
export function packageTransitionSoundUrl(apiBase: string, pageId: string, sound: PackagePresentationTransitionSound): string {
  return `${apiBase}/api/v1/pages/${encodeURIComponent(pageId)}/presentation-native-media/snd-${sound.fingerprint}`;
}

const TRANSITION_SOUND_DATA_URL_RE =
  /^data:audio\/wav;base64,([A-Za-z0-9+/]+={0,2})$/;

function transitionSoundName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const name = value.trim();
  return name &&
    name.length <= MAX_PRESENTATION_TRANSITION_SOUND_NAME_CHARS &&
    !/[\u0000-\u001f\u007f-\u009f]/.test(name)
    ? name
    : null;
}

function storedTransitionSound(
  record: Record<string, unknown>,
): StoredPresentationTransitionSound | null {
  const name = transitionSoundName(record.name);
  const assetId = record.assetId;
  const sizeBytes = record.sizeBytes;
  if (
    Object.keys(record).some(
      (key) =>
        !["sourceKind", "assetId", "mimeType", "name", "sizeBytes"].includes(key),
    ) ||
    record.mimeType !== "audio/wav" ||
    !name ||
    typeof assetId !== "string" ||
    !assetId ||
    assetId !== assetId.trim() ||
    assetId.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(assetId) ||
    !Number.isSafeInteger(sizeBytes) ||
    (sizeBytes as number) < 12
  ) {
    return null;
  }
  return Object.freeze({
    sourceKind: "asset",
    assetId,
    mimeType: "audio/wav",
    name,
    sizeBytes: sizeBytes as number,
  });
}

/** Strict embedded-WAV model used by Transition > Sound and native PPTX. */
export function normalizePresentationTransitionSound(
  value: unknown,
): PresentationTransitionSound | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (record.sourceKind === "asset") return storedTransitionSound(record);
  if (record.sourceKind === "package") return packageTransitionSound(record);
  if (
    Object.keys(record).some(
      (key) => !["dataUrl", "mimeType", "name", "sizeBytes"].includes(key),
    ) ||
    record.mimeType !== "audio/wav" ||
    typeof record.dataUrl !== "string" ||
    !Number.isInteger(record.sizeBytes)
  ) {
    return null;
  }
  const name = transitionSoundName(record.name);
  const sizeBytes = record.sizeBytes as number;
  const match = TRANSITION_SOUND_DATA_URL_RE.exec(record.dataUrl);
  if (
    !name ||
    !match ||
    sizeBytes < 12 ||
    sizeBytes > MAX_PRESENTATION_TRANSITION_SOUND_BYTES ||
    typeof globalThis.atob !== "function"
  ) {
    return null;
  }
  try {
    const decoded = globalThis.atob(match[1]);
    if (
      decoded.length !== sizeBytes ||
      decoded.slice(0, 4) !== "RIFF" ||
      decoded.slice(8, 12) !== "WAVE"
    ) {
      return null;
    }
  } catch {
    return null;
  }
  return Object.freeze({
    dataUrl: record.dataUrl,
    mimeType: "audio/wav",
    name,
    sizeBytes,
  });
}

export function samePresentationTransitionSound(
  left: unknown,
  right: unknown,
): boolean {
  const a = normalizePresentationTransitionSound(left);
  const b = normalizePresentationTransitionSound(right);
  if (a === null || b === null) return a === b;
  const aStored = "assetId" in a ? a.assetId : null;
  const bStored = "assetId" in b ? b.assetId : null;
  const aInline = "dataUrl" in a ? a.dataUrl : null;
  const bInline = "dataUrl" in b ? b.dataUrl : null;
  const aPackage = "fingerprint" in a ? a.fingerprint : null;
  const bPackage = "fingerprint" in b ? b.fingerprint : null;
  return (
    aStored === bStored &&
    aInline === bInline &&
    aPackage === bPackage &&
    a.name === b.name &&
    a.sizeBytes === b.sizeBytes
  );
}

/** The only transition choices authored by PPT.WP.TRN01.BASIC-EFFECT. */
export const BASIC_SLIDE_TRANSITIONS = [
  "none",
  "fade",
  "push",
  "wipe",
  "dissolve",
  "cut",
  "shape",
  "zoom",
  "split",
  "random-bars",
  "checkerboard",
  "blinds",
  "comb",
  "fall-over",
  "drape",
  "curtains",
  "wind",
  "prestige",
  "fracture",
  "crush",
  "peel-off",
  "page-curl",
  "airplane",
  "origami",
  "switch",
  "flip",
  "gallery",
  "ferris-wheel",
  "conveyor",
  "uncover",
  "cover",
  "doors",
  "window",
  "vortex",
  "pan",
  "flash",
  "honeycomb",
  "reveal",
  "ripple",
  "glitter",
  "shred",
  "fly-through",
  "clock",
  "random",
  "cube",
  "box",
  "rotate",
  "orbit",
] as const;
export type BasicSlideTransition = (typeof BASIC_SLIDE_TRANSITIONS)[number];
const BASIC_SLIDE_TRANSITION_SET: ReadonlySet<string> = new Set(
  BASIC_SLIDE_TRANSITIONS,
);

/**
 * Office 2013 `p15:prstTrans` presets whose Effect Options are Left and Right.
 *
 * `invX` mirrors the preset's X axis rather than naming a second effect
 * ([MS-PPTX] CT_PresetTransition), so each preset has one side as written and
 * the other with `invX="1"`. PowerPoint 16.86 reads a bare
 * fallOver/drape/peelOff/pageCurlSingle back as "… left" and a bare
 * wind/airplane/origami as "… right"; the side is the direction the outgoing
 * slide travels. The remaining four presets (Curtains, Prestige, Fracture,
 * Crush) have no side at all.
 */
export const SLIDE_TRANSITION_P15_LEFT_RIGHT_KINDS = [
  "fall-over",
  "drape",
  "peel-off",
  "page-curl",
  "wind",
  "airplane",
  "origami",
] as const;
export type SlideTransitionP15LeftRightKind =
  (typeof SLIDE_TRANSITION_P15_LEFT_RIGHT_KINDS)[number];
const SLIDE_TRANSITION_P15_LEFT_RIGHT_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_P15_LEFT_RIGHT_KINDS,
);

/** Office 2013 `p15:prstTrans` effects with one authenticated preset identity. */
export const SLIDE_TRANSITION_P15_PRESET_KINDS = [
  "fall-over",
  "drape",
  "curtains",
  "wind",
  "prestige",
  "fracture",
  "crush",
  "peel-off",
  "page-curl",
  "airplane",
  "origami",
] as const;
export type SlideTransitionP15PresetKind =
  (typeof SLIDE_TRANSITION_P15_PRESET_KINDS)[number];
const SLIDE_TRANSITION_P15_PRESET_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_P15_PRESET_KINDS,
);

/** Office 2010 `p14` effects whose native option is exactly left or right. */
export const SLIDE_TRANSITION_P14_LEFT_RIGHT_KINDS = [
  "switch",
  "flip",
  "gallery",
  "ferris-wheel",
  "conveyor",
] as const;
export type SlideTransitionP14LeftRightKind =
  (typeof SLIDE_TRANSITION_P14_LEFT_RIGHT_KINDS)[number];
const SLIDE_TRANSITION_P14_LEFT_RIGHT_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_P14_LEFT_RIGHT_KINDS,
);

/** Office 2010 `p14` effects sharing native horizontal/vertical orientation. */
export const SLIDE_TRANSITION_P14_ORIENTATION_KINDS = [
  "doors",
  "window",
] as const;
export type SlideTransitionP14OrientationKind =
  (typeof SLIDE_TRANSITION_P14_ORIENTATION_KINDS)[number];
const SLIDE_TRANSITION_P14_ORIENTATION_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_P14_ORIENTATION_KINDS,
);

/** Office 2010 `p14` effects sharing four-way side-direction semantics. */
export const SLIDE_TRANSITION_P14_SIDE_DIRECTION_KINDS = [
  "vortex",
  "pan",
  "cube",
  "box",
  "rotate",
  "orbit",
] as const;
export type SlideTransitionP14SideDirectionKind =
  (typeof SLIDE_TRANSITION_P14_SIDE_DIRECTION_KINDS)[number];
const SLIDE_TRANSITION_P14_SIDE_DIRECTION_KIND_SET: ReadonlySet<string> =
  new Set(SLIDE_TRANSITION_P14_SIDE_DIRECTION_KINDS);

/** Office 2010 `p14` effects whose native leaf carries no options. */
export const SLIDE_TRANSITION_P14_EMPTY_KINDS = [
  "flash",
  "honeycomb",
] as const;
export type SlideTransitionP14EmptyKind =
  (typeof SLIDE_TRANSITION_P14_EMPTY_KINDS)[number];
const SLIDE_TRANSITION_P14_EMPTY_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_P14_EMPTY_KINDS,
);

/** Office 2010 effects with independently typed p14 Effect Options tuples. */
export const SLIDE_TRANSITION_P14_TYPED_OPTIONS_KINDS = [
  "reveal",
  "ripple",
  "glitter",
  "shred",
  "fly-through",
] as const;
export type SlideTransitionP14TypedOptionsKind =
  (typeof SLIDE_TRANSITION_P14_TYPED_OPTIONS_KINDS)[number];
const SLIDE_TRANSITION_P14_TYPED_OPTIONS_KIND_SET: ReadonlySet<string> =
  new Set(SLIDE_TRANSITION_P14_TYPED_OPTIONS_KINDS);

/** Native 2006 `p:cover` / `p:pull` eight-direction effects. */
export const SLIDE_TRANSITION_EIGHT_DIRECTION_KINDS = [
  "uncover",
  "cover",
] as const;
export type SlideTransitionEightDirectionKind =
  (typeof SLIDE_TRANSITION_EIGHT_DIRECTION_KINDS)[number];
const SLIDE_TRANSITION_EIGHT_DIRECTION_KIND_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_EIGHT_DIRECTION_KINDS,
);

/** Native `p:push/@dir` and `p:wipe/@dir` side-direction identities. */
export const SLIDE_TRANSITION_DIRECTIONS = [
  "left",
  "right",
  "up",
  "down",
] as const;
export type SlideTransitionSideDirection =
  (typeof SLIDE_TRANSITION_DIRECTIONS)[number];
/**
 * Push's Effect Options in PowerPoint's own order -- From Bottom, From Left,
 * From Right, From Top. `p:push/@dir` names the way the slides MOVE: PowerPoint
 * 16.86 reads `dir="u"` back as "entry effect push up" (the new slide rises
 * from the bottom) and `dir="r"` as "push right" (it comes in from the left),
 * and its own save writes Push Left as the bare `<p:push/>`.
 */
export const SLIDE_TRANSITION_PUSH_OPTIONS = ["up", "right", "left", "down"] as const;

/** PowerPoint's name for each Push option: the side the new slide enters from. */
export const SLIDE_TRANSITION_PUSH_OPTION_FROM = {
  up: "bottom",
  right: "left",
  left: "right",
  down: "top",
} as const satisfies Record<(typeof SLIDE_TRANSITION_PUSH_OPTIONS)[number], string>;
export const SLIDE_TRANSITION_LEFT_RIGHT_OPTIONS = ["left", "right"] as const;
export const SLIDE_TRANSITION_CORNER_OPTIONS = [
  "left-down",
  "left-up",
  "right-down",
  "right-up",
] as const;
export type SlideTransitionCornerOption =
  (typeof SLIDE_TRANSITION_CORNER_OPTIONS)[number];
export const SLIDE_TRANSITION_EIGHT_DIRECTION_OPTIONS = [
  ...SLIDE_TRANSITION_DIRECTIONS,
  ...SLIDE_TRANSITION_CORNER_OPTIONS,
] as const;
export const SLIDE_TRANSITION_SHAPE_OPTIONS = [
  "circle",
  "diamond",
  "plus",
] as const;
export type SlideTransitionShapeOption =
  (typeof SLIDE_TRANSITION_SHAPE_OPTIONS)[number];
/**
 * Zoom's In / Out Effect Options. `p:CT_InOutTransition/@dir` defaults to
 * "out" (ISO/IEC 29500-4 A.3), which is how PowerPoint reads a bare Zoom back.
 */
export const SLIDE_TRANSITION_ZOOM_OPTIONS = ["in", "out"] as const;
export type SlideTransitionZoomOption =
  (typeof SLIDE_TRANSITION_ZOOM_OPTIONS)[number];
/** The side a Zoom with no `dir` at all plays. */
export const SLIDE_TRANSITION_ZOOM_DEFAULT_OPTION: SlideTransitionZoomOption =
  "out";
/**
 * The Office 2010 element a Zoom may be spelled with instead of `p:zoom`.
 *
 * `p14:warp` carries the same complex type and the same In / Out options, and
 * it is what PowerPoint's own Zoom gallery entry writes. A slide that came in
 * with it keeps it, so editing the Zoom's direction or timing never hands the
 * deck back to PowerPoint spelled as the other element.
 */
export const SLIDE_TRANSITION_ZOOM_WARP_ELEMENT = "p14:warp" as const;
export type SlideTransitionZoomElement =
  typeof SLIDE_TRANSITION_ZOOM_WARP_ELEMENT;

/**
 * The element a Zoom this editor authors is written with.
 *
 * PowerPoint's own gallery Zoom is `p14:warp` (PowerPoint 16.86 writes it
 * inside an `mc:AlternateContent` carrier with a `p:fade` Fallback), so a
 * newly picked Zoom — or another transition switched to Zoom — is spelled the
 * way PowerPoint spells it. A slide that already carries the 2006 `p:zoom`
 * keeps that element; only leaving Zoom for another effect ends it.
 */
export const SLIDE_TRANSITION_AUTHORED_ZOOM_ELEMENT: SlideTransitionZoomElement =
  SLIDE_TRANSITION_ZOOM_WARP_ELEMENT;

/** Exact guard for the one non-default element a Zoom may name. */
export function isSlideTransitionZoomElement(
  value: unknown,
): value is SlideTransitionZoomElement {
  return value === SLIDE_TRANSITION_ZOOM_WARP_ELEMENT;
}
/** Exact `p:split/@orient + @dir` Effect Options identities. */
export const SLIDE_TRANSITION_SPLIT_OPTIONS = [
  "horizontal-out",
  "horizontal-in",
  "vertical-out",
  "vertical-in",
] as const;
export type SlideTransitionSplitOption =
  (typeof SLIDE_TRANSITION_SPLIT_OPTIONS)[number];
/** Native orientation-transition `@dir` identities, including `p:comb`. */
export const SLIDE_TRANSITION_ORIENTATION_OPTIONS = [
  "horizontal",
  "vertical",
] as const;
export type SlideTransitionOrientationOption =
  (typeof SLIDE_TRANSITION_ORIENTATION_OPTIONS)[number];
/** `p14:reveal/@dir + @thruBlk` composite identities. */
export const SLIDE_TRANSITION_REVEAL_OPTIONS = [
  "left",
  "right",
  "left-through-black",
  "right-through-black",
] as const;
export type SlideTransitionRevealOption =
  (typeof SLIDE_TRANSITION_REVEAL_OPTIONS)[number];
/** `p14:ripple/@dir` corner-or-center identities. */
export const SLIDE_TRANSITION_RIPPLE_OPTIONS = [
  "center",
  "left-up",
  "right-up",
  "left-down",
  "right-down",
] as const;
export type SlideTransitionRippleOption =
  (typeof SLIDE_TRANSITION_RIPPLE_OPTIONS)[number];
/** `p14:glitter/@dir + @pattern` composite identities. */
export const SLIDE_TRANSITION_GLITTER_OPTIONS = [
  "left-diamond",
  "right-diamond",
  "up-diamond",
  "down-diamond",
  "left-hexagon",
  "right-hexagon",
  "up-hexagon",
  "down-hexagon",
] as const;
export type SlideTransitionGlitterOption =
  (typeof SLIDE_TRANSITION_GLITTER_OPTIONS)[number];
/** `p14:shred/@pattern + @dir` composite identities. */
export const SLIDE_TRANSITION_SHRED_OPTIONS = [
  "strips-in",
  "strips-out",
  "rectangles-in",
  "rectangles-out",
] as const;
export type SlideTransitionShredOption =
  (typeof SLIDE_TRANSITION_SHRED_OPTIONS)[number];
/** `p14:flythrough/@dir + @hasBounce` composite identities. */
export const SLIDE_TRANSITION_FLY_THROUGH_OPTIONS = [
  "in",
  "out",
  "in-bounce",
  "out-bounce",
] as const;
export type SlideTransitionFlyThroughOption =
  (typeof SLIDE_TRANSITION_FLY_THROUGH_OPTIONS)[number];
/** Native `p:wheel/@spokes` identities exposed by PowerPoint Effect Options. */
export const SLIDE_TRANSITION_CLOCK_OPTIONS = [
  "one-spoke",
  "two-spokes",
  "three-spokes",
  "four-spokes",
  "eight-spokes",
] as const;
export type SlideTransitionClockOption =
  (typeof SLIDE_TRANSITION_CLOCK_OPTIONS)[number];
export type SlideTransitionP14TypedOption =
  | SlideTransitionRevealOption
  | SlideTransitionRippleOption
  | SlideTransitionGlitterOption
  | SlideTransitionShredOption
  | SlideTransitionFlyThroughOption;
const SLIDE_TRANSITION_P14_TYPED_OPTION_VALUES = [
  "left-through-black",
  "right-through-black",
  "center",
  ...SLIDE_TRANSITION_GLITTER_OPTIONS,
  ...SLIDE_TRANSITION_SHRED_OPTIONS,
  "in-bounce",
  "out-bounce",
] as const;
/** Morph's Effect Options (`p159:morph/@option` byObject / byWord / byChar). */
export const SLIDE_TRANSITION_MORPH_OPTIONS = ["objects", "words", "characters"] as const;
export type SlideTransitionMorphOption = (typeof SLIDE_TRANSITION_MORPH_OPTIONS)[number];
export const SLIDE_TRANSITION_EFFECT_OPTIONS = [
  ...SLIDE_TRANSITION_MORPH_OPTIONS,
  ...SLIDE_TRANSITION_DIRECTIONS,
  ...SLIDE_TRANSITION_CORNER_OPTIONS,
  ...SLIDE_TRANSITION_SHAPE_OPTIONS,
  ...SLIDE_TRANSITION_ZOOM_OPTIONS,
  ...SLIDE_TRANSITION_SPLIT_OPTIONS,
  ...SLIDE_TRANSITION_ORIENTATION_OPTIONS,
  ...SLIDE_TRANSITION_CLOCK_OPTIONS,
  ...SLIDE_TRANSITION_P14_TYPED_OPTION_VALUES,
] as const;
export type SlideTransitionDirection =
  | SlideTransitionSideDirection
  | SlideTransitionCornerOption
  | SlideTransitionShapeOption
  | SlideTransitionZoomOption
  | SlideTransitionSplitOption
  | SlideTransitionOrientationOption
  | SlideTransitionClockOption
  | SlideTransitionP14TypedOption
  | SlideTransitionMorphOption;
const SLIDE_TRANSITION_EFFECT_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_EFFECT_OPTIONS,
);
const SLIDE_TRANSITION_SIDE_DIRECTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_DIRECTIONS,
);
const SLIDE_TRANSITION_CORNER_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_CORNER_OPTIONS,
);
const SLIDE_TRANSITION_SHAPE_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_SHAPE_OPTIONS,
);
const SLIDE_TRANSITION_ZOOM_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_ZOOM_OPTIONS,
);
const SLIDE_TRANSITION_SPLIT_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_SPLIT_OPTIONS,
);
const SLIDE_TRANSITION_ORIENTATION_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_ORIENTATION_OPTIONS,
);
const SLIDE_TRANSITION_CLOCK_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_CLOCK_OPTIONS,
);
const SLIDE_TRANSITION_REVEAL_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_REVEAL_OPTIONS,
);
const SLIDE_TRANSITION_RIPPLE_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_RIPPLE_OPTIONS,
);
const SLIDE_TRANSITION_GLITTER_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_GLITTER_OPTIONS,
);
const SLIDE_TRANSITION_SHRED_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_SHRED_OPTIONS,
);
const SLIDE_TRANSITION_FLY_THROUGH_OPTION_SET: ReadonlySet<string> = new Set(
  SLIDE_TRANSITION_FLY_THROUGH_OPTIONS,
);
const DIRECTIONAL_SLIDE_TRANSITION_DEFAULTS = {
  // A Morph saved before its Effect Options names none: by object.
  morph: "objects",
  cube: "left",
  box: "left",
  rotate: "left",
  orbit: "left",
  // Before these presets carried Effect Options the writer owned exactly one
  // identity each, and every one of them is the RIGHT variant (`invX="1"` on
  // the four left-based presets, the bare preset on the three right-based
  // ones). A slide persisted then names no direction, so right is what it
  // still means — anything else would mirror it on the next export.
  "fall-over": "right",
  drape: "right",
  "peel-off": "right",
  "page-curl": "right",
  wind: "right",
  airplane: "right",
  origami: "right",
  push: "up",
  wipe: "right",
  shape: "circle",
  zoom: "in",
  split: "horizontal-out",
  "random-bars": "horizontal",
  checkerboard: "horizontal",
  blinds: "horizontal",
  comb: "horizontal",
  clock: "one-spoke",
  switch: "right",
  flip: "right",
  gallery: "right",
  "ferris-wheel": "right",
  conveyor: "right",
  uncover: "left",
  cover: "left",
  doors: "horizontal",
  window: "horizontal",
  vortex: "left",
  pan: "left",
  reveal: "left",
  ripple: "center",
  glitter: "left-diamond",
  shred: "strips-in",
  "fly-through": "in",
} as const satisfies Readonly<
  Record<
    | "push"
    | "wipe"
    | "shape"
    | "zoom"
    | "split"
    | "random-bars"
    | "checkerboard"
    | "blinds"
    | "comb"
    | "clock"
    | "morph"
    | SlideTransitionP14LeftRightKind
    | SlideTransitionP14OrientationKind
    | SlideTransitionP14SideDirectionKind
    | SlideTransitionP14TypedOptionsKind
    | SlideTransitionP15LeftRightKind
    | SlideTransitionEightDirectionKind,
    SlideTransitionDirection
  >
>;

/**
 * Every gallery value with a real writer. Morph deliberately stays outside
 * {@link BASIC_SLIDE_TRANSITIONS}:
 * imported p159 AlternateContent is preserve-only today, while local/scratch
 * decks can still author the exact serializer-supported Morph value.
 */
export const AUTHORABLE_SLIDE_TRANSITIONS = [
  "none",
  "morph",
  ...BASIC_SLIDE_TRANSITIONS.filter((kind) => kind !== "none"),
] as const satisfies readonly SlideTransition[];
export type AuthorableSlideTransition =
  (typeof AUTHORABLE_SLIDE_TRANSITIONS)[number];
const AUTHORABLE_SLIDE_TRANSITION_SET: ReadonlySet<string> = new Set(
  AUTHORABLE_SLIDE_TRANSITIONS,
);

/** The first six effects of PowerPoint desktop's transition gallery. */
export const QUICK_SLIDE_TRANSITIONS = [
  "none",
  "morph",
  "fade",
  "push",
  "wipe",
  "split",
] as const satisfies readonly AuthorableSlideTransition[];
const PRESENTATION_SLIDE_TRANSITION_RECEIPT_KEYS: ReadonlySet<string> = new Set([
  "slideTransitionSchemaVersion",
  "slideTransitionDisposition",
  "slideTransitionContainerPresent",
  "slideTransitionEffectPresent",
  "slideTransitionOriginalKind",
  "slideTransitionOriginalDurationMs",
  "slideTransitionOriginalDurationSource",
  "slideTransitionOriginalSpeed",
  "slideTransitionOriginalDirection",
  "slideTransitionOriginalAdvanceOnClick",
  "slideTransitionOriginalAdvanceAfterMs",
  "slideTransitionFingerprint",
  "slideTransitionRejectReason",
]);

/** PowerPoint Mac 16.86 resets a newly selected basic effect to 01.00s. */
export const BASIC_TRANSITION_DURATION_MS = 1000;

/** PowerPoint Mac 16.86's observed Morph gallery default is 02.00s. */
export const MORPH_TRANSITION_DURATION_MS = 2000;

/**
 * What PowerPoint 16.86's Duration field shows while the slide has no
 * transition (a new blank deck, None selected: "02.00"). Display only -- a
 * slide without a transition stores no duration.
 */
export const NONE_TRANSITION_DISPLAY_DURATION_MS = 2000;

/** PowerPoint 16.86 Duration field authoring contract (hundredths of a second). */
export const MIN_TRANSITION_DURATION_MS = 0;
export const MAX_TRANSITION_DURATION_MS = 59_990;
export const TRANSITION_DURATION_STEP_MS = 10;
export const MIN_TRANSITION_DURATION_SECONDS = 0;
export const MAX_TRANSITION_DURATION_SECONDS = 59.99;
export const TRANSITION_DURATION_STEP_SECONDS = 0.01;

/** OOXML `advTm` is an unsigned 32-bit millisecond value. */
export const MIN_TRANSITION_ADVANCE_AFTER_MS = 0;
export const MAX_TRANSITION_ADVANCE_AFTER_MS = 0xffff_ffff;
/** PowerPoint Mac 16.86 exposes the After field to hundredths of a second. */
export const MAX_EDITABLE_TRANSITION_ADVANCE_AFTER_MS = 600_000;
export const TRANSITION_ADVANCE_AFTER_STEP_MS = 10;
export const TRANSITION_ADVANCE_AFTER_STEP_SECONDS = 0.01;

/** Immutable native-receipt boundary shared by browser/API projections. */
export const PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION =
  "pptx-slide-transition-v1" as const;
export const PRESENTATION_SLIDE_TRANSITION_DISPOSITIONS = [
  "native",
  "preserve-only",
] as const;
export type PresentationSlideTransitionDisposition =
  (typeof PRESENTATION_SLIDE_TRANSITION_DISPOSITIONS)[number];

export interface PresentationSlideTransitionReceipt {
  readonly slideTransitionSchemaVersion: typeof PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION;
  readonly slideTransitionDisposition: PresentationSlideTransitionDisposition;
  readonly slideTransitionContainerPresent: boolean;
  readonly slideTransitionEffectPresent: boolean | null;
  readonly slideTransitionOriginalKind: BasicSlideTransition | null;
  readonly slideTransitionOriginalDurationMs: number | null;
  readonly slideTransitionOriginalDurationSource:
    | "exact"
    | "speed"
    | "default"
    | null;
  readonly slideTransitionOriginalSpeed: "fast" | "med" | "slow" | null;
  /** Effective OOXML Effect Options identity for every directional effect. */
  readonly slideTransitionOriginalDirection?: SlideTransitionDirection | null;
  /** `advClick` omission is the OOXML default `true`. */
  readonly slideTransitionOriginalAdvanceOnClick?: boolean | null;
  /** `advTm` omission means automatic advance is disabled. */
  readonly slideTransitionOriginalAdvanceAfterMs?: number | null;
  readonly slideTransitionFingerprint: string;
  readonly slideTransitionRejectReason?: string;
}

/** The transition applied when a slide has no explicit (or invalid) value. */
export const DEFAULT_SLIDE_TRANSITION: SlideTransition = "none";

const SLIDE_TRANSITION_SET: ReadonlySet<string> = new Set(SLIDE_TRANSITIONS);

/**
 * Type guard: is `value` exactly one of the allowed transition identifiers?
 */
export function isSlideTransition(value: unknown): value is SlideTransition {
  return typeof value === "string" && SLIDE_TRANSITION_SET.has(value);
}

/** Exact, case-sensitive guard for the bounded native gallery command values. */
export function isBasicSlideTransition(
  value: unknown,
): value is BasicSlideTransition {
  return typeof value === "string" && BASIC_SLIDE_TRANSITION_SET.has(value);
}

/** Exact guard for the bounded Office 2013 preset-transition registry. */
export function isSlideTransitionP15PresetKind(
  value: unknown,
): value is SlideTransitionP15PresetKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P15_PRESET_KIND_SET.has(value)
  );
}

/** Exact guard for the Office 2013 presets whose Effect Options are sides. */
export function isSlideTransitionP15LeftRightKind(
  value: unknown,
): value is SlideTransitionP15LeftRightKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P15_LEFT_RIGHT_KIND_SET.has(value)
  );
}

/** Exact guard for the bounded Office 2010 left/right transition family. */
export function isSlideTransitionP14LeftRightKind(
  value: unknown,
): value is SlideTransitionP14LeftRightKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P14_LEFT_RIGHT_KIND_SET.has(value)
  );
}

/** Exact guard for the bounded Office 2010 orientation transition family. */
export function isSlideTransitionP14OrientationKind(
  value: unknown,
): value is SlideTransitionP14OrientationKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P14_ORIENTATION_KIND_SET.has(value)
  );
}

/** Exact guard for the bounded Office 2010 four-way side-direction family. */
export function isSlideTransitionP14SideDirectionKind(
  value: unknown,
): value is SlideTransitionP14SideDirectionKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P14_SIDE_DIRECTION_KIND_SET.has(value)
  );
}

/** Exact guard for the bounded Office 2010 optionless leaf family. */
export function isSlideTransitionP14EmptyKind(
  value: unknown,
): value is SlideTransitionP14EmptyKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P14_EMPTY_KIND_SET.has(value)
  );
}

/** Exact guard for the independently typed Office 2010 p14 family. */
export function isSlideTransitionP14TypedOptionsKind(
  value: unknown,
): value is SlideTransitionP14TypedOptionsKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_P14_TYPED_OPTIONS_KIND_SET.has(value)
  );
}

/** Exact guard for native Cover / Uncover eight-direction effects. */
export function isSlideTransitionEightDirectionKind(
  value: unknown,
): value is SlideTransitionEightDirectionKind {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_EIGHT_DIRECTION_KIND_SET.has(value)
  );
}

/** Exact guard for every persisted native Effect Options identity. */
export function isSlideTransitionDirection(
  value: unknown,
): value is SlideTransitionDirection {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_EFFECT_OPTION_SET.has(value)
  );
}

export function isSlideTransitionCornerOption(
  value: unknown,
): value is SlideTransitionCornerOption {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_CORNER_OPTION_SET.has(value)
  );
}

export function isSlideTransitionSideDirection(
  value: unknown,
): value is SlideTransitionSideDirection {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_SIDE_DIRECTION_SET.has(value)
  );
}

export function isSlideTransitionShapeOption(
  value: unknown,
): value is SlideTransitionShapeOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_SHAPE_OPTION_SET.has(value)
  );
}

export function isSlideTransitionZoomOption(
  value: unknown,
): value is SlideTransitionZoomOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_ZOOM_OPTION_SET.has(value)
  );
}

export function isSlideTransitionSplitOption(
  value: unknown,
): value is SlideTransitionSplitOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_SPLIT_OPTION_SET.has(value)
  );
}

export function isSlideTransitionOrientationOption(
  value: unknown,
): value is SlideTransitionOrientationOption {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_ORIENTATION_OPTION_SET.has(value)
  );
}

export function isSlideTransitionRevealOption(
  value: unknown,
): value is SlideTransitionRevealOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_REVEAL_OPTION_SET.has(value)
  );
}

export function isSlideTransitionRippleOption(
  value: unknown,
): value is SlideTransitionRippleOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_RIPPLE_OPTION_SET.has(value)
  );
}

export function isSlideTransitionGlitterOption(
  value: unknown,
): value is SlideTransitionGlitterOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_GLITTER_OPTION_SET.has(value)
  );
}

export function isSlideTransitionShredOption(
  value: unknown,
): value is SlideTransitionShredOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_SHRED_OPTION_SET.has(value)
  );
}

export function isSlideTransitionFlyThroughOption(
  value: unknown,
): value is SlideTransitionFlyThroughOption {
  return (
    typeof value === "string" &&
    SLIDE_TRANSITION_FLY_THROUGH_OPTION_SET.has(value)
  );
}

export function isSlideTransitionClockOption(
  value: unknown,
): value is SlideTransitionClockOption {
  return (
    typeof value === "string" && SLIDE_TRANSITION_CLOCK_OPTION_SET.has(value)
  );
}

export function isDirectionalSlideTransition(
  value: unknown,
): value is
  | "push"
  | "wipe"
  | "shape"
  | "zoom"
  | "split"
  | "random-bars"
  | "checkerboard"
  | "blinds"
  | "comb"
  | "clock"
  | "morph"
  | SlideTransitionP14LeftRightKind
  | SlideTransitionP14OrientationKind
  | SlideTransitionP14SideDirectionKind
  | SlideTransitionP14TypedOptionsKind
  | SlideTransitionP15LeftRightKind
  | SlideTransitionEightDirectionKind {
  return (
    value === "push" ||
    value === "wipe" ||
    value === "shape" ||
    value === "zoom" ||
    value === "split" ||
    value === "random-bars" ||
    value === "checkerboard" ||
    value === "blinds" ||
    value === "comb" ||
    value === "clock" ||
    value === "morph" ||
    isSlideTransitionP14LeftRightKind(value) ||
    isSlideTransitionP14OrientationKind(value) ||
    isSlideTransitionP14SideDirectionKind(value) ||
    isSlideTransitionP14TypedOptionsKind(value) ||
    isSlideTransitionP15LeftRightKind(value) ||
    isSlideTransitionEightDirectionKind(value)
  );
}

/** Exact writer-backed Effect Options for one native transition kind. */
export function slideTransitionEffectOptions(
  kind: unknown,
): readonly SlideTransitionDirection[] {
  if (kind === "push") return SLIDE_TRANSITION_PUSH_OPTIONS;
  if (kind === "wipe") return SLIDE_TRANSITION_DIRECTIONS;
  if (kind === "shape") return SLIDE_TRANSITION_SHAPE_OPTIONS;
  if (kind === "zoom") return SLIDE_TRANSITION_ZOOM_OPTIONS;
  if (kind === "split") return SLIDE_TRANSITION_SPLIT_OPTIONS;
  if (kind === "clock") return SLIDE_TRANSITION_CLOCK_OPTIONS;
  if (kind === "morph") return SLIDE_TRANSITION_MORPH_OPTIONS;
  if (
    isSlideTransitionP14LeftRightKind(kind) ||
    isSlideTransitionP15LeftRightKind(kind)
  ) {
    return SLIDE_TRANSITION_LEFT_RIGHT_OPTIONS;
  }
  if (isSlideTransitionEightDirectionKind(kind)) {
    return SLIDE_TRANSITION_EIGHT_DIRECTION_OPTIONS;
  }
  if (isSlideTransitionP14OrientationKind(kind)) {
    return SLIDE_TRANSITION_ORIENTATION_OPTIONS;
  }
  if (isSlideTransitionP14SideDirectionKind(kind)) {
    return SLIDE_TRANSITION_DIRECTIONS;
  }
  if (kind === "reveal") return SLIDE_TRANSITION_REVEAL_OPTIONS;
  if (kind === "ripple") return SLIDE_TRANSITION_RIPPLE_OPTIONS;
  if (kind === "glitter") return SLIDE_TRANSITION_GLITTER_OPTIONS;
  if (kind === "shred") return SLIDE_TRANSITION_SHRED_OPTIONS;
  if (kind === "fly-through") return SLIDE_TRANSITION_FLY_THROUGH_OPTIONS;
  return kind === "random-bars" ||
    kind === "checkerboard" ||
    kind === "blinds" ||
    kind === "comb"
    ? SLIDE_TRANSITION_ORIENTATION_OPTIONS
    : [];
}

/** Prevent an option from one native effect being written onto another. */
export function isSlideTransitionDirectionForKind(
  kind: unknown,
  value: unknown,
): value is SlideTransitionDirection {
  return (
    isSlideTransitionDirection(value) &&
    slideTransitionEffectOptions(kind).includes(value)
  );
}

/** PowerPoint-compatible effective direction when the editor field is absent. */
export function defaultSlideTransitionDirection(
  kind: unknown,
): SlideTransitionDirection | null {
  return isDirectionalSlideTransition(kind)
    ? DIRECTIONAL_SLIDE_TRANSITION_DEFAULTS[kind]
    : null;
}

/** Resolve a durable direction without granting edit authority. */
export function normalizeSlideTransitionDirection(
  kind: unknown,
  value: unknown,
): SlideTransitionDirection | null {
  const fallback = defaultSlideTransitionDirection(kind);
  if (fallback === null) return null;
  return isSlideTransitionDirectionForKind(kind, value) ? value : fallback;
}

/** Exact, case-sensitive guard for the visible authoring gallery row. */
export function isAuthorableSlideTransition(
  value: unknown,
): value is AuthorableSlideTransition {
  return (
    typeof value === "string" && AUTHORABLE_SLIDE_TRANSITION_SET.has(value)
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Authenticate an imported slide's immutable transition receipt. Locally
 * created slides have no native slide id and remain editable; an older imported
 * slide with a native id but no receipt is deliberately preserve-only.
 */
export function hasEditableSlideTransitionSource(source: unknown): boolean {
  if (source === undefined || source === null) return true;
  if (!isRecord(source)) return false;
  const receiptKeys = Object.keys(source).filter((key) =>
    key.startsWith("slideTransition"),
  );
  if (receiptKeys.length === 0) return source.slideId === undefined;
  if (
    receiptKeys.some(
      (key) => !PRESENTATION_SLIDE_TRANSITION_RECEIPT_KEYS.has(key),
    ) ||
    source.slideTransitionSchemaVersion !==
      PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION ||
    source.slideTransitionDisposition !== "native" ||
    typeof source.slideTransitionContainerPresent !== "boolean" ||
    typeof source.slideTransitionFingerprint !== "string" ||
    !/^[0-9a-f]{64}$/.test(source.slideTransitionFingerprint)
  ) {
    return false;
  }
  if (
    Object.hasOwn(source, "slideTransitionRejectReason") ||
    (source.slideId !== undefined &&
      (typeof source.slideId !== "string" ||
        !/^sld[1-9][0-9]*$/.test(source.slideId)))
  ) {
    return false;
  }
  const containerPresent = source.slideTransitionContainerPresent;
  const effectPresent = source.slideTransitionEffectPresent;
  const kind = source.slideTransitionOriginalKind;
  if (typeof effectPresent !== "boolean") return false;
  if (effectPresent) {
    if (
      !containerPresent ||
      (!isBasicSlideTransition(kind) && kind !== "morph") ||
      kind === "none"
    ) {
      return false;
    }
  } else if (kind !== "none") {
    return false;
  }
  const duration = source.slideTransitionOriginalDurationMs;
  const durationSource = source.slideTransitionOriginalDurationSource;
  if (effectPresent) {
    if (
      !Number.isInteger(duration) ||
      (duration as number) < MIN_TRANSITION_DURATION_MS ||
      (duration as number) > MAX_TRANSITION_DURATION_MS ||
      (durationSource !== "exact" &&
        durationSource !== "speed" &&
        durationSource !== "default")
    ) {
      return false;
    }
  } else if (duration !== null || durationSource !== null) {
    return false;
  }
  const speed = source.slideTransitionOriginalSpeed;
  if (speed !== null && speed !== "fast" && speed !== "med" && speed !== "slow") {
    return false;
  }
  if (!containerPresent && speed !== null) return false;
  if (durationSource === "speed" && speed === null) return false;
  if (durationSource === "default" && speed !== null) return false;
  const hasDirection = Object.hasOwn(
    source,
    "slideTransitionOriginalDirection",
  );
  if (
    isDirectionalSlideTransition(kind) &&
    kind !== "push" &&
    kind !== "wipe" &&
    // A Morph receipt from before its Effect Options names none: by object.
    kind !== "morph" &&
    !hasDirection
  ) {
    return false;
  }
  if (hasDirection) {
    const direction = source.slideTransitionOriginalDirection;
    // A preset receipt stamped before its Left/Right Effect Options names no
    // side. Back then exactly one identity per preset was native and it was
    // always the right-hand one, so the receipt still says which slide it
    // came from; keeping it editable is what stops a re-export mirroring it.
    const legacyDirectionlessPreset =
      direction === null && (isSlideTransitionP15LeftRightKind(kind) || kind === "morph");
    if (
      (isDirectionalSlideTransition(kind) &&
        !legacyDirectionlessPreset &&
        !isSlideTransitionDirectionForKind(kind, direction)) ||
      (!isDirectionalSlideTransition(kind) && direction !== null)
    ) {
      return false;
    }
  }
  const hasAdvanceOnClick = Object.hasOwn(
    source,
    "slideTransitionOriginalAdvanceOnClick",
  );
  const hasAdvanceAfter = Object.hasOwn(
    source,
    "slideTransitionOriginalAdvanceAfterMs",
  );
  // Advance evidence was added after the original v1 receipt shipped. Keep old
  // receipts valid for TRN01/TRN02, but reject a partially forged extension.
  if (hasAdvanceOnClick !== hasAdvanceAfter) return false;
  if (
    hasAdvanceOnClick &&
    (typeof source.slideTransitionOriginalAdvanceOnClick !== "boolean" ||
      (source.slideTransitionOriginalAdvanceAfterMs !== null &&
        (!Number.isInteger(source.slideTransitionOriginalAdvanceAfterMs) ||
          (source.slideTransitionOriginalAdvanceAfterMs as number) <
            MIN_TRANSITION_ADVANCE_AFTER_MS ||
          (source.slideTransitionOriginalAdvanceAfterMs as number) >
            MAX_TRANSITION_ADVANCE_AFTER_MS)))
  ) {
    return false;
  }
  return true;
}

/**
 * Native Morph uses the exact byObject wrapper authenticated by the importer.
 * Legacy preserve-only evidence is never upgraded just from its visual kind.
 */
export function hasAuthorableMorphTransitionSource(source: unknown): boolean {
  if (source === undefined || source === null) return true;
  if (!isRecord(source)) return false;
  if (Object.hasOwn(source, "slideId")) return hasEditableSlideTransitionSource(source);
  return !Object.keys(source).some((key) => key.startsWith("slideTransition"));
}

/** Kind-aware authoring gate shared by live commands and gallery disabled UI. */
export function canAuthorSlideTransitionKind(
  source: unknown,
  kind: unknown,
): kind is AuthorableSlideTransition {
  if (!isAuthorableSlideTransition(kind)) return false;
  return kind === "morph"
    ? hasAuthorableMorphTransitionSource(source)
    : hasEditableSlideTransitionSource(source);
}

/** Default applied when a gallery effect is selected, or null for None. */
export function defaultAuthoringTransitionDuration(
  kind: AuthorableSlideTransition,
): number | null {
  if (kind === "none") return null;
  return kind === "morph"
    ? MORPH_TRANSITION_DURATION_MS
    : BASIC_TRANSITION_DURATION_MS;
}

/**
 * Advance authoring needs explicit immutable evidence on imported slides. The
 * older effect/duration receipt stays usable by its original commands, but it
 * cannot manufacture custody for `advClick` / `advTm`.
 */
export function hasEditableSlideTransitionAdvanceSource(source: unknown): boolean {
  if (source === undefined || source === null) return true;
  if (!hasEditableSlideTransitionSource(source) || !isRecord(source)) return false;
  return (
    Object.hasOwn(source, "slideTransitionOriginalAdvanceOnClick") &&
    Object.hasOwn(source, "slideTransitionOriginalAdvanceAfterMs")
  );
}

/**
 * Imported direction edits require explicit immutable direction evidence.
 * Locally-created slides have no original-package custody boundary.
 */
export function hasEditableSlideTransitionDirectionSource(
  source: unknown,
): boolean {
  if (source === undefined || source === null) return true;
  return (
    hasEditableSlideTransitionSource(source) &&
    isRecord(source) &&
    Object.hasOwn(source, "slideTransitionOriginalDirection")
  );
}

/** OOXML default: a slide advances on click unless `advClick` says otherwise. */
export function transitionAdvanceOnClick(value: unknown): boolean {
  return value === undefined ? true : value === true;
}

/** Strict durable/receipt guard for the automatic-advance millisecond value. */
export function isTransitionAdvanceAfterMs(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= MIN_TRANSITION_ADVANCE_AFTER_MS &&
    value <= MAX_TRANSITION_ADVANCE_AFTER_MS
  );
}

/** True only for values the two-decimal PowerPoint field can author. */
export function isEditableTransitionAdvanceAfterMs(
  value: unknown,
): value is number {
  return (
    isTransitionAdvanceAfterMs(value) &&
    value <= MAX_EDITABLE_TRANSITION_ADVANCE_AFTER_MS &&
    value % TRANSITION_ADVANCE_AFTER_STEP_MS === 0
  );
}

/** Parse seconds without clamping invalid text into a real edit. */
export function parseTransitionAdvanceAfterSeconds(
  value: unknown,
): number | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const text = String(value).trim();
  if (!/^\d+(?:\.\d{0,2})?$/.test(text)) return undefined;
  const seconds = Number(text);
  if (!Number.isFinite(seconds)) return undefined;
  const milliseconds = Math.round(seconds * 1000);
  return isEditableTransitionAdvanceAfterMs(milliseconds)
    ? milliseconds
    : undefined;
}

/**
 * PowerPoint's Duration and After fields show seconds as at least two whole
 * digits and two decimals ("02.00", "00.50", "125.00"); the parsers accept the
 * padded and the plain spelling alike.
 */
export function formatPowerPointSecondsField(milliseconds: number): string {
  const [whole, fraction] = (milliseconds / 1000).toFixed(2).split(".");
  return `${whole.padStart(2, "0")}.${fraction}`;
}

/** Display one persisted `advTm` value in PowerPoint's seconds field. */
export function formatTransitionAdvanceAfterSeconds(value: unknown): string {
  return formatPowerPointSecondsField(
    isTransitionAdvanceAfterMs(value) ? value : 0,
  );
}

/**
 * Normalize an arbitrary persisted value into an allowed transition. Unknown,
 * empty, or non-string values collapse to {@link DEFAULT_SLIDE_TRANSITION} so a
 * stale or hand-edited deck can never inject an unsupported animation. The
 * comparison is case-insensitive and trims surrounding whitespace; everything
 * else (including any other casing/spacing variation) falls back to the default.
 */
export function normalizeSlideTransition(value: unknown): SlideTransition {
  if (typeof value !== "string") return DEFAULT_SLIDE_TRANSITION;
  const candidate = value.trim().toLowerCase();
  return isSlideTransition(candidate) ? candidate : DEFAULT_SLIDE_TRANSITION;
}

/**
 * The CSS class that plays a transition's keyframes when projecting a slide.
 * Returns `null` for `none` (and any value that normalizes to it) so the caller
 * can skip adding an animation class entirely. The class names are defined in
 * `styles/native-page-window.css` (`.presentation-present__slide--enter-*`).
 *
 * `morph` ALSO returns `null`: it is not a CSS entrance animation but a per-block
 * FLIP tween driven separately in present mode (`lib/presentation-morph.ts`). Cut
 * returns `null` because its native meaning is an instantaneous replacement, not
 * an animation. Random resolves to an existing renderer at playback time rather
 * than manufacturing a fixed pseudo-random keyframe.
 */
export function slideTransitionClassName(
  value: unknown,
  direction?: unknown,
  randomSource: () => number = Math.random,
): string | null {
  const transition = normalizeSlideTransition(value);
  if (
    transition === "none" ||
    transition === "morph" ||
    transition === "cut"
  ) {
    return null;
  }
  // The basic-gallery identities reuse the already-shipped visual keyframes;
  // native serialization remains exact and editor-specific.
  if (transition === "push") {
    return `presentation-present__slide--enter-push-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (transition === "wipe") {
    return `presentation-present__slide--enter-wipe-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (transition === "shape") {
    return `presentation-present__slide--enter-shape-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (transition === "zoom") {
    return normalizeSlideTransitionDirection(transition, direction) === "out"
      ? "presentation-present__slide--enter-zoom-out"
      : "presentation-present__slide--enter-zoom";
  }
  if (transition === "split") {
    return `presentation-present__slide--enter-split-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (
    transition === "random-bars" ||
    transition === "checkerboard" ||
    transition === "blinds" ||
    transition === "comb"
  ) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (
    isSlideTransitionP14LeftRightKind(transition) ||
    isSlideTransitionP15LeftRightKind(transition)
  ) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (isSlideTransitionEightDirectionKind(transition)) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (isSlideTransitionP14OrientationKind(transition)) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (isSlideTransitionP14SideDirectionKind(transition)) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (isSlideTransitionP14TypedOptionsKind(transition)) {
    return `presentation-present__slide--enter-${transition}-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (transition === "clock") {
    return `presentation-present__slide--enter-clock-${normalizeSlideTransitionDirection(transition, direction)}`;
  }
  if (transition === "random") {
    return randomSlideTransitionClassName(randomSource);
  }
  return `presentation-present__slide--enter-${transition}`;
}

/**
 * Per-slide transition DURATION bounds (ms). The default mirrors the fixed
 * `320ms` baked into every `.presentation-present__slide--enter-*` rule in
 * `styles/native-page-window.css`; a slide WITHOUT an explicit duration uses
 * that CSS value unchanged (so an old deck renders byte-for-byte as before).
 * When set, the value drives an inline `animation-duration` override clamped to
 * `[MIN, MAX]` so a hand-edited deck can never inject an absurd (negative,
 * NaN, multi-minute) duration.
 */
export const DEFAULT_TRANSITION_DURATION_MS = 320;

/**
 * Existing renderers eligible for native `p:random` playback. PowerPoint owns
 * the choice at render time; the document stores no selected child effect.
 * Keep this bounded to already-supported, non-Morph, non-instantaneous classes.
 */
export const RANDOM_SLIDE_TRANSITION_RENDERER_CLASSES = [
  "presentation-present__slide--enter-fade",
  "presentation-present__slide--enter-push-up",
  "presentation-present__slide--enter-wipe-right",
  "presentation-present__slide--enter-dissolve",
  "presentation-present__slide--enter-shape-circle",
  "presentation-present__slide--enter-zoom",
  "presentation-present__slide--enter-split-horizontal-out",
  "presentation-present__slide--enter-random-bars-horizontal",
] as const;

export function randomSlideTransitionClassName(
  randomSource: () => number = Math.random,
): (typeof RANDOM_SLIDE_TRANSITION_RENDERER_CLASSES)[number] {
  const sample = randomSource();
  const boundedSample = Number.isFinite(sample)
    ? Math.max(0, Math.min(sample, 1 - Number.EPSILON))
    : 0;
  return RANDOM_SLIDE_TRANSITION_RENDERER_CLASSES[
    Math.floor(boundedSample * RANDOM_SLIDE_TRANSITION_RENDERER_CLASSES.length)
  ]!;
}

export interface SlideTransitionPreviewPlayback {
  readonly slideModelId: string;
  readonly className: string | null;
  readonly durationMs: number;
}

export interface MorphSlideTransitionPreviewPlayback {
  readonly kind: "morph";
  readonly fromSlideModelId: string;
  readonly slideModelId: string;
  readonly durationMs: number;
}

/**
 * Resolve the current slide's safe, render-only transition preview. This is
 * deliberately independent from native edit receipts: a uniquely projected
 * legacy transition may still play, but this helper grants no authoring or
 * patchback authority. Cut resolves to an instantaneous playback plan with no
 * fabricated CSS class. `none`, unknown values, and Morph (which needs the
 * present-mode previous-slide FLIP layer) have no single-slide preview here.
 */
export function resolveSlideTransitionPreview(
  source: {
    readonly id?: unknown;
    readonly transition?: unknown;
    readonly transitionDirection?: unknown;
    readonly transitionDurationMs?: unknown;
  },
  randomSource: () => number = Math.random,
): SlideTransitionPreviewPlayback | null {
  if (typeof source.id !== "string" || source.id.length === 0) return null;
  const transition = normalizeSlideTransition(source.transition);
  const className = slideTransitionClassName(
    transition,
    source.transitionDirection,
    randomSource,
  );
  if (!className && transition !== "cut") return null;
  return {
    slideModelId: source.id,
    className,
    durationMs:
      transition === "cut"
        ? 0
        : source.transitionDurationMs === undefined
          ? DEFAULT_TRANSITION_DURATION_MS
          : clampTransitionDuration(source.transitionDurationMs),
  };
}

/**
 * Resolve Morph's editor Preview plan from the live adjacent slide pair. This
 * is render-only authority: it intentionally accepts a safe imported visual
 * projection without granting edit or patchback custody.
 */
export function resolveMorphSlideTransitionPreview(
  slides: readonly {
    readonly id?: unknown;
    readonly transition?: unknown;
    readonly transitionDurationMs?: unknown;
  }[],
  activeIndex: number,
): MorphSlideTransitionPreviewPlayback | null {
  if (
    !Number.isInteger(activeIndex) ||
    activeIndex <= 0 ||
    activeIndex >= slides.length
  ) {
    return null;
  }
  const from = slides[activeIndex - 1];
  const target = slides[activeIndex];
  if (
    typeof from?.id !== "string" ||
    from.id.length === 0 ||
    typeof target?.id !== "string" ||
    target.id.length === 0 ||
    from.id === target.id ||
    normalizeSlideTransition(target.transition) !== "morph"
  ) {
    return null;
  }
  return {
    kind: "morph",
    fromSlideModelId: from.id,
    slideModelId: target.id,
    durationMs:
      target.transitionDurationMs === undefined
        ? MORPH_TRANSITION_DURATION_MS
        : clampTransitionDuration(target.transitionDurationMs),
  };
}

/** True only for values the editable Duration field can author. */
export function isEditableTransitionDurationMs(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= MIN_TRANSITION_DURATION_MS &&
    value <= MAX_TRANSITION_DURATION_MS &&
    value % TRANSITION_DURATION_STEP_MS === 0
  );
}

/**
 * Parse the PowerPoint-style seconds field without clamping invalid input into
 * a real edit. The UI exposes hundredths, while the durable model remains ms.
 */
export function parseTransitionDurationSeconds(
  value: unknown,
): number | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const text = String(value).trim();
  if (!/^\d+(?:\.\d{0,2})?$/.test(text)) return undefined;
  const seconds = Number(text);
  if (
    !Number.isFinite(seconds) ||
    seconds < MIN_TRANSITION_DURATION_SECONDS ||
    seconds > MAX_TRANSITION_DURATION_SECONDS
  ) {
    return undefined;
  }
  const durationMs = Math.round(seconds * 1000);
  return isEditableTransitionDurationMs(durationMs) ? durationMs : undefined;
}

/** Format a persisted duration for the two-decimal seconds field. */
export function formatTransitionDurationSeconds(value: unknown): string {
  const clamped = clampTransitionDuration(value);
  const stepped =
    Math.round(clamped / TRANSITION_DURATION_STEP_MS) *
    TRANSITION_DURATION_STEP_MS;
  return formatPowerPointSecondsField(stepped);
}

/**
 * Clamp an arbitrary persisted transition duration into the supported range,
 * rounding to a whole millisecond. Non-finite / non-number / out-of-range values
 * collapse to the nearest bound (or {@link DEFAULT_TRANSITION_DURATION_MS} when
 * not a usable number at all) so a stale or hand-edited value is always a safe,
 * bounded integer. Use {@link normalizeTransitionDuration} when you need to
 * distinguish "no explicit duration" (byte-stable default) from a real value.
 */
export function clampTransitionDuration(value: unknown): number {
  const n =
    typeof value === "number"
      ? value
      : typeof value === "string"
        ? Number(value.trim())
        : Number.NaN;
  if (!Number.isFinite(n)) return DEFAULT_TRANSITION_DURATION_MS;
  const rounded = Math.round(n);
  if (rounded < MIN_TRANSITION_DURATION_MS) return MIN_TRANSITION_DURATION_MS;
  if (rounded > MAX_TRANSITION_DURATION_MS) return MAX_TRANSITION_DURATION_MS;
  return rounded;
}

/**
 * How long a slide's entrance transition runs in the show, as the show times
 * it: none for None or Cut, a Morph's own clamp (and none when no slide came
 * before it to morph from), the rest their stated length or a second. The
 * slide's own timeline -- its first builds, the time its After counts --
 * starts when it ends, as PowerPoint plays it.
 */
export function slideEntranceTransitionMs(
  slide: { readonly transition?: unknown; readonly transitionDurationMs?: unknown },
  afterAnother = true,
): number {
  const kind = normalizeSlideTransition(slide.transition);
  if (kind === "none" || kind === "cut" || (kind === "morph" && !afterAnother)) return 0;
  if (kind === "morph" || slide.transitionDurationMs !== undefined) return clampTransitionDuration(slide.transitionDurationMs);
  return BASIC_TRANSITION_DURATION_MS;
}

/**
 * Normalize a persisted transition duration for SERIALIZATION. Returns a clamped
 * integer for a usable value, or `undefined` when the value is absent / not a
 * usable number — so a slide that never set a duration persists NO field and
 * renders with the CSS default (byte-for-byte unchanged). A value that equals
 * the default IS retained when explicitly provided so the control round-trips;
 * callers that prefer to drop the default for byte-cleanliness can compare
 * against {@link DEFAULT_TRANSITION_DURATION_MS} themselves.
 */
export function normalizeTransitionDuration(value: unknown): number | undefined {
  if (value === undefined || value === null || value === "") return undefined;
  const n =
    typeof value === "number"
      ? value
      : typeof value === "string"
        ? Number(value.trim())
        : Number.NaN;
  if (!Number.isFinite(n)) return undefined;
  return clampTransitionDuration(n);
}

interface ApplyTransitionToAllSlide {
  readonly id?: unknown;
  readonly transition?: unknown;
  readonly transitionDurationMs?: unknown;
  readonly transitionDirection?: unknown;
  readonly transitionZoomElement?: unknown;
  readonly transitionAdvanceOnClick?: unknown;
  readonly transitionAdvanceAfterMs?: unknown;
  readonly transitionSound?: unknown;
  readonly source?: unknown;
}

interface ApplyTransitionToAllState {
  readonly kind: AuthorableSlideTransition;
  readonly durationMs: number | null;
  readonly direction: SlideTransitionDirection | null;
  readonly zoomElement: SlideTransitionZoomElement | null;
  readonly advanceOnClick: boolean;
  readonly advanceAfterMs: number | null;
  readonly sound: PresentationTransitionSound | null;
}

/**
 * Validate the complete TRN01–03 authoring state without promoting a legacy
 * render-only projection into native edit authority.
 */
function applyTransitionToAllState(
  slide: ApplyTransitionToAllSlide,
): ApplyTransitionToAllState | null {
  const source =
    slide.source && typeof slide.source === "object" && !Array.isArray(slide.source)
      ? (slide.source as Record<string, unknown>)
      : null;
  const importedSoundEditable =
    source?.slideId === undefined ||
    source.transitionSoundDisposition === "absent" ||
    source.transitionSoundDisposition === "native";
  if (
    typeof slide.id !== "string" ||
    slide.id.length === 0 ||
    !hasEditableSlideTransitionSource(slide.source) ||
    !hasEditableSlideTransitionAdvanceSource(slide.source) ||
    !importedSoundEditable
  ) {
    return null;
  }
  const kind =
    slide.transition === undefined
      ? "none"
      : isAuthorableSlideTransition(slide.transition)
        ? slide.transition
        : null;
  if (kind === null || !canAuthorSlideTransitionKind(slide.source, kind)) {
    return null;
  }
  let direction: SlideTransitionDirection | null = null;
  if (isDirectionalSlideTransition(kind)) {
    if (
      !hasEditableSlideTransitionDirectionSource(slide.source) ||
      (slide.transitionDirection !== undefined &&
        !isSlideTransitionDirectionForKind(kind, slide.transitionDirection))
    ) {
      return null;
    }
    direction = normalizeSlideTransitionDirection(
      kind,
      slide.transitionDirection,
    );
  } else if (slide.transitionDirection !== undefined) {
    return null;
  }
  // Only a Zoom names an element, and only the Office 2010 one.
  let zoomElement: SlideTransitionZoomElement | null = null;
  if (slide.transitionZoomElement !== undefined) {
    if (kind !== "zoom" || !isSlideTransitionZoomElement(slide.transitionZoomElement)) {
      return null;
    }
    zoomElement = slide.transitionZoomElement;
  }
  let durationMs: number | null = null;
  if (kind === "none") {
    if (slide.transitionDurationMs !== undefined) return null;
  } else if (slide.transitionDurationMs === undefined) {
    durationMs = defaultAuthoringTransitionDuration(kind);
  } else if (isEditableTransitionDurationMs(slide.transitionDurationMs)) {
    durationMs = slide.transitionDurationMs;
  } else {
    return null;
  }
  if (
    slide.transitionAdvanceOnClick !== undefined &&
    typeof slide.transitionAdvanceOnClick !== "boolean"
  ) {
    return null;
  }
  const advanceAfterMs =
    slide.transitionAdvanceAfterMs === undefined
      ? null
      : isEditableTransitionAdvanceAfterMs(slide.transitionAdvanceAfterMs)
        ? slide.transitionAdvanceAfterMs
        : null;
  if (
    slide.transitionAdvanceAfterMs !== undefined &&
    advanceAfterMs === null
  ) {
    return null;
  }
  const sound =
    slide.transitionSound === undefined
      ? null
      : normalizePresentationTransitionSound(slide.transitionSound);
  if (slide.transitionSound !== undefined && sound === null) return null;
  return {
    kind,
    durationMs,
    direction,
    zoomElement,
    advanceOnClick: transitionAdvanceOnClick(
      slide.transitionAdvanceOnClick,
    ),
    advanceAfterMs,
    sound,
  };
}

/**
 * Copy one live slide's complete TRN01–03 transition state to every slide.
 * The entire deck is authenticated before cloning anything, so one opaque,
 * legacy, malformed, or stale target makes the operation an atomic no-op.
 * Returning the original array identity is the command adapter's no-commit
 * signal for both failed preflight and already-equal decks.
 */
export function applyTransitionToAllSlides<T extends ApplyTransitionToAllSlide>(
  slides: readonly T[],
  sourceSlideModelId: string,
): T[] {
  if (!sourceSlideModelId || slides.length === 0) return slides as T[];
  const seenIds = new Set<string>();
  const states: ApplyTransitionToAllState[] = [];
  let sourceState: ApplyTransitionToAllState | null = null;
  for (const slide of slides) {
    const state = applyTransitionToAllState(slide);
    if (!state || typeof slide.id !== "string" || seenIds.has(slide.id)) {
      return slides as T[];
    }
    seenIds.add(slide.id);
    states.push(state);
    if (slide.id === sourceSlideModelId) {
      if (sourceState !== null) return slides as T[];
      sourceState = state;
    }
  }
  if (sourceState === null) return slides as T[];
  const desired = sourceState;
  // Morph cannot be copied onto an imported slide until its p159 wrapper has
  // authenticated original-package patchback custody. Keep Apply All atomic.
  if (
    !slides.every((slide) =>
      canAuthorSlideTransitionKind(slide.source, desired.kind),
    )
  ) {
    return slides as T[];
  }

  let changed = false;
  const nextSlides = slides.map((slide, index) => {
    const current = states[index];
    if (!current) return slide;
    const zoomElementMatches =
      desired.zoomElement === null
        ? !Object.hasOwn(slide, "transitionZoomElement")
        : slide.transitionZoomElement === desired.zoomElement;
    const transitionMatches =
      desired.kind === "none"
        ? !Object.hasOwn(slide, "transition") &&
          !Object.hasOwn(slide, "transitionDurationMs") &&
          !Object.hasOwn(slide, "transitionDirection") &&
          !Object.hasOwn(slide, "transitionZoomElement")
        : slide.transition === desired.kind &&
          slide.transitionDurationMs === desired.durationMs &&
          zoomElementMatches &&
          (desired.direction === null
            ? !Object.hasOwn(slide, "transitionDirection")
            : normalizeSlideTransitionDirection(
                desired.kind,
                slide.transitionDirection,
              ) === desired.direction);
    const advanceOnClickMatches = desired.advanceOnClick
      ? !Object.hasOwn(slide, "transitionAdvanceOnClick")
      : slide.transitionAdvanceOnClick === false;
    const advanceAfterMatches =
      desired.advanceAfterMs === null
        ? !Object.hasOwn(slide, "transitionAdvanceAfterMs")
        : slide.transitionAdvanceAfterMs === desired.advanceAfterMs;
    const soundMatches = samePresentationTransitionSound(
      slide.transitionSound,
      desired.sound,
    );
    if (
      transitionMatches &&
      advanceOnClickMatches &&
      advanceAfterMatches &&
      soundMatches
    ) {
      return slide;
    }

    changed = true;
    const next = { ...slide } as T;
    if (desired.kind === "none") {
      delete (next as { transition?: unknown }).transition;
      delete (next as { transitionDurationMs?: unknown })
        .transitionDurationMs;
      delete (next as { transitionDirection?: unknown }).transitionDirection;
      delete (next as { transitionZoomElement?: unknown }).transitionZoomElement;
    } else {
      (next as { transition?: unknown }).transition = desired.kind;
      (next as { transitionDurationMs?: unknown }).transitionDurationMs =
        desired.durationMs;
      if (desired.direction === null) {
        delete (next as { transitionDirection?: unknown }).transitionDirection;
      } else {
        (next as { transitionDirection?: unknown }).transitionDirection =
          desired.direction;
      }
      if (desired.zoomElement === null) {
        delete (next as { transitionZoomElement?: unknown })
          .transitionZoomElement;
      } else {
        (next as { transitionZoomElement?: unknown }).transitionZoomElement =
          desired.zoomElement;
      }
    }
    if (desired.advanceOnClick) {
      delete (next as { transitionAdvanceOnClick?: unknown })
        .transitionAdvanceOnClick;
    } else {
      (next as { transitionAdvanceOnClick?: unknown })
        .transitionAdvanceOnClick = false;
    }
    if (desired.advanceAfterMs === null) {
      delete (next as { transitionAdvanceAfterMs?: unknown })
        .transitionAdvanceAfterMs;
    } else {
      (next as { transitionAdvanceAfterMs?: unknown })
        .transitionAdvanceAfterMs = desired.advanceAfterMs;
    }
    if (desired.sound === null) {
      delete (next as { transitionSound?: unknown }).transitionSound;
    } else {
      (next as { transitionSound?: unknown }).transitionSound = desired.sound;
    }
    return next;
  });
  return changed ? nextSlides : (slides as T[]);
}
