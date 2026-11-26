/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { smoothEasingCss } from "./presentation-smooth-timing";

/**
 * Effect Options of one animation step (PowerPoint's Effect Options dialog:
 * the Effect and Timing tabs beyond the effect, its direction and its
 * Start / Delay / Duration).
 *
 * The persisted shape is the API's `effectSettings`
 * (`apps/api/app/services/file_adapters/pptx/animations.py`,
 * `animation_effect_settings`), which spells every key exactly as PowerPoint
 * writes it in `p:timing` (measured on 11,412 PowerPoint decks):
 *
 *  - `accel` / `decel`: Smooth start / Smooth end, the effect node's share of
 *    the duration (100000 = all of it). Stated only when it differs from the
 *    effect's own (a motion path eases 50000 each way).
 *  - `autoReverse`: Auto-reverse, the effect plays forward then back.
 *  - `repeat`: Repeat, a count in thousandths of a play (2000 = twice), or
 *    "until-next-click" / "until-end-of-slide".
 *  - `rewind`: Rewind when done playing (the object returns as it was).
 *  - `afterEffect` (+ `dimColor`): After animation, hide when the effect ends,
 *    hide on the next click, or dim to a colour on the next click.
 *  - `textUnit` (+ `textUnitDelayPct` | `textUnitDelayMs`): Animate text By
 *    word / By letter, each unit starting a share of the duration or a time
 *    after the one before.
 *  - `wholeIterate` (+ `wholeIterateDelayPct` | `wholeIterateDelayMs`): an
 *    iteration over the whole object as the file wrote it (`p:iterate` with
 *    no type, "bare", or `type="el"`, "el"; row 1089). PowerPoint plays the
 *    effect on the object as a whole, as without one, so the show ignores it;
 *    it is kept so the effect is written back as it was. Never beside
 *    `textUnit` (choosing By word / By letter replaces it).
 *  - `flyStart`: "offSlide" -- a Fly In from the bottom, left or right
 *    written on one track from a whole slide away (PowerPoint for Mac's
 *    spelling, row 1096); absent, PowerPoint's own from just past the edge.
 *    It only keeps a file's spelling: the editor offers no such choice, and
 *    a new direction or effect drops it.
 *  - `growSize` (+ `growDirection`): Grow/Shrink's Size and Direction
 *    (`p:animScale/p:by`, row 1094), the size in thousandths of a percent --
 *    Tiny 25000, Smaller 50000, Larger 150000 (the default, never stated),
 *    Huge 400000 or a custom amount -- and the direction "horizontal" (x the
 *    size, y 100 %) or "vertical"; absent, both axes. A file's amount is
 *    kept as it is (0, negative, past 10000 % included: PowerPoint's object
 *    model keeps them, truth anim-grow-size); what a user types is limited
 *    to whole percents 1..9999 ({@link GROW_SIZE_TYPED_MIN_PCT}).
 *
 * Absent keys are the effect's own defaults; a step without the object is
 * byte-for-byte what it was before. This module has no dependencies, so the
 * zod schema, the editor and the show share it; an effect's own defaults
 * (which need the recipe catalog) are `presentation-anim-effect-defaults`.
 */

export const EFFECT_REPEAT_UNTIL = ["until-next-click", "until-end-of-slide"] as const;
export type EffectRepeatUntil = (typeof EFFECT_REPEAT_UNTIL)[number];
export const EFFECT_AFTER_KINDS = ["hide", "hide-on-next-click", "dim"] as const;
export type EffectAfterKind = (typeof EFFECT_AFTER_KINDS)[number];
export const EFFECT_TEXT_UNITS = ["word", "letter"] as const;
export type EffectTextUnit = (typeof EFFECT_TEXT_UNITS)[number];
/** Theme colours a Dim may name (`a:schemeClr`), as PowerPoint writes them. */
export const EFFECT_DIM_SCHEME_COLORS = [
  "bg1",
  "bg2",
  "tx1",
  "tx2",
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;
export type EffectDimSchemeColor = (typeof EFFECT_DIM_SCHEME_COLORS)[number];
/** The most plays one step may repeat, in thousandths (PowerPoint's 9999). */
export const MAX_EFFECT_REPEAT_THOUSANDTHS = 9_999_000;
export const MAX_EFFECT_SHARE = 100_000;
export const MAX_EFFECT_TEXT_DELAY_MS = 60_000;
/** PowerPoint's own delay between words / letters when they are first chosen. */
export const DEFAULT_EFFECT_TEXT_UNIT_DELAY_PCT = 10_000;

/**
 * Effects whose node does not end held (PowerPoint removes or states nothing
 * for them), so Rewind when done playing has no meaning there. Kept equal to
 * the API's `effect_default_fill(variant) != "hold"` plus Transparency (an
 * opacity that simply stays) by `test_effect_settings_web_parity`.
 */
export const PRESENTATION_EFFECTS_WITHOUT_REWIND: ReadonlySet<string> = new Set([
  "bold-reveal",
  "flicker",
  "font",
  "font-style",
  // A media call takes no Effect Options of an animation at all.
  "media-pause",
  "media-play",
  "style-emphasis",
  "transparency",
  "vertical-highlight",
]);

/** Grow/Shrink as the gallery adds it: Larger (150 %), both axes. */
export const DEFAULT_GROW_SIZE = 150_000;
/** The axis a one-way Grow/Shrink leaves alone stays at 100 %. */
export const GROW_UNCHANGED_AXIS = 100_000;
export const EFFECT_GROW_DIRECTIONS = ["horizontal", "vertical"] as const;
export type EffectGrowDirection = (typeof EFFECT_GROW_DIRECTIONS)[number];
/** PowerPoint's Size presets (Effect Options), in thousandths of a percent. */
export const GROW_SIZE_PRESETS = [
  { id: "tiny", size: 25_000 },
  { id: "smaller", size: 50_000 },
  { id: "larger", size: 150_000 },
  { id: "huge", size: 400_000 },
] as const;
export type GrowSizePreset = (typeof GROW_SIZE_PRESETS)[number]["id"];
/** What the Custom box takes when typed: whole percents. */
export const GROW_SIZE_TYPED_MIN_PCT = 1;
export const GROW_SIZE_TYPED_MAX_PCT = 9_999;
/** The widest amount p:by holds (a 32-bit whole number). */
const GROW_SIZE_LIMIT = 2_147_483_647;

export interface PresentationAnimationEffectSettings {
  accel?: number;
  decel?: number;
  autoReverse?: true;
  repeat?: number | EffectRepeatUntil;
  rewind?: true;
  afterEffect?: EffectAfterKind;
  dimColor?: string;
  textUnit?: EffectTextUnit;
  textUnitDelayPct?: number;
  textUnitDelayMs?: number;
  wholeIterate?: EffectWholeIterate;
  wholeIterateDelayPct?: number;
  wholeIterateDelayMs?: number;
  growSize?: number;
  growDirection?: EffectGrowDirection;
  flyStart?: "offSlide";
  /**
   * A templated effect's text colour track as the file states its two
   * colours (Color Typewriter, row 1308), each in its own form: an sRGB
   * value, or a theme / preset colour with its modifiers.
   */
  typewriterTextColors?: [EffectTrackColor, EffectTrackColor];
}

/** One colour of an effect's colour track, in the file's form (row 1308). */
export interface EffectTrackColor {
  kind: "srgbClr" | "schemeClr" | "prstClr";
  val: string;
  mods?: [string, string][];
}

/** An iteration over the whole object: no type stated, or "el" (row 1089). */
export const EFFECT_WHOLE_ITERATES = ["bare", "el"] as const;
export type EffectWholeIterate = (typeof EFFECT_WHOLE_ITERATES)[number];


function wholeIn(value: unknown, low: number, high: number): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= low && value <= high;
}

export function isEffectDimColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (/^#[0-9a-f]{6}$/.test(value) ||
      (EFFECT_DIM_SCHEME_COLORS as readonly string[]).includes(value))
  );
}

/**
 * The well-formed settings of a raw persisted value, or `undefined`. Lenient
 * the way the timeline's normalizers are: a malformed key is dropped, never
 * the step. Effect-specific canonical form (no default stated, no option the
 * effect does not take) is the authoring helpers' and the server's job.
 */
export function normalizeEffectSettings(
  raw: unknown,
): PresentationAnimationEffectSettings | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const out: PresentationAnimationEffectSettings = {};
  if (wholeIn(record.accel, 0, MAX_EFFECT_SHARE)) out.accel = record.accel;
  if (wholeIn(record.decel, 0, MAX_EFFECT_SHARE)) out.decel = record.decel;
  if ((out.accel ?? 0) + (out.decel ?? 0) > MAX_EFFECT_SHARE) {
    delete out.accel;
    delete out.decel;
  }
  if (record.autoReverse === true) out.autoReverse = true;
  if (
    (typeof record.repeat === "string" &&
      (EFFECT_REPEAT_UNTIL as readonly string[]).includes(record.repeat)) ||
    (wholeIn(record.repeat, 1, MAX_EFFECT_REPEAT_THOUSANDTHS) && record.repeat !== 1000)
  ) {
    out.repeat = record.repeat as number | EffectRepeatUntil;
  }
  if (record.rewind === true) out.rewind = true;
  if (
    typeof record.afterEffect === "string" &&
    (EFFECT_AFTER_KINDS as readonly string[]).includes(record.afterEffect)
  ) {
    if (record.afterEffect !== "dim") out.afterEffect = record.afterEffect as EffectAfterKind;
    else if (isEffectDimColor(record.dimColor)) {
      out.afterEffect = "dim";
      out.dimColor = record.dimColor;
    }
  }
  if (
    typeof record.textUnit === "string" &&
    (EFFECT_TEXT_UNITS as readonly string[]).includes(record.textUnit)
  ) {
    if (wholeIn(record.textUnitDelayPct, 0, MAX_EFFECT_SHARE) && record.textUnitDelayMs === undefined) {
      out.textUnit = record.textUnit as EffectTextUnit;
      out.textUnitDelayPct = record.textUnitDelayPct;
    } else if (
      wholeIn(record.textUnitDelayMs, 0, MAX_EFFECT_TEXT_DELAY_MS) &&
      record.textUnitDelayPct === undefined
    ) {
      out.textUnit = record.textUnit as EffectTextUnit;
      out.textUnitDelayMs = record.textUnitDelayMs;
    }
  }
  if (
    out.textUnit === undefined &&
    typeof record.wholeIterate === "string" &&
    (EFFECT_WHOLE_ITERATES as readonly string[]).includes(record.wholeIterate)
  ) {
    if (wholeIn(record.wholeIterateDelayPct, 0, MAX_EFFECT_SHARE) && record.wholeIterateDelayMs === undefined) {
      out.wholeIterate = record.wholeIterate as EffectWholeIterate;
      out.wholeIterateDelayPct = record.wholeIterateDelayPct;
    } else if (
      wholeIn(record.wholeIterateDelayMs, 0, MAX_EFFECT_TEXT_DELAY_MS) &&
      record.wholeIterateDelayPct === undefined
    ) {
      out.wholeIterate = record.wholeIterate as EffectWholeIterate;
      out.wholeIterateDelayMs = record.wholeIterateDelayMs;
    }
  }
  if (
    wholeIn(record.growSize, -GROW_SIZE_LIMIT - 1, GROW_SIZE_LIMIT) &&
    record.growSize !== DEFAULT_GROW_SIZE &&
    !Object.is(record.growSize, -0)
  ) {
    out.growSize = record.growSize;
  }
  if (
    typeof record.growDirection === "string" &&
    (EFFECT_GROW_DIRECTIONS as readonly string[]).includes(record.growDirection)
  ) {
    out.growDirection = record.growDirection as EffectGrowDirection;
  }
  if (record.flyStart === "offSlide") out.flyStart = "offSlide";
  // A key this release does not know is dropped with the rest of the
  // malformed values; the server refuses to write it either way.
  return Object.keys(out).length > 0 ? out : undefined;
}

export function effectSettingsEqual(
  left: PresentationAnimationEffectSettings | undefined,
  right: PresentationAnimationEffectSettings | undefined,
): boolean {
  if (left === right) return true;
  if (!left || !right) return false;
  const keys = new Set([...Object.keys(left), ...Object.keys(right)]);
  for (const key of keys) {
    if (
      (left as Record<string, unknown>)[key] !== (right as Record<string, unknown>)[key]
    ) {
      return false;
    }
  }
  return true;
}

/** What an effect does as the gallery adds it, for the dialog's defaults. */
export interface EffectSettingsDefaults {
  /** Smooth start share the effect has of its own (0..100000). */
  accel: number;
  /** Smooth end share the effect has of its own. */
  decel: number;
  /** The effect already plays text letter by letter (Whip, Wave …). */
  iterates: boolean;
  /** Rewind when done playing applies (the effect's node ends held). */
  rewindable: boolean;
  /** The effect takes a Size and Direction (Grow/Shrink). */
  sizable?: boolean;
  /** The effect may fly from a whole slide away (Fly In from the bottom,
   * left or right, row 1096). */
  fliesFromOffSlide?: boolean;
}

/**
 * The canonical form of `settings` for an effect with `defaults`: defaults
 * omitted, options the effect does not take dropped. `undefined` when
 * nothing is left.
 */
export function canonicalEffectSettings(
  settings: PresentationAnimationEffectSettings | undefined,
  defaults: EffectSettingsDefaults,
): PresentationAnimationEffectSettings | undefined {
  const normalized = normalizeEffectSettings(settings);
  if (!normalized) return undefined;
  const out: PresentationAnimationEffectSettings = { ...normalized };
  const accel = out.accel ?? defaults.accel;
  const decel = out.decel ?? defaults.decel;
  if (accel + decel > MAX_EFFECT_SHARE) return undefined;
  if (out.accel === defaults.accel) delete out.accel;
  if (out.decel === defaults.decel) delete out.decel;
  if (!defaults.rewindable) delete out.rewind;
  if (!defaults.sizable) {
    delete out.growSize;
    delete out.growDirection;
  }
  if (!defaults.fliesFromOffSlide) delete out.flyStart;
  if (defaults.iterates || out.textUnit !== undefined) {
    // An effect that iterates itself takes no iteration of the file's; By
    // word / By letter replaces a whole-object one.
    delete out.wholeIterate;
    delete out.wholeIterateDelayPct;
    delete out.wholeIterateDelayMs;
  }
  if (defaults.iterates) {
    delete out.textUnit;
    delete out.textUnitDelayPct;
    delete out.textUnitDelayMs;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * The scale a Grow/Shrink ends at under its Size and Direction, per axis
 * (1.5 = 150 %): the `p:by` the API writes, divided by 100000.
 */
export function effectGrowScale(
  settings: PresentationAnimationEffectSettings | undefined,
): { x: number; y: number } {
  const size = (settings?.growSize ?? DEFAULT_GROW_SIZE) / 100_000;
  if (settings?.growDirection === "horizontal") return { x: size, y: 1 };
  if (settings?.growDirection === "vertical") return { x: 1, y: size };
  return { x: size, y: size };
}

/** The Size preset a size is, or null for a custom amount. */
export function growSizePreset(size: number): GrowSizePreset | null {
  return GROW_SIZE_PRESETS.find((preset) => preset.size === size)?.id ?? null;
}

/**
 * `settings` with Grow/Shrink's Size and Direction replaced (the default
 * left unstated), or `undefined` when nothing is left.
 */
export function withGrowSettings(
  settings: PresentationAnimationEffectSettings | undefined,
  size: number,
  direction: EffectGrowDirection | "both",
): PresentationAnimationEffectSettings | undefined {
  const { growSize: _size, growDirection: _direction, ...rest } = settings ?? {};
  const out: PresentationAnimationEffectSettings = { ...rest };
  if (size !== DEFAULT_GROW_SIZE) out.growSize = size;
  if (direction !== "both") out.growDirection = direction;
  return Object.keys(out).length > 0 ? out : undefined;
}

/** How many times one play of the effect lasts (repeats x auto-reverse);
 * a repeat with no count counts once, as in the API's timing model. */
export function effectSettingsPlayFactor(
  settings: PresentationAnimationEffectSettings | undefined,
): number {
  if (!settings) return 1;
  const count = typeof settings.repeat === "number" ? settings.repeat / 1000 : 1;
  return count * (settings.autoReverse ? 2 : 1);
}

/**
 * The pieces of `text` By word steps through, white space between them kept
 * as pieces of its own (`word: false`): runs of letters and digits of any
 * script (a Chinese sentence with no spaces is one word) and every other
 * character that is not white space on its own -- the API's
 * `effect_text_words` (truth probe anim-effect-options, PowerPoint
 * 16.0.20527: "Hello, world! 123" is five words, "我们今天讨论动画" one).
 */
export function effectTextWordPieces(text: string): { text: string; word: boolean }[] {
  // PowerPoint's words (row 1310, the API's `effect_text_words`): an emoji with
  // its skin tone, variation selector and ZWJ-joined emoji is one word; a run
  // of letters and digits breaks where Latin meets Han, kana or Hangul; every
  // other character that is not white space is a word of its own.
  const pieces: { text: string; word: boolean }[] = [];
  const pattern =
    /(?:\p{Extended_Pictographic}(?:\p{Emoji_Modifier}|\uFE0F|\u200D\p{Extended_Pictographic})*)|[\p{sc=Han}\p{sc=Hiragana}\p{sc=Katakana}\p{sc=Hangul}][\p{sc=Han}\p{sc=Hiragana}\p{sc=Katakana}\p{sc=Hangul}\p{M}]*|(?:(?![\p{sc=Han}\p{sc=Hiragana}\p{sc=Katakana}\p{sc=Hangul}\p{Extended_Pictographic}])[\p{L}\p{M}\p{N}])+|\s+|[^\p{L}\p{M}\p{N}\s]/gu;
  for (const match of text.matchAll(pattern)) {
    pieces.push({ text: match[0], word: !/^\s/u.test(match[0]) });
  }
  return pieces;
}

/**
 * The words / letters PowerPoint steps through when it animates `text` By
 * word (`effectTextWordPieces`) or By letter (characters that are not white
 * space, punctuation included); the API's `effect_text_unit_count`.
 */
export function effectTextUnitCount(text: string, unit: EffectTextUnit): number {
  if (unit === "word") return effectTextWordPieces(text).filter((piece) => piece.word).length;
  let count = 0;
  for (const character of text) if (!/\s/u.test(character)) count += 1;
  return count;
}

/** The time between two units' starts for an effect of `durationMs`. */
export function effectTextUnitStepMs(
  settings: PresentationAnimationEffectSettings | undefined,
  durationMs: number,
): number | null {
  if (!settings?.textUnit) return null;
  if (typeof settings.textUnitDelayPct === "number") {
    return (settings.textUnitDelayPct / MAX_EFFECT_SHARE) * durationMs;
  }
  return settings.textUnitDelayMs ?? 0;
}

/**
 * How long a step plays in its click under its Effect Options: the duration
 * plus the delays between its words / letters (`unitCount`), times its
 * repeats and auto-reverse. Mirrors the API's `animation_timing_model._duration`
 * (whole milliseconds, rounded down) so the show and the written package
 * start the click's later runs at the same moment.
 */
export function effectSettingsPlayDurationMs(
  durationMs: number,
  settings: PresentationAnimationEffectSettings | undefined,
  unitCount: number,
): number {
  if (!settings || durationMs <= 0) return durationMs;
  let played = durationMs;
  if (settings.textUnit && unitCount > 1) {
    played +=
      typeof settings.textUnitDelayPct === "number"
        ? Math.floor(((unitCount - 1) * durationMs * settings.textUnitDelayPct) / MAX_EFFECT_SHARE)
        : (unitCount - 1) * (settings.textUnitDelayMs ?? 0);
  }
  const repeat = typeof settings.repeat === "number" ? settings.repeat : 1000;
  return Math.floor((played * repeat * (settings.autoReverse ? 2 : 1)) / 1000);
}

/**
 * The CSS easing of PowerPoint's Smooth start / Smooth end (ECMA-376 /
 * SMIL accelerate-decelerate, `lib/presentation-smooth-timing`) for an effect
 * drawn by two keyframes, from its `accel` / `decel` (100000 = all of it).
 * `null` for no easing.
 */
export function effectSmoothEasing(accel: number, decel: number): string | null {
  return smoothEasingCss(accel / MAX_EFFECT_SHARE, decel / MAX_EFFECT_SHARE);
}
