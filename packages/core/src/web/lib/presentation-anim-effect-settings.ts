/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
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
  const pieces: { text: string; word: boolean }[] = [];
  for (const match of text.matchAll(/[\p{L}\p{M}\p{N}]+|\s+|[^\p{L}\p{M}\p{N}\s]/gu)) {
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
 * SMIL accelerate-decelerate): the speed ramps up linearly over the `accel`
 * share, holds, and ramps down over the `decel` share; progress is the
 * integral, sampled into a `linear()` easing. `null` for no easing.
 */
export function effectSmoothEasing(accel: number, decel: number): string | null {
  const a = Math.max(0, Math.min(1, accel / MAX_EFFECT_SHARE));
  const d = Math.max(0, Math.min(1 - a, decel / MAX_EFFECT_SHARE));
  if (a === 0 && d === 0) return null;
  const peak = 1 / (1 - a / 2 - d / 2);
  const progress = (t: number): number => {
    if (t < a) return (peak * t * t) / (2 * a);
    if (t <= 1 - d) return peak * (t - a / 2);
    const tail = t - (1 - d);
    return peak * (t - a / 2 - (tail * tail) / (2 * d));
  };
  const samples: string[] = [];
  const steps = 20;
  for (let index = 0; index <= steps; index += 1) {
    const t = index / steps;
    const value = Math.round(Math.min(1, Math.max(0, progress(t))) * 10_000) / 10_000;
    samples.push(`${value} ${Math.round(t * 1000) / 10}%`);
  }
  return `linear(${samples.join(", ")})`;
}
