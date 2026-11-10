/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pen effects, one model for every editor that inks (Docs, the deck, the
 * worksheet) and every place a stroke is painted: the running stroke's
 * canvas (lib/office-ink-live), the landed SVG of each editor, the replay,
 * the PDF writers (the API's `ink_effects.py` twin) and the file writers.
 *
 * A stroke stores its effect in ONE attribute, `inkEffect`, as one of:
 *
 * - Office's eight (Draw > a pen's Effects: Rainbow, Galaxy, Lava, Ocean,
 *   Rose Gold, Gold, Silver, Bronze). Office keeps them as the InkML brush
 *   property `inkEffects` ([MS-ODRAWXML], Open XML SDK `InkEffectsType`:
 *   none, pencil, rainbow, galaxy, gold, silver, lava, ocean, rosegold,
 *   bronze); the set is closed and takes no colours.
 * - MRAI's own presets ({@link MRAI_INK_EFFECTS}).
 * - A custom effect the user built: `custom-<kind>-<rrggbb>[-<rrggbb>…]`, one
 *   to four colours and one of {@link INK_EFFECT_KINDS}.
 *
 * Every effect paints as LAYERS of the same stroke (no filters, nothing that
 * costs per pixel while the pen moves): the base layer wears a gradient along
 * the stroke's box, a neon's halo is the same stroke wider and translucent
 * under it and its core the stroke narrower and light over it, a metal's sheen
 * a narrow light line along it, a shimmer's glints small dots placed along the
 * stroke by arc length (so the running stroke and the landed one put each
 * glint in the same place). The outline colour (`outline` / `stroke`) holds
 * the effect's solid stand-in, which is what a file reader without effects
 * shows.
 */

import {
  inkStrokePathData,
  inkVariableWidthOutlinePathData,
  type InkPressurePoint,
} from "./office-ink-smooth";

/** Office's eight, in the order the pen menu shows them (MRAI's ids). */
export const OFFICE_INK_EFFECTS = [
  "rainbow",
  "galaxy",
  "lava",
  "ocean",
  "roseGold",
  "gold",
  "silver",
  "bronze",
] as const;
export type OfficeInkEffect = (typeof OFFICE_INK_EFFECTS)[number];

/** MRAI's own presets, in the order the pen menu shows them. */
export const MRAI_INK_EFFECTS = [
  "neonPink",
  "neonBlue",
  "neonLime",
  "ember",
  "glitterGold",
  "glitterRose",
  "starlight",
  "holographic",
  "chrome",
  "pearl",
  "emerald",
  "copper",
  "aurora",
  "pastel",
  "sunset",
  "candy",
] as const;
export type MraiInkEffect = (typeof MRAI_INK_EFFECTS)[number];

/**
 * Draw ▸ Tape's strips, in the order its menu shows them: a pen effect each
 * (a stroke of tape is a pen stroke wearing one), never offered in the pen
 * menu's effects.
 */
export const TAPE_INK_EFFECTS = ["tapeSky", "tapeRose", "tapeMint", "tapeSand"] as const;
export type TapeInkEffect = (typeof TAPE_INK_EFFECTS)[number];

/** How a custom effect paints (the presets are each one of these too). */
export const INK_EFFECT_KINDS = ["gradient", "metallic", "shimmer", "neon", "soft"] as const;
export type InkEffectKind = (typeof INK_EFFECT_KINDS)[number];

export type CustomInkEffect = `custom-${InkEffectKind}-${string}`;
export type InkEffectId = OfficeInkEffect | MraiInkEffect | TapeInkEffect | CustomInkEffect;

/** The longest stored effect (`custom-metallic-` and four colours is 44). */
export const INK_EFFECT_MAX_CHARS = 64;
export const INK_EFFECT_CUSTOM_MAX_COLORS = 4;
const CUSTOM_EFFECT = /^custom-(gradient|metallic|shimmer|neon|soft)((?:-[0-9a-f]{6}){1,4})$/;

/** The InkML `inkEffects` token Office writes for one of its eight. */
export type OfficeInkEffectsToken =
  | "rainbow"
  | "galaxy"
  | "lava"
  | "ocean"
  | "rosegold"
  | "gold"
  | "silver"
  | "bronze";

export interface InkEffectPaint {
  readonly id: InkEffectId;
  readonly kind: InkEffectKind;
  /** The base layer's colours, evenly along the gradient. */
  readonly stops: readonly string[];
  /** The one colour that stands in for the effect (outline attribute, file writers). */
  readonly solid: string;
  /** The colours repeat every this many pen widths (a metal's bands, candy stripes). */
  readonly period?: number;
  /** Each colour is a band with hard edges instead of a blend. */
  readonly hard?: boolean;
  /** A neon's or a soft pen's halo colour. */
  readonly glow?: string;
  /** A metal's sheen line or a neon's core. */
  readonly sheen?: string;
  /** A shimmer's glint colours. */
  readonly sparkle?: readonly string[];
  /** The Office effect that is this effect's closest faithful equivalent, if any. */
  readonly office: OfficeInkEffectsToken | null;
  /** The colours the user chose (a custom effect only). */
  readonly custom?: readonly string[];
}

/**
 * Office's eight. Their colours are MRAI's painting of them (Office paints
 * its own textures; it reads only the token), unchanged since the effects
 * were first drawn so existing strokes keep their look.
 */
const OFFICE_PAINTS: readonly InkEffectPaint[] = [
  { id: "rainbow", kind: "gradient", stops: ["#ff3b30", "#ff9500", "#ffcc00", "#34c759", "#007aff", "#5856d6", "#af52de"], solid: "#34c759", office: "rainbow" },
  { id: "galaxy", kind: "gradient", stops: ["#0b1a4a", "#3b2a8a", "#8a3fd1", "#2e8be6", "#0b1a4a"], solid: "#3b2a8a", office: "galaxy" },
  { id: "lava", kind: "gradient", stops: ["#3a0000", "#c81d11", "#ff7a00", "#ffd166", "#c81d11"], solid: "#c81d11", office: "lava" },
  { id: "ocean", kind: "gradient", stops: ["#013a63", "#0077b6", "#00b4d8", "#90e0ef", "#0077b6"], solid: "#0077b6", office: "ocean" },
  { id: "roseGold", kind: "gradient", stops: ["#b76e79", "#e8b4b8", "#f6d5d0", "#d89aa3", "#b76e79"], solid: "#c98a93", office: "rosegold" },
  { id: "gold", kind: "gradient", stops: ["#8a6d1f", "#d4af37", "#f9e79f", "#d4af37", "#8a6d1f"], solid: "#d4af37", office: "gold" },
  { id: "silver", kind: "gradient", stops: ["#6b6f75", "#c0c4c9", "#f2f4f6", "#c0c4c9", "#6b6f75"], solid: "#a8adb3", office: "silver" },
  { id: "bronze", kind: "gradient", stops: ["#5a3a1a", "#a5672b", "#d9a066", "#a5672b", "#5a3a1a"], solid: "#a5672b", office: "bronze" },
];

/**
 * MRAI's presets. `office` is set only where one of Office's eight is the
 * same family of look (a gold glitter is Office's Gold, a starry night its
 * Galaxy); the rest leave a DOCX as their solid colour.
 * The API keeps the same table (`app/services/ink_effect_presets.json`).
 */
const MRAI_PAINTS: readonly InkEffectPaint[] = [
  { id: "neonPink", kind: "neon", stops: ["#ff4fd8", "#ff2bd6", "#ff6af0"], glow: "#ff2bd6", sheen: "#fff0fb", solid: "#ff2bd6", office: null },
  { id: "neonBlue", kind: "neon", stops: ["#33f0ff", "#00c8ff", "#2f7bff"], glow: "#00b8ff", sheen: "#effcff", solid: "#00a6ee", office: null },
  { id: "neonLime", kind: "neon", stops: ["#d4ff3a", "#7dff2b", "#39ff14"], glow: "#4dff1f", sheen: "#f6ffe8", solid: "#3fd11a", office: null },
  { id: "ember", kind: "neon", stops: ["#ffd36b", "#ff8a1f", "#ff3b1f"], glow: "#ff5a1f", sheen: "#fff4dc", solid: "#ff5a1f", office: "lava" },
  { id: "glitterGold", kind: "shimmer", stops: ["#9a7212", "#e0b83c", "#fff1a8", "#d6a531", "#9a7212"], sparkle: ["#ffffff", "#fff6c8", "#ffe27a"], solid: "#d4a72c", office: "gold" },
  { id: "glitterRose", kind: "shimmer", stops: ["#c76b86", "#f2a7bf", "#ffe1ea", "#e48aa6", "#c76b86"], sparkle: ["#ffffff", "#ffe3ee", "#ffd1a8"], solid: "#e48aa6", office: "rosegold" },
  { id: "starlight", kind: "shimmer", stops: ["#0a1033", "#1d2a6b", "#3a2c8f", "#14306e", "#0a1033"], sparkle: ["#ffffff", "#cfe3ff", "#ffe8a3"], solid: "#1d2a6b", office: "galaxy" },
  { id: "holographic", kind: "metallic", stops: ["#ff9ad5", "#b39dff", "#7fd6ff", "#8cffd9", "#fff59a", "#ff9ad5"], period: 9, sheen: "#ffffff", solid: "#b39dff", office: "rainbow" },
  { id: "chrome", kind: "metallic", stops: ["#3c4652", "#c9d2dc", "#ffffff", "#7d8896", "#2e3640", "#aab4c0", "#3c4652"], period: 6, sheen: "#ffffff", solid: "#8a95a3", office: "silver" },
  { id: "pearl", kind: "metallic", stops: ["#a39bb6", "#efeaf5", "#dcb9c9", "#b6cde0", "#a39bb6"], period: 10, sheen: "#ffffff", solid: "#b3abc6", office: null },
  { id: "emerald", kind: "metallic", stops: ["#043d2a", "#0f8a5f", "#5be3a8", "#0f8a5f", "#043d2a"], period: 8, sheen: "#d9fff0", solid: "#0f8a5f", office: null },
  { id: "copper", kind: "metallic", stops: ["#5e2a12", "#b8622f", "#f3b48a", "#b8622f", "#5e2a12"], period: 8, sheen: "#ffe6d4", solid: "#b8622f", office: "bronze" },
  { id: "aurora", kind: "soft", stops: ["#18e6a0", "#13c6d6", "#6c5cff", "#c04dff", "#ff5fa8"], solid: "#13c6d6", office: null },
  { id: "pastel", kind: "soft", stops: ["#a68cff", "#ff9fb8", "#ffc285", "#7fd8b4", "#86b4ff"], solid: "#a68cff", office: null },
  { id: "sunset", kind: "gradient", stops: ["#ffcf5c", "#ff8a4c", "#ff4d6d", "#c13584", "#5b2a86"], solid: "#ff4d6d", office: null },
  { id: "candy", kind: "gradient", stops: ["#ff4f81", "#ffd1dc", "#36b5ff", "#d6f0ff"], period: 3, hard: true, solid: "#ff4f81", office: null },
];

/**
 * Draw ▸ Tape's strips: a light colour striped with a lighter one, two pen
 * widths a cycle, so a strip reads as paper tape. No Office equivalent (a
 * DOCX keeps the solid colour). The API keeps the same table.
 */
const TAPE_PAINTS: readonly InkEffectPaint[] = [
  { id: "tapeSky", kind: "gradient", stops: ["#8ec5ff", "#b9dcff"], period: 2, hard: true, solid: "#8ec5ff", office: null },
  { id: "tapeRose", kind: "gradient", stops: ["#ff9fbf", "#ffc8da"], period: 2, hard: true, solid: "#ff9fbf", office: null },
  { id: "tapeMint", kind: "gradient", stops: ["#7fdcb4", "#b5efd5"], period: 2, hard: true, solid: "#7fdcb4", office: null },
  { id: "tapeSand", kind: "gradient", stops: ["#f3c97a", "#f9e1ad"], period: 2, hard: true, solid: "#f3c97a", office: null },
];

export const OFFICE_INK_EFFECT_PAINTS: readonly InkEffectPaint[] = OFFICE_PAINTS;
export const MRAI_INK_EFFECT_PAINTS: readonly InkEffectPaint[] = MRAI_PAINTS;
export const TAPE_INK_EFFECT_PAINTS: readonly InkEffectPaint[] = TAPE_PAINTS;
const PRESET_PAINTS = new Map<string, InkEffectPaint>(
  [...OFFICE_PAINTS, ...MRAI_PAINTS, ...TAPE_PAINTS].map((paint) => [paint.id, paint]),
);

export function isTapeInkEffect(value: unknown): value is TapeInkEffect {
  return (TAPE_INK_EFFECTS as readonly unknown[]).includes(value);
}

/** Whether `value` names an effect a stroke or pen may carry. */
export function isInkEffect(value: unknown): value is InkEffectId {
  if (typeof value !== "string" || value.length > INK_EFFECT_MAX_CHARS) return false;
  return PRESET_PAINTS.has(value) || CUSTOM_EFFECT.test(value);
}

export function normalizeInkEffect(value: unknown): InkEffectId | undefined {
  return isInkEffect(value) ? value : undefined;
}

export function isOfficeInkEffect(value: unknown): value is OfficeInkEffect {
  return (OFFICE_INK_EFFECTS as readonly unknown[]).includes(value);
}

export function isCustomInkEffect(value: unknown): value is CustomInkEffect {
  return typeof value === "string" && CUSTOM_EFFECT.test(value);
}

/** A custom effect's kind and colours, or null for anything else. */
export function parseCustomInkEffect(
  value: unknown,
): { kind: InkEffectKind; colors: string[] } | null {
  if (typeof value !== "string") return null;
  const match = CUSTOM_EFFECT.exec(value);
  if (!match) return null;
  return {
    kind: match[1] as InkEffectKind,
    colors: match[2]!.slice(1).split("-").map((hex) => `#${hex}`),
  };
}

/** The stored id of a custom effect (colours `#rrggbb`, 1..4 of them), or null. */
export function customInkEffectId(kind: InkEffectKind, colors: readonly string[]): CustomInkEffect | null {
  if (!(INK_EFFECT_KINDS as readonly string[]).includes(kind)) return null;
  const hexes = colors
    .map((color) => (typeof color === "string" ? color.trim().toLowerCase() : ""))
    .filter((color) => /^#[0-9a-f]{6}$/.test(color))
    .slice(0, INK_EFFECT_CUSTOM_MAX_COLORS)
    .map((color) => color.slice(1));
  if (hexes.length === 0) return null;
  return `custom-${kind}-${hexes.join("-")}` as CustomInkEffect;
}

// ---- colour arithmetic (the API's twin does the same) -------------------

function channels(hex: string): [number, number, number] {
  const value = Number.parseInt(hex.slice(1), 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function hexOf([r, g, b]: readonly number[]): string {
  return `#${[r, g, b]
    .map((channel) => Math.round(Math.min(255, Math.max(0, channel!))).toString(16).padStart(2, "0"))
    .join("")}`;
}

/** `a` moved `share` of the way to `b`. */
export function mixInkColor(a: string, b: string, share: number): string {
  const from = channels(a);
  const to = channels(b);
  return hexOf(from.map((channel, index) => channel + (to[index]! - channel) * share));
}

const lighten = (color: string, share: number) => mixInkColor(color, "#ffffff", share);
const darken = (color: string, share: number) => mixInkColor(color, "#000000", share);

function customPaint(id: CustomInkEffect, kind: InkEffectKind, colors: readonly string[]): InkEffectPaint {
  const first = colors[0]!;
  const last = colors[colors.length - 1]!;
  const middle = colors[Math.floor((colors.length - 1) / 2)]!;
  const base = { id, kind, custom: colors, office: null } as const;
  switch (kind) {
    case "metallic": {
      // Each colour as a polished band: its shadow, itself, its highlight.
      const stops = colors.flatMap((color) => [darken(color, 0.45), color, lighten(color, 0.72), color]);
      return { ...base, stops: [...stops, darken(first, 0.45)], period: 8, sheen: lighten(first, 0.85), solid: middle };
    }
    case "shimmer":
      return {
        ...base,
        stops: colors.length === 1 ? [darken(first, 0.2), first, lighten(first, 0.45), first] : colors,
        sparkle: ["#ffffff", lighten(first, 0.7), lighten(last, 0.6)],
        solid: middle,
      };
    case "neon":
      return {
        ...base,
        stops: colors.length === 1 ? [lighten(first, 0.2), first, lighten(first, 0.2)] : colors,
        glow: first,
        sheen: lighten(first, 0.82),
        solid: first,
      };
    case "soft":
      return {
        ...base,
        stops: colors.length === 1 ? [lighten(first, 0.3), first, lighten(first, 0.3)] : colors,
        solid: middle,
      };
    default:
      return {
        ...base,
        stops: colors.length === 1 ? [darken(first, 0.18), first, lighten(first, 0.32)] : colors,
        solid: middle,
      };
  }
}

const customPaints = new Map<string, InkEffectPaint>();

/** How an effect paints; an unknown value paints as Office's first effect. */
export function inkEffectPaint(effect: InkEffectId | string): InkEffectPaint {
  const preset = PRESET_PAINTS.get(effect);
  if (preset) return preset;
  const cached = customPaints.get(effect);
  if (cached) return cached;
  const custom = parseCustomInkEffect(effect);
  if (!custom) return OFFICE_PAINTS[0]!;
  const paint = customPaint(effect as CustomInkEffect, custom.kind, custom.colors);
  if (customPaints.size > 256) customPaints.clear();
  customPaints.set(effect, paint);
  return paint;
}

/** The InkML `inkEffects` token a DOCX states for this effect (none: a plain colour). */
export function officeInkEffectsToken(effect: InkEffectId | string): OfficeInkEffectsToken | null {
  return isInkEffect(effect) ? inkEffectPaint(effect).office : null;
}

// ---- layers --------------------------------------------------------------

/**
 * One pass of the stroke: `base` wears the effect's gradient, a colour paints
 * itself; `width` is a share of the pen's width, `opacity` of the tool's.
 */
export interface InkEffectLayer {
  readonly role: "under" | "base" | "over";
  readonly paint: "base" | string;
  readonly width: number;
  readonly opacity: number;
}

const layerCache = new Map<string, readonly InkEffectLayer[]>();

export function inkEffectLayers(paint: InkEffectPaint): readonly InkEffectLayer[] {
  const cached = layerCache.get(paint.id);
  if (cached) return cached;
  const base: InkEffectLayer = { role: "base", paint: "base", width: 1, opacity: 1 };
  let layers: InkEffectLayer[];
  switch (paint.kind) {
    case "neon": {
      const glow = paint.glow ?? paint.solid;
      layers = [
        { role: "under", paint: glow, width: 3.1, opacity: 0.14 },
        { role: "under", paint: glow, width: 2, opacity: 0.3 },
        base,
        { role: "over", paint: paint.sheen ?? "#ffffff", width: 0.38, opacity: 0.92 },
      ];
      break;
    }
    case "metallic":
      layers = [base, { role: "over", paint: paint.sheen ?? "#ffffff", width: 0.3, opacity: 0.62 }];
      break;
    case "soft":
      layers = [
        { role: "under", paint: "base", width: 1.75, opacity: 0.26 },
        { ...base, opacity: 0.88 },
      ];
      break;
    default:
      layers = [base];
  }
  layerCache.set(paint.id, layers);
  return layers;
}

/** How far past the pen's width the widest layer reaches (a share of it). */
export function inkEffectReach(paint: InkEffectPaint): number {
  return Math.max(1, ...inkEffectLayers(paint).map((layer) => layer.width));
}

/** Whether the effect paints anything beyond its base gradient. */
export function inkEffectLayered(paint: InkEffectPaint): boolean {
  return inkEffectLayers(paint).length > 1 || Boolean(paint.sparkle?.length);
}

// ---- gradient ------------------------------------------------------------

export interface InkEffectGradientStop {
  readonly offset: number;
  readonly color: string;
}

export interface InkEffectGradient {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  readonly stops: readonly InkEffectGradientStop[];
}

/** The most colour stops one gradient states (a long striped stroke). */
const MAX_GRADIENT_STOPS = 240;

/** One cycle of the colours: each a band with hard edges, or a blend closing on its first. */
function cycleStops(paint: InkEffectPaint, repeating: boolean): InkEffectGradientStop[] {
  const colors = paint.stops;
  const count = colors.length;
  if (count === 1) return [{ offset: 0, color: colors[0]! }, { offset: 1, color: colors[0]! }];
  if (paint.hard) {
    return colors.flatMap((color, index) => [
      { offset: index / count, color },
      { offset: (index + 1) / count, color },
    ]);
  }
  const closing = repeating && colors[0] !== colors[count - 1];
  const sequence = closing ? [...colors, colors[0]!] : colors;
  return sequence.map((color, index) => ({ offset: index / (sequence.length - 1), color }));
}

/**
 * The base layer's gradient for a stroke whose box is `box` (user units, the
 * stroke's own), drawn `widthPx` wide, starting at `anchor` (its first
 * sample). A blend runs from the box's top-left to its bottom-right corner,
 * as Office's effects were always painted here; a repeating one (bands,
 * stripes) runs across the page from the first sample (handwriting travels
 * left to right, so every letter crosses the bands) with one cycle every
 * `period` pen widths, so the bands stay put while the box grows under the
 * pen and the landed stroke shows them where the running one did.
 */
export function inkEffectGradient(
  paint: InkEffectPaint,
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  widthPx: number,
  anchor?: { readonly x: number; readonly y: number } | null,
): InkEffectGradient {
  const width = Math.max(1e-3, box.width);
  const height = Math.max(1e-3, box.height);
  if (!paint.period) {
    return { x1: box.x, y1: box.y, x2: box.x + width, y2: box.y + height, stops: cycleStops(paint, false) };
  }
  const origin = anchor ?? { x: box.x, y: box.y };
  const cycle = Math.max(2, paint.period * Math.max(0.5, widthPx));
  const along = (x: number) => x - origin.x;
  // The box's corners along the axis, with the reach of the widest layer past them.
  const reach = widthPx * 2;
  const ends = [
    along(box.x - reach),
    along(box.x + width + reach),
  ];
  const one = cycleStops(paint, true);
  const firstCycle = Math.floor(Math.min(...ends) / cycle);
  let lastCycle = Math.ceil(Math.max(...ends) / cycle);
  const most = Math.max(1, Math.floor(MAX_GRADIENT_STOPS / one.length));
  if (lastCycle - firstCycle > most) {
    // Too long a stroke for one stop per band: the bands go on from the start.
    lastCycle = firstCycle + most;
  }
  if (lastCycle <= firstCycle) lastCycle = firstCycle + 1;
  const start = firstCycle * cycle;
  const span = (lastCycle - firstCycle) * cycle;
  const stops: InkEffectGradientStop[] = [];
  for (let index = firstCycle; index < lastCycle; index += 1) {
    for (const stop of one) {
      stops.push({ offset: ((index - firstCycle) + stop.offset) / (lastCycle - firstCycle), color: stop.color });
    }
  }
  return {
    x1: origin.x + start,
    y1: origin.y,
    x2: origin.x + start + span,
    y2: origin.y,
    stops,
  };
}

/** The effect as a CSS background (a pen's glyph, a swatch). */
export function inkEffectCssGradient(paint: InkEffectPaint, angle = 135): string {
  if (paint.hard) {
    // Three cycles of its bands, as a short stroke shows them.
    const bands = [...paint.stops, ...paint.stops, ...paint.stops];
    const stops = bands.flatMap((color, index) => [
      `${color} ${((index / bands.length) * 100).toFixed(2)}%`,
      `${color} ${(((index + 1) / bands.length) * 100).toFixed(2)}%`,
    ]);
    return `linear-gradient(${angle}deg, ${stops.join(", ")})`;
  }
  return `linear-gradient(${angle}deg, ${paint.stops.join(", ")})`;
}

// ---- glints --------------------------------------------------------------

export interface InkEffectSparkle {
  readonly x: number;
  readonly y: number;
  readonly r: number;
  /** Which of the effect's sparkle colours. */
  readonly tone: number;
}

/** A small integer hash: the same glint at the same place every time. */
function hash01(seed: number, salt: number): number {
  let value = (seed * 374_761_393 + salt * 668_265_263) | 0;
  value = Math.imul(value ^ (value >>> 13), 1_274_126_177);
  value ^= value >>> 16;
  return (value >>> 0) / 4_294_967_296;
}

/**
 * Glints along a stroke, placed by arc length from its first sample in units
 * of the pen's width: the k-th glint is where it is however much of the
 * stroke is drawn, in whatever units the stroke is measured (a running
 * stroke's host pixels, a landed one's design pixels). Feed samples as they
 * come; the glints found so far are `dots`.
 */
export class InkSparkleWalker {
  readonly dots: InkEffectSparkle[] = [];
  private skipped = 0;
  private last: { x: number; y: number } | null = null;
  private travelled = 0;
  private next = 0;
  private readonly spacing: number;

  constructor(private readonly widthPx: number) {
    this.spacing = Math.max(1.5, widthPx * 1.7);
    this.next = this.placeOf(0);
  }

  private placeOf(index: number): number {
    return (index + 0.3 + hash01(index, 1) * 0.6) * this.spacing;
  }

  push(points: readonly { readonly x: number; readonly y: number }[]): InkEffectSparkle[] {
    const found: InkEffectSparkle[] = [];
    for (const point of points) {
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
      const last = this.last;
      this.last = { x: point.x, y: point.y };
      if (!last) continue;
      const dx = point.x - last.x;
      const dy = point.y - last.y;
      const length = Math.hypot(dx, dy);
      if (length <= 0) continue;
      while (this.next <= this.travelled + length) {
        const index = this.dots.length + this.skipped;
        const share = (this.next - this.travelled) / length;
        const keep = hash01(index, 2) > 0.22;
        if (keep) {
          const side = (hash01(index, 3) - 0.5) * 0.7 * this.widthPx;
          const star = hash01(index, 4) > 0.8;
          const dot: InkEffectSparkle = {
            x: last.x + dx * share - (dy / length) * side,
            y: last.y + dy * share + (dx / length) * side,
            r: star ? Math.max(0.8, this.widthPx * 0.34) : Math.max(0.6, this.widthPx * (0.14 + hash01(index, 5) * 0.14)),
            tone: Math.floor(hash01(index, 6) * 3),
          };
          this.dots.push(dot);
          found.push(dot);
        } else {
          this.skipped += 1;
        }
        this.next = this.placeOf(index + 1);
      }
      this.travelled += length;
    }
    return found;
  }
}

export function inkEffectSparkles(
  points: readonly { readonly x: number; readonly y: number }[],
  widthPx: number,
): InkEffectSparkle[] {
  const walker = new InkSparkleWalker(widthPx);
  walker.push(points);
  return walker.dots;
}

/** One path of round dots (`tone` picks which glints), for SVG. */
export function inkSparklePathData(dots: readonly InkEffectSparkle[], tone: number, precision = 2): string {
  const round = (value: number) => Number(value.toFixed(precision));
  let d = "";
  for (const dot of dots) {
    if (dot.tone !== tone) continue;
    const r = round(dot.r);
    d += `M${round(dot.x - dot.r)} ${round(dot.y)}a${r} ${r} 0 1 0 ${round(2 * dot.r)} 0a${r} ${r} 0 1 0 ${round(-2 * dot.r)} 0Z`;
  }
  return d;
}

// ---- one stroke, painted -----------------------------------------------

export interface InkEffectStrokeGeometry {
  /** The samples, in the user units the paths are drawn in. */
  readonly points: readonly InkPressurePoint[];
  /** The pen's width in the same units. */
  readonly widthPx: number;
  /** Paint the pressure outline (the samples' pressures vary). */
  readonly pressured: boolean;
  readonly corner?: number;
  readonly flatStart?: boolean;
  readonly flatEnd?: boolean;
  /** The line's caps (a point eraser's cut piece ends flat: `butt`), default round. */
  readonly linecap?: "round" | "butt";
  /** A cut piece's round pen tips, stroked over its flat-ended line. */
  readonly tipsD?: string;
  /** The tool's opacity, multiplied into every layer. */
  readonly opacity?: number;
}

/** A path of the plan: filled, stroked, or both (a halo round a pressure outline). */
export interface InkEffectSvgPath {
  readonly role: InkEffectLayer["role"] | "sparkle";
  readonly d: string;
  readonly fill: string;
  readonly stroke: string;
  readonly strokeWidth?: number;
  readonly linecap?: "round" | "butt";
  readonly opacity: number;
}

export interface InkEffectSvgPlan {
  readonly gradientId: string;
  readonly gradient: InkEffectGradient;
  readonly paths: readonly InkEffectSvgPath[];
}

function samplesBox(points: readonly { x: number; y: number }[]) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of points) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  if (!Number.isFinite(minX)) return { x: 0, y: 0, width: 1, height: 1 };
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

/**
 * The SVG a stroke in `paint` is: its gradient (`gradientId`) and its paths
 * in drawing order. Every editor's landed stroke, the Docs running stroke and
 * the replay map this plan to elements one to one.
 */
export function inkEffectSvgPlan(
  paint: InkEffectPaint,
  geometry: InkEffectStrokeGeometry,
  gradientId: string,
): InkEffectSvgPlan {
  const { points, widthPx, pressured } = geometry;
  const toolOpacity = geometry.opacity ?? 1;
  // Always the samples' box, as the running stroke's canvas and the wet stroke
  // use: an editor's drawing box stands off the stroke (its pad, its least
  // size), and a gradient over that shifted the colours when a stroke landed.
  const gradient = inkEffectGradient(paint, samplesBox(points), widthPx, points[0] ?? null);
  const options = { corner: geometry.corner };
  let centre: string | null = null;
  const centreD = () => (centre ??= inkStrokePathData(points, options));
  const outline = pressured
    ? inkVariableWidthOutlinePathData(points, widthPx, {
        ...options,
        flatStart: geometry.flatStart,
        flatEnd: geometry.flatEnd,
      })
    : null;
  const paths: InkEffectSvgPath[] = [];
  for (const layer of inkEffectLayers(paint)) {
    const color = layer.paint === "base" ? `url(#${gradientId})` : layer.paint;
    const opacity = layer.opacity * toolOpacity;
    if (outline && layer.width === 1) {
      // The pen's own width follows its pressure: the outline, filled.
      paths.push({ role: layer.role, d: outline, fill: color, stroke: "none", opacity });
      continue;
    }
    // A halo or a sheen is the centre line at its share of the width (one
    // path: a translucent halo never darkens where it crosses itself).
    const strokeWidth = layer.width * widthPx;
    if (layer.role !== "over" && geometry.linecap === "butt" && geometry.tipsD) {
      // A cut piece: the line with flat ends, then the pen's round tips.
      paths.push({ role: layer.role, d: centreD(), fill: "none", stroke: color, strokeWidth, linecap: "butt", opacity });
      paths.push({ role: layer.role, d: geometry.tipsD, fill: "none", stroke: color, strokeWidth, linecap: "round", opacity });
      continue;
    }
    const linecap = layer.role === "over" ? "round" : geometry.linecap ?? "round";
    paths.push({ role: layer.role, d: centreD(), fill: "none", stroke: color, strokeWidth, linecap, opacity });
  }
  if (paint.sparkle?.length) {
    const dots = inkEffectSparkles(points, widthPx);
    paint.sparkle.forEach((color, tone) => {
      const d = inkSparklePathData(dots, tone);
      if (d) paths.push({ role: "sparkle", d, fill: color, stroke: "none", opacity: 0.95 * toolOpacity });
    });
  }
  return { gradientId, gradient, paths };
}

const SVG_NS = "http://www.w3.org/2000/svg";

/**
 * Append a plan to `host` (its paths) and `defs` (its gradient), for a
 * painter that builds its SVG by hand (the Docs node view and its running
 * stroke). `extra` goes on every path (pointer-events, test ids); `base`
 * on the base layer's paths only.
 */
export function appendInkEffectSvg(
  ownerDocument: Document,
  host: Element,
  defs: Element,
  plan: InkEffectSvgPlan,
  extra: Record<string, string> = {},
  base: Record<string, string> = {},
): void {
  const gradient = ownerDocument.createElementNS(SVG_NS, "linearGradient");
  gradient.setAttribute("id", plan.gradientId);
  gradient.setAttribute("gradientUnits", "userSpaceOnUse");
  gradient.setAttribute("x1", String(plan.gradient.x1));
  gradient.setAttribute("y1", String(plan.gradient.y1));
  gradient.setAttribute("x2", String(plan.gradient.x2));
  gradient.setAttribute("y2", String(plan.gradient.y2));
  for (const stop of plan.gradient.stops) {
    const element = ownerDocument.createElementNS(SVG_NS, "stop");
    element.setAttribute("offset", String(Number(stop.offset.toFixed(5))));
    element.setAttribute("stop-color", stop.color);
    gradient.append(element);
  }
  defs.append(gradient);
  for (const path of plan.paths) {
    const element = ownerDocument.createElementNS(SVG_NS, "path");
    element.setAttribute("d", path.d);
    element.setAttribute("fill", path.fill);
    element.setAttribute("stroke", path.stroke);
    if (path.strokeWidth !== undefined) element.setAttribute("stroke-width", String(path.strokeWidth));
    if (path.stroke !== "none") {
      element.setAttribute("stroke-linecap", path.linecap ?? "round");
      element.setAttribute("stroke-linejoin", "round");
    }
    if (path.opacity < 1) element.setAttribute("opacity", String(Number(path.opacity.toFixed(3))));
    element.setAttribute("data-ink-layer", path.role);
    for (const [name, value] of Object.entries(path.role === "base" ? { ...extra, ...base } : extra)) {
      element.setAttribute(name, value);
    }
    host.append(element);
  }
}

// ---- the running stroke's canvas -----------------------------------------

/** A gradient on a canvas, from a plan's gradient (in the context's user units). */
export function inkEffectCanvasGradient(context: CanvasRenderingContext2D, gradient: InkEffectGradient): CanvasGradient {
  const canvasGradient = context.createLinearGradient(gradient.x1, gradient.y1, gradient.x2, gradient.y2);
  for (const stop of gradient.stops) canvasGradient.addColorStop(Math.min(1, Math.max(0, stop.offset)), stop.color);
  return canvasGradient;
}

/**
 * The stroke whole on a canvas, layer by layer, as the plan paints it. Each
 * layer is one path, so a translucent halo never darkens where it crosses
 * itself. `sparkles` are the glints found so far (the caller keeps a walker
 * so a frame only adds the new ones).
 */
export function paintInkEffectOnCanvas(
  context: CanvasRenderingContext2D,
  paint: InkEffectPaint,
  geometry: InkEffectStrokeGeometry,
  sparkles: readonly InkEffectSparkle[],
): void {
  const { points, widthPx, pressured } = geometry;
  if (points.length === 0) return;
  const box = samplesBox(points);
  const gradient = inkEffectCanvasGradient(context, inkEffectGradient(paint, box, widthPx, points[0]));
  const options = { corner: geometry.corner };
  let centre: Path2D | null = null;
  const centrePath = () => (centre ??= new Path2D(inkStrokePathData(points, options)));
  const outline = pressured ? new Path2D(inkVariableWidthOutlinePathData(points, widthPx, options)) : null;
  context.save();
  context.lineCap = "round";
  context.lineJoin = "round";
  context.setLineDash([]);
  for (const layer of inkEffectLayers(paint)) {
    const color: string | CanvasGradient = layer.paint === "base" ? gradient : layer.paint;
    context.globalAlpha = layer.opacity;
    if (outline && layer.width === 1) {
      context.fillStyle = color;
      context.fill(outline);
      continue;
    }
    context.strokeStyle = color;
    context.lineWidth = Math.max(0.5, layer.width * widthPx);
    context.stroke(centrePath());
  }
  if (paint.sparkle?.length && sparkles.length) {
    context.globalAlpha = 0.95;
    paint.sparkle.forEach((color, tone) => {
      const dots = new Path2D();
      for (const dot of sparkles) {
        if (dot.tone !== tone) continue;
        dots.moveTo(dot.x + dot.r, dot.y);
        dots.arc(dot.x, dot.y, dot.r, 0, Math.PI * 2);
      }
      context.fillStyle = color;
      context.fill(dots);
    });
  }
  context.restore();
}
