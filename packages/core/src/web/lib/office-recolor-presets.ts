/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  hlsToRgb,
  officeColorChannelsFromHex,
  officeColorHexFromChannels,
  officeTransformedHex,
  rgbToHls,
} from "./office-color-transform";
import type { OfficeThemeGridColors, OfficeThemeGridSlot } from "./office-theme-color-grid";

const ACCENTS = [1, 2, 3, 4, 5, 6] as const;

export const OFFICE_RECOLOR_PRESET_IDS = [
  "grayscale",
  "sepia",
  "washout",
  "blackWhite25",
  "blackWhite50",
  "blackWhite75",
  "text2Dark",
  ...ACCENTS.map((n) => `accent${n}Dark` as const),
  "background2Light",
  ...ACCENTS.map((n) => `accent${n}Light` as const),
] as const;
export type OfficeRecolorPresetId = (typeof OFFICE_RECOLOR_PRESET_IDS)[number];

/** The gallery as Office lays it out: three rows of seven, No Recolor first. */
export const OFFICE_RECOLOR_GALLERY_ROWS: readonly (readonly (OfficeRecolorPresetId | "none")[])[] = [
  ["none", "grayscale", "sepia", "washout", "blackWhite25", "blackWhite50", "blackWhite75"],
  ["text2Dark", ...ACCENTS.map((n) => `accent${n}Dark` as const)],
  ["background2Light", ...ACCENTS.map((n) => `accent${n}Light` as const)],
];

const PRESET_SET: ReadonlySet<string> = new Set(OFFICE_RECOLOR_PRESET_IDS);

export function isOfficeRecolorPreset(value: unknown): value is OfficeRecolorPresetId {
  return typeof value === "string" && PRESET_SET.has(value);
}

/** The tile's name as a `console-notebooks` key (PowerPoint's own wording). */
export function officeRecolorLabelKey(id: OfficeRecolorPresetId | "none"): string {
  return `office.recolor.${id}`;
}

/** Washout's luminance, as Corrections deltas (it is one). */
export const OFFICE_WASHOUT_CORRECTIONS = /* own data: OFFICE_WASHOUT_CORRECTIONS */ {"brightness":60,"contrast":-60};

/** How each Recolor tile derives its colours (the editors' own amounts). */
type RecolorRecipe = { readonly lighten: "tint" | "shade"; readonly amount: number; readonly satMod: number };
const RECOLOR_RECIPES: { readonly sepia: RecolorRecipe & { readonly base: string }; readonly dark: RecolorRecipe; readonly light: RecolorRecipe } = /* own data: RECOLOR_RECIPES */ {"sepia":{"base":"#d1b289","lighten":"tint","amount":40000,"satMod":150000},"dark":{"lighten":"tint","amount":50000,"satMod":250000},"light":{"lighten":"shade","amount":30000,"satMod":120000}};

/** The fallback theme colours (the open default theme) for a host without a theme. */
const OFFICE_THEME_FALLBACK: OfficeThemeGridColors = /* own data: OFFICE_THEME_FALLBACK */ {"bg1":"#ffffff","tx1":"#000000","bg2":"#eef1f4","tx2":"#1f2a37","accent1":"#2f6fd6","accent2":"#e07b39","accent3":"#3a9d6b","accent4":"#8c5cc9","accent5":"#d04b6b","accent6":"#2c9aa8"};

/**
 * `tint` / `shade` (linear light, `officeTransformedHex`) then `satMod`, as
 * PowerPoint draws a Recolor preset's colour: the saturation is multiplied
 * WITHOUT being held to 100 % and the RGB that results is clamped.
 */
function presetColor(hex: string, lighten: "tint" | "shade", amount: number, satMod: number): string | null {
  const base = officeTransformedHex(hex, [[lighten, amount]]);
  const channels = base ? officeColorChannelsFromHex(base) : null;
  if (!channels) return null;
  const [hue, lightness, saturation] = rgbToHls(channels[0], channels[1], channels[2]);
  return officeColorHexFromChannels(hlsToRgb(hue, lightness, saturation * (satMod / 100_000)));
}

/**
 * The two colours of a duotone preset, dark end first (`#rrggbb`), from the
 * deck's theme colours and the transforms the preset's XML states.
 */
export function officeRecolorDuotone(
  id: OfficeRecolorPresetId,
  theme?: OfficeThemeGridColors | null,
): readonly [string, string] | null {
  const colors = theme ?? OFFICE_THEME_FALLBACK;
  if (id === "grayscale") return ["#000000", "#ffffff"];
  if (id === "sepia") return ["#000000", presetColor(RECOLOR_RECIPES.sepia.base, RECOLOR_RECIPES.sepia.lighten, RECOLOR_RECIPES.sepia.amount, RECOLOR_RECIPES.sepia.satMod) ?? RECOLOR_RECIPES.sepia.base];
  const dark = /^(text2|accent[1-6])Dark$/.exec(id);
  if (dark) {
    const slot: OfficeThemeGridSlot = dark[1] === "text2" ? "tx2" : (dark[1] as OfficeThemeGridSlot);
    const color = presetColor(colors[slot], RECOLOR_RECIPES.dark.lighten, RECOLOR_RECIPES.dark.amount, RECOLOR_RECIPES.dark.satMod);
    return color ? ["#000000", color] : null;
  }
  const light = /^(background2|accent[1-6])Light$/.exec(id);
  if (light) {
    const slot: OfficeThemeGridSlot = light[1] === "background2" ? "bg2" : (light[1] as OfficeThemeGridSlot);
    const color = presetColor(colors[slot], RECOLOR_RECIPES.light.lighten, RECOLOR_RECIPES.light.amount, RECOLOR_RECIPES.light.satMod);
    return color ? [color, "#ffffff"] : null;
  }
  return null;
}

/**
 * Recolor > More Variations: a Dark or Light tile's duotone around a colour
 * no tile names (`#rrggbb`, the picked colour before the tile's own
 * transforms), dark end first -- the tiles' rule with that colour.
 */
export function officeRecolorVariationDuotone(
  variation: "dark" | "light",
  color: string,
): readonly [string, string] | null {
  if (!/^#[0-9a-fA-F]{6}$/.test(color)) return null;
  if (variation === "dark") {
    const end = presetColor(color, RECOLOR_RECIPES.dark.lighten, RECOLOR_RECIPES.dark.amount, RECOLOR_RECIPES.dark.satMod);
    return end ? ["#000000", end] : null;
  }
  const end = presetColor(color, RECOLOR_RECIPES.light.lighten, RECOLOR_RECIPES.light.amount, RECOLOR_RECIPES.light.satMod);
  return end ? [end, "#ffffff"] : null;
}

/** One SVG filter primitive of the drawing (sRGB values, 0..1). */
export type OfficeAdjustFilterStep =
  | { readonly type: "matrix"; readonly values: readonly number[] }
  | { readonly type: "discrete"; readonly table: readonly number[] }
  | { readonly type: "linear"; readonly slope: number; readonly intercept: number };

export interface OfficeAdjustInput {
  readonly recolor?: OfficeRecolorPresetId | null;
  /** A duotone no tile names (dark end first), drawn when `recolor` is none. */
  readonly duotone?: readonly [string, string] | null;
  /** Corrections, whole signed percentages (Office's -100..100). */
  readonly brightness?: number;
  readonly contrast?: number;
}

const LUMA = [0.2126, 0.7152, 0.0722] as const;

function channels(hex: string): [number, number, number] {
  const value = Number.parseInt(hex.slice(1), 16);
  return [((value >> 16) & 255) / 255, ((value >> 8) & 255) / 255, (value & 255) / 255];
}

const round = (value: number) => Math.round(value * 10_000) / 10_000;

/** Brightness b / Contrast c (fractions) as one per-channel linear map. */
export function officeLuminanceMap(brightness: number, contrast: number): { slope: number; intercept: number } {
  const c = Math.max(-1, Math.min(0.999, contrast));
  const slope = c > 0 ? 1 / (1 - c) : 1 + c;
  return { slope: round(slope), intercept: round(0.5 * (1 - slope) + (brightness * (1 + slope)) / 2) };
}

/**
 * The filter steps for a blip's Color then Corrections; empty when neither
 * changes the picture.
 */
export function officeAdjustFilterSteps(
  input: OfficeAdjustInput,
  theme?: OfficeThemeGridColors | null,
): OfficeAdjustFilterStep[] {
  const steps: OfficeAdjustFilterStep[] = [];
  const recolor = input.recolor && isOfficeRecolorPreset(input.recolor) ? input.recolor : null;
  const threshold = recolor ? /^blackWhite(\d\d)$/.exec(recolor) : null;
  if (threshold) {
    // Luma, then black below the threshold and white from it (equal bins).
    steps.push({ type: "matrix", values: [...LUMA, 0, 0, ...LUMA, 0, 0, ...LUMA, 0, 0, 0, 0, 0, 1, 0] });
    const bins: Record<string, number[]> = { "25": [0, 1, 1, 1], "50": [0, 1], "75": [0, 0, 0, 1] };
    steps.push({ type: "discrete", table: bins[threshold[1]] ?? [0, 1] });
  } else if ((recolor && recolor !== "washout") || (!recolor && input.duotone)) {
    const pair = recolor ? officeRecolorDuotone(recolor, theme) : (input.duotone ?? null);
    if (pair) {
      const [from, to] = [channels(pair[0]), channels(pair[1])];
      const rows: number[] = [];
      for (let channel = 0; channel < 3; channel += 1) {
        const span = to[channel] - from[channel];
        rows.push(...LUMA.map((weight) => round(weight * span)), 0, round(from[channel]));
      }
      steps.push({ type: "matrix", values: [...rows, 0, 0, 0, 1, 0] });
    }
  }
  const corrections =
    recolor === "washout"
      ? OFFICE_WASHOUT_CORRECTIONS
      : { brightness: input.brightness ?? 0, contrast: input.contrast ?? 0 };
  if (corrections.brightness || corrections.contrast) {
    steps.push({ type: "linear", ...officeLuminanceMap(corrections.brightness / 100, corrections.contrast / 100) });
  }
  return steps;
}
