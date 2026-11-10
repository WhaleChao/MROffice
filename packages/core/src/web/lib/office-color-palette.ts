/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  OFFICE_STANDARD_COLORS,
  OFFICE_THEME_COLOR_COLUMNS,
} from "../components/office-ribbon/officeColorCatalog";

/**
 * The shared palette model behind colour controls across the Office editors.
 *
 * A native `<input type="color">` is useful for arbitrary colours, but it does
 * not provide the stable theme, standard and recent choices people repeatedly
 * use in a document. This module owns that common model so editor adapters and
 * shared renderers consume one ordered set and one normalization contract.
 */

export interface OfficeThemePalette {
  id: string;
  /** Translation key suffix; the caller resolves the label. */
  label: string;
  /** The six accent colours, in Office's slot order. */
  accents: readonly string[];
}

/** Theme palettes in the order presented by existing editor controls. */
export const OFFICE_THEME_PALETTES = /* own data: OFFICE_THEME_PALETTES */ [{"id":"office","label":"office","accents":["#2f6fd6","#e07b39","#3a9d6b","#8c5cc9","#d04b6b","#2c9aa8"]},{"id":"blue","label":"blue","accents":["#1c4069","#2b72a1","#4a91cf","#8ec5e1","#234785","#6e9cd8"]},{"id":"warm","label":"warm","accents":["#731811","#b1521b","#e1a937","#ec9b83","#922b16","#e7a35f"]},{"id":"green","label":"green","accents":["#246130","#379537","#74c157","#b1d897","#2d7b54","#84cd79"]},{"id":"grey","label":"grey","accents":["#494d50","#5b5e62","#6c7075","#92969b","#aaadb1","#bfc2c4"]}];

export type OfficeThemePaletteId =
  (typeof OFFICE_THEME_PALETTES)[number]["id"];

export function normalizeOfficeThemePaletteId(
  value: unknown,
): OfficeThemePaletteId | null {
  return typeof value === "string" &&
    OFFICE_THEME_PALETTES.some((palette) => palette.id === value)
    ? (value as OfficeThemePaletteId)
    : null;
}

/** The two neutral slots every generated theme carries before its accents. */
export const THEME_NEUTRALS = ["#ffffff", "#000000"] as const;

/** Lighter 80/60/40, then darker 25/50, in the existing grid order. */
export const THEME_TINT_STEPS = [0.8, 0.6, 0.4, -0.25, -0.5] as const;

/** The standard-colour row shared by every theme. */
export const STANDARD_COLORS = OFFICE_STANDARD_COLORS;

export const MAX_RECENT_COLORS = 10;

const HEX_RE = /^#[0-9a-f]{6}$/i;

/** A #rrggbb hex, lower-cased, or null when the value is not one. */
export function normalizeHexColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const text = value.trim();
  if (HEX_RE.test(text)) return text.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(text)) {
    const [, r, g, b] = text;
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  return null;
}

function channels(hex: string): [number, number, number] {
  return [
    Number.parseInt(hex.slice(1, 3), 16),
    Number.parseInt(hex.slice(3, 5), 16),
    Number.parseInt(hex.slice(5, 7), 16),
  ];
}

function toHex(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value)))
    .toString(16)
    .padStart(2, "0");
}

/** Mix toward white for positive tints and toward black for negative tints. */
export function tintColor(hex: string, tint: number): string {
  const base = normalizeHexColor(hex);
  if (!base) return "#000000";
  const amount = Math.max(-1, Math.min(1, tint));
  const [r, g, b] = channels(base);
  const mix = (channel: number): number =>
    amount >= 0 ? channel + (255 - channel) * amount : channel * (1 + amount);
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
}

/**
 * The five tints/shades Office lists under one theme colour, by that colour's
 * HSL luminance: white is Darker 5/15/25/35/50%, black Lighter 50/35/25/15/5%,
 * a very dark colour (under 20%) Lighter 90/75/50/25/10%, a very light one
 * (over 80%) Darker 10/25/50/75/90%, and every other colour Lighter 80/60/40%
 * then Darker 25/50%. Positive is lighter, negative darker, as SpreadsheetML's
 * `tint` attribute.
 */
export function officeThemeTintSteps(hex: string): readonly number[] {
  const [r, g, b] = channels(normalizeHexColor(hex) ?? "#000000");
  const luminance = (Math.max(r, g, b) + Math.min(r, g, b)) / 510;
  if (luminance >= 1) return [-0.05, -0.15, -0.25, -0.35, -0.5];
  if (luminance <= 0) return [0.5, 0.35, 0.25, 0.15, 0.05];
  if (luminance < 0.2) return [0.9, 0.75, 0.5, 0.25, 0.1];
  if (luminance > 0.8) return [-0.1, -0.25, -0.5, -0.75, -0.9];
  return THEME_TINT_STEPS;
}

/** One generated theme column: the colour itself, then its five tints. */
export function themeColumn(hex: string): string[] {
  return [
    tintColor(hex, 0),
    ...THEME_TINT_STEPS.map((step) => tintColor(hex, step)),
  ];
}

/** The complete theme grid, preserving the curated default Office columns. */
export function themeGrid(palette: OfficeThemePalette): string[][] {
  if (palette.id === "office") {
    return OFFICE_THEME_COLOR_COLUMNS.map((column) => [...column]);
  }
  return [...THEME_NEUTRALS, ...palette.accents].map((hex) => themeColumn(hex));
}

export function paletteById(id: unknown): OfficeThemePalette {
  const match = OFFICE_THEME_PALETTES.find((palette) => palette.id === id);
  return match ?? OFFICE_THEME_PALETTES[0];
}

/** Normalize, de-duplicate and move a colour to the head of recent choices. */
export function pushRecentColor(
  recents: readonly string[],
  color: unknown,
): string[] {
  const normalized = normalizeHexColor(color);
  if (!normalized) return [...recents];
  const existing = recents
    .map((value) => normalizeHexColor(value))
    .filter((value): value is string => value !== null);
  if (existing[0] === normalized) return existing;
  return [
    normalized,
    ...existing.filter((value) => value !== normalized),
  ].slice(0, MAX_RECENT_COLORS);
}

/** Choose white or black ink using the existing WCAG luminance threshold. */
export function readableInkFor(hex: string): "#ffffff" | "#000000" {
  const base = normalizeHexColor(hex) ?? "#ffffff";
  const [r, g, b] = channels(base).map((channel) => {
    const ratio = channel / 255;
    return ratio <= 0.03928 ? ratio / 12.92 : ((ratio + 0.055) / 1.055) ** 2.4;
  });
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luminance > 0.179 ? "#000000" : "#ffffff";
}
