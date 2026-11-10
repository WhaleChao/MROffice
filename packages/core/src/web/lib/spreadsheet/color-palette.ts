/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeThemePalette } from "../office-color-palette";
import { themeColumn as buildThemeColumn } from "../office-color-palette";
import {
  OFFICE_COLOR_SCHEMES,
  OFFICE_THEME_COLOR_SLOTS,
  type OfficeThemeColors,
  type OfficeThemeLabel,
} from "../office-theme-catalog";

export {
  MAX_RECENT_COLORS,
  STANDARD_COLORS,
  THEME_NEUTRALS,
  THEME_TINT_STEPS,
  normalizeHexColor,
  pushRecentColor,
  readableInkFor,
  themeColumn,
  tintColor,
} from "../office-color-palette";

export interface SpreadsheetThemePalette extends OfficeThemePalette {
  /** Office's name for the scheme ({@link officeThemeLabel}). */
  readonly name: OfficeThemeLabel;
  /** Listed in the Colors gallery (a theme's own colors are not). */
  readonly gallery: boolean;
  /** The scheme's name in its XML. */
  readonly schemeName: string;
  readonly dark1: string;
  readonly light1: string;
  readonly dark2: string;
  readonly light2: string;
  readonly hyperlink: string;
  readonly followedHyperlink: string;
  readonly colors: OfficeThemeColors;
}

function palette(
  id: string,
  name: OfficeThemeLabel,
  schemeName: string,
  gallery: boolean,
  colors: OfficeThemeColors,
): SpreadsheetThemePalette {
  return Object.freeze({
    id,
    label: id,
    name,
    schemeName,
    gallery,
    dark1: colors.dk1,
    light1: colors.lt1,
    dark2: colors.dk2,
    light2: colors.lt2,
    accents: Object.freeze([
      colors.accent1,
      colors.accent2,
      colors.accent3,
      colors.accent4,
      colors.accent5,
      colors.accent6,
    ]),
    hyperlink: colors.hlink,
    followedHyperlink: colors.folHlink,
    colors,
  });
}

/**
 * Every Office color scheme: the Colors gallery's entries in Excel's order
 * (`gallery`), then each theme's own colors, which applying that theme uses.
 */
export const SPREADSHEET_THEME_PALETTES: readonly SpreadsheetThemePalette[] =
  Object.freeze(
    OFFICE_COLOR_SCHEMES.map((scheme) =>
      palette(scheme.id, scheme.label, scheme.schemeName, scheme.gallery, scheme.colors),
    ),
  );

/** The Colors gallery, in Excel's order. */
export const SPREADSHEET_THEME_COLOR_GALLERY: readonly SpreadsheetThemePalette[] =
  Object.freeze(SPREADSHEET_THEME_PALETTES.filter((candidate) => candidate.gallery));

export type SpreadsheetThemePaletteId = string;

const PALETTES_BY_ID = new Map(
  SPREADSHEET_THEME_PALETTES.map((candidate) => [candidate.id, candidate]),
);

/** A workbook's own colors (`nativeThemeColors`). */
export interface SpreadsheetNativeThemeColors {
  readonly name: string;
  readonly colors: OfficeThemeColors;
}

const NATIVE_PALETTE_PREFIX = "native:";

/**
 * The render-only palette id of a workbook's own colors. It spells the
 * colors, so {@link paletteById} resolves it anywhere without shared state;
 * it is never persisted as `themeId`.
 */
export function spreadsheetNativeThemePaletteId(
  native: SpreadsheetNativeThemeColors,
): string {
  return (
    NATIVE_PALETTE_PREFIX +
    OFFICE_THEME_COLOR_SLOTS.map((slot) => native.colors[slot].slice(1)).join("")
  );
}

function nativePalette(id: string): SpreadsheetThemePalette | null {
  const spelled = id.slice(NATIVE_PALETTE_PREFIX.length);
  if (!/^[0-9a-f]{72}$/.test(spelled)) return null;
  const colors = Object.fromEntries(
    OFFICE_THEME_COLOR_SLOTS.map((slot, index) => [
      slot,
      `#${spelled.slice(index * 6, index * 6 + 6)}`,
    ]),
  ) as unknown as OfficeThemeColors;
  return palette(id, { en: "" }, "", false, colors);
}

export function paletteById(value: unknown): SpreadsheetThemePalette {
  if (typeof value === "string") {
    const known = PALETTES_BY_ID.get(value);
    if (known) return known;
    if (value.startsWith(NATIVE_PALETTE_PREFIX)) {
      const native = nativePalette(value);
      if (native) return native;
    }
  }
  return SPREADSHEET_THEME_PALETTES[0];
}

export function isSpreadsheetThemePaletteId(value: unknown): value is string {
  return typeof value === "string" && PALETTES_BY_ID.has(value);
}

/** The eight swatches Excel's Colors gallery shows for a scheme. */
export function spreadsheetThemePalettePreviewColors(
  value: unknown,
): readonly string[] {
  const selected = paletteById(value);
  return [selected.dark2, selected.light2, ...selected.accents];
}

function sameColors(left: OfficeThemeColors, right: OfficeThemeColors): boolean {
  return OFFICE_THEME_COLOR_SLOTS.every((slot) => left[slot] === right[slot]);
}

/**
 * The Office scheme with exactly these twelve colors: the one the file names,
 * else a Colors gallery entry, else a theme's own.
 */
export function spreadsheetThemePaletteForNativeColors(
  colors: OfficeThemeColors,
  name = "",
): SpreadsheetThemePalette | null {
  const matches = SPREADSHEET_THEME_PALETTES.filter((candidate) =>
    sameColors(candidate.colors, colors),
  );
  return (
    matches.find((candidate) => candidate.schemeName === name) ??
    matches.find((candidate) => candidate.gallery) ??
    matches[0] ??
    null
  );
}

/** Whether two palettes are the same colors (a theme's own and a gallery entry can be). */
export function spreadsheetThemePalettesEqual(left: unknown, right: unknown): boolean {
  return sameColors(paletteById(left).colors, paletteById(right).colors);
}

/** All ten native theme slots, each expanded into its tint/shade column. */
export function themeGrid(paletteValue: OfficeThemePalette): string[][] {
  const selected = paletteById(paletteValue.id);
  return [
    selected.light1,
    selected.dark1,
    selected.light2,
    selected.dark2,
    ...selected.accents,
  ].map((color) => buildThemeColumn(color));
}
