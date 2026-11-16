/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  OFFICE_EFFECT_SCHEMES,
  OFFICE_THEMES,
  OFFICE_THEME_COLOR_SLOTS,
  OFFICE_THEME_FONT_SCRIPTS,
  type OfficeEffectScheme,
  type OfficeThemeColors,
  type OfficeThemeFontCollection,
  type OfficeThemeFontScript,
  type OfficeThemeLabel,
} from "../office-theme-catalog";
import { officeFamilyFromDisplayName } from "../office-fonts/localized-names";
import {
  isSpreadsheetThemePaletteId,
  paletteById,
  spreadsheetNativeThemePaletteId,
  spreadsheetThemePalettesEqual,
  themeGrid,
  tintColor,
  type SpreadsheetNativeThemeColors,
} from "./color-palette";
import type { SpreadsheetShadow } from "./shape-effects";
import {
  OFFICE_THEME_FONT_GALLERY,
  OFFICE_THEME_FONT_SETS,
  officeThemeFontSet,
  type OfficeThemeFontSet,
} from "../office-theme-fonts";

/**
 * Excel's Page Layout > Themes group: Themes, Colors, Fonts and Effects, all
 * Office's own (lib/office-theme-catalog). A workbook's own color, font or
 * effect scheme (one Office does not ship) is carried as read, shown as the
 * file has it and kept until a gallery choice replaces that one scheme.
 */

/** Effects MRAI offered before its themes were Office's: the Office theme's. */
const LEGACY_EFFECT_IDS = new Set(["subtle", "strong"]);

/** The Effects gallery, in order. */
export const SPREADSHEET_THEME_EFFECTS: readonly OfficeEffectScheme[] =
  OFFICE_EFFECT_SCHEMES;

export type SpreadsheetThemeEffectId = string;

const EFFECTS_BY_ID = new Map(
  SPREADSHEET_THEME_EFFECTS.map((effect) => [effect.id, effect]),
);

/** One theme font scheme as Excel draws it (shared with PowerPoint's Fonts). */
export type SpreadsheetThemeFontSet = OfficeThemeFontSet;

/** Every Office font scheme: the Fonts gallery in Excel's order, then each theme's own. */
export const SPREADSHEET_THEME_FONT_SETS: readonly SpreadsheetThemeFontSet[] =
  OFFICE_THEME_FONT_SETS;

/** The Fonts gallery, in Excel's order. */
export const SPREADSHEET_THEME_FONT_GALLERY: readonly SpreadsheetThemeFontSet[] =
  OFFICE_THEME_FONT_GALLERY;

export type SpreadsheetThemeFontSetId = string;

const FONT_SETS_BY_ID = new Map(
  SPREADSHEET_THEME_FONT_SETS.map((font) => [font.id, font]),
);

/** A workbook's own fonts (`nativeThemeFonts`). */
export interface SpreadsheetNativeThemeFontCollection extends OfficeThemeFontCollection {
  readonly cs: string;
}

export interface SpreadsheetNativeThemeFonts {
  readonly name: string;
  readonly major: SpreadsheetNativeThemeFontCollection;
  readonly minor: SpreadsheetNativeThemeFontCollection;
}

/**
 * A workbook's own effects (`nativeThemeEffects`): their name, and the scheme
 * itself when the page carries it. A colour or font scheme is a handful of
 * values, but an effect scheme is a block of DrawingML with no smaller
 * description, so the page holds it verbatim — that is what lets a workbook
 * written from scratch keep the effects the file was opened with.
 */
export interface SpreadsheetNativeThemeEffects {
  readonly name: string;
  readonly xml?: string;
}

/**
 * Create New Theme Fonts (Word's Design > Fonts and Excel's Page Layout >
 * Fonts > Customize Fonts..., PowerPoint's Variants > Fonts): the theme's own
 * font scheme under the name it is saved as, from the collections it starts
 * from (the theme's own, or a Fonts gallery entry's), with the heading and
 * body faces picked. An Asian text face is the collection's `a:ea` and its
 * `a:font script="Hans"`, the entry Office draws Chinese text from (as
 * PowerPoint's dialog writes it, `_pptx_set_theme_east_asian_face`); every
 * other script face is kept. null when a face or the name is not one a theme
 * can hold.
 */
export function customizedSpreadsheetNativeThemeFonts(
  start: {
    readonly major: OfficeThemeFontCollection & { readonly cs?: string };
    readonly minor: OfficeThemeFontCollection & { readonly cs?: string };
  },
  draft: {
    readonly headingLatin: string;
    readonly bodyLatin: string;
    readonly headingEastAsian?: string;
    readonly bodyEastAsian?: string;
    readonly name: string;
  },
): SpreadsheetNativeThemeFonts | null {
  const face = (value: string | undefined) => (typeof value === "string" ? value.trim() : "");
  const collection = (
    from: OfficeThemeFontCollection & { readonly cs?: string },
    latin: string,
    picked: string,
  ) => {
    // The Asian face the dialog showed is the collection's Chinese one (its
    // Hans entry, else its ea). Given back unchanged -- under its Chinese or
    // its Office name -- it is not a pick: the collection keeps its own
    // ea (an empty ea leaves Japanese and Korean text to their own entries).
    const shown = from.scripts.Hans || from.ea;
    const eastAsian =
      picked && officeFamilyFromDisplayName(picked) !== officeFamilyFromDisplayName(shown) ? picked : "";
    return {
      latin: latin || from.latin,
      ea: eastAsian || from.ea,
      cs: from.cs ?? "",
      scripts: { ...from.scripts, ...(eastAsian ? { Hans: eastAsian } : {}) },
    };
  };
  if (!face(draft.name)) return null;
  return normalizeSpreadsheetNativeThemeFonts({
    name: face(draft.name),
    major: collection(start.major, face(draft.headingLatin), face(draft.headingEastAsian)),
    minor: collection(start.minor, face(draft.bodyLatin), face(draft.bodyEastAsian)),
  });
}

export function spreadsheetNativeThemeFontSet(
  native: SpreadsheetNativeThemeFonts,
): SpreadsheetThemeFontSet {
  return officeThemeFontSet(
    "native",
    { en: native.name },
    native.name,
    false,
    native.major,
    native.minor,
  );
}

const MAX_THEME_NAME_LENGTH = 256;
const HEX_COLOR = /^#[0-9a-f]{6}$/;

function themeName(value: unknown): string | null {
  return typeof value === "string" && value.length <= MAX_THEME_NAME_LENGTH
    ? value
    : null;
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return (
    value !== null &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(value))
  );
}

function hasExactly(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const own = Object.keys(value);
  return own.length === keys.length && keys.every((key) => own.includes(key));
}

/** A `nativeThemeColors` value, normalized, or null (the reader's rules). */
export function normalizeSpreadsheetNativeThemeColors(
  value: unknown,
): SpreadsheetNativeThemeColors | null {
  if (!isPlainRecord(value) || !hasExactly(value, ["name", "colors"])) return null;
  const name = themeName(value.name);
  const colors = value.colors;
  if (name === null || !isPlainRecord(colors) || !hasExactly(colors, OFFICE_THEME_COLOR_SLOTS)) {
    return null;
  }
  const normalized: Record<string, string> = {};
  for (const slot of OFFICE_THEME_COLOR_SLOTS) {
    const color = colors[slot];
    if (typeof color !== "string" || !HEX_COLOR.test(color.toLowerCase())) return null;
    normalized[slot] = color.toLowerCase();
  }
  return { name, colors: normalized as unknown as OfficeThemeColors };
}

function normalizeFontCollection(
  value: unknown,
): SpreadsheetNativeThemeFontCollection | null {
  if (!isPlainRecord(value) || !hasExactly(value, ["latin", "ea", "cs", "scripts"])) {
    return null;
  }
  const faces = [value.latin, value.ea, value.cs].map(themeName);
  const scripts = value.scripts;
  if (faces.some((face) => face === null) || !isPlainRecord(scripts)) return null;
  const known = new Set<string>(OFFICE_THEME_FONT_SCRIPTS);
  if (Object.keys(scripts).some((script) => !known.has(script))) return null;
  const normalizedScripts: Partial<Record<OfficeThemeFontScript, string>> = {};
  for (const script of OFFICE_THEME_FONT_SCRIPTS) {
    if (!(script in scripts)) continue;
    const face = themeName(scripts[script]);
    if (face === null) return null;
    normalizedScripts[script] = face;
  }
  return {
    latin: faces[0] as string,
    ea: faces[1] as string,
    cs: faces[2] as string,
    scripts: normalizedScripts,
  };
}

/** A `nativeThemeFonts` value, normalized, or null (the reader's rules). */
export function normalizeSpreadsheetNativeThemeFonts(
  value: unknown,
): SpreadsheetNativeThemeFonts | null {
  if (!isPlainRecord(value) || !hasExactly(value, ["name", "major", "minor"])) return null;
  const name = themeName(value.name);
  const major = normalizeFontCollection(value.major);
  const minor = normalizeFontCollection(value.minor);
  return name === null || !major || !minor ? null : { name, major, minor };
}

/** An effect scheme is a few kilobytes of DrawingML; beyond this it is dropped. */
const MAX_THEME_EFFECTS_XML = 32_768;

/** A `nativeThemeEffects` value, normalized, or null. */
export function normalizeSpreadsheetNativeThemeEffects(
  value: unknown,
): SpreadsheetNativeThemeEffects | null {
  if (
    !isPlainRecord(value) ||
    !(hasExactly(value, ["name"]) || hasExactly(value, ["name", "xml"]))
  ) {
    return null;
  }
  const name = themeName(value.name);
  if (name === null) return null;
  if (!("xml" in value)) return { name };
  const xml = value.xml;
  if (
    typeof xml !== "string" ||
    xml.length > MAX_THEME_EFFECTS_XML ||
    !/^\s*<a:fmtScheme[\s>]/.test(xml)
  ) {
    return null;
  }
  return { name, xml };
}

/** The workbook theme: per scheme, an Office id or the workbook's own scheme. */
export interface SpreadsheetWorkbookThemeSelection {
  colorId: string;
  fontId: SpreadsheetThemeFontSetId;
  effectId: SpreadsheetThemeEffectId;
  /** The workbook's own colors; `colorId` is not in effect while present. */
  nativeColors?: SpreadsheetNativeThemeColors;
  nativeFonts?: SpreadsheetNativeThemeFonts;
  nativeEffects?: SpreadsheetNativeThemeEffects;
  /**
   * The workbook's own theme part, which MRAI cannot read: the selection
   * names no scheme, it stands for "the theme the file has" (its display
   * accents and whether a whole theme may replace it), so Undo can return
   * to it after a theme was applied.
   */
  preserved?: { accents?: readonly string[]; replaceable?: true };
}

export interface SpreadsheetWorkbookThemePreset {
  id: string;
  label: OfficeThemeLabel;
  selection: SpreadsheetWorkbookThemeSelection;
  tile: { readonly background: string; readonly decoration: string | null };
}

export const DEFAULT_SPREADSHEET_WORKBOOK_THEME = Object.freeze({
  colorId: "office",
  fontId: "office",
  effectId: "none",
}) satisfies SpreadsheetWorkbookThemeSelection;

/** The Themes gallery: Office's themes in Excel's order. */
export const SPREADSHEET_WORKBOOK_THEMES: readonly SpreadsheetWorkbookThemePreset[] =
  Object.freeze(
    OFFICE_THEMES.map((theme) =>
      Object.freeze({
        id: theme.id,
        label: theme.label,
        selection: Object.freeze({
          colorId: theme.colors,
          fontId: theme.fonts,
          effectId: theme.effects,
        }),
        tile: theme.tile,
      }),
    ),
  );

export const SPREADSHEET_THEME_FILE_TYPE = "mrnote.spreadsheet-theme.v1";

export function normalizeSpreadsheetThemeFontId(
  value: unknown,
): SpreadsheetThemeFontSetId {
  return typeof value === "string" && FONT_SETS_BY_ID.has(value)
    ? value
    : DEFAULT_SPREADSHEET_WORKBOOK_THEME.fontId;
}

export function normalizeSpreadsheetThemeEffectId(
  value: unknown,
): SpreadsheetThemeEffectId {
  return typeof value === "string" && EFFECTS_BY_ID.has(value) && !LEGACY_EFFECT_IDS.has(value)
    ? value
    : DEFAULT_SPREADSHEET_WORKBOOK_THEME.effectId;
}

export function spreadsheetThemeEffect(value: unknown): OfficeEffectScheme {
  return (
    EFFECTS_BY_ID.get(normalizeSpreadsheetThemeEffectId(value)) ??
    SPREADSHEET_THEME_EFFECTS[0]
  );
}

export function normalizeSpreadsheetWorkbookTheme(
  value: Partial<SpreadsheetWorkbookThemeSelection>,
): SpreadsheetWorkbookThemeSelection {
  const nativeColors = value.nativeColors
    ? normalizeSpreadsheetNativeThemeColors(value.nativeColors)
    : null;
  const nativeFonts = value.nativeFonts
    ? normalizeSpreadsheetNativeThemeFonts(value.nativeFonts)
    : null;
  const nativeEffects = value.nativeEffects
    ? normalizeSpreadsheetNativeThemeEffects(value.nativeEffects)
    : null;
  return {
    colorId: nativeColors
      ? "office"
      : isSpreadsheetThemePaletteId(value.colorId)
        ? value.colorId
        : "office",
    fontId: nativeFonts ? "office" : normalizeSpreadsheetThemeFontId(value.fontId),
    effectId: nativeEffects ? "none" : normalizeSpreadsheetThemeEffectId(value.effectId),
    ...(nativeColors ? { nativeColors } : {}),
    ...(nativeFonts ? { nativeFonts } : {}),
    ...(nativeEffects ? { nativeEffects } : {}),
    ...(value.preserved
      ? {
          preserved: {
            ...(Array.isArray(value.preserved.accents) &&
            value.preserved.accents.length === 6 &&
            value.preserved.accents.every((color) => /^#[0-9a-f]{6}$/i.test(color))
              ? { accents: value.preserved.accents.map((color) => color.toLowerCase()) }
              : {}),
            ...(value.preserved.replaceable === true ? { replaceable: true as const } : {}),
          },
        }
      : {}),
  };
}

/** The palette id the renderer uses (a workbook's own colors spell theirs). */
export function spreadsheetWorkbookThemePaletteId(
  selection: SpreadsheetWorkbookThemeSelection,
): string {
  return selection.nativeColors
    ? spreadsheetNativeThemePaletteId(selection.nativeColors)
    : selection.colorId;
}

/** The font scheme the renderer uses. */
export function spreadsheetWorkbookThemeFontSet(
  selection: SpreadsheetWorkbookThemeSelection,
): SpreadsheetThemeFontSet {
  return selection.nativeFonts
    ? spreadsheetNativeThemeFontSet(selection.nativeFonts)
    : spreadsheetThemeFontSet(selection.fontId);
}

/**
 * The font scheme as a Simplified Chinese Excel resolves it: a theme font is
 * the scheme's face for the language Office runs in, so the body font of the
 * Office theme is 等线 there -- the Font box reads "等线 (正文)" and digits are
 * drawn in it -- and Aptos Narrow elsewhere (XLS-A052; MEASURED in Excel for
 * Mac zh-CN on a blank workbook). A scheme with no Hans face is unchanged.
 */
export function spreadsheetThemeFontSetForLocale(
  fonts: SpreadsheetThemeFontSet,
  locale: string,
): SpreadsheetThemeFontSet {
  const tag = locale.toLowerCase();
  if (!tag.startsWith("zh") || /-(?:tw|hk|mo|hant)\b/.test(tag)) return fonts;
  const first = (face: string, family: string) => {
    const hans = face.trim();
    if (!hans) return family;
    const quoted = `"${hans.replace(/["'\\;{}()]/g, "")}"`;
    const rest = family
      .split(",")
      .map((part) => part.trim())
      .filter((part) => part && part.toLowerCase() !== quoted.toLowerCase());
    return [quoted, ...rest].join(", ");
  };
  const headingFamily = first(fonts.major.scripts.Hans ?? "", fonts.headingFamily);
  const bodyFamily = first(fonts.minor.scripts.Hans ?? "", fonts.bodyFamily);
  if (headingFamily === fonts.headingFamily && bodyFamily === fonts.bodyFamily) return fonts;
  return Object.freeze({ ...fonts, headingFamily, bodyFamily });
}

/** A UI whose Excel draws a theme-bound font in the slot's Latin face:
 * every language but Chinese, Japanese and Korean (nativeGridFontForLocale). */
export function nativeGridLatinLocale(locale: string): boolean {
  return !/^(?:zh|ja|ko)(?:[-_]|$)/i.test(locale);
}



















export function nativeGridFontForLocale<
  F extends { family: string; themeRole?: "major" | "minor"; bodyFamily?: string; headingFamily?: string },
>(font: F, fonts: SpreadsheetThemeFontSet, locale: string): F {
  const tag = locale.toLowerCase();
  if (!tag.startsWith("zh") || /-(?:tw|hk|mo|hant)\b/.test(tag)) {
    if (!nativeGridLatinLocale(tag) || !font.themeRole) return font;
    const latin = (font.themeRole === "major" ? fonts.headingLatin : fonts.bodyLatin)?.trim();
    if (!latin || latin === font.family) return font;
    return {
      ...font,
      family: latin,
      ...(fonts.bodyLatin?.trim() ? { bodyFamily: fonts.bodyLatin.trim() } : {}),
      ...(fonts.headingLatin?.trim() ? { headingFamily: fonts.headingLatin.trim() } : {}),
    };
  }
  const minor = fonts.minor.scripts.Hans?.trim() || undefined;
  const major = fonts.major.scripts.Hans?.trim() || undefined;
  if (!minor && !major) return font;
  const family =
    font.themeRole === "minor" ? (minor ?? font.family) : font.themeRole === "major" ? (major ?? font.family) : font.family;
  return {
    ...font,
    family,
    ...(minor ? { bodyFamily: minor } : {}),
    ...(major ? { headingFamily: major } : {}),
  };
}

/**
 * The identity of the workbook's font scheme for layout caches: its id, or
 * "native" for the workbook's own (which only its package can change).
 */
export function spreadsheetWorkbookThemeFontKey(
  selection: SpreadsheetWorkbookThemeSelection,
): string {
  return selection.nativeFonts ? "native" : selection.fontId;
}

/** Stable, portable MRAI theme file; this does not claim native THMX support. */
export function serializeSpreadsheetWorkbookThemeFile(
  selection: SpreadsheetWorkbookThemeSelection,
): string {
  return `${JSON.stringify(
    {
      type: SPREADSHEET_THEME_FILE_TYPE,
      selection: normalizeSpreadsheetWorkbookTheme(selection),
    },
    null,
    2,
  )}\n`;
}

/** Strict theme-file reader: unknown schemes fail closed instead of defaulting. */
export function parseSpreadsheetWorkbookThemeFile(
  source: string,
): SpreadsheetWorkbookThemeSelection | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(source);
  } catch {
    return null;
  }
  if (
    !isPlainRecord(parsed) ||
    parsed.type !== SPREADSHEET_THEME_FILE_TYPE ||
    !isPlainRecord(parsed.selection)
  ) {
    return null;
  }
  const { colorId, fontId, effectId, nativeColors, nativeFonts, nativeEffects } =
    parsed.selection;
  const selection: SpreadsheetWorkbookThemeSelection = {
    colorId: typeof colorId === "string" ? colorId : "",
    fontId: typeof fontId === "string" ? fontId : "",
    effectId: typeof effectId === "string" ? effectId : "",
  };
  if (nativeColors !== undefined) {
    const native = normalizeSpreadsheetNativeThemeColors(nativeColors);
    if (!native) return null;
    selection.nativeColors = native;
  } else if (!isSpreadsheetThemePaletteId(colorId)) {
    return null;
  }
  if (nativeFonts !== undefined) {
    const native = normalizeSpreadsheetNativeThemeFonts(nativeFonts);
    if (!native) return null;
    selection.nativeFonts = native;
  } else if (typeof fontId !== "string" || !FONT_SETS_BY_ID.has(fontId)) {
    return null;
  }
  if (nativeEffects !== undefined) {
    const native = normalizeSpreadsheetNativeThemeEffects(nativeEffects);
    if (!native) return null;
    selection.nativeEffects = native;
  } else if (typeof effectId !== "string" || !EFFECTS_BY_ID.has(effectId)) {
    return null;
  }
  return normalizeSpreadsheetWorkbookTheme(selection);
}

export function spreadsheetWorkbookThemePreset(
  value: unknown,
): SpreadsheetWorkbookThemePreset | null {
  return (
    SPREADSHEET_WORKBOOK_THEMES.find((theme) => theme.id === value) ?? null
  );
}

function fontKey(font: SpreadsheetThemeFontSet): string {
  // Script faces differ between Office versions; Latin and East Asian name it.
  return [
    font.major.latin.trim().toLowerCase(),
    font.major.ea.trim(),
    font.minor.latin.trim().toLowerCase(),
    font.minor.ea.trim(),
  ].join("\u0000");
}

/** Whether two font schemes are the same faces (a theme's own and a gallery entry can be). */
export function spreadsheetThemeFontSetsEqual(left: unknown, right: unknown): boolean {
  return fontKey(spreadsheetThemeFontSet(left)) === fontKey(spreadsheetThemeFontSet(right));
}

/** Whether two effect schemes are the same effects (Depth's are Office 2013 - 2022's). */
export function spreadsheetThemeEffectsEqual(left: unknown, right: unknown): boolean {
  return spreadsheetThemeEffect(left).signature === spreadsheetThemeEffect(right).signature;
}

/** The Office theme these schemes are, if they are one theme's. */
export function spreadsheetWorkbookThemePresetFor(
  selection: SpreadsheetWorkbookThemeSelection,
): SpreadsheetWorkbookThemePreset | null {
  const normalized = normalizeSpreadsheetWorkbookTheme(selection);
  if (normalized.nativeColors || normalized.nativeFonts || normalized.nativeEffects) {
    return null;
  }
  return (
    SPREADSHEET_WORKBOOK_THEMES.find(
      (theme) =>
        spreadsheetThemePalettesEqual(theme.selection.colorId, normalized.colorId) &&
        spreadsheetThemeFontSetsEqual(theme.selection.fontId, normalized.fontId) &&
        spreadsheetThemeEffectsEqual(theme.selection.effectId, normalized.effectId),
    ) ?? null
  );
}

export function spreadsheetThemeFontFamily(id: unknown): string {
  return spreadsheetThemeFontSet(id).bodyFamily;
}

export function spreadsheetThemeHeadingFontFamily(id: unknown): string {
  return spreadsheetThemeFontSet(id).headingFamily;
}

export function spreadsheetThemeFontSet(id: unknown): SpreadsheetThemeFontSet {
  return (
    FONT_SETS_BY_ID.get(normalizeSpreadsheetThemeFontId(id)) ??
    SPREADSHEET_THEME_FONT_SETS[0]
  );
}

/**
 * The Office scheme with these faces: the one the file names, else a Fonts
 * gallery entry, else a theme's own. Script faces are not compared: one
 * scheme's differ between Office versions and languages.
 */
export function spreadsheetThemeFontSetForNativeFaces(input: {
  readonly name?: string;
  readonly major: Pick<OfficeThemeFontCollection, "latin" | "ea">;
  readonly minor: Pick<OfficeThemeFontCollection, "latin" | "ea">;
}): SpreadsheetThemeFontSet | null {
  const key = [
    input.major.latin.trim().toLowerCase(),
    input.major.ea.trim(),
    input.minor.latin.trim().toLowerCase(),
    input.minor.ea.trim(),
  ].join("\u0000");
  const matches = SPREADSHEET_THEME_FONT_SETS.filter((font) => fontKey(font) === key);
  return (
    matches.find((font) => input.name !== undefined && font.schemeName === input.name) ??
    matches.find((font) => font.gallery) ??
    matches[0] ??
    null
  );
}

/** The Office effect scheme whose XML signature this is (named one first). */
export function spreadsheetThemeEffectForSignature(
  signature: string,
  name = "",
): OfficeEffectScheme | null {
  const matches = SPREADSHEET_THEME_EFFECTS.filter(
    (effect) => effect.signature === signature,
  );
  return (
    matches.find((effect) => effect.schemeName === name) ?? matches[0] ?? null
  );
}

/** Exact preview swatches shared by Themes and Colors galleries. */
export function spreadsheetThemePreviewColors(colorId: unknown): string[] {
  const palette = paletteById(colorId);
  const grid = themeGrid(palette);
  const accentStart = Math.max(0, grid.length - palette.accents.length);
  return [
    grid[accentStart]?.[0] ?? palette.accents[0],
    grid[accentStart + 1]?.[0] ?? palette.accents[1],
    grid[accentStart + 2]?.[0] ?? palette.accents[2],
    grid[accentStart + 3]?.[0] ?? palette.accents[3],
    grid[accentStart + 4]?.[0] ?? palette.accents[4],
    grid[accentStart + 5]?.[0] ?? palette.accents[5],
  ];
}

export interface SpreadsheetThemeShapeVisual {
  fill: string;
  stroke: string;
  strokeWidth: number;
  shadow?: SpreadsheetShadow;
}

const EMU_PER_PIXEL = 9525;

/**
 * Resolve the visual defaults for a shape created under a workbook theme.
 * Existing shapes retain their explicit visuals; only future insertions consume
 * this helper, matching Office's theme-vs-direct-formatting boundary. The
 * outline is the theme's second line style, as Office's default shape style
 * references it; that style references no effect, so no theme adds a shadow.
 */
export function spreadsheetThemeShapeVisual(
  selection: SpreadsheetWorkbookThemeSelection,
): SpreadsheetThemeShapeVisual {
  const normalized = normalizeSpreadsheetWorkbookTheme(selection);
  const accent =
    paletteById(spreadsheetWorkbookThemePaletteId(normalized)).accents[0] ?? "#156082";
  const effect = spreadsheetThemeEffect(normalized.nativeEffects ? "none" : normalized.effectId);
  const width = (effect.lines[1] ?? 19050) / EMU_PER_PIXEL;
  return {
    fill: tintColor(accent, 0.78),
    stroke: accent,
    strokeWidth: Math.max(1, Math.round(width * 4) / 4),
  };
}

function sameNative(left: unknown, right: unknown): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
}

export function spreadsheetWorkbookThemesEqual(
  left: SpreadsheetWorkbookThemeSelection,
  right: SpreadsheetWorkbookThemeSelection,
): boolean {
  const a = normalizeSpreadsheetWorkbookTheme(left);
  const b = normalizeSpreadsheetWorkbookTheme(right);
  if (a.preserved || b.preserved) {
    return (
      Boolean(a.preserved) === Boolean(b.preserved) &&
      a.preserved?.replaceable === b.preserved?.replaceable &&
      (a.preserved?.accents ?? []).join() === (b.preserved?.accents ?? []).join()
    );
  }
  return (
    a.colorId === b.colorId &&
    a.fontId === b.fontId &&
    a.effectId === b.effectId &&
    sameNative(a.nativeColors, b.nativeColors) &&
    sameNative(a.nativeFonts, b.nativeFonts) &&
    sameNative(a.nativeEffects, b.nativeEffects)
  );
}
