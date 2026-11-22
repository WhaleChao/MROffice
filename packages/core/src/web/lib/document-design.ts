/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A Docs document's design, as Word's Design tab sets it: the document theme
 * (Colors, Fonts and Effects schemes, each one of Office's), the style set
 * (Document Formatting) and the Paragraph Spacing choice.
 *
 * Word keeps the theme in `word/theme/theme1.xml` and the style set's result
 * in `styles.xml`; the editor keeps the choices themselves
 * (`doc.attrs.documentDesign`) and resolves what they mean here, against the
 * shared Office catalog (`office-theme-catalog`) and Word's measured style
 * sets (`office-word-style-sets`):
 *
 * - the document's base (Normal + `w:docDefaults`): the editor's document
 *   default attributes;
 * - each Quick Style's own formatting, with the theme's fonts and colours
 *   resolved, for the Home tab's Styles gallery to write and for a design
 *   change to restyle the paragraphs that wear it;
 * - the theme colours every theme-aware colour reads.
 *
 * Word's Office theme differs from Excel's in its body font (Aptos, not
 * Aptos Narrow); everything else is the catalog's.
 */

import { documentFontSizePx } from "./document-font-size";
import {
  OFFICE_COLOR_SCHEMES,
  OFFICE_EFFECT_SCHEMES,
  OFFICE_FONT_SCHEMES,
  OFFICE_THEMES,
  type OfficeColorScheme,
  type OfficeFontScheme,
  type OfficeThemeColorSlot,
  type OfficeThemeColors,
  type OfficeThemeFontCollection,
} from "./office-theme-catalog";
import { officeThemeFontSet, type OfficeThemeFontSet } from "./office-theme-fonts";
import type { OfficeThemeFileTheme } from "@/lib/office-theme-file";
import type { SpreadsheetNativeThemeColors } from "./spreadsheet/color-palette";
import {
  normalizeSpreadsheetNativeThemeColors,
  normalizeSpreadsheetNativeThemeEffects,
  normalizeSpreadsheetNativeThemeFonts,
  type SpreadsheetNativeThemeEffects,
  type SpreadsheetNativeThemeFonts,
} from "./spreadsheet/workbook-theme";
import {
  documentParagraphSpacingPreset,
  type DocumentParagraphSpacingPresetId,
} from "./document-paragraph-spacing";
import { hlsToRgb, rgbToHls } from "./office-color-transform";
import type { DocumentBorderLine, DocumentParagraphBorders } from "./document-borders";
import {
  WORD_DEFAULT_STYLE_SET_ID,
  isWordStyleSetId,
  wordStyleChain,
  wordStyleSet,
  wordStyleSetBase,
  type WordQuickStyleId,
  type WordStyleBorder,
  type WordStyleColor,
  type WordStyleParagraphProps,
  type WordStyleRunProps,
  type WordStyleSet,
} from "./office-word-style-sets";

export interface DocumentDesign {
  /** The Office theme last applied (Word's theme part name), or null once a scheme changed alone. */
  readonly theme: string | null;
  readonly colors: string;
  readonly fonts: string;
  readonly effects: string;
  /** A Word style set id (`office-word-style-sets`). */
  readonly styleSet: string;
  /** A Paragraph Spacing preset; null keeps the style set's own spacing ("Default"). */
  readonly paragraphSpacing: Exclude<DocumentParagraphSpacingPresetId, "default"> | null;
  /**
   * A theme of the document's own (an imported document's, a browsed
   * theme file's): each scheme no Office scheme is the same as, kept as it
   * is (the shapes the spreadsheet keeps), standing in for the id beside it,
   * and the theme's name. Never with an Office `theme`.
   */
  readonly nativeColors?: SpreadsheetNativeThemeColors;
  readonly nativeFonts?: SpreadsheetNativeThemeFonts;
  readonly nativeEffects?: SpreadsheetNativeThemeEffects;
  readonly themeName?: string;
}

/** Word's new document: the Office theme, the Normal template's styles. */
export const DEFAULT_DOCUMENT_DESIGN: DocumentDesign = Object.freeze({
  theme: "office",
  colors: "office",
  fonts: "office",
  effects: "none",
  styleSet: WORD_DEFAULT_STYLE_SET_ID,
  paragraphSpacing: null,
});

const COLOR_IDS = new Set(OFFICE_COLOR_SCHEMES.map((scheme) => scheme.id));
const FONT_IDS = new Set(OFFICE_FONT_SCHEMES.map((scheme) => scheme.id));
const EFFECT_IDS = new Set(OFFICE_EFFECT_SCHEMES.map((scheme) => scheme.id));
const THEMES_BY_ID = new Map(OFFICE_THEMES.map((theme) => [theme.id, theme]));

function spacingPreset(
  value: unknown,
): DocumentDesign["paragraphSpacing"] {
  const preset = documentParagraphSpacingPreset(value);
  return preset && preset.id !== "default" ? preset.id : null;
}

/** `doc.attrs.documentDesign`, or null for a document the Design tab never touched. */
export function normalizeDocumentDesign(value: unknown): DocumentDesign | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const colors = typeof record.colors === "string" && COLOR_IDS.has(record.colors) ? record.colors : null;
  const fonts = typeof record.fonts === "string" && FONT_IDS.has(record.fonts) ? record.fonts : null;
  const effects =
    typeof record.effects === "string" && EFFECT_IDS.has(record.effects) ? record.effects : null;
  if (!colors || !fonts || !effects) return null;
  const theme =
    typeof record.theme === "string" && THEMES_BY_ID.has(record.theme) ? record.theme : null;
  const styleSet = isWordStyleSetId(record.styleSet) ? record.styleSet : WORD_DEFAULT_STYLE_SET_ID;
  const own: {
    nativeColors?: SpreadsheetNativeThemeColors;
    nativeFonts?: SpreadsheetNativeThemeFonts;
    nativeEffects?: SpreadsheetNativeThemeEffects;
    themeName?: string;
  } = {};
  if (record.nativeColors !== undefined) {
    const native = normalizeSpreadsheetNativeThemeColors(record.nativeColors);
    if (!native) return null;
    own.nativeColors = native;
  }
  if (record.nativeFonts !== undefined) {
    const native = normalizeSpreadsheetNativeThemeFonts(record.nativeFonts);
    if (!native) return null;
    own.nativeFonts = native;
  }
  if (record.nativeEffects !== undefined) {
    const native = normalizeSpreadsheetNativeThemeEffects(record.nativeEffects);
    if (!native) return null;
    own.nativeEffects = native;
  }
  if (record.themeName !== undefined) {
    const name = record.themeName;
    if (typeof name !== "string" || !name.trim() || name.length > 256 || /[\u0000-\u001f\u007f]/.test(name)) {
      return null;
    }
    own.themeName = name;
  }
  // An Office theme is its own three schemes (the API refuses the mix too).
  if (theme && (own.nativeColors || own.nativeFonts || own.nativeEffects)) return null;
  return Object.freeze({
    theme,
    colors,
    fonts,
    effects,
    styleSet,
    paragraphSpacing: spacingPreset(record.paragraphSpacing),
    ...own,
  });
}

function hasOwnSchemes(design: DocumentDesign): boolean {
  return Boolean(design.nativeColors || design.nativeFonts || design.nativeEffects);
}

export function documentDesignEquals(a: DocumentDesign | null, b: DocumentDesign | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (
    a.theme === b.theme &&
    a.colors === b.colors &&
    a.fonts === b.fonts &&
    a.effects === b.effects &&
    a.styleSet === b.styleSet &&
    a.paragraphSpacing === b.paragraphSpacing &&
    a.themeName === b.themeName &&
    JSON.stringify(a.nativeColors ?? null) === JSON.stringify(b.nativeColors ?? null) &&
    JSON.stringify(a.nativeFonts ?? null) === JSON.stringify(b.nativeFonts ?? null) &&
    JSON.stringify(a.nativeEffects ?? null) === JSON.stringify(b.nativeEffects ?? null)
  );
}

/** A design without the document's own schemes or theme name (an Office scheme was chosen). */
function withoutOwn(
  design: DocumentDesign,
  keys: readonly ("nativeColors" | "nativeFonts" | "nativeEffects" | "themeName")[],
): DocumentDesign {
  const next: Record<string, unknown> = { ...design };
  for (const key of keys) delete next[key];
  return next as unknown as DocumentDesign;
}

/** Design > Colors / Fonts / Effects: an Office scheme in place of the document's own. */
export function documentDesignWithScheme(
  design: DocumentDesign,
  axis: "colors" | "fonts" | "effects",
  id: string,
): DocumentDesign {
  const native = axis === "colors" ? "nativeColors" : axis === "fonts" ? "nativeFonts" : "nativeEffects";
  return Object.freeze({ ...withoutOwn(design, [native]), [axis]: id }) as DocumentDesign;
}

/**
 * Design > Fonts > Customize Fonts... (Create New Theme Fonts): a font
 * scheme of the document's own in place of the Office one, under the name it
 * was saved as. The theme is a custom combination from then on (an Office
 * theme is its own three schemes), its colours and effects kept.
 */
export function documentDesignWithCustomFonts(
  design: DocumentDesign,
  fonts: SpreadsheetNativeThemeFonts,
): DocumentDesign | null {
  return normalizeDocumentDesign({ ...design, theme: null, nativeFonts: fonts });
}

/**
 * Browse for Themes: the theme a file carries, in place of the document's
 * three schemes (Word applies a theme file whole). An Office theme by id;
 * otherwise each scheme Office has by id and the others as the file has
 * them, under the file's theme name.
 */
export function documentDesignWithThemeFile(
  design: DocumentDesign,
  file: OfficeThemeFileTheme,
): DocumentDesign | null {
  if (file.theme) return documentDesignWithTheme(design, file.theme);
  const next = {
    ...withoutOwn(design, ["nativeColors", "nativeFonts", "nativeEffects", "themeName"]),
    theme: null,
    colors: "id" in file.colors ? file.colors.id : "office",
    fonts: "id" in file.fonts ? file.fonts.id : "office",
    effects: "id" in file.effects ? file.effects.id : "none",
    ...("native" in file.colors ? { nativeColors: file.colors.native } : {}),
    ...("native" in file.fonts ? { nativeFonts: file.fonts.native } : {}),
    ...("native" in file.effects ? { nativeEffects: file.effects.native } : {}),
    ...(file.name ? { themeName: file.name.slice(0, 256) } : {}),
  };
  return normalizeDocumentDesign(next);
}

/** Applying an Office theme: its three schemes, under its name. */
export function documentDesignWithTheme(design: DocumentDesign, themeId: string): DocumentDesign | null {
  const theme = THEMES_BY_ID.get(themeId);
  if (!theme) return null;
  return Object.freeze({
    ...withoutOwn(design, ["nativeColors", "nativeFonts", "nativeEffects", "themeName"]),
    theme: theme.id,
    colors: theme.colors,
    fonts: theme.fonts,
    effects: theme.effects,
  });
}

/**
 * The theme whose three schemes these are, which the Themes gallery shows
 * as current: the applied theme while it still holds, else any theme that
 * matches, else none (a custom combination).
 */
export function documentDesignThemeId(design: DocumentDesign): string | null {
  if (hasOwnSchemes(design)) return null;
  const holds = (id: string | null) => {
    const theme = id ? THEMES_BY_ID.get(id) : undefined;
    return Boolean(
      theme &&
        theme.colors === design.colors &&
        theme.fonts === design.fonts &&
        theme.effects === design.effects,
    );
  };
  if (holds(design.theme)) return design.theme;
  return OFFICE_THEMES.find((theme) => holds(theme.id))?.id ?? null;
}

/* ------------------------------------------------------------------ theme */

function colorScheme(id: string): OfficeColorScheme {
  return (
    OFFICE_COLOR_SCHEMES.find((scheme) => scheme.id === id) ??
    OFFICE_COLOR_SCHEMES.find((scheme) => scheme.id === "office")!
  );
}

/** The theme's twelve colours (the document's own, else Office's scheme). */
export function documentDesignColors(design: DocumentDesign): OfficeThemeColors {
  return design.nativeColors?.colors ?? colorScheme(design.colors).colors;
}

/** Word's Office theme sets its body text in Aptos (Excel's in Aptos Narrow). */
const WORD_OFFICE_MINOR_LATIN = "Aptos";

function fontScheme(id: string): OfficeFontScheme {
  return (
    OFFICE_FONT_SCHEMES.find((scheme) => scheme.id === id) ??
    OFFICE_FONT_SCHEMES.find((scheme) => scheme.id === "office")!
  );
}

/** The theme's heading (major) and body (minor) fonts, as Word has them. */
export function documentDesignFontScheme(id: string): {
  major: OfficeThemeFontCollection;
  minor: OfficeThemeFontCollection;
  scheme: OfficeFontScheme;
} {
  const scheme = fontScheme(id);
  const minor =
    scheme.id === "office" ? { ...scheme.minor, latin: WORD_OFFICE_MINOR_LATIN } : scheme.minor;
  return { major: scheme.major, minor, scheme };
}

/** The design's heading (major) and body (minor) fonts: the document's own, else Office's scheme. */
export function documentDesignFonts(design: DocumentDesign): {
  major: OfficeThemeFontCollection;
  minor: OfficeThemeFontCollection;
} {
  if (design.nativeFonts) return { major: design.nativeFonts.major, minor: design.nativeFonts.minor };
  const { major, minor } = documentDesignFontScheme(design.fonts);
  return { major, minor };
}

/**
 * The Latin face of the document theme's body (minor) font: the face Word sets
 * a SmartArt graphic's words in when they name none (`a:fontRef idx="minor"`).
 * The page's `documentDesign` says it when the document has one (an imported
 * theme, or the scheme the Design tab chose). An imported document the Design
 * tab never touched states none; its Normal face (`defaultFontFamily`) is the
 * theme's body font there (Word's templates set Normal in `minorHAnsi`). A new
 * document is Word's new-document Office theme.
 */
export function documentThemeBodyLatin(documentDesign: unknown, defaultFontFamily?: unknown): string {
  const design = normalizeDocumentDesign(documentDesign);
  if (design) return documentDesignFonts(design).minor.latin.trim() || WORD_OFFICE_MINOR_LATIN;
  const normal =
    typeof defaultFontFamily === "string"
      ? (defaultFontFamily.split(",")[0] ?? "").trim().replace(/^["']|["']$/g, "").trim()
      : "";
  return normal || documentDesignFonts(DEFAULT_DOCUMENT_DESIGN).minor.latin.trim() || WORD_OFFICE_MINOR_LATIN;
}

/** The Fonts gallery entry the editor draws for a scheme (Word's Office fonts). */
export function documentDesignFontSet(id: string): OfficeThemeFontSet {
  const { major, minor, scheme } = documentDesignFontScheme(id);
  return officeThemeFontSet(scheme.id, scheme.label, scheme.schemeName, scheme.gallery, major, minor);
}

/** The East Asian face a theme font collection gives Chinese text. */
function eastAsianFace(collection: OfficeThemeFontCollection): string | null {
  return collection.ea || collection.scripts.Hans || null;
}

/** Word's theme colour names (styles.xml `w:themeColor`) on the twelve slots. */
const WORD_THEME_COLOR_SLOTS: Readonly<Record<string, OfficeThemeColorSlot>> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  text1: "dk1",
  background1: "lt1",
  text2: "dk2",
  background2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

const BYTE = /^[0-9a-f]{2}$/i;

/**
 * A Word theme colour with its tint or shade (`w:themeTint` / `w:themeShade`,
 * ECMA-376 §17.3.2.6): the luminance moves toward white by the tint, or is
 * scaled by the shade, in HSL, as Word caches the result in `w:val` (to
 * within one unit per channel).
 */
export function wordThemeColorHex(
  colors: OfficeThemeColors,
  themeColor: string,
  tint?: string,
  shade?: string,
): string | null {
  const slot = WORD_THEME_COLOR_SLOTS[themeColor];
  if (!slot) return null;
  const base = colors[slot].toLowerCase();
  const tintByte = tint && BYTE.test(tint) ? parseInt(tint, 16) : null;
  const shadeByte = shade && BYTE.test(shade) ? parseInt(shade, 16) : null;
  if ((tintByte === null && shadeByte === null) || (tintByte ?? shadeByte) === 255) return base;
  const [hue, rawLuminance, saturation] = rgbToHls(
    parseInt(base.slice(1, 3), 16) / 255,
    parseInt(base.slice(3, 5), 16) / 255,
    parseInt(base.slice(5, 7), 16) / 255,
  );
  // Word's cached values (styles.xml under each theme, measured) fit an
  // unrounded luminance best; a channel may still differ from Word's by one.
  let luminance = rawLuminance;
  if (tintByte !== null) luminance = luminance + (1 - luminance) * (1 - tintByte / 255);
  else if (shadeByte !== null) luminance *= shadeByte / 255;
  const channels = hlsToRgb(hue, Math.min(Math.max(luminance, 0), 1), saturation);
  return `#${channels.map((channel) => Math.trunc(channel * 255).toString(16).padStart(2, "0")).join("")}`;
}

/** A style's colour under a theme: the theme reference wins over the cached hex. */
function styleColor(colors: OfficeThemeColors, color: WordStyleColor | undefined): string | null {
  if (!color) return null;
  if (color.themeColor) {
    const resolved = wordThemeColorHex(colors, color.themeColor, color.themeTint, color.themeShade);
    if (resolved) return resolved;
  }
  return color.val && /^[0-9a-f]{6}$/i.test(color.val) ? `#${color.val.toLowerCase()}` : null;
}

function styleFill(colors: OfficeThemeColors, shd: WordStyleColor | undefined): string | null {
  if (!shd) return null;
  if (shd.themeFill) {
    const resolved = wordThemeColorHex(colors, shd.themeFill, shd.themeFillTint, shd.themeFillShade);
    if (resolved) return resolved;
  }
  return shd.fill && /^[0-9a-f]{6}$/i.test(shd.fill) ? `#${shd.fill.toLowerCase()}` : null;
}

/** The ten colours table styles, charts and covers read (`documentThemeColorsForExport`). */
export function documentDesignWordThemeColors(design: DocumentDesign): {
  text1: string;
  background1: string;
  text2: string;
  background2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
} {
  const colors = documentDesignColors(design);
  return {
    text1: colors.dk1,
    background1: colors.lt1,
    text2: colors.dk2,
    background2: colors.lt2,
    accent1: colors.accent1,
    accent2: colors.accent2,
    accent3: colors.accent3,
    accent4: colors.accent4,
    accent5: colors.accent5,
    accent6: colors.accent6,
  };
}

/* ---------------------------------------------------------- style recipes */

function fixed(value: number, digits = 2): number {
  const scale = 10 ** digits;
  return Math.round(value * scale) / scale;
}

function twipsToPt(value: number | undefined): number | null {
  return typeof value === "number" && Number.isFinite(value) ? fixed(value / 20) : null;
}

/** Half-points as the px the editor draws, spelled as the document stores
 * a size (11pt -> "14.6667px"; lib/document-font-size.ts). */
export function wordHalfPointsToPx(halfPoints: number): string {
  return documentFontSizePx(halfPoints / 2);
}

/** `w:line` with `lineRule="auto"` (240ths of a line) as a CSS multiple. */
function lineMultiple(ppr: WordStyleParagraphProps): string | null {
  if (typeof ppr.line !== "number") return null;
  if (ppr.lineRule && ppr.lineRule !== "auto") return null;
  return String(fixed(ppr.line / 240, 4));
}

const ALIGN: Readonly<Record<string, "left" | "center" | "right" | "justify">> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

function border(colors: OfficeThemeColors, edge: WordStyleBorder | undefined): DocumentBorderLine | null {
  if (!edge || !edge.val || edge.val === "nil" || edge.val === "none") return null;
  const color = styleColor(colors, { val: edge.color, themeColor: edge.themeColor, themeTint: edge.themeTint, themeShade: edge.themeShade });
  return {
    style: edge.val,
    ...(typeof edge.sz === "number" ? { sizePt: edge.sz / 8 } : {}),
    ...(color ? { color } : {}),
    ...(typeof edge.space === "number" ? { spacePt: edge.space } : {}),
  };
}

export interface DocumentStyleParagraphRecipe {
  readonly spaceBeforePt: number | null;
  readonly spaceAfterPt: number | null;
  readonly lineHeight: string | null;
  readonly textAlign: "left" | "center" | "right" | "justify" | null;
  readonly indentLeftPt: number | null;
  readonly indentRightPt: number | null;
  readonly firstLineIndentPt: number | null;
  readonly keepNext: boolean | null;
  readonly keepLines: boolean | null;
  readonly contextualSpacing: boolean | null;
  readonly borders: DocumentParagraphBorders | null;
  readonly shading: string | null;
}

export interface DocumentStyleRunRecipe {
  /** The Latin face (one font name, as Word writes `w:ascii`). */
  readonly fontFamily: string | null;
  /** The East Asian face. */
  readonly fontCjkFamily: string | null;
  readonly fontSize: string | null;
  readonly textColor: string | null;
  readonly bold: boolean | null;
  readonly italic: boolean | null;
  readonly allCaps: boolean | null;
  readonly smallCaps: boolean | null;
  readonly charSpacingPt: number | null;
  readonly underline: boolean | null;
}

export interface DocumentStyleRecipe {
  readonly paragraph: DocumentStyleParagraphRecipe;
  readonly run: DocumentStyleRunRecipe;
}

function paragraphRecipe(colors: OfficeThemeColors, ppr: WordStyleParagraphProps): DocumentStyleParagraphRecipe {
  const ind = ppr.ind ?? {};
  const borders: DocumentParagraphBorders = {};
  for (const edge of ["top", "left", "bottom", "right", "between"] as const) {
    const line = border(colors, ppr.pBdr?.[edge]);
    if (line) borders[edge] = line;
  }
  return {
    spaceBeforePt: twipsToPt(ppr.before),
    spaceAfterPt: twipsToPt(ppr.after),
    lineHeight: lineMultiple(ppr),
    textAlign: ppr.jc ? (ALIGN[ppr.jc] ?? null) : null,
    indentLeftPt: twipsToPt(ind.left),
    indentRightPt: twipsToPt(ind.right),
    firstLineIndentPt:
      typeof ind.hanging === "number"
        ? -fixed(ind.hanging / 20)
        : twipsToPt(ind.firstLine),
    keepNext: ppr.keepNext ?? null,
    keepLines: ppr.keepLines ?? null,
    contextualSpacing: ppr.contextualSpacing ?? null,
    borders: Object.keys(borders).length ? borders : null,
    shading: styleFill(colors, ppr.shd),
  };
}

function runRecipe(design: DocumentDesign, rpr: WordStyleRunProps): DocumentStyleRunRecipe {
  const colors = documentDesignColors(design);
  const fonts = documentDesignFonts(design);
  const latin = rpr.font ?? (rpr.fontTheme ? fonts[rpr.fontTheme].latin : null);
  const eastSlot = rpr.eastAsiaTheme ?? rpr.fontTheme;
  const eastAsian = eastSlot ? eastAsianFace(fonts[eastSlot]) : null;
  const underline = rpr.u?.val ? rpr.u.val !== "none" : null;
  return {
    fontFamily: latin || null,
    fontCjkFamily: eastAsian ?? (latin || null),
    fontSize: typeof rpr.sz === "number" && rpr.sz > 0 ? wordHalfPointsToPx(rpr.sz) : null,
    textColor: rpr.color?.val === "auto" && !rpr.color.themeColor ? null : styleColor(colors, rpr.color),
    bold: rpr.b ?? null,
    italic: rpr.i ?? null,
    allCaps: rpr.caps ?? null,
    smallCaps: rpr.smallCaps ?? null,
    charSpacingPt: typeof rpr.spacing === "number" && rpr.spacing !== 0 ? fixed(rpr.spacing / 20) : null,
    underline,
  };
}

/** The design's style set, or null while Word's style sets are not here (lib/office-data). */
function setOf(design: DocumentDesign): WordStyleSet | null {
  return wordStyleSet(design.styleSet) ?? wordStyleSet(WORD_DEFAULT_STYLE_SET_ID) ?? null;
}

/**
 * One Quick Style's own formatting under the design: what its paragraphs
 * wear on top of the document's base (its `basedOn` chain up to Normal),
 * with the theme's fonts and colours resolved. Null while Word's style sets
 * are not here.
 */
export function documentDesignStyleRecipe(
  design: DocumentDesign,
  id: WordQuickStyleId,
): DocumentStyleRecipe | null {
  const set = setOf(design);
  if (!set) return null;
  const { pPr, rPr } = wordStyleChain(set, id);
  return {
    paragraph: paragraphRecipe(documentDesignColors(design), pPr),
    run: runRecipe(design, rPr),
  };
}

/** The document defaults (Normal) the design gives: the editor's `default*` document attributes. */
export interface DocumentDesignBase {
  readonly defaultFontFamily: string;
  readonly defaultFontSize: string;
  readonly defaultTextColor: string | null;
  readonly defaultLineHeight: string;
  readonly defaultTextAlign: "left" | "center" | "right" | "justify" | null;
  readonly defaultParagraphSpaceBefore: number;
  readonly defaultParagraphSpaceAfter: number;
  readonly defaultIndentLeftPt: number | null;
  readonly defaultIndentRightPt: number | null;
  readonly defaultFirstLineIndentPt: number | null;
}

function quotedFace(face: string): string {
  return `"${face.replace(/["'\\;{}()]/g, "").slice(0, 200)}"`;
}

/**
 * The CSS stack the document's body text takes from a theme slot: the
 * theme's Latin face, then its East Asian face, then the platform's.
 */
export function documentDesignFontStack(design: DocumentDesign, slot: "major" | "minor"): string {
  const collection = documentDesignFonts(design)[slot];
  const faces: string[] = [];
  for (const face of [collection.latin, eastAsianFace(collection) ?? ""]) {
    const trimmed = face.trim();
    if (trimmed && !faces.some((existing) => existing.toLowerCase() === quotedFace(trimmed).toLowerCase())) {
      faces.push(quotedFace(trimmed));
    }
  }
  return [...faces, '"PingFang SC"', '"Microsoft YaHei"', "sans-serif"].join(", ");
}

/**
 * The document's base under the design: Normal over `w:docDefaults`, with
 * a Paragraph Spacing choice in place of the style set's own spacing (Word
 * writes the choice into `w:docDefaults`). Word's default size is 10pt and
 * its default spacing none, so a set that states neither has those.
 */
export function documentDesignBase(design: DocumentDesign): DocumentDesignBase | null {
  const set = setOf(design);
  if (!set) return null;
  const { pPr, rPr } = wordStyleSetBase(set);
  const colors = documentDesignColors(design);
  const run = runRecipe(design, rPr);
  const paragraph = paragraphRecipe(colors, pPr);
  const preset = design.paragraphSpacing
    ? documentParagraphSpacingPreset(design.paragraphSpacing)
    : null;
  const fontSlot = rPr.fontTheme ?? "minor";
  return {
    defaultFontFamily: rPr.font
      ? [quotedFace(rPr.font), '"PingFang SC"', '"Microsoft YaHei"', "sans-serif"].join(", ")
      : documentDesignFontStack(design, fontSlot),
    defaultFontSize: run.fontSize ?? wordHalfPointsToPx(20),
    defaultTextColor: run.textColor,
    defaultLineHeight:
      preset && preset.lineHeight !== null ? preset.lineHeight : (paragraph.lineHeight ?? "1"),
    defaultTextAlign: paragraph.textAlign,
    defaultParagraphSpaceBefore:
      preset && preset.beforePt !== null ? preset.beforePt : (paragraph.spaceBeforePt ?? 0),
    defaultParagraphSpaceAfter:
      preset && preset.afterPt !== null ? preset.afterPt : (paragraph.spaceAfterPt ?? 0),
    defaultIndentLeftPt: paragraph.indentLeftPt && paragraph.indentLeftPt > 0 ? paragraph.indentLeftPt : null,
    defaultIndentRightPt: paragraph.indentRightPt && paragraph.indentRightPt > 0 ? paragraph.indentRightPt : null,
    defaultFirstLineIndentPt: paragraph.firstLineIndentPt || null,
  };
}

/** The Quick Styles a design restyles, with the editor's gallery ids. */
export const DOCUMENT_DESIGN_PARAGRAPH_STYLES = [
  "noSpacing",
  "heading1",
  "heading2",
  "heading3",
  "heading4",
  "heading5",
  "heading6",
  "title",
  "subtitle",
  "quote",
  "intenseQuote",
  "listParagraph",
] as const satisfies readonly WordQuickStyleId[];

export const DOCUMENT_DESIGN_CHARACTER_STYLES = [
  "subtleEmphasis",
  "emphasis",
  "intenseEmphasis",
  "strong",
  "subtleReference",
  "intenseReference",
  "bookTitle",
] as const satisfies readonly WordQuickStyleId[];

/**
 * The rules that draw the Quick character styles under a design, scoped to
 * one page surface: each style as the style set defines it, with the theme's
 * colours. The editor paints a character style inline with Word's default
 * definition (`characterStyleCss`); these rules win over it (`!important`)
 * and state every property that definition sets, so a set whose Strong is
 * not bold leaves the text's own weight. A style colour still yields to a
 * direct colour on the same text (`--docs-direct-text-color`). Empty while
 * Word's style sets are not here (Word's default definitions draw).
 */
export function documentDesignCharacterStyleCss(design: DocumentDesign, scope: string): string {
  const rules: string[] = [];
  for (const id of DOCUMENT_DESIGN_CHARACTER_STYLES) {
    const recipe = documentDesignStyleRecipe(design, id);
    if (!recipe) return "";
    const { run } = recipe;
    const declarations = [
      `font-weight: ${run.bold ? "700" : "inherit"}`,
      `font-style: ${run.italic ? "italic" : "inherit"}`,
      `font-variant: ${run.smallCaps ? "small-caps" : "inherit"}`,
      `text-transform: ${run.allCaps ? "uppercase" : "inherit"}`,
      `color: ${run.textColor ? `var(--docs-direct-text-color, ${run.textColor})` : "inherit"}`,
      `letter-spacing: ${run.charSpacingPt ? `${run.charSpacingPt}pt` : "inherit"}`,
      ...(run.underline ? ["text-decoration-line: underline"] : []),
    ];
    rules.push(
      `${scope} .note-editor-content [data-character-style="${id}"] { ${declarations
        .map((declaration) => `${declaration} !important`)
        .join("; ")}; }`,
    );
  }
  return rules.join("\n");
}

/** Word's Fonts gallery: Office's list, with Word's own Office entry. */
export const DOCUMENT_DESIGN_FONT_GALLERY: readonly OfficeThemeFontSet[] = Object.freeze(
  OFFICE_FONT_SCHEMES.filter((scheme) => scheme.gallery).map((scheme) =>
    documentDesignFontSet(scheme.id),
  ),
);

const GRID_SLOT_NAMES: Readonly<Record<string, string>> = {
  bg1: "background1",
  tx1: "text1",
  bg2: "background2",
  tx2: "text2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
};

/**
 * Word's reference for a Theme Colors swatch (`text1`, `accent1:tint:99`,
 * `background1:shade:BF`): Lighter p% is `w:themeTint` (1 - p) of 255,
 * Darker p% is `w:themeShade` (1 - p) of 255, as Word's picker writes them
 * (Lighter 40% = 99, Darker 25% = BF). Null for a transform no swatch makes.
 */
export function wordThemeColorRefForSwatch(ref: {
  readonly slot: string;
  readonly lumMod?: number;
  readonly lumOff?: number;
}): string | null {
  const name = GRID_SLOT_NAMES[ref.slot];
  if (!name) return null;
  const lumMod = ref.lumMod ?? 100000;
  const lumOff = ref.lumOff ?? 0;
  if (lumOff === 0 && lumMod === 100000) return name;
  const byte = (fraction: number) =>
    Math.round(Math.min(Math.max(fraction, 0), 1) * 255)
      .toString(16)
      .toUpperCase()
      .padStart(2, "0");
  if (lumOff === 0) return `${name}:shade:${byte(lumMod / 100000)}`;
  if (lumMod + lumOff === 100000) return `${name}:tint:${byte(lumMod / 100000)}`;
  return null;
}
