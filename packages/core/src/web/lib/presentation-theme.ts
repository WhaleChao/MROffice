/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Deck-level slide master / theme (PowerPoint / WPS parity).
 *
 * A presentation may carry ONE deck-level theme that supplies the DEFAULT
 * background, font family, and title / body text colours for every slide and
 * block. A slide's own `backgroundColor`, or a text block's own `style.color` /
 * `style.fontFamily`, always OVERRIDES the theme; the theme in turn overrides
 * the built-in render defaults (no explicit colour). The precedence is therefore
 * strictly:
 *
 *   block / slide override  >  deck theme  >  built-in default
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the read-only present-mode overlay so the allowed font set, colour
 * validation, preset themes, and the resolution rules live in exactly one place.
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and imported from the zod schema.
 *
 * Everything here is ADDITIVE and OPTIONAL: a deck with no `theme` resolves to
 * the historical built-in defaults, so existing decks render byte-for-byte the
 * same as before.
 */

import {
  normalizeDeckOfficeMaster,
  type DeckOfficeMaster,
} from "./presentation-office-master-model";
import {
  officeThemeFontGalleryEntry,
  officeThemeFontStack,
} from "./office-theme-fonts";
import {
  PRESENTATION_OFFICE_COLOR_SLOTS,
  PRESENTATION_OFFICE_THEME_DEFAULT_ID,
  PRESENTATION_OFFICE_THEME_SHAPE_LINE_EMU,
  PRESENTATION_OFFICE_VARIANT_IDS,
  isOfficeThemeVariantId,
  presentationOfficeTheme,
  presentationOfficeThemeHansFace,
  presentationOfficeThemeVariant,
  type PresentationOfficeVariantId,
} from "./presentation-office-themes";
import { normalizeBgGradient } from "./presentation-slide-bg";
import { officeColorGalleryScheme, officeEffectGalleryScheme } from "./office-theme-catalog";
import { officeTransformedHex } from "./office-color-transform";
import type { OfficeDefaultShapeStyle } from "./office-default-shape-style";
export * from "./presentation-block-title";
import { blockIsTitle } from "./presentation-block-title";

/** Six-digit hex colour, e.g. `#1a2b3c`. Matches the slide background regex. */
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

/**
 * Bounds for a theme-supplied default font size (px). A slide master's title /
 * body text style may pin a default size; the value is clamped to this sane
 * range so a stale / hand-edited theme can never request a degenerate (0 / NaN /
 * absurd) size. Mirrors the structural intent of the per-block `style.fontSize`.
 */
export const MIN_THEME_FONT_SIZE = 6;
export const MAX_THEME_FONT_SIZE = 400;

/**
 * Hard cap on the number of accent (quick-pick) colours a theme may carry. A
 * theme exposes its palette as additive swatches in the shape-fill / text /
 * run colour pickers; the cap keeps a stale / hand-edited theme from injecting
 * an unbounded swatch row. PowerPoint themes ship six accents, so six is a
 * comfortable, faithful ceiling.
 */
export const MAX_THEME_ACCENT_COLORS = 6;

/**
 * The whitelist of allowed theme font families. Constrained so a stale or
 * hand-edited deck can never inject an arbitrary `font-family` string (which
 * could smuggle CSS / a remote `@import`). Each entry is a complete, safe
 * font-stack suitable for use directly as a CSS `font-family` value.
 */
export const THEME_FONT_FAMILIES = [
  "system",
  "serif",
  "mono",
  "rounded",
] as const;

export type ThemeFontFamily = (typeof THEME_FONT_FAMILIES)[number];

/** The font family applied when a theme omits (or mis-specifies) one. */
export const DEFAULT_THEME_FONT_FAMILY: ThemeFontFamily = "system";

const THEME_FONT_FAMILY_SET: ReadonlySet<string> = new Set(THEME_FONT_FAMILIES);

/**
 * Map a whitelisted font family key onto a concrete, safe CSS `font-family`
 * stack. An unknown / missing key collapses to the system stack so the resolved
 * value is always a known-safe string.
 */
export const THEME_FONT_FAMILY_CSS: Record<ThemeFontFamily, string> = {
  system:
    "system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
  serif: "Georgia, Cambria, 'Times New Roman', Times, serif",
  mono: "'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace",
  rounded:
    "'Nunito', 'Quicksand', 'Segoe UI', system-ui, -apple-system, sans-serif",
};

/** Versioned receipt that keeps a gallery selection distinct from its colours. */
export const PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION =
  "presentation-theme-identity-v1" as const;

/** The computed variants of an MRAI catalog / native / custom theme. */
export const PRESENTATION_THEME_DESIGN_VARIANT_IDS = [
  "base",
  "soft",
  "deep",
  "contrast",
] as const;

export type PresentationThemeDesignVariantId =
  (typeof PRESENTATION_THEME_DESIGN_VARIANT_IDS)[number];

/**
 * Every variant id an identity may carry: the four computed variants of the
 * MRAI kinds, plus an Office theme's own variants `v1`..`v8` (`v1` = the
 * theme's base variant; Integral has eight, every other theme four).
 */
export const PRESENTATION_THEME_VARIANT_IDS = [
  ...PRESENTATION_THEME_DESIGN_VARIANT_IDS,
  ...PRESENTATION_OFFICE_VARIANT_IDS,
] as const;

export type PresentationThemeVariantId =
  (typeof PRESENTATION_THEME_VARIANT_IDS)[number];

const THEME_DESIGN_VARIANT_ID_SET: ReadonlySet<string> = new Set(
  PRESENTATION_THEME_DESIGN_VARIANT_IDS,
);

export function isThemeDesignVariantId(
  value: unknown,
): value is PresentationThemeDesignVariantId {
  return typeof value === "string" && THEME_DESIGN_VARIANT_ID_SET.has(value);
}

export const PRESENTATION_THEME_COLOR_SLOTS = [
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

export type PresentationThemeColorSlot =
  (typeof PRESENTATION_THEME_COLOR_SLOTS)[number];

/** The supported theme fields a variant may intentionally replace. */
export interface DeckThemeTokenSnapshot {
  background?: string;
  backgroundGradient?: unknown;
  backgroundPattern?: unknown;
  accentColors?: string[];
  fontFamily?: ThemeFontFamily;
  titleColor?: string;
  bodyTextColor?: string;
  titleFontSize?: number;
  bodyFontSize?: number;
  titleBold?: boolean;
  bodyBold?: boolean;
  titleItalic?: boolean;
  bodyItalic?: boolean;
}

/** Immutable authority stamped by PPTX import for original-package patchback. */
export interface NativeDeckThemeSource {
  part: string;
  fingerprint: string;
  /** Number of distinct themes referenced by slide masters in the source. */
  themeCount?: number;
  name?: string;
  colorSlots?: Partial<Record<PresentationThemeColorSlot, string>>;
  majorLatinFont?: string;
  minorLatinFont?: string;
  baseTheme?: DeckThemeTokenSnapshot;
  /** Authenticates every relationship-bound master carrying this theme. */
  masterDesignFingerprint?: string;
  /**
   * A deck whose masters use several themes: each of them, so a Design pick
   * reaches every master (row 231).
   */
  themeParts?: { part: string; fingerprint: string; masterDesignFingerprint: string }[];
  /**
   * The Effects-gallery entry the theme's own `a:fmtScheme` is (reading row
   * 851): what Design > Variants > Effects checks until a pick
   * (`effectSchemeId`) replaces it. A receipt of the file, never an edit.
   */
  effectSchemeMatch?: string;
}

/**
 * Stable selection identity. `source` survives a later catalog/variant choice so
 * patchback still knows which original theme part it is allowed to address.
 */
export interface DeckThemeIdentity {
  schemaVersion: typeof PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION;
  /**
   * `catalog` = an MRAI preset, `office` = one of PowerPoint's Office themes
   * (`themeId` from lib/presentation-office-themes, `variantId` `v1`..`v8`),
   * `native` = an imported deck's own theme, `custom` = edited by hand.
   */
  selectionKind: "catalog" | "office" | "native" | "custom";
  themeId: string;
  variantId: PresentationThemeVariantId;
  source?: NativeDeckThemeSource;
  /**
   * `package`: the import recognised the deck's OWN theme as this Office
   * theme (its family, or its name and twelve colours) -- nobody chose it
   * here. The file's master, decoration included, is still the deck's, and
   * its graphics are the slides' inherited blocks; only a theme picked in
   * Design (which writes its master over the file's) draws the gallery's
   * artwork instead. Absent on a pick (`applyOfficeThemeSelection`).
   */
  origin?: "package";
}

/**
 * The deck-level theme. Every field is OPTIONAL so a partial theme (e.g. only a
 * background) is valid and the missing fields fall through to the built-in
 * defaults. Colours are normalized six-digit lower-case hex; `fontFamily` is a
 * whitelisted key.
 */
export interface DeckTheme {
  /** Stable gallery/native identity; never inferred from colour coincidence. */
  identity?: DeckThemeIdentity;
  /** Default slide background colour (overridden by `slide.backgroundColor`). */
  background?: string;
  /**
   * Optional deck-level RICH background fill (gradient / picture / pattern), the
   * DEFAULT for every slide that supplies no rich fill of its own. Resolved by
   * `lib/presentation-slide-bg`; typed loosely here (the resolver re-validates)
   * and carried through `normalizeTheme` verbatim. Additive — a theme without
   * these is unchanged.
   */
  backgroundGradient?: unknown;
  backgroundImage?: unknown;
  backgroundPattern?: unknown;
  /**
   * The theme's accent palette: a short list (≤ {@link MAX_THEME_ACCENT_COLORS})
   * of normalized six-digit lower-case hex colours surfaced as ADDITIVE quick-pick
   * swatches in the shape-fill / text / run colour pickers. Purely a convenience
   * palette — it never changes how anything renders on its own, so it has no part
   * in the colour-resolution precedence and a deck without accents is unaffected.
   */
  accentColors?: string[];
  /** Default font family for every text block (overridden by `style.fontFamily`). */
  fontFamily?: ThemeFontFamily;
  /**
   * Design > Variants > Fonts: the Fonts-gallery entry (Office's heading/body
   * pair, lib/office-theme-fonts) the user picked. Written ONLY by that
   * command -- an import never carries it -- so its presence is the theme
   * font edit (SHAPE-IMPORT-AUDIT.md §75.1). Title text renders in its
   * heading face and body text in its body face, ahead of `fontFamily`.
   */
  fontSchemeId?: string;
  /**
   * Design > Variants > Colors: the Colors-gallery entry (Office's colour
   * scheme, lib/office-theme-catalog) the user picked (PPT-A073). Written
   * ONLY by that command -- an import never carries it -- so its presence
   * is the colour-scheme edit; `colorSlots` and the colours materialized
   * from them hold the scheme's colours. A theme or variant pick drops it.
   */
  colorSchemeId?: string;
  /**
   * Design > Variants > Effects: the Effects-gallery entry (an Office effect
   * scheme, lib/office-theme-catalog) the user picked (PPT-A073). Written
   * only by that command; the download makes it the theme's whole
   * `a:fmtScheme`. A theme or variant pick drops it.
   */
  effectSchemeId?: string;
  /**
   * The theme's twelve colour slots (dk1, lt1, dk2, lt2, accent1..6, hlink,
   * folHlink) as lower-case `#rrggbb`. Written when an Office theme or variant
   * is applied; `background` / `titleColor` / `bodyTextColor` /
   * `accentColors` are materialized from it for every existing reader.
   */
  colorSlots?: Record<PresentationThemeColorSlot, string>;
  /**
   * The first master's colour map (`p:clrMap`) of the four names it maps
   * (Background 1 / Text 1 / Background 2 / Text 2 to one of the twelve
   * slots), as an imported deck's master states it -- read only where the
   * map is dark (`bg1` not Light 1: Berlin, an imported dark deck), reading
   * row 826. Text 1 (the slide's ink: a new text box's and a placeholder's
   * words) and the Theme Colors grid's Text / Background cells resolve
   * through it (`presentationThemeColorMap`); absent = the Office theme
   * variant's own map, else the light one.
   */
  colorMap?: DeckThemeColorMap;
  /**
   * The theme's font scheme by face name: headings (`+mj-lt`) and body
   * (`+mn-lt`), plus the Simplified Chinese faces of each (`script="Hans"`).
   * Title / body text render in these faces ahead of `fontSchemeId` and
   * `fontFamily`.
   */
  fontScheme?: DeckThemeFontScheme;
  /** Default colour for title / section-title text blocks + the slide title. */
  titleColor?: string;
  /** Default colour for body / generic text blocks. */
  bodyTextColor?: string;
  /**
   * Default font size (px) for title / section-title text blocks + the slide
   * title. Part of the slide-master TITLE text style. Overridden by a block's own
   * `style.fontSize`. Absent → the renderer's built-in title size (unchanged).
   */
  titleFontSize?: number;
  /**
   * Default font size (px) for body / generic text blocks. Part of the
   * slide-master BODY text style. Overridden by a block's own `style.fontSize`.
   * Absent → the renderer's built-in body size (unchanged).
   */
  bodyFontSize?: number;
  /**
   * Whether title / section-title text defaults to bold. Part of the slide-master
   * TITLE text style. Overridden by a block's own `style.bold`. Absent → no theme
   * weight contribution (the renderer's built-in weight, unchanged).
   */
  titleBold?: boolean;
  /**
   * Whether body / generic text defaults to bold. Part of the slide-master BODY
   * text style. Overridden by a block's own `style.bold`. Absent → no theme weight
   * contribution (the renderer's built-in weight, unchanged).
   */
  bodyBold?: boolean;
  /**
   * Whether title / body text defaults to italic: the slide master's title /
   * body text style Italic (`a:defRPr@i`, PowerPoint's Slide Master view >
   * Font > Italic). Overridden by a block's own `style.italic`; absent = no
   * theme slant, as for bold.
   */
  titleItalic?: boolean;
  bodyItalic?: boolean;
  /**
   * Slide Master view's edits of an Office theme deck's one master
   * (PPT-B030): its text styles and whether its background is the theme's
   * background fields. Only Slide Master view writes it; a theme or variant
   * pick drops it. See lib/presentation-office-master-model.
   */
  officeMaster?: DeckOfficeMaster;
}

/** The four names a master's colour map resolves, each to a theme slot. */
export const DECK_THEME_MAPPED_COLORS = ["bg1", "tx1", "bg2", "tx2"] as const;

export type DeckThemeColorMap = Readonly<
  Record<(typeof DECK_THEME_MAPPED_COLORS)[number], PresentationThemeColorSlot>
>;

/** A theme's heading / body faces (Office font scheme, by name). */
export interface DeckThemeFontScheme {
  majorLatin: string;
  minorLatin: string;
  majorHans?: string;
  minorHans?: string;
  /**
   * The scheme's name (`a:fontScheme@name`): what Create New Theme Fonts
   * saves it as ("Custom 1"), and what the Fonts gallery's Custom section
   * lists it by. Absent on a scheme the file names otherwise or not at all.
   */
  name?: string;
}

/** The longest theme font scheme name Create New Theme Fonts keeps. */
export const MAX_THEME_FONT_SCHEME_NAME = 80;

/**
 * The font scheme Create New Theme Fonts starts from: the theme's own named
 * scheme, else the Fonts-gallery entry it wears (`scheme` overrides it), else
 * an imported deck's own faces as its theme part states them (the import's
 * `identity.source` Latin faces; its Asian faces stay the file's until one is
 * given), else Office's. PowerPoint's dialog opens on the theme's current
 * faces, never on another scheme's.
 */
export function themeFontSchemeStart(
  current: Pick<DeckTheme, "fontScheme" | "fontSchemeId" | "identity"> | null | undefined,
  scheme?: string,
): DeckThemeFontScheme {
  const own = scheme ? undefined : normalizeThemeFontScheme(current?.fontScheme);
  if (own) return own;
  const pickedId = scheme ?? current?.fontSchemeId;
  const identity = current?.identity;
  const native =
    !pickedId && identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source
      : undefined;
  if (native?.majorLatinFont && native.minorLatinFont) {
    return { majorLatin: native.majorLatinFont, minorLatin: native.minorLatinFont };
  }
  const entry = officeThemeFontGalleryEntry(pickedId ?? "office");
  return {
    majorLatin: entry?.headingLatin ?? "Calibri Light",
    minorLatin: entry?.bodyLatin ?? "Calibri",
    ...(entry?.headingEastAsia ? { majorHans: entry.headingEastAsia } : {}),
    ...(entry?.bodyEastAsia ? { minorHans: entry.bodyEastAsia } : {}),
  };
}

/**
 * Design > Variants > Fonts > Customize Fonts (PowerPoint's Create New Theme
 * Fonts): the theme's heading and body faces, Latin and East Asian, each one
 * not given kept from the scheme the edit starts from (a picked gallery
 * entry, else the theme's own scheme, else Office's), named as given. The
 * one computation behind the dialog and the `theme_fonts` command. Undefined
 * when a face is not one a theme may name.
 */
export function customizedThemeFontScheme(
  current: Pick<DeckTheme, "fontScheme" | "fontSchemeId" | "identity"> | null | undefined,
  args: {
    readonly scheme?: string;
    readonly heading?: string;
    readonly body?: string;
    readonly headingEastAsian?: string;
    readonly bodyEastAsian?: string;
    readonly name?: string;
  },
): DeckThemeFontScheme | undefined {
  const start = themeFontSchemeStart(current, args.scheme);
  const name = typeof args.name === "string" ? args.name.trim() : "";
  return normalizeThemeFontScheme({
    ...start,
    ...(args.heading ? { majorLatin: args.heading } : {}),
    ...(args.body ? { minorLatin: args.body } : {}),
    ...(args.headingEastAsian ? { majorHans: args.headingEastAsian } : {}),
    ...(args.bodyEastAsian ? { minorHans: args.bodyEastAsian } : {}),
    ...(name ? { name } : {}),
  });
}

/** A named preset surfaced in the theme editor. */
export interface PresetTheme {
  /** Stable id (used as the React key + the select value). */
  id: string;
  /** i18n key suffix under `presentation.theme.preset.*`. */
  labelKey: string;
  /** The theme this preset applies (already normalized hex / whitelisted font). */
  theme: DeckTheme;
}

export type PresentationThemeDesignFamily =
  "editorial" | "geometric" | "spotlight" | "dark" | "organic" | "minimal";

const THEME_DESIGN_FAMILY_BY_ID: Readonly<
  Record<string, PresentationThemeDesignFamily>
> = {
  classic: "minimal",
  bjorn: "minimal",
  minimal: "minimal",
  ice: "minimal",
  slate: "minimal",
  corporate: "geometric",
  sky: "geometric",
  ocean: "geometric",
  indigo: "geometric",
  vibrant: "spotlight",
  rose: "spotlight",
  lavender: "spotlight",
  berry: "spotlight",
  plum: "spotlight",
  midnight: "dark",
  dark: "dark",
  graphite: "dark",
  forest: "organic",
  meadow: "organic",
  teal: "organic",
  mint: "organic",
  citrus: "organic",
  sunrise: "editorial",
  sand: "editorial",
  clay: "editorial",
  amber: "editorial",
  coral: "editorial",
  parchment: "editorial",
  horizon: "editorial",
  column: "minimal",
  terrace: "organic",
  facet: "geometric",
  prism: "spotlight",
  studio: "minimal",
  skyline: "geometric",
  dune: "editorial",
  tide: "organic",
  ember: "dark",
  aperture: "dark",
  aurora: "spotlight",
};

export function presentationThemeDesignFamily(
  themeId: string,
): PresentationThemeDesignFamily | undefined {
  return THEME_DESIGN_FAMILY_BY_ID[themeId];
}

/** Native DrawingML color bands, not thumbnail-only artwork. The large paper
 * area stays clear for content; equal positions produce crisp band boundaries. */
function bandTheme(
  id: string,
  paper: string,
  ink: string,
  accent: string,
  companion: string,
  angle: number,
  edge: number,
  secondaryEdge: number,
): PresetTheme {
  const stops = [
    { color: paper, position: 0 },
    { color: paper, position: edge },
    { color: accent, position: edge },
    { color: accent, position: secondaryEdge },
    { color: companion, position: secondaryEdge },
    { color: companion, position: 1 },
  ];
  return {
    id,
    labelKey: id,
    theme: {
      background: paper,
      titleColor: ink,
      bodyTextColor: ink,
      fontFamily: "system",
      accentColors: [
        accent,
        companion,
        "#38bdf8",
        "#a78bfa",
        "#fb7185",
        "#eab308",
      ],
      backgroundGradient: { from: paper, to: companion, angle, stops },
    },
  };
}

/**
 * The first generation of the preset palettes, frozen. Only a deck saved
 * before themes carried an identity is matched against them (exact values),
 * so it is still recognised as the preset it was; nothing draws from them.
 */
const LEGACY_PRESET_THEMES: readonly PresetTheme[] = [
  { id: "none", labelKey: "none", theme: {} },
  {
    id: "classic",
    labelKey: "classic",
    theme: {
      background: "#ffffff",
      fontFamily: "system",
      titleColor: "#1e293b",
      bodyTextColor: "#334155",
      accentColors: ["#2563eb", "#0ea5e9", "#10b981", "#f59e0b", "#ef4444"],
    },
  },
  {
    id: "midnight",
    labelKey: "midnight",
    theme: {
      background: "#0f172a",
      fontFamily: "system",
      titleColor: "#f8fafc",
      bodyTextColor: "#cbd5e1",
      accentColors: ["#38bdf8", "#818cf8", "#a78bfa", "#f472b6", "#34d399"],
    },
  },
  {
    id: "bjorn",
    labelKey: "bjorn",
    theme: {
      background: "#ffffff",
      fontFamily: "system",
      titleColor: "#000000",
      bodyTextColor: "#000000",
      accentColors: ["#28ae70", "#ac9f0d", "#ef3475", "#8961f3", "#10add3", "#e78811"],
    },
  },
  {
    id: "sunrise",
    labelKey: "sunrise",
    theme: {
      background: "#fff7ed",
      fontFamily: "rounded",
      titleColor: "#9a3412",
      bodyTextColor: "#7c2d12",
      accentColors: ["#ea580c", "#f59e0b", "#dc2626", "#d97706", "#b45309"],
    },
  },
  {
    id: "forest",
    labelKey: "forest",
    theme: {
      background: "#f0fdf4",
      fontFamily: "serif",
      titleColor: "#14532d",
      bodyTextColor: "#166534",
      accentColors: ["#16a34a", "#65a30d", "#0d9488", "#ca8a04", "#15803d"],
    },
  },
  {
    id: "corporate",
    labelKey: "corporate",
    theme: {
      background: "#f8fafc",
      fontFamily: "system",
      titleColor: "#1e3a8a",
      bodyTextColor: "#475569",
      titleFontSize: 40,
      bodyFontSize: 20,
      titleBold: true,
      accentColors: ["#1e40af", "#0369a1", "#0f766e", "#475569", "#b91c1c"],
    },
  },
  {
    id: "vibrant",
    labelKey: "vibrant",
    theme: {
      background: "#fdf4ff",
      fontFamily: "rounded",
      titleColor: "#a21caf",
      bodyTextColor: "#7e22ce",
      titleFontSize: 44,
      bodyFontSize: 22,
      titleBold: true,
      accentColors: ["#c026d3", "#7c3aed", "#db2777", "#0891b2", "#ea580c"],
    },
  },
  {
    id: "minimal",
    labelKey: "minimal",
    theme: {
      background: "#ffffff",
      fontFamily: "mono",
      titleColor: "#171717",
      bodyTextColor: "#525252",
      titleFontSize: 36,
      bodyFontSize: 18,
      accentColors: ["#171717", "#525252", "#737373", "#a3a3a3"],
    },
  },
  {
    id: "dark",
    labelKey: "dark",
    theme: {
      background: "#111827",
      fontFamily: "system",
      titleColor: "#f9fafb",
      bodyTextColor: "#9ca3af",
      titleFontSize: 40,
      bodyFontSize: 20,
      titleBold: true,
      accentColors: ["#60a5fa", "#a78bfa", "#f472b6", "#fbbf24", "#4ade80"],
    },
  },
  {
    id: "ocean",
    labelKey: "ocean",
    theme: {
      background: "#eff6ff",
      fontFamily: "system",
      titleColor: "#0c4a6e",
      bodyTextColor: "#075985",
      accentColors: ["#0284c7", "#0369a1", "#0891b2", "#0d9488", "#22c55e"],
    },
  },
  {
    id: "meadow",
    labelKey: "meadow",
    theme: {
      background: "#f7fee7",
      fontFamily: "serif",
      titleColor: "#365314",
      bodyTextColor: "#4d7c0f",
      accentColors: ["#65a30d", "#16a34a", "#14b8a6", "#eab308", "#84cc16"],
    },
  },
  {
    id: "sand",
    labelKey: "sand",
    theme: {
      background: "#fffbeb",
      fontFamily: "serif",
      titleColor: "#78350f",
      bodyTextColor: "#92400e",
      accentColors: ["#d97706", "#f59e0b", "#b45309", "#a16207", "#78716c"],
    },
  },
  {
    id: "clay",
    labelKey: "clay",
    theme: {
      background: "#fff7ed",
      fontFamily: "system",
      titleColor: "#7c2d12",
      bodyTextColor: "#9a3412",
      accentColors: ["#c2410c", "#ea580c", "#dc2626", "#a16207", "#57534e"],
    },
  },
  {
    id: "lavender",
    labelKey: "lavender",
    theme: {
      background: "#faf5ff",
      fontFamily: "rounded",
      titleColor: "#581c87",
      bodyTextColor: "#6b21a8",
      accentColors: ["#9333ea", "#c026d3", "#7c3aed", "#db2777", "#6366f1"],
    },
  },
  {
    id: "berry",
    labelKey: "berry",
    theme: {
      background: "#fdf2f8",
      fontFamily: "rounded",
      titleColor: "#831843",
      bodyTextColor: "#9d174d",
      accentColors: ["#db2777", "#e11d48", "#be123c", "#a21caf", "#7c3aed"],
    },
  },
  {
    id: "slate",
    labelKey: "slate",
    theme: {
      background: "#f8fafc",
      fontFamily: "mono",
      titleColor: "#0f172a",
      bodyTextColor: "#334155",
      accentColors: ["#475569", "#64748b", "#0f766e", "#1d4ed8", "#7c3aed"],
    },
  },
  {
    id: "sky",
    labelKey: "sky",
    theme: {
      background: "#f0f9ff",
      fontFamily: "system",
      titleColor: "#0c4a6e",
      bodyTextColor: "#0369a1",
      accentColors: ["#0ea5e9", "#38bdf8", "#0284c7", "#6366f1", "#14b8a6"],
    },
  },
  {
    id: "teal",
    labelKey: "teal",
    theme: {
      background: "#f0fdfa",
      fontFamily: "system",
      titleColor: "#134e4a",
      bodyTextColor: "#115e59",
      accentColors: ["#0d9488", "#14b8a6", "#059669", "#0891b2", "#65a30d"],
    },
  },
  {
    id: "amber",
    labelKey: "amber",
    theme: {
      background: "#fffbeb",
      fontFamily: "rounded",
      titleColor: "#78350f",
      bodyTextColor: "#92400e",
      accentColors: ["#f59e0b", "#eab308", "#d97706", "#ea580c", "#ca8a04"],
    },
  },
  {
    id: "rose",
    labelKey: "rose",
    theme: {
      background: "#fff1f2",
      fontFamily: "rounded",
      titleColor: "#881337",
      bodyTextColor: "#9f1239",
      accentColors: ["#e11d48", "#f43f5e", "#db2777", "#c026d3", "#ef4444"],
    },
  },
  {
    id: "indigo",
    labelKey: "indigo",
    theme: {
      background: "#eef2ff",
      fontFamily: "system",
      titleColor: "#312e81",
      bodyTextColor: "#3730a3",
      accentColors: ["#4f46e5", "#6366f1", "#7c3aed", "#2563eb", "#0891b2"],
    },
  },
  {
    id: "graphite",
    labelKey: "graphite",
    theme: {
      background: "#18181b",
      fontFamily: "mono",
      titleColor: "#fafafa",
      bodyTextColor: "#d4d4d8",
      accentColors: ["#a1a1aa", "#71717a", "#22d3ee", "#f59e0b", "#f43f5e"],
    },
  },
  {
    id: "citrus",
    labelKey: "citrus",
    theme: {
      background: "#f7fee7",
      fontFamily: "rounded",
      titleColor: "#3f6212",
      bodyTextColor: "#4d7c0f",
      accentColors: ["#84cc16", "#eab308", "#f97316", "#22c55e", "#14b8a6"],
    },
  },
  {
    id: "mint",
    labelKey: "mint",
    theme: {
      background: "#ecfdf5",
      fontFamily: "rounded",
      titleColor: "#064e3b",
      bodyTextColor: "#065f46",
      accentColors: ["#10b981", "#34d399", "#14b8a6", "#84cc16", "#0ea5e9"],
    },
  },
  {
    id: "coral",
    labelKey: "coral",
    theme: {
      background: "#fff7ed",
      fontFamily: "rounded",
      titleColor: "#7f1d1d",
      bodyTextColor: "#9a3412",
      accentColors: ["#f97316", "#fb7185", "#ef4444", "#f59e0b", "#db2777"],
    },
  },
  {
    id: "plum",
    labelKey: "plum",
    theme: {
      background: "#fdf4ff",
      fontFamily: "serif",
      titleColor: "#701a75",
      bodyTextColor: "#86198f",
      accentColors: ["#a21caf", "#c026d3", "#9333ea", "#be123c", "#6366f1"],
    },
  },
  {
    id: "ice",
    labelKey: "ice",
    theme: {
      background: "#f8fafc",
      fontFamily: "system",
      titleColor: "#164e63",
      bodyTextColor: "#155e75",
      accentColors: ["#22d3ee", "#38bdf8", "#67e8f9", "#818cf8", "#14b8a6"],
    },
  },
  {
    id: "parchment",
    labelKey: "parchment",
    theme: {
      background: "#fefce8",
      fontFamily: "serif",
      titleColor: "#422006",
      bodyTextColor: "#713f12",
      accentColors: ["#a16207", "#854d0e", "#78716c", "#b91c1c", "#4d7c0f"],
    },
  },
  bandTheme(
    "horizon",
    "#ffffff",
    "#172554",
    "#2563eb",
    "#93c5fd",
    180,
    0.88,
    0.96,
  ),
  bandTheme(
    "column",
    "#fafaf9",
    "#292524",
    "#78716c",
    "#d6d3d1",
    270,
    0.92,
    0.97,
  ),
  bandTheme(
    "terrace",
    "#f7fee7",
    "#14532d",
    "#65a30d",
    "#a3e635",
    180,
    0.86,
    0.94,
  ),
  bandTheme(
    "facet",
    "#f8fafc",
    "#1e293b",
    "#0891b2",
    "#164e63",
    135,
    0.84,
    0.92,
  ),
  bandTheme(
    "prism",
    "#fff7ed",
    "#431407",
    "#f97316",
    "#be123c",
    45,
    0.82,
    0.94,
  ),
  bandTheme(
    "studio",
    "#ffffff",
    "#18181b",
    "#18181b",
    "#f59e0b",
    0,
    0.94,
    0.97,
  ),
  bandTheme(
    "skyline",
    "#eff6ff",
    "#172554",
    "#1d4ed8",
    "#60a5fa",
    90,
    0.9,
    0.96,
  ),
  bandTheme(
    "dune",
    "#fffbeb",
    "#78350f",
    "#d97706",
    "#fcd34d",
    155,
    0.83,
    0.95,
  ),
  bandTheme(
    "tide",
    "#f0fdfa",
    "#134e4a",
    "#0d9488",
    "#5eead4",
    200,
    0.86,
    0.95,
  ),
  bandTheme(
    "ember",
    "#18181b",
    "#fafafa",
    "#ea580c",
    "#fbbf24",
    135,
    0.87,
    0.96,
  ),
  bandTheme(
    "aperture",
    "#0f172a",
    "#f8fafc",
    "#7c3aed",
    "#c4b5fd",
    315,
    0.9,
    0.97,
  ),
  bandTheme(
    "aurora",
    "#f5f3ff",
    "#3b0764",
    "#a855f7",
    "#2dd4bf",
    25,
    0.85,
    0.95,
  ),
] as const;


/** A slide ground: a linear sweep (CSS degrees) or a circular glow from a focus point (% of the slide). */
type ThemeDesignStop = readonly [color: string, position: number];
type ThemeDesignGround =
  | { angle: number; stops: readonly ThemeDesignStop[] }
  | { focus: readonly [x: number, y: number]; stops: readonly ThemeDesignStop[] };

/**
 * One designed preset: its ground (a solid paper, a soft sweep, a glow or
 * crisp native colour bands -- every one a DrawingML fill the download writes
 * as it is drawn), neutral-leaning title and body colours, six spaced
 * accents, and a font pair from Office's own Fonts gallery (heading and body,
 * Latin and Chinese), which the canvas and the download both name.
 */
function designedTheme(design: {
  id: string;
  paper: string;
  ground: ThemeDesignGround | null;
  title: string;
  body: string;
  accents: readonly string[];
  fonts: string;
  fontFamily: ThemeFontFamily;
  titleFontSize: number;
  bodyFontSize: number;
}): PresetTheme {
  const ground = design.ground;
  const stops = ground?.stops.map(([color, position]) => ({ color, position }));
  const path =
    ground && "focus" in ground
      ? {
          kind: "circle" as const,
          fillToRect: {
            l: Math.round(ground.focus[0] * 1000),
            t: Math.round(ground.focus[1] * 1000),
            r: Math.round((100 - ground.focus[0]) * 1000),
            b: Math.round((100 - ground.focus[1]) * 1000),
          },
        }
      : undefined;
  return {
    id: design.id,
    labelKey: design.id,
    theme: {
      background: design.paper,
      ...(ground && stops
        ? {
            backgroundGradient: {
              from: stops[0].color,
              to: stops[stops.length - 1].color,
              angle: "angle" in ground ? ground.angle : 0,
              stops,
              ...(path ? { path } : {}),
            },
          }
        : {}),
      fontFamily: design.fontFamily,
      fontSchemeId: design.fonts,
      titleColor: design.title,
      bodyTextColor: design.body,
      titleFontSize: design.titleFontSize,
      bodyFontSize: design.bodyFontSize,
      titleBold: true,
      accentColors: [...design.accents],
    },
  };
}

/** The white Bjorn card, identified from the Web and kept exactly. */
const LEGACY_BJORN: PresetTheme = {
  id: "bjorn",
  labelKey: "bjorn",
  theme: {
    background: "#ffffff",
    fontFamily: "system",
    titleColor: "#000000",
    bodyTextColor: "#000000",
    accentColors: ["#28ae70", "#ac9f0d", "#ef3475", "#8961f3", "#10add3", "#e78811"],
  },
};

/**
 * The built-in preset themes (Design ▸ Themes ▸ Custom). The first entry
 * (`none`) clears the theme back to the built-in defaults. Hex values are
 * already normalized (lower-case, six digits) so applying a preset is
 * byte-stable.
 */
export const PRESET_THEMES: readonly PresetTheme[] = [
  { id: "none", labelKey: "none", theme: {} },
  designedTheme({
    id: "classic",
    paper: "#ffffff",
    ground: { angle: 180, stops: [["#ffffff", 0], ["#f5f7fa", 1]] },
    title: "#0f172a",
    body: "#334155",
    accents: ["#2563eb", "#0d9488", "#f59e0b", "#e11d48", "#7c3aed", "#64748b"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "midnight",
    paper: "#0a0f1e",
    ground: { focus: [18, 12], stops: [["#1e2b4a", 0], ["#0f172a", 0.6], ["#0a0f1e", 1]] },
    title: "#f8fafc",
    body: "#c7d2e0",
    accents: ["#38bdf8", "#a78bfa", "#34d399", "#fbbf24", "#f472b6", "#94a3b8"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  LEGACY_BJORN,
  designedTheme({
    id: "sunrise",
    paper: "#fffaf4",
    ground: { angle: 165, stops: [["#fffaf4", 0], ["#fff1e4", 0.6], ["#ffe1cc", 1]] },
    title: "#431407",
    body: "#6b4532",
    accents: ["#f97316", "#e11d48", "#f59e0b", "#0ea5e9", "#14b8a6", "#a855f7"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "forest",
    paper: "#f8faf6",
    ground: { focus: [10, 100], stops: [["#e3ecdf", 0], ["#f4f7f1", 0.6], ["#f8faf6", 1]] },
    title: "#14301f",
    body: "#3d5446",
    accents: ["#2f6f4f", "#7a9a3a", "#c08a2e", "#2a7f8f", "#b4533a", "#5b6b5f"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "corporate",
    paper: "#ffffff",
    ground: { angle: 270, stops: [["#ffffff", 0], ["#ffffff", 0.985], ["#0b2545", 0.985], ["#0b2545", 1]] },
    title: "#0b2545",
    body: "#3b4a5c",
    accents: ["#1d4ed8", "#0891b2", "#0f766e", "#f59e0b", "#dc2626", "#64748b"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 40,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "vibrant",
    paper: "#5b21b6",
    ground: { angle: 135, stops: [["#5b21b6", 0], ["#9d174d", 1]] },
    title: "#ffffff",
    body: "#f5e8ff",
    accents: ["#fde047", "#5eead4", "#f9a8d4", "#a5b4fc", "#fdba74", "#ffffff"],
    fonts: "cjkSans",
    fontFamily: "system",
    titleFontSize: 44,
    bodyFontSize: 22,
  }),
  designedTheme({
    id: "minimal",
    paper: "#fafaf9",
    ground: null,
    title: "#1c1917",
    body: "#57534e",
    accents: ["#1c1917", "#78716c", "#d6d3d1", "#ea580c", "#0369a1", "#a8a29e"],
    fonts: "modern",
    fontFamily: "system",
    titleFontSize: 40,
    bodyFontSize: 18,
  }),
  designedTheme({
    id: "dark",
    paper: "#18181b",
    ground: { angle: 160, stops: [["#18181b", 0], ["#09090b", 1]] },
    title: "#fafafa",
    body: "#a1a1aa",
    accents: ["#3b82f6", "#22c55e", "#eab308", "#ef4444", "#a855f7", "#06b6d4"],
    fonts: "cjkSans",
    fontFamily: "system",
    titleFontSize: 40,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "ocean",
    paper: "#f3f8fd",
    ground: { angle: 180, stops: [["#f3f8fd", 0], ["#dcebf8", 1]] },
    title: "#0b2e4a",
    body: "#33536e",
    accents: ["#0369a1", "#0891b2", "#14b8a6", "#f59e0b", "#ef4444", "#6366f1"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "meadow",
    paper: "#fbfdf5",
    ground: { angle: 180, stops: [["#fbfdf5", 0], ["#eef6dc", 1]] },
    title: "#24350f",
    body: "#4a5a2c",
    accents: ["#65a30d", "#0d9488", "#eab308", "#f97316", "#0284c7", "#a855f7"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "sand",
    paper: "#f7f2e8",
    ground: null,
    title: "#33261a",
    body: "#5c4a3a",
    accents: ["#b45309", "#0f766e", "#be123c", "#4d7c0f", "#1d4ed8", "#a16207"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "clay",
    paper: "#f8f0ea",
    ground: { angle: 150, stops: [["#f8f0ea", 0], ["#efdccf", 1]] },
    title: "#4a2416",
    body: "#6b4a3a",
    accents: ["#c2410c", "#0e7490", "#a16207", "#9f1239", "#4d7c0f", "#78716c"],
    fonts: "gill-sans-mt",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "lavender",
    paper: "#faf7ff",
    ground: { angle: 160, stops: [["#faf7ff", 0], ["#ece5fb", 1]] },
    title: "#2e1f57",
    body: "#5b4b7a",
    accents: ["#7c3aed", "#db2777", "#0ea5e9", "#14b8a6", "#f59e0b", "#a78bfa"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "berry",
    paper: "#1f0718",
    ground: { focus: [15, 90], stops: [["#5b1340", 0], ["#2d0a22", 0.7], ["#1f0718", 1]] },
    title: "#fdf2f8",
    body: "#f5c2dc",
    accents: ["#f472b6", "#fb7185", "#c084fc", "#fbbf24", "#2dd4bf", "#f9a8d4"],
    fonts: "gill-sans-mt",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "slate",
    paper: "#f8fafc",
    ground: { angle: 160, stops: [["#f8fafc", 0], ["#e9eef4", 1]] },
    title: "#0f172a",
    body: "#334155",
    accents: ["#334155", "#0ea5e9", "#14b8a6", "#f59e0b", "#f43f5e", "#8b5cf6"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "sky",
    paper: "#ffffff",
    ground: { focus: [90, 5], stops: [["#d6ecff", 0], ["#eef7ff", 0.5], ["#ffffff", 1]] },
    title: "#0c3a5e",
    body: "#3b5b78",
    accents: ["#0ea5e9", "#6366f1", "#14b8a6", "#f59e0b", "#f43f5e", "#64748b"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "teal",
    paper: "#f2fbfa",
    ground: { angle: 165, stops: [["#f2fbfa", 0], ["#d7f1ec", 1]] },
    title: "#0f3d3e",
    body: "#36605e",
    accents: ["#0d9488", "#0891b2", "#f59e0b", "#e11d48", "#6366f1", "#84cc16"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "amber",
    paper: "#fffcf5",
    ground: { angle: 170, stops: [["#fffcf5", 0], ["#fdf0d5", 1]] },
    title: "#3b2410",
    body: "#5f4423",
    accents: ["#d97706", "#b91c1c", "#0f766e", "#2563eb", "#7c3aed", "#a16207"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "rose",
    paper: "#fff7f8",
    ground: { angle: 170, stops: [["#fff7f8", 0], ["#fde4e8", 1]] },
    title: "#4c0519",
    body: "#6b2c3a",
    accents: ["#e11d48", "#f97316", "#8b5cf6", "#0ea5e9", "#10b981", "#fb7185"],
    fonts: "gill-sans-mt",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "indigo",
    paper: "#14123a",
    ground: { focus: [80, 0], stops: [["#3730a3", 0], ["#1e1b4b", 0.65], ["#14123a", 1]] },
    title: "#eef2ff",
    body: "#c7d2fe",
    accents: ["#818cf8", "#f0abfc", "#67e8f9", "#fcd34d", "#86efac", "#c4b5fd"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "graphite",
    paper: "#2a2a2e",
    ground: { angle: 180, stops: [["#2a2a2e", 0], ["#1c1c1f", 1]] },
    title: "#f5f5f7",
    body: "#c7c7cc",
    accents: ["#0a84ff", "#30d158", "#ff9f0a", "#ff375f", "#bf5af2", "#64d2ff"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "citrus",
    paper: "#fffef6",
    ground: { angle: 180, stops: [["#fffef6", 0], ["#fffef6", 0.96], ["#84cc16", 0.96], ["#84cc16", 0.985], ["#f97316", 0.985], ["#f97316", 1]] },
    title: "#1f2d0b",
    body: "#465227",
    accents: ["#65a30d", "#f97316", "#eab308", "#0891b2", "#e11d48", "#84cc16"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "mint",
    paper: "#fafffc",
    ground: { focus: [90, 95], stops: [["#d4f5e6", 0], ["#effcf6", 0.55], ["#fafffc", 1]] },
    title: "#06402f",
    body: "#2f5d4c",
    accents: ["#10b981", "#0ea5e9", "#84cc16", "#f59e0b", "#f43f5e", "#14b8a6"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "coral",
    paper: "#fff8f5",
    ground: { angle: 180, stops: [["#fff8f5", 0], ["#fff8f5", 0.965], ["#f26b5b", 0.965], ["#f26b5b", 1]] },
    title: "#4a1d1f",
    body: "#6b3b3a",
    accents: ["#f26b5b", "#ff9f68", "#2a9d8f", "#3d5a80", "#e9c46a", "#9b5de5"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "plum",
    paper: "#f8f3f8",
    ground: null,
    title: "#3b0a3f",
    body: "#5e3a60",
    accents: ["#86198f", "#be123c", "#0e7490", "#ca8a04", "#4f46e5", "#a21caf"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "ice",
    paper: "#ffffff",
    ground: { focus: [85, 10], stops: [["#e0f2fe", 0], ["#f8fcff", 0.55], ["#ffffff", 1]] },
    title: "#0c3b52",
    body: "#3f6275",
    accents: ["#0891b2", "#38bdf8", "#6366f1", "#14b8a6", "#f59e0b", "#94a3b8"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "parchment",
    paper: "#ead9b8",
    ground: { focus: [50, 45], stops: [["#fbf6ea", 0], ["#f3ead3", 0.75], ["#ead9b8", 1]] },
    title: "#3b2f20",
    body: "#5a4a35",
    accents: ["#8b5a2b", "#7a1f1f", "#2f5d50", "#a77c1f", "#3c4f76", "#6b5b4b"],
    fonts: "cjkSerif",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "horizon",
    paper: "#ffffff",
    ground: { angle: 180, stops: [["#ffffff", 0], ["#ffffff", 0.9], ["#1d4ed8", 0.9], ["#1d4ed8", 0.965], ["#93c5fd", 0.965], ["#93c5fd", 1]] },
    title: "#0b1f4d",
    body: "#33415c",
    accents: ["#1d4ed8", "#0ea5e9", "#14b8a6", "#f59e0b", "#ef4444", "#6366f1"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "column",
    paper: "#fafaf9",
    ground: { angle: 270, stops: [["#fafaf9", 0], ["#fafaf9", 0.94], ["#292524", 0.94], ["#292524", 0.965], ["#a8a29e", 0.965], ["#a8a29e", 1]] },
    title: "#1c1917",
    body: "#57534e",
    accents: ["#292524", "#ea580c", "#0369a1", "#65a30d", "#be123c", "#a8a29e"],
    fonts: "modern",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "terrace",
    paper: "#f7fbf2",
    ground: { angle: 180, stops: [["#f7fbf2", 0], ["#f7fbf2", 0.9], ["#4d7c0f", 0.9], ["#4d7c0f", 0.96], ["#bef264", 0.96], ["#bef264", 1]] },
    title: "#1a2e05",
    body: "#3f5222",
    accents: ["#4d7c0f", "#a16207", "#0f766e", "#ea580c", "#2563eb", "#84cc16"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "facet",
    paper: "#f8fafc",
    ground: { angle: 135, stops: [["#f8fafc", 0], ["#f8fafc", 0.93], ["#0e7490", 0.93], ["#0e7490", 0.97], ["#164e63", 0.97], ["#164e63", 1]] },
    title: "#0f2a36",
    body: "#36505c",
    accents: ["#0e7490", "#164e63", "#f59e0b", "#e11d48", "#6366f1", "#22c55e"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "prism",
    paper: "#fffaf5",
    ground: { angle: 225, stops: [["#fffaf5", 0], ["#fffaf5", 0.93], ["#f97316", 0.93], ["#f97316", 0.97], ["#be123c", 0.97], ["#be123c", 1]] },
    title: "#3b1206",
    body: "#6b3a2a",
    accents: ["#f97316", "#be123c", "#7c3aed", "#0ea5e9", "#14b8a6", "#eab308"],
    fonts: "cjkSans",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "studio",
    paper: "#ffffff",
    ground: { angle: 270, stops: [["#ffffff", 0], ["#ffffff", 0.955], ["#111111", 0.955], ["#111111", 0.975], ["#f59e0b", 0.975], ["#f59e0b", 1]] },
    title: "#111111",
    body: "#4a4a4a",
    accents: ["#111111", "#f59e0b", "#2563eb", "#10b981", "#ef4444", "#8b5cf6"],
    fonts: "cjkSans",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "skyline",
    paper: "#f5f9ff",
    ground: { angle: 180, stops: [["#f5f9ff", 0], ["#f5f9ff", 0.92], ["#1e40af", 0.92], ["#1e40af", 0.97], ["#60a5fa", 0.97], ["#60a5fa", 1]] },
    title: "#0b1f4d",
    body: "#334a6e",
    accents: ["#1e40af", "#60a5fa", "#0d9488", "#f59e0b", "#e11d48", "#8b5cf6"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "dune",
    paper: "#fdf8ee",
    ground: { angle: 155, stops: [["#fdf8ee", 0], ["#fdf8ee", 0.93], ["#d97706", 0.93], ["#d97706", 0.97], ["#fcd34d", 0.97], ["#fcd34d", 1]] },
    title: "#3d2208",
    body: "#6b4a28",
    accents: ["#d97706", "#b45309", "#0f766e", "#be123c", "#2563eb", "#a16207"],
    fonts: "classic",
    fontFamily: "serif",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "tide",
    paper: "#f2fbf9",
    ground: { angle: 200, stops: [["#f2fbf9", 0], ["#f2fbf9", 0.93], ["#0d9488", 0.93], ["#0d9488", 0.97], ["#5eead4", 0.97], ["#5eead4", 1]] },
    title: "#08343a",
    body: "#2f5b5a",
    accents: ["#0d9488", "#0891b2", "#5eead4", "#f59e0b", "#f43f5e", "#6366f1"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "ember",
    paper: "#171717",
    ground: { angle: 160, stops: [["#171717", 0], ["#171717", 0.93], ["#ea580c", 0.93], ["#ea580c", 0.97], ["#fbbf24", 0.97], ["#fbbf24", 1]] },
    title: "#fafafa",
    body: "#d4d4d4",
    accents: ["#f97316", "#fbbf24", "#ef4444", "#38bdf8", "#a3e635", "#e879f9"],
    fonts: "cjkSans",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "aperture",
    paper: "#0f172a",
    ground: { angle: 135, stops: [["#0f172a", 0], ["#0f172a", 0.93], ["#7c3aed", 0.93], ["#7c3aed", 0.97], ["#c4b5fd", 0.97], ["#c4b5fd", 1]] },
    title: "#f8fafc",
    body: "#cbd5e1",
    accents: ["#a78bfa", "#38bdf8", "#34d399", "#fbbf24", "#f472b6", "#c4b5fd"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
  designedTheme({
    id: "aurora",
    paper: "#f5f3ff",
    ground: { angle: 25, stops: [["#f5f3ff", 0], ["#f5f3ff", 0.78], ["#c4b5fd", 0.9], ["#5eead4", 1]] },
    title: "#2e1065",
    body: "#4c3a75",
    accents: ["#8b5cf6", "#14b8a6", "#ec4899", "#0ea5e9", "#f59e0b", "#a855f7"],
    fonts: "office",
    fontFamily: "system",
    titleFontSize: 42,
    bodyFontSize: 20,
  }),
] as const;

/** The four real, base-relative variants exposed for every available theme. */
export const PRESENTATION_THEME_VARIANTS: readonly {
  id: PresentationThemeVariantId;
  labelKey:
    | "presentation.theme.variant.base"
    | "presentation.theme.variant.soft"
    | "presentation.theme.variant.deep"
    | "presentation.theme.variant.contrast";
}[] = [
  { id: "base", labelKey: "presentation.theme.variant.base" },
  { id: "soft", labelKey: "presentation.theme.variant.soft" },
  { id: "deep", labelKey: "presentation.theme.variant.deep" },
  { id: "contrast", labelKey: "presentation.theme.variant.contrast" },
];

/**
 * Type guard: is `value` exactly one of the allowed font-family identifiers?
 */
export function isThemeFontFamily(value: unknown): value is ThemeFontFamily {
  return typeof value === "string" && THEME_FONT_FAMILY_SET.has(value);
}

/**
 * Normalize an arbitrary colour value into a six-digit lower-case hex string, or
 * return `undefined` when the value is missing / malformed. Mirrors the slide
 * background validation so a hand-edited or stale value can never inject
 * arbitrary CSS. Trims surrounding whitespace before validating.
 */
export function normalizeThemeColor(value: unknown): string | undefined {
  if (typeof value === "string" && HEX_COLOR_RE.test(value.trim())) {
    return value.trim().toLowerCase();
  }
  return undefined;
}

/**
 * Normalize an arbitrary persisted font value into an allowed font-family key,
 * or `undefined` when missing / unknown. Case-insensitive + trims whitespace so
 * a stale value can never reach the renderer. Returns `undefined` (rather than
 * the default) so a theme that omits the font keeps the field absent and the
 * resolver can fall through to the built-in default explicitly.
 */
export function normalizeThemeFontFamily(
  value: unknown,
): ThemeFontFamily | undefined {
  if (typeof value !== "string") return undefined;
  const candidate = value.trim().toLowerCase();
  return isThemeFontFamily(candidate) ? candidate : undefined;
}

/**
 * Normalize an arbitrary persisted font size into a finite, rounded px value
 * inside [{@link MIN_THEME_FONT_SIZE}, {@link MAX_THEME_FONT_SIZE}], or
 * `undefined` when missing / non-finite / out of range. Accepts a numeric string
 * (e.g. a stale `"32"`) so a hand-edited theme still validates. A bogus value
 * (0 / NaN / Infinity / negative / absurd) is DROPPED (not clamped to a bound) so
 * an absent size stays distinct from an explicit one and the resolver falls
 * through to the built-in default size.
 */
export function normalizeThemeFontSize(value: unknown): number | undefined {
  let n: number | undefined;
  if (typeof value === "number") n = value;
  else if (typeof value === "string" && value.trim() !== "") n = Number(value.trim());
  if (n === undefined || !Number.isFinite(n)) return undefined;
  const rounded = Math.round(n);
  if (rounded < MIN_THEME_FONT_SIZE || rounded > MAX_THEME_FONT_SIZE) {
    return undefined;
  }
  return rounded;
}




















export function normalizeThemeAccentColors(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const out: string[] = [];
  for (const entry of value) {
    if (out.length >= MAX_THEME_ACCENT_COLORS) break;
    const color = normalizeThemeColor(entry);
    if (!color) continue;
    out.push(color);
  }
  return out.length > 0 ? out : undefined;
}

const THEME_VARIANT_ID_SET: ReadonlySet<string> = new Set(
  PRESENTATION_THEME_VARIANT_IDS,
);
const THEME_COLOR_SLOT_SET: ReadonlySet<string> = new Set(
  PRESENTATION_THEME_COLOR_SLOTS,
);
const THEME_ID_RE = /^[a-z0-9][a-z0-9._:-]{0,159}$/i;
const SHA256_RE = /^[0-9a-f]{64}$/;
const SAFE_NATIVE_FONT_RE = /^[^<>&\u0000-\u001f\u007f]{1,200}$/;

/**
 * Validate the bounded ZIP-part shape of a native presentation theme receipt.
 *
 * The semantic authority is the slide-master relationship graph on the server;
 * PresentationML does not require a theme target to live under `ppt/theme`.
 * The browser therefore keeps any normalized `ppt/...xml` target while still
 * rejecting traversal, URI fragments, backslashes, empty segments and controls.
 */
export function isNativePresentationThemePartPath(
  value: unknown,
): value is string {
  if (typeof value !== "string" || value.length < 9 || value.length > 512) {
    return false;
  }
  if (
    value !== value.trim() ||
    !value.startsWith("ppt/") ||
    !value.endsWith(".xml") ||
    /[\\?#\u0000-\u001f\u007f]/.test(value)
  ) {
    return false;
  }
  const segments = value.split("/");
  return segments.every(
    (segment) => segment.length > 0 && segment !== "." && segment !== "..",
  );
}

function normalizeThemeTokenSnapshot(
  raw: unknown,
): DeckThemeTokenSnapshot | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const result: DeckThemeTokenSnapshot = {};
  const background = normalizeThemeColor(record.background);
  const accents = normalizeThemeAccentColors(record.accentColors);
  const fontFamily = normalizeThemeFontFamily(record.fontFamily);
  const titleColor = normalizeThemeColor(record.titleColor);
  const bodyTextColor = normalizeThemeColor(record.bodyTextColor);
  const titleFontSize = normalizeThemeFontSize(record.titleFontSize);
  const bodyFontSize = normalizeThemeFontSize(record.bodyFontSize);
  if (background) result.background = background;
  if (
    record.backgroundGradient &&
    typeof record.backgroundGradient === "object" &&
    !Array.isArray(record.backgroundGradient)
  ) {
    const gradient = normalizeBgGradient(record.backgroundGradient);
    if (gradient) result.backgroundGradient = gradient;
  }
  if (
    record.backgroundPattern &&
    typeof record.backgroundPattern === "object" &&
    !Array.isArray(record.backgroundPattern)
  ) {
    result.backgroundPattern = record.backgroundPattern;
  }
  if (accents) result.accentColors = accents;
  if (fontFamily) result.fontFamily = fontFamily;
  if (titleColor) result.titleColor = titleColor;
  if (bodyTextColor) result.bodyTextColor = bodyTextColor;
  if (titleFontSize !== undefined) result.titleFontSize = titleFontSize;
  if (bodyFontSize !== undefined) result.bodyFontSize = bodyFontSize;
  if (record.titleBold === true) result.titleBold = true;
  if (record.bodyBold === true) result.bodyBold = true;
  if (record.titleItalic === true) result.titleItalic = true;
  if (record.bodyItalic === true) result.bodyItalic = true;
  return Object.keys(result).length > 0 ? result : undefined;
}

function normalizeNativeThemeSource(
  raw: unknown,
): NativeDeckThemeSource | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const part = typeof record.part === "string" ? record.part.trim() : "";
  const fingerprint =
    typeof record.fingerprint === "string"
      ? record.fingerprint.trim().toLowerCase()
      : "";
  if (
    !isNativePresentationThemePartPath(part) ||
    !SHA256_RE.test(fingerprint)
  ) {
    return undefined;
  }
  const result: NativeDeckThemeSource = { part, fingerprint };
  if (
    Number.isInteger(record.themeCount) &&
    Number(record.themeCount) >= 1 &&
    Number(record.themeCount) <= 64
  ) {
    result.themeCount = Number(record.themeCount);
  }
  if (typeof record.name === "string" && record.name.trim()) {
    // Verbatim: the receipt is the package's theme name, which the server
    
    
    // two themes refused it (multiple_native_themes_not_patchable): every
    // download was rebuilt from the page.
    result.name = record.name.slice(0, 200);
  }
  if (
    record.colorSlots &&
    typeof record.colorSlots === "object" &&
    !Array.isArray(record.colorSlots)
  ) {
    const colorSlots: Partial<Record<PresentationThemeColorSlot, string>> = {};
    for (const [slot, value] of Object.entries(record.colorSlots)) {
      if (!THEME_COLOR_SLOT_SET.has(slot)) continue;
      const color = normalizeThemeColor(value);
      if (color) {
        colorSlots[slot as PresentationThemeColorSlot] = color;
      }
    }
    if (Object.keys(colorSlots).length > 0) result.colorSlots = colorSlots;
  }
  for (const [key, value] of [
    ["majorLatinFont", record.majorLatinFont],
    ["minorLatinFont", record.minorLatinFont],
  ] as const) {
    const candidate = typeof value === "string" ? value.trim() : "";
    if (SAFE_NATIVE_FONT_RE.test(candidate)) result[key] = candidate;
  }
  const baseTheme = normalizeThemeTokenSnapshot(record.baseTheme);
  if (baseTheme) result.baseTheme = baseTheme;
  if (
    typeof record.masterDesignFingerprint === "string" &&
    SHA256_RE.test(record.masterDesignFingerprint.trim().toLowerCase())
  ) {
    result.masterDesignFingerprint = record.masterDesignFingerprint
      .trim()
      .toLowerCase();
  }
  if (Array.isArray(record.themeParts) && record.themeParts.length >= 2 && record.themeParts.length <= 64) {
    const parts = record.themeParts.flatMap((entry: unknown) => {
      const item = entry && typeof entry === "object" ? (entry as Record<string, unknown>) : null;
      const part = typeof item?.part === "string" ? item.part : "";
      const fingerprint = typeof item?.fingerprint === "string" ? item.fingerprint : "";
      const master = typeof item?.masterDesignFingerprint === "string" ? item.masterDesignFingerprint : "";
      return isNativePresentationThemePartPath(part) && SHA256_RE.test(fingerprint) && SHA256_RE.test(master)
        ? [{ part, fingerprint, masterDesignFingerprint: master }]
        : [];
    });
    if (parts.length === record.themeParts.length) result.themeParts = parts;
  }
  if (typeof record.effectSchemeMatch === "string" && officeEffectGalleryScheme(record.effectSchemeMatch)) {
    result.effectSchemeMatch = record.effectSchemeMatch;
  }
  return result;
}

/**
 * The entry Design > Variants > Effects checks (PPT-A073): a pick
 * (`effectSchemeId`); else an Office theme's own effects; else, on a deck
 * still on its file's own theme, the entry that theme's scheme is
 * (`identity.source.effectSchemeMatch`, reading row 851); else none. A theme
 * picked since keeps the file's receipt, which then no longer says what the
 * deck wears, so it is read only while the selection is the file's own.
 */
export function presentationEffectGalleryValue(
  theme: DeckTheme | null | undefined,
  officeThemeEffects: (themeId: string) => string | undefined,
): string {
  if (theme?.effectSchemeId) return theme.effectSchemeId;
  const identity = theme?.identity;
  if (identity?.selectionKind === "office") return officeThemeEffects(identity.themeId) ?? "";
  return identity?.selectionKind === "native" ? (identity.source?.effectSchemeMatch ?? "") : "";
}

/**
 * A complete twelve-slot colour scheme, lower-cased, in slot order; anything
 * short of all twelve valid slots is dropped (a partial scheme cannot be
 * written as `a:clrScheme`).
 */
export function normalizeThemeColorSlots(
  raw: unknown,
): Record<PresentationThemeColorSlot, string> | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const result = {} as Record<PresentationThemeColorSlot, string>;
  for (const slot of PRESENTATION_THEME_COLOR_SLOTS) {
    const color = normalizeThemeColor(record[slot]);
    if (!color) return undefined;
    result[slot] = color;
  }
  return result;
}

/**
 * A master's colour map: all four of bg1 / tx1 / bg2 / tx2, each naming one
 * of the twelve slots; anything else is dropped.
 */
export function normalizeThemeColorMap(raw: unknown): DeckThemeColorMap | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (Object.keys(record).length !== DECK_THEME_MAPPED_COLORS.length) return undefined;
  const result = {} as Record<(typeof DECK_THEME_MAPPED_COLORS)[number], PresentationThemeColorSlot>;
  for (const name of DECK_THEME_MAPPED_COLORS) {
    const slot = record[name];
    if (typeof slot !== "string" || !(PRESENTATION_THEME_COLOR_SLOTS as readonly string[]).includes(slot)) {
      return undefined;
    }
    result[name] = slot as PresentationThemeColorSlot;
  }
  return result;
}

/** A theme font scheme by face name; both Latin faces are required. */
export function normalizeThemeFontScheme(
  raw: unknown,
): DeckThemeFontScheme | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const face = (value: unknown) => {
    const candidate = typeof value === "string" ? value.trim() : "";
    return SAFE_NATIVE_FONT_RE.test(candidate) ? candidate : undefined;
  };
  const majorLatin = face(record.majorLatin);
  const minorLatin = face(record.minorLatin);
  if (!majorLatin || !minorLatin) return undefined;
  const result: DeckThemeFontScheme = { majorLatin, minorLatin };
  const majorHans = face(record.majorHans);
  const minorHans = face(record.minorHans);
  if (majorHans) result.majorHans = majorHans;
  if (minorHans) result.minorHans = minorHans;
  const name = typeof record.name === "string" ? record.name.trim() : "";
  // A name is shown, never parsed: any printable text Office would keep.
  if (name && name.length <= MAX_THEME_FONT_SCHEME_NAME && !/[\u0000-\u001f\u007f]/.test(name)) result.name = name;
  return result;
}

/** Normalize a persisted theme identity without manufacturing native authority. */
export function normalizeThemeIdentity(
  raw: unknown,
): DeckThemeIdentity | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (record.schemaVersion !== PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION) {
    return undefined;
  }
  const selectionKind = record.selectionKind;
  const themeId = typeof record.themeId === "string" ? record.themeId.trim() : "";
  const variantId = record.variantId;
  if (
    !["catalog", "office", "native", "custom"].includes(String(selectionKind)) ||
    !THEME_ID_RE.test(themeId) ||
    typeof variantId !== "string" ||
    !THEME_VARIANT_ID_SET.has(variantId)
  ) {
    return undefined;
  }
  // An Office identity names a catalog theme and one of ITS variants; every
  // other kind keeps the four computed variants.
  if (selectionKind === "office") {
    if (!presentationOfficeThemeVariant(themeId, variantId)) return undefined;
  } else if (!isThemeDesignVariantId(variantId)) {
    return undefined;
  }
  if (
    selectionKind === "catalog" &&
    !PRESET_THEMES.some((preset) => preset.id === themeId && preset.id !== "none")
  ) {
    return undefined;
  }
  const source = normalizeNativeThemeSource(record.source);
  if (selectionKind === "native" && !source) return undefined;
  return {
    schemaVersion: PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
    selectionKind: selectionKind as DeckThemeIdentity["selectionKind"],
    themeId,
    variantId: variantId as PresentationThemeVariantId,
    ...(source ? { source } : {}),
    ...(selectionKind === "office" && source && record.origin === "package" ? { origin: "package" as const } : {}),
  };
}

/**
 * Normalize an arbitrary persisted theme record into a clean {@link DeckTheme},
 * keeping ONLY the fields that validate. Invalid colours and unknown fonts are
 * dropped (never coerced to a bogus value). Returns `undefined` when the input
 * is missing / not an object / yields no valid field, so callers can store an
 * absent `theme` and keep default decks byte-clean.
 */
export function normalizeTheme(raw: unknown): DeckTheme | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const theme: DeckTheme = {};
  const identity = normalizeThemeIdentity(record.identity);
  if (identity) theme.identity = identity;
  const background = normalizeThemeColor(record.background);
  if (background) theme.background = background;
  // Share gradient normalization with the canvas and thumbnail renderer.
  if (
    record.backgroundGradient &&
    typeof record.backgroundGradient === "object" &&
    !Array.isArray(record.backgroundGradient)
  ) {
    const gradient = normalizeBgGradient(record.backgroundGradient);
    if (gradient) theme.backgroundGradient = gradient;
  }
  if (
    record.backgroundImage &&
    typeof record.backgroundImage === "object" &&
    !Array.isArray(record.backgroundImage)
  ) {
    theme.backgroundImage = record.backgroundImage;
  }
  if (
    record.backgroundPattern &&
    typeof record.backgroundPattern === "object" &&
    !Array.isArray(record.backgroundPattern)
  ) {
    theme.backgroundPattern = record.backgroundPattern;
  }
  const accentColors = normalizeThemeAccentColors(record.accentColors);
  if (accentColors) theme.accentColors = accentColors;
  const fontFamily = normalizeThemeFontFamily(record.fontFamily);
  if (fontFamily) theme.fontFamily = fontFamily;
  if (officeThemeFontGalleryEntry(record.fontSchemeId)) {
    theme.fontSchemeId = record.fontSchemeId as string;
  }
  if (officeColorGalleryScheme(record.colorSchemeId)) {
    theme.colorSchemeId = record.colorSchemeId as string;
  }
  if (officeEffectGalleryScheme(record.effectSchemeId)) {
    theme.effectSchemeId = record.effectSchemeId as string;
  }
  const colorSlots = normalizeThemeColorSlots(record.colorSlots);
  if (colorSlots) theme.colorSlots = colorSlots;
  const colorMap = normalizeThemeColorMap(record.colorMap);
  if (colorMap) theme.colorMap = colorMap;
  const fontScheme = normalizeThemeFontScheme(record.fontScheme);
  if (fontScheme) theme.fontScheme = fontScheme;
  const titleColor = normalizeThemeColor(record.titleColor);
  if (titleColor) theme.titleColor = titleColor;
  const bodyTextColor = normalizeThemeColor(record.bodyTextColor);
  if (bodyTextColor) theme.bodyTextColor = bodyTextColor;
  const titleFontSize = normalizeThemeFontSize(record.titleFontSize);
  if (titleFontSize !== undefined) theme.titleFontSize = titleFontSize;
  const bodyFontSize = normalizeThemeFontSize(record.bodyFontSize);
  if (bodyFontSize !== undefined) theme.bodyFontSize = bodyFontSize;
  // Booleans are kept ONLY when explicitly `true`. A `false` (or any non-boolean)
  // contributes nothing — identical to the field being absent — so a theme never
  // carries a redundant `*Bold: false` and a default deck stays byte-clean.
  if (record.titleBold === true) theme.titleBold = true;
  if (record.bodyBold === true) theme.bodyBold = true;
  if (record.titleItalic === true) theme.titleItalic = true;
  if (record.bodyItalic === true) theme.bodyItalic = true;
  const officeMaster = normalizeDeckOfficeMaster(record.officeMaster);
  if (officeMaster) theme.officeMaster = officeMaster;
  return Object.keys(theme).length > 0 ? theme : undefined;
}

const THEME_VARIANT_FIELDS = [
  "background",
  "backgroundGradient",
  "backgroundPattern",
  "accentColors",
  "fontFamily",
  "titleColor",
  "bodyTextColor",
  "titleFontSize",
  "bodyFontSize",
  "titleBold",
  "bodyBold",
] as const satisfies readonly (keyof DeckThemeTokenSnapshot)[];

function themeWithoutIdentity(theme: DeckTheme | undefined): DeckTheme | undefined {
  const normalized = normalizeTheme(theme);
  if (!normalized) return undefined;
  const { identity: _identity, ...tokens } = normalized;
  return Object.keys(tokens).length > 0 ? tokens : undefined;
}

function presetForLegacyTheme(theme: DeckTheme | undefined): PresetTheme | undefined {
  const current = themeWithoutIdentity(theme);
  if (!current) return undefined;
  const currentJson = JSON.stringify(current);
  // A deck that predates theme identity states a preset's values of its day:
  // the current ones, or the first generation's (frozen for this only).
  const match = [...PRESET_THEMES, ...LEGACY_PRESET_THEMES].find(
    (preset) =>
      preset.id !== "none" &&
      JSON.stringify(themeWithoutIdentity(preset.theme)) === currentJson,
  );
  return match ? PRESET_THEMES.find((preset) => preset.id === match.id) : undefined;
}

/** Current gallery identity, with exact legacy-preset recovery only. */
export function resolveThemeSelection(
  theme: DeckTheme | null | undefined,
): Pick<DeckThemeIdentity, "selectionKind" | "themeId" | "variantId"> | null {
  const normalized = normalizeTheme(theme);
  if (normalized?.identity) {
    const { selectionKind, themeId, variantId } = normalized.identity;
    return { selectionKind, themeId, variantId };
  }
  const preset = presetForLegacyTheme(normalized);
  return preset
    ? { selectionKind: "catalog", themeId: preset.id, variantId: "base" }
    : null;
}

function themeBaseTokens(theme: DeckTheme | undefined): DeckThemeTokenSnapshot | undefined {
  const normalized = normalizeTheme(theme);
  const identity = normalized?.identity;
  if (identity?.selectionKind === "catalog") {
    return resolveCatalogThemeDesign(identity.themeId, "base");
  }
  if (identity?.selectionKind === "native") {
    return normalizeThemeTokenSnapshot(identity.source?.baseTheme);
  }
  const legacyPreset = presetForLegacyTheme(normalized);
  return legacyPreset
    ? resolveCatalogThemeDesign(legacyPreset.id, "base")
    : undefined;
}

function mixThemeColor(color: string, target: "#000000" | "#ffffff", weight: number) {
  const source = normalizeThemeColor(color);
  if (!source) return color;
  const mixed = [1, 3, 5]
    .map((offset) => {
      const from = Number.parseInt(source.slice(offset, offset + 2), 16);
      const to = target === "#ffffff" ? 255 : 0;
      return Math.round(from + (to - from) * weight)
        .toString(16)
        .padStart(2, "0");
    })
    .join("");
  return `#${mixed}`;
}

/** `color` moved `weight` of the way to any other colour. */
function blendThemeColors(color: string, target: string, weight: number): string {
  const from = normalizeThemeColor(color);
  const to = normalizeThemeColor(target);
  if (!from || !to) return color;
  return `#${[1, 3, 5]
    .map((offset) => {
      const a = Number.parseInt(from.slice(offset, offset + 2), 16);
      const b = Number.parseInt(to.slice(offset, offset + 2), 16);
      return Math.round(a + (b - a) * weight).toString(16).padStart(2, "0");
    })
    .join("")}`;
}

function themeColorIsLight(color: string | undefined): boolean {
  const value = normalizeThemeColor(color);
  if (!value) return true;
  const red = Number.parseInt(value.slice(1, 3), 16);
  const green = Number.parseInt(value.slice(3, 5), 16);
  const blue = Number.parseInt(value.slice(5, 7), 16);
  return red * 0.2126 + green * 0.7152 + blue * 0.0722 >= 140;
}

/** WCAG relative luminance of a hex colour (0 black .. 1 white). */
function themeLuminance(color: string): number {
  const value = normalizeThemeColor(color) ?? "#ffffff";
  const channel = (offset: number) => {
    const c = Number.parseInt(value.slice(offset, offset + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5);
}

function themeContrast(a: string, b: string): number {
  const [high, low] = [themeLuminance(a), themeLuminance(b)].sort((x, y) => y - x);
  return (high + 0.05) / (low + 0.05);
}

/** `color`, moved toward black or white until it reads on `ground` at `ratio`. */
function readableThemeColor(color: string, ground: string, ratio: number): string {
  const target = themeLuminance(ground) > 0.4 ? "#000000" : "#ffffff";
  for (let step = 0; step <= 20; step += 1) {
    const candidate = mixThemeColor(color, target, step / 20);
    if (themeContrast(candidate, ground) >= ratio) return candidate;
  }
  return target;
}

/**
 * A variant of a designed theme, by colour alone (fonts and sizes stay the
 * theme's, as PowerPoint's Variants keep a theme's fonts):
 * - `soft`: airier -- the ground toward white (a dark theme lifts a little),
 *   its colour bands lightened;
 * - `deep`: the dark counterpart of a light theme (its title colour, deepened,
 *   as the ground) and the light counterpart of a dark one (a pale tint of
 *   its ground);
 * - `contrast`: a bold ground of the theme's strongest dark accent.
 * Text is re-checked against the new ground (titles 7:1, body 4.5:1) and the
 * bands of a native band design keep their places.
 */
function designedThemeVariant(
  base: DeckThemeTokenSnapshot,
  variantId: Exclude<PresentationThemeDesignVariantId, "base">,
): DeckThemeTokenSnapshot {
  const gradient = normalizeBgGradient(base.backgroundGradient);
  const paper =
    normalizeThemeColor(base.background) ??
    gradient?.stops?.[gradient.stops.length - 1]?.color ??
    gradient?.to ??
    "#ffffff";
  const darkPaper = themeLuminance(paper) < 0.2;
  const accents = base.accentColors?.length ? base.accentColors : ["#2563eb"];
  const title = normalizeThemeColor(base.titleColor) ?? (darkPaper ? "#ffffff" : "#111827");
  const body = normalizeThemeColor(base.bodyTextColor) ?? title;
  const stops = gradient?.stops ?? [];
  const banded = stops.some((stop, index) => index > 0 && stop.position === stops[index - 1].position);
  const isBand = (color: string) => banded && color !== paper;
  const recolor = (map: (color: string, index: number) => string) => {
    const nextPaper = map(paper, 0);
    if (!gradient) return { background: nextPaper };
    const nextStops = (gradient.stops ?? [
      { color: gradient.from, position: 0 },
      { color: gradient.to, position: 1 },
    ]).map((stop, index) => ({ ...stop, color: map(stop.color, index) }));
    return {
      background: nextPaper,
      backgroundGradient: {
        ...gradient,
        from: nextStops[0].color,
        to: nextStops[nextStops.length - 1].color,
        stops: nextStops,
      },
    };
  };
  let ground: ReturnType<typeof recolor>;
  let groundColor: string;
  let nextAccents = accents;
  if (variantId === "soft") {
    // A dark ground lifts only as far as white words still read on it at 7:1.
    let lift = darkPaper ? 0.12 : 0.55;
    while (darkPaper && lift > 0 && themeLuminance(mixThemeColor(paper, "#ffffff", lift)) > 0.1) lift -= 0.01;
    ground = recolor((color) => {
      if (isBand(color)) return mixThemeColor(color, "#ffffff", 0.25);
      const airier = mixThemeColor(color, "#ffffff", Math.max(0, lift));
      // A ground already white takes a breath of its first accent instead.
      return themeLuminance(airier) > 0.97 ? blendThemeColors(airier, accents[0], 0.05) : airier;
    });
    groundColor = ground.background;
  } else if (variantId === "deep") {
    if (darkPaper) {
      // Tinted by the first accent, so it is the theme's own light, not grey.
      const light = mixThemeColor(blendThemeColors(paper, accents[0], 0.4), "#ffffff", 0.9);
      ground = recolor((color, index) => (isBand(color) ? color : blendThemeColors(light, paper, 0.05 * Math.min(index, 2))));
    } else {
      const dark = mixThemeColor(title, "#000000", 0.35);
      ground = recolor((color, index) => (isBand(color) ? color : mixThemeColor(dark, "#000000", 0.25 * Math.min(index, 2))));
    }
    groundColor = ground.background;
  } else {
    // The strongest dark accent carries the ground; a theme whose accents
    // are all light (a dark theme's) deepens its first.
    const strong =
      accents.find((color) => {
        const luminance = themeLuminance(color);
        return luminance >= 0.03 && luminance <= 0.28;
      }) ?? accents[0];
    // Deep enough that white words read on it at 7:1.
    let ink = strong;
    for (let step = 0; step < 20 && themeLuminance(ink) > 0.1; step += 1) ink = mixThemeColor(ink, "#000000", 0.08);
    ground = {
      background: ink,
      backgroundGradient: {
        from: ink,
        to: mixThemeColor(ink, "#000000", 0.28),
        angle: 135,
        stops: [
          { color: ink, position: 0 },
          { color: mixThemeColor(ink, "#000000", 0.28), position: 1 },
        ],
      },
    };
    groundColor = ink;
    // The ground's own accent cannot fill a shape on it: the others lead,
    // lightened to stand on it, and it closes the set lightened.
    nextAccents = [...accents.filter((color) => color !== strong), strong].map((color) =>
      mixThemeColor(color, "#ffffff", 0.3),
    );
  }
  const lightGround = themeLuminance(groundColor) > 0.4;
  let nextTitle: string;
  let nextBody: string;
  if (variantId === "soft" && !darkPaper) {
    nextTitle = readableThemeColor(title, groundColor, 7);
    nextBody = readableThemeColor(body, groundColor, 4.5);
  } else if (lightGround) {
    // A dark theme's light counterpart: its own ground, deepened, as ink.
    nextTitle = readableThemeColor(mixThemeColor(paper, "#000000", 0.2), groundColor, 7);
    nextBody = readableThemeColor(blendThemeColors(paper, groundColor, 0.35), groundColor, 4.5);
  } else {
    nextTitle = readableThemeColor("#ffffff", groundColor, 7);
    nextBody = readableThemeColor(blendThemeColors(nextTitle, groundColor, 0.25), groundColor, 4.5);
    if (variantId === "deep") nextAccents = accents.map((color) => mixThemeColor(color, "#ffffff", 0.2));
  }
  return {
    ...base,
    backgroundGradient: undefined,
    backgroundPattern: undefined,
    ...ground,
    titleColor: nextTitle,
    bodyTextColor: nextBody,
    accentColors: nextAccents,
  };
}

export function resolveCatalogThemeDesign(
  themeId: string,
  variantId: PresentationThemeVariantId,
): DeckTheme | undefined {
  if (!isThemeDesignVariantId(variantId)) return undefined;
  const preset = PRESET_THEMES.find(
    (entry) => entry.id === themeId && entry.id !== "none",
  );
  if (!preset) return undefined;
  const base = normalizeTheme(preset.theme);
  if (!base) return undefined;
  if (variantId === "base") return base;
  const tokens = normalizeThemeTokenSnapshot(base);
  return tokens
    ? normalizeTheme({ ...base, ...designedThemeVariant(tokens, variantId) })
    : undefined;
}

/** Resolve one variant as a true partial transformation of its selected base. */
export function resolveThemeVariantTokens(
  theme: DeckTheme | null | undefined,
  variantId: PresentationThemeVariantId,
): (DeckThemeTokenSnapshot & Pick<DeckTheme, "colorSlots" | "fontScheme">) | undefined {
  const identity = normalizeTheme(theme)?.identity;
  // An Office theme's variants are data (its own colour schemes), and only
  // its own `v1`..`vN` apply to it.
  if (identity?.selectionKind === "office") {
    return isOfficeThemeVariantId(variantId)
      ? resolveOfficeThemeDesign(identity.themeId, variantId)
      : undefined;
  }
  if (!isThemeDesignVariantId(variantId)) return undefined;
  const base = themeBaseTokens(normalizeTheme(theme));
  if (!base) return undefined;
  if (variantId === "base") return base;
  if (identity?.selectionKind === "catalog") {
    return resolveCatalogThemeDesign(identity.themeId, variantId);
  }
  return designedThemeVariant(base, variantId);
}

/** Latin faces whose closest whitelisted family is the serif stack. */
const SERIF_FACE_RE =
  /garamond|rockwell|georgia|times|cambria|baskerville|bookman|palatino|constantia|book antiqua|century schoolbook|didot|bodoni/i;

/** The four-key family a legacy reader uses for an Office theme's body face. */
function closestThemeFontFamily(face: string): ThemeFontFamily {
  return SERIF_FACE_RE.test(face) ? "serif" : "system";
}

/**
 * One Office theme variant as deck theme fields: the master's ground as
 * `background`, its title and body text colours (through the master's own
 * colour map, so a dark theme is dark), accent1..6, the full twelve-slot
 * scheme, and the theme's fonts.
 */
export function resolveOfficeThemeDesign(
  themeId: string,
  variantId: PresentationOfficeVariantId,
): (DeckTheme & Required<Pick<DeckTheme, "background" | "titleColor" | "bodyTextColor" | "accentColors" | "colorSlots" | "fontScheme" | "fontFamily">>) | undefined {
  const theme = presentationOfficeTheme(themeId);
  const variant = presentationOfficeThemeVariant(themeId, variantId);
  if (!theme || !variant) return undefined;
  const colorSlots = {} as Record<PresentationThemeColorSlot, string>;
  for (const slot of PRESENTATION_OFFICE_COLOR_SLOTS) {
    colorSlots[slot] = variant.colors[slot];
  }
  const fontScheme: DeckThemeFontScheme = {
    majorLatin: theme.fonts.major.latin,
    minorLatin: theme.fonts.minor.latin,
  };
  const majorHans = presentationOfficeThemeHansFace(theme.fonts.major);
  const minorHans = presentationOfficeThemeHansFace(theme.fonts.minor);
  if (majorHans) fontScheme.majorHans = majorHans;
  if (minorHans) fontScheme.minorHans = minorHans;
  return {
    background: variant.background,
    titleColor: variant.titleColor,
    bodyTextColor: variant.bodyColor,
    accentColors: [
      colorSlots.accent1,
      colorSlots.accent2,
      colorSlots.accent3,
      colorSlots.accent4,
      colorSlots.accent5,
      colorSlots.accent6,
    ],
    colorSlots,
    fontScheme,
    fontFamily: closestThemeFontFamily(fontScheme.minorLatin),
  };
}

/**
 * Apply an Office theme (Design > Themes > Office) and one of its variants:
 * colours, colour slots and fonts come from the catalog; the imported
 * deck's `identity.source` (patchback authority) is kept.
 */
export function applyOfficeThemeSelection(
  currentTheme: DeckTheme | null | undefined,
  themeId: string,
  variantId: PresentationOfficeVariantId = "v1",
): DeckTheme | undefined {
  const design = resolveOfficeThemeDesign(themeId, variantId);
  if (!design) return undefined;
  const source = normalizeTheme(currentTheme)?.identity?.source;
  return normalizeTheme({
    ...design,
    identity: {
      schemaVersion: PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
      selectionKind: "office",
      themeId,
      variantId,
      ...(source ? { source } : {}),
    },
  });
}

const COLOR_SCHEME_GROUND_SLOTS: readonly PresentationThemeColorSlot[] = [
  "lt1",
  "dk1",
  "lt2",
  "dk2",
];

/**
 * Design > Variants > Colors (Microsoft Support, "Change the color scheme of
 * a theme"): one Colors-gallery entry replaces the theme's twelve colours and
 * nothing else (PPT-A073). The ground, title and body text keep the slot
 * they were drawn from -- an Office variant's own colour map, found by its
 * colours (a dark variant's ground is a dark slot); any other theme is the
 * page's lt1 / dk1 / dk2, the slots its export writes them to. Null when the
 * id is not a gallery entry or the scheme is already the theme's.
 */
export function applyThemeColorSchemeSelection(
  currentTheme: DeckTheme | null | undefined,
  schemeId: string,
): DeckTheme | null {
  const scheme = officeColorGalleryScheme(schemeId);
  if (!scheme) return null;
  const current = normalizeTheme(currentTheme) ?? {};
  const slots = {} as Record<PresentationThemeColorSlot, string>;
  for (const slot of PRESENTATION_THEME_COLOR_SLOTS) {
    slots[slot] = scheme.colors[slot].toLowerCase();
  }
  const identity = current.identity;
  const variant =
    identity?.selectionKind === "office"
      ? presentationOfficeThemeVariant(identity.themeId, identity.variantId)
      : undefined;
  const slotOf = (
    value: string | undefined,
    fallback: PresentationThemeColorSlot,
  ): PresentationThemeColorSlot => {
    if (!variant || !value) return fallback;
    const lower = value.toLowerCase();
    return (
      COLOR_SCHEME_GROUND_SLOTS.find(
        (slot) => variant.colors[slot].toLowerCase() === lower,
      ) ?? fallback
    );
  };
  const darkGround = variant ? !themeColorIsLight(variant.background) : false;
  // A deck whose own master maps dark (`colorMap`, reading row 826) keeps
  // that master: its ground and words take the slots the map names (the
  // ground the one of Background 1 / 2 it was drawn from), as the import
  // read them.
  const fileMap = variant ? undefined : current.colorMap;
  const fileSlots =
    current.colorSlots ??
    (identity?.selectionKind === "native" ? identity.source?.colorSlots : undefined);
  const fileGround =
    fileMap &&
    current.background &&
    fileSlots?.[fileMap.bg2]?.toLowerCase() === current.background.toLowerCase() &&
    fileSlots?.[fileMap.bg1]?.toLowerCase() !== current.background.toLowerCase()
      ? fileMap.bg2
      : fileMap?.bg1;
  const groundSlot = fileGround ?? slotOf(variant?.background, darkGround ? "dk1" : "lt1");
  const textSlot: PresentationThemeColorSlot = fileMap?.tx1 ?? (darkGround ? "lt1" : "dk1");
  const next: DeckTheme = {
    ...current,
    background: slots[groundSlot],
    titleColor: slots[slotOf(variant?.titleColor, textSlot)],
    bodyTextColor: slots[variant ? slotOf(variant.bodyColor, textSlot) : (fileMap?.tx1 ?? "dk2")],
    accentColors: [
      slots.accent1,
      slots.accent2,
      slots.accent3,
      slots.accent4,
      slots.accent5,
      slots.accent6,
    ],
    colorSlots: slots,
    colorSchemeId: scheme.id,
  };
  const normalized = normalizeTheme(next) ?? null;
  return normalized &&
    current.colorSchemeId === scheme.id &&
    JSON.stringify(current.colorSlots) === JSON.stringify(slots)
    ? null
    : normalized;
}

/**
 * A theme or variant pick resets the colours: the Colors pick goes with it,
 * and so do Slide Master view's edits of the old master (a theme brings its
 * own master, PPT-B030).
 */
export function withoutColorSchemePick(
  theme: DeckTheme | undefined,
): DeckTheme | undefined {
  if (
    !theme ||
    (theme.colorSchemeId === undefined &&
      theme.effectSchemeId === undefined &&
      theme.officeMaster === undefined)
  ) {
    return theme;
  }
  const next = { ...theme };
  delete next.colorSchemeId;
  delete next.effectSchemeId;
  // The masters Insert Slide Master added are their own themes and stay,
  // with the slides on them; only master one takes the new theme (PPT-B030).
  const additionalMasters = normalizeDeckOfficeMaster(theme.officeMaster)?.additionalMasters;
  if (additionalMasters) next.officeMaster = { additionalMasters };
  else delete next.officeMaster;
  return next;
}

/**
 * The theme a NEW deck starts with. One line chooses it: the catalog's
 * default Office theme ("office"), or set it to an MRAI preset id such as
 * "classic" to start new decks on that preset again.
 */
export const NEW_PRESENTATION_THEME_ID: string =
  PRESENTATION_OFFICE_THEME_DEFAULT_ID;

export function defaultPresentationTheme(): DeckTheme | undefined {
  return presentationOfficeTheme(NEW_PRESENTATION_THEME_ID)
    ? applyOfficeThemeSelection(undefined, NEW_PRESENTATION_THEME_ID, "v1")
    : applyPresetThemeSelection(undefined, NEW_PRESENTATION_THEME_ID);
}

/** What a NEW shape and a NEW line wear in a deck on one theme. */
export interface PresentationThemeObjectStyles {
  readonly shape: OfficeDefaultShapeStyle & { readonly textColor: string };
  readonly line: OfficeDefaultShapeStyle;
}

/**
 * PowerPoint's own object styles over this deck's theme (PPT-H010). A new
 * shape is `a:fillRef idx="1"` accent1, `a:lnRef idx="2"` accent1 with
 * `a:shade 15000` and `a:fontRef minor` lt1: a theme-coloured shape with a
 * darker outline and white text. A new line is the theme's `a:lnDef`,
 * `a:lnRef idx="2"` accent1. The weight is the theme's second line style
 * (1.5 pt in the Office theme); a theme without Office line styles keeps the
 * editor's 2 px, which is the same 1.5 pt. A deck with no theme of its own
 * gets `null` and keeps the editor's historical default.
 */
export function presentationThemeDefaultShapeStyle(
  theme: DeckTheme | null | undefined,
): PresentationThemeObjectStyles | null {
  const normalized = normalizeTheme(theme);
  const accent1 = (
    normalized?.colorSlots?.accent1 ?? normalized?.accentColors?.[0]
  )?.toLowerCase();
  if (!normalized || !accent1 || !HEX_COLOR_RE.test(accent1)) return null;
  const outline = officeTransformedHex(accent1, [["shade", 15000]])?.toLowerCase();
  const identity = normalized.identity;
  const lineEmu =
    identity?.selectionKind === "office"
      ? PRESENTATION_OFFICE_THEME_SHAPE_LINE_EMU[identity.themeId]
      : undefined;
  const weight = lineEmu ? { outlineWidthPt: lineEmu / 12700 } : {};
  const lt1 = normalized.colorSlots?.lt1?.toLowerCase();
  return {
    shape: {
      fill: accent1,
      ...(outline ? { outline } : {}),
      ...weight,
      textColor: lt1 && HEX_COLOR_RE.test(lt1) ? lt1 : "#ffffff",
    },
    line: { outline: accent1, ...weight },
  };
}

/** One tile of the Variants group. */
export interface PresentationThemeVariantOption {
  id: PresentationThemeVariantId;
  /** i18n key of a computed variant's name (MRAI kinds only). */
  labelKey?: (typeof PRESENTATION_THEME_VARIANTS)[number]["labelKey"];
  /** 1-based position of an Office variant (its tiles are numbered). */
  officeIndex?: number;
  theme: DeckTheme;
}

/**
 * The Variants group for a deck theme: an Office theme's own variants (4;
 * Integral 8), else the four computed variants of the selection (a themeless
 * deck counts as Classic, as it always has).
 */
export function availableThemeVariants(
  theme: DeckTheme | null | undefined,
): PresentationThemeVariantOption[] {
  const normalized = normalizeTheme(theme);
  const identity = normalized?.identity;
  if (identity?.selectionKind === "office") {
    const office = presentationOfficeTheme(identity.themeId);
    if (!office) return [];
    return office.variants.flatMap((variant) => {
      const design = resolveOfficeThemeDesign(office.id, variant.id);
      return design
        ? [{ id: variant.id, officeIndex: variant.index, theme: design }]
        : [];
    });
  }
  const source = normalized ?? applyPresetThemeSelection(undefined, "classic");
  return PRESENTATION_THEME_VARIANTS.flatMap((variant) => {
    const variantTheme = resolveThemeVariantTokens(source, variant.id);
    return variantTheme
      ? [{ id: variant.id, labelKey: variant.labelKey, theme: variantTheme }]
      : [];
  });
}

/** Apply a catalog identity while retaining immutable native patchback custody. */
export function applyPresetThemeSelection(
  currentTheme: DeckTheme | null | undefined,
  presetId: string,
): DeckTheme | undefined {
  const preset = PRESET_THEMES.find((entry) => entry.id === presetId);
  if (!preset || preset.id === "none") return undefined;
  const source = normalizeTheme(currentTheme)?.identity?.source;
  return normalizeTheme({
    ...resolveCatalogThemeDesign(preset.id, "base"),
    identity: {
      schemaVersion: PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
      selectionKind: "catalog",
      themeId: preset.id,
      variantId: "base",
      ...(source ? { source } : {}),
    },
  });
}

/** Apply one available variant without touching slide/block local overrides. */
export function applyThemeVariantSelection(
  currentTheme: DeckTheme | null | undefined,
  variantId: PresentationThemeVariantId,
): DeckTheme | undefined {
  if (!THEME_VARIANT_ID_SET.has(variantId)) return normalizeTheme(currentTheme);
  const current = normalizeTheme(currentTheme);
  const selection = resolveThemeSelection(current);
  if (!selection || selection.selectionKind === "custom") return current;
  if (current?.identity?.selectionKind === "office") {
    // Office variant: the variant's colour scheme replaces the colours; the
    // fonts (and every other field) stay as they are.
    const variant = isOfficeThemeVariantId(variantId)
      ? resolveOfficeThemeDesign(current.identity.themeId, variantId)
      : undefined;
    if (!variant) return current;
    // The variant's own master comes in with its colour map: the map the
    // file's master stated (`colorMap`) goes.
    const rest: DeckTheme = { ...current };
    delete rest.colorMap;
    return normalizeTheme({
      ...rest,
      background: variant.background,
      titleColor: variant.titleColor,
      bodyTextColor: variant.bodyTextColor,
      accentColors: variant.accentColors,
      colorSlots: variant.colorSlots,
      identity: { ...current.identity, variantId },
    });
  }
  // The MRAI kinds only know their four computed variants.
  if (!isThemeDesignVariantId(variantId)) return current;
  const legacyPreset = current?.identity ? undefined : presetForLegacyTheme(current);
  const identity: DeckThemeIdentity = current?.identity ?? {
    schemaVersion: PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
    selectionKind: "catalog",
    themeId: legacyPreset?.id ?? selection.themeId,
    variantId: "base",
  };
  if (identity.selectionKind === "catalog") {
    const design = resolveCatalogThemeDesign(identity.themeId, variantId);
    return design
      ? normalizeTheme({
          ...design,
          identity: { ...identity, variantId },
        })
      : current;
  }
  const variant = resolveThemeVariantTokens(
    { ...(current ?? {}), identity },
    variantId,
  );
  if (!variant) return current;
  const next: DeckTheme = { ...(current ?? {}) };
  for (const field of THEME_VARIANT_FIELDS) delete next[field];
  Object.assign(next, variant);
  next.identity = { ...identity, variantId };
  return normalizeTheme(next);
}

/** Rebuild a catalog theme's rich design after native PPTX reopen. */
export function materializeCatalogThemeIdentity(
  theme: DeckTheme | null | undefined,
): DeckTheme | undefined {
  const normalized = normalizeTheme(theme);
  const identity = normalized?.identity;
  if (identity?.selectionKind === "office") {
    // A receipt that arrives without its scheme (e.g. identified at import)
    // gains the catalog's colours and fonts; any field it states wins.
    const design = isOfficeThemeVariantId(identity.variantId)
      ? resolveOfficeThemeDesign(identity.themeId, identity.variantId)
      : undefined;
    return design ? normalizeTheme({ ...design, ...normalized }) : normalized;
  }
  if (identity?.selectionKind !== "catalog") return normalized;
  const design = resolveCatalogThemeDesign(
    identity.themeId,
    identity.variantId,
  );
  if (!design) return normalized;
  // The file is the authority on what the deck wears; the catalog identity
  // names the gallery tile, and fills only what the reading does not state.
  // Rebuilt whole, the preset's values replaced the file's on every reopen
  
  // a Master Font's faces and a master background reverted to the preset's).
  // A gallery pick (`fontSchemeId` ...) is an edit's receipt, never a reading;
  // a background or font the reading states is one whole choice.
  const filled: DeckTheme = { ...design };
  for (const key of CATALOG_PICK_KEYS) delete filled[key];
  if (CATALOG_BACKGROUND_FIELDS.some((key) => normalized![key] !== undefined)) {
    for (const key of CATALOG_BACKGROUND_FIELDS) delete filled[key];
  }
  if (normalized!.fontScheme !== undefined || normalized!.fontFamily !== undefined) {
    delete filled.fontScheme;
    delete filled.fontFamily;
  }
  return normalizeTheme({ ...filled, ...normalized });
}

/** Gallery picks a theme states only once a Design command made them. */
const CATALOG_PICK_KEYS = ["fontSchemeId", "colorSchemeId", "effectSchemeId"] as const satisfies readonly (keyof DeckTheme)[];
/** One background choice: a reading that states any of them states the background. */
const CATALOG_BACKGROUND_FIELDS = [
  "background",
  "backgroundGradient",
  "backgroundPattern",
  "backgroundImage",
] as const satisfies readonly (keyof DeckTheme)[];

/** Mark direct field editing as custom while preserving native source authority. */
export function markThemeSelectionCustom(
  theme: DeckTheme | null | undefined,
): DeckTheme | undefined {
  const normalized = normalizeTheme(theme);
  if (!normalized) return undefined;
  const prior = normalized.identity;
  normalized.identity = {
    schemaVersion: PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
    selectionKind: "custom",
    themeId: prior?.themeId ?? "custom",
    variantId: "base",
    ...(prior?.source ? { source: prior.source } : {}),
  };
  return normalizeTheme(normalized);
}

/** No-op guard used at the command boundary before history/autosave mutation. */
export function deckThemesEqual(
  left: DeckTheme | null | undefined,
  right: DeckTheme | null | undefined,
): boolean {
  return (
    JSON.stringify(normalizeTheme(left) ?? null) ===
    JSON.stringify(normalizeTheme(right) ?? null)
  );
}

/**
 * The CSS stack of the theme's picked Fonts-gallery entry (Design > Fonts):
 * its heading face for title text, its body face for everything else, as
 * PowerPoint's `+mj-lt` / `+mn-lt` resolve. `undefined` without a pick.
 */
export function themeFontSchemeFamily(
  theme: DeckTheme | null | undefined,
  heading: boolean,
): string | undefined {
  // The theme's own font scheme (an Office theme) comes first: its Latin
  // face, then its Simplified Chinese face, so CJK text takes the theme's
  // Hans face as PowerPoint's `script="Hans"` entry resolves it.
  const scheme = normalizeThemeFontScheme(theme?.fontScheme);
  if (scheme) {
    return officeThemeFontStack({
      latin: heading ? scheme.majorLatin : scheme.minorLatin,
      ea: "",
      scripts: {
        Hans: (heading ? scheme.majorHans : scheme.minorHans) ?? "",
      },
    });
  }
  const entry = officeThemeFontGalleryEntry(theme?.fontSchemeId);
  if (!entry) return undefined;
  return heading ? entry.headingFamily : entry.bodyFamily;
}

/**
 * Resolve the effective background colour for a slide, applying the precedence
 *   slide.backgroundColor (override)  >  theme.background  >  built-in default.
 *
 * The slide override is re-validated here (not trusted verbatim) so a stale /
 * hand-edited `backgroundColor` can never inject arbitrary CSS. Returns
 * `undefined` for the built-in default (no explicit background) so the renderer
 * omits the `background-color` entirely, exactly as a themeless deck did before.
 */
export function resolveSlideBackground(
  slide: { backgroundColor?: unknown } | null | undefined,
  theme: DeckTheme | null | undefined,
): string | undefined {
  const override = normalizeThemeColor(slide?.backgroundColor);
  if (override) return override;
  const themed = normalizeThemeColor(theme?.background);
  if (themed) return themed;
  return undefined;
}

/**
 * The resolved theme contribution for a text block: the colour + font-family
 * that should apply WHEN the block has no explicit override. Both fields are
 * optional — `undefined` means "no theme contribution, fall through to the
 * built-in default". `fontFamilyKey` is the whitelisted key; `fontFamily` is the
 * concrete CSS stack (so callers can apply it directly without re-mapping).
 */
export interface ResolvedBlockTextStyle {
  color?: string;
  fontFamilyKey?: ThemeFontFamily;
  fontFamily?: string;
  /**
   * Theme-derived default font size (px) — present only when the theme supplies
   * one AND the block has no explicit `style.fontSize`. `undefined` means "no
   * theme size contribution; keep the built-in default size".
   */
  fontSize?: number;
  /**
   * Theme-derived default bold flag — present (always `true`) only when the theme
   * asks for bold AND the block has no explicit `style.bold`. `undefined` means
   * "no theme weight contribution; keep the built-in weight".
   */
  bold?: boolean;
  /** Theme-derived default italic, as `bold` (only `true`, only when the block states none). */
  italic?: boolean;
}

/**
 * Resolve the theme-derived text style for a block, applying the precedence
 *   block.style.color / .fontFamily (override)  >  theme  >  built-in default.
 *
 * Title-class blocks (placeholder `title` / `sectionTitle`) draw from
 * `theme.titleColor`; every other text block draws from `theme.bodyTextColor`.
 * When the block already specifies its own (valid) colour / font, that override
 * wins and the corresponding theme field is omitted so the caller keeps the
 * explicit per-block value. Non-text blocks (image / chart / table / shape) have
 * no text style and yield an empty result.
 */
export function resolveBlockTextStyle(
  block:
    | {
        type?: unknown;
        placeholder?: unknown;
        placeholderType?: unknown;
        source?: unknown;
        smartArtKind?: unknown;
        style?: {
          color?: unknown;
          fontFamily?: unknown;
          fontSize?: unknown;
          bold?: unknown;
          italic?: unknown;
        } | null;
      }
    | null
    | undefined,
  theme: DeckTheme | null | undefined,
): ResolvedBlockTextStyle {
  if (!block || block.type !== "text") return {};
  const isTitle = blockIsTitle(block);
  const result: ResolvedBlockTextStyle = {};

  // Colour: an explicit (valid) per-block colour overrides the theme. A title
  // block draws from theme.titleColor; everything else from theme.bodyTextColor.
  const overrideColor = normalizeThemeColor(block.style?.color);
  if (!overrideColor) {
    const themed = isTitle
      ? normalizeThemeColor(theme?.titleColor)
      : normalizeThemeColor(theme?.bodyTextColor);
    if (themed) result.color = themed;
  }

  // Font family: an explicit (whitelisted) per-block font overrides the theme.
  const overrideFont = normalizeThemeFontFamily(block.style?.fontFamily);
  if (!overrideFont) {
    const themedFont = normalizeThemeFontFamily(theme?.fontFamily);
    if (themedFont) result.fontFamilyKey = themedFont;
    const schemeFamily = themeFontSchemeFamily(theme, isTitle);
    if (schemeFamily) result.fontFamily = schemeFamily;
    else if (themedFont) result.fontFamily = THEME_FONT_FAMILY_CSS[themedFont];
  }

  // Font size: an explicit per-block size (any finite number) overrides the
  // theme. A title block draws from theme.titleFontSize; everything else from
  // theme.bodyFontSize. The theme value is re-clamped so a stale size is safe.
  const overrideSize = block.style?.fontSize;
  const hasOverrideSize =
    typeof overrideSize === "number" && Number.isFinite(overrideSize);
  if (!hasOverrideSize) {
    const themedSize = normalizeThemeFontSize(
      isTitle ? theme?.titleFontSize : theme?.bodyFontSize,
    );
    if (themedSize !== undefined) result.fontSize = themedSize;
  }

  // Bold: an explicit per-block bold flag (a real boolean either way) overrides
  // the theme. A title block draws from theme.titleBold; everything else from
  // theme.bodyBold. Only a `true` theme flag contributes (mirrors normalizeTheme).
  // An imported text box (no placeholder) is not the master's body: PowerPoint
  // styles it from the master's other-text style, never its body style. The
  // body's bold (theme.bodyBold, read from `p:bodyStyle`) made every such box
  
  // italic drawn bold; both masters' level-one body is b="1").
  // Nor is a SmartArt's words: Office draws a diagram's text from the
  
  // and their reasons, regular Arial, drew bold on a master whose level-one
  // body is b="1").
  const importedTextBox =
    !isTitle &&
    block.placeholder === undefined &&
    block.placeholderType === undefined &&
    ((typeof block.source === "object" &&
      block.source !== null &&
      typeof (block.source as { slideId?: unknown }).slideId === "string") ||
      typeof block.smartArtKind === "string");
  if (typeof block.style?.bold !== "boolean" && !importedTextBox) {
    const themedBold = isTitle ? theme?.titleBold : theme?.bodyBold;
    if (themedBold === true) result.bold = true;
  }
  if (typeof block.style?.italic !== "boolean" && !importedTextBox) {
    const themedItalic = isTitle ? theme?.titleItalic : theme?.bodyItalic;
    if (themedItalic === true) result.italic = true;
  }

  return result;
}

/**
 * Resolve the theme-derived style for the SLIDE TITLE element (the standalone
 * `<h2>` / title input, not a text block). Draws colour from `theme.titleColor`
 * and font from `theme.fontFamily`. Returns an empty result for a missing theme
 * so the title renders with its built-in styling, exactly as before.
 */
export function resolveTitleTextStyle(
  theme: DeckTheme | null | undefined,
): ResolvedBlockTextStyle {
  const result: ResolvedBlockTextStyle = {};
  const color = normalizeThemeColor(theme?.titleColor);
  if (color) result.color = color;
  const font = normalizeThemeFontFamily(theme?.fontFamily);
  if (font) result.fontFamilyKey = font;
  const schemeFamily = themeFontSchemeFamily(theme, true);
  if (schemeFamily) result.fontFamily = schemeFamily;
  else if (font) result.fontFamily = THEME_FONT_FAMILY_CSS[font];
  const size = normalizeThemeFontSize(theme?.titleFontSize);
  if (size !== undefined) result.fontSize = size;
  if (theme?.titleBold === true) result.bold = true;
  if (theme?.titleItalic === true) result.italic = true;
  return result;
}

/**
 * The slide surface's hyperlink colour: the theme's Hyperlink slot (`hlink`),
 * stated as the `--presentation-hyperlink` custom property every run link on
 * the slide draws in (`runStyleCss`). Empty when the theme states no slot.
 */
export function presentationHyperlinkStyle(
  theme: DeckTheme | null | undefined,
  





  masterLinkColors?: { hlink?: string } | null,
): Record<string, string> {
  const own = masterLinkColors?.hlink;
  if (typeof own === "string" && /^#[0-9a-f]{6}$/i.test(own)) {
    return { "--presentation-hyperlink": own };
  }
  // An imported deck's own theme (not one Office/MRAI gallery theme) states
  // its twelve slots on the import's theme source only: its Hyperlink slot
  
  // Office 2007 theme drew black, PowerPoint draws the theme's 0000FF).
  const identity = theme?.identity;
  const nativeHlink =
    identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source?.colorSlots?.hlink
      : undefined;
  const hlink = theme?.colorSlots?.hlink ?? nativeHlink;
  return typeof hlink === "string" && /^#[0-9a-f]{6}$/i.test(hlink)
    ? { "--presentation-hyperlink": hlink }
    : {};
}











export function presentationBodyFontStyle(
  theme: DeckTheme | null | undefined,
): Record<string, string> {
  const scheme = themeFontSchemeFamily(theme, false);
  if (scheme) return { "--presentation-body-font": scheme };
  const identity = theme?.identity;
  const native =
    identity?.selectionKind === "native" && (identity.variantId ?? "base") === "base"
      ? identity.source?.minorLatinFont
      : undefined;
  return native
    ? { "--presentation-body-font": officeThemeFontStack({ latin: native, ea: "", scripts: {} }) }
    : {};
}

/**
 * Resolve the theme's accent palette into a clean list of quick-pick swatch
 * colours, one per accent slot in order (callers read `[1]` as Accent 2),
 * re-validating each entry so a stale theme is render-safe.
 * Returns `[]` (never `undefined`) so callers can map over it unconditionally:
 * an absent / empty palette simply yields no extra swatches and the picker keeps
 * exactly its built-in colours, unchanged.
 */
export function resolveThemeAccentColors(
  theme: DeckTheme | null | undefined,
): string[] {
  return normalizeThemeAccentColors(theme?.accentColors) ?? [];
}
