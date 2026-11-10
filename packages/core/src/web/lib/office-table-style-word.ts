/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Office table gallery (PowerPoint's 74 built-in table styles, the same
 * slots the Presentation editor offers) as WordprocessingML table styles.
 *
 * `office-table-style-word-catalog.json` is generated once from the
 * Presentation catalogue by
 * `apps/api/scripts/generate_office_table_style_word_catalog.py`, which
 * documents the part-by-part mapping; the API holds a byte-identical twin.
 * This module only resolves and serializes it:
 *
 * - a style is a `w:style w:type="table"` whose own `w:rPr`/`w:tblPr`/
 *   `w:tcPr` carry the whole-table look and whose `w:tblStylePr` records carry
 *   the conditional regions (ECMA-376 Part 1 §17.7.6);
 * - every colour is a theme reference (`w:themeColor`/`w:themeFill` with
 *   `w:themeTint`/`w:themeShade`), resolved with Word's own rule: the tint or
 *   shade moves HSL luminance in the byte domain, hue and saturation stay
 *   (the same rule the API's table-style resolver applies on import);
 * - a style is identified by its NAME, as Word identifies styles. Its id is
 *   the name without spaces and punctuation ("Medium Style 2 - Accent 1" ->
 *   `MediumStyle2-Accent1`), suffixed `-2`, `-3`… only when a document
 *   already uses that id for a style of another name: a document's own
 *   style is never overwritten, and a document that already holds the
 *   gallery style keeps its own definition.
 *
 * The styles' looks are served by the API to a signed-in page
 * (`lib/office-data/store.ts`), not bundled: until they arrive, and on a
 * guest page, every gallery style resolves to MRAI's own recipe of its
 * family (`office-table-style-fallback.ts`, the same names and ids, so a
 * table styled from it is written as the style it names), while the
 * gallery's names and slots (`office-table-style-index.ts`) and the default
 * theme are always here.
 */
import { officeData } from "./office-data/store";
import {
  fallbackColorHex,
  fallbackTableStyles,
  type FallbackColor,
  type FallbackLine,
  type FallbackPart,
  type FallbackPartName,
  type FallbackThemeSlot,
} from "./office-table-style-fallback";
import { WORD_TABLE_STYLE_DEFAULT_THEME } from "./office-table-style-index";
import {
  normalizeOfficeThemePaletteId,
  paletteById,
} from "./office-color-palette";
import { hlsToRgb, rgbToHls } from "./office-color-transform";
import { documentDesignWordThemeColors, normalizeDocumentDesign } from "./document-design";
import type {
  DocumentTableLook,
  DocumentTableStyleVisualAuthority,
  DocumentTableStyleVisualBorder,
  DocumentTableStyleVisualLayer,
  DocumentTableVisualCondition,
  DocumentTableVisualEdge,
} from "./document-table-design-authority";

export const OFFICE_WORD_THEME_COLOR_NAMES = [
  "text1",
  "background1",
  "text2",
  "background2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;
export type OfficeWordThemeColorName =
  (typeof OFFICE_WORD_THEME_COLOR_NAMES)[number];
export type OfficeWordThemeColors = Readonly<
  Record<OfficeWordThemeColorName, string>
>;

export interface OfficeWordThemeColorRef {
  readonly theme: OfficeWordThemeColorName;
  readonly tint?: string;
  readonly shade?: string;
}

export interface OfficeWordTableStyleBorder {
  readonly style: "single" | "double" | "nil";
  readonly size?: number;
  readonly color?: OfficeWordThemeColorRef;
}

type BorderEdge = "top" | "left" | "bottom" | "right" | "insideH" | "insideV";

export interface OfficeWordTableStyleLayer {
  readonly run?: { readonly bold?: boolean; readonly color?: OfficeWordThemeColorRef };
  readonly borders?: Partial<Record<BorderEdge, OfficeWordTableStyleBorder>>;
  readonly fill?: OfficeWordThemeColorRef;
}

export type OfficeWordTableStyleGroup = "best" | "light" | "medium" | "dark";

export interface OfficeWordTableStyle {
  readonly slot: number;
  readonly group: OfficeWordTableStyleGroup;
  /** The PowerPoint GUID this style reproduces. */
  readonly presentationStyleId: string;
  /** The preferred `w:styleId`; see {@link officeWordTableStyleIdFor}. */
  readonly styleId: string;
  /** `w:name`: the style's identity in Word and on import. */
  readonly name: string;
  readonly table: OfficeWordTableStyleLayer;
  readonly conditions: Partial<
    Record<DocumentTableVisualCondition, OfficeWordTableStyleLayer>
  >;
}

interface OfficeWordTableStyleTables {
  readonly styles: readonly OfficeWordTableStyle[];
  readonly bySlot: ReadonlyMap<number, OfficeWordTableStyle>;
  readonly byName: ReadonlyMap<string, OfficeWordTableStyle>;
  readonly byId: ReadonlyMap<string, OfficeWordTableStyle>;
}

let servedTables: { readonly source: unknown; readonly tables: OfficeWordTableStyleTables } | null = null;

/** The served catalogue, indexed once per arrival; the recipe's styles
 * until it is here. */
function wordTableStyleTables(): OfficeWordTableStyleTables {
  const source = officeData("word-table-styles") as { styles?: unknown } | undefined;
  if (!source || !Array.isArray(source.styles)) return fallbackWordTableStyleTables();
  if (servedTables?.source !== source) {
    const styles = source.styles as readonly OfficeWordTableStyle[];
    servedTables = {
      source,
      tables: {
        styles,
        bySlot: new Map(styles.map((style) => [style.slot, style])),
        byName: new Map(styles.map((style) => [style.name, style])),
        byId: new Map(styles.map((style) => [style.styleId, style])),
      },
    };
  }
  return servedTables.tables;
}

/** The 74 gallery styles with their looks: the served ones, or the recipe's
 * until they have arrived. */
export function officeWordTableStyles(): readonly OfficeWordTableStyle[] {
  return wordTableStyleTables().styles;
}

/** Whether the gallery is drawing the recipe (the served looks are away). */
export function officeWordTableStylesAreFallback(): boolean {
  return wordTableStyleTables() === fallbackTables;
}

const FALLBACK_SLOT_THEME: Readonly<Record<FallbackThemeSlot, OfficeWordThemeColorName>> = {
  dk1: "text1",
  lt1: "background1",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
};
/** A Presentation line width (pt) as Word's eighths of a point, as the
 * generated twin maps the gallery's widths. */
const FALLBACK_BORDER_SIZE: Readonly<Record<string, number>> = { "1.5": 8, "2": 12, "2.5": 16, "4": 24 };
const FALLBACK_CONDITIONS: readonly (readonly [FallbackPartName, DocumentTableVisualCondition])[] = [
  ["firstRow", "firstRow"],
  ["lastRow", "lastRow"],
  ["firstCol", "firstCol"],
  ["lastCol", "lastCol"],
  ["band1V", "band1Vert"],
  ["band2V", "band2Vert"],
  ["band1H", "band1Horz"],
  ["band2H", "band2Horz"],
  ["neCell", "neCell"],
  ["nwCell", "nwCell"],
  ["seCell", "seCell"],
  ["swCell", "swCell"],
];

/** What a cell shows beneath a part: the table's fill, its background, or
 * the white page. */
interface FallbackUnder {
  readonly hex: string;
  readonly theme: OfficeWordThemeColorName;
}

function hexChannels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16) / 255,
    parseInt(hex.slice(3, 5), 16) / 255,
    parseInt(hex.slice(5, 7), 16) / 255,
  ];
}

/** The theme reference whose Word tint or shade draws closest to `target`
 * (Word moves HSL luminance; the recipe mixes in linear light, as the
 * generated twin's bytes were fitted). */
function fittedWordThemeRef(
  theme: OfficeWordThemeColorName,
  target: string,
): OfficeWordThemeColorRef {
  const base = OFFICE_WORD_DEFAULT_THEME_COLORS[theme].toLowerCase();
  const wanted = target.toLowerCase();
  if (base === wanted) return { theme };
  const [, baseLuminance] = rgbToHls(...hexChannels(base));
  const luminance = Math.round(baseLuminance * 255) / 255;
  const [, targetLuminance] = rgbToHls(...hexChannels(wanted));
  const tint = targetLuminance > luminance;
  const guess = tint
    ? luminance >= 1
      ? 255
      : Math.round(255 * (1 - (targetLuminance - luminance) / (1 - luminance)))
    : luminance <= 0
      ? 255
      : Math.round((255 * targetLuminance) / luminance);
  const want = hexChannels(wanted);
  let best: OfficeWordThemeColorRef = { theme };
  let bestDistance = Number.POSITIVE_INFINITY;
  for (let byte = Math.max(0, guess - 3); byte <= Math.min(255, guess + 3); byte += 1) {
    const value = byte.toString(16).toUpperCase().padStart(2, "0");
    const ref: OfficeWordThemeColorRef = tint ? { theme, tint: value } : { theme, shade: value };
    const drawn = hexChannels(officeWordThemeColorHex(ref));
    const distance = drawn.reduce((sum, channel, i) => sum + (channel - want[i]) ** 2, 0);
    if (distance < bestDistance) {
      best = ref;
      bestDistance = distance;
    }
  }
  return best;
}

function fallbackThemeRef(
  color: FallbackColor,
  under: FallbackUnder | null,
): OfficeWordThemeColorRef {
  // White laid over a colour is that colour, lightened.
  const theme =
    color.alpha !== undefined && color.slot === "lt1" && under
      ? under.theme
      : FALLBACK_SLOT_THEME[color.slot];
  return fittedWordThemeRef(theme, fallbackColorHex(color, under?.hex ?? null));
}

function fallbackBorder(line: FallbackLine): OfficeWordTableStyleBorder {
  if (line.style === "none" || !line.color) return { style: "nil" };
  if (line.style === "double") return { style: "double", size: 6, color: fallbackThemeRef(line.color, null) };
  return {
    style: "single",
    size: FALLBACK_BORDER_SIZE[String(line.width)] ?? Math.round(line.width * 8),
    color: fallbackThemeRef(line.color, null),
  };
}

function fallbackLayer(
  part: FallbackPart,
  under: FallbackUnder | null,
  tableBackground?: FallbackColor,
): OfficeWordTableStyleLayer {
  const fill =
    part.fill && part.fill !== "none"
      ? fallbackThemeRef(part.fill, under)
      : tableBackground
        ? fallbackThemeRef(tableBackground, null)
        : undefined;
  const run =
    part.bold || part.color
      ? {
          ...(part.bold ? { bold: true } : {}),
          ...(part.color ? { color: fallbackThemeRef(part.color, null) } : {}),
        }
      : undefined;
  const borders = part.borders
    ? (Object.fromEntries(
        Object.entries(part.borders).map(([edge, line]) => [edge, fallbackBorder(line)]),
      ) as Partial<Record<BorderEdge, OfficeWordTableStyleBorder>>)
    : undefined;
  return {
    ...(run ? { run } : {}),
    ...(borders ? { borders } : {}),
    ...(fill ? { fill } : {}),
  };
}

let fallbackTables: OfficeWordTableStyleTables | null = null;

/** The recipe as Word table styles, built once. */
function fallbackWordTableStyleTables(): OfficeWordTableStyleTables {
  if (fallbackTables) return fallbackTables;
  const styles = fallbackTableStyles().map((style): OfficeWordTableStyle => {
    const whole = style.parts.wholeTbl ?? {};
    const beneath = whole.fill && whole.fill !== "none" ? whole.fill : (style.background ?? null);
    const under: FallbackUnder = beneath
      ? { hex: fallbackColorHex(beneath), theme: FALLBACK_SLOT_THEME[beneath.slot] }
      : { hex: "#ffffff", theme: "background1" };
    const conditions: Partial<Record<DocumentTableVisualCondition, OfficeWordTableStyleLayer>> = {};
    for (const [part, condition] of FALLBACK_CONDITIONS) {
      const value = style.parts[part];
      if (value) conditions[condition] = fallbackLayer(value, under);
    }
    return {
      slot: style.slot,
      group: style.group,
      presentationStyleId: style.presentationStyleId,
      styleId: style.wordStyleId,
      name: style.name,
      table: fallbackLayer(whole, under, style.background),
      conditions,
    };
  });
  fallbackTables = {
    styles,
    bySlot: new Map(styles.map((style) => [style.slot, style])),
    byName: new Map(styles.map((style) => [style.name, style])),
    byId: new Map(styles.map((style) => [style.styleId, style])),
  };
  return fallbackTables;
}

/**
 * The current Office theme: the one the Presentation gallery is drawn in and
 * a new Word document gets. New Docs pages render their gallery tables in it
 * and both DOCX writers put it in the package (Word resolves every
 * `w:themeFill`/`w:themeColor` against the package's theme, never the cached
 * hex beside it); an imported document keeps its own theme.
 */
export const OFFICE_WORD_DEFAULT_THEME_COLORS: OfficeWordThemeColors =
  WORD_TABLE_STYLE_DEFAULT_THEME;

/** Word's own conditional-format write order (and our XML order). */
const CONDITION_ORDER: readonly DocumentTableVisualCondition[] = [
  "firstRow",
  "lastRow",
  "firstCol",
  "lastCol",
  "band1Vert",
  "band2Vert",
  "band1Horz",
  "band2Horz",
  "neCell",
  "nwCell",
  "seCell",
  "swCell",
];
const EDGE_ORDER: readonly BorderEdge[] = [
  "top",
  "left",
  "bottom",
  "right",
  "insideH",
  "insideV",
];
const HEX = /^#[0-9a-f]{6}$/;
const BYTE = /^[0-9A-F]{2}$/;

export function officeWordTableStyleBySlot(
  slot: number,
): OfficeWordTableStyle | null {
  return wordTableStyleTables().bySlot.get(slot) ?? null;
}

export function officeWordTableStyleByName(
  name: unknown,
): OfficeWordTableStyle | null {
  return typeof name === "string" ? (wordTableStyleTables().byName.get(name) ?? null) : null;
}

/** The gallery style an id was allocated for (`Id` or `Id-2`, `Id-3`…). */
export function officeWordTableStyleForAllocatedId(
  styleId: unknown,
): OfficeWordTableStyle | null {
  if (typeof styleId !== "string") return null;
  const byId = wordTableStyleTables().byId;
  const exact = byId.get(styleId);
  if (exact) return exact;
  const suffixed = /^(.+)-([2-9]|[1-9][0-9]{1,2})$/.exec(styleId);
  return suffixed ? (byId.get(suffixed[1]) ?? null) : null;
}

/** A gallery colour column: the neutral style, or accent 1-6. */
export type OfficeWordTableStyleColor = "neutral" | 1 | 2 | 3 | 4 | 5 | 6;
export const OFFICE_WORD_TABLE_STYLE_COLORS: readonly OfficeWordTableStyleColor[] =
  ["neutral", 1, 2, 3, 4, 5, 6];

/** "Medium Style 2 - Accent 1" -> family "Medium Style 2", accents [1]. */
function styleFamily(style: OfficeWordTableStyle): {
  readonly family: string;
  readonly accents: readonly number[];
} {
  const match = /^(.*) - Accent ([1-6])(?:\/Accent ([1-6]))?$/.exec(style.name);
  if (!match) return { family: style.name, accents: [] };
  return {
    family: match[1],
    accents: [Number(match[2]), ...(match[3] ? [Number(match[3])] : [])],
  };
}

/** The colour column a gallery style sits in. */
export function officeWordTableStyleColor(
  style: OfficeWordTableStyle,
): OfficeWordTableStyleColor {
  const { accents } = styleFamily(style);
  return accents.length
    ? (accents[0] as OfficeWordTableStyleColor)
    : "neutral";
}

/**
 * The same style in another colour column, as the Presentation gallery's
 * columns are: "Medium Style 2 - Accent 1" in accent 3 is "Medium Style 2 -
 * Accent 3"; neutral drops the accent. The family (Light/Medium/Dark and its
 * number) never changes, so a family without that colour (Themed Style has
 * no neutral, the No Style pair no accents) answers `null`. Dark Style 2's
 * two-accent styles answer for either of their accents.
 */
export function officeWordTableStyleRecolor(
  style: OfficeWordTableStyle,
  color: OfficeWordTableStyleColor,
): OfficeWordTableStyle | null {
  const { family } = styleFamily(style);
  return (
    officeWordTableStyles().find((candidate) => {
      const other = styleFamily(candidate);
      if (other.family !== family) return false;
      return color === "neutral"
        ? other.accents.length === 0
        : other.accents.includes(color);
    }) ?? null
  );
}

export interface OfficeWordDefinedStyle {
  readonly styleId: string;
  readonly name: string;
  readonly type?: string;
}

/**
 * Which id a document writes for a gallery style, and whether it already
 * holds that style. A table style of the gallery's NAME is reused as the
 * document defines it; otherwise the first id nobody uses is allocated.
 */
export function officeWordTableStyleIdFor(
  style: OfficeWordTableStyle,
  defined: readonly OfficeWordDefinedStyle[],
): { readonly styleId: string; readonly defined: boolean } {
  const named = defined.find(
    (candidate) =>
      candidate.name === style.name &&
      (candidate.type === undefined || candidate.type === "table"),
  );
  if (named) return { styleId: named.styleId, defined: true };
  const used = new Set(defined.map((candidate) => candidate.styleId));
  if (!used.has(style.styleId)) return { styleId: style.styleId, defined: false };
  for (let suffix = 2; suffix < 1000; suffix += 1) {
    const candidate = `${style.styleId}-${suffix}`;
    if (!used.has(candidate)) return { styleId: candidate, defined: false };
  }
  return { styleId: `${style.styleId}-1000`, defined: false };
}

/** Word's theme tint/shade: HSL luminance in the byte domain, truncating. */
export function officeWordThemeColorHex(
  ref: OfficeWordThemeColorRef,
  theme: OfficeWordThemeColors = OFFICE_WORD_DEFAULT_THEME_COLORS,
): string {
  const base = (theme[ref.theme] ?? OFFICE_WORD_DEFAULT_THEME_COLORS[ref.theme])
    .toLowerCase();
  const tint = ref.tint !== undefined && BYTE.test(ref.tint) ? ref.tint : null;
  const shade =
    ref.shade !== undefined && BYTE.test(ref.shade) ? ref.shade : null;
  // A tint or shade of FF is the identity: the byte-domain round trip would
  // move a channel by one (the server resolver agrees).
  if ((tint === null && shade === null) || (tint ?? shade ?? "").toUpperCase() === "FF") return base;
  const [hue, rawLuminance, saturation] = rgbToHls(
    parseInt(base.slice(1, 3), 16) / 255,
    parseInt(base.slice(3, 5), 16) / 255,
    parseInt(base.slice(5, 7), 16) / 255,
  );
  let luminance = Math.round(rawLuminance * 255) / 255;
  if (tint !== null) {
    luminance = luminance + (1 - luminance) * (1 - parseInt(tint, 16) / 255);
  } else if (shade !== null) {
    luminance *= parseInt(shade, 16) / 255;
  }
  const channels = hlsToRgb(hue, Math.min(Math.max(luminance, 0), 1), saturation);
  return (
    "#" +
    channels
      .map((channel) =>
        Math.trunc(channel * 255)
          .toString(16)
          .padStart(2, "0"),
      )
      .join("")
  );
}

/**
 * `doc.attrs.documentThemeColors`: an imported document's own theme colours
 * (`word/theme/theme1.xml` through `w:clrSchemeMapping`), exactly the ten
 * slots table styles reference; `null` for anything else.
 */
export function normalizeOfficeWordThemeColors(
  value: unknown,
): OfficeWordThemeColors | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (Object.keys(record).length !== OFFICE_WORD_THEME_COLOR_NAMES.length) {
    return null;
  }
  const colors = {} as Record<OfficeWordThemeColorName, string>;
  for (const name of OFFICE_WORD_THEME_COLOR_NAMES) {
    const color = record[name];
    if (typeof color !== "string" || !HEX.test(color)) return null;
    colors[name] = color;
  }
  return colors;
}

/** The text/background slots of the editor's Design colour sets (Office 2013). */
const DESIGN_COLOR_SET_BASE = {
  text1: "#000000",
  background1: "#ffffff",
  text2: "#44546a",
  background2: "#e7e6e6",
} as const;

/**
 * THE theme a Docs document paints in and exports: every theme-referenced
 * colour (gallery table styles, cover pages, their thumbnails, both DOCX
 * writers' theme part) reads this one answer, so what MRAI paints is what
 * Word draws. Precedence: the imported document's own theme
 * (`documentThemeColors`) → the Design tab's document theme
 * (`documentDesign`) → its older colour set (`designColorSet`,
 * with the Office 2013 text/background slots those sets were built on) → the
 * current Office theme (the Presentation gallery's, accent 1 #156082).
 */
export function documentThemeColorsForExport(
  attrs: Readonly<Record<string, unknown>> | null | undefined,
): OfficeWordThemeColors {
  const imported = normalizeOfficeWordThemeColors(attrs?.documentThemeColors);
  if (imported) return imported;
  // The Design tab's document theme (`documentDesign`): its colour scheme.
  const design = normalizeDocumentDesign(attrs?.documentDesign);
  if (design) return documentDesignWordThemeColors(design);
  const colorSet = normalizeOfficeThemePaletteId(attrs?.designColorSet);
  if (colorSet) {
    const accents = paletteById(colorSet).accents.map((color) =>
      color.toLowerCase(),
    );
    return {
      ...DESIGN_COLOR_SET_BASE,
      accent1: accents[0],
      accent2: accents[1],
      accent3: accents[2],
      accent4: accents[3],
      accent5: accents[4],
      accent6: accents[5],
    };
  }
  return OFFICE_WORD_DEFAULT_THEME_COLORS;
}

function visualBorder(
  border: OfficeWordTableStyleBorder,
  theme: OfficeWordThemeColors,
): DocumentTableStyleVisualBorder | null {
  if (border.style === "nil") return null;
  return {
    style: border.style,
    ...(border.size !== undefined ? { sizePt: border.size / 8 } : {}),
    ...(border.color
      ? { color: officeWordThemeColorHex(border.color, theme) }
      : {}),
  };
}

function visualLayer(
  layer: OfficeWordTableStyleLayer,
  theme: OfficeWordThemeColors,
  target: "table" | "cell",
): DocumentTableStyleVisualLayer {
  const lines: Partial<
    Record<DocumentTableVisualEdge, DocumentTableStyleVisualBorder>
  > = {};
  const cleared: DocumentTableVisualEdge[] = [];
  for (const edge of EDGE_ORDER) {
    const border = layer.borders?.[edge];
    if (!border) continue;
    const line = visualBorder(border, theme);
    if (line) lines[edge] = line;
    else cleared.push(edge);
  }
  return {
    ...(layer.fill ? { fill: officeWordThemeColorHex(layer.fill, theme) } : {}),
    ...(layer.run?.color
      ? { color: officeWordThemeColorHex(layer.run.color, theme) }
      : {}),
    ...(layer.run?.bold ? { bold: true } : {}),
    tableBorders: target === "table" ? lines : {},
    cellBorders: target === "cell" ? lines : {},
    clearedTableBorders: target === "table" ? cleared : [],
    clearedCellBorders: target === "cell" ? cleared : [],
  };
}

/**
 * The style as the Docs table renderer draws it: the same cascade model the
 * importer projects from a document's own `styles.xml`, so a gallery style
 * paints exactly like the definition Word reads back.
 */
const visualCache = new Map<string, DocumentTableStyleVisualAuthority>();

export function officeWordTableStyleVisual(
  style: OfficeWordTableStyle,
  styleId: string,
  theme: OfficeWordThemeColors = OFFICE_WORD_DEFAULT_THEME_COLORS,
): DocumentTableStyleVisualAuthority {
  // Every table decoration pass asks again; the answer depends only on these
  // (and on whether the style is the recipe's or the served one).
  const key = `${fallbackTables?.bySlot.get(style.slot) === style ? "fallback" : "served"}|${style.slot}|${styleId}|${OFFICE_WORD_THEME_COLOR_NAMES.map(
    (name) => theme[name],
  ).join(",")}`;
  const cached = visualCache.get(key);
  if (cached) return cached;
  const visual = buildOfficeWordTableStyleVisual(style, styleId, theme);
  if (visualCache.size >= 512) visualCache.clear();
  visualCache.set(key, visual);
  return visual;
}

function buildOfficeWordTableStyleVisual(
  style: OfficeWordTableStyle,
  styleId: string,
  theme: OfficeWordThemeColors,
): DocumentTableStyleVisualAuthority {
  const conditionals: Partial<
    Record<DocumentTableVisualCondition, DocumentTableStyleVisualLayer>
  > = {};
  for (const condition of CONDITION_ORDER) {
    const layer = style.conditions[condition];
    if (layer) conditionals[condition] = visualLayer(layer, theme, "cell");
  }
  return {
    schemaVersion: 1,
    styleId,
    rowBandSize: 1,
    columnBandSize: 1,
    wholeTable: visualLayer(style.table, theme, "table"),
    conditionals,
  };
}

/** Word's look for a table that never chose one (`w:val="04A0"`). */
export const OFFICE_WORD_DEFAULT_TABLE_LOOK: DocumentTableLook = Object.freeze({
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  bandedRows: true,
  bandedColumns: false,
});
