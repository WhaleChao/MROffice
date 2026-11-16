/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Native PowerPoint Web's 74 positions, captured from the owner's real deck.
 * No invented GUIDs or one-off approximated style families. Native XML lives
 * in the API's matching catalogue; this symbolic projection drives the UI.
 *
 * The looks (`presentation-table-native-catalog.json`) are served by the API
 * to a signed-in page (`lib/office-data/store.ts`), not bundled: until they
 * arrive, and on a guest page, a style has its slot, group and name
 * (`office-table-style-index.ts`) and the look MRAI's own recipe of its
 * family draws (`office-table-style-fallback.ts`), and a table draws the
 * look it carries itself. */
import { officeData } from "./office-data/store";
import {
  fallbackPresentationTableCatalog,
  isFallbackPresentationTableCatalog,
  type TableStyleCatalog,
} from "./office-table-style-fallback";
import { PRESENTATION_TABLE_STYLE_INDEX } from "./office-table-style-index";
import {
  officeColorChannelsFromHex,
  officeColorChannelsFromScrgb,
  officeColorHexFromChannels,
  officeColorTransformChannels,
  type OfficeColorChannels,
} from "./office-color-transform";
import {
  normalizeTableStyleLook,
  type TableStyleLook,
} from "./presentation-table-style";

type Token = {
  type: string;
  value: Record<string, string>;
  transforms: string[][];
  base?: string;
};
const presets: Record<string, string> = {
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  blue: "#0000ff",
  green: "#008000",
};
type Catalog = TableStyleCatalog;

/** The served, measured looks, or null until they arrive (a guest never has
 * them). */
function servedCatalog(): Catalog | null {
  const value = officeData("presentation-table-styles") as Catalog | undefined;
  return value && Array.isArray(value.styles) && Array.isArray(value.colors) ? value : null;
}

/** The looks the gallery draws now: the served ones, else the recipe's. */
function nativeCatalog(): Catalog {
  return servedCatalog() ?? fallbackPresentationTableCatalog();
}

/** Whether the gallery is drawing the recipe (the measured looks are away). */
export function webTableStyleGalleryIsFallback(): boolean {
  return isFallbackPresentationTableCatalog(nativeCatalog());
}

/**
 * The deck's own colours a table style's theme colours resolve to besides its
 * accents: Dark 1 / Light 1 / Dark 2 / Light 2 and the link colours, and the
 * master's colour map, which names the slot Text 1 / Background 1 / Text 2 /
 * Background 2 stand for (Light 1 for Text 1 on a dark master). PowerPoint
 * draws a table's style in the slide's theme through that map: "No Style,
 * Table Grid" writes its words and lines in Text 1, white on a dark deck.
 * Without them, Office's own theme (black Dark 1, the light map).
 */
export interface TableStyleSchemeColors {
  readonly slots?: Readonly<Partial<Record<"dk1" | "lt1" | "dk2" | "lt2" | "hlink" | "folHlink", string>>>;
  readonly map?: Readonly<Partial<Record<"tx1" | "tx2" | "bg1" | "bg2", string>>>;
}

const LIGHT_TABLE_STYLE_MAP: Readonly<Record<string, string>> = { tx1: "dk1", tx2: "dk2", bg1: "lt1", bg2: "lt2" };
const TABLE_STYLE_SCHEME_SLOTS = ["dk1", "lt1", "dk2", "lt2", "hlink", "folHlink"] as const;
const TABLE_STYLE_MAPPED_NAMES = ["tx1", "tx2", "bg1", "bg2"] as const;
const TABLE_STYLE_MAP_TARGETS = new Set<string>([...TABLE_STYLE_SCHEME_SLOTS.slice(0, 4)]);

function resolveColor(
  catalog: Catalog,
  value: string,
  theme: Record<string, string>,
  map: Readonly<Record<string, string>> = LIGHT_TABLE_STYLE_MAP,
): { color: string; opacity: number } {
  if (!value.startsWith("$")) return { color: value, opacity: 1 };
  const token = catalog.colors[Number(value.slice(1))] as Token;
  const slot = map[token.value.val] ?? token.value.val;
  const base = token.base
    ? resolveColor(catalog, token.base, theme, map).color
    : token.type === "schemeClr"
      ? (theme[slot] ?? "#000000")
      : token.type === "srgbClr"
        ? "#" + token.value.val
        : token.type === "sysClr"
          ? "#" + token.value.lastClr
          : token.type === "prstClr"
            ? (presets[token.value.val] ?? "#000000")
            : officeColorHexFromChannels(
                officeColorChannelsFromScrgb(
                  Number(token.value.r),
                  Number(token.value.g),
                  Number(token.value.b),
                ),
              );
  // The arithmetic is the one rule every host asks (`office-color-transform`):
  // a Medium style's band is `accent + tint`, which Office mixes in linear
  // light -- `#e9ebf5` for accent 1 + 20 %, where mixing the encoded channels
  // drew `#dae3f3`.
  let channels: OfficeColorChannels = officeColorChannelsFromHex(base) ?? [
      0, 0, 0,
    ],
    opacity = 1;
  for (const [kind, raw] of token.transforms) {
    if (kind === "alpha") opacity = Number(raw) / 100000;
    else
      channels =
        officeColorTransformChannels(channels, kind, Number(raw)) ?? channels;
  }
  return { color: officeColorHexFromChannels(channels), opacity };
}

export type WebTableStyleGroup = "best" | "light" | "medium" | "dark";
export interface WebTableStyle {
  slot: number;
  group: WebTableStyleGroup;
  label: string;
  /** The name as a `console-notebooks` key (`office.tableStyle.*`). */
  labelKey: string;
  look: TableStyleLook;
}
/** The 74 styles with their looks in these accents: the served looks, or
 * the recipe's until they arrive. */
export function createWebTableStyleGallery(
  accentColors: readonly string[] = [],
  locale: "en" | "zh" = "en",
  scheme?: TableStyleSchemeColors,
): readonly WebTableStyle[] {
  return galleryFromCatalog(nativeCatalog(), accentColors, locale, scheme);
}

function galleryFromCatalog(
  catalog: Catalog,
  accentColors: readonly string[],
  locale: "en" | "zh",
  scheme?: TableStyleSchemeColors,
): readonly WebTableStyle[] {
  const theme: Record<string, string> = { ...catalog.defaults };
  for (const slot of TABLE_STYLE_SCHEME_SLOTS) {
    const color = scheme?.slots?.[slot];
    if (typeof color === "string" && /^#[a-f0-9]{6}$/i.test(color)) theme[slot] = color;
  }
  accentColors.slice(0, 6).forEach((color, i) => {
    if (/^#[a-f0-9]{6}$/i.test(color)) theme[`accent${i + 1}`] = color;
  });
  const map: Record<string, string> = { ...LIGHT_TABLE_STYLE_MAP };
  for (const name of TABLE_STYLE_MAPPED_NAMES) {
    const target = scheme?.map?.[name];
    if (typeof target === "string" && TABLE_STYLE_MAP_TARGETS.has(target)) map[name] = target;
  }
  return catalog.styles.map((entry) => {
    const raw = JSON.parse(JSON.stringify(entry)) as Record<string, unknown>;
    const parts = raw.parts as Record<string, Record<string, unknown>>;
    for (const part of Object.values(parts)) {
      for (const field of ["fill", "color"])
        if (typeof part[field] === "string") {
          const color = resolveColor(catalog, part[field] as string, theme, map);
          part[field] = color.color;
          if (field === "fill" && color.opacity < 1)
            part.fillOpacity = color.opacity;
        }
      for (const border of Object.values(
        (part.borders ?? {}) as Record<string, Record<string, unknown>>,
      ))
        if (typeof border.color === "string")
          border.color = resolveColor(catalog, border.color, theme, map).color;
    }
    if (typeof raw.background === "string")
      raw.background = resolveColor(catalog, raw.background, theme, map).color;
    const gradient = raw.backgroundGradient as
      | { angle: number; stops: Array<{ color: string; offset: number }> }
      | undefined;
    if (gradient)
      gradient.stops = gradient.stops.map((stop) => ({
        ...stop,
        color: resolveColor(catalog, stop.color, theme, map).color,
      }));
    const look = normalizeTableStyleLook(raw, entry.styleId)!;
    return {
      slot: entry.slot,
      group: entry.group as WebTableStyleGroup,
      label: webTableStyleLabel(entry.styleName, locale),
      labelKey: webTableStyleLabelKey(entry.styleName),
      look,
    };
  });
}

/**
 * The name a gallery tile shows. The catalog's `label` is the frozen Web
 * menu's text, which names the first dark style "深 样式 2" (a second one)
 * and numbers the Dark Style 2 accent pairs wrongly; `styleName` is Office's
 * own name for the built-in style (its GUID), so both languages come from it.
 */
function webTableStyleLabel(styleName: string, locale: "en" | "zh"): string {
  return locale === "zh"
    ? styleName
    : styleName
        .replace("无样式，无网格", "No Style, No Grid")
        .replace("无样式，网格型", "No Style, Table Grid")
        .replace("主题样式", "Themed Style")
        .replace("浅色样式", "Light Style")
        .replace("中度样式", "Medium Style")
        .replace("深色样式", "Dark Style")
        .replaceAll("强调", "Accent");
}

/**
 * The style's name as a `console-notebooks` catalogue key: Office's English
 * name in camel case ("Medium Style 2 - Accent 1" is
 * `office.tableStyle.mediumStyle2Accent1`), so every locale can name all 74.
 */
export function webTableStyleLabelKey(styleName: string): string {
  const words = webTableStyleLabel(styleName, "en")
    .split(/[^A-Za-z0-9]+/)
    .filter(Boolean);
  return `office.tableStyle.${words
    .map((word, index) =>
      index ? word[0].toUpperCase() + word.slice(1) : word.toLowerCase(),
    )
    .join("")}`;
}

export interface WebTableStyleSlot {
  slot: number;
  group: WebTableStyleGroup;
  label: string;
  labelKey: string;
}

/** The 74 gallery positions without resolving any recipe (every host's menu). */
/** The gallery's slots and names (no looks): always here, served or not. */
export function webTableStyleGallerySlots(
  locale: "en" | "zh" = "en",
): readonly WebTableStyleSlot[] {
  return PRESENTATION_TABLE_STYLE_INDEX.map(([slot, group, , styleName]) => ({
    slot,
    group,
    label: webTableStyleLabel(styleName, locale),
    labelKey: webTableStyleLabelKey(styleName),
  }));
}

/**
 * The catalogue key (`office.tableStyle.*`) a table style is shown under in
 * the UI's language: a built-in's by its GUID, else by its name when the name
 * is one of the 74 (the catalogue's Chinese capture or PowerPoint's English,
 * which a save writes and readings row 893 reads); null for a deck's own
 * style, shown under the name the deck gives it.
 */
export function presentationTableStyleLabelKey(look: { styleId?: unknown; styleName?: unknown }): string | null {
  const id = typeof look.styleId === "string" ? look.styleId.trim().toUpperCase() : "";
  const byId = PRESENTATION_TABLE_STYLE_INDEX.find(([, , styleId]) => styleId.toUpperCase() === id);
  if (byId) return webTableStyleLabelKey(byId[3]);
  if (typeof look.styleName !== "string" || !look.styleName.trim()) return null;
  const key = webTableStyleLabelKey(look.styleName.trim());
  return PRESENTATION_TABLE_STYLE_INDEX.some(([, , , styleName]) => webTableStyleLabelKey(styleName) === key) ? key : null;
}

const NATIVE_STYLE_IDS: ReadonlySet<string> = new Set(
  PRESENTATION_TABLE_STYLE_INDEX.map(([, , styleId]) => styleId.toUpperCase()),
);

/** Whether an id is one of the 74 built-in styles (no look needed). */
export function isNativeWebTableStyleId(id: unknown): boolean {
  return typeof id === "string" && NATIVE_STYLE_IDS.has(id.trim().toUpperCase());
}

/** The gallery slot of a built-in style id (no look needed), else null. */
export function nativeWebTableStyleSlot(id: unknown): number | null {
  if (typeof id !== "string") return null;
  const wanted = id.trim().toUpperCase();
  return PRESENTATION_TABLE_STYLE_INDEX.find(([, , styleId]) => styleId.toUpperCase() === wanted)?.[0] ?? null;
}

let defaultCatalogById: { readonly catalog: Catalog; readonly byId: ReadonlyMap<string, WebTableStyle> } | null = null;

/** A built-in style with its look in the default theme (the served look,
 * or the recipe's until it arrives); null for any other id. */
export function nativeWebTableStyleById(id: unknown): WebTableStyle | null {
  if (typeof id !== "string") return null;
  const catalog = nativeCatalog();
  if (defaultCatalogById?.catalog !== catalog) {
    defaultCatalogById = {
      catalog,
      byId: new Map(galleryFromCatalog(catalog, [], "en").map((entry) => [entry.look.styleId, entry])),
    };
  }
  return defaultCatalogById.byId.get(id.trim().toUpperCase()) ?? null;
}

/**
 * A built-in style's look resolved against a theme's accents, or null for a
 * style that is not one of the built-in ones. A table made here stores the
 * look it was given; when the deck's theme changes, that look is resolved
 * again so "Medium Style 2 - Accent 1" is the new accent 1, not the accent
 * of the day it was picked (PPT-B274).
 */
export function webTableStyleLookForTheme(
  styleId: unknown,
  accentColors: readonly string[],
  scheme?: TableStyleSchemeColors,
): TableStyleLook | null {
  if (!nativeWebTableStyleById(styleId)) return null;
  const id = String(styleId).trim().toUpperCase();
  return (
    createWebTableStyleGallery(accentColors, "en", scheme).find((entry) => entry.look.styleId === id)?.look ??
    null
  );
}

/**
 * The built-in style's look as the recipe draws it in these accents: what a
 * table styled while the measured looks were away carries, so a deck can
 * tell that look from one its file stated when the looks land. Null for a
 * style that is not one of the built-in ones.
 */
export function fallbackWebTableStyleLookForTheme(
  styleId: unknown,
  accentColors: readonly string[],
  scheme?: TableStyleSchemeColors,
): TableStyleLook | null {
  if (!isNativeWebTableStyleId(styleId)) return null;
  const id = String(styleId).trim().toUpperCase();
  return (
    galleryFromCatalog(fallbackPresentationTableCatalog(), accentColors, "en", scheme).find(
      (entry) => entry.look.styleId === id,
    )?.look ?? null
  );
}
