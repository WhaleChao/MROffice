/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word-aligned cover pages: the catalog, the property bindings, and the one
 * drawable "scene" every surface paints.
 *
 * Word's Insert > Cover Page gallery offers seventeen built-in covers. MRAI
 * reproduces the gallery's behaviour (Microsoft Support, "Add a cover page";
 * Microsoft Learn, "Bind a content control to a node in the data store"):
 *
 * - a cover is ONE unit at the very start of the document, on its own page;
 *   inserting another replaces it, and "Remove Current Cover Page" deletes it;
 * - in the package it is a body `w:sdt` whose `w:docPartObj` names the
 *   "Cover Pages" gallery and is `w:docPartUnique`; Word finds the current
 *   cover by that gallery, never by its look;
 * - its words are content controls mapped (`w:dataBinding`) to document
 *   properties: the core properties (title, subject, creator, category), the
 *   extended property Company, and the Cover Page Properties store (Abstract,
 *   PublishDate, CompanyAddress, CompanyPhone, CompanyFax, CompanyEmail).
 *   Typing into one control changes the property, and every other control
 *   mapped to it shows the new value; an empty property shows the control's
 *   placeholder text.
 *
 * The seventeen designs are MRAI's own geometry (bars, blocks, gradients,
 * lines, paths) in theme accent colours: `coverPageCatalog.json` holds them,
 * and its byte-identical twin in the API
 * (`app/services/document_cover_page_catalog.json`) writes the same cover
 * into an imported package. Nothing in either file comes from Microsoft's
 * building blocks; only the gallery names and the binding targets do.
 *
 * A {@link CoverScene} is what a cover looks like on one page size: absolute
 * shapes, text frames and picture placeholders in EMU, back to front. The
 * gallery thumbnails, the editor's cover atom and (through the server) an
 * imported Word cover are all drawn from a scene, so a thumbnail can never
 * disagree with what the insert command puts on the page.
 */

import catalogJson from "./coverPageCatalog.json";
import type { DocumentMetadataKey } from "../../components/console/editor/native-document/documentMetadata";

/** The editor atom a cover page is (one per document, at its start). */
export const DOCUMENT_COVER_PAGE_NODE = "documentCoverPage";

// ---------------------------------------------------------------------------
// Catalog
// ---------------------------------------------------------------------------

export type CoverThemeSlot =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6";

export interface CoverThemePaint {
  theme: CoverThemeSlot;
  lumMod?: number;
  lumOff?: number;
  alpha?: number;
}

export interface CoverGradientPaint {
  gradient: {
    angle: number;
    stops: ReadonlyArray<{ pos: number; paint: CoverThemePaint }>;
  };
}

export type CoverDefinitionPaint = CoverThemePaint | CoverGradientPaint;

export type CoverDefinitionGeometry =
  | "rect"
  | "ellipse"
  | "line"
  | "roundRect"
  | { d: string; open?: boolean };

export interface CoverDefinitionRun {
  field?: CoverFieldKey;
  text?: string;
  size: number;
  paint: CoverThemePaint;
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  /** Letter spacing in points (`w:spacing`). */
  spacing?: number;
  /** A PublishDate control's picture, by role. */
  dateFormat?: "short" | "year";
}

export interface CoverDefinitionParagraph {
  align: "left" | "center" | "right";
  before?: number;
  after?: number;
  runs: readonly CoverDefinitionRun[];
}

export interface CoverDefinitionTextBody {
  anchor: "t" | "ctr" | "b";
  /** Left, top, right, bottom insets in points. */
  inset: readonly [number, number, number, number] | readonly number[];
  paragraphs: readonly CoverDefinitionParagraph[];
}

export interface CoverDefinitionItem {
  type: "shape" | "text" | "picture";
  id: string;
  /** x, y, width, height as fractions of the page. */
  box: readonly number[];
  geom?: CoverDefinitionGeometry;
  fill?: CoverDefinitionPaint;
  line?: { paint: CoverThemePaint; widthPt: number };
  behind: boolean;
  text?: CoverDefinitionTextBody;
}

export interface CoverDefinition {
  id: string;
  names: Readonly<Record<string, string>>;
  fields: readonly CoverFieldKey[];
  items: readonly CoverDefinitionItem[];
}

interface CoverCatalog {
  version: number;
  orders: Readonly<Record<string, readonly string[]>>;
  placeholders: Readonly<Record<string, Readonly<Record<string, string>>>>;
  aliases: Readonly<Record<string, Readonly<Record<string, string>>>>;
  dateFormats: Readonly<Record<string, Readonly<Record<string, string>>>>;
  covers: readonly CoverDefinition[];
}

const CATALOG = catalogJson as unknown as CoverCatalog;

export const DOCUMENT_COVER_PAGE_IDS: readonly string[] = CATALOG.covers.map(
  (cover) => cover.id,
);

const COVER_BY_ID = new Map(CATALOG.covers.map((cover) => [cover.id, cover]));

export function isDocumentCoverPageId(value: unknown): value is string {
  return typeof value === "string" && COVER_BY_ID.has(value);
}

export function documentCoverDefinition(id: unknown): CoverDefinition | null {
  return typeof id === "string" ? (COVER_BY_ID.get(id) ?? null) : null;
}

/**
 * The catalog language a UI or document locale reads: Simplified Chinese
 * (zh, zh-CN, zh-Hans, zh-SG), Traditional (zh-TW, zh-HK, zh-Hant), the
 * locale's own language where the catalog has it, else English.
 */
export function coverCatalogLanguage(locale: unknown): string {
  if (typeof locale !== "string" || !locale) return "en";
  const tag = locale.toLowerCase().replace(/_/g, "-");
  const known = Object.keys(CATALOG.placeholders);
  if (/^zh(?:-(?:cn|sg|hans)(?:-.*)?)?$/.test(tag)) return "zh";
  // Traditional Chinese has its own words; never the Simplified ones.
  if (/^zh-(?:tw|hk|mo|hant)/.test(tag)) return known.includes("zh-TW") ? "zh-TW" : "en";
  const exact = known.find((language) => language.toLowerCase() === tag);
  if (exact) return exact;
  const base = tag.split("-")[0];
  return known.find((language) => language.toLowerCase() === base) ?? "en";
}

/**
 * The gallery order. Word lists its built-in covers sorted by their localized
 * names, so the Chinese gallery reads 奥斯汀, 标题, 裁剪, … and the English one
 * Austin, Badge, Banded, ….
 */
export function documentCoverGalleryOrder(locale: unknown): readonly string[] {
  return CATALOG.orders[coverCatalogLanguage(locale)] ?? CATALOG.orders.en;
}

export function documentCoverName(id: string, locale: unknown): string {
  const cover = COVER_BY_ID.get(id);
  if (!cover) return id;
  return cover.names[coverCatalogLanguage(locale)] ?? cover.names.en ?? id;
}

// ---------------------------------------------------------------------------
// Fields and their Word bindings
// ---------------------------------------------------------------------------

export const COVER_FIELD_KEYS = [
  "title",
  "subtitle",
  "author",
  "category",
  "keywords",
  "company",
  "abstract",
  "publishDate",
  "companyAddress",
  "companyPhone",
  "companyFax",
  "companyEmail",
] as const;

export type CoverFieldKey = (typeof COVER_FIELD_KEYS)[number];

export function isCoverFieldKey(value: unknown): value is CoverFieldKey {
  return (
    typeof value === "string" &&
    (COVER_FIELD_KEYS as readonly string[]).includes(value)
  );
}

/** The data store IDs Word uses for its built-in property parts. */
export const CORE_PROPERTIES_STORE_ID = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}";
export const EXTENDED_PROPERTIES_STORE_ID =
  "{6668398D-A668-4E3E-A5EB-62B293D839F1}";
export const COVER_PAGE_PROPERTIES_STORE_ID =
  "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}";
export const COVER_PAGE_PROPERTIES_NAMESPACE =
  "http://schemas.microsoft.com/office/2006/coverPageProps";

const DC_NS = "http://purl.org/dc/elements/1.1/";
const CP_NS =
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const EP_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

export interface CoverFieldBinding {
  storeItemId: string;
  xpath: string;
  prefixMappings: string;
}

const CORE_PREFIXES = `xmlns:ns0='${DC_NS}' xmlns:ns1='${CP_NS}' `;
const EXTENDED_PREFIXES = `xmlns:ns0='${EP_NS}' `;
const COVER_PREFIXES = `xmlns:ns0='${COVER_PAGE_PROPERTIES_NAMESPACE}' `;

/** The XML mapping Word itself writes for each cover field. */
export const COVER_FIELD_BINDINGS: Readonly<
  Record<CoverFieldKey, CoverFieldBinding>
> = {
  title: {
    storeItemId: CORE_PROPERTIES_STORE_ID,
    xpath: "/ns1:coreProperties[1]/ns0:title[1]",
    prefixMappings: CORE_PREFIXES,
  },
  subtitle: {
    storeItemId: CORE_PROPERTIES_STORE_ID,
    xpath: "/ns1:coreProperties[1]/ns0:subject[1]",
    prefixMappings: CORE_PREFIXES,
  },
  author: {
    storeItemId: CORE_PROPERTIES_STORE_ID,
    xpath: "/ns1:coreProperties[1]/ns0:creator[1]",
    prefixMappings: CORE_PREFIXES,
  },
  category: {
    storeItemId: CORE_PROPERTIES_STORE_ID,
    xpath: "/ns1:coreProperties[1]/ns1:category[1]",
    prefixMappings: CORE_PREFIXES,
  },
  keywords: {
    storeItemId: CORE_PROPERTIES_STORE_ID,
    xpath: "/ns1:coreProperties[1]/ns1:keywords[1]",
    prefixMappings: CORE_PREFIXES,
  },
  company: {
    storeItemId: EXTENDED_PROPERTIES_STORE_ID,
    xpath: "/ns0:Properties[1]/ns0:Company[1]",
    prefixMappings: EXTENDED_PREFIXES,
  },
  abstract: coverStoreBinding("Abstract"),
  publishDate: coverStoreBinding("PublishDate"),
  companyAddress: coverStoreBinding("CompanyAddress"),
  companyPhone: coverStoreBinding("CompanyPhone"),
  companyFax: coverStoreBinding("CompanyFax"),
  companyEmail: coverStoreBinding("CompanyEmail"),
};

function coverStoreBinding(element: string): CoverFieldBinding {
  return {
    storeItemId: COVER_PAGE_PROPERTIES_STORE_ID,
    xpath: `/ns0:CoverPageProperties[1]/ns0:${element}[1]`,
    prefixMappings: COVER_PREFIXES,
  };
}

/** The Cover Page Properties elements, in Word's own order. */
export const COVER_PAGE_PROPERTY_ELEMENTS: ReadonlyArray<
  readonly [string, CoverFieldKey]
> = [
  ["PublishDate", "publishDate"],
  ["Abstract", "abstract"],
  ["CompanyAddress", "companyAddress"],
  ["CompanyPhone", "companyPhone"],
  ["CompanyFax", "companyFax"],
  ["CompanyEmail", "companyEmail"],
];

/** The document property (a `documentMetadata` key) a field edits. */
export const COVER_FIELD_METADATA_KEY: Readonly<
  Record<CoverFieldKey, DocumentMetadataKey>
> = {
  title: "title",
  subtitle: "subject",
  author: "author",
  category: "category",
  keywords: "keywords",
  company: "company",
  abstract: "abstract",
  publishDate: "publishDate",
  companyAddress: "companyAddress",
  companyPhone: "companyPhone",
  companyFax: "companyFax",
  companyEmail: "companyEmail",
};

/**
 * Which field a Word `w:dataBinding` targets, read by namespace URI and local
 * name (never by prefix), or null for a mapping that is not a document
 * property. The step grammar is closed: `/p:name[1]/p:name[1]`.
 */
export function coverFieldFromBinding(
  binding: Partial<CoverFieldBinding> | null | undefined,
): CoverFieldKey | null {
  if (!binding || typeof binding.xpath !== "string") return null;
  const prefixes = new Map<string, string>();
  for (const match of (binding.prefixMappings ?? "").matchAll(
    /xmlns:([A-Za-z_][\w.-]*)\s*=\s*(['"])(.*?)\2/g,
  )) {
    prefixes.set(match[1], match[3]);
  }
  const steps = binding.xpath.split("/").filter(Boolean);
  if (steps.length !== 2) return null;
  const qualified = steps.map((step) => {
    const match = /^(?:([A-Za-z_][\w.-]*):)?([A-Za-z_][\w.-]*)(?:\[1\])?$/.exec(
      step,
    );
    if (!match) return null;
    return { uri: match[1] ? (prefixes.get(match[1]) ?? null) : null, local: match[2] };
  });
  const [root, leaf] = qualified;
  if (!root || !leaf) return null;
  if (root.uri === CP_NS && root.local === "coreProperties") {
    if (leaf.uri === DC_NS) {
      if (leaf.local === "title") return "title";
      if (leaf.local === "subject") return "subtitle";
      if (leaf.local === "creator") return "author";
    }
    if (leaf.uri === CP_NS) {
      if (leaf.local === "category") return "category";
      if (leaf.local === "keywords") return "keywords";
    }
    return null;
  }
  if (root.uri === EP_NS && root.local === "Properties") {
    return leaf.uri === EP_NS && leaf.local === "Company" ? "company" : null;
  }
  if (
    root.uri === COVER_PAGE_PROPERTIES_NAMESPACE &&
    root.local === "CoverPageProperties" &&
    leaf.uri === COVER_PAGE_PROPERTIES_NAMESPACE
  ) {
    const entry = COVER_PAGE_PROPERTY_ELEMENTS.find(
      ([element]) => element === leaf.local,
    );
    return entry ? entry[1] : null;
  }
  return null;
}

export function coverFieldPlaceholder(
  field: CoverFieldKey,
  locale: unknown,
): string {
  const language = coverCatalogLanguage(locale);
  return (
    CATALOG.placeholders[language]?.[field] ??
    CATALOG.placeholders.en?.[field] ??
    `[${field}]`
  );
}

export function coverFieldAlias(field: CoverFieldKey, locale: unknown): string {
  const language = coverCatalogLanguage(locale);
  return (
    CATALOG.aliases[language]?.[field] ?? CATALOG.aliases.en?.[field] ?? field
  );
}

export function coverDatePicture(
  role: "short" | "year" | undefined,
  locale: unknown,
): string {
  const language = coverCatalogLanguage(locale);
  const formats = CATALOG.dateFormats[language] ?? CATALOG.dateFormats.en;
  return formats[role ?? "short"] ?? "yyyy-MM-dd";
}

// ---------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------

const MONTHS_EN = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

/** `YYYY-MM-DD` (optionally with Word's `T00:00:00[Z]` tail) as a date. */
export function parseCoverDate(
  value: string,
): { year: number; month: number; day: number } | null {
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:T[\d:.]+(?:Z|[+-]\d{2}:?\d{2})?)?$/.exec(
    value.trim(),
  );
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { year, month, day };
}

/**
 * What a PublishDate control prints: the `w:dateFormat` picture applied to
 * the stored date. A value that is not a date prints as stored, as in Word.
 */
export function formatCoverDate(value: string, picture: string): string {
  const date = parseCoverDate(value);
  if (!date) return value;
  return picture.replace(
    /yyyy|yy|MMMM|MMM|MM|M|dd|d/g,
    (token) => {
      switch (token) {
        case "yyyy":
          return String(date.year).padStart(4, "0");
        case "yy":
          return String(date.year % 100).padStart(2, "0");
        case "MMMM":
          return MONTHS_EN[date.month - 1];
        case "MMM":
          return MONTHS_EN[date.month - 1].slice(0, 3);
        case "MM":
          return String(date.month).padStart(2, "0");
        case "M":
          return String(date.month);
        case "dd":
          return String(date.day).padStart(2, "0");
        default:
          return String(date.day);
      }
    },
  );
}

/**
 * What typing into a date control stores: an ISO day when the text reads as
 * one (`2026-9-23`, `2026/9/23`, `9/23/2026`), otherwise the text itself.
 */
export function normalizeCoverDateInput(text: string): string {
  const value = text.trim();
  const iso = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(value);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value);
  const parts = iso
    ? [iso[1], iso[2], iso[3]]
    : us
      ? [us[3], us[1], us[2]]
      : null;
  if (!parts) return value;
  const [year, month, day] = parts.map(Number);
  if (month < 1 || month > 12 || day < 1 || day > 31) return value;
  return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

// ---------------------------------------------------------------------------
// Theme colours
// ---------------------------------------------------------------------------

/** The four theme slots the palette catalog does not carry (Office theme). */
const OFFICE_THEME_BASE: Readonly<Record<"dk1" | "lt1" | "dk2" | "lt2", string>> = {
  dk1: "#000000",
  lt1: "#ffffff",
  dk2: "#44546a",
  lt2: "#e7e6e6",
};

export interface CoverThemeColors {
  accents: readonly string[];
  /**
   * The text/background slots (dk1/lt1/dk2/lt2) of the document's theme
   * (`documentThemeColorsForExport`); absent, the Office 2013 ones.
   */
  base?: Readonly<Partial<Record<"dk1" | "lt1" | "dk2" | "lt2", string>>>;
}

function hexToRgb(hex: string): [number, number, number] {
  const value = /^#([0-9a-f]{6})$/i.exec(hex)?.[1] ?? "000000";
  return [
    parseInt(value.slice(0, 2), 16),
    parseInt(value.slice(2, 4), 16),
    parseInt(value.slice(4, 6), 16),
  ];
}

function rgbToHex([r, g, b]: readonly number[]): string {
  return `#${[r, g, b]
    .map((channel) =>
      Math.max(0, Math.min(255, Math.round(channel)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
}

function rgbToHsl([r, g, b]: readonly number[]): [number, number, number] {
  const red = r / 255;
  const green = g / 255;
  const blue = b / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation =
    lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  const hue =
    max === red
      ? (green - blue) / delta + (green < blue ? 6 : 0)
      : max === green
        ? (blue - red) / delta + 2
        : (red - green) / delta + 4;
  return [hue / 6, saturation, lightness];
}

function hslToRgb([h, s, l]: readonly number[]): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t: number) => {
    let value = t;
    if (value < 0) value += 1;
    if (value > 1) value -= 1;
    if (value < 1 / 6) return p + (q - p) * 6 * value;
    if (value < 1 / 2) return q;
    if (value < 2 / 3) return p + (q - p) * (2 / 3 - value) * 6;
    return p;
  };
  return [channel(h + 1 / 3) * 255, channel(h) * 255, channel(h - 1 / 3) * 255];
}

/**
 * A theme slot with DrawingML's luminance modulation applied
 * (`a:lumMod`/`a:lumOff`, ECMA-376 §20.1.2.3.20-21).
 */
export function resolveCoverThemeColor(
  paint: CoverThemePaint,
  theme: CoverThemeColors,
): string {
  const accent = /^accent([1-6])$/.exec(paint.theme);
  const base = accent
    ? (theme.accents[Number(accent[1]) - 1] ?? "#4472c4")
    : (theme.base?.[paint.theme as "dk1" | "lt1" | "dk2" | "lt2"] ??
      OFFICE_THEME_BASE[paint.theme as "dk1" | "lt1" | "dk2" | "lt2"] ??
      "#000000");
  if (paint.lumMod === undefined && paint.lumOff === undefined) {
    return base.toLowerCase();
  }
  const [h, s, l] = rgbToHsl(hexToRgb(base));
  const lightness = Math.max(
    0,
    Math.min(1, l * (paint.lumMod ?? 1) + (paint.lumOff ?? 0)),
  );
  return rgbToHex(hslToRgb([h, s, lightness]));
}

// ---------------------------------------------------------------------------
// Scene
// ---------------------------------------------------------------------------

/**
 * The theme slot a colour was drawn from, with the luminance modulation the
 * design applied (an MRAI design only; an imported scene carries the colour
 * alone). The exporters write it as a theme reference so the cover follows
 * the document theme in Word; `color` is the same slot resolved here.
 */
export interface CoverSceneThemeColor {
  slot: CoverThemeSlot;
  lumMod?: number;
  lumOff?: number;
}

export type CoverScenePaint =
  | { color: string; alpha?: number; theme?: CoverSceneThemeColor }
  | {
      gradient: {
        angle: number;
        stops: ReadonlyArray<{
          pos: number;
          color: string;
          alpha?: number;
          theme?: CoverSceneThemeColor;
        }>;
      };
    };

export interface CoverScenePath {
  d: string;
  /** `none` draws only the outline of this path. */
  fill?: "none" | "norm";
  stroke?: boolean;
}

export interface CoverSceneRun {
  text?: string;
  field?: CoverFieldKey;
  /** Shown while the bound property is empty. */
  placeholder?: string;
  /** A date control's `w:dateFormat` picture. */
  dateFormat?: string;
  sizePt: number;
  color: string;
  /** The theme slot `color` resolves (see {@link CoverSceneThemeColor}). */
  colorTheme?: CoverSceneThemeColor;
  bold?: boolean;
  italic?: boolean;
  caps?: boolean;
  spacingPt?: number;
  fontFamily?: string;
}

export interface CoverSceneParagraph {
  align: "left" | "center" | "right" | "justify";
  beforePt?: number;
  afterPt?: number;
  lineHeight?: number;
  runs: readonly CoverSceneRun[];
}

export interface CoverSceneTextBody {
  anchor: "t" | "ctr" | "b";
  /** Left, top, right, bottom insets in EMU. */
  inset: readonly [number, number, number, number];
  paragraphs: readonly CoverSceneParagraph[];
}

export interface CoverSceneItem {
  kind: "shape" | "text" | "picture";
  id: string;
  /** Page-relative box in EMU. */
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /**
   * Outline paths in a 100 x 100 box stretched over the item (the shared
   * DrawingML path record `{d, fill, stroke}`); absent for text and pictures.
   */
  paths?: readonly CoverScenePath[];
  fill?: CoverScenePaint | null;
  line?: { paint: CoverScenePaint; widthPt: number } | null;
  behind: boolean;
  text?: CoverSceneTextBody | null;
  /**
   * A picture's image (an imported cover's own image part, as a bounded
   * `data:image/png|jpeg|gif|webp` URL). Absent: a picture placeholder.
   */
  src?: string;
  /** `a:srcRect` crop as left, top, right, bottom fractions. */
  crop?: readonly [number, number, number, number];
}

export interface CoverScene {
  version: 1;
  widthEmu: number;
  heightEmu: number;
  /** Back to front. */
  items: readonly CoverSceneItem[];
}

export interface CoverPageGeometry {
  widthEmu: number;
  heightEmu: number;
}

const EMU_PER_PT = 12_700;

function geometryPaths(geom: CoverDefinitionGeometry | undefined): CoverScenePath[] {
  if (geom === undefined || geom === "rect") {
    return [{ d: "M0 0 L100 0 L100 100 L0 100 Z" }];
  }
  if (geom === "ellipse") {
    return [{ d: "M0 50 A50 50 0 1 1 100 50 A50 50 0 1 1 0 50 Z" }];
  }
  if (geom === "line") {
    return [{ d: "M0 0 L100 100", fill: "none" }];
  }
  if (geom === "roundRect") {
    return [
      {
        d: "M8 0 L92 0 Q100 0 100 8 L100 92 Q100 100 92 100 L8 100 Q0 100 0 92 L0 8 Q0 0 8 0 Z",
      },
    ];
  }
  return [{ d: geom.d, ...(geom.open ? { fill: "none" as const } : {}) }];
}

function themeReference(paint: CoverThemePaint): CoverSceneThemeColor {
  return {
    slot: paint.theme,
    ...(paint.lumMod !== undefined ? { lumMod: paint.lumMod } : {}),
    ...(paint.lumOff !== undefined ? { lumOff: paint.lumOff } : {}),
  };
}

function scenePaint(
  paint: CoverDefinitionPaint | undefined,
  theme: CoverThemeColors,
): CoverScenePaint | null {
  if (!paint) return null;
  if ("gradient" in paint) {
    return {
      gradient: {
        angle: paint.gradient.angle,
        stops: paint.gradient.stops.map((stop) => ({
          pos: stop.pos,
          color: resolveCoverThemeColor(stop.paint, theme),
          ...(stop.paint.alpha !== undefined ? { alpha: stop.paint.alpha } : {}),
          theme: themeReference(stop.paint),
        })),
      },
    };
  }
  return {
    color: resolveCoverThemeColor(paint, theme),
    ...(paint.alpha !== undefined ? { alpha: paint.alpha } : {}),
    theme: themeReference(paint),
  };
}

function sceneTextBody(
  body: CoverDefinitionTextBody,
  theme: CoverThemeColors,
  locale: unknown,
): CoverSceneTextBody {
  const [left = 0, top = 0, right = 0, bottom = 0] = body.inset;
  return {
    anchor: body.anchor,
    inset: [
      left * EMU_PER_PT,
      top * EMU_PER_PT,
      right * EMU_PER_PT,
      bottom * EMU_PER_PT,
    ],
    paragraphs: body.paragraphs.map((paragraph) => ({
      align: paragraph.align,
      ...(paragraph.before ? { beforePt: paragraph.before } : {}),
      ...(paragraph.after ? { afterPt: paragraph.after } : {}),
      runs: paragraph.runs.map((run) => ({
        ...(run.field
          ? {
              field: run.field,
              placeholder: coverFieldPlaceholder(run.field, locale),
              ...(run.field === "publishDate"
                ? { dateFormat: coverDatePicture(run.dateFormat, locale) }
                : {}),
            }
          : { text: run.text ?? "" }),
        sizePt: run.size,
        color: resolveCoverThemeColor(run.paint, theme),
        colorTheme: themeReference(run.paint),
        ...(run.bold ? { bold: true } : {}),
        ...(run.italic ? { italic: true } : {}),
        ...(run.caps ? { caps: true } : {}),
        ...(run.spacing ? { spacingPt: run.spacing } : {}),
      })),
    })),
  };
}

/**
 * One MRAI cover drawn on one page: fractions become EMU, theme slots become
 * colours from the document's palette, fields carry their placeholder and
 * date picture in the author's language.
 */
export function coverSceneFromDefinition(
  definition: CoverDefinition,
  page: CoverPageGeometry,
  theme: CoverThemeColors,
  locale: unknown,
): CoverScene {
  const items: CoverSceneItem[] = definition.items.map((item) => {
    const [fx = 0, fy = 0, fw = 0, fh = 0] = item.box;
    const base = {
      id: item.id,
      x: Math.round(fx * page.widthEmu),
      y: Math.round(fy * page.heightEmu),
      w: Math.round(fw * page.widthEmu),
      h: Math.round(fh * page.heightEmu),
      behind: item.behind,
    };
    if (item.type === "picture") {
      return { ...base, kind: "picture" as const };
    }
    if (item.type === "text") {
      return {
        ...base,
        kind: "text" as const,
        text: item.text ? sceneTextBody(item.text, theme, locale) : null,
      };
    }
    return {
      ...base,
      kind: "shape" as const,
      paths: geometryPaths(item.geom),
      fill: scenePaint(item.fill, theme),
      line: item.line
        ? {
            paint: scenePaint(item.line.paint, theme) as CoverScenePaint,
            widthPt: item.line.widthPt,
          }
        : null,
      text: item.text ? sceneTextBody(item.text, theme, locale) : null,
    };
  });
  // Behind-text drawings first, then the rest in authored order.
  const ordered = [
    ...items.filter((item) => item.behind),
    ...items.filter((item) => !item.behind),
  ];
  return {
    version: 1,
    widthEmu: page.widthEmu,
    heightEmu: page.heightEmu,
    items: ordered,
  };
}

// ---------------------------------------------------------------------------
// Imported scenes (from the server projection of a Word cover)
// ---------------------------------------------------------------------------

const MAX_SCENE_ITEMS = 400;
const MAX_SCENE_PARAGRAPHS = 200;
const MAX_SCENE_RUNS = 400;
const MAX_SCENE_TEXT = 4_000;
const MAX_SCENE_PATH = 20_000;
const MAX_EMU = 50_000_000;
const HEX_RE = /^#[0-9a-f]{6}$/;
const PATH_RE = /^[MLCQAZmlcqaz0-9 .,eE+-]*$/;
/** 12 MB of base64: the server's 8 MB picture bound, encoded. */
const MAX_SCENE_PICTURE_CHARS = 12_000_000;
const PICTURE_SRC_RE = /^data:image\/(?:png|jpeg|gif|webp);base64,[A-Za-z0-9+/]+={0,2}$/;

/** A picture source the scene may draw: only an inline raster image. */
export function isCoverPictureSource(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= MAX_SCENE_PICTURE_CHARS &&
    PICTURE_SRC_RE.test(value)
  );
}

function finiteIn(value: unknown, min: number, max: number): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= min && value <= max
    ? value
    : null;
}

function normalizeScenePaint(value: unknown): CoverScenePaint | null {
  if (!value || typeof value !== "object") return null;
  const input = value as Record<string, unknown>;
  if (typeof input.color === "string") {
    const color = input.color.toLowerCase();
    if (!HEX_RE.test(color)) return null;
    const alpha = finiteIn(input.alpha, 0, 1);
    return alpha === null ? { color } : { color, alpha };
  }
  const gradient = input.gradient as Record<string, unknown> | undefined;
  if (!gradient || typeof gradient !== "object") return null;
  const angle = finiteIn(gradient.angle, -360, 360);
  if (angle === null || !Array.isArray(gradient.stops)) return null;
  const stops: Array<{ pos: number; color: string; alpha?: number }> = [];
  for (const raw of gradient.stops.slice(0, 16)) {
    if (!raw || typeof raw !== "object") return null;
    const stop = raw as Record<string, unknown>;
    const pos = finiteIn(stop.pos, 0, 1);
    const color = typeof stop.color === "string" ? stop.color.toLowerCase() : "";
    if (pos === null || !HEX_RE.test(color)) return null;
    const alpha = finiteIn(stop.alpha, 0, 1);
    stops.push(alpha === null ? { pos, color } : { pos, color, alpha });
  }
  return stops.length >= 1 ? { gradient: { angle, stops } } : null;
}

function normalizeSceneText(value: unknown): CoverSceneTextBody | null {
  if (!value || typeof value !== "object") return null;
  const input = value as Record<string, unknown>;
  const anchor =
    input.anchor === "ctr" || input.anchor === "b" ? input.anchor : "t";
  const inset = Array.isArray(input.inset) ? input.inset : [];
  const insets = [0, 1, 2, 3].map(
    (index) => finiteIn(inset[index], 0, MAX_EMU) ?? 0,
  ) as [number, number, number, number];
  if (!Array.isArray(input.paragraphs)) return null;
  let runCount = 0;
  const paragraphs: CoverSceneParagraph[] = [];
  for (const raw of input.paragraphs.slice(0, MAX_SCENE_PARAGRAPHS)) {
    if (!raw || typeof raw !== "object") continue;
    const paragraph = raw as Record<string, unknown>;
    const align =
      paragraph.align === "center" ||
      paragraph.align === "right" ||
      paragraph.align === "justify"
        ? paragraph.align
        : "left";
    const runs: CoverSceneRun[] = [];
    for (const rawRun of Array.isArray(paragraph.runs) ? paragraph.runs : []) {
      if (runCount >= MAX_SCENE_RUNS || !rawRun || typeof rawRun !== "object") {
        continue;
      }
      const run = rawRun as Record<string, unknown>;
      const color =
        typeof run.color === "string" && HEX_RE.test(run.color.toLowerCase())
          ? run.color.toLowerCase()
          : "#000000";
      const sizePt = finiteIn(run.sizePt, 1, 400) ?? 11;
      const field = isCoverFieldKey(run.field) ? run.field : null;
      const text =
        typeof run.text === "string" ? run.text.slice(0, MAX_SCENE_TEXT) : "";
      if (!field && !text) continue;
      runCount += 1;
      runs.push({
        ...(field ? { field } : { text }),
        ...(field && typeof run.placeholder === "string"
          ? { placeholder: run.placeholder.slice(0, MAX_SCENE_TEXT) }
          : {}),
        ...(field && typeof run.dateFormat === "string" && run.dateFormat.length <= 64
          ? { dateFormat: run.dateFormat }
          : {}),
        sizePt,
        color,
        ...(run.bold === true ? { bold: true } : {}),
        ...(run.italic === true ? { italic: true } : {}),
        ...(run.caps === true ? { caps: true } : {}),
        ...(finiteIn(run.spacingPt, -50, 200) ? { spacingPt: run.spacingPt as number } : {}),
        ...(typeof run.fontFamily === "string" && run.fontFamily.length <= 80 && !/[<>;{}]/.test(run.fontFamily)
          ? { fontFamily: run.fontFamily }
          : {}),
      });
    }
    paragraphs.push({
      align,
      ...(finiteIn(paragraph.beforePt, 0, 1_584) ? { beforePt: paragraph.beforePt as number } : {}),
      ...(finiteIn(paragraph.afterPt, 0, 1_584) ? { afterPt: paragraph.afterPt as number } : {}),
      ...(finiteIn(paragraph.lineHeight, 0.5, 5) ? { lineHeight: paragraph.lineHeight as number } : {}),
      runs,
    });
  }
  return { anchor, inset: insets, paragraphs };
}

/**
 * A server-projected scene, bounded and closed: only numbers, hex colours,
 * path data in the SVG path alphabet and plain text survive.
 */
export function normalizeCoverScene(value: unknown): CoverScene | null {
  if (!value || typeof value !== "object") return null;
  const input = value as Record<string, unknown>;
  const widthEmu = finiteIn(input.widthEmu, 1, MAX_EMU);
  const heightEmu = finiteIn(input.heightEmu, 1, MAX_EMU);
  if (widthEmu === null || heightEmu === null || !Array.isArray(input.items)) {
    return null;
  }
  const items: CoverSceneItem[] = [];
  for (const raw of input.items.slice(0, MAX_SCENE_ITEMS)) {
    if (!raw || typeof raw !== "object") continue;
    const item = raw as Record<string, unknown>;
    const kind =
      item.kind === "text" || item.kind === "picture" ? item.kind : "shape";
    const x = finiteIn(item.x, -MAX_EMU, MAX_EMU);
    const y = finiteIn(item.y, -MAX_EMU, MAX_EMU);
    const w = finiteIn(item.w, 0, MAX_EMU);
    const h = finiteIn(item.h, 0, MAX_EMU);
    if (x === null || y === null || w === null || h === null) continue;
    const paths: CoverScenePath[] = [];
    if (kind !== "text" && Array.isArray(item.paths)) {
      for (const rawPath of item.paths.slice(0, 64)) {
        const path = rawPath as Record<string, unknown> | null;
        if (
          !path ||
          typeof path.d !== "string" ||
          path.d.length > MAX_SCENE_PATH ||
          !PATH_RE.test(path.d)
        ) {
          continue;
        }
        paths.push({
          d: path.d,
          ...(path.fill === "none" ? { fill: "none" as const } : {}),
          ...(path.stroke === false ? { stroke: false } : {}),
        });
      }
    }
    const linePaint = item.line && typeof item.line === "object"
      ? normalizeScenePaint((item.line as Record<string, unknown>).paint)
      : null;
    const lineWidth = item.line && typeof item.line === "object"
      ? finiteIn((item.line as Record<string, unknown>).widthPt, 0, 200)
      : null;
    items.push({
      kind,
      id: typeof item.id === "string" ? item.id.slice(0, 80) : `item-${items.length}`,
      x,
      y,
      w,
      h,
      ...(finiteIn(item.rotation, -360, 360) ? { rotation: item.rotation as number } : {}),
      ...(item.flipH === true ? { flipH: true } : {}),
      ...(item.flipV === true ? { flipV: true } : {}),
      ...(kind === "shape" || (kind === "picture" && paths.length) ? { paths } : {}),
      ...(kind === "picture" && isCoverPictureSource(item.src) ? { src: item.src } : {}),
      ...(kind === "picture" &&
      Array.isArray(item.crop) &&
      item.crop.length === 4 &&
      item.crop.every((edge) => finiteIn(edge, -1, 0.99) !== null)
        ? { crop: item.crop as unknown as [number, number, number, number] }
        : {}),
      fill: normalizeScenePaint(item.fill),
      line: linePaint && lineWidth !== null ? { paint: linePaint, widthPt: lineWidth } : null,
      behind: item.behind === true,
      text: normalizeSceneText(item.text),
    });
  }
  return { version: 1, widthEmu, heightEmu, items };
}

/** Every field a scene shows, in first-seen order. */
export function coverSceneFields(scene: CoverScene): CoverFieldKey[] {
  const fields: CoverFieldKey[] = [];
  for (const item of scene.items) {
    for (const paragraph of item.text?.paragraphs ?? []) {
      for (const run of paragraph.runs) {
        if (run.field && !fields.includes(run.field)) fields.push(run.field);
      }
    }
  }
  return fields;
}
