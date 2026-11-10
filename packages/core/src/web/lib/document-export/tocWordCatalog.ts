/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import catalog from "./tocWordCatalog.json";
import {
  TOC_STORED_FORMATS,
  type TocFormat,
  type TocTabLeader,
} from "../../components/console/editor/native-document/referenceSettings";

/**
 * A UI locale the catalogue carries words for ("zh", "en", "de", "zh-TW", …);
 * `tocCatalogLocale` maps any locale onto one, English last.
 */
export type TocCatalogLocale = string;

export interface TocGalleryEntryPlaceholder {
  level: number;
  text: string;
  page: number;
}

export interface TocGalleryRow {
  id: string;
  name: Record<TocCatalogLocale, string>;
  format: TocFormat;
}

interface TocCatalogShape {
  title: Record<TocCatalogLocale, string>;
  noEntries: Record<TocCatalogLocale, string>;
  formats: Record<
    string,
    {
      defaultTabLeader: TocTabLeader;
      defaultRightAlignPageNumbers: boolean;
    }
  >;
  gallery: {
    automatic: TocGalleryRow[];
    manual: (TocGalleryRow & {
      entries: Record<TocCatalogLocale, TocGalleryEntryPlaceholder[]>;
    })[];
  };
}

const CATALOG = catalog as unknown as TocCatalogShape;

export const TOC_STYLE_IDS = [
  "TOC1",
  "TOC2",
  "TOC3",
  "TOC4",
  "TOC5",
  "TOC6",
  "TOC7",
  "TOC8",
  "TOC9",
] as const;

export const TOC_HEADING_STYLE_ID = "TOCHeading";

export function tocCatalogLocale(locale: string | null | undefined): TocCatalogLocale {
  if (typeof locale !== "string" || !locale) return "en";
  const known = Object.keys(CATALOG.title);
  const exact = known.find((tag) => tag.toLowerCase() === locale.toLowerCase());
  if (exact) return exact;
  // zh-Hant/zh-HK read Traditional; every other regional tag its language.
  if (/^zh-(hant|hk|mo|tw)/i.test(locale) && known.includes("zh-TW")) return "zh-TW";
  const base = locale.split(/[-_]/)[0].toLowerCase();
  return known.find((tag) => tag.toLowerCase() === base) ?? "en";
}

function wordsFor(words: Record<TocCatalogLocale, string>, locale: TocCatalogLocale): string {
  return words[locale] ?? words.en;
}

/** The gallery block's heading text ("目录" / "Contents"). */
export function tocGalleryTitle(locale: TocCatalogLocale): string {
  return wordsFor(CATALOG.title, locale);
}

/** Word's cached result of a TOC that found nothing. */
export function tocNoEntriesText(locale: TocCatalogLocale): string {
  return wordsFor(CATALOG.noEntries, locale);
}

/** A gallery row's name in the reader's language. */
export function tocGalleryRowName(row: TocGalleryRow, locale: TocCatalogLocale): string {
  return wordsFor(row.name, locale);
}

export function tocGalleryAutomaticRows(): readonly TocGalleryRow[] {
  return CATALOG.gallery.automatic;
}

export function tocGalleryManualRows(): readonly TocGalleryRow[] {
  return CATALOG.gallery.manual;
}

/** The placeholder entries of a manual gallery block, in document order. */
export function tocGalleryManualEntries(
  id: string,
  locale: TocCatalogLocale,
): TocGalleryEntryPlaceholder[] {
  const row = CATALOG.gallery.manual.find((entry) => entry.id === id);
  return row ? (row.entries[locale] ?? row.entries.en).map((entry) => ({ ...entry })) : [];
}

/** The leader and alignment Word's dialog preselects for a format. */
export function tocFormatDefaults(format: TocFormat): {
  tabLeader: TocTabLeader;
  rightAlignPageNumbers: boolean;
} {
  const entry = CATALOG.formats[format] ?? CATALOG.formats.template;
  return {
    tabLeader: entry.defaultTabLeader,
    rightAlignPageNumbers: entry.defaultRightAlignPageNumbers,
  };
}

export function isTocFormat(value: unknown): value is TocFormat {
  return typeof value === "string" && (TOC_STORED_FORMATS as readonly string[]).includes(value);
}
