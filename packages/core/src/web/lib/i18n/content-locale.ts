/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The language a piece of inserted CONTENT is written in -- an inserted
 * chart's sample data, a drawing template's labels, a starter diagram -- from
 * the viewer's interface locale. Content tables hold one entry per language
 * (not per regional variant): "en-GB" writes "en", "es-419" "es", "zh-Hant"
 * "zh-TW", Norwegian Bokmål / Nynorsk "no"; a language a table does not hold
 * writes English, as the message catalogues fall back (`i18n/locales.ts`).
 *
 * The words are chosen once, when the content goes in, and are never
 * re-translated: once inserted they are the document's own text.
 */

export const CONTENT_LOCALES = [
  "en",
  "zh",
  "zh-TW",
  "ja",
  "ko",
  "de",
  "fr",
  "es",
  "it",
  "nl",
  "pt-BR",
  "da",
  "no",
  "sv",
  "fi",
  "id",
  "th",
  "vi",
] as const;

export type ContentLocale = (typeof CONTENT_LOCALES)[number];

const CONTENT_LOCALE_SET = new Set<string>(CONTENT_LOCALES);

/**
 * The interface language of the page this code runs in (`<html lang>`, which
 * the locale layout sets from the route), for content written outside the
 * React tree -- an AI command, a Copilot structure builder. Undefined on the
 * server and in a document with no language.
 */
export function pageInterfaceLocale(): string | undefined {
  if (typeof document === "undefined") return undefined;
  return document.documentElement.lang || undefined;
}

/**
 * The request header an API request carries this page's content language in
 * (`lib/api.ts`), so the words the SERVER writes into content -- a repaired
 * chart's "Category 1", an empty deck's "Untitled" in its PDF, a
 * PivotTable's "Grand Total" -- are the reader's too. The API reduces it to
 * one of the same content languages (`app/core/content_locale.py`). A route
 * reached through the file-transfer host takes it as the `content_locale`
 * query parameter instead (`CONTENT_LOCALE_QUERY`): that host's preflight
 * allows a fixed header list.
 */
export const CONTENT_LOCALE_HEADER = "X-Content-Locale";
export const CONTENT_LOCALE_QUERY = "content_locale";

/**
 * The content language of the page this code runs in, for an API request;
 * undefined on the server and in a document with no language (the API then
 * reads the browser's Accept-Language, or keeps its own default).
 */
export function requestContentLocale(): ContentLocale | undefined {
  const locale = pageInterfaceLocale();
  return locale ? contentLocaleOf(locale) : undefined;
}

/**
 * `locale`'s content language; `fallback` when no locale is given at all (a
 * caller that cannot know the reader decides what it writes then).
 */
export function contentLocaleOf(
  locale: string | null | undefined,
  fallback: ContentLocale = "en",
): ContentLocale {
  if (typeof locale !== "string" || !locale.trim()) return fallback;
  const lower = locale.trim().replace(/_/g, "-").toLowerCase();
  if (/^zh-(hant|tw|hk|mo)\b/.test(lower)) return "zh-TW";
  if (lower === "zh" || lower.startsWith("zh-")) return "zh";
  if (lower === "pt" || lower.startsWith("pt-")) return "pt-BR";
  const base = lower.split("-")[0];
  if (base === "nb" || base === "nn") return "no";
  return CONTENT_LOCALE_SET.has(base) ? (base as ContentLocale) : "en";
}
