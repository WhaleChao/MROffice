/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The product's locale roster — ONE source of truth for routing, the request
 * config, the language picker, `<html lang/dir>` and every Intl formatter.
 *
 * Resolution contract (i18n/request.ts): a locale's catalogue never has to be
 * complete. Messages resolve along a chain that always ENDS at `BASE_LOCALE`
 * (en), so a key the translators have not reached yet renders in English
 * rather than as a raw `namespace.key`. Regional variants chain through their
 * parent first (es-419 → es → en, en-GB → en, zh-TW → zh → en), which is what
 * makes a regional catalogue able to carry only its genuine deltas
 * ("colour" vs "color") instead of a full copy.
 */

export const LOCALES = [
  "en",
  "id",
  "da",
  "de",
  "en-GB",
  "es",
  "es-419",
  "fr",
  "it",
  "nl",
  "no",
  "pt-BR",
  "fi",
  "sv",
  "vi",
  "th",
  "ko",
  "ja",
  "zh",
  "zh-TW",
] as const;

/**
 * Hebrew ("he") and Arabic ("ar") are shelved for now: they are the roster's
 * only RTL locales, and shipping them means auditing every editor surface
 * under `dir="rtl"` first. Their translated catalogues live in git history
 * (`0928d7172`, messages/he + messages/ar) — restoring them is: re-add the
 * tuple entries + LOCALE_META rows here and in locale-list.json, restore the
 * catalogue directories, regenerate message-manifest.json. The `dir`
 * machinery (localeDirection, `<html dir>`) stays in place for that day.
 */

export type Locale = (typeof LOCALES)[number];

/** The workspace's original locale; unchanged so existing URLs keep working. */
export const DEFAULT_LOCALE: Locale = "zh";

/** The always-complete catalogue every chain terminates at. */
export const BASE_LOCALE: Locale = "en";

export interface LocaleMeta {
  /** Endonym — the picker always shows a language in that language. */
  native: string;
  /** English name, for aria/labels and the settings search index. */
  english: string;
  dir: "ltr" | "rtl";
  /**
   * Next hop when a key is missing. Undefined means "straight to
   * BASE_LOCALE"; BASE_LOCALE itself terminates the chain.
   */
  fallback?: Locale;
}

export const LOCALE_META: Record<Locale, LocaleMeta> = {
  en: { native: "English (US)", english: "English (US)", dir: "ltr" },
  "en-GB": {
    native: "English (UK)",
    english: "English (UK)",
    dir: "ltr",
    fallback: "en",
  },
  id: { native: "Bahasa Indonesia", english: "Indonesian", dir: "ltr" },
  da: { native: "Dansk", english: "Danish", dir: "ltr" },
  de: { native: "Deutsch", english: "German", dir: "ltr" },
  es: { native: "Español (España)", english: "Spanish (Spain)", dir: "ltr" },
  "es-419": {
    native: "Español (Latinoamérica)",
    english: "Spanish (Latin America)",
    dir: "ltr",
    fallback: "es",
  },
  fr: { native: "Français (France)", english: "French", dir: "ltr" },
  it: { native: "Italiano", english: "Italian", dir: "ltr" },
  nl: { native: "Nederlands", english: "Dutch", dir: "ltr" },
  no: { native: "Norsk", english: "Norwegian", dir: "ltr" },
  "pt-BR": {
    native: "Português (Brasil)",
    english: "Portuguese (Brazil)",
    dir: "ltr",
  },
  fi: { native: "Suomi", english: "Finnish", dir: "ltr" },
  sv: { native: "Svenska", english: "Swedish", dir: "ltr" },
  vi: { native: "Tiếng Việt", english: "Vietnamese", dir: "ltr" },
  th: { native: "ภาษาไทย", english: "Thai", dir: "ltr" },
  ko: { native: "한국어", english: "Korean", dir: "ltr" },
  ja: { native: "日本語", english: "Japanese", dir: "ltr" },
  zh: { native: "简体中文", english: "Chinese (Simplified)", dir: "ltr" },
  "zh-TW": {
    native: "繁體中文",
    english: "Chinese (Traditional)",
    dir: "ltr",
    fallback: "zh",
  },
};

export function isLocale(value: unknown): value is Locale {
  return typeof value === "string" && (LOCALES as readonly string[]).includes(value);
}

export function localeDirection(locale: string): "ltr" | "rtl" {
  return isLocale(locale) ? LOCALE_META[locale].dir : "ltr";
}

/**
 * Catalogue merge order for a locale, LEAST specific first, so a plain
 * `Object.assign`-style deep merge down the list leaves the most specific
 * translation on top: `["en", "es", "es-419"]` for es-419.
 */
export function messageChain(locale: Locale): Locale[] {
  const chain: Locale[] = [];
  let cursor: Locale | undefined = locale;
  const guard = new Set<Locale>();
  while (cursor && !guard.has(cursor)) {
    guard.add(cursor);
    chain.push(cursor);
    cursor = LOCALE_META[cursor].fallback;
  }
  if (!guard.has(BASE_LOCALE)) chain.push(BASE_LOCALE);
  return chain.reverse();
}

/**
 * Marketing stages and other two-voice surfaces carry hand-written mock copy
 * in Chinese and English only. They ask for a CONTENT locale: every Chinese
 * variant keeps the Chinese script, everything else reads the English copy.
 */
export function contentLocale(locale: string): "zh" | "en" {
  return locale === "zh" || locale === "zh-TW" ? "zh" : "en";
}

/**
 * Picker order: the reference product lists the current language first, then
 * sorts endonyms with a Latin-first collator (Latin scripts, then Thai,
 * Hangul, Kana, Han).
 */
export function localesForPicker(current: Locale): Locale[] {
  const collator = new Intl.Collator("en");
  const rest = LOCALES.filter((locale) => locale !== current).sort((a, b) =>
    collator.compare(LOCALE_META[a].native, LOCALE_META[b].native),
  );
  return [current, ...rest];
}

/**
 * Drop a leading roster locale segment: "/de/app/x" → "/app/x", "/zh-TW" → "/".
 *
 * Every client-side "is this the notebooks page?" check used to compare
 * against a hand-written `/en/...` + `/zh/...` pair, and the one helper that
 * did strip a prefix used `slice(3)` — which silently mangles the four- and
 * six-character tags in this roster (`zh-TW`, `es-419`). Route comparisons
 * should normalise through here and then test the unprefixed path once.
 */
export function stripLocalePrefix(pathname: string): string {
  const first = pathname.split("/")[1];
  return isLocale(first) ? pathname.slice(first.length + 1) || "/" : pathname;
}

/**
 * "/<locale>" for the locale the current page is in, "" for the default
 * (served unprefixed) or on the server. For client code that builds an
 * absolute path to land on after a redirect.
 */
export function visitorLocalePrefix(): string {
  if (typeof window === "undefined") return "";
  const first = window.location.pathname.split("/")[1];
  return isLocale(first) && first !== DEFAULT_LOCALE ? `/${first}` : "";
}

/**
 * The locale the interface is shown in (the root layout's `<html lang>`),
 * or null on the server. Sent with every AI request as `interface_locale`:
 * the answer's language when the question itself gives no clear signal.
 */
export function interfaceLocale(): Locale | null {
  if (typeof document === "undefined") return null;
  const lang = document.documentElement.lang;
  return isLocale(lang) ? lang : null;
}

/** `body` with the interface locale added (unchanged when none is known). */
export function withInterfaceLocale<T extends object>(
  body: T,
): T & { interface_locale?: Locale } {
  const locale = interfaceLocale();
  return locale ? { ...body, interface_locale: locale } : body;
}
