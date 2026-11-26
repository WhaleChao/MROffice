/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The language the app's pages are shown in (GitHub XinaTian/MROffice#1:
 * English chosen in Settings, the app opened again in Chinese).
 *
 * The web app serves its default locale (zh) without a prefix and every
 * other one under its own (`/en/app/recent`); it never chooses a language
 * for an unprefixed workspace page (apps/web proxy.ts firstEntryLocale: the
 * marketing and sign-in pages only, and only a visitor's first one). The
 * shell loaded every page unprefixed, so each launch, each new window and
 * each document opened from the system showed Chinese, whatever the person
 * chose, and their computer's language was never asked.
 *
 * Now every route the shell loads carries the app's locale
 * (`localizedRoute`). That locale is, in this order:
 *   1. the person's choice: the site's `mr_locale` cookie in the app's
 *      session, which the language picker writes (apps/web
 *      components/i18n/LocalePicker) -- never what a window happens to show
 *      (a link someone shared, in their language, is not a choice); the
 *      shell keeps a copy in desktop-settings.json, read before the session
 *      is, and follows the cookie as it changes;
 *   2. the computer's preferred languages, read as the site reads a
 *      browser's (apps/web/i18n/locale-tags.ts pickLocale, ./web-locales.ts);
 *   3. the web app's default, zh, as before.
 * No Electron here: the rules are tested on their own (test/locale.test.mjs).
 */

import { DEFAULT_LOCALE, LOCALES, pickLocale, type Locale } from "./web-locales";

export { DEFAULT_LOCALE, LOCALES, type Locale };

/** The site's cookie for a remembered language (apps/web/i18n/detect-locale.ts LOCALE_CHOSEN_COOKIE). */
export const LOCALE_CHOSEN_COOKIE = "mr_locale";

const ROSTER: ReadonlySet<string> = new Set(LOCALES);
const escaped = LOCALES.map((locale) => locale.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
/** A path's leading roster locale ("/pt-BR/app" → "/pt-BR"), as the web app's proxy matches it (exact spelling). */
export const LOCALE_PREFIX = new RegExp(`^/(?:${escaped.join("|")})(?=/|$)`);
/** Paths that are not pages: Next's files, the API, the print pages, static files. */
export const NOT_A_PAGE = /^\/(?:_next|api|render|pyodide|fonts|proofing|marketing)(?:\/|$)|\.[A-Za-z0-9]{1,8}$/;

export function isLocale(value: unknown): value is Locale {
  return typeof value === "string" && ROSTER.has(value);
}

/** The roster locale a path starts with, or null. */
export function localeOfPath(pathname: string): Locale | null {
  const found = LOCALE_PREFIX.exec(pathname)?.[0].slice(1);
  return isLocale(found) ? found : null;
}

/** The path without its locale ("/en/app/recent" → "/app/recent", "/en" → "/"). */
export function stripLocale(pathname: string): string {
  const found = LOCALE_PREFIX.exec(pathname)?.[0] ?? "";
  return pathname.slice(found.length) || "/";
}

/**
 * The locale a page of the app's origin is shown in: its prefix, or the
 * default for an unprefixed page. Null for anything else (another origin,
 * not a page, not a URL).
 */
export function shownLocale(raw: string, webOrigin: string): Locale | null {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return null;
  }
  if (url.origin !== webOrigin) return null;
  if (NOT_A_PAGE.test(stripLocale(url.pathname))) return null;
  return localeOfPath(url.pathname) ?? DEFAULT_LOCALE;
}

/** The first of the computer's preferred languages (most preferred first) the site is written in, or null. */
export function systemLocale(languages: readonly string[]): Locale | null {
  for (const language of languages) {
    const found = pickLocale(language, LOCALES);
    if (found) return found;
  }
  return null;
}

/** The locale a launch shows its pages in (see the top of this file). */
export function launchLocale(known: { chosen?: unknown; system?: readonly string[] }): Locale {
  if (isLocale(known.chosen)) return known.chosen;
  return systemLocale(known.system ?? []) ?? DEFAULT_LOCALE;
}

/**
 * The locale a change of the app session's cookies chose: the site's
 * language cookie set (not removed, not expired) for the app's host, with a
 * roster locale. Null for any other change.
 */
export function chosenByCookieChange(
  cookie: { name: string; value: string; domain?: string },
  removed: boolean,
  webHost: string,
): Locale | null {
  if (removed || cookie.name !== LOCALE_CHOSEN_COOKIE) return null;
  const domain = (cookie.domain ?? "").replace(/^\./, "");
  if (domain && domain !== webHost) return null;
  let value = cookie.value;
  try {
    value = decodeURIComponent(value);
  } catch {
    return null;
  }
  return isLocale(value) ? value : null;
}

/**
 * ``route`` ("/app/recent?surface=office") in ``locale``: prefixed, unless
 * it is the default (served unprefixed: the site sends `/zh/...` there) or
 * the route names a locale already (a page's own link, opened in a new
 * window), or it is not a page.
 */
export function localizedRoute(route: string, locale: Locale): string {
  if (locale === DEFAULT_LOCALE || !route.startsWith("/") || route.startsWith("//")) return route;
  const url = new URL(route, "http://app.invalid");
  if (localeOfPath(url.pathname) || NOT_A_PAGE.test(url.pathname)) return route;
  const pathname = url.pathname === "/" ? `/${locale}` : `/${locale}${url.pathname}`;
  return `${pathname}${url.search}${url.hash}`;
}
