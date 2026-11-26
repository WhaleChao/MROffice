/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one language the app speaks in: its pages' locale prefix
 * (./classic/locale.ts) and every word the shell itself shows -- menus, the
 * context menu, dialogs, the launch window, the update prompts -- read here,
 * never from `app.getLocale()` on their own (GitHub XinaTian/MROffice#1:
 * English chosen on a Chinese system kept the app's own words Chinese).
 *
 * The desktop apps set it (`setAppLocale`) from the person's choice, the
 * site's language cookie; before they have, and in an app that never does,
 * it is the computer's preferred languages read with the web app's rules
 * (./classic/web-locales.ts pickLocale), then the app's own locale, then the
 * web app's default.
 */

import { app } from "electron";

import { DEFAULT_LOCALE, LOCALES, pickLocale, type Locale } from "./classic/web-locales";

let chosen: Locale | null = null;
const listeners = new Set<(locale: Locale) => void>();

/** The first of these language tags the site is written in, or null. */
export function localeOfTags(tags: readonly (string | null | undefined)[]): Locale | null {
  for (const tag of tags) {
    const found = pickLocale(tag, LOCALES);
    if (found) return found;
  }
  return null;
}

/** The computer's languages, most preferred first, then the app's own
 * locale; a locale the app was started with (`--lang=en-US`) first. */
export function systemLanguageTags(): string[] {
  let preferred: string[] = [];
  try {
    preferred = app.getPreferredSystemLanguages();
  } catch {
    preferred = [];
  }
  let own = "";
  try {
    own = app.getLocale();
  } catch {
    own = "";
  }
  let asked = false;
  try {
    asked = app.commandLine.hasSwitch("lang");
  } catch {
    asked = false;
  }
  return (asked ? [own, ...preferred] : [...preferred, own]).filter(Boolean);
}

/** The language the app speaks in now. */
export function appLocale(): Locale {
  return chosen ?? localeOfTags(systemLanguageTags()) ?? DEFAULT_LOCALE;
}

/** Whether the app's own words are Chinese (the shell's texts are Chinese or English). */
export function appSpeaksChinese(): boolean {
  return appLocale().toLowerCase().startsWith("zh");
}

/** The app's language, as the person chose it (or the launch worked it out). */
export function setAppLocale(locale: Locale): void {
  if (chosen === locale) return;
  chosen = locale;
  for (const listener of listeners) {
    try {
      listener(locale);
    } catch (error) {
      console.warn(`[desktop] language change not applied: ${error instanceof Error ? error.message : error}`);
    }
  }
}

/** Told when the app's language changes (menus rebuilt in it). */
export function onAppLocaleChange(listener: (locale: Locale) => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}
