/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The title of the outline item MRAI keeps a PDF's own bookmarks under
 * (the Bookmarks pane's managed container), in the language of the reader
 * who first adds one: "MRAI Bookmarks" in English, "MRAI 书签" in Chinese.
 * It is written once, when the container is created; a later edit keeps the
 * title the file's container already has, whoever makes it, and every
 * language's title is recognised as that container.
 *
 * The API's `app/core/content_words.json` (`pdfManagedBookmarks`) holds the
 * same words for the server's own outline writer (a test keeps them equal).
 */

import { contentLocaleOf, type ContentLocale } from "@mrnote/core/web/lib/i18n/content-locale.ts";

const TITLES: Readonly<Record<ContentLocale, string>> = {
  en: "MRAI Bookmarks",
  zh: "MRAI 书签",
  "zh-TW": "MRAI 書籤",
  ja: "MRAI しおり",
  ko: "MRAI 책갈피",
  de: "MRAI-Lesezeichen",
  fr: "Signets MRAI",
  es: "Marcadores de MRAI",
  it: "Segnalibri MRAI",
  nl: "MRAI-bladwijzers",
  "pt-BR": "Marcadores do MRAI",
  da: "MRAI-bogmærker",
  no: "MRAI-bokmerker",
  sv: "MRAI-bokmärken",
  fi: "MRAI-kirjanmerkit",
  id: "Marka buku MRAI",
  th: "บุ๊กมาร์ก MRAI",
  vi: "Dấu trang MRAI",
};

/** Every language's container title. */
export const PDF_MANAGED_BOOKMARKS_TITLES: ReadonlySet<string> = new Set(Object.values(TITLES));

/** A new container's title for a reader of `locale` (no locale: English). */
export function pdfManagedBookmarksTitle(locale: string | null | undefined): string {
  return TITLES[contentLocaleOf(locale, "en")];
}

/** Whether `title` is MRAI's container title in any language. */
export function isPdfManagedBookmarksTitle(title: string | null | undefined): boolean {
  return typeof title === "string" && PDF_MANAGED_BOOKMARKS_TITLES.has(title);
}
