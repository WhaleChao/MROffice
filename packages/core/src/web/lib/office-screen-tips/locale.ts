/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The languages Office's ScreenTip sentences ship in, and which of them serves
 * a UI locale. Shared by the Spreadsheet's tables (XLS-A106) and the
 * Presentation/Docs tables: a regional or script variant reads its language's
 * module (en-GB → en, es-419 → es, zh-CN → zh, zh-Hant-HK → zh-TW, nb → no,
 * pt-PT → pt-BR: i18n/locale-tags), anything else English.
 */

import { pickLocale } from "../../i18n/locale-tags";

export const OFFICE_SCREEN_TIP_LOCALES = [
  "da", "de", "en", "es", "fi", "fr", "id", "it", "ja", "ko", "nl", "no",
  "pt-BR", "sv", "th", "vi", "zh", "zh-TW", "ar", "ru", "tr", "uk", "cs", "pt-PT",
  "ro", "sr-Cyrl", "sr-Latn", "el", "hu", "bg", "hr", "lv", "sk", "ca", "ms", "sl",
  "sq", "az", "gl", "hy", "ur", "be", "eu", "si", "lo", "he", "pl",
] as const;
export type OfficeScreenTipLocale = (typeof OFFICE_SCREEN_TIP_LOCALES)[number];

export function officeScreenTipLocale(locale: string): OfficeScreenTipLocale {
  return pickLocale(locale, OFFICE_SCREEN_TIP_LOCALES) ?? "en";
}
