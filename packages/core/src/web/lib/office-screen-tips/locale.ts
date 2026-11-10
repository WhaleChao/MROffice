/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The languages Office's ScreenTip sentences ship in, and which of them serves
 * a UI locale. Shared by the Spreadsheet's tables (XLS-A106) and the
 * Presentation/Docs tables: a regional or script variant reads its language's
 * module (en-GB → en, es-419 → es, zh-CN → zh, zh-Hant-HK → zh-TW, nb → no),
 * anything else English.
 */

export const OFFICE_SCREEN_TIP_LOCALES = [
  "da", "de", "en", "es", "fi", "fr", "id", "it", "ja", "ko", "nl", "no",
  "pt-BR", "sv", "th", "vi", "zh", "zh-TW",
] as const;
export type OfficeScreenTipLocale = (typeof OFFICE_SCREEN_TIP_LOCALES)[number];

export function officeScreenTipLocale(locale: string): OfficeScreenTipLocale {
  const tag = locale.trim();
  const exact = OFFICE_SCREEN_TIP_LOCALES.find((entry) => entry.toLowerCase() === tag.toLowerCase());
  if (exact) return exact;
  const lower = tag.toLowerCase();
  if (lower === "zh-hant" || lower.startsWith("zh-hant-") || lower === "zh-hk" || lower === "zh-mo") return "zh-TW";
  if (lower === "nb" || lower === "nn" || lower.startsWith("nb-") || lower.startsWith("nn-")) return "no";
  if (lower.startsWith("pt")) return "pt-BR";
  const base = lower.split("-")[0];
  return OFFICE_SCREEN_TIP_LOCALES.find((entry) => entry === base) ?? "en";
}
