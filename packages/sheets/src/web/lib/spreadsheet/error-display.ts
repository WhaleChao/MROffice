/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The text an error value is SHOWN with. Excel localises one of them: a
 * Simplified Chinese Excel shows #SPILL! as #溢出! (its own string table, and
 * MEASURED in Excel for Mac zh-CN; XLS-A089) and every other error -- #VALUE!,
 * #CALC!, #FIELD! -- in English. The value itself, formulas, the display
 * cache and exports keep the code.
 */
const SIMPLIFIED_CHINESE_ERRORS: Readonly<Record<string, string>> = {
  "#SPILL!": "#溢出!",
};

/** Whether `locale` is a Simplified Chinese UI (zh, zh-CN, zh-Hans...). */
function simplifiedChinese(locale: string): boolean {
  const tag = locale.toLowerCase();
  return tag.startsWith("zh") && !/-(?:tw|hk|mo|hant)\b/.test(tag);
}

export function localizedErrorDisplay(display: string, locale: string): string {
  if (display.charCodeAt(0) !== 35 /* # */ || !simplifiedChinese(locale)) return display;
  return SIMPLIFIED_CHINESE_ERRORS[display] ?? display;
}

/** The error code a shown text stands for (#溢出! is #SPILL!). */
export function errorCodeOfDisplay(display: string): string {
  for (const [code, shown] of Object.entries(SIMPLIFIED_CHINESE_ERRORS)) {
    if (shown === display) return code;
  }
  return display;
}
