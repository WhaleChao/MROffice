/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Excel Home > Number format gallery. Pure data keeps UI and tests aligned. */

import {
  normalizeNumberFormatCode,
  numberFormatFamily,
} from "./number-format";
import { buildAccountingCodeFromParams } from "./format-code-builder";

export interface NumberFormatGalleryItem {
  id: string;
  labelKey:
    | "spreadsheet.fmtCatGeneral"
    | "spreadsheet.fmtCatNumber"
    | "spreadsheet.fmtCatCurrency"
    | "spreadsheet.fmtCatAccounting"
    | "spreadsheet.fmtCatDate"
    | "spreadsheet.fmtCatTime"
    | "spreadsheet.fmtCatPercentage"
    | "spreadsheet.fmtCatFraction"
    | "spreadsheet.fmtCatScientific"
    | "spreadsheet.fmtCatText"
    | "spreadsheet.fmtCatCustom";
  menuIcon: NumberFormatMenuIconId;
  code: string;
}

export type NumberFormatMenuIconId =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "date"
  | "long-date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text";

export const ACCOUNTING_USD_FORMAT = buildAccountingCodeFromParams("$", 2);
export const ACCOUNTING_EUR_FORMAT = buildAccountingCodeFromParams("€", 2);
export const ACCOUNTING_GBP_FORMAT = buildAccountingCodeFromParams("£", 2);
// zh-CN Excel's own Accounting (built-in 44 there), not the en-US shape with
// a yuan sign (XLS-A053 / XLS-A105).
export const ACCOUNTING_CNY_FORMAT = buildAccountingCodeFromParams("¥", 2, "zh");

// Excel's fifth entry, "CHF French (Switzerland)" (XLS-A008).
export const ACCOUNTING_CHF_FORMAT = buildAccountingCodeFromParams("CHF", 2);

// Appended to, never reordered: commands address the entries by position.
export const ACCOUNTING_CURRENCY_FORMATS = [
  ["accountingUsd", ACCOUNTING_USD_FORMAT],
  ["accountingEur", ACCOUNTING_EUR_FORMAT],
  ["accountingGbp", ACCOUNTING_GBP_FORMAT],
  ["accountingCny", ACCOUNTING_CNY_FORMAT],
  ["accountingChf", ACCOUNTING_CHF_FORMAT],
] as const;

/**
 * The Accounting Number Format list in Excel's order: the UI language's own
 * currency first (¥ in a Chinese Excel, $ elsewhere), CHF last (XLS-A008).
 */
export function accountingCurrencyMenu(locale: string) {
  const [usd, eur, gbp, cny, chf] = ACCOUNTING_CURRENCY_FORMATS;
  return locale.toLowerCase().startsWith("zh")
    ? [cny, usd, eur, gbp, chf]
    : [usd, gbp, eur, cny, chf];
}

// Currency spells its symbol the way Format Cells > Currency builds it (a bare
// $ or ¥, which the renderer, the family reader and the XLSX writer all take
// as a literal), so the dialog reopens these cells on Currency. zh-CN Excel's
// currency puts the minus after the yuan sign (¥-1,234.10, not -¥1,234.10),
// the same code its Ctrl+Shift+4 applies.
export const CURRENCY_USD_FORMAT = "$#,##0.00";
export const CURRENCY_CNY_FORMAT = "¥#,##0.00;¥-#,##0.00";

/**
 * The regional half of the gallery. Excel's Currency, Accounting, Short Date,
 * Long Date and Time follow the Windows region: en-US gets $, 5/18/1903,
 * "Monday, May 18, 1903" and the long time 1:26:24 PM; zh-CN gets ¥,
 * 1903/5/18, 1903年5月18日 and 13:26:24. The dropdown's Time is the regional
 * LONG time (with seconds). m/d/yyyy, h:mm:ss AM/PM and h:mm:ss are
 * SpreadsheetML built-ins 14, 19 and 21, so the XLSX writer keeps them as
 * built-in ids; the zh-CN dates are written as their own custom codes.
 */
interface RegionalNumberFormats {
  currency: string;
  accounting: string;
  shortDate: string;
  longDate: string;
  time: string;
}

const EN_US_NUMBER_FORMATS: RegionalNumberFormats = {
  currency: CURRENCY_USD_FORMAT,
  accounting: ACCOUNTING_USD_FORMAT,
  shortDate: "m/d/yyyy",
  longDate: "dddd, mmmm d, yyyy",
  time: "h:mm:ss AM/PM",
};

const ZH_CN_NUMBER_FORMATS: RegionalNumberFormats = {
  currency: CURRENCY_CNY_FORMAT,
  accounting: ACCOUNTING_CNY_FORMAT,
  shortDate: "yyyy/m/d",
  longDate: 'yyyy"年"m"月"d"日"',
  time: "h:mm:ss",
};

function regionalNumberFormats(locale: string): RegionalNumberFormats {
  return locale.toLowerCase().startsWith("zh")
    ? ZH_CN_NUMBER_FORMATS
    : EN_US_NUMBER_FORMATS;
}

/** Existing en/zh surfaces use the matching regional default face. */
export function defaultAccountingNumberFormat(locale: string): string {
  return regionalNumberFormats(locale).accounting;
}

/**
 * The code a cell's format is DRAWN with for this viewer. SpreadsheetML
 * built-ins 14 and 22 are the system short date and short date with time:
 * Excel shows them in the viewer's regional short date (m/d/yyyy in en-US,
 * yyyy/m/d in zh-CN). The XLSX importers (browser and API) spell those ids
 * m/d/yyyy and m/d/yyyy h:mm, and the writer maps those spellings back to 14
 * and 22, so only the drawing follows the locale; the stored code, and the
 * id written back, never change. Built-ins 20 and 21 (h:mm, h:mm:ss) read the
 * same in both regions. Every other code is returned as it is.
 */
export function regionalSystemDateDisplayCode(
  code: string,
  locale: string,
): string {
  const regional = regionalNumberFormats(locale);
  const shortDate = regional.shortDate;
  if (code === "m/d/yyyy") return shortDate;
  if (code === "m/d/yyyy h:mm") return `${shortDate} h:mm`;
  // [$-F800] and [$-F400] are Excel's system long date and system time: the
  // rest of the code is only a fallback, the region decides what is drawn
  // (Format Cells marks them "*", XLS-A055).
  if (/^\[\$-F800\]/i.test(code)) return regional.longDate;
  if (/^\[\$-F400\]/i.test(code)) return regional.time;
  return code;
}

/**
 * Excel's list, top to bottom. Number is two decimals WITHOUT a thousands
 * separator (the separator is the Comma Style button's job, which is why the
 * list has no Comma entry); Percentage, Fraction and Scientific are the
 * built-ins 10, 12 and 11.
 */
function numberFormatGalleryFor(
  regional: RegionalNumberFormats,
): readonly NumberFormatGalleryItem[] {
  return [
    {
      id: "general",
      labelKey: "spreadsheet.fmtCatGeneral",
      menuIcon: "general",
      code: "",
    },
    {
      id: "number",
      labelKey: "spreadsheet.fmtCatNumber",
      menuIcon: "number",
      code: "0.00",
    },
    {
      id: "currency",
      labelKey: "spreadsheet.fmtCatCurrency",
      menuIcon: "currency",
      code: regional.currency,
    },
    {
      id: "accounting",
      labelKey: "spreadsheet.fmtCatAccounting",
      menuIcon: "accounting",
      code: regional.accounting,
    },
    {
      id: "date",
      labelKey: "spreadsheet.fmtCatDate",
      menuIcon: "date",
      code: regional.shortDate,
    },
    {
      id: "long-date",
      labelKey: "spreadsheet.fmtCatDate",
      menuIcon: "long-date",
      code: regional.longDate,
    },
    {
      id: "time",
      labelKey: "spreadsheet.fmtCatTime",
      menuIcon: "time",
      code: regional.time,
    },
    {
      id: "percentage",
      labelKey: "spreadsheet.fmtCatPercentage",
      menuIcon: "percentage",
      code: "0.00%",
    },
    {
      id: "fraction",
      labelKey: "spreadsheet.fmtCatFraction",
      menuIcon: "fraction",
      code: "# ?/?",
    },
    {
      id: "scientific",
      labelKey: "spreadsheet.fmtCatScientific",
      menuIcon: "scientific",
      code: "0.00E+00",
    },
    {
      id: "text",
      labelKey: "spreadsheet.fmtCatText",
      menuIcon: "text",
      code: "@",
    },
  ];
}

/** The en-US list (the default for every non-Chinese locale). */
export const NUMBER_FORMAT_GALLERY: readonly NumberFormatGalleryItem[] =
  numberFormatGalleryFor(EN_US_NUMBER_FORMATS);

const ZH_CN_NUMBER_FORMAT_GALLERY: readonly NumberFormatGalleryItem[] =
  numberFormatGalleryFor(ZH_CN_NUMBER_FORMATS);

/** The gallery with the locale's regional Currency, Accounting, dates and time. */
export function localizedNumberFormatGallery(
  locale: string,
): readonly NumberFormatGalleryItem[] {
  return regionalNumberFormats(locale) === ZH_CN_NUMBER_FORMATS
    ? ZH_CN_NUMBER_FORMAT_GALLERY
    : NUMBER_FORMAT_GALLERY;
}

/** The locale's code for one gallery entry (the compact Home buttons reuse it). */
export function localizedNumberFormatGalleryCode(
  locale: string,
  id: "number" | "currency" | "date",
): string {
  return (
    localizedNumberFormatGallery(locale).find((item) => item.id === id)
      ?.code ?? ""
  );
}

const CUSTOM_NUMBER_FORMAT_ITEM: NumberFormatGalleryItem = {
  id: "custom",
  labelKey: "spreadsheet.fmtCatCustom",
  menuIcon: "text",
  code: "",
};

export function numberFormatGalleryItem(
  format: unknown,
  locale = "en",
): NumberFormatGalleryItem {
  const normalized = normalizeNumberFormatCode(format);
  const family = numberFormatFamily(format);
  if (normalized === null || family === null) return NUMBER_FORMAT_GALLERY[0];
  if (family === "custom") {
    return { ...CUSTOM_NUMBER_FORMAT_ITEM, code: normalized };
  }
  const gallery = localizedNumberFormatGallery(locale);
  // A code the other region's list wrote (a zh-CN long date read in English)
  // is still that entry, not the family's first one.
  const exactId = [
    ...gallery,
    ...NUMBER_FORMAT_GALLERY,
    ...ZH_CN_NUMBER_FORMAT_GALLERY,
  ].find((item) => normalizeNumberFormatCode(item.code) === normalized)?.id;
  return (
    gallery.find((item) => item.id === (exactId ?? family)) ??
    NUMBER_FORMAT_GALLERY[0]
  );
}
