/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format-code builder for the Format Cells dialog (设置单元格格式).
 *
 * Maps dialog state → Excel-compatible format code string. Pure functions with
 * no side effects, so they can be unit-tested in isolation and reused by the
 * dialog component without importing React.
 *
 * The produced codes are designed to round-trip through `applyNumberFormat` in
 * number-format.ts — do NOT duplicate formatting logic here.
 */

// A call-time use only (inferDialogState); number-format imports this module
// for its accounting codes, so nothing here may run at module load.
import { applyNumberFormat } from "./number-format";

// ── Category ─────────────────────────────────────────────────────────────────

export const FORMAT_CATEGORIES = [
  "General",
  "Number",
  "Currency",
  "Accounting",
  "Date",
  "Time",
  "Percentage",
  "Fraction",
  "Scientific",
  "Text",
  "Special",
  "Custom",
] as const;

export type FormatCategory = (typeof FORMAT_CATEGORIES)[number];

// ── Negative-number display style ────────────────────────────────────────────

export type NegativeStyle =
  | "minus"         // -1234.56  (plain minus, black); zh-CN currency ¥-1,234.56
  | "red"           // 1234.56   (red, no minus)
  | "red-minus"     // ¥-1,234.56 in red (zh-CN currency)
  | "parens"        // (1234.56) (black, parentheses)
  | "red-parens";   // (1234.56) (red, parentheses)

/**
 * The regional conventions Excel builds Currency / Accounting codes from (the
 * Windows region, which MRAI reads from the UI locale as the Home gallery
 * does). zh-CN puts the minus after the currency sign and writes Accounting as
 * `_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;…` (zh-CN built-ins 41-44, XLS-A053);
 * en-US writes `_($* #,##0.00_);_($* (#,##0.00);…` (built-ins 41-44 en-US).
 */
export type NumberFormatRegion = "en" | "zh";

export function numberFormatRegion(locale: string | undefined): NumberFormatRegion {
  return (locale ?? "").toLowerCase().startsWith("zh") ? "zh" : "en";
}

/** The negative-number choices Excel lists for a category in a region. */
export function negativeStylesFor(
  category: FormatCategory,
  region: NumberFormatRegion,
): NegativeStyle[] {
  if (category === "Currency" && region === "zh") return ["minus", "red", "red-minus"];
  return ["minus", "red", "parens", "red-parens"];
}

// ── Fraction display ─────────────────────────────────────────────────────────

export type FractionType =
  | "up-to-one-digit"
  | "up-to-two-digits"
  | "up-to-three-digits"
  | "halves"
  | "quarters"
  | "eighths"
  | "sixteenths"
  | "tenths"
  | "hundredths";

/**
 * Excel's Fraction Type list (XLS-A057): Up to one/two/three digits, As
 * halves/quarters/eighths/sixteenths, As tenths (3/10), As hundredths
 * (30/100), each shown as a description plus a sample.
 */
export const FRACTION_TYPE_CODES: ReadonlyArray<readonly [FractionType, string]> = [
  ["up-to-one-digit", "# ?/?"],
  ["up-to-two-digits", "# ??/??"],
  ["up-to-three-digits", "# ???/???"],
  ["halves", "# ?/2"],
  ["quarters", "# ?/4"],
  ["eighths", "# ?/8"],
  ["sixteenths", "# ??/16"],
  ["tenths", "# ?/10"],
  ["hundredths", "# ??/100"],
];

// ── Date / Time presets ───────────────────────────────────────────────────────

/** A Date / Time Type list entry is a stored format code. */
export type DatePreset = string;
export type TimePreset = string;

export interface NumberFormatTypeEntry {
  code: string;
  /** Follows the operating system's region (Excel marks these with "*"). */
  system?: boolean;
}

 
export const DATE_TYPE_SAMPLE_SERIAL = 40982 + (13 * 3600 + 30 * 60 + 55) / 86400;

/**
 * Excel's Date Type lists (XLS-A055), top to bottom. The two starred entries
 * are the system short date (built-in 14, stored as m/d/yyyy) and the system
 * long date ([$-F800]); every other entry is a fixed code ending in ";@", as
 * Excel writes it.
 */
const EN_DATE_TYPES: NumberFormatTypeEntry[] = [
  { code: "m/d/yyyy", system: true },
  { code: "[$-F800]dddd, mmmm dd, yyyy", system: true },
  { code: "m/d;@" },
  { code: "m/d/yy;@" },
  { code: "mm/dd/yy;@" },
  { code: "[$-409]d-mmm;@" },
  { code: "[$-409]d-mmm-yy;@" },
  { code: "[$-409]dd-mmm-yy;@" },
  { code: "[$-409]mmm-yy;@" },
  { code: "[$-409]mmmm-yy;@" },
  { code: "[$-409]mmmm d, yyyy;@" },
  { code: "[$-409]m/d/yy h:mm AM/PM;@" },
  { code: "m/d/yy h:mm;@" },
  { code: "[$-409]mmmmm;@" },
  { code: "[$-409]mmmmm-yy;@" },
  { code: "m/d/yyyy;@" },
  { code: "[$-409]d-mmm-yyyy;@" },
];

const ZH_DATE_TYPES: NumberFormatTypeEntry[] = [
  { code: "m/d/yyyy", system: true },
  { code: "[$-F800]dddd, mmmm dd, yyyy", system: true },
  { code: '[DBNum1][$-804]yyyy"年"m"月"d"日";@' },
  { code: '[DBNum1][$-804]yyyy"年"m"月";@' },
  { code: '[DBNum1][$-804]m"月"d"日";@' },
  { code: 'yyyy"年"m"月"d"日";@' },
  { code: 'yyyy"年"m"月";@' },
  { code: 'm"月"d"日";@' },
  { code: "[$-804]aaaa;@" },
  { code: "[$-804]aaa;@" },
  { code: "yyyy/m/d;@" },
  { code: "[$-409]yyyy/m/d h:mm AM/PM;@" },
  { code: "yyyy/m/d h:mm;@" },
  { code: "yy/m/d;@" },
  { code: "m/d;@" },
  { code: "m/d/yy;@" },
  { code: "mm/dd/yy;@" },
  { code: "[$-409]d-mmm;@" },
  { code: "[$-409]d-mmm-yy;@" },
  { code: "[$-409]dd-mmm-yy;@" },
  { code: "[$-409]mmm-yy;@" },
  { code: "[$-409]mmmm-yy;@" },
  { code: "[$-409]mmmmm;@" },
  { code: "[$-409]mmmmm-yy;@" },
];

/** Excel's Time Type lists (XLS-A055); the starred entry is the system time. */
const EN_TIME_TYPES: NumberFormatTypeEntry[] = [
  { code: "[$-F400]h:mm:ss AM/PM", system: true },
  { code: "h:mm;@" },
  { code: "[$-409]h:mm AM/PM;@" },
  { code: "h:mm:ss;@" },
  { code: "[$-409]h:mm:ss AM/PM;@" },
  { code: "mm:ss.0;@" },
  { code: "[h]:mm:ss;@" },
  { code: "[$-409]m/d/yy h:mm AM/PM;@" },
  { code: "m/d/yy h:mm;@" },
];

const ZH_TIME_TYPES: NumberFormatTypeEntry[] = [
  { code: "[$-F400]h:mm:ss AM/PM", system: true },
  { code: "h:mm;@" },
  { code: "[$-409]h:mm AM/PM;@" },
  { code: "h:mm:ss;@" },
  { code: "[$-409]h:mm:ss AM/PM;@" },
  { code: 'h"时"mm"分";@' },
  { code: 'h"时"mm"分"ss"秒";@' },
  { code: '上午/下午h"时"mm"分";@' },
  { code: '上午/下午h"时"mm"分"ss"秒";@' },
  { code: '[DBNum1][$-804]h"时"mm"分";@' },
  { code: '[DBNum1][$-804]上午/下午h"时"mm"分";@' },
  { code: "[$-409]yyyy/m/d h:mm AM/PM;@" },
  { code: "yyyy/m/d h:mm;@" },
];

export function dateTypeList(region: NumberFormatRegion): readonly NumberFormatTypeEntry[] {
  return region === "zh" ? ZH_DATE_TYPES : EN_DATE_TYPES;
}

export function timeTypeList(region: NumberFormatRegion): readonly NumberFormatTypeEntry[] {
  return region === "zh" ? ZH_TIME_TYPES : EN_TIME_TYPES;
}

/**
 * Date and time codes an earlier MRAI dialog or the Home Number Format list
 * wrote. They are not in Excel's Type lists, but a cell carrying one still
 * reopens on Date / Time (with its own code appended to the list) instead of
 * falling to Custom.
 */
const LEGACY_DATE_PRESETS = [
  "dddd, mmmm d, yyyy",
  "yyyy/m/d",
  'yyyy"年"m"月"d"日"',
  'yyyy"年"m"月"',
  'm"月"d"日"',
  "yyyy-mm-dd",
  "dd/mm/yyyy",
  "mm/dd/yyyy",
  "d-mmm-yy",
  "mmmm d, yyyy",
  "d mmm yyyy",
];
const LEGACY_TIME_PRESETS = ["h:mm AM/PM", "h:mm:ss AM/PM", "HH:mm", "HH:mm:ss", "h:mm:ss"];

// ── Special ───────────────────────────────────────────────────────────────────

/**
 * Excel's Special types for the region (XLS-A056): zh-CN 邮政编码 / 中文小写数字
 * / 中文大写数字; en-US Zip Code / Zip Code + 4 / Phone Number / Social
 * Security Number.
 */
export type SpecialTypeId =
  | "postalCodeZh"
  | "chineseLower"
  | "chineseUpper"
  | "zipCode"
  | "zipCodePlus4"
  | "phoneNumber"
  | "socialSecurity";

export function specialTypeList(
  region: NumberFormatRegion,
): ReadonlyArray<{ id: SpecialTypeId; code: string }> {
  return region === "zh"
    ? [
        { id: "postalCodeZh", code: "000000" },
        { id: "chineseLower", code: "[DBNum1][$-804]General" },
        { id: "chineseUpper", code: "[DBNum2][$-804]General" },
      ]
    : [
        { id: "zipCode", code: "00000" },
        { id: "zipCodePlus4", code: "00000-0000" },
        { id: "phoneNumber", code: "[<=9999999]###-####;(###) ###-####" },
        { id: "socialSecurity", code: "000-00-0000" },
      ];
}

const ALL_SPECIAL_CODES = new Set(
  [...specialTypeList("zh"), ...specialTypeList("en")].map((entry) => entry.code),
);

// ── Custom ────────────────────────────────────────────────────────────────────

/**
 * The built-in codes Excel's Custom Type list starts with (XLS-A056), as the
 * region's Excel lists them. The system short date is the stored m/d/yyyy
 * (drawn regionally).
 */
export function customTypeList(region: NumberFormatRegion): readonly string[] {
  const shared = [
    "General",
    "0",
    "0.00",
    "#,##0",
    "#,##0.00",
  ];
  const tail = [
    "0%",
    "0.00%",
    "0.00E+00",
    "##0.0E+0",
    "# ?/?",
    "# ??/??",
    "m/d/yyyy",
    "d-mmm-yy",
    "d-mmm",
    "mmm-yy",
    "h:mm AM/PM",
    "h:mm:ss AM/PM",
    "h:mm",
    "h:mm:ss",
    "m/d/yyyy h:mm",
    "mm:ss",
    "mm:ss.0",
    "@",
    "[h]:mm:ss",
  ];
  if (region === "zh") {
    return [
      ...shared,
      "¥#,##0;¥-#,##0",
      "¥#,##0;[Red]¥-#,##0",
      "¥#,##0.00;¥-#,##0.00",
      "¥#,##0.00;[Red]¥-#,##0.00",
      ...tail,
      buildAccountingCodeFromParams("¥", 0, "zh"),
      buildAccountingCodeFromParams("", 0, "zh"),
      buildAccountingCodeFromParams("¥", 2, "zh"),
      buildAccountingCodeFromParams("", 2, "zh"),
      'yyyy"年"m"月"',
      'm"月"d"日"',
      'yyyy"年"m"月"d"日"',
      'h"时"mm"分"',
      'h"时"mm"分"ss"秒"',
      '上午/下午h"时"mm"分"',
      '上午/下午h"时"mm"分"ss"秒"',
      "[DBNum1][$-804]General",
      "[DBNum2][$-804]General",
    ];
  }
  return [
    ...shared,
    "#,##0_);(#,##0)",
    "#,##0_);[Red](#,##0)",
    "#,##0.00_);(#,##0.00)",
    "#,##0.00_);[Red](#,##0.00)",
    "$#,##0_);($#,##0)",
    "$#,##0_);[Red]($#,##0)",
    "$#,##0.00_);($#,##0.00)",
    "$#,##0.00_);[Red]($#,##0.00)",
    ...tail,
    buildAccountingCodeFromParams("$", 0, "en"),
    buildAccountingCodeFromParams("", 0, "en"),
    buildAccountingCodeFromParams("$", 2, "en"),
    buildAccountingCodeFromParams("", 2, "en"),
  ];
}

// ── Dialog state ──────────────────────────────────────────────────────────────

export interface FormatDialogState {
  category: FormatCategory;
  /** Decimal places (0–30). Used by Number, Currency, Accounting, Percentage, Scientific. */
  decimals: number;
  /** Show thousands separator (e.g. 1,234). Number only. */
  thousands: boolean;
  /** Currency / Accounting symbol. */
  symbol: string;
  /** Negative number display style. Number, Currency, Accounting. */
  negativeStyle: NegativeStyle;
  /** Date format preset. */
  datePreset: DatePreset;
  /** Time format preset. */
  timePreset: TimePreset;
  /** Fraction display type. */
  fractionType: FractionType;
  /** Raw custom format code (Custom category). */
  customCode: string;
  /** Regional conventions for Currency / Accounting codes (default en). */
  region?: NumberFormatRegion;
}

// ── Default state ─────────────────────────────────────────────────────────────

export const DEFAULT_DIALOG_STATE: FormatDialogState = {
  category: "General",
  decimals: 2,
  // Excel's Number category opens with the 1000 separator unchecked.
  thousands: false,
  symbol: "$",
  negativeStyle: "minus",
  datePreset: "m/d/yyyy",
  timePreset: "[$-F400]h:mm:ss AM/PM",
  fractionType: "up-to-one-digit",
  customCode: "General",
};

/** The dialog's defaults in a region: zh-CN currency and accounting open on ¥
 * (XLS-A053 / XLS-A105), en-US on $. */
export function defaultDialogState(region: NumberFormatRegion): FormatDialogState {
  return { ...DEFAULT_DIALOG_STATE, region, symbol: region === "zh" ? "¥" : "$" };
}

// ── Symbol list ───────────────────────────────────────────────────────────────

export const CURRENCY_SYMBOLS = ["$", "€", "£", "¥", "¥", "₹", "₩", "₽", "₪", "₺"] as const;
// Note: first ¥ = U+00A5 (generic Yen), second is kept distinct for future locale use.

/**
 * One Symbol choice. `token` is what the code carries: "" for None, a bare
 * glyph, or a locale-tagged `[$<glyph>-<LCID>]` that Excel lists as
 * "<glyph> <language (region)>" (e.g. "¥ 中文(中国)", "$ 英语(美国)").
 */
export interface CurrencySymbolOption {
  token: string;
  glyph: string;
  /** BCP-47 tag whose display name labels a locale-tagged entry. */
  locale?: string;
}

const TAGGED_CURRENCY_SYMBOLS: CurrencySymbolOption[] = [
  { token: "[$¥-804]", glyph: "¥", locale: "zh-CN" },
  { token: "[$$-409]", glyph: "$", locale: "en-US" },
  { token: "[$£-809]", glyph: "£", locale: "en-GB" },
  { token: "[$€-407]", glyph: "€", locale: "de-DE" },
  { token: "[$€-40C]", glyph: "€", locale: "fr-FR" },
  { token: "[$CHF-100C]", glyph: "CHF", locale: "fr-CH" },
  { token: "[$¥-411]", glyph: "¥", locale: "ja-JP" },
  { token: "[$₩-412]", glyph: "₩", locale: "ko-KR" },
  { token: "[$₹-4009]", glyph: "₹", locale: "en-IN" },
  { token: "[$HK$-C04]", glyph: "HK$", locale: "zh-HK" },
  { token: "[$NT$-404]", glyph: "NT$", locale: "zh-TW" },
];

/** Excel's Symbol list for the region: None, the bare glyphs (the region's own
 * first), then the locale-tagged entries. */
export function currencySymbolOptions(region: NumberFormatRegion): CurrencySymbolOption[] {
  const bare = [...new Set(CURRENCY_SYMBOLS)];
  const first = region === "zh" ? "¥" : "$";
  return [
    { token: "", glyph: "" },
    ...[first, ...bare.filter((glyph) => glyph !== first)].map((glyph) => ({ token: glyph, glyph })),
    ...TAGGED_CURRENCY_SYMBOLS,
  ];
}

// ── Format code builder ───────────────────────────────────────────────────────

/**
 * Build an Excel-compatible format code from the dialog state.
 * Pure function — no React, no side effects.
 */
export function buildFormatCode(state: FormatDialogState): string {
  const { category } = state;

  switch (category) {
    case "General":
      return "General";

    case "Number":
      return buildNumberCode(state);

    case "Currency":
      return buildCurrencyCode(state);

    case "Accounting":
      return buildAccountingCode(state);

    case "Date":
      return state.datePreset;

    case "Time":
      return state.timePreset;

    case "Percentage":
      return buildPercentageCode(state);

    case "Fraction":
      return buildFractionCode(state);

    case "Scientific":
      return buildScientificCode(state);

    case "Text":
      return "@";

    case "Special":
      // Commonly-used special formats (zip / phone / SSN style). We expose
      // a small curated set via the customCode field when Special is chosen.
      return state.customCode || "@";

    case "Custom":
      return state.customCode || "General";

    default: {
      // Exhaustiveness guard (TypeScript will catch unhandled members).
      void (category as never);
      return "General";
    }
  }
}

// ── Number ────────────────────────────────────────────────────────────────────

function buildNumberCode(state: FormatDialogState): string {
  const intToken = state.thousands ? "#,##0" : "0";
  const fracToken = decimalToken(state.decimals);
  const positiveCode = `${intToken}${fracToken}`;

  return applyNegativeStyle(positiveCode, state.negativeStyle, "");
}

// ── Currency ──────────────────────────────────────────────────────────────────

function buildCurrencyCode(state: FormatDialogState): string {
  const sym = escapeCurrencySymbol(state.symbol);
  const intToken = "#,##0";
  const fracToken = decimalToken(state.decimals);
  const positiveCode = `${sym}${intToken}${fracToken}`;

  if (state.region === "zh") {
    // zh-CN built-ins 5-8: the minus after the symbol (¥-1,234.10).
    const negative = `${sym}-${intToken}${fracToken}`;
    switch (state.negativeStyle) {
      case "minus":
        return `${positiveCode};${negative}`;
      case "red":
        return `${positiveCode};[Red]${positiveCode}`;
      case "red-minus":
        return `${positiveCode};[Red]${negative}`;
      default:
        break;
    }
  }
  return applyNegativeStyle(positiveCode, state.negativeStyle, sym);
}

// ── Accounting ────────────────────────────────────────────────────────────────

/**
 * Excel canonical accounting format:
 *   symbol left-aligned (with fill), number right-aligned,
 *   negatives in parentheses, zero as dash.
 *
 * The accounting renderer in number-format.ts detects isAccountingSection
 * on the `_($*` / `_(¥*` etc. prefix.
 */
function buildAccountingCode(state: FormatDialogState): string {
  return buildAccountingCodeFromParams(state.symbol, state.decimals, state.region ?? "en");
}

/**
 * Build the canonical 4-section Excel Accounting code from a RAW currency symbol
 * (e.g. "$" or "₽" — re-encoded via escapeCurrencySymbol) and a decimal count.
 * The single source of truth for the accounting shape, shared by the Format
 * Cells dialog and number-format's Increase/Decrease-Decimals (which must rebuild
 * the whole structure rather than collapse it to a bare numeric code).
 */
export function buildAccountingCodeFromParams(
  symbol: string,
  decimals: number,
  region: NumberFormatRegion = "en",
): string {
  const sym = escapeCurrencySymbol(symbol);
  const fracToken = decimalToken(decimals);
  if (region === "zh") {
    // zh-CN built-ins 41-44: `_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ `
    // (measured in a zh-CN Excel, XLS-A053). None drops the symbol: `_ * …`.
    const lead = `_ ${sym}* `;
    const dash = decimals > 0 ? "??" : "";
    return `${lead}#,##0${fracToken}_ ;${lead}-#,##0${fracToken}_ ;${lead}"-"${dash}_ ;_ @_ `;
  }
  const pos = `_(${sym}* #,##0${fracToken}_)`;
  const neg = `_(${sym}* (#,##0${fracToken})`;
  const zero = `_(${sym}* "-"${decimalDashToken(decimals)}_)`;
  const text = `_(@_)`;
  return `${pos};${neg};${zero};${text}`;
}

// ── Percentage ────────────────────────────────────────────────────────────────

function buildPercentageCode(state: FormatDialogState): string {
  return `0${decimalToken(state.decimals)}%`;
}

// ── Fraction ─────────────────────────────────────────────────────────────────

function buildFractionCode(state: FormatDialogState): string {
  return FRACTION_TYPE_CODES.find(([type]) => type === state.fractionType)?.[1] ?? "# ?/?";
}

// ── Scientific ────────────────────────────────────────────────────────────────

function buildScientificCode(state: FormatDialogState): string {
  return `0${decimalToken(state.decimals)}E+00`;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Return ".000…" for `decimals` > 0, or "" for 0. */
function decimalToken(decimals: number): string {
  if (decimals <= 0) return "";
  return "." + "0".repeat(Math.min(30, decimals));
}

/** Return "??_)" placeholder for the zero-section accounting dash alignment. */
function decimalDashToken(decimals: number): string {
  if (decimals <= 0) return "_)";
  return "??_)";
}

/**
 * Wrap a positive-section code with a negative section, depending on the
 * requested negative style. `sym` is the raw currency symbol (or "") so red /
 * paren variants can replicate it on the negative section.
 */
function applyNegativeStyle(
  positiveCode: string,
  style: NegativeStyle,
  sym: string,
): string {
  switch (style) {
    case "minus":
      // Excel omits a negative section and the engine auto-prefixes "-".
      return positiveCode;

    case "red":
      // Positive black, negative red (no sign).
      return `${positiveCode};[Red]${positiveCode}`;

    // Excel's parenthesised styles reserve the ")" width on the positive
    // section so the digits line up (built-ins 7-8, 39-40: `#,##0.00_);(…)`).
    case "parens":
      return `${positiveCode}_);(${sym}${positiveCode.replace(sym, "")})`;

    case "red-parens":
      return `${positiveCode}_);[Red](${sym}${positiveCode.replace(sym, "")})`;

    case "red-minus":
      return `${positiveCode};[Red]-${positiveCode}`;

    default: {
      void (style as never);
      return positiveCode;
    }
  }
}

/**
 * Escape a currency symbol for use in an Excel format code.
 * Native $ needs no escaping; non-native symbols must be quoted or use the
 * [$SYM] locale bracket — we prefer [$SYM] for round-trip reliability
 * with the engine's CURRENCY_RE which matches $€¥£₹ directly.
 */
function escapeCurrencySymbol(sym: string): string {
  // None, or an already-bracketed locale token ([$¥-804]).
  if (sym === "" || /^\[\$[^\]]*\]$/.test(sym)) return sym;
  // The engine's CURRENCY_RE already matches $, €, ¥, £, ₹ as unquoted literals.
  if (/[$€¥£₹]/.test(sym)) return sym;
  // Other symbols: use [$SYM] bracket (Excel-compatible locale prefix).
  return `[$${sym}]`;
}

// ── State initializer from an existing format string ─────────────────────────

/**
 * Infer the dialog state that best matches an existing format code. Used to
 * pre-populate the dialog when it opens on a cell that already has a format.
 */
export function inferDialogState(formatCode: string, locale = "en"): FormatDialogState {
  const inferred = inferPresetDialogState(formatCode, locale);
  const f = (formatCode || "").trim();
  // Currency and Accounting keep their reading: their codes carry regional
  // spellings (zh-CN's ¥-#,##0.00) the dialog's negative styles do not.
  if (!["Number", "Percentage", "Scientific", "Fraction", "Date", "Time"].includes(inferred.category) || !f) {
    return inferred;
  }
  // A preset the dialog would rebuild differently is not that preset: a
  // three-section `0.00;"N"0.000;"Z"` read as Number would preview and keep
  // `0.00` (live audit B444). Open it as the custom code it is.
  // "Rebuild the same" means showing every value the same way; the builder
  // may spell an equivalent code differently.
  const rebuilt = buildFormatCode(inferred);
  const samples = [1234.567, -1234.567, 0, 0.5, -0.5, 1e7];
  const sameDisplay = samples.every((value) => applyNumberFormat(value, rebuilt) === applyNumberFormat(value, f));
  return sameDisplay ? inferred : { ...defaultDialogState(numberFormatRegion(locale)), category: "Custom", customCode: f };
}

function inferPresetDialogState(formatCode: string, locale: string): FormatDialogState {
  const region = numberFormatRegion(locale);
  const base = defaultDialogState(region);
  const f = (formatCode || "").trim();

  if (!f || f.toLowerCase() === "general") {
    return { ...base, category: "General" };
  }
  if (f === "@") {
    return { ...base, category: "Text", customCode: "@" };
  }
  if (ALL_SPECIAL_CODES.has(f)) {
    return { ...base, category: "Special", customCode: f };
  }

  // Date / Time: Excel's Type lists of both regions, then the codes earlier
  // MRAI lists and the Home gallery wrote.
  const sameCode = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
  for (const entry of [...dateTypeList(region), ...dateTypeList(region === "zh" ? "en" : "zh")]) {
    if (sameCode(entry.code, f)) return { ...base, category: "Date", datePreset: entry.code };
  }
  for (const entry of [...timeTypeList(region), ...timeTypeList(region === "zh" ? "en" : "zh")]) {
    if (sameCode(entry.code, f)) return { ...base, category: "Time", timePreset: entry.code };
  }
  for (const p of LEGACY_DATE_PRESETS) {
    if (sameCode(p, f)) return { ...base, category: "Date", datePreset: p };
  }
  for (const p of LEGACY_TIME_PRESETS) {
    if (sameCode(p, f)) return { ...base, category: "Time", timePreset: p };
  }

  // Percentage
  if (f.endsWith("%")) {
    const inner = f.slice(0, -1);
    const dotIdx = inner.indexOf(".");
    const decimals = dotIdx >= 0 ? (inner.length - dotIdx - 1) : 0;
    return { ...base, category: "Percentage", decimals };
  }

  // Scientific
  if (/e[+\-]/i.test(f)) {
    const dotIdx = f.indexOf(".");
    const eIdx = f.search(/e/i);
    const decimals = dotIdx >= 0 && eIdx > dotIdx ? eIdx - dotIdx - 1 : 0;
    return { ...base, category: "Scientific", decimals };
  }

  // Fraction (the pre-A057 "# ?/16" sixteenths still reopens on Fraction)
  const fraction = FRACTION_TYPE_CODES.find(([, code]) => code === f)?.[0]
    ?? (f === "# ?/16" ? "sixteenths" : undefined);
  if (fraction) {
    return { ...base, category: "Fraction", fractionType: fraction };
  }

  // Accounting: en-US `_($* …`, zh-CN `_ ¥* …` (symbol optional = None). A
  // code the dialog would not write back unchanged stays Custom.
  const accounting = /^_([( ])(\[\$[^\]]*\]|[^*\s"]*)\* #,##0(?:\.(0+))?_/.exec(f);
  if (accounting) {
    const accountingRegion: NumberFormatRegion = accounting[1] === " " ? "zh" : "en";
    const state: FormatDialogState = {
      ...base,
      region: accountingRegion,
      category: "Accounting",
      symbol: accounting[2] ?? "",
      decimals: accounting[3]?.length ?? 0,
    };
    // Stored codes are trimmed, so the zh-CN text section's trailing "_ "
    // reserve is compared without it.
    if (buildFormatCode(state).trim() === f) return state;
    return { ...base, category: "Custom", customCode: f };
  }

  // Currency (a leading symbol, bare or locale-tagged)
  const currencyToken = /^(\[\$[^\]-]+-[^\]]+\]|\[\$[^\]-]+\]|[$€£¥₹₩₽₪₺])/.exec(f)?.[1];
  if (currencyToken) {
    const symbol = /^\[\$[^\]-]+\]$/.test(currencyToken)
      ? currencyToken.slice(2, -1)
      : currencyToken;
    const sections = f.split(";");
    const pos = sections[0] ?? "";
    const neg = sections[1] ?? "";
    const dotIdx = pos.indexOf(".");
    const decimals = dotIdx >= 0 ? (pos.slice(dotIdx + 1).match(/0/g) || []).length : 0;
    // zh-CN's shape puts the minus right after the symbol (¥-#,##0.00).
    const minusAfterSymbol = neg.replace("[Red]", "").startsWith(`${currencyToken}-`);
    const red = neg.includes("[Red]");
    const negStyle: NegativeStyle = minusAfterSymbol
      ? red ? "red-minus" : "minus"
      : red && neg.includes("(")
        ? "red-parens"
        : red
          ? "red"
          : neg.includes("(")
            ? "parens"
            : "minus";
    const currencyRegion: NumberFormatRegion = minusAfterSymbol ? "zh" : region;
    const state: FormatDialogState = {
      ...base,
      region: currencyRegion,
      category: "Currency",
      symbol,
      decimals,
      negativeStyle: negStyle,
    };
    return state;
  }

  // Number (numeric only, possibly with comma grouping)
  if (/^[#0,?.]+(?:\.[0#?]*)?(?:_\))?(?:;.*)?$/.test(f)) {
    const sections = f.split(";");
    const pos = sections[0] ?? "";
    const thousands = pos.includes(",");
    const dotIdx = pos.indexOf(".");
    const decimals = dotIdx >= 0 ? (pos.slice(dotIdx + 1).match(/[0#?]/g) || []).length : 0;
    const negStyle: NegativeStyle =
      sections.length > 1 && (sections[1] ?? "").includes("Red") && (sections[1] ?? "").includes("(")
        ? "red-parens"
        : sections.length > 1 && (sections[1] ?? "").includes("Red")
          ? "red"
          : sections.length > 1 && (sections[1] ?? "").includes("(")
            ? "parens"
            : "minus";
    return { ...base, category: "Number", thousands, decimals, negativeStyle: negStyle };
  }

  // Fallback: Custom
  return { ...base, category: "Custom", customCode: f };
}
