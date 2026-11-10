/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a typed or pasted entry becomes, as in Excel.
 *
 * Excel does not store what was typed into a General cell: it reads the entry
 * as a number, a date, a time, a logical or an error and gives the cell the
 * number format the entry implied (Microsoft Support, "Enter data manually in
 * worksheet cells", "Available number formats in Excel" and "Stop
 * automatically changing numbers to dates"). Measured in Excel for Mac 16
 * (zh UI), entry => value / format:
 *
 *   50% => 0.5 / 0%          12.5% => 0.125 / 0.00%     1,234 => 1234 / #,##0
 *   $1,000 => 1000 / $#,##0_);[Red]($#,##0)             1.50 => 1.5 / General
 *   00123 => 123 / General   1e3 => 1000 / 0.00E+00     (123) => -123 / General
 *   true => TRUE (logical)   10:30 => 0.4375 / h:mm      1 1/2 => 1.5 / # ?/?
 *   2026/9/21 => date / the system short date           9/21 => this year's
 *   September 21 / m"月"d"日" (en-US: d-mmm)
 *
 * The currency symbol and the order of a date's parts follow the region. MRAI
 * has two regional sets, the same two the Home number list uses: zh (¥ and
 * year/month/day, 2026年9月21日 included, as zh-CN Excel reads them) and
 * en-US for every other locale ($ and month/day/year).
 *
 * A cell formatted as Text keeps the entry as typed, and a leading apostrophe
 * forces text: the apostrophe is not part of the value (Excel keeps it as the
 * cell's quote prefix and shows it only while editing). A space before a date
 * keeps it text, as "Stop automatically changing numbers to dates" says.
 *
 * Formulas are not entries of this kind; the caller keeps them as they are.
 *
 * The digits of a number the USER types into a cell or the formula bar are
 * read with the interface language's decimal and thousands separators, as
 * Excel reads an entry with the system separators (Microsoft Support, "Change
 * the character used to separate thousands or decimals": the separators come
 * from the regional settings unless File > Options > Advanced > Use system
 * separators is cleared). MRAI's region is its interface language
 * (locale-number-separators.ts), the same separators the grid draws with: a
 * German entry 5,5 is 5.5 and 1.234,5 is 1234.5, a French 1 234,5 is 1234.5
 * (any space, U+0020, U+00A0 or U+202F, groups the digits there).
 * `typedNumerals` asks for that reading; formulas keep their English form
 * (1.5 and commas between arguments), and every other writer (scenarios,
 * macros, forms, Replace, filters, pastes, AI) keeps the canonical "." form,
 * which a typed entry also falls back to when the regional reading does not
 * match (5.5 typed in German is 5.5, where Excel would read a date or text).
 */

import {
  excelMaxDateSerial,
  excelSystemSerialFromCalendarDate,
} from "./excel-1900-calendar";
import {
  applyNumberFormat,
  formatGeneralNumber,
  normalizeNumberFormatCode,
  numberFormatFamily,
} from "./number-format";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";
import { excelMinDateSerial } from "./excel-1900-calendar";
import {
  readDataConversionOptions,
  type DataConversionOptions,
} from "./data-conversion-options";
import { viewerNumberSeparators } from "./locale-number-separators";

export interface SpreadsheetEntry {
  /** The value to store: a canonical number, TRUE/FALSE, an error code or text. */
  value: string;
  kind: "text" | "number";
  /** The format the entry implies; it is applied only to a General cell. */
  numberFormat?: string;
  /** Entered with a leading apostrophe. */
  quotePrefix?: true;
}

export interface SpreadsheetEntryOptions {
  /** The UI locale; zh* reads the zh-CN region, anything else en-US. */
  locale: string;
  /** The cell is formatted as Text (@): the entry is stored as typed. */
  textFormat?: boolean;
  /** Today, for a date typed without a year. */
  today?: Date;
  /**
   * The workbook uses the 1904 date system: a typed date is stored as its 1904
   * serial (1904-05-05 is 125), and a date before 1904 stays text (Excel keeps
   * "1903-05-05" and "2/29/1900" as typed there).
   */
  date1904?: ExcelDateSystemArg;
  /**
   * The cell is formatted as a percentage. Excel's entry into such a cell
   * reads a number of 1 or more as that many percent and a smaller one as the
   * fraction it already is: 10 and 0.1 are both 10% (Microsoft Support,
   * "Format numbers as percentages"). A number typed with its own % sign, or
   * with any other symbol that implies a format, is read as typed.
   */
  percentFormat?: boolean;
  /**
   * File > Options > Data > Automatic Data Conversion
   * (lib/spreadsheet/data-conversion-options): the conversions this entry may
   * make. Omitted: the user's own choice, as every entry in Excel reads it.
   */
  conversions?: DataConversionOptions;
  /**
   * The entry was typed by the user into a cell or the formula bar: its
   * numeral is read with the interface language's separators first (see the
   * module comment). Omitted, only the canonical "." and "," are read.
   */
  typedNumerals?: boolean;
}

/** Any space a region that groups digits with a space accepts as the group. */
const SPACE_GROUP_RE = /[ \u00a0\u202f\u2009]/g;

/**
 * The separators a typed numeral is read with in `locale`, or null where they
 * are the canonical "." and "," (en, zh, ja, ko, th): a space group (fr, sv,
 * fi, no) is normalized to U+0020.
 */
export function typedNumeralSeparators(locale: string): DecimalNumeralSeparators | null {
  const separators = viewerNumberSeparators(locale);
  if (!separators) return null;
  const decimal = separators.decimal;
  const group = separators.group.replace(SPACE_GROUP_RE, " ") || null;
  if (!decimal || decimal === group) return null;
  return { decimal, group };
}

/** The canonical point a typed entry falls back to, without grouping. */
const CANONICAL_POINT: DecimalNumeralSeparators = { decimal: ".", group: null };

interface EntryRegion {
  zh: boolean;
  currencySymbols: readonly string[];
  currencyFormat(decimals: boolean): string;
  /** The date a month/day entry is shown with. */
  monthDayFormat: string;
  /** The regional short date and time, for the text an edit starts from. */
  shortDate: string;
  longTime: string;
}

const ZH_REGION: EntryRegion = {
  zh: true,
  currencySymbols: ["¥", "￥"],
  // zh-CN Excel's built-ins 6 and 8: the minus after the yuan sign.
  currencyFormat: (decimals) =>
    decimals ? "¥#,##0.00;[Red]¥-#,##0.00" : "¥#,##0;[Red]¥-#,##0",
  monthDayFormat: 'm"月"d"日"',
  shortDate: "yyyy/m/d",
  longTime: "h:mm:ss",
};

const EN_US_REGION: EntryRegion = {
  zh: false,
  currencySymbols: ["$"],
  // en-US Excel's built-ins 6 and 8 (measured: $1,000 => $#,##0_);[Red]($#,##0)).
  currencyFormat: (decimals) =>
    decimals ? "$#,##0.00_);[Red]($#,##0.00)" : "$#,##0_);[Red]($#,##0)",
  monthDayFormat: "d-mmm",
  shortDate: "m/d/yyyy",
  longTime: "h:mm:ss AM/PM",
};

function entryRegion(locale: string): EntryRegion {
  return locale.toLowerCase().startsWith("zh") ? ZH_REGION : EN_US_REGION;
}

/** Built-ins 14 and 22: drawn in the viewer's regional short date. */
const SYSTEM_SHORT_DATE = "m/d/yyyy";
const SYSTEM_SHORT_DATE_TIME = "m/d/yyyy h:mm";

const ERROR_LITERALS = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
];

const MONTH_NAMES = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

function monthFromName(name: string): number | null {
  const lower = name.toLowerCase();
  if (lower.length < 3) return null;
  // Excel reads SEPT as September (SEPT2 is 2-Sep, measured).
  if (lower === "sept") return 9;
  const index = MONTH_NAMES.findIndex((month) =>
    lower.length === 3 ? month.startsWith(lower) : month === lower,
  );
  return index >= 0 ? index + 1 : null;
}








function dateSerial(
  year: number,
  month: number,
  day: number,
  date1904: ExcelDateSystemArg = false,
): number | null {
  if (year < 1900 || year > 9999) return null;
  return excelSystemSerialFromCalendarDate(year, month, day, date1904);
}

/** Excel's two-digit year window: 00-29 is 2000-2029, 30-99 is 1930-1999. */
function fullYear(text: string): number {
  const year = Number(text);
  if (text.length > 2) return year;
  return year < 30 ? 2000 + year : 1900 + year;
}

interface ParsedNumber {
  value: number;
  grouped: boolean;
  decimals: boolean;
  exponent: boolean;
  /** The integer part starts with a zero it does not need (007, 00.5). */
  leadingZeros: boolean;
  /** More significant digits than the 15 a number keeps. */
  long: boolean;
}

/** The decimal and thousands separators a numeral is read with. */
export interface DecimalNumeralSeparators {
  readonly decimal: string;
  /** null: digits are not grouped. */
  readonly group: string | null;
}

/** The parts of an unsigned decimal numeral, as written. */
export interface DecimalNumeralParts {
  /** The integer digits without group separators ("" for `.5`). */
  whole: string;
  /** The fraction digits ("" for `1.` and for `1`). */
  fraction: string;
  /** The decimal separator was written (`1.` and `.5`, not `1`). */
  point: boolean;
  /** The exponent digits with their sign, or undefined. */
  exponent: string | undefined;
  grouped: boolean;
}

const escapeNumeralToken = (value: string): string =>
  value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const numeralPatterns = new Map<string, RegExp>();

/**
 * The one grammar of an unsigned decimal numeral, shared by cell entry and
 * text import: digits (optionally grouped by thousands), a fraction and an
 * exponent, where either side of the decimal separator may be empty but not
 * both, as Excel reads `.5`, `1.`, `1.e3` and `0.5` alike. A grouped numeral
 * takes no exponent. Policy beyond the grammar (15 significant digits for an
 * entry, leading-zero ids for an import) belongs to the caller.
 */
export function readDecimalNumeral(
  text: string,
  separators: DecimalNumeralSeparators = { decimal: ".", group: "," },
): DecimalNumeralParts | null {
  const key = `${separators.decimal}\u0000${separators.group ?? ""}`;
  let pattern = numeralPatterns.get(key);
  if (!pattern) {
    const decimal = escapeNumeralToken(separators.decimal);
    const integer = separators.group
      ? `\\d{1,3}(?:${escapeNumeralToken(separators.group)}\\d{3})+|\\d+`
      : "\\d+";
    pattern = new RegExp(
      `^(${integer})?(?:(${decimal})(\\d*))?(?:[eE]([+-]?\\d+))?$`,
    );
    numeralPatterns.set(key, pattern);
  }
  const match = pattern.exec(text);
  if (!match) return null;
  const grouped = Boolean(
    separators.group && match[1]?.includes(separators.group),
  );
  const whole = separators.group
    ? (match[1] ?? "").split(separators.group).join("")
    : match[1] ?? "";
  const fraction = match[3] ?? "";
  if (whole === "" && fraction === "") return null;
  if (grouped && match[4] !== undefined) return null;
  return {
    whole,
    fraction,
    point: match[2] !== undefined,
    exponent: match[4],
    grouped,
  };
}

/**
 * An unsigned decimal numeral: digits (optionally grouped by thousands), a
 * fraction and an exponent. Excel keeps 15 significant digits and turns the
 * rest into zeros ("Keeping leading zeros and large numbers").
 */
function parseUnsigned(
  text: string,
  separators?: DecimalNumeralSeparators,
): ParsedNumber | null {
  const parts = readDecimalNumeral(
    separators?.group === " " ? text.replace(SPACE_GROUP_RE, " ") : text,
    separators,
  );
  if (!parts) return null;
  const grouped = parts.grouped;
  let whole = parts.whole;
  let fraction = parts.fraction;
  const digits = whole + fraction;
  const lead = digits.length - digits.replace(/^0+/, "").length;
  const long = digits.length - lead > 15;
  if (long) {
    const kept = digits.slice(0, lead + 15).padEnd(digits.length, "0");
    whole = kept.slice(0, whole.length);
    fraction = kept.slice(whole.length);
  }
  const value = Number(
    `${whole || "0"}.${fraction || "0"}e${parts.exponent ?? "0"}`,
  );
  if (!Number.isFinite(value)) return null;
  return {
    value,
    grouped,
    decimals: /[1-9]/.test(fraction) || (fraction.length > 0 && grouped),
    exponent: parts.exponent !== undefined,
    leadingZeros: parts.whole.length > 1 && parts.whole.startsWith("0"),
    long,
  };
}

/**
 * Whether Automatic Data Conversion lets a numeral become a number: a
 * leading-zero numeral (007), one of 16 or more significant digits and one
 * with an exponent (123E5) each need their own option; off, the entry stays
 * the text as written.
 */
function numeralConversionAllowed(
  parsed: Pick<ParsedNumber, "leadingZeros" | "long" | "exponent">,
  conversions: DataConversionOptions,
): boolean {
  return (
    (!parsed.leadingZeros || conversions.leadingZeros) &&
    (!parsed.long || parsed.exponent || conversions.longNumbers) &&
    (!parsed.exponent || conversions.exponent)
  );
}

function numberEntry(value: number, numberFormat?: string): SpreadsheetEntry {
  // -0 is stored as 0, as Excel does.
  const canonical = String(value === 0 ? 0 : value);
  return numberFormat
    ? { value: canonical, kind: "number", numberFormat }
    : { value: canonical, kind: "number" };
}

/** A number, percentage or currency amount, with the format it implies. */
function parseNumeric(
  text: string,
  region: EntryRegion,
  conversions: DataConversionOptions,
  typed: DecimalNumeralSeparators | null = null,
): SpreadsheetEntry | null {
  let body = text;
  let negative = false;
  // (123) is -123; the sign may sit inside the parentheses of neither form.
  const parenthesised = /^\((.*)\)$/.exec(body);
  if (parenthesised) {
    body = parenthesised[1].trim();
    negative = true;
  }
  const takeSign = () => {
    const sign = /^([+-])\s*/.exec(body);
    if (!sign) return true;
    if (parenthesised) return false;
    if (sign[1] === "-") negative = !negative;
    body = body.slice(sign[0].length);
    return true;
  };
  if (!takeSign()) return null;
  let currency = false;
  const symbol = region.currencySymbols.find((candidate) =>
    body.startsWith(candidate),
  );
  if (symbol) {
    currency = true;
    body = body.slice(symbol.length).trimStart();
    if (!takeSign()) return null;
  }
  let percent = false;
  if (body.endsWith("%")) {
    percent = true;
    body = body.slice(0, -1).trimEnd();
  }
  if (currency && percent) return null;
  const parsed = typed
    ? parseUnsigned(body, typed) ?? parseUnsigned(body, CANONICAL_POINT)
    : parseUnsigned(body);
  if (!parsed || (currency && parsed.exponent)) return null;
  if (!numeralConversionAllowed(parsed, conversions)) return null;
  const signed = negative ? -parsed.value : parsed.value;
  if (percent) {
    return numberEntry(
      Number((signed / 100).toPrecision(15)),
      parsed.decimals ? "0.00%" : "0%",
    );
  }
  if (currency) {
    return numberEntry(signed, region.currencyFormat(parsed.decimals));
  }
  if (parsed.exponent) return numberEntry(signed, "0.00E+00");
  if (parsed.grouped) {
    return numberEntry(signed, parsed.decimals ? "#,##0.00" : "#,##0");
  }
  return numberEntry(signed);
}

/** 1 1/2, 0 3/4, -2 1/8: a whole number, a space and a fraction. */
function parseFraction(text: string): SpreadsheetEntry | null {
  const match = /^([+-]?)(\d+)\s+(\d+)\/(\d+)$/.exec(text);
  if (!match) return null;
  const denominator = Number(match[4]);
  if (denominator === 0) return null;
  const magnitude = Number(match[2]) + Number(match[3]) / denominator;
  const width = Math.min(3, match[4].replace(/^0+/, "").length);
  const question = "?".repeat(width);
  return numberEntry(
    match[1] === "-" ? -magnitude : magnitude,
    `# ${question}/${question}`,
  );
}

interface ParsedTime {
  fraction: number;
  format: string;
}

/** 10:30, 10:30:15, 9:00 p, 9:00 PM, 25:00. */
function parseTime(text: string): ParsedTime | null {
  const match =
    /^(\d{1,4}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([AaPp])(?:[Mm])?)?$/.exec(
      text,
    );
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = match[3] === undefined ? 0 : Number(match[3]);
  if (minutes > 59 || seconds >= 60) return null;
  const meridiem = match[4]?.toLowerCase();
  if (meridiem) {
    if (hours > 12) return null;
    if (meridiem === "p" && hours < 12) hours += 12;
    if (meridiem === "a" && hours === 12) hours = 0;
  } else if (hours > 9999) {
    return null;
  }
  const withSeconds = match[3] !== undefined;
  const format = meridiem
    ? withSeconds
      ? "h:mm:ss AM/PM"
      : "h:mm AM/PM"
    : hours >= 24
      ? "[h]:mm:ss"
      : withSeconds
        ? "h:mm:ss"
        : "h:mm";
  return {
    fraction: (hours * 3600 + minutes * 60 + seconds) / 86400,
    format,
  };
}

interface ParsedDate {
  serial: number;
  format: string;
}

/**
 * "Convert continuous letters and numbers to a date" (Automatic Data
 * Conversion), as MEASURED in Excel 365 for Mac through its object model:
 * - month name then digits: a day of this year when the month has that day
 *   (JAN1, SEPT2, MARCH1, DEC12 -> d-mmm), else a month of the year the
 *   digits name (FEB30 -> Feb-1930, JAN0 -> Jan-2000, JAN2024 -> mmm-yy);
 *   three digits (JAN100) stay text;
 * - digits then month name: that day of this year (1JAN, 31DEC -> d-mmm);
 *   no such day (32JAN) stays text;
 * - day, month name, year: 1JAN2024, 12JAN99 -> d-mmm-yy.
 * Month names are the three-letter forms, the full names and SEPT; any case.
 */
export function parseLetterDigitDate(
  text: string,
  today: Date,
  date1904: ExcelDateSystemArg = false,
): ParsedDate | null {
  const thisYear = today.getFullYear();
  const serialOf = (year: number, month: number, day: number) =>
    dateSerial(year, month, day, date1904);
  const withFormat = (serial: number | null, format: string) =>
    serial === null ? null : { serial, format };
  let match = /^([A-Za-z]{3,9})(\d{1,4})$/.exec(text);
  if (match) {
    const month = monthFromName(match[1]);
    if (!month) return null;
    const digits = match[2];
    if (digits.length <= 2) {
      const asDay = Number(digits) >= 1 ? serialOf(thisYear, month, Number(digits)) : null;
      if (asDay !== null) return { serial: asDay, format: "d-mmm" };
      return withFormat(serialOf(fullYear(digits), month, 1), "mmm-yy");
    }
    return digits.length === 4 ? withFormat(serialOf(Number(digits), month, 1), "mmm-yy") : null;
  }
  match = /^(\d{1,2})([A-Za-z]{3,9})(\d{4}|\d{2})?$/.exec(text);
  if (match) {
    const month = monthFromName(match[2]);
    if (!month) return null;
    return match[3]
      ? withFormat(serialOf(fullYear(match[3]), month, Number(match[1])), "d-mmm-yy")
      : withFormat(serialOf(thisYear, month, Number(match[1])), "d-mmm");
  }
  return null;
}

/** A calendar date in the region's forms, with the format it implies. */
function parseDate(
  text: string,
  region: EntryRegion,
  today: Date,
  date1904: ExcelDateSystemArg = false,
): ParsedDate | null {
  const thisYear = today.getFullYear();
  const serialOf = (year: number, month: number, day: number) =>
    dateSerial(year, month, day, date1904);
  const withFormat = (serial: number | null, format: string) =>
    serial === null ? null : { serial, format };
  let match = /^(\d{4})([/-])(\d{1,2})\2(\d{1,2})$/.exec(text);
  if (match) {
    return withFormat(
      serialOf(Number(match[1]), Number(match[3]), Number(match[4])),
      SYSTEM_SHORT_DATE,
    );
  }
  if (region.zh) {
    match = /^(\d{4})年(\d{1,2})月(\d{1,2})日$/.exec(text);
    if (match) {
      return withFormat(
        serialOf(Number(match[1]), Number(match[2]), Number(match[3])),
        'yyyy"年"m"月"d"日"',
      );
    }
    match = /^(\d{4})年(\d{1,2})月$/.exec(text);
    if (match) {
      return withFormat(
        serialOf(Number(match[1]), Number(match[2]), 1),
        'yyyy"年"m"月"',
      );
    }
    match = /^(\d{1,2})月(\d{1,2})日$/.exec(text);
    if (match) {
      return withFormat(
        serialOf(thisYear, Number(match[1]), Number(match[2])),
        region.monthDayFormat,
      );
    }
  } else {
    match = /^(\d{1,2})([/-])(\d{1,2})\2(\d{4}|\d{2})$/.exec(text);
    if (match) {
      return withFormat(
        serialOf(fullYear(match[4]), Number(match[1]), Number(match[3])),
        SYSTEM_SHORT_DATE,
      );
    }
    // 21-Sep, 21 Sep 2026, 21-September-26
    match = /^(\d{1,2})[-\s/]([A-Za-z]{3,9})(?:[-\s/]+(\d{4}|\d{2}))?$/.exec(
      text,
    );
    if (match && monthFromName(match[2])) {
      return withFormat(
        serialOf(
          match[3] ? fullYear(match[3]) : thisYear,
          monthFromName(match[2])!,
          Number(match[1]),
        ),
        match[3] ? "d-mmm-yy" : "d-mmm",
      );
    }
    // Sep 21, Sep 21, 2026
    match = /^([A-Za-z]{3,9})[-\s](\d{1,2})(?:,?\s+(\d{4}))?$/.exec(text);
    if (match && monthFromName(match[1])) {
      return withFormat(
        serialOf(
          match[3] ? Number(match[3]) : thisYear,
          monthFromName(match[1])!,
          Number(match[2]),
        ),
        match[3] ? "d-mmm-yy" : "d-mmm",
      );
    }
  }
  // 9/21, 3-5, 1/2: a month and day of this year.
  match = /^(\d{1,2})([/-])(\d{1,2})$/.exec(text);
  if (match) {
    return withFormat(
      serialOf(thisYear, Number(match[1]), Number(match[3])),
      region.monthDayFormat,
    );
  }
  return null;
}

/**
 * zh-CN's clock forms: 3时30分, 15时30分20秒, 下午3时30分 — with the zh-CN
 * built-ins they imply (h"时"mm"分" … 上午/下午h"时"mm"分"ss"秒"). MEASURED in
 * Excel (zh-CN): 下午3时30分 is 15:30 and 3时30分 is 3:30, while 下午3:30 and
 * 下午3时 (no minutes) stay text, so the 时/分 pair is required (XLS-A092).
 */
function parseChineseTime(text: string): ParsedTime | null {
  const match = /^(上午|下午)?(\d{1,2})时(\d{1,2})分(?:(\d{1,2})秒)?$/.exec(text);
  if (!match) return null;
  let hours = Number(match[2]);
  const minutes = Number(match[3]);
  const seconds = match[4] === undefined ? 0 : Number(match[4]);
  if (minutes > 59 || seconds > 59) return null;
  const meridiem = match[1];
  if (meridiem) {
    if (hours > 12) return null;
    if (meridiem === "下午" && hours < 12) hours += 12;
    if (meridiem === "上午" && hours === 12) hours = 0;
  } else if (hours > 23) {
    return null;
  }
  const clock = match[4] === undefined ? 'h"时"mm"分"' : 'h"时"mm"分"ss"秒"';
  return {
    fraction: (hours * 3600 + minutes * 60 + seconds) / 86400,
    format: meridiem ? `上午/下午${clock}` : clock,
  };
}

/** A date, a time, or a date and a time. */
function parseDateTime(
  text: string,
  region: EntryRegion,
  today: Date,
  date1904: ExcelDateSystemArg = false,
  conversions: DataConversionOptions = readDataConversionOptions(),
): SpreadsheetEntry | null {
  const date =
    parseDate(text, region, today, date1904) ??
    (conversions.letterDates ? parseLetterDigitDate(text, today, date1904) : null);
  if (date) return numberEntry(date.serial, date.format);
  const time = parseTime(text) ?? (region.zh ? parseChineseTime(text) : null);
  if (time) return numberEntry(Number(time.fraction.toPrecision(15)), time.format);
  const split = /^(\S+)\s+(\S+(?:\s*[AaPp][Mm]?)?)$/.exec(text);
  if (!split) return null;
  const datePart = parseDate(split[1], region, today, date1904);
  const timePart = datePart ? parseTime(split[2]) : null;
  if (!datePart || !timePart || timePart.fraction >= 1) return null;
  return numberEntry(
    Number((datePart.serial + timePart.fraction).toPrecision(15)),
    SYSTEM_SHORT_DATE_TIME,
  );
}

/** A plain number typed into a percentage cell (`percentFormat`). */
function percentEntry(entry: SpreadsheetEntry): SpreadsheetEntry {
  const number = Number(entry.value);
  if (entry.kind !== "number" || entry.numberFormat !== undefined || !Number.isFinite(number) || Math.abs(number) < 1) {
    return entry;
  }
  return { ...entry, value: String(Number((number / 100).toPrecision(15))) };
}

/** Read one entry the way Excel's cell entry does. */
export function parseSpreadsheetEntry(
  raw: string,
  options: SpreadsheetEntryOptions,
): SpreadsheetEntry {
  if (raw.startsWith("'")) {
    const text = raw.slice(1);
    return text === ""
      ? { value: "", kind: "text" }
      : { value: text, kind: "text", quotePrefix: true };
  }
  const text = raw.trim();
  if (options.textFormat || text === "" || text.startsWith("=")) {
    return { value: raw, kind: "text" };
  }
  const upper = text.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return { value: upper, kind: "text" };
  }
  const error = ERROR_LITERALS.find((code) => code === upper);
  if (error) return { value: error, kind: "text" };
  const region = entryRegion(options.locale);
  const conversions = options.conversions ?? readDataConversionOptions();
  const typed = options.typedNumerals ? typedNumeralSeparators(options.locale) : null;
  const numeric = parseNumeric(text, region, conversions, typed) ?? parseFraction(text);
  if (numeric) return options.percentFormat ? percentEntry(numeric) : numeric;
  // "A space before you enter a number" keeps a date or time as text.
  if (!/^\s/.test(raw)) {
    const dated = parseDateTime(
      text,
      region,
      options.today ?? new Date(),
      options.date1904,
      conversions,
    );
    if (dated) return dated;
  }
  return { value: raw, kind: "text" };
}

/**
 * Whether `text` typed into a cell is stored as exactly that text: not a
 * formula, a number, a date, a logical or an error. A writer that must keep a
 * string a string (Copy Value from Above, a scenario's value) puts an
 * apostrophe in front otherwise. The parsed kind alone does not say it: a
 * typed TRUE or #N/A is stored with kind "text" and reads as a logical or an
 * error (live audit S-041).
 */
export function entryKeepsText(text: string, options: SpreadsheetEntryOptions): boolean {
  const entry = parseSpreadsheetEntry(text, options);
  if (entry.kind !== "text" || entry.quotePrefix === true || entry.value !== text) return false;
  const upper = text.trim().toUpperCase();
  return (
    !upper.startsWith("=") &&
    upper !== "TRUE" &&
    upper !== "FALSE" &&
    !ERROR_LITERALS.includes(upper)
  );
}

/** A run of digits longer than the 15 significant digits a number keeps. */
const LONG_DIGIT_TEXT = /^\d{16,}$/;

/**
 * Whether the cell holds TEXT whatever its characters look like — Excel's own
 * cell type, not a reading of the value. Two entries make one: a leading
 * apostrophe (kept as the cell's quote prefix) and an entry typed into a Text
 * (@) cell. Such a cell is a string to the engine even when it reads as "123",
 * "TRUE" or "#N/A": SUM ignores it, ="123" is TRUE against it, and only
 * arithmetic coerces it (Microsoft Support, "Enter data manually in worksheet
 * cells", and "Calculation operators and precedence in Excel" for the
 * coercion).
 *
 * The @ format alone is not enough: Excel keeps a number that was already in
 * the cell a number when the format is applied ("Numbers that you entered
 * before you applied the Text format to the cells must be entered again"), and
 * only the re-entry gives the cell its text kind.
 */
export function spreadsheetCellIsText(
  cell:
    | {
        value?: string;
        formula?: string;
        kind?: string;
        stringValue?: unknown;
        style?: { numberFormat?: unknown; quotePrefix?: unknown };
      }
    | undefined
    | null,
): boolean {
  if (!cell || cell.formula) return false;
  // More than 15 digits stored as text can only be text: Excel keeps 15
  // significant digits, so read as a number it becomes other digits. An ID
  // card or account number pasted or written as text showed whole in the grid
  // (flagged as a number stored as text) while formulas read the rounded
  // number, and MID(B2,17,1) of every ID gave "0" (sheet campaign K07,
  
  if (cell.kind === "text" && typeof cell.value === "string" && LONG_DIGIT_TEXT.test(cell.value)) return true;
  // A string the workbook file typed as one (SpreadsheetML t="s" or
  // inlineStr) whose characters read as a logical, error or number: the
  // file's type wins, as it does in Excel (live audit B312).
  if (cell.stringValue === true && cell.kind !== "number") return true;
  // The prefix is the identity of an entry typed after an apostrophe; a number
  // that received it with a pasted or painted format is still the number it was
  // (live audit B086).
  if (cell.style?.quotePrefix === true && cell.kind !== "number") return true;
  if (cell.kind !== "text" || !cell.style?.numberFormat) return false;
  return numberFormatFamily(cell.style.numberFormat) === "text";
}

/**
 * The style a format transfer (Paste Formats, Format Painter) leaves on a
 * target cell: the source's look, with the target's own text identity. The
 * quote prefix is not a format but the record that the entry was typed as
 * text, so it neither travels with the source's style nor is lost from the
 * target's (Excel keeps a cell's value type through a format paste; live
 * audit B086).
 */
export function styleWithOwnTextIdentity<T extends { quotePrefix?: unknown }>(
  transferred: T | undefined,
  own: { quotePrefix?: unknown; [key: string]: unknown } | undefined,
): T | undefined {
  const keep = own?.quotePrefix === true;
  if (transferred === undefined) {
    return keep ? ({ quotePrefix: true } as unknown as T) : undefined;
  }
  if (transferred.quotePrefix === undefined && !keep) return transferred;
  const next = { ...transferred };
  if (keep) next.quotePrefix = true;
  else delete next.quotePrefix;
  return next;
}

/** A number as its lossless edit text: 15 significant digits, no exponent
 *  below 1e21, "E" for the exponent Excel writes. */
function formatEditNumber(value: number): string {
  if (value === 0) return "0";
  const abs = Math.abs(value);
  if (abs >= 1e21 || abs < 1e-6) return formatGeneralNumber(value);
  return String(Number(value.toPrecision(15)));
}

/**
 * The text an edit of the cell starts from (F2, the formula bar): Excel shows
 * a percentage as 50%, a date in the regional short date, a time with its
 * seconds, and every other number as its plain value (a currency amount as
 * 1000). A quote-prefixed text shows its apostrophe again.
 *
 * `typedNumerals` (the cell editor and the formula bar, whose text is entered
 * back with the same option): a number's decimal point is the interface
 * language's, as Excel's formula bar shows 1234,5 in German, so editing a cell
 * and pressing Enter keeps its value (1.234 must not come back as 1234).
 */
export function spreadsheetEntryEditText(
  cell:
    | {
        value?: string;
        formula?: string;
        kind?: string;
        style?: { numberFormat?: unknown; quotePrefix?: unknown };
      }
    | undefined,
  locale: string,
  options?: { date1904?: ExcelDateSystemArg; typedNumerals?: boolean },
): string {
  const value = cell?.value ?? "";
  if (!cell || cell.formula || value === "") return value;
  const date1904 = options?.date1904;
  if (cell.style?.quotePrefix === true && cell.kind !== "number") {
    return `'${value}`;
  }
  if (cell.kind !== "number") return value;
  const number = Number(value);
  const typed = options?.typedNumerals ? typedNumeralSeparators(locale) : null;
  // A canonical numeral ("1234.5", "1.5e-7") with the region's decimal point.
  const numeral = (text: string) =>
    typed && Number.isFinite(Number(text)) ? text.replace(".", typed.decimal) : text;
  const code = normalizeNumberFormatCode(cell.style?.numberFormat);
  if (!code || !Number.isFinite(number)) return numeral(value);
  const family = numberFormatFamily(code);
  if (family === "percentage") {
    // The edit text must round-trip: 1.23456789012345% keeps its 15
    // significant digits, where General's display rounding would cut it to
    // ten decimals and a re-entry would lose the rest (live audit B097).
    return `${numeral(formatEditNumber(Number((number * 100).toPrecision(15))))}%`;
  }
  if (family !== "date" && family !== "time") return numeral(value);
  // Serial 0 ("1900-01-00") and durations of a day or more stay numbers:
  // their calendar reading would not come back as the same value. Serials
  // 1..60 read back through Excel's 1900 calendar, 1900-02-29 included. In
  // the 1904 system serial 0 is a real day, 1904-01-01.
  if (number < 0 || number >= excelMaxDateSerial(date1904) + 1) return numeral(value);
  const region = entryRegion(locale);
  const whole = Math.floor(number);
  if (family === "time") {
    return number < 1 ? applyNumberFormat(number, region.longTime) : numeral(value);
  }
  if (whole < excelMinDateSerial(date1904)) return numeral(value);
  const dateSystem = { date1904 };
  const date = applyNumberFormat(whole, region.shortDate, dateSystem);
  return number === whole
    ? date
    : `${date} ${applyNumberFormat(number, region.longTime, dateSystem)}`;
}
