/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Parse delimited text (CSV / TSV) into native spreadsheet page content.
 *
 * The parser follows RFC 4180 by default: qualified fields may contain the
 * delimiter, embedded newlines, and a doubled qualifier escape. The preview
 * may explicitly select apostrophe or no qualifier through the same authority
 * used by Text to Columns. Numeric-looking cells become numbers, under
 * Excel's Automatic Data Conversion options (lib/spreadsheet/data-conversion-
 * options): a leading-zero id (00123), a numeral of 16+ digits, one with an
 * exponent (123E5) and a run of letters and digits (JAN1) convert only while
 * their option is on, and each converted field says which option it needed,
 * so opening a file can first ask (Excel's notice). Oversized input fails
 * closed while scanning; it is never silently truncated into a different
 * document.
 */

import {
  PageImportProjectionBudget,
  PageImportProjectionBudgetError,
} from "@mrnote/docs/web/lib/document-import/pageImportProjectionBudget.ts";
import { SPREADSHEET_CONTENT_TYPE } from "@mrnote/core/web/lib/native-page-content-types.ts";
import type { NativeSpreadsheetContent } from "@mrnote/core/web/lib/native-page-schemas.ts";
import {
  csvImportTextQualifierCharacter,
  type CsvImportTextQualifier,
} from "@mrnote/core/web/lib/spreadsheet/csv.ts";
import {
  excelSerialFromCalendarDate,
  excelSystemSerial,
  type SpreadsheetDateSystemOptions,
} from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";
import {
  parseLetterDigitDate,
  readDecimalNumeral,
  type DecimalNumeralSeparators,
} from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import {
  DEFAULT_DATA_CONVERSION_OPTIONS,
  readDataConversionOptions,
  type DataConversionKind,
  type DataConversionOptions,
} from "@mrnote/core/web/lib/spreadsheet/data-conversion-options.ts";
import {
  normalizeFixedWidthBreaks,
  splitFixedWidthFields,
} from "@mrnote/core/web/lib/spreadsheet/text-to-columns.ts";
import { excelMinDateSerial } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

// Excel's own grid. The populated-cell and text budgets below are what bound
// the browser's work; a file past them is imported on the server (as tiles).
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
// Mirrors the backend per-sheet cell ceiling (MAX_SPREADSHEET_SHEET_CELLS,
// docs/spreadsheet-page-limits.md).
const MAX_CELLS = 200_000;
const MAX_TEXT_CHARS = 1_000_000;

export type DelimitedImportLimitCode =
  | "csv_row_limit_exceeded"
  | "csv_column_limit_exceeded"
  | "csv_cell_limit_exceeded"
  | "csv_node_limit_exceeded"
  | "csv_text_limit_exceeded"
  | "csv_malformed";

export class DelimitedImportLimitError extends Error {
  constructor(
    readonly code: DelimitedImportLimitCode,
    message: string,
  ) {
    super(message);
    this.name = "DelimitedImportLimitError";
  }
}

interface ScanDelimitedOptions {
  onField: (row: number, column: number, value: string) => void;
  maxMaterializedFields?: number;
  textQualifier?: CsvImportTextQualifier;
  consecutiveDelimiters?: boolean;
}

function scanDelimited(
  input: string,
  delimiter: string,
  options: ScanDelimitedOptions,
): void {
  const qualifier = csvImportTextQualifierCharacter(
    options.textQualifier ?? "double",
  );
  if (
    delimiter.length !== 1 ||
    delimiter === '"' ||
    /[\r\n]/.test(delimiter) ||
    delimiter === qualifier
  ) {
    throw new DelimitedImportLimitError(
      "csv_malformed",
      delimiter === qualifier
        ? "The delimiter and text qualifier must be different characters."
        : "The delimiter must be one non-newline character.",
    );
  }
  const start = input.charCodeAt(0) === 0xfeff ? 1 : 0;
  let row = 0;
  let column = 0;
  let field = "";
  let fieldState: "start" | "unquoted" | "quoted" | "after_quote" = "start";
  let endedWithRowSeparator = false;
  let nonemptyCells = 0;
  let materializedFields = 0;
  let textChars = 0;

  const append = (value: string) => {
    textChars += value.length;
    if (textChars > MAX_TEXT_CHARS) {
      throw new DelimitedImportLimitError(
        "csv_text_limit_exceeded",
        "Delimited text exceeds the import text budget.",
      );
    }
    field += value;
  };

  const emitField = () => {
    if (row >= MAX_ROWS) {
      throw new DelimitedImportLimitError(
        "csv_row_limit_exceeded",
        `Delimited files may contain at most ${MAX_ROWS} rows.`,
      );
    }
    if (column >= MAX_COLS) {
      throw new DelimitedImportLimitError(
        "csv_column_limit_exceeded",
        `Delimited files may contain at most ${MAX_COLS} columns.`,
      );
    }
    materializedFields += 1;
    if (
      options.maxMaterializedFields !== undefined &&
      materializedFields > options.maxMaterializedFields
    ) {
      throw new DelimitedImportLimitError(
        "csv_cell_limit_exceeded",
        "Delimited input contains too many materialized fields.",
      );
    }
    if (field !== "") {
      nonemptyCells += 1;
      if (nonemptyCells > MAX_CELLS) {
        throw new DelimitedImportLimitError(
          "csv_cell_limit_exceeded",
          `Delimited files may contain at most ${MAX_CELLS} populated cells.`,
        );
      }
    }
    options.onField(row, column, field);
    field = "";
    fieldState = "start";
    column += 1;
  };

  const emitRow = () => {
    emitField();
    row += 1;
    column = 0;
    endedWithRowSeparator = true;
  };

  for (let index = start; index < input.length; index += 1) {
    const ch = input[index];
    endedWithRowSeparator = false;
    if (fieldState === "quoted") {
      if (qualifier && ch === qualifier) {
        if (input[index + 1] === qualifier) {
          append(qualifier);
          index += 1;
        } else {
          fieldState = "after_quote";
        }
      } else {
        // A line break inside a qualified field is the field's own text,
        // kept as written (CR, LF or CRLF): a Text column keeps it verbatim,
        // and the other kinds read it as one line break
        // (classifyDelimitedImportCell; user audit S-116).
        append(ch);
      }
    } else if (fieldState === "after_quote") {
      // RFC 4180 permits only a delimiter, record separator, or EOF after the
      // closing quote. Accepting ordinary bytes here silently rewrites malformed
      // source (for example `"a"b` became `ab`).
      if (ch === delimiter) {
        emitField();
        if (options.consecutiveDelimiters) {
          while (input[index + 1] === delimiter) index += 1;
        }
      } else if (ch === "\n" || ch === "\r") {
        emitRow();
        if (ch === "\r" && input[index + 1] === "\n") index += 1;
      } else {
        // Text after a closing quote joins the field, as Excel opens such a
        // file (`"a"b` reads `ab`) instead of refusing it.
        append(ch);
        fieldState = "unquoted";
      }
    } else if (qualifier && ch === qualifier) {
      if (fieldState !== "start") {
        // A quote inside an unquoted field is the field's own text (Excel).
        append(ch);
        continue;
      }
      fieldState = "quoted";
    } else if (ch === delimiter) {
      emitField();
      if (options.consecutiveDelimiters) {
        while (input[index + 1] === delimiter) index += 1;
      }
    } else if (ch === "\n" || ch === "\r") {
      emitRow();
      if (ch === "\r" && input[index + 1] === "\n") index += 1;
    } else {
      append(ch);
      fieldState = "unquoted";
    }
  }
  // An unterminated qualified field runs to the end of the file (Excel).
  if (!endedWithRowSeparator || input.length === start) {
    emitField();
  }
}

/** Split delimited text into a grid of string cells. */
export function parseDelimited(
  input: string,
  delimiter = ",",
  textQualifier: CsvImportTextQualifier = "double",
  consecutiveDelimiters = false,
): string[][] {
  const rows: string[][] = [];
  scanDelimited(input, delimiter, {
    // `parseDelimited` exposes a materialized grid, so cap even empty fields.
    // The page converter below streams fields directly and has no such array.
    maxMaterializedFields: MAX_CELLS,
    textQualifier,
    consecutiveDelimiters,
    onField: (row, _column, value) => {
      (rows[row] ||= []).push(value);
    },
  });
  return rows;
}

/**
 * Parse a fixed-width text file through the same field-boundary authority as
 * Text to Columns. Physical CR/LF records are preserved, trailing record
 * separators do not invent an extra row, and the same import budgets as the
 * delimited scanner fail closed before a large materialized preview is built.
 */
export function parseFixedWidthText(
  input: string,
  breaks: readonly number[],
): string[][] {
  const cuts = normalizeFixedWidthBreaks(breaks);
  const columnCount = cuts.length + 1;
  if (columnCount > MAX_COLS) {
    throw new DelimitedImportLimitError(
      "csv_column_limit_exceeded",
      `Fixed-width files may contain at most ${MAX_COLS} columns.`,
    );
  }
  const source = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  const records = source.split(/\r\n|\r|\n/);
  if (/[\r\n]$/.test(source)) records.pop();
  if (records.length > MAX_ROWS) {
    throw new DelimitedImportLimitError(
      "csv_row_limit_exceeded",
      `Fixed-width files may contain at most ${MAX_ROWS} rows.`,
    );
  }
  if (records.length * columnCount > MAX_CELLS) {
    throw new DelimitedImportLimitError(
      "csv_cell_limit_exceeded",
      `Fixed-width files may contain at most ${MAX_CELLS} materialized fields.`,
    );
  }
  let textChars = 0;
  return records.map((record) => {
    textChars += record.length;
    if (textChars > MAX_TEXT_CHARS) {
      throw new DelimitedImportLimitError(
        "csv_text_limit_exceeded",
        "Fixed-width text exceeds the import text budget.",
      );
    }
    return splitFixedWidthFields(record, cuts);
  });
}

// Spreadsheet date serials follow Excel's 1900 date system (1900-01-01 = 1,
// fictitious 1900-02-29 = 60), shared with the formula engine and the entry
// parser through excel-1900-calendar so an imported date works in formulas.

export interface DelimitedImportCellValue {
  readonly value: string;
  readonly kind: "text" | "number";
  readonly numberFormat?:
    | "yyyy-mm-dd"
    | "@"
    | "0%"
    | "0.00%"
    | "0.00E+00"
    | "d-mmm"
    | "mmm-yy"
    | "d-mmm-yy";
  /** The Automatic Data Conversion option this reading needed (00123 ->
   *  123 needs leadingZeros): what opening the file reports and asks about. */
  readonly conversion?: DataConversionKind;
  /** Explicit Text disposition: never reinterpret a leading '=' as a formula. */
  readonly literalText?: boolean;
  /** Skip disposition: omit this source field from the projected output row. */
  readonly skip?: boolean;
}

export type DelimitedImportColumnDisposition =
  | "general"
  | "text"
  | "date"
  | "skip";
export type DelimitedImportDateOrder = "locale" | "mdy" | "dmy" | "ymd";
export type DelimitedImportDecimalSeparator = "locale" | "." | ",";
export type DelimitedImportThousandsSeparator =
  | "locale"
  | ","
  | "."
  | "space"
  | "none";

/** Explicit number parsing choices from the Text Import Wizard's Advanced step. */
export interface DelimitedImportNumberOptions {
  readonly decimalSeparator?: DelimitedImportDecimalSeparator;
  readonly thousandsSeparator?: DelimitedImportThousandsSeparator;
  readonly trailingMinus?: boolean;
}

function calendarDateSerial(
  year: number,
  month: number,
  day: number,
): number | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const serial = excelSerialFromCalendarDate(year, month, day);
  if (serial === null) return null;
  return serial >= 1 ? serial : null;
}









export function isoDateSerial(raw: string): number | null {
  const m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(raw.trim());
  if (!m) return null;
  return calendarDateSerial(Number(m[1]), Number(m[2]), Number(m[3]));
}

interface ImportNumber {
  readonly value: number;
  /** The option the reading needed, the first in Excel's list order. */
  readonly conversion?: DataConversionKind;
  readonly exponent: boolean;
}

/**
 * The number a signed numeral stands for, read with the cell-entry grammar
 * (`readDecimalNumeral`: `.5`, `1.`, `-.5`, `1.e3` are numbers, as typing them
 * stores them) under Automatic Data Conversion: a leading-zero integer part
 * (007, 012345) is a number only under "Remove leading zeros", a numeral of
 * more than 15 significant digits only under "Keep first 15 digits" (its
 * digits past the 15th become zeros, as Excel keeps it) and one with an
 * exponent only under the letter "E" option; otherwise it stays text, as the
 * source file spells it.
 */
function importNumeral(
  source: string,
  separators: DecimalNumeralSeparators,
  conversions: DataConversionOptions,
): ImportNumber | null {
  const sign = /^[+-]/.exec(source)?.[0] ?? "";
  const parts = readDecimalNumeral(source.slice(sign.length), separators);
  if (!parts) return null;
  const leadingZeros = parts.whole.length > 1 && parts.whole.startsWith("0");
  const exponent = parts.exponent !== undefined;
  let whole = parts.whole;
  let fraction = parts.fraction;
  const digits = whole + fraction;
  const lead = digits.length - digits.replace(/^0+/, "").length;
  const long = !exponent && digits.length - lead > 15;
  if (leadingZeros && !conversions.leadingZeros) return null;
  if (long && !conversions.longNumbers) return null;
  if (exponent && !conversions.exponent) return null;
  if (long) {
    const kept = digits.slice(0, lead + 15).padEnd(digits.length, "0");
    whole = kept.slice(0, whole.length);
    fraction = kept.slice(whole.length);
  }
  const value = Number(
    `${sign === "-" ? "-" : ""}${whole || "0"}.${fraction || "0"}e${parts.exponent ?? "0"}`,
  );
  if (!Number.isFinite(value)) return null;
  const conversion: DataConversionKind | undefined = leadingZeros
    ? "leadingZeros"
    : long
      ? "longNumbers"
      : exponent
        ? "exponent"
        : undefined;
  return { value: value === 0 ? 0 : value, exponent, ...(conversion ? { conversion } : {}) };
}

const PLAIN_NUMERAL: DecimalNumeralSeparators = { decimal: ".", group: null };

/** Whether a raw cell is stored as a number under the user's Automatic Data
 *  Conversion options. */
export function looksNumeric(raw: string): boolean {
  return importNumeral(raw.trim(), PLAIN_NUMERAL, readDataConversionOptions()) !== null;
}

function localeDateSerial(raw: string, locale: string): number | null {
  const trimmed = raw.trim();
  const normalizedLocale = locale.toLowerCase();
  const chinese = /^(\d{4})年(\d{1,2})月(\d{1,2})日$/.exec(trimmed);
  if (chinese && normalizedLocale.startsWith("zh")) {
    return calendarDateSerial(
      Number(chinese[1]),
      Number(chinese[2]),
      Number(chinese[3]),
    );
  }
  const short = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(trimmed);
  if (!short) return null;
  const first = Number(short[1]);
  const second = Number(short[2]);
  const year = Number(short[3]);
  if (normalizedLocale === "en" || normalizedLocale.startsWith("en-us")) {
    return calendarDateSerial(year, first, second);
  }
  if (normalizedLocale.startsWith("en-gb")) {
    return calendarDateSerial(year, second, first);
  }
  return null;
}

function orderedDateSerial(
  raw: string,
  order: Exclude<DelimitedImportDateOrder, "locale">,
): number | null {
  const parts = /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})$/.exec(raw.trim());
  if (!parts) return null;
  if (order === "ymd") {
    if (parts[1].length !== 4) return null;
    return calendarDateSerial(
      Number(parts[1]),
      Number(parts[2]),
      Number(parts[3]),
    );
  }
  if (parts[3].length !== 4) return null;
  const first = Number(parts[1]);
  const second = Number(parts[2]);
  return calendarDateSerial(
    Number(parts[3]),
    order === "mdy" ? first : second,
    order === "mdy" ? second : first,
  );
}

type DelimitedImportNumberSeparators = DecimalNumeralSeparators;

const normalizeNumberWhitespace = (value: string): string =>
  value.replace(/[\u00a0\u202f]/g, " ");

function localeNumberSeparators(
  locale: string,
): DelimitedImportNumberSeparators | null {
  let parts: Intl.NumberFormatPart[];
  try {
    parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
  } catch {
    return null;
  }
  return {
    decimal: normalizeNumberWhitespace(
      parts.find((part) => part.type === "decimal")?.value ?? ".",
    ),
    group:
      normalizeNumberWhitespace(
        parts.find((part) => part.type === "group")?.value ?? "",
      ) || null,
  };
}

function resolveDelimitedImportNumberSeparators(
  locale: string | undefined,
  options: DelimitedImportNumberOptions,
): DelimitedImportNumberSeparators | null {
  const decimalChoice = options.decimalSeparator ?? "locale";
  const groupChoice = options.thousandsSeparator ?? "locale";
  const needsLocale = decimalChoice === "locale" || groupChoice === "locale";
  const localeSeparators =
    needsLocale && locale ? localeNumberSeparators(locale) : null;
  if (needsLocale && !localeSeparators) return null;
  const decimal =
    decimalChoice === "locale" ? localeSeparators!.decimal : decimalChoice;
  const group =
    groupChoice === "locale"
      ? localeSeparators!.group
      : groupChoice === "none"
        ? null
        : groupChoice === "space"
          ? " "
          : groupChoice;
  return !decimal || group === decimal ? null : { decimal, group };
}

/** Invalid separator pairs fail closed instead of silently reclassifying data. */
export function delimitedImportNumberOptionsAreValid(
  locale: string,
  options: DelimitedImportNumberOptions,
): boolean {
  return resolveDelimitedImportNumberSeparators(locale, options) !== null;
}

function configuredNumberValue(
  raw: string,
  locale: string | undefined,
  options: DelimitedImportNumberOptions,
  conversions: DataConversionOptions,
): ImportNumber | null {
  const separators = resolveDelimitedImportNumberSeparators(locale, options);
  if (!separators) return null;
  let source = normalizeNumberWhitespace(raw.trim());
  if (options.trailingMinus && source.endsWith("-")) {
    const magnitude = source.slice(0, -1);
    if (!magnitude || /^[+-]/.test(magnitude)) return null;
    source = `-${magnitude}`;
  }
  return importNumeral(source, separators, conversions);
}

function importNumber(
  raw: string,
  locale: string | undefined,
  numberOptions: DelimitedImportNumberOptions | undefined,
  conversions: DataConversionOptions,
): ImportNumber | null {
  if (numberOptions) return configuredNumberValue(raw, locale, numberOptions, conversions);
  return (
    importNumeral(raw.trim(), PLAIN_NUMERAL, conversions) ??
    (locale
      ? configuredNumberValue(
          raw,
          locale,
          { decimalSeparator: "locale", thousandsSeparator: "locale" },
          conversions,
        )
      : null)
  );
}

/** `50%` / `-12.5 %` under General: the fraction it stands for, with the
 *  percentage format typing it would give (0.00% when the text has decimals). */
function importPercentValue(
  raw: string,
  locale: string | undefined,
  numberOptions: DelimitedImportNumberOptions | undefined,
  conversions: DataConversionOptions,
): DelimitedImportCellValue | null {
  const trimmed = raw.trim();
  if (!trimmed.endsWith("%")) return null;
  const body = trimmed.slice(0, -1).trimEnd();
  if (!body || body.endsWith("%")) return null;
  const parsed = importNumber(body, locale, numberOptions, conversions);
  if (parsed === null) return null;
  const numeric = String(parsed.value);
  const value = parsed.value;
  if (!Number.isFinite(value)) return null;
  const decimals = numeric.includes(".") || /[eE]/.test(numeric);
  return {
    value: String(Number((value / 100).toPrecision(15))),
    kind: "number",
    numberFormat: decimals ? "0.00%" : "0%",
    ...(parsed.conversion ? { conversion: parsed.conversion } : {}),
  };
}

/**
 * Classify one imported field through the same value authority used by native
 * page import and the live worksheet writer. ISO dates are always safe;
 * ambiguous short dates and grouped/decimal numbers are interpreted only when
 * the caller supplies a matching locale. An explicit Text disposition is
 * literal (including a leading '='); Date accepts parsed dates or a nonnegative
 * numeric serial and can use an explicit MDY/DMY/YMD order. Skip marks the
 * source field for projection removal. Under General a numeral or a run of
 * letters and digits is read under Automatic Data Conversion (`conversions`,
 * the user's own options when omitted, as Excel's Text Import and Text to
 * Columns read them): a leading-zero id converts only while "Remove leading
 * zeros" is on, and so on.
 */
export function classifyDelimitedImportCell(
  raw: string,
  locale?: string,
  disposition: DelimitedImportColumnDisposition = "general",
  dateOrder: DelimitedImportDateOrder = "locale",
  numberOptions?: DelimitedImportNumberOptions,
  dateSystem?: SpreadsheetDateSystemOptions,
  conversions: DataConversionOptions = readDataConversionOptions(),
  today?: Date,
): DelimitedImportCellValue {
  if (disposition === "skip") {
    return { value: raw, kind: "text", skip: true };
  }
  if (disposition === "text") {
    // Text keeps the field exactly as the file has it, a CR or CRLF inside a
    // quoted field included (the dialog says so; user audit S-116).
    return {
      value: raw,
      kind: "text",
      numberFormat: "@",
      literalText: true,
    };
  }
  // Any other kind reads a line break inside the field as the cell's one
  // line break, LF, whichever the file wrote.
  raw = raw.replace(/\r\n?/g, "\n");
  // An equals sign alone is not a formula: Excel keeps it as the text "=".
  // Stored bare, the page read it as an empty formula and the saved file
  // held nothing there (xlsx-rounds WALK R288-X24: Text to Columns over
  // "a = b" at spaces left "=" cells that read back empty).
  if (raw.trim() === "=") {
    return { value: raw, kind: "text", literalText: true };
  }
  const calendarSerial =
    disposition === "date" && dateOrder !== "locale"
      ? orderedDateSerial(raw, dateOrder)
      : isoDateSerial(raw) ?? (locale ? localeDateSerial(raw, locale) : null);
  // The serial of the receiving workbook's date system; a date before 1904
  // is not a date in a 1904 workbook and is read like any other text.
  const systemSerial =
    calendarSerial === null
      ? null
      : excelSystemSerial(calendarSerial, dateSystem?.date1904);
  const dateSerial =
    systemSerial !== null && systemSerial >= excelMinDateSerial(dateSystem?.date1904)
      ? systemSerial
      : null;
  if (dateSerial !== null) {
    return {
      value: String(dateSerial),
      kind: "number",
      numberFormat: "yyyy-mm-dd",
    };
  }
  const parsedNumber = importNumber(raw, locale, numberOptions, conversions);
  const numericValue = parsedNumber === null ? null : String(parsedNumber.value);
  // A percentage the way a cell displays one (50%, -12.5%): the number it
  // stands for, shown as a percentage, exactly as typing it into the sheet
  // stores it (live audit B189: the sheet's own CSV came back as text).
  const percent =
    numericValue === null && disposition === "general"
      ? importPercentValue(raw, locale, numberOptions, conversions)
      : null;
  if (percent) return percent;
  const conversion = parsedNumber?.conversion ? { conversion: parsedNumber.conversion } : {};
  if (disposition === "date") {
    if (numericValue !== null && Number(numericValue) >= 0) {
      return {
        value: numericValue,
        kind: "number",
        numberFormat: "yyyy-mm-dd",
        ...conversion,
      };
    }
    return { value: raw, kind: "text" };
  }
  if (numericValue !== null) {
    // 123E5 reads as typing it does: the number in scientific format.
    return parsedNumber?.exponent
      ? { value: numericValue, kind: "number", numberFormat: "0.00E+00", ...conversion }
      : { value: numericValue, kind: "number", ...conversion };
  }
  // JAN1, 1JAN, MARCH1: a date under "Convert continuous letters and numbers
  // to a date", read as typing it reads it.
  if (disposition === "general" && /^[A-Za-z0-9]+$/.test(raw.trim())) {
    const letterDate = parseLetterDigitDate(raw.trim(), today ?? new Date(), dateSystem?.date1904);
    if (letterDate && conversions.letterDates) {
      return {
        value: String(letterDate.serial),
        kind: "number",
        numberFormat: letterDate.format as "d-mmm" | "mmm-yy" | "d-mmm-yy",
        conversion: "letterDates",
      };
    }
  }
  return { value: raw, kind: "text" };
}

/**
 * The Automatic Data Conversion options opening `text` would use, in Excel's
 * list order: what Excel's notice names before it opens a .csv or .txt file
 * ("Remove leading zeros", ...). Empty when the file opens unchanged by
 * them. Reads at most `maxFields` populated fields (a large file's notice is
 * decided by its first part).
 */
export function delimitedImportConversions(
  text: string,
  delimiter: string,
  conversions: DataConversionOptions,
  maxFields = MAX_CELLS,
): DataConversionKind[] {
  const found = new Set<DataConversionKind>();
  let seen = 0;
  const done = new Error("done");
  try {
    scanDelimited(text, delimiter, {
      onField: (_row, _column, raw) => {
        if (raw === "") return;
        seen += 1;
        const conversion = classifyDelimitedImportCell(
          raw,
          undefined,
          "general",
          "locale",
          undefined,
          undefined,
          conversions,
        ).conversion;
        if (conversion) found.add(conversion);
        if (seen >= maxFields || found.size === 4) throw done;
      },
    });
  } catch (error) {
    // A file past the browser's budgets is decided by what was read.
    if (error !== done && !(error instanceof DelimitedImportLimitError)) throw error;
  }
  return (["leadingZeros", "longNumbers", "exponent", "letterDates"] as const).filter(
    (kind) => found.has(kind),
  );
}

/** Convert delimited text into a single-sheet spreadsheet page content, each
 *  field read under `conversions` (Excel's defaults when omitted). */
export function delimitedTextToSpreadsheetContent(
  text: string,
  delimiter = ",",
  projectionBudget = new PageImportProjectionBudget(),
  conversions: DataConversionOptions = DEFAULT_DATA_CONVERSION_OPTIONS,
): NativeSpreadsheetContent {
  const cells: NativeSpreadsheetContent["sheets"][number]["cells"] = {};
  let lastRow = -1;
  let lastCol = -1;
  try {
    // Root content, sheet, and cells-map records plus their string values.
    projectionBudget.reserveObjects(3);
    projectionBudget.reserveString(SPREADSHEET_CONTENT_TYPE);
    projectionBudget.reserveString("sheet-1");
    projectionBudget.reserveString("Sheet1");
    scanDelimited(text, delimiter, {
      onField: (r, c, raw) => {
        // CSV shape is defined by every parsed field, not only populated
        // cells. Keep sparse storage for empty values, but retain trailing and
        // all-empty columns/rows so `a,,` remains a three-column record and a
        // later export can reproduce the source structure.
        if (r > lastRow) lastRow = r;
        if (c > lastCol) lastCol = c;
        if (raw === "") return;
        // Classify: an unambiguous ISO date becomes a serial + date format (so it
        // is numeric/sortable and usable in formulas, like Excel/Sheets); else a
        // numeric-looking cell becomes a number; everything else stays text.
        const classified = classifyDelimitedImportCell(
          raw,
          undefined,
          "general",
          "locale",
          undefined,
          undefined,
          conversions,
        );
        let cell: NativeSpreadsheetContent["sheets"][number]["cells"][string];
        if (classified.kind === "number") {
          cell = {
            value: Number(classified.value),
            kind: "number",
            ...(classified.numberFormat
              ? { style: { numberFormat: classified.numberFormat } }
              : {}),
          };
        } else {
          cell = { value: classified.value, kind: "text" };
        }
        projectionBudget.consumeValue(cell);
        cells[`${r}:${c}`] = cell;
      },
    });
  } catch (error) {
    if (error instanceof PageImportProjectionBudgetError) {
      throw new DelimitedImportLimitError(
        error.reason === "object_nodes"
          ? "csv_node_limit_exceeded"
          : "csv_text_limit_exceeded",
        error.message,
      );
    }
    throw error;
  }

  return {
    type: SPREADSHEET_CONTENT_TYPE,
    sheets: [
      {
        id: "sheet-1",
        name: "Sheet1",
        // Dimensions cover exactly the populated extent and stay schema-valid.
        rowCount: Math.max(lastRow + 1, 1),
        columnCount: Math.max(lastCol + 1, 1),
        cells,
      },
    ],
  };
}
