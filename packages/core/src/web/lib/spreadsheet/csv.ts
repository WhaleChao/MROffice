/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TextToColumnsQualifier } from "./text-to-columns";

/**
 * RFC-4180-ish CSV parser for spreadsheet import.
 *
 * Handles quoted fields, embedded commas / newlines / quotes ("" escape), and
 * both LF and CRLF line endings. Delimiter is configurable (defaults to ",",
 * pass "\t" for TSV). Returns a 2D array of cell strings.
 */
export function parseCsv(input: string, delimiter = ","): string[][] {
  // A leading UTF-8 BOM (Excel "Save As CSV UTF-8" / many exporters) must not
  // bleed into the first cell — strip it before parsing.
  const text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  // A quote opens a quoted field only as the field's first character (RFC
  // 4180, and how Excel and Sheets read a paste). A quote met after the field
  // has content is that content: =EXACT("a","A") pasted as plain text keeps
  // its string arguments, and plain "quoted" text keeps its quotes.
  let atFieldStart = true;
  let i = 0;
  const n = text.length;

  const pushField = () => {
    row.push(field);
    field = "";
    atFieldStart = true;
  };
  const pushRow = () => {
    pushField();
    rows.push(row);
    row = [];
  };

  while (i < n) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      field += c;
      i += 1;
      continue;
    }
    if (c === '"' && atFieldStart) {
      inQuotes = true;
      atFieldStart = false;
      i += 1;
      continue;
    }
    if (c === delimiter) {
      pushField();
      i += 1;
      continue;
    }
    if (c === "\r") {
      // swallow CR (CRLF or lone CR) as a row break
      pushRow();
      if (text[i + 1] === "\n") i += 1;
      i += 1;
      continue;
    }
    if (c === "\n") {
      pushRow();
      i += 1;
      continue;
    }
    field += c;
    atFieldStart = false;
    i += 1;
  }
  // flush trailing field/row unless the input ended exactly on a row break;
  // a field that was opened (a quoted "" at the end) is a record even when it
  // is empty, so NEW\n"" is two records and the second clears its cell.
  if (field !== "" || row.length > 0 || !atFieldStart) {
    pushRow();
  }
  return rows;
}

/** Known delimiters exposed directly by the Spreadsheet text-import UI. */
export type CsvImportKnownDelimiter = "," | ";" | "\t" | "|" | " ";

/**
 * The delimiter a file's declared format fixes, or null when the format only
 * hints at one. TSV is tab-separated by definition (IANA
 * text/tab-separated-values: "fields are separated by tabs"), so a `.tsv` name
 * or that media type decides the delimiter even when the fields contain more
 * semicolons or commas than tabs. A `.csv` is not declared this way: regional
 * Excel writes semicolon-delimited `.csv` files, so it stays auto-detected.
 */
export function declaredCsvImportDelimiter(
  fileName: string,
  mediaType: string | undefined,
): CsvImportKnownDelimiter | null {
  const type = (mediaType || "").split(";", 1)[0].trim().toLowerCase();
  if (
    fileName.toLowerCase().endsWith(".tsv") ||
    type === "text/tab-separated-values"
  ) {
    return "\t";
  }
  return null;
}

declare const csvImportCustomDelimiterBrand: unique symbol;
export type CsvImportCustomDelimiter = string & {
  readonly [csvImportCustomDelimiterBrand]: true;
};

/** A known delimiter or a runtime-validated explicit one-character delimiter. */
export type CsvImportDelimiter =
  | CsvImportKnownDelimiter
  | CsvImportCustomDelimiter;

/** Reuse the same qualifier authority as Data > Text to Columns. */
export type CsvImportTextQualifier = TextToColumnsQualifier;

export function csvImportTextQualifierCharacter(
  qualifier: CsvImportTextQualifier,
): '"' | "'" | null {
  if (qualifier === "single") return "'";
  if (qualifier === "none") return null;
  return '"';
}

export function normalizeCsvImportDelimiter(
  value: string,
): CsvImportDelimiter | null {
  if (value.length !== 1 || value === '"' || /[\r\n]/.test(value)) return null;
  return value as CsvImportDelimiter;
}

export interface CsvImportDialect {
  /** Input with a UTF BOM and an optional Excel `sep=` control row removed. */
  readonly text: string;
  readonly delimiter: CsvImportDelimiter;
}

/** Encodings the Spreadsheet text-import boundary can decode deterministically. */
export type CsvImportEncoding =
  | "utf-8"
  | "utf-16le"
  | "utf-16be"
  | "gb18030";
export type CsvImportEncodingChoice = "auto" | CsvImportEncoding;

export interface DecodedCsvImport {
  readonly text: string;
  readonly encoding: CsvImportEncoding;
  readonly hadBom: boolean;
}

const startsWithBytes = (
  bytes: Uint8Array,
  signature: readonly number[],
): boolean =>
  signature.every((value, index) => bytes[index] === value);

const decodeBytes = (
  bytes: Uint8Array,
  encoding: CsvImportEncoding,
  offset = 0,
): string =>
  new TextDecoder(encoding, { fatal: true }).decode(bytes.subarray(offset));

/**
 * Decode an imported delimited-text file from its original bytes. BOMs are
 * authoritative for UTF-8/UTF-16. Without one, strict UTF-8 is attempted
 * first and GB18030 (the browser Encoding Standard decoder covering GBK input)
 * is the bounded legacy fallback. Unsupported or malformed bytes fail closed
 * instead of silently replacing data with U+FFFD. A non-auto choice applies
 * the user's preview-dialog override and strips only that encoding's own BOM.
 */
export function decodeCsvImportBytes(
  input: ArrayBuffer | Uint8Array,
  encodingChoice: CsvImportEncodingChoice = "auto",
): DecodedCsvImport {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  // UTF-32 begins with the UTF-16LE/BE prefix. Reject it explicitly rather than
  // mis-decoding alternating NUL bytes as spreadsheet content.
  if (
    startsWithBytes(bytes, [0xff, 0xfe, 0x00, 0x00]) ||
    startsWithBytes(bytes, [0x00, 0x00, 0xfe, 0xff])
  ) {
    throw new Error("UTF-32 text import is not supported.");
  }
  if (encodingChoice !== "auto") {
    const bomOffset =
      encodingChoice === "utf-8" &&
      startsWithBytes(bytes, [0xef, 0xbb, 0xbf])
        ? 3
        : encodingChoice === "utf-16le" &&
            startsWithBytes(bytes, [0xff, 0xfe])
          ? 2
          : encodingChoice === "utf-16be" &&
              startsWithBytes(bytes, [0xfe, 0xff])
            ? 2
            : 0;
    return {
      text: decodeBytes(bytes, encodingChoice, bomOffset),
      encoding: encodingChoice,
      hadBom: bomOffset > 0,
    };
  }
  if (startsWithBytes(bytes, [0xef, 0xbb, 0xbf])) {
    return {
      text: decodeBytes(bytes, "utf-8", 3),
      encoding: "utf-8",
      hadBom: true,
    };
  }
  if (startsWithBytes(bytes, [0xff, 0xfe])) {
    return {
      text: decodeBytes(bytes, "utf-16le", 2),
      encoding: "utf-16le",
      hadBom: true,
    };
  }
  if (startsWithBytes(bytes, [0xfe, 0xff])) {
    return {
      text: decodeBytes(bytes, "utf-16be", 2),
      encoding: "utf-16be",
      hadBom: true,
    };
  }
  try {
    return {
      text: decodeBytes(bytes, "utf-8"),
      encoding: "utf-8",
      hadBom: false,
    };
  } catch {
    try {
      return {
        text: decodeBytes(bytes, "gb18030"),
        encoding: "gb18030",
        hadBom: false,
      };
    } catch {
      throw new Error(
        "This text file is not valid UTF-8, UTF-16, GB18030, or GBK.",
      );
    }
  }
}

const CSV_IMPORT_AUTO_DELIMITERS: readonly CsvImportKnownDelimiter[] = [
  ",",
  ";",
  "\t",
  "|",
];
const CSV_DIALECT_SAMPLE_RECORDS = 20;
const CSV_DIALECT_SAMPLE_CHARS = 128 * 1024;

/**
 * Resolve the physical dialect of an imported CSV/Text file without treating
 * separators inside fields enclosed by the selected text qualifier as
 * evidence. A leading `sep=`
 * control row (including a validated custom character) is authoritative and
 * is removed;
 * otherwise up to 20 logical records are scored for a stable comma, semicolon,
 * tab, or pipe shape. An explicit consecutive-delimiter choice scores a run of
 * the same candidate as one boundary. Space is intentionally manual-only so
 * ordinary prose is not guessed as a table. A filename/MIME hint only breaks
 * an equal score.
 */
export function resolveCsvImportDialect(
  input: string,
  hint: CsvImportDelimiter = ",",
  textQualifier: CsvImportTextQualifier = "double",
  consecutiveDelimiters = false,
): CsvImportDialect {
  const text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input;
  const directive = /^sep=([^\r\n])(?:\r\n|\r|\n|$)/i.exec(text);
  const directiveDelimiter = directive
    ? normalizeCsvImportDelimiter(directive[1])
    : null;
  if (directive && directiveDelimiter) {
    return {
      text: text.slice(directive[0].length),
      delimiter: directiveDelimiter,
    };
  }

  const candidates = CSV_IMPORT_AUTO_DELIMITERS.map((delimiter) => ({
    delimiter,
    rows: [] as number[],
    current: 0,
  }));
  let inQuotes = false;
  let recordHasContent = false;
  let sampledRecords = 0;
  let commaCount = 0;
  let decimalCommaCount = 0;
  const sampleLength = Math.min(text.length, CSV_DIALECT_SAMPLE_CHARS);
  const qualifier = csvImportTextQualifierCharacter(textQualifier);

  const finishRecord = () => {
    if (recordHasContent || candidates.some((candidate) => candidate.current)) {
      for (const candidate of candidates) {
        candidate.rows.push(candidate.current);
        candidate.current = 0;
      }
      sampledRecords += 1;
    } else {
      for (const candidate of candidates) candidate.current = 0;
    }
    recordHasContent = false;
  };

  for (
    let index = 0;
    index < sampleLength && sampledRecords < CSV_DIALECT_SAMPLE_RECORDS;
    index += 1
  ) {
    const char = text[index];
    if (qualifier && char === qualifier) {
      recordHasContent = true;
      if (inQuotes && text[index + 1] === qualifier) {
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (inQuotes) {
      recordHasContent ||= !/\s/.test(char);
      continue;
    }
    const candidate = candidates.find((entry) => entry.delimiter === char);
    if (candidate) {
      candidate.current += 1;
      recordHasContent = true;
      if (char === ",") {
        commaCount += 1;
        if (/\d/.test(text[index - 1] ?? "") && /\d/.test(text[index + 1] ?? "")) {
          decimalCommaCount += 1;
        }
      }
      if (consecutiveDelimiters) {
        while (text[index + 1] === char) index += 1;
      }
      continue;
    }
    if (char === "\r" || char === "\n") {
      finishRecord();
      if (char === "\r" && text[index + 1] === "\n") index += 1;
      continue;
    }
    recordHasContent ||= !/\s/.test(char);
  }
  if (
    sampledRecords < CSV_DIALECT_SAMPLE_RECORDS &&
    (recordHasContent || candidates.some((candidate) => candidate.current))
  ) {
    finishRecord();
  }

  const semicolonSeen =
    candidates.find((candidate) => candidate.delimiter === ";")?.rows.some(
      (count) => count > 0,
    ) === true;
  const scores = candidates.map((candidate) => {
    const decimalOnlyComma =
      candidate.delimiter === "," &&
      semicolonSeen &&
      commaCount > 0 &&
      commaCount === decimalCommaCount;
    const positive = decimalOnlyComma
      ? []
      : candidate.rows.filter((count) => count > 0);
    const frequencies = new Map<number, number>();
    for (const count of positive) {
      frequencies.set(count, (frequencies.get(count) ?? 0) + 1);
    }
    return {
      delimiter: candidate.delimiter,
      consistentRows: Math.max(0, ...frequencies.values()),
      coveredRows: positive.length,
      occurrences: positive.reduce((sum, count) => sum + count, 0),
      hinted: candidate.delimiter === hint ? 1 : 0,
    };
  });
  scores.sort(
    (left, right) =>
      right.consistentRows - left.consistentRows ||
      right.coveredRows - left.coveredRows ||
      right.occurrences - left.occurrences ||
      right.hinted - left.hinted ||
      CSV_IMPORT_AUTO_DELIMITERS.indexOf(left.delimiter) -
        CSV_IMPORT_AUTO_DELIMITERS.indexOf(right.delimiter),
  );
  return {
    text,
    delimiter: scores[0]?.coveredRows ? scores[0].delimiter : hint,
  };
}

/** Parse an imported CSV/TSV/Text payload through the resolved physical dialect. */
export function parseCsvImport(
  input: string,
  hint: CsvImportDelimiter = ",",
): string[][] {
  const dialect = resolveCsvImportDialect(input, hint);
  return parseCsv(dialect.text, dialect.delimiter);
}

/**
 * Quote a single field for delimited output when it contains the delimiter, a
 * double-quote, or a line break — matching the spreadsheet-clipboard / RFC-4180
 * convention (wrap in double quotes, double any embedded quote). This is the
 * inverse of `parseCsv`, so values survive a copy → paste / export → import
 * round-trip even when a cell holds a tab, newline or quote.
 */
export function csvEscapeField(value: string, delimiter = ","): string {
  if (
    value.includes(delimiter) ||
    value.includes('"') ||
    value.includes("\n") ||
    value.includes("\r")
  ) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

/**
 * Neutralize CSV / formula injection (CWE-1236) for a field bound for ANOTHER
 * app's clipboard / a downloaded file: a value that Excel / Sheets / WPS would
 * treat as a FORMULA — it begins with `=`, `+`, `-`, `@` or a control char — is
 * prefixed with a leading apostrophe so it renders as TEXT instead of executing
 * (`=cmd|'/c calc'!A1`, `=HYPERLINK(...)`, a DDE payload, …). A genuine finite
 * NUMBER (`-5`, `+3.1`) is left untouched (safe, and prefixing would corrupt it).
 * Quoting alone does NOT help — Excel strips the quotes and still evaluates it.
 *
 * NOT applied to `csvEscapeField` (which is the round-trip-symmetric inverse of
 * `parseCsv`): use this ONLY at an export-to-foreign-app boundary, so an in-app
 * copy→paste still carries raw formulas.
 */
export function escapeCsvFormula(value: string): string {
  return csvFormulaNeedsEscape(value) ? `'${value}` : value;
}

/**
 * A number as a cell DISPLAYS it: an optional sign, currency symbol, digit
 * grouping, decimals, a percent sign or an exponent (`-50%`, `-1,234.5`,
 * `-$5.00`, `+1.5E3`). Such a field is data, not a formula, and must not get
 * the apostrophe that would turn it into text on the way back in (live audit
 * B188: a negative percentage came back as text and dropped out of COUNT).
 */
const DISPLAYED_NUMBER_RE =
  /^[+-]?(?:[$€£¥￥]\s?)?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:\s?%|[eE][+-]?\d+)?$/;

/** Shared predicate lets export budgeting count the prefix before allocating it. */
export function csvFormulaNeedsEscape(value: string): boolean {
  return Boolean(
    value &&
      /^[=+\-@\t\r\n]/.test(value) &&
      !Number.isFinite(Number(value)) &&
      !DISPLAYED_NUMBER_RE.test(value.trim()),
  );
}

/**
 * Serialize a 2D grid of cell strings to delimited text (CSV by default; pass
 * "\t" for the TSV the clipboard / Excel use). Round-trips through `parseCsv`.
 * When `guardFormulaInjection` is set (an export bound for a FOREIGN app — a
 * system-clipboard write or file download), each field is first run through
 * {@link escapeCsvFormula} so a malicious cell cannot execute as a formula when
 * the data is opened in Excel/Sheets; leave it off for an in-app round-trip.
 */
export function toDelimitedText(
  rows: string[][],
  delimiter = ",",
  guardFormulaInjection = false,
): string {
  return rows
    .map((row) =>
      row
        .map((f) =>
          csvEscapeField(guardFormulaInjection ? escapeCsvFormula(f) : f, delimiter),
        )
        .join(delimiter),
    )
    .join("\n");
}
