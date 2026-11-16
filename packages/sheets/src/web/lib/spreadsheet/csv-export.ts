/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { csvFormulaNeedsEscape, toDelimitedText } from "@mrnote/core/web/lib/spreadsheet/csv.ts";

/** Match the API's complete spreadsheet text-export limits. */
export const SPREADSHEET_CSV_EXPORT_LIMITS = {
  maxCells: 1_000_000,
  maxBytes: 32 * 1024 * 1024,
} as const;

export class SpreadsheetCsvExportBudgetError extends Error {
  constructor() {
    super("spreadsheet_text_export_too_large");
    this.name = "SpreadsheetCsvExportBudgetError";
  }
}

/** Count the existing serializer's UTF-8 output without allocating its escaped field. */
function csvFieldBytes(value: string, remaining: number, delimiterCode = 44): number {
  // UTF-8 needs at least one byte per UTF-16 code unit, including lone surrogates.
  if (value.length > remaining) throw new SpreadsheetCsvExportBudgetError();
  let bytes = csvFormulaNeedsEscape(value) ? 1 : 0;
  let quoted = false;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 34) {
      bytes += 1;
      quoted = true;
    } // Doubled quote.
    if (code === delimiterCode || code === 10 || code === 13) quoted = true;
    if (code < 0x80) bytes += 1;
    else if (code < 0x800) bytes += 2;
    else if (
      code >= 0xd800 &&
      code <= 0xdbff &&
      value.charCodeAt(index + 1) >= 0xdc00 &&
      value.charCodeAt(index + 1) <= 0xdfff
    ) {
      bytes += 4;
      index += 1;
    } else bytes += 3; // Includes the replacement encoding of an unpaired surrogate.
    if (bytes + (quoted ? 2 : 0) > remaining)
      throw new SpreadsheetCsvExportBudgetError();
  }
  return bytes + (quoted ? 2 : 0);
}

/** Cells read (or rows written) between two yields of the chunked build. */
const CSV_EXPORT_CHUNK_CELLS = 20_000;

/**
 * Export one complete rectangle of display values. Refuse excessive dimensions
 * before reading a cell. Capture each value once, preflight final escaped UTF-8
 * bytes, then reuse the CSV serializer row by row; never build a dense 2D grid.
 * The BOM, commas and newlines all count toward the download budget.
 */
export function buildSpreadsheetCsvExport(
  rowCount: number,
  columnCount: number,
  readDisplayValue: (row: number, column: number) => string,
  limits: Readonly<{
    maxCells: number;
    maxBytes: number;
  }> = SPREADSHEET_CSV_EXPORT_LIMITS,
): string {
  const steps = csvExportSteps(rowCount, columnCount, readDisplayValue, limits);
  for (;;) {
    const step = steps.next();
    if (step.done) return step.value;
  }
}

/**
 * {@link buildSpreadsheetCsvExport} in chunks, handing the main thread back
 * between them (`yieldToMain`, a macrotask by default): a sheet of up to a
 * million cells no longer freezes the window while its CSV is written. The
 * result, budgets and refusals are the same.
 */
export async function buildSpreadsheetCsvExportInChunks(
  rowCount: number,
  columnCount: number,
  readDisplayValue: (row: number, column: number) => string,
  options: Readonly<{
    yieldToMain?: () => Promise<void>;
    limits?: Readonly<{ maxCells: number; maxBytes: number }>;
    /** "\t" writes Excel's Unicode Text (tab between fields, CRLF between
     * rows; {@link unicodeTextBytes} encodes it); a comma (the default), CSV. */
    delimiter?: "," | "\t";
  }> = {},
): Promise<string> {
  const yieldToMain = options.yieldToMain ?? (() => new Promise<void>((resolve) => setTimeout(resolve, 0)));
  const steps = csvExportSteps(
    rowCount,
    columnCount,
    readDisplayValue,
    options.limits ?? SPREADSHEET_CSV_EXPORT_LIMITS,
    options.delimiter ?? ",",
  );
  for (;;) {
    const step = steps.next();
    if (step.done) return step.value;
    await yieldToMain();
  }
}

function* csvExportSteps(
  rowCount: number,
  columnCount: number,
  readDisplayValue: (row: number, column: number) => string,
  limits: Readonly<{ maxCells: number; maxBytes: number }>,
  delimiter: "," | "\t" = ",",
): Generator<void, string, void> {
  if (
    ![rowCount, columnCount, limits.maxCells, limits.maxBytes].every(
      (value) => Number.isSafeInteger(value) && value > 0,
    ) ||
    rowCount > Math.floor(limits.maxCells / columnCount)
  ) {
    throw new SpreadsheetCsvExportBudgetError();
  }
  // N cells require N - 1 separators across their rectangular CSV, plus a 3-byte BOM.
  let remaining = limits.maxBytes - 3 - (rowCount * columnCount - 1);
  if (remaining < 0) throw new SpreadsheetCsvExportBudgetError();
  const rowsPerChunk = Math.max(1, Math.floor(CSV_EXPORT_CHUNK_CELLS / columnCount));
  const values: string[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    for (let column = 0; column < columnCount; column += 1) {
      const value = readDisplayValue(row, column);
      remaining -= csvFieldBytes(value, remaining, delimiter.charCodeAt(0));
      values.push(value);
    }
    if ((row + 1) % rowsPerChunk === 0 && row + 1 < rowCount) yield;
  }
  const lines: string[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const offset = row * columnCount;
    lines.push(
      toDelimitedText([values.slice(offset, offset + columnCount)], delimiter, true),
    );
    if ((row + 1) % rowsPerChunk === 0 && row + 1 < rowCount) yield;
  }
  // Excel's Unicode Text ends its rows with CRLF; the CSV keeps LF.
  return `\uFEFF${lines.join(delimiter === "\t" ? "\r\n" : "\n")}`;
}

/**
 * Excel's Save As Unicode Text (.txt): the text (its BOM included) as
 * UTF-16 little-endian, the encoding Excel writes it in and reads it back
 * from without asking (a UTF-8 .txt opens in the system code page).
 */
export function unicodeTextBytes(text: string): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(new ArrayBuffer(text.length * 2));
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    bytes[index * 2] = code & 0xff;
    bytes[index * 2 + 1] = code >> 8;
  }
  return bytes;
}

/** A stored cell as far as the CSV extent is concerned. */
interface CsvExtentCell {
  value?: unknown;
  formula?: unknown;
  runs?: readonly unknown[];
}

/**
 * The rectangle a CSV export covers: A1 through the last row and column that
 * hold a stored value, formula or rich text, or a cell of an effective
 * dynamic-array spill (`spilledKeys`, "row:col"). A spill lives in a derived,
 * read-only overlay and never in `cells`, so without it `=SEQUENCE(3,2)` in A1
 * would export only its anchor. An empty sheet is one blank cell. The result
 * still goes through {@link buildSpreadsheetCsvExport}'s cell and byte budgets.
 */
export function spreadsheetCsvExportExtent(
  cells: Readonly<Record<string, CsvExtentCell | undefined>>,
  spilledKeys: Iterable<string> = [],
): { rowCount: number; columnCount: number; empty: boolean } {
  let maxRow = -1;
  let maxCol = -1;
  const include = (key: string) => {
    const separator = key.indexOf(":");
    if (separator <= 0) return;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) return;
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  };
  for (const [key, cell] of Object.entries(cells)) {
    if (
      cell?.value === undefined &&
      cell?.formula === undefined &&
      !cell?.runs?.length
    ) {
      continue;
    }
    include(key);
  }
  for (const key of spilledKeys) include(key);
  const empty = maxRow < 0 || maxCol < 0;
  return {
    rowCount: empty ? 1 : maxRow + 1,
    columnCount: empty ? 1 : maxCol + 1,
    empty,
  };
}
