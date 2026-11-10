/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data a formula reads from outside its workbook: Excel's WEBSERVICE (the
 * text a URL answers) and IMPORTRANGE (a range of another spreadsheet page in
 * MRAI the reader may read).
 *
 * Neither the calculation nor the page fetches anything itself. The browser
 * asks the page's server (`web-data-store`: SSRF-guarded and per-host consent
 * for WEBSERVICE, the reader's own page permissions for IMPORTRANGE), and
 * hands what it holds to the calculation session with each calculation
 * (`BackgroundCalculationRequest.externalData`); a call whose data the session
 * does not hold answers #BUSY! and is named back in the answer
 * (`externalRequests`), which is how the browser learns what to fetch.
 *
 * One key per call: its kind and its normalised arguments. This module is
 * pure (the evaluator, the calculation runtime, the engine host and the
 * browser all read it).
 */

import { nearestKnownCellMap } from "./cell-map-lineage";
import type { ArrayValue, CellValue, FormulaError } from "./formula-engine";

/** Excel: a GET request's URL is at most 2,048 characters. */
export const MAX_WEB_SERVICE_URL = 2_048;
/** Excel's cell text limit: a longer answer is #VALUE!. */
export const MAX_EXTERNAL_TEXT = 32_767;
/** The most cells one IMPORTRANGE brings (a larger range is #VALUE!). */
export const MAX_IMPORT_RANGE_CELLS = 10_000;
/** The longest range text IMPORTRANGE takes. */
export const MAX_IMPORT_RANGE_TEXT = 255;
/** The most calls one calculation holds data for (and names back). */
export const MAX_CALCULATION_EXTERNAL_DATA = 128;
/** The most characters all the data one calculation carries may hold. */
export const MAX_CALCULATION_EXTERNAL_CHARS = 4_000_000;
/** The longest key (a URL, or a page id with a range). */
export const MAX_EXTERNAL_DATA_KEY = 2_400;

const WEB_SERVICE = "webservice";
const IMPORT_RANGE = "importrange";

export type ExternalDataKind = typeof WEB_SERVICE | typeof IMPORT_RANGE;

/** One grid cell as it travels: a value, or an error code. */
export type ExternalCellWire = number | string | boolean | null | { readonly error: string };

/** What the browser holds for one call, as it travels to the session. */
export type ExternalDatumWire =
  | { readonly text: string }
  | { readonly error: string }
  | { readonly grid: ReadonlyArray<ReadonlyArray<ExternalCellWire>> };

/** One call's data: its key and what it reads as. */
export type CalculationExternalDatum = readonly [key: string, value: ExternalDatumWire];

/**
 * How the evaluator reads a call's data: the value (WEBSERVICE's text, an
 * IMPORTRANGE grid or cell) or its error; undefined while the host holds
 * none (#BUSY!).
 */
export type ExternalDataLookup = (key: string) => CellValue | ArrayValue | undefined;

/** The errors a call may read as (anything else from a host is #VALUE!). */
const EXTERNAL_ERRORS = new Set(["#VALUE!", "#REF!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#N/A", "#NAME?", "#NUM!", "#DIV/0!", "#NULL!", "#SPILL!", "#CALC!"]);

export function webServiceDataKey(url: string): string {
  return `${WEB_SERVICE}\u0000${url}`;
}

const STORED_WEB_SERVICE_CALL = /^=\s*(?:_xlfn\.)?WEBSERVICE\(\s*"((?:[^"]|"")*)"\s*\)\s*$/i;

/**
 * The results a workbook's own WEBSERVICE cells were saved with, for the
 * calls among `keys`: a cell whose whole formula is WEBSERVICE of a literal
 * URL and whose stamp (`display`, the file's cached value on import) is a
 * text. Excel for Mac does not fetch on open and shows that cached text; a
 * call this server refuses (#BLOCKED!) shows it too, never fetching (X-330).
 */
export function storedWebServiceAnswers(
  sheets: ReadonlyArray<{ readonly cells: Readonly<Record<string, { readonly value?: unknown; readonly formula?: unknown; readonly display?: unknown } | undefined>> }>,
  keys: readonly string[],
): Map<string, string> {
  const wanted = new Set(keys.filter((key) => parseExternalDataKey(key)?.kind === WEB_SERVICE));
  const found = new Map<string, string>();
  if (!wanted.size) return found;
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells)) {
      const formula = typeof cell?.formula === "string" ? cell.formula : cell?.value;
      if (typeof formula !== "string" || !/WEBSERVICE/i.test(formula)) continue;
      const match = STORED_WEB_SERVICE_CALL.exec(formula.trim());
      const display = cell?.display;
      if (!match || typeof display !== "string" || !display || display.length > MAX_EXTERNAL_TEXT || EXTERNAL_ERRORS.has(display)) continue;
      const key = webServiceDataKey(match[1].replace(/""/g, '"').trim());
      if (wanted.has(key) && !found.has(key)) found.set(key, display);
    }
  }
  return found;
}

export function importRangeDataKey(pageId: string, range: string): string {
  return `${IMPORT_RANGE}\u0000${pageId}\u0000${range}`;
}

/** A key read back: its kind and arguments; null for anything else. */
export function parseExternalDataKey(
  key: unknown,
): { kind: typeof WEB_SERVICE; url: string } | { kind: typeof IMPORT_RANGE; pageId: string; range: string } | null {
  if (typeof key !== "string" || key.length === 0 || key.length > MAX_EXTERNAL_DATA_KEY) return null;
  const parts = key.split("\u0000");
  if (parts[0] === WEB_SERVICE && parts.length === 2 && parts[1] && parts[1].length <= MAX_WEB_SERVICE_URL) {
    return { kind: WEB_SERVICE, url: parts[1] };
  }
  if (parts[0] === IMPORT_RANGE && parts.length === 3 && PAGE_ID.test(parts[1]) && parts[2] && parts[2].length <= MAX_IMPORT_RANGE_TEXT) {
    return { kind: IMPORT_RANGE, pageId: parts[1], range: parts[2] };
  }
  return null;
}

/**
 * WEBSERVICE's url as Excel takes it: http or https with a host, at most
 * 2,048 characters; null where Excel answers #VALUE!.
 */
export function webServiceUrl(text: string): string | null {
  const url = text.trim();
  if (!url || url.length > MAX_WEB_SERVICE_URL) return null;
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return null;
  }
  if ((parsed.protocol !== "https:" && parsed.protocol !== "http:") || !parsed.hostname) return null;
  return url;
}

const PAGE_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const PAGE_ID_IN_URL = /\/pages\/([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})(?:[/?#]|$)/;

/**
 * The page IMPORTRANGE's first argument names: a page's address
 * (`…/pages/<id>`) or its id. Null when it names none.
 */
export function importRangePageId(text: string): string | null {
  const value = text.trim();
  if (!value || value.length > MAX_WEB_SERVICE_URL) return null;
  const bare = value.toLowerCase();
  if (PAGE_ID.test(bare)) return bare;
  const match = PAGE_ID_IN_URL.exec(value);
  return match ? match[1].toLowerCase() : null;
}

// Only tested: `(?:^|[^…])` is the lookbehind Safari < 16.4 cannot compile.
const IMPORT_RANGE_CALL = /(?:^|[^A-Za-z0-9_.])IMPORTRANGE\s*\(/i;

/**
 * Whether a workbook reads another page's data (an IMPORTRANGE call in a
 * cell's formula or a defined name; a mention inside a text literal counts
 * too). Such a workbook saves no computed formula values (display-cache): a
 * reader of this page who may not read the source page must not find its
 * data in what this page stores, prints or exports.
 */
export function workbookImportsRanges(
  sheets: ReadonlyArray<{ readonly cells: Readonly<Record<string, { readonly value?: unknown } | undefined>>; readonly namedRanges?: unknown }>,
): boolean {
  for (const sheet of sheets) {
    if (sheet.namedRanges && IMPORT_RANGE_CALL.test(JSON.stringify(sheet.namedRanges))) return true;
    if (cellsImportRanges(sheet.cells)) return true;
  }
  return false;
}

const cellsImportOfMap = new WeakMap<object, boolean>();

/** By the map; from a map known to hold none, only the keys changed since are read
 * (a walk of every cell was about 8 ms per entry on 30,000). */
function cellsImportRanges(cells: Readonly<Record<string, { readonly value?: unknown } | undefined>>): boolean {
  const kept = cellsImportOfMap.get(cells);
  if (kept !== undefined) return kept;
  const reads = (key: string): boolean => {
    const value = cells[key]?.value;
    return typeof value === "string" && value.charCodeAt(0) === 61 /* = */ && IMPORT_RANGE_CALL.test(value);
  };
  const near = nearestKnownCellMap(cells as never, (map) => (cellsImportOfMap.get(map as object) === false ? true : undefined));
  let found = false;
  if (near) found = near.keys.some(reads);
  else for (const key in cells) if (reads(key)) { found = true; break; }
  cellsImportOfMap.set(cells, found);
  return found;
}

function errorValue(code: string): FormulaError {
  return { __formulaError: true, code: EXTERNAL_ERRORS.has(code) ? code : "#VALUE!" } as FormulaError;
}

function cellOf(wire: unknown): CellValue | undefined {
  if (wire === null) return null;
  if (typeof wire === "number") return Number.isFinite(wire) ? wire : undefined;
  if (typeof wire === "boolean") return wire;
  if (typeof wire === "string") return wire.length <= MAX_EXTERNAL_TEXT ? wire : undefined;
  if (typeof wire === "object" && typeof (wire as { error?: unknown }).error === "string") {
    return errorValue((wire as { error: string }).error);
  }
  return undefined;
}

/**
 * A datum as the evaluator reads it, checked: a text, an error or a
 * rectangular grid within the bounds. Undefined for anything malformed.
 * `chars` adds the characters it carries (the caller bounds the total).
 */
export function decodeExternalDatum(wire: unknown, chars: { count: number }): CellValue | ArrayValue | undefined {
  if (typeof wire !== "object" || wire === null) return undefined;
  const record = wire as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== 1) return undefined;
  if (typeof record.text === "string") {
    if (record.text.length > MAX_EXTERNAL_TEXT) return undefined;
    chars.count += record.text.length;
    return record.text;
  }
  if (typeof record.error === "string") return errorValue(record.error);
  if (!Array.isArray(record.grid) || record.grid.length === 0) return undefined;
  const width = Array.isArray(record.grid[0]) ? record.grid[0].length : 0;
  if (width === 0 || record.grid.length * width > MAX_IMPORT_RANGE_CELLS) return undefined;
  const grid: CellValue[][] = [];
  for (const line of record.grid) {
    if (!Array.isArray(line) || line.length !== width) return undefined;
    const row: CellValue[] = [];
    for (const cell of line) {
      const value = cellOf(cell);
      if (value === undefined) return undefined;
      if (typeof value === "string") chars.count += value.length;
      row.push(value);
    }
    grid.push(row);
  }
  // A single cell reads as that value.
  if (grid.length === 1 && width === 1) return grid[0][0];
  return { __array: true, grid };
}
