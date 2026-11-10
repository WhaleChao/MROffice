/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Automatic Data Conversion (Excel: File > Options > Data; Mac: Excel >
 * Preferences > Edit), read off Microsoft Support "Set automatic data
 * conversions" and the Microsoft 365 Insider post that introduced it:
 *
 * - "Remove leading zeros and convert to a number": 00123 -> 123. Off: the
 *   field stays the text 00123.
 * - "Keep first 15 digits of long numbers and display in scientific notation
 *   if required": a numeral of 16 or more digits becomes a number of 15
 *   significant digits (the rest zeros: 1234567890123459 -> ...450, measured
 *   in Excel). Off: it stays text with every digit.
 * - "Convert digits surrounding the letter "E" to a number in scientific
 *   notation": 123E5 -> 1.23E+07. Off: text.
 * - "Convert continuous letters and numbers to a date": JAN1 -> 1-Jan. Off:
 *   text (a date written with a space or hyphen, JAN 1, still converts).
 * - "When loading a .csv file or similar file, notify me of any automatic
 *   data conversions": opening a .csv/.txt whose fields one of the enabled
 *   conversions would change first asks, and the file can be opened once
 *   without converting.
 *
 * Every option is on in a fresh Excel (the conversions Excel always made).
 * They apply where Excel applies them: opening a .csv or .txt file, typing,
 * pasting, Find and Replace, and Text to Columns. Excel keeps them per
 * application, and the workbook has no field for them, so they are kept per
 * browser here (as Error Checking Options are).
 */

export interface DataConversionOptions {
  readonly leadingZeros: boolean;
  readonly longNumbers: boolean;
  readonly exponent: boolean;
  readonly letterDates: boolean;
  /** Ask before a .csv/.txt file is opened with a conversion. */
  readonly notifyCsv: boolean;
}

/** The four conversions an option can switch off, in Excel's list order. */
export const DATA_CONVERSION_KINDS = [
  "leadingZeros",
  "longNumbers",
  "exponent",
  "letterDates",
] as const;
export type DataConversionKind = (typeof DATA_CONVERSION_KINDS)[number];

export const DEFAULT_DATA_CONVERSION_OPTIONS: DataConversionOptions = Object.freeze({
  leadingZeros: true,
  longNumbers: true,
  exponent: true,
  letterDates: true,
  notifyCsv: true,
});

/** "Don't Convert": the file opened once with none of the four conversions. */
export function withoutDataConversions(options: DataConversionOptions): DataConversionOptions {
  return Object.freeze({
    ...options,
    leadingZeros: false,
    longNumbers: false,
    exponent: false,
    letterDates: false,
  });
}

/**
 * The four switches as the server's CSV reader takes them: one character per
 * conversion in Excel's order, "1" on and "0" off ("1111" is Excel's default).
 */
export function dataConversionFlags(options: DataConversionOptions): string {
  return DATA_CONVERSION_KINDS.map((kind) => (options[kind] ? "1" : "0")).join("");
}

const STORAGE_KEY = "mrai.spreadsheet.dataConversionOptions";

let storageWritable = true;
let memoryRaw: string | null = null;
let snapshot: { raw: string | null; value: DataConversionOptions } | null = null;
const listeners = new Set<() => void>();

export function parseDataConversionOptions(raw: string | null): DataConversionOptions {
  if (!raw) return DEFAULT_DATA_CONVERSION_OPTIONS;
  let stored: unknown;
  try {
    stored = JSON.parse(raw);
  } catch {
    return DEFAULT_DATA_CONVERSION_OPTIONS;
  }
  if (!stored || typeof stored !== "object" || Array.isArray(stored)) {
    return DEFAULT_DATA_CONVERSION_OPTIONS;
  }
  const record = stored as Record<string, unknown>;
  const read = (key: keyof DataConversionOptions) =>
    typeof record[key] === "boolean" ? (record[key] as boolean) : DEFAULT_DATA_CONVERSION_OPTIONS[key];
  return Object.freeze({
    leadingZeros: read("leadingZeros"),
    longNumbers: read("longNumbers"),
    exponent: read("exponent"),
    letterDates: read("letterDates"),
    notifyCsv: read("notifyCsv"),
  });
}

function readRaw(): string | null {
  if (typeof window === "undefined") return null;
  if (!storageWritable) return memoryRaw;
  try {
    return window.localStorage.getItem(STORAGE_KEY);
  } catch {
    return memoryRaw;
  }
}

/** Stable between changes, as useSyncExternalStore requires. */
export function readDataConversionOptions(): DataConversionOptions {
  const raw = readRaw();
  if (snapshot?.raw !== raw) {
    snapshot = { raw, value: parseDataConversionOptions(raw) };
  }
  return snapshot.value;
}

export const serverDataConversionOptions = (): DataConversionOptions =>
  DEFAULT_DATA_CONVERSION_OPTIONS;

export function writeDataConversionOptions(next: DataConversionOptions): void {
  const raw = JSON.stringify({
    leadingZeros: next.leadingZeros,
    longNumbers: next.longNumbers,
    exponent: next.exponent,
    letterDates: next.letterDates,
    notifyCsv: next.notifyCsv,
  });
  memoryRaw = raw;
  try {
    window.localStorage.setItem(STORAGE_KEY, raw);
    storageWritable = true;
  } catch {
    // Private windows and full storage still get the choice for this session.
    storageWritable = false;
  }
  for (const listener of listeners) listener();
}

/** Every open window follows the same choice, as in Excel, and so does
 * another tab of this browser. */
export function subscribeDataConversionOptions(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === STORAGE_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}
