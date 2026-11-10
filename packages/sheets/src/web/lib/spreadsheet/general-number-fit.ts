/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a number is drawn in its column, as Excel draws it.
 *
 * Text runs over empty neighbours; a number never does. Excel's General
 * format shows a number in at most 11 characters and then fits it to the
 * column: it rounds the decimals, or switches to scientific notation, and
 * shows # across the cell only when nothing fits. A number with any other
 * format is never rounded to fit: it shows ##### when the column is too
 * narrow (Microsoft Support, "Available number formats in Excel": General
 * "rounds the numbers with decimals" and "uses scientific (exponential)
 * notation for large numbers (12 or more digits)"; "Display numbers in
 * scientific (exponential) notation"; "How to correct a ##### error").
 *
 * Measured in Excel at the default column width (8.43 characters):
 *   123456789012345 -> 1.23E+14      3.14159265358979 -> 3.141593
 *   0.000012345 -> 1.23E-05          1234567890 -> 1.23E+09
 *   =1/3 -> 0.333333                 =2/3*1000000 -> 666666.7
 *   -123456789012 -> -1.2E+11        #,##0.00 of 123456789 -> ########
 *
 * Between a rounded decimal and scientific notation, the one that keeps more
 * significant digits wins, the decimal on a tie: 0.333333 beats 3.33E-01,
 * 1.23E-05 beats 0.000012. This is display only; the value never changes.
 */

import { DATE_TIME_OUT_OF_RANGE_DISPLAY } from "@mrnote/core/web/lib/spreadsheet/number-format.ts";

/** General's limit, not counting a minus sign. */
const GENERAL_MAX_CHARS = 11;

export interface NumberFitInput {
  value: number;
  /** The text the cell's format gives the value. */
  display: string;
  /** The cell's format is General. */
  general: boolean;
  /** The width the text may take, px. */
  availablePx: number;
  /** The drawn width of a text in the cell's font, px. */
  measure: (text: string) => number;
}

function trimFraction(text: string): string {
  return text.includes(".") ? text.replace(/0+$/, "").replace(/\.$/, "") : text;
}

/**
 * Excel holds a number to 15 significant digits and rounds what it shows on
 * those decimal digits, half away from zero: 0.00015 to four places is
 * 0.0002 (binary rounding, as toFixed does, gives 0.0001). `magnitude` as its
 * 15 digits and the power of ten of the first.
 */
function decimalDigits(magnitude: number): { digits: string; power: number } {
  const [mantissa, exponent] = magnitude.toExponential(14).split("e");
  return { digits: mantissa.replace(".", ""), power: Number(exponent) };
}

/** The first `count` of `digits` rounded half up, and whether that carried
 * into a new leading digit. */
function roundDigits(digits: string, count: number): { kept: bigint; carried: boolean } {
  if (count <= 0) return { kept: count === 0 && digits[0] >= "5" ? 1n : 0n, carried: count === 0 && digits[0] >= "5" };
  const head = BigInt(digits.padEnd(count, "0").slice(0, count));
  const kept = count < digits.length && digits[count] >= "5" ? head + 1n : head;
  return { kept, carried: String(kept).length > count };
}

/** toFixed on the 15-digit decimal value. */
function fixed(magnitude: number, places: number): string {
  const { digits, power } = decimalDigits(magnitude);
  const { kept } = roundDigits(digits, power + 1 + places);
  const text = String(kept).padStart(places + 1, "0");
  return places > 0 ? `${text.slice(0, -places)}.${text.slice(-places)}` : text;
}

/** 1.23457e+14 -> 1.23457E+14, 1.20000e-5 -> 1.2E-05. */
function scientific(magnitude: number, decimals: number): string {
  const { digits, power: first } = decimalDigits(magnitude);
  const { kept, carried } = roundDigits(digits, decimals + 1);
  const power = carried ? first + 1 : first;
  const run = String(kept).slice(0, decimals + 1);
  const mantissa = decimals > 0 ? `${run[0]}.${run.slice(1)}` : run;
  return `${trimFraction(mantissa)}E${power < 0 ? "-" : "+"}${String(Math.abs(power)).padStart(2, "0")}`;
}

function significantDigits(text: string): number {
  return text
    .replace(/^-/, "")
    .replace(/E.*$/, "")
    .replace(".", "")
    .replace(/^0+/, "").length;
}

function hashFill(availablePx: number, measure: (text: string) => number): string {
  const hash = measure("#");
  return "#".repeat(Math.max(1, hash > 0 ? Math.floor(availablePx / hash) : 1));
}

/** The text a number is drawn with in a column `availablePx` wide. */
export function fitNumberText({
  value,
  display,
  general,
  availablePx,
  measure,
}: NumberFitInput): string {
  const fits = (text: string) =>
    text.replace(/^-/, "").length <= GENERAL_MAX_CHARS && measure(text) <= availablePx;
  // What General writes is a decimal or a scientific spelling; any other
  // display (40.00 kg from a conditional format's number format) is a
  // format's, and fits as one.
  const generalSpelling = general && /^-?(\d+\.?\d*|\.\d+)(E[+-]\d+)?$/i.test(display);
  // A date or time the format cannot write is # across the whole cell, as a
  // number too wide for its column is (ED023: 1E12 under yyyy-mm-dd fills
  // its cell in Excel; the format's fixed eight marks sat in its right half).
  if (display === DATE_TIME_OUT_OF_RANGE_DISPLAY) return hashFill(availablePx, measure);
  if (!generalSpelling || !Number.isFinite(value)) {
    return measure(display) <= availablePx ? display : hashFill(availablePx, measure);
  }
  // A plain decimal the cell already shows in full is left as it is.
  if (/^-?(\d+\.?\d*|\.\d+)$/.test(display) && fits(display)) return display;
  if (value === 0) return fits("0") ? "0" : hashFill(availablePx, measure);
  const sign = value < 0 ? "-" : "";
  const magnitude = Math.abs(value);
  let decimal: string | null = null;
  const integerDigits = Math.floor(magnitude) === 0 ? 1 : String(Math.floor(magnitude)).length;
  if (magnitude < 1e21 && integerDigits <= GENERAL_MAX_CHARS) {
    for (let places = Math.max(0, GENERAL_MAX_CHARS - integerDigits - 1); places >= 0; places -= 1) {
      const text = sign + trimFraction(fixed(magnitude, places));
      if (!fits(text)) continue;
      if (Number(text) !== 0) decimal = text;
      break;
    }
  }
  let exponential: string | null = null;
  for (let places = 5; places >= 0; places -= 1) {
    const text = sign + scientific(magnitude, places);
    if (fits(text)) {
      exponential = text;
      break;
    }
  }
  if (decimal !== null && (exponential === null || significantDigits(decimal) >= significantDigits(exponential))) {
    return decimal;
  }
  return exponential ?? hashFill(availablePx, measure);
}
