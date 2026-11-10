/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text transforms for a cell selection — the "smart toolbox" family WPS puts a
 * button on and Excel makes you write a formula for.
 *
 * Every transform here is a pure `string -> string`, so the caller can apply it
 * across a selection, diff the result and undo it as one edit. A transform
 * returns its input unchanged when it has nothing to do, which lets the caller
 * skip writing a cell (and skip an undo entry) rather than rewriting the sheet
 * with identical values.
 *
 * Formulas are the caller's problem, not this module's: a cell holding `=A1+1`
 * must keep its formula, so the caller passes the DISPLAYED text of a formula
 * cell only when the user asked to convert results to text.
 */

export type CellTransformId =
  | "trimSpaces"
  | "collapseSpaces"
  | "upperCase"
  | "lowerCase"
  | "properCase"
  | "halfWidth"
  | "fullWidth"
  | "chineseAmount"
  | "digitsOnly"
  | "stripDigits";

/**
 * Trim leading and trailing whitespace, including the ideographic space U+3000
 * that a Chinese IME produces and `String.prototype.trim` already covers, plus
 * the non-breaking space that a paste from a web page leaves behind.
 */
export function trimSpaces(value: string): string {
  return value.replace(/^[\s 　]+|[\s 　]+$/g, "");
}

/**
 * Excel's TRIM: trim the ends and collapse every inner run of spaces to one.
 * This is the transform people actually mean by "删除空格" on pasted data.
 */
export function collapseSpaces(value: string): string {
  return trimSpaces(value).replace(/[\s 　]+/g, " ");
}

export function upperCase(value: string): string {
  return value.toUpperCase();
}

export function lowerCase(value: string): string {
  return value.toLowerCase();
}

/**
 * Excel's PROPER: capitalise the first letter of every word and lower the rest.
 * Excel treats any non-letter as a word boundary, so "o'neil-smith" becomes
 * "O'Neil-Smith"; this matches that rather than a naive split on spaces.
 */
export function properCase(value: string): string {
  let previousWasLetter = false;
  let out = "";
  for (const char of value) {
    const isLetter = /\p{L}/u.test(char);
    out += isLetter && !previousWasLetter ? char.toUpperCase() : char.toLowerCase();
    previousWasLetter = isLetter;
  }
  return out;
}

// Full-width forms occupy U+FF01..U+FF5E and map to ASCII U+0021..U+007E by a
// fixed offset; the ideographic space is the one exception, pairing with U+0020.
const FULL_WIDTH_START = 0xff01;
const FULL_WIDTH_END = 0xff5e;
const ASCII_START = 0x21;
const WIDTH_OFFSET = FULL_WIDTH_START - ASCII_START;
const IDEOGRAPHIC_SPACE = "　";

/** Full-width characters to their ASCII equivalents (全角转半角). */
export function halfWidth(value: string): string {
  let out = "";
  for (const char of value) {
    const code = char.codePointAt(0) ?? 0;
    if (code >= FULL_WIDTH_START && code <= FULL_WIDTH_END) {
      out += String.fromCodePoint(code - WIDTH_OFFSET);
    } else if (char === IDEOGRAPHIC_SPACE) {
      out += " ";
    } else {
      out += char;
    }
  }
  return out;
}

/** ASCII characters to their full-width equivalents (半角转全角). */
export function fullWidth(value: string): string {
  let out = "";
  for (const char of value) {
    const code = char.codePointAt(0) ?? 0;
    if (code >= ASCII_START && code <= ASCII_START + (FULL_WIDTH_END - FULL_WIDTH_START)) {
      out += String.fromCodePoint(code + WIDTH_OFFSET);
    } else if (char === " ") {
      out += IDEOGRAPHIC_SPACE;
    } else {
      out += char;
    }
  }
  return out;
}

/** Keep only the digits, in order ("A1-B22" -> "122"). */
export function digitsOnly(value: string): string {
  return value.replace(/\D+/g, "");
}

/** Drop every digit, keeping the rest ("A1-B22" -> "A-B"). */
export function stripDigits(value: string): string {
  return value.replace(/\d+/g, "");
}

const CHINESE_DIGITS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
// Positions inside one four-digit group, then the group markers themselves.
const WITHIN_GROUP = ["", "拾", "佰", "仟"];
const GROUP_MARKS = ["", "万", "亿", "万亿"];
// The largest amount the four group markers above can name: four groups of
// four digits of yuan (9999万亿), in cents.
const MAX_CHINESE_AMOUNT_CENTS = 9999_9999_9999_9999_99n;

function chineseGroup(group: number): string {
  let out = "";
  let sawZero = false;
  const digits = String(group).padStart(4, "0").split("").map(Number);
  digits.forEach((digit, index) => {
    const position = 3 - index;
    if (digit === 0) {
      sawZero = true;
      return;
    }
    // One 零 stands for any run of zeros, and never trails the group.
    if (sawZero && out) out += CHINESE_DIGITS[0];
    sawZero = false;
    out += CHINESE_DIGITS[digit] + WITHIN_GROUP[position];
  });
  return out;
}

/**
 * The legal-form Chinese amount used on invoices and cheques (人民币大写), the
 * WPS button Excel has no equivalent for.
 *
 * Rounds to cents like the paper form does, names the units 元/角/分, and closes
 * with 整 when there is no fractional part. Returns an empty string for anything
 * that is not a finite number the unit markers can name, so the caller leaves
 * the cell alone rather than writing a wrong amount onto an invoice.
 */
export function chineseAmount(value: string): string {
  const text = trimSpaces(value).replace(/[,，]/g, "");
  if (!text || !/^-?\d*(\.\d*)?$/.test(text) || !/\d/.test(text)) return "";
  const negative = text.startsWith("-");
  // Round to cents in DECIMAL, straight off the digits. Going through a double
  
  // and an invoice that reads 壹元整 for 1.005 is a wrong invoice.
  const [whole, fraction = ""] = text.replace(/^-/, "").split(".");
  const digits = `${whole || "0"}${fraction.padEnd(3, "0").slice(0, 3)}`;
  if (!/^\d+$/.test(digits)) return "";
  const thousandths = BigInt(digits);
  const cents = thousandths / 10n + (thousandths % 10n >= 5n ? 1n : 0n);
  if (cents > MAX_CHINESE_AMOUNT_CENTS) return "";
  if (cents === 0n) return "零元整";

  const yuan = cents / 100n;
  const jiao = (cents / 10n) % 10n;
  const fen = cents % 10n;

  let integerText = "";
  if (yuan === 0n) {
    integerText = "";
  } else {
    const groups: number[] = [];
    let remaining = yuan;
    while (remaining > 0n) {
      groups.push(Number(remaining % 10000n));
      remaining /= 10000n;
    }
    if (groups.length > GROUP_MARKS.length) return "";
    for (let index = groups.length - 1; index >= 0; index -= 1) {
      const group = groups[index];
      const named = chineseGroup(group);
      if (named) {
        // A group under 1000 that follows a bigger one needs its own 零, or
        // 100,000,001 would read as 壹亿壹 instead of 壹亿零壹.
        if (integerText && group < 1000) integerText += CHINESE_DIGITS[0];
        integerText += named + GROUP_MARKS[index];
      }
    }
    integerText += "元";
  }

  let fractionText = "";
  if (jiao === 0n && fen === 0n) {
    fractionText = integerText ? "整" : "";
  } else {
    if (jiao > 0n) fractionText += CHINESE_DIGITS[Number(jiao)] + "角";
    // A missing 角 between 元 and 分 reads as 零, exactly as the paper form does.
    else if (integerText && fen > 0n) fractionText += CHINESE_DIGITS[0];
    if (fen > 0n) fractionText += CHINESE_DIGITS[Number(fen)] + "分";
  }

  const body = `${integerText}${fractionText}` || "零元整";
  return negative ? `负${body}` : body;
}

const TRANSFORMS: Record<CellTransformId, (value: string) => string> = {
  trimSpaces,
  collapseSpaces,
  upperCase,
  lowerCase,
  properCase,
  halfWidth,
  fullWidth,
  chineseAmount,
  digitsOnly,
  stripDigits,
};

export interface CellTransformResult {
  /** The transformed text, or null when the cell must be left untouched. */
  value: string | null;
}

export interface CellTransformSourceEntry {
  /** Stable worksheet cell key (`row:column`). */
  key: string;
  value: unknown;
}

export interface CellTransformChange {
  key: string;
  before: string;
  after: string;
}

/**
 * Apply one transform to a cell's stored value.
 *
 * A formula cell is never rewritten: replacing `=SUM(A1:A9)` with its uppercase
 * text would silently destroy the formula, which is not what "make this
 * uppercase" asks for. A transform that produces no change, or that declines the
 * input (an amount that cannot be named), also returns null so the caller writes
 * nothing and the edit stays out of the undo stack.
 */
export function applyCellTransform(
  storedValue: unknown,
  transform: CellTransformId,
): CellTransformResult {
  if (typeof storedValue !== "string" || storedValue === "") {
    return { value: null };
  }
  if (storedValue.startsWith("=")) return { value: null };
  const next = TRANSFORMS[transform](storedValue);
  if (next === storedValue || next === "") return { value: null };
  return { value: next };
}

/**
 * Build the complete, deterministic write plan for a selection without
 * mutating it. The UI uses this as its preview/confirmation boundary, then
 * rechecks the live worksheet before applying every change as one history
 * entry. Formula, blank and no-op cells are omitted by `applyCellTransform`.
 */
export function planCellTransform(
  entries: readonly CellTransformSourceEntry[],
  transform: CellTransformId,
): CellTransformChange[] {
  const changes: CellTransformChange[] = [];
  for (const entry of entries) {
    const result = applyCellTransform(entry.value, transform);
    if (result.value === null || typeof entry.value !== "string") continue;
    changes.push({
      key: entry.key,
      before: entry.value,
      after: result.value,
    });
  }
  return changes;
}
