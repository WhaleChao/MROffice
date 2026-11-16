/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one spelling rule for a bare name in a formula — a sheet name written
 * without quotes, a defined name, a LET / LAMBDA name — shared by the formula
 * engine's tokenizer, the defined-name validator, the Evaluate Formula walk and
 * the defined-name rename scanner, so what one of them accepts the others read.
 *
 * Characters are compared by whole code point, so an astral CJK ideograph is a
 * single letter. A letter is a letter of any script (`\p{L}`): Excel stores
 * `=回款!C4` and `=SUM(金额)` unquoted in .xlsx.
 */

import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";

/** Regex source of a character that may start a bare sheet name (`u` flag). */
export const BARE_NAME_START_CLASS = "\\p{L}_";
/**
 * Regex source of a character that may follow the first one (`u` flag):
 * letters with their combining marks (Devanagari and Arabic spell words with
 * them), digits, `_` and `.`.
 */
export const BARE_NAME_PART_CLASS = "\\p{L}\\p{M}\\p{N}_.";
/**
 * A defined name may also start with a backslash, or with a number that is no
 * digit (Unicode Nl / No): workbooks Excel saved define `Ⅲ１図基礎表` and
 * `³_Unknown_value…` (xlsx-rounds R274-X10, R204-X10). A sheet name written
 * bare keeps the letter rule above.
 */
export const DEFINED_NAME_START_CLASS = `${BARE_NAME_START_CLASS}\\p{Nl}\\p{No}\\\\`;
const DEFINED_NAME_NUMBER_START_RE = /^[\p{Nl}\p{No}]$/u;

const BARE_NAME_START_RE = new RegExp(`^[${BARE_NAME_START_CLASS}]$`, "u");
const BARE_NAME_PART_RE = new RegExp(`^[${BARE_NAME_PART_CLASS}]$`, "u");

/** The code point at `src[i]` as a string (two UTF-16 units for an astral one). */
export function characterAt(src: string, i: number): string {
  const point = src.codePointAt(i);
  return point === undefined ? "" : String.fromCodePoint(point);
}

/**
 * Whether one character may appear in a sheet name written WITHOUT quotes.
 * Excel writes a sheet name bare unless it "contains spaces or non-alphabetical
 * characters" (Microsoft Support, "Create an external reference (link) to a
 * cell range in another workbook"), and alphabetical means any script. A
 * leading digit reads as a number, so a name starting with one is quoted.
 */
export function isBareNameCharacter(character: string, first: boolean): boolean {
  if (character.length === 1 && character.charCodeAt(0) < 0x80) {
    if (
      (character >= "A" && character <= "Z") ||
      (character >= "a" && character <= "z") ||
      character === "_"
    ) {
      return true;
    }
    return !first && ((character >= "0" && character <= "9") || character === ".");
  }
  return first
    ? BARE_NAME_START_RE.test(character)
    : BARE_NAME_PART_RE.test(character);
}

/** Whether `name` is spelled entirely in the unquoted sheet-name set. */
export function isBareNameSpelling(name: string): boolean {
  if (name === "") return false;
  for (let i = 0; i < name.length; ) {
    const character = characterAt(name, i);
    if (!isBareNameCharacter(character, i === 0)) return false;
    i += character.length;
  }
  return true;
}

/** A name that is also an A1 cell on Excel's grid: `A1`, `FY2024`, `xfd1`. */
function sheetNameLooksLikeA1Cell(name: string): boolean {
  const match = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(name);
  if (!match) return false;
  let column = 0;
  for (const letter of match[1].toUpperCase()) {
    column = column * 26 + (letter.charCodeAt(0) - 64);
  }
  const row = Number(match[2]);
  return column <= MAX_GRID_COLUMNS && row >= 1 && row <= MAX_GRID_ROWS;
}

/**
 * Whether a reference to sheet `name` must wrap it in single quotes — the one
 * rule every writer of a sheet prefix shares with the formula reader.
 *
 * Excel quotes a name that has spaces or other non-alphabetical characters or
 * starts with a digit (Microsoft Support, "Create an external reference (link)
 * to a cell range in another workbook"), and also a name that would otherwise
 * read as something else: an A1 cell (`'FY2024'!A1`), an R1C1 reference (`R`,
 * `C`, `RC`, `R2C3`) or a boolean. Everything else — `Sheet2`, `回款`,
 * `Données` — is written bare, which is how Excel stores it in .xlsx.
 */
export function sheetNameNeedsQuotes(name: string): boolean {
  if (!isBareNameSpelling(name)) return true;
  return (
    sheetNameLooksLikeA1Cell(name) ||
    /^(?:R[0-9]*)?(?:C[0-9]*)?$/i.test(name) ||
    /^(?:TRUE|FALSE)$/i.test(name)
  );
}

/** `name` as a formula writes it before `!`: bare, or quoted with `'` doubled. */
export function formatSheetNameForFormula(name: string): string {
  return sheetNameNeedsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

/**
 * Split an optional leading SHEET QUALIFIER off a reference, as a formula
 * writes one (`Sheet2!A1:A10`, `回款!$C$4`, `'My Sheet'!$A$1`). A quoted name
 * unescapes doubled single-quotes (`'O''Brien'` → `O'Brien`). No qualifier ⇒
 * no sheet (the common same-sheet case), the text returned unchanged.
 */
export function splitSheetQualifier(text: string): { sheet?: string; ref: string } {
  const quoted = /^'((?:[^']|'')*)'!(.*)$/.exec(text);
  if (quoted) return { sheet: quoted[1].replace(/''/g, "'"), ref: quoted[2].trim() };
  // Bare in any script: Excel stores a reference to sheet 列表 as
  // `列表!$A$1:$A$5`, unquoted.
  const bang = text.indexOf("!");
  const bare = bang > 0 ? text.slice(0, bang) : "";
  if (isBareNameSpelling(bare)) return { sheet: bare, ref: text.slice(bang + 1).trim() };
  return { ref: text };
}

/**
 * Typographic quotation marks Excel reads as characters of a name, first or
 * later: a workbook Excel saved defines the names “ and ”, `=“` reads the
 * cell “ names, and an undefined ‘ is #NAME? -- never a syntax error
 * (S17_tdf165886, X-200). They are never string delimiters.
 */
const DEFINED_NAME_QUOTE_MARKS = new Set(["\u2018", "\u2019", "\u201C", "\u201D"]);










const DEFINED_NAME_CJK_MARKS = ["\u3010", "\u3011", "\u300C", "\u300D", "\u30FB", "\uFF3F"];
/** Regex source (`u` flag) of the marks a defined name may hold after its first character. */
export const DEFINED_NAME_LATER_MARK_CLASS = "?\\\\\\u3010\\u3011\\u300C\\u300D\\u30FB\\uFF3F";

/**
 * Characters Excel itself stores in a defined name after its first one, beyond
 * the documented set: a question mark and a backslash (a workbook Excel saved
 * defines `Profile?_YES_NO_filter`, xlsx-rounds R268-X15). Neither is an
 * operator anywhere in a formula, so a name holding one reads as one token.
 */
const DEFINED_NAME_LATER_MARKS = new Set(["?", "\\", ...DEFINED_NAME_CJK_MARKS]);

/**
 * Whether one character may appear in a defined name. Microsoft Support,
 * "Names in formulas": "The first character of a name must be a letter, an
 * underscore character (_), or a backslash (\). Remaining characters in the
 * name can be letters, numbers, periods, and underscore characters." That is
 * the bare sheet-name set plus a leading backslash, and the `?` and `\` Excel
 * also stores after the first character.
 */
export function isDefinedNameCharacter(character: string, first: boolean): boolean {
  return (
    (first && character === "\\") ||
    (first && DEFINED_NAME_NUMBER_START_RE.test(character)) ||
    (!first && DEFINED_NAME_LATER_MARKS.has(character)) ||
    DEFINED_NAME_QUOTE_MARKS.has(character) ||
    isBareNameCharacter(character, first)
  );
}

/**
 * The index just past the defined-name-shaped run starting at `src[start]`, or
 * `start` itself when no name starts there.
 */
export function definedNameEnd(src: string, start: number): number {
  const first = characterAt(src, start);
  if (!first || !isDefinedNameCharacter(first, true)) return start;
  let end = start + first.length;
  while (end < src.length) {
    const character = characterAt(src, end);
    if (!isDefinedNameCharacter(character, false)) break;
    end += character.length;
  }
  return end;
}

/** Whether every character of `name` is one a defined name may use there. */
export function isDefinedNameSpelling(name: string): boolean {
  return name !== "" && definedNameEnd(name, 0) === name.length;
}

/**
 * Whether `name` reads as a cell reference, which Excel refuses as a name
 * ("Names can't be the same as a cell reference, such as Z$100 or R1C1", and
 * "C", "c", "R" and "r" are the Name Box's row / column shorthands). The A1
 * test is deliberately the formula tokenizer's shape — one to three letters and
 * up to seven digits — so a name is never one that a formula would read as a
 * cell. Letters and digits past Excel's grid are no cell, to Excel and to the
 * formula reader (`cellRefInGrid`): workbooks Excel saved define `a0`, `Mar00`
 * (row 0), `qtr1201314` (row 1,201,314) and `YTD1415` (column 17,424)
 * (xlsx-rounds R202-X04). The R1C1 test covers `R`, `C`, `RC`, `R2`, `C3` and
 * `R2C3`.
 */
export function looksLikeCellReference(name: string): boolean {
  const a1 = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(name);
  if (a1) {
    let column = 0;
    for (const letter of a1[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
    const row = Number(a1[2]);
    if (column <= MAX_GRID_COLUMNS && row >= 1 && row <= MAX_GRID_ROWS) return true;
  }
  return /^(?:R[0-9]*)?(?:C[0-9]*)?$/i.test(name);
}
