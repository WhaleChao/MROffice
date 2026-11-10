/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Error Checking (公式 ▸ 错误检查).
 *
 * Excel and WPS scan a sheet for formulas that are probably wrong, not just
 * formulas that already failed, and walk you through them one at a time. This
 * module is the scan: cells in, ranked findings out. It is pure so the same
 * rules can drive the dialog, a status-bar count, and tests.
 *
 * The rules are Excel's error checking rules (File > Options > Formulas;
 * Microsoft Support, "Detect errors in formulas"), in that list's order:
 *
 *   - `error-value` — the formula evaluates to #DIV/0!, #REF!, #VALUE!, …
 *   - `inconsistent-calculated-column` — a cell of a table's calculated
 *     column that does not hold the column formula: a typed value, a blank
 *     or another formula (Microsoft Support, "Inconsistent calculated column
 *     formula": the cell is an EXCEPTION, the column stays calculated, and
 *     the error button offers "Restore to Calculated Column Formula")
 *   - `two-digit-year` — a text date with a two-digit year, in a cell or a
 *     formula's text ("1/1/31" may be read as 1931 or 2031; XLS-A075)
 *   - `number-stored-as-text` — a numeric literal held as text, which silently
 *     drops out of SUM
 *   - `inconsistent-formula` — a formula whose RELATIVE shape differs from the
 *     matching neighbours on both sides (the classic "one cell in the row was
 *     edited by hand" bug)
 *   - `omits-adjacent-cells` — an aggregate whose range stops short of numbers
 *     that sit directly against it
 *   - `unlocked-formula` — a formula in a cell that is not locked ("Why do I
 *     get an Unprotected Formula message": Excel flags it whether or not the
 *     sheet is protected yet, since the lock is what protection would honour)
 *   - `formula-references-empty` — a reference into a cell that holds nothing
 *   - `invalid-table-data` — a table cell whose value breaks its data
 *     validation rule ("There is a validation error in a table"; XLS-A075)
 *
 * Excel ships every rule ON except "Formulas referring to empty cells"
 * (`EXCEL_DEFAULT_DISABLED_ERROR_RULES`); `checkForErrors` itself checks all
 * of them unless told otherwise. A cell's own Ignore Error choices
 * (`ErrorCheckCell.ignored`, xlsx `<ignoredError>`) silence single rules on
 * that cell only.
 *
 * Every finding names the cell and says what it saw, so the host never has to
 * re-derive the reason to explain it. The same per-cell rules drive the Error
 * Checking dialog (`checkForErrors`, a whole-sheet walk) and the grid's green
 * triangles (`checkCellForErrors`, one rendered cell at a time).
 */

import { errorCodeOfDisplay } from "./error-display";
import {
  columnToLetters,
  isExcelNumericText,
  shiftFormula,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

export type ErrorCheckRule =
  | "error-value"
  | "inconsistent-calculated-column"
  | "inconsistent-formula"
  | "omits-adjacent-cells"
  | "number-stored-as-text"
  | "formula-references-empty"
  | "unlocked-formula"
  | "two-digit-year"
  | "invalid-table-data";

export interface ErrorCheckCell {
  row: number;
  col: number;
  /** Raw cell text: a formula starts with "=". */
  raw: string;
  /** What the grid shows — an error string here is what `error-value` reports. */
  display?: string;
  /** False when the cell stays editable while the sheet is protected. */
  locked?: boolean;
  /**
   * True when the cell is TYPED as text — a leading apostrophe or an `@`
   * (text) number format. A digits-only cell is not a finding on its own:
   * Excel flags a number kept as TEXT, which is the one that silently drops
   * out of SUM, not every number on the sheet.
   */
  textFormatted?: boolean;
  /** Rules this one cell was told to ignore (Ignore Error). */
  ignored?: readonly ErrorCheckRule[];
  /**
   * The formula a table's calculated column expects in THIS cell (the column
   * formula shifted to its row, leading `=`). Set only for a data cell of a
   * calculated column; a cell whose `raw` differs is a calculated-column
   * exception (`calculatedColumnFormulaAt`).
   */
  calculatedColumn?: string;
  /** A cell of a table whose value fails its data validation rule. */
  invalidTableData?: boolean;
}

export interface ErrorCheckFinding {
  rule: ErrorCheckRule;
  row: number;
  col: number;
  /** The cell's formula or literal, for the dialog's "found" line. */
  detail: string;
  /**
   * `inconsistent-formula` only: the neighbour whose formula the region
   * agrees on — above when the column pair agrees, else left — which is what
   * Excel's "Copy Formula from Above / Left" copies.
   */
  source?: { row: number; col: number };
}

export interface ErrorCheckOptions {
  /** Rules the user switched off. */
  disabled?: readonly ErrorCheckRule[];
  /** Stop after this many findings so a broken sheet cannot hang the dialog. */
  limit?: number;
}

export const DEFAULT_ERROR_CHECK_LIMIT = 500;

/** Every rule, in the order Excel's Options list names them. */
export const ERROR_CHECK_RULES: readonly ErrorCheckRule[] = [
  "error-value",
  "inconsistent-calculated-column",
  "two-digit-year",
  "number-stored-as-text",
  "inconsistent-formula",
  "omits-adjacent-cells",
  "unlocked-formula",
  "formula-references-empty",
  "invalid-table-data",
];

/**
 * The rules a fresh Excel install leaves unticked: only "Formulas referring
 * to empty cells". Hosts pass this as `disabled` to check what Excel checks.
 */
export const EXCEL_DEFAULT_DISABLED_ERROR_RULES: readonly ErrorCheckRule[] = [
  "formula-references-empty",
];





export interface IgnoredErrorRange {
  rules: string[];
  range: { r1: number; c1: number; r2: number; c2: number };
  areas?: Array<{ r1: number; c1: number; r2: number; c2: number }>;
}

/** The rules the sheet's ranges ignore at a cell: Excel ignores a rule there
 *  whether the cell carries it or a range covering it does (a value typed
 *  into a blank of an ignored range stays ignored). */
export function ignoredErrorRulesInRanges(
  ranges: unknown,
  row: number,
  col: number,
): ErrorCheckRule[] | undefined {
  if (!Array.isArray(ranges) || ranges.length === 0) return undefined;
  const rules = new Set<string>();
  for (const entry of ranges as IgnoredErrorRange[]) {
    if (!entry || !Array.isArray(entry.rules)) continue;
    const areas = Array.isArray(entry.areas) && entry.areas.length > 0 ? entry.areas : [entry.range];
    if (areas.some((area) => area && row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2)) {
      for (const rule of entry.rules) rules.add(rule);
    }
  }
  return rules.size ? normalizeIgnoredErrorRules([...rules]) : undefined;
}

/** A persisted Ignore Error list, validated and in canonical order. */
export function normalizeIgnoredErrorRules(
  value: unknown,
): ErrorCheckRule[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const rules = ERROR_CHECK_RULES.filter((rule) => value.includes(rule));
  return rules.length > 0 ? rules : undefined;
}

/** The fields of a stored grid cell the rules read. */
export interface ErrorCheckSourceCell {
  value?: unknown;
  /** The stored kind: an imported `t="s"` cell is text however it looks. */
  kind?: unknown;
  locked?: unknown;
  style?: { numberFormat?: unknown; quotePrefix?: unknown } | null;
  ignoredErrors?: unknown;
}

/**
 * Project a stored cell for the rules, so the dialog and the grid read cells
 * the same way. `display` is the live result, and matters only for the cell
 * being checked: a neighbour is read for its raw text alone.
 */
export function errorCheckCellFrom(
  row: number,
  col: number,
  source: ErrorCheckSourceCell | undefined,
  display?: string,
  calculatedColumn?: string,
  invalidTableData?: boolean,
  /** The rules the sheet's ignore ranges cover the cell with
   *  ({@link ignoredErrorRulesInRanges}). */
  rangeIgnored?: readonly ErrorCheckRule[],
): ErrorCheckCell {
  const value = source?.value;
  const raw = typeof value === "string" ? value : String(value ?? "");
  const numberFormat = source?.style?.numberFormat;
  const ignored = normalizeIgnoredErrorRules([
    ...(normalizeIgnoredErrorRules(source?.ignoredErrors) ?? []),
    ...(rangeIgnored ?? []),
  ]);
  return {
    row,
    col,
    raw,
    ...(display !== undefined ? { display } : {}),
    ...(calculatedColumn !== undefined ? { calculatedColumn } : {}),
    ...(invalidTableData ? { invalidTableData: true } : {}),
    ...(source?.locked === false ? { locked: false } : {}),
    // An entry typed after an apostrophe keeps it as the quote prefix, not
    // in the value (entry-parse), and is text exactly like an "@" cell. A
    // cell stored as text (an imported shared string, a pasted-as-text entry)
    // is text too, whatever it looks like: that is the cell Excel flags.
    ...((typeof numberFormat === "string" && numberFormat.trim() === "@") ||
    source?.style?.quotePrefix === true ||
    source?.kind === "text"
      ? { textFormatted: true }
      : {}),
    ...(ignored ? { ignored } : {}),
  };
}

/**
 * The display the error-value rule reads for a formula cell. A formula whose
 * result is TEXT never "returns an error value", even when the text spells
 * one (`="#N/A"`, `=IFERROR(1/0,"#DIV/0!")`): ISERROR is FALSE there, and
 * Excel's error checking leaves the cell alone (live audit S-106). Only a
 * real error result (or a number/boolean display) is passed through.
 */
export function errorCheckFormulaDisplay(computed: unknown, display: string): string {
  return typeof computed === "string" ? "" : display;
}

/** The part of a data table the calculated-column rule reads. */
export interface ErrorCheckCalculatedTable {
  range: { r1: number; c1: number; r2: number; c2: number };
  headerRow?: boolean;
  /** Canonical FIRST-data-row formula per column offset (`DataTable`). */
  calculatedFormulas?: Record<number, string>;
}

/**
 * The formula a calculated column expects in one cell — the column formula
 * (ECMA-376 §18.5.1.3 `calculatedColumnFormula`, written for the first data
 * row) shifted to that row — or undefined when the cell is not a data cell of
 * a calculated column. A cell that holds anything else is an exception.
 */
export function calculatedColumnFormulaAt(
  tables: readonly ErrorCheckCalculatedTable[] | undefined,
  row: number,
  col: number,
): string | undefined {
  if (!tables) return undefined;
  for (const table of tables) {
    const formulas = table.calculatedFormulas;
    if (!formulas) continue;
    const { range } = table;
    if (col < range.c1 || col > range.c2) continue;
    const firstDataRow = range.r1 + (table.headerRow ? 1 : 0);
    if (row < firstDataRow || row > range.r2) continue;
    const canonical = formulas[col - range.c1];
    if (typeof canonical !== "string" || !canonical.startsWith("=")) continue;
    return shiftFormula(canonical, row - firstDataRow, 0);
  }
  return undefined;
}

const ERROR_VALUE_RE =
  /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)$/;
// A1 must not be the tail of a function name: `LOG10(` and `ATAN2(` both
// contain letters followed by digits. The leading group consumes whatever
// separator precedes the reference (or nothing at the start), and the lookahead
// rejects a match that runs into more identifier characters or an opening
// paren. Written without lookbehind so every browser target can run it.
const REFERENCE_RE =
  /(^|[^A-Za-z0-9_$.])\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?![A-Za-z0-9_(])/g;
const AGGREGATE_RE =
  /\b(SUM|AVERAGE|COUNT|COUNTA|MIN|MAX|PRODUCT|MEDIAN|STDEV|VAR)\s*\(\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\s*:\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\s*\)/i;

function columnToIndex(letters: string): number {
  let index = 0;
  for (const char of letters.toUpperCase()) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

/**
 * The runs of a formula that are NOT inside a string literal, each with the
 * offset it starts at. Everything this module reads — references, ranges — is
 * meaningless inside a literal, and a literal that happens to read like a
 * reference (`="see B2 for detail"`) otherwise produces a finding about a cell
 * the formula never touches.
 */
function nonLiteralSegments(
  body: string,
): Array<{ text: string; start: number }> {
  const segments: Array<{ text: string; start: number }> = [];
  let start = 0;
  let inLiteral = false;
  for (let index = 0; index < body.length; index += 1) {
    if (body[index] !== '"') continue;
    if (!inLiteral) {
      segments.push({ text: body.slice(start, index), start });
      inLiteral = true;
      continue;
    }
    // A doubled quote inside a literal is an escaped quote, not its end.
    if (body[index + 1] === '"') {
      index += 1;
      continue;
    }
    inLiteral = false;
    start = index + 1;
  }
  // An unterminated literal leaves nothing to scan, which is the right answer.
  if (!inLiteral) segments.push({ text: body.slice(start), start });
  return segments;
}

/**
 * The cells a formula reads DIRECTLY. Not the text inside string literals, not
 * another sheet's cells (`Sheet2!A1` names a cell this sheet's grid knows
 * nothing about), and not the endpoints of a range — `=SUM(A1:A3)` reads a
 * range, and whether it should have covered more cells is the aggregate rule's
 * question, not the empty-cell rule's.
 */
function directReferences(body: string): Array<{ row: number; col: number }> {
  const references: Array<{ row: number; col: number }> = [];
  for (const segment of nonLiteralSegments(body)) {
    for (const match of segment.text.matchAll(REFERENCE_RE)) {
      const lead = match[1];
      if (lead === "!" || lead === ":") continue;
      const end = (match.index ?? 0) + match[0].length;
      if (segment.text[end] === ":") continue;
      references.push({
        row: Number(match[3]) - 1,
        col: columnToIndex(match[2]),
      });
    }
  }
  return references;
}

/**
 * A formula's shape with every reference rewritten as an offset from its own
 * cell — two cells "have the same formula" when their shapes match, which is
 * exactly how Excel decides a formula is inconsistent with its neighbours.
 * String literals are copied through untouched: rewriting one would make the
 * same formula in two rows look like two different formulas.
 */
export function relativeFormulaShape(
  formula: string,
  row: number,
  col: number,
): string {
  const upper = formula.toUpperCase();
  const rewrite = (text: string) =>
    text.replace(
      REFERENCE_RE,
      (_match, lead: string, letters: string, digits: string) => {
        const refRow = Number(digits) - 1;
        const refCol = columnToIndex(letters);
        return `${lead}R[${refRow - row}]C[${refCol - col}]`;
      },
    );

  let shape = "";
  let cursor = 0;
  for (const segment of nonLiteralSegments(upper)) {
    shape += upper.slice(cursor, segment.start) + rewrite(segment.text);
    cursor = segment.start + segment.text.length;
  }
  return shape + upper.slice(cursor);
}

function isFormula(cell: ErrorCheckCell | undefined): cell is ErrorCheckCell {
  return Boolean(cell && cell.raw.startsWith("="));
}

function looksNumeric(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  return /^[+-]?(?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?$/.test(
    trimmed,
  );
}

/**
 * The number a text entry would have been had it been typed into a General
 * cell: surrounding spaces dropped, thousands separators in their proper
 * groups, and a trailing percent sign scaling by 1/100. The digits themselves
 * must pass the formula engine's own numeric-text test (`isExcelNumericText`),
 * so what the rule flags and what Convert to Number writes are both exactly
 * what a formula would have read as a number.
 */
function parseStoredNumber(
  text: string,
): { value: number; percent: boolean; decimals: boolean } | null {
  let body = text.trim();
  const percent = body.endsWith("%");
  if (percent) body = body.slice(0, -1).trim();
  if (body.includes(",")) {
    if (!/^[+-]?[0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]*)?$/.test(body)) return null;
    body = body.replace(/,/g, "");
  }
  if (!body || !isExcelNumericText(body)) return null;
  const value = Number(body) / (percent ? 100 : 1);
  if (!Number.isFinite(value)) return null;
  return { value, percent, decimals: /\.[0-9]*[1-9]/.test(body) };
}

/** The text of a literal that is kept as text: after `'`, or in an `@` cell. */
function textKeptAsText(cell: ErrorCheckCell): string | null {
  if (cell.raw.startsWith("=")) return null;
  if (cell.raw.startsWith("'")) return cell.raw.slice(1);
  return cell.textFormatted ? cell.raw : null;
}

// XLS-A075 "Cells containing years represented as 2 digits": a text date
// whose year has two digits — 1/1/31, 3-14-12, 14-Mar-12. Excel's fix writes
// the year as 19XX or 20XX.
const TWO_DIGIT_NUMERIC_DATE = /^(\s*(\d{1,2})([/\-.])(\d{1,2})\3)(\d{2})(\s*)$/;
const TWO_DIGIT_NAMED_DATE =
  /^(\s*\d{1,2}[\s-](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s-])(\d{2})(\s*)$/i;

/** The text with its two-digit year written in full, or null when it is not such a date. */
function fullYearDate(text: string, century: 1900 | 2000): string | null {
  const numeric = TWO_DIGIT_NUMERIC_DATE.exec(text);
  if (numeric) {
    const a = Number(numeric[2]);
    const b = Number(numeric[4]);
    const plausible = a >= 1 && b >= 1 && ((a <= 12 && b <= 31) || (a <= 31 && b <= 12));
    if (!plausible) return null;
    return `${numeric[1]}${century / 100}${numeric[5]}${numeric[6]}`;
  }
  const named = TWO_DIGIT_NAMED_DATE.exec(text);
  return named ? `${named[1]}${century / 100}${named[2]}${named[3]}` : null;
}

const FORMULA_TEXT_RE = /"((?:[^"]|"")*)"/g;

/**
 * Excel's Convert XX to 19XX / 20XX for a `two-digit-year` cell: the cell's
 * text, or every such date in its formula's text, with the year in full.
 * Null when there is nothing to convert.
 */
export function twoDigitYearFix(cell: ErrorCheckCell, century: 1900 | 2000): string | null {
  if (cell.raw.startsWith("=")) {
    let changed = false;
    const next = cell.raw.replace(FORMULA_TEXT_RE, (whole, inner: string) => {
      const full = fullYearDate(inner.replace(/""/g, '"'), century);
      if (full === null) return whole;
      changed = true;
      return `"${full.replace(/"/g, '""')}"`;
    });
    return changed ? next : null;
  }
  const text = textKeptAsText(cell);
  if (text === null) return null;
  const full = fullYearDate(text, century);
  return full === null ? null : full;
}

/**
 * Excel's Convert to Number for one `number-stored-as-text` cell: the literal
 * to write (canonical, as General shows it) and, for a percentage, the percent
 * format Excel gives a typed "12%". Null when the cell is not a number kept as
 * text.
 */
export function numberStoredAsTextValue(
  cell: ErrorCheckCell,
): { value: string; numberFormat?: string } | null {
  const text = textKeptAsText(cell);
  const parsed = text === null ? null : parseStoredNumber(text);
  if (!parsed) return null;
  // Number(String(x)) round-trips; General shows at most 15 significant digits.
  const value = String(Number(parsed.value.toPrecision(15)));
  return parsed.percent
    ? { value, numberFormat: parsed.decimals ? "0.00%" : "0%" }
    : { value };
}

/** Reads a neighbouring cell's stored text; undefined when it is empty. */
export type ErrorCheckCellAt = (
  row: number,
  col: number,
) => ErrorCheckCell | undefined;

/**
 * The aggregate range that stops short, and how far its data really reaches.
 * Detection needs only the first omitted cell (`reach` false); the fix walks
 * the whole run.
 */
function omittedAggregate(body: string, at: ErrorCheckCellAt, reach: boolean) {
  const aggregate = AGGREGATE_RE.exec(body);
  if (!aggregate) return null;
  const startRow = Number(aggregate[3]) - 1;
  const startCol = columnToIndex(aggregate[2]);
  const endRow = Number(aggregate[5]) - 1;
  const endCol = columnToIndex(aggregate[4]);
  const sameColumn = startCol === endCol;
  const sameRow = startRow === endRow;
  if (!sameColumn && !sameRow) return null;
  const numberAt = (row: number, col: number) => {
    const neighbour = row >= 0 && col >= 0 ? at(row, col) : undefined;
    return Boolean(
      neighbour &&
        !neighbour.raw.startsWith("=") &&
        looksNumeric(neighbour.raw),
    );
  };
  const firstRow = Math.min(startRow, endRow);
  const firstCol = Math.min(startCol, endCol);
  let reachRow = firstRow;
  let reachCol = firstCol;
  // A column range grows upward and a row range leftward, over the run of
  // numbers that sits directly against it (one cell is all it takes to flag).
  while (sameColumn && numberAt(reachRow - 1, startCol)) {
    reachRow -= 1;
    if (!reach) break;
  }
  while (!sameColumn && numberAt(startRow, reachCol - 1)) {
    reachCol -= 1;
    if (!reach) break;
  }
  if (reachRow === firstRow && reachCol === firstCol) return null;
  return { aggregate, sameColumn, reachRow, reachCol };
}

/**
 * Excel's "Update Formula to Include Cells": the same formula with its
 * aggregate range stretched over the adjacent numbers it left out. The
 * range's start keeps its `$` anchors. Null when nothing is omitted.
 */
export function includeAdjacentCellsFormula(
  cell: ErrorCheckCell,
  at: ErrorCheckCellAt,
): string | null {
  if (!cell.raw.startsWith("=")) return null;
  const body = cell.raw.slice(1);
  const omitted = omittedAggregate(body, at, true);
  if (!omitted) return null;
  const { aggregate, sameColumn, reachRow, reachCol } = omitted;
  const [whole] = aggregate;
  const open = whole.indexOf("(");
  const colon = whole.indexOf(":", open);
  const first = whole.slice(open + 1, colon);
  const second = whole.slice(colon + 1, whole.lastIndexOf(")"));
  const rewrite = (reference: string, row: number, col: number) =>
    reference.replace(
      /^(\s*)(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(\s*)$/,
      (_match, lead, colAnchor, _letters, rowAnchor, _digits, tail) =>
        `${lead}${colAnchor}${columnToLetters(col)}${rowAnchor}${row + 1}${tail}`,
    );
  // The range may be written either way round; the endpoint nearer the data
  // (the smaller row or column) is the one that moves.
  const firstRef = /^\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\s*$/.exec(first);
  const secondRef = /^\s*\$?([A-Za-z]{1,3})\$?([0-9]{1,7})\s*$/.exec(second);
  if (!firstRef || !secondRef) return null;
  const firstIsStart = sameColumn
    ? Number(firstRef[2]) <= Number(secondRef[2])
    : columnToIndex(firstRef[1]) <= columnToIndex(secondRef[1]);
  const moved = firstIsStart ? first : second;
  const movedRef = firstIsStart ? firstRef : secondRef;
  const nextRef = sameColumn
    ? rewrite(moved, reachRow, columnToIndex(movedRef[1]))
    : rewrite(moved, Number(movedRef[2]) - 1, reachCol);
  const range = firstIsStart
    ? `${nextRef}:${second}`
    : `${first}:${nextRef}`;
  const replaced = `${whole.slice(0, open + 1)}${range}${whole.slice(whole.lastIndexOf(")"))}`;
  const index = aggregate.index ?? 0;
  return `=${body.slice(0, index)}${replaced}${body.slice(index + whole.length)}`;
}

/**
 * Every rule that fires on one cell, in Excel's Options order. `skip` holds
 * the rules not to check: switched off, or ignored on this cell. A cell whose
 * formula already failed reports that and nothing else, as Excel does.
 */
function cellFindings(
  cell: ErrorCheckCell,
  at: ErrorCheckCellAt,
  disabled: ReadonlySet<ErrorCheckRule>,
  first: boolean,
): ErrorCheckFinding[] {
  const skip = cell.ignored?.length
    ? new Set([...disabled, ...cell.ignored])
    : disabled;
  const findings: ErrorCheckFinding[] = [];
  const found = (
    rule: ErrorCheckRule,
    extra: Partial<ErrorCheckFinding> = {},
  ) => {
    findings.push({
      rule,
      row: cell.row,
      col: cell.col,
      detail: cell.raw,
      ...extra,
    });
    return first;
  };

  const display = (cell.display ?? "").trim();
  if (ERROR_VALUE_RE.test(errorCodeOfDisplay(display)) && !skip.has("error-value")) {
    found("error-value", { detail: cell.raw || display });
    return findings;
  }

  // A calculated column's cell that does not hold the column formula — a
  // typed value, a blank, a formula of its own — is an exception: Excel keeps
  // it, keeps the column formula, and flags the cell (Microsoft Support,
  // "Inconsistent calculated column formula"). Compared as text: Excel treats
  // an edited formula as inconsistent even when it merely spells the column
  // formula differently.
  if (
    !skip.has("inconsistent-calculated-column") &&
    cell.calculatedColumn !== undefined &&
    cell.raw.trim() !== cell.calculatedColumn
  ) {
    if (found("inconsistent-calculated-column")) return findings;
  }

  if (!skip.has("two-digit-year") && twoDigitYearFix(cell, 1900) !== null) {
    if (found("two-digit-year")) return findings;
  }

  const invalidTableData = () => {
    if (!skip.has("invalid-table-data") && cell.invalidTableData) found("invalid-table-data");
  };

  if (!cell.raw.startsWith("=")) {
    const text = textKeptAsText(cell);
    if (
      text !== null &&
      !skip.has("number-stored-as-text") &&
      parseStoredNumber(text)
    ) {
      if (found("number-stored-as-text")) return findings;
    }
    invalidTableData();
    return findings;
  }

  const body = cell.raw.slice(1);

  // Inconsistent with the neighbours on BOTH sides (row first, then column),
  // which is what distinguishes a hand-edited cell from a legitimate edge.
  if (!skip.has("inconsistent-formula")) {
    const shape = relativeFormulaShape(body, cell.row, cell.col);
    const neighbourPairs: Array<
      [ErrorCheckCell | undefined, ErrorCheckCell | undefined]
    > = [
      [at(cell.row, cell.col - 1), at(cell.row, cell.col + 1)],
      [at(cell.row - 1, cell.col), at(cell.row + 1, cell.col)],
    ];
    for (const [before, after] of neighbourPairs) {
      if (!isFormula(before) || !isFormula(after)) continue;
      const beforeShape = relativeFormulaShape(
        before.raw.slice(1),
        before.row,
        before.col,
      );
      const afterShape = relativeFormulaShape(
        after.raw.slice(1),
        after.row,
        after.col,
      );
      if (beforeShape === afterShape && beforeShape !== shape) {
        if (found("inconsistent-formula", {
          source: { row: before.row, col: before.col },
        })) {
          return findings;
        }
        break;
      }
    }
  }

  // An aggregate whose range stops short of adjacent numbers.
  if (!skip.has("omits-adjacent-cells") && omittedAggregate(body, at, false)) {
    if (found("omits-adjacent-cells")) return findings;
  }

  if (!skip.has("unlocked-formula") && cell.locked === false) {
    if (found("unlocked-formula")) return findings;
  }

  // A reference into a cell with nothing in it.
  if (!skip.has("formula-references-empty")) {
    for (const reference of directReferences(body)) {
      const target = at(reference.row, reference.col);
      if (!target || target.raw.trim() === "") {
        if (found("formula-references-empty")) return findings;
        break;
      }
    }
  }

  invalidTableData();
  return findings;
}

/**
 * The one problem the grid marks on a cell — its green triangle — or null.
 * `at` reads the neighbours the rules compare against; it may leave out
 * `display`, which only the checked cell's own result needs.
 */
export function checkCellForErrors(
  cell: ErrorCheckCell,
  at: ErrorCheckCellAt,
  options: Pick<ErrorCheckOptions, "disabled"> = {},
): ErrorCheckFinding | null {
  return (
    cellFindings(cell, at, new Set(options.disabled ?? []), true)[0] ?? null
  );
}

/** Scan a sheet and rank what a user should look at. */
export function checkForErrors(
  cells: readonly ErrorCheckCell[],
  options: ErrorCheckOptions = {},
): ErrorCheckFinding[] {
  const disabled = new Set(options.disabled ?? []);
  const limit = options.limit ?? DEFAULT_ERROR_CHECK_LIMIT;
  const byKey = new Map(cells.map((cell) => [`${cell.row}:${cell.col}`, cell]));
  const at = (row: number, col: number) => byKey.get(`${row}:${col}`);
  const findings: ErrorCheckFinding[] = [];

  const ordered = [...cells].sort((a, b) =>
    a.row === b.row ? a.col - b.col : a.row - b.row,
  );

  for (const cell of ordered) {
    // The limit exists so a broken sheet cannot hang the dialog, which means
    // stopping the SCAN and not merely capping the list it returns.
    if (findings.length >= limit) break;
    for (const finding of cellFindings(cell, at, disabled, false)) {
      if (findings.length >= limit) break;
      findings.push(finding);
    }
  }

  return findings;
}
