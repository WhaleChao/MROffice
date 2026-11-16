/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet data validation (Excel/WPS parity).
 *
 * A data-validation rule constrains what a user may type into a cell or a
 * rectangular range. The sheet model stores the rules as a list; this module
 * keeps the rule math (does a rule cover a cell? is a value allowed by a rule?
 * how do we coerce a persisted rule into a safe shape?) out of the React
 * component so the validator, the dropdown affordance, the toolbar/dialog and
 * the persistence layer all agree on the model and so each piece can be
 * unit-tested in isolation.
 *
 * Supported rule kinds (matching the Excel "Data > Data Validation" subset):
 *  - "list":   the value must be one of an explicit set of allowed values; the
 *              cell renders a dropdown to pick from. Typing a non-listed value
 *              is rejected.
 *  - "number": the value must be a finite number, optionally constrained by
 *              an Excel comparison operator, optionally an integer.
 *  - "text":   the value's character length must satisfy an optional Excel
 *              comparison operator.
 *  - "formula": the value is valid iff a spreadsheet FORMULA evaluated for the
 *              cell is truthy (Excel's "Custom" validation). The formula is
 *              evaluated by the host's formula engine, threaded in as a
 *              predicate (this module never imports the engine), exactly like
 *              conditional-format formula rules.
 *  - "date":   the value must parse as an ISO date (YYYY-MM-DD), optionally
 *              constrained by an Excel comparison operator, optionally
 *              restricted to certain weekdays, and optionally on-or-before /
 *              on-or-after "today".
 *  - "time":   the value must parse as 24-hour HH:MM[:SS] and satisfy an
 *              optional Excel comparison operator.
 *  - "dependent": a cascade dropdown — a list rule whose allowed options are
 *              chosen by the value of another (controlling) cell, e.g. a State
 *              column whose options depend on the Country cell in the same row.
 *              The controlling cell's value is read through the host (threaded
 *              in like the formula predicate), so the dependency stays additive.
 *
 * The last three kinds are strictly ADDITIVE: list/number/text rules — and any
 * sheet that only uses them — validate byte-for-byte as before, and the pure
 * two-argument `validateValue(rule, value)` call is unchanged for them. The new
 * formula/dependent kinds consult host callbacks. Without the required live
 * resolver they return an explicit fail-closed result instead of authorizing an
 * edit the caller cannot validate.
 *
 * Blank policy: blanks are allowed by default, matching Excel's "Ignore blank"
 * default. A rule may explicitly set `ignoreBlank: false`, in which case a blank
 * is invalid and is handled by the same error-alert policy as any other value.
 *
 * Range model: a rule's `range` is an inclusive rectangle in 0-based grid
 * coordinates ({ r1, c1, r2, c2 }), normalized so r1<=r2 and c1<=c2. A rule
 * "covers" cell (r, c) when the cell lies inside that rectangle. When several
 * rules cover the same cell the FIRST matching rule (by array order) wins, so a
 * later, narrower rule can be layered before a broad one by ordering — matching
 * the renderer's and validator's single-rule-per-cell expectation.
 */

import {
  formatSheetNameForFormula,
  splitSheetQualifier as splitSheetPrefix,
} from "./name-syntax";
import { formulaA1ToR1C1 } from "./r1c1";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { excelMaxDateSerial, excelSystemSerialToUtcMs } from "./excel-1900-calendar";
import { coalesceRanges, fillTileOffsets } from "./conditional-format";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

// A1 cell / range reference grammar for a list rule's optional SOURCE RANGE
// (Excel's "List → Source: =$A$1:$A$10"). Mirrors the formula engine / named-
// ranges grammar so the source accepts exactly what a formula reference would.
const SOURCE_CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;
const SOURCE_RANGE_RE =
  /^(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}):(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})$/;

/** "A" -> 0, "AA" -> 26; -1 for invalid. Local so this module imports no engine. */
function sourceLettersToColumn(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

/** Parse one A1 cell of a source range to 0-based { row, col }, or null. */
function parseSourceCell(token: string): { row: number; col: number } | null {
  const m = SOURCE_CELL_RE.exec(token.trim());
  if (!m) return null;
  const col = sourceLettersToColumn(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || row < 0 || !Number.isFinite(row)) return null;
  return { row, col };
}

/**
 * Parse a list rule's `sourceRange` into a normalized inclusive rectangle
 * (clamped to the grid), accepting either a single cell ("A1") or a range
 * ("A1:A10"), with or without a leading "=" and `$` absolutes. Returns null
 * when the text is not a valid reference. Exposed so the host can resolve the
 * live option set and the validator can reuse the exact same rectangle.
 */
export function parseSourceRange(raw: unknown): ValidationRange | null {
  if (typeof raw !== "string") return null;
  let text = raw.trim();
  if (text === "") return null;
  if (text.startsWith("=")) text = text.slice(1).trim();
  // A cross-sheet source ('Lists'!$A$1:$A$10) carries the values from ANOTHER
  // sheet — a very common Excel pattern (a hidden "Lists" sheet feeding dropdowns
  // on data sheets). Strip the optional sheet qualifier; an unqualified ref keeps
  // sheet === undefined and resolves on the current sheet exactly as before.
  const { sheet, ref } = splitSheetPrefix(text);
  const withSheet = (r: ValidationRange | null): ValidationRange | null =>
    r && sheet ? { ...r, sheet } : r;
  const rangeMatch = SOURCE_RANGE_RE.exec(ref);
  if (rangeMatch) {
    const from = parseSourceCell(rangeMatch[1]);
    const to = parseSourceCell(rangeMatch[2]);
    if (!from || !to) return null;
    return withSheet(
      normalizeRange({ r1: from.row, c1: from.col, r2: to.row, c2: to.col }),
    );
  }
  const cell = parseSourceCell(ref);
  if (cell) {
    return withSheet(
      normalizeRange({ r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col }),
    );
  }
  return null;
}

/** One endpoint of a source reference, with its `$` anchors. */
const SOURCE_CELL_PARTS_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

/**
 * The rectangle a list rule's `sourceRange` reads FOR the cell (row, col).
 * Excel writes the Source against the rule's top-left cell, so a relative
 * part moves with the cell: `=A1` on B1:B2 reads A1 for B1 and A2 for B2,
 * while `$A$1` stays put, `$A1` follows the row only and `A$1` the column
 * only. Without a cell (or at the anchor) it is the source as written. A
 * reference moved off the grid is Excel's #REF! and yields null.
 */
export function listSourceRangeForCell(
  rule: ListRule,
  row?: number,
  col?: number,
): ValidationRange | null {
  const written = parseSourceRange(rule.sourceRange);
  if (!written || row === undefined || col === undefined) return written;
  const rowDelta = row - rule.range.r1;
  const colDelta = col - rule.range.c1;
  if (rowDelta === 0 && colDelta === 0) return written;
  let text = (rule.sourceRange ?? "").trim();
  if (text.startsWith("=")) text = text.slice(1).trim();
  const { ref } = splitSheetPrefix(text);
  const ends: { row: number; col: number }[] = [];
  for (const token of ref.split(":")) {
    const m = SOURCE_CELL_PARTS_RE.exec(token.trim());
    if (!m) return null;
    const endCol = sourceLettersToColumn(m[2]) + (m[1] ? 0 : colDelta);
    const endRow = Number(m[4]) - 1 + (m[3] ? 0 : rowDelta);
    if (endCol < 0 || endCol >= MAX_GRID_COLUMNS || endRow < 0 || endRow >= MAX_GRID_ROWS) {
      return null;
    }
    ends.push({ row: endRow, col: endCol });
  }
  const last = ends[ends.length - 1];
  const moved = normalizeRange({ r1: ends[0].row, c1: ends[0].col, r2: last.row, c2: last.col });
  return moved && written.sheet ? { ...moved, sheet: written.sheet } : moved;
}

/**
 * A list Source that is a formula or a defined name rather than a cell
 * reference: Excel's dependent drop-down (=INDIRECT(A2), each row's list the
 * range its controlling cell names), a named list (=Departments), a growing
 * list (=OFFSET($A$1,0,0,COUNTA($A:$A),1)) or a spill (=$D$2#). Excel
 * evaluates it for each cell, relative references moved from the rule's
 * top-left cell, and lists the cells of the reference it yields. `null` for
 * a plain reference (the source range) or no source.
 */
export function listSourceFormula(rule: { sourceRange?: string }): string | null {
  const text = (rule.sourceRange ?? "").trim().replace(/^=/, "").trim();
  if (!text || parseSourceRange(text)) return null;
  return plausibleListSourceFormula(text) ? text : null;
}

/** A Source text the formula engine may evaluate: bounded, one line, not an
 * inline list ("a,b"), parentheses and quotes balanced. */
function plausibleListSourceFormula(text: string): boolean {
  if (text.length > MAX_FORMULA_LENGTH || /[\r\n\u0000]/.test(text) || text.startsWith('"')) return false;
  let depth = 0;
  let quoted = false;
  for (const ch of text) {
    if (ch === '"') quoted = !quoted;
    else if (!quoted && ch === "(") depth += 1;
    else if (!quoted && ch === ")") {
      depth -= 1;
      if (depth < 0) return false;
    } else if (!quoted && depth === 0 && ch === ",") return false;
  }
  return depth === 0 && !quoted;
}

/** What the calculation answers for a formula Source, per cell: the cells of
 * the reference it yields, non-blank, joined by U+001F (answered with the
 * sheet's other rule formulas, keyed by this text and the rule's anchor). */
export function listSourceOptionsFormula(source: string): string {
  return `=TEXTJOIN(CHAR(31),TRUE,${source})`;
}

/** The options a formula Source's answer lists: trimmed, de-duplicated, in
 * order; none for an error, a pending answer or nothing listed. */
function listFormulaAnswerOptions(answer: unknown, limit: number): string[] {
  if (typeof answer !== "string" && typeof answer !== "number") return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const part of String(answer).split("\u001f")) {
    const value = part.trim();
    if (value === "" || seen.has(value)) continue;
    seen.add(value);
    out.push(value);
    if (out.length >= limit) break;
  }
  return out;
}

/** The evaluator a formula Source is read through (ValidationContext.evalFormula). */
export type ListSourceEvaluator = (
  formula: string,
  row: number,
  col: number,
  anchorRow: number,
  anchorCol: number,
) => unknown;

/**
 * The live dropdown options for a "list" rule: when the rule carries a
 * `sourceRange` AND a `cellValueAt` reader is supplied, the options are the
 * non-blank, trimmed, de-duplicated values of that range's cells (row-major).
 * `at` is the cell the list is for: a relative source is read from that
 * cell's position ({@link listSourceRangeForCell}), so the dropdown and the
 * validator of every cell agree. A declared source range is authoritative: a
 * missing reader, invalid source or blank source returns no options instead
 * of falling back to stale literal values. Pure (reads only through the
 * callback) so the renderer's picker / datalist and the validator share one
 * fail-closed derivation.
 */
export function listRuleOptions(
  rule: ListRule,
  cellValueAt?: (row: number, col: number, sheet?: string) => unknown,
  at?: { row: number; col: number },
  // The dropdown shows at most MAX_LIST_VALUES options; the validator reads
  // the whole declared source (live audit B290).
  limit: number = MAX_LIST_VALUES,
  evalFormula?: ListSourceEvaluator,
): string[] {
  const formula = listSourceFormula(rule);
  if (formula) {
    // Read for the cell (=INDIRECT(A2) on B2:B9 reads A5 for B5).
    const cell = at ?? { row: rule.range.r1, col: rule.range.c1 };
    if (!evalFormula) return [];
    return listFormulaAnswerOptions(
      evalFormula(listSourceOptionsFormula(formula), cell.row, cell.col, rule.range.r1, rule.range.c1),
      limit,
    );
  }
  if (rule.sourceRange) {
    if (!cellValueAt) return [];
    const range = listSourceRangeForCell(rule, at?.row, at?.col);
    if (!range) return [];
    const seen = new Set<string>();
    const out: string[] = [];
    for (let r = range.r1; r <= range.r2; r += 1) {
      for (let c = range.c1; c <= range.c2; c += 1) {
        // Pass the resolved sheet (undefined ⇒ current sheet) so a cross-sheet
        // source reads the OTHER sheet's live values. Unknown/renamed sheets
        // therefore yield an empty set and cannot silently use stale values.
        const raw = cellValueAt(r, c, range.sheet);
        const value =
          typeof raw === "string"
            ? raw.trim()
            : raw == null
              ? ""
              : String(raw).trim();
        if (value === "" || seen.has(value)) continue;
        seen.add(value);
        out.push(value);
        if (out.length >= limit) return out;
      }
    }
    return out;
  }
  return rule.values;
}

/** A rectangular range in inclusive, 0-based grid coordinates. */
export interface ValidationRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /**
   * The sheet the range lives on, when the source ref was sheet-qualified
   * (`Sheet2!A1:A10`). Absent ⇒ the current sheet. Consumed by `listRuleOptions`
   * via the optional 3rd `sheet` arg of its cell reader.
   */
  sheet?: string;
}

/**
 * How an invalid entry is handled (Excel's "Error Alert" style):
 *  - "stop":        the entry is BLOCKED (reverted); the user must fix it.
 *  - "warning":     the entry is flagged but the user may OVERRIDE and keep it.
 *  - "information": purely informational; the entry is always kept.
 * Absent ⇒ "stop" (Excel's default), so a rule with no error style still blocks
 * invalid input as before.
 */
export type ErrorStyle = "stop" | "warning" | "information";

/** Excel's comparison operators shared by number/text-length/date/time rules. */
export type ValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export const VALIDATION_OPERATORS: readonly ValidationOperator[] = [
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
] as const;

/**
 * Optional Excel-style INPUT MESSAGE and ERROR ALERT carried by any rule. The
 * input message (a title + body) is shown when a cell covered by the rule is
 * selected, as guidance BEFORE the user types. The error alert (a title + body +
 * style) is shown when an invalid value is entered; `errorStyle` decides whether
 * the entry is blocked or may be overridden (see ErrorStyle). Every field is
 * optional and additive — a rule without any of them behaves byte-for-byte as
 * before (a stop-style block with a generic translated message).
 */
export interface ValidationMessages {
  /** Excel "Ignore blank"; absent/true means blanks are allowed. */
  ignoreBlank?: boolean;
  /** Whether selection guidance is shown; absent/true is Excel's default. */
  showInputMessage?: boolean;
  /** Whether invalid-entry alerts enforce/block; absent/true is the default. */
  showErrorMessage?: boolean;
  /** Title of the input-guidance tooltip (shown on selection). */
  inputTitle?: string;
  /** Body of the input-guidance tooltip. */
  inputMessage?: string;
  /** Title of the error alert (shown on an invalid entry). */
  errorTitle?: string;
  /** Body of the error alert. */
  errorMessage?: string;
  /** How an invalid entry is handled; absent ⇒ "stop". */
  errorStyle?: ErrorStyle;
}

/** Excel's type="none": no restriction, but input guidance remains editable. */
export interface AnyRule extends ValidationMessages {
  kind: "any";
  range: ValidationRange;
}

/** A "pick from a fixed set" rule. Typing anything not in `values` is rejected. */
export interface ListRule extends ValidationMessages {
  kind: "list";
  range: ValidationRange;
  /**
   * Allowed display values (compared as trimmed strings). When `sourceRange` is
   * present the allowed set comes LIVE from that range. `values` is used only
   * when no source range is wired; it is never a fallback for an unresolved
   * source identity.
   */
  values: string[];
  /**
   * Optional A1 range reference ("$A$1:$A$10" / "A1:A10") whose current cell
   * values are the dropdown options — Excel's "List → Source: =$A$1:$A$10". The
   * options are read live through the host (threaded into ValidationContext like
   * the formula/dependent kinds), so editing the source cells updates the list.
   * Absent ⇒ the rule uses its literal `values` exactly as before (byte-for-byte
   * unchanged persisted shape and validation for source-less list rules).
   */
  sourceRange?: string;
  /** Excel "In-cell dropdown"; absent/true shows the picker. */
  showDropDown?: boolean;
}

/** A numeric rule: finite number, optional [min, max], optional integer-only. */
export interface NumberRule extends ValidationMessages {
  kind: "number";
  range: ValidationRange;
  min?: number;
  max?: number;
  /**
   * A bound written as a formula or reference ("$H$15", "MAX(A:A)"), stored
   * without its leading "=" and evaluated for the validated cell (relative
   * references shift from the rule's top-left, as in Excel). Live audit B149:
   * a bound the editor could not read is kept, not silently dropped.
   */
  minFormula?: string;
  maxFormula?: string;
  /** Absent means the legacy inclusive between/one-sided bounds behavior. */
  operator?: ValidationOperator;
  /** When true, only whole numbers are accepted. */
  integer?: boolean;
}

/** A text-length rule: the entered string's length within [minLength, maxLength]. */
export interface TextRule extends ValidationMessages {
  kind: "text";
  range: ValidationRange;
  minLength?: number;
  maxLength?: number;
  operator?: ValidationOperator;
}

/**
 * A "custom" rule (Excel's formula validation): the value is allowed iff the
 * `formula` — a spreadsheet formula evaluated for the cell — is truthy. The host
 * evaluates the formula (relative references shifted to the cell, Excel-style)
 * and reuses its formula engine; this module only carries the formula text and
 * interprets the engine's result via `isFormulaResultTruthy`.
 */
export interface FormulaRule extends ValidationMessages {
  kind: "formula";
  range: ValidationRange;
  /** A spreadsheet formula (with or without a leading "="); truthy ⇒ valid. */
  formula: string;
}

/**
 * The set of weekdays a date rule allows, as 0=Sunday … 6=Saturday (JS
 * getUTCDay ordering). Absent ⇒ every weekday allowed.
 */
export type Weekday = 0 | 1 | 2 | 3 | 4 | 5 | 6;

/**
 * A date rule: the value must parse as an ISO date (YYYY-MM-DD) and satisfy the
 * optional bounds. `min`/`max` are inclusive ISO date strings. `weekdays`, when
 * present and non-empty, restricts the value to those weekdays. `notBefore` /
 * `notAfter` set to "today" require the value to be on-or-after / on-or-before
 * the current date (evaluated at validation time). All constraints are ANDed.
 */
export interface DateRule extends ValidationMessages {
  kind: "date";
  range: ValidationRange;
  /** Inclusive lower bound, ISO date "YYYY-MM-DD". */
  min?: string;
  /** Inclusive upper bound, ISO date "YYYY-MM-DD". */
  max?: string;
  operator?: ValidationOperator;
  /** Allowed weekdays (0=Sun..6=Sat); empty/absent ⇒ all weekdays allowed. */
  weekdays?: Weekday[];
  /** "today" ⇒ the value must be today or later. */
  notBefore?: "today";
  /** "today" ⇒ the value must be today or earlier. */
  notAfter?: "today";
}

/** A 24-hour time rule. Bounds are normalized HH:MM or HH:MM:SS strings. */
export interface TimeRule extends ValidationMessages {
  kind: "time";
  range: ValidationRange;
  min?: string;
  max?: string;
  operator?: ValidationOperator;
}

/**
 * A dependent (cascade) dropdown rule: a list whose allowed options are selected
 * by the value of another cell. `source` identifies the controlling cell; for a
 * rule covering a multi-row range, `sourceMode` decides whether every covered
 * cell reads the SAME absolute `source` cell ("fixed", the default) or the
 * controlling cell in its OWN row at the source column ("perRow", e.g. a
 * Country column to the left of this State column). `options` maps a controlling
 * value (trimmed, exact) to its allowed list; `optionsDefault` is used when the
 * controlling value matches no key. A value is allowed iff it is one of the
 * options derived for that cell. With no usable derived options every non-blank
 * value is rejected (nothing is allowed for that controlling value).
 */
export interface DependentRule extends ValidationMessages {
  kind: "dependent";
  range: ValidationRange;
  /** The controlling cell (absolute grid coords). */
  source: { row: number; col: number };
  /**
   * "fixed" (default): all covered cells read the absolute `source` cell.
   * "perRow": each covered cell reads (its own row, source.col), so the rule
   * cascades row-by-row like Country→State down a pair of columns.
   */
  sourceMode?: "fixed" | "perRow";
  /** Map of controlling value → allowed options for that value. */
  options: Record<string, string[]>;
  /** Options used when the controlling value matches no `options` key. */
  optionsDefault?: string[];
}

export type ValidationRule =
  | AnyRule
  | ListRule
  | NumberRule
  | TextRule
  | FormulaRule
  | DateRule
  | TimeRule
  | DependentRule;

/** The result of validating a value against a rule. */
export interface ValidationResult {
  ok: boolean;
  /** When not ok, an i18n MESSAGE KEY describing why (the UI translates it). */
  message?: string;
  /** Missing formula/source authority is a hard block, never user-overridable. */
  failure?: "resolver-unavailable";
}

/**
 * Optional cell context for the kinds that need to consult the sheet:
 *  - "formula" rules call `evalFormula(formula, row, col)` to evaluate the
 *    rule's formula FOR THE CELL (the host shifts relative references and runs
 *    its own formula engine — this module never imports it). The result is read
 *    truthily via isFormulaResultTruthy.
 *  - "dependent" rules call `cellValueAt(row, col)` to read the controlling
 *    cell's current value and derive that cell's allowed options.
 *
 * Both callbacks are optional at the type boundary so pure callers can still
 * inspect rules. A non-blank value governed by a rule that needs a missing
 * callback fails closed with `resolver-unavailable`; the host must not persist
 * an edit it cannot validate. List / number / text / date rules without a
 * source range keep their pure two-argument behavior.
 */
export interface ValidationContext {
  /** 0-based row of the cell being validated (for formula/dependent rules). */
  row?: number;
  /** 0-based column of the cell being validated (for formula/dependent rules). */
  col?: number;
  /**
   * Evaluate the formula FOR the cell at (row, col); the rule's range top-left is
   * passed as the anchor so the host can shift relative refs Excel-style. Returns
   * the engine's scalar result (number | string | boolean | null | error shape).
   */
  evalFormula?: (
    formula: string,
    row: number,
    col: number,
    anchorRow: number,
    anchorCol: number,
  ) => unknown;
  /**
   * Read any cell's current value (for a dependent rule's controlling cell, or a
   * list rule's source range). The optional 3rd `sheet` arg lets a list rule's
   * sheet-qualified source range read another sheet; same-sheet callers omit it.
   */
  cellValueAt?: (row: number, col: number, sheet?: string) => unknown;
  /** Prove a parsed source range resolves to an authoritative live sheet. */
  sourceRangeResolvable?: (range: ValidationRange) => boolean;
  /** Override "today" (UTC midnight ms) for deterministic date-rule tests. */
  now?: Date;
  /** The workbook reads serials in the 1904 date system (date rules). */
  date1904?: ExcelDateSystemArg;
}

/**
 * The options a dependent rule allows FOR a specific covered cell, given the
 * cell coordinates and a way to read the controlling cell. Pure: exposed so the
 * renderer can build the cascade dropdown's choices and the validator can reuse
 * the exact same derivation. Returns the trimmed/deduped option list (possibly
 * empty, meaning "nothing is allowed for the current controlling value").
 */
export function dependentOptionsForCell(
  rule: DependentRule,
  row: number,
  col: number,
  cellValueAt: (row: number, col: number) => unknown,
): string[] {
  const srcRow = rule.sourceMode === "perRow" ? row : rule.source.row;
  const srcCol = rule.source.col;
  // Avoid a self-reference reading the very cell being validated (which would be
  // mid-edit) — fall back to the default options in that degenerate case.
  if (srcRow === row && srcCol === col) return rule.optionsDefault ?? [];
  const raw = cellValueAt(srcRow, srcCol);
  const key = typeof raw === "string" ? raw.trim() : raw == null ? "" : String(raw).trim();
  const direct = key !== "" ? rule.options[key] : undefined;
  return direct ?? rule.optionsDefault ?? [];
}

/** Maximum number of allowed values a single list rule may carry. */
export const MAX_LIST_VALUES = 1_000;
/** A list rule's source is checked in full on entry: a value past the
 *  dropdown's first thousand options is still in the list (live audit B290). */
export const MAX_LIST_MEMBERSHIP_SCAN = MAX_GRID_ROWS;
/** Maximum number of rules a sheet may carry (DoS bound). A rule is one
 *  rectangle (the import splits a rule's sqref): Excel sets no count, and one
 *  "IME off" rule of a Japanese workbook spans 1,024 of them (xlsx-rounds
 *  R422-X46), so the bound is 10,000; `ruleForCell` finds a cell's rule by a
 *  row index, not by walking them all. */
export const MAX_VALIDATION_RULES = 10_000;
/** Upper bound for a text-length constraint (Excel caps cell text far higher). */
export const MAX_TEXT_LENGTH = 32_767;
/** Upper bound for a formula/custom rule's source length (DoS bound on parse). */
export const MAX_FORMULA_LENGTH = 1_000;
/** Max distinct controlling-value keys a dependent rule may carry (DoS bound). */
export const MAX_DEPENDENT_KEYS = 1_000;
/** Max length of an input/error message TITLE (DoS bound + UI sanity). */
export const MAX_MESSAGE_TITLE_LENGTH = 255;
/** Max length of an input/error message BODY (DoS bound + UI sanity). */
export const MAX_MESSAGE_BODY_LENGTH = 2_000;

/** The error-alert styles, in the order the dialog lists them. */
export const ERROR_STYLES: readonly ErrorStyle[] = [
  "stop",
  "warning",
  "information",
] as const;

/** Coerce a persisted comparison operator to the supported native set. */
export function normalizeValidationOperator(
  raw: unknown,
): ValidationOperator | undefined {
  return typeof raw === "string" &&
    VALIDATION_OPERATORS.includes(raw as ValidationOperator)
    ? (raw as ValidationOperator)
    : undefined;
}

/** Between/not-between are the only operators that consume a second operand. */
export function validationOperatorNeedsSecondValue(
  operator: ValidationOperator,
): boolean {
  return operator === "between" || operator === "notBetween";
}

/** Trim a raw value to a bounded non-empty string, or undefined when blank. */
function normalizeMessageText(raw: unknown, max: number): string | undefined {
  if (typeof raw !== "string") return undefined;
  const trimmed = raw.trim();
  if (trimmed === "") return undefined;
  return trimmed.slice(0, max);
}

/**
 * Coerce a raw error-alert style to one of the three kinds, or undefined when it
 * is absent/unknown. Absent ⇒ the caller treats it as "stop" (Excel's default),
 * so we only keep an EXPLICIT non-default style to keep the persisted shape lean.
 */
export function normalizeErrorStyle(raw: unknown): ErrorStyle | undefined {
  if (raw === "warning" || raw === "information") return raw;
  if (raw === "stop") return "stop";
  return undefined;
}

/**
 * Extract the optional input/error message fields off a raw rule, dropping blank
 * and out-of-shape values and capping lengths. Returns a partial object with
 * only the present fields, so a rule carrying none of them spreads to nothing
 * and serializes byte-for-byte as before.
 */
export function normalizeMessages(raw: unknown): ValidationMessages {
  if (!raw || typeof raw !== "object") return {};
  const r = raw as Record<string, unknown>;
  const out: ValidationMessages = {};
  // These three controls default ON in Excel. Persist only an explicit false so
  // existing rules keep their byte-lean shape and behavior.
  if (r.ignoreBlank === false) out.ignoreBlank = false;
  if (r.showInputMessage === false) out.showInputMessage = false;
  if (r.showErrorMessage === false) out.showErrorMessage = false;
  const inputTitle = normalizeMessageText(r.inputTitle, MAX_MESSAGE_TITLE_LENGTH);
  if (inputTitle) out.inputTitle = inputTitle;
  const inputMessage = normalizeMessageText(r.inputMessage, MAX_MESSAGE_BODY_LENGTH);
  if (inputMessage) out.inputMessage = inputMessage;
  const errorTitle = normalizeMessageText(r.errorTitle, MAX_MESSAGE_TITLE_LENGTH);
  if (errorTitle) out.errorTitle = errorTitle;
  const errorMessage = normalizeMessageText(r.errorMessage, MAX_MESSAGE_BODY_LENGTH);
  if (errorMessage) out.errorMessage = errorMessage;
  // Only keep a non-default ("warning"/"information") style; "stop" is implicit.
  const errorStyle = normalizeErrorStyle(r.errorStyle);
  if (errorStyle && errorStyle !== "stop") out.errorStyle = errorStyle;
  return out;
}

/**
 * The resolved input-message guidance for a cell (title + body), or null when
 * the rule carries no input message. Pure so the renderer can decide whether to
 * pop the guidance tooltip without re-deriving the shape.
 */
export function resolveInputMessage(
  rule: ValidationMessages,
): { title: string; message: string } | null {
  if (rule.showInputMessage === false) return null;
  const title = rule.inputTitle ?? "";
  const message = rule.inputMessage ?? "";
  if (title === "" && message === "") return null;
  return { title, message };
}

/**
 * The effective error style for a rule: the explicit one, or "stop" by default
 * (Excel parity — a rule with no error style still blocks invalid input).
 */
export function effectiveErrorStyle(rule: ValidationMessages): ErrorStyle {
  return rule.errorStyle ?? "stop";
}

/**
 * Whether an invalid entry under this rule must be BLOCKED (reverted). Only the
 * "stop" style blocks; "warning" and "information" let the user keep the value.
 * The single source of truth for the component's revert decision.
 */
export function shouldBlockInvalidEntry(rule: ValidationMessages): boolean {
  return rule.showErrorMessage !== false && effectiveErrorStyle(rule) === "stop";
}

/**
 * The user's answer to a data-validation error alert: take the entry as typed,
 * go back to editing it, or throw it away and keep what the cell held.
 */
export type ValidationAlertDecision = "accept" | "retry" | "cancel";

export interface ValidationAlertButton {
  readonly decision: ValidationAlertDecision;
  readonly label: "retry" | "yes" | "no" | "ok" | "cancel";
}

const STOP_ALERT_BUTTONS: readonly ValidationAlertButton[] = [
  { decision: "retry", label: "retry" },
  { decision: "cancel", label: "cancel" },
];
const WARNING_ALERT_BUTTONS: readonly ValidationAlertButton[] = [
  { decision: "accept", label: "yes" },
  { decision: "retry", label: "no" },
  { decision: "cancel", label: "cancel" },
];
const INFORMATION_ALERT_BUTTONS: readonly ValidationAlertButton[] = [
  { decision: "accept", label: "ok" },
  { decision: "cancel", label: "cancel" },
];

/**
 * The buttons of the error alert for an error style, in Excel's order. The
 * first is the default button (Enter) and Esc always answers "cancel".
 * Microsoft Support, "More on data validation": Stop offers Retry or Cancel;
 * Warning asks "Continue?" with Yes (accept the entry), No (edit it) or Cancel
 * (remove it); Information offers OK (accept) or Cancel (reject). Only Stop
 * has no way to accept, so an invalid Stop entry is never written.
 */
export function validationAlertButtons(
  style: ErrorStyle,
): readonly ValidationAlertButton[] {
  if (style === "warning") return WARNING_ALERT_BUTTONS;
  if (style === "information") return INFORMATION_ALERT_BUTTONS;
  return STOP_ALERT_BUTTONS;
}

/**
 * A blank value is an EMPTY cell: nothing, or the empty string. A string of
 * spaces is text with a length (LEN 4, ISBLANK FALSE), so "Ignore blank" does
 * not wave it through a text-length or any other rule (live audit B039). A
 * number or boolean is a value, never blank.
 */
export function isBlank(value: unknown): boolean {
  return value === null || value === undefined || value === "";
}

/**
 * Interpret a formula-engine result as a boolean predicate (Excel parity):
 * TRUE / non-zero finite number / non-empty non-error text → truthy; FALSE / 0 /
 * blank / a "#…" error code / null / any non-scalar → falsy. Mirrors the
 * conditional-format truthiness convention so a custom validation behaves like a
 * CF formula rule. Kept local so this module imports no formula engine.
 */
/**
 * The formula a number-rule bound holds, without its "=": a non-numeric,
 * non-blank bound text is a formula or reference. Null for a number or blank.
 */
export function validationBoundFormula(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const text = raw.trim();
  if (text === "") return null;
  if (Number.isFinite(Number(text))) return null;
  return text.startsWith("=") ? text.slice(1).trim() : text;
}

/** The numeric bounds of a number rule for one cell, its formula bounds
 *  evaluated; null when a formula bound yields no number. */
function resolveValidationBounds(
  rule: NumberRule,
  ctx: ValidationContext | undefined,
): { min: number | undefined; max: number | undefined } | null {
  const evaluate = (formula: string): number | null => {
    const evalFormula = ctx?.evalFormula;
    if (!evalFormula || ctx.row === undefined || ctx.col === undefined) return null;
    const result = evalFormula(
      `=${formula}`,
      ctx.row,
      ctx.col,
      rule.range.r1,
      rule.range.c1,
    );
    if (result === null || result === undefined || result === "") return 0;
    if (typeof result === "boolean") return result ? 1 : 0;
    const n = typeof result === "number" ? result : typeof result === "string" ? Number(result.trim()) : NaN;
    return Number.isFinite(n) ? n : null;
  };
  let min = rule.min;
  let max = rule.max;
  if (rule.minFormula !== undefined) {
    const n = evaluate(rule.minFormula);
    if (n === null) return null;
    min = n;
  }
  if (rule.maxFormula !== undefined) {
    const n = evaluate(rule.maxFormula);
    if (n === null) return null;
    max = n;
  }
  if (min !== undefined && max !== undefined && min > max) {
    [min, max] = [max, min];
  }
  return { min, max };
}

export function isFormulaResultTruthy(result: unknown): boolean {
  if (result === null || result === undefined) return false;
  if (typeof result === "boolean") return result;
  if (typeof result === "number") return Number.isFinite(result) && result !== 0;
  if (typeof result === "string") {
    // Excel judges a custom rule as a logical: a text result is not one, so it
    // fails — as IF("invalid",…) is #VALUE! — the same as an error (B043). The
    // literal TRUE/FALSE words are the one text form a logical arrives in.
    const s = result.trim().toUpperCase();
    return s === "TRUE";
  }
  return false; // a FormulaError object (or any other shape) does not match
}

/** Strict ISO date pattern: YYYY-MM-DD (the only date shape we author/persist). */
const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;







export function parseIsoDate(value: unknown): number | null {
  if (typeof value !== "string") return null;
  const m = ISO_DATE_RE.exec(value.trim());
  if (!m) return null;
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const ms = Date.UTC(year, month - 1, day);
  const d = new Date(ms);
  
  if (
    d.getUTCFullYear() !== year ||
    d.getUTCMonth() !== month - 1 ||
    d.getUTCDate() !== day
  ) {
    return null;
  }
  return ms;
}

/** Strict 24-hour time pattern: HH:MM with an optional :SS component. */
const ISO_TIME_RE = /^(\d{2}):(\d{2})(?::(\d{2}))?$/;

/** Parse HH:MM[:SS] to seconds since midnight, or null for an invalid time. */
export function parseIsoTime(value: unknown): number | null {
  if (typeof value !== "string") return null;
  const match = ISO_TIME_RE.exec(value.trim());
  if (!match) return null;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = match[3] === undefined ? 0 : Number(match[3]);
  if (hour > 23 || minute > 59 || second > 59) return null;
  return hour * 3600 + minute * 60 + second;
}

/** A plain decimal number as a cell stores one ("46291", "0.4375", "-2.5e3"). */
const NUMERIC_TEXT_RE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
function storedNumber(text: string): number | null {
  const trimmed = text.trim();
  if (!NUMERIC_TEXT_RE.test(trimmed)) return null;
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : null;
}









function dateValidationMs(text: string, date1904: ExcelDateSystemArg = false): number | null {
  const iso = parseIsoDate(text);
  if (iso !== null) return iso;
  const serial = storedNumber(text);
  if (serial === null || serial < 0 || serial >= excelMaxDateSerial(date1904) + 1) {
    return null;
  }
  return excelSystemSerialToUtcMs(serial, date1904);
}

/**
 * Seconds since midnight a time rule judges. Excel stores a typed 10:30 as
 * the day fraction 0.4375 and compares that number, so a stored number is a
 * day count (a date-time is past any clock bound, as in Excel); HH:MM[:SS]
 * text is still read as a clock time.
 */
function timeValidationSeconds(text: string): number | null {
  const clock = parseIsoTime(text);
  if (clock !== null) return clock;
  const days = storedNumber(text);
  if (days === null || days < 0) return null;
  return Math.round(days * 86_400 * 1_000) / 1_000;
}

/**
 * The text a rule judges for a formula's computed result, so a formula entry
 * is validated by what it evaluates to, as Excel does (`=1+1` is 2 for a
 * whole-number rule): numbers as their shortest decimal, TRUE/FALSE, an
 * error as its code (never a valid number, date or list item) and blank as
 * "".
 */
export function validationCandidateText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "#NUM!";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object" && typeof (value as { code?: unknown }).code === "string") {
    return (value as { code: string }).code;
  }
  return String(value);
}

/**
 * Compare a scalar against Excel's validation operators. Undefined bounds are
 * accepted only for the legacy/default `between` form, preserving old one-sided
 * min/max rules. Explicit operators are normalized to require their operands.
 */
export function validationValueMatchesOperator(
  value: number,
  operator: ValidationOperator,
  first?: number,
  second?: number,
): boolean {
  if (operator === "between") {
    return (first === undefined || value >= first) &&
      (second === undefined || value <= second);
  }
  if (operator === "notBetween") {
    return first !== undefined &&
      second !== undefined &&
      (value < first || value > second);
  }
  if (first === undefined) return false;
  if (operator === "equal") return value === first;
  if (operator === "notEqual") return value !== first;
  if (operator === "greaterThan") return value > first;
  if (operator === "lessThan") return value < first;
  if (operator === "greaterThanOrEqual") return value >= first;
  return value <= first;
}

/** The current date as a UTC-midnight epoch-ms (so "today" comparisons are stable). */
function todayUtcMs(now: Date = new Date()): number {
  return Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate());
}

/**
 * Normalize a raw range into a safe rectangle (r1<=r2, c1<=c2), clamped to the
 * grid bounds, or null when any coordinate is non-finite/negative. Coordinates
 * are floored to integers so a fractional persisted value cannot slip through.
 */
export function normalizeRange(raw: unknown): ValidationRange | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const r1 = Math.floor(Number(r.r1));
  const c1 = Math.floor(Number(r.c1));
  const r2 = Math.floor(Number(r.r2));
  const c2 = Math.floor(Number(r.c2));
  if (
    !Number.isFinite(r1) ||
    !Number.isFinite(c1) ||
    !Number.isFinite(r2) ||
    !Number.isFinite(c2)
  ) {
    return null;
  }
  if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return null;
  const lowR = Math.min(r1, r2);
  const highR = Math.max(r1, r2);
  const lowC = Math.min(c1, c2);
  const highC = Math.max(c1, c2);
  // Clamp to the LAST addressable cell (MAX_GRID_ROWS - 1 / MAX_GRID_COLUMNS - 1),
  // never the bare constant, which would point one cell past the grid edge —
  // matching conditional-format.ts's MAX_RANGE_ROW / MAX_RANGE_COL ceiling.
  return {
    r1: Math.min(lowR, MAX_GRID_ROWS - 1),
    c1: Math.min(lowC, MAX_GRID_COLUMNS - 1),
    r2: Math.min(highR, MAX_GRID_ROWS - 1),
    c2: Math.min(highC, MAX_GRID_COLUMNS - 1),
  };
}

/** True when (row, col) lies inside the (inclusive, normalized) range. */
export function rangeContains(
  range: ValidationRange,
  row: number,
  col: number,
): boolean {
  return (
    row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2
  );
}

/**
 * Coerce a single raw entry into a normalized ValidationRule, or null when it
 * is malformed (unknown kind, bad range, or a list with no usable values).
 * Numeric params that are non-finite are dropped (treated as "no bound"); a
 * min/max pair where min>max is swapped so the rule still makes sense. List
 * values are trimmed-deduplicated and capped at MAX_LIST_VALUES.
 */
export function normalizeRule(raw: unknown): ValidationRule | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const range = normalizeRange(r.range);
  if (!range) return null;
  // The optional input/error message fields are additive and shared by every
  // kind; pull them once and merge into whichever rule we build. A rule that
  // carries none of them produces an empty object that spreads to nothing, so
  // the persisted shape is byte-for-byte unchanged for messageless rules.
  const messages = normalizeMessages(r);

  if (r.kind === "any") {
    return { kind: "any", range, ...messages };
  }

  if (r.kind === "list") {
    const values = normalizeOptionList(r.values);
    // A list rule may source its options either from a literal `values` set OR
    // from a live SOURCE RANGE ("=$A$1:$A$10"). Keep a sourceRange only when it
    // parses to a real reference. Canonicalize the cell address, but retain
    // the sheet qualifier exactly: opening a workbook must not rewrite the
    // native formula or change the spelling passed to its source resolver.
    // A rule with neither a usable value
    // set nor a valid source range can never offer anything → drop it.
    const sourceReference = typeof r.sourceRange === "string"
      ? r.sourceRange.trim().replace(/^=/, "").trim() : "";
    // A valid A1 reference's final ! is the qualifier separator, even when
    // its quoted sheet name also contains an exclamation mark.
    const addressStart = sourceReference.lastIndexOf("!") + 1;
    const sourceRange = parseSourceRange(sourceReference)
      ? sourceReference.slice(0, addressStart) + sourceReference.slice(addressStart).toUpperCase()
      // A formula or defined-name Source (=INDIRECT(A2)) is kept as written:
      // the calculation evaluates it per cell (listSourceFormula).
      : listSourceFormula({ sourceRange: sourceReference })
        ? sourceReference
        : undefined;
    if (values.length === 0 && !sourceRange) return null;
    return {
      kind: "list",
      range,
      values,
      ...(sourceRange ? { sourceRange } : {}),
      ...(r.showDropDown === false ? { showDropDown: false } : {}),
      ...messages,
    };
  }

  if (r.kind === "number") {
    const rule: NumberRule = { kind: "number", range, ...messages };
    const operator = normalizeValidationOperator(r.operator);
    const minFormula = validationBoundFormula(r.minFormula ?? r.min);
    const maxFormula = validationBoundFormula(r.maxFormula ?? r.max);
    const min = minFormula === null ? Number(r.min) : NaN;
    const max = maxFormula === null ? Number(r.max) : NaN;
    const hasMin =
      minFormula !== null || (r.min !== undefined && r.min !== null && Number.isFinite(min));
    const hasMax =
      maxFormula !== null || (r.max !== undefined && r.max !== null && Number.isFinite(max));
    const bothNumeric = minFormula === null && maxFormula === null;
    if (operator === "notBetween") {
      if (!hasMin || !hasMax) return null;
      rule.min = bothNumeric ? Math.min(min, max) : minFormula === null ? min : undefined;
      rule.max = bothNumeric ? Math.max(min, max) : maxFormula === null ? max : undefined;
      rule.operator = operator;
    } else if (operator && operator !== "between") {
      if (!hasMin) return null;
      if (minFormula === null) rule.min = min;
      rule.operator = operator;
    } else {
      if (hasMin && hasMax) {
        rule.min = bothNumeric ? Math.min(min, max) : minFormula === null ? min : undefined;
        rule.max = bothNumeric ? Math.max(min, max) : maxFormula === null ? max : undefined;
      } else if (hasMin) {
        if (minFormula === null) rule.min = min;
      } else if (hasMax) {
        if (maxFormula === null) rule.max = max;
      }
    }
    if (rule.min === undefined) delete rule.min;
    if (rule.max === undefined) delete rule.max;
    if (minFormula !== null && (hasMin || operator === "notBetween")) rule.minFormula = minFormula;
    if (maxFormula !== null && hasMax && (operator === undefined || operator === "between" || operator === "notBetween")) {
      rule.maxFormula = maxFormula;
    }
    if (r.integer === true) rule.integer = true;
    return rule;
  }

  if (r.kind === "text") {
    const rule: TextRule = { kind: "text", range, ...messages };
    const operator = normalizeValidationOperator(r.operator);
    const minLen = Math.floor(Number(r.minLength));
    const maxLen = Math.floor(Number(r.maxLength));
    const hasMin =
      r.minLength !== undefined &&
      r.minLength !== null &&
      Number.isFinite(minLen) &&
      minLen >= 0;
    const hasMax =
      r.maxLength !== undefined &&
      r.maxLength !== null &&
      Number.isFinite(maxLen) &&
      maxLen >= 0;
    if (operator === "notBetween") {
      if (!hasMin || !hasMax) return null;
      rule.minLength = Math.min(minLen, maxLen);
      rule.maxLength = Math.min(Math.max(minLen, maxLen), MAX_TEXT_LENGTH);
      rule.operator = operator;
    } else if (operator && operator !== "between") {
      if (!hasMin) return null;
      rule.minLength = Math.min(minLen, MAX_TEXT_LENGTH);
      rule.operator = operator;
    } else {
      if (hasMin && hasMax) {
        rule.minLength = Math.min(minLen, maxLen);
        rule.maxLength = Math.min(Math.max(minLen, maxLen), MAX_TEXT_LENGTH);
      } else if (hasMin) {
        rule.minLength = Math.min(minLen, MAX_TEXT_LENGTH);
      } else if (hasMax) {
        rule.maxLength = Math.min(maxLen, MAX_TEXT_LENGTH);
      }
    }
    return rule;
  }

  if (r.kind === "formula") {
    // A custom rule needs a non-empty, bounded formula string. Trim so a
    // whitespace-only formula is rejected (it would never be a useful predicate).
    if (typeof r.formula !== "string") return null;
    const formula = r.formula.trim();
    if (formula === "" || formula.length > MAX_FORMULA_LENGTH) return null;
    return { kind: "formula", range, formula, ...messages };
  }

  if (r.kind === "date") {
    const rule: DateRule = { kind: "date", range, ...messages };
    const operator = normalizeValidationOperator(r.operator);
    // Bounds: keep only well-formed ISO dates; swap an inverted [min,max] pair.
    const minMs = parseIsoDate(r.min);
    const maxMs = parseIsoDate(r.max);
    const minIso = minMs === null ? undefined : (r.min as string).trim();
    const maxIso = maxMs === null ? undefined : (r.max as string).trim();
    if (operator === "notBetween") {
      if (minIso === undefined || maxIso === undefined) return null;
      if ((minMs as number) <= (maxMs as number)) {
        rule.min = minIso;
        rule.max = maxIso;
      } else {
        rule.min = maxIso;
        rule.max = minIso;
      }
      rule.operator = operator;
    } else if (operator && operator !== "between") {
      if (minIso === undefined) return null;
      rule.min = minIso;
      rule.operator = operator;
    } else if (minIso !== undefined && maxIso !== undefined) {
      // Reorder to the chronological [earlier, later] pair via the parsed ms.
      if ((minMs as number) <= (maxMs as number)) {
        rule.min = minIso;
        rule.max = maxIso;
      } else {
        rule.min = maxIso;
        rule.max = minIso;
      }
    } else if (minIso !== undefined) {
      rule.min = minIso;
    } else if (maxIso !== undefined) {
      rule.max = maxIso;
    }
    if (Array.isArray(r.weekdays)) {
      const seen = new Set<number>();
      const weekdays: Weekday[] = [];
      for (const d of r.weekdays) {
        const n = Math.floor(Number(d));
        if (!Number.isFinite(n) || n < 0 || n > 6 || seen.has(n)) continue;
        seen.add(n);
        weekdays.push(n as Weekday);
      }
      // Drop the field when empty or all-7 (a no-op constraint), so a rule that
      // does not actually restrict weekdays serializes without the noise.
      if (weekdays.length > 0 && weekdays.length < 7) {
        weekdays.sort((a, b) => a - b);
        rule.weekdays = weekdays;
      }
    }
    if (r.notBefore === "today") rule.notBefore = "today";
    if (r.notAfter === "today") rule.notAfter = "today";
    return rule;
  }

  if (r.kind === "time") {
    const rule: TimeRule = { kind: "time", range, ...messages };
    const operator = normalizeValidationOperator(r.operator);
    const minSeconds = parseIsoTime(r.min);
    const maxSeconds = parseIsoTime(r.max);
    const minTime = minSeconds === null ? undefined : (r.min as string).trim();
    const maxTime = maxSeconds === null ? undefined : (r.max as string).trim();
    if (operator === "notBetween") {
      if (minTime === undefined || maxTime === undefined) return null;
      if ((minSeconds as number) <= (maxSeconds as number)) {
        rule.min = minTime;
        rule.max = maxTime;
      } else {
        rule.min = maxTime;
        rule.max = minTime;
      }
      rule.operator = operator;
    } else if (operator && operator !== "between") {
      if (minTime === undefined) return null;
      rule.min = minTime;
      rule.operator = operator;
    } else if (minTime !== undefined && maxTime !== undefined) {
      if ((minSeconds as number) <= (maxSeconds as number)) {
        rule.min = minTime;
        rule.max = maxTime;
      } else {
        rule.min = maxTime;
        rule.max = minTime;
      }
    } else if (minTime !== undefined) {
      rule.min = minTime;
    } else if (maxTime !== undefined) {
      rule.max = maxTime;
    }
    return rule;
  }

  if (r.kind === "dependent") {
    const src = r.source;
    if (!src || typeof src !== "object") return null;
    const sRow = Math.floor(Number((src as Record<string, unknown>).row));
    const sCol = Math.floor(Number((src as Record<string, unknown>).col));
    if (
      !Number.isFinite(sRow) ||
      !Number.isFinite(sCol) ||
      sRow < 0 ||
      sCol < 0
    ) {
      return null;
    }
    // Clamp the controlling cell to the LAST addressable cell, never the bare
    // constant (one past the grid edge) — same ceiling as normalizeRange above.
    const source = {
      row: Math.min(sRow, MAX_GRID_ROWS - 1),
      col: Math.min(sCol, MAX_GRID_COLUMNS - 1),
    };
    // Normalize the options map: trim keys, dedupe/trim option lists, cap counts.
    const rawOptions =
      r.options && typeof r.options === "object" && !Array.isArray(r.options)
        ? (r.options as Record<string, unknown>)
        : {};
    const options: Record<string, string[]> = {};
    let keyCount = 0;
    for (const [rawKey, rawList] of Object.entries(rawOptions)) {
      const key = rawKey.trim();
      if (key === "") continue;
      const list = normalizeOptionList(rawList);
      if (list.length === 0) continue;
      if (!(key in options)) keyCount += 1;
      options[key] = list;
      if (keyCount >= MAX_DEPENDENT_KEYS) break;
    }
    const optionsDefault = normalizeOptionList(r.optionsDefault);
    // A dependent rule with no options at all (and no default) can never allow
    // any value — that is a malformed rule, so drop it.
    if (Object.keys(options).length === 0 && optionsDefault.length === 0) {
      return null;
    }
    const rule: DependentRule = {
      kind: "dependent",
      range,
      source,
      options,
      ...messages,
    };
    if (r.sourceMode === "perRow") rule.sourceMode = "perRow";
    if (optionsDefault.length > 0) rule.optionsDefault = optionsDefault;
    return rule;
  }

  return null;
}

/** Trim / dedupe / cap an option-value list (shared by list + dependent rules). */
function normalizeOptionList(raw: unknown): string[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of raw) {
    if (typeof v !== "string") continue;
    const trimmed = v.trim();
    if (trimmed === "" || seen.has(trimmed)) continue;
    seen.add(trimmed);
    out.push(trimmed);
    if (out.length >= MAX_LIST_VALUES) break;
  }
  return out;
}

/**
 * Normalize a raw validations array off a persisted sheet: drop malformed
 * entries and cap the list length (DoS bound). Returns a fresh array; array
 * order is preserved so first-match-wins lookup is deterministic.
 */
export function normalizeRules(raw: unknown): ValidationRule[] {
  if (!Array.isArray(raw)) return [];
  const out: ValidationRule[] = [];
  for (const item of raw) {
    const rule = normalizeRule(item);
    if (!rule) continue;
    out.push(rule);
    if (out.length >= MAX_VALIDATION_RULES) break;
  }
  return out;
}

/**
 * `next` (rules worked on in their tidied form, `normalizeRules`) with each
 * rule a command did not change given back as the sheet stored it: the tidying
 * is for reading and comparing, and a stored rule it would rewrite (a list's
 * blank or spaced items, "○, 　," as Excel keeps it) stays word for word
 * (xlsx-rounds R224-X04 E25: adding a rule elsewhere stored that list as "○").
 */
export function withUntouchedRulesAsStored(stored: unknown, next: readonly ValidationRule[]): unknown[] {
  if (!Array.isArray(stored) || stored.length === 0) return [...next];
  const byTidied = new Map<string, unknown[]>();
  for (const raw of stored) {
    const rule = normalizeRule(raw);
    if (!rule) continue;
    const key = JSON.stringify(rule);
    const list = byTidied.get(key);
    if (list) list.push(raw);
    else byTidied.set(key, [raw]);
  }
  return next.map((rule) => {
    const list = byTidied.get(JSON.stringify(rule));
    return list && list.length > 0 ? list.shift() : rule;
  });
}

/**
 * The first rule (by array order) covering cell (row, col), or null when no
 * rule applies. The renderer uses this to decide whether to show a dropdown;
 * the validator uses it to pick which rule to enforce on commit.
 */
export function ruleForCell(
  rules: readonly ValidationRule[],
  row: number,
  col: number,
): ValidationRule | null {
  if (rules.length <= RULE_INDEX_MIN) {
    for (const rule of rules) {
      if (rangeContains(rule.range, row, col)) return rule;
    }
    return null;
  }
  // Asked for each cell drawn: up to MAX_VALIDATION_RULES rules are found by
  // the band of rows they reach, in array order (first match wins).
  let index = RULE_INDEXES.get(rules);
  if (!index) {
    index = ruleIndex(rules);
    RULE_INDEXES.set(rules, index);
  }
  const band = index.bands.get(Math.floor(row / RULE_BAND_ROWS));
  const tall = index.tall;
  let b = 0;
  let t = 0;
  while ((band && b < band.length) || t < tall.length) {
    const next =
      band && b < band.length && (t >= tall.length || band[b] < tall[t])
        ? band[b++]
        : tall[t++];
    if (rangeContains(rules[next].range, row, col)) return rules[next];
  }
  return null;
}

/** Below this many rules a cell's rule is found by walking them. */
const RULE_INDEX_MIN = 64;
/** Rows of one band of the rule index. */
const RULE_BAND_ROWS = 256;
/** A rule reaching more bands than this is asked of every row. */
const RULE_MAX_BANDS = 64;

interface RuleIndex {
  /** The rules reaching each band of rows, by their order. */
  bands: Map<number, number[]>;
  /** The rules too tall to list by band, by their order. */
  tall: number[];
}

const RULE_INDEXES = new WeakMap<readonly ValidationRule[], RuleIndex>();

function ruleIndex(rules: readonly ValidationRule[]): RuleIndex {
  const bands = new Map<number, number[]>();
  const tall: number[] = [];
  rules.forEach((rule, position) => {
    const first = Math.floor(rule.range.r1 / RULE_BAND_ROWS);
    const last = Math.floor(rule.range.r2 / RULE_BAND_ROWS);
    if (last - first + 1 > RULE_MAX_BANDS) {
      tall.push(position);
      return;
    }
    for (let band = first; band <= last; band += 1) {
      const listed = bands.get(band);
      if (listed) listed.push(position);
      else bands.set(band, [position]);
    }
  });
  return { bands, tall };
}

/**
 * Validate a value against a single rule. Returns { ok: true } when the value
 * is allowed (including the always-allowed blank), otherwise { ok: false,
 * message } where `message` is an i18n key the UI translates.
 *
 * Type coercion: number rules parse the trimmed string with Number(); a value
 * that is not a finite number is rejected. Text rules measure the raw string's
 * `.length` (without trimming, so trailing spaces count, matching Excel's
 * character-count semantics). List rules compare the trimmed entry with each
 * option without regard to letter case — Excel's list validation accepts
 * "yes" for a Yes/No list (a case-sensitive list needs a custom EXACT rule).
 *
 * The optional `ctx` carries the cell coordinates and the host callbacks that
 * the formula and dependent kinds need (see ValidationContext). It is ignored by
 * list / number / text / date rules, so the existing two-argument call is
 * unchanged for them. When a formula / dependent rule lacks the context it
 * needs, it fails closed with an explicit resolver-unavailable reason.
 */
export function validateValue(
  rule: ValidationRule,
  value: unknown,
  ctx?: ValidationContext,
): ValidationResult {
  if (rule.kind === "any") return { ok: true };
  // Excel defaults "Ignore blank" ON; an explicit false makes blank an invalid
  // value and lets the host's normal error-alert policy decide block/override.
  if (isBlank(value)) {
    return rule.ignoreBlank === false
      ? { ok: false, message: "spreadsheet.validationBlankRejected" }
      : { ok: true };
  }
  const text = validationCandidateText(value);

  if (rule.kind === "list") {
    // A declared source range owns the allowed set. Missing read authority or
    // an unresolved cross-sheet identity is not permission to accept a value.
    const at =
      ctx?.row !== undefined && ctx.col !== undefined
        ? { row: ctx.row, col: ctx.col }
        : undefined;
    const formulaSource = listSourceFormula(rule);
    if (formulaSource) {
      // Its list is what the calculation answered for this cell; none yet (or
      // no evaluator) is no authority to accept, as an unreadable range.
      const cell = at ?? { row: rule.range.r1, col: rule.range.c1 };
      const answer = ctx?.evalFormula?.(
        listSourceOptionsFormula(formulaSource), cell.row, cell.col, rule.range.r1, rule.range.c1,
      );
      if (answer === undefined || answer === null) {
        return {
          ok: false,
          message: "spreadsheet.validationResolverUnavailable",
          failure: "resolver-unavailable",
        };
      }
      const key = text.toLowerCase();
      return listFormulaAnswerOptions(answer, MAX_LIST_MEMBERSHIP_SCAN).some(
        (v) => v === text || v.toLowerCase() === key,
      )
        ? { ok: true }
        : { ok: false, message: "spreadsheet.validationListRejected" };
    }
    if (rule.sourceRange) {
      const source = listSourceRangeForCell(rule, at?.row, at?.col);
      const sourceResolvable =
        source !== null &&
        ctx?.cellValueAt !== undefined &&
        (!source.sheet || ctx.sourceRangeResolvable?.(source) === true);
      if (!sourceResolvable) {
        return {
          ok: false,
          message: "spreadsheet.validationResolverUnavailable",
          failure: "resolver-unavailable",
        };
      }
    }
    // The entry is compared as typed: Excel's list validation ignores letter
    // case, never spaces, so " Yes" and "Yes " are not the listed "Yes".
    const key = text.toLowerCase();
    const allowed = listRuleOptions(rule, ctx?.cellValueAt, at, MAX_LIST_MEMBERSHIP_SCAN).some(
      (v) => v === text || v.toLowerCase() === key,
    );
    return allowed
      ? { ok: true }
      : { ok: false, message: "spreadsheet.validationListRejected" };
  }

  if (rule.kind === "number") {
    // A formula bound is read for THIS cell; one that does not yield a number
    // makes the rule unsatisfiable rather than one-sided (live audit B149).
    let min = rule.min;
    let max = rule.max;
    if (rule.minFormula !== undefined || rule.maxFormula !== undefined) {
      const resolved = resolveValidationBounds(rule, ctx);
      if (resolved === null) {
        return { ok: false, message: "spreadsheet.validationConditionFailed" };
      }
      min = resolved.min;
      max = resolved.max;
    }
    const trimmed = text.trim();
    // Excel's number syntax only: a JavaScript 0x10 / 0b10 / Infinity literal
    // is text in the cell and must not pass a whole-number rule as 16 or 2
    // (live audit B108).
    const num = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(trimmed)
      ? Number(trimmed)
      : NaN;
    if (trimmed === "" || !Number.isFinite(num)) {
      return { ok: false, message: "spreadsheet.validationNotNumber" };
    }
    if (rule.integer && !Number.isInteger(num)) {
      return { ok: false, message: "spreadsheet.validationNotInteger" };
    }
    if (rule.operator) {
      return validationValueMatchesOperator(num, rule.operator, min, max)
        ? { ok: true }
        : { ok: false, message: "spreadsheet.validationConditionFailed" };
    }
    if (min !== undefined && num < min) {
      return { ok: false, message: "spreadsheet.validationBelowMin" };
    }
    if (max !== undefined && num > max) {
      return { ok: false, message: "spreadsheet.validationAboveMax" };
    }
    return { ok: true };
  }

  if (rule.kind === "text") {
    const len = text.length;
    if (rule.operator) {
      return validationValueMatchesOperator(
        len,
        rule.operator,
        rule.minLength,
        rule.maxLength,
      )
        ? { ok: true }
        : { ok: false, message: "spreadsheet.validationConditionFailed" };
    }
    if (rule.minLength !== undefined && len < rule.minLength) {
      return { ok: false, message: "spreadsheet.validationTooShort" };
    }
    if (rule.maxLength !== undefined && len > rule.maxLength) {
      return { ok: false, message: "spreadsheet.validationTooLong" };
    }
    return { ok: true };
  }

  if (rule.kind === "date") {
    const ms = dateValidationMs(text, ctx?.date1904);
    if (ms === null) {
      return { ok: false, message: "spreadsheet.validationNotDate" };
    }
    if (rule.operator) {
      const minMs = parseIsoDate(rule.min) ?? undefined;
      const maxMs = parseIsoDate(rule.max) ?? undefined;
      if (!validationValueMatchesOperator(ms, rule.operator, minMs, maxMs)) {
        return { ok: false, message: "spreadsheet.validationConditionFailed" };
      }
    } else if (rule.min !== undefined) {
      const minMs = parseIsoDate(rule.min);
      if (minMs !== null && ms < minMs) {
        return { ok: false, message: "spreadsheet.validationDateBeforeMin" };
      }
    }
    if (!rule.operator && rule.max !== undefined) {
      const maxMs = parseIsoDate(rule.max);
      if (maxMs !== null && ms > maxMs) {
        return { ok: false, message: "spreadsheet.validationDateAfterMax" };
      }
    }
    if (rule.weekdays && rule.weekdays.length > 0) {
      const dow = new Date(ms).getUTCDay();
      if (!rule.weekdays.includes(dow as Weekday)) {
        return { ok: false, message: "spreadsheet.validationDateWeekday" };
      }
    }
    if (rule.notBefore === "today" || rule.notAfter === "today") {
      const today = todayUtcMs(ctx?.now);
      if (rule.notBefore === "today" && ms < today) {
        return { ok: false, message: "spreadsheet.validationDateBeforeToday" };
      }
      if (rule.notAfter === "today" && ms > today) {
        return { ok: false, message: "spreadsheet.validationDateAfterToday" };
      }
    }
    return { ok: true };
  }

  if (rule.kind === "time") {
    const seconds = timeValidationSeconds(text);
    if (seconds === null) {
      return { ok: false, message: "spreadsheet.validationNotTime" };
    }
    const minSeconds = parseIsoTime(rule.min) ?? undefined;
    const maxSeconds = parseIsoTime(rule.max) ?? undefined;
    const operator = rule.operator ?? "between";
    return validationValueMatchesOperator(
      seconds,
      operator,
      minSeconds,
      maxSeconds,
    )
      ? { ok: true }
      : { ok: false, message: "spreadsheet.validationConditionFailed" };
  }

  if (rule.kind === "formula") {
    const evalFormula = ctx?.evalFormula;
    if (!evalFormula) {
      return {
        ok: false,
        message: "spreadsheet.validationResolverUnavailable",
        failure: "resolver-unavailable",
      };
    }
    const row = ctx?.row ?? rule.range.r1;
    const col = ctx?.col ?? rule.range.c1;
    const result = evalFormula(
      rule.formula,
      row,
      col,
      rule.range.r1,
      rule.range.c1,
    );
    return isFormulaResultTruthy(result)
      ? { ok: true }
      : { ok: false, message: "spreadsheet.validationFormulaFailed" };
  }

  // dependent
  const cellValueAt = ctx?.cellValueAt;
  if (!cellValueAt) {
    return {
      ok: false,
      message: "spreadsheet.validationResolverUnavailable",
      failure: "resolver-unavailable",
    };
  }
  const row = ctx?.row ?? rule.range.r1;
  const col = ctx?.col ?? rule.range.c1;
  const options = dependentOptionsForCell(rule, row, col, cellValueAt);
  const allowed = options.some((v) => v === text.trim());
  return allowed
    ? { ok: true }
    : { ok: false, message: "spreadsheet.validationDependentRejected" };
}

/** A cell coordinate flagged as violating its validation rule. */
export interface InvalidCellCoord {
  row: number;
  col: number;
}

/** Max invalid-cell circles drawn at once (Excel caps at 255; DoS bound). */
export const MAX_INVALID_CIRCLES = 255;

/**
 * Scan every validation rule's range for cells whose CURRENT value violates the
 * rule (Excel's "Circle Invalid Data"). Returns the violating cell coordinates
 * in row-major order, de-duplicated across overlapping rules, capped at
 * MAX_INVALID_CIRCLES (Excel circles at most 255). Blank cells are flagged only
 * when their rule explicitly turns "Ignore blank" off, matching `validateValue`.
 *
 * Pure: the caller supplies `readCellValue(row, col)` (the cell's own display
 * value) and an optional `contextAt(row, col)` that wires the formula / dependent
 * kinds' callbacks (the same context `validateValue` consumes). When a rule
 * needs a resolver the caller did not supply, it is reported invalid rather
 * than silently disappearing from the audit.
 *
 * @param rules the sheet's normalized validation rules
 * @param readCellValue reads the cell's own current value (display text)
 * @param contextAt builds the per-cell ValidationContext (optional)
 */
export function scanInvalidCells(
  rules: readonly ValidationRule[],
  readCellValue: (row: number, col: number) => unknown,
  contextAt?: (row: number, col: number) => ValidationContext,
): InvalidCellCoord[] {
  const seen = new Set<string>();
  const out: InvalidCellCoord[] = [];
  for (const rule of rules) {
    const { r1, c1, r2, c2 } = rule.range;
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) {
        const key = `${row}:${col}`;
        if (seen.has(key)) continue; // already flagged by an earlier rule
        const value = readCellValue(row, col);
        const result = validateValue(rule, value, contextAt?.(row, col));
        if (result.ok) continue;
        seen.add(key);
        out.push({ row, col });
        if (out.length >= MAX_INVALID_CIRCLES) return out;
      }
    }
  }
  return out;
}

/**
 * Build a normalized rule for a selection rectangle from raw dialog params.
 * Used by the toolbar/dialog to turn the user's choices into a persistable
 * rule. Returns null when the params do not yield a valid rule (e.g. a list
 * with no values), so the caller surfaces a validation error. The optional
 * `messages` carries the Excel-style input/error message fields (additive — a
 * call that omits them yields a rule with no message fields, as before).
 */
export function buildRule(
  range: ValidationRange,
  params:
    | { kind: "any" }
    | {
        kind: "list";
        values: string[];
        sourceRange?: string;
        showDropDown?: boolean;
      }
    | {
        kind: "number";
        /** A number, or a formula / reference text ("=$H$15") kept as a
         *  formula bound (see NumberRule.minFormula). */
        min?: number | string;
        max?: number | string;
        minFormula?: string;
        maxFormula?: string;
        integer?: boolean;
        operator?: ValidationOperator;
      }
    | {
        kind: "text";
        minLength?: number;
        maxLength?: number;
        operator?: ValidationOperator;
      }
    | { kind: "formula"; formula: string }
    | {
        kind: "date";
        min?: string;
        max?: string;
        weekdays?: number[];
        notBefore?: "today";
        notAfter?: "today";
        operator?: ValidationOperator;
      }
    | {
        kind: "time";
        min?: string;
        max?: string;
        operator?: ValidationOperator;
      }
    | {
        kind: "dependent";
        source: { row: number; col: number };
        sourceMode?: "fixed" | "perRow";
        options: Record<string, string[]>;
        optionsDefault?: string[];
      },
  messages?: ValidationMessages,
): ValidationRule | null {
  return normalizeRule({ ...params, ...(messages ?? {}), range });
}

/**
 * Take the given range out of every rule covering part of it: the Data
 * Validation dialog's Clear All on a selection, a new rule replacing the old
 * one on the selected cells, and a paste of validation onto them. Excel
 * changes the selected cells only, so a rule on A1:A3 keeps A1 and A3 when A2
 * is cleared or given another rule. The cut is {@link clearRulesForClearAll}'s:
 * every remaining piece is rebased from the rule's own top-left, so a relative
 * custom formula or list source still reads the same cells from each cell
 * that keeps the rule. `null` when a remainder cannot be represented; the
 * caller must not commit.
 */
export function clearRulesInRange(
  rules: readonly ValidationRule[],
  range: ValidationRange,
  shift: ValidationFormulaShifter,
): ValidationRule[] | null {
  return clearRulesForClearAll(rules, range, shift)?.rules ?? null;
}

/** True when two ranges share at least one cell. */
export function rangesOverlap(a: ValidationRange, b: ValidationRange): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

/** True when any rule's range intersects the given range. */
export function rangeHasRule(
  rules: readonly ValidationRule[],
  range: ValidationRange,
): boolean {
  return rules.some((rule) => rangesOverlap(rule.range, range));
}

/** Formula/reference shifter supplied by the host to keep this module engine-free. */
export type ValidationFormulaShifter = (
  formula: string,
  rowDelta: number,
  colDelta: number,
) => string;

function exactRange(raw: ValidationRange): ValidationRange | null {
  const normalized = normalizeRange(raw);
  if (
    !normalized ||
    normalized.r1 !== raw.r1 ||
    normalized.c1 !== raw.c1 ||
    normalized.r2 !== raw.r2 ||
    normalized.c2 !== raw.c2
  ) {
    return null;
  }
  return normalized;
}

function shiftListSourceRange(
  sourceRange: string,
  rowDelta: number,
  colDelta: number,
  shift: ValidationFormulaShifter,
): string | null {
  const shifted = shift(
    sourceRange.startsWith("=") ? sourceRange : `=${sourceRange}`,
    rowDelta,
    colDelta,
  );
  if (!shifted.startsWith("=")) return null;
  const result = shifted.slice(1);
  return parseSourceRange(result) || listSourceFormula({ sourceRange: result }) ? result : null;
}

function translatedRule(
  rule: ValidationRule,
  range: ValidationRange,
  rowDelta: number,
  colDelta: number,
  shift: ValidationFormulaShifter,
  translateDependentSource: boolean,
): ValidationRule | null {
  const candidate = { ...rule, range } as Record<string, unknown>;
  if (rule.kind === "formula") {
    candidate.formula = shift(rule.formula, rowDelta, colDelta);
  } else if (rule.kind === "number") {
    // A bound written as a formula moves with the rule as the custom formula
    // does: =H45 copied five rows down reads H50, $H$45 stays (Excel). The
    // bounds are stored without their "=", so the shifter sees them with one.
    for (const key of ["minFormula", "maxFormula"] as const) {
      const bound = rule[key];
      if (bound === undefined) continue;
      const shifted = shift(`=${bound}`, rowDelta, colDelta);
      if (!shifted.startsWith("=")) return null;
      candidate[key] = shifted.slice(1);
    }
  } else if (rule.kind === "list" && rule.sourceRange) {
    const sourceRange = shiftListSourceRange(
      rule.sourceRange,
      rowDelta,
      colDelta,
      shift,
    );
    if (!sourceRange) return null;
    candidate.sourceRange = sourceRange;
  } else if (rule.kind === "dependent" && translateDependentSource) {
    const row = rule.source.row + rowDelta;
    const col = rule.source.col + colDelta;
    if (
      !Number.isSafeInteger(row) ||
      !Number.isSafeInteger(col) ||
      row < 0 ||
      row >= MAX_GRID_ROWS ||
      col < 0 ||
      col >= MAX_GRID_COLUMNS
    ) {
      return null;
    }
    candidate.source = { row, col };
  }
  const normalized = normalizeRule(candidate);
  if (
    !normalized ||
    normalized.range.r1 !== range.r1 ||
    normalized.range.c1 !== range.c1 ||
    normalized.range.r2 !== range.r2 ||
    normalized.range.c2 !== range.c2 ||
    (rule.kind === "list" &&
      rule.sourceRange !== undefined &&
      (normalized.kind !== "list" || normalized.sourceRange === undefined))
  ) {
    return null;
  }
  return normalized;
}

/**
 * Plan the data-validation half of Home > Clear > Clear All. Excel removes
 * validation from the cleared cells and from nothing else: a rule wholly
 * inside the selection goes, and a rule that straddles its edge keeps the
 * cells outside it. That remainder is cut into at most four rectangles (the
 * rows above the cleared block, the columns either side of it, the rows
 * below), each rebased from the rule's own top-left the way a clipped Copy is,
 * so a relative custom formula or list source still reads the same cells from
 * every cell that keeps the rule. Rules clear of the selection are returned
 * as the same objects in the same order.
 *
 * `null` means a remainder cannot be represented (a rebased reference falls
 * off the grid, or the split would pass the rule cap); the caller must not
 * commit half of the Clear.
 */
/**
 * The rule as it reads from another base: its relative formulas, bounds and
 * list source shifted from its own top-left to `range`'s (the dialog loads
 * a rule for the selected cell this way, and writes it back for every cell
 * of a same-settings group from that cell's own base; B513). Null when a
 * shifted reference cannot be represented.
 */
export function rebaseValidationRule(
  rule: ValidationRule,
  range: ValidationRange,
  shift: ValidationFormulaShifter,
): ValidationRule | null {
  const normalized = exactRange(range);
  if (!normalized) return null;
  return translatedRule(
    rule,
    normalized,
    normalized.r1 - rule.range.r1,
    normalized.c1 - rule.range.c1,
    shift,
    false,
  );
}

/**
 * A reference-bearing setting (a list Source, a custom formula, a formula
 * bound) as "the same settings" compares it: relative parts measured from the
 * rule's top-left cell (R1C1), sheet names without redundant quotes and, like
 * every name in a formula, without regard to case. String literals stay as
 * written. Measured in Excel: =A1:A2 on B1 and =A5:A6 on B5 are one
 * validation (=A1:A2 on B9 is not), and so are ='Lists'!$A$1:$A$3,
 * =Lists!$A$1:$A$3 and =lists!$A$1:$A$3 (live audit S-040: a copy, whose
 * Source the reference rewriter spells without the quotes, was left out of
 * "Apply these changes to all other cells with the same settings").
 */
function referenceSettingSignature(
  text: string,
  anchor: { row: number; col: number },
): string {
  const body = formulaA1ToR1C1(text.trim().replace(/^=/, "").trim(), anchor);
  let out = "";
  let i = 0;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '"' || ch === "'") {
      let end = i + 1;
      while (end < body.length) {
        if (body[end] === ch) {
          if (body[end + 1] === ch) {
            end += 2;
            continue;
          }
          break;
        }
        end += 1;
      }
      const quoted = body.slice(i, end + 1);
      if (ch === "'" && body[end + 1] === "!") {
        const name = quoted.slice(1, -1).replace(/''/g, "'");
        out += formatSheetNameForFormula(name).toUpperCase();
      } else {
        out += quoted;
      }
      i = end + 1;
      continue;
    }
    out += ch.toUpperCase();
    i += 1;
  }
  return out;
}

function sortedJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(sortedJson).join(",")}]`;
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    return `{${Object.keys(record)
      .filter((key) => record[key] !== undefined)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${sortedJson(record[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
}

/**
 * What "the same settings" compares for a rule: every setting but the cells
 * it covers, with defaults spelled one way (normalizeRule) and references in
 * the form {@link referenceSettingSignature} gives them. A dependent list's
 * controlling cell matches either where it is (a rule split by Clear keeps
 * it) or where it sits from the rule's top-left (a copy moves it with the
 * rule), so it yields both keys.
 */
function validationSettingsSignatures(rule: ValidationRule): string[] {
  const normalized = normalizeRule(rule);
  const settings: Record<string, unknown> = { ...(normalized ?? rule) };
  delete settings.range;
  if (!normalized) return [sortedJson(settings)];
  const anchor = { row: normalized.range.r1, col: normalized.range.c1 };
  if (normalized.kind === "list" && normalized.sourceRange) {
    // A live Source is the list; any literal values are unused.
    delete settings.values;
    settings.sourceRange = referenceSettingSignature(normalized.sourceRange, anchor);
  } else if (normalized.kind === "formula") {
    settings.formula = referenceSettingSignature(normalized.formula, anchor);
  } else if (normalized.kind === "number") {
    for (const key of ["minFormula", "maxFormula"] as const) {
      const bound = normalized[key];
      if (bound !== undefined) settings[key] = referenceSettingSignature(bound, anchor);
    }
  } else if (normalized.kind === "dependent") {
    const { row, col } = normalized.source;
    return [
      sortedJson({ ...settings, source: { row, col } }),
      sortedJson({ ...settings, source: { rowOffset: row - anchor.row, colOffset: col - anchor.col } }),
    ];
  }
  return [sortedJson(settings)];
}

/**
 * Excel's "same settings" for data validation, from `rule`'s side: the test
 * behind the dialog's "Apply these changes to all other cells with the same
 * settings" and Go To Special > Data validation > Same. Two rules match when
 * everything but their range agrees once references are read the way the
 * cells read them, so a copied, filled or split rule still counts.
 */
export function sameValidationSettingsAs(
  rule: ValidationRule,
): (candidate: ValidationRule) => boolean {
  const keys = new Set(validationSettingsSignatures(rule));
  return (candidate) =>
    candidate === rule ||
    validationSettingsSignatures(candidate).some((key) => keys.has(key));
}

export function clearRulesForClearAll(
  rules: readonly ValidationRule[],
  range: ValidationRange,
  shift: ValidationFormulaShifter,
): { rules: ValidationRule[]; changed: boolean } | null {
  const cleared = normalizeRange(range);
  if (!cleared) return null;
  const next: ValidationRule[] = [];
  let changed = false;
  for (const rule of rules) {
    if (!rangesOverlap(rule.range, cleared)) {
      next.push(rule);
      continue;
    }
    changed = true;
    const own = rule.range;
    const bandR1 = Math.max(own.r1, cleared.r1);
    const bandR2 = Math.min(own.r2, cleared.r2);
    const pieces: ValidationRange[] = [];
    if (own.r1 < cleared.r1) {
      pieces.push({ r1: own.r1, c1: own.c1, r2: cleared.r1 - 1, c2: own.c2 });
    }
    if (own.c1 < cleared.c1) {
      pieces.push({ r1: bandR1, c1: own.c1, r2: bandR2, c2: cleared.c1 - 1 });
    }
    if (own.c2 > cleared.c2) {
      pieces.push({ r1: bandR1, c1: cleared.c2 + 1, r2: bandR2, c2: own.c2 });
    }
    if (own.r2 > cleared.r2) {
      pieces.push({ r1: cleared.r2 + 1, c1: own.c1, r2: own.r2, c2: own.c2 });
    }
    for (const piece of pieces) {
      const kept = translatedRule(
        rule,
        piece,
        piece.r1 - own.r1,
        piece.c1 - own.c1,
        shift,
        false,
      );
      if (!kept) return null;
      next.push(kept);
    }
  }
  if (next.length > MAX_VALIDATION_RULES) return null;
  return { rules: changed ? next : [...rules], changed };
}

/**
 * Detach the validation rules owned by a copied rectangle. Rules that extend
 * beyond the Copy selection are clipped; formula/list references are rebased
 * from the original rule anchor to the clipped top-left. `null` means the
 * snapshot cannot be represented faithfully and must never authorize a paste.
 */
export function snapshotRulesForValidationPaste(
  rules: readonly ValidationRule[],
  sourceRange: ValidationRange,
  shift: ValidationFormulaShifter,
): ValidationRule[] | null {
  const source = exactRange(sourceRange);
  if (!source) return null;
  const snapshot: ValidationRule[] = [];
  for (const rawRule of rules) {
    const rule = normalizeRule(rawRule);
    if (!rule || !rangesOverlap(rule.range, source)) continue;
    const clipped: ValidationRange = {
      r1: Math.max(rule.range.r1, source.r1),
      c1: Math.max(rule.range.c1, source.c1),
      r2: Math.min(rule.range.r2, source.r2),
      c2: Math.min(rule.range.c2, source.c2),
    };
    const detached = translatedRule(
      rule,
      clipped,
      clipped.r1 - rule.range.r1,
      clipped.c1 - rule.range.c1,
      shift,
      false,
    );
    if (!detached) return null;
    snapshot.push(detached);
  }
  return snapshot;
}

/**
 * The runs of consecutive rows in `keptRows` (ascending), each with the row
 * it lands on once the rows left out close up under `firstRow`: how a Copy
 * of a filtered table's visible rows pastes.
 */
export function keptRowRuns(
  keptRows: readonly number[],
  firstRow: number,
): Array<{ from: number; to: number; at: number }> {
  const runs: Array<{ from: number; to: number; at: number }> = [];
  keptRows.forEach((row, index) => {
    const last = runs[runs.length - 1];
    if (last && row === last.to + 1) last.to = row;
    else runs.push({ from: row, to: row, at: firstRow + index });
  });
  return runs;
}

/**
 * The validation rules of a Copy that leaves rows out (the visible rows of a
 * filtered table), laid out as that Copy pastes: the kept rows close up under
 * the first one, and each run of them takes its rules along, references
 * rebased as for a paste. Without it the rules kept the hidden rows' places,
 * fell outside the pasted block, and the planner refused the whole Paste
 * without a word (live audit S-045). `null` when a run cannot be represented.
 */
export function snapshotRulesForKeptRowsPaste(
  rules: readonly ValidationRule[],
  sourceRange: ValidationRange,
  keptRows: readonly number[],
  shift: ValidationFormulaShifter,
): ValidationRule[] | null {
  const out: ValidationRule[] = [];
  for (const run of keptRowRuns(keptRows, sourceRange.r1)) {
    const from = { r1: run.from, c1: sourceRange.c1, r2: run.to, c2: sourceRange.c2 };
    const snapshot = snapshotRulesForValidationPaste(rules, from, shift);
    if (!snapshot) return null;
    const placed = planRulesForValidationPaste(
      [],
      snapshot,
      from,
      { ...from, r1: run.at, r2: run.at + run.to - run.from },
      shift,
    );
    if (!placed) return null;
    out.push(...placed.rules);
  }
  return out.length > MAX_VALIDATION_RULES ? null : out;
}

/**
 * Plan the validation-only half of Paste Special. The copied snapshot is
 * translated positionally, destination-intersecting rules are replaced, and
 * unrelated rules retain their order. Returns `null` on any provenance,
 * bounds, reference or rule-cap failure so the component can fail closed before
 * creating history or autosave work.
 */







export function validationPasteOverRuleLimit(
  destinationRules: readonly unknown[],
  copiedRules: readonly unknown[],
): boolean {
  return copiedRules.length > 0 && destinationRules.length + copiedRules.length > MAX_VALIDATION_RULES;
}

export function planRulesForValidationPaste(
  destinationRules: readonly ValidationRule[],
  copiedRules: readonly ValidationRule[],
  sourceRange: ValidationRange,
  destinationRange: ValidationRange,
  shift: ValidationFormulaShifter,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ValidationRule[]; changed: boolean } | null {
  const source = exactRange(sourceRange);
  const destination = exactRange(destinationRange);
  if (
    !source ||
    !destination ||
    source.r2 - source.r1 !== destination.r2 - destination.r1 ||
    source.c2 - source.c1 !== destination.c2 - destination.c1 ||
    (bounds?.maxRow !== undefined && destination.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && destination.c2 > bounds.maxCol)
  ) {
    return null;
  }

  const rowDelta = destination.r1 - source.r1;
  const colDelta = destination.c1 - source.c1;
  const translated: ValidationRule[] = [];
  for (const rawRule of copiedRules) {
    const rule = normalizeRule(rawRule);
    if (
      !rule ||
      rule.range.r1 < source.r1 ||
      rule.range.c1 < source.c1 ||
      rule.range.r2 > source.r2 ||
      rule.range.c2 > source.c2
    ) {
      return null;
    }
    const nextRange: ValidationRange = {
      r1: rule.range.r1 + rowDelta,
      c1: rule.range.c1 + colDelta,
      r2: rule.range.r2 + rowDelta,
      c2: rule.range.c2 + colDelta,
    };
    if (
      (bounds?.maxRow !== undefined && nextRange.r2 > bounds.maxRow) ||
      (bounds?.maxCol !== undefined && nextRange.c2 > bounds.maxCol)
    ) {
      return null;
    }
    const nextRule = translatedRule(
      rule,
      nextRange,
      rowDelta,
      colDelta,
      shift,
      true,
    );
    if (!nextRule) return null;
    translated.push(nextRule);
  }

  const retained = clearRulesInRange(destinationRules, destination, shift);
  if (!retained) return null;
  // At the page's bound, a copied rule that restricts nothing and says nothing
  // need not come along: a list of 1,000 areas of type="none" refused the whole
  
  const pasted =
    translated.length + retained.length > MAX_VALIDATION_RULES
      ? translated.filter((rule) => rule.kind !== "any" || resolveInputMessage(rule) !== null)
      : translated;
  if (pasted.length + retained.length > MAX_VALIDATION_RULES) return null;
  const rules = [...pasted, ...retained];
  return {
    rules,
    changed: JSON.stringify(rules) !== JSON.stringify(destinationRules),
  };
}

/**
 * Plan the data validation of a paste that repeats the copy over a whole
 * number of copies (see planRulesForTiledFormatPaste): each copied rule
 * repeats tile by tile, touching copies joined into one rectangle written
 * against its own corner, so a validated cell pasted down a column is one
 * rule, not one per cell. The selection's own rules lose its cells.
 */
export function planRulesForTiledValidationPaste(
  destinationRules: readonly ValidationRule[],
  copiedRules: readonly ValidationRule[],
  sourceRange: ValidationRange,
  selection: ValidationRange,
  shift: ValidationFormulaShifter,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ValidationRule[]; changed: boolean } | null {
  const source = exactRange(sourceRange);
  const target = exactRange(selection);
  if (!source || !target) return null;
  const rows = source.r2 - source.r1 + 1;
  const cols = source.c2 - source.c1 + 1;
  const tileRows = (target.r2 - target.r1 + 1) / rows;
  const tileCols = (target.c2 - target.c1 + 1) / cols;
  if (
    !Number.isInteger(tileRows) ||
    !Number.isInteger(tileCols) ||
    (bounds?.maxRow !== undefined && target.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && target.c2 > bounds.maxCol)
  ) {
    return null;
  }
  const carried: ValidationRule[] = [];
  for (const rawRule of copiedRules) {
    const rule = normalizeRule(rawRule);
    if (
      !rule ||
      rule.range.r1 < source.r1 ||
      rule.range.c1 < source.c1 ||
      rule.range.r2 > source.r2 ||
      rule.range.c2 > source.c2
    ) {
      return null;
    }
    const copies: ValidationRange[] = [];
    for (let tileRow = 0; tileRow < tileRows; tileRow += 1) {
      for (let tileCol = 0; tileCol < tileCols; tileCol += 1) {
        const rowDelta = target.r1 - source.r1 + tileRow * rows;
        const colDelta = target.c1 - source.c1 + tileCol * cols;
        copies.push({
          r1: rule.range.r1 + rowDelta,
          c1: rule.range.c1 + colDelta,
          r2: rule.range.r2 + rowDelta,
          c2: rule.range.c2 + colDelta,
        });
      }
    }
    for (const range of coalesceRanges(copies)) {
      const next = translatedRule(
        rule,
        range,
        range.r1 - rule.range.r1,
        range.c1 - rule.range.c1,
        shift,
        true,
      );
      if (!next) return null;
      carried.push(next);
    }
  }
  const retained = clearRulesInRange(destinationRules, target, shift);
  if (!retained) return null;
  if (carried.length + retained.length > MAX_VALIDATION_RULES) return null;
  const rules = [...carried, ...retained];
  return {
    rules,
    changed: JSON.stringify(rules) !== JSON.stringify(destinationRules),
  };
}

/**
 * Plan the data validation of a fill (Home > Fill and the fill handle). Excel
 * 16 measured: Fill Down from R164, whose cell must be a whole number > 0,
 * leaves that rule on R164:R166. The source cells' rules repeat over the
 * filled cells as copies (a relative formula read from each cell's own row),
 * joined into one rectangle where the copies meet; a cell filled from a
 * source without a rule loses its own, one rule per cell, as a paste does.
 * `null` when `filled` is not a fill of `source` or a rule cannot be carried.
 */
export function planRulesForValidationFill(
  rules: readonly ValidationRule[],
  source: ValidationRange,
  filled: ValidationRange,
  shift: ValidationFormulaShifter,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ValidationRule[]; changed: boolean } | null {
  const offsets = fillTileOffsets(source, filled);
  if (
    !offsets ||
    (bounds?.maxRow !== undefined && filled.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && filled.c2 > bounds.maxCol)
  ) {
    return null;
  }
  const out: ValidationRule[] = [];
  for (const rule of rules) {
    if (!rangesOverlap(rule.range, filled)) {
      out.push(rule);
      continue;
    }
    const remainder = clearRulesInRange([rule], filled, shift);
    if (!remainder) return null;
    const inSource = {
      r1: Math.max(rule.range.r1, source.r1),
      c1: Math.max(rule.range.c1, source.c1),
      r2: Math.min(rule.range.r2, source.r2),
      c2: Math.min(rule.range.c2, source.c2),
    };
    const copies =
      inSource.r1 <= inSource.r2 && inSource.c1 <= inSource.c2
        ? coalesceRanges(
            offsets
              .map(({ rowDelta, colDelta }) => ({
                r1: Math.max(inSource.r1 + rowDelta, filled.r1),
                c1: Math.max(inSource.c1 + colDelta, filled.c1),
                r2: Math.min(inSource.r2 + rowDelta, filled.r2),
                c2: Math.min(inSource.c2 + colDelta, filled.c2),
              }))
              .filter((range) => range.r1 <= range.r2 && range.c1 <= range.c2),
          )
        : [];
    const carried: ValidationRule[] = [];
    for (const range of copies) {
      // Written against its own top-left, as the copy at that corner is.
      const next = translatedRule(
        rule,
        range,
        range.r1 - rule.range.r1,
        range.c1 - rule.range.c1,
        shift,
        true,
      );
      if (!next) return null;
      carried.push(next);
    }
    out.push(...remainder, ...carried);
  }
  if (out.length > MAX_VALIDATION_RULES) return null;
  return {
    rules: out,
    changed: JSON.stringify(out) !== JSON.stringify(rules),
  };
}

/**
 * Plan data-validation rules for Home > Paste > Transpose.
 *
 * Rule rectangles swap their source-selection-relative axes. Formula, list
 * source-range and dependent-source references remain relative to the rule's
 * clipped top-left and therefore move by that anchor's exact mapped delta.
 */
export function planRulesForTransposedValidationPaste(
  destinationRules: readonly ValidationRule[],
  copiedRules: readonly ValidationRule[],
  sourceRange: ValidationRange,
  destinationRange: ValidationRange,
  shift: ValidationFormulaShifter,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ValidationRule[]; changed: boolean } | null {
  const source = exactRange(sourceRange);
  const destination = exactRange(destinationRange);
  if (
    !source ||
    !destination ||
    source.r2 - source.r1 !== destination.c2 - destination.c1 ||
    source.c2 - source.c1 !== destination.r2 - destination.r1 ||
    (bounds?.maxRow !== undefined && destination.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && destination.c2 > bounds.maxCol)
  ) {
    return null;
  }

  const translated: ValidationRule[] = [];
  for (const rawRule of copiedRules) {
    const rule = normalizeRule(rawRule);
    if (
      !rule ||
      rule.range.r1 < source.r1 ||
      rule.range.c1 < source.c1 ||
      rule.range.r2 > source.r2 ||
      rule.range.c2 > source.c2
    ) {
      return null;
    }
    const nextRange: ValidationRange = {
      r1: destination.r1 + rule.range.c1 - source.c1,
      c1: destination.c1 + rule.range.r1 - source.r1,
      r2: destination.r1 + rule.range.c2 - source.c1,
      c2: destination.c1 + rule.range.r2 - source.r1,
    };
    if (
      (bounds?.maxRow !== undefined && nextRange.r2 > bounds.maxRow) ||
      (bounds?.maxCol !== undefined && nextRange.c2 > bounds.maxCol)
    ) {
      return null;
    }
    const nextRule = translatedRule(
      rule,
      nextRange,
      nextRange.r1 - rule.range.r1,
      nextRange.c1 - rule.range.c1,
      shift,
      true,
    );
    if (!nextRule) return null;
    translated.push(nextRule);
  }

  const retained = clearRulesInRange(destinationRules, destination, shift);
  if (!retained) return null;
  if (translated.length + retained.length > MAX_VALIDATION_RULES) return null;
  const rules = [...translated, ...retained];
  return {
    rules,
    changed: JSON.stringify(rules) !== JSON.stringify(destinationRules),
  };
}

/**
 * Parse a comma/newline-separated list of allowed values into a trimmed,
 * de-duplicated array (the dialog's textarea -> ListRule.values path). Empty
 * entries are dropped. Order is preserved (first occurrence wins).
 */
export function parseListValues(raw: string): string[] {
  if (typeof raw !== "string") return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const part of raw.split(/[\n,]/)) {
    const trimmed = part.trim();
    if (trimmed === "") continue;
    if (seen.has(trimmed)) continue;
    seen.add(trimmed);
    out.push(trimmed);
    if (out.length >= MAX_LIST_VALUES) break;
  }
  return out;
}

/**
 * Parse the dependent-rule options textarea into a `{ key: options[] }` map.
 * Each non-blank LINE is one mapping in the form `key: opt1, opt2, opt3` — the
 * text before the FIRST colon is the controlling value, the rest is a
 * comma-separated option list. Keys and options are trimmed; blank keys, blank
 * options and duplicate options within a key are dropped; a later line for the
 * same key replaces the earlier one. Lines without a colon, or whose option
 * list is empty, are skipped. The map is capped at MAX_DEPENDENT_KEYS keys.
 *
 * Example:
 *   "USA: California, Texas\nCanada: Ontario, Quebec"
 *   => { USA: ["California", "Texas"], Canada: ["Ontario", "Quebec"] }
 */
export function parseDependentOptions(raw: string): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  if (typeof raw !== "string") return out;
  let keyCount = 0;
  for (const line of raw.split(/\n/)) {
    const colon = line.indexOf(":");
    if (colon < 0) continue;
    const key = line.slice(0, colon).trim();
    if (key === "") continue;
    const options = parseListValues(line.slice(colon + 1));
    if (options.length === 0) continue;
    if (!(key in out)) {
      if (keyCount >= MAX_DEPENDENT_KEYS) continue;
      keyCount += 1;
    }
    out[key] = options;
  }
  return out;
}
