/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure, unit-testable predicates for AutoFilter condition filters (Excel / WPS
 * "Text Filters" / "Number Filters" parity).
 *
 * Two categories of condition:
 *
 *   TEXT:   contains / does-not-contain / begins-with / ends-with / equals /
 *           not-equals (all case-insensitive, operand[0] is the user string).
 *
 *   NUMBER: greater-than / less-than / greater-or-equal / less-or-equal /
 *           equals / not-equals / between (operands[0] = lo, operands[1] = hi)
 *           / top/bottom items or percent (operands[0] = count/percent) /
 *           above-average / below-average.
 *
 *   CUSTOM: Excel's Custom AutoFilter — two of the lines above joined by And
 *           or Or ({@link CompoundCondition}), and the four comparisons on a
 *           value that is not a number, which compare the text of text cells.
 *
 * The entry-point is {@link buildRowPredicate}: given a condition descriptor and
 * the set of all data values in the column (needed for ranked / average
 * calculations), it returns a `(displayValue: string) => boolean` that the
 * caller applies row-by-row.
 *
 * All functions are PURE (no side-effects, no state). Blank cells (display value
 * "" after trimming) are treated as non-matching for every numeric operator;
 * text operators match blanks only for "equals ''" (explicit blank match).
 */

import { excelWildcardToRegexSource } from "./formula-engine";
import { compareSpreadsheetText } from "./locale-sort";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type TextOperator =
  | "contains"
  | "does-not-contain"
  | "begins-with"
  | "ends-with"
  | "equals"
  | "not-equals"
  // The operand is a native wildcard pattern (`?` one char, `*` any run, `~`
  // escapes): the shapes the six literal operators cannot spell, `a?c`,
  // `a*c`, `*`. Matched the way Excel matches them, case-insensitively.
  | "matches"
  | "does-not-match"
  // Custom AutoFilter's "is greater than" ... "is less than or equal to" with
  // a value that is not a number: Excel orders the text of TEXT cells against
  // it, case-insensitively, as a sort would. Number and blank cells never
  // match a text comparison.
  | "greater-than"
  | "greater-or-equal"
  | "less-than"
  | "less-or-equal";

export type NumberOperator =
  | "greater-than"
  | "less-than"
  | "greater-or-equal"
  | "less-or-equal"
  | "equals"
  | "not-equals"
  | "between"
  | "top-n"
  | "top-items"
  | "bottom-items"
  | "top-percent"
  | "bottom-percent"
  | "above-average"
  | "below-average";

export type RankedNumberOperator = Extract<
  NumberOperator,
  "top-items" | "bottom-items" | "top-percent" | "bottom-percent"
>;

export interface TextCondition {
  readonly type: "text";
  readonly operator: TextOperator;
  /**
   * [0] = the user-supplied string to test against (case-insensitive).
   * For "equals" / "not-equals" an empty-string operand matches blank cells.
   */
  readonly operands: readonly [string];
}

export interface NumberCondition {
  readonly type: "number";
  readonly operator: NumberOperator;
  /**
   * Numeric bounds (parsed, not raw strings):
   *   - Most operators: operands[0] = the comparison value.
   *   - "between":      operands[0] = lo (inclusive), operands[1] = hi (inclusive).
   *   - "top-n":        legacy distinct-value count (256..500 after normalization).
   *   - "top-items" / "bottom-items": operands[0] = item count (1..255).
   *   - "top-percent" / "bottom-percent": operands[0] = percent (0..100).
   *   - "above-average" / "below-average": operands unused (computed from values).
   */
  readonly operands: readonly number[];
}

/** The six comparisons a Custom AutoFilter line can make against a number. */
export type NumberComparisonOperator = Extract<
  NumberOperator,
  | "greater-than"
  | "less-than"
  | "greater-or-equal"
  | "less-or-equal"
  | "equals"
  | "not-equals"
>;

/**
 * One line of a two-line Custom AutoFilter: any text condition, or a number
 * condition with one of the six comparisons (a ranked, average or between
 * criterion is a whole filter of its own and never one line of a pair).
 */
export type CustomCriterionCondition =
  | TextCondition
  | (NumberCondition & { readonly operator: NumberComparisonOperator });

/**
 * Excel's two-line Custom AutoFilter (`<customFilters and="1|0">` holding two
 * `<customFilter>`): a row is shown when both lines hold (And) or either does
 * (Or). A `>= lo And <= hi` pair with lo <= hi stays the one `between`
 * condition it always was.
 */
export interface CompoundCondition {
  readonly type: "compound";
  readonly join: "and" | "or";
  readonly conditions: readonly [
    CustomCriterionCondition,
    CustomCriterionCondition,
  ];
}

/** A persisted condition filter for one column. */
export type ConditionFilter = TextCondition | NumberCondition | CompoundCondition;

export interface NativeCustomFilterCriterion {
  readonly operator: string;
  readonly value: string;
}

export type NativeAutoFilterCriterion =
  | {
      readonly kind: "custom";
      readonly and: boolean;
      readonly filters: readonly NativeCustomFilterCriterion[];
    }
  | {
      readonly kind: "dynamic";
      readonly dynamicType: string;
    }
  | {
      readonly kind: "top10";
      readonly top: boolean;
      readonly percent: boolean;
      readonly value: number;
    };

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const TEXT_OPERATORS = new Set<TextOperator>([
  "contains",
  "does-not-contain",
  "begins-with",
  "ends-with",
  "equals",
  "not-equals",
  "matches",
  "does-not-match",
  "greater-than",
  "greater-or-equal",
  "less-than",
  "less-or-equal",
]);

/** Text operators whose operand may not be empty. */
const NON_EMPTY_TEXT_OPERATORS = new Set<TextOperator>([
  "contains",
  "does-not-contain",
  "begins-with",
  "ends-with",
  "matches",
  "does-not-match",
  "greater-than",
  "greater-or-equal",
  "less-than",
  "less-or-equal",
]);

const NUMBER_COMPARISON_OPERATORS = new Set<NumberOperator>([
  "greater-than",
  "less-than",
  "greater-or-equal",
  "less-or-equal",
  "equals",
  "not-equals",
]);

const NUMBER_OPERATORS = new Set<NumberOperator>([
  "greater-than",
  "less-than",
  "greater-or-equal",
  "less-or-equal",
  "equals",
  "not-equals",
  "between",
  "top-n",
  "top-items",
  "bottom-items",
  "top-percent",
  "bottom-percent",
  "above-average",
  "below-average",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function finiteNumber(value: unknown): number | null {
  if (typeof value === "string" && value.trim() === "") return null;
  const number = typeof value === "number" ? value : Number(value);
  return Number.isFinite(number) ? number : null;
}

/**
 * Normalize persisted/current condition state into the one canonical model.
 * The small `op`/`value` branch is a bounded migration for editor-state parts
 * emitted before the typed descriptor existed; unknown legacy operators are
 * discarded rather than interpreted.
 */
export function normalizeConditionFilter(value: unknown): ConditionFilter | null {
  if (!isRecord(value)) return null;

  if (value.type === "compound") {
    if (
      (value.join !== "and" && value.join !== "or") ||
      !Array.isArray(value.conditions) ||
      value.conditions.length !== 2
    ) {
      return null;
    }
    const first = normalizeCustomCriterion(value.conditions[0]);
    const second = normalizeCustomCriterion(value.conditions[1]);
    return first && second
      ? { type: "compound", join: value.join, conditions: [first, second] }
      : null;
  }

  if (value.type === "text") {
    if (!TEXT_OPERATORS.has(value.operator as TextOperator)) return null;
    if (
      !Array.isArray(value.operands) ||
      value.operands.length !== 1 ||
      typeof value.operands[0] !== "string" ||
      value.operands[0].length > 1_000 ||
      (value.operands[0] === "" &&
        NON_EMPTY_TEXT_OPERATORS.has(value.operator as TextOperator))
    ) {
      return null;
    }
    return {
      type: "text",
      operator: value.operator as TextOperator,
      operands: [value.operands[0]],
    };
  }

  if (value.type === "number") {
    const rawOperator = value.operator;
    const operator = rawOperator as NumberOperator;
    if (!NUMBER_OPERATORS.has(operator)) return null;
    if (!Array.isArray(value.operands)) return null;
    const operands = value.operands.map(finiteNumber);
    if (operands.some((operand) => operand === null)) return null;
    const numbers = operands as number[];
    if (operator === "above-average" || operator === "below-average") {
      return numbers.length === 0
        ? { type: "number", operator, operands: [] }
        : null;
    }
    if (operator === "between") {
      return numbers.length === 2
        ? { type: "number", operator, operands: [numbers[0], numbers[1]] }
        : null;
    }
    if (operator === "top-n") {
      const count = numbers[0];
      if (
        numbers.length !== 1 ||
        !Number.isInteger(count) ||
        count < 1 ||
        count > 500
      ) {
        return null;
      }
      // Pages written before the native four-mode family used top-n up to
      // 500. Values representable by native Top10 migrate forward; 256..500
      // retain the legacy operator/predicate and are never authored anew.
      return {
        type: "number",
        operator: count <= 255 ? "top-items" : "top-n",
        operands: [count],
      };
    }
    if (operator === "top-items" || operator === "bottom-items") {
      const count = numbers[0];
      return numbers.length === 1 &&
        Number.isInteger(count) &&
        count >= 1 &&
        count <= 255
        ? { type: "number", operator, operands: [count] }
        : null;
    }
    if (operator === "top-percent" || operator === "bottom-percent") {
      const percent = numbers[0];
      return numbers.length === 1 && percent >= 0 && percent <= 100
        ? { type: "number", operator, operands: [percent] }
        : null;
    }
    return numbers.length === 1
      ? { type: "number", operator, operands: [numbers[0]] }
      : null;
  }

  const legacyOperator = typeof value.op === "string" ? value.op.trim() : "";
  const legacyNumberOperator: Record<string, NumberOperator> = {
    gt: "greater-than",
    greaterThan: "greater-than",
    gte: "greater-or-equal",
    ge: "greater-or-equal",
    greaterThanOrEqual: "greater-or-equal",
    lt: "less-than",
    lessThan: "less-than",
    lte: "less-or-equal",
    le: "less-or-equal",
    lessThanOrEqual: "less-or-equal",
    eq: "equals",
    neq: "not-equals",
  };
  const migratedNumberOperator = legacyNumberOperator[legacyOperator];
  const migratedNumber = finiteNumber(value.value);
  if (migratedNumberOperator && migratedNumber !== null) {
    return {
      type: "number",
      operator: migratedNumberOperator,
      operands: [migratedNumber],
    };
  }
  const legacyTextOperator: Record<string, TextOperator> = {
    contains: "contains",
    notContains: "does-not-contain",
    beginsWith: "begins-with",
    endsWith: "ends-with",
    equals: "equals",
    notEquals: "not-equals",
  };
  const migratedTextOperator = legacyTextOperator[legacyOperator];
  if (
    migratedTextOperator &&
    typeof value.value === "string" &&
    value.value.length <= 1_000
  ) {
    return {
      type: "text",
      operator: migratedTextOperator,
      operands: [value.value],
    };
  }
  return null;
}

/** Whether a single condition can be one line of a two-line Custom AutoFilter. */
export function isCustomCriterionCondition(
  condition: ConditionFilter,
): condition is CustomCriterionCondition {
  return (
    condition.type === "text" ||
    (condition.type === "number" &&
      NUMBER_COMPARISON_OPERATORS.has(condition.operator))
  );
}

function normalizeCustomCriterion(
  value: unknown,
): CustomCriterionCondition | null {
  // A line of a pair is always the typed descriptor; the legacy `op`/`value`
  // spelling predates pairs and is never one of their lines.
  if (!isRecord(value) || (value.type !== "text" && value.type !== "number")) {
    return null;
  }
  const condition = normalizeConditionFilter(value);
  return condition && isCustomCriterionCondition(condition) ? condition : null;
}

function decodeNativeWildcardPattern(
  raw: string,
): { value: string; leading: boolean; trailing: boolean } | null {
  const tokens: Array<{ value: string; wildcard: boolean }> = [];
  for (let index = 0; index < raw.length; index += 1) {
    const current = raw[index];
    if (current === "~") {
      const escaped = raw[index + 1];
      if (escaped !== "~" && escaped !== "*" && escaped !== "?") return null;
      tokens.push({ value: escaped, wildcard: false });
      index += 1;
      continue;
    }
    if (current === "?") return null;
    tokens.push({ value: current, wildcard: current === "*" });
  }
  const wildcardIndexes = tokens
    .map((token, index) => (token.wildcard ? index : -1))
    .filter((index) => index >= 0);
  if (
    wildcardIndexes.some(
      (index) => index !== 0 && index !== tokens.length - 1,
    )
  ) {
    return null;
  }
  const leading = tokens[0]?.wildcard === true;
  const trailing = tokens.length > 1 && tokens[tokens.length - 1]?.wildcard === true;
  const value = tokens
    .filter((token) => !token.wildcard)
    .map((token) => token.value)
    .join("");
  if ((leading || trailing) && value === "") return null;
  return { value, leading, trailing };
}

const NATIVE_NUMBER_COMPARISONS: Record<string, NumberComparisonOperator> = {
  greaterThan: "greater-than",
  lessThan: "less-than",
  greaterThanOrEqual: "greater-or-equal",
  lessThanOrEqual: "less-or-equal",
};

/**
 * One native `<customFilter operator val>` line as a canonical condition, or
 * null when the line is not one Excel writes (an unknown operator, an empty
 * comparison value, an over-long value).
 *
 *   greaterThan ... lessThanOrEqual — a number value compares numbers; any
 *     other value compares text (Excel's "is greater than" "m").
 *   equal / notEqual — a literal number is a number equality; text is matched
 *     as Excel's wildcard pattern (`?` one character, `*` any run, `~`
 *     escapes), spelled as the literal operator it amounts to where one exists
 *     (`*x*` contains, `x*` begins with, `*x` ends with) and as `matches` /
 *     `does-not-match` otherwise (`a?c`, `a*c`, `x*` under notEqual).
 */
function conditionFromNativeCustomFilter(
  filter: NativeCustomFilterCriterion,
): CustomCriterionCondition | null {
  if (filter.value.length > 1_000) return null;
  const comparison = NATIVE_NUMBER_COMPARISONS[filter.operator];
  if (comparison) {
    const value = finiteNumber(filter.value);
    if (value !== null) {
      return { type: "number", operator: comparison, operands: [value] };
    }
    if (filter.value === "") return null;
    return { type: "text", operator: comparison, operands: [filter.value] };
  }
  if (filter.operator !== "equal" && filter.operator !== "notEqual") {
    return null;
  }
  const equal = filter.operator === "equal";
  const pattern = decodeNativeWildcardPattern(filter.value);
  if (
    !pattern ||
    (!equal && pattern.leading !== pattern.trailing)
  ) {
    // `a?c`, `a*c`, `*`, and the one-sided "does not begin / end with": a
    // wildcard shape none of the literal operators can spell stays the
    // pattern it is, matched the way Excel matches it.
    if (filter.value === "") return null;
    return {
      type: "text",
      operator: equal ? "matches" : "does-not-match",
      operands: [filter.value],
    };
  }
  if (!pattern.leading && !pattern.trailing) {
    const number = finiteNumber(pattern.value);
    if (number !== null) {
      return {
        type: "number",
        operator: equal ? "equals" : "not-equals",
        operands: [number],
      };
    }
    return {
      type: "text",
      operator: equal ? "equals" : "not-equals",
      operands: [pattern.value],
    };
  }
  if (pattern.leading && pattern.trailing) {
    return {
      type: "text",
      operator: equal ? "contains" : "does-not-contain",
      operands: [pattern.value],
    };
  }
  return {
    type: "text",
    operator: pattern.trailing ? "begins-with" : "ends-with",
    operands: [pattern.value],
  };
}

/**
 * Convert the exact native AutoFilter grammar the editor can execute into the
 * canonical condition model. Every `<customFilters>` of one or two lines Excel
 * writes is kept: one line is its single condition, two lines are the pair
 * joined by the native `and` flag (a `>= lo And <= hi` pair with lo <= hi is
 * `between`). Invalid Top10 bounds, foreign dynamic types and lines no Excel
 * operator spells return null (fail closed).
 */
export function conditionFilterFromNativeAutoFilter(
  criterion: NativeAutoFilterCriterion,
): ConditionFilter | null {
  if (criterion.kind === "dynamic") {
    const type = criterion.dynamicType.trim().toLowerCase();
    if (type === "aboveaverage") {
      return { type: "number", operator: "above-average", operands: [] };
    }
    if (type === "belowaverage") {
      return { type: "number", operator: "below-average", operands: [] };
    }
    return null;
  }
  if (criterion.kind === "top10") {
    const value = criterion.value;
    const operator: RankedNumberOperator = criterion.top
      ? criterion.percent
        ? "top-percent"
        : "top-items"
      : criterion.percent
        ? "bottom-percent"
        : "bottom-items";
    const valid = criterion.percent
      ? value >= 0 && value <= 100
      : Number.isInteger(value) && value >= 1 && value <= 255;
    return valid
      ? { type: "number", operator, operands: [value] }
      : null;
  }

  const filters = criterion.filters;
  if (filters.length === 1) return conditionFromNativeCustomFilter(filters[0]);
  if (filters.length !== 2) return null;
  if (criterion.and) {
    let lower: number | null = null;
    let upper: number | null = null;
    for (const filter of filters) {
      const value = finiteNumber(filter.value);
      if (value === null) break;
      if (filter.operator === "greaterThanOrEqual" && lower === null) {
        lower = value;
      } else if (filter.operator === "lessThanOrEqual" && upper === null) {
        upper = value;
      }
    }
    // An inverted range (>= 10 And <= 5) hides every row in Excel; it stays
    // the pair it is rather than a between that would show 5..10.
    if (lower !== null && upper !== null && lower <= upper) {
      return { type: "number", operator: "between", operands: [lower, upper] };
    }
  }
  const first = conditionFromNativeCustomFilter(filters[0]);
  const second = conditionFromNativeCustomFilter(filters[1]);
  return first && second
    ? {
        type: "compound",
        join: criterion.and ? "and" : "or",
        conditions: [first, second],
      }
    : null;
}

/** `~`, `*` and `?` written as themselves in a native wildcard pattern. */
export function escapeNativeWildcardLiteral(text: string): string {
  return text.replace(/[~*?]/g, "~$&");
}

function nativeNumberText(value: number): string {
  return String(Object.is(value, -0) ? 0 : value);
}

function nativeCustomFilterFromCriterion(
  condition: CustomCriterionCondition,
): NativeCustomFilterCriterion {
  if (condition.type === "number") {
    const value = nativeNumberText(condition.operands[0] ?? 0);
    switch (condition.operator) {
      case "greater-than":
        return { operator: "greaterThan", value };
      case "greater-or-equal":
        return { operator: "greaterThanOrEqual", value };
      case "less-than":
        return { operator: "lessThan", value };
      case "less-or-equal":
        return { operator: "lessThanOrEqual", value };
      case "equals":
        return { operator: "equal", value };
      case "not-equals":
        return { operator: "notEqual", value };
    }
  }
  const operand = condition.operands[0];
  const literal = escapeNativeWildcardLiteral(operand);
  switch (condition.operator) {
    case "equals":
      return { operator: "equal", value: literal };
    case "not-equals":
      return { operator: "notEqual", value: literal };
    case "contains":
      return { operator: "equal", value: `*${literal}*` };
    case "does-not-contain":
      return { operator: "notEqual", value: `*${literal}*` };
    case "begins-with":
      return { operator: "equal", value: `${literal}*` };
    case "ends-with":
      return { operator: "equal", value: `*${literal}` };
    case "matches":
      return { operator: "equal", value: operand };
    case "does-not-match":
      return { operator: "notEqual", value: operand };
    case "greater-than":
      return { operator: "greaterThan", value: operand };
    case "greater-or-equal":
      return { operator: "greaterThanOrEqual", value: operand };
    case "less-than":
      return { operator: "lessThan", value: operand };
    case "less-or-equal":
      return { operator: "lessThanOrEqual", value: operand };
  }
}

/**
 * The native `<customFilters>` a condition is written as, the inverse of
 * {@link conditionFilterFromNativeAutoFilter}: one line for a single
 * comparison or text condition, `>= lo And <= hi` for between, and the two
 * lines of a pair under its join. Ranked and average criteria are not custom
 * filters (they are `top10` / `dynamicFilter`) and return null.
 */
export function nativeCustomFiltersFromCondition(
  condition: ConditionFilter,
): Extract<NativeAutoFilterCriterion, { kind: "custom" }> | null {
  if (condition.type === "compound") {
    return {
      kind: "custom",
      and: condition.join === "and",
      filters: condition.conditions.map(nativeCustomFilterFromCriterion),
    };
  }
  if (condition.type === "number" && condition.operator === "between") {
    const lo = condition.operands[0] ?? 0;
    const hi = condition.operands[1] ?? 0;
    return {
      kind: "custom",
      and: true,
      filters: [
        { operator: "greaterThanOrEqual", value: nativeNumberText(Math.min(lo, hi)) },
        { operator: "lessThanOrEqual", value: nativeNumberText(Math.max(lo, hi)) },
      ],
    };
  }
  if (!isCustomCriterionCondition(condition)) return null;
  return {
    kind: "custom",
    and: false,
    filters: [nativeCustomFilterFromCriterion(condition)],
  };
}

/**
 * Parse a cell's display value to a finite number, or `null` when the cell is
 * blank / non-numeric. Leading/trailing whitespace is trimmed first (matching
 * the sort/filter helpers in sort-filter.ts).
 */
export function parseNumeric(displayValue: string): number | null {
  const trimmed = displayValue.trim();
  if (trimmed === "") return null;
  const n = Number(trimmed);
  if (!Number.isFinite(n)) return null;
  return n;
}

/**
 * Compute the average of the numeric values in a column's value set (for
 * above-average / below-average filters). Blank / non-numeric cells are
 * excluded. Returns `null` when no numeric values are present (every row would
 * be hidden, mirroring Excel).
 */
type NumericSource = string | number | null | undefined;

function numericSourceValue(value: NumericSource): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  return typeof value === "string" ? parseNumeric(value) : null;
}

export function columnAverage(columnValues: readonly NumericSource[]): number | null {
  let sum = 0;
  let count = 0;
  for (const v of columnValues) {
    const n = numericSourceValue(v);
    if (n !== null) {
      sum += n;
      count += 1;
    }
  }
  return count === 0 ? null : sum / count;
}

/**
 * Compute the boundary for one Excel Top10 mode from the ORIGINAL numeric
 * column. Item counts are counts of cells, not distinct values; every tie at
 * the boundary remains visible. Percent uses ceil(numericCount * p / 100),
 * which is the smallest whole-cell population satisfying the requested share.
 *
 * `null` means no numeric row matches (empty numeric domain or 0 percent).
 */
export function rankedNumberThreshold(
  columnValues: readonly NumericSource[],
  operator: RankedNumberOperator,
  amount: number,
): number | null {
  const percent = operator === "top-percent" || operator === "bottom-percent";
  if (
    !Number.isFinite(amount) ||
    (percent && (amount < 0 || amount > 100)) ||
    (!percent && (!Number.isInteger(amount) || amount < 1 || amount > 255))
  ) {
    return null;
  }
  const numeric = columnValues
    .map(numericSourceValue)
    .filter((value): value is number => value !== null)
    .sort((left, right) => left - right);
  if (numeric.length === 0) return null;
  const requested = percent
    ? Math.ceil((numeric.length * amount) / 100)
    : Math.trunc(amount);
  const count = Math.min(numeric.length, Math.max(0, requested));
  if (count === 0) return null;
  return operator === "top-items" || operator === "top-percent"
    ? numeric[numeric.length - count]
    : numeric[count - 1];
}

/** Backward-compatible helper for callers/tests that need the visible strings. */
export function topNValues(
  columnValues: readonly string[],
  n: number,
): ReadonlySet<string> {
  const threshold = rankedNumberThreshold(
    columnValues,
    "top-items",
    Math.max(1, Math.round(n)),
  );
  return new Set(
    threshold === null
      ? []
      : columnValues
          .filter((value) => (parseNumeric(value) ?? -Infinity) >= threshold)
          .map((value) => value.trim()),
  );
}

/** Exact pre-native predicate retained only for legacy top-n 256..500 state. */
function legacyTopNValues(
  columnValues: readonly string[],
  n: number,
): ReadonlySet<string> {
  const count = Math.max(1, Math.round(n));
  const numericMap = new Map<number, string>();
  for (const value of columnValues) {
    const number = parseNumeric(value);
    if (number !== null && !numericMap.has(number)) {
      numericMap.set(number, value.trim());
    }
  }
  const selected = new Set(
    Array.from(numericMap.keys())
      .sort((left, right) => right - left)
      .slice(0, count),
  );
  return new Set(
    columnValues
      .filter((value) => {
        const number = parseNumeric(value);
        return number !== null && selected.has(number);
      })
      .map((value) => value.trim()),
  );
}

// ---------------------------------------------------------------------------
// Predicate builders
// ---------------------------------------------------------------------------

/**
 * Build a row predicate from a {@link TextCondition} and a raw display value.
 * All comparisons are case-insensitive (locale-independent toLowerCase()), but
 * neither side is trimmed: like Excel's custom AutoFilter, spaces are part of
 * the value, so a cell " A" equals the operand " A" and a cell "A" does not.
 * Callers pass the cell's display exactly as shown (no pre-trim).
 */
function textPredicate(
  cond: TextCondition,
): (display: string, numericValue?: number | null) => boolean {
  const { operator, operands } = cond;
  const target = operands[0].toLowerCase();
  // A wildcard pattern never matches a blank cell (Excel's `equals *` hides
  // blanks and `does not equal *` is the way to show them).
  const wildcard =
    operator === "matches" || operator === "does-not-match"
      ? new RegExp(`^${excelWildcardToRegexSource(operands[0])}$`, "is")
      : null;
  // A text comparison orders TEXT cells only: a number (by its authoritative
  // value when the caller has one, else by its display) and a blank never
  // compare against a text value.
  const textOrder = (display: string, numericValue?: number | null) => {
    if (display === "") return null;
    const numeric =
      numericValue === undefined
        ? parseNumeric(display)
        : numericSourceValue(numericValue);
    return numeric === null ? compareSpreadsheetText(display, operands[0]) : null;
  };
  return (display: string, numericValue?: number | null) => {
    const val = display.toLowerCase();
    switch (operator) {
      case "greater-than": {
        const order = textOrder(display, numericValue);
        return order !== null && order > 0;
      }
      case "greater-or-equal": {
        const order = textOrder(display, numericValue);
        return order !== null && order >= 0;
      }
      case "less-than": {
        const order = textOrder(display, numericValue);
        return order !== null && order < 0;
      }
      case "less-or-equal": {
        const order = textOrder(display, numericValue);
        return order !== null && order <= 0;
      }
      case "matches":
        return display !== "" && wildcard!.test(display);
      case "does-not-match":
        return display === "" || !wildcard!.test(display);
      case "contains":
        return val.includes(target);
      case "does-not-contain":
        return !val.includes(target);
      case "begins-with":
        return val.startsWith(target);
      case "ends-with":
        return val.endsWith(target);
      case "equals":
        return val === target;
      case "not-equals":
        return val !== target;
      default: {
        // Exhaustive check: TypeScript ensures this branch is never reached.
        const _: never = operator;
        void _;
        return false;
      }
    }
  };
}

/**
 * Build a row predicate from a {@link NumberCondition}, the column's complete
 * value set (needed for ranked / average), and a raw display value. Blank /
 * non-numeric cells always return false for numeric filters (they are not in
 * the numeric domain).
 */
function numberPredicate(
  cond: NumberCondition,
  columnValues: readonly string[],
  numericColumnValues?: readonly (number | null)[],
): (display: string, numericValue?: number | null) => boolean {
  const { operator, operands } = cond;

  // Pre-compute column-level statistics once (not per-row).
  let rankedThreshold: number | null = null;
  let legacyTopSet: ReadonlySet<string> | null = null;
  let avg: number | null = null;
  if (operator === "top-n") {
    legacyTopSet = legacyTopNValues(columnValues, operands[0] ?? 10);
  } else if (
    operator === "top-items" ||
    operator === "bottom-items" ||
    operator === "top-percent" ||
    operator === "bottom-percent"
  ) {
    rankedThreshold = rankedNumberThreshold(
      numericColumnValues ?? columnValues,
      operator,
      operands[0] ?? 10,
    );
  } else if (operator === "above-average" || operator === "below-average") {
    avg = columnAverage(numericColumnValues ?? columnValues);
  }

  return (display: string, numericValue?: number | null) => {
    // Non-numeric cells are excluded from every numeric filter.
    const n =
      numericValue === undefined
        ? parseNumeric(display)
        : numericSourceValue(numericValue);
    if (n === null) return false;

    switch (operator) {
      case "greater-than":
        return n > (operands[0] ?? 0);
      case "less-than":
        return n < (operands[0] ?? 0);
      case "greater-or-equal":
        return n >= (operands[0] ?? 0);
      case "less-or-equal":
        return n <= (operands[0] ?? 0);
      case "equals":
        return n === (operands[0] ?? 0);
      case "not-equals":
        return n !== (operands[0] ?? 0);
      case "between": {
        const lo = operands[0] ?? 0;
        const hi = operands[1] ?? 0;
        // Inclusive on both sides (Excel parity).
        return n >= Math.min(lo, hi) && n <= Math.max(lo, hi);
      }
      case "top-n":
        return legacyTopSet?.has(display.trim()) === true;
      case "top-items":
      case "top-percent":
        return rankedThreshold !== null && n >= rankedThreshold;
      case "bottom-items":
      case "bottom-percent":
        return rankedThreshold !== null && n <= rankedThreshold;
      case "above-average":
        return avg !== null && n > avg;
      case "below-average":
        return avg !== null && n < avg;
      default: {
        const _: never = operator;
        void _;
        return false;
      }
    }
  };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Build a row predicate for an AutoFilter condition.
 *
 * @param condition    - the condition descriptor ({@link ConditionFilter}).
 * @param columnValues - ALL display values present in the column's data rows
 *   (excluding the header). Required for ranked and average calculations; pass
 *   an empty array for other operators (they ignore it).
 * @returns a `(displayValue: string) => boolean` predicate to apply row-by-row.
 *
 * @example
 * ```ts
 * const pred = buildRowPredicate(
 *   { type: "text", operator: "contains", operands: ["apple"] },
 *   [],
 * );
 * pred("Pineapple"); // → true
 * pred("banana");   // → false
 * ```
 */
export function buildRowPredicate(
  condition: ConditionFilter,
  columnValues: readonly string[],
  numericColumnValues?: readonly (number | null)[],
): (displayValue: string, numericValue?: number | null) => boolean {
  if (condition.type === "compound") {
    // Each line is an ordinary condition; neither uses column statistics, so
    // the pair is decided row by row.
    const [first, second] = condition.conditions.map((line) =>
      buildRowPredicate(line, columnValues, numericColumnValues),
    );
    return condition.join === "and"
      ? (display, numericValue) =>
          first(display, numericValue) && second(display, numericValue)
      : (display, numericValue) =>
          first(display, numericValue) || second(display, numericValue);
  }
  if (condition.type === "text") {
    return textPredicate(condition);
  }
  return numberPredicate(condition, columnValues, numericColumnValues);
}

/**
 * Apply a {@link ConditionFilter} to a column in a row matrix, returning the
 * indices (into `rows`) of rows that PASS the condition. Header row (index 0)
 * is always included. A `null` / undefined condition means "no filter" — every
 * row index is returned.
 *
 * @param rows         - row-major matrix of cells (each cell has `.value`).
 * @param colIndex     - the column to test.
 * @param condition    - the condition to apply (or null for "all visible").
 * @param columnValues - all display values in the column (for ranked / average
 *   calculations). Duplicate rows must remain present because ranking counts
 *   cells, then includes every tie at the boundary.
 */
export function applyConditionFilter(
  rows: ReadonlyArray<ReadonlyArray<{ value?: string } | undefined>>,
  colIndex: number,
  condition: ConditionFilter | null | undefined,
  columnValues: readonly string[],
): number[] {
  const visible: number[] = [];
  if (!condition) {
    return rows.map((_, i) => i);
  }
  const pred = buildRowPredicate(condition, columnValues);
  for (let i = 0; i < rows.length; i += 1) {
    if (i === 0) {
      // Header row is always visible.
      visible.push(i);
      continue;
    }
    // Untrimmed: the text predicate compares the value exactly as shown.
    const display = rows[i][colIndex]?.value ?? "";
    if (pred(display)) visible.push(i);
  }
  return visible;
}
