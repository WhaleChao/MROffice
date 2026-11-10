/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type FilterCellValue = string | number | boolean | null;

export interface AdvancedFilterCriteria {
  /** Criteria header row, aligned to the list's header labels. */
  headers: readonly string[];
  /** One row per alternative; short rows are treated as blank-padded. */
  rows: readonly (readonly FilterCellValue[])[];
  /**
   * Raw worksheet formulas aligned to `rows`. A formula is significant only
   * under a blank/non-list header; formula-looking cells under a real list
   * header keep the ordinary displayed-value criteria semantics.
   */
  formulas?: readonly (readonly (string | null)[])[];
}

export interface AdvancedFilterFormulaPosition {
  /** Zero-based data-row index inside `AdvancedFilterInput.rows`. */
  dataRow: number;
  /** Zero-based criteria-row/column indices inside the criteria body. */
  criteriaRow: number;
  criteriaColumn: number;
}

export interface AdvancedFilterInput {
  /** The list's header labels, in column order. */
  headers: readonly string[];
  /** Data rows, in sheet order. */
  rows: readonly (readonly FilterCellValue[])[];
  criteria: AdvancedFilterCriteria;
  /** Drop rows whose visible content repeats an earlier match. */
  uniqueOnly?: boolean;
  /**
   * Evaluate one calculated criterion after rebasing its relative references
   * to `position.dataRow`. `null` means the expression is unsupported/erroring
   * and therefore fails closed instead of widening the filter.
   */
  evaluateFormulaCriterion?: (
    formula: string,
    position: AdvancedFilterFormulaPosition,
  ) => boolean | null;
}

export interface AdvancedFilterResult {
  /** Indices into `rows` that satisfy the criteria, in order. */
  matched: number[];
  /** Matches removed because `uniqueOnly` had already kept an identical row. */
  duplicatesRemoved: number;
  /**
   * Criteria headers that match no list header. Excel silently ignores them;
   * surfacing them lets the dialog warn instead of returning a puzzling set.
   */
  unmatchedHeaders: string[];
  /**
   * Criteria cells that look like a comparison but carry nothing to compare
   * against — `>=` on its own. Dropping those quietly WIDENS the filter, which
   * is the dangerous direction: the user sees more rows than they asked for and
   * has no way to tell the criterion was ignored.
   */
  invalidCriteria: string[];
}

/** Guard against a pathological criteria range. */
export const MAX_CRITERIA_ROWS = 1_000;

type Comparison = "=" | "<>" | ">" | ">=" | "<" | "<=";

interface ParsedCriterion {
  comparison: Comparison;
  /** Numeric operand when the criterion compares numbers. */
  number: number | null;
  /** Text operand, already lower-cased for case-insensitive matching. */
  text: string;
  /** True for `="…"`, which means "equals exactly", not "begins with". */
  exact: boolean;
  /** True for a bare `=` / `<>` with no operand: is (not) blank. */
  blankTest: boolean;
  /**
   * Compiled matcher for an operand carrying Excel's criteria wildcards, or
   * null when the operand is a plain literal.
   */
  wildcard: RegExp | null;
}

type CriteriaColumn =
  | { kind: "field"; column: number }
  | { kind: "formula"; criteriaColumn: number }
  | { kind: "unmatched" }
  | { kind: "ignored" };

type CompiledCriterion =
  | { kind: "field"; column: number; criterion: ParsedCriterion }
  | {
      kind: "formula";
      formula: string;
      criteriaRow: number;
      criteriaColumn: number;
    };

/**
 * Excel's criteria wildcards: `*` matches any run of characters, `?` matches
 * exactly one, and `~` escapes either so a literal asterisk can be searched
 * for. The pattern is anchored at the start because criteria are "begins
 * with"; `="Sm?th"` anchors both ends, matching Excel's exact form.
 *
 * Runs of `*` collapse into one, so a criteria cell full of asterisks cannot
 * hand the regex engine a backtracking bomb.
 */
function wildcardMatcher(pattern: string, exact: boolean): RegExp {
  const escape = (char: string) => char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let source = "";
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    const next = pattern[index + 1];
    if (char === "~" && (next === "*" || next === "?" || next === "~")) {
      source += escape(next);
      index += 1;
      continue;
    }
    if (char === "*") {
      if (!source.endsWith(".*")) source += ".*";
      continue;
    }
    source += char === "?" ? "." : escape(char);
  }
  return new RegExp(`^${source}${exact ? "$" : ""}`);
}

function matcherFor(operand: string, exact: boolean): RegExp | null {
  return /[*?]/.test(operand)
    ? wildcardMatcher(operand.toLowerCase(), exact)
    : null;
}

function normalizeText(value: FilterCellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

function numericValue(value: FilterCellValue): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  const text = normalizeText(value);
  if (!text) return null;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

/** Parse one criteria cell into a predicate description. */
export function parseCriterion(raw: FilterCellValue): ParsedCriterion | null {
  const text = normalizeText(raw);
  if (!text) return null;
  const match = /^(<>|>=|<=|=|>|<)?\s*(.*)$/s.exec(text);
  const comparison = (match?.[1] ?? "") as Comparison | "";
  const operandRaw = (match?.[2] ?? "").trim();
  const quoted = /^"(.*)"$/s.exec(operandRaw);
  const operand = quoted ? quoted[1] : operandRaw;
  if (!comparison) {
    return {
      comparison: "=",
      number: numericValue(operand),
      text: operand.toLowerCase(),
      exact: false,
      blankTest: false,
      wildcard: matcherFor(operand, false),
    };
  }
  if (!operand) {
    // `=` alone means blank, `<>` alone means non-blank (Excel).
    if (comparison === "=" || comparison === "<>") {
      return {
        comparison,
        number: null,
        text: "",
        exact: true,
        blankTest: true,
        wildcard: null,
      };
    }
    return null;
  }
  // Any explicit operator compares the whole value. Microsoft: a criterion
  // shown as `=Davolio` (entered as ="=Davolio") matches exactly, and only a
  // bare `Davolio` with no operator means "begins with" (the branch above).
  const exact = true;
  return {
    comparison,
    number: numericValue(operand),
    text: operand.toLowerCase(),
    exact,
    blankTest: false,
    wildcard:
      comparison === "=" || comparison === "<>"
        ? matcherFor(operand, exact)
        : null,
  };
}

function satisfies(
  value: FilterCellValue,
  criterion: ParsedCriterion,
): boolean {
  const text = normalizeText(value);
  if (criterion.blankTest) {
    return criterion.comparison === "=" ? text === "" : text !== "";
  }
  const cellNumber = numericValue(value);
  if (criterion.number !== null && cellNumber !== null) {
    switch (criterion.comparison) {
      case "=":
        return cellNumber === criterion.number;
      case "<>":
        return cellNumber !== criterion.number;
      case ">":
        return cellNumber > criterion.number;
      case ">=":
        return cellNumber >= criterion.number;
      case "<":
        return cellNumber < criterion.number;
      case "<=":
        return cellNumber <= criterion.number;
    }
  }
  // A numeric criterion asks a question about numbers, so a cell that holds
  // none does not answer it. Without this an ordering comparison fell through
  // to the text branch and "n/a" > "100" kept a placeholder row in a `>100`
  // filter — the row the filter exists to remove.
  if (criterion.number !== null && cellNumber === null) {
    switch (criterion.comparison) {
      case ">":
      case ">=":
      case "<":
      case "<=":
        return false;
    }
  }

  const lower = text.toLowerCase();
  switch (criterion.comparison) {
    case "=":
      return criterion.wildcard
        ? criterion.wildcard.test(lower)
        : criterion.exact
          ? lower === criterion.text
          : lower.startsWith(criterion.text);
    case "<>":
      return criterion.wildcard
        ? !criterion.wildcard.test(lower)
        : criterion.exact
          ? lower !== criterion.text
          : !lower.startsWith(criterion.text);
    case ">":
      return lower > criterion.text;
    case ">=":
      return lower >= criterion.text;
    case "<":
      return lower < criterion.text;
    case "<=":
      return lower <= criterion.text;
  }
}

function headerIndex(headers: readonly string[], label: string): number {
  const wanted = label.trim().toLowerCase();
  return headers.findIndex((header) => header.trim().toLowerCase() === wanted);
}

/** Which list rows satisfy the criteria range. */
export function applyAdvancedFilter(
  input: AdvancedFilterInput,
): AdvancedFilterResult {
  const unmatchedHeaders: string[] = [];
  const invalidCriteria: string[] = [];
  const formulaRows = input.criteria.formulas ?? [];
  const criteriaColumns: CriteriaColumn[] = input.criteria.headers.map(
    (label, criteriaColumn) => {
      const trimmed = label.trim();
      const target = trimmed ? headerIndex(input.headers, trimmed) : -1;
      if (target >= 0) return { kind: "field", column: target };
      const hasFormula = formulaRows.slice(0, MAX_CRITERIA_ROWS).some((row) => {
        const formula = row[criteriaColumn];
        return typeof formula === "string" && formula.trim().startsWith("=");
      });
      if (hasFormula) return { kind: "formula", criteriaColumn };
      if (!trimmed) return { kind: "ignored" };
      if (!unmatchedHeaders.includes(trimmed)) {
        unmatchedHeaders.push(trimmed);
      }
      return { kind: "unmatched" };
    },
  );

  const alternatives = input.criteria.rows
    .slice(0, MAX_CRITERIA_ROWS)
    .map((row, criteriaRow) =>
      criteriaColumns
        .map((binding, index): CompiledCriterion | null => {
          if (binding.kind === "unmatched") return null;
          if (binding.kind === "formula") {
            const rawFormula = formulaRows[criteriaRow]?.[index];
            const formula =
              typeof rawFormula === "string" &&
              rawFormula.trim().startsWith("=")
                ? rawFormula.trim()
                : null;
            if (formula) {
              return {
                kind: "formula",
                formula,
                criteriaRow,
                criteriaColumn: binding.criteriaColumn,
              };
            }
            const visible = normalizeText(row[index] ?? null);
            if (visible && !invalidCriteria.includes(visible)) {
              invalidCriteria.push(visible);
            }
            return null;
          }
          if (binding.kind === "ignored") {
            const visible = normalizeText(row[index] ?? null);
            if (visible && !invalidCriteria.includes(visible)) {
              invalidCriteria.push(visible);
            }
            return null;
          }
          const raw = row[index] ?? null;
          const criterion = parseCriterion(raw);
          if (!criterion) {
            const text = String(raw ?? "").trim();
            // Blank is a legitimate "no criterion"; a stray operator is not.
            if (text && !invalidCriteria.includes(text)) {
              invalidCriteria.push(text);
            }
            return null;
          }
          return { kind: "field", column: binding.column, criterion };
        })
        .filter((entry): entry is CompiledCriterion => entry !== null),
    );

  if (invalidCriteria.length > 0) {
    return {
      matched: [],
      duplicatesRemoved: 0,
      unmatchedHeaders,
      invalidCriteria,
    };
  }

  // An entirely blank criteria range matches everything, exactly as Excel does.
  const meaningful = alternatives.filter((row) => row.length > 0);
  const matchesAll = meaningful.length === 0;

  const matched: number[] = [];
  const seen = new Set<string>();
  let duplicatesRemoved = 0;
  input.rows.forEach((row, index) => {
    const ok =
      matchesAll ||
      meaningful.some((alternative) =>
        alternative.every((entry) => {
          if (entry.kind === "field") {
            return satisfies(row[entry.column] ?? null, entry.criterion);
          }
          let result: boolean | null = null;
          try {
            result =
              input.evaluateFormulaCriterion?.(entry.formula, {
                dataRow: index,
                criteriaRow: entry.criteriaRow,
                criteriaColumn: entry.criteriaColumn,
              }) ?? null;
          } catch {
            result = null;
          }
          if (result === null && !invalidCriteria.includes(entry.formula)) {
            invalidCriteria.push(entry.formula);
          }
          return result === true;
        }),
      );
    if (!ok) return;
    if (input.uniqueOnly) {
      const key = row
        .map((value) => normalizeText(value).toLowerCase())
        .join(" ");
      if (seen.has(key)) {
        duplicatesRemoved += 1;
        return;
      }
      seen.add(key);
    }
    matched.push(index);
  });

  return {
    matched: invalidCriteria.length > 0 ? [] : matched,
    duplicatesRemoved: invalidCriteria.length > 0 ? 0 : duplicatesRemoved,
    unmatchedHeaders,
    invalidCriteria,
  };
}

/** Rows to HIDE for "filter the list in place" — the complement of the matches. */
export function hiddenRowsForAdvancedFilter(input: AdvancedFilterInput): {
  hidden: number[];
  result: AdvancedFilterResult;
} {
  const result = applyAdvancedFilter(input);
  const kept = new Set(result.matched);
  const hidden: number[] = [];
  for (let index = 0; index < input.rows.length; index += 1) {
    if (!kept.has(index)) hidden.push(index);
  }
  return { hidden, result };
}
