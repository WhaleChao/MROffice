/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { calculationRuleKey, MAX_CALCULATION_RULE_CELLS, MAX_CALCULATION_RULES, type CalculationRuleRequest } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";
import { blankRelativeStopReferences, type ConditionalFormatRule } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { listSourceFormula, listSourceOptionsFormula, type ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";

type FormulaStop = { type?: string; formula?: string } | undefined;

/** The formulas a conditional format evaluates for each of its cells
 * (`ConditionalDisplayOptions.evalFormula`): a formula rule's formula, the
 * formula stops of a colour scale, data bar or icon set. */
export function conditionalRuleFormulas(rule: ConditionalFormatRule): string[] {
  if (rule.kind === "formula") return [rule.formula];
  
  if (rule.kind === "textContains") return rule.formula !== undefined ? [rule.formula] : [];
  const record = rule as unknown as { stops?: FormulaStop[]; min?: FormulaStop; mid?: FormulaStop; max?: FormulaStop };
  const formulas = (stops: FormulaStop[]) => stops
    .filter((stop): stop is { type: string; formula: string } => stop?.type === "formula" && typeof stop.formula === "string" && stop.formula !== "")
    .map((stop) => stop.formula);
  // A colour scale's or data bar's threshold is one value for the range, and
  // a relative reference in it reads blank (Excel: A1*2+8 is 8, $A$1*2+8 is
  // 30): the display asks for the formula with those references blanked
  // (withFormulaStopsResolved), so that is the formula asked of the
  // calculation. Asked as written, the answer never matched the question and
  // the scale stayed at its far colour (P14, screen and print).
  return [...formulas(record.stops ?? []), ...formulas([record.min, record.mid, record.max]).map(blankRelativeStopReferences)];
}

/** The rectangle spanning every area a conditional format applies to. */
function ruleExtent(rule: { range: CalculationRuleRequest["range"]; areas?: CalculationRuleRequest["range"][] }): CalculationRuleRequest["range"] {
  const areas = rule.areas && rule.areas.length > 1 ? rule.areas : null;
  if (!areas) return rule.range;
  return areas.reduce(
    (extent, area) => ({ r1: Math.min(extent.r1, area.r1), c1: Math.min(extent.c1, area.c1), r2: Math.max(extent.r2, area.r2), c2: Math.max(extent.c2, area.c2) }),
    rule.range,
  );
}

/** The formulas a validation rule evaluates for each of its cells
 * (`ValidationContext.evalFormula`), spelled exactly as it passes them: a
 * custom rule's formula, a number rule's formula bounds, a list's formula
 * Source as the options it lists. */
export function validationRuleFormulas(rule: ValidationRule): string[] {
  if (rule.kind === "formula") return [rule.formula];
  // A formula Source (=INDIRECT(A2)): each cell's list (listSourceOptionsFormula).
  if (rule.kind === "list") {
    const source = listSourceFormula(rule);
    return source ? [listSourceOptionsFormula(source)] : [];
  }
  if (rule.kind === "number") {
    return [rule.minFormula, rule.maxFormula].filter((formula): formula is string => formula !== undefined).map((formula) => `=${formula}`);
  }
  return [];
}

/**
 * The rule formulas of one sheet the calculation session evaluates with each
 * calculation (C2), in paint order: conditional formats first, then data
 * validation. A rule the bounds leave out (more than MAX_CALCULATION_RULES
 * formulas or MAX_CALCULATION_RULE_CELLS cells in all) is not sent: it is
 * unanswered (does not format or match; nothing evaluates it in the browser).
 */
export function calculationRuleRequests(
  sheet: { readonly id: string; readonly rowCount: number; readonly columnCount: number },
  conditional: readonly ConditionalFormatRule[],
  validations: readonly ValidationRule[] = [],
  /** Other expressions read at one cell (a =HYPERLINK formula's link
   * location), after the rules: what the budget leaves goes unanswered. */
  cellFormulas: ReadonlyArray<{ readonly row: number; readonly col: number; readonly formula: string }> = [],
): CalculationRuleRequest[] {
  const requests: CalculationRuleRequest[] = [];
  // The answer is keyed by formula and anchor (calculationRuleKey): a cell's
  // value does not depend on how far the range reaches, so two rules sharing
  // both are one request over the larger range.
  const byKey = new Map<string, number>();
  let cells = 0;
  const areaOf = (range: CalculationRuleRequest["range"]) => (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  // `range` is what is evaluated, relative to the anchor (the rule's own
  // top-left); it may reach above or left of it (a later area of the rule).
  const add = (formula: string, range: CalculationRuleRequest["range"], anchorRow = range.r1, anchorCol = range.c1): boolean => {
    const clamped = { r1: range.r1, c1: range.c1, r2: Math.min(range.r2, sheet.rowCount - 1), c2: Math.min(range.c2, sheet.columnCount - 1) };
    if (clamped.r2 < clamped.r1 || clamped.c2 < clamped.c1) return true;
    const request = { sheetId: sheet.id, formula, anchorRow, anchorCol, range: clamped };
    const key = calculationRuleKey(request);
    const index = byKey.get(key);
    if (index !== undefined) {
      const before = requests[index].range;
      const union = { r1: Math.min(before.r1, clamped.r1), c1: Math.min(before.c1, clamped.c1), r2: Math.max(before.r2, clamped.r2), c2: Math.max(before.c2, clamped.c2) };
      const grown = areaOf(union) - areaOf(before);
      if (grown <= 0) return true;
      if (cells + grown > MAX_CALCULATION_RULE_CELLS) return false;
      requests[index] = { ...requests[index], range: union }; cells += grown;
      return true;
    }
    const area = areaOf(clamped);
    if (requests.length >= MAX_CALCULATION_RULES || cells + area > MAX_CALCULATION_RULE_CELLS) return false;
    byKey.set(key, requests.length); cells += area; requests.push(request);
    return true;
  };
  for (const rule of conditional) {
    // Every area of the rule is answered, not just its first (live audit
    // S-031: clearing the formats of P121 left =P120>15 on P120 and P122,
    // and P122 lost its colour). The answer is one rectangle per formula, so
    // it spans the areas; when that does not fit, the first area still does.
    const extent = ruleExtent(rule);
    for (const formula of conditionalRuleFormulas(rule)) {
      if (!add(formula, extent, rule.range.r1, rule.range.c1) && extent !== rule.range) add(formula, rule.range);
    }
  }
  for (const rule of validations) {
    
    const extent = ruleExtent(rule);
    for (const formula of validationRuleFormulas(rule)) {
      if (extent === rule.range) add(formula, rule.range);
      else if (!add(formula, extent, rule.range.r1, rule.range.c1)) add(formula, rule.range);
    }
  }
  for (const cell of cellFormulas) add(cell.formula, { r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col });
  return requests;
}
