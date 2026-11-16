/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ConditionalFormatRule, ConditionalRange } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { parseSourceRange, type ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";

/*
 * Which rules a literal edit cannot reach past its own cell.
 *
 * The grid repaints only the edited rows after a literal edit when every
 * other row's paint is provably unchanged. A rule whose outcome for a cell
 * reads only that cell's own value keeps that proof; a rule that reads other
 * cells (a range statistic, a formula, a source range, a controlling cell)
 * or paints onto a neighbour (a conditional border shared with the next
 * cell) does not, and the sheet keeps the broad repaint.
 */

/** Rule kinds decided by the covered cell's own value. */
const CELL_LOCAL_CONDITIONAL_KINDS = new Set([
  "comparison",
  "textContains",
  "dateOccurring",
]);

function hasBorder(rule: ConditionalFormatRule): boolean {
  // Scale, bar and icon rules carry no differential style at all.
  const edges = "borderEdges" in rule ? rule.borderEdges : undefined;
  return Boolean(edges && Object.keys(edges).length > 0);
}

/**
 * True when every rule judges (and offers choices for) a cell from that cell
 * alone: fixed lists, and number, text, date and time bounds that are not
 * formulas. A list fed by a range, a formula rule and a dependent (cascade)
 * list read other cells.
 */
export function validationRulesAreCellLocal(
  rules: readonly ValidationRule[],
): boolean {
  return rules.every((rule) => {
    switch (rule.kind) {
      case "any":
      case "text":
      case "date":
      case "time":
        return true;
      case "list":
        return !rule.sourceRange;
      case "number":
        return !rule.minFormula && !rule.maxFormula;
      default:
        return false;
    }
  });
}

/** Rule kinds that colour a cell from statistics over their whole range. */
const RANGE_STATISTIC_KINDS = new Set([
  "colorScale",
  "dataBar",
  "iconSet",
  "topBottom",
  "aboveBelowAverage",
  "duplicateValues",
]);

/** A list Source written with both ends absolute on the same sheet. */
const ABSOLUTE_SOURCE_RE = /^=?\s*\$[A-Za-z]{1,3}\$\d{1,7}(?::\$[A-Za-z]{1,3}\$\d{1,7})?\s*$/;

/**
 * The areas a literal edit must stay out of for the other rows' paint to
 * stand, or null when some rule's reach is not known ahead of the edit.
 *
 * Cell-local rules reach nothing past their cell. A range statistic (colour
 * scale, data bar, icon set, top/bottom, average, duplicates) reaches its own
 * areas, and an absolute same-sheet list Source reaches that source: an edit
 * inside either repaints every row, one outside repaints only its own rows.
 * Formula rules, cascades, formula bounds, relative or cross-sheet sources
 * and conditional borders (a cell borrows its neighbour's edges) have no
 * reach that can be known ahead, so such a sheet keeps the broad repaint.
 */
export function literalEditReach(
  conditional: readonly ConditionalFormatRule[],
  validations: readonly ValidationRule[],
): ConditionalRange[] | null {
  const reach: ConditionalRange[] = [];
  for (const rule of conditional) {
    // A rule kept preserve-only (nativeBlocker) is still painted by its kind;
    // the blocker only stops it from being edited or written natively.
    if (hasBorder(rule)) return null;
    const kind = rule.kind ?? "comparison";
    // A text rule with a formula of its own reads what the formula names.
    if (rule.kind === "textContains" && rule.formula !== undefined) return null;
    if (CELL_LOCAL_CONDITIONAL_KINDS.has(kind)) continue;
    if (!RANGE_STATISTIC_KINDS.has(kind)) return null;
    reach.push(...(rule.areas?.length ? rule.areas : [rule.range]));
  }
  for (const rule of validations) {
    if (validationRulesAreCellLocal([rule])) continue;
    if (rule.kind !== "list" || !rule.sourceRange || !ABSOLUTE_SOURCE_RE.test(rule.sourceRange)) return null;
    const source = parseSourceRange(rule.sourceRange);
    if (!source || source.sheet) return null;
    reach.push({ r1: source.r1, c1: source.c1, r2: source.r2, c2: source.c2 });
  }
  return reach;
}

/**
 * The merges that span rows: an edit inside one repaints every row, as the row
 * that shows such a merge may not be the edited one (its anchor scrolled out, a
 * stand-in draws it). A merge along one row is that row's own.
 */
export function multiRowMergeAreas(
  merges: readonly { row: number; col: number; rowSpan?: number; colSpan?: number }[],
): ConditionalRange[] {
  const areas: ConditionalRange[] = [];
  for (const { row, col, rowSpan = 1, colSpan = 1 } of merges) {
    if (rowSpan > 1) areas.push({ r1: row, c1: col, r2: row + rowSpan - 1, c2: col + colSpan - 1 });
  }
  return areas;
}

/** Whether any "row:col" key falls inside one of the areas. */
export function keysTouchAreas(
  keys: Iterable<string>,
  areas: readonly ConditionalRange[],
): boolean {
  if (areas.length === 0) return false;
  for (const key of keys) {
    const split = key.indexOf(":");
    const row = Number(key.slice(0, split));
    const col = Number(key.slice(split + 1));
    for (const area of areas) {
      if (row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2) return true;
    }
  }
  return false;
}
