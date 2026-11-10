/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Conditional Formatting rule writes shared by the window's gallery
 * and the spreadsheet editor commands. Behaviour unchanged.
 */

import {
  isClassicNativeConditionalRuleEditable,
  ruleAreas as conditionalRuleAreas,
  type ConditionalFormatPreset,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { createConditionalRuleId } from "@mrnote/core/web/lib/spreadsheet/conditional-rule-draft.ts";

/** The rule parameters a gallery preset creates (a data bar takes Excel 2010+
 *  defaults, its Gradient / Solid Fill choice included). */
export function conditionalPresetParams(
  preset: ConditionalFormatPreset,
): Record<string, unknown> {
  // A data bar created here follows the defaults Excel has used since
  // 2010, not the 2007 ones: bars run from 0 to the full cell width, and
  // both ends are automatic — min(0, lowest) and max(0, highest), read
  // from the data whenever the bar is drawn — so the shortest bar stands
  // for zero on non-negative data, a negative entered later gets its zero
  // axis and negative bar, and a rule made before and one made after the
  // data crossed zero draw alike (live audit B247). With the 2007
  // defaults (10% to 90%, from the smallest value) 0.05 and 89 drew bars
  // about as long as 1,234 next to a 76,543.
  let presetParams = preset.params as Record<string, unknown>;
  if (presetParams.kind === "dataBar") {
    presetParams = {
      ...presetParams,
      minLength: 0,
      maxLength: 100,
      min: { type: "automatic" },
      max: { type: "automatic" },
    };
    // An imported workbook saves through original-package patchback,
    // which writes a new bar with the x14 twin Excel writes, so the
    // gallery's Gradient / Solid Fill choice is kept there too.
  }
  return presetParams;
}

/**
 * The sheet's rules with a new rule given a fresh id: a new rule goes first
 * (it wins over older ones), except on an original-package workbook, whose
 * native rule order is kept and the new rule appended.
 */
export function conditionalRulesWithNewRule(
  rules: readonly ConditionalFormatRule[],
  rule: ConditionalFormatRule,
  patchback: boolean,
): ConditionalFormatRule[] {
  rule.id = createConditionalRuleId();
  return patchback ? [...rules, rule] : [rule, ...rules];
}

/**
 * Clear Rules from Selected Cells takes only the selected cells out of each
 * rule (the rest of A1:A3 keeps its rule when A2 is cleared), so the cells
 * that change are the selected cells each rule covers; true when one of
 * those is locked under protection.
 */
export function conditionalClearTouchesLockedCells(
  affectedRules: readonly ConditionalFormatRule[],
  range: { r1: number; c1: number; r2: number; c2: number },
  rangeHasLockedCell: (area: { r1: number; c1: number; r2: number; c2: number }) => boolean,
): boolean {
  return affectedRules.some((rule) =>
    conditionalRuleAreas(rule).some((area) => {
      const r1 = Math.max(area.r1, range.r1);
      const c1 = Math.max(area.c1, range.c1);
      const r2 = Math.min(area.r2, range.r2);
      const c2 = Math.min(area.c2, range.c2);
      return r1 <= r2 && c1 <= c2 && rangeHasLockedCell({ r1, c1, r2, c2 });
    }),
  );
}

/** The Manage Rules / gallery writes a sheet's rule collection goes through. */
export type ConditionalFormatMutationAction =
  | "add"
  | "edit"
  | "delete"
  | "reorder"
  | "stopIfTrue"
  | "clearSelection"
  | "clearSheet"
  /** Conditional Formatting Rules Manager OK / Apply: every staged change
   * (edits, new, duplicated and deleted rules, order, Applies to, Stop If
   * True) as one collection write. */
  | "manage";

/**
 * On an original-package workbook a rule the native writer cannot model (a
 * blocker, or a kind it does not edit) may be deleted or moved, as in Excel's
 * Rules Manager: the writer removes its node or rewrites only its priority.
 * It may never change (the dialog keeps the same object) or appear from
 * nowhere.
 */
export function conditionalFormatPatchbackMutationIsLossless(
  currentRules: readonly ConditionalFormatRule[],
  nextRules: readonly ConditionalFormatRule[],
): boolean {
  const readOnlyRules = currentRules.filter(
    (rule) =>
      Boolean(rule.nativeBlocker) ||
      !isClassicNativeConditionalRuleEditable(rule),
  );
  return !nextRules.some(
    (rule) =>
      (Boolean(rule.nativeBlocker) ||
        !isClassicNativeConditionalRuleEditable(rule)) &&
      !readOnlyRules.includes(rule),
  );
}
