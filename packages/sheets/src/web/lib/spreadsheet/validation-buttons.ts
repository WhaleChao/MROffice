/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Data Validation: the dialog's rule from its fields, its OK (a rule
 * on the selection, or "Any value" removing it) and Clear All, as pure
 * transactions over a sheet's validation rules. Moved out of
 * SpreadsheetWindow so the ribbon and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps the dialog authority,
 * protection and native-patchback checks and the error text.
 */

import {
  buildRule,
  clearRulesInRange,
  parseDependentOptions,
  parseListValues,
  rangeHasRule,
  sameValidationSettingsAs,
  validationOperatorNeedsSecondValue,
  validationRuleAreas,
  type ErrorStyle,
  type ValidationFormulaShifter,
  type ValidationMessages,
  type ValidationOperator,
  type ValidationRange,
  type ValidationRule,
} from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";

/** Excel's Allow box. "number" is Decimal, or Whole number with `integer`;
 *  "dependent" is MRAI's cascade list, offered after Excel's own types. */
export type ValidationDialogKind =
  | "any"
  | "list"
  | "number"
  | "text"
  | "formula"
  | "date"
  | "time"
  | "dependent";

/** The Data Validation dialog's fields, as typed. */
export interface ValidationDraft {
  kind: ValidationDialogKind;
  operator: ValidationOperator;
  ignoreBlank: boolean;
  /** List: one Source box; "=..." is a live reference, anything else the items. */
  listValues: string;
  showDropDown: boolean;
  min: string;
  max: string;
  integer: boolean;
  formula: string;
  dateMin: string;
  dateMax: string;
  weekdaysOnly: boolean;
  notBeforeToday: boolean;
  notAfterToday: boolean;
  timeMin: string;
  timeMax: string;
  minLength: string;
  maxLength: string;
  depSourceRow: string;
  depSourceCol: string;
  depPerRow: boolean;
  depOptions: string;
  depDefault: string;
  showInputMessage: boolean;
  inputTitle: string;
  inputMessage: string;
  showErrorMessage: boolean;
  errorTitle: string;
  errorMessage: string;
  errorStyle: ErrorStyle;
}

/** The rule the dialog's fields describe on `range`, or null when invalid. */
/**
 * A date or time rule keeps a bound only as a date or time it can read; a
 * formula (=DATE(2026,12,31)) or unreadable text is not stored for those
 * kinds. Rather than write a one-sided rule that admits what the second
 * bound excluded, the draft is refused (live audit B149, date and time).
 */
function keepsEveryBound(
  rule: ValidationRule | null,
  min: string,
  max: string,
): ValidationRule | null {
  if (!rule || (rule.kind !== "date" && rule.kind !== "time")) return rule;
  if (min.trim() !== "" && rule.min === undefined) return null;
  if (max.trim() !== "" && rule.max === undefined) return null;
  return rule;
}

export function validationDraftRule(
  range: ValidationRange,
  draft: ValidationDraft,
): ValidationRule | null {
  // The Excel-style input/error message fields, shared by every kind. Blank
  // fields are dropped by normalizeMessages so a rule with no messages stays
  // byte-for-byte as before; the "stop" style is the implicit default.
  const messages: ValidationMessages = {
    ignoreBlank: draft.ignoreBlank,
    showInputMessage: draft.showInputMessage,
    showErrorMessage: draft.showErrorMessage,
    inputTitle: draft.inputTitle,
    inputMessage: draft.inputMessage,
    errorTitle: draft.errorTitle,
    errorMessage: draft.errorMessage,
    errorStyle: draft.errorStyle,
  };
  const operator = draft.operator;
  if (draft.kind === "any") {
    return buildRule(range, { kind: "any" }, messages);
  }
  if (draft.kind === "list") {
    // One Source box, as in Excel: "=..." is a live reference to the cells
    // holding the items, anything else is the items themselves.
    const source = draft.listValues.trim();
    const reference = source.startsWith("=");
    return buildRule(
      range,
      {
        kind: "list",
        values: reference ? [] : parseListValues(draft.listValues),
        ...(reference ? { sourceRange: source } : {}),
        showDropDown: draft.showDropDown,
      },
      messages,
    );
  }
  if (draft.kind === "number") {
    // A bound that is not a number is a formula or reference (=$H$15):
    // it is kept as written and evaluated per cell, never dropped for a
    // one-sided rule (live audit B149).
    const boundValue = (text: string): number | string | undefined => {
      const trimmed = text.trim();
      if (trimmed === "") return undefined;
      return Number.isFinite(Number(trimmed)) ? Number(trimmed) : trimmed;
    };
    const min = boundValue(draft.min);
    const max = validationOperatorNeedsSecondValue(operator)
      ? boundValue(draft.max)
      : undefined;
    return buildRule(
      range,
      {
        kind: "number",
        min,
        max,
        integer: draft.integer,
        operator: operator === "between" ? undefined : operator,
      },
      messages,
    );
  }
  if (draft.kind === "formula") {
    return buildRule(range, { kind: "formula", formula: draft.formula }, messages);
  }
  if (draft.kind === "date") {
    return keepsEveryBound(buildRule(
      range,
      {
        kind: "date",
        min: draft.dateMin.trim() === "" ? undefined : draft.dateMin.trim(),
        max:
          validationOperatorNeedsSecondValue(operator) &&
          draft.dateMax.trim() !== ""
            ? draft.dateMax.trim()
            : undefined,
        operator: operator === "between" ? undefined : operator,
        // "Weekdays only" maps to Mon–Fri (1..5 in getUTCDay ordering).
        weekdays: draft.weekdaysOnly ? [1, 2, 3, 4, 5] : undefined,
        notBefore: draft.notBeforeToday ? "today" : undefined,
        notAfter: draft.notAfterToday ? "today" : undefined,
      },
      messages,
    ), draft.dateMin, validationOperatorNeedsSecondValue(operator) ? draft.dateMax : "");
  }
  if (draft.kind === "time") {
    return keepsEveryBound(buildRule(
      range,
      {
        kind: "time",
        min: draft.timeMin.trim() === "" ? undefined : draft.timeMin.trim(),
        max:
          validationOperatorNeedsSecondValue(operator) &&
          draft.timeMax.trim() !== ""
            ? draft.timeMax.trim()
            : undefined,
        operator: operator === "between" ? undefined : operator,
      },
      messages,
    ), draft.timeMin, validationOperatorNeedsSecondValue(operator) ? draft.timeMax : "");
  }
  if (draft.kind === "dependent") {
    // A controlling cell is required; bad coords yield a null rule -> error.
    const parsed = parseDependentSource(draft.depSourceRow, draft.depSourceCol);
    const source = parsed ?? { row: -1, col: -1 };
    return buildRule(
      range,
      {
        kind: "dependent",
        source,
        sourceMode: draft.depPerRow ? "perRow" : "fixed",
        options: parseDependentOptions(draft.depOptions),
        optionsDefault: parseListValues(draft.depDefault),
      },
      messages,
    );
  }
  const minLength =
    draft.minLength.trim() === "" ? undefined : Number(draft.minLength);
  const maxLength =
    validationOperatorNeedsSecondValue(operator) && draft.maxLength.trim() !== ""
      ? Number(draft.maxLength)
      : undefined;
  return buildRule(
    range,
    {
      kind: "text",
      minLength,
      maxLength,
      operator: operator === "between" ? undefined : operator,
    },
    messages,
  );
}

/** What the dialog's OK writes: a rule, or the removal "Any value" means. */
export interface ValidationDefinePlan {
  removing: boolean;
  rule: ValidationRule | null;
}

/**
 * The OK's rule on `range`. Any value with no input message is Excel's
 * "no validation": OK removes the rule rather than writing an empty one.
 * Null when the fields do not make a valid rule.
 */
export function validationDefinePlan(
  range: ValidationRange,
  draft: ValidationDraft,
): ValidationDefinePlan | null {
  const removing =
    draft.kind === "any" && !draft.inputTitle.trim() && !draft.inputMessage.trim();
  const rule = removing ? null : validationDraftRule(range, draft);
  if (!removing && !rule) return null;
  return { removing, rule };
}

/**
 * The sheet's rules after the OK. A new rule on a range replaces the rule(s)
 * on those cells only (Excel applies one rule per cell; A1:A3 keeps A1 and
 * A3 when A2 gets a new rule) and goes first so it wins first-match lookup.
 * With `sameAs` ("Apply these changes to all other cells with the same
 * settings": the rule the dialog opened on) every rule carrying its settings
 * takes the new ones on its own cells, rebuilt with `build`. "unchanged" when removing finds no rule; null
 * when a remainder cannot be represented.
 */
export function validationsAfterDefine(
  validations: readonly ValidationRule[],
  range: ValidationRange,
  plan: ValidationDefinePlan,
  build: (range: ValidationRange) => ValidationRule | null,
  sameAs: ValidationRule | null,
  shift: ValidationFormulaShifter,
): ValidationRule[] | "unchanged" | null {
  const same = sameAs ? validations.filter(sameValidationSettingsAs(sameAs)) : [];
  if (plan.removing && same.length === 0 && !rangeHasRule(validations, range)) {
    return "unchanged";
  }
  const retargeted = plan.removing
    ? []
    : same.flatMap((candidate) => {
        const next = build({ ...candidate.range });
        // A rule over several areas takes the new settings on every area,
        
        if (next && candidate.areas && candidate.areas.length > 1) {
          next.areas = candidate.areas.map((area) => ({ ...area }));
        }
        return next ? [next] : [];
      });
  const covered = same.some((candidate) =>
    validationRuleAreas(candidate).some(
      (area) => area.r1 <= range.r1 && area.c1 <= range.c1 && area.r2 >= range.r2 && area.c2 >= range.c2,
    ),
  );
  const kept = clearRulesInRange(
    validations.filter((candidate) => !same.includes(candidate)),
    range,
    shift,
  );
  if (!kept) return null;
  return [...(plan.rule && !covered ? [plan.rule] : []), ...retargeted, ...kept];
}

/**
 * Clear All on a selection: only the selected cells lose their rule; the rest
 * of it stays. Null when no rule covers the range or a remainder cannot be
 * represented.
 */
export function validationsAfterClear(
  validations: readonly ValidationRule[],
  range: ValidationRange,
  shift: ValidationFormulaShifter,
): ValidationRule[] | null {
  if (!rangeHasRule(validations, range)) return null;
  return clearRulesInRange(validations, range, shift);
}

/**
 * The controlling cell of a cascading list as the dialog shows it: the row
 * as the grid numbers it (from 1) and the column by its letters, so 3 / C is
 * C3 (live audit S-192: the fields took 0-based numbers, and 2 / 2 read C3).
 * The rule keeps its 0-based coordinates.
 */
export function dependentSourceDraft(source: { row: number; col: number }): { row: string; col: string } {
  if (!(source.row >= 0) || !(source.col >= 0)) return { row: "", col: "" };
  return { row: String(source.row + 1), col: columnLettersOf(source.col) };
}

/** The 0-based cell of a row / column draft (a column by letters or by its
 * number from 1), or null when either is missing or not a cell. */
export function parseDependentSource(rowDraft: string, colDraft: string): { row: number; col: number } | null {
  const rowText = rowDraft.trim();
  const colText = colDraft.trim();
  if (!/^\d+$/.test(rowText) || Number(rowText) < 1) return null;
  let col: number;
  if (/^[A-Za-z]{1,3}$/.test(colText)) {
    col = 0;
    for (const letter of colText.toUpperCase()) col = col * 26 + (letter.charCodeAt(0) - 64);
    col -= 1;
  } else if (/^\d+$/.test(colText) && Number(colText) >= 1) {
    col = Number(colText) - 1;
  } else {
    return null;
  }
  return { row: Number(rowText) - 1, col };
}

/** The A1 address the dialog's row / column drafts name, or null. */
export function dependentSourceAddress(rowDraft: string, colDraft: string): string | null {
  const cell = parseDependentSource(rowDraft, colDraft);
  return cell ? `${columnLettersOf(cell.col)}${cell.row + 1}` : null;
}

function columnLettersOf(index: number): string {
  let letters = "";
  for (let rest = index + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  }
  return letters;
}
