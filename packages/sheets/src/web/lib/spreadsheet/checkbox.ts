/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the IN-CELL CHECKBOX (Excel 2024 / WPS "Insert > Checkbox").
 * A checkbox cell is an ordinary cell flagged `style.checkbox = true`; its VALUE
 * is the Excel boolean text "TRUE" / "FALSE", so the cell still participates in
 * formulas (=IF(A1, …), =COUNTIF(A1:A9,TRUE)) exactly like a typed boolean. The
 * grid renders the flagged cell as a checkbox instead of an editable string and
 * toggles the value on click / Space.
 *
 * These functions own only the value <-> checked mapping and the flag plumbing;
 * the React layer writes the toggled value + flag through the normal cell-edit
 * path so undo / autosave / sheet-protection apply. Keeping the logic here makes
 * the (Excel-faithful) parsing rules testable in isolation.
 */

/** Excel's canonical boolean literals. */
export const CHECKBOX_TRUE = "TRUE";
export const CHECKBOX_FALSE = "FALSE";

/**
 * Whether a cell's stored VALUE counts as CHECKED. Excel-faithful and lenient on
 * read: "TRUE" (any case), "1", and a few common truthy spellings ("yes"/"y"/
 * "✓"/"x"/"on") are checked; everything else — including "", "FALSE", "0", "no"
 * — is unchecked. Surrounding whitespace is ignored. Lenient parsing matters
 * because a checkbox can be placed over a column that already holds yes/no or
 * 1/0 data, and the box should reflect that without rewriting every cell first.
 */
/**
 * Whether a cell's COMPUTED value counts as checked: a formula's boolean or
 * number result decides (=H15 follows H15; live audit B168), a text result
 * reads as the stored text does.
 */
export function isCheckboxCheckedValue(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) && value !== 0;
  if (typeof value === "string") return isCheckboxChecked(value);
  return false;
}

export function isCheckboxChecked(value: string | undefined): boolean {
  if (typeof value !== "string") return false;
  const v = value.trim().toLowerCase();
  return (
    v === "true" ||
    v === "1" ||
    v === "yes" ||
    v === "y" ||
    v === "on" ||
    v === "✓" ||
    v === "✔" ||
    v === "x"
  );
}

/**
 * The VALUE to store for a given checked state — always the canonical Excel
 * boolean text, so toggling normalizes a lenient input ("yes", "1") to
 * "TRUE"/"FALSE". This keeps formulas that reference the cell unambiguous.
 */
export function checkboxValue(checked: boolean): string {
  return checked ? CHECKBOX_TRUE : CHECKBOX_FALSE;
}

/** The value AFTER toggling a cell's current value (checked -> FALSE, else TRUE). */
export function toggleCheckboxValue(value: string | undefined): string {
  return checkboxValue(!isCheckboxChecked(value));
}

/**
 * Whether a cell's style marks it as a checkbox cell. Safe on any input
 * (missing style, non-object) — only a literal `checkbox === true` qualifies.
 */
export function isCheckboxCell(
  style: Record<string, unknown> | undefined,
): boolean {
  return Boolean(style) && (style as Record<string, unknown>).checkbox === true;
}

/**
 * The value to seed a cell with when TURNING IT INTO a checkbox. If the cell
 * already holds a value that parses as a boolean (lenient), preserve that
 * checked state (normalized to TRUE/FALSE); otherwise start unchecked (FALSE).
 * This mirrors Excel, which keeps an existing TRUE/1 when you add a checkbox.
 */
export function checkboxSeedValue(existingValue: string | undefined): string {
  const trimmed = (existingValue ?? "").trim();
  if (trimmed === "") return CHECKBOX_FALSE;
  return checkboxValue(isCheckboxChecked(existingValue));
}

/**
 * Whether Insert ▸ Checkbox may turn this cell into a checkbox at all.
 * Microsoft Support, "Using check boxes in Excel": "A check box is composed of
 * the values TRUE and FALSE with checkbox formatting" — so Excel only places
 * one where that is all the cell would lose, i.e. an EMPTY cell or one that
 * already holds TRUE or FALSE. A cell holding anything else keeps its content
 * and gets no checkbox; inserting over a range of labels and totals must not
 * wipe them out.
 *
 * Deliberately stricter than `isCheckboxChecked`, which is lenient on READ so a
 * box placed over a legacy 1/yes column still shows the right state.
 */
export function canBecomeCheckboxCell(
  existing:
    | string
    | {
        value?: unknown;
        formula?: unknown;
        stringValue?: unknown;
        style?: { quotePrefix?: unknown; numberFormat?: unknown } | null;
      }
    | undefined,
): boolean {
  if (existing === undefined) return true;
  const cell = typeof existing === "string" ? { value: existing } : existing;
  if (typeof cell.formula === "string" && cell.formula.trim() !== "") return false;
  const value = typeof cell.value === "string" ? cell.value : cell.value == null ? "" : String(cell.value);
  // Only an EMPTY cell: spaces, a tab or a no-break space are characters
  // someone typed, and a box would replace them with FALSE (live audit B458).
  if (value === "") return true;
  // A cell the workbook or the author made TEXT ('TRUE, a string TRUE from the
  // file, the Text format) is not the logical TRUE a box can stand for, and
  // neither is TRUE with spaces around it.
  const textIdentity =
    cell.stringValue === true ||
    cell.style?.quotePrefix === true ||
    cell.style?.numberFormat === "@";
  if (textIdentity) return false;
  const upper = value.toUpperCase();
  return upper === CHECKBOX_TRUE || upper === CHECKBOX_FALSE;
}
