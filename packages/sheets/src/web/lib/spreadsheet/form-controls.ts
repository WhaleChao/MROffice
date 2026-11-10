/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the IN-CELL FORM CONTROLS beyond the checkbox (Excel / WPS
 * "Insert > Form Controls"): RADIO BUTTON, SPINNER (number stepper) and SLIDER
 * (range). Like the checkbox (see ./checkbox), each control is an ordinary cell
 * flagged with a single `style.<control> = true` boolean; the control's STATE
 * lives in the cell VALUE so the cell still participates in formulas
 * (=IF(A1,…), =SUM(B1:B9)) exactly like a typed value. The grid renders the
 * flagged cell as the matching <input> instead of an editable string and writes
 * the new value back through the normal cell-edit path (so undo / autosave /
 * sheet-protection apply).
 *
 * These functions own only the value <-> control-state mapping plus the flag
 * plumbing; keeping the (Excel-faithful) parsing rules here makes them testable
 * in isolation, mirroring checkbox.ts. The four control flags are mutually
 * exclusive — a cell carries at most one. The React layer enforces that by
 * setting the chosen flag `true` and the other three `false` in one style patch
 * (updateCellStyle strips `false` keys), so turning a cell into one control
 * clears any other control flag.
 */

import { isCheckboxChecked, checkboxValue } from "./checkbox";

// ── RADIO ──────────────────────────────────────────────────────────────────
// A radio cell holds a boolean SELECTED state, stored as the same canonical
// Excel boolean text as the checkbox ("TRUE" / "FALSE") so the two are
// interchangeable in formulas. Selection is lenient on read (reuses the
// checkbox parser) so a radio can be dropped over a column of yes/no or 1/0.

/** Whether a radio cell's stored VALUE counts as SELECTED (lenient, like the checkbox). */
export function isRadioSelected(value: string | undefined): boolean {
  return isCheckboxChecked(value);
}

/** The canonical VALUE to store for a given selected state ("TRUE" / "FALSE"). */
export function radioValue(selected: boolean): string {
  return checkboxValue(selected);
}

/** Whether a cell's style marks it as a radio cell. Safe on any input. */
export function isRadioCell(style: Record<string, unknown> | undefined): boolean {
  return Boolean(style) && (style as Record<string, unknown>).radio === true;
}

/**
 * The value to seed a cell with when TURNING IT INTO a radio. Preserves an
 * existing lenient-boolean value (a 1/yes becomes selected), else starts
 * UNSELECTED (FALSE) — Excel keeps an existing checked state when you add a
 * control.
 */
export function radioSeedValue(existingValue: string | undefined): string {
  const trimmed = (existingValue ?? "").trim();
  if (trimmed === "") return radioValue(false);
  return radioValue(isRadioSelected(existingValue));
}

// ── SPINNER (number stepper) ─────────────────────────────────────────────────
// A spinner cell holds an INTEGER clamped to [min, max] and stepped by `step`.
// The defaults mirror Excel's Spin Button (0..100, step 1). The bounds are
// fixed defaults (not stored per-cell) so the cell stays byte-stable — only the
// integer value is written. A non-numeric / empty existing value seeds at min.

/** Excel-faithful default bounds for the spinner / spin button. */
export const SPINNER_MIN = 0;
export const SPINNER_MAX = 100;
export const SPINNER_STEP = 1;

/** Whether a cell's style marks it as a spinner cell. Safe on any input. */
export function isSpinnerCell(style: Record<string, unknown> | undefined): boolean {
  return Boolean(style) && (style as Record<string, unknown>).spinner === true;
}

/**
 * Normalize a stored value to the spinner's integer in [min, max]. Lenient on
 * read: parses leading numerics, truncates to an integer, and clamps to the
 * bounds; a blank / non-numeric value falls back to `min`. Keeping this pure
 * means the grid and the seed share one rounding/clamping rule.
 */
export function spinnerValueNumber(
  value: string | undefined,
  min: number = SPINNER_MIN,
  max: number = SPINNER_MAX,
): number {
  const lo = Math.min(min, max);
  const hi = Math.max(min, max);
  const n = typeof value === "string" ? Number.parseFloat(value) : NaN;
  if (!Number.isFinite(n)) return lo;
  const truncated = Math.trunc(n);
  return Math.min(hi, Math.max(lo, truncated));
}

/** The clamped integer as the canonical cell text (what gets written back). */
export function spinnerValue(
  value: string | undefined,
  min: number = SPINNER_MIN,
  max: number = SPINNER_MAX,
): string {
  return String(spinnerValueNumber(value, min, max));
}

/**
 * The next spinner value after stepping by `delta` steps (e.g. +1 on the up
 * arrow, -1 on the down arrow), clamped to the bounds. Returns the canonical
 * cell text.
 */
export function stepSpinnerValue(
  value: string | undefined,
  delta: number,
  min: number = SPINNER_MIN,
  max: number = SPINNER_MAX,
  step: number = SPINNER_STEP,
): string {
  const current = spinnerValueNumber(value, min, max);
  return spinnerValue(String(current + delta * step), min, max);
}

/** The value to seed a cell with when TURNING IT INTO a spinner. */
export function spinnerSeedValue(
  existingValue: string | undefined,
  min: number = SPINNER_MIN,
  max: number = SPINNER_MAX,
): string {
  return spinnerValue(existingValue, min, max);
}

// ── SLIDER (range) ───────────────────────────────────────────────────────────
// A slider cell holds a NUMBER in [min, max], bound to the cell value. Excel's
// Scroll Bar / form slider defaults to 0..100. Unlike the spinner the value
// need not be an integer, but the <input type=range> in the grid uses step 1
// (whole-number ticks) so dragging produces clean integers; a programmatic /
// imported fractional value is preserved on read and only clamped.

/** Excel-faithful default bounds for the slider / scroll bar. */
export const SLIDER_MIN = 0;
export const SLIDER_MAX = 100;

/** Whether a cell's style marks it as a slider cell. Safe on any input. */
export function isSliderCell(style: Record<string, unknown> | undefined): boolean {
  return Boolean(style) && (style as Record<string, unknown>).slider === true;
}

/**
 * Normalize a stored value to the slider's number in [min, max]. Lenient on
 * read: parses leading numerics and clamps to the bounds (NO truncation — a
 * slider can hold a fractional value); a blank / non-numeric value falls back
 * to `min`.
 */
export function sliderValueNumber(
  value: string | undefined,
  min: number = SLIDER_MIN,
  max: number = SLIDER_MAX,
): number {
  const lo = Math.min(min, max);
  const hi = Math.max(min, max);
  const n = typeof value === "string" ? Number.parseFloat(value) : NaN;
  if (!Number.isFinite(n)) return lo;
  return Math.min(hi, Math.max(lo, n));
}

/** The clamped number as the canonical cell text (what gets written back). */
export function sliderValue(
  value: string | undefined,
  min: number = SLIDER_MIN,
  max: number = SLIDER_MAX,
): string {
  return String(sliderValueNumber(value, min, max));
}

/** The value to seed a cell with when TURNING IT INTO a slider. */
export function sliderSeedValue(
  existingValue: string | undefined,
  min: number = SLIDER_MIN,
  max: number = SLIDER_MAX,
): string {
  return sliderValue(existingValue, min, max);
}
