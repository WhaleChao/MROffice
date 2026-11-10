/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A worksheet Form control (Developer > Insert > Form Controls) as the sheet
 * draws it. The API reads the control's properties part (`xl/ctrlProps/*`,
 * x14 `formControlPr`) into `object.formControl`; the control itself stays
 * preserve-only (`nativeControlBlocker`), so this module only answers what it
 * shows: a List Box's items come from its input range (`fmlaRange`) and its
 * selection, a Check Box's state, from its cell link (`fmlaLink`), live, as
 * Excel redraws the control when those cells change.
 */
import { parseSourceRange } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";

export type SpreadsheetFormControlType =
  | "Button"
  | "CheckBox"
  | "Drop"
  | "GBox"
  | "Label"
  | "List"
  | "Radio"
  | "Scroll"
  | "Spin"
  | "EditBox"
  | "Dialog";

export interface SpreadsheetFormControl {
  objectType: SpreadsheetFormControlType;
  text?: string;
  textSizePt?: number;
  textBold?: boolean;
  textItalic?: boolean;
  /** The label font's underline (`a:rPr@u` sng / dbl). */
  textUnderline?: "single" | "double";
  /** The label font's strikethrough (`a:rPr@strike`). */
  textStrike?: boolean;
  textFace?: string;
  textColor?: string;
  textAlign?: "left" | "center" | "right";
  checked?: "checked" | "unchecked" | "mixed";
  /** x14 `noThreeD` absent: drawn with 3-D shading. */
  threeD?: boolean;
  fmlaRange?: string;
  fmlaLink?: string;
  /** Literal items (`itemLst`) of a list without an input range. */
  items?: readonly string[];
  /** 1-based selected item (`sel`); 0 = none. */
  selected?: number;
  value?: number;
  min?: number;
  max?: number;
  increment?: number;
  pageChange?: number;
  horizontal?: boolean;
  /** `dx`: the scroll bar's width in pixels. */
  scrollBarWidth?: number;
  dropLines?: number;
}

const TYPES = new Set<string>([
  "Button", "CheckBox", "Drop", "GBox", "Label", "List", "Radio", "Scroll", "Spin", "EditBox", "Dialog",
]);

/** The object's control reading, or null for any other drawing. */
export function spreadsheetFormControlOf(object: { formControl?: unknown }): SpreadsheetFormControl | null {
  const control = object.formControl;
  if (!control || typeof control !== "object") return null;
  const type = (control as { objectType?: unknown }).objectType;
  return typeof type === "string" && TYPES.has(type) ? (control as SpreadsheetFormControl) : null;
}

/** Text of one cell as the control shows it (`sheet` absent = this sheet). */
export type FormControlCellText = (row: number, col: number, sheet?: string) => string | null;

/** Most items a list draws (it shows a window of them; the rest scroll). */
const MAX_ITEMS = 500;

/** The items of a List Box / Combo Box: every cell of its input range, in
 *  order, blanks included (Excel lists an empty cell as an empty row). */
export function formControlItems(
  control: SpreadsheetFormControl,
  cellText: FormControlCellText,
): string[] {
  if (control.fmlaRange) {
    const range = parseSourceRange(control.fmlaRange);
    if (!range) return [];
    const out: string[] = [];
    for (let r = range.r1; r <= range.r2; r += 1) {
      for (let c = range.c1; c <= range.c2; c += 1) {
        out.push(cellText(r, c, range.sheet) ?? "");
        if (out.length >= MAX_ITEMS) return out;
      }
    }
    return out;
  }
  return [...(control.items ?? [])].slice(0, MAX_ITEMS);
}

/** The linked cell's text, or null when the control has no readable link. */
function linkText(control: SpreadsheetFormControl, cellText: FormControlCellText): string | null {
  if (!control.fmlaLink) return null;
  const range = parseSourceRange(control.fmlaLink);
  if (!range) return null;
  return cellText(range.r1, range.c1, range.sheet);
}

/** The 0-based item a list shows selected (-1: none): its cell link's
 *  number when it has one, else the file's `sel`. */
export function formControlSelectedIndex(
  control: SpreadsheetFormControl,
  cellText: FormControlCellText,
): number {
  const linked = linkText(control, cellText);
  if (linked !== null && linked.trim() !== "") {
    const n = Number(linked);
    return Number.isInteger(n) && n > 0 ? n - 1 : -1;
  }
  return (control.selected ?? 0) > 0 ? (control.selected as number) - 1 : -1;
}

/** A Check Box's state: TRUE / FALSE / #N/A in its cell link, else the file's. */
export function formControlCheckState(
  control: SpreadsheetFormControl,
  cellText: FormControlCellText,
): "checked" | "unchecked" | "mixed" {
  if (control.objectType === "CheckBox") {
    const linked = linkText(control, cellText)?.trim().toUpperCase();
    if (linked === "TRUE") return "checked";
    if (linked === "FALSE") return "unchecked";
    if (linked === "#N/A") return "mixed";
  }
  return control.checked ?? "unchecked";
}

/** A Scroll Bar's / Spin Button's current value: its cell link, else `val`. */
export function formControlValue(
  control: SpreadsheetFormControl,
  cellText: FormControlCellText,
): number {
  const linked = Number(linkText(control, cellText) ?? Number.NaN);
  const min = control.min ?? 0;
  const max = Math.max(min, control.max ?? 100);
  const value = Number.isFinite(linked) ? linked : control.value ?? min;
  return Math.min(max, Math.max(min, value));
}
