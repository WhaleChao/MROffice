/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * XLS-A030 — View ▸ Workbook Views ▸ Custom Views.
 *
 * Microsoft Support, "Create, apply, or delete a custom view": a custom view
 * keeps the column widths, display options, frozen panes / window splits, the
 * active sheet and the selected cells as they were when it was added, and,
 * optionally, the print settings and the hidden rows, columns and filter
 * settings ("Include in view" check boxes of the Add View dialog). Showing it
 * puts all of that back. Excel greys Custom Views while the workbook contains
 * a table.
 *
 * The views live at the workbook top level (`customViews`). A view stores,
 * per sheet, only the fields it owns; showing it rewrites exactly those
 * fields on the sheets that still exist.
 */

export const MAX_CUSTOM_VIEWS = 64;
const MAX_CUSTOM_VIEW_NAME = 255;

/** Always part of a view: the sheet's display. Gridlines and headings are
 *  each sheet's own, as in Excel's customSheetView. */
export const CUSTOM_VIEW_DISPLAY_KEYS = [
  "freezeRows",
  "freezeCols",
  "splitRow",
  "splitCol",
  "zoomScale",
  "columnWidths",
  "hideGridlines",
  "hideHeadings",
] as const;
/** "Hidden rows, columns and filter settings". */
export const CUSTOM_VIEW_HIDDEN_KEYS = [
  "manuallyHiddenRows",
  "manuallyHiddenCols",
  "filters",
  "colorFilters",
  "conditionFilters",
  "autoFilterRef",
  "autoFilterId",
] as const;
/** "Print settings". */
export const CUSTOM_VIEW_PRINT_KEYS = ["printSettings"] as const;

type CustomViewSheetKey =
  | (typeof CUSTOM_VIEW_DISPLAY_KEYS)[number]
  | (typeof CUSTOM_VIEW_HIDDEN_KEYS)[number]
  | (typeof CUSTOM_VIEW_PRINT_KEYS)[number];

export interface SpreadsheetCustomViewSelection {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  row: number;
  col: number;
}

export interface SpreadsheetCustomViewSheet {
  sheetId: string;
  /** The owned fields as they were (an owned key that is absent was unset). */
  fields: Partial<Record<CustomViewSheetKey, unknown>>;
}

export interface SpreadsheetCustomView {
  id: string;
  name: string;
  printSettings: boolean;
  hiddenAndFilters: boolean;
  activeSheetId: string;
  selection?: SpreadsheetCustomViewSelection;
  /** Display options a view keeps (View ▸ Show): the formula bar's, and the
   *  active sheet's gridlines and headings. A view added when those two were
   *  one switch for the workbook keeps them only here, for every sheet. */
  display: { hideGridlines: boolean; hideHeadings: boolean; hideFormulaBar: boolean };
  sheets: SpreadsheetCustomViewSheet[];
}

export function customViewNameKey(name: string): string {
  return name.trim().toLocaleLowerCase();
}

function ownedKeys(view: Pick<SpreadsheetCustomView, "printSettings" | "hiddenAndFilters">): CustomViewSheetKey[] {
  return [
    ...CUSTOM_VIEW_DISPLAY_KEYS,
    ...(view.hiddenAndFilters ? CUSTOM_VIEW_HIDDEN_KEYS : []),
    ...(view.printSettings ? CUSTOM_VIEW_PRINT_KEYS : []),
  ];
}

function jsonCopy<T>(value: T): T {
  return value === undefined ? value : (JSON.parse(JSON.stringify(value)) as T);
}

/** Add View: the workbook as it is now. */
export function captureCustomView(input: {
  id: string;
  name: string;
  printSettings: boolean;
  hiddenAndFilters: boolean;
  sheets: readonly ({ id: string } & Record<string, unknown>)[];
  activeSheetId: string;
  selection?: SpreadsheetCustomViewSelection;
  display: SpreadsheetCustomView["display"];
}): SpreadsheetCustomView | null {
  const name = input.name.trim();
  if (!name || name.length > MAX_CUSTOM_VIEW_NAME) return null;
  const keys = ownedKeys(input);
  return {
    id: input.id,
    name,
    printSettings: input.printSettings,
    hiddenAndFilters: input.hiddenAndFilters,
    activeSheetId: input.activeSheetId,
    ...(input.selection ? { selection: { ...input.selection } } : {}),
    display: { ...input.display },
    sheets: input.sheets.map((sheet) => {
      const fields: Partial<Record<CustomViewSheetKey, unknown>> = {};
      for (const key of keys) {
        if (sheet[key] !== undefined) fields[key] = jsonCopy(sheet[key]);
      }
      return { sheetId: sheet.id, fields };
    }),
  };
}

/** Show: the sheets with the view's fields put back (same array when no change). */
export function sheetsWithCustomView<S extends { id: string }>(
  sheets: readonly S[],
  view: SpreadsheetCustomView,
): S[] {
  const keys = ownedKeys(view);
  const byId = new Map(view.sheets.map((entry) => [entry.sheetId, entry]));
  // A view added while gridlines and headings were one workbook switch
  // holds them in `display` alone: they go on every sheet it keeps.
  const workbookWide = (key: "hideGridlines" | "hideHeadings") =>
    view.display[key] === true && view.sheets.every((entry) => entry.fields[key] === undefined);
  let changed = false;
  const next = sheets.map((sheet) => {
    const saved = byId.get(sheet.id);
    if (!saved) return sheet;
    const out = { ...sheet } as Record<string, unknown>;
    for (const key of keys) {
      const value =
        (key === "hideGridlines" || key === "hideHeadings") && workbookWide(key)
          ? true
          : saved.fields[key];
      if (value === undefined) delete out[key];
      else out[key] = jsonCopy(value);
    }
    if (JSON.stringify(out) === JSON.stringify(sheet)) return sheet;
    changed = true;
    return out as S;
  });
  return changed ? next : (sheets as S[]);
}

function isIndex(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
}

/** Fail-closed reader for the workbook's `customViews`. */
export function normalizeCustomViews(raw: unknown): SpreadsheetCustomView[] {
  if (!Array.isArray(raw)) return [];
  const out: SpreadsheetCustomView[] = [];
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_CUSTOM_VIEWS) break;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const candidate = entry as Record<string, unknown>;
    const name = typeof candidate.name === "string" ? candidate.name.trim() : "";
    if (
      typeof candidate.id !== "string" ||
      !/^view-[a-z0-9-]{1,60}$/.test(candidate.id) ||
      ids.has(candidate.id) ||
      !name ||
      name.length > MAX_CUSTOM_VIEW_NAME ||
      names.has(customViewNameKey(name)) ||
      typeof candidate.printSettings !== "boolean" ||
      typeof candidate.hiddenAndFilters !== "boolean" ||
      typeof candidate.activeSheetId !== "string" ||
      !Array.isArray(candidate.sheets)
    ) {
      continue;
    }
    const display = candidate.display as Record<string, unknown> | undefined;
    if (
      !display ||
      typeof display.hideGridlines !== "boolean" ||
      typeof display.hideHeadings !== "boolean" ||
      typeof display.hideFormulaBar !== "boolean"
    ) {
      continue;
    }
    const selection = candidate.selection as Record<string, unknown> | undefined;
    if (
      selection !== undefined &&
      !(["r1", "c1", "r2", "c2", "row", "col"] as const).every((key) => isIndex(selection?.[key]))
    ) {
      continue;
    }
    const allowed = new Set<string>(ownedKeys(candidate as Pick<SpreadsheetCustomView, "printSettings" | "hiddenAndFilters">));
    const sheets: SpreadsheetCustomViewSheet[] = [];
    let valid = true;
    for (const rawSheet of candidate.sheets as unknown[]) {
      const sheet = rawSheet as Record<string, unknown> | null;
      if (
        !sheet ||
        typeof sheet !== "object" ||
        typeof sheet.sheetId !== "string" ||
        !sheet.fields ||
        typeof sheet.fields !== "object" ||
        Array.isArray(sheet.fields) ||
        Object.keys(sheet.fields).some((key) => !allowed.has(key))
      ) {
        valid = false;
        break;
      }
      sheets.push({ sheetId: sheet.sheetId, fields: jsonCopy(sheet.fields as SpreadsheetCustomViewSheet["fields"]) });
    }
    if (!valid) continue;
    ids.add(candidate.id);
    names.add(customViewNameKey(name));
    out.push({
      id: candidate.id,
      name,
      printSettings: candidate.printSettings,
      hiddenAndFilters: candidate.hiddenAndFilters,
      activeSheetId: candidate.activeSheetId,
      ...(selection ? { selection: selection as unknown as SpreadsheetCustomViewSelection } : {}),
      display: {
        hideGridlines: display.hideGridlines,
        hideHeadings: display.hideHeadings,
        hideFormulaBar: display.hideFormulaBar,
      },
      sheets,
    });
  }
  return out;
}
