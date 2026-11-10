/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Name Manager's list as Excel draws it (XLS-A066): Name / Value / Refers To /
 * Scope / Comment columns, table names beside defined names, and the Filter
 * menu (Clear Filter; Names Scoped to Worksheet / Workbook; Names with /
 * without Errors; Defined Names / Table Names).
 *
 * Microsoft Support, "Use the Name Manager in Excel": Value is "the current
 * value of the name, such as the results of a formula, a string constant, a
 * cell range, an error, an array of values, or a placeholder if the formula
 * cannot be evaluated"; each Filter command toggles its operation on or off,
 * so the operations combine. The two commands of a pair (worksheet/workbook,
 * with/without errors, defined/table) exclude each other.
 */

import {
  isArrayValue,
  isFormulaError,
  isLambdaValue,
  isOmittedValue,
  type CellValue,
  type EngineValue,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { splitSheetQualifier } from "@mrnote/core/web/lib/spreadsheet/name-syntax.ts";

/** Excel's Value placeholder for a name it cannot evaluate. */
export const NAME_MANAGER_UNEVALUATED = "{...}";

/** Cells a Value reads off a range before it stops with an ellipsis. */
export const NAME_MANAGER_VALUE_CELL_LIMIT = 64;

export type NameManagerScopeFilter = "all" | "workbook" | "worksheet";
export type NameManagerErrorFilter = "all" | "with" | "without";
export type NameManagerKindFilter = "all" | "defined" | "table";

export interface NameManagerFilter {
  readonly scope: NameManagerScopeFilter;
  readonly errors: NameManagerErrorFilter;
  readonly kind: NameManagerKindFilter;
}

export const NAME_MANAGER_NO_FILTER: NameManagerFilter = {
  scope: "all",
  errors: "all",
  kind: "all",
};

export type NameManagerFilterCommand =
  | "worksheet"
  | "workbook"
  | "with-errors"
  | "without-errors"
  | "defined"
  | "table";

/** A Filter menu command: on when off, off when on; its pair partner yields. */
export function toggleNameManagerFilter(
  filter: NameManagerFilter,
  command: NameManagerFilterCommand,
): NameManagerFilter {
  switch (command) {
    case "worksheet":
    case "workbook":
      return { ...filter, scope: filter.scope === command ? "all" : command };
    case "with-errors":
      return { ...filter, errors: filter.errors === "with" ? "all" : "with" };
    case "without-errors":
      return {
        ...filter,
        errors: filter.errors === "without" ? "all" : "without",
      };
    case "defined":
    case "table":
      return { ...filter, kind: filter.kind === command ? "all" : command };
  }
}

export function isNameManagerFilterOn(
  filter: NameManagerFilter,
  command: NameManagerFilterCommand,
): boolean {
  switch (command) {
    case "worksheet":
    case "workbook":
      return filter.scope === command;
    case "with-errors":
      return filter.errors === "with";
    case "without-errors":
      return filter.errors === "without";
    case "defined":
    case "table":
      return filter.kind === command;
  }
}

export interface NameManagerRowFacts {
  readonly kind: "defined" | "table";
  readonly scope: "workbook" | "worksheet";
  readonly hasError: boolean;
}

export function nameManagerRowMatches(
  row: NameManagerRowFacts,
  filter: NameManagerFilter,
): boolean {
  if (filter.scope !== "all" && row.scope !== filter.scope) return false;
  if (filter.kind !== "all" && row.kind !== filter.kind) return false;
  if (filter.errors === "with" && !row.hasError) return false;
  if (filter.errors === "without" && row.hasError) return false;
  return true;
}

/** A plain cell or range a name refers to, on the sheet it names. */
export interface NameManagerReferenceRange {
  /** The sheet the reference names; absent for a bare reference. */
  readonly sheetName?: string;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

function columnIndex(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value - 1;
}

/**
 * The cell or range `ref` names (`=Data!$A$1:$B$3`, `'Q1''s'!$C$4`, `B2`), or
 * null when it is a formula, a constant or anything else. A reference to a
 * deleted cell (`#REF!`) is not a range.
 */
export function nameManagerReferenceRange(
  ref: string,
): NameManagerReferenceRange | null {
  const body = ref.trim().replace(/^=\s*/, "");
  if (!body || body.includes("#REF!")) return null;
  const { sheet, ref: address } = splitSheetQualifier(body);
  const parts = address.split(":");
  if (parts.length > 2) return null;
  const from = CELL_RE.exec(parts[0] ?? "");
  const to = CELL_RE.exec(parts[1] ?? parts[0] ?? "");
  if (!from || !to) return null;
  const a = { r: Number(from[2]) - 1, c: columnIndex(from[1]) };
  const b = { r: Number(to[2]) - 1, c: columnIndex(to[1]) };
  if (a.r < 0 || b.r < 0 || a.c < 0 || b.c < 0) return null;
  return {
    ...(sheet !== undefined ? { sheetName: sheet } : {}),
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

/**
 * A constant a name holds (`=0.2`, `="EUR"`, `TRUE`), read without the
 * calculation session; undefined when `ref` is not a plain constant.
 */
export function nameManagerConstantValue(ref: string): CellValue | undefined {
  const body = ref.trim().replace(/^=\s*/, "");
  if (/^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(body)) {
    const value = Number(body);
    return Number.isFinite(value) ? value : undefined;
  }
  const text = /^"((?:[^"]|"")*)"$/.exec(body);
  if (text) return text[1].replace(/""/g, '"');
  if (/^TRUE$/i.test(body)) return true;
  if (/^FALSE$/i.test(body)) return false;
  return undefined;
}

/** Whether a name's value is an error (#REF!, #NAME?, #VALUE!, ...). */
/**
 * Whether a name's reference is broken (#REF! outside any quoted string).
 * "Label #REF! note" is a text constant, not a reference error (B482).
 */
export function nameManagerRefIsBroken(ref: string): boolean {
  return ref.replace(/"(?:[^"]|"")*"/g, "").includes("#REF!");
}

export function nameManagerValueIsError(
  value: EngineValue | undefined,
  ref: string,
): boolean {
  if (nameManagerRefIsBroken(ref)) return true;
  return isFormulaError(value);
}

function scalarText(value: CellValue, inArray: boolean): string {
  if (value === null) return inArray ? '""' : "";
  if (isFormulaError(value)) return value.code;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    return Number.isFinite(value) ? String(Number(value.toPrecision(15))) : "#NUM!";
  }
  // Text inside an array reads as an array constant's text does: quoted.
  return inArray ? `"${value.replace(/"/g, '""')}"` : value;
}

/**
 * The Value column's text: a scalar as itself, an error as its code, a range
 * or array as an array constant (`{"a","b";1,2}`: commas between columns,
 * semicolons between rows), cut short after
 * {@link NAME_MANAGER_VALUE_CELL_LIMIT} cells; Excel's `{...}` when the value
 * could not be evaluated.
 */
export function formatNameManagerValue(
  value: EngineValue | readonly (readonly CellValue[])[] | undefined,
  truncated = false,
): string {
  if (value === undefined || isLambdaValue(value) || isOmittedValue(value)) {
    return NAME_MANAGER_UNEVALUATED;
  }
  const grid = Array.isArray(value)
    ? (value as readonly (readonly CellValue[])[])
    : isArrayValue(value)
      ? value.grid
      : null;
  if (!grid) return scalarText(value as CellValue, false);
  if (grid.length === 1 && grid[0]?.length === 1 && !truncated) {
    return scalarText(grid[0][0] ?? null, false);
  }
  let budget = NAME_MANAGER_VALUE_CELL_LIMIT;
  let cut = truncated;
  const rows: string[] = [];
  for (const row of grid) {
    if (budget <= 0) {
      cut = true;
      break;
    }
    const cells: string[] = [];
    for (const cell of row) {
      if (budget <= 0) {
        cut = true;
        break;
      }
      cells.push(scalarText(cell ?? null, true));
      budget -= 1;
    }
    rows.push(cells.join(","));
  }
  return `{${rows.join(";")}${cut ? "…" : ""}}`;
}
