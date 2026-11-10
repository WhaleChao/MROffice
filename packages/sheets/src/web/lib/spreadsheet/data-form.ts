/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The record-at-a-time data form (Excel / WPS 数据 > 记录单).
 *
 * A list with a header row is read as records: one form shows a single row's
 * fields, and the user steps through, edits, adds, deletes, or searches by
 * criteria without scrolling a wide sheet sideways. Everything here is pure and
 * works on plain cell maps, so the dialog stays a thin renderer and every rule
 * below is unit-testable on its own.
 *
 * A formula field is shown but never written back: the form is for data entry,
 * and silently replacing `=B2*C2` with the number the user happened to see is
 * how a spreadsheet loses its arithmetic.
 */

import { spreadsheetCellIsText } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";

export interface DataFormRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface DataFormField {
  /** Header text, or the column letter when the header cell is blank. */
  label: string;
  /** Absolute column index in the sheet. */
  col: number;
  /** A formula field is read-only in the form. */
  readOnly: boolean;
  /** The entry as the user would edit it: a quote-prefixed text keeps its
   *  apostrophe, a formula its source. */
  value: string;
  /** A formula field's calculated result as text, when the reader supplied a
   *  calculator; criteria compare against it (live audit B281). */
  computed?: string;
}

export interface DataFormReadOptions {
  /** The result of the cell at (row, col), for formula fields. */
  computed?: (row: number, col: number) => unknown;
}

export interface DataFormRecord {
  /** Absolute row index in the sheet. */
  row: number;
  /** Zero-based position within the list's body. */
  index: number;
  fields: DataFormField[];
}

export interface DataFormShape {
  range: DataFormRange;
  headers: string[];
  /** Body row count (the header row excluded). */
  recordCount: number;
}

type CellMap = Record<
  string,
  | {
      value?: unknown;
      formula?: string;
      kind?: string;
      style?: { numberFormat?: unknown; quotePrefix?: unknown };
    }
  | undefined
>;

const MAX_FORM_FIELDS = 32;

function cellText(cells: CellMap, row: number, col: number): string {
  const raw = cells[`${row}:${col}`]?.value;
  return raw === undefined || raw === null ? "" : String(raw);
}

function isFormula(value: string): boolean {
  return value.startsWith("=");
}

/**
 * A field as the form shows and edits it. A text entered after an apostrophe
 * reads as =1+1 in the cell but is text, not a formula: it shows with its
 * apostrophe, is editable, and saves as the same kind of entry the grid would
 * take (live audit B281).
 */
function fieldEntryText(cells: CellMap, row: number, col: number): string {
  const cell = cells[`${row}:${col}`];
  const text = cellText(cells, row, col);
  const identity = cell
    ? { value: text, formula: cell.formula, kind: cell.kind, style: cell.style }
    : undefined;
  if (text !== "" && spreadsheetCellIsText(identity) && cell?.style?.quotePrefix === true) {
    return `'${text}`;
  }
  return text;
}

/** Column letter for a zero-based index, so a blank header still names a field. */
export function columnLetter(col: number): string {
  let remaining = col;
  let out = "";
  while (remaining >= 0) {
    out = String.fromCharCode(65 + (remaining % 26)) + out;
    remaining = Math.floor(remaining / 26) - 1;
  }
  return out;
}

/**
 * Read the list's shape: its header row and how many records sit under it.
 *
 * Returns null when the range cannot be a list — no body rows, or more than
 * Excel's 32-field Data Form limit. Enforce the limit in the model so no caller
 * can mount a form that Excel itself could not represent.
 */
export function readDataFormShape(
  cells: CellMap,
  range: DataFormRange,
): DataFormShape | null {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  const fieldCount = c2 - c1 + 1;
  if (r2 <= r1 || fieldCount < 1 || fieldCount > MAX_FORM_FIELDS) return null;
  const headers: string[] = [];
  for (let col = c1; col <= c2; col += 1) {
    const header = cellText(cells, r1, col).trim();
    headers.push(header || columnLetter(col));
  }
  return {
    range: { r1, c1, r2, c2 },
    headers,
    recordCount: r2 - r1,
  };
}

/** The record at a body position, or null when the position is outside the list. */
export function readDataFormRecord(
  cells: CellMap,
  shape: DataFormShape,
  index: number,
  options: DataFormReadOptions = {},
): DataFormRecord | null {
  if (index < 0 || index >= shape.recordCount) return null;
  const row = shape.range.r1 + 1 + index;
  const fields: DataFormField[] = shape.headers.map((label, offset) => {
    const col = shape.range.c1 + offset;
    const value = fieldEntryText(cells, row, col);
    const readOnly = isFormula(value);
    const field: DataFormField = { label, col, readOnly, value };
    if (readOnly && options.computed) {
      const result = options.computed(row, col);
      field.computed = result === undefined || result === null ? "" : String(result);
    }
    return field;
  });
  return { row, index, fields };
}

export interface DataFormWrite {
  key: string;
  value: string;
}

/**
 * The cell writes that save an edited record.
 *
 * Only fields that actually changed are written, and a read-only (formula)
 * field never is, so saving a record the user only looked at touches nothing.
 */
export function dataFormWrites(
  record: DataFormRecord,
  edited: readonly string[],
): DataFormWrite[] {
  const writes: DataFormWrite[] = [];
  record.fields.forEach((field, offset) => {
    if (field.readOnly) return;
    const next = edited[offset];
    if (next === undefined || next === field.value) return;
    writes.push({ key: `${record.row}:${field.col}`, value: next });
  });
  return writes;
}

/**
 * Where a new record goes: the row directly under the last one, so the list
 * grows downward the way Excel's form does.
 */
export function newRecordRow(shape: DataFormShape): number {
  return shape.range.r2 + 1;
}

/**
 * Delete one record from a form-bound range without deleting the worksheet row.
 * Complete cell objects move up inside the form's columns through the final
 * record; cells outside that bounded rectangle (including data directly below
 * the list) remain byte-for-byte where they were.
 */
export function deleteDataFormRecordCells<T>(
  cells: Readonly<Record<string, T>>,
  shape: DataFormShape,
  index: number,
): Record<string, T> | null {
  if (index < 0 || index >= shape.recordCount) return null;
  const deletedRow = shape.range.r1 + 1 + index;
  const out: Record<string, T> = {};
  for (const [key, cell] of Object.entries(cells)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    const inFormColumns =
      Number.isInteger(col) && col >= shape.range.c1 && col <= shape.range.c2;
    if (
      Number.isInteger(row) &&
      inFormColumns &&
      row >= deletedRow &&
      row <= shape.range.r2
    ) {
      if (row !== deletedRow) out[`${row - 1}:${col}`] = cell;
      continue;
    }
    out[key] = cell;
  }
  return out;
}

function escapeCriteriaCharacter(character: string): string {
  return character.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Compile Excel Data Form text criteria. A plain value is a prefix search;
 * once the user supplies wildcard syntax, the whole field must match. `~`
 * quotes the next `*`, `?`, or `~`, matching Excel's documented escape rule.
 */
function dataFormTextCriterion(
  raw: string,
  plainTextIsPrefix: boolean,
): RegExp {
  let source = "";
  let usesWildcardSyntax = false;
  for (let index = 0; index < raw.length; index += 1) {
    const character = raw[index];
    const escaped = raw[index + 1];
    if (
      character === "~" &&
      (escaped === "*" || escaped === "?" || escaped === "~")
    ) {
      usesWildcardSyntax = true;
      source += escapeCriteriaCharacter(escaped);
      index += 1;
    } else if (character === "*") {
      usesWildcardSyntax = true;
      source += ".*";
    } else if (character === "?") {
      usesWildcardSyntax = true;
      source += ".";
    } else {
      source += escapeCriteriaCharacter(character);
    }
  }
  const suffix = usesWildcardSyntax || !plainTextIsPrefix ? "$" : ".*$";
  return new RegExp(`^${source}${suffix}`, "i");
}

/**
 * Match a record against the form's criteria the way Excel's "Criteria" mode
 * does: every non-empty plain-text criterion matches the beginning of its field,
 * case-insensitively; `*` and `?` are wildcards and `~` escapes them. An
 * all-blank criteria set matches everything.
 *
 * A leading comparison operator on a numeric criterion (`>100`, `<=3`) compares
 * numerically, which is what people type into that box and what Excel honours.
 */
export function matchesDataFormCriteria(
  record: DataFormRecord,
  criteria: readonly string[],
): boolean {
  return record.fields.every((field, offset) => {
    const criterion = (criteria[offset] ?? "").trim();
    if (!criterion) return true;
    // A formula field is matched by what it shows, not by its source text.
    const fieldValue = field.computed ?? field.value;
    const comparison = /^(>=|<=|<>|>|<|=)\s*(.+)$/.exec(criterion);
    if (comparison) {
      const [, operator, rawOperand] = comparison;
      const operand = Number(rawOperand);
      const value = Number(fieldValue);
      if (Number.isFinite(operand) && Number.isFinite(value)) {
        switch (operator) {
          case ">":
            return value > operand;
          case "<":
            return value < operand;
          case ">=":
            return value >= operand;
          case "<=":
            return value <= operand;
          case "<>":
            return value !== operand;
          default:
            return value === operand;
        }
      }
      // A non-numeric comparison falls back to exact text for "=" / "<>" and
      // otherwise cannot be answered, so it does not match.
      const target = rawOperand.trim();
      const matches = dataFormTextCriterion(target, false).test(fieldValue);
      if (operator === "=") return matches;
      if (operator === "<>") return !matches;
      return false;
    }
    return dataFormTextCriterion(criterion, true).test(fieldValue);
  });
}

/**
 * The next record index matching the criteria, searching from `from` in
 * `direction`, or null when nothing matches. The search does not wrap: Excel's
 * form stops at the ends, which is what tells the user there are no more.
 */
export function findDataFormRecord(
  cells: CellMap,
  shape: DataFormShape,
  criteria: readonly string[],
  from: number,
  direction: 1 | -1,
  options: DataFormReadOptions = {},
): number | null {
  for (
    let index = from + direction;
    index >= 0 && index < shape.recordCount;
    index += direction
  ) {
    const record = readDataFormRecord(cells, shape, index, options);
    if (record && matchesDataFormCriteria(record, criteria)) return index;
  }
  return null;
}
