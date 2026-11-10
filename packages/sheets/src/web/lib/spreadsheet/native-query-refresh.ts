/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Refresh of an imported Power Query query whose sources are this workbook's
 * own tables (Excel.CurrentWorkbook()), as Excel's Data > Refresh does it:
 * the query runs (power-query-m.ts), its result replaces the rows of the table
 * it loads to, matched to the table's columns by name; a result with more or
 * fewer rows grows or shrinks the table, keeping its header, style and column
 * formats (External Data Properties: Preserve cell formatting, the default).
 *
 * The query's own parts (the DataMashup, connections.xml, the queryTable)
 * stay in the source package untouched; only the table's cells and range
 * change, which the package writer carries like any table edit.
 *
 * Fail-closed: a query MRAI cannot run, a result whose columns differ from
 * the table's, or a row-count change that would move cells under the table
 * (Excel's default "insert cells for new data, delete unused cells" shifts
 * them) refuses with a reason instead of writing something Excel would not.
 */

import type { SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { classifySpreadsheetNumberFormat } from "@mrnote/core/web/lib/spreadsheet/date-filter-tree.ts";
import type { EngineValue } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { planDataTableResizeOutcome, type TableButtonSheet } from "./table-buttons";
import {
  MDateTime,
  MTime,
  PowerQueryError,
  PowerQueryUnsupportedError,
  evaluatePowerQuery,
  powerQueryResultToSheet,
  type MValue,
  type PowerQuerySheetValue,
} from "./power-query-m";
import type { SpreadsheetWorkbookQuery } from "./workbook-import-client";

export type NativeQueryRefreshCode =
  | "unsupported"
  | "m_error"
  | "not_loaded"
  | "external"
  | "table_missing"
  | "columns_changed"
  | "content_below"
  | "resize_refused"
  | "source_unreadable"
  | "formula_pending"
  | "protected";

export class NativeQueryRefreshError extends Error {
  constructor(
    readonly code: NativeQueryRefreshCode,
    readonly detail = "",
  ) {
    super(`${code}${detail ? `: ${detail}` : ""}`);
    this.name = "NativeQueryRefreshError";
  }
}

export interface NativeQueryRefreshOptions<S> {
  /** The computed value of a formula cell; undefined while not calculated. */
  formulaValue: (sheet: S, row: number, col: number) => EngineValue | undefined;
  rangeHasLockedCell: (sheet: S, range: { r1: number; r2: number; c1: number; c2: number }) => boolean;
  now?: Date;
}

const ERROR_TEXT = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$/;

/** Whether this query can be refreshed in MRAI at all (before running it). */
export function nativeQueryRefreshable(query: SpreadsheetWorkbookQuery): NativeQueryRefreshCode | null {
  if (query.sourceKind === "external") return "external";
  if (query.load !== "table" || !query.tableName) return "not_loaded";
  return null;
}

function styleOf(cell: SpreadsheetCell | undefined): SpreadsheetCell["style"] | undefined {
  return cell?.style && typeof cell.style === "object" ? cell.style : undefined;
}

function numberValue(sheetCell: SpreadsheetCell | undefined, value: number): MValue {
  const format = styleOf(sheetCell)?.numberFormat;
  if (typeof format === "string" && format) {
    const semantics = classifySpreadsheetNumberFormat(format, value);
    if (semantics.valid && semantics.hasCalendarDate) return new MDateTime(value);
    if (semantics.valid && semantics.hasClock) return new MTime(value - Math.floor(value));
  }
  return value;
}

function engineToM(value: EngineValue | undefined, cell: SpreadsheetCell, where: string): MValue {
  if (value === undefined) throw new NativeQueryRefreshError("formula_pending", where);
  if (value === null) return null;
  if (typeof value === "number") return numberValue(cell, value);
  if (typeof value === "string") return value === "" ? null : value;
  if (typeof value === "boolean") return value;
  if (typeof value === "object" && "__formulaError" in value)
    throw new NativeQueryRefreshError("source_unreadable", `${where} ${(value as { code: string }).code}`);
  throw new NativeQueryRefreshError("source_unreadable", where);
}

/** A cell as Excel.CurrentWorkbook() reads it. */
function cellToM<S extends TableButtonSheet>(sheet: S, row: number, col: number, options: NativeQueryRefreshOptions<S>): MValue {
  const cell = sheet.cells[`${row}:${col}`] as SpreadsheetCell | undefined;
  if (!cell) return null;
  const where = `${sheet.name}!R${row + 1}C${col + 1}`;
  const style = styleOf(cell);
  const quoted = style?.quotePrefix === true;
  const value = typeof cell.value === "string" ? cell.value : String(cell.value ?? "");
  if (cell.formula || (!quoted && value.startsWith("=") && value.length > 1))
    return engineToM(options.formulaValue(sheet, row, col), cell, where);
  if (value === "") return null;
  if (cell.kind === "number") {
    const number = Number(value);
    if (!Number.isFinite(number)) throw new NativeQueryRefreshError("source_unreadable", where);
    return numberValue(cell, number);
  }
  if (cell.kind === "text" || quoted || style?.numberFormat === "@") return value;
  if (value === "TRUE" || value === "FALSE") return value === "TRUE";
  if (ERROR_TEXT.test(value)) throw new NativeQueryRefreshError("source_unreadable", `${where} ${value}`);
  return value;
}

function headerText(cell: SpreadsheetCell | undefined): string {
  return typeof cell?.value === "string" ? cell.value : String(cell?.value ?? "");
}

/** Each table of the workbook as Excel.CurrentWorkbook() lists it. */
export function nativeQueryWorkbookItems<S extends TableButtonSheet>(sheets: readonly S[], options: NativeQueryRefreshOptions<S>, wanted?: ReadonlySet<string>) {
  const items: Array<{ name: string; columns: string[]; rows: MValue[][] }> = [];
  for (const sheet of sheets) {
    for (const table of sheet.dataTables ?? []) {
      if (wanted && !wanted.has(table.name.toLowerCase())) continue;
      const { r1, c1, r2, c2 } = table.range;
      const columns = table.headerRow
        ? Array.from({ length: c2 - c1 + 1 }, (_, offset) => headerText(sheet.cells[`${r1}:${c1 + offset}`] as SpreadsheetCell | undefined) || table.columns?.[offset]?.name || `Column${offset + 1}`)
        : Array.from({ length: c2 - c1 + 1 }, (_, offset) => table.columns?.[offset]?.name ?? `Column${offset + 1}`);
      const first = table.headerRow ? r1 + 1 : r1;
      if ((r2 - first + 1) * columns.length > 1_000_000) throw new NativeQueryRefreshError("unsupported", `table ${table.name} over a million cells`);
      const rows: MValue[][] = [];
      for (let row = first; row <= r2; row += 1)
        rows.push(columns.map((_, offset) => cellToM(sheet, row, c1 + offset, options)));
      items.push({ name: table.name, columns, rows });
    }
  }
  return items;
}

const DATE_FORMATS: Record<"date" | "datetime" | "time", string> = {
  date: "m/d/yyyy",
  datetime: "m/d/yyyy h:mm",
  time: "h:mm:ss",
};

function writtenCell(value: PowerQuerySheetValue, base: SpreadsheetCell | undefined): SpreadsheetCell | null {
  const style = { ...(styleOf(base) ?? {}) } as NonNullable<SpreadsheetCell["style"]>;
  delete style.quotePrefix;
  const finish = (cell: SpreadsheetCell): SpreadsheetCell => {
    if (Object.keys(style).length) cell.style = style;
    return cell;
  };
  if (value === null) return Object.keys(style).length ? finish({ value: "" }) : null;
  if (value.kind === "number") {
    if (value.format) {
      const existing = typeof style.numberFormat === "string" ? style.numberFormat : "";
      const semantics = existing ? classifySpreadsheetNumberFormat(existing, value.value) : null;
      const fits = value.format === "time" ? semantics?.hasClock && !semantics.hasCalendarDate : semantics?.hasCalendarDate;
      if (!fits) style.numberFormat = DATE_FORMATS[value.format];
    }
    return finish({ value: String(value.value), kind: "number" });
  }
  if (value.kind === "logical") return finish({ value: value.value ? "TRUE" : "FALSE" });
  // A text that reads as a formula or a logical keeps its text identity.
  if (value.value.trim().startsWith("=") || value.value === "TRUE" || value.value === "FALSE") style.quotePrefix = true;
  return finish({ value: value.value, kind: "text" });
}

export interface NativeQueryRefreshPlan<S> {
  sheets: S[];
  sheetId: string;
  tableId: string;
  rowCount: number;
  range: { r1: number; c1: number; r2: number; c2: number };
}

/**
 * Run `name` and lay its result into the table it loads to. Returns the new
 * workbook; throws NativeQueryRefreshError with the reason it refuses.
 */
export function planNativeQueryRefresh<S extends TableButtonSheet & { protected?: boolean }>(
  sheets: readonly S[],
  queries: readonly SpreadsheetWorkbookQuery[],
  name: string,
  options: NativeQueryRefreshOptions<S>,
): NativeQueryRefreshPlan<S> {
  const query = queries.find((entry) => entry.name === name);
  if (!query) throw new NativeQueryRefreshError("not_loaded", name);
  const blocked = nativeQueryRefreshable(query);
  if (blocked === "external") throw new NativeQueryRefreshError("external", query.externalFunctions.join(", "));
  if (blocked) throw new NativeQueryRefreshError(blocked, name);
  const targetSheetIndex = sheets.findIndex((sheet) => sheet.name === query.sheetName);
  const target = targetSheetIndex >= 0 ? sheets[targetSheetIndex] : undefined;
  const table = target?.dataTables?.find((entry) => entry.name.toLowerCase() === query.tableName!.toLowerCase());
  if (!target || !table) throw new NativeQueryRefreshError("table_missing", query.tableName ?? "");
  if (target.protected) throw new NativeQueryRefreshError("protected", target.name);
  if (!table.headerRow || table.totalsRow) throw new NativeQueryRefreshError("resize_refused", table.name);

  let result: ReturnType<typeof powerQueryResultToSheet>;
  try {
    const members = queries.map((entry) => ({ name: entry.name, formula: entry.formula }));
    const value = evaluatePowerQuery(members, query.name, nativeQueryWorkbookItems(sheets, options), {
      culture: query.culture,
      now: options.now,
    });
    result = powerQueryResultToSheet(value);
  } catch (error) {
    if (error instanceof NativeQueryRefreshError) throw error;
    if (error instanceof PowerQueryUnsupportedError) throw new NativeQueryRefreshError("unsupported", error.feature);
    if (error instanceof PowerQueryError) throw new NativeQueryRefreshError("m_error", error.message);
    throw error;
  }

  // Columns by name (a table's column names are unique without case).
  const { r1, c1, r2, c2 } = table.range;
  const tableColumns = Array.from({ length: c2 - c1 + 1 }, (_, offset) => headerText(target.cells[`${r1}:${c1 + offset}`] as SpreadsheetCell | undefined));
  const mapping = tableColumns.map((column) => result.columns.findIndex((entry) => entry.toLowerCase() === column.toLowerCase()));
  if (result.columns.length !== tableColumns.length || mapping.some((index) => index < 0))
    throw new NativeQueryRefreshError("columns_changed", result.columns.join(", "));

  const dataRows = Math.max(1, result.rows.length);
  const nextRange = { r1, c1, r2: r1 + dataRows, c2 };
  if (nextRange.r2 !== r2) {
    // Excel inserts or deletes cells under the table; refuse where that
    // would move anything.
    for (const [key, cell] of Object.entries(target.cells)) {
      const [row, col] = key.split(":").map(Number);
      if (row > r2 && col >= c1 && col <= c2 && cell && String((cell as SpreadsheetCell).value ?? "") !== "")
        throw new NativeQueryRefreshError("content_below", table.name);
    }
    if (target.dataTables?.some((other) => other.id !== table.id && other.range.r1 > r2 && other.range.c1 <= c2 && other.range.c2 >= c1))
      throw new NativeQueryRefreshError("content_below", table.name);
  }
  let resized = target;
  if (nextRange.r2 !== r2) {
    const outcome = planDataTableResizeOutcome(target, table.id, nextRange, options.rangeHasLockedCell);
    if (!outcome.ok) throw new NativeQueryRefreshError("resize_refused", `${table.name} (${outcome.reason})`);
    resized = outcome.sheet;
  }
  const cells = { ...resized.cells };
  for (let rowIndex = 0; rowIndex < dataRows; rowIndex += 1) {
    const row = r1 + 1 + rowIndex;
    const sourceRow = Math.min(row, r2);
    mapping.forEach((resultIndex, offset) => {
      const key = `${row}:${c1 + offset}`;
      const base = (target.cells[key] ?? (row > r2 ? target.cells[`${sourceRow}:${c1 + offset}`] : undefined)) as SpreadsheetCell | undefined;
      const value = rowIndex < result.rows.length ? result.rows[rowIndex][resultIndex] : null;
      const written = writtenCell(value, base);
      if (written) cells[key] = written as (typeof cells)[string];
      else delete cells[key];
    });
  }
  for (let row = nextRange.r2 + 1; row <= r2; row += 1)
    for (let col = c1; col <= c2; col += 1) delete cells[`${row}:${col}`];
  const next = { ...resized, cells } as S;
  const nextSheets = sheets.slice();
  nextSheets[targetSheetIndex] = next;
  return { sheets: nextSheets, sheetId: target.id, tableId: table.id, rowCount: result.rows.length, range: nextRange };
}
