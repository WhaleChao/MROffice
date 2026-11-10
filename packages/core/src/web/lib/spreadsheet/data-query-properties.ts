/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Query Properties (a connection's Usage / Definition / Used In) and
 * External Data Properties (its query table) for MRAI's queries.
 *
 * One query is one connection and one query table, so its properties are one
 * record keyed by queryId in `DataQueryState.properties` on the sheet that
 * holds its output. Stored canonical: a key at its default is dropped, and an
 * absent key is the value MRAI wrote before queries carried properties
 * (foreground refresh, no timer, no refresh on open, data saved, part of
 * Refresh All, fixed column widths, formatting and sort/filter/layout kept).
 * The API's `validate_query_properties` is the same contract
 * (apps/api/app/services/file_adapters/xlsx/data_queries.py), and the native
 * writer maps each field to connections.xml / queryTable*.xml.
 */
import type {
  DataQueryRun,
  DataQuerySheet,
  DataQueryState,
} from "./data-queries";
import type { DataTable } from "./data-table";

export type DataQueryGrowShrinkType =
  | "insertDelete"
  | "insertClear"
  | "overwriteClear";

export interface DataQueryProperties {
  /** connection/@description */
  description?: string;
  /** Usage > Enable background refresh (connection/@background). */
  backgroundRefresh?: boolean;
  /** Usage > Refresh every n minutes (connection/@interval); 0 = off. */
  refreshInterval?: number;
  /** Usage > Refresh data when opening the file (refreshOnLoad). */
  refreshOnLoad?: boolean;
  /** Usage > Remove data from the external data range before saving the
   * workbook (connection/@saveData="0", queryTable/@removeDataOnSave). */
  removeDataOnSave?: boolean;
  /** Usage > Refresh this connection on Refresh All, cleared
   * (x15:connection/@excludeFromRefreshAll). */
  excludeFromRefreshAll?: boolean;
  /** External Data Properties > Include row numbers (queryTable/@rowNumbers). */
  rowNumbers?: boolean;
  /** > Adjust column width (queryTable/@adjustColumnWidth). */
  adjustColumnWidth?: boolean;
  /** > Preserve column sort/filter/layout
   * (queryTableRefresh/@preserveSortFilterLayout). */
  preserveSortFilterLayout?: boolean;
  /** > Preserve cell formatting (queryTable/@preserveFormatting). */
  preserveFormatting?: boolean;
  /** > If the number of rows in the data range changes upon refresh. */
  growShrinkType?: DataQueryGrowShrinkType;
  /** > Fill down formulas in columns adjacent to data (queryTable/@fillFormulas). */
  fillFormulas?: boolean;
}

export type EffectiveDataQueryProperties = Required<
  Omit<DataQueryProperties, "description">
> & { description: string };

export const DATA_QUERY_PROPERTY_DEFAULTS: Readonly<
  Required<Omit<DataQueryProperties, "description">>
> = Object.freeze({
  backgroundRefresh: false,
  refreshInterval: 0,
  refreshOnLoad: false,
  removeDataOnSave: false,
  excludeFromRefreshAll: false,
  rowNumbers: false,
  adjustColumnWidth: false,
  preserveSortFilterLayout: true,
  preserveFormatting: true,
  growShrinkType: "insertDelete",
  fillFormulas: false,
});

/**
 * What Excel gives a query it creates (Get Data > ... > Load): background
 * refresh on and Adjust column width on, everything else at its default.
 */
export const DATA_QUERY_NEW_QUERY_PROPERTIES: Readonly<DataQueryProperties> =
  Object.freeze({ backgroundRefresh: true, adjustColumnWidth: true });

export const DATA_QUERY_MAX_DESCRIPTION = 255;
/** Excel's "Refresh every n minutes" box takes 1 to 32,767. */
export const DATA_QUERY_MAX_REFRESH_INTERVAL = 32_767;
/** The minutes Excel's box shows before a timer was ever set. */
export const DATA_QUERY_DEFAULT_REFRESH_MINUTES = 60;

const GROW_SHRINK_TYPES: readonly DataQueryGrowShrinkType[] = [
  "insertDelete",
  "insertClear",
  "overwriteClear",
];

export class DataQueryPropertiesError extends Error {
  constructor() {
    super("query_properties_invalid");
  }
}

/** One query's properties, canonical (defaults dropped); throws when invalid. */
export function parseDataQueryProperties(input: unknown): DataQueryProperties {
  if (!input || typeof input !== "object" || Array.isArray(input))
    throw new DataQueryPropertiesError();
  const allowed = new Set([
    ...Object.keys(DATA_QUERY_PROPERTY_DEFAULTS),
    "description",
  ]);
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(input)) {
    if (!allowed.has(key)) throw new DataQueryPropertiesError();
    if (key === "description") {
      if (
        typeof value !== "string" ||
        value.length > DATA_QUERY_MAX_DESCRIPTION ||
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u.test(value)
      )
        throw new DataQueryPropertiesError();
      if (value) result.description = value;
      continue;
    }
    if (key === "refreshInterval") {
      if (
        !Number.isSafeInteger(value) ||
        (value as number) < 0 ||
        (value as number) > DATA_QUERY_MAX_REFRESH_INTERVAL
      )
        throw new DataQueryPropertiesError();
    } else if (key === "growShrinkType") {
      if (!GROW_SHRINK_TYPES.includes(value as DataQueryGrowShrinkType))
        throw new DataQueryPropertiesError();
    } else if (typeof value !== "boolean") {
      throw new DataQueryPropertiesError();
    }
    if (
      value !==
      DATA_QUERY_PROPERTY_DEFAULTS[key as keyof typeof DATA_QUERY_PROPERTY_DEFAULTS]
    )
      result[key] = value;
  }
  return result as DataQueryProperties;
}

/** Validate a state's `properties` map against the queries its ledger holds. */
export function parseDataQueryPropertiesMap(
  input: unknown,
  knownQueryIds: ReadonlySet<string>,
  limit: number,
): void {
  if (!input || typeof input !== "object" || Array.isArray(input))
    throw new DataQueryPropertiesError();
  const entries = Object.entries(input);
  if (entries.length > limit) throw new DataQueryPropertiesError();
  for (const [queryId, value] of entries) {
    if (!knownQueryIds.has(queryId)) throw new DataQueryPropertiesError();
    parseDataQueryProperties(value);
  }
}

function stateFor<T extends DataQuerySheet>(
  sheets: readonly T[],
  queryId: string,
): { sheet: T; state: DataQueryState } | null {
  for (const sheet of sheets) {
    const state = sheet.dataQueries;
    if (state && Object.values(state.runs).some((run) => run.queryId === queryId))
      return { sheet, state };
  }
  return null;
}

/** A query's properties with every default filled in. */
export function effectiveDataQueryProperties(
  sheets: readonly DataQuerySheet[],
  queryId: string,
): EffectiveDataQueryProperties {
  const stored = stateFor(sheets, queryId)?.state.properties?.[queryId];
  let parsed: DataQueryProperties = {};
  try {
    parsed = stored ? parseDataQueryProperties(stored) : {};
  } catch {
    parsed = {};
  }
  return { ...DATA_QUERY_PROPERTY_DEFAULTS, description: "", ...parsed };
}

/** The workbook with one query's properties replaced (canonical). */
export function withDataQueryProperties<T extends DataQuerySheet>(
  sheets: readonly T[],
  queryId: string,
  properties: DataQueryProperties,
): T[] {
  const owner = stateFor(sheets, queryId);
  if (!owner) throw new DataQueryPropertiesError();
  const canonical = parseDataQueryProperties(properties);
  return sheets.map((sheet) => {
    if (sheet !== owner.sheet) return sheet;
    const map = { ...owner.state.properties };
    if (Object.keys(canonical).length) map[queryId] = canonical;
    else delete map[queryId];
    const state: DataQueryState = {
      active: owner.state.active,
      runs: owner.state.runs,
    };
    if (Object.keys(map).length) state.properties = map;
    return { ...sheet, dataQueries: state };
  });
}

/** Refresh All refreshes every query except those taken out of it. */
export function dataQueryRunsForRefreshAll(
  sheets: readonly DataQuerySheet[],
  runs: readonly DataQueryRun[],
): DataQueryRun[] {
  return runs.filter(
    (run) => !effectiveDataQueryProperties(sheets, run.queryId).excludeFromRefreshAll,
  );
}

function inRange(area: DataTable["range"], row: number, col: number): boolean {
  return row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2;
}

/**
 * What a refresh does beside writing values, by the query table's External
 * Data Properties (Excel): Preserve cell formatting off drops the formatting
 * of the cells the output covered or now covers (keeping only the quote
 * prefix a "=" text needs); Preserve column sort/filter/layout off drops the
 * table's filter and sort. `before` is the workbook the refresh planned from.
 */
export function applyDataQueryRefreshProperties<T extends DataQuerySheet>(
  before: readonly T[],
  planned: readonly T[],
  queryIds: readonly string[],
): T[] {
  let result = [...planned];
  for (const queryId of queryIds) {
    const properties = effectiveDataQueryProperties(result, queryId);
    if (properties.preserveFormatting && properties.preserveSortFilterLayout) continue;
    const owner = stateFor(result, queryId);
    const runId = owner?.state.active[queryId];
    const run = runId ? owner?.state.runs[runId] : undefined;
    if (!owner || !run) continue;
    const previous = before.find((sheet) => sheet.id === owner.sheet.id);
    const oldRunId = previous?.dataQueries?.active[queryId];
    const oldRun = oldRunId ? previous?.dataQueries?.runs[oldRunId] : undefined;
    const oldArea = oldRun
      ? previous?.dataTables?.find((table) => table.id === oldRun.targetTableId)?.range
      : undefined;
    const table = owner.sheet.dataTables?.find((entry) => entry.id === run.targetTableId);
    if (!table) continue;
    const next = { ...owner.sheet, cells: { ...owner.sheet.cells } } as T;
    if (!properties.preserveFormatting) {
      const areas = [table.range, ...(oldArea ? [oldArea] : [])];
      for (const area of areas)
        for (let row = area.r1; row <= area.r2; row++)
          for (let col = area.c1; col <= area.c2; col++) {
            const key = `${row}:${col}`;
            const cell = next.cells[key];
            if (!cell || !("style" in cell)) continue;
            const kept: Record<string, unknown> = {};
            const style = cell.style as Record<string, unknown> | undefined;
            if (
              style?.quotePrefix === true &&
              inRange(table.range, row, col) &&
              cell.kind !== "number"
            )
              kept.quotePrefix = true;
            const written = { ...cell } as Record<string, unknown>;
            if (Object.keys(kept).length) written.style = kept;
            else delete written.style;
            if (Object.keys(written).length === 1 && written.value === "")
              delete next.cells[key];
            else next.cells[key] = written as (typeof next.cells)[string];
          }
    }
    if (!properties.preserveSortFilterLayout) {
      next.dataTables = (next.dataTables ?? []).map((entry) => {
        if (entry.id !== table.id) return entry;
        const cleared = { ...entry };
        delete cleared.columnFilters;
        delete cleared.sortState;
        return cleared;
      });
    }
    result = result.map((sheet) => (sheet.id === next.id ? next : sheet));
  }
  return result;
}

/**
 * The columns right of a query's output whose first data row holds a
 * formula, nearest first and contiguous: what "Fill down formulas in columns
 * adjacent to data" fills. The API's `_adjacent_formula_columns` is the same.
 */
export function dataQueryAdjacentFormulaColumns(
  sheet: DataQuerySheet,
  area: DataTable["range"],
): number[] {
  const columns: number[] = [];
  if (area.r2 <= area.r1) return columns;
  for (let col = area.c2 + 1; col < sheet.columnCount; col++) {
    const value = sheet.cells[`${area.r1 + 1}:${col}`]?.value;
    if (typeof value !== "string" || !value.startsWith("=")) break;
    columns.push(col);
  }
  return columns;
}

/**
 * Fill down formulas in columns adjacent to data (External Data Properties):
 * when a refresh changes how many rows the output has, each formula next to
 * it is filled down from the first data row to the new last row, and the
 * formulas below the new last row (down to the old one) are removed, as
 * Excel keeps the adjacent formulas level with the data. `shift` moves a
 * formula's relative references as a fill does (formula-engine shiftFormula).
 */
export function fillDataQueryAdjacentFormulas<T extends DataQuerySheet>(
  before: readonly T[],
  planned: readonly T[],
  queryIds: readonly string[],
  shift: (formula: string, rowDelta: number, colDelta: number) => string,
): T[] {
  let result = [...planned];
  for (const queryId of queryIds) {
    if (!effectiveDataQueryProperties(result, queryId).fillFormulas) continue;
    const owner = stateFor(result, queryId);
    const run = owner ? owner.state.runs[owner.state.active[queryId]] : undefined;
    const previous = owner ? before.find((sheet) => sheet.id === owner.sheet.id) : undefined;
    const oldRun = previous?.dataQueries
      ? previous.dataQueries.runs[previous.dataQueries.active[queryId]]
      : undefined;
    const oldArea = oldRun
      ? previous?.dataTables?.find((table) => table.id === oldRun.targetTableId)?.range
      : undefined;
    const area = run
      ? owner?.sheet.dataTables?.find((table) => table.id === run.targetTableId)?.range
      : undefined;
    if (!owner || !previous || !oldArea || !area || oldArea.r2 === area.r2) continue;
    const columns = dataQueryAdjacentFormulaColumns(previous, oldArea);
    if (!columns.length) continue;
    const next = { ...owner.sheet, cells: { ...owner.sheet.cells } } as T;
    const firstRow = oldArea.r1 + 1;
    for (const col of columns) {
      const anchor = previous.cells[`${firstRow}:${col}`];
      const formula = String(anchor.value);
      for (let row = firstRow + 1; row <= area.r2; row++) {
        const value = shift(formula, row - firstRow, 0);
        const filled = { ...anchor, value } as Record<string, unknown>;
        // The cell's formula field is the same shifted formula: left as the
        // anchor's, the XLSX wrote every row's old reference (B521).
        if (typeof anchor.formula === "string" && anchor.formula) {
          filled.formula = anchor.formula.startsWith("=") ? value : value.replace(/^=/, "");
        } else {
          delete filled.formula;
        }
        delete filled.display;
        delete filled.chartValueCache;
        delete filled.spillRef;
        next.cells[`${row}:${col}`] = filled as (typeof next.cells)[string];
      }
      for (let row = Math.max(area.r2 + 1, firstRow + 1); row <= oldArea.r2; row++) {
        const key = `${row}:${col}`;
        const cell = next.cells[key];
        if (typeof cell?.value === "string" && cell.value.startsWith("=")) delete next.cells[key];
      }
    }
    result = result.map((sheet) => (sheet.id === next.id ? next : sheet));
  }
  return result;
}

/**
 * The columns a query table's Adjust column width refits after a refresh:
 * each output column with the texts it shows (header and data), measured by
 * the caller in the cells' own fonts.
 */
export function dataQueryOutputColumnTexts(
  sheet: DataQuerySheet,
  area: DataTable["range"],
): Array<{ col: number; cells: Array<{ text: string; style: unknown }> }> {
  const columns: Array<{ col: number; cells: Array<{ text: string; style: unknown }> }> = [];
  for (let col = area.c1; col <= area.c2; col++) {
    const cells: Array<{ text: string; style: unknown }> = [];
    for (let row = area.r1; row <= area.r2; row++) {
      const cell = sheet.cells[`${row}:${col}`];
      if (cell?.value) cells.push({ text: String(cell.value), style: cell.style });
    }
    columns.push({ col, cells });
  }
  return columns;
}

function columnName(index: number): string {
  let name = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26))
    name = String.fromCharCode(65 + ((value - 1) % 26)) + name;
  return name;
}

/** "$E$1:$F$3", as Excel's Used In tab shows a query table's location. */
export function dataQueryAbsoluteRange(area: DataTable["range"]): string {
  const start = `$${columnName(area.c1)}$${area.r1 + 1}`;
  const end = `$${columnName(area.c2)}$${area.r2 + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/**
 * The Definition tab of a query's connection, as Excel shows a Power Query
 * connection: an OLE DB query on the Mashup provider over this workbook,
 * whose command selects the query by name. The strings are the ones the
 * native writer puts in connections.xml (dbPr).
 */
export function dataQueryConnectionDefinition(name: string): {
  connectionName: string;
  connectionString: string;
  commandText: string;
} {
  return {
    connectionName: `Query - ${name}`,
    connectionString: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${name};Extended Properties=""`,
    commandText: `SELECT * FROM [${name.replaceAll("]", "]]")}]`,
  };
}

/** The Used In rows: the sheet, the table and the cells a query fills. */
export function dataQueryUsedIn(
  sheets: readonly DataQuerySheet[],
  run: DataQueryRun,
): { sheetName: string; tableName: string; location: string } | null {
  const sheet = sheets.find((entry) => entry.id === run.targetSheetId);
  const table = sheet?.dataTables?.find((entry) => entry.id === run.targetTableId);
  if (!sheet || !table) return null;
  return {
    sheetName: sheet.name,
    tableName: table.name,
    location: dataQueryAbsoluteRange(table.range),
  };
}

