/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_DATA_TABLES, type DataTable } from "./data-table";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { classifySpreadsheetNumberFormat } from "./date-filter-tree";
import {
  parseDataQueryPropertiesMap,
  type DataQueryProperties,
  dataQueryAdjacentFormulaColumns,
} from "./data-query-properties";

/** Closed current-workbook query algebra. No I/O, writes, formulas or arbitrary M.
 * Excel Web: Get Data -> Transform -> Close & Load; Refresh is a separate write.
 * The native codec is MS-QDEFF, not a private JSON substitute for Power Query.
 * Visible commands commit only through the authenticated native-query writer.
 */
export const DATA_QUERY_LIMITS = Object.freeze({
  cells: 50_000,
  columns: 256,
  steps: 32,
  queries: 32,
  text: 32_767,
  totalText: 1_000_000,
});

export type QueryScalar = string | number | boolean | null;
export type DataQueryStep =
  | { kind: "selectColumns"; columns: string[] }
  | { kind: "renameColumn"; from: string; to: string }
  | { kind: "filterEquals"; column: string; value: QueryScalar }
  | { kind: "removeBlankRows" }
  | { kind: "distinct"; columns: string[] };
export interface DataQueryDefinition {
  version: 1;
  name: string;
  sourceName: string;
  steps: DataQueryStep[];
}
export interface DataQueryTable {
  columns: string[];
  rows: QueryScalar[][];
}

export interface DataQueryRun {
  id: string;
  queryId: string;
  definition: DataQueryDefinition;
  sourceSheetId: string;
  sourceTableId: string;
  targetSheetId: string;
  targetTableId: string;
  targetName: string;
  range: DataTable["range"];
  columns: NonNullable<DataTable["columns"]>;
  sourceHash: string;
  resultHash: string;
  beforeHash: string;
  priorRunId: string | null;
  createdSourceTable?: DataTable;
}
export interface DataQueryState {
  active: Record<string, string>;
  runs: Record<string, DataQueryRun>;
  /** Query Properties / External Data Properties by queryId
   * (data-query-properties.ts); absent = every default. */
  properties?: Record<string, DataQueryProperties>;
}
export function parseDataQueryState(input: unknown): DataQueryState {
  const state = record(input);
  keys(state, [
    "active",
    "runs",
    ...(Object.hasOwn(state, "properties") ? ["properties"] : []),
  ]);
  const active = record(state.active),
    runs = record(state.runs);
  if (
    Object.keys(active).length > DATA_QUERY_LIMITS.queries ||
    Object.keys(runs).length > 256 ||
    new TextEncoder().encode(JSON.stringify(state)).length > 1_000_000
  )
    fail("query_history_limit");
  const safeId = (id: unknown): string => {
    if (
      typeof id !== "string" ||
      !/^[A-Za-z0-9][A-Za-z0-9_-]{0,199}$/.test(id) ||
      ["__proto__", "constructor", "prototype"].includes(id)
    )
      fail("query_identity_ambiguous");
    return id;
  };
  for (const [id, raw] of Object.entries(runs)) {
    safeId(id);
    const run = record(raw);
    keys(run, [
      "id",
      "queryId",
      "definition",
      "sourceSheetId",
      "sourceTableId",
      "targetSheetId",
      "targetTableId",
      "targetName",
      "range",
      "sourceHash",
      "resultHash",
      "beforeHash",
      "priorRunId",
      "columns",
      ...(Object.hasOwn(run, "createdSourceTable")
        ? ["createdSourceTable"]
        : []),
    ]);
    if (run.id !== id) fail("query_identity_ambiguous");
    safeId(run.queryId);
    for (const key of [
      "sourceSheetId",
      "sourceTableId",
      "targetSheetId",
      "targetTableId",
    ])
      label(run[key], 200);
    parseDataQueryDefinition(run.definition);
    if (!Array.isArray(run.columns)) fail("query_columns_invalid");
    const declaredColumns = run.columns.map((column) => record(column));
    columnList(declaredColumns.map((column) => column.name));
    const columnIds = declaredColumns.map((column) => column.id);
    if (
      new Set(columnIds).size !== columnIds.length ||
      columnIds.some(
        (id) =>
          !Number.isSafeInteger(id) ||
          (id as number) <= 0 ||
          (id as number) >= 2 ** 31,
      )
    )
      fail("query_column_identity_changed");
    declaredColumns.forEach((column) => keys(column, ["id", "name"]));
    label(run.targetName, 200);
    for (const key of ["sourceHash", "resultHash", "beforeHash"]) {
      if (
        typeof run[key] !== "string" ||
        !/^[0-9a-f]{64}$/.test(run[key] as string)
      )
        fail("query_receipt_invalid");
    }
    const area = record(run.range);
    keys(area, ["r1", "c1", "r2", "c2"]);
    if (
      !Object.values(area).every(
        (value) => Number.isSafeInteger(value) && (value as number) >= 0,
      ) ||
      (area.r2 as number) < (area.r1 as number) ||
      (area.c2 as number) < (area.c1 as number)
    )
      fail("query_range_invalid");
    if (run.priorRunId !== null) safeId(run.priorRunId);
    if (
      declaredColumns.length !==
      (area.c2 as number) - (area.c1 as number) + 1
    )
      fail("query_columns_invalid");
    const visited = new Set([id]);
    let prior = run.priorRunId;
    while (prior !== null) {
      const parentId = safeId(prior);
      if (visited.has(parentId) || !Object.hasOwn(runs, parentId))
        fail("query_history_invalid");
      visited.add(parentId);
      const parent = record(runs[parentId]);
      if (
        parent.queryId !== run.queryId ||
        parent.targetTableId !== run.targetTableId ||
        parent.targetSheetId !== run.targetSheetId
      )
        fail("query_history_invalid");
      prior = parent.priorRunId;
    }
  }
  for (const [id, runId] of Object.entries(active)) {
    safeId(id);
    if (
      !Object.hasOwn(runs, safeId(runId)) ||
      record(runs[runId as string]).queryId !== id
    )
      fail("query_identity_ambiguous");
  }
  if (Object.hasOwn(state, "properties")) {
    try {
      parseDataQueryPropertiesMap(
        state.properties,
        new Set(Object.values(runs).map((run) => record(run).queryId as string)),
        DATA_QUERY_LIMITS.queries,
      );
    } catch {
      fail("query_properties_invalid");
    }
  }
  return input as DataQueryState;
}
interface QueryCell {
  value: string;
  kind?: "number" | "text";
  formula?: string;
  runs?: unknown;
  [key: string]: unknown;
}
export interface DataQuerySheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, QueryCell>;
  dataTables?: DataTable[];
  dataQueries?: DataQueryState;
  protected?: boolean;
  namedSheetViewsNativeBlocker?: string;
}
type QueryToken = [string, string];

function checkedRange(area: DataTable["range"], sheet: DataQuerySheet) {
  if (
    ![area.r1, area.c1, area.r2, area.c2].every(Number.isSafeInteger) ||
    area.r1 < 0 ||
    area.c1 < 0 ||
    area.r2 < area.r1 ||
    area.c2 < area.c1 ||
    area.r2 >= sheet.rowCount ||
    area.c2 >= sheet.columnCount
  )
    fail("query_range_invalid");
  if (
    (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) >
    DATA_QUERY_LIMITS.cells
  )
    fail("query_limit");
  return area;
}
function cellToken(cell?: QueryCell): QueryToken {
  if (!cell) return ["", ""];
  if (cell.formula || cell.runs || typeof cell.value !== "string")
    fail("query_cell_unsupported");
  if (!cell.value) return ["", ""];
  const style = cell.style;
  // TRUE/FALSE without a type could be the logical value or the word. An
  // entry typed after an apostrophe, or in a Text-formatted cell, is the
  // word and says so (live audit B380); anything else stays refused.
  const explicitText =
    cell.kind !== "number" &&
    typeof style === "object" &&
    style !== null &&
    ((style as { quotePrefix?: unknown }).quotePrefix === true ||
      (style as { numberFormat?: unknown }).numberFormat === "@");
  if (cell.kind !== "number" && ["TRUE", "FALSE"].includes(cell.value) && !explicitText)
    fail("query_cell_type_ambiguous");
  if (
    cell.kind === "number" &&
    style &&
    typeof style === "object" &&
    "numberFormat" in style &&
    typeof style.numberFormat === "string"
  ) {
    const semantics = classifySpreadsheetNumberFormat(
      style.numberFormat,
      Number(cell.value),
    );
    if (!semantics.valid || semantics.hasCalendarDate || semantics.hasClock)
      fail("query_cell_type_unsupported");
  }
  scalar(cell.value);
  return [cell.kind ?? "text", cell.value];
}
function rangeTokens(
  sheet: DataQuerySheet,
  area: DataTable["range"],
): QueryToken[][] {
  checkedRange(area, sheet);
  return Array.from({ length: area.r2 - area.r1 + 1 }, (_, row) =>
    Array.from({ length: area.c2 - area.c1 + 1 }, (_, col) =>
      cellToken(sheet.cells[`${row + area.r1}:${col + area.c1}`]),
    ),
  );
}
async function queryFingerprint(value: unknown): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(JSON.stringify(value)),
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}
function oneSheet<T extends DataQuerySheet>(
  sheets: readonly T[],
  id: string,
): T {
  const matches = sheets.filter((sheet) => sheet.id === id);
  if (matches.length !== 1) fail("query_sheet_ambiguous");
  return matches[0];
}
function oneTable(sheet: DataQuerySheet, id: string): DataTable {
  const matches = sheet.dataTables?.filter((table) => table.id === id) ?? [];
  if (matches.length !== 1) fail("query_table_ambiguous");
  return matches[0];
}

export function readDataQuerySource(
  sheet: DataQuerySheet,
  table: DataTable,
): DataQueryTable {
  if (
    sheet.protected ||
    table.totalsRow ||
    table.nativeResizeBlocker ||
    table.calculatedFormulas ||
    sheet.namedSheetViewsNativeBlocker
  )
    fail("query_source_unsupported");
  const tokens = rangeTokens(sheet, table.range);
  if (table.headerRow && tokens[0].some(([kind]) => kind !== "text"))
    fail("query_headers_unsupported");
  return {
    columns: columnList(
      table.headerRow
        ? tokens[0].map(([, value]) => value)
        : table.columns?.map((column) => column.name),
    ),
    rows: (table.headerRow ? tokens.slice(1) : tokens).map((row) =>
      row.map(([kind, value]) => {
        if (!value) return null;
        if (kind === "text") return value;
        if (
          kind !== "number" ||
          !/^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?$/.test(value)
        )
          fail("query_cell_unsupported");
        return scalar(Number(value));
      }),
    ),
  };
}

function outputTokens(table: DataQueryTable): QueryToken[][] {
  return [
    table.columns.map((name): QueryToken => ["text", name]),
    ...table.rows.map((row) =>
      row.map((value): QueryToken => {
        if (value === null || value === "") return ["", ""];
        // The current editor cell contract is text/number. A native logical value
        // requires a truthful adapter and is not silently converted to text here.
        if (typeof value === "boolean") return fail("query_cell_unsupported");
        return [typeof value === "number" ? "number" : "text", String(value)];
      }),
    ),
  ];
}

export function activeDataQueryRuns(
  sheets: readonly DataQuerySheet[],
): DataQueryRun[] {
  return sheets.flatMap((sheet) =>
    Object.entries(sheet.dataQueries?.active ?? {}).map(([id, runId]) => {
      const run = sheet.dataQueries?.runs[runId];
      if (
        !run ||
        run.queryId !== id ||
        run.id !== runId ||
        run.targetSheetId !== sheet.id
      )
        return fail("query_identity_ambiguous");
      return run;
    }),
  );
}

/** Browser Undo restores the active pointer but never drops server-verified
 * execution records. It cannot certify a new result by manufacturing history. */
export function mergeForwardDataQueryHistory<T extends DataQuerySheet>(
  restored: readonly T[],
  current: readonly T[],
): T[] {
  return restored.map((sheet) => {
    const live = current.find((candidate) => candidate.id === sheet.id);
    if (!live?.dataQueries) return sheet;
    const runs = { ...live.dataQueries.runs };
    for (const [id, run] of Object.entries(sheet.dataQueries?.runs ?? {})) {
      if (runs[id] && JSON.stringify(runs[id]) !== JSON.stringify(run))
        fail("query_receipt_removed_or_changed");
      runs[id] = run;
    }
    const properties = sheet.dataQueries?.properties;
    return {
      ...sheet,
      dataQueries: {
        active: { ...sheet.dataQueries?.active },
        runs,
        ...(properties ? { properties } : {}),
      },
    };
  });
}

/** Invert only a query transaction against the live workbook. Source/peer edits
 * and formatting are not history snapshots and must never be rolled back. */
export function restoreDataQueryHistory<T extends DataQuerySheet>(
  current: readonly T[],
  desired: readonly T[],
  expected: readonly T[],
): T[] {
  const result = current.map((sheet) => ({
    ...sheet,
    cells: { ...sheet.cells },
    ...(sheet.dataTables ? { dataTables: [...sheet.dataTables] } : {}),
    dataQueries: sheet.dataQueries
      ? {
          active: { ...sheet.dataQueries.active },
          runs: { ...sheet.dataQueries.runs },
          ...(sheet.dataQueries.properties
            ? { properties: sheet.dataQueries.properties }
            : {}),
        }
      : undefined,
  }));
  const before = new Map(
    activeDataQueryRuns(expected).map((run) => [run.queryId, run]),
  );
  const after = new Map(
    activeDataQueryRuns(desired).map((run) => [run.queryId, run]),
  );
  for (const id of new Set([...before.keys(), ...after.keys()])) {
    const oldRun = before.get(id),
      run = after.get(id);
    if (oldRun?.id === run?.id) continue;
    const owner = (run ?? oldRun)!;
    const live = oneSheet(result, owner.targetSheetId);
    if (live.protected || live.namedSheetViewsNativeBlocker)
      fail("query_target_protected");
    if (live.dataQueries?.active[id] !== oldRun?.id) fail("query_source_stale");
    const expectedSheet = oneSheet(expected, owner.targetSheetId);
    const desiredSheet = oneSheet(desired, owner.targetSheetId);
    const oldTable = oldRun
      ? oneTable(expectedSheet, owner.targetTableId)
      : undefined;
    const newTable = run
      ? oneTable(desiredSheet, owner.targetTableId)
      : undefined;
    const liveTable = oldRun ? oneTable(live, owner.targetTableId) : undefined;
    if (
      !oldRun &&
      live.dataTables?.some((table) => table.id === owner.targetTableId)
    )
      fail("query_target_identity_changed");
    if (
      liveTable &&
      (JSON.stringify(liveTable.columns) !==
        JSON.stringify(oldTable!.columns) ||
        liveTable.range.r2 - liveTable.range.r1 !==
          oldTable!.range.r2 - oldTable!.range.r1 ||
        liveTable.range.c2 - liveTable.range.c1 !==
          oldTable!.range.c2 - oldTable!.range.c1)
    )
      fail("query_output_modified");
    const origin = liveTable?.range ?? newTable!.range;
    const restoredArea = newTable
      ? checkedRange(
          {
            r1: origin.r1,
            c1: origin.c1,
            r2: origin.r1 + newTable.range.r2 - newTable.range.r1,
            c2: origin.c1 + newTable.range.c2 - newTable.range.c1,
          },
          live,
        )
      : undefined;
    const oldArea = liveTable?.range;
    const height = Math.max(
      oldTable ? oldTable.range.r2 - oldTable.range.r1 + 1 : 0,
      newTable ? newTable.range.r2 - newTable.range.r1 + 1 : 0,
    );
    const width = Math.max(
      oldTable ? oldTable.range.c2 - oldTable.range.c1 + 1 : 0,
      newTable ? newTable.range.c2 - newTable.range.c1 + 1 : 0,
    );
    for (const table of live.dataTables ?? []) {
      if (table.id === owner.targetTableId || !restoredArea) continue;
      const other = table.range;
      if (!(
        restoredArea.r2 < other.r1 ||
        restoredArea.r1 > other.r2 ||
        restoredArea.c2 < other.c1 ||
        restoredArea.c1 > other.c2
      ))
        fail("query_output_table_collision");
    }
    for (let row = 0; row < height; row++)
      for (let col = 0; col < width; col++) {
        const inArea = (area: DataTable["range"] | undefined) =>
          area && row <= area.r2 - area.r1 && col <= area.c2 - area.c1;
        if (!inArea(oldArea) && !inArea(restoredArea)) continue;
        const key = `${origin.r1 + row}:${origin.c1 + col}`;
        const oldCell =
          oldTable && inArea(oldTable.range)
            ? expectedSheet.cells[
                `${oldTable.range.r1 + row}:${oldTable.range.c1 + col}`
              ]
            : undefined;
        if (
          JSON.stringify(cellToken(live.cells[key])) !==
          JSON.stringify(cellToken(oldCell))
        )
          fail("query_output_modified");
        const restoredCell =
          newTable && inArea(newTable.range)
            ? desiredSheet.cells[
                `${newTable.range.r1 + row}:${newTable.range.c1 + col}`
              ]
            : undefined;
        const [kind, value] = cellToken(restoredCell);
        const cell: QueryCell = { ...live.cells[key], value };
        delete cell.kind;
        delete cell.display;
        delete cell.chartValueCache;
        if (value) cell.kind = kind === "number" ? "number" : "text";
        if (!value && Object.keys(cell).length === 1) delete live.cells[key];
        else live.cells[key] = cell;
      }
    // The formulas the refresh filled beside the output (External Data
    // Properties) are the same step's writes: an undo takes them back to
    // what they were, a redo re-fills them (B520). Only a cell the step
    // left as the live sheet still shows is touched.
    if (oldTable && newTable) {
      const columns = dataQueryAdjacentFormulaColumns(expectedSheet, oldTable.range);
      const firstRow = oldTable.range.r1 + 1;
      const lastRow = Math.max(oldTable.range.r2, newTable.range.r2);
      for (const col of columns) {
        for (let row = firstRow + 1; row <= lastRow; row++) {
          const key = `${row}:${col}`;
          const expectedCell = expectedSheet.cells[key];
          const liveCell = live.cells[key];
          if (JSON.stringify(liveCell ?? null) !== JSON.stringify(expectedCell ?? null)) continue;
          const desiredCell = desiredSheet.cells[key];
          if (desiredCell) live.cells[key] = { ...desiredCell };
          else delete live.cells[key];
        }
      }
    }
    live.dataTables = (live.dataTables ?? []).filter(
      (table) => table.id !== owner.targetTableId,
    );
    if (newTable)
      live.dataTables.push({
        ...(liveTable ?? newTable),
        range: restoredArea!,
        columns: newTable.columns,
      });
    live.dataQueries ??= { active: {}, runs: {} };
    if (run) {
      live.dataQueries.active[id] = run.id;
      live.dataQueries.runs[run.id] = run;
    } else delete live.dataQueries.active[id];
    // A query transaction that changed the query's properties too (a rename
    // from Query Properties) restores them with it.
    const restoredProperties = desiredSheet.dataQueries?.properties?.[id];
    const properties = { ...live.dataQueries.properties };
    if (restoredProperties) properties[id] = restoredProperties;
    else delete properties[id];
    if (Object.keys(properties).length) live.dataQueries.properties = properties;
    else delete live.dataQueries.properties;
    const created = (oldRun ?? run)?.createdSourceTable;
    if (created) {
      const source = oneSheet(result, owner.sourceSheetId);
      const expectedSource = oneSheet(expected, owner.sourceSheetId);
      const desiredSource = oneSheet(desired, owner.sourceSheetId);
      const wasPresent = expectedSource.dataTables?.some(
        (table) => table.id === created.id,
      );
      const shouldBePresent = desiredSource.dataTables?.some(
        (table) => table.id === created.id,
      );
      if (wasPresent !== shouldBePresent) {
        if (source.protected) fail("query_source_unsupported");
        if (wasPresent) {
          if (
            JSON.stringify(oneTable(source, created.id)) !==
              JSON.stringify(created) ||
            activeDataQueryRuns(result).some(
              (entry) =>
                entry.sourceSheetId === source.id &&
                entry.sourceTableId === created.id,
            )
          )
            fail("query_source_identity_changed");
          source.dataTables = source.dataTables?.filter(
            (table) => table.id !== created.id,
          );
        } else {
          if (
            source.dataTables?.some(
              (table) =>
                table.id === created.id ||
                table.name.toLowerCase() === created.name.toLowerCase(),
            )
          )
            fail("query_source_identity_changed");
          source.dataTables = [...(source.dataTables ?? []), created];
        }
      }
    }
  }
  return mergeForwardDataQueryHistory(result, current);
}

/** Ordinary sheet operations cannot copy a query's identities or orphan it. */
export function assertDataQueryBindings(
  sheets: readonly DataQuerySheet[],
): void {
  const identities = new Set<string>();
  for (const run of activeDataQueryRuns(sheets)) {
    if (identities.has(run.queryId)) fail("query_identity_ambiguous");
    identities.add(run.queryId);
    oneTable(oneSheet(sheets, run.sourceSheetId), run.sourceTableId);
    oneTable(oneSheet(sheets, run.targetSheetId), run.targetTableId);
  }
}

export type DataQueryStatus = "current" | "stale" | "edited" | "blocked";
export function dataQueryOutputRange(
  sheets: readonly DataQuerySheet[],
  run: DataQueryRun,
): DataTable["range"] {
  const target = oneSheet(sheets, run.targetSheetId);
  return checkedRange(oneTable(target, run.targetTableId).range, target);
}
export async function inspectDataQueryRun(
  sheets: readonly DataQuerySheet[],
  run: DataQueryRun,
): Promise<DataQueryStatus> {
  try {
    const source = oneSheet(sheets, run.sourceSheetId),
      table = oneTable(source, run.sourceTableId);
    const sourceData = readDataQuerySource(source, table);
    const target = oneSheet(sheets, run.targetSheetId);
    if (target.protected) return "blocked";
    if (
      (await queryFingerprint(
        rangeTokens(target, dataQueryOutputRange(sheets, run)),
      )) !== run.resultHash
    )
      return "edited";
    const area = table.range;
    const fingerprint = await queryFingerprint([
      source.id,
      table.id,
      table.name,
      table.headerRow,
      [area.r1, area.c1, area.r2, area.c2],
      rangeTokens(source, area),
    ]);
    if (fingerprint !== run.sourceHash) return "stale";
    return (await queryFingerprint(
      outputTokens(previewDataQuery(run.definition, sourceData)),
    )) === run.resultHash
      ? "current"
      : "stale";
  } catch {
    return "blocked";
  }
}

export interface DataQueryRunRequest {
  definition: DataQueryDefinition;
  queryId: string;
  runId: string;
  sourceSheetId: string;
  sourceTableId: string;
  targetSheetId: string;
  targetTableId: string;
  targetName: string;
  row: number;
  col: number;
  createdSourceTable?: DataTable;
}

function queryColumnLineage(
  definition: DataQueryDefinition,
  names: string[],
): Array<{ origin: number; name: string }> {
  let columns = names.map((name, origin) => ({ origin, name }));
  for (const step of definition.steps) {
    if (step.kind === "selectColumns")
      columns = step.columns.map((name) => {
        const column = columns.find((entry) => entry.name === name);
        if (!column) fail("query_column_missing");
        return column;
      });
    else if (step.kind === "renameColumn") {
      if (!columns.some((column) => column.name === step.from))
        fail("query_column_missing");
      columns = columns.map((column) =>
        column.name === step.from ? { ...column, name: step.to } : column,
      );
    }
  }
  return columns;
}

function queryOutputColumns(
  definition: DataQueryDefinition,
  sourceNames: string[],
  oldRun?: DataQueryRun,
): NonNullable<DataTable["columns"]> {
  const lineage = queryColumnLineage(definition, sourceNames);
  if (!oldRun)
    return lineage.map((column, index) => ({
      id: index + 1,
      name: column.name,
    }));
  const oldLineage = queryColumnLineage(oldRun.definition, sourceNames);
  // A source column renamed in place is the same column (B355): a step that
  // names it fails above as a missing column, and a column the steps pass
  // through takes its new header. A source that gained or lost columns is
  // another shape.
  if (oldLineage.length !== oldRun.columns.length)
    fail("query_source_schema_changed");
  const byOrigin = new Map(
    oldLineage.map((column, index) => [
      column.origin,
      oldRun.columns[index].id,
    ]),
  );
  let nextId = Math.max(...oldRun.columns.map((column) => column.id)) + 1;
  return lineage.map((column) => ({
    id: byOrigin.get(column.origin) ?? nextId++,
    name: column.name,
  }));
}

/** One immutable candidate for Load/Refresh All. The UI must recheck the owning
 * workbook reference after await; the server revalidates independently at save. */
export async function planDataQueryRuns<T extends DataQuerySheet>(
  sheets: readonly T[],
  requests: readonly DataQueryRunRequest[],
): Promise<T[]> {
  if (!requests.length || requests.length > DATA_QUERY_LIMITS.queries)
    fail("query_limit");
  const targets = activeDataQueryRuns(sheets).map(
    (run) => `${run.targetSheetId}:${run.targetTableId}`,
  );
  const result = sheets.map((sheet) => ({
    ...sheet,
    cells: { ...sheet.cells },
    ...(sheet.dataTables ? { dataTables: [...sheet.dataTables] } : {}),
    ...(sheet.dataQueries
      ? {
          dataQueries: {
            active: { ...sheet.dataQueries.active },
            runs: { ...sheet.dataQueries.runs },
            ...(sheet.dataQueries.properties
              ? { properties: sheet.dataQueries.properties }
              : {}),
          },
        }
      : {}),
  }));
  const seen = new Set<string>();
  for (const request of requests) {
    if (seen.has(request.queryId)) fail("query_identity_ambiguous");
    seen.add(request.queryId);
    const definition = parseDataQueryDefinition(request.definition);
    const source = oneSheet(sheets, request.sourceSheetId);
    const sourceTable =
      request.createdSourceTable ?? oneTable(source, request.sourceTableId);
    if (request.createdSourceTable) {
      if (
        source.dataTables?.some(
          (table) =>
            table.id === sourceTable.id ||
            table.name.toLowerCase() === sourceTable.name.toLowerCase(),
        )
      )
        fail("query_identity_ambiguous");
      const nextSource = oneSheet(result, source.id);
      nextSource.dataTables = [...(nextSource.dataTables ?? []), sourceTable];
    }
    if (sourceTable.name !== definition.sourceName)
      fail("query_source_identity_changed");
    if (
      targets.includes(`${source.id}:${sourceTable.id}`) ||
      requests.some(
        (entry) =>
          entry.targetSheetId === source.id &&
          entry.targetTableId === sourceTable.id,
      )
    )
      fail("query_dependency_unsupported");
    const oldTarget = oneSheet(sheets, request.targetSheetId);
    const target = oneSheet(result, request.targetSheetId);
    if (oldTarget.protected || oldTarget.namedSheetViewsNativeBlocker)
      fail("query_target_protected");
    const state = target.dataQueries ?? { active: {}, runs: {} };
    const oldRun = state.runs[state.active[request.queryId]];
    const oldArea = oldRun ? dataQueryOutputRange(sheets, oldRun) : undefined;
    const liveTargetName = oldRun
      ? oneTable(oldTarget, oldRun.targetTableId).name
      : request.targetName;
    if (state.runs[request.runId]) fail("query_identity_ambiguous");
    if (
      oldRun &&
      (oldRun.targetTableId !== request.targetTableId ||
        liveTargetName !== request.targetName ||
        oldArea?.r1 !== request.row ||
        oldArea?.c1 !== request.col)
    )
      fail("query_target_identity_changed");
    if (
      oldRun &&
      (await queryFingerprint(rangeTokens(oldTarget, oldArea!))) !==
        oldRun.resultHash
    )
      fail("query_output_modified");
    const sourceData = readDataQuerySource(source, sourceTable);
    const preview = previewDataQuery(definition, sourceData);
    const outputColumns = queryOutputColumns(
      definition,
      sourceData.columns,
      oldRun,
    );
    const tokens = outputTokens(preview);
    // Excel's grid has no edge a query result can fall off; a result that
    // reaches past the sheet's current size grows the sheet, as pasting past
    // its last row does, up to the grid's own limits.
    target.rowCount = Math.min(
      MAX_GRID_ROWS,
      Math.max(target.rowCount, request.row + tokens.length),
    );
    target.columnCount = Math.min(
      MAX_GRID_COLUMNS,
      Math.max(target.columnCount, request.col + tokens[0].length),
    );
    const area = checkedRange(
      {
        r1: request.row,
        c1: request.col,
        r2: request.row + tokens.length - 1,
        c2: request.col + tokens[0].length - 1,
      },
      target,
    );
    for (const table of target.dataTables ?? []) {
      if (table.id === request.targetTableId) continue;
      const other = table.range;
      if (!(
        area.r2 < other.r1 ||
        area.r1 > other.r2 ||
        area.c2 < other.c1 ||
        area.c1 > other.c2
      ))
        fail("query_output_table_collision");
    }
    for (let row = area.r1; row <= area.r2; row++)
      for (let col = area.c1; col <= area.c2; col++) {
        const old = oldArea;
        if (
          !(
            old &&
            row >= old.r1 &&
            row <= old.r2 &&
            col >= old.c1 &&
            col <= old.c2
          ) &&
          cellToken(target.cells[`${row}:${col}`])[1] !== ""
        )
          fail("query_output_occupied");
      }
    const sourceRange = sourceTable.range;
    const run: DataQueryRun = {
      id: request.runId,
      queryId: request.queryId,
      definition,
      sourceSheetId: source.id,
      sourceTableId: sourceTable.id,
      targetSheetId: target.id,
      targetTableId: request.targetTableId,
      targetName: request.targetName,
      range: area,
      columns: outputColumns,
      priorRunId: oldRun?.id ?? null,
      sourceHash: await queryFingerprint([
        source.id,
        sourceTable.id,
        sourceTable.name,
        sourceTable.headerRow,
        [sourceRange.r1, sourceRange.c1, sourceRange.r2, sourceRange.c2],
        rangeTokens(source, sourceRange),
      ]),
      resultHash: await queryFingerprint(tokens),
      // The cells the result lands on, as they were: on the sheet before
      // this load, read within its grown size (a first load past the old
      // last column reads blanks there instead of failing).
      beforeHash: await queryFingerprint(
        rangeTokens(
          {
            ...oldTarget,
            rowCount: target.rowCount,
            columnCount: target.columnCount,
          },
          oldArea ?? area,
        ),
      ),
      ...(request.createdSourceTable
        ? { createdSourceTable: request.createdSourceTable }
        : {}),
    };
    if (oldArea)
      for (let row = oldArea.r1; row <= oldArea.r2; row++)
        for (let col = oldArea.c1; col <= oldArea.c2; col++) {
          const key = `${row}:${col}`,
            cell = target.cells[key];
          if (cell) {
            const cleared: QueryCell = { ...cell, value: "" };
            delete cleared.kind;
            delete cleared.display;
            delete cleared.chartValueCache;
            if (Object.keys(cleared).length === 1) delete target.cells[key];
            else target.cells[key] = cleared;
          }
        }
    tokens.forEach((row, rowIndex) =>
      row.forEach(([kind, value], colIndex) => {
        const key = `${area.r1 + rowIndex}:${area.c1 + colIndex}`;
        if (value) {
          const existing = target.cells[key];
          const written: Record<string, unknown> = {
            ...existing,
            value,
            kind: kind === "number" ? "number" : "text",
          };
          // A text scalar that starts with "=" stays the text it is (the
          // editor reads an unprefixed "=1+1" as a formula; live audit B277).
          const style = { ...((existing?.style as Record<string, unknown> | undefined) ?? {}) };
          // The words TRUE and FALSE keep their text identity the same way.
          if (kind !== "number" && (value.trim().startsWith("=") || value === "TRUE" || value === "FALSE")) style.quotePrefix = true;
          else delete style.quotePrefix;
          if (Object.keys(style).length > 0) written.style = style;
          else delete written.style;
          target.cells[key] = written as (typeof target.cells)[string];
        }
      }),
    );
    const oldTable = target.dataTables?.find(
      (table) => table.id === request.targetTableId,
    );
    const table: DataTable = {
      ...oldTable,
      id: request.targetTableId,
      name: request.targetName,
      range: area,
      headerRow: true,
      columns: outputColumns,
    };
    target.dataTables = [
      ...(target.dataTables ?? []).filter((entry) => entry.id !== table.id),
      table,
    ];
    state.active[request.queryId] = request.runId;
    state.runs[request.runId] = run;
    if (
      Object.keys(state.runs).length > 256 ||
      new TextEncoder().encode(JSON.stringify(state)).length > 1_000_000
    )
      fail("query_history_limit");
    if (Object.keys(state.active).length > DATA_QUERY_LIMITS.queries)
      fail("query_limit");
    target.dataQueries = state;
    parseDataQueryState(state);
  }
  const activeNames = activeDataQueryRuns(result).map((run) =>
    run.definition.name.toLowerCase(),
  );
  if (new Set(activeNames).size !== activeNames.length)
    fail("query_name_collision");
  if (
    result.some((sheet) => (sheet.dataTables?.length ?? 0) > MAX_DATA_TABLES) ||
    activeDataQueryRuns(result).length > DATA_QUERY_LIMITS.queries ||
    new TextEncoder().encode(
      JSON.stringify(result.map((sheet) => sheet.dataQueries ?? null)),
    ).length > 2_000_000
  )
    fail("query_history_limit");
  return result;
}
export class DataQueryError extends Error {
  constructor(public readonly code: string) {
    super(code);
  }
}
function fail(code: string): never {
  throw new DataQueryError(code);
}
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value))
    fail("query_invalid");
  return value as Record<string, unknown>;
}
function keys(value: Record<string, unknown>, allowed: string[]) {
  if (
    Object.keys(value).some((key) => !allowed.includes(key)) ||
    allowed.some((key) => !Object.hasOwn(value, key))
  )
    fail("query_unsupported");
}
function label(value: unknown, limit = 255): string {
  if (
    typeof value !== "string" ||
    !value.trim() ||
    value.length > limit ||
    /[\u0000-\u001f\u007f]/u.test(value) ||
    /[\ud800-\udfff]/u.test(value)
  )
    fail("query_name_invalid");
  return value;
}
function scalar(value: unknown): QueryScalar {
  if (value === null || typeof value === "boolean") return value;
  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    Math.abs(value) <= Number.MAX_SAFE_INTEGER
  )
    return value === 0 ? 0 : value;
  if (
    typeof value === "string" &&
    value.length <= DATA_QUERY_LIMITS.text &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff]/u.test(value)
  )
    return value;
  return fail("query_scalar_unsupported");
}
function columnList(value: unknown): string[] {
  if (
    !Array.isArray(value) ||
    !value.length ||
    value.length > DATA_QUERY_LIMITS.columns
  )
    fail("query_columns_invalid");
  const result = value.map((name) => label(name));
  if (new Set(result.map((name) => name.toLowerCase())).size !== result.length)
    fail("query_columns_invalid");
  return result;
}
export function parseDataQueryDefinition(input: unknown): DataQueryDefinition {
  const value = record(input);
  keys(value, ["version", "name", "sourceName", "steps"]);
  if (
    value.version !== 1 ||
    !Array.isArray(value.steps) ||
    value.steps.length > DATA_QUERY_LIMITS.steps
  )
    fail("query_unsupported");
  const name = label(value.name);
  // MS-QDEFF 2.3.2 section-member restrictions, including leading/trailing space.
  if (name.trim() !== name || /[.";]/u.test(name)) fail("query_name_invalid");
  const sourceName = label(value.sourceName);
  const steps = value.steps.map((inputStep): DataQueryStep => {
    const step = record(inputStep);
    switch (step.kind) {
      case "selectColumns":
      case "distinct":
        keys(step, ["kind", "columns"]);
        return { kind: step.kind, columns: columnList(step.columns) };
      case "renameColumn":
        keys(step, ["kind", "from", "to"]);
        return { kind: step.kind, from: label(step.from), to: label(step.to) };
      case "filterEquals":
        keys(step, ["kind", "column", "value"]);
        return {
          kind: step.kind,
          column: label(step.column),
          value: scalar(step.value),
        };
      case "removeBlankRows":
        keys(step, ["kind"]);
        return { kind: step.kind };
      default:
        return fail("query_step_unsupported");
    }
  });
  return { version: 1, name, sourceName, steps };
}

/** Returns new arrays. Missing columns/error cells/type coercions never become blanks. */
export function previewDataQuery(
  definition: unknown,
  source: DataQueryTable,
): DataQueryTable {
  const query = parseDataQueryDefinition(definition);
  keys(record(source), ["columns", "rows"]);
  let columns = columnList(source.columns);
  if (
    !Array.isArray(source.rows) ||
    (source.rows.length + 1) * columns.length > DATA_QUERY_LIMITS.cells
  )
    fail("query_limit");
  let totalText = 0;
  let rows = source.rows.map((row) => {
    if (!Array.isArray(row) || row.length !== columns.length)
      fail("query_row_invalid");
    return row.map((value) => {
      const result = scalar(value);
      if (typeof result === "string") totalText += result.length;
      if (totalText > DATA_QUERY_LIMITS.totalText) fail("query_limit");
      return result;
    });
  });
  const indices = (names: string[]) =>
    names.map((name) => {
      const index = columns.indexOf(name);
      if (index < 0) fail("query_column_missing");
      return index;
    });
  for (const step of query.steps) {
    switch (step.kind) {
      case "selectColumns": {
        const picked = indices(step.columns);
        rows = rows.map((row) => picked.map((index) => row[index]));
        columns = [...step.columns];
        break;
      }
      case "renameColumn": {
        const [index] = indices([step.from]);
        if (
          columns.some(
            (name, position) =>
              position !== index &&
              name.toLowerCase() === step.to.toLowerCase(),
          )
        )
          fail("query_column_collision");
        columns[index] = step.to;
        break;
      }
      case "filterEquals": {
        const [index] = indices([step.column]);
        rows = rows.filter((row) => row[index] === step.value);
        break;
      }
      case "removeBlankRows":
        rows = rows.filter((row) =>
          row.some((value) => value !== null && value !== ""),
        );
        break;
      case "distinct": {
        const picked = indices(step.columns);
        const seen = new Set<string>();
        rows = rows.filter((row) => {
          const key = JSON.stringify(picked.map((index) => row[index]));
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        });
      }
    }
  }
  return { columns, rows };
}

/** Plan all results before any caller writes. V1 rejects query-output sources,
 * including cycles; it cannot accidentally refresh them in iteration order. */
export function previewDataQueryBatch(
  definitions: unknown[],
  sources: ReadonlyMap<string, DataQueryTable>,
  destinationNames: string[],
): DataQueryTable[] {
  if (
    !Array.isArray(definitions) ||
    !definitions.length ||
    definitions.length > DATA_QUERY_LIMITS.queries ||
    destinationNames.length !== definitions.length
  )
    fail("query_limit");
  const queries = definitions.map(parseDataQueryDefinition);
  const names = queries.map((query) => query.name.toLowerCase());
  const destinations = destinationNames.map((name) =>
    label(name).toLowerCase(),
  );
  if (
    new Set(names).size !== names.length ||
    new Set(destinations).size !== destinations.length
  )
    fail("query_identity_ambiguous");
  return queries.map((query) => {
    if (destinations.includes(query.sourceName.toLowerCase()))
      fail("query_dependency_unsupported");
    const source = sources.get(query.sourceName);
    if (!source) fail("query_source_missing");
    return previewDataQuery(query, source);
  });
}

/** M escaping is not JSON escaping. Escape # first to prevent injected #(lf). */
function mText(value: string): string {
  return `"${value
    .replaceAll("#", "#(#)")
    .replaceAll('"', '""')
    .replaceAll("\r", "#(cr)")
    .replaceAll("\n", "#(lf)")
    .replaceAll("\t", "#(tab)")}"`;
}
function mValue(value: QueryScalar): string {
  return typeof value === "string" ? mText(value) : String(value);
}
function mColumns(columns: string[]): string {
  return `{${columns.map(mText).join(", ")}}`;
}

// The step names Excel's Power Query Editor gives these transforms; a repeated
// one takes the next number ("Filtered Rows1"). Same text as the server's
// serialize_query_m (data_queries.py), which writes the workbook's query.
const EXCEL_STEP_NAMES: Record<DataQueryStep["kind"], string> = {
  selectColumns: "Removed Other Columns",
  renameColumn: "Renamed Columns",
  filterEquals: "Filtered Rows",
  removeBlankRows: "Removed Blank Rows",
  distinct: "Removed Duplicates",
};
const M_KEYWORDS = new Set([
  "and", "as", "each", "else", "error", "false", "if", "in", "is", "let", "meta",
  "not", "null", "or", "otherwise", "section", "shared", "then", "true", "try", "type",
]);
function mName(value: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(value) && !M_KEYWORDS.has(value) ? value : `#${mText(value)}`;
}
function mField(value: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*(?: [A-Za-z0-9_]+)*$/.test(value) && !M_KEYWORDS.has(value)
    ? `[${value}]`
    : `[#${mText(value)}]`;
}

/** Native closed M recipe, written as Excel's editor writes the same steps. */
export function serializeDataQueryM(input: unknown): string {
  const query = parseDataQueryDefinition(input);
  let previous = "Source";
  const used = new Set(["Source"]);
  const expressions = [
    `Source = Excel.CurrentWorkbook(){[Name=${mText(query.sourceName)}]}[Content]`,
  ];
  query.steps.forEach((step) => {
    const here = mName(previous);
    let expression: string;
    switch (step.kind) {
      case "selectColumns":
        expression = `Table.SelectColumns(${here},${mColumns(step.columns)})`;
        break;
      case "renameColumn":
        expression = `Table.RenameColumns(${here},{{${mText(step.from)}, ${mText(step.to)}}})`;
        break;
      case "filterEquals":
        expression = `Table.SelectRows(${here}, each (${mField(step.column)} = ${mValue(step.value)}))`;
        break;
      case "removeBlankRows":
        expression = `Table.SelectRows(${here}, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))`;
        break;
      case "distinct":
        expression = `Table.Distinct(${here}, ${mColumns(step.columns)})`;
        break;
    }
    const base = EXCEL_STEP_NAMES[step.kind];
    let name = base;
    for (let suffix = 1; used.has(name); suffix += 1) name = `${base}${suffix}`;
    used.add(name);
    previous = name;
    expressions.push(`${mName(name)} = ${expression}`);
  });
  return `shared ${mName(query.name)} = let\n    ${expressions.join(",\n    ")}\nin\n    ${mName(previous)};`;
}
