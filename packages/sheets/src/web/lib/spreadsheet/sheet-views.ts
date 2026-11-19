/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizeFilterDateGroups,
  type FilterDateGroup,
} from "@mrnote/core/web/lib/spreadsheet/date-filter-tree.ts";
import { iconSetCount } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import {
  normalizeDataTableColumnFilter,
  normalizeDataTableIconFilter,
  normalizeDataTableSortState,
  type DataTableColumn,
  type DataTableColumnFilter,
  type DataTableIconFilter,
  type DataTableSortState,
} from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import type { ChineseCollation } from "@mrnote/core/web/lib/spreadsheet/locale-sort.ts";
import {
  compareCellsAtLevel,
  type SortDisplayKeyCell,
  type SortLevel,
} from "./multi-sort";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { columnHeaderLabel } from "@mrnote/core/web/lib/spreadsheet/r1c1.ts";
import {
  normalizeFillColor,
  type SortFilterCell,
  type SpreadsheetTextOrder,
} from "./sort-filter";
import {
  normalizeConditionFilter,
  type ConditionFilter,
} from "@mrnote/core/web/lib/spreadsheet/filter-conditions.ts";

export const MAX_NAMED_SHEET_VIEWS = 256;
export const MAX_NAMED_SHEET_VIEW_NAME = 127;
export const NAMED_SHEET_VIEW_NATIVE_BLOCKER =
  "named_sheet_view_native_graph_unsupported";
/** One nsvFilter per AutoFilter; a sheet holds at most this many tables. */
export const MAX_SHEET_VIEW_TABLES = 64;
/** Excel's custom-list bound for one sort key (the entries of one list). */
export const MAX_SHEET_VIEW_CUSTOM_LIST_ENTRIES = 1_000;

/**
 * A Table AutoFilter's state inside one Sheet View ([MS-XLSX] nsvFilter with
 * a non-zero tableId). The criteria and sort keys use exactly the Table's own
 * grammar (keyed by stable tableColumn ids), so a view holds whatever the
 * Table's arrow can set. `filterId` is that table AutoFilter's `xr:uid`; a
 * definition read from or written to a file always carries it, the in-session
 * snapshot of the default view may not.
 */
export interface SpreadsheetSheetViewTableState {
  tableId: string;
  filterId?: string;
  columnFilters?: Record<string, DataTableColumnFilter>;
  sortState?: DataTableSortState;
}

export interface SpreadsheetNamedSheetView {
  /** Native xsv:namedSheetView GUID, retained across rename and duplicate. */
  id: string;
  name: string;
  /**
   * Native CT_NsvFilter identity. An empty view has neither field. The bounded
   * authoring subset accepts zero or one nsvFilter; filterId is required when
   * the filter exists, while ref is optional in OOXML and is retained only
   * when present. Column-filter keys below are absolute worksheet indexes.
   */
  filterId?: string;
  ref?: string;
  filters?: Record<string, string[]>;
  /**
   * Excel's date-tree criteria (year / month / day ... prefixes) per column,
   * beside or instead of the literal values in `filters`.
   */
  dateGroups?: Record<string, FilterDateGroup[]>;
  conditionFilters?: Record<string, ConditionFilter>;
  colorFilters?: Record<string, string[]>;
  sortState?: SpreadsheetSheetViewSortState;
  /** The view's Table AutoFilter states, one entry per filtered/sorted table. */
  tables?: SpreadsheetSheetViewTableState[];
  /** Unsupported native graphs stay visible but immutable. */
  nativeBlocker?: string;
}

/**
 * Replace, rather than merge, a view's mutable nsvFilter payload. This keeps a
 * filter-only edit from retaining an old sort (and vice versa). An existing
 * nsvFilter association survives with no criteria; a truly empty `{id,name}`
 * view does not acquire an AutoFilter merely because it was opened. A Table's
 * nsvFilter exists exactly while the view filters or sorts that table.
 */
export function rebuildNamedSheetViewState(
  view: SpreadsheetNamedSheetView,
  state: SpreadsheetSheetViewFilterState,
  filterId?: string,
  ref?: string,
): SpreadsheetNamedSheetView {
  const identity: SpreadsheetNamedSheetView = { ...view };
  delete identity.filterId;
  delete identity.ref;
  delete identity.filters;
  delete identity.dateGroups;
  delete identity.conditionFilters;
  delete identity.colorFilters;
  delete identity.sortState;
  delete identity.tables;
  const { tables: rawTables, ...snapshot } = snapshotSheetViewFilters(state);
  const hasState = Boolean(
    snapshot.sortState || sheetViewStateHasFilters(snapshot),
  );
  const retainAssociation = Boolean(filterId && (view.filterId || hasState));
  const tables = (rawTables ?? []).flatMap((table) => {
    const filterIdForTable =
      table.filterId ??
      view.tables?.find((candidate) => candidate.tableId === table.tableId)
        ?.filterId;
    return filterIdForTable && sheetViewTableStateHasCriteria(table)
      ? [{ ...table, filterId: filterIdForTable }]
      : [];
  });
  const base =
    retainAssociation && filterId
      ? {
          ...identity,
          filterId,
          ...(ref ? { ref } : {}),
          ...snapshot,
        }
      : identity;
  return tables.length > 0 ? { ...base, tables } : base;
}

/** Whether a Table's view state filters or sorts it at all. */
export function sheetViewTableStateHasCriteria(
  table: SpreadsheetSheetViewTableState,
): boolean {
  return Boolean(
    table.sortState || Object.keys(table.columnFilters ?? {}).length > 0,
  );
}

/**
 * What one Sheet View holds, per [MS-XLSX] 2.6 Named Sheet Views: "sort
 * and/or filter settings that are associated with an AutoFilter". A
 * columnFilter carries any AutoFilter criterion, so a view keeps the
 * column's value list, custom/dynamic condition, or cell-colour filter; the
 * sortRules keep its value sort. Cell content is never part of a view.
 */
export interface SpreadsheetSheetViewFilterState {
  filters?: Record<string, string[]>;
  dateGroups?: Record<string, FilterDateGroup[]>;
  /**
   * In-session only: the display values each column's date groups resolved
   * to when the view was applied, so a snapshot can tell the view's own
   * date criteria from a selection the user changed.
   */
  dateGroupValues?: Record<string, string[]>;
  conditionFilters?: Record<string, ConditionFilter>;
  colorFilters?: Record<string, string[]>;
  sortState?: SpreadsheetSheetViewSortState;
  /**
   * The Table AutoFilters' state. Present (even empty) on a whole-sheet
   * snapshot, so applying it also clears a table the state does not filter;
   * absent on a worksheet-only change, which leaves the tables alone.
   */
  tables?: SpreadsheetSheetViewTableState[];
}

/** Whether the WORKSHEET AutoFilter part of a state holds any criterion. */
export function sheetViewStateHasFilters(
  state: SpreadsheetSheetViewFilterState,
): boolean {
  return (
    Object.keys(state.filters ?? {}).length > 0 ||
    Object.keys(state.dateGroups ?? {}).length > 0 ||
    Object.keys(state.conditionFilters ?? {}).length > 0 ||
    Object.keys(state.colorFilters ?? {}).length > 0
  );
}

/** Excel's Sort > Sort On for one key: cell values or a displayed format. */
export type SpreadsheetSheetViewSortBy = "cellColor" | "fontColor" | "icon";

/**
 * One sort key of a Sheet View (an [MS-XLSX] sortRule). A value key orders by
 * the cells (optionally through a custom list); a format key brings one cell
 * colour, font colour or conditional-format icon to the top (`descending`
 * absent) or to the bottom (`descending: true`), Excel's "On Top / On Bottom".
 */
export interface SpreadsheetSheetViewSortCondition {
  ref: string;
  descending?: boolean;
  sortBy?: SpreadsheetSheetViewSortBy;
  /** Lower-case "#rrggbb", present exactly for a cell/font colour key. */
  color?: string;
  /** The icon, present exactly for an icon key. */
  icon?: DataTableIconFilter;
  /** Order > Custom List: the list, for a value key only. */
  customList?: string[];
}

export interface SpreadsheetSheetViewSortState {
  ref: string;
  conditions: SpreadsheetSheetViewSortCondition[];
  /** Sort Options > Case sensitive (one switch for every key, as in Excel). */
  caseSensitive?: true;
  /** Sort Options > Method for Chinese text: pinyin (letter) or stroke. */
  sortMethod?: "pinYin" | "stroke";
}

export interface SpreadsheetSheetViewSource
  extends SpreadsheetSheetViewFilterState {
  rowCount: number;
  columnCount: number;
  dataTables?: Array<{
    columnFilters?: Record<string, unknown>;
    sortState?: unknown;
  }>;
  advancedFilterState?: unknown;
  namedSheetViewsNativeBlocker?: unknown;
}

// ST_Guid is a lexical GUID, not an RFC-4122 version/variant receipt. Excel
// legitimately emits IDs outside the UUID version/variant bit pattern.
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

export function normalizeNamedSheetViewId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  if (value !== value.trim()) return null;
  const normalized = value.toUpperCase();
  return GUID.test(normalized) ? normalized : null;
}

export function createNamedSheetViewId(): string {
  return `{${crypto.randomUUID().toUpperCase()}}`;
}

/**
 * The name Keep gives a temporary view: Excel's "View1", "View2", ... The
 * first number whose name no view on the sheet already carries (view names are
 * unique without regard to case). `format` supplies the localized pattern.
 */
export function nextDefaultNamedSheetViewName(
  views: readonly { name: string }[],
  format: (index: number) => string,
): string {
  const taken = new Set(views.map((view) => view.name.toLocaleLowerCase()));
  for (let index = 1; index <= views.length; index += 1) {
    const name = format(index);
    if (!taken.has(name.toLocaleLowerCase())) return name;
  }
  return format(views.length + 1);
}

export function normalizeNamedSheetViewName(value: unknown): string | null {
  if (typeof value !== "string") return null;
  if (
    value !== value.trim() ||
    value.length === 0 ||
    Array.from(value).length > MAX_NAMED_SHEET_VIEW_NAME ||
    /[\u0000-\u001F\u007F]/.test(value) ||
    /[^\S ]/u.test(value) ||
    value.includes(";") ||
    /^_xlnsv\./i.test(value)
  ) {
    return null;
  }
  return value;
}

function normalizeValueFilters(
  value: unknown,
  columnCount: number,
): Record<string, string[]> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const normalized: Record<string, string[]> = {};
  for (const [key, rawValues] of Object.entries(value)) {
    const column = Number(key);
    if (
      !Number.isInteger(column) ||
      column < 0 ||
      column >= columnCount ||
      !Array.isArray(rawValues) ||
      rawValues.length === 0 ||
      rawValues.length > 10_000
    ) {
      return undefined;
    }
    const values: string[] = [];
    for (const rawValue of rawValues) {
      if (typeof rawValue !== "string" || rawValue.length > 1_000) {
        return undefined;
      }
      values.push(rawValue);
    }
    if (new Set(values).size !== values.length) return undefined;
    normalized[String(column)] = values;
  }
  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function normalizeColorFilters(
  value: unknown,
  columnCount: number,
): Record<string, string[]> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const normalized: Record<string, string[]> = {};
  for (const [key, rawValues] of Object.entries(value)) {
    const column = Number(key);
    if (
      !Number.isInteger(column) ||
      column < 0 ||
      column >= columnCount ||
      !Array.isArray(rawValues) ||
      rawValues.length === 0 ||
      rawValues.length > 256
    ) {
      return undefined;
    }
    const values: string[] = [];
    for (const rawValue of rawValues) {
      if (typeof rawValue !== "string" || rawValue.length > 64) return undefined;
      const color = normalizeFillColor(rawValue);
      if (!values.includes(color)) values.push(color);
    }
    normalized[String(column)] = values;
  }
  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function normalizeConditionFilters(
  value: unknown,
  columnCount: number,
): Record<string, ConditionFilter> | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const normalized: Record<string, ConditionFilter> = {};
  for (const [key, raw] of Object.entries(value)) {
    const column = Number(key);
    const condition = normalizeConditionFilter(raw);
    if (
      !Number.isInteger(column) ||
      column < 0 ||
      column >= columnCount ||
      !condition
    ) {
      return undefined;
    }
    normalized[String(column)] = condition;
  }
  return Object.keys(normalized).length > 0 ? normalized : undefined;
}

function rangeA1(bounds: {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}): string {
  const from = `${columnHeaderLabel(bounds.c1, "a1")}${bounds.r1 + 1}`;
  const to = `${columnHeaderLabel(bounds.c2, "a1")}${bounds.r2 + 1}`;
  return `${from}:${to}`;
}

export function normalizeSheetViewSortState(
  raw: unknown,
  viewRef: string,
  rowCount: number,
  columnCount: number,
): SpreadsheetSheetViewSortState | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const view = parseA1Range(viewRef.replace(/\$/g, ""));
  const sort =
    typeof record.ref === "string"
      ? parseA1Range(record.ref.trim().toUpperCase().replace(/\$/g, ""))
      : null;
  if (
    !view ||
    !sort ||
    sort.r1 < view.r1 ||
    sort.c1 < view.c1 ||
    sort.r2 > view.r2 ||
    sort.c2 > view.c2 ||
    sort.r2 >= rowCount ||
    sort.c2 >= columnCount
  ) {
    return undefined;
  }
  const dataR1 = sort.r1 === view.r1 ? sort.r1 + 1 : sort.r1;
  if (dataR1 > sort.r2 || !Array.isArray(record.conditions)) return undefined;
  if (record.conditions.length < 1 || record.conditions.length > 64) {
    return undefined;
  }
  if (
    (record.caseSensitive !== undefined && record.caseSensitive !== true) ||
    (record.sortMethod !== undefined &&
      record.sortMethod !== "pinYin" &&
      record.sortMethod !== "stroke")
  ) {
    return undefined;
  }
  const identities = new Set<string>();
  const conditions: SpreadsheetSheetViewSortCondition[] = [];
  for (const rawCondition of record.conditions) {
    if (
      !rawCondition ||
      typeof rawCondition !== "object" ||
      Array.isArray(rawCondition)
    ) {
      return undefined;
    }
    const condition = rawCondition as Record<string, unknown>;
    if (
      typeof condition.ref !== "string" ||
      (condition.descending !== undefined &&
        typeof condition.descending !== "boolean")
    ) {
      return undefined;
    }
    const bounds = parseA1Range(
      condition.ref.trim().toUpperCase().replace(/\$/g, ""),
    );
    if (
      !bounds ||
      bounds.c1 !== bounds.c2 ||
      bounds.c1 < sort.c1 ||
      bounds.c1 > sort.c2 ||
      bounds.r1 !== dataR1 ||
      bounds.r2 !== sort.r2
    ) {
      return undefined;
    }
    const key = normalizeSheetViewSortKey(condition);
    if (!key) return undefined;
    const identity = sheetViewSortKeyIdentity(bounds.c1, key);
    if (identities.has(identity)) return undefined;
    identities.add(identity);
    conditions.push({
      ref: rangeA1(bounds),
      ...(condition.descending === true ? { descending: true } : {}),
      ...key,
    });
  }
  return {
    ref: rangeA1({ ...sort, r1: dataR1 }),
    conditions,
    ...(record.caseSensitive === true ? { caseSensitive: true as const } : {}),
    ...(record.sortMethod === "pinYin" || record.sortMethod === "stroke"
      ? { sortMethod: record.sortMethod }
      : {}),
  };
}

/**
 * What makes two sort levels the same level. Excel's Sort dialog takes a
 * column once for values, but once per colour or icon for a format key:
 * "Sort by cell color, font color or icon" adds the same column again for
 * every colour that should come first (live audit B474). A repeat of the
 * same identity cannot change the order, and Excel refuses it.
 */
function sheetViewSortKeyIdentity(
  column: number,
  key: Omit<SpreadsheetSheetViewSortCondition, "ref" | "descending">,
): string {
  if (key.sortBy === "cellColor" || key.sortBy === "fontColor") {
    return `${column}:${key.sortBy}:${key.color ?? ""}`;
  }
  if (key.sortBy === "icon") {
    return `${column}:icon:${key.icon?.iconSet ?? ""}:${key.icon?.iconId ?? ""}`;
  }
  return `${column}:value`;
}

/**
 * The Sort On part of one view sort key, or null when it is malformed. A value
 * key may carry a custom list; a colour key carries exactly one "#rrggbb"; an
 * icon key carries exactly one standard icon. Nothing else rides along.
 */
function normalizeSheetViewSortKey(
  condition: Record<string, unknown>,
): Omit<SpreadsheetSheetViewSortCondition, "ref" | "descending"> | null {
  const sortBy = condition.sortBy;
  if (sortBy === undefined) {
    if (condition.color !== undefined || condition.icon !== undefined) {
      return null;
    }
    if (condition.customList === undefined) return {};
    const list = normalizeSheetViewCustomList(condition.customList);
    return list ? { customList: list } : null;
  }
  if (condition.customList !== undefined) return null;
  if (sortBy === "cellColor" || sortBy === "fontColor") {
    if (
      condition.icon !== undefined ||
      typeof condition.color !== "string" ||
      !/^#[0-9a-f]{6}$/i.test(condition.color)
    ) {
      return null;
    }
    return { sortBy, color: condition.color.toLowerCase() };
  }
  if (sortBy === "icon") {
    const icon = normalizeDataTableIconFilter(condition.icon);
    return icon && condition.color === undefined ? { sortBy, icon } : null;
  }
  return null;
}

/**
 * One custom list as a view sort key holds it: 1..1000 non-empty entries
 * without a comma (OOXML stores the list as one comma-separated string).
 */
export function normalizeSheetViewCustomList(raw: unknown): string[] | null {
  if (
    !Array.isArray(raw) ||
    raw.length === 0 ||
    raw.length > MAX_SHEET_VIEW_CUSTOM_LIST_ENTRIES
  ) {
    return null;
  }
  const list: string[] = [];
  for (const entry of raw) {
    if (
      typeof entry !== "string" ||
      entry.trim() === "" ||
      entry.length > 255 ||
      entry.includes(",")
    ) {
      return null;
    }
    list.push(entry);
  }
  return list;
}

function normalizeSheetViewTables(
  raw: unknown,
  knownTables?: ReadonlyArray<SheetViewTableShape>,
): SpreadsheetSheetViewTableState[] | null {
  if (!Array.isArray(raw) || raw.length > MAX_SHEET_VIEW_TABLES) return null;
  const seen = new Set<string>();
  const out: SpreadsheetSheetViewTableState[] = [];
  for (const candidate of raw) {
    if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
      return null;
    }
    const record = candidate as Record<string, unknown>;
    const tableId = record.tableId;
    const filterId = normalizeNamedSheetViewId(record.filterId);
    if (
      typeof tableId !== "string" ||
      tableId.length === 0 ||
      tableId.length > 200 ||
      seen.has(tableId) ||
      !filterId ||
      Object.keys(record).some(
        (key) =>
          !["tableId", "filterId", "columnFilters", "sortState"].includes(key),
      )
    ) {
      return null;
    }
    const known = knownTables?.find((table) => table.id === tableId);
    if (knownTables && (!known || !known.headerRow || !known.columns)) {
      return null;
    }
    let columnFilters: Record<string, DataTableColumnFilter> | undefined;
    if (record.columnFilters !== undefined) {
      if (
        !record.columnFilters ||
        typeof record.columnFilters !== "object" ||
        Array.isArray(record.columnFilters)
      ) {
        return null;
      }
      for (const [key, rawFilter] of Object.entries(
        record.columnFilters as Record<string, unknown>,
      )) {
        const columnId = Number(key);
        const filter = normalizeDataTableColumnFilter(rawFilter);
        if (
          !Number.isInteger(columnId) ||
          String(columnId) !== key ||
          !filter ||
          (known?.columns &&
            !known.columns.some((column) => column.id === columnId))
        ) {
          return null;
        }
        (columnFilters ??= {})[key] = filter;
      }
    }
    let sortState: DataTableSortState | undefined;
    if (record.sortState !== undefined) {
      const columns: DataTableColumn[] =
        known?.columns ??
        (Array.isArray((record.sortState as { conditions?: unknown })?.conditions)
          ? (
              (record.sortState as { conditions: unknown[] }).conditions as Array<{
                columnId?: unknown;
              }>
            ).flatMap((condition) =>
              typeof condition?.columnId === "number"
                ? [{ id: condition.columnId, name: String(condition.columnId) }]
                : [],
            )
          : []);
      const normalized = normalizeDataTableSortState(record.sortState, columns);
      if (!normalized) return null;
      sortState = normalized;
    }
    seen.add(tableId);
    const state: SpreadsheetSheetViewTableState = {
      tableId,
      filterId,
      ...(columnFilters ? { columnFilters } : {}),
      ...(sortState ? { sortState } : {}),
    };
    // An nsvFilter with no criteria (`<nsvFilter filterId=.. tableId=../>`)
    // is how Excel writes a view that shows that Table unfiltered: it stays,
    // so loading and saving a workbook leaves its namedSheetView part alone
    
    out.push(state);
  }
  return out;
}

/**
 * Keep each view's Table entries in step with the sheet's Tables: an entry
 * for a removed Table (or one that lost its header row) goes, and so do the
 * criteria and sort keys of columns the Table no longer has.
 */
export function pruneSheetViewTables(
  views: SpreadsheetNamedSheetView[],
  tables: ReadonlyArray<SheetViewTableShape>,
): SpreadsheetNamedSheetView[] {
  return views.map((view) => {
    if (!view.tables) return view;
    const kept = view.tables.flatMap((entry) => {
      const table = tables.find((candidate) => candidate.id === entry.tableId);
      if (!table?.headerRow || !table.columns) return [];
      const ids = new Set(table.columns.map((column) => column.id));
      const columnFilters = Object.fromEntries(
        Object.entries(entry.columnFilters ?? {}).filter(([key]) =>
          ids.has(Number(key)),
        ),
      );
      const conditions = (entry.sortState?.conditions ?? []).filter((condition) =>
        ids.has(condition.columnId),
      );
      const next: SpreadsheetSheetViewTableState = {
        tableId: entry.tableId,
        ...(entry.filterId ? { filterId: entry.filterId } : {}),
        ...(Object.keys(columnFilters).length > 0 ? { columnFilters } : {}),
        ...(conditions.length > 0 ? { sortState: { conditions } } : {}),
      };
      return sheetViewTableStateHasCriteria(next) ||
        !sheetViewTableStateHasCriteria(entry)
        ? [next]
        : [];
    });
    if (
      kept.length === view.tables.length &&
      kept.every((entry, index) => JSON.stringify(entry) === JSON.stringify(view.tables![index]))
    ) {
      return view;
    }
    const next = { ...view };
    if (kept.length > 0) next.tables = kept;
    else delete next.tables;
    return next;
  });
}

/** What the view normalizer needs to know about a sheet's Tables. */
export interface SheetViewTableShape {
  id: string;
  headerRow: boolean;
  columns?: DataTableColumn[];
  range?: { r1: number; c1: number; r2: number; c2: number };
}

export function normalizeNamedSheetViews(
  raw: unknown,
  rowCount: number,
  columnCount: number,
  autoFilterRef?: string,
  /** The sheet's Tables; when given, a view's table entries must name them. */
  knownTables?: ReadonlyArray<SheetViewTableShape>,
): SpreadsheetNamedSheetView[] {
  if (!Array.isArray(raw)) return [];
  const result: SpreadsheetNamedSheetView[] = [];
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const candidate of raw.slice(0, MAX_NAMED_SHEET_VIEWS)) {
    if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
      continue;
    }
    const record = candidate as Record<string, unknown>;
    const id = normalizeNamedSheetViewId(record.id);
    const name = normalizeNamedSheetViewName(record.name);
    const filterId =
      record.filterId === undefined
        ? undefined
        : normalizeNamedSheetViewId(record.filterId);
    const rawRef =
      typeof record.ref === "string" && record.ref === record.ref.trim()
        ? record.ref.toUpperCase()
        : undefined;
    const effectiveRef = rawRef ?? autoFilterRef;
    const range = effectiveRef
      ? parseA1Range(effectiveRef.replace(/\$/g, ""))
      : null;
    const hasFilterPayload =
      record.filterId !== undefined ||
      record.ref !== undefined ||
      record.filters !== undefined ||
      record.dateGroups !== undefined ||
      record.conditionFilters !== undefined ||
      record.colorFilters !== undefined ||
      record.sortState !== undefined;
    if (
      !id ||
      !name ||
      (hasFilterPayload && (!filterId || !range)) ||
      (range !== null &&
        (range.r2 >= rowCount || range.c2 >= columnCount)) ||
      ids.has(id) ||
      names.has(name.toLocaleLowerCase())
    ) {
      continue;
    }
    const filters = hasFilterPayload
      ? normalizeValueFilters(record.filters, columnCount)
      : undefined;
    if (
      record.filters !== undefined &&
      (!record.filters ||
        typeof record.filters !== "object" ||
        Array.isArray(record.filters) ||
        (Object.keys(record.filters).length > 0 && !filters))
    ) {
      continue;
    }
    let dateGroups: Record<string, FilterDateGroup[]> | undefined;
    let dateGroupsInvalid = false;
    if (record.dateGroups !== undefined) {
      if (!record.dateGroups || typeof record.dateGroups !== "object" || Array.isArray(record.dateGroups)) {
        dateGroupsInvalid = true;
      } else {
        for (const [key, rawGroups] of Object.entries(record.dateGroups as Record<string, unknown>)) {
          const column = Number(key);
          const groups = normalizeFilterDateGroups(rawGroups);
          if (!Number.isInteger(column) || column < 0 || column >= columnCount || String(column) !== key || !groups) {
            dateGroupsInvalid = true;
            break;
          }
          (dateGroups ??= {})[key] = groups;
        }
      }
    }
    if (dateGroupsInvalid) continue;
    const conditionFilters = hasFilterPayload
      ? normalizeConditionFilters(record.conditionFilters, columnCount)
      : undefined;
    if (record.conditionFilters !== undefined && !conditionFilters) continue;
    const colorFilters = hasFilterPayload
      ? normalizeColorFilters(record.colorFilters, columnCount)
      : undefined;
    if (record.colorFilters !== undefined && !colorFilters) continue;
    const outsideRange = (key: string) => {
      const column = Number(key);
      return !range || column < range.c1 || column > range.c2;
    };
    if (
      Object.keys(filters ?? {}).some(outsideRange) ||
      Object.keys(conditionFilters ?? {}).some(outsideRange) ||
      Object.keys(colorFilters ?? {}).some(outsideRange) ||
      // A column holds one criterion kind, as one AutoFilter filterColumn does.
      Object.keys(conditionFilters ?? {}).some((key) => filters?.[key])
    ) {
      continue;
    }
    const sortState = effectiveRef
      ? normalizeSheetViewSortState(
          record.sortState,
          effectiveRef,
          rowCount,
          columnCount,
        )
      : undefined;
    if (record.sortState !== undefined && !sortState) continue;
    const tables =
      record.tables === undefined
        ? undefined
        : normalizeSheetViewTables(record.tables, knownTables);
    if (record.tables !== undefined && !tables) continue;
    const nativeBlocker =
      typeof record.nativeBlocker === "string" && record.nativeBlocker
        ? record.nativeBlocker.slice(0, 255)
        : undefined;
    ids.add(id);
    names.add(name.toLocaleLowerCase());
    result.push({
      id,
      name,
      ...(filterId ? { filterId } : {}),
      ...(rawRef ? { ref: rawRef } : {}),
      ...(filters ? { filters } : {}),
      ...(dateGroups ? { dateGroups } : {}),
      ...(conditionFilters ? { conditionFilters } : {}),
      ...(colorFilters ? { colorFilters } : {}),
      ...(sortState ? { sortState } : {}),
      ...(tables && tables.length > 0 ? { tables } : {}),
      ...(nativeBlocker ? { nativeBlocker } : {}),
    });
  }
  return result;
}

/**
 * Whether an AutoFilter range would cover part of a Table (header and totals
 * rows included). Excel never lets a worksheet AutoFilter overlap a Table.
 */
export function sheetViewRefOverlapsTable(
  ref: string,
  tables: ReadonlyArray<{
    range: { r1: number; c1: number; r2: number; c2: number };
    totalsRow?: boolean;
  }>,
): boolean {
  const range = parseA1Range(ref.replace(/\$/g, ""));
  if (!range) return false;
  return tables.some((table) => {
    const lastRow = table.totalsRow ? table.range.r2 + 1 : table.range.r2;
    return (
      table.range.r1 <= range.r2 &&
      lastRow >= range.r1 &&
      table.range.c1 <= range.c2 &&
      table.range.c2 >= range.c1
    );
  });
}

/** A view's filter range for a block of cells, its first row the header. */
export function sheetViewRefForRange(range: { r1: number; c1: number; r2: number; c2: number }): string {
  return rangeA1(range);
}

export function defaultNamedSheetViewRef(source: {
  rowCount: number;
  columnCount: number;
}): string {
  const lastColumn = columnHeaderLabel(Math.max(0, source.columnCount - 1), "a1");
  return `A1:${lastColumn}${Math.max(1, source.rowCount)}`;
}

export function snapshotSheetViewFilters(
  source: SpreadsheetSheetViewFilterState,
): SpreadsheetSheetViewFilterState {
  let filters = source.filters
    ? Object.fromEntries(
        Object.entries(source.filters).map(([key, values]) => [key, [...values]]),
      )
    : undefined;
  // A column still showing exactly what its date groups resolved to keeps
  // them as date groups; a selection the user changed becomes plain values.
  let dateGroups: Record<string, FilterDateGroup[]> | undefined;
  for (const [key, groups] of Object.entries(source.dateGroups ?? {})) {
    const resolved = source.dateGroupValues?.[key];
    const current = filters?.[key];
    if (resolved === undefined) {
      (dateGroups ??= {})[key] = groups.map((group) => ({ ...group }));
      continue;
    }
    const resolvedSet = new Set(resolved);
    if (!current || !resolved.every((value) => current.includes(value))) continue;
    const literal = current.filter((value) => !resolvedSet.has(value));
    (dateGroups ??= {})[key] = groups.map((group) => ({ ...group }));
    if (filters) {
      if (literal.length > 0) filters[key] = literal;
      else delete filters[key];
    }
  }
  if (filters && Object.keys(filters).length === 0) filters = undefined;
  const conditionFilters = source.conditionFilters
    ? Object.fromEntries(
        Object.entries(source.conditionFilters).map(([key, condition]) => [
          key,
          { ...condition },
        ]),
      )
    : undefined;
  const colorFilters = source.colorFilters
    ? Object.fromEntries(
        Object.entries(source.colorFilters).map(([key, values]) => [key, [...values]]),
      )
    : undefined;
  const sortState = source.sortState
    ? cloneSheetViewSortState(source.sortState)
    : undefined;
  return {
    ...(filters && Object.keys(filters).length > 0 ? { filters } : {}),
    ...(dateGroups ? { dateGroups } : {}),
    ...(conditionFilters && Object.keys(conditionFilters).length > 0
      ? { conditionFilters }
      : {}),
    ...(colorFilters && Object.keys(colorFilters).length > 0
      ? { colorFilters }
      : {}),
    ...(sortState ? { sortState } : {}),
    ...(source.tables !== undefined
      ? { tables: source.tables.map(cloneSheetViewTableState) }
      : {}),
  };
}

function cloneSheetViewSortState(
  state: SpreadsheetSheetViewSortState,
): SpreadsheetSheetViewSortState {
  return {
    ...state,
    conditions: state.conditions.map((condition) => ({
      ...condition,
      ...(condition.icon ? { icon: { ...condition.icon } } : {}),
      ...(condition.customList ? { customList: [...condition.customList] } : {}),
    })),
  };
}

function cloneSheetViewTableState(
  table: SpreadsheetSheetViewTableState,
): SpreadsheetSheetViewTableState {
  return structuredClone(table);
}

/**
 * The view state the live sheet shows right now: its worksheet criteria (the
 * view's, while a view is active), the session's sort, and every Table's
 * criteria with the session's table sorts. `tables` is always present, so the
 * result is a whole-sheet state that also restores tables when applied.
 */
export function sheetViewStateOfSheet(
  sheet: SpreadsheetSheetViewFilterState & {
    dataTables?: ReadonlyArray<SheetViewTableShape & {
      columnFilters?: Record<string, DataTableColumnFilter>;
    }>;
  },
  session?: {
    sortState?: SpreadsheetSheetViewSortState;
    tables?: ReadonlyArray<SpreadsheetSheetViewSessionTable>;
  } | null,
): SpreadsheetSheetViewFilterState {
  const tables: SpreadsheetSheetViewTableState[] = [];
  for (const table of sheet.dataTables ?? []) {
    const entry = session?.tables?.find(
      (candidate) => candidate.tableId === table.id,
    );
    const state: SpreadsheetSheetViewTableState = {
      tableId: table.id,
      ...(entry?.filterId ? { filterId: entry.filterId } : {}),
      ...(table.columnFilters && Object.keys(table.columnFilters).length > 0
        ? { columnFilters: structuredClone(table.columnFilters) }
        : {}),
      ...(entry?.sortState ? { sortState: structuredClone(entry.sortState) } : {}),
    };
    if (sheetViewTableStateHasCriteria(state)) tables.push(state);
  }
  return snapshotSheetViewFilters({
    filters: sheet.filters,
    dateGroups: sheet.dateGroups,
    dateGroupValues: sheet.dateGroupValues,
    conditionFilters: sheet.conditionFilters,
    colorFilters: sheet.colorFilters,
    ...(session?.sortState ? { sortState: session.sortState } : {}),
    tables,
  });
}

/** A view's state as a whole-sheet state: no table entry means unfiltered. */
export function wholeSheetViewState(
  state: SpreadsheetSheetViewFilterState,
): SpreadsheetSheetViewFilterState {
  return { ...state, tables: state.tables ?? [] };
}

/**
 * The session entries for a view's tables: identity and the view's own table
 * sorts (a Table sort inside a view is a display order, never a cell move).
 */
export function sheetViewSessionTables(
  state: SpreadsheetSheetViewFilterState,
): SpreadsheetSheetViewSessionTable[] {
  return (state.tables ?? []).flatMap((table) =>
    table.filterId
      ? [
          {
            tableId: table.tableId,
            filterId: table.filterId,
            ...(table.sortState
              ? { sortState: structuredClone(table.sortState) }
              : {}),
          },
        ]
      : [],
  );
}

export function applySheetViewFilters<
  T extends SpreadsheetSheetViewFilterState & { dataTables?: unknown },
>(
  source: T,
  state: SpreadsheetSheetViewFilterState,
  /**
   * The display values of a column that fall in the given date groups. With
   * it, a view's date criteria filter rows like the values they stand for.
   */
  resolveDateGroups?: (column: number, groups: readonly FilterDateGroup[]) => string[],
): T {
  const next = { ...source } as T;
  if (state.filters && Object.keys(state.filters).length > 0) {
    next.filters = Object.fromEntries(
      Object.entries(state.filters).map(([key, values]) => [key, [...values]]),
    );
  } else {
    delete next.filters;
  }
  delete next.dateGroups;
  delete next.dateGroupValues;
  if (state.dateGroups && Object.keys(state.dateGroups).length > 0) {
    next.dateGroups = Object.fromEntries(
      Object.entries(state.dateGroups).map(([key, groups]) => [
        key,
        groups.map((group) => ({ ...group })),
      ]),
    );
    if (resolveDateGroups) {
      const resolvedByColumn: Record<string, string[]> = {};
      const filters: Record<string, string[]> = { ...(next.filters ?? {}) };
      for (const [key, groups] of Object.entries(state.dateGroups)) {
        const resolved = resolveDateGroups(Number(key), groups);
        resolvedByColumn[key] = resolved;
        filters[key] = [...new Set([...(filters[key] ?? []), ...resolved])];
      }
      next.filters = filters;
      next.dateGroupValues = resolvedByColumn;
    }
  }
  if (state.conditionFilters && Object.keys(state.conditionFilters).length > 0) {
    next.conditionFilters = Object.fromEntries(
      Object.entries(state.conditionFilters).map(([key, condition]) => [
        key,
        { ...condition },
      ]),
    );
  } else {
    delete next.conditionFilters;
  }
  if (state.colorFilters && Object.keys(state.colorFilters).length > 0) {
    next.colorFilters = Object.fromEntries(
      Object.entries(state.colorFilters).map(([key, values]) => [key, [...values]]),
    );
  } else {
    delete next.colorFilters;
  }
  // A whole-sheet state also sets every Table's criteria: the table the state
  // names shows its view criteria, any other table shows unfiltered. The
  // tables' own sortState is the default view's (a physical order) and stays.
  const withTables = next as T & {
    dataTables?: Array<{ id: string; columnFilters?: Record<string, DataTableColumnFilter> }>;
  };
  if (state.tables !== undefined && Array.isArray(withTables.dataTables)) {
    withTables.dataTables = withTables.dataTables.map((table) => {
      const entry = state.tables?.find((candidate) => candidate.tableId === table.id);
      const filters = entry?.columnFilters;
      const current = table.columnFilters ?? {};
      const wanted = filters ?? {};
      if (JSON.stringify(current) === JSON.stringify(wanted)) return table;
      const nextTable = { ...table };
      if (filters && Object.keys(filters).length > 0) {
        nextTable.columnFilters = structuredClone(filters);
      } else {
        delete nextTable.columnFilters;
      }
      return nextTable;
    });
  }
  return next;
}

export function sheetViewFilterStatesEqual(
  left: SpreadsheetSheetViewFilterState,
  right: SpreadsheetSheetViewFilterState,
): boolean {
  const stable = (value: unknown): string =>
    JSON.stringify(value ?? null, (_key, entry: unknown) =>
      entry && typeof entry === "object" && !Array.isArray(entry)
        ? Object.fromEntries(
            Object.entries(entry as Record<string, unknown>).sort(([a], [b]) =>
              a < b ? -1 : a > b ? 1 : 0,
            ),
          )
        : entry,
    );
  const tables = (state: SpreadsheetSheetViewFilterState) =>
    [...(state.tables ?? [])]
      .filter(sheetViewTableStateHasCriteria)
      .map((table) => ({
        tableId: table.tableId,
        columnFilters: table.columnFilters ?? {},
        sortState: table.sortState ?? null,
      }))
      .sort((a, b) => (a.tableId < b.tableId ? -1 : a.tableId > b.tableId ? 1 : 0));
  return (
    stable(left.filters) === stable(right.filters) &&
    stable(left.dateGroups) === stable(right.dateGroups) &&
    stable(left.conditionFilters) === stable(right.conditionFilters) &&
    stable(left.colorFilters) === stable(right.colorFilters) &&
    stable(left.sortState) === stable(right.sortState) &&
    stable(tables(left)) === stable(tables(right))
  );
}

/**
 * Why a sheet cannot open or hold a Sheet View. Advanced Filter hides rows
 * destructively, outside any AutoFilter a view addresses; an unreadable
 * native graph is preserve-only. Worksheet and Table AutoFilter criteria and
 * sorts are all view state.
 */
export function sheetViewAuthoringBlocker(
  source: SpreadsheetSheetViewSource,
): "advanced" | "native" | null {
  if (source.namedSheetViewsNativeBlocker) return "native";
  if (source.advancedFilterState) return "advanced";
  return null;
}

/**
 * One block of rows a Sheet View sorts: the worksheet AutoFilter's data rows
 * or one Table's data rows, with the keys as sort levels on absolute columns.
 */
export interface SheetViewSortRegion {
  r1: number;
  r2: number;
  levels: SortLevel[];
  /** The Table this region sorts; absent for the worksheet AutoFilter. */
  tableId?: string;
}

/** The sort levels of a worksheet view sort state (absolute key columns). */
export function sheetViewSortLevels(
  sortState: SpreadsheetSheetViewSortState,
  textOrder?: Pick<SpreadsheetTextOrder, "locale">,
): SortLevel[] | null {
  const method: ChineseCollation | undefined =
    sortState.sortMethod === "pinYin"
      ? "pinyin"
      : sortState.sortMethod === "stroke"
        ? "stroke"
        : undefined;
  const levels: SortLevel[] = [];
  for (const condition of sortState.conditions) {
    const range = parseA1Range(condition.ref.replace(/\$/g, ""));
    if (!range) return null;
    const placement = condition.descending === true ? "last" : "first";
    if (condition.sortBy === "cellColor" || condition.sortBy === "fontColor") {
      if (!condition.color) return null;
      levels.push({
        column: range.c1,
        direction: "asc",
        sortOn: condition.sortBy,
        formatKey: condition.color,
        placement,
      });
      continue;
    }
    if (condition.sortBy === "icon") {
      if (!condition.icon) return null;
      levels.push({
        column: range.c1,
        direction: "asc",
        sortOn: "icon",
        formatKey: sheetViewIconKey(condition.icon),
        placement,
      });
      continue;
    }
    levels.push({
      column: range.c1,
      direction: condition.descending === true ? "desc" : "asc",
      caseSensitive: sortState.caseSensitive === true,
      ...(textOrder?.locale ? { locale: textOrder.locale } : {}),
      ...(method ? { method } : {}),
      ...(condition.customList ? { orderKeys: condition.customList } : {}),
    });
  }
  return levels;
}

/** The display key a conditional-format icon sorts under (multi-sort). */
export function sheetViewIconKey(icon: DataTableIconFilter): string {
  return `${icon.iconSet}:${icon.iconId}:${iconSetCount(icon.iconSet)}`;
}

/**
 * Every block a view sorts: the worksheet AutoFilter, then each Table sorted
 * in the view, in the order the Tables sit on the sheet. A Table without its
 * header row or column identities cannot be sorted and is skipped.
 */
export function sheetViewSortRegions(
  sortState: SpreadsheetSheetViewSortState | undefined,
  tables: ReadonlyArray<SpreadsheetSheetViewSessionTable> | undefined,
  dataTables: ReadonlyArray<SheetViewTableShape> | undefined,
  textOrder?: Pick<SpreadsheetTextOrder, "locale">,
): SheetViewSortRegion[] {
  const regions: SheetViewSortRegion[] = [];
  if (sortState) {
    const bounds = parseA1Range(sortState.ref.replace(/\$/g, ""));
    const levels = sheetViewSortLevels(sortState, textOrder);
    if (bounds && levels) regions.push({ r1: bounds.r1, r2: bounds.r2, levels });
  }
  for (const table of dataTables ?? []) {
    const entry = tables?.find((candidate) => candidate.tableId === table.id);
    if (!entry?.sortState || !table.headerRow || !table.columns || !table.range) {
      continue;
    }
    const levels: SortLevel[] = [];
    for (const condition of entry.sortState.conditions) {
      const offset = table.columns.findIndex(
        (column) => column.id === condition.columnId,
      );
      if (offset < 0) continue;
      const column = table.range.c1 + offset;
      const placement = condition.descending === true ? "last" : "first";
      if (condition.icon) {
        levels.push({
          column,
          direction: "asc",
          sortOn: "icon",
          formatKey: sheetViewIconKey(condition.icon),
          placement,
        });
      } else if (condition.color) {
        levels.push({
          column,
          direction: "asc",
          sortOn: condition.color.kind === "cell" ? "cellColor" : "fontColor",
          formatKey: condition.color.value,
          placement,
        });
      } else {
        levels.push({
          column,
          direction: condition.descending === true ? "desc" : "asc",
          caseSensitive: false,
          ...(textOrder?.locale ? { locale: textOrder.locale } : {}),
        });
      }
    }
    if (levels.length > 0 && table.range.r2 > table.range.r1) {
      regions.push({
        r1: table.range.r1 + 1,
        r2: table.range.r2,
        levels,
        tableId: table.id,
      });
    }
  }
  return regions;
}

/**
 * The first region whose rows overlap another's, or null. A view orders whole
 * display rows, so two side-by-side blocks sharing rows cannot both be sorted.
 */
export function sheetViewSortRegionConflict(
  regions: ReadonlyArray<SheetViewSortRegion>,
): SheetViewSortRegion | null {
  for (let index = 1; index < regions.length; index += 1) {
    const region = regions[index]!;
    if (
      regions
        .slice(0, index)
        .some((other) => region.r1 <= other.r2 && other.r1 <= region.r2)
    ) {
      return region;
    }
  }
  return null;
}

/**
 * Build a display-slot -> physical-row permutation for a Sheet View's sorts.
 * Cell coordinates and formulas never move. Rows outside every sorted block
 * keep their original slot, and equal keys remain stable.
 *
 * `keyAt` answers with a whole sort key, not display text: the authenticated
 * numeric scalar and the logical/error type ride along so a view sorts in
 * Excel's type order (numbers, text, FALSE, TRUE, errors, blanks last) instead
 * of ordering "$1,200", TRUE and #N/A as the strings they are drawn as; the
 * displayed fill, font colour and icon ride along for format keys.
 */
export function sheetViewRegionsRowOrder(
  rowCount: number,
  regions: ReadonlyArray<SheetViewSortRegion>,
  /**
   * `display` is true for a colour or icon key: only then does the caller
   * need to resolve what the cell displays (conditional formats included).
   */
  keyAt: (row: number, column: number, display: boolean) => SortDisplayKeyCell,
): number[] | null {
  // The normal workbook path is identity. Keep it allocation-free so a
  // million-row virtual sheet does not materialize a million-entry array on
  // every cell edit merely because Sheet Views support is mounted.
  if (regions.length === 0) return null;
  const order = Array.from({ length: Math.max(0, rowCount) }, (_, row) => row);
  const applied: SheetViewSortRegion[] = [];
  for (const region of regions) {
    if (
      region.r1 < 0 ||
      region.r2 >= rowCount ||
      region.r2 < region.r1 ||
      applied.some((other) => region.r1 <= other.r2 && other.r1 <= region.r2)
    ) {
      continue;
    }
    applied.push(region);
    const rows = order.slice(region.r1, region.r2 + 1);
    // Each key is read once per row, not once per comparison.
    const keys = new Map(
      rows.map((row) => [
        row,
        region.levels.map((level) =>
          keyAt(row, level.column, (level.sortOn ?? "values") !== "values"),
        ),
      ]),
    );
    rows.sort((left, right) => {
      const leftKeys = keys.get(left)!;
      const rightKeys = keys.get(right)!;
      for (let index = 0; index < region.levels.length; index += 1) {
        const comparison = compareCellsAtLevel(
          leftKeys[index],
          rightKeys[index],
          region.levels[index]!,
        );
        if (comparison !== 0) return comparison;
      }
      return left - right;
    });
    order.splice(region.r1, rows.length, ...rows);
  }
  return order;
}

/** The worksheet-only form of {@link sheetViewRegionsRowOrder}. */
export function sheetViewRowOrder(
  rowCount: number,
  sortState: SpreadsheetSheetViewSortState | undefined,
  keyAt: (row: number, column: number) => SortDisplayKeyCell | SortFilterCell,
  textOrder?: Pick<SpreadsheetTextOrder, "locale">,
): number[] | null {
  if (!sortState) return null;
  const identity = () =>
    Array.from({ length: Math.max(0, rowCount) }, (_, row) => row);
  const regions = sheetViewSortRegions(sortState, undefined, undefined, textOrder);
  return sheetViewRegionsRowOrder(rowCount, regions, keyAt) ?? identity();
}

/**
 * Resolve the keyboard navigation sequence for an active Sheet View. The
 * identity/no-filter path stays null and allocation-free; sorted or filtered
 * views expose only visible source rows in visual order.
 */
export function sheetViewVisibleRowOrder(
  rowCount: number,
  displayOrder: readonly number[] | null,
  hiddenRows: ReadonlySet<number>,
): readonly number[] | null {
  if (!displayOrder && hiddenRows.size === 0) return null;
  if (displayOrder && hiddenRows.size === 0) return displayOrder;
  const visible: number[] = [];
  for (let visualRow = 0; visualRow < Math.max(0, rowCount); visualRow += 1) {
    const sourceRow = displayOrder?.[visualRow] ?? visualRow;
    if (!hiddenRows.has(sourceRow)) visible.push(sourceRow);
  }
  return visible;
}

export function createSheetViewValueSortState(
  viewRef: string,
  column: number,
  direction: "asc" | "desc",
): SpreadsheetSheetViewSortState | null {
  return createSheetViewMultiValueSortState(viewRef, [
    { column, direction },
  ]);
}

export function createSheetViewMultiValueSortState(
  viewRef: string,
  conditions: ReadonlyArray<{
    column: number;
    direction: "asc" | "desc";
  }>,
): SpreadsheetSheetViewSortState | null {
  return createSheetViewSortState(
    viewRef,
    conditions.map((condition) => ({
      column: condition.column,
      ...(condition.direction === "desc" ? { descending: true } : {}),
    })),
  );
}

/** One key of a view sort, on an absolute worksheet column. */
export interface SheetViewSortKeyInput {
  column: number;
  /** Value keys: Z to A. Format keys: the colour/icon "On Bottom". */
  descending?: boolean;
  sortBy?: SpreadsheetSheetViewSortBy;
  color?: string;
  icon?: DataTableIconFilter;
  customList?: readonly string[];
}

/** Why a view sort state could not be built from the dialog's keys. */
export type SheetViewSortStateProblem =
  /** A key column lies outside the view's filter range. */
  | "outsideView"
  /** A key repeats an earlier level (same column and Sort On identity). */
  | "repeatedLevel"
  /** No data rows, no keys or too many, or a malformed key. */
  | "invalid";

/**
 * The view sort state for keys over the whole AutoFilter range (header row
 * first), or null when a key is outside it, repeated or malformed.
 */
export function createSheetViewSortState(
  viewRef: string,
  keys: ReadonlyArray<SheetViewSortKeyInput>,
  options: { caseSensitive?: boolean; sortMethod?: "pinYin" | "stroke" } = {},
): SpreadsheetSheetViewSortState | null {
  const built = buildSheetViewSortState(viewRef, keys, options);
  return "state" in built ? built.state : null;
}

/**
 * {@link createSheetViewSortState} with the reason a state could not be
 * built, so the caller names the rule the keys broke. One column may appear
 * in several levels when they differ in Sort On or colour/icon.
 */
export function buildSheetViewSortState(
  viewRef: string,
  keys: ReadonlyArray<SheetViewSortKeyInput>,
  options: { caseSensitive?: boolean; sortMethod?: "pinYin" | "stroke" } = {},
): { state: SpreadsheetSheetViewSortState } | { problem: SheetViewSortStateProblem } {
  const view = parseA1Range(viewRef.replace(/\$/g, ""));
  if (!view || view.r2 <= view.r1 || keys.length < 1 || keys.length > 64) {
    return { problem: "invalid" };
  }
  const identities = new Set<string>();
  const data = { ...view, r1: view.r1 + 1 };
  const conditions: SpreadsheetSheetViewSortCondition[] = [];
  for (const key of keys) {
    if (!Number.isInteger(key.column)) return { problem: "invalid" };
    if (key.column < view.c1 || key.column > view.c2) {
      return { problem: "outsideView" };
    }
    const sortKey = normalizeSheetViewSortKey({
      ...(key.sortBy ? { sortBy: key.sortBy } : {}),
      ...(key.color !== undefined ? { color: key.color } : {}),
      ...(key.icon !== undefined ? { icon: key.icon } : {}),
      ...(key.customList !== undefined ? { customList: [...key.customList] } : {}),
    });
    if (!sortKey) return { problem: "invalid" };
    const identity = sheetViewSortKeyIdentity(key.column, sortKey);
    if (identities.has(identity)) return { problem: "repeatedLevel" };
    identities.add(identity);
    conditions.push({
      ref: rangeA1({ ...data, c1: key.column, c2: key.column }),
      ...(key.descending === true ? { descending: true } : {}),
      ...sortKey,
    });
  }
  return {
    state: {
      ref: rangeA1(data),
      conditions,
      ...(options.caseSensitive === true ? { caseSensitive: true as const } : {}),
      ...(options.sortMethod ? { sortMethod: options.sortMethod } : {}),
    },
  };
}

/** A Table AutoFilter inside the active view: its identity and view sort. */
export interface SpreadsheetSheetViewSessionTable {
  tableId: string;
  filterId: string;
  sortState?: DataTableSortState;
}

export interface SpreadsheetSheetViewSessionState {
  sheetId: string;
  activeViewId: string | "temporary";
  defaultState: SpreadsheetSheetViewFilterState;
  defaultAutoFilterId?: string;
  defaultAutoFilterRef?: string;
  defaultOwnsAutoFilter?: true;
  filterId?: string;
  ref?: string;
  sortState?: SpreadsheetSheetViewSortState;
  ownsAutoFilter?: true;
  /**
   * The Table AutoFilters this view addresses: their nsvFilter identity and
   * the view's own Table sorts (a display order; the Tables' cells and their
   * default sortState never move). Table criteria live on the live sheet's
   * Tables while the view is open, like the worksheet criteria do.
   */
  tables?: SpreadsheetSheetViewSessionTable[];
  /**
   * Named views only: the definition as this client last saw it in the
   * shared workbook. A merge whose definition differs from it is another
   * person's change to the view and is adopted; a merge that still matches
   * it is an echo of the shared copy while this client's own edit is unsaved,
   * and the local projection stays.
   */
  sharedState?: SpreadsheetSheetViewFilterState;
}

type SheetViewSessionSheet = SpreadsheetSheetViewFilterState & {
  id: string;
  autoFilterId?: string;
  autoFilterRef?: string;
  namedSheetViewsOwnsAutoFilter?: true;
  namedSheetViews?: SpreadsheetNamedSheetView[];
  dataTables?: ReadonlyArray<
    SheetViewTableShape & { columnFilters?: Record<string, DataTableColumnFilter> }
  >;
};

/**
 * The nsvFilter identity a view uses for one Table: the one this session or
 * any view on the sheet already gave it (every view of a table shares its
 * AutoFilter's uid), else a new one from `create`.
 */
export function sheetViewTableFilterId(
  sheet: { namedSheetViews?: ReadonlyArray<SpreadsheetNamedSheetView> },
  session: { tables?: ReadonlyArray<SpreadsheetSheetViewSessionTable> } | null | undefined,
  tableId: string,
  create: () => string,
): string {
  return (
    session?.tables?.find((table) => table.tableId === tableId)?.filterId ??
    sheet.namedSheetViews
      ?.flatMap((view) => view.tables ?? [])
      .find((table) => table.tableId === tableId && table.filterId)?.filterId ??
    create()
  );
}

/** Give every table entry of a view state its nsvFilter identity. */
export function withSheetViewTableFilterIds(
  sheet: { namedSheetViews?: ReadonlyArray<SpreadsheetNamedSheetView> },
  session: { tables?: ReadonlyArray<SpreadsheetSheetViewSessionTable> } | null | undefined,
  state: SpreadsheetSheetViewFilterState,
): SpreadsheetSheetViewFilterState {
  if (!state.tables || state.tables.every((table) => table.filterId)) return state;
  return {
    ...state,
    tables: state.tables.map((table) =>
      table.filterId
        ? table
        : {
            ...table,
            filterId: sheetViewTableFilterId(
              sheet,
              session,
              table.tableId,
              createNamedSheetViewId,
            ),
          },
    ),
  };
}

/**
 * The session with an entry for every Table the live sheet filters or the
 * view sorts, each with its nsvFilter identity, and none for a Table that no
 * longer exists. Returns the same object when nothing changes.
 */
export function ensureSheetViewSessionTables<
  S extends SpreadsheetSheetViewSessionState,
>(
  session: S,
  sheet: SheetViewSessionSheet,
): S {
  const tableIds = new Set((sheet.dataTables ?? []).map((table) => table.id));
  const entries = (session.tables ?? []).filter((entry) =>
    tableIds.has(entry.tableId),
  );
  let changed = entries.length !== (session.tables ?? []).length;
  for (const table of sheet.dataTables ?? []) {
    if (entries.some((entry) => entry.tableId === table.id)) continue;
    if (Object.keys(table.columnFilters ?? {}).length === 0) continue;
    entries.push({
      tableId: table.id,
      filterId: sheetViewTableFilterId(sheet, session, table.id, createNamedSheetViewId),
    });
    changed = true;
  }
  if (!changed) return session;
  const next = { ...session };
  if (entries.length > 0) next.tables = entries;
  else delete next.tables;
  return next;
}

/**
 * Re-project this client's Sheet View onto sheets that just arrived from a
 * collaborator's merge. Microsoft: "When using Sheet View, you will be
 * unaffected by sorts and filters from others" — the peer's cells and their
 * default filter land, the view's own filters and sort stay on screen.
 *
 * Returns null when the view can no longer be held: the sheet is gone, the
 * named view was deleted, or the sheet now carries filters a view cannot
 * express (table, advanced, condition, colour). The caller then exits the
 * view the same way it would for a fresh load.
 */
export function reprojectSheetViewSession<
  T extends SheetViewSessionSheet & Partial<SpreadsheetSheetViewSource>,
>(
  loadedSheets: readonly T[],
  session: SpreadsheetSheetViewSessionState,
  liveSheet: T | undefined,
): { sheets: T[]; session: SpreadsheetSheetViewSessionState } | null {
  const loaded = loadedSheets.find((sheet) => sheet.id === session.sheetId);
  if (!loaded || !liveSheet || liveSheet.id !== session.sheetId) return null;
  if (sheetViewAuthoringBlocker(loaded as SpreadsheetSheetViewSource)) return null;
  let viewState = sheetViewStateOfSheet(liveSheet, session);
  let sessionTables = session.tables;
  let sharedState = session.sharedState;
  if (session.activeViewId !== "temporary") {
    const view = (loaded.namedSheetViews ?? []).find(
      (candidate) => candidate.id === session.activeViewId,
    );
    if (!view || view.nativeBlocker) return null;
    const merged = wholeSheetViewState(snapshotSheetViewFilters(view));
    // A named view is one shared definition. Someone else's change to it
    // shows here at once; the shared copy echoing back while this client's
    // own edit is still unsaved does not undo that edit.
    if (!sharedState || !sheetViewFilterStatesEqual(merged, sharedState)) {
      viewState = merged;
      sessionTables = sheetViewSessionTables(merged);
    }
    sharedState = merged;
  }
  const projected = applySheetViewFilters({ ...loaded }, viewState);
  delete projected.autoFilterId;
  delete projected.autoFilterRef;
  delete projected.namedSheetViewsOwnsAutoFilter;
  if (liveSheet.autoFilterId) projected.autoFilterId = liveSheet.autoFilterId;
  if (liveSheet.autoFilterRef) projected.autoFilterRef = liveSheet.autoFilterRef;
  if (liveSheet.namedSheetViewsOwnsAutoFilter) {
    projected.namedSheetViewsOwnsAutoFilter = true;
  }
  // The persisted sheet carries the shared default state: the peer's filter
  // when they changed it, this client's own snapshot when the merge merely
  // echoes an earlier save. Either way it is what leaving the view restores.
  const nextSession: SpreadsheetSheetViewSessionState = {
    ...session,
    defaultState: sheetViewStateOfSheet(loaded, null),
    ...(viewState.sortState
      ? { sortState: viewState.sortState }
      : {}),
    ...(sessionTables && sessionTables.length > 0 ? { tables: sessionTables } : {}),
    ...(sharedState ? { sharedState } : {}),
  };
  if (!viewState.sortState) delete nextSession.sortState;
  if (!sessionTables || sessionTables.length === 0) delete nextSession.tables;
  const echoesOwnAssociation =
    Boolean(session.filterId) && loaded.autoFilterId === session.filterId;
  if (!echoesOwnAssociation) {
    delete nextSession.defaultAutoFilterId;
    delete nextSession.defaultAutoFilterRef;
    delete nextSession.defaultOwnsAutoFilter;
    if (loaded.autoFilterId) nextSession.defaultAutoFilterId = loaded.autoFilterId;
    if (loaded.autoFilterRef) nextSession.defaultAutoFilterRef = loaded.autoFilterRef;
    if (loaded.namedSheetViewsOwnsAutoFilter) {
      nextSession.defaultOwnsAutoFilter = true;
    }
  }
  return {
    sheets: loadedSheets.map((sheet) => (sheet.id === loaded.id ? projected : sheet)),
    session: nextSession,
  };
}

export function persistedSheetsForSheetViewSession<T extends SheetViewSessionSheet>(
  sheets: readonly T[],
  session: SpreadsheetSheetViewSessionState | null,
): T[] {
  if (!session) return [...sheets];
  return sheets.map((sheet) => {
    if (sheet.id !== session.sheetId) return sheet;
    let persistedSheet = sheet;
    let namedSheetViews = sheet.namedSheetViews;
    if (session.activeViewId === "temporary") {
      persistedSheet = { ...sheet };
      delete persistedSheet.autoFilterId;
      delete persistedSheet.autoFilterRef;
      delete persistedSheet.namedSheetViewsOwnsAutoFilter;
      const defaultOwnedAssociationStillExists = Boolean(
        session.defaultAutoFilterId &&
          (namedSheetViews ?? []).some(
            (view) => view.filterId === session.defaultAutoFilterId,
          ),
      );
      const mayRestoreDefaultAssociation =
        session.defaultOwnsAutoFilter !== true ||
        defaultOwnedAssociationStillExists;
      if (session.defaultAutoFilterId && mayRestoreDefaultAssociation) {
        persistedSheet.autoFilterId = session.defaultAutoFilterId;
      }
      if (session.defaultAutoFilterRef && mayRestoreDefaultAssociation) {
        persistedSheet.autoFilterRef = session.defaultAutoFilterRef;
      }
      if (session.defaultOwnsAutoFilter && defaultOwnedAssociationStillExists) {
        persistedSheet.namedSheetViewsOwnsAutoFilter = true;
      }
    } else {
      const state = withSheetViewTableFilterIds(
        sheet,
        session,
        sheetViewStateOfSheet(sheet, session),
      );
      const hasFilterState = Boolean(
        state.sortState || sheetViewStateHasFilters(state),
      );
      namedSheetViews = (namedSheetViews ?? []).map((view) =>
        view.id === session.activeViewId
          ? rebuildNamedSheetViewState(
              view,
              state,
              session.filterId,
              session.ref,
            )
          : view,
      );
      if (hasFilterState && session.filterId) {
        persistedSheet = {
          ...sheet,
          autoFilterId: session.filterId,
          ...(session.ref ? { autoFilterRef: session.ref } : {}),
          ...(session.ownsAutoFilter
            ? { namedSheetViewsOwnsAutoFilter: true as const }
            : {}),
        };
      }
    }
    if (
      persistedSheet.namedSheetViewsOwnsAutoFilter === true &&
      !(
        persistedSheet.autoFilterId &&
        (namedSheetViews ?? []).some(
          (view) => view.filterId === persistedSheet.autoFilterId,
        )
      )
    ) {
      persistedSheet = { ...persistedSheet };
      delete persistedSheet.autoFilterId;
      delete persistedSheet.autoFilterRef;
      delete persistedSheet.namedSheetViewsOwnsAutoFilter;
      if (Object.keys(persistedSheet.filters ?? {}).length === 0) {
        delete persistedSheet.filters;
      }
    }
    const restored = applySheetViewFilters(
      persistedSheet,
      session.defaultState,
    );
    return namedSheetViews && namedSheetViews.length > 0
      ? { ...restored, namedSheetViews }
      : restored;
  });
}

// ---------------------------------------------------------------------------
// Sheet View > Options (Rename, Duplicate, Delete) as pure transactions, shared
// by the window's Sheet View dialog and the spreadsheet editor commands.

/** A view name as typed, or why it cannot be one ("invalid" text, or a name
 *  another view of the sheet has in any case). */
export function sheetViewNameFor(
  views: readonly SpreadsheetNamedSheetView[],
  raw: unknown,
  renamingId?: string,
): { name: string } | { error: "invalid" | "duplicate" } {
  const name = normalizeNamedSheetViewName(raw);
  if (!name) return { error: "invalid" };
  const collision = views.some(
    (view) =>
      view.name.toLocaleLowerCase() === name.toLocaleLowerCase() &&
      (renamingId === undefined || view.id !== renamingId),
  );
  return collision ? { error: "duplicate" } : { name };
}

/** Rename one view; null when it is unknown or preserve-only. */
export function sheetViewsRenamed(
  views: readonly SpreadsheetNamedSheetView[],
  id: string,
  name: string,
): SpreadsheetNamedSheetView[] | null {
  const selected = views.find((view) => view.id === id);
  if (!selected || selected.nativeBlocker) return null;
  return views.map((view) => (view.id === id ? { ...view, name } : view));
}

/** Duplicate one view as "<name> (2)" (the first free number); null when it
 *  is unknown, preserve-only, or the sheet holds its last view. */
export function sheetViewsWithDuplicate(
  views: readonly SpreadsheetNamedSheetView[],
  id: string,
): SpreadsheetNamedSheetView[] | null {
  const source = views.find((view) => view.id === id);
  if (!source || source.nativeBlocker || views.length >= MAX_NAMED_SHEET_VIEWS) return null;
  const names = new Set(views.map((view) => view.name.toLocaleLowerCase()));
  let index = 2;
  const duplicateName = (suffix: number) => {
    const tail = ` (${suffix})`;
    return `${source.name.slice(0, Math.max(1, 127 - tail.length))}${tail}`;
  };
  let name = duplicateName(index);
  while (names.has(name.toLocaleLowerCase())) {
    index += 1;
    name = duplicateName(index);
  }
  return [...views, { ...source, id: createNamedSheetViewId(), name }];
}

/**
 * Delete one view. An open view first gives the sheet back its default
 * filters (`restoreDefault`). The last view the editor's own AutoFilter
 * existed for takes that AutoFilter identity with it. A preserve-only view
 * can be deleted; a sheet whose whole Sheet View part is preserve-only
 * cannot. Null when the view is unknown or the sheet's views are read-only.
 */
export function sheetWithoutSheetView<
  S extends {
    namedSheetViews?: SpreadsheetNamedSheetView[];
    namedSheetViewsOwnsAutoFilter?: true;
    autoFilterId?: string;
    autoFilterRef?: string;
    filters?: Record<string, string[]>;
    namedSheetViewsNativeBlocker?: unknown;
  },
>(
  sheet: S,
  id: string,
  restoreDefault?: (sheet: S) => S,
): { sheet: S; removesFinalOwnedView: boolean } | null {
  const views = sheet.namedSheetViews ?? [];
  const selected = views.find((view) => view.id === id);
  if (!selected || sheet.namedSheetViewsNativeBlocker) return null;
  const nextViews = views.filter((view) => view.id !== selected.id);
  const nextSheetBase: S = { ...sheet, namedSheetViews: nextViews.length > 0 ? nextViews : undefined };
  let nextSheet = restoreDefault ? restoreDefault(nextSheetBase) : nextSheetBase;
  const removesFinalOwnedView = nextViews.length === 0 && nextSheet.namedSheetViewsOwnsAutoFilter === true;
  if (removesFinalOwnedView) {
    nextSheet = { ...nextSheet };
    delete nextSheet.autoFilterId;
    delete nextSheet.autoFilterRef;
    delete nextSheet.namedSheetViewsOwnsAutoFilter;
    if (Object.keys(nextSheet.filters ?? {}).length === 0) delete nextSheet.filters;
  }
  return { sheet: nextSheet, removesFinalOwnedView };
}
