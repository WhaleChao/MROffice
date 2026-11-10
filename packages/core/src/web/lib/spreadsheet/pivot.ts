/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for spreadsheet PIVOT TABLES (Excel/WPS parity).
 *
 * A pivot is computed from a SOURCE MATRIX (a rectangular block of display
 * strings read out of the sheet, FIRST row = header labels, the rest = data
 * rows) plus a CONFIG that says which source columns become ROW fields, which
 * become COLUMN fields, and which become aggregated VALUE fields. The result is
 * a grouped grid: one logical row per distinct combination of the row-field
 * values, one logical column per (distinct column-field combination x value
 * field), each cell holding the aggregate of the matching source rows, with
 * grand totals on the trailing row / column.
 *
 * Everything here is dependency-free and deterministic so it can be unit-tested
 * without React: the component layer only reads a source range out of the sheet
 * (via cellDisplayValue), hands the matrix + config here, and renders the
 * returned grid read-only. Numeric coercion is forgiving (currency symbols,
 * thousands separators, leading/trailing space) and blanks are ignored, exactly
 * as a user pasting "$1,200" expects SUM to behave.
 *
 * The persisted config is ADDITIVE/optional on the sheet schema (see
 * spreadsheetSheetSchema in lib/native-page-schemas — `pivot`), so an existing
 * sheet without one still validates and the pivot recomputes from current data
 * on every reload (the GRID is never persisted, only the recipe).
 */

import {
  EXCEL_1900_MARCH_FIRST_SERIAL,
  excelSerialFromUtcMs,
  excelSerialToUtcMs,
} from "./excel-1900-calendar";
import { columnToLetters, lettersToColumn } from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { applyNumberFormat } from "./number-format";
import {
  applyPivotFieldFilters,
  normalizePivotFieldFilters,
  pivotFieldFilterNeedsValue,
  remapPivotFieldFilterLevels,
  type PivotFieldFilter,
} from "./pivot-field-filters";

export type { PivotFieldFilter } from "./pivot-field-filters";

// Grid bounds mirror the schema / SpreadsheetWindow clamps so a normalized range
// can never address outside the grid.
// The last addressable row/column index of a sheet (Excel's 1,048,576 rows by
// 16,384 columns), as every other range reader clamps (conditional-format,
// data-table): a source past the old 10,000 x 1,000 grid lost its rows and
// fields past it.
const MAX_RANGE_ROW = MAX_GRID_ROWS - 1;
const MAX_RANGE_COL = MAX_GRID_COLUMNS - 1;

/**
 * The eleven functions of Excel's "Summarize value field by" (XLS-A071). Each
 * id is the Open XML ST_DataConsolidateFunction value a dataField's `subtotal`
 * attribute carries, so the model and the file spell them the same way.
 */
export type PivotAgg =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNums"
  | "stdDev"
  | "stdDevp"
  | "var"
  | "varp";

/** Excel's own order in Value Field Settings and "Summarize Values By". */
export const PIVOT_AGGS: readonly PivotAgg[] = [
  "sum",
  "count",
  "average",
  "max",
  "min",
  "product",
  "countNums",
  "stdDev",
  "stdDevp",
  "var",
  "varp",
] as const;

/**
 * The "Show Values As" post-aggregation display transform on a value field
 * (Excel "Show Values As"). Each transforms the value field's aggregated cells
 * over the result grid AFTER aggregation — the underlying aggregate is unchanged,
 * only how it is presented:
 *  - `normal` — the raw aggregate (default; omit/undefined behaves the same).
 *  - `percentOfGrandTotal` — each cell / the value field's grand total.
 *  - `percentOfRowTotal` — each cell / its output ROW total (within this value
 *    field's columns).
 *  - `percentOfColumnTotal` — each cell / its output COLUMN total (down all
 *    rows).
 *  - `runningTotalInRows` / `runningTotalInColumns` — cumulative sum down the
 *    rows / across the columns in axis order.
 *  - `rankSmallestToLargest` / `rankLargestToSmallest` — competition rank of the
 *    cell among the value field's body cells in its COLUMN (ties share a rank,
 *    next rank skips — Excel RANK.EQ semantics).
 *  - `differenceFrom` / `percentDifferenceFrom` — the cell minus (or as a % of)
 *    the value at the chosen BASE item along a base axis field (see
 *    `PivotShowValuesAsConfig.base`).
 */
export type PivotShowValuesAs =
  | "normal"
  | "percentOfGrandTotal"
  | "percentOfRowTotal"
  | "percentOfColumnTotal"
  | "percentOfBaseItem"
  | "percentOfParentRowTotal"
  | "percentOfParentColumnTotal"
  | "percentOfParentTotal"
  | "runningTotalInRows"
  | "runningTotalInColumns"
  | "percentOfRunningTotalIn"
  | "rankSmallestToLargest"
  | "rankLargestToSmallest"
  | "differenceFrom"
  | "percentDifferenceFrom"
  | "index";

/** Excel's own menu order, which the value settings dialog follows. */
export const PIVOT_SHOW_VALUES_AS: readonly PivotShowValuesAs[] = [
  "normal",
  "percentOfGrandTotal",
  "percentOfColumnTotal",
  "percentOfRowTotal",
  "percentOfBaseItem",
  "percentOfParentRowTotal",
  "percentOfParentColumnTotal",
  "percentOfParentTotal",
  "differenceFrom",
  "percentDifferenceFrom",
  "runningTotalInRows",
  "runningTotalInColumns",
  "percentOfRunningTotalIn",
  "rankSmallestToLargest",
  "rankLargestToSmallest",
  "index",
] as const;

/** True when a transform divides by a total / base and so renders as a percent. */
export function isPercentShowValuesAs(mode: PivotShowValuesAs): boolean {
  return (
    mode === "percentOfGrandTotal" ||
    mode === "percentOfRowTotal" ||
    mode === "percentOfColumnTotal" ||
    mode === "percentOfBaseItem" ||
    mode === "percentOfParentRowTotal" ||
    mode === "percentOfParentColumnTotal" ||
    mode === "percentOfParentTotal" ||
    mode === "percentOfRunningTotalIn" ||
    mode === "percentDifferenceFrom"
  );
}

/**
 * The BASE FIELD a transform runs along, and for the two "Difference From"
 * modes the item it measures against. `axis` is "row" or "col"; `pos` is the
 * field's position within that axis (config.rows[pos] / config.columns[pos]);
 * `item` is the base value's display string (matched case/space-insensitively,
 * like GETPIVOTDATA). A base item that matches no group leaves every cell null
 * (Excel shows blank for an unknown base). Excel names a base field for every
 * mode but the shares of a total, and writes it as the dataField's `baseField`.
 */
export interface PivotShowValuesAsBase {
  axis: "row" | "col";
  pos: number;
  /** Required (and only meaningful) for the differenceFrom modes. */
  item?: string;
}

/** The full "Show Values As" setting on a value field (mode + optional base). */
export interface PivotShowValuesAsConfig {
  mode: PivotShowValuesAs;
  /**
   * The field the transform runs along (and the item it measures against).
   * A running total or a rank without one runs along the innermost field of
   * its axis, which is what a report with one field on that axis means.
   */
  base?: PivotShowValuesAsBase;
}

/**
 * One value field: either a SOURCE column aggregated (the original shape — a
 * source `colIndex` + an `agg`), or a CALCULATED field referenced by name via
 * `calcRef` (then `colIndex`/`agg` are placeholders the engine ignores). An
 * optional `showAs` applies a "Show Values As" display transform after
 * aggregation. The persisted shape stays byte-identical for an ordinary
 * {colIndex, agg} field (calcRef/showAs are additive + optional), so existing
 * sheets round-trip unchanged.
 */
export interface PivotValueField {
  colIndex: number;
  agg: PivotAgg;
  /** Optional report caption; source field identity remains unchanged. */
  caption?: string;
  /** Excel number format applied to every displayed value in this field. */
  numberFormat?: string;
  /** When set, this value field is the calculated field of that name. */
  calcRef?: string;
  /** When set, a post-aggregation display transform over the result grid. */
  showAs?: PivotShowValuesAsConfig;
}

/**
 * A user-defined CALCULATED FIELD: a `name` + a `formula` that combines OTHER
 * source fields by header label. Excel computes a calculated field over the
 * SUMMED components of each group — i.e. for group g the field's value is the
 * formula evaluated with each referenced source field replaced by SUM(field)
 * over g (NOT the per-row formula averaged). This module replicates that exact
 * documented semantic via evaluatePivotCalcFormula. Names are referenced by a
 * value field's `calcRef`.
 */
export interface PivotFieldSort {
  colIndex: number;
  order: "ascending" | "descending";
  /**
   * Sort the field's items by this value field (its index in `values`)
   * instead of by their labels: Excel's "Sort Largest to Smallest" on a value,
   * the PivotField's `autoSortScope`. Each outer item's items are ordered by
   * their value over every column.
   */
  byValue?: number;
}

/**
 * The order Excel lists an unsorted field's items in (the PivotField's own
 * item list, which a user reorders by hand): carried with the recipe and
 * written back, so the report keeps it. Items the list does not name follow
 * in first-seen order.
 */
export interface PivotFieldItemOrder {
  colIndex: number;
  items: string[];
}

export interface PivotFieldNumberFormat {
  colIndex: number;
  numberFormat: string;
}

export interface PivotCalculatedField {
  name: string;
  formula: string;
}

/**
 * A REPORT / PAGE filter (Excel PivotTable "Filters" area): a SOURCE column whose
 * value restricts which source rows feed the whole pivot, BEFORE any grouping or
 * aggregation. `selected` is the set of allowed display values; an EMPTY
 * `selected` means "(All)" — no restriction — so a freshly-added filter (nothing
 * picked yet) shows everything, matching Excel. A row passes only when, for every
 * filter, its value in `colIndex` is one of that filter's `selected` values.
 */
export interface PivotFilterField {
  colIndex: number;
  /** Allowed source display values; empty ⇒ (All), no restriction. */
  selected: string[];
  /**
   * A date field grouped by several units is laid out as several axis fields
   * (Years, Quarters, then the field itself); an item filter on one of those
   * levels names the level's unit, and `selected` holds that level's group
   * items ("2024", "Qtr1", "Jan"). The derived levels have no source column,
   * so the filter stays on the date field's own `colIndex`. Absent: the
   * field's own items, which are its group items once the field is grouped.
   */
  unit?: PivotDateUnit;
  /**
   * A field grouped by hand (Excel's "Group Selection") gains a group field
   * ("Region2") beside it; an item filter on that group field is stored on
   * the base field's `colIndex` with `group: true`, and `selected` holds the
   * group field's items: a group's name, or an item outside every group as
   * itself. Never together with `unit`.
   */
  group?: true;
}

/**
 * A TIMELINE date filter (Excel PivotTable timeline, x15 `dateBetween` with
 * `useWholeDay`): only source rows whose `colIndex` value falls on a calendar
 * day in [start, end] (inclusive, `YYYY-MM-DD`) feed the pivot. A timeline
 * owns the filter on its source field; `timeline` is the native filter name
 * (Excel writes the field's), and `filterId` and `evalOrder` keep the native
 * filter identity. Driven only by a timeline (see lib/spreadsheet/pivot-timeline).
 */
export interface PivotDateFilter {
  colIndex: number;
  start: string;
  end: string;
  timeline: string;
  filterId?: number;
  evalOrder?: number;
}

/** Date-grouping granularity (Excel PivotTable "Group field" for dates). */
export const PIVOT_DATE_UNITS = ["year", "quarter", "month"] as const;
export type PivotDateUnit = (typeof PIVOT_DATE_UNITS)[number];













export interface PivotDiscreteGroup {
  name: string;
  items: string[];
}

export interface PivotGrouping {
  colIndex: number;
  kind: "number" | "date" | "discrete";
  /** number kind: bin width (> 0). */
  size?: number;
  /** number kind: bin origin (default 0). */
  start?: number;
  /** date kind: granularity. */
  unit?: PivotDateUnit;
  /** discrete kind (Excel's manual "Group Selection"): the named groups; an
   * item outside every group stays its own. */
  groups?: PivotDiscreteGroup[];
  /** discrete kind: how the axis shows it — the groups with the items nested
   * under them (Excel's default after grouping), the groups alone, or the
   * items alone with the group field parked. */
  display?: "groups+items" | "groups" | "items";
}

/**
 * Stable identity for one cache field in an imported native PivotTable.
 * `id` is derived from the original pivot/cache relationship graph and never
 * from the field's mutable display position. `sourceIndex` is the source-range
 * column the field owns; grouped cache fields point back to their base column.
 */
export interface PivotNativeFieldAuthority {
  id: string;
  name: string;
  sourceIndex: number;
  cacheIndex: number;
  groupedFromCacheIndex?: number;
}

export interface PivotNativePartAuthority {
  path: string;
  sha256: string;
}

export type PivotNativeScalarKind = "s" | "n" | "d" | "b" | "e" | "m";

/** Exact native cache scalar authority; `raw` is never display-normalized. */
export interface PivotNativeScalar {
  kind: PivotNativeScalarKind;
  raw: string;
}

export type PivotNativeScalarMatrix = PivotNativeScalar[][];

/**
 * The explicit native Pivot cache snapshot. Imported PivotTables read from this
 * immutable matrix until Refresh succeeds; ordinary worksheet edits therefore
 * mark the Pivot dirty instead of silently recomputing from live cells.
 */
export interface PivotNativeCacheSnapshot {
  revision: number;
  fingerprint: string;
  fieldIds: string[];
  matrix: PivotMatrix;
  scalarMatrix: PivotNativeScalarMatrix;
}

/**
 * A cache snapshot can also belong to a newly-authored PivotTable before the
 * first XLSX export has allocated native part/relationship identities.  The
 * snapshot is still exact typed authority; only the package identity is absent.
 */
export type PivotCacheSnapshot = PivotNativeCacheSnapshot;

export const PIVOT_REPORT_LAYOUTS = ["compact", "outline", "tabular"] as const;
export type PivotReportLayout = (typeof PIVOT_REPORT_LAYOUTS)[number];
export const PIVOT_SUBTOTAL_POSITIONS = ["none", "top", "bottom"] as const;
export type PivotSubtotalPosition = (typeof PIVOT_SUBTOTAL_POSITIONS)[number];
export const PIVOT_GRAND_TOTAL_MODES = ["both", "rows", "columns", "none"] as const;
export type PivotGrandTotalMode = (typeof PIVOT_GRAND_TOTAL_MODES)[number];
export const PIVOT_STYLE_NAMES = Object.freeze(
  (["Light", "Medium", "Dark"] as const).flatMap((tone) =>
    Array.from({ length: 28 }, (_, index) => `PivotStyle${tone}${index + 1}`),
  ),
) as readonly string[];

export function isPivotStyleName(value: unknown): value is string {
  return typeof value === "string" && PIVOT_STYLE_NAMES.includes(value);
}

export interface PivotNativeAuthority {
  version: 1;
  disposition: "editable" | "preserve-only";
  blocker?: string;
  pivotName: string;
  sourceSheetName: string;
  sourceRef: string;
  outputRef: string;
  styleName?: string;
  worksheetRelationshipId: string;
  cacheId: number;
  cacheRelationshipId: string;
  recordsRelationshipId?: string;
  pivotTable: PivotNativePartAuthority;
  cacheDefinition: PivotNativePartAuthority;
  cacheRecords?: PivotNativePartAuthority;
  fields: PivotNativeFieldAuthority[];
  snapshot: PivotNativeCacheSnapshot;
  /**
   * Set when the PivotTable reads a cache other PivotTables read too. In the
   * persisted workbook the snapshot then lives once under `pivotCaches`
   * (see `sharePivotCacheSnapshots`); in memory every sharing PivotTable holds
   * the same snapshot object, and a refresh of one refreshes them all.
   */
  cacheKey?: string;
  /**
   * Hidden items (a slicer's selection) the imported output cells were
   * computed with. Immutable receipt; the recipe's own `itemFilters` move with
   * the slicer.
   */
  itemFilters?: PivotFilterField[];
  /**
   * Timeline filters the imported output cells were computed with. Immutable
   * receipt; the recipe's own `dateFilters` move with the timeline.
   */
  dateFilters?: PivotDateFilter[];
  /**
   * Field filters the imported output cells were computed with. Immutable
   * receipt; the recipe's own `fieldFilters` move with the report.
   */
  fieldFilters?: PivotFieldFilter[];
  /**
   * The cells the PivotTable occupied at import (`pivotOutputFingerprint`). An
   * export recognises them as its own however Excel laid them out and however
   * the recipe has changed since.
   */
  outputFingerprint?: string;
  /**
   * 2 when the recipe's layout keys were read from each field's own flags, as
   * Excel applies them. Pages imported before carry the older reading, which
   * the writer takes for an unchanged layout.
   */
  layoutReading?: 2;
  /**
   * The cache refreshes when the file opens (`refreshOnLoad`): the report
   * shows its refreshed plan, never the cells the file cached (X-158).
   */
  refreshOnLoad?: true;
  /**
   * Classic PivotTable layout (the part's `gridDropZones`): the column area
   * keeps its own header row with no column field, so one value field's
   * caption heads the report and its column reads "Total" (X-007).
   */
  classicLayout?: true;
  /**
   * PivotTable Options > Show the Values row cleared (the part's x14
   * `hideValuesRow`): several value fields with no column field head their
   * columns in the one header row, without a "Values" row (X-266).
   */
  hideValuesRow?: true;
  /** The caption the report heads its value fields with (`dataCaption`),
   * shown as stated in any UI language (X-249). */
  dataCaption?: string;
  /**
   * Where the report's areas lie in the cells the file cached (its
   * `location` and `rowItems` / `colItems`): a report the grid does not draw
   * wears its PivotTable style over those cells from this note (X-235).
   * An import note; malformed, it is dropped, never repaired.
   */
  savedLayout?: PivotNativeSavedLayout;
  /**
   * The report's own formats on every item label of one field (`format` /
   * `pivotArea labelOnly`, a reference naming the field's items all): the
   * source field and the dxf's fill, font colour and weight (CX118, X-283).
   * An import note; malformed entries are dropped.
   */
  labelFormats?: PivotNativeLabelFormat[];
  /** The report's formats on its grand totals, field buttons and the values
   * of chosen items (X-294a). An import note. */
  areaFormats?: PivotNativeAreaFormat[];
  /** Fields whose one Custom subtotal shows at the top of each group
   * (`subtotalTop`, X-311). An import note. */
  subtotalTopFields?: number[];
}

export interface PivotNativeLabelFormat {
  field: number;
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

export function normalizePivotNativeLabelFormats(raw: unknown): PivotNativeLabelFormat[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const hex = (value: unknown) => (typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined);
  const out: PivotNativeLabelFormat[] = [];
  for (const entry of raw.slice(0, 16)) {
    if (typeof entry !== "object" || entry === null) continue;
    const record = entry as Record<string, unknown>;
    if (!Number.isInteger(record.field) || (record.field as number) < 0) continue;
    const format: PivotNativeLabelFormat = { field: record.field as number };
    if (hex(record.backgroundColor)) format.backgroundColor = hex(record.backgroundColor);
    if (hex(record.color)) format.color = hex(record.color);
    if (typeof record.bold === "boolean") format.bold = record.bold;
    if (typeof record.italic === "boolean") format.italic = record.italic;
    if (Object.keys(format).length > 1) out.push(format);
  }
  return out.length ? out : undefined;
}

/**
 * A report's format on one of its areas (Excel's `format/pivotArea`): the
 * grand total row or column (its label too when not `dataOnly`), a field's
 * button (the cell of its caption), or the values under chosen items of the
 * axis fields and, when `valueIndex` is stated, of one value field.
 */
export type PivotNativeAreaFormat = (
  | { area: "grandRow" | "grandCol"; dataOnly?: false }
  | { area: "button"; field: number }
  | { area: "data"; items: { field: number; item: string }[]; valueIndex?: number }
) & { backgroundColor?: string; color?: string; bold?: boolean; italic?: boolean };

export function normalizePivotNativeAreaFormats(raw: unknown): PivotNativeAreaFormat[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const hex = (value: unknown) => (typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined);
  const field = (value: unknown) => (Number.isInteger(value) && (value as number) >= 0 ? (value as number) : undefined);
  const out: PivotNativeAreaFormat[] = [];
  for (const entry of raw.slice(0, 32)) {
    if (typeof entry !== "object" || entry === null) continue;
    const record = entry as Record<string, unknown>;
    const look = {
      ...(hex(record.backgroundColor) ? { backgroundColor: hex(record.backgroundColor) } : {}),
      ...(hex(record.color) ? { color: hex(record.color) } : {}),
      ...(typeof record.bold === "boolean" ? { bold: record.bold } : {}),
      ...(typeof record.italic === "boolean" ? { italic: record.italic } : {}),
    };
    if (!Object.keys(look).length) continue;
    if (record.area === "grandRow" || record.area === "grandCol") {
      out.push({ area: record.area, ...(record.dataOnly === false ? { dataOnly: false as const } : {}), ...look });
    } else if (record.area === "button" && field(record.field) !== undefined) {
      out.push({ area: "button", field: field(record.field)!, ...look });
    } else if (record.area === "data" && Array.isArray(record.items)) {
      const items = record.items.slice(0, 8).flatMap((item) => {
        const own = item as Record<string, unknown>;
        return typeof item === "object" && item !== null && field(own.field) !== undefined && typeof own.item === "string"
          ? [{ field: field(own.field)!, item: own.item }]
          : [];
      });
      const valueIndex = field(record.valueIndex);
      if (items.length !== (record.items as unknown[]).length || (!items.length && valueIndex === undefined)) continue;
      out.push({ area: "data", items, ...(valueIndex !== undefined ? { valueIndex } : {}), ...look });
    }
  }
  return out.length ? out : undefined;
}

/** One cached row or column item: kind (0 item, 1 subtotal, 2 grand total,
 * 3 blank line), its `r` (members repeated from the item above) and the
 * count of its own members. */
export type PivotNativeSavedItem = readonly [0 | 1 | 2 | 3, number, number];

export interface PivotNativeSavedLayout {
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
  rowFieldCount: number;
  pageFieldCount?: number;
  rowPageCount?: number;
  colPageCount?: number;
  pageOverThenDown?: true;
  rowItems: readonly PivotNativeSavedItem[];
  colItems: readonly PivotNativeSavedItem[];
}

const PIVOT_NATIVE_MAX_SAVED_ITEMS = 20_000;

function normalizePivotNativeSavedLayout(raw: unknown): PivotNativeSavedLayout | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const source = raw as Record<string, unknown>;
  const count = (value: unknown, max: number): number | null =>
    typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= max ? value : null;
  const items = (value: unknown): PivotNativeSavedItem[] | null => {
    if (!Array.isArray(value) || value.length > PIVOT_NATIVE_MAX_SAVED_ITEMS) return null;
    const out: PivotNativeSavedItem[] = [];
    for (const item of value) {
      if (!Array.isArray(item) || item.length !== 3) return null;
      const [kind, level, members] = item;
      if ((kind !== 0 && kind !== 1 && kind !== 2 && kind !== 3) || count(level, 255) === null || count(members, 255) === null) {
        return null;
      }
      out.push([kind, level as number, members as number]);
    }
    return out;
  };
  const firstHeaderRow = count(source.firstHeaderRow, MAX_RANGE_ROW);
  const firstDataRow = count(source.firstDataRow, MAX_RANGE_ROW);
  const firstDataCol = count(source.firstDataCol, MAX_RANGE_COL);
  const rowFieldCount = count(source.rowFieldCount, 255);
  const rowItems = items(source.rowItems);
  const colItems = items(source.colItems);
  if (
    firstHeaderRow === null || firstDataRow === null || firstDataCol === null ||
    rowFieldCount === null || !rowItems || !colItems
  ) {
    return undefined;
  }
  const pageFieldCount = source.pageFieldCount === undefined ? undefined : count(source.pageFieldCount, 256);
  const rowPageCount = source.rowPageCount === undefined ? undefined : count(source.rowPageCount, 256);
  const colPageCount = source.colPageCount === undefined ? undefined : count(source.colPageCount, 256);
  if (pageFieldCount === null || rowPageCount === null || colPageCount === null) return undefined;
  return {
    firstHeaderRow,
    firstDataRow,
    firstDataCol,
    rowFieldCount,
    ...(pageFieldCount ? { pageFieldCount, rowPageCount: rowPageCount || 1, colPageCount: colPageCount || 1 } : {}),
    ...(pageFieldCount && source.pageOverThenDown === true ? { pageOverThenDown: true as const } : {}),
    rowItems,
    colItems,
  };
}

const PIVOT_NATIVE_MAX_FIELDS = 128;
// A snapshot's rows are bounded by its cells (the next line): 3 fields hold

const PIVOT_NATIVE_MAX_SNAPSHOT_ROWS = 200_000;
const PIVOT_NATIVE_MAX_SNAPSHOT_CELLS = 200_000;
// A source cell's whole text is its identity: two labels that differ only in
// their 2,001st character are two groups. The cap is the cell's own limit.
const PIVOT_NATIVE_MAX_CELL_TEXT = 32_767;
// A relationship-resolved package part: its folder need not be xl/ (an Open
// XML SDK package keeps its PivotCache at /pivotCache/, K07; X-008). The
// page schema's own rule: canonical forward-slash segments, no traversal.
const PIVOT_NATIVE_PART = /^[A-Za-z0-9_-][A-Za-z0-9_.-]*(?:\/[A-Za-z0-9_-][A-Za-z0-9_.-]*)*\.xml$/i;
const PIVOT_NATIVE_SHA256 = /^[0-9a-f]{64}$/;

function pivotFingerprintFeed(
  rows: readonly (readonly string[])[],
): string {
  let left = 0x811c9dc5;
  let right = 0x9e3779b9;
  const feed = (code: number) => {
    left ^= code;
    left = Math.imul(left, 0x01000193) >>> 0;
    right ^= code + 0x9e3779b9 + ((right << 6) >>> 0) + (right >>> 2);
    right >>>= 0;
  };
  feed(rows.length);
  for (const row of rows) {
    feed(row.length);
    for (const value of row) {
      const text = String(value);
      feed(text.length);
      for (let index = 0; index < text.length; index += 1) {
        feed(text.charCodeAt(index));
      }
      feed(0xffff);
    }
    feed(0xfffe);
  }
  return `${left.toString(16).padStart(8, "0")}${right
    .toString(16)
    .padStart(8, "0")}`;
}

/** The most cells an imported PivotTable's output fingerprint covers. */
const PIVOT_OUTPUT_FINGERPRINT_MAX_CELLS = 200_000;

/**
 * Identify the cells an imported PivotTable occupies, as the page holds them.
 * The server computes the same value (`_xlsx_pivot_output_fingerprint`).
 */
export function pivotOutputFingerprint(
  cells: Readonly<Record<string, unknown>> | undefined,
  outputRef: string,
): string | null {
  const range = normalizePivotRange(outputRef);
  if (
    !cells ||
    !range ||
    (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1) >
      PIVOT_OUTPUT_FINGERPRINT_MAX_CELLS
  ) {
    return null;
  }
  const token = (cell: unknown): string => {
    if (!cell || typeof cell !== "object") return "m\u0000";
    const { formula, value, kind } = cell as Record<string, unknown>;
    if (typeof formula === "string" && formula.trim()) return `f\u0000${formula}`;
    if (value === undefined || value === null || value === "") return "m\u0000";
    return `${kind === "number" ? "n" : "s"}\u0000${String(value)}`;
  };
  const rows: string[][] = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    const line: string[] = [];
    for (let col = range.c1; col <= range.c2; col += 1) {
      line.push(token(cells[`${row}:${col}`]));
    }
    rows.push(line);
  }
  return pivotFingerprintFeed(rows);
}

/** A deterministic identity over exact native scalar kind + lexical value. */
export function pivotCacheSnapshotFingerprint(
  matrix: readonly (readonly PivotNativeScalar[])[],
): string {
  return pivotFingerprintFeed(
    matrix.map((row) => row.map((value) => `${value.kind}\u0000${value.raw}`)),
  );
}

function pivotLegacyDisplayFingerprint(matrix: PivotMatrix): string {
  return pivotFingerprintFeed(matrix);
}

export function pivotNativeScalarDisplay(value: PivotNativeScalar): string {
  if (value.kind === "m") return "";
  if (value.kind === "b") return value.raw === "1" ? "TRUE" : "FALSE";
  return value.raw;
}

function clonePivotMatrix(matrix: readonly (readonly string[])[]): PivotMatrix {
  return matrix.map((row) => row.map((value) => String(value)));
}

function clonePivotScalarMatrix(
  matrix: readonly (readonly PivotNativeScalar[])[],
): PivotNativeScalarMatrix {
  return matrix.map((row) =>
    row.map((value) => ({ kind: value.kind, raw: value.raw })),
  );
}

function pivotConfigCacheSnapshot(
  config: PivotConfig,
): PivotCacheSnapshot | null {
  return config.native?.snapshot ?? config.cacheSnapshot ?? null;
}

/** Freeze the selected source as the cache used by a newly-created Pivot. */
export function createPivotCacheSnapshot(
  fields: readonly PivotInferredField[],
  matrix: PivotMatrix,
  scalarMatrix: PivotNativeScalarMatrix,
): PivotCacheSnapshot | null {
  const width = fields.length;
  if (
    width < 1 ||
    width > PIVOT_NATIVE_MAX_FIELDS ||
    matrix.length < 2 ||
    matrix.length > PIVOT_NATIVE_MAX_SNAPSHOT_ROWS ||
    matrix.length * width > PIVOT_NATIVE_MAX_SNAPSHOT_CELLS ||
    scalarMatrix.length !== matrix.length ||
    matrix.some((row) => row.length !== width) ||
    scalarMatrix.some((row) => row.length !== width) ||
    scalarMatrix.some((row, rowIndex) =>
      row.some(
        (value, columnIndex) =>
          pivotNativeScalarDisplay(value) !== matrix[rowIndex]?.[columnIndex],
      ),
    )
  ) {
    return null;
  }
  const fieldIds = fields.map(
    (field) => `source-field:${field.columnIndex}`,
  );
  if (new Set(fieldIds).size !== fieldIds.length) return null;
  const clonedScalars = clonePivotScalarMatrix(scalarMatrix);
  return {
    revision: 0,
    fingerprint: pivotCacheSnapshotFingerprint(clonedScalars),
    fieldIds,
    matrix: clonePivotMatrix(matrix),
    scalarMatrix: clonedScalars,
  };
}

/**
 * Return the exact data authority used by Pivot evaluation. Imported native
 * pivots and recommended pivots remain pinned to their cache snapshot until an
 * explicit Refresh transaction replaces it. Legacy manual pivots without a
 * snapshot continue to read live cells.
 */
/**
 * The range a PivotTable's source covers NOW. A table source follows its
 * table as it grows or shrinks: header plus data rows (a table's `range` never
 * includes its total row), which is what Excel reads for
 * `<worksheetSource name="Table1"/>`. A plain range source, or a table that no
 * longer exists (deleted, converted to a range, header row off), is the stored
 * `source`. `sourceTables` are the tables of the SOURCE worksheet.
 */
export function pivotLiveSourceRange(
  config: Pick<PivotConfig, "source" | "sourceTableId" | "sourceTableColumn">,
  sourceTables:
    | readonly { id: string; range: PivotRange; headerRow?: boolean }[]
    | undefined,
): PivotRange | null {
  // One column of a table reads that column's own range.
  const table = config.sourceTableId && !config.sourceTableColumn
    ? sourceTables?.find(
        (candidate) =>
          candidate.id === config.sourceTableId && candidate.headerRow !== false,
      )
    : undefined;
  return table
    ? {
        r1: table.range.r1,
        c1: table.range.c1,
        r2: table.range.r2,
        c2: table.range.c2,
      }
    : normalizePivotRange(config.source);
}

export function pivotComputationMatrix(
  config: PivotConfig,
  liveSourceMatrix: PivotMatrix,
): PivotMatrix {
  const snapshot = pivotConfigCacheSnapshot(config);
  return snapshot ? clonePivotMatrix(snapshot.matrix) : liveSourceMatrix;
}

/**
 * Why an imported PivotTable is kept as it came in, as a message key: what
 * its native graph carries that the writer cannot rewrite as Excel would.
 * `blocker` is the importer's comma-joined list (the browser's and the
 * server's share the codes); the first matching family names the reason,
 * the most fundamental first.
 */
export function pivotPreservedReasonKey(blocker: string | undefined): string {
  const codes = (blocker ?? "").split(",").map((code) => code.trim()).filter(Boolean);
  const has = (...tests: Array<string | RegExp>) =>
    codes.some((code) =>
      tests.some((test) => (typeof test === "string" ? code === test : test.test(code))),
    );
  // Past the snapshot's 200,000 cells: nothing else could refresh it, so this
  // is its reason whatever else it carries.
  if (has("cache-records-boundary")) return "spreadsheet.pivotPreserved.sourceTooLarge";
  if (has("external-or-data-model-source", "slicer-connection-olap")) {
    return "spreadsheet.pivotPreserved.dataModel";
  }
  if (has("pivot-filter", "pivot-date-filter", "timeline-pivot-filter")) {
    return "spreadsheet.pivotPreserved.filters";
  }
  if (has("calculated-cache-field")) return "spreadsheet.pivotPreserved.calculated";
  if (has("field-group", "multiple-group-levels", "source-field-authority")) {
    return "spreadsheet.pivotPreserved.grouping";
  }
  if (has(/^slicer-connection-/, /^timeline-/)) return "spreadsheet.pivotPreserved.slicer";
  if (has("pivot-field-sort")) return "spreadsheet.pivotPreserved.sort";
  if (has(/^aggregation-/, /^show-values-as-/, "value-field", "value-field-count", "data-fields-structure", "empty-field-layout")) {
    return "spreadsheet.pivotPreserved.values";
  }
  if (has("pivot-area-format", "value-number-format", "pivot-field-number-format")) {
    return "spreadsheet.pivotPreserved.formats";
  }
  if (
    has(
      "mixed-subtotals",
      "colFields-values-position",
      "axis-count",
      "axis-overlap",
      "pivot-field-caption",
      "page-field",
      "page-field-selection",
      /-field$/,
      /-structure$/,
    )
  ) {
    return "spreadsheet.pivotPreserved.layout";
  }
  return "spreadsheet.pivotPreserved.other";
}

/** Whether a PivotTable is a Scenario PivotTable report (no worksheet source). */
export function pivotIsScenarioReport(
  config: Pick<PivotConfig, "cacheSource" | "cacheSnapshot" | "native">,
): boolean {
  return (
    config.cacheSource === "scenario" &&
    config.cacheSnapshot !== undefined &&
    config.native === undefined
  );
}

export function pivotNativeSourceIsDirty(
  config: PivotConfig,
  liveSourceScalarMatrix: PivotNativeScalarMatrix,
): boolean {
  // A scenario report has no worksheet source to drift from.
  if (pivotIsScenarioReport(config)) return false;
  const snapshot = pivotConfigCacheSnapshot(config);
  return Boolean(
    snapshot &&
      pivotCacheSnapshotFingerprint(liveSourceScalarMatrix) !==
        snapshot.fingerprint,
  );
}

/** Why a Refresh could not replace a PivotTable's cache snapshot. */
export type PivotRefreshRefusal =
  /** The PivotTable reads live cells; there is no snapshot to replace. */
  | "no-snapshot"
  /** An imported PivotTable this editor only preserves. */
  | "preserve-only"
  /** The source range reads as nothing (no rows, or a ragged read). */
  | "source-unreadable"
  /**
   * An imported PivotTable whose source now has columns of the same name
   * twice: Excel's Refresh finds each field again by its name, which two
   * columns cannot share.
   */
  | "duplicate-field-names"
  /**
   * A Scenario PivotTable report: its cache is the scenarios as they were,
   * and Excel's scenario reports are made anew rather than refreshed.
   */
  | "scenario-report";

export type PivotRefreshOutcome =
  | { config: PivotConfig; refusal?: undefined }
  | { config?: undefined; refusal: PivotRefreshRefusal };

/**
 * Build the one permitted Refresh mutation of a PivotTable that computes from
 * a cache snapshot, or say why there is none. The caller still owns live
 * workbook/sheet/source identity revalidation and the single Undo
 * transaction; this pure helper only replaces the authenticated snapshot.
 *
 * Excel's Refresh reads the source as it is NOW (Microsoft Support, "Refresh
 * PivotTable data"): a column added to the source arrives as a new field in
 * the field list, unplaced, and a column no longer there takes its field off
 * the layout. For a PivotTable authored here the snapshot's fields are the
 * source columns in order, so a wider or narrower read simply becomes the new
 * snapshot; the recipe's field offsets were already kept on their columns by
 * the structural remap, and any offset now past the source is dropped by
 * `normalizePivotConfig` at commit. An IMPORTED PivotTable's fields are the
 * cache fields of its OOXML cache definition: a wider or narrower source
 * rebuilds them (`rebasePivotNativeSnapshot`), each placed field found again
 * by its NAME as Excel's Refresh finds it, so a column inserted before a
 * placed field does not move the report onto its neighbour. Refused only when
 * two source columns now share a name.
 */
export function refreshPivotSnapshot(
  config: PivotConfig,
  liveSourceMatrix: PivotMatrix,
  liveSourceScalarMatrix: PivotNativeScalarMatrix,
): PivotRefreshOutcome {
  const native = config.native;
  const snapshot = pivotConfigCacheSnapshot(config);
  if (!snapshot) return { refusal: "no-snapshot" };
  if (pivotIsScenarioReport(config)) return { refusal: "scenario-report" };
  if (native?.disposition === "preserve-only" || native?.blocker) {
    return { refusal: "preserve-only" };
  }
  const width = liveSourceMatrix[0]?.length ?? 0;
  if (
    liveSourceMatrix.length < 1 ||
    width < 1 ||
    liveSourceMatrix.some((row) => row.length !== width) ||
    liveSourceScalarMatrix.length !== liveSourceMatrix.length ||
    liveSourceScalarMatrix.some(
      (row, index) => row.length !== liveSourceMatrix[index]?.length,
    )
  ) {
    return { refusal: "source-unreadable" };
  }
  const matrix = clonePivotMatrix(liveSourceMatrix);
  const scalarMatrix = clonePivotScalarMatrix(liveSourceScalarMatrix);
  if (native && width !== snapshot.fieldIds.length) {
    const fields = pivotNativeFieldsForHeader(native, matrix[0], snapshot.revision + 1);
    if (!fields) return { refusal: "duplicate-field-names" };
    const rebased = rebasePivotNativeSnapshot(
      config,
      {
        revision: snapshot.revision + 1,
        fingerprint: pivotCacheSnapshotFingerprint(scalarMatrix),
        fieldIds: fields.map((field) => field.id),
        matrix,
        scalarMatrix,
      },
      fields,
    );
    return rebased ? { config: rebased } : { refusal: "duplicate-field-names" };
  }
  const fieldIds =
    width === snapshot.fieldIds.length
      ? [...snapshot.fieldIds]
      : Array.from({ length: width }, (_, index) => `source-field:${index}`);
  const nextSnapshot: PivotCacheSnapshot = {
    revision: snapshot.revision + 1,
    fingerprint: pivotCacheSnapshotFingerprint(scalarMatrix),
    fieldIds,
    matrix,
    scalarMatrix,
  };
  if (!native) {
    return { config: { ...config, cacheSnapshot: nextSnapshot } };
  }
  return {
    config: {
      ...config,
      native: {
        ...native,
        snapshot: nextSnapshot,
      },
    },
  };
}

/** The index of the one header cell named `name` (Excel compares field names
 * without regard to case), or null when none or several are. */
function pivotFieldIndexByName(header: readonly string[], name: string): number | null {
  const wanted = name.trim().toLowerCase();
  let found: number | null = null;
  for (let index = 0; index < header.length; index += 1) {
    if (String(header[index] ?? "").trim().toLowerCase() !== wanted) continue;
    if (found !== null) return null;
    found = index;
  }
  return found;
}

/**
 * The cache fields of an imported PivotTable whose source now has `header`:
 * a field whose column is still there (by name) keeps its identity, a new
 * column is a new field. Null when two columns share a name.
 */
export function pivotNativeFieldsForHeader(
  native: PivotNativeAuthority,
  header: readonly string[],
  revision: number,
): PivotNativeFieldAuthority[] | null {
  const seen = new Set<string>();
  for (const cell of header) {
    const key = String(cell ?? "").trim().toLowerCase();
    if (seen.has(key)) return null;
    seen.add(key);
  }
  const prefix = `${native.cacheDefinition.path}#${native.cacheId}:field`;
  const usedIds = new Set<string>();
  return header.map((cell, index) => {
    const name = String(cell ?? "").trim().slice(0, 255) || `Field ${index + 1}`;
    const kept = native.fields.find(
      (field) => field.name.trim().toLowerCase() === name.toLowerCase(),
    );
    let id = kept?.id ?? `${prefix}:r${revision}:${index}`;
    for (let suffix = 2; usedIds.has(id); suffix += 1) {
      id = `${prefix}:r${revision}:${index}:${suffix}`;
    }
    usedIds.add(id);
    return { id, name, sourceIndex: index, cacheIndex: index };
  });
}

/**
 * Every source-field reference of a recipe moved by `map` (old column ->
 * new column, or null when the field is gone): its axes, values, filters,
 * groupings, sorts, item orders, number formats, collapsed items, timeline
 * filters and the base fields of Show Values As. A reference to a field that
 * is gone is dropped, as Excel's Refresh takes a removed field off the
 * layout.
 */
export function remapPivotFieldIndexes(
  config: PivotConfig,
  map: (index: number) => number | null,
): PivotConfig {
  const mapAll = (indexes: readonly number[]) =>
    indexes.flatMap((index) => {
      const next = map(index);
      return next === null ? [] : [next];
    });
  const mapEntries = <T extends { colIndex: number }>(entries: readonly T[] | undefined) =>
    entries?.flatMap((entry) => {
      const next = map(entry.colIndex);
      return next === null ? [] : [{ ...entry, colIndex: next }];
    });
  const rows = mapAll(config.rows);
  const columns = mapAll(config.columns);
  const values = config.values.flatMap((value) => {
    const next = value.calcRef ? (map(value.colIndex) ?? 0) : map(value.colIndex);
    if (next === null) return [];
    const base = value.showAs?.base;
    if (!value.showAs || !base) return [{ ...value, colIndex: next }];
    const oldAxis = base.axis === "row" ? config.rows : config.columns;
    const newAxis = base.axis === "row" ? rows : columns;
    const moved = oldAxis[base.pos] === undefined ? null : map(oldAxis[base.pos]);
    const pos = moved === null ? -1 : newAxis.indexOf(moved);
    const showAs: PivotShowValuesAsConfig =
      pos >= 0 ? { ...value.showAs, base: { ...base, pos } } : { mode: value.showAs.mode };
    return [{ ...value, colIndex: next, showAs }];
  });
  const next: PivotConfig = {
    ...config,
    rows,
    columns,
    values,
    filters: mapEntries(config.filters) ?? [],
    grouping: mapEntries(config.grouping) ?? [],
  };
  const itemFilters = mapEntries(config.itemFilters);
  const fieldSorts = mapEntries(config.fieldSorts);
  const fieldItemOrders = mapEntries(config.fieldItemOrders);
  const fieldNumberFormats = mapEntries(config.fieldNumberFormats);
  const collapsed = mapEntries(config.collapsed);
  const dateFilters = mapEntries(config.dateFilters);
  const fieldFilters = mapEntries(config.fieldFilters);
  const fieldCaptions = mapEntries(config.fieldCaptions);
  const fieldSubtotals = mapEntries(config.fieldSubtotals);
  const calculatedItems = mapEntries(config.calculatedItems);
  delete next.fieldSubtotals;
  delete next.calculatedItems;
  const itemCaptions = mapEntries(config.itemCaptions);
  delete next.fieldFilters;
  delete next.fieldCaptions;
  delete next.itemCaptions;
  delete next.itemFilters;
  delete next.fieldSorts;
  delete next.fieldItemOrders;
  delete next.fieldNumberFormats;
  delete next.collapsed;
  delete next.dateFilters;
  if (itemFilters?.length) next.itemFilters = itemFilters;
  if (fieldSorts?.length) next.fieldSorts = fieldSorts;
  if (fieldItemOrders?.length) next.fieldItemOrders = fieldItemOrders;
  if (fieldNumberFormats?.length) next.fieldNumberFormats = fieldNumberFormats;
  if (collapsed?.length) next.collapsed = collapsed;
  if (dateFilters?.length) next.dateFilters = dateFilters;
  if (fieldFilters?.length) next.fieldFilters = fieldFilters;
  if (fieldCaptions?.length) next.fieldCaptions = fieldCaptions;
  if (fieldSubtotals?.length) next.fieldSubtotals = fieldSubtotals;
  if (calculatedItems?.length) next.calculatedItems = calculatedItems;
  if (itemCaptions?.length) next.itemCaptions = itemCaptions;
  return next;
}

/**
 * An imported PivotTable on a cache whose source fields are now `fields`
 * (Refresh over a source that gained or lost columns): its recipe finds each
 * field again by name and its snapshot is `snapshot`. Every PivotTable on
 * the cache takes the same rebase, each its own layout. Null when a field
 * the recipe or a timeline needs cannot be found and nothing else stands in
 * for it.
 */
export function rebasePivotNativeSnapshot(
  config: PivotConfig,
  snapshot: PivotCacheSnapshot,
  fields: PivotNativeFieldAuthority[],
): PivotConfig | null {
  const native = config.native;
  if (!native) return null;
  const oldHeader = (native.snapshot?.matrix[0] ?? native.fields.map((field) => field.name)).map(
    (cell) => String(cell ?? ""),
  );
  const newHeader = fields.map((field) => field.name);
  const oldWidth = native.fields.length;
  const map = (index: number): number | null => {
    if (!Number.isInteger(index) || index < 0 || index >= oldWidth) return null;
    const name = oldHeader[index];
    return name === undefined ? null : pivotFieldIndexByName(newHeader, name);
  };
  const remapped = remapPivotFieldIndexes(config, map);
  return {
    ...remapped,
    native: { ...native, fields, snapshot },
  };
}

/** `refreshPivotSnapshot` as a config or null (the refusal is dropped). */
export function refreshPivotNativeSnapshot(
  config: PivotConfig,
  liveSourceMatrix: PivotMatrix,
  liveSourceScalarMatrix: PivotNativeScalarMatrix,
): PivotConfig | null {
  return (
    refreshPivotSnapshot(config, liveSourceMatrix, liveSourceScalarMatrix)
      .config ?? null
  );
}

/** Exact bounded recipe admitted by the existing-native Pivot graph writer. */
export function pivotNativeRecipeIsPatchable(config: PivotConfig): boolean {
  const native = config.native;
  if (!native) return true;
  // The writer lays out what Excel does: a fully cleared PivotTable, or fields
  // on the axes and in Values, including the partial layouts of one being
  // built. Without row fields it writes one plain value field.
  const clearedLayout =
    config.rows.length === 0 &&
    config.columns.length === 0 &&
    config.values.length === 0 &&
    config.filters.length === 0 &&
    config.grouping.length === 0;
  if (
    native.disposition !== "editable" ||
    native.blocker ||
    config.calculatedFields.length > 32 ||
    config.rows.length > 16 ||
    config.columns.length > 16 ||
    config.values.length > 8 ||
    config.filters.length > 16 ||
    config.grouping.length > 8 ||
    (!clearedLayout &&
      config.rows.length === 0 &&
      config.columns.length === 0 &&
      config.values.length === 0) ||
    // Without row fields, several value fields beside column fields are
    // laid out in compact form only (Excel-refreshed ClosedXML corpus K06;
    // the other forms are unverified). The xlsx writer applies the same rule.
    (config.rows.length === 0 &&
      ((config.columns.length === 0 &&
        config.values.some((field) => (field.showAs?.mode ?? "normal") !== "normal")) ||
        (config.values.length > 1 &&
          config.columns.length > 0 &&
          (config.reportLayout ?? "tabular") !== "compact")))
  ) {
    return false;
  }
  if (
    (config.subtotalPosition !== undefined &&
      !PIVOT_SUBTOTAL_POSITIONS.includes(config.subtotalPosition)) ||
    (config.grandTotals !== undefined &&
      !PIVOT_GRAND_TOTAL_MODES.includes(config.grandTotals)) ||
    (config.reportLayout !== undefined &&
      !PIVOT_REPORT_LAYOUTS.includes(config.reportLayout)) ||
    (config.styleName !== undefined && !isPivotStyleName(config.styleName)) ||
    (config.repeatItemLabels === true &&
      (config.reportLayout ?? "tabular") === "compact")
  ) {
    return false;
  }
  const source = normalizePivotRange(config.source);
  const output = normalizePivotRange(config.output ?? "");
  const width = native.snapshot.fieldIds.length;
  if (
    !source ||
    !output ||
    output.r1 !== output.r2 ||
    output.c1 !== output.c2 ||
    source.c2 - source.c1 + 1 !== width
  ) {
    return false;
  }
  const axis = [...config.rows, ...config.columns];
  const inRange = (index: number) =>
    Number.isInteger(index) && index >= 0 && index < width;
  if (
    axis.some((index) => !inRange(index)) ||
    new Set(axis).size !== axis.length ||
    config.filters.some((field) => !inRange(field.colIndex)) ||
    new Set(config.filters.map((field) => field.colIndex)).size !==
      config.filters.length ||
    config.filters.some((field) => axis.includes(field.colIndex))
  ) {
    return false;
  }
  const allowedAgg = new Set<PivotAgg>(PIVOT_AGGS);
  // What the writer can name in a dataField: Excel's own showDataAs modes,
  // plus the two ranks it keeps in its 2010 extension.
  const allowedShow = new Set<PivotShowValuesAs>(PIVOT_SHOW_VALUES_AS);
  const baseNamesAnAxisField = (field: PivotValueField): boolean => {
    const mode = field.showAs?.mode ?? "normal";
    const base = field.showAs?.base;
    if (!base) {
      // Only a difference has nothing to measure against without one; the
      // rest fall back to the innermost field of the axis they name.
      return !showValuesAsNeedsBaseItem(mode);
    }
    if (!showValuesAsNeedsBaseField(mode)) return false;
    const axisFields = base.axis === "row" ? config.rows : config.columns;
    return base.pos >= 0 && base.pos < axisFields.length;
  };
  const calcNames = new Set(
    config.calculatedFields.map((field) => field.name.trim().toLowerCase()),
  );
  if (
    config.values.some(
      (field) =>
        // A calculated value names a field the recipe defines; the writer
        // gives it its own cache field, as Excel does.
        (field.calcRef !== undefined &&
          !calcNames.has(field.calcRef.trim().toLowerCase())) ||
        (field.calcRef === undefined && !inRange(field.colIndex)) ||
        !allowedAgg.has(field.agg) ||
        (field.numberFormat !== undefined &&
          (!field.numberFormat.trim() ||
            field.numberFormat.length > MAX_PIVOT_VALUE_NUMBER_FORMAT ||
            /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(
              field.numberFormat,
            ))) ||
        !allowedShow.has(field.showAs?.mode ?? "normal") ||
        !baseNamesAnAxisField(field),
    )
  ) {
    return false;
  }
  const grouped = new Set<string>();
  for (const grouping of config.grouping) {
    const key =
      grouping.kind === "date" ? `${grouping.colIndex}:${grouping.unit}` : `${grouping.colIndex}`;
    if (
      !inRange(grouping.colIndex) ||
      (grouping.kind !== "discrete" && !axis.includes(grouping.colIndex)) ||
      grouped.has(key) ||
      (grouping.kind === "discrete" && !(grouping.groups && grouping.groups.length > 0)) ||
      (grouping.kind === "number" && grouped.has(`${grouping.colIndex}:date`)) ||
      (grouping.kind === "number" &&
        (!(grouping.size && grouping.size > 0) ||
          !Number.isFinite(grouping.size) ||
          (grouping.start !== undefined && !Number.isFinite(grouping.start)))) ||
      (grouping.kind === "date" &&
        !["year", "quarter", "month"].includes(grouping.unit ?? ""))
    ) {
      return false;
    }
    grouped.add(key);
    if (grouping.kind === "date") grouped.add(`${grouping.colIndex}:date`);
  }
  if (
    (config.fieldFilters ?? []).some(
      (filter) =>
        !inRange(filter.colIndex) ||
        (filter.valueIndex !== undefined && filter.valueIndex >= config.values.length),
    )
  ) {
    return false;
  }
  const dateColumns = (config.dateFilters ?? []).map((filter) => filter.colIndex);
  if (
    dateColumns.length > MAX_PIVOT_DATE_FILTERS ||
    dateColumns.some((index) => !inRange(index)) ||
    new Set(dateColumns).size !== dateColumns.length ||
    config.filters.some((field) => dateColumns.includes(field.colIndex))
  ) {
    return false;
  }
  return true;
}

function boundedNativeText(value: unknown, max: number): string | null {
  if (typeof value !== "string") return null;
  const text = value.trim();
  return text.length > 0 && text.length <= max ? text : null;
}

function normalizePivotNativePart(value: unknown): PivotNativePartAuthority | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const path = boundedNativeText(source.path, 512);
  const sha256 = boundedNativeText(source.sha256, 64);
  if (!path || !PIVOT_NATIVE_PART.test(path) || !sha256 || !PIVOT_NATIVE_SHA256.test(sha256)) {
    return null;
  }
  return { path, sha256 };
}

function normalizePivotCacheSnapshot(
  input: unknown,
  expectedWidth?: number,
): PivotCacheSnapshot | null {
  if (!input || typeof input !== "object") return null;
  const source = input as Record<string, unknown>;
  const revision = source.revision;
  const fingerprint = boundedNativeText(source.fingerprint, 64);
  const fieldIds = Array.isArray(source.fieldIds) ? source.fieldIds : [];
  const rawMatrix = Array.isArray(source.matrix) ? source.matrix : [];
  const rawScalarMatrix = Array.isArray(source.scalarMatrix)
    ? source.scalarMatrix
    : [];
  const width = fieldIds.length;
  if (
    !Number.isInteger(revision) ||
    (revision as number) < 0 ||
    !fingerprint ||
    width < 1 ||
    width > PIVOT_NATIVE_MAX_FIELDS ||
    (expectedWidth !== undefined && width !== expectedWidth) ||
    rawMatrix.length < 2 ||
    rawMatrix.length > PIVOT_NATIVE_MAX_SNAPSHOT_ROWS ||
    rawScalarMatrix.length !== rawMatrix.length
  ) {
    return null;
  }
  const normalizedIds: string[] = [];
  for (const value of fieldIds) {
    const fieldId = boundedNativeText(value, 512);
    if (!fieldId) return null;
    normalizedIds.push(fieldId);
  }
  if (new Set(normalizedIds).size !== normalizedIds.length) return null;

  const matrix: PivotMatrix = [];
  const scalarMatrix: PivotNativeScalarMatrix = [];
  let cellCount = 0;
  for (let rowIndex = 0; rowIndex < rawMatrix.length; rowIndex += 1) {
    const rawRow = rawMatrix[rowIndex];
    const rawScalarRow = rawScalarMatrix[rowIndex];
    if (
      !Array.isArray(rawRow) ||
      !Array.isArray(rawScalarRow) ||
      rawRow.length !== width ||
      rawScalarRow.length !== width
    ) {
      return null;
    }
    cellCount += width;
    if (cellCount > PIVOT_NATIVE_MAX_SNAPSHOT_CELLS) return null;
    const row = rawRow.map((value) =>
      String(value).slice(0, PIVOT_NATIVE_MAX_CELL_TEXT),
    );
    const scalarRow: PivotNativeScalar[] = [];
    for (const rawValue of rawScalarRow) {
      if (!rawValue || typeof rawValue !== "object") return null;
      const candidate = rawValue as Record<string, unknown>;
      const kind = candidate.kind;
      const raw = candidate.raw;
      if (
        !["s", "n", "d", "b", "e", "m"].includes(String(kind)) ||
        typeof raw !== "string" ||
        raw.length > PIVOT_NATIVE_MAX_CELL_TEXT ||
        (kind === "m" && raw !== "") ||
        (kind === "b" && raw !== "0" && raw !== "1")
      ) {
        return null;
      }
      scalarRow.push({ kind: kind as PivotNativeScalarKind, raw });
    }
    if (
      scalarRow.some(
        (value, columnIndex) =>
          pivotNativeScalarDisplay(value) !== row[columnIndex],
      )
    ) {
      return null;
    }
    matrix.push(row);
    scalarMatrix.push(scalarRow);
  }
  if (pivotCacheSnapshotFingerprint(scalarMatrix) !== fingerprint) return null;
  return {
    revision: revision as number,
    fingerprint,
    fieldIds: normalizedIds,
    matrix,
    scalarMatrix,
  };
}

/**
 * Validate imported native custody without repairing or guessing it. Any
 * malformed identity returns null; callers then keep the Pivot preserve-only.
 */
/**
 * The receipt's item filters exactly as imported (order and all), or null when
 * malformed. The save path compares the receipt byte for byte with the one the
 * import wrote, so nothing here may be repaired or reordered.
 */
function normalizePivotNativeItemFilters(
  input: unknown,
  fieldCount: number,
): PivotFilterField[] | null {
  if (!Array.isArray(input) || input.length > MAX_PIVOT_ITEM_FILTERS) return null;
  const filters: PivotFilterField[] = [];
  for (const raw of input) {
    if (!raw || typeof raw !== "object") return null;
    const entry = raw as Record<string, unknown>;
    const { colIndex, selected, unit, group } = entry;
    if (
      Object.keys(entry).some(
        (key) => key !== "colIndex" && key !== "selected" && key !== "unit" && key !== "group",
      ) ||
      (unit !== undefined && !PIVOT_DATE_UNITS.includes(unit as PivotDateUnit)) ||
      (group !== undefined && typeof group !== "boolean") ||
      !Number.isInteger(colIndex) ||
      (colIndex as number) < 0 ||
      (colIndex as number) >= fieldCount ||
      !Array.isArray(selected) ||
      selected.length > MAX_PIVOT_FILTER_SELECTED ||
      selected.some(
        (value) => typeof value !== "string" || value.length > MAX_PIVOT_FILTER_VALUE,
      )
    ) {
      return null;
    }
    filters.push({
      colIndex: colIndex as number,
      selected: [...(selected as string[])],
      ...(unit === undefined ? {} : { unit: unit as PivotDateUnit }),
      ...(group === true ? { group: true as const } : {}),
    });
  }
  return filters;
}

export function normalizePivotNativeAuthority(
  input: unknown,
): PivotNativeAuthority | null {
  if (!input || typeof input !== "object") return null;
  const source = input as Record<string, unknown>;
  if (source.version !== 1) return null;
  const disposition =
    source.disposition === "editable" || source.disposition === "preserve-only"
      ? source.disposition
      : null;
  const blocker =
    source.blocker === undefined ? undefined : boundedNativeText(source.blocker, 200);
  if (!disposition || (source.blocker !== undefined && !blocker)) return null;
  if (disposition === "preserve-only" && !blocker) return null;
  if (disposition === "editable" && blocker) return null;

  const pivotName = boundedNativeText(source.pivotName, 255);
  const sourceSheetName = boundedNativeText(source.sourceSheetName, 255);
  const sourceRef = boundedNativeText(source.sourceRef, 64);
  const outputRef = boundedNativeText(source.outputRef, 64);
  const worksheetRelationshipId = boundedNativeText(
    source.worksheetRelationshipId,
    255,
  );
  const cacheRelationshipId = boundedNativeText(source.cacheRelationshipId, 255);
  const recordsRelationshipId =
    source.recordsRelationshipId === undefined
      ? undefined
      : boundedNativeText(source.recordsRelationshipId, 255);
  const cacheId = source.cacheId;
  const pivotTable = normalizePivotNativePart(source.pivotTable);
  const cacheDefinition = normalizePivotNativePart(source.cacheDefinition);
  const cacheRecords =
    source.cacheRecords === undefined
      ? undefined
      : normalizePivotNativePart(source.cacheRecords);
  if (
    !pivotName ||
    !sourceSheetName ||
    !sourceRef ||
    !outputRef ||
    !worksheetRelationshipId ||
    !cacheRelationshipId ||
    (source.recordsRelationshipId !== undefined && !recordsRelationshipId) ||
    !Number.isInteger(cacheId) ||
    (cacheId as number) < 0 ||
    !pivotTable ||
    !cacheDefinition ||
    (source.cacheRecords !== undefined && !cacheRecords)
  ) {
    return null;
  }

  const rawFields = Array.isArray(source.fields) ? source.fields : [];
  if (rawFields.length === 0 || rawFields.length > PIVOT_NATIVE_MAX_FIELDS) return null;
  const fields: PivotNativeFieldAuthority[] = [];
  const fieldIds = new Set<string>();
  const cacheIndexes = new Set<number>();
  for (const raw of rawFields) {
    if (!raw || typeof raw !== "object") return null;
    const field = raw as Record<string, unknown>;
    const id = boundedNativeText(field.id, 512);
    const name = boundedNativeText(field.name, 255);
    const sourceIndex = field.sourceIndex;
    const cacheIndex = field.cacheIndex;
    const groupedFromCacheIndex = field.groupedFromCacheIndex;
    if (
      !id ||
      !name ||
      fieldIds.has(id) ||
      !Number.isInteger(sourceIndex) ||
      (sourceIndex as number) < 0 ||
      !Number.isInteger(cacheIndex) ||
      (cacheIndex as number) < 0 ||
      cacheIndexes.has(cacheIndex as number) ||
      (groupedFromCacheIndex !== undefined &&
        (!Number.isInteger(groupedFromCacheIndex) ||
          (groupedFromCacheIndex as number) < 0))
    ) {
      return null;
    }
    fieldIds.add(id);
    cacheIndexes.add(cacheIndex as number);
    fields.push({
      id,
      name,
      sourceIndex: sourceIndex as number,
      cacheIndex: cacheIndex as number,
      ...(groupedFromCacheIndex === undefined
        ? {}
        : { groupedFromCacheIndex: groupedFromCacheIndex as number }),
    });
  }

  if (!source.snapshot || typeof source.snapshot !== "object") return null;
  const rawSnapshot = source.snapshot as Record<string, unknown>;
  const revision = rawSnapshot.revision;
  const fingerprint = boundedNativeText(rawSnapshot.fingerprint, 64);
  const snapshotFieldIds = Array.isArray(rawSnapshot.fieldIds)
    ? rawSnapshot.fieldIds
    : [];
  const rawMatrix = Array.isArray(rawSnapshot.matrix) ? rawSnapshot.matrix : [];
  if (
    !Number.isInteger(revision) ||
    (revision as number) < 0 ||
    !fingerprint ||
    snapshotFieldIds.length !== fields.length ||
    rawMatrix.length < 1 ||
    rawMatrix.length > PIVOT_NATIVE_MAX_SNAPSHOT_ROWS
  ) {
    return null;
  }
  const matrix: string[][] = [];
  let cells = 0;
  for (const rawRow of rawMatrix) {
    if (!Array.isArray(rawRow) || rawRow.length !== fields.length) return null;
    cells += rawRow.length;
    if (cells > PIVOT_NATIVE_MAX_SNAPSHOT_CELLS) return null;
    const row = rawRow.map((value) => String(value).slice(0, PIVOT_NATIVE_MAX_CELL_TEXT));
    matrix.push(row);
  }
  const rawScalarMatrix = Array.isArray(rawSnapshot.scalarMatrix)
    ? rawSnapshot.scalarMatrix
    : [];
  const scalarMatrix: PivotNativeScalarMatrix = [];
  const scalarAuthorityMissing = rawScalarMatrix.length === 0;
  if (!scalarAuthorityMissing) {
    if (rawScalarMatrix.length !== matrix.length) return null;
    for (const [rowIndex, rawRow] of rawScalarMatrix.entries()) {
      if (!Array.isArray(rawRow) || rawRow.length !== fields.length) return null;
      const row: PivotNativeScalar[] = [];
      for (const rawValue of rawRow) {
        if (!rawValue || typeof rawValue !== "object") return null;
        const candidate = rawValue as Record<string, unknown>;
        const kind = candidate.kind;
        const raw = candidate.raw;
        if (
          !["s", "n", "d", "b", "e", "m"].includes(String(kind)) ||
          typeof raw !== "string" ||
          raw.length > PIVOT_NATIVE_MAX_CELL_TEXT ||
          (kind === "m" && raw !== "") ||
          (kind === "b" && raw !== "0" && raw !== "1")
        ) {
          return null;
        }
        row.push({ kind: kind as PivotNativeScalarKind, raw });
      }
      if (
        row.some(
          (value, columnIndex) =>
            pivotNativeScalarDisplay(value) !== matrix[rowIndex]?.[columnIndex],
        )
      ) {
        return null;
      }
      scalarMatrix.push(row);
    }
  } else {
    for (const row of matrix) {
      scalarMatrix.push(row.map((raw) => ({ kind: "s", raw })));
    }
  }
  if (
    snapshotFieldIds.some(
      (value, index) => value !== fields[index]?.id,
    ) ||
    (scalarAuthorityMissing
      ? pivotLegacyDisplayFingerprint(matrix)
      : pivotCacheSnapshotFingerprint(scalarMatrix)) !== fingerprint
  ) {
    return null;
  }
  const styleName =
    source.styleName === undefined ? undefined : boundedNativeText(source.styleName, 64);
  if (source.styleName !== undefined && !styleName) return null;
  // The receipt is immutable identity: a malformed one invalidates custody
  // instead of being silently repaired.
  const nativeItemFilters =
    source.itemFilters === undefined
      ? undefined
      : normalizePivotNativeItemFilters(source.itemFilters, fields.length);
  if (nativeItemFilters === null) return null;
  const nativeDateFilters =
    source.dateFilters === undefined
      ? undefined
      : normalizePivotDateFilters(source.dateFilters, fields.length - 1);
  if (
    nativeDateFilters !== undefined &&
    (!Array.isArray(source.dateFilters) ||
      nativeDateFilters.length !== source.dateFilters.length)
  ) {
    return null;
  }
  const nativeFieldFilters =
    source.fieldFilters === undefined
      ? undefined
      : normalizePivotFieldFilters(source.fieldFilters, fields.length - 1);
  if (
    nativeFieldFilters !== undefined &&
    (!Array.isArray(source.fieldFilters) ||
      nativeFieldFilters.length !== source.fieldFilters.length)
  ) {
    return null;
  }
  const outputFingerprint = source.outputFingerprint;
  if (
    (outputFingerprint !== undefined &&
      (typeof outputFingerprint !== "string" ||
        !/^[0-9a-f]{16}$/.test(outputFingerprint))) ||
    (source.layoutReading !== undefined && source.layoutReading !== 2) ||
    (source.classicLayout !== undefined && source.classicLayout !== true) ||
    (source.hideValuesRow !== undefined && source.hideValuesRow !== true) ||
    (source.dataCaption !== undefined &&
      (typeof source.dataCaption !== "string" || !source.dataCaption.trim() || source.dataCaption.length > 255)) ||
    (source.refreshOnLoad !== undefined && source.refreshOnLoad !== true)
  ) {
    return null;
  }
  const savedLayout = normalizePivotNativeSavedLayout(source.savedLayout);
  const labelFormats = normalizePivotNativeLabelFormats(source.labelFormats);
  const areaFormats = normalizePivotNativeAreaFormats(source.areaFormats);
  const subtotalTopFields = Array.isArray(source.subtotalTopFields)
    ? source.subtotalTopFields.filter((field): field is number => Number.isInteger(field) && (field as number) >= 0).slice(0, 128)
    : [];
  return {
    version: 1,
    disposition: scalarAuthorityMissing ? "preserve-only" : disposition,
    ...((scalarAuthorityMissing
      ? [blocker, "native-cache-scalar-authority-missing"].filter(Boolean).join(",")
      : blocker)
      ? {
          blocker: (scalarAuthorityMissing
            ? [blocker, "native-cache-scalar-authority-missing"]
                .filter(Boolean)
                .join(",")
            : blocker) as string,
        }
      : {}),
    pivotName,
    sourceSheetName,
    sourceRef,
    outputRef,
    ...(styleName ? { styleName } : {}),
    worksheetRelationshipId,
    cacheId: cacheId as number,
    cacheRelationshipId,
    ...(recordsRelationshipId ? { recordsRelationshipId } : {}),
    pivotTable,
    cacheDefinition,
    ...(cacheRecords ? { cacheRecords } : {}),
    fields,
    snapshot: {
      revision: revision as number,
      fingerprint,
      fieldIds: fields.map((field) => field.id),
      matrix,
      scalarMatrix,
    },
    ...(typeof source.cacheKey === "string" &&
    source.cacheKey.length > 0 &&
    source.cacheKey.length <= 512
      ? { cacheKey: source.cacheKey }
      : {}),
    ...(nativeItemFilters !== undefined ? { itemFilters: nativeItemFilters } : {}),
    ...(nativeDateFilters !== undefined ? { dateFilters: nativeDateFilters } : {}),
    ...(nativeFieldFilters !== undefined ? { fieldFilters: nativeFieldFilters } : {}),
    ...(outputFingerprint !== undefined ? { outputFingerprint } : {}),
    ...(source.layoutReading === 2 ? { layoutReading: 2 as const } : {}),
    ...(source.classicLayout === true ? { classicLayout: true as const } : {}),
    ...(source.hideValuesRow === true ? { hideValuesRow: true as const } : {}),
    ...(typeof source.dataCaption === "string" ? { dataCaption: source.dataCaption } : {}),
    ...(source.refreshOnLoad === true ? { refreshOnLoad: true as const } : {}),
    ...(savedLayout ? { savedLayout } : {}),
    ...(labelFormats ? { labelFormats } : {}),
    ...(areaFormats ? { areaFormats } : {}),
    ...(subtotalTopFields.length ? { subtotalTopFields } : {}),
  };
}

/** Every PivotTable of a sheet: its first (`pivot`), then `morePivots`. */
export function sheetPivots<P>(sheet: {
  pivot?: P;
  morePivots?: readonly P[];
}): P[] {
  return [
    ...(sheet.pivot !== undefined ? [sheet.pivot] : []),
    ...(sheet.morePivots ?? []),
  ];
}

/** The PivotTable in a sheet's slot: 0 is `pivot`, k is `morePivots[k-1]`. */
export function sheetPivotAt<P>(
  sheet: { pivot?: P; morePivots?: readonly P[] } | null | undefined,
  slot: number,
): P | undefined {
  if (!sheet) return undefined;
  return slot <= 0 ? sheet.pivot : sheet.morePivots?.[slot - 1];
}

/**
 * The sheet with one slot's PivotTable replaced, or removed (`undefined`): the
 * PivotTables after a removed one move up, so `pivot` stays the first.
 */
export function sheetWithPivotAt<
  S extends { pivot?: unknown; morePivots?: unknown[] },
>(sheet: S, slot: number, pivot: NonNullable<S["pivot"]> | undefined): S {
  const all: unknown[] = sheetPivots(sheet);
  const index = Math.max(0, slot);
  if (pivot === undefined) all.splice(index, 1);
  else if (index < all.length) all[index] = pivot;
  else all.push(pivot);
  const next: { pivot?: unknown; morePivots?: unknown[] } = { ...sheet };
  if (all.length > 0) next.pivot = all[0];
  else delete next.pivot;
  if (all.length > 1) next.morePivots = all.slice(1);
  else delete next.morePivots;
  return next as S;
}

/** The key one PivotTable's derived report is planned under. */
export function pivotPlanKey(sheetId: string, slot: number): string {
  return slot <= 0 ? sheetId : `${sheetId}#pivot:${slot}`;
}

/**
 * A raw persisted PivotTable carrying the snapshot it stands on: a PivotTable
 * on a shared cache keeps none of its own and names the workbook's entry.
 */
export function pivotWithSharedSnapshot(
  rawPivot: unknown,
  pivotCaches: unknown,
): unknown {
  if (!rawPivot || typeof rawPivot !== "object") return rawPivot;
  const native = (rawPivot as { native?: unknown }).native;
  if (!native || typeof native !== "object") return rawPivot;
  const { snapshot, cacheKey } = native as {
    snapshot?: unknown;
    cacheKey?: unknown;
  };
  if (snapshot !== undefined || typeof cacheKey !== "string") return rawPivot;
  const shared = expandSharedPivotCache(
    pivotCaches && typeof pivotCaches === "object"
      ? (pivotCaches as Record<string, unknown>)[cacheKey]
      : undefined,
  );
  return shared === null
    ? rawPivot
    : { ...rawPivot, native: { ...native, snapshot: shared } };
}

const SHARED_PIVOT_CACHE_KINDS = new Set(["s", "n", "d", "b", "e", "m"]);

/**
 * A shared cache as the workbook keeps it: one string per value, its OOXML
 * kind then its lexical value ("n12.5", "sBaden", "m"). The display matrix
 * follows from the values, so it is not stored a second time. Mirror of
 * `xlsx_compact_shared_pivot_cache` on the server.
 */
export function compactSharedPivotCache(snapshot: {
  revision?: number;
  fingerprint?: string;
  fieldIds?: readonly string[];
  scalarMatrix?: ReadonlyArray<ReadonlyArray<{ kind: string; raw: string }>>;
}): Record<string, unknown> {
  return {
    revision: snapshot.revision ?? 0,
    fingerprint: snapshot.fingerprint,
    fieldIds: [...(snapshot.fieldIds ?? [])],
    records: (snapshot.scalarMatrix ?? []).map((row) =>
      row.map((value) => `${value.kind}${value.raw}`),
    ),
  };
}

/** The snapshot a PivotTable reads, from the workbook's compact entry. */
export function expandSharedPivotCache(
  shared: unknown,
): Record<string, unknown> | null {
  if (!shared || typeof shared !== "object") return null;
  const { revision, fingerprint, fieldIds, records } = shared as Record<string, unknown>;
  if (!Array.isArray(records) || !Array.isArray(fieldIds)) return null;
  const scalarMatrix: Array<Array<{ kind: string; raw: string }>> = [];
  for (const row of records) {
    if (!Array.isArray(row)) return null;
    const scalars: Array<{ kind: string; raw: string }> = [];
    for (const value of row) {
      if (typeof value !== "string" || !SHARED_PIVOT_CACHE_KINDS.has(value[0] ?? "")) {
        return null;
      }
      scalars.push({ kind: value[0], raw: value.slice(1) });
    }
    scalarMatrix.push(scalars);
  }
  return {
    revision: typeof revision === "number" ? revision : 0,
    fingerprint,
    fieldIds: [...fieldIds],
    matrix: scalarMatrix.map((row) =>
      row.map((value) =>
        value.kind === "m"
          ? ""
          : value.kind === "b"
            ? value.raw === "1"
              ? "TRUE"
              : "FALSE"
            : value.raw,
      ),
    ),
    scalarMatrix,
  };
}

/**
 * In memory, PivotTables on one cache hold ONE snapshot object: normalizing
 * each would otherwise keep a copy of the cache per PivotTable.
 */
export function unifyPivotCacheSnapshots<
  S extends { pivot?: PivotConfig; morePivots?: PivotConfig[] },
>(sheets: S[]): S[] {
  const byKey = new Map<string, PivotNativeCacheSnapshot>();
  const unify = (pivot: PivotConfig): PivotConfig => {
    const key = pivot.native?.cacheKey;
    if (!key || !pivot.native) return pivot;
    const known = byKey.get(key);
    if (!known) {
      byKey.set(key, pivot.native.snapshot);
      return pivot;
    }
    return known === pivot.native.snapshot ||
      known.fingerprint !== pivot.native.snapshot.fingerprint
      ? pivot
      : { ...pivot, native: { ...pivot.native, snapshot: known } };
  };
  return sheets.map((sheet) => {
    if (!sheet.pivot && !sheet.morePivots?.length) return sheet;
    return {
      ...sheet,
      ...(sheet.pivot ? { pivot: unify(sheet.pivot) } : {}),
      ...(sheet.morePivots ? { morePivots: sheet.morePivots.map(unify) } : {}),
    };
  });
}

/**
 * The persisted form of shared caches: each snapshot once under `pivotCaches`,
 * taken off the PivotTables that name it. The most refreshed snapshot of a
 * cache speaks for it.
 */
export function sharePivotCacheSnapshots<
  S extends { pivot?: unknown; morePivots?: unknown },
>(sheets: S[]): { sheets: S[]; pivotCaches?: Record<string, unknown> } {
  const newest: Record<string, { revision?: number }> = {};
  const strip = (raw: unknown): unknown => {
    if (!raw || typeof raw !== "object") return raw;
    const native = (raw as { native?: unknown }).native as
      | { cacheKey?: unknown; snapshot?: { revision?: number } }
      | undefined;
    if (!native || typeof native.cacheKey !== "string" || !native.snapshot) {
      return raw;
    }
    const known = newest[native.cacheKey];
    if (!known || (native.snapshot.revision ?? 0) > (known.revision ?? 0)) {
      newest[native.cacheKey] = native.snapshot;
    }
    const { snapshot: _snapshot, ...rest } = native;
    return { ...raw, native: rest };
  };
  let touched = false;
  const next = sheets.map((sheet) => {
    const pivot = strip(sheet.pivot);
    const more = Array.isArray(sheet.morePivots)
      ? sheet.morePivots.map(strip)
      : sheet.morePivots;
    const changed =
      pivot !== sheet.pivot ||
      (Array.isArray(more) &&
        more.some((item, index) => item !== (sheet.morePivots as unknown[])[index]));
    if (!changed) return sheet;
    touched = true;
    return {
      ...sheet,
      ...(sheet.pivot !== undefined ? { pivot } : {}),
      ...(sheet.morePivots !== undefined ? { morePivots: more } : {}),
    } as S;
  });
  if (!touched) return { sheets };
  return {
    sheets: next,
    pivotCaches: Object.fromEntries(
      Object.entries(newest).map(([key, snapshot]) => [
        key,
        compactSharedPivotCache(snapshot),
      ]),
    ),
  };
}







export function parseGroupableDate(
  value: string,
): { year: number; month: number; day: number } | null {
  const s = value.trim();
  let m: RegExpMatchArray | null;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/))) {
    const [, y, mo, d] = m;
    return finiteDate(+y, +mo, +d);
  }
  if ((m = s.match(/^(\d{4})-(\d{1,2})$/))) {
    const [, y, mo] = m;
    return finiteDate(+y, +mo, 1);
  }
  if ((m = s.match(/^(\d{4})$/))) {
    return finiteDate(+m[1], 1, 1);
  }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
    const [, mo, d, y] = m; // US M/D/YYYY (matches the editor's auto-date)
    return finiteDate(+y, +mo, +d);
  }
  return null;
}

function finiteDate(
  year: number,
  month: number,
  day: number,
): { year: number; month: number; day: number } | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { year, month, day };
}

/**
 * The bucket label a grouped field value collapses to, or the raw string when the
 * value cannot be parsed for the grouping kind (so it stays its own group rather
 * than silently merging). Pure + exported so the bucketing is unit-testable apart
 * from the full pivot compute.
 */
export const PIVOT_MONTH_LABELS = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
] as const;
export const PIVOT_QUARTER_LABELS = ["Qtr1", "Qtr2", "Qtr3", "Qtr4"] as const;

/**
 * The group item a value falls in, spelled as Excel spells its group items:
 * "Jan", "Qtr1", "2024", and "0-9" / "10-19" for a number bin. Grouping by
 * months alone merges the Januaries of every year, as Excel does; Years is
 * its own level (see `expandPivotGroupLevels`). Mirror of the server writer.
 */
export function pivotBucketLabel(value: string, grouping: PivotGrouping): string {
  if (grouping.kind === "discrete") {
    return grouping.groups?.find((group) => group.items.includes(value))?.name ?? value;
  }
  if (grouping.kind === "number") {
    const n = coerceNumber(value);
    if (n === null) return value;
    const size = grouping.size && grouping.size > 0 ? grouping.size : 1;
    const start = Number.isFinite(grouping.start) ? (grouping.start as number) : 0;
    const lo = start + Math.floor((n - start) / size) * size;
    const integerBins = Number.isInteger(size) && Number.isInteger(start);
    return `${trimNum(lo)}-${trimNum(integerBins ? lo + size - 1 : lo + size)}`;
  }
  const parts = parseGroupableDate(value);
  if (!parts) return value;
  if (grouping.unit === "year") return String(parts.year);
  if (grouping.unit === "quarter") {
    return PIVOT_QUARTER_LABELS[Math.floor((parts.month - 1) / 3)];
  }
  return PIVOT_MONTH_LABELS[parts.month - 1];
}

const PIVOT_DATE_UNIT_RANK: Record<PivotDateUnit, number> = {
  year: 0,
  quarter: 1,
  month: 2,
};
const PIVOT_DERIVED_LEVEL_NAMES: Record<PivotDateUnit, string> = {
  year: "Years",
  quarter: "Quarters",
  month: "Months",
};

/**
 * The order group items sort in on an axis: calendar order for months and
 * quarters, ascending for years and number bins (Excel lists them so, not
 * first-seen). `null` when the field is not grouped.
 */
export function pivotGroupItemOrder(
  grouping: PivotGrouping | undefined,
  labels: Iterable<string>,
): Map<string, number> | null {
  if (!grouping) return null;
  if (grouping.kind === "discrete") {
    const order = new Map<string, number>();
    const names = new Set((grouping.groups ?? []).map((group) => group.name));
    for (const label of labels) if (!names.has(label) && !order.has(label)) order.set(label, order.size);
    for (const group of grouping.groups ?? []) if (!order.has(group.name)) order.set(group.name, order.size);
    return order;
  }
  const rank = (label: string): number => {
    if (grouping.kind === "date") {
      const table: readonly string[] =
        grouping.unit === "month"
          ? PIVOT_MONTH_LABELS
          : grouping.unit === "quarter"
            ? PIVOT_QUARTER_LABELS
            : [];
      const place = table.indexOf(label);
      if (place >= 0) return place;
      const year = Number(label);
      return Number.isFinite(year) ? year : Number.POSITIVE_INFINITY;
    }
    const low = Number(label.split("-")[0]);
    return Number.isFinite(low) ? low : Number.POSITIVE_INFINITY;
  };
  const order = new Map<string, number>();
  for (const label of [...new Set(labels)].sort((a, b) => rank(a) - rank(b))) {
    order.set(label, order.size);
  }
  return order;
}

/**
 * A date field grouped by several units is, as Excel lays it out, several
 * axis fields: "Years" and "Quarters" derived from it (coarse to fine) and
 * then the field itself showing the finest unit. The derived levels become
 * virtual source columns appended to the matrix, so every reader of the
 * expanded config sees ordinary axis fields. Idempotent: a config with at
 * most one unit per field comes back unchanged.
 */
export function expandPivotGroupLevels(
  matrix: PivotMatrix,
  config: PivotConfig,
): { matrix: PivotMatrix; config: PivotConfig; levels: Map<number, PivotGroupLevel> } {
  const dateUnitsByCol = new Map<number, PivotGrouping[]>();
  for (const grouping of config.grouping ?? []) {
    if (grouping.kind !== "date") continue;
    dateUnitsByCol.set(grouping.colIndex, [
      ...(dateUnitsByCol.get(grouping.colIndex) ?? []),
      grouping,
    ]);
  }
  // Every expanded column that is a level of a date field, named as the
  // recipe names it: the field's source column and the level's unit.
  const levels = new Map<number, PivotGroupLevel>();
  for (const [colIndex, units] of dateUnitsByCol) {
    if (units.length === 1) levels.set(colIndex, { colIndex, unit: units[0].unit ?? "month" });
  }
  const multi = [...dateUnitsByCol.entries()].filter(([, units]) => units.length > 1);
  const discrete = (config.grouping ?? []).filter((item) => item.kind === "discrete");
  if ((multi.length === 0 && discrete.length === 0) || !Array.isArray(matrix) || matrix.length === 0) {
    return { matrix, config, levels };
  }
  const width = matrix[0]?.length ?? 0;
  const header = [...(matrix[0] ?? [])];
  const rows = matrix.slice(1).map((row) => [...row]);
  const expansion = new Map<number, number[]>();
  const grouping: PivotGrouping[] = (config.grouping ?? []).filter(
    (item) =>
      item.kind !== "discrete" &&
      (item.kind !== "date" || (dateUnitsByCol.get(item.colIndex)?.length ?? 0) < 2),
  );
  const taken = new Set(header.map((label) => String(label ?? "").toLowerCase()));
  let nextColumn = width;
  for (const [colIndex, units] of multi) {
    const ordered = [...units].sort(
      (a, b) => PIVOT_DATE_UNIT_RANK[a.unit ?? "month"] - PIVOT_DATE_UNIT_RANK[b.unit ?? "month"],
    );
    const virtual: number[] = [];
    for (const level of ordered.slice(0, -1)) {
      const base = PIVOT_DERIVED_LEVEL_NAMES[level.unit ?? "month"];
      let name = base;
      for (let ordinal = 2; taken.has(name.toLowerCase()); ordinal += 1) {
        name = `${base}${ordinal}`;
      }
      taken.add(name.toLowerCase());
      header[nextColumn] = name;
      for (const row of rows) {
        row[nextColumn] = pivotBucketLabel(
          row[colIndex] == null ? "" : String(row[colIndex]),
          level,
        );
      }
      grouping.push({ ...level, colIndex: nextColumn });
      levels.set(nextColumn, { colIndex, unit: level.unit ?? "month" });
      virtual.push(nextColumn);
      nextColumn += 1;
    }
    const finest = ordered[ordered.length - 1];
    grouping.push({ ...finest });
    levels.set(colIndex, { colIndex, unit: finest.unit ?? "month" });
    expansion.set(colIndex, virtual);
  }
  // A manual group is a derived column beside the untouched field; the axis
  // shows the groups, the groups with the items nested, or the items alone.
  const discreteExpansion = new Map<number, number[]>();
  for (const item of discrete) {
    const base = `${String(header[item.colIndex] ?? "")}2`;
    let name = base;
    for (let ordinal = 3; taken.has(name.toLowerCase()); ordinal += 1) {
      name = `${String(header[item.colIndex] ?? "")}${ordinal}`;
    }
    taken.add(name.toLowerCase());
    header[nextColumn] = name;
    for (const row of rows) {
      row[nextColumn] = pivotBucketLabel(
        row[item.colIndex] == null ? "" : String(row[item.colIndex]),
        item,
      );
    }
    // The derived column keeps the group order: its labels are already groups.
    grouping.push({ colIndex: nextColumn, kind: "discrete", groups: item.groups, display: "groups" });
    levels.set(nextColumn, { colIndex: item.colIndex, group: true });
    const display = item.display ?? "groups+items";
    discreteExpansion.set(item.colIndex, [
      ...(display !== "items" ? [nextColumn] : []),
      ...(display !== "groups" ? [item.colIndex] : []),
    ]);
    nextColumn += 1;
  }
  const expand = (fields: readonly number[]): number[] =>
    fields.flatMap((field) =>
      discreteExpansion.has(field)
        ? discreteExpansion.get(field)!
        : [...(expansion.get(field) ?? []), field],
    );
  // An item filter on a level of a date field moves to that level's column
  // (the derived column for a coarser unit, the field itself for the finest);
  // one naming a unit the field is no longer grouped by has nothing to filter.
  // A filter on a manual grouping's group field moves to that derived column.
  const itemFilters = config.itemFilters?.flatMap((entry) => {
    if (entry.unit === undefined && entry.group === undefined) return [entry];
    for (const [column, level] of levels) {
      if (
        level.colIndex === entry.colIndex &&
        level.unit === entry.unit &&
        level.group === entry.group
      ) {
        return [{ ...entry, colIndex: column }];
      }
    }
    return [];
  });
  return {
    matrix: [header, ...rows],
    config: {
      ...config,
      rows: expand(config.rows),
      columns: expand(config.columns),
      grouping,
      ...(itemFilters === undefined ? {} : { itemFilters }),
      ...(config.fieldFilters === undefined
        ? {}
        : { fieldFilters: remapPivotFieldFilterLevels(config.fieldFilters, levels) }),
    },
    levels,
  };
}

/** A level of a grouped field, as the recipe names it: a date field's level
 * by its unit, or the group field a manual grouping adds (`group`). */
export interface PivotGroupLevel {
  colIndex: number;
  unit?: PivotDateUnit;
  group?: true;
}

/**
 * What a drawn field caption's dropdown edits and lists. The report is laid
 * out from the expanded recipe, so the caption of a derived level ("Years")
 * names a virtual column with no source column behind it: reading the source
 * by that index finds nothing. The dropdown's field is the recipe's name for
 * the level (the date field's column and the level's unit), and its items are
 * the values the report groups by, in the order the axis shows them (years
 * ascending, months in calendar order, bins ascending), which is what Excel's
 * field dropdown lists for a grouped field.
 */
export function pivotFieldMenuItems(
  matrix: PivotMatrix,
  config: PivotConfig,
  field: number,
): { field: PivotGroupLevel; items: string[] } {
  const expanded = expandPivotGroupLevels(matrix, config);
  const grouping = (expanded.config.grouping ?? []).find((entry) => entry.colIndex === field);
  const items: string[] = [];
  const seen = new Set<string>();
  for (const row of expanded.matrix.slice(1)) {
    const raw = row[field] == null ? "" : String(row[field]);
    const item = grouping ? pivotBucketLabel(raw, grouping) : raw;
    if (seen.has(item)) continue;
    seen.add(item);
    items.push(item);
  }
  const order = pivotGroupItemOrder(grouping, items);
  if (order) items.sort((a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0));
  return { field: expanded.levels.get(field) ?? { colIndex: field }, items };
}

/**
 * A field's sort-and-filter list as its dropdown shows it: the items, the
 * ones shown, and the edits its check boxes and sort rows make. A report
 * filter (`kind: "page"`) keeps its selection in `filters`; a row or column
 * field keeps its hidden items in `itemFilters` (a date level by its unit).
 * Shared by the report's header dropdowns and a PivotChart's field buttons.
 */
export function pivotFieldMenuModel(
  matrix: PivotMatrix,
  config: PivotConfig,
  field: number,
  kind: "page" | "axis",
): {
  items: string[];
  selected: string[];
  sortOrder: PivotFieldSort["order"] | undefined;
  selectionPatch: (shown: readonly string[]) => Partial<PivotConfig>;
  sortPatch: (order: PivotFieldSort["order"] | undefined) => Partial<PivotConfig>;
  /** The field's sort by a value field, and the patch that sets or drops it. */
  valueSort: PivotFieldSort | undefined;
  valueSortPatch: (
    order: PivotFieldSort["order"] | undefined,
    valueIndex: number,
  ) => Partial<PivotConfig>;
  /** The field's Label (or Date) filter and its Value filter or Top 10. */
  labelFilter: PivotFieldFilter | undefined;
  valueFilter: PivotFieldFilter | undefined;
  dateField: boolean;
  fieldFilterPatch: (
    next: Omit<PivotFieldFilter, "colIndex" | "unit" | "group"> | null,
    family: "label" | "value",
  ) => Partial<PivotConfig>;
  clearPatch: () => Partial<PivotConfig>;
} {
  const menu = pivotFieldMenuItems(matrix, config, field);
  const ownsFilter = (entry: PivotFilterField) =>
    entry.colIndex === menu.field.colIndex &&
    entry.unit === menu.field.unit &&
    entry.group === menu.field.group;
  const ownsFieldFilter = (entry: PivotFieldFilter) =>
    entry.colIndex === menu.field.colIndex &&
    entry.unit === menu.field.unit &&
    entry.group === menu.field.group;
  const pageFilter = config.filters.find((entry) => entry.colIndex === field);
  const itemFilter = config.itemFilters?.find(ownsFilter);
  return {
    items: menu.items,
    selected: kind === "page" ? pageFilter?.selected ?? [] : itemFilter?.selected ?? [],
    sortOrder: config.fieldSorts?.find((entry) => entry.colIndex === field && entry.byValue === undefined)?.order,
    valueSort: config.fieldSorts?.find((entry) => entry.colIndex === field && entry.byValue !== undefined),
    valueSortPatch: (order, valueIndex) => {
      const others = (config.fieldSorts ?? []).filter((entry) => entry.colIndex !== field);
      return {
        fieldSorts: order ? [...others, { colIndex: field, order, byValue: valueIndex }] : others,
      };
    },
    selectionPatch: (shown) => {
      const all = shown.length === 0 || menu.items.every((item) => shown.includes(item));
      if (kind === "page") {
        return {
          filters: config.filters.map((entry) =>
            entry.colIndex === field ? { ...entry, selected: all ? [] : [...shown] } : entry,
          ),
        };
      }
      const others = (config.itemFilters ?? []).filter((entry) => !ownsFilter(entry));
      return { itemFilters: all ? others : [...others, { ...menu.field, selected: [...shown] }] };
    },
    sortPatch: (order) => {
      const others = (config.fieldSorts ?? []).filter((entry) => entry.colIndex !== field);
      return { fieldSorts: order ? [...others, { colIndex: field, order }] : others };
    },
    labelFilter: config.fieldFilters?.find(
      (entry) => ownsFieldFilter(entry) && !pivotFieldFilterNeedsValue(entry.type),
    ),
    valueFilter: config.fieldFilters?.find(
      (entry) => ownsFieldFilter(entry) && pivotFieldFilterNeedsValue(entry.type),
    ),
    // Excel offers Date Filters instead of Label Filters on a field of dates.
    dateField:
      menu.field.unit === undefined &&
      !menu.field.group &&
      menu.items.some((item) => item !== "") &&
      menu.items.every((item) => item === "" || /^\d{4}-\d{2}-\d{2}(T|$)/.test(item)),
    fieldFilterPatch: (next, family) => {
      // A field has one Label (or Date) filter and one Value filter (or Top
      // 10); applying one replaces the field's filter of the same family.
      const kept = (config.fieldFilters ?? []).filter(
        (entry) =>
          !ownsFieldFilter(entry) ||
          (family === "value") !== pivotFieldFilterNeedsValue(entry.type),
      );
      return {
        fieldFilters: next
          ? [...kept, { ...next, colIndex: menu.field.colIndex, ...(menu.field.unit ? { unit: menu.field.unit } : {}), ...(menu.field.group ? { group: true as const } : {}) }]
          : kept,
      };
    },
    clearPatch: () => ({
      // Excel's "Clear Filter From": the hidden items and both field filters.
      itemFilters: (config.itemFilters ?? []).filter((entry) => !ownsFilter(entry)),
      fieldFilters: (config.fieldFilters ?? []).filter((entry) => !ownsFieldFilter(entry)),
    }),
  };
}

/** Compact a finite number for a bin label (drop a trailing ".0"-style suffix). */
function trimNum(n: number): string {
  return Number.isInteger(n) ? String(n) : String(Number(n.toFixed(4)));
}

/**
 * The pivot recipe. `source` is the A1 range the matrix was read from (kept so
 * the component can re-read current cells on reload). `rows` / `columns` /
 * `values` are SOURCE column indices (0-based, relative to the source range's
 * first column). `filters` (additive) are report/page-filter fields that pare the
 * source rows before grouping. `calculatedFields` (additive, optional) are
 * user-defined formula fields a value field may reference via `calcRef`. Every
 * field is required in the in-memory shape (filters / calculatedFields default to
 * []); the persisted form is normalized through normalizePivotConfig so a
 * hand-edited / partial value still produces a safe config.
 */
export interface PivotConfig {
  source: string;
  /** Stable source worksheet identity when an authored Pivot lives elsewhere. */
  sourceSheetId?: string;
  /**
   * Set when the source is an Excel TABLE (Create PivotTable proposed "Table1"
   * rather than a range; OOXML `<worksheetSource name="Table1"/>`). The table
   * is named by its stable id on the source sheet, so a rename cannot detach
   * it. `source` keeps the table's range as of the last refresh - header and
   * data rows, never the total row - and is what every range-only reader, and
   * a PivotTable whose table was deleted, falls back to. See
   * `pivotLiveSourceRange`.
   */
  sourceTableId?: string;
  /**
   * The report is built on one column of that table (`Table1[Col]`): the
   * cache names the column, and `source` is the column's header and data
   * rows. Imported only.
   */
  sourceTableColumn?: string;
  /** Native pivotTableDefinition@name used by Excel Navigation. */
  name?: string;
  /**
   * Optional A1 anchor where the derived pivot grid is placed on this sheet.
   * The output grid itself is never persisted: it is recomputed from `source`.
   * Absent keeps the legacy side-panel-only pivot byte-for-byte unchanged.
   */
  output?: string;
  rows: number[];
  columns: number[];
  values: PivotValueField[];
  filters: PivotFilterField[];
  /**
   * Manual item filters (Excel's hidden PivotTable items, `item h="1"`): the
   * VISIBLE values of a field that is not a report filter. They pare source
   * rows exactly like `filters` but are not a Filters-area field — they sit on
   * row/column fields or on fields outside the layout, which is how a slicer
   * connected to the PivotTable filters it. Absent or empty = every item shows.
   */
  itemFilters?: PivotFilterField[];
  grouping: PivotGrouping[];
  calculatedFields: PivotCalculatedField[];
  /**
   * Show per-outer-group SUBTOTAL rows (Excel PivotTable subtotals). OPTIONAL +
   * additive — absent/false leaves the pivot byte-for-byte unchanged (flat leaf
   * rows, first-seen order). When on with ≥2 row fields, rows are sorted so each
   * outer group is contiguous and a subtotal row is emitted after each group.
   */
  rowSubtotals?: boolean;
  subtotalPosition?: PivotSubtotalPosition;
  grandTotals?: PivotGrandTotalMode;
  reportLayout?: PivotReportLayout;
  blankRows?: boolean;
  repeatItemLabels?: boolean;
  /**
   * Σ Values in the Rows area, as the innermost row field (Excel's
   * `dataOnRows`): each value field is a row under every innermost row item
   * instead of a column under every column item. Only meaningful with two or
   * more value fields; absent = Values in the Columns area (the default).
   */
  valuesOnRows?: true;
  /**
   * Excel's per-field item sort (PivotField `sortType`): the source column's
   * items run ascending or descending instead of in first-seen order. A field
   * outside the layout keeps its sort for when it is placed on an axis.
   */
  fieldSorts?: PivotFieldSort[];
  /** An unsorted field's manual item order (see PivotFieldItemOrder). */
  fieldItemOrders?: PivotFieldItemOrder[];
  /**
   * A field's own number format (PivotField `numFmtId`), how Excel reads its
   * items on an axis. Carried with the recipe and written back.
   */
  fieldNumberFormats?: PivotFieldNumberFormat[];
  /**
   * Excel for the web PivotTable Analyze -> Show -> +/- Buttons. Absent is the
   * native default (visible); false hides only the disclosure controls and
   * leaves the current expanded/collapsed result state untouched.
   */
  showDrillButtons?: boolean;
  /**
   * Items whose details are hidden (Excel's Collapse, the PivotField item's
   * `sd="0"`): an item's own flag, so the item shows collapsed under every
   * parent, in the report and in its PivotCharts alike. `colIndex` is the row
   * field as the report lays it out (a date field grouped by several units
   * lays its outer units out as fields of their own, `expandPivotGroupLevels`);
   * `items` are the raw items (a group's own label, "" for blank). Absent =
   * every item shows its details.
   */
  collapsed?: PivotCollapsedField[];
  styleName?: string;
  showRowHeaders?: boolean;
  showColumnHeaders?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
  /**
   * Typed cache authority for a newly-authored PivotTable. It exists before
   * XLSX export assigns package part/relationship identities and is replaced by
   * `native.snapshot` when that workbook is imported again.
   */
  cacheSnapshot?: PivotCacheSnapshot;
  /**
   * Excel's Scenario PivotTable report (Scenario Manager ▸ Summary ▸
   * Scenario PivotTable report; OOXML `cacheSource type="scenario"`): the
   * cache holds the scenarios' values from when the report was made and has
   * no worksheet range behind it. `source` only gives the cache's shape; the
   * grid is never read, the report is never "out of date", and Refresh is
   * refused — "Scenario reports don't automatically recalculate" (Microsoft
   * Support, "Switch between various sets of values by using scenarios").
   * Kept only beside an authored `cacheSnapshot`.
   */
  cacheSource?: "scenario";
  /** Original-package native Pivot/cache custody and explicit cache snapshot. */
  native?: PivotNativeAuthority;
  /** Timeline (x15) date filters, set only by a connected native timeline. */
  dateFilters?: PivotDateFilter[];
  /**
   * A row or column field's Label, Date and Value filters and its Top 10
   * (Excel's field filters, the definition's `filters`), applied in order
   * after the report filters and hidden items. See pivot-field-filters.
   */
  fieldFilters?: PivotFieldFilter[];
  /** A source field's own name in the report (Field Settings > Custom Name,
   * the PivotField's `name`), shown in its header, filter row and buttons. */
  fieldCaptions?: PivotFieldCaption[];
  /** An item typed over in the report (the item's `n`): `item` is its value,
   * `caption` what the report shows for it. */
  itemCaptions?: PivotItemCaption[];
  /**
   * A field's own subtotals (Field Settings > Subtotals): `functions` []
   * is None, a list is Custom (one subtotal per function, below the group).
   * A field without an entry follows the report's Subtotals setting.
   */
  fieldSubtotals?: PivotFieldSubtotal[];
  /** Calculated items (see PivotCalculatedItem). */
  calculatedItems?: PivotCalculatedItem[];
  /** PivotTable Options > For error values show (absent: the error itself). */
  errorCaption?: string;
  /** PivotTable Options > For empty cells show (absent: blank). */
  emptyCaption?: string;
}

export interface PivotFieldSubtotal {
  colIndex: number;
  functions: PivotAgg[];
}

/**
 * A calculated item (Excel's Fields, Items & Sets > Calculated Item): an item
 * `name` of a source field whose value is `formula` over the field's other
 * items, in Excel's own spelling (`Product[Apple]+Product[Banana]`, or the
 * bare `Apple+Banana` typed in its dialog). For each combination of the
 * report's other fields it is the formula over those items' sums, and the
 * subtotals and grand totals count it too, as Excel's do.
 */
export interface PivotCalculatedItem {
  colIndex: number;
  name: string;
  formula: string;
}

export function normalizePivotCalculatedItems(raw: unknown, cap: number): PivotCalculatedItem[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotCalculatedItem[] = [];
  const seen = new Set<string>();
  for (const entry of raw.slice(0, 64)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, name, formula } = entry as Record<string, unknown>;
    if (
      typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 || colIndex > cap ||
      typeof name !== "string" || !name.trim() || name.length > 255 ||
      typeof formula !== "string" || !formula.trim() || formula.length > 1_024 ||
      PIVOT_CONTROL_CHARS.test(name + formula)
    ) {
      continue;
    }
    const key = `${colIndex}\u0000${name.toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ colIndex, name, formula });
  }
  return out;
}

/** A calculated item's formula in the calculated-field grammar: each item
 * reference (`Field[Item]`, `Field['Item']`) becomes the quoted item. */
export function pivotCalculatedItemExpression(formula: string): string {
  return formula
    .replace(/^\s*=/, "")
    .replace(/(?:'(?:[^']|'')*'|[^\s+\-*/()'[\]]+)\s*\[\s*('(?:[^']|'')*'|[^\]]*)\s*\]/gu, (_match, item: string) => {
      const name = item.startsWith("'") ? item.slice(1, -1).replace(/''/g, "'") : item.trim();
      return `'${name.replace(/'/g, "''")}'`;
    });
}

/**
 * The records a report's calculated items add: for each combination of the
 * other axis fields' items, one record holding the item's value for each
 * value field's column. Summed with the field's own records, they give the
 * item's cells and its share of every subtotal and grand total.
 */
export function pivotCalculatedItemRecords(
  rows: readonly string[][],
  header: readonly string[],
  config: PivotConfig,
): string[][] {
  const items = (config.calculatedItems ?? []).filter(
    (item) => config.rows.includes(item.colIndex) || config.columns.includes(item.colIndex),
  );
  if (items.length === 0) return [];
  const groupingByCol = new Map<number, PivotGrouping>();
  for (const g of config.grouping ?? []) groupingByCol.set(g.colIndex, g);
  const itemOf = (row: readonly string[], col: number) => {
    const raw = row[col] == null ? "" : String(row[col]);
    const g = groupingByCol.get(col);
    return g ? pivotBucketLabel(raw, g) : raw;
  };
  const valueColumns = [
    ...new Set(config.values.filter((value) => !value.calcRef).map((value) => value.colIndex)),
  ];
  const width = header.length;
  const out: string[][] = [];
  for (const item of items) {
    const expression = pivotCalculatedItemExpression(item.formula);
    const context = [...config.rows, ...config.columns].filter((field) => field !== item.colIndex);
    const groups = new Map<string, { sample: readonly string[]; byItem: Map<string, string[][]> }>();
    for (const row of rows) {
      const key = context.map((field) => itemOf(row, field)).join("\u0000");
      let group = groups.get(key);
      if (!group) groups.set(key, (group = { sample: row, byItem: new Map() }));
      const own = itemOf(row, item.colIndex).toLowerCase();
      const list = group.byItem.get(own);
      if (list) list.push(row);
      else group.byItem.set(own, [row]);
    }
    for (const group of groups.values()) {
      const record = Array.from({ length: width }, () => "");
      for (const field of context) record[field] = itemOf(group.sample, field);
      record[item.colIndex] = item.name;
      for (const col of valueColumns) {
        const sums: Record<string, number> = {};
        for (const [name, list] of group.byItem) {
          sums[name] = aggSum(list.map((row) => (row[col] == null ? "" : String(row[col])))) ?? 0;
        }
        for (const name of calcFormulaFieldNames(expression)) {
          const key = name.trim().toLowerCase();
          if (!(key in sums)) sums[key] = 0;
        }
        const value = evaluatePivotCalcFormula(expression, sums);
        record[col] = value === null || !Number.isFinite(value) ? "" : String(value);
      }
      out.push(record);
    }
  }
  return out;
}

export function normalizePivotFieldSubtotals(raw: unknown, cap: number): PivotFieldSubtotal[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldSubtotal[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, 128)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, functions } = entry as Record<string, unknown>;
    if (
      typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 || colIndex > cap ||
      seen.has(colIndex) || !Array.isArray(functions)
    ) {
      continue;
    }
    const picked = PIVOT_AGGS.filter((agg) => functions.includes(agg));
    seen.add(colIndex);
    out.push({ colIndex, functions: picked });
  }
  return out;
}

/** "Sum", "Count", ... as a custom subtotal row names its function. */
export function pivotSubtotalFunctionLabel(agg: PivotAgg): string {
  return AGG_CAPTION_VERB[agg] ?? "Sum";
}

export interface PivotFieldCaption {
  colIndex: number;
  caption: string;
}

export interface PivotItemCaption {
  colIndex: number;
  item: string;
  caption: string;
}

const PIVOT_CONTROL_CHARS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/;

function normalizePivotCaptionText(value: unknown): string | undefined {
  return typeof value === "string" && value.length <= 255 && !PIVOT_CONTROL_CHARS.test(value)
    ? value
    : undefined;
}

export function normalizePivotFieldCaptions(raw: unknown, cap: number): PivotFieldCaption[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldCaption[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, 128)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, caption } = entry as Record<string, unknown>;
    const text = normalizePivotCaptionText(caption);
    if (
      typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 || colIndex > cap ||
      seen.has(colIndex) || !text || !text.trim()
    ) {
      continue;
    }
    seen.add(colIndex);
    out.push({ colIndex, caption: text });
  }
  return out;
}

export function normalizePivotItemCaptions(raw: unknown, cap: number): PivotItemCaption[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotItemCaption[] = [];
  const seen = new Set<string>();
  for (const entry of raw.slice(0, 10_000)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, item, caption } = entry as Record<string, unknown>;
    const text = normalizePivotCaptionText(caption);
    if (
      typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 || colIndex > cap ||
      typeof item !== "string" || item.length > 2_000 || !text
    ) {
      continue;
    }
    const key = `${colIndex}\u0000${item}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ colIndex, item, caption: text });
  }
  return out;
}

/** A field's name in the report: its own caption, else its source header. */
export function pivotFieldCaption(
  config: Pick<PivotConfig, "fieldCaptions">,
  matrix: PivotMatrix,
  colIndex: number,
): string {
  return (
    config.fieldCaptions?.find((entry) => entry.colIndex === colIndex)?.caption ??
    headerLabel(matrix, colIndex)
  );
}

/**
 * Restrict source data rows to those that pass EVERY report filter (a row passes
 * a filter when its `colIndex` value is in that filter's `selected` set, or the
 * set is empty = "(All)"). Pure + exported so the gating is unit-testable apart
 * from the full pivot compute. Returns the input rows verbatim when there are no
 * active filters, so an unfiltered pivot is byte-for-byte unchanged.
 *
 * A grouped field's items are its group items ("2024", "Jan", "0-99"), so a
 * filter on a column `grouping` covers matches the row's group item. A filter
 * without a `unit` also matches the raw value, so one written before the field
 * was grouped keeps hiding the rows it named.
 */
export function applyPivotFilters(
  dataRows: PivotMatrix,
  filters: readonly PivotFilterField[] | undefined,
  dateFilters?: readonly PivotDateFilter[],
  grouping?: readonly PivotGrouping[],
): PivotMatrix {
  // Timeline date filters: a separate whole-day pass before report filters.
  if (dateFilters && dateFilters.length > 0) {
    dataRows = applyPivotDateFilters(dataRows, dateFilters);
  }
  const active = (filters ?? []).filter((f) => f.selected.length > 0);
  if (active.length === 0) return dataRows;
  const groupingByCol = new Map<number, PivotGrouping>();
  for (const g of grouping ?? []) groupingByCol.set(g.colIndex, g);
  const allow = active.map((f) => ({
    col: f.colIndex,
    set: new Set(f.selected),
    // A native cache names date items as ISO date-times while the live
    // source cells hold serials: the same instant either way
    // (S26_pivottable_date_field_filter showed an empty report).
    instants: f.selected.every((item) => PIVOT_ISO_DAY.test(item))
      ? new Set(f.selected.map(pivotItemInstant).filter((item): item is string => item !== null))
      : null,
    grouping: groupingByCol.get(f.colIndex),
    rawToo: f.unit === undefined && f.group === undefined,
  }));
  return dataRows.filter((row) =>
    allow.every(({ col, set, instants, grouping: g, rawToo }) => {
      const raw = row[col] == null ? "" : String(row[col]);
      if (!g) {
        if (set.has(raw)) return true;
        const instant = instants ? pivotItemInstant(raw) : null;
        return instant !== null && instants!.has(instant);
      }
      return set.has(pivotBucketLabel(raw, g)) || (rawToo && set.has(raw));
    }),
  );
}

const PIVOT_ISO_DAY =
  /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d{1,7})?)?Z?$/;
const PIVOT_US_DAY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
const PIVOT_SERIAL_DAY = /^\d{1,7}(?:\.\d+)?$/;
const MAX_PIVOT_DATE_FILTERS = 16;

function pivotDayKey(year: number, month: number, day: number): string | null {
  const time = Date.UTC(year, month - 1, day);
  const date = new Date(time);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day ||
    year < 1 ||
    year > 9_999
  ) {
    return null;
  }
  return date.toISOString().slice(0, 10);
}

/** The instant (`YYYY-MM-DDTHH:MM:SS`, whole seconds) an ISO date-time or a
 *  1900-system serial from 1900-03-01 names; null for anything else. */
function pivotItemInstant(value: string): string | null {
  const text = value.trim();
  const match = PIVOT_ISO_DAY.exec(text);
  if (match) {
    return `${match[1]}-${match[2]}-${match[3]}T${match[4] ?? "00"}:${match[5] ?? "00"}:${match[6] ?? "00"}`;
  }
  if (!PIVOT_SERIAL_DAY.test(text)) return null;
  const serial = Number(text);
  if (!Number.isFinite(serial) || serial < EXCEL_1900_MARCH_FIRST_SERIAL) return null;
  const ms = Math.round(excelSerialToUtcMs(serial) / 1000) * 1000;
  return new Date(ms).toISOString().slice(0, 19);
}

/**
 * The calendar day (`YYYY-MM-DD`) a pivot source value falls on: an ISO date
 * or dateTime (native cache `d` items), US `M/D/YYYY`, or an Excel serial
 * (1900 system, 1900-03-01 onwards). Mirrors timeline_slicer._scalar_day.
 */
export function pivotFilterDay(value: unknown): string | null {
  const text = String(value ?? "").trim();
  let match = PIVOT_ISO_DAY.exec(text);
  if (match) return pivotDayKey(+match[1], +match[2], +match[3]);
  match = PIVOT_US_DAY.exec(text);
  if (match) return pivotDayKey(+match[3], +match[1], +match[2]);
  if (PIVOT_SERIAL_DAY.test(text)) {
    // Serials before 1900-03-01 stay unmatched, exactly as the API's
    // timeline_slicer._scalar_day refuses them, so both sides filter alike.
    const serial = Math.floor(Number(text));
    if (serial < EXCEL_1900_MARCH_FIRST_SERIAL) return null;
    return new Date(excelSerialToUtcMs(serial)).toISOString().slice(0, 10);
  }
  return null;
}

/**
 * The caption of a date item (a native cache's `d` items arrive as ISO
 * date-times): Excel shows it in the short date, with the time of day when
 * it is not midnight. Anything else is its own caption.
 */
/**
 * A field item's caption in a field's filter list: through the field's
 * number format when it has one (Excel's field settings; "KW "00 over a
 * week number lists KW 01), else as {@link pivotItemCaption} reads it.
 */
export function pivotFieldItemCaption(
  config: Pick<PivotConfig, "fieldNumberFormats"> | null | undefined,
  field: number,
  item: string,
): string {
  const format = config?.fieldNumberFormats?.find((entry) => entry.colIndex === field)?.numberFormat;
  if (format && item.trim() !== "") {
    const numeric = Number(item);
    if (Number.isFinite(numeric)) return applyNumberFormat(numeric, format);
  }
  return pivotItemCaption(item);
}

export function pivotItemCaption(item: string): string {
  const match = PIVOT_ISO_DAY.exec(item);
  if (!match || match[4] === undefined) return item;
  const time = Date.UTC(+match[1], +match[2] - 1, +match[3], +match[4], +match[5], +match[6]);
  if (Number.isNaN(time)) return item;
  // Captions follow Excel's 1900 calendar, so 1900-01-01 captions as 1/1/1900.
  const serial = excelSerialFromUtcMs(time);
  if (serial < 1) return item;
  const midnight = match[4] === "00" && match[5] === "00" && match[6] === "00";
  return applyNumberFormat(serial, midnight ? "m/d/yyyy" : "m/d/yyyy h:mm");
}

/** Keep rows whose date lies in every timeline filter's whole-day range. */
export function applyPivotDateFilters(
  dataRows: PivotMatrix,
  dateFilters: readonly PivotDateFilter[],
): PivotMatrix {
  if (dateFilters.length === 0) return dataRows;
  return dataRows.filter((row) =>
    dateFilters.every((filter) => {
      const day = pivotFilterDay(row[filter.colIndex]);
      return day !== null && day >= filter.start && day <= filter.end;
    }),
  );
}

function normalizePivotDateFilters(
  value: unknown,
  cap: number,
): PivotDateFilter[] {
  const raw = Array.isArray(value) ? value : [];
  const out: PivotDateFilter[] = [];
  const columns = new Set<number>();
  const timelines = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_PIVOT_DATE_FILTERS) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const colIndex = clampIndex(e.colIndex, cap);
    const start = typeof e.start === "string" ? pivotFilterDay(e.start) : null;
    const end = typeof e.end === "string" ? pivotFilterDay(e.end) : null;
    const timeline =
      typeof e.timeline === "string" &&
      e.timeline.trim() &&
      e.timeline.length <= 255 &&
      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(e.timeline)
        ? e.timeline
        : null;
    if (
      colIndex === null ||
      !start ||
      start !== e.start ||
      !end ||
      end !== e.end ||
      start > end ||
      !timeline ||
      columns.has(colIndex) ||
      timelines.has(timeline.toLowerCase())
    ) {
      continue;
    }
    columns.add(colIndex);
    timelines.add(timeline.toLowerCase());
    const filterId =
      Number.isInteger(e.filterId) &&
      (e.filterId as number) >= 0 &&
      (e.filterId as number) <= 2_147_483_647
        ? (e.filterId as number)
        : undefined;
    const evalOrder =
      Number.isInteger(e.evalOrder) &&
      (e.evalOrder as number) >= -1 &&
      (e.evalOrder as number) <= 2_147_483_647
        ? (e.evalOrder as number)
        : undefined;
    out.push({
      colIndex,
      start,
      end,
      timeline,
      ...(filterId !== undefined ? { filterId } : {}),
      ...(evalOrder !== undefined ? { evalOrder } : {}),
    });
  }
  return out;
}

/**
 * The computed pivot grid.
 * - `rowHeaders`: one entry per output row; each is the tuple of row-field
 *   display values for that group (e.g. ["East", "2024"]). Empty array entry =
 *   the grand-total row.
 * - `colHeaders`: one entry per output column; each is the tuple of
 *   column-field display values PLUS the value-field label for that column
 *   (e.g. ["Q1", "Sum of Sales"]). When there are no column fields, there is
 *   one column per value field labelled just by the value field.
 * - `cells[r][c]`: the aggregate (or null when no source row matched).
 * - `totals`: grand totals — `row[c]` is the column-c total across all groups,
 *   `column[r]` is the row-r total across all output columns, `grand[v]` is the
 *   overall total per value field, indexed by value-field position.
 */
export interface PivotResult {
  rowHeaders: string[][];
  colHeaders: PivotColumnHeader[];
  cells: (number | null)[][];
  totals: {
    row: (number | null)[];
    column: (number | null)[];
    grand: (number | null)[];
  };
  /**
   * Per-OUTER-GROUP subtotal rows (Excel PivotTable subtotals), present ONLY when
   * `rowSubtotals` is on and there are ≥2 row fields. Each is rendered AFTER the
   * leaf row at `afterRowIndex` (rows are sorted when subtotals are on, so each
   * outer group is contiguous). `outerValue` is the group's first-field value (the
   * render labels it "<value> Total"); `cells` aggregate the group across the same
   * output columns as `cells[r]`. Absent for an ordinary (no-subtotal) pivot, so
   * existing consumers + GETPIVOTDATA are byte-for-byte unaffected.
   */
  subtotals?: PivotSubtotalRow[];
  /**
   * The column-field tuples in output order, present whenever the recipe has a
   * column field. `colHeaders` repeats them per value field, so a layout
   * without values (a column label list) has its items only here.
   */
  columnTuples?: string[][];
}

/** One outer-group subtotal row inserted into the rendered pivot. */
export interface PivotSubtotalRow {
  /** Insert this subtotal row AFTER this (0-based) leaf row index. */
  afterRowIndex: number;
  /** The outer group's first-row-field value (rendered as "<value> Total"). */
  outerValue: string;
  /** Aggregate per output column, parallel to a `cells` row. */
  cells: (number | null)[];
}

/** A single output column: the column-field value tuple + which value field. */
export interface PivotColumnHeader {
  /** Display values of the column fields, in config order ([] when none). */
  columnValues: string[];
  /** Index into config.values that this output column aggregates. */
  valueIndex: number;
}

/** A 2D matrix of source display strings; matrix[0] is the header row. */
export type PivotMatrix = string[][];

// A source cell that is TEXT reading as a number ('20, a Text cell) is not a
// number to Sum, Average, Min or Max, as it is not in Excel; Count still
// counts it. The matrix keeps the display string for grouping, and the cells
// that are text are remembered beside it, by the matrix's identity.
const PIVOT_TEXT_CELLS = new WeakMap<PivotMatrix, ReadonlySet<string>>();
const PIVOT_TEXT_MARK = "\u0000";

/** Remember which cells of `matrix` are text that reads as a number, from the
 * typed scalars beside it; returns `matrix` for chaining. */
export function markPivotTextCells(
  matrix: PivotMatrix,
  scalarMatrix: PivotNativeScalarMatrix,
): PivotMatrix {
  const cells = new Set<string>();
  scalarMatrix.forEach((row, rowIndex) => {
    if (rowIndex === 0) return;
    row.forEach((scalar, colIndex) => {
      if (scalar.kind === "s" && coerceNumber(scalar.raw) !== null) {
        cells.add(`${rowIndex}:${colIndex}`);
      }
    });
  });
  if (cells.size > 0) PIVOT_TEXT_CELLS.set(matrix, cells);
  else PIVOT_TEXT_CELLS.delete(matrix);
  return matrix;
}

// ---------------------------------------------------------------------------
// Numeric coercion (shared by every aggregator).
// ---------------------------------------------------------------------------

/**
 * Coerce a raw cell string to a finite number, or null when it is blank /
 * non-numeric. Tolerant of what a user actually types: surrounding whitespace,
 * a leading currency symbol ($ € £ ¥), thousands separators (commas), a
 * trailing percent (scaled to a fraction), and parenthesized negatives
 * "(1,200)" = -1200. A truly empty or word value yields null so it is IGNORED
 * by aggregation (never counted as 0).
 */
export function coerceNumber(raw: unknown): number | null {
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  if (typeof raw !== "string") return null;
  // A pooled value marked as text is not a number, whatever it reads as.
  if (raw.startsWith(PIVOT_TEXT_MARK)) return null;
  let s = raw.trim();
  if (s === "") return null;

  // Parenthesized negative: (1,234.5) -> -1234.5
  let sign = 1;
  if (s.startsWith("(") && s.endsWith(")")) {
    sign = -1;
    s = s.slice(1, -1).trim();
  }

  // Trailing percent -> divide by 100.
  let percent = false;
  if (s.endsWith("%")) {
    percent = true;
    s = s.slice(0, -1).trim();
  }

  // Strip a single leading currency symbol.
  s = s.replace(/^[$€£¥]\s*/u, "");

  // A leading +/- sign (after the currency strip).
  if (s.startsWith("+")) {
    s = s.slice(1);
  } else if (s.startsWith("-")) {
    sign *= -1;
    s = s.slice(1);
  }

  // Remove thousands separators (commas only — keep the decimal point).
  s = s.replace(/,/gu, "");

  if (s === "") return null;
  // Reject anything that still has stray letters/symbols so "12abc" -> null.
  // A decimal exponent is part of a number, not a stray letter: a finite value
  // read from a cell is serialized with String(n), which writes 1e-7 and 1e+21
  // in scientific notation, and Excel reads "1E-7" as a number too. A bare or
  // dangling exponent ("1e", "e5", "1e+") still has no digits after it and is
  // rejected, as are JS-only forms (hex, "Infinity").
  if (!/^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/u.test(s)) return null;

  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return sign * (percent ? n / 100 : n);
}

// ---------------------------------------------------------------------------
// Aggregators. Each takes the RAW cell strings of the matching source rows for
// one value column; blanks / non-numeric are ignored. `count` counts NON-BLANK
// entries (Excel COUNTA semantics for a pivot count). Returns null when the
// group has no usable value (so the cell renders empty, not 0) — except COUNT,
// which returns 0 for "no matching non-blank rows".
// ---------------------------------------------------------------------------

export function aggSum(values: string[]): number | null {
  let total = 0;
  let any = false;
  for (const v of values) {
    const n = coerceNumber(v);
    if (n !== null) {
      total += n;
      any = true;
    }
  }
  return any ? total : null;
}

export function aggCount(values: string[]): number {
  let count = 0;
  for (const v of values) {
    // Count is COUNTA: a stored text of spaces is a non-empty cell, only a
    // truly empty cell is not counted (live audit B288).
    if (typeof v === "string" ? v !== "" : v != null) count += 1;
  }
  return count;
}

export function aggAverage(values: string[]): number | null {
  let total = 0;
  let count = 0;
  for (const v of values) {
    const n = coerceNumber(v);
    if (n !== null) {
      total += n;
      count += 1;
    }
  }
  return count > 0 ? total / count : null;
}

export function aggMin(values: string[]): number | null {
  let min: number | null = null;
  for (const v of values) {
    const n = coerceNumber(v);
    if (n !== null && (min === null || n < min)) min = n;
  }
  return min;
}

export function aggMax(values: string[]): number | null {
  let max: number | null = null;
  for (const v of values) {
    const n = coerceNumber(v);
    if (n !== null && (max === null || n > max)) max = n;
  }
  return max;
}

function numbersOf(values: string[]): number[] {
  const out: number[] = [];
  for (const v of values) {
    const n = coerceNumber(v);
    if (n !== null) out.push(n);
  }
  return out;
}

/** Product: the PRODUCT of the numeric values (XLS-A071). */
export function aggProduct(values: string[]): number | null {
  const numbers = numbersOf(values);
  if (numbers.length === 0) return null;
  const product = numbers.reduce((acc, n) => acc * n, 1);
  // A product past the float range has no value a cell can hold.
  return Number.isFinite(product) ? product : null;
}

/** Count Numbers: Excel's COUNT, the numeric values only (XLS-A071). */
export function aggCountNums(values: string[]): number {
  return numbersOf(values).length;
}

/**
 * What an aggregation yields where Excel shows #DIV/0!: a sample StdDev or
 * Var over fewer than two numbers (XLS-A071). It stays NaN inside the numeric
 * result grid; the report and the panel render it as the error.
 */
export const PIVOT_AGG_DIVIDE_BY_ZERO = Number.NaN;

/** True for an aggregate Excel shows as #DIV/0! (see PIVOT_AGG_DIVIDE_BY_ZERO). */
export function isPivotAggDivideByZero(value: number | null | undefined): boolean {
  return typeof value === "number" && Number.isNaN(value);
}

function pivotVariance(values: string[], sample: boolean): number | null {
  const numbers = numbersOf(values);
  if (numbers.length === 0) return null;
  // VAR / STDEV divide by n-1, so a lone number is Excel's #DIV/0!.
  if (sample && numbers.length < 2) return PIVOT_AGG_DIVIDE_BY_ZERO;
  const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
  const squares = numbers.reduce((sum, n) => sum + (n - mean) ** 2, 0);
  return squares / (sample ? numbers.length - 1 : numbers.length);
}

/** Var: the sample variance (n-1), Excel's VAR. */
export function aggVar(values: string[]): number | null {
  return pivotVariance(values, true);
}

/** Varp: the population variance (n), Excel's VARP. */
export function aggVarp(values: string[]): number | null {
  return pivotVariance(values, false);
}

/** StdDev: the sample standard deviation, Excel's STDEV. */
export function aggStdDev(values: string[]): number | null {
  const v = pivotVariance(values, true);
  return v === null ? null : Math.sqrt(v);
}

/** StdDevp: the population standard deviation, Excel's STDEVP. */
export function aggStdDevp(values: string[]): number | null {
  const v = pivotVariance(values, false);
  return v === null ? null : Math.sqrt(v);
}

const AGGREGATORS: Record<PivotAgg, (values: string[]) => number | null> = {
  sum: aggSum,
  count: aggCount,
  average: aggAverage,
  max: aggMax,
  min: aggMin,
  product: aggProduct,
  countNums: aggCountNums,
  stdDev: aggStdDev,
  stdDevp: aggStdDevp,
  var: aggVar,
  varp: aggVarp,
};

/** Apply a named aggregation to a group's raw value strings. */
export function aggregate(agg: PivotAgg, values: string[]): number | null {
  return (AGGREGATORS[agg] ?? aggSum)(values);
}

/**
 * The words of Excel's English default value-field name. Count Numbers reads
 * "Count of", as Excel names it.
 */
const AGG_CAPTION_VERB: Record<PivotAgg, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNums: "Count",
  stdDev: "StdDev",
  stdDevp: "StdDevp",
  var: "Var",
  varp: "Varp",
};

/** Human label for an aggregation, e.g. "Sum of Sales". Pure (no i18n here). */
export function aggLabel(agg: PivotAgg, fieldName: string): string {
  return `${AGG_CAPTION_VERB[agg] ?? "Sum"} of ${fieldName}`;
}

/**
 * The display label for a value field: Excel sums a CALCULATED field and names
 * it after the field ("Sum of Margin", as the dataField of a workbook Excel
 * wrote spells it); a SOURCE field shows its own aggregation label ("Sum of
 * Sales"). `sourceName` is the header label of the field's source column (only
 * used for a source field). Pure — the UI passes the resolved column header in.
 */
export function valueFieldLabel(field: PivotValueField, sourceName: string): string {
  if (field.caption && field.caption.trim() !== "") return field.caption.trim();
  if (field.calcRef && field.calcRef.trim() !== "") {
    return `Sum of ${field.calcRef.trim()}`;
  }
  return aggLabel(field.agg, sourceName);
}

// ---------------------------------------------------------------------------
// Calculated fields (Excel "Calculated Field").
//
// A calculated field is a user formula that combines OTHER source fields by their
// header label, e.g. `Sales - Cost` or `Profit / Sales`. Excel computes a
// calculated field over the SUMMED components of each group — for a group g the
// field's value is the formula evaluated with each referenced field replaced by
// SUM(field) across g (so `Profit % = Profit / Sales` for a region = that
// region's summed Profit / summed Sales, NOT the per-row ratios averaged). This
// module replicates that documented semantic: the engine sums each referenced
// source field per group and feeds those sums to evaluatePivotCalcFormula.
//
// The expression grammar is a tiny, dependency-free arithmetic over + - * / and
// parentheses (left-associative, * / before + -, unary +/-), with operands that
// are numeric literals or FIELD NAMES (bare identifiers, or any text wrapped in
// single quotes for names with spaces/operators, Excel-style). A field name that
// is unknown or a malformed expression yields null (the cell renders blank)
// rather than throwing — a pivot must never crash on a hand-typed formula; a
// division by zero is Excel's #DIV/0! (PIVOT_AGG_DIVIDE_BY_ZERO). A name may
// be another calculated field, evaluated over the same group.
// ---------------------------------------------------------------------------

type CalcToken =
  | { kind: "num"; value: number }
  | { kind: "name"; value: string }
  | { kind: "op"; value: "+" | "-" | "*" | "/" }
  | { kind: "lparen" }
  | { kind: "rparen" };

/** Tokenize a calc-field formula; returns null on any unrecognized character. */
function tokenizeCalcFormula(formula: string): CalcToken[] | null {
  const tokens: CalcToken[] = [];
  const s = formula;
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      i += 1;
      continue;
    }
    if (ch === "+" || ch === "-" || ch === "*" || ch === "/") {
      tokens.push({ kind: "op", value: ch });
      i += 1;
      continue;
    }
    if (ch === "(") {
      tokens.push({ kind: "lparen" });
      i += 1;
      continue;
    }
    if (ch === ")") {
      tokens.push({ kind: "rparen" });
      i += 1;
      continue;
    }
    // Single-quoted field name: 'Gross Sales' (doubled '' = literal quote).
    if (ch === "'") {
      let j = i + 1;
      let name = "";
      let closed = false;
      while (j < s.length) {
        if (s[j] === "'") {
          if (s[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          closed = true;
          j += 1;
          break;
        }
        name += s[j];
        j += 1;
      }
      if (!closed) return null;
      tokens.push({ kind: "name", value: name.trim() });
      i = j;
      continue;
    }
    // Numeric literal (with optional decimal point).
    if ((ch >= "0" && ch <= "9") || ch === ".") {
      let j = i;
      let dot = false;
      while (j < s.length) {
        const c = s[j];
        if (c >= "0" && c <= "9") {
          j += 1;
        } else if (c === "." && !dot) {
          dot = true;
          j += 1;
        } else {
          break;
        }
      }
      const lit = s.slice(i, j);
      const n = Number(lit);
      if (!Number.isFinite(n)) return null;
      tokens.push({ kind: "num", value: n });
      i = j;
      continue;
    }
    // Bare field name: a letter/underscore/CJK start, then word-ish characters
    // (letters, digits, underscore, spaces, common name punctuation). We stop at
    // an operator/paren so `Sales-Cost` splits into Sales, -, Cost. A bare name
    // is right-trimmed so trailing spaces before an operator don't stick.
    if (/[^\s+\-*/()'0-9.]/u.test(ch)) {
      let j = i;
      while (j < s.length && /[^+\-*/()']/u.test(s[j])) {
        j += 1;
      }
      const name = s.slice(i, j).trim();
      if (name === "") return null;
      tokens.push({ kind: "name", value: name });
      i = j;
      continue;
    }
    return null; // unrecognized character
  }
  return tokens;
}

/**
 * Evaluate a calculated-field formula against the per-group SUMS of its
 * referenced source fields. `fieldSums` maps a source field's header label
 * (case/space-insensitively) to that field's SUM over the current group (a field
 * with no numeric data in the group contributes 0, matching Excel SUM). Returns
 * the numeric result, null when the formula is malformed or references an
 * unknown field (the pivot then shows a blank cell), or
 * PIVOT_AGG_DIVIDE_BY_ZERO when it divides by zero (Excel's #DIV/0!).
 */
export function evaluatePivotCalcFormula(
  formula: string,
  fieldSums: Record<string, number>,
): number | null {
  if (typeof formula !== "string" || formula.trim() === "") return null;
  const tokens = tokenizeCalcFormula(formula);
  if (!tokens || tokens.length === 0) return null;

  // Case/space-insensitive field lookup (mirrors looseEqual elsewhere).
  const lookup = new Map<string, number>();
  for (const [k, v] of Object.entries(fieldSums)) {
    lookup.set(k.trim().toLowerCase(), v);
  }

  let pos = 0;
  let failed = false;
  const peek = (): CalcToken | undefined => tokens[pos];

  // Recursive-descent: expr := term (('+'|'-') term)*; term := factor
  // (('*'|'/') factor)*; factor := number | name | '(' expr ')' | ('+'|'-')
  // factor. A division by zero or unknown name sets `failed` -> overall null.
  const parseExpr = (): number => {
    let value = parseTerm();
    for (;;) {
      const tk = peek();
      if (tk && tk.kind === "op" && (tk.value === "+" || tk.value === "-")) {
        pos += 1;
        const rhs = parseTerm();
        value = tk.value === "+" ? value + rhs : value - rhs;
      } else {
        break;
      }
    }
    return value;
  };

  const parseTerm = (): number => {
    let value = parseFactor();
    for (;;) {
      const tk = peek();
      if (tk && tk.kind === "op" && (tk.value === "*" || tk.value === "/")) {
        pos += 1;
        const rhs = parseFactor();
        if (tk.value === "*") {
          value *= rhs;
        } else if (rhs === 0) {
          // Excel shows #DIV/0! (the grid's PIVOT_AGG_DIVIDE_BY_ZERO), and
          // the error carries through the rest of the formula and into any
          // calculated field that reads this one (LS034_calcfields, X-265).
          value = PIVOT_AGG_DIVIDE_BY_ZERO;
        } else {
          value /= rhs;
        }
      } else {
        break;
      }
    }
    return value;
  };

  const parseFactor = (): number => {
    const tk = peek();
    if (!tk) {
      failed = true;
      return 0;
    }
    if (tk.kind === "op" && (tk.value === "+" || tk.value === "-")) {
      pos += 1;
      const operand = parseFactor();
      return tk.value === "-" ? -operand : operand;
    }
    if (tk.kind === "num") {
      pos += 1;
      return tk.value;
    }
    if (tk.kind === "name") {
      pos += 1;
      const v = lookup.get(tk.value.trim().toLowerCase());
      if (v === undefined) {
        failed = true; // unknown field reference
        return 0;
      }
      return v;
    }
    if (tk.kind === "lparen") {
      pos += 1;
      const inner = parseExpr();
      const close = peek();
      if (!close || close.kind !== "rparen") {
        failed = true;
        return 0;
      }
      pos += 1;
      return inner;
    }
    failed = true; // a stray operator / rparen where a factor was expected
    return 0;
  };

  const result = parseExpr();
  // Leftover tokens (e.g. "1 2" or unbalanced ")") mean a malformed formula.
  if (failed || pos !== tokens.length) return null;
  if (Number.isNaN(result)) return PIVOT_AGG_DIVIDE_BY_ZERO;
  if (!Number.isFinite(result)) return null;
  return result;
}

/** How deep one calculated field may read another (a cycle stops here). */
const MAX_PIVOT_CALC_NESTING = 8;

/**
 * A calculated field's value over one group: each source field it names as
 * its SUM there (`sumOf`, by lower-cased name; undefined for no such field),
 * each other calculated field it names evaluated over the same group, as
 * Excel lets one calculated field read another (LS034_calcfields' Field5 =
 * 'Field4'/1, X-265). A name that is neither stays unknown (blank).
 */
function pivotCalcFieldValue(
  formula: string,
  sumOf: (nameKey: string) => number | undefined,
  calcByName: ReadonlyMap<string, PivotCalculatedField>,
  depth = 0,
): number | null {
  const sums: Record<string, number> = {};
  for (const name of calcFormulaFieldNames(formula)) {
    const key = name.trim().toLowerCase();
    const own = sumOf(key);
    if (own !== undefined) {
      sums[key] = own;
      continue;
    }
    const nested = calcByName.get(key);
    if (nested && depth < MAX_PIVOT_CALC_NESTING) {
      const value = pivotCalcFieldValue(nested.formula, sumOf, calcByName, depth + 1);
      if (value !== null) sums[key] = value;
    }
  }
  return evaluatePivotCalcFormula(formula, sums);
}

/** The source-field names a calculated field reads, through the calculated
 * fields it names (each read once). */
function pivotCalcSourceFieldNames(
  formula: string,
  calcByName: ReadonlyMap<string, PivotCalculatedField>,
): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  const visit = (text: string, depth: number) => {
    for (const name of calcFormulaFieldNames(text)) {
      const key = name.trim().toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(name);
      const nested = calcByName.get(key);
      if (nested && depth < MAX_PIVOT_CALC_NESTING) visit(nested.formula, depth + 1);
    }
  };
  visit(formula, 0);
  return out;
}

/**
 * The set of source-field header labels a calc formula references (its bare /
 * quoted name tokens), in first-seen order, deduped case-insensitively. Used to
 * decide which source columns the engine must sum per group for this field. A
 * malformed formula yields []. Numeric literals are not fields.
 */
export function calcFormulaFieldNames(formula: string): string[] {
  const tokens = tokenizeCalcFormula(formula);
  if (!tokens) return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const tk of tokens) {
    if (tk.kind === "name") {
      const key = tk.value.trim().toLowerCase();
      if (key !== "" && !seen.has(key)) {
        seen.add(key);
        out.push(tk.value.trim());
      }
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Distinct values (drives the row/column header axes).
// ---------------------------------------------------------------------------

/**
 * The distinct DATA values of a source column, in first-seen order (Excel pivot
 * default ordering — NOT sorted, so a user's natural data order is preserved).
 * matrix[0] is the header and is skipped. Blank cells contribute the empty
 * string "" exactly once (so a blank group still appears, like Excel's
 * "(blank)"). A column index out of range yields [].
 */
export function distinctValues(matrix: PivotMatrix, colIndex: number): string[] {
  if (!Array.isArray(matrix) || matrix.length <= 1) return [];
  if (!Number.isInteger(colIndex) || colIndex < 0) return [];
  // Out-of-range column (wider than any row, header included) -> []. A blank
  // cell INSIDE the data width still contributes its "" group below.
  const width = matrix.reduce(
    (max, row) => (Array.isArray(row) ? Math.max(max, row.length) : max),
    0,
  );
  if (colIndex >= width) return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (let r = 1; r < matrix.length; r += 1) {
    const row = matrix[r];
    const raw = Array.isArray(row) ? row[colIndex] : undefined;
    const v = raw == null ? "" : String(raw);
    if (!seen.has(v)) {
      seen.add(v);
      out.push(v);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Config normalization.
// ---------------------------------------------------------------------------

function clampIndex(value: unknown, max: number): number | null {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return null;
  const i = Math.trunc(n);
  if (i < 0 || i > max) return null;
  return i;
}

function normalizeAgg(value: unknown): PivotAgg {
  const a = typeof value === "string" ? value.trim().toLowerCase() : "";
  return PIVOT_AGGS.find((agg) => agg.toLowerCase() === a) ?? "sum";
}

// Bounds for calculated fields so a hand-edited file cannot build a huge list
// (DoS on iteration) or a pathological formula (DoS on parse).
const MAX_CALC_FIELDS = 32;
const MAX_CALC_NAME = 128;
export const MAX_PIVOT_VALUE_CAPTION = 255;
export const MAX_PIVOT_VALUE_NUMBER_FORMAT = 200;
const MAX_CALC_FORMULA = 1024;
const MAX_SHOWAS_ITEM = 256;
// Report/page-filter bounds: at most this many filter fields, each allowing at
// most this many selected values, each value clamped to this length — so a
// hand-edited / forged config can never carry an unbounded filter payload.
const MAX_PIVOT_FILTERS = 16;
/** One item filter per cache field at most; mirrors PIVOT_NATIVE_MAX_FIELDS. */
const MAX_PIVOT_ITEM_FILTERS = 128;
const MAX_PIVOT_GROUPING = 64;
const MAX_PIVOT_FILTER_SELECTED = 10_000;
const MAX_PIVOT_FILTER_VALUE = 500;

function normalizeShowValuesAs(value: unknown): PivotShowValuesAsConfig | undefined {
  if (!value || typeof value !== "object") return undefined;
  const o = value as Record<string, unknown>;
  const rawMode = typeof o.mode === "string" ? o.mode.trim() : "";
  const mode = (PIVOT_SHOW_VALUES_AS as readonly string[]).includes(rawMode)
    ? (rawMode as PivotShowValuesAs)
    : "normal";
  // "normal" is the default — drop it so a no-op transform never persists.
  if (mode === "normal") return undefined;
  const out: PivotShowValuesAsConfig = { mode };
  // A base names the field the mode runs along; the shares of a total have
  // none, and only the differenceFrom modes carry an item as well.
  if (showValuesAsNeedsBaseField(mode)) {
    const rawBase = o.base;
    if (rawBase && typeof rawBase === "object") {
      const b = rawBase as Record<string, unknown>;
      const fixed =
        mode === "runningTotalInRows"
          ? "row"
          : mode === "runningTotalInColumns"
            ? "col"
            : null;
      const axis = fixed ?? (b.axis === "col" ? "col" : b.axis === "row" ? "row" : null);
      const posN = clampIndex(b.pos, MAX_RANGE_COL);
      const item =
        typeof b.item === "string" ? b.item.slice(0, MAX_SHOWAS_ITEM) : "";
      if (axis !== null && posN !== null) {
        out.base = showValuesAsNeedsBaseItem(mode)
          ? { axis, pos: posN, item }
          : { axis, pos: posN };
      }
    }
  }
  return out;
}

function normalizeCalculatedFields(value: unknown): PivotCalculatedField[] {
  const arr = Array.isArray(value) ? value : [];
  const out: PivotCalculatedField[] = [];
  const seen = new Set<string>();
  for (const entry of arr) {
    if (out.length >= MAX_CALC_FIELDS) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const name = typeof e.name === "string" ? e.name.trim().slice(0, MAX_CALC_NAME) : "";
    const formula =
      typeof e.formula === "string" ? e.formula.slice(0, MAX_CALC_FORMULA) : "";
    if (name === "" || formula.trim() === "") continue; // both required
    const key = name.toLowerCase();
    if (seen.has(key)) continue; // a calc-field name must be unique
    seen.add(key);
    out.push({ name, formula });
  }
  return out;
}

/**
 * Normalize a free-form / partial pivot config into a safe, in-bounds shape.
 * - row/column field indices are coerced to in-grid integers; bad entries are
 *   dropped and duplicates removed (a column can't be its own group twice).
 * - value fields are coerced ({colIndex, agg}) plus optional additive
 *   `calcRef` (kept only when it names a defined calculated field) and `showAs`
 *   (a Show-Values-As transform; "normal"/garbage is dropped). A bad colIndex
 *   drops a SOURCE value field; an unknown agg falls back to "sum". A calc value
 *   field (valid calcRef) keeps its placeholder colIndex/agg, which the engine
 *   ignores.
 * - `calculatedFields` are coerced ({name, formula}, both required, names unique,
 *   list + lengths bounded).
 * - `maxCol` (when given, e.g. the source range width - 1) further clamps the
 *   indices so a stale config from a wider range can't address past the data.
 * - the same column may appear in BOTH rows and columns is disallowed (it is
 *   removed from `columns` to keep the axes disjoint, matching Excel).
 */
export function normalizePivotConfig(
  input: unknown,
  maxCol = MAX_RANGE_COL,
): PivotConfig {
  const o =
    input && typeof input === "object" ? (input as Record<string, unknown>) : {};
  const native = normalizePivotNativeAuthority(o.native);
  // An imported report's fields are its cache's, whatever its source range
  // reads now: a source column deleted since the last Refresh still has its
  // cached field until Refresh drops it.
  const cap = Math.min(
    MAX_RANGE_COL,
    Math.max(0, Math.trunc(maxCol), native ? native.fields.length - 1 : 0),
  );

  const dedupeIndices = (value: unknown): number[] => {
    const arr = Array.isArray(value) ? value : [];
    const seen = new Set<number>();
    const out: number[] = [];
    for (const entry of arr) {
      const i = clampIndex(entry, cap);
      if (i !== null && !seen.has(i)) {
        seen.add(i);
        out.push(i);
      }
    }
    return out;
  };

  const rows = dedupeIndices(o.rows);
  const rowSet = new Set(rows);
  // Columns are kept disjoint from rows (a field can't be both an axis).
  const columns = dedupeIndices(o.columns).filter((i) => !rowSet.has(i));

  const calculatedFields = normalizeCalculatedFields(o.calculatedFields);
  const calcNames = new Set(calculatedFields.map((f) => f.name.toLowerCase()));

  const values: PivotValueField[] = [];
  const rawValues = Array.isArray(o.values) ? o.values : [];
  for (const entry of rawValues) {
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const showAs = normalizeShowValuesAs(e.showAs);
    const caption =
      typeof e.caption === "string" &&
      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(e.caption)
        ? e.caption.trim().slice(0, MAX_PIVOT_VALUE_CAPTION)
        : "";
    const numberFormat =
      typeof e.numberFormat === "string" &&
      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(e.numberFormat)
        ? e.numberFormat.trim().slice(0, MAX_PIVOT_VALUE_NUMBER_FORMAT)
        : "";
    const rawCalcRef =
      typeof e.calcRef === "string" ? e.calcRef.trim().slice(0, MAX_CALC_NAME) : "";
    // A calc value field is kept only when its calcRef names a defined calc
    // field; otherwise it falls through to the source-column path.
    if (rawCalcRef !== "" && calcNames.has(rawCalcRef.toLowerCase())) {
      const field: PivotValueField = { colIndex: 0, agg: "sum", calcRef: rawCalcRef };
      if (caption) field.caption = caption;
      if (numberFormat && numberFormat !== "General") {
        field.numberFormat = numberFormat;
      }
      if (showAs) field.showAs = showAs;
      values.push(field);
      continue;
    }
    const colIndex = clampIndex(e.colIndex, cap);
    if (colIndex === null) continue;
    const field: PivotValueField = { colIndex, agg: normalizeAgg(e.agg) };
    if (caption) field.caption = caption;
    if (numberFormat && numberFormat !== "General") {
      field.numberFormat = numberFormat;
    }
    if (showAs) field.showAs = showAs;
    values.push(field);
  }

  // Report/page filters: a source column + its selected allowed values. Kept
  // DISJOINT from the row/column AXIS fields (a field is an axis OR a filter, the
  // same way columns are kept disjoint from rows). Deduped by column; the selected
  // value list is de-duplicated, length-clamped per value, and capped in count.
  const axisSet = new Set<number>([...rows, ...columns]);
  const filterSeen = new Set<number>();
  const filters: PivotFilterField[] = [];
  const rawFilters = Array.isArray(o.filters) ? o.filters : [];
  for (const entry of rawFilters) {
    if (filters.length >= MAX_PIVOT_FILTERS) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const colIndex = clampIndex(e.colIndex, cap);
    if (colIndex === null || axisSet.has(colIndex) || filterSeen.has(colIndex)) {
      continue;
    }
    filterSeen.add(colIndex);
    const rawSelected = Array.isArray(e.selected) ? e.selected : [];
    const selSeen = new Set<string>();
    const selected: string[] = [];
    for (const s of rawSelected) {
      if (selected.length >= MAX_PIVOT_FILTER_SELECTED) break;
      const v = (s == null ? "" : String(s)).slice(0, MAX_PIVOT_FILTER_VALUE);
      if (!selSeen.has(v)) {
        selSeen.add(v);
        selected.push(v);
      }
    }
    filters.push({ colIndex, selected });
  }

  // Field grouping (date period / number bin) applies to ROW/COLUMN AXIS fields
  // ONLY — grouping a non-axis column has no effect, so drop it. Deduped by
  // column; a number grouping needs a positive bin `size` (else dropped), a date
  // grouping defaults to "month" for an unknown unit.
  const groupSeen = new Set<string>();
  const grouping: PivotGrouping[] = [];
  const rawGrouping = Array.isArray(o.grouping) ? o.grouping : [];
  for (const entry of rawGrouping) {
    if (grouping.length >= MAX_PIVOT_GROUPING) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const colIndex = clampIndex(e.colIndex, cap);
    if (colIndex === null) continue;
    if (e.kind === "discrete") {
      // Excel's manual groups stay in the cache whether or not the field is
      // placed, so they are kept for a field off the axes too.
      if (groupSeen.has(`${colIndex}`) || groupSeen.has(`${colIndex}:date`)) continue;
      const rawGroups = Array.isArray(e.groups) ? e.groups : [];
      const usedNames = new Set<string>();
      const usedItems = new Set<string>();
      const groups: PivotDiscreteGroup[] = [];
      for (const rawGroup of rawGroups) {
        if (!rawGroup || typeof rawGroup !== "object" || groups.length >= 1_000) continue;
        const { name, items } = rawGroup as { name?: unknown; items?: unknown };
        if (typeof name !== "string" || name.length === 0 || name.length > 255) continue;
        if (usedNames.has(name)) continue;
        const members = (Array.isArray(items) ? items : [])
          .filter((item): item is string => typeof item === "string" && item.length <= 500)
          .filter((item) => !usedItems.has(item))
          .slice(0, 10_000);
        if (members.length === 0) continue;
        usedNames.add(name);
        for (const item of members) usedItems.add(item);
        groups.push({ name, items: members });
      }
      if (groups.length === 0) continue;
      const display =
        e.display === "groups" || e.display === "items" || e.display === "groups+items"
          ? e.display
          : undefined;
      groupSeen.add(`${colIndex}`);
      grouping.push({ colIndex, kind: "discrete", groups, ...(display ? { display } : {}) });
      continue;
    }
    if (!axisSet.has(colIndex)) continue;
    if (e.kind === "number") {
      // A number field groups once; a field grouped by dates cannot also bin.
      if (groupSeen.has(`${colIndex}`) || groupSeen.has(`${colIndex}:date`)) continue;
      const size =
        typeof e.size === "number" && Number.isFinite(e.size) && e.size > 0
          ? e.size
          : null;
      if (size === null) continue;
      const start =
        typeof e.start === "number" && Number.isFinite(e.start) ? e.start : 0;
      groupSeen.add(`${colIndex}`);
      grouping.push({ colIndex, kind: "number", size, start });
    } else if (e.kind === "date") {
      // A date field groups by any set of units (Excel's Years, Quarters,
      // Months on one field); each unit once.
      const unit = PIVOT_DATE_UNITS.includes(e.unit as PivotDateUnit)
        ? (e.unit as PivotDateUnit)
        : "month";
      if (groupSeen.has(`${colIndex}`) || groupSeen.has(`${colIndex}:${unit}`)) continue;
      groupSeen.add(`${colIndex}:date`);
      groupSeen.add(`${colIndex}:${unit}`);
      grouping.push({ colIndex, kind: "date", unit });
    }
  }

  // Manual item filters: any field that is not a report filter may hide items
  // (row/column fields and fields outside the layout alike). An entry with no
  // visible values means "no filter" and is dropped, so an unfiltered pivot
  // round-trips without an `itemFilters` key. A date field grouped by several
  // units is several axis fields (Years, Quarters, the field), each with its
  // own filter named by `unit`; a unit the field is not grouped by names a
  // level the report does not have, so that entry is dropped.
  const itemFilterSeen = new Set<string>([...filterSeen].map(String));
  const itemFilters: PivotFilterField[] = [];
  const rawItemFilters = Array.isArray(o.itemFilters) ? o.itemFilters : [];
  for (const entry of rawItemFilters) {
    if (itemFilters.length >= MAX_PIVOT_ITEM_FILTERS) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const colIndex = clampIndex(e.colIndex, cap);
    if (colIndex === null || filterSeen.has(colIndex)) continue;
    const unit =
      e.unit === undefined
        ? undefined
        : PIVOT_DATE_UNITS.includes(e.unit as PivotDateUnit)
          ? (e.unit as PivotDateUnit)
          : null;
    if (unit === null) continue;
    if (
      unit !== undefined &&
      !grouping.some((g) => g.colIndex === colIndex && g.kind === "date" && g.unit === unit)
    ) {
      continue;
    }
    // A manual grouping's group field: only while the field is grouped by hand.
    const group = e.group === true;
    if (e.group !== undefined && !group) continue;
    if (group && (unit !== undefined || !grouping.some((g) => g.colIndex === colIndex && g.kind === "discrete"))) {
      continue;
    }
    const seenKey = group ? `${colIndex}:group` : unit === undefined ? `${colIndex}` : `${colIndex}:${unit}`;
    if (itemFilterSeen.has(seenKey)) continue;
    itemFilterSeen.add(seenKey);
    const rawSelected = Array.isArray(e.selected) ? e.selected : [];
    const selSeen = new Set<string>();
    const selected: string[] = [];
    for (const s of rawSelected) {
      if (selected.length >= MAX_PIVOT_FILTER_SELECTED) break;
      const v = (s == null ? "" : String(s)).slice(0, MAX_PIVOT_FILTER_VALUE);
      if (!selSeen.has(v)) {
        selSeen.add(v);
        selected.push(v);
      }
    }
    if (selected.length > 0) {
      itemFilters.push({
        colIndex,
        selected,
        ...(unit === undefined ? {} : { unit }),
        ...(group ? { group: true as const } : {}),
      });
    }
  }

  const source = typeof o.source === "string" ? o.source.slice(0, 500) : "";
  const rawOutput = normalizePivotRange(o.output);
  const output = rawOutput
    ? pivotRangeToA1({
        r1: rawOutput.r1,
        c1: rawOutput.c1,
        r2: rawOutput.r1,
        c2: rawOutput.c1,
      })
    : "";

  const dateFilters = normalizePivotDateFilters(o.dateFilters, cap);
  const fieldFilters = normalizePivotFieldFilters(o.fieldFilters, cap).filter(
    (entry) => entry.valueIndex === undefined || entry.valueIndex < values.length,
  );
  // An imported report names its source sheet by the file's sheet name until
  // Change Data Source points it elsewhere; then, like any report, by id.
  const sourceSheetId =
    typeof o.sourceSheetId === "string" &&
    o.sourceSheetId.trim() &&
    o.sourceSheetId.length <= 255 &&
    !/[\u0000-\u001f\u007f-\u009f]/.test(o.sourceSheetId)
      ? o.sourceSheetId.trim()
      : undefined;
  const sourceTableId =
    typeof o.sourceTableId === "string" &&
    o.sourceTableId.trim() &&
    o.sourceTableId.length <= 255 &&
    !/[\u0000-\u001f\u007f-\u009f]/.test(o.sourceTableId)
      ? o.sourceTableId.trim()
      : undefined;
  const nameSource = typeof o.name === "string" ? o.name : native?.pivotName;
  const name =
    typeof nameSource === "string" && nameSource.trim()
      ? nameSource.trim().slice(0, 255)
      : undefined;
  const cacheSnapshot = native
    ? null
    : normalizePivotCacheSnapshot(o.cacheSnapshot, cap + 1);
  const subtotalPosition = PIVOT_SUBTOTAL_POSITIONS.includes(
    o.subtotalPosition as PivotSubtotalPosition,
  )
    ? (o.subtotalPosition as PivotSubtotalPosition)
    : o.rowSubtotals === true
      ? "bottom"
      : undefined;
  const grandTotals = PIVOT_GRAND_TOTAL_MODES.includes(
    o.grandTotals as PivotGrandTotalMode,
  )
    ? (o.grandTotals as PivotGrandTotalMode)
    : undefined;
  const reportLayout = PIVOT_REPORT_LAYOUTS.includes(
    o.reportLayout as PivotReportLayout,
  )
    ? (o.reportLayout as PivotReportLayout)
    : undefined;
  const styleName = isPivotStyleName(o.styleName)
    ? o.styleName
    : native?.styleName && isPivotStyleName(native.styleName)
      ? native.styleName
      : undefined;
  return {
    source,
    ...(sourceSheetId ? { sourceSheetId } : {}),
    ...(sourceTableId ? { sourceTableId } : {}),
    ...(sourceTableId &&
    typeof o.sourceTableColumn === "string" &&
    o.sourceTableColumn.trim() &&
    o.sourceTableColumn.length <= 255 &&
    !/[\[\]'\u0000-\u001f]/.test(o.sourceTableColumn)
      ? { sourceTableColumn: o.sourceTableColumn }
      : {}),
    ...(name ? { name } : {}),
    ...(output ? { output } : {}),
    rows,
    columns,
    values,
    filters,
    ...(itemFilters.length > 0 ? { itemFilters } : {}),
    grouping,
    calculatedFields,
    // Additive opt-in flag: kept ONLY when explicitly true, so an ordinary pivot
    // (the common case) round-trips byte-for-byte with no `rowSubtotals` key.
    ...(subtotalPosition && subtotalPosition !== "none"
      ? { rowSubtotals: true, subtotalPosition }
      : subtotalPosition === "none"
        ? { subtotalPosition }
        : {}),
    ...(grandTotals ? { grandTotals } : {}),
    ...(reportLayout ? { reportLayout } : {}),
    ...(o.blankRows === true ? { blankRows: true } : {}),
    ...(o.repeatItemLabels === true ? { repeatItemLabels: true } : {}),
    ...(o.valuesOnRows === true ? { valuesOnRows: true as const } : {}),
    ...(normalizePivotFieldSorts(o.fieldSorts).length > 0
      ? { fieldSorts: normalizePivotFieldSorts(o.fieldSorts) }
      : {}),
    ...(normalizePivotFieldItemOrders(o.fieldItemOrders).length > 0
      ? { fieldItemOrders: normalizePivotFieldItemOrders(o.fieldItemOrders) }
      : {}),
    ...(normalizePivotFieldNumberFormats(o.fieldNumberFormats).length > 0
      ? { fieldNumberFormats: normalizePivotFieldNumberFormats(o.fieldNumberFormats) }
      : {}),
    ...(o.showDrillButtons === false ? { showDrillButtons: false } : {}),
    ...(normalizePivotCollapsed(o.collapsed).length > 0
      ? { collapsed: normalizePivotCollapsed(o.collapsed) }
      : {}),
    ...(styleName ? { styleName } : {}),
    ...(o.showRowHeaders === false ? { showRowHeaders: false } : {}),
    ...(o.showColumnHeaders === false ? { showColumnHeaders: false } : {}),
    ...(o.bandedRows === true ? { bandedRows: true } : {}),
    ...(o.bandedColumns === true ? { bandedColumns: true } : {}),
    ...(cacheSnapshot ? { cacheSnapshot } : {}),
    // A scenario report is its frozen cache; without one it is no report.
    ...(cacheSnapshot && !native && o.cacheSource === "scenario"
      ? { cacheSource: "scenario" as const }
      : {}),
    ...(native ? { native } : {}),
    ...(dateFilters.length > 0 ? { dateFilters } : {}),
    ...(fieldFilters.length > 0 ? { fieldFilters } : {}),
    ...(normalizePivotCalculatedItems(o.calculatedItems, cap).length > 0
      ? { calculatedItems: normalizePivotCalculatedItems(o.calculatedItems, cap) }
      : {}),
    ...(normalizePivotFieldSubtotals(o.fieldSubtotals, cap).length > 0
      ? { fieldSubtotals: normalizePivotFieldSubtotals(o.fieldSubtotals, cap) }
      : {}),
    ...(normalizePivotFieldCaptions(o.fieldCaptions, cap).length > 0
      ? { fieldCaptions: normalizePivotFieldCaptions(o.fieldCaptions, cap) }
      : {}),
    ...(normalizePivotItemCaptions(o.itemCaptions, cap).length > 0
      ? { itemCaptions: normalizePivotItemCaptions(o.itemCaptions, cap) }
      : {}),
    ...(normalizePivotCaptionText(o.errorCaption) !== undefined
      ? { errorCaption: normalizePivotCaptionText(o.errorCaption) }
      : {}),
    ...(normalizePivotCaptionText(o.emptyCaption)
      ? { emptyCaption: normalizePivotCaptionText(o.emptyCaption) }
      : {}),
  };
}

/** One row field's collapsed items ({@link PivotConfig.collapsed}). */
export interface PivotCollapsedField {
  colIndex: number;
  items: string[];
}

const MAX_PIVOT_COLLAPSED_FIELDS = 16;
const MAX_PIVOT_COLLAPSED_ITEMS = 4096;

/** A bounded, de-duplicated `collapsed` list; [] when nothing is collapsed. */
export function normalizePivotCollapsed(raw: unknown): PivotCollapsedField[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotCollapsedField[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, MAX_PIVOT_COLLAPSED_FIELDS)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, items } = entry as { colIndex?: unknown; items?: unknown };
    if (
      typeof colIndex !== "number" ||
      !Number.isInteger(colIndex) ||
      colIndex < 0 ||
      colIndex > MAX_RANGE_COL + 64 ||
      seen.has(colIndex) ||
      !Array.isArray(items)
    ) {
      continue;
    }
    const kept = [
      ...new Set(
        items
          .slice(0, MAX_PIVOT_COLLAPSED_ITEMS)
          // An item is a source label, up to the length a source cell holds
          // (live audit B330: a 256-character group must collapse too).
          .filter((item): item is string => typeof item === "string" && item.length <= PIVOT_NATIVE_MAX_CELL_TEXT),
      ),
    ];
    if (kept.length === 0) continue;
    seen.add(colIndex);
    out.push({ colIndex, items: kept });
  }
  return out;
}

/**
 * The recipe with one item's details shown or hidden (Excel's Expand and
 * Collapse on an item, or on every item of the field when `items` is
 * omitted with `allItems`).
 */
export function pivotWithCollapsedItems(
  config: PivotConfig,
  colIndex: number,
  items: readonly string[],
  collapse: boolean,
): PivotConfig {
  const current = normalizePivotCollapsed(config.collapsed);
  const field = current.find((entry) => entry.colIndex === colIndex);
  const nextItems = new Set(field?.items ?? []);
  for (const item of items) {
    if (collapse) nextItems.add(item);
    else nextItems.delete(item);
  }
  const others = current.filter((entry) => entry.colIndex !== colIndex);
  const next =
    nextItems.size > 0 ? [...others, { colIndex, items: [...nextItems] }] : others;
  const { collapsed: _previous, ...rest } = config;
  return next.length > 0 ? { ...rest, collapsed: next } : rest;
}

/** True when a config can actually produce a grid (at least one value field). */
export function isPivotConfigComputable(config: PivotConfig): boolean {
  return config.values.length > 0;
}

// ---------------------------------------------------------------------------
// Create-from-selection: infer fields from a selected range + build a default
// recipe so "Insert PivotTable" produces a populated, non-empty starting pivot
// the user can refine in the existing builder panel. Pure (no React / no sheet
// access) so it unit-tests in isolation; the component only feeds it the source
// display strings it already reads for the engine.
// ---------------------------------------------------------------------------

/** One inferred source field: header name + 0-based column + sniffed type. */
export interface PivotInferredField {
  /** Header label (matrix[0][columnIndex]); "" when the header cell is blank. */
  name: string;
  /** 0-based column index, relative to the selection's first column. */
  columnIndex: number;
  /** Deterministic source role used by Recommended PivotTables. */
  type: "text" | "number" | "date";
  /**
   * The source column's own number format (its first non-blank data cell's),
   * when it has one other than General. A value field made from this column
   * starts with it, the way Excel gives a new data field the number format of
   * the field it summarizes.
   */
  numberFormat?: string;
}

/**
 * The value field Excel adds for a source field: SUM of a numeric field,
 * COUNT of any other, carrying the source field's number format when it has
 * one. Excel sets a new data field's format from the source field once, when
 * the field is added; Value Field Settings > Number Format then edits the data
 * field's own format and a later change to the source does not follow.
 */
export function newPivotValueField(
  colIndex: number,
  source: Pick<PivotInferredField, "type" | "numberFormat"> | undefined,
  agg: PivotAgg = source?.type === "number" ? "sum" : "count",
): PivotValueField {
  const field: PivotValueField = { colIndex, agg };
  const numberFormat = source?.numberFormat?.trim() ?? "";
  if (
    numberFormat !== "" &&
    numberFormat !== "General" &&
    numberFormat.length <= MAX_PIVOT_VALUE_NUMBER_FORMAT
  ) {
    field.numberFormat = numberFormat;
  }
  return field;
}

export interface PivotFieldProfile extends PivotInferredField {
  nonBlankCount: number;
  blankCount: number;
  uniqueCount: number;
}

/** The result of sniffing a selected range: its fields + how many data rows. */
export interface PivotFieldInference {
  fields: PivotInferredField[];
  /** Count of DATA rows (rows below the header); 0 for a header-only/empty range. */
  dataRowCount: number;
}

/**
 * Infer the available pivot fields from a selected range's display strings.
 * matrix[0] is treated as the HEADER row (field names), the rest as data. Each
 * column's type is sniffed over its data cells: a column counts as "number" when
 * it has at least one non-blank data cell AND every non-blank data cell coerces
 * to a number (coerceNumber) — so a "Sales" column of "$1,200" reads numeric but
 * a "Region" column of words reads text, and one stray label demotes a column to
 * text (safer for grouping than mis-summing). A header-only or empty selection
 * yields the header fields (all typed "text", since there is no data to sniff)
 * and dataRowCount 0. Field width follows the HEADER row's length so a ragged
 * data row can't invent phantom columns. `columnNumberFormats[c]`, when the
 * caller reads it off the source cells, is column c's own number format (see
 * PivotInferredField.numberFormat).
 */
export function inferPivotFieldsFromRange(
  values: string[][],
  columnNumberFormats?: readonly (string | undefined)[],
): PivotFieldInference {
  if (!Array.isArray(values) || values.length === 0) {
    return { fields: [], dataRowCount: 0 };
  }
  const header = Array.isArray(values[0]) ? values[0] : [];
  const width = header.length;
  if (width === 0) return { fields: [], dataRowCount: 0 };

  const dataRows = values.slice(1).filter((row) => Array.isArray(row));
  const dataRowCount = dataRows.length;

  const fields: PivotInferredField[] = [];
  for (let c = 0; c < width; c += 1) {
    const rawName = header[c];
    const name = rawName == null ? "" : String(rawName);
    let nonBlank = 0;
    let numeric = 0;
    for (const row of dataRows) {
      const raw = row[c];
      const s = raw == null ? "" : String(raw);
      if (s.trim() === "") continue;
      nonBlank += 1;
      if (coerceNumber(s) !== null) numeric += 1;
    }
    const type: "text" | "number" =
      nonBlank > 0 && numeric === nonBlank ? "number" : "text";
    const numberFormat = columnNumberFormats?.[c]?.trim() ?? "";
    fields.push(
      numberFormat !== "" && numberFormat !== "General"
        ? { name, columnIndex: c, type, numberFormat }
        : { name, columnIndex: c, type },
    );
  }

  return { fields, dataRowCount };
}

/**
 * Rich, deterministic profiling for Recommended PivotTables.  Date detection
 * intentionally requires an explicit date separator; a numeric year/ID column
 * therefore remains numeric instead of being guessed as a calendar field.
 */
export function profilePivotFieldsFromRange(
  values: string[][],
  scalarMatrix?: readonly (readonly PivotNativeScalar[])[],
  columnNumberFormats?: readonly (string | undefined)[],
): { fields: PivotFieldProfile[]; dataRowCount: number } {
  const inference = inferPivotFieldsFromRange(values, columnNumberFormats);
  const dataRows = values.slice(1).filter((row) => Array.isArray(row));
  const fields = inference.fields.map((field) => {
    const valuesForField = dataRows.map((row) =>
      row[field.columnIndex] == null ? "" : String(row[field.columnIndex]),
    );
    const nonBlankValues = valuesForField.filter(
      (value) => value.trim() !== "",
    );
    const explicitDateCount = nonBlankValues.filter(
      (value) => /[-/]/.test(value) && parseGroupableDate(value) !== null,
    ).length;
    const typedValuesForField = scalarMatrix
      ?.slice(1)
      .map((row) => row[field.columnIndex]);
    const typedDateValues = typedValuesForField?.filter(
      (value) => value && value.kind !== "m",
    );
    const hasExactTypedDateAuthority =
      typedDateValues !== undefined &&
      typedDateValues.length === nonBlankValues.length &&
      typedDateValues.length > 0 &&
      typedDateValues.every((value) => value.kind === "d");
    const type: PivotInferredField["type"] =
      hasExactTypedDateAuthority ||
      (nonBlankValues.length > 0 &&
        explicitDateCount === nonBlankValues.length)
        ? "date"
        : field.type;
    return {
      ...field,
      type,
      nonBlankCount: nonBlankValues.length,
      blankCount: valuesForField.length - nonBlankValues.length,
      uniqueCount: new Set(nonBlankValues).size,
    };
  });
  return { fields, dataRowCount: inference.dataRowCount };
}

/**
 * Build a sensible DEFAULT pivot recipe from inferred fields + the A1 source
 * range, so "Insert PivotTable" opens a non-empty, editable pivot rather than a
 * blank one. Defaults (Excel "recommended pivot" spirit, kept deterministic):
 *  - ROW field = the first TEXT column (a natural grouping dimension); when
 *    every column is numeric, fall back to the first column so there is still a
 *    row axis to group by.
 *  - VALUE field = the first NUMBER column, aggregated with SUM (the Excel
 *    default), as long as it is not the chosen row field; when no numeric column
 *    exists, COUNT the first column instead so the pivot still computes a grid.
 *  - no COLUMN fields by default (a flat row breakdown the user can cross-tab).
 * With no fields at all the recipe is source-only (rows/columns/values empty) —
 * still a valid, persistable in-progress pivot. The returned shape is the exact
 * PivotConfig the schema validates and normalizePivotConfig produces, so a
 * created pivot round-trips identically to an imported one.
 */
export function buildDefaultPivotConfig(
  fields: PivotInferredField[],
  sourceRangeRef: string,
): PivotConfig {
  const source = typeof sourceRangeRef === "string" ? sourceRangeRef : "";
  const list = Array.isArray(fields) ? fields : [];
  if (list.length === 0) {
    return { source, rows: [], columns: [], values: [], filters: [], grouping: [], calculatedFields: [] };
  }

  const firstText = list.find((f) => f.type === "text");
  const firstDate = list.find((f) => f.type === "date");
  const rowField = firstText ?? firstDate ?? list[0];

  const rows = [rowField.columnIndex];

  // Prefer SUM of the first numeric column that is NOT the row field (a value
  // should summarize a different column than the one it groups by).
  const sumField = list.find(
    (f) => f.type === "number" && f.columnIndex !== rowField.columnIndex,
  );
  let values: PivotValueField[];
  if (sumField) {
    values = [newPivotValueField(sumField.columnIndex, sumField, "sum")];
  } else {
    // No usable numeric value column -> COUNT a column other than the row field
    // (falling back to the row field only when it is the lone column) so the
    // pivot still produces a populated grid.
    const countField =
      list.find((f) => f.columnIndex !== rowField.columnIndex) ?? rowField;
    values = [newPivotValueField(countField.columnIndex, countField, "count")];
  }

  return { source, rows, columns: [], values, filters: [], grouping: [], calculatedFields: [] };
}

/**
 * Recommended layouts are computed eagerly for the gallery thumbnail and again
 * for the output plan. Keep the Cartesian result comfortably bounded before
 * either caller can allocate rowTuple × columnTuple pools.
 */
export const MAX_RECOMMENDED_PIVOT_RESULT_CELLS = 20_000;

function recommendationFitsResultBudget(
  matrix: PivotMatrix,
  config: PivotConfig,
  maxResultCells: number,
): boolean {
  if (!Array.isArray(matrix) || matrix.length <= 1) return false;
  if (!Number.isInteger(maxResultCells) || maxResultCells < 1) return false;
  const rowTuples = new Set<string>();
  const columnTuples = new Set<string>();
  if (config.rows.length === 0) rowTuples.add("[]");
  if (config.columns.length === 0) columnTuples.add("[]");
  const tupleFor = (row: string[], columns: number[]) =>
    JSON.stringify(
      columns.map((column) => {
        const value = row[column];
        return value == null ? "" : String(value);
      }),
    );

  for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
    const candidate = matrix[rowIndex];
    if (!Array.isArray(candidate)) continue;
    if (config.rows.length > 0) {
      rowTuples.add(tupleFor(candidate, config.rows));
    }
    if (config.columns.length > 0) {
      columnTuples.add(tupleFor(candidate, config.columns));
    }
    const outputColumns =
      Math.max(1, columnTuples.size) * Math.max(1, config.values.length);
    if (
      outputColumns > maxResultCells ||
      Math.max(1, rowTuples.size) >
        Math.floor(maxResultCells / outputColumns)
    ) {
      return false;
    }
  }
  return true;
}

/** Deterministic, bounded layouts for Insert ▸ Recommended PivotTables. */
export function buildRecommendedPivotConfigs(
  fields: PivotInferredField[],
  sourceRangeRef: string,
  matrix: PivotMatrix,
): PivotConfig[] {
  const source = typeof sourceRangeRef === "string" ? sourceRangeRef : "";
  const list = Array.isArray(fields)
    ? fields.filter(
        (field) =>
          field !== null &&
          typeof field === "object" &&
          Number.isInteger(field.columnIndex) &&
          field.columnIndex >= 0 &&
          (["text", "number", "date"] as const).includes(field.type),
      )
    : [];
  if (list.length === 0) return [];

  const recommendations: PivotConfig[] = [];
  const seen = new Set<string>();
  const add = (config: PivotConfig) => {
    if (!isPivotConfigComputable(config)) return;
    if (
      !recommendationFitsResultBudget(
        matrix,
        config,
        MAX_RECOMMENDED_PIVOT_RESULT_CELLS,
      )
    ) {
      return;
    }
    const key = JSON.stringify({
      rows: config.rows,
      columns: config.columns,
      values: config.values,
    });
    if (seen.has(key)) return;
    seen.add(key);
    recommendations.push(config);
  };

  const base = buildDefaultPivotConfig(list, source);
  add(base);
  const textFields = list.filter((field) => field.type === "text");
  const dateFields = list.filter((field) => field.type === "date");
  const numericFields = list.filter((field) => field.type === "number");
  const baseRow = base.rows[0];
  const baseValue = base.values[0];
  const firstDate = dateFields.find(
    (field) => field.columnIndex !== baseRow,
  );
  if (
    firstDate &&
    baseValue &&
    baseValue.colIndex !== firstDate.columnIndex
  ) {
    add({ ...base, columns: [firstDate.columnIndex] });
    add({ ...base, rows: [firstDate.columnIndex], columns: [] });
  }
  const secondText = textFields.find(
    (field) => field.columnIndex !== baseRow,
  );
  if (
    secondText &&
    baseValue &&
    baseValue.colIndex !== secondText.columnIndex
  ) {
    add({ ...base, columns: [secondText.columnIndex] });
    add({ ...base, rows: [secondText.columnIndex], columns: [] });
  }
  if (recommendations.length < 3 && baseValue) {
    const alternateValue = numericFields.find(
      (field) =>
        field.columnIndex !== baseRow &&
        field.columnIndex !== baseValue.colIndex,
    );
    if (alternateValue) {
      add({
        ...base,
        values: [newPivotValueField(alternateValue.columnIndex, alternateValue, "sum")],
      });
    }
  }
  if (baseValue && secondText && firstDate) {
    add({
      ...base,
      rows: [secondText.columnIndex, firstDate.columnIndex],
      columns: [],
    });
  }
  if (baseValue) {
    const alternateValue = numericFields.find(
      (field) =>
        field.columnIndex !== baseRow &&
        field.columnIndex !== baseValue.colIndex,
    );
    if (alternateValue) {
      add({
        ...base,
        values: [newPivotValueField(alternateValue.columnIndex, alternateValue, "sum")],
      });
    }
  }
  return recommendations.slice(0, 6);
}

// ---------------------------------------------------------------------------
// A1 source-range helpers (so the component can persist/read a string range).
// ---------------------------------------------------------------------------

/** A 0-based inclusive rectangle, normalized so r1<=r2 and c1<=c2. */
export interface PivotRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const A1_CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

function parseA1Cell(token: string): { row: number; col: number } | null {
  const m = A1_CELL_RE.exec(token.trim());
  if (!m) return null;
  const col = lettersToColumn(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || row < 0 || !Number.isFinite(row)) return null;
  return { row, col };
}

function clampRangeIndex(value: number, max: number): number {
  if (!Number.isFinite(value)) return 0;
  const n = Math.trunc(value);
  if (n < 0) return 0;
  if (n > max) return max;
  return n;
}

/**
 * Parse an A1 range ("A1:D20", a single "A1", or a {r1,c1,r2,c2} object) into a
 * normalized in-grid rectangle, or null when unrecognizable. Drops a leading
 * sheet qualifier ("Sheet1!A1:B2").
 */
export function normalizePivotRange(input: unknown): PivotRange | null {
  if (input && typeof input === "object" && !Array.isArray(input)) {
    const o = input as Record<string, unknown>;
    if (
      typeof o.r1 === "number" &&
      typeof o.c1 === "number" &&
      typeof o.r2 === "number" &&
      typeof o.c2 === "number"
    ) {
      const r1 = clampRangeIndex(o.r1, MAX_RANGE_ROW);
      const r2 = clampRangeIndex(o.r2, MAX_RANGE_ROW);
      const c1 = clampRangeIndex(o.c1, MAX_RANGE_COL);
      const c2 = clampRangeIndex(o.c2, MAX_RANGE_COL);
      return {
        r1: Math.min(r1, r2),
        c1: Math.min(c1, c2),
        r2: Math.max(r1, r2),
        c2: Math.max(c1, c2),
      };
    }
    return null;
  }

  if (typeof input !== "string") return null;
  const trimmed = input.trim();
  if (!trimmed) return null;
  const bang = trimmed.lastIndexOf("!");
  const body = bang >= 0 ? trimmed.slice(bang + 1) : trimmed;

  const parts = body.split(":");
  if (parts.length === 1) {
    const cell = parseA1Cell(parts[0]);
    if (!cell) return null;
    const row = clampRangeIndex(cell.row, MAX_RANGE_ROW);
    const col = clampRangeIndex(cell.col, MAX_RANGE_COL);
    return { r1: row, c1: col, r2: row, c2: col };
  }
  if (parts.length !== 2) return null;
  const a = parseA1Cell(parts[0]);
  const b = parseA1Cell(parts[1]);
  if (!a || !b) return null;
  const r1 = clampRangeIndex(a.row, MAX_RANGE_ROW);
  const r2 = clampRangeIndex(b.row, MAX_RANGE_ROW);
  const c1 = clampRangeIndex(a.col, MAX_RANGE_COL);
  const c2 = clampRangeIndex(b.col, MAX_RANGE_COL);
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/** Render a normalized PivotRange back to its A1 string ("A1:D20" or "A1"). */
export function pivotRangeToA1(range: PivotRange | null): string {
  if (!range) return "";
  const from = `${columnToLetters(range.c1)}${range.r1 + 1}`;
  if (range.r1 === range.r2 && range.c1 === range.c2) return from;
  const to = `${columnToLetters(range.c2)}${range.r2 + 1}`;
  return `${from}:${to}`;
}

// ---------------------------------------------------------------------------
// "Show Values As" transforms (Excel "Show Values As").
//
// Excel runs a transform ALONG ITS BASE FIELD: within one report, the cells a
// cell is measured against are the ones that differ from it only in that
// field's item, every other coordinate held fixed (including a total, which is
// a coordinate of its own). Three consequences, each measured against Excel on

//  - a cell carries a transform only when it names an item of the base field,
//    so a subtotal row carries one when the base field is OUTER of it, and the
//    grand total of that axis never does;
//  - the other axis takes part in full, its grand total included: a running
//    total down a row field fills the Grand Total column, one across a column
//    field fills the Grand Total row;
//  - a base field inside a group restarts at every group ("East" then "West"),
//    because the outer item is one of the coordinates held fixed.
// The percent-of-a-total modes have no base field and so transform every cell.
// A cell with no records counts as 0 where a running total or a difference
// reads it, and stays blank where a rank does.
// ---------------------------------------------------------------------------

/** The Excel error a "% Difference From" writes where it cannot subtract. */
export const PIVOT_SHOW_AS_NO_CELL = "#NULL!";
/** The Excel error a "% Difference From" writes for a zero base value. */
export const PIVOT_SHOW_AS_DIVIDE_BY_ZERO = "#DIV/0!";
/** What Excel writes through a value field whose base field left the layout. */
export const PIVOT_SHOW_AS_NO_FIELD = "#N/A";

/** A displayed value: the transformed number, an Excel error, or blank. */
export type PivotShownValue = number | string | null;

/**
 * One value field's aggregated cells, isolated from the full result grid so a
 * transform can reason about it alone. `cells[r][c]` is the aggregate at output
 * row r, column group c (within this value field). `rowTotals[r]` is the row's
 * total across the value field's column groups; `colTotals[c]` the column group's
 * total down all rows; `grand` the overall total. Totals are the engine's true
 * re-aggregates, so percentages divide by the correct denominator.
 */
export interface PivotValueGrid {
  cells: (number | null)[][];
  rowTotals: (number | null)[];
  colTotals: (number | null)[];
  grand: number | null;
}

/**
 * A cell as a share of a total: a cell with no records is 0 of it (Excel shows
 * 0, not a blank), a total of zero is the division Excel reports as #DIV/0!,
 * and a total that is not there at all leaves the cell blank.
 */
export function showAsShareOfTotal(
  cell: number | null,
  total: number | null,
): PivotShownValue {
  if (total === null) return null;
  if (total === 0) return PIVOT_SHOW_AS_DIVIDE_BY_ZERO;
  return (cell ?? 0) / total;
}

/**
 * Excel's Index: how much of this cell is its own, rather than what its row
 * and column would predict — (cell x grand total) / (row total x column
 * total). Every cell takes part, the totals included.
 */
export function showAsIndex(
  cell: number | null,
  rowTotal: number | null,
  columnTotal: number | null,
  grand: number | null,
): PivotShownValue {
  if (rowTotal === null || columnTotal === null || grand === null) return null;
  if (rowTotal === 0 || columnTotal === 0) return PIVOT_SHOW_AS_DIVIDE_BY_ZERO;
  return ((cell ?? 0) * grand) / (rowTotal * columnTotal);
}

/** True for the modes that read a base FIELD (and so run along one axis). */
export function showValuesAsNeedsBaseField(mode: PivotShowValuesAs): boolean {
  return (
    mode === "differenceFrom" ||
    mode === "percentDifferenceFrom" ||
    mode === "percentOfBaseItem" ||
    mode === "percentOfParentTotal" ||
    mode === "runningTotalInRows" ||
    mode === "runningTotalInColumns" ||
    mode === "percentOfRunningTotalIn" ||
    mode === "rankSmallestToLargest" ||
    mode === "rankLargestToSmallest"
  );
}

/** True for the modes that also name a base ITEM to measure against. */
export function showValuesAsNeedsBaseItem(mode: PivotShowValuesAs): boolean {
  return (
    mode === "differenceFrom" ||
    mode === "percentDifferenceFrom" ||
    mode === "percentOfBaseItem"
  );
}

/**
 * One cell of a transform that runs along a base field.
 *
 * `members` are the aggregates of the cells the base field spans for this cell
 * (its own included), in report order; `index` is this cell's place among them
 * and `baseIndex` the base item's place (-1 when the report has no such item).
 * Pure, and the whole of what the writer and the planner share — the server's
 * `_pivot_show_values_as` computes the same list for the same cell.
 */
export function showValuesAsAlongBase(
  mode: PivotShowValuesAs,
  members: readonly (number | null)[],
  index: number,
  baseIndex: number,
): PivotShownValue {
  if (index < 0 || index >= members.length) return null;
  const cell = members[index] ?? null;
  if (mode === "runningTotalInRows" || mode === "runningTotalInColumns") {
    // Excel carries the sum so far through a blank and shows 0 before the
    // first record, so a running total is never itself blank.
    let running = 0;
    for (let k = 0; k <= index; k += 1) running += members[k] ?? 0;
    return running;
  }
  if (mode === "rankSmallestToLargest" || mode === "rankLargestToSmallest") {
    if (cell === null) return null; // a cell with no records is not ranked
    const largestFirst = mode === "rankLargestToSmallest";
    // Competition rank (Excel RANK.EQ): ties share a rank, the next skips.
    let better = 0;
    for (const other of members) {
      if (other === null) continue;
      if (largestFirst ? other > cell : other < cell) better += 1;
    }
    return better + 1;
  }
  if (mode === "percentOfRunningTotalIn") {
    // The running total so far as a share of the whole run.
    let total = 0;
    let any = false;
    let running = 0;
    for (let k = 0; k < members.length; k += 1) {
      const member = members[k];
      if (member !== null) {
        total += member;
        any = true;
      }
      if (k <= index) running += member ?? 0;
    }
    if (!any) return null;
    if (total === 0) return PIVOT_SHOW_AS_DIVIDE_BY_ZERO;
    return running / total;
  }
  if (
    mode !== "differenceFrom" &&
    mode !== "percentDifferenceFrom" &&
    mode !== "percentOfBaseItem"
  ) {
    return cell;
  }
  if (baseIndex < 0 || baseIndex >= members.length) return null;
  if (index === baseIndex && mode !== "percentOfBaseItem") {
    return null; // its difference from itself; "% Of" reads 100% instead
  }
  const base = members[baseIndex] ?? null;
  if (mode === "differenceFrom") {
    // Either side without records counts as 0; neither side, and there is
    // nothing to measure.
    if (cell === null && base === null) return null;
    return (cell ?? 0) - (base ?? 0);
  }
  if (base === null) return null;
  if (cell === null) return PIVOT_SHOW_AS_NO_CELL;
  if (base === 0) return PIVOT_SHOW_AS_DIVIDE_BY_ZERO;
  return mode === "percentOfBaseItem" ? cell / base : (cell - base) / base;
}

/** % of the value field's GRAND total: cell / grand (blank when grand is 0). */
export function showAsPercentOfGrandTotal(
  grid: PivotValueGrid,
): PivotShownValue[][] {
  return grid.cells.map((row) =>
    row.map((cell) => showAsShareOfTotal(cell, grid.grand)),
  );
}

/** % of each cell's output ROW total (within this value field's columns). */
export function showAsPercentOfRowTotal(
  grid: PivotValueGrid,
): PivotShownValue[][] {
  return grid.cells.map((row, r) =>
    row.map((cell) => showAsShareOfTotal(cell, grid.rowTotals[r] ?? null)),
  );
}

/** % of each cell's output COLUMN total (down all rows of this value field). */
export function showAsPercentOfColumnTotal(
  grid: PivotValueGrid,
): PivotShownValue[][] {
  return grid.cells.map((row) =>
    row.map((cell, c) => showAsShareOfTotal(cell, grid.colTotals[c] ?? null)),
  );
}

/** Apply a transform down every column of a grid (its rows are the members). */
function alongRows(
  grid: PivotValueGrid,
  mode: PivotShowValuesAs,
  baseRowIndex: number,
): PivotShownValue[][] {
  const rows = grid.cells.length;
  const cols = rows > 0 ? grid.cells[0].length : 0;
  const out: PivotShownValue[][] = grid.cells.map((row) => row.map(() => null));
  for (let c = 0; c < cols; c += 1) {
    const members = grid.cells.map((row) => row[c] ?? null);
    for (let r = 0; r < rows; r += 1) {
      out[r][c] = showValuesAsAlongBase(mode, members, r, baseRowIndex);
    }
  }
  return out;
}

/** Apply a transform across every row of a grid (its columns are the members). */
function alongColumns(
  grid: PivotValueGrid,
  mode: PivotShowValuesAs,
  baseColIndex: number,
): PivotShownValue[][] {
  return grid.cells.map((row) =>
    row.map((_cell, c) =>
      showValuesAsAlongBase(
        mode,
        row.map((value) => value ?? null),
        c,
        baseColIndex,
      ),
    ),
  );
}

/**
 * Running total cumulating DOWN the rows of a grid, which is what the mode
 * means when the base field is the only field of the row axis.
 */
export function showAsRunningTotalInRows(
  grid: PivotValueGrid,
): PivotShownValue[][] {
  return alongRows(grid, "runningTotalInRows", -1);
}

/** Running total cumulating ACROSS the column groups of a grid, per row. */
export function showAsRunningTotalInColumns(
  grid: PivotValueGrid,
): PivotShownValue[][] {
  return alongColumns(grid, "runningTotalInColumns", -1);
}

/**
 * Competition rank of every body cell among the value field's cells in its
 * COLUMN (the grid form of a rank whose base field is the only row field).
 */
export function showAsRank(
  grid: PivotValueGrid,
  largestFirst: boolean,
): PivotShownValue[][] {
  return alongRows(
    grid,
    largestFirst ? "rankLargestToSmallest" : "rankSmallestToLargest",
    -1,
  );
}

/**
 * Difference (or % difference) of each cell FROM the value at a base row (row
 * axis) or base column group (column axis). `baseRowIndex` / `baseColIndex` is
 * the resolved 0-based index of the base item within this value field's
 * sub-grid; a negative index means the report has no such item, and every cell
 * is blank, as Excel shows for an unknown base.
 */
export function showAsDifferenceFrom(
  grid: PivotValueGrid,
  axis: "row" | "col",
  baseRowIndex: number,
  baseColIndex: number,
  percent: boolean,
): PivotShownValue[][] {
  const mode: PivotShowValuesAs = percent
    ? "percentDifferenceFrom"
    : "differenceFrom";
  return axis === "col"
    ? alongColumns(grid, mode, baseColIndex)
    : alongRows(grid, mode, baseRowIndex);
}

// ---------------------------------------------------------------------------
// The pivot engine.
// ---------------------------------------------------------------------------

// Join a tuple of field values into a stable group key. The NUL char
// (U+0000) cannot appear in a spreadsheet display string, so two different
// tuples can never collapse to the same key (e.g. ["a","b c"] vs
// ["a b","c"] stay distinct).
const KEY_SEP = String.fromCharCode(0);
function tupleKey(values: string[]): string {
  return values.join(KEY_SEP);
}

/** Read the header-label for a source column (matrix[0]); "" when absent. */
export function headerLabel(matrix: PivotMatrix, colIndex: number): string {
  const header = Array.isArray(matrix) ? matrix[0] : undefined;
  const raw = Array.isArray(header) ? header[colIndex] : undefined;
  return raw == null ? "" : String(raw);
}

/**
 * Compute the pivot grid from a source matrix (header row + data rows) and a
 * config. Returns empty axes / no cells when the matrix has no data rows or the
 * config has no value fields. The result GRID is read-only and never persisted —
 * only the config is, and this recomputes from current data on every call.
 *
 * Algorithm:
 *  1. distinct row-field tuples (first-seen order) -> rowHeaders.
 *  2. distinct column-field tuples (first-seen order) x value fields ->
 *     colHeaders. With no column fields, one column per value field.
 *  3. bucket each data row by (rowTuple, colTuple) and accumulate the raw cell
 *     strings of every SOURCE column any value field needs (a source field needs
 *     its own column; a CALCULATED field needs every source column its formula
 *     references). Strings are pooled per (row, col, sourceColumn) so every
 *     aggregate AND its totals re-read the true underlying values.
 *  4. evaluate each output cell + grand totals. A source value field aggregates
 *     its column with its agg; a CALCULATED field follows Excel's documented
 *     semantic — for each group it SUMs every referenced source column over that
 *     group and evaluates the formula on those sums (so a calc field's totals are
 *     recomputed on the summed components of the broader group, never the sum of
 *     sub-cell calc values).
 *  5. apply each value field's optional "Show Values As" display transform over
 *     the assembled result grid (a pure post-aggregation step).
 */
/**
 * One value field's aggregate over some records: a source field by its
 * summary function, a calculated field from its fields' sums (as
 * computePivot evaluates them). What a Value filter or a Top 10 judges by.
 */
export function pivotValueFieldAggregate(
  config: Pick<PivotConfig, "values" | "calculatedFields">,
  header: readonly string[],
  rows: readonly string[][],
  valueIndex: number,
): number | null {
  const field = config.values[valueIndex];
  if (!field) return null;
  const column = (col: number) => rows.map((row) => (row[col] == null ? "" : String(row[col])));
  const calc = field.calcRef
    ? (config.calculatedFields ?? []).find(
        (entry) => entry.name.trim().toLowerCase() === field.calcRef!.trim().toLowerCase(),
      )
    : undefined;
  if (calc) {
    const calcByName = new Map(
      (config.calculatedFields ?? []).map((entry) => [entry.name.trim().toLowerCase(), entry] as const),
    );
    const sumOf = (want: string): number | undefined => {
      const col = header.findIndex((label) => String(label ?? "").trim().toLowerCase() === want);
      return col >= 0 ? (aggSum(column(col)) ?? 0) : undefined;
    };
    return pivotCalcFieldValue(calc.formula, sumOf, calcByName);
  }
  return aggregate(field.agg, column(field.colIndex));
}

/** The records a report's field filters leave (pivot-field-filters), on the
 * expanded recipe: each record's item on a field is its group label there. */
export function pivotRowsAfterFieldFilters(
  rows: string[][],
  header: readonly string[],
  config: PivotConfig,
  today?: Date,
): string[][] {
  if (!config.fieldFilters?.length) return rows;
  const groupingByCol = new Map<number, PivotGrouping>();
  for (const g of config.grouping ?? []) groupingByCol.set(g.colIndex, g);
  return applyPivotFieldFilters(rows, config.fieldFilters, {
    rowFields: config.rows,
    columnFields: config.columns,
    itemOf: (row, col) => {
      const raw = row[col] == null ? "" : String(row[col]);
      const g = groupingByCol.get(col);
      return g ? pivotBucketLabel(raw, g) : raw;
    },
    measure: (group, valueIndex) =>
      pivotValueFieldAggregate(config, header, group as string[][], valueIndex),
    valueCount: config.values.length,
    today,
  });
}

export function computePivot(
  sourceMatrix: PivotMatrix,
  sourceConfig: PivotConfig,
): PivotResult {
  // A date field grouped by several units is several axis fields.
  const { matrix, config } = expandPivotGroupLevels(sourceMatrix, sourceConfig);
  // The expanded rows keep the source order, so a row's position is its
  // source row; the cells marked as text are looked up by it when pooled.
  const textCells = PIVOT_TEXT_CELLS.get(sourceMatrix);
  const sourceRowOf = textCells
    ? new Map<string[], number>(matrix.slice(1).map((row, index) => [row, index + 1]))
    : null;
  const empty: PivotResult = {
    rowHeaders: [],
    colHeaders: [],
    cells: [],
    totals: { row: [], column: [], grand: [] },
  };

  if (!Array.isArray(matrix) || matrix.length <= 1) return empty;
  // A pivot with axis fields but no value fields is Excel's label list: the
  // tuples are still grouped, there are simply no value columns.
  if (
    config.values.length === 0 &&
    config.rows.length === 0 &&
    config.columns.length === 0
  ) {
    return empty;
  }

  // Report/page filters pare the source rows BEFORE any grouping/aggregation, so
  // every header, cell and total downstream reflects only the kept rows; a
  // field's Label / Date / Value filters then pare what they leave.
  const reportFiltered = applyPivotFilters(
    matrix.slice(1),
    [...config.filters, ...(config.itemFilters ?? [])],
    config.dateFilters,
    config.grouping,
  );
  const dataRows = pivotRowsAfterFieldFilters(
    [...reportFiltered, ...pivotCalculatedItemRecords(reportFiltered, matrix[0] ?? [], config)],
    matrix[0] ?? [],
    config,
  );

  // Resolve, per value field, the SOURCE columns it needs and how to evaluate
  // it from per-column string pools. A calc field maps each referenced header
  // name to a source column (via headerLabel); an unresolved name is dropped
  // from its needs and treated as 0 by the formula evaluator. The union of all
  // needed columns is bucketed once per group.
  const calcByName = new Map<string, PivotCalculatedField>();
  // Defensive: a hand-built / legacy config may omit calculatedFields entirely.
  for (const cf of config.calculatedFields ?? []) {
    calcByName.set(cf.name.trim().toLowerCase(), cf);
  }
  const colByHeader = (label: string): number | null => {
    const want = label.trim().toLowerCase();
    if (want === "") return null;
    const header = Array.isArray(matrix[0]) ? matrix[0] : [];
    for (let c = 0; c < header.length; c += 1) {
      const h = header[c] == null ? "" : String(header[c]);
      if (h.trim().toLowerCase() === want) return c;
    }
    return null;
  };

  // For each value field: the set of source columns it reads + an evaluator that
  // turns per-column string pools (for one group) into the field's number.
  interface ValuePlan {
    neededCols: number[];
    evaluate: (pool: Map<number, string[]>) => number | null;
  }
  const neededColSet = new Set<number>();
  const plans: ValuePlan[] = config.values.map((vf) => {
    if (vf.calcRef && calcByName.has(vf.calcRef.trim().toLowerCase())) {
      const cf = calcByName.get(vf.calcRef.trim().toLowerCase())!;
      // Map each referenced field NAME -> source column (drop unknown names),
      // through the calculated fields it names as well.
      const nameToCol = new Map<string, number>();
      for (const name of pivotCalcSourceFieldNames(cf.formula, calcByName)) {
        const col = colByHeader(name);
        if (col !== null) {
          nameToCol.set(name.trim().toLowerCase(), col);
          neededColSet.add(col);
        }
      }
      return {
        neededCols: [...new Set(nameToCol.values())],
        evaluate: (pool) => {
          // Excel: each referenced field = SUM(field) over the group, then the
          // formula. A field with no numeric value in the group sums to 0.
          const sumOf = (nameKey: string): number | undefined => {
            const col = nameToCol.get(nameKey);
            return col === undefined ? undefined : (aggSum(pool.get(col) ?? []) ?? 0);
          };
          return pivotCalcFieldValue(cf.formula, sumOf, calcByName);
        },
      };
    }
    // Source value field: aggregate its own column with its agg.
    neededColSet.add(vf.colIndex);
    const colIndex = vf.colIndex;
    const agg = vf.agg;
    return {
      neededCols: [colIndex],
      evaluate: (pool) => {
        // Excel leaves a cell with nothing to count empty rather than writing 0.
        const result = aggregate(agg, pool.get(colIndex) ?? []);
        return (agg === "count" || agg === "countNums") && result === 0
          ? null
          : result;
      },
    };
  });
  const neededCols = [...neededColSet];

  // Field GROUPING (Excel "Group field"): bucket a ROW/COLUMN axis value before
  // the tuples form, so the pivot groups by the bucket (date period / number bin)
  // instead of the raw value. Aggregated VALUE columns are never bucketed; an
  // unparseable value passes through ungrouped; no grouping for a column → its raw
  // string, byte-for-byte as before.
  const groupingByCol = new Map<number, PivotGrouping>();
  for (const g of config.grouping ?? []) groupingByCol.set(g.colIndex, g);
  // Excel compares a PivotTable's text items without regard to case: "Scone"
  // and "SCONE" are one item, spelled as the source first spells it (the
  // server's cache writer folds them the same way, or Excel repairs it).
  const firstSpelling = new Map<number, Map<string, string>>();
  for (const ci of new Set([...config.rows, ...config.columns])) {
    const spellings = new Map<string, string>();
    // The spelling the source's first row of that item has, filtered or not
    // (the cache is built from every source row).
    for (const row of [...matrix.slice(1), ...dataRows]) {
      const raw = row[ci] == null ? "" : String(row[ci]);
      const folded = raw.toLowerCase();
      if (!spellings.has(folded)) spellings.set(folded, raw);
    }
    firstSpelling.set(ci, spellings);
  }
  const fieldValueOf = (row: string[], ci: number): string => {
    const raw = row[ci] == null ? "" : String(row[ci]);
    const g = groupingByCol.get(ci);
    if (g) return pivotBucketLabel(raw, g);
    return firstSpelling.get(ci)?.get(raw.toLowerCase()) ?? raw;
  };
  // Group items sort in their own order (calendar, ascending), as Excel
  // lists them, however the source rows come.
  const groupOrders = new Map<number, Map<string, number>>();
  for (const ci of [...config.rows, ...config.columns]) {
    const order = pivotGroupItemOrder(
      groupingByCol.get(ci),
      dataRows.map((row) => fieldValueOf(row, ci)),
    );
    if (order) groupOrders.set(ci, order);
  }

  // --- 1. Row-field tuples (first-seen order). ---
  const rowTupleOrder: string[][] = [];
  const rowTupleIndex = new Map<string, number>();
  const rowKeyOf = (row: string[]): string =>
    tupleKey(config.rows.map((ci) => fieldValueOf(row, ci)));
  // When there are no row fields the whole table is one (empty-tuple) group.
  if (config.rows.length === 0) {
    rowTupleOrder.push([]);
    rowTupleIndex.set(tupleKey([]), 0);
  } else {
    for (const row of dataRows) {
      const tuple = config.rows.map((ci) => fieldValueOf(row, ci));
      const key = tupleKey(tuple);
      if (!rowTupleIndex.has(key)) {
        rowTupleIndex.set(key, rowTupleOrder.length);
        rowTupleOrder.push(tuple);
      }
    }
  }

  sortAxisTuples(rowTupleOrder, rowTupleIndex, config.rows, config.fieldSorts, groupOrders, config.fieldItemOrders);

  // --- 2. Column-field tuples (first-seen order). ---
  const colTupleOrder: string[][] = [];
  const colTupleIndex = new Map<string, number>();
  if (config.columns.length === 0) {
    colTupleOrder.push([]);
    colTupleIndex.set(tupleKey([]), 0);
  } else {
    for (const row of dataRows) {
      const tuple = config.columns.map((ci) => fieldValueOf(row, ci));
      const key = tupleKey(tuple);
      if (!colTupleIndex.has(key)) {
        colTupleIndex.set(key, colTupleOrder.length);
        colTupleOrder.push(tuple);
      }
    }
  }

  sortAxisTuples(colTupleOrder, colTupleIndex, config.columns, config.fieldSorts, groupOrders, config.fieldItemOrders);

  const colKeyOf = (row: string[]): string =>
    tupleKey(config.columns.map((ci) => fieldValueOf(row, ci)));

  // Output columns = (column tuple x value field), value field is the inner loop
  // so the same column tuple's value fields stay adjacent (Excel layout).
  const colHeaders: PivotColumnHeader[] = [];
  for (let ct = 0; ct < colTupleOrder.length; ct += 1) {
    for (let vi = 0; vi < config.values.length; vi += 1) {
      colHeaders.push({ columnValues: colTupleOrder[ct], valueIndex: vi });
    }
  }

  // --- 3. Pool raw cell strings per (rowTuple, colTuple, sourceColumn). ---
  // pool[rowIdx][colTupleIdx].get(sourceCol) = string[]. Only the source
  // columns any value field needs are pooled (a source value field needs its
  // own column; a calc field needs each referenced column).
  const newPool = (): Map<number, string[]> => {
    const m = new Map<number, string[]>();
    for (const c of neededCols) m.set(c, []);
    return m;
  };
  const pool: Map<number, string[]>[][] = rowTupleOrder.map(() =>
    colTupleOrder.map(() => newPool()),
  );
  for (const row of dataRows) {
    const ri = config.rows.length === 0 ? 0 : rowTupleIndex.get(rowKeyOf(row));
    const ci =
      config.columns.length === 0 ? 0 : colTupleIndex.get(colKeyOf(row));
    if (ri === undefined || ci === undefined) continue;
    const cell = pool[ri][ci];
    const sourceRow = sourceRowOf?.get(row);
    for (const c of neededCols) {
      const raw = row[c];
      const text = raw == null ? "" : String(raw);
      cell.get(c)!.push(
        sourceRow !== undefined && textCells!.has(`${sourceRow}:${c}`)
          ? PIVOT_TEXT_MARK + text
          : text,
      );
    }
  }

  // --- 4. Evaluate every output cell + grand totals (per value plan). ---
  const cells: (number | null)[][] = [];
  const totalRow: (number | null)[] = colHeaders.map(() => null); // per output col
  const totalColumn: (number | null)[] = []; // per output row
  // grand[valueIdx] across the whole table (all rows + all column tuples).
  const grand: (number | null)[] = config.values.map(() => null);
  // Pool the RAW strings across the whole table per output column and per
  // (row, value) so totals are a true re-evaluate (correct for average/min/max
  // and for calc fields recomputed on summed components, not a sum of cells).
  const colTotalPool: Map<number, string[]>[] = colHeaders.map(() => newPool());
  const grandPool: Map<number, string[]>[] = config.values.map(() => newPool());

  const mergePool = (
    dest: Map<number, string[]>,
    src: Map<number, string[]>,
  ): void => {
    for (const c of neededCols) {
      const into = dest.get(c)!;
      for (const s of src.get(c)!) into.push(s);
    }
  };

  for (let ri = 0; ri < rowTupleOrder.length; ri += 1) {
    const outRow: (number | null)[] = [];
    const rowTotalPool: Map<number, string[]>[] = config.values.map(() =>
      newPool(),
    );
    let outCol = 0;
    for (let ct = 0; ct < colTupleOrder.length; ct += 1) {
      const groupPool = pool[ri][ct];
      for (let vi = 0; vi < config.values.length; vi += 1) {
        outRow.push(plans[vi].evaluate(groupPool));
        mergePool(colTotalPool[outCol], groupPool);
        mergePool(rowTotalPool[vi], groupPool);
        mergePool(grandPool[vi], groupPool);
        outCol += 1;
      }
    }
    cells.push(outRow);

    // Row grand total: re-evaluate this row's pooled strings per value field.
    // When there are multiple value fields the row total is per value field, so
    // we expose the FIRST value field's row total as the canonical column[ri]
    // (matching Excel, which shows one grand-total column per value field; the
    // component renders the per-value totals from `cells` + this array). We keep
    // a single meaningful number only when there is one value field.
    if (config.values.length === 1) {
      totalColumn.push(plans[0].evaluate(rowTotalPool[0]));
    } else {
      // Multi-value: row total column is not a single scalar; expose null and
      // let the renderer omit it (Excel shows per-value subtotals instead).
      totalColumn.push(null);
    }
  }

  // Column grand totals: re-evaluate each output column's pooled strings.
  for (let oc = 0; oc < colHeaders.length; oc += 1) {
    totalRow[oc] = plans[colHeaders[oc].valueIndex].evaluate(colTotalPool[oc]);
  }
  // Overall per-value-field grand total.
  for (let vi = 0; vi < config.values.length; vi += 1) {
    grand[vi] = plans[vi].evaluate(grandPool[vi]);
  }

  // --- 4b. Optional SUBTOTALS (opt-in): sort rows so each outer group is
  // contiguous (Excel sorts pivot rows by default), then emit one subtotal row
  // per outer group re-aggregated from the source. Skipped when a "Show Values
  // As" transform is active (subtotals would mix raw aggregates with transformed
  // leaf cells). Default-off leaves the rows in first-seen order, byte-stable. ---
  let subtotals: PivotSubtotalRow[] | undefined;
  const showAsActive = config.values.some(
    (v) => v.showAs && v.showAs.mode !== "normal",
  );
  if (
    config.rowSubtotals &&
    config.rows.length >= 2 &&
    rowTupleOrder.length > 0 &&
    !showAsActive
  ) {
    // Sort the row-parallel arrays (headers + cells + per-row total) together by
    // the full row tuple. This matches Excel's default sort AND groups each outer
    // value contiguously so its subtotal can sit after its last leaf.
    const order = rowTupleOrder.map((_unused, i) => i);
    order.sort((a, b) => compareTuples(rowTupleOrder[a], rowTupleOrder[b]));
    const sortedHeaders = order.map((i) => rowTupleOrder[i]);
    const sortedCells = order.map((i) => cells[i]);
    const sortedTotalColumn = order.map((i) => totalColumn[i]);
    rowTupleOrder.splice(0, rowTupleOrder.length, ...sortedHeaders);
    cells.splice(0, cells.length, ...sortedCells);
    totalColumn.splice(0, totalColumn.length, ...sortedTotalColumn);

    subtotals = [];
    let start = 0;
    while (start < rowTupleOrder.length) {
      const outer = rowTupleOrder[start][0] ?? "";
      let end = start;
      while (
        end + 1 < rowTupleOrder.length &&
        (rowTupleOrder[end + 1][0] ?? "") === outer
      ) {
        end += 1;
      }
      // Re-pool the group's source rows (matching the OUTER field's bucketed
      // value) per column tuple, then evaluate per value field — a true
      // re-aggregate (correct for average/min/max + calc fields), not a sum of
      // the leaf cells.
      const groupPools = colTupleOrder.map(() => newPool());
      for (const row of dataRows) {
        if (fieldValueOf(row, config.rows[0]) !== outer) continue;
        const ci =
          config.columns.length === 0 ? 0 : colTupleIndex.get(colKeyOf(row));
        if (ci === undefined) continue;
        for (const c of neededCols) {
          const raw = row[c];
          groupPools[ci].get(c)!.push(raw == null ? "" : String(raw));
        }
      }
      const subCells: (number | null)[] = [];
      for (let ct = 0; ct < colTupleOrder.length; ct += 1) {
        for (let vi = 0; vi < config.values.length; vi += 1) {
          subCells.push(plans[vi].evaluate(groupPools[ct]));
        }
      }
      subtotals.push({ afterRowIndex: end, outerValue: outer, cells: subCells });
      start = end + 1;
    }
  }

  const result: PivotResult = {
    rowHeaders: rowTupleOrder,
    colHeaders,
    cells,
    totals: { row: totalRow, column: totalColumn, grand },
    ...(subtotals ? { subtotals } : {}),
    ...(config.columns.length > 0 ? { columnTuples: colTupleOrder } : {}),
  };

  // --- 5. Apply each value field's "Show Values As" display transform. ---
  return applyShowValuesAsToResult(result, config);
}

/**
 * Compare two row tuples for the subtotal sort: element by element, numerically
 * when both sides parse as finite numbers (so "9" sorts before "10"), else by
 * string order. Shorter tuples sort first when one is a prefix of the other.
 */
/** Validated, de-duplicated per-field number formats. */
export function normalizePivotFieldNumberFormats(raw: unknown): PivotFieldNumberFormat[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldNumberFormat[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, 128)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, numberFormat } = entry as Record<string, unknown>;
    if (
      typeof colIndex !== "number" ||
      !Number.isInteger(colIndex) ||
      colIndex < 0 ||
      seen.has(colIndex) ||
      typeof numberFormat !== "string" ||
      numberFormat === "" ||
      numberFormat.length > 255
    ) {
      continue;
    }
    seen.add(colIndex);
    out.push({ colIndex, numberFormat });
  }
  return out;
}

/** Validated, de-duplicated per-field sorts (first entry per column wins). */
export function normalizePivotFieldItemOrders(raw: unknown): PivotFieldItemOrder[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldItemOrder[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, 128)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, items } = entry as Record<string, unknown>;
    if (
      typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 ||
      seen.has(colIndex) || !Array.isArray(items) || items.length === 0 || items.length > 10_000 ||
      items.some((item) => typeof item !== "string")
    ) {
      continue;
    }
    seen.add(colIndex);
    out.push({ colIndex, items: [...(items as string[])] });
  }
  return out;
}

export function normalizePivotFieldSorts(raw: unknown): PivotFieldSort[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldSort[] = [];
  const seen = new Set<number>();
  for (const entry of raw.slice(0, 128)) {
    if (!entry || typeof entry !== "object") continue;
    const { colIndex, order, byValue } = entry as Record<string, unknown>;
    if (
      typeof colIndex !== "number" ||
      !Number.isInteger(colIndex) ||
      colIndex < 0 ||
      seen.has(colIndex) ||
      (order !== "ascending" && order !== "descending")
    ) {
      continue;
    }
    seen.add(colIndex);
    out.push({
      colIndex,
      order,
      ...(typeof byValue === "number" && Number.isInteger(byValue) && byValue >= 0 && byValue < 8
        ? { byValue }
        : {}),
    });
  }
  return out;
}

/**
 * Order an axis's tuples by its fields' sorts: a sorted field's items run
 * ascending / descending (numbers numerically), an unsorted one keeps
 * first-seen order. Rebuilds the tuple index in place.
 */
function sortAxisTuples(
  order: string[][],
  index: Map<string, number>,
  fields: readonly number[],
  sorts: readonly PivotFieldSort[] | undefined,
  groupOrders?: ReadonlyMap<number, ReadonlyMap<string, number>>,
  itemOrders?: readonly PivotFieldItemOrder[],
): void {
  const byField = fields.map(
    (field) => sorts?.find((sort) => sort.colIndex === field)?.order,
  );
  const byGroup = fields.map((field) => groupOrders?.get(field));
  // An unsorted field's manual item order, from its PivotField item list.
  const byManual = fields.map((field) => {
    const entry = itemOrders?.find((order) => order.colIndex === field);
    return entry ? new Map(entry.items.map((item, position) => [item, position] as const)) : undefined;
  });
  if (!byField.some(Boolean) && !byGroup.some(Boolean) && !byManual.some(Boolean)) return;
  const firstSeen = fields.map(() => new Map<string, number>());
  for (const tuple of order) {
    tuple.forEach((value, level) => {
      const seen = firstSeen[level];
      if (seen && !seen.has(value)) seen.set(value, seen.size);
    });
  }
  order.sort((a, b) => {
    for (let level = 0; level < fields.length; level += 1) {
      const x = a[level] ?? "";
      const y = b[level] ?? "";
      if (x === y) continue;
      const direction = byField[level];
      const groupOrder = byGroup[level];
      if (!direction && groupOrder) {
        // A grouped field's items keep their own order (calendar, ascending).
        return (groupOrder.get(x) ?? 0) - (groupOrder.get(y) ?? 0);
      }
      if (!direction) {
        const manual = byManual[level];
        if (manual) {
          // Named items in the list's order; the rest after them, first-seen.
          const px = manual.get(x);
          const py = manual.get(y);
          if (px !== undefined || py !== undefined) {
            if (px === undefined) return 1;
            if (py === undefined) return -1;
            return px - py;
          }
        }
        return (firstSeen[level]?.get(x) ?? 0) - (firstSeen[level]?.get(y) ?? 0);
      }
      const compared = compareTuples([x], [y]);
      return direction === "descending" ? -compared : compared;
    }
    return 0;
  });
  index.clear();
  order.forEach((tuple, position) => index.set(tupleKey(tuple), position));
}

function compareTuples(a: string[], b: string[]): number {
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n; i += 1) {
    const x = a[i] ?? "";
    const y = b[i] ?? "";
    if (x === y) continue;
    const nx = Number(x);
    const ny = Number(y);
    const bothNumeric =
      x.trim() !== "" &&
      y.trim() !== "" &&
      Number.isFinite(nx) &&
      Number.isFinite(ny);
    if (bothNumeric) return nx - ny;
    return x < y ? -1 : 1;
  }
  return 0;
}

/**
 * Apply every value field's `showAs` display transform to a computed result,
 * returning a NEW result with the transformed cells/totals (an untransformed
 * field is left exactly as-is). Each value field owns a disjoint set of output
 * columns (the ones whose `valueIndex` is its index), so a transform rewrites
 * only those columns + the matching grand-total cells.
 *
 * This is the flat leaf grid the formula surface reads (GETPIVOTDATA); the
 * report the sheet shows is laid out by `planPivotOutput`, which runs the same
 * per-cell rules over its subtotal and grand-total rows as well. A #DIV/0!
 * (a base of 0) is kept as the grid's #DIV/0! mark (PIVOT_AGG_DIVIDE_BY_ZERO),
 * which the Field List preview, the chart and GETPIVOTDATA read as the error
 * -- read as blank, the preview showed a failed calculation as no data (user
 * audit S-027); #N/A has no numeric mark and reads as blank here, the report
 * carrying it. Pure: it reads only the aggregates already in `result`.
 */
export function applyShowValuesAsToResult(
  result: PivotResult,
  config: PivotConfig,
): PivotResult {
  const hasAny = config.values.some((v) => v.showAs && v.showAs.mode !== "normal");
  if (!hasAny) return result;

  const rowCount = result.rowHeaders.length;
  // Deep-copy the cells + totals so the transform never mutates the input.
  const cells = result.cells.map((row) => row.slice());
  const totalRow = result.totals.row.slice();
  const totalColumn = result.totals.column.slice();
  const grand = result.totals.grand.slice();
  const singleValue = config.values.length === 1;
  const numeric = (value: PivotShownValue): number | null =>
    typeof value === "number"
      ? value
      : value === PIVOT_SHOW_AS_DIVIDE_BY_ZERO
        ? PIVOT_AGG_DIVIDE_BY_ZERO
        : null;

  for (let vi = 0; vi < config.values.length; vi += 1) {
    const showAs = config.values[vi].showAs;
    if (!showAs || showAs.mode === "normal") continue;
    const mode = showAs.mode;

    // The output-column indices belonging to this value field, in axis order.
    const outCols: number[] = [];
    for (let c = 0; c < result.colHeaders.length; c += 1) {
      if (result.colHeaders[c].valueIndex === vi) outCols.push(c);
    }
    if (outCols.length === 0) continue;

    // This field's own totals: per row across its columns, per column down all
    // rows, and overall. A lone value field has them precomputed.
    const rowTotalOf = (r: number): number | null => {
      if (singleValue) return result.totals.column[r] ?? null;
      let sum = 0;
      let any = false;
      for (const c of outCols) {
        const value = result.cells[r][c];
        if (value !== null && value !== undefined) {
          sum += value;
          any = true;
        }
      }
      return any ? sum : null;
    };
    const grandOf = result.totals.grand[vi] ?? null;

    if (!showValuesAsNeedsBaseField(mode)) {
      // A share of a total: every cell takes part, the totals included, so the
      // report's own Grand Total row and column read as 100%. The parent of a
      // leaf row here is the group it belongs to, summed from the group's own
      // cells (this grid holds no subtotal rows; the report computes them).
      const parentRowTotal = (r: number, c: number): number | null => {
        const tuple = result.rowHeaders[r];
        if (tuple.length <= 1) return c < 0 ? grandOf : (result.totals.row[c] ?? null);
        const prefix = tupleKey(tuple.slice(0, -1));
        let sum = 0;
        let any = false;
        for (let other = 0; other < rowCount; other += 1) {
          if (tupleKey(result.rowHeaders[other].slice(0, -1)) !== prefix) continue;
          const value = c < 0 ? rowTotalOf(other) : (result.cells[other][c] ?? null);
          if (value !== null) {
            sum += value;
            any = true;
          }
        }
        return any ? sum : null;
      };
      const denominator = (r: number, c: number): number | null => {
        switch (mode) {
          case "percentOfGrandTotal":
            return grandOf;
          case "percentOfRowTotal":
            return rowTotalOf(r);
          case "percentOfParentRowTotal":
            return parentRowTotal(r, c);
          case "percentOfParentColumnTotal":
            // One column field here, so a column group's parent is the row.
            return rowTotalOf(r);
          default:
            return c < 0 ? grandOf : (result.totals.row[c] ?? null);
        }
      };
      const shownAt = (r: number, c: number, cell: number | null): PivotShownValue =>
        mode === "index"
          ? showAsIndex(
              cell,
              rowTotalOf(r),
              c < 0 ? grandOf : (result.totals.row[c] ?? null),
              grandOf,
            )
          : showAsShareOfTotal(cell, denominator(r, c));
      for (let r = 0; r < rowCount; r += 1) {
        for (const c of outCols) {
          cells[r][c] = numeric(shownAt(r, c, result.cells[r][c] ?? null));
        }
        // The Grand Total column, which every share divides as well.
        if (singleValue) totalColumn[r] = numeric(shownAt(r, -1, rowTotalOf(r)));
      }
      for (const c of outCols) {
        // The Grand Total row: its own column total, so "% of column" reads 1.
        const columnTotal = result.totals.row[c] ?? null;
        totalRow[c] = numeric(
          mode === "index"
            ? showAsIndex(columnTotal, grandOf, columnTotal, grandOf)
            : showAsShareOfTotal(
                columnTotal,
                mode === "percentOfColumnTotal" || mode === "percentOfParentRowTotal"
                  ? columnTotal
                  : grandOf,
              ),
        );
      }
      grand[vi] = numeric(
        mode === "index"
          ? showAsIndex(grandOf, grandOf, grandOf, grandOf)
          : showAsShareOfTotal(grandOf, grandOf),
      );
      continue;
    }

    // A transform along a base field. Its axis takes part only through that
    // field's items, so the axis' own grand total goes blank; the other axis
    // takes part in full, its grand total included.
    const base = showAs.base;
    const axis =
      base?.axis ?? (mode === "runningTotalInColumns" ? "col" : "row");
    const axisFields = axis === "row" ? config.rows : config.columns;
    // Without a named field the mode runs along the innermost field of its
    // axis, which is what a report with one field on it always means.
    const pos = base && base.pos < axisFields.length ? base.pos : axisFields.length - 1;
    const wantsItem = showValuesAsNeedsBaseItem(mode);
    // Excel writes #N/A through a whole value field whose base field is not on
    // an axis at all (it reads as blank in this numeric grid); a mode that
    // names no field of its own keeps running along the innermost one.
    const blank =
      pos < 0 ||
      (base !== undefined && base.pos >= axisFields.length) ||
      (wantsItem && (!base || base.axis !== axis));
    // "% of Parent Total" divides by the total of the base field's own item,
    // which this grid holds only as the sum of the group's leaf cells.
    const parentTotalOf = (members: number[], line: (number | null)[]): number | null => {
      let sum = 0;
      let any = false;
      for (let k = 0; k < members.length; k += 1) {
        const value = line[k];
        if (value !== null) {
          sum += value;
          any = true;
        }
      }
      return any ? sum : null;
    };

    if (axis === "row") {
      // Rows that differ only at `pos` are one another's members.
      const groups = new Map<string, number[]>();
      for (let r = 0; r < rowCount; r += 1) {
        const tuple = result.rowHeaders[r];
        if (tuple.length <= pos) continue;
        const key = tupleKey(tuple.filter((_value, index) => index !== pos));
        const members = groups.get(key);
        if (members) members.push(r);
        else groups.set(key, [r]);
      }
      for (const members of groups.values()) {
        const baseIndex = wantsItem
          ? members.findIndex((r) =>
              looseEqual(result.rowHeaders[r][pos] ?? "", base?.item ?? ""),
            )
          : -1;
        const columns = singleValue ? [...outCols, -1] : outCols;
        for (const c of columns) {
          const source = (r: number) =>
            c < 0 ? rowTotalOf(r) : (result.cells[r][c] ?? null);
          const line = members.map(source);
          const parentTotal =
            mode === "percentOfParentTotal" ? parentTotalOf(members, line) : null;
          members.forEach((r, index) => {
            const shown = blank
              ? null
              : numeric(
                  mode === "percentOfParentTotal"
                    ? showAsShareOfTotal(line[index], parentTotal)
                    : showValuesAsAlongBase(mode, line, index, baseIndex),
                );
            if (c < 0) totalColumn[r] = shown;
            else cells[r][c] = shown;
          });
        }
      }
      for (let r = 0; r < rowCount; r += 1) {
        if (result.rowHeaders[r].length > pos) continue;
        for (const c of outCols) cells[r][c] = null;
        if (singleValue) totalColumn[r] = null;
      }
      for (const c of outCols) totalRow[c] = null;
      grand[vi] = null;
      continue;
    }

    // Column axis: the column groups that differ only at `pos` are members,
    // and every row takes part, the Grand Total row included.
    const groups = new Map<string, number[]>();
    for (const c of outCols) {
      const tuple = result.colHeaders[c].columnValues;
      if (tuple.length <= pos) continue;
      const key = tupleKey(tuple.filter((_value, index) => index !== pos));
      const members = groups.get(key);
      if (members) members.push(c);
      else groups.set(key, [c]);
    }
    for (const members of groups.values()) {
      const baseIndex = wantsItem
        ? members.findIndex((c) =>
            looseEqual(result.colHeaders[c].columnValues[pos] ?? "", base?.item ?? ""),
          )
        : -1;
      for (let r = -1; r < rowCount; r += 1) {
        const source = (c: number) =>
          r < 0 ? (result.totals.row[c] ?? null) : (result.cells[r][c] ?? null);
        const line = members.map(source);
        const parentTotal =
          mode === "percentOfParentTotal" ? parentTotalOf(members, line) : null;
        members.forEach((c, index) => {
          const shown = blank
            ? null
            : numeric(
                mode === "percentOfParentTotal"
                  ? showAsShareOfTotal(line[index], parentTotal)
                  : showValuesAsAlongBase(mode, line, index, baseIndex),
              );
          if (r < 0) totalRow[c] = shown;
          else cells[r][c] = shown;
        });
      }
    }
    for (const c of outCols) {
      if (result.colHeaders[c].columnValues.length > pos) continue;
      for (let r = 0; r < rowCount; r += 1) cells[r][c] = null;
      totalRow[c] = null;
    }
    if (singleValue) for (let r = 0; r < rowCount; r += 1) totalColumn[r] = null;
    grand[vi] = null;
  }

  return {
    rowHeaders: result.rowHeaders,
    colHeaders: result.colHeaders,
    cells,
    totals: { row: totalRow, column: totalColumn, grand },
    ...(result.columnTuples ? { columnTuples: result.columnTuples } : {}),
  };
}

// ---------------------------------------------------------------------------
// GETPIVOTDATA support (Excel parity).
//
// GETPIVOTDATA(data_field, pivot_table, [field1, item1], …) reads ONE aggregated
// value out of a pivot table by criteria, rather than by cell position — so a
// formula keeps pointing at "Sum of Sales for East / 2024" even when the pivot
// re-lays-out. The formula engine resolves the `pivot_table` argument to an
// ANCHOR cell and hands it (plus the data field + the field/item pairs) to this
// module via the EvalContext.resolvePivotData seam; everything here is pure so it
// unit-tests without React, exactly like computePivot.
//
// Resolution:
//  1. Find the placement whose OUTPUT range contains the anchor cell. None ->
//     "not-pivot" (engine: #REF!).
//  2. Recompute the pivot grid (computePivot) and pick the value-field column
//     matching `data_field` (its "Sum of Sales" label, or the bare source-field
//     name when unambiguous).
//  3. Filter the rows/columns to the ones whose field values satisfy EVERY
//     (field, item) pair; with no pairs, take the grand total of that value
//     field. A single surviving cell -> its value; none, an unknown field/item,
//     or an unknown data field -> "no-match" (engine: #N/A).
//
// DEFERRED (documented, not faked): when two value fields would BOTH match an
// ambiguous bare `data_field`, this returns no-match rather than guessing; and a
// (field,item) referring to a CALCULATED field/item is not modelled (the native
// pivot has no calculated fields yet) — such a field name simply finds no match.
// ---------------------------------------------------------------------------

/**
 * One pivot table placed on a sheet, as GETPIVOTDATA needs to see it: the cell
 * RANGE its output occupies (so an anchor cell can be matched to it), plus the
 * recipe + source matrix needed to (re)compute its grid. `output` is a 0-based
 * inclusive rectangle; the anchor argument of GETPIVOTDATA must fall inside it
 * for this placement to be selected.
 */
export interface PivotPlacement {
  output: PivotRange;
  config: PivotConfig;
  matrix: PivotMatrix;
}

/** A resolved GETPIVOTDATA lookup outcome (the engine maps each to a value/error). */
export type PivotDataResult =
  | { kind: "value"; value: number | string | null }
  | { kind: "no-match" }
  | { kind: "not-pivot" };

/** One GETPIVOTDATA criterion: a field (header label) + the item to match. */
export interface PivotDataCriterion {
  field: string;
  item: string;
}

/** True when (row,col) lies inside an inclusive output rectangle. */
export function rangeContains(range: PivotRange, row: number, col: number): boolean {
  return (
    row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2
  );
}

/** Case/space-insensitive equality for field labels + item display strings. */
function looseEqual(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * Find which config.values index `dataField` names. Matches (case/space-
 * insensitively) the value field's LABEL — "Sum of Sales" for a source field,
 * the calc-field NAME for a calculated field — then the bare source field NAME
 * ("Sales") when exactly one value field uses that column. Returns the value
 * index, or null when no/an AMBIGUOUS match (Excel needs the labelled form to
 * disambiguate, so an ambiguous bare name is treated as not-found rather than
 * guessed). A calculated value field is addressed by its name (so
 * GETPIVOTDATA("Profit", …) resolves a calc field, no longer a DEFER).
 */
export function matchDataFieldIndex(
  matrix: PivotMatrix,
  config: PivotConfig,
  dataField: string,
): number | null {
  const want = dataField.trim();
  if (want === "") return null;
  // 1. Exact value-field-label match ("Sum of Sales" / calc-field name).
  for (let vi = 0; vi < config.values.length; vi += 1) {
    const field = config.values[vi];
    const label = valueFieldLabel(field, headerLabel(matrix, field.colIndex));
    if (looseEqual(label, want)) return vi;
  }
  // 2. Bare source-field-name match ("Sales"); only when exactly one SOURCE
  //    value field aggregates that column (else ambiguous -> not-found). A calc
  //    field has no source column, so it is matched only by its name above.
  const byName: number[] = [];
  for (let vi = 0; vi < config.values.length; vi += 1) {
    const field = config.values[vi];
    if (field.calcRef && field.calcRef.trim() !== "") continue;
    const name = headerLabel(matrix, field.colIndex);
    if (looseEqual(name, want)) byName.push(vi);
  }
  if (byName.length === 1) return byName[0];
  return null;
}

/**
 * Resolve a GETPIVOTDATA lookup against ONE placement's computed grid.
 * `dataField` selects the value field; `criteria` constrains which group cell.
 * Returns a value result (possibly null for an empty group), or "no-match" when
 * the data field / a criterion field / an item is unknown or no cell survives.
 */
function lookupInPlacement(
  placement: PivotPlacement,
  dataField: string,
  criteria: readonly PivotDataCriterion[],
): PivotDataResult {
  const { config, matrix } = placement;
  const valueIndex = matchDataFieldIndex(matrix, config, dataField);
  if (valueIndex === null) return { kind: "no-match" };

  // Map each criterion FIELD to the axis it constrains: a row-field position, a
  // column-field position, or unknown. An unknown field -> no-match.
  type Bound = { axis: "row" | "col"; pos: number; item: string };
  const bounds: Bound[] = [];
  for (const crit of criteria) {
    const rowPos = config.rows.findIndex((ci) =>
      looseEqual(headerLabel(matrix, ci), crit.field),
    );
    if (rowPos >= 0) {
      bounds.push({ axis: "row", pos: rowPos, item: crit.item });
      continue;
    }
    const colPos = config.columns.findIndex((ci) =>
      looseEqual(headerLabel(matrix, ci), crit.field),
    );
    if (colPos >= 0) {
      bounds.push({ axis: "col", pos: colPos, item: crit.item });
      continue;
    }
    return { kind: "no-match" }; // criterion field is not an axis of this pivot
  }

  const result = computePivot(matrix, config);

  // No criteria -> the grand total of the chosen value field.
  if (bounds.length === 0) {
    return { kind: "value", value: result.totals.grand[valueIndex] ?? null };
  }

  const rowBounds = bounds.filter((b) => b.axis === "row");
  const colBounds = bounds.filter((b) => b.axis === "col");

  // Rows whose row-field tuple satisfies every row criterion.
  const matchRows: number[] = [];
  for (let r = 0; r < result.rowHeaders.length; r += 1) {
    const tuple = result.rowHeaders[r];
    if (rowBounds.every((b) => looseEqual(tuple[b.pos] ?? "", b.item))) {
      matchRows.push(r);
    }
  }
  // Output columns whose (column tuple + value field) satisfy every col criterion
  // AND aggregate the chosen value field.
  const matchCols: number[] = [];
  for (let c = 0; c < result.colHeaders.length; c += 1) {
    const header = result.colHeaders[c];
    if (header.valueIndex !== valueIndex) continue;
    if (
      colBounds.every((b) => looseEqual(header.columnValues[b.pos] ?? "", b.item))
    ) {
      matchCols.push(c);
    }
  }

  // A fully-specified cross (one row group × one column group) addresses a single
  // body cell. When only one axis is constrained, the OTHER must collapse to the
  // grand total along it (Excel: an unconstrained axis means "all", i.e. the
  // total) — so we read from totals rather than picking an arbitrary group cell.
  const rowsConstrained = rowBounds.length > 0;
  const colsConstrained = colBounds.length > 0;

  if (rowsConstrained && colsConstrained) {
    if (matchRows.length !== 1 || matchCols.length !== 1) {
      return { kind: "no-match" };
    }
    const value = result.cells[matchRows[0]]?.[matchCols[0]] ?? null;
    return { kind: "value", value };
  }
  if (rowsConstrained) {
    // Row group total for the chosen value field. With a single value field the
    // row total column is the canonical scalar; otherwise re-read via the value
    // field's own column total within that row across all column groups.
    if (matchRows.length !== 1) return { kind: "no-match" };
    const r = matchRows[0];
    const value = rowGroupTotal(result, config, r, valueIndex);
    return { kind: "value", value };
  }
  // Only columns constrained: the column group's grand total down all rows.
  if (matchCols.length !== 1) return { kind: "no-match" };
  const value = result.totals.row[matchCols[0]] ?? null;
  return { kind: "value", value };
}

/**
 * The total of one value field across ALL column groups within a single output
 * ROW (so GETPIVOTDATA("Sum of Sales", …, "Region", "East") = East's row total).
 * For a single value field this is the precomputed `totals.column[r]`; for
 * multiple value fields that column is null (per-value subtotals aren't a single
 * scalar there), so we re-sum the value field's own output columns in that row.
 */
function rowGroupTotal(
  result: PivotResult,
  config: PivotConfig,
  rowIndex: number,
  valueIndex: number,
): number | null {
  if (config.values.length === 1) {
    return result.totals.column[rowIndex] ?? null;
  }
  const agg = config.values[valueIndex].agg;
  const cellsInRow: (number | null)[] = [];
  for (let c = 0; c < result.colHeaders.length; c += 1) {
    if (result.colHeaders[c].valueIndex !== valueIndex) continue;
    cellsInRow.push(result.cells[rowIndex]?.[c] ?? null);
  }
  // SUM/COUNT re-add across column groups; AVERAGE/MIN/MAX cannot be recombined
  // from sub-aggregates without the raw strings, so for those a multi-value
  // single-axis total is not exposed (no-match upstream would be cleaner, but a
  // best-effort numeric combine keeps SUM/COUNT — the common case — correct).
  const nums = cellsInRow.filter((n): n is number => n !== null);
  if (nums.length === 0) return null;
  if (agg === "min") return Math.min(...nums);
  if (agg === "max") return Math.max(...nums);
  if (agg === "product") return nums.reduce((a, b) => a * b, 1);
  // A deviation or variance cannot be rebuilt from its parts (XLS-A071).
  if (agg === "stdDev" || agg === "stdDevp" || agg === "var" || agg === "varp") {
    return null;
  }
  // sum / count / average(best-effort): sum the parts (average across groups is
  // not exact; documented DEFER — multi-value single-axis average).
  return nums.reduce((a, b) => a + b, 0);
}

/**
 * Build the EvalContext.resolvePivotData lookup over a set of placed pivots.
 * Given an anchor cell + a data field + (field,item) criteria it finds the pivot
 * whose output range contains the anchor and resolves the value, mirroring how
 * {@link makeTableRefResolver} wires structured references. Anchors outside every
 * placement -> "not-pivot" (#REF!); a matched pivot with no qualifying cell ->
 * "no-match" (#N/A). The host (SpreadsheetWindow) constructs the placements from
 * its on-sheet pivot output ranges and passes this into each formula's context.
 */
export function makePivotDataResolver(
  placements: readonly PivotPlacement[],
): (
  anchor: { row: number; col: number },
  dataField: string,
  criteria: readonly PivotDataCriterion[],
) => PivotDataResult {
  return (anchor, dataField, criteria) => {
    const hit = placements.find((p) =>
      rangeContains(p.output, anchor.row, anchor.col),
    );
    if (!hit) return { kind: "not-pivot" };
    return lookupInPlacement(hit, dataField, criteria);
  };
}
