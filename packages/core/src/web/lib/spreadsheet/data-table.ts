/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for STRUCTURED DATA TABLES (Excel ListObject / WPS "Format as
 * Table" parity).
 *
 * A data table is a NAMED structured range over existing grid cells: its first
 * row is either a visible header or data when Header Row is off; stable column
 * ids/names remain table metadata in both modes. An optional totals row appended
 * below the range owns real worksheet cells. The table's
 * style remains an overlay, but its header/data/totals values always live in the
 * ordinary grid cell map. Totals recipes produce real structured-reference
 * `SUBTOTAL` formulas, so evaluation, Undo, persistence, reload and XLSX export
 * all consume the same content authority.
 *
 * This module is dependency-light and deterministic so it can be unit-tested in
 * isolation: it carries the table MODEL, the build-from-range default, the
 * per-column aggregate, the cell-role mapping the renderer keys its style by,
 * and the normalize/persist coercion (a persisted value can come from an import,
 * an AI write, or a hand-edited JSON, so every field is clamped/dropped rather
 * than trusted). The structural remap (the range shifts on row/col insert/delete)
 * lives in structural.ts alongside validations — a table is a `StructuralRanged`.
 */

import { spreadsheetDefaultNames } from "./default-names";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import {
  aggregate,
  coerceNumber,
  type PivotAgg,
} from "./pivot";
import {
  normalizeConditionFilter,
  type ConditionFilter,
} from "./filter-conditions";
import {
  normalizeFilterDateGroups,
  type FilterDateGroup,
} from "./date-filter-tree";
import {
  iconSetCount,
  type IconSetName,
} from "./conditional-format";
import {
  paletteById,
} from "./color-palette";

import { builtInTableCellPaint, tableRecipeCellPaint } from "./table-style-recipes";
import { normalizeTableOwnBorders, type TableOwnBorders } from "./table-own-borders";
import {
  findNativeTableStyle,
  type SpreadsheetNativeTableStyle,
} from "./native-table-styles";

/**
 * Per-column totals aggregation. "none" means the totals cell is blank (Excel's
 * default for non-numeric columns); the other five reuse the pivot aggregators
 * over the column's data-row strings. (Pivot's `count` maps the same — it counts
 * non-blank entries, COUNTA semantics.)
 */
export type TableTotalKind =
  | "none"
  // A total row offers only these five of the PivotTable's functions (Excel
  // has no Product / StdDevp / Varp on a table's Total Row; XLS-A071).
  | Extract<PivotAgg, "sum" | "count" | "average" | "min" | "max">
  | "countNums"
  | "stdDev"
  | "var"
  | "custom";

export const TABLE_TOTAL_KINDS: readonly TableTotalKind[] = [
  "none",
  "average",
  "count",
  "countNums",
  "max",
  "min",
  "sum",
  "stdDev",
  "var",
  "custom",
] as const;

const TABLE_SUBTOTAL_CODES: Readonly<
  Record<Exclude<TableTotalKind, "none" | "custom">, number>
> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};

/**
 * Excel table-column names are case-insensitively unique and never blank.
 * Left to right, a later duplicate takes the next free number ("Amount2").
 * `keep` marks columns that already own their name (a table's existing
 * columns during a resize): they claim it first, so a column the resize adds
 * with the same header text is the one that gets renumbered, as in Excel.
 */
export function normalizeTableColumnNames(
  rawNames: readonly string[],
  options: { keep?: readonly boolean[] } = {},
): string[] {
  const used = new Set<string>();
  const out: string[] = new Array(rawNames.length);
  const claim = (index: number) => {
    const base = rawNames[index].trim().slice(0, 255) || `${spreadsheetDefaultNames().column}${index + 1}`;
    let candidate = base;
    let suffix = 2;
    while (used.has(candidate.toLowerCase())) {
      const marker = String(suffix);
      candidate = `${base.slice(0, Math.max(0, 255 - marker.length))}${marker}`;
      suffix += 1;
    }
    used.add(candidate.toLowerCase());
    out[index] = candidate;
  };
  const keep = options.keep;
  if (keep) {
    for (let index = 0; index < rawNames.length; index += 1) {
      if (keep[index]) claim(index);
    }
  }
  for (let index = 0; index < rawNames.length; index += 1) {
    if (!keep?.[index]) claim(index);
  }
  return out;
}

/**
 * The name Excel gives a table column that has no header text: "Column1",
 * "Column2", ... in English, "列1" in Chinese, taking the FIRST free number
 * that does not collide case-insensitively with a name the table already has
 * (Microsoft Support, "Resize a table by adding or removing rows and
 * columns"; the same rule Create Table applies to a blank header). `pattern`
 * renders the localized text for one number; a pattern that does not vary
 * with the number (a translator stub) falls back to Excel's English form so
 * the loop always ends on a fresh name.
 */
export function defaultTableColumnName(
  usedNames: Iterable<string>,
  pattern: (index: number) => string = (index) => `Column${index}`,
): string {
  const used = new Set<string>();
  for (const name of usedNames) used.add(name.trim().toLowerCase());
  const render =
    pattern(1).trim() !== pattern(2).trim()
      ? pattern
      : (index: number) => `Column${index}`;
  for (let index = 1; ; index += 1) {
    const candidate = render(index).trim().slice(0, 255);
    if (candidate !== "" && !used.has(candidate.toLowerCase())) {
      return candidate;
    }
  }
}

/**
 * The number format a table column's totals cell takes: Excel formats the
 * total like the column it totals, so "#,##0.00" data sums to "81,427.32"
 * (Microsoft Support, "Total the data in an Excel table"). The last data
 * cell's format wins, as an inserted row takes the row above; failing that,
 * the format most of the column's formatted data cells share. Null when no
 * data cell of the column carries a number format.
 */
export function tableColumnNumberFormat(
  cells: Readonly<
    Record<string, { style?: { numberFormat?: unknown } } | undefined>
  >,
  table: DataTable,
  offset: number,
): string | null {
  const col = table.range.c1 + offset;
  if (offset < 0 || col > table.range.c2) return null;
  const firstDataRow = table.range.r1 + (table.headerRow ? 1 : 0);
  const counts = new Map<string, number>();
  let lastRowFormat: string | null = null;
  let dominant: string | null = null;
  let dominantCount = 0;
  for (let row = firstDataRow; row <= table.range.r2; row += 1) {
    const format = cells[`${row}:${col}`]?.style?.numberFormat;
    if (typeof format !== "string" || format.trim() === "") continue;
    if (row === table.range.r2) lastRowFormat = format;
    const count = (counts.get(format) ?? 0) + 1;
    counts.set(format, count);
    if (count > dominantCount) {
      dominant = format;
      dominantCount = count;
    }
  }
  return lastRowFormat ?? dominant;
}

export interface DataTableColumn {
  /** Stable native ListObject tableColumn identity. */
  id: number;
  /** Canonical structured-reference name, independent of header visibility. */
  name: string;
}

/**
 * One table-owned AutoFilter criterion, keyed by the stable native
 * `tableColumn/@id` rather than by a worksheet column coordinate. Excel treats
 * the checkbox value set and a typed condition as mutually exclusive modes.
 */
export interface DataTableColumnFilter {
  values?: string[];
  dateGroups?: FilterDateGroup[];
  /** Explicit native calendar identity; omission is the Gregorian default. */
  calendarType?: "gregorian";
  condition?: ConditionFilter;
  /** Native table `<colorFilter>`; one exact fill or font colour per column. */
  color?: DataTableColorFilter;
  /** Native table `<iconFilter>`; one displayed icon in one standard family. */
  icon?: DataTableIconFilter;
  /** Base-OOXML dynamicFilter date period; val/maxVal are native caches only. */
  relativeDate?: DataTableRelativeDateFilter;
}

export interface DataTableRelativeDateFilter {
  period: DataTableRelativeDatePeriod;
  /** Native threshold cache preserved for an imported fixed point. */
  val?: number;
  /** Optional upper threshold cache; valid only together with val. */
  maxVal?: number;
}

/** The base-OOXML ST_DynamicFilterType date modes this table model can own. */
export const DATA_TABLE_RELATIVE_DATE_PERIODS = [
  "yesterday",
  "today",
  "tomorrow",
  "lastWeek",
  "thisWeek",
  "nextWeek",
  "lastMonth",
  "thisMonth",
  "nextMonth",
] as const;
export type DataTableRelativeDatePeriod =
  (typeof DATA_TABLE_RELATIVE_DATE_PERIODS)[number];

export function isDataTableRelativeDatePeriod(
  value: unknown,
): value is DataTableRelativeDatePeriod {
  return (
    typeof value === "string" &&
    (DATA_TABLE_RELATIVE_DATE_PERIODS as readonly string[]).includes(value)
  );
}

/** Strictly normalize the supported base-OOXML dynamic relative-date shape. */
export function normalizeDataTableRelativeDateFilter(
  raw: unknown,
): DataTableRelativeDateFilter | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const value = raw as Record<string, unknown>;
  const keys = Object.keys(value);
  if (
    !keys.every((key) => ["period", "val", "maxVal"].includes(key)) ||
    !keys.includes("period") ||
    !isDataTableRelativeDatePeriod(value.period) ||
    (value.val !== undefined &&
      (typeof value.val !== "number" || !Number.isFinite(value.val))) ||
    (value.maxVal !== undefined &&
      (typeof value.maxVal !== "number" || !Number.isFinite(value.maxVal))) ||
    (value.maxVal !== undefined && value.val === undefined) ||
    (typeof value.val === "number" &&
      typeof value.maxVal === "number" &&
      value.maxVal < value.val)
  ) {
    return null;
  }
  return {
    period: value.period,
    ...(typeof value.val === "number" ? { val: value.val } : {}),
    ...(typeof value.maxVal === "number" ? { maxVal: value.maxVal } : {}),
  };
}

export interface DataTableColorFilter {
  kind: "cell" | "font";
  value: string;
}

/**
 * The base SpreadsheetML icon families openpyxl and the standard `<iconFilter>`
 * element can round-trip without an x14/custom-icon extension. The three
 * extended families stay conditional-format display options, but are
 * deliberately excluded from table filter authority.
 */
export const DATA_TABLE_ICON_FILTER_SETS = [
  "3-arrows",
  "3-arrows-gray",
  "3-flags",
  "3-traffic-lights",
  "3-traffic-lights-rimmed",
  "3-signs",
  "3-symbols-circled",
  "3-symbols",
  "4-arrows",
  "4-arrows-gray",
  "4-red-to-black",
  "4-ratings",
  "4-traffic-lights",
  "5-arrows",
  "5-arrows-gray",
  "5-ratings",
  "5-quarters",
] as const satisfies readonly IconSetName[];

export type DataTableIconSetName = (typeof DATA_TABLE_ICON_FILTER_SETS)[number];

const DATA_TABLE_ICON_FILTER_SET = new Set<string>(
  DATA_TABLE_ICON_FILTER_SETS,
);

export interface DataTableIconFilter {
  iconSet: DataTableIconSetName;
  /** Zero-based icon identity inside the family, matching OOXML `iconId`. */
  iconId: number;
}

export function isDataTableIconSetName(
  raw: unknown,
): raw is DataTableIconSetName {
  return typeof raw === "string" && DATA_TABLE_ICON_FILTER_SET.has(raw);
}

/** Strictly normalize one base-OOXML table iconFilter descriptor. */
export function normalizeDataTableIconFilter(
  raw: unknown,
): DataTableIconFilter | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const value = raw as Record<string, unknown>;
  if (
    Object.keys(value).length !== 2 ||
    !Object.prototype.hasOwnProperty.call(value, "iconSet") ||
    !Object.prototype.hasOwnProperty.call(value, "iconId") ||
    !isDataTableIconSetName(value.iconSet) ||
    !Number.isInteger(value.iconId) ||
    Number(value.iconId) < 0 ||
    Number(value.iconId) >= iconSetCount(value.iconSet)
  ) {
    return null;
  }
  return { iconSet: value.iconSet, iconId: Number(value.iconId) };
}

/**
 * OOXML color filters point at one dxf colour, never a partially normalized
 * CSS token or a set of colours. Keep the editor representation just as strict
 * so import/export can preserve or refuse the native criterion atomically.
 */
export function normalizeDataTableFilterColor(
  raw: unknown,
): DataTableColorFilter | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const value = raw as Record<string, unknown>;
  if (
    Object.keys(value).length !== 2 ||
    !Object.prototype.hasOwnProperty.call(value, "kind") ||
    !Object.prototype.hasOwnProperty.call(value, "value") ||
    (value.kind !== "cell" && value.kind !== "font") ||
    typeof value.value !== "string" ||
    !/^#[0-9a-f]{6}$/i.test(value.value)
  ) {
    return null;
  }
  return { kind: value.kind, value: value.value.toLowerCase() };
}

/** One native table sort key, addressed by stable `tableColumn/@id`. */
export interface DataTableSortCondition {
  columnId: number;
  descending?: boolean;
  color?: DataTableColorFilter;
  icon?: DataTableIconFilter;
}

/**
 * The table-owned `<sortState>` behind Excel's header arrow and Reapply.
 * Its range is deliberately derived from the current table body, so a safe
 * resize cannot leave stale A1 coordinates in persisted editor state.
 */
export interface DataTableSortState {
  conditions: DataTableSortCondition[];
}

const MAX_TABLE_FILTER_VALUES = 10_000;
const MAX_TABLE_FILTER_VALUE_LENGTH = 1_000;

export function normalizeDataTableColumnFilter(
  raw: unknown,
): DataTableColumnFilter | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const value = raw as Record<string, unknown>;
  const keys = Object.keys(value);
  const hasValues = Object.prototype.hasOwnProperty.call(value, "values");
  const hasDateGroups = Object.prototype.hasOwnProperty.call(
    value,
    "dateGroups",
  );
  const hasCalendarType = Object.prototype.hasOwnProperty.call(
    value,
    "calendarType",
  );
  const hasCondition = Object.prototype.hasOwnProperty.call(value, "condition");
  const hasColor = Object.prototype.hasOwnProperty.call(value, "color");
  const hasIcon = Object.prototype.hasOwnProperty.call(value, "icon");
  const hasRelativeDate = Object.prototype.hasOwnProperty.call(
    value,
    "relativeDate",
  );
  if (
    !keys.every((key) =>
      [
        "values",
        "dateGroups",
        "calendarType",
        "condition",
        "color",
        "icon",
        "relativeDate",
      ].includes(key),
    ) ||
    (hasCondition &&
      (hasValues ||
        hasDateGroups ||
        hasCalendarType ||
        hasColor ||
        hasIcon ||
        hasRelativeDate)) ||
    (hasColor &&
      (hasValues ||
        hasDateGroups ||
        hasCalendarType ||
        hasCondition ||
        hasIcon ||
        hasRelativeDate)) ||
    (hasIcon &&
      (hasValues ||
        hasDateGroups ||
        hasCalendarType ||
        hasCondition ||
        hasColor ||
        hasRelativeDate)) ||
    (hasRelativeDate &&
      (hasValues ||
        hasDateGroups ||
        hasCalendarType ||
        hasCondition ||
        hasColor ||
        hasIcon)) ||
    (!hasCondition &&
      !hasColor &&
      !hasIcon &&
      !hasRelativeDate &&
      !hasValues &&
      !hasDateGroups)
  ) {
    return null;
  }

  if (hasCondition) {
    if (keys.length !== 1) return null;
    const condition = normalizeConditionFilter(value.condition);
    return condition ? { condition } : null;
  }

  if (hasColor) {
    if (keys.length !== 1) return null;
    const color = normalizeDataTableFilterColor(value.color);
    return color ? { color } : null;
  }

  if (hasIcon) {
    if (keys.length !== 1) return null;
    const icon = normalizeDataTableIconFilter(value.icon);
    return icon ? { icon } : null;
  }

  if (hasRelativeDate) {
    if (keys.length !== 1) return null;
    const relativeDate = normalizeDataTableRelativeDateFilter(
      value.relativeDate,
    );
    return relativeDate ? { relativeDate } : null;
  }

  let values: string[] | undefined;
  if (hasValues) {
    if (!Array.isArray(value.values)) return null;
    if (
      value.values.length === 0 ||
      value.values.length > MAX_TABLE_FILTER_VALUES
    ) {
      return null;
    }
    values = [];
    for (const entry of value.values) {
      if (
        typeof entry !== "string" ||
        entry.length > MAX_TABLE_FILTER_VALUE_LENGTH
      ) {
        return null;
      }
      values.push(entry);
    }
    if (new Set(values).size !== values.length) return null;
  }
  const dateGroups = hasDateGroups
    ? normalizeFilterDateGroups(value.dateGroups)
    : undefined;
  if (hasDateGroups && !dateGroups) return null;
  if (
    hasCalendarType &&
    (value.calendarType !== "gregorian" || !dateGroups)
  ) {
    return null;
  }
  return {
    ...(values ? { values } : {}),
    ...(dateGroups ? { dateGroups } : {}),
    ...(hasCalendarType ? { calendarType: "gregorian" as const } : {}),
  };
}

export function normalizeDataTableSortState(
  raw: unknown,
  columns: readonly DataTableColumn[],
): DataTableSortState | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const value = raw as Record<string, unknown>;
  if (Object.keys(value).length !== 1 || !Array.isArray(value.conditions)) {
    return null;
  }
  if (value.conditions.length < 1 || value.conditions.length > 64) return null;
  const knownIds = new Set(columns.map((column) => column.id));
  const seen = new Set<number>();
  const conditions: DataTableSortCondition[] = [];
  for (const rawCondition of value.conditions) {
    if (
      !rawCondition ||
      typeof rawCondition !== "object" ||
      Array.isArray(rawCondition)
    ) {
      return null;
    }
    const condition = rawCondition as Record<string, unknown>;
    const keys = Object.keys(condition);
    if (
      !keys.every(
        (key) =>
          key === "columnId" ||
          key === "descending" ||
          key === "color" ||
          key === "icon",
      ) ||
      !keys.includes("columnId") ||
      typeof condition.columnId !== "number" ||
      !Number.isInteger(condition.columnId) ||
      !knownIds.has(condition.columnId) ||
      seen.has(condition.columnId) ||
      (condition.descending !== undefined &&
        typeof condition.descending !== "boolean") ||
      (condition.color !== undefined &&
        normalizeDataTableFilterColor(condition.color) === null) ||
      (condition.icon !== undefined &&
        normalizeDataTableIconFilter(condition.icon) === null) ||
      (condition.color !== undefined && condition.icon !== undefined)
    ) {
      return null;
    }
    seen.add(condition.columnId);
    conditions.push({
      columnId: condition.columnId,
      ...(condition.descending === true ? { descending: true } : {}),
      ...(condition.color !== undefined
        ? { color: normalizeDataTableFilterColor(condition.color)! }
        : {}),
      ...(condition.icon !== undefined
        ? { icon: normalizeDataTableIconFilter(condition.icon)! }
        : {}),
    });
  }
  return { conditions };
}

/** Validate a complete left-to-right tableColumn identity set. */
export function normalizeDataTableColumns(
  raw: unknown,
  width: number,
): DataTableColumn[] | null {
  if (!Array.isArray(raw) || raw.length !== width || width <= 0) return null;
  const ids = new Set<number>();
  const names = new Set<string>();
  const columns: DataTableColumn[] = [];
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") return null;
    const value = entry as Record<string, unknown>;
    const id = Number(value.id);
    // The name as it is stored, a space at either end included: a table
    // column is named by its header cell's text, and Excel keeps "Note " as
    // it stands. Trimmed when the window opened the workbook, the column
    // read as renamed and the next save rewrote the table's part, its name
    
    // "All Disablement Pensioners ", R450-X05 "Note ", R454-X01). Names are
    // still told apart as a structured reference reads them, trimmed.
    const stored = typeof value.name === "string" ? value.name : "";
    const name = stored.trim() === "" ? "" : stored;
    const folded = name.trim().toLowerCase();
    if (
      !Number.isInteger(id) ||
      id <= 0 ||
      id > 2_147_483_647 ||
      ids.has(id) ||
      name === "" ||
      // As long as a header cell's text may be (a Google Forms export names
      // a column by its whole question: xlsx-rounds R96-X13, 414 chars).
      name.length > 32_767 ||
      names.has(folded)
    ) {
      return null;
    }
    ids.add(id);
    names.add(folded);
    columns.push({ id, name });
  }
  return columns;
}

/** Escape one literal table header for an Excel structured reference. */
function structuredColumnName(name: string): string {
  return name.replace(/[\[\]#'@:\u0000]/g, (character) => `'${character}`);
}

/** The real value owned by one physical totals-row cell. */
export function tableTotalCellValue(
  table: DataTable,
  offset: number,
  columnName: string,
): string {
  const kind = table.totals?.[offset] ?? "none";
  if (kind === "custom") return table.totalFormulas?.[offset] ?? "";
  if (kind !== "none") {
    return `=SUBTOTAL(${TABLE_SUBTOTAL_CODES[kind]},[${structuredColumnName(
      columnName,
    )}])`;
  }
  return table.totalLabels?.[offset] ?? "";
}

function gridColumnLabel(column: number): string {
  let value = column + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

/** Equivalent A1 formula used when Convert to Range removes table identity. */
export function tableTotalA1Formula(
  table: DataTable,
  offset: number,
): string | null {
  const kind = table.totals?.[offset] ?? "none";
  const firstDataIndex = table.range.r1 + (table.headerRow ? 1 : 0);
  if (
    kind === "none" ||
    kind === "custom" ||
    table.range.r2 < firstDataIndex
  ) {
    return null;
  }
  const column = gridColumnLabel(table.range.c1 + offset);
  const firstDataRow = table.range.r1 + (table.headerRow ? 2 : 1);
  return `=SUBTOTAL(${TABLE_SUBTOTAL_CODES[kind]},${column}${
    firstDataRow
  }:${column}${table.range.r2 + 1})`;
}

/**
 * Whether a physical totals cell still belongs to its persisted recipe.
 * Native imports may carry the equivalent A1 form, so both that and the Web
 * structured-reference form are accepted; an arbitrary user formula/value is
 * never silently cleared or replaced.
 */
export function tableTotalCellOwnsValue(
  table: DataTable,
  offset: number,
  columnName: string,
  currentValue: string,
): boolean {
  const expected = tableTotalCellValue(table, offset, columnName);
  if (currentValue === expected) return true;
  const kind = table.totals?.[offset] ?? "none";
  if (kind === "none") return false;
  const normalizeFormula = (formula: string) =>
    formula.replace(/\s+/g, "").toLowerCase();
  // What Excel writes in the file: the column qualified by the table's
  // name, =SUBTOTAL(109,Sales[Qty]).
  if (kind !== "custom" && table.name) {
    const qualified = `=SUBTOTAL(${TABLE_SUBTOTAL_CODES[kind]},${table.name}[${structuredColumnName(
      columnName,
    )}])`;
    if (normalizeFormula(currentValue) === normalizeFormula(qualified)) return true;
  }
  const a1Formula = tableTotalA1Formula(table, offset);
  if (!a1Formula) return false;
  return normalizeFormula(currentValue) === normalizeFormula(a1Formula);
}

/**
 * Curated native table-style presets. The ids are persisted in the workbook
 * model; their palette stays centralized so the ribbon swatches and grid
 * renderer cannot drift apart.
 */
const LEGACY_TABLE_STYLE_IDS = [
  "slate-light",
  "blue-light",
  "orange-light",
  "green-light",
  "teal-light",
  "purple-light",
  "lime-light",
  "slate-medium",
  "orange-medium",
  "purple-medium",
  "lime-medium",
  "slate-dark",
  "blue-dark",
  "orange-dark",
  "green-dark",
  "teal-dark",
  "purple-dark",
  "lime-dark",
  "teal-medium",
  "blue-medium",
  "green-medium",
] as const;

const ADDITIONAL_TABLE_STYLE_NATIVE_IDS = [
  ["light-8", "light", 8],
  ["light-9", "light", 9],
  ["light-10", "light", 10],
  ["light-11", "light", 11],
  ["light-12", "light", 12],
  ["light-13", "light", 13],
  ["light-14", "light", 14],
  ["light-15", "light", 15],
  ["light-16", "light", 16],
  ["light-17", "light", 17],
  ["light-18", "light", 18],
  ["light-19", "light", 19],
  ["light-20", "light", 20],
  ["light-21", "light", 21],
  ["medium-8", "medium", 8],
  ["medium-9", "medium", 9],
  ["medium-10", "medium", 10],
  ["medium-11", "medium", 11],
  ["medium-12", "medium", 12],
  ["medium-13", "medium", 13],
  ["medium-14", "medium", 14],
  ["medium-15", "medium", 15],
  ["medium-16", "medium", 16],
  ["medium-17", "medium", 17],
  ["medium-18", "medium", 18],
  ["medium-19", "medium", 19],
  ["medium-20", "medium", 20],
  ["medium-21", "medium", 21],
  ["medium-22", "medium", 22],
  ["medium-23", "medium", 23],
  ["medium-24", "medium", 24],
  ["medium-25", "medium", 25],
  ["medium-26", "medium", 26],
  ["medium-27", "medium", 27],
  ["medium-28", "medium", 28],
  ["dark-8", "dark", 8],
  ["dark-9", "dark", 9],
  ["dark-10", "dark", 10],
  ["dark-11", "dark", 11],
] as const;

export const TABLE_STYLE_IDS = [
  "none",
  ...LEGACY_TABLE_STYLE_IDS,
  ...ADDITIONAL_TABLE_STYLE_NATIVE_IDS.map(([id]) => id),
] as const;

export type TableStyleId = (typeof TABLE_STYLE_IDS)[number];

export interface TableStylePreset {
  id: TableStyleId;
  label: { en: string; zh: string };
  /** Excel's gallery section. Kept in the model so Home and Table Design agree. */
  tone: "light" | "medium" | "dark";
  /** Native OOXML built-in used by server export and XLSX import round-trips. */
  excelStyleName: string;
  headerBackground: string;
  headerColor: string;
  bandBackground: string;
  totalsBackground: string;
  accent: string;
}

// Persisted legacy ids are identities, not literal colour names.
const LEGACY_TABLE_STYLE_PRESETS: readonly TableStylePreset[] = [
  { id: "slate-light", tone: "light", excelStyleName: "TableStyleLight1" },
  { id: "blue-light", tone: "light", excelStyleName: "TableStyleLight2" },
  { id: "orange-light", tone: "light", excelStyleName: "TableStyleLight3" },
  { id: "green-light", tone: "light", excelStyleName: "TableStyleLight4" },
  { id: "teal-light", tone: "light", excelStyleName: "TableStyleLight5" },
  { id: "purple-light", tone: "light", excelStyleName: "TableStyleLight6" },
  { id: "lime-light", tone: "light", excelStyleName: "TableStyleLight7" },
  { id: "slate-medium", tone: "medium", excelStyleName: "TableStyleMedium1" },
  { id: "teal-medium", tone: "medium", excelStyleName: "TableStyleMedium5" },
  { id: "blue-medium", tone: "medium", excelStyleName: "TableStyleMedium2" },
  { id: "orange-medium", tone: "medium", excelStyleName: "TableStyleMedium3" },
  { id: "green-medium", tone: "medium", excelStyleName: "TableStyleMedium4" },
  { id: "purple-medium", tone: "medium", excelStyleName: "TableStyleMedium6" },
  { id: "lime-medium", tone: "medium", excelStyleName: "TableStyleMedium7" },
  { id: "slate-dark", tone: "dark", excelStyleName: "TableStyleDark1" },
  { id: "blue-dark", tone: "dark", excelStyleName: "TableStyleDark2" },
  { id: "orange-dark", tone: "dark", excelStyleName: "TableStyleDark3" },
  { id: "green-dark", tone: "dark", excelStyleName: "TableStyleDark4" },
  { id: "teal-dark", tone: "dark", excelStyleName: "TableStyleDark5" },
  { id: "purple-dark", tone: "dark", excelStyleName: "TableStyleDark6" },
  { id: "lime-dark", tone: "dark", excelStyleName: "TableStyleDark7" },
].map((identity) => ({
  ...identity,
  label: { en: "", zh: "" }, headerBackground: "#ffffff", headerColor: "#000000",
  bandBackground: "#ffffff", totalsBackground: "#ffffff", accent: "#000000",
})) as TableStylePreset[];

function nativeTableStyleOrdinal(preset: TableStylePreset): number {
  return Number.parseInt(preset.excelStyleName.match(/\d+$/)?.[0] ?? "0", 10);
}

const TABLE_STYLE_TONE_ORDER = { light: 0, medium: 1, dark: 2 };

// Office Web reference gallery: colour descriptions belong to each native
// recipe, not just its accent slot (e.g. Light 8 is White, Dark 3 is Brown).
const OFFICE_TABLE_STYLE_COLOR_NAMES = {
  light: [
    "浅灰色", "青绿", "浅橙色", "浅绿色", "青绿", "粉红", "浅绿色",
    "白色", "深青", "橙色", "深绿色", "青绿", "梅红", "绿色",
    "浅灰色", "青绿", "浅橙色", "浅绿色", "青绿", "粉红", "浅绿色",
  ],
  medium: [
    "浅灰色", "深青", "橙色", "深绿色", "青绿", "梅红", "绿色",
    "灰色", "深青", "橙色", "深绿色", "青绿", "梅红", "绿色",
    "浅灰色", "深青", "橙色", "深绿色", "青绿", "梅红", "绿色",
    "灰色", "青绿", "浅橙色", "浅绿色", "青绿", "粉红", "浅绿色",
  ],
  dark: [
    "深灰色", "深青", "褐色", "深绿色", "青绿", "梅红", "深绿色",
    "灰色", "橙色", "青绿", "绿色",
  ],
} as const;
const TABLE_STYLE_COLOR_EN: Record<string, string> = {
  浅灰色: "Light Gray", 青绿: "Turquoise", 浅橙色: "Light Orange",
  浅绿色: "Light Green", 粉红: "Pink", 白色: "White", 深青: "Dark Teal",
  橙色: "Orange", 深绿色: "Dark Green", 梅红: "Magenta", 绿色: "Green",
  灰色: "Gray", 深灰色: "Dark Gray", 褐色: "Brown", 主题色: "Theme Color",
};

/** Read-only display metadata shared with the catalogue adapter; never a style writer. */
export function tableStyleLabelMetadata(base: TableStylePreset, accents: readonly string[] = paletteById("office").accents) {
  const ordinal = nativeTableStyleOrdinal(base);
  const family = base.tone === "dark" && ordinal > 7 ? (ordinal - 8) * 2 : (ordinal - 1) % 7;
  const accent = family > 0 ? accents[family - 1] : "#000000";
  const officeAccents = paletteById("office").accents;
  const isOffice = accents.length === officeAccents.length && accents.every(
    (value, index) => value.toLowerCase() === officeAccents[index],
  );
  const colorZh = ((isOffice || family === 0) && OFFICE_TABLE_STYLE_COLOR_NAMES[base.tone][ordinal - 1]) || "主题色";
  const tone = base.tone === "light" ? ["Light", "浅色"] : base.tone === "medium" ? ["Medium", "中等"] : ["Dark", "深色"];
  return { ordinal, colorZh, color: TABLE_STYLE_COLOR_EN[colorZh], tone, accent };
}

function resolveTableStylePreset(base: TableStylePreset, accents: readonly string[]): TableStylePreset {
  const look = { range: { r1: 0, c1: 0, r2: 5, c2: 4 }, headerRow: true, bandedRows: true, totalsRow: true };
  const paint = (row: number) => builtInTableCellPaint(base.excelStyleName, accents, look, row, 1);
  const header = paint(0), band = paint(1), totals = paint(6);
  const { ordinal, colorZh, tone, accent } = tableStyleLabelMetadata(base, accents);
  return {
    ...base,
    label: base.id === "none" ? { en: "No Style", zh: "无" } : {
      en: TABLE_STYLE_COLOR_EN[colorZh] + ", Table Style " + tone[0] + " " + ordinal,
      zh: colorZh + "，表样式 " + tone[1] + " " + ordinal,
    },
    headerBackground: header.backgroundColor ?? "#ffffff",
    headerColor: header.color ?? "#000000",
    bandBackground: band.backgroundColor ?? "#ffffff",
    totalsBackground: totals.backgroundColor ?? "#ffffff",
    accent: accent ?? "#000000",
  };
}

/** Native built-ins, ordered like Excel Web. None clears the style, not the table. */
export const TABLE_STYLE_PRESETS: readonly TableStylePreset[] = [
  { id: "none" as const, label: { en: "No Style", zh: "无样式" }, tone: "light" as const,
    excelStyleName: "", headerBackground: "#ffffff", headerColor: "#000000",
    bandBackground: "#ffffff", totalsBackground: "#ffffff", accent: "#000000" },
  ...LEGACY_TABLE_STYLE_PRESETS,
  ...ADDITIONAL_TABLE_STYLE_NATIVE_IDS.map(([id, tone, ordinal]) => ({
    ...LEGACY_TABLE_STYLE_PRESETS[0], id, tone,
    excelStyleName: "TableStyle" + tone[0].toUpperCase() + tone.slice(1) + ordinal,
  })),
].sort((a, b) => TABLE_STYLE_TONE_ORDER[a.tone] - TABLE_STYLE_TONE_ORDER[b.tone] || nativeTableStyleOrdinal(a) - nativeTableStyleOrdinal(b))
  .map((preset) => resolveTableStylePreset(preset, paletteById("office").accents));

const TABLE_STYLE_BY_EXCEL_NAME = new Map(TABLE_STYLE_PRESETS.map((preset) => [preset.excelStyleName.toLowerCase(), preset.id]));

export function tableStylePreset(styleId: TableStyleId | undefined): TableStylePreset {
  return TABLE_STYLE_PRESETS.find((preset) => preset.id === styleId)
    ?? TABLE_STYLE_PRESETS.find((preset) => preset.id === "teal-medium")!;
}

export function tableStylePresetForTheme(styleId: TableStyleId | undefined, accents: readonly string[]): TableStylePreset {
  const base = tableStylePreset(styleId);
  return styleId === undefined ? base : resolveTableStylePreset(base, accents);
}

/**
 * A Table Styles gallery choice: a built-in gallery id, or one of the
 * workbook's own styles (the gallery's Custom section) by name.
 */
export type NativeTableStyleChoice = `native:${string}`;
export type TableStyleChoice = TableStyleId | NativeTableStyleChoice;

/**
 * The style a new table takes: Excel's default table style, Table Style
 * Medium 2 (Accent 1; every Excel styles.xml names it defaultTableStyle and
 * Excel for Mac 16.86 gives a new table it, XLS-A109). The legacy fallbacks
 * for a stored style this build does not know stay Medium 5.
 */
export const DEFAULT_NEW_TABLE_STYLE_ID: TableStyleId = "blue-medium";

export function nativeTableStyleChoice(name: string): NativeTableStyleChoice {
  return `native:${name}`;
}

/** The workbook style name a choice names, or undefined for a built-in. */
export function nativeTableStyleChoiceName(
  choice: TableStyleChoice | undefined,
): string | undefined {
  return typeof choice === "string" && choice.startsWith("native:")
    ? choice.slice("native:".length)
    : undefined;
}

/**
 * The workbook style a table paints with: the one its `nativeStyleName`
 * names when that is not a built-in Excel name, or the one a gallery choice
 * names. Undefined when it uses a built-in (or a style the page lacks).
 */
export function tableNativeStyle(
  table: Pick<DataTable, "styleId" | "nativeStyleName">,
  nativeStyles: readonly SpreadsheetNativeTableStyle[] | undefined,
  choice?: TableStyleChoice,
): SpreadsheetNativeTableStyle | undefined {
  const chosen = nativeTableStyleChoiceName(choice);
  if (chosen !== undefined) return findNativeTableStyle(nativeStyles, chosen);
  if (choice !== undefined || table.styleId !== undefined) return undefined;
  if (!table.nativeStyleName || tableStyleIdFromExcelName(table.nativeStyleName)) {
    return undefined;
  }
  return findNativeTableStyle(nativeStyles, table.nativeStyleName);
}

/**
 * A table with a gallery choice applied: a built-in id with its Excel name,
 * or a workbook style by name (no gallery id; the name is what the file
 * keeps in tableStyleInfo@name).
 */
export function tableWithStyleChoice(
  table: DataTable,
  choice: TableStyleChoice,
): DataTable {
  const name = nativeTableStyleChoiceName(choice);
  if (name !== undefined) {
    const next: DataTable = { ...table, nativeStyleName: name };
    delete next.styleId;
    return next;
  }
  const styleId = choice as TableStyleId;
  const excelStyleName = tableStylePreset(styleId).excelStyleName;
  // None has no Excel name: the table states none (an empty name failed the
  // page check, so None could not be picked from a command).
  if (!excelStyleName) {
    const { nativeStyleName: _previous, ...rest } = table;
    return { ...rest, styleId };
  }
  return { ...table, styleId, nativeStyleName: excelStyleName };
}

/** Whether a gallery choice names a style this workbook can use. */
export function tableStyleChoiceAvailable(
  choice: TableStyleChoice,
  nativeStyles: readonly SpreadsheetNativeTableStyle[] | undefined,
): boolean {
  const name = nativeTableStyleChoiceName(choice);
  return name !== undefined
    ? findNativeTableStyle(nativeStyles, name) !== undefined
    : (TABLE_STYLE_IDS as readonly string[]).includes(choice);
}

/** The gallery choice a table currently shows as selected. */
export function tableStyleChoiceOf(
  table: Pick<DataTable, "styleId" | "nativeStyleName">,
  nativeStyles: readonly SpreadsheetNativeTableStyle[] | undefined,
): TableStyleChoice {
  const native = tableNativeStyle(table, nativeStyles);
  return native ? nativeTableStyleChoice(native.name) : table.styleId ?? "teal-medium";
}

/**
 * Shared paint authority for grid, pinned header, conversion and gallery
 * tiles. `styleId` is the built-in (or previewed) choice; a table on one of
 * the workbook's own styles paints from that style's elements.
 */
export function tableStyleCellPaint(
  styleId: TableStyleChoice | undefined,
  accents: readonly string[],
  table: DataTable,
  row: number,
  col: number,
  nativeStyles?: readonly SpreadsheetNativeTableStyle[],
) {
  const native = tableNativeStyle(table, nativeStyles, styleId);
  if (native) return tableRecipeCellPaint(native, accents, table, row, col);
  const builtIn = nativeTableStyleChoiceName(styleId) === undefined
    ? (styleId as TableStyleId | undefined)
    : undefined;
  return builtInTableCellPaint(tableStylePreset(builtIn).excelStyleName, accents, table, row, col);
}

/** Map a native Excel built-in table style back to the editor gallery id. */
export function tableStyleIdFromExcelName(
  value: unknown,
): TableStyleId | undefined {
  if (typeof value !== "string") return undefined;
  return TABLE_STYLE_BY_EXCEL_NAME.get(value.trim().toLowerCase());
}

/** The role a cell inside (or just below) a table range plays, for styling. */
export type TableCellRole = "header" | "data" | "banded" | "totals" | null;

/** Native ListObject state that makes an in-place range resize lossy. */
export type TableNativeResizeBlocker =
  | "totals"
  | "activeFilter"
  | "sortState"
  | "query"
  | "extension";

/**
 * The table's own AutoFilter / sort holds criteria MRAI does not model
 * ("activeFilter": a column criterion that may own hidden rows; "sortState":
 * only sort or button metadata). The in-place XLSX writer carries them exactly
 * through every shape change - resize, move, rename, restyle, Total Row, a
 * header rename - and drops them when the Header Row is turned off, as Excel
 * does. Only this table's filter and sort commands stay closed.
 */
export function tableNativeFilterCustody(
  table: Pick<DataTable, "nativeResizeBlocker">,
): boolean {
  return (
    table.nativeResizeBlocker === "activeFilter" ||
    table.nativeResizeBlocker === "sortState"
  );
}

/** Native table data that keeps the table's structure closed. */
export function tableNativeStructureBlocked(
  table: Pick<DataTable, "nativeResizeBlocker">,
): boolean {
  return Boolean(table.nativeResizeBlocker) && !tableNativeFilterCustody(table);
}

const TABLE_NATIVE_RESIZE_BLOCKERS = new Set<TableNativeResizeBlocker>([
  "totals",
  "activeFilter",
  "sortState",
  "query",
  "extension",
]);

/**
 * A structured data table over a rectangular grid range. `range` (r1/c1/r2/c2,
 * inclusive, 0-based) covers the optional HEADER row plus DATA rows.
 * `columns` owns stable ListObject ids/names even while `headerRow` is false;
 * legacy persisted tables may omit it and are resolved from visible headers.
 * `bandedRows` toggles the
 * alternating data-row stripe (on by default). `totalsRow` owns the appended
 * physical row at `range.r2 + 1`; `totals` is the per-column aggregation keyed
 * by the column's offset from `range.c1` (0-based, in-range). A missing/`"none"`
 * entry leaves that column's totals cell blank unless `totalLabels` owns it.
 */
export interface DataTable {
  id: string;
  name: string;
  range: { r1: number; c1: number; r2: number; c2: number };
  headerRow: boolean;
  columns?: DataTableColumn[];
  /** Header disclosure visibility. Omitted is the Excel default (`true`). */
  showFilterButton?: boolean;
  /**
   * Table-local criteria keyed by stable `tableColumn.id`. Hiding the buttons
   * deliberately preserves this map; showing them again restores the criteria.
   */
  columnFilters?: Record<string, DataTableColumnFilter>;
  /** Supported value-sort keys in precedence order, owned by this table. */
  sortState?: DataTableSortState;
  styleId?: TableStyleId;
  /** Imported Excel style name when it is outside the curated gallery. */
  nativeStyleName?: string;
  /** Office 2010+ x14 table alternate-text description. */
  altText?: string;
  /** Office 2010+ x14 table alternate-text summary (the UI's short title). */
  altTextSummary?: string;
  /** The table's own borders (tableBorderDxfId, headerRowBorderDxfId,
   *  totalsRowBorderDxfId), drawn over the cells' own edges (table-own-borders). */
  ownBorders?: TableOwnBorders;
  /**
   * Imported native state the bounded ListObject writer deliberately preserves.
   * Commands that would change `range` must fail before mutating the grid when
   * this is present; the low-level package writer repeats the same check.
   */
  nativeResizeBlocker?: TableNativeResizeBlocker;
  /**
   * With `nativeResizeBlocker` "activeFilter" / "sortState": the stable
   * column ids the table's own (unmodelled) filter criteria and sort keys sit
   * on. A reshape that drops one of these columns is refused, the same rule
   * modelled `columnFilters` follow.
   */
  nativeFilterColumnIds?: number[];
  /**
   * The table is the landing area of a legacy external-data query (text
   * import, web query, deleted connection). Excel lets its rows, names,
   * style, filters and sorts change; its columns belong to the query.
   */
  queryBacked?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  totalsRow?: boolean;
  totals?: Record<number, TableTotalKind>;
  /** Canonical first-data-row formula for each native calculated table column. */
  calculatedFormulas?: Record<number, string>;
  /** Calculated columns the file marks `array="1"`, by column offset: each
   *  row's cell is its own array formula, and the page holds no recipe. */
  arrayCalculatedColumns?: Record<number, true>;
  /**
   * Exact user-authored formulas for columns whose totals kind is `custom`.
   * Values use the editor's normal leading-`=` form; XLSX adapters remove/add
   * that marker only at the OOXML boundary (`totalsRowFormula` omits it).
   */
  totalFormulas?: Record<number, string>;
  /**
   * Excel `totalsRowLabel` per in-range column OFFSET — the literal text the
   * totals row shows in that column (typically "Total" in the first column)
   * instead of an aggregate. A column with both a label and an aggregate kind
   * keeps the aggregate; Excel itself never persists both.
   */
  totalLabels?: Record<number, string>;
}

/** Max characters honored for one totals-row label (matches Excel's cap). */
export const MAX_TOTALS_LABEL_LENGTH = 255;
/** Excel's worksheet formula text limit, including the editor's leading `=`. */
export const MAX_TABLE_TOTAL_FORMULA_LENGTH = 8_192;
const XML_ILLEGAL_CONTROL_RE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/;

/** True when a custom total is bounded and safe for worksheet XML. */
export function isAuthorableTableTotalFormula(value: unknown): value is string {
  if (typeof value !== "string") return false;
  const formula = value.trim();
  return (
    formula.startsWith("=") &&
    formula.length > 1 &&
    formula.length <= MAX_TABLE_TOTAL_FORMULA_LENGTH &&
    !XML_ILLEGAL_CONTROL_RE.test(formula)
  );
}
/** SpreadsheetML x14 table alternate-text bounds (MS-XLSX CT_Table). */
export const MAX_TABLE_ALT_TEXT_LENGTH = 25_000;
export const MAX_TABLE_ALT_TEXT_SUMMARY_LENGTH = 50_000;

/** Max data tables a sheet may carry (DoS bound on iteration / overlay render).
 *  Excel sets no count of its own; workbooks with 68 and 123 tables on one
 *  sheet are in the corpus, and a table past the bound is dropped from the page
 *  while its file part stays where it was: a row inserted above it moved its
 *  cells and left the table behind (xlsx-rounds R312-X25: Table65-68 of 68). */
export const MAX_DATA_TABLES = 1024;
/** Max length of a table id / name (DoS bound on a hand-edited value). */
export const MAX_DATA_TABLE_ID_LENGTH = 200;
export const MAX_DATA_TABLE_NAME_LENGTH = 200;
// Grid bounds mirror the grid's own (grid-sizing) so a normalized range can
// never address outside it. The old 10,000 x 1,000 grid's bounds clamped a
// table over every column (A1:XFD272) to column 999: its columns no longer

const MAX_RANGE_ROW = MAX_GRID_ROWS - 1;
const MAX_RANGE_COL = MAX_GRID_COLUMNS - 1;

let tableSeq = 0;

function normalizeTableAltText(value: unknown, maxLength: number): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .trim()
    .slice(0, maxLength);
}

/**
 * Excel/WPS-compatible table names: identifier-like, workbook-scoped, and not
 * ambiguous with A1/R1C1 cell references.
 */
export function isValidDataTableName(name: string): boolean {
  const trimmed = name.trim();
  if (
    trimmed === "" ||
    trimmed.length > MAX_DATA_TABLE_NAME_LENGTH ||
    // Any letter of any script starts a name (销售, ÉVentes), as in Excel;
    // digits and periods may follow (live audit B106).
    !/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(trimmed)
  ) {
    return false;
  }
  if (/^[RC]$/i.test(trimmed)) return false;
  if (/^[A-Za-z]{1,3}[1-9][0-9]*$/.test(trimmed)) return false;
  if (/^R[1-9][0-9]*C[1-9][0-9]*$/i.test(trimmed)) return false;
  return true;
}

/**
 * Rename a table in one formula, as Excel does when a table is renamed: every
 * qualified reference (`Sales[Amount]`, `Sales[#All]`) and every bare
 * whole-table reference (`=SUM(Sales)`) follows the new name. Quoted strings,
 * quoted sheet names, sheet prefixes (`Sales!A1`), function calls, external
 * workbook brackets and similarly prefixed names are never touched.
 */
export function renameDataTableReferences(
  formula: string,
  oldName: string,
  newName: string,
): string {
  if (!formula.startsWith("=") || oldName === "" || oldName === newName) {
    return formula;
  }
  return rewriteStructuredReferences(formula, {
    tableName: oldName,
    nextTableName: newName,
  });
}









export function breakDataTableReferences(formula: string, tableName: string): string {
  if (!formula.startsWith("=") || tableName === "") return formula;
  return rewriteStructuredReferences(formula, { tableName, broken: true });
}

/** Find the closing bracket of one possibly nested structured-reference token. */
function structuredReferenceBracketEnd(
  source: string,
  start: number,
): number {
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    const character = source[index];
    // Excel escapes a structural character inside a column name with `'`.
    // Skip the escaped character so e.g. `Cost']Net` cannot close the token.
    if (character === "'" && index + 1 < source.length) {
      index += 1;
      continue;
    }
    if (character === "[") depth += 1;
    else if (character === "]") {
      depth -= 1;
      if (depth === 0) return index;
      if (depth < 0) return -1;
    }
  }
  return -1;
}

/** Rewrite exact leaf column tokens inside one structured-reference selector. */
function renameStructuredColumnSelector(
  selector: string,
  oldColumn: string,
  newColumn: string,
): string {
  const oldLower = oldColumn.toLowerCase();
  const renameLeaf = (value: string): string => {
    const rowPrefix = value.startsWith("@") ? "@" : "";
    const candidate = rowPrefix ? value.slice(1) : value;
    return candidate.toLowerCase() === oldLower
      ? `${rowPrefix}${newColumn}`
      : value;
  };

  let out = "";
  let index = 0;
  while (index < selector.length) {
    if (selector[index] !== "[") {
      out += selector[index];
      index += 1;
      continue;
    }
    const end = structuredReferenceBracketEnd(selector, index);
    if (end < 0) return selector;
    const inner = selector.slice(index + 1, end);
    const renamedLeaf = renameLeaf(inner);
    out += `[${
      renamedLeaf === inner
        ? renameStructuredColumnSelector(inner, oldColumn, newColumn)
        : renamedLeaf
    }]`;
    index = end + 1;
  }
  return out;
}

/**
 * Rename one table column in the structured references of one formula.
 * Matching is case-insensitive like Excel table identities, but exact:
 * similarly named tables/columns survive. Quoted strings are never inspected,
 * while simple, row-scoped, range and nested `[#Data]` selectors all share the
 * same bracket-aware path.
 *
 * Explicit `Table[Column]` references are renamed wherever the formula sits.
 * The unqualified in-table forms (`[@Column]`, `[Column]`) address the table
 * that CONTAINS the formula, so they are renamed only when the caller says the
 * formula belongs to that table (`shortFormOwnedByTable`): a cell inside the
 * table's range, or one of its calculated-column / totals-row formulas.
 */
export function renameDataTableColumnReferences(
  formula: string,
  tableName: string,
  oldColumnName: string,
  newColumnName: string,
  options: { shortFormOwnedByTable?: boolean } = {},
): string {
  if (
    !formula.startsWith("=") ||
    tableName === "" ||
    oldColumnName === "" ||
    oldColumnName === newColumnName
  ) {
    return formula;
  }
  return rewriteStructuredReferences(formula, {
    tableName,
    column: { from: oldColumnName, to: newColumnName },
    shortFormOwnedByTable: options.shortFormOwnedByTable === true,
  });
}

/**
 * Rename a column inside a formula owned by that table (its Totals Row or a
 * calculated-column template). In this context Excel permits the short
 * `[Column]` / `[@Column]` forms as well as `Table[Column]`.
 */
export function renameDataTableTotalFormulaColumn(
  formula: string,
  tableName: string,
  oldColumnName: string,
  newColumnName: string,
): string {
  return renameDataTableColumnReferences(
    formula,
    tableName,
    oldColumnName,
    newColumnName,
    { shortFormOwnedByTable: true },
  );
}

// Any script's letters and digits, as table names take them (销售, 新Sales):
// an ASCII-only class split 新Sales into 新 + Sales (O008).
const FORMULA_NAME_CHARACTER = /[\p{L}\p{N}_.\\]/u;

/** Index just past a quoted run (`"..."` or `'...'`, doubled quote escapes). */
function quotedRunEnd(source: string, start: number): number {
  const quote = source[start];
  let index = start + 1;
  while (index < source.length) {
    if (source[index] === quote) {
      if (source[index + 1] === quote) {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return source.length;
}

/**
 * The one token walk behind every table/column rename. It reads a formula as
 * a sequence of string literals, quoted sheet names, identifiers (a table name
 * possibly followed by its `[...]` selector) and unqualified `[...]` selectors,
 * and rewrites only the tokens that address `tableName`.
 */
function rewriteStructuredReferences(
  formula: string,
  plan: {
    tableName: string;
    nextTableName?: string;
    column?: { from: string; to: string };
    shortFormOwnedByTable?: boolean;
    /** Each reference to the table becomes `#REF!` (breakDataTableReferences). */
    broken?: boolean;
  },
): string {
  const tableLower = plan.tableName.toLowerCase();
  const oldColumn = plan.column ? structuredColumnName(plan.column.from) : "";
  const newColumn = plan.column ? structuredColumnName(plan.column.to) : "";
  const renameSelector = (selector: string) =>
    oldColumn && oldColumn !== newColumn
      ? renameStructuredColumnSelector(selector, oldColumn, newColumn)
      : selector;
  let out = "";
  let index = 0;
  while (index < formula.length) {
    const character = formula[index];
    if (character === '"' || character === "'") {
      const end = quotedRunEnd(formula, index);
      out += formula.slice(index, end);
      index = end;
      continue;
    }
    if (character === "[") {
      const end = structuredReferenceBracketEnd(formula, index);
      if (end < 0) {
        out += formula.slice(index);
        break;
      }
      const selector = formula.slice(index, end + 1);
      // `[Book.xlsx]Sheet1!A1` is an external-workbook prefix, not a column.
      const after = formula[end + 1];
      const externalWorkbook = Boolean(after && /[A-Za-z0-9_'!]/.test(after));
      out +=
        plan.shortFormOwnedByTable && !externalWorkbook
          ? renameSelector(selector)
          : selector;
      index = end + 1;
      continue;
    }
    if (FORMULA_NAME_CHARACTER.test(character)) {
      let end = index;
      while (end < formula.length && FORMULA_NAME_CHARACTER.test(formula[end])) {
        end += 1;
      }
      const word = formula.slice(index, end);
      const next = formula[end];
      const matches =
        word.toLowerCase() === tableLower && formula[index - 1] !== "!";
      if (next === "[") {
        const selectorEnd = structuredReferenceBracketEnd(formula, end);
        if (selectorEnd >= 0) {
          const selector = formula.slice(end, selectorEnd + 1);
          out += matches
            ? plan.broken
              ? "#REF!"
              : `${plan.nextTableName ?? word}${renameSelector(selector)}`
            : `${word}${selector}`;
          index = selectorEnd + 1;
          continue;
        }
      }
      // A bare table name is the whole data body (`=SUM(Sales)`); a sheet
      // prefix (`Sales!A1`) or a function call (`Sales(`) is something else.
      out +=
        matches && next !== "(" && next !== "!" && (plan.nextTableName || plan.broken)
          ? plan.broken
            ? "#REF!"
            : plan.nextTableName
          : word;
      index = end;
      continue;
    }
    out += character;
    index += 1;
  }
  return out;
}

/**
 * Turn a custom Totals Row reference to a deleted table column into `#REF!`.
 * Exact short and explicit selectors are handled outside string literals;
 * similarly named tables/columns and external-workbook brackets survive.
 */
export function invalidateDataTableTotalFormulaColumn(
  formula: string,
  tableName: string,
  deletedColumnName: string,
): string {
  if (!formula.startsWith("=") || !deletedColumnName) return formula;
  const deletedColumn = structuredColumnName(deletedColumnName);
  const replacementColumn = structuredColumnName(
    "__MRAI_DELETED_TABLE_COLUMN__",
  );
  const selectorTouchesDeletedColumn = (selector: string) =>
    renameStructuredColumnSelector(
      selector,
      deletedColumn,
      replacementColumn,
    ) !== selector;
  const tableLower = tableName.toLowerCase();
  let out = "";
  let inString = false;
  let index = 0;
  while (index < formula.length) {
    const character = formula[index];
    if (character === '"') {
      if (inString && formula[index + 1] === '"') {
        out += '""';
        index += 2;
        continue;
      }
      inString = !inString;
      out += character;
      index += 1;
      continue;
    }
    if (inString) {
      out += character;
      index += 1;
      continue;
    }
    const previous = index > 0 ? formula[index - 1] : "";
    const previousIsNameCharacter = /[A-Za-z0-9_.\\]/.test(previous);
    const selectorStart = index + tableName.length;
    if (
      tableName &&
      !previousIsNameCharacter &&
      formula.slice(index, selectorStart).toLowerCase() === tableLower &&
      formula[selectorStart] === "["
    ) {
      const selectorEnd = structuredReferenceBracketEnd(
        formula,
        selectorStart,
      );
      if (selectorEnd >= 0) {
        const selector = formula.slice(selectorStart, selectorEnd + 1);
        out += selectorTouchesDeletedColumn(selector)
          ? "#REF!"
          : formula.slice(index, selectorEnd + 1);
        index = selectorEnd + 1;
        continue;
      }
    }
    if (character === "[" && !previousIsNameCharacter) {
      let selectorEnd = structuredReferenceBracketEnd(formula, index);
      if (
        selectorEnd >= 0 &&
        formula[selectorEnd + 1] === ":" &&
        formula[selectorEnd + 2] === "["
      ) {
        const spanEnd = structuredReferenceBracketEnd(
          formula,
          selectorEnd + 2,
        );
        if (spanEnd >= 0) selectorEnd = spanEnd;
      }
      if (selectorEnd >= 0) {
        const selector = formula.slice(index, selectorEnd + 1);
        const after = formula[selectorEnd + 1];
        const externalWorkbookRef = Boolean(
          after && /[A-Za-z_'!]/.test(after),
        );
        out +=
          !externalWorkbookRef && selectorTouchesDeletedColumn(selector)
            ? "#REF!"
            : selector;
        index = selectorEnd + 1;
        continue;
      }
    }
    out += character;
    index += 1;
  }
  return out;
}

/**
 * Build a default table model from a selected range: the selection becomes the
 * range (normalized so r1<=r2 / c1<=c2 and clamped into the grid), the first row
 * is the header, banded rows are ON (Excel's default table style). `name` falls
 * back to "Table" + a sequence number; `id` is unique within the process.
 * Returns null when the range is empty/degenerate (a table needs at least a
 * header row — a single row is allowed: header only, zero data rows).
 */
export function buildDataTableFromRange(
  range: { r1: number; c1: number; r2: number; c2: number },
  name?: string,
  columnNames?: readonly string[],
): DataTable | null {
  const r1 = clampIndex(Math.min(range.r1, range.r2), MAX_RANGE_ROW);
  const r2 = clampIndex(Math.max(range.r1, range.r2), MAX_RANGE_ROW);
  const c1 = clampIndex(Math.min(range.c1, range.c2), MAX_RANGE_COL);
  const c2 = clampIndex(Math.max(range.c1, range.c2), MAX_RANGE_COL);
  if (r1 === null || r2 === null || c1 === null || c2 === null) return null;
  tableSeq += 1;
  const trimmed = (name ?? "").trim();
  const finalName =
    trimmed !== "" && trimmed.length <= MAX_DATA_TABLE_NAME_LENGTH
      ? trimmed
      : `Table${tableSeq}`;
  const width = c2 - c1 + 1;
  const names = normalizeTableColumnNames(
    columnNames?.length === width
      ? columnNames
      : Array.from({ length: width }, (_, index) => `${spreadsheetDefaultNames().column}${index + 1}`),
  );
  return {
    id: `table-${tableSeq}-${Date.now().toString(36)}`,
    name: finalName,
    range: { r1, c1, r2, c2 },
    headerRow: true,
    columns: names.map((columnName, index) => ({
      id: index + 1,
      name: columnName,
    })),
    styleId: DEFAULT_NEW_TABLE_STYLE_ID,
    bandedRows: true,
  };
}

/**
 * Compute one column's totals aggregate over its DATA-row strings. `values` are
 * the raw display strings of the column's data rows (r1+1..r2) in order; `kind`
 * is the column's aggregation. Reuses the pivot `aggregate` + `coerceNumber`
 * (identical numeric coercion to SUM/AVERAGE in a pivot), so blanks/non-numeric
 * are IGNORED (never counted as 0). Returns:
 *   - ""    when kind is "none" (blank totals cell)
 *   - a number string for sum/average/min/max/stdDev/var (or "" when the
 *     numeric sample is empty/too small)
 *   - an integer string for count (non-blank entries, COUNTA semantics; 0 shows)
 *   - an integer string for countNums (numeric entries only; 0 shows)
 */
export function tableColumnAggregate(
  values: readonly string[],
  kind: TableTotalKind,
): string {
  if (kind === "none" || kind === "custom") return "";
  const strings = values.map((v) =>
    typeof v === "string" ? v : String(v ?? ""),
  );
  if (kind === "countNums") {
    return String(
      strings.filter((value) => coerceNumber(value) !== null).length,
    );
  }
  if (kind === "stdDev" || kind === "var") {
    const numbers = strings
      .map((value) => coerceNumber(value))
      .filter((value): value is number => value !== null);
    if (numbers.length < 2) return "";
    const mean =
      numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
    const variance =
      numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) /
      (numbers.length - 1);
    const result = kind === "stdDev" ? Math.sqrt(variance) : variance;
    return Number.isInteger(result)
      ? String(result)
      : String(roundTotal(result));
  }
  const result = aggregate(kind, strings);
  if (result === null) return "";
  // Trim float noise the way a totals row should read; integers stay integral.
  return Number.isInteger(result) ? String(result) : String(roundTotal(result));
}

/**
 * The styling role of a grid cell at (row, col) for a table's render layer:
 *   - "header"  : the table's header row (row === r1), any in-range column
 *   - "totals"  : the physical totals row just below the range (row === r2+1)
 *                 when totalsRow is on, any in-range column
 *   - "banded"  : an ODD-offset data row (1st data row is plain, 2nd banded, …)
 *                 when bandedRows is on
 *   - "data"    : any other in-range data row
 *   - null      : the cell is outside the table (including its totals row)
 *
 * The banded stripe alternates by the data row's offset from the FIRST data row
 * (r1+1), so the first data row is "data" and the second "banded" — matching
 * Excel's default table style (header tinted, then alternating bands below it).
 * When `bandedRows` is off every data row is "data". A single-row range (header
 * only, r1===r2) yields only "header" (+ "totals" below when on).
 */
export function tableCellRole(
  table: DataTable,
  row: number,
  col: number,
): TableCellRole {
  const { r1, c1, r2, c2 } = table.range;
  if (col < c1 || col > c2) return null;
  if (table.headerRow && row === r1) return "header";
  if (table.totalsRow && row === r2 + 1) return "totals";
  const firstDataRow = r1 + (table.headerRow ? 1 : 0);
  if (row < firstDataRow || row > r2) return null;
  // Band by offset from the first data row.
  if (table.bandedRows) {
    const dataOffset = row - firstDataRow;
    return dataOffset % 2 === 1 ? "banded" : "data";
  }
  return "data";
}

/**
 * The physical totals row's grid row index (one below the range), or null when
 * the table has no totals row.
 */
export function tableTotalsRowIndex(table: DataTable): number | null {
  return table.totalsRow ? table.range.r2 + 1 : null;
}

function clampIndex(raw: number, max: number): number | null {
  if (!Number.isFinite(raw)) return null;
  const v = Math.floor(raw);
  if (v < 0) return null;
  return Math.min(v, max);
}

/** Round a totals float to a stable 10-decimal value, dropping FP noise. */
function roundTotal(n: number): number {
  return Math.round(n * 1e10) / 1e10;
}

function normalizeTotalKind(raw: unknown): TableTotalKind | null {
  if (typeof raw !== "string") return null;
  return (TABLE_TOTAL_KINDS as readonly string[]).includes(raw)
    ? (raw as TableTotalKind)
    : null;
}

function normalizeTableStyleId(raw: unknown): TableStyleId | null {
  if (typeof raw !== "string") return null;
  return (TABLE_STYLE_IDS as readonly string[]).includes(raw)
    ? (raw as TableStyleId)
    : null;
}

/**
 * Coerce a raw persisted entry into a safe DataTable, or null when malformed
 * (missing/over-long id or name, non-rectangular or out-of-grid range). The
 * range is normalized (r1<=r2 / c1<=c2) and clamped into the grid; `bandedRows`
 * / `totalsRow` are kept only when truthy; `totals` keeps only in-range column
 * offsets with a valid non-"none" kind. Extra fields are ignored.
 *
 * @param raw the persisted entry
 * @param bounds the sheet's row/column count, so a stale table addressing past
 *   the current grid is clamped (a range fully past the grid is dropped)
 */
export function normalizeDataTable(
  raw: unknown,
  bounds: { rowCount: number; columnCount: number },
): DataTable | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const id = typeof r.id === "string" ? r.id.trim() : "";
  if (id === "" || id.length > MAX_DATA_TABLE_ID_LENGTH) return null;
  const name = typeof r.name === "string" ? r.name.trim() : "";
  if (name === "" || name.length > MAX_DATA_TABLE_NAME_LENGTH) return null;
  const rawRange = r.range;
  if (!rawRange || typeof rawRange !== "object") return null;
  const rr = rawRange as Record<string, unknown>;
  const lastRow = Math.max(0, Math.floor(bounds.rowCount) - 1);
  const lastCol = Math.max(0, Math.floor(bounds.columnCount) - 1);
  const maxRow = Number.isFinite(lastRow)
    ? Math.min(lastRow, MAX_RANGE_ROW)
    : MAX_RANGE_ROW;
  const maxCol = Number.isFinite(lastCol)
    ? Math.min(lastCol, MAX_RANGE_COL)
    : MAX_RANGE_COL;
  const r1 = clampIndex(Math.min(Number(rr.r1), Number(rr.r2)), maxRow);
  const r2 = clampIndex(Math.max(Number(rr.r1), Number(rr.r2)), maxRow);
  const c1 = clampIndex(Math.min(Number(rr.c1), Number(rr.c2)), maxCol);
  const c2 = clampIndex(Math.max(Number(rr.c1), Number(rr.c2)), maxCol);
  if (r1 === null || r2 === null || c1 === null || c2 === null) return null;
  const out: DataTable = {
    id,
    name,
    range: { r1, c1, r2, c2 },
    headerRow: r.headerRow !== false,
  };
  const columns = normalizeDataTableColumns(r.columns, c2 - c1 + 1);
  if (columns) out.columns = columns;
  const invalidColumnIdentity =
    (r.columns !== undefined && !columns) || (!out.headerRow && !columns);
  const styleId = normalizeTableStyleId(r.styleId);
  if (styleId) out.styleId = styleId;
  if (typeof r.nativeStyleName === "string") {
    const nativeStyleName = r.nativeStyleName.trim();
    if (nativeStyleName && nativeStyleName.length <= 255) {
      out.nativeStyleName = nativeStyleName;
    }
  }
  const altText = normalizeTableAltText(r.altText, MAX_TABLE_ALT_TEXT_LENGTH);
  if (altText) out.altText = altText;
  const altTextSummary = normalizeTableAltText(
    r.altTextSummary,
    MAX_TABLE_ALT_TEXT_SUMMARY_LENGTH,
  );
  if (altTextSummary) out.altTextSummary = altTextSummary;
  const ownBorders = normalizeTableOwnBorders(r.ownBorders);
  if (ownBorders) out.ownBorders = ownBorders;
  if (
    typeof r.nativeResizeBlocker === "string" &&
    TABLE_NATIVE_RESIZE_BLOCKERS.has(
      r.nativeResizeBlocker as TableNativeResizeBlocker,
    )
  ) {
    out.nativeResizeBlocker = r.nativeResizeBlocker as TableNativeResizeBlocker;
  }
  if (invalidColumnIdentity && !out.nativeResizeBlocker) {
    out.nativeResizeBlocker = "extension";
  }
  if (
    tableNativeFilterCustody(out) &&
    Array.isArray(r.nativeFilterColumnIds) &&
    columns
  ) {
    const known = new Set(columns.map((column) => column.id));
    const ids = Array.from(
      new Set(
        r.nativeFilterColumnIds.filter(
          (id): id is number => typeof id === "number" && known.has(id),
        ),
      ),
    );
    if (ids.length > 0) out.nativeFilterColumnIds = ids;
  }
  if (r.queryBacked === true) out.queryBacked = true;
  if (r.showFilterButton === false) out.showFilterButton = false;
  if (
    r.columnFilters &&
    typeof r.columnFilters === "object" &&
    !Array.isArray(r.columnFilters) &&
    columns
  ) {
    const knownIds = new Set(columns.map((column) => String(column.id)));
    const columnFilters: Record<string, DataTableColumnFilter> = {};
    let invalidFilter = !out.headerRow;
    for (const [columnId, filter] of Object.entries(
      r.columnFilters as Record<string, unknown>,
    )) {
      if (!knownIds.has(columnId)) {
        invalidFilter = true;
        continue;
      }
      const normalized = normalizeDataTableColumnFilter(filter);
      if (normalized) columnFilters[columnId] = normalized;
      else invalidFilter = true;
    }
    if (!invalidFilter && Object.keys(columnFilters).length > 0) {
      out.columnFilters = columnFilters;
    } else if (invalidFilter && !out.nativeResizeBlocker) {
      out.nativeResizeBlocker = "extension";
    }
  } else if (r.columnFilters !== undefined && !out.nativeResizeBlocker) {
    out.nativeResizeBlocker = "extension";
  }
  if (r.sortState !== undefined) {
    const sortState =
      columns && out.headerRow
        ? normalizeDataTableSortState(r.sortState, columns)
        : null;
    if (sortState) out.sortState = sortState;
    else if (!out.nativeResizeBlocker) out.nativeResizeBlocker = "extension";
  }
  if (r.bandedRows === true) out.bandedRows = true;
  if (r.bandedColumns === true) out.bandedColumns = true;
  if (r.firstColumn === true) out.firstColumn = true;
  if (r.lastColumn === true) out.lastColumn = true;
  if (r.totalsRow === true) out.totalsRow = true;
  if (r.totals && typeof r.totals === "object" && !Array.isArray(r.totals)) {
    const totals: Record<number, TableTotalKind> = {};
    const width = c2 - c1;
    for (const [k, v] of Object.entries(r.totals as Record<string, unknown>)) {
      const offset = Number(k);
      if (!Number.isInteger(offset) || offset < 0 || offset > width) continue;
      const kind = normalizeTotalKind(v);
      if (!kind || kind === "none") continue;
      totals[offset] = kind;
    }
    if (Object.keys(totals).length > 0) out.totals = totals;
  }
  if (
    r.totalFormulas &&
    typeof r.totalFormulas === "object" &&
    !Array.isArray(r.totalFormulas) &&
    out.totals
  ) {
    const totalFormulas: Record<number, string> = {};
    const width = c2 - c1;
    for (const [k, v] of Object.entries(
      r.totalFormulas as Record<string, unknown>,
    )) {
      const offset = Number(k);
      if (
        !Number.isInteger(offset) ||
        offset < 0 ||
        offset > width ||
        out.totals[offset] !== "custom" ||
        typeof v !== "string"
      ) {
        continue;
      }
      const formula = v.trim();
      if (isAuthorableTableTotalFormula(formula)) {
        totalFormulas[offset] = formula;
      } else {
        delete out.totals[offset];
      }
    }
    for (const [offset, kind] of Object.entries(out.totals)) {
      if (kind === "custom" && totalFormulas[Number(offset)] === undefined) {
        delete out.totals[Number(offset)];
      }
    }
    if (Object.keys(totalFormulas).length > 0) {
      out.totalFormulas = totalFormulas;
    }
    if (Object.keys(out.totals).length === 0) delete out.totals;
  } else if (out.totals) {
    for (const [offset, kind] of Object.entries(out.totals)) {
      if (kind === "custom") delete out.totals[Number(offset)];
    }
    if (Object.keys(out.totals).length === 0) delete out.totals;
  }
  if (
    r.calculatedFormulas &&
    typeof r.calculatedFormulas === "object" &&
    !Array.isArray(r.calculatedFormulas)
  ) {
    const calculatedFormulas: Record<number, string> = {};
    const width = c2 - c1;
    for (const [rawOffset, rawFormula] of Object.entries(
      r.calculatedFormulas as Record<string, unknown>,
    )) {
      const offset = Number(rawOffset);
      if (
        Number.isInteger(offset) &&
        offset >= 0 &&
        offset <= width &&
        isAuthorableTableTotalFormula(rawFormula)
      ) {
        calculatedFormulas[offset] = rawFormula.trim();
      }
    }
    if (Object.keys(calculatedFormulas).length > 0) {
      out.calculatedFormulas = calculatedFormulas;
    }
  }
  if (
    r.arrayCalculatedColumns &&
    typeof r.arrayCalculatedColumns === "object" &&
    !Array.isArray(r.arrayCalculatedColumns)
  ) {
    const arrayCalculatedColumns: Record<number, true> = {};
    for (const rawOffset of Object.keys(r.arrayCalculatedColumns)) {
      const offset = Number(rawOffset);
      if (Number.isInteger(offset) && offset >= 0 && offset <= c2 - c1) {
        arrayCalculatedColumns[offset] = true;
      }
    }
    if (Object.keys(arrayCalculatedColumns).length > 0) {
      out.arrayCalculatedColumns = arrayCalculatedColumns;
    }
  }
  if (
    r.totalLabels &&
    typeof r.totalLabels === "object" &&
    !Array.isArray(r.totalLabels)
  ) {
    const totalLabels: Record<number, string> = {};
    const width = c2 - c1;
    for (const [k, v] of Object.entries(
      r.totalLabels as Record<string, unknown>,
    )) {
      const offset = Number(k);
      if (!Number.isInteger(offset) || offset < 0 || offset > width) continue;
      if (typeof v !== "string") continue;
      const label = v.trim().slice(0, MAX_TOTALS_LABEL_LENGTH);
      if (label === "") continue;
      // An aggregate wins over a label on the same column, mirroring Excel.
      if (out.totals && out.totals[offset]) continue;
      totalLabels[offset] = label;
    }
    if (Object.keys(totalLabels).length > 0) out.totalLabels = totalLabels;
  }
  return out;
}

/**
 * Normalize a raw `dataTables` array off a persisted sheet: drop malformed
 * entries, drop a duplicate id, and cap the count (DoS bound). Returns a fresh
 * array; empty when nothing valid remains so the no-table path stays
 * byte-identical (the field is never serialized when empty).
 */
export function normalizeDataTables(
  raw: unknown,
  bounds: { rowCount: number; columnCount: number },
): DataTable[] {
  if (!Array.isArray(raw)) return [];
  const out: DataTable[] = [];
  const usedIds = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_DATA_TABLES) break;
    const table = normalizeDataTable(entry, bounds);
    if (!table) continue;
    if (usedIds.has(table.id)) continue;
    usedIds.add(table.id);
    out.push(table);
  }
  return out;
}

/** The 0-based column offset of `col` within a table's range (or null if out). */
export function tableColumnOffset(
  table: DataTable,
  col: number,
): number | null {
  if (col < table.range.c1 || col > table.range.c2) return null;
  return col - table.range.c1;
}
