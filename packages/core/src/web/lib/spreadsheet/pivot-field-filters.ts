/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { excelSerialFromUtcMs } from "./excel-1900-calendar";

export const PIVOT_LABEL_FILTER_TYPES = [
  "captionEqual",
  "captionNotEqual",
  "captionBeginsWith",
  "captionNotBeginsWith",
  "captionEndsWith",
  "captionNotEndsWith",
  "captionContains",
  "captionNotContains",
  "captionGreaterThan",
  "captionGreaterThanOrEqual",
  "captionLessThan",
  "captionLessThanOrEqual",
  "captionBetween",
  "captionNotBetween",
] as const;

export const PIVOT_VALUE_FILTER_TYPES = [
  "valueEqual",
  "valueNotEqual",
  "valueGreaterThan",
  "valueGreaterThanOrEqual",
  "valueLessThan",
  "valueLessThanOrEqual",
  "valueBetween",
  "valueNotBetween",
] as const;

/** Top 10 by items, by percent and by sum. */
export const PIVOT_TOP_FILTER_TYPES = ["count", "percent", "sum"] as const;

export const PIVOT_DATE_FILTER_TYPES = [
  "dateEqual",
  "dateNotEqual",
  "dateOlderThan",
  "dateOlderThanOrEqual",
  "dateNewerThan",
  "dateNewerThanOrEqual",
  "dateBetween",
  "dateNotBetween",
  "tomorrow",
  "today",
  "yesterday",
  "nextWeek",
  "thisWeek",
  "lastWeek",
  "nextMonth",
  "thisMonth",
  "lastMonth",
  "nextQuarter",
  "thisQuarter",
  "lastQuarter",
  "nextYear",
  "thisYear",
  "lastYear",
  "yearToDate",
  "Q1",
  "Q2",
  "Q3",
  "Q4",
  "M1",
  "M2",
  "M3",
  "M4",
  "M5",
  "M6",
  "M7",
  "M8",
  "M9",
  "M10",
  "M11",
  "M12",
] as const;

export type PivotLabelFilterType = (typeof PIVOT_LABEL_FILTER_TYPES)[number];
export type PivotValueFilterType = (typeof PIVOT_VALUE_FILTER_TYPES)[number];
export type PivotTopFilterType = (typeof PIVOT_TOP_FILTER_TYPES)[number];
export type PivotDateFieldFilterType = (typeof PIVOT_DATE_FILTER_TYPES)[number];
export type PivotFieldFilterType =
  | PivotLabelFilterType
  | PivotValueFilterType
  | PivotTopFilterType
  | PivotDateFieldFilterType;

export type PivotFieldFilterKind = "label" | "date" | "value" | "top";

export interface PivotFieldFilter {
  /** The field, as item filters name it: a source column, with `unit` for a
   * level of a date field grouped by several units, `group` for the group
   * field of a manual grouping. */
  colIndex: number;
  unit?: "year" | "quarter" | "month";
  group?: true;
  type: PivotFieldFilterType;
  /** Label filters: the text typed (Excel's stringValue1/2). Value filters:
   * the number(s). Top 10: N. Date filters: the day(s) as a 1900 serial. */
  value1?: string;
  value2?: string;
  /** Value filters and Top 10: the value field they judge by. */
  valueIndex?: number;
  /** Top 10: the bottom items instead of the top. */
  bottom?: true;
  /** The native filter's id, kept through a rewrite. */
  id?: number;
}

const LABEL = new Set<string>(PIVOT_LABEL_FILTER_TYPES);
const VALUE = new Set<string>(PIVOT_VALUE_FILTER_TYPES);
const TOP = new Set<string>(PIVOT_TOP_FILTER_TYPES);
const DATE = new Set<string>(PIVOT_DATE_FILTER_TYPES);
const TWO_OPERANDS = new Set<string>([
  "captionBetween",
  "captionNotBetween",
  "valueBetween",
  "valueNotBetween",
  "dateBetween",
  "dateNotBetween",
]);
const NO_OPERAND = new Set<string>(PIVOT_DATE_FILTER_TYPES.slice(8));

export function pivotFieldFilterKind(type: PivotFieldFilterType): PivotFieldFilterKind {
  if (LABEL.has(type)) return "label";
  if (VALUE.has(type)) return "value";
  if (TOP.has(type)) return "top";
  return "date";
}

/** Whether the filter judges items by a value field (Value Filters, Top 10). */
export function pivotFieldFilterNeedsValue(type: PivotFieldFilterType): boolean {
  return VALUE.has(type) || TOP.has(type);
}

export function pivotFieldFilterOperands(type: PivotFieldFilterType): 0 | 1 | 2 {
  if (NO_OPERAND.has(type)) return 0;
  return TWO_OPERANDS.has(type) ? 2 : 1;
}

const MAX_FIELD_FILTERS = 64;
const MAX_OPERAND = 255;

function operand(value: unknown): string | undefined {
  if (typeof value !== "string" && typeof value !== "number") return undefined;
  const text = String(value);
  if (text.length > MAX_OPERAND || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) {
    return undefined;
  }
  return text;
}

function finite(value: string | undefined): number | null {
  if (value === undefined || value.trim() === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

/**
 * A bounded list of well-formed filters: one label-or-date and one value-or-
 * Top-10 filter per field level (Excel replaces a field's filter of the same
 * family when another is applied), each with the operands its type needs.
 */
export function normalizePivotFieldFilters(raw: unknown, maxCol: number): PivotFieldFilter[] {
  if (!Array.isArray(raw)) return [];
  const out: PivotFieldFilter[] = [];
  const seen = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_FIELD_FILTERS) break;
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const colIndex = e.colIndex;
    if (typeof colIndex !== "number" || !Number.isInteger(colIndex) || colIndex < 0 || colIndex > maxCol) {
      continue;
    }
    const type = e.type as PivotFieldFilterType;
    if (!LABEL.has(type) && !VALUE.has(type) && !TOP.has(type) && !DATE.has(type)) continue;
    const unit = e.unit === "year" || e.unit === "quarter" || e.unit === "month" ? e.unit : undefined;
    if (e.unit !== undefined && unit === undefined) continue;
    const group = e.group === true ? (true as const) : undefined;
    if (group && unit) continue;
    const value1 = operand(e.value1);
    const value2 = operand(e.value2);
    const operands = pivotFieldFilterOperands(type);
    if ((operands >= 1 && value1 === undefined) || (operands === 2 && value2 === undefined)) continue;
    const kind = pivotFieldFilterKind(type);
    if ((kind === "value" || kind === "top") && finite(value1) === null) continue;
    if (kind === "value" && operands === 2 && finite(value2) === null) continue;
    if (kind === "date" && operands >= 1 && (finite(value1) === null || (operands === 2 && finite(value2) === null))) {
      continue;
    }
    const valueIndex =
      typeof e.valueIndex === "number" && Number.isInteger(e.valueIndex) && e.valueIndex >= 0 && e.valueIndex < 8
        ? e.valueIndex
        : undefined;
    if ((kind === "value" || kind === "top") && valueIndex === undefined) continue;
    const family = kind === "value" || kind === "top" ? "value" : "label";
    const key = `${colIndex}:${unit ?? ""}:${group ? "g" : ""}:${family}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const id =
      typeof e.id === "number" && Number.isInteger(e.id) && e.id >= 0 && e.id <= 2_147_483_647 ? e.id : undefined;
    out.push({
      colIndex,
      ...(unit ? { unit } : {}),
      ...(group ? { group } : {}),
      type,
      ...(operands >= 1 && value1 !== undefined ? { value1 } : {}),
      ...(operands === 2 && value2 !== undefined ? { value2 } : {}),
      ...(kind === "value" || kind === "top" ? { valueIndex } : {}),
      ...(kind === "top" && e.bottom === true ? { bottom: true as const } : {}),
      ...(id !== undefined ? { id } : {}),
    });
  }
  return out;
}

/** Excel's label-filter text match: case-insensitive, `*` and `?` wildcards,
 * `~` escaping either. */
function wildcardRegExp(pattern: string, anchorStart: boolean, anchorEnd: boolean): RegExp {
  let source = "";
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === "~" && index + 1 < pattern.length && "*?~".includes(pattern[index + 1])) {
      source += `\\${pattern[index + 1]}`;
      index += 1;
    } else if (char === "*") source += "[\\s\\S]*";
    else if (char === "?") source += "[\\s\\S]";
    else source += char.replace(/[.+^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`${anchorStart ? "^" : ""}${source}${anchorEnd ? "$" : ""}`, "i");
}

function compareCaptions(item: string, operandText: string): number {
  const left = finite(item);
  const right = finite(operandText);
  if (left !== null && right !== null) return left - right;
  return item.localeCompare(operandText, undefined, { sensitivity: "accent" });
}

export function pivotLabelFilterPasses(filter: PivotFieldFilter, caption: string): boolean {
  const first = filter.value1 ?? "";
  const second = filter.value2 ?? "";
  switch (filter.type) {
    case "captionEqual":
      return wildcardRegExp(first, true, true).test(caption);
    case "captionNotEqual":
      return !wildcardRegExp(first, true, true).test(caption);
    case "captionBeginsWith":
      return wildcardRegExp(first, true, false).test(caption);
    case "captionNotBeginsWith":
      return !wildcardRegExp(first, true, false).test(caption);
    case "captionEndsWith":
      return wildcardRegExp(first, false, true).test(caption);
    case "captionNotEndsWith":
      return !wildcardRegExp(first, false, true).test(caption);
    case "captionContains":
      return wildcardRegExp(first, false, false).test(caption);
    case "captionNotContains":
      return !wildcardRegExp(first, false, false).test(caption);
    case "captionGreaterThan":
      return compareCaptions(caption, first) > 0;
    case "captionGreaterThanOrEqual":
      return compareCaptions(caption, first) >= 0;
    case "captionLessThan":
      return compareCaptions(caption, first) < 0;
    case "captionLessThanOrEqual":
      return compareCaptions(caption, first) <= 0;
    case "captionBetween":
    case "captionNotBetween": {
      const low = compareCaptions(first, second) <= 0 ? first : second;
      const high = low === first ? second : first;
      const inside = compareCaptions(caption, low) >= 0 && compareCaptions(caption, high) <= 0;
      return filter.type === "captionBetween" ? inside : !inside;
    }
    default:
      return true;
  }
}

const ISO_DAY = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?)?Z?$/;

/** An item's date as a 1900 serial (an ISO date or dateTime, a serial). */
export function pivotItemSerial(item: string): number | null {
  const match = ISO_DAY.exec(item.trim());
  if (match) {
    const ms = Date.UTC(+match[1], +match[2] - 1, +match[3], +(match[4] ?? 0), +(match[5] ?? 0), +(match[6] ?? 0));
    return Number.isNaN(ms) ? null : excelSerialFromUtcMs(ms);
  }
  const number = finite(item);
  return number !== null && number >= 1 ? number : null;
}

function serialOf(year: number, monthIndex: number, day: number): number {
  return excelSerialFromUtcMs(Date.UTC(year, monthIndex, day));
}

/** The [start, end) day range of a dynamic date filter, today being `today`. */
function dynamicRange(type: string, today: Date): [number, number] | null {
  const y = today.getFullYear();
  const m = today.getMonth();
  const d = today.getDate();
  const day = serialOf(y, m, d);
  // Excel's weeks start on Sunday.
  const weekStart = day - today.getDay();
  const quarter = Math.floor(m / 3);
  switch (type) {
    case "today":
      return [day, day + 1];
    case "yesterday":
      return [day - 1, day];
    case "tomorrow":
      return [day + 1, day + 2];
    case "thisWeek":
      return [weekStart, weekStart + 7];
    case "lastWeek":
      return [weekStart - 7, weekStart];
    case "nextWeek":
      return [weekStart + 7, weekStart + 14];
    case "thisMonth":
      return [serialOf(y, m, 1), serialOf(y, m + 1, 1)];
    case "lastMonth":
      return [serialOf(y, m - 1, 1), serialOf(y, m, 1)];
    case "nextMonth":
      return [serialOf(y, m + 1, 1), serialOf(y, m + 2, 1)];
    case "thisQuarter":
      return [serialOf(y, quarter * 3, 1), serialOf(y, quarter * 3 + 3, 1)];
    case "lastQuarter":
      return [serialOf(y, quarter * 3 - 3, 1), serialOf(y, quarter * 3, 1)];
    case "nextQuarter":
      return [serialOf(y, quarter * 3 + 3, 1), serialOf(y, quarter * 3 + 6, 1)];
    case "thisYear":
      return [serialOf(y, 0, 1), serialOf(y + 1, 0, 1)];
    case "lastYear":
      return [serialOf(y - 1, 0, 1), serialOf(y, 0, 1)];
    case "nextYear":
      return [serialOf(y + 1, 0, 1), serialOf(y + 2, 0, 1)];
    case "yearToDate":
      return [serialOf(y, 0, 1), day + 1];
    default:
      return null;
  }
}

function monthOfSerial(serial: number): number {
  // 1900 serials from March 1900 on map onto the proleptic calendar.
  const ms = Date.UTC(1899, 11, 30) + Math.floor(serial) * 86_400_000;
  return new Date(ms).getUTCMonth();
}

export function pivotDateFilterPasses(filter: PivotFieldFilter, item: string, today: Date = new Date()): boolean {
  const serial = pivotItemSerial(item);
  if (serial === null) return false;
  const itemDay = Math.floor(serial);
  const first = Math.floor(finite(filter.value1) ?? 0);
  const second = Math.floor(finite(filter.value2) ?? 0);
  switch (filter.type) {
    case "dateEqual":
      return itemDay === first;
    case "dateNotEqual":
      return itemDay !== first;
    case "dateOlderThan":
      return itemDay < first;
    case "dateOlderThanOrEqual":
      return itemDay <= first;
    case "dateNewerThan":
      return itemDay > first;
    case "dateNewerThanOrEqual":
      return itemDay >= first;
    case "dateBetween":
    case "dateNotBetween": {
      const inside = itemDay >= Math.min(first, second) && itemDay <= Math.max(first, second);
      return filter.type === "dateBetween" ? inside : !inside;
    }
    default:
      break;
  }
  const quarterMatch = /^Q([1-4])$/.exec(filter.type);
  if (quarterMatch) return Math.floor(monthOfSerial(itemDay) / 3) === +quarterMatch[1] - 1;
  const monthMatch = /^M(\d{1,2})$/.exec(filter.type);
  if (monthMatch) return monthOfSerial(itemDay) === +monthMatch[1] - 1;
  const range = dynamicRange(filter.type, today);
  return range !== null && itemDay >= range[0] && itemDay < range[1];
}

function valuePasses(filter: PivotFieldFilter, value: number | null): boolean {
  if (value === null || !Number.isFinite(value)) return false;
  const first = finite(filter.value1) ?? 0;
  const second = finite(filter.value2) ?? 0;
  switch (filter.type) {
    case "valueEqual":
      return value === first;
    case "valueNotEqual":
      return value !== first;
    case "valueGreaterThan":
      return value > first;
    case "valueGreaterThanOrEqual":
      return value >= first;
    case "valueLessThan":
      return value < first;
    case "valueLessThanOrEqual":
      return value <= first;
    case "valueBetween":
    case "valueNotBetween": {
      const inside = value >= Math.min(first, second) && value <= Math.max(first, second);
      return filter.type === "valueBetween" ? inside : !inside;
    }
    default:
      return true;
  }
}

/** The items a Top 10 keeps among one outer item's items. */
function topItems(filter: PivotFieldFilter, measured: readonly { item: string; value: number | null }[]): Set<string> {
  const ranked = measured
    .filter((entry): entry is { item: string; value: number } => entry.value !== null && Number.isFinite(entry.value))
    .sort((a, b) => (filter.bottom ? a.value - b.value : b.value - a.value));
  const limit = finite(filter.value1) ?? 0;
  const kept = new Set<string>();
  if (filter.type === "count") {
    const count = Math.floor(limit);
    if (count <= 0 || ranked.length === 0) return kept;
    const threshold = ranked[Math.min(count, ranked.length) - 1].value;
    for (const entry of ranked) {
      if (filter.bottom ? entry.value <= threshold : entry.value >= threshold) kept.add(entry.item);
    }
    return kept;
  }
  const total = ranked.reduce((sum, entry) => sum + entry.value, 0);
  const goal = filter.type === "percent" ? (total * limit) / 100 : limit;
  // Largest first (smallest for Bottom) until the items reach the goal; the
  // item that reaches it is kept.
  let running = 0;
  for (const entry of ranked) {
    if (goal <= 0) break;
    kept.add(entry.item);
    running += entry.value;
    if (running >= goal) break;
  }
  return kept;
}

export interface PivotFieldFilterContext<Row> {
  rowFields: readonly number[];
  columnFields: readonly number[];
  /** The item a record has on a field (its group label on a grouped one). */
  itemOf: (row: Row, field: number) => string;
  /** A value field's aggregate over some records. */
  measure: (rows: readonly Row[], valueIndex: number) => number | null;
  valueCount: number;
  today?: Date;
}

/**
 * The records left after a report's field filters, applied in order. A
 * filter on a field off both axes, or naming a value field the report no
 * longer has, filters nothing (Excel drops such a filter).
 */
export function applyPivotFieldFilters<Row>(
  rows: readonly Row[],
  filters: readonly PivotFieldFilter[] | undefined,
  context: PivotFieldFilterContext<Row>,
): Row[] {
  let kept = [...rows];
  for (const filter of filters ?? []) {
    const rowDepth = context.rowFields.indexOf(filter.colIndex);
    const columnDepth = context.columnFields.indexOf(filter.colIndex);
    const axis = rowDepth >= 0 ? context.rowFields : columnDepth >= 0 ? context.columnFields : null;
    const depth = rowDepth >= 0 ? rowDepth : columnDepth;
    if (!axis) continue;
    const kind = pivotFieldFilterKind(filter.type);
    if (kind === "label" || kind === "date") {
      kept = kept.filter((row) => {
        const item = context.itemOf(row, filter.colIndex);
        return kind === "label"
          ? pivotLabelFilterPasses(filter, item)
          : pivotDateFilterPasses(filter, item, context.today);
      });
      continue;
    }
    if (filter.valueIndex === undefined || filter.valueIndex >= context.valueCount) continue;
    // Each item is judged within its outer items on the same axis.
    const parents = new Map<string, Map<string, Row[]>>();
    for (const row of kept) {
      const parentKey = axis
        .slice(0, depth)
        .map((field) => context.itemOf(row, field))
        .join("\u0000");
      const item = context.itemOf(row, filter.colIndex);
      let items = parents.get(parentKey);
      if (!items) parents.set(parentKey, (items = new Map()));
      const group = items.get(item);
      if (group) group.push(row);
      else items.set(item, [row]);
    }
    const survivors = new Set<Row>();
    for (const items of parents.values()) {
      const measured = [...items.entries()].map(([item, group]) => ({
        item,
        value: context.measure(group, filter.valueIndex as number),
      }));
      const passing =
        kind === "top"
          ? topItems(filter, measured)
          : new Set(measured.filter((entry) => valuePasses(filter, entry.value)).map((entry) => entry.item));
      for (const [item, group] of items) {
        if (passing.has(item)) for (const row of group) survivors.add(row);
      }
    }
    kept = kept.filter((row) => survivors.has(row));
  }
  return kept;
}

/** Move each filter on a level of a grouped field to that level's column of
 * the expanded report (see `expandPivotGroupLevels`); a level the field no
 * longer has drops its filter. */
export function remapPivotFieldFilterLevels(
  filters: readonly PivotFieldFilter[] | undefined,
  levels: ReadonlyMap<number, { colIndex: number; unit?: string; group?: true }>,
): PivotFieldFilter[] | undefined {
  if (filters === undefined) return undefined;
  return filters.flatMap((entry) => {
    if (entry.unit === undefined && entry.group === undefined) return [entry];
    for (const [column, level] of levels) {
      if (level.colIndex === entry.colIndex && level.unit === entry.unit && level.group === entry.group) {
        return [{ ...entry, colIndex: column }];
      }
    }
    return [];
  });
}
