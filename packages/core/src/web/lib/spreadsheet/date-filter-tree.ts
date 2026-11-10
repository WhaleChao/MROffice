/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------

import {
  EXCEL_1900_FICTITIOUS_LEAP_SERIAL,
  excelSystemSerialFromUtcMs,
  excelSystemSerialToUtcMs,
} from "./excel-1900-calendar";
import {
  dateInPeriod,
  type DateOccurringPeriod,
} from "./conditional-format";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";
import { isExcelDate1904, excelSystemHasFictitiousLeapDay } from "./excel-1900-calendar";
// AutoFilter date grouping (Excel / WPS parity).
//
// When a filtered column holds dates, Excel/WPS replace the flat distinct-value
// checkbox list with an expandable Year > Month > Day tree carrying tri-state
// checkboxes (a year is checked when all its months are, indeterminate when only
// some are, etc.). This module is PURE (no React/DOM): it parses a column's
// DISPLAY strings into calendar dates, groups them into that tree, and provides
// the tri-state math + the set of leaf display-strings a node controls — so the
// existing value `filters[col]` (a list of allowed *display strings*) stays the
// single source of truth and `applyFilter` is unchanged.
//
// Parsing operates on the DISPLAYED text (what the grid shows after number
// formatting), mirroring how the flat value list is built. Recognized forms:
//   - ISO-ish:        YYYY-MM-DD or YYYY/MM/DD            (unambiguous)
//   - US slash/dot:   M/D/YYYY  M-D-YYYY  M.D.YYYY        (month-first)
//   - Month name:     "Mar 15, 2024"  "March 15 2024"  "15 March 2024"
// A 2-digit year is windowed like Excel (00-29 -> 2000-2029, 30-99 -> 1930-1999).
// Anything that does not parse (or is blank) is collected into a single
// "non-date" bucket so a column with a few stray labels still groups its dates.
// ---------------------------------------------------------------------------

/** A parsed calendar date with an optional recoverable 24-hour time prefix. */
export interface FilterDate {
  year: number;
  /** 1-12 */
  month: number;
  /** 1-31 */
  day: number;
  /** 0-23; present only when the displayed value carries an explicit time. */
  hour?: number;
  /** 0-59; present whenever hour is present. */
  minute?: number;
  /** 0-59; omitted when the display does not reveal seconds. */
  second?: number;
}

export interface FilterDateAuthority {
  /** Exact displayed value -> actual Gregorian components. */
  byDisplay: ReadonlyMap<string, FilterDate>;
  /** False when any nonblank display is unrecoverable or maps ambiguously. */
  complete: boolean;
}

/** Gregorian calendar day used to evaluate a dynamic relative-date filter. */
export interface FilterDateToday {
  year: number;
  month: number;
  day: number;
}

/** Semantic components exposed by one structurally valid numeric format. */
export interface SpreadsheetNumberFormatSemantics {
  valid: boolean;
  /** A Gregorian calendar component (year/day/era, or an unambiguous month). */
  hasCalendarDate: boolean;
  /** A clock or elapsed-time component. */
  hasClock: boolean;
  /** True only for bracketed elapsed-time tokens such as `[h]` or `[mm]`. */
  hasElapsedClock: boolean;
  /** Safe to reuse after a workbook-epoch serial translation. */
  epochInvariant: boolean;
}

interface ParsedNumberFormatSection
  extends Omit<SpreadsheetNumberFormatSemantics, "epochInvariant"> {
  condition?: {
    operator: "<" | "<=" | "=" | "<>" | ">=" | ">";
    operand: number;
  };
}

const INVALID_NUMBER_FORMAT: SpreadsheetNumberFormatSemantics = Object.freeze({
  valid: false,
  hasCalendarDate: false,
  hasClock: false,
  hasElapsedClock: false,
  epochInvariant: false,
});

const FORMAT_CONDITION = /^\s*(<=|>=|<>|=|<|>)\s*(-?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?)\s*$/;

function splitSpreadsheetNumberFormat(format: string): string[] | null {
  const sections: string[] = [];
  let current = "";
  let quoted = false;
  let bracketed = false;
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index] || "";
    if (quoted) {
      current += character;
      if (character === '"') quoted = false;
      continue;
    }
    if (bracketed) {
      if (character === "[") return null;
      current += character;
      if (character === "]") bracketed = false;
      continue;
    }
    if (character === '"') {
      quoted = true;
      current += character;
      continue;
    }
    if (character === "[") {
      bracketed = true;
      current += character;
      continue;
    }
    if (character === "\\" || character === "_" || character === "*") {
      if (index + 1 >= format.length) return null;
      current += character + (format[++index] || "");
      continue;
    }
    if (character === ";") {
      sections.push(current);
      current = "";
      continue;
    }
    current += character;
  }
  if (quoted || bracketed) return null;
  sections.push(current);
  return sections.length <= 4 ? sections : null;
}

function parseSpreadsheetNumberFormatSection(
  section: string,
): ParsedNumberFormatSection {
  let body = "";
  let condition: ParsedNumberFormatSection["condition"];
  let hasElapsedClock = false;
  for (let index = 0; index < section.length; index += 1) {
    const character = section[index] || "";
    if (character === '"') {
      index += 1;
      while (index < section.length && section[index] !== '"') index += 1;
      if (index >= section.length) return INVALID_NUMBER_FORMAT;
      continue;
    }
    if (character === "\\" || character === "_" || character === "*") {
      if (index + 1 >= section.length) return INVALID_NUMBER_FORMAT;
      index += 1;
      continue;
    }
    if (character === "[") {
      const close = section.indexOf("]", index + 1);
      if (close < 0) return INVALID_NUMBER_FORMAT;
      const token = section.slice(index + 1, close);
      const comparison = FORMAT_CONDITION.exec(token);
      if (comparison) {
        const operand = Number(comparison[2]);
        if (condition || !Number.isFinite(operand)) return INVALID_NUMBER_FORMAT;
        condition = {
          operator: comparison[1] as NonNullable<
            ParsedNumberFormatSection["condition"]
          >["operator"],
          operand,
        };
      } else if (/^\s*[<>=]/.test(token)) {
        return INVALID_NUMBER_FORMAT;
      } else if (/^(?:h+|m+|s+)$/i.test(token.trim())) {
        hasElapsedClock = true;
      }
      index = close;
      continue;
    }
    body += character;
  }

  const semanticBody = body
    .replace(/general/gi, "")
    .replace(/G\/通用格式/g, "")
    .replace(/通用格式/g, "")
    .replace(/[eE][+-]?[0#?]+/g, "");
  const hasYearOrDay = /[yd]/i.test(semanticBody);
  const hasEra = /(^|[^A-Za-z])e+($|[^A-Za-z])/i.test(semanticBody);
  const hasHourOrSecond = /[hs]/i.test(semanticBody);
  const hasAmPm = /(?:am\/pm|a\/p)/i.test(semanticBody);
  const hasClock = hasElapsedClock || hasHourOrSecond || hasAmPm;
  const hasCalendarDate =
    hasYearOrDay || hasEra || (/[m]/i.test(semanticBody) && !hasClock);
  return {
    valid: true,
    hasCalendarDate,
    hasClock,
    hasElapsedClock,
    ...(condition ? { condition } : {}),
  };
}

function numberFormatConditionMatches(
  value: number,
  condition: NonNullable<ParsedNumberFormatSection["condition"]>,
): boolean {
  switch (condition.operator) {
    case "<":
      return value < condition.operand;
    case "<=":
      return value <= condition.operand;
    case "=":
      return value === condition.operand;
    case "<>":
      return value !== condition.operand;
    case ">=":
      return value >= condition.operand;
    case ">":
      return value > condition.operand;
  }
}

/**
 * Parse one SpreadsheetML/Excel number format for a specific numeric value.
 *
 * Quotes, escapes and bracket tokens are lexical, not date evidence. Numeric
 * multi-section formats select the same positive/negative/zero section Excel
 * would use; supported comparison conditions are evaluated before semantics.
 * Malformed structure fails closed so import, menu authority and runtime never
 * disagree about whether a serial is an actual calendar or clock value.
 */
export function classifySpreadsheetNumberFormat(
  format: string,
  numericValue: number,
): SpreadsheetNumberFormatSemantics {
  if (!Number.isFinite(numericValue)) return INVALID_NUMBER_FORMAT;
  const rawSections = splitSpreadsheetNumberFormat(format);
  if (!rawSections || rawSections.length === 0) return INVALID_NUMBER_FORMAT;
  const sections = rawSections.map(parseSpreadsheetNumberFormatSection);
  if (sections.some((section) => !section.valid)) {
    return INVALID_NUMBER_FORMAT;
  }
  const numericSections = sections.slice(0, 3);
  const hasConditions = numericSections.some((section) => section.condition);
  let selected: ParsedNumberFormatSection | undefined;
  if (hasConditions) {
    let sawFallback = false;
    for (const section of numericSections) {
      if (!section.condition) {
        if (sawFallback) return INVALID_NUMBER_FORMAT;
        selected = section;
        sawFallback = true;
        continue;
      }
      if (sawFallback) return INVALID_NUMBER_FORMAT;
      if (numberFormatConditionMatches(numericValue, section.condition)) {
        selected = section;
        break;
      }
    }
  } else {
    selected =
      numericValue > 0
        ? numericSections[0]
        : numericValue < 0
          ? (numericSections[1] ?? numericSections[0])
          : (numericSections[2] ?? numericSections[0]);
  }
  return selected
    ? {
        valid: selected.valid,
        hasCalendarDate: selected.hasCalendarDate,
        hasClock: selected.hasClock,
        hasElapsedClock: selected.hasElapsedClock,
        epochInvariant: !hasConditions && numericSections.length === 1,
      }
    : INVALID_NUMBER_FORMAT;
}

/**
 * Resolve the user's local Gregorian day from an injected clock instant.
 * Keeping the clock at the boundary makes midnight behavior deterministic in
 * tests while the table predicate itself remains pure.
 */
export function localFilterDateToday(now: Date = new Date()): FilterDateToday {
  return {
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    day: now.getDate(),
  };
}

/** Match one authoritative cell date against one supported dynamic period. */
export function filterDateMatchesPeriod(
  actual: FilterDate | null,
  period: DateOccurringPeriod,
  today: FilterDateToday,
): boolean {
  return actual !== null && dateInPeriod(actual, period, today);
}

// Serials follow Excel's 1900 date system (excel-1900-calendar): 1900-01-01
// is serial 1 and serial 60 is the fictitious 1900-02-29.
const FILTER_DATE_MS_PER_DAY = 86_400_000;
const ISO_CALENDAR_VALUE_RE =
  /^([0-9]{4})-([0-9]{2})-([0-9]{2})(?:T([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\.([0-9]{1,3}))?)?(?:(Z)|([+-])([0-9]{2}):([0-9]{2}))?)?$/;

/**
 * Parse the strict ISO calendar subset accepted for SpreadsheetML `t="d"`.
 *
 * A date-only value is interpreted as a Gregorian calendar day. A timestamp
 * keeps its written wall-clock components, including when a syntactically
 * valid offset is present. This matches openpyxl's base-OOXML `t="d"`
 * projection; offsets are validated but do not shift the editor serial. The
 * grammar is deliberately narrower than JavaScript's Date parser so Web and
 * API import grant native filter authority to exactly the same source values.
 */
export function spreadsheetIsoCalendarProjection(
  value: string,
  date1904: ExcelDateSystemArg = false,
): {
  serial: number;
  numberFormat: "yyyy-mm-dd" | "yyyy-mm-dd hh:mm:ss";
} | null {
  const match = ISO_CALENDAR_VALUE_RE.exec(value);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = match[4] === undefined ? 0 : Number(match[4]);
  const minute = match[5] === undefined ? 0 : Number(match[5]);
  const second = match[6] === undefined ? 0 : Number(match[6]);
  const millisecond = Number((match[7] || "").padEnd(3, "0") || "0");
  if (
    !isValidYmd(year, month, day) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59 ||
    second < 0 ||
    second > 59
  ) {
    return null;
  }
  const offsetHour = match[10] === undefined ? 0 : Number(match[10]);
  const offsetMinute = match[11] === undefined ? 0 : Number(match[11]);
  if (
    offsetHour > 14 ||
    offsetMinute > 59 ||
    (offsetHour === 14 && offsetMinute !== 0)
  ) {
    return null;
  }
  const local = new Date(0);
  local.setUTCFullYear(year, month - 1, day);
  local.setUTCHours(hour, minute, second, millisecond);
  if (
    local.getUTCFullYear() !== year ||
    local.getUTCMonth() + 1 !== month ||
    local.getUTCDate() !== day ||
    local.getUTCHours() !== hour ||
    local.getUTCMinutes() !== minute ||
    local.getUTCSeconds() !== second ||
    local.getUTCMilliseconds() !== millisecond
  ) {
    return null;
  }
  const instantMs = local.getTime();
  const instant = new Date(instantMs);
  if (
    !Number.isFinite(instantMs) ||
    instant.getUTCFullYear() < 1 ||
    instant.getUTCFullYear() > 9999
  ) {
    return null;
  }
  // A workbook in the 1904 date system has no day before 1904-01-01.
  if (isExcelDate1904(date1904) && instant.getUTCFullYear() < 1904) return null;
  return {
    serial: Number(excelSystemSerialFromUtcMs(instantMs, date1904).toFixed(10)),
    numberFormat:
      match[4] === undefined ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss",
  };
}

/** The serial of one strict SpreadsheetML `t="d"` value in the workbook's
 *  date system. */
export function spreadsheetSerialFromIsoCalendarValue(
  value: string,
  date1904: ExcelDateSystemArg = false,
): number | null {
  return spreadsheetIsoCalendarProjection(value, date1904)?.serial ?? null;
}

/**
 * Resolve a serial in the workbook's date system into UTC components. In the
 * 1900 system Excel's fictitious serial 60 groups as 1900-02-29, as Excel's
 * filter shows it; the 1904 system reads serial 0 as 1904-01-01.
 */
export function filterDateFromSpreadsheetSerial(
  serial: number,
  date1904: ExcelDateSystemArg = false,
): FilterDate | null {
  if (!Number.isFinite(serial)) return null;
  const daySeconds = FILTER_DATE_MS_PER_DAY / 1_000;
  const wholeSeconds = Math.round(serial * daySeconds);
  const wholeDay = Math.floor(wholeSeconds / daySeconds);
  const secondOfDay = wholeSeconds - wholeDay * daySeconds;
  if (excelSystemHasFictitiousLeapDay(date1904) && wholeDay === EXCEL_1900_FICTITIOUS_LEAP_SERIAL) {
    return {
      year: 1900,
      month: 2,
      day: 29,
      hour: Math.floor(secondOfDay / 3_600),
      minute: Math.floor((secondOfDay % 3_600) / 60),
      second: secondOfDay % 60,
    };
  }
  const date = new Date(
    excelSystemSerialToUtcMs(wholeDay, date1904) + secondOfDay * 1_000,
  );
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1;
  const day = date.getUTCDate();
  if (!Number.isFinite(date.getTime()) || !isValidYmd(year, month, day)) {
    return null;
  }
  return {
    year,
    month,
    day,
    hour: date.getUTCHours(),
    minute: date.getUTCMinutes(),
    second: date.getUTCSeconds(),
  };
}

/** Which part of a date-time a filter needs the cell's format to show. */
export type FilterDateRequirement = "calendar" | "clock" | "calendar-or-clock";

/**
 * The serial a cell filters by as a date, or NaN when it is not one: its
 * actual number (a literal's, or a formula's computed result such as
 * =DATE(2024,3,1)'s - never the stored characters) under a number format
 * that shows that number as a calendar date and/or a clock time. Text, an
 * error or an unanswered formula has no number and so no date identity, and
 * neither has a number under a non-date format (live audit B459).
 */
export function filterDateSerialForCell(
  numericValue: number | null,
  numberFormat: unknown,
  requirement: FilterDateRequirement = "calendar-or-clock",
): number {
  if (numericValue === null || !Number.isFinite(numericValue)) return Number.NaN;
  if (typeof numberFormat !== "string") return Number.NaN;
  const semantics = classifySpreadsheetNumberFormat(numberFormat, numericValue);
  const matches =
    semantics.valid === true &&
    (requirement === "calendar"
      ? semantics.hasCalendarDate
      : requirement === "clock"
        ? semantics.hasClock
        : semantics.hasCalendarDate || semantics.hasClock);
  return matches ? numericValue : Number.NaN;
}

function sameFilterDate(left: FilterDate, right: FilterDate): boolean {
  return (
    left.year === right.year &&
    left.month === right.month &&
    left.day === right.day &&
    left.hour === right.hour &&
    left.minute === right.minute &&
    left.second === right.second
  );
}

/**
 * Bind localized display text to the actual numeric date serial that produced
 * it. A display shared by different instants is deliberately omitted: a value
 * checkbox cannot identify which native date-group member the user intended.
 */
export function buildFilterDateAuthority(
  entries: readonly { displayValue: string; serial: number }[],
  date1904: ExcelDateSystemArg = false,
): FilterDateAuthority {
  const byDisplay = new Map<string, FilterDate>();
  const ambiguous = new Set<string>();
  let complete = entries.length > 0;
  for (const entry of entries) {
    const displayValue = entry.displayValue;
    const actual = filterDateFromSpreadsheetSerial(entry.serial, date1904);
    if (displayValue === "" || !actual) {
      complete = false;
      continue;
    }
    if (ambiguous.has(displayValue)) {
      complete = false;
      continue;
    }
    const previous = byDisplay.get(displayValue);
    if (previous && !sameFilterDate(previous, actual)) {
      byDisplay.delete(displayValue);
      ambiguous.add(displayValue);
      complete = false;
      continue;
    }
    byDisplay.set(displayValue, actual);
  }
  return { byDisplay, complete };
}

function resolveFilterDate(
  displayValue: string,
  authority?: ReadonlyMap<string, FilterDate>,
): FilterDate | null {
  return authority === undefined
    ? parseFilterDate(displayValue)
    : authority.get(displayValue) ?? null;
}

/** Native table AutoFilter date groups that this editor can round-trip. */
export type FilterDateGroup =
  | { grouping: "year"; year: number }
  | { grouping: "month"; year: number; month: number }
  | { grouping: "day"; year: number; month: number; day: number }
  | {
      grouping: "hour";
      year: number;
      month: number;
      day: number;
      hour: number;
    }
  | {
      grouping: "minute";
      year: number;
      month: number;
      day: number;
      hour: number;
      minute: number;
    }
  | {
      grouping: "second";
      year: number;
      month: number;
      day: number;
      hour: number;
      minute: number;
      second: number;
    };

const MAX_FILTER_DATE_GROUPS = 10_000;

function dateGroupKey(group: FilterDateGroup): string {
  return `${group.grouping}:${group.year}:${
    "month" in group ? group.month : ""
  }:${"day" in group ? group.day : ""}:${
    "hour" in group ? group.hour : ""
  }:${"minute" in group ? group.minute : ""}:${
    "second" in group ? group.second : ""
  }`;
}

/** Strictly normalize the supported Gregorian base-OOXML date-group subset. */
export function normalizeFilterDateGroups(
  raw: unknown,
): FilterDateGroup[] | null {
  if (
    !Array.isArray(raw) ||
    raw.length === 0 ||
    raw.length > MAX_FILTER_DATE_GROUPS
  ) {
    return null;
  }
  const groups: FilterDateGroup[] = [];
  const seen = new Set<string>();
  for (const entry of raw) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) return null;
    const value = entry as Record<string, unknown>;
    const grouping = value.grouping;
    const year = value.year;
    if (
      !["year", "month", "day", "hour", "minute", "second"].includes(
        String(grouping),
      ) ||
      typeof year !== "number" ||
      !Number.isInteger(year) ||
      year < 1 ||
      year > 9999
    ) {
      return null;
    }
    let group: FilterDateGroup;
    const month = value.month;
    const day = value.day;
    const hour = value.hour;
    const minute = value.minute;
    const second = value.second;
    if (grouping === "year") {
      if (Object.keys(value).length !== 2) return null;
      group = { grouping, year };
    } else if (
      typeof month !== "number" ||
      !Number.isInteger(month) ||
      month < 1 ||
      month > 12
    ) {
      return null;
    } else if (grouping === "month") {
      if (Object.keys(value).length !== 3) return null;
      group = { grouping, year, month };
    } else if (
      typeof day !== "number" ||
      !Number.isInteger(day) ||
      !isValidYmd(year, month, day)
    ) {
      return null;
    } else if (grouping === "day") {
      if (Object.keys(value).length !== 4) return null;
      group = { grouping, year, month, day };
    } else if (
      typeof hour !== "number" ||
      !Number.isInteger(hour) ||
      hour < 0 ||
      hour > 23
    ) {
      return null;
    } else if (grouping === "hour") {
      if (Object.keys(value).length !== 5) return null;
      group = { grouping, year, month, day, hour };
    } else if (
      typeof minute !== "number" ||
      !Number.isInteger(minute) ||
      minute < 0 ||
      minute > 59
    ) {
      return null;
    } else if (grouping === "minute") {
      if (Object.keys(value).length !== 6) return null;
      group = { grouping, year, month, day, hour, minute };
    } else if (
      Object.keys(value).length !== 7 ||
      typeof second !== "number" ||
      !Number.isInteger(second) ||
      second < 0 ||
      second > 59
    ) {
      return null;
    } else {
      group = { grouping: "second", year, month, day, hour, minute, second };
    }
    const key = dateGroupKey(group);
    if (seen.has(key)) return null;
    seen.add(key);
    groups.push(group);
  }
  return groups;
}

/** Whether one actual date belongs to one supported native date group. */
export function filterDateMatchesGroup(
  parsed: FilterDate | null,
  group: FilterDateGroup,
): boolean {
  if (!parsed || parsed.year !== group.year) return false;
  if (group.grouping === "year") return true;
  if (parsed.month !== group.month) return false;
  if (group.grouping === "month") return true;
  if (parsed.day !== group.day) return false;
  if (group.grouping === "day") return true;
  if (parsed.hour !== group.hour) return false;
  if (group.grouping === "hour") return true;
  if (parsed.minute !== group.minute) return false;
  if (group.grouping === "minute") return true;
  return parsed.second === group.second;
}

export function filterDateValueMatchesGroup(
  displayValue: string,
  group: FilterDateGroup,
  authority?: ReadonlyMap<string, FilterDate>,
): boolean {
  return filterDateMatchesGroup(resolveFilterDate(displayValue, authority), group);
}

/** Rehydrate native date groups into the menu's existing display-value draft. */
export function selectedValuesForDateGroups(
  options: readonly string[],
  groups: readonly FilterDateGroup[],
  literalValues: readonly string[] = [],
  authority?: ReadonlyMap<string, FilterDate>,
): Set<string> {
  const selected = new Set(literalValues);
  for (const value of options) {
    if (
      groups.some((group) =>
        filterDateValueMatchesGroup(value, group, authority),
      )
    ) {
      selected.add(value);
    }
  }
  return selected;
}

/**
 * Compress selected display values into the narrowest exact native date groups.
 * A partially selected prefix remains literal unless all siblings controlled by
 * that year/month/day/hour/minute/second prefix are selected.
 */
export function compressDateFilterSelection(
  options: readonly string[],
  selected: ReadonlySet<string>,
  authority?: ReadonlyMap<string, FilterDate>,
): { dateGroups: FilterDateGroup[]; values?: string[] } {
  const tree = buildDateFilterTree(options, authority);
  const dateGroups: FilterDateGroup[] = [];
  const values: string[] = [];
  for (const year of tree.years) {
    if (year.values.every((value) => selected.has(value))) {
      dateGroups.push({ grouping: "year", year: year.year });
      continue;
    }
    for (const month of year.months) {
      if (month.values.every((value) => selected.has(value))) {
        dateGroups.push({
          grouping: "month",
          year: year.year,
          month: month.month,
        });
        continue;
      }
      const valuesByDay = new Map<number, string[]>();
      for (const leaf of month.days) {
        const siblings = valuesByDay.get(leaf.day) ?? [];
        siblings.push(leaf.value);
        valuesByDay.set(leaf.day, siblings);
      }
      for (const [day, siblingValues] of valuesByDay) {
        const selectedSiblings = siblingValues.filter((value) =>
          selected.has(value),
        );
        if (selectedSiblings.length === siblingValues.length) {
          dateGroups.push({
            grouping: "day",
            year: year.year,
            month: month.month,
            day,
          });
        } else {
          const parsedSiblings = siblingValues.map((value) => ({
            value,
            parsed: resolveFilterDate(value, authority),
          }));
          const literalDateValues = parsedSiblings
            .filter(({ parsed }) => parsed?.hour === undefined)
            .map(({ value }) => value);
          values.push(
            ...literalDateValues.filter((value) => selected.has(value)),
          );
          const byHour = new Map<number, typeof parsedSiblings>();
          for (const sibling of parsedSiblings) {
            if (sibling.parsed?.hour === undefined) continue;
            const hourSiblings = byHour.get(sibling.parsed.hour) ?? [];
            hourSiblings.push(sibling);
            byHour.set(sibling.parsed.hour, hourSiblings);
          }
          for (const [hour, hourSiblings] of byHour) {
            if (hourSiblings.every(({ value }) => selected.has(value))) {
              dateGroups.push({
                grouping: "hour",
                year: year.year,
                month: month.month,
                day,
                hour,
              });
              continue;
            }
            const byMinute = new Map<number, typeof hourSiblings>();
            for (const sibling of hourSiblings) {
              if (sibling.parsed?.minute === undefined) {
                if (selected.has(sibling.value)) values.push(sibling.value);
                continue;
              }
              const minuteSiblings =
                byMinute.get(sibling.parsed.minute) ?? [];
              minuteSiblings.push(sibling);
              byMinute.set(sibling.parsed.minute, minuteSiblings);
            }
            for (const [minute, minuteSiblings] of byMinute) {
              if (minuteSiblings.every(({ value }) => selected.has(value))) {
                dateGroups.push({
                  grouping: "minute",
                  year: year.year,
                  month: month.month,
                  day,
                  hour,
                  minute,
                });
                continue;
              }
              const bySecond = new Map<number, typeof minuteSiblings>();
              for (const sibling of minuteSiblings) {
                if (sibling.parsed?.second === undefined) {
                  if (selected.has(sibling.value)) values.push(sibling.value);
                  continue;
                }
                const secondSiblings =
                  bySecond.get(sibling.parsed.second) ?? [];
                secondSiblings.push(sibling);
                bySecond.set(sibling.parsed.second, secondSiblings);
              }
              for (const [second, secondSiblings] of bySecond) {
                const selectedSecondValues = secondSiblings.filter(({ value }) =>
                  selected.has(value),
                );
                if (
                  selectedSecondValues.length === secondSiblings.length &&
                  secondSiblings.length > 0
                ) {
                  dateGroups.push({
                    grouping: "second",
                    year: year.year,
                    month: month.month,
                    day,
                    hour,
                    minute,
                    second,
                  });
                } else {
                  values.push(
                    ...selectedSecondValues.map(({ value }) => value),
                  );
                }
              }
            }
          }
        }
      }
    }
  }
  for (const value of tree.nonDates) {
    if (selected.has(value)) values.push(value);
  }
  return {
    dateGroups,
    ...(values.length > 0 ? { values } : {}),
  };
}

const MONTH_NAMES_FULL = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

/** Map a (full or 3+ letter) month name to its 1-based index, else 0. */
function monthNameToIndex(raw: string): number {
  const s = raw.trim().toLowerCase();
  if (s.length < 3) return 0;
  for (let i = 0; i < MONTH_NAMES_FULL.length; i += 1) {
    // Match on the first 3 letters so "Mar"/"March"/"Mar." all resolve.
    if (MONTH_NAMES_FULL[i].slice(0, 3) === s.slice(0, 3)) return i + 1;
  }
  return 0;
}

/** Window a 2-digit year exactly like Excel (00-29 → 2000s, 30-99 → 1900s). */
function windowYear(n: number): number {
  if (n >= 100) return n;
  return n <= 29 ? 2000 + n : 1900 + n;
}

/** True when (y,m,d) is a real calendar date (validates day-of-month + leap). */
function isValidYmd(year: number, month: number, day: number): boolean {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day)
  ) {
    return false;
  }
  if (month < 1 || month > 12 || day < 1 || year < 1 || year > 9999) {
    return false;
  }
  const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  const dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day <= dim[month - 1];
}

const ISO_RE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
const SLASH_RE = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/;
// "Mar 15, 2024" / "March 15 2024" / "Mar. 15, 2024"
const MONTH_FIRST_RE = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{2,4})$/;
// "15 March 2024" / "15 Mar 2024"
const DAY_FIRST_RE = /^(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{2,4})$/;
const TIME_SUFFIX_RE = /^(.*?)(?:T|\s+)(\d{1,2}):(\d{2})(?::(\d{2}))?$/;

/**
 * Parse a column DISPLAY string into a {@link FilterDate}, or null if it is not
 * a recognized date. The string is trimmed first. Ambiguous numeric forms are
 * read month-first (US / Excel default); a value that cannot be a valid date
 * month-first is retried day-first before giving up (so "25/12/2024" still
 * parses as 25 Dec).
 */
export function parseFilterDate(displayValue: string): FilterDate | null {
  const s = (displayValue ?? "").trim();
  if (s === "") return null;

  const timeMatch = TIME_SUFFIX_RE.exec(s);
  const dateText = timeMatch ? timeMatch[1].trim() : s;
  const hour = timeMatch ? Number(timeMatch[2]) : undefined;
  const minute = timeMatch ? Number(timeMatch[3]) : undefined;
  const second = timeMatch?.[4] === undefined ? undefined : Number(timeMatch[4]);
  if (
    (hour !== undefined &&
      (!Number.isInteger(hour) || hour < 0 || hour > 23)) ||
    (minute !== undefined &&
      (!Number.isInteger(minute) || minute < 0 || minute > 59)) ||
    (second !== undefined &&
      (!Number.isInteger(second) || second < 0 || second > 59))
  ) {
    return null;
  }
  const withTime = (date: FilterDate): FilterDate =>
    hour === undefined || minute === undefined
      ? date
      : {
          ...date,
          hour,
          minute,
          ...(second === undefined ? {} : { second }),
        };

  const iso = ISO_RE.exec(dateText);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    return isValidYmd(year, month, day)
      ? withTime({ year, month, day })
      : null;
  }

  const slash = SLASH_RE.exec(dateText);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const year = windowYear(Number(slash[3]));
    // Month-first (US default). Fall back to day-first when that is invalid but
    // the swap is valid (e.g. 25/12/2024).
    if (isValidYmd(year, a, b)) {
      return withTime({ year, month: a, day: b });
    }
    if (isValidYmd(year, b, a)) {
      return withTime({ year, month: b, day: a });
    }
    return null;
  }

  const mf = MONTH_FIRST_RE.exec(dateText);
  if (mf) {
    const month = monthNameToIndex(mf[1]);
    const day = Number(mf[2]);
    const year = windowYear(Number(mf[3]));
    if (month > 0 && isValidYmd(year, month, day)) {
      return withTime({ year, month, day });
    }
    return null;
  }

  const df = DAY_FIRST_RE.exec(dateText);
  if (df) {
    const day = Number(df[1]);
    const month = monthNameToIndex(df[2]);
    const year = windowYear(Number(df[3]));
    if (month > 0 && isValidYmd(year, month, day)) {
      return withTime({ year, month, day });
    }
    return null;
  }

  return null;
}

/**
 * Heuristic: does this column's distinct value list look like a DATE column?
 * True when at least half of the non-blank distinct values parse as dates AND
 * there are at least two parseable dates (so a lone date among labels does not
 * flip a text column into date-grouping mode). Blanks are ignored in the ratio.
 * This mirrors the FilterMenu's numeric-column heuristic for condition filters.
 */
export function isDateColumn(values: readonly string[]): boolean {
  let dateCount = 0;
  let nonBlank = 0;
  for (const v of values) {
    if ((v ?? "").trim() === "") continue;
    nonBlank += 1;
    if (parseFilterDate(v)) dateCount += 1;
  }
  return dateCount >= 2 && dateCount / nonBlank >= 0.5;
}

/** A leaf (a single day) in the date filter tree. */
export interface DateTreeDay {
  day: number;
  hour?: number;
  minute?: number;
  second?: number;
  /** The exact original display string this day maps back to (for the filter). */
  value: string;
}

/** A month node grouping its days. */
export interface DateTreeMonth {
  month: number;
  days: DateTreeDay[];
  /** All leaf display strings under this month (for tri-state + toggling). */
  values: string[];
}

/** A year node grouping its months. */
export interface DateTreeYear {
  year: number;
  months: DateTreeMonth[];
  values: string[];
}

/** The whole grouped tree plus a flat "non-date" bucket. */
export interface DateFilterTree {
  years: DateTreeYear[];
  /** Display strings that did not parse as a date (kept as a flat list). */
  nonDates: string[];
  /** Every leaf display string the tree controls (years + nonDates), de-duped. */
  allValues: string[];
}

/**
 * Build the Year > Month > Day tree from a column's distinct DISPLAY strings
 * (the same list the flat checkbox UI uses, blanks included as ""). Years,
 * months and days are each sorted ascending; the non-date bucket keeps the input
 * order. A display string maps to exactly one (year, month, day) leaf — if two
 * different strings parse to the same calendar day they become two leaves under
 * that day's parent month (so toggling the day controls both), which is correct
 * because the value filter keys on the literal display string.
 */
export function buildDateFilterTree(
  values: readonly string[],
  authority?: ReadonlyMap<string, FilterDate>,
): DateFilterTree {
  const years = new Map<number, Map<number, DateTreeDay[]>>();
  const nonDates: string[] = [];
  const seenNonDate = new Set<string>();

  for (const value of values) {
    const parsed = resolveFilterDate(value, authority);
    if (!parsed) {
      if (!seenNonDate.has(value)) {
        seenNonDate.add(value);
        nonDates.push(value);
      }
      continue;
    }
    let months = years.get(parsed.year);
    if (!months) {
      months = new Map<number, DateTreeDay[]>();
      years.set(parsed.year, months);
    }
    let days = months.get(parsed.month);
    if (!days) {
      days = [];
      months.set(parsed.month, days);
    }
    days.push({
      day: parsed.day,
      ...(parsed.hour === undefined ? {} : { hour: parsed.hour }),
      ...(parsed.minute === undefined ? {} : { minute: parsed.minute }),
      ...(parsed.second === undefined ? {} : { second: parsed.second }),
      value,
    });
  }

  const yearNodes: DateTreeYear[] = [];
  for (const year of [...years.keys()].sort((a, b) => a - b)) {
    const monthsMap = years.get(year)!;
    const monthNodes: DateTreeMonth[] = [];
    const yearValues: string[] = [];
    for (const month of [...monthsMap.keys()].sort((a, b) => a - b)) {
      const days = monthsMap
        .get(month)!
        .slice()
        .sort(
          (a, b) =>
            a.day - b.day ||
            (a.hour ?? -1) - (b.hour ?? -1) ||
            (a.minute ?? -1) - (b.minute ?? -1) ||
            (a.second ?? -1) - (b.second ?? -1) ||
            a.value.localeCompare(b.value),
        );
      const monthValues = days.map((d) => d.value);
      yearValues.push(...monthValues);
      monthNodes.push({ month, days, values: monthValues });
    }
    yearNodes.push({ year, months: monthNodes, values: yearValues });
  }

  const allValues: string[] = [];
  for (const y of yearNodes) allValues.push(...y.values);
  allValues.push(...nonDates);

  return { years: yearNodes, nonDates, allValues };
}

/** Tri-state of a tree node against a set of currently-checked display values. */
export type TriState = "checked" | "unchecked" | "indeterminate";

/**
 * The tri-state of a node given the set of display strings whose leaves are
 * currently checked. A node is "checked" when ALL the values it controls are in
 * the set, "unchecked" when NONE are, and "indeterminate" otherwise. An empty
 * `values` list (a node controlling nothing) reads as "unchecked".
 */
export function nodeTriState(
  values: readonly string[],
  checked: ReadonlySet<string>,
): TriState {
  if (values.length === 0) return "unchecked";
  let on = 0;
  for (const v of values) {
    if (checked.has(v)) on += 1;
  }
  if (on === 0) return "unchecked";
  if (on === values.length) return "checked";
  return "indeterminate";
}

/**
 * Toggle a node: returns a NEW checked-set with every value the node controls
 * either all added or all removed. The direction follows Excel's checkbox: if
 * the node is fully checked it clears (removes all its values); otherwise
 * (unchecked OR indeterminate) it fills (adds all its values). The input set is
 * never mutated.
 */
export function toggleNode(
  values: readonly string[],
  checked: ReadonlySet<string>,
): Set<string> {
  const next = new Set(checked);
  const state = nodeTriState(values, checked);
  if (state === "checked") {
    for (const v of values) next.delete(v);
  } else {
    for (const v of values) next.add(v);
  }
  return next;
}
