/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native Excel PivotTable timelines (x15 Timeline views, "日程表").
 *
 * Excel for the web renders an existing timeline and lets the viewer USE it:
 * click or drag contiguous periods, Shift+click to extend, switch the time
 * level (years / quarters / months / days), scroll, and Clear Filter.
 * Inserting a timeline, editing its report connections and formatting it are
 * desktop-only, so the sheet model carries imported timelines as views whose
 * settings are read-only and whose use-state (`selection`, `level`,
 * `selectionLevel`, `scrollPosition`) is the only thing the editor writes.
 *
 * The timeline cache state is authoritative for the selection; each connected
 * PivotTable mirrors it as the `PivotConfig.dateFilters` entry on the
 * timeline's source field, which the pivot engine applies as a whole-day
 * range. Every change
 * therefore rewrites the timeline, every other view of the same cache, and
 * every connected PivotTable in one workbook transaction.
 *
 * Package graph and projection: apps/api/.../xlsx/timeline_slicer.py (server)
 * and pivot-timeline-import.ts (browser import).
 */

import type { PivotDateFilter } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";

export const PIVOT_TIMELINE_LEVELS = [
  "years",
  "quarters",
  "months",
  "days",
] as const;
export type PivotTimelineLevel = (typeof PIVOT_TIMELINE_LEVELS)[number];

export interface PivotTimelineRange {
  /** First day, `YYYY-MM-DD`. */
  start: string;
  /** Selection: last day (inclusive). Bounds: Excel's exclusive upper day. */
  end: string;
}

export interface PivotTimelineConnection {
  sheetName: string;
  pivotName: string;
}

export interface PivotTimelineNativeAuthority {
  version: 1;
  disposition: "editable" | "preserve-only";
  blocker?: string;
  timelinePart: string;
  cachePart: string;
}

export interface PivotTimeline {
  /** Stable identity: the native view name (unique in the workbook). */
  id: string;
  name: string;
  cacheName: string;
  sourceName: string;
  caption: string;
  connections: PivotTimelineConnection[];
  level: PivotTimelineLevel;
  selectionLevel: PivotTimelineLevel;
  scrollPosition?: string;
  selection?: PivotTimelineRange;
  bounds: PivotTimelineRange;
  showHeader: boolean;
  showSelectionLabel: boolean;
  showTimeLevel: boolean;
  showHorizontalScrollbar: boolean;
  style: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  native?: PivotTimelineNativeAuthority;
}

export const MAX_PIVOT_TIMELINES = 64;
export const DEFAULT_PIVOT_TIMELINE_WIDTH = 336;
export const DEFAULT_PIVOT_TIMELINE_HEIGHT = 140;
const MAX_TIMELINE_PERIODS = 4_000;
const DAY_MS = 86_400_000;
const ISO_DAY = /^(\d{4})-(\d{2})-(\d{2})$/;
const CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/;

// --------------------------------------------------------------------------
// Days
// --------------------------------------------------------------------------

interface DayParts {
  year: number;
  month: number;
  day: number;
}

function dayParts(value: string): DayParts | null {
  const match = ISO_DAY.exec(value);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
    ? { year, month, day }
    : null;
}

export function isPivotTimelineDay(value: unknown): value is string {
  return typeof value === "string" && dayParts(value) !== null;
}

function dayNumber(value: string): number {
  const parts = dayParts(value);
  return parts ? Date.UTC(parts.year, parts.month - 1, parts.day) / DAY_MS : NaN;
}

function dayFromNumber(value: number): string {
  return new Date(value * DAY_MS).toISOString().slice(0, 10);
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

function isoFromParts(year: number, month: number, day: number): string {
  return dayFromNumber(Date.UTC(year, month - 1, day) / DAY_MS);
}

// --------------------------------------------------------------------------
// Normalization (render model; the stored object is never rebuilt from it)
// --------------------------------------------------------------------------

function boundedText(value: unknown, max = 255): string | null {
  return typeof value === "string" &&
    value.trim() &&
    value.length <= max &&
    !CONTROL.test(value)
    ? value
    : null;
}

function normalizeRange(value: unknown): PivotTimelineRange | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  if (
    !isPivotTimelineDay(source.start) ||
    !isPivotTimelineDay(source.end) ||
    source.start > source.end
  ) {
    return null;
  }
  return { start: source.start, end: source.end };
}

function normalizeLevel(value: unknown): PivotTimelineLevel | null {
  return PIVOT_TIMELINE_LEVELS.includes(value as PivotTimelineLevel)
    ? (value as PivotTimelineLevel)
    : null;
}

function finiteOrUndefined(value: unknown, min: number, max: number): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value >= min && value <= max
    ? value
    : undefined;
}

/** One stored timeline as a render model, or null when it cannot be shown. */
export function normalizePivotTimeline(value: unknown): PivotTimeline | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const id = boundedText(source.id);
  const cacheName = boundedText(source.cacheName);
  const sourceName = boundedText(source.sourceName);
  const level = normalizeLevel(source.level);
  const bounds = normalizeRange(source.bounds);
  if (!id || !cacheName || !sourceName || !level || !bounds) return null;
  const selection =
    source.selection === undefined || source.selection === null
      ? undefined
      : normalizeRange(source.selection);
  if (selection === null) return null;
  const connections: PivotTimelineConnection[] = [];
  for (const raw of Array.isArray(source.connections) ? source.connections.slice(0, 64) : []) {
    if (!raw || typeof raw !== "object") continue;
    const connection = raw as Record<string, unknown>;
    const sheetName = boundedText(connection.sheetName);
    const pivotName = boundedText(connection.pivotName);
    if (sheetName && pivotName) connections.push({ sheetName, pivotName });
  }
  const rawNative =
    source.native && typeof source.native === "object"
      ? (source.native as Record<string, unknown>)
      : null;
  const disposition =
    rawNative?.disposition === "editable" || rawNative?.disposition === "preserve-only"
      ? rawNative.disposition
      : null;
  const native: PivotTimelineNativeAuthority | undefined =
    rawNative && disposition
      ? {
          version: 1,
          disposition,
          ...(boundedText(rawNative.blocker, 200)
            ? { blocker: boundedText(rawNative.blocker, 200) as string }
            : {}),
          timelinePart: boundedText(rawNative.timelinePart, 512) ?? "",
          cachePart: boundedText(rawNative.cachePart, 512) ?? "",
        }
      : undefined;
  return {
    id,
    name: boundedText(source.name) ?? id,
    cacheName,
    sourceName,
    caption: boundedText(source.caption) ?? sourceName,
    connections,
    level,
    selectionLevel: normalizeLevel(source.selectionLevel) ?? level,
    ...(isPivotTimelineDay(source.scrollPosition)
      ? { scrollPosition: source.scrollPosition }
      : {}),
    ...(selection ? { selection } : {}),
    bounds,
    showHeader: source.showHeader !== false,
    showSelectionLabel: source.showSelectionLabel !== false,
    showTimeLevel: source.showTimeLevel !== false,
    showHorizontalScrollbar: source.showHorizontalScrollbar !== false,
    style: boundedText(source.style, 255) ?? "TimelineStyleLight1",
    ...(finiteOrUndefined(source.x, 0, 1_000_000) !== undefined
      ? { x: source.x as number }
      : {}),
    ...(finiteOrUndefined(source.y, 0, 1_000_000) !== undefined
      ? { y: source.y as number }
      : {}),
    ...(finiteOrUndefined(source.width, 1, 100_000) !== undefined
      ? { width: source.width as number }
      : {}),
    ...(finiteOrUndefined(source.height, 1, 100_000) !== undefined
      ? { height: source.height as number }
      : {}),
    ...(native ? { native } : {}),
  };
}

export function normalizePivotTimelines(value: unknown): PivotTimeline[] {
  if (!Array.isArray(value)) return [];
  const out: PivotTimeline[] = [];
  const ids = new Set<string>();
  for (const raw of value.slice(0, MAX_PIVOT_TIMELINES)) {
    const timeline = normalizePivotTimeline(raw);
    if (!timeline || ids.has(timeline.id.toLowerCase())) continue;
    ids.add(timeline.id.toLowerCase());
    out.push(timeline);
  }
  return out;
}

// --------------------------------------------------------------------------
// Periods
// --------------------------------------------------------------------------

export interface PivotTimelinePeriod {
   
  key: string;
  /** First and last calendar day covered, `YYYY-MM-DD`. */
  start: string;
  end: string;
  year: number;
  /** 1-4 for quarters, 1-12 for months and days. */
  part: number;
  day: number;
  /** Parent band: the year (years/quarters/months) or the month (days). */
  groupKey: string;
}

function periodStart(parts: DayParts, level: PivotTimelineLevel): DayParts {
  switch (level) {
    case "years":
      return { year: parts.year, month: 1, day: 1 };
    case "quarters":
      return {
        year: parts.year,
        month: Math.floor((parts.month - 1) / 3) * 3 + 1,
        day: 1,
      };
    case "months":
      return { year: parts.year, month: parts.month, day: 1 };
    case "days":
      return parts;
  }
}

function nextPeriodStart(parts: DayParts, level: PivotTimelineLevel): DayParts {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  switch (level) {
    case "years":
      date.setUTCFullYear(date.getUTCFullYear() + 1);
      break;
    case "quarters":
      date.setUTCMonth(date.getUTCMonth() + 3);
      break;
    case "months":
      date.setUTCMonth(date.getUTCMonth() + 1);
      break;
    case "days":
      date.setUTCDate(date.getUTCDate() + 1);
      break;
  }
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}

/**
 * The last day a timeline displays. Excel writes the cache bounds with an
 * exclusive upper day (the day after the data, snapped to a period start),
 * the same convention as a pivot cache's `maxDate`.
 */
export function pivotTimelineLastDay(bounds: PivotTimelineRange): string {
  return bounds.end > bounds.start
    ? dayFromNumber(dayNumber(bounds.end) - 1)
    : bounds.end;
}

/** The ordered period tiles between the bounds at one time level. */
export function pivotTimelinePeriods(
  bounds: PivotTimelineRange,
  level: PivotTimelineLevel,
): PivotTimelinePeriod[] {
  const first = dayParts(bounds.start);
  const lastDay = pivotTimelineLastDay(bounds);
  if (!first || !dayParts(lastDay)) return [];
  const out: PivotTimelinePeriod[] = [];
  let cursor = periodStart(first, level);
  const lastNumber = dayNumber(lastDay);
  while (out.length < MAX_TIMELINE_PERIODS) {
    const start = isoFromParts(cursor.year, cursor.month, cursor.day);
    if (dayNumber(start) > lastNumber) break;
    const following = nextPeriodStart(cursor, level);
    const end = dayFromNumber(
      Date.UTC(following.year, following.month - 1, following.day) / DAY_MS - 1,
    );
    const quarter = Math.floor((cursor.month - 1) / 3) + 1;
    out.push({
      key:
        level === "years"
          ? String(cursor.year)
          : level === "quarters"
            ? `${cursor.year}-Q${quarter}`
            : level === "months"
              ? `${cursor.year}-${pad2(cursor.month)}`
              : start,
      start,
      end,
      year: cursor.year,
      part: level === "quarters" ? quarter : cursor.month,
      day: cursor.day,
      groupKey:
        level === "days"
          ? `${cursor.year}-${pad2(cursor.month)}`
          : String(cursor.year),
    });
    cursor = following;
  }
  return out;
}

export type PivotTimelineTileState = "selected" | "partial" | "none";

export function pivotTimelineTileState(
  period: PivotTimelinePeriod,
  selection: PivotTimelineRange | undefined,
): PivotTimelineTileState {
  if (!selection) return "none";
  if (period.end < selection.start || period.start > selection.end) return "none";
  return period.start >= selection.start && period.end <= selection.end
    ? "selected"
    : "partial";
}

/** The inclusive whole-period range between two tiles, in either order. */
export function pivotTimelineRangeForTiles(
  periods: readonly PivotTimelinePeriod[],
  anchor: number,
  focus: number,
): PivotTimelineRange | null {
  const low = periods[Math.min(anchor, focus)];
  const high = periods[Math.max(anchor, focus)];
  return low && high ? { start: low.start, end: high.end } : null;
}

/** First and last tile a selection touches, or null when it touches none. */
export function pivotTimelineSelectedTileSpan(
  periods: readonly PivotTimelinePeriod[],
  selection: PivotTimelineRange | undefined,
): { first: number; last: number } | null {
  if (!selection) return null;
  let first = -1;
  let last = -1;
  periods.forEach((period, index) => {
    if (pivotTimelineTileState(period, selection) === "none") return;
    if (first < 0) first = index;
    last = index;
  });
  return first < 0 ? null : { first, last };
}

/** The tile index the view scrolls to: scrollPosition, else the selection. */
export function pivotTimelineScrollIndex(
  periods: readonly PivotTimelinePeriod[],
  timeline: Pick<PivotTimeline, "scrollPosition" | "selection">,
): number {
  const anchor = timeline.scrollPosition ?? timeline.selection?.start;
  if (!anchor) return 0;
  const index = periods.findIndex((period) => period.end >= anchor);
  return index < 0 ? Math.max(0, periods.length - 1) : index;
}

// --------------------------------------------------------------------------
// Labels
// --------------------------------------------------------------------------

/**
 * The page's translator (`console-notebooks`), for what a calendar cannot
 * name: Excel's quarters, `spreadsheet.pivotTimeline.quarter` ("Q{quarter}")
 * and `spreadsheet.pivotTimeline.quarterYear` ("Q{quarter} {year}"). Years,
 * months and days come from `Intl.DateTimeFormat` in the viewer's locale.
 */
export type PivotTimelineTranslate = (
  key: string,
  values?: Record<string, string | number>,
) => string;

function isChinese(locale: string | undefined): boolean {
  return (locale ?? "").toLowerCase().startsWith("zh");
}

const DATE_FORMATS = new Map<string, Intl.DateTimeFormat>();

/** A Gregorian, UTC formatter for the locale (cached; an unknown tag reads English). */
function dateFormat(
  locale: string | undefined,
  options: Intl.DateTimeFormatOptions,
): Intl.DateTimeFormat {
  const key = `${locale ?? ""}|${JSON.stringify(options)}`;
  let format = DATE_FORMATS.get(key);
  if (!format) {
    const settings = { ...options, calendar: "gregory", timeZone: "UTC" };
    try {
      format = new Intl.DateTimeFormat(locale || "en", settings);
    } catch {
      format = new Intl.DateTimeFormat("en", settings);
    }
    DATE_FORMATS.set(key, format);
  }
  return format;
}

function utcDay(parts: { year: number; month: number; day?: number }): Date {
  return new Date(Date.UTC(parts.year, parts.month - 1, parts.day ?? 1));
}

function formatDay(
  locale: string | undefined,
  options: Intl.DateTimeFormatOptions,
  parts: { year: number; month: number; day?: number },
): string {
  return dateFormat(locale, options).format(utcDay(parts));
}

/** True where the locale writes the year before the month ("2020年1月"). */
function yearComesFirst(locale: string | undefined): boolean {
  const types = dateFormat(locale, { year: "numeric", month: "short" })
    .formatToParts(utcDay({ year: 2020, month: 1 }))
    .map((part) => part.type);
  return types.indexOf("year") < types.indexOf("month");
}

function quarterText(
  quarter: number,
  year: number | null,
  locale: string | undefined,
  translate?: PivotTimelineTranslate,
): string {
  if (translate) {
    return year === null
      ? translate("spreadsheet.pivotTimeline.quarter", { quarter })
      : translate("spreadsheet.pivotTimeline.quarterYear", {
          quarter,
          year: String(year),
        });
  }
  const zh = isChinese(locale);
  if (year === null) return zh ? `第${quarter}季度` : `Q${quarter}`;
  return zh ? `${year}年第${quarter}季度` : `Q${quarter} ${year}`;
}

/** Label of one tile (Excel: `2020`, `Q1`, `JAN`, `5`). */
export function pivotTimelineTileLabel(
  period: PivotTimelinePeriod,
  level: PivotTimelineLevel,
  locale?: string,
  translate?: PivotTimelineTranslate,
): string {
  switch (level) {
    case "years":
      return formatDay(
        locale,
        { year: "numeric" },
        { year: period.year, month: 1 },
      );
    case "quarters":
      return quarterText(period.part, null, locale, translate);
    case "months":
      return formatDay(
        locale,
        { month: "short" },
        { year: period.year, month: period.part },
      ).toLocaleUpperCase(locale || "en");
    case "days":
      return String(period.day);
  }
}

/** Label of the band above the tiles (the year, or the month for days). */
export function pivotTimelineGroupLabel(
  period: PivotTimelinePeriod,
  level: PivotTimelineLevel,
  locale?: string,
): string {
  if (level === "days") {
    return formatDay(
      locale,
      { year: "numeric", month: "short" },
      { year: period.year, month: period.part },
    ).toLocaleUpperCase(locale || "en");
  }
  return formatDay(
    locale,
    { year: "numeric" },
    { year: period.year, month: 1 },
  );
}

function periodText(
  parts: DayParts,
  level: PivotTimelineLevel,
  locale: string | undefined,
  withYear: boolean,
  translate?: PivotTimelineTranslate,
): string {
  const quarter = Math.floor((parts.month - 1) / 3) + 1;
  switch (level) {
    case "years":
      return formatDay(locale, { year: "numeric" }, parts);
    case "quarters":
      return quarterText(
        quarter,
        withYear ? parts.year : null,
        locale,
        translate,
      );
    case "months":
      return formatDay(
        locale,
        withYear ? { year: "numeric", month: "short" } : { month: "short" },
        parts,
      );
    case "days":
      return formatDay(
        locale,
        withYear
          ? { year: "numeric", month: "short", day: "numeric" }
          : { month: "short", day: "numeric" },
        parts,
      );
  }
}

/**
 * The selection label Excel shows under the caption: "All Periods" when no
 * filter, else the selected span at its selection level ("Jan - Mar 2020",
 * "Q4 2019 - Q2 2020", "2019 - 2020", "Jan 5 - 9, 2020"). The year stands
 * once, where the locale writes it: last in English, first in Chinese
 * ("2020年1月 - 3月").
 */
export function pivotTimelineSelectionLabel(
  selection: PivotTimelineRange | undefined,
  level: PivotTimelineLevel,
  allPeriodsLabel: string,
  locale?: string,
  translate?: PivotTimelineTranslate,
): string {
  if (!selection) return allPeriodsLabel;
  const start = dayParts(selection.start);
  const end = dayParts(selection.end);
  if (!start || !end) return allPeriodsLabel;
  const text = (parts: DayParts, withYear: boolean) =>
    periodText(parts, level, locale, withYear, translate);
  const endStart = periodStart(end, level);
  const samePeriod =
    periodStart(start, level).year === endStart.year &&
    periodStart(start, level).month === endStart.month &&
    periodStart(start, level).day === endStart.day;
  if (samePeriod) return text(start, true);
  const sameYear = start.year === end.year;
  if (level === "years") return `${text(start, true)} - ${text(end, true)}`;
  const yearFirst = yearComesFirst(locale);
  if (level === "days" && sameYear && start.month === end.month) {
    if (yearFirst) {
      return `${text(start, true)} - ${formatDay(locale, { day: "numeric" }, end)}`;
    }
    // One month, one year: the locale's own day range ("Jan 5 - 9, 2020",
    // "5 - 9 Jan 2020"), with Excel's spaced hyphen between the days.
    return dateFormat(locale, {
      year: "numeric",
      month: "short",
      day: "numeric",
    })
      .formatRange(utcDay(start), utcDay(end))
      .replace(/\s*[–—-]\s*/u, " - ");
  }
  return yearFirst
    ? `${text(start, true)} - ${text(end, !sameYear)}`
    : `${text(start, !sameYear)} - ${text(end, true)}`;
}

// --------------------------------------------------------------------------
// Styles
// --------------------------------------------------------------------------

export interface PivotTimelineStyleRecipe {
  surface: string;
  border: string;
  headerText: string;
  labelText: string;
  tile: string;
  tileText: string;
  selected: string;
  selectedText: string;
  partial: string;
  scrollbar: string;
}

/** Office 2013-2022 theme accents, used when the workbook theme is unknown. */
const DEFAULT_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

function mix(hex: string, other: string, weight: number): string {
  const parse = (value: string) => {
    const match = /^#?([0-9a-f]{6})$/i.exec(value);
    const number = match ? Number.parseInt(match[1], 16) : 0x808080;
    return [(number >> 16) & 255, (number >> 8) & 255, number & 255];
  };
  const a = parse(hex);
  const b = parse(other);
  const channel = (index: number) =>
    Math.round(a[index] * (1 - weight) + b[index] * weight)
      .toString(16)
      .padStart(2, "0");
  return `#${channel(0)}${channel(1)}${channel(2)}`;
}

/**
 * Built-in TimelineStyleLight1-6 / TimelineStyleDark1-6 on the workbook's
 * theme accents. A custom workbook style keeps its name in the file and
 * renders as TimelineStyleLight1.
 */
export function pivotTimelineStyleRecipe(
  style: string,
  accents?: readonly string[],
): PivotTimelineStyleRecipe {
  const match = /^TimelineStyle(Light|Dark)([1-6])$/.exec(style);
  const dark = match?.[1] === "Dark";
  const index = match ? Number(match[2]) - 1 : 0;
  const accent =
    (accents?.[index] && /^#?[0-9a-f]{6}$/i.test(accents[index])
      ? `#${accents[index].replace(/^#/, "")}`
      : DEFAULT_ACCENTS[index]) ?? DEFAULT_ACCENTS[0];
  return dark
    ? {
        surface: "#262626",
        border: mix(accent, "#000000", 0.25),
        headerText: "#ffffff",
        labelText: "#d9d9d9",
        tile: "#404040",
        tileText: "#f2f2f2",
        selected: accent,
        selectedText: "#ffffff",
        partial: mix(accent, "#262626", 0.55),
        scrollbar: "#595959",
      }
    : {
        surface: "#ffffff",
        border: "#bfbfbf",
        headerText: "#262626",
        labelText: "#595959",
        tile: "#e7e6e6",
        tileText: "#404040",
        selected: accent,
        selectedText: "#ffffff",
        partial: mix(accent, "#ffffff", 0.6),
        scrollbar: "#c8c8c8",
      };
}

// --------------------------------------------------------------------------
// Workbook transaction
// --------------------------------------------------------------------------

export interface PivotTimelineSheetLike {
  id: string;
  name: string;
  pivot?: unknown;
  /** The sheet's other PivotTables (a sheet may hold several, as Excel's do). */
  morePivots?: unknown;
  timelines?: unknown;
}

export type PivotTimelineBlock =
  | "native"
  | "unconnected"
  | "connection"
  | "pivot";

interface ConnectedPivot {
  sheetIndex: number;
  /** The PivotTable's place among its sheet's PivotTables (0 = `pivot`). */
  slot: number;
  colIndex: number;
}

/**
 * Every PivotTable of a sheet, in slot order: `pivot`, then `morePivots`. A
 * timeline connects to PivotTables by sheet and name, wherever they sit on
 * the sheet (Excel's Report Connections list each PivotTable on one cache).
 */
function pivotRecords(sheet: PivotTimelineSheetLike): Array<Record<string, unknown> | null> {
  const record = (value: unknown) =>
    value && typeof value === "object" ? (value as Record<string, unknown>) : null;
  return [
    record(sheet.pivot),
    ...(Array.isArray(sheet.morePivots) ? sheet.morePivots.map(record) : []),
  ];
}

/** `sheet` with the PivotTable in `slot` replaced. */
function withPivotAt<T extends PivotTimelineSheetLike>(
  sheet: T,
  slot: number,
  pivot: Record<string, unknown>,
): T {
  if (slot === 0) return { ...sheet, pivot };
  const more = Array.isArray(sheet.morePivots) ? [...sheet.morePivots] : [];
  more[slot - 1] = pivot;
  return { ...sheet, morePivots: more };
}

function nativeRecord(pivot: Record<string, unknown>): Record<string, unknown> | null {
  return pivot.native && typeof pivot.native === "object"
    ? (pivot.native as Record<string, unknown>)
    : null;
}

/** A connection names the PivotTable as imported; a rename keeps native.pivotName. */
function answersToPivotName(pivot: Record<string, unknown>, name: string): boolean {
  const wanted = name.toLowerCase();
  return [pivot.name, nativeRecord(pivot)?.pivotName].some(
    (candidate) => typeof candidate === "string" && candidate.toLowerCase() === wanted,
  );
}

function sourceColumn(pivot: Record<string, unknown>, sourceName: string): number | null {
  const fields = nativeRecord(pivot)?.fields;
  if (!Array.isArray(fields)) return null;
  const matches = fields.filter(
    (field) =>
      field &&
      typeof field === "object" &&
      (field as Record<string, unknown>).name === sourceName &&
      Number.isInteger((field as Record<string, unknown>).sourceIndex),
  );
  return matches.length === 1
    ? ((matches[0] as Record<string, unknown>).sourceIndex as number)
    : null;
}

function resolveConnections(
  sheets: readonly PivotTimelineSheetLike[],
  timeline: PivotTimeline,
  isPivotEditable: (pivot: unknown) => boolean,
): { pivots: ConnectedPivot[]; block?: PivotTimelineBlock } {
  if (timeline.native?.disposition !== "editable") return { pivots: [], block: "native" };
  if (timeline.connections.length === 0) return { pivots: [], block: "unconnected" };
  const pivots: ConnectedPivot[] = [];
  for (const connection of timeline.connections) {
    const matches = sheets.flatMap((sheet, sheetIndex) =>
      sheet.name === connection.sheetName
        ? pivotRecords(sheet).flatMap((pivot, slot) =>
            pivot && answersToPivotName(pivot, connection.pivotName)
              ? [{ sheetIndex, slot, pivot }]
              : [],
          )
        : [],
    );
    if (matches.length !== 1) return { pivots: [], block: "connection" };
    const { sheetIndex, slot, pivot } = matches[0];
    if (pivots.some((entry) => entry.sheetIndex === sheetIndex && entry.slot === slot)) {
      continue;
    }
    const colIndex = sourceColumn(pivot, timeline.sourceName);
    if (colIndex === null) return { pivots: [], block: "connection" };
    if (!isPivotEditable(pivot)) return { pivots: [], block: "pivot" };
    pivots.push({ sheetIndex, slot, colIndex });
  }
  return { pivots };
}

/** Why a timeline cannot be used right now, or null when it can. */
export function pivotTimelineBlock(
  sheets: readonly PivotTimelineSheetLike[],
  timeline: PivotTimeline,
  isPivotEditable: (pivot: unknown) => boolean,
): PivotTimelineBlock | null {
  return resolveConnections(sheets, timeline, isPivotEditable).block ?? null;
}

export interface PivotTimelineChange {
  selection: PivotTimelineRange | null;
  selectionLevel?: PivotTimelineLevel;
  level?: PivotTimelineLevel;
  scrollPosition?: string;
}

function sameRange(
  left: PivotTimelineRange | undefined | null,
  right: PivotTimelineRange | undefined | null,
): boolean {
  return (left?.start ?? null) === (right?.start ?? null) &&
    (left?.end ?? null) === (right?.end ?? null);
}

function withTimelineState(
  raw: Record<string, unknown>,
  change: PivotTimelineChange,
  isTarget: boolean,
): Record<string, unknown> {
  const next: Record<string, unknown> = { ...raw };
  if (change.selection) {
    next.selection = { start: change.selection.start, end: change.selection.end };
  } else {
    delete next.selection;
  }
  if (change.selectionLevel) next.selectionLevel = change.selectionLevel;
  if (isTarget && change.level) next.level = change.level;
  if (isTarget && change.scrollPosition && isPivotTimelineDay(change.scrollPosition)) {
    next.scrollPosition = change.scrollPosition;
  }
  return next;
}

/** A filter name no kept filter uses (the reader requires distinct names). */
function freeFilterName(wanted: string, kept: readonly PivotDateFilter[]): string {
  const taken = new Set(kept.map((item) => item.timeline?.toLowerCase()));
  let name = wanted;
  for (let suffix = 2; taken.has(name.toLowerCase()); suffix += 1) {
    name = `${wanted} ${suffix}`;
  }
  return name;
}

/**
 * Excel links a timeline to the date filter on its source field, whatever the
 * filter's name; a new filter is named after the field, as Excel names it.
 */
function nextDateFilters(
  pivot: Record<string, unknown>,
  sourceName: string,
  colIndex: number,
  selection: PivotTimelineRange | null,
): PivotDateFilter[] {
  const current = Array.isArray(pivot.dateFilters)
    ? (pivot.dateFilters as PivotDateFilter[])
    : [];
  const previous = current.find((item) => item.colIndex === colIndex);
  const kept = current.filter((item) => item.colIndex !== colIndex);
  if (!selection) return kept;
  const receipts = nativeRecord(pivot)?.dateFilters;
  const receipt = Array.isArray(receipts)
    ? (receipts as PivotDateFilter[]).find((item) => item.colIndex === colIndex)
    : undefined;
  const usedIds = [...current, ...(Array.isArray(receipts) ? (receipts as PivotDateFilter[]) : [])]
    .map((item) => item.filterId)
    .filter((id): id is number => Number.isInteger(id));
  const filterId =
    previous?.filterId ?? receipt?.filterId ?? Math.max(0, ...usedIds) + 1;
  const evalOrder = previous?.evalOrder ?? receipt?.evalOrder ?? -1;
  return [
    ...kept,
    {
      colIndex,
      start: selection.start,
      end: selection.end,
      timeline: previous?.timeline ?? receipt?.timeline ?? freeFilterName(sourceName, kept),
      filterId,
      evalOrder,
    },
  ];
}

/**
 * Apply one timeline use-state change to the workbook: the target view, every
 * other view of the same cache (they share one filter state) and every
 * connected PivotTable's `dateFilters`. Returns null when the timeline cannot
 * be used or nothing would change.
 */
export function applyPivotTimelineChange<T extends PivotTimelineSheetLike>(
  sheets: readonly T[],
  target: { sheetId: string; timelineId: string },
  change: PivotTimelineChange,
  isPivotEditable: (pivot: unknown) => boolean,
): T[] | null {
  const sheet = sheets.find((candidate) => candidate.id === target.sheetId);
  const stored = Array.isArray(sheet?.timelines) ? (sheet?.timelines as unknown[]) : [];
  const raw = stored.find(
    (entry) =>
      entry &&
      typeof entry === "object" &&
      String((entry as Record<string, unknown>).id ?? "").toLowerCase() ===
        target.timelineId.toLowerCase(),
  ) as Record<string, unknown> | undefined;
  const timeline = raw ? normalizePivotTimeline(raw) : null;
  if (!sheet || !raw || !timeline) return null;
  if (
    change.selection &&
    (!isPivotTimelineDay(change.selection.start) ||
      !isPivotTimelineDay(change.selection.end) ||
      change.selection.start > change.selection.end)
  ) {
    return null;
  }
  const resolved = resolveConnections(sheets, timeline, isPivotEditable);
  if (resolved.block) return null;
  const selectionChanged = !sameRange(timeline.selection, change.selection);
  const levelChanged = Boolean(change.level && change.level !== timeline.level);
  const selectionLevelChanged = Boolean(
    change.selectionLevel && change.selectionLevel !== timeline.selectionLevel,
  );
  const scrollChanged = Boolean(
    change.scrollPosition && change.scrollPosition !== timeline.scrollPosition,
  );
  if (!selectionChanged && !levelChanged && !selectionLevelChanged && !scrollChanged) {
    return null;
  }
  const cacheKey = timeline.cacheName.toLowerCase();
  let changed = false;
  const next = sheets.map((candidate, sheetIndex) => {
    let updated: T = candidate;
    if (Array.isArray(candidate.timelines)) {
      let touched = false;
      const views = (candidate.timelines as unknown[]).map((entry) => {
        if (!entry || typeof entry !== "object") return entry;
        const view = entry as Record<string, unknown>;
        if (String(view.cacheName ?? "").toLowerCase() !== cacheKey) return entry;
        const isTarget = candidate.id === target.sheetId && view === raw;
        if (!isTarget && !selectionChanged) return entry;
        touched = true;
        return withTimelineState(view, change, isTarget);
      });
      if (touched) updated = { ...updated, timelines: views };
    }
    for (const connected of resolved.pivots) {
      if (connected.sheetIndex !== sheetIndex || !selectionChanged) continue;
      const pivot = pivotRecords(updated)[connected.slot] as Record<string, unknown>;
      const dateFilters = nextDateFilters(
        pivot,
        timeline.sourceName,
        connected.colIndex,
        change.selection,
      );
      const nextPivot: Record<string, unknown> = { ...pivot };
      if (dateFilters.length > 0) nextPivot.dateFilters = dateFilters;
      else delete nextPivot.dateFilters;
      updated = withPivotAt(updated, connected.slot, nextPivot);
    }
    if (updated !== candidate) changed = true;
    return updated;
  });
  return changed ? next : null;
}

/** Follow a worksheet rename into every timeline's report connections. */
export function renamePivotTimelineConnections<T extends PivotTimelineSheetLike>(
  sheet: T,
  oldName: string,
  newName: string,
): T {
  if (!Array.isArray(sheet.timelines) || oldName === newName) return sheet;
  let changed = false;
  const timelines = (sheet.timelines as unknown[]).map((entry) => {
    if (!entry || typeof entry !== "object") return entry;
    const view = entry as Record<string, unknown>;
    if (!Array.isArray(view.connections)) return entry;
    let touched = false;
    const connections = (view.connections as unknown[]).map((connection) => {
      if (
        connection &&
        typeof connection === "object" &&
        (connection as Record<string, unknown>).sheetName === oldName
      ) {
        touched = true;
        return { ...(connection as Record<string, unknown>), sheetName: newName };
      }
      return connection;
    });
    if (!touched) return entry;
    changed = true;
    return { ...view, connections };
  });
  return changed ? { ...sheet, timelines } : sheet;
}
