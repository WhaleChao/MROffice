/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's 1900 date system, including its fictitious 1900-02-29.
 *
 * Excel numbers days from 1900-01-01 = 1 and, for Lotus 1-2-3 compatibility,
 * treats 1900 as a leap year: serial 60 is "February 29, 1900", a day that
 * never existed, and every real date from 1900-03-01 (serial 61) onward is
 * exactly `days since 1899-12-30`. Only the first 60 serials differ from that
 * plain epoch count, so this module keeps the whole compatibility window in one
 * place:
 *
 *   serial 0   -> "January 0, 1900" (Excel shows 1/0/1900)
 *   serial 1   -> 1900-01-01
 *   serial 59  -> 1900-02-28
 *   serial 60  -> 1900-02-29 (fictitious; typed "2/29/1900" is accepted as 60)
 *   serial 61  -> 1900-03-01, and plain epoch arithmetic from here on
 *
 * The weekday of a serial is pure arithmetic (serial 1 is a Sunday), which is
 * what Excel's WEEKDAY answers for every serial, the window included.
 *
 * Reference: Microsoft Support, "Excel incorrectly assumes that the year 1900
 * is a leap year" and "Differences between the 1900 and the 1904 date system".
 */

export const EXCEL_1900_EPOCH_MS = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

/** The first serial on which the plain 1899-12-30 epoch count is right. */
export const EXCEL_1900_MARCH_FIRST_SERIAL = 61;
/** Excel's fictitious 1900-02-29. */
export const EXCEL_1900_FICTITIOUS_LEAP_SERIAL = 60;

function epochDays(year: number, month0: number, day: number): number {
  return Math.round((Date.UTC(year, month0, day) - EXCEL_1900_EPOCH_MS) / DAY_MS);
}

/**
 * The serial DATE(year, month, day) answers: month and day roll over the way
 * Excel's DATE does (DATE(1900,1,60) = 60, DATE(1900,3,0) = 60), counting the
 * fictitious February 29 in 1900. `year` must already be a full year.
 */
export function excelSerialFromYmd(
  year: number,
  month: number,
  day: number,
): number {
  const firstOfMonth = epochDays(year, month - 1, 1);
  // The first of every month before March 1900 sits one day earlier in Excel's
  // numbering (Jan -> 1, Feb -> 32); March 1900 onward (61+) is unchanged.
  const excelFirst =
    firstOfMonth < EXCEL_1900_MARCH_FIRST_SERIAL ? firstOfMonth - 1 : firstOfMonth;
  return excelFirst + Math.trunc(day) - 1;
}

/**
 * The serial of a typed calendar date that must be real, or null. 1900-02-29
 * is accepted as Excel's serial 60 (Excel takes that entry too); every other
 * date must exist in the Gregorian calendar.
 */
export function excelSerialFromCalendarDate(
  year: number,
  month: number,
  day: number,
): number | null {
  if (month < 1 || month > 12 || day < 1) return null;
  if (year === 1900 && month === 2 && day === 29) {
    return EXCEL_1900_FICTITIOUS_LEAP_SERIAL;
  }
  const back = new Date(Date.UTC(year, month - 1, day));
  if (
    back.getUTCFullYear() !== year ||
    back.getUTCMonth() !== month - 1 ||
    back.getUTCDate() !== day
  ) {
    return null;
  }
  return excelSerialFromYmd(year, month, day);
}

export interface ExcelCalendarParts {
  year: number;
  /** 1..12 */
  month: number;
  /** 0..31; 0 only for serial 0 ("January 0, 1900"). */
  day: number;
  /** 0 = Sunday .. 6 = Saturday, by serial arithmetic. */
  weekday: number;
}

/** The weekday of a whole serial, 0 = Sunday: serial 0 is a Saturday. */
export function excelSerialWeekday(serial: number): number {
  return ((Math.trunc(serial) % 7) + 13) % 7;
}

/**
 * Calendar parts of the whole-day part of a serial as Excel displays it,
 * including serial 0 (1900-01-00) and serial 60 (1900-02-29). Negative serials
 * fall back to plain epoch arithmetic (Excel has no dates there; callers that
 * care reject them before asking).
 */
export function excelSerialToParts(serial: number): ExcelCalendarParts {
  const whole = Math.floor(serial);
  const weekday = excelSerialWeekday(whole);
  if (whole === 0) return { year: 1900, month: 1, day: 0, weekday };
  if (whole === EXCEL_1900_FICTITIOUS_LEAP_SERIAL) {
    return { year: 1900, month: 2, day: 29, weekday };
  }
  const shifted =
    whole > 0 && whole < EXCEL_1900_FICTITIOUS_LEAP_SERIAL ? whole + 1 : whole;
  const date = new Date(EXCEL_1900_EPOCH_MS + shifted * DAY_MS);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday,
  };
}

/**
 * A UTC Date for date arithmetic on a whole serial. Serials 1..59 map to their
 * real January/February 1900 day and serial 0 to 1899-12-31, the day before.
 * The fictitious serial 60 has no real day; it maps to 1900-02-28 so that
 * month arithmetic (EDATE, EOMONTH) stays in February. From 61 on it is the
 * plain epoch count. Use {@link excelSerialToParts} when the displayed
 * year/month/day matters.
 */
export function excelSerialToUtcDate(serial: number): Date {
  return new Date(excelSerialToUtcMs(Math.trunc(serial)));
}

/**
 * The UTC instant of a serial, time of day included: the fractional form of
 * {@link excelSerialToUtcDate} (serial 1.5 is 1900-01-01 12:00, serial 60.25
 * is 1900-02-28 06:00 since 02-29 never existed).
 */
export function excelSerialToUtcMs(serial: number): number {
  const shifted =
    serial >= 0 && serial < EXCEL_1900_FICTITIOUS_LEAP_SERIAL ? serial + 1 : serial;
  return EXCEL_1900_EPOCH_MS + shifted * DAY_MS;
}

/**
 * The serial (time of day as the fraction) of a real UTC instant: the plain
 * 1899-12-30 epoch count from 1900-03-01 on, one day less before it
 * (1900-01-01 00:00 is 1, 1900-02-28 18:00 is 59.75, 1899-12-31 is 0).
 * Instants before 1899-12-31 keep the plain count; Excel has no dates there.
 */
export function excelSerialFromUtcMs(ms: number): number {
  const epochSerial = (ms - EXCEL_1900_EPOCH_MS) / DAY_MS;
  return epochSerial >= 1 && epochSerial < EXCEL_1900_MARCH_FIRST_SERIAL
    ? epochSerial - 1
    : epochSerial;
}

// ---------------------------------------------------------------------------
// The workbook's date system (Excel's "Use 1904 date system").
//
// A workbook reads its serials in one of three systems (SpreadsheetML
// `workbookPr@date1904` and `@dateCompatibility`):
//
// - "1900", the default: the calendar above, fictitious 1900-02-29 included.
// - "1904": serial 0 is 1904-01-01 and every day is the 1900 system's serial
//   minus 1462 (four years and one day): no fictitious 1900-02-29 and no date
//   before 1904. Read through the 1900 functions with the 1462-day offset.
// - "1900-noncompat": the non-backward-compatible 1900 date base a Strict
//   workbook states with `dateCompatibility="0"` (ECMA-376 Part 1, 18.2.28):
//   serial n is plainly 1899-12-30 + n days, with no fictitious leap day.
//   Measured in Excel for Mac (X-236): 0 is 12/30/1899, 1 is 12/31/1899, 2 is
//   1/1/1900, 60 is 2/28/1900, 61 is 3/1/1900; WEEKDAY(0) = 7; a negative
//   serial is no date (#######).
//
// The system decides how a number is READ as a date; it never changes the
// number, so arithmetic is identical in all of them (Microsoft Support, "Date
// systems in Excel").
//
// Callers carry the system as an `ExcelDateSystemArg`: the historical boolean
// (`true` is the 1904 system) or the system itself. Every reader below
// resolves it through `excelDateSystemOf`, the one accessor.

/** The workbook date systems, as page content stores them (`dateSystem`). */
export type SpreadsheetDateSystem = "1900" | "1904" | "1900-noncompat";

/** The non-backward-compatible 1900 date base (Strict `dateCompatibility="0"`). */
export const EXCEL_1900_NONCOMPAT_SYSTEM = "1900-noncompat" as const;

/** A workbook's date system as a calendar reader takes it: `true` (or
 *  "1904") is the 1904 system, "1900-noncompat" the non-compatible 1900 base,
 *  anything else the default 1900 system. */
export type ExcelDateSystemArg = boolean | SpreadsheetDateSystem | undefined;

/** The date-system option a calendar reader or formatter accepts. */
export interface SpreadsheetDateSystemOptions {
  /** The workbook's date system (`true`: 1904). Omitted/false: 1900. */
  date1904?: ExcelDateSystemArg;
}

/** The one accessor: the system a carried value names. */
export function excelDateSystemOf(value: unknown): SpreadsheetDateSystem {
  if (value === true || value === "1904") return "1904";
  if (value === EXCEL_1900_NONCOMPAT_SYSTEM) return EXCEL_1900_NONCOMPAT_SYSTEM;
  return "1900";
}

/** Whether a carried value is the 1904 system. */
export function isExcelDate1904(value: unknown): boolean {
  return excelDateSystemOf(value) === "1904";
}

/** The value a page's stored system is carried as: `true` for 1904 (as every
 *  caller carried it before), the non-compatible base by name, else false. */
export function excelDateSystemArg(system: unknown): ExcelDateSystemArg {
  const resolved = excelDateSystemOf(system);
  return resolved === "1904" ? true : resolved === EXCEL_1900_NONCOMPAT_SYSTEM ? resolved : false;
}

/** Whether serial 60 is Excel's fictitious 1900-02-29 in this system (only
 *  the default 1900 system has it). */
export function excelSystemHasFictitiousLeapDay(value: ExcelDateSystemArg): boolean {
  return excelDateSystemOf(value) === "1900";
}

/** The first whole serial that is a date the system writes: 1 in the 1900
 *  system (0 is "1/0/1900", no day), 0 in the others. */
export function excelMinDateSerial(value: ExcelDateSystemArg): number {
  return excelDateSystemOf(value) === "1900" ? 1 : 0;
}

/** Days between the systems: the 1900 serial of 1904-01-01. */
export const EXCEL_DATE1904_OFFSET_DAYS = 1462;

/** The last serial a system writes (9999-12-31). */
export const EXCEL_1900_MAX_DATE_SERIAL = 2_958_465;
export const EXCEL_1904_MAX_DATE_SERIAL =
  EXCEL_1900_MAX_DATE_SERIAL - EXCEL_DATE1904_OFFSET_DAYS;

/** Whether a stored date-system value is the 1904 system. */
export function isDate1904System(value: unknown): boolean {
  return value === "1904";
}

/** The page-content value for a system (`undefined` keeps 1900 implicit). */
export function normalizeSpreadsheetDateSystem(
  value: unknown,
): SpreadsheetDateSystem | undefined {
  return value === "1904" || value === "1900" || value === EXCEL_1900_NONCOMPAT_SYSTEM ? value : undefined;
}

/** The non-compatible base's serial of a real UTC instant: the plain epoch
 *  count (it equals the 1900 system's from 1900-03-01 on). */
function noncompatSerialFromUtcMs(ms: number): number {
  return (ms - EXCEL_1900_EPOCH_MS) / DAY_MS;
}

/**
 * The 1900-system serial that reads as the same calendar day and time. The
 * non-compatible base's serials before 1900-03-01 have one (one less) from 2
 * on; 0 and 1 (1899-12-30 / -31) have none and map to 0 (callers that show a
 * calendar day ask {@link excelSystemSerialToParts}, which reads them right).
 */
export function excelCalendarSerial(serial: number, date1904?: ExcelDateSystemArg): number {
  const system = excelDateSystemOf(date1904);
  if (system === "1904") return serial + EXCEL_DATE1904_OFFSET_DAYS;
  if (system === EXCEL_1900_NONCOMPAT_SYSTEM) {
    return serial >= 1 && serial < EXCEL_1900_MARCH_FIRST_SERIAL ? Math.max(0, serial - 1) : serial;
  }
  return serial;
}

/** The workbook-system serial of a 1900-system serial. */
export function excelSystemSerial(serial1900: number, date1904?: ExcelDateSystemArg): number {
  const system = excelDateSystemOf(date1904);
  if (system === "1904") return serial1900 - EXCEL_DATE1904_OFFSET_DAYS;
  if (system === EXCEL_1900_NONCOMPAT_SYSTEM) {
    return serial1900 >= 1 && serial1900 < EXCEL_1900_FICTITIOUS_LEAP_SERIAL ? serial1900 + 1 : serial1900;
  }
  return serial1900;
}

/** The last date serial (9999-12-31) of the workbook's system. */
export function excelMaxDateSerial(date1904?: ExcelDateSystemArg): number {
  return isExcelDate1904(date1904) ? EXCEL_1904_MAX_DATE_SERIAL : EXCEL_1900_MAX_DATE_SERIAL;
}

/** Whether a whole serial is a date the system can write (0..9999-12-31). */
export function excelSerialInCalendar(serial: number, date1904?: ExcelDateSystemArg): boolean {
  return serial >= 0 && serial <= excelMaxDateSerial(date1904);
}

/** {@link excelSerialToParts} in the workbook's system (1904: serial 0 is
 *  1904-01-01, a Friday; the non-compatible base: serial 0 is 1899-12-30, a
 *  Saturday, and no 1900-02-29). */
export function excelSystemSerialToParts(
  serial: number,
  date1904?: ExcelDateSystemArg,
): ExcelCalendarParts {
  if (excelDateSystemOf(date1904) === EXCEL_1900_NONCOMPAT_SYSTEM) {
    const whole = Math.floor(serial);
    const date = new Date(EXCEL_1900_EPOCH_MS + whole * DAY_MS);
    return {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
      weekday: excelSerialWeekday(whole),
    };
  }
  return excelSerialToParts(excelCalendarSerial(serial, date1904));
}

/**
 * {@link excelSerialFromYmd} in the workbook's system: a date before
 * 1904-01-01 is a negative 1904 serial (callers reject it). In the
 * non-compatible base the month and day roll over on the real calendar
 * (DATE(1900,1,60) = 61, DATE(1900,2,30) = 62), except that Excel answers
 * #NUM! for the two spellings of the 1900 system's fictitious day,
 * DATE(1900,2,29) and DATE(1900,3,0): NaN here (measured, X-236).
 */
export function excelSystemSerialFromYmd(
  year: number,
  month: number,
  day: number,
  date1904?: ExcelDateSystemArg,
): number {
  if (excelDateSystemOf(date1904) === EXCEL_1900_NONCOMPAT_SYSTEM) {
    const months = year * 12 + (month - 1);
    const rolledYear = Math.floor(months / 12);
    const rolledMonth = months - rolledYear * 12 + 1;
    const wholeDay = Math.trunc(day);
    if (rolledYear === 1900 && ((rolledMonth === 2 && wholeDay === 29) || (rolledMonth === 3 && wholeDay === 0))) {
      return Number.NaN;
    }
    return Math.round((Date.UTC(rolledYear, rolledMonth - 1, 1) - EXCEL_1900_EPOCH_MS) / DAY_MS) + wholeDay - 1;
  }
  return excelSystemSerial(excelSerialFromYmd(year, month, day), date1904);
}

/**
 * {@link excelSerialFromCalendarDate} in the workbook's system: a real date
 * the system can hold, or null. The 1904 system has no 1900-02-29 and no
 * date before 1904-01-01, so a typed "2/29/1900" or "1903-05-05" is not a date
 * there (measured in Excel: both stay text). The non-compatible base has no
 * 1900-02-29 either, and a typed date before 1900 (12/31/1899) stays text
 * though it has a serial (measured, X-236).
 */
export function excelSystemSerialFromCalendarDate(
  year: number,
  month: number,
  day: number,
  date1904?: ExcelDateSystemArg,
): number | null {
  const system = excelDateSystemOf(date1904);
  if (system === "1904" && year < 1904) return null;
  if (system === EXCEL_1900_NONCOMPAT_SYSTEM && (year < 1900 || (year === 1900 && month === 2 && day === 29))) return null;
  const serial = excelSerialFromCalendarDate(year, month, day);
  return serial === null ? null : excelSystemSerial(serial, date1904);
}

/** {@link excelSerialToUtcMs} in the workbook's system. */
export function excelSystemSerialToUtcMs(serial: number, date1904?: ExcelDateSystemArg): number {
  if (excelDateSystemOf(date1904) === EXCEL_1900_NONCOMPAT_SYSTEM) return EXCEL_1900_EPOCH_MS + serial * DAY_MS;
  return excelSerialToUtcMs(excelCalendarSerial(serial, date1904));
}

/** {@link excelSerialToUtcDate} in the workbook's system. */
export function excelSystemSerialToUtcDate(serial: number, date1904?: ExcelDateSystemArg): Date {
  return new Date(excelSystemSerialToUtcMs(Math.trunc(serial), date1904));
}

/** {@link excelSerialFromUtcMs} in the workbook's system. */
export function excelSystemSerialFromUtcMs(ms: number, date1904?: ExcelDateSystemArg): number {
  if (excelDateSystemOf(date1904) === EXCEL_1900_NONCOMPAT_SYSTEM) return noncompatSerialFromUtcMs(ms);
  return excelSystemSerial(excelSerialFromUtcMs(ms), date1904);
}

/** {@link excelSerialWeekday} in the workbook's system: 1904 serial 0 is a
 *  Friday (WEEKDAY(0) is 6 there); the non-compatible base's serial 0 is a
 *  Saturday, as the 1900 system's (WEEKDAY(0) is 7). */
export function excelSystemSerialWeekday(serial: number, date1904?: ExcelDateSystemArg): number {
  if (excelDateSystemOf(date1904) === EXCEL_1900_NONCOMPAT_SYSTEM) return excelSerialWeekday(Math.trunc(serial));
  return excelSerialWeekday(excelCalendarSerial(Math.trunc(serial), date1904));
}
