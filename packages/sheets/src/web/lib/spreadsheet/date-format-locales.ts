/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Format Cells > Number > Date / Time: "Locale (location)" and "Calendar
// type" (XLS-A055).
//
// Microsoft Support, "Format a date the way you want" / "Format numbers as
// dates or times": the Type list shows the dates of the chosen locale; a type
// from another locale is written with that locale's [$-LCID] tag, and a
// calendar type other than Gregorian is written in the tag's calendar byte
// ([$-NCCLLLL]: N numerals, CC calendar, LLLL language). The workbook keeps
// only the code, so the two dropdowns are read back from it.
//
// The Type lists of the UI's own region are Excel's curated ones
// (format-code-builder); every other locale's list is built from the
// browser's own patterns for that locale (Intl), each written with the
// locale's tag, so the list is that locale's real date and time shapes.
// ---------------------------------------------------------------------------

import type { NumberFormatTypeEntry } from "@mrnote/core/web/lib/spreadsheet/format-code-builder.ts";

/** The locales offered, as Excel's language ids. */
export const DATE_FORMAT_LOCALES: ReadonlyArray<{ lid: number; tag: string }> = [
  { lid: 0x0804, tag: "zh-CN" },
  { lid: 0x0404, tag: "zh-TW" },
  { lid: 0x0c04, tag: "zh-HK" },
  { lid: 0x0409, tag: "en-US" },
  { lid: 0x0809, tag: "en-GB" },
  { lid: 0x0411, tag: "ja-JP" },
  { lid: 0x0412, tag: "ko-KR" },
  { lid: 0x0407, tag: "de-DE" },
  { lid: 0x040c, tag: "fr-FR" },
  { lid: 0x0c0a, tag: "es-ES" },
  { lid: 0x0410, tag: "it-IT" },
  { lid: 0x0416, tag: "pt-BR" },
  { lid: 0x0413, tag: "nl-NL" },
  { lid: 0x041d, tag: "sv-SE" },
  { lid: 0x0406, tag: "da-DK" },
  { lid: 0x0414, tag: "nb-NO" },
  { lid: 0x040b, tag: "fi-FI" },
  { lid: 0x0419, tag: "ru-RU" },
  { lid: 0x0415, tag: "pl-PL" },
  { lid: 0x041f, tag: "tr-TR" },
  { lid: 0x0421, tag: "id-ID" },
  { lid: 0x041e, tag: "th-TH" },
  { lid: 0x042a, tag: "vi-VN" },
  { lid: 0x0401, tag: "ar-SA" },
  { lid: 0x040d, tag: "he-IL" },
];

/** Excel's calendar ids (the tag's CC byte) a locale offers beside Gregorian. */
export type DateCalendarId = 1 | 3 | 4 | 5 | 6 | 7 | 8 | 0x17;
const LOCALE_CALENDARS: Record<number, readonly DateCalendarId[]> = {
  0x0411: [3],
  0x0404: [4],
  0x0412: [5],
  0x0401: [6, 0x17],
  0x041e: [7],
  0x040d: [8],
};
const INTL_CALENDAR: Partial<Record<DateCalendarId, string>> = {
  3: "japanese",
  4: "roc",
  6: "islamic",
  7: "buddhist",
  8: "hebrew",
  0x17: "islamic-umalqura",
};

/** Gregorian first, then the locale's own calendars. */
export function calendarOptionsFor(lid: number): DateCalendarId[] {
  return [1, ...(LOCALE_CALENDARS[lid] ?? [])];
}

/** The UI region's locale id: zh-CN for Chinese, en-US otherwise. */
export function regionLocaleId(locale: string): number {
  const tag = locale.toLowerCase();
  if (tag === "zh-tw" || tag === "zh-hant") return 0x0404;
  return tag.startsWith("zh") ? 0x0804 : 0x0409;
}

/** The locale and calendar a date / time code is written for. */
export function dateCodeLocale(code: string, fallbackLid: number): { lid: number; calendar: DateCalendarId } {
  const tag = /\[\$[^\]]*?-([0-9a-f]{3,8})\]/i.exec(code)?.[1];
  if (!tag || /^f[48]00$/i.test(tag)) return { lid: fallbackLid, calendar: 1 };
  const value = Number.parseInt(tag, 16);
  const lid = value & 0xffff;
  const calendarId = (value >> 16) & 0xff;
  const known = DATE_FORMAT_LOCALES.some((entry) => entry.lid === lid) ? lid : fallbackLid;
  const calendar = (calendarOptionsFor(known) as number[]).includes(calendarId)
    ? (calendarId as DateCalendarId)
    : 1;
  return { lid: known, calendar };
}

/** The `[$-…]` tag Excel writes for a locale and calendar. */
export function dateLocaleTag(lid: number, calendar: DateCalendarId): string {
  const value = calendar === 1 ? lid : (calendar << 16) | lid;
  return `[$-${value.toString(16).toUpperCase()}]`;
}



// sample text.
const PROBE = new Date(Date.UTC(2012, 2, 4, 13, 5, 9));

function quoteLiteral(text: string): string {
  if (text === "") return "";
  // Separators Excel writes bare in a date code.
  if (/^[\s/\-.,:]+$/.test(text)) return text;
  return `"${text.replace(/"/g, "")}"`;
}

/** An Intl pattern for the locale as an Excel date / time picture, or null. */
function excelPicture(
  locale: string,
  options: Intl.DateTimeFormatOptions,
  calendar: DateCalendarId,
): string | null {
  const intlCalendar = INTL_CALENDAR[calendar];
  let parts: Intl.DateTimeFormatPart[];
  try {
    parts = new Intl.DateTimeFormat(intlCalendar ? `${locale}-u-ca-${intlCalendar}` : locale, {
      ...options,
      timeZone: "UTC",
    }).formatToParts(PROBE);
  } catch {
    return null;
  }
  const eraYear = calendar === 3 || calendar === 4;
  let out = "";
  for (const part of parts) {
    switch (part.type) {
      case "year":
        out += eraYear ? (part.value.length >= 2 && /^0/.test(part.value) ? "ee" : "e") : part.value.length <= 2 ? "yy" : "yyyy";
        break;
      case "month":
        if (/^\d+$/.test(part.value)) out += part.value.length === 2 ? "mm" : "m";
        else out += options.month === "short" ? "mmm" : "mmmm";
        break;
      case "day":
        out += part.value.length === 2 ? "dd" : "d";
        break;
      case "weekday":
        out += options.weekday === "short" ? "ddd" : "dddd";
        break;
      case "era":
        out += "ggg";
        break;
      case "hour":
        out += part.value.length === 2 ? "hh" : "h";
        break;
      case "minute":
        out += "mm";
        break;
      case "second":
        out += "ss";
        break;
      case "dayPeriod":
        out += "AM/PM";
        break;
      case "literal":
        out += quoteLiteral(part.value);
        break;
      default:
        return null;
    }
  }
  return out;
}

const DATE_SHAPES: Intl.DateTimeFormatOptions[] = [
  { dateStyle: "short" },
  { dateStyle: "medium" },
  { dateStyle: "long" },
  { dateStyle: "full" },
  { year: "numeric", month: "long" },
  { month: "long", day: "numeric" },
  { year: "2-digit", month: "2-digit", day: "2-digit" },
  { year: "numeric", month: "short", day: "numeric" },
  { weekday: "long" },
];
const TIME_SHAPES: Intl.DateTimeFormatOptions[] = [
  { hour: "numeric", minute: "2-digit" },
  { hour: "numeric", minute: "2-digit", second: "2-digit" },
  { hour: "numeric", minute: "2-digit", hour12: true },
  { hour: "numeric", minute: "2-digit", second: "2-digit", hour12: true },
  { hour: "numeric", minute: "2-digit", hour12: false },
];

function builtList(
  lid: number,
  calendar: DateCalendarId,
  shapes: Intl.DateTimeFormatOptions[],
): NumberFormatTypeEntry[] {
  const locale = DATE_FORMAT_LOCALES.find((entry) => entry.lid === lid)?.tag ?? "en-US";
  const tag = dateLocaleTag(lid, calendar);
  const seen = new Set<string>();
  const list: NumberFormatTypeEntry[] = [];
  for (const shape of shapes) {
    const picture = excelPicture(locale, shape, calendar);
    if (!picture || seen.has(picture)) continue;
    seen.add(picture);
    list.push({ code: `${tag}${picture};@` });
  }
  return list;
}

/**
 * The Date Type list for a locale and calendar: Excel's curated list for the
 * UI's own region in the Gregorian calendar, else the locale's own shapes.
 */
export function dateTypesFor(
  lid: number,
  calendar: DateCalendarId,
  regionLid: number,
  regionList: readonly NumberFormatTypeEntry[],
): readonly NumberFormatTypeEntry[] {
  if (lid === regionLid && calendar === 1) return regionList;
  return builtList(lid, calendar, DATE_SHAPES);
}

/** The Time Type list for a locale (time has no calendar). */
export function timeTypesFor(
  lid: number,
  regionLid: number,
  regionList: readonly NumberFormatTypeEntry[],
): readonly NumberFormatTypeEntry[] {
  if (lid === regionLid) return regionList;
  return builtList(lid, 1, TIME_SHAPES);
}
