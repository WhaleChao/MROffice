/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  EXCEL_1900_EPOCH_MS,
  EXCEL_1900_MARCH_FIRST_SERIAL,
  excelCalendarSerial,
  excelMaxDateSerial,
  excelSerialToParts,
  excelSystemHasFictitiousLeapDay,
  isExcelDate1904,
  type ExcelDateSystemArg,
  type SpreadsheetDateSystemOptions,
} from "./excel-1900-calendar";
import { buildAccountingCodeFromParams } from "./format-code-builder";

// Spreadsheet date serials follow Excel's 1900 date system: day 1 = 1900-01-01
// and serial 60 is the fictitious 1900-02-29, so the 1899-12-30 epoch count is
// only exact from serial 61 (1900-03-01) on; the window before it is read
// through excel-1900-calendar, matching the formula engine and entry parser.
const DATE_EPOCH_MS = EXCEL_1900_EPOCH_MS;
const MS_PER_DAY = 86_400_000;
const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const WEEKDAYS = [
  "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
];

/**
 * The language of a `[$-LID]` / `[$€-LID]` locale tag (the low 16 bits of
 * Excel's LCID; the high bits carry calendar and digit flags) as a BCP 47 tag,
 * for the month and weekday names a date picture spells out. English (the
 * default, and the 0409 family) keeps the built-in tables.
 */
const LID_LOCALES: Record<number, string> = {
  0x0401: "ar-SA", 0x0404: "zh-TW", 0x0405: "cs-CZ", 0x0406: "da-DK", 0x0407: "de-DE",
  0x0408: "el-GR", 0x040a: "es-ES", 0x040b: "fi-FI", 0x040c: "fr-FR", 0x040d: "he-IL",
  0x040e: "hu-HU", 0x0410: "it-IT", 0x0411: "ja-JP", 0x0412: "ko-KR", 0x0413: "nl-NL",
  0x0414: "nb-NO", 0x0415: "pl-PL", 0x0416: "pt-BR", 0x0418: "ro-RO", 0x0419: "ru-RU",
  0x041a: "hr-HR", 0x041b: "sk-SK", 0x041d: "sv-SE", 0x041e: "th-TH", 0x041f: "tr-TR",
  0x0421: "id-ID", 0x0422: "uk-UA", 0x042a: "vi-VN", 0x0804: "zh-CN", 0x0807: "de-CH",
  0x080a: "es-MX", 0x080c: "fr-BE", 0x0816: "pt-PT", 0x0c04: "zh-HK", 0x0c07: "de-AT",
  0x0c0a: "es-ES", 0x0c0c: "fr-CA", 0x100c: "fr-CH", 0x1404: "zh-MO",
  // More of Windows' language ids (MS-LCID), so a `[$-LID]` date in another
  // language names its months and days in it (S08_tdf73063's [$-1C1A]
  // Serbian Cyrillic "субота", X-160).
  0x0402: "bg-BG", 0x0403: "ca-ES", 0x040f: "is-IS", 0x041c: "sq-AL", 0x0420: "ur-PK",
  0x0423: "be-BY", 0x0424: "sl-SI", 0x0425: "et-EE", 0x0426: "lv-LV", 0x0427: "lt-LT",
  0x0429: "fa-IR", 0x042b: "hy-AM", 0x042c: "az-Latn-AZ", 0x042d: "eu-ES", 0x042f: "mk-MK",
  0x0436: "af-ZA", 0x0437: "ka-GE", 0x0439: "hi-IN", 0x043e: "ms-MY", 0x043f: "kk-KZ",
  0x0441: "sw-KE", 0x0443: "uz-Latn-UZ", 0x0445: "bn-IN", 0x0449: "ta-IN", 0x0456: "gl-ES",
  0x0810: "it-CH", 0x0813: "nl-BE", 0x0814: "nn-NO", 0x081a: "sr-Latn-CS", 0x0c1a: "sr-Cyrl-CS",
  0x101a: "hr-BA", 0x141a: "bs-Latn-BA", 0x181a: "sr-Latn-BA", 0x1c1a: "sr-Cyrl-BA",
  0x201a: "bs-Cyrl-BA", 0x241a: "sr-Latn-RS", 0x281a: "sr-Cyrl-RS", 0x2c1a: "sr-Latn-ME",
  0x301a: "sr-Cyrl-ME", 0x1004: "zh-SG", 0x140c: "fr-LU", 0x1407: "de-LI", 0x1007: "de-LU",
  0x2c0a: "es-AR", 0x240a: "es-CO", 0x340a: "es-CL", 0x280a: "es-PE", 0x200a: "es-VE",
};

function dateNamesLocale(section: string): string | undefined {
  const tag = /\[\$[^\]]*?-([^\]]+)\]/.exec(section)?.[1]?.trim();
  if (!tag) return undefined;
  if (/^[0-9a-f]{1,8}$/i.test(tag)) {
    const lid = Number.parseInt(tag, 16) & 0xffff;
    const locale = LID_LOCALES[lid];
    return locale && !locale.startsWith("en") ? locale : undefined;
  }
  // A spelled tag ([$-de-DE]); English keeps the tables.
  return /^[a-z]{2,3}(?:-[a-z0-9]{2,8})*$/i.test(tag) && !/^en\b/i.test(tag) ? tag : undefined;
}

/**
 * The calendar a `[$-NCCLLLL]` tag asks for (XLS-A055, Format Cells > Date >
 * Calendar type): CC is Excel's calendar id — 3 Japanese emperor era,
 * 4 Taiwan (Minguo), 5 Korean Tangun era, 6 Hijri, 7 Thai Buddhist,
 * 8 Hebrew lunar, 0x17 Umm al-Qura; 0-2 and 9-12 are Gregorian variants.
 * Null for the Gregorian calendar.
 */
export type SpreadsheetDateCalendar =
  | "japanese"
  | "roc"
  | "tangun"
  | "islamic"
  | "buddhist"
  | "hebrew"
  | "islamic-umalqura";
const CALENDAR_IDS: Record<number, SpreadsheetDateCalendar> = {
  3: "japanese",
  4: "roc",
  5: "tangun",
  6: "islamic",
  7: "buddhist",
  8: "hebrew",
  0x17: "islamic-umalqura",
};

export function dateCalendarOfSection(section: string): SpreadsheetDateCalendar | null {
  const tag = /\[\$[^\]]*?-([0-9a-f]{5,8})\]/i.exec(section)?.[1];
  if (!tag) return null;
  return CALENDAR_IDS[(Number.parseInt(tag, 16) >> 16) & 0xff] ?? null;
}

/** A Gregorian date read in another calendar: its year, month, day and era. */
function calendarParts(
  calendar: SpreadsheetDateCalendar,
  year: number,
  month: number,
  day: number,
  namesLocale: string | undefined,
): { year: number; month: number; day: number; era: string; eraShort: string; monthLong: string; monthShort: string } | null {
  if (calendar === "tangun") {
    return { year: year + 2333, month, day, era: "", eraShort: "", monthLong: "", monthShort: "" };
  }
  try {
    const at = new Date(Date.UTC(year, month - 1, day));
    const locale = `${namesLocale ?? "en-US"}-u-ca-${calendar}`;
    const numeric = new Intl.DateTimeFormat(`en-US-u-ca-${calendar}`, {
      year: "numeric", month: "numeric", day: "numeric", timeZone: "UTC",
    }).formatToParts(at);
    const named = new Intl.DateTimeFormat(locale, { era: "long", month: "long", timeZone: "UTC" }).formatToParts(at);
    const short = new Intl.DateTimeFormat(locale, { era: "short", month: "short", timeZone: "UTC" }).formatToParts(at);
    const part = (parts: Intl.DateTimeFormatPart[], type: string) => parts.find((p) => p.type === type)?.value ?? "";
    const toNumber = (value: string) => Number(value.replace(/[^0-9-]/g, ""));
    return {
      year: toNumber(part(numeric, "year")),
      month: toNumber(part(numeric, "month")),
      day: toNumber(part(numeric, "day")),
      era: part(named, "era"),
      eraShort: part(short, "era"),
      monthLong: part(named, "month"),
      monthShort: part(short, "month"),
    };
  } catch {
    return null;
  }
}

const dateNameFormatters = new Map<string, Intl.DateTimeFormat | null>();
function dateNameFormatter(locale: string, part: "month" | "weekday", long: boolean) {
  const key = `${locale}|${part}|${long ? "long" : "short"}`;
  let formatter = dateNameFormatters.get(key);
  if (formatter === undefined) {
    try {
      formatter = new Intl.DateTimeFormat(locale, {
        [part]: long ? "long" : "short",
        timeZone: "UTC",
      });
    } catch {
      formatter = null;
    }
    dateNameFormatters.set(key, formatter);
  }
  return formatter;
}

/** The month name of `month` (1-12) in the picture's language, full or
 *  abbreviated like the English mmmm / mmm. */
function monthName(locale: string | undefined, month: number, long: boolean): string {
  const english = long ? MONTHS[month - 1] : MONTHS[month - 1].slice(0, 3);
  const formatter = locale ? dateNameFormatter(locale, "month", long) : null;
  if (!formatter) return english;
  return formatter.format(new Date(Date.UTC(2000, month - 1, 1))) || english;
}

function weekdayName(locale: string | undefined, weekday: number, long: boolean): string {
  const english = long ? WEEKDAYS[weekday] : WEEKDAYS[weekday].slice(0, 3);
  const formatter = locale ? dateNameFormatter(locale, "weekday", long) : null;
  if (!formatter) return english;
  
  return formatter.format(new Date(Date.UTC(2000, 0, 2 + weekday))) || english;
}

// A number-spec placeholder character (digits, grouping, decimal point).
const NUM_SPEC_RE = /[0#?,.]/;

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

/**
 * Split a format string into its `;`-separated sections, treating separators
 * inside quoted strings, `\` escapes and `[...]` brackets as literal.
 */
function splitSections(format: string): string[] {
  const sections: string[] = [];
  let cur = "";
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') {
      cur += ch;
      i++;
      while (i < format.length && format[i] !== '"') cur += format[i++];
      if (i < format.length) cur += format[i]; // closing quote
    } else if (ch === "\\") {
      cur += ch;
      if (i + 1 < format.length) cur += format[++i];
    } else if (ch === "[") {
      cur += ch;
      i++;
      while (i < format.length && format[i] !== "]") cur += format[i++];
      if (i < format.length) cur += format[i];
    } else if (ch === ";") {
      sections.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  sections.push(cur);
  return sections;
}

/** Strip literals and layout operands, leaving only the value-format tokens. */
/**
 * Whether a section carries the Text placeholder. A quoted "@" or an escaped
 * \@ is a literal at-sign the number renderer prints after the digits
 * (0"@" shows 42 as 42@, exactly like 0"#"), not the placeholder (B289).
 */
function isTextPlaceholderSection(section: string): boolean {
  return stripLiterals(section).includes("@");
}

/** Characters with no meaning in a number picture that Excel prints where
 *  they stand, unquoted (`|` in `|#.#|E-|#|`, POI's approximation tests). */
const BARE_PICTURE_LITERALS = new Set(["|", " ", "!", "^", "&", "'", "~", "{", "}", "<", ">", "=", "(", ")", "$", ":"]);

/** `section` with each bare literal escaped (`\|`), so the scientific reader
 *  keeps it in place between the placeholders rather than in its core. */
function escapeBarePictureLiterals(section: string): string {
  let out = "";
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end < 0 ? section.length : end + 1;
      out += section.slice(i, stop);
      i = stop - 1;
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      out += section.slice(i, i + 2);
      i++;
    } else if (ch === "[") {
      const end = section.indexOf("]", i + 1);
      const stop = end < 0 ? section.length : end + 1;
      out += section.slice(i, stop);
      i = stop - 1;
    } else {
      out += BARE_PICTURE_LITERALS.has(ch) ? `\\${ch}` : ch;
    }
  }
  return out;
}











export function textFunctionRejectsFormat(format: string): boolean {
  for (const section of splitSections(format)) {
    if (/@|general|通用格式/i.test(section)) continue;
    let placeholdersBefore = 0;
    let placeholdersAfter = 0;
    let exponent = false;
    let dateCode = false;
    let lastLetter = "";
    let secondsFraction = false;
    let percents = 0;
    for (let i = 0; i < section.length; i++) {
      const ch = section[i];
      // More than one % in a section is no format to TEXT: =TEXT(1,"0%%"),
      // "0.0%%" and "0 %%" are #VALUE! while "0%" is 100% (Excel for Mac
      
      if (ch === "%" && ++percents > 1) return true;
      if (ch === '"') {
        const end = section.indexOf('"', i + 1);
        i = end < 0 ? section.length : end;
        secondsFraction = false;
        continue;
      }
      if (ch === "\\" || ch === "_" || ch === "*") {
        i++;
        secondsFraction = false;
        continue;
      }
      if (ch === "[") {
        const end = section.indexOf("]", i + 1);
        i = end < 0 ? section.length : end;
        continue;
      }
      if (ch === "0" && secondsFraction) continue;
      secondsFraction = ch === "." && lastLetter === "s";
      if (ch === "0" || ch === "#" || ch === "?") {
        if (exponent) placeholdersAfter += 1;
        else placeholdersBefore += 1;
        continue;
      }
      if (ch === "E" && (section[i + 1] === "+" || section[i + 1] === "-")) {
        if (exponent) return true;
        exponent = true;
        i++;
        continue;
      }
      if (/[a-z]/i.test(ch)) {
        lastLetter = ch.toLowerCase();
        if ("dmyhse".includes(lastLetter)) dateCode = true;
      }
    }
    if (dateCode && placeholdersBefore + placeholdersAfter > 0) return true;
    if (exponent && (placeholdersBefore === 0 || placeholdersAfter === 0)) return true;
  }
  return false;
}

function stripLiterals(section: string): string {
  let out = "";
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      while (i < section.length && section[i] !== '"') i++;
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      i++; // An escaped, reserved or repeated character is not a date/digit token.
    } else if (ch === "[") {
      while (i < section.length && section[i] !== "]") i++;
    } else {
      out += ch;
    }
  }
  return out;
}

/**
 * Decode a literal region (a slice of a section OUTSIDE the number/spec core)
 * into the visible text Excel would print: quoted "…" runs and `\x` escapes
 * unwrap to their contents, `[…]` colour/condition brackets vanish, `_x` emits
 * a single space (reserved width), `*x` emits nothing (column fill, no pixel
 * metrics), and any other char (currency $, parens, punctuation) is literal.
 * Mirrors the prefix/suffix handling in `renderNumberSection`.
 */
function decodeLiterals(region: string): string {
  let out = "";
  for (let i = 0; i < region.length; i++) {
    const ch = region[i];
    if (ch === '"') {
      i++;
      while (i < region.length && region[i] !== '"') out += region[i++];
    } else if (ch === "\\") {
      if (i + 1 < region.length) out += region[++i];
    } else if (ch === "[") {
      while (i < region.length && region[i] !== "]") i++; // drop colour/condition
    } else if (ch === "_") {
      out += " ";
      if (i + 1 < region.length) i++; // swallow the spacing char
    } else if (ch === "*") {
      if (i + 1 < region.length) i++; // swallow the fill char, emit nothing
    } else {
      out += ch;
    }
  }
  return out;
}

/**
 * Split a section into a leading literal region, the placeholder/spec CORE, and
 * a trailing literal region, so a renderer can rebuild `${prefix}${inner}${suffix}`
 * instead of discarding the surrounding text. `isCore(ch, prevCore)` classifies
 * a TOP-LEVEL char (outside quotes/escapes/brackets) as part of the core span;
 * `prevCore` is the previous core char (for context like "+/- only after E").
 * The core span runs from the first to the last core char; everything before is
 * the prefix region, everything after the suffix region (both returned RAW, to
 * be passed through `decodeLiterals`). The core slice is returned as its
 * placeholder tokens only, and the literals quoted or escaped INSIDE it come
 * back separately in `literals`, positioned against those tokens.
 */
function splitCore(
  section: string,
  isCore: (ch: string, prevCore: string | undefined) => boolean,
): { prefix: string; core: string; suffix: string; literals: EmbeddedLiteral[] } {
  let firstCore = -1;
  let lastCore = -1;
  let prevCore: string | undefined;
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      while (i < section.length && section[i] !== '"') i++;
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      i++; // Literal/layout operands never belong to a numeric core.
    } else if (ch === "[") {
      while (i < section.length && section[i] !== "]") i++;
    } else if (isCore(ch, prevCore)) {
      if (firstCore === -1) firstCore = i;
      lastCore = i;
      prevCore = ch;
    }
  }
  if (firstCore === -1) {
    return { prefix: "", core: "", suffix: decodeLiterals(section), literals: [] };
  }
  const { core, literals } = readCoreLiterals(section.slice(firstCore, lastCore + 1));
  return {
    prefix: section.slice(0, firstCore),
    core,
    suffix: section.slice(lastCore + 1),
    literals,
  };
}

/**
 * Read a numeric core span into its placeholder tokens (exactly what
 * `stripLiterals` keeps) and the literal text that sits BETWEEN those tokens,
 * each recorded at its offset into the token string. Microsoft Learn,
 * NumberingFormats: a character inside quotation marks or after a backslash
 * is always a literal and is displayed where the code places it — so
 * `0"X".00E+00` shows 123 as 1X.23E+02 and `# "X"?/?` shows 1.25 as 1 X1/4
 * (live audit B472). The literal decoding matches `renderNumberSection`:
 * quoted runs and `\x` print their text, `_x` a space, `*x` nothing, and a
 * `[$sym-locale]` bracket its currency symbol.
 */
function readCoreLiterals(span: string): { core: string; literals: EmbeddedLiteral[] } {
  let core = "";
  const literals: EmbeddedLiteral[] = [];
  const emit = (text: string) => {
    if (text) literals.push({ at: core.length, text });
  };
  for (let i = 0; i < span.length; i++) {
    const ch = span[i];
    if (ch === '"') {
      i++;
      let inner = "";
      while (i < span.length && span[i] !== '"') inner += span[i++];
      emit(inner);
    } else if (ch === "\\") {
      if (i + 1 < span.length) emit(span[++i]);
    } else if (ch === "_") {
      emit(" ");
      i++;
    } else if (ch === "*") {
      i++;
    } else if (ch === "[") {
      let inner = "";
      i++;
      while (i < span.length && span[i] !== "]") inner += span[i++];
      if (inner.startsWith("$")) emit(inner.slice(1).split("-")[0]);
    } else {
      core += ch;
    }
  }
  return { core, literals };
}

/** A section is date/time when it carries an unquoted d/m/y/h/s token, or an
 *  elapsed-time bracket like [h]/[mm]/[ss] (which stripLiterals would remove). */
function isDateTimeSection(section: string): boolean {
  // `a` is the Chinese weekday token (aaa 三 / aaaa 星期三) and `e` the
  // four-digit year — MEASURED in WPS. Without them a section that is nothing
  // but weekday, such as "aaa", never reaches the date renderer and echoes
  // itself as a literal.
  // "General" and its Chinese spelling are whole-format words, not pictures —
  // and now that `a` and `e` are date letters, "General" would otherwise read
  // as one (G-e-n-e-r-a-l). Remove them before looking for date letters.
  const body = stripLiterals(escapeBarePictureLiterals(section))
    .replace(/general/gi, "")
    .replace(/G\/通用格式/g, "")
    .replace(/通用格式/g, "")
    // A scientific exponent is an E followed by a sign or a placeholder
    // (0.00E+00), not the year token — remove it before looking for `e`.
    .replace(/[eE][+-]?[0#?]+/g, "");
  return /[dmyhsae]/i.test(body) || /\[[hms]+\]/i.test(section);
}

interface SectionPick {
  section: string;
  /** Render the absolute value and prefix a "-" (single-section negatives). */
  autoMinus: boolean;
  /** The (signed) value, used by a "General" section. */
  signed: number;
  /** Every section carried a condition and none of them matched the value. */
  unmatched?: boolean;
}

type Comparator = ">=" | "<=" | "<>" | ">" | "<" | "=";
interface Condition {
  op: Comparator;
  rhs: number;
}

/**
 * Extract a leading bracketed comparison condition (e.g. "[>=100]…", "[<0]…")
 * from a section. Returns `null` for a colour-only bracket ([Red]/[Blue]) or no
 * bracket — only a `<op><number>` payload counts as a condition.
 */
function extractCondition(section: string): Condition | null {
  // The condition may sit behind a colour or locale bracket — Excel writes the
  // colour first ("[Red][>=100]0"), and "[>=100][Red]0" is the same format —
  // so every leading bracket is read until a comparison is found.
  const bracket = /^\s*\[([^\]]*)\]/;
  let rest = section;
  for (;;) {
    const m = bracket.exec(rest);
    if (!m) return null;
    const cond = /^\s*(>=|<=|<>|>|<|=)\s*(-?\d+(?:\.\d+)?)\s*$/.exec(m[1]);
    if (cond) return { op: cond[1] as Comparator, rhs: Number(cond[2]) };
    rest = rest.slice(m[0].length);
  }
}

function conditionMatches(cond: Condition, value: number): boolean {
  switch (cond.op) {
    case ">=":
      return value >= cond.rhs;
    case "<=":
      return value <= cond.rhs;
    case "<>":
      return value !== cond.rhs;
    case ">":
      return value > cond.rhs;
    case "<":
      return value < cond.rhs;
    case "=":
      return value === cond.rhs;
  }
}

/**
 * Choose the section for `value`'s sign. With one section, negatives reuse it
 * with an automatic minus. With ≥2 sections the negative/zero sections carry
 * their own sign glyphs (parens, literal "-") so no minus is auto-added.
 *
 * When ANY section carries a bracketed comparison condition ([>=100], [<0]),
 * selection switches to first-match-wins against the SIGNED value (Excel: a
 * condition overrides the default sign-based routing). The first section whose
 * condition matches — or which has no condition (an unconditional default) —
 * wins; if none match, the last section is the fallback. Colour-only brackets
 * are NOT conditions, so they leave the sign-based path intact.
 */
function pickSection(value: number, sections: string[]): SectionPick {
  const pos = sections[0] ?? "";
  const conditions = sections.map(extractCondition);
  if (conditions.some((c) => c !== null)) {
    for (let i = 0; i < sections.length; i++) {
      const cond = conditions[i];
      if (cond === null || conditionMatches(cond, value)) {
        // A format of one conditioned section shows a negative value with
        // its minus, as Excel does (POI FormatConditionTests in
        // P09_FormatConditionTests: -11 in [<-10]#" Wow" is "-11 Wow", X-163).
        return { section: sections[i] ?? "", autoMinus: sections.length === 1 && value < 0, signed: value };
      }
    }
    if (sections.length === 1) {
      // One conditioned section the value fails: Excel shows the value in
      // General (10 in [<10]#" Wow" is "10", X-163), not a cell of #.
      return { section: "General", autoMinus: value < 0, signed: value };
    }
    // No condition matched and no unconditional default: the value has no
    // section to show it. Excel fills the cell with # rather than borrowing
    // the last section's sign glyphs (50 under [>=100]0;[<0]-0 is not -50,
    // live audit B300).
    return {
      section: sections[sections.length - 1] ?? "",
      autoMinus: false,
      signed: value,
      unmatched: true,
    };
  }
  if (value >= 0) {
    if (value === 0 && sections.length >= 3) {
      return { section: sections[2] ?? "", autoMinus: false, signed: value };
    }
    return { section: pos, autoMinus: false, signed: value };
  }
  // value < 0
  if (sections.length >= 2) {
    return { section: sections[1] ?? "", autoMinus: false, signed: value };
  }
  return { section: pos, autoMinus: true, signed: value };
}

/**
 * The characters a number's decimal point and thousands separator are DRAWN
 * with. A format code always spells them "." and "," (the file's and the
 * code's own syntax); Excel draws them with the viewer's system separators
 * (Microsoft Support, "Review guidelines for customizing a number format":
 * the regional settings determine the decimal symbol and the thousands
 * separator), so 5000 through `#,##0.00 [$€-407]` is 5,000.00 € for an
 * English viewer and 5.000,00 € for a German one. The `[$€-407]` tag only
 * names the symbol. Literal text in a code (`"1.5 kg"`, `\.`, a date's
 * `dd.mm.yyyy`) is never a separator and is drawn as written.
 */
export interface NumberFormatSeparators {
  readonly decimal: string;
  readonly group: string;
}

/** The separators a format code is written with, and every stored, cached,
 * exported or AI-read display text is drawn with. */
export const CODE_NUMBER_SEPARATORS: NumberFormatSeparators = Object.freeze({ decimal: ".", group: "," });

/** applyNumberFormat's options: the workbook's date system and, for a
 * viewer's own drawing only, the separators to draw with. */
export interface NumberFormatRenderOptions extends SpreadsheetDateSystemOptions {
  separators?: NumberFormatSeparators;
}

// The separators of the applyNumberFormat call being rendered. Rendering is
// synchronous and pure, so the call sets them for its own extent and every
// section renderer reads them here instead of each taking another argument.
let drawnSeparators: NumberFormatSeparators = CODE_NUMBER_SEPARATORS;

function sameSeparators(a: NumberFormatSeparators, b: NumberFormatSeparators): boolean {
  return a.decimal === b.decimal && a.group === b.group;
}

/** A plain numeral ("-1234.5", "1.2E+05", General's spelling, which never
 * groups) drawn with the active decimal separator. */
function drawnNumeral(text: string): string {
  return drawnSeparators.decimal === "." ? text : text.replace(".", drawnSeparators.decimal);
}

/** A General number drawn with the active separators. */
function drawnGeneralNumber(value: number): string {
  return drawnNumeral(formatGeneralNumber(value));
}

/**
 * General's text for `value` drawn with `separators` (the viewer's), for a
 * number shown without any format code.
 */
export function formatGeneralNumberWith(
  value: number,
  separators: NumberFormatSeparators | undefined,
): string {
  const text = formatGeneralNumber(value);
  return separators && separators.decimal !== "." ? text.replace(".", separators.decimal) : text;
}

function groupThousands(intPart: string): string {
  return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, drawnSeparators.group);
}

/** Excel's 8 named number-format colours -> a CSS colour. (The numbered
 *  [Color1]..[Color56] palette indices are not mapped — they return null.) */
const NAMED_FORMAT_COLORS: Record<string, string> = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  green: "#008000",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00",
};

/** The CSS colour a format section names via a `[Red]`-style bracket, or null.
 *  Scans only ALPHA-only brackets so a condition ([>=100]) or an elapsed date
 *  token ([h]/[mm]) is never mistaken for a colour. */
function sectionColor(section: string): string | null {
  const re = /\[([A-Za-z]+)\]/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(section)) !== null) {
    const name = m[1].toLowerCase();
    if (name in NAMED_FORMAT_COLORS) return NAMED_FORMAT_COLORS[name];
  }
  return null;
}

/**
 * The text colour a number format applies to `value` — e.g. the "[Red]" in the
 * negative section of `0.00;[Red]-0.00` makes a negative render red. Returns a
 * CSS colour string or null. Uses the SAME section-selection as
 * {@link applyNumberFormat} (sign-based routing AND `[>=100]`-style conditions),
 * so the colour always matches the section whose text is shown. PURE.
 */
export function numberFormatColor(value: number, format: string): string | null {
  if (!format || !Number.isFinite(value)) return null;
  const sections = splitSections(format);
  if (sections.length === 0) return null;
  return sectionColor(pickSection(value, sections).section);
}

/** Render `abs` (already non-negative) through a numeric section's tokens. */
/**
 * Round to `decimals` places the way a spreadsheet does — on the number's own
 * DECIMAL spelling, half away from zero.
 *
 * `toFixed` rounds the binary double, and 1.005 is really 1.00499999999999989,
 * so it produces "1.00". MEASURED in WPS: TEXT(1.005,"0.00") is 1.01. The
 * shortest round-trip decimal (what `String` gives) is the spelling the user
 * typed, so rounding that reproduces it — and 0.125, which IS exact in binary,
 * still goes to 0.13 because the tie rounds away from zero either way.
 */
function toFixedDecimal(value: number, decimals: number): string {
  if (!Number.isFinite(value)) return value.toFixed(decimals);
  const shortest = String(Math.abs(value));
  // An exponent form (1e-7, 1.5e+21) has no plain decimal spelling to walk.
  if (shortest.includes("e") || shortest.includes("E")) {
    return value.toFixed(decimals);
  }
  const sign = value < 0 || Object.is(value, -0) ? "-" : "";
  // Excel rounds the value's 15-significant-digit decimal, so a product the
  // binary arithmetic left a hair under a half (0.1785 x 100 for 0.0% =
  // 17.849999999999998) still rounds as 17.85 does (17.9%; P20_AverageTaxRates,
  // X-205).
  const fifteen = String(Math.abs(Number(value.toPrecision(15))));
  const [intPart, fracPart = ""] = (fifteen.includes("e") || fifteen.includes("E") ? shortest : fifteen).split(".");
  if (fracPart.length <= decimals) {
    return sign + intPart + (decimals > 0 ? "." + fracPart.padEnd(decimals, "0") : "");
  }
  const keep = (intPart + fracPart.slice(0, decimals)).split("");
  const roundUp = Number(fracPart[decimals]) >= 5;
  if (roundUp) {
    let i = keep.length - 1;
    for (; i >= 0; i--) {
      if (keep[i] === "9") {
        keep[i] = "0";
        continue;
      }
      keep[i] = String(Number(keep[i]) + 1);
      break;
    }
    if (i < 0) keep.unshift("1");
  }
  const digits = keep.join("");
  const cut = digits.length - decimals;
  const head = digits.slice(0, cut) || "0";
  return sign + head + (decimals > 0 ? "." + digits.slice(cut) : "");
}

/** A literal run recorded at its offset into a section's placeholder spec. */
interface EmbeddedLiteral {
  at: number;
  text: string;
}

/**
 * Place the literals that sit BETWEEN the integer digit placeholders.
 *
 * Microsoft Support, "Number format codes" / "Create or delete a custom number
 * format": a code may mix literal characters (`-`, space, `(`, `)`, `/`, quoted
 * text) in among the `0`/`#`/`?` placeholders, and Excel fills the digits into
 * the placeholders RIGHT-TO-LEFT while every literal stays where the code wrote
 * it. That is what makes the 特殊 category work — MEASURED in Excel:
 *   7            "000-0000"        000-0007
 *   13912345678  "000-0000-0000"   139-1234-5678
 * A value wider than the picture puts its extra leading digits before the FIRST
 * placeholder, and a segment whose spec carries a thousands comma is grouped on
 * its own.
 */
function spliceIntegerLiterals(
  digits: string,
  intSpec: string,
  literals: readonly EmbeddedLiteral[],
): string {
  // Cut the integer spec at each literal offset; several literals may share one.
  const byOffset = new Map<number, string>();
  for (const literal of literals) {
    const at = Math.max(0, Math.min(literal.at, intSpec.length));
    byOffset.set(at, (byOffset.get(at) ?? "") + literal.text);
  }
  const offsets = [...byOffset.keys()].sort((a, b) => a - b);
  const segments: string[] = [];
  const separators: string[] = [];
  let start = 0;
  for (const at of offsets) {
    segments.push(intSpec.slice(start, at));
    separators.push(byOffset.get(at)!);
    start = at;
  }
  segments.push(intSpec.slice(start));

  // Fill right-to-left: each segment takes as many digits as it has
  // placeholders, and the leftmost one keeps whatever is left over.
  const filled: string[] = new Array(segments.length).fill("");
  let rest = digits;
  for (let i = segments.length - 1; i >= 1; i--) {
    const width = (segments[i].match(/[0#?]/g) || []).length;
    const take = Math.min(width, rest.length);
    filled[i] = rest.slice(rest.length - take);
    rest = rest.slice(0, rest.length - take);
  }
  filled[0] = rest;

  let out = "";
  for (let i = 0; i < segments.length; i++) {
    out += segments[i].includes(",") && filled[i].trim()
      ? groupThousands(filled[i])
      : filled[i];
    if (i < separators.length) out += separators[i];
  }
  return out;
}

/**
 * Place the literals that sit between the FRACTION placeholders. A decimal
 * position is fixed, so the digits fill left-to-right and a literal is printed
 * after the placeholders that precede it — MEASURED: 1.23 through `0.0"x"0`
 * is 1.2x3.
 */
function spliceFractionLiterals(
  fracChars: readonly string[],
  fracSpec: string,
  literals: readonly EmbeddedLiteral[],
): string {
  const placed = literals
    .map((literal) => ({
      slot: (fracSpec.slice(0, literal.at).match(/[0#?]/g) || []).length,
      text: literal.text,
    }))
    .sort((a, b) => a.slot - b.slot);
  let out = "";
  let next = 0;
  for (let i = 0; i <= fracChars.length; i++) {
    while (next < placed.length && placed[next].slot === i) out += placed[next++].text;
    if (i < fracChars.length) out += fracChars[i];
  }
  while (next < placed.length) out += placed[next++].text;
  return out;
}

function renderNumberSection(abs: number, section: string, autoMinus: boolean): string {
  // Walk the section, separating literal runs (prefix/suffix) from the
  // contiguous number-placeholder run in the middle.
  let prefix = "";
  let suffix = "";
  let numSpec = "";
  let seenNum = false;
  let isPercent = false;
  // A literal seen after the first placeholder is recorded at its offset into
  // `numSpec`: one still inside the placeholder run keeps its position (see
  // spliceIntegerLiterals); one past the last placeholder is a plain suffix.
  const embedded: EmbeddedLiteral[] = [];

  const emitLiteral = (text: string) => {
    if (seenNum) embedded.push({ at: numSpec.length, text });
    else prefix += text;
  };

  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      let inner = "";
      while (i < section.length && section[i] !== '"') inner += section[i++];
      emitLiteral(inner);
    } else if (ch === "\\") {
      if (i + 1 < section.length) emitLiteral(section[++i]);
    } else if (ch === "[") {
      // Read the bracket body. A `[$<sym>]` / `[$<sym>-<locale>]` token is a
      // CURRENCY symbol and must render (Excel parity, and the form the Currency
      // dialog's escapeCurrencySymbol emits for ₩/₽/₪/₺ and that imported .xlsx
      // carries for locale-tagged currencies like [$£-809]). Everything else —
      // colour ([Red]), condition ([>=100]), elapsed-time ([h]/[mm]/[ss]) and a
      // locale-only prefix ([$-409], no symbol before the dash) — renders nothing.
      // Mirrors accountingParams' decode so the plain and accounting paths agree.
      let inner = "";
      i++; // move past "["
      while (i < section.length && section[i] !== "]") inner += section[i++];
      // i now sits on "]" (or end); the loop's i++ skips it, as before.
      if (inner.startsWith("$")) {
        const sym = inner.slice(1).split("-")[0];
        if (sym) emitLiteral(sym);
      }
    } else if (ch === "%") {
      isPercent = true;
      emitLiteral("%");
    } else if (ch === "_") {
      // Width placeholder "_x": Excel reserves the width of the following char
      // and renders nothing visible — approximate as a single space and swallow
      // the spacing char (else "_)" leaks a literal "_)" into the value).
      emitLiteral(" ");
      if (i + 1 < section.length) i++;
    } else if (ch === "*") {
      // Fill char "*x": repeats the next char to fill the column. With no pixel
      // metrics here, render nothing and swallow the fill char (else "* " leaks).
      if (i + 1 < section.length) i++;
    } else if (NUM_SPEC_RE.test(ch)) {
      seenNum = true;
      numSpec += ch;
    } else {
      emitLiteral(ch);
    }
  }

  // A comma AFTER the last digit placeholder scales the value down by 1000 each
  // (Excel: "#,##0," shows thousands, "0.0,," shows millions). Strip those
  // trailing commas before reading grouping / decimals.
  const trailingCommas = (numSpec.match(/,+$/) || [""])[0].length;
  const numSpecTrimmed = trailingCommas
    ? numSpec.slice(0, numSpec.length - trailingCommas)
    : numSpec;
  // A comma BEFORE any digit placeholder groups nothing: it is a literal
  // comma (Excel: ",#" of 1234567 -> ",1234567", X-136).
  const leadingCommas = (numSpecTrimmed.match(/^,+/) || [""])[0].length;
  const numSpecCore = numSpecTrimmed.slice(leadingCommas);
  if (leadingCommas) prefix += ",".repeat(leadingCommas);
  // Split the number spec into integer / fraction placeholder runs.
  const dotIdx = numSpecCore.indexOf(".");
  const intSpec = dotIdx >= 0 ? numSpecCore.slice(0, dotIdx) : numSpecCore;
  const fracSpec = dotIdx >= 0 ? numSpecCore.slice(dotIdx + 1) : "";
  // The fraction placeholder run, in order, dropping any non-placeholder chars
  // so each spec character lines up 1:1 with a `toFixed` fraction digit below.
  const fracPlaceholders = (fracSpec.match(/[0#?]/g) || []);
  const decimals = fracPlaceholders.length;
  // REQUIRED fraction positions: '0' forces a digit; '#'/'?' are optional and
  // their trailing zeros are trimmed ('#' dropped, '?' -> alignment space).
  const minFrac = (fracSpec.match(/0/g) || []).length;
  const minInt = (intSpec.match(/0/g) || []).length;
  // Integer-side '?' reserve width with a leading space for an absent digit.
  const intPlaceholderCount = (intSpec.match(/[0?]/g) || []).length;
  const useThousands = intSpec.includes(",");
  const hasDigits = /[0#?]/.test(numSpecCore);

  // Route each recorded literal: inside the integer run, inside the fraction
  // run, or past the last placeholder — the last case is the plain suffix this
  // renderer has always produced (`0 "kg"`, `#,##0.00_)`, `0%`).
  const intLiterals: EmbeddedLiteral[] = [];
  const fracLiterals: EmbeddedLiteral[] = [];
  for (const literal of embedded) {
    if (!hasDigits || literal.at >= numSpecCore.length) {
      suffix += literal.text;
    } else if (dotIdx >= 0 && literal.at > dotIdx) {
      fracLiterals.push({ at: literal.at - dotIdx - 1, text: literal.text });
    } else {
      intLiterals.push(literal);
    }
  }

  let body = "";
  let roundsToZero = true;
  if (hasDigits) {
    let scaled = isPercent ? abs * 100 : abs;
    if (trailingCommas) scaled /= 1000 ** trailingCommas;
    const fixed = toFixedDecimal(scaled, decimals);
    roundsToZero = Number(fixed) === 0;
    const [intRaw, fracRaw] = fixed.split(".");
    let intPart = intRaw;
    if (intPart.length < minInt) intPart = intPart.padStart(minInt, "0");
    // A '#'-only integer run (no '0' placeholder) does not force a digit, so a
    // bare leading zero for |value| < 1 is suppressed (Excel: "#.##" of 0.456 ->
    // ".46", "#" of 0 -> ""). Strip before grouping so it never sees the digit.
    if (minInt === 0 && intPart === "0") intPart = "";
    // Integer side '?': pad the shortfall below the reserved placeholder width
    // with leading spaces (one per unfilled '?'), before grouping. Only the
    // '?' slots reserve width — '0' slots are already filled by minInt padding,
    // and '#' slots reserve nothing.
    if (intPart.length < intPlaceholderCount) {
      intPart = intPart.padStart(intPlaceholderCount, " ");
    }
    if (intLiterals.length) intPart = spliceIntegerLiterals(intPart, intSpec, intLiterals);
    else if (useThousands) intPart = groupThousands(intPart);

    // Per-placeholder trim of OPTIONAL trailing fraction positions. Walk from
    // the right; a position i (0-based from left) past the required width
    // (i >= minFrac) is trimmable: '#' drops a trailing '0', '?' replaces it
    // with an alignment space. Stop at the first non-trimmable position — a '0'
    // placeholder, a significant non-zero digit, or a '?' once a later digit is
    // significant — so interior zeros before a later digit are preserved.
    let fracChars = fracRaw === undefined ? [] : fracRaw.split("");
    if (fracChars.length) {
      for (let i = fracChars.length - 1; i >= 0; i--) {
        if (i < minFrac) break; // required position
        const spec = fracPlaceholders[i];
        if (spec === "#" && fracChars[i] === "0") {
          fracChars = fracChars.slice(0, i);
        } else if (spec === "?" && fracChars[i] === "0") {
          fracChars[i] = " ";
        } else {
          break; // '0' placeholder or a significant digit — stop trimming.
        }
      }
    }
    const fracPart = fracLiterals.length
      ? spliceFractionLiterals(fracChars, fracSpec, fracLiterals)
      : fracChars.join("");
    // The decimal point stays when every optional fraction position was
    // dropped: Excel draws "0.##" of 5 as "5." and "#,##.#" of 1234567 as
    // "1,234,567." (P05_NumberFormatTests, X-136).
    body = fracRaw !== undefined ? `${intPart}${drawnSeparators.decimal}${fracPart}` : intPart;
  }

  // Suppress the auto sign when the value rounds to zero at this precision, so a
  // tiny negative (e.g. -0.001 with "0.00") shows "0.00" — not "-0.00" — matching
  // Excel / Google Sheets. (Two-section formats pick the negative section and
  // never set autoMinus, so they are unaffected.)
  // (Read from the ROUNDED digits, not from `body`: an embedded literal makes
  // the assembled text non-numeric.)
  const showMinus = autoMinus && hasDigits && !roundsToZero;
  return `${showMinus ? "-" : ""}${prefix}${body}${suffix}`;
}

interface DateToken {
  lit: boolean;
  text: string;
  letter?: string; // 'd' | 'm' | 'y' | 'h' | 's' for non-literal tokens
  ampm?: boolean;
  elapsed?: "h" | "m" | "s"; // [h]/[mm]/[ss] total-duration token
  subSec?: boolean; // a seconds token carrying a ".0+" fractional suffix (ss.00)
}

/** Render a date serial through a date/time section like "yyyy-mm-dd hh:mm AM/PM". */
//: Sunday-first, matching Date#getUTCDay. MEASURED in WPS: aaa on a Sunday is
//: 日 and on a Wednesday 三; aaaa prefixes 星期.
const CN_WEEKDAYS = ["日", "一", "二", "三", "四", "五", "六"];

function renderDateTimeSection(
  serial: number,
  section: string,
  date1904: ExcelDateSystemArg = false,
): string {
  // A sub-second token ("ss.0" / "ss.00" / "[ss].0") needs fractional-second
  // precision; otherwise round to WHOLE seconds (the historical default) to kill
  // IEEE-754 noise. p = requested fractional digits; the round unit shrinks to
  // 10^-p seconds only when a sub-second token is present (so every existing
  // whole-second format is byte-identical).
  const subMatch = section.match(/(?:\[s+\]|s+)\.(0+)/i);
  const subSecP = subMatch ? subMatch[1].length : 0;
  const roundUnitMs = subSecP > 0 ? 1000 / 10 ** subSecP : 1000;
  // A section that shows no time (no hour, second, elapsed or AM/PM token;
  // a bare "m" is a month) reads the calendar day the serial lies in: the
  // last half second of 1 March is still 1 March, as DAY() says (live audit
  // B198). Rounding to whole seconds is for the time fields it would show.
  const showsTime = /h|s|am\/pm|a\/p|\[m+\]/i.test(section.replace(/"[^"]*"/g, ""));
  // The workbook's date system decides which calendar day a serial is (the
  // 1904 system reads serial s as the 1900 system's s + 1462; the
  // non-compatible 1900 base reads every serial as the plain 1899-12-30
  // count, X-236). Only the calendar reading moves: the elapsed [h]/[m]/[s]
  // totals below count from serial 0 in every system.
  const leapWindow = excelSystemHasFictitiousLeapDay(date1904);
  const calendarSerial = isExcelDate1904(date1904) ? excelCalendarSerial(serial, date1904) : serial;
  const ms = showsTime
    ? Math.round((DATE_EPOCH_MS + calendarSerial * MS_PER_DAY) / roundUnitMs) * roundUnitMs
    : DATE_EPOCH_MS + Math.floor(calendarSerial) * MS_PER_DAY;
  const date = new Date(ms);
  let year = date.getUTCFullYear();
  let month = date.getUTCMonth() + 1;
  let day = date.getUTCDate();
  let weekday = date.getUTCDay();
  // Before 1900-03-01 Excel's calendar runs one day behind the epoch count and
  // holds a 1900-02-29; the time of day is unaffected (whole-day shift).
  const wholeDay = Math.floor((ms - DATE_EPOCH_MS) / MS_PER_DAY);
  if (leapWindow && wholeDay >= 0 && wholeDay < EXCEL_1900_MARCH_FIRST_SERIAL) {
    ({ year, month, day, weekday } = excelSerialToParts(wholeDay));
  }
  const hours24 = date.getUTCHours();
  const minutes = date.getUTCMinutes();
  const seconds = date.getUTCSeconds();
  // A [$-407] / [$-40C] tag spells the month and weekday names in that
  // language (Januar / Montag, janvier / lundi), not in English (B299).
  const namesLocale = dateNamesLocale(section);
  // XLS-A055: a calendar type in the tag reads the date in that calendar
  // (Thai Buddhist 2555, Minguo 101, Heisei 24, Hebrew, Hijri…).
  // An era code (g / gg / ggg) with no calendar in the tag reads the
  // Japanese emperor era, as Excel does: S06_tdf161301's
  // [$]ggge"年"m"月"d"日" is 令和6年5月28日 (X-159).
  const statedCalendar = dateCalendarOfSection(section);
  const calendar = statedCalendar ?? (/g/i.test(stripLiterals(section)) ? "japanese" : null);
  const converted = calendar
    ? calendarParts(calendar, year, month, day, namesLocale ?? (statedCalendar ? undefined : "ja-JP"))
    : null;
  const lunarMonthNames = calendar === "hebrew" || calendar === "islamic" || calendar === "islamic-umalqura";
  if (converted) {
    year = converted.year;
    month = converted.month;
    day = converted.day;
  }
  // The fractional-second suffix (".25"), derived from the precision-rounded ms
  // so the integer seconds and the fraction stay consistent through any carry.
  const secFracStr =
    subSecP > 0 ? drawnNumeral((date.getUTCMilliseconds() / 1000).toFixed(subSecP).slice(1)) : "";
  // Elapsed/duration totals for [h]/[m]/[s] bracket tokens (total hours,
  // minutes or seconds since serial 0 — NOT wrapped at 24/60). A sub-second
  // request keeps the fractional remainder on the seconds total.
  const totalSecExact =
    subSecP > 0
      ? Math.round(serial * 86_400 * 10 ** subSecP) / 10 ** subSecP
      : Math.round(serial * 86_400);
  const totalSec = Math.trunc(totalSecExact);
  const elapsedSecFracStr =
    subSecP > 0 ? drawnNumeral((totalSecExact - totalSec).toFixed(subSecP).slice(1)) : "";
  const elapsed: Record<"h" | "m" | "s", number> = {
    h: Math.floor(totalSec / 3600),
    m: Math.floor(totalSec / 60),
    s: totalSec,
  };

  // Tokenize.
  const tokens: DateToken[] = [];
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      let inner = "";
      while (i < section.length && section[i] !== '"') inner += section[i++];
      tokens.push({ lit: true, text: inner });
    } else if (ch === "\\") {
      if (i + 1 < section.length) tokens.push({ lit: true, text: section[++i] });
    } else if (ch === "[") {
      i++;
      let inner = "";
      while (i < section.length && section[i] !== "]") inner += section[i++];
      const e = inner.toLowerCase();
      // [h]/[hh]/[m]/[mm]/[s]/[ss] are elapsed-time tokens (total duration);
      // any other bracket ([Red], [<100]) is a colour/condition code -> dropped.
      if (/^h+$/.test(e)) tokens.push({ lit: false, text: inner, elapsed: "h" });
      else if (/^m+$/.test(e)) tokens.push({ lit: false, text: inner, elapsed: "m" });
      else if (/^s+$/.test(e)) {
        const tok: DateToken = { lit: false, text: inner, elapsed: "s" };
        if (section[i + 1] === "." && section[i + 2] === "0") {
          i += 1; // move onto the "."
          while (i + 1 < section.length && section[i + 1] === "0") i += 1;
          tok.subSec = true;
        }
        tokens.push(tok);
      }
    } else if (ch === "_") {
      // Width placeholder "_x" -> a single space; swallow the spacing char (else
      // "m/d/yyyy_)" leaks a literal "_)" into the rendered date).
      tokens.push({ lit: true, text: " " });
      if (i + 1 < section.length) i++;
    } else if (ch === "*") {
      // Fill char "*x" -> nothing here; swallow the fill char.
      if (i + 1 < section.length) i++;
    } else {
      const ampm = section.slice(i, i + 5);
      if (ampm.toUpperCase() === "AM/PM") {
        tokens.push({ lit: false, text: ampm, ampm: true });
        i += 4;
        continue;
      }
      // "上午/下午" is the Chinese 12-hour marker, a picture token of its own
      // beside AM/PM and A/P (XLS-A035). MEASURED in Excel:
      // TEXT(0.75,"上午/下午h""时""mm""分""") is 下午6时00分 and
      // TEXT(0.5,"上午/下午h:mm") 下午12:00.
      if (ampm === "上午/下午") {
        tokens.push({ lit: false, text: ampm, ampm: true });
        i += 4;
        continue;
      }
      // "A/P" is the short 12-hour marker (Excel: "h:mm A/P" is 3:30 P); it is
      // not the East Asian weekday letter a.
      const ap = section.slice(i, i + 3);
      if (ap.toUpperCase() === "A/P") {
        tokens.push({ lit: false, text: ap, ampm: true });
        i += 2;
        continue;
      }
      const letter = ch.toLowerCase();
      // g/gg/ggg is the era of a Japanese or Taiwan calendar section.
      const eraCalendar = calendar === "japanese" || calendar === "roc";
      if ("dmyhsae".includes(letter) || (letter === "g" && eraCalendar)) {
        let run = ch;
        while (i + 1 < section.length && section[i + 1].toLowerCase() === letter) {
          run += section[++i];
        }
        const tok: DateToken = { lit: false, text: run, letter };
        // A seconds token may carry a fractional sub-second suffix ".0+" (ss.00);
        // consume the "." and the zeros so they are not echoed as literals.
        if (letter === "s" && section[i + 1] === "." && section[i + 2] === "0") {
          i += 1; // move onto the "."
          while (i + 1 < section.length && section[i + 1] === "0") i += 1;
          tok.subSec = true;
        }
        tokens.push(tok);
      } else {
        tokens.push({ lit: true, text: ch });
      }
    }
  }

  // Disambiguate 'm': minutes if adjacent (ignoring literals) to an h or s
  // token, otherwise month.
  const nonLit = tokens.filter((t) => !t.lit && !t.ampm);
  const unitOf = (tok?: DateToken): string | undefined => tok?.letter ?? tok?.elapsed;
  const mIsMinute = new Map<DateToken, boolean>();
  for (let k = 0; k < nonLit.length; k++) {
    if (nonLit[k].letter !== "m") continue;
    // Minute when adjacent to an hour or second token — including an elapsed
    // [h]/[s] bracket (which carries `elapsed`, not `letter`).
    const prev = unitOf(nonLit[k - 1]);
    const next = unitOf(nonLit[k + 1]);
    mIsMinute.set(nonLit[k], prev === "h" || next === "s");
  }
  const has12h = tokens.some((t) => t.ampm);
  const hour12 = hours24 % 12 === 0 ? 12 : hours24 % 12;

  let out = "";
  for (const t of tokens) {
    if (t.lit) {
      out += t.text;
      continue;
    }
    if (t.ampm) {
      if (t.text === "上午/下午") {
        out += hours24 < 12 ? "上午" : "下午";
        continue;
      }
      const full = hours24 < 12 ? "AM" : "PM";
      const label = t.text.length === 3 ? full[0] : full;
      out += t.text === t.text.toLowerCase() ? label.toLowerCase() : label;
      continue;
    }
    if (t.elapsed) {
      out += String(elapsed[t.elapsed]).padStart(t.text.length, "0");
      if (t.elapsed === "s" && t.subSec) out += elapsedSecFracStr;
      continue;
    }
    const len = t.text.length;
    switch (t.letter) {
      case "y":
        out += len <= 2 ? pad2(year % 100) : String(year);
        break;
      case "m":
        if (mIsMinute.get(t)) {
          out += len >= 2 ? pad2(minutes) : String(minutes);
        } else if (len === 1) out += String(month);
        else if (len === 2) out += pad2(month);
        else if (lunarMonthNames && converted) out += len === 3 ? converted.monthShort : converted.monthLong;
        else if (len === 3) out += monthName(namesLocale, month, false);
        else if (len >= 5) out += monthName(namesLocale, month, true).slice(0, 1); // mmmmm -> "J"
        else out += monthName(namesLocale, month, true); // mmmm -> full name
        break;
      case "d":
        if (len === 1) out += String(day);
        else if (len === 2) out += pad2(day);
        else if (len === 3) out += weekdayName(namesLocale, weekday, false);
        else out += weekdayName(namesLocale, weekday, true);
        break;
      case "h": {
        const h = has12h ? hour12 : hours24;
        out += len >= 2 ? pad2(h) : String(h);
        break;
      }
      case "s":
        out += len >= 2 ? pad2(seconds) : String(seconds);
        if (t.subSec) out += secFracStr;
        break;
      case "a":
        // The Chinese weekday tokens. MEASURED in WPS over four dates:
        // aaa is 日/一/二/三 and aaaa is 星期日/星期三 — the English names stay
        // on ddd/dddd, which this switch already renders.
        out += len >= 4 ? `星期${CN_WEEKDAYS[weekday]}` : CN_WEEKDAYS[weekday];
        break;
      case "e":
        // MEASURED: both e and ee render the full four-digit year in WPS's
        // Chinese locale — it is not an era offset there. In a Japanese or
        // Taiwan calendar section it is the year of the era (ee pads to 2).
        out += converted && (calendar === "japanese" || calendar === "roc") && len >= 2
          ? pad2(year)
          : String(year);
        break;
      case "g":
        // The era: g its initial, gg the short name, ggg the full name.
        out += !converted ? "" : len >= 3 ? converted.era : len === 2 ? converted.eraShort : converted.eraShort.slice(0, 1);
        break;
    }
  }
  return out;
}

export function formatGeneralNumber(value: number): string {
  const abs = Math.abs(value);
  // Excel's General never writes a non-zero number as a plain "0": a magnitude
  // the noise trim below would flatten (anything under 1e-10, such as the
  // 7.61985E-24 tail of NORM.S.DIST(-10,TRUE)) is written in scientific
  // notation, as is a magnitude past 1e21, where JavaScript switches to an
  // exponent of its own. Excel spells the exponent with a capital E.
  if (abs !== 0 && (abs < 1e-10 || abs >= 1e21)) {
    return generalExponentSpelling(String(Number(value.toPrecision(15))));
  }
  // Trim floating-point noise (0.1 + 0.2 -> 0.3). JavaScript writes a
  // magnitude under 1e-6 with an exponent of its own (1e-7), which Excel
  // spells 1E-07.
  return generalExponentSpelling(String(Math.round(value * 1e10) / 1e10));
}

/** JavaScript's exponent (3e-104, 1e-7, 1e+21) as Excel's General writes
 * it: a capital E, a sign and at least two digits (3E-104, 1E-07, 1E+21). */
function generalExponentSpelling(text: string): string {
  const exponent = /e([+-]?)(\d+)$/.exec(text);
  if (!exponent) return text;
  return `${text.slice(0, exponent.index)}E${exponent[1] || "+"}${exponent[2].padStart(2, "0")}`;
}

// ---------------------------------------------------------------------------
// Chinese numeral formats ([DBNum1] / [DBNum2] / 人民币大写)
//
// Excel & WPS expose Chinese numerals through bracketed locale directives in a
// number-format code: `[DBNum1]` renders the 小写 (everyday) numerals
// ○一二三四五六七八九 with 十百千万亿 positional units; `[DBNum2]` renders the
// 大写 (financial / anti-tamper) numerals 壹贰叁肆伍陆柒捌玖 with 拾佰仟萬億;
// and the 人民币大写 currency format spells a yuan amount out in 大写 with
// 元/角/分/整 and a 负 prefix for negatives. These are implemented as pure
// functions so the numeral conversion is unit-testable in isolation, and are
// detected ahead of the date/number paths in {@link applyNumberFormat} (the
// `[DBNum*]` directive is otherwise stripped as a colour/condition bracket).
// ---------------------------------------------------------------------------

// 小写 (DBNum1) glyph tables. Index 0..9 maps a decimal digit to its numeral.
// digits[0] is the STANDALONE zero (the whole number is 0); a collapsed/bridging
// internal zero is always 零 (see `BRIDGE_ZERO`), so 100000001 → 一亿零一 but a bare
// 0 → ○ (U+25CB, the everyday zero — not the ideograph U+3007).
const CN_LOWER_DIGITS = ["○", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
// Within-group positional units for the ones/tens/hundreds/thousands place.
const CN_LOWER_UNITS = ["", "十", "百", "千"];
// 大写 (DBNum2) glyph tables — financial numerals and their positional units.
const CN_UPPER_DIGITS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
const CN_UPPER_UNITS = ["", "拾", "佰", "仟"];
// 4-digit group markers, indexed [万 (10^4), 亿 (10^8)] — 简体 vs 繁体 variants.
// Higher tiers compose these (10^12 = 万亿, 10^16 = 亿亿) in `groupMarker`.
const CN_LOWER_GROUP_MARKERS = ["万", "亿"];
const CN_UPPER_GROUP_MARKERS = ["萬", "億"];
// The collapsed/bridging interior zero is 零 in BOTH styles (Excel renders a bare
// DBNum1 zero as ○ but an interior gap as 零, e.g. 101 → 一百零一).
const BRIDGE_ZERO = "零";

interface ChineseNumeralStyle {
  digits: string[];
  units: string[];
  groupMarkers: string[];
}

const CN_LOWER_STYLE: ChineseNumeralStyle = {
  digits: CN_LOWER_DIGITS,
  units: CN_LOWER_UNITS,
  groupMarkers: CN_LOWER_GROUP_MARKERS,
};
const CN_UPPER_STYLE: ChineseNumeralStyle = {
  digits: CN_UPPER_DIGITS,
  units: CN_UPPER_UNITS,
  groupMarkers: CN_UPPER_GROUP_MARKERS,
};

/**
 * Convert a non-negative integer STRING (arbitrary length, digits only) into a
 * grouped Chinese numeral using the supplied glyph style. The number is split
 * into 4-digit groups from the right; each group is rendered with 千/百/十 units
 * and joined with 万/亿 markers. Internal zero runs collapse to a single 零 and no
 * trailing zero unit is emitted, so 1004 → 一千零四, 10000 → 一万, 100000000 → 一亿,
 * 100010001 → 一亿零一万零一. The markers compose (10^12 → 万亿, 10^16 → 亿亿), so
 * values well beyond Number.MAX_SAFE_INTEGER stay structurally correct. PURE —
 * operates on a digit string, immune to float precision.
 */
function intStringToChineseNumeral(intStr: string, style: ChineseNumeralStyle): string {
  // Normalise: drop leading zeros (but keep one digit for "0").
  const digits = intStr.replace(/^0+(?=\d)/, "");
  if (digits === "0") return style.digits[0];

  // Split into 4-digit groups, least-significant first (groups[0] = ones group).
  const groups: string[] = [];
  for (let end = digits.length; end > 0; end -= 4) {
    groups.push(digits.slice(Math.max(0, end - 4), end));
  }

  // Render one 4-digit group's body (no 万/亿 marker). Returns the spelled text and
  // whether the group is entirely zero. A `0` between non-zero digits collapses to
  // a single bridging 零; trailing zeros emit nothing (no dangling unit).
  const renderGroup = (group: string): { text: string; allZero: boolean } => {
    const padded = group.padStart(4, "0");
    let text = "";
    let pendingZero = false; // a zero seen since the last emitted non-zero digit
    let emittedAny = false;
    for (let i = 0; i < 4; i++) {
      const d = padded.charCodeAt(i) - 48;
      const unitIdx = 3 - i; // 3=千 .. 0=ones
      if (d === 0) {
        if (emittedAny) pendingZero = true; // defer; only kept if a digit follows
        continue;
      }
      if (pendingZero) {
        text += BRIDGE_ZERO;
        pendingZero = false;
      }
      text += style.digits[d] + style.units[unitIdx];
      emittedAny = true;
    }
    return { text, allZero: !emittedAny };
  };

  // The 10^4 group marker for group index g (g=0 ones): 亿 (10^8) composed
  // floor(g/2) times, with a leading 万 on odd g — so g=1 万, g=2 亿, g=3 万亿,
  // g=4 亿亿, g=5 万亿亿. Higher groups are more significant, so the 万 precedes the
  // 亿 run (10^12 reads 万亿). Structurally correct for arbitrarily long inputs.
  const groupMarker = (g: number): string => {
    if (g === 0) return "";
    const yi = style.groupMarkers[1].repeat(Math.floor(g / 2));
    return (g % 2 === 1 ? style.groupMarkers[0] : "") + yi;
  };

  // Build most-significant group first. A bridging 零 is inserted only immediately
  // BEFORE a group that actually emits text, when the running output has real
  // content and this group's thousands place is empty (1,0004 → 一万零四) or an
  // entirely-zero group sits between two non-zero ones (1,0000,0001 → 一亿零一). A
  // group that is all-zero never appends its own marker, so no trailing 零/marker
  // can dangle (10000 → 一万, not 一万零).
  let out = "";
  let zeroPending = false; // an all-zero gap awaiting a following non-zero group
  for (let g = groups.length - 1; g >= 0; g--) {
    const { text, allZero } = renderGroup(groups[g]);
    if (allZero) {
      if (out !== "") zeroPending = true; // remember the gap; do not emit yet
      continue;
    }
    if (out !== "") {
      // Bridge a gap: either a fully-zero higher group, or this group's missing
      // thousands place. Collapse to a single 零.
      if (zeroPending || groups[g].padStart(4, "0")[0] === "0") {
        if (!out.endsWith(BRIDGE_ZERO)) out += BRIDGE_ZERO;
      }
    }
    zeroPending = false;
    out += text + groupMarker(g);
  }
  return out;
}

/**
 * Render a number through a [DBNum1]/[DBNum2] Chinese-numeral format. The integer
 * part is spelled out with grouped numerals; any fractional part is rendered as a
 * 点 followed by its digits spelled one-by-one (Excel: 1.5 [DBNum1] → 一点五).
 * Negatives gain a 负 prefix. The value is rounded to the format's decimal
 * placeholders when present, else to a stable 10-digit precision (General-like).
 * PURE.
 */
function renderChineseNumeralSection(
  value: number,
  section: string,
  upper: boolean,
): string {
  const style = upper ? CN_UPPER_STYLE : CN_LOWER_STYLE;
  const sign = value < 0 ? "负" : "";
  const abs = Math.abs(value);

  // Decimal precision: honour explicit fraction placeholders ("0.00"); otherwise
  // spell however many fractional digits the General rendering would show.
  const spec = stripLiterals(section);
  const dot = spec.indexOf(".");
  const fracPlaceholders = dot >= 0 ? (spec.slice(dot + 1).match(/[0#?]/g) || []).length : -1;

  let fixed: string;
  if (fracPlaceholders >= 0) {
    fixed = toFixedDecimal(abs, Math.min(20, fracPlaceholders));
  } else {
    // General-like: trim float noise, keep whatever decimals remain.
    fixed = String(Math.round(abs * 1e10) / 1e10);
  }
  const [intRaw, fracRaw = ""] = fixed.split(".");
  const intText = intStringToChineseNumeral(intRaw, style);

  if (!fracRaw) return sign + intText;
  // Spell the fraction digit-by-digit after 点; `digits[0]` is the standalone zero
  // glyph (○ 小写 / 零 大写), so a "0" fraction digit uses the right form.
  let fracText = "点";
  for (const ch of fracRaw) {
    fracText += style.digits[ch.charCodeAt(0) - 48];
  }
  return sign + intText + fracText;
}

/**
 * Render a number as 人民币大写 (uppercase RMB). Spells the integer yuan in 大写
 * numerals + 元, then 角 (tenths) and 分 (hundredths) for the two-decimal cents.
 * Follows the formal financial conventions:
 *  - 整 is appended when there are no 角/分 (whole yuan): 1234 → …元整.
 *  - 零 bridges a missing 角 when 分 is present (1.05 → 壹元零伍分).
 *  - 角 with zero 分 does not append 分 (1.50 → 壹元伍角); 分 with zero 角 inserts 零.
 *  - A zero amount renders 零元整; a negative gains a 负 prefix.
 *  - Sub-分 precision is rounded to the nearest 分 (2 decimals).
 * PURE — the cents arithmetic uses an integer count of 分 to avoid float drift.
 */
function renderRmbUpper(value: number): string {
  const sign = value < 0 ? "负" : "";
  const abs = Math.abs(value);
  // Total 分 (hundredths), rounded to the nearest 分; integer math then keeps the
  // 元/角/分 split exact. (A value sitting exactly on a half-分 IEEE-754 boundary,
  // e.g. 1.005, rounds however its float representation lands — the same caveat
  // Excel itself has.)
  const totalFen = Math.round(abs * 100);
  const yuan = Math.floor(totalFen / 100);
  const jiao = Math.floor((totalFen % 100) / 10);
  const fen = totalFen % 10;

  const yuanText = intStringToChineseNumeral(String(yuan), CN_UPPER_STYLE);
  let out = yuanText + "元";

  if (jiao === 0 && fen === 0) {
    return sign + out + "整";
  }
  if (jiao === 0 && yuan > 0) {
    // 分 present but no 角 after a non-zero yuan: insert a bridging 零 (壹元零伍分).
    out += "零";
  }
  if (jiao > 0) out += CN_UPPER_DIGITS[jiao] + "角";
  if (fen > 0) out += CN_UPPER_DIGITS[fen] + "分";
  return sign + out;
}

/**
 * Detect a Chinese-numeral directive on a format section. Excel/WPS write a
 * `[DBNum1]` / `[DBNum2]` bracket (optionally with `[$-zh-CN]`/`[$-804]`). The
 * 人民币大写 currency preset is a `[DBNum2]` format that ALSO spells a yuan amount
 * — recognised by the sentinel `人民币` marker or an `"元"` literal alongside DBNum2.
 * Returns the kind so {@link applyNumberFormat} can route before the date path.
 */
function chineseNumeralKind(
  section: string,
): "rmb" | "dbnum3" | "dbnum2" | "dbnum1" | null {
  const hasDb1 = /\[dbnum1\]/i.test(section);
  const hasDb2 = /\[dbnum2\]/i.test(section);
  const hasDb3 = /\[dbnum3\]/i.test(section);
  if (!hasDb1 && !hasDb2 && !hasDb3) return null;
  // RMB大写: a DBNum2 format carrying the 人民币 marker or a yuan literal.
  if (hasDb2 && (/人民币/.test(section) || /[元圆]/.test(section))) return "rmb";
  if (hasDb3) return "dbnum3";
  if (hasDb2) return "dbnum2";
  return "dbnum1";
}

//: Digit-by-digit glyphs. DBNum1's zero is ○ (U+25CB) and DBNum2's is 零 — MEASURED:
//: TEXT(1004,"[DBNum2]0") is 壹零零肆 while TEXT(101,"[DBNum1]0") is 一○一.
const DBNUM_DIGIT_GLYPHS: Record<"dbnum1" | "dbnum2" | "dbnum3", string[]> = {
  dbnum1: CN_LOWER_DIGITS,
  dbnum2: ["零", ...CN_UPPER_DIGITS.slice(1)],
  dbnum3: ["０", "１", "２", "３", "４", "５", "６", "７", "８", "９"],
};

/** Replace every ASCII digit of an already-rendered string, leaving the rest. */
function substituteDigits(
  rendered: string,
  kind: "dbnum1" | "dbnum2" | "dbnum3",
): string {
  const glyphs = DBNUM_DIGIT_GLYPHS[kind];
  return rendered.replace(/[0-9]/g, (digit) => glyphs[Number(digit)]);
}

//: The colloquial positional form a DATE field takes: 10 是 十 not 一十, and the
//: financial style keeps its leading 壹. MEASURED: [DBNum1] renders 十日/十五日/
//: 三十一日 and [DBNum2] renders 壹拾伍日.
function chineseDateNumber(value: number, financial: boolean): string {
  const digits = financial ? CN_UPPER_DIGITS : CN_LOWER_DIGITS;
  const ten = financial ? "拾" : "十";
  if (value < 10) return digits[value];
  const tens = Math.floor(value / 10);
  const ones = value % 10;
  const head = tens === 1 && !financial ? ten : digits[tens] + ten;
  return ones === 0 ? head : head + digits[ones];
}

/**
 * A date rendered with Chinese numerals. The YEAR is spelled digit by digit and
 * every other field takes the colloquial positional form — MEASURED in WPS:
 *   [DBNum1]yyyy年m月d日  二○二三年三月十五日   二○二三年十二月三十一日
 *                        二○二四年一月十日     二○二五年一月五日
 *   [DBNum2]yyyy年m月d日  贰零贰叁年叁月壹拾伍日
 *   [DBNum1]m月d日        三月十五日
 *   [DBNum1]aaaa          星期日   <- the weekday name is not a numeral
 * [DBNum3] is not here: it is a pure width change, so it goes through
 * substituteDigits like any other rendering.
 */
function renderChineseDateSection(
  serial: number,
  section: string,
  financial: boolean,
  date1904: ExcelDateSystemArg = false,
): string {
  const plain = renderDateTimeSection(serial, stripDbNumDirectives(section), date1904);
  const digits = financial ? CN_UPPER_DIGITS : CN_LOWER_DIGITS;
  // A four-digit run is a year; anything shorter is a month, day or clock field.
  return plain.replace(/[0-9]+/g, (run) =>
    run.length >= 4
      ? run
          .split("")
          .map((digit) => digits[Number(digit)])
          .join("")
      : chineseDateNumber(Number(run), financial),
  );
}

/** Drop the [DBNum*] bracket so the ordinary renderers see a plain picture. */
function stripDbNumDirectives(section: string): string {
  return section.replace(/\[dbnum[123]\]/gi, "");
}

/** True when the section carries an explicit numeric picture rather than General. */
function hasExplicitNumberPicture(section: string): boolean {
  return /[0#?]/.test(stripDbNumDirectives(section).replace(/\[[^\]]*\]/g, ""));
}

// ---------------------------------------------------------------------------
// Accounting number format (会计专用)
//
// Excel / WPS "Accounting" (会计专用) differs from plain Currency:
//   - The currency symbol is left-aligned (separated from the number by padding).
//   - Numbers are right-aligned with thousands separators.
//   - Negatives are shown in parentheses, e.g.  $ (1,234.50).
//   - Zero shows as a dash:  $ -
//   - Decimals are vertically aligned across rows.
//
// Excel writes this as a four-section format code whose positive section
// starts with `_(<symbol>*` (underscore+open-paren reserves symbol-width
// space, asterisk fills the gap between symbol and digits). Detecting this
// pattern lets us route accounting codes to a specialised renderer that
// produces symbol-left / number-right output even without pixel metrics.
//
// The canonical preset code stored per-cell is the Excel USD accounting code:
//   _($* #,##0.00_);_($* (#,##0.00)_);_($* "-"??_);_(@_)
// Custom codes for other currencies follow the same structural pattern.
// ---------------------------------------------------------------------------

/**
 * True when a positive section (already sign-selected) is an Accounting-style
 * format: `_(<symbol>*` at the front after any leading bracket. This covers the
 * canonical Excel/WPS pattern for all currency accounting codes.
 */
function isAccountingSection(section: string): boolean {
  // Strip leading bracket directives ([Red], [$-409], etc.) and whitespace, then
  // look for the structural  _( <currency> * sequence that marks accounting codes.
  // The currency may be $, €, ¥, £, ₹, or a bracket-enclosed [$<sym>-<locale>].
  const stripped = section.replace(/^\s*(\[[^\]]*\]\s*)*/g, "");
  return /^_\((\[[$][^\]]*\]|[$€¥£₹])\s*\*/.test(stripped)
    || /^_\(\s*[$€¥£₹]\s*\*/.test(stripped);
}

/**
 * Extract the currency symbol and decimal count from an Accounting positive
 * section. Returns defaults ($, 2) if parsing fails.
 */
function accountingParams(section: string): { symbol: string; decimals: number } {
  // Currency symbol: either a bracket locale form [$<sym>-<locale>] or a bare glyph.
  const bracketMatch = section.match(/\[[$]([^-\]]*)/);
  // Without a bracket form the glyph decides: defaulting to "$" here drew
  // every bare ¥ / € / £ accounting code as dollars (and Increase Decimal
  // rebuilt the stored code with $).
  let symbol = bracketMatch ? bracketMatch[1] : "";
  if (!symbol) {
    const glyphMatch = section.match(/[$€¥£₹]/);
    symbol = glyphMatch ? glyphMatch[0] : "$";
  }
  if (!symbol) symbol = "$";
  // Decimal count: look for `0.00` or `0.0` in the number spec portion.
  const numMatch = section.match(/[0#?],[0#?]*\.([0#?]+)/);
  const decimals = numMatch ? numMatch[1].length : 2;
  return { symbol, decimals };
}

/**
 * Render a numeric value through an Accounting format. Produces:
 *   Positive:  `$ 1,234.50 ` (symbol left, number right, trailing space)
 *   Negative:  `$(1,234.50)` (symbol left, parens around number)
 *   Zero:      `$ -         ` (symbol left, dash instead of 0.00)
 * The symbol is followed by a single space (approximation of Excel's fill
 * alignment; pixel-level alignment requires CSS/table layout). PURE.
 */
export function formatAccounting(
  value: number,
  symbol: string,
  decimals: number,
): string {
  // Round to the requested precision (decimal half-up at 15 significant
  // digits, like the plain number format: 2.675 -> 2.68, live audit B133) so
  // -0.001 with 2dp correctly shows zero.
  const rounded = Number(toFixedDecimal(value, decimals));
  const abs = Math.abs(rounded);
  const intRaw = Math.floor(abs).toString();
  const intGrouped = groupThousands(intRaw);
  const fracPart = decimals > 0 ? toFixedDecimal(abs, decimals).split(".")[1] : "";
  const numStr = fracPart ? `${intGrouped}${drawnSeparators.decimal}${fracPart}` : intGrouped;

  if (rounded === 0) {
    // Zero: symbol + space + dash + trailing space (matches Excel visual).
    return `${symbol} -`;
  }
  if (rounded < 0) {
    // Negative: symbol + space + open-paren + number + close-paren.
    return `${symbol} (${numStr})`;
  }
  // Positive: symbol + space + number + trailing space.
  return `${symbol} ${numStr}`;
}

/**
 * Render a value through a full Accounting format code (all 4 sections).
 * Routes the value's sign to the correct section's symbol/decimal params.
 */
function renderAccountingFormat(value: number, format: string): string {
  const sections = splitSections(format);
  // Positive section (section 0) carries the currency symbol and decimal count.
  const posSection = sections[0] ?? "";
  const { symbol, decimals } = accountingParams(posSection);
  return formatAccounting(value, symbol, decimals);
}

/** A scientific/exponential format carries an `E+`/`E-` token before a digit. */
function isScientificSection(section: string): boolean {
  return /[eE][+-]?[0#?]/.test(stripLiterals(escapeBarePictureLiterals(section)));
}

/** Render a number through a scientific format like "0.00E+00" / "##0.0E+0". */
function renderScientificSection(abs: number, section: string, autoMinus: boolean): string {
  // A bare literal between the placeholders stays where it stands (X-228).
  section = escapeBarePictureLiterals(section);
  // Isolate the "<mantissa>E<exp>" core from any surrounding literal text
  // ($ prefix, parens, " units" suffix). Inside the core: digits / grouping /
  // decimal point always count, and a +/- only when it directly follows the E
  // (so a trailing ")" is a suffix literal, not part of the exponent).
  const { prefix, core: spec, suffix, literals } = splitCore(section, (ch, prevCore) => {
    if (/[0#?.,eE]/.test(ch)) return true;
    if ((ch === "+" || ch === "-") && (prevCore === "e" || prevCore === "E")) {
      return true;
    }
    return false;
  });
  const pre = decodeLiterals(prefix);
  const suf = decodeLiterals(suffix);
  const eIdx = spec.search(/[eE]/);
  const mantissaFmt = spec.slice(0, eIdx);
  // The modifiers the plain number format applies apply here too (live audit
  // B441): each unescaped % multiplies by 100, and each comma ending the
  // mantissa divides by 1000, BEFORE the exponent is taken.
  const percents = (stripLiterals(section).match(/%/g) ?? []).length;
  const scaleCommas = (mantissaFmt.match(/,+$/) ?? [""])[0].length;
  abs = (abs * 100 ** percents) / 1000 ** scaleCommas;
  const expPart = spec.slice(eIdx + 1);
  const expSignAlways = expPart.startsWith("+");
  const dot = mantissaFmt.indexOf(".");
  const intFmt = dot >= 0 ? mantissaFmt.slice(0, dot) : mantissaFmt;
  const fracFmt = dot >= 0 ? mantissaFmt.slice(dot + 1) : "";
  const intDigits = Math.max(1, (intFmt.match(/[0#?]/g) || []).length);
  const decimals = (fracFmt.match(/[0#?]/g) || []).length;
  const expDigits = Math.max(1, (expPart.match(/[0#?]/g) || []).length);
  const sign = autoMinus ? "-" : "";
  // Literals quoted/escaped between the placeholders keep their position (B472):
  // the mantissa's integer digits fill right-to-left around them, its decimals
  // left-to-right, and the exponent digits right-to-left — the same placement
  // the plain number renderer uses.
  const mantLiterals = literals.filter((literal) => literal.at <= eIdx);
  const expLiterals = literals
    .filter((literal) => literal.at > eIdx)
    .map((literal) => ({ at: literal.at - eIdx - 1, text: literal.text }));
  const placeMantissa = (mant: string) => {
    if (!mantLiterals.length) return drawnNumeral(mant);
    const mantDot = mant.indexOf(".");
    const intText = mantDot >= 0 ? mant.slice(0, mantDot) : mant;
    const fracText = mantDot >= 0 ? mant.slice(mantDot + 1) : "";
    const intLits = mantLiterals.filter((literal) => dot < 0 || literal.at <= dot);
    const fracLits = mantLiterals
      .filter((literal) => dot >= 0 && literal.at > dot)
      .map((literal) => ({ at: literal.at - dot - 1, text: literal.text }));
    const intOut = intLits.length ? spliceIntegerLiterals(intText, intFmt, intLits) : intText;
    if (mantDot < 0) {
      // Every decimal was dropped (or none exist): literals written after the
      // point still print, after the integer digits.
      return intOut + fracLits.map((literal) => literal.text).join("");
    }
    const fracOut = fracLits.length
      ? spliceFractionLiterals(fracText.split(""), fracFmt, fracLits)
      : fracText;
    return `${intOut}${drawnSeparators.decimal}${fracOut}`;
  };
  const placeExponent = (expSign: string, digits: string) => {
    if (!expLiterals.length) return `E${expSign}${digits}`;
    // The exponent spec may open with its +/- sign; literals before the
    // first exponent digit stand before the sign, which Excel writes next to
    // the digits (`E+|#` shows E|+5, `E-|#` E|-5), the rest sit among them.
    const signWidth = /^[+-]/.test(expPart) ? 1 : 0;
    const head = expLiterals
      .filter((literal) => literal.at <= signWidth)
      .map((literal) => literal.text)
      .join("");
    const digitLits = expLiterals
      .filter((literal) => literal.at > signWidth)
      .map((literal) => ({ at: literal.at - signWidth, text: literal.text }));
    const body = digitLits.length
      ? spliceIntegerLiterals(digits, expPart.slice(signWidth), digitLits)
      : digits;
    return `E${head}${expSign}${body}`;
  };

  if (abs === 0) {
    const mant = decimals > 0 ? `0.${"0".repeat(decimals)}` : "0";
    const exp0 = placeExponent(expSignAlways ? "+" : "", "0".repeat(expDigits));
    return `${sign}${pre}${placeMantissa(mant)}${exp0}${suf}`;
  }
  let exp = Math.floor(Math.log10(abs));
  // A grouping comma between the mantissa's integer placeholders ("0,00E+00")
  // turns engineering notation off: Excel keeps one integer digit and pads it
  // with the zero placeholders, 44227 -> 004E+04 (X-305, r68 SP041).
  const groupedMantissa = /[0#?],+[0#?]/.test(intFmt);
  // Engineering notation: >1 integer placeholder groups the exponent so the
  // mantissa carries up to `intDigits` integer digits (exponent a multiple of it).
  if (intDigits > 1 && !groupedMantissa) exp = intDigits * Math.floor(exp / intDigits);
  // The mantissa rounds like every other digit run (decimal, not binary:
  // 2.675E+00 shows 2.68E+00, live audit B133).
  let mantStr = toFixedDecimal(abs / 10 ** exp, decimals);
  // Rounding can push the mantissa to 10^intDigits (e.g. 9.99 -> "10.0"); bump.
  const mantissaStep = groupedMantissa ? 1 : intDigits;
  if (Number(mantStr) >= 10 ** mantissaStep) {
    exp += mantissaStep;
    mantStr = toFixedDecimal(abs / 10 ** exp, decimals);
  }
  if (groupedMantissa) {
    const zeros = (intFmt.match(/0/g) ?? []).length;
    const point = mantStr.indexOf(".");
    const whole = point >= 0 ? mantStr.slice(0, point) : mantStr;
    mantStr = whole.padStart(zeros, "0") + (point >= 0 ? mantStr.slice(point) : "");
  }
  // `#` shows a fraction digit only when it is significant and `?` keeps its
  // width as a space; only `0` forces the digit (1.2 under 0.0#E+0 is 1.2E+0,
  // under 0.0?E+0 is "1.2 E+0", live audit B308).
  const fracPlaceholders = fracFmt.match(/[0#?]/g) ?? [];
  const mantDot = mantStr.indexOf(".");
  if (mantDot >= 0 && fracPlaceholders.some((ch) => ch !== "0")) {
    let digits = mantStr.slice(mantDot + 1);
    let trailing = "";
    for (let i = digits.length - 1; i >= 0; i -= 1) {
      const placeholder = fracPlaceholders[i] ?? "0";
      if (digits[i] !== "0" || placeholder === "0") break;
      digits = digits.slice(0, i);
      if (placeholder === "?") trailing = ` ${trailing}`;
    }
    mantStr = `${mantStr.slice(0, mantDot)}${digits ? `.${digits}` : ""}${trailing}`;
  }
  const expSign = exp < 0 ? "-" : expSignAlways ? "+" : "";
  const expStr = String(Math.abs(exp)).padStart(expDigits, "0");
  return `${sign}${pre}${placeMantissa(mantStr)}${placeExponent(expSign, expStr)}${suf}`;
}

/** A fraction format has a digit placeholder, a slash, then a denominator. */
function isFractionSection(section: string): boolean {
  return /[0#?] *\/ *[0-9#?]/.test(stripLiterals(section));
}

/**
 * The fraction a "?/??" placeholder denominator shows: the LAST continued-
 * fraction convergent of `x` whose denominator fits in `maxDen` — not the best
 * rational approximation under that bound (XLS-A034). MEASURED in Excel over
 * 168 values and the "# ?/?", "# ??/??" and "# ???/???" codes: every result
 * is that convergent. 0.14159… through "??/??" is 1/7 (the next convergent,
 * 15/106, needs three digits) where the closest two-digit fraction is 14/99,
 * and 0.0909 through "?/?" is 0/1 although 1/9 is nearer.
 */
function continuedFractionConvergent(x: number, maxDen: number): { num: number; den: number } {
  let num = Math.round(x);
  let den = 1;
  // h(-2)/k(-2) = 0/1 and h(-1)/k(-1) = 1/0 seed the convergent recurrence.
  let prevNum = 0;
  let prevDen = 1;
  let curNum = 1;
  let curDen = 0;
  let rest = x;
  for (let step = 0; step < 64; step += 1) {
    const term = Math.floor(rest);
    const nextNum = term * curNum + prevNum;
    const nextDen = term * curDen + prevDen;
    if (nextDen > maxDen) break;
    prevNum = curNum;
    prevDen = curDen;
    curNum = nextNum;
    curDen = nextDen;
    num = curNum;
    den = curDen;
    const remainder = rest - term;
    if (remainder < 1e-12 || Math.abs(x - curNum / curDen) < 1e-15) break;
    rest = 1 / remainder;
  }
  return { num, den };
}

/** Render a number as a (mixed or improper) fraction like "# ?/?" / "# ?/8". */
function renderFractionSection(abs: number, section: string, autoMinus: boolean): string {
  // Isolate the "<int> <num>/<den>" core from surrounding literal text so a
  // suffix like " cups" or a "$" prefix survives. Interior spaces (between the
  // integer and numerator) are kept by the firstCore..lastCore slice.
  const { prefix, core: spec, suffix, literals } = splitCore(section, (ch) =>
    /[0#?/0-9]/.test(ch),
  );
  const pre = decodeLiterals(prefix);
  const suf = decodeLiterals(suffix);
  const slash = spec.indexOf("/");
  const left = spec.slice(0, slash);
  const denFmt = spec.slice(slash + 1).trim();
  // The integer format (if any) precedes the numerator, separated by a space
  // -- or by a literal standing between two placeholder runs: Excel writes
  // its own fraction formats with an escaped space (`#\ ?/?`), and any other
  // literal divides them as well (`#"-"?/?` shows 3-1/7). MEASURED in Excel
  // for Mac (X-213): `\-#\ #/#####` shows -PI() as -3 14093/99532, where
  // reading `#\ #` as one integer run ran 31268 and 9 together.
  const lastSpace = left.trimEnd().lastIndexOf(" ");
  const boundary =
    lastSpace >= 0
      ? lastSpace
      : literals
          .filter((literal) => literal.at > 0 && literal.at < slash
            && /[0#?]/.test(left.slice(0, literal.at)) && /[0#?]/.test(left.slice(literal.at)))
          .reduce((at, literal) => Math.max(at, literal.at), -1);
  const numStart = lastSpace >= 0 ? lastSpace + 1 : Math.max(0, boundary);
  const intFmt = boundary >= 0 ? left.slice(0, boundary) : "";
  const hasInteger = /[0#?]/.test(intFmt);
  // The text between the integer and the numerator: the space, and every
  // literal written where the numerator starts (`# "X"?/?` is " X").
  const separator = hasInteger
    ? (lastSpace >= 0 ? " " : "") + literals.filter((literal) => literal.at === numStart).map((literal) => literal.text).join("")
    : "";
  const sign = autoMinus ? "-" : "";
  const wrap = (inner: string) => `${sign}${pre}${inner}${suf}`;

  // A fixed denominator ("0/8", "?/16") keeps it for zero as for any value
  // that rounds to zero: 0/8, not 0/1.
  const fixedDenominator = /^\d+$/.test(denFmt) && !/^0+$/.test(denFmt);
  if (abs === 0 && !hasInteger && !fixedDenominator) return wrap("0/1");

  const whole = hasInteger ? Math.floor(abs) : 0;
  const fracVal = abs - whole;
  let num: number;
  let den: number;
  if (fixedDenominator) {
    // A fixed denominator (e.g. "?/8") is used as-is and NOT reduced.
    //
    // All-zeros is NOT a fixed denominator: `0` is a digit PLACEHOLDER, so
    // "# 0/0" asks for a computed denominator exactly like "# ?/?" does.
    // Reading it as the literal 0 fell back to a denominator of 1, which made
    // every proper fraction round to 0/1 and print as a blank whole number.
    den = Number(denFmt) || 1;
    num = Math.round(fracVal * den);
  } else {
    // Placeholder denominator: den <= 10^digits - 1. The digit count is CLAMPED
    // (a format string holds up to 255 chars); real fraction formats use 1-3
    // denominator digits and 6 is far beyond any genuine use.
    const MAX_FRACTION_DEN_DIGITS = 6;
    const denDigits = Math.min(
      MAX_FRACTION_DEN_DIGITS,
      Math.max(1, (denFmt.match(/[0#?]/g) || []).length),
    );
    ({ num, den } = continuedFractionConvergent(fracVal, 10 ** denDigits - 1));
  }

  let wholeOut = whole;
  if (hasInteger && den > 0 && num >= den) {
    wholeOut += Math.floor(num / den);
    num %= den;
  }
  // The numerator/denominator placeholder runs.
  const numFmt = left.slice(numStart);
  const slotsOf = (fmt: string): string[] => [...(fmt.match(/[0#?]/g) ?? [])];
  const intSlots = slotsOf(intFmt);
  const numSlots = slotsOf(numFmt);
  const denSlots = slotsOf(denFmt);

  // Literals quoted/escaped among the placeholders keep their position (B472):
  // `# "X"?/?` shows 1.25 as 1 X1/4. Integer and numerator digits fill
  // right-to-left around them, denominator digits left-to-right — a literal
  // sits where the code wrote it, exactly as in the plain number renderer.
  const intLits: EmbeddedLiteral[] = [];
  const numLits: EmbeddedLiteral[] = [];
  const denLits: EmbeddedLiteral[] = [];
  for (const literal of literals) {
    if (literal.at > slash) {
      denLits.push({ at: literal.at - slash - 1, text: literal.text });
    } else if (hasInteger && literal.at === numStart) {
      continue; // part of the separator
    } else if (hasInteger && literal.at < numStart) {
      intLits.push(literal);
    } else {
      numLits.push({ at: Math.max(0, literal.at - numStart), text: literal.text });
    }
  }
  const placeInt = (text: string) =>
    intLits.length ? spliceIntegerLiterals(text, intFmt, intLits) : text;
  const placeNum = (text: string) =>
    numLits.length ? spliceIntegerLiterals(text, numFmt, numLits) : text;
  // A fixed denominator's digits are positions too ("?/1"X"6" keeps X after 1).
  const denSpecRaw = spec.slice(slash + 1).replace(/[1-9]/g, "0");
  const placeDen = (text: string) =>
    denLits.length ? spliceFractionLiterals(text.split(""), denSpecRaw, denLits) : text;

  if (!hasInteger) return wrap(`${placeNum(`${wholeOut * den + num}`)}/${placeDen(`${den}`)}`);

  /**
   * Fill the placeholder positions the digits did not reach. This is the whole
   * point of `?`: it holds a position open with a SPACE so a column of
   * fractions lines up on the slash; `0` pads with a zero and `#` with
   * nothing. A numerator is right-aligned, so its padding goes before it; a
   * denominator's spaces go after it and its zeros before it -- MEASURED in
   * Excel: TEXT(0.125,"# ??/?") "  1/8", TEXT(1.25,"# ??/??") "1  1/4 ",
   * TEXT(0.125,"# 00/00") "01/08".
   */
  const fill = (digits: string, slots: string[], side: "left" | "right") => {
    const missing = slots.length - digits.length;
    if (missing <= 0) return digits;
    const unused = side === "left" ? slots.slice(0, missing) : slots.slice(-missing);
    if (side === "left") return unused.map((c) => (c === "?" ? " " : c === "0" ? "0" : "")).join("") + digits;
    const zeros = unused.filter((c) => c === "0").length;
    const spaces = unused.filter((c) => c === "?").length;
    return "0".repeat(zeros) + digits + " ".repeat(spaces);
  };
  const spacePadded = numSlots.includes("?") || denSlots.includes("?");

  if (num === 0 && !numSlots.includes("0")) {
    // A whole number in a fraction format: the whole part shows, even a zero
    // one, and the fraction BLANKS -- to spaces of its width when a `?`
    // holds the column open, to nothing otherwise. MEASURED in Excel:
    // TEXT(2,"# ?/?") "2    ", 0.05 through "# ?/?" "0    " (XLS-A034),
    // TEXT(2,"# ?/8") "2    ", but TEXT(2,"# #/#") and TEXT(2,"# #/8") "2".
    // A `0` numerator is never blank: TEXT(2,"# 0/0") is "2 0/1".
    if (!spacePadded) return wrap(placeInt(`${wholeOut}`));
    if (!literals.length) {
      const denWidth = (denFmt.match(/[0-9#?]/g) ?? []).length;
      const width = separator.length + numSlots.length + 1 + denWidth;
      return wrap(`${wholeOut}` + " ".repeat(width));
    }
    return wrap(
      `${placeInt(`${wholeOut}`)}${" ".repeat(separator.length)}`
        + `${placeNum(" ".repeat(numSlots.length))} ${placeDen(" ".repeat(denSlots.length))}`,
    );
  }

  // A '#'/'?' integer placeholder (no forcing '0') suppresses a ZERO whole
  // part: "# ?/?" of 0.125 is " 1/8" and "0 ?/?" "0 1/8". The separator then
  // holds its width in spaces when the numerator has a `?` and goes with the
  // whole part otherwise, and a `?` integer slot keeps its space -- MEASURED
  // in Excel: 0.25 through `#"and"?/?` "   1/4", `# "X"?/?` "  1/4",
  // `?\ ?/?` "  1/4", but `# #/#####`, `# 0/0` and `#"-"#/#` "1/4".
  const forcesInteger = intSlots.includes("0");
  const suppressed = wholeOut === 0 && !forcesInteger;
  const intOut = suppressed ? fill("", intSlots, "left") : placeInt(`${wholeOut}`);
  const separatorOut = !suppressed
    ? separator
    : numSlots.includes("?") ? " ".repeat(separator.length) : "";
  return wrap(
    `${intOut}${separatorOut}${placeNum(fill(`${num}`, numSlots, "left"))}`
      + `/${placeDen(fill(`${den}`, denSlots, "right"))}`,
  );
}

/**
 * Render a TEXT section (one containing the `@` placeholder) for a value whose
 * General text is `text`. `@` is replaced by that text; quoted "…" runs and `\x`
 * escapes are literals; `_x` emits a space and `*x` is skipped (alignment, no
 * pixel metrics here). Used when a NUMERIC cell carries a Text ("@") custom
 * format — Excel/Sheets echo the value as text rather than printing a literal @.
 */
function renderTextSection(text: string, section: string): string {
  let out = "";
  for (let i = 0; i < section.length; i += 1) {
    const ch = section[i];
    if (ch === "@") {
      out += text;
    } else if (ch === "[") {
      // A bracketed directive — color ([Red]), condition ([<100]) or locale
      // ([$-409]) — never appears in the rendered text; skip to its `]`. A `[`
      // with no closing bracket is treated as a literal (Excel tolerates it).
      const end = section.indexOf("]", i + 1);
      if (end === -1) {
        out += ch;
      } else {
        i = end;
      }
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      if (end === -1) {
        out += section.slice(i + 1);
        break;
      }
      out += section.slice(i + 1, end);
      i = end;
    } else if (ch === "\\" && i + 1 < section.length) {
      out += section[i + 1];
      i += 1;
    } else if ((ch === "_" || ch === "*") && i + 1 < section.length) {
      if (ch === "_") out += " ";
      i += 1;
    } else if (ch === "_" || ch === "*") {
      // A stored code is trimmed, so zh-CN Accounting's closing `_ ` reserve
      // arrives as a lone `_`: still the reserved space, never a literal `_`
      // (" 甲乙_" under `_ @_`). A trailing `*` repeats nothing.
      if (ch === "_") out += " ";
    } else {
      out += ch;
    }
  }
  return out;
}

/** What a cell shows for a date or time Excel cannot write. */
export const DATE_TIME_OUT_OF_RANGE_DISPLAY = "########";

/**
 * Whether a date/time section cannot write `value`: past 9999-12-31 in the
 * workbook's system (serial 2958465 in 1900, 2957003 in 1904), or a negative
 * value with no negative section in the 1900 system. The 1904 system writes a
 * negative date or time as its magnitude with a minus sign (measured in
 * Excel: -1 is -1904-01-02, -0.25 as h:mm is -6:00), so only the 1900 system
 * refuses one.
 */
function dateTimeOutOfRange(
  value: number,
  autoMinus: boolean,
  date1904: ExcelDateSystemArg = false,
): boolean {
  return (autoMinus && !isExcelDate1904(date1904)) || Math.abs(value) >= excelMaxDateSerial(date1904) + 1;
}

/**
 * Whether `format` would show `value` through a date/time section it cannot
 * write (a negative value with no negative section, or a date past 9999). The
 * TEXT function answers #VALUE! for it where a cell shows ######.
 */
export function numberFormatDateTimeOutOfRange(
  value: number,
  format: string,
  options?: SpreadsheetDateSystemOptions,
): boolean {
  if (!Number.isFinite(value)) return false;
  const raw = format || "";
  const f = raw.trim();
  if (!f || f.toLowerCase() === "general") return false;
  const sections = splitSections(raw);
  if (sections.length > 0 && isAccountingSection((sections[0] ?? "").trim())) return false;
  const { section, autoMinus } = pickSection(value, sections);
  const sec = section.trim();
  if (sec === "" || sec.toLowerCase() === "general" || isTextPlaceholderSection(sec)) return false;
  if (chineseNumeralKind(sec)) return false;
  return isDateTimeSection(sec) && dateTimeOutOfRange(value, autoMinus, options?.date1904);
}

/**
 * `value` rendered through a number format code. `options.date1904` reads a
 * date/time section in the workbook's 1904 date system (serial 0 is
 * 1904-01-01); every number-only section is identical in both systems.
 */
/** Rendered numbers by value, format, date system and separators. Rendering
 * is pure, and a grid asks for the same few thousand again on each pass over
 * the screen: half a second of an open of 250,000 formatted cells. */
const renderedNumbers = new Map<string, string>();
const MAX_RENDERED_NUMBERS = 50_000;

export function applyNumberFormat(
  value: number,
  format: string,
  options?: NumberFormatRenderOptions,
): string {
  // -0 reads as "0" in a key and prints its sign under some formats.
  if (Object.is(value, -0)) return renderWithSeparators(value, format, options);
  const separators = options?.separators ?? drawnSeparators;
  const key = `${value}\u0000${format}\u0000${String(options?.date1904)}\u0000${separators.decimal}${separators.group}`;
  const known = renderedNumbers.get(key);
  if (known !== undefined) return known;
  const text = renderWithSeparators(value, format, options);
  if (renderedNumbers.size >= MAX_RENDERED_NUMBERS) renderedNumbers.clear();
  renderedNumbers.set(key, text);
  return text;
}

function renderWithSeparators(
  value: number,
  format: string,
  options?: NumberFormatRenderOptions,
): string {
  const separators = options?.separators;
  if (!separators || sameSeparators(separators, drawnSeparators)) return renderNumberFormat(value, format, options);
  const outer = drawnSeparators;
  drawnSeparators = separators;
  try {
    return renderNumberFormat(value, format, options);
  } finally {
    drawnSeparators = outer;
  }
}

function renderNumberFormat(
  value: number,
  format: string,
  options?: NumberFormatRenderOptions,
): string {
  if (!Number.isFinite(value)) return String(value);
  const date1904 = options?.date1904 ?? false;
  // A space in a format code is a literal Excel prints, at either end: a
  // TRAILING space is how `0.00_ ;-0.00 ` lines a negative up under the width
  // `_)` reserves on a positive (MEASURED: -5 through that code is "-5.00 "),
  // and a LEADING one is the " 12" that TEXT(12," 0") returns (its LEN is 3).
  // So split the sections from the untrimmed code and let the plain-number
  // renderer see its section's literals; only the General test trims.
  const raw = format || "";
  const f = raw.trim();
  if (!f || f.toLowerCase() === "general") return drawnGeneralNumber(value);

  const sections = splitSections(raw);

  // Accounting format (会计专用): detected on the POSITIVE section (section 0)
  // so the routing applies to the entire format string regardless of which
  // section the value's sign selects. Route before sign-based section picking
  // so the accounting renderer owns all three value cases (positive/negative/zero).
  //
  // Only when every section the value could select is an accounting one: a
  // code whose negative or zero section is its own (`_($* #,##0.00_);"DEBT
  // "0.000;"ZERO"`) shows those sections as written (live audit B443).
  const accountingThroughout = sections
    .slice(0, 3)
    .every((part) => part.trim() === "" || isAccountingSection(part.trim()));
  if (sections.length > 0 && isAccountingSection((sections[0] ?? "").trim()) && accountingThroughout) {
    return renderAccountingFormat(value, f);
  }

  const { section, autoMinus, unmatched } = pickSection(value, sections);
  if (unmatched) return DATE_TIME_OUT_OF_RANGE_DISPLAY;
  const sec = section.trim();
  // The value's own section is an accounting one among sections that are not.
  if (isAccountingSection(sec)) return renderAccountingFormat(value, sec);

  // An explicitly empty section hides the value (Excel's ";;" idiom).
  if (sec === "") return "";
  if (sec.toLowerCase() === "general") {
    return `${autoMinus ? "-" : ""}${drawnGeneralNumber(Math.abs(value))}`;
  }
  // A Text ("@") section echoes the value's General text at the placeholder (the
  // signed General text already carries any minus sign). This MUST be tested
  // before the date/scientific/fraction probes: a literal like "Mr. @" otherwise
  // matches isDateTimeSection (the "M" reads as a month token) and renders the
  // value as a bogus date ("5r. @"). The `@` is the unambiguous text marker.
  if (isTextPlaceholderSection(sec)) {
    return renderTextSection(drawnGeneralNumber(value), sec);
  }
  // Chinese-numeral directives ([DBNum1]/[DBNum2], 人民币大写) MUST be probed
  // before the date path: a `[DBNum2]…G/通用格式"元"` code carries `m`/`d`-free
  // text, but a numeral format like `[DBNum1]0` would otherwise fall through to
  // the plain numeric renderer and lose the spelled-out form. The directive owns
  // the whole section, so this routes on the SELECTED section (sign-correct).
  const cnKind = chineseNumeralKind(sec);
  if (cnKind === "rmb") return renderRmbUpper(value);
  if (cnKind) {
    const bare = stripDbNumDirectives(sec);
    // A DATE picture keeps its fields and converts them one by one; [DBNum3] is
    // only a width change so it needs no field awareness.
    if (isDateTimeSection(bare)) {
      return cnKind === "dbnum3"
        ? substituteDigits(renderDateTimeSection(value, bare, date1904), "dbnum3")
        : renderChineseDateSection(value, sec, cnKind === "dbnum2", date1904);
    }
    // An explicit numeric picture is rendered normally and then transliterated
    // DIGIT BY DIGIT — MEASURED: TEXT(1234,"[DBNum1]0") is 一二三四, not
    // 一千二百三十四, and TEXT(1234,"[DBNum1]#,##0") keeps the comma (一,二三四).
    if (hasExplicitNumberPicture(sec)) {
      return substituteDigits(applyNumberFormat(value, bare || "General", options), cnKind);
    }
    // No picture at all ([DBNum1] / [DBNum1]General) is the POSITIONAL form —
    // MEASURED: those render 一千二百三十四 and 四万五千.
    if (cnKind === "dbnum3") return substituteDigits(drawnGeneralNumber(value), "dbnum3");
    return renderChineseNumeralSection(value, sec, cnKind === "dbnum2");
  }
  if (isDateTimeSection(sec)) {
    // Excel has no negative dates or times, nor dates past 9999-12-31: such a
    // value in a date/time section shows as ###### instead of a clock that
    // wrapped to the day before (-1 hour read 23:00:00). An explicit negative
    // section formats the magnitude, as Excel does.
    if (dateTimeOutOfRange(value, autoMinus, date1904)) return DATE_TIME_OUT_OF_RANGE_DISPLAY;
    const rendered = renderDateTimeSection(Math.abs(value), sec, date1904);
    // Only the 1904 system reaches here with a negative value and no negative
    // section: Excel writes the magnitude behind a minus sign.
    return autoMinus ? `-${rendered}` : rendered;
  }
  if (isScientificSection(sec)) {
    return renderScientificSection(Math.abs(value), sec, autoMinus);
  }
  if (isFractionSection(sec)) {
    return renderFractionSection(Math.abs(value), sec, autoMinus);
  }

  return renderNumberSection(Math.abs(value), section, autoMinus);
}

export type NumberFormatLayoutPart =
  | { kind: "text"; text: string }
  | { kind: "reserve" | "fill"; character: string };

/** Retain the layout operators while reusing the existing value/section
 * renderer. `_x` reserves x's advance and `*x` fills the available width;
 * neither operator is an ordinary space in a saved or copied cell value.
 * Quoted/escaped operators remain literal. This applies to custom formats,
 * currencies and dates, without identifying particular Accounting presets. */
export function numberFormatLayout(
  value: number,
  format: string,
  options?: NumberFormatRenderOptions,
): NumberFormatLayoutPart[] | undefined {
  if (!Number.isFinite(value) || !format || format.length > 4096 || !/[_*]/.test(format)) return undefined;
  // Private-use markers are temporary quoted literals, never persisted. Pick
  // an unused prefix so user-authored literal text cannot become an operator.
  let marker = "\ue000";
  while (format.includes(marker)) marker += "\ue000";
  const operators = new Map<string, NumberFormatLayoutPart>();
  let encoded = "";
  for (let i = 0; i < format.length; i += 1) {
    const ch = format[i];
    if (ch === '"' || ch === "[") {
      const end = ch === '"' ? '"' : "]";
      encoded += ch;
      while (++i < format.length) {
        encoded += format[i];
        if (format[i] === end) break;
      }
    } else if (ch === "\\") {
      encoded += ch;
      if (i + 1 < format.length) encoded += format[++i];
    } else if ((ch === "_" || ch === "*") && i + 1 < format.length) {
      const character = String.fromCodePoint(format.codePointAt(i + 1)!);
      i += character.length;
      const token = `${marker}${String.fromCharCode(0xe100 + operators.size)}${marker}`;
      operators.set(token, { kind: ch === "_" ? "reserve" : "fill", character });
      encoded += `"${token}"`;
    } else encoded += ch;
  }
  if (!operators.size) return undefined;
  const rendered = applyNumberFormat(value, encoded, options);
  const parts: NumberFormatLayoutPart[] = [];
  let cursor = 0;
  while (cursor < rendered.length) {
    let next = rendered.length;
    let found: [string, NumberFormatLayoutPart] | undefined;
    for (const entry of operators) {
      const index = rendered.indexOf(entry[0], cursor);
      if (index >= 0 && index < next) { next = index; found = entry; }
    }
    if (next > cursor) parts.push({ kind: "text", text: rendered.slice(cursor, next) });
    if (!found) break;
    parts.push(found[1]);
    cursor = next + found[0].length;
  }
  // A layout projection must not quietly change an existing numeric result.
  // Specialized value renderers may own additional semantics (for example
  // rounded Accounting zero); retain that path until it supplies layout atoms.
  const visible = parts.filter(part => part.kind === "text").map(part => part.text).join("");
  if (visible.replace(/\s/gu, "") !== applyNumberFormat(value, format, options).replace(/\s/gu, "")) return undefined;
  if (parts.filter(part => part.kind === "fill").length > 1) return undefined;
  return parts.some(part => part.kind !== "text") ? parts : undefined;
}

/**
 * Format a TEXT value through a custom format's text section — the `@`
 * placeholder echoes the text (`TEXT("hi","@@@")` → "hihihi"; `"Mr. @"` →
 * "Mr. hi"). Excel's text section is the 4th `;`-section; a single-section format
 * counts as a text section only when it actually contains `@`. When the format
 * has no applicable text section the text passes through unchanged (Excel leaves
 * text alone under a purely numeric format). PURE.
 */
export function applyTextFormat(text: string, format: string): string {
  const f = (format || "").trim();
  if (!f || f.toLowerCase() === "general") return text;
  const sections = splitSections(f);
  const textSection =
    sections.length >= 4
      ? sections[3]
      : sections.length === 1 && isTextPlaceholderSection(sections[0])
        ? sections[0]
        : null;
  if (textSection === null) return text;
  return renderTextSection(text, textSection.trim());
}

/**
 * A format that is a text section alone (one section holding `@`) shapes text
 * only. MEASURED in Excel for Mac 16: 12 in a cell formatted `"文字:"@` or
 * `@"x"` reads 12, as General, while its text reads 文字:甲乙. TEXT() keeps
 * {@link applyNumberFormat}.
 */
export function isTextOnlyNumberFormat(format: string): boolean {
  const sections = splitSections((format || "").trim());
  return sections.length === 1 && isTextPlaceholderSection(sections[0]);
}

/**
 * What a cell shows for a number under `format`: {@link applyNumberFormat},
 * except that a text-only format ({@link isTextOnlyNumberFormat}) leaves the
 * number in General. PURE.
 */
export function applyCellNumberFormat(
  value: number,
  format: string,
  options?: NumberFormatRenderOptions,
): string {
  return Number.isFinite(value) && isTextOnlyNumberFormat(format)
    ? formatGeneralNumberWith(value, options?.separators)
    : applyNumberFormat(value, format, options);
}

/** True when the format string is one this renderer meaningfully handles. */
export function isNumberFormat(format: string | undefined): format is string {
  return typeof format === "string" && format.trim() !== "";
}

/** Excel's editable custom number-format box accepts at most 255 characters. */
export const MAX_NUMBER_FORMAT_CODE_LENGTH = 255;

export type NumberFormatFamily =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "date"
  | "time"
  | "percentage"
  | "fraction"
  | "scientific"
  | "text"
  | "custom";

/**
 * Validate and canonicalize the raw style authority used by Home > Number.
 * Missing/empty/General are the same stored General state. Runtime page data is
 * deliberately passthrough elsewhere, so a non-string or control-bearing value
 * must fail closed before a range command acquires mutation authority.
 */
export function normalizeNumberFormatCode(format: unknown): string | null {
  if (format === undefined) return "";
  if (typeof format !== "string") return null;
  let normalized = format.trim();
  // A trailing "_x" pads the width of character x and "\\x" writes x: when x is
  // a space, the space IS the code. "0.00_ " trimmed to "0.00_" is a dangling
  // padding Excel cannot read -- a formula that inherited it made Excel repair
  
  if (/(?:^|[^_\\])(?:_|\\)$/.test(normalized) && format.length > format.trimEnd().length) {
    normalized += format.slice(format.trimEnd().length, format.trimEnd().length + 1);
  }
  if (
    normalized.length > MAX_NUMBER_FORMAT_CODE_LENGTH ||
    /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(normalized)
  ) {
    return null;
  }
  if (
    normalized === "" ||
    normalized.toLowerCase() === "general" ||
    normalized === "G/通用格式" ||
    normalized === "通用格式"
  ) {
    return "";
  }
  return normalized;
}

/**
 * Whether the first picture is one of Home > Number's decimal/grouping forms.
 * Digit masks with embedded literals or repeated mandatory integer zeroes are
 * Excel Custom formats (for example `000-00-0000`), not Number. Keeping that
 * boundary explicit prevents Increase/Decrease Decimal from rewriting an
 * identifier mask as though it were a numeric presentation preset.
 */
function isStandardNumberSection(section: string): boolean {
  if (/["\\]/.test(section)) return false;
  const picture = section
    .replace(/\[[^\]]*\]/g, "")
    .replace(/_.?/g, "")
    .replace(/\*./g, "")
    .trim();
  return /^[#,]*0(?:\.[0#?]+)?$/.test(picture);
}

/** Semantic family projected by Excel's Home number-format combo. */
export function numberFormatFamily(format: unknown): NumberFormatFamily | null {
  const normalized = normalizeNumberFormatCode(format);
  if (normalized === null) return null;
  if (normalized === "") return "general";

  const section = splitSections(normalized)[0] ?? "";
  // The zh-CN (`_ ¥* #,##0.00_ `), en-GB/euro (`_-€* …`) and symbol-less
  // (`_(* …`) Accounting shapes are Accounting too (XLS-A053); they render
  // through the generic `_x` / `*x` layout path.
  if (
    isAccountingSection(section) ||
    /^_[ (-](?:\[\$[^\]]*\]|[^\s*"0#?;]*)\s*\*/.test(section.trimStart())
  ) {
    return "accounting";
  }
  if (isDateTimeSection(section)) {
    const tokens = stripLiterals(section)
      .replace(/[eE][+-]?[0#?]+/g, "")
      .replace(/general/gi, "")
      .replace(/G\/通用格式/g, "")
      .replace(/通用格式/g, "")
      // The A of AM/PM is not the Chinese weekday letter `a`: without this,
      // h:mm:ss AM/PM (Excel's en-US Time) read as a date.
      .replace(/AM\/PM/gi, "");
    const hasClock = /[hs]/i.test(tokens) || /AM\/PM/i.test(section) || /\[[hms]+\]/i.test(section);
    const hasDate = /[dyae]/i.test(tokens) || (!hasClock && /m/i.test(tokens));
    return hasDate ? "date" : "time";
  }

  const tokens = stripLiterals(section);
  if (/[eE][+-]?[0#?]+/.test(tokens)) return "scientific";
  if (tokens.includes("%")) return "percentage";
  if (isFractionSection(section)) return "fraction";
  if (tokens.includes("@") && !/[0#?]/.test(tokens)) return "text";
  if (/\[\$(?!-)[^\]]+\]|[$€¥£₹]/.test(section)) return "currency";
  if (isStandardNumberSection(section)) return "number";
  if (/[0#?]/.test(tokens)) return "custom";
  return "custom";
}

/** Whether the Comma Style face should be pressed for this live format. */
export function isCommaStyleNumberFormat(format: unknown): boolean | null {
  const family = numberFormatFamily(format);
  if (family === null) return null;
  // Excel's Comma style is the symbol-less Accounting shape (`_(* …` / `_ * …`).
  if (family === "accounting") {
    const code = normalizeNumberFormatCode(format) ?? "";
    return /^_[( ]\* /.test(code);
  }
  if (family !== "number") return false;
  const normalized = normalizeNumberFormatCode(format);
  if (!normalized) return false;
  return stripLiterals(splitSections(normalized)[0] ?? "").includes(",");
}

interface DecimalTokenScan {
  dotIndex: number;
  exponentIndex: number;
  fractionalPlaceholders: number[];
  lastMantissaPlaceholder: number;
}

/** Locate top-level decimal placeholders without counting exponent digits. */
function scanDecimalTokens(section: string): DecimalTokenScan {
  const tokens: Array<{ index: number; char: string }> = [];
  for (let index = 0; index < section.length; index += 1) {
    const char = section[index];
    if (char === '"') {
      index += 1;
      while (index < section.length && section[index] !== '"') index += 1;
    } else if (char === "\\") {
      index += 1;
    } else if (char === "[") {
      while (index < section.length && section[index] !== "]") index += 1;
    } else {
      tokens.push({ index, char });
    }
  }

  let exponentToken = -1;
  for (let index = 0; index < tokens.length; index += 1) {
    if (!/[eE]/.test(tokens[index]?.char ?? "")) continue;
    let next = index + 1;
    if (/[+-]/.test(tokens[next]?.char ?? "")) next += 1;
    if (/[0#?]/.test(tokens[next]?.char ?? "")) {
      exponentToken = index;
      break;
    }
  }
  const mantissaTokens =
    exponentToken >= 0 ? tokens.slice(0, exponentToken) : tokens;
  const dotToken = mantissaTokens.findIndex((token) => token.char === ".");
  const fractionalPlaceholders = mantissaTokens
    .slice(dotToken >= 0 ? dotToken + 1 : mantissaTokens.length)
    .filter((token) => /[0#?]/.test(token.char))
    .map((token) => token.index);
  const lastMantissaPlaceholder = [...mantissaTokens]
    .reverse()
    .find((token) => /[0#?]/.test(token.char))?.index ?? -1;
  return {
    dotIndex: dotToken >= 0 ? (mantissaTokens[dotToken]?.index ?? -1) : -1,
    exponentIndex:
      exponentToken >= 0 ? (tokens[exponentToken]?.index ?? -1) : -1,
    fractionalPlaceholders,
    lastMantissaPlaceholder,
  };
}

function adjustSectionDecimalPlaceholders(
  section: string,
  delta: number,
): string {
  if (delta === 0 || isDateTimeSection(section) || isFractionSection(section)) {
    return section;
  }
  const scan = scanDecimalTokens(section);
  if (scan.lastMantissaPlaceholder < 0) return section;
  const current = scan.fractionalPlaceholders.length;
  const next = Math.min(20, Math.max(0, current + delta));
  if (next === current) return section;

  if (next > current) {
    const addition = "0".repeat(next - current);
    if (scan.dotIndex >= 0) {
      const insertAt =
        scan.exponentIndex >= 0
          ? scan.exponentIndex
          : (scan.fractionalPlaceholders.at(-1) ?? scan.dotIndex) + 1;
      return section.slice(0, insertAt) + addition + section.slice(insertAt);
    }
    const insertAt = scan.lastMantissaPlaceholder + 1;
    return (
      section.slice(0, insertAt) +
      `.${addition}` +
      section.slice(insertAt)
    );
  }

  const remove = new Set(scan.fractionalPlaceholders.slice(next));
  let adjusted = "";
  for (let index = 0; index < section.length; index += 1) {
    if (!remove.has(index)) adjusted += section[index];
  }
  if (next === 0 && scan.dotIndex >= 0) {
    const adjustedDot = scan.dotIndex;
    adjusted = adjusted.slice(0, adjustedDot) + adjusted.slice(adjustedDot + 1);
  }
  return adjusted;
}

/** Count of decimal places a numeric format renders (0 for General/dates). */
export function formatDecimalPlaces(format: string | undefined): number {
  const f = (format || "").trim();
  if (!f || f.toLowerCase() === "general") return 0;
  // Use the positive (first) section so a negative/zero section's decimals do
  // not bleed in.
  const seg = splitSections(f)[0] ?? "";
  if (isDateTimeSection(seg)) return 0;
  return scanDecimalTokens(seg).fractionalPlaceholders.length;
}

/**
 * The decimal places a General cell holding `value` displays: 1.5 shows one,
 * 1.234 three, 3 none. Excel's Increase and Decrease Decimal start a General
 * cell from what it shows, so 1.5 goes to 0.00 (1.50) or to 0 (2). Text, a
 * logical, an error, an empty cell and a number General writes in scientific
 * notation count as none.
 */
export function generalDisplayDecimals(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 0;
  const shown = formatGeneralNumber(value);
  if (/e/i.test(shown)) return 0;
  const dot = shown.indexOf(".");
  return dot < 0 ? 0 : shown.length - dot - 1;
}

/**
 * Increase or decrease the decimal places of a numeric format by `delta`,
 * preserving each section's literals, grouping, currency/percent tokens and
 * scientific exponent. General starts from a plain `0` format with the
 * `generalDecimals` the cell currently displays (see
 * {@link generalDisplayDecimals}). Date/time, fraction, text and unsupported
 * custom formats are semantic no-ops, matching Excel's Home command instead of
 * flattening them into an unrelated number. Decimals are clamped to [0, 20]
 * (matching Number.prototype.toFixed limits).
 */
export function adjustFormatDecimals(
  format: string | undefined,
  delta: number,
  generalDecimals = 0,
): string {
  const normalized = normalizeNumberFormatCode(format);
  if (normalized === null) return (format || "").trim();
  const f = normalized;
  const seg = splitSections(f)[0] ?? "";
  if (!f) {
    const shown =
      Number.isSafeInteger(generalDecimals) && generalDecimals > 0
        ? generalDecimals
        : 0;
    const next = Math.min(20, Math.max(0, shown + delta));
    return next > 0 ? `0.${"0".repeat(next)}` : "0";
  }

  const family = numberFormatFamily(f);
  if (
    family === "date" ||
    family === "time" ||
    family === "fraction" ||
    family === "text" ||
    family === "custom" ||
    family === null
  ) {
    return f;
  }

  // An Accounting format is a 4-section structure (`_($* #,##0.00_);…`). Its
  // currency token sits AFTER the leading `_(`. Rebuild the full accounting
  // code at the new decimal count so the symbol, parenthesised negative and
  // dash-zero semantics stay intact.
  if (family === "accounting" && isAccountingSection(seg)) {
    const current = formatDecimalPlaces(f);
    const next = Math.min(20, Math.max(0, current + delta));
    return buildAccountingCodeFromParams(accountingParams(seg).symbol, next);
  }
  return splitSections(f)
    .map((section) => adjustSectionDecimalPlaceholders(section, delta))
    .join(";");
}
