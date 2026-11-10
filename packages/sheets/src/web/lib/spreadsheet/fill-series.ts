/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeHexColor } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import {
  excelSystemSerialFromUtcMs,
  excelSystemSerialToUtcMs,
  type SpreadsheetDateSystemOptions,
} from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";
import { numberFormatFamily } from "@mrnote/core/web/lib/spreadsheet/number-format.ts";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";
import { excelMinDateSerial } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

/**
 * Drag-to-fill series math (Excel/WPS parity), kept pure so the SpreadsheetWindow
 * fill-handle drag and the Ctrl+D fill-down shortcut share one tested
 * implementation. The component owns the model mutation (it writes the returned
 * values through the existing cell-update / paste path so undo + autosave fire);
 * this module only decides WHAT values a fill produces.
 *
 * Excel fill rules this models:
 *  - Single non-numeric source  -> copy the constant value `count` times,
 *                                  UNLESS it is a known TEXT SERIES token
 *                                  (weekday/month name, "Q1", or a label that
 *                                  starts or ends with a number, like "Item1"
 *                                  or "1st Period"), which extends.
 *  - Single numeric source      -> copy the constant value (no implied step).
 *  - Two+ numeric sources       -> extend the arithmetic series by its last step
 *                                  (e.g. 1,2 -> 3,4,5 ; 10,8 -> 6,4,2), any
 *                                  other run along its best-fit line
 *                                  (3,5,8 -> 10.33..., 12.83...).
 *  - Single date source         -> step by one day in the drag direction.
 *  - Two+ date sources          -> extend by months when every date falls on
 *                                  the same day of the month, else by days;
 *                                  each value keeps its seed's text shape
 *                                  ("2026/9/21" -> 2026/9/22).
 *  - Text-series sources        -> "Mon"->Tue,Wed; "Jan"->Feb; "Item1"->Item2;
 *                                  quarters wrap ("Q3"->Q4,Q1,Q2); a two-cell
 *                                  block sets the step ("Mon","Wed"->Fri,Sun).
 *  - Mixed / unparseable        -> cycle the source values (constant copy of the
 *                                  block), which is Excel's behaviour for a
 *                                  non-series selection.
 *
 * `direction` only flips the SIGN of the implied step for a *single* numeric or
 * date source (dragging up/left counts down, down/right counts up). For a
 * multi-cell source the step is already encoded by the source order, so the
 * caller is responsible for ordering `sourceCells` from the far edge toward the
 * fill so the extension continues naturally.
 */
export type FillDirection = "down" | "up" | "right" | "left";

const MS_PER_DAY = 86_400_000;

/**
 * Parse a cell's display string as a finite number, or null. Empty/whitespace
 * and non-numeric text (including things `Number()` coerces loosely, like "" or
 * "  ") are rejected so only a genuine numeric literal seeds a numeric series.
 */
function parseNumber(raw: string): number | null {
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

/**
 * Whether two consecutive numeric steps are equal within a small epsilon, so a
 * floating-point series like 0.1,0.2,0.3 is still treated as arithmetic.
 */
function nearlyEqual(a: number, b: number): boolean {
  const scale = Math.max(1, Math.abs(a), Math.abs(b));
  return Math.abs(a - b) <= 1e-9 * scale;
}

/** Step sign for a SINGLE-cell source dragged in a given direction. */
function directionSign(direction: FillDirection): 1 | -1 {
  return direction === "up" || direction === "left" ? -1 : 1;
}

// --- Text series (Excel built-in fill lists) --------------------------------
// Excel auto-extends a handful of built-in text lists (weekday + month names,
// long and short) and any trailing-number label. Each list is a CYCLIC ring, so
// dragging past the end wraps (Sun -> Mon, Dec -> Jan), matching Excel.

interface TextSeriesList {
  /** Lower-cased members in order; index 0 is the list's first item. */
  readonly lower: readonly string[];
  /** Canonical-cased members in order (used when echoing capitalised input). */
  readonly canonical: readonly string[];
}

const WEEKDAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const WEEKDAYS_LONG = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];
const MONTHS_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const MONTHS_LONG = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

// --- Built-in CJK (Chinese) fill lists (Excel zh-CN parity) ------------------
// The lists zh-CN Excel auto-extends on drag-fill, in the order its Edit Custom
// Lists dialog shows them. They are ordinary cyclic rings like the weekday/month
// lists; CJK characters have no upper/lower case, so `toLowerCase()` is the
// identity here and the casing logic simply echoes the canonical member verbatim
// ("title"/"as-is" both return it). Several lists share members (二月 is in both
// month lists, 一 is a weekday and a numeral): a single seed takes the FIRST list
// holding it, which is the order below, and a multi-cell seed takes the first
// list holding every seed.
//
// Microsoft Support "Create or delete a custom list for sorting and filling
// data" (zh-CN) names the built-in lists as 周日…周六, 星期日…星期六, 1月…12月
// and 一月…十二月. The dialog of zh-CN Excel also carries:
//  - CJK_WEEKDAY_NUMERALS: 日 一 二 … 六  (Sun..Sat as single characters)
//  - CJK_QUARTERS:     第一季 … 第四季
//  - LUNAR_MONTHS:     正月 二月 … 十一月 腊月  (the traditional month names)
//  - EARTHLY_BRANCHES: 子 丑 寅 … 亥 (the twelve 地支 Earthly Branches)
//  - HEAVENLY_STEMS:   甲 乙 丙 … 癸  (the ten 天干 Heavenly Stems)
// CJK_NUMERALS (一 … 十) is not an Excel list; it comes last so it only answers
// for the numerals the weekday list lacks (七 八 九 十), which WPS extends.
const CJK_WEEKDAYS_SHORT = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"];
const CJK_WEEKDAY_NUMERALS = ["日", "一", "二", "三", "四", "五", "六"];
const CJK_WEEKDAYS = [
  "星期日",
  "星期一",
  "星期二",
  "星期三",
  "星期四",
  "星期五",
  "星期六",
];
const CJK_MONTHS_ARABIC = Array.from({ length: 12 }, (_, i) => `${i + 1}月`);
const CJK_MONTHS = [
  "一月",
  "二月",
  "三月",
  "四月",
  "五月",
  "六月",
  "七月",
  "八月",
  "九月",
  "十月",
  "十一月",
  "十二月",
];
const LUNAR_MONTHS = ["正月", ...CJK_MONTHS.slice(1, 11), "腊月"];
const CJK_NUMERALS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
const HEAVENLY_STEMS = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"];
const EARTHLY_BRANCHES = [
  "子",
  "丑",
  "寅",
  "卯",
  "辰",
  "巳",
  "午",
  "未",
  "申",
  "酉",
  "戌",
  "亥",
];
const CJK_QUARTERS = ["第一季", "第二季", "第三季", "第四季"];

/**
 * The built-in CJK fill lists, exported so the UI can surface them (e.g. a
 * "fill with series" affordance or a read-only listing alongside the user's
 * custom lists). These are auto-detected by `computeFillValues` with no caller
 * action — exporting them is purely for display.
 */
export const BUILTIN_CJK_FILL_LISTS: ReadonlyArray<{
  id: string;
  members: readonly string[];
}> = [
  { id: "cjkWeekdaysShort", members: CJK_WEEKDAYS_SHORT },
  { id: "cjkWeekdayNumerals", members: CJK_WEEKDAY_NUMERALS },
  { id: "cjkWeekdays", members: CJK_WEEKDAYS },
  { id: "cjkMonthsArabic", members: CJK_MONTHS_ARABIC },
  { id: "cjkMonths", members: CJK_MONTHS },
  { id: "cjkQuarters", members: CJK_QUARTERS },
  { id: "lunarMonths", members: LUNAR_MONTHS },
  { id: "earthlyBranches", members: EARTHLY_BRANCHES },
  { id: "heavenlyStems", members: HEAVENLY_STEMS },
  { id: "cjkNumerals", members: CJK_NUMERALS },
];

function makeList(canonical: readonly string[]): TextSeriesList {
  return { lower: canonical.map((m) => m.toLowerCase()), canonical };
}

const BUILTIN_TEXT_SERIES_LISTS: readonly TextSeriesList[] = [
  makeList(WEEKDAYS_SHORT),
  makeList(WEEKDAYS_LONG),
  makeList(MONTHS_SHORT),
  makeList(MONTHS_LONG),
  ...BUILTIN_CJK_FILL_LISTS.map((list) => makeList(list.members)),
];

/**
 * Normalize a caller-supplied CUSTOM fill list (Excel's "Custom Lists" feature)
 * into a `TextSeriesList`. Trims each member, drops empties, and de-duplicates
 * (case-insensitively) so a malformed list can never create an ambiguous match
 * (two members that lower-case the same would make `indexOf` non-deterministic).
 * A list with fewer than two distinct members is dropped (returns null) — a
 * one-item "list" is just a constant and the constant-copy path already handles
 * it. The members are matched case-insensitively and echoed in the user's
 * stored casing, exactly like the built-in weekday/month lists.
 */
function normalizeCustomList(raw: readonly string[]): TextSeriesList | null {
  const canonical: string[] = [];
  const seenLower = new Set<string>();
  for (const item of raw) {
    if (typeof item !== "string") continue;
    const trimmed = item.trim();
    if (trimmed === "") continue;
    const lower = trimmed.toLowerCase();
    if (seenLower.has(lower)) continue;
    seenLower.add(lower);
    canonical.push(trimmed);
  }
  return canonical.length >= 2 ? makeList(canonical) : null;
}

/**
 * The full candidate list set for a fill: the built-ins plus any normalized
 * custom lists. User custom lists come FIRST so that if a token appears in both
 * a custom list and a built-in (e.g. a custom list that redefines weekday
 * order), the user's list wins — matching Excel, where a custom list takes
 * precedence over the built-in series for drag-fill.
 */
function resolveSeriesLists(
  customLists?: ReadonlyArray<readonly string[]>,
): readonly TextSeriesList[] {
  if (!customLists || customLists.length === 0) {
    return BUILTIN_TEXT_SERIES_LISTS;
  }
  const custom: TextSeriesList[] = [];
  for (const list of customLists) {
    const normalized = normalizeCustomList(list);
    if (normalized) custom.push(normalized);
  }
  return custom.length > 0
    ? [...custom, ...BUILTIN_TEXT_SERIES_LISTS]
    : BUILTIN_TEXT_SERIES_LISTS;
}

/** Casing of a recognised list member, so the extension echoes the input. */
type SeriesCasing = "upper" | "lower" | "title" | "as-is";

function detectCasing(raw: string, canonical: string): SeriesCasing {
  if (raw === canonical) return "as-is";
  if (raw === raw.toUpperCase() && raw !== raw.toLowerCase()) return "upper";
  if (raw === raw.toLowerCase()) return "lower";
  // Title-case (first letter upper, rest lower) is the canonical shape; anything
  // else (e.g. mixed) falls back to the canonical casing.
  return "title";
}

function applyCasing(
  canonical: string,
  casing: SeriesCasing,
): string {
  switch (casing) {
    case "upper":
      return canonical.toUpperCase();
    case "lower":
      return canonical.toLowerCase();
    case "title":
    case "as-is":
      return canonical;
  }
}

interface ListMatch {
  list: TextSeriesList;
  index: number;
  casing: SeriesCasing;
}

/**
 * Match a token against any of the candidate `lists` (case-insensitive), or
 * null. The lists are tried in order, so a custom list placed ahead of the
 * built-ins wins a tie.
 */
function matchList(
  raw: string,
  lists: readonly TextSeriesList[],
): ListMatch | null {
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  const lower = trimmed.toLowerCase();
  for (const list of lists) {
    const index = list.lower.indexOf(lower);
    if (index >= 0) {
      return {
        list,
        index,
        casing: detectCasing(trimmed, list.canonical[index]),
      };
    }
  }
  return null;
}

/** Render a list item at a (possibly out-of-range) ordinal, wrapping cyclically. */
function renderListAt(match: ListMatch, ordinal: number): string {
  const size = match.list.canonical.length;
  const wrapped = ((ordinal % size) + size) % size;
  return applyCasing(match.list.canonical[wrapped], match.casing);
}

// --- Quarters ----------------------------------------------------------------
// Microsoft's AutoFill table: "Qtr3 (or Q3 or Quarter3) -> Qtr4, Qtr1, Qtr2".
// A quarter is a four-member ring, so it wraps instead of counting on to Q5 the
// way an ordinary numbered label ("Item4" -> "Item5") does. The word may be in
// any case and may be separated from the digit by spaces; both are echoed
// verbatim. The ordinal form ("1st Quarter", "2nd Qtr") wraps the same way.
const QUARTER_RE = /^(q|qtr|quarter)(\s*)([1-4])$/i;
const ORDINAL_QUARTER_RE = /^([1-4])(st|nd|rd|th)(\s*)(quarter|qtr)$/i;

/** English ordinal suffix for a whole number: 1st 2nd 3rd 4th … 11th 21st. */
function ordinalSuffix(value: number): string {
  const n = Math.abs(Math.trunc(value));
  const lastTwo = n % 100;
  if (lastTwo >= 11 && lastTwo <= 13) return "th";
  switch (n % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}

interface QuarterLabel {
  /** 0-based quarter (Q1 = 0). */
  index: number;
  /** The label with the quarter number abstracted out, so two labels of the
   * same spelling compare equal ("Qtr #"). */
  shape: string;
  /** Spell another quarter (0-based) in this label's own shape. */
  render: (index: number) => string;
}

function matchQuarterLabel(raw: string): QuarterLabel | null {
  const text = raw.trim();
  const plain = QUARTER_RE.exec(text);
  if (plain) {
    const [, word, gap, digit] = plain;
    return {
      index: Number(digit) - 1,
      shape: `${word}${gap}#`,
      render: (index) => `${word}${gap}${index + 1}`,
    };
  }
  const ordinal = ORDINAL_QUARTER_RE.exec(text);
  if (ordinal) {
    const [, digit, suffix, gap, word] = ordinal;
    const quarter = Number(digit);
    if (suffix.toLowerCase() !== ordinalSuffix(quarter)) return null;
    const upper = suffix === suffix.toUpperCase();
    return {
      index: quarter - 1,
      shape: `#${upper ? "ST" : "st"}${gap}${word}`,
      render: (index) => {
        const next = ordinalSuffix(index + 1);
        return `${index + 1}${upper ? next.toUpperCase() : next}${gap}${word}`;
      },
    };
  }
  return null;
}

// A label with a number at its END ("Item1", "Week 3", "A-09", "项目1") or at
// its START ("1st Period", "3楼"). Microsoft's AutoFill table shows both:
// "Product 1 -> Product 2, Product 3" and "1st Period -> 2nd Period, 3rd
// Period". A number INSIDE the text ("第1季度") is neither and the text is
// copied: Microsoft documents no rule for it. Zero-padding width is preserved
// so "Item01" -> "Item02", and an English ordinal right after a leading number
// follows it (1st, 2nd, 3rd, 4th, … 11th, 21st).
const TRAILING_NUMBER_RE = /^(.*?)(\d+)$/;
const LEADING_NUMBER_RE = /^(\d+)(\D*)$/;
const LEADING_ORDINAL_RE = /^(st|nd|rd|th)(?![a-z])/i;
const HAS_LETTER_RE = /\p{L}/u;

interface NumberLabel {
  /** Text before the number ("" for a leading number). */
  prefix: string;
  /** Text after the number and its ordinal ("" for a trailing number). */
  suffix: string;
  /** The number's ordinal suffix is re-spelled per value, in this case. */
  ordinal: "lower" | "upper" | null;
  value: number;
  width: number;
  /** The label with the number abstracted out, to compare two labels. */
  shape: string;
}

function matchNumberLabel(raw: string): NumberLabel | null {
  // No trim: a label's surrounding spaces are part of its text. Reject a pure
  // number (handled by the numeric branch) and an empty digit run.
  const trailing = TRAILING_NUMBER_RE.exec(raw);
  if (trailing) {
    const prefix = trailing[1];
    const digits = trailing[2];
    // A pure-number string (empty prefix) is NOT a label — let the numeric
    // series own it so "1","2" still steps numerically and a lone "5" copies.
    // The prefix must contain a LETTER of any script, so a genuine text label
    // ("Item1", "Week 3", "A-09", "编号7") increments but a digits-and-
    
    // is left as text and copied — matching Excel, which never increments
    // those.
    if (prefix === "" || !HAS_LETTER_RE.test(prefix)) return null;
    const value = Number(digits);
    if (!Number.isFinite(value)) return null;
    return {
      prefix,
      suffix: "",
      ordinal: null,
      value,
      width: digits.length,
      shape: `${prefix}#`,
    };
  }
  // A leading number is the only number in the text, followed by a letter.
  const leading = LEADING_NUMBER_RE.exec(raw);
  if (!leading || !HAS_LETTER_RE.test(leading[2])) return null;
  const digits = leading[1];
  const value = Number(digits);
  if (!Number.isFinite(value)) return null;
  let suffix = leading[2];
  let ordinal: NumberLabel["ordinal"] = null;
  const ordinalMatch = LEADING_ORDINAL_RE.exec(suffix);
  if (ordinalMatch && ordinalMatch[1].toLowerCase() === ordinalSuffix(value)) {
    ordinal = ordinalMatch[1] === ordinalMatch[1].toUpperCase() ? "upper" : "lower";
    suffix = suffix.slice(2);
  }
  return {
    prefix: "",
    suffix,
    ordinal,
    value,
    width: digits.length,
    shape: `#${ordinal ? `(${ordinal})` : ""}${suffix}`,
  };
}

function renderNumberLabel(label: NumberLabel, value: number): string {
  const sign = value < 0 ? "-" : "";
  const digits = String(Math.abs(Math.trunc(value))).padStart(label.width, "0");
  const suffix = ordinalSuffix(value);
  const ordinal =
    label.ordinal === "upper"
      ? suffix.toUpperCase()
      : label.ordinal === "lower"
        ? suffix
        : "";
  return `${label.prefix}${sign}${digits}${ordinal}${label.suffix}`;
}

/**
 * Extend a TEXT series (built-in list or trailing-number label). Returns the
 * `count` next values, or null when the source block is not a recognisable text
 * series (so the caller falls back to cycling).
 *
 * Single source: step by ±1 in the drag direction. Two+ sources: the step is
 * the difference of the first two ordinals (so "Mon","Wed" -> Fri,Sun and
 * "Item1","Item3" -> Item5), required to be CONSISTENT across the whole block.
 */
function computeTextSeries(
  sourceCells: readonly string[],
  count: number,
  direction: FillDirection,
  lists: readonly TextSeriesList[],
): string[] | null {
  // Built-in / custom list (weekday / month / CJK / user list) — all sources
  // must hit the SAME list: the first one that holds every source, so
  // 十一月,腊月 reads on the lunar-month list although 十一月 alone would take
  // the 一月…十二月 list.
  const listMatches = sourceCells.map((cell) => matchList(cell, lists));
  if (listMatches.every((m) => m !== null)) {
    const common = lists.find((list) =>
      sourceCells.every((cell) =>
        list.lower.includes(cell.trim().toLowerCase()),
      ),
    );
    if (common) {
      const ms = sourceCells.map(
        (cell) => matchList(cell, [common]) as ListMatch,
      );
      const casing = ms[0].casing;
      const proto: ListMatch = { list: ms[0].list, index: 0, casing };
      if (ms.length === 1) {
        const step = directionSign(direction);
        const start = ms[0].index;
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(renderListAt(proto, start + step * i));
        }
        return out;
      }
      // Multi-cell: ordinal step from the first pair, must be consistent. The
      // ordinals are taken on the list's ring, using the minimal signed delta
      // between consecutive members so "Sat","Sun" reads as +1 (not -6).
      const step = ringStep(ms[0].index, ms[1].index, ms[0].list.lower.length);
      let consistent = true;
      for (let i = 2; i < ms.length; i += 1) {
        if (ringStep(ms[i - 1].index, ms[i].index, ms[0].list.lower.length) !== step) {
          consistent = false;
          break;
        }
      }
      if (consistent) {
        const last = ms[ms.length - 1].index;
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(renderListAt(proto, last + step * i));
        }
        return out;
      }
    }
    return null;
  }

  // Quarters ("Q1", "Qtr 3", "quarter2", "1st Quarter") — a ring of four, all
  // sources spelled the same way.
  const quarters = sourceCells.map(matchQuarterLabel);
  if (quarters.every((q) => q !== null)) {
    const qs = quarters as QuarterLabel[];
    if (qs.every((q) => q.shape === qs[0].shape)) {
      const render = qs[0].render;
      const wrap = (ordinal: number) => render(((ordinal % 4) + 4) % 4);
      if (qs.length === 1) {
        const step = directionSign(direction);
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(wrap(qs[0].index + step * i));
        }
        return out;
      }
      const step = ringStep(qs[0].index, qs[1].index, 4);
      const consistent = qs.every(
        (q, i) => i < 2 || ringStep(qs[i - 1].index, q.index, 4) === step,
      );
      if (consistent) {
        const last = qs[qs.length - 1].index;
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(wrap(last + step * i));
        }
        return out;
      }
    }
    return null;
  }

  // Numbered labels ("Item1", "Week 3", "Q5", "1st Period") — all sources must
  // share the SAME text around the number and digit width so the series is
  // well-defined.
  const labels = sourceCells.map(matchNumberLabel);
  if (labels.every((l) => l !== null)) {
    const ls = labels as NumberLabel[];
    const sameShape = ls.every((l) => l.shape === ls[0].shape);
    const sameWidth = ls.every((l) => l.width === ls[0].width);
    if (sameShape && sameWidth) {
      if (ls.length === 1) {
        const step = directionSign(direction);
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(renderNumberLabel(ls[0], ls[0].value + step * i));
        }
        return out;
      }
      const step = ls[1].value - ls[0].value;
      let consistent = true;
      for (let i = 2; i < ls.length; i += 1) {
        if (ls[i].value - ls[i - 1].value !== step) {
          consistent = false;
          break;
        }
      }
      if (consistent) {
        const last = ls[ls.length - 1].value;
        const out: string[] = [];
        for (let i = 1; i <= count; i += 1) {
          out.push(renderNumberLabel(ls[0], last + step * i));
        }
        return out;
      }
    }
    return null;
  }

  return null;
}

/**
 * Minimal signed step from index `a` to `b` on a ring of `size`, choosing the
 * shorter direction so two ADJACENT list members read as ±1 even across the
 * wrap seam ("Sat"(6) -> "Sun"(0) is +1, not -6). Ties (exactly half the ring)
 * resolve to the positive direction.
 */
function ringStep(a: number, b: number, size: number): number {
  let delta = (((b - a) % size) + size) % size; // 0..size-1
  if (delta > size / 2) delta -= size; // choose the shorter signed step
  return delta;
}

// --- Text dates ----------------------------------------------------------------
// A date typed into a cell is stored as the text the user typed, and the formula
// engine reads that text as a date wherever it needs a number (toNumber falls
// back to DATEVALUE in formula-engine.ts). The fill recognises the same forms
// DATEVALUE does, so a value that computes as a date also fills as one:

//   m/d/yyyy / m-d-yyyy        (9/21/2026)
//   d Mon yyyy / d-Mon-yyyy    (21-Sep-2026, 21 September 2026)
//   Mon d, yyyy / Mon d yyyy   (Sep 21, 2026, September 21 2026)
// Each filled value is written back in its seed's own shape: the separators,
// the month name's length and case, and whether month/day carry a leading zero.
const YMD_DATE_RE = /^(\d{4})([-/])(\d{1,2})([-/])(\d{1,2})$/;
const MDY_DATE_RE = /^(\d{1,2})([-/])(\d{1,2})([-/])(\d{4})$/;
const DAY_MONTH_NAME_DATE_RE = /^(\d{1,2})([\s-]+)([A-Za-z]{3,9})([\s-]+)(\d{4})$/;
const MONTH_NAME_DAY_DATE_RE = /^([A-Za-z]{3,9})([\s-]+)(\d{1,2})(,?\s+)(\d{4})$/;
const MONTH_NAME_LISTS: readonly TextSeriesList[] = [
  makeList(MONTHS_SHORT),
  makeList(MONTHS_LONG),
];

interface TextDate {
  /** The date as a UTC-midnight epoch-ms. */
  ms: number;
  /** Spell another date (UTC-midnight epoch-ms) in this seed's shape. */
  render: (ms: number) => string;
}

 
function calendarDateMs(year: number, month: number, day: number): number | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const ms = Date.UTC(year, month - 1, day);
  const back = new Date(ms);
  return back.getUTCFullYear() === year &&
    back.getUTCMonth() + 1 === month &&
    back.getUTCDate() === day
    ? ms
    : null;
}

/** Whether a written month/day field shows zero padding: "09" yes, "9" no,
 * "21" cannot tell. */
function fieldPadding(field: string): boolean | null {
  if (field.length === 1) return false;
  return field.startsWith("0") ? true : null;
}

function padField(value: number, padded: boolean): string {
  return padded ? String(value).padStart(2, "0") : String(value);
}

function dateParts(ms: number): { year: string; month: number; day: number } {
  const date = new Date(ms);
  return {
    year: String(date.getUTCFullYear()).padStart(4, "0"),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}

/** The pieces of a date written with a month name, day first or month first. */
function namedMonthDateFields(text: string): {
  dayFirst: boolean;
  dayText: string;
  monthName: string;
  yearText: string;
  gap1: string;
  gap2: string;
} | null {
  const dayFirst = DAY_MONTH_NAME_DATE_RE.exec(text);
  if (dayFirst) {
    const [, dayText, gap1, monthName, gap2, yearText] = dayFirst;
    return { dayFirst: true, dayText, monthName, yearText, gap1, gap2 };
  }
  const monthFirst = MONTH_NAME_DAY_DATE_RE.exec(text);
  if (monthFirst) {
    const [, monthName, gap1, dayText, gap2, yearText] = monthFirst;
    return { dayFirst: false, dayText, monthName, yearText, gap1, gap2 };
  }
  return null;
}

function parseTextDate(raw: string): TextDate | null {
  const text = raw.trim();
  const numeric = YMD_DATE_RE.exec(text) ?? MDY_DATE_RE.exec(text);
  if (numeric) {
    const yearFirst = numeric[1].length === 4;
    const [yearText, monthText, dayText] = yearFirst
      ? [numeric[1], numeric[3], numeric[5]]
      : [numeric[5], numeric[1], numeric[3]];
    const ms = calendarDateMs(
      Number(yearText),
      Number(monthText),
      Number(dayText),
    );
    if (ms === null) return null;
    const sep1 = numeric[2];
    const sep2 = numeric[4];
    // A field that does not show its padding follows the other field, then
    // the form's own habit: ISO yyyy-mm-dd pads, slashed and US dates do not.
    const fallback = yearFirst && sep1 === "-" && sep2 === "-";
    const padMonth =
      fieldPadding(monthText) ?? fieldPadding(dayText) ?? fallback;
    const padDay = fieldPadding(dayText) ?? fieldPadding(monthText) ?? fallback;
    return {
      ms,
      render: (next) => {
        const { year, month, day } = dateParts(next);
        const m = padField(month, padMonth);
        const d = padField(day, padDay);
        return yearFirst
          ? `${year}${sep1}${m}${sep2}${d}`
          : `${m}${sep1}${d}${sep2}${year}`;
      },
    };
  }
  const named = namedMonthDateFields(text);
  if (!named) return null;
  const { dayFirst, dayText, monthName, yearText, gap1, gap2 } = named;
  const monthMatch = matchList(monthName, MONTH_NAME_LISTS);
  if (!monthMatch) return null;
  const ms = calendarDateMs(
    Number(yearText),
    monthMatch.index + 1,
    Number(dayText),
  );
  if (ms === null) return null;
  const padDay = fieldPadding(dayText) ?? false;
  return {
    ms,
    render: (next) => {
      const { year, month, day } = dateParts(next);
      const name = renderListAt(monthMatch, month - 1);
      const d = padField(day, padDay);
      return dayFirst
        ? `${d}${gap1}${name}${gap2}${year}`
        : `${name}${gap1}${d}${gap2}${year}`;
    },
  };
}

/** Months since year 0 of a UTC-midnight epoch-ms, to measure month steps. */
function monthOrdinal(ms: number): number {
  const date = new Date(ms);
  return date.getUTCFullYear() * 12 + date.getUTCMonth();
}

/**
 * Compute the values a fill produces for `count` target cells, given the source
 * block (ordered from the far edge toward the fill) and the drag direction.
 *
 * `customLists` (optional) are the user's saved Custom Lists (Excel's
 * File > Options > Custom Lists) plus they take precedence over the built-in
 * weekday / month / CJK lists for the text-series branch. Each is a series of
 * tokens (e.g. ["Spring","Summer","Autumn","Winter"] or a CJK list); a list
 * with fewer than two distinct members is ignored.
 *
 * Always returns exactly `max(0, count)` strings. Safe on non-numeric / mixed /
 * empty input (falls back to cycling the source values).
 *
 * `mode` is the Auto Fill Options choice (see AutoFillValueMode); a plain
 * drag is "auto". `numberFormats` are the source cells' number formats: a
 * number under a date format is a date (a typed date is stored as its
 * serial), and one under a time format steps by the hour.
 */
export function computeFillValues(
  sourceCells: readonly string[],
  count: number,
  direction: FillDirection,
  customLists?: ReadonlyArray<readonly string[]>,
  mode: AutoFillValueMode = "auto",
  numberFormats?: FillNumberFormats,
  dateSystem?: SpreadsheetDateSystemOptions,
): string[] {
  const target = Math.max(0, Math.floor(count));
  if (target === 0) return [];
  if (sourceCells.length === 0) {
    return Array.from({ length: target }, () => "");
  }
  if (mode === "copy") return cycleValues(sourceCells, target);
  if (mode !== "auto" && mode !== "series") {
    const dated = fillDatesBy(
      sourceCells,
      target,
      direction,
      mode,
      numberFormats,
      dateSystem,
    );
    if (dated) return dated;
  }
  const serialDates = serialDateSeeds(sourceCells, numberFormats, dateSystem);
  // Date-times a part of a day apart (9:00 and 10:00 of one day) step by
  // their difference as numbers do -- 11:00, 12:00 -- as Excel fills them;
  // only seeds at one time of day continue by days, months or years.
  if (serialDates && !seedTimesOfDayDiffer(sourceCells)) {
    return extendDateSeeds(serialDates, sourceCells, target, direction);
  }

  // --- Numeric series -------------------------------------------------------
  const numbers = sourceCells.map(parseNumber);
  if (numbers.every((n) => n !== null)) {
    const nums = numbers as number[];
    const padWidth = zeroPaddedWidth(sourceCells);
    const render = (n: number) =>
      padWidth !== null && Number.isInteger(n) && n >= 0
        ? String(n).padStart(padWidth, "0")
        : formatNumber(n);
    if (nums.length === 1) {
      // One time steps an hour (Microsoft Support's AutoFill examples:
      // 9:00 -> 10:00, 11:00, 12:00).
      if (numberFormatFamily(numberFormats?.[0]) === "time") {
        const hour = directionSign(direction) / 24;
        return Array.from({ length: target }, (_, i) =>
          formatSerial(nums[0] + hour * (i + 1)),
        );
      }
      // Single numeric source: Excel copies the constant (no implied step) on a
      // plain drag. Fill Series (or a Ctrl+drag) steps it by 1 in the drag
      // direction instead.
      if (mode === "series") {
        const step = directionSign(direction);
        return Array.from({ length: target }, (_, i) =>
          render(nums[0] + step * (i + 1)),
        );
      }
      return Array.from({ length: target }, () => sourceCells[0]);
    }
    // Detect a consistent arithmetic step across the source block.
    const step = nums[1] - nums[0];
    let arithmetic = true;
    for (let i = 2; i < nums.length; i += 1) {
      if (!nearlyEqual(nums[i] - nums[i - 1], step)) {
        arithmetic = false;
        break;
      }
    }
    if (arithmetic) {
      const last = nums[nums.length - 1];
      const out: string[] = [];
      for (let i = 1; i <= target; i += 1) {
        out.push(render(last + step * i));
      }
      return out;
    }
    // Any other run of numbers continues along its best-fit line (least
    // squares, as TREND): 3, 5, 8 dragged on fill 10.33..., 12.83... and the
    // source keeps its values (Microsoft Support, "Project values in a
    // series"; Excel 16 measured). Copy Cells still repeats the block.
    const meanX = (nums.length + 1) / 2;
    const meanY = nums.reduce((sum, value) => sum + value, 0) / nums.length;
    let covariance = 0;
    let variance = 0;
    nums.forEach((value, index) => {
      covariance += (index + 1 - meanX) * (value - meanY);
      variance += (index + 1 - meanX) ** 2;
    });
    const slope = covariance / variance;
    return Array.from({ length: target }, (_, i) =>
      render(meanY + slope * (nums.length + i + 1 - meanX)),
    );
  }

  // --- Date series ----------------------------------------------------------
  // Dates typed as text in older pages ("2026/9/21"). The i-th filled cell
  // takes its format from source cell i mod n (the caller copies formatting
  // that way), so it also takes that seed's text shape.
  const dates = sourceCells.map(parseTextDate);
  if (dates.every((d) => d !== null)) {
    return extendDateSeeds(dates as TextDate[], sourceCells, target, direction);
  }

  // --- Text series (weekday / month / CJK / custom / Q1 / Item1 …) ----------
  // Tried only after numeric + date, so a numeric block never reaches here. A
  // null result means the block is not a recognisable text series → cycle.
  const textSeries = computeTextSeries(
    sourceCells,
    target,
    direction,
    resolveSeriesLists(customLists),
  );
  if (textSeries) return textSeries;

  // --- Numbered labels among other text -------------------------------------
  const mixed = computeMixedLabelSeries(
    sourceCells,
    target,
    direction,
    resolveSeriesLists(customLists),
  );
  if (mixed) return mixed;

  // --- Constant / mixed -----------------------------------------------------
  return cycleValues(sourceCells, target);
}

/**
 * A block that is no series as a whole but holds numbered labels. Microsoft's
 * AutoFill table ("Fill data automatically in worksheet cells", the Initial
 * selection / Extended series examples) has the row
 * "text1, textA -> text2, textA, text3, textA, ...": the numbered label counts
 * on from one repeat of the block to the next while the text beside it is
 * repeated. Neighbouring labels spelled the same way ("text1", "text2" beside
 * "textA") are one series between them, extended as they would be alone.
 * Everything else in the block (plain text, numbers, dates) is repeated, as
 * before. Null when the block holds no numbered label.
 */
function computeMixedLabelSeries(
  sourceCells: readonly string[],
  count: number,
  direction: FillDirection,
  lists: readonly TextSeriesList[],
): string[] | null {
  const size = sourceCells.length;
  if (size < 2) return null;
  const labels = sourceCells.map(matchNumberLabel);
  if (labels.every((label) => label === null)) return null;
  const repeats = Math.ceil(count / size);
  // byPosition[p][r]: the value at source position p in the r-th repeat.
  const byPosition: string[][] = [];
  let start = 0;
  while (start < size) {
    const label = labels[start];
    let end = start + 1;
    if (label) {
      while (
        end < size &&
        labels[end]?.shape === label.shape &&
        labels[end]?.width === label.width
      ) {
        end += 1;
      }
    }
    const run = sourceCells.slice(start, end);
    const length = run.length * repeats;
    const series =
      (label ? computeTextSeries(run, length, direction, lists) : null) ??
      cycleValues(run, length);
    for (let offset = 0; offset < run.length; offset += 1) {
      byPosition[start + offset] = Array.from(
        { length: repeats },
        (_, repeat) => series[repeat * run.length + offset],
      );
    }
    start = end;
  }
  return Array.from(
    { length: count },
    (_, index) => byPosition[index % size][Math.floor(index / size)],
  );
}

/**
 * The Auto Fill Options button Excel shows after a fill-handle drag
 * (Microsoft Support, "Fill data automatically in worksheet cells") redoes the
 * fill another way. For the values of a line that is not all formulas:
 * - "auto": what a plain drag does (a series where the source is one, else a
 *   copy);
 * - "copy": Copy Cells, the source repeated;
 * - "series": Fill Series, which also steps a single number by 1;
 * - a date unit: Fill Days / Fill Weekdays / Fill Months / Fill Years, for a
 *   source of dates only (any other source fills as "auto").
 * Holding Ctrl while dragging swaps the plain drag's copy and series.
 */
export type AutoFillValueMode = "auto" | "copy" | "series" | SeriesDateUnit;

/**
 * What the Auto Fill Options button needs to know about one line: whether a
 * plain drag repeats it (Copy Cells, which Ctrl+drag turns into Fill Series,
 * and the other way round) and whether Fill Series would extend it at all
 * (Excel leaves Fill Series out for text that is no series).
 */
export function autoFillLineOptions(
  sourceCells: readonly string[],
  count: number,
  direction: FillDirection,
  customLists?: ReadonlyArray<readonly string[]>,
  numberFormats?: FillNumberFormats,
  dateSystem?: SpreadsheetDateSystemOptions,
): { copies: boolean; extendable: boolean } {
  const copy = cycleValues(sourceCells, Math.max(0, Math.floor(count)));
  const same = (values: readonly string[]) =>
    values.every((value, index) => value === copy[index]);
  const fill = (mode: AutoFillValueMode) =>
    computeFillValues(
      sourceCells,
      count,
      direction,
      customLists,
      mode,
      numberFormats,
      dateSystem,
    );
  return { copies: same(fill("auto")), extendable: !same(fill("series")) };
}

/** Whether every cell of the line is a date the fill recognises, which is
 * when Excel adds Fill Days / Weekdays / Months / Years to its options. */
export function autoFillLineIsDates(
  sourceCells: readonly string[],
  numberFormats?: FillNumberFormats,
  dateSystem?: SpreadsheetDateSystemOptions,
): boolean {
  return (
    sourceCells.length > 0 &&
    (serialDateSeeds(sourceCells, numberFormats, dateSystem) !== null ||
      sourceCells.every((value) => parseTextDate(value) !== null))
  );
}

/** The source cells' number formats, parallel to their values. */
export type FillNumberFormats = readonly (string | null | undefined)[];

/**
 * Date seeds for cells that hold a serial under a date number format, which
 * is how a typed date is stored. Each filled value is a serial again and
 * keeps its seed's time of day. Null unless every cell is one.
 */
function serialDateSeeds(
  sourceCells: readonly string[],
  numberFormats: FillNumberFormats | undefined,
  dateSystem?: SpreadsheetDateSystemOptions,
): TextDate[] | null {
  if (!numberFormats || sourceCells.length === 0) return null;
  const seeds: TextDate[] = [];
  for (let i = 0; i < sourceCells.length; i += 1) {
    if (numberFormatFamily(numberFormats[i]) !== "date") return null;
    const seed = serialDateSeed(sourceCells[i], dateSystem);
    if (!seed) return null;
    seeds.push(seed);
  }
  return seeds;
}

/** Whether serial seeds carry different times of day (to the millisecond). */
function seedTimesOfDayDiffer(sourceCells: readonly string[]): boolean {
  const times = sourceCells.map((raw) => {
    const serial = parseNumber(raw) ?? 0;
    return Math.round((serial - Math.floor(serial)) * 86_400_000);
  });
  return times.some((time) => time !== times[0]);
}

/**
 * A single serial-number seed read as a date in the workbook's system: any
 * number >= 1 in the 1900 system (serial 1 is 1900-01-01; the fictitious
 * 1900-02-29, 60, steps from 1900-02-28), any number >= 0 in 1904 (serial 0
 * is 1904-01-01). The filled cells are serials of the same system and keep
 * the seed's time of day.
 */
function serialDateSeed(
  raw: string,
  dateSystem?: SpreadsheetDateSystemOptions,
): TextDate | null {
  const date1904 = dateSystem?.date1904;
  const serial = parseNumber(raw);
  if (serial === null || !Number.isFinite(serial) || serial < excelMinDateSerial(date1904)) {
    return null;
  }
  const whole = Math.floor(serial);
  const time = serial - whole;
  return {
    ms: excelSystemSerialToUtcMs(whole, date1904),
    render: (ms) =>
      formatSerial(Math.round(excelSystemSerialFromUtcMs(ms, date1904)) + time),
  };
}

/** A date or time serial without float noise; 15 significant digits keep
 * a time of day exact. */
function formatSerial(value: number): string {
  return String(Number(value.toPrecision(15)));
}

/**
 * Fill Days / Weekdays / Months / Years: dates stepped by the unit. One seed
 * steps by 1 in the drag direction; several step by the difference of the
 * first two in that unit (months and years counted from the last seed, so a
 * month end stays one). Null when the line is not all dates.
 */
function fillDatesBy(
  sourceCells: readonly string[],
  count: number,
  direction: FillDirection,
  unit: SeriesDateUnit,
  numberFormats?: FillNumberFormats,
  dateSystem?: SpreadsheetDateSystemOptions,
): string[] | null {
  const parsed =
    serialDateSeeds(sourceCells, numberFormats, dateSystem) ??
    sourceCells.map(parseTextDate);
  if (!parsed.every((d) => d !== null)) return null;
  const ds = parsed as TextDate[];
  const first = ds[0].ms;
  const last = ds[ds.length - 1].ms;
  let step: number = directionSign(direction);
  if (ds.length > 1) {
    const second = ds[1].ms;
    const measured =
      unit === "month"
        ? monthOrdinal(second) - monthOrdinal(first)
        : unit === "year"
          ? new Date(second).getUTCFullYear() - new Date(first).getUTCFullYear()
          : unit === "day"
            ? Math.round((second - first) / MS_PER_DAY)
            : Math.sign(second - first);
    if (measured !== 0) step = measured;
  }
  const out: string[] = [];
  let current = last;
  for (let i = 1; i <= count; i += 1) {
    current =
      unit === "month" || unit === "year"
        ? advanceDate(last, unit, step * i)
        : advanceDate(current, unit, step);
    out.push(ds[(i - 1) % ds.length].render(current));
  }
  return out;
}

/**
 * A plain drag of dates (Microsoft Support, "Fill data automatically in
 * worksheet cells"): one date steps a day in the drag direction; dates on the
 * same day of the month that step by whole months continue by months (and so
 * the same month and day a year apart continue by years); any other even step
 * continues by that many days; anything else is repeated.
 */
function extendDateSeeds(
  ds: readonly TextDate[],
  sourceCells: readonly string[],
  target: number,
  direction: FillDirection,
): string[] {
  const renderAt = (i: number, ms: number) =>
    ds[(i - 1) % ds.length].render(ms);
  if (ds.length === 1) {
    const stepDays = directionSign(direction);
    const out: string[] = [];
    for (let i = 1; i <= target; i += 1) {
      out.push(renderAt(i, ds[0].ms + stepDays * i * MS_PER_DAY));
    }
    return out;
  }
  const last = ds[ds.length - 1].ms;
  // Dates on the same day of the month that step by whole months continue
  // by months (Jan 15, Apr 15 -> Jul 15, Oct 15), each counted from the last
  // seed so a month end never drifts (Jan 31, Mar 31 -> May 31, Jul 31,
  // Sep 30, Nov 30).
  const dayOfMonth = new Date(ds[0].ms).getUTCDate();
  const stepMonths = monthOrdinal(ds[1].ms) - monthOrdinal(ds[0].ms);
  const byMonths =
    stepMonths !== 0 &&
    ds.every(
      (d, i) =>
        new Date(d.ms).getUTCDate() === dayOfMonth &&
        (i === 0 ||
          monthOrdinal(d.ms) - monthOrdinal(ds[i - 1].ms) === stepMonths),
    );
  if (byMonths) {
    const out: string[] = [];
    for (let i = 1; i <= target; i += 1) {
      out.push(renderAt(i, advanceDate(last, "month", stepMonths * i)));
    }
    return out;
  }
  const stepMs = ds[1].ms - ds[0].ms;
  const arithmetic = ds.every(
    (d, i) => i === 0 || d.ms - ds[i - 1].ms === stepMs,
  );
  if (arithmetic) {
    const out: string[] = [];
    for (let i = 1; i <= target; i += 1) {
      out.push(renderAt(i, last + stepMs * i));
    }
    return out;
  }
  return cycleValues(sourceCells, target);
}

/**
 * Repeat the source values to fill `count` slots (Excel's behaviour for a
 * non-series block: AAB -> AABAAB…). For a single source this is a plain copy.
 */
function cycleValues(sourceCells: readonly string[], count: number): string[] {
  const out: string[] = [];
  for (let i = 0; i < count; i += 1) {
    out.push(sourceCells[i % sourceCells.length]);
  }
  return out;
}

/**
 * Render a computed numeric series value back to a string without trailing
 * float noise (e.g. 0.1+0.2 patterns) while keeping integers integer. Rounds to
 * 15 SIGNIFICANT digits (what Excel keeps), so a tiny series (1e-11, 2e-11,
 * 3e-11) stays itself instead of rounding to 0 at a fixed decimal place.
 */
function formatNumber(value: number): string {
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toPrecision(15)));
}

/**
 * '001, '002: digit-only text of one width with a leading zero is a numbered
 * code, and its series keeps the width (003, 004), as Excel fills it. The
 * width, or null when the block is ordinary numbers.
 */
function zeroPaddedWidth(cells: readonly string[]): number | null {
  const width = cells[0]?.length ?? 0;
  if (width < 2) return null;
  if (!cells.every((cell) => cell.length === width && /^\d+$/.test(cell))) return null;
  return cells.some((cell) => cell.startsWith("0")) ? width : null;
}

// --- Explicit Fill > Series dialog (Excel "Home > Fill > Series…") -----------
// Unlike the drag-fill auto-detection above, the Series DIALOG takes an explicit
// step value, an optional stop value, and a series TYPE, and generates values
// until either `count` cells are filled or the stop value is passed. This is the
// pure math behind that dialog; the component owns the model write (it pastes the
// returned strings through the same paste path so undo + autosave fire).

/** The four series types the dialog offers (Excel parity). */
export type SeriesType = "linear" | "growth" | "date" | "autoFill";

/** The date unit for a `date` series (Excel's Date Unit radio group). */
export type SeriesDateUnit = "day" | "weekday" | "month" | "year";

export interface SeriesOptions {
  /** Series type: arithmetic (linear), geometric (growth), or calendar (date). */
  type: SeriesType;
  /** Step / increment. Linear: added each step. Growth: multiplied each step.
   * Date: the number of date-units added each step. Defaults to 1 when omitted. */
  step?: number;
  /** Optional stop value: generation halts once the next value would PASS it
   * (above it for an increasing series, below for a decreasing one). For a date
   * series the stop is a date in any form the fill recognises (compared
   * chronologically). */
  stop?: number | string | null;
  /** Date unit for a `date` series; ignored otherwise. Defaults to "day". */
  dateUnit?: SeriesDateUnit;
  /** The workbook uses the 1904 date system: a serial seed or stop is read,
   * and the filled serials written, in that system. */
  date1904?: ExcelDateSystemArg;
}

/** Advance a UTC-midnight epoch-ms by `n` calendar units of the given kind. */
function advanceDate(ms: number, unit: SeriesDateUnit, n: number): number {
  const d = new Date(ms);
  switch (unit) {
    case "day":
      return ms + n * MS_PER_DAY;
    case "weekday": {
      // Step `n` business days (skip Sat/Sun). Direction follows the sign of n.
      if (n === 0) return ms;
      const dir = n > 0 ? 1 : -1;
      let remaining = Math.abs(Math.trunc(n));
      let cur = ms;
      while (remaining > 0) {
        cur += dir * MS_PER_DAY;
        const dow = new Date(cur).getUTCDay(); // 0 Sun … 6 Sat
        if (dow !== 0 && dow !== 6) remaining -= 1;
      }
      return cur;
    }
    case "month": {
      const year = d.getUTCFullYear();
      const month = d.getUTCMonth();
      const day = d.getUTCDate();
      const target = new Date(Date.UTC(year, month + n, 1));
      // Clamp the day to the target month's length (Excel: Jan 31 +1mo -> Feb 28).
      const lastDay = new Date(
        Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0),
      ).getUTCDate();
      return Date.UTC(
        target.getUTCFullYear(),
        target.getUTCMonth(),
        Math.min(day, lastDay),
      );
    }
    case "year": {
      const year = d.getUTCFullYear();
      const month = d.getUTCMonth();
      const day = d.getUTCDate();
      const target = new Date(Date.UTC(year + n, month, 1));
      const lastDay = new Date(
        Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0),
      ).getUTCDate();
      return Date.UTC(
        target.getUTCFullYear(),
        target.getUTCMonth(),
        Math.min(day, lastDay),
      );
    }
  }
}

/**
 * Generate an explicit Fill > Series from a START value and dialog options,
 * producing AT MOST `maxCount` values (the size of the target range) and halting
 * early when a `stop` value is set and the next value would pass it. The START
 * cell itself is NOT included — these are the values written into the cells AFTER
 * it (matching how the dialog fills the rest of a selection whose first cell is
 * the seed). Always returns at most `max(0, maxCount)` strings.
 *
 *  - linear: start + step, start + 2·step, …
 *  - growth: start · step, start · step², … (a step of 0 or 1 degenerates; we
 *            still emit values so the dialog never silently produces nothing).
 *  - date:   start advanced by `step` date-units each step (day/weekday/month/
 *            year), written in the start's own text shape (2026/9/21 ->
 *            2026/9/22). A start that is not a date the fill recognises falls
 *            back to a linear numeric series on its number value, or an empty
 *            result when it is not numeric either.
 *
 * `autoFill` is not handled here (it routes to computeFillValues); passing it
 * yields an empty array so the caller falls back explicitly.
 */
export function computeSeriesValues(
  start: string,
  options: SeriesOptions,
  maxCount: number,
): string[] {
  const max = Math.max(0, Math.floor(maxCount));
  if (max === 0) return [];
  const step = Number.isFinite(Number(options.step)) ? Number(options.step) : 1;

  if (options.type === "autoFill") return [];

  if (options.type === "date") {
    // A date entered normally is stored as its serial (2026/1/31 is 46053):
    // in a Date series it is that date, as Excel reads any number there, and
    // the filled cells keep the serial form for the seed's date format (live
    // audit B153).
    const dateSystem = { date1904: options.date1904 };
    const startDate = parseTextDate(start) ?? serialDateSeed(start, dateSystem);
    if (startDate !== null) {
      const unit = options.dateUnit ?? "day";
      const stopText =
        typeof options.stop === "string" && options.stop.trim() !== ""
          ? options.stop
          : typeof options.stop === "number"
            ? String(options.stop)
            : null;
      const stopMs =
        stopText !== null
          ? ((parseTextDate(stopText) ?? serialDateSeed(stopText, dateSystem))?.ms ??
            null)
          : null;
      const out: string[] = [];
      let cur = startDate.ms;
      const stepN = Math.trunc(step);
      // Months and years count from the start, so a month end stays one:
      // Microsoft Learn's Range.DataSeries example fills 31-JAN-1996 by months
      // into "the last day of every month in 1996" (Feb 29, Mar 31, Apr 30).
      const fromStart = unit === "month" || unit === "year";
      for (let i = 0; i < max; i += 1) {
        cur = fromStart
          ? advanceDate(startDate.ms, unit, stepN * (i + 1))
          : advanceDate(cur, unit, stepN);
        if (stopMs !== null) {
          if (stepN >= 0 && cur > stopMs) break;
          if (stepN < 0 && cur < stopMs) break;
        }
        out.push(startDate.render(cur));
      }
      return out;
    }
    // Non-date start: fall through to a numeric linear series if it parses.
  }

  const startNum = parseNumber(start);
  if (startNum === null) return [];

  const stopNum =
    options.stop === null || options.stop === undefined || options.stop === ""
      ? null
      : Number(options.stop);
  const hasStop = stopNum !== null && Number.isFinite(stopNum);

  const out: string[] = [];
  let value = startNum;
  for (let i = 0; i < max; i += 1) {
    const previous = value;
    if (options.type === "growth") {
      value = value * step;
    } else {
      value = value + step;
    }
    if (hasStop) {
      // Halt once the series PASSES the stop, in the direction it moves: a
      // linear series by the sign of its step, a growth series by where this
      // term lies against the previous one (-1, -2, -4, -8 is decreasing,
      // whatever its magnitudes do).
      const increasing =
        options.type === "growth" ? value > previous : step >= 0;
      if (increasing && value > (stopNum as number)) break;
      if (!increasing && value < (stopNum as number)) break;
    }
    out.push(formatNumber(value));
  }
  return out;
}

// --- Home > Fill > Series over a selection -------------------------------------
// Excel's Series dialog (Microsoft Support "Project values in a series" and
// "Fill data automatically in worksheet cells"; Range.DataSeries on Microsoft
// Learn) fills every row or every column of the selection on its own, from
// that line's first cell. "Series in" picks rows or columns; left alone,
// DataSeries takes it from "the size and shape of the range" — a selection
// taller than it is wide runs down its columns, any other runs across rows.
// With Trend checked, a Linear or Growth line is refitted from its starting
// values (least squares y = mx + b, or y = b·m^x — what TREND and GROWTH
// return), the step is ignored, and "the resulting series replaces the
// original selected values". With only one cell selected and a stop value,
// the series runs on from that cell until it passes the stop.

export type SeriesIn = "rows" | "columns";

/** "Series in" for a selection of this size: columns when it is taller than
 * it is wide, rows otherwise. */
export function defaultSeriesIn(rowCount: number, columnCount: number): SeriesIn {
  return rowCount > columnCount ? "columns" : "rows";
}

export interface FillSeriesRange {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

export interface FillSeriesRequest extends SeriesOptions {
  /** The selection (0-based, inclusive). */
  range: FillSeriesRange;
  seriesIn: SeriesIn;
  /** Fit Linear/Growth to the starting values instead of using the step. */
  trend?: boolean;
  /** Row / column counts of the grid, bounding a one-cell selection that
   * runs on toward its stop value. */
  rowLimit: number;
  columnLimit: number;
}

export interface FillSeriesWrite {
  row: number;
  col: number;
  value: string;
}

/** Least-squares fit through (1, y1), (2, y2), …: y = slope·x + intercept. */
function leastSquares(ys: readonly number[]): { slope: number; intercept: number } {
  const n = ys.length;
  let sx = 0;
  let sy = 0;
  let sxy = 0;
  let sxx = 0;
  ys.forEach((y, i) => {
    const x = i + 1;
    sx += x;
    sy += y;
    sxy += x * y;
    sxx += x * x;
  });
  const slope = (n * sxy - sx * sy) / (n * sxx - sx * sx);
  return { slope, intercept: (sy - slope * sx) / n };
}

/** The best-fit value at x (1-based) through the starting values, or null
 * when Growth cannot fit them (GROWTH needs every value above zero). */
function trendFit(
  type: "linear" | "growth",
  seeds: readonly number[],
): ((x: number) => number) | null {
  if (type === "linear") {
    const { slope, intercept } = leastSquares(seeds);
    return (x) => intercept + slope * x;
  }
  if (seeds.some((y) => y <= 0)) return null;
  const { slope, intercept } = leastSquares(seeds.map(Math.log));
  return (x) => Math.exp(intercept + slope * x);
}

/**
 * One line of the Series dialog: the values written from `start` on, or null
 * when the line cannot seed a series. `read(i)` is the i-th cell's value (a
 * formula's result); `length` is how many cells the line may cover.
 */
function planSeriesLine(
  read: (index: number) => string,
  length: number,
  request: FillSeriesRequest,
  runsToStop: boolean,
): { start: number; values: string[] } | null {
  const type = request.type;
  if (type === "autoFill") return null;
  const stop =
    request.stop === null || request.stop === undefined
      ? ""
      : String(request.stop).trim();
  if ((type === "linear" || type === "growth") && request.trend) {
    const seeds: number[] = [];
    for (let i = 0; i < length; i += 1) {
      const raw = read(i);
      if (raw.trim() === "") break;
      const value = parseNumber(raw);
      if (value === null) return null;
      seeds.push(value);
    }
    if (seeds.length >= 2) {
      const fit = trendFit(type, seeds);
      if (!fit) return null;
      const stopValue = stop === "" ? null : Number(stop);
      const rising = fit(seeds.length + 1) >= fit(seeds.length);
      const values: string[] = [];
      for (let x = 1; x <= length; x += 1) {
        const value = fit(x);
        if (
          x > seeds.length &&
          stopValue !== null &&
          Number.isFinite(stopValue) &&
          (rising ? value > stopValue : value < stopValue)
        ) {
          break;
        }
        values.push(formatNumber(value));
      }
      return { start: 0, values };
    }
    // Fewer than two starting values: an ordinary series from the first cell.
  }
  const first = read(0);
  const step = Number.isFinite(Number(request.step)) ? Number(request.step) : 1;
  if (runsToStop) {
    // A series that never moves (step 0, Growth by ±1 or from 0) would run to
    // the edge of the sheet without meeting its stop value.
    const startNumber = parseNumber(first);
    if (
      step === 0 ||
      (type === "growth" && (Math.abs(step) === 1 || startNumber === 0))
    ) {
      return null;
    }
  }
  const values = computeSeriesValues(first, request, length - 1);
  if (values.length === 0) return null;
  // Running on from one cell must end at the stop value, not the sheet's edge.
  if (runsToStop && values.length === length - 1) return null;
  return { start: 1, values };
}

/**
 * Plan Home > Fill > Series (Linear, Growth or Date) over a selection: every
 * cell the dialog writes, in one list so the caller commits them as one undo
 * step. `valueAt(row, col)` reads a cell's value as a number or date sees it
 * (a formula's result, else its text). AutoFill is not planned here: it is
 * the fill handle's own fill, from {@link planAutoFillSource}.
 */
export function planFillSeries(
  request: FillSeriesRequest,
  valueAt: (row: number, col: number) => string,
): FillSeriesWrite[] {
  const { range } = request;
  const columns = request.seriesIn === "columns";
  const hasStop =
    request.stop !== null &&
    request.stop !== undefined &&
    String(request.stop).trim() !== "";
  const runsToStop = range.r1 === range.r2 && range.c1 === range.c2 && hasStop;
  const lineCount = columns ? range.c2 - range.c1 + 1 : range.r2 - range.r1 + 1;
  const length = runsToStop
    ? columns
      ? request.rowLimit - range.r1
      : request.columnLimit - range.c1
    : columns
      ? range.r2 - range.r1 + 1
      : range.c2 - range.c1 + 1;
  if (length < 2) return [];
  const cellAt = (line: number, index: number) =>
    columns
      ? { row: range.r1 + index, col: range.c1 + line }
      : { row: range.r1 + line, col: range.c1 + index };
  const writes: FillSeriesWrite[] = [];
  for (let line = 0; line < lineCount; line += 1) {
    const plan = planSeriesLine(
      (index) => {
        const { row, col } = cellAt(line, index);
        return valueAt(row, col);
      },
      length,
      request,
      runsToStop,
    );
    if (!plan) continue;
    plan.values.forEach((value, i) => {
      writes.push({ ...cellAt(line, plan.start + i), value });
    });
  }
  return writes;
}

/**
 * The fill-handle drag that Series > AutoFill performs: the selection's
 * leading rows (Series in columns) or columns (Series in rows) that hold
 * anything are the source, and the rest of the selection is filled from them,
 * so AutoFill extends lists, numbered labels, dates and formulas exactly as
 * dragging the fill handle does. Null when there is no source or nothing left
 * to fill.
 */
export function planAutoFillSource(
  range: FillSeriesRange,
  seriesIn: SeriesIn,
  isBlank: (row: number, col: number) => boolean,
): { source: FillSeriesRange; target: { row: number; col: number } } | null {
  if (seriesIn === "columns") {
    let r = range.r1;
    while (r <= range.r2) {
      let blank = true;
      for (let c = range.c1; c <= range.c2 && blank; c += 1) {
        blank = isBlank(r, c);
      }
      if (blank) break;
      r += 1;
    }
    if (r === range.r1 || r > range.r2) return null;
    return {
      source: { ...range, r2: r - 1 },
      target: { row: range.r2, col: range.c1 },
    };
  }
  let c = range.c1;
  while (c <= range.c2) {
    let blank = true;
    for (let r = range.r1; r <= range.r2 && blank; r += 1) {
      blank = isBlank(r, c);
    }
    if (blank) break;
    c += 1;
  }
  if (c === range.c1 || c > range.c2) return null;
  return {
    source: { ...range, c2: c - 1 },
    target: { row: range.r1, col: range.c2 },
  };
}

/**
 * Normalize a raw sheet tab color into the canonical lowercased `#rrggbb` form,
 * or null when it is not a valid hex color. Funnels through the same
 * normalizeHexColor used by cell/conditional-format colors so the tab-color
 * field can never disagree with the rest of the spreadsheet about what a valid
 * color is (named colors, rgb(), CSS-injection attempts all reject), and an
 * empty / whitespace value clears the color (returns null).
 */
export function normalizeTabColor(raw: unknown): string | null {
  if (typeof raw === "string" && raw.trim() === "") return null;
  return normalizeHexColor(raw);
}

// Excel's built-in Custom Lists, which its Sort dialog offers under Order >
// Custom List ahead of the user's own lists: the English day and month lists
// in every language and, in a Simplified Chinese Excel, the zh-CN lists above
// (BUILTIN_CJK_FILL_LISTS less the numerals, which Excel does not ship).
const EXCEL_CUSTOM_LISTS: readonly (readonly string[])[] = [
  WEEKDAYS_SHORT,
  WEEKDAYS_LONG,
  MONTHS_SHORT,
  MONTHS_LONG,
];
const EXCEL_ZH_CN_CUSTOM_LISTS: readonly (readonly string[])[] = [
  ...EXCEL_CUSTOM_LISTS,
  ...BUILTIN_CJK_FILL_LISTS.filter((list) => list.id !== "cjkNumerals").map(
    (list) => list.members,
  ),
];

/**
 * Excel's built-in Custom Lists for a UI locale, in the order its Custom Lists
 * dialog shows them. The same array comes back for the same locale, so a
 * caller can memoize on it.
 */
export function excelBuiltInCustomLists(
  locale: string,
): readonly (readonly string[])[] {
  const tag = locale.toLowerCase();
  return tag === "zh" ||
    tag.startsWith("zh-cn") ||
    tag.startsWith("zh-sg") ||
    tag.startsWith("zh-hans")
    ? EXCEL_ZH_CN_CUSTOM_LISTS
    : EXCEL_CUSTOM_LISTS;
}
