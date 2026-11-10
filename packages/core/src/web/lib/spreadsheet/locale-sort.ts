/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Locale-aware text sort helpers for Chinese collation (WPS Smart Toolbox /
 * Excel "Custom Sort" parity).
 *
 * WPS and Excel both expose two Chinese collation orders for text sort:
 *   - 拼音 (Pinyin): alphabetical by romanisation — the default in V8/Node for
 *     'zh' / 'zh-Hans'. Explicitly requested via the Unicode extension
 *     'zh-Hans-u-co-pinyin'.
 *   - 笔画 (Stroke): by the number and shape of strokes in the first
 *     character, then subsequent characters. Requested via 'zh-Hans-u-co-stroke'.
 *
 * The collators are created once at module scope (or lazily on first use) via
 * featureDetect() so that environments without the Chinese collation extension
 * (e.g. some test runners) fall back gracefully to a standard locale compare.
 * The fallback produces the same ordering as the existing `localeCompare`
 * without an explicit locale, i.e. the engine's default — so existing sort
 * behavior is preserved for users whose JS engine doesn't support the extension.
 */

export type ChineseCollation = "pinyin" | "stroke" | "default";

interface CollatorEntry {
  collator: Intl.Collator;
  /** True when the requested BCP 47 extension was actually honoured. */
  supported: boolean;
}

/**
 * Probe whether `Intl.Collator` with the given locale tag produces a different
 * ordering than a plain 'en' collator on a known discriminating pair
 * ("中" vs "心": pinyin 'zhong' > 'xin'; stroke 中(4) < 心(4+) is equal-or-
 * reversed depending on stroke sub-order but crucially differs from the
 * pinyin order for '中' vs '一').
 *
 * We probe with a pair where pinyin != stroke definitively:
 *   '一' (yī, 1 stroke) vs '中' (zhōng, 4 strokes)
 *   pinyin: '一' < '中' (y < z)
 *   stroke: '一' < '中' (1 < 4) — same direction but *different* from '王' vs '一'
 *
 * A cleaner probe set: ['王','一','中']
 *   pinyin order: 一(y→) < 王(w→) < 中(z→)? No: w < y < z → 王 一 中.
 *   stroke order: 一(1) < 中(4) < 王(4+5=not right – 王 has 4 strokes).
 *
 * We use the canonical probe from the ICU test suite:
 *   pinyin: 啊 < 芭 (ā < bā)
 *   stroke: 啊(11 strokes) > 芭(10 strokes) → 芭 < 啊 in stroke order
 *
 * If the produced order on ['啊','芭'] matches the expected direction for the
 * requested collation, the extension is supported; otherwise we fall back to
 * the standard locale collator.
 */
function buildCollator(localeTag: string): CollatorEntry {
  let collator: Intl.Collator;
  try {
    collator = new Intl.Collator(localeTag, { sensitivity: "variant" });
  } catch {
    // Locale tag not parseable — fall back to default.
    return { collator: new Intl.Collator(), supported: false };
  }

  // Probe pair: '芭' (bā, 10 strokes) vs '啊' (ā, 11 strokes).
  // Expected: pinyin → 啊 < 芭 (ā < bā, i.e. compare('啊','芭') < 0)
  //           stroke → 芭 < 啊 (10 < 11, i.e. compare('芭','啊') < 0)
  const cmpResult = collator.compare("啊", "芭");

  let supported: boolean;
  if (localeTag.includes("co-pinyin")) {
    // pinyin: 啊(ā) should come before 芭(bā) → compare < 0
    supported = cmpResult < 0;
  } else if (localeTag.includes("co-stroke")) {
    // stroke: 啊(11) should come after 芭(10) → compare > 0
    supported = cmpResult > 0;
  } else {
    supported = true; // 'default' / no extension — always OK
  }

  if (!supported) {
    // Extension not honoured; fall back to a plain locale collator.
    return { collator: new Intl.Collator("zh"), supported: false };
  }
  return { collator, supported: true };
}

// Module-level cache so the collators are built at most once per process.
let _pinyinEntry: CollatorEntry | null = null;
let _strokeEntry: CollatorEntry | null = null;

function getPinyinCollator(): CollatorEntry {
  if (!_pinyinEntry) _pinyinEntry = buildCollator("zh-Hans-u-co-pinyin");
  return _pinyinEntry;
}

function getStrokeCollator(): CollatorEntry {
  if (!_strokeEntry) _strokeEntry = buildCollator("zh-Hans-u-co-stroke");
  return _strokeEntry;
}

/**
 * Return the `Intl.Collator` for a given Chinese collation method, with
 * graceful fallback when the requested extension is unsupported.
 *
 * @param method - "pinyin" | "stroke" | "default"
 *   - "default": standard locale compare (no Chinese extension)
 *   - "pinyin": zh-Hans with pinyin collation; falls back to zh if unsupported
 *   - "stroke": zh-Hans with stroke collation; falls back to zh if unsupported
 */
export function getChineseCollator(method: ChineseCollation): Intl.Collator {
  if (method === "pinyin") return getPinyinCollator().collator;
  if (method === "stroke") return getStrokeCollator().collator;
  return new Intl.Collator();
}

/**
 * Whether the engine actually supports the requested Chinese collation
 * extension. Useful for feature-gating UI labels.
 */
export function isChineseCollationSupported(
  method: "pinyin" | "stroke",
): boolean {
  if (method === "pinyin") return getPinyinCollator().supported;
  return getStrokeCollator().supported;
}

/**
 * How a spreadsheet sort orders TEXT against text (Excel's Sort Options).
 * Every field is optional, so a sort level can be passed as-is.
 */
export interface SpreadsheetTextOrder {
  /**
   * Sort Options ▸ Case sensitive. Off (Excel's default): "Apple" and "apple"
   * tie, and the stable sort keeps their input order. On: the lowercase form
   * sorts first, measured in Excel as a | A | ab | Ab | b | B.
   */
  readonly caseSensitive?: boolean;
  /** Locale whose collation orders text (the UI locale). Absent: the engine's. */
  readonly locale?: string;
  /**
   * Sort Options ▸ Method for Chinese text: "pinyin" (letter sort) or "stroke".
   * Absent or "default": the locale's own Han order.
   */
  readonly method?: ChineseCollation;
}

interface TextCollation {
  /** Whole-string order: punctuation, digits, Latin, ..., CJK, as Excel has it. */
  readonly text: Intl.Collator;
  /** Base-letter equality, used to find the first real difference. */
  readonly base: Intl.Collator;
  /** Order of two Han characters when a Chinese order applies, else null. */
  readonly han: Intl.Collator | null;
}

const HAN_SCRIPT = /\p{Script=Han}/u;
const EMPTY_TEXT_ORDER: SpreadsheetTextOrder = {};
const textCollationByKey = new Map<string, TextCollation>();
const textCollationByOrder = new WeakMap<SpreadsheetTextOrder, TextCollation>();
let engineLocale: string | null = null;

/**
 * The locale text sorts in when an order names none. Unset, this process's
 * (the browser's UI locale); the server's calculation session sets the
 * user's, so SORT, MATCH and friends order text as they would in the browser
 * (DESIGN-WAVE2 5.3.6). `null` restores this process's own.
 */
export function setDefaultSortLocale(locale: string | null): void {
  let resolved: string | null = null;
  if (locale) {
    try {
      resolved = new Intl.Collator(locale).resolvedOptions().locale;
    } catch {
      resolved = null;
    }
  }
  engineLocale = resolved;
}

function defaultSortLocale(): string {
  if (engineLocale === null) {
    try {
      engineLocale = new Intl.Collator().resolvedOptions().locale;
    } catch {
      engineLocale = "en";
    }
  }
  return engineLocale;
}

function collatorFor(
  tag: string,
  options: Intl.CollatorOptions,
): Intl.Collator {
  try {
    return new Intl.Collator(tag, options);
  } catch {
    // A structurally invalid tag: the root order still sorts deterministically.
    return new Intl.Collator("en", options);
  }
}

function textCollationFor(order: SpreadsheetTextOrder): TextCollation {
  const known = textCollationByOrder.get(order);
  if (known) return known;
  const locale = order.locale?.trim() || defaultSortLocale();
  const method = order.method ?? "default";
  const caseSensitive = order.caseSensitive === true;
  const key = `${locale}|${method}|${caseSensitive ? "1" : "0"}`;
  let collation = textCollationByKey.get(key);
  if (!collation) {
    const options: Intl.CollatorOptions = caseSensitive
      ? { sensitivity: "variant", caseFirst: "lower" }
      : { sensitivity: "accent" };
    const base = collatorFor("en", { sensitivity: "base" });
    if (method !== "default" || /^zh(?:-|$)/i.test(locale)) {
      // ICU's Chinese collations move every Han character ahead of Latin;
      // Excel keeps Latin first (measured: Zed | 王 | 一 | 中) and orders only
      // Han against Han by pinyin or stroke. The root order carries the rest.
      const hanTag =
        method === "pinyin" || method === "stroke"
          ? isChineseCollationSupported(method)
            ? `zh-Hans-u-co-${method}`
            : "zh"
          : locale;
      collation = {
        text: collatorFor("en", options),
        base,
        han: collatorFor(hanTag, options),
      };
    } else {
      collation = { text: collatorFor(locale, options), base, han: null };
    }
    textCollationByKey.set(key, collation);
  }
  textCollationByOrder.set(order, collation);
  return collation;
}

/**
 * Where two strings first differ in their base letters (case and accents
 * aside), as the characters at that point and their offsets; null when one is
 * a base-letter prefix of the other.
 */
function firstBaseDifference(
  a: string,
  b: string,
  base: Intl.Collator,
): { a: string; b: string; atA: number; atB: number } | null {
  let atA = 0;
  let atB = 0;
  while (atA < a.length && atB < b.length) {
    const charA = String.fromCodePoint(a.codePointAt(atA)!);
    const charB = String.fromCodePoint(b.codePointAt(atB)!);
    if (
      charA !== charB &&
      charA.toLowerCase() !== charB.toLowerCase() &&
      base.compare(charA, charB) !== 0
    ) {
      return { a: charA, b: charB, atA, atB };
    }
    atA += charA.length;
    atB += charB.length;
  }
  return null;
}

// ── Text in a sort, as Excel orders it (Windows' word sort) ─────────────────
//

// (34 groups), xprobe-arabicsort2 (6) and xprobe-arabicsort (R166-X13), each
// group sorted as written and reversed, so a tie shows as the input order kept.
// - Case does not count; full and half width do not count (Ａ = A, ｱ = ア).
// - Apostrophe, hyphen and the dashes (- ‐ ‑ – —) are ignored at first
//   (coop < co-op < cooperate < co-operate; "1" < "-1"). Between texts equal
//   without them: fewer first, then the later one first, then ' < - < ‐ < ‑ <
//   – < — (a < a- < -a < --a; don't < don-t).
// - Space, no-break space and punctuation count, before digits, before letters,
//   in Excel's order: ' ' NBSP ! " ‘ ’ “ ” « » # $ £ € % & ( ) * , . / : ; ? @
//   [ \ ] ^ _ ` { | } ~ + < = > ± × ÷ © ® °; other symbols after, by code.
// - Digits by value, one by one ("10" < "2"); Arabic-Indic digits after the
//   ASCII ones they equal.
// - Letters in ICU's root order, one by one; æ = ae, œ = oe, ß = ss (ties);
//   accents after the letters, left to right, é < è < ê < ë < å (ä before å),
//   ø after ö; katakana before hiragana (ア < あ).
// - Arabic script: tatweel ignored; marks after the letters; the alef forms
//   one letter (ٱ < ا < إ < أ < آ); ؤ and ئ after و and ي; ں just before ن,
//   ۀ just after ی.
// - Han: by the order's Han collation (pinyin or stroke in Chinese), else ICU's
//   root order; a Chinese Excel measured Latin < kana < Han, by pinyin.
// ICU's own string order (compareSpreadsheetText) differed on 44 of the 68
// Latin sorts and on the Arabic ones.

const WORD_SORT_PUNCTUATION = [
  " ", " ", "!", '"', "‘", "’", "“", "”", "«", "»", "#", "$", "£", "€",
  "%", "&", "(", ")", "*", ",", ".", "/", ":", ";", "?", "@", "[", "\\", "]", "^", "_", "`", "{", "|", "}", "~",
  "+", "<", "=", ">", "±", "×", "÷", "©", "®", "°",
];
const WORD_SORT_PUNCTUATION_RANK = new Map(WORD_SORT_PUNCTUATION.map((ch, index) => [ch, index]));
const WORD_SORT_IGNORED = new Map([["'", 0], ["-", 1], ["‐", 2], ["‑", 3], ["–", 4], ["—", 5]]);
const WORD_SORT_EXPANSION: Readonly<Record<string, string>> = {
  "æ": "ae", "Æ": "ae", "œ": "oe", "Œ": "oe", "ß": "ss",
};
const WORD_SORT_STROKE: Readonly<Record<string, string>> = {
  "ø": "o", "Ø": "o", "đ": "d", "Đ": "d", "ł": "l", "Ł": "l", "ħ": "h", "Ħ": "h",
};
/** The order of the marks on one letter: acute, grave, breve, circumflex, caron, diaeresis, ring... */
const WORD_SORT_MARK: Readonly<Record<string, number>> = {
  "́": 1, "̀": 2, "̆": 3, "̂": 4, "̌": 5, "̈": 6, "̊": 7, "̋": 8,
  "̃": 9, "̇": 10, "̨": 11, "̧": 12, "̄": 13,
};
const ARABIC_ALEF_WEIGHT: Readonly<Record<string, number>> = {
  "ٱ": -1, "ا": 0, "إ": 1, "أ": 2, "آ": 3,
};
const ARABIC_HAMZA_SEAT: Readonly<Record<string, string>> = { "ؤ": "و", "ئ": "ي" };
const KATAKANA_SCRIPT = /\p{Script=Katakana}/u;
const SYMBOL_OR_SPACE = /[\p{P}\p{S}\p{Z}]/u;
const LETTER = /\p{L}/u;

interface WordSortKey {
  /** Class and rank of each element, as one number (1 punctuation, 2 digit, 3 letter). */
  primary: number[];
  /** Per element: the weight after its letter (alef form, hamza seat, digit script). */
  secondary: number[];
  /** Per element: its Arabic marks, as written. */
  marks: string[];
  /** Per element: its accents (WORD_SORT_MARK, read left to right). */
  accents: number[];
  /** Per element: 0 katakana, 1 anything else. */
  kana: number[];
  /** The ignored characters: -position * 8 + their own order. */
  ignored: number[];
}

let letterRanks: Map<string, number> | null = null;
const hanRanks = new Map<string, Map<string, number>>();

function ranksOf(collator: Intl.Collator, take: (ch: string) => boolean): Map<string, number> {
  const chars: string[] = [];
  for (let code = 0x41; code < 0x10000; code += 1) {
    if (code >= 0xd800 && code < 0xe000) continue;
    const ch = String.fromCharCode(code);
    if (take(ch)) chars.push(ch);
  }
  chars.sort(collator.compare);
  const ranks = new Map<string, number>();
  let rank = 0;
  chars.forEach((ch, index) => {
    if (index > 0 && collator.compare(chars[index - 1], ch) !== 0) rank += 1;
    ranks.set(ch, rank);
  });
  return ranks;
}

function letterRank(ch: string, han: Intl.Collator | null): number {
  letterRanks ??= ranksOf(
    new Intl.Collator("en", { sensitivity: "base" }),
    (candidate) => LETTER.test(candidate) && !HAN_SCRIPT.test(candidate) && candidate.normalize("NFD").length === 1,
  );
  const known = letterRanks.get(ch.toLowerCase()) ?? letterRanks.get(ch);
  if (known !== undefined) return known;
  if (HAN_SCRIPT.test(ch)) {
    const collator = han ?? new Intl.Collator("en", { sensitivity: "base" });
    const key = han ? JSON.stringify(han.resolvedOptions()) : "root";
    let ranks = hanRanks.get(key);
    if (!ranks) {
      ranks = ranksOf(collator, (candidate) => HAN_SCRIPT.test(candidate));
      hanRanks.set(key, ranks);
    }
    return 100_000 + (ranks.get(ch) ?? ch.codePointAt(0)!);
  }
  return 1_000_000 + ch.codePointAt(0)!;
}

const wordSortKeys = new Map<string, WordSortKey>();
const MAX_WORD_SORT_KEYS = 200_000;
let wordSortKeysHan: Intl.Collator | null = null;

function wordSortKey(text: string, han: Intl.Collator | null): WordSortKey {
  if (han !== wordSortKeysHan) {
    wordSortKeys.clear();
    wordSortKeysHan = han;
  }
  const known = wordSortKeys.get(text);
  if (known) return known;
  const key: WordSortKey = { primary: [], secondary: [], marks: [], accents: [], kana: [], ignored: [] };
  const push = (cls: number, rank: number, secondary = 0, accents = 0, kana = 1) => {
    key.primary.push(cls * 10_000_000 + rank);
    key.secondary.push(secondary);
    key.marks.push("");
    key.accents.push(accents);
    key.kana.push(kana);
  };
  const letter = (ch: string, place: number) => letterRank(ch, han) * 4 + place;
  let position = 0;
  for (const raw of text) {
    const code = raw.codePointAt(0)!;
    if (code === 0x640) continue;
    if ((code >= 0x64b && code <= 0x65f) || code === 0x670) {
      const last = key.marks.length - 1;
      if (last >= 0) key.marks[last] += raw;
      continue;
    }
    const ignored = WORD_SORT_IGNORED.get(raw);
    if (ignored !== undefined) {
      key.ignored.push(-position * 8 + ignored);
      continue;
    }
    position += 1;
    const punctuation = WORD_SORT_PUNCTUATION_RANK.get(raw);
    if (punctuation !== undefined) {
      push(1, punctuation);
      continue;
    }
    for (const ch of WORD_SORT_EXPANSION[raw] ?? raw.normalize("NFKC")) {
      const c = ch.codePointAt(0)!;
      const folded = WORD_SORT_PUNCTUATION_RANK.get(ch);
      if (c >= 0x30 && c <= 0x39) push(2, c - 0x30);
      else if (c >= 0x660 && c <= 0x669) push(2, c - 0x660, 1);
      else if (c >= 0x6f0 && c <= 0x6f9) push(2, c - 0x6f0, 2);
      else if (folded !== undefined) push(1, folded);
      else if (SYMBOL_OR_SPACE.test(ch)) push(1, 1000 + c);
      else if (ch in ARABIC_ALEF_WEIGHT) push(3, letter("ا", 2), ARABIC_ALEF_WEIGHT[ch]);
      else if (ch in ARABIC_HAMZA_SEAT) push(3, letter(ARABIC_HAMZA_SEAT[ch], 2), 1);
      else if (ch === "ں") push(3, letter("ن", 1));
      else if (ch === "ۀ") push(3, letter("ی", 3));
      else if (ch in WORD_SORT_STROKE) push(3, letter(WORD_SORT_STROKE[ch], 2), 0, 900);
      else {
        const decomposed = ch.normalize("NFD");
        let accents = 0;
        for (const mark of decomposed.slice(1)) accents = accents * 32 + (WORD_SORT_MARK[mark] ?? 30);
        push(3, letter(decomposed[0], 2), 0, accents, KATAKANA_SCRIPT.test(ch) ? 0 : 1);
      }
    }
  }
  if (wordSortKeys.size >= MAX_WORD_SORT_KEYS) wordSortKeys.clear();
  wordSortKeys.set(text, key);
  return key;
}

/**
 * Excel's order of two texts in a sort (case-insensitive), as measured: see
 * the notes above. `han` orders Han characters among themselves (null: ICU's
 * root order).
 */
export function compareExcelSortText(a: string, b: string, han: Intl.Collator | null = null): number {
  const left = wordSortKey(a, han);
  const right = wordSortKey(b, han);
  const n = Math.min(left.primary.length, right.primary.length);
  for (let i = 0; i < n; i += 1) {
    if (left.primary[i] !== right.primary[i]) return left.primary[i] - right.primary[i];
  }
  if (left.primary.length !== right.primary.length) return left.primary.length - right.primary.length;
  for (let i = 0; i < n; i += 1) {
    if (left.secondary[i] !== right.secondary[i]) return left.secondary[i] - right.secondary[i];
    if (left.marks[i] !== right.marks[i]) return left.marks[i] < right.marks[i] ? -1 : 1;
    if (left.accents[i] !== right.accents[i]) return left.accents[i] - right.accents[i];
    if (left.kana[i] !== right.kana[i]) return left.kana[i] - right.kana[i];
  }
  if (left.ignored.length !== right.ignored.length) return left.ignored.length - right.ignored.length;
  for (let i = 0; i < left.ignored.length; i += 1) {
    if (left.ignored[i] !== right.ignored[i]) return left.ignored[i] - right.ignored[i];
  }
  return 0;
}

/**
 * Excel's order of two text values in a sort (ascending sign; the caller
 * applies direction). Case-insensitive by default, lowercase first when case
 * sensitive. With a Chinese locale or an explicit method, Latin text still
 * sorts before Chinese and Han characters order by pinyin (or stroke) among
 * themselves, the order measured in zh Excel: Zed | 王 | 一 | 中.
 */
export function compareSpreadsheetText(
  a: string,
  b: string,
  order: SpreadsheetTextOrder = EMPTY_TEXT_ORDER,
): number {
  const collation = textCollationFor(order);
  if (collation.han === null || !HAN_SCRIPT.test(a) || !HAN_SCRIPT.test(b)) {
    return collation.text.compare(a, b);
  }
  const difference = firstBaseDifference(a, b, collation.base);
  if (
    difference &&
    HAN_SCRIPT.test(difference.a) &&
    HAN_SCRIPT.test(difference.b)
  ) {
    const byHan = collation.han.compare(
      a.slice(difference.atA),
      b.slice(difference.atB),
    );
    if (byHan !== 0) return byHan;
  }
  return collation.text.compare(a, b);
}







export function compareSpreadsheetSortText(
  a: string,
  b: string,
  order: SpreadsheetTextOrder = EMPTY_TEXT_ORDER,
): number {
  if (order.caseSensitive === true) return compareSpreadsheetText(a, b, order);
  return compareExcelSortText(a, b, textCollationFor(order).han);
}

/**
 * Build a comparator function suitable for `Array.prototype.sort` that sorts
 * text values using the given Chinese collation order. Numbers and blanks are
 * handled by the caller (this comparator is only invoked for text-vs-text
 * comparisons in the sort pipeline; integrate via the `compareCells` wiring).
 *
 * @param method  - collation method
 * @param dir     - "asc" | "desc"
 */
export function buildChineseTextComparator(
  method: ChineseCollation,
  dir: "asc" | "desc",
): (a: string, b: string) => number {
  // One text order for every sort: the spreadsheet's own comparator, with the
  // requested Chinese method (Latin still ahead of Han, as in Excel).
  const order: SpreadsheetTextOrder = { method };
  return (a, b) => {
    const cmp = compareSpreadsheetText(a, b, order);
    return dir === "desc" ? -cmp : cmp;
  };
}

/**
 * Sort a list of string values using the specified Chinese collation, returning
 * a new sorted array. Blank strings always sink to the end regardless of
 * direction (mirrors the spreadsheet sort convention).
 *
 * This is the pure, unit-testable entry point used by tests.
 *
 * @param values  - array of strings to sort
 * @param method  - collation method
 * @param dir     - "asc" | "desc"
 */
export function sortChineseStrings(
  values: string[],
  method: ChineseCollation,
  dir: "asc" | "desc" = "asc",
): string[] {
  const compare = buildChineseTextComparator(method, dir);
  // Separate blanks (always last) from non-blanks.
  const nonBlanks = values.filter((v) => v.trim() !== "");
  const blanks = values.filter((v) => v.trim() === "");
  const sorted = [...nonBlanks].sort(compare);
  return [...sorted, ...blanks];
}
