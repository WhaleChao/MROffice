/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { lineBreakOpportunities } from "./line-break";
import { KINSOKU_EAST_ASIAN, OFFICE_KINSOKU } from "./kinsoku";

/** The rules a run's words break by: a language's, or "none" (no list, no hanging). */
export type PowerPointLineLanguage = "ja" | "zh-TW" | "zh-CN" | "none";

const set = (chars: string): ReadonlySet<string> => new Set([...chars]);

/** Office's line-start marks PowerPoint lets start a line all the same (hang-full-probe "NONE"). */
const NOT_HELD: Readonly<Record<Exclude<PowerPointLineLanguage, "none">, ReadonlySet<string>>> = {
  ja: set("¢"),
  "zh-TW": set("¢•‧╴］｝､"),
  "zh-CN": set("%>¢°ˉ―‰′″›℃〞︶︺︾﹀﹄﹚﹜﹞％￠"),
};

const withoutHeld = (language: Exclude<PowerPointLineLanguage, "none">): ReadonlySet<string> =>
  new Set([...OFFICE_KINSOKU[language].noStart].filter((char) => !NOT_HELD[language].has(char)));

const NO_START: Readonly<Record<PowerPointLineLanguage, ReadonlySet<string>>> = {
  ja: withoutHeld("ja"),
  "zh-TW": withoutHeld("zh-TW"),
  "zh-CN": withoutHeld("zh-CN"),
  none: new Set(),
};

const NO_END: Readonly<Record<PowerPointLineLanguage, ReadonlySet<string>>> = {
  ja: OFFICE_KINSOKU.ja.noEnd,
  "zh-TW": OFFICE_KINSOKU["zh-TW"].noEnd,
  "zh-CN": OFFICE_KINSOKU["zh-CN"].noEnd,
  none: new Set(),
};

/** The marks PowerPoint hangs past a line's end (hang-full-probe "HANG"). */
const HANGS: Readonly<Record<PowerPointLineLanguage, ReadonlySet<string>>> = {
  ja: set(",.、。，．"),
  "zh-TW": set("!),.:;?]}’”、。〉》」』】〕！），：；？"),
  "zh-CN": set("!),.:;?]}’”、。〉》」』】〕！），：；？〗］｝"),
  none: new Set(),
};

/** A body's run whose language nothing states: the rules every slide body had before row 1001. */
const LEGACY = OFFICE_KINSOKU["zh-CN"];

/** A letter of an East Asian script (a mark is not one). */
const EAST_ASIAN_LETTER = /^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const NO_BREAK_BEFORE = new Set([" ", " ", "　", "\t", "\n", "\u000b", " "]);
/** A language with lists (and hanging): ja, zh-TW, zh-CN. */
const isListed = (language: PowerPointLineLanguage | null): boolean => language !== null && language !== "none";
/** Middle dots that join two words' letters ("교·강사", "A・B"). */
const WORD_DOTS = new Set(["\u00b7", "\u30fb", "\u318d"]);

function eastAsianOf(tag: string): PowerPointLineLanguage | null {
  const parts = tag.toLowerCase().split("-");
  const primary = parts[0];
  if (primary === "ja") return "ja";
  if (primary === "ko") return "none";
  if (primary !== "zh") return null;
  return parts.some((part) => part === "tw" || part === "hk" || part === "mo" || part === "hant") ? "zh-TW" : "zh-CN";
}

/**
 * The rules a run breaks by from its `lang` and `altLang` (its own, else what
 * its row inherits): an East Asian `lang` its own; a Western one the East Asian
 * `altLang`'s, else "none". Null when no `lang` is known and no East Asian
 * `altLang` either: the caller keeps the old rules for such a run (probe4: on
 * the truth VM, whose editing language is en-GB, such a run breaks as "none";
 * whether that is the editing language's doing is not known, so pages without a
 * language do not move).
 */
export function powerPointLineLanguage(lang: unknown, altLang: unknown): PowerPointLineLanguage | null {
  const own = typeof lang === "string" && lang.trim() ? lang.trim() : null;
  const alternate = typeof altLang === "string" && altLang.trim() ? eastAsianOf(altLang.trim()) : null;
  if (own !== null) return eastAsianOf(own) ?? alternate ?? "none";
  return alternate;
}

/** Whether `mark` may not start a line in `language`'s run (null: the old list). */
export function powerPointHoldsAtLineStart(mark: string, language: PowerPointLineLanguage | null): boolean {
  return (language === null ? LEGACY.noStart : NO_START[language]).has(mark);
}

/** Whether PowerPoint hangs `mark` past a line's end in `language`'s run. */
export function powerPointHangs(mark: string, language: PowerPointLineLanguage | null): boolean {
  return language !== null && HANGS[language].has(mark);
}

/**
 * Whether a line may break before each grapheme (index 0 always false), each
 * grapheme in its run's language. Between two graphemes whose language nothing
 * states, exactly the old rules (UAX #14 with PowerPoint's tailoring and Office's
 * Simplified Chinese list); between two of "none", Latin words' rules, East
 * Asian letters read as Latin letters.
 */
export function powerPointLineBreakOpportunities(
  graphemes: readonly string[],
  languages: readonly (PowerPointLineLanguage | null)[],
  options: { readonly latinLineBreak?: boolean; readonly eastAsianRules?: boolean } = {},
): boolean[] {
  const base = lineBreakOpportunities(graphemes, { kinsoku: null, powerpoint: true });
  const legacy = lineBreakOpportunities(graphemes, { kinsoku: "zh-CN", powerpoint: true });
  // "none": a letter of a word -- an East Asian letter, or a middle dot between two.
  const wordLetter = (index: number): boolean =>
    EAST_ASIAN_LETTER.test(graphemes[index] ?? "") ||
    (WORD_DOTS.has(graphemes[index] ?? "") && EAST_ASIAN_LETTER.test(graphemes[index - 1] ?? "") && EAST_ASIAN_LETTER.test(graphemes[index + 1] ?? ""));
  const out = new Array<boolean>(graphemes.length).fill(false);
  for (let index = 1; index < graphemes.length; index += 1) {
    const before = languages[index - 1] ?? null;
    const after = languages[index] ?? null;
    if (before === null && after === null) {
      out[index] = legacy[index];
      continue;
    }
    if (before === "none" && after === "none") {
      out[index] = base[index] && (options.latinLineBreak === true || !(wordLetter(index - 1) && wordLetter(index)));
      continue;
    }
    const previous = graphemes[index - 1];
    const next = graphemes[index];
    if (
      options.eastAsianRules === false && (isListed(before) || isListed(after)) &&
      (KINSOKU_EAST_ASIAN.test(previous) || KINSOKU_EAST_ASIAN.test(next)) &&
      !NO_BREAK_BEFORE.has(next) && !NO_BREAK_BEFORE.has(previous)
    ) {
      out[index] = true;
      continue;
    }
    let open = base[index];
    if (
      !open && isListed(after) && (EAST_ASIAN_LETTER.test(previous) || EAST_ASIAN_LETTER.test(next)) &&
      !NO_BREAK_BEFORE.has(next) && !NO_BREAK_BEFORE.has(previous)
    ) open = true;
    if (open && (KINSOKU_EAST_ASIAN.test(previous) || KINSOKU_EAST_ASIAN.test(next))) {
      const start = after === null ? LEGACY.noStart : NO_START[after];
      const end = before === null ? LEGACY.noEnd : NO_END[before];
      if (start.has(next) || end.has(previous)) open = false;
    }
    out[index] = open;
  }
  return out;
}
