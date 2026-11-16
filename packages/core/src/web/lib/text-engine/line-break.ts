/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Line breaking over measured graphemes.
 *
 * Opportunities follow UAX #14 (the Unicode line breaking algorithm the
 * browsers' ICU implements, rules LB7-LB31, with CSS `line-break: normal`'s
 * treatment of small kana) on grapheme clusters, plus Office's kinsoku lists
 * (kinsoku.ts) between East Asian characters. Lines are filled greedily the
 * way CSS `white-space: pre-wrap; overflow-wrap: anywhere` fills them:
 * spaces at a break hang past the edge, and a word wider than the line
 * breaks between graphemes. The same graphemes and widths give the same
 * lines in every browser.
 */
import { KINSOKU_EAST_ASIAN, OFFICE_KINSOKU, type KinsokuList } from "./kinsoku";

type Lb =
  | "SP" | "ZW" | "GL" | "WJ" | "CL" | "CP" | "EX" | "IS" | "SY" | "OP" | "QU" | "NS" | "BA" | "HY"
  | "B2" | "BB" | "IN" | "AL" | "NU" | "PR" | "PO" | "ID";

const EXPLICIT: ReadonlyMap<number, Lb> = (() => {
  const map = new Map<number, Lb>();
  const put = (cls: Lb, chars: string) => { for (const char of chars) map.set(char.codePointAt(0)!, cls); };
  const range = (cls: Lb, from: number, to: number) => { for (let code = from; code <= to; code += 1) map.set(code, cls); };
  map.set(0x20, "SP");
  map.set(0x200b, "ZW");
  for (const code of [0xa0, 0x202f, 0x2007, 0x2011, 0x0f0c, 0x0f12, 0x034f, 0x180e]) map.set(code, "GL");
  for (const code of [0x2060, 0xfeff]) map.set(code, "WJ");
  put("EX", "!?！？");
  put("CP", ")]");
  put("CL", "}、。〉》」』】〕〗〙〛〞〟︐︑︒﹐﹒），．］｝｠｡｣､︶︸︺︼︾﹀﹂﹄﹈﹚﹜﹞");
  put("OP", "([{¡¿‚„〈《「『【〔〖〘〚〝︵︷︹︻︽︿﹁﹃﹇﹙﹛﹝（［｛｟｢");
  put("QU", "\"'«»‘’‛“”‟‹›❛❜❝❞⸀⸁⸂⸃⸄⸅⸆⸇⸈⸉⸊⸋⸌⸍⸜⸝⸠⸡");
  put("IS", ",.:;;։،؍߸⁄︐︓︔");
  put("SY", "/");
  put("HY", "-");
  put("BA", "\t\u00ad֊‐‒–‧\u3000|\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2008\u2009\u200a\u205f⳺⳻⳼⳽⳿⸎⸏⸐⸑⸒⸓⸔⸕⸗⸙⸪⸫⸬⸭⸰⸱⸳⸴⸼⸽⸾");
  put("B2", "—⸺⸻");
  put("BB", "´ˈˌ˟´꡴꡵");
  put("IN", "․‥…︙");
  put("NS", "៖‼‽⁇⁈⁉々〜〻〼゛゜ゝゞ゠・ヽヾꀕ﹔﹕：；･ﾞﾟ");
  put("PR", "$+\\£¥±№−∓＄￡￥￦");
  range("PR", 0x20a0, 0x20a6);
  range("PR", 0x20a8, 0x20b5);
  range("PR", 0x20b7, 0x20ba);
  range("PR", 0x20bc, 0x20bd);
  range("PR", 0x20bf, 0x20cf);
  put("PO", "%¢°؉؊؋٪‰‱′″‴‵‶‷₧₶₻₾℃℉﷼﹪％￠");
  return map;
})();

/** What Word's mode-13 Latin breaking takes as letters (ideographsAsLetters). */
const LATIN_SET_IDEOGRAPH = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}⺀-⿿\u3001-〿㇀-㇯㈀-㏿︰-﹏\uff00-\uffef]/u;

const HAN_LIKE = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}⺀-⿿\u3000-〿㇀-㇯㈀-㏿︰-﹏\uff00-\uffef]/u;

function classOf(grapheme: string): Lb {
  const code = grapheme.codePointAt(0)!;
  const explicit = EXPLICIT.get(code);
  if (explicit) return explicit;
  if (code >= 0x30 && code <= 0x39) return "NU";
  if (code >= 0xff10 && code <= 0xff19) return "NU";
  if (/\p{Nd}/u.test(grapheme.charAt(0))) return "NU";
  // Hangul jamo sequences and syllables, Han, kana (small kana included:
  // CSS line-break: normal) and the full-width forms break like ideographs.
  if (HAN_LIKE.test(String.fromCodePoint(code))) return "ID";
  return "AL";
}

export interface LineBreakOptions {
  /** Office's East Asian line breaking list; null for none. */
  readonly kinsoku: KinsokuList | null;
  /** Word's own tailoring (Docs): a slash inside a word does not open a
   * break ("and/or", "https://host/1234/path" stay whole), as Word for Mac and
   * the browsers keep them (corpus R279-D21, R280-D14, R280-D30). */
  readonly word?: boolean;
  





  readonly powerpoint?: boolean;
  






  readonly browser?: boolean;
  





  readonly continuesText?: boolean;
  /** Windows Excel's wrap of a cell (see {@link excelLineBreakOpportunities}):
   * every other option is ignored. */
  readonly excel?: boolean;
  /** Windows Word's East Asian text in a document whose defaults name no
   * East Asian face, in compatibility mode 13 and earlier
   * (wordSetsIdeographsAsLetters): ideographs, kana and the full-width forms
   * are letters of a word, broken only where a Latin word would be. Hangul
   * and the ideographic space keep their own class. */
  readonly ideographsAsLetters?: boolean;
}














export function excelLineBreakOpportunities(graphemes: readonly string[]): boolean[] {
  const out = new Array<boolean>(graphemes.length).fill(false);
  let worded = false;
  for (let i = 1; i < graphemes.length; i += 1) {
    const before = graphemes[i - 1];
    if (before !== " ") worded = true;
    out[i] = worded && graphemes[i] !== " " && (before === " " || before === "-");
  }
  return out;
}

/** Closing (final) quotation marks, `\p{Pf}` among UAX #14's QU. */
const FINAL_QUOTES = new Set(Array.from("’”›»⸃⸅⸊⸍⸝⸡"));
/** Opening (initial) quotation marks, `\p{Pi}` among UAX #14's QU. */
const INITIAL_QUOTES = new Set(Array.from("‘“‹«‛‟⸂⸄⸉⸌⸜⸠"));











function quoteBreaksAmongEastAsian(graphemes: readonly string[], classes: readonly Lb[], i: number): boolean {
  const eastAsian = (grapheme: string | undefined) => grapheme !== undefined && KINSOKU_EAST_ASIAN.test(grapheme);
  if (classes[i] === "QU" && !(INITIAL_QUOTES.has(graphemes[i]!) && eastAsian(graphemes[i - 1]) && eastAsian(graphemes[i + 1]))) return false;
  if (classes[i - 1] === "QU" && !(FINAL_QUOTES.has(graphemes[i - 1]!) && eastAsian(graphemes[i]) && eastAsian(graphemes[i - 2]))) return false;
  return true;
}

/** After these, Chromium breaks only before an opening bracket. */
const ASCII_BREAK_BEFORE_OPEN = new Set(Array.from('!"#%&)*+,.:;=>\\]|}~'));
const ASCII_OPENING = new Set(Array.from("(<[{"));
/** What Chromium breaks before after a hyphen-minus (a digit only after a
 * letter or digit: "-5" and "(-5" stay whole, "ABCD-1234" breaks). */
const ASCII_AFTER_HYPHEN = new Set(Array.from("\"#%&'(*+-<=>@[\\^_`{|~"));
/** What Chromium breaks before after a question mark. */
const ASCII_AFTER_QUESTION = new Set(Array.from("#$%&(*+-0123456789<=>@[\\^_`{|~"));

function asciiPrintable(grapheme: string | undefined): boolean {
  if (grapheme === undefined || grapheme.length !== 1) return false;
  const code = grapheme.charCodeAt(0);
  return code > 0x20 && code < 0x7f;
}

function asciiAlphanumeric(grapheme: string | undefined): boolean {
  return grapheme !== undefined && grapheme.length === 1 && /[0-9A-Za-z]/.test(grapheme);
}

/** Chromium's break between two printable ASCII characters, `previous`
 * the one before `before` (or undefined). */
function chromiumAsciiBreak(previous: string | undefined, before: string, after: string): boolean {
  if (before === "-") {
    if (after >= "0" && after <= "9") return asciiAlphanumeric(previous);
    return /[A-Za-z]/.test(after) || ASCII_AFTER_HYPHEN.has(after);
  }
  if (before === "?") return /[A-Za-z]/.test(after) || ASCII_AFTER_QUESTION.has(after);
  return ASCII_BREAK_BEFORE_OPEN.has(before) && ASCII_OPENING.has(after);
}

/**
 * Whether a line may break before each grapheme (index 0 is always false).
 */
export function lineBreakOpportunities(graphemes: readonly string[], options: LineBreakOptions): boolean[] {
  if (options.excel) return excelLineBreakOpportunities(graphemes);
  const count = graphemes.length;
  // The browser breaks between pictographs (emoji), as between ideographs.
  const classes = graphemes.map((grapheme): Lb =>
    options.browser && /^\p{Extended_Pictographic}/u.test(grapheme)
      ? "ID"
      : options.ideographsAsLetters && LATIN_SET_IDEOGRAPH.test(String.fromCodePoint(grapheme.codePointAt(0) ?? 0x20))
        ? "AL"
        : classOf(grapheme),
  );
  const out = new Array<boolean>(count).fill(false);
  const kinsoku = options.kinsoku ? OFFICE_KINSOKU[options.kinsoku] : null;
  // Index of the last non-space before position i (for the "X SP* ×" rules).
  let lastNonSpace = -1;
  for (let i = 1; i < count; i += 1) {
    if (classes[i - 1] !== "SP") lastNonSpace = i - 1;
    const before = classes[i - 1];
    const after = classes[i];
    const beforeSpaces = lastNonSpace >= 0 ? classes[lastNonSpace] : undefined;
    // After a run of spaces, before anything else: Chromium takes every space
    // run as a break, past UAX #14's OP SP* × and × CL and past the CJK line
    // start rules ("( £12" broke after the bracket, "xx 。" before the stop:
    
    const afterSpaceRun = options.browser === true && graphemes[i - 1] === " " && graphemes[i] !== " " && after !== "SP";
    out[i] =
      options.browser && asciiPrintable(graphemes[i - 1]) && asciiPrintable(graphemes[i])
        ? chromiumAsciiBreak(graphemes[i - 2], graphemes[i - 1], graphemes[i])
        : afterSpaceRun || pairBreaks(before, after, beforeSpaces, lastNonSpace === i - 1,
          (options.word === true || options.browser === true) && lastNonSpace === i - 1 && quoteBreaksAmongEastAsian(graphemes, classes, i));
    // Nor after the spaces a text opens with: no line of spaces alone.
    if (out[i] && options.browser && lastNonSpace < 0 && !options.continuesText) out[i] = false;
    // After a number too: Word for Mac's PDF of corpus R282-D08 (p22) left 215 pt of a line empty
    // before "https://www.unicef.org/media/179001/file/..." whose head to "179001/" takes 190 pt,
    // and R283-D09 (p23) 435 pt before a head of 333 pt ending "/245181468009603809/".
    if (out[i] && options.word && before === "SY" && after !== "SP" && i >= 2) out[i] = false;
    // A slash that stands apart opens a line like a word: Word for Mac's PDF of
    // corpus R279-D27 (p8, a 54 pt cell) draws "Final TP / FP" and "/ FN" under
    // it, as the browsers do; UAX #14's "SP* x SY" took "FP" down with it.
    // (One place in the corpus; no fixture yet.)
    if (!out[i] && options.word && after === "SY" && before === "SP") out[i] = true;
    if (out[i] && options.powerpoint && (before === "SY" || before === "EX") && (after === "AL" || after === "NU") && i >= 2 && classes[i - 2] !== "SP") out[i] = false;
    // A closing quotation mark and spaces before an opening bracket: PowerPoint
    
    // ends a line at the quote, as Chromium does there too): today's UAX #14
    // has no QU SP* × OP (LB15a/b replaced it). Only a closing quote: a
    // straight one may open the bracket's quotation, and keeps the old rule.
    if (!out[i] && options.powerpoint && after === "OP" && before === "SP" && beforeSpaces === "QU" && FINAL_QUOTES.has(graphemes[lastNonSpace])) out[i] = true;
    // A hyphen is a minus sign only where a number starts ("-5", "(-5"):
    // after a word or a number it is a hyphen a line may break after
    // ("pre-|2025", "2025-|2027"), as Word and the browsers break it
    // (corpus R280-D30; Unicode 15.1's LB25).
    if (!out[i] && options.word && before === "HY" && after === "NU" && i >= 2 && (classes[i - 2] === "AL" || classes[i - 2] === "NU" || classes[i - 2] === "ID")) out[i] = true;
    if (out[i] && kinsoku && !afterSpaceRun) {
      const previous = graphemes[i - 1];
      const next = graphemes[i];
      const eastAsian = KINSOKU_EAST_ASIAN.test(previous) || KINSOKU_EAST_ASIAN.test(next);
      if (eastAsian && (kinsoku.noStart.has(next) || kinsoku.noEnd.has(previous))) out[i] = false;
    }
  }
  return out;
}

/** UAX #14 LB7-LB31 for the pair (before, after); `lead` is the class
 * before any spaces ending at `before`, `adjacent` whether no space
 * separates them. */
function pairBreaks(before: Lb, after: Lb, lead: Lb | undefined, adjacent: boolean, quoteAmongEastAsian = false): boolean {
  // LB7
  if (after === "SP" || after === "ZW") return false;
  // LB8
  if (before === "ZW" || (!adjacent && lead === "ZW")) return true;
  // LB11
  if (before === "WJ" || after === "WJ") return false;
  // LB12, LB12a
  if (before === "GL") return false;
  if (after === "GL" && before !== "SP" && before !== "BA" && before !== "HY") return false;
  // LB13
  if (after === "CL" || after === "CP" || after === "EX" || after === "IS" || after === "SY") return false;
  // LB14: OP SP* ×
  if (lead === "OP") return false;
  // LB15: QU SP* × OP
  if (lead === "QU" && after === "OP") return false;
  // LB16: (CL|CP) SP* × NS
  if ((lead === "CL" || lead === "CP") && after === "NS") return false;
  // LB17: B2 SP* × B2
  if (lead === "B2" && after === "B2") return false;
  // LB18
  if (before === "SP") return true;
  // LB19 (LB19a: not among East Asian text, quoteBreaksAmongEastAsian)
  if ((before === "QU" || after === "QU") && !quoteAmongEastAsian) return false;
  // LB21
  if (after === "BA" || after === "HY" || after === "NS" || before === "BB") return false;
  // LB22
  if (after === "IN") return false;
  // LB23
  if ((before === "AL" && after === "NU") || (before === "NU" && after === "AL")) return false;
  // LB23a
  if ((before === "PR" && after === "ID") || (before === "ID" && after === "PO")) return false;
  // LB24
  if ((before === "PR" || before === "PO") && after === "AL") return false;
  if (before === "AL" && (after === "PR" || after === "PO")) return false;
  // LB25 (the pair form)
  if ((before === "CL" || before === "CP" || before === "NU") && (after === "PO" || after === "PR")) return false;
  if ((before === "PO" || before === "PR") && (after === "OP" || after === "NU")) return false;
  if ((before === "HY" || before === "IS" || before === "NU" || before === "SY") && after === "NU") return false;
  // LB28
  if (before === "AL" && after === "AL") return false;
  // LB29
  if (before === "IS" && after === "AL") return false;
  // LB30
  if ((before === "AL" || before === "NU") && after === "OP") return false;
  if (before === "CP" && (after === "AL" || after === "NU")) return false;
  // LB31
  return true;
}

export interface BreakableGrapheme {
  readonly text: string;
  readonly advance: number;
}

/** Tolerance of a fit: the browsers' layout unit is 1/64 px. */
const FIT_EPSILON = 1 / 128;

/**
 * Greedy line filling. Returns the index of the first grapheme of every line
 * (the first is 0). Spaces before a break stay on their line and take no
 * room there; a line holds at least one grapheme.
 */
export function breakLines(graphemes: readonly BreakableGrapheme[], width: number, opportunities: readonly boolean[]): number[] {
  const starts = [0];
  let lineStart = 0;
  for (;;) {
    let used = 0;
    let pendingSpace = 0;
    let lastBreak = -1;
    let overflowAt = -1;
    for (let i = lineStart; i < graphemes.length; i += 1) {
      if (i > lineStart && opportunities[i]) lastBreak = i;
      const grapheme = graphemes[i];
      if (grapheme.text === " ") {
        pendingSpace += grapheme.advance;
        continue;
      }
      if (i > lineStart && used + pendingSpace + grapheme.advance > width + FIT_EPSILON) {
        overflowAt = i;
        break;
      }
      used += pendingSpace + grapheme.advance;
      pendingSpace = 0;
    }
    if (overflowAt < 0) return starts;
    // The last opportunity on the line; a word wider than the line breaks
    // where it overflows (overflow-wrap: anywhere).
    lineStart = lastBreak > lineStart ? lastBreak : overflowAt;
    starts.push(lineStart);
  }
}
