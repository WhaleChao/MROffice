/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Windows Excel's screen wrap of a cell's text: the lines GDI breaks it
 * into, from its whole-pixel hinted advance widths at the cell's pixel em
 * (office-wrap-advances.json, built by xlsx-rounds/lookwin/tools/
 * build_wrap_advances.py: FreeType's v35 TrueType interpreter reproduces
 * GDI's advances; numbers only, no font data). The room is
 * officeHintedWrapRoom of the column's pixels; a line breaks after a space or a hyphen, and a word wider
 * than the room breaks between its letters. Measured against Windows Excel's
 * own rows (lookwin probes p20/p20b): 96.5-99% of rows' line counts across
 * Calibri, Arial, Aptos Narrow, Times New Roman, Verdana and Tahoma, where
 * the browser's widths reached about 89%.
 *
 * East Asian text wraps by the same rule (lookwin probes p35/p36: 98.4-100%
 * of rows across MS Gothic, MS PGothic, Yu Gothic, DengXian, SimSun,
 * Microsoft YaHei and Calibri cells): a run of ideographs and kana is one
 * "word" that breaks between any two characters when it does not fit, and
 * Excel keeps no line-start or line-end rule (kinsoku). A Western face takes
 * the characters it lacks from Yu Gothic, as Windows Excel draws them in a
 * Calibri cell; every ideograph of a face is one advance.
 *
 * Undefined for a face, style, size or character the table does not hold:
 * the caller keeps its own measure. The API's
 * office_hinted_wrap.py is the same rule on an identical table, and
 * office-hinted-wrap.cases.json holds the answers both sides must give.
 */
import TABLE from "./office-wrap-advances.json";
import { excelWindowsFaceKey, type ExcelFontRecordClass } from "../office-fonts/excel-windows-faces";
import { officeArabicWrapAdvances } from "./office-arabic";
import { officeRegionalLetterAdvance } from "./office-regional-letters";

type FaceTable = Readonly<Record<string, Readonly<Record<string, Readonly<Record<string, string>>>>>>;

interface WrapAdvanceTable {
  readonly alphabet: string;
  readonly chars: string;
  readonly faces: FaceTable;
  /** East Asian advances: `chars`, then one symbol for every ideograph. */
  readonly cjk: { readonly chars: string; readonly fallback: string; readonly faces: FaceTable };
}

const table = TABLE as unknown as WrapAdvanceTable;
const UNKNOWN = table.alphabet.length - 1;
const charIndex = new Map<string, number>([...table.chars].map((ch, index) => [ch, index]));
const cjkCharIndex = new Map<string, number>([...table.cjk.chars].map((ch, index) => [ch, index]));
const IDEOGRAPH_INDEX = [...table.cjk.chars].length;
const alphabetIndex = new Map<string, number>([...table.alphabet].map((ch, index) => [ch, index]));

/**
 * The characters past Latin-1 whose widths office-wrap-advances-extra.json
 * holds (lookwin p257-lw: Latin Extended-A U+0100-013F, the Greek letters and
 * − ≤ ≥ ≠ ≈ ∞ wrap at FreeType's hinted widths as the Latin-1 table does). The
 * table is a chunk of its own, fetched the first time a cell's text holds one
 * of them: until it lands such text keeps the caller's own measure, and the
 * sheet is measured again when it has (officeWrapExtraEpoch). The API's
 * office_wrap_advances_extra.json is the same table, read whole.
 */
const EXTRA_CHARS = /[Ā-ĿΑ-ΡΣ-Ωα-ω−≤≥≠≈∞]/u;

interface ExtraWrapAdvanceTable {
  readonly chars: string;
  readonly faces: FaceTable;
}

let extraTable: { readonly index: ReadonlyMap<string, number>; readonly faces: FaceTable } | null = null;
let extraLoad: Promise<void> | null = null;
let extraEpoch = 0;
const extraListeners = new Set<() => void>();

/** Moves once the extra table has landed: a key part for anything built on hinted wraps. */
export function officeWrapExtraEpoch(): number {
  return extraEpoch;
}

/** Calls `listener` when the extra table lands; returns the unsubscribe. */
export function onOfficeWrapExtraLoaded(listener: () => void): () => void {
  extraListeners.add(listener);
  return () => { extraListeners.delete(listener); };
}

/** Fetches the extra table once (a failed fetch leaves those characters to the caller's measure). */
export function loadOfficeWrapExtra(): Promise<void> {
  extraLoad ??= import("./office-wrap-advances-extra.json").then((module) => {
    const data = ((module as { default?: unknown }).default ?? module) as ExtraWrapAdvanceTable;
    extraTable = { index: new Map([...data.chars].map((ch, index) => [ch, index])), faces: data.faces };
    extraEpoch += 1;
    for (const listener of [...extraListeners]) listener();
  }, () => undefined);
  return extraLoad;
}

function isIdeograph(ch: string): boolean {
  const code = ch.codePointAt(0) ?? 0;
  // ... or a Hangul syllable or compatibility jamo: every Korean face of
  // Windows sets one a whole em, as it sets an ideograph (Malgun Gothic's
  // hinted advances; printed, lookwin p187-lw).
  return (code >= 0x3400 && code <= 0x4dbf) || (code >= 0x4e00 && code <= 0x9fff)
    || (code >= 0xf900 && code <= 0xfaff) || (code >= 0x20000 && code <= 0x3ffff)
    || (code >= 0xac00 && code <= 0xd7a3) || (code >= 0x3131 && code <= 0x318e);
}

export const OFFICE_HINTED_WRAP_ROOM_STEP_PX = 29;
export const OFFICE_HINTED_WRAP_ROOM_STEP_FROM_PX = 56;

/** The workbook-wide inputs of Windows Excel's screen wrap beyond a cell's own face. */
export interface OfficeHintedWrapOptions {
  /** The workbook's Normal font's widest digit, px (its max digit width, Excel's column unit): sets the room's
   * steps (officeHintedWrapRoom). */
  readonly normalDigitPx?: number;
  /** Excel's wrap after it has fitted a wrapped complex-script cell in this load (officeComplexWrapTrigger): a
   * line broken at a space keeps room for the space after it. */
  readonly spaced?: boolean;
}

/**
 * The room a wrapped line of a `columnPx` column breaks at, for a face of
 * `emPx` whole pixels per em: the column less a margin of 2 p + 1 px,
 * p = (em - 1) // 8 + 2 (7 px for 9-16 px ems, 8-12 pt; 9 px for 17-24,
 * 13-18 pt; 5 px at 6 pt), then a pixel less again every 29 px of that past
 * 56 (lookwin p110b-lw: Windows Excel's own row heights on resave, 99 column
 * widths of 20-600 px at Calibri 11; p110/p110c/p110d-lw: Calibri, Arial and
 * Times New Roman at 6-18 pt alike, 260 of 269 widths x sizes, the misses at
 * 19 pt and over). A line fits while its hinted width is at most the room.
 * The linear 0.978 x px - 5.2 before it gave a 300 px column 288; Excel's is
 * 285. The API's office_hinted_wrap_room is the same rule.
 *
 * `normalDigitPx`, the workbook's Normal font's widest digit d, sets the
 * steps: a pixel less every 4 d + 1 px past 8 d (29 past 56 at 7 px; lookwin
 * p116/p116b-lw: Arial 8 and Calibri 8, 6 px, every 25 past 48; Times New
 * Roman 12 and Verdana 10, 8 px, every 33 past 64; 0 misses over 146 widths).
 * `digitPx`, the cell face's own widest digit, sets the margin:
 * p = (d - 1) // 4 + 2 (Arial 12, 9 px: 9 px where its 16 px em gave 7).
 */
export function officeHintedWrapRoom(columnPx: number, emPx = 15, normalDigitPx?: number, digitPx?: number): number {
  const margin = digitPx !== undefined && Number.isInteger(digitPx) && digitPx > 0
    ? 2 * (Math.floor((digitPx - 1) / 4) + 2) + 1
    : 2 * (Math.floor((Math.max(1, Math.trunc(emPx)) - 1) / 8) + 2) + 1;
  const width = columnPx - margin;
  const stepped = normalDigitPx !== undefined && Number.isInteger(normalDigitPx) && normalDigitPx > 0;
  const start = stepped ? 8 * normalDigitPx : OFFICE_HINTED_WRAP_ROOM_STEP_FROM_PX;
  const step = stepped ? 4 * normalDigitPx + 1 : OFFICE_HINTED_WRAP_ROOM_STEP_PX;
  return width - Math.max(0, Math.floor((width - start) / step));
}

/** The widest digit of a face's regular cut at the whole-pixel ems under the
 * hinted table's (3-7 px), FreeType's v35 interpreter as the table: lookwin
 * p244-lw, Normal Arial 10 at 50% steps its room by a 4 px digit; p239-lw,
 * Calibri 11 at 40% by a 3 px one. The API's _SMALL_EM_DIGIT_PX. */
const SMALL_EM_DIGIT_PX: Readonly<Record<string, readonly number[]>> = {
  // ems 3, 4, 5, 6, 7
  calibri: [2, 2, 3, 3, 4], arial: [2, 2, 3, 3, 4], "times new roman": [2, 2, 3, 3, 4],
  verdana: [2, 3, 3, 4, 4], tahoma: [2, 2, 3, 3, 4], "aptos narrow": [2, 2, 3, 3, 4],
};

/** The widest zoom the small-em digits were measured at (p244-lw 50%,
 * p239-lw 40%). The API's OFFICE_SMALL_EM_DIGIT_MAX_ZOOM. */
export const OFFICE_SMALL_EM_DIGIT_MAX_ZOOM = 55;

/** A zoomed-out Normal font's widest digit where its em is under the hinted
 * table's (3-7 px), regular cut only; undefined elsewhere. Asked only for a
 * sheet at OFFICE_SMALL_EM_DIGIT_MAX_ZOOM or less (the API's
 * office_small_em_digit_px). */
export function officeSmallEmDigitPx(face: OfficeHintedWrapFace): number | undefined {
  if (face.bold || face.italic || !Number.isFinite(face.sizePt) || face.sizePt <= 0) return undefined;
  const small = SMALL_EM_DIGIT_PX[faceKey(excelWindowsFaceKey(face.family, face.record))];
  const ppem = Math.floor(face.sizePt * 96 / 72 + 0.5);
  return small && ppem >= 3 && ppem <= 7 ? small[ppem - 3] : undefined;
}

/** The widest of a face's ten digits at its whole-pixel em (Excel's max digit width), or undefined (the API's
 * office_digit_px). */
export function officeDigitPx(face: OfficeHintedWrapFace): number | undefined {
  const advance = advancesOf(face);
  if (!advance) return undefined;
  let widest = 0;
  for (const digit of "0123456789") {
    const value = advance(digit);
    if (value === undefined) return undefined;
    widest = Math.max(widest, value);
  }
  return widest;
}

/**
 * Letters of the complex scripts Windows Excel shapes through Uniscribe
 * (Hebrew, Arabic, Syriac, Thaana, the Indic scripts, Thai, Lao, Tibetan,
 * Myanmar, Khmer, the Hebrew/Arabic presentation forms): the API's
 * _COMPLEX_WRAP_TEXT (office-hinted-wrap.cases.json complexCases holds the
 * answers both give).
 */
export const OFFICE_COMPLEX_WRAP_TEXT = /[\u0590-\u08ff\u0900-\u0dff\u0e00-\u0eff\u0f00-\u0fff\u1000-\u109f\u1780-\u17ff\ufb1d-\ufdff\ufe70-\ufeff]/u;
/** The text whose wrapped cell turns Windows Excel's later wraps of the load
 * "spaced" (office-complex-wrap): the complex scripts' letters, and a
 * combining diacritical mark, which Uniscribe shapes too (lookwin p217-lw: a
 * wrapped "o" + U+0301 does, as a Hebrew alef; a zero width space, a tab and
 * curly quotes do not). The API's _SPACED_WRAP_TRIGGER_TEXT. */
export const OFFICE_SPACED_WRAP_TRIGGER_TEXT = /[\u0300-\u036f\u0590-\u08ff\u0900-\u0dff\u0e00-\u0eff\u0f00-\u0fff\u1000-\u109f\u1780-\u17ff\ufb1d-\ufdff\ufe70-\ufeff]/u;

/** GDI's whole pixels per em of a `sizePt` face at 96 dpi (the advances table's key). */
function emPxOf(sizePt: number): number {
  return Math.floor((sizePt * 96) / 72 + 0.5);
}

export interface OfficeHintedWrapFace {
  readonly family: string;
  readonly sizePt: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** The font record's class where Windows lacks its face (excelFontRecordOf). */
  readonly record?: ExcelFontRecordClass | null;
}

/** The metric-compatible faces a page draws an Office face in (the web's
 * substitutes) wrap as that face does, and the East Asian faces' own names;
 * keyed by the NFKC lower-case name (the API's _ALIASES). */
const ALIASES: Readonly<Record<string, string>> = {
  carlito: "calibri", arimo: "arial", "liberation sans": "arial", tinos: "times new roman", "liberation serif": "times new roman",
  "ms ゴシック": "ms gothic", "ms pゴシック": "ms pgothic", "游ゴシック": "yu gothic", yugothic: "yu gothic",
  "等线": "dengxian", "宋体": "simsun", "微软雅黑": "microsoft yahei", "맑은 고딕": "malgun gothic",
  "ms 明朝": "ms mincho", "ms p明朝": "ms pmincho", "メイリオ": "meiryo", "游明朝": "yu mincho", yumincho: "yu mincho",
};

function faceKey(family: string): string {
  const key = family.normalize("NFKC").toLowerCase().split(/\s+/).filter(Boolean).join(" ");
  return ALIASES[key] ?? key;
}

/**
 * Office's fixed-pitch Latin faces: every glyph's advance in 2048ths of the em
 * (Courier New 1229, Consolas 1126, Lucida Console 1234). GDI sets each at its
 * whole-pixel em, rounded: Courier New 10 is 8 px at 13 px per em. Their lines
 * are GDI's own, so their rows are fitted, not kept at a stated height
 * (R379-X18's one-line Courier New 10 names, 18 px). The API's
 * _FIXED_PITCH_UNITS.
 */
const FIXED_PITCH_UNITS: Readonly<Record<string, number>> = { "courier new": 1229, consolas: 1126, "lucida console": 1234 };
/** The fixed-pitch faces Office ships without a bold cut: GDI emboldens each glyph a pixel wider. */
const FIXED_PITCH_NO_BOLD = new Set(["lucida console"]);

/** The face's advances at its pixel em, or undefined where the table has none. */
/** The characters Windows Excel's wrap gives no width: a combining
 * diacritical mark (it stands over its base), a zero width space and a tab
 * (lookwin p218-lw: an Arial 10 line of 1117 px keeps its width with one or
 * three of each inside its words; a tab between two words in place of their
 * space takes the space's 4 px off). The API's _zero_advance. */
function zeroAdvance(ch: string): boolean {
  return ch === "\t" || ch === "\u200b" || (ch >= "\u0300" && ch <= "\u036f");
}

export function advancesOf(face: OfficeHintedWrapFace): ((ch: string) => number | undefined) | undefined {
  // The face Windows Excel draws the name in (one it lacks, Helvetica Neue,
  // wraps in Arial's widths; lookwin faces probe), then the table's own name.
  const key = faceKey(excelWindowsFaceKey(face.family, face.record));
  const styles = table.faces[key];
  if (!Number.isFinite(face.sizePt) || face.sizePt <= 0) return undefined;
  const fixedPitch = FIXED_PITCH_UNITS[key];
  if (!styles && fixedPitch !== undefined) {
    const width = Math.floor((fixedPitch * Math.floor((face.sizePt * 96) / 72 + 0.5)) / 2048 + 0.5)
      + (face.bold && FIXED_PITCH_NO_BOLD.has(key) ? 1 : 0);
    // The table's Latin characters only: a face's other scripts fall back.
    return (ch) => (charIndex.has(ch) ? width : undefined);
  }
  if (!styles) return undefined;
  let style = face.bold ? (face.italic ? "bi" : "b") : face.italic ? "i" : "r";
  // A face Office ships without a bold cut (ＭＳ Ｐゴシック, ＭＳ ゴシック,
  // SimSun) is emboldened by GDI: its regular advances, a pixel wider each
  // (lookwin p166-lw / R361-X03 A131: a bold ＭＳ Ｐゴシック 11 label wraps five
  // 16 px kana a line where regular 15 px ones made six). The API's _advances.
  let widen = 0;
  if (face.bold && !styles[style] && styles.r) {
    style = "r";
    widen = 1;
  }
  const ppem = String(Math.floor((face.sizePt * 96) / 72 + 0.5));
  const encoded = styles[style]?.[ppem];
  if (!encoded) return undefined;
  // The face's own East Asian advances, then the fallback face's (a character the face lacks).
  const cjkTables = [...new Set([key, table.cjk.fallback])]
    .filter((name) => table.cjk.faces[name])
    .map((name) => table.cjk.faces[name][style]?.[ppem] ?? "");
  const decode = (symbol: string | undefined) => {
    const value = symbol === undefined ? undefined : alphabetIndex.get(symbol);
    return value === undefined || value === UNKNOWN ? undefined : value;
  };
  return (ch) => {
    if (zeroAdvance(ch)) return 0;
    // Segoe UI Emoji's letter, whatever the face (office-regional-letters.ts, p254-lw).
    const regional = officeRegionalLetterAdvance(ch);
    if (regional !== undefined) return Math.floor(regional * Number(ppem) + 0.5);
    const at = charIndex.get(ch);
    const own = at === undefined ? undefined : decode(encoded[at]);
    if (own !== undefined) return own + widen;
    if (at === undefined && EXTRA_CHARS.test(ch)) {
      // Greek, Latin Extended-A and the math signs from their own chunk, once it has landed.
      if (!extraTable) {
        void loadOfficeWrapExtra();
        return undefined;
      }
      const extraAt = extraTable.index.get(ch);
      const extra = extraAt === undefined ? undefined : decode(extraTable.faces[key]?.[style]?.[ppem]?.[extraAt]);
      return extra === undefined ? undefined : extra + widen;
    }
    const cjkAt = cjkCharIndex.get(ch) ?? (isIdeograph(ch) ? IDEOGRAPH_INDEX : undefined);
    if (cjkAt === undefined) return undefined;
    for (const cjk of cjkTables) {
      const value = cjk ? decode(cjk[cjkAt]) : undefined;
      if (value !== undefined) return value + widen;
    }
    return undefined;
  };
}

/** The tokens a line may break after: a word up to and including a hyphen,
 * a word, or one space. */
function tokens(paragraph: string): string[] {
  return paragraph.match(/[^ -]*-|[^ ]+| /g) ?? [];
}

/**
 * The lines Windows Excel's screen wraps `text` into in a column `columnPx`
 * wide, or undefined when the table cannot say (the caller measures as it
 * did). A typed line break starts a new line; an empty paragraph is a line.
 */
export function officeHintedWrapLines(text: string, face: OfficeHintedWrapFace, columnPx: number, options: OfficeHintedWrapOptions = {}): string[] | undefined {
  const advance = advancesOf(face);
  if (!advance) return undefined;
  if (OFFICE_COMPLEX_WRAP_TEXT.test(text)) {
    // A Calibri cell's Arabic at Windows Calibri's own joined, kerned widths
    // (officeArabicWrapAdvances, lookwin p220-lw), in Excel's room of the
    // column scaled to the face's size (complexWrapWidth; the API's
    // cell_text_lines); other complex text the table cannot say for.
    const shares = officeArabicWrapAdvances(text, face.family, face.bold === true, face.sizePt);
    if (!shares) return undefined;
    const scale = complexRoomScale(face.sizePt);
    return wrapWith(text, (at) => (shares[at] ?? 0) * scale, columnPx, emPxOf(face.sizePt), options, officeDigitPx(face));
  }
  return wrapWith(text, (_at, ch) => advance(ch), columnPx, emPxOf(face.sizePt), options, officeDigitPx(face),
    face.italic === true ? officeItalicWrapLess(emPxOf(face.sizePt)) : 0);
}

/** The pixels an italic cell's wrap room is short of its upright one's,
 * about a fifth of its em (lookwin p250/p251-lw: 2 px up to a 15 px em, 3 at
 * 16-19, 4 at 21-24, 5 at 27). The API's office_italic_wrap_less. */
export function officeItalicWrapLess(emPx: number): number {
  return Math.max(0, Math.floor(Math.trunc(emPx) * 4 / 21));
}

/** Complex-script text wraps in Excel's room of the column at the face's
 * whole-pixel em scaled to its size (lookwin p119-lw): its advances over that
 * scale, against the room at the em. */
function complexRoomScale(sizePt: number): number {
  return emPxOf(sizePt) / (sizePt * 96 / 72);
}

/** One rich-text run: its text and the face it is set in. */
export interface OfficeHintedWrapRun extends OfficeHintedWrapFace {
  readonly text: string;
}

/**
 * officeHintedWrapLines over rich text: each character at its own run's
 * hinted advance (the server's office_hinted_wrap_runs; R195-X01's footnote
 * marks, R226-X05's notes); undefined when the table cannot say for any run.
 */
export function officeHintedWrapRunLines(runs: readonly OfficeHintedWrapRun[], columnPx: number, options: OfficeHintedWrapOptions = {}): string[] | undefined {
  let text = "";
  const owners: ((ch: string) => number | undefined)[] = [];
  let emPx: number | undefined;
  let digitPx: number | undefined;
  let firstRun = true;
  // A run in italic anywhere takes the italic cell's room off the cell, at
  // its em (lookwin p259-lw; the server's office_hinted_wrap_runs).
  let slanted = false;
  const complex = runs.some((run) => OFFICE_COMPLEX_WRAP_TEXT.test(String(run.text)));
  const scale = complex ? complexRoomScale(runs.find((run) => String(run.text))?.sizePt ?? 11) : 1;
  for (const run of runs) {
    const piece = String(run.text).replace(/\r\n?/g, "\n");
    if (!piece) continue;
    const advance = advancesOf(run);
    if (!advance) return undefined;
    text += piece;
    if (run.italic === true) slanted = true;
    if (complex && OFFICE_COMPLEX_WRAP_TEXT.test(piece)) {
      // Each run's words at its face's Arabic widths (officeHintedWrapLines).
      const shares = officeArabicWrapAdvances(piece, run.family, run.bold === true, run.sizePt);
      if (!shares || /[\ud800-\udfff]/.test(piece)) return undefined;
      owners.push(...shares.map((share) => () => share * scale));
    } else if (complex) {
      owners.push(...Array.from(piece, () => (ch: string) => {
        const value = advance(ch);
        return value === undefined ? undefined : value * scale;
      }));
    } else owners.push(...Array.from(piece, () => advance));
    // The cell's margin is its first run's face's.
    emPx ??= emPxOf(run.sizePt);
    if (firstRun) {
      digitPx = officeDigitPx(run);
      firstRun = false;
    }
  }
  if (!owners.length) return undefined;
  return wrapWith(text, (at, ch) => owners[at]?.(ch), columnPx, emPx ?? 15, options, digitPx,
    slanted ? officeItalicWrapLess(emPx ?? 15) : 0);
}

/** The rule over `text`, `advanceAt(index, character)` the advance of the
 * character at that code point index.
 *
 * `spaced`: once Windows Excel has wrapped a cell of complex-script text in a
 * row it fits, every later wrap of the load keeps room for one space after a
 * line it breaks at a space: the line and that space must fit (lookwin
 * p115b-lw: Arial 10's first line of "A B tail" kept at room - 4; a hyphen
 * break and a last word keep the plain room, a paragraph's trailing spaces
 * count as one; p117b/p117d/p118-lw; R379-X18's formula notes wrap 9 and 19
 * lines, not 8 and 18). The API's _wrap is the same rule. */
function wrapWith(text: string, advanceAt: (at: number, ch: string) => number | undefined, columnPx: number, emPx = 15,
  options: OfficeHintedWrapOptions = {}, digitPx?: number, roomLess = 0): string[] | undefined {
  if (!Number.isFinite(columnPx)) return undefined;
  const room = officeHintedWrapRoom(columnPx, emPx, options.normalDigitPx, digitPx) - roomLess;
  const spaced = options.spaced === true;
  const lines: string[] = [];
  let offset = 0;
  const paragraphs = text.replace(/\r\n?/g, "\n").split("\n");
  // The break a text ends with opens a last, empty line -- but not after a
  // line of spaces alone, nor, in the spaced wrap, after a line that ends in
  // a space (lookwin p212-lw: "abc\n" and "abc \n" stand two lines,
  // "abc\n \n" two; p217-lw, after a wrapped complex-script cell: "abc \n"
  // one line, "abc \n\n" three still; R454-X16's notes ending ". \n").
  const beforeLast = paragraphs.length >= 2 && paragraphs[paragraphs.length - 1] === "" ? paragraphs[paragraphs.length - 2] : undefined;
  if (beforeLast !== undefined && beforeLast.endsWith(" ") && (spaced || beforeLast.replace(/ /g, "") === "")) paragraphs.pop();
  for (const paragraph of paragraphs) {
    let line = "";
    let width = 0;
    let pendingSpaces = "";
    let pendingWidth = 0;
    let first = true;
    let at = offset;
    const all = tokens(paragraph);
    for (const [tokenIndex, token] of all.entries()) {
      const tokenAt = at;
      const chars = [...token];
      at += chars.length;
      if (token === " ") {
        const space = advanceAt(tokenAt, " ");
        if (space === undefined) return undefined;
        pendingSpaces += " ";
        pendingWidth += space;
        continue;
      }
      const widths: number[] = [];
      for (let index = 0; index < chars.length; index += 1) {
        const value = advanceAt(tokenAt + index, chars[index]);
        if (value === undefined) return undefined;
        widths.push(value);
      }
      const tokenWidth = widths.reduce((sum, value) => sum + value, 0);
      let follow = 0;
      if (spaced && !token.endsWith("-") && all[tokenIndex + 1] === " ") follow = advanceAt(at, " ") ?? 0;
      if (!first && width + pendingWidth + tokenWidth + follow <= room + 1e-9) {
        line += pendingSpaces + token;
        width += pendingWidth + tokenWidth;
        pendingSpaces = "";
        pendingWidth = 0;
        continue;
      }
      // The spaces a paragraph starts with stand on its first line at their
      // advances (lookwin p216-lw: one, two and three leading spaces of Arial
      // 10 take 4, 8 and 12 px of the room; the spaces a line ends with take
      // none); the spaces a wrap breaks at start no line.
      if (first) {
        line = pendingSpaces;
        width = pendingWidth;
      } else {
        lines.push(line);
        line = "";
        width = 0;
      }
      first = false;
      pendingSpaces = "";
      pendingWidth = 0;
      // ...alone, where the first word does not fit after them: it starts
      // the next line (R454-X16's U58, a space before a 17-digit coordinate
      // wider than its 127 px column: five lines, 25 in all, not 24).
      if (line && width + tokenWidth > room + 1e-9) {
        lines.push(line);
        line = "";
        width = 0;
      }
      // A token wider than the room breaks between its letters.
      for (let index = 0; index < chars.length; index += 1) {
        const value = widths[index];
        if (width + value > room + 1e-9 && width > 0) {
          lines.push(line);
          line = "";
          width = 0;
        }
        line += chars[index];
        width += value;
      }
    }
    lines.push(line);
    offset += [...paragraph].length + 1;
  }
  return lines;
}
