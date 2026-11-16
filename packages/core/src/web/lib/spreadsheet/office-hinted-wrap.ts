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

function isIdeograph(ch: string): boolean {
  const code = ch.codePointAt(0) ?? 0;
  return (code >= 0x3400 && code <= 0x4dbf) || (code >= 0x4e00 && code <= 0x9fff)
    || (code >= 0xf900 && code <= 0xfaff) || (code >= 0x20000 && code <= 0x3ffff);
}

/**
 * The room a wrapped line of a `columnPx` column breaks at: 0.978 x px - 5.2,
 * not a fixed padding (lookwin p66-lw: 39 Calibri 11 columns of 40-300 px,
 * the room px - 7 up to 100 px, px - 10 at 150-160, px - 14 at 270; p65-lw's
 * 47 px column wraps at 40). Over 1,418 rows 1,398 give Excel's line count
 * (px - 8: 1,342); p20 Calibri 199/200 and Arial 198/200 (px - 8: 193, 196);
 * p35 East Asian 1,427/1,440 (1,416).
 */
export function officeHintedWrapRoom(columnPx: number): number {
  return OFFICE_HINTED_WRAP_ROOM_SCALE * columnPx - OFFICE_HINTED_WRAP_ROOM_LESS_PX;
}
export const OFFICE_HINTED_WRAP_ROOM_SCALE = 0.978;
export const OFFICE_HINTED_WRAP_ROOM_LESS_PX = 5.2;

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
  "等线": "dengxian", "宋体": "simsun", "微软雅黑": "microsoft yahei",
};

function faceKey(family: string): string {
  const key = family.normalize("NFKC").toLowerCase().split(/\s+/).filter(Boolean).join(" ");
  return ALIASES[key] ?? key;
}

/** The face's advances at its pixel em, or undefined where the table has none. */
export function advancesOf(face: OfficeHintedWrapFace): ((ch: string) => number | undefined) | undefined {
  // The face Windows Excel draws the name in (one it lacks, Helvetica Neue,
  // wraps in Arial's widths; lookwin faces probe), then the table's own name.
  const key = faceKey(excelWindowsFaceKey(face.family, face.record));
  const styles = table.faces[key];
  if (!styles || !Number.isFinite(face.sizePt) || face.sizePt <= 0) return undefined;
  const style = face.bold ? (face.italic ? "bi" : "b") : face.italic ? "i" : "r";
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
    const at = charIndex.get(ch);
    const own = at === undefined ? undefined : decode(encoded[at]);
    if (own !== undefined) return own;
    const cjkAt = cjkCharIndex.get(ch) ?? (isIdeograph(ch) ? IDEOGRAPH_INDEX : undefined);
    if (cjkAt === undefined) return undefined;
    for (const cjk of cjkTables) {
      const value = cjk ? decode(cjk[cjkAt]) : undefined;
      if (value !== undefined) return value;
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
export function officeHintedWrapLines(text: string, face: OfficeHintedWrapFace, columnPx: number): string[] | undefined {
  const advance = advancesOf(face);
  if (!advance) return undefined;
  return wrapWith(text, (_at, ch) => advance(ch), columnPx);
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
export function officeHintedWrapRunLines(runs: readonly OfficeHintedWrapRun[], columnPx: number): string[] | undefined {
  let text = "";
  const owners: ((ch: string) => number | undefined)[] = [];
  for (const run of runs) {
    const piece = String(run.text).replace(/\r\n?/g, "\n");
    if (!piece) continue;
    const advance = advancesOf(run);
    if (!advance) return undefined;
    text += piece;
    owners.push(...Array.from(piece, () => advance));
  }
  if (!owners.length) return undefined;
  return wrapWith(text, (at, ch) => owners[at]?.(ch), columnPx);
}

/** The rule over `text`, `advanceAt(index, character)` the advance of the
 * character at that code point index. */
function wrapWith(text: string, advanceAt: (at: number, ch: string) => number | undefined, columnPx: number): string[] | undefined {
  if (!Number.isFinite(columnPx)) return undefined;
  const room = officeHintedWrapRoom(columnPx);
  const lines: string[] = [];
  let offset = 0;
  for (const paragraph of text.replace(/\r\n?/g, "\n").split("\n")) {
    let line = "";
    let width = 0;
    let pendingSpaces = "";
    let pendingWidth = 0;
    let first = true;
    let at = offset;
    for (const token of tokens(paragraph)) {
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
      if (!first && width + pendingWidth + tokenWidth <= room + 1e-9) {
        line += pendingSpaces + token;
        width += pendingWidth + tokenWidth;
        pendingSpaces = "";
        pendingWidth = 0;
        continue;
      }
      if (!first) lines.push(line);
      first = false;
      pendingSpaces = "";
      pendingWidth = 0;
      line = "";
      width = 0;
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
