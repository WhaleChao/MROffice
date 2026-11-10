/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Spreadsheet Find & Replace (Ctrl/Cmd+H — Excel parity)
//
// This module is PURE: no React, no DOM. Given a sheet's cell map, a query and
// a set of options it produces the set of matching cell keys and a replaceAll
// that returns the updated cell map plus a replacement count.
//
// What is searched:
//   - The DISPLAYED text of each cell (literal text, or the computed result of
//     a formula, with any number format applied) — this mirrors what the user
//     sees in the grid.
//   - When `inFormulas` is set, the RAW formula text (e.g. "=A1+B1") is searched
//     instead, so a query can target formula source.
//
// What is replaced (replaceAll):
//   - With `inFormulas`, the raw formula text is rewritten and the cell's stored
//     `value`/`formula` updated. Display values follow from re-evaluation.
//   - Without `inFormulas`, only LITERAL (non-formula) cell values are rewritten
//     — replacing the displayed result of a formula would be meaningless (the
//     formula recomputes it), so formula cells are skipped during replacement.
//
// Replacement is intentionally substring-based (Excel "Find & Replace" replaces
// every occurrence within a cell), unless `wholeCell` is set, in which case the
// whole cell text must equal the query and is swapped for the replacement.
//
// Excel's "Options >>" (Microsoft Support, "Find or replace text and numbers on
// a worksheet") are modelled by `lookIn`, `searchOrder` and `matchByte`:
//   - Look in: Formulas (a formula's source, a constant as the formula bar
//     shows it — never the formatted text), Values (what the cell displays),
//     Notes, Comments. The Replace tab only offers Formulas.
//     A constant's formula-bar text is its EDIT text (entry-parse): Excel reads
//     a typed 50% into 0.5 with the format 0%, shows 50% again in the formula
//     bar, and Find matches that "50%", not the stored 0.5; a date matches its
//     regional short date, not its serial. Replace rewrites that same text and
//     enters it again, so "50%" -> "60%" is stored as 0.6, still a percentage.
//   - Search: By Rows / By Columns decides the order Find Next and Find All
//     walk the hits in.
//   - Match byte (区分全/半角, East Asian editions): off by default, when full-
//     and half-width forms of the same character are equal.
// ---------------------------------------------------------------------------

import { SOURCE_XF_STYLE_KEY } from "./style-source-xf";
import { recordFontDrawn } from "@mrnote/core/web/lib/spreadsheet/cell-record-font.ts";
import {
  columnToLetters,
  formatCellValue,
  lettersToColumn,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { storedFormulaResult } from "@mrnote/core/web/lib/spreadsheet/stored-formula-results.ts";
import {
  applyNumberFormat,
  normalizeNumberFormatCode,
} from "@mrnote/core/web/lib/spreadsheet/number-format.ts";
import {
  parseSpreadsheetEntry,
  spreadsheetEntryEditText,
} from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import { phoneticAfterValueWrite } from "@mrnote/core/web/lib/spreadsheet/cell-phonetic.ts";
import { SPREADSHEET_DEFAULT_FONT_PT, spreadsheetPtToPx } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";

/** Minimal structural shape of a spreadsheet cell this module relies on. */
export interface FindReplaceCell {
  value?: string;
  formula?: string;
  /** The editor cell's entry kind; a Look in Formulas rewrite re-enters it. */
  kind?: "text" | "number";
  style?: {
    numberFormat?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

export type FindReplaceCellMap = Record<string, FindReplaceCell>;

/**
 * What a cell displays, answered by the host's authoritative calculation (the
 * grid's own display: workbook-wide formulas, spills, number formats). Look
 * in Values matches this text, so Find sees exactly what the grid shows,
 * including a result that reads another sheet (=Second!A1).
 */
export type FindDisplayText = (
  row: number,
  col: number,
  cell: FindReplaceCell,
) => string;

/** Excel's "Look in" list, in Excel's order. */
export type FindLookIn = "formulas" | "values" | "notes" | "comments";
/** Excel's "Search" list: By Rows or By Columns. */
export type FindSearchOrder = "rows" | "columns";

export interface FindReplaceOptions {
  /** Match case exactly. Default false (case-insensitive). */
  caseSensitive?: boolean;
  /** The whole cell text must equal the query (Excel "Match entire cell"). */
  wholeCell?: boolean;
  /**
   * Legacy switch: search/replace only the formula source of formula cells
   * (constants never match). Superseded by `lookIn`, which wins when set.
   */
  inFormulas?: boolean;
  /** Excel "Look in". Unset keeps the legacy displayed-value search. */
  lookIn?: FindLookIn;
  /**
   * Excel "Match byte" (区分全/半角). Unset/false: a full-width letter, digit,
   * sign or space equals its half-width form, as in Excel by default.
   */
  matchByte?: boolean;
  /** Excel "Search": order of the hits. Default "rows". */
  searchOrder?: FindSearchOrder;
  /**
   * The UI locale, which decides the region a constant's edit text is written
   * in and read back with (zh: 2026/9/21 and ¥; otherwise en-US). Look in
   * Formulas only. Default "en-US".
   */
  locale?: string;
  /**
   * The displayed text of a cell of the searched sheet (Look in Values). A
   * host with a live workbook supplies it so Find never rebuilds a separate,
   * single-sheet calculation; without it a pure caller falls back to
   * evaluating the cell map alone.
   */
  displayText?: FindDisplayText;
  /**
   * Cell keys ("r:c") that show a value without a stored cell: a dynamic
   * array's spilled results. Look in Values visits them through `displayText`
   * (they hold no formula or note of their own), so a spilled "Audit126_B" is
   * found like a typed one (live audit B126).
   */
  extraKeys?: Iterable<string>;
  /**
   * Excel's Find "Format…" (XLS-A103): only a cell whose format is this one
   * is a hit. With it set, an empty "Find what" finds every cell of that
   * format, as in Excel. The format is a cell style snapshot (Choose Format
   * From Cell takes the picked cell's whole style).
   */
  format?: FindCellFormat;
  /**
   * Excel's Replace "Format…": every replaced cell takes this format on top
   * of its own. With it set, an empty "Find what" and "Replace with" change
   * the format of the found cells and leave their contents alone.
   */
  replaceFormat?: FindCellFormat;
}

/** A Find / Replace format: the cell style keys it pins. */
export type FindCellFormat = Readonly<Record<string, unknown>>;

/**
 * Marks a format built in the Find Format dialog (XLS-A103): only the
 * attributes it names are compared (Excel's unset fields match anything). A
 * format picked from a cell has no mark and is the cell's whole format.
 */
export const FIND_FORMAT_PARTIAL = "@findFormatPartial";

/** A style value as Find Format compares it: unset, false and "" are alike. */
function formatValueKey(value: unknown): string {
  if (value === undefined || value === null || value === false || value === "") {
    return "";
  }
  return JSON.stringify(value);
}

/**
 * The value a style key has in effect, as Find Format compares it. A cell
 * that sets no size is at the default size, so it matches a format naming
 * 11 pt (B526: every plain Beijing was missed by an 11 pt Find Format), and
 * a size is compared in hundredths of a pixel, past the point/pixel
 * conversion's float noise. Other keys compare as stored.
 */
function effectiveFormatValueKey(style: Record<string, unknown>, key: string): string {
  const value = style[key];
  if (key !== "fontSize") return formatValueKey(value);
  const px =
    typeof value === "number" && Number.isFinite(value) && value > 0
      ? value
      : spreadsheetPtToPx(SPREADSHEET_DEFAULT_FONT_PT);
  return JSON.stringify(Math.round(px * 100) / 100);
}

/**
 * Whether a cell has the Find format: every style key of either side agrees,
 * so a format picked from a plain cell finds plain cells and one picked from
 * a bold red cell finds bold red cells only (XLS-A103).
 */
export function cellHasFindFormat(
  cell: FindReplaceCell | undefined,
  format: FindCellFormat,
): boolean {
  const style = (cell?.style ?? {}) as Record<string, unknown>;
  const partial = format[FIND_FORMAT_PARTIAL] === true;
  const keys = new Set(partial ? Object.keys(format) : [...Object.keys(format), ...Object.keys(style)]);
  keys.delete(FIND_FORMAT_PARTIAL);
  // An entry detail, not a format.
  keys.delete("quotePrefix");
  keys.delete(SOURCE_XF_STYLE_KEY);
  for (const key of keys) {
    if (effectiveFormatValueKey(style, key) !== effectiveFormatValueKey(format as Record<string, unknown>, key)) return false;
  }
  return true;
}

/** Whether a Find with this query and options can match anything at all. */
export function findCanMatch(query: string, options: FindReplaceOptions): boolean {
  return query !== "" || options.format !== undefined;
}

/** The cell with the Replace format laid over its own (XLS-A103). */
function withReplaceFormat(
  cell: FindReplaceCell,
  format: FindCellFormat | undefined,
): FindReplaceCell {
  if (!format) return cell;
  const style: Record<string, unknown> = { ...(cell.style ?? {}) };
  for (const [key, value] of Object.entries(format)) {
    if (key === "quotePrefix" || key === FIND_FORMAT_PARTIAL) continue;
    if (formatValueKey(value) === "") delete style[key];
    else style[key] = value;
  }
  const next: FindReplaceCell = { ...cell };
  if (Object.keys(style).length > 0) next.style = style as FindReplaceCell["style"];
  else delete next.style;
  return next;
}

export interface ReplaceAllResult {
  /** New cell map with replacements applied (input is never mutated). */
  cells: FindReplaceCellMap;
  /** Number of cells whose stored text was changed. */
  count: number;
}

const CELL_KEY_RE = /^(\d+):(\d+)$/;

function parseCellKey(key: string): { row: number; col: number } | null {
  const m = CELL_KEY_RE.exec(key);
  if (!m) return null;
  return { row: Number(m[1]), col: Number(m[2]) };
}

function cellRawValue(cell: FindReplaceCell | undefined): string {
  if (!cell) return "";
  return typeof cell.value === "string" ? cell.value : "";
}

/**
 * A constant whose text starts with "=" is still text when the cell says so:
 * an entry typed with a leading apostrophe ('=1+2) or in a Text cell stores
 * the "=" as content. Only a cell that claims neither is read as a legacy
 * formula written into `value`.
 */
function isTextConstant(cell: FindReplaceCell): boolean {
  return cell.kind === "text" || cell.style?.quotePrefix === true;
}

function cellFormulaSource(cell: FindReplaceCell | undefined): string {
  if (!cell) return "";
  if (typeof cell.formula === "string" && cell.formula) return cell.formula;
  if (isTextConstant(cell)) return "";
  const raw = cellRawValue(cell);
  return raw.trim().startsWith("=") ? raw : "";
}

function isFormulaCell(cell: FindReplaceCell | undefined): boolean {
  if (!cell) return false;
  if (typeof cell.formula === "string" && cell.formula.trim().startsWith("=")) {
    return true;
  }
  if (isTextConstant(cell)) return false;
  return cellRawValue(cell).trim().startsWith("=");
}

/** What a cell is searched by, per `lookIn` (or the legacy `inFormulas`). */
type HaystackMode =
  | "display"
  | "formulaSource"
  | "formulas"
  | "notes"
  | "comments";

function haystackMode(options: FindReplaceOptions): HaystackMode {
  if (options.lookIn) {
    return options.lookIn === "values" ? "display" : options.lookIn;
  }
  return options.inFormulas ? "formulaSource" : "display";
}

const DEFAULT_ENTRY_LOCALE = "en-US";

/**
 * Excel "Look in: Formulas": a formula's source, else the constant as the
 * formula bar shows it (the text an edit starts from: 50%, 2026/9/21) --
 * without the apostrophe that marks text, which is the cell's prefix and not
 * its contents (Excel 16 measured: "audit_text" whole-cell finds a cell
 * entered as 'audit_text, "'audit_text" finds nothing; user audit
 * P-FIND-QUOTE).
 */
function cellEntryText(cell: FindReplaceCell, locale: string): string {
  const source = cellFormulaSource(cell);
  if (source) return source;
  const entry = spreadsheetEntryEditText(cell, locale);
  return entryTextPrefixed(cell, entry) ? entry.slice(1) : entry;
}

/** Whether the edit text leads with the text marker the stored value lacks. */
function entryTextPrefixed(cell: FindReplaceCell, entry: string): boolean {
  return entry.startsWith("'") && !(cell.value ?? "").startsWith("'");
}

function cellCommentObject(
  cell: FindReplaceCell,
): Record<string, unknown> | null {
  const comment = cell.comment;
  return comment && typeof comment === "object"
    ? (comment as Record<string, unknown>)
    : null;
}

/** A classic Note's text; null when the cell carries no note to search. */
function cellNoteText(cell: FindReplaceCell): string | null {
  const comment = cellCommentObject(cell);
  if (!comment || comment.kind !== "note") return null;
  return typeof comment.text === "string" ? comment.text : "";
}

/**
 * A threaded comment's messages, oldest first, one per line; null when the
 * cell carries no comment to search.
 */
function cellCommentText(cell: FindReplaceCell): string | null {
  const comment = cellCommentObject(cell);
  if (!comment || comment.kind === "note") return null;
  if (Array.isArray(comment.thread) && comment.thread.length > 0) {
    return comment.thread
      .map((message) =>
        message && typeof (message as { text?: unknown }).text === "string"
          ? (message as { text: string }).text
          : "",
      )
      .join("\n");
  }
  return typeof comment.text === "string" ? comment.text : "";
}

/** The text each cell is matched against under the given options. */
function makeHaystackResolver(
  cells: FindReplaceCellMap,
  options: FindReplaceOptions,
): (key: string, cell: FindReplaceCell) => string | null {
  const mode = haystackMode(options);
  if (mode === "display") return makeDisplayResolver(options.displayText);
  if (mode === "formulaSource") return (_key, cell) => cellFormulaSource(cell);
  if (mode === "formulas") {
    const locale = options.locale ?? DEFAULT_ENTRY_LOCALE;
    return (_key, cell) => cellEntryText(cell, locale);
  }
  if (mode === "notes") return (_key, cell) => cellNoteText(cell);
  return (_key, cell) => cellCommentText(cell);
}








function makeDisplayResolver(
  displayText?: FindDisplayText,
): (key: string, cell: FindReplaceCell) => string {
  if (displayText) {
    return (key, cell) => {
      const parsed = parseCellKey(key);
      return parsed ? displayText(parsed.row, parsed.col, cell) : cellRawValue(cell);
    };
  }
  return (key, cell) => {
    const parsed = parseCellKey(key);
    const fmt =
      typeof cell.style?.numberFormat === "string"
        ? cell.style.numberFormat.trim()
        : "";
    if (parsed && isFormulaCell(cell)) {
      if (typeof cell.display === "string") return cell.display;
      const stored = storedFormulaResult(typeof cell.value === "string" ? cell.value : null, { display: cell.display, chartValueCache: cell.chartValueCache });
      if (stored === undefined) return "";
      return fmt && typeof stored === "number" ? applyNumberFormat(stored, fmt) : formatCellValue(stored as Parameters<typeof formatCellValue>[0]);
    }
    const raw = cellRawValue(cell);
    if (fmt && raw.trim() !== "") {
      const num = Number(raw);
      if (Number.isFinite(num)) return applyNumberFormat(num, fmt);
    }
    return raw;
  };
}

const REGEX_META_RE = /[\\^$.*+?()[\]{}|]/g;

function escapeRegexLiteral(value: string): string {
  return value.replace(REGEX_META_RE, "\\$&");
}

const WIDE_FORM_RE = /[　！-～￠-￦\uff61-\uff9f\u30a0-\u30ff]/;

/**
 * Half-width katakana and its full-width twin fold onto one spelling: the
 * base kana in full width with the voiced/semi-voiced mark as a separate
 * combining unit (NFKC then NFD), so ｶﾞ (two units) and ガ (one unit) read
 * the same (live audit B169). Only the kana blocks are folded this way;
 * circled numbers, superscripts and the like are not width variants.
 */
function foldKanaWidth(character: string): string {
  return character.normalize("NFKC").normalize("NFD");
}

const KANA_RE = /[\uff61-\uff9f\u30a0-\u30ff]/;

/**
 * Width-folded text with, for every unit of the folded text, the index of the
 * unit of the original text it came from; `map[folded.length]` is the
 * original's length. A folded unit that expands (ガ -> カ + ゛) maps every
 * expanded unit to the same original unit.
 */
export function foldCharacterWidthWithMap(text: string): {
  folded: string;
  map: number[];
} {
  const map: number[] = [];
  if (!WIDE_FORM_RE.test(text)) {
    for (let index = 0; index <= text.length; index += 1) map.push(index);
    return { folded: text, map };
  }
  let out = "";
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const piece = KANA_RE.test(character)
      ? foldKanaWidth(character)
      : String.fromCharCode(narrowCode(text.charCodeAt(index)));
    for (let k = 0; k < piece.length; k += 1) map.push(index);
    out += piece;
  }
  map.push(text.length);
  return { folded: out, map };
}

/** Half-width twin of one UTF-16 unit (itself when it has none). */
function narrowCode(code: number): number {
  if (code >= 0xff01 && code <= 0xff5e) return code - 0xfee0;
  switch (code) {
    case 0x3000: return 0x20; // ideographic space
    case 0xffe0: return 0xa2; // ￠
    case 0xffe1: return 0xa3; // ￡
    case 0xffe2: return 0xac; // ￢
    case 0xffe3: return 0xaf; // ￣
    case 0xffe4: return 0xa6; // ￤
    case 0xffe5: return 0xa5; // ￥
    case 0xffe6: return 0x20a9; // ￦
    default: return code;
  }
}

/**
 * Fold full-width forms onto their half-width twins, one UTF-16 unit for one,
 * so an index into the folded text is the same index into the original and a
 * replacement can be spliced into the original text.
 */
export function foldCharacterWidth(text: string): string {
  return foldCharacterWidthWithMap(text).folded;
}

/**
 * Excel wildcard grammar: `*` matches any run, `?` one character, and `~`
 * escapes `*`, `?`, or `~`. The generated expression is always safe because
 * every non-wildcard character is escaped before compilation. Only the ASCII
 * signs are wildcards; with width folding a full-width ＊ is a literal that
 * finds both * and ＊.
 */
function excelFindPatternSource(query: string, foldWidth: boolean): string {
  const characters = Array.from(query);
  const literal = (character: string) =>
    escapeRegexLiteral(foldWidth ? foldCharacterWidth(character) : character);
  let source = "";
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    if (character === "~") {
      const next = characters[index + 1];
      if (next === "*" || next === "?" || next === "~") {
        source += escapeRegexLiteral(next);
        index += 1;
      } else {
        source += literal(character);
      }
    } else if (character === "*") {
      source += ".*";
    } else if (character === "?") {
      source += ".";
    } else {
      source += literal(character);
    }
  }
  return source;
}

function foldsWidth(options: FindReplaceOptions): boolean {
  return options.matchByte !== true;
}

function buildFindRegex(
  query: string,
  options: FindReplaceOptions,
  global = false,
): RegExp {
  const source = excelFindPatternSource(query, foldsWidth(options));
  // `(?=.)` prevents a wildcard-only expression from producing a second empty
  // match at the end during Replace All.
  const bounded = options.wholeCell ? `^(?:${source})$` : `(?=.)${source}`;
  return new RegExp(
    bounded,
    `${global ? "g" : ""}${options.caseSensitive ? "" : "i"}su`,
  );
}

/** Whether `haystack` matches `query` under the given options. */
function textMatches(
  haystack: string,
  query: string,
  options: FindReplaceOptions,
): boolean {
  if (query === "") return false;
  return buildFindRegex(query, options).test(
    foldsWidth(options) ? foldCharacterWidth(haystack) : haystack,
  );
}

/** Replace every occurrence of `query` within `source` (case option aware). */
function replaceOccurrences(
  source: string,
  query: string,
  replacement: string,
  options: FindReplaceOptions,
): string {
  if (query === "") return source;
  const regex = buildFindRegex(query, options, true);
  const { folded, map } = foldsWidth(options)
    ? foldCharacterWidthWithMap(source)
    : { folded: source, map: null };
  // Callback replacement keeps `$`, `$&`, and other user text literal.
  if (folded === source || !map) return source.replace(regex, () => replacement);
  // Matched on the folded text; the map takes each match back to the
  // original units, so the untouched parts keep their own width.
  let out = "";
  let last = 0;
  for (const match of folded.matchAll(regex)) {
    const start = map[match.index ?? 0];
    const end = map[(match.index ?? 0) + match[0].length];
    if (start < last) continue;
    out += source.slice(last, start) + replacement;
    last = end;
  }
  return out + source.slice(last);
}

/**
 * Find the keys of all cells matching `query`. Keys are sorted in row-major
 * order (top-to-bottom, then left-to-right) so Find Next traverses the grid the
 * way a user reads it — or column-major with `searchOrder: "columns"`.
 */
export function findMatches(
  cells: FindReplaceCellMap,
  query: string,
  options: FindReplaceOptions = {},
): string[] {
  if (!findCanMatch(query, options)) return [];
  const resolveHaystack = makeHaystackResolver(cells, options);
  // With a Find format, the cell's format must match too; with an empty
  // "Find what" the format alone decides (XLS-A103).
  const hit = (cell: FindReplaceCell, haystack: () => string | null) =>
    (!options.format || cellHasFindFormat(cell, options.format)) &&
    (query === "" ||
      (() => {
        // A cell without the searched note or comment is no candidate: "*"
        // with Match entire cell contents would otherwise match its missing
        // text.
        const text = haystack();
        return text !== null && textMatches(text, query, options);
      })());
  const matched: string[] = [];
  for (const [key, cell] of Object.entries(cells)) {
    if (!cell) continue;
    if (hit(cell, () => resolveHaystack(key, cell))) matched.push(key);
  }
  if (options.extraKeys && haystackMode(options) === "display") {
    const seen = new Set(matched);
    for (const key of options.extraKeys) {
      if (seen.has(key) || cells[key]) continue;
      seen.add(key);
      const spilled = { value: "" };
      if (hit(spilled, () => resolveHaystack(key, spilled))) matched.push(key);
    }
  }
  matched.sort(cellKeyComparator(options.searchOrder));
  return matched;
}

/** Row-major comparator for `row:col` keys. */
export function compareCellKeys(a: string, b: string): number {
  const pa = parseCellKey(a);
  const pb = parseCellKey(b);
  if (!pa || !pb) return a.localeCompare(b);
  if (pa.row !== pb.row) return pa.row - pb.row;
  return pa.col - pb.col;
}

/** Column-major comparator for `row:col` keys (Excel "Search: By Columns"). */
export function compareCellKeysByColumns(a: string, b: string): number {
  const pa = parseCellKey(a);
  const pb = parseCellKey(b);
  if (!pa || !pb) return a.localeCompare(b);
  if (pa.col !== pb.col) return pa.col - pb.col;
  return pa.row - pb.row;
}

/** The comparator Find Next walks the hits with. */
export function cellKeyComparator(
  order: FindSearchOrder | undefined,
): (a: string, b: string) => number {
  return order === "columns" ? compareCellKeysByColumns : compareCellKeys;
}

/**
 * Replace all matches across the sheet, returning a new cell map and the number
 * of cells changed. The input map is never mutated.
 *
 * - Without `inFormulas`: only literal (non-formula) cells are rewritten;
 *   formula cells are skipped (their displayed result is derived, not stored).
 * - With `inFormulas`: the raw formula source is rewritten and both `value` and
 *   `formula` are updated so the engine re-evaluates the new source.
 */
export function replaceAll(
  cells: FindReplaceCellMap,
  query: string,
  replacement: string,
  options: FindReplaceOptions = {},
): ReplaceAllResult {
  if (!findCanMatch(query, options)) return { cells, count: 0 };
  const mode = haystackMode(options);
  // A Note or a comment is found, never replaced: Excel's Replace tab only
  // looks in Formulas.
  if (mode === "notes" || mode === "comments") return { cells, count: 0 };
  const next: FindReplaceCellMap = { ...cells };
  let count = 0;

  const resolveDisplay =
    mode === "display" ? makeDisplayResolver(options.displayText) : null;

  for (const [key, cell] of Object.entries(cells)) {
    if (!cell) continue;
    if (options.format && !cellHasFindFormat(cell, options.format)) continue;

    // An empty "Find what" with a Find format (XLS-A103): the contents stay
    // and only the Replace format, if any, is laid on.
    if (query === "") {
      const formatted = withReplaceFormat(cell, options.replaceFormat);
      if (JSON.stringify(formatted) === JSON.stringify(cell)) continue;
      next[key] = formatted;
      count += 1;
      continue;
    }

    if (mode === "formulas" || mode === "formulaSource") {
      // Look in Formulas rewrites a formula's source and a constant's stored
      // entry alike (a rewrite that starts with "=" becomes a formula); the
      // legacy switch rewrites formula sources only.
      const locale = options.locale ?? DEFAULT_ENTRY_LOCALE;
      const source =
        mode === "formulas"
          ? cellEntryText(cell, locale)
          : cellFormulaSource(cell);
      if (source === "" || !textMatches(source, query, options)) continue;
      const updated = options.wholeCell
        ? replacement
        : replaceOccurrences(source, query, replacement, options);
      if (updated === source && !options.replaceFormat) continue;
      // A cell marked as text keeps its marker, as Excel keeps the prefix:
      // '007 with 7 replaced by 8 is the text 008, not the number 8.
      const marked =
        mode === "formulas" &&
        !cellFormulaSource(cell) &&
        entryTextPrefixed(cell, spreadsheetEntryEditText(cell, locale));
      next[key] = withReplaceFormat(
        updated === source
          ? cell
          : mode === "formulas"
            ? applyEntryRewrite(cell, marked ? `'${updated}` : updated, locale)
            : applyFormulaRewrite(cell, updated),
        options.replaceFormat,
      );
      count += 1;
      continue;
    }

    // Literal-cell replacement only (skip formula cells).
    if (isFormulaCell(cell)) continue;
    const display = resolveDisplay!(key, cell);
    if (!textMatches(display, query, options)) continue;
    const original = cellRawValue(cell);
    const updated = options.wholeCell
      ? replacement
      : replaceOccurrences(original, query, replacement, options);
    if (updated === original && !options.replaceFormat) continue;
    next[key] = withReplaceFormat({ ...cell, value: updated }, options.replaceFormat);
    count += 1;
  }

  return { cells: next, count };
}

/**
 * Apply a rewritten formula/text to a cell, keeping `formula` in sync with the
 * stored `value` exactly as the editor's commit path does.
 */
function applyFormulaRewrite(
  cell: FindReplaceCell,
  updated: string,
): FindReplaceCell {
  const nextCell: FindReplaceCell = { ...cell, value: updated };
  if (updated.trim().startsWith("=")) {
    nextCell.formula = updated.trim();
  } else {
    delete nextCell.formula;
  }
  return nextCell;
}

/**
 * Enter the rewritten edit text again, as typing it into the cell would
 * (the editor's entry path): a formula stays a formula; anything else is read
 * by parseSpreadsheetEntry, so "60%" is stored as 0.6 and a date as its
 * serial. A General cell takes the format the entry implies, any other format
 * stays, a Text (@) cell keeps the text as typed, and a leading apostrophe
 * becomes the quote prefix. Rich-text runs describe the old text and go with
 * it.
 */
function applyEntryRewrite(
  cell: FindReplaceCell,
  updated: string,
  locale: string,
): FindReplaceCell {
  const style = { ...(cell.style ?? {}) };
  delete style.quotePrefix;
  const withStyle = (next: FindReplaceCell): FindReplaceCell => {
    if (Object.keys(style).length > 0) next.style = style;
    else delete next.style;
    if (next.value !== cell.value) {
      delete next.runs;
      // The new text takes the cell's own font (cell-record-font).
      if (next.style) next.style = recordFontDrawn(next.style);
      // Furigana over the characters the replacement left keep their
      // readings (lib/spreadsheet/cell-phonetic).
      const phonetic = phoneticAfterValueWrite(
        cell.phonetic,
        cell.value ?? "",
        next.value ?? "",
        {
          editedInPlace: true,
          textual: !next.formula && next.kind !== "number",
        },
      );
      if (phonetic) next.phonetic = phonetic;
      else delete next.phonetic;
    }
    return next;
  };
  if (updated.trim().startsWith("=")) {
    return withStyle(applyFormulaRewrite(cell, updated));
  }
  const format = normalizeNumberFormatCode(style.numberFormat);
  const entry = parseSpreadsheetEntry(updated, {
    locale,
    textFormat: format === "@",
  });
  const next: FindReplaceCell = { ...cell, value: entry.value, kind: entry.kind };
  delete next.formula;
  if (entry.quotePrefix) style.quotePrefix = true;
  if (entry.numberFormat && format === "") {
    style.numberFormat = entry.numberFormat;
  }
  return withStyle(next);
}

// ---------------------------------------------------------------------------
// Workbook-scope Find & Replace (Excel/WPS "Within: Workbook" — search EVERY
// sheet tab). These wrap the single-sheet primitives above so the matching and
// replacement rules are byte-identical; the only addition is iterating the
// sheets in tab order and tagging each hit with its sheet index so the caller
// can jump to the right sheet + cell. Still PURE: a sheet is just its cell map.
// ---------------------------------------------------------------------------

/** A single match in a multi-sheet workbook. */
export interface WorkbookMatch {
  /** Index of the sheet (in the same order as the input array) holding the hit. */
  sheetIndex: number;
  /** The `row:col` key within that sheet. */
  key: string;
  /** Parsed row (0-based) for convenience — always derivable from `key`. */
  row: number;
  /** Parsed column (0-based) for convenience. */
  col: number;
}

/**
 * Find every match across an ordered list of sheet cell maps. Matches are
 * ordered first by SHEET (the input array order = tab order), then row-major
 * within each sheet — so "Find Next" walks tab 1 top-to-bottom, then tab 2, etc.
 * exactly like Excel's workbook-scope search. Returns an empty list for an empty
 * query. The single-sheet {@link findMatches} rules (case / wholeCell / formulas)
 * apply unchanged per sheet. `displayTextFor` gives each sheet's displayed text
 * from the host's workbook, so a value computed from another sheet is found
 * where it is shown.
 */
export function findMatchesInWorkbook(
  sheets: ReadonlyArray<FindReplaceCellMap>,
  query: string,
  options: FindReplaceOptions = {},
  displayTextFor?: (sheetIndex: number) => FindDisplayText | undefined,
  /** Cells a sheet shows without storing them (spilled results, a
   *  PivotTable report), visited through its display text. */
  extraKeysFor?: (sheetIndex: number) => Iterable<string> | undefined,
): WorkbookMatch[] {
  if (!findCanMatch(query, options)) return [];
  const out: WorkbookMatch[] = [];
  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex += 1) {
    const displayText = displayTextFor?.(sheetIndex);
    const extraKeys = extraKeysFor?.(sheetIndex);
    const keys = findMatches(sheets[sheetIndex], query, {
      ...options,
      ...(displayText ? { displayText } : {}),
      ...(extraKeys ? { extraKeys } : {}),
    });
    for (const key of keys) {
      const parsed = parseCellKey(key);
      out.push({
        sheetIndex,
        key,
        row: parsed ? parsed.row : 0,
        col: parsed ? parsed.col : 0,
      });
    }
  }
  return out;
}

/** The per-sheet outcome of a workbook-scope Replace All. */
export interface ReplaceAllInWorkbookResult {
  /** New cell maps, one per input sheet, in the same order (inputs untouched). */
  sheets: FindReplaceCellMap[];
  /** Total cells changed across all sheets. */
  count: number;
  /** Per-sheet change counts (parallel to `sheets`), for diagnostics. */
  countsBySheet: number[];
}

/**
 * Replace all matches across EVERY sheet, returning a new cell map per sheet and
 * the total number of cells changed. Each sheet is replaced with the exact same
 * rules as {@link replaceAll}; the inputs are never mutated. A sheet with no hit
 * gets its original cell-map reference back (so identity is preserved for the
 * caller's change-detection / undo bookkeeping).
 */
export function replaceAllInWorkbook(
  sheets: ReadonlyArray<FindReplaceCellMap>,
  query: string,
  replacement: string,
  options: FindReplaceOptions = {},
): ReplaceAllInWorkbookResult {
  if (!findCanMatch(query, options)) {
    return {
      sheets: sheets.map((s) => s),
      count: 0,
      countsBySheet: sheets.map(() => 0),
    };
  }
  const nextSheets: FindReplaceCellMap[] = [];
  const countsBySheet: number[] = [];
  let count = 0;
  for (const cells of sheets) {
    const result = replaceAll(cells, query, replacement, options);
    // Preserve the original reference when nothing changed (stable identity).
    nextSheets.push(result.count > 0 ? result.cells : cells);
    countsBySheet.push(result.count);
    count += result.count;
  }
  return { sheets: nextSheets, count, countsBySheet };
}

// ---------------------------------------------------------------------------
// Find All (Excel: "Find All" lists every hit under the dialog, one row per
// cell with the columns Book | Sheet | Name | Cell | Value | Formula, and a
// column heading sorts the list). Pure description of the hits; the dialog
// renders it and selects a row's cell when the row is chosen.
// ---------------------------------------------------------------------------

/** One sheet as Find All describes it. */
export interface FindAllSheet {
  id: string;
  name: string;
  cells: FindReplaceCellMap;
}

/** A defined name, reduced to what the Name column needs. */
export interface FindAllDefinedName {
  name: string;
  ref: string;
  scope?: "workbook" | "worksheet";
  scopeSheetId?: string;
}

/** One row of the Find All list. */
export interface FindAllRow {
  /** Stable row identity: `${sheetIndex}!${key}`. */
  id: string;
  sheetIndex: number;
  sheetId: string;
  sheetName: string;
  key: string;
  row: number;
  col: number;
  /** Absolute A1 address, as Excel lists it: `$B$3`. */
  cell: string;
  /** Defined names that refer to exactly this cell, comma separated. */
  name: string;
  /** What the cell displays. */
  value: string;
  /** The formula, or "" for a constant. */
  formula: string;
}

const SINGLE_CELL_NAME_REF_RE =
  /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

function namedCellTarget(
  entry: FindAllDefinedName,
): { sheetName: string | null; row: number; col: number } | null {
  const ref = entry.ref.trim().replace(/^=/, "");
  const match = SINGLE_CELL_NAME_REF_RE.exec(ref);
  if (!match) return null;
  const col = lettersToColumn(match[3]);
  const row = Number(match[4]) - 1;
  if (col < 0 || row < 0) return null;
  const quoted = match[1]?.replace(/''/g, "'");
  return { sheetName: quoted ?? match[2] ?? null, row, col };
}

/**
 * Describe `matches` (as {@link findMatchesInWorkbook} returns them, sheet
 * indices into `sheets`) as Find All rows, in the order given. The Value
 * column is `displayTextFor`'s text when the host supplies it.
 */
export function describeFindMatches(
  sheets: ReadonlyArray<FindAllSheet>,
  matches: ReadonlyArray<WorkbookMatch>,
  names: ReadonlyArray<FindAllDefinedName> = [],
  displayTextFor?: (sheetIndex: number) => FindDisplayText | undefined,
): FindAllRow[] {
  const resolvers = new Map<number, (key: string, cell: FindReplaceCell) => string>();
  const targets = names.flatMap((entry) => {
    const target = namedCellTarget(entry);
    return target ? [{ entry, target }] : [];
  });
  const rows: FindAllRow[] = [];
  for (const match of matches) {
    const sheet = sheets[match.sheetIndex];
    if (!sheet) continue;
    const cell = sheet.cells[match.key] ?? {};
    let resolve = resolvers.get(match.sheetIndex);
    if (!resolve) {
      resolve = makeDisplayResolver(displayTextFor?.(match.sheetIndex));
      resolvers.set(match.sheetIndex, resolve);
    }
    const cellNames = targets
      .filter(({ entry, target }) => {
        if (target.row !== match.row || target.col !== match.col) return false;
        if (target.sheetName !== null) {
          return target.sheetName.toLowerCase() === sheet.name.toLowerCase();
        }
        // An unqualified reference only names a cell of its own sheet.
        return entry.scope === "worksheet" && entry.scopeSheetId === sheet.id;
      })
      .map(({ entry }) => entry.name);
    rows.push({
      id: `${match.sheetIndex}!${match.key}`,
      sheetIndex: match.sheetIndex,
      sheetId: sheet.id,
      sheetName: sheet.name,
      key: match.key,
      row: match.row,
      col: match.col,
      cell: `$${columnToLetters(match.col)}$${match.row + 1}`,
      name: [...new Set(cellNames)].join(", "),
      value: resolve(match.key, cell),
      formula: isFormulaCell(cell) ? cellFormulaSource(cell) : "",
    });
  }
  return rows;
}

/** The Find All columns a heading click sorts by (Book is one file: no-op). */
export type FindAllSortColumn =
  | "book"
  | "sheet"
  | "name"
  | "cell"
  | "value"
  | "formula";

/**
 * Sort Find All rows by one column. Text columns compare with `collator`
 * (numeric aware, so 9 sorts before 10); Cell compares by sheet, then row,
 * then column. The sort is stable: rows that tie keep the search order.
 */
export function sortFindAllRows(
  rows: ReadonlyArray<FindAllRow>,
  column: FindAllSortColumn,
  descending: boolean,
  collator: Intl.Collator = new Intl.Collator(undefined, {
    numeric: true,
    sensitivity: "base",
  }),
): FindAllRow[] {
  const direction = descending ? -1 : 1;
  const compare = (a: FindAllRow, b: FindAllRow): number => {
    switch (column) {
      case "book":
        return 0;
      case "sheet":
        return collator.compare(a.sheetName, b.sheetName);
      case "cell":
        return a.sheetIndex - b.sheetIndex || a.row - b.row || a.col - b.col;
      default:
        return collator.compare(a[column], b[column]);
    }
  };
  return [...rows].sort((a, b) => direction * compare(a, b));
}
