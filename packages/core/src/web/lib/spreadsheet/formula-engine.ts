/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
} from "./grid-sizing";
import { FORMULA_FUNCTION_NAMES } from "./formula-function-names";
import {
  applyNumberFormat,
  formatGeneralNumber,
} from "./number-format";
import {
  characterAt,
  formatSheetNameForFormula,
  isBareNameCharacter,
  isBareNameSpelling,
  isDefinedNameCharacter,
  sheetNameNeedsQuotes,
} from "./name-syntax";
import type { WhatIfTableDefinition } from "./what-if-table-definition";
import type { ExternalDataLookup } from "./external-data";
import type { KeptFormulaInput } from "./kept-formula-results";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

/**
 * Whether a reference to sheet `name` must wrap it in single quotes, and the
 * prefix a formula writes for it — the one rule every writer of a sheet prefix
 * shares with the reader above. Defined in ./name-syntax, beside the spelling
 * it extends, so modules that only write references need not load the engine.
 */
export { formatSheetNameForFormula, sheetNameNeedsQuotes };

// ---------------------------------------------------------------------------
// Value model
// ---------------------------------------------------------------------------

export interface FormulaError {
  readonly __formulaError: true;
  readonly code: string;
}

export type CellValue = number | string | boolean | null | FormulaError;

/**
 * A dynamic-array result (Excel 365 / Google Sheets "spill").
 *
 * Produced by the array functions (FILTER, UNIQUE, SORT, SEQUENCE, TRANSPOSE).
 * `grid` is a rectangular rows×cols block (always ≥ 1×1). It is an INTERNAL
 * evaluation value: a formula whose top-level result is an ArrayValue collapses
 * to the block's top-left cell for normal single-cell display (Excel parity —
 * the anchor cell shows the first element), while `getSpill()` exposes the whole
 * block so the grid can spill it into adjacent cells. `CellValue` (the persisted
 * / public scalar type) intentionally does NOT include this, so every existing
 * scalar code path is unaffected.
 */
export interface ArrayValue {
  readonly __array: true;
  readonly grid: CellValue[][];
}

/**
 * A first-class function value (Excel 365 LAMBDA). Produced by `LAMBDA(p1, …, pn,
 * body)` WITHOUT evaluating `body`; it captures its parameter names, the body
 * AST, and a snapshot of the lexical scope in force at definition time (its
 * `closure`, the enclosing LET bindings) so the lambda is a proper closure. It
 * is an INTERNAL evaluation value only: like {@link ArrayValue} it is NOT part of
 * {@link CellValue}, so it can never be persisted. A bare LAMBDA used as a cell
 * value collapses to `#CALC!` (Excel parity — a lambda must be CALLED). It is
 * invoked by the `apply` node (`f(args)` / `LAMBDA(...)(args)`) and by the
 * lambda-helper functions (MAP/REDUCE/SCAN/BYROW/BYCOL/MAKEARRAY).
 */
export interface LambdaValue {
  readonly __lambda: true;
  readonly params: readonly string[];
  readonly body: Node;
  readonly closure: ReadonlyMap<string, EngineValue>;
}

/**
 * The sentinel an EXPLICITLY OMITTED LAMBDA argument binds to — the value Excel
 * hands a parameter whose call slot was left empty, so `ISOMITTED(param)` can
 * detect it. It arises ONLY from an `empty` argument node in a lambda call
 * (`LAMBDA(x, y, …)(1,)` → param `y` is OMITTED); a slot filled with any real
 * value (including an explicit blank cell or 0) is NOT omitted. Like
 * {@link ArrayValue}/{@link LambdaValue} it is an INTERNAL evaluation value, NOT
 * part of {@link CellValue}, so it can never be persisted; used anywhere a scalar
 * is required it collapses to a blank (null) — Excel parity, an unchecked omitted
 * optional arg reads as empty. The unique symbol-keyed brand can't be forged by a
 * cell value.
 */
export interface OmittedValue {
  readonly __omitted: true;
}

export function isOmittedValue(v: unknown): v is OmittedValue {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __omitted?: unknown }).__omitted === true
  );
}

/** The value an expression can evaluate to internally: a scalar, an array, a
 *  lambda (a first-class function value), or the omitted-argument sentinel. */
export type EngineValue = CellValue | ArrayValue | LambdaValue | OmittedValue;

export function isArrayValue(v: unknown): v is ArrayValue {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __array?: unknown }).__array === true
  );
}

export function isLambdaValue(v: unknown): v is LambdaValue {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __lambda?: unknown }).__lambda === true
  );
}

export const ERR = (code: string): FormulaError => ({ __formulaError: true, code });
export const ERRORS = {
  DIV0: ERR("#DIV/0!"),
  VALUE: ERR("#VALUE!"),
  NAME: ERR("#NAME?"),
  REF: ERR("#REF!"),
  NA: ERR("#N/A"),
  NUM: ERR("#NUM!"),
  SPILL: ERR("#SPILL!"),
  CIRCULAR: ERR("#CIRCULAR!"),
  CALC: ERR("#CALC!"),
  NULL: ERR("#NULL!"),
  PARSE: ERR("#ERROR!"),
  /** IMAGE while its picture is still loading (Excel's #BUSY!). */
  BUSY: ERR("#BUSY!"),
  /** IMAGE whose source cannot be reached (Excel's #CONNECT!). */
  CONNECT: ERR("#CONNECT!"),
  /** IMAGE whose source is refused for security (Excel's #BLOCKED!). */
  BLOCKED: ERR("#BLOCKED!"),
} as const;

export function isFormulaError(v: unknown): v is FormulaError {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __formulaError?: unknown }).__formulaError === true &&
    (v as { __image?: unknown }).__image !== true
  );
}

/**
 * A picture placed IN a cell (Excel's Place in Cell) as a formula value.
 *
 * Excel stores such a cell as the error `#VALUE!` with a rich-value
 * attachment, so a reader that does not know pictures sees `#VALUE!`; one that
 * does sees the picture. The value keeps that shape: it is structurally a
 * `#VALUE!` error (any unaware formatter shows Excel's own fallback text), but
 * it is NOT an error to the engine. {@link isFormulaError} refuses it, so
 * ISERROR is FALSE, IFERROR passes it through, TYPE answers 128 (compound
 * data), COUNTA counts it, ISBLANK/ISTEXT/ISNUMBER are FALSE, and arithmetic
 * or concatenation on it is `#VALUE!` (Microsoft Support, "TYPE function",
 * "IMAGE function").
 *
 * The bytes are not carried: the value names the cell the picture lives in
 * (the sheet name the workbook registered, 0-based row and column) and its
 * alternative text, and the host reads the picture from that cell when it
 * draws a reference to it. That keeps data URLs out of memo snapshots and out
 * of the background calculation's messages.
 */
export interface ImageCellValue extends FormulaError {
  readonly __image: true;
  readonly code: "#VALUE!";
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  readonly alt: string;
}

export function isImageCellValue(v: unknown): v is ImageCellValue {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __image?: unknown }).__image === true &&
    typeof (v as { sheet?: unknown }).sheet === "string" &&
    Number.isSafeInteger((v as { row?: unknown }).row) &&
    Number.isSafeInteger((v as { col?: unknown }).col)
  );
}

/**
 * The picture Excel's IMAGE function returns (`=IMAGE(source, [alt_text],
 * [sizing], [height], [width])`): a value of the same compound kind as a
 * picture placed in a cell -- TYPE 128, not an error to ISERROR, #VALUE! to
 * arithmetic and text -- that names its picture by its https `source`
 * instead of a cell, with the sizing mode it is drawn in (Microsoft Support,
 * "IMAGE function"): 0 fit keeping the aspect ratio, 1 fill the cell, 2 the
 * original size, 3 the given `height`/`width` in pixels (one of them alone
 * keeps the aspect ratio).
 */
export interface WebImageValue extends FormulaError {
  readonly __image: true;
  readonly code: "#VALUE!";
  readonly source: string;
  readonly alt: string;
  readonly sizing: 0 | 1 | 2 | 3;
  readonly height?: number;
  readonly width?: number;
}

export function isWebImageValue(v: unknown): v is WebImageValue {
  return (
    typeof v === "object" &&
    v !== null &&
    (v as { __image?: unknown }).__image === true &&
    typeof (v as { source?: unknown }).source === "string"
  );
}

/** A picture as a value: one placed in a cell or one IMAGE returned. */
export function isPictureValue(v: unknown): v is ImageCellValue | WebImageValue {
  return isImageCellValue(v) || isWebImageValue(v);
}

export function createWebImageValue(
  source: string,
  alt: string,
  sizing: 0 | 1 | 2 | 3,
  size: { height?: number; width?: number } = {},
): WebImageValue {
  return Object.freeze({
    __formulaError: true as const,
    __image: true as const,
    code: "#VALUE!" as const,
    source,
    alt,
    sizing,
    ...(size.height !== undefined ? { height: size.height } : {}),
    ...(size.width !== undefined ? { width: size.width } : {}),
  });
}

/**
 * Where an IMAGE source's picture stands for the workbook being evaluated:
 * still loading, refused with one of Excel's IMAGE errors, or shown. The
 * host that fetches pictures supplies it ({@link WorkbookOptions.webImageStatus});
 * without one every valid source reads as its picture value.
 */
export type WebImageStatus = "loading" | "ready" | "#CONNECT!" | "#BLOCKED!" | "#VALUE!";

export function createImageCellValue(
  sheet: string,
  row: number,
  col: number,
  alt = "",
): ImageCellValue {
  return Object.freeze({
    __formulaError: true as const,
    __image: true as const,
    code: "#VALUE!" as const,
    sheet,
    row,
    col,
    alt,
  });
}

// ---------------------------------------------------------------------------
// A1 notation
// ---------------------------------------------------------------------------

/** 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnToLetters(col: number): string {
  let n = col;
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/** "A" -> 0, "Z" -> 25, "AA" -> 26. Returns -1 for invalid. */
export function lettersToColumn(letters: string): number {
  if (!letters) return -1;
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

export interface CellRef {
  row: number; // 0-indexed
  col: number; // 0-indexed
}

export const CELL_REF_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

/** A word that is a cell reference inside Excel's grid (A1:XFD1048576). */
function cellRefInGrid(word: string): boolean {
  const match = CELL_REF_RE.exec(word);
  if (!match) return false;
  let column = 0;
  for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
  const row = Number(match[2]);
  return column <= 16_384 && row >= 1 && row <= 1_048_576;
}
/** One side of a whole-column (`$A`) or whole-row (`$12`) reference. */
const WHOLE_AXIS_SIDE_RE = {
  column: /^\$?([A-Za-z]{1,3})/,
  row: /^\$?([0-9]{1,7})/,
} as const;

/**
 * Read `A:A` / `$B:$D` / `1:1` / `$2:$4` starting at `start`, or null.
 *
 * A whole-axis reference is BOTH sides bare: `A:A1` and `A1:A` are not one (a
 * side carrying the other axis makes it an ordinary range or a name), so the
 * side must not be followed by a character that could continue the address.
 */
function _matchWholeAxisRef(
  src: string,
  start: number,
  axis: "column" | "row",
): { token: Token; next: number } | null {
  const side = WHOLE_AXIS_SIDE_RE[axis];
  const left = side.exec(src.slice(start));
  if (!left) return null;
  let cursor = start + left[0].length;
  if (src[cursor] !== ":") return null;
  const right = side.exec(src.slice(cursor + 1));
  if (!right) return null;
  const after = src[cursor + 1 + right[0].length];
  // `A:A1` / `1:1A` — the far side keeps going, so this is not a whole axis.
  if (after !== undefined && /[A-Za-z0-9$_.]/.test(after)) return null;
  cursor += 1 + right[0].length;
  return {
    // `value` is the written spelling, `$` included, so a formula rewritten
    // from tokens (a sheet rename, a fill) keeps `$A:$A` absolute.
    token: { type: "wholeref", value: src.slice(start, cursor), axis, from: left[1], to: right[1] },
    next: cursor,
  };
}

/** Parse "A1" / "$B$3" -> { row: 0-indexed, col: 0-indexed }. null if invalid. */
export function parseCellRef(ref: string): CellRef | null {
  const m = CELL_REF_RE.exec(ref.trim());
  if (!m) return null;
  const col = lettersToColumn(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || row < 0 || !Number.isFinite(row)) return null;
  return { row, col };
}

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

type TokenType =
  | "number"
  // Excel's INTERSECTION operator: a space between two references. Emitted by
  // the tokenizer only where a space can mean nothing else (see the whitespace
  // branch in `tokenize`), so ordinary spacing stays whitespace.
  | "isect"
  | "string"
  | "ident" // function name or boolean
  | "ref"
  | "tableref" // a structured / ListObject reference: Table1[Amount], [@Col]
  | "error" // an error literal like #REF! (from a structural rewrite)
  | "op"
  | "lparen"
  | "rparen"
  | "lbrace" // `{` opening an inline array constant: {1,2;3,4}
  | "rbrace" // `}` closing an inline array constant
  | "semicolon" // `;` the array-constant ROW separator
  | "comma"
  // `A:A` / `1:1` — a WHOLE-COLUMN or WHOLE-ROW reference. Both sides are
  // bare (no row on a column ref, no column on a row ref), which is what tells
  // it apart from an ordinary `A1:A5` range.
  | "wholeref"
  | "colon";

// Error literals that can appear inside a formula (e.g. after a row delete
// rewrites a reference to #REF!). Longest-match order is not required since
// they are mutually non-prefixing except for distinct suffixes.
const ERROR_LITERALS = [
  "#DIV/0!",
  "#VALUE!",
  "#NAME?",
  "#N/A",
  "#NUM!",
  "#NULL!",
  "#REF!",
  "#SPILL!",
  "#ERROR!",
  "#CIRCULAR!",
  "#CALC!",
];

/**
 * The address Excel leaves behind `#REF!` when the sheet a reference named is
 * deleted: a cell, a cell range, or a whole-column / whole-row span. It must not
 * run on into a longer name (`#REF!A1B`), a call or another sheet prefix.
 */
const DELETED_SHEET_REF_TAIL =
  /^(?:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[0-9]{1,7}:\$?[0-9]{1,7})(?![A-Za-z0-9_.(!\\])/;

export interface Token {
  type: TokenType;
  value: string;
  // Only on a `ref` or `wholeref` token produced by the sheet-prefix grammar
  // (Sheet2!A1 / 'My Sheet'!A1 / Sheet1:Sheet3!A1 / Sheet2!A:A). `sheet` is the
  // already-unquoted target
  // sheet name; `endSheet`, when present, is the last sheet of a 3D span (the
  // span runs `sheet`..`endSheet` in workbook order). Absent on a same-sheet
  // ref, so an unprefixed token is byte-identical to before.
  sheet?: string;
  endSheet?: string;
  // Only on a `tableref` token (a structured / ListObject reference). `table` is
  // the table name (null for the unqualified this-row form `[@Col]` that targets
  // the formula's own table). `spec` is the decoded bracket selector — see
  // {@link StructuredSpec}. Absent on every non-table token, so an ordinary ref
  // / ident / string token is byte-identical to before.
  table?: string | null;
  spec?: StructuredSpec;
  /** Exact source spelling of one structured ref, used by formula rewrites. */
  raw?: string;
  // Only on an `error` token spelling a deleted-sheet reference (`#REF!A1`):
  // the error it evaluates to, while `value` keeps the whole spelling.
  errorCode?: string;
  // Only on a `wholeref` token: which axis it spans and its two bare endpoints
  // (`from`/`to` are the column letters or the row numbers, $ already stripped).
  axis?: "column" | "row";
  from?: string;
  to?: string;
  // Only on a `ref` token immediately followed by a `#` SPILL-RANGE operator
  // (`A1#`) — the reference addresses the ENTIRE dynamic-array block that spilled
  // from anchor cell `value`. Absent on every ordinary ref, so a `#`-less ref is
  // byte-identical to before. A single-sheet qualifier may precede it
  // (`Sheet2!A1#`); a 3D span may not, as in Excel.
  spill?: boolean;
}

// ---------------------------------------------------------------------------
// Structured (ListObject) references — Table1[Column], [#Headers], [@Col], …
// ---------------------------------------------------------------------------
//
// A structured reference selects a BAND (rows) × COLUMNS rectangle of an Excel
// "ListObject" table by NAME rather than by A1 address, so it survives row/col
// inserts. The decoded selector a `tableref` token carries is a {@link
// StructuredSpec}: which item band (#Data default / #Headers / #Totals / #All
// / this-row `@`) and which column(s) (a single name, an inclusive [A]:[B] span,
// or null = the whole table width). Resolution maps it to an A1 box against the
// table's registered geometry, then reuses the SAME range/aggregate machinery
// as an ordinary ref — see `resolveTableNode`.

/** The row band a structured reference selects. */
export type StructuredBand =
  | "data" // the data rows (the default for `Table[Col]` / a bare `Table`)
  | "headers" // the header row
  | "totals" // the totals row (or #REF! when the table has none)
  | "all" // header + data + totals
  | "headersData" // `[[#Headers],[#Data]]`: the header row and the data rows
  | "dataTotals" // `[[#Data],[#Totals]]`: the data rows and the totals row
  | "thisrow"; // the column cell(s) on the formula's OWN row (`[@Col]`)

/**
 * A decoded structured-reference selector. `band` is the row band; `column` is a
 * single column name, `columnEnd` (when present) makes it the inclusive span
 * `[column]:[columnEnd]`, and a null `column` means "the whole table width"
 * (every column). Column names are matched case-insensitively at resolution.
 */
export interface StructuredSpec {
  band: StructuredBand;
  column: string | null;
  columnEnd?: string;
}

// A non-special-item, non-bracketed column NAME inside [...] may contain any
// char except the structured-ref metacharacters (`[ ] # ' @`); those must be
// escaped with `'` (Excel parity). `readStructuredName` below decodes the
// escapes, so the returned name is the literal column header text.

/**
 * Read ONE column name starting at `src[i]` inside a `[[ … ]]` (or bare `[ … ]`)
 * column selector. Stops at an unescaped `]` or `:`. An apostrophe escapes the
 * next char (so a header literally containing `]`/`[`/`#`/`@`/`'` is written
 * `'…`). Returns the decoded name + the next index (the stop char), or null if a
 * metacharacter appears unescaped (malformed). Leading/trailing spaces are kept
 * verbatim — Excel matches the exact header text.
 */
function readStructuredName(
  src: string,
  i: number,
  // A bare name ends at a `:` (the `[Col1]:[Col2]` span separator); inside
  // its own brackets a column named "Time:Start" keeps its colon (B303).
  stopAtColon = true,
): { name: string; next: number } | null {
  let name = "";
  let j = i;
  const n = src.length;
  while (j < n) {
    const c = src[j];
    if (c === "'") {
      // Escape: the NEXT char is taken literally (Excel's structured-ref escape).
      if (j + 1 >= n) return null;
      name += src[j + 1];
      j += 2;
      continue;
    }
    if (c === "]" || (stopAtColon && c === ":")) return { name, next: j };
    // An unescaped meta character is malformed -- but an `@` or `#` inside a
    // name, past its first character, is the name's own: Excel writes and
    // calculates Table3[[#This Row],[LogS @pH 11]] so (xlsx-rounds formula
    // truth R230-X06). At the start they keep their meaning ([@Col], [#Totals]).
    if (c === "[" || ((c === "#" || c === "@") && name === "")) return null;
    if (c === "#" || c === "@") {
      name += c;
      j += 1;
      continue;
    }
    name += c;
    j += 1;
  }
  return null; // ran off the end without a closing bracket
}

/**
 * Read the FULL bracketed selector of a structured reference starting at the
 * opening `[` (`src[i] === "["`). Handles every Excel form:
 *   - `[Column]` / `[ #Headers ]` / `[@Column]` — a single un-nested item
 *   - `[[Column With ] Special]]` — a column whose name needs `[[ … ]]`
 *   - `[[#Totals],[Amount]]` — a special item plus a column
 *   - `[[Col1]:[Col2]]` — an inclusive multi-column span
 *   - `[@[Col1]:[Col2]]` / `[@],[Col]` style this-row forms
 * Returns the decoded {@link StructuredSpec} + the index just AFTER the closing
 * `]`, or null when the brackets are malformed (so the tokenizer rejects the
 * whole formula rather than silently mis-parsing). `band` defaults to "data".
 */
function readStructuredSelector(
  src: string,
  i: number,
): { spec: StructuredSpec; next: number } | null {
  const n = src.length;
  if (src[i] !== "[") return null;
  let j = i + 1;
  let band: StructuredBand = "data";
  let bandSeen = false;
  let column: string | null = null;
  let columnEnd: string | undefined;
  // A leading `@` (this-row) may appear bare: [@Column] or [@[Col]:[Col]].
  if (src[j] === "@") {
    band = "thisrow";
    bandSeen = true;
    j += 1;
  }

  // Read one "part" — either a #special item, a [bracketed] name, or a bare name
  // running to the next `]`/`,`/`:`. Returns the next index or null on error.
  const setBand = (raw: string): boolean => {
    const key = raw.trim().toLowerCase();
    const map: Record<string, StructuredBand> = {
      "#data": "data",
      "#headers": "headers",
      "#totals": "totals",
      "#all": "all",
      "#this row": "thisrow",
    };
    if (!(key in map)) return false;
    // Two special items name a contiguous pair of bands (Microsoft: #Headers
    // with #Data, #Data with #Totals); the second item used to replace the
    // first, dropping the header row (live audit B302).
    const next = map[key];
    if (bandSeen) {
      const pair = [band, next].sort().join("+");
      if (pair === "data+headers") band = "headersData";
      else if (pair === "data+totals") band = "dataTotals";
      else band = next;
    } else {
      band = next;
    }
    // `@` (thisrow) and a #special item are mutually exclusive bands; if `@`
    // already set thisrow, only "#this row" is compatible.
    bandSeen = true;
    return true;
  };
  const setColumn = (name: string): boolean => {
    if (column === null) {
      column = name;
      return true;
    }
    if (columnEnd === undefined) {
      columnEnd = name;
      return true;
    }
    return false; // a third column name is malformed
  };

  while (j < n) {
    const c = src[j];
    if (c === "]") {
      j += 1;
      return { spec: buildSpec(band, bandSeen, column, columnEnd), next: j };
    }
    if (c === ",") {
      j += 1;
      continue;
    }
    if (c === ":") {
      // span separator between two [bracketed] columns; handled by the column
      // setter (the next part becomes columnEnd).
      j += 1;
      continue;
    }
    if (c === "#") {
      // A special-item word runs to the next `]`/`,`.
      let word = "";
      while (j < n && src[j] !== "]" && src[j] !== ",") word += src[j++];
      if (!setBand(word)) return null;
      continue;
    }
    if (c === "[") {
      // A bracketed part: a column name (possibly with special chars) or a
      // bracketed #special. Read to the matching `]`.
      const inner = src[j + 1] === "#" ? readSpecialInBracket(src, j + 1) : null;
      if (inner) {
        if (!setBand(inner.name)) return null;
        j = inner.next;
        if (src[j] !== "]") return null;
        j += 1;
        continue;
      }
      const name = readStructuredName(src, j + 1, false);
      if (!name) return null;
      if (src[name.next] !== "]") return null;
      if (!setColumn(name.name)) return null;
      j = name.next + 1;
      continue;
    }
    // A bare (un-bracketed) name part: runs to the next `]`/`,`/`:`.
    const name = readStructuredName(src, j);
    if (!name || name.name === "") return null;
    if (!setColumn(name.name)) return null;
    j = name.next;
  }
  return null; // no closing `]`
}

/** Read a bracketed `#special` item word (`[#Totals]`) starting at the `#`. */
function readSpecialInBracket(
  src: string,
  i: number,
): { name: string; next: number } | null {
  if (src[i] !== "#") return null;
  let word = "";
  let j = i;
  const n = src.length;
  while (j < n && src[j] !== "]" && src[j] !== ",") word += src[j++];
  return { name: word, next: j };
}

/** Assemble the final spec, applying the bare-table default (band "data"). */
function buildSpec(
  band: StructuredBand,
  bandSeen: boolean,
  column: string | null,
  columnEnd: string | undefined,
): StructuredSpec {
  const spec: StructuredSpec = {
    band: bandSeen ? band : "data",
    column,
  };
  if (columnEnd !== undefined) spec.columnEnd = columnEnd;
  return spec;
}

// ---------------------------------------------------------------------------
// Structured-reference RESOLUTION (registry -> A1 box)
// ---------------------------------------------------------------------------

/**
 * One table (ListObject) in the registry the structured-ref resolver consults.
 * `range` is 0-indexed/inclusive and covers the HEADER row (r1) plus the DATA
 * rows (r1+1..r2); `hasTotals` is true when an appended totals row exists (it
 * occupies grid row r2+1). `columns` are the header texts in left-to-right order
 * (one per column c1..c2), used for case-insensitive column-name matching.
 */
export interface TableDefinition {
  name: string;
  range: { r1: number; c1: number; r2: number; c2: number };
  /** Omitted by legacy callers; defaults to the historical header-present mode. */
  hasHeader?: boolean;
  hasTotals: boolean;
  columns: readonly string[];
}

/**
 * A resolved structured ref: the A1 box it addresses, or a formula error.
 * `sheet` names the table's OWN sheet when that is not the formula's sheet —
 * table names are workbook-scoped in Excel, so `=SUM(Table1[Amount])` on a
 * summary sheet reads Table1's cells where Table1 lives. Absent means the
 * formula's own sheet.
 */
export type TableRefResult =
  | { from: string; to: string; sheet?: string }
  | FormulaError;

/** True iff a `TableRefResult` is the error branch. */
export function isTableRefError(r: TableRefResult): r is FormulaError {
  return isFormulaError(r);
}

/**
 * Build a deterministic structured-reference resolver over a fixed set of tables.
 * The returned function is the `resolveTableRef` EvalContext callback: it maps
 * `(tableName, spec, currentCell)` to the A1 box the structured ref addresses,
 * or a FormulaError. Pure + side-effect-free, so a host can build it once per
 * sheet and inject the same instance into every formula's context.
 *
 * Matching rules (Excel parity):
 *   - table name + column name match CASE-INSENSITIVELY (trimmed);
 *   - an unknown table / unknown column / `#Totals` on a table with no totals row
 *     all yield ERRORS.REF;
 *   - the unqualified this-row form (`[@Col]`, tableName === null) finds the table
 *     CONTAINING the formula's own cell; outside any table -> ERRORS.VALUE;
 *   - a `@` (this-row) form whose formula row is the header/totals row, or any
 *     other row outside the data band, yields ERRORS.VALUE (Excel's "@ outside the
 *     table" behavior).
 */
export function makeTableRefResolver(
  tables: readonly TableDefinition[],
): (
  tableName: string | null,
  spec: StructuredSpec,
  currentCell: { row: number; col: number } | undefined,
) => TableRefResult {
  const byLower = new Map<string, TableDefinition>();
  for (const t of tables) {
    const key = t.name.trim().toLowerCase();
    if (key !== "" && !byLower.has(key)) byLower.set(key, t);
  }

  // The table whose RANGE (header..data, or its totals row) contains (row,col).
  const tableContaining = (
    cell: { row: number; col: number } | undefined,
  ): TableDefinition | null => {
    if (!cell) return null;
    for (const t of tables) {
      const { r1, c1, r2, c2 } = t.range;
      const lastRow = t.hasTotals ? r2 + 1 : r2;
      if (cell.row >= r1 && cell.row <= lastRow && cell.col >= c1 && cell.col <= c2) {
        return t;
      }
    }
    return null;
  };

  // The 0-indexed grid column of a named column inside a table, or null.
  const columnIndex = (table: TableDefinition, name: string): number | null => {
    const want = name.trim().toLowerCase();
    for (let k = 0; k < table.columns.length; k += 1) {
      if (table.columns[k].trim().toLowerCase() === want) return table.range.c1 + k;
    }
    return null;
  };

  return function resolve(tableName, spec, currentCell): TableRefResult {
    const table =
      tableName === null
        ? tableContaining(currentCell)
        : byLower.get(tableName.trim().toLowerCase()) ?? null;
    if (!table) {
      // A this-row `[@Col]` outside ANY table is #VALUE! (Excel); a named but
      // unknown table is #REF!.
      return tableName === null ? ERRORS.VALUE : ERRORS.REF;
    }

    const { r1, c1, r2, c2 } = table.range;
    const firstDataRow = r1 + (table.hasHeader === false ? 0 : 1);
    const totalsRow = table.hasTotals ? r2 + 1 : null;

    // Resolve the column span [c1..c2] of the box. A null column = whole width.
    let colStart = c1;
    let colEnd = c2;
    if (spec.column !== null) {
      const start = columnIndex(table, spec.column);
      if (start === null) return ERRORS.REF; // unknown column
      if (spec.columnEnd !== undefined) {
        const end = columnIndex(table, spec.columnEnd);
        if (end === null) return ERRORS.REF;
        colStart = Math.min(start, end);
        colEnd = Math.max(start, end);
      } else {
        colStart = start;
        colEnd = start;
      }
    }

    // Resolve the row band [rowStart..rowEnd].
    let rowStart: number;
    let rowEnd: number;
    switch (spec.band) {
      case "headers":
        if (table.hasHeader === false) return ERRORS.REF;
        rowStart = r1;
        rowEnd = r1;
        break;
      case "totals":
        if (totalsRow === null) return ERRORS.REF; // no totals row -> #REF!
        rowStart = totalsRow;
        rowEnd = totalsRow;
        break;
      case "all":
        rowStart = r1;
        rowEnd = totalsRow ?? r2;
        break;
      case "headersData":
        if (table.hasHeader === false) return ERRORS.REF;
        rowStart = r1;
        rowEnd = r2;
        break;
      case "dataTotals":
        if (totalsRow === null) return ERRORS.REF;
        if (r2 < firstDataRow) return ERRORS.REF;
        rowStart = firstDataRow;
        rowEnd = totalsRow;
        break;
      case "thisrow": {
        if (!currentCell) return ERRORS.VALUE;
        const row = currentCell.row;
        // The `@` form is valid only when the formula sits on a DATA row of THIS
        // table (Excel: `@` on the header/totals row or outside is #VALUE!).
        if (row < firstDataRow || row > r2) return ERRORS.VALUE;
        rowStart = row;
        rowEnd = row;
        break;
      }
      case "data":
      default:
        // A headerless table begins its data body at r1. A header-only table
        // has no body and therefore resolves to #REF!.
        if (r2 < firstDataRow) return ERRORS.REF;
        rowStart = firstDataRow;
        rowEnd = r2;
        break;
    }

    const from = `${columnToLetters(colStart)}${rowStart + 1}`;
    const to = `${columnToLetters(colEnd)}${rowEnd + 1}`;
    return { from, to };
  };
}

function absoluteTableCellReference(
  reference: string,
  relativeRow = false,
): string | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(reference);
  return match ? `$${match[1]}${relativeRow ? "" : "$"}${match[2]}` : null;
}

function formulaSheetQualifier(sheetName: string): string {
  return `${formatSheetNameForFormula(sheetName)}!`;
}

/**
 * Convert every reference to one soon-to-be-removed ListObject into an absolute
 * A1 reference, matching Excel's Convert to Range contract. The scanner reuses
 * the formula engine's structured-selector parser and resolver, never touches
 * string literals or similarly-prefixed names, and returns `null` instead of
 * guessing when a selector cannot be represented as one contiguous A1 box.
 */
export function convertRemovedTableReferencesToA1(
  formula: string,
  table: TableDefinition,
  currentCell: { row: number; col: number } | undefined,
  tableSheetName: string,
  formulaSheetName?: string,
  options: { thisRowRelative?: boolean } = {},
): string | null {
  if (!formula || !table.name.trim()) return formula;
  const resolver = makeTableRefResolver([table]);
  const tableName = table.name.trim();
  const tableLower = tableName.toLowerCase();
  const sameSheet =
    formulaSheetName !== undefined &&
    formulaSheetName.toLowerCase() === tableSheetName.toLowerCase();
  const currentInsideRemovedTable = Boolean(
    currentCell &&
      currentCell.row >= table.range.r1 &&
      currentCell.row <= table.range.r2 + (table.hasTotals ? 1 : 0) &&
      currentCell.col >= table.range.c1 &&
      currentCell.col <= table.range.c2,
  );
  const qualifier = sameSheet ? "" : formulaSheetQualifier(tableSheetName);
  // Table names take any script's letters (新Sales), so a name boundary is
  // any letter or digit, not only ASCII ones (O008).
  const nameCharacter = (value: string | undefined) =>
    Boolean(value && /[\p{L}\p{N}_.$\\]/u.test(value));
  const replacement = (
    explicitTableName: string | null,
    spec: StructuredSpec,
  ): string | null => {
    const resolved = resolver(explicitTableName, spec, currentCell);
    if (isTableRefError(resolved)) return null;
    const relativeThisRow =
      spec.band === "thisrow" && options.thisRowRelative === true;
    const from = absoluteTableCellReference(resolved.from, relativeThisRow);
    const to = absoluteTableCellReference(resolved.to, relativeThisRow);
    if (!from || !to) return null;
    return `${qualifier}${from}${from === to ? "" : `:${to}`}`;
  };

  let output = "";
  let index = 0;
  let inString = false;
  while (index < formula.length) {
    const character = formula[index];
    if (character === '"') {
      if (inString && formula[index + 1] === '"') {
        output += '""';
        index += 2;
        continue;
      }
      inString = !inString;
      output += character;
      index += 1;
      continue;
    }
    if (inString) {
      output += character;
      index += 1;
      continue;
    }
    // A quoted sheet name ('Sales East'!A1, 'Sales'!A1) names a worksheet,
    // never the table: copy it whole, '' escapes included (O009).
    if (character === "'") {
      let end = index + 1;
      while (end < formula.length) {
        if (formula[end] === "'") {
          if (formula[end + 1] === "'") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      output += formula.slice(index, end);
      index = end;
      continue;
    }

    const previous = index > 0 ? formula[index - 1] : undefined;
    const selectorStart = index + tableName.length;
    const exactTableName =
      !nameCharacter(previous) &&
      previous !== "!" &&
      formula.slice(index, selectorStart).toLowerCase() === tableLower;
    if (exactTableName && formula[selectorStart] === "[") {
      const selected = readStructuredSelector(formula, selectorStart);
      if (!selected) return null;
      const rawSelector = formula.slice(selectorStart, selected.next);
      if ((rawSelector.match(/#(?:data|headers|totals|all|this row)/gi) ?? []).length > 1) {
        return null;
      }
      const converted = replacement(tableName, selected.spec);
      if (!converted) return null;
      output += converted;
      index = selected.next;
      continue;
    }
    if (
      exactTableName &&
      !nameCharacter(formula[selectorStart]) &&
      formula[selectorStart] !== "(" &&
      formula[selectorStart] !== "!"
    ) {
      const converted = replacement(tableName, {
        band: "data",
        column: null,
      });
      if (!converted) return null;
      output += converted;
      index = selectorStart;
      continue;
    }

    if (character === "[") {
      const selected = readStructuredSelector(formula, index);
      if (!selected) return null;
      const after = formula[selected.next];
      const rawSelector = formula.slice(index, selected.next);
      const looksExternal = Boolean(after && /[A-Za-z_'!]/.test(after));
      if (!sameSheet || !currentInsideRemovedTable || looksExternal) {
        output += rawSelector;
        index = selected.next;
        continue;
      }
      if ((rawSelector.match(/#(?:data|headers|totals|all|this row)/gi) ?? []).length > 1) {
        return null;
      }
      const converted = replacement(null, selected.spec);
      if (!converted) return null;
      output += converted;
      index = selected.next;
      continue;
    }

    output += character;
    index += 1;
  }
  return output;
}

// A single (unquoted) A1 cell ref including an optional sheet-qualifying
// prefix. Used by the sheet-prefix reader to find where a qualified ref ends.
const QUALIFIED_CELL_TAIL_RE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/;

/**
 * The result of reading a sheet-name prefix that terminates in `!`: the parsed
 * sheet name(s) and the index just AFTER the `!`. `end` (the 3D span terminus)
 * is set only for the `Sheet1:Sheet3!` form.
 */
interface SheetPrefix {
  sheet: string;
  end?: string;
  next: number; // index of the first char after the closing "!"
  /** Set when the qualified target is a whole column / row, not an A1 cell. */
  wholeAxis?: { token: Token; next: number };
  /**
   * Set when the qualified target is a DEFINED NAME (`Sheet1!Budget`,
   * `'My Sheet'!Rate`) — Excel's way to address a sheet-scoped name from any
   * sheet. `next` is the index just after the name.
   */
  definedName?: { name: string; next: number };
  /** Set when the qualified target is the #REF! Excel writes in place of a
   * deleted range (`Sheet1!#REF!`); `next` is the index after it. */
  refError?: { next: number };
  /** Set for a span whose ends are quoted apart (`'A':'B'!B1`), which is no
   * reference to Excel: the whole text reads #NAME? (X-179). */
  nameError?: { next: number };
}

/**
 * Whether one character (a whole code point, so an astral CJK ideograph is one
 * character) may appear in a sheet name written WITHOUT quotes: a letter of any
 * script or `_`, then letters, combining marks, digits, `_` and `.` — Excel
 * stores `=回款!C4`, `=Données!A1` and `=Лист1!A1` unquoted in .xlsx. The set
 * lives in ./name-syntax, which defined names share (plus a leading `\`).
 */
const isBareSheetNameCharacter = isBareNameCharacter;

/**
 * Whether `name` is spelled entirely in the unquoted sheet-name charset, so a
 * reader meeting `name!` takes it as a sheet (`回款!C4`, `Lists!$A$1`). A
 * writer additionally quotes names that are ALSO a cell, R1C1 ref or boolean
 * — see {@link sheetNameNeedsQuotes}; once `!` follows, those read unambiguously.
 */
export const isBareSheetName = isBareNameSpelling;

/**
 * Read ONE sheet name starting at `src[i]`. A name is either:
 *   - quoted: `'…'` where an embedded quote is doubled (`''`), so the name may
 *     contain spaces, `!`, `:` and most punctuation; or
 *   - bare: a letter (of any script) or `_`, then letters, combining marks,
 *     digits, `_` and `.` — see {@link isBareSheetNameCharacter}. A leading
 *     digit means this is a number, not a name.
 * Returns the decoded name + the next index, or null if no name starts here.
 * Reading a name commits to nothing: the caller only treats it as a sheet when
 * `!` follows, so a function, defined or table name is never swallowed.
 */
function readSheetName(
  src: string,
  i: number,
  digitFirst = false,
): { name: string; next: number } | null {
  const n = src.length;
  if (src[i] === "'") {
    let name = "";
    let j = i + 1;
    while (j < n) {
      if (src[j] === "'") {
        if (src[j + 1] === "'") {
          name += "'";
          j += 2;
          continue;
        }
        // closing quote
        return { name, next: j + 1 };
      }
      name += src[j++];
    }
    return null; // unterminated quote
  }
  const first = characterAt(src, i);
  if (!first || !isBareSheetNameCharacter(first, !digitFirst)) return null;
  let j = i + first.length;
  while (j < n) {
    const character = characterAt(src, j);
    if (!isBareSheetNameCharacter(character, false)) break;
    j += character.length;
  }
  return { name: src.slice(i, j), next: j };
}

/**
 * Try to read a sheet-qualifying prefix at `src[i]`: `Sheet2!`, `'My Sheet'!`,
 * or the 3D `Sheet1:Sheet3!`. Succeeds ONLY when a name (or name:name pair) is
 * immediately followed by `!` AND the text after the `!` begins with a valid A1
 * cell ref or is a whole column / row — otherwise this is not a sheet
 * reference (e.g. a bare name, a string
 * literal, or `a:b` slicing) and the caller falls back to the ordinary token
 * grammar. Returns the parsed prefix or null.
 */
function readSheetPrefix(src: string, i: number, digitFirst = false): SheetPrefix | null {
  const first = readSheetName(src, i, digitFirst);
  if (!first) return null;
  let sheet = first.name;
  let end: string | undefined;
  let pos = first.next;
  // Optional 3D span: a second name after a ":" (the colon must be BETWEEN two
  // names and BEFORE the "!", distinguishing it from a range colon like A1:B2).
  if (src[pos] === ":") {
    const second = readSheetName(src, pos + 1);
    if (!second) return null;
    if ((src[i] === "'" || src[pos + 1] === "'") && src[second.next] === "!") {
      // Excel quotes a 3-D span as one pair ('Q 1:Q 3'!A1); two names quoted
      // apart and joined by ":" are not a reference, and Excel shows #NAME?
      // (S11_tdf163554's 'time (pnrst)':'time (misc) - last'!B1:B1, X-179).
      const tail = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?/.exec(src.slice(second.next + 1));
      if (!tail) return null;
      return { sheet, next: second.next + 1, nameError: { next: second.next + 1 + tail[0].length } };
    }
    end = second.name;
    pos = second.next;
  } else if (src[i] === "'" && sheet.includes(":")) {
    // Excel quotes a 3D span as ONE pair — `'Q 1:Q 3'!A1`, the MS-XLSX
    // sheet-range form — whenever either end needs quotes. A sheet name cannot
    // contain ":" in Excel, so the colon inside the quotes is the span.
    const colon = sheet.indexOf(":");
    end = sheet.slice(colon + 1);
    sheet = sheet.slice(0, colon);
    if (sheet === "" || end === "") return null;
  }
  if (src[pos] !== "!") return null;
  pos += 1;
  // The qualified target must be a real A1 cell ref, or a whole column / row
  // (`Sheet2!A:A`, `Sheet2!$1:$3`); otherwise back off so we don't swallow e.g.
  // a name followed by "!" that is not a reference.
  const prefix: SheetPrefix =
    end !== undefined ? { sheet, end, next: pos } : { sheet, next: pos };
  if (QUALIFIED_CELL_TAIL_RE.test(src.slice(pos))) return prefix;
  // `Sheet1!#REF!`: a reference whose cells were deleted -- the error value
  // #REF! (a defined name `sale_1` = Sheet1!#REF! reads #REF!, P26_49783,
  // X-120).
  if (src.slice(pos, pos + 5).toUpperCase() === "#REF!") return { ...prefix, refError: { next: pos + 5 } };
  const wholeAxis =
    _matchWholeAxisRef(src, pos, "column") ?? _matchWholeAxisRef(src, pos, "row");
  if (wholeAxis) return { ...prefix, wholeAxis };
  // `Sheet1!Budget` — a defined name read in that sheet's scope (Microsoft,
  // "Names in formulas": a worksheet-level name is referred to from another
  // sheet as Sheet1!Budget_FY08). Never a 3D span, a call or a table.
  if (end !== undefined) return null;
  const nameStart = characterAt(src, pos);
  if (!nameStart || !isDefinedNameCharacter(nameStart, true)) return null;
  let nameEnd = pos + nameStart.length;
  while (nameEnd < src.length) {
    const character = characterAt(src, nameEnd);
    if (!isDefinedNameCharacter(character, false)) break;
    nameEnd += character.length;
  }
  if (src[nameEnd] === "(" || src[nameEnd] === "[" || src[nameEnd] === "!") return null;
  return { ...prefix, definedName: { name: src.slice(pos, nameEnd), next: nameEnd } };
}

/**
 * Split a sheet-qualified defined-name spelling (`Sheet1!Budget`,
 * `'My Sheet'!Rate`) into its unquoted sheet and bare name; null for a plain
 * name. A defined name can never contain `!`, so only the qualified form does.
 */
export function splitQualifiedName(raw: string): { sheet: string; name: string } | null {
  if (!raw.includes("!")) return null;
  const prefix = readSheetPrefix(raw, 0);
  if (!prefix?.definedName || prefix.definedName.next !== raw.length) return null;
  return { sheet: prefix.sheet, name: prefix.definedName.name };
}

export function tokenize(src: string): Token[] | null {
  const tokens: Token[] = [];
  let i = 0;
  const n = src.length;
  const isDigit = (c: string) => c >= "0" && c <= "9";

  // Emit a sheet-qualified ref token: the prefix has been parsed; read the A1
  // cell ref that follows (its length is bounded by QUALIFIED_CELL_TAIL_RE) and
  // push a single `ref` token carrying the (unquoted) sheet name(s) on it. The
  // colon-range part (A1:B2) stays as separate ref/colon tokens, so the parser
  // builds the cross-sheet range exactly as it builds a same-sheet one. A whole
  // column / row target is one `wholeref` token carrying the same sheet fields.
  const emitSheetRef = (prefix: SheetPrefix, start: number): number => {
    // `Sheet1!Budget`: one ident token spelled exactly as written, so every
    // rewrite that re-joins token values reproduces it byte-for-byte; the name
    // resolver splits the sheet back off (see splitQualifiedName).
    if (prefix.definedName) {
      tokens.push({ type: "ident", value: src.slice(start, prefix.definedName.next) });
      return prefix.definedName.next;
    }
    if (prefix.refError) {
      // Spelled as written, read as #REF!.
      tokens.push({ type: "error", value: src.slice(start, prefix.refError.next), errorCode: "#REF!" });
      return prefix.refError.next;
    }
    if (prefix.nameError) {
      tokens.push({ type: "error", value: src.slice(start, prefix.nameError.next), errorCode: "#NAME?" });
      return prefix.nameError.next;
    }
    if (prefix.wholeAxis) {
      tokens.push(
        prefix.end !== undefined
          ? { ...prefix.wholeAxis.token, sheet: prefix.sheet, endSheet: prefix.end }
          : { ...prefix.wholeAxis.token, sheet: prefix.sheet },
      );
      return prefix.wholeAxis.next;
    }
    const tail = QUALIFIED_CELL_TAIL_RE.exec(src.slice(prefix.next))!;
    const cell = tail[0];
    const next = prefix.next + cell.length;
    if (prefix.end !== undefined) {
      tokens.push({ type: "ref", value: cell, sheet: prefix.sheet, endSheet: prefix.end });
      return next;
    }
    // `Sheet2!A1#` — the block spilled from Sheet2!A1, the same operator an
    // unqualified anchor takes. (A 3D span above has no spill form.)
    if (src[next] === "#") {
      tokens.push({ type: "ref", value: cell, sheet: prefix.sheet, spill: true });
      return next + 1;
    }
    tokens.push({ type: "ref", value: cell, sheet: prefix.sheet });
    return next;
  };

  while (i < n) {
    const c = src[i];
    // `!A1` with no sheet before the "!": in a defined name, the cell on the
    // sheet the name is used on (Excel's sheet-relative reference;
    // S25_forum-mso-en4-145327's PrevSheet = CELL("filename",!A1); X-175).
    // Only where an operand may start, so `Sheet1!A1` keeps its prefix grammar.
    if (c === "!") {
      const before = tokens[tokens.length - 1];
      const tail = QUALIFIED_CELL_TAIL_RE.exec(src.slice(i + 1));
      if (
        tail &&
        !/[A-Za-z0-9_.(]/.test(src[i + 1 + tail[0].length] ?? "") &&
        (before === undefined || before.type === "op" || before.type === "lparen" || before.type === "comma" || before.type === "semicolon")
      ) {
        tokens.push({ type: "ref", value: tail[0] });
        i += 1 + tail[0].length;
        continue;
      }
    }
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i++;
      while (i < n && (src[i] === " " || src[i] === "\t" || src[i] === "\n" || src[i] === "\r")) {
        i++;
      }
      // A space BETWEEN TWO REFERENCES is Excel's intersection operator:
      // SUM(A1:A5 A3:A5) totals only A3:A5. Everywhere else a space is just
      // layout, so this emits the operator on a deliberately narrow shape —
      // the previous token addresses cells, and the next character can begin a
      // reference. `A1 + 1`, `SUM (A1)`, `"a" "b"` and a leading space are all
      // unaffected because they fail one of those two tests.
      //
      // A defined name (`Sales Costs`, `Sales A1:B1`) is an `ident` token, which
      // is also how a function name lexes — so a name is a left operand only
      // when the next character is not `(`: `SUM (A1)` stays a call with layout
      // space, while `Sales A1:B1` intersects (live audit B360). A CLOSING PAREN
      // is covered, because the left operand of an intersection is exactly what
      // a grouping paren produces — MEASURED in WPS: SUM((A1:A2,B1:B2) B1:B2) is
      // 30. That shape was a parse error here.
      const prev = tokens[tokens.length - 1];
      if (
        prev !== undefined &&
        (prev.type === "ref" ||
          prev.type === "tableref" ||
          prev.type === "rparen" ||
          prev.type === "wholeref" ||
          (prev.type === "ident" && src[i] !== "(")) &&
        i < n &&
        // The right operand starts like a reference: `$A$1`, a grouping paren,
        // a cell / column / name / bare sheet word (`A1`, `Sales`, `回款!C4`)
        // or a QUOTED sheet name (`'Sheet 1'!A1`, live audit B360). A digit can
        // only begin a reference as a WHOLE-ROW one (`2:2`), so it is admitted
        // only when it actually is one — MEASURED in WPS: SUM(A:A 2:2) is 2 and
        // SUM(A:C 2:2) is 222.
        (src[i] === "$" ||
          src[i] === "(" ||
          src[i] === "'" ||
          isDefinedNameCharacter(characterAt(src, i), true) ||
          _matchWholeAxisRef(src, i, "row") !== null)
      ) {
        tokens.push({ type: "isect", value: " " });
      }
      continue;
    }
    if (c === '"') {
      let s = "";
      i++;
      while (i < n) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i++];
      }
      tokens.push({ type: "string", value: s });
      continue;
    }
    if (c === "'") {
      // A single quote only ever opens a QUOTED SHEET NAME ('My Sheet'!A1). It
      // must terminate in `!` + a valid A1 ref; if not, it is malformed (Excel
      // has no other use for a bare leading quote in a formula) -> reject.
      const prefix = readSheetPrefix(src, i);
      if (!prefix) return null;
      i = emitSheetRef(prefix, i);
      continue;
    }
    if (isDigit(c)) {
      // A sheet whose name starts with a digit, unquoted (`1_Busquedas!G5`,
      // LibreOffice's spelling): a reference, or a row put in that sheet does
      
      const prefix = src.includes("!", i) ? readSheetPrefix(src, i, true) : null;
      if (prefix) {
        i = emitSheetRef(prefix, i);
        continue;
      }
    }
    if (isDigit(c) || (c === "." && isDigit(src[i + 1]))) {
      // `1:1` / `$2:$4` — a WHOLE-ROW reference. Digits reach the tokenizer as a
      // number, so the row form has to be recognised here before the numeric
      // scan swallows them. The column form is in the word branch below.
      const wholeRow = _matchWholeAxisRef(src, i, "row");
      if (wholeRow) {
        tokens.push(wholeRow.token);
        i = wholeRow.next;
        continue;
      }
      let num = "";
      while (i < n && (isDigit(src[i]) || src[i] === ".")) num += src[i++];
      // scientific notation
      if (src[i] === "e" || src[i] === "E") {
        num += src[i++];
        if (src[i] === "+" || src[i] === "-") num += src[i++];
        while (i < n && isDigit(src[i])) num += src[i++];
      }
      tokens.push({ type: "number", value: num });
      continue;
    }
    if (c === "#") {
      const rest = src.slice(i);
      const match = ERROR_LITERALS.find((code) => rest.startsWith(code));
      if (!match) return null;
      // A reference whose sheet was deleted keeps its address after the error
      // (Excel writes =#REF!A1, =SUM(#REF!A1:B2)). The address belongs to the
      // one invalid reference, not a second operand, so it stays in the token's
      // spelling and every rewrite reproduces it byte-for-byte.
      const tail =
        match === ERRORS.REF.code
          ? (DELETED_SHEET_REF_TAIL.exec(rest.slice(match.length))?.[0] ?? "")
          : "";
      tokens.push({
        type: "error",
        value: match + tail,
        ...(tail ? { errorCode: match } : {}),
      });
      i += match.length + tail.length;
      continue;
    }
    if (c === "$" || isDefinedNameCharacter(characterAt(src, i), true)) {
      // A sheet-qualified ref (Sheet2!A1 / Sheet1:Sheet3!A1 / 回款!C4) begins
      // with a bare sheet name. Try that FIRST: it only matches when the name is
      // followed by `!` (or `:Name!`) AND a valid A1 ref, so an ordinary function
      // name, defined name, or same-sheet ref never triggers it (no regression).
      // A `$` or `\` can't start a sheet name, so `$A$1` falls through unchanged.
      if (c !== "$") {
        const prefix = readSheetPrefix(src, i);
        if (prefix) {
          i = emitSheetRef(prefix, i);
          continue;
        }
      }
      // The word is a cell / column / row reference, a function or a defined
      // name — one run of name characters (./name-syntax, the set the Name
      // Manager validates against: `金额`, `\Rate`, `Données`) plus `$`.
      const wordStart = i;
      i += characterAt(src, i).length;
      while (i < n) {
        const character = characterAt(src, i);
        if (character !== "$" && !isDefinedNameCharacter(character, false)) break;
        i += character.length;
      }
      const word = src.slice(wordStart, i);
      // A NAME immediately followed by `[` is a structured (ListObject) ref:
      // `Table1[Amount]`, `Table1[#Totals]`, `Table1[@Col]`. The `[` after a name
      // is otherwise never legal in a formula, so this can't shadow an existing
      // cell ref / function / defined name. (`word` may technically match the A1
      // pattern, but Excel forbids a table named like a cell ref, and the `[`
      // disambiguates regardless.)
      if (src[i] === "[" && word.length > 0) {
        const sel = readStructuredSelector(src, i);
        if (!sel) return null; // malformed bracket -> reject the whole formula
        tokens.push({
          type: "tableref",
          value: word,
          table: word,
          spec: sel.spec,
          raw: `${word}${src.slice(i, sel.next)}`,
        });
        i = sel.next;
        continue;
      }
      // `A:A` / `$B:$D` — a WHOLE-COLUMN reference. `word` stopped at the colon
      // holding just the column letters, which is not a cell ref, so without
      // this it would become an `ident` (a defined name) and fail to parse.
      // `$1:$1` reaches this branch too — a leading `$` is a word character, so
      // the digit branch above never saw it. Try both axes from the same anchor.
      const wholeAxis =
        _matchWholeAxisRef(src, i - word.length, "column") ??
        _matchWholeAxisRef(src, i - word.length, "row");
      if (wholeAxis) {
        tokens.push(wholeAxis.token);
        i = wholeAxis.next;
        continue;
      }
      // Is it a cell ref (letters + digits, optional $)? Letters past XFD
      // are no column: Excel reads "zzz1" as a name (an undefined one is
      // #NAME?, X-306), as it reads rows past 1048576.
      if (cellRefInGrid(word)) {
        // A trailing `#` is the SPILL-RANGE operator (`A1#`): the ref addresses
        // the whole dynamic-array block anchored at `word`. The `#` is consumed
        // here (so it never starts an error-literal token) and flagged on the
        // ref. Only a plain `A1` anchor takes it — `A1:B2#` / a `#` after a range
        // endpoint is not Excel's anchor form and is left for the error-literal /
        // parser paths (the next `#REF!`-style scan or a trailing-token reject).
        if (src[i] === "#") {
          tokens.push({ type: "ref", value: word, spill: true });
          i += 1;
        } else {
          tokens.push({ type: "ref", value: word });
        }
      } else {
        tokens.push({ type: "ident", value: word });
      }
      continue;
    }
    if (c === "[" && src.startsWith("[0]!", i)) {
      // `[0]!Name`: external book 0 is the workbook itself, so the name is
      // this workbook's own (POI writes it; Excel resolves it as the bare name,
      // P16_56737, X-119).
      i += 4;
      continue;
    }
    if (c === "[") {
      // A bare `[…]` (no table name) is the this-row / same-table form:
      // `[@Amount]`, `[Amount]`, `[#Headers]`. It targets the table that CONTAINS
      // the formula's own cell (resolved against the registry at eval time). A `[`
      // was previously an "unexpected char" (rejected), so this is purely additive.
      const sel = readStructuredSelector(src, i);
      if (!sel) return null;
      tokens.push({
        type: "tableref",
        value: "",
        table: null,
        spec: sel.spec,
        raw: src.slice(i, sel.next),
      });
      i = sel.next;
      continue;
    }
    if (c === "(") {
      tokens.push({ type: "lparen", value: c });
      i++;
      continue;
    }
    if (c === ")") {
      tokens.push({ type: "rparen", value: c });
      i++;
      continue;
    }
    if (c === "{") {
      tokens.push({ type: "lbrace", value: c });
      i++;
      continue;
    }
    if (c === "}") {
      tokens.push({ type: "rbrace", value: c });
      i++;
      continue;
    }
    if (c === ";") {
      tokens.push({ type: "semicolon", value: c });
      i++;
      continue;
    }
    if (c === ",") {
      tokens.push({ type: "comma", value: c });
      i++;
      continue;
    }
    if (c === ":") {
      tokens.push({ type: "colon", value: c });
      i++;
      continue;
    }
    // operators (multi-char first)
    const two = src.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      tokens.push({ type: "op", value: two });
      i += 2;
      continue;
    }
    if ("+-*/^%&=<>".includes(c)) {
      tokens.push({ type: "op", value: c });
      i++;
      continue;
    }
    // `@` is the implicit-intersection operator (Excel 365). Inside `[...]` it is
    // already consumed by the structured-ref decoder above, so a bare `@` here is
    // always the prefix operator; the parser only accepts it in prefix position.
    if (c === "@") {
      tokens.push({ type: "op", value: "@" });
      i++;
      continue;
    }
    // A sheet or defined name that starts with a letter of any script
    // (`回款!C4`, `Лист1!A1`, `金额`) was read by the word branch above.
    return null; // unexpected char
  }
  return tokens;
}

// ---------------------------------------------------------------------------
// AST
// ---------------------------------------------------------------------------

export type Node =
  | { kind: "num"; value: number }
  | { kind: "str"; value: string }
  | { kind: "bool"; value: boolean }
  // A cell reference. `sheet` is the (already-unquoted) sheet name for a
  // cross-sheet reference like Sheet2!A1; absent/undefined means the same
  // (owning) sheet — every pre-existing formula. (A 3D single-cell ref like
  // Sheet1:Sheet3!A1 is modeled as a degenerate `range` with from === to, so a
  // `ref` node is never 3D.)
  | { kind: "ref"; ref: string; sheet?: string }
  // `A1:A5 A3:A5` — the rectangle where two references overlap. Excel writes it
  // with a space and calls it the intersection operator; it is #NULL! when the
  // two do not overlap, which is the only thing #NULL! means.
  | { kind: "isect"; left: Node; right: Node }
  // `(A1:A2,B1:B2)` — Excel's UNION operator: one reference holding several
  // rectangles ("areas"). The comma only means union inside a GROUPING paren;
  // between a function's arguments it stays an argument separator, which is why
  // SUM(A1:A2,B1:B2) and SUM((A1:A2,B1:B2)) are different formulas that happen
  // to agree (both 33, MEASURED). Areas keep their written order and their
  // duplicates: SUM((A1:A2,A1:A2)) is 6, not 3, and AREAS of it is 2.
  | { kind: "union"; parts: Node[] }
  // `A:A` / `$B:$D` / `1:1` — a whole column or whole row. It addresses the
  // sheet's full height (or width), so it resolves against MAX_GRID_ROWS /
  // MAX_GRID_COLUMNS rather than a written endpoint.
  | { kind: "wholeRange"; axis: "column" | "row"; from: string; to: string }
  // A SPILL-RANGE reference: `A1#` — the entire dynamic-array block that spilled
  // from anchor cell `ref`. Resolves to the spilled rectangle (so SUM(A1#) totals
  // it and =A1# re-spills it); a non-anchor / 1×1 anchor degrades to #REF! (Excel).
  // `sheet` is set for another sheet's anchor (`Sheet2!A1#`).
  | { kind: "spill"; ref: string; sheet?: string }
  // A structured / ListObject reference: =SUM(Table1[Amount]), [@Col], …
  // `table` is null for the unqualified this-row form (`[@Col]`) that targets the
  // table containing the formula's own cell. `spec` is the decoded selector.
  | { kind: "tableref"; table: string | null; spec: StructuredSpec }
  | { kind: "name"; name: string }
  | { kind: "error"; code: string }
  | { kind: "range"; from: string; to: string; sheet?: string; endSheet?: string }
  // A range operator whose endpoint is computed: `A1:INDEX(A1:A9,n)`,
  // `OFFSET(A1,1,0):B5`. Each side must resolve to a reference; the result is
  // the rectangle spanning both (Excel's range operator on references).
  | { kind: "rangeExpr"; left: Node; right: Node }
  | { kind: "unary"; op: string; operand: Node }
  | { kind: "binary"; op: string; left: Node; right: Node }
  | { kind: "call"; name: string; args: Node[] }
  // Applying a computed function value to arguments: `callee(args)`, where the
  // callee is any expression that evaluates to a LambdaValue — the inline form
  // `=LAMBDA(x,x+1)(4)` (callee is a `call`) or a parenthesized lambda. (A bare
  // identifier `f(args)` is parsed as a `call` named F and, at evaluation time,
  // re-routed to lambda application when F is bound to a LambdaValue.)
  | { kind: "apply"; callee: Node; args: Node[] }
  // An inline array constant: `{1,2;3,4}`. `rows` is row-major (each inner array
  // is one row's elements). Rows may be ragged on input; evaluation pads short
  // rows to the widest with #N/A (Excel array-constant behaviour).
  | { kind: "array"; rows: Node[][] }
  | { kind: "empty" };

// ---------------------------------------------------------------------------
// Parser (recursive descent with Excel-ish precedence)
// ---------------------------------------------------------------------------

export class Parser {
  private pos = 0;
  constructor(private readonly tokens: Token[]) {}

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }
  private next(): Token | undefined {
    return this.tokens[this.pos++];
  }
  private expect(type: TokenType): Token {
    const t = this.next();
    if (!t || t.type !== type) throw new Error("parse");
    return t;
  }

  parse(): Node {
    const node = this.parseComparison();
    if (this.pos !== this.tokens.length) throw new Error("trailing");
    return node;
  }

  private parseComparison(): Node {
    let left = this.parseConcat();
    while (this.peek()?.type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(this.peek()!.value)) {
      const op = this.next()!.value;
      left = { kind: "binary", op, left, right: this.parseConcat() };
    }
    return left;
  }

  private parseConcat(): Node {
    let left = this.parseAddSub();
    while (this.peek()?.type === "op" && this.peek()!.value === "&") {
      this.next();
      left = { kind: "binary", op: "&", left, right: this.parseAddSub() };
    }
    return left;
  }

  private parseAddSub(): Node {
    let left = this.parseMulDiv();
    while (this.peek()?.type === "op" && ["+", "-"].includes(this.peek()!.value)) {
      const op = this.next()!.value;
      left = { kind: "binary", op, left, right: this.parseMulDiv() };
    }
    return left;
  }

  private parseMulDiv(): Node {
    let left = this.parsePower();
    while (this.peek()?.type === "op" && ["*", "/"].includes(this.peek()!.value)) {
      const op = this.next()!.value;
      left = { kind: "binary", op, left, right: this.parsePower() };
    }
    return left;
  }

  // Excel operator precedence (tightest first): negation (unary -), %, ^, * /.
  // So unary minus binds TIGHTER than ^ — `=-2^2` is `(-2)^2` = 4, not `-(2^2)`.
  // The chain is therefore parseMulDiv → parsePower(^) → parsePercent(%) →
  // parseNegation(-/+/@) → parsePostfixApply.
  private parsePower(): Node {
    let left = this.parsePercent();
    while (this.peek()?.type === "op" && this.peek()!.value === "^") {
      this.next();
      left = { kind: "binary", op: "^", left, right: this.parsePercent() };
    }
    return left;
  }

  private parsePercent(): Node {
    let left = this.parseNegation();
    while (this.peek()?.type === "op" && this.peek()!.value === "%") {
      this.next();
      left = { kind: "unary", op: "%", operand: left };
    }
    return left;
  }

  private parseNegation(): Node {
    if (this.peek()?.type === "op" && (this.peek()!.value === "-" || this.peek()!.value === "+")) {
      const op = this.next()!.value;
      return { kind: "unary", op, operand: this.parseNegation() };
    }
    // `@expr` — implicit intersection: collapse an array/spill operand to its
    // anchor scalar (Excel 365).
    if (this.peek()?.type === "op" && this.peek()!.value === "@") {
      this.next();
      return { kind: "unary", op: "@", operand: this.parseNegation() };
    }
    return this.parseIntersection();
  }

  /**
   * `ref ref` — Excel's intersection operator. It binds TIGHTER than unary
   * minus and looser than the `:` range colon (which `parsePrimary` has already
   * folded into a single range node), which is where Excel puts it.
   */
  private parseIntersection(): Node {
    let expr = this.parsePostfixApply();
    while (this.peek()?.type === "isect") {
      this.next();
      expr = { kind: "isect", left: expr, right: this.parsePostfixApply() };
    }
    return expr;
  }

  /**
   * A primary expression followed by zero or more argument lists: `expr(a,b)(c)`.
   * This is what lets a computed function value be CALLED — the inline LAMBDA
   * form `=LAMBDA(x,x+1)(4)` (primary is a `call` to LAMBDA, then an apply with
   * (4)) and a parenthesized lambda `=(LAMBDA(x,x))(5)`. Each `(…)` after the
   * primary becomes an `apply` node whose callee is the expression so far; the
   * argument list reuses the same grammar as a function call. A bare identifier
   * `f(...)` is still parsed as a `call` by `parsePrimary` (its name is unknown
   * until evaluation) and re-routed to lambda application there.
   */
  private parsePostfixApply(): Node {
    let expr = this.parsePrimary();
    while (this.peek()?.type === "lparen") {
      this.next(); // consume "("
      const args: Node[] = [];
      if (this.peek()?.type !== "rparen") {
        args.push(this.parseArgOrEmpty());
        while (this.peek()?.type === "comma") {
          this.next();
          args.push(this.parseArgOrEmpty());
        }
      }
      this.expect("rparen");
      expr = { kind: "apply", callee: expr, args };
    }
    // `INDEX(...):A9` — a reference-returning call as the LEFT endpoint of the
    // range operator (live audit B096).
    if (expr.kind === "call" && this.peek()?.type === "colon") {
      this.next();
      expr = { kind: "rangeExpr", left: expr, right: this.parsePrimary() };
    }
    return expr;
  }

  /**
   * A reference that a parenthesised group or a spill ref spells, as the LEFT
   * endpoint of the range operator: `F31#:H31`, `(F25,F26):F27`, `(F25):F27`
   * (Excel spans them, LibreOffice LS018; X-312). Anything else is returned
   * as it is.
   */
  private rangeTail(left: Node): Node {
    if (this.peek()?.type !== "colon") return left;
    this.next();
    return { kind: "rangeExpr", left, right: this.parsePrimary() };
  }

  private parsePrimary(): Node {
    const t = this.peek();
    if (!t) throw new Error("eof");
    if (t.type === "number") {
      this.next();
      return { kind: "num", value: Number(t.value) };
    }
    if (t.type === "string") {
      this.next();
      return { kind: "str", value: t.value };
    }
    if (t.type === "wholeref") {
      this.next();
      if (t.sheet !== undefined) {
        // `Sheet2!A:A` / `Sheet2!1:1` / `Sheet1:Sheet3!A:A` is the sheet-qualified
        // range spanning the grid on that axis, so it takes the cross-sheet and
        // 3D range path every `Sheet2!A1:B5` takes — range2D bounds a whole
        // axis by the TARGET sheet's used range. A `wholeRange` node stays
        // same-sheet only; the functions reading it as a reference would
        // otherwise resolve it against the formula's own sheet.
        const node: Node =
          t.axis === "column"
            ? { kind: "range", from: `${t.from!}1`, to: `${t.to!}${MAX_GRID_ROWS}` }
            : {
                kind: "range",
                from: `A${t.from!}`,
                to: `${columnToLetters(MAX_GRID_COLUMNS - 1)}${t.to!}`,
              };
        node.sheet = t.sheet;
        if (t.endSheet !== undefined) node.endSheet = t.endSheet;
        return node;
      }
      return {
        kind: "wholeRange",
        axis: t.axis!,
        from: t.from!,
        to: t.to!,
      };
    }
    if (t.type === "tableref") {
      // A structured / ListObject reference. The tokenizer already decoded the
      // table name + selector, so this is a leaf — it resolves to a range/grid at
      // evaluation time via the EvalContext table resolver.
      this.next();
      return { kind: "tableref", table: t.table ?? null, spec: t.spec! };
    }
    if (t.type === "error") {
      this.next();
      return { kind: "error", code: t.errorCode ?? t.value };
    }
    if (t.type === "lparen") {
      this.next();
      const inner = this.parseComparison();
      // A comma inside a GROUPING paren is Excel's union operator, not an
      // argument separator (that comma is consumed by parseCall/parsePostfixApply,
      // which never reaches here). Nested unions flatten: AREAS(((A1,B1),C1)) is
      // 3 and SUM(((A1,B1),C1)) is 111 — MEASURED in WPS — so a union part that
      // is itself a union contributes its own parts, not one nested area.
      if (this.peek()?.type === "comma") {
        const parts: Node[] = inner.kind === "union" ? [...inner.parts] : [inner];
        while (this.peek()?.type === "comma") {
          this.next();
          const part = this.parseComparison();
          if (part.kind === "union") parts.push(...part.parts);
          else parts.push(part);
        }
        this.expect("rparen");
        return this.rangeTail({ kind: "union", parts });
      }
      this.expect("rparen");
      return this.rangeTail(inner);
    }
    if (t.type === "lbrace") {
      return this.parseArrayLiteral();
    }
    if (t.type === "ref") {
      // A SPILL-RANGE ref (`A1#`) is a leaf addressing the whole spilled block.
      // It is never a function name nor a range endpoint (the `#` already closed
      // the reference), so it short-circuits the call/range grammar below.
      if (t.spill) {
        this.next();
        return this.rangeTail(
          t.sheet !== undefined
            ? { kind: "spill", ref: t.value, sheet: t.sheet }
            : { kind: "spill", ref: t.value },
        );
      }
      // A "ref" immediately followed by "(" is actually a function whose name
      // happens to match the A1 pattern (e.g. LOG10). A SHEET-QUALIFIED ref
      // (it carries a `sheet`) is never a function name, so the call shortcut is
      // only taken for an unqualified ref.
      if (t.sheet === undefined && this.tokens[this.pos + 1]?.type === "lparen") {
        this.next();
        return this.parseCall(t.value.toUpperCase());
      }
      this.next();
      // range? The sheet prefix lives on the LEFT endpoint (Sheet1!A1:B5 or the
      // 3D Sheet1:Sheet3!A1:B5); the right endpoint is a bare ref. Excel does not
      // allow a sheet prefix on the right side, so a right token carrying its own
      // `sheet` is a malformed reference -> reject.
      if (this.peek()?.type === "colon") {
        this.next();
        // `A1:INDEX(...)` — the right endpoint is a call that yields a
        // reference (live audit B096); a plain ref is the ordinary range.
        const rightTok = this.peek();
        const rightIsCall =
          ((rightTok?.type === "ref" || rightTok?.type === "ident") &&
            this.tokens[this.pos + 1]?.type === "lparen") ||
          // `A1:(A2,A3)` / `A1:(A2)`: a parenthesised endpoint (X-312).
          rightTok?.type === "lparen" ||
          // `A1:B1#`: a spill ref as the right endpoint.
          (rightTok?.type === "ref" && rightTok.spill === true);
        if (rightIsCall) {
          const left: Node = { kind: "ref", ref: t.value };
          if (t.sheet !== undefined) left.sheet = t.sheet;
          return { kind: "rangeExpr", left, right: this.parsePrimary() };
        }
        const to = this.expect("ref");
        // `Data!D3:Data!D99`: the same sheet named on both ends is the ordinary range,
        // and Excel keeps it spelled so in the file (R168-X06). Two sheets are no range.
        if (to.sheet !== undefined && to.sheet.toLowerCase() !== t.sheet?.toLowerCase()) {
          throw new Error("range-sheet-on-rhs");
        }
        const node: Node = { kind: "range", from: t.value, to: to.value };
        if (t.sheet !== undefined) node.sheet = t.sheet;
        if (t.endSheet !== undefined) node.endSheet = t.endSheet;
        return node;
      }
      // A 3D single-cell ref (Sheet1:Sheet3!A1) is modeled as a degenerate range
      // (from === to) so it ALWAYS flows through the 3D range machinery — a
      // single cell read across the whole sheet span, which is the only sensible
      // 3D semantic. (A plain single-sheet/2D-prefixed ref stays a `ref` node so
      // every existing scalar path is unchanged.)
      if (t.endSheet !== undefined) {
        return {
          kind: "range",
          from: t.value,
          to: t.value,
          sheet: t.sheet!,
          endSheet: t.endSheet,
        };
      }
      const node: Node = { kind: "ref", ref: t.value };
      if (t.sheet !== undefined) node.sheet = t.sheet;
      return node;
    }
    if (t.type === "ident") {
      this.next();
      const upper = t.value.toUpperCase();
      // An ident immediately followed by "(" is a function call — including the
      // zero-arg TRUE()/FALSE() forms Excel allows. A bare TRUE/FALSE is a
      // boolean literal; anything else is a defined name (named range).
      if (this.peek()?.type === "lparen") {
        return this.parseCall(upper);
      }
      if (upper === "TRUE") return { kind: "bool", value: true };
      if (upper === "FALSE") return { kind: "bool", value: false };
      return { kind: "name", name: stripFormulaParameterPrefix(t.value) };
    }
    throw new Error("unexpected");
  }

  private parseCall(rawName: string): Node {
    const name = stripFutureFunctionPrefix(rawName);
    this.expect("lparen");
    const args: Node[] = [];
    if (this.peek()?.type !== "rparen") {
      args.push(this.parseArgOrEmpty());
      while (this.peek()?.type === "comma") {
        this.next();
        args.push(this.parseArgOrEmpty());
      }
    }
    this.expect("rparen");
    return { kind: "call", name, args };
  }

  /**
   * One call argument, or an `empty` node when the slot is omitted — Excel/Sheets
   * allow `SUM(1,,3)` / `TEXTSPLIT(t,c,,TRUE)` where a missing middle argument is
   * a blank value (NOT a syntax error). An omitted slot is one where the next
   * token is the argument separator or the closing paren.
   */
  private parseArgOrEmpty(): Node {
    const t = this.peek()?.type;
    if (t === "comma" || t === "rparen") return { kind: "empty" };
    return this.parseComparison();
  }

  /**
   * An inline array constant: `{1,2;3,4}` — `,` separates columns within a row,
   * `;` separates rows. Each element is a full (comparison-level) expression, so
   * `{-1,2;TRUE,"x"}` and `{1+1,2}` parse; ragged rows are tolerated here and
   * padded at evaluation time. An empty `{}` has no first element and throws
   * (→ #ERROR!), matching Excel's rejection of an empty array constant.
   */
  private parseArrayLiteral(): Node {
    this.expect("lbrace");
    const rows: Node[][] = [];
    let row: Node[] = [this.parseComparison()];
    for (;;) {
      const t = this.peek()?.type;
      if (t === "comma") {
        this.next();
        row.push(this.parseComparison());
      } else if (t === "semicolon") {
        this.next();
        rows.push(row);
        row = [this.parseComparison()];
      } else {
        break;
      }
    }
    rows.push(row);
    this.expect("rbrace");
    return { kind: "array", rows };
  }
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

/**
 * The outcome of a GETPIVOTDATA lookup through {@link EvalContext.resolvePivotData}.
 * `"value"` carries the matched aggregate (null = an empty group cell); the engine
 * maps `"no-match"` to #N/A and `"not-pivot"` to #REF!. Structurally identical to
 * the PivotDataResult produced by lib/spreadsheet/pivot's makePivotDataResolver,
 * declared here so the engine never imports the pivot module (pivot.ts imports the
 * engine, so the dependency stays one-way).
 */
export type PivotDataLookupResult =
  | { kind: "value"; value: number | string | null }
  | { kind: "no-match" }
  | { kind: "not-pivot" };

export interface EvalContext {
  /** Resolve a single cell reference to its computed value. */
  resolveRef(ref: string): CellValue;
  /** Resolve a range "A1:B2" to a flat array of computed values. */
  resolveRange(from: string, to: string): CellValue[];
  /** Resolve a range "A1:B2" to a 2D grid (rows x cols) of computed values. */
  resolveRange2D(from: string, to: string): CellValue[][];
  /** Value-only aggregate iteration may omit implicit whole-axis blanks.
   * The reference resolver itself still returns its complete address extent. */
  aggregateValueBounds?(box: RefBox): RefBox | FormulaError;
  /**
   * Resolve a CROSS-SHEET single cell ("Sheet2!A1") to its computed value.
   * `sheetName` is the already-unquoted target sheet name (case-insensitive
   * match is the resolver's job). An unknown sheet must return ERRORS.REF.
   * Optional: a single-sheet context omits it, in which case any sheet-qualified
   * ref evaluates to #REF! (it has no other sheet to reach).
   */
  resolveSheetRef?(sheetName: string, ref: string): CellValue;
  /** Resolve a cross-sheet range ("Sheet2!A1:B2") to a 2D grid. Unknown sheet -> [[REF]]. */
  resolveSheetRange2D?(sheetName: string, from: string, to: string): CellValue[][];
  /**
   * Resolve a 3D range — the SAME cell/range read across a CONSECUTIVE span of
   * sheets (Sheet1:Sheet3!A1). Returns one 2D grid PER sheet in the span (in
   * workbook order), which the aggregate/flatten path concatenates. An unknown
   * endpoint or a reversed/empty span -> a single [[REF]] grid. Optional.
   */
  resolveSheet3DRange?(
    startSheet: string,
    endSheet: string,
    from: string,
    to: string,
  ): CellValue[][][];
  /**
   * Resolve a STRUCTURED (ListObject) reference — `Table1[Amount]`, `[@Col]`,
   * `Table1[#Totals]`, … — to the A1 box it addresses. A box on another sheet
   * (table names are workbook-scoped) carries that sheet's name in `sheet`.
   * `tableName` is null for the unqualified this-row form (`[@Col]`); the resolver
   * uses `currentCell` to find the table that contains the formula's own cell.
   * Returns a resolved A1 box (`{ from, to }`, e.g. `{from:"B2",to:"B10"}`) or a
   * FormulaError: ERRORS.REF for an unknown table / column / band-without-totals,
   * ERRORS.VALUE for a `[@Col]` whose formula sits OUTSIDE the table's rows.
   * Optional: a context without tables omits it, in which case any structured ref
   * evaluates to #REF! (there is no table registry to consult). Same-cell matching
   * is case-insensitive (the resolver's job).
   */
  resolveTableRef?(
    tableName: string | null,
    spec: StructuredSpec,
    currentCell: { row: number; col: number } | undefined,
  ): TableRefResult;
  /**
   * Resolve a GETPIVOTDATA lookup. `anchor` is the cell the `pivot_table`
   * argument addressed (0-indexed); `dataField` is the value-field name
   * ("Sum of Sales" or the bare source field); `criteria` are the (field,item)
   * pairs that constrain the group. Returns the matched aggregate (a number, or
   * null for an empty group), `"no-match"` when no qualifying cell / unknown
   * field / item (engine: #N/A), or `"not-pivot"` when the anchor is not inside
   * any pivot's output (engine: #REF!). Optional: a context without pivots omits
   * it, in which case every GETPIVOTDATA is #REF! (no pivot registry to consult).
   * The pivot math lives in lib/spreadsheet/pivot (makePivotDataResolver); this
   * is the seam the host wires it through, mirroring {@link resolveTableRef}.
   */
  resolvePivotData?(
    anchor: { row: number; col: number; sheet?: string },
    dataField: string,
    criteria: ReadonlyArray<{ field: string; item: string }>,
  ): PivotDataLookupResult;
  /**
   * Resolve a SPILL-RANGE reference `A1#` to the FULL dynamic-array block that
   * spilled from anchor cell `ref` — the live `CellValue[][]` grid (`[[1],[2],
   * [3]]` for a 3×1 spill). The host evaluates the anchor's formula and returns
   * its array result directly; the spilled non-anchor cells are NOT in cell
   * storage (they are an overlay), so the block must come from the anchor's
   * array, not from re-reading a box. When `ref` is NOT a live spill anchor — it
   * holds a constant, a scalar formula (a 1×1 result), or nothing — Excel's
   * `A1#` yields #REF! ("the spill range is empty"), so the resolver returns
   * null and the engine produces ERRORS.REF. Optional: a context without
   * dynamic-array materialization omits it, in which case every `A1#` is #REF!.
   */
  resolveSpillArray?(ref: string): CellValue[][] | null;
  /**
   * {@link resolveSpillArray} for an anchor on ANOTHER sheet (`Sheet2!A1#`),
   * same contract; null for an unknown sheet too. A context without it makes
   * every other-sheet `A1#` #REF!.
   */
  resolveSheetSpillArray?(sheetName: string, ref: string): CellValue[][] | null;
  /**
   * The rectangular EXTENT (top-left + height×width) of the spill anchored at
   * `ref`, used when `A1#` appears as the reference argument of OFFSET/INDEX/
   * COLUMN/ROW (which compose by coordinates). null when `ref` is not a live
   * spill anchor. Companion to {@link resolveSpillArray}; a host that wires one
   * normally wires both. Optional — absent contexts make `A1#` in a reference
   * argument #REF!.
   */
  resolveSpillBox?(ref: string): { from: string; to: string } | null;
  /**
   * Resolve a defined name (named range) to its A1 reference string — a single
   * cell ("B2") or a range ("B2:B10") — or null/undefined when the name is not
   * defined. Optional: contexts without named ranges may omit it, in which case
   * any bare name evaluates to #NAME?.
   */
  resolveName?(name: string): string | null;
  /**
   * Resolve a defined name in ANOTHER sheet's scope — `Sheet1!Budget` reads
   * Sheet1's local Budget (falling back to a workbook-level Budget, the way that
   * sheet itself would). Same contract as {@link resolveName}; a bare cell /
   * range it returns addresses THAT sheet. Optional: without it every
   * sheet-qualified name is #NAME?.
   */
  resolveSheetName?(sheetName: string, name: string): string | null;
  /**
   * The set of named-FORMULA names currently being expanded (uppercased), used
   * to break a name→…→name cycle (a named formula that, directly or transitively,
   * refers back to itself) — re-entering a name in the set yields #REF! instead
   * of recursing forever. Evaluation-scope only; never persisted. A plain cell /
   * range named-range never enters this path, so it stays untouched.
   */
  resolvingNames?: Set<string>;
  /**
   * Lexical bindings introduced by an enclosing LET(...) during evaluation —
   * consulted BEFORE named ranges. Evaluation-scope only; never persisted.
   */
  letBindings?: ReadonlyMap<string, EngineValue>;
  /**
   * The current LAMBDA call-stack depth. Each lambda invocation increments it;
   * when it exceeds {@link MAX_LAMBDA_DEPTH} the call yields #NUM! instead of
   * recursing further. This bounds a self-recursive lambda (e.g. =LET(f,
   * LAMBDA(n, f(n)), f(1))) that, unlike a cell→cell cycle, recurses entirely
   * within ONE cell's evaluation and so is invisible to the per-cell cycle
   * guard. Absent on a fresh context = depth 0.
   */
  lambdaDepth?: number;
  /**
   * The cell currently being evaluated (0-indexed row/col), used by the no-arg
   * COLUMN()/ROW() forms. Optional: contexts that don't supply it make a no-arg
   * COLUMN()/ROW() return #VALUE! (the with-reference forms work regardless).
   */
  cell?: { row: number; col: number };
  /**
   * Row visibility for SUBTOTAL (Excel parity). The evaluator is per-sheet, so
   * the callbacks take a 0-indexed row only. `isRowHiddenByFilter` marks rows
   * hidden by an AutoFilter — SUBTOTAL skips those for EVERY function_num
   * (1-11 and 101-111). `isRowManuallyHidden` marks rows hidden by hand
   * (Excel's "Hide Rows"); only the 101-111 family skips those. Both optional:
   * contexts without them treat every row as visible.
   */
  isRowHiddenByFilter?(row: number): boolean;
  isRowManuallyHidden?(row: number): boolean;
  /**
   * The raw formula text ("=…") of the cell at (row, col), or null when the
   * cell holds no formula. SUBTOTAL uses it to EXCLUDE cells whose own formula
   * contains a SUBTOTAL/AGGREGATE call (Excel ignores nested subtotals so a
   * grand total over running totals never double-counts). Optional: contexts
   * without it skip the nested-subtotal exclusion.
   */
  getCellFormula?(row: number, col: number): string | null;
  /**
   * The raw formula text of a cell on ANOTHER sheet — the cross-sheet companion
   * of {@link getCellFormula}, used by FORMULATEXT / ISFORMULA / CELL when their
   * reference is sheet-qualified (`FORMULATEXT(Sheet2!A1)`). Returns the "=…"
   * text, null when the cell holds no formula, or null for an unknown sheet.
   * Optional: a single-sheet context omits it (a sheet-qualified ref then has no
   * formula to read).
   */
  getSheetCellFormula?(sheetName: string, row: number, col: number): string | null;
  /**
   * Whether the formula of the cell at (row, col) is withheld from display —
   * its sheet is protected and the cell is formatted Hidden. The formula still
   * calculates; only FORMULATEXT, which returns what the formula bar shows,
   * answers #N/A for it. The cross-sheet companion answers a sheet-qualified
   * reference. Optional: without them every formula is displayable.
   */
  isCellFormulaHidden?(row: number, col: number): boolean;
  isSheetCellFormulaHidden?(sheetName: string, row: number, col: number): boolean;
  /**
   * The phonetic reading (furigana) of the text cell at (row, col), or null
   * when it has none — what PHONETIC returns for it. The cross-sheet companion
   * answers a sheet-qualified reference. Optional: without them PHONETIC
   * reads a cell's own text.
   */
  getCellPhonetic?(row: number, col: number): string | null;
  getSheetCellPhonetic?(sheetName: string, row: number, col: number): string | null;
  /**
   * The width of column `col` of the formula's sheet (or of `sheetName`) in
   * characters, as Column Width reads it, or null when the host gave none --
   * what CELL("width") rounds.
   */
  getSheetColumnWidth?(sheetName: string | undefined, col: number): number | null;
  /** Whether that column has its sheet's default width, or null when the host gave none. */
  getSheetColumnWidthIsDefault?(sheetName: string | undefined, col: number): boolean | null;
  /**
   * Another sheet's row visibility, for SUBTOTAL/AGGREGATE over
   * `Sheet2!A1:A10`: the callbacks that sheet registered (its AutoFilter and
   * its hidden rows), or null for an unknown sheet. Absent callbacks mean
   * every row of that sheet is visible.
   */
  sheetVisibility?(sheetName: string): {
    isRowHiddenByFilter?: (row: number) => boolean;
    isRowManuallyHidden?: (row: number) => boolean;
  } | null;
  /**
   * SHEET([value]) support. With no argument SHEET() returns the 1-based number
   * of the sheet the formula lives on (`sheetNumber()`); SHEET("Name") /
   * SHEET(Sheet2!A1) returns the number of the named/referenced sheet. The
   * callback maps a sheet NAME (case-insensitively) to its 1-based workbook
   * position, or null when no such sheet exists (#N/A). Called with no name for
   * the owning sheet. Optional: a single-sheet context omits it, in which case
   * SHEET() is 1 and SHEET(name) of any other name is #N/A.
   */
  sheetNumber?(name?: string): number | null;
  /**
   * SHEETS() support: the total number of sheets in the workbook (SHEETS() with
   * no argument). The ref-spanning form SHEETS(ref) is computed from the
   * reference's own sheet span and does not need this. Optional: a single-sheet
   * context omits it, in which case SHEETS() is 1.
   */
  sheetCount?(): number;
  /** SHEETSNAME() — the name of the workbook's FIRST sheet. */
  firstSheetName?(): string | null;
  /** BOOKNAME() — the workbook's file name, when the host supplies one. */
  workbookName?(): string | null;
  /** The folder the workbook's location names, ending in "/", before
   * "[Book.xlsx]Sheet" in CELL("filename") (X-194); null when the host has
   * no location to give. */
  workbookFolder?(): string | null;
  /** The name of the sheet the formula is evaluated on (CELL("filename")). */
  ownSheetName?(): string;
  /**
   * The workbook reads its serials in the 1904 date system (Excel's "Use 1904
   * date system", SpreadsheetML `workbookPr@date1904`): serial 0 is
   * 1904-01-01. Only the calendar reading changes — DATE/YEAR/WEEKDAY/TEXT,
   * DATEVALUE, TODAY/NOW and the other date functions — never arithmetic.
   * Omitted inherits the evaluation in progress (1900 at the top level).
   */
  date1904?: ExcelDateSystemArg;
  /**
   * Where an IMAGE source's picture stands ({@link WebImageStatus}); omitted
   * inherits the evaluation in progress (no host: every valid source reads
   * as its picture value).
   */
  webImageStatus?: (source: string) => WebImageStatus | null;
  /**
   * The data WEBSERVICE and IMPORTRANGE read, as the host holds it
   * (`external-data`); omitted inherits the evaluation in progress (no host:
   * every such call reads as #BUSY!).
   */
  externalData?: ExternalDataLookup;
  /** Formulas parsed already, by text (a calculation session's, kept across
   * its calculations): a formula found here is not tokenized and parsed
   * again. The parse of a text is the same wherever it stands. */
  parsedFormulas?: Map<string, unknown>;
}

/**
 * Strip the prefix a spreadsheet FILE carries on a "future function".
 *
 * This is not cosmetic. Every function added after Excel 2007 is stored in
 * xlsx as `_xlfn.NAME` — AGGREGATE, IFS, TEXTJOIN, CONCAT, SWITCH, MAXIFS,
 * STDEV.P, NORM.S.INV, XLOOKUP and the rest — and worksheet-only ones as
 * `_xlfn._xlws.NAME` (FILTER, SORT). The prefix is how a 2007 reader knows it
 * cannot evaluate the formula; a reader that DOES implement the function is
 * expected to ignore it.
 *
 * Nothing in this codebase did. A workbook saved by Excel or WPS using any of
 * those functions imported with the prefix intact and every such cell resolved
 * to #NAME? — found by differential-testing against WPS, where
 * `_xlfn.AGGREGATE(9,0,A1:A5)` returned 15 and this engine returned #NAME?.
 *
 * Names TYPED by a user never carry the prefix, so stripping it here costs
 * nothing and is not a fallback that can mask a genuine unknown-name error:
 * `_xlfn.NOTAFUNCTION` still resolves to #NAME?, just without the prefix.
 */
/**
 * Strip the prefix a spreadsheet FILE carries on a LAMBDA or LET parameter.
 *
 * Excel stores `=LAMBDA(a,b,a*b)(3,4)` as
 * `_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a*_xlpm.b)(3,4)`: `_xlpm.` marks a
 * name as a parameter for a reader that predates LAMBDA ([MS-XLSX] 2.2.2,
 * future functions). A reader that implements LAMBDA reads the name without
 * it. Left on, the name carried a period, which no LAMBDA parameter may, so
 * every LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY a workbook
 * saved by Excel uses was #VALUE! (XLS-A084); LET tolerated the period and
 * so read correctly by accident. A name a user types never has the prefix.
 */
export function stripFormulaParameterPrefix(name: string): string {
  return /^_xlpm\./i.test(name) ? name.slice("_xlpm.".length) : name;
}

export function stripFutureFunctionPrefix(upperName: string): string {
  let name = upperName;
  if (name.startsWith("_XLFN.")) name = name.slice("_XLFN.".length);
  if (name.startsWith("_XLWS.")) name = name.slice("_XLWS.".length);
  // WPS marks ITS OWN extensions the same way, with `_wpsfn.`. Saving a workbook
  // through WPS rewrites ROUNDBANK, SUBSTITUTES, REGEXP, REPTARRAY, BOOKNAME and
  // SHEETSNAME with the prefix (measured; BAHTTEXT, NUMBERSTRING and DATESTRING
  // keep their bare names, and PERCENTOF gets `_xlfn.` instead), so without this
  // every one of those cells would import as #NAME?.
  if (name.startsWith("_WPSFN.")) name = name.slice("_WPSFN.".length);
  return name;
}

const STORED_FORMULA_PREFIX = /^(?:_xlfn\.|_xlws\.|_xlpm\.|_xleta\.)/i;

/**
 * A formula as Excel SHOWS it -- formula bar, Show Formulas, FORMULATEXT --
 * without the storage prefixes a file keeps on future functions (`_xlfn.`,
 * `_xlfn._xlws.`) and LAMBDA / LET parameters (`_xlpm.`): Excel shows
 * `=SORT(A2:A5)` for a cell stored `_xlfn._xlws.SORT(A2:A5)` (MS-XLSX 2.2.2;
 * X-169, S17_SimpleSpillOperatorFixture). Text in quotes, quoted sheet names
 * and bracketed structured-reference / workbook parts are left as written;
 * the stored formula itself is unchanged (the writer re-adds the prefixes).
 */
export function formulaWithoutStoragePrefixes(formula: string): string {
  if (!/_xl(?:fn|ws|pm|eta)\./i.test(formula)) return formula;
  let out = "";
  let depth = 0;
  for (let i = 0; i < formula.length; ) {
    const ch = formula[i];
    if (depth === 0 && (ch === '"' || ch === "'")) {
      let end = i + 1;
      while (end < formula.length) {
        if (formula[end] === ch) {
          if (formula[end + 1] === ch) { end += 2; continue; }
          break;
        }
        end += 1;
      }
      out += formula.slice(i, end + 1);
      i = end + 1;
      continue;
    }
    if (ch === "[") depth += 1;
    else if (ch === "]" && depth > 0) depth -= 1;
    if (depth === 0 && ch === "_" && !/[A-Za-z0-9_.\\]/.test(formula[i - 1] ?? "")) {
      let match = STORED_FORMULA_PREFIX.exec(formula.slice(i));
      if (match) {
        while (match) {
          i += match[0].length;
          match = STORED_FORMULA_PREFIX.exec(formula.slice(i));
        }
        continue;
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

/** Excel's numeric TEXT forms: optional sign, decimal, optional exponent. Used to
 *  REJECT JS-only syntaxes (hex 0x1F→31, 0b/0o, "Infinity", "NaN") that Excel
 *  treats as plain text, keeping every numeric-coercion site consistent. */
const EXCEL_NUMBER_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
export function isExcelNumericText(s: string): boolean {
  const t = s.trim();
  return EXCEL_NUMBER_RE.test(t) && Number.isFinite(Number(t));
}

export type FnArgs = Array<{
  scalar: CellValue;
  list: CellValue[];
  grid: CellValue[][];
  /** Full address dimensions when a sparse whole-row/whole-column reference
   *  materializes only the sheet's used rectangle. */
  referenceRows?: number;
  referenceColumns?: number;
  referenceCellCount?: number;
  /**
   * True when this argument came from a cell reference/range rather than a
   * direct literal. Excel counts logicals/text DIFFERENTLY by source: a direct
   * literal TRUE/numeric-text is coerced, but logicals/text inside a reference
   * are ignored. Aggregate functions use this via {@link flattenNumericArgs}.
   */
  fromRange?: boolean;
  /** The value is a nested expression's bare scalar result, carried in a 1×1 grid (B507). */
  scalarResult?: boolean;
  /**
   * True when the argument evaluated to an array VALUE (an array constant,
   * an array function's result, a name bound to one), not to a reference. A
   * 1×1 array keeps its array identity here although its grid looks like a
   * single cell: TYPE({1}) is 64 while TYPE(1) and TYPE(A1) are not.
   */
  arrayValue?: boolean;
  /**
   * True for an explicitly empty slot (`SUM(1,,3)`). It is NOT the same as a
   * blank cell: an optional argument left empty keeps its default, while a
   * reference to an empty cell is the value 0 / FALSE / "".
   */
  omitted?: boolean;
  /**
   * For a REFERENCE argument: read the rectangle that starts at the
   * reference's top-left cell and has the given size, on the reference's own
   * sheet (cells past the grid read as blanks). SUMIF / AVERAGEIF size their
   * sum_range / average_range this way — Microsoft's AVERAGEIF example reads
   * B1:B5 for average_range B1:B3 against range A1:A5 — so a shorter or 1×1
   * third argument still reaches the spatially matching cells.
   */
  resizeFrom?: (rows: number, cols: number) => CellValue[][];
  /** For a D-function's criteria reference: the computed criteria it holds
   *  (see {@link attachDatabaseCriteriaFormulas}). */
  criteriaFormulaAt?: DatabaseCriteriaFormula;
}>;

/** Excel criteria string -> predicate. Supports >, >=, <, <=, <>, = and
 *  bare values (numeric or text, with simple * / ? wildcards). */
/** Compile an Excel wildcard pattern (``*`` = any run, ``?`` = any single char,
 *  ``~`` escapes a literal ``*``/``?``/``~``) to a regex SOURCE (no anchors).
 *  Shared by makeCriteria and wildcardEqual so they stay consistent. */
export function excelWildcardToRegexSource(pat: string): string {
  let out = "";
  for (let i = 0; i < pat.length; i++) {
    const ch = pat[i];
    if (
      ch === "~" &&
      i + 1 < pat.length &&
      (pat[i + 1] === "?" || pat[i + 1] === "*" || pat[i + 1] === "~")
    ) {
      out += pat[i + 1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      i++;
    } else if (ch === "?") {
      out += ".";
    } else if (ch === "*") {
      out += ".*";
    } else {
      out += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
  }
  return out;
}

/**
 * The wall clock TODAY, NOW and a year-less DATEVALUE read. Unset, it is this
 * process's local time (the browser's). A calculation host that runs for
 * someone elsewhere (the server's calculation session, DESIGN-WAVE2 5.3.6)
 * sets the user's instant and UTC offset for the duration of its calculation,
 * so the server's own time zone never reaches a result.
 */
export interface CalculationClock {
  readonly epochMs: number;
  /** Minutes east of UTC at `epochMs` (`-new Date().getTimezoneOffset()`). */
  readonly utcOffsetMinutes: number;
}

/** Per-record evaluation of a computed criterion: the formula in criteria cell
 *  (critRow, critCol), re-read for the record `offset` rows below the first. */
export type DatabaseCriteriaFormula = (
  critRow: number,
  critCol: number,
  offset: number,
) => CellValue | undefined;

// ---------------------------------------------------------------------------
// Reference-returning functions (OFFSET / INDIRECT) — dynamic addressing
// ---------------------------------------------------------------------------
//
// Unlike every entry in FUNCTIONS above (which receives already-RESOLVED arg
// VALUES), OFFSET and INDIRECT need to compute a *reference* — an address
// rectangle — and only then read the cells behind it. OFFSET in particular
// needs the ADDRESS of its first argument (where it points), not its value, so
// it cannot be expressed over `FnArgs`. These functions therefore live in a
// SEPARATE registry that receives the un-resolved argument AST nodes plus the
// EvalContext, and is dispatched in the "call" case BEFORE the FnArgs path is
// built — so the existing FnArgs dispatch (and every function in it) is
// completely untouched. Both return an EngineValue (a scalar for a 1×1 result,
// an ArrayValue for a multi-cell block) so they reuse the existing array/spill
// machinery: a block spills, and SUM(OFFSET(...)) flows the whole grid.

/** A resolved reference rectangle: a 0-indexed top-left plus its dimensions. */
export interface RefBox {
  row: number;
  col: number;
  height: number; // ≥ 1
  width: number; // ≥ 1
  /**
   * The OTHER sheet the rectangle lies on (`Sheet2!A1:B3`); absent for the
   * formula's own sheet, including an own-sheet name written out
   * (`Sheet1!A1` on Sheet1). Every reader of a box goes through
   * {@link refBoxGrid} / {@link refBoxToValue}, which honor it.
   */
  sheet?: string;
}

// ---------------------------------------------------------------------------
// Reference shifting (fill-down / move) — relative refs move, absolute stay.
// ---------------------------------------------------------------------------

const SHIFT_REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

function shiftRef(ref: string, rowDelta: number, colDelta: number): string {
  const m = SHIFT_REF_RE.exec(ref);
  if (!m) return ref;
  const colAbs = m[1];
  const rowAbs = m[3];
  let col = lettersToColumn(m[2]);
  let row = Number(m[4]) - 1;
  if (!colAbs) col += colDelta;
  if (!rowAbs) row += rowDelta;
  if (col < 0 || col >= MAX_GRID_COLUMNS || row < 0 || row >= MAX_GRID_ROWS)
    return "#REF!";
  return `${colAbs}${columnToLetters(col)}${rowAbs}${row + 1}`;
}

/** The two written sides of a whole-column / whole-row token, `$` kept apart. */
const WHOLE_COLUMN_SIDES_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/;
const WHOLE_ROW_SIDES_RE = /^(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})$/;

/** `B:$D` -> { from: 1, fromLock: "", to: 3, toLock: "$" } (0-indexed lines). */
interface WholeAxisSides {
  axis: "column" | "row";
  fromLock: string;
  from: number;
  toLock: string;
  to: number;
}

function readWholeAxisSides(tok: Token): WholeAxisSides | null {
  if (tok.axis === "column") {
    const m = WHOLE_COLUMN_SIDES_RE.exec(tok.value);
    if (!m) return null;
    const from = lettersToColumn(m[2]);
    const to = lettersToColumn(m[4]);
    return from < 0 || to < 0
      ? null
      : { axis: "column", fromLock: m[1], from, toLock: m[3], to };
  }
  const m = WHOLE_ROW_SIDES_RE.exec(tok.value);
  if (!m) return null;
  const from = Number(m[2]) - 1;
  const to = Number(m[4]) - 1;
  return from < 0 || to < 0
    ? null
    : { axis: "row", fromLock: m[1], from, toLock: m[3], to };
}

function writeWholeAxisSides(sides: WholeAxisSides, from: number, to: number): string {
  const side = (line: number) =>
    sides.axis === "column" ? columnToLetters(line) : String(line + 1);
  return `${sides.fromLock}${side(from)}:${sides.toLock}${side(to)}`;
}

/**
 * Fill / copy for a whole column or row: `A:A` is `A1:A1048576`, so it moves
 * across columns (=SUM(A:A) filled right is =SUM(B:B)) and stays put down the
 * rows; `1:1` is the transpose. `$A:$A` is locked, and each side keeps its own
 * `$` like a mixed cell reference (`$A:C` filled right is `$A:D`). A side
 * pushed off the sheet is #REF!, as for a cell reference.
 */
function shiftWholeAxisRef(tok: Token, rowDelta: number, colDelta: number): string {
  const sides = readWholeAxisSides(tok);
  if (!sides) return tok.value;
  const delta = sides.axis === "column" ? colDelta : rowDelta;
  if (delta === 0 || (sides.fromLock && sides.toLock)) return tok.value;
  const limit = sides.axis === "column" ? MAX_GRID_COLUMNS : MAX_GRID_ROWS;
  const from = sides.fromLock ? sides.from : sides.from + delta;
  const to = sides.toLock ? sides.to : sides.to + delta;
  if (from < 0 || to < 0 || from >= limit || to >= limit) return ERRORS.REF.code;
  return writeWholeAxisSides(sides, from, to);
}

/**
 * Re-emit the `Sheet2!` / `'My Sheet'!` / `Sheet1:Sheet3!` prefix the tokenizer
 * split onto a sheet-qualified ref token's `sheet`/`endSheet` fields (empty for
 * an unprefixed token). The inverse of readSheetPrefix: a tokenize → transform →
 * `.join("")` round-trip MUST re-attach this, or the cross-sheet target is
 * silently dropped and the ref rebinds to the active sheet (wrong data).
 *
 * Quoting follows {@link sheetNameNeedsQuotes}, so `回款` is written bare, as
 * Excel writes it. A 3D span needing quotes on either end is quoted as ONE
 * pair, `'Q 1:Q 3'!` — the only quoted span form Excel reads (MS-XLSX 2.2.2
 * sheet-range); `'Q 1':'Q 3'!` is #NAME? to Excel and here (X-179).
 */
function sheetPrefixOf(tok: Token): string {
  if (tok.sheet === undefined) return "";
  if (tok.endSheet === undefined) return `${formatSheetNameForFormula(tok.sheet)}!`;
  const span = `${tok.sheet}:${tok.endSheet}`;
  return sheetNameNeedsQuotes(tok.sheet) || sheetNameNeedsQuotes(tok.endSheet)
    ? `'${span.replace(/'/g, "''")}'!`
    : `${span}!`;
}

/**
 * The spill-range operator a `ref` token carried (`A1#`), re-attached after its
 * rewritten address: a fill, move or insert moves the anchor and the reference
 * still names the whole spilled block (=SUM(A1#) filled down is =SUM(A2#)). A
 * rewrite that broke the address drops it — Excel writes =SUM(#REF!).
 */
function spillSuffixOf(tok: Token, address: string): string {
  return tok.spill === true && address !== ERRORS.REF.code ? "#" : "";
}

/**
 * Shift the relative cell references in a formula by (rowDelta, colDelta),
 * leaving absolute ($-prefixed) parts unchanged. Re-serializes from tokens, so
 * whitespace is normalized. Function names that match the A1 pattern (LOG10)
 * are left alone.
 */
/**
 * The link_location of a formula whose whole value is a HYPERLINK call —
 * `=HYPERLINK(A1,"Open")` answers `=A1` — or null for any other formula.
 *
 * HYPERLINK's VALUE is its display text (see FUNCTIONS.HYPERLINK); what makes
 * the cell clickable is its target, which the host evaluates from this
 * formula in the cell's own context, so a target read from another cell
 * follows that cell (Microsoft, HYPERLINK: clicking the cell opens
 * link_location). Only the top-level call counts, the way the engine reads
 * the cell's value; a HYPERLINK nested in other arithmetic is plain text.
 */
export function hyperlinkFormulaLocation(formula: string): string | null {
  const match = /^=\s*(?:_xlfn\.)?HYPERLINK\s*\(/i.exec(formula);
  if (!match) return null;
  let depth = 0;
  let argEnd = -1;
  let i = match[0].length;
  for (; i < formula.length; i += 1) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // A string literal or a quoted sheet name; a doubled quote escapes one.
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) {
            j += 2;
            continue;
          }
          break;
        }
        j += 1;
      }
      i = j;
      continue;
    }
    if (ch === "(" || ch === "{" || ch === "[") depth += 1;
    else if (ch === "]" || ch === "}") depth -= 1;
    else if (ch === ")") {
      if (depth === 0) break;
      depth -= 1;
    } else if (ch === "," && depth === 0 && argEnd < 0) argEnd = i;
  }
  // The call must close and be the whole formula.
  if (i >= formula.length || formula.slice(i + 1).trim() !== "") return null;
  const location = formula.slice(match[0].length, argEnd < 0 ? i : argEnd).trim();
  return location === "" ? null : `=${location}`;
}

/**
 * The cells and ranges a formula READS, without evaluating it.
 *
 * Exposed so a caller can walk the dependency graph iteratively — see
 * lib/spreadsheet/formula-dependencies, which uses it to resolve a deep chain
 * bottom-up instead of letting the recursive evaluator run out of stack. Only
 * same-sheet, non-spill references are reported: a cross-sheet or spill
 * reference is left to the evaluator, which already handles both.
 *
 * Returns `null` when the formula does not tokenize, so a caller can tell
 * "reads nothing" from "cannot be read".
 */
export function referencedRangesOf(
  formula: string,
): Array<{ r1: number; c1: number; r2: number; c2: number }> | null {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return null;
  const out: Array<{ r1: number; c1: number; r2: number; c2: number }> = [];
  const plain = (index: number) => {
    const tok = tokens[index];
    if (!tok || tok.type !== "ref") return null;
    // A "ref" immediately followed by "(" is a function name, not a reference.
    if (tokens[index + 1]?.type === "lparen") return null;
    // Another sheet's cell, or a spill anchor: the evaluator owns those.
    if (tok.sheet || tok.endSheet || tok.spill) return null;
    return parseCellRef(tok.value.replace(/\$/g, ""));
  };
  for (let i = 0; i < tokens.length; i += 1) {
    const from = plain(i);
    if (!from) continue;
    // A range is three tokens — ref, colon, ref — not one (see `tokenize`).
    const ranged = tokens[i + 1]?.type === "colon" ? plain(i + 2) : null;
    const to = ranged ?? from;
    out.push({
      r1: Math.min(from.row, to.row),
      c1: Math.min(from.col, to.col),
      r2: Math.max(from.row, to.row),
      c2: Math.max(from.col, to.col),
    });
    if (ranged) i += 2;
  }
  return out;
}

/**
 * Function names whose result can change without any referenced cell changing.
 * A formula calling one of these can never be carried across an edit: NOW/RAND
 * change on their own, OFFSET/INDIRECT read cells the tokens do not name,
 * CELL/INFO/SHEET(S) read workbook state, GETPIVOTDATA reads a pivot plan.
 */
const VOLATILE_OR_OPAQUE_FUNCTIONS = new Set([
  "NOW",
  "TODAY",
  "RAND",
  "RANDBETWEEN",
  "RANDARRAY",
  "OFFSET",
  "INDIRECT",
  "CELL",
  "INFO",
  "SHEET",
  "SHEETS",
  "GETPIVOTDATA",
  // Its result follows the host's picture fetch (loading, shown, refused).
  "IMAGE",
  // Their results follow what the host fetched (external-data).
  "WEBSERVICE",
  "IMPORTRANGE",
  // A What-If Data Table's result: its formula and value cells are named by
  // the table, not the text, and Excel recalculates every data table
  // whenever the worksheet recalculates.
  "TABLE",
]);

const XLSX_MAX_ROW = 1_048_575;
const XLSX_MAX_COL = 16_383;

/**
 * The COMPLETE set of same-sheet cells a formula reads — or null when that set
 * cannot be proven complete.
 *
 * This is the strict sibling of {@link referencedRangesOf}, and the difference
 * is the whole point: that helper SKIPS what it does not understand, which is
 * fine for warming a cache (a missed precedent just warms less) and fatal for
 * INVALIDATION (a missed dependency is a stale value on screen). Here, anything
 * that could read a cell the tokens do not name returns null — a cross-sheet or
 * 3D or spill reference, a structured table reference, a defined name, a
 * volatile function, a dynamic range endpoint (`A1:INDEX(...)`). Callers treat
 * null as "this formula is dirty on every edit", which is always sound.
 *
 * Whole-row/column references (`A:A`, `2:7`) are returned as their full
 * rectangles rather than rejected — `SUM(A:A)` is too common to give up on,
 * and a containment test against a large rectangle is exact.
 */
/**
 * The cells SUMIF / AVERAGEIF read through their third argument, which is
 * sized to the criteria range from its top-left cell. `lparen` is the call's
 * opening parenthesis. Undefined when nothing beyond the written refs is read
 * (no third argument, or one that is not a reference); null when the size or
 * the anchor cannot be pinned from the tokens.
 */
function resizedSumRangeOf(
  tokens: readonly Token[],
  lparen: number,
  plainBox: (index: number) => { row: number; col: number } | null,
): { r1: number; c1: number; r2: number; c2: number } | undefined | null {
  const args: Array<[number, number]> = [];
  let depth = 0;
  let start = lparen + 1;
  let end = -1;
  for (let j = lparen + 1; j < tokens.length; j += 1) {
    const type = tokens[j].type;
    if (type === "lparen" || type === "lbrace") depth += 1;
    else if ((type === "rparen" || type === "rbrace") && depth > 0) depth -= 1;
    else if (type === "rparen") {
      args.push([start, j]);
      end = j;
      break;
    } else if (type === "comma" && depth === 0) {
      args.push([start, j]);
      start = j + 1;
    }
  }
  if (end < 0) return null;
  if (args.length < 3) return undefined;
  const box = ([from, to]: [number, number]) => {
    if (to - from === 1) {
      const cell = plainBox(from);
      return cell ? { r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col } : null;
    }
    if (to - from === 3 && tokens[from + 1]?.type === "colon") {
      const a = plainBox(from);
      const b = plainBox(from + 2);
      if (!a || !b) return null;
      return {
        r1: Math.min(a.row, b.row),
        c1: Math.min(a.col, b.col),
        r2: Math.max(a.row, b.row),
        c2: Math.max(a.col, b.col),
      };
    }
    return null;
  };
  const sum = box(args[2]);
  if (!sum) {
    const [from, to] = args[2];
    const mayBeReference = tokens
      .slice(from, to)
      .some((token) => token.type === "ref" || token.type === "lparen" || token.type === "wholeref");
    return mayBeReference ? null : undefined;
  }
  const range = box(args[0]);
  if (!range) return null;
  return {
    r1: sum.r1,
    c1: sum.c1,
    r2: Math.min(XLSX_MAX_ROW, sum.r1 + (range.r2 - range.r1)),
    c2: Math.min(XLSX_MAX_COL, sum.c1 + (range.c2 - range.c1)),
  };
}

export function dependencyRangesOf(
  formula: string,
): Array<{ r1: number; c1: number; r2: number; c2: number }> | null {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return null;
  const out: Array<{ r1: number; c1: number; r2: number; c2: number }> = [];
  const plainBox = (index: number) => {
    const tok = tokens[index];
    if (!tok || tok.type !== "ref") return null;
    if (tok.sheet || tok.endSheet || tok.spill) return null;
    return parseCellRef(tok.value.replace(/\$/g, ""));
  };
  for (let i = 0; i < tokens.length; i += 1) {
    const tok = tokens[i];
    const isCall =
      (tok.type === "ref" || tok.type === "ident") &&
      tokens[i + 1]?.type === "lparen";
    if (isCall) {
      const name = tok.value.toUpperCase();
      if (VOLATILE_OR_OPAQUE_FUNCTIONS.has(name)) {
        return null;
      }
      // SUMIF / AVERAGEIF read their sum range at the criteria range's size
      // from its top-left cell (SUMIF(F55:F59,"A",G55) reads G55:G59), so the
      // cells past the written sum range are read too (live audit S-101: a
      // change in G59 left the total stale).
      if (name === "SUMIF" || name === "AVERAGEIF") {
        const resized = resizedSumRangeOf(tokens, i + 1, plainBox);
        if (resized === null) return null;
        if (resized) out.push(resized);
      }
      continue;
    }
    switch (tok.type) {
      case "tableref":
        // Structured references resolve through a registry this walk cannot
        // see move.
        return null;
      case "wholeref": {
        // `Sheet2!A:A` reads another sheet's cells — opaque here, like any
        // sheet-qualified ref.
        if (tok.sheet !== undefined) return null;
        const cols = /^\$?([A-Za-z]{1,3})\$?:\$?([A-Za-z]{1,3})$/.exec(
          tok.value,
        );
        if (cols) {
          const a = parseCellRef(`${cols[1].toUpperCase()}1`);
          const b = parseCellRef(`${cols[2].toUpperCase()}1`);
          if (!a || !b) return null;
          out.push({
            r1: 0,
            r2: XLSX_MAX_ROW,
            c1: Math.min(a.col, b.col),
            c2: Math.max(a.col, b.col),
          });
          break;
        }
        const rows = /^\$?([0-9]{1,7})\$?:\$?([0-9]{1,7})$/.exec(tok.value);
        if (rows) {
          const a = Number(rows[1]) - 1;
          const b = Number(rows[2]) - 1;
          if (a < 0 || b < 0) return null;
          out.push({
            r1: Math.min(a, b),
            r2: Math.max(a, b),
            c1: 0,
            c2: XLSX_MAX_COL,
          });
          break;
        }
        return null;
      }
      case "ref": {
        const from = plainBox(i);
        // Not a call (handled above), not a plain A1 box: a defined name, a
        // sheet-qualified ref, a spill anchor — all opaque here.
        if (!from) return null;
        // `INDEX(...):A9`: the range's other endpoint is computed.
        if (tokens[i - 1]?.type === "colon" && tokens[i - 2]?.type === "rparen") return null;
        if (tokens[i + 1]?.type === "colon") {
          const to = plainBox(i + 2);
          // `A1:INDEX(...)` and friends: a range endpoint this walk cannot pin.
          if (!to || tokens[i + 3]?.type === "lparen") return null;
          out.push({
            r1: Math.min(from.row, to.row),
            c1: Math.min(from.col, to.col),
            r2: Math.max(from.row, to.row),
            c2: Math.max(from.col, to.col),
          });
          i += 2;
          break;
        }
        out.push({ r1: from.row, c1: from.col, r2: from.row, c2: from.col });
        break;
      }
      case "ident": {
        const upper = tok.value.toUpperCase();
        // A bare identifier that is not TRUE/FALSE is a defined name.
        if (upper !== "TRUE" && upper !== "FALSE") return null;
        break;
      }
      default:
        break;
    }
  }
  return out;
}

/** One cell whose raw content differs from the snapshot the memo came from. */
export interface WarmStartChangedCell {
  sheet: string;
  row: number;
  col: number;
}

/** One formula currently in the workbook (the caller enumerates; the engine
 *  reads through callbacks and cannot). */
export interface WarmStartFormula {
  sheet: string;
  row: number;
  col: number;
  formula: string;
}

export interface WorkbookWarmStart {
  /** The previous workbook's memo, exactly as memoSnapshot() returned it. */
  snapshot: ReadonlyMap<string, EngineValue>;
  changed: readonly WarmStartChangedCell[];
  formulas: readonly WarmStartFormula[];
  /** The memo keys not to carry, when the caller found them already (a
   * dependency index it keeps across edits, `WarmDependencyIndex`): then
   * `formulas` need only list the formulas among them. */
  dirty?: ReadonlySet<string>;
}

/**
 * Formula-result authority for a selective workbook calculation.
 *
 * A frozen sheet may still expose CURRENT literal cells, but every formula on
 * that sheet must resolve from this exact memo key instead of executing. That
 * is the engine boundary SpreadsheetWindow needs for Excel's Shift+F9: formulas
 * on the active sheet run, formulas on every other dirty sheet keep their last
 * calculated result. The seed is deliberately distinct from WorkbookWarmStart:
 * warm-start is a performance hint that may safely fall back to computation;
 * this contract is semantic authority and a missing entry must fail closed.
 */
export interface WorkbookFormulaResultSeed {
  snapshot: ReadonlyMap<string, EngineValue>;
  /** Sheet names whose formula cells must never execute in this generation. */
  frozenSheets: readonly string[];
  /** Rendering may surface #N/A; publication paths must request "throw". */
  onMissing: "error" | "throw";
  /**
   * Formula cells on frozen sheets that evaluate now regardless of the
   * snapshot: the entries made since the last calculation. Excel calculates
   * the formula you enter even in Manual mode (its dependents wait for F9).
   */
  evaluate?: ReadonlySet<string>;
}






export class FormulaResultUnavailableError extends Error {
  readonly sheetName: string;
  readonly row: number;
  readonly col: number;

  constructor(sheetName: string, row: number, col: number, message = `No calculated result for ${sheetName}!${row}:${col}.`) {
    // A control-flow signal, thrown once for every formula a save finds no
    // result for: no stack is captured (a deep render stack is most of what
    // one costs; the cell it names is the useful part).
    const limit = Error.stackTraceLimit;
    Error.stackTraceLimit = 0;
    try {
      super(message);
    } finally {
      Error.stackTraceLimit = limit;
    }
    this.name = "FormulaResultUnavailableError";
    this.sheetName = sheetName;
    this.row = row;
    this.col = col;
  }
}

export function shiftFormula(
  formula: string,
  rowDelta: number,
  colDelta: number,
): string {
  const moved = shiftedFormulaPieces(formula, rowDelta, colDelta);
  if (!moved) return formula;
  if (!/\s/.test(formula)) return `${formula.startsWith("=") ? "=" : ""}${moved.join("")}`;
  return piecesInTheFormulasSpacing(formula, moved, shiftedFormulaPieces(formula, 0, 0));
}

function shiftedFormulaPieces(formula: string, rowDelta: number, colDelta: number): string[] | null {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return null;
  const out = tokens
    .map((tok, i): string => {
      if (tok.type === "ref" && tokens[i + 1]?.type !== "lparen") {
        // A sheet-qualified ref keeps its target sheet on a fill/copy: only the
        // relative cell coordinates move (=Sheet2!A1 filled down -> =Sheet2!A2),
        // exactly as in Excel. sheetPrefixOf is "" for a same-sheet ref, so this
        // stays byte-identical for unprefixed formulas.
        const shifted = shiftRef(tok.value, rowDelta, colDelta);
        return `${sheetPrefixOf(tok)}${shifted}${spillSuffixOf(tok, shifted)}`;
      }
      if (tok.type === "string") {
        return `"${tok.value.replace(/"/g, '""')}"`;
      }
      if (tok.type === "wholeref") {
        // A whole column pushed off the grid keeps its sheet, as a cell does:
        // Windows Excel writes =INDEX('Data Sheet'!#REF!,...) for a copy of
        // =INDEX('Data Sheet'!A:A,...) one column left (office-truth copy
        
        const shifted = shiftWholeAxisRef(tok, rowDelta, colDelta);
        return `${sheetPrefixOf(tok)}${shifted}`;
      }
      if (tok.type === "tableref") return tok.raw ?? tok.value;
      return tok.value;
    });
  // A range one of whose corners left the grid is #REF! whole, as Excel
  
  // three rows is =SUM(#REF!), not =SUM(#REF!:B2)).
  for (let i = 0; i + 2 < tokens.length; i += 1) {
    if (
      tokens[i].type === "ref" &&
      tokens[i + 1].type === "colon" &&
      tokens[i + 2].type === "ref" &&
      tokens[i + 3]?.type !== "lparen" &&
      (out[i].endsWith(ERRORS.REF.code) || out[i + 2].endsWith(ERRORS.REF.code))
    ) {
      out[i] = `${sheetPrefixOf(tokens[i])}${ERRORS.REF.code}`;
      out[i + 1] = "";
      out[i + 2] = "";
      i += 2;
    }
  }
  return out;
}

/**
 * A formula pasted with Paste Special ▸ Transpose. Excel reflects every cell
 * reference across the copied block's diagonal: the address at (row, column)
 * offset from the copied block's top-left lands at (column, row) offset from
 * the paste's top-left, and its `$` locks swap axes with it. That holds for
 * references outside the block and for absolute ones too (Excel 16, copying
 * A1:C2 to G4 transposed: `=C5+A1+B1` -> `=K6+G4+G5`, `$A$1` -> `$G$4`,
 * `A$1` -> `$G4`, `SUM(A1:B1)` -> `SUM(G4:G5)`). A plain per-cell shift kept
 * the references pointing the copied way, so the results came out wrong.
 * Whole rows/columns and table references keep the per-cell shift.
 */
export function transposeFormula(
  formula: string,
  sourceRow: number,
  sourceCol: number,
  sourceOriginRow: number,
  sourceOriginCol: number,
  destinationOriginRow: number,
  destinationOriginCol: number,
): string {
  const hasEq = formula.startsWith("=");
  const src = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return formula;
  const destinationRow = destinationOriginRow + (sourceCol - sourceOriginCol);
  const destinationCol = destinationOriginCol + (sourceRow - sourceOriginRow);
  const rowDelta = destinationRow - sourceRow;
  const colDelta = destinationCol - sourceCol;
  const reflect = (ref: string): string => {
    const m = SHIFT_REF_RE.exec(ref);
    if (!m) return ref;
    const col = lettersToColumn(m[2]);
    const row = Number(m[4]) - 1;
    const nextRow = destinationOriginRow + (col - sourceOriginCol);
    const nextCol = destinationOriginCol + (row - sourceOriginRow);
    if (
      nextCol < 0 ||
      nextCol >= MAX_GRID_COLUMNS ||
      nextRow < 0 ||
      nextRow >= MAX_GRID_ROWS
    )
      return "#REF!";
    // The row's lock now guards the column and the column's the row.
    return `${m[3]}${columnToLetters(nextCol)}${m[1]}${nextRow + 1}`;
  };
  const out = tokens
    .map((tok, i) => {
      if (tok.type === "ref" && tokens[i + 1]?.type !== "lparen") {
        const moved = reflect(tok.value);
        return `${sheetPrefixOf(tok)}${moved}${spillSuffixOf(tok, moved)}`;
      }
      if (tok.type === "string") {
        return `"${tok.value.replace(/"/g, '""')}"`;
      }
      if (tok.type === "wholeref") {
        const shifted = shiftWholeAxisRef(tok, rowDelta, colDelta);
        return shifted === ERRORS.REF.code ? shifted : `${sheetPrefixOf(tok)}${shifted}`;
      }
      if (tok.type === "tableref") return tok.raw ?? tok.value;
      return tok.value;
    })
    .join("");
  return `${hasEq ? "=" : ""}${out}`;
}

// ---------------------------------------------------------------------------
// Structural reference adjustment (insert/delete rows & columns)
// ---------------------------------------------------------------------------

export type StructuralAxis = "row" | "col";

/**
 * Adjust one reference for a structural change. Unlike fill-down shifting,
 * inserting/deleting a row or column shifts BOTH relative and absolute refs
 * (Excel moves $A$5 to $A$6 when a row is inserted). A reference to the deleted
 * line becomes #REF!.
 *
 * @param at    0-indexed row/column where the insert/delete happens.
 * @param delta +1 to insert a line at `at`, -1 to delete the line at `at`.
 */
function adjustStructuralRef(
  ref: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  endpoint?: "start" | "end",
): string {
  const m = SHIFT_REF_RE.exec(ref);
  if (!m) return ref;
  const colAbs = m[1];
  const rowAbs = m[3];
  let col = lettersToColumn(m[2]);
  let row = Number(m[4]) - 1;
  // When the deleted line falls exactly on a RANGE endpoint, Excel SHRINKS the
  // range (A1:A10 delete row 10 -> A1:A9; delete row 1 -> A1:A9) instead of
  // breaking the whole formula. A standalone ref (no endpoint role) on the
  // deleted line still becomes #REF!. The "start" endpoint stays at `at` (the
  // cell now occupying it); the "end" endpoint clamps to `at-1` (the last
  // surviving row/col). A fully-collapsed range falls through the bounds check
  // below to #REF!. A range's end on the grid's last line stays there through an
  // insert and a delete before it (Excel 16 Mac, xlsx-rounds evidence/
  // xprobe-gridend: $B$10:$B$1048576 reads so after either, A1048570:A1048576
  // reads A1048571:A1048576 and A1048569:A1048576); a start or a single cell
  // pushed past the grid is #REF!.
  if (axis === "row") {
    if (delta > 0) {
      if (row >= at) row += delta;
      if (endpoint === "end") row = Math.min(row, MAX_GRID_ROWS - 1);
    } else if (row === at) {
      if (endpoint === "start") row = at;
      else if (endpoint === "end") row = at - 1;
      else return ERRORS.REF.code;
    } else if (row > at && !(endpoint === "end" && row === MAX_GRID_ROWS - 1)) {
      row += delta;
    }
  } else if (delta > 0) {
    if (col >= at) col += delta;
    if (endpoint === "end") col = Math.min(col, MAX_GRID_COLUMNS - 1);
  } else if (col === at) {
    if (endpoint === "start") col = at;
    else if (endpoint === "end") col = at - 1;
    else return ERRORS.REF.code;
  } else if (col > at && !(endpoint === "end" && col === MAX_GRID_COLUMNS - 1)) {
    col += delta;
  }
  if (col < 0 || col >= MAX_GRID_COLUMNS || row < 0 || row >= MAX_GRID_ROWS)
    return ERRORS.REF.code;
  return `${colAbs}${columnToLetters(col)}${rowAbs}${row + 1}`;
}

/**
 * The whole-column / whole-row counterpart of {@link adjustStructuralRef}. Only
 * an edit on the reference's own axis moves it (a row insert never touches
 * `A:A`), and `$` does not pin it — a structural edit moves every reference.
 * Inserting at or before a side pushes it along, so =SUM(A:A) reads =SUM(B:B)
 * after a column is inserted before A and A:C grows to A:D when one is
 * inserted inside it; an end pushed past the sheet edge stays on the last line.
 * Deleting lines shrinks the span from whichever end they cut, and deleting
 * every line it names is #REF!.
 */
function adjustStructuralWholeRef(
  tok: Token,
  axis: StructuralAxis,
  at: number,
  delta: number,
): string {
  const sides = readWholeAxisSides(tok);
  if (!sides || (sides.axis === "column") !== (axis === "col")) return tok.value;
  const limit = sides.axis === "column" ? MAX_GRID_COLUMNS : MAX_GRID_ROWS;
  const low = Math.min(sides.from, sides.to);
  const high = Math.max(sides.from, sides.to);
  let nextLow: number;
  let nextHigh: number;
  if (delta > 0) {
    nextLow = low >= at ? low + delta : low;
    nextHigh = high >= at ? Math.min(high + delta, limit - 1) : high;
    if (nextLow >= limit) return ERRORS.REF.code;
  } else {
    const last = at - delta - 1; // the deleted lines are at..last
    nextLow = low < at ? low : low > last ? low + delta : at;
    // A side on the grid's last line stays there through a delete before it,
    // as a cell range's end does (Windows Excel 16, xlsx-rounds
    // night/n58/x41/probe: $5:$1048576 and ROWS($3:$1048576) kept after row 11
    // was deleted; R468-X41 E6, a name over Overall!$1:$1048576).
    nextHigh = high < at ? high : high === limit - 1 && high > last ? high : high > last ? high + delta : at - 1;
    if (nextHigh < nextLow) return ERRORS.REF.code;
  }
  if (nextLow === low && nextHigh === high) return tok.value;
  // Keep the written order (`C:A` stays high-first) along with each side's `$`.
  return sides.from <= sides.to
    ? writeWholeAxisSides(sides, nextLow, nextHigh)
    : writeWholeAxisSides(sides, nextHigh, nextLow);
}

/**
 * Which references an insert/delete of rows or columns on one worksheet moves.
 * Excel rewrites every reference whose TARGET is the edited sheet, however it
 * is written: `A1` in a formula on that sheet, `First!A1` on that sheet or on
 * any other, and nothing that points at a different sheet.
 */
export interface StructuralChangeScope {
  /** The edited worksheet's name (compared case-insensitively, as Excel does). */
  readonly sheetName: string;
  /** The formula lives on ANOTHER sheet: its unqualified references are that
   * sheet's own and stay; only references qualified with `sheetName` move. */
  readonly qualifiedOnly?: boolean;
}

/** Whether a structural edit in `scope` moves the reference `tok`. Without a
 * scope (legacy callers) only unqualified references move. A 3-D reference
 * (`First:Third!A1`) never moves here; see structuralChangeMovesThreeDReference. */
function structuralScopeMoves(tok: Token, scope: StructuralChangeScope | undefined): boolean {
  if (tok.endSheet !== undefined) return false;
  if (tok.sheet === undefined) return !scope?.qualifiedOnly;
  return scope !== undefined && tok.sheet.toUpperCase() === scope.sheetName.toUpperCase();
}

/**
 * Rewrite the references in a formula for an inserted/deleted row or column,
 * so dependent formulas keep pointing at the right cells (or surface #REF!).
 * `scope` names the edited sheet; without it only unqualified references move.
 */
export function adjustFormulaForStructuralChange(
  formula: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string {
  const moved = structuralFormulaPieces(formula, axis, at, delta, scope);
  if (!moved) return formula;
  if (!/\s/.test(formula)) return `${formula.startsWith("=") ? "=" : ""}${moved.join("")}`;
  // (at -1, nothing on the edited line: a zero-line edit at a real line
  // would read as a delete of it)
  const spelled = structuralFormulaPieces(formula, axis, -1, 0, scope);
  // Nothing in it moved: the formula is what it was, whatever the tokenizer
  // made of its text. Rebuilt, text a file stores as a formula that does not
  // parse as one lost its spaces on every row insert (xlsx-rounds R381-X11:
  // <f>1 if the school is a small-scale rural school ...</f>).
  if (spelled && spelled.length === moved.length && spelled.every((piece, index) => piece === moved[index])) return formula;
  return piecesInTheFormulasSpacing(formula, moved, spelled);
}

/**
 * The author's spacing and line breaks (Alt+Enter in the formula bar) stay
 * where they were, as Excel keeps them when references move: rebuilt from
 * tokens, a long laid-out formula came back as one line (xlsx-rounds R13-X13,
 * a LaTeX dictionary template). `spelled` is the same rewrite with nothing
 * moving, each token as the tokenizer reads it; where every one of those is
 * the original's own text, the whitespace between them goes back around
 * `moved`.
 */
function piecesInTheFormulasSpacing(formula: string, moved: readonly string[], spelled: readonly string[] | null): string {
  const hasEq = formula.startsWith("=");
  const rewritten = `${hasEq ? "=" : ""}${moved.join("")}`;
  if (!spelled || spelled.length !== moved.length) return rewritten;
  const src = hasEq ? formula.slice(1) : formula;
  let position = 0;
  const kept: string[] = [];
  let broken = false;
  for (let index = 0; index < spelled.length; index += 1) {
    const piece = spelled[index];
    if (piece === "" || piece.trim() === "") return rewritten;
    let gap = position;
    while (gap < src.length && /\s/.test(src[gap])) gap += 1;
    if (!src.startsWith(piece, gap)) return rewritten;
    // A reference the edit broke takes the space before it away, and an area
    // closed to one #REF! the spaces inside it: Excel 16 Mac writes
    // `IF(x, $G$2:$G$9)` with column G deleted as `IF(x,#REF!)`, its line
    
    if (moved[index].endsWith(ERRORS.REF.code) && !piece.endsWith(ERRORS.REF.code)) broken = true;
    else if (moved[index] !== "") broken = false;
    kept.push(broken ? "" : src.slice(position, gap), moved[index]);
    position = gap + piece.length;
  }
  if (src.slice(position).trim() !== "") return rewritten;
  kept.push(src.slice(position));
  return `${hasEq ? "=" : ""}${kept.join("")}`;
}

/** The structural rewrite of a formula's text as one string per token (a
 *  range's endpoints and colon apart; a range closed by a delete as one
 *  `#REF!`), or null when the text does not tokenize. */
function structuralFormulaPieces(
  formula: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string[] | null {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return null;
  const qualified = (tok: Token, address: string) =>
    `${sheetPrefixOf(tok)}${address}${spillSuffixOf(tok, address)}`;
  const out: string[] = [];
  for (let i = 0; i < tokens.length; i += 1) {
    const tok = tokens[i];
    if (tok.type === "ref" && tokens[i + 1]?.type !== "lparen") {
      const moves = structuralScopeMoves(tok, scope);
      const end = tokens[i + 2];
      if (tokens[i + 1]?.type === "colon" && end?.type === "ref" && tokens[i + 3]?.type !== "lparen") {
        // A range: the start's qualifier governs its end (`Second!A1:A10` —
        // the `A10` token carries no sheet). A reference into another sheet
        // is untouched: inserting a row here never moves that sheet's cells.
        if (!moves) {
          out.push(qualified(tok, tok.value), ":", qualified(end, end.value));
          i += 2;
          continue;
        }
        // Endpoints shrink (not #REF!) when their line is deleted; a range
        // whose every line is deleted is #REF!, never a silent inverted range
        // over the neighbouring cells. `A12:A10` names the lines of `A10:A12`:
        // which endpoint is the range's first line is its coordinate on the
        // edited axis, not the order it was written in, and the written order
        // is kept.
        const before = [parseCellRef(tok.value), parseCellRef(end.value)];
        const writtenReversed =
          before[0] && before[1]
            ? axis === "row"
              ? before[0].row > before[1].row
              : before[0].col > before[1].col
            : false;
        const start = adjustStructuralRef(
          tok.value, axis, at, delta, writtenReversed ? "end" : "start");
        const finish = adjustStructuralRef(
          end.value, axis, at, delta, writtenReversed ? "start" : "end");
        const after = [parseCellRef(start), parseCellRef(finish)];
        const inverted = (pair: Array<CellRef | null>) =>
          pair[0] && pair[1] ? (axis === "row" ? pair[0].row > pair[1].row : pair[0].col > pair[1].col) : false;
        // The grid's last line itself deleted: a new line comes in at the end, and a
        // range reaching it from before still reaches it (xlsx-rounds evidence/
        // xprobe-gridend2: B10:B1048576 reads the same, A1048576:B1048576 is #REF!).
        const lastLine = axis === "row" ? MAX_GRID_ROWS - 1 : MAX_GRID_COLUMNS - 1;
        const line = (ref: CellRef | null) => (ref ? (axis === "row" ? ref.row : ref.col) : -1);
        if (
          delta < 0 && at === lastLine && after[0] && after[1] && !inverted(before) && !inverted(after) &&
          line(before[writtenReversed ? 0 : 1]) === lastLine
        ) {
          out.push(qualified(tok, tok.value), ":", qualified(end, end.value));
          i += 2;
          continue;
        }
        if (!after[0] || !after[1] || (!inverted(before) && inverted(after))) {
          // A deleted area keeps the sheet it named, as Excel writes it
          
          // three pieces stay three (the two after it empty), so the
          // author's spacing still lines up (piecesInTheFormulasSpacing).
          out.push(`${sheetPrefixOf(tok)}${ERRORS.REF.code}`, "", "");
        } else {
          out.push(`${sheetPrefixOf(tok)}${start}`, ":", qualified(end, finish));
        }
        i += 2;
        continue;
      }
      if (!moves) {
        out.push(qualified(tok, tok.value));
        continue;
      }
      // A ref after a colon that did not start with a cell (`INDEX(...):B5`)
      // is still a range end, which shrinks rather than breaking.
      const adjusted = adjustStructuralRef(tok.value, axis, at, delta,
        tokens[i - 1]?.type === "colon" ? "end" : undefined);
      out.push(adjusted === ERRORS.REF.code ? `${sheetPrefixOf(tok)}${adjusted}` : qualified(tok, adjusted));
      continue;
    }
    if (tok.type === "string") {
      out.push(`"${tok.value.replace(/"/g, '""')}"`);
    } else if (tok.type === "wholeref") {
      // Another sheet's column is untouched, exactly like a qualified cell.
      if (!structuralScopeMoves(tok, scope)) {
        out.push(`${sheetPrefixOf(tok)}${tok.value}`);
      } else {
        // A broken whole column keeps its sheet too (`Second!#REF!`, which the
        // tokenizer reads as the error it is).
        const adjusted = adjustStructuralWholeRef(tok, axis, at, delta);
        out.push(`${sheetPrefixOf(tok)}${adjusted}`);
      }
    } else if (tok.type === "tableref") {
      out.push(tok.raw ?? tok.value);
    } else {
      out.push(tok.value);
    }
  }
  return out;
}

/**
 * Whether a structural edit would have to move a 3-D reference
 * (`First:Third!A5`) whose span includes the edited sheet. Only one sheet of
 * the span moves, so no single rewrite of the reference is right for every
 * sheet it names; the caller refuses the edit instead of leaving the span
 * silently pointing at shifted cells. `spansEditedSheet` receives the span's
 * two sheet names and decides by workbook order.
 */
export function structuralChangeMovesThreeDReference(
  formula: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  spansEditedSheet: (from: string, to: string) => boolean,
): boolean {
  const tokens = tokenize(formula.startsWith("=") ? formula.slice(1) : formula);
  if (!tokens) return false;
  for (let i = 0; i < tokens.length; i += 1) {
    const tok = tokens[i];
    if (tok.sheet === undefined || tok.endSheet === undefined || !spansEditedSheet(tok.sheet, tok.endSheet)) continue;
    if (tok.type === "wholeref" && adjustStructuralWholeRef(tok, axis, at, delta) !== tok.value) return true;
    if (tok.type !== "ref") continue;
    const end = tokens[i + 1]?.type === "colon" && tokens[i + 2]?.type === "ref" ? tokens[i + 2] : undefined;
    if (adjustStructuralRef(tok.value, axis, at, delta, end ? "start" : undefined) !== tok.value) return true;
    if (end && adjustStructuralRef(end.value, axis, at, delta, "end") !== end.value) return true;
  }
  return false;
}

export interface FormulaCellShiftOperation {
  readonly range: { r1: number; c1: number; r2: number; c2: number };
  readonly direction: "down" | "right" | "up" | "left";
  /** Current sheet name; lets explicit `Sheet1!A1` refs move with this sheet. */
  readonly sheetName?: string;
  /** Peer worksheets: rewrite only references qualified with sheetName. */
  readonly qualifiedOnly?: boolean;
}

export interface FormulaCellShiftResult {
  readonly formula: string;
  readonly supported: boolean;
}

function normalizeCellShiftRange(
  range: FormulaCellShiftOperation["range"],
): FormulaCellShiftOperation["range"] {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}

function mapCellShiftRef(
  ref: string,
  operation: FormulaCellShiftOperation,
  endpoint?: "start" | "end",
): string {
  const match = SHIFT_REF_RE.exec(ref);
  if (!match) return ref;
  const range = normalizeCellShiftRange(operation.range);
  const height = range.r2 - range.r1 + 1;
  const width = range.c2 - range.c1 + 1;
  let column = lettersToColumn(match[2]);
  let row = Number(match[4]) - 1;
  const vertical = operation.direction === "down" || operation.direction === "up";
  const inCorridor = vertical
    ? column >= range.c1 && column <= range.c2
    : row >= range.r1 && row <= range.r2;
  if (!inCorridor) return ref;

  if (operation.direction === "down" && row >= range.r1) {
    row += height;
  } else if (operation.direction === "right" && column >= range.c1) {
    column += width;
  } else if (operation.direction === "up") {
    if (row >= range.r1 && row <= range.r2) {
      if (endpoint === "start") row = range.r1;
      else if (endpoint === "end") row = range.r1 - 1;
      else return ERRORS.REF.code;
    } else if (row > range.r2) {
      row -= height;
    }
  } else if (operation.direction === "left") {
    if (column >= range.c1 && column <= range.c2) {
      if (endpoint === "start") column = range.c1;
      else if (endpoint === "end") column = range.c1 - 1;
      else return ERRORS.REF.code;
    } else if (column > range.c2) {
      column -= width;
    }
  }
  if (
    column < 0 ||
    column >= MAX_GRID_COLUMNS ||
    row < 0 ||
    row >= MAX_GRID_ROWS
  ) {
    return ERRORS.REF.code;
  }
  return `${match[1]}${columnToLetters(column)}${match[3]}${row + 1}`;
}

/**
 * Whether an Insert/Delete Cells shift moves, stretches or shrinks the range
 * `start:end` as a whole. It does when the range lies clear of the moved band,
 * or wholly inside the band's rows (a horizontal shift) or columns (a vertical
 * one). A range that the band crosses only in part is not the shift's to
 * rewrite: Excel leaves such a reference as written (`=SUM(C4:C10)` stays put
 * when Insert Cells shifts row 7 right from B7), and so does MRAI.
 */
function cellShiftMovesWholeRange(
  start: CellRef,
  end: CellRef,
  operation: FormulaCellShiftOperation,
): boolean {
  const range = normalizeCellShiftRange(operation.range);
  const ref = {
    r1: Math.min(start.row, end.row),
    r2: Math.max(start.row, end.row),
    c1: Math.min(start.col, end.col),
    c2: Math.max(start.col, end.col),
  };
  const vertical = operation.direction === "down" || operation.direction === "up";
  if (vertical) {
    const intersectsCorridor = ref.c2 >= range.c1 && ref.c1 <= range.c2;
    if (!intersectsCorridor || ref.r2 < range.r1) return true;
    return ref.c1 >= range.c1 && ref.c2 <= range.c2;
  }
  const intersectsCorridor = ref.r2 >= range.r1 && ref.r1 <= range.r2;
  if (!intersectsCorridor || ref.c2 < range.c1) return true;
  return ref.r1 >= range.r1 && ref.r2 <= range.r2;
}

/**
 * Adjust same-sheet A1 references after an insert/delete CELLS operation.
 * A range the moved band crosses only in part keeps its text, as in Excel
 * (see cellShiftMovesWholeRange). `supported=false` is left for a formula that
 * cannot be tokenized or a range endpoint that cannot be read, and for a 3-D
 * span in the workbook-scoped (`qualifiedOnly`) pass.
 */
// Read-only tokenization shared by capability checks and the actual rewrite.
// Bound both text and token residency; never cache result values or coordinates.
const cellShiftTokenMemo = new Map<string, Token[] | null>();
let cellShiftMemoCharacters = 0;
let cellShiftMemoTokens = 0;
function cellShiftTokens(source: string): Token[] | null {
  if (cellShiftTokenMemo.has(source)) return cellShiftTokenMemo.get(source)!;
  const tokens = tokenize(source);
  if (source.length <= 65_536 && (tokens?.length ?? 0) <= 10_000) {
    while (cellShiftTokenMemo.size >= 20_000 || cellShiftMemoCharacters + source.length > 2_000_000 || cellShiftMemoTokens + (tokens?.length ?? 0) > 250_000) {
      const oldest = cellShiftTokenMemo.keys().next().value;
      if (oldest === undefined) break;
      cellShiftMemoCharacters -= oldest.length;
      cellShiftMemoTokens -= cellShiftTokenMemo.get(oldest)?.length ?? 0;
      cellShiftTokenMemo.delete(oldest);
    }
    cellShiftTokenMemo.set(source, tokens);
    cellShiftMemoCharacters += source.length;
    cellShiftMemoTokens += tokens?.length ?? 0;
  }
  return tokens;
}

const SCALAR_AGGREGATE_ROOTS = new Set(["SUM", "COUNT", "COUNTA", "AVERAGE", "MIN", "MAX", "SUBTOTAL"]);

/** Syntactic proof for spill discovery, NOT a cached result or evaluator.
 * Call only with evaluator.getFormula(row,col), never cell.formula metadata.
 * The caller must prove this is a standalone, unseeded evaluator: manual seeds
 * or result overrides may retain an array even when the current raw is scalar.
 * Local lambdas can shadow FUNCTIONS dispatch. Missing context is conservative.
 * Only complete root calls with simple local arguments are proven; operators,
 * strings, names, qualified names, nested calls and dynamic forms stay unknown.
 */
export function isDefinitelyScalarAggregateFormula(
  formula: string | null,
  context?: { standaloneUnseeded?: boolean; letBindings?: ReadonlyMap<string, unknown> },
): boolean {
  if (context?.standaloneUnseeded !== true || context.letBindings?.size || typeof formula !== "string" || formula.length > 8192) return false;
  const normalized = formula.trimStart();
  const source = normalized.startsWith("=") ? normalized.slice(1) : normalized;
  const tokens = cellShiftTokens(source);
  if (!tokens || tokens.length > 2048 || tokens.some(token => token.type === "ident" && token.value.includes("."))) return false;
  try {
    const root = new Parser(tokens).parse(); // parse() rejects trailing operators/tokens.
    if (root.kind !== "call" || !SCALAR_AGGREGATE_ROOTS.has(root.name)) return false;
    return root.args.every(arg => {
      switch (arg.kind) {
        case "num": case "bool": case "empty": case "error": case "wholeRange": return true;
        case "ref": return arg.sheet === undefined;
        case "range": return arg.sheet === undefined && arg.endSheet === undefined;
        default: return false;
      }
    });
  } catch { return false; }
}

/** Check the writer's refusals (an unreadable formula or range endpoint)
 * without constructing rewritten formulas. Ribbon availability needs only this predicate, not 4 full rewrites
 * of every formula whenever the selection moves. */
/** The parts of a formula a cell-shift capability check reads: its range
 * references (sheet qualifier upper-cased, endpoints parsed once) and whether
 * any reference spans sheets. Tokenizing happens once per formula text; the
 * check itself then only compares rectangles. */
export interface PreparedCellShiftFormula {
  readonly ranges: ReadonlyArray<{
    readonly sheet: string | undefined;
    readonly start: CellRef | null;
    readonly end: CellRef | null;
  }>;
  readonly threeD: boolean;
}

/** null when the formula cannot be tokenized; such a formula never adjusts. */
export function prepareFormulaForCellShift(formula: string): PreparedCellShiftFormula | null {
  const tokens = cellShiftTokens(formula.startsWith("=") ? formula.slice(1) : formula);
  if (!tokens) return null;
  const ranges: Array<PreparedCellShiftFormula["ranges"][number]> = [];
  let threeD = false;
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.type !== "ref" || tokens[index + 1]?.type === "lparen" || tokens[index - 1]?.type === "colon") continue;
    if (token.endSheet !== undefined) threeD = true;
    const end = tokens[index + 2];
    if (tokens[index + 1]?.type !== "colon" || end?.type !== "ref") continue;
    ranges.push({ sheet: token.sheet?.toUpperCase(), start: parseCellRef(token.value), end: parseCellRef(end.value) });
    index += 2;
  }
  return { ranges, threeD };
}

export function canAdjustPreparedFormulaForCellShift(prepared: PreparedCellShiftFormula | null, operation: FormulaCellShiftOperation): boolean {
  if (!prepared) return false;
  if (operation.qualifiedOnly !== undefined && prepared.threeD) return false;
  const sheetName = operation.sheetName?.toUpperCase();
  for (const range of prepared.ranges) {
    if (operation.qualifiedOnly && range.sheet === undefined) continue;
    if (range.sheet !== undefined && (!sheetName || range.sheet !== sheetName)) continue;
    // A range the band crosses only in part stays as written, so only an
    // unreadable endpoint refuses (cellShiftMovesWholeRange).
    if (!range.start || !range.end) return false;
  }
  return true;
}

export function canAdjustFormulaForCellShift(formula: string, operation: FormulaCellShiftOperation): boolean {
  return canAdjustPreparedFormulaForCellShift(prepareFormulaForCellShift(formula), operation);
}

export function adjustFormulaForCellShift(
  formula: string,
  operation: FormulaCellShiftOperation,
): FormulaCellShiftResult {
  const hasEq = formula.startsWith("=");
  const source = hasEq ? formula.slice(1) : formula;
  const tokens = cellShiftTokens(source);
  if (!tokens) return { formula, supported: false };
  const replacements = new Map<number, string>();
  let changed = false;

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.type !== "ref" || tokens[index + 1]?.type === "lparen") continue;
    if (tokens[index - 1]?.type === "colon") continue;
    // A 3D span may include the edited sheet between its endpoints. The
    // workbook-scoped caller must not silently shift the whole span or leave
    // just one member stale. Legacy partial-cell callers retain their policy.
    if (operation.qualifiedOnly !== undefined && token.endSheet !== undefined)
      return { formula, supported: false };
    if (operation.qualifiedOnly && token.sheet === undefined) continue;

    const colon = tokens[index + 1];
    const rangeEnd = tokens[index + 2];
    if (colon?.type === "colon" && rangeEnd?.type === "ref") {
      // A qualified range belongs to another sheet unless its qualifier names
      // the sheet being edited. Older callers omit sheetName and retain the
      // conservative cross-sheet behavior.
      if (
        token.sheet !== undefined &&
        (!operation.sheetName ||
          token.sheet.toUpperCase() !== operation.sheetName.toUpperCase())
      ) {
        continue;
      }
      const start = parseCellRef(token.value);
      const end = parseCellRef(rangeEnd.value);
      if (!start || !end) return { formula, supported: false };
      if (!cellShiftMovesWholeRange(start, end, operation)) {
        // Excel leaves a range the shift crosses only in part as written.
        index += 2;
        continue;
      }
      const nextStart = mapCellShiftRef(token.value, operation, "start");
      const nextEnd = mapCellShiftRef(rangeEnd.value, operation, "end");
      const parsedStart = parseCellRef(nextStart);
      const parsedEnd = parseCellRef(nextEnd);
      if (
        nextStart === ERRORS.REF.code ||
        nextEnd === ERRORS.REF.code ||
        !parsedStart ||
        !parsedEnd ||
        parsedStart.row > parsedEnd.row ||
        parsedStart.col > parsedEnd.col
      ) {
        replacements.set(index, ERRORS.REF.code);
        replacements.set(index + 1, "");
        replacements.set(index + 2, "");
        changed = true;
      } else {
        replacements.set(
          index,
          `${token.sheet !== undefined ? sheetPrefixOf(token) : ""}${nextStart}`,
        );
        replacements.set(index + 2, nextEnd);
        changed ||= nextStart !== token.value || nextEnd !== rangeEnd.value;
      }
      index += 2;
      continue;
    }

    if (
      token.sheet !== undefined &&
      (!operation.sheetName ||
        token.sheet.toUpperCase() !== operation.sheetName.toUpperCase())
    ) {
      continue;
    }
    const next = mapCellShiftRef(token.value, operation);
    replacements.set(
      index,
      next === ERRORS.REF.code
        ? next
        : `${token.sheet !== undefined ? sheetPrefixOf(token) : ""}${next}${spillSuffixOf(token, next)}`,
    );
    changed ||= next !== token.value;
  }

  if (!changed) return { formula, supported: true };
  const output = tokens
    .map((token, index) => {
      const replacement = replacements.get(index);
      if (replacement !== undefined) return replacement;
      if (token.type === "ref") {
        return `${sheetPrefixOf(token)}${token.value}${spillSuffixOf(token, token.value)}`;
      }
      if (token.type === "string") return `"${token.value.replace(/"/g, '""')}"`;
      if (token.type === "wholeref") return `${sheetPrefixOf(token)}${token.value}`;
      if (token.type === "tableref") return token.raw ?? token.value;
      return token.value;
    })
    .join("");
  return { formula: `${hasEq ? "=" : ""}${output}`, supported: true };
}

export interface FormulaCellMoveOperation {
  /** The cut source rectangle on the sheet named `sheetName`. */
  readonly from: { r1: number; c1: number; r2: number; c2: number };
  readonly rowDelta: number;
  readonly colDelta: number;
  /** The sheet the move happened on; qualified refs naming it move too. */
  readonly sheetName?: string;
  /** Formulas on OTHER sheets: only refs qualified with sheetName apply. */
  readonly qualifiedOnly?: boolean;
}

/**
 * Cut and Paste MOVES cells, and Excel keeps every formula pointing at them:
 * `=C4*2` reads `=G4*2` once C4 has been cut to G4, on this sheet and (through
 * a qualified ref) on every other. A reference follows when the whole of what
 * it names moved — a single cell inside the source, or a range whose two
 * endpoints are both inside it. What the moved block lands ON is destroyed, so
 * a reference to a cell it overwrote becomes #REF!, as in Excel. A range that
 * only partly overlaps the source or the destination is left as written.
 * Absolute markers are kept: a move is not a fill, `$C$4` follows as `$G$4`.
 * Returns the ORIGINAL string when nothing applies.
 */
export function adjustFormulaForCellMove(
  formula: string,
  operation: FormulaCellMoveOperation,
): string {
  if (!formula || (operation.rowDelta === 0 && operation.colDelta === 0)) {
    return formula;
  }
  const hasEq = formula.startsWith("=");
  const source = hasEq ? formula.slice(1) : formula;
  const tokens = cellShiftTokens(source);
  if (!tokens) return formula;
  const from = normalizeCellShiftRange(operation.from);
  const to = {
    r1: from.r1 + operation.rowDelta,
    r2: from.r2 + operation.rowDelta,
    c1: from.c1 + operation.colDelta,
    c2: from.c2 + operation.colDelta,
  };
  const inside = (
    rect: typeof from,
    cell: { row: number; col: number },
  ): boolean =>
    cell.row >= rect.r1 &&
    cell.row <= rect.r2 &&
    cell.col >= rect.c1 &&
    cell.col <= rect.c2;
  const moved = (ref: string): string => {
    const match = SHIFT_REF_RE.exec(ref);
    if (!match) return ref;
    const column = lettersToColumn(match[2]) + operation.colDelta;
    const row = Number(match[4]) - 1 + operation.rowDelta;
    if (
      column < 0 ||
      column >= MAX_GRID_COLUMNS ||
      row < 0 ||
      row >= MAX_GRID_ROWS
    ) {
      return ERRORS.REF.code;
    }
    return `${match[1]}${columnToLetters(column)}${match[3]}${row + 1}`;
  };
  const replacements = new Map<number, string>();

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.type !== "ref" || tokens[index + 1]?.type === "lparen") continue;
    if (tokens[index - 1]?.type === "colon") continue;
    // A 3D span names several sheets at once; leave it as written.
    if (token.endSheet !== undefined) continue;
    if (token.sheet === undefined ? operation.qualifiedOnly : (
      !operation.sheetName ||
      token.sheet.toUpperCase() !== operation.sheetName.toUpperCase()
    )) {
      continue;
    }
    const prefix = token.sheet !== undefined ? sheetPrefixOf(token) : "";
    const start = parseCellRef(token.value);
    if (!start) continue;
    const colon = tokens[index + 1];
    const rangeEnd = tokens[index + 2];
    if (colon?.type === "colon" && rangeEnd?.type === "ref") {
      const end = parseCellRef(rangeEnd.value);
      index += 2;
      if (!end) continue;
      if (inside(from, start) && inside(from, end)) {
        const nextStart = moved(token.value);
        const nextEnd = moved(rangeEnd.value);
        if (nextStart === ERRORS.REF.code || nextEnd === ERRORS.REF.code) {
          replacements.set(index - 2, ERRORS.REF.code);
          replacements.set(index - 1, "");
          replacements.set(index, "");
        } else {
          replacements.set(index - 2, `${prefix}${nextStart}`);
          replacements.set(index, nextEnd);
        }
      } else if (
        inside(to, start) &&
        inside(to, end) &&
        !(inside(from, start) || inside(from, end))
      ) {
        replacements.set(index - 2, ERRORS.REF.code);
        replacements.set(index - 1, "");
        replacements.set(index, "");
      }
      continue;
    }
    if (inside(from, start)) {
      const next = moved(token.value);
      replacements.set(
        index,
        next === ERRORS.REF.code
          ? next
          : `${prefix}${next}${spillSuffixOf(token, next)}`,
      );
    } else if (inside(to, start)) {
      replacements.set(index, ERRORS.REF.code);
    }
  }

  if (replacements.size === 0) return formula;
  const output = tokens
    .map((token, index) => {
      const replacement = replacements.get(index);
      if (replacement !== undefined) return replacement;
      if (token.type === "ref") {
        return `${sheetPrefixOf(token)}${token.value}${spillSuffixOf(token, token.value)}`;
      }
      if (token.type === "string") return `"${token.value.replace(/"/g, '""')}"`;
      if (token.type === "wholeref") return `${sheetPrefixOf(token)}${token.value}`;
      if (token.type === "tableref") return token.raw ?? token.value;
      return token.value;
    })
    .join("");
  return `${hasEq ? "=" : ""}${output}`;
}

export interface FormulaCrossSheetMoveOperation {
  /** The cut source rectangle on `fromSheet`. */
  readonly from: { r1: number; c1: number; r2: number; c2: number };
  readonly rowDelta: number;
  readonly colDelta: number;
  /** The sheet the cells were cut from. */
  readonly fromSheet: string;
  /** The sheet they were pasted on. */
  readonly toSheet: string;
  /** The sheet the formula being rewritten lives on, after the move. */
  readonly hostSheet: string;
  /**
   * True for a formula that travelled with the cells: it lives on `toSheet`
   * now, but its unqualified references were written on `fromSheet` and
   * still name that sheet's cells.
   */
  readonly movedFormula?: boolean;
}

/**
 * Cut on one sheet, Paste on another MOVES the cells across the workbook,
 * and Excel keeps every formula pointing at them: `=Sheet2!A2*2` on Sheet3
 * reads `=C3*2` once Sheet2!A2 has been cut to Sheet3!C3, and `=A2*2` left
 * behind on Sheet2 reads `=Sheet3!C3*2`. A reference follows when the whole
 * of what it names moved (a cell inside the source, or a range with both
 * endpoints inside it); what the block lands on is destroyed, so a reference
 * to a destination cell becomes #REF!; a partial overlap is left as written.
 * A moved formula's own untouched references keep naming the source sheet,
 * spelled out now that the formula no longer lives there (`=B2` becomes
 * `=Sheet2!B2`). Absolute markers are kept: a move is not a fill. Returns the
 * ORIGINAL string when nothing applies. With the two sheet names equal this
 * is the same-sheet rule of {@link adjustFormulaForCellMove}.
 */
export function adjustFormulaForCrossSheetMove(
  formula: string,
  operation: FormulaCrossSheetMoveOperation,
): string {
  if (!formula) return formula;
  const same = (a: string | undefined, b: string): boolean =>
    a !== undefined && a.trim().toUpperCase() === b.trim().toUpperCase();
  if (same(operation.fromSheet, operation.toSheet)) {
    return adjustFormulaForCellMove(formula, {
      from: operation.from,
      rowDelta: operation.rowDelta,
      colDelta: operation.colDelta,
      sheetName: operation.fromSheet,
      qualifiedOnly: !same(operation.hostSheet, operation.fromSheet),
    });
  }
  const hasEq = formula.startsWith("=");
  const source = hasEq ? formula.slice(1) : formula;
  const tokens = cellShiftTokens(source);
  if (!tokens) return formula;
  const from = normalizeCellShiftRange(operation.from);
  const to = {
    r1: from.r1 + operation.rowDelta,
    r2: from.r2 + operation.rowDelta,
    c1: from.c1 + operation.colDelta,
    c2: from.c2 + operation.colDelta,
  };
  const inside = (
    rect: typeof from,
    cell: { row: number; col: number },
  ): boolean =>
    cell.row >= rect.r1 &&
    cell.row <= rect.r2 &&
    cell.col >= rect.c1 &&
    cell.col <= rect.c2;
  const moved = (ref: string): string => {
    const match = SHIFT_REF_RE.exec(ref);
    if (!match) return ref;
    const column = lettersToColumn(match[2]) + operation.colDelta;
    const row = Number(match[4]) - 1 + operation.rowDelta;
    if (
      column < 0 ||
      column >= MAX_GRID_COLUMNS ||
      row < 0 ||
      row >= MAX_GRID_ROWS
    ) {
      return ERRORS.REF.code;
    }
    return `${match[1]}${columnToLetters(column)}${match[3]}${row + 1}`;
  };
  const toPrefix = same(operation.hostSheet, operation.toSheet)
    ? ""
    : `${formatSheetNameForFormula(operation.toSheet)}!`;
  const fromPrefix = `${formatSheetNameForFormula(operation.fromSheet)}!`;
  // The sheet an unqualified reference names.
  const unqualifiedSheet = operation.movedFormula
    ? operation.fromSheet
    : operation.hostSheet;
  // An untouched unqualified reference of a moved formula, spelled with its
  // source sheet; null when the reference is written as it stands.
  const kept = (token: Token): string | null =>
    token.sheet === undefined && operation.movedFormula === true
      ? `${fromPrefix}${token.value}`
      : null;
  const replacements = new Map<number, string>();
  let changed = false;

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.type === "wholeref") {
      // A whole row or column never moves with a cell cut; a moved formula's
      // own unqualified one still names the source sheet.
      const keep = kept(token);
      if (keep !== null) {
        replacements.set(index, keep);
        changed = true;
      }
      continue;
    }
    if (token.type !== "ref" || tokens[index + 1]?.type === "lparen") continue;
    if (tokens[index - 1]?.type === "colon") continue;
    // A 3D span names several sheets at once; leave it as written.
    if (token.endSheet !== undefined) continue;
    const named = token.sheet ?? unqualifiedSheet;
    const namesFrom = same(named, operation.fromSheet);
    const namesTo = same(named, operation.toSheet);
    const start = parseCellRef(token.value);
    if (!start) continue;
    const colon = tokens[index + 1];
    const rangeEnd = tokens[index + 2];
    if (colon?.type === "colon" && rangeEnd?.type === "ref") {
      const end = parseCellRef(rangeEnd.value);
      index += 2;
      if (!end) continue;
      if (namesFrom && inside(from, start) && inside(from, end)) {
        const nextStart = moved(token.value);
        const nextEnd = moved(rangeEnd.value);
        if (nextStart === ERRORS.REF.code || nextEnd === ERRORS.REF.code) {
          replacements.set(index - 2, ERRORS.REF.code);
          replacements.set(index - 1, "");
          replacements.set(index, "");
        } else {
          replacements.set(index - 2, `${toPrefix}${nextStart}`);
          replacements.set(index, nextEnd);
        }
        changed = true;
      } else if (namesTo && inside(to, start) && inside(to, end)) {
        replacements.set(index - 2, ERRORS.REF.code);
        replacements.set(index - 1, "");
        replacements.set(index, "");
        changed = true;
      } else {
        const keep = kept(token);
        if (keep !== null) {
          replacements.set(index - 2, keep);
          changed = true;
        }
      }
      continue;
    }
    if (namesFrom && inside(from, start)) {
      const next = moved(token.value);
      replacements.set(
        index,
        next === ERRORS.REF.code
          ? next
          : `${toPrefix}${next}${spillSuffixOf(token, next)}`,
      );
      changed = true;
    } else if (namesTo && inside(to, start)) {
      replacements.set(index, ERRORS.REF.code);
      changed = true;
    } else {
      const keep = kept(token);
      if (keep !== null) {
        replacements.set(index, `${keep}${spillSuffixOf(token, token.value)}`);
        changed = true;
      }
    }
  }

  if (!changed) return formula;
  const output = tokens
    .map((token, index) => {
      const replacement = replacements.get(index);
      if (replacement !== undefined) return replacement;
      if (token.type === "ref") {
        return `${sheetPrefixOf(token)}${token.value}${spillSuffixOf(token, token.value)}`;
      }
      if (token.type === "string") return `"${token.value.replace(/"/g, '""')}"`;
      if (token.type === "wholeref") return `${sheetPrefixOf(token)}${token.value}`;
      if (token.type === "tableref") return token.raw ?? token.value;
      return token.value;
    })
    .join("");
  return `${hasEq ? "=" : ""}${output}`;
}

/**
 * Rewrite every sheet-qualified reference naming `oldName` to `newName`, so
 * cross-sheet formulas keep resolving after a sheet is renamed (renaming Sheet2
 * to Data turns `=Sheet2!A1` into `=Data!A1`). Sheet names match
 * case-insensitively (Excel parity) on both the start and end sheet of a 3D
 * span; `newName` is quoted iff {@link sheetNameNeedsQuotes} says Excel would
 * quote it, so renaming 回款 to "回款 2026" quotes the prefix and renaming it
 * back writes it bare again.
 *
 * References to OTHER sheets and same-sheet refs are preserved (the former keep
 * their own prefix; the latter are untouched). When NOTHING in the formula names
 * `oldName`, the ORIGINAL string is returned byte-for-byte (no re-serialization,
 * so a formula that does not touch the renamed sheet can never be perturbed).
 * Accepts a formula with or without a leading "=", and equally a bare ref string
 * such as a named-range target or a validation source (`'Lists'!$A$1:$A$10`).
 */
/** A sheet name as references compare it: case-insensitive, and with the
 *  outer spaces a name may carry ("Quest. ", xlsx-rounds R16-X02) compared
 *  the same way on both sides -- the name a caller passes was trimmed while
 *  the reference kept its space, so a rename or delete missed it. */
function sheetKey(name: string): string {
  return name.trim().toUpperCase();
}

export function renameSheetInFormula(
  formula: string,
  oldName: string,
  newName: string,
): string {
  if (!formula) return formula;
  const oldKey = oldName.trim().toUpperCase();
  if (oldKey === "" || oldKey === newName.trim().toUpperCase()) return formula;
  const hasEq = formula.startsWith("=");
  const src = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return formula;
  let changed = false;
  const out = tokens
    .map((tok) => {
      if ((tok.type === "ref" || tok.type === "wholeref") && tok.sheet !== undefined) {
        const sheet = sheetKey(tok.sheet) === oldKey ? newName : tok.sheet;
        const endSheet =
          tok.endSheet !== undefined && sheetKey(tok.endSheet) === oldKey
            ? newName
            : tok.endSheet;
        if (sheet !== tok.sheet || endSheet !== tok.endSheet) changed = true;
        return `${sheetPrefixOf({ ...tok, sheet, endSheet })}${tok.value}${spillSuffixOf(tok, tok.value)}`;
      }
      if (tok.type === "string") {
        return `"${tok.value.replace(/"/g, '""')}"`;
      }
      if (tok.type === "ident") {
        const qualified = splitQualifiedName(tok.value);
        if (qualified && sheetKey(qualified.sheet) === oldKey) {
          changed = true;
          return `${formatSheetNameForFormula(newName)}!${qualified.name}`;
        }
      }
      // `Outputs!#REF!`: the cells are gone, the sheet it names is not, and a
      // rename follows it (xlsx-rounds R132-X02, R132-X04).
      if (tok.type === "error" && tok.value.toUpperCase().endsWith("!#REF!")) {
        const prefix = tok.value.slice(0, -"!#REF!".length);
        const named = prefix.startsWith("'") && prefix.endsWith("'") ? prefix.slice(1, -1).replace(/''/g, "'") : prefix;
        if (sheetKey(named) === oldKey) {
          changed = true;
          return `${formatSheetNameForFormula(newName)}!#REF!`;
        }
      }
      if (tok.type === "ref") return `${tok.value}${spillSuffixOf(tok, tok.value)}`;
      if (tok.type === "wholeref") return `${sheetPrefixOf(tok)}${tok.value}`;
      if (tok.type === "tableref") return tok.raw ?? tok.value;
      return tok.value;
    });
  // Untouched formulas stay byte-identical; only a real rename re-serializes,
  // and that keeps the author's spacing and line breaks (Alt+Enter in the
  // formula bar) around the renamed references, as Excel does (xlsx-rounds
  // R101-X04: a SUMIFS laid out over three lines came back on one).
  if (!changed) return formula;
  return piecesInTheFormulasSpacing(formula, out, shiftedFormulaPieces(formula, 0, 0));
}

/**
 * Break every reference to a sheet that is being deleted, the way Excel does:
 * `=Sheet2!A1` becomes `=#REF!A1` and `=SUM(Sheet2!A1:B3)` becomes
 * `=SUM(#REF!A1:B3)`, so a sheet created later under the same name can never
 * silently take the reference over. A 3-D span that ENDS on the deleted sheet
 * shrinks to the sheets still between its ends (Microsoft Support, "Create a
 * reference to the same cell range on multiple worksheets": deleting Sheet2 or
 * Sheet6 of `Sheet2:Sheet6` adjusts the calculation to the range between
 * them); a span of that sheet alone breaks. A sheet deleted from the middle of
 * a span needs no text change.
 *
 * `sheetOrder` is the workbook's sheet names in tab order BEFORE the delete.
 * String literals are untouched, and a formula naming no deleted sheet comes
 * back byte-for-byte. Accepts a formula with or without "=", or a bare
 * reference such as a named-range target.
 */
/**
 * Adjust the 3-D spans of a formula for a sheet moved to another tab position
 * (Microsoft Support, "Create a reference to the same cell range on multiple
 * worksheets"): a sheet moved between the endpoints joins the span and one
 * moved out of them leaves it, which the span's endpoints already express;
 * an ENDPOINT moved past the other endpoint leaves the span, which then runs
 * over the sheets that remain between them (live audit B112: moving 数据 of
 * 数据:透视 to the end gives 表格:透视). A span left with no sheet breaks.
 * `orderBefore` / `orderAfter` are the tab orders around the move.
 */
export function adjustFormulaForSheetMove(
  formula: string,
  orderBefore: readonly string[],
  orderAfter: readonly string[],
  movedSheet: string,
): string {
  if (!formula || !formula.includes(":")) return formula;
  const upperBefore = orderBefore.map((name) => name.toUpperCase());
  const upperAfter = orderAfter.map((name) => name.toUpperCase());
  const movedKey = movedSheet.toUpperCase();
  const hasEq = formula.startsWith("=");
  const src = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return formula;
  let changed = false;
  const out = tokens
    .map((tok) => {
      if ((tok.type === "ref" || tok.type === "wholeref") && tok.sheet !== undefined && tok.endSheet !== undefined) {
        const startKey = tok.sheet.toUpperCase();
        const endKey = tok.endSheet.toUpperCase();
        if (startKey === movedKey || endKey === movedKey) {
          const b1 = upperBefore.indexOf(startKey);
          const b2 = upperBefore.indexOf(endKey);
          const a1 = upperAfter.indexOf(startKey);
          const a2 = upperAfter.indexOf(endKey);
          if (b1 >= 0 && b2 >= 0 && a1 >= 0 && a2 >= 0 && b1 !== b2) {
            const flipped = (b1 <= b2) !== (a1 <= a2);
            if (flipped) {
              // The moved endpoint left the span: the members that remain
              // (the old members without it) bound the new span.
              const lo = Math.min(b1, b2);
              const hi = Math.max(b1, b2);
              const remaining = upperBefore
                .slice(lo, hi + 1)
                .filter((name) => name !== movedKey)
                .map((name) => upperAfter.indexOf(name))
                .filter((index) => index >= 0)
                .sort((x, y) => x - y);
              changed = true;
              if (remaining.length === 0) {
                return `${ERRORS.REF.code}${tok.value}${spillSuffixOf(tok, ERRORS.REF.code)}`;
              }
              const sheet = orderAfter[remaining[0]];
              const endSheet = orderAfter[remaining[remaining.length - 1]];
              if (sheet === endSheet) {
                return `${sheetPrefixOf({ ...tok, sheet, endSheet: undefined })}${tok.value}${spillSuffixOf(tok, tok.value)}`;
              }
              return `${sheetPrefixOf({ ...tok, sheet, endSheet })}${tok.value}${spillSuffixOf(tok, tok.value)}`;
            }
          }
        }
        return `${sheetPrefixOf(tok)}${tok.value}${spillSuffixOf(tok, tok.value)}`;
      }
      if (tok.type === "string") return `"${tok.value.replace(/"/g, '""')}"`;
      if (tok.type === "ref") return `${sheetPrefixOf(tok)}${tok.value}${spillSuffixOf(tok, tok.value)}`;
      if (tok.type === "wholeref") return `${sheetPrefixOf(tok)}${tok.value}`;
      if (tok.type === "tableref") return tok.raw ?? tok.value;
      return tok.value;
    });
  if (!changed) return formula;
  return piecesInTheFormulasSpacing(formula, out, shiftedFormulaPieces(formula, 0, 0));
}

export function invalidateSheetInFormula(
  formula: string,
  deletedName: string,
  sheetOrder: readonly string[],
): string {
  if (!formula) return formula;
  const deletedKey = deletedName.trim().toUpperCase();
  if (deletedKey === "") return formula;
  const hasEq = formula.startsWith("=");
  const src = hasEq ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  if (!tokens) return formula;
  const order = sheetOrder.map((name) => name.trim().toUpperCase());
  const deletedIndex = order.indexOf(deletedKey);
  let changed = false;
  const out = tokens
    .map((tok) => {
      if ((tok.type === "ref" || tok.type === "wholeref") && tok.sheet !== undefined) {
        const startHit = sheetKey(tok.sheet) === deletedKey;
        const endHit =
          tok.endSheet !== undefined && sheetKey(tok.endSheet) === deletedKey;
        if (startHit || endHit) {
          changed = true;
          const start = order.indexOf(sheetKey(tok.sheet));
          const end =
            tok.endSheet === undefined ? start : order.indexOf(sheetKey(tok.endSheet));
          if (tok.endSheet === undefined || start < 0 || end < 0 || start === end || deletedIndex < 0) {
            return `${ERRORS.REF.code}${tok.value}`;
          }
          // Step the deleted end one sheet toward the other end.
          const step = start < end ? 1 : -1;
          const sheet = startHit ? sheetOrder[start + step] : tok.sheet;
          const endSheet = endHit ? sheetOrder[end - step] : tok.endSheet;
          const collapsed = sheet.toUpperCase() === endSheet.toUpperCase();
          const next = { ...tok, sheet, endSheet: collapsed ? undefined : endSheet };
          return `${sheetPrefixOf(next)}${tok.value}${spillSuffixOf(next, tok.value)}`;
        }
        return `${sheetPrefixOf(tok)}${tok.value}${spillSuffixOf(tok, tok.value)}`;
      }
      // `'X'!#REF!`, cells of X already gone, X now deleted: Excel shows
      // `#REF!#REF!` and stores the bare `#REF!` (xlsx-rounds
      // evidence/xprobe-refprefix, Excel 16 for Mac; R159-X01 E18).
      if (tok.type === "error" && tok.value.toUpperCase().endsWith("!#REF!")) {
        const prefix = tok.value.slice(0, -"!#REF!".length);
        const named = prefix.startsWith("'") && prefix.endsWith("'") ? prefix.slice(1, -1).replace(/''/g, "'") : prefix;
        if (sheetKey(named) === deletedKey) {
          changed = true;
          return `${ERRORS.REF.code}${ERRORS.REF.code}`;
        }
      }
      if (tok.type === "string") return `"${tok.value.replace(/"/g, '""')}"`;
      if (tok.type === "ref") return `${tok.value}${spillSuffixOf(tok, tok.value)}`;
      if (tok.type === "wholeref") return `${sheetPrefixOf(tok)}${tok.value}`;
      if (tok.type === "tableref") return tok.raw ?? tok.value;
      return tok.value;
    });
  if (!changed) return formula;
  return piecesInTheFormulasSpacing(formula, out, shiftedFormulaPieces(formula, 0, 0));
}

// ---------------------------------------------------------------------------
// Sheet evaluator (memoized, cycle-detecting)
// ---------------------------------------------------------------------------

/**
 * The error values a cell can HOLD, keyed by the code Excel stores for them.
 * A cell whose stored contents are one of these is an error value, not the
 * text of one: Excel's own NA page says "You can also type the value #N/A
 * directly into a cell", and its IS-functions page evaluates a cell holding
 * #N/A with ISNA → TRUE (so ISERROR → TRUE). #GETTING_DATA is deliberately
 * absent: it is a transient display state, not a value ISERROR reports on.
 */
const STORED_ERROR_LITERALS: Readonly<Record<string, FormulaError>> = {
  "#NULL!": ERRORS.NULL,
  "#DIV/0!": ERRORS.DIV0,
  "#VALUE!": ERRORS.VALUE,
  "#REF!": ERRORS.REF,
  "#NAME?": ERRORS.NAME,
  "#NUM!": ERRORS.NUM,
  "#N/A": ERRORS.NA,
  "#SPILL!": ERRORS.SPILL,
  "#CALC!": ERRORS.CALC,
};

/**
 * True when a cell's stored characters would not read back as that text: a
 * logical (TRUE/FALSE), a stored error (#N/A) or an Excel numeric literal.
 * An importer marks such a cell when the file typed it as a string, so the
 * cell keeps its string type (live audit B312).
 */
export function literalReadsAsNonText(raw: string): boolean {
  const value = coerceLiteral(raw);
  return value !== raw && value !== null;
}

/** Coerce a raw cell string (literal) to a typed value. */
export function coerceLiteral(raw: string): CellValue {
  if (raw === "") return null;
  const t = raw.trim();
  if (t === "") return raw;
  const upper = t.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  // An entry the cell stores as an error IS that error to every reader of the
  // cell, exactly as a formula's own #DIV/0! is. A cell whose type is text
  // (a quote prefix, or a Text-formatted cell) never reaches here.
  const storedError = Object.prototype.hasOwnProperty.call(
    STORED_ERROR_LITERALS,
    upper,
  )
    ? STORED_ERROR_LITERALS[upper]
    : undefined;
  if (storedError) return storedError;
  // Excel-correct numeric coercion: a JS-only literal (hex 0x1F, 0b/0o, Infinity,
  // NaN) stays TEXT — `Number()` alone would turn "0x1F" into 31, diverging from
  // Excel/Sheets and from the formula-arg path (toNumber) which already guards
  // this. Keeps =A1+1 -> #VALUE! and ISNUMBER(A1) -> FALSE for those literals.
  if (isExcelNumericText(t)) {
    const n = Number(t);
    // Excel keeps no subnormal numbers: an entry such as 1E-310, below the
    // smallest normal double, stays TEXT in the cell (audit P-NUMERIC-SUBNORMAL).
    return isSubnormalNumber(n) ? raw : n;
  }
  return raw;
}

/** Smallest positive normal double (2.2250738585072014E-308). Excel has no
 * subnormal numbers: Microsoft Learn, "Floating-point arithmetic may give
 * inaccurate results in Excel", notes Excel does not support denormalized
 * numbers, so a result below this magnitude is 0. */
export const EXCEL_MIN_NORMAL_NUMBER = 2.2250738585072014e-308;

/** A finite, non-zero number below Excel's smallest normal magnitude. */
export function isSubnormalNumber(value: number): boolean {
  return value !== 0 && Math.abs(value) < EXCEL_MIN_NORMAL_NUMBER;
}

/**
 * The typed value a literal cell's characters are read as when the host does
 * not type the cell as text (no quote prefix, no Text format): "123" is a
 * number, "TRUE" a logical, "#N/A" an error, "" blank and anything else text.
 * Writers that must keep a string a string (Paste Values of a text result)
 * use it to tell when the characters alone would change the type.
 */
export function readLiteralCellValue(raw: string): CellValue {
  return coerceLiteral(raw);
}

export interface SheetEvaluator {
  /**
   * The workbook reads serials in the 1904 date system. Hosts format the
   * evaluator's numbers with it (`applyNumberFormat(v, fmt, { date1904 })`),
   * so a display path that holds an evaluator reads dates the way the
   * formulas behind it do. Absent/false is the 1900 system.
   */
  readonly date1904?: ExcelDateSystemArg;
  /** Computed value at (row, col), evaluating formulas transitively. */
  getComputed(row: number, col: number): CellValue;
  /**
   * The exact raw formula registered for this evaluator generation, or null
   * for a literal/empty cell. Hosts use this instead of rebuilding a reader
   * from live sheet state, which would bypass a manual-calculation snapshot.
   */
  getFormula(row: number, col: number): string | null;
  /**
   * Atomically reserve this start cell's single deep-recovery attempt for the
   * current workbook generation. All evaluator views of the same sheet share
   * the reservation, so asking evaluatorFor() again cannot retry it.
   */
  beginDeepRecovery(row: number, col: number): boolean;
  /** Convenience: the display string for a cell. */
  getDisplay(row: number, col: number): string;
  /**
   * The dynamic-array block a formula cell would SPILL, or null when the cell is
   * not a formula or its result is a single scalar. The grid uses this to lay a
   * spill block over adjacent cells (with a #SPILL! guard) without persisting
   * the spilled cells. A 1×1 block (a scalar array) returns null — there is
   * nothing to spill. Evaluation shares the same memoized, cycle-detecting
   * machinery as getComputed, so refs/circular refs behave identically.
   */
  getArray(row: number, col: number): CellValue[][] | null;
  /**
   * Retry one already-planned, same-sheet dependency chain in this evaluator.
   * The caller must supply a complete, deepest-first bounded plan. Scalar memo
   * and stale cycle-guard entries for the plan/start are evicted together, and
   * the derived range cache is cleared before the plan is warmed bottom-up.
   */
  recomputeWithPrecedents(
    row: number,
    col: number,
    precedents: readonly { row: number; col: number }[],
  ): CellValue;
  /**
   * Evaluate a formula no cell holds -- a conditional format's or a data
   * validation rule's -- as if it stood at (row, col), with everything a cell
   * formula there reads: other sheets, the workbook's names and the sheet's
   * used range for whole rows and columns. Nothing is cached for the cell.
   * `resolveName` answers names when the workbook's sheets do not.
   */
  evaluateFormulaAt?(
    formula: string,
    row: number,
    col: number,
    fallback?: { resolveName?: EvalContext["resolveName"] },
  ): CellValue;
}

/** The engine option over a sheet's enumerated legacy array formulas. */
export function legacyArrayExtentOption(
  entries: ReadonlyArray<readonly [string, number, number]> | undefined,
): Pick<SheetEvaluatorOptions, "legacyArrayExtent"> {
  if (!entries || entries.length === 0) return {};
  const extents = new Map(entries.map(([key, height, width]) => [key, { height, width }]));
  return { legacyArrayExtent: (row: number, col: number) => extents.get(`${row}:${col}`) ?? null };
}

/**
 * A sheet's column widths in characters, as Column Width reads them: the
 * default column's and every other column's ([column index, characters]; a
 * hidden column 0). Plain data, so the background calculation gets the same.
 */
export interface ColumnWidthReading {
  readonly defaultWidth: number;
  readonly columns: ReadonlyArray<readonly [number, number]>;
}

/** The engine option over a sheet's column widths (CELL("width")). */
export function columnWidthOption(
  reading: ColumnWidthReading | undefined,
): Pick<SheetEvaluatorOptions, "columnWidth" | "columnWidthIsDefault"> {
  if (!reading || !Number.isFinite(reading.defaultWidth)) return {};
  const widths = new Map(reading.columns);
  return {
    columnWidth: (col: number) => widths.get(col) ?? reading.defaultWidth,
    // A column the sheet gives no width of its own (nor hides) is the default.
    columnWidthIsDefault: (col: number) => !widths.has(col),
  };
}

/** Optional capabilities a sheet can expose to the evaluator. */
export interface SheetEvaluatorOptions {
  /**
   * Materialized worksheet extent. Whole-row/whole-column references keep
   * Excel's full address dimensions, but aggregate only this bounded rectangle
   * plus implicit trailing blanks. Hosts should pass normalized sheet counts;
   * omitted legacy/test contexts infer the addressed whole-axis strip once,
   * with a hard scan budget, instead of truncating at the formula cell.
   */
  usedRange?: { rowCount: number; columnCount: number };
  /**
   * Resolve a defined name (named range) to its A1 reference string ("B2" or
   * "B2:B10"), or null when undefined. Wired into every formula's EvalContext so
   * names work inside SUM(), arithmetic, etc. Omit when the sheet has no names.
   */
  resolveName?: (name: string) => string | null;
  /**
   * Whether a row (0-indexed) is hidden by an AutoFilter. SUBTOTAL skips such
   * rows for EVERY function_num (1-11 and 101-111). Omit when the sheet has no
   * filter (every row counts).
   */
  isRowHiddenByFilter?: (row: number) => boolean;
  /**
   * Whether a row (0-indexed) is MANUALLY hidden (Excel's "Hide Rows"). Only
   * SUBTOTAL's 101-111 family skips such rows. The native sheet model has no
   * manual row-hide feature yet, so production hosts leave this unwired and
   * 101-111 behave as filter-hidden-only; the seam exists for Excel parity.
   */
  isRowManuallyHidden?: (row: number) => boolean;
  /**
   * The cells whose formulas may spill (the sheet's formula cells). With it, a
   * plain reference to a blank cell inside a dynamic array's footprint reads
   * the element the grid shows there — =G21 is what SEQUENCE spilled into G21
   * and SUM(F20:G22) sums the block, as in Excel — instead of the blank the
   * storage holds. Omitted, only `A1#` sees a spill.
   */
  spillAnchors?: () => Iterable<{ row: number; col: number }>;
  /**
   * Whether the host counts (row, col) as occupied, so a spill stops there the
   * way the grid's own overlay stops (a merge's covered cells, for one).
   * Omitted, a cell with a value occupies and a blank one does not.
   */
  isSpillTargetOccupied?: (row: number, col: number) => boolean;
  /**
   * The size of the LEGACY array formula range (Ctrl+Shift+Enter, a file's
   * `<f t="array" ref="A1:A3">` without a dynamic-array flag) whose anchor is
   * (row, col), or null. Its result is laid over exactly that range, never
   * beyond: a single value repeats in every cell, a single row or column
   * repeats along the other axis, and cells past a larger result show #N/A
   * (Microsoft Support, "Guidelines and examples of array formulas"). Omit
   * for a host with no legacy array formulas.
   */
  legacyArrayExtent?: (row: number, col: number) => { height: number; width: number } | null;
  /**
   * Resolve a STRUCTURED (ListObject) reference — `Table1[Amount]`, `[@Col]`,
   * `Table1[#Totals]` — against THIS sheet's tables, to the A1 box it
   * addresses on this sheet, or a FormulaError. Build one with
   * {@link makeTableRefResolver} over the sheet's tables and pass it here; omit
   * when the sheet has no tables. Inside a workbook ({@link createWorkbook}) a
   * formula also reaches the tables registered on every OTHER sheet (Excel's
   * table names are workbook-scoped) and reads them on their own sheet.
   */
  resolveTableRef?: (
    tableName: string | null,
    spec: StructuredSpec,
    currentCell: { row: number; col: number } | undefined,
  ) => TableRefResult;
  /**
   * Resolve a GETPIVOTDATA lookup against THIS sheet's pivots — build one with
   * lib/spreadsheet/pivot's `makePivotDataResolver` over the sheet's placed pivot
   * output ranges and pass it here. Omit when the sheet has no pivots, in which
   * case any GETPIVOTDATA evaluates to #REF! (there is no pivot registry).
   */
  resolvePivotData?: (
    anchor: { row: number; col: number },
    dataField: string,
    criteria: ReadonlyArray<{ field: string; item: string }>,
  ) => PivotDataLookupResult;
  /**
   * Whether the cell at (row, col) holds TEXT whatever its characters look
   * like. Excel stores a cell's type, not only its characters: an entry made
   * with a leading apostrophe (the cell's quote prefix) or typed into a Text
   * (@) cell is a string even when it reads as "123", "TRUE" or "#N/A"
   * (Microsoft Support, "Enter data manually in worksheet cells"). Such a cell
   * is ignored by SUM and the other aggregates, is TRUE against ="123", and
   * still coerces in arithmetic (=A1+0 is 123) because operators convert
   * numeric text. Omit for a host with no cell types: every literal is then
   * read from its characters alone, as before.
   */
  isTextCell?: (row: number, col: number) => boolean;
  /**
   * The picture placed in the literal cell at (row, col), or null. Such a
   * cell's value is the picture (an {@link ImageCellValue}), not the
   * `#VALUE!` its characters store for readers that cannot draw it. Only the
   * alternative text crosses: the host reads the picture itself from the
   * cell. Omit for a host with no in-cell pictures.
   */
  cellPicture?: (row: number, col: number) => { readonly alt: string } | null;
  /**
   * The phonetic reading of the text cell at (row, col) — its text with the
   * furigana in place of the characters they read, in the cell's character
   * type (lib/spreadsheet/cell-phonetic `cellPhoneticReading`) — or null when
   * the cell has no furigana. PHONETIC(reference) returns it.
   */
  cellPhonetic?: (row: number, col: number) => string | null;
  /**
   * The width of column `col` in characters of the Normal font's digit, as
   * Column Width reads it (two decimals; 0 for a hidden column) --
   * CELL("width") returns it rounded (display-cache `columnWidthEngineOption`).
   */
  columnWidth?: (col: number) => number;
  /** Whether column `col` has the sheet's default width (CELL("width")'s second item). */
  columnWidthIsDefault?: (col: number) => boolean;
  /**
   * The file's own last value of the formula cell at (row, col) when that
   * formula reads external data MRAI cannot reach (an OLAP cube, RTD, another
   * workbook), or undefined. Excel keeps showing that value until a refresh
   * (lib/spreadsheet/kept-formula-results); the cell answers it instead of
   * being evaluated, for itself and every formula reading it.
   */
  keptFormulaResult?: (row: number, col: number) => CellValue | undefined;
  /**
   * What a held result (a formula calling a function this engine does not
   * have, kept-formula-results) read in the file: the areas and the values
   * stored in them. The cell answers its kept result only while each still
   * holds that value; once one does not, the cell is calculated (#NAME?), as
   * Excel recalculates it with its macros off.
   */
  keptFormulaInputs?: (row: number, col: number) => readonly KeptFormulaInput[] | undefined;
  /** Set precision as displayed (calcPr@fullPrecision="0"): the rounding the
   *  cell's number format gives its number, a formula's or a constant's
   *  (displayedPrecisionRounder), null where it keeps the number whole.
   *  Readers of the cell read it rounded. */
  displayPrecision?: (row: number, col: number) => ((value: number) => number) | null;
  /**
   * Whether the formula of the cell at (row, col) cannot be displayed: the
   * sheet's protection is in effect and the cell is formatted Hidden
   * (lib/spreadsheet/protection `isCellContentHidden`). FORMULATEXT of such a
   * cell is #N/A (Microsoft Support, "FORMULATEXT function"); the cell itself
   * calculates as usual. Omit when nothing is hidden.
   */
  isFormulaHidden?: (row: number, col: number) => boolean;
  /**
   * The sheet's What-If Data Tables (lib/spreadsheet/what-if-table-definition).
   * A cell inside one's results whose raw content is its `=TABLE(...)` is
   * that table's result: its formula recalculated with the row's and/or
   * column's value typed into the input cells, every time the workbook
   * calculates. Omit when the sheet has none (TABLE is then unknown).
   */
  whatIfTables?: readonly WhatIfTableDefinition[];
}

/**
 * A workbook accessor the cross-sheet resolver uses to reach OTHER sheets. It is
 * created once via {@link createWorkbook}; every sheet evaluator built from it
 * shares its cache + cycle stack. Sheet-name matching is case-insensitive
 * (Excel parity). The host supplies, per sheet name, the same `getRawCell` +
 * options it would pass to a standalone `createSheetEvaluator`, so the target
 * sheet's named ranges / hidden rows are honored.
 */
export interface Workbook {
  /** True iff a sheet with this name (case-insensitively) exists. */
  hasSheet(sheetName: string): boolean;
  /** Computed scalar at (sheet,row,col); #REF! for an unknown sheet. */
  getComputed(sheetName: string, row: number, col: number): CellValue;
  /** Computed 2D grid for sheet!from:to; [[#REF!]] for an unknown sheet. */
  getRange2D(
    sheetName: string,
    from: string,
    to: string,
  ): CellValue[][];
  /**
   * The sheet names in workbook order — used to expand a 3D span Sheet1:Sheet3.
   * Names are returned as registered (display case).
   */
  sheetOrder(): readonly string[];
}

/** One sheet's inputs to a workbook: its raw-cell reader + evaluator options. */
export interface WorkbookSheetInput {
  name: string;
  getRawCell: (row: number, col: number) => string | null;
  options?: SheetEvaluatorOptions;
}

/**
 * Build a workbook over MULTIPLE sheets that share ONE memoization cache and ONE
 * cross-sheet cycle-detection stack. Returns the workbook accessor plus a
 * factory for each sheet's public {@link SheetEvaluator}. This is the entry
 * point a multi-sheet host (SpreadsheetWindow) uses so Sheet2!A1 reaches the
 * real Sheet2 grid and a cross-sheet cycle is caught.
 */
export interface WorkbookOptions {
  /** The document's file name, for BOOKNAME(). Omitted leaves BOOKNAME() at #N/A
   *  rather than guessing — the engine has no filesystem of its own. */
  workbookName?: string;
  /**
   * The folder of the document's location, ending in "/": CELL("filename")
   * answers folder + "[Book.xlsx]Sheet", as Excel for the web answers the
   * URL folder of the document (X-194). Omitted, "[Book.xlsx]Sheet" alone.
   */
  workbookFolder?: string;
  /**
   * File > Options > Formulas > Enable iterative calculation (SpreadsheetML
   * `calcPr@iterate`, `@iterateCount`, `@iterateDelta`): a circular reference
   * is calculated by repeated passes, each reading the values the previous
   * pass left, until no cell of the cycle changes by `maxChange` or more, or
   * `maxIterations` passes ran (Microsoft Support, "Remove or allow a
   * circular reference"; ECMA-376 defaults 100 and 0.001). Omitted, a
   * circular reference is #CIRCULAR!.
   */
  iteration?: {
    readonly maxIterations: number;
    readonly maxChange: number;
    /**
     * Each cell's current value, keyed as the engine keys a cell
     * ("sheet-lowercase\u0000row:col"): Excel starts every iteration from the
     * values the cells hold -- the last calculation's, or the file's cached
     * ones -- not from blank. Omitted cells start at 0.
     */
    readonly initialValues?: ReadonlyMap<string, EngineValue>;
  };
  /**
   * The workbook uses the 1904 date system (page content `dateSystem:
   * "1904"`, SpreadsheetML `workbookPr@date1904`). Every formula of the
   * workbook reads its serials as 1904 dates; omitted is the 1900 system.
   */
  date1904?: ExcelDateSystemArg;
  /**
   * Where the pictures IMAGE formulas name stand: the host that fetches them
   * answers loading (#BUSY!), refused (#CONNECT! / #BLOCKED! / #VALUE!) or
   * ready. IMAGE is volatile, so a warm start recomputes it after a fetch.
   * Omitted: every valid IMAGE source reads as its picture value.
   */
  webImageStatus?: (source: string) => WebImageStatus | null;
  /**
   * The data WEBSERVICE and IMPORTRANGE read (`external-data`): the session
   * holds what the browser fetched; a call it does not hold reads #BUSY!.
   * Both are volatile, so a warm start recomputes them when it changes.
   */
  externalData?: ExternalDataLookup;
  /** Formulas parsed already, by text (a calculation session's, kept across
   * its calculations): a formula found here is not tokenized and parsed
   * again. The parse of a text is the same wherever it stands. */
  parsedFormulas?: Map<string, unknown>;
  /**
   * Reuse the materialized grid when the same range is read more than once.
   *
   * Cells are already memoized, but the ARRAY a range assembles is not: a sheet
   * where two hundred summary formulas each read `$A$1:$A$5000` walks and
   * rebuilds that same five-thousand-cell grid two hundred times. Measured on
   * such a sheet, one edit cost 93ms of pure re-materialization.
   *
   * Values cannot change inside one workbook — every cell is computed at most
   * once and the workbook is rebuilt from scratch on the next edit — so a grid
   * cached here is correct for as long as it can be reached. Off by default
   * while the flag is dark.
   */
  cacheRangeValues?: boolean;
  /**
   * Optional bound for one workbook evaluation. Each distinct computed cell
   * consumes one slot, and a range must fit before its grid is traversed or
   * allocated. Omitted preserves the historical unbounded evaluator behavior.
   */
  maxComputedCells?: number;
  /**
   * Exact formula results published by a prior authenticated worksheet-only
   * calculation. Only matching live formula cells consume an override; literals
   * and every formula outside this map continue to evaluate normally.
   */
  formulaResultOverrides?: ReadonlyMap<string, EngineValue>;
  /** Display-only background-calculation projection. Formula cells missing
   * from overrides remain blank instead of synchronously running the formula.
   * Never use this evaluator for an authoritative save, export or command.
   * Those consumers must wait for a complete, version-checked result snapshot. */
  deferMissingFormulaResults?: boolean;
  /**
   * Carry the previous workbook's memo across an edit, minus everything the
   * edit could have changed.
   *
   * The workbook is immutable once built, so every edit rebuilds it and drops
   * the memo — correct, and wasteful: one keystroke in a constants column
   * recomputes every formula on the sheet. With this option the new workbook
   * starts from the previous memo with the dirty set removed, where dirty =
   * changed cells ∪ unprovable formulas ∪ everything reachable from either
   * (see {@link warmStartDirtyKeys}). A formula whose reads cannot be proven
   * complete — volatile, INDIRECT/OFFSET, cross-sheet, defined names,
   * structured refs — is dirty on every edit, which is always sound, never
   * fast-but-wrong. If the propagation cannot be bounded, nothing is seeded.
   *
   * The caller owns the equally important other half: only offer this when
   * everything EXCEPT cell values is identical to the snapshot's build (same
   * sheets, same names/merges/hidden rows/tables, no pivot overlay).
   */
  warmStart?: WorkbookWarmStart;
  /**
   * Semantic formula cache used by worksheet-scoped recalculation. Frozen
   * sheets use only these formula results; literals remain current. Unlike
   * warmStart, this never computes through a missing frozen formula.
   */
  formulaResultSeed?: WorkbookFormulaResultSeed;
  /**
   * Internal: this workbook recalculates one What-If Data Table result
   * (formula-evaluator, TABLE). `overrides` are the input cells' typed
   * values by memo key, read before the cells themselves; a table result
   * read inside it is #N/A (tables do not nest); `budget` counts the cells
   * every such recalculation of the outer workbook may compute.
   */
  whatIfProbe?: {
    readonly overrides: ReadonlyMap<string, EngineValue>;
    readonly budget: { left: number; readonly limit: number };
  };
}

// ---------------------------------------------------------------------------
// Dynamic-array spill materialization (#SPILL! guard)
// ---------------------------------------------------------------------------

/**
 * The materialized overlay of every dynamic-array spill on a sheet. Keys are
 * "row:col". `spilled` holds the value displayed at each NON-ANCHOR target
 * cell; `blocked` holds the anchors whose footprint collided with existing
 * content — those render ERRORS.SPILL ("#SPILL!") instead of their top-left
 * element, and contribute NOTHING to `spilled` (no partial writes).
 */
export interface SpillOverlay {
  spilled: Map<string, CellValue>;
  blocked: Set<string>;
}

/**
 * Compute the spill overlay for a sheet (Excel 365 #SPILL! semantics). A
 * formula whose result is a multi-cell array claims its rows×cols footprint
 * from the anchor; if ANY target beyond the anchor is occupied — the host's
 * own content (`isOccupied`: a non-empty cell, merge cover, …) or a cell
 * already claimed by an earlier spill — the anchor is blocked (#SPILL!) and
 * nothing spills from it. Blocking content includes constants and other
 * formulas' cells alike. A formula's OWN previous spill never blocks it: the
 * overlay is recomputed from scratch on every call and spilled targets are
 * never written into the host's cell storage, so re-spilling over the prior
 * footprint — or recalculating after the blocking cell is emptied — resolves
 * cleanly. A 1×1 result never blocks (`getArray` returns null — nothing to
 * spill). Anchors are processed in deterministic (row, col) order so
 * overlapping-spill resolution is stable across recalculations.
 */
export function computeSpillOverlay(params: Parameters<typeof computeSpillOverlaySteps>[0]): SpillOverlay {
  const steps = computeSpillOverlaySteps(params);
  for (;;) { const next = steps.next(); if (next.done) return next.value; }
}

/** Same spill algorithm with cancellation/yield boundaries for background preparation. */
export function* computeSpillOverlaySteps(params: {
  /** Candidate anchor cells (the sheet's formula cells), in any order. */
  anchors: ReadonlyArray<{ row: number; col: number }>;
  /** An anchor's array block, or null when there is nothing to spill
   *  (a scalar formula or a 1×1 array) — SheetEvaluator.getArray's contract. */
  getArray(row: number, col: number): CellValue[][] | null;
  /** Whether the host considers (row, col) occupied by its own content. The
   *  anchor cell itself is never asked (it owns its block). */
  isOccupied(row: number, col: number): boolean;
  /** An anchor that may never spill where it stands: a formula inside a
   *  table (Excel: a table cannot hold a spilled array) reads #SPILL!
   *  whatever its footprint would find (live audit B323). */
  isAnchorBlocked?(row: number, col: number): boolean;
  /** Each spilled cell as it is laid down, before the next anchor's array is
   *  asked for: an anchor below one it reads finds that one's values already. */
  onLaid?(row: number, col: number, value: CellValue): void;
}): Generator<void, SpillOverlay> {
  const { anchors, getArray, isOccupied, isAnchorBlocked, onLaid } = params;
  const spilled = new Map<string, CellValue>();
  const blocked = new Set<string>();
  // The anchor that claimed each target, so two spills can never overlap and a
  // later anchor can never spill onto an earlier anchor's claimed cell.
  const claimedBy = new Map<string, string>();
  // Cap on a single spill's footprint so a pathological SEQUENCE(1e6) can never
  // build a runaway overlay map (the engine already caps the array itself).
  const MAX_SPILL_CELLS = 50_000;
  const ordered = [...anchors]
    .filter(({ row, col }) => Number.isFinite(row) && Number.isFinite(col))
    .sort((a, b) => a.row - b.row || a.col - b.col);
  let work = 0;
  for (const { row, col } of ordered) {
    if (++work % 64 === 0) yield;
    const key = `${row}:${col}`;
    const block = getArray(row, col);
    if (!block) continue; // scalar formula (or 1×1) — nothing to spill
    const height = block.length;
    const width = block[0]?.length ?? 0;
    if (height * width > MAX_SPILL_CELLS || isAnchorBlocked?.(row, col)) {
      blocked.add(key);
      continue;
    }
    // First pass: does any non-anchor target collide with existing content?
    let collision = false;
    for (let r = 0; r < height && !collision; r += 1) {
      for (let c = 0; c < width; c += 1) {
        if (++work % 512 === 0) yield;
        const tk = `${row + r}:${col + c}`;
        if (tk === key) continue; // the anchor's own cell is exempt
        if (claimedBy.has(tk) || isOccupied(row + r, col + c)) {
          collision = true;
          break;
        }
      }
    }
    if (collision) {
      blocked.add(key);
      continue;
    }
    // Second pass: lay the block down (skip the anchor — it renders top-left).
    for (let r = 0; r < height; r += 1) {
      for (let c = 0; c < width; c += 1) {
        if (++work % 512 === 0) yield;
        const tk = `${row + r}:${col + c}`;
        claimedBy.set(tk, key);
        if (r === 0 && c === 0) continue;
        spilled.set(tk, block[r][c] ?? null);
        onLaid?.(row + r, col + c, block[r][c] ?? null);
      }
    }
  }
  return { spilled, blocked };
}

/** Format a computed value for display in a cell. */
export function formatCellValue(v: CellValue): string {
  // A picture's text is Excel's own fallback for a reader that cannot draw it.
  if (isPictureValue(v)) return v.code;
  if (isFormulaError(v)) return v.code;
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") {
    if (!Number.isFinite(v)) return "#NUM!";
    // The same General text the number formatter writes, so an unformatted cell
    // and a General-formatted one cannot disagree.
    return formatGeneralNumber(v);
  }
  return v;
}

// ---------------------------------------------------------------------------
// Function names (autocomplete, name validation)
// ---------------------------------------------------------------------------

const KNOWN_FUNCTION_NAMES = new Set<string>(FORMULA_FUNCTION_NAMES);

/**
 * The names of every function the engine can evaluate -- the source behind
 * formula autocomplete/intellisense. A fresh sorted array (uppercase,
 * including dotted names like "STDEV.S") a caller may mutate. Generated from
 * the evaluator's table (formula-function-names.ts).
 */
export function getFunctionNames(): string[] {
  return [...FORMULA_FUNCTION_NAMES];
}

/** True iff `name` (case-insensitive) is a known engine function. */
export function isKnownFunction(name: string): boolean {
  return KNOWN_FUNCTION_NAMES.has(stripFutureFunctionPrefix(name.toUpperCase()));
}
