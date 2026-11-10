/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for five daily Excel/WPS grid interactions whose logic is kept
 * OUT of the React component so each can be unit-tested in isolation:
 *
 *  - Fill-handle DOUBLE-CLICK auto-fill: double-clicking the fill handle fills
 *    the selection down to the bottom of the adjacent data (Excel detects the
 *    extent of the contiguous neighbour column). `autoFillDownExtent` returns
 *    the bottom row the fill should reach (or null when there is nothing to
 *    follow), reading cells through a `hasValue(row, col)` probe.
 *  - Ctrl+Enter fill selection: entering a value and pressing Ctrl/Cmd+Enter
 *    fills the WHOLE current selection with that value in one undo step.
 *    `fillSelectionGrid` builds the rows×cols grid of the entered value.
 *  - Ctrl+; insert date / Ctrl+Shift+; insert time: `formatDateForCell` /
 *    `formatTimeForCell` turn a `Date` into the ISO `YYYY-MM-DD` / 24-hour
 *    `HH:mm:ss` text the formula engine recognises (DATEVALUE / TIMEVALUE). The
 *    `now` is threaded in as a parameter so the helpers stay pure; the handler
 *    passes `new Date()`.
 *  - Internal workbook hyperlinks (`#Sheet2!A1` / `#'My Sheet'!A1:B3`): a link
 *    whose target is a cell/range on ANOTHER sheet rather than a URL.
 *    `formatInternalLink` builds the stored `#…` string and `parseInternalLink`
 *    decodes it back to { sheetName, ref } so the click handler can jump there.
 *  - Trace Precedents / Dependents (formula-audit arrows): `extractPrecedents`
 *    parses a formula's same-sheet cell/range references into the set of source
 *    cells; `extractSheetQualifiedPrecedents` does the same for references that
 *    explicitly name another worksheet; `findDependents` scans the sheet for
 *    formulas that reference a given cell. The renderer turns these plain
 *    coordinates into arrows.
 *
 * Every helper is side-effect free. A1 parsing / column-letter formatting reuse
 * the formula engine's shared helpers so the reference grammar agrees with the
 * evaluator everywhere.
 */

import {
  columnToLetters,
  formatSheetNameForFormula,
  lettersToColumn,
  parseCellRef,
  shiftFormula,
  type CellRef,
} from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import { splitSheetQualifier } from "./name-syntax";
import { execBehind } from "../regexp-behind";

// ---------------------------------------------------------------------------
// Fill-handle double-click — auto-fill down to the bottom of adjacent data
// ---------------------------------------------------------------------------

export interface SelectionRect {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/**
 * The bottom row a fill-handle double-click should extend the selection to, or
 * null when there is nothing to follow (no adjacent data below the selection).
 *
 * Excel rule: double-clicking the fill handle fills DOWN to match the run of
 * the neighbouring column. We measure the contiguous filled run, starting at
 * the first row BELOW the selection, of the column immediately to the LEFT of
 * the selection; if that column is empty there (or the selection touches the
 * left edge) we fall back to the column immediately to the RIGHT. The fill
 * stops at the last row where the chosen neighbour column still has a value.
 *
 * Returns a row STRICTLY below `sel.r2` (so the caller always extends the fill)
 * — null when neither neighbour has data on the row just below the selection,
 * matching Excel (a double-click with no adjacent data does nothing). The
 * returned row is clamped to the last grid row.
 *
 * @param sel       the current selection rectangle (the fill source block).
 * @param hasValue  probe: does cell (row, col) hold a non-empty value?
 * @param rowCount  the sheet's row count (the inclusive scan upper bound).
 */
export function autoFillDownExtent(
  sel: SelectionRect,
  hasValue: (row: number, col: number) => boolean,
  rowCount: number,
): number | null {
  const firstBelow = sel.r2 + 1;
  const lastRow = Math.min(rowCount, MAX_GRID_ROWS) - 1;
  if (firstBelow > lastRow) return null;

  // Pick the neighbour column to follow: the column to the LEFT of the
  // selection when it exists and has data on the row just below the selection,
  // else the column to the RIGHT. (Excel prefers the left neighbour.)
  const leftCol = sel.c1 - 1;
  const rightCol = sel.c2 + 1;
  let probeCol: number | null = null;
  if (leftCol >= 0 && hasValue(firstBelow, leftCol)) {
    probeCol = leftCol;
  } else if (rightCol < MAX_GRID_COLUMNS && hasValue(firstBelow, rightCol)) {
    probeCol = rightCol;
  }
  if (probeCol === null) return null;

  // Walk down the neighbour column from the first row below the selection while
  // it stays filled; the last filled row is how far the fill reaches.
  let target = firstBelow;
  for (let r = firstBelow; r <= lastRow; r += 1) {
    if (!hasValue(r, probeCol)) break;
    target = r;
  }
  return target > sel.r2 ? target : null;
}

// ---------------------------------------------------------------------------
// Ctrl+Enter — fill the whole selection with the entered value
// ---------------------------------------------------------------------------

/**
 * Build the rows×cols grid that fills the selection rectangle with a single
 * `value` (Ctrl/Cmd+Enter). The grid is anchored at (sel.r1, sel.c1) and spans
 * every selected row and column, so writing it through the shared paste path
 * fills the entire selection in one undo step.
 *
 * A plain value is copied verbatim into every cell. A FORMULA (leading `=`) has
 * its relative references rebased PER CELL relative to `anchor` (the active cell
 * the user typed into) — Excel parity: selecting B1:B3, typing `=A1` and pressing
 * Ctrl+Enter yields B1=`=A1`, B2=`=A2`, B3=`=A3`. Absolute ($) refs are left
 * alone by shiftFormula. When no `anchor` is given the formula is copied
 * verbatim (no rebasing).
 */
export function fillSelectionGrid(
  sel: SelectionRect,
  value: string,
  anchor?: CellRef,
): string[][] {
  const isFormula = value.trim().startsWith("=");
  const rows: string[][] = [];
  for (let r = sel.r1; r <= sel.r2; r += 1) {
    const cols: string[] = [];
    for (let c = sel.c1; c <= sel.c2; c += 1) {
      if (isFormula && anchor) {
        cols.push(shiftFormula(value, r - anchor.row, c - anchor.col));
      } else {
        cols.push(value);
      }
    }
    rows.push(cols);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Ctrl+; date / Ctrl+Shift+; time — formatted "now" insertion
// ---------------------------------------------------------------------------

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/**
 * The ISO `YYYY-MM-DD` date text for `now` (LOCAL date, Excel's Ctrl+; inserts
 * the user's calendar date). The format is one DATEVALUE recognises, so the
 * inserted text behaves as a date in formulas. Threaded `now` keeps it pure.
 */
export function formatDateForCell(now: Date): string {
  return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
}

/**
 * The 24-hour `HH:mm:ss` time text for `now` (LOCAL time, Excel's Ctrl+Shift+;
 * inserts the current time). The format is one TIMEVALUE recognises. Threaded
 * `now` keeps it pure.
 */
export function formatTimeForCell(now: Date): string {
  return `${pad2(now.getHours())}:${pad2(now.getMinutes())}:${pad2(now.getSeconds())}`;
}

// ---------------------------------------------------------------------------
// Internal workbook hyperlinks — #Sheet2!A1 / #'My Sheet'!A1:B3
// ---------------------------------------------------------------------------

export interface InternalLinkTarget {
  /** Target sheet name, unquoted (matched case-insensitively by the caller). */
  sheetName: string;
  /** The A1 cell or range reference on that sheet, e.g. "A1" or "A1:B3". */
  ref: string;
}

/**
 * Build the stored internal-link string for a sheet + A1 ref ("#Sheet2!A1").
 * The sheet is quoted exactly when a formula would quote it (Excel's rule, see
 * sheetNameNeedsQuotes): `#回款!A1`, `#'My Sheet'!A1`, `#'O''Brien'!A1`.
 */
export function formatInternalLink(sheetName: string, ref: string): string {
  return `#${formatSheetNameForFormula(sheetName)}!${ref}`;
}

/**
 * A link to a defined name (`#Budget`): Excel stores the name itself as the
 * hyperlink's location, so the link follows the name when its range is
 * redefined, where `#Sheet1!A1` stays on that cell (live audit B359). Returns
 * the name, or null when the text is not a bare name (a sheet-qualified place
 * and a plain cell reference are parseInternalLink's).
 */
export function parseNamedLink(raw: unknown): string | null {
  if (typeof raw !== "string" || !raw.startsWith("#")) return null;
  const body = raw.slice(1);
  // A name scoped to a sheet is written with the sheet before it, as a
  // formula names one (`#'Annex G'!_ftnref1`, Word's footnote back-links);
  // it comes back whole for findScopedName to read. Refused, it failed the
  // page schema and the whole workbook opened read-only (xlsx-rounds
  
  const { sheet, ref: name } = splitSheetQualifier(body);
  if (sheet !== undefined && sheet === "") return null;
  if (name.length === 0 || name.length > 255) return null;
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u.test(name)) return null;
  if (parseCellRef(name) !== null) return null;
  return sheet === undefined ? name : body;
}

// Validate an A1 single-cell or range reference (each side optionally
// $-absolute) — the same grammar the formula engine's parseCellRef accepts,
// extended to a "A1:B3" range. Returns the normalised (uppercased) ref or null.
function normalizeRef(raw: string): string | null {
  const trimmed = raw.trim();
  if (trimmed === "") return null;
  const parts = trimmed.split(":");
  if (parts.length > 2) return null;
  const cells = parts.map((p) => parseCellRef(p));
  if (cells.some((c) => c === null)) return null;
  return trimmed.toUpperCase();
}

/**
 * Decode a stored internal-link string ("#Sheet2!A1", "#'My Sheet'!A1:B3") to
 * its { sheetName, ref }. Returns null when the string is not a well-formed
 * internal link — including a plain URL (no leading `#`), a missing `!`
 * separator, an unparseable sheet name, or an invalid A1 reference — so the
 * caller can fall back to treating it as an external URL.
 */
export function parseInternalLink(raw: unknown): InternalLinkTarget | null {
  if (typeof raw !== "string") return null;
  if (!raw.startsWith("#")) return null;
  const body = raw.slice(1);
  if (body === "") return null;

  let sheetName: string;
  let rest: string;
  if (body.startsWith("'")) {
    // Quoted sheet name: read to the closing quote, honouring '' as an escaped
    // quote. The character after the closing quote must be '!'.
    let i = 1;
    let name = "";
    let closed = false;
    while (i < body.length) {
      const ch = body[i];
      if (ch === "'") {
        if (body[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        closed = true;
        i += 1;
        break;
      }
      name += ch;
      i += 1;
    }
    if (!closed || body[i] !== "!") return null;
    sheetName = name;
    rest = body.slice(i + 1);
  } else {
    const bang = body.indexOf("!");
    if (bang <= 0) return null;
    sheetName = body.slice(0, bang);
    rest = body.slice(bang + 1);
  }

  if (sheetName === "") return null;
  const ref = normalizeRef(rest);
  if (ref === null) return null;
  return { sheetName, ref };
}

/** How many cells of `sheets` hold an internal link naming `sheetName` (case
 * apart, as a link resolves): a link's target is its text, so a rename leaves
 * them on the old name. */
export function internalLinksToSheet(
  sheets: ReadonlyArray<{ readonly cells?: Readonly<Record<string, { hyperlink?: unknown } | null | undefined>> }>,
  sheetName: string,
): number {
  const wanted = sheetName.toLowerCase();
  let count = 0;
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells ?? {})) {
      if (parseInternalLink(cell?.hyperlink)?.sheetName.toLowerCase() === wanted) count += 1;
    }
  }
  return count;
}

/** True when a stored hyperlink string is an internal workbook link (#…). */
export function isInternalLink(raw: unknown): boolean {
  return parseInternalLink(raw) !== null;
}

/**
 * Resolve an internal link's target sheet to its index in `sheetNames` and the
 * top-left cell of its reference, clamped to that sheet's extent. Sheet match
 * is case-insensitive (Excel parity). Returns null when the link is malformed
 * or names a sheet that no longer exists.
 *
 * @param link        the stored internal-link string ("#Sheet2!A1").
 * @param sheetNames  the workbook's sheet names, in tab order.
 * @param extentOf    the (rowCount, columnCount) of the sheet at a given index
 *                    (so the jump target never lands past the rendered grid).
 */
export function resolveInternalLinkTarget(
  link: unknown,
  sheetNames: readonly string[],
  extentOf: (sheetIndex: number) => { rowCount: number; columnCount: number },
): { sheetIndex: number; row: number; col: number } | null {
  const parsed = parseInternalLink(link);
  if (!parsed) return null;
  // A sheet is named by every character of its name: " Report " and
  // "Report" are two sheets, and a link to one must not land on the other
  // (live audit B370). Only when no sheet carries the name as written does a
  // name that differs by outer spaces stand in for it.
  const named = parsed.sheetName.toLowerCase();
  const exact = sheetNames.findIndex((name) => name.toLowerCase() === named);
  const want = named.trim();
  const sheetIndex = exact >= 0 ? exact : sheetNames.findIndex(
    (name) => name.trim().toLowerCase() === want,
  );
  if (sheetIndex < 0) return null;
  // The top-left cell of the reference (range or single cell).
  const first = parsed.ref.split(":")[0];
  const cell = parseCellRef(first);
  if (!cell) return null;
  // The link names a worksheet address, and the grid shows any legal one on
  // demand: the target is clamped to the sheet's address space, never to the
  // rows and columns that happen to hold data (live audit B176: A200 on a
  // 45-row extent jumped to A45). `extentOf` stays for hosts that pass it.
  const extent = extentOf(sheetIndex);
  const maxRow = Math.max(0, Math.max(extent.rowCount, MAX_GRID_ROWS) - 1);
  const maxCol = Math.max(0, Math.max(extent.columnCount, MAX_GRID_COLUMNS) - 1);
  // An address past the worksheet (A1048577) is no place at all: not a
  // link, rather than one that lands on some other cell (live audit S-137:
  // it was accepted and landed on the data's last row).
  if (cell.row > maxRow || cell.col > maxCol) return null;
  return { sheetIndex, row: cell.row, col: cell.col };
}

// ---------------------------------------------------------------------------
// Trace Precedents / Dependents — formula-audit arrow extraction
// ---------------------------------------------------------------------------

// A SAME-SHEET A1 cell or range reference inside a formula. Sheet-qualified
// refs (Sheet2!A1) and table refs are intentionally excluded — trace arrows are
// drawn on the active sheet only, so a cross-sheet precedent has no on-sheet
// arrow (Excel draws those to a sheet icon, which this grid does not render).
// The leading boundary FORMULA_CELL_OR_RANGE_BEHIND (no character of
// [\w$.!':] just before) keeps "A1" inside "TA1" / "Sheet1!A1" / a function
// name from matching; the trailing `(?![\w(])` rejects a function call like
// "SUM(" and a longer token. The boundary is checked in code (execBehind), not
// as a lookbehind, which Safari before 16.4 cannot compile.
const FORMULA_CELL_OR_RANGE_BEHIND = { char: /[\w$.!':]/, not: true } as const;
const FORMULA_CELL_OR_RANGE_RE =
  /(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})(?::(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?(?![\w(])/g;

// An A1 cell/range explicitly scoped to a worksheet in THIS workbook. External
// workbook refs (`[Book.xlsx]Sheet1!A1`) and 3D refs (`Sheet1:Sheet3!A1`) are
// rejected by the boundary check in extractSheetQualifiedPrecedents. Excel
// requires sheet names containing spaces/punctuation to be single-quoted and
// escapes an apostrophe by doubling it, both handled here. A bare name is a
// letter of any script (\u56de\u6b3e, \u041b\u0438\u0441\u04421, Donn\u00e9es) \u2014 the formula engine's charset.
const FORMULA_SHEET_QUALIFIED_CELL_OR_RANGE_RE =
  /('(?:[^']|'')+'|[\p{L}_][\p{L}\p{M}\p{N}_.]*)!(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})(?:\s*:\s*(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}))?(?![\w(])/gu;

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/**
 * The set of PRECEDENT cells a formula reads — every same-sheet cell named
 * directly or covered by a range, de-duplicated and returned as { row, col }
 * coordinates. A non-formula value (no leading `=`) has no precedents.
 *
 * Ranges are expanded to their covered cells, bounded to the grid maximums and
 * capped so a whole-column reference (A:A → up to 1,048,576 cells) cannot blow
 * up the arrow set; `maxCells` (default 1000) caps the expansion. String
 * literals inside the formula are stripped first so a quoted "A1" is not
 * mistaken for a reference.
 */
export function extractPrecedents(
  formula: string | undefined,
  maxCells = 1000,
): CellRef[] {
  if (typeof formula !== "string" || !formula.startsWith("=")) return [];
  // Strip strings and whole sheet-qualified tokens before looking for local
  // refs. Masking the complete token also prevents the tail of a spaced range
  // such as Sheet2!A1 : A2 from being projected onto the current sheet.
  const body = formula
    .slice(1)
    .replace(/"(?:[^"]|"")*"/g, " ")
    .replace(FORMULA_SHEET_QUALIFIED_CELL_OR_RANGE_RE, " ");
  const seen = new Set<string>();
  const out: CellRef[] = [];
  let match: RegExpExecArray | null;
  FORMULA_CELL_OR_RANGE_RE.lastIndex = 0;
  while ((match = execBehind(FORMULA_CELL_OR_RANGE_RE, body, FORMULA_CELL_OR_RANGE_BEHIND)) !== null) {
    const start = parseCellRef(match[1]);
    if (!start) continue;
    if (match[2]) {
      const end = parseCellRef(match[2]);
      if (!end) continue;
      const r1 = Math.min(start.row, end.row);
      const r2 = Math.min(Math.max(start.row, end.row), MAX_GRID_ROWS - 1);
      const c1 = Math.min(start.col, end.col);
      const c2 = Math.min(Math.max(start.col, end.col), MAX_GRID_COLUMNS - 1);
      for (let r = r1; r <= r2; r += 1) {
        for (let c = c1; c <= c2; c += 1) {
          const key = cellKey(r, c);
          if (seen.has(key)) continue;
          seen.add(key);
          out.push({ row: r, col: c });
          if (out.length >= maxCells) return out;
        }
      }
    } else {
      if (start.row >= MAX_GRID_ROWS || start.col >= MAX_GRID_COLUMNS) continue;
      const key = cellKey(start.row, start.col);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(start);
      if (out.length >= maxCells) return out;
    }
  }
  return out;
}

/** A cell/range precedent whose formula explicitly names a worksheet. */
export interface SheetQualifiedCellRef extends CellRef {
  sheetName: string;
}

function unquoteFormulaSheetName(token: string): string {
  return token.startsWith("'") && token.endsWith("'")
    ? token.slice(1, -1).replace(/''/g, "'")
    : token;
}

/**
 * The workbook-local, sheet-qualified A1 cells read by a formula. Ranges are
 * expanded and the whole result is de-duplicated/capped exactly like
 * extractPrecedents. Unqualified refs are intentionally absent: callers merge
 * this result with extractPrecedents when they need both relationship kinds.
 *
 * External-workbook, 3D, structured, named and dynamic/INDIRECT references are
 * fail-closed. The host resolves sheetName against the CURRENT workbook before
 * creating a relationship, so a stale or missing worksheet never gains an
 * arrow/navigation target.
 */
export function extractSheetQualifiedPrecedents(
  formula: string | undefined,
  maxCells = 1000,
): SheetQualifiedCellRef[] {
  if (
    typeof formula !== "string" ||
    !formula.startsWith("=") ||
    maxCells <= 0
  ) {
    return [];
  }
  const body = formula.slice(1).replace(/"(?:[^"]|"")*"/g, " ");
  const seen = new Set<string>();
  const out: SheetQualifiedCellRef[] = [];
  let match: RegExpExecArray | null;
  FORMULA_SHEET_QUALIFIED_CELL_OR_RANGE_RE.lastIndex = 0;
  while (
    (match = FORMULA_SHEET_QUALIFIED_CELL_OR_RANGE_RE.exec(body)) !== null
  ) {
    // Do not reinterpret the tail of [Book]Sheet!A1 or Sheet1:Sheet3!A1 as a
    // workbook-local reference. Excel sheet names cannot contain brackets.
    const previous = body[match.index - 1] ?? "";
    if (
      previous &&
      (/[\w$.!']/u.test(previous) || previous === "]" || previous === ":")
    ) {
      continue;
    }
    const sheetName = unquoteFormulaSheetName(match[1]);
    if (
      !sheetName ||
      sheetName.includes(":") ||
      sheetName.includes("[") ||
      sheetName.includes("]")
    ) {
      continue;
    }
    const start = parseCellRef(match[2]);
    if (!start) continue;
    const end = match[3] ? parseCellRef(match[3]) : start;
    if (!end) continue;
    const r1 = Math.min(start.row, end.row);
    const r2 = Math.min(Math.max(start.row, end.row), MAX_GRID_ROWS - 1);
    const c1 = Math.min(start.col, end.col);
    const c2 = Math.min(Math.max(start.col, end.col), MAX_GRID_COLUMNS - 1);
    const sheetKey = sheetName.trim().toLowerCase();
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) {
        const key = `${sheetKey}\u0000${cellKey(row, col)}`;
        if (seen.has(key)) continue;
        seen.add(key);
        out.push({ sheetName, row, col });
        if (out.length >= maxCells) return out;
      }
    }
  }
  return out;
}

/** A cell with its formula, as supplied to {@link findDependents}. */
export interface FormulaCellEntry {
  row: number;
  col: number;
  formula: string;
}

/**
 * The set of DEPENDENT cells — every formula cell on the sheet whose precedents
 * include the target (row, col). Returned as de-duplicated { row, col } coords.
 * `formulaCells` is the list of cells that carry a formula (the caller filters
 * the sheet down to these); each is expanded via {@link extractPrecedents} so
 * the dependency test matches the precedent grammar exactly. `maxCells` caps
 * the result for a heavily-referenced cell.
 */
export function findDependents(
  target: CellRef,
  formulaCells: readonly FormulaCellEntry[],
  maxCells = 1000,
): CellRef[] {
  const out: CellRef[] = [];
  for (const entry of formulaCells) {
    if (entry.row === target.row && entry.col === target.col) continue;
    const precedents = extractPrecedents(entry.formula);
    if (
      precedents.some((p) => p.row === target.row && p.col === target.col)
    ) {
      out.push({ row: entry.row, col: entry.col });
      if (out.length >= maxCells) break;
    }
  }
  return out;
}

// Re-export the column-letter helpers so a caller importing trace utilities can
// label an arrow endpoint (A1 coordinate) without a second import.
export { columnToLetters, lettersToColumn };
