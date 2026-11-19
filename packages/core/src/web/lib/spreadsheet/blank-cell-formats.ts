/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nearestKnownCellMap } from "./cell-map-lineage";
import { isMaterializedLineStyleBlank } from "./line-styles";
/**
 * Formatted blank cells stored as rectangular ranges (`sheet.blankCellFormats`).
 *
 * SpreadsheetML formats a cell in three places (ECMA-376 Part 1 §18.3.1):
 * a `<c s>` record with no value (a formatted blank), a `<row s
 * customFormat="1">` style for the row's unrecorded cells, and a `<col style>`
 * default "for cells not yet allocated in the column(s)". The cell's own record
 * wins over its row, which wins over its column. Both importers project the
 * row/column formats onto the used data range, so every formatted blank becomes
 * an explicit cell; a template that formats a large block ahead of data entry
 * then produced one page entry per cell (75,476 in one real workbook).
 *
 * This module is the browser half of the contract shared with
 * `apps/api/app/services/spreadsheet_blank_formats.py`:
 *
 *   "blankCellFormats": [
 *     { "style": {...}, "locked": false, "ranges": ["B2:Z3000", "AB7"] }
 *   ]
 *
 * - A position with no `cells` entry inside a range reads exactly as the entry
 *   `{ value: "", kind: "text", ...group }` (without `ranges`) of the FIRST
 *   group/range in array order that covers it.
 * - An explicit `cells` entry always wins, whole. Typing into a formatted blank
 *   writes an entry that already carries the format, the way Excel keeps a
 *   cell's format when a value is entered (and on Clear Contents).
 * - A group carries a non-empty `style`, `locked: false`, or both. Ranges are
 *   canonical A1 (`A1` / `A1:C3`, upper case, no `$`, start before end) inside
 *   the declared grid, within the per-sheet group/range/area budgets and the
 *   workbook area budget.
 *
 * The editor keeps its established explicit-cell model in memory: loading
 * hydrates covered positions into ordinary cells (so rendering, typing, paste,
 * fill, Clear Formats, sort and undo behave exactly as they always have), and
 * saving compacts them again. The ranges a sheet was loaded with ride along
 * under a symbol key as a stability hint, so an edit that leaves a range's
 * blanks alone leaves its persisted range alone too.
 */

export const BLANK_CELL_FORMATS_KEY = "blankCellFormats" as const;
/** Distinct blank-cell formats per sheet (Excel allows 65,490 per workbook). */
export const MAX_BLANK_CELL_FORMAT_GROUPS = 4_096;
/** Ranges across all groups of one sheet. */
export const MAX_BLANK_CELL_FORMAT_RANGES = 20_000;
/** Positions the ranges of one sheet may cover. Mirrors the importer's
 * per-sheet cell ceiling: covered cells are materialized when a sheet opens. */
export const MAX_BLANK_CELL_FORMAT_AREA = 200_000;
/** Positions the ranges of a whole workbook may cover. */
export const MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA = 1_000_000;
/** Formatted blanks a sheet must hold before they are stored as ranges.
 * Ranges exist for a formatted BLOCK; a sheet with a handful of formatted
 * blanks — a cell someone bolded and left empty — keeps them as cells,
 * byte-for-byte as before. A sheet already storing ranges keeps storing them. */
export const MIN_BLANK_CELL_FORMAT_CELLS = 16;
/** "XFD1048576:XFD1048576" is the longest canonical reference. */
export const MAX_BLANK_CELL_FORMAT_REF_LENGTH = 21;

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
const REF_PATTERN = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/;
const CELL_KEY_PATTERN = /^(0|[1-9][0-9]{0,6}):(0|[1-9][0-9]{0,4})$/;
const TEMPLATE_CELL_KEYS = new Set(["value", "kind", "style", "locked", "display"]);
const GROUP_KEYS = new Set(["style", "locked", "ranges"]);

type JsonRecord = Record<string, unknown>;

export interface BlankCellRect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface BlankCellFormatGroup {
  style?: JsonRecord;
  locked?: false;
  ranges: string[];
}

export interface BlankCell {
  value: "";
  kind: "text";
  style?: JsonRecord;
  locked?: false;
}

/** The loaded ranges of an in-memory sheet. A symbol key never reaches
 * persisted JSON (JSON.stringify, structuredClone and Object.keys skip it),
 * while object spread and history snapshots carry it along. */
export const BLANK_CELL_FORMATS_HINT: unique symbol = Symbol(
  "mrai.spreadsheet.blankCellFormats",
);

export interface BlankCellFormatsHintCarrier {
  [BLANK_CELL_FORMATS_HINT]?: readonly BlankCellFormatGroup[];
}

function isRecord(value: unknown): value is JsonRecord {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function columnIndex(letters: string): number {
  let index = 0;
  for (let i = 0; i < letters.length; i += 1) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}

function columnLetters(index: number): string {
  let letters = "";
  let rest = index + 1;
  while (rest > 0) {
    const remainder = (rest - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    rest = Math.floor((rest - 1) / 26);
  }
  return letters;
}

/** Zero-based inclusive bounds of a canonical reference, else null. */
export function parseBlankCellRangeRef(ref: unknown): BlankCellRect | null {
  if (typeof ref !== "string" || ref.length > MAX_BLANK_CELL_FORMAT_REF_LENGTH) {
    return null;
  }
  const match = REF_PATTERN.exec(ref);
  if (!match) return null;
  const c1 = columnIndex(match[1]);
  const r1 = Number(match[2]) - 1;
  let c2 = c1;
  let r2 = r1;
  if (match[3] !== undefined) {
    c2 = columnIndex(match[3]);
    r2 = Number(match[4]) - 1;
    if (r1 === r2 && c1 === c2) return null;
  }
  if (r1 > r2 || c1 > c2 || r2 >= MAX_ROWS || c2 >= MAX_COLUMNS) return null;
  return { r1, c1, r2, c2 };
}

export function formatBlankCellRangeRef({ r1, c1, r2, c2 }: BlankCellRect): string {
  const start = `${columnLetters(c1)}${r1 + 1}`;
  return r1 === r2 && c1 === c2 ? start : `${start}:${columnLetters(c2)}${r2 + 1}`;
}

/**
 * The first structural violation of a persisted field, for the page schema.
 * Styles are validated by the schema's cell style parser; this checks what a
 * per-group parser cannot: canonical references, the grid, and the budgets.
 */
export function blankCellFormatsIssue(
  value: unknown,
  rowCount: number,
  columnCount: number,
): { path: Array<string | number>; message: string } | null {
  if (value === undefined) return null;
  if (!Array.isArray(value)) return { path: [], message: "must be a list" };
  if (value.length > MAX_BLANK_CELL_FORMAT_GROUPS) {
    return {
      path: [],
      message: `must contain at most ${MAX_BLANK_CELL_FORMAT_GROUPS} groups`,
    };
  }
  let ranges = 0;
  let area = 0;
  for (const [groupIndex, group] of value.entries()) {
    if (!isRecord(group)) return { path: [groupIndex], message: "must be an object" };
    const unknown = Object.keys(group).filter((key) => !GROUP_KEYS.has(key));
    if (unknown.length > 0) {
      return { path: [groupIndex], message: "contains unsupported fields" };
    }
    if (
      "style" in group &&
      (!isRecord(group.style) || Object.keys(group.style).length === 0)
    ) {
      return { path: [groupIndex, "style"], message: "must be a non-empty object" };
    }
    if ("locked" in group && group.locked !== false) {
      return { path: [groupIndex, "locked"], message: "may only be false" };
    }
    if (!("style" in group) && !("locked" in group)) {
      return { path: [groupIndex], message: "must carry a style or locked: false" };
    }
    if (!Array.isArray(group.ranges) || group.ranges.length === 0) {
      return { path: [groupIndex, "ranges"], message: "must be a non-empty list" };
    }
    for (const [rangeIndex, ref] of group.ranges.entries()) {
      ranges += 1;
      if (ranges > MAX_BLANK_CELL_FORMAT_RANGES) {
        return {
          path: [groupIndex, "ranges"],
          message: `a sheet may hold at most ${MAX_BLANK_CELL_FORMAT_RANGES} ranges`,
        };
      }
      const rect = parseBlankCellRangeRef(ref);
      if (!rect) {
        return {
          path: [groupIndex, "ranges", rangeIndex],
          message: "must be a canonical A1 range",
        };
      }
      if (rect.r2 >= rowCount || rect.c2 >= columnCount) {
        return {
          path: [groupIndex, "ranges", rangeIndex],
          message: "is outside the declared sheet bounds",
        };
      }
      area += (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
      if (area > MAX_BLANK_CELL_FORMAT_AREA) {
        return {
          path: [groupIndex, "ranges", rangeIndex],
          message: `ranges may cover at most ${MAX_BLANK_CELL_FORMAT_AREA} cells`,
        };
      }
    }
  }
  return null;
}

/** Positions the canonical ranges of one sheet's field cover. */
export function blankCellFormatsArea(value: unknown): number {
  let area = 0;
  for (const group of Array.isArray(value) ? value : []) {
    const ranges = isRecord(group) && Array.isArray(group.ranges) ? group.ranges : [];
    for (const ref of ranges) {
      const rect = parseBlankCellRangeRef(ref);
      if (rect) area += (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
    }
  }
  return area;
}

function cloneJson<T>(value: T): T {
  if (Array.isArray(value)) return value.map(cloneJson) as T;
  if (!isRecord(value)) return value;
  const copy: JsonRecord = {};
  for (const key of Object.keys(value)) copy[key] = cloneJson(value[key]);
  return copy as T;
}

function groupTemplate(group: JsonRecord): { style?: JsonRecord; locked?: false } | null {
  const template: { style?: JsonRecord; locked?: false } = {};
  if (group.style !== undefined) {
    if (!isRecord(group.style) || Object.keys(group.style).length === 0) return null;
    template.style = group.style;
  }
  if (group.locked === false) template.locked = false;
  return template.style || template.locked === false ? template : null;
}

/** A fresh explicit cell for a group; the style is deep-copied. */
export function blankCellFromGroup(group: BlankCellFormatGroup): BlankCell {
  const cell: BlankCell = { value: "", kind: "text" };
  if (group.style) cell.style = cloneJson(group.style);
  if (group.locked === false) cell.locked = false;
  return cell;
}

function sheetDimension(value: unknown, ceiling: number): number {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.max(0, Math.min(Math.trunc(value), ceiling))
    : 0;
}

/**
 * Usable ranges in precedence order: malformed entries skipped, ranges clipped
 * to the grid, and the walk stopped once the area budget is spent, so a
 * hand-edited page can neither address past its grid nor force an unbounded
 * allocation.
 */
export function* blankCellFormatRanges(
  groups: unknown,
  rowCount: number,
  columnCount: number,
): Generator<{ rect: BlankCellRect; group: BlankCellFormatGroup; index: number }> {
  if (!Array.isArray(groups)) return;
  const rows = sheetDimension(rowCount, MAX_ROWS);
  const columns = sheetDimension(columnCount, MAX_COLUMNS);
  let areaLeft = MAX_BLANK_CELL_FORMAT_AREA;
  let rangesLeft = MAX_BLANK_CELL_FORMAT_RANGES;
  for (const [index, raw] of groups.slice(0, MAX_BLANK_CELL_FORMAT_GROUPS).entries()) {
    if (!isRecord(raw) || !Array.isArray(raw.ranges)) continue;
    const template = groupTemplate(raw);
    if (!template) continue;
    const group: BlankCellFormatGroup = { ...template, ranges: raw.ranges as string[] };
    for (const ref of raw.ranges) {
      if (rangesLeft <= 0) return;
      rangesLeft -= 1;
      const parsed = parseBlankCellRangeRef(ref);
      if (!parsed) continue;
      const rect = {
        ...parsed,
        r2: Math.min(parsed.r2, rows - 1),
        c2: Math.min(parsed.c2, columns - 1),
      };
      if (rect.r1 > rect.r2 || rect.c1 > rect.c2) continue;
      const area = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
      if (area > areaLeft) return;
      areaLeft -= area;
      yield { rect, group, index };
    }
  }
}

/** The groups a sheet can actually use (clipped, budgeted), for the hint. */
export function readBlankCellFormats(
  groups: unknown,
  rowCount: number,
  columnCount: number,
): BlankCellFormatGroup[] | undefined {
  const usable: BlankCellFormatGroup[] = [];
  let lastIndex = -1;
  for (const { rect, group, index } of blankCellFormatRanges(groups, rowCount, columnCount)) {
    if (index !== lastIndex) {
      lastIndex = index;
      usable.push({ ...group, ranges: [] });
    }
    usable[usable.length - 1].ranges.push(formatBlankCellRangeRef(rect));
  }
  return usable.length > 0 ? usable : undefined;
}

/**
 * Materialize every covered position into `cells` (mutated). A position that
 * `explicit` owns -- even one the caller normalized away -- is left alone:
 * the persisted entry decided that cell. Returns the usable groups.
 */
export function hydrateBlankCellFormats<C>(
  cells: Record<string, C | BlankCell>,
  explicit: Record<string, unknown>,
  groups: unknown,
  rowCount: number,
  columnCount: number,
): BlankCellFormatGroup[] | undefined {
  const usable = readBlankCellFormats(groups, rowCount, columnCount);
  if (!usable) return undefined;
  for (const group of usable) {
    for (const ref of group.ranges) {
      const rect = parseBlankCellRangeRef(ref)!;
      for (let row = rect.r1; row <= rect.r2; row += 1) {
        for (let col = rect.c1; col <= rect.c2; col += 1) {
          const key = `${row}:${col}`;
          if (Object.hasOwn(explicit, key) || Object.hasOwn(cells, key)) continue;
          cells[key] = blankCellFromGroup(group);
        }
      }
    }
  }
  return usable;
}

type RawSheet = {
  cells?: unknown;
  rowCount?: unknown;
  columnCount?: unknown;
  [BLANK_CELL_FORMATS_KEY]?: unknown;
};

/**
 * Point reads of a raw (not hydrated) sheet: the explicit entry at a
 * position, else a fresh copy of the first covering range's blank, else
 * undefined. Parse once, read many.
 */
export function spreadsheetCellReader(
  sheet: RawSheet,
): (row: number, col: number) => unknown {
  const cells = isRecord(sheet.cells) ? sheet.cells : {};
  const ranges = [
    ...blankCellFormatRanges(
      sheet[BLANK_CELL_FORMATS_KEY],
      Number(sheet.rowCount),
      Number(sheet.columnCount),
    ),
  ];
  return (row, col) => {
    const key = `${row}:${col}`;
    if (Object.hasOwn(cells, key)) return cells[key];
    for (const { rect, group } of ranges) {
      if (row >= rect.r1 && row <= rect.r2 && col >= rect.c1 && col <= rect.c2) {
        return blankCellFromGroup(group);
      }
    }
    return undefined;
  };
}

/** One point read of a raw sheet (see spreadsheetCellReader). */
export function spreadsheetCellAt(sheet: RawSheet, row: number, col: number): unknown {
  return spreadsheetCellReader(sheet)(row, col);
}

interface TemplateMemo {
  style: unknown;
  value: unknown;
  kind: unknown;
  locked: unknown;
  display: unknown;
  size: number;
  key: string | null;
}

const templateMemo = new WeakMap<object, TemplateMemo>();







function blankTemplateKey(cell: unknown): string | null {
  if (!isRecord(cell) || isMaterializedLineStyleBlank(cell)) return null;
  const keys = Object.keys(cell);
  const memo = templateMemo.get(cell);
  if (
    memo &&
    memo.style === cell.style &&
    memo.value === cell.value &&
    memo.kind === cell.kind &&
    memo.locked === cell.locked &&
    memo.display === cell.display &&
    memo.size === keys.length
  ) {
    return memo.key;
  }
  let key: string | null = null;
  if (
    keys.every((name) => TEMPLATE_CELL_KEYS.has(name) || cell[name] === undefined) &&
    cell.value === "" &&
    cell.kind === "text" &&
    (cell.display === undefined || cell.display === "") &&
    (cell.locked === undefined || cell.locked === false) &&
    (cell.style === undefined ||
      (isRecord(cell.style) && Object.keys(cell.style).length > 0)) &&
    (cell.style !== undefined || cell.locked === false)
  ) {
    try {
      key = JSON.stringify([cell.style ?? null, cell.locked === false]);
    } catch {
      key = null;
    }
  }
  templateMemo.set(cell, {
    style: cell.style,
    value: cell.value,
    kind: cell.kind,
    locked: cell.locked,
    display: cell.display,
    size: keys.length,
    key,
  });
  return key;
}

function groupKey(group: BlankCellFormatGroup): string {
  return JSON.stringify([group.style ?? null, group.locked === false]);
}

/** Disjoint rectangles covering exactly the sorted, unique positions. Each row
 * is cut into runs; a run extends the rectangle above it when both span the
 * same columns on consecutive rows. Mirrors the server's `_rectangles`. */
function rectangles(positions: Array<[number, number]>): BlankCellRect[] {
  const out: BlankCellRect[] = [];
  let open = new Map<string, { r1: number; r2: number; c1: number; c2: number }>();
  let index = 0;
  while (index < positions.length) {
    const row = positions[index][0];
    const runs: Array<[number, number]> = [];
    let start = positions[index][1];
    let previous = start;
    index += 1;
    while (index < positions.length && positions[index][0] === row) {
      const col = positions[index][1];
      if (col !== previous + 1) {
        runs.push([start, previous]);
        start = col;
      }
      previous = col;
      index += 1;
    }
    runs.push([start, previous]);
    const next = new Map<string, { r1: number; r2: number; c1: number; c2: number }>();
    for (const [c1, c2] of runs) {
      const span = `${c1}:${c2}`;
      let rect = open.get(span);
      open.delete(span);
      if (rect && rect.r2 === row - 1) {
        rect.r2 = row;
      } else {
        if (rect) out.push({ ...rect });
        rect = { r1: row, r2: row, c1, c2 };
      }
      next.set(span, rect);
    }
    for (const rect of open.values()) out.push({ ...rect });
    open = next;
  }
  for (const rect of open.values()) out.push({ ...rect });
  return out.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
}

export interface BlankCellFormatsPlan {
  /** Explicit cells that the ranges now stand for. */
  removable: ReadonlySet<string>;
  groups: BlankCellFormatGroup[] | undefined;
}

const planCache = new WeakMap<
  object,
  {
    hint: readonly BlankCellFormatGroup[] | undefined;
    rowCount: number;
    columnCount: number;
    areaBudget: number;
    plan: BlankCellFormatsPlan;
  }
>();

/**
 * Decide which explicit blanks of an in-memory (hydrated) sheet are persisted
 * as ranges. A hint range survives while every position it covers still holds
 * an entry, none of them is a blank of another format, and at least one is
 * still its own blank: then typing into its cells, or clearing their contents,
 * never rewrites the range. Everything else collapses into fresh rectangles.
 * A position with no entry is never covered: in memory, no entry means no
 * format, so a cleared format can never come back from a range. A sheet with
 * no ranges yet and fewer than MIN_BLANK_CELL_FORMAT_CELLS formatted blanks
 * keeps every one of them as a cell.
 */
export function planBlankCellFormats(
  cells: Record<string, unknown>,
  hint: readonly BlankCellFormatGroup[] | undefined,
  rowCount: number,
  columnCount: number,
  areaBudget = MAX_BLANK_CELL_FORMAT_AREA,
): BlankCellFormatsPlan {
  const cached = planCache.get(cells);
  if (
    cached &&
    cached.hint === hint &&
    cached.rowCount === rowCount &&
    cached.columnCount === columnCount &&
    cached.areaBudget === areaBudget
  ) {
    return cached.plan;
  }
  // An edit that touched no formatted blank, on positions that hold an entry
  // before and after and that no range stands for, leaves every input of the plan
  // as it was: the walk below over each cell was about 25 ms per entry on 30,000.
  const near = nearestKnownCellMap(cells, (map) => {
    const known = planCache.get(map);
    return known && known.hint === hint && known.rowCount === rowCount && known.columnCount === columnCount &&
      known.areaBudget === areaBudget ? { map, plan: known.plan } : undefined;
  });
  if (near && near.keys.every((key) =>
    Object.hasOwn(cells, key) && Object.hasOwn(near.value.map, key) && !near.value.plan.removable.has(key) &&
    blankTemplateKey(cells[key]) === null && blankTemplateKey(near.value.map[key]) === null,
  )) {
    planCache.set(cells, { hint, rowCount, columnCount, areaBudget, plan: near.value.plan });
    return near.value.plan;
  }
  const rows = sheetDimension(rowCount, MAX_ROWS);
  const columns = sheetDimension(columnCount, MAX_COLUMNS);
  const blanks = new Map<string, string>();
  const templates = new Map<string, BlankCellFormatGroup>();
  for (const key of Object.keys(cells)) {
    const template = blankTemplateKey(cells[key]);
    if (template === null) continue;
    const match = CELL_KEY_PATTERN.exec(key);
    if (!match || Number(match[1]) >= rows || Number(match[2]) >= columns) continue;
    blanks.set(key, template);
    if (!templates.has(template)) {
      if (templates.size >= MAX_BLANK_CELL_FORMAT_GROUPS) {
        blanks.delete(key);
        continue;
      }
      const cell = cells[key] as JsonRecord;
      templates.set(template, {
        ...(cell.style !== undefined ? { style: cell.style as JsonRecord } : {}),
        ...(cell.locked === false ? { locked: false as const } : {}),
        ranges: [],
      });
    }
  }

  const removable = new Set<string>();
  const covered = new Set<string>();
  const output = new Map<string, BlankCellFormatGroup>();
  if (!hint && blanks.size < MIN_BLANK_CELL_FORMAT_CELLS) {
    // Too few to be a formatted block: the sheet keeps them as cells.
    const plan: BlankCellFormatsPlan = { removable, groups: undefined };
    planCache.set(cells, { hint, rowCount, columnCount, areaBudget, plan });
    return plan;
  }
  let rangesLeft = MAX_BLANK_CELL_FORMAT_RANGES;
  let areaLeft = Math.max(0, Math.min(areaBudget, MAX_BLANK_CELL_FORMAT_AREA));
  const emit = (template: string, rect: BlankCellRect) => {
    let group = output.get(template);
    if (!group) {
      const source = templates.get(template)!;
      group = {
        ...(source.style !== undefined ? { style: cloneJson(source.style) } : {}),
        ...(source.locked === false ? { locked: false as const } : {}),
        ranges: [],
      };
      output.set(template, group);
    }
    group.ranges.push(formatBlankCellRangeRef(rect));
  };

  for (const group of hint ?? []) {
    const template = groupKey(group);
    if (!templates.has(template)) continue;
    for (const ref of group.ranges) {
      const rect = parseBlankCellRangeRef(ref);
      if (!rect || rect.r2 >= rows || rect.c2 >= columns) continue;
      const area = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
      if (rangesLeft <= 0 || area > areaLeft) continue;
      let intact = true;
      let own = 0;
      for (let row = rect.r1; intact && row <= rect.r2; row += 1) {
        for (let col = rect.c1; col <= rect.c2; col += 1) {
          const key = `${row}:${col}`;
          if (!Object.hasOwn(cells, key) || covered.has(key)) {
            intact = false;
            break;
          }
          const blank = blanks.get(key);
          if (blank === undefined) continue;
          if (blank !== template) {
            intact = false;
            break;
          }
          own += 1;
        }
      }
      if (!intact || own === 0) continue;
      rangesLeft -= 1;
      areaLeft -= area;
      for (let row = rect.r1; row <= rect.r2; row += 1) {
        for (let col = rect.c1; col <= rect.c2; col += 1) {
          const key = `${row}:${col}`;
          covered.add(key);
          if (blanks.has(key)) removable.add(key);
        }
      }
      emit(template, rect);
    }
  }

  const loose = new Map<string, Array<[number, number]>>();
  for (const [key, template] of blanks) {
    if (covered.has(key)) continue;
    const separator = key.indexOf(":");
    const position: [number, number] = [
      Number(key.slice(0, separator)),
      Number(key.slice(separator + 1)),
    ];
    const list = loose.get(template);
    if (list) list.push(position);
    else loose.set(template, [position]);
  }
  const ordered = [...loose.entries()].map(([template, positions]) => {
    positions.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    return { template, positions };
  });
  ordered.sort(
    (a, b) =>
      a.positions[0][0] - b.positions[0][0] ||
      a.positions[0][1] - b.positions[0][1],
  );
  for (const { template, positions } of ordered) {
    for (const rect of rectangles(positions)) {
      const area = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
      if (rangesLeft <= 0 || area > areaLeft) continue;
      rangesLeft -= 1;
      areaLeft -= area;
      for (let row = rect.r1; row <= rect.r2; row += 1) {
        for (let col = rect.c1; col <= rect.c2; col += 1) {
          removable.add(`${row}:${col}`);
        }
      }
      emit(template, rect);
    }
  }

  const plan: BlankCellFormatsPlan = {
    removable,
    groups: output.size > 0 ? [...output.values()] : undefined,
  };
  planCache.set(cells, { hint, rowCount, columnCount, areaBudget, plan });
  return plan;
}

/**
 * Persisted form of one in-memory sheet. The plan comes from the live cells
 * (identity-stable across saves); `persisted` is the same sheet after the
 * serializer's display cache. A planned blank whose cache now says something
 * (a spilled array element) keeps its entry: explicit cells always win.
 */
export function compactSheetBlankCellFormats<
  S extends { cells: Record<string, unknown>; rowCount: number; columnCount: number },
>(live: S, persisted: S, areaBudget = MAX_BLANK_CELL_FORMAT_AREA): S {
  const hint = (live as BlankCellFormatsHintCarrier)[BLANK_CELL_FORMATS_HINT];
  const plan = planBlankCellFormats(
    live.cells,
    hint,
    live.rowCount,
    live.columnCount,
    areaBudget,
  );
  const hasField = Object.hasOwn(persisted, BLANK_CELL_FORMATS_KEY);
  if (plan.removable.size === 0 && !plan.groups && !hasField) return persisted;
  const cells: Record<string, unknown> = {};
  for (const key of Object.keys(persisted.cells)) {
    const cell = persisted.cells[key];
    if (
      plan.removable.has(key) &&
      (!isRecord(cell) || cell.display === undefined || cell.display === "")
    ) {
      continue;
    }
    cells[key] = cell;
  }
  const next = { ...persisted, cells } as S;
  const writable = next as unknown as JsonRecord & BlankCellFormatsHintCarrier;
  if (plan.groups) writable[BLANK_CELL_FORMATS_KEY] = plan.groups;
  else delete writable[BLANK_CELL_FORMATS_KEY];
  delete writable[BLANK_CELL_FORMATS_HINT];
  return next;
}

/** Persisted form of a workbook's sheets, sharing the workbook area budget
 * in sheet order (see compactSheetBlankCellFormats). */
export function compactWorkbookBlankCellFormats<
  S extends { cells: Record<string, unknown>; rowCount: number; columnCount: number },
>(live: readonly S[], persisted: readonly S[]): S[] {
  let areaLeft = MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA;
  return persisted.map((sheet, index) => {
    const compacted = compactSheetBlankCellFormats(live[index] ?? sheet, sheet, areaLeft);
    areaLeft -= blankCellFormatsArea((compacted as JsonRecord)[BLANK_CELL_FORMATS_KEY]);
    return compacted;
  });
}

/**
 * Let a batch of cell edits run against a persisted-shape sheet (mutated in
 * place; `sheet.cells` must already be the caller's own copy). Covered blanks
 * become ordinary cells so an edit keeps their format and a sort moves it with
 * its row; the returned function folds whatever is still blank back into
 * ranges, keeping the ranges the edits did not touch.
 */
export function hydrateSheetForEditing(sheet: {
  cells: Record<string, unknown>;
  rowCount: number;
  columnCount: number;
  [key: string]: unknown;
}): () => void {
  const groups = sheet[BLANK_CELL_FORMATS_KEY];
  if (!Array.isArray(groups) || groups.length === 0) return () => {};
  const hint = hydrateBlankCellFormats(
    sheet.cells,
    sheet.cells,
    groups,
    sheet.rowCount,
    sheet.columnCount,
  );
  delete sheet[BLANK_CELL_FORMATS_KEY];
  return () => {
    const plan = planBlankCellFormats(sheet.cells, hint, sheet.rowCount, sheet.columnCount);
    if (plan.removable.size > 0) {
      const kept: Record<string, unknown> = {};
      for (const key of Object.keys(sheet.cells)) {
        if (!plan.removable.has(key)) kept[key] = sheet.cells[key];
      }
      sheet.cells = kept;
    }
    if (plan.groups) sheet[BLANK_CELL_FORMATS_KEY] = plan.groups;
  };
}

/**
 * A persisted page with every sheet's ranges hydrated into ordinary cells
 * (copies), each sheet carrying the ranges it had as the save hint. Pair with
 * compactHydratedContentBlankCellFormats to edit a saved page the way the
 * editor does and hand back the same persisted shape.
 */
export function hydrateContentBlankCellFormats<T>(content: T): T {
  if (!isRecord(content) || !Array.isArray(content.sheets)) return content;
  let changed = false;
  const sheets = content.sheets.map((sheet) => {
    if (!isRecord(sheet) || !isRecord(sheet.cells) || !Array.isArray(sheet[BLANK_CELL_FORMATS_KEY])) {
      return sheet;
    }
    changed = true;
    const cells: Record<string, unknown> = { ...sheet.cells };
    const hint = hydrateBlankCellFormats(
      cells,
      sheet.cells,
      sheet[BLANK_CELL_FORMATS_KEY],
      Number(sheet.rowCount),
      Number(sheet.columnCount),
    );
    const next: JsonRecord & BlankCellFormatsHintCarrier = { ...sheet, cells };
    delete next[BLANK_CELL_FORMATS_KEY];
    if (hint) next[BLANK_CELL_FORMATS_HINT] = hint;
    return next;
  });
  return changed ? ({ ...content, sheets } as T) : content;
}

/**
 * Persisted form of a hydrated page whose sheets may have lost their hints
 * (a JSON merge, say): the hints are taken by sheet id from `hinted`, the
 * hydrateContentBlankCellFormats result the page was derived from, so ranges
 * nobody touched come back exactly as they were saved.
 */
export function compactHydratedContentBlankCellFormats<T>(content: T, hinted: unknown): T {
  if (!isRecord(content) || !Array.isArray(content.sheets)) return content;
  const hints = new Map<unknown, readonly BlankCellFormatGroup[]>();
  for (const sheet of isRecord(hinted) && Array.isArray(hinted.sheets) ? hinted.sheets : []) {
    const hint = isRecord(sheet)
      ? (sheet as BlankCellFormatsHintCarrier)[BLANK_CELL_FORMATS_HINT]
      : undefined;
    if (hint) hints.set((sheet as JsonRecord).id, hint);
  }
  let changed = false;
  let areaLeft = MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA;
  const sheets = content.sheets.map((sheet) => {
    // Only sheets that were saved with ranges go back to ranges: a page that
    // still stores every blank as a cell keeps exactly that shape.
    const hint = isRecord(sheet) ? hints.get(sheet.id) : undefined;
    if (
      !hint ||
      !isRecord(sheet) ||
      !isRecord(sheet.cells) ||
      Object.hasOwn(sheet, BLANK_CELL_FORMATS_KEY)
    ) {
      return sheet;
    }
    const plan = planBlankCellFormats(
      sheet.cells,
      hint,
      Number(sheet.rowCount),
      Number(sheet.columnCount),
      areaLeft,
    );
    areaLeft -= blankCellFormatsArea(plan.groups);
    if (plan.removable.size === 0 && !plan.groups) return sheet;
    changed = true;
    const next: JsonRecord = { ...sheet, cells: {} };
    const kept = next.cells as JsonRecord;
    for (const key of Object.keys(sheet.cells)) {
      if (!plan.removable.has(key)) kept[key] = sheet.cells[key];
    }
    if (plan.groups) next[BLANK_CELL_FORMATS_KEY] = plan.groups;
    return next;
  });
  return changed ? ({ ...content, sheets } as T) : content;
}

/**
 * Collapse the formatted blanks of a parsed page (importers). Existing ranges
 * are folded in first, so the result depends only on the effective cells.
 */
export function compactContentBlankCellFormats<T>(content: T): T {
  if (!isRecord(content) || !Array.isArray(content.sheets)) return content;
  let changed = false;
  let areaLeft = MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA;
  const sheets = content.sheets.map((sheet) => {
    if (!isRecord(sheet) || !isRecord(sheet.cells)) return sheet;
    const rowCount = Number(sheet.rowCount);
    const columnCount = Number(sheet.columnCount);
    const cells: Record<string, unknown> = { ...sheet.cells };
    hydrateBlankCellFormats(
      cells,
      sheet.cells,
      sheet[BLANK_CELL_FORMATS_KEY],
      rowCount,
      columnCount,
    );
    const plan = planBlankCellFormats(cells, undefined, rowCount, columnCount, areaLeft);
    areaLeft -= blankCellFormatsArea(plan.groups);
    if (
      plan.removable.size === 0 &&
      !plan.groups &&
      !Object.hasOwn(sheet, BLANK_CELL_FORMATS_KEY)
    ) {
      return sheet;
    }
    changed = true;
    const next: JsonRecord = { ...sheet, cells: {} };
    const kept = next.cells as JsonRecord;
    for (const key of Object.keys(cells)) {
      if (!plan.removable.has(key)) kept[key] = cells[key];
    }
    if (plan.groups) next[BLANK_CELL_FORMATS_KEY] = plan.groups;
    else delete next[BLANK_CELL_FORMATS_KEY];
    return next;
  });
  return changed ? ({ ...content, sheets } as T) : content;
}

/**
 * Keep a sheet's hint in step with a structural insert/delete of one line, so
 * the ranges an edit did not touch persist unchanged. An inserted line is
 * blank in this editor, so a range it crosses splits around it; a deleted line
 * shrinks the ranges it crossed and drops the ones it held alone.
 */
export function shiftBlankCellFormatsHint(
  hint: readonly BlankCellFormatGroup[] | undefined,
  axis: "row" | "col",
  at: number,
  delta: 1 | -1,
): BlankCellFormatGroup[] | undefined {
  if (!hint) return hint;
  const limit = axis === "row" ? MAX_ROWS : MAX_COLUMNS;
  const next: BlankCellFormatGroup[] = [];
  for (const group of hint) {
    const ranges: string[] = [];
    for (const ref of group.ranges) {
      const rect = parseBlankCellRangeRef(ref);
      if (!rect) continue;
      const low = axis === "row" ? rect.r1 : rect.c1;
      const high = axis === "row" ? rect.r2 : rect.c2;
      const spans: Array<[number, number]> = [];
      if (delta > 0) {
        if (at <= low) spans.push([low + 1, high + 1]);
        else if (at > high) spans.push([low, high]);
        else spans.push([low, at - 1], [at + 1, high + 1]);
      } else if (at < low) {
        spans.push([low - 1, high - 1]);
      } else if (at > high) {
        spans.push([low, high]);
      } else if (high > low) {
        spans.push([low, high - 1]);
      }
      for (const [start, end] of spans) {
        if (start > end || end >= limit) continue;
        ranges.push(
          formatBlankCellRangeRef(
            axis === "row"
              ? { ...rect, r1: start, r2: end }
              : { ...rect, c1: start, c2: end },
          ),
        );
      }
    }
    if (ranges.length > 0) next.push({ ...group, ranges });
  }
  return next.length > 0 ? next : undefined;
}
