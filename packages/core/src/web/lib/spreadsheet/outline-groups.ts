/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure model + ops for spreadsheet row/column OUTLINE GROUPING (Excel / WPS
 * parity). A group is a contiguous run of lines (rows or columns) on one axis
 * that can be COLLAPSED to hide its body and show a single [+] toggle, or
 * EXPANDED to show its body and a [-] toggle. Groups nest by CONTAINMENT — an
 * inner group's range must sit fully inside its parent's — to form an outline
 * up to 8 levels deep (Excel's hard limit). Crossing ranges (partial overlap)
 * are illegal and rejected.
 *
 * Everything here is pure and data-only so it can be unit-tested in isolation
 * and reused by the React grid: the grid stores `outlineGroups` on the sheet
 * model, calls createGroup / removeGroup / toggleCollapse / setOutlineLevel to
 * mutate it, then derives the hidden-line sets and the gutter render model from
 * the selectors. Nothing here touches the DOM or the formula engine; the grid
 * wires hiddenRowsFromGroups into both the evaluator's filter/manual-hidden
 * seams and the rendered row band.
 *
 * Collapse semantics match Excel: collapsing a group hides EVERY line in its
 * range EXCEPT the group's own summary line. Excel puts the summary BELOW the
 * detail by default (the row under the grouped block), so the LAST line of a
 * row group / column group stays visible and shows the toggle; the lines above
 * it are the hidden detail. We follow that convention — the toggle sits on the
 * line just past the group's body, i.e. `end + 1` is the summary, and the
 * hidden body is `[start, end]`.
 *
 * Excel's Data > Outline > Settings can move the summary: clearing "Summary
 * rows below detail" / "Summary columns to right of detail" (OOXML
 * `sheetPr/outlinePr summaryBelow="0"` / `summaryRight="0"`) puts it on the
 * line BEFORE the body, `start - 1`, where the toggle is drawn and the file
 * keeps the collapsed flag. The body, the levels and the hidden lines are the
 * same either way; only the summary line moves (see outlineSummaryLine).
 */

/** Which axis a group lives on. Row groups hide rows; column groups hide cols. */
export type OutlineAxis = "row" | "col";







export const MAX_OUTLINE_DEPTH = 7;

/**
 * A single outline group: a contiguous, inclusive line range on one axis plus
 * its collapsed flag. `start <= end`; both are 0-indexed line numbers. Nesting
 * is implicit in the ranges (a group contained in another is one level deeper)
 * — there is no parent pointer, so the model stays a flat, order-independent
 * list that round-trips trivially.
 */
export interface OutlineGroup {
  axis: OutlineAxis;
  start: number;
  end: number;
  collapsed: boolean;
}

/**
 * The sheet's Outline Settings as the page stores them (Excel's Data > Outline
 * > Settings). Absent means Excel's default: summary rows below the detail and
 * summary columns to the right of it; only `false` is ever stored.
 */
export interface OutlineSummarySettings {
  outlineSummaryBelow?: boolean;
  outlineSummaryRight?: boolean;
}

/** Whether `axis`'s summary lines follow their groups (Excel's default). */
export function outlineSummaryAfter(
  settings: OutlineSummarySettings | null | undefined,
  axis: OutlineAxis,
): boolean {
  return (
    (axis === "row"
      ? settings?.outlineSummaryBelow
      : settings?.outlineSummaryRight) !== false
  );
}

/**
 * A group's SUMMARY line: `end + 1` when summaries follow the detail, else
 * `start - 1`. It carries the [+]/[-] toggle and, in the file, the collapsed
 * flag. Null for a group on line 0 whose summary would precede it: Excel has
 * no line to draw its toggle on (the level buttons still collapse it).
 */
export function outlineSummaryLine(
  group: { start: number; end: number },
  summaryAfter = true,
): number | null {
  const line = summaryAfter ? group.end + 1 : group.start - 1;
  return line >= 0 ? line : null;
}

/** A normalized, finite, ordered [start, end] range, or null when invalid. */
function normalizeRange(
  start: unknown,
  end: unknown,
): { start: number; end: number } | null {
  const a = Number(start);
  const b = Number(end);
  if (!Number.isInteger(a) || !Number.isInteger(b) || a < 0 || b < 0) {
    return null;
  }
  return { start: Math.min(a, b), end: Math.max(a, b) };
}

/** Two ranges on the same axis. Containment / disjoint / crossing classifier. */
function relation(
  a: { start: number; end: number },
  b: { start: number; end: number },
): "disjoint" | "equal" | "a-contains-b" | "b-contains-a" | "crossing" {
  if (a.start === b.start && a.end === b.end) return "equal";
  if (a.end < b.start || b.end < a.start) return "disjoint";
  if (a.start <= b.start && a.end >= b.end) return "a-contains-b";
  if (b.start <= a.start && b.end >= a.end) return "b-contains-a";
  return "crossing";
}

/**
 * The nesting depth (1-based) a NEW range would have given the existing groups
 * on its axis: 1 + the number of existing groups that would CONTAIN it. Returns
 * null when the range crosses any existing group (illegal) — the caller rejects
 * the create. Equal ranges count as containment for the cap check (a duplicate
 * is rejected separately).
 */
function depthOfRange(
  range: { start: number; end: number },
  existingSameAxis: readonly OutlineGroup[],
): number | null {
  let containers = 0;
  for (const g of existingSameAxis) {
    const rel = relation({ start: g.start, end: g.end }, range);
    if (rel === "crossing") return null;
    if (rel === "a-contains-b" || rel === "equal") containers += 1;
  }
  return containers + 1;
}

/** Sort groups so the OUTERMOST (widest span) comes first; stable on start. */
function bySpanDesc(a: OutlineGroup, b: OutlineGroup): number {
  const sa = a.end - a.start;
  const sb = b.end - b.start;
  if (sa !== sb) return sb - sa;
  return a.start - b.start;
}

/**
 * Create one outline group from each contiguous range on `axis`, validating
 * Excel's nesting rules against the existing groups. A range that:
 *   - is malformed (non-integer / negative)            -> skipped
 *   - crosses an existing group (partial overlap)       -> skipped
 *   - exactly duplicates an existing group on this axis -> skipped
 *   - would nest deeper than MAX_OUTLINE_DEPTH          -> skipped
 * is dropped; every other range is added as an expanded group. Returns a NEW
 * array (input untouched). Ranges are validated one at a time against the
 * accumulating set so two new ranges in the same call can legally nest.
 */
export function createGroup(
  groups: readonly OutlineGroup[],
  ranges: ReadonlyArray<{ start: number; end: number }>,
  axis: OutlineAxis,
): OutlineGroup[] {
  const next = [...groups];
  for (const raw of ranges) {
    const range = normalizeRange(raw.start, raw.end);
    if (!range) continue;
    const sameAxis = next.filter((g) => g.axis === axis);
    // Reject an exact duplicate on this axis (Excel won't stack two identical
    // groups; it would just toggle the existing one).
    if (sameAxis.some((g) => g.start === range.start && g.end === range.end)) {
      continue;
    }
    const depth = depthOfRange(range, sameAxis);
    if (depth === null || depth > MAX_OUTLINE_DEPTH) continue;
    next.push({ axis, start: range.start, end: range.end, collapsed: false });
  }
  return next;
}

/**
 * Remove the group on `axis` whose range exactly matches [start, end]. When no
 * group matches exactly, remove the INNERMOST (narrowest) group that contains
 * the range — Excel's Ungroup peels one level off the selection. Children of a
 * removed group survive (they just shift one level outward). Returns a NEW array.
 */
export function removeGroup(
  groups: readonly OutlineGroup[],
  range: { start: number; end: number },
  axis: OutlineAxis,
): OutlineGroup[] {
  const norm = normalizeRange(range.start, range.end);
  if (!norm) return [...groups];
  const sameAxis = groups.filter((g) => g.axis === axis);
  const exact = sameAxis.find(
    (g) => g.start === norm.start && g.end === norm.end,
  );
  let target: OutlineGroup | undefined = exact;
  if (!target) {
    // Innermost (narrowest span) container of the range.
    const containers = sameAxis
      .filter((g) => g.start <= norm.start && g.end >= norm.end)
      .sort((a, b) => a.end - a.start - (b.end - b.start));
    target = containers[0];
  }
  if (!target) return [...groups];
  return groups.filter((g) => g !== target);
}

/**
 * Excel keeps an outline as a LEVEL per line (`row@outlineLevel`,
 * `col@outlineLevel`), and every group is a maximal run of lines at or above
 * one level. Data > Group raises the selected lines one level and Ungroup
 * lowers them one, which is why in Excel two groups made side by side read
 * back as ONE group, a range that crosses an existing group is legal, a
 * single row groups on its own, and ungrouping one line of a group splits it.
 * `groupLines` / `ungroupLines` apply those per-line edits to the range model
 * and rebuild the groups exactly as the XLSX reader (and the API) derive them
 * from levels, so what the grid shows is what the saved file reads back as.
 *
 * The arithmetic runs on segments between group boundaries, never per line,
 * so a whole-column selection costs no more than a few cells.
 */
type LevelSegment = { start: number; end: number; level: number };

function levelSegments(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  extra: { start: number; end: number } | null,
): LevelSegment[] {
  const sameAxis = groups.filter((g) => g.axis === axis);
  const points = new Set<number>();
  for (const g of sameAxis) {
    points.add(g.start);
    points.add(g.end + 1);
  }
  if (extra) {
    points.add(extra.start);
    points.add(extra.end + 1);
  }
  const sorted = [...points].sort((a, b) => a - b);
  const segments: LevelSegment[] = [];
  for (let index = 0; index + 1 < sorted.length; index += 1) {
    const start = sorted[index];
    const end = sorted[index + 1] - 1;
    let level = 0;
    for (const g of sameAxis) {
      if (g.start <= start && g.end >= end) level += 1;
    }
    segments.push({ start, end, level });
  }
  return segments;
}

function groupsFromSegments(
  segments: readonly LevelSegment[],
  axis: OutlineAxis,
  previous: readonly OutlineGroup[],
): OutlineGroup[] {
  const groups: OutlineGroup[] = [];
  const maxLevel = segments.reduce((max, s) => Math.max(max, s.level), 0);
  for (let depth = 1; depth <= maxLevel; depth += 1) {
    let runStart = -1;
    let runEnd = -1;
    const flush = () => {
      if (runStart < 0) return;
      const kept = previous.find(
        (g) => g.axis === axis && g.start === runStart && g.end === runEnd,
      );
      groups.push({
        axis,
        start: runStart,
        end: runEnd,
        collapsed: kept?.collapsed ?? false,
      });
      runStart = -1;
    };
    for (const segment of segments) {
      const inside = segment.level >= depth;
      if (inside && runStart >= 0 && segment.start === runEnd + 1) {
        runEnd = segment.end;
      } else if (inside) {
        flush();
        runStart = segment.start;
        runEnd = segment.end;
      } else {
        flush();
      }
    }
    flush();
  }
  return groups;
}

function relevelLines(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  range: { start: number; end: number },
  delta: 1 | -1,
): OutlineGroup[] {
  const norm = normalizeRange(range.start, range.end);
  if (!norm) return [...groups];
  let changed = false;
  const segments = levelSegments(groups, axis, norm).map((segment) => {
    if (segment.start < norm.start || segment.end > norm.end) return segment;
    const level = Math.min(
      MAX_OUTLINE_DEPTH,
      Math.max(0, segment.level + delta),
    );
    if (level !== segment.level) changed = true;
    return { ...segment, level };
  });
  if (!changed) return [...groups];
  return [
    ...groups.filter((g) => g.axis !== axis),
    ...groupsFromSegments(segments, axis, groups),
  ];
}

/** Excel's Data > Group on a selection: its lines go one level deeper
 * (Excel nests seven deep; lines already at the deepest level stay there). */
export function groupLines(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  range: { start: number; end: number },
): OutlineGroup[] {
  return relevelLines(groups, axis, range, 1);
}

/** Excel's Data > Ungroup on a selection: its grouped lines come up one
 * level, splitting a group when only part of it is selected. */
export function ungroupLines(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  range: { start: number; end: number },
): OutlineGroup[] {
  return relevelLines(groups, axis, range, -1);
}

/** Whether two outline lists describe the same groups (order ignored). */
export function sameOutlineGroups(
  a: readonly OutlineGroup[],
  b: readonly OutlineGroup[],
): boolean {
  if (a.length !== b.length) return false;
  const key = (g: OutlineGroup) =>
    `${g.axis}:${g.start}:${g.end}:${g.collapsed ? 1 : 0}`;
  const left = a.map(key).sort();
  const right = b.map(key).sort();
  return left.every((value, index) => value === right[index]);
}

/**
 * Toggle (or set) the collapsed flag of the group on `axis` matching exactly
 * [start, end]. When `collapsed` is omitted the flag flips; otherwise it is set
 * to the given value. No matching group -> unchanged. Returns a NEW array.
 */
export function toggleCollapse(
  groups: readonly OutlineGroup[],
  range: { start: number; end: number },
  axis: OutlineAxis,
  collapsed?: boolean,
): OutlineGroup[] {
  const norm = normalizeRange(range.start, range.end);
  if (!norm) return [...groups];
  return groups.map((g) => {
    if (g.axis !== axis || g.start !== norm.start || g.end !== norm.end) {
      return g;
    }
    return { ...g, collapsed: collapsed ?? !g.collapsed };
  });
}

/**
 * The 1-based outline level of a group: 1 + the number of OTHER same-axis
 * groups that strictly or equally contain it. Level 1 = outermost.
 */
export function levelOf(
  group: OutlineGroup,
  groups: readonly OutlineGroup[],
): number {
  let containers = 0;
  for (const g of groups) {
    if (g === group || g.axis !== group.axis) continue;
    if (g.start <= group.start && g.end >= group.end) containers += 1;
  }
  return containers + 1;
}

/**
 * The deepest outline level present on `axis` (0 when there are no groups).
 * The corner "level buttons" run 1..(this + 1): button N collapses every group
 * at level >= N and expands the rest, so the highest button shows the full
 * detail.
 */
export function maxOutlineLevel(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
): number {
  let max = 0;
  for (const g of groups) {
    if (g.axis !== axis) continue;
    const lvl = levelOf(g, groups);
    if (lvl > max) max = lvl;
  }
  return max;
}

/**
 * Apply an outline LEVEL button (Excel's 1/2/3… corner buttons) to `axis`:
 * groups at level >= `depth` collapse, groups at level < `depth` expand. So
 * depth 1 collapses everything; depth (maxLevel + 1) expands everything. Groups
 * on the OTHER axis are untouched. Returns a NEW array.
 */
export function setOutlineLevel(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  depth: number,
): OutlineGroup[] {
  const d = Math.max(1, Math.trunc(Number(depth) || 1));
  return groups.map((g) => {
    if (g.axis !== axis) return g;
    const collapsed = levelOf(g, groups) >= d;
    return collapsed === g.collapsed ? g : { ...g, collapsed };
  });
}

/**
 * The set of line indices hidden by COLLAPSED groups on `axis`. A collapsed
 * group hides its body [start, end] (the summary line `end + 1` stays visible).
 * Nested collapsed groups union naturally. A line is hidden if ANY collapsed
 * group on the axis covers it — so collapsing an outer group hides every inner
 * line even when the inner group is itself expanded (Excel parity: an outer
 * collapse hides the whole block).
 */
function hiddenLines(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
): Set<number> {
  const hidden = new Set<number>();
  for (const g of groups) {
    if (g.axis !== axis || !g.collapsed) continue;
    for (let i = g.start; i <= g.end; i += 1) hidden.add(i);
  }
  return hidden;
}

/** Rows hidden by collapsed row groups (see hiddenLines). */
export function hiddenRowsFromGroups(
  groups: readonly OutlineGroup[],
): Set<number> {
  return hiddenLines(groups, "row");
}

/** Columns hidden by collapsed column groups (see hiddenLines). */
export function hiddenColsFromGroups(
  groups: readonly OutlineGroup[],
): Set<number> {
  return hiddenLines(groups, "col");
}

/**
 * The render model for one line in the outline gutter beside the row numbers /
 * above the column letters:
 *   - level:    the deepest level of any group whose body covers this line
 *               (drives the bracket indent); 0 when no group covers it.
 *   - toggle:   present when this line is the SUMMARY line of a group (the line
 *               just past a group's body, group.end + 1, or with summaries
 *               before the detail the line just before it, group.start - 1);
 *               carries the group's range + collapsed state so the UI can
 *               render [+]/[-] and call toggleCollapse on click. When several
 *               groups summarize on the same line, the innermost (narrowest)
 *               wins — that's the toggle the user expects to flip first.
 *   - bracketTop / bracketBottom: whether a bracket end-cap is drawn here (the
 *               first / last body line of the deepest covering group), so the
 *               gutter can draw a spanning vertical (rows) / horizontal (cols)
 *               line with caps.
 */
export interface GutterCell {
  index: number;
  level: number;
  /** The innermost group summarized on this line. */
  toggle: { start: number; end: number; collapsed: boolean } | null;
  /** Every group summarized on this line, outermost first: nested groups
   *  that end on the same line each keep their own button (Excel draws one
   *  per outline level). */
  toggles: Array<{ start: number; end: number; collapsed: boolean }>;
  bracketTop: boolean;
  bracketBottom: boolean;
}

/**
 * Build the gutter render model for every line of `axis` in [0, lineCount).
 * Pure: the UI maps each cell to a bracket segment + optional toggle button.
 * A line hidden by a collapsed group still gets a cell (the renderer skips
 * hidden lines on its own); the toggle on a collapsed group's summary line is
 * what the user clicks to expand it back.
 */
export function groupBracketsForGutter(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  lineCount: number,
  summaryAfter = true,
): GutterCell[] {
  const count = Math.max(0, Math.trunc(Number(lineCount) || 0));
  const sameAxis = groups.filter((g) => g.axis === axis);
  const cells: GutterCell[] = [];
  for (let i = 0; i < count; i += 1) {
    // Deepest group whose BODY covers this line drives the bracket indent.
    let level = 0;
    let bracketTop = false;
    let bracketBottom = false;
    for (const g of sameAxis) {
      if (i < g.start || i > g.end) continue;
      const lvl = levelOf(g, sameAxis);
      if (lvl > level) {
        level = lvl;
        bracketTop = i === g.start;
        bracketBottom = i === g.end;
      } else if (lvl === level) {
        // Same depth (siblings cannot overlap, so this is the same group's
        // single covering); keep the caps already computed.
        bracketTop = bracketTop || i === g.start;
        bracketBottom = bracketBottom || i === g.end;
      }
    }
    // The toggle sits on a group's SUMMARY line (end + 1, or start - 1 with
    // summaries before the detail). Nested groups ending on the same line
    // each keep theirs (live audit S-047: the outer group's button was lost
    // and only the outline level buttons could collapse it).
    const toggles = sameAxis
      .filter((g) => outlineSummaryLine(g, summaryAfter) === i)
      .sort((a, b) => (b.end - b.start) - (a.end - a.start))
      .map((g) => ({ start: g.start, end: g.end, collapsed: g.collapsed }));
    const toggle = toggles[toggles.length - 1] ?? null;
    cells.push({ index: i, level, toggle, toggles, bracketTop, bracketBottom });
  }
  return cells;
}

/**
 * Normalize a persisted outline-group list (round-trip / hand-edited file).
 * Drops malformed entries, clamps each range into [0, lineCount-1] for its
 * axis, rejects crossing/duplicate entries (rebuilding the legal subset via
 * createGroup so the nesting + depth-cap invariants always hold), and bounds
 * the count so a malicious file cannot build a huge list. Returns a NEW array;
 * empty when nothing valid survives so the no-group render path is unchanged.
 */
export function normalizeOutlineGroups(
  raw: unknown,
  bounds: { rowCount: number; columnCount: number },
): OutlineGroup[] {
  if (!Array.isArray(raw)) return [];
  const rowMax = Math.max(0, Math.trunc(bounds.rowCount) - 1);
  const colMax = Math.max(0, Math.trunc(bounds.columnCount) - 1);
  // Hard cap on entries (DoS guard); each axis can hold at most one group per
  // addressable line, so the grid dimensions bound it generously.
  const limit = Math.max(rowMax, colMax) + 1;
  let rows: OutlineGroup[] = [];
  let cols: OutlineGroup[] = [];
  for (const entry of raw.slice(0, limit * 2)) {
    if (!entry || typeof entry !== "object") continue;
    const e = entry as Record<string, unknown>;
    const axis: OutlineAxis = e.axis === "col" ? "col" : "row";
    const max = axis === "row" ? rowMax : colMax;
    const norm = normalizeRange(e.start, e.end);
    if (!norm) continue;
    const start = Math.min(norm.start, max);
    const end = Math.min(norm.end, max);
    // A single line is a legal group: collapsing it hides that line and its
    // toggle sits on the line below, as Excel draws a one-row group.
    if (end < start) continue;
    const collapsed = e.collapsed === true;
    const held = axis === "row" ? rows : cols;
    // The same lines twice are two levels, as the file holds them: a lone
    // <col outlineLevel="2"/> reads as a group inside a group over one column
    // (xlsx-rounds R166-X01). Dropped here, a later outline command wrote it
    // back at level 1. Only a stored outline stacks them; createGroup, what a
    // new group goes through, still takes one.
    const repeats = held.some((g) => g.start === start && g.end === end);
    if (repeats) {
      const depth = depthOfRange({ start, end }, held);
      if (depth !== null && depth <= MAX_OUTLINE_DEPTH) {
        held.push({ axis, start, end, collapsed });
      }
      continue;
    }
    const added = createGroup(
      axis === "row" ? rows : cols,
      [{ start, end }],
      axis,
    );
    // createGroup may reject (crossing/dup/depth); only adopt when it grew.
    if (axis === "row") {
      if (added.length > rows.length) {
        rows = added;
        if (collapsed) {
          rows = toggleCollapse(rows, { start, end }, "row", true);
        }
      }
    } else if (added.length > cols.length) {
      cols = added;
      if (collapsed) {
        cols = toggleCollapse(cols, { start, end }, "col", true);
      }
    }
  }
  return [...rows, ...cols];
}

// ---------------------------------------------------------------------------
// Auto Outline (Excel "数据 > 分级显示 > 自动建立分级显示")
//
// Excel derives an outline from the SUMMARY-FORMULA structure of a sheet: a cell
// holding an aggregate over a contiguous block of cells in the SAME column
// directly above it (e.g. B5 = SUM(B2:B4)) is a row summary, so [B2:B4] becomes a
// collapsible row group whose summary line is B5. The mirror holds for columns:
// E2 = SUM(B2:D2) makes [B2:D2] a column group summarized by E2. We follow the
// "summary BELOW / RIGHT OF detail" convention the rest of this module uses
// (group body = [start, end], summary = end + 1) unless the sheet's Outline
// Settings put summaries above / left of the detail: then a summary cell
// heads the block directly below / right of it (B2 = SUM(B3:B5) groups rows
// 3..5), as Excel's Auto Outline reads such a sheet.
//
// Detection is purely structural and works off each cell's RAW formula text via a
// `formulaAt(row, col)` accessor (returns the stored string, "" / null for a
// non-formula cell). It recognises the aggregate functions Excel's auto-outline
// keys on (SUM and the SUBTOTAL/aggregate family) over a single rectangular
// A1 range whose far edge abuts the summary cell. Everything is pure and
// data-only so it is unit-testable without the grid or the formula engine.
// ---------------------------------------------------------------------------

/** Aggregate functions whose presence marks a cell as an outline summary. */
const OUTLINE_SUMMARY_FUNCS = new Set([
  "SUM",
  "SUBTOTAL",
  "AGGREGATE",
  "AVERAGE",
  "COUNT",
  "COUNTA",
  "MAX",
  "MIN",
  "PRODUCT",
]);

/** A1 column letters → 0-based column index (A→0, Z→25, AA→26). */
function lettersToCol(letters: string): number {
  let n = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i += 1) {
    n = n * 26 + (up.charCodeAt(i) - 64);
  }
  return n - 1;
}

/** A parsed rectangular A1 range (0-based, inclusive, ordered). */
interface ParsedRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * Parse the FIRST simple rectangular A1 range out of a summary formula — e.g.
 * "=SUM(B2:B4)" → {r1:1,c1:1,r2:3,c2:1}. Only a bare `A1:B2` range (optionally
 * with `$` anchors, which are ignored) inside the formula is recognised; a
 * single cell, a 3-D / cross-sheet ref, a named range or a multi-area union
 * returns null (auto-outline only groups a straightforward contiguous block).
 */
function parseSummaryRange(formula: string): ParsedRange | null {
  // A two-endpoint range with no sheet qualifier (no "!") and optional $ anchors.
  const m = formula.match(
    /(^|[^A-Za-z0-9_!.])\$?([A-Za-z]{1,3})\$?([0-9]{1,7}):\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?![A-Za-z0-9_(])/,
  );
  if (!m) return null;
  const c1 = lettersToCol(m[2]);
  const r1 = Number(m[3]) - 1;
  const c2 = lettersToCol(m[4]);
  const r2 = Number(m[5]) - 1;
  if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return null;
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/** True when the formula's leading function token is an outline aggregate. */
function isSummaryFormula(formula: string): boolean {
  const m = formula.match(/^\s*=\s*([A-Za-z][A-Za-z0-9_.]*)\s*\(/);
  if (!m) return false;
  return OUTLINE_SUMMARY_FUNCS.has(m[1].toUpperCase());
}

export interface AutoOutlineOptions {
  rowCount: number;
  columnCount: number;
  /** Raw stored text of a cell (a formula keeps its leading `=`); "" when empty. */
  formulaAt: (row: number, col: number) => string | null | undefined;
  /** The sheet's Outline Settings; absent means summaries below / right. */
  settings?: OutlineSummarySettings | null;
}

/**
 * Derive outline groups from a sheet's summary-formula structure (Excel's
 * "Auto Outline"). Scans every cell for an aggregate summary formula
 * (SUM/SUBTOTAL/…) over a single contiguous range that abuts the summary cell:
 *
 *  - ROW group: a VERTICAL range in the summary cell's own column whose last row
 *    is exactly the row above the summary (B5 = SUM(B2:B4) → row group rows 1..3,
 *    summary row 4). The range must be at least two rows tall.
 *  - COL group: a HORIZONTAL range in the summary cell's own row whose last
 *    column is exactly the column left of the summary (E2 = SUM(B2:D2) → col
 *    group cols 1..3, summary col 4). At least two columns wide.
 *
 * Detected ranges are fed through {@link createGroup} so the result already
 * satisfies the module's nesting / containment / depth invariants (a summary of
 * summaries naturally nests, crossing ranges are dropped). Returns a NEW group
 * list; empty when nothing structural is found. PURE — no DOM, no engine.
 */
export function autoOutlineGroups(options: AutoOutlineOptions): OutlineGroup[] {
  const rowCount = Math.max(0, Math.trunc(options.rowCount));
  const colCount = Math.max(0, Math.trunc(options.columnCount));
  const rowRanges: Array<{ start: number; end: number }> = [];
  const colRanges: Array<{ start: number; end: number }> = [];
  const rowSummaryAfter = outlineSummaryAfter(options.settings, "row");
  const colSummaryAfter = outlineSummaryAfter(options.settings, "col");

  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      const raw = options.formulaAt(row, col);
      if (typeof raw !== "string" || raw === "" || raw[0] !== "=") continue;
      if (!isSummaryFormula(raw)) continue;
      const range = parseSummaryRange(raw);
      if (!range) continue;

      // ROW summary: a vertical block in THIS column whose bottom edge is the row
      // directly above this cell. (range.c1 === range.c2 === col, r2 === row - 1.)
      // With summary rows above the detail, its top edge is the row below.
      if (
        range.c1 === col &&
        range.c2 === col &&
        (rowSummaryAfter ? range.r2 === row - 1 : range.r1 === row + 1) &&
        range.r1 < range.r2 // ≥ 2 rows
      ) {
        rowRanges.push({ start: range.r1, end: range.r2 });
      }

      // COL summary: a horizontal block in THIS row whose right edge is the
      // column directly left of this cell. (range.r1 === range.r2 === row,
      // c2 === col - 1.) With summary columns left of the detail, its left
      // edge is the column to the right.
      if (
        range.r1 === row &&
        range.r2 === row &&
        (colSummaryAfter ? range.c2 === col - 1 : range.c1 === col + 1) &&
        range.c1 < range.c2 // ≥ 2 cols
      ) {
        colRanges.push({ start: range.c1, end: range.c2 });
      }
    }
  }

  // Apply the widest ranges first so an outer group is added before the inner
  // ones it contains (createGroup computes nesting depth against what is already
  // present; an inner range added before its container would be at level 1).
  const byWidest = (
    a: { start: number; end: number },
    b: { start: number; end: number },
  ) => b.end - b.start - (a.end - a.start) || a.start - b.start;
  rowRanges.sort(byWidest);
  colRanges.sort(byWidest);

  let groups: OutlineGroup[] = [];
  groups = createGroup(groups, rowRanges, "row");
  groups = createGroup(groups, colRanges, "col");
  return groups;
}

/**
 * XLS-A027 — Data ▸ Outline ▸ Show Detail / Hide Detail on the active line:
 * Hide Detail collapses the innermost open group the line is in (or whose
 * summary line it is); Show Detail expands the innermost collapsed group
 * whose summary line it is, or that it sits in. Null when no group answers.
 */
export function outlineDetailTarget(
  groups: readonly OutlineGroup[],
  axis: OutlineAxis,
  line: number,
  summaryAfter: boolean,
  mode: "show" | "hide",
): { start: number; end: number } | null {
  const candidates = groups.filter((group) => {
    if (group.axis !== axis) return false;
    if (group.collapsed !== (mode === "show")) return false;
    const summary = outlineSummaryLine(group, summaryAfter);
    return (line >= group.start && line <= group.end) || summary === line;
  });
  if (candidates.length === 0) return null;
  const innermost = candidates.reduce((best, group) =>
    group.end - group.start < best.end - best.start ? group : best,
  );
  return { start: innermost.start, end: innermost.end };
}

export { bySpanDesc };
