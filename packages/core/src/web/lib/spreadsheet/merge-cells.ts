/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet merged cells (Excel/WPS parity).
 *
 * A merge joins a rectangular block of cells into one visual cell: the
 * top-left ("anchor") cell keeps its value and is the only one rendered; every
 * other cell in the block is "covered" and not separately rendered. The sheet
 * model stores the merges as a list of spans (`merges`); this module keeps all
 * the merge math out of the React component so it can be unit-tested in
 * isolation and so the toolbar (which adds/removes spans), the renderer (which
 * resolves a cell's rowspan/colspan and which cells are covered), and the
 * formula engine (which must read the anchor value for a covered reference)
 * all agree on the model.
 *
 * Rendering model: the grid is an HTML table. The anchor cell is emitted with
 * `rowSpan` / `colSpan`; covered cells are skipped entirely so the browser's
 * native rowspan/colspan layout takes over. Selection and editing always target
 * the anchor (a click on a covered cell resolves to its anchor).
 *
 * Formula behavior: a covered cell holds no value of its own, so a formula
 * reference to a covered cell reads BLANK (null), while a reference to the
 * anchor reads the anchor's value — matching Excel, where only the top-left of
 * a merged range carries the value.
 */

/** A single merged block, anchored at its top-left (row, col). */
export interface MergeSpan {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  





  coveredStyles?: Readonly<Record<string, Readonly<Record<string, unknown>>>>;
  



  coveredLink?: Readonly<{ hyperlink: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string }>;
  



  coveredLinks?: Readonly<Record<string, Readonly<{ hyperlink: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string; hyperlinkInArea?: boolean }>>>;
  /** The link over the area the file lists after the anchor's own one-cell
   *  link (reading row 941: A11, then A11:M11 the merge): unmerged, the anchor
   *  keeps it beneath its own (`hyperlinkArea`), as the file reads once Excel
   *  unmerges (xlsx-rounds R447-X08 E16, Windows Excel probe n30/x08). With
   *  the links' places in the file (`hyperlinkRank`, `ownRank`). */
  coveredAnchorArea?: Readonly<{ hyperlink: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string; hyperlinkRank?: number; ownRank?: number }>;
  


  coveredValues?: Readonly<Record<string, Readonly<{ value: string; kind: "text" | "number" }>>>;
}

/** A 0-based grid coordinate. */
export interface MergeCoord {
  row: number;
  col: number;
}

/** A rectangular selection (inclusive bounds, any corner order). */
export interface MergeRange {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/**
 * Build the merge span for a selected range. Returns null when the range is a
 * single cell (1x1) — there is nothing to merge — or when any bound is
 * non-finite/negative. Bounds are normalized so the anchor is always the
 * top-left corner regardless of which corner the user dragged from.
 */
export function mergeSpanFor(range: MergeRange): MergeSpan | null {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  if (
    !Number.isFinite(r1) ||
    !Number.isFinite(c1) ||
    !Number.isFinite(r2) ||
    !Number.isFinite(c2)
  ) {
    return null;
  }
  if (r1 < 0 || c1 < 0) return null;
  const rowSpan = r2 - r1 + 1;
  const colSpan = c2 - c1 + 1;
  if (rowSpan <= 1 && colSpan <= 1) return null; // 1x1 is not a merge
  return { row: r1, col: c1, rowSpan, colSpan };
}

/**
 * Build the merge spans for a "Merge Across" of a selected range: each ROW in
 * the selection is merged INDEPENDENTLY across the selected columns (Excel/WPS
 * "Merge Across"). The result is one horizontal span per row whose colSpan is
 * the selection's full width, anchored at that row's leftmost selected column.
 *
 * Returns [] when there is nothing to merge across: a single-column selection
 * (each row would be a 1x1 "merge"), or any non-finite/negative bound. Bounds
 * are normalized so a drag from any corner yields the same spans.
 */
export function mergeAcrossSpans(range: MergeRange): MergeSpan[] {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  if (
    !Number.isFinite(r1) ||
    !Number.isFinite(c1) ||
    !Number.isFinite(r2) ||
    !Number.isFinite(c2)
  ) {
    return [];
  }
  if (r1 < 0 || c1 < 0) return [];
  const colSpan = c2 - c1 + 1;
  if (colSpan <= 1) return []; // single column: each row is 1x1, nothing to merge
  const spans: MergeSpan[] = [];
  for (let r = r1; r <= r2; r += 1) {
    spans.push({ row: r, col: c1, rowSpan: 1, colSpan });
  }
  return spans;
}

/**
 * Coerce a single raw span into a safe, normalized MergeSpan, or null if it is
 * malformed or degenerate (1x1). Spans are clamped so rowSpan/colSpan are
 * positive integers >= 1; a 1x1 result is dropped because it covers nothing.
 */
export function normalizeSpan(raw: unknown): MergeSpan | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const row = Math.floor(Number(r.row));
  const col = Math.floor(Number(r.col));
  const rowSpan = Math.floor(Number(r.rowSpan));
  const colSpan = Math.floor(Number(r.colSpan));
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  if (row < 0 || col < 0) return null;
  const safeRowSpan = Number.isFinite(rowSpan) ? Math.max(1, rowSpan) : 1;
  const safeColSpan = Number.isFinite(colSpan) ? Math.max(1, colSpan) : 1;
  if (safeRowSpan <= 1 && safeColSpan <= 1) return null;
  const covered = r.coveredStyles;
  const link = r.coveredLink as Record<string, unknown> | undefined;
  return {
    row, col, rowSpan: safeRowSpan, colSpan: safeColSpan,
    ...(covered && typeof covered === "object" && !Array.isArray(covered)
      ? { coveredStyles: covered as MergeSpan["coveredStyles"] }
      : {}),
    ...(link && typeof link === "object" && typeof link.hyperlink === "string" && link.hyperlink
      ? { coveredLink: link as unknown as NonNullable<MergeSpan["coveredLink"]> }
      : {}),
    ...(r.coveredLinks && typeof r.coveredLinks === "object" && !Array.isArray(r.coveredLinks)
      ? { coveredLinks: r.coveredLinks as NonNullable<MergeSpan["coveredLinks"]> }
      : {}),
    ...(r.coveredValues && typeof r.coveredValues === "object" && !Array.isArray(r.coveredValues)
      ? { coveredValues: r.coveredValues as NonNullable<MergeSpan["coveredValues"]> }
      : {}),
    ...(r.coveredAnchorArea && typeof r.coveredAnchorArea === "object" && !Array.isArray(r.coveredAnchorArea)
      && typeof (r.coveredAnchorArea as Record<string, unknown>).hyperlink === "string"
      && (r.coveredAnchorArea as Record<string, unknown>).hyperlink
      ? { coveredAnchorArea: r.coveredAnchorArea as NonNullable<MergeSpan["coveredAnchorArea"]> }
      : {}),
  };
}

/**
 * Normalize a raw merges array off a persisted sheet: drop malformed/degenerate
 * spans and de-duplicate any spans sharing the same anchor (last one wins).
 * Returns a fresh array (never mutates the input).
 */
export function normalizeMerges(raw: unknown): MergeSpan[] {
  if (!Array.isArray(raw)) return [];
  const byAnchor = new Map<string, MergeSpan>();
  for (const item of raw) {
    const span = normalizeSpan(item);
    if (!span) continue;
    byAnchor.set(`${span.row}:${span.col}`, span);
  }
  return [...byAnchor.values()];
}

/** True when (row, col) is the anchor (top-left) of the given span. */
export function isAnchor(row: number, col: number, span: MergeSpan): boolean {
  return span.row === row && span.col === col;
}

/** True when (row, col) lies anywhere inside the span's rectangle. */
export function spanContains(
  row: number,
  col: number,
  span: MergeSpan,
): boolean {
  return (
    row >= span.row &&
    row < span.row + span.rowSpan &&
    col >= span.col &&
    col < span.col + span.colSpan
  );
}

/** True when two spans' rectangles overlap (share at least one cell). */
export function spansOverlap(a: MergeSpan, b: MergeSpan): boolean {
  const aR2 = a.row + a.rowSpan - 1;
  const aC2 = a.col + a.colSpan - 1;
  const bR2 = b.row + b.rowSpan - 1;
  const bC2 = b.col + b.colSpan - 1;
  return a.row <= bR2 && b.row <= aR2 && a.col <= bC2 && b.col <= aC2;
}

/**
 * Expand a selection until every merge it touches is fully contained.
 *
 * Excel/WPS treat a merged cell as one indivisible selection unit. A clipboard
 * snapshot that starts from only its anchor (or any covered coordinate) must
 * therefore include the whole span instead of flattening the unselected half.
 * Expansion can expose another adjacent merge, so scan to a fixed point.
 */
export function expandRangeToMergeClosure(
  range: MergeRange,
  spans: readonly MergeSpan[],
): MergeRange {
  let r1 = Math.min(range.r1, range.r2);
  let r2 = Math.max(range.r1, range.r2);
  let c1 = Math.min(range.c1, range.c2);
  let c2 = Math.max(range.c1, range.c2);
  if (
    !Number.isFinite(r1) ||
    !Number.isFinite(r2) ||
    !Number.isFinite(c1) ||
    !Number.isFinite(c2)
  ) {
    return { r1, r2, c1, c2 };
  }

  let changed = true;
  while (changed) {
    changed = false;
    const probe: MergeSpan = {
      row: r1,
      col: c1,
      rowSpan: r2 - r1 + 1,
      colSpan: c2 - c1 + 1,
    };
    for (const span of spans) {
      if (!spansOverlap(probe, span)) continue;
      const nextR1 = Math.min(r1, span.row);
      const nextR2 = Math.max(r2, span.row + span.rowSpan - 1);
      const nextC1 = Math.min(c1, span.col);
      const nextC2 = Math.max(c2, span.col + span.colSpan - 1);
      if (
        nextR1 !== r1 ||
        nextR2 !== r2 ||
        nextC1 !== c1 ||
        nextC2 !== c2
      ) {
        r1 = nextR1;
        r2 = nextR2;
        c1 = nextC1;
        c2 = nextC2;
        changed = true;
      }
    }
  }
  return { r1, r2, c1, c2 };
}

/**
 * The set of covered (non-anchor) cell keys ("row:col") across every span.
 * These are the cells the renderer must skip and the formula engine must read
 * as blank. The anchor of each span is never in the set.
 */
export function coveredCells(spans: readonly MergeSpan[]): Set<string> {
  const covered = new Set<string>();
  for (const span of spans) {
    for (let r = 0; r < span.rowSpan; r += 1) {
      for (let c = 0; c < span.colSpan; c += 1) {
        if (r === 0 && c === 0) continue; // anchor
        covered.add(`${span.row + r}:${span.col + c}`);
      }
    }
  }
  return covered;
}

/**
 * The anchor cell keys ("row:col") of every span: with `coveredCells` these
 * are all the cells a merge owns. A dynamic array may not spill into any of
 * them, the blank anchor included (Excel: #SPILL! when the spill range
 * "contains a merged cell"), while the renderer still skips only the covered.
 */
export function mergeAnchorCells(spans: readonly MergeSpan[]): Set<string> {
  const anchors = new Set<string>();
  for (const span of spans) anchors.add(`${span.row}:${span.col}`);
  return anchors;
}

/** Where a merge whose anchor is not drawn stands, and how far it reaches. */
export interface MergeStandIn {
  /** The merge's own anchor ("row:col"), hidden or outside the drawn rows. */
  anchor: string;
  /** The drawn rows and columns of the merge from the stand-in on. */
  rows: readonly number[];
  cols: readonly number[];
}

/**
 * The cells that stand in for a merge whose anchor is not drawn: its row or
 * column is hidden (or outside the rows drawn), while other cells of the
 * merge are. The renderer skips every covered cell, so the drawn rest of the
 * merge left no cells at all and the row's later cells slid left under the
 * wrong columns (S-077: J7:K8 with row 7 hidden put L8 under J). The merge's
 * first drawn cell stands in, spanning the drawn rest -- blank, as Excel shows
 * a merge whose top row is hidden (its value lives in the hidden anchor).
 * Keyed by the stand-in cell ("row:col"); `drawnRows`/`drawnCols` in order.
 */
export function mergeStandIns(
  spans: readonly MergeSpan[],
  drawnRows: readonly number[],
  drawnCols: readonly number[],
): Map<string, MergeStandIn> {
  const standIns = new Map<string, MergeStandIn>();
  if (!spans.length) return standIns;
  const rowSet = new Set(drawnRows);
  const colSet = new Set(drawnCols);
  for (const span of spans) {
    if (span.rowSpan * span.colSpan <= 1) continue;
    if (rowSet.has(span.row) && colSet.has(span.col)) continue;
    // A span's own lines when it is shorter than the drawn window, else the
    // drawn window's lines inside it: a merge far from the window costs its
    // size, not the window's.
    const within = (start: number, size: number, drawn: readonly number[], set: ReadonlySet<number>) =>
      size <= drawn.length
        ? Array.from({ length: size }, (_, offset) => start + offset).filter((line) => set.has(line))
        : drawn.filter((line) => line >= start && line < start + size);
    const rows = within(span.row, span.rowSpan, drawnRows, rowSet);
    const cols = within(span.col, span.colSpan, drawnCols, colSet);
    if (!rows.length || !cols.length) continue;
    standIns.set(`${rows[0]}:${cols[0]}`, { anchor: `${span.row}:${span.col}`, rows, cols });
  }
  return standIns;
}

/** True when (row, col) is covered by (but is not the anchor of) some span. */
export function isCovered(
  row: number,
  col: number,
  spans: readonly MergeSpan[],
): boolean {
  for (const span of spans) {
    if (isAnchor(row, col, span)) continue;
    if (spanContains(row, col, span)) return true;
  }
  return false;
}

/**
 * Resolve the span whose ANCHOR is exactly (row, col), or null when no merge is
 * anchored there. Used by the renderer to decide a cell's rowspan/colspan.
 */
export function spanAt(
  row: number,
  col: number,
  spans: readonly MergeSpan[],
): MergeSpan | null {
  for (const span of spans) {
    if (isAnchor(row, col, span)) return span;
  }
  return null;
}

/**
 * Resolve any (row, col) — anchor or covered — to its merge's anchor coordinate.
 * Returns the input coordinate unchanged when the cell is not part of any merge.
 * Used so selection/editing of a covered cell snaps to the anchor.
 */
export function anchorOf(
  row: number,
  col: number,
  spans: readonly MergeSpan[],
): MergeCoord {
  for (const span of spans) {
    if (spanContains(row, col, span)) {
      return { row: span.row, col: span.col };
    }
  }
  return { row, col };
}

/**
 * Where a jump (Name Box, Go To, a link to a cell) lands: a merged cell is one
 * unit. A single cell inside a merged area is that merged cell, at its
 * top-left -- the one cell that holds its content and takes the keyboard --
 * as a click on it is; a range grows to hold every merge it touches, from
 * its top-left corner (live audit S-058: B17 inside A17:E17 left the
 * keyboard nowhere and the formula bar empty).
 */
export function mergeAwareJump(
  range: MergeRange,
  spans: readonly MergeSpan[],
): { active: MergeCoord; anchor: MergeCoord } {
  if (range.r1 === range.r2 && range.c1 === range.c2) {
    const at = anchorOf(range.r1, range.c1, spans);
    return { active: at, anchor: at };
  }
  const grown = expandRangeToMergeClosure(range, spans);
  return {
    active: { row: grown.r1, col: grown.c1 },
    anchor: { row: grown.r2, col: grown.c2 },
  };
}

/**
 * The cells AutoFit passes over (Excel: AutoFit cannot be used on merged
 * cells). A merge spread across several columns lays its text out over their
 * combined width, so it never says how wide ONE column must be; likewise a
 * merge spread across several rows says nothing about one row's height.
 * Returns a predicate for (row, col): true when the cell, anchor or covered,
 * belongs to a merge spanning more than one column (`"column"`: fitting a
 * width) or more than one row (`"row"`: fitting a height).
 */
export function autoFitSkipsMergedCell(
  spans: readonly MergeSpan[],
  axis: "column" | "row",
): (row: number, col: number) => boolean {
  const spread = spans.filter((span) =>
    axis === "column" ? span.colSpan > 1 : span.rowSpan > 1,
  );
  if (spread.length === 0) return () => false;
  return (row, col) => spread.some((span) => spanContains(row, col, span));
}

/**
 * What merging a span does to the cells' contents.
 *
 * Microsoft Support, "Merge and unmerge cells": "the contents of only one cell
 * (the upper-left cell for left-to-right languages ...) appear in the merged
 * cell. The contents of the other cells that you merge are deleted." Excel asks
 * before it deletes anything: the confirmation is raised only when MORE THAN
 * ONE of the cells holds data, because a single value is never lost — it ends
 * up in the merged cell wherever in the span it sat.
 *
 * `dataCells` counts the cells of the span holding data (any non-empty value,
 * including a formula), and `keep` is the upper-left-most of them in reading
 * order — the content the merged cell shows — or null for an empty span.
 */
export function mergeContentCarry(
  span: MergeSpan,
  valueAt: (row: number, col: number) => string | null | undefined,
): { dataCells: number; keep: MergeCoord | null } {
  let dataCells = 0;
  let keep: MergeCoord | null = null;
  for (let r = 0; r < span.rowSpan; r += 1) {
    for (let c = 0; c < span.colSpan; c += 1) {
      const row = span.row + r;
      const col = span.col + c;
      const value = valueAt(row, col);
      if (typeof value !== "string" || value === "") continue;
      dataCells += 1;
      if (!keep) keep = { row, col };
    }
  }
  return { dataCells, keep };
}

/**
 * True when every cell of the range belongs to a merge — the state in which
 * Excel draws Merge & Center pressed and a click on it unmerges instead of
 * merging (the button reports Range.MergeCells, true only when the whole
 * selection is merged). Pass a range already expanded to its merge closure.
 */
export function rangeFullyMerged(
  spans: readonly MergeSpan[],
  range: MergeRange,
): boolean {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  if (!Number.isFinite(r1) || !Number.isFinite(c1)) return false;
  if (spans.length === 0) return false;
  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      if (!spans.some((span) => spanContains(row, col, span))) return false;
    }
  }
  return true;
}

/**
 * Step a keyboard move past the merge it starts in, then land on the anchor.
 *
 * Excel treats a merged range as ONE cell for the arrow keys, Tab and Enter
 * (Microsoft Support, "Merge and unmerge cells": merged cells behave as a
 * single cell). From a merge at A1:B1 the right arrow goes to C1, never to the
 * covered B1; coming back from C1 selects the merge with A1 active. The same
 * holds vertically. This helper takes the origin the move started from and the
 * naive target the caller computed, and returns the coordinate the move should
 * actually land on:
 *
 *  - when the origin lies in a merge and the naive target is still inside that
 *    merge, the target is pushed just past the merge's edge in the direction of
 *    travel (a diagonal move — Enter back to a Tab run's first column — pushes
 *    on each axis it moves along);
 *  - when the landing cell is covered by a merge it resolves to that merge's
 *    anchor, so the active cell is the one that actually carries the value and
 *    owns the rendered input. `keepCovered` (a Shift extension, whose far end is
 *    only a rectangle corner) skips that last resolution.
 *
 * Coordinates are not clamped to the grid; the caller clamps as usual.
 */
export function mergeAwareMoveTarget(
  spans: readonly MergeSpan[],
  origin: MergeCoord,
  target: MergeCoord,
  options: { keepCovered?: boolean } = {},
): MergeCoord {
  if (spans.length === 0) return target;
  let { row, col } = target;
  const originSpan = spans.find((span) =>
    spanContains(origin.row, origin.col, span),
  );
  if (originSpan) {
    const lastRow = originSpan.row + originSpan.rowSpan - 1;
    const lastCol = originSpan.col + originSpan.colSpan - 1;
    if (row > origin.row && row <= lastRow) row = lastRow + 1;
    else if (row < origin.row && row >= originSpan.row) row = originSpan.row - 1;
    if (col > origin.col && col <= lastCol) col = lastCol + 1;
    else if (col < origin.col && col >= originSpan.col) col = originSpan.col - 1;
  }
  if (options.keepCovered) return { row, col };
  return anchorOf(row, col, spans);
}

/**
 * Add a new merge span, removing any existing spans it overlaps (Excel merges
 * over an existing merge by replacing it). Degenerate/duplicate spans are
 * de-duplicated by anchor. Returns a fresh array.
 */
export function addMerge(
  spans: readonly MergeSpan[],
  next: MergeSpan,
): MergeSpan[] {
  const kept = spans.filter((existing) => !spansOverlap(existing, next));
  const byAnchor = new Map<string, MergeSpan>();
  for (const span of kept) byAnchor.set(`${span.row}:${span.col}`, span);
  byAnchor.set(`${next.row}:${next.col}`, next);
  return [...byAnchor.values()];
}

/**
 * Remove every merge span that intersects the given range (unmerge). A range
 * touching any part of a merged block clears that whole block's merge, like
 * Excel's Unmerge. Returns a fresh array.
 */
export function removeMergesInRange(
  spans: readonly MergeSpan[],
  range: MergeRange,
): MergeSpan[] {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  const probe: MergeSpan = {
    row: r1,
    col: c1,
    rowSpan: r2 - r1 + 1,
    colSpan: c2 - c1 + 1,
  };
  return spans.filter((span) => !spansOverlap(span, probe));
}









export function cellsAfterMergesRemoved<C extends { rowspan?: number; colspan?: number }>(
  cells: Record<string, C>,
  before: readonly MergeSpan[],
  after: readonly MergeSpan[],
): Record<string, C> {
  const kept = new Set(after.map((span) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  let next = cells;
  for (const span of before) {
    if (kept.has(`${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`)) continue;
    const key = `${span.row}:${span.col}`;
    const cell = next[key];
    if (!cell || (cell.rowspan === undefined && cell.colspan === undefined)) continue;
    const rest = { ...cell };
    delete rest.rowspan;
    delete rest.colspan;
    if (next === cells) next = { ...cells };
    next[key] = rest;
  }
  return next;
}

/** True when any merge span intersects the given range. */
export function rangeHasMerge(
  spans: readonly MergeSpan[],
  range: MergeRange,
): boolean {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  const probe: MergeSpan = {
    row: r1,
    col: c1,
    rowSpan: r2 - r1 + 1,
    colSpan: c2 - c1 + 1,
  };
  return spans.some((span) => spansOverlap(span, probe));
}

/**
 * Remap merge spans through a row permutation produced by a row sort.
 *
 * `order[destRow] = srcRow` describes where each destination row's data came
 * from (the same array a sort uses to permute the cell map). For each span, its
 * covered source rows must remain CONTIGUOUS and in their original relative
 * order after the permutation — otherwise the merge rectangle would be torn
 * apart and the span is DROPPED. Surviving spans keep their column position and
 * size; only the anchor row is moved to its new destination. Columns are never
 * permuted by a row sort, so colSpan and col are untouched.
 *
 * Returns a fresh array containing only the spans that survive the permutation.
 */
export function remapMergesByRowPermutation(
  spans: readonly MergeSpan[],
  order: readonly number[],
): MergeSpan[] {
  // Invert the permutation: srcRow -> destRow.
  const srcToDest = new Map<number, number>();
  order.forEach((srcRow, destRow) => {
    srcToDest.set(srcRow, destRow);
  });
  const result: MergeSpan[] = [];
  for (const span of spans) {
    // A single-row span (rowSpan === 1) only needs its anchor row remapped.
    let torn = false;
    let firstDest: number | null = null;
    for (let r = 0; r < span.rowSpan; r += 1) {
      const dest = srcToDest.get(span.row + r);
      if (dest === undefined) {
        torn = true;
        break;
      }
      if (firstDest === null) {
        firstDest = dest;
      } else if (dest !== firstDest + r) {
        // Covered rows did not stay contiguous & ordered: tear the merge apart.
        torn = true;
        break;
      }
    }
    if (torn || firstDest === null) continue;
    result.push({ ...span, row: firstDest });
  }
  return result;
}
