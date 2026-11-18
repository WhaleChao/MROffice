/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Merge & Center, Merge Across, Merge Cells and Unmerge Cells as pure
 * sheet transactions. The window's buttons and the spreadsheet editor commands
 * share them; behaviour is unchanged (the window still asks before a merge
 * discards values; a command proceeds, as Excel's Range.Merge does).
 */

import { styleStatesAFormat } from "./style-source-xf";
import { updateCellStyle, type CellButtonsSheet, type CellMap, type SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { cellHorizontalChoicePatch } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { styleWithOwnTextIdentity } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import { normalizeBorderEdges } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { findEqualRuns, mergeIdenticalCells } from "./merge-identical-cells";
import {
  addMerge,
  expandRangeToMergeClosure,
  mergeAcrossSpans,
  mergeContentCarry,
  mergeSpanFor,
  normalizeMerges,
  rangeFullyMerged,
  rangeHasMerge,
  cellsAfterMergesRemoved,
  removeMergesInRange,
  spansOverlap,
  type MergeSpan,
} from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { isMaterializedLineStyleBlank } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/**
 * The cell map after merging: each span keeps the upper-left-most value it
 * covers (moved onto the anchor when it sat elsewhere, so a lone value is
 * never lost) and every other cell of the span is cleared, so a formula
 * reading a covered cell reads blank.
 */
export function cellMapAfterMerge(spans: readonly MergeSpan[], cells: CellMap): CellMap {
  const next: CellMap = { ...cells };
  for (const span of spans) {
    const { keep } = mergeContentCarry(
      span,
      (row, col) => cells[`${row}:${col}`]?.value,
    );
    const carried =
      keep && (keep.row !== span.row || keep.col !== span.col)
        ? cells[`${keep.row}:${keep.col}`]
        : undefined;
    for (let r = 0; r < span.rowSpan; r += 1) {
      for (let c = 0; c < span.colSpan; c += 1) {
        if (r === 0 && c === 0) continue;
        delete next[`${span.row + r}:${span.col + c}`];
      }
    }
    const anchorKey = `${span.row}:${span.col}`;
    if (carried) {
      // The merged cell shows that content in the anchor's own formatting:
      // Excel takes the merged cell's format from the upper-left cell. The
      // content keeps ITS identity: a quote-prefixed '=1+1 carried into the
      // anchor stays the text =1+1 (live audit B154).
      const anchorCell = next[anchorKey];
      const style = styleWithOwnTextIdentity(anchorCell?.style, carried.style);
      const merged: SpreadsheetCell = {
        ...anchorCell,
        value: carried.value,
        ...(carried.formula === undefined
          ? {}
          : { formula: carried.formula }),
        ...(carried.runs ? { runs: carried.runs } : {}),
        ...(carried.kind ? { kind: carried.kind } : {}),
      };
      // The carried text brings its own furigana; the anchor's described
      // the text it no longer shows.
      if (carried.phonetic) merged.phonetic = carried.phonetic;
      else delete merged.phonetic;
      if (style) merged.style = style;
      else delete merged.style;
      next[anchorKey] = merged;
    }
  }
  return next;
}

/** Whether a merge of `plannedSpans` touches a locked cell under protection:
 *  addMerge removes every existing span it overlaps, so those count too. */
export function mergeLockedByProtection(
  mergeSpans: readonly MergeSpan[],
  plannedSpans: readonly MergeSpan[],
  rangeHasLockedCell: (range: Rect) => boolean,
): boolean {
  if (plannedSpans.length === 0) return false;
  // addMerge removes every existing span it overlaps. Authenticate those
  // complete spans too, not only the newly selected rectangle, so a partial
  // overlap cannot indirectly unmerge a locked cell outside the selection.
  const affectedSpans = [
    ...plannedSpans,
    ...mergeSpans.filter((existing) =>
      plannedSpans.some((planned) => spansOverlap(existing, planned)),
    ),
  ];
  return affectedSpans.some((span) =>
    rangeHasLockedCell({
      r1: span.row,
      r2: span.row + span.rowSpan - 1,
      c1: span.col,
      c2: span.col + span.colSpan - 1,
    }),
  );
}

export type MergeButton = "merge_center" | "merge_across" | "merge_cells" | "unmerge";

/**
 * The sheet after one of the merge buttons on `rect` (grown to the merges it
 * touches, as the selection is): "protected" when protection refuses it, null
 * when the button does nothing there (a single cell, nothing merged). Merge &
 * Center on a fully merged range unmerges it, as the pressed button does.
 */
/** Whether a rectangle reaches into a table (its header, body or total row). */
export function mergeRectTouchesTable(sheet: object, rect: Rect): boolean {
  const tables = (sheet as { dataTables?: unknown }).dataTables;
  if (!Array.isArray(tables)) return false;
  return tables.some((table) => {
    const range = (table as { range?: Partial<Rect> } | null)?.range;
    if (!range || ![range.r1, range.r2, range.c1, range.c2].every(Number.isInteger)) return false;
    const bottom = (range.r2 as number) + ((table as { totalsRow?: unknown }).totalsRow ? 1 : 0);
    return rect.r1 <= bottom && rect.r2 >= (range.r1 as number) && rect.c1 <= (range.c2 as number) && rect.c2 >= (range.c1 as number);
  });
}

export function sheetAfterMergeButton<S extends CellButtonsSheet & { merges?: MergeSpan[] }>(
  sheet: S,
  rect: Rect,
  button: MergeButton,
  rangeHasLockedCell: (range: Rect) => boolean,
): S | "protected" | null {
  const mergeSpans = normalizeMerges(sheet.merges);
  const target = expandRangeToMergeClosure(rect, mergeSpans);
  const unmerge = (): S | "protected" | null => {
    if (!rangeHasMerge(mergeSpans, target)) return null;
    const selectionSpan: MergeSpan = {
      row: target.r1,
      col: target.c1,
      rowSpan: target.r2 - target.r1 + 1,
      colSpan: target.c2 - target.c1 + 1,
    };
    const affected = mergeSpans.filter((span) => spansOverlap(span, selectionSpan));
    if (mergeLockedByProtection(mergeSpans, affected, rangeHasLockedCell)) return "protected";
    const nextMerges = removeMergesInRange(mergeSpans, target);
    const next: S = { ...sheet, cells: cellsAfterUnmerge(sheet.cells, mergeSpans, nextMerges) };
    if (nextMerges.length > 0) next.merges = nextMerges;
    else delete next.merges;
    return next;
  };
  if (button === "unmerge") return unmerge();
  if (button === "merge_center" && rangeFullyMerged(mergeSpans, target)) return unmerge();
  const spans =
    button === "merge_across"
      ? mergeAcrossSpans(target)
      : [mergeSpanFor(target)].filter((span): span is MergeSpan => span !== null);
  if (spans.length === 0) return null;
  // Excel greys Merge for cells of a table: a merge across its records
  // breaks the one-row-per-record structure its filters, sorts and
  // structured references read (user audit S-019). Unmerge stays open.
  if (mergeRectTouchesTable(sheet, target)) return null;
  if (mergeLockedByProtection(mergeSpans, spans, rangeHasLockedCell)) return "protected";
  let merges = mergeSpans;
  for (const span of spans) merges = addMerge(merges, span);
  let next: S = { ...sheet, cells: cellMapAfterMerge(spans, sheet.cells), merges };
  if (button === "merge_center") {
    // Merge & Center is the Center pick of the Horizontal list: it replaces
    // a Center Across Selection (or Fill, Justify) the cell had rather than
    // sitting beside it, which the file cannot say and the export dropped
    // (xlsx-rounds R16-X02 E15).
    next = updateCellStyle(next, spans[0].row, spans[0].col, cellHorizontalChoicePatch("center"));
  }
  const covered = cellsCoveredInTheAnchorsFormat(next.cells, sheet.cells, spans);
  const lined = mergedAreaLinesAsExcelLeavesThem(covered, sheet.cells, spans);
  return {
    ...next,
    cells: anchorsShowTheirCoveredRightEdge(lined.cells, sheet.cells, lined.unread),
  };
}

const AREA_SIDES = [
  ["top", "borderTop"],
  ["right", "borderRight"],
  ["bottom", "borderBottom"],
  ["left", "borderLeft"],
] as const;
type AreaSide = (typeof AREA_SIDES)[number][0];















function mergedAreaLinesAsExcelLeavesThem(
  cells: CellMap,
  before: CellMap,
  spans: readonly MergeSpan[],
): { cells: CellMap; unread: MergeSpan[] } {
  let next = cells;
  const unread: MergeSpan[] = [];
  for (const span of spans) {
    if (span.rowSpan * span.colSpan < 2) continue;
    const r2 = span.row + span.rowSpan - 1;
    const c2 = span.col + span.colSpan - 1;
    const own = new Map<string, Partial<Record<AreaSide, unknown>>>();
    let legacy = false;
    for (let r = span.row; r <= r2 && !legacy; r += 1) {
      for (let c = span.col; c <= c2; c += 1) {
        const key = `${r}:${c}`;
        const style = before[key]?.style as Record<string, unknown> | undefined;
        const edges = normalizeBorderEdges(style?.borderEdges) ?? {};
        if (style?.border === true && AREA_SIDES.some(([side]) => !edges[side])) legacy = true;
        if (AREA_SIDES.some(([side, flag]) => style?.[flag] === true && !edges[side])) legacy = true;
        if (legacy) break;
        own.set(key, edges);
      }
    }
    if (legacy) {
      unread.push(span);
      continue;
    }
    const lineOf = (r: number, c: number, side: AreaSide) => own.get(`${r}:${c}`)?.[side];
    const anchorSides = AREA_SIDES.map(([side]) => side).filter((side) => lineOf(span.row, span.col, side));
    const spread =
      anchorSides.length === 1
      && ((anchorSides[0] === "right" && span.colSpan > 1) || (anchorSides[0] === "bottom" && span.rowSpan > 1))
        ? anchorSides[0]
        : null;
    const along = (side: AreaSide): Array<[number, number]> => {
      const out: Array<[number, number]> = [];
      if (side === "top" || side === "bottom") {
        for (let c = span.col; c <= c2; c += 1) out.push([side === "top" ? span.row : r2, c]);
      } else {
        for (let r = span.row; r <= r2; r += 1) out.push([r, side === "left" ? span.col : c2]);
      }
      return out;
    };
    const kept = new Set<AreaSide>();
    for (const [side] of AREA_SIDES) {
      if (side === spread) continue;
      const cellsAlong = along(side);
      const first = lineOf(cellsAlong[0]![0], cellsAlong[0]![1], side);
      if (first && cellsAlong.every(([r, c]) => JSON.stringify(lineOf(r, c, side)) === JSON.stringify(first))) kept.add(side);
    }
    const onSide = (r: number, c: number, side: AreaSide) =>
      side === "top" ? r === span.row : side === "bottom" ? r === r2 : side === "left" ? c === span.col : c === c2;
    const result = new Map<string, Record<string, unknown>>();
    for (let r = span.row; r <= r2; r += 1) {
      for (let c = span.col; c <= c2; c += 1) {
        const key = `${r}:${c}`;
        const edges: Record<string, unknown> = { ...(own.get(key) ?? {}) };
        for (const [side] of AREA_SIDES) {
          if (side === spread) edges[side] = lineOf(span.row, span.col, side);
          else if (!(kept.has(side) && onSide(r, c, side))) delete edges[side];
        }
        result.set(key, edges);
      }
    }
    // The anchor draws the area: its outline, not its own record's far sides.
    const anchorKey = `${span.row}:${span.col}`;
    const anchorEdges = { ...(result.get(anchorKey) ?? {}) };
    const farRight = result.get(`${span.row}:${c2}`)?.right;
    const farBottom = result.get(`${r2}:${span.col}`)?.bottom;
    if (farRight) anchorEdges.right = farRight;
    else delete anchorEdges.right;
    if (farBottom) anchorEdges.bottom = farBottom;
    else delete anchorEdges.bottom;
    result.set(anchorKey, anchorEdges);
    for (const [key, edges] of result) {
      const cell = next[key];
      const current = normalizeBorderEdges(cell?.style?.borderEdges);
      const wanted = normalizeBorderEdges(edges);
      if (JSON.stringify(current ?? null) === JSON.stringify(wanted ?? null)) continue;
      if (!cell && !wanted) continue;
      const style: Record<string, unknown> = { ...(cell?.style ?? {}) };
      delete style.border;
      if (wanted) style.borderEdges = wanted;
      else delete style.borderEdges;
      for (const [side, flag] of AREA_SIDES) {
        if (wanted?.[side]) style[flag] = true;
        else delete style[flag];
      }
      if (next === cells) next = { ...cells };
      next[key] = {
        ...(cell ?? { value: "", kind: "text" }),
        style: style as NonNullable<SpreadsheetCell["style"]>,
      } as CellMap[string];
    }
  }
  return { cells: next, unread };
}

/** The cells once the merges of `before` that `after` no longer holds are
 *  unmerged, as Excel leaves them: every covered cell has its own format,
 *  link and value back, and the anchor the lines of its own record. The
 *  Unmerge button's, and Create Table's, which unmerges what it covers
 *  (xlsx-rounds WALK R285-X13 / R285-X16: the anchor kept the area's far line
 *  on the page, the file's record never had it and the save read back
 *  without it). */
export function cellsAfterUnmerge(cells: CellMap, before: readonly MergeSpan[], after: readonly MergeSpan[]): CellMap {
  const kept = new Set(after.map((span) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  const removed = before.filter((span) => !kept.has(`${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  return anchorsBackToTheirOwnEdges(
    cellsWithTheirCoveredValues(
      cellsWithTheirCoveredLinks(
        cellsWithTheirCoveredFormats(cellsAfterMergesRemoved(cells, before, after), before, after),
        before,
        after,
      ),
      before,
      after,
    ),
    removed,
  );
}













function anchorsShowTheirCoveredRightEdge(
  cells: CellMap,
  before: CellMap,
  spans: readonly MergeSpan[],
): CellMap {
  let next = cells;
  const sides = [
    ["right", "borderRight", (span: MergeSpan) => span.colSpan > 1, (span: MergeSpan) => `${span.row}:${span.col + span.colSpan - 1}`],
    ["bottom", "borderBottom", (span: MergeSpan) => span.rowSpan > 1, (span: MergeSpan) => `${span.row + span.rowSpan - 1}:${span.col}`],
  ] as const;
  for (const span of spans) {
    for (const [side, flag, spans2, lastKey] of sides) {
      if (!spans2(span)) continue;
      const line = normalizeBorderEdges(before[lastKey(span)]?.style?.borderEdges)?.[side];
      if (!line) continue;
      const key = `${span.row}:${span.col}`;
      const anchor = next[key];
      const edges = normalizeBorderEdges(anchor?.style?.borderEdges) ?? {};
      if (JSON.stringify(edges[side]) === JSON.stringify(line)) continue;
      const style: Record<string, unknown> = {
        ...(anchor?.style ?? {}),
        borderEdges: normalizeBorderEdges({ ...edges, [side]: line }),
        [flag]: true,
      };
      delete style.border;
      if (next === cells) next = { ...cells };
      next[key] = { ...(anchor ?? { value: "", kind: "text" }), style: style as NonNullable<SpreadsheetCell["style"]> } as CellMap[string];
    }
  }
  return next;
}






export function cellsWithTheirCoveredFormats(
  cells: CellMap,
  before: readonly MergeSpan[],
  after: readonly MergeSpan[],
): CellMap {
  const kept = new Set(after.map((span) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  let next = cells;
  for (const span of before) {
    if (!span.coveredStyles || kept.has(`${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`)) continue;
    for (const [offset, own] of Object.entries(span.coveredStyles)) {
      const [r, c] = offset.split(":").map(Number);
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= span.rowSpan || c >= span.colSpan) continue;
      if ((r === 0 && c === 0) || Object.keys(own).length === 0) continue;
      const key = `${span.row + r}:${span.col + c}`;
      // A blank only its column's format wrote is no cell the page holds: under
      // a merge it stood in for the record (xlsx-rounds R139-X02: a title block
      // came back plain on a sheet with a font on every column).
      if (next[key] && !isMaterializedLineStyleBlank(next[key])) continue;
      if (next === cells) next = { ...cells };
      next[key] = { value: "", kind: "text", style: structuredClone(own) as NonNullable<SpreadsheetCell["style"]> };
    }
  }
  return next;
}






export function cellsWithTheirCoveredLinks(
  cells: CellMap,
  before: readonly MergeSpan[],
  after: readonly MergeSpan[],
): CellMap {
  const kept = new Set(after.map((span) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  let next = cells;
  for (const span of before) {
    const each = span.coveredLinks;
    const whole = span.coveredLink;
    const anchorArea = span.coveredAnchorArea;
    if ((!each && !whole && !anchorArea) || kept.has(`${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`)) continue;
    // The anchor linked on its own keeps the area's link beneath (reading row
    // 941; R447-X08 E16: A11 over A11:M11, which Excel's unmerge leaves).
    if (anchorArea?.hyperlink) {
      const key = `${span.row}:${span.col}`;
      const anchor = next[key] as (CellMap[string] & { hyperlinkArea?: unknown; hyperlinkRank?: unknown }) | undefined;
      if (anchor?.hyperlink && !anchor.hyperlinkArea) {
        const { ownRank, ...area } = anchorArea;
        if (next === cells) next = { ...cells };
        next[key] = {
          ...anchor,
          hyperlinkArea: { ...area },
          ...(typeof ownRank === "number" && anchor.hyperlinkRank === undefined ? { hyperlinkRank: ownRank } : {}),
        } as CellMap[string];
      }
    }
    if (!each && !whole) continue;
    
    // anchor's own; it holds only while the anchor still has it.
    if (!each && next[`${span.row}:${span.col}`]?.hyperlink !== whole?.hyperlink) continue;
    for (let r = 0; r < span.rowSpan; r += 1) {
      for (let c = 0; c < span.colSpan; c += 1) {
        if (r === 0 && c === 0) continue;
        const link = each ? each[`${r}:${c}`] : whole;
        if (!link?.hyperlink) continue;
        const key = `${span.row + r}:${span.col + c}`;
        const own = next[key];
        if (own?.hyperlink) continue;
        if (next === cells) next = { ...cells };
        next[key] = {
          ...(own ?? { value: "", kind: "text" }),
          hyperlink: link.hyperlink,
          ...(link.hyperlinkTooltip ? { hyperlinkTooltip: link.hyperlinkTooltip } : {}),
          ...(link.hyperlinkDisplay ? { hyperlinkDisplay: link.hyperlinkDisplay } : {}),
          ...("hyperlinkInArea" in link && link.hyperlinkInArea === true ? { hyperlinkInArea: true } : {}),
        } as CellMap[string];
      }
    }
  }
  return next;
}





export function cellsWithTheirCoveredValues(
  cells: CellMap,
  before: readonly MergeSpan[],
  after: readonly MergeSpan[],
): CellMap {
  const kept = new Set(after.map((span) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`));
  let next = cells;
  for (const span of before) {
    const values = span.coveredValues;
    if (!values || kept.has(`${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`)) continue;
    for (const [offset, stored] of Object.entries(values)) {
      const [r, c] = offset.split(":").map(Number);
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= span.rowSpan || c >= span.colSpan) continue;
      if ((r === 0 && c === 0) || typeof stored?.value !== "string" || stored.value === "") continue;
      const key = `${span.row + r}:${span.col + c}`;
      const own = next[key];
      if (own && own.value !== "" && own.value !== undefined) continue;
      if (next === cells) next = { ...cells };
      next[key] = { ...(own ?? {}), value: stored.value, kind: stored.kind === "number" ? "number" : "text" } as CellMap[string];
    }
  }
  return next;
}





function anchorsBackToTheirOwnEdges(cells: CellMap, spans: readonly MergeSpan[]): CellMap {
  let next = cells;
  for (const span of spans) {
    const key = `${span.row}:${span.col}`;
    const anchor = next[key];
    const last = next[`${span.row + span.rowSpan - 1}:${span.col + span.colSpan - 1}`];
    const edges = normalizeBorderEdges(anchor?.style?.borderEdges);
    const lastEdges = normalizeBorderEdges(last?.style?.borderEdges);
    // What the anchor's own record states on the far sides (`coveredStyles`
    
    const own = normalizeBorderEdges(span.coveredStyles?.["0:0"]?.borderEdges);
    if (!anchor?.style || (!(edges && lastEdges) && !own)) continue;
    const taken = ([["right", span.colSpan], ["bottom", span.rowSpan]] as const)
      .filter(([side, size]) => size > 1 && edges?.[side] && JSON.stringify(edges[side]) === JSON.stringify(lastEdges?.[side]))
      .map(([side]) => side);
    if (taken.length === 0 && !own) continue;
    const kept = { ...(edges ?? {}) };
    for (const side of taken) delete kept[side];
    Object.assign(kept, own ?? {});
    const style: Record<string, unknown> = { ...anchor.style, borderEdges: normalizeBorderEdges(kept) };
    delete style.border;
    for (const [side, flag] of [["top", "borderTop"], ["right", "borderRight"], ["bottom", "borderBottom"], ["left", "borderLeft"]] as const) {
      if (kept[side]) style[flag] = true;
      else delete style[flag];
    }
    if (style.borderEdges === undefined) delete style.borderEdges;
    if (next === cells) next = { ...cells };
    next[key] = { ...anchor, style: style as NonNullable<typeof anchor.style> };
  }
  return next;
}





function cellsCoveredInTheAnchorsFormat(
  cells: CellMap,
  before: CellMap,
  spans: readonly MergeSpan[],
): CellMap {
  const isBorder = (key: string) => key.startsWith("border");
  let next = cells;
  for (const span of spans) {
    const anchor = cells[`${span.row}:${span.col}`];
    if (!anchor?.style || !styleStatesAFormat(anchor.style)) continue;
    const format = Object.entries(anchor.style).filter(([key]) => !isBorder(key));
    for (let r = 0; r < span.rowSpan; r += 1) {
      for (let c = 0; c < span.colSpan; c += 1) {
        if (r === 0 && c === 0) continue;
        const key = `${span.row + r}:${span.col + c}`;
        const own = Object.entries(before[key]?.style ?? {}).filter(([field]) => isBorder(field));
        if (next === cells) next = { ...cells };
        next[key] = {
          value: "",
          kind: "text",
          style: Object.fromEntries([...format, ...own]) as NonNullable<typeof anchor.style>,
          ...(anchor.locked === false ? { locked: false } : {}),
        };
      }
    }
  }
  return next;
}

/** Merge Identical Cells (WPS): the vertical runs of equal values in each
 *  column of the rectangle, one span per run. */
export function mergeIdenticalPlannedSpans(cells: CellMap, rect: Rect): MergeSpan[] {
  const { r1, r2, c1, c2 } = rect;
  const plannedSpans: MergeSpan[] = [];
  for (let col = c1; col <= c2; col += 1) {
    for (const run of findEqualRuns(cells, col, r1, r2)) {
      plannedSpans.push({
        row: run.row,
        col,
        rowSpan: run.count,
        colSpan: 1,
      });
    }
  }
  return plannedSpans;
}

/** The sheet after Merge Identical Cells on `rect`, or null when no run merges. */
export function sheetAfterMergeIdentical<S extends CellButtonsSheet & { merges?: MergeSpan[] }>(
  sheet: S,
  rect: Rect,
): S | null {
  if (rect.r1 === rect.r2) return null;
  const result = mergeIdenticalCells(sheet.cells, normalizeMerges(sheet.merges), rect.r1, rect.r2, rect.c1, rect.c2);
  if (result.spansCreated === 0) return null;
  const next: S = { ...sheet, cells: result.cells as CellMap };
  if (result.merges.length > 0) next.merges = result.merges;
  else delete next.merges;
  return next;
}

export type { SpreadsheetCell };
