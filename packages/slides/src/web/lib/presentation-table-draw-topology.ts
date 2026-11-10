/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  buildOccupancyGrid,
  MAX_PRESENTATION_TABLE_CELLS,
  MAX_TABLE_SPLIT_COLUMNS,
  MAX_TABLE_SPLIT_ROWS,
  type MergeableCell,
} from "@mrnote/core/web/lib/presentation-table-merge.ts";

/** Pencil topology is local cell rectangles, projected onto DrawingML's shared
 * physical axes. Inserting an axis does not split unrelated neighbouring cells. */
export type DrawTableGrid = {
  rows: MergeableCell[][];
  colWidths: number[];
  rowHeights: number[];
};
type Rect = {
  cell: MergeableCell;
  row: number;
  col: number;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
};
const MIN_SIZE = 4;
const rounded = (value: number) => Math.round(value * 9525) / 9525;
const bounds = (sizes: number[]) =>
  sizes.reduce<number[]>((a, n) => [...a, rounded(a[a.length - 1] + n)], [0]);

function rectangles(grid: DrawTableGrid): Rect[] | null {
  if (
    // The table's own bounds: the pen draws on any table the page holds.
    grid.colWidths.length > MAX_TABLE_SPLIT_COLUMNS ||
    grid.rowHeights.length > MAX_TABLE_SPLIT_ROWS ||
    grid.colWidths.length * grid.rowHeights.length > MAX_PRESENTATION_TABLE_CELLS
  )
    return null;
  if (
    !grid.colWidths.every((n) => Number.isFinite(n) && n > 0) ||
    !grid.rowHeights.every((n) => Number.isFinite(n) && n > 0) ||
    grid.rows.length !== grid.rowHeights.length
  )
    return null;
  const occupancy = buildOccupancyGrid(grid.rows);
  if (!occupancy) return null;
  const xs = bounds(grid.colWidths),
    ys = bounds(grid.rowHeights);
  const result: Rect[] = [];
  for (let r = 0; r < grid.rows.length; r++) {
    for (let c = 0; c < grid.colWidths.length; c++) {
      const at = occupancy[r]?.[c];
      if (!at) return null;
      if (!at.isAnchor) continue;
      const cell = grid.rows[at.anchorRow][at.cellIndex];
      const right = c + (cell.colspan ?? 1),
        bottom = r + (cell.rowspan ?? 1);
      if (right >= xs.length || bottom >= ys.length) return null;
      result.push({
        cell:
          typeof cell.drawCellId === "string"
            ? cell
            : {
                ...cell,
                drawCellId: crypto.randomUUID(),
                drawSource: { row: r, col: c },
              },
        row: r,
        col: at.cellIndex,
        x0: xs[c],
        x1: xs[right],
        y0: ys[r],
        y1: ys[bottom],
      });
    }
  }
  return result.length === grid.rows.reduce((sum, row) => sum + row.length, 0)
    ? result
    : null;
}

function materialize(rects: Rect[]): DrawTableGrid | null {
  const xs = [
    ...new Set(rects.flatMap((r) => [rounded(r.x0), rounded(r.x1)])),
  ].sort((a, b) => a - b);
  const ys = [
    ...new Set(rects.flatMap((r) => [rounded(r.y0), rounded(r.y1)])),
  ].sort((a, b) => a - b);
  if (
    xs.length > MAX_TABLE_SPLIT_COLUMNS + 1 ||
    ys.length > MAX_TABLE_SPLIT_ROWS + 1 ||
    (xs.length - 1) * (ys.length - 1) > MAX_PRESENTATION_TABLE_CELLS
  )
    return null;
  const rows: MergeableCell[][] = Array.from(
    { length: ys.length - 1 },
    () => [],
  );
  for (const rect of [...rects].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0)) {
    if (rect.x1 - rect.x0 < MIN_SIZE || rect.y1 - rect.y0 < MIN_SIZE)
      return null;
    const row = ys.indexOf(rounded(rect.y0)),
      col = xs.indexOf(rounded(rect.x0));
    const cell = { ...rect.cell };
    const rowspan = ys.indexOf(rounded(rect.y1)) - row;
    const colspan = xs.indexOf(rounded(rect.x1)) - col;
    if (rowspan > 1) cell.rowspan = rowspan;
    else delete cell.rowspan;
    if (colspan > 1) cell.colspan = colspan;
    else delete cell.colspan;
    rows[row].push(cell);
  }
  const result = {
    rows,
    colWidths: xs.slice(1).map((v, i) => rounded(v - xs[i])),
    rowHeights: ys.slice(1).map((v, i) => rounded(v - ys[i])),
  };
  return rectangles(result) ? result : null;
}

function emptyCell(source: MergeableCell): MergeableCell {
  const cell: MergeableCell = { ...source, text: "" };
  for (const key of [
    "runs",
    "paragraphs",
    "href",
    "hyperlink",
    "nativeText",
    "nativeCell",
    "drawRowFamily",
    "drawSource",
  ])
    delete cell[key];
  cell.drawCellId = crypto.randomUUID();
  return cell;
}








export function evenRowSplitTop(cell: { readonly top: number; readonly height: number }, thickness: number): number {
  return cell.top + cell.height / 2 - thickness / 2;
}

export function splitDrawTable(
  grid: DrawTableGrid,
  targets: readonly { row: number; col: number }[],
  axis: "row" | "column",
  coordinate: number,
  border: unknown,
): DrawTableGrid | null {
  const rects = rectangles(grid);
  if (!rects || !Number.isFinite(coordinate)) return null;
  const selected = rects.filter((r) =>
    targets.some((t) => t.row === r.row && t.col === r.col),
  );
  if (!selected.length || selected.length !== targets.length) return null;
  let next = [...rects];
  const processed = new Set<Rect>();
  for (const target of selected) {
    if (processed.has(target)) continue;
    if (axis === "column") {
      const at = rounded(coordinate);
      if (at - target.x0 < MIN_SIZE || target.x1 - at < MIN_SIZE) return null;
      const left: MergeableCell = { ...target.cell, borderRight: border };
      delete left.drawRowFamily;
      next = next.filter((r) => r !== target);
      next.push(
        { ...target, cell: left, x1: at },
        {
          ...target,
          cell: { ...emptyCell(target.cell), borderLeft: border },
          x0: at,
        },
      );
      continue;
    }
    const family = target.cell.drawRowFamily;
    const siblings =
      typeof family === "string"
        ? rects
            .filter(
              (r) =>
                r.cell.drawRowFamily === family &&
                r.x0 === target.x0 &&
                r.x1 === target.x1,
            )
            .sort((a, b) => a.y0 - b.y0)
        : [target];
    if (siblings.some((r, i) => i > 0 && r.y0 !== siblings[i - 1].y1))
      return null;
    const top = siblings[0].y0,
      bottom = siblings[siblings.length - 1].y1;
    const pieces = siblings.length + 1;
    if ((bottom - top) / pieces < MIN_SIZE) return null;
    const id =
      typeof family === "string"
        ? family
        : `pencil:${target.x0}:${target.x1}:${top}:${bottom}`;
    const children = siblings.map((r) => r.cell);
    children.splice(siblings.indexOf(target) + 1, 0, emptyCell(target.cell));
    next = next.filter((r) => !siblings.includes(r));
    siblings.forEach((r) => processed.add(r));
    children.forEach((cell, i) =>
      next.push({
        ...target,
        cell: {
          ...cell,
          drawRowFamily: id,
          ...(i > 0 ? { borderTop: border } : {}),
          ...(i < pieces - 1 ? { borderBottom: border } : {}),
        },
        y0: rounded(top + ((bottom - top) * i) / pieces),
        y1: rounded(top + ((bottom - top) * (i + 1)) / pieces),
      }),
    );
  }
  return materialize(next);
}

/** Move one connected boundary segment, not every unrelated line sharing an
 * underlying OOXML grid coordinate. Neighbouring cells keep their text owner. */
export function moveDrawTableBoundary(
  grid: DrawTableGrid,
  target: { row: number; col: number },
  axis: "row" | "column",
  coordinate: number,
): DrawTableGrid | null {
  const rects = rectangles(grid);
  const selected = rects?.find(
    (r) => r.row === target.row && r.col === target.col,
  );
  if (!rects || !selected || !Number.isFinite(coordinate)) return null;
  const horizontal = axis === "row";
  const edge = horizontal ? selected.y1 : selected.x1;
  let start = horizontal ? selected.x0 : selected.y0;
  let end = horizontal ? selected.x1 : selected.y1;
  let touched: Rect[] = [];
  for (let pass = 0; pass <= rects.length; pass++) {
    const hits = rects.filter(
      (r) =>
        (horizontal
          ? r.y0 === edge || r.y1 === edge
          : r.x0 === edge || r.x1 === edge) &&
        (horizontal ? r.x0 < end && r.x1 > start : r.y0 < end && r.y1 > start),
    );
    if (hits.length === touched.length) break;
    touched = hits;
    start = Math.min(...hits.map((r) => (horizontal ? r.x0 : r.y0)));
    end = Math.max(...hits.map((r) => (horizontal ? r.x1 : r.y1)));
  }
  if (!touched.some((r) => (horizontal ? r.y0 === edge : r.x0 === edge)))
    return null;
  const at = rounded(coordinate);
  if (at === edge) return null;
  return materialize(
    rects.map((r) => {
      if (!touched.includes(r)) return r;
      const cell = { ...r.cell };
      delete cell.drawRowFamily;
      return {
        ...r,
        cell,
        ...(horizontal
          ? r.y0 === edge
            ? { y0: at }
            : { y1: at }
          : r.x0 === edge
            ? { x0: at }
            : { x1: at }),
      };
    }),
  );
}
