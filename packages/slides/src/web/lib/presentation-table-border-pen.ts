/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation table BORDER PEN — the pointer half of PowerPoint / WPS
 * "Draw Table" and "Eraser".
 *
 * The ribbon border commands apply a border TARGET to the whole cell selection.
 * The pen is different: the user picks up a tool, hovers a single cell edge,
 * presses, and drags across the table drawing or erasing one native edge at a
 * time. That is the interaction WPS users reach for when a table needs one
 * missing rule rather than a rectangle of them.
 *
 * Everything here is pure — no React, no DOM. The component measures the
 * rendered `<td>` boxes, feeds them in, and commits the returned rows as ONE
 * undoable block replacement per gesture. Keeping the geometry maths and the
 * model rewrite out of the component is what makes both unit testable without a
 * browser.
 *
 * An internal edge is written on BOTH adjacent cells, exactly like
 * `tableBorderPatchForSelection`, so the exported DrawingML stays editable in
 * PowerPoint and WPS even though neither collapses neighbouring `a:ln*` nodes.
 */

import { buildOccupancyGrid, type MergeableCell } from "@mrnote/core/web/lib/presentation-table-merge.ts";
import {
  CELL_STYLE_KEYS,
  normalizeCellStyle,
  persistedCellStyle,
  type CellStyleInput,
  type TableCellBorder,
} from "@mrnote/core/web/lib/presentation-table-style.ts";

/** The two pointer tools. `null` (in the component) means "no pen held". */
export const TABLE_BORDER_PEN_MODES = ["draw", "erase"] as const;
export type TableBorderPenMode = (typeof TABLE_BORDER_PEN_MODES)[number];

export type TableBorderEdgeSide = "top" | "right" | "bottom" | "left";

/**
 * One cell edge. `row` / `col` are SOURCE ARRAY indices (the same coordinates
 * the renderer and `updateTableCellStyle` use), never logical grid positions —
 * a merged table's array index and logical column differ.
 */
export interface TableBorderEdgeRef {
  row: number;
  col: number;
  side: TableBorderEdgeSide;
}

/** A rendered cell's box in OVERLAY-LOCAL pixels. */
export interface TableBorderCellBox extends Omit<TableBorderEdgeRef, "side"> {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** The hovered edge plus the overlay-local bar used to highlight it. */
export interface TableBorderEdgeHit extends TableBorderEdgeRef {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Perpendicular pointer distance in px; smaller wins. */
  distance: number;
}

/** How close (px) the pointer must come to an edge before it is a candidate. */
export const TABLE_BORDER_PEN_TOLERANCE = 7;
/** Thickness (px) of the hover / stroke marker drawn over a candidate edge. */
export const TABLE_BORDER_PEN_MARKER_THICKNESS = 3;

const EDGE_STYLE_KEY: Record<TableBorderEdgeSide, keyof CellStyleInput> = {
  top: "borderTop",
  right: "borderRight",
  bottom: "borderBottom",
  left: "borderLeft",
};

const OPPOSITE_SIDE: Record<TableBorderEdgeSide, TableBorderEdgeSide> = {
  top: "bottom",
  right: "left",
  bottom: "top",
  left: "right",
};

/** Stable identity for an edge — used to de-duplicate a drag path. */
export function tableBorderEdgeKey(ref: TableBorderEdgeRef): string {
  return `${ref.row}:${ref.col}:${ref.side}`;
}

/**
 * The edge a pointer hit (a PHYSICAL side of a drawn cell) as the side the
 * table stores. A cell's `borderLeft` / `borderRight` are DrawingML's `a:lnL`
 * / `a:lnR`, and on a right-to-left table (`a:tblPr@rtl`) PowerPoint draws
 * those on the cell's start and end: `a:lnL` on its physical RIGHT (measured
 * in PowerPoint for Mac). Every grid helper below works in that stored,
 * logical frame, where column `c - 1` lies across `left`.
 */
export function tableBorderStoredEdge<T extends TableBorderEdgeRef>(
  edge: T,
  rightToLeft: boolean,
): T {
  if (!rightToLeft || (edge.side !== "left" && edge.side !== "right")) {
    return edge;
  }
  return { ...edge, side: edge.side === "left" ? "right" : "left" };
}

/**
 * The existing borders a pencil stroke from `from` to `to` lies ON, in the
 * order it meets them: the border under a click, or the borders a drag
 * traces along (PowerPoint redraws those in the pen). Empty when the stroke
 * does not start on a border, or crosses the border it starts on (a line
 * drawn through a cell, which is new structure).
 */
export function traceTableBorderEdges(
  boxes: readonly TableBorderCellBox[],
  from: { x: number; y: number },
  to: { x: number; y: number },
  tolerance: number = TABLE_BORDER_PEN_TOLERANCE,
): TableBorderEdgeHit[] {
  const first = hitTestTableBorderEdge(boxes, from, tolerance);
  if (!first) return [];
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const length = Math.hypot(dx, dy);
  if (!Number.isFinite(length)) return [];
  if (length < 4) return [first];
  const horizontal = (side: TableBorderEdgeSide) =>
    side === "top" || side === "bottom";
  if ((Math.abs(dx) >= Math.abs(dy)) !== horizontal(first.side)) return [];
  const edges: TableBorderEdgeHit[] = [];
  const seen = new Set<string>();
  const steps = Math.ceil(length / 4);
  for (let step = 0; step <= steps; step += 1) {
    const hit = hitTestTableBorderEdge(
      boxes,
      { x: from.x + (dx * step) / steps, y: from.y + (dy * step) / steps },
      tolerance,
    );
    if (!hit || horizontal(hit.side) !== horizontal(first.side)) continue;
    const key = tableBorderEdgeKey(hit);
    if (seen.has(key)) continue;
    seen.add(key);
    edges.push(hit);
  }
  return edges;
}

/** The cell-style key an edge writes (`top` ⇒ `borderTop`). */
export function tableBorderEdgeStyleKey(
  side: TableBorderEdgeSide,
): keyof CellStyleInput {
  return EDGE_STYLE_KEY[side];
}

interface EdgeGeometry {
  /** Fixed coordinate of the edge line (y for horizontal, x for vertical). */
  position: number;
  /** Span start along the edge. */
  from: number;
  /** Span end along the edge. */
  to: number;
  horizontal: boolean;
}

function edgeGeometry(
  box: TableBorderCellBox,
  side: TableBorderEdgeSide,
): EdgeGeometry {
  const right = box.left + box.width;
  const bottom = box.top + box.height;
  switch (side) {
    case "top":
      return { position: box.top, from: box.left, to: right, horizontal: true };
    case "bottom":
      return { position: bottom, from: box.left, to: right, horizontal: true };
    case "left":
      return { position: box.left, from: box.top, to: bottom, horizontal: false };
    case "right":
      return { position: right, from: box.top, to: bottom, horizontal: false };
  }
}

/**
 * Nearest cell edge to a pointer, or `null` when the pointer is not within
 * `tolerance` of any edge. Ties resolve to the FIRST box in document order so a
 * shared internal line always reports the same owning cell — the drag path then
 * de-duplicates cleanly even when the pointer wobbles across the line.
 */
export function hitTestTableBorderEdge(
  boxes: readonly TableBorderCellBox[],
  point: { x: number; y: number },
  tolerance: number = TABLE_BORDER_PEN_TOLERANCE,
): TableBorderEdgeHit | null {
  let best: TableBorderEdgeHit | null = null;
  const half = TABLE_BORDER_PEN_MARKER_THICKNESS / 2;
  for (const box of boxes) {
    if (!(box.width > 0) || !(box.height > 0)) continue;
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const geo = edgeGeometry(box, side);
      const across = geo.horizontal ? point.y : point.x;
      const along = geo.horizontal ? point.x : point.y;
      const distance = Math.abs(across - geo.position);
      // A non-finite pointer coordinate makes every comparison false, so a NaN
      // would sail past the tolerance check AND past the "closer than best" test
      // and claim an arbitrary edge. Require a real distance.
      if (!Number.isFinite(distance) || distance > tolerance) continue;
      // Stay inside the edge's own span (a small overhang keeps corners
      // reachable) so a pointer far off the end of a short cell does not claim
      // that cell's edge.
      if (along < geo.from - tolerance || along > geo.to + tolerance) continue;
      if (best !== null && distance >= best.distance) continue;
      best = {
        row: box.row,
        col: box.col,
        side,
        distance,
        left: geo.horizontal ? geo.from : geo.position - half,
        top: geo.horizontal ? geo.position - half : geo.from,
        width: geo.horizontal
          ? geo.to - geo.from
          : TABLE_BORDER_PEN_MARKER_THICKNESS,
        height: geo.horizontal
          ? TABLE_BORDER_PEN_MARKER_THICKNESS
          : geo.to - geo.from,
      };
    }
  }
  return best;
}

/**
 * A pencil stroke through a cell's INTERIOR — WPS's other pencil gesture. Drag
 * across a cell and the pencil draws a new gridline there, splitting it. The
 * axis names what the new line ADDS: a horizontal stroke divides the cell into
 * two rows, a vertical stroke into two columns.
 */
export interface TableBorderPenSplit {
  /** Source-array coordinates of the crossed cell. */
  row: number;
  col: number;
  axis: "row" | "column";
  /** Overlay-local preview bar for the line about to be drawn. */
  left: number;
  top: number;
  width: number;
  height: number;
}

/** Minimum drag (px) before an interior stroke counts as a new gridline. */
export const TABLE_BORDER_PEN_SPLIT_MIN_TRAVEL = 12;

/** The cell a point falls inside, ignoring its edges. */
export function hitTestTableCellInterior(
  boxes: readonly TableBorderCellBox[],
  point: { x: number; y: number },
): TableBorderCellBox | null {
  for (const box of boxes) {
    if (!(box.width > 0) || !(box.height > 0)) continue;
    if (
      point.x >= box.left &&
      point.x <= box.left + box.width &&
      point.y >= box.top &&
      point.y <= box.top + box.height
    ) {
      return box;
    }
  }
  return null;
}

/**
 * Resolve an interior drag into the gridline it draws, or `null` when the
 * pointer barely moved (a click inside a cell must not restructure a table).
 * The dominant direction wins, so a slightly wobbly horizontal stroke still
 * reads as horizontal — the same forgiveness WPS's pencil has.
 */
export function resolveTableBorderPenSplit(
  box: TableBorderCellBox,
  from: { x: number; y: number },
  to: { x: number; y: number },
  minTravel: number = TABLE_BORDER_PEN_SPLIT_MIN_TRAVEL,
): TableBorderPenSplit | null {
  const dx = Math.abs(to.x - from.x);
  const dy = Math.abs(to.y - from.y);
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return null;
  if (Math.max(dx, dy) < minTravel) return null;
  const half = TABLE_BORDER_PEN_MARKER_THICKNESS / 2;
  if (dx >= dy) {
    const y = (from.y + to.y) / 2;
    return {
      row: box.row,
      col: box.col,
      axis: "row",
      left: box.left,
      top: y - half,
      width: box.width,
      height: TABLE_BORDER_PEN_MARKER_THICKNESS,
    };
  }
  const x = (from.x + to.x) / 2;
  return {
    row: box.row,
    col: box.col,
    axis: "column",
    left: x - half,
    top: box.top,
    width: TABLE_BORDER_PEN_MARKER_THICKNESS,
    height: box.height,
  };
}

/**
 * Every cell edge that describes the SAME physical line: the requested edge
 * plus the facing edge of each neighbour across it. A merged cell faces several
 * neighbours along one side, so all of them are returned. Falls back to the
 * requested edge alone when the table has overlapping spans (an inconsistent
 * grid) — the pen then still writes one editable side rather than nothing.
 */
export function tableBorderEdgeSiblings(
  rows: readonly (readonly MergeableCell[])[],
  ref: TableBorderEdgeRef,
): TableBorderEdgeRef[] {
  const result: TableBorderEdgeRef[] = [ref];
  const seen = new Set<string>([tableBorderEdgeKey(ref)]);
  if (!rows[ref.row]?.[ref.col]) return result;
  const grid = buildOccupancyGrid(rows as MergeableCell[][]);
  if (!grid) return result;
  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  for (let r = 0; r < grid.length; r += 1) {
    for (let c = 0; c < grid[r].length; c += 1) {
      const occupant = grid[r][c];
      if (!occupant) continue;
      if (occupant.anchorRow !== ref.row || occupant.cellIndex !== ref.col) {
        continue;
      }
      minRow = Math.min(minRow, r);
      maxRow = Math.max(maxRow, r);
      minCol = Math.min(minCol, c);
      maxCol = Math.max(maxCol, c);
    }
  }
  if (!Number.isFinite(minRow) || !Number.isFinite(minCol)) return result;
  const neighbours: Array<{ r: number; c: number }> = [];
  if (ref.side === "top") {
    for (let c = minCol; c <= maxCol; c += 1) neighbours.push({ r: minRow - 1, c });
  } else if (ref.side === "bottom") {
    for (let c = minCol; c <= maxCol; c += 1) neighbours.push({ r: maxRow + 1, c });
  } else if (ref.side === "left") {
    for (let r = minRow; r <= maxRow; r += 1) neighbours.push({ r, c: minCol - 1 });
  } else {
    for (let r = minRow; r <= maxRow; r += 1) neighbours.push({ r, c: maxCol + 1 });
  }
  for (const { r, c } of neighbours) {
    const occupant = grid[r]?.[c];
    if (!occupant) continue;
    const sibling: TableBorderEdgeRef = {
      row: occupant.anchorRow,
      col: occupant.cellIndex,
      side: OPPOSITE_SIDE[ref.side],
    };
    const key = tableBorderEdgeKey(sibling);
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(sibling);
  }
  return result;
}

/** The logical rectangle an eraser stroke merges (inclusive on all sides). */
export interface TableBorderEraseMerge {
  row1: number;
  col1: number;
  row2: number;
  col2: number;
}

/** The logical rectangle a source-array cell covers, or `null`. */
function logicalRect(
  grid: readonly (readonly ({ anchorRow: number; cellIndex: number } | null)[])[],
  row: number,
  col: number,
): TableBorderEraseMerge | null {
  let row1 = Number.POSITIVE_INFINITY;
  let row2 = Number.NEGATIVE_INFINITY;
  let col1 = Number.POSITIVE_INFINITY;
  let col2 = Number.NEGATIVE_INFINITY;
  for (let r = 0; r < grid.length; r += 1) {
    for (let c = 0; c < grid[r].length; c += 1) {
      const occupant = grid[r][c];
      if (!occupant || occupant.anchorRow !== row || occupant.cellIndex !== col) {
        continue;
      }
      row1 = Math.min(row1, r);
      row2 = Math.max(row2, r);
      col1 = Math.min(col1, c);
      col2 = Math.max(col2, c);
    }
  }
  return Number.isFinite(row1) && Number.isFinite(col1)
    ? { row1, col1, row2, col2 }
    : null;
}

/**
 * WPS's eraser removes the RULE BETWEEN cells, which merges them. Resolve which
 * rectangles a stroke collapses: cells joined by an ERASED rule form one group,
 * and each group merges on its own.
 *
 * Grouping matters. Dragging down a table's middle vertical erases that rule on
 * every row, so each row's pair merges — the result is a narrower table, not one
 * giant cell. A bounding box over all the pairs would swallow the horizontal
 * rules the stroke never touched.
 *
 * Returns an empty list when the stroke only touched the table's outer frame:
 * there is nothing on the far side to merge with, so the caller suppresses those
 * borders instead.
 */
export function resolveTableBorderEraseMerge(
  rows: readonly (readonly MergeableCell[])[],
  edges: readonly TableBorderEdgeRef[],
): TableBorderEraseMerge[] {
  if (edges.length === 0) return [];
  const grid = buildOccupancyGrid(rows as MergeableCell[][]);
  if (!grid) return [];
  // Union-find over source cells, joined only where a rule was erased.
  const parent = new Map<string, string>();
  const find = (key: string): string => {
    let root = key;
    while (parent.get(root) !== undefined && parent.get(root) !== root) {
      root = parent.get(root)!;
    }
    return root;
  };
  const union = (a: string, b: string): void => {
    if (!parent.has(a)) parent.set(a, a);
    if (!parent.has(b)) parent.set(b, b);
    const rootA = find(a);
    const rootB = find(b);
    if (rootA !== rootB) parent.set(rootA, rootB);
  };
  for (const edge of edges) {
    const siblings = tableBorderEdgeSiblings(rows, edge);
    // A single-entry result means the edge faces nothing: an outer frame line.
    if (siblings.length < 2) continue;
    const anchor = `${siblings[0].row}:${siblings[0].col}`;
    for (const ref of siblings.slice(1)) {
      union(anchor, `${ref.row}:${ref.col}`);
    }
  }
  const groups = new Map<string, TableBorderEraseMerge>();
  for (const key of parent.keys()) {
    const [row, col] = key.split(":").map((part) => Number.parseInt(part, 10));
    const rect = logicalRect(grid, row, col);
    if (!rect) continue;
    const root = find(key);
    const current = groups.get(root);
    groups.set(
      root,
      current
        ? {
            row1: Math.min(current.row1, rect.row1),
            col1: Math.min(current.col1, rect.col1),
            row2: Math.max(current.row2, rect.row2),
            col2: Math.max(current.col2, rect.col2),
          }
        : rect,
    );
  }
  return [...groups.values()].filter(
    (rect) => rect.row1 !== rect.row2 || rect.col1 !== rect.col2,
  );
}

/** The edges of a stroke that face nothing — the table's outer frame. */
export function outerTableBorderEdges(
  rows: readonly (readonly MergeableCell[])[],
  edges: readonly TableBorderEdgeRef[],
): TableBorderEdgeRef[] {
  return edges.filter((edge) => tableBorderEdgeSiblings(rows, edge).length < 2);
}

function sameBorder(
  existing: TableCellBorder | null,
  next: TableCellBorder | null,
): boolean {
  if (existing === null || next === null) return existing === next;
  return (
    existing.style === next.style &&
    existing.color === next.color &&
    existing.width === next.width &&
    existing.compound === next.compound
  );
}

/**
 * Apply one pen GESTURE — every edge the pointer touched — to the table rows,
 * returning new rows. Returns the ORIGINAL rows reference when the gesture
 * changes nothing (re-erasing an already-erased edge), so the caller can skip
 * the commit and keep undo history free of no-op entries.
 */
export function applyTableBorderPenStroke<C extends MergeableCell>(
  rows: readonly C[][],
  edges: readonly TableBorderEdgeRef[],
  border: TableCellBorder,
): C[][] {
  if (edges.length === 0) return rows as C[][];
  const normalized = normalizeCellStyle({ borderTop: border }).borderTop;
  if (!normalized) return rows as C[][];
  const patches = new Map<string, Partial<CellStyleInput>>();
  for (const edge of edges) {
    for (const sibling of tableBorderEdgeSiblings(rows, edge)) {
      const key = `${sibling.row}:${sibling.col}`;
      const patch = patches.get(key) ?? {};
      (patch as Record<string, unknown>)[EDGE_STYLE_KEY[sibling.side]] =
        normalized;
      patches.set(key, patch);
    }
  }
  let changed = false;
  const next = rows.map((row, rowIndex) =>
    row.map((cell, cellIndex) => {
      const patch = patches.get(`${rowIndex}:${cellIndex}`);
      if (!patch) return cell;
      const resolved = normalizeCellStyle(cell as CellStyleInput);
      const redundant = Object.keys(patch).every((key) =>
        sameBorder(
          (resolved as unknown as Record<string, TableCellBorder | null>)[key] ??
            null,
          normalized,
        ),
      );
      if (redundant) return cell;
      changed = true;
      const merged: CellStyleInput = {
        ...(cell as Record<string, unknown>),
        ...patch,
      };
      const base = { ...(cell as Record<string, unknown>) };
      for (const key of CELL_STYLE_KEYS) delete base[key];
      return {
        ...base,
        text: cell.text,
        ...persistedCellStyle(merged),
      } as C;
    }),
  );
  return changed ? next : (rows as C[][]);
}

/**
 * The border a pen gesture writes. `erase` keeps the pen colour and width so an
 * erased edge can be redrawn identically, but sets `style: "none"` — which the
 * PPTX projection emits as DrawingML `a:noFill`, i.e. an explicitly suppressed
 * native line rather than a deleted one.
 */
export function tableBorderPenBorder(
  mode: TableBorderPenMode,
  pen: TableCellBorder,
): TableCellBorder {
  if (mode === "erase") return { ...pen, style: "none" };
  return pen.style === "none" ? { ...pen, style: "solid" } : pen;
}
