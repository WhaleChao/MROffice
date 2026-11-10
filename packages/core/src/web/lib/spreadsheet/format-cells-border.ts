/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Cells > Border as Excel models it (XLS-A060): a selection is one
 * block of cells, and the dialog's eight border buttons name LINES of that
 * block, not edges of every cell.
 *
 *   top / bottom / left / right   the block's perimeter: the top edge of its
 *                                 first row, the right edge of its last
 *                                 column, ...
 *   insideHorizontal / Vertical   every boundary between two rows (columns)
 *                                 of the block; offered only when the block
 *                                 has more than one row (column)
 *   diagonalUp / diagonalDown     the diagonal of every cell of the block
 *
 * (Microsoft Learn, "Control settings in the Format Cells dialog box";
 * Microsoft Support, "Apply or remove cell borders on a worksheet": "A
 * selected range of cells is formatted as a single block of cells".)
 *
 * Both cells of an interior boundary carry it, as the Inside preset and the
 * Home Borders gallery write it, so a later edit on either side still finds
 * its own edge. SpreadsheetML keeps ONE diagonal per cell (a style, a colour
 * and diagonalUp / diagonalDown flags), so the two diagonal lines share their
 * style and colour: the last one drawn restyles both.
 *
 * Pure: the summary reads the selection, the per-cell patch says what one
 * cell's edges become. The worksheet commit (lib/spreadsheet/cell-buttons)
 * applies it.
 */

import {
  BORDER_EDGES,
  BORDER_LINE_STYLES,
  type BorderEdge,
  type BorderLineStyle,
  type CellBorderEdges,
} from "./cell-format";
import { normalizeHexColor } from "./color-palette";
import type { SpreadsheetThemeColor } from "./theme-colors";

export type FormatCellsBorderLine =
  | "top"
  | "insideHorizontal"
  | "bottom"
  | "diagonalUp"
  | "left"
  | "insideVertical"
  | "right"
  | "diagonalDown";

/** Excel's button order: the column beside the preview, then the row under it. */
export const FORMAT_CELLS_BORDER_LINES: readonly FormatCellsBorderLine[] = [
  "top",
  "insideHorizontal",
  "bottom",
  "diagonalUp",
  "left",
  "insideVertical",
  "right",
  "diagonalDown",
] as const;

/** A drawn line: a style and an explicit colour (absent = Automatic). */
export interface FormatCellsBorderLineValue {
  readonly style: BorderLineStyle;
  readonly color?: string;
}

/** What one line of the block shows: one agreed value, none, or mixed. */
export interface FormatCellsBorderLineState {
  readonly value?: FormatCellsBorderLineValue;
  readonly mixed: boolean;
}

export type FormatCellsBorderSummary = Readonly<
  Record<FormatCellsBorderLine, FormatCellsBorderLineState>
>;

/** One line set from the dialog: a pen (style + colour) or removal (null). */
export interface FormatCellsBorderLinePatch {
  readonly style: BorderLineStyle | null;
  /** null = Automatic. */
  readonly color?: string | null;
  readonly themeColor?: SpreadsheetThemeColor;
}

export type FormatCellsBorderLinesPatch = Partial<
  Record<FormatCellsBorderLine, FormatCellsBorderLinePatch>
>;

/** What one cell's edge becomes. `style: null` removes it; an absent style
 * keeps the edge's own (a diagonal losing one of its two directions). */
export interface CellBorderEdgePatch {
  style?: BorderLineStyle | null;
  color?: string | null;
  themeColor?: SpreadsheetThemeColor;
  /** Diagonal only: the directions the cell's diagonal is drawn in. */
  up?: boolean;
  down?: boolean;
}

export type CellBorderEdgesPatch = Partial<Record<BorderEdge, CellBorderEdgePatch>>;

export interface FormatCellsBlock {
  readonly r1: number;
  readonly r2: number;
  readonly c1: number;
  readonly c2: number;
}

/** The block line a cell's box edge belongs to. */
export function formatCellsBorderLineOfEdge(
  edge: Exclude<BorderEdge, "diagonal">,
  row: number,
  col: number,
  block: FormatCellsBlock,
): FormatCellsBorderLine {
  switch (edge) {
    case "top":
      return row === block.r1 ? "top" : "insideHorizontal";
    case "bottom":
      return row === block.r2 ? "bottom" : "insideHorizontal";
    case "left":
      return col === block.c1 ? "left" : "insideVertical";
    case "right":
      return col === block.c2 ? "right" : "insideVertical";
  }
}

const BOX_EDGES = ["top", "right", "bottom", "left"] as const;

type EdgeReading =
  | { readonly malformed: true }
  | {
      readonly malformed: false;
      readonly edges: Partial<
        Record<BorderEdge, { style: BorderLineStyle; color?: string; up?: boolean; down?: boolean }>
      >;
    };

/** A stored style's rich edges, or malformed when any part is not a valid
 * edge (the dialog then has no border authority over the selection). */
function readEdges(style: unknown): EdgeReading {
  const raw = (style as { borderEdges?: unknown } | null | undefined)?.borderEdges;
  if (raw === undefined) return { malformed: false, edges: {} };
  if (raw === null || typeof raw !== "object" || Array.isArray(raw)) {
    return { malformed: true };
  }
  const map = raw as Record<string, unknown>;
  const edges: Partial<
    Record<BorderEdge, { style: BorderLineStyle; color?: string; up?: boolean; down?: boolean }>
  > = {};
  for (const edge of BORDER_EDGES) {
    if (!Object.prototype.hasOwnProperty.call(map, edge)) continue;
    const candidate = map[edge];
    if (candidate === null || typeof candidate !== "object" || Array.isArray(candidate)) {
      return { malformed: true };
    }
    const entry = candidate as Record<string, unknown>;
    if (
      typeof entry.style !== "string" ||
      !BORDER_LINE_STYLES.includes(entry.style as BorderLineStyle)
    ) {
      return { malformed: true };
    }
    let color: string | undefined;
    if (entry.color !== undefined) {
      const normalized = normalizeHexColor(entry.color);
      if (!normalized) return { malformed: true };
      color = normalized;
    }
    for (const flag of ["up", "down"] as const) {
      if (entry[flag] !== undefined && typeof entry[flag] !== "boolean") {
        return { malformed: true };
      }
    }
    edges[edge] = {
      style: entry.style as BorderLineStyle,
      ...(color ? { color } : {}),
      ...(entry.up === true ? { up: true } : {}),
      ...(entry.down === true ? { down: true } : {}),
    };
  }
  return { malformed: false, edges };
}

/** A diagonal's two directions (neither flag set reads as down, the
 * split-header form, as the grid draws it). */
function diagonalDirections(
  diagonal: { up?: boolean; down?: boolean } | undefined,
): { up: boolean; down: boolean } {
  if (!diagonal) return { up: false, down: false };
  return { up: diagonal.up === true, down: diagonal.down === true || diagonal.up !== true };
}

const NONE_KEY = "none";
const lineKey = (value: FormatCellsBorderLineValue | undefined) =>
  value ? `${value.style}|${value.color ?? ""}` : NONE_KEY;

/**
 * Every line of the block as the dialog shows it. Null when a stored edge is
 * malformed. A selection larger than `denseLimit` cells reads only its stored
 * cells, plus "no line" for the positions no stored cell covers, so a whole
 * column never walks a million empty coordinates.
 */
export function formatCellsBorderSummary(
  cells: Readonly<Record<string, { style?: unknown } | undefined>>,
  block: FormatCellsBlock,
  denseLimit = 20_000,
): FormatCellsBorderSummary | null {
  const rows = block.r2 - block.r1 + 1;
  const cols = block.c2 - block.c1 + 1;
  const seen = Object.fromEntries(
    FORMAT_CELLS_BORDER_LINES.map((line) => [
      line,
      new Map<string, FormatCellsBorderLineValue | undefined>(),
    ]),
  ) as Record<FormatCellsBorderLine, Map<string, FormatCellsBorderLineValue | undefined>>;
  const observed = Object.fromEntries(
    FORMAT_CELLS_BORDER_LINES.map((line) => [line, 0]),
  ) as Record<FormatCellsBorderLine, number>;
  const add = (line: FormatCellsBorderLine, value: FormatCellsBorderLineValue | undefined) => {
    seen[line].set(lineKey(value), value);
    observed[line] += 1;
  };
  const observe = (row: number, col: number, style: unknown): boolean => {
    const reading = readEdges(style);
    if (reading.malformed) return false;
    for (const edge of BOX_EDGES) {
      const own = reading.edges[edge];
      add(
        formatCellsBorderLineOfEdge(edge, row, col, block),
        own ? { style: own.style, ...(own.color ? { color: own.color } : {}) } : undefined,
      );
    }
    const diagonal = reading.edges.diagonal;
    const directions = diagonalDirections(diagonal);
    const drawn = diagonal
      ? { style: diagonal.style, ...(diagonal.color ? { color: diagonal.color } : {}) }
      : undefined;
    add("diagonalUp", directions.up ? drawn : undefined);
    add("diagonalDown", directions.down ? drawn : undefined);
    return true;
  };

  if (rows * cols <= denseLimit) {
    for (let row = block.r1; row <= block.r2; row += 1) {
      for (let col = block.c1; col <= block.c2; col += 1) {
        if (!observe(row, col, cells[`${row}:${col}`]?.style)) return null;
      }
    }
  } else {
    for (const key in cells) {
      if (!Object.prototype.hasOwnProperty.call(cells, key)) continue;
      const separator = key.indexOf(":");
      if (separator <= 0) continue;
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (
        !Number.isSafeInteger(row) ||
        !Number.isSafeInteger(col) ||
        row < block.r1 ||
        row > block.r2 ||
        col < block.c1 ||
        col > block.c2
      ) {
        continue;
      }
      if (!observe(row, col, cells[key]?.style)) return null;
    }
  }
  // Positions no stored cell covers have no line.
  const expected: Record<FormatCellsBorderLine, number> = {
    top: cols,
    bottom: cols,
    left: rows,
    right: rows,
    insideHorizontal: 2 * (rows - 1) * cols,
    insideVertical: 2 * (cols - 1) * rows,
    diagonalUp: rows * cols,
    diagonalDown: rows * cols,
  };
  const summary = {} as Record<FormatCellsBorderLine, FormatCellsBorderLineState>;
  for (const line of FORMAT_CELLS_BORDER_LINES) {
    if (observed[line] < expected[line]) seen[line].set(NONE_KEY, undefined);
    const values = [...seen[line].values()];
    summary[line] =
      values.length === 1
        ? Object.freeze({ ...(values[0] ? { value: values[0] } : {}), mixed: false })
        : Object.freeze({ mixed: values.length > 1 });
  }
  return Object.freeze(summary);
}

/** One cell's own edges read as the dialog's lines (a single-cell block). */
export function formatCellsBorderSummaryOfEdges(
  edges: CellBorderEdges | undefined,
): FormatCellsBorderSummary {
  return (
    formatCellsBorderSummary({ "0:0": { style: { borderEdges: edges } } }, { r1: 0, r2: 0, c1: 0, c2: 0 }) ??
    formatCellsBorderSummary({}, { r1: 0, r2: 0, c1: 0, c2: 0 })!
  );
}

/**
 * The edge patch the dialog's line patch makes on one cell of the block.
 * `current` is that cell's own edges, which a diagonal needs: drawing one
 * direction keeps the other the cell already has.
 */
export function formatCellsBorderPatchForCell(
  row: number,
  col: number,
  block: FormatCellsBlock,
  lines: FormatCellsBorderLinesPatch,
  current: CellBorderEdges | undefined,
): CellBorderEdgesPatch {
  const patch: CellBorderEdgesPatch = {};
  const has = (line: FormatCellsBorderLine) =>
    Object.prototype.hasOwnProperty.call(lines, line) && lines[line] !== undefined;
  for (const edge of BOX_EDGES) {
    const line = formatCellsBorderLineOfEdge(edge, row, col, block);
    if (!has(line)) continue;
    const pen = lines[line]!;
    patch[edge] =
      pen.style === null
        ? { style: null }
        : {
            style: pen.style,
            color: pen.color ?? null,
            ...(pen.themeColor ? { themeColor: pen.themeColor } : {}),
          };
  }
  const upSet = has("diagonalUp");
  const downSet = has("diagonalDown");
  if (upSet || downSet) {
    const own = diagonalDirections(current?.diagonal);
    const up = upSet ? lines.diagonalUp!.style !== null : own.up;
    const down = downSet ? lines.diagonalDown!.style !== null : own.down;
    if (!up && !down) {
      patch.diagonal = { style: null };
    } else {
      const pen =
        upSet && lines.diagonalUp!.style !== null
          ? lines.diagonalUp!
          : downSet && lines.diagonalDown!.style !== null
            ? lines.diagonalDown!
            : undefined;
      patch.diagonal = {
        ...(pen
          ? {
              style: pen.style,
              color: pen.color ?? null,
              ...(pen.themeColor ? { themeColor: pen.themeColor } : {}),
            }
          : {}),
        up,
        down,
      };
    }
  }
  return patch;
}
