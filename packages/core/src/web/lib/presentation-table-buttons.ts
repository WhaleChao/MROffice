/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's table buttons as pure transforms of a table block:
 * Insert Above / Below / Left / Right, Delete Rows / Columns, cell shading,
 * alignment, margins and borders over a cell selection, Table Style Options,
 * Table Styles and a cell's text. The ribbon (`PresentationWindow`) resolves
 * the live cell selection; a headless editor command passes the cells it
 * names. Both commit what these return.
 */

import { CSS_PX_PER_POINT } from "./presentation/text-units";
import { tableCellListRows } from "./presentation/table-cell-lists";
import { normalizeOfficeTextLayout, officeTextLayoutFollowingText } from "./office-text-layout";
import type {
  PresentationTableBlock,
  PresentationTableCell,
} from "../components/console/editor/native-document/documentCommands";
import {
  CELL_CHARACTER_FIELDS,
  cellCharacterRuns,
  clearCellCharacterFormat,
  formatCellTextRange,
  type CellTextRange,
} from "./presentation-cell-rich-text";
import type { RunStyle, TextRun } from "./presentation-runs";
import {
  buildOccupancyGrid,
  deleteTableColumn,
  deleteTableRow,
  insertTableColumn,
  insertTableRow,
  tableLogicalWidth,
} from "./presentation-table-merge";
import {
  tableCellIntersectsSelection,
  type TableAxisSelectionRange,
} from "./presentation-table-selection";
import {
  deleteAxisSize,
  equalizedAxisSizeRange,
  equalizedAxisSizes,
  insertAxisSize,
  materializedAxisSizeAt,
  normalizeAxisSizes,
  tableFrameAfterAxisChange,
} from "./presentation-table-size";
import {
  normalizePresentationTableBackground,
  withPresentationTableEffect,
  type PresentationTableEffectKind,
  type PresentationTableEffects,
} from "./presentation-table-effects";
import {
  applyTableStylePreset,
  CELL_STYLE_KEYS,
  clearNativeTableStyle,
  normalizeCellStyle,
  normalizeTableStyleId,
  normalizeTableStyleLook,
  persistedCellStyle,
  TABLE_STYLE_PRESETS,
  tableStyleLookCellPaint,
  type CellStyleInput,
  type ResolvedCellStyle,
  type TableBorderSelectionPosition,
  type TableStyleLook,
  type TableStylePresetId,
} from "./presentation-table-style";


export type PresentationTableStyleSelection =
  | { readonly kind: "preset"; readonly presetId: TableStylePresetId }
  | { readonly kind: "look"; readonly look: TableStyleLook }
  | { readonly kind: "clear" };

/** A logical cell rectangle, 0-based, as the table selection states it. */
export interface PresentationTableCellRectangle {
  row1: number;
  row2: number;
  col1: number;
  col2: number;
}

function lookIsCarriedByTable(
  table: PresentationTableBlock,
  styleId: string | null,
): boolean {
  return (
    normalizeTableStyleLook(table.tableStyleLook, table.tableStyleId)
      ?.styleId === styleId
  );
}

/** The same colour, either way it is spelt. */
function sameHex(a: unknown, b: unknown): boolean {
  return typeof a === "string" && typeof b === "string" && a.trim().toLowerCase() === b.trim().toLowerCase();
}

/** Two edges that draw the same line. */
function sameEdge(a: unknown, b: unknown): boolean {
  const left = a as { style?: unknown; color?: unknown; width?: unknown } | null | undefined;
  const right = b as { style?: unknown; color?: unknown; width?: unknown } | null | undefined;
  if (!left || !right) return false;
  if ((left.style === "none") !== (right.style === "none")) return false;
  if (left.style === "none") return true;
  return (
    left.style === right.style &&
    sameHex(left.color, right.color) &&
    Math.abs(Number(left.width ?? 0) - Number(right.width ?? 0)) < 0.01
  );
}










function cellsWithoutOldStylePaint(table: PresentationTableBlock): PresentationTableBlock {
  const look = normalizeTableStyleLook(table.tableStyleLook, table.tableStyleId);
  if (!look) return table;
  const flags = table as unknown as Parameters<typeof tableStyleLookCellPaint>[1];
  const rowCount = table.rows.length;
  let changed = false;
  const rows = table.rows.map((row, rowIndex) => {
    let rowChanged = false;
    const next = row.map((cell, colIndex) => {
      const paint = tableStyleLookCellPaint(look, flags, rowIndex, rowCount, colIndex, row.length);
      if (!paint) return cell;
      const record = cell as unknown as Record<string, unknown>;
      const drop: string[] = [];
      if (paint.color !== null && sameHex(record.color, paint.color)) drop.push("color");
      if ("bold" in record && record.bold === paint.bold) drop.push("bold");
      if ("italic" in record && record.italic === paint.italic) drop.push("italic");
      for (const side of ["top", "bottom", "left", "right"] as const) {
        const key = `border${side[0].toUpperCase()}${side.slice(1)}`;
        const edge = paint.edges[side] ?? { style: "none", color: "#000000", width: 1 };
        if (key in record && sameEdge(record[key], edge)) drop.push(key);
      }
      if (!drop.length) return cell;
      rowChanged = true;
      const stripped = { ...record };
      for (const key of drop) delete stripped[key];
      return stripped as unknown as typeof cell;
    });
    if (!rowChanged) return row;
    changed = true;
    return next;
  });
  return changed ? { ...table, rows } : table;
}

/**
 * An explicit style application replaces cell shading so the new table look
 * is visible. Opening/rendering never clears a fill, and subsequent cell shading
 * still overrides the style. Text, borders, merges and source custody survive.
 * Returning the original object
 * is the no-op signal consumed by the editor's one-save/one-Undo boundary.
 */
export function applyPresentationTableStyleSelection(
  table: PresentationTableBlock,
  selection: PresentationTableStyleSelection,
): PresentationTableBlock {
  if (selection.kind !== "clear") {
    let changed = false;
    // Shading > Picture is the cell's shading too (`fillImage`).
    const shaded = (cell: PresentationTableBlock["rows"][number][number]) =>
      cell.fill !== undefined || (cell as { fillImage?: unknown }).fillImage !== undefined;
    const rows = table.rows.map((row) => {
      if (!row.some(shaded)) return row;
      changed = true;
      return row.map((cell) => {
        if (!shaded(cell)) return cell;
        const next = { ...cell } as typeof cell & { fillImage?: unknown };
        delete next.fill;
        delete next.fillImage;
        return next;
      });
    });
    if (changed) table = { ...table, rows };
  }
  if (selection.kind === "look") {
    const styleId = selection.look.styleId;
    if (
      normalizeTableStyleId(table.tableStyleId) === styleId &&
      lookIsCarriedByTable(table, styleId)
    ) {
      return table;
    }
    // A table style draws its own lines (`wholeTbl` borders): the editor's
    // plain grey grid (`borders`) stays off under it, or "No Style, No
    // Grid" kept a full grey grid on the slide, in the show and in the
    
    const styled: PresentationTableBlock = {
      ...cellsWithoutOldStylePaint(table),
      tableStyleId: styleId,
      tableStyleLook: selection.look,
    };
    delete (styled as { borders?: unknown }).borders;
    return styled;
  }

  if (selection.kind === "clear") {
    const cleared = clearNativeTableStyle(table) as PresentationTableBlock;
    const styleId = normalizeTableStyleId(cleared.tableStyleId);
    if (
      normalizeTableStyleId(table.tableStyleId) === styleId &&
      !table.tableStyleLook
    ) {
      return table;
    }
    const next = { ...cleared };
    delete next.tableStyleLook;
    return next;
  }

  const applied = applyTableStylePreset(
    table,
    selection.presetId,
  ) as PresentationTableBlock;
  const styleId = normalizeTableStyleId(applied.tableStyleId);
  const preset = TABLE_STYLE_PRESETS.find(
    (candidate) => candidate.id === selection.presetId,
  );
  const flagsAlreadyApplied =
    (!preset?.headerRow || table.headerRow === true) &&
    (!preset?.bandedRows || table.bandedRows === true);
  const keepsLook = lookIsCarriedByTable(table, styleId);
  if (
    normalizeTableStyleId(table.tableStyleId) === styleId &&
    flagsAlreadyApplied &&
    (!table.tableStyleLook || keepsLook)
  ) {
    return table;
  }
  if (keepsLook || !table.tableStyleLook) return applied;
  const next = { ...applied };
  delete next.tableStyleLook;
  return next;
}


/** One row/column insert or delete, as the table stood just before it. */
export interface PresentationTableAxisEdit {
  axis: "row" | "column";
  kind: "insert" | "delete";
  index: number;
  count: number;
  /** An insert's source row (column), in the table before it: the one the
   * new cells take their formatting from (the save copies its end marks'
   * language, as PowerPoint does). */
  from?: number;
}

/** The longest `nativeAxisEdits` list a table keeps (the API's
 * `_MAX_STATED_AXIS_EDITS`, table_projection.py). */
export const MAX_TABLE_AXIS_EDITS = 64;









function withAxisEdit(
  block: PresentationTableBlock,
  edit: PresentationTableAxisEdit,
): PresentationTableBlock {
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return block;
  const previous = Array.isArray(block.nativeAxisEdits)
    ? (block.nativeAxisEdits as PresentationTableAxisEdit[])
    : [];
  const edits = [...previous, edit];
  if (edits.length > MAX_TABLE_AXIS_EDITS) {
    const next = { ...block };
    delete next.nativeAxisEdits;
    return next;
  }
  return { ...block, nativeAxisEdits: edits };
}

/** Insert Above / Below: one row beside the range, sized like its neighbour. Null when it cannot. */
export function tableWithInsertedRow(
  block: PresentationTableBlock,
  range: TableAxisSelectionRange,
  placement: "before" | "after",
): PresentationTableBlock | null {
  const atRow = placement === "before" ? range.start : range.end + 1;
  const sourceRow = placement === "before" ? range.start : range.end;
  // PowerPoint for the web inserts exactly one row per command, even when
  // the live cell selection spans multiple rows.
  // The new row takes the formatting of the row it is inserted from.
  const rows = insertTableRow(
    block.rows,
    atRow,
    sourceRow,
  ) as PresentationTableCell[][];
  if (rows === block.rows) return null;
  const rowHeights = insertAxisSize(
    block.rowHeights,
    block.rows.length,
    atRow,
    sourceRow,
  );
  return withFrameAfterAxisChange(
    block,
    withAxisEdit({ ...block, rows, rowHeights }, { axis: "row", kind: "insert", index: atRow, count: 1, from: sourceRow }),
    "row",
  );
}

/** Insert Left / Right: one column beside the range, sized like its neighbour. Null when it cannot. */
export function tableWithInsertedColumn(
  block: PresentationTableBlock,
  range: TableAxisSelectionRange,
  placement: "before" | "after",
): PresentationTableBlock | null {
  const width = tableLogicalWidth(block.rows);
  const atCol = placement === "before" ? range.start : range.end + 1;
  const sourceColumn = placement === "before" ? range.start : range.end;
  // The Web command is single-shot here as well; desktop-style multi-axis
  // insertion remains outside the visible Web inventory.
  // The new column takes the formatting of the column it is inserted from.
  const rows = insertTableColumn(
    block.rows,
    atCol,
    sourceColumn,
  ) as PresentationTableCell[][];
  if (rows === block.rows) return null;
  const colWidths = insertAxisSize(
    block.colWidths,
    width,
    atCol,
    sourceColumn,
  );
  return withFrameAfterAxisChange(
    block,
    withAxisEdit({ ...block, rows, colWidths }, { axis: "column", kind: "insert", index: atCol, count: 1, from: sourceColumn }),
    "col",
  );
}

/** Delete Rows: every row of the range, never the whole table. Null when it cannot. */
export function tableWithDeletedRows(
  block: PresentationTableBlock,
  range: TableAxisSelectionRange,
): PresentationTableBlock | null {
  if (range.count >= block.rows.length) return null;
  let rows = block.rows as PresentationTableCell[][];
  let rowHeights = block.rowHeights;
  let rowCount = block.rows.length;
  for (let offset = 0; offset < range.count; offset += 1) {
    const deleted = deleteTableRow(
      rows,
      range.start,
    ) as PresentationTableCell[][];
    if (deleted === rows) return null;
    rowHeights = deleteAxisSize(rowHeights, rowCount, range.start);
    rows = deleted;
    rowCount -= 1;
  }
  return withFrameAfterAxisChange(
    block,
    withAxisEdit(
      { ...block, rows, rowHeights },
      { axis: "row", kind: "delete", index: range.start, count: range.count },
    ),
    "row",
  );
}

/** Delete Columns: every column of the range, never the whole table. Null when it cannot. */
export function tableWithDeletedColumns(
  block: PresentationTableBlock,
  range: TableAxisSelectionRange,
): PresentationTableBlock | null {
  const width = tableLogicalWidth(block.rows);
  if (range.count >= width) return null;
  let rows = block.rows as PresentationTableCell[][];
  let colWidths = block.colWidths;
  let columnCount = width;
  for (let offset = 0; offset < range.count; offset += 1) {
    const deleted = deleteTableColumn(
      rows,
      range.start,
    ) as PresentationTableCell[][];
    if (deleted === rows) return null;
    colWidths = deleteAxisSize(colWidths, columnCount, range.start);
    rows = deleted;
    columnCount -= 1;
  }
  return withFrameAfterAxisChange(
    block,
    withAxisEdit(
      { ...block, rows, colWidths },
      { axis: "column", kind: "delete", index: range.start, count: range.count },
    ),
    "col",
  );
}

/** The table with its frame grown or shrunk by the rows (columns) inserted
 * or deleted, as PowerPoint draws it (`tableFrameAfterAxisChange`). */
function withFrameAfterAxisChange(
  before: PresentationTableBlock,
  after: PresentationTableBlock,
  axis: "col" | "row",
): PresentationTableBlock {
  const extent = tableFrameAfterAxisChange(before, after, axis);
  if (extent === null) return after;
  return axis === "col" ? { ...after, width: extent } : { ...after, height: extent };
}

/** Table Style Options: a flag set to true, or deleted. */
export function tableWithFlags(
  block: PresentationTableBlock,
  patch: Partial<
    Pick<
      PresentationTableBlock,
      | "borders"
      | "headerRow"
      | "bandedRows"
      | "totalRow"
      | "firstColumn"
      | "lastColumn"
      | "bandedColumns"
    >
  >,
): PresentationTableBlock {
  const next = { ...block } as Record<string, unknown>;
  for (const [key, value] of Object.entries(patch)) {
    if (value === true) next[key] = true;
    else delete next[key];
  }
  return next as unknown as PresentationTableBlock;
}

/**
 * Table Design > Effects > Shadow / Reflection: one pick on the whole table
 * (`value` null = "No Shadow" / "No Reflection"); the same block when nothing
 * changes (PPT-A137).
 */
export function tableWithEffect<K extends PresentationTableEffectKind>(
  block: PresentationTableBlock,
  kind: K,
  value: NonNullable<PresentationTableEffects[K]> | null,
): PresentationTableBlock {
  const effects = withPresentationTableEffect(block.tableEffects, kind, value);
  if (JSON.stringify(effects ?? null) === JSON.stringify(block.tableEffects ?? null)) return block;
  const next = { ...block } as Record<string, unknown>;
  if (effects) next.tableEffects = effects;
  else delete next.tableEffects;
  return next as unknown as PresentationTableBlock;
}

/**
 * Table Design > Shading > Table Background: the colour under every cell of
 * the table (`null` = No Fill); the same block when nothing changes.
 */
export function tableWithBackground(
  block: PresentationTableBlock,
  color: string | null,
): PresentationTableBlock {
  const value = color === null ? null : normalizePresentationTableBackground(color);
  if (color !== null && value === null) return block;
  if ((block.tableBackground ?? null) === value) return block;
  const next = { ...block } as Record<string, unknown>;
  if (value) next.tableBackground = value;
  else delete next.tableBackground;
  return next as unknown as PresentationTableBlock;
}

/** A cell's text (and runs); runs that no longer spell the text go. */
export function tableWithCellText(
  block: PresentationTableBlock,
  rowIndex: number,
  cellIndex: number,
  text: string,
  runs?: TextRun[],
): PresentationTableBlock {
  const rows = block.rows.map((row, nextRowIndex) =>
    nextRowIndex === rowIndex
      ? row.map((cell, nextCellIndex) =>
          nextCellIndex === cellIndex
            ? tableCellListsFollowingText(
                { ...cell, text, ...(runs ? { runs } : cell.text !== text ? { runs: undefined } : {}) },
                cell,
              )
            : cell,
        )
      : row,
  );
  return { ...block, rows };
}

/**
 * The cell's Bullets and Numbering (`textLayout`) carried to its words as
 * typed: Enter in a bulleted paragraph keeps the bullet on both halves, and
 * the stored rows stay one per paragraph for every reader and writer.
 */
function tableCellListsFollowingText(
  next: PresentationTableCell,
  before: PresentationTableCell,
): PresentationTableCell {
  if (before.textLayout === undefined || next.text === before.text) return next;
  const layout = normalizeOfficeTextLayout(before.textLayout);
  if (!layout) return next;
  const rows = officeTextLayoutFollowingText(
    { ...layout, paragraphs: tableCellListRows(before) },
    typeof before.text === "string" ? before.text : "",
    typeof next.text === "string" ? next.text : "",
  );
  return rows ? { ...next, textLayout: rows } : next;
}

/**
 * Clear Formatting over part of a cell that a row or cell operation sized
 * or coloured (PPT-B212): one write presses the cell's direct format into
 * the characters outside the range and drops it from the cell and the
 * range, which return to the table style. Null when the cell holds none of
 * `keys` itself (the plain character clear is enough then).
 */
export function tableWithCellCharacterFormatCleared(
  block: PresentationTableBlock,
  rowIndex: number,
  cellIndex: number,
  range: CellTextRange,
  keys: readonly string[],
): PresentationTableBlock | null {
  const cell = block.rows[rowIndex]?.[cellIndex];
  if (!cell) return null;
  const result = clearCellCharacterFormat(
    cell as PresentationTableCell & Record<string, unknown>,
    range,
    keys,
  );
  if (!result) return null;
  const next = { ...cell } as Record<string, unknown>;
  for (const key of result.cleared) delete next[key];
  if (result.runs) next.runs = result.runs;
  else delete next.runs;
  const rows = block.rows.map((row, nextRowIndex) =>
    nextRowIndex === rowIndex
      ? row.map((item, nextCellIndex) =>
          nextCellIndex === cellIndex ? (next as PresentationTableCell) : item,
        )
      : row,
  );
  return { ...block, rows };
}

/**
 * A cell format (shading, alignment, margins, borders, character format)
 * over every cell the selection touches. An edge written on the outside of
 * the selection is written on the neighbour's facing side too, so the line
 * between two cells is one line (PPT-B080).
 */
export function tableWithCellStyle(
  block: PresentationTableBlock,
  positions: ReadonlyArray<ReadonlyArray<{ row: number; col: number } | null | undefined>>,
  selection: PresentationTableCellRectangle,
  patch:
    | Partial<CellStyleInput>
    | ((
        position: TableBorderSelectionPosition,
        style: ResolvedCellStyle,
      ) => Partial<CellStyleInput>),
  options: {
    /**
     * Grow / Shrink Font: each run of a cell's words that states its own size
     * steps from that size (pt in, pt out; null at the end of the list), as
     * PowerPoint steps words of mixed sizes. Without it the patch's size goes
     * on every run.
     */
    readonly stepRunFontSizePt?: (pt: number) => number | null;
  } = {},
): PresentationTableBlock {
  // Outside edges this edit wrote, to give the neighbour across each the
  // same line (below).
  const writtenOuterEdges: Array<{
    row: number;
    col: number;
    rowEnd: number;
    colEnd: number;
    patch: Partial<CellStyleInput>;
  }> = [];
  const rows = block.rows.map((row, nextRowIndex) =>
    row.map((cell, nextCellIndex) => {
      const logicalPosition =
        positions[nextRowIndex]?.[nextCellIndex] ?? null;
      const selectionContainsCell =
        logicalPosition !== null &&
        tableCellIntersectsSelection(cell, logicalPosition, selection);
      if (!selectionContainsCell) return cell;
      const resolvedPatch =
        typeof patch === "function"
          ? patch(
              {
                row: logicalPosition.row,
                col: logicalPosition.col,
                row1: selection.row1,
                row2: selection.row2,
                col1: selection.col1,
                col2: selection.col2,
                rowSpan: cell.rowspan,
                colSpan: cell.colspan,
              },
              normalizeCellStyle(cell),
            )
          : patch;
      if (Object.keys(resolvedPatch).length === 0) return cell;
      if (
        "borderLeft" in resolvedPatch ||
        "borderRight" in resolvedPatch ||
        "borderTop" in resolvedPatch ||
        "borderBottom" in resolvedPatch
      ) {
        writtenOuterEdges.push({
          row: logicalPosition.row,
          col: logicalPosition.col,
          rowEnd: logicalPosition.row + Math.max(1, cell.rowspan ?? 1) - 1,
          colEnd: logicalPosition.col + Math.max(1, cell.colspan ?? 1) - 1,
          patch: resolvedPatch,
        });
      }
      const merged: CellStyleInput = {
        ...(cell as Record<string, unknown>),
        ...resolvedPatch,
      };
      const base = { ...cell } as Record<string, unknown>;
      for (const key of CELL_STYLE_KEYS) delete base[key];
      // Bold turned OFF (`false`, not the cleared `undefined`) on a row a
      // table option makes bold -- Header Row, Total Row -- is stated on the
      // cell's words, as PowerPoint writes `b="0"`: the cell itself keeps no
      // `false`, so the row's own bold would show again (PPT-B354). Anywhere
      // else "off" is simply the flag removed, as before.
      const flags = block as { headerRow?: unknown; totalRow?: unknown };
      const rowMakesBold =
        (flags.headerRow === true && logicalPosition.row === 0) ||
        (flags.totalRow === true &&
          logicalPosition.row + Math.max(1, cell.rowspan ?? 1) - 1 === block.rows.length - 1);
      const negatesRowBold =
        resolvedPatch.bold === false && rowMakesBold && cell.text.length > 0;
      return {
        ...base,
        text: cell.text,
        ...persistedCellStyle(merged),
        ...(Array.isArray(cell.runs) || negatesRowBold ? {
          runs: formatCellTextRange(cellCharacterRuns(cell.text, cell.runs), { start: 0, end: cell.text.length }, ((runPatch: Partial<RunStyle>) => {
            const step = options.stepRunFontSizePt;
            if (!step || typeof runPatch.fontSize !== "number") return runPatch;
            // A run with a size of its own steps from it; the rest take the
            // cell's stepped size.
            return (style: RunStyle | undefined): Partial<RunStyle> => {
              if (typeof style?.fontSize !== "number") return runPatch;
              const next = step(Math.round((style.fontSize / CSS_PX_PER_POINT) * 100) / 100);
              return { ...runPatch, fontSize: next === null ? style.fontSize : next * CSS_PX_PER_POINT };
            };
          })(Object.fromEntries(Object.entries(resolvedPatch).filter(([key]) => CELL_CHARACTER_FIELDS.has(key)).map(([key, value]) => [key, key === "fontSize" && typeof value === "number" ? value * CSS_PX_PER_POINT : key === "bold" && value === false && !negatesRowBold ? undefined : value])) as Partial<RunStyle>)),
        } : {}),
      } as PresentationTableCell;
    }),
  );
  // A line between two cells is one line: an edge written on the outside
  // of the selection is written on the neighbour's facing side too, so
  // the neighbour's old line (thicker, it wins the collapsed border) does
  // not keep showing (PPT-B080). Inside the selection both sides are
  // already written.
  const occupancy = writtenOuterEdges.length
    ? buildOccupancyGrid(rows)
    : null;
  if (occupancy) {
    const neighbourWrites = new Map<string, Partial<CellStyleInput>>();
    const mirror = (
      logicalRow: number,
      logicalCol: number,
      side: "borderLeft" | "borderRight" | "borderTop" | "borderBottom",
      border: unknown,
    ) => {
      if (
        logicalRow >= selection.row1 &&
        logicalRow <= selection.row2 &&
        logicalCol >= selection.col1 &&
        logicalCol <= selection.col2
      ) {
        return;
      }
      const occupant = occupancy[logicalRow]?.[logicalCol];
      if (!occupant) return;
      const key = `${occupant.anchorRow}:${occupant.cellIndex}`;
      neighbourWrites.set(key, {
        ...(neighbourWrites.get(key) ?? {}),
        [side]: border,
      });
    };
    for (const edge of writtenOuterEdges) {
      const { patch: edgePatch } = edge;
      for (let r = edge.row; r <= edge.rowEnd; r += 1) {
        if ("borderLeft" in edgePatch) mirror(r, edge.col - 1, "borderRight", edgePatch.borderLeft);
        if ("borderRight" in edgePatch) mirror(r, edge.colEnd + 1, "borderLeft", edgePatch.borderRight);
      }
      for (let c = edge.col; c <= edge.colEnd; c += 1) {
        if ("borderTop" in edgePatch) mirror(edge.row - 1, c, "borderBottom", edgePatch.borderTop);
        if ("borderBottom" in edgePatch) mirror(edge.rowEnd + 1, c, "borderTop", edgePatch.borderBottom);
      }
    }
    for (const [key, neighbourPatch] of neighbourWrites) {
      const [anchorRow, cellIndex] = key.split(":").map(Number);
      const cell = rows[anchorRow]?.[cellIndex];
      if (!cell) continue;
      const base = { ...cell } as Record<string, unknown>;
      for (const styleKey of CELL_STYLE_KEYS) delete base[styleKey];
      rows[anchorRow][cellIndex] = {
        ...base,
        text: cell.text,
        ...persistedCellStyle({
          ...(cell as Record<string, unknown>),
          ...neighbourPatch,
        }),
      } as PresentationTableCell;
    }
  }
  return { ...block, rows };
}

/**
 * Distribute Rows / Distribute Columns: every row (column) of the table, or
 * the selected range of them, made the same size. Columns are counted on the
 * LOGICAL grid the table draws (a merged cell spans several), never as a
 * row's physical cells (PPT-B034). A partial distribution shares out the
 * size the selected slots HAVE as drawn (`drawnExtent`, the table's drawn
 * width or height in CSS px): an automatic slot's is its share of the frame
 * (PPT-B289). Null when the sizes already are distributed.
 */
export function tableWithDistributedAxis(
  block: PresentationTableBlock,
  axis: "col" | "row",
  range?: (TableAxisSelectionRange & { drawnExtent: number }) | null,
): PresentationTableBlock | null {
  const count = axis === "col" ? tableLogicalWidth(block.rows) : block.rows.length;
  const key = axis === "col" ? "colWidths" : "rowHeights";
  const stated = (block as Record<string, unknown>)[key];
  let current = stated;
  if (range && range.drawnExtent > 0) {
    const drawn = Array.from({ length: count }, (_, index) =>
      materializedAxisSizeAt(stated, count, index, range.drawnExtent),
    );
    if (drawn.every((size) => size !== null)) current = drawn;
  }
  const nextSizes = range
    ? equalizedAxisSizeRange(current, count, range.start, range.end)
    : equalizedAxisSizes(current, count);
  const currentSizes = normalizeAxisSizes(stated, count);
  const unchanged = nextSizes
    ? currentSizes !== undefined &&
      currentSizes.length === nextSizes.length &&
      nextSizes.every((size, index) => Object.is(size, currentSizes[index]))
    : currentSizes === undefined;
  if (unchanged) return null;
  const next: PresentationTableBlock = { ...block };
  if (nextSizes) next[key] = nextSizes;
  else delete next[key];
  return next;
}
