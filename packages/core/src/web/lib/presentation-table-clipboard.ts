/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native presentation-table clipboard helpers.
 *
 * PowerPoint and WPS copy a selected table row/column/cell rectangle as a
 * bounded matrix: pasting into another selected rectangle replaces cell
 * contents (and, for an in-app cell payload, direct cell formatting) without
 * inserting/removing table axes. Plain TSV from Excel/WPS replaces text only,
 * so the destination table's formatting remains intact.
 *
 * The persisted table uses HTML-style merged rows. All reads and writes go
 * through the logical occupancy grid; covered continuation coordinates are
 * never materialized as fake cells and a paste never flattens a native merge.
 */

import {
  buildOccupancyGrid,
  logicalCellPositions,
  tableLogicalWidth,
  type MergeableCell,
} from "./presentation-table-merge";
import {
  tableCellIntersectsSelection,
  type TableSelectionRectangle,
} from "./presentation-table-selection";
import { persistedCellRuns } from "./presentation-cell-rich-text";

export const PRESENTATION_TABLE_CLIPBOARD_MIME =
  "application/x-mrai-presentation-table+json";

const TABLE_CLIPBOARD_KIND = "mrai.presentation-table.cells.v1";
const MAX_CLIPBOARD_AXIS = 256;
const MAX_CLIPBOARD_TEXT = 1_000_000;

export interface PresentationTableClipboardPayload {
  kind: typeof TABLE_CLIPBOARD_KIND;
  mode: "cells" | "text";
  width: number;
  height: number;
  cells: Array<Array<MergeableCell | null>>;
}

export interface ApplyTableClipboardResult {
  rows: MergeableCell[][];
  applied: number;
  skippedMergedContinuations: number;
  /**
   * Copied cells with text that would land only inside a merged cell and so
   * be lost. The paste is refused whole when this is above 0: nothing is
   * written (PPT-B048).
   */
  lostSourceCells: number;
  /**
   * Copied cells with text that would land past the table's last row or
   * column. Refused whole the same way, and said so (PPT-B185): the paste
   * wrote the cells that fitted and dropped the rest without a word.
   */
  beyondTableSourceCells?: number;
}

function positiveSpan(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 1;
  return Math.max(1, Math.floor(value));
}

function cloneCell(cell: MergeableCell): MergeableCell {
  return JSON.parse(JSON.stringify(cell)) as MergeableCell;
}

function normalizedSelection(
  rows: MergeableCell[][],
  selection: TableSelectionRectangle,
): TableSelectionRectangle | null {
  const grid = buildOccupancyGrid(rows);
  const positions = logicalCellPositions(rows);
  const width = tableLogicalWidth(rows);
  if (!grid || width <= 0 || rows.length === 0) return null;

  let next: TableSelectionRectangle = {
    row1: Math.max(0, Math.min(selection.row1, selection.row2)),
    row2: Math.min(rows.length - 1, Math.max(selection.row1, selection.row2)),
    col1: Math.max(0, Math.min(selection.col1, selection.col2)),
    col2: Math.min(width - 1, Math.max(selection.col1, selection.col2)),
  };
  if (next.row1 > next.row2 || next.col1 > next.col2) return null;

  // A native merged cell is indivisible. Expand until every cell intersecting
  // the rectangle is fully contained (a newly-expanded edge can reveal another
  // span, hence the small fixed-point loop).
  for (let pass = 0; pass < rows.length * Math.max(1, width); pass += 1) {
    const previous = { ...next };
    for (let row = 0; row < rows.length; row += 1) {
      for (let cellIndex = 0; cellIndex < rows[row].length; cellIndex += 1) {
        const cell = rows[row][cellIndex];
        const position = positions[row]?.[cellIndex] ?? null;
        if (!tableCellIntersectsSelection(cell, position, next) || !position) {
          continue;
        }
        next = {
          row1: Math.min(next.row1, position.row),
          row2: Math.min(
            rows.length - 1,
            Math.max(next.row2, position.row + positiveSpan(cell.rowspan) - 1),
          ),
          col1: Math.min(next.col1, position.col),
          col2: Math.min(
            width - 1,
            Math.max(next.col2, position.col + positiveSpan(cell.colspan) - 1),
          ),
        };
      }
    }
    if (
      previous.row1 === next.row1 &&
      previous.row2 === next.row2 &&
      previous.col1 === next.col1 &&
      previous.col2 === next.col2
    ) {
      return next;
    }
  }
  return next;
}

export function captureTableSelection(
  rows: MergeableCell[][],
  selection: TableSelectionRectangle,
): PresentationTableClipboardPayload | null {
  const normalized = normalizedSelection(rows, selection);
  const grid = buildOccupancyGrid(rows);
  if (!normalized || !grid) return null;
  const height = normalized.row2 - normalized.row1 + 1;
  const width = normalized.col2 - normalized.col1 + 1;
  if (
    height <= 0 ||
    width <= 0 ||
    height > MAX_CLIPBOARD_AXIS ||
    width > MAX_CLIPBOARD_AXIS
  ) {
    return null;
  }

  const cells = Array.from({ length: height }, (_, rowOffset) =>
    Array.from({ length: width }, (_, colOffset) => {
      const occupant =
        grid[normalized.row1 + rowOffset]?.[normalized.col1 + colOffset] ??
        null;
      if (!occupant?.isAnchor) return null;
      const cell = rows[occupant.anchorRow]?.[occupant.cellIndex];
      return cell ? cloneCell(cell) : null;
    }),
  );
  return {
    kind: TABLE_CLIPBOARD_KIND,
    mode: "cells",
    width,
    height,
    cells,
  };
}

/**
 * Split clipboard text into TSV records, or `null` as soon as the text exceeds
 * the row / column budget. The budget is enforced WHILE parsing so a payload of
 * many short lines is refused in bounded work instead of first growing an
 * unbounded record list (`"\n".repeat(180000)` used to reach `Math.max(...)`
 * over every row and overflow the call stack before the row limit was checked).
 */
function parseDelimitedRows(text: string): string[][] | null {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const endField = (): boolean => {
    row.push(field);
    field = "";
    return row.length <= MAX_CLIPBOARD_AXIS;
  };
  // One record beyond the axis budget is tolerated here: a terminal newline
  // produces an empty trailing record that is trimmed below.
  const endRow = (): boolean => {
    if (!endField()) return false;
    rows.push(row);
    row = [];
    return rows.length <= MAX_CLIPBOARD_AXIS + 1;
  };
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"' && field.length === 0) {
      quoted = true;
    } else if (char === "\t") {
      if (!endField()) return null;
    } else if (char === "\r" && text[index + 1] === "\n") {
      if (!endRow()) return null;
      index += 1;
    } else if (char === "\n" || char === "\r") {
      if (!endRow()) return null;
    } else {
      field += char;
    }
  }
  if (!endRow()) return null;
  // Clipboard producers commonly append one terminal newline. It terminates the
  // final record; it is not an extra blank table row.
  if (
    rows.length > 1 &&
    rows[rows.length - 1].length === 1 &&
    rows[rows.length - 1][0] === "" &&
    /(?:\r\n|\r|\n)$/.test(text)
  ) {
    rows.pop();
  }
  return rows;
}

export function parseTableClipboardText(
  raw: string,
): PresentationTableClipboardPayload | null {
  if (raw.length > MAX_CLIPBOARD_TEXT) return null;
  const parsed = parseDelimitedRows(raw);
  if (!parsed) return null;
  const height = parsed.length;
  let width = 0;
  for (const row of parsed) width = Math.max(width, row.length);
  if (
    height <= 0 ||
    width <= 0 ||
    height > MAX_CLIPBOARD_AXIS ||
    width > MAX_CLIPBOARD_AXIS
  ) {
    return null;
  }
  return {
    kind: TABLE_CLIPBOARD_KIND,
    mode: "text",
    width,
    height,
    cells: parsed.map((row) =>
      Array.from({ length: width }, (_, column) => ({
        text: row[column] ?? "",
      })),
    ),
  };
}

function quoteClipboardField(text: string): string {
  return /[\t\r\n"]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

export function serializeTableClipboardText(
  payload: PresentationTableClipboardPayload,
): string {
  return payload.cells
    .map((row) =>
      row
        .map((cell) => quoteClipboardField(String(cell?.text ?? "")))
        .join("\t"),
    )
    .join("\n");
}

function escapeHtml(text: string): string {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function serializeTableClipboardHtml(
  payload: PresentationTableClipboardPayload,
): string {
  const payloadAttribute =
    payload.mode === "cells"
      ? ` data-mrai-presentation-table="${encodeURIComponent(
          encodeTableClipboardPayload(payload),
        )}"`
      : "";
  const rows = payload.cells.map(
    (row) =>
      `<tr>${row
        .map((cell) => {
          // A null slot is a covered continuation of a merged anchor. Emitting
          // a second <td> beside a colspan/rowspan anchor would create a
          // malformed Office table and silently change its logical width.
          if (!cell) return "";
          const colspan = positiveSpan(cell.colspan);
          const rowspan = positiveSpan(cell.rowspan);
          const spans = `${rowspan > 1 ? ` rowspan="${rowspan}"` : ""}${
            colspan > 1 ? ` colspan="${colspan}"` : ""
          }`;
          return `<td${spans}>${escapeHtml(String(cell.text ?? "")).replaceAll(
            "\n",
            "<br>",
          )}</td>`;
        })
        .join("")}</tr>`,
  );
  // Safari can reject custom clipboard MIME types. Embed the same bounded,
  // versioned payload in the HTML flavor so an in-app paste still preserves
  // cell formatting without relying on a process-local cache. PowerPoint,
  // WPS, Excel, and browsers ignore the unknown data attribute.
  return `<table${payloadAttribute}>${rows.join("")}</table>`;
}

export function encodeTableClipboardPayload(
  payload: PresentationTableClipboardPayload,
): string {
  return JSON.stringify(payload);
}

export function decodeTableClipboardPayload(
  raw: string,
): PresentationTableClipboardPayload | null {
  if (!raw || raw.length > MAX_CLIPBOARD_TEXT) return null;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return null;
  }
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (
    record.kind !== TABLE_CLIPBOARD_KIND ||
    record.mode !== "cells" ||
    !Number.isInteger(record.width) ||
    !Number.isInteger(record.height)
  ) {
    return null;
  }
  const width = Number(record.width);
  const height = Number(record.height);
  if (
    width <= 0 ||
    height <= 0 ||
    width > MAX_CLIPBOARD_AXIS ||
    height > MAX_CLIPBOARD_AXIS ||
    !Array.isArray(record.cells) ||
    record.cells.length !== height
  ) {
    return null;
  }
  const cells: Array<Array<MergeableCell | null>> = [];
  for (const rawRow of record.cells) {
    if (!Array.isArray(rawRow) || rawRow.length !== width) return null;
    const row: Array<MergeableCell | null> = [];
    for (const rawCell of rawRow) {
      if (rawCell === null) {
        row.push(null);
        continue;
      }
      if (
        !rawCell ||
        typeof rawCell !== "object" ||
        typeof (rawCell as { text?: unknown }).text !== "string"
      ) {
        return null;
      }
      row.push(cloneCell(rawCell as MergeableCell));
    }
    cells.push(row);
  }
  return {
    kind: TABLE_CLIPBOARD_KIND,
    mode: "cells",
    width,
    height,
    cells,
  };
}

export function decodeTableClipboardHtml(
  raw: string,
): PresentationTableClipboardPayload | null {
  if (!raw || raw.length > MAX_CLIPBOARD_TEXT * 2) return null;
  const match = raw.match(
    /<table\b[^>]*\bdata-mrai-presentation-table=(["'])(.*?)\1/i,
  );
  if (!match) return null;
  try {
    return decodeTableClipboardPayload(decodeURIComponent(match[2]));
  } catch {
    return null;
  }
}

function replaceCellContent(
  target: MergeableCell,
  source: MergeableCell,
  mode: PresentationTableClipboardPayload["mode"],
): MergeableCell {
  if (mode === "text") {
    // Plain text replaces the cell's characters, so the old character runs no
    // longer describe them: the pasted text takes the cell-level formatting.
    const next: MergeableCell = { ...target, text: source.text };
    delete next.runs;
    return next;
  }
  const cloned = cloneCell(source);
  delete cloned.colspan;
  delete cloned.rowspan;
  // An in-app cell payload carries its own runs; persist them only when they
  // describe the pasted text (and drop a run set the source cannot vouch for).
  const runs = persistedCellRuns(cloned.text, cloned.runs);
  if (runs) cloned.runs = runs;
  else delete cloned.runs;
  return {
    ...cloned,
    ...(positiveSpan(target.colspan) > 1
      ? { colspan: positiveSpan(target.colspan) }
      : {}),
    ...(positiveSpan(target.rowspan) > 1
      ? { rowspan: positiveSpan(target.rowspan) }
      : {}),
  };
}

export function applyTableClipboard(
  rows: MergeableCell[][],
  payload: PresentationTableClipboardPayload,
  start: { row: number; col: number },
  targetSelection?: TableSelectionRectangle | null,
): ApplyTableClipboardResult {
  const grid = buildOccupancyGrid(rows);
  if (!grid) {
    return { rows, applied: 0, skippedMergedContinuations: 0, lostSourceCells: 0 };
  }
  const selection = targetSelection
    ? normalizedSelection(rows, targetSelection)
    : null;
  const startRow = selection?.row1 ?? start.row;
  const startCol = selection?.col1 ?? start.col;
  const height = selection
    ? selection.row2 - selection.row1 + 1
    : payload.height;
  const width = selection ? selection.col2 - selection.col1 + 1 : payload.width;
  if (height <= 0 || width <= 0) {
    return { rows, applied: 0, skippedMergedContinuations: 0, lostSourceCells: 0 };
  }

  const next = rows.map((row) => row.map((cell) => ({ ...cell })));
  const updatedAnchors = new Set<string>();
  let applied = 0;
  let skippedMergedContinuations = 0;
  const writtenSources = new Set<string>();
  const coveredSources = new Set<string>();
  const beyondSources = new Set<string>();
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const targetRow = startRow + rowOffset;
      const targetCol = startCol + colOffset;
      const sourceKey = `${payload.height === 1 ? 0 : rowOffset % payload.height}:${
        payload.width === 1 ? 0 : colOffset % payload.width
      }`;
      const occupant = grid[targetRow]?.[targetCol] ?? null;
      if (!occupant) {
        // No cell there: past the last row or column. Its content has
        // nowhere to land.
        beyondSources.add(sourceKey);
        continue;
      }
      if (!occupant.isAnchor) {
        skippedMergedContinuations += 1;
        coveredSources.add(sourceKey);
        continue;
      }
      writtenSources.add(sourceKey);
      const anchorKey = `${occupant.anchorRow}:${occupant.cellIndex}`;
      if (updatedAnchors.has(anchorKey)) continue;
      const sourceRow = payload.height === 1 ? 0 : rowOffset % payload.height;
      const sourceCol = payload.width === 1 ? 0 : colOffset % payload.width;
      const source = payload.cells[sourceRow]?.[sourceCol] ?? null;
      const target = next[occupant.anchorRow]?.[occupant.cellIndex];
      if (!source || !target) continue;
      next[occupant.anchorRow][occupant.cellIndex] = replaceCellContent(
        target,
        source,
        payload.mode,
      );
      updatedAnchors.add(anchorKey);
      applied += 1;
    }
  }
  let lostSourceCells = 0;
  for (const key of coveredSources) {
    if (writtenSources.has(key)) continue;
    const [row, col] = key.split(":").map(Number);
    if ((payload.cells[row]?.[col]?.text ?? "") !== "") lostSourceCells += 1;
  }
  if (lostSourceCells > 0) {
    return { rows, applied: 0, skippedMergedContinuations, lostSourceCells };
  }
  let beyondTableSourceCells = 0;
  for (const key of beyondSources) {
    if (writtenSources.has(key)) continue;
    const [row, col] = key.split(":").map(Number);
    if ((payload.cells[row]?.[col]?.text ?? "") !== "") beyondTableSourceCells += 1;
  }
  if (beyondTableSourceCells > 0) {
    return {
      rows,
      applied: 0,
      skippedMergedContinuations,
      lostSourceCells: 0,
      beyondTableSourceCells,
    };
  }
  return {
    rows: applied > 0 ? next : rows,
    applied,
    skippedMergedContinuations,
    lostSourceCells,
  };
}

export function clearTableSelection(
  rows: MergeableCell[][],
  selection: TableSelectionRectangle,
): MergeableCell[][] {
  const normalized = normalizedSelection(rows, selection);
  const grid = buildOccupancyGrid(rows);
  if (!normalized || !grid) return rows;
  const next = rows.map((row) => row.map((cell) => ({ ...cell })));
  const cleared = new Set<string>();
  for (let row = normalized.row1; row <= normalized.row2; row += 1) {
    for (let col = normalized.col1; col <= normalized.col2; col += 1) {
      const occupant = grid[row]?.[col] ?? null;
      if (!occupant?.isAnchor) continue;
      const key = `${occupant.anchorRow}:${occupant.cellIndex}`;
      if (cleared.has(key)) continue;
      const cell = next[occupant.anchorRow]?.[occupant.cellIndex];
      if (!cell || (cell.text === "" && cell.runs === undefined)) continue;
      // Cleared text has no characters left for the runs to describe.
      const emptied: MergeableCell = { ...cell, text: "" };
      delete emptied.runs;
      next[occupant.anchorRow][occupant.cellIndex] = emptied;
      cleared.add(key);
    }
  }
  return cleared.size > 0 ? next : rows;
}
