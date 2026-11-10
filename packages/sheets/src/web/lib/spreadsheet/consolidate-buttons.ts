/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Consolidate as a pure plan over the workbook: the source references
 * (this sheet or another, never another workbook), the values the grid shows
 * in them, the consolidated block (with links to the sources when asked) and
 * every refusal checked before the first write. Moved out of SpreadsheetWindow
 * so the dialog and the spreadsheet editor commands run the same code;
 * behaviour is unchanged. The window keeps the dialog text and the commit.
 */

import { writeCellValueInto, type CellMap } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { rangesOverlap } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import {
  consolidate,
  consolidateLinkFormula,
  type ConsolidateFunction,
  type ConsolidateValue,
} from "./consolidate";
import { resolveCellChartScalar, resolveCellDisplay } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import {
  columnToLetters,
  isFormulaError,
  parseCellRef,
  type SheetEvaluator,
  type SpillOverlay,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { rangeHasMerge, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { formatInternalLink, parseInternalLink } from "@mrnote/core/web/lib/spreadsheet/sheet-interactions.ts";

type Rect = { r1: number; r2: number; c1: number; c2: number };

export type ConsolidateSourceParseResult =
  | {
      readonly kind: "ok";
      readonly sheetName?: string;
      readonly range: Rect;
    }
  | { readonly kind: "external" }
  | { readonly kind: "invalid" };

/** Parse one Consolidate source as A1:B4 or 'Other Sheet'!$A$1:$B$4. */
export function parseConsolidateSource(raw: string): ConsolidateSourceParseResult {
  const text = raw.trim();
  if (text === "") return { kind: "invalid" };
  // External references require relationship/provider custody that this
  // workbook-local command does not own.
  if (text.includes("[") || text.includes("]")) return { kind: "external" };
  const qualified = text.includes("!") ? parseInternalLink(`#${text}`) : null;
  if (text.includes("!") && !qualified) return { kind: "invalid" };
  const ref = qualified?.ref ?? text;
  const parts = ref.split(":");
  if (parts.length < 1 || parts.length > 2) return { kind: "invalid" };
  const first = parseCellRef(parts[0]);
  const second = parts.length === 2 ? parseCellRef(parts[1]) : first;
  if (!first || !second) return { kind: "invalid" };
  return {
    kind: "ok",
    ...(qualified ? { sheetName: qualified.sheetName } : {}),
    range: {
      r1: Math.min(first.row, second.row),
      r2: Math.max(first.row, second.row),
      c1: Math.min(first.col, second.col),
      c2: Math.max(first.col, second.col),
    },
  };
}

export function absoluteCellReference(row: number, col: number): string {
  return `$${columnToLetters(col)}$${row + 1}`;
}

export interface ConsolidateRequest {
  /** The destination's top-left cell, as typed. */
  destination: string;
  /** One reference per source (2-64). */
  sources: readonly string[];
  fn: ConsolidateFunction;
  mode: "position" | "category";
  topRow: boolean;
  leftColumn: boolean;
  createLinks: boolean;
}

export type ConsolidateError =
  | "bad-dest"
  | "need-label"
  | "bad-source"
  | "external"
  | "no-room"
  | "overlap"
  | "links-same-sheet"
  | "merged"
  | "link-too-long"
  | "spill"
  | "locked";

export interface ConsolidateSheet {
  id: string;
  name: string;
  cells: CellMap;
  rowCount: number;
  columnCount: number;
  merges?: MergeSpan[];
}

/** A source cell's value as a formula reading it gets it. */
function consolidateCellValue(
  cell: CellMap[string] | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
): ConsolidateValue {
  let value = resolveCellChartScalar(cell, row, col, evaluator, overlay);
  // A constant's characters are typed as formulas type them (a number, a
  // logical, a stored #N/A, or text when the cell is text).
  if (!cell?.formula && cell?.value && typeof value === "string") value = evaluator.getComputed(row, col);
  if (isFormulaError(value)) return { error: value.code };
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (value === null || value === "") return null;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return typeof value === "string" ? value : "";
}

/**
 * Consolidate into `host` at the destination. Values are read through the
 * same workbook evaluator and spill authority the grid uses (formulas on
 * other sheets included); link mode uses the helper's contributor buckets, so
 * position and category matching cannot diverge from the static result: label
 * cells stay literal and aggregate cells become absolute, sheet-qualified
 * formulas. Protection, merges and dynamic-array spill footprints are checked
 * before the first write, so a refusal changes nothing.
 */
export function consolidatePlan<S extends ConsolidateSheet>(
  sheets: readonly S[],
  host: S,
  request: ConsolidateRequest,
  read: {
    evaluatorFor: (sheet: S) => SheetEvaluator;
    overlayFor: (sheet: S, evaluator: SheetEvaluator) => SpillOverlay;
    isEditable: (row: number, col: number) => boolean;
  },
): { ok: false; error: ConsolidateError } | { ok: true; sheet: S; rows: number; cols: number } {
  const destRef = parseCellRef(request.destination.trim());
  if (!destRef) return { ok: false, error: "bad-dest" };
  if (request.mode === "category" && !request.topRow && !request.leftColumn) {
    return { ok: false, error: "need-label" };
  }
  const sourceLines = request.sources.map((line) => line.trim()).filter((line) => line !== "");
  if (sourceLines.length < 2 || sourceLines.length > 64) return { ok: false, error: "bad-source" };
  const rangeInBounds = (range: Rect, targetSheet: S) =>
    range.r1 >= 0 &&
    range.r2 >= range.r1 &&
    range.r2 < targetSheet.rowCount &&
    range.c1 >= 0 &&
    range.c2 >= range.c1 &&
    range.c2 < targetSheet.columnCount;
  const resolvedSources: Array<{ sheet: S; range: Rect }> = [];
  for (const line of sourceLines) {
    const parsed = parseConsolidateSource(line);
    if (parsed.kind === "external") return { ok: false, error: "external" };
    if (parsed.kind !== "ok") return { ok: false, error: "bad-source" };
    let sourceSheet = host;
    if (parsed.sheetName !== undefined) {
      const wanted = parsed.sheetName.trim().toLowerCase();
      const matches = sheets.filter((candidate) => (candidate.name || candidate.id).trim().toLowerCase() === wanted);
      if (matches.length !== 1) return { ok: false, error: "bad-source" };
      sourceSheet = matches[0];
    }
    if (!rangeInBounds(parsed.range, sourceSheet)) return { ok: false, error: "bad-source" };
    resolvedSources.push({ sheet: sourceSheet, range: parsed.range });
  }
  const sourceGrids: string[][][] = [];
  // What the functions aggregate: the values formulas read, not the display
  // (1,000 under a thousands format is 1000; a quoted '7 is text; #N/A is an
  // error that Excel's consolidation carries out).
  const sourceValues: ConsolidateValue[][][] = [];
  let hostOverlay: SpillOverlay | null = null;
  try {
    for (const source of resolvedSources) {
      const evaluator = read.evaluatorFor(source.sheet);
      const overlay = read.overlayFor(source.sheet, evaluator);
      if (source.sheet === host) hostOverlay = overlay;
      const rows: string[][] = [];
      const values: ConsolidateValue[][] = [];
      for (let r = source.range.r1; r <= source.range.r2; r += 1) {
        const row: string[] = [];
        const valueRow: ConsolidateValue[] = [];
        for (let c = source.range.c1; c <= source.range.c2; c += 1) {
          const cell = source.sheet.cells[`${r}:${c}`];
          row.push(resolveCellDisplay(cell, r, c, evaluator, overlay));
          valueRow.push(consolidateCellValue(cell, r, c, evaluator, overlay));
        }
        rows.push(row);
        values.push(valueRow);
      }
      sourceGrids.push(rows);
      sourceValues.push(values);
    }
  } catch {
    return { ok: false, error: "bad-source" };
  }
  const result = consolidate({
    sources: sourceGrids,
    values: sourceValues,
    fn: request.fn,
    mode:
      request.mode === "position"
        ? { kind: "position" }
        : { kind: "category", useTopRow: request.topRow, useLeftColumn: request.leftColumn },
  });
  if (result.rows === 0 || result.cols === 0) return { ok: false, error: "bad-source" };
  const outputRange = {
    r1: destRef.row,
    r2: destRef.row + result.rows - 1,
    c1: destRef.col,
    c2: destRef.col + result.cols - 1,
  };
  if (!rangeInBounds(outputRange, host)) return { ok: false, error: "no-room" };
  if (resolvedSources.some((source) => source.sheet.id === host.id && rangesOverlap(source.range, outputRange))) {
    return { ok: false, error: "overlap" };
  }
  if (request.createLinks && resolvedSources.some((source) => source.sheet.id === host.id)) {
    return { ok: false, error: "links-same-sheet" };
  }
  if (host.merges && rangeHasMerge(host.merges, outputRange)) return { ok: false, error: "merged" };
  const outputValues = result.grid.map((row) => [...row]);
  if (request.createLinks) {
    for (let r = 0; r < result.rows; r += 1) {
      for (let c = 0; c < result.cols; c += 1) {
        const refs = result.contributors[r][c].map((contributor) => {
          const source = resolvedSources[contributor.sourceIndex];
          return formatInternalLink(
            source.sheet.name || source.sheet.id,
            absoluteCellReference(source.range.r1 + contributor.row, source.range.c1 + contributor.col),
          ).slice(1);
        });
        if (refs.length === 0) continue;
        const formula = consolidateLinkFormula(request.fn, refs);
        if (!formula) return { ok: false, error: "link-too-long" };
        outputValues[r][c] = formula;
      }
    }
  }
  const spill = hostOverlay ?? read.overlayFor(host, read.evaluatorFor(host));
  for (let r = 0; r < result.rows; r += 1) {
    for (let c = 0; c < result.cols; c += 1) {
      const row = destRef.row + r;
      const col = destRef.col + c;
      if (spill.spilled.has(`${row}:${col}`)) return { ok: false, error: "spill" };
      if (!read.isEditable(row, col)) return { ok: false, error: "locked" };
    }
  }
  const nextCells: CellMap = { ...host.cells };
  for (let r = 0; r < result.rows; r += 1) {
    for (let c = 0; c < result.cols; c += 1) {
      writeCellValueInto(nextCells, destRef.row + r, destRef.col + c, outputValues[r][c]);
    }
  }
  return { ok: true, sheet: { ...host, cells: nextCells }, rows: result.rows, cols: result.cols };
}
