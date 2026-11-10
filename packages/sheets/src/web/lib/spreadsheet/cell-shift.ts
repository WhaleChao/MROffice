/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert / Delete CELLS with a directional shift (Excel/WPS "Insert Cells…" and
 * "Delete Cells…" — the dialog with Shift cells down/right/up/left), as opposed
 * to the whole-row/whole-column ops in structural.ts.
 *
 * A cell-range shift moves only the cells on the SAME rows (for a horizontal
 * shift) or SAME columns (for a vertical shift) as the selected rectangle — the
 * rest of the grid is untouched:
 *
 *  - INSERT + shift DOWN  : in each column of the range, cells at row ≥ r1 move
 *    down by the range height; the range cells become blank.
 *  - INSERT + shift RIGHT : in each row of the range, cells at col ≥ c1 move right
 *    by the range width; the range cells become blank.
 *  - DELETE + shift UP    : in each column of the range, the range cells are
 *    removed and cells below (row > r2) move up by the range height.
 *  - DELETE + shift LEFT  : in each row of the range, the range cells are removed
 *    and cells to the right (col > c2) move left by the range width.
 *
 * FORMULAS: references to moved cells adjust automatically, including absolute
 * references. A formula range the moved band crosses only in part is left as
 * written, as Excel leaves it (so are defined names, the print area, chart
 * series and the other formula-shaped references below). A validation or
 * conditional-format target the band crosses in part is cut along the band
 * and follows piece by piece, as Excel's multi-area sqref does (XLS-A009).
 * Range-owned metadata Excel itself refuses to cut (a Table, a merged area, a
 * PivotTable) still fails closed before mutation.
 *
 * Pure and generic (operates on a "row:col" cell map) so the remap is
 * unit-testable and reusable; the component supplies its full SpreadsheetCell.
 */

import { protectRelativeNameRefs } from "@mrnote/core/web/lib/spreadsheet/name-relative-refs.ts";
import {
  adjustFormulaForCellShift,
  canAdjustFormulaForCellShift,
  canAdjustPreparedFormulaForCellShift,
  prepareFormulaForCellShift,
  shiftFormula,
  type FormulaCellShiftOperation,
  type PreparedCellShiftFormula,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { MAX_CONDITIONAL_AREAS } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { cellHoldsSheetEndContent, drawingAnchorClosedByDelete, shiftDrawingAnchorLines } from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import { spreadsheetObjectOffsetsLayoutRelative } from "@mrnote/core/web/lib/spreadsheet/layout-relative-anchor.ts";

export interface ShiftCellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type ShiftDirection = "down" | "right" | "up" | "left";

export interface ShiftFormulaCell {
  value?: string;
  formula?: string;
  [key: string]: unknown;
}

export interface ShiftMetadataRange {
  range: ShiftCellRange;
}

export interface ShiftValidationMetadataRange extends ShiftMetadataRange {
  range: ShiftCellRange & { sheet?: string };
  kind?: string;
  formula?: string;
  sourceRange?: string;
  source?: { row: number; col: number };
  sourceMode?: "fixed" | "perRow";
}

export interface ShiftConditionalMetadataRange extends ShiftMetadataRange {
  kind?: string;
  formula?: string;
  nativeBlocker?: string;
  /** A rule's further rectangles, moved with its first (live audit B304). */
  areas?: ShiftCellRange[];
}

export interface ShiftDataTableMetadataRange extends ShiftMetadataRange {
  id?: string;
  columns?: Array<{ id: number }>;
  nativeResizeBlocker?: string;
}

export interface ShiftNamedMetadataRange {
  ref: string;
  scope?: "workbook" | "worksheet";
  scopeSheetId?: string;
  nativeBlocker?: string;
}

// These metadata contracts name only the fields the shift reads. They carry
// no index signature: the editor's SheetModel members are interfaces, which
// never satisfy one, and every remap is generic so the caller's full type
// (a Sparkline's colours, a Scenario cell's value, a Slicer's layout) comes
// back unchanged.
export interface ShiftScenarioMetadata {
  cells: ReadonlyArray<{ ref: string }>;
}

export interface ShiftSparklineMetadata {
  range?: string;
}

export interface ShiftPivotMetadata {
  source: string;
  /** The worksheet the source range is on when that is not the PivotTable's
   * own sheet: by id for an authored report, by name for an imported one. */
  sourceSheetId?: string;
  output?: string;
  native?: { sourceSheetName?: string } | unknown;
  /** "scenario": a Scenario PivotTable report, whose source is no range. */
  cacheSource?: string;
}

export interface ShiftSlicerMetadata {
  column: number;
  source?:
    | { kind: "table"; tableId: string; tableColumnId: number }
    | { kind: "pivot"; pivotOutput: string; fieldId: string };
}

export interface CellShiftMetadataSheet {
  id?: string;
  name?: string;
  validations?: ShiftValidationMetadataRange[];
  conditionalFormats?: ShiftConditionalMetadataRange[];
  conditionalFormatsNativeBlocker?: string;
  dataTables?: ShiftDataTableMetadataRange[];
  merges?: Array<{ row: number; col: number; rowSpan: number; colSpan: number }>;
  charts?: Array<ShiftMetadataRange & { anchor?: { row: number; col: number } }>;
  objects?: ShiftDrawingMetadata[];
  namedRanges?: ShiftNamedMetadataRange[];
  pivot?: ShiftPivotMetadata;
  morePivots?: ShiftPivotMetadata[];
  scenarios?: ShiftScenarioMetadata[];
  sparklines?: Record<string, ShiftSparklineMetadata>;
  watches?: Array<{ row: number; col: number }>;
  slicers?: ShiftSlicerMetadata[];
  filters?: Record<string | number, unknown>;
  colorFilters?: Record<string | number, unknown>;
  conditionFilters?: Record<string | number, unknown>;
  printSettings?: { printArea?: string };
}

export interface ShiftDrawingAnchorPoint {
  row?: number;
  column?: number;
  col?: number;
  rowOffset?: number;
  columnOffset?: number;
  colOffset?: number;
}

export interface ShiftDrawingMetadata {
  type?: string;
  range?: string;
  ranges?: string[];
  nativeChartBlocker?: string;
  anchor?: {
    kind?: "oneCell" | "twoCell" | "absolute";
    from?: ShiftDrawingAnchorPoint;
    to?: ShiftDrawingAnchorPoint;
  };
}

export type CellShiftMetadataBlocker =
  | "validation"
  | "conditional-format"
  | "table"
  | "merge"
  | "chart"
  | "drawing"
  | "defined-name"
  | "pivot"
  | "scenario"
  | "sparkline"
  | "watch"
  | "slicer"
  | "filter"
  | "print-area";

/** Normalize so r1<=r2 and c1<=c2 (an anchor/active pair may be inverted). */
function normalize(range: ShiftCellRange): ShiftCellRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}

function rangeTouchesShift(
  candidate: ShiftCellRange,
  target: ShiftCellRange,
  direction: ShiftDirection,
): boolean {
  const value = normalize(candidate);
  const shift = normalize(target);
  if (direction === "down" || direction === "up") {
    return (
      value.c2 >= shift.c1 &&
      value.c1 <= shift.c2 &&
      value.r2 >= shift.r1
    );
  }
  return (
    value.r2 >= shift.r1 &&
    value.r1 <= shift.r2 &&
    value.c2 >= shift.c1
  );
}

export interface CellShiftRangeRemapResult {
  range: ShiftCellRange | null;
  supported: boolean;
}

function rangeMappingSupported(
  candidate: ShiftCellRange,
  target: ShiftCellRange,
  direction: ShiftDirection,
): boolean {
  const value = normalize(candidate);
  const shift = normalize(target);
  const vertical = direction === "down" || direction === "up";
  if (vertical) {
    const intersectsCorridor = value.c2 >= shift.c1 && value.c1 <= shift.c2;
    if (!intersectsCorridor || value.r2 < shift.r1) return true;
    return value.c1 >= shift.c1 && value.c2 <= shift.c2;
  }
  const intersectsCorridor = value.r2 >= shift.r1 && value.r1 <= shift.r2;
  if (!intersectsCorridor || value.c2 < shift.c1) return true;
  return value.r1 >= shift.r1 && value.r2 <= shift.r2;
}

/**
 * Remap one range through an Insert/Delete Cells corridor. A range that crosses
 * only part of the corridor would become non-rectangular and is therefore
 * unsupported. `null` means the whole range was deleted.
 */
export function remapCellShiftRange(
  candidate: ShiftCellRange,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftRangeRemapResult {
  const value = normalize(candidate);
  const shift = normalize(target);
  if (!rangeMappingSupported(value, shift, direction)) {
    return { range: value, supported: false };
  }
  if (!rangeTouchesShift(value, shift, direction)) {
    return { range: value, supported: true };
  }

  const height = shift.r2 - shift.r1 + 1;
  const width = shift.c2 - shift.c1 + 1;
  let { r1, c1, r2, c2 } = value;
  if (direction === "down") {
    if (r1 >= shift.r1) r1 += height;
    if (r2 >= shift.r1) r2 += height;
  } else if (direction === "right") {
    if (c1 >= shift.c1) c1 += width;
    if (c2 >= shift.c1) c2 += width;
  } else if (direction === "up") {
    if (r1 >= shift.r1 && r1 <= shift.r2) r1 = shift.r1;
    else if (r1 > shift.r2) r1 -= height;
    if (r2 >= shift.r1 && r2 <= shift.r2) r2 = shift.r1 - 1;
    else if (r2 > shift.r2) r2 -= height;
  } else {
    if (c1 >= shift.c1 && c1 <= shift.c2) c1 = shift.c1;
    else if (c1 > shift.c2) c1 -= width;
    if (c2 >= shift.c1 && c2 <= shift.c2) c2 = shift.c1 - 1;
    else if (c2 > shift.c2) c2 -= width;
  }
  return r1 > r2 || c1 > c2
    ? { range: null, supported: true }
    : { range: { r1, c1, r2, c2 }, supported: true };
}

/** One rectangle of a range after a shift, with the ORIGINAL cell that is
 * now its top-left (what a relative formula must be rebased from). */
export interface CellShiftRangePiece {
  range: ShiftCellRange;
  origin: { row: number; col: number };
}

/**
 * XLS-A009 — Excel keeps a range-owned target (a data validation's or a
 * conditional format's Applies to) through ANY Insert/Delete Cells shift: the
 * part of the range on the moved rows (columns) travels with its cells and the
 * rest stays, so a range the band crosses only in part becomes a multi-area
 * range (E2:E30 with a cell inserted at B3, shifting right, is E2,E4:E30,F3).
 * The range is cut along the band's edges and each part is remapped on its
 * own; a part the delete erases is gone. Pieces keep the original reading
 * order (the part holding the original top-left first).
 */
export function splitCellShiftRange(
  candidate: ShiftCellRange,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftRangePiece[] {
  const value = normalize(candidate);
  const shift = normalize(target);
  const pieceOf = (part: ShiftCellRange): CellShiftRangePiece | null => {
    const moved = remapCellShiftRange(part, shift, direction);
    if (!moved.range) return null;
    // Only a delete can drop the original top-left: the new top-left is then
    // the first line after the deleted band.
    const origin = { row: part.r1, col: part.c1 };
    if (direction === "up" && part.r1 >= shift.r1 && part.r1 <= shift.r2) {
      origin.row = shift.r2 + 1;
    } else if (direction === "left" && part.c1 >= shift.c1 && part.c1 <= shift.c2) {
      origin.col = shift.c2 + 1;
    }
    return { range: moved.range, origin };
  };
  if (rangeMappingSupported(value, shift, direction)) {
    const piece = pieceOf(value);
    return piece ? [piece] : [];
  }
  const vertical = direction === "down" || direction === "up";
  const parts: ShiftCellRange[] = vertical
    ? [
        { ...value, c2: shift.c1 - 1 },
        { ...value, c1: Math.max(value.c1, shift.c1), c2: Math.min(value.c2, shift.c2) },
        { ...value, c1: shift.c2 + 1 },
      ]
    : [
        { ...value, r2: shift.r1 - 1 },
        { ...value, r1: Math.max(value.r1, shift.r1), r2: Math.min(value.r2, shift.r2) },
        { ...value, r1: shift.r2 + 1 },
      ];
  const pieces: CellShiftRangePiece[] = [];
  for (const part of parts) {
    const clipped = vertical
      ? { ...part, c1: Math.max(part.c1, value.c1), c2: Math.min(part.c2, value.c2) }
      : { ...part, r1: Math.max(part.r1, value.r1), r2: Math.min(part.r2, value.r2) };
    if (clipped.r1 > clipped.r2 || clipped.c1 > clipped.c2) continue;
    const piece = pieceOf(clipped);
    if (piece) pieces.push(piece);
  }
  return pieces;
}

/** A formula relative to `from` restated for the cell `to` (fill semantics). */
function rebaseRelativeFormula(
  formula: string,
  from: { row: number; col: number },
  to: { row: number; col: number },
): string {
  return to.row === from.row && to.col === from.col
    ? formula
    : shiftFormula(formula, to.row - from.row, to.col - from.col);
}

function remapCellShiftPoint(
  point: { row: number; col: number },
  target: ShiftCellRange,
  direction: ShiftDirection,
): { row: number; col: number } | null {
  const shift = normalize(target);
  const height = shift.r2 - shift.r1 + 1;
  const width = shift.c2 - shift.c1 + 1;
  let { row, col } = point;
  const vertical = direction === "down" || direction === "up";
  const inCorridor = vertical
    ? col >= shift.c1 && col <= shift.c2
    : row >= shift.r1 && row <= shift.r2;
  if (!inCorridor) return point;
  if (direction === "down" && row >= shift.r1) row += height;
  else if (direction === "right" && col >= shift.c1) col += width;
  else if (direction === "up") {
    if (row >= shift.r1 && row <= shift.r2) return null;
    if (row > shift.r2) row -= height;
  } else if (direction === "left") {
    if (col >= shift.c1 && col <= shift.c2) return null;
    if (col > shift.c2) col -= width;
  }
  return { row, col };
}

function remapCellShiftAnchorPoint(
  point: { row: number; col: number },
  target: ShiftCellRange,
  direction: ShiftDirection,
): { row: number; col: number } {
  const shift = normalize(target);
  const mapped = remapCellShiftPoint(point, shift, direction);
  if (mapped) return mapped;
  return direction === "up"
    ? { row: shift.r1, col: point.col }
    : { row: point.row, col: shift.c1 };
}

export interface CellShiftCollectionRemap<T> {
  items: T[] | undefined;
  supported: boolean;
}

function remapEmbeddedFormula(
  value: string,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): { value: string; supported: boolean } {
  const result = adjustFormulaForCellShift(value, {
    range: target,
    direction,
    sheetName,
  });
  return { value: result.formula, supported: result.supported };
}

/** Remap validation targets plus every modeled same-sheet dependency. */
export function remapCellShiftValidations<T extends ShiftValidationMetadataRange>(
  rules: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): CellShiftCollectionRemap<T> {
  if (!rules) return { items: rules, supported: true };
  const items: T[] = [];
  for (const rule of rules) {
    if (rule.range.sheet !== undefined) {
      if (!rangeTouchesShift(rule.range, target, direction)) {
        items.push(rule);
        continue;
      }
      return { items: rules, supported: false };
    }
    // A target the band crosses in part becomes one rule per rectangle
    // (MRAI's validations are single-rectangle; Excel's sqref is the union).
    // Each one's relative formula is restated for its own top-left, since a
    // custom/list formula is relative to the rule's first cell (XLS-A009).
    const pieces = splitCellShiftRange(rule.range, target, direction);
    // A per-row cascade reads its own row's control cell; a split would need
    // per-piece control columns this model does not have.
    if (pieces.length > 1 && rule.source && rule.sourceMode === "perRow") {
      return { items: rules, supported: false };
    }
    const anchor = { row: Math.min(rule.range.r1, rule.range.r2), col: Math.min(rule.range.c1, rule.range.c2) };
    for (const piece of pieces) {
      let next: ShiftValidationMetadataRange = {
        ...rule,
        range: { ...rule.range, ...piece.range },
      };
      if (typeof rule.formula === "string") {
        const formula = remapEmbeddedFormula(
          rebaseRelativeFormula(rule.formula, anchor, piece.origin),
          target,
          direction,
          sheetName,
        );
        if (!formula.supported) return { items: rules, supported: false };
        next = { ...next, formula: formula.value };
      }
      if (typeof rule.sourceRange === "string") {
        const hasEquals = rule.sourceRange.startsWith("=");
        const source = remapEmbeddedFormula(
          rebaseRelativeFormula(
            hasEquals ? rule.sourceRange : `=${rule.sourceRange}`,
            anchor,
            piece.origin,
          ),
          target,
          direction,
          sheetName,
        );
        if (!source.supported || source.value.includes("#REF!")) {
          return { items: rules, supported: false };
        }
        next = {
          ...next,
          sourceRange: hasEquals ? source.value : source.value.slice(1),
        };
      }
      if (rule.source) {
        const source = remapCellShiftPoint(rule.source, target, direction);
        if (!source) return { items: rules, supported: false };
        next = { ...next, source };
      }
      items.push(next as T);
    }
  }
  return { items, supported: true };
}

/** Remap conditional-format targets and modeled formula rules. */
export function remapCellShiftConditionalFormats<
  T extends ShiftConditionalMetadataRange,
>(
  rules: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): CellShiftCollectionRemap<T> {
  if (!rules) return { items: rules, supported: true };
  const items: T[] = [];
  for (const rule of rules) {
    const sourceAreas = rule.areas && rule.areas.length > 1 ? rule.areas : [rule.range];
    if (rule.nativeBlocker) {
      if (!sourceAreas.some((area) => rangeTouchesShift(area, target, direction))) {
        items.push(rule);
        continue;
      }
      return { items: rules, supported: false };
    }
    // Every area travels with the cells it covers; an area the band crosses
    // in part is cut along the band (XLS-A009), so the rule's Applies to
    // becomes the union of the pieces, as Excel's sqref does.
    const anchor = {
      row: Math.min(rule.range.r1, rule.range.r2),
      col: Math.min(rule.range.c1, rule.range.c2),
    };
    const pieces = sourceAreas.flatMap((area) =>
      splitCellShiftRange(area, target, direction),
    );
    if (pieces.length === 0) continue;
    const anchorMoved =
      pieces[0]!.origin.row !== anchor.row || pieces[0]!.origin.col !== anchor.col;
    const split = pieces.length !== sourceAreas.length || anchorMoved;
    const relativeFormula =
      typeof rule.formula === "string" &&
      shiftFormula(rule.formula, 1, 1) !== rule.formula;
    const restate = (
      next: ShiftConditionalMetadataRange,
      origin: { row: number; col: number },
    ): ShiftConditionalMetadataRange | null => {
      if (typeof rule.formula !== "string") return next;
      const formula = remapEmbeddedFormula(
        rebaseRelativeFormula(rule.formula, anchor, origin),
        target,
        direction,
        sheetName,
      );
      return formula.supported ? { ...next, formula: formula.value } : null;
    };
    if (split && relativeFormula) {
      // A relative formula is read from the rule's first cell; pieces that
      // moved differently cannot share one anchor, so each becomes its own
      // rule (adjacent, same settings: the same cells format the same way).
      for (const [index, piece] of pieces.entries()) {
        const { areas: _areas, ...rest } = rule;
        void _areas;
        const base: ShiftConditionalMetadataRange = { ...rest, range: piece.range };
        const id = (rule as { id?: unknown }).id;
        if (index > 0 && typeof id === "string") {
          (base as { id?: string }).id = `${id}~${index}`;
        }
        const next = restate(base, piece.origin);
        if (!next) return { items: rules, supported: false };
        items.push(next as T);
      }
      continue;
    }
    if (pieces.length > MAX_CONDITIONAL_AREAS) {
      return { items: rules, supported: false };
    }
    let next: ShiftConditionalMetadataRange = {
      ...rule,
      range: pieces[0]!.range,
    };
    // `areas` lists every rectangle, the first being the rule's own range; a
    // rule left with one rectangle is a plain rule again.
    if (pieces.length > 1) next.areas = pieces.map((piece) => piece.range);
    else delete next.areas;
    const restated = restate(next, pieces[0]!.origin);
    if (!restated) return { items: rules, supported: false };
    next = restated;
    items.push(next as T);
  }
  return { items, supported: true };
}

/** Remap merge rectangles; a merge collapsed to one cell is removed. */
export function remapCellShiftMerges<
  T extends { row: number; col: number; rowSpan: number; colSpan: number },
>(
  merges: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftCollectionRemap<T> {
  if (!merges) return { items: merges, supported: true };
  const items: T[] = [];
  for (const merge of merges) {
    const result = remapCellShiftRange(
      {
        r1: merge.row,
        c1: merge.col,
        r2: merge.row + Math.max(1, merge.rowSpan) - 1,
        c2: merge.col + Math.max(1, merge.colSpan) - 1,
      },
      target,
      direction,
    );
    if (!result.supported) return { items: merges, supported: false };
    if (result.range === null) continue;
    const rowSpan = result.range.r2 - result.range.r1 + 1;
    const colSpan = result.range.c2 - result.range.c1 + 1;
    if (rowSpan <= 1 && colSpan <= 1) continue;
    items.push({
      ...merge,
      row: result.range.r1,
      col: result.range.c1,
      rowSpan,
      colSpan,
    });
  }
  return { items, supported: true };
}

/**
 * A ListObject may ride through a partial shift only as a whole rectangle. Any
 * operation that would resize, split, or erase it remains owned by Table Design.
 */
export function remapCellShiftDataTables<
  T extends ShiftDataTableMetadataRange,
>(
  tables: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftCollectionRemap<T> {
  if (!tables) return { items: tables, supported: true };
  const items: T[] = [];
  for (const table of tables) {
    const result = remapCellShiftRange(table.range, target, direction);
    if (!result.supported) return { items: tables, supported: false };
    if (result.range === null) return { items: tables, supported: false };
    const changed =
      result.range.r1 !== table.range.r1 ||
      result.range.c1 !== table.range.c1 ||
      result.range.r2 !== table.range.r2 ||
      result.range.c2 !== table.range.c2;
    const sameShape =
      result.range.r2 - result.range.r1 === table.range.r2 - table.range.r1 &&
      result.range.c2 - result.range.c1 === table.range.c2 - table.range.c1;
    // A table whose own filter/sort MRAI only carries ("activeFilter" /
    // "sortState", see tableNativeFilterCustody) moves with its cells: the
    // writer follows its columns by id, so only other native data pins it.
    const pinned =
      Boolean(table.nativeResizeBlocker) &&
      table.nativeResizeBlocker !== "activeFilter" &&
      table.nativeResizeBlocker !== "sortState";
    if (!sameShape || (changed && pinned)) {
      return { items: tables, supported: false };
    }
    items.push({ ...table, range: result.range });
  }
  return { items, supported: true };
}

export function remapCellShiftNamedRanges<T extends ShiftNamedMetadataRange>(
  names: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetId?: string,
  sheetName?: string,
): CellShiftCollectionRemap<T> {
  if (!names) return { items: names, supported: true };
  const items: T[] = [];
  for (const name of names) {
    if (
      name.scope === "worksheet" &&
      name.scopeSheetId &&
      sheetId &&
      name.scopeSheetId !== sheetId
    ) {
      items.push(name);
      continue;
    }
    // A relative name's offsets are not cells: they stay as they are while
    // its absolute axes move with the cells, as in Excel.
    const relative = protectRelativeNameRefs(
      name.ref,
      direction === "down" || direction === "up" ? "row" : "col",
    );
    const result = remapEmbeddedFormula(
      relative?.text ?? name.ref,
      target,
      direction,
      sheetName,
    );
    const value = relative ? relative.restore(result.value) : result.value;
    if (!result.supported || (name.nativeBlocker && value !== name.ref)) {
      return { items: names, supported: false };
    }
    items.push(
      value === name.ref ? name : ({ ...name, ref: value } as T),
    );
  }
  return { items, supported: true };
}

export interface ShiftPrintSettings {
  printArea?: string;
}

export function remapCellShiftPrintSettings<T extends ShiftPrintSettings>(
  settings: T | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): { item: T | undefined; supported: boolean } {
  if (!settings?.printArea) return { item: settings, supported: true };
  // Excel keeps several comma-separated areas in one name; each block is
  // remapped on its own and an erased block leaves the list.
  const blocks: string[] = [];
  for (const block of settings.printArea.split(",")) {
    const source = block.trim();
    if (!source) continue;
    const result = remapEmbeddedFormula(source, target, direction, sheetName);
    if (!result.supported) return { item: settings, supported: false };
    if (!result.value.includes("#REF!")) blocks.push(result.value);
  }
  const value = blocks.join(",");
  if (value === settings.printArea) {
    return { item: settings, supported: true };
  }
  const next = { ...settings };
  if (!value) delete next.printArea;
  else next.printArea = value;
  // Settings left empty (the only print area erased) are no settings, as
  // Clear Print Area leaves them; `{}` read back as none (sheet model fuzz).
  return { item: Object.keys(next).length > 0 ? next : undefined, supported: true };
}

/**
 * Remap native chart data ranges and their floating cell anchors. A chart's
 * series are formulas in Excel, so a data range the band crosses only in part
 * stays as written; one the shift deletes outright still refuses.
 */
export function remapCellShiftCharts<
  T extends ShiftMetadataRange & { anchor?: { row: number; col: number } },
>(
  charts: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftCollectionRemap<T> {
  if (!charts) return { items: charts, supported: true };
  const items: T[] = [];
  for (const chart of charts) {
    // A #REF! source (see StructuralChart) names no cells; only its anchor moves.
    const brokenSource = (chart as { sourceRefError?: unknown }).sourceRefError === true;
    const range = brokenSource
      ? { range: chart.range, supported: true }
      : remapCellShiftRange(chart.range, target, direction);
    if (range.supported && range.range === null) {
      return { items: charts, supported: false };
    }
    const anchor = chart.anchor
      ? remapCellShiftAnchorPoint(chart.anchor, target, direction)
      : undefined;
    // The chart's cell anchor (lib/spreadsheet/chart-anchor) follows the
    // shifted cells as any drawing's does.
    const drawingAnchor = (chart as { drawingAnchor?: NonNullable<ShiftDrawingMetadata["anchor"]> }).drawingAnchor;
    const movedDrawing = drawingAnchor ? cellShiftDrawingAnchor(drawingAnchor, target, direction, false) : undefined;
    if (movedDrawing === null) continue;
    items.push({
      ...chart,
      range: range.supported && range.range ? range.range : chart.range,
      ...(anchor ? { anchor } : {}),
      ...(movedDrawing && movedDrawing !== drawingAnchor ? { drawingAnchor: movedDrawing } : {}),
    });
  }
  return { items, supported: true };
}

/** Remap image/shape/chart DrawingML cell anchors and chart source strings. */
export function remapCellShiftDrawings<T extends ShiftDrawingMetadata>(
  objects: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): CellShiftCollectionRemap<T> {
  if (!objects) return { items: objects, supported: true };
  const items: T[] = [];
  for (const object of objects) {
    let next: ShiftDrawingMetadata = object;
    if (object.type === "chart") {
      const sources = [object.range, ...(object.ranges ?? [])].filter(
        (value): value is string => typeof value === "string",
      );
      const remapped: string[] = [];
      for (const source of sources) {
        const result = remapEmbeddedFormula(
          source,
          target,
          direction,
          sheetName,
        );
        if (
          !result.supported ||
          result.value.includes("#REF!") ||
          (object.nativeChartBlocker && result.value !== source)
        ) {
          return { items: objects, supported: false };
        }
        remapped.push(result.value);
      }
      const hasSingle = typeof object.range === "string";
      next = {
        ...next,
        ...(hasSingle ? { range: remapped[0] } : {}),
        ...(object.ranges
          ? { ranges: remapped.slice(hasSingle ? 1 : 0) }
          : {}),
      };
    }
    const anchor = object.anchor;
    if (anchor) {
      const moved = cellShiftDrawingAnchor(anchor, target, direction, spreadsheetObjectOffsetsLayoutRelative(object));
      if (moved === null) continue;
      if (moved !== anchor) next = { ...next, anchor: moved };
    }
    items.push(next as T);
  }
  return { items, supported: true };
}












function cellShiftDrawingAnchor<A extends NonNullable<ShiftDrawingMetadata["anchor"]>>(
  anchor: A,
  target: ShiftCellRange,
  direction: ShiftDirection,
  keepOffsets: boolean,
): A | null {
  if (anchor.kind === "absolute") return anchor;
  const shift = normalize(target);
  const vertical = direction === "down" || direction === "up";
  const lead = vertical ? (anchor.from?.column ?? anchor.from?.col) : anchor.from?.row;
  const [low, high] = vertical ? [shift.c1, shift.c2] : [shift.r1, shift.r2];
  if (!Number.isInteger(lead) || lead! < low || lead! > high) return anchor;
  const axis = vertical ? "row" : "col";
  const at = vertical ? shift.r1 : shift.c1;
  const lines = vertical ? shift.r2 - shift.r1 + 1 : shift.c2 - shift.c1 + 1;
  const moved = shiftDrawingAnchorLines(
    anchor,
    axis,
    at,
    direction === "down" || direction === "right" ? lines : -lines,
    { keepOffsets },
  );
  return drawingAnchorClosedByDelete(anchor, moved, axis) ? null : moved;
}

export function remapCellShiftScenarios<T extends ShiftScenarioMetadata>(
  scenarios: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): CellShiftCollectionRemap<T> {
  if (!scenarios) return { items: scenarios, supported: true };
  const items: T[] = [];
  for (const scenario of scenarios) {
    const cells: Array<T["cells"][number]> = [];
    for (const cell of scenario.cells) {
      const result = remapEmbeddedFormula(
        cell.ref,
        target,
        direction,
        sheetName,
      );
      if (!result.supported) return { items: scenarios, supported: false };
      if (result.value.includes("#REF!")) continue;
      cells.push({ ...cell, ref: result.value });
    }
    items.push({ ...scenario, cells } as T);
  }
  return { items, supported: true };
}

export function remapCellShiftSparklines<T extends ShiftSparklineMetadata>(
  sparklines: Record<string, T> | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): { item: Record<string, T> | undefined; supported: boolean } {
  if (!sparklines) return { item: sparklines, supported: true };
  const item: Record<string, T> = {};
  for (const [key, sparkline] of Object.entries(sparklines)) {
    const [row, col] = key.split(":").map(Number);
    if (!Number.isInteger(row) || !Number.isInteger(col)) {
      item[key] = sparkline;
      continue;
    }
    const anchor = remapCellShiftPoint({ row, col }, target, direction);
    if (!anchor) continue;
    let next = sparkline;
    if (typeof sparkline.range === "string" && sparkline.range !== "") {
      const source = remapEmbeddedFormula(
        sparkline.range,
        target,
        direction,
        sheetName,
      );
      if (!source.supported) return { item: sparklines, supported: false };
      if (source.value.includes("#REF!")) continue;
      if (source.value !== sparkline.range) {
        next = { ...sparkline, range: source.value };
      }
    }
    item[`${anchor.row}:${anchor.col}`] = next;
  }
  return { item, supported: true };
}

export function remapCellShiftWatches<T extends { row: number; col: number }>(
  watches: T[] | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
): CellShiftCollectionRemap<T> {
  if (!watches) return { items: watches, supported: true };
  const items: T[] = [];
  for (const watch of watches) {
    const point = remapCellShiftPoint(watch, target, direction);
    if (!point) continue;
    items.push({ ...watch, ...point });
  }
  return { items, supported: true };
}

/**
 * A Pivot recipe may follow only a whole-source translation. Resizing its source
 * would change field offsets/cache semantics, while moving its output anchor
 * requires the computed output footprint that this pure layer does not own.
 *
 * `parts` says which of the PivotTable's two ranges are on the edited sheet:
 * a report on another worksheet has only its SOURCE there (0b0fd2f48c), and a
 * report whose data is elsewhere has only its OUTPUT there. Default: both.
 */
export function remapCellShiftPivot<T extends ShiftPivotMetadata>(
  pivot: T | undefined,
  target: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
  parts: { source?: boolean; output?: boolean } = {},
): { item: T | undefined; supported: boolean } {
  if (!pivot) return { item: pivot, supported: true };
  const { source: sourceInParts = true, output: outputHere = true } = parts;
  // A Scenario PivotTable report's `source` is only its cache's shape.
  const sourceHere = sourceInParts && pivot.cacheSource !== "scenario";
  let nextSource = pivot.source;
  if (sourceHere) {
    const sourceRange = parseBoundA1Range(pivot.source);
    if (!sourceRange) return { item: pivot, supported: false };
    const source = remapCellShiftRange(sourceRange, target, direction);
    if (!source.supported || source.range === null) {
      return { item: pivot, supported: false };
    }
    const sameShape =
      source.range.r2 - source.range.r1 === sourceRange.r2 - sourceRange.r1 &&
      source.range.c2 - source.range.c1 === sourceRange.c2 - sourceRange.c1;
    // Records added to or deleted from the source (its header row and its
    // columns kept) leave every field where it was: the PivotTable reads the
    // longer or shorter list, as Excel's does over a table that grows or
    // shrinks (live audit S-090: deleting a record of a PivotTable's source
    // table was refused). A move of the header or a change of columns still
    // is not followed.
    const recordsOnly =
      source.range.c1 === sourceRange.c1 &&
      source.range.c2 === sourceRange.c2 &&
      source.range.r1 === sourceRange.r1 &&
      source.range.r2 > source.range.r1;
    if (!sameShape && !(recordsOnly && !pivot.native)) return { item: pivot, supported: false };
    const sourceFormula = remapEmbeddedFormula(
      pivot.source,
      target,
      direction,
      sheetName,
    );
    if (!sourceFormula.supported || sourceFormula.value.includes("#REF!")) {
      return { item: pivot, supported: false };
    }
    nextSource = sourceFormula.value;
  }
  if (outputHere && pivot.output) {
    const outputFormula = remapEmbeddedFormula(
      pivot.output,
      target,
      direction,
      sheetName,
    );
    if (!outputFormula.supported || outputFormula.value !== pivot.output) {
      return { item: pivot, supported: false };
    }
  }
  if (pivot.native && nextSource !== pivot.source) {
    return { item: pivot, supported: false };
  }
  return {
    supported: true,
    item:
      nextSource === pivot.source
        ? pivot
        : ({ ...pivot, source: nextSource } as T),
  };
}

/**
 * Whether `pivot`, a PivotTable of `host`, reads its source from `edited`.
 * Mirrors the editor's source-sheet lookup (and structural.ts's): by
 * `sourceSheetId` first, then an imported report's sheet name, else the
 * PivotTable's own sheet.
 */
export function cellShiftPivotSourceIsOnSheet(
  pivot: ShiftPivotMetadata,
  host: Pick<CellShiftMetadataSheet, "id" | "name">,
  edited: Pick<CellShiftMetadataSheet, "id" | "name">,
): boolean {
  if (typeof pivot.sourceSheetId === "string" && pivot.sourceSheetId) {
    return pivot.sourceSheetId === edited.id;
  }
  const nativeName = (pivot.native as { sourceSheetName?: unknown } | undefined)
    ?.sourceSheetName;
  if (typeof nativeName === "string" && nativeName) {
    return nativeName === edited.name;
  }
  return host === edited || (host.id !== undefined && host.id === edited.id);
}

/**
 * Every PivotTable of `host` through an Insert/Delete Cells shift on `edited`,
 * which may be `host` itself or another worksheet. Excel moves a PivotTable's
 * source range like any reference when the whole block it names shifts on
 * the SOURCE sheet, wherever the report lives, and leaves a range the band
 * crosses only in part as written; this editor refuses that partial case
 * before mutation, for a report on any sheet alike. A report's OUTPUT anchor
 * answers only to a shift on its own sheet. Returns `host` itself when no
 * PivotTable of it changes.
 */
export function remapCellShiftPivotsOnSheet<T extends CellShiftMetadataSheet>(
  host: T,
  edited: Pick<CellShiftMetadataSheet, "id" | "name">,
  target: ShiftCellRange,
  direction: ShiftDirection,
): { sheet: T; supported: boolean } {
  if (!host.pivot && !host.morePivots?.length) {
    return { sheet: host, supported: true };
  }
  const outputHere =
    host === edited || (host.id !== undefined && host.id === edited.id);
  let supported = true;
  let changed = false;
  const remap = <P extends ShiftPivotMetadata>(pivot: P): P => {
    const plan = remapCellShiftPivot(pivot, target, direction, edited.name, {
      source: cellShiftPivotSourceIsOnSheet(pivot, host, edited),
      output: outputHere,
    });
    if (!plan.supported) supported = false;
    if (plan.item !== undefined && plan.item !== pivot) changed = true;
    return plan.item ?? pivot;
  };
  const pivot = host.pivot ? remap(host.pivot) : host.pivot;
  const morePivots = host.morePivots?.map(remap);
  if (!supported) return { sheet: host, supported: false };
  if (!changed) return { sheet: host, supported: true };
  return {
    sheet: {
      ...host,
      ...(host.pivot ? { pivot } : {}),
      ...(host.morePivots ? { morePivots } : {}),
    },
    supported: true,
  };
}

/** Keep slicer placement aligned with a whole-translated source table. */
export function remapCellShiftSlicers<T extends ShiftSlicerMetadata>(
  slicers: T[] | undefined,
  beforeTables: ShiftDataTableMetadataRange[] | undefined,
  afterTables: ShiftDataTableMetadataRange[] | undefined,
): CellShiftCollectionRemap<T> {
  if (!slicers) return { items: slicers, supported: true };
  const beforeById = new Map(
    (beforeTables ?? []).flatMap((table) =>
      table.id ? [[table.id, table] as const] : [],
    ),
  );
  const afterById = new Map(
    (afterTables ?? []).flatMap((table) =>
      table.id ? [[table.id, table] as const] : [],
    ),
  );
  const items = slicers.map((slicer) => {
    const source = slicer.source;
    if (source?.kind !== "table") return slicer;
    const before = beforeById.get(source.tableId);
    const after = afterById.get(source.tableId);
    if (!before || !after) return slicer;
    const declaredOffset = before.columns?.findIndex(
      (column) => column.id === source.tableColumnId,
    );
    const fallbackOffset = slicer.column - before.range.c1;
    const offset =
      declaredOffset !== undefined && declaredOffset >= 0
        ? declaredOffset
        : fallbackOffset;
    if (offset < 0 || after.range.c1 + offset > after.range.c2) return slicer;
    const column = after.range.c1 + offset;
    return column === slicer.column ? slicer : ({ ...slicer, column } as T);
  });
  return { items, supported: true };
}

function parseBoundA1Range(value: string | undefined): ShiftCellRange | null {
  if (!value) return null;
  const bang = value.lastIndexOf("!");
  const body = (bang >= 0 ? value.slice(bang + 1) : value).replace(/\$/g, "");
  return parseA1Range(body);
}

/**
 * Return the first native metadata family whose coordinates would become stale
 * under a partial cell shift. Unaffected objects elsewhere on the sheet do not
 * block the operation.
 */
export function cellShiftMetadataBlocker(
  sheet: CellShiftMetadataSheet,
  range: ShiftCellRange,
  direction: ShiftDirection,
  /** The workbook's other worksheets: a PivotTable there that reads `sheet`
   * has its source range in the shift's way too. */
  otherSheets: readonly CellShiftMetadataSheet[] = [],
): CellShiftMetadataBlocker | null {
  if (
    !remapCellShiftValidations(
      sheet.validations,
      range,
      direction,
      sheet.name,
    ).supported
  ) {
    return "validation";
  }
  if (
    (sheet.conditionalFormatsNativeBlocker &&
      sheet.conditionalFormats?.some((item) =>
        rangeTouchesShift(item.range, range, direction),
      )) ||
    !remapCellShiftConditionalFormats(
      sheet.conditionalFormats,
      range,
      direction,
      sheet.name,
    ).supported
  ) {
    return "conditional-format";
  }
  if (!remapCellShiftDataTables(sheet.dataTables, range, direction).supported) {
    return "table";
  }
  if (!remapCellShiftMerges(sheet.merges, range, direction).supported) {
    return "merge";
  }
  if (!remapCellShiftCharts(sheet.charts, range, direction).supported) {
    return "chart";
  }
  if (
    !remapCellShiftDrawings(
      sheet.objects,
      range,
      direction,
      sheet.name,
    ).supported
  ) {
    return "drawing";
  }
  if (
    !remapCellShiftNamedRanges(
      sheet.namedRanges,
      range,
      direction,
      sheet.id,
      sheet.name,
    ).supported
  ) {
    return "defined-name";
  }
  if (
    !remapCellShiftPivotsOnSheet(sheet, sheet, range, direction).supported ||
    otherSheets.some(
      (other) =>
        other !== sheet &&
        !remapCellShiftPivotsOnSheet(other, sheet, range, direction).supported,
    )
  ) {
    return "pivot";
  }
  if (
    !remapCellShiftScenarios(
      sheet.scenarios,
      range,
      direction,
      sheet.name,
    ).supported
  ) {
    return "scenario";
  }
  if (
    !remapCellShiftSparklines(
      sheet.sparklines,
      range,
      direction,
      sheet.name,
    ).supported
  ) {
    return "sparkline";
  }
  if (!remapCellShiftWatches(sheet.watches, range, direction).supported) {
    return "watch";
  }
  return remapCellShiftPrintSettings(
    sheet.printSettings,
    range,
    direction,
    sheet.name,
  ).supported
    ? null
    : "print-area";
}

function formulaSource(cell: ShiftFormulaCell): string | null {
  if (typeof cell.formula === "string" && cell.formula.startsWith("=")) {
    return cell.formula;
  }
  return typeof cell.value === "string" && cell.value.startsWith("=")
    ? cell.value
    : null;
}

/** Build a read-only shortlist for repeated selection capability checks.
 * A colon range is the only place the writer reads range endpoints, and an
 * unreadable endpoint refuses. Without a colon (including a 3-D sheet span), a
 * successfully tokenized formula is supported for every selection. The existing predicate validates
 * that fast path; tokenization failures remain in the shortlist and fail closed.
 * Callers may memoize this snapshot against their immutable cell map. Actual
 * commands must still validate the full current map before writing.
 * `formulaKeys` may name every stored formula key of this map so a large sheet
 * is not enumerated; each named key is still read from the current map. */
export function cellShiftRangeSensitiveFormulas<T extends ShiftFormulaCell>(
  cells: Record<string, T>,
  formulaKeys?: readonly string[],
): Record<string, T> {
  const candidates: Record<string, T> = Object.create(null);
  const seen = new Set<string>();
  const probe: FormulaCellShiftOperation = {
    range: { r1: 0, c1: 0, r2: 0, c2: 0 }, direction: "down",
  };
  for (const key of formulaKeys ?? Object.keys(cells)) {
    const cell = cells[key];
    if (!cell) continue;
    const source = formulaSource(cell);
    if (source === null || seen.has(source)) continue;
    seen.add(source);
    if (source.includes(":") || !canAdjustFormulaForCellShift(source, probe)) {
      candidates[key] = cell;
    }
  }
  return candidates;
}

export function canAdjustCellShiftFormulas<T extends ShiftFormulaCell>(
  cells: Record<string, T>,
  range: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): boolean {
  const operation: FormulaCellShiftOperation = { range, direction, sheetName };
  return Object.values(cells).every((candidate) => {
    const source = formulaSource(candidate);
    return source === null || canAdjustFormulaForCellShift(source, operation);
  });
}

/** Tokenize a candidate set once so the capability check for each selection
 * only compares rectangles. Same decision as canAdjustCellShiftFormulas. */
export function prepareCellShiftFormulas<T extends ShiftFormulaCell>(
  cells: Record<string, T>,
): ReadonlyArray<PreparedCellShiftFormula | null> {
  const prepared: Array<PreparedCellShiftFormula | null> = [];
  for (const candidate of Object.values(cells)) {
    const source = formulaSource(candidate);
    if (source !== null) prepared.push(prepareFormulaForCellShift(source));
  }
  return prepared;
}

export function canAdjustPreparedCellShiftFormulas(
  prepared: ReadonlyArray<PreparedCellShiftFormula | null>,
  range: ShiftCellRange,
  direction: ShiftDirection,
  sheetName?: string,
): boolean {
  const operation: FormulaCellShiftOperation = { range, direction, sheetName };
  return prepared.every((formula) => canAdjustPreparedFormulaForCellShift(formula, operation));
}

function rewriteCellFormula<T extends ShiftFormulaCell>(
  cell: T,
  operation: FormulaCellShiftOperation,
): T {
  const source = formulaSource(cell);
  if (source === null) return cell;
  const result = adjustFormulaForCellShift(source, operation);
  if (!result.supported) {
    throw new Error("Unsupported partial cell-shift formula range");
  }
  if (result.formula === source) return cell;
  const next: ShiftFormulaCell = { ...cell, value: result.formula };
  if (typeof cell.formula === "string") next.formula = result.formula;
  return next as T;
}









export function insertCellsShiftDropsCells(
  cells: Record<string, unknown>,
  range: ShiftCellRange,
  direction: "down" | "right",
  maxRows: number,
  maxCols: number,
): boolean {
  const r = normalize(range);
  const height = r.r2 - r.r1 + 1;
  const width = r.c2 - r.c1 + 1;
  for (const key of Object.keys(cells)) {
    const separator = key.indexOf(":");
    if (separator <= 0) continue;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    const cell = cells[key];
    if (!cell || typeof cell !== "object" || !cellHoldsSheetEndContent(cell)) continue;
    if (direction === "down") {
      if (col >= r.c1 && col <= r.c2 && row >= r.r1 && row + height >= maxRows) return true;
    } else if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col + width >= maxCols) {
      return true;
    }
  }
  return false;
}

/**
 * Insert blank cells at `range`, pushing existing cells DOWN or RIGHT. Only cells
 * on the range's columns (down) or rows (right) move; everything else is copied
 * through unchanged. Returns a NEW map (the input is never mutated). Non-integer
 * keys ride through verbatim.
 *
 * `maxRows`/`maxCols` are the sheet's hard bounds: a cell that would shift past
 * the last valid line is DROPPED (it has nowhere to go), mirroring how the grid
 * cannot hold data outside its declared dimensions. Pass the schema bounds.
 */
export function insertCellsShift<T extends ShiftFormulaCell>(
  cells: Record<string, T>,
  range: ShiftCellRange,
  direction: "down" | "right",
  maxRows: number,
  maxCols: number,
  sheetName?: string,
): Record<string, T> {
  const r = normalize(range);
  const operation: FormulaCellShiftOperation = {
    range: r,
    direction,
    sheetName,
  };
  if (!canAdjustCellShiftFormulas(cells, r, direction, sheetName)) {
    throw new Error("Unsupported partial cell-shift formula range");
  }
  const height = r.r2 - r.r1 + 1;
  const width = r.c2 - r.c1 + 1;
  const out: Record<string, T> = {};
  for (const key of Object.keys(cells)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) {
      out[key] = cells[key];
      continue;
    }
    if (direction === "down") {
      // Same columns as the range, at or below its top: push down by height.
      if (col >= r.c1 && col <= r.c2 && row >= r.r1) {
        const nextRow = row + height;
        if (nextRow >= maxRows) continue; // shifted past the last row -> drop
        out[`${nextRow}:${col}`] = rewriteCellFormula(cells[key], operation);
      } else {
        out[key] = rewriteCellFormula(cells[key], operation);
      }
    } else {
      // direction === "right": same rows, at or right of the left edge.
      if (row >= r.r1 && row <= r.r2 && col >= r.c1) {
        const nextCol = col + width;
        if (nextCol >= maxCols) continue; // shifted past the last column -> drop
        out[`${row}:${nextCol}`] = rewriteCellFormula(cells[key], operation);
      } else {
        out[key] = rewriteCellFormula(cells[key], operation);
      }
    }
  }
  return out;
}

/**
 * Delete the cells in `range`, pulling later cells UP or LEFT to fill the gap.
 * Only cells on the range's columns (up) or rows (left) move; everything else is
 * copied through. The range cells themselves are removed. Returns a NEW map (the
 * input is never mutated). Non-integer keys ride through verbatim.
 */
export function deleteCellsShift<T extends ShiftFormulaCell>(
  cells: Record<string, T>,
  range: ShiftCellRange,
  direction: "up" | "left",
  sheetName?: string,
): Record<string, T> {
  const r = normalize(range);
  const operation: FormulaCellShiftOperation = {
    range: r,
    direction,
    sheetName,
  };
  if (!canAdjustCellShiftFormulas(cells, r, direction, sheetName)) {
    throw new Error("Unsupported partial cell-shift formula range");
  }
  const height = r.r2 - r.r1 + 1;
  const width = r.c2 - r.c1 + 1;
  const out: Record<string, T> = {};
  for (const key of Object.keys(cells)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) {
      out[key] = cells[key];
      continue;
    }
    if (direction === "up") {
      if (col >= r.c1 && col <= r.c2) {
        // On a range column: cells inside the band are deleted; below pull up.
        if (row >= r.r1 && row <= r.r2) continue; // deleted
        if (row > r.r2) {
          out[`${row - height}:${col}`] = rewriteCellFormula(
            cells[key],
            operation,
          );
        } else {
          out[key] = rewriteCellFormula(cells[key], operation); // above the band -> unchanged
        }
      } else {
        out[key] = rewriteCellFormula(cells[key], operation);
      }
    } else {
      // direction === "left": on a range row, cells in the band are deleted; to
      // the right pull left.
      if (row >= r.r1 && row <= r.r2) {
        if (col >= r.c1 && col <= r.c2) continue; // deleted
        if (col > r.c2) {
          out[`${row}:${col - width}`] = rewriteCellFormula(
            cells[key],
            operation,
          );
        } else {
          out[key] = rewriteCellFormula(cells[key], operation); // left of the band -> unchanged
        }
      } else {
        out[key] = rewriteCellFormula(cells[key], operation);
      }
    }
  }
  return out;
}
