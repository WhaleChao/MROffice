/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the spreadsheet Format Painter (Excel/WPS 格式刷).
 *
 * The Format Painter copies one cell's formatting and "paints" it onto a target
 * cell or range. The React component owns the wiring (the toolbar button, the
 * grid mouse-up that triggers a paint, protection enforcement, and the commit
 * through `updateCellStyle`); this module keeps the two delicate parts pure so
 * they can be unit-tested in isolation:
 *
 *  1. The arm/disarm STATE MACHINE (single-use vs sticky), so re-clicking the
 *     button, finishing a single paint, or pressing Esc never leaves the tool
 *     stuck armed.
 *  2. The captured style is a deep SNAPSHOT taken at arm time — later edits to
 *     the source cell must not mutate what gets painted.
 *
 * Replace (not merge) semantics: painting REPLACES the target's whole style with
 * the captured one (Excel's Format Painter overwrites formatting). Since the grid
 * merges style patches in `updateCellStyle` (and prunes falsy keys), we build a
 * patch that explicitly clears every style key the target might already carry and
 * then sets the captured keys on top. The cleared keys are pruned away, leaving
 * the target with exactly the captured style.
 */

import {
  normalizeShadow,
  type SpreadsheetShadow,
} from "@mrnote/core/web/lib/spreadsheet/shape-effects.ts";
import {
  normalizeShapeStyle,
  type SpreadsheetShapeStyle,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import {
  normalizeOfficeShapeTextFont,
  OFFICE_SHAPE_TEXT_FONT_KEYS,
  type OfficeShapeTextFont,
} from "@mrnote/core/web/lib/office-shape-text-font.ts";

/** The cell style object the grid stores (a loose bag of formatting fields). */
export type CellStyleObject = Record<string, unknown>;

/** The painter's arm mode: not armed, single-use (paints once), or sticky. */
export type PainterMode = "off" | "single" | "sticky";

/**
 * The painter's full state: its arm mode plus the captured style snapshot (null
 * when off). When armed, `clipboard` is the deep copy taken at arm time.
 */
export interface FormatPainterState {
  mode: PainterMode;
  clipboard: CellStyleObject | null;
}

/** A normalized worksheet rectangle, inclusive on every edge. */
export interface FormatPainterRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** The addressable dimensions of the current worksheet. */
export interface FormatPainterBounds {
  rows: number;
  columns: number;
}

/**
 * A detached row-major snapshot of every source cell's complete style bag.
 * Rows and style bags are readonly so the pending painter cannot accidentally
 * observe a later source edit through a shared container.
 */
export interface FormatPainterStyleMatrix {
  readonly rows: number;
  readonly columns: number;
  readonly styles: readonly (readonly Readonly<CellStyleObject>[])[];
}

/**
 * The visual-only part of a spreadsheet drawing shape captured by Format
 * Painter. Shape identity, geometry/kind and label CONTENT deliberately never
 * enter the snapshot; the target keeps all of those when this formatting lands.
 * `null` is authoritative absence for optional visual fields, so painting a
 * plain source also clears stale WordArt/shadow styling on the target.
 */
export interface SpreadsheetShapeVisualSnapshot {
  readonly fill: string;
  readonly stroke: string;
  readonly strokeWidth: number;
  readonly textArt: NonNullable<SpreadsheetShapeStyle["textArt"]> | null;
  readonly textSize: number | null;
  readonly textColor: string | null;
  /**
   * The label's font (family, bold, italic, underline, strikethrough). Excel's
   * Format Painter copies a shape's text formatting with its fill and line; an
   * empty object is the authoritative "none", as `null` is for the others.
   */
  readonly textFont: Readonly<OfficeShapeTextFont>;
  readonly shadow: Readonly<SpreadsheetShadow> | null;
}

/** A fully planned visual-only shape paint, detached from source and target. */
export interface SpreadsheetShapeVisualPaintPlan {
  readonly shape: SpreadsheetShapeStyle;
  readonly shadow: SpreadsheetShadow | undefined;
  readonly changed: boolean;
}

/**
 * The only picture visual currently represented by the spreadsheet object
 * model. `null` is authoritative absence, so a shadowless source picture can
 * clear a stale shadow without capturing any picture content or geometry.
 */
export interface SpreadsheetPictureShadowSnapshot {
  readonly shadow: Readonly<SpreadsheetShadow> | null;
}

/** A detached, validated shadow replacement for one target picture. */
export interface SpreadsheetPictureShadowPaintPlan {
  readonly shadow: SpreadsheetShadow | undefined;
  readonly changed: boolean;
}

/** The disarmed resting state. */
export const FORMAT_PAINTER_OFF: FormatPainterState = {
  mode: "off",
  clipboard: null,
};

/**
 * Every style key the Format Painter knows how to clear, so a paint REPLACES the
 * target style instead of merging onto it. Mirrors the `style` shape on a
 * spreadsheet cell (lib/spreadsheet cell model + the toolbar's controls). Keeping
 * this list explicit (rather than diffing the live target) keeps the produced
 * patch deterministic and independent of whatever the target happens to carry.
 */
export const PAINTABLE_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "color",
  "backgroundColor",
  "fillPattern",
  "numberFormat",
  "border",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
  "textAlign",
  "verticalAlign",
  "wrapText",
] as const;

function cloneStyleValue<T>(value: T): T {
  return value !== null && typeof value === "object"
    ? structuredClone(value)
    : value;
}

/**
 * Deep-copy a style object so the painter's clipboard is a SNAPSHOT, fully
 * detached from the source cell. Later edits to the source (or its style object)
 * must not change what a pending paint will apply. A null/undefined style (an
 * unstyled source cell) snapshots to an empty object so painting it sanely clears
 * the target rather than crashing.
 *
 * Style values are JSON-safe primitives (booleans / strings), so a structured
 * clone of the own enumerable keys is a faithful, independent copy.
 */
export function snapshotStyle(
  style: Readonly<CellStyleObject> | null | undefined,
): CellStyleObject {
  const snapshot: CellStyleObject = {};
  if (!style) return snapshot;
  for (const key of Object.keys(style)) {
    const value = style[key];
    // Skip undefined so an explicitly-undefined key does not masquerade as set.
    if (value === undefined) continue;
    snapshot[key] = cloneStyleValue(value);
  }
  return snapshot;
}

/**
 * Capture a drawing shape's complete visual formatting at arm time. The shape
 * normalizer is the same authority used by the worksheet renderer/editor, so a
 * legacy partial shape produces the exact visible defaults rather than leaking
 * malformed fields into the pending painter. Every nested value is detached.
 */
export function snapshotSpreadsheetShapeVisual(
  shape: SpreadsheetShapeStyle | null | undefined,
  shadow: SpreadsheetShadow | null | undefined,
): Readonly<SpreadsheetShapeVisualSnapshot> {
  const normalizedShape = normalizeShapeStyle(shape);
  const normalizedShadow = normalizeShadow(shadow);
  return Object.freeze({
    fill: normalizedShape.fill,
    stroke: normalizedShape.stroke,
    strokeWidth: normalizedShape.strokeWidth,
    textArt: normalizedShape.textArt ?? null,
    textSize: normalizedShape.textSize ?? null,
    textColor: normalizedShape.textColor ?? null,
    textFont: Object.freeze(normalizeOfficeShapeTextFont(normalizedShape)),
    shadow: normalizedShadow
      ? Object.freeze({ ...normalizedShadow })
      : null,
  });
}

/** Capture only a picture's normalized shadow, detached at arm time. */
export function snapshotSpreadsheetPictureShadow(
  shadow: SpreadsheetShadow | null | undefined,
): Readonly<SpreadsheetPictureShadowSnapshot> {
  const normalizedShadow = normalizeShadow(shadow);
  return Object.freeze({
    shadow: normalizedShadow
      ? Object.freeze({ ...normalizedShadow })
      : null,
  });
}

function isNormalizedRange(range: FormatPainterRange): boolean {
  return (
    Number.isSafeInteger(range.r1) &&
    Number.isSafeInteger(range.c1) &&
    Number.isSafeInteger(range.r2) &&
    Number.isSafeInteger(range.c2) &&
    range.r1 >= 0 &&
    range.c1 >= 0 &&
    range.r2 >= range.r1 &&
    range.c2 >= range.c1
  );
}

function isValidBounds(bounds: FormatPainterBounds): boolean {
  return (
    Number.isSafeInteger(bounds.rows) &&
    Number.isSafeInteger(bounds.columns) &&
    bounds.rows > 0 &&
    bounds.columns > 0
  );
}

function isStyleBag(value: unknown): value is Readonly<CellStyleObject> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isValidStyleMatrix(
  matrix: FormatPainterStyleMatrix | null | undefined,
): matrix is FormatPainterStyleMatrix {
  return Boolean(
    matrix &&
    Number.isSafeInteger(matrix.rows) &&
    Number.isSafeInteger(matrix.columns) &&
    matrix.rows > 0 &&
    matrix.columns > 0 &&
    matrix.styles.length === matrix.rows &&
    matrix.styles.every(
      (row) =>
        Array.isArray(row) &&
        row.length === matrix.columns &&
        row.every(isStyleBag),
    ),
  );
}

/**
 * Capture a complete rectangular style matrix. The callback is read exactly
 * once per source coordinate; every returned style is deeply detached through
 * `snapshotStyle`, so overlapping/later paints never re-read live source data.
 */
export function snapshotStyleMatrix(
  range: FormatPainterRange,
  bounds: FormatPainterBounds,
  styleAt: (row: number, column: number) => CellStyleObject | null | undefined,
): FormatPainterStyleMatrix | null {
  if (
    !isNormalizedRange(range) ||
    !isValidBounds(bounds) ||
    range.r2 >= bounds.rows ||
    range.c2 >= bounds.columns
  ) {
    return null;
  }
  const rows = range.r2 - range.r1 + 1;
  const columns = range.c2 - range.c1 + 1;
  if (
    !Number.isSafeInteger(rows) ||
    !Number.isSafeInteger(columns)
  ) {
    return null;
  }
  try {
    const styles = Array.from({ length: rows }, (_, rowOffset) =>
      Object.freeze(
        Array.from({ length: columns }, (_, columnOffset) =>
          Object.freeze(
            snapshotStyle(
              styleAt(range.r1 + rowOffset, range.c1 + columnOffset),
            ),
          ),
        ),
      ),
    );
    return Object.freeze({
      rows,
      columns,
      styles: Object.freeze(styles),
    });
  } catch {
    // A malformed/non-cloneable style must not arm a partially captured matrix.
    return null;
  }
}

/**
 * Resolve the rectangle a matrix paint will cover.
 *
 * - A one-cell target expands to the source matrix dimensions.
 * - Every other target must be an exact positive-integer tiling of the source.
 * - Malformed/ragged matrices, shape mismatches, and worksheet overflow fail
 *   closed before callers begin constructing a worksheet mutation.
 */
export function planFormatPainterDestination(
  matrix: FormatPainterStyleMatrix | null | undefined,
  target: FormatPainterRange,
  bounds: FormatPainterBounds,
): FormatPainterRange | null {
  if (
    !isValidStyleMatrix(matrix) ||
    !isNormalizedRange(target) ||
    !isValidBounds(bounds) ||
    target.r1 >= bounds.rows ||
    target.c1 >= bounds.columns
  ) {
    return null;
  }
  const targetRows = target.r2 - target.r1 + 1;
  const targetColumns = target.c2 - target.c1 + 1;
  if (
    !Number.isSafeInteger(targetRows) ||
    !Number.isSafeInteger(targetColumns)
  ) {
    return null;
  }
  const isOneCellAnchor = targetRows === 1 && targetColumns === 1;
  const rows = isOneCellAnchor ? matrix.rows : targetRows;
  const columns = isOneCellAnchor ? matrix.columns : targetColumns;
  if (
    rows % matrix.rows !== 0 ||
    columns % matrix.columns !== 0 ||
    target.r1 + rows > bounds.rows ||
    target.c1 + columns > bounds.columns
  ) {
    return null;
  }
  return {
    r1: target.r1,
    c1: target.c1,
    r2: target.r1 + rows - 1,
    c2: target.c1 + columns - 1,
  };
}

/** Return a fresh style bag for one tiled destination offset, or null if stale. */
export function formatPainterStyleAt(
  matrix: FormatPainterStyleMatrix | null | undefined,
  rowOffset: number,
  columnOffset: number,
): CellStyleObject | null {
  if (
    !matrix ||
    !Number.isSafeInteger(matrix.rows) ||
    !Number.isSafeInteger(matrix.columns) ||
    matrix.rows <= 0 ||
    matrix.columns <= 0 ||
    !Number.isSafeInteger(rowOffset) ||
    !Number.isSafeInteger(columnOffset) ||
    rowOffset < 0 ||
    columnOffset < 0
  ) {
    return null;
  }
  const row = matrix.styles[rowOffset % matrix.rows];
  const style = row?.[columnOffset % matrix.columns];
  if (!isStyleBag(style)) return null;
  try {
    return snapshotStyle(style);
  } catch {
    return null;
  }
}

function styleValueEquals(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) || Array.isArray(right)) {
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((value, index) => styleValueEquals(value, right[index]))
    );
  }
  if (
    left === null ||
    right === null ||
    typeof left !== "object" ||
    typeof right !== "object"
  ) {
    return false;
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).filter(
    (key) => leftRecord[key] !== undefined,
  );
  const rightKeys = Object.keys(rightRecord).filter(
    (key) => rightRecord[key] !== undefined,
  );
  return (
    leftKeys.length === rightKeys.length &&
    leftKeys.every(
      (key) =>
        Object.prototype.hasOwnProperty.call(rightRecord, key) &&
        styleValueEquals(leftRecord[key], rightRecord[key]),
    )
  );
}

function normalizeShapeVisualSnapshot(
  snapshot: Readonly<SpreadsheetShapeVisualSnapshot> | null | undefined,
): {
  fill: string;
  stroke: string;
  strokeWidth: number;
  textArt: NonNullable<SpreadsheetShapeStyle["textArt"]> | null;
  textSize: number | null;
  textColor: string | null;
  textFont: OfficeShapeTextFont;
  shadow: SpreadsheetShadow | undefined;
} | null {
  if (
    !snapshot ||
    typeof snapshot.fill !== "string" ||
    typeof snapshot.stroke !== "string" ||
    typeof snapshot.strokeWidth !== "number" ||
    !Number.isFinite(snapshot.strokeWidth)
  ) {
    return null;
  }
  const normalizedShape = normalizeShapeStyle({
    kind: "rectangle",
    text: "format painter",
    fill: snapshot.fill,
    stroke: snapshot.stroke,
    strokeWidth: snapshot.strokeWidth,
    ...(snapshot.textArt !== null ? { textArt: snapshot.textArt } : {}),
    ...(snapshot.textSize !== null ? { textSize: snapshot.textSize } : {}),
    ...(snapshot.textColor !== null ? { textColor: snapshot.textColor } : {}),
    ...(snapshot.textFont ?? {}),
  });
  const textFont = normalizeOfficeShapeTextFont(normalizedShape);
  if (
    normalizedShape.fill !== snapshot.fill ||
    normalizedShape.stroke !== snapshot.stroke ||
    normalizedShape.strokeWidth !== snapshot.strokeWidth ||
    (normalizedShape.textArt ?? null) !== snapshot.textArt ||
    (normalizedShape.textSize ?? null) !== snapshot.textSize ||
    (normalizedShape.textColor ?? null) !== snapshot.textColor ||
    !styleValueEquals(textFont, snapshot.textFont ?? {})
  ) {
    return null;
  }
  const shadow =
    snapshot.shadow === null ? undefined : normalizeShadow(snapshot.shadow);
  if (
    snapshot.shadow !== null &&
    (!shadow || !styleValueEquals(snapshot.shadow, shadow))
  ) {
    return null;
  }
  return {
    fill: normalizedShape.fill,
    stroke: normalizedShape.stroke,
    strokeWidth: normalizedShape.strokeWidth,
    textArt: normalizedShape.textArt ?? null,
    textSize: normalizedShape.textSize ?? null,
    textColor: normalizedShape.textColor ?? null,
    textFont,
    shadow,
  };
}

/**
 * Resolve a detached shape-visual snapshot against a live target. Only the
 * declared visual keys are replaced: target kind/text and any future nonvisual
 * shape metadata survive byte-for-byte. Optional source fields are authoritative
 * clears. A malformed snapshot fails closed before the caller constructs a
 * worksheet mutation.
 */
export function planSpreadsheetShapeVisualPaint(
  snapshot: Readonly<SpreadsheetShapeVisualSnapshot> | null | undefined,
  targetShape: SpreadsheetShapeStyle | null | undefined,
  targetShadow: SpreadsheetShadow | null | undefined,
): SpreadsheetShapeVisualPaintPlan | null {
  const visual = normalizeShapeVisualSnapshot(snapshot);
  if (!visual) return null;

  const normalizedTarget = normalizeShapeStyle(targetShape);
  const nextShapeDraft = {
    ...(targetShape && typeof targetShape === "object" ? targetShape : {}),
    kind: normalizedTarget.kind,
    fill: visual.fill,
    stroke: visual.stroke,
    strokeWidth: visual.strokeWidth,
    ...(normalizedTarget.text ? { text: normalizedTarget.text } : {}),
  } as SpreadsheetShapeStyle;
  if (!normalizedTarget.text) delete nextShapeDraft.text;
  if (normalizedTarget.text && visual.textArt !== null) {
    nextShapeDraft.textArt = visual.textArt;
  } else {
    delete nextShapeDraft.textArt;
  }
  if (normalizedTarget.text && visual.textSize !== null) {
    nextShapeDraft.textSize = visual.textSize;
  } else {
    delete nextShapeDraft.textSize;
  }
  if (normalizedTarget.text && visual.textColor !== null) {
    nextShapeDraft.textColor = visual.textColor;
  } else {
    delete nextShapeDraft.textColor;
  }
  // The font replaces the target's whole font: a key the source does not
  // state is cleared, like every other optional field here.
  for (const key of OFFICE_SHAPE_TEXT_FONT_KEYS) delete nextShapeDraft[key];
  Object.assign(nextShapeDraft, visual.textFont);
  // Re-run the target-kind normalizer after applying the visual fields. It
  // preserves the target's semantic kind/text while enforcing invariants such
  // as line/arrow shapes having no interior fill.
  const nextShape = {
    ...nextShapeDraft,
    ...normalizeShapeStyle(nextShapeDraft),
  } as SpreadsheetShapeStyle;

  const nextShadow = visual.shadow ? { ...visual.shadow } : undefined;
  return {
    shape: nextShape,
    shadow: nextShadow,
    changed:
      !styleValueEquals(targetShape, nextShape) ||
      !styleValueEquals(targetShadow, nextShadow),
  };
}

/**
 * Validate a detached picture-shadow snapshot against a live target shadow.
 * The result contains no picture identity, source, metadata, or geometry, so
 * the component can replace exactly this one visual field and nothing else.
 */
export function planSpreadsheetPictureShadowPaint(
  snapshot: Readonly<SpreadsheetPictureShadowSnapshot> | null | undefined,
  targetShadow: SpreadsheetShadow | null | undefined,
): SpreadsheetPictureShadowPaintPlan | null {
  if (!snapshot) return null;
  const shadow =
    snapshot.shadow === null ? undefined : normalizeShadow(snapshot.shadow);
  if (
    snapshot.shadow !== null &&
    (!shadow || !styleValueEquals(snapshot.shadow, shadow))
  ) {
    return null;
  }
  const nextShadow = shadow ? { ...shadow } : undefined;
  return {
    shadow: nextShadow,
    changed: !styleValueEquals(targetShadow, nextShadow),
  };
}

/** Semantic style equality (object-key order and undefined keys are irrelevant). */
export function formatPainterStylesEqual(
  left: CellStyleObject | null | undefined,
  right: CellStyleObject | null | undefined,
): boolean {
  const normalizedLeft =
    left && Object.keys(left).length > 0 ? left : undefined;
  const normalizedRight =
    right && Object.keys(right).length > 0 ? right : undefined;
  return styleValueEquals(normalizedLeft, normalizedRight);
}

/**
 * Build the single per-cell style patch that REPLACES a target's style with the
 * captured one. Every known style key is first set to a falsy sentinel (so
 * `updateCellStyle` prunes it from the target), then the captured keys are laid
 * on top. Painting an empty/unstyled snapshot therefore CLEARS the target's
 * formatting (Excel parity), and painting a styled snapshot overwrites it wholly
 * — no stale field from the previous formatting survives.
 *
 * The patch is a fresh object on every call and shares no references with the
 * clipboard, so the caller can apply it to many cells without aliasing.
 */
export function buildPaintPatch(
  clipboard: CellStyleObject | null,
): Record<string, unknown> {
  const patch: Record<string, unknown> = {};
  // Clear every paintable key first (false is pruned by updateCellStyle).
  for (const key of PAINTABLE_STYLE_KEYS) {
    patch[key] = false;
  }
  if (!clipboard) return patch;
  // Lay the captured keys on top. Copy values (primitives) so the patch never
  // aliases the clipboard. Unknown extra keys captured from the source are
  // carried through too, so a custom style field still round-trips.
  for (const key of Object.keys(clipboard)) {
    const value = clipboard[key];
    if (value === undefined) continue;
    patch[key] = cloneStyleValue(value);
  }
  return patch;
}

/**
 * The arm transition for a single click on the painter button.
 *  - When OFF: arm in single-use mode, capturing the given source-style snapshot.
 *  - When ALREADY ARMED (single or sticky): a click DISARMS (toggle off), so the
 *    button never gets stuck — clicking an armed painter cancels it.
 *
 * `sourceStyle` is snapshotted here (deep copy) so the returned clipboard is
 * detached from the live source cell.
 */
export function armSingle(
  current: FormatPainterState,
  sourceStyle: Readonly<CellStyleObject> | null | undefined,
): FormatPainterState {
  if (current.mode !== "off") return FORMAT_PAINTER_OFF;
  return { mode: "single", clipboard: snapshotStyle(sourceStyle) };
}

/**
 * The arm transition for a double-click on the painter button: STICKY mode. It
 * captures (or re-captures) the source snapshot and stays armed across paints
 * until Esc / another click disarms it. A double-click always arms sticky (even
 * from an already-single-armed state) so the gesture is predictable.
 */
export function armSticky(
  _current: FormatPainterState,
  sourceStyle: Readonly<CellStyleObject> | null | undefined,
): FormatPainterState {
  return { mode: "sticky", clipboard: snapshotStyle(sourceStyle) };
}

/**
 * The post-paint transition. After a paint lands:
 *  - single-use: disarm (one shot, then off).
 *  - sticky: STAY armed with the same clipboard (repeat painting).
 *  - off: stays off (defensive — a paint should not fire while off).
 */
export function afterPaint(current: FormatPainterState): FormatPainterState {
  if (current.mode === "sticky") return current;
  return FORMAT_PAINTER_OFF;
}

/** True when the painter is armed (single or sticky) and ready to paint. */
export function isArmed(state: FormatPainterState): boolean {
  return state.mode !== "off" && state.clipboard !== null;
}
