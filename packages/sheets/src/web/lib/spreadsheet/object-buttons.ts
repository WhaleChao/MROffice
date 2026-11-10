/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pictures and shapes on a worksheet: the drawing objects' geometry (anchor
 * cells, offsets and sizes in CSS pixels) and the Insert, Shape Format and
 * Picture Format buttons, as pure functions over a sheet's `objects[]`.
 * Moved out of SpreadsheetWindow so the window's buttons and the spreadsheet
 * editor commands (lib/editor-commands/spreadsheet) run the same code;
 * behaviour is unchanged. Receipts, protection notices, selection, undo and
 * the contextual tab stay with the caller.
 *
 * A button returns the patch the window merges into the object
 * (`{ ...object, ...patch }`), or null when it would change nothing or does
 * not apply to that object.
 */


import { canonicalizeDrawingMlCropRect, clampCropRect, isIdentityCrop, normalizeSignedCropRect, sameCropRect, type CropRect } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { readOfficeLineEnds } from "@mrnote/core/web/lib/office-line-ends.ts";
import { normalizeOfficeTextWarp } from "@mrnote/core/web/lib/office-text-warp.ts";
import { officeTextLayoutFollowingText } from "@mrnote/core/web/lib/office-text-layout.ts";
import { OFFICE_MENU_DASHES, officeMenuDashFields, type OfficeMenuDash } from "@mrnote/core/web/lib/office-stroke-dash.ts";
import { normalizeOfficePictureEffects, withOfficePictureEffect } from "@mrnote/core/web/lib/office-picture-effects.ts";
import type { OfficeShadowEffect } from "@mrnote/core/web/lib/office-shadow-effects.ts";
import { spreadsheetShapeOfficeEffects } from "./shape-office-effects";
import {
  OFFICE_OUTLINE_SHOWN_WEIGHT_PT,
  officeLineWeightPtToPx,
  officeLineWeightPxLadder,
  officeOutlineWeightWhenShown,
} from "@mrnote/core/web/lib/office-line-weight.ts";
import { isOfficePictureStyleId } from "@mrnote/core/web/lib/office-picture-styles/catalog.ts";
import { normalizeOfficeShapeEffectPreset, OFFICE_SHAPE_EFFECT_PRESETS } from "@mrnote/core/web/lib/office-shape-effects.ts";
import { normalizeOfficeTextPaint, patchOfficeTextPaint, sameOfficeTextPaint, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import { readSmartArtGroupMembers } from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import { chartDrawingGroupId } from "./chart-groups";
import type { OfficeSmartArtHostNode, SmartArtBlockMarker } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { officeOutlinePtToPx, type OfficeDefaultShapeStyle } from "@mrnote/core/web/lib/office-default-shape-style.ts";
import type { OfficeShapeStyleRecipe } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import type { OfficeWordArtRecipe } from "@mrnote/core/web/lib/office-wordart-style-catalog.ts";
import { smartArtConnectorsFollowNodes } from "@mrnote/core/web/lib/office-smartart/connector-follow.ts";
import { readSmartArtCustom, smartArtLayoutBox, smartArtMemberAfterBoxEdit } from "@mrnote/core/web/lib/office-smartart/custom-geometry.ts";
import { normalizeHexColor } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import { recomputeSpreadsheetGluedConnectors, spreadsheetConnectorGlueAfterMove } from "./connector-glue";
import {
  applyDrawingStack,
  moveInDrawingStack,
  sheetDrawingStack,
  type DrawingStackEntry,
  type DrawingStackKind,
  type DrawingStackMove,
} from "./drawing-stack";
import {
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
  DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  offsetOfLine,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { DEFAULT_SHADOW_COLOR, normalizeShadow, SPREADSHEET_SHADOW_PRESETS } from "@mrnote/core/web/lib/spreadsheet/shape-effects.ts";
import { spreadsheetObjectOffsetsLayoutRelative } from "@mrnote/core/web/lib/spreadsheet/layout-relative-anchor.ts";
import { spreadsheetShapeLineEndsPatch } from "./shape-line-ends";
import {
  createSpreadsheetShapeObject,
  DEFAULT_SHAPE_STROKE,
  isSpreadsheetLineLikeShapeKind,
  isSpreadsheetPresetShapeKind,
  MAX_SHAPE_STROKE_WIDTH,
  MAX_SHAPE_TEXT_LENGTH,
  MAX_SHAPE_TEXT_SIZE,
  MIN_SHAPE_TEXT_SIZE,
  normalizeShapeStyle,
  spreadsheetShapeHasPictureFill,
  spreadsheetShapeStyleWithoutPictureFill,
  type SpreadsheetPresetShapeKind,
  type SpreadsheetShapeStyle,
  type SpreadsheetShapeObject,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import { spreadsheetEquationBoxLatex, withSpreadsheetEquationBoxBody } from "./equation-box";
import { officePictureFillSrc } from "@mrnote/core/web/lib/office-picture-fill.ts";
import { normalizeOfficePictureFillOptions, officePictureFillOptionsStated } from "@mrnote/core/web/lib/office-picture-fill-options.ts";
import {
  normalizeSpreadsheetShapeTextRuns,
  spreadsheetShapeTextRunsClipped,
  spreadsheetShapeTextRunsFilled,
  spreadsheetShapeTextRunsFollowing,
  spreadsheetShapeTextRunsWithout,
} from "@mrnote/core/web/lib/spreadsheet/shape-text-runs.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";
import { officeObjectHasImportIdentity } from "@mrnote/core/web/lib/office-object-provenance.ts";
import { smartArtSlotPictureFormatOf } from "@mrnote/core/web/lib/office-smartart/slot-picture-format.ts";
import type { SpreadsheetThemeShapeVisual } from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

/** What every drawing object carries that the geometry reads. */
export interface DrawingObjectLike {
  id: string;
  type?: string;
  anchor?: SpreadsheetObjectAnchor;
  source?: unknown;
}

export interface SpreadsheetAnchorPoint {
  row: number;
  column?: number;
  col?: number;
  rowOffset?: number;
  columnOffset?: number;
  colOffset?: number;
}

export interface SpreadsheetObjectAnchor {
  kind?: "oneCell" | "twoCell" | "absolute";
  from?: SpreadsheetAnchorPoint;
  to?: SpreadsheetAnchorPoint;
  xEmu?: number;
  yEmu?: number;
  widthEmu?: number;
  heightEmu?: number;
  /**
   * Where an absolute position is measured from: "sheet" is the worksheet's
   * own origin (cell A1's corner, an imported `xdr:absoluteAnchor`); absent,
   * the drawing layer's, which also holds the headings (objects drawn here).
   */
  origin?: "sheet";
  /** DrawingML xfrm rotation projected as user-facing clockwise degrees. */
  rotation?: number;
  /** DrawingML xfrm mirror flags; shape text remains readable/upright. */
  flipH?: boolean;
  flipV?: boolean;
  /**
   * A Move-but-don't-size or Don't-move-or-size placement chosen in
   * Properties (equal to `kind`): the export writes it as Excel does, a
   * `twoCellAnchor` with this `editAs` on the box's two markers. Absent, an
   * anchor read from a file keeps the element it came in.
   */
  editAs?: "oneCell" | "absolute";
  [key: string]: unknown;
}

export interface SpreadsheetObjectBounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

export function anchorPointColumn(point: SpreadsheetAnchorPoint | undefined): number {
  return Math.max(0, Number(point?.column ?? point?.col ?? 0));
}

export function anchorPointRow(point: SpreadsheetAnchorPoint | undefined): number {
  return Math.max(0, Number(point?.row ?? 0));
}

export function anchorPointOffsetPixels(
  point: SpreadsheetAnchorPoint | undefined,
  axis: "row" | "column",
): number {
  const raw =
    axis === "row"
      ? point?.rowOffset
      : point?.columnOffset ?? point?.colOffset;
  const emu = Number(raw ?? 0);
  // A negative offset is a frame turned a quarter about an Excel box near
  // the sheet's first row or column, starting before it (X-166).
  return Number.isFinite(emu) && emu !== 0 ? emu / EMU_PER_CSS_PIXEL : 0;
}









export function anchorPointOffsetInLinePixels(
  point: SpreadsheetAnchorPoint | undefined,
  axis: "row" | "column",
  lineSize: number,
): number {
  return Math.min(anchorPointOffsetPixels(point, axis), Math.max(0, lineSize));
}

export const EMU_PER_CSS_PIXEL = 9525;

export const DEFAULT_OBJECT_WIDTH_PX = 120;

export const DEFAULT_OBJECT_HEIGHT_PX = 84;

export function emuToCssPixels(value: number | undefined): number | undefined {
  if (!Number.isFinite(value) || value === undefined || value <= 0)
    return undefined;
  return Math.max(1, Math.round(value / EMU_PER_CSS_PIXEL));
}

export function cssPixelsToEmu(value: number): number {
  return Math.max(1, Math.round(value * EMU_PER_CSS_PIXEL));
}

export function positiveInteger(value: string, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? Math.max(1, Math.floor(parsed)) : fallback;
}

export function anchorPointWithDefaults(
  point: SpreadsheetAnchorPoint | undefined,
): SpreadsheetAnchorPoint {
  return {
    ...(point || {}),
    row: anchorPointRow(point),
    column: anchorPointColumn(point),
  };
}

/**
 * `anchor` at an explicit size, its top-left corner kept. An absolute anchor
 * stays absolute at its own x/y (and origin): turned into a one-cell anchor it
 * lost its position, which it has no `from` cell for, and every object drawn
 * with Insert > Shapes jumped to the headings' corner on a Size edit (user
 * audit S-008). A cell anchor keeps its `from` cell and offsets.
 */
export function anchorWithExplicitSize(
  anchor: SpreadsheetObjectAnchor,
  size: Pick<SpreadsheetObjectAnchor, "widthEmu" | "heightEmu">,
): SpreadsheetObjectAnchor {
  if (anchor.kind === "absolute") return { ...anchor, ...size };
  const rest = { ...anchor };
  delete rest.to;
  return {
    ...rest,
    kind: "oneCell",
    ...size,
  };
}

export function anchorWithDefaults(
  anchor: SpreadsheetObjectAnchor | undefined,
): SpreadsheetObjectAnchor {
  return {
    ...(anchor || {}),
    kind: anchor?.kind || "oneCell",
    from: anchorPointWithDefaults(anchor?.from),
  };
}

export function objectGridSize(object: DrawingObjectLike): {
  width: number;
  height: number;
};
export function objectGridSize(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): { width: number; height: number };
export function objectGridSize(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number = () =>
    DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  rowHeight: (row: number) => number = () =>
    DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
): { width: number; height: number } {
  const from = object.anchor?.from;
  const to = object.anchor?.to;
  if (object.anchor?.kind === "twoCell" && to) {
    const fromRow = anchorPointRow(from);
    const fromCol = anchorPointColumn(from);
    const toRow = Math.max(fromRow, anchorPointRow(to));
    const toCol = Math.max(fromCol, anchorPointColumn(to));
    const fromX =
      offsetOfLine(fromCol, fromCol, columnWidth) +
      anchorPointOffsetInLinePixels(from, "column", columnWidth(fromCol));
    const fromY =
      offsetOfLine(fromRow, fromRow, rowHeight) +
      anchorPointOffsetInLinePixels(from, "row", rowHeight(fromRow));
    const toX =
      offsetOfLine(toCol, toCol, columnWidth) +
      anchorPointOffsetInLinePixels(to, "column", columnWidth(toCol));
    const toY =
      offsetOfLine(toRow, toRow, rowHeight) +
      anchorPointOffsetInLinePixels(to, "row", rowHeight(toRow));
    return {
      // A two-cell marker is a boundary coordinate, not an inclusive cell.
      // Excel therefore sizes A1:C3 from the leading edge of A1 to the leading
      // edge of C3 (plus offsets), rather than adding C3's full width/height.
      // With every line it spans hidden it is 0, and Excel draws nothing
      // (spreadsheetObjectCollapsedByCells).
      width: Math.max(0, toX - fromX),
      height: Math.max(0, toY - fromY),
    };
  }
  const fromRow = anchorPointRow(from);
  const fromCol = anchorPointColumn(from);
  const toRow = Math.max(fromRow, anchorPointRow(to));
  const toCol = Math.max(fromCol, anchorPointColumn(to));
  return {
    width: Math.max(
      DEFAULT_OBJECT_WIDTH_PX,
      (toCol - fromCol + 1) * DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
    ),
    height: Math.max(
      DEFAULT_OBJECT_HEIGHT_PX,
      (toRow - fromRow + 1) * DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
    ),
  };
}

export function objectSizePixels(object: DrawingObjectLike): {
  width: number;
  height: number;
};
export function objectSizePixels(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): { width: number; height: number };
export function objectSizePixels(
  object: DrawingObjectLike,
  columnWidth?: (column: number) => number,
  rowHeight?: (row: number) => number,
): { width: number; height: number } {
  const gridSize =
    columnWidth && rowHeight
      ? objectGridSize(object, columnWidth, rowHeight)
      : objectGridSize(object);
  // A stroke of ink keeps its box's fractions: its points are shares of the
  // box, and a box rounded to whole pixels stretched each stroke its own way
  
  const exact = (object as { shape?: { kind?: unknown } }).shape?.kind === "ink";
  const size = (emu: number | undefined) =>
    // An extent stated as 0 is a line along the other axis (a connector
    // with `cy="0"`, NP013's L03): one pixel, not the grid's default size,
    // which drew it as a diagonal (X-331).
    emu === 0
      ? 1
      : exact && emu !== undefined && Number.isFinite(emu) && emu > 0
      ? Math.max(1, emu / EMU_PER_CSS_PIXEL)
      : emuToCssPixels(emu);
  // Move and size with cells: Excel sizes the object from its two markers
  // alone (an extent stated beside them is never read, measured on Excel for
  
  // (A collapsed one is not drawn; its frame keeps a pixel for the editing
  // arithmetic that divides by it.)
  if (object.anchor?.kind === "twoCell" && object.anchor.to && columnWidth && rowHeight) {
    return { width: Math.max(1, gridSize.width), height: Math.max(1, gridSize.height) };
  }
  return {
    width: size(object.anchor?.widthEmu) ?? gridSize.width,
    height: size(object.anchor?.heightEmu) ?? gridSize.height,
  };
}

export function objectBoundsPixels(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds {
  // A SmartArt member is drawn inside its graphic's frame, which is the one
  // anchored object Excel has (lib/spreadsheet SmartArt frame, below).
  const frame = readSmartArtFrame(object.anchor);
  if (frame) return smartArtMemberBounds(object, frame, columnWidth, rowHeight);
  return objectBoundsOnGrid(object, columnWidth, rowHeight);
}

// ---------------------------------------------------------------------------
// The SmartArt graphic frame.
//
// In Excel a SmartArt graphic is ONE anchored object, an `xdr:graphicFrame`
// in a two-cell anchor (Move and size with cells, as Excel writes it): rows
// and columns inserted, deleted, resized or hidden inside its span stretch
// or move the whole graphic, and its shapes are laid out inside the new
// frame (ECMA-376 §20.5.2.16; Microsoft Support, "Move or size cells with
// objects"). The sheet keeps the graphic's shapes as members anchored to the
// frame's corner cell with their offsets in the graphic's layout, and the
// frame on each member's anchor (`smartArtFrame`): its own anchor, plus the
// size the members were laid out at. A member is drawn where its place in
// that layout falls in the frame as the grid now sizes it.

/** A SmartArt graphic's frame: its anchor and the size its shapes were laid out at. */
export type SpreadsheetSmartArtFrame = SpreadsheetObjectAnchor & {
  designWidthEmu: number;
  designHeightEmu: number;
};

/** The frame a member's anchor carries, or null (no frame: drawn as anchored). */
export function readSmartArtFrame(anchor: SpreadsheetObjectAnchor | undefined): SpreadsheetSmartArtFrame | null {
  const raw = anchor?.smartArtFrame;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const frame = raw as SpreadsheetSmartArtFrame;
  const width = Number(frame.designWidthEmu);
  const height = Number(frame.designHeightEmu);
  if (!(width > 0) || !(height > 0) || !Number.isFinite(width) || !Number.isFinite(height)) return null;
  if (frame.kind === "twoCell" ? !frame.from || !frame.to : frame.kind === "oneCell" ? !frame.from : frame.kind !== "absolute") {
    return null;
  }
  return frame;
}

/** The frame's anchor alone (what Excel's graphic frame states). */
export function smartArtFrameAnchor(frame: SpreadsheetSmartArtFrame): SpreadsheetObjectAnchor {
  return smartArtFrameAnchorOnly(frame);
}

/** Where the frame is drawn on this grid. */
export function smartArtFrameBounds(
  frame: SpreadsheetSmartArtFrame,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds {
  return objectBoundsOnGrid({ id: "", anchor: smartArtFrameAnchor(frame) }, columnWidth, rowHeight);
}

/** A point's place on the grid with its offsets as stated (not clamped). */
function rawTopLeft(
  anchor: SpreadsheetObjectAnchor,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): { left: number; top: number } {
  if (anchor.kind === "absolute") {
    const sheetOrigin = anchor.origin === "sheet";
    return {
      left: (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_WIDTH_PX : 0) + Math.max(0, Number(anchor.xEmu ?? 0) / EMU_PER_CSS_PIXEL),
      top: (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX : 0) + Math.max(0, Number(anchor.yEmu ?? 0) / EMU_PER_CSS_PIXEL),
    };
  }
  const from = anchor.from;
  const row = anchorPointRow(from);
  const column = anchorPointColumn(from);
  return {
    left: DEFAULT_SPREADSHEET_HEADER_WIDTH_PX + offsetOfLine(column, column, columnWidth) + anchorPointOffsetPixels(from, "column"),
    top: DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX + offsetOfLine(row, row, rowHeight) + anchorPointOffsetPixels(from, "row"),
  };
}

/** A member's box in its graphic's layout: offset from the frame's corner, and size (px at the design size). */
function smartArtMemberLayoutBox(
  anchor: SpreadsheetObjectAnchor,
  frame: SpreadsheetSmartArtFrame,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds {
  const member = rawTopLeft(anchor, columnWidth, rowHeight);
  const corner = rawTopLeft(smartArtFrameAnchor(frame), columnWidth, rowHeight);
  return {
    left: member.left - corner.left,
    top: member.top - corner.top,
    width: Math.max(1, Number(anchor.widthEmu ?? EMU_PER_CSS_PIXEL) / EMU_PER_CSS_PIXEL),
    height: Math.max(1, Number(anchor.heightEmu ?? EMU_PER_CSS_PIXEL) / EMU_PER_CSS_PIXEL),
  };
}

function smartArtMemberBounds(
  object: DrawingObjectLike,
  frame: SpreadsheetSmartArtFrame,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds {
  const live = smartArtFrameBounds(frame, columnWidth, rowHeight);
  const box = smartArtMemberLayoutBox(object.anchor!, frame, columnWidth, rowHeight);
  const scaleX = live.width / (frame.designWidthEmu / EMU_PER_CSS_PIXEL);
  const scaleY = live.height / (frame.designHeightEmu / EMU_PER_CSS_PIXEL);
  return {
    left: live.left + box.left * scaleX,
    top: live.top + box.top * scaleY,
    width: Math.max(0, box.width * scaleX),
    height: Math.max(0, box.height * scaleY),
  };
}

/**
 * A member's anchor for the box it is now drawn in (`bounds`, drawing-layer
 * px), its graphic's frame where it is: the box's place in the frame, as the
 * frame's design size has it, from the frame's corner cell.
 */
export function smartArtMemberAnchorForBounds(
  anchor: SpreadsheetObjectAnchor,
  frame: SpreadsheetSmartArtFrame,
  bounds: SpreadsheetObjectBounds,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectAnchor {
  const live = smartArtFrameBounds(frame, columnWidth, rowHeight);
  const scaleX = live.width > 0 ? live.width / (frame.designWidthEmu / EMU_PER_CSS_PIXEL) : 1;
  const scaleY = live.height > 0 ? live.height / (frame.designHeightEmu / EMU_PER_CSS_PIXEL) : 1;
  const relLeft = (bounds.left - live.left) / scaleX;
  const relTop = (bounds.top - live.top) / scaleY;
  const sized = {
    widthEmu: cssPixelsToEmu(Math.max(1, bounds.width / scaleX)),
    heightEmu: cssPixelsToEmu(Math.max(1, bounds.height / scaleY)),
  };
  const corner = smartArtFrameAnchor(frame);
  if (corner.kind === "absolute") {
    return {
      ...anchor,
      kind: "absolute",
      ...(corner.origin === "sheet" ? { origin: "sheet" } : {}),
      xEmu: Math.max(0, Math.round(Number(corner.xEmu ?? 0) + relLeft * EMU_PER_CSS_PIXEL)),
      yEmu: Math.max(0, Math.round(Number(corner.yEmu ?? 0) + relTop * EMU_PER_CSS_PIXEL)),
      ...sized,
    };
  }
  const from = anchorPointWithDefaults(corner.from);
  const rest = { ...anchor };
  delete rest.to;
  return {
    ...rest,
    kind: "oneCell",
    from: {
      row: from.row,
      column: from.column,
      // The member's place in the graphic's layout, measured from its corner
      // cell; a quarter-turned box starts before it (a negative offset).
      rowOffset: Math.round(Number(from.rowOffset ?? 0) + relTop * EMU_PER_CSS_PIXEL),
      columnOffset: Math.round(Number(from.columnOffset ?? from.colOffset ?? 0) + relLeft * EMU_PER_CSS_PIXEL),
    },
    ...sized,
  };
}

/**
 * A frame placed on `bounds` (drawing-layer px), keeping its placement, laid
 * out at that size: what a graphic moved, resized or (re)built takes.
 */
export function smartArtFrameForBounds(
  frame: SpreadsheetObjectAnchor | undefined,
  bounds: SpreadsheetObjectBounds,
  grid: DrawingGrid,
): SpreadsheetSmartArtFrame {
  const base: SpreadsheetObjectAnchor = frame
    ? smartArtFrameAnchorOnly(frame)
    : { kind: "twoCell", from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
  const placed = anchorForObjectBounds(base, bounds, grid.rowCount, grid.columnCount, grid.colWidth, grid.rowHeight);
  return {
    ...placed,
    designWidthEmu: cssPixelsToEmu(Math.max(1, bounds.width)),
    designHeightEmu: cssPixelsToEmu(Math.max(1, bounds.height)),
  };
}

/**
 * The objects with each SmartArt frame that has no design size yet given the
 * size the grid draws it at now (the same array when none lacked one).
 *
 * A diagram converted from a workbook is laid out where its graphic frame
 * stands, but the frame's two-cell size is only known on the grid it sits on,
 * which the conversion does not have; the window settles it on arrival, so
 * the members are drawn exactly as converted and later stretch with the frame.
 * Until then a member is drawn as anchored (`readSmartArtFrame` is null).
 */
export function smartArtFramesSettled<O extends DrawingObjectLike>(
  objects: readonly O[],
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): readonly O[] {
  let changed = false;
  const settled = objects.map((object) => {
    const raw = object.anchor?.smartArtFrame;
    if (!raw || typeof raw !== "object" || Array.isArray(raw) || readSmartArtFrame(object.anchor)) return object;
    const frame = smartArtFrameAnchorOnly(raw as SpreadsheetObjectAnchor);
    if (frame.kind === "twoCell" ? !frame.from || !frame.to : frame.kind === "oneCell" ? !frame.from : frame.kind !== "absolute") {
      return object;
    }
    const live = objectBoundsOnGrid({ id: "", anchor: frame }, columnWidth, rowHeight);
    if (!(live.width > 0) || !(live.height > 0)) return object;
    changed = true;
    return {
      ...object,
      anchor: {
        ...object.anchor,
        smartArtFrame: { ...frame, designWidthEmu: cssPixelsToEmu(live.width), designHeightEmu: cssPixelsToEmu(live.height) },
      },
    };
  });
  return changed ? settled : objects;
}

function smartArtFrameAnchorOnly(frame: SpreadsheetObjectAnchor): SpreadsheetObjectAnchor {
  const copy = { ...frame } as Record<string, unknown>;
  delete copy.designWidthEmu;
  delete copy.designHeightEmu;
  delete copy.smartArtFrame;
  return copy as SpreadsheetObjectAnchor;
}

/**
 * The members of one graphic laid into `frame`: each member's anchor for
 * its box (`boxes`, drawing-layer px, in member order) inside that frame.
 */
export function smartArtMembersInFrame<O extends DrawingObjectLike>(
  members: readonly O[],
  boxes: readonly SpreadsheetObjectBounds[],
  frame: SpreadsheetSmartArtFrame,
  grid: DrawingGrid,
): O[] {
  return members.map((member, index) => ({
    ...member,
    anchor: {
      ...smartArtMemberAnchorForBounds(
        { ...(member.anchor ?? { kind: "oneCell" }), smartArtFrame: frame },
        frame,
        boxes[index],
        grid.colWidth,
        grid.rowHeight,
      ),
      smartArtFrame: frame,
    },
  }));
}

function objectBoundsOnGrid(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds {
  if (object.anchor?.kind === "absolute") {
    const { width, height } = objectSizePixels(
      object,
      columnWidth,
      rowHeight,
    );
    // A worksheet-origin position sits past the headings, beside the
    // cell-anchored objects (live audit B415).
    const sheetOrigin = object.anchor.origin === "sheet";
    const left =
      (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_WIDTH_PX : 0) +
      Math.max(0, Number(object.anchor.xEmu ?? 0) / EMU_PER_CSS_PIXEL);
    const top =
      (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX : 0) +
      Math.max(0, Number(object.anchor.yEmu ?? 0) / EMU_PER_CSS_PIXEL);
    return { left, top, width, height };
  }
  const from = object.anchor?.from;
  const fromRow = anchorPointRow(from);
  const fromCol = anchorPointColumn(from);
  const { width, height } = objectSizePixels(
    object,
    columnWidth,
    rowHeight,
  );
  // A SmartArt member's offset is its place in the graphic's layout, measured
  // from the graphic's corner cell (the graphic is ONE Excel object), not a
  // marker Excel clamps to its cell.
  const layoutRelative = spreadsheetObjectOffsetsLayoutRelative(object);
  const offset = (axis: "row" | "column", size: number) =>
    layoutRelative ? anchorPointOffsetPixels(from, axis) : anchorPointOffsetInLinePixels(from, axis, size);
  return {
    left:
      DEFAULT_SPREADSHEET_HEADER_WIDTH_PX +
      offsetOfLine(fromCol, fromCol, columnWidth) +
      offset("column", columnWidth(fromCol)),
    top:
      DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX +
      offsetOfLine(fromRow, fromRow, rowHeight) +
      offset("row", rowHeight(fromRow)),
    width,
    height,
  };
}








export function spreadsheetObjectCollapsedByCells(
  object: DrawingObjectLike,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): boolean {
  // A SmartArt member goes with its graphic's frame.
  const frame = readSmartArtFrame(object.anchor);
  if (frame) {
    if (frame.kind !== "twoCell" || !frame.to) return false;
    const live = objectGridSize({ id: "", anchor: smartArtFrameAnchor(frame) }, columnWidth, rowHeight);
    return live.width <= 0 || live.height <= 0;
  }
  if (object.anchor?.kind !== "twoCell" || !object.anchor.to) return false;
  const { width, height } = objectGridSize(object, columnWidth, rowHeight);
  return width <= 0 || height <= 0;
}

export function spreadsheetObjectRotation(anchor: SpreadsheetObjectAnchor | undefined) {
  const raw = Number(anchor?.rotation ?? 0);
  if (!Number.isFinite(raw)) return 0;
  const normalized = ((raw % 360) + 360) % 360;
  return normalized > 180 ? normalized - 360 : normalized;
}

/**
 * A picture or shape this page inserted into an imported workbook: it has no
 * import identity (everything the importer read carries one; where it came
 * from, if recorded, is its `provenance`, and an older page's provenance
 * record in `source` is no identity either: lib/office-object-provenance,
 * one rule with every editor and the API). The in-place writer writes such an object
 * whole on every save (`XlsxInsertedObjectPatch`), so a move, resize, format
 * or delete of it saves like any other edit. An ink stroke is not an
 * inserted object and keeps its own rule.
 */
export function spreadsheetObjectInsertedByPage(
  object: DrawingObjectLike & { shape?: unknown },
): boolean {
  if (object.type !== "image" && object.type !== "shape") return false;
  if (officeObjectHasImportIdentity(object)) return false;
  const shape = object.shape;
  return !(
    object.type === "shape" &&
    shape !== null &&
    typeof shape === "object" &&
    (shape as { kind?: unknown }).kind === "ink"
  );
}

export function spreadsheetObjectHasPatchbackTransformIdentity(
  object: DrawingObjectLike,
): boolean {
  if (spreadsheetObjectInsertedByPage(object)) return true;
  const source = object.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  const record = source as Record<string, unknown>;
  const relationshipIdentity = [
    record.part,
    record.drawingPart,
    record.relationshipId,
  ].every((value) => typeof value === "string" && value.trim().length > 0);
  const nonVisualId = record.nonVisualId;
  const shapeIdentity =
    object.type === "shape" &&
    typeof record.drawingPart === "string" &&
    record.drawingPart.trim().length > 0 &&
    ((typeof nonVisualId === "string" && /^\d+$/.test(nonVisualId.trim())) ||
      (typeof nonVisualId === "number" &&
        Number.isSafeInteger(nonVisualId) &&
        nonVisualId >= 0));
  return relationshipIdentity || shapeIdentity;
}

export function gridAnchorPointAtPixels(
  absolute: number,
  header: number,
  count: number,
  resolveSize: (line: number) => number,
  axis: "row" | "column",
): SpreadsheetAnchorPoint {
  let remaining = Math.max(0, absolute - header);
  let index = 0;
  while (index < Math.max(0, count - 1)) {
    const size = Math.max(0, resolveSize(index));
    if (remaining < size) break;
    remaining -= size;
    index += 1;
  }
  // A point exactly on the last line's far edge is the next line's start,
  // as Excel writes it (row 8 + 0, never row 7 + its full height): the line
  // the marker names then carries the far marker's insert rule.
  const last = Math.max(0, resolveSize(index));
  if (count > 0 && index === count - 1 && last > 0 && Math.abs(remaining - last) < 1e-6) {
    return axis === "row"
      ? { row: index + 1, rowOffset: 0 }
      : { row: 0, column: index + 1, columnOffset: 0 };
  }
  // A marker offset is allowed to sit exactly on a grid line. The general
  // size helper deliberately clamps extents to one EMU, so do not reuse that
  // clamp here or every resize would drift a zero offset to one EMU.
  const offset = Math.max(0, Math.round(remaining * EMU_PER_CSS_PIXEL));
  return axis === "row"
    ? { row: index, rowOffset: offset }
    : { row: 0, column: index, columnOffset: offset };
}

/**
 * A SmartArt graphic's frame on the sheet: the union of its members' LAYOUT
 * boxes, in CSS pixels. A member the author resized or moved inside the
 * graphic counts at the box the layout gave it (its `smartArtCustom` factors
 * taken back off), not at the box it shows: the graphic's frame does not
 * follow one shape's resize, and a rebuild re-lays the diagram out in the
 * frame it had and then re-applies the shape's factors. Measured at the shown
 * boxes, a shape shrunk at the graphic's edge narrowed the frame, the rebuild
 * laid every member out narrower and shrank that shape again, and each text
 * edit took the whole graphic in further (the deck's PPT-B082, on the sheet).
 */
export function sheetSmartArtLayoutFrame(
  members: readonly DrawingObjectLike[],
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectBounds | null {
  // A graphic with its own frame (the anchored graphic frame) is laid out in
  // that frame, where the grid now draws it.
  const frame = members.length ? readSmartArtFrame(members[0].anchor) : null;
  if (frame) return smartArtFrameBounds(frame, columnWidth, rowHeight);
  const boxes = members.map((member) =>
    smartArtLayoutBox(
      objectBoundsPixels(member, columnWidth, rowHeight),
      readSmartArtCustom((member as { smartArtCustom?: unknown }).smartArtCustom),
    ),
  );
  if (boxes.length === 0) return null;
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * The farthest up and left a drawing's box can be carried or stretched, in
 * the drawing layer's px: cell A1's corner for one anchored to the cells or
 * placed from the worksheet's origin (Excel stops a drag at the sheet's
 * edge), the layer's own corner for an absolute one placed from it. What
 * `anchorForObjectBounds` writes for a box past it is this corner, so a
 * preview let past it -- over the row or column headings -- jumped back on
 * release (and a two-cell drawing lost the overshoot from its size). A
 * SmartArt member is placed in its graphic's layout and keeps the layer's.
 */
export function spreadsheetObjectBoundsFloor(anchor: SpreadsheetObjectAnchor | undefined): { left: number; top: number } {
  if (anchor && (readSmartArtFrame(anchor) || (anchor.kind === "absolute" && anchor.origin !== "sheet"))) {
    return { left: 0, top: 0 };
  }
  return { left: DEFAULT_SPREADSHEET_HEADER_WIDTH_PX, top: DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX };
}

export function anchorForObjectBounds(
  anchor: SpreadsheetObjectAnchor,
  bounds: SpreadsheetObjectBounds,
  rowCount: number,
  columnCount: number,
  columnWidth: (column: number) => number,
  rowHeight: (row: number) => number,
): SpreadsheetObjectAnchor {
  // A SmartArt member moved or sized inside its graphic: its place in the
  // graphic's layout, the frame where it is.
  const frame = readSmartArtFrame(anchor);
  if (frame) return smartArtMemberAnchorForBounds(anchor, frame, bounds, columnWidth, rowHeight);
  if (anchor.kind === "absolute") {
    const sheetOrigin = anchor.origin === "sheet";
    return {
      ...anchor,
      xEmu: cssPixelsToEmu(Math.max(0, bounds.left - (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_WIDTH_PX : 0))),
      yEmu: cssPixelsToEmu(Math.max(0, bounds.top - (sheetOrigin ? DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX : 0))),
      widthEmu: cssPixelsToEmu(bounds.width),
      heightEmu: cssPixelsToEmu(bounds.height),
    };
  }
  // The markers name the line the edge is on in the WHOLE grid, as Excel's
  // do, not the stored extent: an object below the sheet's used rows sits on
  // its own row, which the extent then grows to hold (sheetCoveringItsCells).
  // Capped at the extent, its markers named the last stored line with an
  // offset past it, which Excel (and the grid) clamp to that line's end.
  rowCount = Math.max(rowCount, MAX_GRID_ROWS);
  columnCount = Math.max(columnCount, MAX_GRID_COLUMNS);
  const fromRow = gridAnchorPointAtPixels(
    bounds.top,
    DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
    rowCount,
    rowHeight,
    "row",
  );
  const fromColumn = gridAnchorPointAtPixels(
    bounds.left,
    DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
    columnCount,
    columnWidth,
    "column",
  );
  const from: SpreadsheetAnchorPoint = {
    ...anchor.from,
    row: fromRow.row,
    rowOffset: fromRow.rowOffset,
    column: fromColumn.column,
    columnOffset: fromColumn.columnOffset,
  };
  if (anchor.kind === "twoCell") {
    const toRow = gridAnchorPointAtPixels(
      bounds.top + bounds.height,
      DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX,
      rowCount,
      rowHeight,
      "row",
    );
    const toColumn = gridAnchorPointAtPixels(
      bounds.left + bounds.width,
      DEFAULT_SPREADSHEET_HEADER_WIDTH_PX,
      columnCount,
      columnWidth,
      "column",
    );
    return {
      ...anchor,
      from,
      to: {
        ...anchor.to,
        row: toRow.row,
        rowOffset: toRow.rowOffset,
        column: toColumn.column,
        columnOffset: toColumn.columnOffset,
      },
      // An imported two-cell anchor may also state its extent, which the
      // size reads first: it follows the new bounds too.
      ...(anchor.widthEmu !== undefined ? { widthEmu: cssPixelsToEmu(bounds.width) } : {}),
      ...(anchor.heightEmu !== undefined ? { heightEmu: cssPixelsToEmu(bounds.height) } : {}),
    };
  }
  return {
    ...anchor,
    kind: "oneCell",
    from,
    widthEmu: cssPixelsToEmu(bounds.width),
    heightEmu: cssPixelsToEmu(bounds.height),
  };
}

/** A line or connector with no label: Office offers it no text, nor do we. */
export function spreadsheetLineShapeTakesNoText(
  shape: SpreadsheetShapeStyle | undefined,
): boolean {
  if (!shape || shape.text) return false;
  return (
    isSpreadsheetPresetShapeKind(shape.kind) &&
    isSpreadsheetLineLikeShapeKind(shape.kind)
  );
}


// ---------------------------------------------------------------------------
// Shape Format and Picture Format buttons.

/** The fields of a picture or shape object the format buttons read. */
interface DrawingObjectFields extends DrawingObjectLike {
  shape?: unknown;
  shadow?: { preset?: string; color?: string } & Record<string, unknown>;
  effect?: unknown;
  alt?: string;
  caption?: string;
  lockAspectRatio?: boolean;
  /** Format > Properties > Print object off (`fPrintsWithSheet="0"`). */
  printObject?: false;
  cropRect?: unknown;
  pictureStyle?: string;
  nativePictureStyleBlocker?: string;
}

/** A patch the caller merges into the object (`{ ...object, ...patch }`). */
export type DrawingObjectPatch = Record<string, unknown>;

/** The grid's line sizes, which place and size a cell-anchored object. */
export interface DrawingGrid {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly colWidth: (column: number) => number;
  readonly rowHeight: (row: number) => number;
}

function fieldsOf(object: object): DrawingObjectFields {
  return object as DrawingObjectFields;
}

/** Alt Text, Name: control characters out, trimmed, bounded. */
export function normalizedTableAltTextInput(value: string, maxLength: number): string {
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .trim()
    .slice(0, maxLength);
}

export const MAX_PICTURE_ALT_TEXT_LENGTH = 2_000;
export const MAX_PICTURE_NAME_LENGTH = 255;

/** Shape Format > Shape Styles. */
export function shapeQuickStylePatch(
  object: object,
  parameter: string,
  recipes: readonly OfficeShapeStyleRecipe[],
): DrawingObjectPatch | null {
  const styleId = parameter;
  const recipe = recipes.find(
    (candidate) => candidate.id === styleId,
  );
  if (!recipe) return null;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  // A style's fill replaces a picture fill, as Office's Shape Styles do.
  const next = normalizeShapeStyle({
    ...spreadsheetShapeStyleWithoutPictureFill(object, current),
    fill: recipe.fill,
    fillGradient:recipe.fillGradient,
    stroke: recipe.stroke,
    strokeWidth: recipe.strokeWidth,
    textColor: recipe.textColor,
  });
  const shadow = recipe.shapeShadow
    ? { preset: "outer-bottom" as const, ...recipe.shapeShadow }
    : undefined;
  if (
    JSON.stringify(next) === JSON.stringify(current) &&
    JSON.stringify(shadow ?? null) === JSON.stringify(fieldsOf(object).shadow ?? null)
  ) return null;
  return { shape: next, shadow };
}

/**
 * Picture Format on a SmartArt picture slot's picture: the slot's next crop
 * and Picture Style (`{}` is Reset Picture), written on the object as
 * `smartArtPictureCrop` / `smartArtPictureStyle`
 * (`lib/office-smartart/slot-picture-format`). The slot's picture is its
 * fill, so the command rides the Format Shape fill writer. Only a slot that
 * holds a picture takes it.
 */
export function smartArtSlotPicturePatch(object: object, value: unknown): DrawingObjectPatch | null {
  const record = object as { smartArtPicturePlaceholder?: unknown; smartArtPictureCrop?: unknown; smartArtPictureStyle?: unknown };
  const shape = normalizeShapeStyle(fieldsOf(object).shape);
  if (record.smartArtPicturePlaceholder !== true || typeof shape.src !== "string" || !shape.src) return null;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const next = smartArtSlotPictureFormatOf({
    smartArtPictureCrop: (value as { crop?: unknown }).crop,
    smartArtPictureStyle: (value as { style?: unknown }).style,
  });
  const current = smartArtSlotPictureFormatOf(record);
  if (JSON.stringify(next) === JSON.stringify(current)) return null;
  return {
    smartArtPictureCrop: next.crop,
    smartArtPictureStyle: next.style,
  };
}

/** Shape Format > Edit Shape > Change Shape. */
export function shapeChangePatch(object: object, parameter: string): DrawingObjectPatch | null {
  const kind = parameter as SpreadsheetPresetShapeKind;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if (
    kind === "text-box" ||
    !isSpreadsheetPresetShapeKind(kind) ||
    kind === current.kind
  ) {
    return null;
  }
  // A new preset has its own handles: the ones dragged on the old
  // preset are named for it and go back to the defaults, as Word does.
  const shape = normalizeShapeStyle({ ...current, kind, adjustments: undefined });
  // One shape of a SmartArt graphic: the shape is the AUTHOR's from now on
  // (Office keeps it as `a:prstGeom` on the item's presentation point), so a
  // relayout draws the item in it again and Reset Shape / Reset Graphic put
  // the layout's back (Microsoft Support, "Restore the default layout and
  // color of your SmartArt graphic"). A picture slot keeps the layout's frame.
  const fields = object as { smartArtRole?: unknown; smartArtKind?: unknown; smartArtPicturePlaceholder?: unknown };
  if (fields.smartArtRole === "node" && typeof fields.smartArtKind === "string") {
    if (fields.smartArtPicturePlaceholder === true) return null;
    return { shape, smartArtAuthorShape: kind };
  }
  return { shape };
}

/**
 * Shape Format > Outline weight, in the PIXELS a sheet object stores, derived
 * from Excel's own point ladder through the ONE shared table -- so the sheet,
 * the deck and the document offer the same nine weights instead of three
 * hand-written sets in two units (audit 64). Zero stays at the front: it is
 * "no outline", which is not a weight Office's submenu lists.
 */
export const SPREADSHEET_SHAPE_OUTLINE_WIDTHS: readonly number[] = Object.freeze([
  0,
  ...officeLineWeightPxLadder(MAX_SHAPE_STROKE_WIDTH),
]);

/** Shape Outline > Weight, in the pixels a sheet object stores (0 = none). */
export function shapeOutlineWeightPatch(
  object: object,
  parameter: string,
): DrawingObjectPatch | null {
  const strokeWidth = Number(parameter);
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  const currentWidth = current.stroke === "none" ? 0 : current.strokeWidth;
  if (
    !SPREADSHEET_SHAPE_OUTLINE_WIDTHS.includes(
      strokeWidth,
    ) ||
    strokeWidth === currentWidth
  ) {
    return null;
  }
  return {
    shape: normalizeShapeStyle({
      ...current,
      stroke:
        strokeWidth === 0
          ? "none"
          : current.stroke === "none"
            ? DEFAULT_SHAPE_STROKE
            : current.stroke,
      strokeWidth,
    }),
  };
}

/** Shape Fill: a colour or "none". */
export function shapeFillColorPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const color =
    parameter === "none"
      ? "none"
      : normalizeHexColor(parameter);
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  const unpictured = spreadsheetShapeStyleWithoutPictureFill(object, current);
  if (!color || (color === current.fill && !current.fillGradient && unpictured === current)) return null;
  // A colour replaces a picture fill, as Shape Fill does in Office.
  return {
    shape: normalizeShapeStyle({ ...unpictured, fill: color, fillGradient: undefined }),
  };
}

/** Shape Outline: a colour or "none". */
export function shapeOutlineColorPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const color =
    parameter === "none"
      ? "none"
      : normalizeHexColor(parameter);
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if (!color) return null;
  // A colour turns the line on: a shape left with no weight (an Intense
  // Effect style) shows it at Office's 0.75 pt, as Excel does.
  const strokeWidth =
    color === "none"
      ? current.strokeWidth
      : officeOutlineWeightWhenShown(
          current.strokeWidth,
          officeLineWeightPtToPx(OFFICE_OUTLINE_SHOWN_WEIGHT_PT),
        );
  if (color === current.stroke && strokeWidth === current.strokeWidth) return null;
  return {
    shape: normalizeShapeStyle({ ...current, stroke: color, strokeWidth }),
  };
}

/**
 * Shape Outline > Dashes: one of Office's eight (`OFFICE_MENU_DASHES`). The
 * exact dash on the style (none for Solid), Round Dot's round cap (and a cap
 * the earlier Round Dot set goes with it); a shape with no line shows one,
 * as a colour pick does.
 */
export function shapeOutlineDashPatch(object: object, parameter: string): DrawingObjectPatch | null {
  if (!(OFFICE_MENU_DASHES as readonly string[]).includes(parameter)) return null;
  const dash = parameter as OfficeMenuDash;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  const outlineDash = dash === "solid" ? undefined : dash;
  const outlineCap =
    officeMenuDashFields(dash).cap ??
    (current.outlineCap === "round" && current.outlineDash === "sysDot" ? undefined : current.outlineCap);
  const stroke = current.stroke === "none" ? DEFAULT_SHAPE_STROKE : current.stroke;
  const strokeWidth = officeOutlineWeightWhenShown(
    current.strokeWidth,
    officeLineWeightPtToPx(OFFICE_OUTLINE_SHOWN_WEIGHT_PT),
  );
  if (
    current.outlineDash === outlineDash &&
    current.outlineCap === outlineCap &&
    stroke === current.stroke &&
    strokeWidth === current.strokeWidth
  ) {
    return null;
  }
  return {
    shape: normalizeShapeStyle({ ...current, stroke, strokeWidth, outlineDash, outlineCap }),
  };
}

/** Shape Outline > Sketched: "none", "curved", "freehand" or "scribble";
 * a shape with no line shows one, as a dash pick does. */
export function shapeOutlineSketchPatch(object: object, parameter: string): DrawingObjectPatch | null {
  if (!["none", "curved", "freehand", "scribble"].includes(parameter)) return null;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  const outlineSketch = parameter === "none" ? undefined : (parameter as "curved" | "freehand" | "scribble");
  const showsLine = outlineSketch !== undefined;
  const stroke = showsLine && current.stroke === "none" ? DEFAULT_SHAPE_STROKE : current.stroke;
  const strokeWidth = showsLine
    ? officeOutlineWeightWhenShown(current.strokeWidth, officeLineWeightPtToPx(OFFICE_OUTLINE_SHOWN_WEIGHT_PT))
    : current.strokeWidth;
  if (current.outlineSketch === outlineSketch && stroke === current.stroke && strokeWidth === current.strokeWidth) {
    return null;
  }
  return { shape: normalizeShapeStyle({ ...current, stroke, strokeWidth, outlineSketch }) };
}

/** Text Effects > Transform: a preset warp token, or "none". A shape with
 * no words takes none (its style keeps a warp only beside a label). */
export function shapeTextTransformPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if (!current.text) return null;
  const warp = parameter === "none" ? null : normalizeOfficeTextWarp({ preset: parameter });
  if (parameter !== "none" && !warp) return null;
  if ((current.textWarp?.preset ?? null) === (warp?.preset ?? null)) return null;
  return { shape: normalizeShapeStyle({ ...current, textWarp: warp ?? undefined }) };
}

/**
 * Format Shape > Line: Compound type, Cap type and Join type, as JSON
 * (`{outlineCompound?, outlineCap?, outlineJoin?}`; "single" clears the
 * compound). Each held to its vocabulary; a join other than miter drops the
 * miter limit.
 */
export function shapeLineDetailPatch(object: object, parameter: string): DrawingObjectPatch | null {
  let picked: Record<string, unknown>;
  try {
    picked = JSON.parse(parameter) as Record<string, unknown>;
  } catch {
    return null;
  }
  if (!picked || typeof picked !== "object") return null;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  const next: Record<string, unknown> = { ...current };
  if ("outlineCompound" in picked) {
    next.outlineCompound = picked.outlineCompound === "single" ? undefined : picked.outlineCompound;
  }
  if ("outlineCap" in picked) next.outlineCap = picked.outlineCap;
  if ("outlineJoin" in picked) {
    next.outlineJoin = picked.outlineJoin;
    if (picked.outlineJoin !== "miter") next.outlineMiterLimit = undefined;
  }
  const shape = normalizeShapeStyle(next);
  return JSON.stringify(shape) === JSON.stringify(current) ? null : { shape };
}

/**
 * Format Shape > Fill & Line (the shared page): the shape's fill and line as
 * the page wrote them, as JSON -- `fill`, `fillGradient`, `fillPattern`,
 * `fillTransparency`, `stroke`, `outlineGradient`, `outlineTransparency`,
 * `strokeWidth`; a key given as null takes that part off (written out, so
 * the file's goes too). Re-read through the shape's own normalizer.
 */
export function shapeFillLinePatch(object: object, parameter: string): DrawingObjectPatch | null {
  let picked: Record<string, unknown>;
  try {
    picked = JSON.parse(parameter) as Record<string, unknown>;
  } catch {
    return null;
  }
  if (!picked || typeof picked !== "object" || Array.isArray(picked)) return null;
  if ("smartArtPicture" in picked) {
    return Object.keys(picked).length === 1 ? smartArtSlotPicturePatch(object, picked.smartArtPicture) : null;
  }
  const keys = ["fill", "fillGradient", "fillPattern", "fillTransparency", "stroke", "outlineGradient", "outlineTransparency", "strokeWidth"];
  if (Object.keys(picked).some((key) => !keys.includes(key) && key !== "fillPicture" && key !== "fillPictureOptions")) return null;
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if ("fillPictureOptions" in picked && !("fillPicture" in picked)) {
    // Picture or texture fill's options (tile, offsets, Transparency, Rotate
    // with shape): only on a shape whose fill IS its picture, and alone.
    const options = normalizeOfficePictureFillOptions(picked.fillPictureOptions);
    if (!options || Object.keys(picked).length !== 1 || !spreadsheetShapeHasPictureFill({ ...object, shape: current })) return null;
    const next: Record<string, unknown> = { ...current };
    if (officePictureFillOptionsStated(options)) next.fillPictureOptions = options;
    else delete next.fillPictureOptions;
    const shape = normalizeShapeStyle(next);
    return JSON.stringify(shape) === JSON.stringify(current) ? null : { shape };
  }
  if ("fillPicture" in picked) {
    // Picture or texture fill: the picture IS the fill, stretched over the
    // box; the colour, gradient and pattern underneath are not drawn and are
    // replaced when another kind is picked. Nothing else rides with it.
    // A texture brings its tile with it (one Undo step), as Office's does.
    const src = officePictureFillSrc(picked.fillPicture);
    const withOptions = "fillPictureOptions" in picked;
    const options = withOptions ? normalizeOfficePictureFillOptions(picked.fillPictureOptions) : null;
    if (!src || Object.keys(picked).length !== (withOptions ? 2 : 1) || (withOptions && !options)) return null;
    if (current.kind === "ink" || (isSpreadsheetPresetShapeKind(current.kind) && isSpreadsheetLineLikeShapeKind(current.kind))) return null;
    const pictured: Record<string, unknown> = {
      ...current,
      fillGradient: undefined,
      fillPattern: undefined,
      fillTransparency: undefined,
      src,
      fillPicture: true,
    };
    if (options && officePictureFillOptionsStated(options)) pictured.fillPictureOptions = options;
    else if (withOptions) delete pictured.fillPictureOptions;
    const shape = normalizeShapeStyle(pictured);
    return JSON.stringify(shape) === JSON.stringify(current) ? null : { shape };
  }
  // Any other fill the page writes replaces the picture.
  const fillWritten = ["fill", "fillGradient", "fillPattern", "fillTransparency"].some((key) => key in picked);
  const next: Record<string, unknown> = {
    ...(fillWritten ? spreadsheetShapeStyleWithoutPictureFill(object, current) : current),
  };
  for (const key of keys) {
    if (!(key in picked)) continue;
    // A gradient fill the page takes off is no key (the model keeps none as null).
    if (key === "fillGradient" && picked[key] === null) delete next.fillGradient;
    else next[key] = picked[key];
  }
  const shape = normalizeShapeStyle(next);
  return JSON.stringify(shape) === JSON.stringify(current) ? null : { shape };
}

/** Shape Outline > Arrows: the picked line-end state as JSON. */
export function shapeLineEndsPatch(object: object, parameter: string, grid: DrawingGrid): DrawingObjectPatch | null {
  // Shape Outline > Arrows. The receipt carries the picked state; it
  // is re-read through the shared vocabulary, never trusted as typed.
  let picked: unknown;
  try {
    picked = JSON.parse(parameter);
  } catch {
    return null;
  }
  if (!picked || typeof picked !== "object") return null;
  const bounds = objectBoundsPixels(fieldsOf(object), grid.colWidth, grid.rowHeight);
  const next = spreadsheetShapeLineEndsPatch(
    fieldsOf(object).shape,
    readOfficeLineEnds(picked as Record<string, unknown>),
    bounds.width,
    bounds.height,
  );
  return next ? { shape: next } : null;
}

/** Shape Effects > Shadow: a preset or "none". */
export function objectShadowPresetPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const preset = parameter;
  if (
    preset !== "none" &&
    !(SPREADSHEET_SHADOW_PRESETS as readonly string[]).includes(preset)
  ) {
    return null;
  }
  // A shape that states Office's effects keeps its shadow there: the older
  // preset is read as the Office shadow it draws.
  const officeEffects = normalizeOfficePictureEffects((object as { shapeEffects?: unknown }).shapeEffects);
  if (officeEffects) {
    const shadow =
      preset === "none" ? null : (spreadsheetShapeOfficeEffects({ shadow: { preset } })?.shadow ?? null);
    const next = withOfficePictureEffect(officeEffects, "shadow", shadow as OfficeShadowEffect | null, false);
    return JSON.stringify(next ?? null) === JSON.stringify(officeEffects) ? null : { shapeEffects: next };
  }
  const currentPreset = fieldsOf(object).shadow?.preset ?? "none";
  if (preset === currentPreset) return null;
  return {
    shadow: normalizeShadow({
      ...(fieldsOf(object).shadow ?? {}),
      preset,
    }),
  };
}

/** Shape Effects > Shadow > Color, on an object that has a shadow. */
export function objectShadowColorPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const officeEffects = normalizeOfficePictureEffects((object as { shapeEffects?: unknown }).shapeEffects);
  if (officeEffects) {
    const color = normalizeHexColor(parameter);
    const current = officeEffects.shadow && officeEffects.shadow !== "none" ? officeEffects.shadow : null;
    if (!color || !current || color === current.color) return null;
    return { shapeEffects: withOfficePictureEffect(officeEffects, "shadow", { ...current, color }, false) };
  }
  const shadow = fieldsOf(object).shadow;
  const color = normalizeHexColor(parameter);
  if (
    !color ||
    !shadow ||
    color === (shadow.color ?? DEFAULT_SHADOW_COLOR)
  ) {
    return null;
  }
  return {
    shadow: normalizeShadow({ ...shadow, color }),
  };
}

/**
 * The text a shape shows ("" removes it). `textRuns` are the typed words'
 * own runs (null: none); absent, the shape's runs follow the new words.
 */
export function shapeTextPatch(
  object: object,
  parameter: string,
  textRuns?: readonly TextRun[] | null,
): DrawingObjectPatch | null {
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  // A line takes no text, in Excel either; one that already carries a
  // label may still change or lose it.
  if (spreadsheetLineShapeTakesNoText(current)) return null;
  const text = parameter.slice(0, MAX_SHAPE_TEXT_LENGTH);
  const runs =
    textRuns === undefined
      ? spreadsheetShapeTextRunsFollowing(current.textRuns, current.text ?? "", text)
      : textRuns
        ? spreadsheetShapeTextRunsClipped(textRuns, MAX_SHAPE_TEXT_LENGTH)
        : undefined;
  const rest: SpreadsheetShapeStyle = { ...current };
  delete rest.textRuns;
  // The file's paragraph layout follows the paragraphs typed (a split
  // paragraph keeps its bullet in both halves), as Office carries it.
  const textLayout = officeTextLayoutFollowingText(current.textLayout, current.text ?? "", text);
  if (textLayout) rest.textLayout = textLayout;
  else delete rest.textLayout;
  const typed = normalizeShapeStyle({ ...rest, text, ...(runs ? { textRuns: runs } : {}) });
  if (typed.text === current.text && JSON.stringify(typed.textRuns) === JSON.stringify(current.textRuns)) return null;
  // Words typed beside an equation box's equation leave it the box Insert >
  // Equation made: Excel keeps its text body (no wrap, no insets, Resize
  // shape to fit text), so the words go on the equation's line and the box
  // grows to hold them. A box saved before it carried that body takes it now;
  // a file's box keeps the body its file states (the write-back keeps it).
  const next =
    !officeObjectHasImportIdentity(object) &&
    spreadsheetEquationBoxLatex(current) !== null &&
    spreadsheetEquationBoxLatex(typed) === null
      ? normalizeShapeStyle(withSpreadsheetEquationBoxBody(typed))
      : typed;
  return { shape: next };
}

/** Font size of a shape's text. */
export function shapeTextSizePatch(object: object, parameter: string): DrawingObjectPatch | null {
  const size = Number(parameter);
  if (
    !Number.isFinite(size) ||
    size < MIN_SHAPE_TEXT_SIZE ||
    size > MAX_SHAPE_TEXT_SIZE
  ) {
    return null;
  }
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  // The selected shape's size is every run's: a run's own size goes.
  const textRuns = spreadsheetShapeTextRunsWithout(current.textRuns, current.text ?? "", ["fontSize"]);
  const rest: SpreadsheetShapeStyle = { ...current };
  delete rest.textRuns;
  const next = normalizeShapeStyle({ ...rest, textSize: size, ...(textRuns ? { textRuns } : {}) });
  if (next.textSize === current.textSize && JSON.stringify(next.textRuns) === JSON.stringify(current.textRuns)) return null;
  return { shape: next };
}

/**
 * Font colour of a shape's text. A WordArt paint draws the text over its
 * plain colour, so with one the colour is the paint's fill as well: setting
 * the plain colour alone changed a field nothing drew (B524: Format Shape's
 * Text Fill turned green while the WordArt text stayed blue). The ribbon's
 * Text Fill already wrote the paint; every caller of the command does now.
 */
export function shapeTextColorPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const color = normalizeHexColor(parameter);
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if (!color) return null;
  const paint = current.textPaint ? patchOfficeTextPaint(current.textPaint, { fill: color }) : undefined;
  const paintChanged = paint !== undefined && !sameOfficeTextPaint(paint, current.textPaint);
  // Every run takes the colour: its own colour goes, its own WordArt takes
  // it as the fill.
  const textRuns = spreadsheetShapeTextRunsFilled(current.textRuns, current.text ?? "", color);
  const runsChanged = JSON.stringify(textRuns) !== JSON.stringify(current.textRuns);
  if (color === current.textColor && !paintChanged && !runsChanged) return null;
  const rest: SpreadsheetShapeStyle = { ...current };
  delete rest.textRuns;
  return {
    shape: normalizeShapeStyle({
      ...rest,
      textColor: color,
      ...(paint ? { textPaint: paint } : {}),
      ...(textRuns ? { textRuns } : {}),
    }),
  };
}

/** WordArt Styles on a shape's text: the whole paint as JSON, or "none". */
export function shapeTextPaintPatch(
  object: object,
  parameter: string,
  // The command as an edit of each paint (Text Fill, Text Outline, Text
  // Effects): a run wearing WordArt of its own takes the same edit, so it
  // keeps what the edit leaves alone. Absent, the parameter is every
  // character's paint (a WordArt style, or none).
  edit?: (paint: OfficeTextPaint) => OfficeTextPaint,
): DrawingObjectPatch | null {
  // The parameter is the whole paint (JSON) or "none" to clear it; it
  // is re-read through the shared normaliser, never trusted as typed.
  const current = normalizeShapeStyle(fieldsOf(object).shape);
  if (spreadsheetLineShapeTakesNoText(current) || !current.text) return null;
  let paint: OfficeTextPaint | undefined;
  if (parameter !== "none") {
    try {
      paint = normalizeOfficeTextPaint(JSON.parse(parameter));
    } catch {
      return null;
    }
    if (!paint) return null;
  }
  const { textPaint: _previous, textArt: _art, ...rest } = current;
  delete rest.textRuns;
  // The selected shape's WordArt is every character's: a run's own paint is
  // edited the same way (or goes, for a whole style), and a run's own
  // colour goes once the fill changes.
  const fillChanged = (paint?.fill ?? null) !== (current.textPaint?.fill ?? null);
  const text = current.text ?? "";
  const textRuns = current.textRuns
    ? normalizeSpreadsheetShapeTextRuns(
        current.textRuns.map((run) => {
          // An equation among the words stays one (its characters are restyled).
          if (!run.style) return run.math ? { text: run.text, math: run.math } : { text: run.text };
          const style = { ...run.style } as Record<string, unknown>;
          if (fillChanged) delete style.color;
          if (run.style.textPaint && edit && paint) {
            try {
              style.textPaint = edit(run.style.textPaint);
            } catch {
              delete style.textPaint;
            }
            // Edited into the label's own: the run wears the label's.
            if (style.textPaint && sameOfficeTextPaint(style.textPaint as OfficeTextPaint, paint)) delete style.textPaint;
          } else delete style.textPaint;
          return run.math ? { text: run.text, style, math: run.math } : { text: run.text, style };
        }),
        text,
      )
    : undefined;
  const next = normalizeShapeStyle({ ...rest, ...(paint ? { textPaint: paint } : {}), ...(textRuns ? { textRuns } : {}) });
  if (
    sameOfficeTextPaint(next.textPaint ?? null, current.textPaint ?? null) &&
    !current.textArt &&
    JSON.stringify(next.textRuns) === JSON.stringify(current.textRuns)
  )
    return null;
  return { shape: next };
}

/** Shape Effects: one preset at a time. Shadow is the sheet's own shadow
 * model (its first preset); the five others are `effect`, stored as "none"
 * when cleared so the in-place writer can tell a removal from a workbook
 * that predates the field. */
export function shapeEffectPresetPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const choice = parameter;
  // Office's Shape Effects menu: None, Shadow and the five presets.
  if (!(["none", "shadow", ...OFFICE_SHAPE_EFFECT_PRESETS] as readonly string[]).includes(choice)) return null;
  const preset = normalizeOfficeShapeEffectPreset(choice);
  // A shape that states Office's effects: the one-of-seven pick becomes them.
  if (normalizeOfficePictureEffects((object as { shapeEffects?: unknown }).shapeEffects)) {
    return {
      shapeEffects:
        choice === "none"
          ? undefined
          : choice === "shadow"
            ? spreadsheetShapeOfficeEffects({ shadow: { preset: SPREADSHEET_SHADOW_PRESETS[0] } })
            : spreadsheetShapeOfficeEffects({ effect: preset }),
    };
  }
  const currentPreset = normalizeOfficeShapeEffectPreset(fieldsOf(object).effect);
  const currentShadow = fieldsOf(object).shadow ? "shadow" : null;
  if ((currentPreset ?? currentShadow ?? "none") === choice) return null;
  return {
    shadow:
      choice === "shadow"
        ? normalizeShadow({
            ...(fieldsOf(object).shadow ?? {}),
            preset: SPREADSHEET_SHADOW_PRESETS[0],
          })
        : undefined,
    effect: preset ?? (currentPreset || "effect" in object ? "none" : undefined),
  };
}

/**
 * Whether a drawing keeps its proportions (Size > Lock aspect ratio): a
 * picture does unless it was unlocked -- Excel inserts every picture with the
 * box checked, so the corner handles and the Size boxes scale it as a whole --
 * and a shape, chart or diagram only when it was locked (Excel's default for
 * them is unchecked).
 */
export function objectAspectLocked(object: object): boolean {
  const record = fieldsOf(object);
  return record.type === "image"
    ? record.lockAspectRatio !== false
    : record.lockAspectRatio === true;
}

/** Size > Lock aspect ratio (a toggle). A picture, locked by default, stores
 * only its unlocking (`false`); any other drawing stores its locking. */
/**
 * Format > Size & Properties > Properties > Print object: whether the object
 * prints with the sheet (`xdr:clientData@fPrintsWithSheet`, default on).
 * Only "off" is stored.
 */
export function objectPrints(object: object): boolean {
  return (object as { printObject?: unknown }).printObject !== false;
}

export function objectPrintObjectPatch(object: object): DrawingObjectPatch {
  return { printObject: objectPrints(object) ? false : undefined };
}

export function objectLockAspectPatch(object: object): DrawingObjectPatch {
  const locked = objectAspectLocked(object);
  if (fieldsOf(object).type === "image") return { lockAspectRatio: locked ? false : undefined };
  return { lockAspectRatio: locked ? undefined : true };
}

/**
 * Format Shape / Picture > Size & Properties > Properties, Excel's three
 * "Object positioning" choices, as the anchor kind holds them (ECMA-376
 * `xdr:twoCellAnchor@editAs`): two-cell = Move and size with cells, one-cell
 * = Move but don't size with cells, absolute = Don't move or size with cells.
 */
export type SpreadsheetObjectPlacement = "moveAndSize" | "move" | "free";

export const SPREADSHEET_OBJECT_PLACEMENTS: readonly SpreadsheetObjectPlacement[] = [
  "moveAndSize",
  "move",
  "free",
];

export function spreadsheetObjectPlacement(
  anchor: SpreadsheetObjectAnchor | undefined,
): SpreadsheetObjectPlacement {
  return anchor?.kind === "twoCell" ? "moveAndSize" : anchor?.kind === "absolute" ? "free" : "move";
}

/**
 * The object at `placement`, where it stands now: Excel changes only how it
 * follows the cells, never its place or size, when the choice changes. The
 * box it is drawn in becomes the new anchor (two markers, a top-left marker
 * and an extent, or a position from the worksheet's corner). Null when the
 * object already has that placement.
 */
export function objectPlacementPatch(
  object: object,
  placement: string,
  grid: DrawingGrid,
): DrawingObjectPatch | null {
  if (!(SPREADSHEET_OBJECT_PLACEMENTS as readonly string[]).includes(placement)) return null;
  const record = fieldsOf(object);
  if (spreadsheetObjectPlacement(record.anchor) === placement) return null;
  const bounds = objectBoundsPixels(record, grid.colWidth, grid.rowHeight);
  const current = record.anchor ?? {};
  const transform = {
    ...(current.rotation !== undefined ? { rotation: current.rotation } : {}),
    ...(current.flipH ? { flipH: true } : {}),
    ...(current.flipV ? { flipV: true } : {}),
  };
  const kind = placement === "moveAndSize" ? "twoCell" : placement === "move" ? "oneCell" : "absolute";
  // Excel writes every placement chosen here as a two-cell anchor with
  // `editAs` (the export reads `editAs`; Move and size needs none).
  const base: SpreadsheetObjectAnchor =
    kind === "absolute"
      ? { kind, origin: "sheet", editAs: "absolute", ...transform }
      : { kind, from: anchorPointWithDefaults(current.from), ...(kind === "twoCell" ? { to: anchorPointWithDefaults(current.from) } : { editAs: "oneCell" }), ...transform };
  return {
    anchor: anchorForObjectBounds(base, bounds, grid.rowCount, grid.columnCount, grid.colWidth, grid.rowHeight),
  };
}

/**
 * An object's anchor at a new size, as the Size boxes resize it: its top-left
 * corner and its kind kept. An absolute object stays at its x/y; a one-cell
 * object on its `from` cell; a two-cell object (Move and size with cells)
 * stays two-cell, its `to` marker moved to the new bottom-right corner.
 */
function anchorAtSize(
  record: DrawingObjectFields,
  size: { width: number; height: number },
  grid: DrawingGrid,
): SpreadsheetObjectAnchor {
  const anchor = record.anchor?.kind === "absolute" ? record.anchor : anchorWithDefaults(record.anchor);
  if (anchor.kind === "twoCell" && anchor.to) {
    const bounds = objectBoundsPixels(record, grid.colWidth, grid.rowHeight);
    const resized = anchorForObjectBounds(
      anchor,
      { ...bounds, ...size },
      grid.rowCount,
      grid.columnCount,
      grid.colWidth,
      grid.rowHeight,
    );
    return { ...resized, from: anchor.from };
  }
  return anchorWithExplicitSize(anchor, {
    widthEmu: cssPixelsToEmu(size.width),
    heightEmu: cssPixelsToEmu(size.height),
  });
}

/** A shape's Size boxes: width or height in pixels; Lock aspect ratio keeps
 * the proportion through them as it does through the handles. */
export function shapeSizePatch(
  object: object,
  axis: "width" | "height",
  parameter: string,
  grid: DrawingGrid,
): DrawingObjectPatch | null {
  const record = fieldsOf(object);
  const current = objectSizePixels(record, grid.colWidth, grid.rowHeight);
  if (axis === "width") {
    const width = positiveInteger(parameter, current.width);
    if (width === current.width) return null;
    // Lock W:H keeps the proportion through the size boxes as it does
    // through the handles (live audit B160).
    const locked = record.lockAspectRatio === true && current.width > 0;
    return {
      anchor: anchorAtSize(record, {
        width,
        height: locked ? Math.max(1, Math.round((width * current.height) / current.width)) : current.height,
      }, grid),
    };
  }
  const height = positiveInteger(parameter, current.height);
  if (height === current.height) return null;
  const locked = record.lockAspectRatio === true && current.height > 0;
  return {
    anchor: anchorAtSize(record, {
      width: locked ? Math.max(1, Math.round((height * current.width) / current.height)) : current.width,
      height,
    }, grid),
  };
}

/** A picture's Size boxes: width or height in pixels (1..10,000). With Lock
 * aspect ratio on, the other side follows in proportion, as Excel scales a
 * picture whose box is locked (Format Picture > Size). */
export function pictureSizePatch(
  object: object,
  axis: "width" | "height",
  parameter: string,
  grid: DrawingGrid,
): DrawingObjectPatch | null {
  const record = fieldsOf(object);
  const current = objectSizePixels(record, grid.colWidth, grid.rowHeight);
  const value = Number(parameter);
  if (!Number.isFinite(value) || value < 1 || value > 10_000) {
    return null;
  }
  if (Math.abs(value - current[axis]) < 0.5) return null;
  const other = axis === "width" ? "height" : "width";
  const locked = objectAspectLocked(record) && current[axis] > 0;
  return {
    anchor: anchorAtSize(
      record,
      {
        ...current,
        [axis]: value,
        ...(locked ? { [other]: Math.max(1, Math.round((value * current[other]) / current[axis])) } : {}),
      },
      grid,
    ),
  };
}

/** Rotate > Rotate Left 90° ("-90") / Rotate Right 90° ("90"). */
export function objectRotateStepPatch(object: object, parameter: string): DrawingObjectPatch | null {
  if (parameter !== "-90" && parameter !== "90") return null;
  const current = spreadsheetObjectRotation(fieldsOf(object).anchor);
  const step = parameter === "90" ? 90 : -90;
  const next = ((current + step + 180) % 360 + 360) % 360 - 180;
  return {
    anchor: {
      ...anchorWithDefaults(fieldsOf(object).anchor),
      rotation: next,
    },
  };
}

/** Rotate > Flip Vertical / Flip Horizontal ("toggle"). */
export function objectFlipPatch(object: object, axis: "flipH" | "flipV", parameter: string): DrawingObjectPatch | null {
  if (parameter !== "toggle") return null;
  const anchor = anchorWithDefaults(fieldsOf(object).anchor);
  const updated = { ...anchor };
  if (anchor[axis] === true) delete updated[axis];
  else updated[axis] = true;
  return { anchor: updated };
}

/** Size > Rotation: degrees in [0, 360). */
export function objectRotationPatch(object: object, parameter: string): DrawingObjectPatch | null {
  const degrees = Number(parameter);
  if (!Number.isFinite(degrees) || degrees < 0 || degrees >= 360) {
    return null;
  }
  const current = spreadsheetObjectRotation(fieldsOf(object).anchor);
  const next =
    Math.round(
      ((((degrees + 180) % 360) + 360) % 360 - 180) * 10,
    ) / 10;
  if (next === current) return null;
  return {
    anchor: {
      ...anchorWithDefaults(fieldsOf(object).anchor),
      rotation: next,
    },
  };
}

/** The members of the SmartArt graphic `objectId` belongs to, or null. */
export function sheetSmartArtGroupAt<O extends DrawingObjectLike>(
  objects: readonly O[],
  objectId: string | null,
): { marker: SmartArtBlockMarker; members: O[] } | null {
  const primary = objectId
    ? objects.find((object) => object.id === objectId)
    : undefined;
  if (!primary) return null;
  const group = readSmartArtGroupMembers(
    objects as unknown as readonly OfficeSmartArtHostNode[],
    primary as unknown as OfficeSmartArtHostNode,
  );
  if (!group) return null;
  const ids = new Set(group.members.map((member) => member.id));
  const members = objects.filter((object) => ids.has(object.id));
  return members.length === group.members.length
    ? { marker: group.marker, members }
    : null;
}

/**
 * Shape Format > Alt Text: the sheet's objects with the shape's description
 * (control characters out, bounded). A SmartArt graphic has ONE Alt Text,
 * set on the graphic and not on a shape of it (Microsoft Support, "Add
 * alternative text to a shape, picture, chart, SmartArt graphic, or other
 * object"); the worksheet has no object for the graphic, so every member
 * carries it. Null when nothing changes.
 */
export function objectsWithShapeAltText<O extends DrawingObjectLike>(
  objects: readonly O[],
  objectId: string,
  parameter: string,
): O[] | null {
  const object = objects.find((item) => item.id === objectId);
  if (!object) return null;
  // The description keeps its own line breaks (S-089: two lines were joined
  // into one word), as a picture's does; other control characters go.
  const next = normalizedTableAltTextInput(parameter.replace(/\r\n?/g, "\n"), MAX_PICTURE_ALT_TEXT_LENGTH)
    .replace(/[\u007f-\u009f]/g, "");
  if (next === (fieldsOf(object).alt ?? "")) return null;
  const group = sheetSmartArtGroupAt(objects, objectId);
  const members = new Set(group ? group.members.map((member) => member.id) : [objectId]);
  return objects.map((item) => (members.has(item.id) ? { ...item, alt: next } : item));
}

/**
 * Picture Format > Alt Text (the description) and the Selection pane's name
 * (`caption`): a field not given is left exactly as it is. Null when
 * nothing changes.
 */
export function pictureAltTextPatch(
  object: object,
  nextAlt: string | undefined,
  nextCaption: string | undefined,
): DrawingObjectPatch | null {
  // A field the caller does not give is left exactly as it is: the
  // Alt Text pane writes the description alone and never touches
  // the picture's name (`caption`), which the Selection and
  // Navigation panes' Rename write alone.
  const patch: DrawingObjectPatch = {};
  if (nextAlt !== undefined) {
    const value = normalizedTableAltTextInput(
      nextAlt,
      MAX_PICTURE_ALT_TEXT_LENGTH,
    );
    if (value !== (fieldsOf(object).alt ?? "")) patch.alt = value;
  }
  if (nextCaption !== undefined) {
    const value = normalizedTableAltTextInput(
      nextCaption,
      MAX_PICTURE_NAME_LENGTH,
    );
    if (value !== (fieldsOf(object).caption ?? "")) {
      patch.caption = value;
    }
  }
  return Object.keys(patch).length === 0 ? null : patch;
}

/** Picture Format > Crop: the DrawingML keep rectangle as JSON; the object
 * as it is after (an identity crop removes the crop). Null when unchanged. */
export function pictureCroppedObject<O extends object>(object: O, parameter: string): O | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  const next = canonicalizeDrawingMlCropRect(parsed);
  if (!next || sameCropRect(next, fieldsOf(object).cropRect)) return null;
  const nextObject = { ...object } as O & { cropRect?: unknown };
  if (isIdentityCrop(next)) delete nextObject.cropRect;
  else nextObject.cropRect = next;
  return nextObject;
}

/** The crop a picture holds, as its crop handles start from (a Fit's outset kept). */
export function pictureCropStart(cropRect: unknown): CropRect {
  return normalizeSignedCropRect(cropRect) ?? clampCropRect(cropRect);
}

/**
 * A crop handle's request, `{"crop": {...}, "frame": "follow"}`: the crop to
 * apply with the frame moving by it. Null for a plain crop (the pane, Fit,
 * Fill, aspect ratio), which keeps the frame.
 */
export function pictureCropHandleRequest(parameter: string): CropRect | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(parameter);
  } catch {
    return null;
  }
  if (!parsed || typeof parsed !== "object" || (parsed as { frame?: unknown }).frame !== "follow") return null;
  return canonicalizeDrawingMlCropRect((parsed as { crop?: unknown }).crop);
}

/**
 * The frame a crop handle leaves (Excel and PowerPoint: "drag inward on the
 * side cropping handle"). The source keeps the scale it is drawn at, so the
 * frame's edges move with the crop's: 40 px cropped off the left moves the
 * left edge 40 px in and nothing stretches. `bounds` is the unturned box; a
 * turned or mirrored picture moves along its own sides.
 */
export function pictureCropFrameBounds(
  bounds: SpreadsheetObjectBounds,
  from: CropRect,
  to: CropRect,
  rotationDegrees = 0,
  flip: { flipH?: boolean; flipV?: boolean } = {},
): SpreadsheetObjectBounds {
  if (!(from.w > 0) || !(from.h > 0)) return bounds;
  const pxX = bounds.width / from.w;
  const pxY = bounds.height / from.h;
  const width = to.w * pxX;
  const height = to.h * pxY;
  // The centre's move in the picture's own axes, before its mirror and turn.
  let cx = (to.x - from.x + (to.w - from.w) / 2) * pxX;
  let cy = (to.y - from.y + (to.h - from.h) / 2) * pxY;
  if (flip.flipH) cx = -cx;
  if (flip.flipV) cy = -cy;
  const radians = (rotationDegrees * Math.PI) / 180;
  const centreX = bounds.left + bounds.width / 2 + cx * Math.cos(radians) - cy * Math.sin(radians);
  const centreY = bounds.top + bounds.height / 2 + cx * Math.sin(radians) + cy * Math.cos(radians);
  return { left: centreX - width / 2, top: centreY - height / 2, width, height };
}

/** Picture Format > Picture Styles: one of Office's 28 styles, by id. */
export function pictureStylePatch(object: object, parameter: string): DrawingObjectPatch | null {
  if (fieldsOf(object).nativePictureStyleBlocker) return null;
  const next = parameter;
  if (!isOfficePictureStyleId(next)) return null;
  if (fieldsOf(object).pictureStyle === next) return null;
  return {
    pictureStyle: next,
  };
}

/** Delete: the objects without `objectId`, or null when it is not there. */
export function objectsWithout<O extends DrawingObjectLike>(objects: readonly O[], objectId: string): O[] | null {
  const next = objects.filter(
    (object) => object.id !== objectId,
  );
  return next.length === objects.length ? null : next;
}

/**
 * Arrange > Bring Forward / Bring to Front / Send Backward / Send to Back on
 * the sheet-wide stack (pictures, shapes AND native charts, live audit
 * B161): both arrays come back in stack order and every drawing carries its
 * position. A native chart with no stored x/y sits at a default that depends
 * on its array index (`chartPosition`), so a chart array that changes order
 * first pins each chart where it is. Null when the move changes nothing.
 */
export function sheetArrangedForDrawing<S extends object>(
  base: S,
  kind: DrawingStackKind,
  id: string,
  move: DrawingStackMove,
  chartPosition: (chart: StackedChart, ordinal: number) => { x: number; y: number },
  /** The Selection Pane's row of a group member: it moves among its group only. */
  options: { withinGroup?: boolean } = {},
): S | null {
  const sheet = base as { objects?: readonly DrawingObjectLike[]; charts?: readonly StackedChart[] };
  const objects: readonly DrawingObjectLike[] = sheet.objects ?? [];
  const charts: readonly StackedChart[] = sheet.charts ?? [];
  // A group (pictures, shapes and charts grouped together, or a SmartArt
  // graphic's members) moves as one, its members in their own order.
  const groups = new Map<string, string>();
  for (const object of objects) {
    const group = chartDrawingGroupId((object as { groupId?: unknown }).groupId);
    if (group) groups.set(`object:${object.id}`, group);
  }
  for (const chart of charts) {
    const group = chartDrawingGroupId(chart.groupId);
    if (group) groups.set(`chart:${chart.id}`, group);
  }
  const moved = moveInDrawingStack(
    sheetDrawingStack(objects, charts),
    kind,
    id,
    move,
    (entry) => groups.get(`${entry.kind}:${entry.id}`),
    options.withinGroup === true,
  );
  if (!moved) return null;
  return sheetWithDrawingStack(base, moved, chartPosition);
}

/**
 * The sheet with its drawings restacked to `stack` (bottom first): both
 * arrays in stack order, every drawing carrying its position, a reordered
 * chart array pinned where each chart is (see `sheetArrangedForDrawing`).
 * Shared by the Arrange moves and Reorder Objects.
 */
export function sheetWithDrawingStack<S extends object>(
  base: S,
  stack: readonly DrawingStackEntry[],
  chartPosition: (chart: StackedChart, ordinal: number) => { x: number; y: number },
): S {
  const sheet = base as { objects?: readonly DrawingObjectLike[]; charts?: readonly StackedChart[] };
  const objects: readonly DrawingObjectLike[] = sheet.objects ?? [];
  const charts: readonly StackedChart[] = sheet.charts ?? [];
  let applied = applyDrawingStack(stack, objects, charts);
  if (applied.chartsReordered) {
    applied = applyDrawingStack(
      stack,
      objects,
      charts.map((chart, ordinal) => ({ ...chart, ...chartPosition(chart, ordinal) })),
    );
  }
  return {
    ...base,
    ...(sheet.objects || applied.objects.length ? { objects: applied.objects } : {}),
    ...(sheet.charts || applied.charts.length ? { charts: applied.charts } : {}),
  } as S;
}

/** A floating chart as the drawing stack sees it. */
export interface StackedChart {
  id: string;
  [key: string]: unknown;
}

// ---------------------------------------------------------------------------
// Insert > Shapes, Text Box and WordArt.

/**
 * Insert Shape (Excel/WPS "Insert ▸ Shapes"): a drawing shape of the chosen
 * kind at the cell, in the theme's shape visual, or the Set as Default
 * Shape formatting when one is stored. A text box is an invisible text
 * container, not a themed AutoShape: it keeps its factory fill:none /
 * stroke:none under every workbook theme.
 */
export function insertedShapeObject(
  kind: SpreadsheetPresetShapeKind,
  cell: { row: number; col: number },
  themeVisual: SpreadsheetThemeShapeVisual,
  defaults: OfficeDefaultShapeStyle | null,
  /**
   * Excel's click-and-drag insertion: the box the user drew, in stage CSS
   * pixels, as an absolute anchor (the one a Draw-tab stroke takes). Without
   * it the shape anchors at `cell` at its factory size.
   */
  frame?: { left: number; top: number; width: number; height: number; flipV?: boolean },
  






  grid?: DrawingGrid,
): SpreadsheetShapeObject {
  const anchored = createSpreadsheetShapeObject(kind, {
    row: cell.row,
    col: cell.col,
  });
  const placed: SpreadsheetShapeObject = frame
    ? {
        ...anchored,
        ...(frame.flipV ? { shape: normalizeShapeStyle({ ...anchored.shape, flipV: true }) } : {}),
        anchor: {
          kind: "absolute",
          xEmu: Math.max(0, Math.round(frame.left * 9_525)),
          yEmu: Math.max(0, Math.round(frame.top * 9_525)),
          widthEmu: Math.max(1, Math.round(frame.width * 9_525)),
          heightEmu: Math.max(1, Math.round(frame.height * 9_525)),
        },
      }
    : anchored;
  const baseShapeObject: SpreadsheetShapeObject = grid
    ? ({
        ...placed,
        ...(objectPlacementPatch(placed, "moveAndSize", grid) ?? {}),
      } as SpreadsheetShapeObject)
    : placed;
  // A text box is an invisible text container, not a themed AutoShape.
  // Preserve its factory-owned fill:none / stroke:none visual under every
  // workbook theme, including effect themes that would otherwise add an
  // unintended shadow. Ordinary drawing shapes continue to consume the
  // live theme visual at insertion time.
  return kind === "text-box"
    ? baseShapeObject
    : {
        ...baseShapeObject,
        shape: normalizeShapeStyle({
          ...baseShapeObject.shape,
          fill: defaults?.fill ?? themeVisual.fill,
          stroke: defaults?.outline ?? themeVisual.stroke,
          strokeWidth:
            defaults?.outlineWidthPt !== undefined
              ? officeOutlinePtToPx(defaults.outlineWidthPt)
              : themeVisual.strokeWidth,
          ...(defaults?.textPaint ? { textPaint: defaults.textPaint } : {}),
        }),
        ...(themeVisual.shadow ? { shadow: themeVisual.shadow } : {}),
      };
}

/**
 * Excel's Insert > WordArt: the text-box factory object (invisible
 * container) with the text and ONE shared WordArt recipe as its text paint --
 * the same object the Shape Format tab's WordArt Styles edit.
 */
export function insertedWordArtObject(
  recipe: OfficeWordArtRecipe,
  cell: { row: number; col: number },
  text: string,
): ReturnType<typeof createSpreadsheetShapeObject> {
  const base = createSpreadsheetShapeObject("text-box", {
    row: cell.row,
    col: cell.col,
  });
  return {
    ...base,
    shape: normalizeShapeStyle({
      ...base.shape,
      text,
      textSize: 36,
      textPaint: recipe.paint,
    }),
  };
}

/** A drawing object a move or resize may reach (connectors carry glue). */
export interface TransformableObject extends DrawingObjectLike {
  type: string;
  shape?: unknown;
  hidden?: unknown;
  startObjectId?: unknown;
  endObjectId?: unknown;
  startAnchor?: unknown;
  endAnchor?: unknown;
}

/**
 * One drawing object moved, resized or rotated to `bounds` (pixels in the
 * grid frame, headers included), as a drag ends: its anchor follows in its
 * own kind, a SmartArt shape edited on its own carries Office's customization
 * factors, a connector dropped with an end on a shape's connection site joins
 * it (one moved off lets go), and every connector glued at both ends follows
 * its shapes. Shared by the window's drag and the spreadsheet editor commands.
 */
export function objectsAfterBoxEdit<O extends TransformableObject>(
  objects: readonly O[],
  source: O,
  bounds: SpreadsheetObjectBounds,
  rotation: number | undefined,
  grid: DrawingGrid,
  /** The object's mirror flags after a resize carried through the pinned
   * edge (Excel flips the shape in the same gesture); a false flag is deleted. */
  flip?: { flipH: boolean; flipV: boolean },
): O[] {
  const nextAnchor = anchorForObjectBounds(
    anchorWithDefaults(source.anchor),
    bounds,
    grid.rowCount,
    grid.columnCount,
    grid.colWidth,
    grid.rowHeight,
  );
  if (rotation !== undefined) nextAnchor.rotation = rotation;
  if (flip) {
    for (const axis of ["flipH", "flipV"] as const) {
      if (flip[axis]) nextAnchor[axis] = true;
      else delete nextAnchor[axis];
    }
  }
  const rectOf = (object: O) => objectBoundsPixels(object, grid.colWidth, grid.rowHeight);
  let nextObjects: readonly O[] = objects.map((candidate) =>
    candidate === source
      ? smartArtMemberAfterBoxEdit(candidate, { ...candidate, anchor: nextAnchor }, rectOf)
      : candidate,
  );
  // The arrows and lines that join a moved SmartArt shape follow it, as
  // Office's connector algorithm re-routes them (`connector-follow`).
  const diagram = sheetSmartArtGroupAt(objects, source.id);
  if (diagram && diagram.marker.role === "node") {
    const ids = new Set(diagram.members.map((member) => member.id));
    const followed = smartArtConnectorsFollowNodes(
      diagram.marker.kind,
      diagram.members,
      nextObjects.filter((object) => ids.has(object.id)),
      {
        id: (object) => object.id,
        box: rectOf,
        rotation: (object) => {
          const turn = (object.anchor as { rotation?: unknown } | undefined)?.rotation;
          return typeof turn === "number" && Number.isFinite(turn) ? turn : 0;
        },
        shape: (object) => (object.shape as { kind?: unknown } | undefined)?.kind,
        write: (object, box, turn) => {
          const anchor = anchorForObjectBounds(
            anchorWithDefaults(object.anchor),
            box,
            grid.rowCount,
            grid.columnCount,
            grid.colWidth,
            grid.rowHeight,
          );
          if (turn !== 0 || anchor.rotation !== undefined) anchor.rotation = turn;
          return { ...object, anchor };
        },
      },
    );
    const byId = new Map(followed.map((object) => [object.id, object]));
    nextObjects = nextObjects.map((object) => byId.get(object.id) ?? object);
  }
  const connectorKind = (object: O) => {
    if (object.type !== "shape" || !object.shape) return false;
    const kind = normalizeShapeStyle(object.shape).kind;
    return isSpreadsheetPresetShapeKind(kind) && isSpreadsheetLineLikeShapeKind(kind);
  };
  const moved = nextObjects.find((object) => object.id === source.id);
  if (moved && rotation === undefined && connectorKind(moved)) {
    const targets = nextObjects
      .filter((object) => object.id !== moved.id && object.type === "shape" && !connectorKind(object) && !object.hidden)
      .map((object) => ({ id: object.id, rect: rectOf(object) }));
    const glue = spreadsheetConnectorGlueAfterMove(
      rectOf(moved),
      moved.anchor?.flipH === true,
      moved.anchor?.flipV === true,
      targets,
    );
    const reglued: O = { ...moved };
    delete reglued.startObjectId;
    delete reglued.endObjectId;
    delete reglued.startAnchor;
    delete reglued.endAnchor;
    const joined: O = { ...reglued, ...glue };
    nextObjects = nextObjects.map((object) => (object === moved ? joined : object));
  }
  return [
    ...recomputeSpreadsheetGluedConnectors(nextObjects, rectOf, (connector, box, flipH, flipV) => ({
      ...connector,
      anchor: {
        ...anchorForObjectBounds(
          anchorWithDefaults(connector.anchor),
          box,
          grid.rowCount,
          grid.columnCount,
          grid.colWidth,
          grid.rowHeight,
        ),
        ...(flipH ? { flipH: true } : { flipH: undefined }),
        ...(flipV ? { flipV: true } : { flipV: undefined }),
      },
    })),
  ];
}
