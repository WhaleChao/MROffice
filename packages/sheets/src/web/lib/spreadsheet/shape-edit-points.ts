/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Edit Points on a worksheet shape: the outline the sheet's Edit Points
 * overlay (SpreadsheetWindow) edits and the AI's
 * spreadsheet.object.shape_edit_points replaces, and the one write both make.
 */

import {
  officeCustomPathEditBox,
  officeEditPathFromCustomPath,
  officeEditPathFromPreset,
  type OfficeEditPath,
} from "@mrnote/core/web/lib/office-custom-path-edit.ts";
import { normalizePresentationCustomPath } from "@mrnote/core/web/lib/presentation-freeform.ts";
import {
  anchorForObjectBounds,
  anchorWithDefaults,
  objectBoundsPixels,
  spreadsheetObjectRotation,
  type DrawingObjectLike,
  type SpreadsheetObjectAnchor,
} from "./object-buttons";
import { spreadsheetShapeDrawingPreset } from "./shape-preset";
import { normalizeShapeStyle, SPREADSHEET_CUSTOM_SHAPE_KIND, type SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

/**
 * Where a sheet shape's drawn geometry sits inside its element, as fractions
 * of the element's box: the whole box, as Excel paints a shape to its anchored
 * box. A preset is drawn by the shared preview sized out past the element by
 * the preview's own bleed (`officeShapePreviewBleedFrame`), so its geometry
 * lands on the element; a custom outline fills the element by itself. The
 * adjustment handles, Edit Points and the label's text rectangle are laid over
 * exactly that box.
 */
export function spreadsheetShapeGeometryInset(_shape: SpreadsheetShapeStyle) {
  return { left: 0, top: 0, right: 0, bottom: 0 };
}

/** The outline Edit Points edits, in the geometry box's own pixels. */
export function spreadsheetShapeEditPath(
  shape: SpreadsheetShapeStyle,
  width: number,
  height: number,
): OfficeEditPath | null {
  if (shape.kind === "text-box" || shape.kind === "ink" || !(width > 0) || !(height > 0)) return null;
  if (shape.customPath) return officeEditPathFromCustomPath(shape.customPath, width, height);
  const preset = spreadsheetShapeDrawingPreset(shape);
  return preset ? officeEditPathFromPreset(preset, width, height, shape.adjustments) : null;
}

export interface SpreadsheetShapeEditGrid {
  readonly colWidth: (column: number) => number;
  readonly rowHeight: (row: number) => number;
  readonly rowCount: number;
  readonly columnCount: number;
}

/**
 * One finished Edit Points edit (`officeCustomPathFromEdit`'s result: the
 * outline and its box relative to the drawn geometry's box): the shape
 * becomes a custom outline, its adjust values and points dropped, and its
 * anchor follows the outline. Null for a text box or ink, or an edit that
 * names no stored outline.
 */
export function spreadsheetShapeEditedPathPatch(
  object: DrawingObjectLike & { shape?: unknown },
  edit: { customPath: unknown; left: number; top: number; width: number; height: number },
  grid: SpreadsheetShapeEditGrid,
): { shape: SpreadsheetShapeStyle; anchor: SpreadsheetObjectAnchor } | null {
  const customPath = normalizePresentationCustomPath(edit.customPath);
  const { left, top, width, height } = edit;
  if (!customPath || ![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;
  const current = normalizeShapeStyle(object.shape);
  if (current.kind === "text-box" || current.kind === "ink") return null;
  // The edit was made in the drawn geometry's box; the new element box is the
  // one a custom outline fills (no preset bleed).
  const bounds = objectBoundsPixels(object, grid.colWidth, grid.rowHeight);
  const inset = spreadsheetShapeGeometryInset(current);
  const inner = {
    x: bounds.left + inset.left * bounds.width,
    y: bounds.top + inset.top * bounds.height,
    width: bounds.width * (1 - inset.left - inset.right),
    height: bounds.height * (1 - inset.top - inset.bottom),
  };
  const box = officeCustomPathEditBox(
    inner,
    { left, top, width, height },
    spreadsheetObjectRotation(object.anchor),
    object.anchor?.flipH === true,
    object.anchor?.flipV === true,
  );
  // The adjust values and points described the preset, not this outline.
  const rest: Partial<SpreadsheetShapeStyle> = { ...current };
  delete rest.adjustments;
  delete rest.customPoints;
  return {
    shape: normalizeShapeStyle({ ...rest, kind: SPREADSHEET_CUSTOM_SHAPE_KIND, customPath }),
    anchor: anchorForObjectBounds(
      anchorWithDefaults(object.anchor),
      {
        left: Math.max(0, box.x),
        top: Math.max(0, box.y),
        width: Math.max(1, box.width),
        height: Math.max(1, box.height),
      },
      grid.rowCount,
      grid.columnCount,
      grid.colWidth,
      grid.rowHeight,
    ),
  };
}
