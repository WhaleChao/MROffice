/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The DrawingML preset a worksheet shape is drawn from: what its view draws
 * (SpreadsheetShapeView), what its adjustment handles and Edit Points read,
 * and what the AI's spreadsheet.object.shape_adjust pins its values against.
 */

import { isOfficeShapePreset, officeShapeDefinition } from "@mrnote/core/web/lib/office-shape-catalog.ts";
import { presentationShapeCatalogEntry } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import { normalizeShapeKind as normalizePresentationShapeKind, type ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { SPREADSHEET_CUSTOM_SHAPE_KIND, type SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

/** The shared preview kind a sheet shape kind draws as. */
export function spreadsheetShapePreviewKind(kind: SpreadsheetShapeStyle["kind"]): ShapeKind {
  // "custom-geometry" draws from its own points; the rectangle is only what a
  // shape whose points failed to normalize falls back to.
  if (kind === "text-box" || kind === "ink" || kind === "custom-geometry") {
    return "rectangle";
  }
  return isOfficeShapePreset(kind)
    ? officeShapeDefinition(kind).previewKind
    : normalizePresentationShapeKind(kind);
}

/** The DrawingML preset of a sheet shape, or null for a text box, ink or a custom outline. */
export function spreadsheetShapeDrawingPreset(shape: SpreadsheetShapeStyle): string | null {
  if (shape.kind === "text-box" || shape.kind === "ink" || shape.kind === SPREADSHEET_CUSTOM_SHAPE_KIND || shape.customPath) {
    return null;
  }
  const entry = presentationShapeCatalogEntry(spreadsheetShapePreviewKind(shape.kind));
  return entry?.geometry === "preset" ? entry.drawingMlPreset : null;
}
