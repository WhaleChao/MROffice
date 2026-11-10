/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture's own preset geometry: Picture Format > Crop > Crop to Shape.
 *
 * Office semantics (Microsoft Support, "Crop a picture to fit in a shape";
 * ISO/IEC 29500-1 §20.1.9.18 `a:prstGeom`, §20.2.2.6 `pic:spPr`):
 * - Crop to Shape replaces the picture's `a:prstGeom/@prst` and nothing else.
 *   The picture keeps its frame (extent), its crop (`a:srcRect`) and any
 *   picture style paint; the shape is stretched to the frame, so an ellipse
 *   on a 4:3 picture is a 4:3 ellipse (verified against Word 16 for Mac).
 * - The kept source fills the frame as before and only the part inside the
 *   shape's filled paths is drawn; a preset's holes (a donut) stay empty.
 * - The preset's handles live in `a:avLst`; Crop to Shape writes the
 *   defaults (an empty list), an imported picture may carry dragged values.
 * - Choosing Rectangle draws the ordinary picture again.
 *
 * The geometry comes from the one trusted DrawingML preset catalog the deck,
 * the sheet and the document shapes already evaluate (all 187
 * `ST_ShapeType` presets), never from document data.
 */

import { evaluateDrawingmlDefaultPreset } from "./drawingml-default-preset-program";
import { presentationPresetAdjustNames } from "./presentation-authored-preset-geometry";
import { DEFAULT_PRESET_PROGRAMS } from "./presentation-default-preset-programs.generated";

/** The plain picture's `@prst`. */
export const OFFICE_PICTURE_RECTANGLE_GEOMETRY = "rect";

const MAX_PICTURE_GEOMETRY_ADJUSTMENTS = 8;

/** One `ST_ShapeType` name the preset catalog knows, else null. */
export function normalizeOfficePictureGeometry(value: unknown): string | null {
  return typeof value === "string" &&
    /^[A-Za-z0-9]{1,40}$/.test(value) &&
    Object.prototype.hasOwnProperty.call(DEFAULT_PRESET_PROGRAMS, value)
    ? value
    : null;
}

/**
 * The picture geometry's `a:avLst` values by handle name, or null for the
 * preset's defaults. Only literal integers for handles this preset declares
 * survive; anything else is not a statement about this preset.
 */
export function normalizeOfficePictureGeometryAdjustments(
  preset: unknown,
  value: unknown,
): Readonly<Record<string, number>> | null {
  const geometry = normalizeOfficePictureGeometry(preset);
  if (!geometry || !value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const declared = presentationPresetAdjustNames(geometry) ?? [];
  const entries = Object.entries(value as Record<string, unknown>);
  if (entries.length === 0 || entries.length > MAX_PICTURE_GEOMETRY_ADJUSTMENTS) {
    return null;
  }
  const adjustments: Record<string, number> = {};
  for (const [name, amount] of entries) {
    if (
      !declared.includes(name) ||
      typeof amount !== "number" ||
      !Number.isSafeInteger(amount) ||
      Math.abs(amount) > 2_147_483_647
    ) {
      return null;
    }
    adjustments[name] = amount;
  }
  return adjustments;
}

/**
 * The SVG path data of the picture's shape at its drawn size, in that box's
 * own pixels: every filled path of the preset, joined (DrawingML fills each
 * path; their union is what the picture shows through). Null for the plain
 * rectangle, a degenerate box or a geometry the catalog cannot evaluate.
 */
export function officePictureGeometryPath(
  preset: unknown,
  adjustments: unknown,
  widthPx: number,
  heightPx: number,
): string | null {
  const geometry = normalizeOfficePictureGeometry(preset);
  if (
    !geometry ||
    geometry === OFFICE_PICTURE_RECTANGLE_GEOMETRY ||
    !Number.isFinite(widthPx) ||
    !Number.isFinite(heightPx) ||
    widthPx <= 0 ||
    heightPx <= 0
  ) {
    return null;
  }
  const program = DEFAULT_PRESET_PROGRAMS[geometry] as { program: unknown } | undefined;
  if (!program) return null;
  const evaluated = evaluateDrawingmlDefaultPreset(
    program.program,
    widthPx * 9525,
    heightPx * 9525,
    normalizeOfficePictureGeometryAdjustments(geometry, adjustments),
  );
  if (!evaluated) return null;
  const filled = evaluated.paths.filter((path) => path.fill !== "none");
  return filled.length ? filled.map((path) => path.d).join(" ") : null;
}

/** The same shape as a CSS `clip-path` value, or "" for the plain picture. */
export function officePictureGeometryClipPath(
  preset: unknown,
  adjustments: unknown,
  widthPx: number,
  heightPx: number,
): string {
  const d = officePictureGeometryPath(preset, adjustments, widthPx, heightPx);
  return d ? `path(${JSON.stringify(d)})` : "";
}
