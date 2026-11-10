/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  drawingmlDefaultPresetAdjustNames,
  evaluateDrawingmlDefaultPreset,
} from "./drawingml-default-preset-program";
import { DEFAULT_PRESET_PROGRAMS } from "./presentation-default-preset-programs.generated";
import type { NativeDrawingGeometry } from "./native-drawing-geometry";
import { presentationShapeCatalogEntry } from "./presentation-shape-catalog";
import {
  PRESENTATION_CONNECTOR_PRESETS,
  type PresentationConnectorPreset,
  type ShapeKind,
} from "./presentation-shapes";

const CONNECTOR_PRESET_SET: ReadonlySet<string> = new Set(
  PRESENTATION_CONNECTOR_PRESETS,
);

/**
 * The route one DrawingML connector preset draws, evaluated in the connector's
 * own box.
 *
 * A connector's `@prst` IS its route: `bentConnector4` is four segments with
 * three bends, `curvedConnector5` is four cubics. Hand-written route strings
 * can express the three-segment default and nothing else, which is why the
 * other six had no representation at all. Ends live on `a:ln` and are drawn by
 * the caller's markers, never here.
 */
export function presentationConnectorRouteGeometry(
  connectorPreset: PresentationConnectorPreset,
  widthPixels: number,
  heightPixels: number,
  /**
   * The route's OWN dragged handles (`a:avLst`), substituted by name exactly
   * as {@link presentationPresetGeometryForKind} substitutes a shape's: an
   * elbow whose bend was dragged draws the bend where it was dragged.
   */
  adjustments?: Readonly<Record<string, number>> | null,
): NativeDrawingGeometry | null {
  if (!CONNECTOR_PRESET_SET.has(connectorPreset)) return null;
  if (!Number.isFinite(widthPixels) || !Number.isFinite(heightPixels) || widthPixels <= 0 || heightPixels <= 0) return null;
  const preset = DEFAULT_PRESET_PROGRAMS[connectorPreset] as { program: unknown } | undefined;
  return preset
    ? evaluateDrawingmlDefaultPreset(preset.program, widthPixels * 9525, heightPixels * 9525, adjustments)
    : null;
}

/** Local catalog defaults only; never accepts a program from document data.
 *
 * `adjustments` are the shape's OWN dragged handles, read off its `a:avLst` at
 * import. They are substituted into the catalog program by name, which is how
 * the server evaluates the same shape (`file_adapters/pptx/native_geometry`);
 * the program itself always comes from the catalog, never from the document.
 */
export function presentationPresetGeometryForKind(
  kind: ShapeKind,
  widthPixels: number,
  heightPixels: number,
  adjustments?: Readonly<Record<string, number>> | null,
): NativeDrawingGeometry | null {
  if (!Number.isFinite(widthPixels) || !Number.isFinite(heightPixels) || widthPixels <= 0 || heightPixels <= 0) return null;
  const entry = presentationShapeCatalogEntry(kind);
  if (entry?.geometry !== "preset") return null;
  const preset = DEFAULT_PRESET_PROGRAMS[entry.drawingMlPreset] as { program: unknown } | undefined;
  return preset
    ? evaluateDrawingmlDefaultPreset(preset.program, widthPixels * 9525, heightPixels * 9525, adjustments)
    : null;
}

/** The adjust handles one DrawingML preset declares, or null for no such preset. */
export function presentationPresetAdjustNames(
  drawingMlPreset: string,
): readonly string[] | null {
  const preset = DEFAULT_PRESET_PROGRAMS[drawingMlPreset] as { program: unknown } | undefined;
  return preset ? drawingmlDefaultPresetAdjustNames(preset.program) : null;
}

/** Newly authored presets have no imported adjustments or cached geometry.
 * Retain the imported renderer whenever any source or explicit geometry exists.
 * This result is a view, not native authoring authority or persisted metadata.
 */
export function presentationAuthoredPresetGeometry(
  block: Record<string, unknown>,
  widthPixels: number,
  heightPixels: number,
): NativeDrawingGeometry | null {
  if (block.type !== "shape" || typeof block.shape !== "string" ||
    ["source", "points", "customPath"].some(key => Object.prototype.hasOwnProperty.call(block, key))) return null;
  return presentationPresetGeometryForKind(block.shape as ShapeKind, widthPixels, heightPixels);
}
