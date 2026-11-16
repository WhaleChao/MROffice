/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isNativePresentationThemePartPath } from "./presentation-theme";
import { readNativeDrawingGeometry } from "./native-drawing-geometry";
import { normalizeOfficeShapeGradient, normalizeOfficeShapeShadow, type OfficeShapeGradient, officeGradientAngle } from "./office-shape-paint";

const record = (value: unknown): Record<string, unknown> | null =>
  value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;






export function presentationHasNativeShapeDrawing(value: unknown): boolean {
  const block = record(value), source = record(block?.source), drawing = record(source?.nativeShapeDrawing);
  const inherited = source?.inheritedFrom === "layout" || source?.inheritedFrom === "master";
  return block?.type === "shape" && typeof source?.slideId === "string" && (typeof source?.shapeId === "string" || inherited) &&
    drawing?.schemaVersion === "pptx-shape-drawing-v1" && drawing.shapeKind === block.shape &&
    readNativeDrawingGeometry(drawing.geometry) !== null;
}

export interface PresentationNativeThemePaint {
  fill?: string;
  fillTo?: string;
  fillGradient?: OfficeShapeGradient;
  gradientAngle?: number;
  stroke?: string;
  strokeWidth?: number;
}

/** Bounded display-only receipt; the API authenticates it against native bytes. */
export function readPresentationNativeThemePaint(value: unknown): PresentationNativeThemePaint | undefined {
  if (!presentationHasNativeShapeDrawing(value)) return undefined;
  const source = record(record(value)?.source), drawing = record(source?.nativeShapeDrawing);
  const receipt = record(drawing?.themePaint), paint = record(receipt?.paint);
  if (receipt?.schemaVersion !== "pptx-native-theme-paint-v1" || !paint ||
    typeof receipt.themePart !== "string" || (!receipt.themePart.startsWith("/") || !isNativePresentationThemePartPath(receipt.themePart.slice(1))) ||
    ["themeSha256", "shapeStyleSha256", "propertiesSha256", "colorMapSha256"].some(k => typeof receipt[k] !== "string" || !/^[0-9a-f]{64}$/.test(receipt[k] as string)) ||
    Object.keys(paint).some(k => !["fill", "fillTo", "fillGradient", "gradientAngle", "stroke", "strokeWidth"].includes(k))) return undefined;
  for (const k of ["fill", "stroke", "fillTo"])
    if (paint[k] !== undefined && (typeof paint[k] !== "string" || !/^(?:#[0-9a-f]{6}|none)$/i.test(paint[k] as string))) return undefined;
  if (paint.strokeWidth !== undefined && (typeof paint.strokeWidth !== "number" || !Number.isFinite(paint.strokeWidth) || paint.strokeWidth < 0 || paint.strokeWidth > 1000)) return undefined;
  try {
    const gradient = normalizeOfficeShapeGradient(paint.fillGradient);
    if (gradient && (paint.fill !== gradient.stops[0].color || paint.fillTo !== gradient.stops.at(-1)?.color || paint.gradientAngle !== officeGradientAngle(gradient))) return undefined;
    if (!gradient && (paint.fillTo !== undefined || paint.gradientAngle !== undefined)) return undefined;
    return { ...paint, ...(gradient ? { fillGradient: gradient } : {}) } as PresentationNativeThemePaint;
  } catch { return undefined; }
}

/** Page-fixed shadow for a proven rigid, preserve-only group member. */
export function readPresentationNativeThemeShadow(value: unknown) {
  if (!presentationHasNativeShapeDrawing(value)) return undefined;
  const block = record(value)!, source = record(block.source)!, drawing = record(source.nativeShapeDrawing)!;
  if (typeof block.groupId !== "string" || !block.groupId || "nativeGroupAuthority" in source ||
    block.shadow !== undefined || block.shapeShadow !== undefined) return undefined;
  const receipt = record(drawing.themeShadow), geometry = readNativeDrawingGeometry(drawing.geometry);
  if (receipt?.schemaVersion !== "pptx-native-theme-shadow-v1" || receipt.profile !== "rigid-group-page-shadow-v1" || !geometry?.viewBox ||
    typeof receipt.themePart !== "string" || (!receipt.themePart.startsWith("/") || !isNativePresentationThemePartPath(receipt.themePart.slice(1))) ||
    ["themeSha256", "shapeStyleSha256", "propertiesSha256", "colorMapSha256", "groupTransformSha256"].some(k => typeof receipt[k] !== "string" || !/^[0-9a-f]{64}$/.test(receipt[k] as string))) return undefined;
  const [width, height] = geometry.viewBox;
  if (typeof block.width !== "number" || typeof block.height !== "number" || !Number.isFinite(block.width) || !Number.isFinite(block.height) ||
    Math.abs(block.width / 9525 - width) > 1e-6 || Math.abs(block.height / 9525 - height) > 1e-6) return undefined;
  const rotation = block.rotation ?? 0;
  if (typeof rotation !== "number" || !Number.isFinite(rotation)) return undefined;
  try {
    const shadow = normalizeOfficeShapeShadow(receipt.shadow);
    if (!shadow) return undefined;
    const radians = (rotation % 360) * Math.PI / 180;
    // Inverse rigid rotation of the page-space downward offset. The source
    // profile and 1:1 viewport check exclude scaling/skew and resized views.
    const dx = shadow.distance * Math.sin(radians), dy = shadow.distance * Math.cos(radians);
    const padding = shadow.blur * 1.5 + shadow.distance + 1;
    return { shadow, dx, dy, bounds: { x: -padding, y: -padding, width: width + 2 * padding, height: height + 2 * padding } };
  } catch { return undefined; }
}
