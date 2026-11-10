/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Crop > Crop to Shape (PPT-A132; Microsoft Support, "Crop a
 * picture to fit in a shape"): the picture keeps its frame, crop and
 * formatting and draws inside the outline of a preset shape -- its own
 * `p:spPr/a:prstGeom@prst`, default adjustments. The page keeps the preset
 * name as `clipShape` (the importer's `ellipse` since PPT-B364, any closed
 * preset now); a rectangle is no clip at all.
 *
 * The outline is the shared picture geometry Docs crops with
 * (`lib/office-picture-geometry`: the DrawingML preset program evaluated at
 * the picture's own size), so a rounded or snipped corner keeps its size as
 * the frame stretches. This module adds only the deck's model (`clipShape`,
 * gallery kinds) around it. Pure.
 */

import type { PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  OFFICE_PICTURE_RECTANGLE_GEOMETRY,
  normalizeOfficePictureGeometry,
  officePictureGeometryPath,
} from "@mrnote/core/web/lib/office-picture-geometry.ts";
import { presentationShapeCatalogEntry, presentationShapeCatalogEntryForPreset } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

/** The preset a shape-gallery kind crops to, or null for a line / connector / unknown. */
export function pictureCropShapePreset(kind: ShapeKind): string | null {
  const entry = presentationShapeCatalogEntry(kind);
  if (!entry || entry.geometry !== "preset" || entry.category === "lines") return null;
  return normalizeOfficePictureGeometry(entry.drawingMlPreset);
}

/** The page's crop-to-shape preset, validated; null when there is none. */
export function normalizePictureClipShape(value: unknown): string | null {
  const preset = normalizeOfficePictureGeometry(value);
  if (!preset || preset === OFFICE_PICTURE_RECTANGLE_GEOMETRY) return null;
  const entry = presentationShapeCatalogEntryForPreset(preset);
  return entry && entry.geometry === "preset" && entry.category !== "lines" ? preset : null;
}

/** The gallery kind the picture is cropped to ("rectangle" when it is not). */
export function pictureClipShapeKind(block: PresentationImageBlock): ShapeKind {
  const preset = normalizePictureClipShape((block as Record<string, unknown>).clipShape);
  return (preset ? presentationShapeCatalogEntryForPreset(preset)?.kind : null) ?? "rectangle";
}

/**
 * The block cropped to `kind`'s outline (a rectangle removes the clip); null
 * when nothing changes. Rounded corners (`cornerRadius`) are the rounded
 * rectangle's own adjustment, so another shape replaces them.
 */
export function withPictureCropShape(
  block: PresentationImageBlock,
  kind: ShapeKind,
): PresentationImageBlock | null {
  const preset = kind === "rectangle" ? null : pictureCropShapePreset(kind);
  if (kind !== "rectangle" && !preset) return null;
  const current = normalizePictureClipShape((block as Record<string, unknown>).clipShape);
  if (current === preset && !(preset && (block as Record<string, unknown>).cornerRadius)) return null;
  const next = { ...block } as Record<string, unknown>;
  delete next.clipShape;
  // Dragged handles belong to the shape they were dragged on.
  delete next.clipShapeAdjustments;
  if (preset) {
    next.clipShape = preset;
    delete next.cornerRadius;
  }
  return next as PresentationImageBlock;
}

/**
 * The outline at the picture's size (EMU in, px out), for an SVG `clipPath`
 * in `objectBoundingBox` units scaled by `1 / width, 1 / height`; null when
 * the preset cannot be evaluated.
 */
export function pictureClipShapeOutline(
  preset: string,
  widthEmu: number,
  heightEmu: number,
  adjustments: unknown = null,
): { width: number; height: number; d: string } | null {
  if (!Number.isFinite(widthEmu) || !Number.isFinite(heightEmu) || widthEmu <= 0 || heightEmu <= 0) return null;
  const width = widthEmu / 9525;
  const height = heightEmu / 9525;
  
  // frames); none is the preset's own defaults.
  const d = officePictureGeometryPath(preset, adjustments, width, height);
  return d ? { width, height, d } : null;
}
