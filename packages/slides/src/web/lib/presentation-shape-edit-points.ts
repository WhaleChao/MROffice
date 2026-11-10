/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Edit Points on a slide shape: the outline the deck's Edit Points overlay
 * (PresentationWindow) edits and the AI's presentation.shape.edit_points
 * replaces, and the one write both commit.
 */

import type { PresentationShapeBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { normalizeDocumentShapeAdjustments } from "@mrnote/core/web/lib/document-shapes.ts";
import {
  officeCustomPathEditBox,
  officeCustomPathFromEdit,
  officeEditPathFromCustomPath,
  officeEditPathFromPreset,
  type OfficeEditPath,
} from "@mrnote/core/web/lib/office-custom-path-edit.ts";
import {
  normalizePresentationCustomPath,
  presentationCustomPathIsOpenLineOnly,
  presentationCustomPathToLegacyPoints,
} from "@mrnote/core/web/lib/presentation-freeform.ts";
import { blockClaimsNativeGroup, blockFlag, geometryNumber, isPresentationCameoBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { presentationShapeCatalogEntry } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { asRecord } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/** The outline Edit Points edits (the block's own path, or its preset's), in box pixels. */
export function presentationShapeEditPath(block: PresentationShapeBlock): OfficeEditPath | null {
  const width = (geometryNumber(block, "width") ?? 0) / 9525;
  const height = (geometryNumber(block, "height") ?? 0) / 9525;
  if (!(width > 0) || !(height > 0) || block.groupId || blockClaimsNativeGroup(block) || isPresentationCameoBlock(block) ||
    asRecord(asRecord(block.source)?.nativeSmartArtView)) return null;
  // A slide freeform carries no words in this model (the page schema and
  // both PPTX writers keep text off custom geometry), so a shape with text
  // keeps its preset rather than losing its label.
  const record = block as unknown as Record<string, unknown>;
  if ((typeof record.text === "string" && record.text.trim()) ||
    (Array.isArray(record.runs) && record.runs.length > 0) ||
    (Array.isArray(record.paragraphs) && record.paragraphs.length > 0)) return null;
  const customPath = normalizePresentationCustomPath(block.customPath);
  if (customPath) return officeEditPathFromCustomPath(customPath, width, height);
  if (Object.prototype.hasOwnProperty.call(block, "points")) return null;
  const entry = presentationShapeCatalogEntry(normalizeShapeKind(block.shape));
  return entry?.geometry === "preset"
    ? officeEditPathFromPreset(entry.drawingMlPreset, width, height, normalizeDocumentShapeAdjustments(block.adjustments))
    : null;
}

/**
 * One finished Edit Points edit: the block becomes (or stays) a freeform
 * whose box follows its outline, as Office writes the preset's current
 * outline as `a:custGeom` once Edit Points touches it. An outline that is one
 * open run of straight lines stays an ink line. Null when the path cannot be
 * stored (too many commands, nothing drawn).
 */
export function presentationShapeWithEditedPath(
  block: PresentationShapeBlock,
  path: OfficeEditPath,
): PresentationShapeBlock | null {
  const result = officeCustomPathFromEdit(path);
  if (!result) return null;
  const box = officeCustomPathEditBox(
    {
      x: (geometryNumber(block, "left") ?? 0) / 9525,
      y: (geometryNumber(block, "top") ?? 0) / 9525,
      width: (geometryNumber(block, "width") ?? 0) / 9525,
      height: (geometryNumber(block, "height") ?? 0) / 9525,
    },
    result,
    geometryNumber(block, "rotation") ?? 0,
    blockFlag(block, "flipH"),
    blockFlag(block, "flipV"),
  );
  const openLineOnly = presentationCustomPathIsOpenLineOnly(result.customPath);
  // An open line (an "ink" block) carries no fill, so a path opened and
  // closed again came back as a custom geometry with no fill of its own and
  
  // PowerPoint's Close Path leaves an unfilled freeform unfilled: a closed
  // shape made from a line states No Fill until the author picks one.
  const wasLine = normalizeShapeKind(block.shape) === "ink";
  const fillPatch =
    !openLineOnly && wasLine && block.fill === undefined ? { fill: "none" } : {};
  return shapeBlockWithPatch(block, {
    shape: openLineOnly ? "ink" : "custom-geometry",
    ...fillPatch,
    customPath: result.customPath,
    points: openLineOnly ? (presentationCustomPathToLegacyPoints(result.customPath) ?? undefined) : undefined,
    adjustments: undefined,
    left: Math.round(box.x * 9525),
    top: Math.round(box.y * 9525),
    width: Math.max(1, Math.round(box.width * 9525)),
    height: Math.max(1, Math.round(box.height * 9525)),
  } as Partial<PresentationShapeBlock>);
}
