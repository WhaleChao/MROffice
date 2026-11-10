/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which DrawingML preset's yellow adjustment handles a slide shape wears: the
 * deck's Adjust Shape handles (PresentationWindow) and the AI's
 * presentation.shape.adjust read the same answer.
 */

import type { PresentationShapeBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { normalizeDocumentShapeAdjustments } from "@mrnote/core/web/lib/document-shapes.ts";
import { officeShapeHasAdjustHandles } from "@mrnote/core/web/lib/office-shape-adjust-handles.ts";
import { normalizePresentationCustomPath } from "@mrnote/core/web/lib/presentation-freeform.ts";
import { blockClaimsNativeGroup, isPresentationCameoBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { presentationShapeCatalogEntry } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { asRecord } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/**
 * The DrawingML preset whose yellow handles this block wears, or null. An
 * imported block qualifies when its outline is the catalog's own (the
 * package's defaults) or already carries its handle values; one whose
 * outline the package drew some other way keeps its native drawing.
 */
export function presentationShapeAdjustPreset(block: PresentationShapeBlock): string | null {
  if (block.groupId || blockClaimsNativeGroup(block) || isPresentationCameoBlock(block) ||
    normalizePresentationCustomPath(block.customPath) || Object.prototype.hasOwnProperty.call(block, "points")) return null;
  const entry = presentationShapeCatalogEntry(normalizeShapeKind(block.shape));
  if (entry?.geometry !== "preset" || !officeShapeHasAdjustHandles(entry.drawingMlPreset)) return null;
  const nativeDrawing = asRecord(asRecord(block.source)?.nativeShapeDrawing);
  if (nativeDrawing && nativeDrawing.presetDefaults !== true && !normalizeDocumentShapeAdjustments(block.adjustments)) return null;
  if (asRecord(asRecord(block.source)?.nativeSmartArtView)) return null;
  return entry.drawingMlPreset;
}
