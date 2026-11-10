/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";

type SlideSnapshot = { id: string; blocks: readonly unknown[]; widthEmu?: unknown; heightEmu?: unknown };
const row = (value: unknown): Record<string, unknown> | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown> : null;

/** Remote paint/text changes can coexist with an unfinished geometry gesture.
 * Geometry, membership and native source changes cannot: the pointer session
 * owns a captured coordinate frame and must not weaken its native authority.
 */
export function presentationGestureSurvivesRemote(
  before: readonly SlideSnapshot[],
  incoming: readonly SlideSnapshot[],
  slideId: string,
  blockIds: readonly string[],
  additionalFields: readonly string[] = [],
): boolean {
  const oldSlides = before.filter((slide) => slide.id === slideId);
  const newSlides = incoming.filter((slide) => slide.id === slideId);
  if (oldSlides.length !== 1 || newSlides.length !== 1 || blockIds.length === 0) return false;
  const oldSlide = oldSlides[0], newSlide = newSlides[0];
  if (oldSlide.widthEmu !== newSlide.widthEmu || oldSlide.heightEmu !== newSlide.heightEmu) return false;
  const oldRows = oldSlide.blocks.map(row), newRows = newSlide.blocks.map(row);
  if (oldRows.some((block) => !block) || newRows.some((block) => !block)) return false;
  const protectedFields = ["type", "source", "groupId", "groupPath", "left", "top", "width", "height", "rotation",
    "src", "assetId", "sourceKind", "hidden", "locked", ...additionalFields];
  for (const id of new Set(blockIds)) {
    const originals = oldRows.filter((block) => block?.id === id);
    const replacements = newRows.filter((block) => block?.id === id);
    if (originals.length !== 1 || replacements.length !== 1) return false;
    const original = originals[0]!, replacement = replacements[0]!;
    if (protectedFields.some((key) => !pageDocumentJsonEquals(original[key], replacement[key]))) return false;
    if (typeof original.groupId === "string") {
      const members = (blocks: typeof oldRows) => blocks.filter((block) => block?.groupId === original.groupId)
        .map((block) => block!.id).sort();
      if (!pageDocumentJsonEquals(members(oldRows), members(newRows))) return false;
    }
  }
  return true;
}
