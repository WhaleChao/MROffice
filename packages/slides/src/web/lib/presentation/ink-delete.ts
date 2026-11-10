/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Review ▸ Hide Ink ▾'s two delete commands (PPT-A102): Delete All Ink on
 * Slide and Delete All Ink in Presentation. Ink is the persisted pen strokes
 * (`shape: "ink"` blocks the Draw tab and the slide-show pen keep); every
 * other object, including a shape converted from ink, stays. A locked stroke
 * is not removed, as no other command removes a locked object.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { persistedInkPathSignature } from "./ink-annotation";
import { canEditBlockContent } from "@mrnote/core/web/lib/presentation-text-case.ts";

export type InkDeleteScope = { kind: "slide"; index: number } | { kind: "presentation" };

function isDeletableInk(block: PresentationBlock): boolean {
  return persistedInkPathSignature(block) !== null && canEditBlockContent(block);
}

/** Whether `scope` holds any ink the command would delete. */
export function hasDeletableInk(
  slides: ReadonlyArray<{ blocks: readonly PresentationBlock[] }>,
  scope: InkDeleteScope,
): boolean {
  const pool = scope.kind === "slide" ? slides.slice(scope.index, scope.index + 1) : slides;
  return pool.some((slide) => slide.blocks.some(isDeletableInk));
}

/**
 * The slides with the scope's ink removed, each untouched slide kept by
 * reference; null when there is nothing to delete.
 */
export function slidesWithoutInk<S extends { blocks: PresentationBlock[] }>(
  slides: readonly S[],
  scope: InkDeleteScope,
): S[] | null {
  let changed = false;
  const next = slides.map((slide, index) => {
    if (scope.kind === "slide" && index !== scope.index) return slide;
    const blocks = slide.blocks.filter((block) => !isDeletableInk(block));
    if (blocks.length === slide.blocks.length) return slide;
    changed = true;
    return { ...slide, blocks };
  });
  return changed ? next : null;
}
