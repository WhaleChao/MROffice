/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** PowerPoint's first Duplicate offset: 12 pt right and down, in EMU. */
export const PRESENTATION_DUPLICATE_OFFSET_EMU = 152_400;

export interface PresentationDuplicateOffset {
  readonly dx: number;
  readonly dy: number;
}

interface PositionedBlock {
  readonly id: string;
  readonly left?: unknown;
  readonly top?: unknown;
}

interface DuplicateSlide {
  readonly id: string;
  readonly blocks: ReadonlyArray<PositionedBlock>;
  readonly animationSequence?: unknown;
}

/** What the last Duplicate made: its copies, where their originals stood,
 *  and the slide as it was apart from the copies' positions. */
export interface PresentationDuplicateMemory {
  readonly slideId: string;
  readonly cloneIds: readonly string[];
  /** Per copy: the original's position when it was duplicated. */
  readonly originById: Readonly<Record<string, { left: number; top: number }>>;
  readonly fingerprint: string;
}

const position = (block: PositionedBlock | undefined) =>
  block &&
  typeof block.left === "number" &&
  Number.isFinite(block.left) &&
  typeof block.top === "number" &&
  Number.isFinite(block.top)
    ? { left: block.left, top: block.top }
    : null;

/** The slide with the copies' `left`/`top` left out: a move of the copies
 *  keeps it, any other edit changes it. */
function slideFingerprint(slide: DuplicateSlide, cloneIds: ReadonlySet<string>): string {
  return JSON.stringify({
    blocks: slide.blocks.map((block) => {
      if (!cloneIds.has(block.id)) return block;
      const rest: Record<string, unknown> = { ...block };
      delete rest.left;
      delete rest.top;
      return rest;
    }),
    animationSequence: slide.animationSequence ?? null,
  });
}

/** Remember a Duplicate: `pairs` maps each original (as it stood) to its copy. */
export function rememberPresentationDuplicate(
  slideAfter: DuplicateSlide,
  pairs: ReadonlyArray<{ original: PositionedBlock; cloneId: string }>,
): PresentationDuplicateMemory | null {
  const originById: Record<string, { left: number; top: number }> = {};
  for (const { original, cloneId } of pairs) {
    const origin = position(original);
    if (!origin || !slideAfter.blocks.some((block) => block.id === cloneId)) return null;
    originById[cloneId] = origin;
  }
  const cloneIds = pairs.map((pair) => pair.cloneId);
  if (cloneIds.length === 0) return null;
  return {
    slideId: slideAfter.id,
    cloneIds,
    originById,
    fingerprint: slideFingerprint(slideAfter, new Set(cloneIds)),
  };
}

/**
 * The offset the next Duplicate of `selectedIds` takes: the remembered
 * original -> copy offset when the selection is exactly the last Duplicate's
 * copies and the slide changed only by moving them; else 12 pt right and down.
 */
export function presentationDuplicateOffset(
  memory: PresentationDuplicateMemory | null,
  slide: DuplicateSlide | undefined,
  selectedIds: ReadonlySet<string>,
): PresentationDuplicateOffset {
  const fallback = {
    dx: PRESENTATION_DUPLICATE_OFFSET_EMU,
    dy: PRESENTATION_DUPLICATE_OFFSET_EMU,
  };
  if (
    !memory ||
    !slide ||
    slide.id !== memory.slideId ||
    selectedIds.size !== memory.cloneIds.length ||
    memory.cloneIds.some((id) => !selectedIds.has(id)) ||
    slideFingerprint(slide, selectedIds) !== memory.fingerprint
  ) {
    return fallback;
  }
  const primaryId = memory.cloneIds[0];
  const now = position(slide.blocks.find((block) => block.id === primaryId));
  const origin = memory.originById[primaryId];
  if (!now || !origin) return fallback;
  return { dx: now.left - origin.left, dy: now.top - origin.top };
}
