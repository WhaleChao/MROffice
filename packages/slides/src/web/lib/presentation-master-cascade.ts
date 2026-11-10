/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide-master placeholder cascade (PowerPoint "edit the master, it propagates"):
 * a slide INHERITS its master's placeholder blocks unless it provides its own
 * block with the same id, in which case the slide's block wins. Inherited master
 * blocks render BEHIND the slide's own blocks and are tagged so the editor draws
 * them read-only (you edit them on the master, not the slide).
 *
 * Pure + generic over the block shape so the cascade is exhaustively testable
 * without the full presentation model. An absent / empty master block list yields
 * the slide's own blocks UNCHANGED (byte-for-byte) — a deck whose master carries
 * no placeholder blocks is unaffected.
 */

/** A block tagged as inherited from the master (rendered read-only, behind own). */
export type InheritedBlock<B> = B & { inheritedFromMaster: true };
/** A block inherited specifically from the selected native child layout. */
export type LayoutInheritedBlock<B> = InheritedBlock<B> & {
  inheritedFromLayout: true;
};

function placeholderIdentity(block: unknown): string | undefined {
  if (!block || typeof block !== "object" || Array.isArray(block)) {
    return undefined;
  }
  const record = block as Record<string, unknown>;
  const placeholderType =
    typeof record.placeholderType === "string"
      ? record.placeholderType.trim()
      : typeof record.placeholder === "string"
        ? record.placeholder.trim()
        : "";
  if (!placeholderType) return undefined;
  const rawIndex = record.placeholderIndex;
  const index =
    typeof rawIndex === "number" && Number.isInteger(rawIndex)
      ? rawIndex
      : typeof rawIndex === "string" && /^\d+$/.test(rawIndex)
        ? Number(rawIndex)
        : 0;
  return `${placeholderType}:${index}`;
}

function blockKeys(block: { id?: string }): string[] {
  const keys: string[] = [];
  if (block.id) keys.push(`id:${block.id}`);
  const placeholder = placeholderIdentity(block);
  if (placeholder) keys.push(`placeholder:${placeholder}`);
  return keys;
}

function blockKeySet<B extends { id?: string }>(blocks: readonly B[]): Set<string> {
  return new Set(blocks.flatMap(blockKeys));
}

function isOverridden<B extends { id?: string }>(
  block: B,
  keys: ReadonlySet<string>,
): boolean {
  return blockKeys(block).some((key) => keys.has(key));
}

/**
 * Resolve the blocks to RENDER for a slide: the master's placeholder blocks the
 * slide does NOT override (matched by id), tagged `inheritedFromMaster`, placed
 * BEHIND the slide's own blocks (master first in the array = drawn underneath).
 * A slide block with the same id as a master placeholder REPLACES it (the slide's
 * content wins). Master blocks are de-duped by id. Returns the slide's own blocks
 * unchanged when the master has no blocks.
 */
export function resolveSlideRenderBlocks<B extends { id?: string }>(
  masterBlocks: readonly B[] | undefined,
  slideBlocks: readonly B[],
  layoutBlocks?: readonly B[],
): Array<B | InheritedBlock<B> | LayoutInheritedBlock<B>> {
  if (
    (!masterBlocks || masterBlocks.length === 0) &&
    (!layoutBlocks || layoutBlocks.length === 0)
  ) {
    return [...slideBlocks];
  }
  const ownKeys = blockKeySet(slideBlocks);
  const layoutKeys = blockKeySet(layoutBlocks ?? []);
  const inherited: Array<InheritedBlock<B>> = [];
  const seen = new Set<string>();
  for (const block of masterBlocks ?? []) {
    const keys = blockKeys(block);
    // Skip a master placeholder the slide overrides, and de-dupe master blocks by
    // stable id OR OOXML placeholder identity. A selected layout also overrides
    // the matching master placeholder before the slide-level override is applied.
    if (
      isOverridden(block, ownKeys) ||
      isOverridden(block, layoutKeys) ||
      keys.some((key) => seen.has(key))
    ) {
      continue;
    }
    keys.forEach((key) => seen.add(key));
    inherited.push({ ...block, inheritedFromMaster: true });
  }
  const inheritedLayouts: Array<LayoutInheritedBlock<B>> = [];
  const seenLayouts = new Set<string>();
  for (const block of layoutBlocks ?? []) {
    const keys = blockKeys(block);
    if (
      isOverridden(block, ownKeys) ||
      keys.some((key) => seenLayouts.has(key))
    ) {
      continue;
    }
    keys.forEach((key) => seenLayouts.add(key));
    inheritedLayouts.push({
      ...block,
      inheritedFromMaster: true,
      inheritedFromLayout: true,
    });
  }
  return [...inherited, ...inheritedLayouts, ...slideBlocks];
}

/** Whether a render block came from the master cascade (so it edits read-only). */
export function isInheritedBlock<B>(block: B): block is InheritedBlock<B> {
  return (
    typeof block === "object" &&
    block !== null &&
    (block as { inheritedFromMaster?: unknown }).inheritedFromMaster === true
  );
}

/** Whether an inherited render block came from the selected child layout. */
export function isLayoutInheritedBlock<B>(
  block: B,
): block is LayoutInheritedBlock<B> {
  return (
    isInheritedBlock(block) &&
    (block as LayoutInheritedBlock<B>).inheritedFromLayout === true
  );
}
