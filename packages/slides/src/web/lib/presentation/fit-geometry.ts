/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";










export const FIT_GEOMETRY_KEYS = ["left", "top", "width", "height", "rowHeights"] as const;
export type FitGeometryKey = (typeof FIT_GEOMETRY_KEYS)[number];
export type FitGeometry = Partial<Record<FitGeometryKey, unknown>>;

function sameValue(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((value, index) => value === right[index]);
  }
  return false;
}

/**
 * `block` as it stands with only the fit's geometry put on it: every other
 * field (its rows, cells, words, runs) is the very value `block` holds, so a
 * fit measured from an older drawing of the block writes nothing but its frame.
 * The same block when the fit changes nothing.
 */
export function blockWithFitGeometry<B extends PresentationBlock>(block: B, fit: FitGeometry): B {
  let next: B | null = null;
  for (const key of FIT_GEOMETRY_KEYS) {
    if (!(key in fit)) continue;
    const value = fit[key];
    if (value === undefined || sameValue((block as Record<string, unknown>)[key], value)) continue;
    next ??= { ...block };
    (next as Record<string, unknown>)[key] = value;
  }
  return next ?? block;
}

/**
 * The live slide's blocks with each fit's geometry on the block it names (as
 * the block now stands, not as the fit measured it), or null when no block
 * changes. Every block it does not fit is the same object.
 */
export function blocksWithFitGeometry(
  blocks: readonly PresentationBlock[],
  fits: ReadonlyMap<string, FitGeometry>,
): PresentationBlock[] | null {
  let changed = false;
  const out = blocks.map((block) => {
    const fit = fits.get(block.id);
    if (!fit) return block;
    const next = blockWithFitGeometry(block, fit);
    if (next !== block) changed = true;
    return next;
  });
  return changed ? out : null;
}
