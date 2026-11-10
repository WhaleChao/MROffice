/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Reorder Objects and Reorder Overlapping Objects (Home ▸
 * Arrange, Shape Format ▸ Reorder Objects ▾; PPT-A045 / PPT-A119).
 *
 * PowerPoint for Mac lays the slide's objects out as layers, front first, and
 * lets the user drag a layer to a new depth; Apply writes the new stacking
 * order, Cancel leaves it. "Overlapping" limits the layers to the selected
 * objects and the objects that overlap them, so the other objects keep their
 * depth. Stacking order is the order the slide paints its objects in -- the
 * blocks' array order (presentation-z-order.ts) -- so a new order is a
 * permutation of the slide's objects: a grouped or multi-block object moves
 * whole, as it does in the Selection Pane.
 *
 * Pure: the dialog edits a list of layers; `applyReorderedLayers` turns it
 * into the slide's next blocks, refusing anything that would move a locked
 * object or name objects that are not the ones listed.
 */

import { rotatedBlockBounds } from "@mrnote/core/web/lib/presentation-multi-transform.ts";
import { presentationRestackedPlaces } from "@mrnote/core/web/lib/presentation-paint-order.ts";
import {
  isBlockHidden,
  selectionPaneObjectGroups,
  type SelectablePaneBlock,
} from "@mrnote/core/web/lib/presentation-selection-pane.ts";

export interface ReorderLayer<T extends SelectablePaneBlock> {
  /** The Selection Pane identity of the object (stable across reorders). */
  readonly objectId: string;
  readonly blocks: readonly T[];
  readonly blockIds: readonly string[];
  readonly label: string;
  readonly hidden: boolean;
  /** A locked object keeps its depth; the dialog cannot move it. */
  readonly locked: boolean;
  /** One of the objects that were selected when the dialog opened. */
  readonly selected: boolean;
}

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

function blockBox(block: SelectablePaneBlock): Box | null {
  const record = block as Record<string, unknown>;
  const read = (key: string) =>
    typeof record[key] === "number" && Number.isFinite(record[key]) ? (record[key] as number) : null;
  const left = read("left");
  const top = read("top");
  const width = read("width");
  const height = read("height");
  if (left === null || top === null || width === null || height === null) return null;
  return rotatedBlockBounds({ left, top, width, height }, read("rotation"));
}

function boxesOverlap(a: Box, b: Box): boolean {
  return (
    a.left < b.left + b.width &&
    b.left < a.left + a.width &&
    a.top < b.top + b.height &&
    b.top < a.top + a.height
  );
}

/**
 * The layers the dialog lists, FRONT first. With `overlappingWith`, only the
 * objects holding one of those block ids and the objects whose bounds overlap
 * theirs; null when that leaves fewer than two layers (nothing to reorder) or
 * the slide's objects cannot be addressed (the Selection Pane's own rule).
 */
export function reorderLayers<T extends SelectablePaneBlock>(
  blocks: readonly T[],
  options: {
    readonly nameOf: (block: T) => string;
    readonly isLocked: (block: T) => boolean;
    readonly selectedIds: ReadonlySet<string>;
    readonly overlappingOnly?: boolean;
  },
): ReorderLayer<T>[] | null {
  const groups = selectionPaneObjectGroups(blocks);
  if (!groups) return null;
  let layers = groups.map((group) => ({
    objectId: group.objectId,
    blocks: group.blocks,
    blockIds: group.blocks.map((block) => block.id),
    label: options.nameOf(group.blocks[0]),
    hidden: group.blocks.every(isBlockHidden),
    locked: group.blocks.some(options.isLocked),
    selected: group.blocks.some((block) => options.selectedIds.has(block.id)),
  }));
  if (options.overlappingOnly) {
    const selectedBoxes = layers
      .filter((layer) => layer.selected)
      .flatMap((layer) => layer.blocks.map(blockBox))
      .filter((box): box is Box => box !== null);
    if (selectedBoxes.length === 0) return null;
    layers = layers.filter(
      (layer) =>
        layer.selected ||
        layer.blocks.some((block) => {
          const box = blockBox(block);
          return box !== null && selectedBoxes.some((selected) => boxesOverlap(box, selected));
        }),
    );
  }
  if (layers.length < 2) return null;
  return layers.reverse();
}

/**
 * The slide's blocks with `layers` (front first, a reordering of the layers
 * `reorderLayers` listed) restacked: the listed objects take the depths they
 * held between them in the new order, every other object keeps its own. Null
 * when the order is unchanged, names other objects, or moves a locked one.
 */
export function applyReorderedLayers<T extends SelectablePaneBlock>(
  blocks: readonly T[],
  original: readonly ReorderLayer<T>[],
  reordered: readonly string[],
): T[] | null {
  const groups = selectionPaneObjectGroups(blocks);
  if (!groups) return null;
  const listed = new Set(original.map((layer) => layer.objectId));
  if (reordered.length !== listed.size || reordered.some((id) => !listed.has(id))) {
    return null;
  }
  if (new Set(reordered).size !== reordered.length) return null;
  if (reordered.every((id, index) => id === original[index]?.objectId)) return null;
  for (const [index, layer] of original.entries()) {
    if (layer.locked && reordered[index] !== layer.objectId) return null;
  }
  const byId = new Map(groups.map((group) => [group.objectId, group]));
  // The slide changed under the dialog (an object deleted or regrouped): the
  // listed objects no longer all exist as listed, so nothing is restacked.
  if (reordered.some((id) => !byId.has(id))) return null;
  for (const layer of original) {
    const group = byId.get(layer.objectId)!;
    if (group.blocks.length !== layer.blockIds.length ||
        group.blocks.some((block, index) => block.id !== layer.blockIds[index])) {
      return null;
    }
  }
  // The listed objects' slots back-to-front, filled from the new order read
  // back-to-front (the dialog lists front first).
  const backToFront = [...reordered].reverse();
  let next = 0;
  const restacked = groups.map((group) => {
    if (!listed.has(group.objectId)) return group;
    const replacement = byId.get(backToFront[next]!);
    next += 1;
    return replacement ?? group;
  });
  return presentationRestackedPlaces(restacked.flatMap((group) => group.blocks));
}

/** Move the layer `objectId` to `index` (front first) in a layer order. */
export function moveLayer(
  order: readonly string[],
  objectId: string,
  index: number,
): string[] {
  const from = order.indexOf(objectId);
  if (from < 0) return [...order];
  const target = Math.max(0, Math.min(order.length - 1, index));
  const next = [...order];
  next.splice(from, 1);
  next.splice(target, 0, objectId);
  return next;
}
