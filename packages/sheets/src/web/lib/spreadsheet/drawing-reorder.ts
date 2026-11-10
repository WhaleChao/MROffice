/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Reorder Objects ▾ (Reorder Objects / Reorder Overlapping
 * Objects) on a worksheet's drawing stack, through the deck's own pure layer
 * rules (`presentation-reorder-objects`): a group is one layer, a locked
 * drawing keeps its depth, and the listed drawings swap the places they held
 * between them while every other drawing keeps its own.
 *
 * PowerPoint is the Office application with this dialog; Excel restacks one
 * step at a time (Bring Forward / Send Backward) or in the Selection Pane.
 * MRAI offers the dialog on a sheet too so the Shape Format tab is one tab.
 */

import {
  applyReorderedLayers,
  reorderLayers,
  type ReorderLayer,
} from "@mrnote/slides/web/lib/presentation-reorder-objects.ts";
import type { DrawingStackEntry } from "./drawing-stack";

/** One drawing as the dialog reads it, keyed `${kind}:${id}` as the stack is. */
export interface SpreadsheetReorderDrawing {
  readonly label: string;
  readonly hidden: boolean;
  readonly locked: boolean;
  readonly groupId?: string;
  readonly box: { readonly left: number; readonly top: number; readonly width: number; readonly height: number } | null;
}

export interface SpreadsheetReorderBlock {
  id: string;
  groupId?: string;
  hidden?: boolean;
  name: string;
  locked: boolean;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  [key: string]: unknown;
}

export type SpreadsheetDrawingReorderLayer = ReorderLayer<SpreadsheetReorderBlock>;

export function drawingStackKey(entry: DrawingStackEntry): string {
  return `${entry.kind}:${entry.id}`;
}

function stackBlocks(
  stack: readonly DrawingStackEntry[],
  drawings: ReadonlyMap<string, SpreadsheetReorderDrawing>,
): SpreadsheetReorderBlock[] {
  return stack.map((entry) => {
    const key = drawingStackKey(entry);
    const drawing = drawings.get(key);
    return {
      id: key,
      name: drawing?.label ?? entry.id,
      locked: drawing?.locked ?? true,
      ...(drawing?.hidden ? { hidden: true } : {}),
      ...(drawing?.groupId ? { groupId: drawing.groupId } : {}),
      ...(drawing?.box ?? {}),
    };
  });
}

/**
 * The sheet's drawings as the dialog's layers, front first; with
 * `overlappingOnly`, only the selected drawings and those overlapping them.
 * Null when fewer than two layers are left.
 */
export function spreadsheetDrawingReorderLayers(
  stack: readonly DrawingStackEntry[],
  drawings: ReadonlyMap<string, SpreadsheetReorderDrawing>,
  selectedKeys: ReadonlySet<string>,
  overlappingOnly = false,
): SpreadsheetDrawingReorderLayer[] | null {
  return reorderLayers(stackBlocks(stack, drawings), {
    nameOf: (block) => block.name,
    isLocked: (block) => block.locked,
    selectedIds: selectedKeys,
    overlappingOnly,
  });
}

/**
 * The stack after the dialog's Apply (`order`: the layers' object ids, front
 * first), or null when nothing moves, the sheet changed under the dialog, or
 * a locked drawing would move.
 */
export function spreadsheetDrawingStackReordered(
  stack: readonly DrawingStackEntry[],
  drawings: ReadonlyMap<string, SpreadsheetReorderDrawing>,
  layers: readonly SpreadsheetDrawingReorderLayer[],
  order: readonly string[],
): DrawingStackEntry[] | null {
  const blocks = applyReorderedLayers(stackBlocks(stack, drawings), layers, order);
  if (!blocks) return null;
  const byKey = new Map(stack.map((entry) => [drawingStackKey(entry), entry]));
  const next = blocks.map((block) => byKey.get(block.id));
  return next.every((entry): entry is DrawingStackEntry => entry !== undefined) ? next : null;
}

/** A layer's box (the union of its drawings), for the dialog's preview. */
export function spreadsheetReorderLayerBox(
  layer: SpreadsheetDrawingReorderLayer,
): { width: number; height: number } | null {
  const boxes = layer.blocks.filter(
    (block) => typeof block.width === "number" && typeof block.height === "number",
  );
  if (boxes.length === 0) return null;
  const left = Math.min(...boxes.map((block) => block.left ?? 0));
  const top = Math.min(...boxes.map((block) => block.top ?? 0));
  const right = Math.max(...boxes.map((block) => (block.left ?? 0) + (block.width ?? 0)));
  const bottom = Math.max(...boxes.map((block) => (block.top ?? 0) + (block.height ?? 0)));
  return { width: right - left, height: bottom - top };
}
