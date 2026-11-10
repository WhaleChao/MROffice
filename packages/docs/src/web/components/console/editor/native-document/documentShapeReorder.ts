/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shape Format > Reorder Objects ▾ (Reorder Objects / Reorder Overlapping
 * Objects) on the document's ONE floating stack: the layers the shared
 * reorder dialog (`PresentationReorderObjectsDialog`) lists, front first.
 *
 * PowerPoint is the Office application with this dialog; Word reorders the
 * same stack one step at a time (Bring Forward / Send Backward). MRAI offers
 * the dialog in Docs too so the Shape Format tab is one tab. Applying writes
 * through `reorderLiveDocumentStack` (shapeButtons), which swaps the listed
 * members' `relativeHeight`s and leaves every other floating object alone.
 */

import type { EditorView } from "@tiptap/pm/view";
import type { EditorState } from "@tiptap/pm/state";
import { NodeSelection } from "@tiptap/pm/state";
import { resolveDocumentShapeStack } from "@mrnote/core/web/lib/document-shapes.ts";
import { collectDocumentGoToTargets } from "./documentGoTo";
import { documentShapeMultiSelection } from "@mrnote/core/web/components/console/editor/native-document/documentShapeExtension.ts";

export interface DocumentShapeReorderLayer {
  /** The member's position, as text (the dialog's object id). */
  readonly objectId: string;
  readonly pos: number;
  readonly label: string;
  readonly selected: boolean;
  /** The member's drawn box on screen (the preview's proportions). */
  readonly box: { readonly width: number; readonly height: number } | null;
}

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

function overlaps(a: Box, b: Box): boolean {
  return a.left < b.left + b.width && b.left < a.left + a.width && a.top < b.top + b.height && b.top < a.top + a.height;
}

/** The positions of the floating objects selected now (one, or the Shift-clicked set). */
export function documentReorderSelectedPositions(state: EditorState): ReadonlySet<number> {
  const selected = new Set(documentShapeMultiSelection(state));
  if (state.selection instanceof NodeSelection) selected.add(state.selection.from);
  return selected;
}

/**
 * The floating stack's layers, front first; with `overlappingOnly`, only the
 * selected objects and those whose drawn boxes overlap them. Null when the
 * stack cannot be renumbered (an imported member the page does not show) or
 * fewer than two layers are left.
 */
export function documentShapeReorderLayers(
  view: EditorView,
  options: {
    readonly overlappingOnly?: boolean;
    readonly fallbackLabel: (type: string, ordinal: number) => string;
  },
): DocumentShapeReorderLayer[] | null {
  const { state } = view;
  const stack = resolveDocumentShapeStack(state.doc);
  if (!stack.complete || stack.backToFront.length < 2) return null;
  const selected = documentReorderSelectedPositions(state);
  const targets = new Map(
    collectDocumentGoToTargets(state.doc)
      .filter((target) => target.kind === "object" && target.selection === "node")
      .map((target) => [target.from, target]),
  );
  const boxOf = (pos: number): Box | null => {
    const dom = view.nodeDOM(pos);
    if (!(dom instanceof HTMLElement)) return null;
    const rect = dom.getBoundingClientRect();
    return rect.width > 0 || rect.height > 0
      ? { left: rect.left, top: rect.top, width: rect.width, height: rect.height }
      : null;
  };
  let layers = [...stack.backToFront].reverse().map((entry) => {
    const target = targets.get(entry.pos);
    const type = target?.id.split(":", 3)[1] || state.doc.nodeAt(entry.pos)?.type.name || "object";
    const raw = target?.label?.trim();
    const box = boxOf(entry.pos);
    return {
      objectId: String(entry.pos),
      pos: entry.pos,
      label: raw && raw !== type ? raw : options.fallbackLabel(type, target?.ordinal ?? 0),
      selected: selected.has(entry.pos),
      box,
    };
  });
  if (options.overlappingOnly) {
    const picked = layers.filter((layer) => layer.selected && layer.box).map((layer) => layer.box!);
    if (picked.length === 0) return null;
    layers = layers.filter(
      (layer) => layer.selected || (layer.box !== null && picked.some((box) => overlaps(box as Box, layer.box as Box))),
    );
  }
  return layers.length >= 2
    ? layers.map((layer) => ({
        ...layer,
        box: layer.box ? { width: layer.box.width, height: layer.box.height } : null,
      }))
    : null;
}
