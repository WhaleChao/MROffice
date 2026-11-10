/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The visitor's homepage edits, made in the page's own editor after sign-up:
 * one ordinary transaction that replaces each block they changed, inserts
 * each block they added and deletes each block they removed, by block id,
 * exactly where those blocks stand in the editor. The editor then saves it as
 * it saves any edit (its collaboration model, its version, the DOCX
 * patch-back), and the visitor can undo it as one step.
 *
 * `current` is the editor's own document as it would save it now; `target`
 * the same document with the visitor's edits placed (lib/classic-start/rebase).
 * Refused (nothing dispatched) when the two do not line up block for block
 * outside the edits, when a kept block moved, or when the edit changed the
 * document's own settings rather than its blocks.
 */

import type { Node as ProseMirrorNode, Schema } from "@tiptap/pm/model";
import type { EditorState, Transaction } from "@tiptap/pm/state";

type Json = Record<string, unknown>;

export type ClassicEditorApply =
  | {
      readonly ok: true;
      readonly transaction: Transaction | null;
      /** The first operation the transaction makes (its first steps). */
      readonly first?: { readonly kind: "replace" | "delete" | "insert"; readonly blockId: string };
    }
  | { readonly ok: false; readonly reason: string };

/** Root attributes a save restamps from the layout; never the visitor's edit. */
const LAYOUT_ROOT_KEYS = new Set([
  "pageBlockSizes",
  "pageBlockLayoutKey",
  "pageRegionLayoutReceipt",
  "paginationProjection",
]);

function blockId(node: unknown): string | null {
  const attrs = (node as { attrs?: Json } | null)?.attrs;
  const id = attrs?.block_id;
  return typeof id === "string" && id ? id : null;
}

function stable(value: unknown): string {
  return JSON.stringify(value);
}

interface RuntimeRange {
  from: number;
  to: number;
}

/** Where each top-level block stands in the editor (inside its page sheet). */
function runtimeRanges(doc: ProseMirrorNode, pageType: string): Map<string, RuntimeRange> {
  const ranges = new Map<string, RuntimeRange>();
  const visit = (node: ProseMirrorNode, pos: number) => {
    const id = blockId(node.toJSON());
    if (!id) return;
    const known = ranges.get(id);
    // a block split across two pages is one block: its parts' whole span
    if (known) known.to = pos + node.nodeSize;
    else ranges.set(id, { from: pos, to: pos + node.nodeSize });
  };
  doc.forEach((child, offset) => {
    if (child.type.name === pageType) {
      child.forEach((inner, innerOffset) => visit(inner, offset + 1 + innerOffset));
    } else {
      visit(child, offset);
    }
  });
  return ranges;
}

/** Replace only what differs between `old` (at `pos`) and `next`, in place:
 *  down to the textblocks that changed when the structure is the same (a
 *  table's cells, a list's items), so each change is the one typing makes. */
function replaceInside(tr: Transaction, pos: number, old: ProseMirrorNode, next: ProseMirrorNode): void {
  const sameShape =
    !old.isTextblock &&
    old.childCount === next.childCount &&
    old.childCount > 0 &&
    Array.from({ length: old.childCount }, (_, i) => old.child(i).type === next.child(i).type).every(Boolean);
  if (sameShape) {
    // the later children first: the earlier positions stay where they were
    const offsets: number[] = [];
    old.forEach((_child, offset) => offsets.push(offset));
    for (let i = old.childCount - 1; i >= 0; i -= 1) {
      const before = old.child(i);
      const after = next.child(i);
      if (!before.eq(after)) replaceInside(tr, pos + 1 + offsets[i]!, before, after);
    }
    if (!old.hasMarkup(next.type, next.attrs, next.marks)) {
      tr.setNodeMarkup(tr.mapping.map(pos), next.type, next.attrs, next.marks);
    }
    return;
  }
  const start = old.content.findDiffStart(next.content);
  if (start !== null) {
    const end = old.content.findDiffEnd(next.content);
    let endA = end ? end.a : old.content.size;
    let endB = end ? end.b : next.content.size;
    // the two ends may overlap where text repeats: widen to a valid range
    const overlap = start - Math.min(endA, endB);
    if (overlap > 0) {
      endA += overlap;
      endB += overlap;
    }
    tr.replace(pos + 1 + start, pos + 1 + endA, next.slice(start, endB));
  }
  if (!old.hasMarkup(next.type, next.attrs, next.marks)) {
    tr.setNodeMarkup(tr.mapping.map(pos), next.type, next.attrs, next.marks);
  }
}

export function classicEditTransaction(input: {
  readonly state: EditorState;
  readonly schema: Schema;
  readonly current: Json;
  readonly target: Json;
  /** Storage blocks as the editor holds them (its hydration, for one block). */
  readonly toRuntimeBlock: (block: Json, rootAttrs: Json | undefined) => Json;
  readonly pageType?: string;
}): ClassicEditorApply {
  const { state, schema, current, target, toRuntimeBlock } = input;
  const pageType = input.pageType ?? "pageBlock";
  const currentBlocks = Array.isArray(current.content) ? (current.content as Json[]) : null;
  const targetBlocks = Array.isArray(target.content) ? (target.content as Json[]) : null;
  if (!currentBlocks || !targetBlocks) return { ok: false, reason: "not_a_document" };

  // the document's own settings: only the layout's restamps may differ
  const currentAttrs = (current.attrs as Json | undefined) ?? {};
  const targetAttrs = (target.attrs as Json | undefined) ?? {};
  for (const key of new Set([...Object.keys(currentAttrs), ...Object.keys(targetAttrs)])) {
    if (LAYOUT_ROOT_KEYS.has(key)) continue;
    if (stable(currentAttrs[key]) !== stable(targetAttrs[key])) {
      return { ok: false, reason: `root_attr:${key}` };
    }
  }

  const currentById = new Map<string, Json>();
  for (const block of currentBlocks) {
    const id = blockId(block);
    if (!id || currentById.has(id)) return { ok: false, reason: "current_block_identity" };
    currentById.set(id, block);
  }
  const targetIds = targetBlocks.map(blockId);
  if (targetIds.some((id) => !id) || new Set(targetIds).size !== targetIds.length) {
    return { ok: false, reason: "target_block_identity" };
  }
  // kept blocks keep their order
  const keptInTarget = (targetIds as string[]).filter((id) => currentById.has(id));
  const keptInCurrent = currentBlocks.map((block) => blockId(block) as string).filter((id) => keptInTarget.includes(id));
  if (stable(keptInTarget) !== stable(keptInCurrent)) return { ok: false, reason: "kept_blocks_moved" };

  const ranges = runtimeRanges(state.doc, pageType);
  for (const id of currentById.keys()) {
    if (!ranges.has(id)) return { ok: false, reason: "block_not_in_editor" };
  }

  type Op =
    | { kind: "replace"; from: number; to: number; order: number; block: Json; id: string }
    | { kind: "delete"; from: number; to: number; order: number; id: string }
    | { kind: "insert"; at: number; order: number; block: Json; id: string };
  const ops: Op[] = [];
  // insertion point: before the next kept block, or at the end of the last page
  const lastPage = state.doc.lastChild;
  const endOfContent =
    lastPage && lastPage.type.name === pageType ? state.doc.content.size - 1 : state.doc.content.size;
  let anchor = endOfContent;
  for (let index = targetBlocks.length - 1; index >= 0; index -= 1) {
    const block = targetBlocks[index]!;
    const id = targetIds[index] as string;
    const before = currentById.get(id);
    if (before) {
      const range = ranges.get(id)!;
      if (stable(before) !== stable(block)) {
        ops.push({ kind: "replace", from: range.from, to: range.to, order: index, block, id });
      }
      anchor = range.from;
    } else {
      ops.push({ kind: "insert", at: anchor, order: index, block, id });
    }
  }
  const targetIdSet = new Set(targetIds as string[]);
  for (const [id] of currentById) {
    if (targetIdSet.has(id)) continue;
    const range = ranges.get(id)!;
    ops.push({ kind: "delete", from: range.from, to: range.to, order: -1, id });
  }
  if (ops.length === 0) return { ok: true, transaction: null };

  const position = (op: Op) => (op.kind === "insert" ? op.at : op.from);
  // from the end of the document backwards, so earlier positions stay valid;
  // at one position the block there changes first, then what goes before it
  ops.sort((a, b) => {
    const delta = position(b) - position(a);
    if (delta !== 0) return delta;
    if (a.kind !== "insert" && b.kind === "insert") return -1;
    if (a.kind === "insert" && b.kind !== "insert") return 1;
    return b.order - a.order;
  });
  const tr = state.tr;
  try {
    for (const op of ops) {
      if (op.kind === "delete") {
        tr.delete(op.from, op.to);
        continue;
      }
      const node = schema.nodeFromJSON(toRuntimeBlock(op.block, targetAttrs));
      if (op.kind === "insert") {
        tr.insert(op.at, node);
        continue;
      }
      const old = tr.doc.nodeAt(op.from);
      if (old && op.from + old.nodeSize === op.to && old.type === node.type) {
        // as typing does it: only the changed stretch inside the block, and
        // the block's own attributes only if they changed
        replaceInside(tr, op.from, old, node);
      } else {
        tr.replaceWith(op.from, op.to, node);
      }
    }
  } catch (error) {
    return { ok: false, reason: `transaction:${error instanceof Error ? error.message.slice(0, 80) : "error"}` };
  }
  const first = ops[0]!;
  return { ok: true, transaction: tr, first: { kind: first.kind, blockId: first.id } };
}
