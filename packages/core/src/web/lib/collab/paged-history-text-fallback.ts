/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Transform } from "@tiptap/pm/transform";

/** Text blocks by their identity (`block_id`), wherever they stand. */
function textblocksById(doc: ProseMirrorNode): Map<string, ProseMirrorNode> {
  const found = new Map<string, ProseMirrorNode>();
  doc.descendants((node) => {
    if (!node.isTextblock) return true;
    const id = node.attrs.block_id;
    if (typeof id === "string" && id) found.set(id, node);
    return false;
  });
  return found;
}

/** The most text blocks one step's fallback restores. */
const MAX_RESTORED_BLOCKS = 64;

/**
 * The room document with the text of a writer's own step taken back, block
 * by block identity, or null when there is nothing it can take back safely.
 *
 * WUA 149: A types Q into a list item, B then changes that item's level. The
 * shared document has no move: the item and its paragraph are written again
 * as B's, so A's Undo step names text that no longer exists and undoes
 * nothing, while A's history is spent. The paragraph keeps its identity
 * through the move, so what A's step did to its words can still be undone:
 * each text block the step changed only in its content goes back to the
 * content it had before the step -- only where nobody has changed those
 * words since (its content is still what the step left). A block the step
 * created, removed or re-typed is left alone.
 */
export function restoreStepTextByIdentity(
  current: ProseMirrorNode,
  stepAfter: ProseMirrorNode,
  stepBefore: ProseMirrorNode,
): ProseMirrorNode | null {
  const before = textblocksById(stepBefore);
  const after = textblocksById(stepAfter);
  const wanted = new Map<string, ProseMirrorNode>();
  for (const [id, left] of after) {
    const previous = before.get(id);
    if (!previous || previous.type !== left.type || previous.content.eq(left.content)) continue;
    wanted.set(id, previous);
    if (wanted.size > MAX_RESTORED_BLOCKS) return null;
  }
  if (!wanted.size) return null;
  const transform = new Transform(current);
  const targets: Array<{ pos: number; node: ProseMirrorNode; restore: ProseMirrorNode }> = [];
  current.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const id = node.attrs.block_id;
    const restore = typeof id === "string" ? wanted.get(id) : undefined;
    // Only words still exactly as the step left them.
    if (restore && node.content.eq(after.get(id as string)!.content)) targets.push({ pos, node, restore });
    return false;
  });
  if (!targets.length) return null;
  for (const { pos, node, restore } of targets) {
    const from = transform.mapping.map(pos + 1);
    const to = transform.mapping.map(pos + 1 + node.content.size);
    transform.replaceWith(from, to, restore.content);
  }
  return transform.docChanged ? transform.doc : null;
}
