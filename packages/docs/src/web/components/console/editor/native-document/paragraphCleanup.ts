/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

/**
 * Paragraph cleanup (段落整理 / 删除空段) — the Word/WPS Writer "tidy up the
 * document" gesture exposed in the Review menu. It does two things in ONE
 * transaction over the whole doc:
 *
 *   (a) DELETE EMPTY PARAGRAPHS — a `paragraph` whose visible text content is
 *       empty or whitespace-only is removed, collapsing runs of blank lines.
 *   (b) TRIM TRAILING WHITESPACE — for every surviving paragraph, trailing
 *       spaces/tabs at the very end of the block are dropped.
 *
 * It is a pure EDITOR TRANSFORM: no schema change, no serialization change.
 * `planParagraphCleanup` is the small, testable core — it inspects the doc and
 * returns the exact set of edits to make (which paragraph ranges to delete +
 * which text tails to trim), or `null` when the document is already clean. When
 * it returns `null` the transform dispatches NO transaction, so a document with
 * no empty paragraphs (and no trailing whitespace) is left byte-for-byte
 * unchanged.
 *
 * SAFETY RULES (mirrored in the planner):
 *   - Only nodes whose type is `paragraph` are ever deleted. Headings, list
 *     items, table cells, blockquotes, code blocks, images (a block node here),
 *     page/section breaks, math blocks, etc. are never touched.
 *   - A paragraph that holds an INLINE ATOM (footnote / citation / inline-math —
 *     any non-text inline leaf) is NOT empty even with no characters, so it is
 *     never deleted. (Images are block-level in this schema and so are never a
 *     paragraph child, but the atom guard would protect an inline one too.)
 *   - The last remaining block of any parent is preserved, so a container (the
 *     doc, a table cell, …) is never emptied below the schema's `block+`
 *     requirement. A document that is entirely blank paragraphs collapses to a
 *     single empty paragraph rather than an invalid empty doc.
 *
 * DEFER (noted, not implemented here):
 *   - Merging soft line breaks (<br> / hardBreak) into spaces.
 *   - Full-width / half-width (全角/半角) punctuation & space normalization — so
 *     "blank" and the trailing trim deliberately consider only ASCII whitespace,
 *     leaving a lone full-width space (U+3000) untouched.
 */

/** ASCII whitespace trimmed from the tail of a surviving paragraph. */
const TRAILING_ASCII_WHITESPACE = /[ \t\f\v]+$/;
/** A paragraph is "blank" when its text is empty or only ASCII whitespace. */
const ASCII_WHITESPACE_ONLY = /^[ \t\r\n\f\v]*$/;

/** A paragraph range to delete (absolute open/close positions in the doc). */
export interface ParagraphDeletion {
  from: number;
  to: number;
}

/** A trailing-whitespace trim on one text leaf (absolute positions + new text). */
export interface TrailingTrim {
  from: number;
  to: number;
  text: string;
}

export interface ParagraphCleanupPlan {
  deletions: ParagraphDeletion[];
  trims: TrailingTrim[];
}

/** A node visited by {@link walkWithParent}, with its absolute geometry. */
interface VisitedNode {
  node: ProseMirrorNode;
  /** Absolute position where this node opens. */
  pos: number;
  /** Absolute position of the parent node (the doc is reported as -1). */
  parentPos: number;
  /** The parent's child count (lets the planner spot a would-be-empty parent). */
  parentChildCount: number;
}

/**
 * True when `node` is a `paragraph` that carries no meaningful content: every
 * child is a text node whose text is whitespace-only, and there is no inline
 * atom (footnote / citation / inline math / any non-text leaf). An empty
 * paragraph (no children at all) qualifies.
 */
export function isBlankParagraph(node: ProseMirrorNode): boolean {
  if (node.type.name !== "paragraph") return false;
  for (let index = 0; index < node.childCount; index += 1) {
    const child = node.child(index);
    if (!child.isText) {
      // An inline atom / non-text leaf (image, footnote, citation, inline math,
      // hard break, …) makes the paragraph non-empty — keep it.
      return false;
    }
    if (!ASCII_WHITESPACE_ONLY.test(child.text ?? "")) return false;
  }
  return true;
}

/**
 * Walk the document depth-first, collecting every node with its absolute
 * position, its parent's absolute position, and the parent's child count.
 * Mirrors ProseMirror's own descent order so positions line up with the live
 * doc. The doc node itself is not visited; its direct children are reported
 * with `parentPos === -1`.
 */
function walkWithParent(doc: ProseMirrorNode): VisitedNode[] {
  const visited: VisitedNode[] = [];
  const descend = (parent: ProseMirrorNode, parentPos: number): void => {
    let childPos = parentPos < 0 ? 0 : parentPos + 1;
    for (let index = 0; index < parent.childCount; index += 1) {
      const child = parent.child(index);
      visited.push({
        node: child,
        pos: childPos,
        parentPos,
        parentChildCount: parent.childCount,
      });
      if (child.childCount > 0) descend(child, childPos);
      childPos += child.nodeSize;
    }
  };
  descend(doc, -1);
  return visited;
}

/**
 * Inspect `doc` and produce the cleanup plan, or `null` when nothing would
 * change (no blank paragraphs to drop and no trailing whitespace to trim).
 *
 * Blank paragraphs are grouped by their parent; within each parent, all blank
 * paragraphs are scheduled for deletion EXCEPT when that would remove the
 * parent's every child — in that case the first blank in the parent is kept so
 * the container still satisfies the schema's `block+` content. Trailing-trim
 * edits are computed only for paragraphs that survive (a paragraph slated for
 * deletion needs no trim).
 *
 * Positions are absolute; deletions cover the paragraph's full node range
 * `[pos, pos + nodeSize)`. The caller applies edits back-to-front so each
 * earlier position stays valid.
 */
export function planParagraphCleanup(
  doc: ProseMirrorNode,
): ParagraphCleanupPlan | null {
  const nodes = walkWithParent(doc);

  // How many blank paragraphs each parent holds, keyed by parent position. A
  // parent whose blank count equals its child count is ALL blanks → keep one.
  const blanksPerParent = new Map<number, number>();
  for (const { node, parentPos } of nodes) {
    if (isBlankParagraph(node)) {
      blanksPerParent.set(parentPos, (blanksPerParent.get(parentPos) ?? 0) + 1);
    }
  }

  const deletions: ParagraphDeletion[] = [];
  const deletedFrom = new Set<number>();
  const keptOnePerParent = new Set<number>();
  for (const { node, pos, parentPos, parentChildCount } of nodes) {
    if (!isBlankParagraph(node)) continue;
    const blankChildren = blanksPerParent.get(parentPos) ?? 0;
    // If EVERY child of this parent is a blank paragraph, retain the first one
    // we encounter so the parent is never emptied below `block+`.
    if (blankChildren >= parentChildCount && !keptOnePerParent.has(parentPos)) {
      keptOnePerParent.add(parentPos);
      continue;
    }
    deletions.push({ from: pos, to: pos + node.nodeSize });
    deletedFrom.add(pos);
  }

  // Trailing-whitespace trims for surviving paragraphs. Trailing whitespace
  // lives at the block's tail, so only the LAST text child of each surviving
  // paragraph can need trimming.
  const trims: TrailingTrim[] = [];
  for (const { node, pos } of nodes) {
    if (node.type.name !== "paragraph") continue;
    if (deletedFrom.has(pos)) continue;
    const count = node.childCount;
    if (count === 0) continue;
    const last = node.child(count - 1);
    if (!last.isText) continue;
    const original = last.text ?? "";
    const trimmed = original.replace(TRAILING_ASCII_WHITESPACE, "");
    if (trimmed === original) continue;
    // Absolute start of the last child = paragraph content start (pos + 1) plus
    // the sizes of every earlier child.
    let offset = pos + 1;
    for (let index = 0; index < count - 1; index += 1) {
      offset += node.child(index).nodeSize;
    }
    trims.push({ from: offset, to: offset + last.nodeSize, text: trimmed });
  }

  if (deletions.length === 0 && trims.length === 0) return null;
  return { deletions, trims };
}

/**
 * Apply {@link planParagraphCleanup} to the editor in a SINGLE transaction.
 * Returns true when an edit was dispatched, false when the document was already
 * clean (no transaction → byte-for-byte unchanged).
 *
 * Deletions and trims are merged into one position-sorted list and applied
 * back-to-front (highest position first) so each remaining edit's absolute
 * positions stay valid as earlier content is removed/replaced.
 */
export function runParagraphCleanup(editor: Editor): boolean {
  const { state } = editor;
  const plan = planParagraphCleanup(state.doc);
  if (!plan) return false;

  type Edit =
    | { kind: "delete"; from: number; to: number }
    | { kind: "trim"; from: number; to: number; text: string };
  const edits: Edit[] = [
    ...plan.deletions.map((deletion): Edit => ({ kind: "delete", ...deletion })),
    ...plan.trims.map((trim): Edit => ({ kind: "trim", ...trim })),
  ];
  // Highest `from` first so a length change never shifts an earlier edit.
  edits.sort((a, b) => b.from - a.from);

  let transaction = state.tr;
  for (const edit of edits) {
    if (edit.kind === "delete") {
      transaction = transaction.delete(edit.from, edit.to);
    } else {
      // Only the trailing spaces go: the words before them keep their text
      // node, and with it their bold, font, size and colour. Replaced by a
      // new bare text node, the run lost every mark it had (corpus R395-D23
      // walk W35: a bold table-cell label came out plain, and its save was
      // refused as a changed run).
      transaction = transaction.delete(edit.from + edit.text.length, edit.to);
    }
  }

  if (!transaction.docChanged) return false;
  editor.view.dispatch(transaction.scrollIntoView());
  return true;
}
