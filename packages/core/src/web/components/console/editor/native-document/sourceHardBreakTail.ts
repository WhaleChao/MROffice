/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { SheetFact } from "../../../../lib/editor/sheet-facts";
import { plainTextEdits } from "./textblockOnlyEdits";
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import styles from "./sourceHardBreakTail.module.css";

type Segment = {
  base: string;
  index: number;
  count: number;
  breakBefore: "page" | "column" | null;
  nonPrintingEmptyTail: boolean;
  styleId: unknown;
  styleName: unknown;
};

function sourceSegment(node: ProseMirrorNode): Segment | null {
  if (node.type.name !== "paragraph" && node.type.name !== "heading") return null;
  const source = node.attrs.source;
  const marker = source?.hardBreakSegment;
  if (!source || source.adapter !== "docx" || source.nodeType !== "paragraph" ||
      !marker || marker.version !== 1 || marker.splitParagraphMark !== false ||
      !Number.isSafeInteger(marker.index) || !Number.isSafeInteger(marker.count) ||
      marker.index < 0 || marker.count < 2 || marker.index >= marker.count ||
      typeof marker.nonPrintingEmptyTail !== "boolean") return null;
  const match = typeof source.sourceAnchor === "string"
    ? /^(word\/document\.xml#p\[\d+\])(?:\/seg\[([1-9]\d*)\])?$/.exec(source.sourceAnchor) : null;
  if (!match || Number(match[2] ?? 0) !== marker.index ||
      (marker.index === 0 ? marker.breakBefore !== null : !["page", "column"].includes(marker.breakBefore))) return null;
  if (marker.nonPrintingEmptyTail && (marker.index !== marker.count - 1 || marker.breakBefore !== "page")) return null;
  return { base: match[1], index: marker.index, count: marker.count, breakBefore: marker.breakBefore,
    nonPrintingEmptyTail: marker.nonPrintingEmptyTail, styleId: source.styleId, styleName: source.styleName };
}

function paragraphMarkup(node: ProseMirrorNode): ProseMirrorNode {
  const attrs = { ...node.attrs };
  // The paragraph mark's own run formatting is the last piece's alone (it
  // holds the mark): a tail stating the mark's proofing language while its
  // first piece states none is still the same paragraph (corpus R73-D22:
  // every break paragraph drew its mark as a bordered blank line on the
  // next page, which Word does not draw).
  // The empty paragraph's size the mark gives (`paragraphFontSize`) is the
  // mark's too (corpus R280-D19: a tail stating 12 pt after a head stating
  // none drew its line at the top of page 2 and pushed the floating table
  // there 13.8 pt down).
  for (const key of ["source", "block_id", "anchor_id", "toc_anchor_id", "paragraphMarkMarks", "paragraphFontSize"]) delete attrs[key];
  return node.type.create(attrs, null, node.marks);
}

/** A block drawn from an object inside the source paragraph `base`. */
function objectOfSourceParagraph(node: ProseMirrorNode, base: string): boolean {
  if (node.type.name === "paragraph" || node.type.name === "heading") return false;
  const source = node.attrs.source;
  return Boolean(source && source.adapter === "docx" && typeof source.sourceAnchor === "string" &&
    source.sourceAnchor.startsWith(`${base}/`));
}

const LIST_NODES = new Set(["orderedList", "bulletList"]);

/** Current flow only: provenance is retained, never used as write authority.
 * A numbered paragraph's pieces stand in lists: each piece is read from its
 * list item, and a tail that is its list's one item and paragraph is the
 * list (corpus R158-D19: "Following submission" + page break, numbered;
 * Word draws neither the mark's line nor a number on the next page, where
 * the page drew a "9." line). */
export function sourceHardBreakTailPositions(doc: ProseMirrorNode): number[] {
  const flow: Array<{ node: ProseMirrorNode; pos: number; block: number | null }> = [];
  function collect(parent: ProseMirrorNode, start: number) {
    parent.forEach((node, offset) => {
      const pos = start + offset;
      if (node.type.name === "pageBlock") collect(node, pos + 1);
      else if (LIST_NODES.has(node.type.name)) {
        const lone = node.childCount === 1 && node.firstChild!.childCount === 1;
        node.forEach((item, itemOffset) => {
          const itemPos = pos + 1 + itemOffset;
          item.forEach((child, childOffset) => {
            flow.push({ node: child, pos: itemPos + 1 + childOffset, block: lone ? pos : null });
          });
        });
      } else flow.push({ node, pos, block: pos });
    });
  }
  collect(doc, 0);
  const result: number[] = [];
  for (let end = 0; end < flow.length; end++) {
    const tail = flow[end];
    const marker = sourceSegment(tail.node);
    if (!marker?.nonPrintingEmptyTail || tail.node.content.size !== 0 || marker.count > flow.length) continue;
    const start = end - 2 * (marker.count - 1);
    if (start < 0) continue;
    if (sourceSegment(flow[start].node)?.index !== 0) {
      // A first piece holding only an object Word anchors there (R87-D17: a
      // VML text box, then the page break): the page draws the object as
      // its own block and keeps no empty paragraph for the piece, so the
      // tail follows the object and its break directly. Word still draws
      // no line for the paragraph mark after the break (the chapter title
      // under it opens the page).
      if (marker.count === 2 && objectOfSourceParagraph(flow[start].node, marker.base) &&
          flow[end - 1].node.type.name === "pageBreak" && tail.block !== null) result.push(tail.block);
      continue;
    }
    const markup = paragraphMarkup(flow[start].node);
    let valid = true;
    for (let index = 0; index < marker.count; index++) {
      const candidate = flow[start + 2 * index].node;
      const part = sourceSegment(candidate);
      if (!part || part.base !== marker.base || part.count !== marker.count || part.index !== index ||
          part.styleId !== marker.styleId || part.styleName !== marker.styleName ||
          !paragraphMarkup(candidate).sameMarkup(markup)) { valid = false; break; }
      if (index > 0 && flow[start + 2 * index - 1].node.type.name !==
          (part.breakBefore === "page" ? "pageBreak" : "columnBreak")) { valid = false; break; }
    }
    if (valid && tail.block !== null) result.push(tail.block);
  }
  return result;
}

const HOLDS_SOURCE_SEGMENT = new SheetFact((node) => sourceSegment(node) !== null);

/**
 * The pieces of paragraphs a page break inside them split, as Word lays them
 * out (Word 16.86, corpus R04-D09: six paragraphs each opening with a page
 * break, 6 pages):
 * - `leads`: an empty first piece, the line holding only the leading page
 *   (or column) break. Word draws it at the foot of the page before, under the text,
 *   without the paragraph's space before, and it never opens a page of its
 *   own. Laid out as a line of its own with that space, it took a page.
 * - `continuations`: the text after a page break inside the paragraph. It
 *   goes on at the top of the next page without the paragraph's space
 *   before: that space is the paragraph's first line's.
 */
export function sourceHardBreakPiecePositions(doc: ProseMirrorNode): { leads: number[]; continuations: number[] } {
  if (!HOLDS_SOURCE_SEGMENT.inDoc(doc)) return { leads: [], continuations: [] };
  const flow: Array<{ node: ProseMirrorNode; pos: number }> = [];
  function collect(parent: ProseMirrorNode, start: number) {
    parent.forEach((node, offset) => {
      const pos = start + offset;
      if (node.type.name === "pageBlock") collect(node, pos + 1);
      else flow.push({ node, pos });
    });
  }
  collect(doc, 0);
  const leads: number[] = [];
  const continuations: number[] = [];
  flow.forEach(({ node, pos }, index) => {
    const marker = sourceSegment(node);
    if (!marker) return;
    const breakAfter = flow[index + 1]?.node.type.name;
    if (marker.index === 0 && node.content.size === 0 && (breakAfter === "pageBreak" || breakAfter === "columnBreak")) {
      const kind = breakAfter === "pageBreak" ? "page" : "column";
      // Only before text of its own paragraph: a paragraph holding nothing
      // but the break keeps its line (its tail is the non-printing one).
      const after = flow[index + 2]?.node;
      const next = after ? sourceSegment(after) : null;
      if (next && next.base === marker.base && next.index === 1 && next.breakBefore === kind &&
          !next.nonPrintingEmptyTail && after!.content.size > 0) leads.push(pos);
      // A line holding only the break keeps no space before either (Word
      // 16.86, corpus R73-D22: a paragraph of a page break alone, and one
      // of a column break before an empty piece, stand right under the line
      // above them in Word; with their 60 pt before they took a column).
      else continuations.push(pos);
    } else if (marker.index >= 1 && marker.breakBefore !== null && !marker.nonPrintingEmptyTail &&
        flow[index - 1]?.node.type.name === (marker.breakBefore === "page" ? "pageBreak" : "columnBreak")) {
      continuations.push(pos);
    }
  });
  return { leads, continuations };
}

type TailState = { positions: number[]; decorations: DecorationSet };
export const sourceHardBreakTailKey = new PluginKey<TailState>("sourceHardBreakTail");

function decorations(doc: ProseMirrorNode, positions: number[], caret: number | null): DecorationSet {
  const pieces = sourceHardBreakPiecePositions(doc);
  return DecorationSet.create(doc, [
    ...positions.map((pos) => {
      const node = doc.nodeAt(pos)!;
      // A tail standing as its own list: the list's number is not drawn either.
      const list = LIST_NODES.has(node.type.name) ? ` ${styles.nonPrintingTailList}` : "";
      return Decoration.node(pos, pos + node.nodeSize, {
        class: `${styles.nonPrintingTail}${list}${caret === pos + 1 ? ` ${styles.focusedCaret}` : ""}`,
        "data-mrnote-source-break-tail": "true",
      });
    }),
    // The empty line of a leading page break: no physical line of its own.
    ...pieces.leads.map((pos) => Decoration.node(pos, pos + doc.nodeAt(pos)!.nodeSize, {
      class: `${styles.nonPrintingTail}${caret === pos + 1 ? ` ${styles.focusedCaret}` : ""}`,
      "data-mrnote-source-break-lead": "true",
    })),
    ...pieces.continuations.map((pos) => Decoration.node(pos, pos + doc.nodeAt(pos)!.nodeSize, {
      class: styles.breakContinuation,
    })),
  ]);
}

export const SourceHardBreakTail = Extension.create({
  name: "sourceHardBreakTail",
  addProseMirrorPlugins() {
    return [new Plugin<TailState>({
      key: sourceHardBreakTailKey,
      state: {
        init: (_, state) => {
          const positions = sourceHardBreakTailPositions(state.doc);
          return { positions, decorations: decorations(state.doc, positions, state.selection.empty ? state.selection.from : null) };
        },
        // Equal-content document replacement still recreates its positions.
        apply: (tr, previous) => {
          if (!tr.docChanged && !tr.selectionSet) return previous;
          // Typing in a block that holds text moves the tails (empty blocks
          // and their source markers) without making or ending one.
          const positions = !tr.docChanged
            ? previous.positions
            : plainTextEdits(tr)
              ? previous.positions.map((position) => tr.mapping.map(position))
              : sourceHardBreakTailPositions(tr.doc);
          return { positions, decorations: decorations(tr.doc, positions, tr.selection.empty ? tr.selection.from : null) };
        },
      },
      props: { decorations: (state) => sourceHardBreakTailKey.getState(state)?.decorations ?? DecorationSet.empty },
    })];
  },
});
