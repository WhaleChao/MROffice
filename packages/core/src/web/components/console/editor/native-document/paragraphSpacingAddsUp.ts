/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { plainTextEdits } from "./textblockOnlyEdits";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

import { floatFlowMode, normalizeFloatAnchor } from "../../../../lib/office-float-anchor";

/**
 * Word's "Don't use HTML paragraph auto spacing" (`w:compat/
 * w:doNotUseHTMLParagraphAutoSpacing`, the page's `wordSpacingAddsUp`): a
 * paragraph's space after and the next one's space before ADD UP (Word
 * 16.86, measured on R27-D13's styles: 10 pt and 10 pt print 34.2 pt apart
 * with the switch, 24.15 pt without). Block margins collapse, so the second
 * paragraph of each pair on one page draws the sum as its own space before;
 * `GAP_ADDS_AFTER_ATTR` says how much of it is the first one's space after,
 * for the page ledger, which reads each paragraph's own spacing and adds the
 * two itself (`ledgerGap`). A paragraph opening its page has no pair: Word
 * starts a page with its own space before, as without the switch.
 */
export const GAP_ADDS_AFTER_ATTR = "data-gap-adds-after-pt";

const paragraphSpacingAddsUpKey = new PluginKey<DecorationSet>("paragraphSpacingAddsUp");

const PAIR_CONTAINERS = new Set(["doc", "pageBlock", "tableCell", "tableHeader"]);

function points(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 1584 ? value : null;
}

function spacingLine(node: ProseMirrorNode): boolean {
  return node.isTextblock && (node.type.name === "paragraph" || node.type.name === "heading");
}

/** A table Word floats over the text (`w:tblpPr`): it stands between two
 * paragraphs without taking a line of their flow. */
function floatingTable(node: ProseMirrorNode): boolean {
  return node.type.name === "table" && node.attrs?.tablePositioning != null;
}

/** An object drawn out of the flow (a floating shape or text box, not a Top
 * and Bottom band): it stands between two paragraphs in the page's blocks and
 * between none in Word's lines (corpus R280-D12: rules anchored between two
 * paragraphs of a form, whose 15.75 pt and 23.7 pt add up in Word). */
function anchoredObject(node: ProseMirrorNode): boolean {
  if (node.isTextblock || node.attrs?.wrapMode === "inline") return false;
  // A node never changes, and the document is walked on every change: a page
  // of ink read every stroke's anchor per pen landing.
  let anchored = anchoredObjects.get(node);
  if (anchored === undefined) {
    const anchor = normalizeFloatAnchor(node.attrs?.floatAnchor);
    anchored = anchor !== null && floatFlowMode(anchor) !== "band";
    anchoredObjects.set(node, anchored);
  }
  return anchored;
}

const anchoredObjects = new WeakMap<ProseMirrorNode, boolean>();

export function paragraphSpacingAddsUpDecorations(doc: ProseMirrorNode): DecorationSet {
  const everywhere = doc.attrs?.wordSpacingAddsUp === true;
  const decorations: Decoration[] = [];
  const visit = (container: ProseMirrorNode, start: number) => {
    let previous: ProseMirrorNode | null = null;
    // A floating table between two paragraphs keeps their spacing apart:
    // Word sets the second paragraph's space before below the first one's
    // space after, where the table stands (Word 16.86, corpus R83-D15 H9: a
    // 10pt after, the table, a 15pt before -- the text starts 15pt below the
    // table's top, 25pt below the first paragraph; collapsed, 5pt).
    let acrossFloat = false;
    container.forEach((child, offset) => {
      const position = start + offset;
      if (anchoredObject(child)) return;
      if (floatingTable(child) && previous && spacingLine(previous)) {
        acrossFloat = true;
        visit(child, position + 1);
        return;
      }
      const pairs = everywhere || acrossFloat;
      acrossFloat = false;
      if (
        pairs &&
        previous &&
        spacingLine(previous) &&
        spacingLine(child) &&
        previous.attrs?.contextualSpacing !== true &&
        child.attrs?.contextualSpacing !== true
      ) {
        const after = points(previous.attrs?.spaceAfter);
        const before = points(child.attrs?.spaceBefore);
        if (after !== null && before !== null && after > 0) {
          decorations.push(
            Decoration.node(position, position + child.nodeSize, {
              style: `margin-top: ${before + after}pt`,
              [GAP_ADDS_AFTER_ATTR]: String(after),
            }),
          );
        }
      }
      if (!child.isTextblock && !child.isLeaf && (PAIR_CONTAINERS.has(child.type.name) || child.type.name.startsWith("table"))) {
        visit(child, position + 1);
      }
      previous = child;
    });
  };
  visit(doc, 0);
  return decorations.length ? DecorationSet.create(doc, decorations) : DecorationSet.empty;
}

export function createParagraphSpacingAddsUpPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: paragraphSpacingAddsUpKey,
    state: {
      init: (_config, state) => paragraphSpacingAddsUpDecorations(state.doc),
      apply: (transaction, previous, _old, state) => {
        if (transaction.getMeta("documentAttrs")) return paragraphSpacingAddsUpDecorations(state.doc);
        if (!transaction.docChanged) return previous;
        // Spacing comes from the blocks' attributes and order, not their text.
        return plainTextEdits(transaction)
          ? previous.map(transaction.mapping, transaction.doc)
          : paragraphSpacingAddsUpDecorations(state.doc);
      },
    },
    props: {
      decorations: (state) => paragraphSpacingAddsUpKey.getState(state) ?? DecorationSet.empty,
    },
  });
}
