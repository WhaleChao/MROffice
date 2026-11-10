"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An empty paragraph that holds nothing but the mark of a CONTINUOUS section
 * break takes no room on Word's page: "Hello", an empty paragraph carrying
 * the section mark, then "World" in a continuous section set "World" exactly
 * one line pitch below "Hello" -- with or without a field's code in the empty
 * paragraph (Word 16.86 probe, `scratchpad/sectfld`; corpus R30-D04, whose
 * INDEX field begins in such a paragraph). The import keeps that paragraph
 * (it may hold a field's code, a bookmark), so the page draws it as zero
 * height instead of as a blank line.
 */

import { SheetFact } from "../../../../lib/editor/sheet-facts";
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

export const SECTION_MARK_ONLY_ATTR = "data-section-mark-only";
/**
 * An empty paragraph whose paragraph mark is hidden text (`w:pPr/w:rPr/w:vanish`,
 * read as the page's `paragraphMarkMarks` hiddenText) takes no room either,
 * in the body, in a cell, or between two tables: Word 16.86 set "B" one line
 * pitch below "A" across such a paragraph, and a table right under the one
 * before it (probe `scratchpad/hidmark`; corpus R03-D23, a job posting whose
 * HTML conversion puts one between its nested tables). Not when it is the
 * last block of its container, whose mark Word keeps, and not while hidden
 * text is shown (the stylesheet keys on `data-show-hidden-text`).
 */
export const HIDDEN_MARK_ONLY_ATTR = "data-hidden-mark-only";
/**
 * The same mark alone on the last line of a paragraph, after a line break
 * (`w:br` closing the text): that empty line takes no room either, and the
 * paragraph's space after is not added. Word 16.86 set the next section's
 * first line at the top of where the empty line would be, one line pitch
 * below the paragraph's last words; with the break removed, one line pitch
 * plus the paragraph's 12 pt after; with text after the break, both lines
 * and the 12 pt (corpus R93-D23 and edited copies of it, probe
 * `scratchpad/r93d23/v`; one or two columns in the next section alike).
 */
export const SECTION_MARK_AFTER_BREAK_ATTR = "data-section-mark-after-break";

function hiddenMarkOnly(node: ProseMirrorNode): boolean {
  if (node.type.name !== "paragraph" || node.content.size !== 0) return false;
  const marks = node.attrs?.paragraphMarkMarks;
  return Array.isArray(marks) && marks.some(
    (mark) => !!mark && typeof mark === "object" && (mark as { type?: unknown }).type === "hiddenText",
  );
}

/** The paragraph's text closes with a line break: a hard break node, or
 * the newline an imported `w:br` became at the end of its text. */
function endsInLineBreak(node: ProseMirrorNode): boolean {
  const last = node.lastChild;
  if (!last) return false;
  if (last.type.name === "hardBreak") return true;
  return Boolean(last.isText && last.text?.endsWith("\n"));
}

const sectionMarkParagraphKey = new PluginKey<DecorationSet>("documentSectionMarkParagraph");

/**
 * An empty paragraph the import folded a dedicated section-mark paragraph
 * INTO (`source.sectionMarkFolded`): the mark was the paragraph after it, and
 * this one is an ordinary blank line, which Word draws (Word 16.86, corpus
 * R281-D18: removing the dedicated paragraph moves nothing, removing this one
 * moves the next section up a line).
 */
function sectionMarkFolded(node: ProseMirrorNode): boolean {
  const source = node.attrs?.source as { sectionMarkFolded?: unknown } | null | undefined;
  return source?.sectionMarkFolded === 1;
}

function continuousBreak(node: ProseMirrorNode | null | undefined): boolean {
  return node?.type.name === "sectionBreak" && node.attrs.breakType === "continuous";
}

const HOLDS_HIDDEN_MARK = new SheetFact((node) => hiddenMarkOnly(node));
const HOLDS_CONTINUOUS_BREAK = new SheetFact((node) => continuousBreak(node));

export function buildSectionMarkParagraphDecorations(doc: ProseMirrorNode): DecorationSet {
  const decorations: Decoration[] = [];
  if (HOLDS_HIDDEN_MARK.inDoc(doc)) {
    const visitHidden = (parent: ProseMirrorNode, start: number) => {
      parent.forEach((child, offset, index) => {
        const pos = start + offset;
        if (child.isTextblock) {
          if (index + 1 < parent.childCount && hiddenMarkOnly(child)) {
            decorations.push(Decoration.node(pos, pos + child.nodeSize, { [HIDDEN_MARK_ONLY_ATTR]: "true" }));
          }
        } else if (!child.isAtom && child.childCount) {
          visitHidden(child, pos + 1);
        }
      });
    };
    visitHidden(doc, 0);
  }
  const visit = (parent: ProseMirrorNode, start: number) => {
    parent.forEach((child, offset, index) => {
      const pos = start + offset;
      if (child.isTextblock) {
        if (
          child.type.name === "paragraph" &&
          child.content.size === 0 &&
          !sectionMarkFolded(child) &&
          continuousBreak(index + 1 < parent.childCount ? parent.child(index + 1) : null)
        ) {
          decorations.push(Decoration.node(pos, pos + child.nodeSize, { [SECTION_MARK_ONLY_ATTR]: "true" }));
        } else if (
          child.type.name === "paragraph" &&
          endsInLineBreak(child) &&
          continuousBreak(index + 1 < parent.childCount ? parent.child(index + 1) : null)
        ) {
          decorations.push(Decoration.node(pos, pos + child.nodeSize, { [SECTION_MARK_AFTER_BREAK_ATTR]: "true" }));
          // An imported break is the text's own closing newline (white-space
          // pre-wrap draws the empty row after it): that newline is not drawn.
          if (child.lastChild?.isText) {
            const end = pos + child.nodeSize - 1;
            decorations.push(Decoration.inline(end - 1, end, { style: "display: none" }));
          }
        }
      } else if (!child.isAtom && child.childCount && child.type.name !== "table") {
        visit(child, pos + 1);
      }
    });
  };
  if (HOLDS_CONTINUOUS_BREAK.inDoc(doc)) visit(doc, 0);
  return decorations.length ? DecorationSet.create(doc, decorations) : DecorationSet.empty;
}

const DocumentSectionMarkParagraph = Extension.create({
  name: "documentSectionMarkParagraph",

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: sectionMarkParagraphKey,
        state: {
          init: (_config, state) => buildSectionMarkParagraphDecorations(state.doc),
          apply: (tr, value, _old, state) => (tr.docChanged ? buildSectionMarkParagraphDecorations(state.doc) : value),
        },
        props: {
          decorations: (state) => sectionMarkParagraphKey.getState(state) ?? DecorationSet.empty,
        },
      }),
    ];
  },
});

export default DocumentSectionMarkParagraph;
