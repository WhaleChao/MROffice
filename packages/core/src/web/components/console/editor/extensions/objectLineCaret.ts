"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The insertion point on a line that holds an in-line picture or chart.
 *
 * Word draws the insertion point as tall as the text it stands in; only right
 * beside an In Line with Text object is it as tall as that object. WebKit
 * draws its native caret as tall as the whole LINE box instead, so a line an
 * in-line picture or chart makes 250 px high got a 250 px caret wherever the
 * caret stood on it -- beside "我是" as much as beside the picture.
 *
 * While the caret stands in a paragraph that holds an in-line drawing, the
 * native caret is made transparent and this plugin draws one from the
 * caret's own client rect: the text's box beside text, the object's box
 * beside the object (Chromium's and Word's reading). It follows focus,
 * scrolling and layout, blinks like a caret and steps aside during IME
 * composition and while the caret is parked beside a block object (that
 * caret is drawn by `objectPlacement.ts`).
 */

import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, TextSelection, type EditorState } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

import { documentBoundaryInteractionKey } from "../native-document/documentBoundaryInteraction";
import { isInlineDrawingNode } from "../native-document/documentObjectNodes";
import { isImeCompositionActive } from "../native-document/imeCompositionGuard";
import { parkedBlockObjectSide } from "./objectPlacement";
import { clipToVisibleEditor } from "../native-document/visibleEditorClip";

/**
 * On the paragraph whose caret the drawn one replaces. The stylesheet makes
 * that paragraph's native caret transparent while the editor has focus and
 * no composition runs (the root's `data-object-line-caret-composing`). It
 * used to be a class on the editor root, and caret-color is inherited: every
 * paint (each update and scroll with the caret in such a paragraph) took
 * the class off to read the caret colour and put it back, two restyles of
 * the whole document, each 300 ms and more in WebKit on a long document.
 */
export const OBJECT_LINE_CARET_ATTR = "data-object-line-caret";
const COMPOSING_ATTR = "data-object-line-caret-composing";

/**
 * An in-line object the caret can stand beside: a drawing, or an in-line
 * equation. The browser draws no caret beside a non-editable equation at the
 * start of its line (nothing to measure there), and one as tall as the line
 * box after it; Word draws it at the equation's edge, as tall as the equation.
 */
function isCaretObjectNode(node: { readonly type: { readonly name: string } }): boolean {
  return isInlineDrawingNode(node) || node.type.name === "inlineMath";
}

function paragraphHoldsDrawing(state: EditorState): boolean {
  const { selection } = state;
  if (!(selection instanceof TextSelection) || !selection.empty) return false;
  const parent = selection.$head.parent;
  if (!parent.isTextblock) return false;
  let holds = false;
  parent.forEach((child) => {
    if (!holds && isCaretObjectNode(child)) holds = true;
  });
  return holds;
}

/** Whether the document's state lets the drawn caret stand in for the
 * native one; focus and composition are the view's part. */
function drawnCaretState(state: EditorState): boolean {
  return (
    paragraphHoldsDrawing(state) &&
    // No caret at all while the boundary policy keeps it hidden (a
    // pointer dismissal, a caret parked beside a block object).
    documentBoundaryInteractionKey.getState(state) !== "idle" &&
    !parkedBlockObjectSide(state)
  );
}

function drawnCaretDecorations(state: EditorState): DecorationSet {
  if (!drawnCaretState(state)) return DecorationSet.empty;
  const { $head } = state.selection;
  const from = $head.before();
  return DecorationSet.create(state.doc, [
    Decoration.node(from, from + $head.parent.nodeSize, { [OBJECT_LINE_CARET_ATTR]: "true" }),
  ]);
}

/** The caret's client rect as the browser places it, text-high beside text. */
function caretRect(
  view: EditorView,
): { left: number; top: number; height: number } | null {
  const selection = view.dom.ownerDocument.getSelection();
  if (selection && selection.rangeCount > 0 && selection.isCollapsed) {
    const range = selection.getRangeAt(0);
    if (view.dom.contains(range.startContainer)) {
      const rects = Array.from(range.getClientRects()).filter(
        (rect) => rect.height > 0,
      );
      const rect = rects[rects.length - 1];
      if (rect) return { left: rect.left, top: rect.top, height: rect.height };
    }
  }
  // Between two atoms (no text node to measure): the neighbour decides, the
  // character before the caret first, as Word takes its formatting from it.
  const { $head } = view.state.selection;
  const side = $head.nodeBefore ? -1 : 1;
  try {
    const coords = view.coordsAtPos($head.pos, side);
    return {
      left: side < 0 ? coords.right : coords.left,
      top: coords.top,
      height: coords.bottom - coords.top,
    };
  } catch {
    return null;
  }
}

class ObjectLineCaretView {
  private readonly caret: HTMLElement;
  private frame = 0;
  private lastKey = "";

  constructor(private view: EditorView) {
    const doc = view.dom.ownerDocument;
    this.caret = doc.createElement("div");
    this.caret.className = "note-editor-object-line-caret";
    this.caret.setAttribute("aria-hidden", "true");
    this.caret.hidden = true;
    doc.body.append(this.caret);
    view.dom.addEventListener("focus", this.schedule);
    view.dom.addEventListener("blur", this.schedule);
    view.dom.addEventListener("compositionstart", this.onCompositionStart);
    view.dom.addEventListener("compositionend", this.onCompositionEnd);
    doc.addEventListener("scroll", this.schedule, true);
    doc.defaultView?.addEventListener("resize", this.schedule);
    this.schedule();
  }

  update(view: EditorView): void {
    this.view = view;
    this.schedule();
  }

  // The native caret shows during a composition (the candidate text is the
  // browser's to place); the drawn one steps aside in `paint`.
  private readonly onCompositionStart = () => {
    this.view.dom.setAttribute(COMPOSING_ATTR, "true");
    this.schedule();
  };

  private readonly onCompositionEnd = () => {
    this.view.dom.removeAttribute(COMPOSING_ATTR);
    this.schedule();
  };

  private readonly schedule = () => {
    if (this.frame) return;
    this.frame = window.requestAnimationFrame(() => {
      this.frame = 0;
      this.paint();
    });
  };

  private paint(): void {
    const { view } = this;
    const active =
      view.editable &&
      view.hasFocus() &&
      !isImeCompositionActive(view) &&
      drawnCaretState(view.state);
    const caret = active ? caretRect(view) : null;
    // Beside a tall in-line picture scrolled half out of view the caret runs
    // past the scroller; the part the page does not show is not drawn.
    const rect = caret ? clipToVisibleEditor(view.dom, { ...caret, width: 0 }) : null;
    if (!rect) {
      // Nothing to draw. This runs on every scroll of the document for
      // every editor on it, and the colour read below forced a style
      // recalc each time (a page of unfocused editors paid ~40 ms a frame
      // while scrolling).
      this.caret.hidden = true;
      this.lastKey = "";
      return;
    }
    // The caret colour the theme gives the native caret: the root's, which
    // the paragraph's transparent caret leaves as it is.
    const style = getComputedStyle(view.dom);
    const color =
      style.caretColor && style.caretColor !== "auto"
        ? style.caretColor
        : style.color;
    this.caret.hidden = false;
    this.caret.style.color = color;
    this.caret.style.left = `${rect.left}px`;
    this.caret.style.top = `${rect.top}px`;
    this.caret.style.height = `${rect.height}px`;
    // A caret that moves shows at once and starts blinking again.
    const key = `${view.state.selection.head}:${Math.round(rect.left)}:${Math.round(rect.top)}`;
    if (key !== this.lastKey) {
      this.lastKey = key;
      this.caret.style.animation = "none";
      void this.caret.offsetWidth;
      this.caret.style.animation = "";
    }
  }

  destroy(): void {
    if (this.frame) window.cancelAnimationFrame(this.frame);
    const doc = this.view.dom.ownerDocument;
    this.view.dom.removeEventListener("focus", this.schedule);
    this.view.dom.removeEventListener("blur", this.schedule);
    this.view.dom.removeEventListener("compositionstart", this.onCompositionStart);
    this.view.dom.removeEventListener("compositionend", this.onCompositionEnd);
    doc.removeEventListener("scroll", this.schedule, true);
    doc.defaultView?.removeEventListener("resize", this.schedule);
    this.view.dom.removeAttribute(COMPOSING_ATTR);
    this.caret.remove();
  }
}

const objectLineCaretKey = new PluginKey("documentObjectLineCaret");

export const ObjectLineCaret = Extension.create({
  name: "objectLineCaret",

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: objectLineCaretKey,
        props: { decorations: drawnCaretDecorations },
        view: (view) => new ObjectLineCaretView(view),
      }),
    ];
  },
});

export default ObjectLineCaret;
