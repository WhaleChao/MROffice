/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A caret a person just put down is kept when the editor's own work redraws
 * the document before ProseMirror has read it.
 *
 * A click into text is left to the browser: it moves the DOM selection, and
 * ProseMirror takes it over on the "selectionchange" event that follows (a
 * task later, tens to hundreds of ms on a loaded machine). A transaction that
 * redraws the view in between -- one that changes the document or its
 * decorations without choosing a selection: a Word page's layout passes,
 * page rebalancing, line projections -- writes ProseMirror's previous
 * selection back into the DOM (`EditorView.updateStateInner` ->
 * `selectionToDOM`; it holds back only while a mouse button is down), and
 * the event then finds nothing new. The person's click is gone: the caret
 * stays where it was, and what they type lands there. With the cold first
 * layout a Word document keeps laying out its pages for seconds after it
 * turns editable; a click and 40 keys typed at once landed at the end of the
 * document in about a third of the runs on a loaded machine (cold10 gate,
 * float.docx: the keys went into its trailing paragraph, not the clicked one).
 *
 * Such a transaction now carries the selection the DOM holds and ProseMirror
 * has not read yet (`appendTransaction`, before the view redraws): the same
 * one ProseMirror would read, mapped through the transaction. Transactions
 * that set a selection of their own (typing, commands, remote edits restoring
 * the caret) are left as they are, as is a view that is not focused, is
 * composing, or ignores the selection where it is.
 */
import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { Mapping } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";

interface DomSelectionRange {
  readonly anchorNode: Node | null;
  readonly anchorOffset: number;
  readonly focusNode: Node | null;
  readonly focusOffset: number;
}

/** ProseMirror's own view of the DOM selection (prosemirror-view internals,
 * read only; absent in another version, and then nothing is done). */
interface ObservedView {
  readonly composing?: boolean;
  domSelectionRange?(): DomSelectionRange;
  readonly domObserver?: {
    readonly currentSelection?: { eq(range: DomSelectionRange): boolean };
    readonly suppressingSelectionUpdates?: boolean;
    ignoreSelectionChange?(range: DomSelectionRange): boolean;
  };
}

export const pendingDomSelectionKey = new PluginKey("pendingDomSelection");

/** The DOM selection a person made that `view` has not read yet, as document
 * positions of the state the view shows; null when there is none. */
export function unreadDomSelection(view: EditorView): { anchor: number; head: number } | null {
  const observed = view as unknown as ObservedView;
  const observer = observed.domObserver;
  if (!observer?.currentSelection || typeof observed.domSelectionRange !== "function") return null;
  if (observed.composing || observer.suppressingSelectionUpdates || !view.hasFocus()) return null;
  const range = observed.domSelectionRange();
  const { anchorNode, focusNode } = range;
  if (!anchorNode || !focusNode || observer.currentSelection.eq(range)) return null;
  if (!view.dom.contains(anchorNode) || !view.dom.contains(focusNode)) return null;
  if (observer.ignoreSelectionChange?.(range)) return null;
  // Only a selection in the document's editable text: one inside a widget or
  // an uneditable island is the view's to settle, as before.
  for (const node of [anchorNode, focusNode]) {
    const element = node.nodeType === 1 ? (node as Element) : node.parentElement;
    const island = element?.closest('[contenteditable="false"]');
    if (!element || (island && view.dom.contains(island))) return null;
  }
  try {
    const anchor = view.posAtDOM(anchorNode, range.anchorOffset);
    const head = view.posAtDOM(focusNode, range.focusOffset);
    const size = view.state.doc.content.size;
    if (!(anchor >= 0 && head >= 0 && anchor <= size && head <= size)) return null;
    const current = view.state.selection;
    if (current.anchor === anchor && current.head === head) return null;
    return { anchor, head };
  } catch {
    return null;
  }
}

/** The transaction that keeps `pending` (positions in `before`) through
 * `transactions` on `after`, or null when it cannot be kept. */
export function keepPendingSelection(
  pending: { anchor: number; head: number },
  transactions: readonly Transaction[],
  after: EditorState,
): Transaction | null {
  const mapping = new Mapping();
  for (const tr of transactions) mapping.appendMapping(tr.mapping);
  const anchor = mapping.mapResult(pending.anchor, -1);
  const head = mapping.mapResult(pending.head, -1);
  if (anchor.deleted || head.deleted) return null;
  const size = after.doc.content.size;
  if (anchor.pos > size || head.pos > size) return null;
  const selection = TextSelection.between(after.doc.resolve(anchor.pos), after.doc.resolve(head.pos));
  if (selection.eq(after.selection)) return null;
  return after.tr.setSelection(selection).setMeta("addToHistory", false).setMeta(pendingDomSelectionKey, true);
}

export function pendingDomSelectionPlugin(): Plugin {
  let view: EditorView | null = null;
  return new Plugin({
    key: pendingDomSelectionKey,
    view(editorView) {
      view = editorView;
      return { destroy() { if (view === editorView) view = null; } };
    },
    appendTransaction(transactions, before, after) {
      // Only while the view still shows `before` (a dispatch about to
      // redraw it), and only for the editor's own work, which chose no
      // selection of its own.
      if (!view || view.state !== before) return null;
      if (transactions.some((tr) => tr.selectionSet || tr.getMeta(pendingDomSelectionKey))) return null;
      const pending = unreadDomSelection(view);
      return pending ? keepPendingSelection(pending, transactions, after) : null;
    },
  });
}

export const PendingDomSelection = Extension.create({
  name: "pendingDomSelection",
  addProseMirrorPlugins() {
    return [pendingDomSelectionPlugin()];
  },
});
