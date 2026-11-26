"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension } from "@tiptap/core";
import { Plugin, PluginKey, type Transaction } from "@tiptap/pm/state";
import type { Slice } from "@tiptap/pm/model";
import { ReplaceStep } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";





















const GUARD_NAME = "readBackLossGuard";

/** Fewest characters a refused read-back loses beyond what the input may. */
export const LOSS_MIN_CHARS = 8;
/** How long after its beforeinput a read-back still belongs to that input. */
const INPUT_WINDOW_MS = 1000;

export interface TypedInput {
  inputType: string;
  data: string;
  at: number;
}

interface FlushingObserver {
  flush(): void;
  readBackLossGuardWrapped?: boolean;
}

const inputs = new WeakMap<EditorView, TypedInput>();
const readBackDepth = new WeakMap<EditorView, number>();
let refusedCount = 0;

/** How many read-backs this page refused (diagnostics). */
export function refusedReadBacks(): number {
  return refusedCount;
}

/** True while the view reads its DOM back (DOMObserver.flush). */
export function isReadingBack(view: EditorView): boolean {
  return (readBackDepth.get(view) ?? 0) > 0;
}

function wrapFlush(view: EditorView): void {
  const observer = (view as EditorView & { domObserver?: FlushingObserver }).domObserver;
  if (!observer || typeof observer.flush !== "function" || observer.readBackLossGuardWrapped) return;
  const flush = observer.flush;
  observer.flush = function guardedFlush(this: FlushingObserver) {
    readBackDepth.set(view, (readBackDepth.get(view) ?? 0) + 1);
    try {
      return flush.call(this);
    } finally {
      readBackDepth.set(view, Math.max(0, (readBackDepth.get(view) ?? 1) - 1));
    }
  };
  observer.readBackLossGuardWrapped = true;
}

/** Characters the transaction's replace steps delete from the document. */
export function deletedCharacters(tr: Transaction): number {
  let deleted = 0;
  tr.steps.forEach((step, index) => {
    if (!(step instanceof ReplaceStep)) return;
    const { from, to } = step as unknown as { from: number; to: number };
    if (to > from) deleted += tr.docs[index].textBetween(from, to, "\n", "\ufffc").length;
  });
  return deleted;
}

/** Characters the transaction's replace steps put in. */
export function insertedCharacters(tr: Transaction): number {
  let inserted = 0;
  for (const step of tr.steps) {
    if (!(step instanceof ReplaceStep)) continue;
    const content = (step as unknown as { slice: Slice }).slice.content;
    inserted += content.textBetween(0, content.size, "\n", "\ufffc").length;
  }
  return inserted;
}

/**
 * The read-back loses text the last input cannot have deleted: a plain
 * insertText of one or two characters, `selected` the characters the
 * selection held when it was typed.
 */
export function readBackLosesText(tr: Transaction, input: TypedInput | undefined, selected: number, now: number): boolean {
  if (!input || now - input.at > INPUT_WINDOW_MS) return false;
  if (input.inputType !== "insertText" || input.data.length === 0 || input.data.length > 2) return false;
  if (!tr.docChanged) return false;
  // What it takes out beyond what it puts in: an autocorrection the key set
  // off (a misspelled word replaced) puts as much back.
  return deletedCharacters(tr) - insertedCharacters(tr) >= selected + LOSS_MIN_CHARS;
}

/** Harness switch (`window.__docsReadBackLossGuardOff = true`): every
 * read-back goes on as before the guard (A/B of one build). */
function guardOff(view: EditorView): boolean {
  const win = view.dom.ownerDocument.defaultView as (Window & { __docsReadBackLossGuardOff?: boolean }) | null;
  return win?.__docsReadBackLossGuardOff === true;
}

function typeAtSelection(view: EditorView, data: string): void {
  const { from, to } = view.state.selection;
  const insert = () => view.state.tr.insertText(data, from, to);
  if (!view.someProp("handleTextInput", (handle) => handle(view, from, to, data, insert))) {
    view.dispatch(insert().scrollIntoView());
  }
}

export const ReadBackLossGuard = Extension.create({
  name: GUARD_NAME,

  dispatchTransaction({ transaction, next }) {
    let view: EditorView | null = null;
    try {
      view = this.editor.view;
    } catch {
      view = null;
    }
    if (!view || view.composing || !isReadingBack(view) || guardOff(view)) {
      next(transaction);
      return;
    }
    const input = inputs.get(view);
    const { from, to } = view.state.selection;
    const selected = to > from ? view.state.doc.textBetween(from, to, "\n", "\ufffc").length : 0;
    if (!input || !readBackLosesText(transaction, input, selected, Date.now())) {
      next(transaction);
      return;
    }
    inputs.delete(view);
    refusedCount += 1;
    console.warn(`[docs] ${GUARD_NAME}: refused a read-back deleting ${deletedCharacters(transaction)} characters for one typed character`);
    // The flush that carried it redraws the range from the document; the key
    // goes in after that, as an update of its own.
    const target = view;
    queueMicrotask(() => {
      if (!target.isDestroyed) typeAtSelection(target, input.data);
    });
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey(GUARD_NAME),
        props: {
          handleDOMEvents: {
            beforeinput: (view, event) => {
              const input = event as InputEvent;
              inputs.set(view, { inputType: input.inputType, data: input.data ?? "", at: Date.now() });
              return false;
            },
          },
        },
        view(view) {
          wrapFlush(view);
          return {};
        },
      }),
    ];
  },
});

export default ReadBackLossGuard;
