"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension } from "@tiptap/core";
import { Selection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { Mapping } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";































const GUARD_NAME = "derivedDispatchGuard";

/** Meta of the one nudge that follows a dropped derived transaction. */
export const DERIVED_DISPATCH_RETRY_META = "docsDerivedDispatchRetry";

/**
 * Origins that are not derived even without history: a remote edit or a
 * shared-history step arriving through the collaboration bridge / y-sync, and
 * the history plugin's own Undo/Redo. They are never held, replayed or dropped
 * here: they keep the behaviour they had.
 */
const PROTECTED_ORIGIN_METAS = [
  "y-sync$",
  "mrnotePagedCanonicalRemote",
  "mrnotePagedCanonicalHistory",
  "mrnotePagedHistoryApplied",
  "history$",
] as const;

/**
 * Derived writes that carry steps: the paginator's cuts, moves and joins
 * (structuralRebalance: the ledger, its reconcile and column passes), the
 * note-list slice repair, and the floating text box's flow. A step-carrying
 * transaction without history that names none of these (a fill's held edits
 * replayed by the collaboration bridge, a style-set restyle) is someone's
 * edit, not a projection: it is left as it was.
 */
const LAYOUT_STEP_METAS = ["structuralRebalance", "mrnoteNoteSliceIntegrity", "floatingTextBoxFlow"] as const;

/** Drops in one window before a diagnostic; nudges stop past it. */
const DROP_DIAGNOSTIC_LIMIT = 8;
const DROP_WINDOW_MS = 2000;

interface PendingDomInput {
  flush(): void;
  forceFlush?: () => void;
  flushingSoon?: number;
  queue?: MutationRecord[];
  observer?: MutationObserver | null;
}

interface DropBook {
  windowStart: number;
  count: number;
  warned: boolean;
  nudgeFrame: number | null;
}

/** Views reading back their DOM now, and the mapping of what that read applied. */
const reading = new WeakMap<EditorView, Mapping>();
const drops = new WeakMap<EditorView, DropBook>();

/**
 * A transaction that changes no content (a projection's decorations, a
 * plugin's meta, a settle style, a caret), or one of the layout writes above;
 * from no remote or history origin.
 */
export function isDerivedTransaction(tr: Transaction): boolean {
  if (tr.getMeta(DERIVED_DISPATCH_RETRY_META) === true) return false;
  if (tr.steps.length > 0 && !LAYOUT_STEP_METAS.some((key) => tr.getMeta(key) === true)) return false;
  return PROTECTED_ORIGIN_METAS.every((key) => tr.getMeta(key) === undefined);
}

/**
 * The DOM changes the view has not read yet. Taken from the observer the way
 * DOMObserver.stop() takes them (into its queue, which flush() reads first), so
 * nothing is lost by looking. Inside the view's own read-back the queue is
 * already empty: a selection it dispatches there never reads again.
 */
function unreadDomChanges(observer: PendingDomInput): number {
  const queue = observer.queue;
  if (!Array.isArray(queue)) return 0;
  const taken = observer.observer?.takeRecords?.() ?? [];
  for (const record of taken) queue.push(record);
  return queue.length;
}

/**
 * Have the view read back the DOM changes it has not read yet (typed text,
 * and the selection with it). The mapping of what that applied when it changed
 * the view's state, else null. Never during an IME composition (the composing
 * DOM belongs to the IME and ProseMirror reads it at its own time), and never
 * re-entered.
 */
export function readPendingDomInput(view: EditorView): Mapping | null {
  if (reading.has(view) || view.composing || view.isDestroyed) return null;
  const observer = (view as EditorView & { domObserver?: PendingDomInput }).domObserver;
  if (!observer || typeof observer.flush !== "function") return null;
  if (unreadDomChanges(observer) === 0) return null;
  const before = view.state;
  const mapping = new Mapping();
  reading.set(view, mapping);
  try {
    if (typeof observer.forceFlush === "function" && typeof observer.flushingSoon === "number" && observer.flushingSoon > -1) {
      observer.forceFlush();
    } else {
      observer.flush();
    }
  } catch {
    // A read-back that fails leaves the state as it was: the transaction goes on.
  } finally {
    reading.delete(view);
  }
  return view.state !== before ? mapping : null;
}

/**
 * The derived transaction's steps, meta and own selection on `state`;
 * `mapping` carries a selection it set across what the read-back applied
 * (step-less transactions only: one with steps is replayed only on the
 * document it was made for).
 */
export function replayOn(tr: Transaction, state: EditorState, mapping?: Mapping): Transaction | null {
  const next = state.tr;
  for (const step of tr.steps) {
    if (next.maybeStep(step).failed) return null;
  }
  const meta = (tr as unknown as { meta?: Record<string, unknown> }).meta;
  if (meta) for (const key of Object.keys(meta)) next.setMeta(key, meta[key]);
  if (tr.selectionSet) {
    try {
      const own = Selection.fromJSON(next.doc.eq(tr.doc) ? next.doc : tr.doc, tr.selection.toJSON());
      next.setSelection(next.doc.eq(tr.doc) ? own : own.map(next.doc, mapping ?? new Mapping()));
    } catch {
      return null;
    }
  }
  if (tr.storedMarksSet) next.setStoredMarks(tr.storedMarks);
  if (tr.scrolledIntoView) next.scrollIntoView();
  next.setTime(tr.time);
  return next;
}

function noteDropped(view: EditorView): void {
  const now = Date.now();
  let book = drops.get(view);
  if (!book || now - book.windowStart > DROP_WINDOW_MS) {
    book = { windowStart: now, count: 0, warned: book?.warned ?? false, nudgeFrame: book?.nudgeFrame ?? null };
    drops.set(view, book);
  }
  book.count += 1;
  if (book.count > DROP_DIAGNOSTIC_LIMIT) {
    if (!book.warned) {
      book.warned = true;
      console.warn(`[docs] ${GUARD_NAME}: ${book.count} derived transactions dropped within ${DROP_WINDOW_MS} ms for unread input`);
    }
    return;
  }
  if (book.nudgeFrame !== null) return;
  const win = view.dom.ownerDocument.defaultView;
  if (!win) return;
  const entry = book;
  const nudge = () => {
    entry.nudgeFrame = null;
    if (view.isDestroyed) return;
    view.dispatch(view.state.tr.setMeta(DERIVED_DISPATCH_RETRY_META, true).setMeta("addToHistory", false));
  };
  entry.nudgeFrame = typeof win.requestAnimationFrame === "function" ? win.requestAnimationFrame(nudge) : win.setTimeout(nudge, 16);
}

/**
 * Decide what becomes of a derived transaction about to be dispatched on
 * `view`: itself, its replay on the state the read-back left, or nothing.
 */
/** Harness switch (`window.__docsDerivedDispatchGuardOff = true`): every
 * transaction goes on as dispatched, as before the guard (A/B of one build). */
function guardOff(view: EditorView): boolean {
  const win = view.dom.ownerDocument.defaultView as (Window & { __docsDerivedDispatchGuardOff?: boolean }) | null;
  return win?.__docsDerivedDispatchGuardOff === true;
}

export function guardDerivedTransaction(view: EditorView, tr: Transaction): Transaction | null {
  if (guardOff(view)) return tr;
  const outer = reading.get(view);
  if (outer) {
    // A transaction the read-back itself applies: what it maps is what a
    // replayed caret is carried across.
    outer.appendMapping(tr.mapping);
    return tr;
  }
  if (!isDerivedTransaction(tr)) return tr;
  const mapping = readPendingDomInput(view);
  if (!mapping) return tr;
  const current = view.state;
  // A transaction with no steps applies to any document: once the typed text
  // is read, its redraw replaces nothing unread. One with steps applies only
  // to the document it was made for.
  const replayed = tr.steps.length === 0 || current.doc.eq(tr.before) ? replayOn(tr, current, mapping) : null;
  if (replayed) return replayed;
  noteDropped(view);
  return null;
}

export const DerivedDispatchGuard = Extension.create({
  name: GUARD_NAME,

  dispatchTransaction({ transaction, next }) {
    let view: EditorView | null = null;
    try {
      view = this.editor.view;
    } catch {
      view = null;
    }
    if (!view) {
      next(transaction);
      return;
    }
    const guarded = guardDerivedTransaction(view, transaction);
    if (guarded) next(guarded);
  },
});

export default DerivedDispatchGuard;
