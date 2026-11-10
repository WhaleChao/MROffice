/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { forEachDiagnostic } from "@codemirror/lint";
import {
  Transaction,
  type EditorState,
  type Extension,
} from "@codemirror/state";
import {
  activateHover,
  closeHoverTooltips,
  hasHoverTooltips,
  keymap,
  ViewPlugin,
  type EditorView,
  type ViewUpdate,
} from "@codemirror/view";

import { isTouchLikePointer } from "@mrnote/core/web/lib/pointer-capability.ts";

/** How long a tapped caret rests before its hover opens (touch only). */
export const CARET_HOVER_DELAY_MS = 400;

/** An opened hover closes on the next edit or caret move (VS Code). */
function closesShownHover(transaction: Transaction): boolean {
  return transaction.docChanged || transaction.selection !== undefined;
}

function isWordChar(char: string): boolean {
  return /[\p{L}\p{N}_$]/u.test(char);
}

/**
 * The side of `pos` the hover reads: the character after the caret, unless
 * only the one before it belongs to a word (a caret at a word's end).
 */
export function hoverSideAt(state: EditorState, pos: number): -1 | 1 {
  const after = state.doc.sliceString(pos, pos + 1);
  const before = pos > 0 ? state.doc.sliceString(pos - 1, pos) : "";
  if (!isWordChar(after) && isWordChar(before)) return -1;
  return 1;
}

/** True when the caret touches a word or lies inside a diagnostic's range. */
export function caretHasHoverSubject(state: EditorState, pos: number): boolean {
  if (state.wordAt(pos)) return true;
  let onDiagnostic = false;
  forEachDiagnostic(state, (_diagnostic, from, to) => {
    if (from <= pos && pos <= to) onDiagnostic = true;
  });
  return onDiagnostic;
}

/** Open every hover tooltip at `pos` (default: the main caret). */
export function showHoverAt(view: EditorView, pos?: number): void {
  const at = pos ?? view.state.selection.main.head;
  activateHover(view, at, hoverSideAt(view.state, at), {
    until: closesShownHover,
  });
}

/** Escape closes an open hover; otherwise it falls through to the rest. */
function closeHoverOnEscape(view: EditorView): boolean {
  if (!hasHoverTooltips(view.state)) return false;
  view.dispatch({ effects: closeHoverTooltips });
  return true;
}

export interface CaretHoverOnTouchOptions {
  delayMs?: number;
}

/** How long after a finger touched the text its caret move counts as a tap. */
const TAP_SELECTION_WINDOW_MS = 1000;

/**
 * Escape-to-close for any hover, plus the touch caret hover. Mount once per
 * editor next to the hover tooltip sources.
 */
export function showHoverSupport(
  options: CaretHoverOnTouchOptions = {},
): Extension {
  const delay = options.delayMs ?? CARET_HOVER_DELAY_MS;
  const caretHover = ViewPlugin.fromClass(
    class {
      timer: ReturnType<typeof setTimeout> | null = null;
      /** When a finger last pressed the text (pointerdown, touch-like). */
      touchedAt = Number.NEGATIVE_INFINITY;

      constructor(readonly view: EditorView) {}

      update(update: ViewUpdate) {
        if (!update.selectionSet && !update.docChanged) return;
        this.clear();
        if (update.docChanged) return;
        // CodeMirror leaves a finger's tap to the browser and reads the caret
        // back from the DOM, so the move carries no user event (a mouse's
        // carries "select.pointer"); a keyboard's arrows carry "select".
        if (Date.now() - this.touchedAt > TAP_SELECTION_WINDOW_MS) return;
        const tapped = update.transactions.every((transaction) => {
          const event = transaction.annotation(Transaction.userEvent);
          return event === undefined || event === "select.pointer";
        });
        if (!tapped) return;
        const { selection } = update.state;
        if (selection.ranges.length !== 1 || !selection.main.empty) return;
        const pos = selection.main.head;
        this.timer = setTimeout(() => {
          this.timer = null;
          const { state } = this.view;
          const main = state.selection.main;
          if (state.selection.ranges.length !== 1 || !main.empty) return;
          if (main.head !== pos || hasHoverTooltips(state)) return;
          if (!caretHasHoverSubject(state, pos)) return;
          showHoverAt(this.view, pos);
        }, delay);
      }

      clear() {
        if (this.timer !== null) clearTimeout(this.timer);
        this.timer = null;
      }

      destroy() {
        this.clear();
      }
    },
    {
      eventHandlers: {
        pointerdown(event) {
          if (isTouchLikePointer(event)) this.touchedAt = Date.now();
        },
      },
    },
  );
  return [keymap.of([{ key: "Escape", run: closeHoverOnEscape }]), caretHover];
}
