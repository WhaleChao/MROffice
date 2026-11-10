/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded undo/redo history for MANUAL PDF content edits (text boxes, ink,
 * image overlays) in the native PDF editor.
 *
 * Adapted from pdf.js's CommandManager (src/display/editor/tools.js,
 * https://github.com/mozilla/pdf.js, Apache-2.0): commands are cmd/undo
 * closure pairs on a bounded stack with a position cursor; adding while the
 * cursor is mid-stack truncates the redo tail; consecutive same-`type`
 * commands can fold into one entry (`overwriteIfSameType` + `keepUndo`) so a
 * continuous gesture — a color-picker scrub, repeated font-size steps —
 * undoes in a single step. Extended with a change subscription so a React
 * toolbar can mirror canUndo/canRedo without polling, and with a string
 * `type` (undefined never folds) instead of pdf.js's NaN-typed numbers.
 *
 * Deliberately NOT entangled with the AI document-edit snapshot mechanism
 * (`rememberCappedSnapshot`) or the formula source editor's own bounded undo
 * stack — both keep working independently; the keyboard layer defers to
 * focused inputs so those surfaces win while focused.
 */

import { isTextEditingTarget } from "./keyboard-nav";

export interface PdfEditCommand {
  /** Applies (and re-applies on redo) the change. Must be idempotent. */
  cmd: () => void;
  /** Reverts the change. Must be idempotent. */
  undo: () => void;
  /** Execute `cmd` immediately as part of `add` (single application path). */
  mustExec?: boolean;
  /**
   * Fold marker: when the entry at the cursor has the same type and the new
   * command sets `overwriteIfSameType`, the new command REPLACES it instead
   * of pushing (one undo step per gesture). `undefined` never folds.
   */
  type?: string;
  overwriteIfSameType?: boolean;
  /** When folding, keep the replaced entry's undo (revert to gesture start). */
  keepUndo?: boolean;
}

export const PDF_EDIT_HISTORY_LIMIT = 50;

interface StoredCommand {
  cmd: () => void;
  undo: () => void;
  type?: string;
}

export class PdfEditHistory {
  #commands: StoredCommand[] = [];

  #locked = false;

  #maxSize: number;

  #position = -1;

  #version = 0;

  #listeners = new Set<() => void>();

  constructor(maxSize = PDF_EDIT_HISTORY_LIMIT) {
    this.#maxSize = Math.max(1, maxSize);
  }

  #notify() {
    this.#version += 1;
    for (const listener of this.#listeners) listener();
  }

  /** Monotonic change counter (useSyncExternalStore snapshot). */
  getVersion = (): number => this.#version;

  /** Subscribe to any history change; returns the unsubscribe function. */
  subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  };

  add({
    cmd,
    undo,
    mustExec = false,
    type,
    overwriteIfSameType = false,
    keepUndo = false,
  }: PdfEditCommand): void {
    if (mustExec) {
      cmd();
    }
    // Re-entrant adds during an undo/redo replay must not corrupt the stack.
    if (this.#locked) {
      return;
    }
    const save: StoredCommand = { cmd, undo, type };
    if (this.#position === -1) {
      // Everything has been undone: a new command clears the redo tail.
      if (this.#commands.length > 0) {
        this.#commands.length = 0;
      }
      this.#position = 0;
      this.#commands.push(save);
      this.#notify();
      return;
    }
    const current = this.#commands[this.#position];
    if (
      overwriteIfSameType &&
      type !== undefined &&
      current.type === type
    ) {
      if (keepUndo) {
        save.undo = current.undo;
      }
      this.#commands[this.#position] = save;
      this.#notify();
      return;
    }
    const next = this.#position + 1;
    if (next === this.#maxSize) {
      this.#commands.splice(0, 1);
    } else {
      this.#position = next;
      if (next < this.#commands.length) {
        this.#commands.splice(next);
      }
    }
    this.#commands.push(save);
    this.#notify();
  }

  undo(): void {
    if (this.#position === -1) {
      return;
    }
    this.#locked = true;
    try {
      this.#commands[this.#position].undo();
    } finally {
      this.#locked = false;
    }
    this.#position -= 1;
    this.#notify();
  }

  redo(): void {
    if (this.#position >= this.#commands.length - 1) {
      return;
    }
    this.#position += 1;
    this.#locked = true;
    try {
      this.#commands[this.#position].cmd();
    } finally {
      this.#locked = false;
    }
    this.#notify();
  }

  canUndo(): boolean {
    return this.#position !== -1;
  }

  canRedo(): boolean {
    return this.#position < this.#commands.length - 1;
  }

  /** Drop the whole history (document reloaded / replaced). */
  clear(): void {
    if (this.#commands.length === 0 && this.#position === -1) return;
    this.#commands.length = 0;
    this.#position = -1;
    this.#notify();
  }
}

// ---- id-keyed before/after records --------------------------------------

/**
 * One box/overlay/stroke change: `before === undefined` means the record did
 * not exist (create); `after === undefined` means it was removed (delete).
 */
export interface KeyedEditHistoryChange<T> {
  id: string;
  before?: T;
  after?: T;
  /**
   * The page this record lives on, when it is NOT the entry's own page. Only a
   * gesture that spans pages sets it — a whole-document find & replace is one
   * user action and has to undo in one step, and folding several single-page
   * entries together cannot express that: folding keeps the FIRST entry's undo
   * and would silently strip the later pages' revert.
   */
  page?: number;
}

/** A user gesture, usually on one page but occasionally spanning several. */
export interface KeyedEditHistoryEntry<T> {
  page: number;
  changes: KeyedEditHistoryChange<T>[];
  /**
   * Fold key for continuous gestures (e.g. `color:0:tb-1`): consecutive
   * entries with the same key collapse into ONE history entry that undoes
   * back to the state before the first step.
   */
  foldKey?: string;
}

/**
 * Record (and immediately apply) an id-keyed entry. `applyRecords` must
 * replace-or-remove each id on the given page with the supplied record — the
 * same replace-by-id semantics the live editing setters use, so replaying a
 * command is idempotent regardless of what happened in between.
 */
export function addKeyedEditHistoryEntry<T>(
  history: PdfEditHistory,
  applyRecords: (
    page: number,
    records: Array<{ id: string; record?: T }>,
  ) => void,
  entry: KeyedEditHistoryEntry<T>,
): void {
  const { page, changes, foldKey } = entry;
  if (changes.length === 0) return;
  // Group by the change's own page (falling back to the entry's) so one command
  // can carry a sweep that crossed pages; applyRecords is called once per page,
  // which is exactly what it already expects.
  const byPage = new Map<number, KeyedEditHistoryChange<T>[]>();
  for (const change of changes) {
    const target = change.page ?? page;
    const bucket = byPage.get(target);
    if (bucket) bucket.push(change);
    else byPage.set(target, [change]);
  }
  const applyEach = (pick: (change: KeyedEditHistoryChange<T>) => T | undefined) => {
    for (const [target, group] of byPage) {
      applyRecords(
        target,
        group.map((change) => ({ id: change.id, record: pick(change) })),
      );
    }
  };
  history.add({
    cmd: () => applyEach((change) => change.after),
    undo: () => applyEach((change) => change.before),
    mustExec: true,
    type: foldKey,
    overwriteIfSameType: Boolean(foldKey),
    keepUndo: true,
  });
}

// ---- keyboard routing -----------------------------------------------------

/** The keydown fields the shortcut router reads. */
export interface EditHistoryKeyEvent {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
}









export function resolveEditHistoryShortcut(
  event: EditHistoryKeyEvent,
  target?: EventTarget | null,
): "undo" | "redo" | null {
  if (!event || typeof event.key !== "string") return null;
  if (!(event.metaKey || event.ctrlKey) || event.altKey) return null;
  if (event.key.toLowerCase() !== "z") return null;
  if (target !== undefined && isTextEditingTarget(target)) return null;
  return event.shiftKey ? "redo" : "undo";
}
