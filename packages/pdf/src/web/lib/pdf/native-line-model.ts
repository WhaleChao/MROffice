/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the reader's edits have made of the page's lines, ahead of the bytes.
 *
 * The working copy's revision of an edit takes a moment to be compiled, drawn
 * and read in; the reader does not wait for it. For that moment this model is
 * what a line reads: leaving a line, pressing it again, an Undo and a Redo all
 * read and move it at once, and the line is opened with what it says. Lines it
 * says nothing of read as the page shows them.
 *
 * Each edit is kept by the operation that staged it, with the words before and
 * after, so a step back or on is known here without reading a revision. A step
 * it does not know (a format, a move, a saved edit taken back) leaves it unable
 * to say what any line reads until the revision is up: `clear`.
 */

/** A line's box on its page, in fractions of the page as shown. */
export type PdfNativeLineFrame = Readonly<{ pageNumber: number; x: number; y: number; width: number; height: number }>;

/** The same line: on one page, each box's centre inside the other (a line
 * above or below, or the next box along a row, has neither). */
export function samePdfNativeLineFrame(left: PdfNativeLineFrame | null | undefined,
  right: PdfNativeLineFrame | null | undefined): boolean {
  if (!left || !right || left.pageNumber !== right.pageNumber) return false;
  const inside = (outer: PdfNativeLineFrame, inner: PdfNativeLineFrame) => {
    const x = inner.x + inner.width / 2, y = inner.y + inner.height / 2;
    return x >= outer.x && x <= outer.x + outer.width && y >= outer.y && y <= outer.y + outer.height;
  };
  return inside(left, right) && inside(right, left);
}

/** A line as a press or a finish names it: by the key of its items, its id
 * across revisions, the words its box was marked with, or where it stands. */
export type PdfNativeLineName = Readonly<{ key: string; lineId?: string | null;
  frame?: PdfNativeLineFrame | null; markedWords?: string | null }>;

export type PdfNativeLineRecord = Readonly<{ key: string; lineId: string | null;
  frame: PdfNativeLineFrame | null; words: string }>;

type Edit = Readonly<{ key: string; lineId: string | null; frame: PdfNativeLineFrame | null;
  before: string | null; after: string }>;

const KEPT_LINES = 8;
const KEPT_EDITS = 64;

export interface PdfNativeLineModel {
  /** An edit left on a line: the line reads `after` from now on. `before`: what
   * it read, when known (a step back needs it). `isOperation`: this edit is
   * that operation of the working copy, once staged. */
  edited(line: PdfNativeLineRecord & { before?: string | null }): Readonly<{ isOperation(opId: string): void }>;
  /** The edit left on `key` went nowhere: the line reads what the page shows. */
  dropped(key: string): void;
  /** The words a line reads ahead of the bytes; null when it reads as shown. */
  wordsOf(name: PdfNativeLineName): PdfNativeLineRecord | null;
  /** Whether an edit left on this line is still being staged. */
  inFlight(name: Pick<PdfNativeLineName, "key" | "lineId">): boolean;
  began(name: Pick<PdfNativeLineName, "key" | "lineId">): () => void;
  /** An Undo or Redo of `opId`: true when the model knows what it leaves (and
   * now reads it); false leaves it cleared. `line` gets what the line reads. */
  stepped(opId: string | undefined, direction: "undo" | "redo"): PdfNativeLineRecord | null;
  clear(): void;
}

export function createPdfNativeLineModel(): PdfNativeLineModel {
  let lines: readonly PdfNativeLineRecord[] = [];
  const edits = new Map<string, Edit>();
  const flying = new Map<string, number>();
  const flightKeys = (name: Pick<PdfNativeLineName, "key" | "lineId">) =>
    [name.key, ...(name.lineId ? [`id:${name.lineId}`] : [])];
  const put = (record: PdfNativeLineRecord) => {
    lines = [...lines.filter(entry => entry.key !== record.key &&
      !(record.lineId && entry.lineId === record.lineId)), record].slice(-KEPT_LINES);
  };
  const remember = (opId: string, edit: Edit) => {
    edits.delete(opId);
    edits.set(opId, edit);
    while (edits.size > KEPT_EDITS) edits.delete(edits.keys().next().value as string);
  };
  return {
    edited(line) {
      const { before = null, ...record } = line;
      put(record);
      const edit: Edit = { key: record.key, lineId: record.lineId, frame: record.frame, before, after: record.words };
      return { isOperation(opId) { remember(opId, edit); } };
    },
    dropped(key) {
      lines = lines.filter(entry => entry.key !== key);
    },
    wordsOf(name) {
      let found: PdfNativeLineRecord | null = null;
      for (const record of lines) {
        if (record.key === name.key || (record.lineId && record.lineId === name.lineId) ||
          (typeof name.markedWords === "string" && name.markedWords === record.words) ||
          samePdfNativeLineFrame(record.frame, name.frame)) found = record;
      }
      return found;
    },
    inFlight(name) {
      return flightKeys(name).some(key => (flying.get(key) ?? 0) > 0);
    },
    began(name) {
      const keys = flightKeys(name);
      for (const key of keys) flying.set(key, (flying.get(key) ?? 0) + 1);
      let ended = false;
      return () => {
        if (ended) return;
        ended = true;
        for (const key of keys) {
          const left = (flying.get(key) ?? 1) - 1;
          if (left > 0) flying.set(key, left); else flying.delete(key);
        }
      };
    },
    stepped(opId, direction) {
      const edit = opId ? edits.get(opId) : undefined;
      if (!edit || (direction === "undo" && edit.before === null)) { this.clear(); return null; }
      const record = Object.freeze({ key: edit.key, lineId: edit.lineId, frame: edit.frame,
        words: direction === "undo" ? edit.before! : edit.after });
      put(record);
      return record;
    },
    clear() {
      lines = [];
    },
  };
}
