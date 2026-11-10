/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

/** A state smaller than this is written whole with every copy, as before. */
export const PAGE_CAUSAL_JOURNAL_MIN_BYTES = 256 * 1024;
/** The tail that makes the next base: this share of the base, at least
 * {@link PAGE_CAUSAL_JOURNAL_TAIL_MIN_BYTES}. */
const TAIL_SHARE_OF_BASE = 1 / 8;
const PAGE_CAUSAL_JOURNAL_TAIL_MIN_BYTES = 64 * 1024;
/** A tail this many shares long (no new base stored all that while) ends the
 * journal: the copies are whole again. */
const TAIL_GIVE_UP_SHARES = 8;
/** A kept base is stored again (its age refreshed: the store drops entries a
 * day old) this often. */
export const PAGE_CAUSAL_JOURNAL_BASE_REFRESH_MS = 60 * 60 * 1_000;

interface JournalBase {
  readonly base64: string;
  readonly bytes: number;
  /** When the store last confirmed it holds this base (0: not yet). */
  storedAt: number;
}

interface DocJournal {
  base: JournalBase;
  /** The document's updates since `base` was taken, in order. */
  updates: Uint8Array[];
  updateBytes: number;
  /** A newer base being stored, and how many of `updates` it already holds. */
  candidate: (JournalBase & { readonly holds: number }) | null;
  /** The last merge of the first `count` of `updates`. */
  merged: { readonly count: number; readonly tail: Uint8Array } | null;
  detach: () => void;
}

/** `updates` merged, from the last merge and the updates after it. */
function mergedTail(journal: DocJournal): Uint8Array {
  const { updates, merged } = journal;
  if (updates.length === 0) return new Uint8Array(0);
  if (merged?.count === updates.length) return merged.tail;
  const from = merged && merged.count < updates.length ? merged.count : 0;
  const parts = from > 0 ? [merged!.tail, ...updates.slice(from)] : updates;
  const tail = parts.length === 1 ? parts[0] : Y.mergeUpdates(parts);
  journal.merged = { count: updates.length, tail };
  return tail;
}

function forget(doc: Y.Doc, journal: DocJournal): void {
  journal.detach();
  if (journals.get(doc) === journal) journals.delete(doc);
}

/** What a recovery copy can carry for its state instead of the whole of it. */
export interface PageCausalJournalForm {
  /** The base, as base64 (the blob the copy names). */
  readonly baseBase64: string;
  /** Whether the store has confirmed it holds the base. */
  readonly baseStored: boolean;
  /** The updates since the base, merged; empty when there are none. */
  readonly tail: Uint8Array;
  /** The body the capture was for, when it is the state's projection itself
   * (a copy that carries this very body may leave it out). */
  readonly body: object | null;
  /** Strings the store should keep for this journal (the base when it is not
   * confirmed or is due for a refresh, a compaction's new base), and the call
   * that tells the journal the store now holds them. */
  readonly toStore: readonly string[];
  stored(values: readonly string[]): void;
}

const journals = new WeakMap<Y.Doc, DocJournal>();
const forms = new WeakMap<object, PageCausalJournalForm>();

/**
 * Note a capture of `doc`'s whole state (`update`, its base64 `updateBase64`)
 * as recovery object `recovery`, made for `body` (null when the body is not
 * the state's projection). The caller has just encoded `update` from `doc`
 * with no transaction in between.
 */
export function notePageCausalCapture(
  doc: Y.Doc,
  recovery: object,
  update: Uint8Array,
  updateBase64: string,
  body: object | null,
): void {
  let journal = journals.get(doc);
  if (!journal) {
    if (update.length < PAGE_CAUSAL_JOURNAL_MIN_BYTES) return;
    journal = { base: { base64: updateBase64, bytes: update.length, storedAt: 0 },
      updates: [], updateBytes: 0, candidate: null, merged: null, detach: () => undefined };
    const created = journal;
    created.detach = () => { doc.off("update", onUpdate); doc.off("destroy", onDestroy); };
    const onUpdate = (change: Uint8Array) => {
      created.updates.push(change);
      created.updateBytes += change.length;
    };
    const onDestroy = () => forget(doc, created);
    doc.on("update", onUpdate);
    doc.on("destroy", onDestroy);
    journals.set(doc, journal);
  }
  const current = journal;
  // The state as it is now: the base and the tail make it exactly.
  let tail = mergedTail(current);
  const share = Math.max(PAGE_CAUSAL_JOURNAL_TAIL_MIN_BYTES, current.base.bytes * TAIL_SHARE_OF_BASE);
  if (tail.length > share && current.base.storedAt === 0) {
    // No copy names a base the store never confirmed: the state now is the
    // base instead (what a store that keeps refusing costs is a whole copy
    // each time, as before, never a tail that grows without end).
    current.base = { base64: updateBase64, bytes: update.length, storedAt: 0 };
    current.updates = [];
    current.updateBytes = 0;
    current.candidate = null;
    current.merged = null;
    tail = new Uint8Array(0);
  } else if (tail.length > share * TAIL_GIVE_UP_SHARES) {
    // The store holds the base but not one compaction since: whole copies.
    forget(doc, current);
    return;
  } else if (!current.candidate && current.base.storedAt > 0 && tail.length > share) {
    // Grown past its share of the base: the state now is the next base. The
    // copies go on naming the current one until the store holds the new one.
    current.candidate = { base64: updateBase64, bytes: update.length, storedAt: 0, holds: current.updates.length };
  }
  const base = current.base;
  const now = Date.now();
  const toStore: string[] = [];
  if (base.storedAt === 0 || now - base.storedAt >= PAGE_CAUSAL_JOURNAL_BASE_REFRESH_MS) toStore.push(base.base64);
  if (current.candidate) toStore.push(current.candidate.base64);
  const form: PageCausalJournalForm = Object.freeze({
    baseBase64: base.base64,
    baseStored: base.storedAt > 0,
    tail,
    body,
    toStore,
    stored: (values: readonly string[]) => {
      const at = Date.now();
      if (journals.get(doc) !== current) return;
      if (values.includes(current.base.base64)) current.base.storedAt = at;
      const next = current.candidate;
      if (next && values.includes(next.base64)) {
        // The new base is kept: the copies name it from now on, with the
        // updates it does not hold yet as their tail.
        current.base = { base64: next.base64, bytes: next.bytes, storedAt: at };
        current.updates = current.updates.slice(next.holds);
        current.updateBytes = current.updates.reduce((sum, change) => sum + change.length, 0);
        current.candidate = null;
        current.merged = null;
      }
    },
  });
  forms.set(recovery, form);
}

/** The journal form of a recovery object this module was told of, or null. */
export function pageCausalJournalForm(recovery: unknown): PageCausalJournalForm | null {
  return recovery !== null && typeof recovery === "object" ? forms.get(recovery) ?? null : null;
}
