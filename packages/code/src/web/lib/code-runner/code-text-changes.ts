/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Edits of a code page's text that carry their own places.
 *
 * Code repeats characters everywhere (brackets, indentation, `==`). An edit
 * found by comparing the old and new text can only guess which copy of a
 * repeated character its letters follow: "(a" typed before "(" reads as "a("
 * typed after it. In the shared Y.Text that guess decides which character the
 * new letters are anchored to, so a collaborator's key typed at the same
 * moment came out inside the run (TRIAGE 47). An edit that knows where it was
 * made (a Find-in-Files hit, a refactor's text change, an editor transaction)
 * is therefore written at its own places; a whole-text write is placed hunk
 * by hunk and refused when any hunk's place is a guess.
 */

import { diff } from "@codemirror/merge";
import type * as Y from "yjs";

/** Replace `[from, to)` of the text the edit was made against with `insert`. */
export interface CodeTextChange {
  readonly from: number;
  readonly to: number;
  readonly insert: string;
}

/** Changes of `expected`, the text they were computed against. */
export interface CodeTextChangesAgainst {
  readonly changes: readonly CodeTextChange[];
  readonly expected: string;
}

/**
 * `changes` sorted, checked to be disjoint and inside `base`, with no-ops
 * dropped; null when they do not describe edits of `base`.
 */
export function normalizeCodeTextChanges(
  base: string,
  changes: readonly CodeTextChange[],
): CodeTextChange[] | null {
  const sorted = [...changes].sort(
    (left, right) => left.from - right.from || left.to - right.to,
  );
  const out: CodeTextChange[] = [];
  let previousTo = 0;
  for (const change of sorted) {
    if (
      !Number.isSafeInteger(change.from) ||
      !Number.isSafeInteger(change.to) ||
      typeof change.insert !== "string" ||
      change.from < previousTo ||
      change.to < change.from ||
      change.to > base.length
    ) {
      return null;
    }
    previousTo = change.to;
    if (change.from === change.to && change.insert === "") continue;
    out.push({ from: change.from, to: change.to, insert: change.insert });
  }
  return out;
}

/** `base` with sorted, disjoint `changes` applied. */
export function applyCodeTextChanges(
  base: string,
  changes: readonly CodeTextChange[],
): string {
  let out = "";
  let at = 0;
  for (const change of changes) {
    out += base.slice(at, change.from) + change.insert;
    at = change.to;
  }
  return out + base.slice(at);
}

/**
 * The changes that take `applyCodeTextChanges(base, changes)` back to `base`,
 * each at its own place (the undo of a positioned edit is positioned too).
 */
export function invertCodeTextChanges(
  base: string,
  changes: readonly CodeTextChange[],
): CodeTextChange[] {
  const out: CodeTextChange[] = [];
  let shift = 0;
  for (const change of changes) {
    const from = change.from + shift;
    out.push({
      from,
      to: from + change.insert.length,
      insert: base.slice(change.from, change.to),
    });
    shift += change.insert.length - (change.to - change.from);
  }
  return out;
}

/**
 * Whether `change`'s inserted letters could equally stand one place earlier
 * or later in `base`: the letter before it is the run's last letter, or the
 * letter after it the run's first. Which copy the run follows is then a
 * guess. Conservative for a replacement (a deletion alone never splits a run).
 */
export function codeTextChangeIsAmbiguous(
  base: string,
  change: CodeTextChange,
): boolean {
  const { insert } = change;
  if (!insert) return false;
  return (
    (change.from > 0 && base[change.from - 1] === insert[insert.length - 1]) ||
    (change.to < base.length && base[change.to] === insert[0])
  );
}

/**
 * The hunks turning `from` into `to`, and whether the place of any of them is
 * a guess (`codeTextChangeIsAmbiguous`). A whole-text write with no positions
 * of its own is placed this way, never as one splice spanning the first to
 * the last difference (which re-created every character between them, so a
 * collaborator's key typed anywhere in that span lost its place).
 */
export function codeTextHunks(
  from: string,
  to: string,
): { hunks: CodeTextChange[]; ambiguous: boolean } {
  if (from === to) return { hunks: [], ambiguous: false };
  const hunks = diff(from, to, { scanLimit: 500, timeout: 20 }).map(
    (change) => ({
      from: change.fromA,
      to: change.toA,
      insert: to.slice(change.fromB, change.toB),
    }),
  );
  if (applyCodeTextChanges(from, hunks) !== to) {
    // Never expected from the diff; one splice is then still exact text.
    return {
      hunks: [{ from: 0, to: from.length, insert: to }],
      ambiguous: true,
    };
  }
  return {
    hunks,
    ambiguous: hunks.some((hunk) => codeTextChangeIsAmbiguous(from, hunk)),
  };
}

/**
 * Write sorted, disjoint `changes` (of `text`'s current string) into the
 * Y.Text at their own places, the last first so each index still holds.
 * The caller owns the transaction.
 */
export function writeCodeTextChanges(
  text: Y.Text,
  changes: readonly CodeTextChange[],
): boolean {
  for (let index = changes.length - 1; index >= 0; index -= 1) {
    const change = changes[index]!;
    if (change.to > change.from) text.delete(change.from, change.to - change.from);
    if (change.insert) text.insert(change.from, change.insert);
  }
  return changes.length > 0;
}
