/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure cursor NAVIGATION HISTORY (VSCode parity — the "Go Back" /
 * "Go Forward" navigation stack, `workbench.action.navigateBack` /
 * `...navigateForward`).
 *
 * VSCode records a navigation point whenever the caret JUMPS non-incrementally —
 * a go-to-definition, a Go-to-Symbol pick, a reference jump, a Go-to-Line — so
 * the user can hop "back" to where they were before the jump and "forward"
 * again. This module is the framework-free stack behind that: CodeMirrorEditor
 * captures the caret offset/selection at each jump boundary and feeds it here;
 * Back / Forward return the position to restore.
 *
 * The model is exactly a BROWSER history list:
 *  - `entries` is the visited trail, oldest → newest;
 *  - `index` points at the CURRENT position within it (`-1` when empty);
 *  - a NEW navigation truncates any "forward" entries past `index` (you can't go
 *    forward to a branch you've abandoned) and appends, just like opening a new
 *    page after pressing Back drops the old forward history;
 *  - Back walks `index` toward 0; Forward walks it toward the end.
 *
 * Each navigation records BOTH endpoints (the origin you LEFT and the
 * destination you arrived at), so the very first jump from A→B yields the trail
 * `[A, B]` with the cursor "at" B — Back then restores A, Forward re-restores B.
 * Consecutive endpoints that coincide are de-duplicated so a jump whose origin
 * already sits on the current entry doesn't double it.
 *
 * Everything here is total (never throws), immutable (each operation returns a
 * fresh state — nothing is mutated in place, so React state updates stay safe),
 * and capped so a long editing session can't grow the trail without bound.
 */

/**
 * A recorded caret position: the selection's `anchor` and `head` offsets (so a
 * jump that SELECTED a span — e.g. go-to-definition lands on the whole
 * identifier — is restored as a selection, not just a bare cursor). For a plain
 * cursor, `anchor === head`. Offsets are 0-based document offsets; the editor
 * clamps them into the live document on restore (a buffer edit between record
 * and restore can shift them).
 */
export interface NavPosition {
  anchor: number;
  head: number;
}

/**
 * The navigation-history state. `entries[index]` is the position the caret is
 * currently "at" within the history (after the most recent record / Back /
 * Forward). `index` is `-1` exactly when `entries` is empty.
 */
export interface NavHistory {
  entries: NavPosition[];
  index: number;
  /** Max trail length; older entries are evicted from the FRONT past this. */
  readonly limit: number;
}

/** The default trail cap — generous for a session, bounded against runaway growth. */
export const NAV_HISTORY_LIMIT = 50;

/** A fresh, empty history. `limit` caps the trail (defaults to {@link NAV_HISTORY_LIMIT}). */
export function createNavHistory(limit: number = NAV_HISTORY_LIMIT): NavHistory {
  // A non-finite / non-positive limit would break the cap math; floor to >= 1.
  const cap = Number.isFinite(limit) && limit >= 1 ? Math.floor(limit) : NAV_HISTORY_LIMIT;
  return { entries: [], index: -1, limit: cap };
}

/** Whether two positions are the same caret/selection (used to de-dupe the trail). */
function samePosition(a: NavPosition, b: NavPosition): boolean {
  return a.anchor === b.anchor && a.head === b.head;
}

/** Normalize an offset to a finite, non-negative integer (defensive). */
function normOffset(n: number): number {
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

/** Normalize a position's offsets (defensive against NaN / negative inputs). */
function normPosition(pos: NavPosition): NavPosition {
  return { anchor: normOffset(pos.anchor), head: normOffset(pos.head) };
}

/**
 * Append `pos` as the new newest entry, dropping any forward history past the
 * current `index`, de-duplicating against the entry that would precede it, and
 * enforcing the front-eviction cap. `index` ends pointing at the appended (or
 * coalesced) entry. Internal helper shared by {@link recordNavigation}.
 */
function appendEntry(history: NavHistory, pos: NavPosition): NavHistory {
  // Drop forward entries: a new navigation invalidates anything you'd "Forward"
  // into (browser semantics). `index + 1` keeps entries[0..index] inclusive.
  const kept = history.entries.slice(0, history.index + 1);
  const last = kept.length > 0 ? kept[kept.length - 1] : null;
  // Coalesce: don't record the same position twice in a row.
  if (last && samePosition(last, pos)) {
    return { ...history, entries: kept, index: kept.length - 1 };
  }
  kept.push(pos);
  // Front-evict to the cap, keeping the cursor anchored at the newest entry.
  const overflow = kept.length - history.limit;
  const entries = overflow > 0 ? kept.slice(overflow) : kept;
  return { ...history, entries, index: entries.length - 1 };
}

/**
 * Record a navigation JUMP from `from` (where the caret was) to `to` (where it
 * landed). Both endpoints are added to the trail — so the first jump A→B yields
 * `[A, B]` with `index` at B — with consecutive duplicates coalesced. Any
 * forward history is dropped (a new jump abandons the old forward branch). The
 * caret ends "at" `to`.
 *
 * Call this at every non-incremental caret jump (go-to-definition, symbol pick,
 * reference jump, Go-to-Line) so Back returns to the pre-jump spot. It is NOT
 * called for ordinary typing / arrow-key motion (that would flood the trail);
 * the editor only feeds it the genuine navigation boundaries, matching VSCode.
 */
export function recordNavigation(
  history: NavHistory,
  from: NavPosition,
  to: NavPosition,
): NavHistory {
  // Record the origin first (so Back can return to it), then the destination.
  const afterFrom = appendEntry(history, normPosition(from));
  return appendEntry(afterFrom, normPosition(to));
}

/** Whether Back can move (there is an older entry than the current one). */
export function canGoBack(history: NavHistory): boolean {
  return history.index > 0;
}

/** Whether Forward can move (there is a newer entry than the current one). */
export function canGoForward(history: NavHistory): boolean {
  return history.index >= 0 && history.index < history.entries.length - 1;
}

/**
 * Step BACK one entry. Returns the updated history plus the {@link NavPosition}
 * to restore the caret to, or `null` when already at the oldest entry (the
 * editor then no-ops, leaving the caret put). Pure: the input history is
 * unchanged.
 */
export function goBack(
  history: NavHistory,
): { history: NavHistory; position: NavPosition } | null {
  if (!canGoBack(history)) return null;
  const index = history.index - 1;
  return { history: { ...history, index }, position: history.entries[index] };
}

/**
 * Step FORWARD one entry. Returns the updated history plus the
 * {@link NavPosition} to restore, or `null` when already at the newest entry.
 * Pure: the input history is unchanged.
 */
export function goForward(
  history: NavHistory,
): { history: NavHistory; position: NavPosition } | null {
  if (!canGoForward(history)) return null;
  const index = history.index + 1;
  return { history: { ...history, index }, position: history.entries[index] };
}
