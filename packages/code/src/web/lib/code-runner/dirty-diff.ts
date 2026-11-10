/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Change markers against a baseline (VS Code's `scm.diffDecorations`).
 *
 * VS Code marks every line you have changed since the last commit, and lets you
 * revert one of those changes without touching the rest. There is no Git here,
 * but there is the same question — "what have I changed since the last version
 * I kept?" — and the page's own saved snapshots answer it. This module is the
 * pure part: given a baseline and the live text, which lines changed, and what
 * exact edit puts one of them back.
 *
 * Everything is bounded. A very large document, a pathological diff, or a
 * rewrite that touches thousands of places all fail closed to "no markers"
 * rather than freezing the editor with decorations nobody asked for.
 */

import { diff, type Change, type DiffConfig } from "@codemirror/merge";

/** Documents larger than this are not marked; the diff cost is not worth it. */
export const DIRTY_DIFF_CODE_UNIT_LIMIT = 1024 * 1024;
/** Hunks rendered before the whole thing is treated as a rewrite. */
export const DIRTY_DIFF_HUNK_LIMIT = 500;
export const DIRTY_DIFF_SCAN_LIMIT = 50_000;
export const DIRTY_DIFF_TIMEOUT_MS = 100;

export type DirtyDiffKind = "added" | "modified" | "deleted";

export interface DirtyDiffHunk {
  kind: DirtyDiffKind;
  /**
   * 1-based inclusive line range in the CURRENT document. Markers are drawn per
   * line, so this is deliberately widened to whole lines.
   */
  fromLine: number;
  toLine: number;
  /**
   * EXACT character range the change occupies in the current document — not
   * widened. Reverting replaces precisely this, so an insertion in the middle
   * of a line does not take the rest of the line with it.
   */
  from: number;
  to: number;
  /** Baseline text this hunk replaced — what reverting writes back. */
  baselineText: string;
}

function lineStartOffsets(text: string): number[] {
  const starts = [0];
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) === 10) starts.push(index + 1);
  }
  return starts;
}

function lineIndexAt(starts: readonly number[], offset: number): number {
  let low = 0;
  let high = starts.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (starts[middle]! <= offset) low = middle + 1;
    else high = middle;
  }
  return Math.max(0, low - 1);
}

/** Grow an offset range to whole lines, so a marker covers the line it is on. */
function expandToLines(
  text: string,
  starts: readonly number[],
  from: number,
  to: number,
): { from: number; to: number; fromLine: number; toLine: number } {
  const safeFrom = Math.max(0, Math.min(from, text.length));
  const safeTo = Math.max(safeFrom, Math.min(to, text.length));
  const fromLine = lineIndexAt(starts, safeFrom);
  // A change ending exactly at a line start belongs to the previous line.
  const endLine =
    safeTo > safeFrom && starts[lineIndexAt(starts, safeTo)] === safeTo
      ? lineIndexAt(starts, safeTo - 1)
      : lineIndexAt(starts, safeTo);
  const lineFrom = starts[fromLine] ?? 0;
  const nextStart = starts[endLine + 1];
  const lineTo = nextStart === undefined ? text.length : nextStart - 1;
  return {
    from: lineFrom,
    to: Math.max(lineFrom, lineTo),
    fromLine: fromLine + 1,
    toLine: endLine + 1,
  };
}

/**
 * Which lines of `current` differ from `baseline`, and what to write back.
 *
 * Returns an empty list when there is nothing to compare against, when the
 * texts match, or when any bound is exceeded — an empty list always means "show
 * no markers", never "there are no changes".
 */
export function computeDirtyDiffHunks(
  baseline: string,
  current: string,
  options: {
    codeUnitLimit?: number;
    hunkLimit?: number;
    scanLimit?: number;
    timeoutMs?: number;
  } = {},
): DirtyDiffHunk[] {
  const codeUnitLimit = options.codeUnitLimit ?? DIRTY_DIFF_CODE_UNIT_LIMIT;
  const hunkLimit = options.hunkLimit ?? DIRTY_DIFF_HUNK_LIMIT;
  if (
    baseline === current ||
    baseline.length > codeUnitLimit ||
    current.length > codeUnitLimit
  ) {
    return [];
  }

  const config: DiffConfig = {
    scanLimit: options.scanLimit ?? DIRTY_DIFF_SCAN_LIMIT,
    timeout: options.timeoutMs ?? DIRTY_DIFF_TIMEOUT_MS,
  };
  let changes: readonly Change[];
  try {
    changes = diff(baseline, current, config);
  } catch {
    return [];
  }
  if (changes.length === 0 || changes.length > hunkLimit) return [];

  const currentStarts = lineStartOffsets(current);
  const hunks: DirtyDiffHunk[] = [];
  for (const change of changes) {
    const removed = change.toA > change.fromA;
    const added = change.toB > change.fromB;
    if (!removed && !added) continue;

    const currentRange = expandToLines(
      current,
      currentStarts,
      change.fromB,
      change.toB,
    );
    // Classify the way a line gutter is read, not the way a character diff is
    // computed. A line that already existed and now says something different is
    // MODIFIED, even when the edit only inserted characters into it; only a
    // change that carries a line break actually adds or removes lines. (A
    // character diff aligns on shared text, so "beta = 2" -> "beta = 222" is an
    // insertion — but nobody reads that as a new line.)
    const insertedText = current.slice(change.fromB, change.toB);
    const removedText = baseline.slice(change.fromA, change.toA);
    const kind: DirtyDiffKind =
      removed && added
        ? "modified"
        : added
          ? insertedText.includes("\n")
            ? "added"
            : "modified"
          : removedText.includes("\n")
            ? "deleted"
            : "modified";
    hunks.push({
      kind,
      fromLine: currentRange.fromLine,
      toLine: currentRange.toLine,
      // The exact span, so reverting rewrites what changed and nothing else.
      from: change.fromB,
      to: change.toB,
      baselineText: baseline.slice(change.fromA, change.toA),
    });
  }
  return hunks;
}

/**
 * The single edit that puts one hunk back to its baseline.
 *
 * Returns null when the hunk no longer fits the document it was computed
 * against — reverting stale offsets would corrupt unrelated text.
 */
export function revertDirtyDiffHunk(
  current: string,
  hunk: DirtyDiffHunk,
): { from: number; to: number; insert: string } | null {
  if (hunk.from < 0 || hunk.to < hunk.from || hunk.to > current.length) {
    return null;
  }
  // Every kind is the same operation: put the baseline text back where the
  // change is. A deletion simply has an empty range to put it back into.
  return { from: hunk.from, to: hunk.to, insert: hunk.baselineText };
}
