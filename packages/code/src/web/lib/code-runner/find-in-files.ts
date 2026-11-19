/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for a VSCode-style "Find in Files" (Ctrl+Shift+F) panel over a set
 * of named buffers — the workspace's code cells / open documents.
 *
 * Where {@link crossCellMatches} flattens every cell's match into ONE ordered
 * navigation list (next/prev across boundaries), this module instead GROUPS the
 * matches the way the Find-in-Files results tree does: one group per buffer that
 * has at least one hit, and within each group one entry per match carrying its
 * 1-based line / column, its offset range, and a one-line snippet for display.
 * Clicking an entry navigates to that buffer at the match's offset (the caller
 * wires the result of {@link findInFilesHitTarget} to its existing per-cell /
 * per-document selection path).
 *
 * It reuses the SAME matcher as the per-cell find bar and the cross-cell list
 * ({@link findMatches} from find-replace.ts), so the literal / regex / case /
 * whole-word semantics — and therefore the hit count — agree exactly across all
 * three surfaces. Buffers with no hit are omitted (an empty results tree), and
 * an empty query yields no groups.
 *
 * Framework-free + fully unit-tested; panel UIs iterate the groups and render
 * them. Replacement and its bounded undo use these same loaded buffer snapshots.
 * A backend ripgrep search over unauthorized host files remains deliberately out
 * of scope — callers may only supply buffers already loaded through their
 * notebook/workspace authorization path.
 */

import {
  findMatches,
  FindReplacementWorkspaceTooLargeError,
  MAX_FIND_REPLACEMENT_MATCHES,
  replaceAllWithEdits,
  type FindOptions,
  type Match,
} from "./find-replace";
import {
  invertCodeTextChanges,
  type CodeTextChange,
} from "./code-text-changes";

/** A buffer to search: a code cell or an open document, with a display label. */
export interface SearchBuffer {
  /** Stable id used to navigate back to this buffer (cell id / document id). */
  id: string;
  /** Human label shown as the group header (e.g. "Cell 2", a filename). */
  label: string;
  /** The buffer's full current text. */
  text: string;
}

/** A single match located within one buffer, with display + navigation data. */
export interface FindInFilesHit {
  /** Id of the buffer this hit lives in (mirrors {@link SearchBuffer.id}). */
  bufferId: string;
  /** 1-based line number of the match start within that buffer. */
  line: number;
  /** 1-based column (UTF-16 code units) of the match start within its line. */
  column: number;
  /** Start offset of the match within the buffer's text. */
  start: number;
  /** End offset of the match within the buffer's text. */
  end: number;
  /** A bounded preview of the source line the match starts on. */
  lineText: string;
  /** Column inside a bounded preview when `lineText` was shortened. */
  previewColumn?: number;
  /** The exact matched substring (for highlighting within the snippet). */
  matchText: string;
}

/** One buffer's worth of hits — a group in the results tree. Never empty. */
export interface FindInFilesGroup {
  /** Id of the buffer this group represents. */
  bufferId: string;
  /** Display label for the group header. */
  label: string;
  /** The hits within this buffer, in left-to-right / top-to-bottom order. */
  hits: FindInFilesHit[];
}

/** The whole search result: the non-empty groups plus the total hit count. */
export interface FindInFilesResult {
  groups: FindInFilesGroup[];
  /** Displayed hits across every group (== sum of group.hits.length). */
  total: number;
  /** True when more hits exist beyond the bounded workspace result window. */
  truncated?: boolean;
}

export const FIND_IN_FILES_RESULT_LIMIT = 2_000;
export const FIND_IN_FILES_SNIPPET_LIMIT = 400;
const FIND_IN_FILES_MATCH_TEXT_LIMIT = 512;

/**
 * Build a locator for the already ordered matches in one buffer. It advances
 * through line breaks once and retains only the current line, so a file with
 * millions of short lines does not allocate a whole-file line-start array just
 * to display a bounded result window.
 */
function createOrderedMatchLocator(
  text: string,
): (match: Match) => {
  line: number;
  column: number;
  lineText: string;
  previewColumn?: number;
} {
  let line = 1;
  let lineStart = 0;
  let nextLineBreak = text.indexOf("\n");
  return (match) => {
    while (nextLineBreak !== -1 && nextLineBreak < match.start) {
      line += 1;
      lineStart = nextLineBreak + 1;
      nextLineBreak = text.indexOf("\n", lineStart);
    }
    const lineEnd = nextLineBreak === -1 ? text.length : nextLineBreak;
    let lineText = text.slice(lineStart, lineEnd);
    if (lineText.endsWith("\r")) lineText = lineText.slice(0, -1);
    const offsetInLine = match.start - lineStart;
    const column = offsetInLine + 1;
    if (lineText.length <= FIND_IN_FILES_SNIPPET_LIMIT) {
      return { line, column, lineText };
    }

    const contextBefore = Math.floor(FIND_IN_FILES_SNIPPET_LIMIT * 0.3);
    let previewStart = Math.max(0, offsetInLine - contextBefore);
    let previewEnd = Math.min(
      lineText.length,
      previewStart + FIND_IN_FILES_SNIPPET_LIMIT,
    );
    previewStart = Math.max(
      0,
      previewEnd - FIND_IN_FILES_SNIPPET_LIMIT,
    );
    previewEnd = Math.min(
      lineText.length,
      previewStart + FIND_IN_FILES_SNIPPET_LIMIT,
    );
    const hasPrefix = previewStart > 0;
    const hasSuffix = previewEnd < lineText.length;
    lineText = `${hasPrefix ? "…" : ""}${lineText.slice(
      previewStart,
      previewEnd,
    )}${hasSuffix ? "…" : ""}`;
    return {
      line,
      column,
      lineText,
      previewColumn: offsetInLine - previewStart + (hasPrefix ? 2 : 1),
    };
  };
}

/**
 * Group every match of `query` across `buffers` into a results tree. Buffers are
 * visited in order; within each buffer matches are left-to-right ({@link
 * findMatches} order). A buffer with no match is omitted entirely, so `groups`
 * contains only buffers that actually matched. An empty query (or no buffers)
 * yields `{ groups: [], total: 0 }`.
 */
export function findInFiles(
  buffers: readonly SearchBuffer[],
  query: string,
  opts: FindOptions = {},
  limit = FIND_IN_FILES_RESULT_LIMIT,
): FindInFilesResult {
  if (!query) return { groups: [], total: 0 };
  const safeLimit = Number.isFinite(limit)
    ? Math.max(0, Math.min(FIND_IN_FILES_RESULT_LIMIT, Math.trunc(limit)))
    : FIND_IN_FILES_RESULT_LIMIT;
  const groups: FindInFilesGroup[] = [];
  let total = 0;
  let truncated = false;
  for (const buffer of buffers) {
    const text = buffer.text ?? "";
    const remaining = Math.max(0, safeLimit - total);
    const matches = findMatches(text, query, opts, remaining + 1);
    if (matches.length === 0) continue;
    if (remaining === 0) {
      truncated = true;
      break;
    }
    if (matches.length > remaining) truncated = true;
    const visibleMatches = matches.slice(0, remaining);
    const locateMatch = createOrderedMatchLocator(text);
    const hits: FindInFilesHit[] = visibleMatches.map((m) => {
      const { line, column, lineText, previewColumn } = locateMatch(m);
      return {
        bufferId: buffer.id,
        line,
        column,
        start: m.start,
        end: m.end,
        lineText,
        ...(previewColumn === undefined ? {} : { previewColumn }),
        matchText: text.slice(
          m.start,
          Math.min(m.end, m.start + FIND_IN_FILES_MATCH_TEXT_LIMIT),
        ),
      };
    });
    groups.push({ bufferId: buffer.id, label: buffer.label, hits });
    total += hits.length;
    if (truncated) break;
  }
  return truncated ? { groups, total, truncated: true } : { groups, total };
}

/**
 * The navigation target for a clicked hit: the buffer id and the offset range to
 * select. The panel hands this straight to its existing per-buffer selection
 * path (e.g. `selectCellRange`), so opening a result reuses the same code as
 * cross-cell go-to-definition / find navigation.
 */
export function findInFilesHitTarget(
  hit: FindInFilesHit,
): { bufferId: string; from: number; to: number } {
  return { bufferId: hit.bufferId, from: hit.start, to: hit.end };
}

// ---------------------------------------------------------------------------
// Replace-in-Files
// ---------------------------------------------------------------------------

/** The outcome of a replace-in-files operation — which buffers changed. */
export interface ReplaceInFilesResult {
  /** Map of bufferId -> new text for every buffer that was modified. */
  changes: Map<string, string>;
  /** Total number of replacements made across all buffers. */
  count: number;
  /** Map of bufferId -> the replaced matches of its text (sorted, disjoint),
   * so a shared text is written at the matches' places, not re-found by
   * comparing the old and new text (code-text-changes). */
  edits?: Map<string, CodeTextChange[]>;
}

export const FIND_IN_FILES_UNDO_CODE_UNIT_LIMIT = 32 * 1024 * 1024;
export const FIND_IN_FILES_REPLACEMENT_RESULT_CODE_UNIT_LIMIT =
  FIND_IN_FILES_UNDO_CODE_UNIT_LIMIT;

export interface ReplaceInFilesUndoEntry {
  /** Exact text before the workspace replacement. */
  before: string;
  /** Exact text written by the workspace replacement. */
  after: string;
  /** The replacement's edits of `before`, when known. */
  edits?: readonly CodeTextChange[];
}

export interface ReplaceInFilesUndoSnapshot {
  entries: ReadonlyMap<string, ReplaceInFilesUndoEntry>;
}

export interface UndoReplaceInFilesResult {
  /** Original text to write back, keyed by the still-open buffer id. */
  changes: Map<string, string>;
  /** The edits of each restored buffer's live text that write it back, at
   * the replaced matches' places, when the replacement's edits are known. */
  edits?: Map<string, CodeTextChange[]>;
  /** Buffers restored because their text still matched the replacement result. */
  restoredBufferCount: number;
  /** Closed or subsequently edited buffers deliberately left untouched. */
  skippedBufferCount: number;
}

/**
 * Capture one bounded workspace-replace transaction for a single-step Undo.
 *
 * Strings are immutable references, so retaining `after` does not duplicate
 * the editor's live value. The budget counts only displaced pre-replacement
 * text, which is the additional memory the undo record keeps alive.
 */
export function createReplaceInFilesUndoSnapshot(
  buffers: readonly SearchBuffer[],
  result: ReplaceInFilesResult,
  maxCodeUnits = FIND_IN_FILES_UNDO_CODE_UNIT_LIMIT,
): ReplaceInFilesUndoSnapshot | null {
  const safeLimit = Number.isFinite(maxCodeUnits)
    ? Math.max(0, Math.trunc(maxCodeUnits))
    : FIND_IN_FILES_UNDO_CODE_UNIT_LIMIT;
  const entries = new Map<string, ReplaceInFilesUndoEntry>();
  let retainedCodeUnits = 0;
  for (const buffer of buffers) {
    const after = result.changes.get(buffer.id);
    if (after === undefined || after === buffer.text) continue;
    retainedCodeUnits += buffer.text.length;
    if (retainedCodeUnits > safeLimit) return null;
    const edits = result.edits?.get(buffer.id);
    entries.set(
      buffer.id,
      edits ? { before: buffer.text, after, edits } : { before: buffer.text, after },
    );
  }
  return entries.size > 0 ? { entries } : null;
}

/**
 * Resolve a workspace replacement undo against live buffers.
 *
 * A buffer is restored only when it still contains the exact text written by
 * that transaction. This avoids overwriting typing or another automated edit
 * that happened after Replace All. Closed buffers are counted as skipped.
 */
export function undoReplaceInFiles(
  buffers: readonly SearchBuffer[],
  snapshot: ReplaceInFilesUndoSnapshot,
): UndoReplaceInFilesResult {
  const liveById = new Map(buffers.map((buffer) => [buffer.id, buffer.text]));
  const changes = new Map<string, string>();
  const edits = new Map<string, CodeTextChange[]>();
  let skippedBufferCount = 0;
  for (const [bufferId, entry] of snapshot.entries) {
    if (liveById.get(bufferId) === entry.after) {
      changes.set(bufferId, entry.before);
      if (entry.edits) {
        edits.set(bufferId, invertCodeTextChanges(entry.before, entry.edits));
      }
    } else {
      skippedBufferCount += 1;
    }
  }
  return {
    changes,
    edits,
    restoredBufferCount: changes.size,
    skippedBufferCount,
  };
}

/**
 * Apply a literal replacement to all matches of {@link query} across
 * {@link buffers}, returning a map of changed buffer ids to their new texts.
 * Buffers with no match (or an empty query) are omitted from the result —
 * the caller only needs to write back the changed buffers.
 *
 * Uses the SAME {@link findMatches} matcher as {@link findInFiles}, so the
 * options (case / whole-word / regex) and hit count agree exactly.
 */
export function replaceInFiles(
  buffers: readonly SearchBuffer[],
  query: string,
  replacement: string,
  opts: FindOptions = {},
  maxChangedCodeUnits = FIND_IN_FILES_REPLACEMENT_RESULT_CODE_UNIT_LIMIT,
): ReplaceInFilesResult {
  if (!query) return { changes: new Map(), count: 0, edits: new Map() };
  const safeResultLimit = Number.isFinite(maxChangedCodeUnits)
    ? Math.max(
        0,
        Math.min(
          FIND_IN_FILES_REPLACEMENT_RESULT_CODE_UNIT_LIMIT,
          Math.trunc(maxChangedCodeUnits),
        ),
      )
    : FIND_IN_FILES_REPLACEMENT_RESULT_CODE_UNIT_LIMIT;
  const changes = new Map<string, string>();
  const edits = new Map<string, CodeTextChange[]>();
  let count = 0;
  let changedCodeUnits = 0;
  for (const buffer of buffers) {
    const text = buffer.text ?? "";
    const result = replaceAllWithEdits(
      text,
      query,
      replacement,
      opts,
      MAX_FIND_REPLACEMENT_MATCHES - count,
    );
    if (result.count === 0) continue;
    count += result.count;
    changedCodeUnits += result.text.length;
    if (changedCodeUnits > safeResultLimit) {
      throw new FindReplacementWorkspaceTooLargeError();
    }
    changes.set(buffer.id, result.text);
    edits.set(buffer.id, result.edits);
  }
  return { changes, count, edits };
}
