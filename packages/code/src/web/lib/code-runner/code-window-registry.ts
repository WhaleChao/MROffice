/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Module-level registry of currently-open standalone code windows.
 *
 * Each {@link CodeWindow} registers itself on mount (with its page id, a
 * display label, and live accessors for its buffer) and deregisters on
 * unmount. Any other window — or the Find-in-Files / Compare panels — can
 * query the full set of open code buffers through {@link getOpenCodeBuffers}
 * without going through React context or the WindowManager.
 *
 * Why a module singleton rather than context?
 *  - CodeWindow is rendered deep in the window tree; threading context through
 *    every ancestor to reach all peer windows is impractical.
 *  - The registry is read-only from the outside (callers can't mutate the
 *    buffers directly; they call `setText` which is the window's own setter).
 *  - Tests can pre-populate the registry with stubs before mounting, making
 *    cross-window behaviour unit-testable without the full WindowManager.
 *
 * Framework-free: no React imports. Fully unit-testable.
 */

import type { CodeTextChangesAgainst } from "./code-text-changes";

export type CodeWindowProblemSeverity = "error" | "warning" | "info" | "hint";

/** One bounded lint problem exposed by an open code window. */
export interface CodeWindowProblem {
  from: number;
  to: number;
  severity: CodeWindowProblemSeverity;
  message: string;
  line: number;
  col: number;
}

export interface CodeWindowEntry {
  /** Notebook scope that owns this page; unscoped entries fail closed for scoped reads. */
  notebookId?: string;
  /** The page id of the code window (stable for the window's lifetime). */
  pageId: string;
  /** The window instance id (may be undefined for old callers). */
  windowId: string | undefined;
  /** Human label shown in the Find-in-Files results tree / Compare picker. */
  label: string;
  /** Returns the window's CURRENT buffer text (live, may change on every edit). */
  getText: () => string;
  /** Returns the window's current language. */
  getLanguage: () => string;
  /**
   * Overwrite the window's buffer with `text` and persist the change.
   * Used by Replace-in-Files to write replacements back into the right window.
   * The implementation is CodeWindow's own `setContent` + `persist` pair.
   * `positions`, when the caller knows where its edits are (the replaced
   * matches, a refactor's text changes): the shared text is written at those
   * places. Without them a change whose place is a guess may be refused in a
   * live room, and the window then shows the room's text again.
   */
  setText: (text: string, positions?: CodeTextChangesAgainst) => void;
  /** Returns this window's bounded, line-sorted diagnostic snapshot. */
  getProblems?: () => readonly CodeWindowProblem[];
  /** Returns the full live diagnostic count, including rows beyond the UI cap. */
  getProblemCount?: () => number;
  /** Focus this window's live editor without changing its current selection. */
  focus?: () => void;
  /** Select and reveal a diagnostic range in this window's live editor. */
  revealRange?: (from: number, to: number) => void;
}

/** One inert read of an open window's live accessors. */
export interface CodeWindowSnapshot {
  bufferId: string;
  notebookId: string | undefined;
  pageId: string;
  windowId: string | undefined;
  label: string;
  language: string;
  text: string;
  entry: CodeWindowEntry;
}

/** Stable opaque id used by cross-window search and code intelligence. */
export function codeWindowEntryId(
  entry: Pick<CodeWindowEntry, "pageId" | "windowId">,
): string {
  return `${entry.pageId}:::${entry.windowId ?? ""}`;
}

/** The single shared registry, keyed by a stable registration token. */
const _registry = new Map<string, CodeWindowEntry>();
const _listeners = new Set<(origin?: CodeWindowEntry) => void>();
const _sourceListeners = new Set<(bufferId?: string) => void>();
const _editListeners = new Set<(edit: CodeWindowEdit) => void>();
let _seq = 0;

/** Maximum diagnostic rows rendered across all open code windows. */
export const WORKSPACE_PROBLEM_DISPLAY_LIMIT = 500;
const MAX_PROBLEM_MESSAGE_LENGTH = 2_048;

export interface WorkspaceProblemGroup {
  pageId: string;
  windowId: string | undefined;
  label: string;
  problems: CodeWindowProblem[];
  totalCount: number;
  revealRange?: (from: number, to: number) => void;
}

export interface WorkspaceProblemSnapshot {
  groups: WorkspaceProblemGroup[];
  /** Full diagnostic count across windows, even when displayed rows are capped. */
  totalCount: number;
  displayedCount: number;
  truncated: boolean;
}

/** `origin` is the window whose registration or diagnostics changed, when
 * known: a listener refreshes its own window at once and its peers' changes
 * as background work. */
function emitRegistryChange(origin?: CodeWindowEntry): void {
  for (const listener of [..._listeners]) {
    listener(origin);
  }
}

function emitSourceChange(bufferId?: string): void {
  for (const listener of [..._sourceListeners]) {
    listener(bufferId);
  }
}

/**
 * Register a code window. Returns a unique token that must be passed to
 * {@link unregisterCodeWindow} when the window unmounts.
 */
export function registerCodeWindow(entry: CodeWindowEntry): string {
  _seq += 1;
  const token = `cw-${entry.pageId}-${_seq}`;
  _registry.set(token, entry);
  emitRegistryChange(entry);
  emitSourceChange(codeWindowEntryId(entry));
  return token;
}

/** Deregister a previously-registered code window by its token. */
export function unregisterCodeWindow(token: string): void {
  const entry = _registry.get(token);
  if (_registry.delete(token)) {
    emitRegistryChange(entry);
    emitSourceChange(entry ? codeWindowEntryId(entry) : undefined);
  }
}

/** Subscribe to registry membership or live diagnostic-snapshot changes. */
export function subscribeCodeWindowRegistry(
  listener: (origin?: CodeWindowEntry) => void,
): () => void {
  _listeners.add(listener);
  return () => {
    _listeners.delete(listener);
  };
}

/**
 * Subscribe only to open-buffer membership/source/language changes.
 *
 * Kept separate from the general registry channel because diagnostics publish
 * through that channel; using it to refresh project diagnostics would form a
 * lint -> Problems -> registry -> lint loop.
 */
export function subscribeCodeWindowSources(
  listener: (bufferId?: string) => void,
): () => void {
  _sourceListeners.add(listener);
  return () => {
    _sourceListeners.delete(listener);
  };
}

/** Notify peer editors that one open buffer's source or language changed. */
export function notifyCodeWindowSourceChange(bufferId: string): void {
  emitSourceChange(bufferId);
}

/** A member's own edit in a code window (typing, not a load or a peer's save). */
export interface CodeWindowEdit {
  readonly notebookId: string | undefined;
  readonly pageId: string;
  /** The file's workspace path ("paper/chapters/intro.tex"). */
  readonly path: string;
}

/**
 * Subscribe to members' edits in every open code window. Unlike the source
 * channel, a page loading, a window opening or a save acknowledgement is not
 * an edit, so a listener can act on it (a LaTeX main file compiling when a
 * file of its project is typed in) without reacting to its own refresh.
 */
export function subscribeCodeWindowEdits(
  listener: (edit: CodeWindowEdit) => void,
): () => void {
  _editListeners.add(listener);
  return () => {
    _editListeners.delete(listener);
  };
}

/** Announce an edit the member made in a code window. */
export function notifyCodeWindowEdit(edit: CodeWindowEdit): void {
  for (const listener of [..._editListeners]) {
    listener(edit);
  }
}

/** Notify subscribers after a registered window publishes new diagnostics. */
export function notifyCodeWindowRegistryChange(origin?: CodeWindowEntry): void {
  emitRegistryChange(origin);
}

/**
 * Return every currently-open code window's buffer descriptor, optionally
 * excluding the caller's own window (by page id + window id). Order is
 * registration order (i.e. open order).
 */
export function getOpenCodeBuffers(
  excludePageId?: string,
  excludeWindowId?: string,
): CodeWindowEntry[] {
  const out: CodeWindowEntry[] = [];
  for (const entry of _registry.values()) {
    if (
      excludePageId !== undefined &&
      entry.pageId === excludePageId &&
      (excludeWindowId === undefined || entry.windowId === excludeWindowId)
    ) {
      continue;
    }
    out.push(entry);
  }
  return out;
}

/**
 * Read open-window metadata without letting one detached view break every
 * registry consumer. Write operations still use the live entry and recheck
 * their own authority immediately before mutation.
 */
export function getOpenCodeWindowSnapshots(
  excludePageId?: string,
  excludeWindowId?: string,
  notebookId?: string,
): CodeWindowSnapshot[] {
  const snapshots: CodeWindowSnapshot[] = [];
  for (const entry of getOpenCodeBuffers(excludePageId, excludeWindowId)) {
    try {
      if (
        notebookId !== undefined &&
        entry.notebookId !== notebookId
      ) {
        continue;
      }
      const label = entry.label;
      const language = entry.getLanguage();
      const text = entry.getText();
      if (
        typeof label !== "string" ||
        typeof language !== "string" ||
        typeof text !== "string"
      ) {
        continue;
      }
      snapshots.push({
        bufferId: codeWindowEntryId(entry),
        notebookId: entry.notebookId,
        pageId: entry.pageId,
        windowId: entry.windowId,
        label,
        language,
        text,
        entry,
      });
    } catch {
      // A stale/unmounted view is absent from this snapshot, not fatal.
    }
  }
  return snapshots;
}

function safeProblemCount(value: unknown, fallback: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return fallback;
  return Math.max(
    fallback,
    Math.max(0, Math.min(Number.MAX_SAFE_INTEGER, Math.trunc(value))),
  );
}

function normalizeProblem(value: CodeWindowProblem): CodeWindowProblem | null {
  if (
    !value ||
    typeof value.from !== "number" ||
    !Number.isFinite(value.from) ||
    typeof value.to !== "number" ||
    !Number.isFinite(value.to) ||
    typeof value.line !== "number" ||
    !Number.isFinite(value.line) ||
    typeof value.col !== "number" ||
    !Number.isFinite(value.col) ||
    typeof value.message !== "string" ||
    !["error", "warning", "info", "hint"].includes(value.severity)
  ) {
    return null;
  }
  const from = Math.max(0, Math.trunc(value.from));
  return {
    from,
    to: Math.max(from, Math.trunc(value.to)),
    severity: value.severity,
    message: value.message.slice(0, MAX_PROBLEM_MESSAGE_LENGTH),
    line: Math.max(1, Math.trunc(value.line)),
    col: Math.max(1, Math.trunc(value.col)),
  };
}

interface WorkspaceProblemSource {
  entry: CodeWindowEntry;
  problems: readonly CodeWindowProblem[];
  totalCount: number;
}

function readWorkspaceProblemSource(
  entry: CodeWindowEntry,
): WorkspaceProblemSource | null {
  let problems: readonly CodeWindowProblem[];
  try {
    const rawProblems: unknown = entry.getProblems?.() ?? [];
    if (!Array.isArray(rawProblems)) return null;
    // Reading length can itself fail for a revoked Proxy-backed provider.
    void rawProblems.length;
    problems = rawProblems as readonly CodeWindowProblem[];
  } catch {
    return null;
  }

  let reportedCount: unknown;
  try {
    reportedCount = entry.getProblemCount?.();
  } catch {
    reportedCount = undefined;
  }
  return {
    entry,
    problems,
    totalCount: safeProblemCount(reportedCount, problems.length),
  };
}

/**
 * Project live per-window diagnostics into one bounded workspace snapshot.
 *
 * Windows remain in open order and rows remain in their editor-provided order.
 * Only `limit * 4` candidate rows are inspected, so a hostile/legacy provider
 * returning a huge invalid array cannot turn the toolbar into an unbounded scan.
 */
export function collectWorkspaceProblems(
  entries: readonly CodeWindowEntry[] = getOpenCodeBuffers(),
  limit = WORKSPACE_PROBLEM_DISPLAY_LIMIT,
): WorkspaceProblemSnapshot {
  const safeLimit = Math.max(
    0,
    Math.min(WORKSPACE_PROBLEM_DISPLAY_LIMIT, Math.trunc(limit) || 0),
  );
  const groups: WorkspaceProblemGroup[] = [];
  let totalCount = 0;
  let displayedCount = 0;
  let inspectedCount = 0;
  const inspectLimit = safeLimit * 4;

  // Duplicate View creates multiple registry entries for one canonical page.
  // Count that page once, preferring the most complete live diagnostic view.
  // A broken/revoked provider is isolated; a healthy duplicate can still own
  // the page's Problems group.
  const sources: WorkspaceProblemSource[] = [];
  const sourceIndexByPageId = new Map<string, number>();
  for (const entry of entries) {
    const candidate = readWorkspaceProblemSource(entry);
    if (!candidate) continue;
    const existingIndex = sourceIndexByPageId.get(entry.pageId);
    if (existingIndex === undefined) {
      sourceIndexByPageId.set(entry.pageId, sources.length);
      sources.push(candidate);
      continue;
    }
    const existing = sources[existingIndex];
    if (
      candidate.totalCount > existing.totalCount ||
      (candidate.totalCount === existing.totalCount &&
        candidate.problems.length > existing.problems.length)
    ) {
      sources[existingIndex] = candidate;
    }
  }

  for (const source of sources) {
    const { entry, problems: sourceProblems, totalCount: entryTotal } = source;
    totalCount = Math.min(Number.MAX_SAFE_INTEGER, totalCount + entryTotal);

    const problems: CodeWindowProblem[] = [];
    for (
      let index = 0;
      index < sourceProblems.length &&
      displayedCount < safeLimit &&
      inspectedCount < inspectLimit;
      index += 1
    ) {
      inspectedCount += 1;
      let rawProblem: CodeWindowProblem;
      try {
        rawProblem = sourceProblems[index];
      } catch {
        continue;
      }
      const problem = normalizeProblem(rawProblem);
      if (!problem) continue;
      problems.push(problem);
      displayedCount += 1;
    }
    if (problems.length > 0) {
      groups.push({
        pageId: entry.pageId,
        windowId: entry.windowId,
        label: entry.label,
        problems,
        totalCount: entryTotal,
        revealRange: entry.revealRange,
      });
    }
  }

  return {
    groups,
    totalCount,
    displayedCount,
    truncated: totalCount > displayedCount,
  };
}

/**
 * Clear the registry. Only used in tests to reset state between test cases.
 * @internal
 */
export function _clearRegistryForTest(): void {
  _registry.clear();
  _listeners.clear();
  _sourceListeners.clear();
  _editListeners.clear();
}
