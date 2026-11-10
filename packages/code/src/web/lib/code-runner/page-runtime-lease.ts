/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reference-counted ownership of the runtimes that belong to a PAGE, not to a
 * window.
 *
 * A code page's Python interpreter, its in-memory SQL database and its console
 * realm are keyed by page id and shared by everything looking at that page.
 * The windows that use them are not: a code page opens as a `note`-type window,
 * and every note titlebar carries a "duplicate view" button that opens a second
 * window on the same page with `force_new`. So one page can have two views, and
 * before this module each view tore the page's runtimes down on unmount —
 * closing either one killed the other's interpreter mid-run, taking the user's
 * variables and tables with it. Nothing errored; the survivor simply started
 * answering as if it had just booted.
 *
 * The shape is the one Code-OSS uses for text models, which have the same
 * problem in a sharper form (one model, many editors): a resource is looked up
 * by its identity, the caller gets a handle rather than the resource, and the
 * resource is disposed when the last handle is released. Reimplemented here for
 * MRAI's page-scoped runtimes — no upstream code was read or transcribed, and
 * the naming deliberately follows this repo rather than theirs.
 *
 * Deliberate rules, each of which a caller can rely on:
 *
 *  - THE FIRST ACQUIRE'S DISPOSER WINS. Later acquires for the same key ignore
 *    the disposer they pass. Every holder of one page builds the same teardown,
 *    so they are identical by construction; honouring a later one would make
 *    what happens at disposal depend on the order the views happened to open.
 *
 *  - RELEASING TWICE IS A NO-OP, not a second decrement. React can run an
 *    effect cleanup more than once for one setup, and a double decrement would
 *    dispose a resource somebody is still holding — the exact bug this module
 *    exists to prevent, reintroduced from the other side.
 *
 *  - THE DISPOSER RUNS EXACTLY ONCE, after the entry is already gone from the
 *    registry. A disposer that throws therefore cannot leave a half-dead entry
 *    behind, and one that re-acquires its own key gets a clean first acquire.
 *
 * With a single holder — the ordinary case — acquire/release is exactly the
 * mount/unmount teardown it replaces, so the common path is unchanged.
 *
 * Framework-free: no React imports, fully unit-testable.
 */

/** A holder's handle on one keyed resource. Release it exactly once. */
export interface RuntimeLease {
  /** The resource identity this lease is against (a page id here). */
  readonly key: string;
  /**
   * Record a server content version learned through hydration/reload. This is
   * coordination metadata only; it never initiates a save or a fetch.
   */
  observeContentVersion(version: number): void;
  /**
   * Record a version produced by this view's autosave. Returns true exactly
   * once for that advancing version when another view of the page is alive, so
   * the caller can ask peers to reconcile their optimistic-save base.
   */
  publishSavedContentVersion(version: number): boolean;
  /** Give up this holder's claim. Idempotent; safe to call from a cleanup. */
  release(): void;
}

interface RuntimeEntry {
  holders: number;
  dispose: () => void;
  highestContentVersion: number | null;
}

const _entries = new Map<string, RuntimeEntry>();

/** The three Code surfaces that may exclusively execute against one page. */
export type PageExecutionKind = "run" | "console" | "debug";

/** Stable external-store snapshot for one page's current execution owner. */
export interface PageExecutionState {
  readonly ownerId: string;
  readonly kind: PageExecutionKind;
  /** Exact language captured when the execution was dispatched. */
  readonly language: string;
}

/**
 * Exclusive handle over one page's execution authority.
 *
 * A retired handle is deliberately inert: an older async completion can never
 * release a newer execution that acquired the same page after it.
 */
export interface PageExecutionClaim {
  readonly key: string;
  readonly state: PageExecutionState;
  isActive(): boolean;
  release(): void;
}

interface PageExecutionEntry {
  readonly token: number;
  readonly state: PageExecutionState;
}

const _pageExecutions = new Map<string, PageExecutionEntry>();
const _pageExecutionListeners = new Map<string, Set<() => void>>();
let _pageExecutionSequence = 0;

function notifyPageExecution(key: string): void {
  const listeners = _pageExecutionListeners.get(key);
  if (!listeners) return;
  for (const listener of [...listeners]) listener();
}

/**
 * Atomically claim the one execution slot for `key`.
 *
 * The registry is synchronous, so two views dispatching during one React frame
 * cannot both pass based on a stale rendered `disabled` state.
 */
export function tryAcquirePageExecution(
  key: string,
  state: PageExecutionState,
): PageExecutionClaim | null {
  if (_pageExecutions.has(key)) return null;

  const token = (_pageExecutionSequence += 1);
  const entry: PageExecutionEntry = {
    token,
    // Keep this exact object alive until release. `useSyncExternalStore`
    // requires referentially stable snapshots between notifications.
    state: { ...state },
  };
  _pageExecutions.set(key, entry);
  notifyPageExecution(key);

  let released = false;
  return {
    key,
    state: entry.state,
    isActive() {
      return !released && _pageExecutions.get(key)?.token === token;
    },
    release() {
      if (released) return;
      released = true;
      if (_pageExecutions.get(key)?.token !== token) return;
      _pageExecutions.delete(key);
      notifyPageExecution(key);
    },
  };
}

/** Stable snapshot of the page's current execution, or null while idle. */
export function getPageExecutionState(key: string): PageExecutionState | null {
  return _pageExecutions.get(key)?.state ?? null;
}

/** Subscribe a duplicate view to page execution ownership changes. */
export function subscribePageExecution(
  key: string,
  listener: () => void,
): () => void {
  let listeners = _pageExecutionListeners.get(key);
  if (!listeners) {
    listeners = new Set();
    _pageExecutionListeners.set(key, listeners);
  }
  listeners.add(listener);
  return () => {
    listeners?.delete(listener);
    if (listeners?.size === 0) _pageExecutionListeners.delete(key);
  };
}

/**
 * Claim the runtimes belonging to `key`, disposing them when the last holder
 * releases.
 *
 * `dispose` is used only if this is the first claim; see the module note.
 */
export function acquirePageRuntime(
  key: string,
  dispose: () => void,
): RuntimeLease {
  const existing = _entries.get(key);
  if (existing) {
    existing.holders += 1;
  } else {
    _entries.set(key, {
      holders: 1,
      dispose,
      highestContentVersion: null,
    });
  }

  let released = false;
  const recordVersion = (version: number): RuntimeEntry | null => {
    if (released || !Number.isSafeInteger(version) || version < 1) {
      return null;
    }
    const entry = _entries.get(key);
    if (!entry) return null;
    if (
      entry.highestContentVersion === null ||
      version > entry.highestContentVersion
    ) {
      entry.highestContentVersion = version;
      return entry;
    }
    return null;
  };
  return {
    key,
    observeContentVersion(version) {
      recordVersion(version);
    },
    publishSavedContentVersion(version) {
      const advanced = recordVersion(version);
      return advanced !== null && advanced.holders > 1;
    },
    release() {
      if (released) return;
      released = true;
      const entry = _entries.get(key);
      if (!entry) return;
      entry.holders -= 1;
      if (entry.holders > 0) return;
      // Remove BEFORE disposing, so a throwing or re-entrant disposer cannot
      // observe (or leave behind) an entry with no holders.
      _entries.delete(key);
      entry.dispose();
    },
  };
}

/**
 * How many holders `key` currently has; 0 when nothing holds it.
 *
 * For tests and diagnostics. Not a way to decide whether to dispose — that is
 * what releasing a lease is for.
 */
export function pageRuntimeHolderCount(key: string): number {
  return _entries.get(key)?.holders ?? 0;
}

/**
 * Drop every entry WITHOUT disposing. Tests only: disposing here would run
 * teardown against whatever mocks the finished test left behind.
 * @internal
 */
export function _clearPageRuntimeLeasesForTest(): void {
  _entries.clear();
  _pageExecutions.clear();
  _pageExecutionListeners.clear();
  _pageExecutionSequence = 0;
}
