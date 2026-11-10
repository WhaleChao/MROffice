/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * When a LaTeX page compiles without being asked, as Overleaf's "Auto
 * compile" does, and what a compile shortcut pressed during a compile means.
 *
 * Overleaf compiles once typing has paused (its compiler's debounce is
 * 2.5 s), never starts a second compile while one runs (it tries again after),
 * and keeps the switch per member. Every compile here is a run on the
 * server's sandbox and counts as one, so the scheduler compiles only after an
 * edit, only when the text differs from what the last compile read, and at
 * most one compile follows the running one. A compile that fails is not
 * retried: the next edit is what asks again.
 *
 * The scheduler is framework-free; the Code window supplies what it knows
 * (the switch, whether a compile may start, the text a compile would read)
 * and starts the compile itself.
 */

import { editorPrefStorageKey, loadEditorPref, saveEditorPref } from "./editor-prefs";

/** Overleaf's pause before an automatic compile (AUTO_COMPILE_DEBOUNCE). */
export const LATEX_AUTO_COMPILE_DELAY_MS = 2_500;
/** How soon a compile held back by a save in flight looks again. */
export const LATEX_COMPILE_RETRY_MS = 500;
/** How many times it looks again (10 s) before waiting for the next edit. */
export const LATEX_COMPILE_MAX_RETRIES = 20;

// --- The switch ---------------------------------------------------------------

let cachedEnabled: boolean | null = null;
const enabledListeners = new Set<() => void>();

/** Whether Auto compile is on for this member in this browser (default off). */
export function latexAutoCompileEnabled(): boolean {
  if (cachedEnabled === null) cachedEnabled = loadEditorPref("latexAutoCompile", false);
  return cachedEnabled;
}

export function setLatexAutoCompileEnabled(enabled: boolean): void {
  cachedEnabled = enabled;
  saveEditorPref("latexAutoCompile", enabled);
  for (const listener of [...enabledListeners]) listener();
}

/** Every Code window follows the switch, in this tab and the member's others. */
export function subscribeLatexAutoCompile(listener: () => void): () => void {
  enabledListeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key !== editorPrefStorageKey("latexAutoCompile")) return;
    cachedEnabled = loadEditorPref("latexAutoCompile", false);
    listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    enabledListeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

export function resetLatexAutoCompileForTests(): void {
  cachedEnabled = null;
  enabledListeners.clear();
}

// --- Which file compiles ------------------------------------------------------

const ROOT_DOCUMENT_LINE = /^[^%\n]*\\document(?:class|style)\b/m;
/** \documentclass sits in the preamble; a long file is not read past this. */
const ROOT_DOCUMENT_SCAN_CHARS = 64_000;

/**
 * Whether `source` is a document of its own (it has a \documentclass outside
 * a comment), not a chapter or a bibliography that another file reads. Only
 * such a file compiles by itself: a chapter's edit compiles the main file
 * whose project it is in.
 */
export function isLatexRootDocument(source: string): boolean {
  return ROOT_DOCUMENT_LINE.test(source.slice(0, ROOT_DOCUMENT_SCAN_CHARS));
}

// --- The scheduler ------------------------------------------------------------

/**
 * Whether a compile may start now:
 * - `ready`: yes;
 * - `compiling`: this window's own compile is running, one may follow it;
 * - `wait`: not yet, for a moment (a save in flight, the page's version
 *   settling), so look again shortly;
 * - `blocked`: no (read-only, LaTeX unavailable, the page busy elsewhere).
 */
export type LatexCompileGate = "ready" | "compiling" | "wait" | "blocked";

/** Who asked: a pause in typing, or a member (Ctrl/Cmd+Enter, Ctrl/Cmd+S). */
export type LatexCompileKind = "auto" | "manual";

/** The text a compile reads, by workspace path. */
export type LatexCompileInput = ReadonlyMap<string, string>;

export interface LatexCompileSchedulerHost {
  /** Auto compile is on and this file is a document of its own (read when a
   * pause ends). */
  autoCompile(): boolean;
  gate(kind: LatexCompileKind): LatexCompileGate;
  /** The main file and the open files of its project, as they are now. */
  input(): LatexCompileInput;
  /** Start one compile; the host reports it with compileStarted(). */
  compile(kind: LatexCompileKind): void;
}

export interface LatexCompileScheduler {
  /** The member edited the main file or a file of its project. */
  noteEdit(): void;
  /**
   * A compile shortcut: compile now; during this window's own compile, once
   * it ends when the text changed since it started; or later, when a save in
   * flight holds it back.
   */
  requestCompile(): "started" | "queued" | "ignored";
  /** A compile of this window started, whoever asked for it. */
  compileStarted(): void;
  /** The text that compile actually read, once its project is assembled. */
  recordInput(input: LatexCompileInput): void;
  /** That compile ended: succeeded, failed or abandoned (not stopped). */
  compileSettled(): void;
  /** Stop or close: forget any waiting compile. */
  cancel(): void;
}

export function createLatexCompileScheduler(
  host: LatexCompileSchedulerHost,
  options: { readonly delayMs?: number; readonly retryMs?: number } = {},
): LatexCompileScheduler {
  const delayMs = options.delayMs ?? LATEX_AUTO_COMPILE_DELAY_MS;
  const retryMs = options.retryMs ?? LATEX_COMPILE_RETRY_MS;
  let timer: ReturnType<typeof setTimeout> | null = null;
  // Who the armed timer compiles for.
  let pendingKind: LatexCompileKind | null = null;
  let retries = 0;
  // The compile that follows the running one, if any; a member's wins.
  let followUp: LatexCompileKind | null = null;
  // What the last compile read; null before the first.
  let lastInput: Map<string, string> | null = null;

  const clearTimer = () => {
    if (timer !== null) clearTimeout(timer);
    timer = null;
    pendingKind = null;
  };

  // A file opened since the last compile is taken as it read it (its saved
  // text); only a path whose text moved counts as a change.
  const inputChanged = (): boolean => {
    if (!lastInput) return true;
    let changed = false;
    for (const [path, text] of host.input()) {
      const before = lastInput.get(path);
      if (before === undefined) lastInput.set(path, text);
      else if (before !== text) changed = true;
    }
    return changed;
  };

  const arm = (kind: LatexCompileKind, ms: number) => {
    clearTimer();
    pendingKind = kind;
    timer = setTimeout(() => {
      timer = null;
      pendingKind = null;
      attempt(kind);
    }, ms);
  };

  function attempt(kind: LatexCompileKind): void {
    if (kind === "auto" && (!host.autoCompile() || !inputChanged())) return;
    const gate = host.gate(kind);
    if (gate === "ready") {
      host.compile(kind);
      return;
    }
    if (gate === "compiling") {
      if (followUp !== "manual") followUp = kind;
      return;
    }
    if (gate === "wait" && retries < LATEX_COMPILE_MAX_RETRIES) {
      retries += 1;
      arm(kind, retryMs);
    }
    // Blocked: nothing runs until the member edits again.
  }

  return {
    noteEdit() {
      // A member's compile waiting on a save already takes the newest text.
      // The switch is read when the pause ends, not on every keystroke.
      if (pendingKind === "manual") return;
      retries = 0;
      arm("auto", delayMs);
    },
    requestCompile() {
      clearTimer();
      retries = 0;
      const gate = host.gate("manual");
      if (gate === "ready") {
        host.compile("manual");
        return "started";
      }
      if (gate === "compiling") {
        if (!inputChanged()) return "ignored";
        followUp = "manual";
        return "queued";
      }
      if (gate === "wait") {
        retries = 1;
        arm("manual", retryMs);
        return "queued";
      }
      return "ignored";
    },
    compileStarted() {
      clearTimer();
      retries = 0;
      followUp = null;
      lastInput = new Map(host.input());
    },
    recordInput(input) {
      lastInput = new Map(input);
    },
    compileSettled() {
      const next = followUp;
      followUp = null;
      if (next === null) return;
      // The member is still typing: the pause will compile.
      if (next === "auto" && timer !== null) return;
      retries = 0;
      // A tick later, so the window has retired the compile that ended.
      arm(next, 0);
    },
    cancel() {
      clearTimer();
      retries = 0;
      followUp = null;
    },
  };
}
