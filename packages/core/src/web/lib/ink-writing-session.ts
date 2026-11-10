/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { inkPointerArbiter } from "./ink-pointer-policy";

/** No stroke for this long: the writing is over. */
export const INK_WRITING_IDLE_MS = 1_800;
/** A save held this long goes at the next lift even while the writing goes on. */
export const INK_WRITING_MAX_HOLD_MS = 20_000;










export const INK_WRITING_STALL_MS = 120;
/** How soon the look after a stall looks again. */
export const INK_WRITING_STALL_RECHECK_MS = 32;
/** A page busy for longer than this many re-looks saves anyway. */
const INK_WRITING_STALL_RECHECKS = 6;

function clock(): number {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

export interface InkWritingClock {
  /** The pen is on the glass (a stroke is running). */
  readonly penDown: boolean;
  /** When the last stroke ended (`performance.now()` time), or null. */
  readonly lastStrokeEndedAt: number | null;
}

function arbiterClock(): InkWritingClock {
  return {
    // A stroke whose lift never reached a surface does not hold the save.
    penDown: inkPointerArbiter.strokeRunning(),
    lastStrokeEndedAt: inkPointerArbiter.lastStrokeEndedAt,
  };
}

/** How long until the session ends, 0 when none runs. */
export function inkWritingRemainingMs(
  now: number = clock(),
  pen: InkWritingClock = arbiterClock(),
  idleMs: number = INK_WRITING_IDLE_MS,
): number {
  if (pen.penDown) return idleMs;
  const ended = pen.lastStrokeEndedAt;
  if (ended === null || !Number.isFinite(ended)) return 0;
  const since = now - ended;
  if (since < 0) return 0;
  return Math.max(0, idleMs - since);
}

/**
 * After the last stroke, the saves a document makes leave out the page's
 * pagination records for this long (they are measured from the page's
 * layout and digested over the whole document: a forced layout and several
 * whole-document passes, which landed under the pen when a long session's
 * save went, or when the writer went on just after a pause). The records
 * come with the save that follows once this has passed (NoteEditor).
 */
export const INK_LAYOUT_SETTLE_MS = 8_000;

/** How long until the pen has been away long enough for the layout records
 * to be measured again; 0 when it has (or never wrote). */
export function inkLayoutSettlesInMs(
  now: number = clock(),
  pen: InkWritingClock = arbiterClock(),
  hovering: boolean = inkPointerArbiter.penHovering(),
): number {
  if (pen.penDown || hovering) return INK_LAYOUT_SETTLE_MS;
  const ended = pen.lastStrokeEndedAt;
  if (ended === null || !Number.isFinite(ended)) return 0;
  const since = now - ended;
  if (since < 0) return 0;
  return Math.max(0, INK_LAYOUT_SETTLE_MS - since);
}

/** The pen wrote within INK_LAYOUT_SETTLE_MS (or is on, or over, the glass). */
export function inkLayoutUnsettled(
  now: number = clock(),
  pen: InkWritingClock = arbiterClock(),
  hovering: boolean = inkPointerArbiter.penHovering(),
): boolean {
  return inkLayoutSettlesInMs(now, pen, hovering) > 0;
}

/** A writing session is running. */
export function inkWritingActive(
  now: number = clock(),
  pen: InkWritingClock = arbiterClock(),
  idleMs: number = INK_WRITING_IDLE_MS,
): boolean {
  return inkWritingRemainingMs(now, pen, idleMs) > 0;
}

/**
 * Whether a save first held at `heldSince` (`performance.now()` time) still
 * waits: the session runs, and it has not waited its longest yet -- past
 * that it waits only for the pen to lift, never into a stroke.
 */
export function inkWritingHoldsSave(
  heldSince: number | null,
  now: number = clock(),
  pen: InkWritingClock = arbiterClock(),
  idleMs: number = INK_WRITING_IDLE_MS,
  maxHoldMs: number = INK_WRITING_MAX_HOLD_MS,
): boolean {
  if (!inkWritingActive(now, pen, idleMs)) return false;
  if (pen.penDown) return true;
  return heldSince === null || now - heldSince < maxHoldMs;
}










export const INK_WRITING_IDLE_PER_SAVE_COST = 4;
export const INK_WRITING_IDLE_COSTLY_MAX_MS = 8_000;
export const INK_WRITING_HOLD_PER_SAVE_COST = 30;
export const INK_WRITING_HOLD_COSTLY_MAX_MS = 120_000;








export const INK_WRITING_LATE_PAUSE_MS = 450;

export interface InkWritingSaveGate {
  /**
   * Whether the automatic save asked for now waits. When it does, `resume`
   * runs (once) when it may go: the session ended, or held long enough and
   * the pen is up. Asking again before then replaces the pending resume.
   */
  hold(resume: () => void): boolean;
  /** Drop a pending resume (the owner went away or saved by other means). */
  cancel(): void;
  /** What the last save cost the main thread, ms: a costly one waits longer. */
  noteSaveCost(ms: number): void;
}

/**
 * The gate an autosave puts in front of its timer-driven saves. One gate per
 * save owner; it remembers when its first held save was asked for, so a long
 * session cannot hold it forever.
 */
export function createInkWritingSaveGate(
  timers: {
    readonly setTimeout: (run: () => void, ms: number) => unknown;
    readonly clearTimeout: (handle: unknown) => void;
    readonly now?: () => number;
    readonly pen?: () => InkWritingClock;
    /** The pause that ends a session (another writer's, e.g. typing). */
    readonly idleMs?: number;
    /** The longest a save waits while the session goes on. */
    readonly maxHoldMs?: number;
    /** The pause a save past its longest hold waits for (0: the lift itself). */
    readonly latePauseMs?: number;
  } = {
    setTimeout: (run, ms) => globalThis.setTimeout(run, ms),
    clearTimeout: (handle) => globalThis.clearTimeout(handle as ReturnType<typeof setTimeout>),
  },
): InkWritingSaveGate {
  const now = timers.now ?? clock;
  const pen = timers.pen ?? arbiterClock;
  const baseIdleMs = timers.idleMs ?? INK_WRITING_IDLE_MS;
  const baseMaxHoldMs = timers.maxHoldMs ?? INK_WRITING_MAX_HOLD_MS;
  const latePauseMs = timers.latePauseMs ?? INK_WRITING_LATE_PAUSE_MS;
  let idleMs = baseIdleMs;
  let maxHoldMs = baseMaxHoldMs;
  // The saves' recent cost, decaying: one cheap save after costly ones does
  // not end the longer waits at once.
  let saveCost = 0;
  let heldSince: number | null = null;
  let timer: unknown = null;
  // When the scheduled look was due (`now()` time).
  let dueAt = 0;
  let stallRechecks = 0;
  // The pending save waits only for one look confirming the session ended.
  let confirming = false;
  let pending: (() => void) | null = null;
  // The resume is running: the save it asks for goes (a long session's save
  // released at a lift must not be held again by its own request).
  let resuming = false;
  const stop = () => {
    if (timer !== null) timers.clearTimeout(timer);
    timer = null;
  };
  const schedule = (wait: number, at: number) => {
    dueAt = at + wait;
    timer = timers.setTimeout(check, wait);
  };
  // How long the pen has been up, ms (Infinity before any stroke).
  const awayFor = (state: InkWritingClock, at: number) => {
    const ended = state.lastStrokeEndedAt;
    return ended === null || !Number.isFinite(ended) ? Infinity : Math.max(0, at - ended);
  };
  const pastLongest = (at: number) => heldSince !== null && at - heldSince >= maxHoldMs;
  const holds = (at: number, state: InkWritingClock) => {
    if (inkWritingHoldsSave(heldSince, at, state, idleMs, maxHoldMs)) return true;
    // Past its longest hold: a short pause, not the lift itself.
    return latePauseMs > 0 && heldSince !== null && !state.penDown && inkWritingActive(at, state, idleMs) &&
      at - heldSince < 2 * maxHoldMs && awayFor(state, at) < latePauseMs;
  };
  const scheduleFor = (state: InkWritingClock, at: number) => {
    // While the pen is down the next look is soon (the lift may be the
    // moment a long-held save goes); otherwise when the session would end,
    // or, past the longest hold, when the pen has been up a short pause.
    const wait = state.penDown
      ? 120
      : latePauseMs > 0 && pastLongest(at)
        ? Math.max(40, latePauseMs - awayFor(state, at))
        : Math.max(40, inkWritingRemainingMs(at, state, idleMs));
    schedule(Math.min(wait, idleMs), at);
  };
  const release = () => {
    const run = pending;
    pending = null;
    heldSince = null;
    stallRechecks = 0;
    confirming = false;
    if (!run) return;
    resuming = true;
    try {
      run();
    } finally {
      resuming = false;
    }
  };
  function check() {
    timer = null;
    if (!pending) return;
    const at = now();
    if (at - dueAt > INK_WRITING_STALL_MS && stallRechecks < INK_WRITING_STALL_RECHECKS) {
      // The page was busy past this look: hear the pen's queued events first.
      stallRechecks += 1;
      schedule(INK_WRITING_STALL_RECHECK_MS, at);
      return;
    }
    stallRechecks = 0;
    const state = pen();
    if (holds(at, state)) {
      // The writing went on after all: an ordinary held save from here.
      confirming = false;
      if (heldSince === null) heldSince = at;
      scheduleFor(state, at);
      return;
    }
    release();
  }
  return {
    hold(resume) {
      if (resuming) return false;
      const at = now();
      const state = pen();
      if (!holds(at, state)) {
        const ended = state.lastStrokeEndedAt;
        if (pending === null && ended !== null && Number.isFinite(ended) && at - ended >= 0 && at - ended < 2 * idleMs) {
          // The session ended just now, by the clock. The timer asking may
          // itself have run late, ahead of a stroke's queued events: one
          // short look confirms it before the save goes.
          pending = resume;
          confirming = true;
          stop();
          schedule(INK_WRITING_STALL_RECHECK_MS, at);
          return true;
        }
        if (confirming && timer !== null) {
          // A confirming look is on its way: this request is the save it lets go.
          pending = resume;
          return true;
        }
        stop();
        pending = null;
        heldSince = null;
        confirming = false;
        return false;
      }
      if (heldSince === null) heldSince = at;
      if (confirming) {
        // Held for real now: the confirming look becomes an ordinary one.
        confirming = false;
        stop();
      }
      pending = resume;
      if (timer === null) scheduleFor(state, at);
      return true;
    },
    noteSaveCost(ms) {
      if (!Number.isFinite(ms) || ms < 0) return;
      saveCost = Math.max(ms, saveCost * 0.7);
      idleMs = Math.max(baseIdleMs, Math.min(INK_WRITING_IDLE_COSTLY_MAX_MS, saveCost * INK_WRITING_IDLE_PER_SAVE_COST));
      maxHoldMs = Math.max(baseMaxHoldMs, Math.min(INK_WRITING_HOLD_COSTLY_MAX_MS, saveCost * INK_WRITING_HOLD_PER_SAVE_COST));
    },
    cancel() {
      stop();
      pending = null;
      heldSince = null;
      stallRechecks = 0;
      confirming = false;
    },
  };
}
