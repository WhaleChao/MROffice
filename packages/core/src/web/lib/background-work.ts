/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { requestQuietCallback } from "./quiet-callback";

 
export const INPUT_QUIET_MS = 300;
 
export const MAX_DEFER_MS = 1500;




export const FORCED_GAP_MS = 1000;
/** A slice's length while nothing else is going on. */
export const IDLE_SLICE_MS = 8;
/** A slice's length while the person is working elsewhere. */
export const ACTIVE_SLICE_MS = 4;







export const FIRST_SCREEN_SLICE_MS = 40;
 
export const ACTIVE_SHARE = 0.25;
 
export const OWN_INPUT_BOOST_MS = 1500;
/** Unspent credit is capped at this much (a pause does not save up a burst). */
const MAX_CREDIT_MS = 2 * ACTIVE_SLICE_MS;

export type BackgroundPriority = "first-screen" | "background";

export interface BackgroundLaneOptions {
  /** Shown in the diagnostics (`globalThis.__backgroundWork`). */
  label: string;
  


  owner?: () => Element | null | undefined;
  priority?: BackgroundPriority;
}

export interface BackgroundLane {
  

  turn(): Promise<void>;
  



  tryTurn(): boolean;
   
  inputQuiet(): boolean;
  



  sinceLastKeyOrPress(): number;
  /** The current turn's slice is used up (or input is waiting, where the
   * browser can tell): a step that can stop here should, and ask for the
   * next turn. */
  shouldYield(): boolean;
  

  release(): void;
  setPriority(priority: BackgroundPriority): void;
  

  dispose(): void;
  readonly disposed: boolean;
  readonly label: string;
}

interface LaneState {
  readonly label: string;
  readonly owner: (() => Element | null | undefined) | null;
  priority: BackgroundPriority;
  disposed: boolean;
  waiting: { resolve: () => void; since: number } | null;
  lastOwnInputAt: number;
  lastForeignInputAt: number;
  /** For the A/B switch only: the key, press, wheel or touch start elsewhere
   * the page ledger used to wait on. */
  lastLegacyInputAt: number;
   
  lastGrantAt: number;
}

interface Stats {
  grants: number;
  idleGrants: number;
  sharedGrants: number;
  forcedGrants: number;
  firstScreenGrants: number;
  maxWaitMs: number;
  maxSliceMs: number;
  lanes: Record<string, { grants: number; busyMs: number; maxSliceMs: number; maxWaitMs: number }>;
}

type GlobalSwitches = typeof globalThis & {
  /** A probe's A/B switch: true schedules as before this scheduler (every
   * turn at once; the page ledger's own earlier rule for its passes). */
  __backgroundWorkOff?: boolean;
  __backgroundWork?: Stats;
  scheduler?: { yield?: () => Promise<void> };
  navigator?: Navigator & { scheduling?: { isInputPending?: (options?: { includeContinuous?: boolean }) => boolean } };
};

const g = globalThis as GlobalSwitches;

let testClockOffset = 0;

function clock(): number {
  return (typeof performance === "undefined" ? Date.now() : performance.now()) + testClockOffset;
}

const lanes = new Set<LaneState>();
let lastInputAt = Number.NEGATIVE_INFINITY;
let lastKeyOrPressAt = Number.NEGATIVE_INFINITY;
let listening = false;
let credit = MAX_CREDIT_MS;
let creditAt = 0;
/** When the last turn given while the person worked elsewhere ended. */
let lastHeldTurnEndAt = Number.NEGATIVE_INFINITY;
let current: { lane: LaneState; start: number; budget: number; shared: boolean } | null = null;
let pumpQueued = false;
let frameQueued = false;
let timer: ReturnType<typeof setTimeout> | null = null;
let channel: MessageChannel | null = null;
const channelCallbacks: Array<() => void> = [];

function stats(): Stats {
  return (g.__backgroundWork ??= {
    grants: 0, idleGrants: 0, sharedGrants: 0, forcedGrants: 0, firstScreenGrants: 0, maxWaitMs: 0, maxSliceMs: 0, lanes: {},
  });
}

/** A task of its own, before timers, after input already waiting. */
function postTask(callback: () => void): void {
  if (typeof MessageChannel === "function") {
    if (!channel) {
      channel = new MessageChannel();
      channel.port1.onmessage = () => {
        const run = channelCallbacks.shift();
        run?.();
      };
    }
    channelCallbacks.push(callback);
    channel.port2.postMessage(null);
    return;
  }
  setTimeout(callback, 0);
}

function onInput(event: Event): void {
  if (!event.isTrusted) return;
  if (event.type === "pointermove") {
    const pointer = event as PointerEvent;
    // A hover is not work the person waits on; a drag, a pen or a finger is.
    if (pointer.buttons === 0 && pointer.pointerType === "mouse") return;
  }
  noteInput(event.target instanceof Node ? event.target : null, LEGACY_INPUT.has(event.type));
}

const LEGACY_INPUT = new Set(["keydown", "pointerdown", "wheel", "touchstart"]);

function noteInput(target: Node | null, legacy = true): void {
  const now = clock();
  lastInputAt = now;
  if (legacy) lastKeyOrPressAt = now;
  for (const lane of lanes) {
    let inside = false;
    if (lane.owner && target) {
      try { inside = lane.owner()?.contains(target) ?? false; } catch { inside = false; }
    }
    if (inside) lane.lastOwnInputAt = now;
    else {
      lane.lastForeignInputAt = now;
      if (legacy) lane.lastLegacyInputAt = now;
    }
  }
}

function listen(): void {
  if (listening || typeof window === "undefined") return;
  listening = true;
  const options = { capture: true, passive: true } as const;
  for (const type of ["keydown", "beforeinput", "compositionupdate", "pointerdown", "pointermove", "wheel", "touchstart", "touchmove"]) {
    window.addEventListener(type, onInput, options);
  }
}

function hidden(): boolean {
  return typeof document !== "undefined" && document.visibilityState === "hidden";
}

type Verdict = "first-screen" | "own" | "idle" | "forced" | "shared" | "hold";

function legacyVerdict(lane: LaneState, now: number): Verdict {
  if (lane.label !== "docs-ledger" || now - lane.lastLegacyInputAt >= 300) return "idle";
  const since = lane.waiting ? lane.waiting.since : now;
  return now - since >= 2000 ? "forced" : "hold";
}

function verdict(lane: LaneState, now: number, afterPaint: boolean): Verdict {
  if (g.__backgroundWorkOff) return legacyVerdict(lane, now);
  if (lane.priority === "first-screen") return "first-screen";
  if (now - lane.lastOwnInputAt < OWN_INPUT_BOOST_MS && lane.lastOwnInputAt >= lane.lastForeignInputAt) return "own";
  // A hidden page takes no input: nothing to wait for.
  if (hidden() || now - lane.lastForeignInputAt >= INPUT_QUIET_MS) return "idle";
  if (now >= forcedAt(lane, now)) return "forced";
  if (afterPaint && credit > 0) return "shared";
  return "hold";
}

 
function forcedAt(lane: LaneState, now: number): number {
  const since = lane.waiting ? Math.max(lane.waiting.since, lane.lastGrantAt) : now;
  return Math.max(since + MAX_DEFER_MS, lastHeldTurnEndAt + FORCED_GAP_MS);
}

const RANK: Record<Verdict, number> = { "own": 0, "first-screen": 1, "forced": 2, "idle": 3, "shared": 4, "hold": 9 };

function accrue(now: number): void {
  const busy = now - lastInputAt < INPUT_QUIET_MS;
  if (!busy) credit = MAX_CREDIT_MS;
  else credit = Math.min(MAX_CREDIT_MS, credit + Math.max(0, now - creditAt) * ACTIVE_SHARE);
  creditAt = now;
}

function endTurn(): void {
  const turn = current;
  if (!turn) return;
  current = null;
  const now = clock();
  const ms = now - turn.start;
  if (turn.shared) lastHeldTurnEndAt = now;
  // A step's time is what the person waits on whenever it runs while they
  // work: it spends the shared credit, however it was granted.
  if (now - lastInputAt < INPUT_QUIET_MS + ms) {
    accrue(now);
    credit -= ms;
  }
  const s = stats();
  s.maxSliceMs = Math.max(s.maxSliceMs, ms);
  const lane = (s.lanes[turn.lane.label] ??= { grants: 0, busyMs: 0, maxSliceMs: 0, maxWaitMs: 0 });
  lane.busyMs += ms;
  lane.maxSliceMs = Math.max(lane.maxSliceMs, ms);
}

function schedulePump(): void {
  if (pumpQueued) return;
  pumpQueued = true;
  postTask(() => {
    pumpQueued = false;
    pump(false);
  });
}

function scheduleFrame(): void {
  if (frameQueued) return;
  frameQueued = true;
  const afterPaint = () => postTask(() => {
    frameQueued = false;
    pump(true);
  });
  if (typeof requestAnimationFrame === "function" && !hidden()) requestAnimationFrame(afterPaint);
  else setTimeout(afterPaint, 16);
}

function scheduleTimer(at: number): void {
  if (timer !== null) clearTimeout(timer);
  timer = setTimeout(() => {
    timer = null;
    pump(false);
  }, Math.max(0, at - clock()));
}

function pump(afterPaint: boolean): void {
  // The previous turn's step ran in the task that granted it; this is a new task.
  endTurn();
  const now = clock();
  accrue(now);
  let best: LaneState | null = null;
  let bestVerdict: Verdict = "hold";
  let held = false;
  let wakeAt = Number.POSITIVE_INFINITY;
  for (const lane of lanes) {
    if (!lane.waiting) continue;
    const v = verdict(lane, now, afterPaint);
    if (v === "hold") {
      held = true;
      wakeAt = Math.min(wakeAt, lane.lastForeignInputAt + INPUT_QUIET_MS, forcedAt(lane, now));
      continue;
    }
    if (best === null || RANK[v] < RANK[bestVerdict] || (RANK[v] === RANK[bestVerdict] && lane.waiting.since < best.waiting!.since)) {
      best = lane;
      bestVerdict = v;
    }
  }
  if (best) {
    grant(best, bestVerdict, now);
    schedulePump();
    return;
  }
  if (held) {
    // Right after the next frame paints: the shared turn; at the latest when
    // the quiet period or the deferral cap ends.
    scheduleFrame();
    if (Number.isFinite(wakeAt)) scheduleTimer(wakeAt);
  }
}

function grant(lane: LaneState, v: Verdict, now: number): void {
  const waiting = lane.waiting!;
  lane.waiting = null;
  const waited = now - Math.max(waiting.since, lane.lastGrantAt);
  lane.lastGrantAt = now;
  const shared = v === "shared" || v === "forced";
  const budget = shared ? ACTIVE_SLICE_MS
    : v === "first-screen" && now - lastInputAt >= INPUT_QUIET_MS ? FIRST_SCREEN_SLICE_MS
      : IDLE_SLICE_MS;
  current = { lane, start: now, budget, shared };
  const s = stats();
  s.grants += 1;
  if (v === "idle" || v === "own") s.idleGrants += 1;
  else if (v === "shared") s.sharedGrants += 1;
  else if (v === "forced") s.forcedGrants += 1;
  else s.firstScreenGrants += 1;
  s.maxWaitMs = Math.max(s.maxWaitMs, waited);
  const entry = (s.lanes[lane.label] ??= { grants: 0, busyMs: 0, maxSliceMs: 0, maxWaitMs: 0 });
  entry.grants += 1;
  entry.maxWaitMs = Math.max(entry.maxWaitMs, waited);
  waiting.resolve();
}

export function createBackgroundLane(options: BackgroundLaneOptions): BackgroundLane {
  listen();
  const state: LaneState = {
    label: options.label,
    owner: options.owner ?? null,
    priority: options.priority ?? "background",
    disposed: false,
    waiting: null,
    lastOwnInputAt: Number.NEGATIVE_INFINITY,
    lastForeignInputAt: lastInputAt,
    lastLegacyInputAt: Number.NEGATIVE_INFINITY,
    lastGrantAt: Number.NEGATIVE_INFINITY,
  };
  lanes.add(state);
  const lane: BackgroundLane = {
    get disposed() { return state.disposed; },
    get label() { return state.label; },
    turn() {
      if (current?.lane === state) endTurn();
      if (state.disposed) return Promise.resolve();
      if (state.waiting) {
        // One turn at a time: a second caller shares the first's.
        const first = state.waiting;
        return new Promise<void>((resolve) => {
          const previous = first.resolve;
          first.resolve = () => { previous(); resolve(); };
        });
      }
      return new Promise<void>((resolve) => {
        state.waiting = { resolve, since: clock() };
        schedulePump();
      });
    },
    tryTurn() {
      
      // is still running (the caller is inside it).
      if (state.disposed || state.waiting || (current && current.lane !== state)) return false;
      if (current?.lane === state) endTurn();
      const now = clock();
      accrue(now);
      state.waiting = { resolve: () => {}, since: now };
      const v = verdict(state, now, false);
      if (v === "hold" || v === "shared") {
        state.waiting = null;
        return false;
      }
      grant(state, v, now);
      // The turn ends with the task it runs in.
      postTask(() => { if (current?.lane === state) endTurn(); });
      return true;
    },
    inputQuiet() {
      if (state.disposed) return false;
      const v = verdict(state, clock(), false);
      return v === "idle" || v === "own" || v === "first-screen";
    },
    sinceLastKeyOrPress() {
      return clock() - lastKeyOrPressAt;
    },
    shouldYield() {
      if (state.disposed) return true;
      const turn = current;
      if (!turn || turn.lane !== state) return true;
      if (clock() - turn.start >= turn.budget) return true;
      try {
        return g.navigator?.scheduling?.isInputPending?.() === true;
      } catch {
        return false;
      }
    },
    release() {
      if (current?.lane === state) endTurn();
    },
    setPriority(priority) {
      if (state.priority === priority) return;
      state.priority = priority;
      if (state.waiting) schedulePump();
    },
    dispose() {
      if (state.disposed) return;
      state.disposed = true;
      if (current?.lane === state) endTurn();
      lanes.delete(state);
      const waiting = state.waiting;
      state.waiting = null;
      waiting?.resolve();
    },
  };
  return lane;
}






export async function runInLane(lane: BackgroundLane, step: (lane: BackgroundLane) => boolean | Promise<boolean>): Promise<void> {
  for (;;) {
    await lane.turn();
    if (lane.disposed) return;
    let more: boolean;
    try {
      more = await step(lane);
    } catch (error) {
      lane.release();
      throw error;
    }
    if (!more || lane.disposed) {
      lane.release();
      return;
    }
  }
}






export function requestLaneFrame(lane: BackgroundLane, callback: () => void, win: Window = window): () => void {
  let cancelled = false;
  const frame = win.requestAnimationFrame(() => {
    if (cancelled) return;
    if (lane.tryTurn()) {
      callback();
      return;
    }
    void lane.turn().then(() => {
      if (!cancelled && !lane.disposed) callback();
    });
  });
  return () => {
    cancelled = true;
    win.cancelAnimationFrame(frame);
  };
}

/** Tests only: input on `target` (synthetic events are not the person's). */








/** How long since the person last did anything here (a key, a press, a
 * wheel, a touch, a drag). */
export function msSinceInput(): number {
  listen();
  return clock() - lastInputAt;
}

export function requestInputQuietCallback(task: () => void, quietMs = 800): () => void {
  listen();
  let cancelled = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let cancelQuiet: (() => void) | null = null;
  const attempt = () => {
    timer = null;
    if (cancelled) return;
    const since = clock() - lastInputAt;
    if (since < quietMs) {
      timer = setTimeout(attempt, Math.max(16, quietMs - since));
      return;
    }
    cancelQuiet = requestQuietCallback(() => {
      cancelQuiet = null;
      if (cancelled) return;
      // input came while waiting for the idle moment: wait for quiet again
      if (clock() - lastInputAt < quietMs) {
        attempt();
        return;
      }
      task();
    });
  };
  attempt();
  return () => {
    cancelled = true;
    if (timer !== null) clearTimeout(timer);
    cancelQuiet?.();
  };
}

export function noteInputForTests(target: Node | null): void {
  noteInput(target);
}

/** Tests only: a step that takes `ms` of the thread. */
export function spendTimeForTests(ms: number): void {
  testClockOffset += ms;
}

 
export function resetBackgroundWorkForTests(): void {
  for (const lane of lanes) lane.waiting = null;
  lanes.clear();
  lastInputAt = Number.NEGATIVE_INFINITY;
  lastKeyOrPressAt = Number.NEGATIVE_INFINITY;
  credit = MAX_CREDIT_MS;
  lastHeldTurnEndAt = Number.NEGATIVE_INFINITY;
  current = null;
  pumpQueued = false;
  frameQueued = false;
  if (timer !== null) clearTimeout(timer);
  timer = null;
  channelCallbacks.length = 0;
  testClockOffset = 0;
  delete g.__backgroundWork;
}
