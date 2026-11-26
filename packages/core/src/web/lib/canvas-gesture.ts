/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A burst of wheel or scroll steps is over this long after its last step. */
export const BURST_REST_MS = 220;
/** A pointer gesture whose end never came stops holding work after this. */
export const MAX_GESTURE_MS = 15_000;

type GestureSwitches = typeof globalThis & {
  __canvasGestureHoldOff?: boolean;
  __canvasGestures?: { started: number; held: number };
};

const g = globalThis as GestureSwitches;

function clock(): number {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

const pointerGestures = new Map<number, number>();
let nextToken = 1;
let burstUntil = Number.NEGATIVE_INFINITY;
let endTimer: ReturnType<typeof setTimeout> | null = null;
const endListeners = new Set<() => void>();

function holdOff(): boolean {
  return g.__canvasGestureHoldOff === true;
}

function expire(now: number): void {
  for (const [token, since] of pointerGestures) {
    if (now - since >= MAX_GESTURE_MS) pointerGestures.delete(token);
  }
}

/** A gesture is in progress now (always false with the kill switch on). */
export function canvasGestureActive(now = clock()): boolean {
  if (holdOff()) return false;
  if (now < burstUntil) return true;
  if (pointerGestures.size === 0) return false;
  expire(now);
  return pointerGestures.size > 0;
}

/** When the current gesture can end at the earliest (for a wake-up timer). */
export function canvasGestureRestsAt(now = clock()): number {
  if (!canvasGestureActive(now)) return now;
  if (pointerGestures.size > 0) {
    let first = Number.POSITIVE_INFINITY;
    for (const since of pointerGestures.values()) first = Math.min(first, since);
    return Math.max(burstUntil, first + MAX_GESTURE_MS);
  }
  return burstUntil;
}

function notifyEnd(): void {
  if (endTimer !== null) {
    clearTimeout(endTimer);
    endTimer = null;
  }
  const now = clock();
  if (canvasGestureActive(now)) {
    endTimer = setTimeout(notifyEnd, Math.max(1, canvasGestureRestsAt(now) - now));
    return;
  }
  for (const listener of [...endListeners]) {
    try {
      listener();
    } catch {
      /* a listener's failure is its own */
    }
  }
  if (heldTasks.length > 0) scheduleRelease();
}

function armEnd(): void {
  if (endTimer !== null) clearTimeout(endTimer);
  const now = clock();
  endTimer = setTimeout(notifyEnd, Math.max(1, canvasGestureRestsAt(now) - now));
}

function count(): void {
  const s = (g.__canvasGestures ??= { started: 0, held: 0 });
  s.started += 1;
}

/** A pointer gesture starts; call the returned function when it ends. */
export function beginCanvasGesture(): () => void {
  if (holdOff()) return () => {};
  const token = nextToken++;
  pointerGestures.set(token, clock());
  count();
  armEnd();
  let ended = false;
  return () => {
    if (ended) return;
    ended = true;
    pointerGestures.delete(token);
    notifyEnd();
  };
}

/** One step of a wheel or scroll burst (zoom, pan). */
export function noteCanvasGestureBurst(restMs = BURST_REST_MS): void {
  if (holdOff()) return;
  const now = clock();
  if (!canvasGestureActive(now)) count();
  burstUntil = Math.max(burstUntil, now + restMs);
  armEnd();
}

/** How long a pointer that reached a title bar holds background work: a
 * person who points at a title bar presses it within about a second. */
export const APPROACH_HOLD_MS = 1_200;








export function noteCanvasGestureApproach(): void {
  if ((g as GestureSwitches & { __canvasGestureApproachOff?: boolean }).__canvasGestureApproachOff === true) return;
  noteCanvasGestureBurst(APPROACH_HOLD_MS);
}

/** Called each time a gesture has ended (and no other is in progress). */
export function onCanvasGestureEnd(listener: () => void): () => void {
  endListeners.add(listener);
  return () => endListeners.delete(listener);
}

/** Work held for a gesture starts this long after it ends, so the next
 * gesture (a pan right after a pinch) finds the thread free, and then a few
 * milliseconds of it a frame: let go all at once in the first frame after
 * the gesture, the passes it held (every Docs editor's line slices, object
 * layout, font news) were one long frame the next pan started into. */
export const RELEASE_GRACE_MS = 250;
const RELEASE_FRAME_BUDGET_MS = 8;

interface HeldTask {
  run: () => void;
  cancelled: boolean;
}

const heldTasks: HeldTask[] = [];
let releaseTimer: ReturnType<typeof setTimeout> | null = null;
let releaseFrame = 0;

/** Held tasks are being let go (or about to be): new ones queue behind. */
function releasing(): boolean {
  return heldTasks.length > 0 && !holdOff();
}

function holdTask(run: () => void): HeldTask {
  const s = (g.__canvasGestures ??= { started: 0, held: 0 });
  s.held += 1;
  const task = { run, cancelled: false };
  heldTasks.push(task);
  if (!canvasGestureActive()) scheduleRelease();
  return task;
}

function scheduleRelease(): void {
  if (releaseTimer !== null || releaseFrame) return;
  releaseTimer = setTimeout(() => {
    releaseTimer = null;
    releaseSome();
  }, RELEASE_GRACE_MS);
}

function releaseSome(): void {
  releaseFrame = 0;
  // A gesture began meanwhile: its end lets them go.
  if (canvasGestureActive()) return;
  const started = clock();
  let ran = 0;
  while (heldTasks.length > 0 && (ran === 0 || clock() - started < RELEASE_FRAME_BUDGET_MS)) {
    const task = heldTasks.shift()!;
    if (task.cancelled) continue;
    ran += 1;
    try {
      task.run();
    } catch {
      /* a task's failure is its own */
    }
  }
  if (heldTasks.length === 0) return;
  if (typeof requestAnimationFrame === "function") releaseFrame = requestAnimationFrame(releaseSome) as unknown as number;
  else releaseTimer = setTimeout(() => { releaseTimer = null; releaseSome(); }, 16);
}

/**
 * `requestAnimationFrame` for background layout: the callback runs in the
 * next frame, or, while a canvas gesture is in progress, in the first frame
 * after it. Returns the cancel.
 */
export function requestQuietFrame(win: Window, callback: () => void): () => void {
  let cancelled = false;
  let frame = 0;
  let held: HeldTask | null = null;
  frame = win.requestAnimationFrame(() => {
    frame = 0;
    if (cancelled) return;
    if (canvasGestureActive() || releasing()) {
      held = holdTask(callback);
      return;
    }
    callback();
  });
  return () => {
    cancelled = true;
    if (frame) win.cancelAnimationFrame(frame);
    if (held) held.cancelled = true;
    held = null;
  };
}
/**
 * Runs `task` now, or once the canvas gesture in progress has ended: for
 * news that sets layout work going (a face landing, a face declared) and can
 * wait the length of a drag or a pinch.
 */
export function afterCanvasGesture(task: () => void): () => void {
  if (!canvasGestureActive() && !releasing()) {
    task();
    return () => {};
  }
  const held = holdTask(task);
  return () => {
    held.cancelled = true;
  };
}

/** Tests only. */
export function resetCanvasGesturesForTests(): void {
  pointerGestures.clear();
  burstUntil = Number.NEGATIVE_INFINITY;
  if (endTimer !== null) clearTimeout(endTimer);
  endTimer = null;
  endListeners.clear();
  heldTasks.length = 0;
  if (releaseTimer !== null) clearTimeout(releaseTimer);
  releaseTimer = null;
  if (releaseFrame && typeof cancelAnimationFrame === "function") cancelAnimationFrame(releaseFrame);
  releaseFrame = 0;
}
