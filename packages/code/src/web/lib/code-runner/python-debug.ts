/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Python breakpoint debugging: RECORD-AND-REPLAY over the live kernel.
 *
 * The Pyodide worker cannot truly block mid-Python (no JSPI guarantee, and
 * SharedArrayBuffer would force COOP/COEP onto the whole app), so a debug run
 * executes to completion ONCE under sys.settrace — recording every breakpoint
 * hit plus a bounded window of line-steps after each hit (locals previews +
 * user-frame stack) — and THIS module replays the recording with the exact
 * Continue / Step / paused-panel semantics of the JS debugger
 * ({@link DebugRunController} shape parity). Compared to the JS instrumentation
 * approach, breakpoints work INSIDE function bodies, and a "paused" kernel can
 * never wedge: side effects (prints, plots, kernel globals) land exactly once.
 *
 * The controller resolves `result` when playback finishes: after the last
 * event, on Continue with no further breakpoint hits, on abort, or immediately
 * when nothing was recorded.
 */

import {
  debugRunPython,
  type PythonDebugBreakpoint,
  type PythonDebugEvent,
  type PythonDebugStackFrame,
  type PythonHitCondition,
} from "@mrnote/core/web/lib/code-runner/python-runner.ts";
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";

/** Generous cap for a condition expression (one line of Python, not a program). */
export const MAX_PYTHON_BREAKPOINT_CONDITION_LENGTH = 500;

/**
 * Parse a VSCode-style hit-count expression into a normalized predicate, or
 * `null` when the text is empty/garbage (caller then treats the breakpoint as
 * having NO hit-count — a plain/conditional breakpoint, unchanged).
 *
 * Grammar (whitespace-insensitive):
 *   - `N`            → `>= N`  (bare number: "this hit onward", VSCode default)
 *   - `>= N`, `> N`, `<= N`, `< N`
 *   - `== N`, `= N`  → `== N`
 *   - `% N`          → every Nth hit (remainder 0)
 *   - `% N == R`     → hit count ≡ R (mod N)
 * N/R must be non-negative integers; anything else → `null`.
 */
export function parseHitCondition(raw: string): PythonHitCondition | null {
  if (typeof raw !== "string") return null;
  const text = raw.trim();
  if (text === "") return null;

  // `% N` or `% N == R` (every-Nth, optionally offset).
  const mod = /^%\s*(\d+)(?:\s*==?\s*(\d+))?$/.exec(text);
  if (mod) {
    const value = Number(mod[1]);
    if (value <= 0) return null; // a 0 modulus is meaningless — ignore.
    const rem = mod[2] === undefined ? 0 : Number(mod[2]) % value;
    return { op: "%", value, rem };
  }

  // `<op> N`, where op is a comparison; bare `N` defaults to `>=`.
  const cmp = /^(>=|<=|==|=|>|<)?\s*(\d+)$/.exec(text);
  if (cmp) {
    const value = Number(cmp[2]);
    if (!Number.isFinite(value) || value < 0) return null;
    const raw2 = cmp[1];
    const op: PythonHitCondition["op"] =
      raw2 === undefined || raw2 === ">="
        ? ">="
        : raw2 === "=" || raw2 === "=="
          ? "=="
          : (raw2 as ">" | "<" | "<=");
    return { op, value };
  }

  return null; // garbage → no hit-condition (the hit always counts).
}

/**
 * Whether a hit-count predicate is satisfied at the given 1-based reach count
 * (the running tally of conditionally-true reaches for that line). A `null`
 * predicate (no hit-count) is always satisfied. Mirrors the worker tracer's
 * `_hit_ok`; used by tests and as the canonical TS reference.
 */
export function hitMatches(hit: PythonHitCondition | null | undefined, count: number): boolean {
  if (!hit) return true;
  switch (hit.op) {
    case ">=":
      return count >= hit.value;
    case ">":
      return count > hit.value;
    case "<=":
      return count <= hit.value;
    case "<":
      return count < hit.value;
    case "==":
      return count === hit.value;
    case "%":
      return hit.value === 0 ? true : count % hit.value === (hit.rem ?? 0);
    default:
      return true;
  }
}

/**
 * Pre-validate a Python breakpoint condition: bounded, single-line, non-empty.
 * Unlike the JS validator there is no in-browser Python parser to lean on, so
 * syntax is enforced at run time instead: the worker evals the condition in the
 * frame scope and a raising condition records the hit with `conditionError`
 * (VSCode semantics) — a broken condition can never break the run itself.
 */
export function isValidPythonBreakpointCondition(condition: string): boolean {
  if (typeof condition !== "string") return false;
  const trimmed = condition.trim();
  if (trimmed === "" || trimmed.length > MAX_PYTHON_BREAKPOINT_CONDITION_LENGTH) {
    return false;
  }
  // A condition is a single expression: newlines/NULs can only smuggle
  // statements (eval would reject them anyway — fail fast in the UI).
  return !/[\n\r\0]/.test(condition);
}

/**
 * The paused-state message handed to the panel at each replayed event. A
 * superset of the JS debugger's WorkerDebugPausedMessage (line/locals/step), so
 * one panel renders both engines; `stack`/`kind` are the recording extras.
 */
export interface PythonDebugPausedMessage {
  type: "debug-paused";
  /** 1-based source line of this event. */
  line: number;
  /** `{ name: previewString }` locals of the paused frame. */
  locals: Record<string, string>;
  /** 1-based ordinal of this pause within the replay. */
  step: number;
  /** Breakpoint hit, a line-step, or the terminal uncaught-exception frame. */
  kind: "breakpoint" | "step" | "exception";
  /** User-frame call stack, innermost first. */
  stack: PythonDebugStackFrame[];
  /**
   * A per-event condition error (legacy field — the tracer no longer records a
   * raising condition as a hit, so this is unset from recordings; kept so the
   * panel and any caller that sets it directly still render).
   */
  conditionError?: string;
  /**
   * A one-time, non-fatal note: the FIRST breakpoint condition that raised
   * during recording (the run kept going; that breakpoint was skipped, VSCode
   * "failed condition"). Surfaced on the first pause only.
   */
  conditionNote?: string;
  /** The recording hit its cap — later hits were not captured. */
  truncated?: boolean;
  /** Bounded display text for a terminal uncaught-exception pause. */
  exception?: string;
}

/**
 * The replay index a depth-based step lands on, given the recorded events and
 * the index currently paused at. All three are PURE over the recording (the
 * stack depth = `stack.length`, innermost-first), so they are unit-tested in
 * isolation. A return >= events.length means "no further matching event" — the
 * controller finishes the run there (VSCode: step-out at the top frame, or any
 * step past the last recorded line, ends execution).
 *
 *  - step-into: the very next event (descend into a call if it is deeper).
 *  - step-over: the next event at the current depth OR shallower (skip frames
 *    pushed by a call on this line — they are strictly deeper).
 *  - step-out:  the next event strictly shallower than now (return to a caller).
 */
export function nextStepIntoIndex(_events: PythonDebugEvent[], index: number): number {
  return index + 1;
}

/**
 * The TRUE user-frame depth at an event. Prefer the uncapped `depth` field
 * (recorded since the worker started emitting it); fall back to the displayed
 * `stack.length` only for older recordings that lack it. The `stack` slice is
 * capped at 20 frames, so its length saturates in deep recursion and cannot
 * distinguish frames beyond that — using it for targeting picks the wrong frame.
 */
function eventDepth(event: PythonDebugEvent | undefined): number {
  if (!event) return 0;
  return typeof event.depth === "number" ? event.depth : event.stack.length;
}

export function nextStepOverIndex(events: PythonDebugEvent[], index: number): number {
  const depth = eventDepth(events[index]);
  for (let i = index + 1; i < events.length; i += 1) {
    if (eventDepth(events[i]) <= depth) return i;
  }
  return events.length;
}

export function nextStepOutIndex(events: PythonDebugEvent[], index: number): number {
  const depth = eventDepth(events[index]);
  for (let i = index + 1; i < events.length; i += 1) {
    if (eventDepth(events[i]) < depth) return i;
  }
  return events.length;
}

/**
 * Resolve the locals to show for the frame at `frameIndex` (0 = innermost) of
 * the event at `eventIndex`. The recording stores locals only for the CURRENT
 * (innermost) frame of each event, so for an OUTER frame we surface the locals
 * from the most recent event (at or before `eventIndex`) where THAT exact frame
 * was the innermost one.
 *
 * We identify "that exact frame" by the FULL caller chain from it outward — the
 * frame plus every caller below it (name + pinned line) must match. While the
 * frame is on the stack its callers do not move, so the chain `stack.slice(k)`
 * is invariant from the moment the frame was innermost until now. Matching the
 * whole chain (not just the frame's own name+line) is what disambiguates
 * RECURSION: two `foo` frames at the same source line are distinguished by their
 * different caller chains. Returns null when no such event exists (the frame was
 * entered before the recording window opened).
 */
export function resolveFrameLocals(
  events: PythonDebugEvent[],
  eventIndex: number,
  frameIndex: number,
): Record<string, string> | null {
  const event = events[eventIndex];
  if (!event) return null;
  // The innermost frame's locals are exactly what the event already carries.
  if (frameIndex <= 0) return event.locals;
  if (frameIndex >= event.stack.length) return null;
  // The target frame as innermost: its stack was exactly this tail (the frame +
  // its callers), so find the most recent event whose stack equals it.
  const tail = event.stack.slice(frameIndex);
  for (let i = eventIndex; i >= 0; i -= 1) {
    const s = events[i].stack;
    if (
      s.length === tail.length &&
      s.every((frame, j) => frame.name === tail[j].name && frame.line === tail[j].line)
    ) {
      return events[i].locals;
    }
  }
  return null;
}

export interface PythonDebugOptions {
  breakpoints: number[];
  /** Per-line condition expressions (only valid ones are sent). */
  conditions?: Record<number, string>;
  /**
   * Per-line hit-count expressions (VSCode hit-count breakpoints). Parsed via
   * {@link parseHitCondition}; only parseable ones are sent. Unparseable text is
   * ignored — the breakpoint keeps its plain/conditional behaviour.
   */
  hitConditions?: Record<number, string>;
  /** Kernel namespace (the page id) — same kernel as normal runs. */
  namespace?: string;
  /** Pre-supplied stdin for input() calls, exactly like a normal run. */
  stdin?: string;
  /**
   * Aborting during recording physically retires its worker generation;
   * aborting after the recording arrived only ends the local replay.
   */
  signal?: AbortSignal;
  /** Replay one terminal pause when an exception escapes user code. */
  pauseOnUncaughtExceptions?: boolean;
  onPause: (pause: PythonDebugPausedMessage) => void;
  timeoutMs?: number;
}

/** Same surface as the JS DebugRunController so the panel drives either. */
export interface PythonDebugController {
  readonly result: Promise<RunResult>;
  continue(): void;
  /** Back-compat alias for {@link PythonDebugController.stepInto}. */
  step(): void;
  /** Advance to the next event in the current frame or shallower (VSCode). */
  stepOver(): void;
  /** Advance to the very next recorded event (descend into a call). */
  stepInto(): void;
  /** Advance to the next event in a caller frame (return out). */
  stepOut(): void;
  /**
   * Locals for the given stack frame (0 = innermost) of the current pause,
   * resolved from the recording. Null when execution has finished or the
   * outer frame's locals were never captured.
   */
  frameLocals(frameIndex: number): Record<string, string> | null;
}

/**
 * Start a Python debug session: one recorded run, then interactive replay.
 * `continue()` advances to the next breakpoint hit; `stepInto()` (a.k.a.
 * `step()`) to the very next recorded event; `stepOver()`/`stepOut()` navigate
 * by recorded stack depth (see {@link nextStepOverIndex}/{@link nextStepOutIndex}).
 * Always resolves `result` (never rejects).
 */
export function runPythonDebug(
  code: string,
  options: PythonDebugOptions,
): PythonDebugController {
  const payload: PythonDebugBreakpoint[] = options.breakpoints.map((line) => {
    const bp: PythonDebugBreakpoint = { line };
    const condition = options.conditions?.[line];
    if (condition && isValidPythonBreakpointCondition(condition)) {
      bp.condition = condition.trim();
    }
    const hitText = options.hitConditions?.[line];
    if (hitText) {
      const hit = parseHitCondition(hitText);
      if (hit) bp.hit = hit;
    }
    return bp;
  });

  let settled = false;
  let resolveResult!: (result: RunResult) => void;
  const result = new Promise<RunResult>((resolve) => (resolveResult = resolve));

  let events: PythonDebugEvent[] = [];
  let truncated = false;
  /** First raising-condition note from the recording (surfaced once). */
  let conditionNote: string | undefined;
  let outcome: RunResult = { logs: [] };
  /** Index of the event currently paused at; -1 before the first pause. */
  let index = -1;
  /** The worker run finished and replay is active. */
  let playing = false;
  /** The low-level worker still owns the physical recording. */
  let recording = true;
  let abortListener: (() => void) | null = null;

  const finish = () => {
    if (settled) return;
    settled = true;
    recording = false;
    playing = false;
    if (options.signal && abortListener) {
      options.signal.removeEventListener("abort", abortListener);
    }
    resolveResult(outcome);
  };

  const pauseAt = (i: number) => {
    if (settled) return;
    if (i >= events.length) {
      finish();
      return;
    }
    index = i;
    const event = events[i];
    options.onPause({
      type: "debug-paused",
      line: event.line,
      locals: event.locals,
      step: i + 1,
      kind: event.kind,
      stack: event.stack,
      ...(event.conditionError ? { conditionError: event.conditionError } : {}),
      // Surface the cap and the first raising-condition note once, on the first
      // pause, so the user sees them before walking the replay.
      ...(i === 0 && truncated ? { truncated: true } : {}),
      ...(i === 0 && conditionNote ? { conditionNote } : {}),
      ...(event.kind === "exception" && event.exception
        ? { exception: event.exception }
        : {}),
    });
  };

  abortListener = () => {
    // During recording the runner owns cancellation: it must physically retire
    // the exact worker generation and return its interrupted receipt. Resolving
    // here would merely hide a still-running Python program behind a stopped UI.
    if (!recording) finish();
  };

  if (options.signal?.aborted) {
    // Never create or touch a kernel for a controller its owner already retired.
    recording = false;
    outcome = { logs: [], error: "Execution interrupted.", interrupted: true };
    finish();
  } else {
    options.signal?.addEventListener("abort", abortListener, { once: true });
    // Close the check/register window before the worker request is created.
    if (options.signal?.aborted) {
      recording = false;
      outcome = {
        logs: [],
        error: "Execution interrupted.",
        interrupted: true,
      };
      finish();
    } else {
      void debugRunPython(code, payload, {
        namespace: options.namespace,
        stdin: options.stdin,
        timeoutMs: options.timeoutMs,
        signal: options.signal,
        pauseOnUncaughtExceptions:
          options.pauseOnUncaughtExceptions === true,
      })
        .then((res) => {
          // The runner removes its physical-abort listener when the recording
          // receipt is routed. From here on cancellation is replay-only.
          recording = false;
          if (settled) return;
          outcome = res;
          // Abort may land after the worker receipt but before this continuation.
          // Preserve the committed result, but never hydrate a stale pause UI.
          if (options.signal?.aborted) {
            finish();
            return;
          }
          events = res.debugEvents ?? [];
          truncated = res.debugTruncated === true;
          conditionNote = res.debugConditionNote;
          playing = true;
          if (events.length === 0) {
            finish();
            return;
          }
          pauseAt(0);
        })
        .catch((err) => {
          recording = false;
          if (settled) return;
          // debugRunPython never rejects by contract; belt-and-braces.
          outcome = {
            logs: [],
            error: err instanceof Error ? err.message : String(err),
          };
          finish();
        });
    }
  }

  const stepInto = () => {
    if (settled || !playing) return;
    pauseAt(nextStepIntoIndex(events, index));
  };
  const stepOver = () => {
    if (settled || !playing) return;
    pauseAt(nextStepOverIndex(events, index));
  };
  const stepOut = () => {
    if (settled || !playing) return;
    pauseAt(nextStepOutIndex(events, index));
  };

  return {
    result,
    continue: () => {
      if (settled || !playing) return;
      // Run to the next GENUINE breakpoint hit. `kind === "breakpoint"` IS the
      // recorded per-step hit flag: the tracer emits a `breakpoint` event only
      // when the line's condition was truthy AND its hit-count matched (both
      // need Python eval, done once at recording time), so replay honours
      // conditional/hit-count breakpoints without re-running anything.
      for (let i = index + 1; i < events.length; i += 1) {
        if (
          events[i].kind === "breakpoint" ||
          events[i].kind === "exception"
        ) {
          pauseAt(i);
          return;
        }
      }
      finish();
    },
    step: stepInto,
    stepInto,
    stepOver,
    stepOut,
    frameLocals: (frameIndex: number) => {
      if (!playing || index < 0) return null;
      return resolveFrameLocals(events, index, frameIndex);
    },
  };
}
