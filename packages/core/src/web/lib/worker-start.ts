/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// A worker's start, survived once
//

// chunk (lib0, under Yjs) failed to load ("Uncaught NetworkError: Failed to
// execute 'importScripts' on 'WorkerGlobalScope'"). The chunk was there, the
// network was not, for a moment. Every worker of the app fell back to the
// page's own thread, but the error went up as the page's own uncaught error,
// and the page-causal derive worker never came back for the rest of the
// session: one network blip made every later page read slower.
//
// startResilientWorker runs a worker's start the same way everywhere:
// - an error before the worker says it is ready is handled here (the
//   browser no longer reports it as the page's uncaught error) and reported
//   as what it is, a worker that did not start (WorkerStartFailed, with the
//   worker's name and the error's own words, the chunk among them);
// - the caller gets one more try, after a pause (a chunk that failed to
//   load usually loads a moment later); a second failure is final, and the
//   caller's own fallback (its work on the page's thread) takes over. A
//   worker the page cannot do without is never given up: its starts are
//   paced, the pause doubling from 1.5 s to at most 30 s.
//
// Webpack's chunk loading inside the worker is not touched: the worker is
// made again from the page, which is where a retry can be paced and seen.
// ---------------------------------------------------------------------------

import { reportClientError } from "./errors/report-beacon";

/** The pause before the first retry; each failed start after it doubles it. */
export const WORKER_START_RETRY_MS = 1_500;
/** The longest pause between starts (a worker that is never given up, PDF
 * native text, is not asked again every 1.5 s while its chunk stays away). */
export const WORKER_START_MAX_RETRY_MS = 30_000;
/** Failed starts before a worker is given up: the first and one retry. A
 * worker that started is not counted: a caller may replace a working one
 * (a timed-out request) as often as it needs. */
export const WORKER_START_ATTEMPTS = 2;

export interface ResilientWorkerStart {
  /**
   * A new worker from `factory`, or null while the retry's pause runs or
   * after the last start failed (the caller does the work itself then).
   * A factory that throws counts as a failed start.
   */
  create(factory: () => Worker): Worker | null;
  /** The worker said it has loaded: errors from now on are its work's. */
  started(worker: Worker): void;
  /**
   * The worker did not start (an error, a message before its "ready", no
   * word by its deadline). Reported, and the caller terminates it. "retry":
   * `create` makes the next one after the pause; "gave_up": it will not.
   */
  failed(worker: Worker | null, reason: string): "retry" | "gave_up";
  /** No start is left. */
  readonly gaveUp: boolean;
}

export interface ResilientWorkerOptions {
  retryDelayMs?: number;
  maxRetryDelayMs?: number;
  attempts?: number;
  /** For tests: the clock the retry's pause is read on. */
  now?: () => number;
}

export function startResilientWorker(name: string, options: ResilientWorkerOptions = {}): ResilientWorkerStart {
  const retryDelayMs = options.retryDelayMs ?? WORKER_START_RETRY_MS;
  const maxRetryDelayMs = options.maxRetryDelayMs ?? WORKER_START_MAX_RETRY_MS;
  const attempts = options.attempts ?? WORKER_START_ATTEMPTS;
  const now = options.now ?? (() => Date.now());
  let failures = 0;
  let notBefore = 0;
  let gaveUp = false;
  const ready = new WeakSet<Worker>();
  const lastError = new WeakMap<Worker, string>();

  const failed = (worker: Worker | null, reason: string): "retry" | "gave_up" => {
    if (gaveUp) return "gave_up";
    if (worker && ready.has(worker)) return "gave_up";
    const detail = worker ? lastError.get(worker) : undefined;
    failures += 1;
    const error = new Error(`${name} did not start (attempt ${failures} of ${attempts}): ${reason}${detail ? ` — ${detail}` : ""}`);
    error.name = "WorkerStartFailed";
    reportClientError(error);
    if (failures >= attempts) {
      gaveUp = true;
      return "gave_up";
    }
    // 1.5 s, 3 s, 6 s, … up to 30 s: doubled after each failed start.
    notBefore = now() + Math.min(retryDelayMs * 2 ** (failures - 1), maxRetryDelayMs);
    return "retry";
  };

  return {
    create(factory) {
      if (gaveUp || now() < notBefore) return null;
      let worker: Worker;
      try {
        worker = factory();
      } catch (error) {
        failed(null, `could not be made: ${String(error).slice(0, 200)}`);
        return null;
      }
      worker.addEventListener?.("error", (event: ErrorEvent) => {
        if (ready.has(worker)) return;
        // Handled here: a worker that did not start is reported as one
        // (failed), not as the page's own uncaught error.
        event.preventDefault();
        const where = event.filename ? ` at ${event.filename}` : "";
        lastError.set(worker, `${String(event.message ?? "error").slice(0, 300)}${where}`.slice(0, 400));
      });
      return worker;
    },
    started(worker) {
      ready.add(worker);
    },
    failed,
    get gaveUp() {
      return gaveUp;
    },
  };
}
