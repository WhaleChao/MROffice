/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BackgroundCalculationPlanRequest, BackgroundCalculationPlanResult } from "./background-calculation-plans";
import { worksheetFormulaResultKey, type WorksheetCalculationToken } from "./calculation-mode";
import { BackgroundCalculationMirror } from "./background-calculation-mirror";
import { BackgroundCalculationCancelledError, BackgroundCalculationUnavailableError, sameCalculationToken, type CalculationUnavailableReason, type BackgroundCalculationSnapshot, type BackgroundCalculationDelta, type BackgroundCalculationRequest, type BackgroundCalculationResult, type BackgroundCalculationMessage, type BackgroundCalculationResponse } from "./background-calculation-protocol";
import { RemoteCalculationWorker } from "./remote-calculation-worker";
import type { EngineValue } from "./formula-engine";

/** Narrow worker seam also supports deterministic transport tests. */
export interface BackgroundCalculationWorker {
  postMessage(message: BackgroundCalculationMessage): void;
  terminate(): void;
  onmessage: ((event: MessageEvent<BackgroundCalculationResponse>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  /**
   * `cancel` is posted and the worker kept (the server's calculation
   * session keeps its warm memo); any other worker is terminated instead.
   */
  readonly cancelInPlace?: boolean;
  /** Its own silent-request bound, when the client was given none. */
  readonly timeoutMs?: number;
  /** It would refuse this workbook unasked, as too large, until a reset. */
  readonly refusesWorkbook?: boolean;
}
/** What a worker factory may read: the client mirror's current snapshot. */
export interface BackgroundCalculationWorkerContext {
  exportSnapshot(): BackgroundCalculationSnapshot;
}
export interface BackgroundCalculationClientOptions {
  /** Existing worksheet authority owns this token; the client never advances it. */
  getCurrentToken(): WorksheetCalculationToken | null;
  createWorker?: (context: BackgroundCalculationWorkerContext) => BackgroundCalculationWorker | null;
  /** Maximum silent worker request time before the request fails as unavailable. */
  workerTimeoutMs?: number;
}
interface Pending {
  readonly id: number;
  readonly request: BackgroundCalculationRequest;
  readonly resolve: (result: BackgroundCalculationResult) => void;
  readonly reject: (error: Error) => void;
}
interface PendingPlan {
  readonly id: number;
  readonly request: BackgroundCalculationPlanRequest;
  readonly resolve: (result: BackgroundCalculationPlanResult) => void;
  readonly reject: (error: Error) => void;
}
export function backgroundCalculationCapability(snapshot: BackgroundCalculationSnapshot): { ok: true } | { ok: false; reason: "uncloneable-input" } {
  try { structuredClone(snapshot); return { ok: true }; }
  catch { return { ok: false, reason: "uncloneable-input" }; }
}


function defaultWorker(context: BackgroundCalculationWorkerContext): BackgroundCalculationWorker | null {
  return new RemoteCalculationWorker(context);
}

/**
 * One mounted workbook: reset once, apply accepted deltas thereafter. Await
 * calculate before export/copy/sort/validation/save consumes that revision.
 * Rejection is a barrier failure, never permission to consume older results.
 * This class has no model writer. Caller owns projection publication.
 */
export class BackgroundCalculationClient {
  readonly #options: BackgroundCalculationClientOptions;
  readonly #mirror = new BackgroundCalculationMirror();
  #worker: BackgroundCalculationWorker | null = null;
  #workerDisabled = false;
  #workerInitialized = false;
  #pending: Pending | null = null;
  #pendingPlan: PendingPlan | null = null;
  #sequence = 0;
  #disposed = false;
  readonly #workerTimeoutMs: number;
  #workerTimer: ReturnType<typeof setTimeout> | null = null;
  readonly #workerContext: BackgroundCalculationWorkerContext = { exportSnapshot: () => this.#mirror.exportSnapshot() };
  constructor(options: BackgroundCalculationClientOptions) {
    this.#options = options;
    this.#workerTimeoutMs = options.workerTimeoutMs ?? 30_000;
    if (!Number.isSafeInteger(this.#workerTimeoutMs) || this.#workerTimeoutMs < 1 || this.#workerTimeoutMs > 2_147_483_647) throw new TypeError("Invalid calculation worker timeout.");
  }
  /** See `BackgroundCalculationWorker.refusesWorkbook`. */
  get refusesWorkbook(): boolean {
    return this.#worker?.refusesWorkbook === true;
  }
  reset(snapshot: BackgroundCalculationSnapshot): void {
    this.#assertOpen(); this.#mirror.reset(snapshot); this.cancel();
    if (this.#worker) {
      try { this.#worker.postMessage({ kind: "reset", snapshot }); this.#workerInitialized = true; }
      catch { this.#disableWorker(); }
    }
  }
  applyDelta(delta: BackgroundCalculationDelta): void {
    this.#assertOpen();
    // Validate atomically before disturbing the current projection.
    this.#mirror.applyDelta(delta); this.cancel();
    if (this.#worker && this.#workerInitialized) {
      try { this.#worker.postMessage({ kind: "delta", delta }); }
      catch { this.#disableWorker(); }
    }
  }
  calculate(request: BackgroundCalculationRequest): Promise<BackgroundCalculationResult> {
    this.#assertOpen();
    const current = this.#options.getCurrentToken();
    if (!current || !sameCalculationToken(current, request.token)) return Promise.reject(new BackgroundCalculationCancelledError());
    this.cancel();
    // A frozen coordinate list (the host's, unchanged since its formulas
    // last moved) is already the client's own: only the rest is copied.
    const { coordinates, ...rest } = request;
    const owned: BackgroundCalculationRequest = Object.isFrozen(coordinates) && coordinates.every(Object.isFrozen)
      ? { ...structuredClone(rest), coordinates }
      : structuredClone(request);
    // The mirror holds what the session is asked to hold: a session opened
    // again from it keeps the PivotTables (a refused one fails this request).
    if (owned.pivots !== undefined) {
      try { this.#mirror.replacePivots(owned.pivots); }
      catch (error) { return Promise.reject(error instanceof Error ? error : new Error("Invalid PivotTables.")); }
    }
    return new Promise((resolve, reject) => {
      const pending: Pending = { id: ++this.#sequence, request: owned, resolve, reject };
      this.#pending = pending;
      try {
        if (!this.#worker && !this.#workerDisabled) {
          this.#worker = (this.#options.createWorker ?? defaultWorker)(this.#workerContext);
          if (this.#worker) this.#bindWorker(this.#worker); else this.#workerDisabled = true;
        }
        if (this.#worker) {
          this.#armWorkerTimeout(pending);
          if (!this.#workerInitialized) {
            this.#worker.postMessage({ kind: "reset", snapshot: this.#mirror.exportSnapshot() });
            this.#workerInitialized = true;
          }
          this.#worker.postMessage({ kind: "calculate", requestId: pending.id, request: owned });
        } else this.#fail(pending);
      } catch { this.#disableWorker(); this.#fail(pending); }
    });
  }
  /** Run after the exact calculation barrier; the existing command alone applies the plan. */
  plan(request: BackgroundCalculationPlanRequest): Promise<BackgroundCalculationPlanResult> {
    this.#assertOpen();
    const current = this.#options.getCurrentToken();
    if (!current || !sameCalculationToken(current, request.token)) return Promise.reject(new BackgroundCalculationCancelledError());
    this.cancel();
    const owned = structuredClone(request);
    return new Promise((resolve, reject) => {
      const pending: PendingPlan = { id: ++this.#sequence, request: owned, resolve, reject };
      this.#pendingPlan = pending;
      try {
        if (!this.#worker && !this.#workerDisabled) {
          this.#worker = (this.#options.createWorker ?? defaultWorker)(this.#workerContext);
          if (this.#worker) this.#bindWorker(this.#worker); else this.#workerDisabled = true;
        }
        if (this.#worker) {
          this.#armWorkerTimeout(pending);
          this.#worker.postMessage({ kind: "plan", requestId: pending.id, request: owned });
        } else this.#failPlan(pending);
      } catch { this.#disableWorker(); this.#failPlan(pending); }
    });
  }
  /**
   * What-If Analysis (Goal Seek, Data Table) against the calculation
   * session, which holds this client's current workbook (C2). Null when this
   * client's worker is not the server's session or the server cannot answer:
   * the caller then computes as before.
   */
  async whatIf(input: Parameters<RemoteCalculationWorker["whatIf"]>[0], signal?: AbortSignal): Promise<unknown | null> {
    this.#assertOpen();
    if (!this.#worker && !this.#workerDisabled) {
      this.#worker = (this.#options.createWorker ?? defaultWorker)(this.#workerContext);
      if (this.#worker) this.#bindWorker(this.#worker); else this.#workerDisabled = true;
    }
    const worker = this.#worker;
    if (!(worker instanceof RemoteCalculationWorker)) return null;
    if (!this.#workerInitialized) {
      worker.postMessage({ kind: "reset", snapshot: this.#mirror.exportSnapshot() });
      this.#workerInitialized = true;
    }
    return worker.whatIf(input, signal);
  }
  /** Expressions typed at one cell, against the calculation session (see
   * RemoteCalculationWorker.evaluate). Null without the server's session. */
  async evaluate(input: Parameters<RemoteCalculationWorker["evaluate"]>[0], signal?: AbortSignal): Promise<Array<EngineValue | undefined> | null> {
    this.#assertOpen();
    if (!this.#worker && !this.#workerDisabled) {
      this.#worker = (this.#options.createWorker ?? defaultWorker)(this.#workerContext);
      if (this.#worker) this.#bindWorker(this.#worker); else this.#workerDisabled = true;
    }
    const worker = this.#worker;
    if (!(worker instanceof RemoteCalculationWorker)) return null;
    if (!this.#workerInitialized) {
      worker.postMessage({ kind: "reset", snapshot: this.#mirror.exportSnapshot() });
      this.#workerInitialized = true;
    }
    return worker.evaluate(input, signal);
  }
  /** Rule formulas over their ranges, against the calculation session (see
   * RemoteCalculationWorker.evaluateRules). Null without the server's session. */
  async evaluateRules(input: Parameters<RemoteCalculationWorker["evaluateRules"]>[0], signal?: AbortSignal): Promise<Map<string, ReadonlyMap<string, EngineValue>> | null> {
    this.#assertOpen();
    if (!this.#worker && !this.#workerDisabled) {
      this.#worker = (this.#options.createWorker ?? defaultWorker)(this.#workerContext);
      if (this.#worker) this.#bindWorker(this.#worker); else this.#workerDisabled = true;
    }
    const worker = this.#worker;
    if (!(worker instanceof RemoteCalculationWorker)) return null;
    if (!this.#workerInitialized) {
      worker.postMessage({ kind: "reset", snapshot: this.#mirror.exportSnapshot() });
      this.#workerInitialized = true;
    }
    return worker.evaluateRules(input, signal);
  }
  cancel(): void {
    this.#clearWorkerTimeout();
    const pending = this.#pending; const plan = this.#pendingPlan;
    this.#pending = null; this.#pendingPlan = null;
    pending?.reject(new BackgroundCalculationCancelledError());
    plan?.reject(new BackgroundCalculationCancelledError());
    // A queued cancel cannot interrupt an atomic formula. Terminate its worker
    // immediately; the next request recreates it from the current sparse mirror.
    // The server's session is kept: its next call supersedes the old one.
    if (pending || plan) {
      const worker = this.#worker;
      if (worker?.cancelInPlace) {
        try { worker.postMessage({ kind: "cancel", requestId: (pending ?? plan)!.id }); }
        catch { this.#stopWorker(false); }
      } else this.#stopWorker(false);
    }
  }
  dispose(): void {
    if (this.#disposed) return;
    this.cancel(); this.#disableWorker(); this.#mirror.clear(); this.#disposed = true;
  }
  #assertOpen(): void { if (this.#disposed) throw new Error("Calculation client disposed."); }
  #clearWorkerTimeout(): void {
    if (this.#workerTimer !== null) clearTimeout(this.#workerTimer);
    this.#workerTimer = null;
  }
  #armWorkerTimeout(pending: Pending | PendingPlan): void {
    this.#clearWorkerTimeout();
    const worker = this.#worker;
    const timer = setTimeout(() => {
      if (this.#workerTimer !== timer) return;
      this.#workerTimer = null;
      if (worker !== this.#worker || (pending !== this.#pending && pending !== this.#pendingPlan)) return;
      // A silent server: the next request opens a fresh session.
      this.#stopWorker(false);
      const current = this.#options.getCurrentToken();
      if (!current || !sameCalculationToken(current, pending.request.token)) { this.cancel(); return; }
      if (pending === this.#pending) this.#fail(this.#pending);
      else if (pending === this.#pendingPlan) this.#failPlan(this.#pendingPlan);
    }, this.#options.workerTimeoutMs === undefined ? Math.max(this.#workerTimeoutMs, worker?.timeoutMs ?? 0) : this.#workerTimeoutMs);
    this.#workerTimer = timer;
  }
  #disableWorker(): void { this.#stopWorker(true); }
  #stopWorker(disable: boolean): void {
    this.#clearWorkerTimeout();
    const worker = this.#worker;
    this.#worker = null; this.#workerInitialized = false;
    if (disable) this.#workerDisabled = true;
    if (worker) { worker.onmessage = null; worker.onerror = null; worker.onmessageerror = null; worker.terminate(); }
  }
  #bindWorker(worker: BackgroundCalculationWorker): void {
    worker.onmessage = (event) => {
      if (worker !== this.#worker) return;
      const response = event.data;
      // A reply outside the protocol (no request id, an unknown kind) is not a
      // stale answer to an earlier request but a worker that is not ours, for
      // example a host's stub for another document kind. Waiting for the
      // timeout would stall calculation for its full length; fail closed now.
      const shape = response as { kind?: unknown; requestId?: unknown } | null;
      const known = shape?.kind === "error" || shape?.kind === "result" || shape?.kind === "plan-result";
      const addressed = typeof shape?.requestId === "number" || (shape?.kind === "error" && shape.requestId === null);
      if (!known || !addressed) {
        this.#disableWorker();
        if (this.#pending) this.#fail(this.#pending);
        if (this.#pendingPlan) this.#failPlan(this.#pendingPlan);
        return;
      }
      const plan = this.#pendingPlan;
      if (plan) {
        if (response.kind === "error") {
          if (response.requestId !== null && response.requestId !== plan.id) return;
          // The server could not answer: the worker (and its back-off) stays.
          this.#failPlan(plan, response.reason); return;
        }
        if (response.requestId !== plan.id) return;
        if (response.kind !== "plan-result" || !sameCalculationToken(response.token, plan.request.token) || !this.#validPlanResult(plan, response.plan)) {
          this.#stopWorker(false); this.#failPlan(plan); return;
        }
        this.#finishPlan(plan, { token: response.token, plan: response.plan, execution: "worker" }); return;
      }
      const pending = this.#pending;
      if (!pending) return;
      if (response.kind === "error") {
        if (response.requestId !== null && response.requestId !== pending.id) return;
        this.#fail(pending, response.reason); return;
      }
      if (response.requestId !== pending.id) return;
      if (response.kind !== "result") { this.#stopWorker(false); this.#fail(pending); return; }
      const expected = new Set(pending.request.coordinates.map((cell) => worksheetFormulaResultKey(cell.sheetName, cell.row, cell.col)));
      if (!sameCalculationToken(response.token, pending.request.token) || !(response.results instanceof Map) || response.results.size !== expected.size || [...expected].some((key) => !response.results.has(key) || response.results.get(key) === undefined)) {
        this.#stopWorker(false); this.#fail(pending); return;
      }
      this.#finish(pending, { token: response.token, results: response.results, execution: "worker", ...(response.ruleValues ? { ruleValues: response.ruleValues } : {}),
        ...(response.seq !== undefined ? { seq: response.seq } : {}), ...(response.pivotsId ? { pivotsId: response.pivotsId } : {}),
        ...(response.externalRequests?.length ? { externalRequests: response.externalRequests } : {}) });
    };
    const fail = () => {
      if (worker !== this.#worker) return;
      this.#stopWorker(false); if (this.#pending) this.#fail(this.#pending);
      if (this.#pendingPlan) this.#failPlan(this.#pendingPlan);
    };
    worker.onerror = fail; worker.onmessageerror = fail;
  }
  /** No answer for this request: there is no local evaluator to ask (C3). */
  #fail(pending: Pending, reason?: CalculationUnavailableReason): void {
    if (this.#pending !== pending) return;
    this.#clearWorkerTimeout();
    this.#pending = null;
    pending.reject(new BackgroundCalculationUnavailableError(reason));
  }
  #validPlanResult(pending: PendingPlan, plan: BackgroundCalculationPlanResult["plan"]): boolean {
    if (pending.request.kind === "validate-draft") {
      return plan.kind === "validation" && typeof plan.result?.ok === "boolean" &&
        (!plan.result.ok || !("failure" in plan.result));
    }
    if (pending.request.kind === "aggregate") {
      return plan.kind === "aggregate" && plan.values.length === pending.request.groups.length && plan.values.every((value, index) => value.id === (pending.request as Extract<BackgroundCalculationPlanRequest, { kind: "aggregate" }>).groups[index].id && (value.value === null || typeof value.value === "number"));
    }
    if (plan.kind !== "rows" || plan.sourceOrder.length !== pending.request.sourceRows.length) return false;
    const sources = new Set(pending.request.sourceRows); const ordered = new Set(plan.sourceOrder);
    return ordered.size === sources.size && plan.sourceOrder.every((row) => sources.has(row)) && new Set(plan.visibleSourceRows).size === plan.visibleSourceRows.length && plan.visibleSourceRows.every((row) => sources.has(row));
  }
  #failPlan(pending: PendingPlan, reason?: CalculationUnavailableReason): void {
    if (this.#pendingPlan !== pending) return;
    this.#clearWorkerTimeout();
    this.#pendingPlan = null;
    pending.reject(new BackgroundCalculationUnavailableError(reason));
  }
  #finishPlan(pending: PendingPlan, result: BackgroundCalculationPlanResult): void {
    if (this.#pendingPlan !== pending) return;
    this.#clearWorkerTimeout();
    this.#pendingPlan = null;
    const current = this.#options.getCurrentToken();
    if (!current || !sameCalculationToken(current, pending.request.token)) { pending.reject(new BackgroundCalculationCancelledError()); return; }
    pending.resolve(result);
  }
  #finish(pending: Pending, result: BackgroundCalculationResult): void {
    if (this.#pending !== pending) return;
    this.#clearWorkerTimeout();
    this.#pending = null;
    const current = this.#options.getCurrentToken();
    if (!current || !sameCalculationToken(current, pending.request.token)) { pending.reject(new BackgroundCalculationCancelledError()); return; }
    // Raw arrays and LAMBDA captures remain read-only after transport.
    pending.resolve(result);
  }
}
