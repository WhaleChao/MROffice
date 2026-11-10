"use client";































/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { callEngine, EngineCallCancelledError, EngineRefusedError, EngineUnavailableError, isEngineAbort } from "../engine-host/client";
import { EngineSession } from "../engine-host/session";
import type { SheetCalcCalculateInput, SheetCalcCalculateOutput, SheetCalcEnvironment, SheetCalcEvaluateInput, SheetCalcOpenInput, SheetCalcWhatIfInput } from "@/lib/engine-host/sheet-calc-types";
import type {
  BackgroundCalculationDelta,
  BackgroundCalculationMessage,
  BackgroundCalculationRequest,
  BackgroundCalculationResponse,
  BackgroundCalculationSnapshot,
  CalculationRuleRequest,
} from "./background-calculation-protocol";
import { calculationRuleKey, type CalculationUnavailableReason } from "./background-calculation-protocol";
import type { BackgroundCalculationWorker } from "./background-calculation-client";
import { worksheetFormulaResultKey } from "./calculation-mode";
import { decodeCalcValue, encodeCalcCoordinates, encodeCalcValue, type CalcCoordinates } from "./calc-wire";
import type { EngineValue } from "./formula-engine";
import type { CalculationExternalDatum } from "./external-data";

/** After the server failed a request, how long new requests fail at once. */
export const REMOTE_CALCULATION_BACKOFF_MS = 30_000;
/** Queued edits past which the session is re-opened from the snapshot. */
const MAX_QUEUED_DELTAS = 200;
const MAX_QUEUED_DELTA_CHARS = 4 * 1024 * 1024;
/** The least one cell of a delta spells: `{"sheetId":"","row":0,"col":0,"raw":null}`. */
const MIN_DELTA_CELL_CHARS = 40;

let remoteDownUntil = 0;

type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;

/**
 * How long new requests fail at once after a failure: a busy pool or a
 * used-up calculation budget says when to come back (`Retry-After`); an
 * outage waits the full back-off. Only the server being unable to answer
 * backs off: a refusal of one request (its input, its size) is that
 * request's answer, and the next request goes to the server as usual.
 */
function backoffFor(error: unknown): number {
  if (!(error instanceof EngineUnavailableError)) return 0;
  if (error.reason === "busy" || error.reason === "rate_limited") {
    return Math.min(REMOTE_CALCULATION_BACKOFF_MS, Math.max(1_000, error.retryAfterMs ?? 1_000));
  }
  return REMOTE_CALCULATION_BACKOFF_MS;
}

/** Refusals that mean "this workbook is over the calculation's budget". */
const BUDGET_REFUSALS = new Set(["budget_exceeded", "snapshot_too_large", "selection_too_large", "too_large"]);

/** Why a failed request was not answered, as the status bar names it (C4). */
export function calculationUnavailableReason(error: unknown): CalculationUnavailableReason {
  if (typeof navigator !== "undefined" && navigator.onLine === false) return "offline";
  if (error instanceof EngineUnavailableError) {
    if (error.reason === "offline") return "offline";
    if (error.reason === "busy" || error.reason === "rate_limited") return "busy";
    return "unavailable";
  }
  if (error instanceof EngineRefusedError && BUDGET_REFUSALS.has(error.reason)) return "budget";
  return "unavailable";
}

/** Why requests fail at once during the back-off. */
let remoteDownReason: CalculationUnavailableReason = "unavailable";

function backOff(error: unknown): void {
  const wait = backoffFor(error);
  if (wait > 0) {
    remoteDownUntil = Math.max(remoteDownUntil, Date.now() + wait);
    remoteDownReason = calculationUnavailableReason(error);
  }
}

/**
 * The session's workbook is not the one this worker's edits describe (it
 * asks for a reset, or holds another revision than the request names):
 * open it again from the client's snapshot and ask once more. Without a
 * local evaluator (C3) this is the only recovery, so a session that drifted
 * never leaves the request unanswered.
 */
function sessionDiverged(error: unknown): boolean {
  return error instanceof EngineRefusedError && (error.reason === "reset_required" || error.reason === "stale_token");
}

/** Forget a back-off: the user asked to calculate again now (and tests). */
export function resetRemoteCalculationBackoff(): void {
  remoteDownUntil = 0;
}

export interface RemoteCalculationContext {
  /** The client mirror's current snapshot (reset + every delta since). */
  exportSnapshot(): BackgroundCalculationSnapshot;
}

export function calculationEnvironment(): SheetCalcEnvironment {
  let timeZone: string | undefined;
  let locale: string | undefined;
  try {
    timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    locale = new Intl.Collator().resolvedOptions().locale;
  } catch {
    /* the server's defaults */
  }
  return {
    epochMs: Date.now(),
    utcOffsetMinutes: -new Date().getTimezoneOffset(),
    ...(timeZone ? { timeZone } : {}),
    ...(locale ? { locale } : {}),
  };
}

const cellKey = (sheetId: string, row: number, col: number) => `${sheetId}\u0000${row}:${col}`;

/** The data's fingerprint, once per list (the host hands the same list until it changes). */
const externalPrints = new WeakMap<readonly CalculationExternalDatum[], string>();
function externalFingerprint(data: readonly CalculationExternalDatum[]): string {
  let print = externalPrints.get(data);
  if (print === undefined) externalPrints.set(data, (print = JSON.stringify(data)));
  return print;
}

export class RemoteCalculationWorker implements BackgroundCalculationWorker {
  /** `cancel` keeps this worker (and the server's memo). */
  readonly cancelInPlace = true;
  /** The server's own deadline (30 s) plus the round trip. */
  readonly timeoutMs = 35_000;
  onmessage: ((event: MessageEvent<BackgroundCalculationResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;

  readonly #context: RemoteCalculationContext;
  readonly #session: EngineSession<"sheet-calc">;
  #deltas: BackgroundCalculationDelta[] = [];
  #deltaChars = 0;
  /** Every value the server sent this worker, by result key. */
  #values = new Map<string, EngineValue>();
  #ack = 0;
  /** The session's formula cells (`sheetId\0row:col`), as the mirror holds them. */
  #formulas = new Set<string>();
  #formulasBySheet = new Map<string, number>();
  /** Bumped whenever the formula cells above change. */
  #formulasVersion = 0;
  /** The last frozen coordinate list found to be exactly the session's formulas. */
  #allChecked: { list: readonly unknown[]; target: string | null; version: number } | null = null;
  /** Result keys of a frozen coordinate list, built once for it. */
  #resultKeys = new WeakMap<readonly unknown[], readonly string[]>();
  #inFlight: { requestId: number; controller: AbortController } | null = null;
  /** The WEBSERVICE / IMPORTRANGE data the open session was last handed
   * (its fingerprint): it travels again only when it changes or the session
   * is opened anew, not with every keystroke's calculation. */
  #externalSent: string | null = null;
  #terminated = false;
  /** The snapshot's cells, and how many a session open was refused for as
   * too large: a workbook at least that large is not offered again (a sorted
   * 728,000-cell sheet sent its 14.7 MB snapshot, and was refused, on every
   * calculation: four times while its save waited). */
  #cellCount = 0;
  #refusedCells: number | null = null;
  readonly #onPageHide = () => this.#session.reset({ keepalive: true });

  constructor(context: RemoteCalculationContext) {
    this.#context = context;
    this.#session = new EngineSession("sheet-calc", {
      snapshot: (): SheetCalcOpenInput => ({ snapshot: this.#context.exportSnapshot(), environment: calculationEnvironment() }),
      // The page: the server keeps an owner's few sessions a page, and this
      // tab its newest one (`EngineSession`).
      scopeOf: (input) => {
        const pageId = (input.snapshot as { token?: { pageId?: unknown } } | null)?.token?.pageId;
        return typeof pageId === "string" && /^[A-Za-z0-9:._-]{1,128}$/.test(pageId) ? pageId : undefined;
      },
      onOpened: () => {
        // The snapshot holds every edit so far; this worker starts from nothing.
        this.#deltas = [];
        this.#deltaChars = 0;
        this.#values = new Map();
        this.#ack = 0;
        this.#externalSent = null;
      },
    });
    if (typeof window !== "undefined") window.addEventListener("pagehide", this.#onPageHide);
  }

  /** This workbook's session was refused as too large and is not offered
   * again until a new snapshot says it is smaller (`#calculate`). */
  get refusesWorkbook(): boolean {
    return !this.#session.id && this.#refusedCells !== null && this.#cellCount >= this.#refusedCells;
  }

  /** Session statistics for diagnostics and tests. */
  get reopens(): number {
    return this.#session.reopens;
  }

  postMessage(message: BackgroundCalculationMessage): void {
    if (this.#terminated) return;
    if (message.kind === "reset") {
      this.#abortInFlight();
      this.#session.reset();
      this.#deltas = [];
      this.#deltaChars = 0;
      this.#trackSnapshot(message.snapshot);
      return;
    }
    if (message.kind === "delta") {
      this.#trackDelta(message.delta);
      this.#deltas.push(message.delta);
      // Measured only when it can fit: a sort's edit of 730,000 cells was
      // 0.4 s of the frame after it spelled out to be found too long.
      this.#deltaChars += message.delta.cells.length * MIN_DELTA_CELL_CHARS > MAX_QUEUED_DELTA_CHARS
        ? MAX_QUEUED_DELTA_CHARS + 1
        : JSON.stringify(message.delta).length;
      if (this.#deltas.length > MAX_QUEUED_DELTAS || this.#deltaChars > MAX_QUEUED_DELTA_CHARS) {
        // Cheaper to open again from the snapshot than to replay.
        this.#session.reset();
        this.#deltas = [];
        this.#deltaChars = 0;
      }
      return;
    }
    if (message.kind === "cancel") {
      if (this.#inFlight?.requestId === message.requestId) this.#abortInFlight();
      return;
    }
    this.#abortInFlight();
    const controller = new AbortController();
    this.#inFlight = { requestId: message.requestId, controller };
    if (Date.now() < remoteDownUntil) {
      this.#unavailable(message.requestId, remoteDownReason);
      return;
    }
    if (message.kind === "plan") void this.#plan(message.requestId, message.request, controller.signal);
    else void this.#calculate(message.requestId, message.request, controller.signal);
  }

  terminate(): void {
    if (this.#terminated) return;
    this.#terminated = true;
    this.#abortInFlight();
    if (typeof window !== "undefined") window.removeEventListener("pagehide", this.#onPageHide);
    this.#session.close({ keepalive: true });
  }

  #abortInFlight(): void {
    this.#inFlight?.controller.abort();
    this.#inFlight = null;
  }

  #current(requestId: number, signal: AbortSignal): boolean {
    return !this.#terminated && !signal.aborted && this.#inFlight?.requestId === requestId;
  }

  #emit(response: BackgroundCalculationResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<BackgroundCalculationResponse>);
  }

  #trackSnapshot(snapshot: BackgroundCalculationSnapshot): void {
    this.#formulasVersion += 1;
    this.#formulas = new Set();
    this.#formulasBySheet = new Map();
    this.#cellCount = 0;
    for (const sheet of snapshot.sheets) {
      this.#cellCount += sheet.cells.length;
      for (const [key, raw] of sheet.cells) {
        if (!raw.trim().startsWith("=")) continue;
        const [row, col] = key.split(":").map(Number);
        this.#formulas.add(cellKey(sheet.id, row, col));
        this.#formulasBySheet.set(sheet.id, (this.#formulasBySheet.get(sheet.id) ?? 0) + 1);
      }
    }
  }

  #trackDelta(delta: BackgroundCalculationDelta): void {
    for (const cell of delta.cells) {
      const key = cellKey(cell.sheetId, cell.row, cell.col);
      const formula = cell.raw?.trim().startsWith("=") === true;
      const had = this.#formulas.has(key);
      if (formula !== had) this.#formulasVersion += 1;
      if (formula && !had) {
        this.#formulas.add(key);
        this.#formulasBySheet.set(cell.sheetId, (this.#formulasBySheet.get(cell.sheetId) ?? 0) + 1);
      } else if (!formula && had) {
        this.#formulas.delete(key);
        this.#formulasBySheet.set(cell.sheetId, (this.#formulasBySheet.get(cell.sheetId) ?? 1) - 1);
      }
    }
  }

  /** `{all: true}` when the request names exactly the session's formulas. */
  #coordinates(request: BackgroundCalculationRequest): CalcCoordinates {
    const target = request.token.targetSheetId;
    const expected = target === null ? this.#formulas.size : (this.#formulasBySheet.get(target) ?? 0);
    const frozen = Object.isFrozen(request.coordinates);
    const checked = this.#allChecked;
    if (frozen && checked && checked.list === request.coordinates && checked.target === target && checked.version === this.#formulasVersion) return { all: true };
    if (request.coordinates.length === expected) {
      const seen = new Set<string>();
      const all = request.coordinates.every((cell) => {
        const key = cellKey(cell.sheetId, cell.row, cell.col);
        if (seen.has(key) || !this.#formulas.has(key) || (target !== null && cell.sheetId !== target)) return false;
        seen.add(key);
        return true;
      });
      if (all) {
        if (frozen) this.#allChecked = { list: request.coordinates, target, version: this.#formulasVersion };
        return { all: true };
      }
    }
    return encodeCalcCoordinates(request.coordinates);
  }

  async #calculate(requestId: number, request: BackgroundCalculationRequest, signal: AbortSignal, reopened = false): Promise<void> {
    let opening = false;
    try {
      if (!this.#session.id && this.#refusedCells !== null && this.#cellCount >= this.#refusedCells) {
        throw new EngineRefusedError("too_large");
      }
      opening = !this.#session.id;
      await this.#session.open(signal);
      opening = false;
      if (!this.#current(requestId, signal)) return;
      const external = request.externalData;
      const externalPrint = external !== undefined ? externalFingerprint(external) : null;
      const build = (ack: number, reopenedSession = false): SheetCalcCalculateInput => ({
        ...(this.#deltas.length ? { deltas: [...this.#deltas] } : {}),
        request: {
          token: request.token,
          coordinates: this.#coordinates(request),
          ...(request.formulaResultSeed ? { formulaResultSeed: encodeCalcValue(request.formulaResultSeed) } : {}),
          ...(request.maxComputedCells !== undefined ? { maxComputedCells: request.maxComputedCells } : {}),
          ...(request.rules?.length ? { rules: request.rules } : {}),
          ...(request.pivots !== undefined ? { pivots: request.pivots } : {}),
          ...(request.webImages !== undefined ? { webImages: request.webImages } : {}),
          ...(external !== undefined && (reopenedSession || externalPrint !== this.#externalSent) ? { externalData: external } : {}),
        },
        ack,
        environment: calculationEnvironment(),
      });
      let answer = await this.#session.call("calculate", build(this.#ack), {
        signal,
        // Re-opened from the current snapshot: no edits to replay, nothing held.
        retryInput: () => build(0, true),
      });
      if (!this.#current(requestId, signal)) return;
      let results = this.#apply(answer, request);
      if (!results) {
        // This worker lacks a value the server thinks it holds: ask for all.
        answer = await this.#session.call("calculate", build(-1), { signal, retryInput: () => build(0, true) });
        if (!this.#current(requestId, signal)) return;
        results = this.#apply(answer, request);
      }
      if (!results) throw new Error("Calculation session answered an incomplete result.");
      this.#inFlight = null;
      if (externalPrint !== null) this.#externalSent = externalPrint;
      let ruleValues: Map<string, ReadonlyMap<string, EngineValue>> | undefined;
      if (request.rules?.length && answer.rules) {
        ruleValues = new Map();
        for (const [index, rule] of request.rules.entries()) {
          const cells = answer.rules[index];
          if (cells) ruleValues.set(calculationRuleKey(rule), new Map(cells.map(([cell, value]) => [cell, decodeCalcValue(value) as EngineValue])));
        }
      }
      this.#emit({ kind: "result", requestId, token: request.token, results, ...(ruleValues ? { ruleValues } : {}),
        ...(answer.seq !== null ? { seq: answer.seq } : {}), ...(answer.pivotsId ? { pivotsId: answer.pivotsId } : {}),
        ...(answer.externalRequests?.length ? { externalRequests: answer.externalRequests } : {}) });
    } catch (error) {
      if (opening && error instanceof EngineRefusedError && (error.reason === "too_large" || error.reason === "snapshot_too_large")) {
        this.#refusedCells = this.#refusedCells === null ? this.#cellCount : Math.min(this.#refusedCells, this.#cellCount);
      }
      if (isEngineAbort(error) || error instanceof EngineCallCancelledError || !this.#current(requestId, signal)) return;
      if (sessionDiverged(error) && !reopened) {
        // The session's edits and ours diverged: open it again from the
        // client's snapshot and ask once more.
        this.#session.reset();
        return this.#calculate(requestId, request, signal, true);
      }
      backOff(error);
      this.#unavailable(requestId, calculationUnavailableReason(error));
    }
  }

  /**
   * What-If Analysis against the session's current workbook (Goal Seek, Data
   * Table; C2): the edits typed so far travel with it. Null when the server
   * cannot answer (the caller computes in the browser, C1/C2 only).
   */
  async whatIf(
    input: DistributiveOmit<SheetCalcWhatIfInput, "deltas" | "environment">,
    signal?: AbortSignal,
    reopened = false,
  ): Promise<unknown | null> {
    if (this.#terminated || Date.now() < remoteDownUntil) return null;
    try {
      await this.#session.open(signal);
      let newest: BackgroundCalculationDelta["token"] | null = null;
      const build = (): SheetCalcWhatIfInput => {
        newest = this.#deltas.at(-1)?.token ?? null;
        return { ...input, ...(this.#deltas.length ? { deltas: [...this.#deltas] } : {}), environment: calculationEnvironment() } as SheetCalcWhatIfInput;
      };
      const answer = await this.#session.call("whatIf", build(), { signal, retryInput: build });
      if (newest) this.#prune(newest);
      return decodeCalcValue(answer);
    } catch (error) {
      if (isEngineAbort(error)) throw error;
      if (sessionDiverged(error) && !reopened) {
        this.#session.reset();
        return this.whatIf(input, signal, true);
      }
      return null;
    }
  }

  /**
   * Expressions typed at one cell of the session's current workbook
   * (Evaluate Formula, Function Arguments; C2): each value, `undefined` where
   * the engine refuses one. Null when the server cannot answer.
   */
  async evaluate(
    input: Omit<SheetCalcEvaluateInput, "deltas" | "environment">,
    signal?: AbortSignal,
    reopened = false,
  ): Promise<Array<EngineValue | undefined> | null> {
    if (this.#terminated || Date.now() < remoteDownUntil) return null;
    try {
      await this.#session.open(signal);
      let newest: BackgroundCalculationDelta["token"] | null = null;
      const build = (): SheetCalcEvaluateInput => {
        newest = this.#deltas.at(-1)?.token ?? null;
        return { ...input, ...(this.#deltas.length ? { deltas: [...this.#deltas] } : {}), environment: calculationEnvironment() };
      };
      const answer = await this.#session.call("evaluate", build(), { signal, retryInput: build });
      if (newest) this.#prune(newest);
      return answer.values.map((entry) => entry === null ? undefined : decodeCalcValue(entry.value) as EngineValue);
    } catch (error) {
      if (isEngineAbort(error)) throw error;
      if (sessionDiverged(error) && !reopened) {
        this.#session.reset();
        return this.evaluate(input, signal, true);
      }
      return null;
    }
  }

  /**
   * Rule formulas over their ranges against the session's current workbook
   * (advanced filter criteria; C2), keyed like calculate's rule values (FALSE
   * omitted). Null when the server cannot answer; a rule it declined is absent.
   */
  async evaluateRules(
    input: { readonly token: SheetCalcEvaluateInput["token"]; readonly rules: readonly CalculationRuleRequest[] },
    signal?: AbortSignal,
    reopened = false,
  ): Promise<Map<string, ReadonlyMap<string, EngineValue>> | null> {
    if (this.#terminated || Date.now() < remoteDownUntil || input.rules.length === 0) return null;
    const first = input.rules[0];
    try {
      await this.#session.open(signal);
      let newest: BackgroundCalculationDelta["token"] | null = null;
      const build = (): SheetCalcEvaluateInput => {
        newest = this.#deltas.at(-1)?.token ?? null;
        return { token: input.token, sheetId: first.sheetId, row: first.anchorRow, col: first.anchorCol, expressions: [], rules: input.rules,
          ...(this.#deltas.length ? { deltas: [...this.#deltas] } : {}), environment: calculationEnvironment() };
      };
      const answer = await this.#session.call("evaluate", build(), { signal, retryInput: build });
      if (newest) this.#prune(newest);
      const values = new Map<string, ReadonlyMap<string, EngineValue>>();
      for (const [index, rule] of input.rules.entries()) {
        const cells = answer.rules?.[index];
        if (cells) values.set(calculationRuleKey(rule), new Map(cells.map(([cell, value]) => [cell, decodeCalcValue(value) as EngineValue])));
      }
      return values;
    } catch (error) {
      if (isEngineAbort(error)) throw error;
      if (sessionDiverged(error) && !reopened) {
        this.#session.reset();
        return this.evaluateRules(input, signal, true);
      }
      return null;
    }
  }

  /** Edits the session has applied (an answer came back for them) leave the queue. */
  #prune(token: { pageId: string; loadEpoch: number; generation: number }): void {
    this.#deltas = this.#deltas.filter((delta) => delta.token.pageId !== token.pageId ||
      delta.token.loadEpoch !== token.loadEpoch || delta.token.generation > token.generation);
    this.#deltaChars = this.#deltas.reduce((total, delta) => total + JSON.stringify(delta).length, 0);
  }

  /** Merges an answer into this worker's values; the request's full result
   * map, or null when a value the server did not resend is missing here. */
  #apply(
    answer: SheetCalcCalculateOutput,
    request: BackgroundCalculationRequest,
  ): Map<string, EngineValue> | null {
    // The edits that travelled are in the session now.
    this.#prune(answer.token);
    if (answer.full) this.#values = new Map();
    for (const [key, value] of answer.changed) this.#values.set(key, decodeCalcValue(value) as EngineValue);
    if (answer.seq !== null) this.#ack = answer.seq;
    const results = new Map<string, EngineValue>();
    const coordinates = request.coordinates;
    let keys = this.#resultKeys.get(coordinates);
    if (!keys) {
      keys = coordinates.map((cell) => worksheetFormulaResultKey(cell.sheetName, cell.row, cell.col));
      if (Object.isFrozen(coordinates)) this.#resultKeys.set(coordinates, keys);
    }
    for (const key of keys) {
      const value = this.#values.get(key);
      if (value === undefined) return null;
      results.set(key, value);
    }
    return results;
  }

  async #plan(requestId: number, request: Extract<BackgroundCalculationMessage, { kind: "plan" }>["request"], signal: AbortSignal): Promise<void> {
    try {
      const answer = await callEngine("sheet-calc", "plan", { request: encodeCalcValue(request) }, { signal });
      if (!this.#current(requestId, signal)) return;
      this.#inFlight = null;
      this.#emit({ kind: "plan-result", requestId, token: request.token, plan: decodeCalcValue(answer.plan) as never });
    } catch (error) {
      if (isEngineAbort(error) || !this.#current(requestId, signal)) return;
      backOff(error);
      this.#unavailable(requestId, calculationUnavailableReason(error));
    }
  }

  /** The server cannot answer this request (outage / back-off). */
  #unavailable(requestId: number, reason: CalculationUnavailableReason): void {
    if (this.#inFlight?.requestId === requestId) this.#inFlight = null;
    this.#emit({ kind: "error", requestId, message: "Calculation unavailable.", reason });
  }
}
