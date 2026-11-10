/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Formula results of workbooks the calculation session does not hold (C2):
 * a Copilot preview's planned workbook, step by step. The planners stay
 * synchronous (the editors' shared Copilot surface asks them synchronously);
 * they read a workbook's results here, and a miss asks the server's stateless
 * `sheet-calc.results` once and reports "pending" until it lands, when the
 * owner re-plans (`onSettled`). Keys are content fingerprints memoized per
 * sheet object, so an unchanged sheet is never serialized twice.
 */

import { hash53 } from "../content-hash";
import { callEngine } from "../engine-host/client";
import { calculationRuleKey, type BackgroundCalculationSnapshot, type CalculationRuleRequest } from "./background-calculation-protocol";
import { decodeCalcValue } from "./calc-wire";
import { MAX_PIVOT_REPLACEMENTS, type CalculationPivots } from "./calculation-pivots";
import { isArrayValue, type EngineValue } from "./formula-engine";
import { calculationEnvironment } from "./remote-calculation-worker";

export type WorkbookValuesLookup =
  | { readonly kind: "values"; readonly results: ReadonlyMap<string, EngineValue> }
  /** Asked for; the owner is told when it lands. */
  | { readonly kind: "pending" }
  /** The server cannot answer this workbook: the caller computes as before (until C3). */
  | { readonly kind: "local" };

export interface WorkbookValuesSource {
  readonly sheets: readonly object[];
  readonly dateSystem?: unknown;
}

/** Rule formulas answered beside the results: per rule key
 * (calculationRuleKey), the "row:col" values that are not FALSE. */
export type WorkbookRuleValues = ReadonlyMap<string, ReadonlyMap<string, EngineValue>>;
export interface WorkbookResultsAnswer {
  readonly results: ReadonlyMap<string, EngineValue>;
  readonly ruleValues?: WorkbookRuleValues;
}

/** A stateless calculation of a snapshot: the server's (`requestWorkbookAnswer`),
 * or the engine host's own runtime when the server itself prints a workbook. */
export type WorkbookAnswerRequest = (snapshot: BackgroundCalculationSnapshot, signal?: AbortSignal,
  rules?: readonly CalculationRuleRequest[]) => Promise<WorkbookResultsAnswer | ReadonlyMap<string, EngineValue> | null>;

/** Every formula result of a snapshot from the server, and the rule formulas
 * asked beside them (a rule the server declines is absent); null when it
 * cannot answer. */
export async function requestWorkbookAnswer(snapshot: BackgroundCalculationSnapshot, signal?: AbortSignal,
  rules?: readonly CalculationRuleRequest[]): Promise<WorkbookResultsAnswer | null> {
  try {
    const answer = await callEngine("sheet-calc", "results", { snapshot, environment: calculationEnvironment(),
      ...(rules?.length ? { rules } : {}) }, { signal });
    const results = decodeCalcValue(answer.results);
    if (!(results instanceof Map)) return null;
    if (!rules?.length) return { results: results as ReadonlyMap<string, EngineValue> };
    const ruleValues = new Map<string, ReadonlyMap<string, EngineValue>>();
    const answered = Array.isArray(answer.rules) ? answer.rules as unknown[] : [];
    rules.forEach((rule, index) => {
      const cells = answered[index];
      if (Array.isArray(cells)) {
        ruleValues.set(calculationRuleKey(rule), new Map((cells as Array<[string, unknown]>).map(([cell, value]) => [cell, decodeCalcValue(value) as EngineValue])));
      }
    });
    return { results: results as ReadonlyMap<string, EngineValue>, ruleValues };
  } catch {
    return null;
  }
}

/** Every formula result of a snapshot from the server; null when it cannot answer. */
export async function requestWorkbookResults(snapshot: BackgroundCalculationSnapshot, signal?: AbortSignal): Promise<ReadonlyMap<string, EngineValue> | null> {
  return (await requestWorkbookAnswer(snapshot, signal))?.results ?? null;
}

/**
 * Every formula result of a snapshot whose PivotTables' reports are read
 * from its own answer (pivot option C, stateless): while the reports read
 * from an answer are not the ones it was calculated with, ask again with
 * those -- at most MAX_PIVOT_REPLACEMENTS times; a third change answers the
 * last results with `pivotStale` (a report reading another report's output).
 * Without `derive` (no PivotTable) this is one request. `rules` (rule
 * formulas, e.g. conditional formats a colour filter reads) are answered
 * with the last results.
 */
export async function requestSettledWorkbookResults(
  snapshot: BackgroundCalculationSnapshot,
  derive?: (results: ReadonlyMap<string, EngineValue>) => CalculationPivots | null,
  signal?: AbortSignal,
  request: WorkbookAnswerRequest = requestWorkbookAnswer,
  rules?: readonly CalculationRuleRequest[],
): Promise<{ results: ReadonlyMap<string, EngineValue>; pivotStale: boolean; ruleValues?: WorkbookRuleValues } | null> {
  const ask = async (candidate: BackgroundCalculationSnapshot): Promise<WorkbookResultsAnswer | null> => {
    const answer = await request(candidate, signal, rules);
    return answer instanceof Map ? { results: answer } : answer as WorkbookResultsAnswer | null;
  };
  let sent: CalculationPivots | null = snapshot.pivots ?? null;
  let answer = await ask(snapshot);
  for (let round = 0; derive; round += 1) {
    if (!answer) return null;
    const pivots = derive(answer.results);
    if ((pivots?.id ?? null) === (sent?.id ?? null)) break;
    if (round === MAX_PIVOT_REPLACEMENTS) return { ...answer, pivotStale: true };
    sent = pivots;
    const { pivots: _previous, ...rest } = snapshot;
    void _previous;
    answer = await ask(pivots ? { ...rest, pivots } : rest);
  }
  return answer ? { ...answer, pivotStale: false } : null;
}

const sheetFingerprints = new WeakMap<object, string>();


/** Sheet parts no formula reads: a planned chart or object does not make the
 * workbook's results another workbook's (the apply path re-plans a chart
 * over the workbook that already holds it). */
const NOT_CALCULATED = new Set(["charts", "objects"]);

function sheetFingerprint(sheet: object): string {
  let fingerprint = sheetFingerprints.get(sheet);
  if (fingerprint === undefined) {
    const text = JSON.stringify(sheet, function (this: unknown, key: string, value: unknown) {
      return this === sheet && NOT_CALCULATED.has(key) ? undefined : value;
    });
    fingerprint = `${hash53(text)}.${text.length}`;
    sheetFingerprints.set(sheet, fingerprint);
  }
  return fingerprint;
}

/** A workbook's identity: its sheets' fingerprints (memoized per sheet object) and its date system. */
export function workbookValuesKey(source: WorkbookValuesSource): string {
  return `${source.dateSystem === "1904" ? "1904" : "1900"}|${source.sheets.map(sheetFingerprint).join("|")}`;
}

function approximateBytes(results: ReadonlyMap<string, EngineValue>): number {
  let bytes = 0;
  for (const [key, value] of results) {
    bytes += key.length * 2 + 48;
    if (isArrayValue(value)) bytes += value.grid.reduce((total, row) => total + row.length * 24, 0);
    else if (typeof value === "string") bytes += value.length * 2;
  }
  return bytes;
}

/** Times a failed workbook is asked again by itself before it stays local. */
const RETRIES = 3;

export class WorkbookValuesCache {
  readonly #fetch: (source: WorkbookValuesSource) => Promise<ReadonlyMap<string, EngineValue> | null>;
  readonly #onSettled: () => void;
  readonly #maxEntries: number;
  readonly #maxBytes: number;
  readonly #retryAfterMs: number;
  readonly #viewKey: (() => string) | undefined;
  /** Insertion order is recency (least recent first). */
  readonly #entries = new Map<string, { results: ReadonlyMap<string, EngineValue>; bytes: number }>();
  /** Workbooks the server could not answer, with when it said so. */
  readonly #local = new Map<string, number>();
  readonly #inFlight = new Set<string>();
  /** Workbooks whose failed answer is asked again by itself (a server or
   * engine that was not ready yet), with how often each failed. */
  readonly #retrying = new Set<string>();
  readonly #failures = new Map<string, number>();
  #bytes = 0;
  #generation = 0;

  constructor(options: {
    fetch(source: WorkbookValuesSource): Promise<ReadonlyMap<string, EngineValue> | null>;
    /** A requested workbook's answer landed (or failed): plan again. */
    onSettled(): void;
    maxEntries?: number;
    maxBytes?: number;
    /** After this long a workbook the server could not answer is asked again
     * (default never). */
    retryAfterMs?: number;
    /** View state the answer depends on beyond the workbook (a PivotTable
     * report's labels and collapsed items): part of every entry's key. */
    viewKey?: () => string;
  }) {
    this.#fetch = options.fetch;
    this.#onSettled = options.onSettled;
    this.#maxEntries = options.maxEntries ?? 16;
    this.#maxBytes = options.maxBytes ?? 32 * 1024 * 1024;
    this.#retryAfterMs = options.retryAfterMs ?? Infinity;
    this.#viewKey = options.viewKey;
  }

  



  get pending(): boolean { return this.#inFlight.size > 0 || this.#retrying.size > 0; }

  lookup(source: WorkbookValuesSource): WorkbookValuesLookup {
    const key = this.#viewKey ? `${workbookValuesKey(source)}#${this.#viewKey()}` : workbookValuesKey(source);
    const entry = this.#entries.get(key);
    if (entry) {
      this.#entries.delete(key); this.#entries.set(key, entry);
      return { kind: "values", results: entry.results };
    }
    const refused = this.#local.get(key);
    if (refused !== undefined) {
      if (Date.now() - refused < this.#retryAfterMs) return { kind: "local" };
      this.#local.delete(key);
    }
    // One request per workbook, however often the chain is planned again.
    if (this.#inFlight.has(key)) return { kind: "pending" };
    this.#inFlight.add(key);
    const generation = this.#generation;
    void this.#fetch(source).catch(() => null).then((results) => {
      if (generation !== this.#generation) return; // Cleared meanwhile (the page closed).
      this.#inFlight.delete(key);
      if (results) { this.#failures.delete(key); this.#store(key, results); }
      else this.#failed(key, generation);
      this.#onSettled();
    });
    return { kind: "pending" };
  }

  /** Drops every answer and every request in flight (their answers are discarded). */
  clear(): void {
    this.#generation += 1;
    this.#entries.clear(); this.#local.clear(); this.#inFlight.clear(); this.#bytes = 0;
    this.#retrying.clear(); this.#failures.clear();
  }

  /** A failed answer is remembered for retryAfterMs; with a finite delay it
   * is asked again by itself (the owner re-plans) up to RETRIES times, then
   * stays local: a workbook the server cannot take is not "calculating". */
  #failed(key: string, generation: number): void {
    this.#remember(key);
    const failures = (this.#failures.get(key) ?? 0) + 1;
    this.#failures.set(key, failures);
    if (!Number.isFinite(this.#retryAfterMs) || failures >= RETRIES) { this.#retrying.delete(key); return; }
    this.#retrying.add(key);
    // A little past the window: the lookup it prompts reads the window off
    // Date.now(), which a timer can be a millisecond ahead of, and found
    // still local, it would not ask again while pending had turned false.
    setTimeout(() => {
      if (generation !== this.#generation || !this.#retrying.delete(key)) return;
      this.#onSettled();
    }, this.#retryAfterMs + 50);
  }

  #store(key: string, results: ReadonlyMap<string, EngineValue>): void {
    const bytes = approximateBytes(results);
    if (bytes > this.#maxBytes) { this.#remember(key); return; }
    this.#entries.set(key, { results, bytes });
    this.#bytes += bytes;
    for (const [oldest, entry] of this.#entries) {
      if (this.#entries.size <= this.#maxEntries && this.#bytes <= this.#maxBytes) break;
      this.#entries.delete(oldest); this.#bytes -= entry.bytes;
    }
  }

  #remember(key: string): void {
    this.#local.set(key, Date.now());
    if (this.#local.size > 64) this.#local.delete(this.#local.keys().next().value!);
  }
}
