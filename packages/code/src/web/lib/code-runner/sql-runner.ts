/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-local SQL execution in a dedicated sql.js Worker.
 *
 * One worker generation owns one page namespace and its in-memory database.
 * Run + Console therefore keep notebook-style state, while Stop, timeout and
 * Restart physically terminate the exact generation. The host validates and
 * bounds every final receipt before it can reach the renderer or persistence.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import type { VariableInfo } from "@/lib/code-runner/inspector";
import {
  MAX_EXECUTION_TEXT_CHARS,
  boundExecutionResult,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";

export const SQL_EXECUTION_PROTOCOL = "mrnote.sql-execution.v1";
export const SQL_EXECUTION_RUNTIME = "sqljs-1.14.1";

const WORKER_URL = "/sql-execution-worker.js";
const READY_TIMEOUT_MS = 20_000;
const DEFAULT_TIMEOUT_MS = 5_000;
const MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const MAX_OUTPUT_BYTES = 300_000;
const MAX_COLUMNS = 100;
const MAX_ROWS = 1_000;
const MAX_CELL_CHARS = 2_000;
const MAX_COLUMN_CHARS = 1_000;
const MAX_TABLES = 100;
const MAX_TABLE_METADATA_BYTES = 32 * 1024;
const MAX_EARLIER_RESULTS = 20;
const MAX_EARLIER_ROWS = 50;
const encoder = new TextEncoder();

type SqlReceiptStatus = "succeeded" | "runtime-error";
type SqlReceiptKind = "query" | "statement";

export interface SqlWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

interface SqlTableSummary {
  name: string;
  rowCount: number;
  columns: string[];
}

/** A result set returned before the run's last one. */
interface SqlEarlierResult {
  columns: string[];
  rows: unknown[][];
  totalRows: number;
  truncated: boolean;
}

interface SqlFinalReceipt {
  status: SqlReceiptStatus;
  kind?: SqlReceiptKind;
  columns?: string[];
  rows?: unknown[][];
  totalRows?: number;
  earlier?: SqlEarlierResult[];
  rowsModified?: number;
  truncated?: boolean;
  error?: string;
  tables: SqlTableSummary[];
}

interface PendingSqlRun {
  id: string;
  resolve: (result: RunResult) => void;
  timer: ReturnType<typeof setTimeout>;
  outputBytes: number;
  signal?: AbortSignal;
  abortListener?: () => void;
}

interface SqlWorkerHandle {
  worker: SqlWorkerLike;
  namespace: string;
  generation: number;
  sequence: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingSqlRun>;
  inspector: VariableInfo[];
}

export interface SqlRunOptions {
  namespace?: string;
  signal?: AbortSignal;
  timeoutMs?: number;
  outputBytes?: number;
}

let generation = 0;
const workers = new Map<string, SqlWorkerHandle>();

const defaultWorkerFactory = (): SqlWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("SQL execution requires a browser Worker.");
  }
  return new Worker(WORKER_URL, { name: "mrnote-sql-execution" });
};

let workerFactory: () => SqlWorkerLike = defaultWorkerFactory;

function byteLength(value: string): number {
  return encoder.encode(value).byteLength;
}

function hasUnpairedSurrogate(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return true;
      index += 1;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      return true;
    }
  }
  return false;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function interruptedResult(): RunResult {
  return { logs: [], interrupted: true };
}

function settlePending(
  handle: SqlWorkerHandle,
  pending: PendingSqlRun,
  result: RunResult,
): void {
  if (handle.pending.get(pending.id) !== pending) return;
  handle.pending.delete(pending.id);
  clearTimeout(pending.timer);
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  pending.resolve(boundExecutionResult(result));
}

function retireHandle(handle: SqlWorkerHandle, result: RunResult): void {
  if (handle.retired) return;
  handle.retired = true;
  clearTimeout(handle.readyTimer);
  if (!handle.readySettled) {
    handle.readySettled = true;
    handle.settleReady(false);
  }
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, result);
  }
  handle.inspector = [];
  try {
    handle.worker.terminate();
  } catch {
    // The generation is already unusable; termination is best effort.
  }
  if (workers.get(handle.namespace) === handle) workers.delete(handle.namespace);
}

function parseReady(value: unknown): boolean {
  if (!isRecord(value)) return false;
  try {
    return (
      value.protocol === SQL_EXECUTION_PROTOCOL &&
      value.runtime === SQL_EXECUTION_RUNTIME &&
      value.type === "ready"
    );
  } catch {
    return false;
  }
}

function parseBoundedString(value: unknown, limit: number): string | null {
  return typeof value === "string" && value.length <= limit ? value : null;
}

function parseCell(value: unknown): string | number | boolean | null | undefined {
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "number") return Number.isFinite(value) ? value : undefined;
  if (typeof value === "string" && value.length <= MAX_CELL_CHARS) return value;
  return undefined;
}

function parseTables(value: unknown): SqlTableSummary[] | null {
  if (!Array.isArray(value) || value.length > MAX_TABLES) return null;
  const tables: SqlTableSummary[] = [];
  let metadataBytes = 0;
  for (const candidate of value) {
    if (!isRecord(candidate)) return null;
    const name = parseBoundedString(candidate.name, MAX_COLUMN_CHARS);
    if (
      name === null ||
      name.length === 0 ||
      !Number.isSafeInteger(candidate.rowCount) ||
      (candidate.rowCount as number) < 0 ||
      !Array.isArray(candidate.columns) ||
      candidate.columns.length > MAX_COLUMNS
    ) {
      return null;
    }
    const columns = candidate.columns.map((column) =>
      parseBoundedString(column, MAX_COLUMN_CHARS),
    );
    if (columns.some((column) => column === null)) return null;
    metadataBytes +=
      byteLength(name) +
      (columns as string[]).reduce(
        (total, column) => total + byteLength(column),
        0,
      );
    if (metadataBytes > MAX_TABLE_METADATA_BYTES) return null;
    tables.push({
      name,
      rowCount: candidate.rowCount as number,
      columns: columns as string[],
    });
  }
  return tables;
}

function parseFinalReceipt(
  value: unknown,
  pending: PendingSqlRun,
): SqlFinalReceipt | null {
  if (!isRecord(value)) return null;
  try {
    if (
      value.protocol !== SQL_EXECUTION_PROTOCOL ||
      value.runtime !== SQL_EXECUTION_RUNTIME ||
      value.type !== "result" ||
      value.id !== pending.id ||
      (value.status !== "succeeded" && value.status !== "runtime-error")
    ) {
      return null;
    }
    const tables = parseTables(value.tables);
    if (!tables) return null;
    if (value.status === "runtime-error") {
      const error = parseBoundedString(value.error, MAX_EXECUTION_TEXT_CHARS);
      if (
        error === null ||
        error.length === 0 ||
        value.kind !== undefined ||
        value.earlier !== undefined ||
        value.columns !== undefined ||
        value.rows !== undefined ||
        value.totalRows !== undefined ||
        value.rowsModified !== undefined
      ) {
        return null;
      }
      return { status: "runtime-error", error, tables };
    }
    if (
      value.error !== undefined ||
      (value.kind !== "query" && value.kind !== "statement")
    ) {
      return null;
    }
    if (value.kind === "statement") {
      if (
        !Number.isSafeInteger(value.rowsModified) ||
        (value.rowsModified as number) < 0 ||
        value.columns !== undefined ||
        value.rows !== undefined ||
        value.totalRows !== undefined ||
        value.truncated !== undefined ||
        value.earlier !== undefined
      ) {
        return null;
      }
      return {
        status: "succeeded",
        kind: "statement",
        rowsModified: value.rowsModified as number,
        tables,
      };
    }
    if (
      !Array.isArray(value.columns) ||
      value.columns.length === 0 ||
      value.columns.length > MAX_COLUMNS ||
      !Array.isArray(value.rows) ||
      value.rows.length > MAX_ROWS ||
      !Number.isSafeInteger(value.totalRows) ||
      (value.totalRows as number) < value.rows.length ||
      typeof value.truncated !== "boolean"
    ) {
      return null;
    }
    let retainedBytes = 0;
    const parseResultSet = (
      rawColumns: unknown[],
      rawRows: unknown[],
    ): { columns: string[]; rows: unknown[][] } | null => {
      const columns = rawColumns.map((column) =>
        parseBoundedString(column, MAX_COLUMN_CHARS),
      );
      if (columns.some((column) => column === null)) return null;
      retainedBytes += (columns as string[]).reduce(
        (total, column) => total + byteLength(column),
        0,
      );
      const rows: unknown[][] = [];
      for (const candidate of rawRows) {
        if (!Array.isArray(candidate) || candidate.length !== columns.length) {
          return null;
        }
        const row: unknown[] = [];
        for (const cell of candidate) {
          const parsed = parseCell(cell);
          if (parsed === undefined) return null;
          retainedBytes += byteLength(String(parsed ?? ""));
          row.push(parsed);
        }
        if (retainedBytes > pending.outputBytes) return null;
        rows.push(row);
      }
      return { columns: columns as string[], rows };
    };
    let earlier: SqlEarlierResult[] | undefined;
    if (value.earlier !== undefined) {
      if (
        !Array.isArray(value.earlier) ||
        value.earlier.length === 0 ||
        value.earlier.length > MAX_EARLIER_RESULTS
      ) {
        return null;
      }
      earlier = [];
      for (const candidate of value.earlier) {
        if (
          !isRecord(candidate) ||
          !Array.isArray(candidate.columns) ||
          candidate.columns.length === 0 ||
          candidate.columns.length > MAX_COLUMNS ||
          !Array.isArray(candidate.rows) ||
          candidate.rows.length > MAX_EARLIER_ROWS ||
          !Number.isSafeInteger(candidate.totalRows) ||
          (candidate.totalRows as number) < candidate.rows.length ||
          typeof candidate.truncated !== "boolean" ||
          Object.keys(candidate).some(
            (key) => !["columns", "rows", "totalRows", "truncated"].includes(key),
          )
        ) {
          return null;
        }
        const set = parseResultSet(candidate.columns, candidate.rows);
        if (!set) return null;
        earlier.push({
          ...set,
          totalRows: candidate.totalRows as number,
          truncated: candidate.truncated,
        });
      }
    }
    const last = parseResultSet(value.columns, value.rows);
    if (!last) return null;
    return {
      status: "succeeded",
      kind: "query",
      columns: last.columns,
      rows: last.rows,
      totalRows: value.totalRows as number,
      truncated: value.truncated,
      ...(earlier ? { earlier } : {}),
      tables,
    };
  } catch {
    return null;
  }
}

function sqlTextCell(value: unknown): string {
  if (value === null || value === undefined) return "NULL";
  return String(value);
}

function resultForReceipt(receipt: SqlFinalReceipt): RunResult {
  if (receipt.status === "runtime-error") {
    return boundExecutionResult({
      logs: [],
      error: receipt.error ?? "SQL execution failed.",
    });
  }
  if (receipt.kind === "statement") {
    const modified = receipt.rowsModified ?? 0;
    return boundExecutionResult({
      logs: [modified > 0 ? `${modified} row(s) affected` : "OK"],
      logStreams: ["stdout"],
    });
  }
  const columns = receipt.columns ?? [];
  const rows = receipt.rows ?? [];
  const totalRows = receipt.totalRows ?? rows.length;
  const hiddenRows = Math.max(0, totalRows - rows.length);
  // Every SELECT before the last one is printed as text, as the sqlite3
  // shell prints it; the last one is the table.
  const earlierLogs = (receipt.earlier ?? []).flatMap((set) => {
    const hidden = Math.max(0, set.totalRows - set.rows.length);
    return [
      set.columns.join("\t"),
      ...set.rows.map((row) => row.map(sqlTextCell).join("\t")),
      hidden > 0
        ? `(${set.totalRows} row(s), ${hidden} not shown)`
        : `(${set.totalRows} row(s))`,
      "",
    ];
  });
  const logs = [
    ...earlierLogs,
    ...(hiddenRows > 0 ? [`… ${hiddenRows} more row(s)`] : []),
  ];
  return boundExecutionResult({
    logs,
    ...(logs.length > 0
      ? { logStreams: logs.map(() => "stdout" as const) }
      : {}),
    rich: { kind: "table", columns, rows },
    result: `${totalRows} row(s)`,
  });
}

function updateInspector(handle: SqlWorkerHandle, tables: SqlTableSummary[]): void {
  handle.inspector = tables.map((table) => ({
    name: table.name,
    type: "table",
    preview: `${table.rowCount} row(s) · (${table.columns.join(", ")})`,
  }));
}

function receiveWorkerMessage(handle: SqlWorkerHandle, value: unknown): void {
  if (handle.retired || workers.get(handle.namespace) !== handle) return;
  if (!isRecord(value)) return;
  try {
    if (value.type === "ready") {
      if (!parseReady(value) || handle.ready) {
        retireHandle(handle, {
          logs: [],
          error: "SQL runtime returned a malformed readiness receipt.",
        });
        return;
      }
      handle.ready = true;
      clearTimeout(handle.readyTimer);
      if (!handle.readySettled) {
        handle.readySettled = true;
        handle.settleReady(true);
      }
      return;
    }
    if (
      value.type === "fatal" &&
      value.protocol === SQL_EXECUTION_PROTOCOL &&
      value.runtime === SQL_EXECUTION_RUNTIME
    ) {
      retireHandle(handle, {
        logs: [],
        error: "SQL runtime generation failed.",
      });
      return;
    }
    if (typeof value.id !== "string") return;
    const pending = handle.pending.get(value.id);
    if (!pending) return;
    const receipt = parseFinalReceipt(value, pending);
    if (!receipt) {
      retireHandle(handle, {
        logs: [],
        error: "SQL runtime returned a malformed execution receipt.",
      });
      return;
    }
    updateInspector(handle, receipt.tables);
    settlePending(handle, pending, resultForReceipt(receipt));
  } catch {
    retireHandle(handle, {
      logs: [],
      error: "SQL runtime receipt validation failed.",
    });
  }
}

function createHandle(namespace: string): SqlWorkerHandle | null {
  let worker: SqlWorkerLike;
  try {
    worker = workerFactory();
  } catch {
    return null;
  }
  let settleReady!: (ready: boolean) => void;
  const readyPromise = new Promise<boolean>((resolve) => {
    settleReady = resolve;
  });
  const handle = {
    worker,
    namespace,
    generation: ++generation,
    sequence: 0,
    ready: false,
    retired: false,
    readySettled: false,
    readyPromise,
    settleReady,
    readyTimer: 0 as unknown as ReturnType<typeof setTimeout>,
    pending: new Map<string, PendingSqlRun>(),
    inspector: [],
  } satisfies SqlWorkerHandle;
  handle.readyTimer = setTimeout(
    () =>
      retireHandle(handle, {
        logs: [],
        error: "SQL runtime did not become ready.",
      }),
    READY_TIMEOUT_MS,
  );
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () =>
    retireHandle(handle, { logs: [], error: "SQL runtime worker failed." });
  workers.set(namespace, handle);
  return handle;
}

function getHandle(namespace: string): SqlWorkerHandle | null {
  const current = workers.get(namespace);
  if (current && !current.retired) return current;
  return createHandle(namespace);
}

function waitForReady(
  handle: SqlWorkerHandle,
  signal?: AbortSignal,
): Promise<boolean> {
  if (!signal) return handle.readyPromise;
  if (signal.aborted) return Promise.resolve(false);
  return new Promise((resolve) => {
    let settled = false;
    const finish = (ready: boolean) => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      resolve(ready);
    };
    const onAbort = () => finish(false);
    signal.addEventListener("abort", onAbort, { once: true });
    void handle.readyPromise.then(finish);
    if (signal.aborted) onAbort();
  });
}

/** Snapshot the last trusted worker table inventory for the namespace. */
export function inspectSql(namespace = "__default__"): VariableInfo[] {
  const handle = workers.get(namespace);
  return handle && !handle.retired
    ? handle.inspector.map((item) => ({ ...item }))
    : [];
}

export async function runSql(
  source: string,
  options: SqlRunOptions = {},
): Promise<RunResult> {
  const namespace = options.namespace ?? "__default__";
  if (options.signal?.aborted) return interruptedResult();
  if (
    !namespace ||
    hasUnpairedSurrogate(source) ||
    byteLength(source) > MAX_SOURCE_BYTES
  ) {
    return {
      logs: [],
      error: "SQL execution input exceeds its trusted budget.",
    };
  }
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const outputBytes = options.outputBytes ?? MAX_OUTPUT_BYTES;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000 ||
    !Number.isSafeInteger(outputBytes) ||
    outputBytes <= 0 ||
    outputBytes > MAX_OUTPUT_BYTES
  ) {
    return { logs: [], error: "SQL execution limits are invalid." };
  }

  const handle = getHandle(namespace);
  if (!handle) {
    return { logs: [], error: "SQL runtime is unavailable." };
  }
  const ready = await waitForReady(handle, options.signal);
  if (!ready) {
    if (
      options.signal?.aborted &&
      !handle.retired &&
      workers.get(namespace) === handle
    ) {
      retireHandle(handle, interruptedResult());
    }
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: "SQL runtime is unavailable." };
  }
  if (
    options.signal?.aborted ||
    handle.retired ||
    workers.get(namespace) !== handle ||
    !handle.ready
  ) {
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: "SQL runtime generation is unavailable." };
  }

  const id = `sql:${handle.generation}:${++handle.sequence}`;
  return new Promise<RunResult>((resolve) => {
    let dispatched = false;
    const pending: PendingSqlRun = {
      id,
      resolve,
      outputBytes,
      timer: 0 as unknown as ReturnType<typeof setTimeout>,
      signal: options.signal,
    };
    pending.timer = setTimeout(() => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(handle, {
        logs: [],
        error: `SQL execution timed out after ${Math.round(
          timeoutMs / 1_000,
        )}s and was stopped.`,
        timedOut: true,
      });
    }, timeoutMs);
    pending.abortListener = () => {
      if (handle.pending.get(id) !== pending) return;
      if (!dispatched) {
        settlePending(handle, pending, interruptedResult());
        return;
      }
      retireHandle(handle, interruptedResult());
    };
    handle.pending.set(id, pending);
    options.signal?.addEventListener("abort", pending.abortListener, {
      once: true,
    });
    if (options.signal?.aborted) {
      pending.abortListener();
      return;
    }
    try {
      dispatched = true;
      handle.worker.postMessage({
        protocol: SQL_EXECUTION_PROTOCOL,
        type: "run",
        id,
        source,
        outputBytes,
      });
    } catch {
      retireHandle(handle, {
        logs: [],
        error: "SQL runtime transport failed.",
      });
    }
  });
}

/** Physically stop the exact page worker generation. */
export function interruptSql(namespace = "__default__"): boolean {
  const handle = workers.get(namespace);
  if (!handle || handle.retired) return false;
  retireHandle(handle, interruptedResult());
  return true;
}

/** Restart/dispose one page database (or every test/runtime namespace). */
export function resetSqlRuntime(namespace?: string): void {
  if (namespace !== undefined) {
    const handle = workers.get(namespace);
    if (handle) {
      retireHandle(handle, {
        logs: [],
        error: "SQL runtime was restarted.",
        interrupted: true,
      });
    }
    return;
  }
  for (const handle of [...workers.values()]) {
    retireHandle(handle, {
      logs: [],
      error: "SQL runtime was restarted.",
      interrupted: true,
    });
  }
}

export function __setSqlWorkerFactoryForTests(
  factory: (() => SqlWorkerLike) | null,
): void {
  resetSqlRuntime();
  workerFactory = factory ?? defaultWorkerFactory;
}

export function __resetSqlRunnerForTests(): void {
  __setSqlWorkerFactoryForTests(null);
}

/** Backward-compatible test cleanup name; production never owns a page module. */
export function __resetSqlModuleForTests(): void {
  __resetSqlRunnerForTests();
}
