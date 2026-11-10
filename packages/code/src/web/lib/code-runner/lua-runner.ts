/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-local Lua 5.5 execution.
 *
 * The checked-in worker/WASM pair owns the interpreter. This host layer owns
 * request generations, cancellation and the second trust boundary around
 * streamed/final output. A namespace keeps one interpreter for Run + Console;
 * Stop, timeout and Restart physically retire that exact worker generation.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import type { OnOutput } from "@mrnote/core/web/lib/code-runner/output-coalescer.ts";
import {
  MAX_EXECUTION_LOG_CHARS,
  MAX_EXECUTION_LOG_LINE_CHARS,
  MAX_EXECUTION_LOG_LINES,
  MAX_EXECUTION_TEXT_CHARS,
  boundExecutionResult,
  createExecutionStreamLimiter,
  type ExecutionStreamLimiter,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";

export const LUA_EXECUTION_PROTOCOL = "mrnote.lua-execution.v1";
export const LUA_EXECUTION_RUNTIME = "lua-5.5.0";

const WORKER_URL = "/lua-execution-worker.js";
const READY_TIMEOUT_MS = 20_000;
const DEFAULT_TIMEOUT_MS = 5_000;
const MAX_SOURCE_BYTES = 2 * 1024 * 1024;
const MAX_STDIN_BYTES = 100 * 1024;
const MAX_OUTPUT_BYTES = 300_000;
const DEFAULT_INSTRUCTION_BUDGET = 20_000_000;
const encoder = new TextEncoder();

type LuaWorkerStatus =
  | "succeeded"
  | "compile-error"
  | "runtime-error"
  | "quota-exceeded";

export interface LuaWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

interface LuaFinalReceipt {
  status: LuaWorkerStatus;
  logs: string[];
  logStreams: ("stdout" | "stderr")[];
  error?: string;
}

interface PendingLuaRun {
  id: string;
  resolve: (result: RunResult) => void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
  onOutput?: OnOutput;
  streamLimiter: ExecutionStreamLimiter;
}

interface LuaWorkerHandle {
  worker: LuaWorkerLike;
  namespace: string;
  generation: number;
  sequence: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingLuaRun>;
}

export interface RunLuaOptions {
  namespace: string;
  stdin?: string;
  signal?: AbortSignal;
  onOutput?: OnOutput;
  timeoutMs?: number;
  instructionBudget?: number;
  outputBytes?: number;
}

let generation = 0;
const workers = new Map<string, LuaWorkerHandle>();

const defaultWorkerFactory = (): LuaWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Lua execution requires a browser Worker.");
  }
  return new Worker(WORKER_URL, { name: "mrnote-lua-execution" });
};

let workerFactory: () => LuaWorkerLike = defaultWorkerFactory;

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
  handle: LuaWorkerHandle,
  pending: PendingLuaRun,
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

function settleAll(handle: LuaWorkerHandle, result: RunResult): void {
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, result);
  }
}

function retireHandle(handle: LuaWorkerHandle, result: RunResult): void {
  if (handle.retired) return;
  handle.retired = true;
  clearTimeout(handle.readyTimer);
  if (!handle.readySettled) {
    handle.readySettled = true;
    handle.settleReady(false);
  }
  settleAll(handle, result);
  try {
    handle.worker.terminate();
  } catch {
    // A failed generation is already unusable; termination is best effort.
  }
  if (workers.get(handle.namespace) === handle) {
    workers.delete(handle.namespace);
  }
}

function parseReady(value: unknown): boolean {
  if (!isRecord(value)) return false;
  try {
    return (
      value.protocol === LUA_EXECUTION_PROTOCOL &&
      value.runtime === LUA_EXECUTION_RUNTIME &&
      value.type === "ready"
    );
  } catch {
    return false;
  }
}

function parseFinalReceipt(
  value: unknown,
  pending: PendingLuaRun,
): LuaFinalReceipt | null {
  if (!isRecord(value)) return null;
  try {
    if (
      value.protocol !== LUA_EXECUTION_PROTOCOL ||
      value.runtime !== LUA_EXECUTION_RUNTIME ||
      value.type !== "result" ||
      value.id !== pending.id ||
      (value.status !== "succeeded" &&
        value.status !== "compile-error" &&
        value.status !== "runtime-error" &&
        value.status !== "quota-exceeded") ||
      !Array.isArray(value.logs) ||
      !Array.isArray(value.logStreams) ||
      value.logs.length !== value.logStreams.length ||
      value.logs.length > MAX_EXECUTION_LOG_LINES
    ) {
      return null;
    }
    let totalChars = 0;
    const logs: string[] = [];
    const logStreams: ("stdout" | "stderr")[] = [];
    for (let index = 0; index < value.logs.length; index += 1) {
      const log = value.logs[index];
      const stream = value.logStreams[index];
      if (
        typeof log !== "string" ||
        log.length > MAX_EXECUTION_LOG_LINE_CHARS ||
        (stream !== "stdout" && stream !== "stderr")
      ) {
        return null;
      }
      totalChars += log.length;
      if (totalChars > MAX_EXECUTION_LOG_CHARS) return null;
      logs.push(log);
      logStreams.push(stream);
    }
    const succeeded = value.status === "succeeded";
    if (
      (succeeded && value.error !== undefined) ||
      (!succeeded &&
        (typeof value.error !== "string" ||
          value.error.length === 0 ||
          value.error.length > MAX_EXECUTION_TEXT_CHARS))
    ) {
      return null;
    }
    return {
      status: value.status,
      logs,
      logStreams,
      ...(succeeded ? {} : { error: value.error as string }),
    };
  } catch {
    return null;
  }
}

function resultForReceipt(receipt: LuaFinalReceipt): RunResult {
  const base = { logs: receipt.logs, logStreams: receipt.logStreams };
  if (receipt.status === "succeeded") return boundExecutionResult(base);
  const detail = receipt.error ?? "Lua execution failed.";
  if (receipt.status === "compile-error") {
    return boundExecutionResult({ ...base, error: `Lua compile error: ${detail}` });
  }
  if (receipt.status === "quota-exceeded") {
    return boundExecutionResult({
      ...base,
      error: `Lua execution quota exceeded: ${detail}`,
    });
  }
  return boundExecutionResult({ ...base, error: `Lua runtime error: ${detail}` });
}

function receiveWorkerMessage(handle: LuaWorkerHandle, value: unknown): void {
  if (handle.retired || workers.get(handle.namespace) !== handle) return;
  if (isRecord(value)) {
    try {
      if (value.type === "ready") {
        if (!parseReady(value) || handle.ready) {
          retireHandle(handle, {
            logs: [],
            error: "Lua runtime returned a malformed readiness receipt.",
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
        value.protocol === LUA_EXECUTION_PROTOCOL &&
        value.runtime === LUA_EXECUTION_RUNTIME
      ) {
        retireHandle(handle, {
          logs: [],
          error: "Lua runtime generation failed.",
        });
        return;
      }
      if (typeof value.id !== "string") return;
      const pending = handle.pending.get(value.id);
      if (!pending) return;
      if (value.type === "output") {
        if (
          value.protocol !== LUA_EXECUTION_PROTOCOL ||
          value.runtime !== LUA_EXECUTION_RUNTIME ||
          typeof value.chunk !== "string" ||
          (value.stream !== "stdout" && value.stream !== "stderr")
        ) {
          retireHandle(handle, {
            logs: [],
            error: "Lua runtime returned malformed streamed output.",
          });
          return;
        }
        const retained = pending.streamLimiter.push(value.chunk);
        if (retained !== null) pending.onOutput?.(retained, value.stream);
        return;
      }
      const receipt = parseFinalReceipt(value, pending);
      if (!receipt) {
        retireHandle(handle, {
          logs: [],
          error: "Lua runtime returned a malformed execution receipt.",
        });
        return;
      }
      settlePending(handle, pending, resultForReceipt(receipt));
    } catch {
      retireHandle(handle, {
        logs: [],
        error: "Lua runtime receipt validation failed.",
      });
    }
  }
}

function createHandle(namespace: string): LuaWorkerHandle | null {
  let worker: LuaWorkerLike;
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
    pending: new Map<string, PendingLuaRun>(),
  } satisfies LuaWorkerHandle;
  handle.readyTimer = setTimeout(
    () =>
      retireHandle(handle, {
        logs: [],
        error: "Lua runtime did not become ready.",
      }),
    READY_TIMEOUT_MS,
  );
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () =>
    retireHandle(handle, {
      logs: [],
      error: "Lua runtime worker failed.",
    });
  workers.set(namespace, handle);
  return handle;
}

function getHandle(namespace: string): LuaWorkerHandle | null {
  const current = workers.get(namespace);
  if (current && !current.retired) return current;
  return createHandle(namespace);
}

function waitForReady(
  handle: LuaWorkerHandle,
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

export async function runLua(
  source: string,
  options: RunLuaOptions,
): Promise<RunResult> {
  const stdin = options.stdin ?? "";
  if (options.signal?.aborted) return interruptedResult();
  if (
    !options.namespace ||
    hasUnpairedSurrogate(source) ||
    hasUnpairedSurrogate(stdin) ||
    byteLength(source) > MAX_SOURCE_BYTES ||
    byteLength(stdin) > MAX_STDIN_BYTES
  ) {
    return { logs: [], error: "Lua execution input exceeds its trusted budget." };
  }
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const instructionBudget =
    options.instructionBudget ?? DEFAULT_INSTRUCTION_BUDGET;
  const outputBytes = options.outputBytes ?? MAX_OUTPUT_BYTES;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000 ||
    !Number.isSafeInteger(instructionBudget) ||
    instructionBudget < 1_000 ||
    instructionBudget > 50_000_000 ||
    !Number.isSafeInteger(outputBytes) ||
    outputBytes <= 0 ||
    outputBytes > MAX_OUTPUT_BYTES
  ) {
    return { logs: [], error: "Lua execution limits are invalid." };
  }

  const handle = getHandle(options.namespace);
  if (!handle || !(await waitForReady(handle, options.signal))) {
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: "Lua runtime is unavailable." };
  }
  if (
    options.signal?.aborted ||
    handle.retired ||
    workers.get(options.namespace) !== handle ||
    !handle.ready
  ) {
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: "Lua runtime generation is unavailable." };
  }

  const id = `lua:${handle.generation}:${++handle.sequence}`;
  return new Promise<RunResult>((resolve) => {
    let dispatched = false;
    const pending: PendingLuaRun = {
      id,
      resolve,
      timer: 0 as unknown as ReturnType<typeof setTimeout>,
      signal: options.signal,
      onOutput: options.onOutput,
      streamLimiter: createExecutionStreamLimiter(),
    };
    pending.timer = setTimeout(() => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(handle, {
        logs: [],
        error: `Lua execution timed out after ${Math.round(timeoutMs / 1000)}s and was stopped.`,
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
        protocol: LUA_EXECUTION_PROTOCOL,
        type: "run",
        id,
        source,
        stdin,
        outputBytes,
        instructionBudget,
      });
    } catch {
      retireHandle(handle, {
        logs: [],
        error: "Lua runtime transport failed.",
      });
    }
  });
}

/** Physically stop the exact namespace generation. */
export function interruptLua(namespace: string): boolean {
  const handle = workers.get(namespace);
  if (!handle || handle.retired || handle.pending.size === 0) return false;
  retireHandle(handle, interruptedResult());
  return true;
}

/** Restart/dispose the page interpreter. The next request creates a clean one. */
export function resetLuaRuntime(namespace: string): void {
  const handle = workers.get(namespace);
  if (!handle) return;
  retireHandle(handle, {
    logs: [],
    error: "Lua runtime was restarted.",
    interrupted: true,
  });
}

export function __setLuaWorkerFactoryForTests(
  factory: (() => LuaWorkerLike) | null,
): void {
  for (const handle of [...workers.values()]) {
    retireHandle(handle, interruptedResult());
  }
  workerFactory = factory ?? defaultWorkerFactory;
}

export function __resetLuaRunnerForTests(): void {
  __setLuaWorkerFactoryForTests(null);
}
