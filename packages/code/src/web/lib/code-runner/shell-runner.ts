/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Host authority for the dedicated single-run Bashkit WASM Worker.
 *
 * A page namespace owns at most one Worker generation. A run owns that
 * generation even while the WASM asset is loading, so Stop can physically
 * terminate it before dispatch. Every final receipt retires the Worker and the
 * next run starts with a fresh virtual filesystem.
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
import {
  SHELL_EXECUTION_PROTOCOL,
  SHELL_EXECUTION_RUNTIME,
  SHELL_MAX_OUTPUT_BYTES,
  SHELL_MAX_SOURCE_BYTES,
  SHELL_MAX_STDIN_BYTES,
  type ShellExecutionStatus,
} from "./shell-execution-runtime";

export { SHELL_EXECUTION_PROTOCOL, SHELL_EXECUTION_RUNTIME };

export const SHELL_RUNTIME_UNAVAILABLE_MESSAGE =
  "The isolated Shell WebAssembly runtime is unavailable here.";

const READY_TIMEOUT_MS = 20_000;
const DEFAULT_TIMEOUT_MS = 5_000;
const encoder = new TextEncoder();

export type ShellRuntimeAvailability = "checking" | "available" | "unavailable";

export interface ShellWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

interface ShellFinalReceipt {
  status: ShellExecutionStatus;
  logs: string[];
  logStreams: Array<"stdout" | "stderr">;
  exitCode: number | null;
  error?: string;
  retireGeneration: true;
}

export interface ShellRunResult extends RunResult {
  /** Exact process-style status returned by the trusted Shell generation. */
  readonly exitCode?: number;
}

interface PendingShellRun {
  id: string;
  resolve: (result: ShellRunResult) => void;
  timer: ReturnType<typeof setTimeout> | null;
  signal?: AbortSignal;
  abortListener?: () => void;
  onOutput?: OnOutput;
  streamLimiter: ExecutionStreamLimiter;
  streamBytes: number;
  stdoutBytes: number;
  stderrBytes: number;
  timeoutMs: number;
  outputBytes: number;
  source: string;
  stdin: string;
  dispatched: boolean;
}

interface ShellWorkerHandle {
  worker: ShellWorkerLike;
  namespace: string;
  generation: number;
  sequence: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingShellRun>;
}

interface AvailabilityState {
  status: ShellRuntimeAvailability;
  listeners: Set<() => void>;
}

export interface RunShellOptions {
  namespace: string;
  stdin?: string;
  signal?: AbortSignal;
  onOutput?: OnOutput;
  timeoutMs?: number;
  outputBytes?: number;
}

let generation = 0;
const workers = new Map<string, ShellWorkerHandle>();
const availability = new Map<string, AvailabilityState>();

const defaultWorkerFactory = (): ShellWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Shell execution requires a browser Worker.");
  }
  return new Worker(new URL("./shell-execution-worker.ts", import.meta.url), {
    name: "mrnote-shell-execution",
    type: "module",
  });
};

let workerFactory: () => ShellWorkerLike = defaultWorkerFactory;

function browserSupportsShell(): boolean {
  return (
    (workerFactory !== defaultWorkerFactory || typeof Worker !== "undefined") &&
    typeof WebAssembly !== "undefined" &&
    typeof WebAssembly.compile === "function" &&
    typeof TextEncoder !== "undefined" &&
    typeof TextDecoder !== "undefined" &&
    typeof globalThis.crypto?.subtle !== "undefined"
  );
}

function availabilityState(namespace: string): AvailabilityState {
  let state = availability.get(namespace);
  if (!state) {
    state = { status: "checking", listeners: new Set() };
    availability.set(namespace, state);
  }
  return state;
}

function setAvailability(namespace: string, status: ShellRuntimeAvailability): void {
  const state = availabilityState(namespace);
  if (state.status === status) return;
  state.status = status;
  for (const listener of [...state.listeners]) listener();
}

export function getShellRuntimeAvailability(
  namespace: string,
): ShellRuntimeAvailability {
  if (!namespace || !browserSupportsShell()) return "unavailable";
  return availabilityState(namespace).status;
}

export function subscribeShellRuntimeAvailability(
  namespace: string,
  listener: () => void,
): () => void {
  const state = availabilityState(namespace);
  state.listeners.add(listener);
  return () => state.listeners.delete(listener);
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

function isOwnRecord(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  try {
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return false;
    return Object.values(Object.getOwnPropertyDescriptors(value)).every(
      (descriptor) => "value" in descriptor,
    );
  } catch {
    return false;
  }
}

function ownArray(value: unknown): unknown[] | null {
  if (!Array.isArray(value)) return null;
  try {
    const descriptors = Object.getOwnPropertyDescriptors(value);
    if (
      Reflect.ownKeys(descriptors).some(
        (key) =>
          typeof key !== "string" ||
          (key !== "length" && !/^(0|[1-9][0-9]*)$/.test(key)),
      )
    ) {
      return null;
    }
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)];
      if (!descriptor || !("value" in descriptor)) return null;
    }
    return Array.from(value);
  } catch {
    return null;
  }
}

function interruptedResult(): RunResult {
  return { logs: [], interrupted: true };
}

function settlePending(
  handle: ShellWorkerHandle,
  pending: PendingShellRun,
  result: ShellRunResult,
): void {
  if (handle.pending.get(pending.id) !== pending) return;
  handle.pending.delete(pending.id);
  if (pending.timer) clearTimeout(pending.timer);
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  const bounded = boundExecutionResult(result);
  pending.resolve(
    Number.isSafeInteger(result.exitCode) && Number(result.exitCode) >= 0
      ? { ...bounded, exitCode: Number(result.exitCode) }
      : bounded,
  );
}

function settleAll(handle: ShellWorkerHandle, result: ShellRunResult): void {
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, result);
  }
}

function retireHandle(
  handle: ShellWorkerHandle,
  result: ShellRunResult,
  runtimeUnavailable = false,
): void {
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
    // The generation is detached before best-effort physical termination.
  }
  if (workers.get(handle.namespace) === handle) workers.delete(handle.namespace);
  if (runtimeUnavailable) setAvailability(handle.namespace, "unavailable");
}

function parseReady(value: unknown): boolean {
  return (
    isOwnRecord(value) &&
    value.protocol === SHELL_EXECUTION_PROTOCOL &&
    value.runtime === SHELL_EXECUTION_RUNTIME &&
    value.type === "ready"
  );
}

function parseFinalReceipt(
  value: unknown,
  pending: PendingShellRun,
): ShellFinalReceipt | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      value.protocol !== SHELL_EXECUTION_PROTOCOL ||
      value.runtime !== SHELL_EXECUTION_RUNTIME ||
      value.type !== "result" ||
      value.id !== pending.id ||
      (value.status !== "succeeded" &&
        value.status !== "compile-error" &&
        value.status !== "runtime-error" &&
        value.status !== "quota-exceeded") ||
      value.retireGeneration !== true
    ) {
      return null;
    }
    const rawLogs = ownArray(value.logs);
    const rawStreams = ownArray(value.logStreams);
    if (
      !rawLogs ||
      !rawStreams ||
      rawLogs.length !== rawStreams.length ||
      rawLogs.length > MAX_EXECUTION_LOG_LINES
    ) {
      return null;
    }
    let aggregateBytes = 0;
    let stdoutBytes = 0;
    let stderrBytes = 0;
    const logs: string[] = [];
    const logStreams: Array<"stdout" | "stderr"> = [];
    for (let index = 0; index < rawLogs.length; index += 1) {
      const log = rawLogs[index];
      const stream = rawStreams[index];
      if (
        typeof log !== "string" ||
        log.length > MAX_EXECUTION_LOG_LINE_CHARS ||
        (stream !== "stdout" && stream !== "stderr")
      ) {
        return null;
      }
      const bytes = encoder.encode(log).byteLength;
      aggregateBytes += bytes;
      if (stream === "stderr") stderrBytes += bytes;
      else stdoutBytes += bytes;
      if (
        aggregateBytes > pending.outputBytes ||
        aggregateBytes > MAX_EXECUTION_LOG_CHARS ||
        stdoutBytes > pending.outputBytes ||
        stderrBytes > pending.outputBytes
      ) {
        return null;
      }
      logs.push(log);
      logStreams.push(stream);
    }
    const succeeded = value.status === "succeeded";
    const exitCode = value.exitCode;
    if (
      (exitCode !== null &&
        (!Number.isSafeInteger(exitCode) ||
          Number(exitCode) < 0 ||
          Number(exitCode) > 255)) ||
      (succeeded && exitCode !== 0) ||
      (value.status === "runtime-error" && exitCode === 0) ||
      ((value.status === "compile-error" || value.status === "quota-exceeded") &&
        exitCode !== null) ||
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
      exitCode: exitCode as number | null,
      ...(succeeded ? {} : { error: value.error as string }),
      retireGeneration: true,
    };
  } catch {
    return null;
  }
}

function resultForReceipt(receipt: ShellFinalReceipt): ShellRunResult {
  const base = { logs: receipt.logs, logStreams: receipt.logStreams };
  const withExitCode =
    receipt.exitCode === null
      ? base
      : { ...base, exitCode: receipt.exitCode };
  if (receipt.status === "succeeded") return withExitCode;
  const detail = receipt.error ?? "Shell execution failed.";
  if (receipt.status === "compile-error") {
    return { ...withExitCode, error: `Shell compile error: ${detail}` };
  }
  if (receipt.status === "quota-exceeded") {
    return {
      ...withExitCode,
      error: `Shell execution quota exceeded: ${detail}`,
    };
  }
  return { ...withExitCode, error: `Shell runtime error: ${detail}` };
}

function receiveWorkerMessage(handle: ShellWorkerHandle, value: unknown): void {
  if (handle.retired || workers.get(handle.namespace) !== handle) return;
  if (isOwnRecord(value) && value.type === "ready") {
    if (!parseReady(value) || handle.ready) {
      retireHandle(
        handle,
        { logs: [], error: "Shell runtime returned a malformed readiness receipt." },
        true,
      );
      return;
    }
    handle.ready = true;
    clearTimeout(handle.readyTimer);
    if (!handle.readySettled) {
      handle.readySettled = true;
      handle.settleReady(true);
    }
    setAvailability(handle.namespace, "available");
    return;
  }
  if (isOwnRecord(value) && value.type === "fatal") {
    const validFatal =
      value.protocol === SHELL_EXECUTION_PROTOCOL &&
      value.runtime === SHELL_EXECUTION_RUNTIME &&
      (value.phase === "load" ||
        value.phase === "request" ||
        value.phase === "execution");
    retireHandle(
      handle,
      { logs: [], error: "Shell runtime generation failed." },
      !handle.ready || !validFatal,
    );
    return;
  }
  if (!isOwnRecord(value) || typeof value.id !== "string") return;
  const pending = handle.pending.get(value.id);
  if (!pending) return;
  if (value.type === "output") {
    if (
      value.protocol !== SHELL_EXECUTION_PROTOCOL ||
      value.runtime !== SHELL_EXECUTION_RUNTIME ||
      typeof value.chunk !== "string" ||
      value.chunk === "" ||
      (value.stream !== "stdout" && value.stream !== "stderr")
    ) {
      retireHandle(handle, {
        logs: [],
        error: "Shell runtime returned malformed streamed output.",
      });
      return;
    }
    const bytes = encoder.encode(value.chunk).byteLength;
    pending.streamBytes += bytes;
    if (value.stream === "stderr") pending.stderrBytes += bytes;
    else pending.stdoutBytes += bytes;
    if (
      pending.streamBytes > pending.outputBytes ||
      pending.stdoutBytes > pending.outputBytes ||
      pending.stderrBytes > pending.outputBytes
    ) {
      retireHandle(handle, {
        logs: [],
        error: "Shell execution quota exceeded: streamed output exceeded its trusted budget.",
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
      error: "Shell runtime returned a malformed execution receipt.",
    });
    return;
  }
  settlePending(handle, pending, resultForReceipt(receipt));
  retireHandle(handle, { logs: [], interrupted: true });
}

function createHandle(namespace: string): ShellWorkerHandle | null {
  if (!browserSupportsShell()) {
    setAvailability(namespace, "unavailable");
    return null;
  }
  let worker: ShellWorkerLike;
  try {
    worker = workerFactory();
  } catch {
    setAvailability(namespace, "unavailable");
    return null;
  }
  // Every final receipt retires its generation, so replacing a proven runtime
  // before the next run is routine, not a capability change: dropping back to
  // "checking" here made the Code window retire the user's stdin after every
  // run. Runs still wait for readiness, and a replacement that fails to load
  // marks the runtime unavailable.
  if (availabilityState(namespace).status !== "available") {
    setAvailability(namespace, "checking");
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
    pending: new Map<string, PendingShellRun>(),
  } satisfies ShellWorkerHandle;
  handle.readyTimer = setTimeout(() => {
    retireHandle(
      handle,
      { logs: [], error: "Shell runtime did not become ready." },
      true,
    );
  }, READY_TIMEOUT_MS);
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () =>
    retireHandle(
      handle,
      { logs: [], error: "Shell runtime worker failed." },
      !handle.ready,
    );
  workers.set(namespace, handle);
  return handle;
}

function getHandle(namespace: string): ShellWorkerHandle | null {
  const current = workers.get(namespace);
  if (current && !current.retired) return current;
  return createHandle(namespace);
}

export function prepareShellRuntime(namespace: string): void {
  if (namespace) getHandle(namespace);
}

export function runShell(
  source: string,
  options: RunShellOptions,
): Promise<ShellRunResult> {
  const stdin = options.stdin ?? "";
  if (options.signal?.aborted) return Promise.resolve(interruptedResult());
  if (
    !options.namespace ||
    source.includes("\0") ||
    stdin.includes("\0") ||
    hasUnpairedSurrogate(source) ||
    hasUnpairedSurrogate(stdin) ||
    encoder.encode(source).byteLength > SHELL_MAX_SOURCE_BYTES ||
    encoder.encode(stdin).byteLength > SHELL_MAX_STDIN_BYTES
  ) {
    return Promise.resolve({
      logs: [],
      error: "Shell execution input exceeds its trusted budget.",
    });
  }
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const outputBytes = options.outputBytes ?? SHELL_MAX_OUTPUT_BYTES;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000 ||
    !Number.isSafeInteger(outputBytes) ||
    outputBytes <= 0 ||
    outputBytes > SHELL_MAX_OUTPUT_BYTES
  ) {
    return Promise.resolve({ logs: [], error: "Shell execution limits are invalid." });
  }

  const handle = getHandle(options.namespace);
  if (!handle) {
    return Promise.resolve({ logs: [], error: SHELL_RUNTIME_UNAVAILABLE_MESSAGE });
  }
  if (handle.pending.size !== 0) {
    return Promise.resolve({ logs: [], error: "Shell runtime is busy." });
  }

  const id = `shell:${handle.generation}:${++handle.sequence}`;
  return new Promise<ShellRunResult>((resolve) => {
    const pending: PendingShellRun = {
      id,
      resolve,
      timer: null,
      signal: options.signal,
      onOutput: options.onOutput,
      streamLimiter: createExecutionStreamLimiter(),
      streamBytes: 0,
      stdoutBytes: 0,
      stderrBytes: 0,
      timeoutMs,
      outputBytes,
      source,
      stdin,
      dispatched: false,
    };
    pending.abortListener = () => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(handle, interruptedResult());
    };
    handle.pending.set(id, pending);
    options.signal?.addEventListener("abort", pending.abortListener, { once: true });
    if (options.signal?.aborted) {
      pending.abortListener();
      return;
    }
    void handle.readyPromise.then((ready) => {
      if (handle.pending.get(id) !== pending) return;
      if (
        !ready ||
        handle.retired ||
        workers.get(options.namespace) !== handle ||
        !handle.ready
      ) {
        retireHandle(handle, {
          logs: [],
          error: options.signal?.aborted
            ? undefined
            : SHELL_RUNTIME_UNAVAILABLE_MESSAGE,
          ...(options.signal?.aborted ? { interrupted: true } : {}),
        });
        return;
      }
      if (options.signal?.aborted) {
        pending.abortListener?.();
        return;
      }
      pending.timer = setTimeout(() => {
        if (handle.pending.get(id) !== pending) return;
        retireHandle(handle, {
          logs: [],
          error: `Shell execution timed out after ${Math.round(timeoutMs / 1000)}s and was stopped.`,
          timedOut: true,
        });
      }, timeoutMs);
      try {
        pending.dispatched = true;
        handle.worker.postMessage({
          protocol: SHELL_EXECUTION_PROTOCOL,
          type: "run",
          id,
          source,
          stdin,
          timeoutMs,
          outputBytes,
        });
      } catch {
        retireHandle(handle, { logs: [], error: "Shell runtime transport failed." });
      }
    });
  });
}

/** Physically terminate the exact page generation, including a loading run. */
export function interruptShell(namespace: string): boolean {
  const handle = workers.get(namespace);
  if (!handle || handle.retired || handle.pending.size === 0) return false;
  retireHandle(handle, interruptedResult());
  return true;
}

/** Restart/dispose the page generation. The next request creates a fresh Worker. */
export function resetShellRuntime(namespace: string): void {
  const handle = workers.get(namespace);
  if (!handle) return;
  retireHandle(handle, {
    logs: [],
    error: "Shell runtime was restarted.",
    interrupted: true,
  });
}

export function __setShellWorkerFactoryForTests(
  factory: (() => ShellWorkerLike) | null,
): void {
  for (const handle of [...workers.values()]) {
    retireHandle(handle, interruptedResult());
  }
  availability.clear();
  workerFactory = factory ?? defaultWorkerFactory;
}

export function __resetShellRunnerForTests(): void {
  __setShellWorkerFactoryForTests(null);
  generation = 0;
}
