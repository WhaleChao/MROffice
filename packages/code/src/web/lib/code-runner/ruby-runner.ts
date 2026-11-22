/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Host authority for the dedicated Ruby WASI Worker.
 *
 * One page namespace owns one Worker generation. Stop/timeout/Restart
 * terminate that exact Worker; stale messages are ignored by handle identity.
 * Capability stays in `checking` until the pinned runtime has genuinely loaded.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  MAX_EXECUTION_LOG_CHARS,
  MAX_EXECUTION_LOG_LINE_CHARS,
  MAX_EXECUTION_LOG_LINES,
  MAX_EXECUTION_TEXT_CHARS,
  boundExecutionResult,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";
import {
  RUBY_EXECUTION_PROTOCOL,
  RUBY_EXECUTION_RUNTIME,
  RUBY_MAX_OUTPUT_BYTES,
  RUBY_MAX_SOURCE_BYTES,
  RUBY_MAX_STDIN_BYTES,
  type RubyExecutionStatus,
} from "./ruby-execution-runtime";

export { RUBY_EXECUTION_PROTOCOL, RUBY_EXECUTION_RUNTIME };

export const RUBY_RUNTIME_UNAVAILABLE_MESSAGE =
  "Ruby 4.0 WASI is unavailable here.";

const READY_TIMEOUT_MS = 30_000;
const DEFAULT_TIMEOUT_MS = 5_000;
const encoder = new TextEncoder();

export type RubyRuntimeAvailability =
  | "checking"
  | "available"
  | "unavailable";

export interface RubyWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

interface RubyFinalReceipt {
  status: RubyExecutionStatus;
  logs: string[];
  logStreams: Array<"stdout" | "stderr">;
  error?: string;
  retireGeneration: boolean;
}

interface PendingRubyRun {
  id: string;
  resolve: (result: RunResult) => void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
  outputBytes: number;
}

interface RubyWorkerHandle {
  worker: RubyWorkerLike;
  namespace: string;
  generation: number;
  sequence: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingRubyRun>;
}

interface AvailabilityState {
  status: RubyRuntimeAvailability;
  listeners: Set<() => void>;
}

export interface RunRubyOptions {
  namespace: string;
  stdin?: string;
  signal?: AbortSignal;
  timeoutMs?: number;
  outputBytes?: number;
}

let generation = 0;
const workers = new Map<string, RubyWorkerHandle>();
const availability = new Map<string, AvailabilityState>();

const defaultWorkerFactory = (): RubyWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("Ruby execution requires a browser Worker.");
  }
  return new Worker(new URL("./ruby-execution-worker.ts", import.meta.url), {
    name: "mrnote-ruby-execution",
    type: "module",
  });
};

let workerFactory: () => RubyWorkerLike = defaultWorkerFactory;

function browserSupportsRuby(): boolean {
  return (
    (workerFactory !== defaultWorkerFactory || typeof Worker !== "undefined") &&
    typeof WebAssembly !== "undefined" &&
    typeof WebAssembly.compile === "function" &&
    typeof TextEncoder !== "undefined" &&
    typeof TextDecoder !== "undefined"
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

function setAvailability(
  namespace: string,
  status: RubyRuntimeAvailability,
): void {
  const state = availabilityState(namespace);
  if (state.status === status) return;
  state.status = status;
  for (const listener of [...state.listeners]) listener();
}

export function getRubyRuntimeAvailability(
  namespace: string,
): RubyRuntimeAvailability {
  if (!namespace || !browserSupportsRuby()) return "unavailable";
  return availabilityState(namespace).status;
}

export function subscribeRubyRuntimeAvailability(
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

function interruptedResult(): RunResult {
  return { logs: [], interrupted: true };
}

function settlePending(
  handle: RubyWorkerHandle,
  pending: PendingRubyRun,
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

function settleAll(handle: RubyWorkerHandle, result: RunResult): void {
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, result);
  }
}

function retireHandle(
  handle: RubyWorkerHandle,
  result: RunResult,
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
    // A retired generation cannot be reused; termination remains best effort.
  }
  if (workers.get(handle.namespace) === handle) {
    workers.delete(handle.namespace);
  }
  if (runtimeUnavailable) setAvailability(handle.namespace, "unavailable");
}

function parseReady(value: unknown): boolean {
  if (!isOwnRecord(value)) return false;
  return (
    value.protocol === RUBY_EXECUTION_PROTOCOL &&
    value.runtime === RUBY_EXECUTION_RUNTIME &&
    value.type === "ready"
  );
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

function parseFinalReceipt(
  value: unknown,
  pending: PendingRubyRun,
): RubyFinalReceipt | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      value.protocol !== RUBY_EXECUTION_PROTOCOL ||
      value.runtime !== RUBY_EXECUTION_RUNTIME ||
      value.type !== "result" ||
      value.id !== pending.id ||
      (value.status !== "succeeded" &&
        value.status !== "compile-error" &&
        value.status !== "runtime-error" &&
        value.status !== "quota-exceeded") ||
      (value.retireGeneration !== undefined &&
        value.retireGeneration !== true)
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
      retireGeneration: value.retireGeneration === true,
    };
  } catch {
    return null;
  }
}

function resultForReceipt(receipt: RubyFinalReceipt): RunResult {
  const base = { logs: receipt.logs, logStreams: receipt.logStreams };
  if (receipt.status === "succeeded") return boundExecutionResult(base);
  const detail = receipt.error ?? "Ruby execution failed.";
  if (receipt.status === "compile-error") {
    return boundExecutionResult({ ...base, error: `Ruby compile error: ${detail}` });
  }
  if (receipt.status === "quota-exceeded") {
    return boundExecutionResult({
      ...base,
      error: `Ruby execution quota exceeded: ${detail}`,
    });
  }
  return boundExecutionResult({ ...base, error: `Ruby runtime error: ${detail}` });
}

function receiveWorkerMessage(handle: RubyWorkerHandle, value: unknown): void {
  if (handle.retired || workers.get(handle.namespace) !== handle) return;
  if (isOwnRecord(value) && value.type === "ready") {
    if (!parseReady(value) || handle.ready) {
      retireHandle(
        handle,
        { logs: [], error: "Ruby runtime returned a malformed readiness receipt." },
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
      value.protocol === RUBY_EXECUTION_PROTOCOL &&
      value.runtime === RUBY_EXECUTION_RUNTIME &&
      (value.phase === "load" ||
        value.phase === "request" ||
        value.phase === "execution");
    retireHandle(
      handle,
      { logs: [], error: "Ruby runtime generation failed." },
      !handle.ready || !validFatal,
    );
    return;
  }
  if (!isOwnRecord(value) || typeof value.id !== "string") return;
  const pending = handle.pending.get(value.id);
  if (!pending) return;
  const receipt = parseFinalReceipt(value, pending);
  if (!receipt) {
    retireHandle(handle, {
      logs: [],
      error: "Ruby runtime returned a malformed execution receipt.",
    });
    return;
  }
  settlePending(handle, pending, resultForReceipt(receipt));
  if (receipt.retireGeneration) {
    retireHandle(handle, {
      logs: [],
      error: "Ruby runtime generation was retired.",
    });
  }
}

function createHandle(namespace: string): RubyWorkerHandle | null {
  if (!browserSupportsRuby()) {
    setAvailability(namespace, "unavailable");
    return null;
  }
  let worker: RubyWorkerLike;
  try {
    worker = workerFactory();
  } catch {
    setAvailability(namespace, "unavailable");
    return null;
  }
  // Stop and quota receipts retire their generation. Replacing a proven
  // runtime is routine, not a capability change: dropping back to "checking"
  // made the Code window retire the user's stdin. Runs still wait for
  // readiness, and a replacement that fails to load marks it unavailable.
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
    pending: new Map<string, PendingRubyRun>(),
  } satisfies RubyWorkerHandle;
  handle.readyTimer = setTimeout(() => {
    retireHandle(
      handle,
      { logs: [], error: "Ruby runtime did not become ready." },
      true,
    );
  }, READY_TIMEOUT_MS);
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () =>
    retireHandle(
      handle,
      { logs: [], error: "Ruby runtime worker failed." },
      !handle.ready,
    );
  workers.set(namespace, handle);
  return handle;
}

function getHandle(namespace: string): RubyWorkerHandle | null {
  const current = workers.get(namespace);
  if (current && !current.retired) return current;
  return createHandle(namespace);
}

export function prepareRubyRuntime(namespace: string): void {
  if (!namespace) return;
  // A language switch or page reopen is an explicit retry boundary for a
  // transient asset/Worker load failure. Unsupported browsers still fail
  // closed immediately in createHandle().
  getHandle(namespace);
}

function waitForReady(
  handle: RubyWorkerHandle,
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

export async function runRuby(
  source: string,
  options: RunRubyOptions,
): Promise<RunResult> {
  const stdin = options.stdin ?? "";
  if (options.signal?.aborted) return interruptedResult();
  if (
    !options.namespace ||
    source.includes("\0") ||
    stdin.includes("\0") ||
    hasUnpairedSurrogate(source) ||
    hasUnpairedSurrogate(stdin) ||
    encoder.encode(source).byteLength > RUBY_MAX_SOURCE_BYTES ||
    encoder.encode(stdin).byteLength > RUBY_MAX_STDIN_BYTES
  ) {
    return { logs: [], error: "Ruby execution input exceeds its trusted budget." };
  }
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const outputBytes = options.outputBytes ?? RUBY_MAX_OUTPUT_BYTES;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000 ||
    !Number.isSafeInteger(outputBytes) ||
    outputBytes <= 0 ||
    outputBytes > RUBY_MAX_OUTPUT_BYTES
  ) {
    return { logs: [], error: "Ruby execution limits are invalid." };
  }

  const handle = getHandle(options.namespace);
  if (!handle || !(await waitForReady(handle, options.signal))) {
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: RUBY_RUNTIME_UNAVAILABLE_MESSAGE };
  }
  if (
    options.signal?.aborted ||
    handle.retired ||
    workers.get(options.namespace) !== handle ||
    !handle.ready
  ) {
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: RUBY_RUNTIME_UNAVAILABLE_MESSAGE };
  }
  if (handle.pending.size !== 0) {
    return { logs: [], error: "Ruby runtime is busy." };
  }

  const id = `ruby:${handle.generation}:${++handle.sequence}`;
  return new Promise<RunResult>((resolve) => {
    let dispatched = false;
    const pending: PendingRubyRun = {
      id,
      resolve,
      timer: 0 as unknown as ReturnType<typeof setTimeout>,
      signal: options.signal,
      outputBytes,
    };
    pending.timer = setTimeout(() => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(handle, {
        logs: [],
        error: `Ruby execution timed out after ${Math.round(timeoutMs / 1000)}s and was stopped.`,
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
        protocol: RUBY_EXECUTION_PROTOCOL,
        type: "run",
        id,
        source,
        stdin,
        outputBytes,
      });
    } catch {
      retireHandle(handle, {
        logs: [],
        error: "Ruby runtime transport failed.",
      });
    }
  });
}

/** Physically terminate the exact page generation that owns an active run. */
export function interruptRuby(namespace: string): boolean {
  const handle = workers.get(namespace);
  if (!handle || handle.retired || handle.pending.size === 0) return false;
  retireHandle(handle, interruptedResult());
  return true;
}

/** Restart/dispose the interpreter. The next request creates a fresh Worker. */
export function resetRubyRuntime(namespace: string): void {
  const handle = workers.get(namespace);
  if (!handle) return;
  retireHandle(handle, {
    logs: [],
    error: "Ruby runtime was restarted.",
    interrupted: true,
  });
}

export function __setRubyWorkerFactoryForTests(
  factory: (() => RubyWorkerLike) | null,
): void {
  for (const handle of [...workers.values()]) {
    retireHandle(handle, { logs: [], interrupted: true });
  }
  availability.clear();
  workerFactory = factory ?? defaultWorkerFactory;
}

export function __resetRubyRunnerForTests(): void {
  __setRubyWorkerFactoryForTests(null);
  generation = 0;
}
