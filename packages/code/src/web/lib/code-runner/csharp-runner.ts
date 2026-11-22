/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  MAX_EXECUTION_TEXT_CHARS,
  boundExecutionResult,
  createExecutionLogCapture,
  type ExecutionLogCapture,
} from "@mrnote/core/web/lib/code-runner/execution-output-budget.ts";
import {
  CSHARP_DEFAULT_TIMEOUT_MS,
  CSHARP_EXECUTION_PROTOCOL,
  CSHARP_EXECUTION_RUNTIME,
  CSHARP_MAX_OUTPUT_BYTES,
  CSHARP_MAX_SOURCE_BYTES,
  CSHARP_MAX_STDIN_BYTES,
  CSHARP_WORKER_URL,
} from "./csharp-execution-runtime";

export {
  CSHARP_EXECUTION_PROTOCOL,
  CSHARP_EXECUTION_RUNTIME,
};

export const CSHARP_RUNTIME_UNAVAILABLE_MESSAGE =
  "The trusted .NET 10 C# Worker is unavailable here.";

const READY_TIMEOUT_MS = 60_000;
const MAX_DIAGNOSTICS = 100;
const encoder = new TextEncoder();

export type CSharpRuntimeAvailability =
  | "checking"
  | "available"
  | "unavailable";

export interface CSharpWorkerLike {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  onerror: ((event: ErrorEvent) => void) | null;
  postMessage(message: unknown): void;
  terminate(): void;
}

export interface CSharpDiagnostic {
  code: string;
  message: string;
  line: number | null;
  column: number | null;
}

export interface CSharpCompileResult {
  diagnostics: readonly CSharpDiagnostic[];
  error?: string;
  interrupted?: boolean;
}

interface CSharpFinalReceipt {
  status: "succeeded" | "compile-error" | "runtime-error" | "quota-exceeded";
  success: boolean;
  exitCode: number | null;
  stdout: string;
  stderr: string;
  error: string | null;
  diagnostics: CSharpDiagnostic[];
}

interface CSharpCompileReceipt {
  status: "clean" | "compile-error" | "quota-exceeded";
  success: boolean;
  error: string | null;
  diagnostics: CSharpDiagnostic[];
}

type CSharpRunResult = RunResult & { readonly exitCode?: number };

interface PendingCSharpRun {
  kind: "run";
  id: string;
  resolve: (result: CSharpRunResult) => void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
  outputBytes: number;
}

interface PendingCSharpCompile {
  kind: "compile";
  id: string;
  resolve: (result: CSharpCompileResult) => void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
}

type PendingCSharpRequest = PendingCSharpRun | PendingCSharpCompile;

interface CSharpWorkerHandle {
  worker: CSharpWorkerLike;
  namespace: string;
  generation: number;
  sequence: number;
  ready: boolean;
  retired: boolean;
  readySettled: boolean;
  readyPromise: Promise<boolean>;
  settleReady: (ready: boolean) => void;
  readyTimer: ReturnType<typeof setTimeout>;
  pending: Map<string, PendingCSharpRequest>;
}

interface AvailabilityState {
  status: CSharpRuntimeAvailability;
  listeners: Set<() => void>;
}

export interface RunCSharpOptions {
  namespace: string;
  stdin?: string;
  signal?: AbortSignal;
  timeoutMs?: number;
  outputBytes?: number;
}

export interface CompileCSharpOptions {
  namespace: string;
  signal?: AbortSignal;
  timeoutMs?: number;
}

let generation = 0;
const workers = new Map<string, CSharpWorkerHandle>();
const availability = new Map<string, AvailabilityState>();
// Whether this page has already booted the pinned runtime once. Every window
// used to keep its own idle generation booted just to learn that the runtime
// loads; a notebook with a dozen C# windows then booted a dozen .NET runtimes
// at once (each tens of MB of WebAssembly), which starved the page's own
// scripts of connections and left it blank. One boot answers for all.
let runtimeVerified = false;
let verifyingHandle: CSharpWorkerHandle | null = null;

const defaultWorkerFactory = (): CSharpWorkerLike => {
  if (typeof Worker === "undefined") {
    throw new Error("C# execution requires a browser Worker.");
  }
  return new Worker(CSHARP_WORKER_URL, {
    name: "mrnote-csharp-execution",
    type: "module",
  });
};

let workerFactory: () => CSharpWorkerLike = defaultWorkerFactory;

function browserSupportsCSharp(): boolean {
  return (
    (workerFactory !== defaultWorkerFactory || typeof Worker !== "undefined") &&
    typeof WebAssembly !== "undefined" &&
    typeof WebAssembly.compile === "function" &&
    typeof TextEncoder !== "undefined"
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
  status: CSharpRuntimeAvailability,
): void {
  const state = availabilityState(namespace);
  if (state.status === status) return;
  state.status = status;
  for (const listener of [...state.listeners]) listener();
}

export function getCSharpRuntimeAvailability(
  namespace: string,
): CSharpRuntimeAvailability {
  if (!namespace || !browserSupportsCSharp()) return "unavailable";
  return availabilityState(namespace).status;
}

export function subscribeCSharpRuntimeAvailability(
  namespace: string,
  listener: () => void,
): () => void {
  const state = availabilityState(namespace);
  state.listeners.add(listener);
  return () => state.listeners.delete(listener);
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
    for (let index = 0; index < value.length; index += 1) {
      const descriptor = descriptors[String(index)];
      if (!descriptor || !("value" in descriptor)) return null;
    }
    return Array.from(value);
  } catch {
    return null;
  }
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

function interruptedResult(): CSharpRunResult {
  return { logs: [], interrupted: true };
}

function settlePending(
  handle: CSharpWorkerHandle,
  pending: PendingCSharpRequest,
  runResult: CSharpRunResult,
  compileResult: CSharpCompileResult,
): void {
  if (handle.pending.get(pending.id) !== pending) return;
  handle.pending.delete(pending.id);
  clearTimeout(pending.timer);
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  if (pending.kind === "compile") {
    pending.resolve(compileResult);
    return;
  }
  const bounded = boundExecutionResult(runResult);
  pending.resolve(
    runResult.exitCode === undefined
      ? bounded
      : ({ ...bounded, exitCode: runResult.exitCode } as CSharpRunResult),
  );
}

function settleAll(
  handle: CSharpWorkerHandle,
  runResult: CSharpRunResult,
  compileResult: CSharpCompileResult,
): void {
  for (const pending of [...handle.pending.values()]) {
    settlePending(handle, pending, runResult, compileResult);
  }
}

function retireHandle(
  handle: CSharpWorkerHandle,
  result: CSharpRunResult,
  runtimeUnavailable = false,
  compileResult: CSharpCompileResult = {
    diagnostics: [],
    ...(result.error ? { error: result.error } : {}),
    ...(result.interrupted ? { interrupted: true } : {}),
  },
): void {
  if (handle.retired) return;
  handle.retired = true;
  clearTimeout(handle.readyTimer);
  if (!handle.readySettled) {
    handle.readySettled = true;
    handle.settleReady(false);
  }
  settleAll(handle, result, compileResult);
  try {
    handle.worker.terminate();
  } catch {
    // The exact generation is already retired; termination is best effort.
  }
  if (workers.get(handle.namespace) === handle) workers.delete(handle.namespace);
  if (runtimeUnavailable) setAvailability(handle.namespace, "unavailable");
}

function parseReady(value: unknown): boolean {
  if (!isOwnRecord(value)) return false;
  try {
    return (
      value.protocol === CSHARP_EXECUTION_PROTOCOL &&
      value.runtime === CSHARP_EXECUTION_RUNTIME &&
      value.type === "ready" &&
      value.ok === true &&
      value.error === undefined
    );
  } catch {
    return false;
  }
}

function parseDiagnostic(value: unknown): CSharpDiagnostic | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      typeof value.code !== "string" ||
      value.code.length === 0 ||
      value.code.length > 32 ||
      hasUnpairedSurrogate(value.code) ||
      typeof value.message !== "string" ||
      value.message.length === 0 ||
      value.message.length > MAX_EXECUTION_TEXT_CHARS ||
      hasUnpairedSurrogate(value.message) ||
      ((value.line === null) !== (value.column === null)) ||
      (value.line !== null &&
        (!Number.isSafeInteger(value.line) || Number(value.line) <= 0)) ||
      (value.column !== null &&
        (!Number.isSafeInteger(value.column) || Number(value.column) <= 0))
    ) {
      return null;
    }
    return {
      code: value.code,
      message: value.message,
      line: value.line === null ? null : Number(value.line),
      column: value.column === null ? null : Number(value.column),
    };
  } catch {
    return null;
  }
}

/**
 * The worker's report that the runtime itself stopped while serving this
 * request (`ok: false`, e.g. "Program terminated with exit(1)" after a Mono
 * assertion), as opposed to a malformed receipt. The program is what stopped
 * it, so this request fails with that reason; the runtime stays available,
 * and the next request gets a fresh Worker as every request does.
 */
function parseRuntimeStopped(
  value: unknown,
  id: string,
  type: "result" | "compile-result",
): string | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      value.protocol !== CSHARP_EXECUTION_PROTOCOL ||
      value.runtime !== CSHARP_EXECUTION_RUNTIME ||
      value.type !== type ||
      value.id !== id ||
      value.ok !== false ||
      typeof value.error !== "string" ||
      value.error.length === 0 ||
      value.error.length > MAX_EXECUTION_TEXT_CHARS ||
      hasUnpairedSurrogate(value.error)
    ) {
      return null;
    }
    return `C# runtime stopped unexpectedly: ${value.error}`;
  } catch {
    return null;
  }
}

function parseFinalReceipt(
  value: unknown,
  pending: PendingCSharpRun,
): CSharpFinalReceipt | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      value.protocol !== CSHARP_EXECUTION_PROTOCOL ||
      value.runtime !== CSHARP_EXECUTION_RUNTIME ||
      value.type !== "result" ||
      value.id !== pending.id ||
      value.ok !== true ||
      !isOwnRecord(value.receipt)
    ) {
      return null;
    }
    const receipt = value.receipt;
    const diagnosticsRaw = ownArray(receipt.diagnostics);
    if (
      (receipt.status !== "succeeded" &&
        receipt.status !== "compile-error" &&
        receipt.status !== "runtime-error" &&
        receipt.status !== "quota-exceeded") ||
      typeof receipt.success !== "boolean" ||
      (receipt.exitCode !== null &&
        (!Number.isSafeInteger(receipt.exitCode) ||
          Number(receipt.exitCode) < 0 ||
          Number(receipt.exitCode) > 255)) ||
      typeof receipt.stdout !== "string" ||
      typeof receipt.stderr !== "string" ||
      (receipt.error !== null &&
        (typeof receipt.error !== "string" ||
          receipt.error.length === 0 ||
          receipt.error.length > MAX_EXECUTION_TEXT_CHARS ||
          hasUnpairedSurrogate(receipt.error))) ||
      !diagnosticsRaw ||
      diagnosticsRaw.length > MAX_DIAGNOSTICS ||
      hasUnpairedSurrogate(receipt.stdout) ||
      hasUnpairedSurrogate(receipt.stderr)
    ) {
      return null;
    }
    const outputBytes =
      encoder.encode(receipt.stdout).byteLength +
      encoder.encode(receipt.stderr).byteLength;
    if (outputBytes > pending.outputBytes || outputBytes > CSHARP_MAX_OUTPUT_BYTES) {
      return null;
    }
    const diagnostics: CSharpDiagnostic[] = [];
    for (const item of diagnosticsRaw) {
      const diagnostic = parseDiagnostic(item);
      if (!diagnostic) return null;
      diagnostics.push(diagnostic);
    }
    const exitCode = receipt.exitCode === null ? null : Number(receipt.exitCode);
    if (
      (receipt.success !== (receipt.status === "succeeded")) ||
      (receipt.success &&
        (exitCode !== 0 || receipt.error !== null || diagnostics.length > 0)) ||
      (!receipt.success && receipt.error === null) ||
      (diagnostics.length > 0 && exitCode !== null) ||
      ((receipt.status === "compile-error" ||
        receipt.status === "quota-exceeded") &&
        exitCode !== null)
    ) {
      return null;
    }
    return {
      status: receipt.status,
      success: receipt.success,
      exitCode,
      stdout: receipt.stdout,
      stderr: receipt.stderr,
      error: receipt.error,
      diagnostics,
    };
  } catch {
    return null;
  }
}

function parseCompileReceipt(
  value: unknown,
  pending: PendingCSharpCompile,
): CSharpCompileReceipt | null {
  if (!isOwnRecord(value)) return null;
  try {
    if (
      value.protocol !== CSHARP_EXECUTION_PROTOCOL ||
      value.runtime !== CSHARP_EXECUTION_RUNTIME ||
      value.type !== "compile-result" ||
      value.id !== pending.id ||
      value.ok !== true ||
      !isOwnRecord(value.receipt)
    ) {
      return null;
    }
    const receipt = value.receipt;
    const diagnosticsRaw = ownArray(receipt.diagnostics);
    if (
      (receipt.status !== "clean" &&
        receipt.status !== "compile-error" &&
        receipt.status !== "quota-exceeded") ||
      typeof receipt.success !== "boolean" ||
      (receipt.error !== null &&
        (typeof receipt.error !== "string" ||
          receipt.error.length === 0 ||
          receipt.error.length > MAX_EXECUTION_TEXT_CHARS ||
          hasUnpairedSurrogate(receipt.error))) ||
      !diagnosticsRaw ||
      diagnosticsRaw.length > MAX_DIAGNOSTICS ||
      receipt.stdout !== undefined ||
      receipt.stderr !== undefined ||
      receipt.exitCode !== undefined
    ) {
      return null;
    }
    const diagnostics: CSharpDiagnostic[] = [];
    for (const item of diagnosticsRaw) {
      const diagnostic = parseDiagnostic(item);
      if (!diagnostic) return null;
      diagnostics.push(diagnostic);
    }
    if (
      receipt.success !== (receipt.status === "clean") ||
      (receipt.success && (receipt.error !== null || diagnostics.length > 0)) ||
      (!receipt.success && receipt.error === null) ||
      (receipt.status === "quota-exceeded" && diagnostics.length > 0)
    ) {
      return null;
    }
    return {
      status: receipt.status,
      success: receipt.success,
      error: receipt.error,
      diagnostics,
    };
  } catch {
    return null;
  }
}

// The Worker returns each stream whole, already bounded by encoded bytes in
// parseFinalReceipt. Many short lines or one very long line is ordinary program
// output, so the shared capture truncates it with the visible marker; treating
// it as a malformed receipt retired the runtime for the whole page.
function appendStream(
  value: string,
  stream: "stdout" | "stderr",
  capture: ExecutionLogCapture,
): void {
  if (value.length === 0) return;
  const lines = value.split(/\r?\n/);
  if (lines.at(-1) === "") lines.pop();
  for (const line of lines) {
    if (capture.push(line, stream) === null) return;
  }
}

// The host reports every uncaught exception with the fixed text below and
// appends the exception's ToString() to stderr. That block moves into the
// error, as Ruby and JavaScript report theirs, so it is shown once; the frames
// under the user's own frames are the host's reflection call into the entry
// point, which a `dotnet run` trace never shows.
const HOST_UNCAUGHT_ERROR = "C# runtime error.";
const HOST_INVOKE_FRAME = /^\s+at System\.(?:Reflection|RuntimeMethodHandle)\./;

function uncaughtException(
  stderr: string,
): { trace: string; stderr: string } | null {
  const lines = stderr.split(/\r?\n/);
  if (lines.at(-1) === "") lines.pop();
  let start = lines.length;
  while (start > 0 && /^\s+at |^\s+--- End of /.test(lines[start - 1])) {
    start -= 1;
  }
  // Inner exceptions print as "Outer: msg\n ---> Inner: msg"; the header is
  // the first line of that chain.
  while (start > 0 && /^\s*---> /.test(lines[start - 1] ?? "")) start -= 1;
  const headerIndex = start - 1;
  const header = lines[headerIndex];
  if (
    header === undefined ||
    !/^[A-Za-z_][\w.`+]*(?:Exception|Error)\b/.test(header)
  ) {
    return null;
  }
  const trace = lines
    .slice(headerIndex)
    .filter((line) => !HOST_INVOKE_FRAME.test(line));
  const before = lines.slice(0, headerIndex);
  return {
    trace: trace.join("\n"),
    stderr: before.length > 0 ? `${before.join("\n")}\n` : "",
  };
}

function resultForReceipt(receipt: CSharpFinalReceipt): CSharpRunResult {
  const uncaught =
    receipt.status === "runtime-error" && receipt.error === HOST_UNCAUGHT_ERROR
      ? uncaughtException(receipt.stderr)
      : null;
  const capture = createExecutionLogCapture();
  appendStream(receipt.stdout, "stdout", capture);
  appendStream(uncaught?.stderr ?? receipt.stderr, "stderr", capture);
  const base: CSharpRunResult = {
    logs: capture.logs,
    logStreams: capture.logStreams,
    ...(receipt.exitCode === null ? {} : { exitCode: receipt.exitCode }),
  };
  if (receipt.status === "succeeded") return base;
  if (receipt.status === "compile-error") {
    const first = receipt.diagnostics[0];
    const location = first
      ? first.line === null
        ? "main.cs"
        : `main.cs:${first.line}:${first.column ?? 1}`
      : "main.cs";
    const detail = first?.message ?? receipt.error ?? "C# compilation failed.";
    return {
      ...base,
      error: `C# compile error: ${location}: ${detail}`,
    };
  }
  if (receipt.status === "quota-exceeded") {
    return {
      ...base,
      error: `C# execution quota exceeded: ${receipt.error ?? "C# execution exceeded its quota."}`,
    };
  }
  if (receipt.error === HOST_UNCAUGHT_ERROR) {
    return {
      ...base,
      error: uncaught
        ? `C# runtime error: ${uncaught.trace}`
        : HOST_UNCAUGHT_ERROR,
    };
  }
  return {
    ...base,
    error: `C# runtime error: ${receipt.error ?? "C# execution failed."}`,
  };
}

function compileResultForReceipt(
  receipt: CSharpCompileReceipt,
): CSharpCompileResult {
  return {
    diagnostics: receipt.diagnostics,
    ...(receipt.error ? { error: receipt.error } : {}),
  };
}

function receiveWorkerMessage(handle: CSharpWorkerHandle, value: unknown): void {
  if (handle.retired || workers.get(handle.namespace) !== handle) return;
  if (isOwnRecord(value) && value.type === "ready") {
    if (!parseReady(value) || handle.ready) {
      retireHandle(
        handle,
        { logs: [], error: "C# runtime returned malformed readiness." },
        true,
      );
      return;
    }
    handle.ready = true;
    runtimeVerified = true;
    clearTimeout(handle.readyTimer);
    if (!handle.readySettled) {
      handle.readySettled = true;
      handle.settleReady(true);
    }
    setAvailability(handle.namespace, "available");
    return;
  }
  if (!isOwnRecord(value) || typeof value.id !== "string") return;
  const pending = handle.pending.get(value.id);
  if (!pending) return;
  if (pending.kind === "compile") {
    const stopped = parseRuntimeStopped(value, pending.id, "compile-result");
    if (stopped) {
      settlePending(handle, pending, { logs: [] }, { diagnostics: [], error: stopped });
      retireHandle(handle, { logs: [], interrupted: true });
      return;
    }
    const receipt = parseCompileReceipt(value, pending);
    if (!receipt) {
      retireHandle(
        handle,
        { logs: [], error: "C# runtime returned a malformed compile receipt." },
        true,
        {
          diagnostics: [],
          error: "C# runtime returned a malformed compile receipt.",
        },
      );
      return;
    }
    const compileResult = compileResultForReceipt(receipt);
    settlePending(handle, pending, { logs: [] }, compileResult);
    retireHandle(handle, { logs: [], interrupted: true });
    return;
  }
  const stopped = parseRuntimeStopped(value, pending.id, "result");
  if (stopped) {
    settlePending(handle, pending, { logs: [], error: stopped }, { diagnostics: [] });
    retireHandle(handle, { logs: [], interrupted: true });
    return;
  }
  const receipt = parseFinalReceipt(value, pending);
  const result = receipt ? resultForReceipt(receipt) : null;
  if (!result) {
    retireHandle(
      handle,
      { logs: [], error: "C# runtime returned a malformed receipt." },
      true,
    );
    return;
  }
  settlePending(handle, pending, result, { diagnostics: [] });
  retireHandle(handle, { logs: [], interrupted: true });
}

function createHandle(namespace: string): CSharpWorkerHandle | null {
  let worker: CSharpWorkerLike;
  try {
    worker = workerFactory();
  } catch {
    setAvailability(namespace, "unavailable");
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
    pending: new Map<string, PendingCSharpRequest>(),
  } satisfies CSharpWorkerHandle;
  handle.readyTimer = setTimeout(
    () =>
      retireHandle(
        handle,
        { logs: [], error: "C# runtime did not become ready." },
        true,
      ),
    READY_TIMEOUT_MS,
  );
  worker.onmessage = (event) => receiveWorkerMessage(handle, event.data);
  worker.onerror = () =>
    retireHandle(
      handle,
      { logs: [], error: "C# runtime Worker failed." },
      true,
    );
  workers.set(namespace, handle);
  return handle;
}

function getHandle(namespace: string): CSharpWorkerHandle | null {
  const current = workers.get(namespace);
  if (current && !current.retired) return current;
  return createHandle(namespace);
}

function waitForReady(
  handle: CSharpWorkerHandle,
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

export function prepareCSharpRuntime(namespace: string): void {
  if (!namespace || !browserSupportsCSharp()) {
    setAvailability(namespace, "unavailable");
    return;
  }
  if (availabilityState(namespace).status === "available") return;
  // Runs and compiles use single-use generations anyway, so a window needs no
  // idle generation of its own once the page knows the runtime boots.
  if (runtimeVerified) {
    setAvailability(namespace, "available");
    return;
  }
  const verifying = verifyingHandle;
  if (verifying && !verifying.retired && verifying.namespace !== namespace) {
    void verifying.readyPromise.then((ready) => {
      if (ready || runtimeVerified) setAvailability(namespace, "available");
      // The shared boot failed: this window tries its own, as before.
      else prepareCSharpRuntime(namespace);
    });
    return;
  }
  const handle = getHandle(namespace);
  if (!handle) return;
  if (!runtimeVerified) verifyingHandle = handle;
  void handle.readyPromise.then((ready) => {
    if (verifyingHandle === handle) verifyingHandle = null;
    if (!ready && workers.get(namespace) === handle) {
      setAvailability(namespace, "unavailable");
    }
  });
}

export async function runCSharp(
  source: string,
  options: RunCSharpOptions,
): Promise<CSharpRunResult> {
  const stdin = options.stdin ?? "";
  if (options.signal?.aborted) return interruptedResult();
  if (
    !options.namespace ||
    hasUnpairedSurrogate(source) ||
    hasUnpairedSurrogate(stdin) ||
    encoder.encode(source).byteLength > CSHARP_MAX_SOURCE_BYTES ||
    encoder.encode(stdin).byteLength > CSHARP_MAX_STDIN_BYTES
  ) {
    return { logs: [], error: "C# execution input exceeds its trusted budget." };
  }
  const timeoutMs = options.timeoutMs ?? CSHARP_DEFAULT_TIMEOUT_MS;
  const outputBytes = options.outputBytes ?? CSHARP_MAX_OUTPUT_BYTES;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000 ||
    !Number.isSafeInteger(outputBytes) ||
    outputBytes <= 0 ||
    outputBytes > CSHARP_MAX_OUTPUT_BYTES
  ) {
    return { logs: [], error: "C# execution limits are invalid." };
  }
  const handle = getHandle(options.namespace);
  if (!handle || !(await waitForReady(handle, options.signal))) {
    if (options.signal?.aborted && handle) {
      retireHandle(handle, interruptedResult());
      return interruptedResult();
    }
    return { logs: [], error: CSHARP_RUNTIME_UNAVAILABLE_MESSAGE };
  }
  if (
    options.signal?.aborted ||
    handle.retired ||
    workers.get(options.namespace) !== handle ||
    !handle.ready
  ) {
    if (options.signal?.aborted) retireHandle(handle, interruptedResult());
    return options.signal?.aborted
      ? interruptedResult()
      : { logs: [], error: "C# runtime generation is unavailable." };
  }
  const id = `csharp:${handle.generation}:${++handle.sequence}`;
  return new Promise<CSharpRunResult>((resolve) => {
    const pending: PendingCSharpRun = {
      kind: "run",
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
        error: `C# execution timed out after ${Math.round(timeoutMs / 1000)}s and was stopped.`,
        timedOut: true,
      });
    }, timeoutMs);
    pending.abortListener = () => {
      if (handle.pending.get(id) !== pending) return;
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
      handle.worker.postMessage({
        protocol: CSHARP_EXECUTION_PROTOCOL,
        type: "run",
        id,
        source,
        stdin,
      });
    } catch {
      retireHandle(handle, {
        logs: [],
        error: "C# runtime transport failed.",
      });
    }
  });
}

export async function compileCSharp(
  source: string,
  options: CompileCSharpOptions,
): Promise<CSharpCompileResult> {
  if (options.signal?.aborted) {
    return { diagnostics: [], interrupted: true };
  }
  if (
    !options.namespace ||
    hasUnpairedSurrogate(source) ||
    encoder.encode(source).byteLength > CSHARP_MAX_SOURCE_BYTES
  ) {
    return {
      diagnostics: [],
      error: "C# compile input exceeds its trusted budget.",
    };
  }
  const timeoutMs = options.timeoutMs ?? CSHARP_DEFAULT_TIMEOUT_MS;
  if (
    !Number.isSafeInteger(timeoutMs) ||
    timeoutMs <= 0 ||
    timeoutMs > 60_000
  ) {
    return { diagnostics: [], error: "C# compile limits are invalid." };
  }
  const handle = getHandle(options.namespace);
  if (!handle || !(await waitForReady(handle, options.signal))) {
    if (options.signal?.aborted && handle) {
      retireHandle(handle, interruptedResult());
      return { diagnostics: [], interrupted: true };
    }
    return { diagnostics: [], error: CSHARP_RUNTIME_UNAVAILABLE_MESSAGE };
  }
  if (
    options.signal?.aborted ||
    handle.retired ||
    workers.get(options.namespace) !== handle ||
    !handle.ready
  ) {
    if (options.signal?.aborted) retireHandle(handle, interruptedResult());
    return options.signal?.aborted
      ? { diagnostics: [], interrupted: true }
      : { diagnostics: [], error: "C# compile generation is unavailable." };
  }
  if (handle.pending.size > 0) {
    return { diagnostics: [], error: "C# compile generation is busy." };
  }
  const id = `csharp-compile:${handle.generation}:${++handle.sequence}`;
  return new Promise<CSharpCompileResult>((resolve) => {
    const pending: PendingCSharpCompile = {
      kind: "compile",
      id,
      resolve,
      timer: 0 as unknown as ReturnType<typeof setTimeout>,
      signal: options.signal,
    };
    pending.timer = setTimeout(() => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(
        handle,
        { logs: [], error: "C# compile timed out and was stopped.", timedOut: true },
        false,
        { diagnostics: [], error: "C# compile timed out and was stopped." },
      );
    }, timeoutMs);
    pending.abortListener = () => {
      if (handle.pending.get(id) !== pending) return;
      retireHandle(
        handle,
        interruptedResult(),
        false,
        { diagnostics: [], interrupted: true },
      );
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
      handle.worker.postMessage({
        protocol: CSHARP_EXECUTION_PROTOCOL,
        type: "compile",
        id,
        source,
      });
    } catch {
      retireHandle(
        handle,
        { logs: [], error: "C# compile transport failed." },
        false,
        { diagnostics: [], error: "C# compile transport failed." },
      );
    }
  });
}

export function interruptCSharp(namespace: string): boolean {
  const handle = workers.get(namespace);
  if (!handle || handle.retired) return false;
  retireHandle(handle, interruptedResult());
  return true;
}

export function resetCSharpRuntime(namespace: string): void {
  const handle = workers.get(namespace);
  if (!handle) return;
  retireHandle(handle, {
    logs: [],
    interrupted: true,
    error: "C# runtime was restarted.",
  });
}

export function __setCSharpWorkerFactoryForTests(
  factory: (() => CSharpWorkerLike) | null,
): void {
  for (const handle of [...workers.values()]) {
    retireHandle(handle, interruptedResult());
  }
  availability.clear();
  runtimeVerified = false;
  verifyingHandle = null;
  workerFactory = factory ?? defaultWorkerFactory;
}

export function __resetCSharpRunnerForTests(): void {
  __setCSharpWorkerFactoryForTests(null);
}
