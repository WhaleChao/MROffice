/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client-side Python execution via Pyodide (CPython compiled to WebAssembly).
 *
 * In a real browser, Python runs in a Web Worker (`/python-worker.js`) so a
 * runaway snippet (e.g. `while True: pass`) can be stopped by terminating the
 * worker instead of freezing the tab — the worker thread is isolated from the
 * UI thread. Each worker keeps a live Pyodide interpreter, so cells in one
 * notebook share kernel state (keyed by namespace); separate notebooks get
 * separate workers. WHERE that worker lives depends on configuration:
 *
 *  - NEXT_PUBLIC_SANDBOX_ORIGIN set (independent-origin mode): the workers are
 *    hosted by /python-sandbox.html on that SEPARATE origin, embedded here in a
 *    hidden cross-origin iframe and reached over postMessage. Pyodide cannot
 *    run in the JS sandbox's opaque-origin iframe (it needs IndexedDB + nested
 *    workers, which an opaque origin blocks), so isolation comes from the
 *    same-origin policy instead: the sandbox origin shares no cookies/storage
 *    with the app, and the app CSP drops the Pyodide CDN connect-src grant
 *    ('wasm-unsafe-eval' stays — the in-origin sql.js runner needs it; see
 *    proxy.ts).
 *
 *  - Otherwise (local development/tests only): a same-origin worker, served
 *    with its own permissive CSP (see proxy.ts). Production fails closed when
 *    the independent origin is absent or resolves to the app origin; user
 *    Python must never inherit the authenticated application origin there.
 *
 * In tests / SSR (no Worker, or an injected loader) execution falls back to an
 * in-process Pyodide instance so the orchestration stays unit-testable without
 * the real WASM runtime.
 */

import { desktopOndemand, ensureDesktopOndemand } from "../desktop-ondemand";
import type { LogStream } from "./cell-output-format";
import type { RunResult } from "./js-runner";
import type { OnOutput } from "./output-coalescer";
import {
  boundExecutionResult,
  createExecutionLogCapture,
  createExecutionStreamLimiter,
  type ExecutionStreamLimiter,
} from "./execution-output-budget";
import { type VariableInfo } from "@/lib/code-runner/inspector";
import {
  desktopPythonSandboxOrigin,
  IS_DESKTOP_CLASSIC,
} from "../desktop-classic";

const PYODIDE_VERSION = "0.26.2";
const PYODIDE_CDN = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;
// Subresource Integrity for the pinned bootstrap script (sha384 of
// pyodide.js@0.26.2 on jsDelivr; the version is immutable so the hash is
// stable). Used by the in-process loader; the worker uses importScripts (no SRI
// hook) under a CDN-restricted CSP.
const PYODIDE_SCRIPT_SRI =
  "sha384-tVslJOEkg7nVRW3Y3/ReGX0NnonNrbcmt1R5qFbQXQdGa2chRkoJYHAjAsv3zoTq";

/** Same-origin worker that hosts Pyodide (served with its own CSP). */
const PYTHON_WORKER_URL = "/python-worker.js";
/** Relay page that hosts the workers on the independent sandbox origin. */
const PYTHON_SANDBOX_PAGE_PATH = "/python-sandbox.html";
const DEFAULT_NAMESPACE = "__default__";
/** Stop a snippet that runs longer than this (anti-freeze). */
export const DEFAULT_PYTHON_TIMEOUT_MS = 20_000;
/** Give the cross-origin sandbox iframe this long to signal readiness. */
const FRAME_READY_TIMEOUT_MS = 15_000;
const PRIVATE_SANDBOX_TICKET_PATH = "/__beta/python-ticket";
const PRIVATE_SANDBOX_TICKET_MAX_AGE_MS = 60_000;
const PRIVATE_SANDBOX_RESPONSE_MAX_BYTES = 8_192;

function usesPrivateSandboxBootstrap(): boolean {
  return process.env.NEXT_PUBLIC_PRIVATE_BETA_SANDBOX === "true";
}

async function readPrivateSandboxResponse(response: Response): Promise<string> {
  const reader = response.body?.getReader();
  if (!reader) throw new Error("private sandbox response body missing");
  const decoder = new TextDecoder("utf-8", { fatal: true });
  let size = 0;
  let body = "";
  try {
    while (true) {
      const chunk = await reader.read();
      if (chunk.done) return body + decoder.decode();
      size += chunk.value.byteLength;
      if (size > PRIVATE_SANDBOX_RESPONSE_MAX_BYTES) {
        throw new Error("private sandbox authorization response too large");
      }
      body += decoder.decode(chunk.value, { stream: true });
    }
  } finally {
    await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}

export interface PyodideLike {
  runPythonAsync(code: string): Promise<unknown>;
  setStdout(options: { batched: (text: string) => void }): void;
  setStderr(options: { batched: (text: string) => void }): void;
  /**
   * Pre-load any Pyodide-provided packages a snippet imports (numpy, pandas,
   * …) so `import numpy` works without an explicit install step. Optional so
   * the orchestration stays testable without the real runtime.
   */
  loadPackagesFromImports?(code: string): Promise<unknown>;
}

export type PyodideLoader = () => Promise<PyodideLike>;

let cached: Promise<PyodideLike> | null = null;

/** Reset the cached runtime (tests). */
export function __resetPyodideCacheForTests(): void {
  cached = null;
}

function loadScriptOnce(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (typeof document === "undefined") {
      reject(new Error("No document"));
      return;
    }
    const existing = document.querySelector<HTMLScriptElement>(
      `script[data-pyodide="1"]`,
    );
    if (existing) {
      if (
        typeof (globalThis as { loadPyodide?: unknown }).loadPyodide ===
        "function"
      ) {
        resolve();
        return;
      }
      if (existing.dataset.pyodideState === "loading") {
        const settle = (error?: Error) => {
          existing.removeEventListener("load", onLoad);
          existing.removeEventListener("error", onError);
          if (error) reject(error);
          else resolve();
        };
        const onLoad = () => settle();
        const onError = () => {
          existing.remove();
          settle(new Error("Failed to load Pyodide"));
        };
        existing.addEventListener("load", onLoad, { once: true });
        existing.addEventListener("error", onError, { once: true });
        return;
      }
      // A completed/failed tag without the loader is not reusable. Retire it
      // so this request performs a real network retry instead of false success.
      existing.remove();
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.crossOrigin = "anonymous";
    if (src.endsWith("pyodide.js")) script.integrity = PYODIDE_SCRIPT_SRI;
    script.dataset.pyodide = "1";
    script.dataset.pyodideState = "loading";
    script.onload = () => {
      script.dataset.pyodideState = "loaded";
      resolve();
    };
    script.onerror = () => {
      script.remove();
      reject(new Error("Failed to load Pyodide"));
    };
    document.head.appendChild(script);
  });
}

async function defaultLoadPyodide(): Promise<PyodideLike> {
  await loadScriptOnce(`${PYODIDE_CDN}pyodide.js`);
  const loader = (
    globalThis as unknown as {
      loadPyodide?: (opts: { indexURL: string }) => Promise<PyodideLike>;
    }
  ).loadPyodide;
  if (!loader) throw new Error("Pyodide loader unavailable");
  return loader({ indexURL: PYODIDE_CDN });
}

function getPyodide(loader: PyodideLoader): Promise<PyodideLike> {
  if (!cached) {
    cached = loader().catch((err) => {
      cached = null; // allow retry after a failed load
      throw err;
    });
  }
  return cached;
}

// ---------------------------------------------------------------------------
// Worker path (real browsers): isolated, terminable, per-namespace kernels.
// ---------------------------------------------------------------------------

interface PendingRun {
  resolve: (result: PythonDebugRunResult) => void;
  timer: ReturnType<typeof setTimeout>;
  /**
   * Cancellation belongs to this exact request and worker generation. The
   * listener is removed on every settlement path so a later abort cannot retain
   * (or retire) a replacement generation for the same namespace.
   */
  signal?: AbortSignal;
  abortListener?: () => void;
  /**
   * Incremental output sink for THIS run: the worker posts `{id, type:"stream",
   * stream, chunk}` messages as the kernel produces stdout/stderr, and they are
   * routed here so the cell renders progressively. Absent ⇒ no streaming (the
   * caller didn't ask for it); the final batched result is unaffected.
   */
  onOutput?: OnOutput;
  /** Independent host-side guard even when a worker/relay violates its contract. */
  streamLimiter: ExecutionStreamLimiter;
  /**
   * The namespace whose page shows this run as a game, once its first frame
   * came (python-worker.js runGame): from then on it runs until stopped.
   */
  gameNamespace?: string;
}



/**
 * One frame of a game: its window's pixels and its caption. The page's own
 * Python sends RGBA pixels; a desktop program in a cloud window
 * (cloud-gui.ts) sends a decoded picture of its windows.
 */
export interface PythonGameFrame {
  width: number;
  height: number;
  caption: string;
  pixels?: Uint8ClampedArray<ArrayBuffer>;
  bitmap?: ImageBitmap;
  /** It plays in a cloud window, not in the page (its hint says so). */
  cloud?: boolean;
}

export type PythonGameSignal =
  | { kind: "frame"; frame: PythonGameFrame }
  | { kind: "ended" };

/**
 * A game played somewhere other than the page's own Python worker (a cloud
 * window, cloud-gui.ts): the page's keys, mouse and Stop go to it.
 */
export interface PythonGameController {
  input(events: readonly PythonGameInput[]): boolean;
  stop(): void;
}

/** The page's keys and mouse, as python-worker.js acceptGameInput reads them. */
export interface PythonGameInput {
  t: "kd" | "ku" | "mm" | "md" | "mu" | "wheel" | "blur" | "quit";
  key?: string;
  code?: string;
  x?: number;
  y?: number;
  dy?: number;
  button?: number;
  shift?: boolean;
  ctrl?: boolean;
  alt?: boolean;
  meta?: boolean;
  repeat?: boolean;
}

const GAME_MAX_SIDE = 2048;
/** How long a stopped game has to end itself before its kernel is stopped. */
const GAME_STOP_GRACE_MS = 1_500;
const gameListeners = new Map<string, Set<(signal: PythonGameSignal) => void>>();
const gameRuns = new Map<string, { handle: WorkerHandle; id: number }>();
const gameControllers = new Map<string, PythonGameController>();

/** Route a namespace's game input and Stop to ``controller`` until the returned
 * function is called. */
export function registerPythonGameController(
  namespace: string,
  controller: PythonGameController,
): () => void {
  gameControllers.set(namespace, controller);
  return () => {
    if (gameControllers.get(namespace) === controller) gameControllers.delete(namespace);
  };
}

/** Show a frame, or the end, of a game played outside the page's worker. */
export function emitPythonGame(namespace: string, signal: PythonGameSignal): void {
  emitGame(namespace, signal);
}

function emitGame(namespace: string, signal: PythonGameSignal): void {
  for (const listener of Array.from(gameListeners.get(namespace) ?? [])) {
    try {
      listener(signal);
    } catch {
      /* a listener's failure is its own */
    }
  }
}

function boundedGameFrame(data: Record<string, unknown>): PythonGameFrame | null {
  const { width, height, pixels } = data;
  if (
    typeof width !== "number" ||
    typeof height !== "number" ||
    !Number.isSafeInteger(width) ||
    !Number.isSafeInteger(height) ||
    width < 1 ||
    height < 1 ||
    width > GAME_MAX_SIDE ||
    height > GAME_MAX_SIDE ||
    !(pixels instanceof ArrayBuffer) ||
    pixels.byteLength !== width * height * 4
  ) {
    return null;
  }
  return {
    width,
    height,
    caption: typeof data.caption === "string" ? data.caption.slice(0, 200) : "",
    pixels: new Uint8ClampedArray(pixels),
  };
}

/** Follow the game a namespace's run plays: its frames, then its end. */
export function subscribePythonGame(
  namespace: string,
  listener: (signal: PythonGameSignal) => void,
): () => void {
  let listeners = gameListeners.get(namespace);
  if (!listeners) {
    listeners = new Set();
    gameListeners.set(namespace, listeners);
  }
  listeners.add(listener);
  return () => {
    const current = gameListeners.get(namespace);
    current?.delete(listener);
    if (current && current.size === 0) gameListeners.delete(namespace);
  };
}

/** Hand the running game the page's keys and mouse. False when none runs. */
export function sendPythonGameInput(
  namespace: string,
  events: readonly PythonGameInput[],
): boolean {
  const controller = gameControllers.get(namespace);
  if (controller) return controller.input(events.slice(0, 200));
  const run = gameRuns.get(namespace);
  if (!run || !run.handle.active || !run.handle.pending.has(run.id)) return false;
  try {
    run.handle.transport.postMessage({ id: run.id, op: "game-input", events: events.slice(0, 200) });
    return true;
  } catch {
    return false;
  }
}

/**
 * The run's own Stop on a page whose Python plays a game: stopped as the
 * game's Stop button stops it (QUIT, then the kernel), so the run ends as a
 * game does, not as an interrupted program. False when no game plays there.
 */
export function stopPlayingPythonGame(namespace: string): boolean {
  if (!gameControllers.has(namespace) && !gameRuns.has(namespace)) return false;
  stopPythonGame(namespace);
  return true;
}

/** Close the game's window (a QUIT it can answer), then stop it if it does not. */
export function stopPythonGame(namespace: string): void {
  const controller = gameControllers.get(namespace);
  if (controller) {
    controller.stop();
    return;
  }
  const run = gameRuns.get(namespace);
  if (!run) return;
  sendPythonGameInput(namespace, [{ t: "quit" }]);
  setTimeout(() => {
    if (gameRuns.get(namespace) === run && run.handle.pending.has(run.id)) {
      interruptPython(namespace);
    }
  }, GAME_STOP_GRACE_MS);
}

interface PendingInspect {
  resolve: (variables: VariableInfo[]) => void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * The first SyntaxError from a parse-only `compile()` of a snippet, or
 * `{ ok: true }`. Mirrors the worker's `lint` op response payload.
 */
export interface PythonLintResult {
  ok: boolean;
  lineno?: number | null;
  offset?: number | null;
  endLineno?: number | null;
  endOffset?: number | null;
  msg?: string | null;
}

/** One pyflakes semantic warning (undefined name, unused import, etc.). */
export interface PythonFlakeWarning {
  /** 1-based line number of the warning. */
  row: number;
  /** 1-based column number of the warning. */
  col: number;
  /** Human-readable warning message (e.g. "imported but unused"). */
  message: string;
  /** pyflakes message class name (e.g. "UnusedImport", "UndefinedName"). */
  type: string;
}

/**
 * Combined lint + semantic warning result from the worker's `lint` op.
 * `lint` holds the SyntaxError (if any); `warnings` holds the pyflakes
 * semantic warnings (populated only when `lint.ok === true`).
 */
export interface PythonLintAndWarnings {
  lint: PythonLintResult;
  warnings: PythonFlakeWarning[];
}

interface PendingLint {
  resolve: (result: PythonLintAndWarnings) => void;
  timer: ReturnType<typeof setTimeout>;
}

export interface PythonFormatReceipt {
  ok: boolean;
  formatted: string;
}

// The sandbox relay already preserves bounded `error` strings for format
// replies. Reuse that authenticated field as a success receipt so this runner
// stays compatible with the independent-origin relay without widening its
// protocol whitelist. Missing/different values (including old workers) fail closed.
const PYTHON_FORMAT_SUCCESS_RECEIPT = "mrnote.format.ok.v1";

interface PendingFormat {
  resolve: (receipt: PythonFormatReceipt) => void;
  timer: ReturnType<typeof setTimeout>;
  /** Original source; oversized or malformed formatter replies fail closed to it. */
  fallback: string;
  signal?: AbortSignal;
  abortListener?: () => void;
}

/** One jedi completion suggestion (kind is jedi's type string, e.g. "function"). */
export interface PythonCompletionItem {
  name: string;
  kind: string | null;
}

interface PendingComplete {
  resolve: (items: PythonCompletionItem[]) => void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * jedi hover info for the symbol under the cursor: `signature` is the callable
 * signature (or jedi's description for non-callables), `doc` the raw docstring
 * (worker-capped). `kind` is jedi's type string ("function", "module", …).
 */
export interface PythonHoverInfo {
  name: string;
  kind: string | null;
  signature: string | null;
  doc: string | null;
}

interface PendingHover {
  resolve: (info: PythonHoverInfo | null) => void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * jedi parameter hints for the enclosing call: `label` is the full signature
 * (jedi's `Signature.to_string()`, e.g. `foo(a, b, *, c=1)`), `parameters` the
 * per-parameter names, and `activeParameter` the 0-based argument slot the caret
 * sits in. Mirrors the TS `SignatureHelpInfo` shape (minus its HTML escaping —
 * jedi output is plain text and is escaped at the render site).
 */
export interface PythonSignatureInfo {
  label: string;
  parameters: string[];
  activeParameter: number;
}

interface PendingSignature {
  resolve: (info: PythonSignatureInfo | null) => void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * jedi go-to-definition result: the definition's `line` (1-based) / `column`
 * (0-based) and whether it lives `inBuffer` (this single buffer — navigable) or
 * out of file (stdlib / another module — the editor shows a "not found" notice).
 */
export interface PythonDefinitionInfo {
  line: number;
  column: number;
  inBuffer: boolean;
}

interface PendingDefinition {
  resolve: (info: PythonDefinitionInfo | null) => void;
  timer: ReturnType<typeof setTimeout>;
}

/** One in-buffer reference: `line` (1-based), `column` (0-based), identifier `length`. */
export interface PythonReference {
  line: number;
  column: number;
  length: number;
}

interface PendingReferences {
  resolve: (refs: PythonReference[]) => void;
  timer: ReturnType<typeof setTimeout>;
}

/**
 * A normalized hit-count predicate (VSCode hit-count breakpoints). Parsed on the
 * TS side from a user string (">= 3", "== 5", "% 2 == 0", "5"); the tracer just
 * compares the running conditional-reach count against it (no Python parsing).
 *  - comparison ops compare `count <op> value`;
 *  - `%` matches when `count % value == rem` (every Nth hit, optionally offset).
 */
export interface PythonHitCondition {
  op: ">=" | ">" | "<=" | "<" | "==" | "%";
  value: number;
  /** Only for `%`: the required remainder (default 0). */
  rem?: number;
}

/** A line breakpoint for a debug-run; `condition` is a Python expression. */
export interface PythonDebugBreakpoint {
  line: number;
  condition?: string;
  /** Optional hit-count predicate; only stop when the count matches. */
  hit?: PythonHitCondition;
}

/** One user-source frame of the stack recorded at a debug event. */
export interface PythonDebugStackFrame {
  name: string;
  line: number;
}

/**
 * One recorded debug event: a breakpoint hit, or a line-step inside the
 * bounded window that follows a hit. Locals are preview strings (inspector
 * capping discipline); the stack covers user-source frames innermost-first.
 */
export interface PythonDebugEvent {
  kind: "breakpoint" | "step" | "exception";
  line: number;
  locals: Record<string, string>;
  stack: PythonDebugStackFrame[];
  /**
   * TRUE (uncapped) user-frame depth at this event. The `stack` slice is capped
   * at 20 frames for display, so its length saturates in deep recursion; this
   * field is the real depth used for step-over/step-out targeting. Absent on
   * older recordings (callers fall back to `stack.length`).
   */
  depth?: number;
  /** Set when the breakpoint condition raised (the hit is still recorded). */
  conditionError?: string;
  /** Bounded display text for a terminal uncaught-exception pause. */
  exception?: string;
}

/** A normal run result plus the debug recording (debug-run op only). */
export interface PythonDebugRunResult extends RunResult {
  readonly debugEvents?: PythonDebugEvent[];
  /** The recording hit its event cap; execution still ran to completion. */
  readonly debugTruncated?: boolean;
  /**
   * The first breakpoint condition that RAISED during recording (preview
   * string). A raising condition is treated as not-hit (VSCode skips it); this
   * one-time note surfaces the failure without aborting the run.
   */
  readonly debugConditionNote?: string;
}

// ---------------------------------------------------------------------------
// Transport layer: where the Pyodide worker lives.
//
// Local development/test mode: a same-origin `new Worker(PYTHON_WORKER_URL)`.
// Independent-origin mode (NEXT_PUBLIC_SANDBOX_ORIGIN): the workers live in a
// hidden cross-origin iframe (/python-sandbox.html on the sandbox origin) and
// are reached over postMessage; one shared iframe hosts every namespace, so
// relayed messages carry a `namespace` field for routing.
// ---------------------------------------------------------------------------

/** The slice of Worker the kernel orchestration needs (local or remote). */
interface PythonTransport {
  postMessage(message: Record<string, unknown>): void;
  terminate(): void;
}

const WORKER_PROTOCOL_MARKER = "mrnote.execution-worker.v1";
const WORKER_PROTOCOL_TOKEN_PATTERN = /^[0-9a-f]{32}$/;

let sandboxOriginOverride: string | null | undefined;

/** Force the sandbox origin (tests). Pass undefined to clear the override. */
export function __setPythonSandboxOriginForTests(
  value: string | null | undefined,
): void {
  sandboxOriginOverride = value;
  // The frame channel is bound to the origin it was created with — drop it so
  // the next run reflects the override instead of a stale origin.
  frameChannel?.dispose?.();
  frameChannel = null;
}

/**
 * The configured independent sandbox origin, or null when unset/invalid.
 * A value equal to the app's own origin is treated as unset: it would provide
 * no isolation. Local development/tests may use the plain worker; production
 * execution separately fails closed in {@link pythonExecutionConfigurationError}.
 */
function pythonSandboxOrigin(): string | null {
  // MRNote Classic: the app's own server on another loopback host, at the
  // port of this launch (lib/desktop-classic.ts).
  const raw =
    sandboxOriginOverride !== undefined
      ? sandboxOriginOverride
      : IS_DESKTOP_CLASSIC
        ? typeof window !== "undefined"
          ? desktopPythonSandboxOrigin(window.location.host)
          : null
        : process.env.NEXT_PUBLIC_SANDBOX_ORIGIN;
  if (!raw) return null;
  try {
    const url = new URL(raw);
    if (url.protocol !== "http:" && url.protocol !== "https:") return null;
    if (
      usesPrivateSandboxBootstrap() &&
      (url.protocol !== "https:" || raw !== url.origin)
    ) {
      return null;
    }
    if (
      typeof window !== "undefined" &&
      url.origin === window.location.origin
    ) {
      return null;
    }
    
    // itself is served from a loopback host. Through a tunnel/remote origin
    // the iframe would point at the VIEWER's machine and hang for 15s —
    // treat the config as unset so the failure is an instant, actionable
    // error instead. Real deployments configure a non-loopback origin.
    if (typeof window !== "undefined") {
      const loopback = new Set(["localhost", "127.0.0.1", "::1", "[::1]"]);
      if (
        loopback.has(url.hostname) &&
        !loopback.has(window.location.hostname)
      ) {
        return null;
      }
      if (
        url.protocol === "http:" &&
        (window.location.protocol === "https:" ||
          (process.env.NODE_ENV === "production" &&
            !loopback.has(url.hostname)))
      ) {
        return null;
      }
    } else if (
      url.protocol === "http:" &&
      process.env.NODE_ENV === "production"
    ) {
      return null;
    }
    return url.origin;
  } catch {
    return null;
  }
}

/**
 * Production Python must never execute in an app-origin worker. Pyodide exposes
 * worker browser APIs through its JS bridge, including origin-scoped storage
 * and messaging capabilities. The worker CSP currently blocks app-origin
 * network requests, but CSP is defense in depth rather than an origin-isolation
 * boundary. Tests and local development retain the injectable/local path; a
 * production build without a genuinely cross-origin host fails closed with an
 * actionable configuration error.
 */
function pythonExecutionConfigurationError(): string | null {
  if (process.env.NODE_ENV !== "production" && !usesPrivateSandboxBootstrap()) {
    return null;
  }
  if (pythonSandboxOrigin() !== null) return null;
  return (
    "Python execution requires an independent Python sandbox origin in " +
    "production. Configure NEXT_PUBLIC_SANDBOX_ORIGIN with a separate HTTPS " +
    "origin and NEXT_PUBLIC_SANDBOX_EMBEDDER_ORIGIN with the app origin."
  );
}

/** Static deployment readiness used by the execution-provider capability UI. */
export function getPythonExecutionAvailability():
  | "available"
  | "unavailable" {
  return pythonExecutionConfigurationError() ? "unavailable" : "available";
}

type FrameSubscriber = (data: Record<string, unknown>) => void;

interface SandboxFrameChannel {
  post(message: Record<string, unknown>): void;
  subscribe(namespace: string, handler: FrameSubscriber): () => void;
  /** Release the real iframe/listener channel; optional for injected fakes. */
  dispose?(): void;
}

let frameChannel: SandboxFrameChannel | null = null;
let frameChannelFactoryOverride:
  ((origin: string) => SandboxFrameChannel) | null = null;

/** Inject a fake frame channel (tests). Pass null to restore the real one. */
export function __setPythonSandboxChannelFactoryForTests(
  factory: ((origin: string) => SandboxFrameChannel) | null,
): void {
  frameChannel?.dispose?.();
  frameChannelFactoryOverride = factory;
  frameChannel = null;
}

function getSandboxFrameChannel(origin: string): SandboxFrameChannel {
  if (frameChannel) return frameChannel;
  frameChannel = (frameChannelFactoryOverride ?? createSandboxFrameChannel)(
    origin,
  );
  return frameChannel;
}

function createSandboxFrameChannel(origin: string): SandboxFrameChannel {
  const subscribers = new Map<string, Set<FrameSubscriber>>();
  let iframe: HTMLIFrameElement | null = null;
  // The relay page registers its message listener only once it has loaded, so
  // it signals readiness with { ready: true }; posts are buffered until then.
  let ready = false;
  let failed = false;
  let disposed = false;
  let pending: Record<string, unknown>[] = [];
  let readyTimer: ReturnType<typeof setTimeout> | null = null;
  let bootstrapAbort: AbortController | null = null;
  let bootstrapForm: HTMLFormElement | null = null;

  function dispose(): void {
    if (disposed) return;
    disposed = true;
    pending = [];
    subscribers.clear();
    bootstrapAbort?.abort();
    bootstrapAbort = null;
    if (bootstrapForm) {
      bootstrapForm.replaceChildren();
      bootstrapForm.remove();
      bootstrapForm = null;
    }
    window.removeEventListener("message", onWindowMessage);
    if (readyTimer) {
      clearTimeout(readyTimer);
      readyTimer = null;
    }
    if (iframe?.parentNode) iframe.parentNode.removeChild(iframe);
    iframe = null;
    if (frameChannel === channel) frameChannel = null;
  }

  // An unreachable/misconfigured sandbox origin would otherwise surface as a
  // generic run timeout. Fail every live kernel with a pointed message and
  // drop the channel so a later run retries with a fresh iframe.
  function fail(message: string): void {
    if (failed || disposed) return;
    failed = true;
    for (const handlers of Array.from(subscribers.values())) {
      for (const handler of Array.from(handlers)) {
        handler({ workerError: message });
      }
    }
    dispose();
  }

  function flush(): void {
    if (disposed || !ready || !iframe?.contentWindow) return;
    const queued = pending;
    pending = [];
    for (const message of queued) {
      iframe.contentWindow.postMessage(message, origin);
    }
  }

  // Only the sandbox iframe is trusted: match both the source window and the
  // exact sandbox origin (the frame keeps its REAL origin — it is not opaque).
  const onWindowMessage = (event: MessageEvent) => {
    if (!iframe || event.source !== iframe.contentWindow) return;
    if (event.origin !== origin) return;
    const data = event.data as Record<string, unknown> | undefined;
    if (!data || typeof data !== "object") return;
    if (data.ready === true) {
      ready = true;
      if (readyTimer) {
        clearTimeout(readyTimer);
        readyTimer = null;
      }
      flush();
      return;
    }
    const namespace =
      typeof data.namespace === "string" ? data.namespace : DEFAULT_NAMESPACE;
    const handlers = subscribers.get(namespace);
    if (!handlers) return;
    for (const handler of Array.from(handlers)) handler(data);
  };

  function makeFrame(): HTMLIFrameElement {
    const frame = document.createElement("iframe");
    // allow-same-origin keeps ONLY the separate host's origin. The relay
    // cannot submit forms, open windows, or navigate the top-level app.
    frame.setAttribute("sandbox", "allow-scripts allow-same-origin");
    frame.setAttribute("aria-hidden", "true");
    frame.style.display = "none";
    return frame;
  }

  async function bootstrapPrivateFrame(): Promise<void> {
    const abort = new AbortController();
    bootstrapAbort = abort;
    let ticket = "";
    let ticketInput: HTMLInputElement | null = null;
    try {
      const response = await fetch(PRIVATE_SANDBOX_TICKET_PATH, {
        method: "POST",
        credentials: "same-origin",
        cache: "no-store",
        redirect: "error",
        headers: {
          "Content-Type": "application/json",
          "X-MRNote-Python-Bootstrap": "1",
        },
        body: "{}",
        signal: abort.signal,
      });
      if (disposed) return;
      if (
        !response.ok ||
        response.redirected ||
        response.headers.get("content-type")?.split(";")[0].trim() !==
          "application/json"
      ) {
        throw new Error("private sandbox authorization rejected");
      }
      let body = await readPrivateSandboxResponse(response);
      if (disposed) return;
      let payload: unknown = JSON.parse(body);
      body = "";
      if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
        throw new Error("invalid private sandbox authorization");
      }
      const result = payload as Record<string, unknown>;
      const now = Date.now();
      if (
        typeof result.ticket !== "string" ||
        !/^[A-Za-z0-9._~-]{32,4096}$/.test(result.ticket) ||
        typeof result.expiresAt !== "number" ||
        !Number.isSafeInteger(result.expiresAt) ||
        result.expiresAt <= now ||
        result.expiresAt - now > PRIVATE_SANDBOX_TICKET_MAX_AGE_MS ||
        result.sandboxOrigin !== origin
      ) {
        throw new Error("invalid private sandbox authorization");
      }
      ticket = result.ticket;
      result.ticket = "";
      payload = null;
      if (disposed || abort.signal.aborted) return;

      const frame = makeFrame();
      frame.name = `mrnote-python-${crypto.randomUUID()}`;
      iframe = frame;
      document.body.appendChild(frame);

      // A native POST navigation carries the one-use authorization in the
      // body, never in a URL/referrer or a worker message. The
      // parent submits it; the child gains no allow-forms permission.
      const form = document.createElement("form");
      bootstrapForm = form;
      form.method = "POST";
      form.action = `${origin}${PYTHON_SANDBOX_PAGE_PATH}`;
      form.target = frame.name;
      form.style.display = "none";
      ticketInput = document.createElement("input");
      ticketInput.type = "hidden";
      ticketInput.name = "ticket";
      ticketInput.value = ticket;
      ticket = "";
      form.appendChild(ticketInput);
      document.body.appendChild(form);
      // No awaits between this check and submit: disposal must never turn a
      // removed named iframe into an orphan popup target.
      if (disposed || abort.signal.aborted || !frame.isConnected) return;
      HTMLFormElement.prototype.submit.call(form);
    } catch {
      if (!disposed) {
        fail(
          "Python sandbox authorization failed. Unlock the private site and retry.",
        );
      }
    } finally {
      ticket = "";
      if (ticketInput) ticketInput.value = "";
      bootstrapForm?.replaceChildren();
      bootstrapForm?.remove();
      bootstrapForm = null;
      if (bootstrapAbort === abort) bootstrapAbort = null;
    }
  }

  window.addEventListener("message", onWindowMessage);
  function startFrame(): void {
    if (disposed || failed) return;
    if (!usesPrivateSandboxBootstrap()) {
      const frame = makeFrame();
      frame.src = `${origin}${PYTHON_SANDBOX_PAGE_PATH}`;
      document.body.appendChild(frame);
      iframe = frame;
    }
    readyTimer = setTimeout(() => {
      fail(
        `Python sandbox origin ${origin} did not become ready within ` +
          `${Math.round(FRAME_READY_TIMEOUT_MS / 1000)}s — check that the ` +
          "sandbox host serves /python-sandbox.html (NEXT_PUBLIC_SANDBOX_ORIGIN).",
      );
    }, FRAME_READY_TIMEOUT_MS);
    if (usesPrivateSandboxBootstrap()) void bootstrapPrivateFrame();
  }
  // The desktop app fetches Python after installing (lib/desktop-ondemand):
  // the frame starts once it is on this computer; posts wait meanwhile.
  if (desktopOndemand()) {
    ensureDesktopOndemand("pyodide").then(startFrame, (error: unknown) =>
      fail(
        `Python is not on this computer yet and could not be downloaded (${
          error instanceof Error ? error.message : String(error)
        }). Connect to the internet and run it again.`,
      ),
    );
  } else {
    startFrame();
  }

  const channel: SandboxFrameChannel = {
    post(message) {
      if (failed || disposed) return;
      if (ready && iframe?.contentWindow) {
        iframe.contentWindow.postMessage(message, origin);
      } else {
        pending.push(message);
      }
    },
    subscribe(namespace, handler) {
      if (disposed) return () => {};
      let handlers = subscribers.get(namespace);
      if (!handlers) {
        handlers = new Set();
        subscribers.set(namespace, handlers);
      }
      handlers.add(handler);
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        handlers.delete(handler);
        if (handlers.size === 0) subscribers.delete(namespace);
        // Each subscription represents one live remote Python transport. Keep
        // the expensive iframe while any page kernel uses it; the final release
        // destroys the child browsing context, ready watchdog and global
        // message listener instead of pinning them for the app's lifetime.
        if (subscribers.size === 0) dispose();
      };
    },
    dispose,
  };
  return channel;
}

function createLocalTransport(
  onMessage: (data: unknown) => void,
  onError: (message: string) => void,
): PythonTransport {
  const worker = new Worker(PYTHON_WORKER_URL);
  let protocolToken: string | null = null;
  worker.onmessage = (event: MessageEvent) => {
    const envelope = event.data;
    if (
      Array.isArray(envelope) &&
      envelope.length === 3 &&
      envelope[0] === WORKER_PROTOCOL_MARKER &&
      typeof envelope[1] === "string" &&
      WORKER_PROTOCOL_TOKEN_PATTERN.test(envelope[1]) &&
      envelope[2] !== null &&
      typeof envelope[2] === "object"
    ) {
      const token = envelope[1];
      const payload = envelope[2] as Record<string, unknown>;
      if (protocolToken === null) {
        if (payload.type !== "protocol-ready") return;
        protocolToken = token;
        return;
      }
      if (token !== protocolToken || payload.type === "protocol-ready") return;
      onMessage(payload);
      return;
    }
    // Compatibility for injected test workers and an old cached development
    // worker. The real worker's ready envelope is emitted at script startup,
    // before authored Python can run; once seen, raw messages fail closed.
    if (protocolToken === null) onMessage(envelope);
  };
  worker.onerror = (event: ErrorEvent) =>
    onError(event.message || "Python worker crashed");
  return {
    postMessage: (message) => worker.postMessage(message),
    terminate: () => worker.terminate(),
  };
}

let transportGeneration = 0;

function createRemoteTransport(
  namespace: string,
  origin: string,
  onMessage: (data: unknown) => void,
  onError: (message: string) => void,
): PythonTransport {
  const channel = getSandboxFrameChannel(origin);
  // Wire key: unique per transport GENERATION, not just per namespace. After
  // a kernel kill (timeout/interrupt/restart) a reply from the dying worker
  // can still be in flight; routed by plain namespace it would settle a run
  // of the namespace's REPLACEMENT kernel (per-kernel ids restart at 1, so
  // they collide). The relay treats the key as opaque, and stale messages
  // carry a wire key nobody subscribes to anymore — dropped, matching the
  // old local path where worker.terminate() cut delivery off.
  transportGeneration += 1;
  const wireNamespace = `${namespace}#${transportGeneration}`;
  const unsubscribe = channel.subscribe(wireNamespace, (data) => {
    if (typeof data.workerError === "string") {
      onError(data.workerError);
      return;
    }
    onMessage(data);
  });
  return {
    postMessage: (message) =>
      channel.post({ ...message, namespace: wireNamespace }),
    terminate: () => {
      // worker.terminate() is not reachable cross-frame; the relay page kills
      // this kernel's worker on this control op. Post BEFORE releasing the
      // subscription: the last release disposes the iframe channel itself.
      channel.post({ op: "terminate", namespace: wireNamespace });
      unsubscribe();
    },
  };
}

interface WorkerHandle {
  transport: PythonTransport;
  /** False once this exact transport generation has been retired. */
  active: boolean;
  seq: number;
  pending: Map<number, PendingRun>;
  inspects: Map<number, PendingInspect>;
  lints: Map<number, PendingLint>;
  formats: Map<number, PendingFormat>;
  completes: Map<number, PendingComplete>;
  hovers: Map<number, PendingHover>;
  signatures: Map<number, PendingSignature>;
  definitions: Map<number, PendingDefinition>;
  references: Map<number, PendingReferences>;
}

function failedPythonFormat(fallback: string): PythonFormatReceipt {
  return { ok: false, formatted: fallback };
}

function settlePendingFormat(
  handle: WorkerHandle,
  id: number,
  pending: PendingFormat,
  receipt: PythonFormatReceipt,
): void {
  if (handle.formats.get(id) !== pending) return;
  clearTimeout(pending.timer);
  handle.formats.delete(id);
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  pending.resolve(receipt);
}

function settlePendingRun(
  handle: WorkerHandle,
  id: number,
  pending: PendingRun,
  result: PythonDebugRunResult,
): void {
  if (handle.pending.get(id) !== pending) return;
  handle.pending.delete(id);
  clearTimeout(pending.timer);
  if (pending.gameNamespace !== undefined) {
    const run = gameRuns.get(pending.gameNamespace);
    if (run && run.handle === handle && run.id === id) gameRuns.delete(pending.gameNamespace);
    emitGame(pending.gameNamespace, { kind: "ended" });
  }
  if (pending.signal && pending.abortListener) {
    pending.signal.removeEventListener("abort", pending.abortListener);
  }
  pending.resolve({ ...result, logs: [...result.logs] });
}

const workers = new Map<string, WorkerHandle>();

function settleWorkerOperations(
  handle: WorkerHandle,
  runResult: PythonDebugRunResult,
): void {
  for (const [id, pending] of Array.from(handle.pending.entries())) {
    settlePendingRun(handle, id, pending, runResult);
  }
  for (const inspect of handle.inspects.values()) {
    clearTimeout(inspect.timer);
    inspect.resolve([]);
  }
  handle.inspects.clear();
  for (const lint of handle.lints.values()) {
    clearTimeout(lint.timer);
    lint.resolve({ lint: { ok: true }, warnings: [] });
  }
  handle.lints.clear();
  for (const [id, fmt] of Array.from(handle.formats.entries())) {
    settlePendingFormat(handle, id, fmt, failedPythonFormat(fmt.fallback));
  }
  for (const complete of handle.completes.values()) {
    clearTimeout(complete.timer);
    complete.resolve([]);
  }
  handle.completes.clear();
  for (const hover of handle.hovers.values()) {
    clearTimeout(hover.timer);
    hover.resolve(null);
  }
  handle.hovers.clear();
  for (const signature of handle.signatures.values()) {
    clearTimeout(signature.timer);
    signature.resolve(null);
  }
  handle.signatures.clear();
  for (const definition of handle.definitions.values()) {
    clearTimeout(definition.timer);
    definition.resolve(null);
  }
  handle.definitions.clear();
  for (const references of handle.references.values()) {
    clearTimeout(references.timer);
    references.resolve([]);
  }
  handle.references.clear();
}

function retireWorkerGeneration(namespace: string, handle: WorkerHandle): void {
  if (!handle.active || workers.get(namespace) !== handle) return;
  // Invalidate routing before touching the transport. A re-entrant next run
  // therefore creates a fresh generation and no late relay message can enter
  // this handle.
  handle.active = false;
  workers.delete(namespace);
  handle.transport.terminate();
}

function terminateWorker(namespace: string): void {
  const handle = workers.get(namespace);
  if (!handle) return;
  // Page close / explicit reset has no caller left to own these watchdogs.
  // Settle every continuation before retiring the transport so no component,
  // timer or large source snapshot remains pinned until its original timeout.
  settleWorkerOperations(handle, {
    logs: [],
    error: "Python runtime reset.",
    interrupted: true,
  });
  retireWorkerGeneration(namespace, handle);
}

/** Retire only the generation that initiated an operation which later timed out. */
function terminateWorkerGeneration(
  namespace: string,
  handle: WorkerHandle,
): void {
  retireWorkerGeneration(namespace, handle);
}

const MAX_PYTHON_PROTOCOL_ITEMS = 500;
const MAX_PYTHON_COMPLETIONS = 80;
const MAX_PYTHON_PROTOCOL_TOTAL_CHARS = 100_000;
const MAX_PYTHON_PROTOCOL_NAME_CHARS = 256;
const MAX_PYTHON_PROTOCOL_TYPE_CHARS = 128;
const MAX_PYTHON_VARIABLE_PREVIEW_CHARS = 120;
const MAX_PYTHON_PROTOCOL_FIELD_CHARS = 4_096;
const MAX_PYTHON_HOVER_DOC_CHARS = 1_500;
const MAX_PYTHON_FORMATTED_CHARS = 2_000_000;
const MAX_PYTHON_PROTOCOL_COORDINATE = 10_000_000;
const MAX_PYTHON_DEBUG_EXCEPTION_CHARS = 500;

interface ProtocolStringBudget {
  remaining: number;
}

function boundedProtocolString(
  value: unknown,
  maxChars: number,
  budget: ProtocolStringBudget,
  allowEmpty = true,
): string | null {
  if (typeof value !== "string" || budget.remaining <= 0) return null;
  const bounded = value.slice(0, Math.min(maxChars, budget.remaining));
  if (!allowEmpty && bounded.length === 0) return null;
  budget.remaining -= bounded.length;
  return bounded;
}

function pythonResultToJs(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  const proxy = value as {
    toJs?: (options?: {
      dict_converter?: (
        entries: Iterable<readonly [PropertyKey, unknown]>,
      ) => object;
      create_pyproxies?: boolean;
    }) => unknown;
    destroy?: () => void;
  };
  if (typeof proxy.toJs !== "function") return value;
  try {
    return proxy.toJs({
      dict_converter: Object.fromEntries,
      create_pyproxies: false,
    });
  } finally {
    try {
      proxy.destroy?.();
    } catch {
      // The PyProxy may already have been released by an injected runtime.
    }
  }
}

function decodePythonResult(value: unknown): unknown {
  const plain = pythonResultToJs(value);
  // Backward-compatible test/injected runtimes may still return a JSON string.
  // Real Pyodide helpers below return native dict/list proxies and never trust the
  // persistent interpreter's mutable json module for protocol serialization.
  if (typeof plain !== "string") return plain;
  try {
    return JSON.parse(plain) as unknown;
  } catch {
    return plain;
  }
}

function boundedVariables(value: unknown): VariableInfo[] {
  if (!Array.isArray(value)) return [];
  const result: VariableInfo[] = [];
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  for (const candidate of value.slice(0, MAX_PYTHON_PROTOCOL_ITEMS)) {
    if (
      !candidate ||
      typeof candidate !== "object" ||
      Array.isArray(candidate)
    ) {
      continue;
    }
    const record = candidate as Record<string, unknown>;
    const name = boundedProtocolString(
      record.name,
      MAX_PYTHON_PROTOCOL_NAME_CHARS,
      budget,
      false,
    );
    if (name === null) continue;
    const type = boundedProtocolString(
      record.type,
      MAX_PYTHON_PROTOCOL_TYPE_CHARS,
      budget,
    );
    const preview = boundedProtocolString(
      record.preview,
      MAX_PYTHON_VARIABLE_PREVIEW_CHARS,
      budget,
    );
    if (type === null || preview === null) break;
    result.push({ name, type, preview });
  }
  return result;
}

function boundedLintResult(value: unknown): PythonLintResult {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return { ok: true };
  }
  const record = value as Record<string, unknown>;
  if (record.ok !== false) return { ok: true };
  const boundedCoordinate = (candidate: unknown): number | null =>
    Number.isSafeInteger(candidate) &&
    Number(candidate) >= 0 &&
    Number(candidate) <= MAX_PYTHON_PROTOCOL_COORDINATE
      ? Number(candidate)
      : null;
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  return {
    ok: false,
    lineno: boundedCoordinate(record.lineno),
    offset: boundedCoordinate(record.offset),
    endLineno: boundedCoordinate(record.endLineno),
    endOffset: boundedCoordinate(record.endOffset),
    msg: boundedProtocolString(
      record.msg,
      MAX_PYTHON_PROTOCOL_FIELD_CHARS,
      budget,
    ),
  };
}

function boundedLintWarnings(value: unknown): PythonFlakeWarning[] {
  if (!Array.isArray(value)) return [];
  const result: PythonFlakeWarning[] = [];
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  for (const candidate of value.slice(0, MAX_PYTHON_PROTOCOL_ITEMS)) {
    if (
      !candidate ||
      typeof candidate !== "object" ||
      Array.isArray(candidate)
    ) {
      continue;
    }
    const record = candidate as Record<string, unknown>;
    if (
      !Number.isSafeInteger(record.row) ||
      Number(record.row) < 1 ||
      Number(record.row) > MAX_PYTHON_PROTOCOL_COORDINATE ||
      !Number.isSafeInteger(record.col) ||
      Number(record.col) < 1 ||
      Number(record.col) > MAX_PYTHON_PROTOCOL_COORDINATE
    ) {
      continue;
    }
    const message = boundedProtocolString(
      record.message,
      MAX_PYTHON_PROTOCOL_FIELD_CHARS,
      budget,
      false,
    );
    const type = boundedProtocolString(
      record.type,
      MAX_PYTHON_PROTOCOL_TYPE_CHARS,
      budget,
      false,
    );
    if (message === null || type === null) break;
    result.push({
      row: Number(record.row),
      col: Number(record.col),
      message,
      type,
    });
  }
  return result;
}

function boundedCompletions(value: unknown): PythonCompletionItem[] {
  if (!Array.isArray(value)) return [];
  const result: PythonCompletionItem[] = [];
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  for (const candidate of value.slice(0, MAX_PYTHON_COMPLETIONS)) {
    if (
      !candidate ||
      typeof candidate !== "object" ||
      Array.isArray(candidate)
    ) {
      continue;
    }
    const record = candidate as Record<string, unknown>;
    const name = boundedProtocolString(
      record.name,
      MAX_PYTHON_PROTOCOL_NAME_CHARS,
      budget,
      false,
    );
    if (name === null) continue;
    const kind =
      record.kind === null || record.kind === undefined
        ? null
        : boundedProtocolString(
            record.kind,
            MAX_PYTHON_PROTOCOL_TYPE_CHARS,
            budget,
          );
    if (record.kind !== null && record.kind !== undefined && kind === null)
      break;
    result.push({ name, kind });
  }
  return result;
}

function boundedHover(value: unknown): PythonHoverInfo | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (record.found !== true) return null;
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  const name = boundedProtocolString(
    record.name,
    MAX_PYTHON_PROTOCOL_NAME_CHARS,
    budget,
    false,
  );
  if (name === null) return null;
  const optional = (candidate: unknown, maxChars: number): string | null =>
    candidate === null || candidate === undefined
      ? null
      : boundedProtocolString(candidate, maxChars, budget);
  const kind = optional(record.kind, MAX_PYTHON_PROTOCOL_TYPE_CHARS);
  const signature = optional(record.signature, MAX_PYTHON_PROTOCOL_FIELD_CHARS);
  const doc = optional(record.doc, MAX_PYTHON_HOVER_DOC_CHARS);
  if (
    (record.kind != null && kind === null) ||
    (record.signature != null && signature === null) ||
    (record.doc != null && doc === null)
  ) {
    return null;
  }
  return { name, kind, signature, doc };
}

function boundedSignature(value: unknown): PythonSignatureInfo | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (record.found !== true || !Array.isArray(record.params)) return null;
  const budget = { remaining: MAX_PYTHON_PROTOCOL_TOTAL_CHARS };
  const label = boundedProtocolString(
    record.label,
    MAX_PYTHON_PROTOCOL_FIELD_CHARS,
    budget,
    false,
  );
  if (label === null) return null;
  const parameters: string[] = [];
  for (const candidate of record.params.slice(0, MAX_PYTHON_PROTOCOL_ITEMS)) {
    const parameter = boundedProtocolString(
      candidate,
      MAX_PYTHON_PROTOCOL_NAME_CHARS,
      budget,
    );
    if (parameter === null) break;
    parameters.push(parameter);
  }
  const requested =
    Number.isSafeInteger(record.activeParameter) &&
    Number(record.activeParameter) >= 0
      ? Number(record.activeParameter)
      : 0;
  return {
    label,
    parameters,
    activeParameter: Math.min(requested, Math.max(parameters.length - 1, 0)),
  };
}

function boundedDefinition(value: unknown): PythonDefinitionInfo | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    record.found !== true ||
    !Number.isSafeInteger(record.line) ||
    Number(record.line) < 1 ||
    Number(record.line) > MAX_PYTHON_PROTOCOL_COORDINATE ||
    !Number.isSafeInteger(record.column) ||
    Number(record.column) < 0 ||
    Number(record.column) > MAX_PYTHON_PROTOCOL_COORDINATE
  ) {
    return null;
  }
  return {
    line: Number(record.line),
    column: Number(record.column),
    inBuffer: record.inBuffer === true,
  };
}

function boundedReferences(value: unknown): PythonReference[] {
  if (!value || typeof value !== "object" || Array.isArray(value)) return [];
  const raw = (value as Record<string, unknown>).refs;
  if (!Array.isArray(raw)) return [];
  const result: PythonReference[] = [];
  for (const candidate of raw.slice(0, MAX_PYTHON_PROTOCOL_ITEMS)) {
    if (
      !candidate ||
      typeof candidate !== "object" ||
      Array.isArray(candidate)
    ) {
      continue;
    }
    const record = candidate as Record<string, unknown>;
    if (
      !Number.isSafeInteger(record.line) ||
      Number(record.line) < 1 ||
      Number(record.line) > MAX_PYTHON_PROTOCOL_COORDINATE ||
      !Number.isSafeInteger(record.column) ||
      Number(record.column) < 0 ||
      Number(record.column) > MAX_PYTHON_PROTOCOL_COORDINATE ||
      !Number.isSafeInteger(record.length) ||
      Number(record.length) < 1 ||
      Number(record.length) > MAX_PYTHON_PROTOCOL_FIELD_CHARS
    ) {
      continue;
    }
    result.push({
      line: Number(record.line),
      column: Number(record.column),
      length: Number(record.length),
    });
  }
  return result;
}

function getWorker(namespace: string): WorkerHandle {
  const existing = workers.get(namespace);
  if (existing) return existing;
  const handle: WorkerHandle = {
    transport: null as unknown as PythonTransport,
    active: true,
    seq: 0,
    pending: new Map(),
    inspects: new Map(),
    lints: new Map(),
    formats: new Map(),
    completes: new Map(),
    hovers: new Map(),
    signatures: new Map(),
    definitions: new Map(),
    references: new Map(),
  };
  const onMessage = (raw: unknown) => {
    // A terminated local Worker should not deliver again, but injected transports
    // and cross-origin relays can race a final structured-clone message. Bind every
    // reply to the exact live generation before looking at ids (which restart at 1).
    if (!handle.active || workers.get(namespace) !== handle) return;
    const data = (raw ?? {}) as {
      id?: number;
      type?: string;
      stream?: LogStream;
      chunk?: string;
      logs?: string[];
      logStreams?: LogStream[];
      result?: string;
      error?: string;
      images?: string[];
      html?: string;
      variables?: VariableInfo[];
      lint?: PythonLintResult;
      warnings?: PythonFlakeWarning[];
      formatted?: string;
      completions?: PythonCompletionItem[];
      hover?: { found?: boolean } & Partial<PythonHoverInfo>;
      signature?: {
        found?: boolean;
        label?: string;
        params?: unknown;
        activeParameter?: unknown;
      };
      definition?: {
        found?: boolean;
        line?: unknown;
        column?: unknown;
        inBuffer?: unknown;
      };
      references?: { found?: boolean; refs?: unknown };
      debugEvents?: PythonDebugEvent[];
      debugTruncated?: boolean;
      debugConditionNote?: string;
    };
    if (typeof data.id !== "number") return;
    // Route by the exact pending owner before inspecting payload shape. A forged,
    // malformed, or cross-kind reply using a live format id must fail that request
    // immediately rather than fall into another shape branch and leave its timer
    // armed until it kills the shared kernel.
    const fmt = handle.formats.get(data.id);
    if (fmt) {
      const valid =
        data.error === PYTHON_FORMAT_SUCCESS_RECEIPT &&
        typeof data.formatted === "string" &&
        (fmt.fallback.trim().length === 0 || data.formatted.trim().length > 0) &&
        data.formatted.length <= MAX_PYTHON_FORMATTED_CHARS;
      settlePendingFormat(
        handle,
        data.id,
        fmt,
        valid
          ? { ok: true, formatted: data.formatted as string }
          : failedPythonFormat(fmt.fallback),
      );
      return;
    }
    // A lint response carries `lint` (parse-only syntax check + pyflakes warnings).
    if (data.lint && typeof data.lint === "object") {
      const lint = handle.lints.get(data.id);
      if (!lint) return;
      clearTimeout(lint.timer);
      handle.lints.delete(data.id);
      lint.resolve({
        lint: boundedLintResult(data.lint),
        warnings: boundedLintWarnings(data.warnings),
      });
      return;
    }
    // A completion response carries `completions` (jedi suggestions).
    if (Array.isArray(data.completions)) {
      const complete = handle.completes.get(data.id);
      if (!complete) return;
      clearTimeout(complete.timer);
      handle.completes.delete(data.id);
      complete.resolve(boundedCompletions(data.completions));
      return;
    }
    // A hover response carries `hover` ({found:false} ⇒ nothing under cursor).
    if (data.hover && typeof data.hover === "object") {
      const hover = handle.hovers.get(data.id);
      if (!hover) return;
      clearTimeout(hover.timer);
      handle.hovers.delete(data.id);
      hover.resolve(boundedHover(data.hover));
      return;
    }
    // A signature response carries `signature` ({found:false} ⇒ no call hint).
    if (data.signature && typeof data.signature === "object") {
      const signature = handle.signatures.get(data.id);
      if (!signature) return;
      clearTimeout(signature.timer);
      handle.signatures.delete(data.id);
      signature.resolve(boundedSignature(data.signature));
      return;
    }
    // A definition response carries `definition` ({found:false} ⇒ nothing to jump
    // to). `inBuffer` distinguishes a navigable same-buffer def from an out-of-file
    // one (the caller surfaces a notice for the latter, like the TS path).
    if (data.definition && typeof data.definition === "object") {
      const definition = handle.definitions.get(data.id);
      if (!definition) return;
      clearTimeout(definition.timer);
      handle.definitions.delete(data.id);
      definition.resolve(boundedDefinition(data.definition));
      return;
    }
    // A references response carries `references` ({refs:[]} ⇒ no in-buffer usages).
    if (data.references && typeof data.references === "object") {
      const references = handle.references.get(data.id);
      if (!references) return;
      clearTimeout(references.timer);
      handle.references.delete(data.id);
      references.resolve(boundedReferences(data.references));
      return;
    }
    // An incremental stream chunk (progressive output): route it to the pending
    // run's onOutput sink WITHOUT settling the run. The final message (below)
    // still carries the authoritative complete logs that get reconciled in.
    // A game's frame: the run is played, not timed, from its first frame on.
    if (data.type === "game-frame") {
      const pending = handle.pending.get(data.id);
      if (!pending) return;
      const frame = boundedGameFrame(raw as Record<string, unknown>);
      if (!frame) return;
      if (pending.gameNamespace === undefined) {
        pending.gameNamespace = namespace;
        clearTimeout(pending.timer);
        gameRuns.set(namespace, { handle, id: data.id as number });
      }
      emitGame(namespace, { kind: "frame", frame });
      return;
    }
    if (data.type === "stream") {
      const pending = handle.pending.get(data.id);
      if (!pending?.onOutput || typeof data.chunk !== "string") return;
      const chunk = pending.streamLimiter.push(data.chunk);
      if (chunk !== null) {
        pending.onOutput(chunk, data.stream === "stderr" ? "stderr" : "stdout");
      }
      return;
    }
    // An inspect response carries `variables` (no run output).
    if (Array.isArray(data.variables)) {
      const inspect = handle.inspects.get(data.id);
      if (!inspect) return;
      clearTimeout(inspect.timer);
      handle.inspects.delete(data.id);
      inspect.resolve(boundedVariables(data.variables));
      return;
    }
    const pending = handle.pending.get(data.id);
    if (!pending) return;
    const bounded = boundExecutionResult({
      logs: data.logs ?? [],
      logStreams: Array.isArray(data.logStreams) ? data.logStreams : undefined,
      result: data.result,
      error: data.error,
      images: data.images,
      html: data.html,
    });
    settlePendingRun(handle, data.id, pending, {
      ...bounded,
      // Debug-run recordings ride the normal final message (absent otherwise).
      // The worker itself caps at 300; repeat that at the trust boundary. A
      // terminal exception additionally needs an exact positive source line
      // and bounded non-empty text before it may drive editor navigation/UI.
      debugEvents: Array.isArray(data.debugEvents)
        ? data.debugEvents
            .slice(0, 300)
            .filter(
              (event): event is PythonDebugEvent =>
                !!event &&
                typeof event === "object" &&
                Number.isSafeInteger(event.line) &&
                event.line >= 1 &&
                (event.kind === "breakpoint" ||
                  event.kind === "step" ||
                  (event.kind === "exception" &&
                    typeof event.exception === "string" &&
                    event.exception.trim().length > 0)),
            )
            .map((event) =>
              event.kind === "exception"
                ? {
                    ...event,
                    exception: event.exception?.slice(
                      0,
                      MAX_PYTHON_DEBUG_EXCEPTION_CHARS,
                    ),
                  }
                : event,
            )
        : undefined,
      debugTruncated: data.debugTruncated === true ? true : undefined,
      debugConditionNote:
        typeof data.debugConditionNote === "string" && data.debugConditionNote
          ? data.debugConditionNote
          : undefined,
    });
  };
  const onError = (message: string) => {
    // A fatal worker error invalidates the interpreter: fail every pending run
    // and every auxiliary request, cancel their timers, then drop the worker so
    // the next run boots a fresh kernel.
    settleWorkerOperations(handle, {
      logs: [],
      error: `Python runtime error: ${message}`,
    });
    terminateWorkerGeneration(namespace, handle);
  };
  const origin = pythonSandboxOrigin();
  handle.transport =
    origin !== null && typeof document !== "undefined"
      ? createRemoteTransport(namespace, origin, onMessage, onError)
      : createLocalTransport(onMessage, onError);
  workers.set(namespace, handle);
  return handle;
}

function interruptedPythonRunResult(): PythonDebugRunResult {
  return {
    logs: [],
    error: "Execution interrupted.",
    interrupted: true,
  };
}

function runViaWorker(
  code: string,
  namespace: string,
  timeoutMs: number,
  onOutput?: OnOutput,
  stdin?: string,
  debugBreakpoints?: PythonDebugBreakpoint[],
  signal?: AbortSignal,
  pauseOnUncaughtExceptions = false,
): Promise<PythonDebugRunResult> {
  const interrupted = interruptedPythonRunResult();
  // A request cancelled before dispatch owns no worker generation. In
  // particular, it must not create a new kernel or retire an existing kernel in
  // the same namespace merely because its caller is already gone.
  if (signal?.aborted) return Promise.resolve(interrupted);

  let handle: WorkerHandle;
  try {
    handle = getWorker(namespace);
  } catch (error) {
    return Promise.resolve({
      logs: [],
      error: `Python runtime transport failed: ${
        error instanceof Error ? error.message : String(error)
      }`,
    });
  }
  const id = (handle.seq += 1);
  return new Promise<PythonDebugRunResult>((resolve) => {
    let dispatched = false;
    const timer = setTimeout(() => {
      const pending = handle.pending.get(id);
      if (!pending) return;
      settlePendingRun(handle, id, pending, {
        logs: [],
        error: `Execution timed out after ${Math.round(
          timeoutMs / 1000,
        )}s and was stopped.`,
        timedOut: true,
      });
      // Terminate the runaway interpreter; the next run recreates the worker
      // (equivalent to a kernel restart for this namespace). First settle every
      // other owner on this exact generation so no format/listener/timer leaks.
      settleWorkerOperations(handle, {
        logs: [],
        error: "Python runtime generation retired after execution timeout.",
        timedOut: true,
      });
      terminateWorkerGeneration(namespace, handle);
    }, timeoutMs);
    const pending: PendingRun = {
      resolve,
      timer,
      signal,
      onOutput,
      streamLimiter: createExecutionStreamLimiter(),
    };
    const abortListener = () => {
      if (handle.pending.get(id) !== pending) return;
      if (!dispatched) {
        // The request never entered this generation. Retiring the namespace
        // here could kill an older, unrelated run sharing its live kernel.
        settlePendingRun(handle, id, pending, interrupted);
        return;
      }
      // The worker protocol has no per-operation cancellation. Once dispatched,
      // terminating this exact generation is the only physical stop; settle
      // every sibling first so no continuation remains pinned to a dead worker.
      settleWorkerOperations(handle, interrupted);
      terminateWorkerGeneration(namespace, handle);
    };
    pending.abortListener = abortListener;
    handle.pending.set(id, pending);
    signal?.addEventListener("abort", abortListener, { once: true });
    // Close the check/register race: AbortSignal dispatch is synchronous, so an
    // abort immediately before listener registration is observed here.
    if (signal?.aborted) {
      abortListener();
      return;
    }
    // `stream: true` tells the worker to also post incremental `{type:"stream"}`
    // chunks as the kernel produces output. Older workers ignore the flag, so a
    // version skew degrades gracefully to the final-batch behaviour.
    try {
      // This line is the ownership boundary: an abort synchronously triggered
      // by postMessage itself must retire the generation, while an abort before
      // it must leave a pre-existing kernel alone.
      dispatched = true;
      handle.transport.postMessage({
        id,
        code,
        stream: Boolean(onOutput),
        ...(typeof stdin === "string" && stdin !== "" ? { stdin } : {}),
        ...(debugBreakpoints
          ? {
              op: "debug-run",
              breakpoints: debugBreakpoints,
              pauseOnUncaughtExceptions,
            }
          : {}),
      });
    } catch (error) {
      // A synchronous transport failure invalidates this generation just like a
      
      // their timers/listeners, then retire only this exact handle.
      settleWorkerOperations(handle, {
        logs: [],
        error: `Python runtime transport failed: ${
          error instanceof Error ? error.message : String(error)
        }`,
      });
      terminateWorkerGeneration(namespace, handle);
    }
  });
}

export interface PythonRunOptions {
  loadPyodide?: PyodideLoader;
  /** Per-notebook namespace so separate notebooks don't share globals. */
  namespace?: string;
  /** Override the anti-freeze timeout (worker path only). */
  timeoutMs?: number;
  /**
   * Progressive-output sink: called with each `(chunk, stream)` of stdout/stderr
   * AS the kernel produces it, so the cell can render live before the run ends
   * (Jupyter streaming parity). OPTIONAL/additive — when omitted the run behaves
   * exactly as before (final batched result only). The final RunResult still
   * carries the complete logs and remains the source of truth.
   */
  onOutput?: OnOutput;
  /**
   * Pre-supplied standard input. `input()` consumes it line by line; reading
   * past the last line raises Python's normal EOFError. Worker-path only.
   */
  stdin?: string;
  /**
   * Cancel a worker-backed execution. Before dispatch this is a request-local
   * cancellation; after dispatch it physically retires that exact kernel
   * generation because Pyodide has no cooperative per-run interrupt channel.
   * The in-process fallback can only honour a signal that is already aborted.
   */
  signal?: AbortSignal;
}

export interface PythonDebugRunOptions extends PythonRunOptions {
  /** Record one terminal pause when an exception escapes the user's program. */
  pauseOnUncaughtExceptions?: boolean;
}

function usesWorkerTransport(loadPyodide?: PyodideLoader): boolean {
  // Injected loader → in-process path (unit tests). Otherwise use the worker
  // when the platform supports it: a local Worker, or — in independent-origin
  // mode — the sandbox iframe (which hosts the workers itself, so the app
  // context does not need the Worker constructor).
  if (loadPyodide || typeof document === "undefined") return false;
  return typeof Worker !== "undefined" || pythonSandboxOrigin() !== null;
}

/** Whether a Python runtime is already loaded (so the UI can warn on first run). */
export function isPythonRuntimeReady(): boolean {
  return cached !== null || workers.size > 0;
}

/**
 * Drop running interpreters so the next run boots fresh empty globals — the
 * "restart kernel" action. With no namespace, resets every kernel; otherwise
 * just the given notebook's worker.
 */
export function resetPythonRuntime(namespace?: string): void {
  cached = null;
  if (namespace !== undefined) {
    terminateWorker(namespace);
    return;
  }
  for (const key of Array.from(workers.keys())) terminateWorker(key);
}

/**
 * Stop the in-flight Python run for a namespace (Jupyter "interrupt"): resolve
 * the pending run as interrupted and terminate the worker. Pyodide runs on the
 * worker thread with no cooperative interrupt buffer wired here, so terminating
 * the worker is the reliable way to abort a runaway snippet; the next run boots a
 * fresh kernel (equivalent to a restart). Returns true if a run was aborted.
 */
export function interruptPython(namespace = DEFAULT_NAMESPACE): boolean {
  const handle = workers.get(namespace);
  if (!handle) return false;
  const aborted = handle.pending.size > 0;
  settleWorkerOperations(handle, {
    logs: [],
    error: "Execution interrupted.",
    interrupted: true,
  });
  terminateWorkerGeneration(namespace, handle);
  return aborted;
}

/** Timeout for the (fast) inspect snapshot so a wedged kernel can't hang it. */
const INSPECT_TIMEOUT_MS = 5_000;
/** Timeout for the parse-only lint check (compile() is fast; cap a wedged kernel). */
const LINT_TIMEOUT_MS = 5_000;
/**
 * Timeout for format requests: autopep8 is fast but the first request pays the
 * one-off micropip install cost; cap a wedged kernel but give the install room.
 */
const FORMAT_TIMEOUT_MS = 30_000;
// Completions race a keystroke: a busy kernel (mid-run) should degrade to the
// keyword/identifier completions quickly rather than stall the popup. The
// first request also pays the one-off jedi package load, hence not too tight.
const COMPLETE_TIMEOUT_MS = 4_000;
// Hover shares the completion budget: a busy kernel degrades to "no tooltip".
const HOVER_TIMEOUT_MS = 4_000;
// Signature hints race a keystroke like completions/hover: a busy kernel degrades
// to "no parameter hint" within the same budget.
const SIGNATURE_TIMEOUT_MS = 4_000;
// Go-to-definition / rename are explicit USER actions (F12 / F2 / Cmd-click), not
// keystroke-triggered, so they tolerate a slightly longer budget than the typing-
// race ops; a busy kernel still degrades cleanly (no jump / no rename) rather than
// hanging the action.
const DEFINITION_TIMEOUT_MS = 6_000;
const REFERENCES_TIMEOUT_MS = 6_000;

export interface LintOptions {
  loadPyodide?: PyodideLoader;
  namespace?: string;
}

export interface CompleteOptions {
  namespace?: string;
}

/**
 * Whether a live kernel already exists for the namespace. Lets the completion
 * source skip semantic requests BEFORE the first run — booting Pyodide (a
 * multi-MB download) as a side effect of typing would be a hostile surprise.
 */
export function hasPythonKernel(namespace?: string): boolean {
  return workers.has(namespace ?? DEFAULT_NAMESPACE);
}

/**
 * Semantic Python completions from the LIVE kernel (jedi over the worker's
 * global namespace): `df.gro` completes pandas members after the user ran
 * `import pandas`. Worker-path only and only once a kernel exists — without
 * one it resolves `[]` and the keyword/identifier source still applies.
 * `line` is 1-based, `column` 0-based (jedi's own convention). Always resolves.
 */
export async function completePython(
  code: string,
  line: number,
  column: number,
  options: CompleteOptions = {},
): Promise<PythonCompletionItem[]> {
  if (!usesWorkerTransport(undefined) || typeof window === "undefined") {
    return [];
  }
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  if (!workers.has(namespace)) return [];
  const handle = getWorker(namespace);
  const id = (handle.seq += 1);
  return new Promise<PythonCompletionItem[]>((resolve) => {
    const timer = setTimeout(() => {
      handle.completes.delete(id);
      resolve([]);
    }, COMPLETE_TIMEOUT_MS);
    handle.completes.set(id, { resolve, timer });
    handle.transport.postMessage({ id, op: "complete", code, line, column });
  });
}

/**
 * jedi hover info (signature + docstring) from the LIVE kernel for the symbol
 * at `line` (1-based) / `column` (0-based). Same gating as `completePython`:
 * worker-path only, never boots Pyodide — without a kernel it resolves null
 * and the editor simply shows no tooltip. Always resolves.
 */
export async function hoverPython(
  code: string,
  line: number,
  column: number,
  options: CompleteOptions = {},
): Promise<PythonHoverInfo | null> {
  if (!usesWorkerTransport(undefined) || typeof window === "undefined") {
    return null;
  }
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  if (!workers.has(namespace)) return null;
  const handle = getWorker(namespace);
  const id = (handle.seq += 1);
  return new Promise<PythonHoverInfo | null>((resolve) => {
    const timer = setTimeout(() => {
      handle.hovers.delete(id);
      resolve(null);
    }, HOVER_TIMEOUT_MS);
    handle.hovers.set(id, { resolve, timer });
    handle.transport.postMessage({ id, op: "hover", code, line, column });
  });
}

/**
 * jedi parameter hints (the enclosing call's signature + active argument index)
 * from the LIVE kernel for the call at `line` (1-based) / `column` (0-based).
 * Same gating as `completePython` / `hoverPython`: worker-path only, never boots
 * Pyodide — without a kernel it resolves null and the editor shows no hint.
 * Always resolves.
 */
export async function signaturePython(
  code: string,
  line: number,
  column: number,
  options: CompleteOptions = {},
): Promise<PythonSignatureInfo | null> {
  if (!usesWorkerTransport(undefined) || typeof window === "undefined") {
    return null;
  }
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  if (!workers.has(namespace)) return null;
  const handle = getWorker(namespace);
  const id = (handle.seq += 1);
  return new Promise<PythonSignatureInfo | null>((resolve) => {
    const timer = setTimeout(() => {
      handle.signatures.delete(id);
      resolve(null);
    }, SIGNATURE_TIMEOUT_MS);
    handle.signatures.set(id, { resolve, timer });
    handle.transport.postMessage({ id, op: "signature", code, line, column });
  });
}

/**
 * jedi go-to-definition from the LIVE kernel for the symbol at `line` (1-based) /
 * `column` (0-based). Same gating as `completePython` / `hoverPython`: worker-path
 * only, never boots Pyodide — without a kernel it resolves null and the editor
 * shows the "not found" notice. `inBuffer` tells the caller whether the resolved
 * definition is navigable in THIS buffer (jump) or out of file (notice). Always
 * resolves.
 */
export async function definitionPython(
  code: string,
  line: number,
  column: number,
  options: CompleteOptions = {},
): Promise<PythonDefinitionInfo | null> {
  if (!usesWorkerTransport(undefined) || typeof window === "undefined") {
    return null;
  }
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  if (!workers.has(namespace)) return null;
  const handle = getWorker(namespace);
  const id = (handle.seq += 1);
  return new Promise<PythonDefinitionInfo | null>((resolve) => {
    const timer = setTimeout(() => {
      handle.definitions.delete(id);
      resolve(null);
    }, DEFINITION_TIMEOUT_MS);
    handle.definitions.set(id, { resolve, timer });
    handle.transport.postMessage({ id, op: "definition", code, line, column });
  });
}

/**
 * jedi find-references from the LIVE kernel for the symbol at `line` (1-based) /
 * `column` (0-based), FILTERED to this buffer (rename only edits the current
 * source). Same gating as `completePython` / `hoverPython`: worker-path only,
 * never boots Pyodide — without a kernel it resolves `[]` and the editor reports
 * no symbol to rename. Each reference is `{line, column, length}` (a 1-based line,
 * 0-based column, and the identifier's text length). Always resolves.
 */
export async function referencesPython(
  code: string,
  line: number,
  column: number,
  options: CompleteOptions = {},
): Promise<PythonReference[]> {
  if (!usesWorkerTransport(undefined) || typeof window === "undefined") {
    return [];
  }
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  if (!workers.has(namespace)) return [];
  const handle = getWorker(namespace);
  const id = (handle.seq += 1);
  return new Promise<PythonReference[]>((resolve) => {
    const timer = setTimeout(() => {
      handle.references.delete(id);
      resolve([]);
    }, REFERENCES_TIMEOUT_MS);
    handle.references.set(id, { resolve, timer });
    handle.transport.postMessage({ id, op: "references", code, line, column });
  });
}

// Parse-only syntax-check snippet for the in-process path: reads the source from
// a global the caller seeds, never executes it. Mirrors the worker's LINT_CHECK.
const PYTHON_LINT_IN_PROCESS = [
  "def _mrnote_lint(_src):",
  "    try:",
  "        compile(_src, '<lint>', 'exec')",
  "        return {'ok': True}",
  "    except SyntaxError as _e:",
  "        return {",
  "            'ok': False,",
  "            'lineno': getattr(_e, 'lineno', None),",
  "            'offset': getattr(_e, 'offset', None),",
  "            'endLineno': getattr(_e, 'end_lineno', None),",
  "            'endOffset': getattr(_e, 'end_offset', None),",
  "            'msg': getattr(_e, 'msg', None) or str(_e),",
  "        }",
  "    except Exception:",
  "        return {'ok': True}",
  "_mrnote_lint(__mrnote_lint_src)",
].join("\n");

/**
 * Parse-check Python `code` WITHOUT executing it (anti-freeze) and resolve the
 * first SyntaxError plus any pyflakes semantic warnings, or `{ lint: { ok: true },
 * warnings: [] }` when it compiles cleanly. Always resolves — empty source, an
 * unavailable runtime, or any failure yields the clean sentinel so the caller
 * never surfaces a spurious squiggle.
 *
 * Worker path (real browser): asks the live kernel via the `lint` op, which runs
 * both the syntax check (compile()) and the pyflakes static-analysis pass.
 * In-process path (tests/SSR with an injected loader): runs only the syntax check
 * (pyflakes requires micropip which is not available in the test runner).
 */
export async function lintPython(
  code: string,
  options: LintOptions = {},
): Promise<PythonLintAndWarnings> {
  if (code.trim().length === 0) return { lint: { ok: true }, warnings: [] };
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  // Worker path (real browser): parse-check + pyflakes on the kernel thread.
  if (usesWorkerTransport(options.loadPyodide)) {
    const handle = getWorker(namespace);
    const id = (handle.seq += 1);
    return new Promise<PythonLintAndWarnings>((resolve) => {
      const timer = setTimeout(() => {
        handle.lints.delete(id);
        resolve({ lint: { ok: true }, warnings: [] });
      }, LINT_TIMEOUT_MS);
      handle.lints.set(id, { resolve, timer });
      handle.transport.postMessage({ id, op: "lint", code });
    });
  }

  // In-process / test path: seed the source as a global and run the syntax check.
  // pyflakes is not run here (micropip not available in test/SSR environments).
  const loader = options.loadPyodide ?? defaultLoadPyodide;
  let pyodide: PyodideLike;
  try {
    pyodide = await getPyodide(loader);
  } catch {
    return { lint: { ok: true }, warnings: [] };
  }
  try {
    // Bind the source as data via globals so arbitrary text round-trips safely.
    const globals = (
      pyodide as unknown as {
        globals?: {
          set?: (k: string, v: unknown) => void;
          delete?: (k: string) => void;
        };
      }
    ).globals;
    if (globals?.set) {
      globals.set("__mrnote_lint_src", code);
    } else {
      // Fallback for fakes without a globals map: assign via a tiny statement.
      // JSON.stringify yields a Python-compatible literal for typical source.
      await pyodide.runPythonAsync(
        `__mrnote_lint_src = ${JSON.stringify(code)}`,
      );
    }
    const raw = await pyodide.runPythonAsync(PYTHON_LINT_IN_PROCESS);
    try {
      globals?.delete?.("__mrnote_lint_src");
    } catch {
      /* ignore cleanup failure */
    }
    return { lint: boundedLintResult(decodePythonResult(raw)), warnings: [] };
  } catch {
    return { lint: { ok: true }, warnings: [] };
  }
}

export interface FormatOptions {
  loadPyodide?: PyodideLoader;
  namespace?: string;
  signal?: AbortSignal;
}

/**
 * Format Python `code` via autopep8 (PEP 8 style), loaded lazily via micropip
 * on the first call. Returns the formatted source string on success, or the
 * ORIGINAL `code` unchanged when autopep8 is unavailable, the source has a syntax
 * error autopep8 cannot fix, or anything else goes wrong. Never throws.
 *
 * Worker path (real browser): dispatches to the Pyodide worker's `format` op.
 * In-process / test path: returns the original text (autopep8 requires micropip
 * which is unavailable in the test runner — the routing is the testable part).
 */
export async function formatPythonWithReceipt(
  code: string,
  options: FormatOptions = {},
): Promise<PythonFormatReceipt> {
  if (code.trim().length === 0) return { ok: true, formatted: code };
  if (options.signal?.aborted) return failedPythonFormat(code);
  // Worker path (real browser): format via the kernel thread.
  if (usesWorkerTransport(options.loadPyodide)) {
    const namespace = options.namespace ?? DEFAULT_NAMESPACE;
    let handle: WorkerHandle;
    try {
      handle = getWorker(namespace);
    } catch {
      return failedPythonFormat(code);
    }
    const id = (handle.seq += 1);
    return new Promise<PythonFormatReceipt>((resolve) => {
      const timer = setTimeout(() => {
        if (!handle.formats.has(id)) return;
        // A hard format timeout means this interpreter generation may be wedged.
        // Settle every owner before terminating it; the next run/format boots a
        // fresh generation. Ordinary AbortSignal retirement below deliberately
        // does NOT terminate the shared notebook kernel.
        settleWorkerOperations(handle, {
          logs: [],
          error: "Python formatter timed out.",
          timedOut: true,
        });
        terminateWorkerGeneration(namespace, handle);
      }, FORMAT_TIMEOUT_MS);
      const abortListener = () => {
        const current = handle.formats.get(id);
        if (current) {
          settlePendingFormat(handle, id, current, failedPythonFormat(code));
        }
      };
      const pending: PendingFormat = {
        resolve,
        timer,
        fallback: code,
        signal: options.signal,
        abortListener,
      };
      handle.formats.set(id, pending);
      options.signal?.addEventListener("abort", abortListener, { once: true });
      if (options.signal?.aborted) {
        abortListener();
        return;
      }
      try {
        handle.transport.postMessage({ id, op: "format", code });
      } catch {
        // A synchronous transport failure means this generation cannot accept
        // work. Settle the exact format and every sibling before retirement;
        // the public receipt remains total and never rejects.
        settleWorkerOperations(handle, {
          logs: [],
          error: "Python formatter transport failed.",
        });
        terminateWorkerGeneration(namespace, handle);
      }
    });
  }
  // In-process / test path: autopep8 is not available, so this is a failure
  // receipt rather than a false "already formatted" success.
  return failedPythonFormat(code);
}

/** Backward-compatible string-only wrapper retained for existing callers. */
export async function formatPython(
  code: string,
  options: FormatOptions = {},
): Promise<string> {
  return (await formatPythonWithReceipt(code, options)).formatted;
}

export interface InspectOptions {
  loadPyodide?: PyodideLoader;
  namespace?: string;
}

/**
 * Snapshot the kernel's user-defined globals as {name,type,preview}. Cheap and
 * safe: the worker enumerates `globals()` minus builtins/modules/functions and
 * never runs a pathological repr (capped). On the in-process/test path it does
 * the equivalent via a small introspection snippet. Always resolves (empty list
 * when no kernel exists yet or the runtime is unavailable).
 */
export async function inspectPython(
  options: InspectOptions = {},
): Promise<VariableInfo[]> {
  const namespace = options.namespace ?? DEFAULT_NAMESPACE;
  // Worker path (real browser): ask the live kernel for its globals.
  if (usesWorkerTransport(options.loadPyodide)) {
    // No kernel yet → nothing defined.
    if (!workers.has(namespace)) return [];
    const handle = getWorker(namespace);
    const id = (handle.seq += 1);
    return new Promise<VariableInfo[]>((resolve) => {
      const timer = setTimeout(() => {
        if (!handle.inspects.delete(id)) return;
        resolve([]);
        // A timed-out inspect can be a hostile __repr__/poisoned helper which is
        // still consuming the worker thread. Retire this exact generation so a
        // late reply cannot settle a future id and the next run starts clean.
        // Settle sibling operations first so their timers/listeners do not retain
        // a dead handle (especially a pending formatter on this namespace).
        settleWorkerOperations(handle, {
          logs: [],
          error: "Python runtime generation retired after inspect timeout.",
          timedOut: true,
        });
        terminateWorkerGeneration(namespace, handle);
      }, INSPECT_TIMEOUT_MS);
      handle.inspects.set(id, { resolve, timer });
      handle.transport.postMessage({ id, op: "inspect" });
    });
  }

  // In-process / test path: use the same no-user-repr snapshot policy. Real
  // Pyodide returns a list PyProxy; injected legacy fakes may return JSON.
  const loader = options.loadPyodide ?? defaultLoadPyodide;
  let pyodide: PyodideLike;
  try {
    pyodide = await getPyodide(loader);
  } catch {
    return [];
  }
  try {
    const snapshot = await pyodide.runPythonAsync(INSPECT_SAFE_IN_PROCESS);
    return boundedVariables(decodePythonResult(snapshot));
  } catch {
    return [];
  }
}

const INSPECT_SAFE_IN_PROCESS = [
  "def _mrai_safe_snapshot():",
  "    import types as _types",
  "    _safe_type = type",
  "    _safe_len = len",
  "    _safe_type_name = type.__getattribute__",
  "    _safe_str_repr = str.__repr__",
  "    _safe_int_repr = int.__repr__",
  "    _safe_float_repr = float.__repr__",
  "    _skip_types = (_types.ModuleType, _types.FunctionType, _types.BuiltinFunctionType)",
  "    def _type_name(_value):",
  "        try:",
  "            _name = _safe_type_name(_safe_type(_value), '__name__')",
  "            return _name[:128] if _safe_type(_name) is str else '?'",
  "        except BaseException:",
  "            return '?'",
  "    def _preview(_value, _kind):",
  "        try:",
  "            if _kind is str:",
  "                _text = _safe_str_repr(_value)",
  "            elif _kind is bool:",
  "                _text = 'True' if _value else 'False'",
  "            elif _kind is int:",
  "                _text = _safe_int_repr(_value)",
  "            elif _kind is float:",
  "                _text = _safe_float_repr(_value)",
  "            elif _value is None:",
  "                _text = 'None'",
  "            else:",
  "                return '<' + _type_name(_value) + '>'",
  "            return _text if _safe_len(_text) <= 120 else _text[:119] + '\u2026'",
  "        except BaseException:",
  "            return '<' + _type_name(_value) + '>'",
  "    _out = []",
  "    for _name, _value in list(globals().items()):",
  "        if not _name or _name.startswith('_'):",
  "            continue",
  "        _kind = _safe_type(_value)",
  "        if _kind in _skip_types:",
  "            continue",
  "        _out.append({'name': _name[:256], 'type': _type_name(_value), 'preview': _preview(_value, _kind)})",
  "        if _safe_len(_out) >= 500:",
  "            break",
  "    return _out",
  "_mrai_safe_snapshot()",
].join("\n");

export async function runPython(
  code: string,
  options: PythonRunOptions = {},
): Promise<RunResult> {
  if (options.signal?.aborted) return interruptedPythonRunResult();
  const configurationError = pythonExecutionConfigurationError();
  if (configurationError) {
    return { logs: [], error: configurationError };
  }
  if (usesWorkerTransport(options.loadPyodide)) {
    return runViaWorker(
      code,
      options.namespace ?? DEFAULT_NAMESPACE,
      options.timeoutMs ?? DEFAULT_PYTHON_TIMEOUT_MS,
      options.onOutput,
      options.stdin,
      undefined,
      options.signal,
    );
  }

  return runPythonInProcess(code, options);
}

/**
 * Recorded breakpoint debugging (worker `debug-run` op): runs the code to
 * completion in the SAME kernel/namespace as normal runs while recording each
 * breakpoint hit + a bounded step window (locals previews, user-frame stack).
 * The editor replays the recording with Continue/Step (see python-debug.ts).
 * Worker-path only: without one the result simply reports the limitation.
 * Default timeout is doubled — settrace adds per-line overhead.
 */
export async function debugRunPython(
  code: string,
  breakpoints: PythonDebugBreakpoint[],
  options: PythonDebugRunOptions = {},
): Promise<PythonDebugRunResult> {
  if (options.signal?.aborted) return interruptedPythonRunResult();
  const configurationError = pythonExecutionConfigurationError();
  if (configurationError) {
    return { logs: [], error: configurationError };
  }
  if (
    !usesWorkerTransport(options.loadPyodide) ||
    typeof window === "undefined"
  ) {
    return {
      logs: [],
      error: "Debugging requires the in-browser Python worker runtime.",
    };
  }
  return runViaWorker(
    code,
    options.namespace ?? DEFAULT_NAMESPACE,
    options.timeoutMs ?? DEFAULT_PYTHON_TIMEOUT_MS * 2,
    options.onOutput,
    options.stdin,
    breakpoints,
    options.signal,
    options.pauseOnUncaughtExceptions === true,
  );
}

async function runPythonInProcess(
  code: string,
  options: PythonRunOptions = {},
): Promise<RunResult> {
  // In-process fallback (tests / SSR). Pyodide runs in-origin; the app CSP
  // grants only 'wasm-unsafe-eval' for this.
  const loader = options.loadPyodide ?? defaultLoadPyodide;
  let pyodide: PyodideLike;
  try {
    pyodide = await getPyodide(loader);
  } catch (err) {
    return {
      logs: [],
      error: `Failed to load Python runtime: ${
        err instanceof Error ? err.message : String(err)
      }`,
    };
  }

  const capture = createExecutionLogCapture();
  const onOutput = options.onOutput;
  pyodide.setStdout({
    batched: (text) => {
      const retained = capture.push(text, "stdout");
      // Mirror each chunk to the progressive sink so the cell renders live; the
      // batched text from Pyodide is exactly one stream block per call.
      if (retained !== null) onOutput?.(retained, "stdout");
    },
  });
  pyodide.setStderr({
    batched: (text) => {
      const retained = capture.push(text, "stderr");
      if (retained !== null) onOutput?.(retained, "stderr");
    },
  });

  if (typeof pyodide.loadPackagesFromImports === "function") {
    try {
      await pyodide.loadPackagesFromImports(code);
    } catch {
      /* fall through to execution */
    }
  }

  try {
    const ret = await pyodide.runPythonAsync(code);
    const result = ret === undefined || ret === null ? undefined : String(ret);
    return boundExecutionResult({
      logs: capture.logs,
      logStreams: capture.logStreams,
      result,
    });
  } catch (err) {
    return boundExecutionResult({
      logs: capture.logs,
      logStreams: capture.logStreams,
      error: err instanceof Error ? err.message : String(err),
    });
  }
}
