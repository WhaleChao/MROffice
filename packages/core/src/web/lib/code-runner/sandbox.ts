/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Sandboxed JavaScript execution (isolated, same-origin).
 *
 * User JS runs inside a `sandbox="allow-scripts"` iframe pointed at the
 * same-origin page /code-sandbox.html, which the middleware serves with its OWN
 * permissive CSP (script-src 'unsafe-eval'; worker-src blob:). The sandbox
 * attribute makes it an OPAQUE origin (no access to app cookies/storage/DOM),
 * and a real URL uses its own response CSP (unlike blob/srcdoc which inherit
 * the parent's) — so the APP origin's CSP needs NO 'unsafe-eval'. Code runs in
 * a Worker inside the iframe (off the main thread, terminable on timeout).
 *
 * This is the host side: a deterministic request/response protocol with timeout
 * + reset. The browser channel (iframe creation) is behind an injectable
 * interface so the protocol is unit-testable without a DOM. Browser-verified;
 * see the findings note above getBrowserChannel. (Python does NOT use this —
 * Pyodide stalls in an opaque origin — it runs in-origin with 'wasm-unsafe-eval'.)
 */

import type { LogStream } from "./cell-output-format";
import type { RunResult } from "./js-runner";
import type { OnOutput } from "./output-coalescer";
import {
  boundExecutionResult,
  createExecutionStreamLimiter,
} from "./execution-output-budget";
import {
  decodeSandboxRichOutput,
  placeSandboxErrorFrames,
  wrapSandboxJavaScriptForRichOutput,
} from "./sandbox-rich-output";

export type SandboxLanguage = "javascript" | "python";

/**
 * A debug session, carried on the same channel as a normal run.
 *
 * Debug used to be the ONE execution path that never entered the sandbox: it
 * built a blob Worker on the app's own origin, which inherits the app's CSP. In
 * development that ran user JavaScript beside the session cookie; in production,
 * where `unsafe-eval` was deliberately removed once execution moved in here, the
 * same worker could not evaluate anything at all. Neither environment showed
 * both halves, so neither was noticed.
 *
 * The harness travels with the request rather than living in the sandbox page.
 * The parent already sends the user's code to be evaluated in that worker, so
 * this is the same trust edge and not a new one — and it keeps ONE copy of a
 * harness that is assembled from shared fragments and covered by its own test,
 * instead of a second, drifting transcription inside static HTML.
 */
export interface SandboxDebugStart {
  /** The worker harness source (js-runner's DEBUG_WORKER_SCRIPT). */
  script: string;
  /** The first message the harness expects: `{type:"debug", code, …}`. */
  start: unknown;
}

export type SandboxPromptMode = "plain" | "await";

export interface SandboxRequest {
  id: string;
  language: SandboxLanguage;
  code: string;
  /** Optional per-notebook Python namespace so notebooks don't share globals. */
  namespace?: string;
  /**
   * Console session id. Submissions that carry one share a single long-lived
   * realm, so what a prompt defines is still there for the next line. Absent
   * (a file run), each execution gets a fresh realm and is torn down after it.
   */
  session?: string;
  /**
   * A console prompt submission. It only changes how the result is reported:
   * error frames name the submission's own lines (`<console>:2:9`) and the
   * evaluation wrapper's frames are left out, a returned Promise is shown with
   * its state, and an "await" submission is waited for and reported by the
   * value it settles to.
   */
  prompt?: SandboxPromptMode;
  /** Start a debug session for this id. */
  debug?: SandboxDebugStart;
  /** A message for a running debug session (resume, step). */
  debugControl?: unknown;
  /** Tear down a running debug session. */
  debugTerminate?: true;
  /** When true, the sandbox also posts incremental {@link SandboxStreamMessage}s. */
  stream?: boolean;
  /**
   * Terminate `session`'s realm or `targetId`'s one-shot worker instead of
   * running anything. This is the only way to stop code that will not yield
   * (`while (true) {}`), so Stop/timeout retire the exact execution.
   */
  reset?: boolean;
  /** Exact non-session run to terminate; keeps sibling runs in the page alive. */
  targetId?: string;
}

export interface SandboxResponse {
  id: string;
  result: RunResult;
}

/**
 * An incremental output message the sandbox posts before the final response (when
 * `stream` is requested): one stdout/stderr chunk. Distinguished from a
 * {@link SandboxResponse} by carrying `stream` instead of `result`.
 */
export interface SandboxStreamMessage {
  id: string;
  stream: { stream: LogStream; chunk: string };
}

/**
 * A message a debug session's worker produced, relayed verbatim. The debug
 * protocol is the worker's own (`debug-paused`, `stream`, the final result), so
 * the sandbox forwards rather than interprets.
 */
export interface SandboxDebugMessage {
  id: string;
  debugMessage: unknown;
}

/** A message from the sandbox: a final result, a chunk, or a debug relay. */
export type SandboxMessage =
  | SandboxResponse
  | SandboxStreamMessage
  | SandboxDebugMessage;

/** Type guard: a relayed message from a debug session's worker. */
export function isSandboxDebugMessage(
  message: SandboxMessage,
): message is SandboxDebugMessage {
  return "debugMessage" in message;
}

/** Type guard: a sandbox message that is an incremental stream chunk. */
export function isSandboxStreamMessage(
  message: SandboxMessage,
): message is SandboxStreamMessage {
  return (
    typeof (message as SandboxStreamMessage).stream === "object" &&
    (message as SandboxStreamMessage).stream !== null
  );
}

/** Bidirectional channel to the isolated execution context. */
export interface SandboxChannel {
  post(request: SandboxRequest): void;
  subscribe(handler: (message: SandboxMessage) => void): () => void;
  /** Tear down + recreate this page context after a channel/load failure. */
  reset(): void;
  /** Permanently release the iframe, listeners, timers and nested workers. */
  dispose?(): void;
  /**
   * Call `callback` once the request with this id has reached the sandbox page
   * (it is buffered until the page has loaded). A run's time limit counts from
   * there: the page's own load is not the program's time. Returns a cancel.
   */
  whenDelivered?(id: string, callback: () => void): () => void;
}

export interface SandboxRunOptions {
  timeoutMs?: number;
  namespace?: string;
   
  sandboxOwner?: string;
  /** Injectable channel (tests / custom hosts). Defaults to the iframe channel. */
  channel?: SandboxChannel;
  /** Abort the run (user Stop): kills its exact worker and resolves interrupted. */
  signal?: AbortSignal;
  /** Console session id — see {@link SandboxRequest.session}. */
  session?: string;
  /** Console prompt reporting — see {@link SandboxRequest.prompt}. */
  prompt?: SandboxPromptMode;
  /**
   * Progressive-output sink: called with each `(chunk, stream)` as the sandbox
   * produces it. OPTIONAL/additive — when omitted behaviour is unchanged.
   */
  onOutput?: OnOutput;
}

const DEFAULT_TIMEOUT_MS: Record<SandboxLanguage, number> = {
  javascript: 5_000,
  python: 30_000,
};

let idCounter = 0;
function nextRequestId(): string {
  idCounter += 1;
  return `sbx-${idCounter}`;
}

/**
 * Execute code in the sandbox, resolving with its RunResult. Always resolves
 * (never rejects): on timeout it terminates the exact worker and resolves with
 * a timedOut result.
 */
export function runInSandbox(
  language: SandboxLanguage,
  code: string,
  options: SandboxRunOptions = {},
): Promise<RunResult> {
  const channel =
    options.channel ?? getBrowserChannel(options.sandboxOwner);
  if (!channel) {
    return Promise.resolve({
      logs: [],
      error: "Sandbox runtime is unavailable in this environment.",
    });
  }
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS[language];
  const id = nextRequestId();
  const signal = options.signal;
  const prepared =
    language === "javascript"
      ? wrapSandboxJavaScriptForRichOutput(code, id)
      : { code, marker: null, lines: undefined };

  return new Promise<RunResult>((resolve) => {
    let settled = false;
    const streamLimiter = createExecutionStreamLimiter();
    let unsubscribe = () => {};
    let cancelDelivery = () => {};
    const cleanup = () => {
      unsubscribe();
      cancelDelivery();
      if (signal) signal.removeEventListener("abort", onAbort);
    };
    // Kill only this worker. Console realms are keyed by session; file/cell
    // workers by the original request id. Resetting the whole page channel here
    // would also kill a duplicate view's unrelated in-flight run.
    const killRunaway = () => {
      channel.post({
        id: nextRequestId(),
        language,
        code: "",
        ...(options.session
          ? { session: options.session }
          : { targetId: id }),
        reset: true,
      });
    };
    let timer: ReturnType<typeof setTimeout> | undefined;
    const armTimer = () => {
      if (settled || timer !== undefined) return;
      timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        cleanup();
        killRunaway();
        resolve({
          logs: [],
          error: `Execution timed out after ${timeoutMs} ms.`,
          timedOut: true,
        });
      }, timeoutMs);
    };

    const onAbort = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      cleanup();
      killRunaway();
      resolve({ logs: [], error: "Execution interrupted.", interrupted: true });
    };
    if (signal) {
      if (signal.aborted) {
        onAbort();
        return;
      }
      signal.addEventListener("abort", onAbort);
    }

    unsubscribe = channel.subscribe((message) => {
      if (message.id !== id || settled) return;
      // An incremental stream chunk: forward to onOutput WITHOUT settling.
      if (isSandboxStreamMessage(message)) {
        const chunk = streamLimiter.push(message.stream.chunk);
        if (chunk !== null) {
          options.onOutput?.(chunk, message.stream.stream);
        }
        return;
      }
      // A debug relay never settles a normal run: this subscriber belongs to
      // one, and the two share the channel but not the id.
      if (isSandboxDebugMessage(message)) return;
      settled = true;
      clearTimeout(timer);
      cleanup();
      resolve(
        boundExecutionResult(
          decodeSandboxRichOutput(
            placeSandboxErrorFrames(
              boundExecutionResult(message.result),
              prepared.lines,
            ),
            prepared.marker,
          ),
        ),
      );
    });

    // The time limit starts when the program reaches the sandbox, not while
    // the sandbox page itself is still loading (a first run): a channel that
    // cannot tell starts it now, as before.
    if (channel.whenDelivered) {
      cancelDelivery = channel.whenDelivered(id, armTimer);
    } else {
      armTimer();
    }
    channel.post({
      id,
      language,
      code: prepared.code,
      namespace: options.namespace,
      session: options.session,
      ...(options.prompt ? { prompt: options.prompt } : {}),
      stream: Boolean(options.onOutput),
    });
  });
}

/**
 * End a console session's realm and everything it was holding.
 *
 * A session worker is deliberately immortal between submissions, so something
 * has to bury it: the console's owner does, when its page goes away. Without
 * this, every page whose prompt was ever used would keep a worker (and its
 * globals, timers and intervals) alive for as long as the tab stayed open.
 *
 * Safe to call for a session that never ran or was already ended.
 */
export function endSandboxSession(
  session: string,
  channel: SandboxChannel | null = getBrowserChannel(),
): void {
  if (!session || !channel) return;
  channel.post({
    id: nextRequestId(),
    language: "javascript",
    code: "",
    session,
    reset: true,
  });
}

// ---------------------------------------------------------------------------
// Feature flag
// ---------------------------------------------------------------------------

let sandboxEnabledOverride: boolean | null = null;

/** Force the flag (tests). Pass null to clear. */
export function __setSandboxEnabledForTests(value: boolean | null): void {
  sandboxEnabledOverride = value;
}

/**
 * Whether JS execution should route through the isolated sandbox page.
 *
 * Browser-verified: the same-origin route /code-sandbox.html, served with its
 * own permissive CSP and loaded in a sandbox="allow-scripts" iframe, runs JS
 * under that own CSP in an OPAQUE origin (no app cookies/storage) — so the app
 * CSP needs no 'unsafe-eval'. Enabled in any browser; SSR has no DOM (null
 * channel) and falls back. (Python does NOT use this — Pyodide stalls in an
 * opaque origin — it runs in-origin with 'wasm-unsafe-eval'.)
 */
export function isSandboxExecutionEnabled(): boolean {
  if (sandboxEnabledOverride !== null) return sandboxEnabledOverride;
  return typeof document !== "undefined";
}

// ---------------------------------------------------------------------------
// Browser channel (iframe + worker) — pending browser verification
// ---------------------------------------------------------------------------

/**
 * BROWSER-VERIFIED FINDINGS (standalone harness, real browser):
 *  - A blob/srcdoc/data iframe INHERITS the app CSP — it cannot eval. But a
 *    REAL URL (even same-origin /code-sandbox.html) served with its OWN
 *    permissive CSP response header, loaded in a sandbox="allow-scripts"
 *    iframe, uses that own CSP and is an OPAQUE origin. So the sandbox does NOT
 *    need a separate domain — a same-origin route with its own CSP works.
 *  - JS verified: executes (logs captured), an infinite loop is timed out, and
 *    the iframe CANNOT read app cookies (SecurityError — opaque origin). So JS
 *    can be isolated WITHOUT 'unsafe-eval' in the app CSP.
 *  - Python NOT yet working in the sandbox: Pyodide stalls in the opaque-origin
 *    iframe (it uses storage/IndexedDB + workers that an opaque origin blocks;
 *    likely needs COEP/cross-origin-isolation or storage shims). Until solved,
 *    Python uses the in-origin path (app CSP 'wasm-unsafe-eval' + SRI).
 * Result: JS runs in this sandbox (default in any browser) and the app CSP
 * carries NO 'unsafe-eval' (only 'wasm-unsafe-eval' for in-origin Python).
 */
export const SANDBOX_BROWSER_CHECKLIST = true;

const DEFAULT_SANDBOX_OWNER = "__shared-sandbox__";
const cachedChannels = new Map<string, SandboxChannel>();
let channelOverride: SandboxChannel | null = null;

const VIBE_SHARED_SANDBOX_PATH = "/code-sandbox.html?profile=vibe-shared-v1";
const VIBE_SHARED_SANDBOX_CSP =
  "default-src 'none'; script-src 'unsafe-eval' 'unsafe-inline'; worker-src blob:; frame-ancestors 'self'; connect-src 'none'; base-uri 'none'; form-action 'none'";
const sharedAppChannels = new Map<
  string,
  {
    controller: AbortController;
    pending: Promise<SandboxChannel>;
    channel?: SandboxChannel;
  }
>();

/**
 * A separate, fixed no-network profile. Never fall back to the ordinary Code
 * sandbox (whose development policy deliberately remains unchanged). A stale
 * deployment without the explicit CSP must not execute shared application code.
 */
export function getVibeSharedSandboxChannel(
  owner: string,
): Promise<SandboxChannel> {
  if (!owner || typeof document === "undefined") {
    return Promise.reject(new Error("Shared application sandbox unavailable."));
  }
  const existing = sharedAppChannels.get(owner);
  if (existing) return existing.pending;
  const controller = new AbortController();
  const entry: {
    controller: AbortController;
    pending: Promise<SandboxChannel>;
    channel?: SandboxChannel;
  } = {
    controller,
    pending: Promise.resolve().then(async () => {
      const timer = setTimeout(() => controller.abort(), 5_000);
      try {
        const url = new URL(VIBE_SHARED_SANDBOX_PATH, window.location.origin);
        // The probe goes the way the frame itself will: a same-origin
        // navigation carries the session. Sent without credentials, it was
        // refused (401) by the private-beta gate that fronts every product
        // path in production, and every shared window read "sandbox policy
        
        // checked below, so a gate's login page can never pass as the sandbox.
        const response = await fetch(url.href, {
          method: "HEAD",
          credentials: "same-origin",
          cache: "no-store",
          redirect: "error",
          signal: controller.signal,
        });
        if (
          controller.signal.aborted ||
          sharedAppChannels.get(owner) !== entry ||
          !response.ok ||
          response.redirected ||
          response.url !== url.href ||
          response.headers.get("content-security-policy") !==
            VIBE_SHARED_SANDBOX_CSP
        ) {
          throw new Error("Shared application sandbox policy unavailable.");
        }
        entry.channel = createIframeChannel(url.href);
        return entry.channel;
      } catch (error) {
        if (sharedAppChannels.get(owner) === entry)
          sharedAppChannels.delete(owner);
        throw error;
      } finally {
        clearTimeout(timer);
      }
    }),
  };
  sharedAppChannels.set(owner, entry);
  return entry.pending;
}

/** Retire only this verified instance's profile, including a pending probe. */
export function disposeVibeSharedSandbox(owner: string): void {
  const entry = sharedAppChannels.get(owner);
  if (!entry) return;
  sharedAppChannels.delete(owner);
  entry.controller.abort();
  entry.channel?.dispose?.();
}

function sandboxOwnerKey(owner: string | undefined): string {
  return owner || DEFAULT_SANDBOX_OWNER;
}

/**
 * Stand a channel in for the iframe (tests). Pass null to clear.
 *
 * An import-level mock cannot reach this: callers inside this module resolve
 * `getBrowserChannel` locally, so a test that only patched the export would be
 * testing its own stub instead of the routing.
 */
export function __setBrowserChannelForTests(
  channel: SandboxChannel | null,
): void {
  for (const cached of cachedChannels.values()) cached.dispose?.();
  cachedChannels.clear();
  channelOverride = channel;
}

/** The same-origin sandbox-page iframe channel (browser only; null under SSR). */
export function getBrowserChannel(
  sandboxOwner?: string,
): SandboxChannel | null {
  if (channelOverride) return channelOverride;
  if (typeof document === "undefined") return null;
  const owner = sandboxOwnerKey(sandboxOwner);
  const cached = cachedChannels.get(owner);
  if (cached) return cached;
  const channel = createIframeChannel("/code-sandbox.html");
  cachedChannels.set(owner, channel);
  return channel;
}

/**
 * Release one page's complete JavaScript execution container.
 *
 * This deliberately does not call {@link getBrowserChannel}: closing a page
 * that never ran JavaScript must not create an iframe just to tear it down.
 * Removing the page-owned iframe kills its file, console and debug workers in
 * one browser-owned operation; other page keys remain untouched.
 */
export function disposeSandboxRuntime(sandboxOwner: string): void {
  if (!sandboxOwner) return;
  if (channelOverride) {
    channelOverride.dispose?.();
    return;
  }
  const owner = sandboxOwnerKey(sandboxOwner);
  const channel = cachedChannels.get(owner);
  if (!channel) return;
  cachedChannels.delete(owner);
  channel.dispose?.();
}

/**
 * How long the sandbox page has to say hello before we call it broken.
 *
 * Generous: this is a same-origin document from `public/`, so the only reasons
 * to miss it are a bad deploy or something blocking the frame — not slowness.
 */
const FRAME_READY_TIMEOUT_MS = 15_000;

/** The `WorkerLike` shape `runJavaScriptDebug` drives, restated to avoid a cycle. */
interface SandboxDebugWorker {
  postMessage(data: unknown): void;
  terminate(): void;
  onmessage: ((event: { data: never }) => void) | null;
  onerror: ((event: { message?: string }) => void) | null;
}

/**
 * A debug worker that lives INSIDE the sandbox.
 *
 * `runJavaScriptDebug` already speaks to its worker through a tiny interface —
 * post, terminate, onmessage, onerror — so moving debug into the sandbox is a
 * matter of giving it one backed by the iframe instead of a same-origin blob.
 * The debug protocol crosses unchanged and uninterpreted; the sandbox relays.
 *
 * `script` is the harness to run there. See {@link SandboxDebugStart} for why it
 * travels with the request instead of being transcribed into the sandbox page.
 */
export function createSandboxDebugWorker(
  script: string,
  sandboxOwner?: string,
): SandboxDebugWorker {
  const channel = getBrowserChannel(sandboxOwner);
  if (!channel) {
    throw new Error("The code sandbox is unavailable, so debugging cannot run.");
  }
  const id = nextRequestId();
  let started = false;
  let closed = false;

  const worker: SandboxDebugWorker = {
    onmessage: null,
    onerror: null,
    postMessage(data: unknown) {
      if (closed) return;
      if (!started) {
        started = true;
        // The first message is the harness's own start payload; everything
        // after it is a control message for a session already running.
        channel.post({
          id,
          language: "javascript",
          code: "",
          debug: { script, start: data },
        });
        return;
      }
      channel.post({ id, language: "javascript", code: "", debugControl: data });
    },
    terminate() {
      if (closed) return;
      closed = true;
      channel.post({ id, language: "javascript", code: "", debugTerminate: true });
      unsubscribe();
    },
  };

  const unsubscribe = channel.subscribe((message: SandboxMessage) => {
    if (closed || message.id !== id) return;
    if (isSandboxDebugMessage(message)) {
      worker.onmessage?.({ data: message.debugMessage as never });
      return;
    }
    // A plain result on a debug id is the sandbox reporting that the session
    // could not start at all (it never got as far as the harness).
    if ("result" in message) {
      const error = message.result?.error;
      worker.onerror?.({ message: error || "Sandbox debug session failed" });
    }
  });

  return worker;
}

/**
 * The real iframe channel, for tests that need to exercise its load handling.
 * Production reaches it through the cached accessor above.
 */
export function __createIframeChannelForTests(src: string): SandboxChannel {
  return createIframeChannel(src);
}

function createIframeChannel(src: string): SandboxChannel {
  const handlers = new Set<(message: SandboxMessage) => void>();
  let iframe: HTMLIFrameElement | null = null;
  let disposed = false;
  // The iframe registers its message listener only once /code-sandbox.html has
  // loaded, so it signals readiness with a { ready: true } message. Until then,
  // posts are buffered — otherwise a run fired immediately after mount (before
  // the iframe is live) would be dropped and spuriously time out.
  let ready = false;
  let pending: SandboxRequest[] = [];
  // `ready` is not proof the page is alive: `onload` fires for an error
  // document too, so a 404 would flush into a page with no listener and the run
  // would die of its own timeout, reporting "execution timed out" about a
  // two-line program. `handshaken` is the page ACTUALLY answering.
  let handshaken = false;
  let loadFailure: string | null = null;
  let readyTimer: ReturnType<typeof setTimeout> | null = null;
  // Ids posted but not yet answered, so a load failure can settle them instead
  // of leaving each run to time out separately with the wrong diagnosis.
  const inFlight = new Set<string>();
  // Runs waiting for their request to reach the page (see whenDelivered).
  const deliveryWaiters = new Map<string, () => void>();

  function delivered(id: string) {
    const callback = deliveryWaiters.get(id);
    if (!callback) return;
    deliveryWaiters.delete(id);
    callback();
  }

  /**
   * The sandbox page is not going to answer. Say THAT, to every run waiting on
   * it and to every run that starts before a reset — a timeout tells the user
   * to look for an infinite loop in code that has none, and editing the code
   * can never fix it. Python's channel has had this guard; this one did not.
   */
  function failLoad(message: string) {
    if (disposed || loadFailure) return;
    loadFailure = message;
    if (readyTimer !== null) {
      clearTimeout(readyTimer);
      readyTimer = null;
    }
    const waiting = [
      ...pending.map((request) => request.id),
      ...inFlight,
    ];
    pending = [];
    inFlight.clear();
    for (const id of new Set(waiting)) answerFailure(id);
  }

  function answer(
    id: string,
    result: RunResult,
  ) {
    const response: SandboxMessage = {
      id,
      result,
    };
    for (const handler of Array.from(handlers)) handler(response);
  }

  function answerFailure(id: string) {
    answer(id, {
      logs: [],
      error: loadFailure ?? "Sandbox unavailable.",
    });
  }

  function flush() {
    if (disposed || !ready || !iframe?.contentWindow) return;
    const queued = pending;
    pending = [];
    for (const request of queued) {
      iframe.contentWindow.postMessage(request, "*");
      delivered(request.id);
    }
  }

  // The sandbox iframe is sandbox="allow-scripts" -> OPAQUE origin, so its
  // messages arrive with event.origin "null"; match by source window instead.
  const onWindowMessage = (event: MessageEvent) => {
    if (disposed || !iframe || event.source !== iframe.contentWindow) return;
    const data = event.data as
      | (SandboxMessage & { ready?: boolean })
      | undefined;
    if (!data) return;
    if (data.ready === true) {
      handshaken = true;
      if (readyTimer !== null) {
        clearTimeout(readyTimer);
        readyTimer = null;
      }
      ready = true;
      flush();
      return;
    }
    if (typeof data.id !== "string") return;
    // A normal final result retires the id. Stream chunks and debug relays do
    // not: a debug relay may be a pause, and the worker is retired explicitly
    // when its controller sends debugTerminate.
    if (
      !isSandboxStreamMessage(data as SandboxMessage) &&
      !isSandboxDebugMessage(data as SandboxMessage)
    ) {
      inFlight.delete(data.id);
    }
    for (const handler of handlers) handler(data);
  };

  function mount() {
    if (disposed) return;
    ready = false;
    handshaken = false;
    loadFailure = null;
    const frame = document.createElement("iframe");
    // allow-scripts only (no allow-same-origin): the sandbox can't reach the
    // app's storage/cookies even though it is a real, CSP-permissive origin.
    frame.setAttribute("sandbox", "allow-scripts");
    frame.setAttribute("aria-hidden", "true");
    frame.style.display = "none";
    frame.src = src;
    // Fallback: if the ready handshake is missed, flush once the frame loads.
    // Kept, because a missed handshake on a healthy page should still run — but
    // it is NOT treated as proof of life; the deadline below still applies.
    frame.onload = () => {
      if (!ready) {
        ready = true;
        flush();
      }
    };
    frame.onerror = () => {
      failLoad(
        `The code sandbox (${src}) failed to load, so nothing can run. ` +
          "Reload the page; if it persists the deployment is missing that file.",
      );
    };
    document.body.appendChild(frame);
    iframe = frame;
    if (readyTimer !== null) clearTimeout(readyTimer);
    readyTimer = setTimeout(() => {
      readyTimer = null;
      if (handshaken) return;
      failLoad(
        `The code sandbox (${src}) did not respond within ` +
          `${Math.round(FRAME_READY_TIMEOUT_MS / 1000)}s, so nothing can run. ` +
          "Reload the page; if it persists the deployment is missing that file.",
      );
    }, FRAME_READY_TIMEOUT_MS);
  }

  function unmount() {
    if (iframe) {
      iframe.onload = null;
      iframe.onerror = null;
      if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
    }
    iframe = null;
  }

  function dispose() {
    if (disposed) return;
    disposed = true;
    if (readyTimer !== null) {
      clearTimeout(readyTimer);
      readyTimer = null;
    }
    const waiting = new Set([
      ...pending.map((request) => request.id),
      ...inFlight,
    ]);
    pending = [];
    inFlight.clear();
    window.removeEventListener("message", onWindowMessage);
    unmount();
    // Settle before dropping handlers so page-close never leaves a run promise
    // and its source/timer alive until the original timeout. Generation guards
    // in the hosts still reject any continuation after unmount.
    for (const id of waiting) {
      answer(id, {
        logs: [],
        error: "Sandbox runtime released.",
        interrupted: true,
      });
    }
    handlers.clear();
  }

  window.addEventListener("message", onWindowMessage);
  mount();

  return {
    post(request) {
      if (disposed) return;
      // Already known dead: answer now rather than let this run spend its whole
      // timeout budget to arrive at the wrong conclusion.
      if (loadFailure) {
        answerFailure(request.id);
        return;
      }
      if (request.debugTerminate === true) {
        inFlight.delete(request.id);
      } else if (request.debugControl === undefined) {
        inFlight.add(request.id);
      }
      // Buffer until the iframe is ready; otherwise post (opaque target -> "*").
      if (ready && iframe?.contentWindow) {
        iframe.contentWindow.postMessage(request, "*");
        delivered(request.id);
      } else {
        pending.push(request);
      }
    },
    whenDelivered(id, callback) {
      deliveryWaiters.set(id, callback);
      return () => {
        if (deliveryWaiters.get(id) === callback) deliveryWaiters.delete(id);
      };
    },
    subscribe(handler) {
      if (disposed) return () => {};
      handlers.add(handler);
      return () => handlers.delete(handler);
    },
    reset() {
      if (disposed) return;
      // Abandon any buffered/in-flight requests (the caller already timed out).
      // A run still waiting for delivery starts its time limit now, so it
      // ends the way it did before delivery was awaited.
      const abandoned = pending;
      pending = [];
      inFlight.clear();
      for (const request of abandoned) delivered(request.id);
      unmount();
      mount();
    },
    dispose,
  };
}
