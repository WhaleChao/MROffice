/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { announceAllowanceSpent } from "@mrnote/core/web/lib/allowance-spent.ts";
import { ApiRequestError, apiGet, apiPost } from "@mrnote/core/web/lib/api.ts";
import { codeRuntimeGuardFindings } from "./code-runtime-guard";
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import type { OnOutput } from "@mrnote/core/web/lib/code-runner/output-coalescer.ts";
import {
  emitPythonGame,
  registerPythonGameController,
  type PythonGameInput,
} from "@mrnote/core/web/lib/code-runner/python-runner.ts";
import { realtimeSocketOrigin } from "@mrnote/core/web/lib/realtime/socket-origin.ts";

export interface CloudGuiCapability {
  readonly available: boolean;
  readonly idleSeconds: number;
  readonly maxSeconds: number;
}

export type CloudGuiEndReason =
  | "exit"
  | "stopped"
  | "idle"
  | "limit"
  | "ai_allowance_exhausted"
  | "busy"
  | "unavailable"
  | "failed"
  | "superseded"
  | "left";

const END_REASONS: ReadonlySet<string> = new Set<CloudGuiEndReason>([
  "exit",
  "stopped",
  "idle",
  "limit",
  "ai_allowance_exhausted",
  "busy",
  "unavailable",
  "failed",
  "superseded",
  "left",
]);

/** The page's words for a window's end (CodeWindow passes its translations). */
export type CloudGuiMessage = (
  key:
    | "starting"
    | "exitCode"
    | "ended.cancelled"
    | `ended.${Exclude<CloudGuiEndReason, "exit" | "stopped" | "left">}`,
  values?: Record<string, string | number>,
) => string;

/** A cloud run in progress on a page: its Stop, from the moment Run starts it. */
interface ActiveCloudRun {
  stopRequested: boolean;
  stop: (() => void) | null;
}

const activeRuns = new Map<string, ActiveCloudRun>();

/** A cloud window runs (or is opening) for this page. */
export function cloudGuiRunActive(namespace: string): boolean {
  return activeRuns.has(namespace);
}

/**
 * Stop the page's cloud window the way its own Stop button does: the program
 * is ended, the run settles by itself with its output kept (never an
 * "interrupted" error), the session is closed and charged to that moment.
 * One still being opened is not opened. False when none runs.
 */
export function stopCloudGui(namespace: string): boolean {
  const run = activeRuns.get(namespace);
  if (!run) return false;
  run.stopRequested = true;
  run.stop?.();
  return true;
}

const UNAVAILABLE: CloudGuiCapability = Object.freeze({ available: false, idleSeconds: 0, maxSeconds: 0 });
/** A capability that could not be read is asked again after this long. */
const CAPABILITY_RETRY_MS = 60_000;
const MAX_SIDE = 2048;
const MAX_FRAME_BYTES = 6 * 1024 * 1024;
const MAX_TEXT_MESSAGE = 128 * 1024;
const MAX_CAPTION = 200;
/** Input waits while this much is still unsent (a slow link): moves are dropped. */
const INPUT_BACKLOG_BYTES = 256 * 1024;
const STOP_CLOSE_MS = 6_000;

let capability: CloudGuiCapability | null = null;
let capabilityLoad: Promise<CloudGuiCapability> | null = null;
let capabilityFailedAt = 0;

function parseCapability(value: unknown): CloudGuiCapability {
  if (!value || typeof value !== "object") return UNAVAILABLE;
  const record = value as Record<string, unknown>;
  const seconds = (field: unknown) =>
    typeof field === "number" && Number.isSafeInteger(field) && field > 0 && field <= 86_400 ? field : 0;
  if (record.available !== true) return UNAVAILABLE;
  return Object.freeze({
    available: true,
    idleSeconds: seconds(record.idle_seconds),
    maxSeconds: seconds(record.max_seconds),
  });
}

/** Whether cloud windows are offered here (read once; a failure is unavailable). */
export function cloudGuiCapability(): Promise<CloudGuiCapability> {
  if (capability) return Promise.resolve(capability);
  if (capabilityLoad) return capabilityLoad;
  if (capabilityFailedAt && Date.now() - capabilityFailedAt < CAPABILITY_RETRY_MS) {
    return Promise.resolve(UNAVAILABLE);
  }
  capabilityLoad = apiGet<unknown>("/api/v1/code-execution/gui/capability", undefined, {
    suppressUnauthorizedHandling: true,
  })
    .then((value) => {
      capability = parseCapability(value);
      return capability;
    })
    .catch(() => {
      capabilityFailedAt = Date.now();
      return UNAVAILABLE;
    })
    .finally(() => {
      capabilityLoad = null;
    });
  return capabilityLoad;
}

/** The capability already read, or null. */
export function cloudGuiCapabilityNow(): CloudGuiCapability | null {
  return capability;
}

export function __resetCloudGuiForTests(value: CloudGuiCapability | null = null): void {
  capability = value;
  capabilityLoad = null;
  capabilityFailedAt = 0;
}

/** A Python source that needs a window, a terminal or processes, and cloud
 * windows are offered: it runs in one instead of the page. */
export function runsInCloudGui(source: string): boolean {
  return Boolean(capability?.available) && codeRuntimeGuardFindings("python", source).length > 0;
}

function isImage(bytes: Uint8Array): "image/jpeg" | "image/png" | null {
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  ) {
    return "image/png";
  }
  return null;
}

/** One input event rebuilt from checked fields (the API and the sandbox check again). */
function boundedInput(event: PythonGameInput): Record<string, unknown> | null {
  const point = (value: unknown) =>
    typeof value === "number" && Number.isFinite(value) ? Math.max(0, Math.min(MAX_SIDE, Math.floor(value))) : null;
  switch (event.t) {
    case "kd":
    case "ku":
      if (typeof event.key !== "string" || typeof event.code !== "string") return null;
      if (event.key.length > 32 || event.code.length > 32) return null;
      return { t: event.t, key: event.key, code: event.code, repeat: event.repeat === true };
    case "mm":
    case "md":
    case "mu":
    case "wheel": {
      const x = point(event.x);
      const y = point(event.y);
      if (x === null || y === null) return null;
      if (event.t === "md" || event.t === "mu") {
        const button = typeof event.button === "number" && event.button >= 1 && event.button <= 5 ? event.button : null;
        return button === null ? null : { t: event.t, x, y, button };
      }
      if (event.t === "wheel") {
        const dy = typeof event.dy === "number" && Number.isFinite(event.dy) ? Math.max(-100_000, Math.min(100_000, event.dy)) : 0;
        return { t: "wheel", x, y, dy };
      }
      return { t: "mm", x, y };
    }
    case "blur":
      return { t: "blur" };
    default:
      return null;
  }
}

export interface CloudGuiRunOptions {
  /** The code page that runs it (its Run rules decide who may). */
  readonly pageId: string;
  /** The page's game namespace (PythonGameSurface's). */
  readonly namespace: string;
  readonly stdin?: string;
  readonly onOutput?: OnOutput;
  readonly signal?: AbortSignal;
  readonly message: CloudGuiMessage;
  /** Test seams. */
  readonly socketFactory?: (url: string) => WebSocket;
  readonly decode?: (blob: Blob) => Promise<ImageBitmap>;
}

const MAX_LOG_LINES = 10_000;

/**
 * Run ``source`` in a cloud window until it ends: the program exits, the
 * reader stops it, it is left idle or reaches its time, or the allowance runs
 * out. Resolves with its output; never rejects.
 */
export async function runPythonInCloudGui(source: string, options: CloudGuiRunOptions): Promise<RunResult> {
  const run: ActiveCloudRun = { stopRequested: false, stop: null };
  activeRuns.set(options.namespace, run);
  try {
    return await runCloudSession(source, options, run);
  } finally {
    if (activeRuns.get(options.namespace) === run) activeRuns.delete(options.namespace);
  }
}

async function runCloudSession(source: string, options: CloudGuiRunOptions, run: ActiveCloudRun): Promise<RunResult> {
  const { message, namespace } = options;
  const notStarted = (error: string): RunResult => ({ logs: [], error, notStarted: true });
  // Cancelled before it started (the page went away): said, never a silent
  // "finished" with no output.
  const cancelled = (): RunResult => ({ logs: [], interrupted: true, notStarted: true, error: message("ended.cancelled") });
  if (options.signal?.aborted) return cancelled();
  const origin = realtimeSocketOrigin();
  if (!origin) return notStarted(message("ended.unavailable"));
  let ticket: string;
  try {
    const minted = await apiPost<{ ticket?: unknown }>(
      "/api/v1/code-execution/gui/sessions",
      { page_id: options.pageId, source, stdin: options.stdin ?? "" },
      options.signal ? { signal: options.signal } : undefined,
    );
    if (typeof minted?.ticket !== "string" || minted.ticket.length < 16 || minted.ticket.length > 256) {
      return notStarted(message("ended.unavailable"));
    }
    ticket = minted.ticket;
  } catch (error) {
    if (run.stopRequested) return { logs: [] };
    if (options.signal?.aborted) return cancelled();
    if (error instanceof ApiRequestError && error.code === "ai_allowance_exhausted") {
      // lib/api.ts already opened the allowance dialog for the 402.
      return notStarted(message("ended.ai_allowance_exhausted"));
    }
    if (error instanceof ApiRequestError && error.status === 503) {
      return notStarted(message("ended.unavailable"));
    }
    return notStarted(error instanceof Error && error.message ? error.message : message("ended.failed"));
  }

  // Stopped while its ticket was minted: no window is opened.
  if (run.stopRequested) return { logs: [] };
  if (options.signal?.aborted) return cancelled();
  const url = `${origin}/api/v1/code-execution/gui/ws?ticket=${encodeURIComponent(ticket)}`;
  const socket = (options.socketFactory ?? ((target: string) => new WebSocket(target)))(url);
  socket.binaryType = "arraybuffer";
  const decode = options.decode ?? ((blob: Blob) => createImageBitmap(blob));

  const logs: string[] = [];
  const logStreams: ("stdout" | "stderr")[] = [];
  const partial: Record<"stdout" | "stderr", string> = { stdout: "", stderr: "" };
  const addLine = (stream: "stdout" | "stderr", line: string) => {
    if (logs.length >= MAX_LOG_LINES) return;
    logs.push(line);
    logStreams.push(stream);
  };
  const pushLines = (stream: "stdout" | "stderr", text: string) => {
    const lines = (partial[stream] + text).split("\n");
    partial[stream] = lines.pop() ?? "";
    for (const line of lines) addLine(stream, line);
  };
  const flushLines = (stream: "stdout" | "stderr") => {
    if (partial[stream]) addLine(stream, partial[stream]);
    partial[stream] = "";
  };

  let size: { width: number; height: number; caption: string } | null = null;
  let started = false;
  let ended = false;
  let endReason: CloudGuiEndReason | null = null;
  let exitCode: number | null = null;
  let lastBitmap: ImageBitmap | null = null;
  let decoding = 0;
  let decodeChain: Promise<void> = Promise.resolve();

  const send = (data: Record<string, unknown>): boolean => {
    if (socket.readyState !== WebSocket.OPEN) return false;
    try {
      socket.send(JSON.stringify(data));
      return true;
    } catch {
      return false;
    }
  };

  // Both Stop buttons (the window's and the run's) end it here: the program
  // is stopped and the API closes the session; if it does not answer, the
  // socket is closed.
  let closeTimer: ReturnType<typeof setTimeout> | null = null;
  const stopSession = () => {
    if (!send({ t: "stop" }) && socket.readyState !== WebSocket.OPEN) {
      try {
        socket.close(1000);
      } catch {
        /* closed */
      }
      return;
    }
    if (closeTimer !== null) return;
    closeTimer = setTimeout(() => {
      try {
        socket.close(1000);
      } catch {
        /* closed */
      }
    }, STOP_CLOSE_MS);
  };
  run.stop = stopSession;
  const unregister = registerPythonGameController(namespace, {
    input(events) {
      if (!started || ended) return false;
      const congested = socket.bufferedAmount > INPUT_BACKLOG_BYTES;
      const checked = events
        .map(boundedInput)
        .filter((event): event is Record<string, unknown> => event !== null)
        // On a slow link the pointer's path is dropped; presses never are.
        .filter((event) => !congested || event.t !== "mm");
      return checked.length > 0 ? send({ t: "in", e: checked.slice(0, 200) }) : false;
    },
    stop() {
      run.stopRequested = true;
      stopSession();
    },
  });

  const onAbort = () => stopSession();
  options.signal?.addEventListener("abort", onAbort, { once: true });

  const showFrame = (bytes: ArrayBuffer) => {
    if (!size || bytes.byteLength > MAX_FRAME_BYTES) return;
    const view = new Uint8Array(bytes);
    const type = isImage(view);
    if (!type) return;
    // Pictures are decoded in order; when they come faster than they decode,
    // the waiting ones are dropped (the next one shows the same windows).
    if (decoding >= 2) return;
    decoding += 1;
    const caption = size.caption;
    decodeChain = decodeChain
      .then(async () => {
        const bitmap = await decode(new Blob([bytes], { type }));
        if (ended || bitmap.width < 1 || bitmap.height < 1 || bitmap.width > MAX_SIDE || bitmap.height > MAX_SIDE) {
          bitmap.close?.();
          return;
        }
        emitPythonGame(namespace, {
          kind: "frame",
          frame: { width: bitmap.width, height: bitmap.height, caption, bitmap, cloud: true },
        });
        // The surface drew it (or will on its next animation frame, before
        // the next picture is decoded): the one before is done with.
        lastBitmap?.close?.();
        lastBitmap = bitmap;
      })
      .catch(() => undefined)
      .finally(() => {
        decoding -= 1;
      });
  };

  const handleText = (text: string) => {
    if (text.length > MAX_TEXT_MESSAGE) return;
    let data: unknown;
    try {
      data = JSON.parse(text);
    } catch {
      return;
    }
    if (!data || typeof data !== "object" || Array.isArray(data)) return;
    const record = data as Record<string, unknown>;
    switch (record.t) {
      case "status":
        if (record.s === "running") started = true;
        return;
      case "win": {
        const width = record.w;
        const height = record.h;
        if (
          typeof width !== "number" ||
          typeof height !== "number" ||
          !Number.isSafeInteger(width) ||
          !Number.isSafeInteger(height) ||
          width < 1 ||
          height < 1 ||
          width > MAX_SIDE ||
          height > MAX_SIDE
        ) {
          return;
        }
        const caption = typeof record.caption === "string" ? record.caption.slice(0, MAX_CAPTION) : "";
        size = { width, height, caption };
        return;
      }
      case "out": {
        const stream = record.s === "stderr" ? "stderr" : record.s === "stdout" ? "stdout" : null;
        if (!stream || typeof record.d !== "string") return;
        pushLines(stream, record.d);
        options.onOutput?.(record.d, stream);
        return;
      }
      case "out-cut":
        pushLines("stderr", "\n[output truncated]\n");
        return;
      case "end": {
        const reason = typeof record.reason === "string" && END_REASONS.has(record.reason) ? record.reason : "failed";
        endReason = reason as CloudGuiEndReason;
        exitCode = typeof record.code === "number" && Number.isSafeInteger(record.code) ? record.code : null;
        if (endReason === "ai_allowance_exhausted") {
          const details =
            record.details && typeof record.details === "object" && !Array.isArray(record.details)
              ? (record.details as Record<string, unknown>)
              : {};
          announceAllowanceSpent(details);
        }
        return;
      }
      default:
        return;
    }
  };

  const finished = new Promise<void>((resolve) => {
    socket.onmessage = (event: MessageEvent) => {
      if (typeof event.data === "string") handleText(event.data);
      else if (event.data instanceof ArrayBuffer) showFrame(event.data);
    };
    socket.onclose = () => resolve();
    socket.onerror = () => undefined;
  });
  await finished;
  ended = true;
  run.stop = null;
  if (closeTimer !== null) clearTimeout(closeTimer);
  options.signal?.removeEventListener("abort", onAbort);
  unregister();
  await decodeChain.catch(() => undefined);
  emitPythonGame(namespace, { kind: "ended" });
  lastBitmap = null;
  flushLines("stdout");
  flushLines("stderr");

  const result = (extra: Partial<RunResult>): RunResult => ({ logs, logStreams, ...extra });
  // Set by the socket's handler (TypeScript does not follow that closure).
  const said = endReason as CloudGuiEndReason | null;
  const code = exitCode as number | null;
  const reason: CloudGuiEndReason = said ?? (options.signal?.aborted || run.stopRequested ? "stopped" : "failed");
  switch (reason) {
    case "exit":
      return code ? result({ error: message("exitCode", { code }) }) : result({});
    case "stopped":
    case "left":
      return result(options.signal?.aborted ? { interrupted: true } : {});
    case "idle":
      return result({ error: message("ended.idle", { minutes: Math.round((capability?.idleSeconds ?? 900) / 60) }) });
    case "limit":
      return result({ error: message("ended.limit", { minutes: Math.round((capability?.maxSeconds ?? 1800) / 60) }) });
    case "ai_allowance_exhausted":
    case "busy":
    case "unavailable":
      return result({ error: message(`ended.${reason}`), ...(started ? {} : { notStarted: true }) });
    default:
      return result({ error: message(`ended.${reason}`), ...(started ? {} : { notStarted: true }) });
  }
}
