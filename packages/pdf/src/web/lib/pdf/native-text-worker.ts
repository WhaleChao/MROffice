/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { startResilientWorker } from "@mrnote/core/web/lib/worker-start.ts";

export type PdfNativeTextWorkerChannel = "prepare" | "plan" | "planner";

export class PdfNativeTextWorkerUnavailable extends Error {
  constructor(reason: string) {
    super(reason);
    this.name = "PdfNativeTextWorkerUnavailable";
  }
}




export type PdfNativeTextWorkerLane = "interactive" | "background" | "compile";

type Pending = Readonly<{
  channel: PdfNativeTextWorkerChannel;
  owner: Worker;
  resolve: (payload: unknown) => void;
  reject: (error: unknown) => void;
}>;

const workers = new Map<PdfNativeTextWorkerLane, Worker>();
const nativeTextWorkerStart = startResilientWorker("pdf-native-text", { attempts: Number.POSITIVE_INFINITY });
const pending = new Map<string, Pending>();
/** What a channel does with an answer nobody waits for any more (an
 * aborted request's candidate bytes, cleared). */
const discards = new Map<PdfNativeTextWorkerChannel, (payload: unknown) => void>();

function retire(target: Worker, reason: string): void {
  for (const [lane, held] of workers) if (held === target) workers.delete(lane);
  target.terminate();
  const failed = [...pending].filter(([, entry]) => entry.owner === target);
  for (const [requestId] of failed) pending.delete(requestId);
  for (const [, entry] of failed) entry.reject(new PdfNativeTextWorkerUnavailable(reason));
}

function current(lane: PdfNativeTextWorkerLane = "interactive"): Worker {
  const held = workers.get(lane);
  if (held) return held;
  // A failed start is reported and the next one waits out a pause
  // (lib/worker-start.ts); it is never given up: an edit has no other way
  // to be prepared here, and waits ("later") while the Worker is away.
  const spawned = nativeTextWorkerStart.create(() => new Worker(new URL("./native-text.worker.ts", import.meta.url),
    { type: "module", name: lane === "background" ? "mrnote-pdf-native-text-background"
      : lane === "compile" ? "mrnote-pdf-native-text-compile" : "mrnote-pdf-native-text" }));
  if (!spawned) throw new PdfNativeTextWorkerUnavailable("worker-starting-again");
  spawned.onmessage = (event: MessageEvent<{ channel?: PdfNativeTextWorkerChannel; payload?: { requestId?: string } }>) => {
    nativeTextWorkerStart.started(spawned);
    const payload = event.data?.payload;
    const requestId = payload?.requestId ?? "";
    const entry = pending.get(requestId);
    if (!entry) {
      const channel = event.data?.channel;
      if (channel) discards.get(channel)?.(payload);
      return;
    }
    pending.delete(requestId);
    entry.resolve(payload);
  };
  spawned.onerror = (event) => {
    nativeTextWorkerStart.failed(spawned, event.message || "worker-error");
    retire(spawned, event.message || "worker-error");
  };
  spawned.onmessageerror = () => retire(spawned, "invalid-answer");
  workers.set(lane, spawned);
  return spawned;
}

/** Start the Worker ahead of its first request (its modules load while the
 * reader opens a box). */
export function prewarmPdfNativeTextWorker(): void {
  if (typeof Worker === "undefined") return;
  try { current(); } catch { /* the first request says so */ }
}

/** How a channel clears an answer that comes after its request was given up. */
export function onPdfNativeTextWorkerDiscard(channel: PdfNativeTextWorkerChannel,
  discard: (payload: unknown) => void): void {
  discards.set(channel, discard);
}

/** The Worker failed this request's way (a preparation's own "worker-failed"):
 * it is replaced for the next. */
export function retirePdfNativeTextWorker(reason: string): void {
  for (const held of [...workers.values()]) retire(held, reason);
}

/**
 * Ask the Worker on a channel. Aborted, the request is let finish and its
 * answer dropped (the Worker keeps what it read); timed out, the Worker is
 * replaced.
 */
export function askPdfNativeTextWorker<Answer>(channel: PdfNativeTextWorkerChannel,
  payload: Readonly<{ requestId: string }>, options: Readonly<{
    timeoutMs: number; signal?: AbortSignal; transfer?: Transferable[];
    /** Which Worker answers: a read nobody waits on takes the background one. */
    lane?: PdfNativeTextWorkerLane;
  }>): Promise<Answer> {
  let target: Worker;
  try { target = current(options.lane); }
  catch (error) { return Promise.reject(new PdfNativeTextWorkerUnavailable(error instanceof Error ? error.message : "no-worker")); }
  return new Promise<Answer>((resolve, reject) => {
    const timer = setTimeout(() => retire(target, "timeout"), options.timeoutMs);
    const done = () => { clearTimeout(timer); options.signal?.removeEventListener("abort", onAbort); };
    const onAbort = () => {
      pending.delete(payload.requestId);
      done();
      // A plan not started yet is not started at all (native-text.worker.ts).
      if (channel === "plan") {
        try { target.postMessage({ channel, payload: { requestId: payload.requestId, cancel: true } }); } catch { /* gone */ }
      }
      reject(new DOMException("Cancelled", "AbortError"));
    };
    if (options.signal?.aborted) { done(); reject(new DOMException("Cancelled", "AbortError")); return; }
    options.signal?.addEventListener("abort", onAbort, { once: true });
    pending.set(payload.requestId, Object.freeze({
      channel,
      owner: target,
      resolve: (answer: unknown) => { done(); resolve(answer as Answer); },
      reject: (error: unknown) => { done(); reject(error); },
    }));
    try {
      target.postMessage({ channel, payload }, options.transfer ?? []);
    } catch (error) {
      pending.delete(payload.requestId);
      done();
      reject(new PdfNativeTextWorkerUnavailable(`post:${error instanceof Error ? error.message : String(error)}`));
    }
  });
}
