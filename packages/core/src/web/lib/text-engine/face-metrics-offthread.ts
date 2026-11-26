/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Served faces' metrics read in a worker (face-metrics.worker.ts): the page's
 * thread neither decompresses a WOFF2 program nor loads fontkit to do it.
 * `readFaceMetricsOffThread(url)` answers with the metrics, or rejects when
 * there is no worker (none in this environment, it did not start, it broke)
 * or the worker could not read that face: the caller then reads it on the
 * page, as before (face-metrics.ts), so a face is never left unread because
 * of the worker. One request per URL at a time; a few workers (poolSize).
 */
import type { SfntMetrics } from "../pdf/engine/native/sfnt/sfnt-reader";
import { startResilientWorker } from "../worker-start";

type Response = { ready?: true; id?: number; metrics?: SfntMetrics; error?: string };

/**
 * A worker that has not said it loaded by then is taken for one that never
 * will. Long: on a cold cache over a slow link its module (fontkit and the
 * WOFF2 decoder) takes more than a few seconds, and a start given up at 4 s
 * was terminated mid-load, made again and given up again (production
 * WorkerStartFailed "face-metrics ... no word by its start deadline", about
 * three a day, mostly a try-it notebook's first load): every face of the
 * session was then read on the page's thread.
 */
const WORKER_START_MS = 30_000;
/**
 * How long a read waits for a worker still loading before it is read on the
 * page instead (the caller's fallback). The worker goes on loading and takes
 * the reads after it.
 */
export const FACE_METRICS_READ_WAIT_MS = 4_000;









function poolSize(): number {
  const asked = typeof window === "undefined" ? undefined : (window as { __faceMetricsWorkers?: unknown }).__faceMetricsWorkers;
  if (typeof asked === "number" && Number.isFinite(asked)) return Math.max(1, Math.min(8, Math.floor(asked)));
  const cores = typeof navigator === "undefined" ? 0 : navigator.hardwareConcurrency || 0;
  return Math.max(1, Math.min(4, Math.floor(cores / 4)));
}

interface Slot {
  worker: Worker | null;
  ready: Promise<boolean> | null;
  broken: boolean;
  start: ReturnType<typeof startResilientWorker>;
  /** Reads sent to this worker and not answered yet. */
  waiting: Map<number, Read>;
}
interface Read {
  id: number;
  url: string;
  resolve(metrics: SfntMetrics): void;
  reject(error: Error): void;
}
const slots: Slot[] = [];
let sequence = 0;
const inFlight = new Map<string, Promise<SfntMetrics>>();
/** Reads no worker has taken yet: each started worker takes the next one as
 * it has room (READS_PER_WORKER), so a worker still starting holds none up and
 * a burst of faces is shared by the workers as they come. */
const queued: Read[] = [];
const READS_PER_WORKER = 2;

function slotAt(index: number): Slot {
  let slot = slots[index];
  if (!slot) {
    slot = { worker: null, ready: null, broken: false, start: startResilientWorker("face-metrics"), waiting: new Map() };
    slots[index] = slot;
  }
  return slot;
}

function failAll(slot: Slot, reason: string): void {
  for (const read of slot.waiting.values()) read.reject(new Error(reason));
  slot.waiting.clear();
}

/** Hand queued reads to the started workers with room, the least busy first. */
function dispatch(): void {
  while (queued.length > 0) {
    let best: Slot | null = null;
    for (const slot of slots) {
      if (!slot || slot.broken || !slot.worker || slot.waiting.size >= READS_PER_WORKER) continue;
      if (!best || slot.waiting.size < best.waiting.size) best = slot;
    }
    if (!best?.worker) return;
    const read = queued.shift()!;
    best.waiting.set(read.id, read);
    best.worker.postMessage({ id: read.id, url: read.url });
  }
}

function startWorker(slot: Slot): Promise<boolean> {
  if (slot.ready) return slot.ready;
  if (slot.broken || typeof Worker === "undefined" || typeof window === "undefined") return Promise.resolve(false);
  // One more start after a pause when this one fails (lib/worker-start.ts);
  // none while the pause runs or once given up: the caller works here then.
  const started = slot.start.create(() => new Worker(new URL("./face-metrics.worker.ts", import.meta.url), { type: "module", name: "mrnote-face-metrics" }));
  if (!started) {
    if (slot.start.gaveUp) slot.broken = true;
    return Promise.resolve(false);
  }
  slot.ready = new Promise<boolean>((resolve) => {
    let startedNow = false;
    let over = false;
    const giveUp = (reason: string) => {
      // Once: an error and the start deadline must not count as two starts.
      if (over) return;
      over = true;
      clearTimeout(timer);
      // A worker that failed at its work, or a second failed start, is given
      // up; a first failed start leaves the next call a fresh one.
      if (slot.start.failed(started, startedNow ? `failed at work: ${reason}` : reason) === "gave_up") slot.broken = true;
      if (!slot.broken) slot.ready = null;
      started.terminate();
      if (slot.worker === started) slot.worker = null;
      failAll(slot, "face metrics worker stopped");
      resolve(false);
      // No worker left to take the queued reads (none started, none starting):
      // they are read on the page now, not after their wait.
      if (!slots.some((other) => other && !other.broken && (other.worker || other.ready))) {
        for (const read of queued.splice(0)) read.reject(new Error("no face metrics worker"));
      }
    };
    const timer = setTimeout(() => giveUp("no word by its start deadline"), WORKER_START_MS);
    started.onerror = () => giveUp("error");
    started.onmessageerror = () => giveUp("message error");
    started.onmessage = (event: MessageEvent<Response>) => {
      const data = event.data;
      if (data?.ready) {
        clearTimeout(timer);
        startedNow = true;
        slot.start.started(started);
        slot.worker = started;
        resolve(true);
        dispatch();
        return;
      }
      if (typeof data?.id !== "number") return;
      const read = slot.waiting.get(data.id);
      slot.waiting.delete(data.id);
      if (read) {
        if (data.metrics) read.resolve(data.metrics);
        else read.reject(new Error(data.error ?? "face metrics not read"));
      }
      dispatch();
    };
  });
  return slot.ready;
}

/** Queued reads past which one more worker is started: a deck's or a PDF's
 * handful of faces is read by the one worker as before (its start, a
 * decoder's compile, costs more than it saves), a document asking for its
 * CJK face's 200-odd slices at once gets the pool. */
const READS_PER_EXTRA_WORKER = 8;

/** The first worker, and one more for each READS_PER_EXTRA_WORKER reads queued. */
function startPool(): void {
  const wanted = Math.min(poolSize(), 1 + Math.floor(queued.length / READS_PER_EXTRA_WORKER));
  for (let index = 0; index < wanted; index += 1) void startWorker(slotAt(index));
}











export function warmFaceMetricsWorker(): void {
  if (typeof window === "undefined" || typeof Worker === "undefined") return;
  if ((window as { __faceMetricsWarmOff?: boolean }).__faceMetricsWarmOff === true) return;
  void startWorker(slotAt(0));
}

/** How long the app's page is left alone before its idle warm-up of the worker. */
export const FACE_METRICS_IDLE_WARM_MS = 8_000;

/**
 * The desktop app's page, idle: its face-metrics worker started ahead of any
 * document (its chunk and its WOFF2 decoder loaded while nothing waits, about
 * 20-30 MB kept). Started at the Word import's answer only, it was still
 * loading when the engine asked for the first page's faces in some opens of
 * a loaded machine: every read waited its 4 s and went to the page. The
 * returned function cancels a warm-up not begun. `window.__faceMetricsIdleWarmOff
 * = true`: not ahead (at the import's answer, as before).
 */
export function warmFaceMetricsWorkerWhenIdle(delayMs = FACE_METRICS_IDLE_WARM_MS): () => void {
  if (typeof window === "undefined" || typeof Worker === "undefined") return () => undefined;
  if ((window as { __faceMetricsIdleWarmOff?: boolean }).__faceMetricsIdleWarmOff === true) return () => undefined;
  let idle: number | null = null;
  const timer = window.setTimeout(() => {
    const go = () => {
      idle = null;
      warmFaceMetricsWorker();
    };
    if (typeof window.requestIdleCallback === "function") idle = window.requestIdleCallback(go, { timeout: 10_000 });
    else go();
  }, delayMs);
  return () => {
    window.clearTimeout(timer);
    if (idle !== null && typeof window.cancelIdleCallback === "function") window.cancelIdleCallback(idle);
  };
}

/** Whether faces can be read off the page's thread here (a browser with workers). */
export function faceMetricsOffThreadAvailable(): boolean {
  if (typeof Worker === "undefined" || typeof window === "undefined") return false;
  const size = poolSize();
  for (let index = 0; index < size; index += 1) if (!slotAt(index).broken) return true;
  return false;
}

export function readFaceMetricsOffThread(url: string): Promise<SfntMetrics> {
  const known = inFlight.get(url);
  if (known) return known;
  const absolute = typeof document === "undefined" ? url : new URL(url, document.baseURI).href;
  const read = new Promise<SfntMetrics>((resolve, reject) => {
    const entry: Read = { id: ++sequence, url: absolute, resolve, reject };
    queued.push(entry);
    // Read on the page instead when no worker has taken it by then (none
    // started, or they are all still loading); it stays a worker's once taken.
    const timer = setTimeout(() => {
      const at = queued.indexOf(entry);
      if (at < 0) return;
      queued.splice(at, 1);
      reject(new Error("no face metrics worker"));
    }, FACE_METRICS_READ_WAIT_MS);
    const settle = <T,>(done: (value: T) => void) => (value: T) => {
      clearTimeout(timer);
      done(value);
    };
    entry.resolve = settle(resolve);
    entry.reject = settle(reject);
    startPool();
    dispatch();
  });
  inFlight.set(url, read);
  const forget = () => {
    if (inFlight.get(url) === read) inFlight.delete(url);
  };
  read.then(forget, forget);
  return read;
}
