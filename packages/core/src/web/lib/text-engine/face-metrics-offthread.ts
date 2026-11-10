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
 * of the worker. One request per URL at a time.
 */
import type { SfntMetrics } from "../pdf/engine/native/sfnt/sfnt-reader";

type Response = { ready?: true; id?: number; metrics?: SfntMetrics; error?: string };

/** A worker that has not said it loaded by then is taken for one that never will. */
const WORKER_START_MS = 4_000;
let worker: Worker | null = null;
let ready: Promise<boolean> | null = null;
let broken = false;
let sequence = 0;
const waiting = new Map<number, { resolve(metrics: SfntMetrics): void; reject(error: Error): void }>();
const inFlight = new Map<string, Promise<SfntMetrics>>();

function failAll(reason: string): void {
  for (const settle of waiting.values()) settle.reject(new Error(reason));
  waiting.clear();
}

function startWorker(): Promise<boolean> {
  if (ready) return ready;
  if (broken || typeof Worker === "undefined" || typeof window === "undefined") return Promise.resolve(false);
  ready = new Promise<boolean>((resolve) => {
    let started: Worker;
    try {
      started = new Worker(new URL("./face-metrics.worker.ts", import.meta.url), { type: "module", name: "mrnote-face-metrics" });
    } catch {
      broken = true;
      resolve(false);
      return;
    }
    const giveUp = () => {
      broken = true;
      started.terminate();
      if (worker === started) worker = null;
      failAll("face metrics worker stopped");
      resolve(false);
    };
    const timer = setTimeout(giveUp, WORKER_START_MS);
    started.onerror = giveUp;
    started.onmessageerror = giveUp;
    started.onmessage = (event: MessageEvent<Response>) => {
      const data = event.data;
      if (data?.ready) {
        clearTimeout(timer);
        worker = started;
        resolve(true);
        return;
      }
      if (typeof data?.id !== "number") return;
      const settle = waiting.get(data.id);
      waiting.delete(data.id);
      if (!settle) return;
      if (data.metrics) settle.resolve(data.metrics);
      else settle.reject(new Error(data.error ?? "face metrics not read"));
    };
  });
  return ready;
}

/** Whether faces can be read off the page's thread here (a browser with workers). */
export function faceMetricsOffThreadAvailable(): boolean {
  return !broken && typeof Worker !== "undefined" && typeof window !== "undefined";
}

export function readFaceMetricsOffThread(url: string): Promise<SfntMetrics> {
  const known = inFlight.get(url);
  if (known) return known;
  const absolute = typeof document === "undefined" ? url : new URL(url, document.baseURI).href;
  const read = startWorker().then((started) => {
    if (!started || !worker) throw new Error("no face metrics worker");
    const id = ++sequence;
    const answer = new Promise<SfntMetrics>((resolve, reject) => waiting.set(id, { resolve, reject }));
    worker.postMessage({ id, url: absolute });
    return answer;
  });
  inFlight.set(url, read);
  const forget = () => {
    if (inFlight.get(url) === read) inFlight.delete(url);
  };
  read.then(forget, forget);
  return read;
}
