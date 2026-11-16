/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The source risk inspection of a PDF, run in a worker (source-risk.worker.ts).
 *
 * The reader inspects a file once after it loads (encryption, signatures,
 * XFA, attachments, tags). pdf-lib parses the whole file and walks its object
 * graph for that: seconds of the reader's thread on a long book, during which
 * no click on the page answered. Here the bytes go to a worker and the answer
 * comes back; without a worker, or when it fails, the inspection runs on the
 * main thread as before. The caller still owns the profile cache and the
 * digest binding (save-policy.ts); this is only where the work happens.
 */
import type { PdfSourceRiskProfile } from "./save-policy";

type Response = { ready?: true; id?: number; profile?: PdfSourceRiskProfile; error?: string };

/** A worker that has not said it loaded by then is taken for one that never will. */
const WORKER_START_MS = 4_000;
let worker: Worker | null = null;
let ready: Promise<boolean> | null = null;
let broken = false;
let sequence = 0;
const pending = new Map<number, (profile: PdfSourceRiskProfile | null) => void>();

function startWorker(): Promise<boolean> {
  if (ready) return ready;
  // The unit-test environment installs a Worker that never answers: waiting
  // WORKER_START_MS for it held every reader's inspection 4 s and timed the
  
  if (broken || typeof Worker === "undefined" || process.env.NODE_ENV === "test") return Promise.resolve(false);
  ready = new Promise<boolean>((resolve) => {
    let started: Worker;
    try {
      started = new Worker(new URL("./source-risk.worker.ts", import.meta.url), { type: "module", name: "mrnote-pdf-source-risk" });
    } catch {
      broken = true;
      resolve(false);
      return;
    }
    const giveUp = () => {
      broken = true;
      started.terminate();
      if (worker === started) worker = null;
      for (const settle of pending.values()) settle(null);
      pending.clear();
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
      const settle = pending.get(data.id);
      pending.delete(data.id);
      settle?.(data.profile ?? null);
    };
  });
  return ready;
}

/**
 * The profile of `bytes` from the worker, or null when there is no worker
 * (then the caller inspects on its own thread). The bytes are copied for
 * the worker; the caller's copy stays its own.
 */
export async function inspectPdfRiskProfileOffThread(
  bytes: Uint8Array,
): Promise<PdfSourceRiskProfile | null> {
  if (!(await startWorker()) || !worker) return null;
  const id = ++sequence;
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return new Promise<PdfSourceRiskProfile | null>((resolve) => {
    pending.set(id, resolve);
    try {
      worker!.postMessage({ id, bytes: copy.buffer }, [copy.buffer]);
    } catch {
      pending.delete(id);
      resolve(null);
    }
  });
}

/**
 * Start the worker now (no-op where there is none, or one is started): for
 * a caller that will inspect in a moment -- the reader asks the shared store
 * of profiles first, and a file the store does not hold then found the
 * worker still to start.
 */
export function warmPdfSourceRiskWorker(): void {
  void startWorker().catch(() => undefined);
}

/** Tests: forget the worker so the next call starts afresh. */
export function resetPdfSourceRiskWorkerForTests(): void {
  worker?.terminate();
  worker = null;
  ready = null;
  broken = false;
  for (const settle of pending.values()) settle(null);
  pending.clear();
}
