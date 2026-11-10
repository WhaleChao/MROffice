"use client";






















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PDFWorker } from "pdfjs-dist";
import { createPdfNativeTextWorkerSlots } from "./native-text-worker-slots";

/** How long a started worker waits for a document before it is ended. */
const SPARE_IDLE_MS = 60_000;

type WorkerSlots = ReturnType<typeof createPdfNativeTextWorkerSlots>;
let slots: WorkerSlots | null = null;
let retireTimer: ReturnType<typeof setTimeout> | null = null;

function available(): boolean {
  return typeof window !== "undefined" && typeof Worker === "function" &&
    // The unit-test environment's Worker never answers (source-risk-offthread).
    process.env.NODE_ENV !== "test";
}

function keepForAWhile(): void {
  if (retireTimer) clearTimeout(retireTimer);
  retireTimer = setTimeout(() => {
    retireTimer = null;
    slots?.dispose();
    slots = null;
  }, SPARE_IDLE_MS);
  (retireTimer as { unref?: () => void }).unref?.();
}

/** Start the worker the next PDF document opens on (no-op when one is
 * already started or there are no workers here). */
export function warmPdfReaderWorker(): void {
  if (!available()) return;
  slots ??= createPdfNativeTextWorkerSlots();
  slots.warm();
  keepForAWhile();
}

/**
 * A started worker for one document (the next one starts at once), or
 * undefined, when pdf.js then starts its own as before. The caller owns it:
 * `bindPdfWorkerToTask` ends it with the document's loading task.
 */
export async function takePdfReaderWorker(): Promise<PDFWorker | undefined> {
  if (!available()) return undefined;
  slots ??= createPdfNativeTextWorkerSlots();
  keepForAWhile();
  try {
    return await slots.takePreview();
  } catch {
    return undefined;
  }
}

/** End `worker` with `task`: destroying the loading task (directly, or
 * through its document's `destroy`) destroys the worker after it. */
export function bindPdfWorkerToTask<T extends { destroy: () => Promise<void> }>(task: T, worker: PDFWorker): T {
  const destroy = task.destroy.bind(task);
  let ended = false;
  const end = () => {
    if (ended) return;
    ended = true;
    try { worker.destroy(); } catch { /* an ended worker stays ended */ }
  };
  task.destroy = () => destroy().finally(end);
  return task;
}

/** Tests only. */
export function resetPdfReaderWorkerForTests(): void {
  if (retireTimer) clearTimeout(retireTimer);
  retireTimer = null;
  slots?.dispose();
  slots = null;
}
