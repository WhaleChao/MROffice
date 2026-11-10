/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The reference index of a PDF, built in a worker (reference-index.worker.ts).
 *
 * The viewer reads each page's text as before and hands it here page by page;
 * the folding and scanning of the whole document's text, which held the
 * reader's thread for a quarter of a second while a 300-page book opened, runs
 * in the worker. The index is the one `buildReferenceIndex` gives, except that
 * `entries` holds a page's searchable text only for the pages that have
 * mentions (the only ones the viewer measures). Without a worker, or when it
 * fails, the index is built here, whole, as before.
 */
import { buildReferenceIndex, type PdfReferenceIndex, type PdfReferencePageInput } from "./reference-index";

type Response = { ready?: true; id?: number; index?: PdfReferenceIndex; error?: string };

/** A worker that has not said it loaded by then is taken for one that never will. */
const WORKER_START_MS = 4_000;
let worker: Worker | null = null;
let ready: Promise<boolean> | null = null;
let broken = false;
let sequence = 0;
const pending = new Map<number, (index: PdfReferenceIndex | null) => void>();

function startWorker(): Promise<boolean> {
  if (ready) return ready;
  if (broken || typeof Worker === "undefined") return Promise.resolve(false);
  ready = new Promise<boolean>((resolve) => {
    let started: Worker;
    try {
      started = new Worker(new URL("./reference-index.worker.ts", import.meta.url), { type: "module", name: "mrnote-pdf-reference-index" });
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
      settle?.(data.index ?? null);
    };
  });
  return ready;
}

export interface ReferenceIndexBuilder {
  /** The next page in the viewer's visible-page order (null: its text could not be read). */
  add(page: PdfReferencePageInput | null): void;
  build(): Promise<PdfReferenceIndex>;
}

export function createReferenceIndexBuilder(): ReferenceIndexBuilder {
  const kept: Array<PdfReferencePageInput | null> = [];
  const id = ++sequence;
  const started = startWorker();
  let sent = 0;
  const flush = () => {
    if (!worker) return;
    for (; sent < kept.length; sent += 1) worker.postMessage({ id, page: kept[sent] });
  };
  void started.then(flush);
  return {
    add(page) {
      kept.push(page);
      // Each page crosses as it is read: no one long copy at the end.
      if (worker) flush();
    },
    async build() {
      if (await started) {
        flush();
        const index = await new Promise<PdfReferenceIndex | null>((resolve) => {
          pending.set(id, resolve);
          try {
            worker!.postMessage({ id, build: true });
          } catch {
            pending.delete(id);
            resolve(null);
          }
        });
        if (index) return index;
      }
      return buildReferenceIndex(kept);
    },
  };
}
