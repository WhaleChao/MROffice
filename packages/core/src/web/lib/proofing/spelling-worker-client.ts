/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The page side of the spelling worker (public/proofing/spelling-worker.js):
 * one module worker for the page session, started on the first spelling
 * request, which loads Hunspell (WebAssembly) and every dictionary a check
 * asks for. A worker that cannot start, or dies, fails its requests; the
 * caller reports the dictionary as unavailable rather than checking on the
 * page thread.
 */

export const SPELLING_WORKER_URL = "/proofing/spelling-worker.js";

export type SpellingWorkerRequest =
  | { op: "load"; tag: string; aff: string; dic: string }
  | { op: "check"; tag: string; words: readonly string[]; lenientUppercaseAccents?: boolean }
  | { op: "suggest"; tag: string; word: string };

type SpellingWorkerReply = { id: number; ok: boolean; value?: unknown; error?: string };

/** What the client needs of a Worker (tests hand in an in-process one). */
export interface SpellingWorkerLike {
  postMessage(message: unknown): void;
  terminate(): void;
  onmessage: ((event: MessageEvent<SpellingWorkerReply>) => void) | null;
  onerror: ((event: Event | ErrorEvent) => void) | null;
}

type WorkerFactory = () => SpellingWorkerLike;

const defaultFactory: WorkerFactory = () =>
  new Worker(SPELLING_WORKER_URL, {
    type: "module",
    name: "mrnote-spelling",
  }) as unknown as SpellingWorkerLike;

let factory: WorkerFactory | null = null;
let worker: SpellingWorkerLike | null = null;
let sequence = 0;
const pending = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();

/** Whether this page can check spelling off the main thread. */
export function spellingWorkerAvailable(): boolean {
  return factory !== null || (typeof window !== "undefined" && typeof Worker !== "undefined");
}

function failAll(error: Error): void {
  for (const entry of pending.values()) entry.reject(error);
  pending.clear();
}

function started(): SpellingWorkerLike {
  if (worker) return worker;
  const spawned = (factory ?? defaultFactory)();
  spawned.onmessage = (event) => {
    const reply = event.data;
    const entry = typeof reply?.id === "number" ? pending.get(reply.id) : undefined;
    if (!entry) return;
    pending.delete(reply.id);
    if (reply.ok) entry.resolve(reply.value);
    else entry.reject(new Error(reply.error ?? "spelling worker request failed"));
  };
  spawned.onerror = () => {
    // A dead worker forgets every dictionary: the next request starts a new
    // one, which loads them again.
    if (worker === spawned) worker = null;
    spawned.terminate();
    failAll(new Error("spelling worker stopped"));
  };
  worker = spawned;
  return spawned;
}

export function spellingWorkerRequest<T>(request: SpellingWorkerRequest): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    let target: SpellingWorkerLike;
    try {
      target = started();
    } catch (error) {
      reject(error instanceof Error ? error : new Error(String(error)));
      return;
    }
    const id = ++sequence;
    pending.set(id, { resolve: resolve as (value: unknown) => void, reject });
    try {
      target.postMessage({ ...request, id });
    } catch (error) {
      pending.delete(id);
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
}

/** Tests only: run the worker through `next` (null: the real Worker). */
export function setSpellingWorkerFactoryForTests(next: WorkerFactory | null): void {
  worker?.terminate();
  worker = null;
  failAll(new Error("spelling worker replaced"));
  factory = next;
}
