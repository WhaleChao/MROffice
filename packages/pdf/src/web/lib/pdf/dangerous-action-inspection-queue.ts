"use client";










/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const MAX_IN_FLIGHT = 2;
const KEEP = 64;

let running = 0;
const waiting: Array<() => void> = [];
const answers = new Map<string, unknown>();
const inFlight = new Map<string, Promise<unknown>>();

function pump(): void {
  while (running < MAX_IN_FLIGHT && waiting.length) {
    running += 1;
    waiting.shift()!();
  }
}

function remember(key: string, value: unknown): void {
  answers.delete(key);
  answers.set(key, value);
  while (answers.size > KEEP) answers.delete(answers.keys().next().value as string);
}

function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) return Promise.reject(new DOMException("The operation was aborted", "AbortError"));
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(new DOMException("The operation was aborted", "AbortError"));
    signal.addEventListener("abort", onAbort, { once: true });
    promise.then(
      (value) => { signal.removeEventListener("abort", onAbort); resolve(value); },
      (error: unknown) => { signal.removeEventListener("abort", onAbort); reject(error); },
    );
  });
}

/** The inventory for `key`, fetched by `fetcher` (never with the caller's
 * signal: one window closing does not cancel the request another waits
 * on). A failed answer is not kept: the next window asks again. */
export function fetchPdfDangerousActionInspection(
  key: string,
  fetcher: () => Promise<unknown>,
  signal: AbortSignal,
): Promise<unknown> {
  if (answers.has(key)) {
    const value = answers.get(key);
    remember(key, value);
    return signal.aborted ? Promise.reject(new DOMException("The operation was aborted", "AbortError"))
      : Promise.resolve(value);
  }
  let shared = inFlight.get(key);
  if (!shared) {
    shared = new Promise<void>((started) => { waiting.push(started); pump(); })
      .then(fetcher)
      .then((value) => { remember(key, value); return value; })
      .finally(() => {
        inFlight.delete(key);
        running -= 1;
        pump();
      });
    inFlight.set(key, shared);
  }
  return abortable(shared, signal);
}

/** Tests. */
export function resetPdfDangerousActionInspectionQueue(): void {
  running = 0;
  waiting.length = 0;
  answers.clear();
  inFlight.clear();
}
