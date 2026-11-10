/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The metrics of every served face file the engine has been asked about,
 * fetched and parsed once, lazily. Parsing is pure (sfnt-reader's
 * readSfntMetrics: no DOM, no outlines), so this module runs unchanged in a
 * worker; on the page it costs one fetch (normally a cache hit: the browser
 * fetched the same file to draw it) and well under a millisecond of parsing
 * for a Latin slice.
 *
 * `faceMetrics(url)` answers synchronously: the metrics, "pending" (the
 * load has started; a listener hears when it lands) or "failed".
 */
import { readSfntMetrics, type SfntMetrics } from "../pdf/engine/native/sfnt/sfnt-reader";
import { faceMetricsOffThreadAvailable, readFaceMetricsOffThread } from "./face-metrics-offthread";

export type FaceMetrics = SfntMetrics;
export type FaceMetricsState = FaceMetrics | "pending" | "failed";

const states = new Map<string, FaceMetricsState>();
const listeners = new Set<() => void>();
let generation = 0;
let notifyScheduled = false;

/** Bumped whenever a face finishes loading (or fails): measurements made
 * while it was pending may now answer differently. */
export function faceMetricsGeneration(): number {
  return generation;
}

/** Hear when faces land (batched to one call per task). Returns the unsubscribe. */
export function onFaceMetricsLoaded(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

function settled(): void {
  generation += 1;
  if (notifyScheduled) return;
  notifyScheduled = true;
  const run = () => {
    notifyScheduled = false;
    for (const listener of [...listeners]) {
      try { listener(); } catch { /* a listener's failure is its own */ }
    }
  };
  setTimeout(run, 0);
}

/** Other face data the engine loads beside the files (Office's pair
 * kerning, office-kerning.ts) landed: measurements may answer differently. */
export function noteFaceDataSettled(): void {
  settled();
}

type Fetcher = (url: string) => Promise<ArrayBuffer>;
const fetchServed: Fetcher = async (url) => {
  const response = await fetch(url, { cache: "force-cache", credentials: "same-origin" });
  if (!response.ok) throw new Error(`font ${response.status}`);
  return response.arrayBuffer();
};
let fetcher: Fetcher = fetchServed;

/** Tests and workers: where face bytes come from. */
export function setFaceFetcher(next: Fetcher): void {
  fetcher = next;
}

/** On the page, read in the metrics worker (face-metrics-offthread.ts); where
 * it is not to be had, or could not read the face, here, as before. */
function readFace(url: string): Promise<FaceMetrics> {
  const onPage = () => fetcher(url).then((buffer) => readSfntMetrics(new Uint8Array(buffer)));
  if (fetcher !== fetchServed || !faceMetricsOffThreadAvailable()) return onPage();
  return readFaceMetricsOffThread(url).catch(onPage);
}

const failures = new Map<string, number>();
const RETRIES = 3;

export function faceMetrics(url: string): FaceMetricsState {
  const known = states.get(url);
  if (known) return known;
  states.set(url, "pending");
  void readFace(url)
    .then(
      (metrics) => states.set(url, metrics),
      () => {
        // A fetch that failed (a busy dev server, a dropped connection) is
        // tried again a few times before the face is given up on: a face
        // given up on in one browser and not another breaks the agreement.
        const count = (failures.get(url) ?? 0) + 1;
        failures.set(url, count);
        if (count < RETRIES) {
          states.delete(url);
          setTimeout(() => faceMetrics(url), 2000 * count);
          return;
        }
        states.set(url, "failed");
      },
    )
    .finally(settled);
  return "pending";
}

/** Wait for faces to settle (tests, probes). */
export async function faceMetricsReady(urls: readonly string[]): Promise<void> {
  for (const url of urls) faceMetrics(url);
  for (let guard = 0; guard < 2000 && urls.some((url) => states.get(url) === "pending"); guard += 1) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

/** Register already-parsed metrics (tests, a worker handing results back). */
export function primeFaceMetrics(url: string, metrics: FaceMetrics): void {
  states.set(url, metrics);
  settled();
}

export function resetFaceMetricsForTests(): void {
  states.clear();
  generation += 1;
}
