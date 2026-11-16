/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The page's side of the glyph outlines (office-text-outline.worker.ts):
 * asks for the glyphs it needs, a face at a time, and keeps what comes back
 * by face and glyph, size-independent (font units), least recently used out
 * once the kept path data passes {@link OUTLINE_CACHE_BYTES}.
 * `releaseOutlineGlyphs` drops it all (a document closed).
 */
import { NO_PATH_UNION, OutlineEngine, type FaceOutlines } from "./outline-engine";
import { startOutlineWorker } from "./outline-worker-start";
import { startResilientWorker } from "../worker-start";

export interface GlyphPath {
  readonly d: string;
  readonly merged: boolean;
  readonly bbox: readonly [number, number, number, number] | null;
  readonly unitsPerEm: number;
  /** Horizontal advance, font units. */
  readonly advance: number;
  /** The vertical origin's height above the baseline, font units. */
  readonly vOrigin: number;
}

/** Path data kept, in UTF-16 code units x2 (bytes), plus a fixed cost per glyph. */
export const OUTLINE_CACHE_BYTES = 8 * 1024 * 1024;
const ENTRY_OVERHEAD = 96;

const kept = new Map<string, GlyphPath | null>();
let keptBytes = 0;
const pending = new Map<string, Set<number>>();
const asked = new Set<string>();
const listeners = new Set<() => void>();
let flushScheduled = false;

const sizeOf = (path: GlyphPath | null) => ENTRY_OVERHEAD + (path ? path.d.length * 2 : 0);
const keyOf = (url: string, glyph: number) => `${glyph}@${url}`;

function remember(key: string, path: GlyphPath | null): void {
  const old = kept.get(key);
  if (old !== undefined) {
    keptBytes -= sizeOf(old);
    kept.delete(key);
  }
  kept.set(key, path);
  keptBytes += sizeOf(path);
  for (const [oldest, value] of kept) {
    if (keptBytes <= OUTLINE_CACHE_BYTES) break;
    kept.delete(oldest);
    keptBytes -= sizeOf(value);
  }
}

/** The kept outline of `glyph` in the face at `url`: the path, null when the
 * face cannot give one, "pending" while it is asked for. */
export function glyphPath(url: string, glyph: number): GlyphPath | null | "pending" {
  const key = keyOf(url, glyph);
  const known = kept.get(key);
  if (known !== undefined) {
    // Most recently used goes last.
    kept.delete(key);
    kept.set(key, known);
    return known;
  }
  if (!asked.has(key)) {
    asked.add(key);
    let glyphs = pending.get(url);
    if (!glyphs) pending.set(url, (glyphs = new Set()));
    glyphs.add(glyph);
    if (!flushScheduled) {
      flushScheduled = true;
      queueMicrotask(flush);
    }
  }
  return "pending";
}

/** Called when outlines asked for have come back. */
export function onGlyphPaths(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function outlineCacheBytes(): number {
  return keptBytes;
}

/** Drop every kept outline and the faces read for them. */
export function releaseOutlineGlyphs(): void {
  kept.clear();
  keptBytes = 0;
  asked.clear();
  pending.clear();
  if (worker) worker.postMessage({ release: true });
  pageEngine?.release();
}

function flush(): void {
  flushScheduled = false;
  const batch = [...pending];
  pending.clear();
  for (const [url, glyphs] of batch) {
    const ids = [...glyphs];
    void request(url, ids).then(
      (outlines) => {
        for (const [glyph, d, merged, bbox, advance, vOrigin] of outlines.glyphs) {
          remember(keyOf(url, glyph), { d, merged, bbox, unitsPerEm: outlines.unitsPerEm, advance, vOrigin });
        }
        for (const glyph of outlines.failed) remember(keyOf(url, glyph), null);
      },
      () => {
        for (const glyph of ids) remember(keyOf(url, glyph), null);
      },
    ).finally(() => {
      for (const glyph of ids) asked.delete(keyOf(url, glyph));
      for (const listener of [...listeners]) listener();
    });
  }
}

// --- where the outlines are made -------------------------------------------

type WorkerResponse = { ready?: true; id?: number; outlines?: FaceOutlines; union?: string[]; error?: string };
let worker: Worker | null = null;
let workerReady: Promise<boolean> | null = null;
let workerBroken = false;
const workerStart = startResilientWorker("text-outline");
let sequence = 0;
const waiting = new Map<number, { resolve(value: FaceOutlines | string[]): void; reject(error: Error): void }>();
let pageEngine: OutlineEngine | null = null;
let engineForTests: OutlineEngine | null = null;

/** Tests: make the outlines with this engine, on the page. */
export function setOutlineEngineForTests(engine: OutlineEngine | null): void {
  engineForTests = engine;
  releaseOutlineGlyphs();
}

function startWorker(): Promise<boolean> {
  if (workerReady) return workerReady;
  if (workerBroken || typeof Worker === "undefined" || typeof window === "undefined") return Promise.resolve(false);
  // One more start after a pause when this one fails (lib/worker-start.ts):
  // a classic worker, which loads PathKit with importScripts.
  const started = workerStart.create(startOutlineWorker);
  if (!started) {
    if (workerStart.gaveUp) workerBroken = true;
    return Promise.resolve(false);
  }
  workerReady = new Promise<boolean>((resolve) => {
    let startedNow = false;
    let over = false;
    const giveUp = (reason: string) => {
      // Once: an error and the start deadline must not count as two starts.
      if (over) return;
      over = true;
      clearTimeout(timer);
      if (workerStart.failed(started, startedNow ? `failed at work: ${reason}` : reason) === "gave_up") workerBroken = true;
      if (!workerBroken) workerReady = null;
      started.terminate();
      if (worker === started) worker = null;
      for (const settle of waiting.values()) settle.reject(new Error("text outline worker stopped"));
      waiting.clear();
      resolve(false);
    };
    const timer = setTimeout(() => giveUp("no word by its start deadline"), 4_000);
    started.onerror = () => giveUp("error");
    started.onmessageerror = () => giveUp("message error");
    started.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const data = event.data;
      if (data?.ready) {
        clearTimeout(timer);
        startedNow = true;
        workerStart.started(started);
        worker = started;
        resolve(true);
        return;
      }
      if (typeof data?.id !== "number") return;
      const settle = waiting.get(data.id);
      waiting.delete(data.id);
      if (!settle) return;
      if (data.outlines) settle.resolve(data.outlines);
      else if (data.union) settle.resolve(data.union);
      else settle.reject(new Error(data.error ?? "outlines not read"));
    };
  });
  return workerReady;
}

async function request(url: string, glyphs: number[]): Promise<FaceOutlines> {
  if (engineForTests) return engineForTests.outlines(url, glyphs);
  const absolute = typeof document === "undefined" ? url : new URL(url, document.baseURI).href;
  if (await startWorker()) {
    const id = ++sequence;
    const answer = new Promise<FaceOutlines | string[]>((resolve, reject) => waiting.set(id, { resolve, reject }));
    worker!.postMessage({ id, url: absolute, glyphs });
    return answer as Promise<FaceOutlines>;
  }
  // No worker here: the same engine on the page, without PathKit (a
  // glyph's overlapping contours are then stroked as they stand).
  pageEngine ??= new OutlineEngine(
    async (face) => {
      const response = await fetch(face, { cache: "force-cache", credentials: "same-origin" });
      if (!response.ok) throw new Error(`font ${response.status}`);
      return response.arrayBuffer();
    },
    NO_PATH_UNION,
  );
  return pageEngine.outlines(absolute, glyphs);
}

/**
 * Glyph outlines given as path data (a PDF's embedded face, read by the
 * caller), each with its overlapping contours merged by PathKit in the
 * outline worker; as they stand where there is no worker.
 */
export async function unionGlyphPaths(paths: readonly string[]): Promise<string[]> {
  if (engineForTests) return engineForTests.unionPaths(paths);
  if (!paths.length || !(await startWorker())) return [...paths];
  const id = ++sequence;
  const answer = new Promise<FaceOutlines | string[]>((resolve, reject) => waiting.set(id, { resolve, reject }));
  worker!.postMessage({ id, union: [...paths] });
  try {
    return (await answer) as string[];
  } catch {
    return [...paths];
  }
}
