"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart pictures from the server's chart engine, for every editor.
 *
 * The planner and the drawing run on the server (engine `charts`,
 * `lib/engine-host/charts-types.ts`); a host asks for the picture of a
 * chart's model at a size and shows it. This module is the one way to ask:
 *
 * - every chart asked for within one frame (16 ms) goes in ONE `render` call
 *   (a document opening asks for its charts in view together), at most 32 a
 *   call; identical requests share one answer;
 * - answers are kept in memory (an exact repeat answers at once) and in the
 *   browser's persistent store (`chart-picture-store.ts`), so a chart seen
 *   once draws on reopen and offline;
 * - a caller's abort detaches it; a queued request nobody waits for any more
 *   is never sent, so the newest of a burst of requests (a gallery hover, a
 *   theme pick) is the one drawn;
 * - a busy server is asked once more after its Retry-After.
 *
 * Style gallery tiles go the same way through `gallery`: every tile of one
 * gallery open in one request.
 */

import {
  callEngine,
  canonicalJson,
  EngineUnavailableError,
} from "../engine-host/client";
import {
  CHARTS_ENGINE_MAX_CHARTS,
  CHARTS_ENGINE_MAX_TILES,
  type ChartGalleryTile,
  type ChartGalleryTileResult,
  type ChartRenderPicture,
  type ChartRenderRequest,
  type ChartRenderResult,
} from "../engine-host/charts-types";
import { readStoredChartPicture, storeChartPicture } from "./chart-picture-store";

/** The engine refused this chart (invalid model, a picture too large). */
export class ChartPictureRefusedError extends Error {
  readonly code: string;

  constructor(code: string) {
    super(`chart picture refused: ${code}`);
    this.name = "ChartPictureRefusedError";
    this.code = code;
  }
}

const BATCH_WINDOW_MS = 16;
const MEMORY_ENTRIES = 96;
const MEMORY_CHARS = 12_000_000;
const MAX_BUSY_WAIT_MS = 2_000;

const abortError = () => new DOMException("The operation was aborted", "AbortError");

// ---------------------------------------------------------------------------
// Memory
// ---------------------------------------------------------------------------

const memory = new Map<string, ChartRenderPicture>();
let memoryChars = 0;

function remember(key: string, picture: ChartRenderPicture): void {
  const cost = key.length + picture.svg.length;
  if (cost > MEMORY_CHARS) return;
  const existing = memory.get(key);
  if (existing) {
    memoryChars -= key.length + existing.svg.length;
    memory.delete(key);
  }
  memory.set(key, picture);
  memoryChars += cost;
  while (memory.size > MEMORY_ENTRIES || memoryChars > MEMORY_CHARS) {
    const oldest = memory.keys().next().value as string | undefined;
    if (oldest === undefined) break;
    memoryChars -= oldest.length + memory.get(oldest)!.svg.length;
    memory.delete(oldest);
  }
}

/** The identity of a picture request (sorted-key JSON of the request). */
export function chartPictureKey(request: ChartRenderRequest): string {
  return canonicalJson(request);
}

/** The picture already given for this request in this page, if any. */
export function cachedChartPicture(key: string): ChartRenderPicture | undefined {
  const hit = memory.get(key);
  if (hit) {
    memory.delete(key);
    memory.set(key, hit);
  }
  return hit;
}

async function persistentHash(key: string): Promise<string | null> {
  const subtle = typeof crypto !== "undefined" ? crypto.subtle : undefined;
  if (!subtle) return null;
  try {
    const digest = await subtle.digest("SHA-256", new TextEncoder().encode(`charts.render\u0000${key}`));
    return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
  } catch {
    return null;
  }
}

/**
 * The picture this browser stored for the request on an earlier visit (the
 * persistent store), or null. A host shows it while the server's answer is on
 * its way, and keeps it when the server cannot be reached.
 */
export async function storedChartPicture(request: ChartRenderRequest): Promise<ChartRenderPicture | null> {
  const key = chartPictureKey(request);
  const hit = cachedChartPicture(key);
  if (hit) return hit;
  const hash = await persistentHash(key);
  return hash ? readStoredChartPicture(hash) : null;
}

// ---------------------------------------------------------------------------
// Batching
// ---------------------------------------------------------------------------

interface Waiter<T> {
  resolve(value: T): void;
  reject(error: unknown): void;
}

interface Job<Q, T> {
  key: string;
  input: Q;
  waiters: Set<Waiter<T>>;
  batch?: { jobs: Job<Q, T>[]; controller: AbortController };
}

/**
 * A queue that sends every input asked for within one frame in one call,
 * shares identical inputs, and drops inputs nobody waits for.
 */
class Batcher<Q, T> {
  private readonly queued = new Map<string, Job<Q, T>>();
  private readonly sent = new Map<string, Job<Q, T>>();
  private timer: ReturnType<typeof setTimeout> | null = null;

  constructor(
    private readonly max: number,
    private readonly send: (inputs: Q[], signal: AbortSignal) => Promise<(T | Error)[]>,
    private readonly settled: (key: string, value: T) => void,
  ) {}

  ask(key: string, input: Q, signal?: AbortSignal): Promise<T> {
    if (signal?.aborted) return Promise.reject(abortError());
    const sent = this.sent.get(key);
    // A sent call every waiter left is being aborted: ask again.
    let job = this.queued.get(key) ?? (sent && !sent.batch?.controller.signal.aborted ? sent : undefined);
    if (!job) {
      job = { key, input, waiters: new Set() };
      this.queued.set(key, job);
      this.schedule();
    }
    const owner = job;
    return new Promise<T>((resolve, reject) => {
      const waiter: Waiter<T> = { resolve, reject };
      owner.waiters.add(waiter);
      signal?.addEventListener(
        "abort",
        () => {
          if (!owner.waiters.delete(waiter)) return;
          reject(abortError());
          this.release(owner);
        },
        { once: true },
      );
    });
  }

  /** A job nobody waits for: unsent, it is dropped; sent, its whole call is
   * aborted once no job of that call has a waiter. */
  private release(job: Job<Q, T>): void {
    if (job.waiters.size) return;
    if (this.queued.get(job.key) === job) {
      this.queued.delete(job.key);
      return;
    }
    const batch = job.batch;
    if (batch && batch.jobs.every((entry) => entry.waiters.size === 0)) batch.controller.abort();
  }

  private schedule(): void {
    if (this.queued.size >= this.max) {
      void this.flush();
      return;
    }
    this.timer ??= setTimeout(() => void this.flush(), BATCH_WINDOW_MS);
  }

  private async flush(): Promise<void> {
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    const jobs = [...this.queued.values()].slice(0, this.max);
    if (!jobs.length) return;
    for (const job of jobs) {
      this.queued.delete(job.key);
      this.sent.set(job.key, job);
    }
    if (this.queued.size) this.schedule();
    const batch = { jobs, controller: new AbortController() };
    for (const job of jobs) job.batch = batch;
    let answers: (T | Error)[] | null = null;
    let failure: unknown = null;
    try {
      answers = await this.send(jobs.map((job) => job.input), batch.controller.signal);
    } catch (error) {
      failure = error;
    }
    jobs.forEach((job, index) => {
      if (this.sent.get(job.key) === job) this.sent.delete(job.key);
      const answer = answers?.[index];
      if (answers && !(answer instanceof Error) && answer !== undefined) this.settled(job.key, answer);
      for (const waiter of job.waiters) {
        if (!answers) waiter.reject(failure);
        else if (answer instanceof Error || answer === undefined) waiter.reject(answer ?? new ChartPictureRefusedError("chart_render_failed"));
        else waiter.resolve(answer);
      }
      job.waiters.clear();
    });
  }
}

/** One engine call, asked once more when the server says it is busy. */
async function engineWithRetry<T>(call: () => Promise<T>, signal: AbortSignal): Promise<T> {
  try {
    return await call();
  } catch (error) {
    if (!(error instanceof EngineUnavailableError) || error.reason !== "busy" || signal.aborted) throw error;
    await new Promise<void>((resolve, reject) => {
      const timer = setTimeout(resolve, Math.min(MAX_BUSY_WAIT_MS, error.retryAfterMs ?? 1000));
      signal.addEventListener("abort", () => {
        clearTimeout(timer);
        reject(abortError());
      }, { once: true });
    });
    return call();
  }
}

const pictures = new Batcher<ChartRenderRequest, ChartRenderPicture>(
  CHARTS_ENGINE_MAX_CHARTS,
  async (charts, signal) => {
    const answer = await engineWithRetry(
      () => callEngine("charts", "render", { charts }, { signal }),
      signal,
    );
    const results: ChartRenderResult[] = Array.isArray(answer?.results) ? answer.results : [];
    return charts.map((_, index) => {
      const result = results[index];
      if (!result) return new ChartPictureRefusedError("chart_render_failed");
      return result.ok ? result : new ChartPictureRefusedError(result.code);
    });
  },
  (key, picture) => {
    remember(key, picture);
    void persistentHash(key).then((hash) => {
      if (hash) storeChartPicture(hash, picture);
    });
  },
);

/**
 * The server's picture of a chart: from memory when this page was already
 * given it, else in the next batch. Rejects with the engine's error
 * (`EngineUnavailableError`: offline, busy, failed), with
 * {@link ChartPictureRefusedError} for a chart the engine refuses, or with the
 * caller's AbortError.
 */
export function requestChartPicture(
  request: ChartRenderRequest,
  options: { signal?: AbortSignal } = {},
): Promise<ChartRenderPicture> {
  const key = chartPictureKey(request);
  const hit = cachedChartPicture(key);
  if (hit) return Promise.resolve(hit);
  return pictures.ask(key, request, options.signal);
}

/** Each named series' drawn colour in a picture, for keys a host draws beside
 * the chart (a data table's legend keys): the series' own paint, else the
 * colour cycle at its place (the planner's `chartV2PlanSeriesColors`). */
export function chartPictureSeriesColors(picture: ChartRenderPicture, names: readonly string[]): string[] {
  const cycle = picture.palette;
  return names.map((name, index) =>
    picture.series.find((series) => series.name === name)?.color ??
    (cycle.length ? cycle[index % cycle.length] : "#64748b"),
  );
}

// ---------------------------------------------------------------------------
// Style gallery tiles
// ---------------------------------------------------------------------------

export interface ChartGalleryTileRequest {
  attrs: Record<string, unknown>;
  nativeLook?: Record<string, unknown>;
  themeAccents?: readonly string[];
  /** The chart's size on its page (px): the tile is laid out at it. */
  width?: number;
  height?: number;
}

const tileMemory = new Map<string, ChartGalleryTileResult | null>();
const MAX_TILES_REMEMBERED = 96;

const tiles = new Batcher<ChartGalleryTileRequest, ChartGalleryTileResult | null>(
  CHARTS_ENGINE_MAX_TILES,
  async (requests, signal) => {
    // One chart's tiles share its look and accents: each is sent once.
    const looks: Record<string, unknown>[] = [];
    const lookKeys: string[] = [];
    const accents: string[][] = [];
    const accentKeys: string[] = [];
    const indexOf = <V,>(value: V, keys: string[], values: V[]) => {
      const key = canonicalJson(value);
      let index = keys.indexOf(key);
      if (index < 0) {
        index = keys.length;
        keys.push(key);
        values.push(value);
      }
      return index;
    };
    const wire: ChartGalleryTile[] = requests.map((request) => ({
      attrs: request.attrs,
      ...(request.nativeLook ? { look: indexOf(request.nativeLook, lookKeys, looks) } : {}),
      ...(request.themeAccents?.length ? { accents: indexOf([...request.themeAccents], accentKeys, accents) } : {}),
      ...(request.width && request.height ? { width: request.width, height: request.height } : {}),
    }));
    const answer = await engineWithRetry(
      () =>
        callEngine(
          "charts",
          "gallery",
          { tiles: wire, ...(looks.length ? { looks } : {}), ...(accents.length ? { accents } : {}) },
          { signal },
        ),
      signal,
    );
    return requests.map((_, index) => answer?.tiles?.[index] ?? null);
  },
  (key, tile) => {
    tileMemory.set(key, tile);
    while (tileMemory.size > MAX_TILES_REMEMBERED) tileMemory.delete(tileMemory.keys().next().value as string);
  },
);

export function chartGalleryTileKey(request: ChartGalleryTileRequest): string {
  return canonicalJson(request);
}

/** A tile already drawn in this page (`undefined` when not asked yet). */
export function cachedChartGalleryTile(key: string): ChartGalleryTileResult | null | undefined {
  return tileMemory.get(key);
}

/** One style tile, batched with every tile mounted in the same frame. */
export function requestChartGalleryTile(
  request: ChartGalleryTileRequest,
  options: { signal?: AbortSignal } = {},
): Promise<ChartGalleryTileResult | null> {
  const key = chartGalleryTileKey(request);
  if (tileMemory.has(key)) return Promise.resolve(tileMemory.get(key) ?? null);
  return tiles.ask(key, request, options.signal);
}

/** Tests: forget every picture and tile this page was given. */
export function resetChartPictureClientForTests(): void {
  memory.clear();
  memoryChars = 0;
  tileMemory.clear();
}
