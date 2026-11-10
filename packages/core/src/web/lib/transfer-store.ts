/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The store behind the transfer line (lib/transfer-progress), free of React
 * so request helpers that server code also imports can report to it.
 *
 * Every file transfer in flight (a part upload, an export, a parallel
 * download), so one quiet progress line can show them all: how far, how
 * fast, and about how long is left.
 *
 * A transfer reports the bytes done so far; speed is an exponentially
 * smoothed rate over those reports, so a burst of parts finishing together
 * does not make the estimate jump. Finished transfers leave the list.
 */

export type TransferDirection = "upload" | "download";
export type TransferUmbrella = "primary" | "secondary";

export interface TransferSnapshot {
  readonly id: number;
  readonly direction: TransferDirection;
  readonly name: string;
  readonly totalBytes: number;
  readonly doneBytes: number;
  /** Bytes per second, smoothed; 0 until two reports give a rate. */
  readonly bytesPerSecond: number;
  /** Seconds the server has been preparing it (an export it is still
   * writing), before any byte moves; absent for an ordinary transfer. */
  readonly preparingSeconds?: number;
  /** A whole operation (an export from click to file), shown only while
   * nothing more specific moves; a "primary" one (the editor's, in its own
   * words) before a "secondary" one (the request's, named by format). */
  readonly umbrella?: TransferUmbrella;
}

export interface TransferHandle {
  /** The bytes done so far (monotonic; smaller values are ignored). */
  update(doneBytes: number): void;
  /** The transfer ended, completed or not; it leaves the list. */
  end(): void;
  /** Still being prepared, for this many seconds (``preparing`` transfers). */
  preparing(seconds: number): void;
}

interface Entry {
  snapshot: TransferSnapshot;
  lastAt: number;
  lastBytes: number;
}

const SMOOTHING = 0.3;
const entries = new Map<number, Entry>();
const listeners = new Set<() => void>();
let nextId = 1;
/** No transfers: one array, which the server snapshot returns too
 * (lib/transfer-progress), so the first client render matches by identity. */
export const NO_TRANSFERS: readonly TransferSnapshot[] = [];
let published: readonly TransferSnapshot[] = NO_TRANSFERS;

function publish(): void {
  published = [...entries.values()].map((entry) => entry.snapshot);
  for (const listener of listeners) listener();
}

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

export function beginTransfer(options: {
  direction: TransferDirection;
  name: string;
  totalBytes: number;
  /** The server is preparing it (no bytes yet): shown with its time so far. */
  preparing?: boolean;
  /** Covers a whole operation; a more specific transfer inside it (the
   * export's own download) is shown instead while it runs. */
  umbrella?: TransferUmbrella;
}): TransferHandle {
  const id = nextId++;
  const entry: Entry = {
    snapshot: {
      id,
      direction: options.direction,
      name: options.name,
      totalBytes: Math.max(0, options.totalBytes),
      doneBytes: 0,
      bytesPerSecond: 0,
      ...(options.preparing || options.umbrella ? { preparingSeconds: 0 } : {}),
      ...(options.umbrella ? { umbrella: options.umbrella } : {}),
    },
    lastAt: now(),
    lastBytes: 0,
  };
  entries.set(id, entry);
  publish();
  let ended = false;
  return {
    update(doneBytes) {
      if (ended) return;
      const done = Math.min(entry.snapshot.totalBytes, Math.max(0, doneBytes));
      if (done <= entry.snapshot.doneBytes) return;
      const at = now();
      const seconds = (at - entry.lastAt) / 1000;
      let rate = entry.snapshot.bytesPerSecond;
      if (seconds > 0.05) {
        const instant = (done - entry.lastBytes) / seconds;
        rate = rate > 0 ? rate + SMOOTHING * (instant - rate) : instant;
        entry.lastAt = at;
        entry.lastBytes = done;
      }
      entry.snapshot = { ...entry.snapshot, doneBytes: done, bytesPerSecond: rate };
      publish();
    },
    end() {
      if (ended) return;
      ended = true;
      entries.delete(id);
      publish();
    },
    preparing(seconds) {
      if (ended || entry.snapshot.preparingSeconds === undefined) return;
      entry.snapshot = { ...entry.snapshot, preparingSeconds: Math.max(0, seconds) };
      publish();
    },
  };
}

/** Seconds left for one transfer, or null while its rate is unknown. */
export function transferSecondsLeft(transfer: TransferSnapshot): number | null {
  if (transfer.bytesPerSecond <= 0) return null;
  return Math.max(0, (transfer.totalBytes - transfer.doneBytes) / transfer.bytesPerSecond);
}

export function subscribeTransfers(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The transfers in flight now (for a poll outside React). */
export function currentTransfers(): readonly TransferSnapshot[] {
  return published;
}

/** Test seam: forget every transfer. */
export function resetTransfersForTests(): void {
  entries.clear();
  publish();
}

/**
 * Read a response body to a Blob while the transfer line shows its progress
 * (when the size is known; otherwise it is read as before).
 */
export async function readResponseWithProgress(response: Response, name: string): Promise<Blob> {
  const total = Number(response.headers.get("content-length") || 0);
  if (!response.body || !(total > 0)) return response.blob();
  const type = response.headers.get("content-type") || "";
  const transfer = beginTransfer({ direction: "download", name, totalBytes: total });
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let done = 0;
  try {
    for (;;) {
      const { value, done: finished } = await reader.read();
      if (finished) break;
      chunks.push(value);
      done += value.byteLength;
      transfer.update(done);
    }
  } finally {
    transfer.end();
  }
  return new Blob(chunks as BlobPart[], { type });
}
