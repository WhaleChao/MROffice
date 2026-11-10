/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A large file downloaded in parallel ranges straight to a file the person
 * picks, with progress on the window's transfer line.
 *
 * Over the long path to the file host one TCP connection carries only part
 * of a line's speed; the host answers byte ranges over HTTP/1.1, so six
 * ranges in flight use six connections. Where the browser cannot write a
 * picked file (the File System Access API is Chromium's), or the file is
 * small, the caller falls back to the browser's own download.
 */

import { beginTransfer } from "./transfer-store";

/** Below this a plain download finishes before parallel ranges would help. */
export const PARALLEL_DOWNLOAD_MIN_BYTES = 32 * 1024 * 1024;
const RANGE_BYTES = 8 * 1024 * 1024;
const PARALLEL_RANGES = 6;
const RANGE_ATTEMPTS = 3;

interface WritableFile {
  write(data: { type: "write"; position: number; data: Blob }): Promise<void>;
  close(): Promise<void>;
  abort(reason?: unknown): Promise<void>;
}

interface SaveFileHandle {
  createWritable(): Promise<WritableFile>;
}

type SaveFilePicker = (options: { suggestedName: string }) => Promise<SaveFileHandle>;

function saveFilePicker(): SaveFilePicker | null {
  const picker = (globalThis as { showSaveFilePicker?: SaveFilePicker }).showSaveFilePicker;
  return typeof picker === "function" ? picker.bind(globalThis) : null;
}

/** Whether a file of this size downloads here in parallel ranges. */
export function canDownloadInParallel(sizeBytes: number): boolean {
  return sizeBytes >= PARALLEL_DOWNLOAD_MIN_BYTES && saveFilePicker() !== null;
}

async function fetchRange(url: string, start: number, end: number, signal: AbortSignal): Promise<Blob> {
  let lastError: unknown = null;
  for (let attempt = 0; attempt < RANGE_ATTEMPTS; attempt += 1) {
    try {
      const response = await fetch(url, {
        credentials: "include",
        cache: "no-store",
        headers: { Range: `bytes=${start}-${end}` },
        signal,
      });
      if (response.status !== 206) throw new Error(`range_${response.status}`);
      const blob = await response.blob();
      if (blob.size !== end - start + 1) throw new Error("range_short");
      return blob;
    } catch (error) {
      if (signal.aborted) throw error;
      lastError = error;
    }
  }
  throw lastError;
}

/**
 * Ask where to save, then download in parallel ranges. Call it straight from
 * the click (the picker needs the gesture); `resolveUrl` runs after the pick.
 * Resolves "saved" or "cancelled" (the person closed the picker).
 */
export async function downloadInParallel(options: {
  filename: string;
  sizeBytes: number;
  resolveUrl: () => Promise<string>;
}): Promise<"saved" | "cancelled"> {
  const picker = saveFilePicker();
  if (!picker) throw new Error("parallel_download_unavailable");
  let handle: SaveFileHandle;
  try {
    handle = await picker({ suggestedName: options.filename });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") return "cancelled";
    throw error;
  }
  const url = await options.resolveUrl();
  const writable = await handle.createWritable();
  const controller = new AbortController();
  const transfer = beginTransfer({
    direction: "download",
    name: options.filename,
    totalBytes: options.sizeBytes,
  });
  const total = options.sizeBytes;
  let nextStart = 0;
  let done = 0;
  const worker = async () => {
    while (nextStart < total) {
      const start = nextStart;
      const end = Math.min(total, start + RANGE_BYTES) - 1;
      nextStart = end + 1;
      const blob = await fetchRange(url, start, end, controller.signal);
      await writable.write({ type: "write", position: start, data: blob });
      done += blob.size;
      transfer.update(done);
    }
  };
  try {
    await Promise.all(Array.from({ length: PARALLEL_RANGES }, worker));
    await writable.close();
    return "saved";
  } catch (error) {
    controller.abort();
    await writable.abort(error).catch(() => undefined);
    throw error;
  } finally {
    transfer.end();
  }
}
