/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Bytes gathered before they are written: one write per megabyte. */
const PIECE_BYTES = 1 << 20;
/** Bytes allowed to wait for the disk before the recording waits for it. */
const PENDING_LIMIT_BYTES = 16 << 20;

const SPOOL_DIRECTORY = "mrai-video-export";
/** How long a finished file is kept for its download. */
const KEEP_FINISHED_MS = 30 * 60 * 1000;

export interface ByteSpool {
  /** Add bytes at the end (kept in order). */
  append(data: Uint8Array): void;
  /** How many bytes were appended. */
  readonly size: number;
  /** Resolves once little enough waits to be written (the recorder's back-pressure). */
  drained(): Promise<void>;
  /** Every byte appended, as one Blob that refers to where they were written. */
  seal(): Promise<Blob>;
  /** Drop what was written. */
  discard(): Promise<void>;
}

/** The pieces kept as Blobs: the browser holds them, not the page. */
export function createBlobSpool(): ByteSpool {
  const parts: Blob[] = [];
  let pending: Uint8Array[] = [];
  let pendingBytes = 0;
  let size = 0;
  const flush = () => {
    if (pending.length === 0) return;
    parts.push(new Blob(pending as BlobPart[]));
    pending = [];
    pendingBytes = 0;
  };
  return {
    append(data) {
      if (data.length === 0) return;
      pending.push(data);
      pendingBytes += data.length;
      size += data.length;
      if (pendingBytes >= PIECE_BYTES) flush();
    },
    get size() {
      return size;
    },
    async drained() {},
    async seal() {
      flush();
      return new Blob(parts);
    },
    async discard() {
      parts.length = 0;
      pending = [];
      pendingBytes = 0;
    },
  };
}

type WritableFile = {
  write(data: BufferSource | Blob): Promise<void>;
  close(): Promise<void>;
  abort?(): Promise<void>;
};
type FileHandle = { kind?: string; getFile(): Promise<File>; createWritable(options?: { keepExistingData?: boolean }): Promise<WritableFile> };
type DirectoryHandle = {
  getDirectoryHandle(name: string, options?: { create?: boolean }): Promise<DirectoryHandle>;
  getFileHandle(name: string, options?: { create?: boolean }): Promise<FileHandle>;
  removeEntry(name: string, options?: { recursive?: boolean }): Promise<void>;
  entries?(): AsyncIterableIterator<[string, FileHandle]>;
};

/** Names this page's recordings are writing now: never swept. */
const liveNames = new Set<string>();

async function spoolDirectory(): Promise<DirectoryHandle | null> {
  try {
    const storage = typeof navigator !== "undefined" ? navigator.storage : undefined;
    if (!storage || typeof storage.getDirectory !== "function") return null;
    const root = (await storage.getDirectory()) as unknown as DirectoryHandle;
    return await root.getDirectoryHandle(SPOOL_DIRECTORY, { create: true });
  } catch {
    // Not offered here (a private window may refuse it).
    return null;
  }
}

/** What earlier recordings left, half an hour on: their downloads were taken long ago. */
async function sweep(directory: DirectoryHandle, now = Date.now()): Promise<void> {
  if (typeof directory.entries !== "function") return;
  try {
    const stale: string[] = [];
    for await (const [name, handle] of directory.entries()) {
      if (liveNames.has(name)) continue;
      try {
        const file = await handle.getFile();
        if (now - file.lastModified >= KEEP_FINISHED_MS) stale.push(name);
      } catch {
        // Not a file of ours, or being written elsewhere: left alone.
      }
    }
    await Promise.all(stale.map((name) => directory.removeEntry(name).catch(() => undefined)));
  } catch {
    // Another tab's recording holds one: it is swept next time.
  }
}

/**
 * The bytes written to a file of the site's private storage, or kept as
 * Blob pieces where this browser offers no writable file there.
 */
export async function createByteSpool(name: string): Promise<ByteSpool> {
  const directory = await spoolDirectory();
  if (!directory) return createBlobSpool();
  await sweep(directory);
  const fileName = `${name}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}.bin`;
  let handle: FileHandle;
  let writable: WritableFile;
  try {
    handle = await directory.getFileHandle(fileName, { create: true });
    if (typeof handle.createWritable !== "function") throw new Error("no writable");
    writable = await handle.createWritable({ keepExistingData: false });
  } catch {
    await directory.removeEntry(fileName).catch(() => undefined);
    return createBlobSpool();
  }
  liveNames.add(fileName);
  let piece: Uint8Array[] = [];
  let pieceBytes = 0;
  let size = 0;
  let waiting = 0;
  let failure: Error | null = null;
  let chain: Promise<void> = Promise.resolve();
  const waiters: Array<() => void> = [];
  const wake = () => {
    if (waiting > PENDING_LIMIT_BYTES && !failure) return;
    for (const resolve of waiters.splice(0)) resolve();
  };
  const write = () => {
    if (pieceBytes === 0) return;
    const bytes = new Uint8Array(pieceBytes);
    let at = 0;
    for (const part of piece) {
      bytes.set(part, at);
      at += part.length;
    }
    piece = [];
    pieceBytes = 0;
    waiting += bytes.length;
    chain = chain.then(async () => {
      if (failure) return;
      try {
        await writable.write(bytes);
      } catch (error) {
        failure = error instanceof Error ? error : new Error(String(error));
      } finally {
        waiting -= bytes.length;
        wake();
      }
    });
  };
  return {
    append(data) {
      if (failure) throw failure;
      if (data.length === 0) return;
      piece.push(data);
      pieceBytes += data.length;
      size += data.length;
      if (pieceBytes >= PIECE_BYTES) write();
    },
    get size() {
      return size;
    },
    drained() {
      if (failure) return Promise.reject(failure);
      if (waiting <= PENDING_LIMIT_BYTES) return Promise.resolve();
      return new Promise<void>((resolve, reject) => {
        waiters.push(() => (failure ? reject(failure) : resolve()));
      });
    },
    async seal() {
      write();
      await chain;
      if (failure) throw failure;
      await writable.close();
      // Finished: a later recording sweeps it once its download is long taken.
      liveNames.delete(fileName);
      return handle.getFile();
    },
    async discard() {
      piece = [];
      pieceBytes = 0;
      await chain.catch(() => undefined);
      try {
        await (writable.abort ? writable.abort() : writable.close());
      } catch {
        // Already closed.
      }
      liveNames.delete(fileName);
      await directory.removeEntry(fileName).catch(() => undefined);
    },
  };
}
