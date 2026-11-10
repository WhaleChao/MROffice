/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Large strings a page's local recovery draft holds, kept out of
 * sessionStorage.
 *
 * A native PDF edit that installs a font (the font extension) carries the
 * font program in its journal record, base64, often more than a megabyte. The
 * recovery draft is written to sessionStorage before the edit may be sent, and
 * sessionStorage holds a few megabytes per origin: such a draft did not fit,
 * the write failed, and the edit was refused before it reached the server.
 * A draft that does not fit whole now stores a placeholder for every string
 * of 64K characters or more and keeps the string itself in IndexedDB; a
 * restored draft is whole again only when every placeholder resolves, and is
 * dropped otherwise. A draft that fits is written whole, as before.
 */

const DB_NAME = "mrnote-page-draft-blobs";
const DB_VERSION = 1;
const STORE_NAME = "blobs";
const MARKER = "$mraiPageDraftBlob";
const MAX_DEPTH = 64;
const BLOB_MAX_AGE_MS = 24 * 60 * 60 * 1_000;
/** A kept string is stored again (its age refreshed) at most this often. */
const BLOB_REFRESH_MS = 60 * 60 * 1_000;

export const PAGE_DRAFT_BLOB_MIN_CHARS = 64 * 1024;
/** How often a tab sweeps out strings older than BLOB_MAX_AGE_MS. */
const SWEEP_INTERVAL_MS = 10 * 60 * 1_000;
let lastSweepAt = Number.NEGATIVE_INFINITY;

interface StoredBlob {
  readonly key: string;
  readonly pageId: string;
  readonly value: string;
  readonly savedAt: number;
}

export interface PageDraftBlob {
  readonly key: string;
  readonly value: string;
}

let databasePromise: Promise<IDBDatabase | null> | null = null;

function openDatabase(): Promise<IDBDatabase | null> {
  if (databasePromise) return databasePromise;
  if (typeof indexedDB === "undefined") {
    databasePromise = Promise.resolve(null);
    return databasePromise;
  }
  databasePromise = new Promise((resolve) => {
    let settled = false;
    const finish = (database: IDBDatabase | null) => {
      if (settled) return;
      settled = true;
      if (!database) databasePromise = null;
      resolve(database);
    };
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME, { keyPath: "key" });
        }
      };
      request.onsuccess = () => finish(request.result);
      request.onerror = () => finish(null);
      request.onblocked = () => finish(null);
    } catch {
      finish(null);
    }
  });
  return databasePromise;
}

function isPlaceholder(value: unknown): value is Readonly<{ [MARKER]: string; chars: number }> {
  return Boolean(
    value && typeof value === "object" && !Array.isArray(value) &&
      typeof (value as Record<string, unknown>)[MARKER] === "string" &&
      typeof (value as Record<string, unknown>).chars === "number" &&
      Object.keys(value).length === 2,
  );
}




const recentKeys = new Map<string, string>();
const RECENT_KEYS = 16;

function contentKey(pageId: string, value: string): string {
  // The same strings come back with every draft of the page (each edit
  // writes one): their keys are remembered, a lookup the engine answers from
  // the string's own cached hash. Hashing them again read megabytes of a
  
  const remembered = recentKeys.get(value);
  if (remembered !== undefined && remembered.startsWith(`${pageId}:`)) return remembered;
  const key = hashedContentKey(pageId, value);
  recentKeys.delete(value);
  recentKeys.set(value, key);
  if (recentKeys.size > RECENT_KEYS) recentKeys.delete(recentKeys.keys().next().value as string);
  return key;
}

function hashedContentKey(pageId: string, value: string): string {
  let first = 0x811c9dc5;
  let second = 0x01000193 ^ value.length;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193) >>> 0;
    second = Math.imul(second ^ (code + index), 0x01000193) >>> 0;
  }
  return `${pageId}:${value.length}:${first.toString(36)}${second.toString(36)}`;
}

/** The blob holding `value` for a draft of `pageId`, and the placeholder a
 * draft writes in its place (a draft can name a string kept already, without
 * carrying it: page-causal-journal.ts). */
export function pageDraftBlobFor(pageId: string, value: string): PageDraftBlob {
  return { key: contentKey(pageId, value), value };
}
export function pageDraftBlobPlaceholder(blob: PageDraftBlob): Readonly<{ [MARKER]: string; chars: number }> {
  return { [MARKER]: blob.key, chars: blob.value.length };
}

// Frozen objects (a kept block's JSON: an edit keeps every block it did not
// touch, the very same object from one draft to the next) found holding no
// string long enough to take out, by the depth they were walked from; a walk
// from there or deeper gives them back unchanged. Every draft walked every
// block of a long document again (a 296-page book: ~0.3 s of each pause).
const withoutBlobs = new WeakMap<object, number>();

/** A copy of `value` with every string of PAGE_DRAFT_BLOB_MIN_CHARS or more
 * replaced by a placeholder, and the strings it took out. */
export function externalizePageDraftBlobs<T>(
  value: T,
  pageId: string,
): Readonly<{ value: T; blobs: readonly PageDraftBlob[] }> {
  const blobs: PageDraftBlob[] = [];
  const walk = (entry: unknown, depth: number): unknown => {
    if (typeof entry === "string") {
      if (entry.length < PAGE_DRAFT_BLOB_MIN_CHARS) return entry;
      const key = contentKey(pageId, entry);
      if (!blobs.some((blob) => blob.key === key)) blobs.push({ key, value: entry });
      return { [MARKER]: key, chars: entry.length };
    }
    if (!entry || typeof entry !== "object" || depth > MAX_DEPTH) return entry;
    const frozen = Object.isFrozen(entry);
    if (frozen) {
      const cleanFrom = withoutBlobs.get(entry);
      if (cleanFrom !== undefined && depth >= cleanFrom) return entry;
    }
    if (Array.isArray(entry)) {
      let changed = false;
      const next = entry.map((item) => {
        const mapped = walk(item, depth + 1);
        if (mapped !== item) changed = true;
        return mapped;
      });
      if (!changed && frozen) withoutBlobs.set(entry, depth);
      return changed ? next : entry;
    }
    let changed = false;
    const next: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(entry as Record<string, unknown>)) {
      const mapped = walk(item, depth + 1);
      if (mapped !== item) changed = true;
      next[key] = mapped;
    }
    if (!changed && frozen) withoutBlobs.set(entry, depth);
    return changed ? next : entry;
  };
  return Object.freeze({ value: walk(value, 0) as T, blobs: Object.freeze(blobs) });
}

/** Whether a restored draft still names strings kept in IndexedDB. */
export function pageDraftHasBlobs(value: unknown): boolean {
  const walk = (entry: unknown, depth: number): boolean => {
    if (!entry || typeof entry !== "object" || depth > MAX_DEPTH) return false;
    if (isPlaceholder(entry)) return true;
    return Array.isArray(entry)
      ? entry.some((item) => walk(item, depth + 1))
      : Object.values(entry as Record<string, unknown>).some((item) => walk(item, depth + 1));
  };
  return walk(value, 0);
}

/** Keep the strings a draft took out; true only when every one is stored. */
export async function storePageDraftBlobs(
  pageId: string,
  blobs: readonly PageDraftBlob[],
): Promise<boolean> {
  if (blobs.length === 0) return true;
  const database = await openDatabase();
  if (!database) return false;
  const savedAt = Date.now();
  return new Promise((resolve) => {
    try {
      const transaction = database.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      for (const blob of blobs) {
        // An entry already holding the string is only refreshed in age, and
        // not again within BLOB_REFRESH_MS: refreshing is writing the whole
        // record, megabytes for a large page's state, every draft.
        const existing = store.get(blob.key);
        existing.onsuccess = () => {
          const record = existing.result as StoredBlob | undefined;
          const held = record && record.pageId === pageId && record.value.length === blob.value.length;
          if (held && typeof record.savedAt === "number" && savedAt - record.savedAt >= 0 &&
              savedAt - record.savedAt < BLOB_REFRESH_MS) return;
          store.put(held
            ? { ...record, savedAt }
            : { key: blob.key, pageId, value: blob.value, savedAt } satisfies StoredBlob);
        };
      }
      // Strings older than a draft may live are no one's any more. The sweep
      // reads every record (values included), so it runs once in a while,
      // not with every draft: each edit of a long document paid for it.
      const sweep = savedAt - lastSweepAt >= SWEEP_INTERVAL_MS ? store.openCursor() : null;
      if (sweep) lastSweepAt = savedAt;
      if (sweep) sweep.onsuccess = () => {
        const cursor = sweep.result;
        if (!cursor) return;
        const record = cursor.value as StoredBlob;
        if (typeof record?.savedAt !== "number" || savedAt - record.savedAt > BLOB_MAX_AGE_MS) cursor.delete();
        cursor.continue();
      };
      transaction.oncomplete = () => resolve(true);
      transaction.onerror = () => resolve(false);
      transaction.onabort = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

/** The draft with every placeholder replaced by its string, or null when one
 * of them is missing (the tab's IndexedDB was cleared, or another page's). */
export async function hydratePageDraftBlobs<T>(pageId: string, value: T): Promise<T | null> {
  const keys = new Set<string>();
  const collect = (entry: unknown, depth: number) => {
    if (!entry || typeof entry !== "object" || depth > MAX_DEPTH) return;
    if (isPlaceholder(entry)) {
      keys.add(entry[MARKER]);
      return;
    }
    for (const item of Array.isArray(entry) ? entry : Object.values(entry as Record<string, unknown>)) {
      collect(item, depth + 1);
    }
  };
  collect(value, 0);
  if (keys.size === 0) return value;
  const database = await openDatabase();
  if (!database) return null;
  const found = await new Promise<Map<string, string> | null>((resolve) => {
    try {
      const store = database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME);
      const values = new Map<string, string>();
      let pending = keys.size;
      let failed = false;
      for (const key of keys) {
        const request = store.get(key);
        request.onsuccess = () => {
          const record = request.result as StoredBlob | undefined;
          if (record && record.pageId === pageId && typeof record.value === "string") values.set(key, record.value);
          else failed = true;
          pending -= 1;
          if (pending === 0) resolve(failed ? null : values);
        };
        request.onerror = () => {
          failed = true;
          pending -= 1;
          if (pending === 0) resolve(null);
        };
      }
    } catch {
      resolve(null);
    }
  });
  if (!found) return null;
  let intact = true;
  const walk = (entry: unknown, depth: number): unknown => {
    if (!entry || typeof entry !== "object" || depth > MAX_DEPTH) return entry;
    if (isPlaceholder(entry)) {
      const restored = found.get(entry[MARKER]);
      if (restored === undefined || restored.length !== entry.chars) intact = false;
      return restored ?? entry;
    }
    if (Array.isArray(entry)) return entry.map((item) => walk(item, depth + 1));
    return Object.fromEntries(
      Object.entries(entry as Record<string, unknown>).map(([key, item]) => [key, walk(item, depth + 1)]),
    );
  };
  const hydrated = walk(value, 0) as T;
  return intact ? hydrated : null;
}
