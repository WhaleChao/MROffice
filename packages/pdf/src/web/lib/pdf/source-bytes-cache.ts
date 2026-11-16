/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readCookie } from "@mrnote/core/web/lib/cookie.ts";
import {
  clearPrincipalCacheKeys,
  principalCacheKey,
  principalCacheKeyGeneration,
} from "@mrnote/core/web/lib/principal-cache-key.ts";

const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
// Keys of the earlier tab-scoped store; only ever removed now.
const LEGACY_KEY_STORAGE_PREFIX = "mrnote:pdf-source-bytes-key:v1:";
const DB_NAME = "mrnote-pdf-source-bytes-v1";
const STORE_NAME = "sources";
const DB_VERSION = 1;
const RECORD_VERSION = 1;
const MAX_ENTRY_BYTES = 48 * 1024 * 1024;
const MAX_TOTAL_BYTES = 160 * 1024 * 1024;
const MAX_RECORDS = 8;
const MAX_AGE_MS = 14 * 24 * 60 * 60 * 1_000;
const WAIT_MS = 400;
const SHA256 = /^[0-9a-f]{64}$/u;
const TIMEOUT = Symbol("pdf-source-bytes-timeout");

interface StoredSourceRecord {
  cacheKey: string;
  principal: string;
  sha256: string;
  savedAt: number;
  bytes: number;
  iv: ArrayBuffer;
  ciphertext: ArrayBuffer;
  version: number;
}

let databasePromise: Promise<IDBDatabase | null> | null = null;
let sessionEpoch = 0;

function currentPrincipal(): string | null {
  if (typeof document === "undefined") return null;
  const value = readCookie(PRINCIPAL_COOKIE);
  return value && /^[0-9a-f]{64}$/u.test(value) ? value : null;
}

/** The login principal the kept bytes are partitioned by (null: none):
 * bytes held in memory by digest (lib/pdf/source-prefetch) are shared only
 * under the principal that read them, as the records here are. */
export function pdfSourceBytesPrincipal(): string | null {
  return currentPrincipal();
}

function additionalData(principal: string, sha256: string): ArrayBuffer {
  return Uint8Array.from(
    new TextEncoder().encode(`${RECORD_VERSION}\u0000${principal}\u0000${sha256}`),
  ).buffer;
}

function recordKey(principal: string, sha256: string): string {
  return `${principal}:${sha256}`;
}

async function bounded<T>(promise: Promise<T>): Promise<T | typeof TIMEOUT> {
  let timer: ReturnType<typeof globalThis.setTimeout> | null = null;
  const timeout = new Promise<typeof TIMEOUT>((resolve) => {
    timer = globalThis.setTimeout(() => resolve(TIMEOUT), WAIT_MS);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timer !== null) globalThis.clearTimeout(timer);
  }
}

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
          database.createObjectStore(STORE_NAME, { keyPath: "cacheKey" });
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

function requestResult<T>(request: IDBRequest<T>): Promise<T | null> {
  return new Promise((resolve) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
  });
}

function validRecord(
  record: unknown, principal: string, sha256: string, now: number,
): record is StoredSourceRecord {
  const value = record as Partial<StoredSourceRecord> | null;
  return Boolean(
    value &&
      value.version === RECORD_VERSION &&
      value.cacheKey === recordKey(principal, sha256) &&
      value.principal === principal &&
      value.sha256 === sha256 &&
      typeof value.savedAt === "number" &&
      now - value.savedAt >= 0 &&
      now - value.savedAt <= MAX_AGE_MS &&
      isArrayBuffer(value.iv) &&
      isArrayBuffer(value.ciphertext),
  );
}

/** Realm-independent: a structured clone may come from another realm. */
function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === "[object ArrayBuffer]";
}

/**
 * The cached bytes claimed to be `sha256` for this account, or null.
 * The caller hashes them before use.
 */
export async function readCachedPdfSourceBytes(sha256: string | null | undefined): Promise<Uint8Array | null> {
  const digest = typeof sha256 === "string" ? sha256.trim().toLowerCase() : "";
  const principal = currentPrincipal();
  if (!SHA256.test(digest) || !principal) return null;
  const epoch = sessionEpoch;
  const keyGeneration = principalCacheKeyGeneration();
  try {
    const result = await bounded((async () => {
      // Asked first: this also drops another account's keys (an account
      // switch), whatever this read finds.
      const key = await principalCacheKey(principal, "pdf-source");
      const database = await openDatabase();
      if (!database || !key) return null;
      const record = await requestResult(
        database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).get(recordKey(principal, digest)),
      );
      if (!validRecord(record, principal, digest, Date.now())) return null;
      const plain = await globalThis.crypto.subtle.decrypt(
        { name: "AES-GCM", iv: record.iv, additionalData: additionalData(principal, digest) },
        key, record.ciphertext,
      );
      return new Uint8Array(plain);
    })());
    if (result === TIMEOUT || !result || epoch !== sessionEpoch ||
        keyGeneration !== principalCacheKeyGeneration() || currentPrincipal() !== principal) return null;
    return result.byteLength > 0 && result.byteLength <= MAX_ENTRY_BYTES ? result : null;
  } catch {
    return null;
  }
}

/** How long the digests kept here may take to list before the caller goes
 * on without them (it then asks as a reader that keeps nothing would not). */
const LIST_WAIT_MS = 150;

/**
 * The digests of the files kept here for this account (their keys alone; no
 * record is read or opened), or null when they could not be listed in time.
 * A reader names them when it asks for a file's bytes, so the server can
 * answer "you hold these" instead of sending them (lib/pdf/source-prefetch).
 * A listed record may since have expired: its read then finds nothing.
 */
export async function listCachedPdfSourceDigests(): Promise<string[] | null> {
  const principal = currentPrincipal();
  if (!principal) return [];
  const epoch = sessionEpoch;
  try {
    let timer: ReturnType<typeof globalThis.setTimeout> | null = null;
    const timeout = new Promise<typeof TIMEOUT>((resolve) => {
      timer = globalThis.setTimeout(() => resolve(TIMEOUT), LIST_WAIT_MS);
    });
    const listed = (async () => {
      const database = await openDatabase();
      if (!database) return [] as string[];
      const keys = await requestResult(
        database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).getAllKeys(),
      );
      const prefix = `${principal}:`;
      const digests: string[] = [];
      for (const key of keys ?? []) {
        if (typeof key !== "string" || !key.startsWith(prefix)) continue;
        const digest = key.slice(prefix.length);
        if (SHA256.test(digest)) digests.push(digest);
      }
      return digests.slice(0, MAX_RECORDS * 2);
    })();
    let result: string[] | typeof TIMEOUT;
    try {
      result = await Promise.race([listed, timeout]);
    } finally {
      if (timer !== null) globalThis.clearTimeout(timer);
    }
    if (result === TIMEOUT || epoch !== sessionEpoch || currentPrincipal() !== principal) return null;
    return result;
  } catch {
    return null;
  }
}

/** Keep `bytes` (already verified to be `sha256`) for this account. */
export async function storeCachedPdfSourceBytes(sha256: string | null | undefined, bytes: Uint8Array): Promise<void> {
  const digest = typeof sha256 === "string" ? sha256.trim().toLowerCase() : "";
  const principal = currentPrincipal();
  if (!SHA256.test(digest) || !principal || bytes.byteLength === 0 || bytes.byteLength > MAX_ENTRY_BYTES) return;
  const epoch = sessionEpoch;
  const keyGeneration = principalCacheKeyGeneration();
  const current = () => epoch === sessionEpoch && keyGeneration === principalCacheKeyGeneration() &&
    currentPrincipal() === principal;
  const owned = Uint8Array.from(bytes);
  try {
    const key = await principalCacheKey(principal, "pdf-source", { create: true });
    const database = await openDatabase();
    if (!key || !database || !current()) return;
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
    const ciphertext = await globalThis.crypto.subtle.encrypt(
      { name: "AES-GCM", iv, additionalData: additionalData(principal, digest) }, key, owned,
    );
    if (!current()) return;
    const record: StoredSourceRecord = {
      cacheKey: recordKey(principal, digest), principal, sha256: digest,
      savedAt: Date.now(), bytes: owned.byteLength, iv: iv.buffer, ciphertext, version: RECORD_VERSION,
    };
    const store = database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME);
    await requestResult(store.put(record));
    await pruneRecords(database, principal);
  } catch {
    // Optional: the next open downloads the source as before.
  } finally {
    owned.fill(0);
  }
}

/** Bounds the store, and drops whatever another account left in it. */
async function pruneRecords(database: IDBDatabase, principal: string): Promise<void> {
  const all = await requestResult(
    database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).getAll(),
  ) as StoredSourceRecord[] | null;
  if (!all) return;
  const now = Date.now();
  const keep = all
    .filter((record) => record && record.principal === principal &&
      typeof record.savedAt === "number" && now - record.savedAt <= MAX_AGE_MS)
    .sort((left, right) => right.savedAt - left.savedAt);
  const kept = new Set<string>();
  let total = 0;
  for (const record of keep) {
    if (kept.size >= MAX_RECORDS || total + record.bytes > MAX_TOTAL_BYTES) break;
    kept.add(record.cacheKey);
    total += record.bytes;
  }
  const store = database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME);
  for (const record of all) {
    if (record && !kept.has(record.cacheKey)) store.delete(record.cacheKey);
  }
}

/** Logout: forget the keys (the records become unreadable) and drop them. */
export function clearPdfSourceBytesCacheSession(): void {
  sessionEpoch += 1;
  void clearPrincipalCacheKeys();
  if (typeof window !== "undefined") {
    try {
      const keys: string[] = [];
      for (let index = 0; index < window.sessionStorage.length; index += 1) {
        const key = window.sessionStorage.key(index);
        if (key?.startsWith(LEGACY_KEY_STORAGE_PREFIX)) keys.push(key);
      }
      for (const key of keys) window.sessionStorage.removeItem(key);
    } catch {
      // Without the key no record can be read back anyway.
    }
  }
  void openDatabase().then((database) => {
    if (!database) return;
    try {
      database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).clear();
    } catch {
      // Best effort.
    }
  });
}

export function resetPdfSourceBytesCacheForTests(): void {
  databasePromise = null;
  sessionEpoch = 0;
}
