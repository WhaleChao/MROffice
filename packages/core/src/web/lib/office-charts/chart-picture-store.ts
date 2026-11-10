/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The browser's persistent copy of chart pictures it was already given.
 *
 * Charts are drawn on the server (engine `charts`); a picture is a pure
 * function of its request, so this store keeps answers by the request's
 * hash: a chart seen once draws again on reopen before the server answers,
 * and stays visible offline. It is a per-viewer convenience, never authority:
 * an editor still asks the server and replaces a stored picture with the
 * answer, and every failure here (private window, quota, blocked storage) is
 * a miss.
 *
 * Records belong to the signed-in principal (the page cache's opaque
 * principal cookie, or "guest"), are capped by total size and age, and the
 * whole store is deleted at sign-out (`lib/auth-state.ts`).
 */
import { readCookie } from "../cookie";
import type { ChartRenderPicture } from "../engine-host/charts-types";

const DB_NAME = "mrai-chart-pictures-v1";
const STORE = "pictures";
const AT_INDEX = "at";
const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
/** About 20 MB of SVG text (DESIGN-WAVE2.md 3.2, decision 3). */
const MAX_TOTAL_CHARS = 20_000_000;
const MAX_RECORD_CHARS = 2_000_000;
const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
const WAIT_MS = 400;
/** Every this many writes, the oldest records past the cap are dropped. */
const TRIM_EVERY = 25;

interface StoredPicture {
  key: string;
  at: number;
  chars: number;
  picture: ChartRenderPicture;
}

let database: Promise<IDBDatabase | null> | null = null;
let writes = 0;

function principal(): string {
  try {
    const value = typeof document === "undefined" ? null : readCookie(PRINCIPAL_COOKIE);
    return value && /^[0-9a-f]{64}$/u.test(value) ? value : "guest";
  } catch {
    return "guest";
  }
}

function open(): Promise<IDBDatabase | null> {
  if (database) return database;
  database = new Promise<IDBDatabase | null>((resolve) => {
    try {
      if (typeof indexedDB === "undefined") return resolve(null);
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        const store = request.result.createObjectStore(STORE, { keyPath: "key" });
        store.createIndex(AT_INDEX, "at");
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return database;
}

function bounded<T>(promise: Promise<T>, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    const timer = setTimeout(() => resolve(fallback), WAIT_MS);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      () => {
        clearTimeout(timer);
        resolve(fallback);
      },
    );
  });
}

function recordKey(hash: string): string {
  return `${principal()}:${hash}`;
}

/** The stored answer to the request whose hash is `hash`, or null. */
export function readStoredChartPicture(hash: string): Promise<ChartRenderPicture | null> {
  const read = async (): Promise<ChartRenderPicture | null> => {
    const db = await open();
    if (!db) return null;
    return new Promise((resolve) => {
      try {
        const request = db.transaction(STORE, "readonly").objectStore(STORE).get(recordKey(hash));
        request.onsuccess = () => {
          const record = request.result as StoredPicture | undefined;
          resolve(
            record && Date.now() - record.at < MAX_AGE_MS && record.picture?.ok === true && typeof record.picture.svg === "string"
              ? record.picture
              : null,
          );
        };
        request.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  };
  return bounded(read(), null);
}

/** Keep the answer to the request whose hash is `hash` (best effort). */
export function storeChartPicture(hash: string, picture: ChartRenderPicture): void {
  const chars = picture.svg.length;
  if (chars > MAX_RECORD_CHARS) return;
  void (async () => {
    const db = await open();
    if (!db) return;
    try {
      const transaction = db.transaction(STORE, "readwrite");
      transaction.objectStore(STORE).put({ key: recordKey(hash), at: Date.now(), chars, picture } satisfies StoredPicture);
    } catch {
      return;
    }
    writes += 1;
    if (writes % TRIM_EVERY === 0) trim(db);
  })().catch(() => {});
}

/** Drop records past the age limit, then the oldest past the size cap. */
function trim(db: IDBDatabase): void {
  try {
    const transaction = db.transaction(STORE, "readwrite");
    const records: { key: string; at: number; chars: number }[] = [];
    const cursor = transaction.objectStore(STORE).index(AT_INDEX).openCursor(null, "prev");
    cursor.onsuccess = () => {
      const current = cursor.result;
      if (current) {
        const value = current.value as StoredPicture;
        records.push({ key: value.key, at: value.at, chars: value.chars });
        current.continue();
        return;
      }
      // Newest first: keep until the cap, drop the rest and the expired.
      let total = 0;
      const store = transaction.objectStore(STORE);
      for (const record of records) {
        total += record.chars;
        if (total > MAX_TOTAL_CHARS || Date.now() - record.at > MAX_AGE_MS) store.delete(record.key);
      }
    };
  } catch {
    /* a later write trims again */
  }
}

/** Sign-out: no picture of this account's charts stays in the browser. */
export function clearChartPictureStore(): void {
  const pending = database;
  database = null;
  void (async () => {
    const db = await pending;
    db?.close();
    try {
      if (typeof indexedDB !== "undefined") indexedDB.deleteDatabase(DB_NAME);
    } catch {
      /* nothing stored, or storage blocked */
    }
  })();
}
