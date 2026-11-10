/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The sealing keys of this browser's local copies (saved pages, PDF source
 * bytes), one per signed-in account and purpose, kept across tabs and
 * restarts so a page reopened in a new tab or the next morning opens from
 * the copy this browser already holds instead of waiting for the network.
 *
 * - A key is a non-extractable AES-GCM CryptoKey in IndexedDB: page script
 *   can seal and open with it but never read its bytes, and nothing secret
 *   is ever written to cookies, localStorage or the records themselves.
 * - Keys are partitioned by the login principal cookie (an HMAC of the user
 *   id, `mrnote_page_cache_principal`). Asking for one account's key deletes
 *   every other account's keys in the same transaction, so after an account
 *   switch the previous account's copies cannot be opened again even if
 *   their records are still on disk. Logout clears every key
 *   (`clearPrincipalCacheKeys`), which makes all sealed copies unreadable at
 *   once.
 * - A key is retired after KEY_MAX_AGE_MS; copies sealed under it then fail
 *   to open and are simply read from the network again.
 *
 * Every operation is best-effort and never throws: storage that is missing,
 * blocked or cleared answers `null`, and callers read the network as before.
 */

const DB_NAME = "mrnote-principal-cache-keys-v1";
const STORE_NAME = "keys";
const DB_VERSION = 1;
const KEY_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1_000;
const MAX_FUTURE_SKEW_MS = 5 * 60 * 1_000;
const PRINCIPAL = /^[0-9a-f]{64}$/u;

export type PrincipalCacheKeyPurpose = "saved-page" | "pdf-source";

interface StoredKeyRecord {
  id: string;
  principal: string;
  purpose: PrincipalCacheKeyPurpose;
  createdAt: number;
  key: CryptoKey;
}

let databasePromise: Promise<IDBDatabase | null> | null = null;
let generation = 0;
const resolvedKeys = new Map<string, CryptoKey>();
let lastClear: Promise<void> = Promise.resolve();

function recordId(principal: string, purpose: PrincipalCacheKeyPurpose): string {
  return `${principal}:${purpose}`;
}

function isAesGcmKey(value: unknown): value is CryptoKey {
  if (!value || typeof value !== "object") return false;
  const key = value as Partial<CryptoKey>;
  return (
    key.type === "secret" &&
    (key.algorithm as { name?: unknown } | undefined)?.name === "AES-GCM" &&
    Array.isArray(key.usages) &&
    key.usages.includes("encrypt") &&
    key.usages.includes("decrypt")
  );
}

function validRecord(
  value: unknown,
  principal: string,
  purpose: PrincipalCacheKeyPurpose,
): value is StoredKeyRecord {
  if (!value || typeof value !== "object") return false;
  const record = value as Partial<StoredKeyRecord>;
  const now = Date.now();
  return (
    record.id === recordId(principal, purpose) &&
    record.principal === principal &&
    record.purpose === purpose &&
    typeof record.createdAt === "number" &&
    Number.isFinite(record.createdAt) &&
    record.createdAt <= now + MAX_FUTURE_SKEW_MS &&
    now - record.createdAt < KEY_MAX_AGE_MS &&
    isAesGcmKey(record.key)
  );
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (databasePromise) return databasePromise;
  if (typeof indexedDB === "undefined") {
    databasePromise = Promise.resolve(null);
    return databasePromise;
  }
  const opening = new Promise<IDBDatabase | null>((resolve) => {
    let settled = false;
    const finish = (database: IDBDatabase | null) => {
      if (settled) return;
      settled = true;
      resolve(database);
    };
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME, { keyPath: "id" });
        }
      };
      request.onsuccess = () => {
        const database = request.result;
        // Another tab upgrading or deleting the store: let it, and open
        // again on the next use.
        try {
          database.onversionchange = () => {
            database.close();
            if (databasePromise === opening) databasePromise = null;
          };
        } catch {
          // A store without the event still works for this tab.
        }
        finish(database);
      };
      request.onerror = () => finish(null);
      request.onblocked = () => finish(null);
    } catch {
      finish(null);
    }
  });
  databasePromise = opening;
  // A failed open is not remembered: storage may come back (a private
  // window's quota, a store another tab was deleting).
  void opening.then((database) => {
    if (!database && databasePromise === opening) databasePromise = null;
  });
  return opening;
}

/**
 * Read (and with `create`, make) the key, and drop every other account's
 * keys, in one read-write transaction. Two tabs creating a key at the same
 * moment agree on whichever committed first.
 */
function keyTransaction(
  database: IDBDatabase,
  principal: string,
  purpose: PrincipalCacheKeyPurpose,
  candidate: CryptoKey | null,
): Promise<CryptoKey | null> {
  return new Promise((resolve) => {
    let result: CryptoKey | null = null;
    try {
      const transaction = database.transaction(STORE_NAME, "readwrite");
      const store = transaction.objectStore(STORE_NAME);
      const ids = store.getAllKeys();
      ids.onsuccess = () => {
        for (const id of ids.result as IDBValidKey[]) {
          if (typeof id !== "string" || !id.startsWith(`${principal}:`)) store.delete(id);
        }
      };
      const request = store.get(recordId(principal, purpose));
      request.onsuccess = () => {
        if (validRecord(request.result, principal, purpose)) {
          result = request.result.key;
          return;
        }
        if (request.result !== undefined) store.delete(recordId(principal, purpose));
        if (!candidate) return;
        const record: StoredKeyRecord = {
          id: recordId(principal, purpose),
          principal,
          purpose,
          createdAt: Date.now(),
          key: candidate,
        };
        store.put(record);
        result = candidate;
      };
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => resolve(null);
      transaction.onabort = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

/**
 * This account's key for `purpose`, or null (no principal, no storage, no
 * WebCrypto, or no key yet and `create` not asked). Never throws.
 */
export async function principalCacheKey(
  principal: string | null,
  purpose: PrincipalCacheKeyPurpose,
  options: { create?: boolean } = {},
): Promise<CryptoKey | null> {
  if (!principal || !PRINCIPAL.test(principal)) return null;
  const id = recordId(principal, purpose);
  const known = resolvedKeys.get(id);
  if (known) return known;
  const startedIn = generation;
  try {
    const subtle = globalThis.crypto?.subtle;
    if (!subtle) return null;
    const database = await openDatabase();
    if (!database || startedIn !== generation) return null;
    const candidate = options.create
      ? ((await subtle.generateKey({ name: "AES-GCM", length: 256 }, false, [
          "encrypt",
          "decrypt",
        ])) as CryptoKey)
      : null;
    if (startedIn !== generation) return null;
    const key = await keyTransaction(database, principal, purpose, candidate);
    // A logout while the transaction ran: the key it returned was cleared.
    if (!key || startedIn !== generation) return null;
    resolvedKeys.set(id, key);
    return key;
  } catch {
    return null;
  }
}

/** Monotonic counter that changes whenever every key was forgotten. */
export function principalCacheKeyGeneration(): number {
  return generation;
}

/**
 * Forget every key (logout, session expiry). The returned promise settles
 * once the store is cleared or storage proved unavailable; it never rejects.
 */
export function clearPrincipalCacheKeys(): Promise<void> {
  generation += 1;
  resolvedKeys.clear();
  lastClear = (async () => {
    const database = await openDatabase();
    if (!database) return;
    await new Promise<void>((resolve) => {
      try {
        const transaction = database.transaction(STORE_NAME, "readwrite");
        transaction.objectStore(STORE_NAME).clear();
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => resolve();
        transaction.onabort = () => resolve();
      } catch {
        resolve();
      }
    });
  })().catch(() => undefined);
  return lastClear;
}

/** The last `clearPrincipalCacheKeys` run, for a logout that waits on it. */
export function principalCacheKeysCleared(): Promise<void> {
  return lastClear;
}

export function resetPrincipalCacheKeysForTests(): void {
  databasePromise = null;
  generation = 0;
  resolvedKeys.clear();
  lastClear = Promise.resolve();
}
