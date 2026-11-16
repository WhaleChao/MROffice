/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A kept queue of unsaved PDF edits as its holder last had it, or the word
 * that the queue was finished (`kept: null`), at the count of changes `seq`.
 */
export type PdfWorkingRecoveryShadowRecord = Readonly<{
  /** `scope:queue`, as the mirror names the queue. */
  name: string;
  seq: number;
  at: number;
  kept: Readonly<{
    claim: Readonly<{ owner: string | null; at: number }>;
    data: string;
    /** The records mirrored beside the queue, by the name each is kept under. */
    with: Readonly<Record<string, string>>;
  }> | null;
}>;

/**
 * Where the queues are written a second time, by a store that has them on
 * disk as each write completes (lib/pdf/working-recovery-store, "the
 * shadow"). Every call may reject: the mirror then stands alone, as it did.
 */
export type PdfWorkingRecoveryShadow = {
  /**
   * Keeps the record unless one of the queue with as high a count is kept
   * already (a write that arrives late never replaces a later one). Answers
   * whether it was kept, and the count the shadow has for the queue now.
   */
  put(record: PdfWorkingRecoveryShadowRecord): Promise<{ stored: boolean; seq: number }>;
  get(name: string): Promise<PdfWorkingRecoveryShadowRecord | null>;
  all(): Promise<PdfWorkingRecoveryShadowRecord[]>;
  remove(name: string): Promise<void>;
};

const DB_NAME = "mrnote-pdf-working-recovery";
const DB_VERSION = 1;
const STORE_NAME = "queues";
/** A database that does not open in this long (an upgrade blocked by another
 * tab, a profile on a stalled disk) is done without for this attempt. */
const OPEN_WITHIN_MS = 2_000;

function isRecord(value: unknown): value is PdfWorkingRecoveryShadowRecord {
  if (!value || typeof value !== "object") return false;
  const record = value as Record<string, unknown>;
  if (typeof record.name !== "string" || !Number.isSafeInteger(record.seq) || (record.seq as number) < 0 ||
    typeof record.at !== "number") return false;
  if (record.kept === null) return true;
  const kept = record.kept as Record<string, unknown> | undefined;
  const claim = kept?.claim as Record<string, unknown> | undefined;
  return Boolean(kept && typeof kept.data === "string" && claim &&
    (claim.owner === null || typeof claim.owner === "string") && typeof claim.at === "number" &&
    kept.with && typeof kept.with === "object" &&
    Object.values(kept.with as Record<string, unknown>).every((entry) => typeof entry === "string"));
}

/** One transaction on a store; `work` makes its requests and answers what to
 * read off them once all of it is written. */
export type PdfRecoveryIdbRun = <T,>(mode: IDBTransactionMode, work: (store: IDBObjectStore) => () => T) => Promise<T>;

/**
 * A store of this browser's IndexedDB whose writes are on disk when they are
 * reported done (records keyed by their `name`); null where there is none.
 * The queues' shadow below, and the typing drafts' (native-text-typing-drafts).
 */
export function openPdfRecoveryIdbStore(dbName: string, storeName: string): PdfRecoveryIdbRun | null {
  if (typeof indexedDB === "undefined") return null;
  let opening: Promise<IDBDatabase> | null = null;
  const open = (): Promise<IDBDatabase> => {
    if (opening) return opening;
    const attempt = new Promise<IDBDatabase>((resolve, reject) => {
      let settled = false;
      const fail = (error: unknown) => {
        if (settled) return;
        settled = true;
        if (opening === attempt) opening = null;
        reject(error instanceof Error ? error : new Error("The PDF recovery shadow did not open."));
      };
      const timer = setTimeout(() => fail(null), OPEN_WITHIN_MS);
      try {
        const request = indexedDB.open(dbName, DB_VERSION);
        request.onupgradeneeded = () => {
          if (!request.result.objectStoreNames.contains(storeName)) {
            request.result.createObjectStore(storeName, { keyPath: "name" });
          }
        };
        request.onsuccess = () => {
          clearTimeout(timer);
          const database = request.result;
          if (settled) { database.close(); return; }
          settled = true;
          const forget = () => { if (opening === attempt) opening = null; };
          database.onversionchange = () => { forget(); database.close(); };
          database.onclose = forget;
          resolve(database);
        };
        request.onerror = () => { clearTimeout(timer); fail(request.error); };
        request.onblocked = () => { clearTimeout(timer); fail(null); };
      } catch (error) {
        clearTimeout(timer);
        fail(error);
      }
    });
    opening = attempt;
    return attempt;
  };
  return async <T,>(mode: IDBTransactionMode, work: (store: IDBObjectStore) => () => T): Promise<T> => {
    const database = await open();
    return new Promise<T>((resolve, reject) => {
      // "strict": the write is on disk when it is reported done, which is
      // all this store is for.
      const transaction = database.transaction(storeName, mode, mode === "readwrite" ? { durability: "strict" } : undefined);
      const result = work(transaction.objectStore(storeName));
      transaction.oncomplete = () => resolve(result());
      transaction.onerror = () => reject(transaction.error ?? new Error("The PDF recovery shadow refused the request."));
      transaction.onabort = () => reject(transaction.error ?? new Error("The PDF recovery shadow gave the write up."));
    });
  };
}

/** The shadow in this browser's IndexedDB; null where there is none. */
export function createIndexedDbPdfWorkingRecoveryShadow(): PdfWorkingRecoveryShadow | null {
  const run = openPdfRecoveryIdbStore(DB_NAME, STORE_NAME);
  if (!run) return null;
  return {
    put(record) {
      return run("readwrite", (store) => {
        let outcome = { stored: false, seq: record.seq };
        const read = store.get(record.name);
        read.onsuccess = () => {
          const existing: unknown = read.result;
          if (isRecord(existing) && existing.seq >= record.seq) {
            outcome = { stored: false, seq: existing.seq };
            return;
          }
          store.put(record);
          outcome = { stored: true, seq: record.seq };
        };
        return () => outcome;
      });
    },
    async get(name) {
      const found: unknown = await run("readonly", (store) => {
        const request = store.get(name);
        return () => request.result as unknown;
      });
      return isRecord(found) ? found : null;
    },
    async all() {
      const found = await run("readonly", (store) => {
        const request = store.getAll();
        return () => request.result as unknown[];
      });
      return found.filter(isRecord);
    },
    async remove(name) {
      await run("readwrite", (store) => {
        store.delete(name);
        return () => undefined;
      });
    },
  };
}
