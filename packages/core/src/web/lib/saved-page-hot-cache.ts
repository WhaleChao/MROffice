/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readCookie } from "./cookie";
import { encodeJsonInSlices } from "./json-slices";
import type { PageItem } from "./notebook-sdk";
import {
  clearPrincipalCacheKeys,
  principalCacheKey,
  principalCacheKeyGeneration,
  principalCacheKeysCleared,
} from "./principal-cache-key";
import {
  expandCellStyleTable,
  hasCellStyleTable,
  packCellStyleTable,
} from "./spreadsheet/cell-style-table";
import { mountWorkspaceFor } from "./workspace-mounts";

const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
const WORKSPACE_COOKIE = "mrnote_workspace_id";
// Session-only bookkeeping of this tab: deletions it made (tombstones) and
// the order of its own writes. The legacy key prefix is only cleared.
const LEGACY_KEY_PREFIX = "mrnote:saved-page-hot-key:v1:";
const TOMBSTONE_PREFIX = "mrnote:saved-page-hot-blocked:v2:";
const LEGACY_TOMBSTONE_PREFIX = "mrnote:saved-page-hot-blocked:v1:";
const OPERATION_EPOCH_PREFIX = "mrnote:saved-page-hot-epoch:v2:";
const LEGACY_OPERATION_EPOCH_PREFIX = "mrnote:saved-page-hot-epoch:v1:";
const LEGACY_TAB_ID_KEY = "mrnote:saved-page-hot-tab:v1";
const DB_NAME = "mrnote-page-cache-v2";
const LEGACY_DB_NAME = "mrnote-saved-page-hot-cache-v1";
const PAGES_STORE = "pages";
const META_STORE = "meta";
const DB_VERSION = 1;
const RECORD_VERSION = 2;
// A long imported document is 6.6 MB with its checkpoint; a 62,791-cell
// workbook is 3.6 MB in its stored form (`cellStyleTable`), which is how a
// spreadsheet is kept here (27 MB inline).
const MAX_PAGE_BYTES = 24 * 1024 * 1024;
const MAX_TOTAL_BYTES = 160 * 1024 * 1024;
const MAX_RECORDS = 48;
const MAX_AGE_MS = 14 * 24 * 60 * 60 * 1_000;
const MAX_FUTURE_SKEW_MS = 5 * 60 * 1_000;
export const SAVED_PAGE_HOT_CACHE_WAIT_MS = 500;
/**
 * How long a window that is opening a page waits for the copy it kept, to
 * ask the server only what changed since. A document being mounted holds the
 * thread for longer than the preview's wait above, and the answers IndexedDB
 * had ready were given up on while they queued behind it: the reload then
 * downloaded the whole page it was holding.
 */
export const SAVED_PAGE_HOT_CACHE_OPEN_WAIT_MS = 2_500;
/** Keeping a copy is background work: it is not given up because the thread
 * was busy when the store answered (the copy of a page just opened was never
 * kept for exactly that reason). */
const WRITE_WAIT_MS = 15_000;

const SUPPORTED_PAGE_TYPES = new Set([
  "document",
  "presentation",
  "pdf",
  "code",
  "spreadsheet",
]);
const EDITOR_OWNED_PAGE_TYPES = new Set(["code", "spreadsheet"]);
const TIMEOUT = Symbol("saved-page-hot-cache-timeout");

interface StoredPageRecord {
  cacheKey: string;
  principal: string;
  workspaceId: string;
  notebookId: string;
  pageId: string;
  savedAt: number;
  expiresAt: number;
  bytes: number;
  contentVersion: number;
  iv: ArrayBuffer;
  ciphertext: ArrayBuffer;
  version: number;
}

/** What eviction and version ordering read, without the ciphertext. */
interface StoredPageMeta {
  cacheKey: string;
  savedAt: number;
  usedAt: number;
  expiresAt: number;
  bytes: number;
  contentVersion: number;
}

let databasePromise: Promise<IDBDatabase | null> | null = null;
const operationEpochs = new Map<string, number>();
const blockedKeys = new Set<string>();
const purgedFor = new Set<string>();
let cacheSessionEpoch = 0;
let lastWipe: Promise<void> = Promise.resolve();

export interface SavedPageHotCacheScope {
  readonly principal: string | null;
  readonly sessionEpoch: number;
}

/** Capture before HTTP work; a late receipt cannot borrow another account. */
export function captureSavedPageHotCacheScope(): SavedPageHotCacheScope {
  return { principal: currentPrincipal(), sessionEpoch: cacheSessionEpoch };
}

export function isSavedPageHotCacheScopeCurrent(
  scope: SavedPageHotCacheScope,
): boolean {
  return (
    scope.principal === currentPrincipal() &&
    scope.sessionEpoch === cacheSessionEpoch
  );
}

function storedOperationEpoch(key: string): number | null {
  if (typeof window === "undefined") return null;
  try {
    const value = Number(
      window.sessionStorage.getItem(`${OPERATION_EPOCH_PREFIX}${key}`),
    );
    return Number.isSafeInteger(value) && value > 0 ? value : null;
  } catch {
    return null;
  }
}

function currentOperationEpoch(key: string): number {
  return storedOperationEpoch(key) ?? operationEpochs.get(key) ?? 0;
}

function nextOperationEpoch(key: string): number {
  const current = currentOperationEpoch(key);
  const epoch = current >= Number.MAX_SAFE_INTEGER ? 1 : current + 1;
  operationEpochs.set(key, epoch);
  if (typeof window !== "undefined") {
    try {
      window.sessionStorage.setItem(
        `${OPERATION_EPOCH_PREFIX}${key}`,
        String(epoch),
      );
    } catch {
      // The in-memory epoch still orders operations in this live module.
    }
  }
  return epoch;
}

function currentPrincipal(): string | null {
  if (typeof document === "undefined") return null;
  const value = readCookie(PRINCIPAL_COOKIE);
  return value && /^[0-9a-f]{64}$/u.test(value) ? value : null;
}

/** The workspace a read of a page in this notebook is made in (lib/api
 * `resolveApiWorkspaceId`: a shared notebook's own, else the user's). */
function pageWorkspace(notebookId: string): string {
  if (typeof document === "undefined") return "";
  try {
    return mountWorkspaceFor(notebookId) ?? readCookie(WORKSPACE_COOKIE) ?? "";
  } catch {
    return "";
  }
}

function cacheKey(principal: string, pageId: string): string | null {
  if (typeof window === "undefined" || !pageId) return null;
  return `${principal}:${pageId}`;
}

function tombstoneStorageKey(key: string): string {
  return `${TOMBSTONE_PREFIX}${key}`;
}

function blockCacheKey(key: string): void {
  blockedKeys.add(key);
  if (typeof window === "undefined") return;
  try {
    window.sessionStorage.setItem(tombstoneStorageKey(key), "1");
  } catch {
    // The in-memory tombstone still protects this page for the live session.
  }
}

function unblockCacheKey(key: string): void {
  blockedKeys.delete(key);
  if (typeof window === "undefined") return;
  try {
    window.sessionStorage.removeItem(tombstoneStorageKey(key));
  } catch {
    // A retained tombstone only disables an optional future preview.
  }
}

function cacheKeyIsBlocked(key: string): boolean {
  if (blockedKeys.has(key)) return true;
  if (typeof window === "undefined") return false;
  try {
    const blocked =
      window.sessionStorage.getItem(tombstoneStorageKey(key)) === "1";
    if (blocked) blockedKeys.add(key);
    return blocked;
  } catch {
    return false;
  }
}

function additionalData(
  principal: string,
  workspaceId: string,
  pageId: string,
): ArrayBuffer {
  return Uint8Array.from(
    new TextEncoder().encode(
      `${RECORD_VERSION}\u0000${principal}\u0000${workspaceId}\u0000${pageId}`,
    ),
  ).buffer;
}

async function bounded<T>(
  promise: Promise<T>,
  waitMs = SAVED_PAGE_HOT_CACHE_WAIT_MS,
): Promise<T | typeof TIMEOUT> {
  let timer: ReturnType<typeof globalThis.setTimeout> | null = null;
  const timeout = new Promise<typeof TIMEOUT>((resolve) => {
    timer = globalThis.setTimeout(() => resolve(TIMEOUT), waitMs);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timer !== null) globalThis.clearTimeout(timer);
  }
}

let legacyDatabaseDropped = false;
function dropLegacyDatabase(): void {
  if (legacyDatabaseDropped) return;
  legacyDatabaseDropped = true;
  try {
    // Records of the tab-scoped store can never be opened again (their keys
    // lived in sessionStorage); free the space.
    indexedDB.deleteDatabase?.(LEGACY_DB_NAME);
  } catch {
    // Nothing kept, or storage blocked.
  }
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (databasePromise) return databasePromise;
  if (typeof indexedDB === "undefined") {
    databasePromise = Promise.resolve(null);
    return databasePromise;
  }
  dropLegacyDatabase();
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
        if (!database.objectStoreNames.contains(PAGES_STORE)) {
          database.createObjectStore(PAGES_STORE, { keyPath: "cacheKey" });
        }
        if (!database.objectStoreNames.contains(META_STORE)) {
          database.createObjectStore(META_STORE, { keyPath: "cacheKey" });
        }
      };
      request.onsuccess = () => {
        const database = request.result;
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
  void opening.then((database) => {
    if (!database && databasePromise === opening) databasePromise = null;
  });
  return opening;
}

/**
 * Once per account in this tab: delete every record and key another account
 * left (an account switch whose logout cleanup never ran). Only keys are
 * read.
 */
function purgeOtherPrincipals(
  database: IDBDatabase,
  principal: string,
): Promise<void> {
  const marker = `${cacheSessionEpoch}:${principal}`;
  if (purgedFor.has(marker)) return Promise.resolve();
  return new Promise((resolve) => {
    try {
      const transaction = database.transaction(
        [PAGES_STORE, META_STORE],
        "readwrite",
      );
      for (const name of [PAGES_STORE, META_STORE]) {
        const store = transaction.objectStore(name);
        const keys = store.getAllKeys();
        keys.onsuccess = () => {
          for (const key of keys.result as IDBValidKey[]) {
            if (typeof key !== "string" || !key.startsWith(`${principal}:`)) {
              store.delete(key);
            }
          }
        };
      }
      transaction.oncomplete = () => {
        purgedFor.add(marker);
        // Asking for this account's key drops every other account's keys.
        void principalCacheKey(principal, "saved-page").then(() => resolve(), () => resolve());
      };
      transaction.onerror = () => resolve();
      transaction.onabort = () => resolve();
    } catch {
      resolve();
    }
  });
}

function readRecord(
  database: IDBDatabase,
  key: string,
): Promise<StoredPageRecord | null> {
  return new Promise((resolve) => {
    try {
      const request = database
        .transaction(PAGES_STORE, "readonly")
        .objectStore(PAGES_STORE)
        .get(key);
      request.onsuccess = () =>
        resolve((request.result as StoredPageRecord | undefined) ?? null);
      request.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

/** A read is a use: the least recently used copies are evicted first. */
function touchRecord(database: IDBDatabase, key: string): void {
  try {
    const store = database
      .transaction(META_STORE, "readwrite")
      .objectStore(META_STORE);
    const request = store.get(key);
    request.onsuccess = () => {
      const meta = request.result as StoredPageMeta | undefined;
      if (meta?.cacheKey === key) store.put({ ...meta, usedAt: Date.now() });
    };
  } catch {
    // Recency is an eviction hint only.
  }
}

function deleteRecord(database: IDBDatabase, key: string): void {
  try {
    const transaction = database.transaction(
      [PAGES_STORE, META_STORE],
      "readwrite",
    );
    transaction.objectStore(PAGES_STORE).delete(key);
    transaction.objectStore(META_STORE).delete(key);
  } catch {
    // Optional cache cleanup never affects the document load path.
  }
}

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return (
    typeof value === "object" &&
    value !== null &&
    Object.prototype.toString.call(value) === "[object ArrayBuffer]"
  );
}

function storedRecordIsValid(
  record: StoredPageRecord,
  principal: string,
  pageId: string,
): boolean {
  const now = Date.now();
  return (
    record?.version === RECORD_VERSION &&
    record.cacheKey === cacheKey(principal, pageId) &&
    record.principal === principal &&
    record.pageId === pageId &&
    typeof record.workspaceId === "string" &&
    record.workspaceId.length <= 128 &&
    typeof record.notebookId === "string" &&
    record.notebookId.length > 0 &&
    record.notebookId.length <= 128 &&
    Number.isFinite(record.savedAt) &&
    record.savedAt > 0 &&
    record.savedAt <= now + MAX_FUTURE_SKEW_MS &&
    now - record.savedAt <= MAX_AGE_MS &&
    Number.isFinite(record.expiresAt) &&
    record.expiresAt >= now &&
    record.expiresAt <= record.savedAt + MAX_AGE_MS &&
    Number.isSafeInteger(record.bytes) &&
    record.bytes > 0 &&
    record.bytes <= MAX_PAGE_BYTES &&
    Number.isSafeInteger(record.contentVersion) &&
    record.contentVersion > 0 &&
    isArrayBuffer(record.iv) &&
    record.iv.byteLength === 12 &&
    isArrayBuffer(record.ciphertext) &&
    record.ciphertext.byteLength > 16 &&
    record.ciphertext.byteLength <= MAX_PAGE_BYTES + 32
  );
}

function parsedPage(value: unknown, pageId: string): PageItem | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const page = value as Record<string, unknown>;
  if (
    page.id !== pageId ||
    typeof page.title !== "string" ||
    page.title.length > 16_384 ||
    typeof page.page_type !== "string" ||
    !SUPPORTED_PAGE_TYPES.has(page.page_type) ||
    typeof page.notebook_id !== "string" ||
    page.notebook_id.length === 0 ||
    !Number.isSafeInteger(page.content_version) ||
    (page.content_version as number) <= 0 ||
    !page.content_json ||
    typeof page.content_json !== "object" ||
    Array.isArray(page.content_json)
  ) {
    return null;
  }
  return page as unknown as PageItem;
}

/** A spreadsheet is kept in its stored form and read back inline, exactly
 * as `getPage` hands it to every reader (lib/spreadsheet/cell-style-table). */
function storedForm(page: PageItem): PageItem {
  if (page.page_type !== "spreadsheet" || !page.content_json) return page;
  try {
    const packed = packCellStyleTable(page.content_json);
    return packed === page.content_json ? page : { ...page, content_json: packed };
  } catch {
    return page;
  }
}

function inlineForm(page: PageItem): PageItem {
  if (!hasCellStyleTable(page.content_json)) return page;
  return { ...page, content_json: expandCellStyleTable(page.content_json) };
}

export async function readSavedPageHotCache(
  pageId: string,
  expectedPageType?: "code" | "spreadsheet",
  /** `opening`: the caller is opening the page and measures its read of the
   * server against this copy (SAVED_PAGE_HOT_CACHE_OPEN_WAIT_MS). */
  options: { opening?: boolean } = {},
): Promise<PageItem | null> {
  if (!pageId) return null;
  const principal = currentPrincipal();
  if (!principal) return null;
  const key = cacheKey(principal, pageId);
  if (!key) return null;
  if (cacheKeyIsBlocked(key)) return null;
  const operationEpoch = currentOperationEpoch(key);
  const sessionEpoch = cacheSessionEpoch;
  const keyGeneration = principalCacheKeyGeneration();
  const current = () =>
    cacheSessionEpoch === sessionEpoch &&
    principalCacheKeyGeneration() === keyGeneration &&
    currentPrincipal() === principal &&
    currentOperationEpoch(key) === operationEpoch &&
    !cacheKeyIsBlocked(key);
  const result = await bounded(
    (async () => {
      const database = await openDatabase();
      if (!database) return null;
      await purgeOtherPrincipals(database, principal);
      const record = await readRecord(database, key);
      if (!record || !storedRecordIsValid(record, principal, pageId)) {
        return null;
      }
      // Kept for another workspace than the one this page is read in now.
      if (record.workspaceId !== pageWorkspace(record.notebookId)) return null;
      const keyObject = await principalCacheKey(principal, "saved-page");
      if (!keyObject) return null;
      try {
        const plaintext = await globalThis.crypto.subtle.decrypt(
          {
            name: "AES-GCM",
            iv: record.iv,
            additionalData: additionalData(
              principal,
              record.workspaceId,
              pageId,
            ),
          },
          keyObject,
          record.ciphertext,
        );
        if (plaintext.byteLength !== record.bytes) return null;
        const page = parsedPage(
          JSON.parse(new TextDecoder().decode(plaintext)),
          pageId,
        );
        if (
          !page ||
          page.notebook_id !== record.notebookId ||
          (expectedPageType
            ? page.page_type !== expectedPageType
            : EDITOR_OWNED_PAGE_TYPES.has(page.page_type ?? ""))
        )
          return null;
        if (page.content_version !== record.contentVersion || !current()) {
          return null;
        }
        touchRecord(database, key);
        return inlineForm(page);
      } catch {
        // A failed read has no write authority: a newer receipt may already
        // have replaced this record while decryption was pending, or the
        // key it was sealed with was retired.
        return null;
      }
    })(),
    options.opening
      ? SAVED_PAGE_HOT_CACHE_OPEN_WAIT_MS
      : SAVED_PAGE_HOT_CACHE_WAIT_MS,
  );
  return result === TIMEOUT ? null : result;
}

/**
 * Whether this browser may hold a copy of the page that a window opening it
 * would read the server against (`readSavedPageHotCache`): only the record's
 * bookkeeping is looked at, nothing is decrypted. Unsure (no answer in time,
 * a storage error) is "may": a window that is opening then reads the page
 * as it always did.
 */
export async function savedPageHotCacheMayHold(pageId: string): Promise<boolean> {
  if (!pageId) return false;
  const principal = currentPrincipal();
  if (!principal) return false;
  const key = cacheKey(principal, pageId);
  if (!key || cacheKeyIsBlocked(key)) return false;
  const result = await bounded(
    (async () => {
      const database = await openDatabase();
      if (!database) return false;
      return new Promise<boolean>((resolve) => {
        try {
          const request = database
            .transaction(META_STORE, "readonly")
            .objectStore(META_STORE)
            .get(key);
          request.onsuccess = () => {
            const meta = request.result as StoredPageMeta | undefined;
            resolve(Boolean(meta && meta.cacheKey === key && !(meta.expiresAt < Date.now())));
          };
          request.onerror = () => resolve(true);
        } catch {
          resolve(true);
        }
      });
    })(),
  );
  return result === TIMEOUT ? true : result;
}

function writeRecord(
  database: IDBDatabase,
  record: StoredPageRecord,
  stillCurrent: () => boolean,
): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const transaction = database.transaction(
        [PAGES_STORE, META_STORE],
        "readwrite",
      );
      const pages = transaction.objectStore(PAGES_STORE);
      const metas = transaction.objectStore(META_STORE);
      const metaRequest = metas.getAll();
      metaRequest.onsuccess = () => {
        if (!stillCurrent()) {
          transaction.abort();
          return;
        }
        const now = Date.now();
        const all = (metaRequest.result as StoredPageMeta[]).filter(
          (candidate) => candidate && typeof candidate.cacheKey === "string",
        );
        const existing = all.find(
          (candidate) => candidate.cacheKey === record.cacheKey,
        );
        // A late older receipt never replaces a newer copy.
        const keepExisting =
          existing !== undefined &&
          Number.isSafeInteger(existing.contentVersion) &&
          existing.contentVersion > record.contentVersion &&
          existing.expiresAt >= now;
        const selectedBytes = keepExisting ? existing.bytes : record.bytes;
        const ranked = all
          .filter(
            (candidate) =>
              candidate.cacheKey !== record.cacheKey &&
              candidate.cacheKey.startsWith(`${record.principal}:`) &&
              Number.isSafeInteger(candidate.bytes) &&
              candidate.bytes > 0 &&
              Number.isFinite(candidate.expiresAt) &&
              candidate.expiresAt >= now,
          )
          .sort(
            (left, right) =>
              Math.max(right.usedAt || 0, right.savedAt || 0) -
              Math.max(left.usedAt || 0, left.savedAt || 0),
          );
        const keep = new Set<string>([record.cacheKey]);
        let totalBytes = selectedBytes;
        for (const candidate of ranked) {
          if (
            keep.size >= MAX_RECORDS ||
            totalBytes + candidate.bytes > MAX_TOTAL_BYTES
          ) {
            continue;
          }
          keep.add(candidate.cacheKey);
          totalBytes += candidate.bytes;
        }
        if (!keepExisting) {
          pages.put(record);
          metas.put({
            cacheKey: record.cacheKey,
            savedAt: record.savedAt,
            usedAt: record.savedAt,
            expiresAt: record.expiresAt,
            bytes: record.bytes,
            contentVersion: record.contentVersion,
          } satisfies StoredPageMeta);
        }
        for (const candidate of all) {
          if (!keep.has(candidate.cacheKey)) metas.delete(candidate.cacheKey);
        }
        // Records without their meta (an interrupted write) go too.
        const pageKeys = pages.getAllKeys();
        pageKeys.onsuccess = () => {
          for (const key of pageKeys.result as IDBValidKey[]) {
            if (typeof key !== "string" || !keep.has(key)) pages.delete(key);
          }
        };
      };
      metaRequest.onerror = () => transaction.abort();
      transaction.oncomplete = () => resolve(true);
      transaction.onerror = () => resolve(false);
      transaction.onabort = () => resolve(false);
    } catch {
      resolve(false);
    }
  });
}

export async function writeSavedPageHotCache(
  page: PageItem,
  scope: SavedPageHotCacheScope = captureSavedPageHotCacheScope(),
): Promise<void> {
  if (!isSavedPageHotCacheScopeCurrent(scope)) return;
  const pageId = page.id;
  const principal = scope.principal;
  if (!principal || !pageId) return;
  const key = cacheKey(principal, pageId);
  if (!key) return;
  const operationEpoch = nextOperationEpoch(key);
  const sessionEpoch = cacheSessionEpoch;
  const keyGeneration = principalCacheKeyGeneration();
  const stillCurrent = () =>
    cacheSessionEpoch === sessionEpoch &&
    principalCacheKeyGeneration() === keyGeneration &&
    currentPrincipal() === principal &&
    currentOperationEpoch(key) === operationEpoch;
  if (
    !page.page_type ||
    !SUPPORTED_PAGE_TYPES.has(page.page_type) ||
    !page.content_json ||
    typeof page.notebook_id !== "string" ||
    !page.notebook_id ||
    !Number.isSafeInteger(page.content_version) ||
    (page.content_version ?? 0) <= 0
  ) {
    await deleteSavedPageHotCache(pageId);
    return;
  }
  let plaintext: Uint8Array;
  try {
    // In slices: a long document's page was one task of hundreds of
    // milliseconds that froze every other window (lib/json-slices).
    plaintext = await encodeJsonInSlices(storedForm(page));
  } catch {
    return;
  }
  if (!stillCurrent()) return;
  if (plaintext.byteLength === 0 || plaintext.byteLength > MAX_PAGE_BYTES) {
    await deleteSavedPageHotCache(pageId);
    return;
  }
  const workspaceId = pageWorkspace(page.notebook_id);
  try {
    const database = await bounded(openDatabase(), WRITE_WAIT_MS);
    if (database === TIMEOUT || !database || !stillCurrent()) return;
    await purgeOtherPrincipals(database, principal);
    const keyObject = await principalCacheKey(principal, "saved-page", {
      create: true,
    });
    if (!keyObject || !stillCurrent()) return;
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
    const ciphertext = await globalThis.crypto.subtle.encrypt(
      {
        name: "AES-GCM",
        iv,
        additionalData: additionalData(principal, workspaceId, pageId),
      },
      keyObject,
      // The encoder's own buffer, not a copy made byte by byte.
      (plaintext.byteOffset === 0 && plaintext.byteLength === plaintext.buffer.byteLength
        ? plaintext.buffer : plaintext.slice().buffer) as ArrayBuffer,
    );
    if (!stillCurrent()) return;
    const savedAt = Date.now();
    const written = await bounded(
      writeRecord(
        database,
        {
          cacheKey: key,
          principal,
          workspaceId,
          notebookId: page.notebook_id,
          pageId,
          savedAt,
          expiresAt: savedAt + MAX_AGE_MS,
          bytes: plaintext.byteLength,
          contentVersion: page.content_version as number,
          iv: Uint8Array.from(iv).buffer,
          ciphertext,
          version: RECORD_VERSION,
        },
        stillCurrent,
      ),
      WRITE_WAIT_MS,
    );
    if (written === true && stillCurrent()) unblockCacheKey(key);
  } catch {
    // Optional local preview cannot affect the authoritative save receipt.
  }
}

export async function deleteSavedPageHotCache(pageId: string): Promise<void> {
  const principal = currentPrincipal();
  if (!principal || !pageId) return;
  const key = cacheKey(principal, pageId);
  if (!key) return;
  nextOperationEpoch(key);
  blockCacheKey(key);
  const result = await bounded(openDatabase());
  if (result === TIMEOUT || !result) return;
  deleteRecord(result, key);
}

function clearAllRecords(): Promise<void> {
  return (async () => {
    const database = await openDatabase();
    if (!database) return;
    await new Promise<void>((resolve) => {
      try {
        const transaction = database.transaction(
          [PAGES_STORE, META_STORE],
          "readwrite",
        );
        transaction.objectStore(PAGES_STORE).clear();
        transaction.objectStore(META_STORE).clear();
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => resolve();
        transaction.onabort = () => resolve();
      } catch {
        resolve();
      }
    });
  })().catch(() => undefined);
}

/**
 * Logout, session expiry or an account change: nothing kept may be shown
 * again. The keys go first (every sealed copy is unreadable from that
 * moment), then the records. Synchronous for the caller; await
 * `savedPageHotCacheCleared()` before leaving the page.
 */
export function clearSavedPageHotCacheSession(): void {
  if (typeof window === "undefined") return;
  cacheSessionEpoch += 1;
  purgedFor.clear();
  operationEpochs.clear();
  const keysCleared = clearPrincipalCacheKeys();
  lastWipe = Promise.all([keysCleared, clearAllRecords()]).then(
    () => undefined,
    () => undefined,
  );
  try {
    const keys: string[] = [];
    for (let index = 0; index < window.sessionStorage.length; index += 1) {
      const key = window.sessionStorage.key(index);
      if (
        key?.startsWith(LEGACY_KEY_PREFIX) ||
        key?.startsWith(TOMBSTONE_PREFIX) ||
        key?.startsWith(LEGACY_TOMBSTONE_PREFIX) ||
        key?.startsWith(OPERATION_EPOCH_PREFIX) ||
        key?.startsWith(LEGACY_OPERATION_EPOCH_PREFIX) ||
        key === LEGACY_TAB_ID_KEY
      ) {
        keys.push(key);
      }
    }
    for (const key of keys) window.sessionStorage.removeItem(key);
  } catch {
    // The cleared keys already make every kept copy unreadable.
  }
}

/** Settles (never rejects) once the last clear reached the store, or the
 * store proved unavailable. */
export function savedPageHotCacheCleared(): Promise<void> {
  return Promise.all([lastWipe, principalCacheKeysCleared()]).then(
    () => undefined,
    () => undefined,
  );
}

export function resetSavedPageHotCacheForTests(): void {
  databasePromise = null;
  operationEpochs.clear();
  blockedKeys.clear();
  purgedFor.clear();
  cacheSessionEpoch = 0;
  lastWipe = Promise.resolve();
  legacyDatabaseDropped = false;
}
