"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The page layout this browser last reached for a long document, so its next
 * opening starts from the finished pages instead of laying the book out again.
 *
 * A Docs page saved by an edit carries its layout receipt (which page each
 * block, or each piece of a split block, sits on: paginationProjectionReceipt).
 * A page that is only opened and read is never saved, so every opening of a
 * long import began again from provisional sheets and repaginated the whole
 * book (a 93-page import: 123 provisional sheets moved into 93, every time).
 * The receipt the paginator reached is kept here for the next opening.
 *
 * Safe by construction:
 * - A receipt is checked against the content it is applied to
 *   (restorePaginationProjection compares the canonical hash and the restored
 *   content exactly): any other content, an older or newer version of the
 *   page included, does not take it, and the page opens as before.
 * - It holds no text: block paths, text cut offsets, page geometry and a
 *   hash. Records are kept per signed-in principal (the page-cache principal
 *   cookie, an HMAC of the user id) and per page; another account reads none.
 * - The store is bounded (MAX_RECORDS, MAX_RECORD_CHARS, MAX_AGE_MS) and every
 *   operation is best-effort: storage that is missing, blocked or full
 *   answers nothing and the page opens from provisional sheets as before.
 * - Nothing is sent to the server: opening a page stays a read.
 */
import { readCookie } from "./cookie";

const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
const DB_NAME = "mrnote-docs-layout-receipts-v1";
const STORE = "receipts";
const MAX_RECORDS = 40;
const MAX_RECORD_CHARS = 4 * 1024 * 1024;
const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

interface ReceiptRecord {
  key: string;
  /** The exact layout (pieces of split blocks included), when it could be
   * taken; checked against the content it is applied to. */
  receipt: unknown;
  /** How many whole top-level blocks each page holds: the fallback when the
   * receipt does not restore (a different form of the same content). */
  sizes: number[] | null;
  json: string;
  at: number;
}

/** Receipts read this session, by key: the opening path is synchronous. */
const memory = new Map<string, ReceiptRecord>();
const loading = new Map<string, Promise<unknown>>();

function principal(): string | null {
  if (typeof document === "undefined") return null;
  const value = readCookie(PRINCIPAL_COOKIE);
  return value && /^[0-9a-f]{64}$/u.test(value) ? value : null;
}

function keyFor(pageId: string): string | null {
  const who = principal();
  return who && pageId ? `${who}:${pageId}` : null;
}

let database: Promise<IDBDatabase | null> | null = null;

function open(): Promise<IDBDatabase | null> {
  if (database) return database;
  database = new Promise((resolve) => {
    try {
      if (typeof indexedDB === "undefined") {
        resolve(null);
        return;
      }
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "key" });
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

function fresh(record: ReceiptRecord | undefined | null): record is ReceiptRecord {
  return Boolean(record && typeof record.json === "string" && Date.now() - record.at < MAX_AGE_MS);
}

/**
 * Start reading the receipt kept for `pageId`; resolves once it is in memory
 * (or known absent). Call as soon as the page is known to open, so the
 * receipt is ready by the time the document is laid out.
 */
export function preloadDocsLayoutReceipt(pageId: string): Promise<unknown> {
  const key = keyFor(pageId);
  if (!key) return Promise.resolve(null);
  const known = memory.get(key);
  if (known) return Promise.resolve(known.receipt);
  const pending = loading.get(key);
  if (pending) return pending;
  const read = open().then(
    (db) =>
      new Promise<unknown>((resolve) => {
        if (!db) {
          resolve(null);
          return;
        }
        try {
          const request = db.transaction(STORE, "readonly").objectStore(STORE).get(key);
          request.onsuccess = () => {
            const record = request.result as ReceiptRecord | undefined;
            if (fresh(record)) {
              memory.set(key, record);
              resolve(record.receipt);
            } else {
              resolve(null);
            }
          };
          request.onerror = () => resolve(null);
        } catch {
          resolve(null);
        }
      }),
  );
  loading.set(key, read);
  void read.finally(() => loading.delete(key));
  return read;
}

/** The receipt kept for `pageId`, once read (preloadDocsLayoutReceipt). */
export function cachedDocsLayoutReceipt(pageId: string): unknown {
  const key = keyFor(pageId);
  return key ? (memory.get(key)?.receipt ?? null) : null;
}

/** The whole-block page distribution kept for `pageId`, once read. */
export function cachedDocsLayoutSizes(pageId: string): number[] | null {
  const key = keyFor(pageId);
  const sizes = key ? memory.get(key)?.sizes : null;
  return Array.isArray(sizes) ? sizes : null;
}

/** Keep the layout reached for `pageId`: its receipt and its page distribution. */
export function storeDocsLayoutReceipt(pageId: string, receipt: unknown, sizes: number[] | null): void {
  const key = keyFor(pageId);
  if (!key || (!receipt && !sizes)) return;
  let json: string;
  try {
    json = JSON.stringify({ receipt, sizes });
  } catch {
    return;
  }
  if (json.length > MAX_RECORD_CHARS) return;
  if (memory.get(key)?.json === json) return;
  const record: ReceiptRecord = { key, receipt: receipt ?? null, sizes, json, at: Date.now() };
  memory.set(key, record);
  void open().then((db) => {
    if (!db) return;
    try {
      const store = db.transaction(STORE, "readwrite").objectStore(STORE);
      store.put(record);
      // Bounded: expired records, and the oldest beyond MAX_RECORDS, go.
      const all = store.getAll();
      all.onsuccess = () => {
        const records = (all.result as ReceiptRecord[]).sort((a, b) => b.at - a.at);
        records.forEach((item, index) => {
          if (index >= MAX_RECORDS || !fresh(item)) store.delete(item.key);
        });
      };
    } catch {
      // Best-effort: the next opening lays the pages out again.
    }
  });
}

/** Logout or session expiry: nothing of the account stays behind. */
export function clearDocsLayoutReceipts(): void {
  memory.clear();
  void open().then((db) => {
    if (!db) return;
    try {
      db.transaction(STORE, "readwrite").objectStore(STORE).clear();
    } catch {
      // Best-effort.
    }
  });
}

/** Whether a harness traces the receipts (worth a costlier diagnosis). */
export function tracingDocsLayoutReceipt(): boolean {
  if (typeof window === "undefined") return false;
  return Array.isArray((window as Window & { __docsLayoutReceiptTrace?: unknown[] }).__docsLayoutReceiptTrace);
}

/** A harness that pre-creates `__docsLayoutReceiptTrace` sees each step. */
export function traceDocsLayoutReceipt(entry: string): void {
  if (typeof window === "undefined") return;
  const sink = (window as Window & { __docsLayoutReceiptTrace?: unknown[] }).__docsLayoutReceiptTrace;
  if (Array.isArray(sink)) sink.push(`${Math.round(performance.now())} ${entry}`);
}
