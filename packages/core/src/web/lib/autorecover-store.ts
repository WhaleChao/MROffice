/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * AutoRecover: the copy of a page's unsaved edits this browser keeps until
 * the server has them, so closing the tab or the browser, a crash, an
 * expired session or a lost permission never loses them (Office's
 * AutoRecover / Document Recovery). The editor notes every change; the copy
 * is written a moment later (coalesced), and again when the tab is hidden.
 *
 * - One "current" copy per page and account, plus "set-aside" copies: edits
 *   the window could no longer save (permission lost, discarded for the
 *   remote version, the server version kept at recovery). They are only
 *   ever offered to the person; nothing here is applied by itself.
 * - Records are partitioned by the login principal cookie (an HMAC of the
 *   user id) and sealed with AES-GCM under that account's non-extractable
 *   key kept in the same store; the additional data binds principal, page
 *   and slot. Another account on this browser lists nothing of them.
 * - Unlike the saved-page cache, an expired session does not clear the
 *   store: surviving it is the point. Sign-out (lib/api `logout`) does.
 * - Bounded: a current copy lives MAX_CURRENT_AGE_MS, a set-aside one
 *   MAX_ASIDE_AGE_MS, at most MAX_RECORDS per account (oldest set-aside
 *   evicted first) and MAX_RECORD_BYTES each.
 *
 * Every operation is best-effort and never throws.
 */
import { rememberPageCost, rememberedPageCost } from "./remembered-page-costs";
import { readCookie } from "./cookie";
import { encodeJsonInSlices } from "./json-slices";
import { inkTraceWork } from "./ink-input-trace";

const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
const PRINCIPAL = /^[0-9a-f]{64}$/u;
const DB_NAME = "mrai-autorecover-v1";
const DB_VERSION = 1;
const RECORDS = "records";
const KEYS = "keys";
const MAX_CURRENT_AGE_MS = 7 * 24 * 60 * 60 * 1_000;
const MAX_ASIDE_AGE_MS = 30 * 24 * 60 * 60 * 1_000;
const MAX_RECORDS = 40;
const MAX_RECORD_BYTES = 48 * 1024 * 1024;
/** How long after a change its copy is written (later changes join it). */
export const AUTORECOVER_WRITE_DELAY_MS = 600;







const AUTORECOVER_PAUSE_PER_COST = 12;
const AUTORECOVER_LONGEST_PAUSE_MS = 30_000;
/** Pauses that short never came for this long: the shortest pause will do. */
const AUTORECOVER_OVERDUE_MS = 5 * 60 * 1000;
/** The longest a copy waits for the page's own save on its way (noteAutoRecoverChange). */
export const AUTORECOVER_SAVE_WAIT_MS = 30_000;

export type AutoRecoverReason =
  | "unsaved"
  | "permission-lost"
  | "signed-out"
  | "discarded-for-remote"
  | "kept-server-version"
  | "not-applied";

export interface AutoRecoverPayload {
  /** The page body as the editor saves it. */
  content: Record<string, unknown>;
  title?: string;
  /** Editor-specific extra state (a tiled workbook's unsent cell edits). */
  extra?: Record<string, unknown>;
}

export interface AutoRecoverCopy {
  pageId: string;
  /** "current:<tab>" (a tab's live copy) or "aside:<time>-<n>". */
  slot: string;
  savedAt: number;
  /** The server version the edits were made on (0 when unknown). */
  baseContentVersion: number;
  reason: AutoRecoverReason;
  payload: AutoRecoverPayload;
  /** A tab's current copy written before that tab's last change: the tab
   * closed before the copy caught up (an Undo a moment before closing), so it
   * holds an earlier state than the tab ended with. */
  stale?: boolean;
}

export interface AutoRecoverEntry {
  payload: AutoRecoverPayload;
  baseContentVersion: number;
  reason?: AutoRecoverReason;
}

interface StoredRecord {
  id: string;
  principal: string;
  pageId: string;
  slot: string;
  savedAt: number;
  baseContentVersion: number;
  reason: AutoRecoverReason;
  bytes: number;
  iv: ArrayBuffer;
  ciphertext: ArrayBuffer;
}

interface StoredKey {
  principal: string;
  key: CryptoKey;
}

let databasePromise: Promise<IDBDatabase | null> | null = null;
const keys = new Map<string, Promise<CryptoKey | null>>();
let clearGeneration = 0;
let asideCounter = 0;

/**
 * Each tab writes its own current copy ("current:<tab>"), and holds a Web
 * Lock named after the tab for as long as it is open: a copy whose tab still
 * holds its lock belongs to a window that is still working (its edits reach
 * the others through the server), so another tab never offers it as a
 * recovery. Without Web Locks every other tab's copy counts as left behind.
 */
const TAB_ID = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
const TAB_LOCK_PREFIX = "mrai-autorecover-tab:";
export const AUTORECOVER_CURRENT_SLOT = `current:${TAB_ID}`;
let tabLockHeld = false;

/** This tab's id (its AutoRecover slot's and its room presence's). */
export function autoRecoverTabId(): string {
  return TAB_ID;
}

const CLOSED_TABS_KEY = "mrai-closed-tabs-v1";
const CLOSED_TAB_MEMORY_MS = 5 * 60_000;

/** Tabs of this device that closed in the last few minutes. */
export function recentlyClosedTabs(): Set<string> {
  if (typeof window === "undefined") return new Set();
  try {
    const parsed = JSON.parse(window.localStorage.getItem(CLOSED_TABS_KEY) ?? "{}") as Record<string, number>;
    const now = Date.now();
    return new Set(Object.entries(parsed).filter(([, at]) => typeof at === "number" && now - at < CLOSED_TAB_MEMORY_MS).map(([tab]) => tab));
  } catch {
    return new Set();
  }
}

function noteTabClosed(): void {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(CLOSED_TABS_KEY) ?? "{}") as Record<string, number>;
    const now = Date.now();
    const kept = Object.fromEntries(Object.entries(parsed).filter(([, at]) => typeof at === "number" && now - at < CLOSED_TAB_MEMORY_MS));
    kept[TAB_ID] = now;
    window.localStorage.setItem(CLOSED_TABS_KEY, JSON.stringify(kept));
  } catch {
    // no storage: a closed tab's presence is waited out
  }
}

export function isCurrentAutoRecoverSlot(slot: string): boolean {
  return slot === "current" || slot.startsWith("current:");
}

function holdTabLock(): void {
  if (tabLockHeld || typeof navigator === "undefined") return;
  const locks = (navigator as Navigator & { locks?: LockManager }).locks;
  if (!locks?.request) return;
  tabLockHeld = true;
  void locks.request(`${TAB_LOCK_PREFIX}${TAB_ID}`, () => new Promise<void>(() => {})).catch(() => {
    tabLockHeld = false;
  });
}

async function liveTabIds(): Promise<Set<string>> {
  const live = new Set<string>([TAB_ID]);
  try {
    const locks = (navigator as Navigator & { locks?: LockManager }).locks;
    const snapshot = await locks?.query?.();
    for (const lock of snapshot?.held ?? []) {
      if (lock.name?.startsWith(TAB_LOCK_PREFIX)) live.add(lock.name.slice(TAB_LOCK_PREFIX.length));
    }
  } catch {
    // no Web Locks: only this tab is known to be alive
  }
  return live;
}

export function autoRecoverPrincipal(): string | null {
  const value = readCookie(PRINCIPAL_COOKIE);
  return value && PRINCIPAL.test(value) ? value : null;
}

function recordId(principal: string, pageId: string, slot: string): string {
  return `${principal}:${pageId}:${slot}`;
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (databasePromise) return databasePromise;
  if (typeof indexedDB === "undefined") return Promise.resolve(null);
  const opening = new Promise<IDBDatabase | null>((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(RECORDS)) {
          database.createObjectStore(RECORDS, { keyPath: "id" }).createIndex("principal", "principal");
        }
        if (!database.objectStoreNames.contains(KEYS)) database.createObjectStore(KEYS, { keyPath: "principal" });
      };
      request.onsuccess = () => {
        const database = request.result;
        try {
          database.onversionchange = () => {
            database.close();
            if (databasePromise === opening) databasePromise = null;
          };
        } catch {
          // still usable by this tab
        }
        resolve(database);
      };
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  databasePromise = opening;
  void opening.then((database) => {
    if (!database && databasePromise === opening) databasePromise = null;
  });
  return opening;
}

function transact<T>(
  stores: string[],
  mode: IDBTransactionMode,
  run: (transaction: IDBTransaction, done: (value: T) => void) => void,
  fallback: T,
): Promise<T> {
  return openDatabase().then((database) => new Promise<T>((resolve) => {
    if (!database) {
      resolve(fallback);
      return;
    }
    let result = fallback;
    try {
      const transaction = database.transaction(stores, mode);
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => resolve(fallback);
      transaction.onabort = () => resolve(fallback);
      run(transaction, (value) => { result = value; });
    } catch {
      resolve(fallback);
    }
  }));
}

function principalKey(principal: string): Promise<CryptoKey | null> {
  const known = keys.get(principal);
  if (known) return known;
  const generation = clearGeneration;
  const pending = (async () => {
    const subtle = globalThis.crypto?.subtle;
    if (!subtle) return null;
    const candidate = (await subtle.generateKey({ name: "AES-GCM", length: 256 }, false, [
      "encrypt",
      "decrypt",
    ])) as CryptoKey;
    const key = await transact<CryptoKey | null>([KEYS], "readwrite", (transaction, done) => {
      const store = transaction.objectStore(KEYS);
      const request = store.get(principal);
      request.onsuccess = () => {
        const stored = request.result as StoredKey | undefined;
        if (stored?.key) {
          done(stored.key);
          return;
        }
        store.put({ principal, key: candidate } satisfies StoredKey);
        done(candidate);
      };
    }, null);
    return generation === clearGeneration ? key : null;
  })().catch(() => null);
  keys.set(principal, pending);
  void pending.then((key) => {
    if (!key && keys.get(principal) === pending) keys.delete(principal);
  });
  return pending;
}

function additionalData(principal: string, pageId: string, slot: string): Uint8Array<ArrayBuffer> {
  return new TextEncoder().encode(`mrai-autorecover:v1:${principal}:${pageId}:${slot}`);
}

async function seal(
  principal: string,
  pageId: string,
  slot: string,
  entry: AutoRecoverEntry,
  savedAt: number,
): Promise<StoredRecord | null> {
  const key = await principalKey(principal);
  if (!key) return null;
  const plain = await encodeJsonInSlices(entry.payload);
  if (plain.byteLength > MAX_RECORD_BYTES) return null;
  const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await globalThis.crypto.subtle.encrypt(
    { name: "AES-GCM", iv, additionalData: additionalData(principal, pageId, slot) },
    key,
    plain,
  );
  return {
    id: recordId(principal, pageId, slot),
    principal,
    pageId,
    slot,
    savedAt,
    baseContentVersion: Number.isSafeInteger(entry.baseContentVersion) ? entry.baseContentVersion : 0,
    reason: entry.reason ?? "unsaved",
    bytes: plain.byteLength,
    iv: iv.buffer,
    ciphertext,
  };
}

async function open(record: StoredRecord): Promise<AutoRecoverCopy | null> {
  try {
    const key = await principalKey(record.principal);
    if (!key) return null;
    const plain = await globalThis.crypto.subtle.decrypt(
      { name: "AES-GCM", iv: new Uint8Array(record.iv), additionalData: additionalData(record.principal, record.pageId, record.slot) },
      key,
      record.ciphertext,
    );
    const payload = JSON.parse(new TextDecoder().decode(plain)) as AutoRecoverPayload;
    if (!payload || typeof payload !== "object" || !payload.content || typeof payload.content !== "object") return null;
    return {
      pageId: record.pageId,
      slot: record.slot,
      savedAt: record.savedAt,
      baseContentVersion: record.baseContentVersion,
      reason: record.reason,
      payload,
    };
  } catch {
    return null;
  }
}

function expired(record: StoredRecord, now: number): boolean {
  const age = now - record.savedAt;
  return age < -5 * 60_000 || age > (isCurrentAutoRecoverSlot(record.slot) ? MAX_CURRENT_AGE_MS : MAX_ASIDE_AGE_MS);
}

async function put(record: StoredRecord): Promise<boolean> {
  return transact<boolean>([RECORDS], "readwrite", (transaction, done) => {
    const store = transaction.objectStore(RECORDS);
    store.put(record);
    done(true);
    // Bound this account's records: expired first, then the oldest set-aside.
    const all = store.index("principal").getAll(record.principal);
    all.onsuccess = () => {
      const now = Date.now();
      const kept = (all.result as StoredRecord[]).filter((entry) => {
        if (entry.id === record.id || !expired(entry, now)) return true;
        store.delete(entry.id);
        return false;
      });
      const excess = kept.length - MAX_RECORDS;
      if (excess > 0) {
        kept
          .filter((entry) => !isCurrentAutoRecoverSlot(entry.slot) && entry.id !== record.id)
          .sort((left, right) => left.savedAt - right.savedAt)
          .slice(0, excess)
          .forEach((entry) => store.delete(entry.id));
      }
    };
  }, false);
}

// ── the current copy, written a moment after each change ───────────────────

interface PendingWrite {
  principal: string;
  build: () => AutoRecoverEntry | null;
  timer: ReturnType<typeof setTimeout> | null;
  /** Bumped by every note and cancel: a write started before is superseded. */
  generation: number;
  /** When the oldest change this write carries was made. */
  since: number;
}
/** How long reading each page's copy off the page last took, ms. */
const copyReadCosts = new Map<string, number>();
/** A copy that took at least this long to read from the page waits for the
 * page's own save (an editor's `deferWhile`): a 6.4 MB handwritten note's copy,
 * three bodies serialized and sealed, was 0.4 s of the editor's thread at every
 * pause of the pen, for a copy the save then cleared a moment later. */
export const AUTORECOVER_COSTLY_COPY_MS = 30;
/** How long the page's last copy took to read (0 before its first, in this
 * visit or an earlier one in this browser). */
export function autoRecoverCopyCost(pageId: string): number {
  const measured = copyReadCosts.get(pageId);
  if (measured !== undefined) return measured;
  return rememberedPageCost("recover-copy", pageId);
}

const pendingWrites = new Map<string, PendingWrite>();
const writeGenerations = new Map<string, number>();
const inFlightWrites = new Set<Promise<unknown>>();
let lifecycleInstalled = false;

/** When each page last changed in this tab (cleared once the server holds it). */
const lastChanges = new Map<string, { principal: string; at: number }>();
const STALE_MARKS_KEY = "mrai-autorecover-stale-v1";
const markKey = (principal: string, pageId: string, slot: string) => `${principal}|${pageId}|${slot}`;

function readStaleMarks(): Record<string, number> {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(STALE_MARKS_KEY) ?? "{}") as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, number> : {};
  } catch {
    return {};
  }
}

function writeStaleMarks(marks: Record<string, number>): void {
  try {
    if (Object.keys(marks).length) window.localStorage.setItem(STALE_MARKS_KEY, JSON.stringify(marks));
    else window.localStorage.removeItem(STALE_MARKS_KEY);
  } catch {
    // no storage: a copy is taken as written
  }
}

/**
 * The tab is going: say, synchronously, when each of its pages last changed.
 * Its copies are written asynchronously (sealed, then stored) and can stop
 * short of the last change -- an Undo a moment before closing left a copy
 * that still held the undone words, and the reopened view took them back
 * (a45c sac22/sac23). A copy saved before its mark is known to be behind.
 */
export function markAutoRecoverCopiesClosing(): void {
  if (typeof window === "undefined" || !lastChanges.size) return;
  const marks = readStaleMarks();
  for (const [pageId, change] of lastChanges) marks[markKey(change.principal, pageId, AUTORECOVER_CURRENT_SLOT)] = change.at;
  writeStaleMarks(marks);
  // The change the copy has not caught up with, written now while the tab
  // still can: the page as it is, synchronously. The edits a closing tab had
  // queued behind a save still on its way were otherwise only in this tab's
  // session storage, gone with it -- an Undo 100 ms before closing came back
  // undone (the save before it landed, the Undo's never left; a45c sac24).
  for (const [pageId, pending] of pendingWrites) {
    let entry: AutoRecoverEntry | null = null;
    try {
      entry = pending.build();
    } catch {
      entry = null;
    }
    if (!entry) continue;
    try {
      let text = JSON.stringify({ savedAt: Date.now(), entry });
      // Too large with every body it carries: the page as it is and the one
      // base the editor merges against first (the extra bodies are bases).
      for (const keep of [2, 1, 0]) {
        if (text.length <= CLOSING_COPY_MAX_CHARS) break;
        const extra = entry.payload.extra;
        if (!extra) break;
        const kept = Object.fromEntries(Object.entries(extra).slice(0, keep));
        text = JSON.stringify({ savedAt: Date.now(), entry: { ...entry, payload: { ...entry.payload, extra: kept } } });
      }
      if (text.length > CLOSING_COPY_MAX_CHARS) continue;
      window.localStorage.setItem(closingKey(pending.principal, pageId, AUTORECOVER_CURRENT_SLOT), text);
    } catch {
      // over quota: the encrypted copy and its stale mark are what is left
    }
  }
}

const HELD_MARKS_KEY = "mrai-autorecover-held-v1";

function readHeldMarks(): Record<string, number> {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(HELD_MARKS_KEY) ?? "{}") as unknown;
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed as Record<string, number> : {};
  } catch {
    return {};
  }
}

function writeHeldMarks(marks: Record<string, number>): void {
  try {
    if (Object.keys(marks).length) window.localStorage.setItem(HELD_MARKS_KEY, JSON.stringify(marks));
    else window.localStorage.removeItem(HELD_MARKS_KEY);
  } catch {
    // no storage: the copy is read as ever
  }
}

/** Each page's editors open in this tab, each answering whether it holds
 * nothing the stored page lacks (markAutoRecoverClosingHeld). The page's copy
 * is the tab's, not one editor's. */
const closingEditors = new Map<string, Set<() => boolean>>();

/** An editor of `pageId` in this tab, answering (synchronously) whether it
 * holds nothing the stored page lacks. Returns its release. */
export function noteAutoRecoverClosingEditor(pageId: string, holdsNothing: () => boolean): () => void {
  let editors = closingEditors.get(pageId);
  if (!editors) {
    editors = new Set();
    closingEditors.set(pageId, editors);
  }
  editors.add(holdsNothing);
  return () => {
    const now = closingEditors.get(pageId);
    now?.delete(holdsNothing);
    if (now && now.size === 0) closingEditors.delete(pageId);
  };
}

/**
 * The tab is going and the page is stored as the tab last held it: nothing
 * it holds differs from the stored page, nor from a save on its way. Its
 * copies of the page were all written before now, so each is behind that
 * state (a long page's copy still holding a word an Undo took out a moment
 * before closing, too large for a closing copy): where read, they are
 * dropped, never offered or merged. Synchronous.
 *
 * `editor` (noteAutoRecoverClosingEditor) says it of itself; another editor
 * of the page in this tab that holds edits, or cannot say, keeps the copy:
 * the tab's copy is that editor's too. False when nothing was dropped.
 */
export function markAutoRecoverClosingHeld(pageId: string, editor?: () => boolean): boolean {
  if (typeof window === "undefined" || !pageId) return false;
  for (const other of closingEditors.get(pageId) ?? []) {
    if (other === editor) continue;
    let holdsNothing = false;
    try {
      holdsNothing = other();
    } catch {
      holdsNothing = false;
    }
    if (!holdsNothing) return false;
  }
  const principal = autoRecoverPrincipal();
  if (!principal) return false;
  // Nothing more of it is written as the tab goes (the flush on pagehide
  // would seal a copy dated after this mark): a write waiting or sealing is
  // cancelled, and the copy already stored is deleted where it can be.
  const pending = pendingWrites.get(pageId);
  if (pending?.timer) clearTimeout(pending.timer);
  pendingWrites.delete(pageId);
  lastChanges.delete(pageId);
  nextGeneration(pageId);
  void deleteAutoRecoverCopy(pageId, AUTORECOVER_CURRENT_SLOT, principal);
  const now = Date.now();
  const marks = readHeldMarks();
  for (const [key, at] of Object.entries(marks)) if (typeof at !== "number" || now - at > CLOSING_COPY_MAX_AGE_MS) delete marks[key];
  marks[markKey(principal, pageId, AUTORECOVER_CURRENT_SLOT)] = now;
  writeHeldMarks(marks);
  return true;
}

function forgetHeldMark(principal: string, pageId: string, slot: string): void {
  const marks = readHeldMarks();
  const key = markKey(principal, pageId, slot);
  if (!(key in marks)) return;
  delete marks[key];
  writeHeldMarks(marks);
}

/** A closing copy is kept in local storage only this large (an origin has
 * about 5 MB of it); a larger page relies on its encrypted copy. */
const CLOSING_COPY_MAX_CHARS = 2_500_000;
const CLOSING_COPY_PREFIX = "mrai-autorecover-closing-v1:";
/** A closing copy never read again is dropped after a week (an encrypted copy's age). */
const CLOSING_COPY_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
const closingKey = (principal: string, pageId: string, slot: string) => `${CLOSING_COPY_PREFIX}${principal}|${pageId}|${slot}`;

function readClosingCopy(principal: string, pageId: string, slot: string): { savedAt: number; entry: AutoRecoverEntry } | null {
  if (typeof window === "undefined") return null;
  try {
    const parsed = JSON.parse(window.localStorage.getItem(closingKey(principal, pageId, slot)) ?? "null") as
      { savedAt?: unknown; entry?: AutoRecoverEntry } | null;
    if (!parsed || typeof parsed.savedAt !== "number" || !parsed.entry?.payload?.content) return null;
    return { savedAt: parsed.savedAt, entry: parsed.entry };
  } catch {
    return null;
  }
}

function forgetClosingCopy(principal: string, pageId: string, slot: string): void {
  try {
    window.localStorage.removeItem(closingKey(principal, pageId, slot));
  } catch {
    // nothing kept
  }
}

function forgetStaleMark(principal: string, pageId: string, slot: string): void {
  if (typeof window === "undefined") return;
  if (isCurrentAutoRecoverSlot(slot)) {
    forgetClosingCopy(principal, pageId, slot);
    forgetHeldMark(principal, pageId, slot);
  }
  const marks = readStaleMarks();
  const key = markKey(principal, pageId, slot);
  if (!(key in marks)) return;
  delete marks[key];
  writeStaleMarks(marks);
}

function installLifecycle(): void {
  if (lifecycleInstalled || typeof window === "undefined") return;
  lifecycleInstalled = true;
  const flush = () => { void flushAutoRecoverWrites(); };
  window.addEventListener("beforeunload", markAutoRecoverCopiesClosing);
  window.addEventListener("pagehide", (event) => {
    markAutoRecoverCopiesClosing();
    if (!event.persisted) noteTabClosed();
    flush();
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") flush();
  });
}

function nextGeneration(pageId: string): number {
  const next = (writeGenerations.get(pageId) ?? 0) + 1;
  writeGenerations.set(pageId, next);
  return next;
}

async function writeCurrent(pageId: string, pending: PendingWrite): Promise<void> {
  const readStarted = Date.now();
  const entry = pending.build();
  copyReadCosts.set(pageId, Date.now() - readStarted);
  rememberPageCost("recover-copy", pageId, Date.now() - readStarted);
  // A pen contact reported late names the copy's read when it held the page.
  inkTraceWork("recover-read", Date.now() - readStarted);
  if (!entry) return;
  const generation = pending.generation;
  const cleared = clearGeneration;
  const record = await seal(pending.principal, pageId, AUTORECOVER_CURRENT_SLOT, entry, Date.now());
  // A newer change, a save that confirmed everything, or a sign-out came in
  // while this copy was being sealed: it is not written over them.
  if (!record || writeGenerations.get(pageId) !== generation || cleared !== clearGeneration) return;
  await put(record);
  if (writeGenerations.get(pageId) !== generation) {
    // Cancelled while writing: the record must not outlive the cancel.
    await deleteAutoRecoverCopy(pageId, AUTORECOVER_CURRENT_SLOT, pending.principal);
  }
}

/**
 * The page changed: its copy is written AUTORECOVER_WRITE_DELAY_MS later
 * from what `build` returns then (later changes join the same write).
 */
export function noteAutoRecoverChange(
  pageId: string,
  build: () => AutoRecoverEntry | null,
  




  deferWhile?: () => boolean,
): void {
  if (typeof window === "undefined" || !pageId) return;
  const principal = autoRecoverPrincipal();
  if (!principal) return;
  installLifecycle();
  holdTabLock();
  lastChanges.set(pageId, { principal, at: Date.now() });
  const previous = pendingWrites.get(pageId);
  if (previous?.timer) clearTimeout(previous.timer);
  const startedAt = Date.now();
  const since = previous?.since ?? startedAt;
  const pending: PendingWrite = { principal, build, timer: null, generation: nextGeneration(pageId), since };
  const pause = startedAt - since >= AUTORECOVER_OVERDUE_MS
    ? AUTORECOVER_WRITE_DELAY_MS
    : Math.min(
        AUTORECOVER_LONGEST_PAUSE_MS,
        Math.max(AUTORECOVER_WRITE_DELAY_MS, (copyReadCosts.get(pageId) ?? 0) * AUTORECOVER_PAUSE_PER_COST),
      );
  const fire = () => {
    pending.timer = null;
    if (pendingWrites.get(pageId) !== pending) return;
    // Waited for from the first change the copy holds: edits made without a
    // pause (offline, a save refused) never keep it from being written longer.
    if (deferWhile?.() && Date.now() - since < AUTORECOVER_SAVE_WAIT_MS) {
      pending.timer = setTimeout(fire, AUTORECOVER_WRITE_DELAY_MS);
      return;
    }
    pendingWrites.delete(pageId);
    const write = writeCurrent(pageId, pending).catch(() => undefined);
    inFlightWrites.add(write);
    void write.finally(() => inFlightWrites.delete(write));
  };
  pending.timer = setTimeout(fire, pause);
  pendingWrites.set(pageId, pending);
}

/** Write every copy waiting for its delay now (the tab is being hidden). */
export async function flushAutoRecoverWrites(): Promise<void> {
  const waiting = [...pendingWrites.entries()];
  pendingWrites.clear();
  const writes = waiting.map(([pageId, pending]) => {
    if (pending.timer) clearTimeout(pending.timer);
    pending.timer = null;
    return writeCurrent(pageId, pending).catch(() => undefined);
  });
  await Promise.all([...writes, ...inFlightWrites]);
}

/**
 * The server holds everything the page's copy held: forget the copy (and a
 * write still waiting for it).
 */
export function clearCurrentAutoRecover(pageId: string): void {
  if (!pageId) return;
  const pending = pendingWrites.get(pageId);
  if (pending?.timer) clearTimeout(pending.timer);
  pendingWrites.delete(pageId);
  lastChanges.delete(pageId);
  nextGeneration(pageId);
  const principal = autoRecoverPrincipal();
  if (principal) void deleteAutoRecoverCopy(pageId, AUTORECOVER_CURRENT_SLOT, principal);
}

// ── reading, setting aside, deleting ───────────────────────────────────────

/**
 * This account's copies of the page, newest first: the set-aside ones and
 * the current copies tabs left behind (not this tab's, nor a tab's that is
 * still open). Expired ones are dropped.
 */
export async function listAutoRecoverCopies(pageId: string): Promise<AutoRecoverCopy[]> {
  const principal = autoRecoverPrincipal();
  if (!principal || !pageId) return [];
  // A copy still waiting for its delay is part of the answer.
  if (pendingWrites.has(pageId) || inFlightWrites.size > 0) await flushAutoRecoverWrites();
  const records = await transact<StoredRecord[]>([RECORDS], "readonly", (transaction, done) => {
    const request = transaction.objectStore(RECORDS).index("principal").getAll(principal);
    request.onsuccess = () => done((request.result as StoredRecord[]).filter((record) => record.pageId === pageId));
  }, []);
  const now = Date.now();
  const live = await liveTabIds();
  const marks = typeof window === "undefined" ? {} : readStaleMarks();
  const held = typeof window === "undefined" ? {} : readHeldMarks();
  // Written before its tab closed on a page stored as that tab held it.
  const heldPast = (slot: string, savedAt: number) => {
    const at = held[markKey(principal, pageId, slot)];
    return typeof at === "number" && at >= savedAt;
  };
  const copies: AutoRecoverCopy[] = [];
  for (const record of records) {
    if (expired(record, now)) {
      void deleteAutoRecoverCopy(pageId, record.slot, principal);
      continue;
    }
    if (record.slot.startsWith("current:") && live.has(record.slot.slice("current:".length))) continue;
    if (isCurrentAutoRecoverSlot(record.slot) && heldPast(record.slot, record.savedAt)) {
      void deleteAutoRecoverCopy(pageId, record.slot, principal);
      continue;
    }
    const copy = await open(record);
    if (!copy) continue;
    const closedAt = marks[markKey(principal, pageId, record.slot)];
    if (typeof closedAt === "number" && closedAt > record.savedAt) {
      // The closing tab's last state, written as it went: the copy is that.
      const closing = readClosingCopy(principal, pageId, record.slot);
      if (closing && closing.savedAt >= closedAt) {
        copy.payload = closing.entry.payload;
        copy.baseContentVersion = closing.entry.baseContentVersion;
        copy.savedAt = closing.savedAt;
      } else copy.stale = true;
    }
    copies.push(copy);
  }
  // A tab that closed before its first copy was written left only its
  // closing copy.
  if (typeof window !== "undefined") {
    const prefix = `${CLOSING_COPY_PREFIX}${principal}|${pageId}|`;
    let keys: string[] = [];
    try { keys = Object.keys(window.localStorage).filter((key) => key.startsWith(prefix)); } catch { keys = []; }
    for (const key of keys) {
      const slot = key.slice(prefix.length);
      if (!isCurrentAutoRecoverSlot(slot) || live.has(slot.slice("current:".length)) ||
          copies.some((copy) => copy.slot === slot)) continue;
      const closing = readClosingCopy(principal, pageId, slot);
      if (!closing || now - closing.savedAt > CLOSING_COPY_MAX_AGE_MS || heldPast(slot, closing.savedAt)) {
        forgetClosingCopy(principal, pageId, slot);
        forgetHeldMark(principal, pageId, slot);
        continue;
      }
      copies.push({
        pageId, slot, savedAt: closing.savedAt, baseContentVersion: closing.entry.baseContentVersion,
        reason: closing.entry.reason ?? "unsaved", payload: closing.entry.payload,
      });
    }
  }
  // A closed tab's mark has done its work once read: none of that tab's
  // copies is left (each was dropped above), and the tab writes no more.
  if (typeof window !== "undefined") {
    const prefix = markKey(principal, pageId, "");
    const left = Object.keys(held).filter((key) => key.startsWith(prefix) && !live.has(key.slice(prefix.length + "current:".length)));
    if (left.length) {
      const marks = readHeldMarks();
      for (const key of left) delete marks[key];
      writeHeldMarks(marks);
    }
  }
  return copies.sort((left, right) => right.savedAt - left.savedAt);
}

// ── keys typed while a document's view was handed over ─────────────────────

/**
 * The keys a document window held while its view was handed to another (a
 * page read again as parts, lib/editor/held-keys holdKeysThroughHandover),
 * kept with the whole view's set-aside copy (`slot`, when it has one) and the
 * caret's place they go at. Written synchronously on every key, to local
 * storage, so a tab closed in the seconds of the handover leaves them as it
 * leaves a closing copy (`markAutoRecoverCopiesClosing`); read back by the
 * page's next opening, which types them at their place once the copy is in.
 * One record per page and tab; at most HANDOVER_KEYS_MAX_CHARS of it, and
 * a week old at most, as a closing copy.
 */
export interface AutoRecoverHandoverKeys {
  /** The tab that held them (its record's key). */
  readonly tab: string;
  readonly savedAt: number;
  /** The set-aside copy of the whole view they were typed after, if any. */
  readonly slot: string | null;
  /** Where they go (native-document documentCaretPlace). */
  readonly place: unknown;
  /** The keys (held-keys HeldKey), in the order typed. */
  readonly keys: readonly unknown[];
  



  readonly edits?: { readonly blocks: readonly unknown[]; readonly base: readonly unknown[] } | null;
}
const HANDOVER_KEYS_PREFIX = "mrai-autorecover-handover-keys-v1:";
/** A record's JSON at most this long: what does not fit is still typed by the
 * tab itself, only not kept past its closing. */
export const HANDOVER_KEYS_MAX_CHARS = 256_000;
const handoverKeysKey = (principal: string, pageId: string, tab: string) => `${HANDOVER_KEYS_PREFIX}${principal}|${pageId}|${tab}`;
/** Pages whose handover keys a view of this tab still holds itself: its record is not left behind. */
const handoverKeysHeld = new Map<string, number>();

/** Keep (or with no keys, forget) this tab's handover keys for the page.
 * False when they could not be kept (no account, no storage, too many). */
export function noteAutoRecoverHandoverKeys(
  pageId: string,
  entry: { slot: string | null; place: unknown; keys: readonly unknown[];
    edits?: { readonly blocks: readonly unknown[]; readonly base: readonly unknown[] } | null },
): boolean {
  if (typeof window === "undefined" || !pageId) return false;
  const principal = autoRecoverPrincipal();
  if (!principal) return false;
  const key = handoverKeysKey(principal, pageId, TAB_ID);
  try {
    const edits = entry.edits?.blocks.length ? entry.edits : null;
    if (!entry.keys.length && !edits) {
      window.localStorage.removeItem(key);
      return true;
    }
    const record: AutoRecoverHandoverKeys = { tab: TAB_ID, savedAt: Date.now(), slot: entry.slot, place: entry.place, keys: entry.keys,
      ...(edits ? { edits } : {}) };
    let text = JSON.stringify(record);
    // Edits too long to keep here are left to the set-aside copy alone.
    if (text.length > HANDOVER_KEYS_MAX_CHARS && edits) {
      if (!entry.keys.length) return false;
      text = JSON.stringify({ ...record, edits: undefined });
    }
    if (text.length > HANDOVER_KEYS_MAX_CHARS) return false;
    holdTabLock();
    window.localStorage.setItem(key, text);
    return true;
  } catch {
    return false;
  }
}

/** Says a view of this tab holds the page's handover keys itself (so its
 * record is not read as left behind); the returned function lets go. */
export function holdAutoRecoverHandoverKeys(pageId: string): () => void {
  handoverKeysHeld.set(pageId, (handoverKeysHeld.get(pageId) ?? 0) + 1);
  let held = true;
  return () => {
    if (!held) return;
    held = false;
    const left = (handoverKeysHeld.get(pageId) ?? 1) - 1;
    if (left > 0) handoverKeysHeld.set(pageId, left);
    else handoverKeysHeld.delete(pageId);
  };
}

/** The handover keys of the page that no open view holds: a closed tab's, or
 * this tab's own left by a view that went before it could type them. */
export async function listAutoRecoverHandoverKeys(pageId: string): Promise<AutoRecoverHandoverKeys[]> {
  if (typeof window === "undefined" || !pageId) return [];
  const principal = autoRecoverPrincipal();
  if (!principal) return [];
  const prefix = `${HANDOVER_KEYS_PREFIX}${principal}|${pageId}|`;
  const names: string[] = [];
  try {
    const storage = window.localStorage;
    for (let index = 0; index < storage.length; index += 1) {
      const name = storage.key(index);
      if (name?.startsWith(prefix)) names.push(name);
    }
  } catch {
    // no storage: nothing kept
  }
  if (!names.length) return [];
  const live = await liveTabIds();
  const now = Date.now();
  const found: AutoRecoverHandoverKeys[] = [];
  for (const name of names) {
    const tab = name.slice(prefix.length);
    if (tab === TAB_ID ? handoverKeysHeld.has(pageId) : live.has(tab)) continue;
    let record: AutoRecoverHandoverKeys | null = null;
    try { record = JSON.parse(window.localStorage.getItem(name) ?? "null") as AutoRecoverHandoverKeys | null; } catch { record = null; }
    const editsKept = Boolean(record?.edits && Array.isArray(record.edits.blocks) && record.edits.blocks.length &&
      Array.isArray(record.edits.base));
    if (!record || typeof record.savedAt !== "number" || !Array.isArray(record.keys) || (!record.keys.length && !editsKept) ||
        now - record.savedAt > CLOSING_COPY_MAX_AGE_MS) {
      try { window.localStorage.removeItem(name); } catch { /* nothing kept */ }
      continue;
    }
    found.push({ ...record, tab });
  }
  return found.sort((left, right) => left.savedAt - right.savedAt);
}

let leftCounter = 0;





export function leaveAutoRecoverHandoverKeys(pageId: string): void {
  if (typeof window === "undefined" || !pageId) return;
  const principal = autoRecoverPrincipal();
  if (!principal) return;
  try {
    const from = handoverKeysKey(principal, pageId, TAB_ID);
    const text = window.localStorage.getItem(from);
    if (!text) return;
    leftCounter += 1;
    window.localStorage.setItem(handoverKeysKey(principal, pageId, `${TAB_ID}~left${leftCounter}`), text);
    window.localStorage.removeItem(from);
  } catch {
    // no storage: nothing kept
  }
}

/** The keys are in a view now (typed, its own edits from then on): their record goes. */
export function forgetAutoRecoverHandoverKeys(pageId: string, tab: string = TAB_ID): void {
  if (typeof window === "undefined" || !pageId) return;
  const principal = autoRecoverPrincipal();
  if (!principal) return;
  try { window.localStorage.removeItem(handoverKeysKey(principal, pageId, tab)); } catch { /* nothing kept */ }
}

/** Keep `entry` as a set-aside copy of the page; resolves to its slot. */
export async function setAsideAutoRecoverCopy(pageId: string, entry: AutoRecoverEntry): Promise<string | null> {
  const principal = autoRecoverPrincipal();
  if (!principal || !pageId) return null;
  asideCounter += 1;
  const slot = `aside:${Date.now().toString(36)}-${asideCounter}`;
  const record = await seal(principal, pageId, slot, entry, Date.now()).catch(() => null);
  if (!record) return null;
  return (await put(record)) ? slot : null;
}

export async function deleteAutoRecoverCopy(pageId: string, slot: string, principal = autoRecoverPrincipal()): Promise<void> {
  if (!principal || !pageId) return;
  forgetStaleMark(principal, pageId, slot);
  await transact<boolean>([RECORDS], "readwrite", (transaction, done) => {
    transaction.objectStore(RECORDS).delete(recordId(principal, pageId, slot));
    done(true);
  }, false);
}

/** Sign-out: every copy and key on this browser goes. */
export function clearAutoRecoverStore(): Promise<void> {
  clearGeneration += 1;
  keys.clear();
  for (const pending of pendingWrites.values()) if (pending.timer) clearTimeout(pending.timer);
  pendingWrites.clear();
  lastChanges.clear();
  if (typeof window !== "undefined") {
    writeStaleMarks({});
    writeHeldMarks({});
    try {
      for (const key of Object.keys(window.localStorage)) {
        if (key.startsWith(CLOSING_COPY_PREFIX) || key.startsWith(HANDOVER_KEYS_PREFIX)) window.localStorage.removeItem(key);
      }
    } catch {
      // no storage
    }
  }
  return transact<boolean>([RECORDS, KEYS], "readwrite", (transaction, done) => {
    transaction.objectStore(RECORDS).clear();
    transaction.objectStore(KEYS).clear();
    done(true);
  }, false).then(() => undefined);
}

export function resetAutoRecoverStoreForTests(): void {
  databasePromise = null;
  keys.clear();
  pendingWrites.clear();
  writeGenerations.clear();
  inFlightWrites.clear();
  lastChanges.clear();
  clearGeneration = 0;
}
