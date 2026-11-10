/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The typeset boxes of a block's formulas, kept for the next opening.
 *
 * The page ledger measures each block holding a formula in a typeset clone
 * before it can lay the block out (ledgerTypesetMath): a long book of
 * equations spent some 20 s of every opening doing it again. The boxes are a
 * function of the block's content, its width and how this build of the
 * editor typesets (its KaTeX and its styles), so they are kept by those and
 * nothing else, and read back only where the clone would have measured them
 * (ledgerTypesetMath decides when).
 *
 * - It holds no text: numbers, under a 53-bit hash of the block and its
 *   length. Records are kept per signed-in principal (the page-cache
 *   principal cookie); another account reads none, and signing out clears it.
 * - Bounded (MAX_ENTRIES, MAX_AGE_MS) and best-effort: storage that is
 *   missing, blocked or full answers nothing and the clone measures.
 * - Nothing is sent to the server.
 */
import { readCookie } from "./cookie";

const PRINCIPAL_COOKIE = "mrnote_page_cache_principal";
const DB_NAME = "mrnote-docs-math-geometry-v1";
const STORE = "boxes";
const MAX_ENTRIES = 40_000;
const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
const FLUSH_DELAY_MS = 4_000;

interface BoxRecord {
  key: string;
  boxes: unknown;
  at: number;
}

let database: Promise<IDBDatabase | null> | null = null;
let loaded: Promise<void> | null = null;
const memory = new Map<string, unknown>();
const pending = new Map<string, unknown>();
let flushTimer: ReturnType<typeof setTimeout> | null = null;

function principal(): string | null {
  if (typeof document === "undefined") return null;
  const value = readCookie(PRINCIPAL_COOKIE);
  return value && /^[0-9a-f]{64}$/u.test(value) ? value : null;
}

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

/** The key of a block's boxes: the principal, what typesets them, and the
 * block's own key (content hash, width, kind). Null without a principal. */
export function mathGeometryKey(typesetter: string, blockKey: string): string | null {
  const who = principal();
  return who ? `${who}|${typesetter}|${blockKey}` : null;
}

/** Start reading this principal's kept boxes into memory, once. */
export function preloadMathGeometry(): Promise<void> {
  if (loaded) return loaded;
  const who = principal();
  if (!who) return Promise.resolve();
  loaded = open().then(
    (db) =>
      new Promise<void>((resolve) => {
        if (!db) {
          resolve();
          return;
        }
        try {
          const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
          request.onsuccess = () => {
            const now = Date.now();
            for (const record of request.result as BoxRecord[]) {
              if (record && typeof record.key === "string" && record.key.startsWith(`${who}|`) && now - record.at < MAX_AGE_MS) {
                memory.set(record.key, record.boxes);
              }
            }
            trace({ preloaded: memory.size });
            resolve();
          };
          request.onerror = () => resolve();
        } catch {
          resolve();
        }
      }),
  );
  return loaded;
}

/** The boxes kept under `key`, if they have been read. */
export function keptMathGeometry(key: string): unknown {
  return memory.get(key);
}

/** Harness switch (`window.__docsMathGeometryTrace = []`): what the cache
 * read, kept and wrote, and why a write failed. */
function trace(entry: Record<string, unknown>): void {
  if (typeof window === "undefined") return;
  const sink = (window as Window & { __docsMathGeometryTrace?: unknown[] }).__docsMathGeometryTrace;
  if (Array.isArray(sink) && sink.length < 2000) sink.push({ at: Math.round(performance.now()), ...entry });
}

function flush(): void {
  flushTimer = null;
  if (pending.size === 0) return;
  const batch = [...pending];
  pending.clear();
  void open().then((db) => {
    if (!db) return;
    try {
      const transaction = db.transaction(STORE, "readwrite");
      transaction.oncomplete = () => trace({ written: batch.length });
      transaction.onerror = () => trace({ writeFailed: String(transaction.error?.name ?? "error") });
      transaction.onabort = () => trace({ writeAborted: String(transaction.error?.name ?? "abort") });
      const store = transaction.objectStore(STORE);
      const at = Date.now();
      for (const [key, boxes] of batch) store.put({ key, boxes, at } satisfies BoxRecord);
      const count = store.count();
      count.onsuccess = () => {
        if (count.result <= MAX_ENTRIES) return;
        // Bounded: the oldest beyond MAX_ENTRIES go.
        const all = store.getAll();
        all.onsuccess = () => {
          const records = (all.result as BoxRecord[]).sort((a, b) => b.at - a.at);
          records.slice(MAX_ENTRIES).forEach((record) => store.delete(record.key));
        };
      };
    } catch {
      // Best-effort: the next opening measures again.
    }
  });
}

/** Keep `boxes` under `key` (written a few seconds later, in one batch). */
export function keepMathGeometry(key: string, boxes: unknown): void {
  trace({ kept: 1 });
  memory.set(key, boxes);
  pending.set(key, boxes);
  if (flushTimer === null && typeof setTimeout === "function") flushTimer = setTimeout(flush, FLUSH_DELAY_MS);
}

/** Boxes measured under a typesetting found stale (a formula face that
 * arrived late): those not written yet are dropped, those written go. */
export function forgetMathGeometry(keys: Iterable<string>): void {
  const gone = [...keys];
  trace({ forgot: gone.length });
  for (const key of gone) {
    memory.delete(key);
    pending.delete(key);
  }
  if (gone.length === 0) return;
  void open().then((db) => {
    if (!db) return;
    try {
      const store = db.transaction(STORE, "readwrite").objectStore(STORE);
      for (const key of gone) store.delete(key);
    } catch {
      // Best-effort.
    }
  });
}

/** Logout or session expiry: nothing of the account stays behind. */
export function clearMathGeometry(): void {
  clearFormulaBoxes();
  memory.clear();
  pending.clear();
  loaded = null;
  void open().then((db) => {
    if (!db) return;
    try {
      db.transaction(STORE, "readwrite").objectStore(STORE).clear();
    } catch {
      // Best-effort.
    }
  });
}

/** A 53-bit hash of `value` (cyrb53), as a number. */
export function mathGeometryHashNumber(value: string): number {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < value.length; i += 1) {
    const ch = value.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return 4294967296 * (2097151 & h2) + (h1 >>> 0);
}

/** {@link mathGeometryHashNumber} as a base-36 string. */
export function mathGeometryHash(value: string): string {
  return mathGeometryHashNumber(value).toString(36);
}

// ---- One formula's box, wherever and in whichever document it stands -------

/**
 * The typeset box of one formula, by what lays it out and nothing about where
 * it stands (ledgerFormulaBoxes makes the key): a book repeats its formulas (4,573
 * different ones among 7,750), and the next document repeats them again.
 *
 * Kept as one record per typesetter (ledgerFormulaBoxes.typesetterOf): two
 * typed arrays, read with one `get` and written whole. A record per formula
 * cost a `getAll` of tens of thousands of objects at every opening.
 */
export interface FormulaBox {
  readonly width: number;
  readonly height: number;
  readonly baselineDepth: number;
  readonly kind: "inline" | "block";
  /** KaTeX left places to break it: whether it does depends on its line. */
  readonly breaks: boolean;
  /** An in-line formula on Word's box that fits one slot of the line grid:
   * it leaves its line to the words (note-editor.css), so what it adds to the
   * line is nothing, and its height of 0 is its box, not a reading of nothing. */
  readonly leavesLine?: boolean;
}

interface FormulaShard {
  key: string;
  hashes: Float64Array;
  /** width, height, baselineDepth, flags for each hash. */
  boxes: Float64Array;
  at: number;
}

const FORMULA_DB_NAME = "mrnote-docs-math-formula-v1";
const FORMULA_STORE = "shards";
const FORMULA_FIELDS = 4;
const FORMULA_BLOCK_FLAG = 1;
const FORMULA_BREAKS_FLAG = 2;
const FORMULA_LEAVES_LINE_FLAG = 4;
/** The most formulas one typesetter keeps; the oldest go first. */
const MAX_FORMULAS = 100_000;
/** The typesetters a browser keeps (two zoom levels, a browser update). */
const MAX_FORMULA_SHARDS = 3;

let formulaDatabase: Promise<IDBDatabase | null> | null = null;
let formulaTypesetter: string | null = null;
let formulaLoaded: Promise<void> | null = null;
const formulaMemory = new Map<number, FormulaBox>();
const formulaPending = new Map<number, FormulaBox>();
const formulaForgotten = new Set<number>();
let formulaFlushTimer: ReturnType<typeof setTimeout> | null = null;

function openFormulas(): Promise<IDBDatabase | null> {
  if (formulaDatabase) return formulaDatabase;
  formulaDatabase = new Promise((resolve) => {
    try {
      if (typeof indexedDB === "undefined") {
        resolve(null);
        return;
      }
      const request = indexedDB.open(FORMULA_DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(FORMULA_STORE)) db.createObjectStore(FORMULA_STORE, { keyPath: "key" });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return formulaDatabase;
}

/** Whether a signed-in principal keeps formula boxes (a guest's last the page). */
export function formulaBoxesCanBeKept(): boolean {
  return principal() !== null;
}

function shardKey(typesetter: string): string | null {
  const who = principal();
  return who ? `${who}|${typesetter}` : null;
}

/** Read the formulas kept under `typesetter` into memory, once. */
export function preloadFormulaBoxes(typesetter: string): Promise<void> {
  if (formulaLoaded && formulaTypesetter === typesetter) return formulaLoaded;
  formulaTypesetter = typesetter;
  formulaMemory.clear();
  formulaPending.clear();
  formulaForgotten.clear();
  const key = shardKey(typesetter);
  if (!key) return Promise.resolve();
  formulaLoaded = openFormulas().then(
    (db) =>
      new Promise<void>((resolve) => {
        if (!db) {
          resolve();
          return;
        }
        try {
          const request = db.transaction(FORMULA_STORE, "readonly").objectStore(FORMULA_STORE).get(key);
          request.onsuccess = () => {
            const shard = request.result as FormulaShard | undefined;
            if (shard && Date.now() - shard.at < MAX_AGE_MS && formulaTypesetter === typesetter) {
              const { hashes, boxes } = shard;
              for (let at = 0; at < hashes.length; at += 1) {
                // A formula measured since (or forgotten since) stands.
                if (formulaMemory.has(hashes[at]) || formulaForgotten.has(hashes[at])) continue;
                const flags = boxes[at * FORMULA_FIELDS + 3];
                formulaMemory.set(hashes[at], {
                  width: boxes[at * FORMULA_FIELDS],
                  height: boxes[at * FORMULA_FIELDS + 1],
                  baselineDepth: boxes[at * FORMULA_FIELDS + 2],
                  kind: flags & FORMULA_BLOCK_FLAG ? "block" : "inline",
                  breaks: (flags & FORMULA_BREAKS_FLAG) !== 0,
                  ...((flags & FORMULA_LEAVES_LINE_FLAG) !== 0 ? { leavesLine: true } : {}),
                });
              }
            }
            trace({ formulasPreloaded: formulaMemory.size });
            resolve();
          };
          request.onerror = () => resolve();
        } catch {
          resolve();
        }
      }),
  );
  return formulaLoaded;
}

/** The box kept for formula `hash`, if it has been read. */
export function keptFormulaBox(hash: number): FormulaBox | undefined {
  const box = formulaMemory.get(hash);
  // One without a box (kept before such were refused) is not used; one that
  // leaves its line has the height of none (FormulaBox.leavesLine).
  return box && box.width > 0 && (box.height > 0 || (box.leavesLine === true && box.height === 0 && box.baselineDepth === 0)) ? box : undefined;
}

function flushFormulas(): void {
  formulaFlushTimer = null;
  const typesetter = formulaTypesetter;
  const key = typesetter ? shardKey(typesetter) : null;
  if (!key || (formulaPending.size === 0 && formulaForgotten.size === 0)) return;
  const added = [...formulaPending];
  const gone = new Set(formulaForgotten);
  formulaPending.clear();
  formulaForgotten.clear();
  void openFormulas().then((db) => {
    if (!db) return;
    try {
      const transaction = db.transaction(FORMULA_STORE, "readwrite");
      transaction.oncomplete = () => trace({ formulasWritten: added.length });
      transaction.onerror = () => trace({ formulasWriteFailed: String(transaction.error?.name ?? "error") });
      transaction.onabort = () => trace({ formulasWriteAborted: String(transaction.error?.name ?? "abort") });
      const store = transaction.objectStore(FORMULA_STORE);
      // Read and written in one transaction: another tab's formulas stay.
      const read = store.get(key);
      read.onsuccess = () => {
        const shard = read.result as FormulaShard | undefined;
        const hashes: number[] = [];
        const boxes: number[] = [];
        const incoming = new Set(added.map(([hash]) => hash));
        if (shard) {
          for (let at = 0; at < shard.hashes.length; at += 1) {
            const hash = shard.hashes[at];
            if (gone.has(hash) || incoming.has(hash)) continue;
            hashes.push(hash);
            for (let field = 0; field < FORMULA_FIELDS; field += 1) boxes.push(shard.boxes[at * FORMULA_FIELDS + field]);
          }
        }
        for (const [hash, box] of added) {
          hashes.push(hash);
          boxes.push(
            box.width, box.height, box.baselineDepth,
            (box.kind === "block" ? FORMULA_BLOCK_FLAG : 0) | (box.breaks ? FORMULA_BREAKS_FLAG : 0) |
              (box.leavesLine ? FORMULA_LEAVES_LINE_FLAG : 0),
          );
        }
        const over = Math.max(0, hashes.length - MAX_FORMULAS);
        store.put({
          key,
          hashes: Float64Array.from(over ? hashes.slice(over) : hashes),
          boxes: Float64Array.from(over ? boxes.slice(over * FORMULA_FIELDS) : boxes),
          at: Date.now(),
        } satisfies FormulaShard);
      };
      // A typesetter past (an update of the browser, a zoom left behind).
      const keys = store.getAllKeys();
      keys.onsuccess = () => {
        const others = (keys.result as string[]).filter((other) => other !== key);
        if (others.length >= MAX_FORMULA_SHARDS) others.forEach((other) => store.delete(other));
      };
    } catch {
      // Best-effort: the next opening measures again.
    }
  });
}

/** Keep these formulas' boxes (written a few seconds later, in one batch). */
export function keepFormulaBoxes(entries: Iterable<readonly [number, FormulaBox]>): void {
  if (!formulaTypesetter) return;
  let kept = 0;
  for (const [hash, box] of entries) {
    formulaMemory.set(hash, box);
    formulaPending.set(hash, box);
    formulaForgotten.delete(hash);
    kept += 1;
  }
  trace({ formulasKept: kept });
  if (kept > 0 && formulaFlushTimer === null && typeof setTimeout === "function") {
    formulaFlushTimer = setTimeout(flushFormulas, FLUSH_DELAY_MS);
  }
}

/** Boxes measured under a typesetting found stale (see forgetMathGeometry). */
export function forgetFormulaBoxes(hashes: Iterable<number>): void {
  let forgot = 0;
  for (const hash of hashes) {
    formulaMemory.delete(hash);
    formulaPending.delete(hash);
    formulaForgotten.add(hash);
    forgot += 1;
  }
  trace({ formulasForgot: forgot });
  if (forgot > 0 && formulaFlushTimer === null && typeof setTimeout === "function") {
    formulaFlushTimer = setTimeout(flushFormulas, FLUSH_DELAY_MS);
  }
}

function clearFormulaBoxes(): void {
  formulaMemory.clear();
  formulaPending.clear();
  formulaForgotten.clear();
  formulaLoaded = null;
  formulaTypesetter = null;
  void openFormulas().then((db) => {
    if (!db) return;
    try {
      db.transaction(FORMULA_STORE, "readwrite").objectStore(FORMULA_STORE).clear();
    } catch {
      // Best-effort.
    }
  });
}
