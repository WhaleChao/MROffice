/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Office reference tables the editors read, held for the page.
 *
 * Tables we measured or reconstructed from real Office (PowerPoint's and
 * Word's table style galleries, Word's synthetic-bold advances, Word's style
 * sets, the font registry and the embedding manifest) are not part of the
 * browser bundle:
 * the API serves them to a signed-in user (`GET /api/v1/office-data/{name}`,
 * app/routers/office_data.py) and this store holds what arrived. A guest is
 * served the font registry's mapping alone ({@link officeDataServedToGuests}:
 * which open face and slice files draw an Office family), so a visitor's
 * deck registers its fonts as an account's does; the measured tables stay an
 * account's, and a guest's readers of those draw their fallback. The tables
 * themselves live in the files `lib/office-data/datasets.json` names, which
 * `scripts/check-production-bundle.mjs --source` keeps out of the browser's
 * module graph.
 *
 * Readers stay synchronous: {@link officeData} answers what has arrived (or
 * `undefined`, and the caller draws its fallback), {@link loadOfficeData}
 * asks once and shares the request, and {@link subscribeOfficeData} tells a
 * surface to draw again when a table lands. An editor prefetches the tables
 * it draws with when it opens (and the workspace after sign-in), so a page's
 * first layout normally has them.
 *
 * A host that holds the tables itself (the engine host, the editor-command
 * host, unit tests) installs a synchronous source
 * (`lib/office-data/server-tables.ts`), and every read answers at once.
 *
 * Dependency-free on purpose: the transport (`./transport`, the API client)
 * is loaded on first use, so a test setup can install a source before any
 * app module loads.
 */

export type OfficeDataWholeName =
  | "presentation-table-styles"
  | "word-table-styles"
  | "word-synthetic-bold-advance"
  | "word-style-sets";

/** Tables served by part (one font family at a time): too large to hand
 * over whole, and a page needs only the families it names. */
export type OfficeDataPartitionedName = "font-registry" | "font-embed-manifest";

export type OfficeDataName = OfficeDataWholeName | OfficeDataPartitionedName;

export const OFFICE_DATA_NAMES: readonly OfficeDataName[] = Object.freeze([
  "presentation-table-styles",
  "word-table-styles",
  "word-synthetic-bold-advance",
  "word-style-sets",
  "font-registry",
  "font-embed-manifest",
]);

const PARTITIONED: ReadonlySet<OfficeDataName> = new Set<OfficeDataName>(["font-registry", "font-embed-manifest"]);

/** Tables the API also answers a guest (no session): the font registry, as
 * its mapping (app/services/office_data.py office_data_guest_answer). */
const SERVED_TO_GUESTS: ReadonlySet<OfficeDataName> = new Set<OfficeDataName>(["font-registry"]);

/** Whether a page without a session is answered this table too. */
export function officeDataServedToGuests(name: OfficeDataName): boolean {
  return SERVED_TO_GUESTS.has(name);
}

/** A table as the files hold it (server-side hosts, tests). */
export type OfficeDataSource = (name: OfficeDataName) => unknown;

/** Why the browser has no table: a guest, no connection, or a refusal. */
export type OfficeDataUnavailableReason = "signed_out" | "offline" | "rate_limited" | "failed";

export class OfficeDataUnavailableError extends Error {
  readonly reason: OfficeDataUnavailableReason;
  constructor(reason: OfficeDataUnavailableReason, detail?: string) {
    super(detail ? `office data ${reason}: ${detail}` : `office data ${reason}`);
    this.name = "OfficeDataUnavailableError";
    this.reason = reason;
  }
}

/** How the browser asks the API (`./transport`). */
export interface OfficeDataTransport {
  /** Whether a request can succeed at all (a signed-in page; a guest only
   * for {@link officeDataServedToGuests}). */
  signedIn(): boolean;
  whole(name: OfficeDataWholeName): Promise<unknown>;
  /** The fragment holding these parts, in the table's own shape. */
  parts(name: OfficeDataPartitionedName, keys: readonly string[]): Promise<Record<string, unknown>>;
}

/** At most this many parts in one request (the API refuses more). */
export const OFFICE_DATA_PARTS_PER_REQUEST = 64;
/** A failed table is asked for again after this long. */
const RETRY_AFTER_MS = 30_000;

interface WholeEntry {
  value?: unknown;
  pending?: Promise<unknown>;
  failedAt?: number;
}

interface PartitionedEntry {
  /** Every fragment that arrived, merged, in the table's own shape. */
  merged: Record<string, unknown>;
  /** Parts asked for and answered (present in `merged` or unknown to the table). */
  known: Set<string>;
  pending: Map<string, Promise<void>>;
  failedAt: Map<string, number>;
  /** Parts waiting for this turn's request. */
  queued: Map<string, { resolve: () => void }>;
  flushScheduled: boolean;
}

let source: OfficeDataSource | null = null;
let transport: OfficeDataTransport | null = null;
let transportLoading: Promise<OfficeDataTransport> | null = null;
const whole = new Map<OfficeDataWholeName, WholeEntry>();
const partitioned = new Map<OfficeDataPartitionedName, PartitionedEntry>();
const sourceValues = new Map<OfficeDataName, unknown>();
const listeners = new Set<() => void>();
let version = 0;
const versions = new Map<OfficeDataName, number>();

function notify(name?: OfficeDataName): void {
  version += 1;
  for (const each of name ? [name] : OFFICE_DATA_NAMES) versions.set(each, (versions.get(each) ?? 0) + 1);
  for (const listener of [...listeners]) {
    try {
      listener();
    } catch {
      // One reader's failure never stops the others from redrawing.
    }
  }
}

function fromSource(name: OfficeDataName): unknown {
  if (!source) return undefined;
  if (sourceValues.has(name)) return sourceValues.get(name);
  const value = source(name);
  if (value !== undefined) sourceValues.set(name, value);
  return value;
}

async function currentTransport(): Promise<OfficeDataTransport> {
  if (transport) return transport;
  transportLoading ??= import("./transport").then((module) => {
    transport ??= module.officeDataApiTransport;
    return transport;
  });
  return transportLoading;
}

/** The signed-in check the transport answers, before it has loaded: the
 * `auth_state` cookie `lib/auth-state.ts` keeps (a guest asks nothing). */
function signedInNow(): boolean {
  if (transport) return transport.signedIn();
  if (typeof document === "undefined") return false;
  return /(?:^|;\s*)auth_state=1(?:;|$)/.test(document.cookie ?? "");
}

function isPartitioned(name: OfficeDataName): name is OfficeDataPartitionedName {
  return PARTITIONED.has(name);
}

/** A whole table if it is here, else `undefined` (draw the fallback). */
export function officeData(name: OfficeDataWholeName): unknown {
  const local = fromSource(name);
  if (local !== undefined) return local;
  return whole.get(name)?.value;
}

/**
 * A whole table, asked for once: resolves to the table, or to `undefined`
 * when this page cannot have it (a guest, offline, refused). Never rejects.
 */
export function loadOfficeData(name: OfficeDataWholeName): Promise<unknown> {
  const local = fromSource(name);
  if (local !== undefined) return Promise.resolve(local);
  let entry = whole.get(name);
  if (!entry) whole.set(name, (entry = {}));
  if (entry.value !== undefined) return Promise.resolve(entry.value);
  if (entry.pending) return entry.pending;
  if (entry.failedAt !== undefined && Date.now() - entry.failedAt < RETRY_AFTER_MS) return Promise.resolve(undefined);
  // A guest asks nothing; signing in later asks then.
  if (!signedInNow()) return Promise.resolve(undefined);
  const current = entry;
  current.failedAt = undefined;
  current.pending = currentTransport()
    .then((client) => client.whole(name))
    .then(
      (value) => {
        current.pending = undefined;
        if (value === undefined || value === null) {
          current.failedAt = Date.now();
          notify(name);
          return undefined;
        }
        current.value = value;
        notify(name);
        return value;
      },
      () => {
        current.pending = undefined;
        current.failedAt = Date.now();
        notify(name);
        return undefined;
      },
    );
  return current.pending;
}

function partitionedEntry(name: OfficeDataPartitionedName): PartitionedEntry {
  let entry = partitioned.get(name);
  if (!entry) {
    entry = {
      merged: {},
      known: new Set(),
      pending: new Map(),
      failedAt: new Map(),
      queued: new Map(),
      flushScheduled: false,
    };
    partitioned.set(name, entry);
  }
  return entry;
}

/** Merge one fragment: an object-valued field gains the fragment's keys; any
 * other field (a version) takes the fragment's value. */
function mergeFragment(target: Record<string, unknown>, fragment: Record<string, unknown>): void {
  for (const [key, value] of Object.entries(fragment)) {
    if (value && typeof value === "object" && !Array.isArray(value)) {
      const existing = target[key];
      target[key] = Object.assign(
        existing && typeof existing === "object" && !Array.isArray(existing) ? (existing as Record<string, unknown>) : {},
        value as Record<string, unknown>,
      );
    } else {
      target[key] = value;
    }
  }
}

function flushParts(name: OfficeDataPartitionedName, entry: PartitionedEntry): void {
  entry.flushScheduled = false;
  const queued = [...entry.queued.entries()];
  entry.queued.clear();
  if (!queued.length) return;
  for (let start = 0; start < queued.length; start += OFFICE_DATA_PARTS_PER_REQUEST) {
    const batch = queued.slice(start, start + OFFICE_DATA_PARTS_PER_REQUEST);
    const keys = batch.map(([key]) => key).sort();
    void currentTransport()
      .then((client) => client.parts(name, keys))
      .then(
        (fragment) => {
          mergeFragment(entry.merged, fragment);
          for (const key of keys) {
            entry.known.add(key);
            entry.failedAt.delete(key);
          }
        },
        () => {
          const now = Date.now();
          for (const key of keys) entry.failedAt.set(key, now);
        },
      )
      .finally(() => {
        for (const [key, waiter] of batch) {
          entry.pending.delete(key);
          waiter.resolve();
        }
        notify(name);
      });
  }
}

/**
 * The table holding at least these parts (font families, lower-cased),
 * asked for together: parts requested in the same turn travel in one
 * request. Resolves to everything this page has of the table (in its own
 * shape), or `undefined` when it has none of it. Never rejects.
 */
export async function loadOfficeDataParts(
  name: OfficeDataPartitionedName,
  keys: readonly string[],
): Promise<Record<string, unknown> | undefined> {
  const local = fromSource(name);
  if (local !== undefined) return local as Record<string, unknown>;
  const entry = partitionedEntry(name);
  const now = Date.now();
  const waits: Promise<void>[] = [];
  const askable = signedInNow() || officeDataServedToGuests(name);
  for (const raw of keys) {
    const key = raw.trim().toLowerCase();
    if (!key || entry.known.has(key)) continue;
    const pending = entry.pending.get(key);
    if (pending) {
      waits.push(pending);
      continue;
    }
    const failedAt = entry.failedAt.get(key);
    if (failedAt !== undefined && now - failedAt < RETRY_AFTER_MS) continue;
    if (!askable) continue;
    let resolve!: () => void;
    const wait = new Promise<void>((done) => (resolve = done));
    entry.pending.set(key, wait);
    entry.queued.set(key, { resolve });
    waits.push(wait);
  }
  if (entry.queued.size && !entry.flushScheduled) {
    entry.flushScheduled = true;
    queueMicrotask(() => flushParts(name, entry));
  }
  if (waits.length) await Promise.all(waits);
  return Object.keys(entry.merged).length ? entry.merged : undefined;
}

/** Whether a part has been answered (present, or unknown to the table). */
export function officeDataPartKnown(name: OfficeDataPartitionedName, key: string): boolean {
  if (fromSource(name) !== undefined) return true;
  return partitioned.get(name)?.known.has(key.trim().toLowerCase()) ?? false;
}

/** Whether this page can ask the API at all (signed in, or a host that holds
 * the tables): a guest's readers settle for their fallback for good. With a
 * name, whether it can ask for that table (a guest can for
 * {@link officeDataServedToGuests}). */
export function officeDataObtainable(name?: OfficeDataName): boolean {
  return source !== null || signedInNow() || (name !== undefined && officeDataServedToGuests(name));
}

/** What has arrived of a partitioned table (`undefined` when nothing has). */
export function officeDataParts(name: OfficeDataPartitionedName): Record<string, unknown> | undefined {
  const local = fromSource(name);
  if (local !== undefined) return local as Record<string, unknown>;
  const merged = partitioned.get(name)?.merged;
  return merged && Object.keys(merged).length ? merged : undefined;
}

/**
 * Whether every named whole table has been answered (arrived, or this page
 * cannot have it): what an editor waits for before it reveals its first
 * layout. A guest is settled at once.
 */
export function officeDataSettled(names: readonly OfficeDataWholeName[]): boolean {
  return names.every((name) => {
    if (fromSource(name) !== undefined) return true;
    const entry = whole.get(name);
    if (entry?.value !== undefined) return true;
    if (entry?.pending) return false;
    if (entry?.failedAt !== undefined) return true;
    return !signedInNow();
  });
}

/** Ask for these tables now (an editor opening, the workspace after sign-in). */
export function prefetchOfficeData(names: readonly OfficeDataName[]): void {
  for (const name of names) {
    if (isPartitioned(name)) continue;
    void loadOfficeData(name);
  }
}

export function subscribeOfficeData(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Advances every time a table (or a part of one) arrives or fails; with a
 * name, only for that table. */
export function officeDataVersion(name?: OfficeDataName): number {
  return name ? (versions.get(name) ?? 0) : version;
}

/**
 * Hosts that hold the tables (the engine host, the editor-command host, unit
 * tests) answer every read from them; `null` removes the source.
 */
export function installOfficeDataSource(next: OfficeDataSource | null): void {
  source = next;
  sourceValues.clear();
  notify();
}

/** A transport in place of the API client: tests, and the page renderer's
 * browser (`./print-transport`, the tables its driver serves). */
export function setOfficeDataTransport(next: OfficeDataTransport | null): void {
  transport = next;
  transportLoading = null;
}

/** Forget everything this page holds (tests; a sign-out). */
export function resetOfficeData(): void {
  whole.clear();
  partitioned.clear();
  sourceValues.clear();
  notify();
}
