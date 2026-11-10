/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The data WEBSERVICE and IMPORTRANGE read, as this browser tab has it
 * (`external-data`), fetched from the page's server -- never from the source
 * itself, as the IMAGE pictures are (`web-image-store`):
 *
 * - WEBSERVICE: `POST /api/v1/office/sheets/web-services` fetches the URL
 *   (SSRF-guarded, redirects refused, text only, capped) for a host the
 *   member allowed; any other host is #BLOCKED! and the sheet asks once per
 *   host per tab ("Allow" is kept on the server for the member; a separate
 *   consent from IMAGE's: allowing pictures from a host does not let formulas
 *   read data from it). An answer that cannot be returned is #VALUE!.
 * - IMPORTRANGE: `POST /api/v1/office/sheets/import-range` reads the range of
 *   another spreadsheet page of this workspace with the reader's own
 *   permissions; anything the reader may not read is #REF!.
 *
 * The calculation session names the calls it read without data
 * (`externalRequests`); each is fetched once per tab ("loading" meanwhile:
 * the cell shows #BUSY!) and handed to that page's session with its next
 * calculation. Calculate Now fetches everything again, the old answers
 * standing until the new ones arrive.
 */

import { ApiRequestError, apiPost } from "@mrnote/core/web/lib/api.ts";
import {
  MAX_CALCULATION_EXTERNAL_CHARS,
  MAX_CALCULATION_EXTERNAL_DATA,
  MAX_EXTERNAL_TEXT,
  MAX_IMPORT_RANGE_CELLS,
  parseExternalDataKey,
  type CalculationExternalDatum,
  type ExternalCellWire,
  type ExternalDatumWire,
} from "@mrnote/core/web/lib/spreadsheet/external-data.ts";

export const SPREADSHEET_WEB_SERVICE_PATH = "/api/v1/office/sheets/web-services";
export const SPREADSHEET_WEB_SERVICE_HOSTS_PATH = "/api/v1/office/sheets/web-services/hosts";
export const SPREADSHEET_IMPORT_RANGE_PATH = "/api/v1/office/sheets/import-range";

type Fetcher = (key: string, pageId: string | null) => Promise<ExternalDatumWire>;

interface Entry {
  readonly status: "loading" | "ready";
  /** The answer (while loading: the previous one, kept until the new one comes). */
  readonly data?: ExternalDatumWire;
  /** Which answer this is (the list handed to a session names it). */
  readonly answer?: number;
}

/** The most calls one tab keeps. */
const MAX_ENTRIES = 512;
const entries = new Map<string, Entry>();
/** Per page (the session that asked), the keys it read. */
const requestedBy = new Map<string, Set<string>>();
/** A refused call's stand-in per page: the text its cell was saved with (X-330). */
const storedBy = new Map<string, Map<string, string>>();
const listeners = new Set<() => void>();
let revision = 0;

let answers = 0;

const scopeOf = (pageId: string | null) => pageId ?? "";
/** WEBSERVICE answers serve every page; IMPORTRANGE ones the page they were read for. */
function entryKey(key: string, pageId: string | null): string {
  return parseExternalDataKey(key)?.kind === "importrange" ? `${scopeOf(pageId)}\u0001${key}` : `\u0001${key}`;
}

function notify(): void {
  revision += 1;
  listCache.clear();
  for (const listener of [...listeners]) listener();
}

function errorCode(error: unknown): string | undefined {
  const details = error instanceof ApiRequestError ? (error.details as { code?: unknown } | undefined) : undefined;
  return typeof details?.code === "string" ? details.code : undefined;
}

function checkedGrid(value: unknown): ExternalDatumWire | null {
  if (!Array.isArray(value) || value.length === 0 || !Array.isArray(value[0])) return null;
  const width = value[0].length;
  if (width === 0 || value.length * width > MAX_IMPORT_RANGE_CELLS) return null;
  const grid: ExternalCellWire[][] = [];
  for (const line of value) {
    if (!Array.isArray(line) || line.length !== width) return null;
    const row: ExternalCellWire[] = [];
    for (const cell of line) {
      if (cell === null || typeof cell === "boolean" || (typeof cell === "number" && Number.isFinite(cell))) row.push(cell);
      else if (typeof cell === "string" && cell.length <= MAX_EXTERNAL_TEXT) row.push(cell);
      else if (typeof cell === "object" && typeof (cell as { error?: unknown }).error === "string") row.push({ error: (cell as { error: string }).error });
      else return null;
    }
    grid.push(row);
  }
  return { grid };
}

async function fetchFromServer(key: string, pageId: string | null): Promise<ExternalDatumWire> {
  const call = parseExternalDataKey(key);
  if (!call) return { error: "#VALUE!" };
  if (call.kind === "webservice") {
    try {
      const response = await apiPost<{ text?: unknown }>(SPREADSHEET_WEB_SERVICE_PATH, { url: call.url });
      return typeof response.text === "string" && response.text.length <= MAX_EXTERNAL_TEXT
        ? { text: response.text }
        : { error: "#VALUE!" };
    } catch (error) {
      const code = errorCode(error);
      const details = error instanceof ApiRequestError ? (error.details as { reason?: unknown; host?: unknown } | undefined) : undefined;
      if (code === "#BLOCKED!" && details?.reason === "host_not_approved" && typeof details.host === "string") {
        noteHostQuestion(key, details.host);
      }
      // Excel: anything that cannot be returned is #VALUE!.
      return { error: code === "#BLOCKED!" ? "#BLOCKED!" : "#VALUE!" };
    }
  }
  if (!pageId) return { error: "#REF!" };
  try {
    const response = await apiPost<{ grid?: unknown }>(SPREADSHEET_IMPORT_RANGE_PATH, {
      page_id: pageId,
      source_page_id: call.pageId,
      range: call.range,
    });
    return checkedGrid(response.grid) ?? { error: "#REF!" };
  } catch (error) {
    return { error: errorCode(error) === "#VALUE!" ? "#VALUE!" : "#REF!" };
  }
}

let fetcher: Fetcher = fetchFromServer;

/** Fetch one call; resolves true when its answer differs from `previous`. */
function start(key: string, pageId: string | null, previous?: ExternalDatumWire): Promise<boolean> {
  const at = entryKey(key, pageId);
  const before = entries.get(at);
  const entry: Entry = previous ? { status: "loading", data: previous, ...(before?.answer ? { answer: before.answer } : {}) } : { status: "loading" };
  entries.set(at, entry);
  const settle = (data: ExternalDatumWire): boolean => {
    // Forgotten or fetched again meanwhile: this answer is dropped.
    if (entries.get(at) !== entry) return false;
    answers += 1;
    entries.set(at, { status: "ready", data, answer: answers });
    notify();
    return JSON.stringify(data) !== JSON.stringify(previous);
  };
  return fetcher(key, pageId).then(settle, () => settle({ error: parseExternalDataKey(key)?.kind === "importrange" ? "#REF!" : "#VALUE!" }));
}

/**
 * The calls a page's calculation read without data: each is fetched unless
 * this tab already has (or is getting) it.
 */
export function requestSpreadsheetExternalData(
  pageId: string | null,
  keys: readonly string[],
  stored?: ReadonlyMap<string, string>,
): void {
  let changed = false;
  const asked = requestedBy.get(scopeOf(pageId)) ?? new Set<string>();
  requestedBy.set(scopeOf(pageId), asked);
  if (stored?.size) {
    const kept = storedBy.get(scopeOf(pageId)) ?? new Map<string, string>();
    storedBy.set(scopeOf(pageId), kept);
    for (const [key, text] of stored) {
      if (kept.get(key) === text) continue;
      kept.set(key, text);
      changed = true;
    }
  }
  for (const key of keys) {
    if (!parseExternalDataKey(key)) continue;
    if (!asked.has(key)) {
      asked.add(key);
      changed = true;
    }
    if (entries.has(entryKey(key, pageId))) continue;
    if (entries.size >= MAX_ENTRIES) continue;
    void start(key, pageId);
  }
  // A key another page fetched already: this page's list now names it.
  if (changed) notify();
}

const listCache = new Map<string, { key: string; data: readonly CalculationExternalDatum[] }>();
/** The last list each page was handed: kept while its key holds, so a change
 * elsewhere hands the page the very same list (no calculation again). */
const lastLists = new Map<string, { key: string; data: readonly CalculationExternalDatum[] }>();

/**
 * What a page's session is handed: the data of the calls it read, as this
 * tab holds it, bounded as the session bounds it; `key` changes with it.
 */
export function spreadsheetExternalDataFor(pageId: string | null): { key: string; data: readonly CalculationExternalDatum[] } {
  const scope = scopeOf(pageId);
  const cached = listCache.get(scope);
  if (cached) return cached;
  const data: CalculationExternalDatum[] = [];
  const named: string[] = [];
  let chars = 0;
  for (const key of [...(requestedBy.get(scope) ?? [])].sort()) {
    if (data.length >= MAX_CALCULATION_EXTERNAL_DATA) break;
    const fetched = entries.get(entryKey(key, pageId))?.data;
    if (!fetched) continue;
    // A call this server refuses shows the text its cell was saved with, as
    // Excel for Mac shows a workbook's cached WEBSERVICE result (X-330).
    const standIn = "error" in fetched && fetched.error === "#BLOCKED!" ? storedBy.get(scope)?.get(key) : undefined;
    const held: ExternalDatumWire = standIn !== undefined ? { text: standIn } : fetched;
    const size = "text" in held ? held.text.length : "grid" in held ? JSON.stringify(held.grid).length : 0;
    if (chars + size > MAX_CALCULATION_EXTERNAL_CHARS) continue;
    chars += size;
    data.push([key, held]);
    named.push(`${entries.get(entryKey(key, pageId))?.answer ?? 0}${standIn !== undefined ? "s" : ""}`);
  }
  // Changes only when this page's data does (another page's fetch leaves it).
  const list = { key: data.length ? named.join(",") : "", data };
  const previous = lastLists.get(scope);
  if (previous && previous.key === list.key) {
    listCache.set(scope, previous);
    return previous;
  }
  lastLists.set(scope, list);
  listCache.set(scope, list);
  return list;
}

/** Hear every change; returns the unsubscribe. */
export function subscribeSpreadsheetExternalData(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Changes whenever an answer arrives or a page asks for another call. */
export function spreadsheetExternalDataRevision(): number {
  return revision;
}

/**
 * Calculate Now: every call is fetched again; the old answers stand until the
 * new ones arrive. Resolves true once they have, when any answer changed (the
 * caller calculates again then).
 */
export function refreshSpreadsheetExternalData(): Promise<boolean> {
  const fetches: Array<Promise<boolean>> = [];
  for (const [at, entry] of [...entries]) {
    if (entry.status === "loading") continue;
    const split = at.indexOf("\u0001");
    const scope = at.slice(0, split);
    fetches.push(start(at.slice(split + 1), scope || null, entry.data));
  }
  if (!fetches.length) return Promise.resolve(false);
  notify();
  return Promise.all(fetches).then((changed) => changed.some(Boolean));
}

// ── Per-host consent (WEBSERVICE) ─────────────────────────────────────────
// A URL can be computed from cells (and the AI writes formulas), so it can
// carry workbook text to whoever runs the host. The server answers #BLOCKED!
// for a host the member has not allowed; the sheet asks once per host per
// tab. "Allow" is recorded on the server (per member + host); "Don't allow"
// lasts for this tab.

const hostKeys = new Map<string, Set<string>>();
const deniedHosts = new Set<string>();
let hostQuestions: readonly string[] = [];

function noteHostQuestion(key: string, host: string): void {
  const keys = hostKeys.get(host) ?? new Set<string>();
  keys.add(key);
  hostKeys.set(host, keys);
  if (!deniedHosts.has(host) && !hostQuestions.includes(host)) {
    hostQuestions = [...hostQuestions, host];
    queueMicrotask(notify);
  }
}

/** The hosts this tab is waiting for the member to allow or refuse, in order. */
export function spreadsheetWebServiceHostQuestions(): readonly string[] {
  return hostQuestions;
}

/** The member allows formulas to read data from ``host``: record it, then fetch its calls again. */
export async function allowSpreadsheetWebServiceHost(host: string): Promise<void> {
  await apiPost(SPREADSHEET_WEB_SERVICE_HOSTS_PATH, { host });
  hostQuestions = hostQuestions.filter((candidate) => candidate !== host);
  for (const key of hostKeys.get(host) ?? []) {
    const at = entryKey(key, null);
    if (entries.get(at)?.status === "ready") void start(key, null);
  }
  hostKeys.delete(host);
  notify();
}

/** The member refuses ``host`` for this tab: its calls keep #BLOCKED!. */
export function denySpreadsheetWebServiceHost(host: string): void {
  deniedHosts.add(host);
  hostQuestions = hostQuestions.filter((candidate) => candidate !== host);
  notify();
}

/**
 * A printed workbook's data (the print page holds no session to ask the
 * server with): the server fetched the literal WEBSERVICE / IMPORTRANGE
 * calls it may, as the printing member, and hands each answer in by key; any
 * other call reads as #BLOCKED! (WEBSERVICE) or #REF! (IMPORTRANGE).
 * Forgets every entry.
 */
export function answerSpreadsheetExternalDataFrom(answers: Readonly<Record<string, unknown>>): void {
  entries.clear();
  requestedBy.clear();
  notify();
  fetcher = async (key) => {
    const answer = Object.prototype.hasOwnProperty.call(answers, key) ? answers[key] : undefined;
    const fallback: ExternalDatumWire = { error: parseExternalDataKey(key)?.kind === "importrange" ? "#REF!" : "#BLOCKED!" };
    if (typeof answer !== "object" || answer === null) return fallback;
    const record = answer as Record<string, unknown>;
    if (typeof record.text === "string" && record.text.length <= MAX_EXTERNAL_TEXT) return { text: record.text };
    if (typeof record.error === "string") return { error: record.error };
    return checkedGrid(record.grid) ?? fallback;
  };
}

/** Tests: replace the server fetch and forget every entry. */
export function resetSpreadsheetExternalDataForTests(next?: Fetcher): void {
  entries.clear();
  requestedBy.clear();
  storedBy.clear();
  listCache.clear();
  lastLists.clear();
  answers = 0;
  hostKeys.clear();
  deniedHosts.clear();
  hostQuestions = [];
  revision = 0;
  fetcher = next ?? fetchFromServer;
}
