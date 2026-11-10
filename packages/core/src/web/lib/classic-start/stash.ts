/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a signed-out visitor started on the homepage's workbench (classic
 * mode): up to one document, one spreadsheet, one presentation (and one
 * PDF), each a file they dropped or a blank page they made, with what they
 * did to it and where its window stood. Kept in this browser until it is
 * placed in their workspace after sign-up (/open/start). A file's bytes
 * travel as the File itself, never in localStorage and never to the server
 * before sign-up. A record older than its expiry is dropped unread; an item
 * is removed once it is placed (and its edits are on the page), or kept
 * while it still holds edits the visitor has not been able to copy out.
 */

export const CLASSIC_START_DB = "mrai-classic-start";
export const CLASSIC_START_STORE = "sessions";
export const CLASSIC_START_KEY = "current";
/** Long enough for a sign-up that waits on an e-mailed code. */
export const CLASSIC_START_MAX_AGE_MS = 24 * 60 * 60 * 1000;

export type ClassicPageType = "document" | "spreadsheet" | "presentation" | "pdf" | "code";

export interface ClassicRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * The guest editor's own serializations, before the first edit and now.
 * Both come from the editor's save path, so what the visitor did is their
 * difference and nothing else.
 */
export interface ClassicGuestEdits {
  readonly base: Record<string, unknown>;
  readonly final: Record<string, unknown>;
}

/** After a resume whose edits could not be placed on the imported page. */
export interface ClassicKeptOutcome {
  readonly notebookId: string;
  readonly pageId: string;
  readonly reason: string;
}

export interface ClassicItem {
  /** One per page type: the workbench holds one window of each. */
  readonly type: ClassicPageType;
  readonly kind: "file" | "blank";
  readonly title: string;
  /** The original file, for `kind: "file"`. */
  readonly file?: File;
  /** False when the guest reader could not open the file; it is still
   *  imported after sign-up, as it is. */
  readonly guestEditable: boolean;
  readonly refusal?: string;
  readonly edits?: ClassicGuestEdits;
  /** The window, relative to the workbench's canvas. */
  readonly frame?: ClassicRect;
  /** The document's own zoom (1 = 100%), where the editor has one. */
  readonly zoom?: number;
  /** The page /open/start made of it (a resume goes on from there). */
  readonly placed?: { readonly notebookId: string; readonly pageId: string };
  readonly kept?: ClassicKeptOutcome;
}

export interface ClassicStartRecord {
  readonly v: 2;
  readonly savedAt: number;
  readonly expiresAt: number;
  /** The workbench canvas the frames are relative to. */
  readonly canvas?: { readonly width: number; readonly height: number };
  readonly items: readonly ClassicItem[];
}

export type ClassicStartDraft = Omit<ClassicStartRecord, "v" | "savedAt" | "expiresAt">;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CLASSIC_START_DB, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(CLASSIC_START_STORE)) {
        request.result.createObjectStore(CLASSIC_START_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function withStore<T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  const db = await openDb();
  try {
    return await new Promise<T>((resolve, reject) => {
      const transaction = db.transaction(CLASSIC_START_STORE, mode);
      const request = run(transaction.objectStore(CLASSIC_START_STORE));
      transaction.oncomplete = () => resolve(request.result);
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  } finally {
    db.close();
  }
}

/** Keep (or replace) the visitor's session; an empty one is deleted. */
export async function writeClassicStart(draft: ClassicStartDraft, now = Date.now()): Promise<void> {
  if (draft.items.length === 0) {
    await clearClassicStart();
    return;
  }
  const record: ClassicStartRecord = {
    ...draft,
    v: 2,
    savedAt: now,
    expiresAt: now + CLASSIC_START_MAX_AGE_MS,
  };
  await withStore("readwrite", (store) => store.put(record, CLASSIC_START_KEY));
}

function validItem(value: unknown): value is ClassicItem {
  if (!value || typeof value !== "object") return false;
  const item = value as Partial<ClassicItem>;
  return (
    typeof item.type === "string" &&
    (item.kind === "file" || item.kind === "blank") &&
    (item.kind === "blank" || item.file instanceof Blob)
  );
}

/** A record of the first version (one item, frames in viewport pixels). */
function fromFirstVersion(value: Record<string, unknown>): ClassicStartRecord | null {
  const desk = value.desk as { frame?: ClassicRect; viewport?: { width: number; height: number }; zoom?: number } | undefined;
  const item = {
    type: value.pageType,
    kind: value.kind,
    title: value.title,
    ...(value.file ? { file: value.file } : {}),
    guestEditable: value.guestEditable !== false,
    ...(value.refusal ? { refusal: value.refusal } : {}),
    ...(value.edits ? { edits: value.edits } : {}),
    ...(desk?.frame ? { frame: desk.frame } : {}),
    ...(typeof desk?.zoom === "number" ? { zoom: desk.zoom } : {}),
    ...(value.placed ? { placed: value.placed } : {}),
    ...(value.kept ? { kept: value.kept } : {}),
  };
  if (!validItem(item) || typeof value.expiresAt !== "number") return null;
  return {
    v: 2,
    savedAt: Number(value.savedAt) || 0,
    expiresAt: value.expiresAt,
    ...(desk?.viewport ? { canvas: desk.viewport } : {}),
    items: [item],
  };
}

function asRecord(value: unknown): ClassicStartRecord | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (record.v === 1) return fromFirstVersion(record);
  if (record.v !== 2 || typeof record.expiresAt !== "number" || !Array.isArray(record.items)) return null;
  if (!record.items.every(validItem)) return null;
  return record as unknown as ClassicStartRecord;
}

/** The kept session; none when there is none, it is malformed or stale (then it is deleted). */
export async function readClassicStart(now = Date.now()): Promise<ClassicStartRecord | null> {
  try {
    const stored = await withStore<unknown>("readonly", (store) => store.get(CLASSIC_START_KEY));
    if (stored === undefined) return null;
    const record = asRecord(stored);
    if (!record || now > record.expiresAt || record.items.length === 0) {
      await clearClassicStart();
      return null;
    }
    return record;
  } catch {
    return null;
  }
}

export async function clearClassicStart(): Promise<void> {
  try {
    await withStore("readwrite", (store) => store.delete(CLASSIC_START_KEY));
  } catch {
    // nothing kept, or storage is unavailable: nothing to clear
  }
}

/** Change one item (by page type) of the kept session, or drop it (null). */
export async function updateClassicItem(
  type: ClassicPageType,
  change: (item: ClassicItem) => ClassicItem | null,
): Promise<ClassicStartRecord | null> {
  const record = await readClassicStart();
  if (!record) return null;
  const items: ClassicItem[] = [];
  for (const item of record.items) {
    if (item.type !== type) {
      items.push(item);
      continue;
    }
    const next = change(item);
    if (next) items.push(next);
  }
  const draft: ClassicStartDraft = { ...(record.canvas ? { canvas: record.canvas } : {}), items };
  await writeClassicStart(draft);
  return items.length ? { ...record, items } : null;
}
