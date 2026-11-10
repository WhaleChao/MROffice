/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a window closing on a page still has to say, and the word it leaves for
 * the window opened next, for every editor that sends one:
 *
 * - a workbook whose rooms are still preparing sends its cells
 *   (lib/spreadsheet/workbook-closing-edits.ts, `workbook-collab-v2/closing-edits`);
 * - a page held as one shared document (Docs) sends what its document holds
 *   beyond the checkpoint it adopted: original operations, an Undo a deletion
 *   of a few bytes however long the page (`/pages/{id}/closing-edits`,
 *   services/page_closing_edits.py).
 *
 * Both go out as the page closes (`keepalive`, at most 64 KiB), and both leave
 * a note here; the window opened next reads the page only once they are in it,
 * and from then on shows the stored page: nothing is guessed from a room or a
 * local copy (a45c/ab27 sac21/sac22: one closing in four kept an undone word).
 * Edits too large for the request are left to the local recovery copy, as before.
 */

/** A request sent as a page closes may carry at most 64 KiB (`keepalive`). */
export const PAGE_CLOSING_EDITS_MAX_BYTES = 60 * 1024;

const CLOSING_EDITS_SENT_PREFIX = "mrnote:workbook-closing-edits-sent:";
/** A note older than this is not waited on. */
const CLOSING_EDITS_WAIT_MS = 2 * 60 * 1000;
/** The longest a window opening waits for the edits before it opens anyway. */
const CLOSING_EDITS_OPEN_WAIT_MS = 30 * 1000;
/** Never heard of after this long, a request is taken as lost with its window. */
const CLOSING_EDITS_UNHEARD_MS = 8 * 1000;
const CLOSING_EDITS_OPEN_POLL_MS = 500;

export interface ClosingEditsSent {
  at: number;
  /** The page version the edits were made from. */
  base: number;
  /** A shared document's request names itself: what became of it is asked by this. */
  id?: string;
  /** The tab that sent them (its AutoRecover slot, its closing copy). */
  tab?: string;
  /** Once in the page: the blocks the server kept twice (someone else had
   * changed them meanwhile), each block and the copy beside it. */
  conflicts?: Array<{ kept: string; copy: string }>;
}

/** The body, when it fits a request sent as a page closes. */
export function closingEditsBody(edits: unknown): string | null {
  const body = JSON.stringify(edits);
  return new Blob([body]).size <= PAGE_CLOSING_EDITS_MAX_BYTES ? body : null;
}

export function newClosingEditsId(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function readClosingEditsSent(pageId: string, now = Date.now()): ClosingEditsSent | null {
  try {
    const key = CLOSING_EDITS_SENT_PREFIX + pageId;
    const raw = window.localStorage.getItem(key);
    if (raw === null) return null;
    const sent = JSON.parse(raw) as Partial<ClosingEditsSent> | null;
    if (typeof sent?.at === "number" && typeof sent.base === "number" && now - sent.at < CLOSING_EDITS_WAIT_MS) {
      return {
        at: sent.at, base: sent.base,
        ...(typeof sent.id === "string" ? { id: sent.id } : {}),
        ...(typeof sent.tab === "string" ? { tab: sent.tab } : {}),
      };
    }
    window.localStorage.removeItem(key);
  } catch {
    // Read as not sent.
  }
  return null;
}

export function noteClosingEditsSent(
  pageId: string,
  baseContentVersion: number,
  now = Date.now(),
  request: { id?: string; tab?: string } = {},
): void {
  try {
    window.localStorage.setItem(
      CLOSING_EDITS_SENT_PREFIX + pageId,
      JSON.stringify({ at: now, base: baseContentVersion, ...request } satisfies ClosingEditsSent),
    );
  } catch {
    // Without the note the page opens, and a copy is offered, at once.
  }
}

export function forgetClosingEditsSent(pageId: string): void {
  try {
    window.localStorage.removeItem(CLOSING_EDITS_SENT_PREFIX + pageId);
  } catch {
    // nothing kept
  }
}

export function closingEditsSentRecently(pageId: string, now = Date.now()): boolean {
  return readClosingEditsSent(pageId, now) !== null;
}

/** What the server says of a closing request: in the page, being saved now
 * (it has it: a long page's save, a Word file written back, takes as long as
 * it takes, up to the note's age), still to come, or never to be. */
export type ClosingEditsReading = "landed" | "working" | "pending" | "failed";

/** The closing edits the page took, by page, until the view opened on it asks
 * (takeLandedClosingEdits): its recovery copies of the same edits are done. */
const landedClosings = new Map<string, ClosingEditsSent>();

export function takeLandedClosingEdits(pageId: string): ClosingEditsSent | null {
  const landed = landedClosings.get(pageId) ?? null;
  landedClosings.delete(pageId);
  return landed;
}

/**
 * Before reading a page that a window just left with edits on their way: wait
 * until the page holds them. True when it does; false when nothing was sent,
 * the server refused them, or nothing came within the wait (the page then
 * opens, and the recovery copy answers for those edits as ever).
 */
export async function waitForClosingEdits(
  pageId: string,
  read: (pageId: string, sent: ClosingEditsSent) => Promise<ClosingEditsReading | null>,
  isCancelled: () => boolean,
): Promise<boolean> {
  const sent = readClosingEditsSent(pageId, Date.now());
  if (!sent) return false;
  let deadline = Math.min(sent.at + CLOSING_EDITS_WAIT_MS, Date.now() + CLOSING_EDITS_OPEN_WAIT_MS);
  while (!isCancelled() && Date.now() < deadline) {
    let reading: ClosingEditsReading | null;
    try {
      reading = await read(pageId, sent);
    } catch {
      return false;
    }
    // A server without the read does not answer it: nothing to wait on.
    if (reading === null) return false;
    if (reading === "landed") {
      // A named request is answered for good; a workbook's note stays its
      // two minutes (its rooms and its recovery copy read it too).
      if (sent.id) {
        landedClosings.set(pageId, sent);
        forgetClosingEditsSent(pageId);
      }
      return true;
    }
    if (reading === "failed") {
      if (sent.id) forgetClosingEditsSent(pageId);
      return false;
    }
    // The server has them and is saving them: opened before that ends, the
    // view would show the page without them.
    if (reading === "working") deadline = sent.at + CLOSING_EDITS_WAIT_MS;
    await new Promise((resolve) => setTimeout(resolve, CLOSING_EDITS_OPEN_POLL_MS));
  }
  return false;
}

/** A named request's state as the server answers it, read for waitForClosingEdits. */
export function closingEditsReadingOf(
  state: string | null | undefined,
  sent: ClosingEditsSent,
  now = Date.now(),
): ClosingEditsReading {
  if (state === "applied" || state === "unchanged") return "landed";
  if (state === "refused") return "failed";
  if (state === "received") return "working";
  // Not heard of: still on its way a moment after closing, else lost with the window.
  return now - sent.at < CLOSING_EDITS_UNHEARD_MS ? "pending" : "failed";
}

/** Base64 of a closing request's bytes (at most the request's budget). */
export function closingBytesBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return btoa(binary);
}

/** One block a closing window changed, as `/pages/{id}/closing-edits` takes
 * it for a document saved as JSON (services/page_closing_edits.py). */
export interface ClosingBlockChange {
  id: string;
  /** null: the window deleted it. */
  block: unknown;
  /** Its readings the window knew before (the stored part's, the save on its way's). */
  before: unknown[];
  /** Present for a block placed anew (new, or after another than before). */
  after?: string | null;
}

function closingBlockId(block: unknown): string | null {
  const attrs = block && typeof block === "object" ? (block as { attrs?: unknown }).attrs : null;
  const id = attrs && typeof attrs === "object" ? (attrs as { block_id?: unknown }).block_id : null;
  return typeof id === "string" && id ? id : null;
}

/**
 * What `current` changed against the earlier readings of the same part
 * (`earlier`: the stored part, the save on its way; null for one not
 * known): only those blocks, so an Undo in a long document is one block.
 * Null when a block has no id or one twice (not told apart by id).
 */
export function closingBlockChanges(current: readonly unknown[], earlier: ReadonlyArray<readonly unknown[] | null>): ClosingBlockChange[] | null {
  type Reading = { json: string; block: unknown; after: string | null };
  const read = (blocks: readonly unknown[]): Map<string, Reading> | null => {
    const map = new Map<string, Reading>();
    let previous: string | null = null;
    for (const block of blocks) {
      const id = closingBlockId(block);
      if (!id || map.has(id)) return null;
      map.set(id, { json: JSON.stringify(block), block, after: previous });
      previous = id;
    }
    return map;
  };
  const now = read(current);
  if (!now) return null;
  const before: Map<string, Reading>[] = [];
  for (const blocks of earlier) {
    if (!blocks) continue;
    const map = read(blocks);
    if (!map) return null;
    before.push(map);
  }
  const changes: ClosingBlockChange[] = [];
  const readingsOf = (id: string, except?: string) => {
    const seen = new Set<string>();
    const out: unknown[] = [];
    for (const map of before) {
      const reading = map.get(id);
      if (reading && reading.json !== except && !seen.has(reading.json)) { seen.add(reading.json); out.push(reading.block); }
    }
    return out;
  };
  for (const [id, mine] of now) {
    const known = before.map((map) => map.get(id));
    const changed = known.some((reading) => !reading || reading.json !== mine.json);
    const placed = known.every((reading) => !reading || reading.after !== mine.after);
    if (!changed && !placed) continue;
    changes.push({ id, block: mine.block, before: readingsOf(id, mine.json), ...(placed ? { after: mine.after } : {}) });
  }
  const deleted = new Set<string>();
  for (const map of before) for (const id of map.keys()) if (!now.has(id)) deleted.add(id);
  for (const id of deleted) changes.unshift({ id, block: null, before: readingsOf(id) });
  return changes;
}
