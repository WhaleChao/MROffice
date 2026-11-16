/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The depth (xref sections) at which the server rewrites a saved chain:
 * a copy this deep with edits unsaved has no room left behind the rewrite. */
export const PDF_WORKING_REVISION_ROOM_DEPTH = 28;

const STARTXREF = [0x73, 0x74, 0x61, 0x72, 0x74, 0x78, 0x72, 0x65, 0x66]; // "startxref"

/**
 * How many times `startxref` is written in `bytes` from `from` on: never
 * fewer than the file's revisions there (each ends with one), so a count
 * below the depth says "below the depth" without parsing the file. (More
 * only where a stream holds the word; the exact depth is then read.)
 * A copy grown by one edit is counted from its last eight old bytes on.
 */
export function countPdfStartxref(bytes: Uint8Array, from = 0): number {
  let count = 0;
  const last = bytes.length - STARTXREF.length;
  for (let index = Math.max(0, from); index <= last; index += 1) {
    if (bytes[index] !== 0x73 || bytes[index + 1] !== 0x74 || bytes[index + 5] !== 0x78) continue;
    let matched = true;
    for (let offset = 2; offset < STARTXREF.length; offset += 1) {
      if (bytes[index + offset] !== STARTXREF[offset]) { matched = false; break; }
    }
    if (matched) { count += 1; index += STARTXREF.length - 1; }
  }
  return count;
}

/** The revisions the parser reads: bytes that hold this many take no edit. */
export const PDF_WORKING_REVISION_CEILING = 32;

/**
 * That bound for a head of the working copy, kept by digest: a head one edit
 * on from a head already counted is read from its old end only.
 */
export function pdfWorkingRevisionBound(counts: Map<string, number>, head: string, bytes: () => Uint8Array,
  grownFrom?: Readonly<{ sha256: string; length: number }> | null): number {
  const known = counts.get(head);
  if (known !== undefined) return known;
  const data = bytes();
  const before = grownFrom && grownFrom.length <= data.length ? counts.get(grownFrom.sha256) : undefined;
  const bound = before !== undefined
    ? before + countPdfStartxref(data, grownFrom!.length - 8) : countPdfStartxref(data);
  counts.set(head, bound);
  while (counts.size > 8) counts.delete(counts.keys().next().value!);
  return bound;
}










export function pdfRequestOnRewrittenHead<T>(request: T, from: string, to: string, depth = 0): T {
  if (depth > 6 || request === null || typeof request !== "object" || from === to) return request;
  if (Array.isArray(request)) {
    const mapped = request.map((entry) => pdfRequestOnRewrittenHead(entry, from, to, depth + 1));
    return (Object.isFrozen(request) ? Object.freeze(mapped) : mapped) as T;
  }
  const prototype = Object.getPrototypeOf(request);
  if (prototype !== Object.prototype && prototype !== null) return request;
  const mapped: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(request as Record<string, unknown>)) {
    mapped[key] = key === "sourceSha256" && typeof entry === "string" && entry.toLowerCase() === from.toLowerCase()
      ? to : pdfRequestOnRewrittenHead(entry, from, to, depth + 1);
  }
  return (Object.isFrozen(request) ? Object.freeze(mapped) : mapped) as T;
}

/** How long an edit that asks for room waits before a save that failed for
 * the same bytes is tried again. */
const RETRY_AFTER_FAILURE_MS = 10_000;

export type PdfWorkingLimitSaveOutcome = "none" | "saved" | "failed";

export interface PdfWorkingLimitSaveHost {
  autosaveOn(): boolean;
  /** The working copy as it stands; null when none is open. */
  read(): Readonly<{ workingSha256: string; pending: number; saving: boolean }> | null;
  /** The xref sections of the copy's bytes (null: not read). */
  depth(): Promise<number | null>;
  /** An edit is being planned or staged right now: the save waits for the
   * copy to move, and is asked again then. */
  busy(): boolean;
  /** The Save command: true when everything held was written. */
  save(): Promise<boolean>;
  notify(outcome: "saved" | "failed"): void;
  now?(): number;
}

export interface PdfWorkingLimitSave {
  /** The check or save running now, if any. */
  inFlight(): Promise<PdfWorkingLimitSaveOutcome> | null;
  /**
   * Saves once when the copy is at the depth and holds unsaved edits.
   * "saved": the Save command wrote everything; "failed": it did not (said
   * once per attempt); "none": nothing was due. `asked`: an edit is waiting
   * for the room -- a save that failed for these same bytes is then tried
   * again (ten seconds apart), where the copy merely changing is not.
   */
  check(options?: Readonly<{ asked?: boolean }>): Promise<PdfWorkingLimitSaveOutcome>;
}

export function createPdfWorkingLimitSave(host: PdfWorkingLimitSaveHost): PdfWorkingLimitSave {
  const now = () => host.now?.() ?? Date.now();
  let flight: Promise<PdfWorkingLimitSaveOutcome> | null = null;
  let failed: { head: string; at: number } | null = null;
  let measured: { head: string; depth: number | null } | null = null;

  const run = async (asked: boolean): Promise<PdfWorkingLimitSaveOutcome> => {
    // The copy may move while its depth is read: looked at again, a few times.
    for (let attempt = 0; attempt < 4; attempt += 1) {
      if (host.autosaveOn()) return "none";
      const before = host.read();
      if (!before || before.pending < 1 || before.saving || host.busy()) return "none";
      const head = before.workingSha256;
      if (failed?.head === head && (!asked || now() - failed.at < RETRY_AFTER_FAILURE_MS)) return "none";
      if (measured?.head !== head) {
        let depth: number | null = null;
        try { depth = await host.depth(); } catch { depth = null; }
        // (Read for the bytes the copy held when it was asked.)
        if (host.read()?.workingSha256 !== head) continue;
        measured = { head, depth };
      }
      if (measured.depth === null || measured.depth < PDF_WORKING_REVISION_ROOM_DEPTH) return "none";
      const current = host.read();
      if (host.autosaveOn() || !current || current.pending < 1 || current.saving || host.busy()) return "none";
      if (current.workingSha256 !== head) continue;
      let saved = false;
      try { saved = await host.save(); } catch { saved = false; }
      if (saved) {
        failed = null;
        host.notify("saved");
        return "saved";
      }
      failed = { head: host.read()?.workingSha256 ?? head, at: now() };
      host.notify("failed");
      return "failed";
    }
    return "none";
  };

  return {
    inFlight: () => flight,
    check(options) {
      if (flight) return flight;
      const mine: Promise<PdfWorkingLimitSaveOutcome> = run(options?.asked === true).finally(() => {
        if (flight === mine) flight = null;
      });
      flight = mine;
      return mine;
    },
  };
}
