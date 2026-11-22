/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

import { sendErrorBeacon } from "../errors/report-beacon";

/** The page's checkpoint as a read hands it (`serialize_page_causal_checkpoint`,
 * plus the room's custody not yet folded). */
export interface PageCheckpointReceipt {
  readonly epoch?: unknown;
  readonly checkpoint?: { readonly update_base64?: unknown } | null;
  readonly room_custody_updates_base64?: readonly unknown[] | null;
}

export type GapRepairOutcome =
  /** The merge integrated what was pending: nothing waits any more. */
  | "repaired"
  /** Merged, and something still waits on ancestors the checkpoint lacks too. */
  | "still_pending"
  /** The checkpoint is of another causal epoch: not merged. */
  | "epoch_mismatch"
  /** No checkpoint to merge (the page has none, or its update is unreadable). */
  | "no_checkpoint";

export interface GapPendingSummary {
  readonly missing: number;
  readonly ds: boolean;
}

/** What `doc` holds pending: how many clients' structs wait, and whether a
 * delete set does. */
export function gapPending(doc: Y.Doc): GapPendingSummary {
  const store = doc.store as unknown as { pendingStructs: { missing: Map<number, number> } | null; pendingDs: Uint8Array | null };
  return { missing: store.pendingStructs ? store.pendingStructs.missing.size : 0, ds: Boolean(store.pendingDs) };
}

function gapped(doc: Y.Doc): boolean {
  const pending = gapPending(doc);
  return pending.missing > 0 || pending.ds;
}

function decodeBase64(value: unknown): Uint8Array | null {
  if (typeof value !== "string" || value.length === 0) return null;
  try {
    const binary = atob(value);
    const out = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) out[index] = binary.charCodeAt(index);
    return out;
  } catch {
    return null;
  }
}

/** A causal epoch as both sides name it: 0 and 1 are the page's first epoch. */
function causalEpoch(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0) return null;
  return value <= 1 ? 1 : value;
}

/**
 * Merge `receipt` into `doc` as the room's own (`origin`, which keeps it from
 * being sent back), when it is of the room's causal epoch. Never throws.
 */
export function applyCheckpointRepair(
  doc: Y.Doc,
  receipt: PageCheckpointReceipt | null | undefined,
  roomEpoch: number,
  origin: unknown,
): GapRepairOutcome {
  if (!receipt) return "no_checkpoint";
  const epoch = causalEpoch(receipt.epoch);
  if (epoch === null || epoch !== causalEpoch(roomEpoch)) return "epoch_mismatch";
  const update = decodeBase64(receipt.checkpoint?.update_base64);
  if (!update) return "no_checkpoint";
  const custody = (receipt.room_custody_updates_base64 ?? []).map(decodeBase64).filter((item): item is Uint8Array => item !== null);
  try {
    doc.transact(() => {
      Y.applyUpdate(doc, update, origin);
      for (const item of custody) Y.applyUpdate(doc, item, origin);
    }, origin);
  } catch {
    return "no_checkpoint";
  }
  return gapped(doc) ? "still_pending" : "repaired";
}

/** At most one checkpoint read per page a minute, whichever window asks. */
export const GAP_REPAIR_INTERVAL_MS = 60_000;
const lastRepairAt = new Map<string, number>();

/** Whether `pageId` may be repaired now (and claims it), or how long to wait. */
export function claimGapRepair(pageId: string, now: number = Date.now()): { readonly claimed: boolean; readonly waitMs: number } {
  const last = lastRepairAt.get(pageId);
  if (last !== undefined && now - last < GAP_REPAIR_INTERVAL_MS) {
    return { claimed: false, waitMs: GAP_REPAIR_INTERVAL_MS - (now - last) };
  }
  lastRepairAt.set(pageId, now);
  return { claimed: true, waitMs: 0 };
}

/** Test seam. */
export function resetGapRepairForTests(): void {
  lastRepairAt.clear();
  incomplete.clear();
  listeners.clear();
}

// The page's "some collaborators' edits cannot be shown" signal: one entry per
// provider (window) that holds an unrepaired gap; the page is incomplete while
// any does. Each entry says whether the room keeps every edit that window made
// (`editsHeld`): a document holding a gap cannot be captured into the page
// (`checkpoint_incomplete`), so its window never reads "saved" until it is
// reopened, and what makes the reopen safe is the room keeping its edits.
const incomplete = new Map<string, Map<symbol, () => boolean>>();
const listeners = new Map<string, Set<() => void>>();

export function setPageGapIncomplete(
  pageId: string,
  providerId: symbol,
  on: boolean,
  editsHeld: () => boolean = () => false,
): void {
  const holders = incomplete.get(pageId) ?? new Map<symbol, () => boolean>();
  const before = holders.size > 0;
  if (on) holders.set(providerId, editsHeld);
  else holders.delete(providerId);
  if (holders.size > 0) incomplete.set(pageId, holders);
  else incomplete.delete(pageId);
  if (before !== holders.size > 0) for (const listener of [...(listeners.get(pageId) ?? [])]) listener();
}

/** A gap holder's room may keep more (or less) of its edits now: whoever reads
 * `pageGapEditsHeld` looks again. Nothing when the page holds no gap. */
export function notePageGapEditsHeldChanged(pageId: string): void {
  if (!incomplete.has(pageId)) return;
  for (const listener of [...(listeners.get(pageId) ?? [])]) listener();
}

/** One reading for both the reopen and the save face of a page: no gap
 * (`none`), a gap whose room keeps every edit its windows made (`held`), or a
 * gap it does not yet (`unheld`). */
export type PageGapState = "none" | "held" | "unheld";

export function pageGapState(pageId: string): PageGapState {
  if (!isPageGapIncomplete(pageId)) return "none";
  return pageGapEditsHeld(pageId) ? "held" : "unheld";
}

/** A save face during a gap, from the face it would show otherwise: "synced"
 * exactly when the reopen would go (`held`), never "saved" while the room has
 * not confirmed every edit (`unheld`: "saving", or the failure as it stands). */
export function gapSaveFace<T extends string>(state: PageGapState, face: T): T | "synced" | "saving" {
  if (state === "held") return "synced";
  if (state === "unheld" && face === "saved") return "saving";
  return face;
}

/** Whether `pageId` holds a gap in some window and the room keeps every edit
 * each of those windows made (in its database, not only relayed). */
export function pageGapEditsHeld(pageId: string): boolean {
  const holders = incomplete.get(pageId);
  if (!holders || holders.size === 0) return false;
  for (const held of holders.values()) {
    try {
      if (!held()) return false;
    } catch {
      return false;
    }
  }
  return true;
}

export function isPageGapIncomplete(pageId: string): boolean {
  return (incomplete.get(pageId)?.size ?? 0) > 0;
}

export function subscribePageGapIncomplete(pageId: string, listener: () => void): () => void {
  const set = listeners.get(pageId) ?? new Set<() => void>();
  set.add(listener);
  listeners.set(pageId, set);
  return () => {
    set.delete(listener);
    if (set.size === 0) listeners.delete(pageId);
  };
}

/** The code the server's client-error log is grepped for. */
export const GAP_TELEMETRY_NAME = "PageUpdateGap";

/** One line per gap: which outcome, and what was pending -- never any content. */
export function reportGapOutcome(detail: {
  readonly pageId: string;
  readonly epoch: number;
  readonly outcome: GapRepairOutcome | "resume_limit";
  readonly before: GapPendingSummary;
  readonly after?: GapPendingSummary;
}): void {
  try {
    sendErrorBeacon({
      name: GAP_TELEMETRY_NAME,
      message: JSON.stringify({
        outcome: detail.outcome, page: detail.pageId, epoch: detail.epoch,
        missingBefore: detail.before.missing, dsBefore: detail.before.ds,
        ...(detail.after ? { missingAfter: detail.after.missing, dsAfter: detail.after.ds } : {}),
      }),
      path: typeof location === "undefined" ? "" : location.pathname,
      source: "window",
    });
  } catch {
    // Telemetry never stands in the way.
  }
}

/** A part room joined again because its document held a gap (each time), and
 * a part room's gap three of those did not fill. */
export const PART_RESYNC_TELEMETRY_NAME = "PartRoomPendingResync";
export const PART_GAP_TELEMETRY_NAME = "PartRoomGap";

/** One line per part room resync or unfilled gap: the part, how long it has
 * been pending, which resync and how it was asked, what is missing (client and
 * clock, no more) -- never any content. */
export function reportPartGap(
  name: typeof PART_RESYNC_TELEMETRY_NAME | typeof PART_GAP_TELEMETRY_NAME,
  detail: {
    readonly pageId: string;
    readonly partKey: string;
    readonly pendingMs: number;
    readonly attempt: number;
    readonly pending: GapPendingSummary;
    readonly how: "vector" | "given-up";
    readonly generation: number | null;
    readonly missing: readonly string[];
  },
): void {
  try {
    sendErrorBeacon({
      name,
      message: JSON.stringify({
        page: detail.pageId, part: detail.partKey, pendingMs: Math.round(detail.pendingMs), attempt: detail.attempt,
        how: detail.how, generation: detail.generation, missing: detail.pending.missing, ds: detail.pending.ds,
        lacks: detail.missing,
      }),
      path: typeof location === "undefined" ? "" : location.pathname,
      source: "window",
    });
  } catch {
    // Telemetry never stands in the way.
  }
}
