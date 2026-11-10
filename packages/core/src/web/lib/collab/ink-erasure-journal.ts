/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Awareness } from "y-protocols/awareness";

export const INK_ERASED_FIELD = "inkErased";
/** How long a reader's own deletion explains the room taking those strokes out. */
export const INK_ERASURE_KEEP_MS = 15_000;
/** The most entries one view keeps (and puts in its presence). */
export const INK_ERASURE_MAX = 64;

export interface InkErasureEntry {
  /** Unique per deletion of one stroke: how a view knows it already used it. */
  readonly i: string;
  /** The stroke's fingerprint (inkFingerprint). */
  readonly f: string;
  /** When it was written down, in this tab's clock. */
  readonly t: number;
}

interface PresenceState {
  readonly v: 1;
  readonly e: readonly { readonly i: string; readonly f: string }[];
}

interface Session {
  readonly read: Awareness | null;
  readonly write: Awareness | null;
}

const local = new Map<string, InkErasureEntry[]>();
const sessions = new Map<string, Set<Session>>();
let counter = 0;

function now(): number {
  return Date.now();
}

function nextId(): string {
  counter = (counter + 1) % 1_000_000;
  return `${now().toString(36)}${counter.toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

function stable(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    return `{${Object.keys(record)
      .filter((key) => record[key] !== undefined && record[key] !== null)
      .sort()
      .map((key) => `${key}:${stable(record[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
}

/** A stroke's identity across views: two FNV-1a passes over its attributes. */
export function inkFingerprint(typeName: string, attrs: unknown): string {
  const text = `${typeName}|${stable(attrs)}`;
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ text.length;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    a = Math.imul(a ^ code, 0x01000193) >>> 0;
    b = Math.imul(b ^ code, 0x5bd1e995) >>> 0;
  }
  return `${a.toString(36)}${b.toString(36)}`;
}

function pruned(entries: readonly InkErasureEntry[]): InkErasureEntry[] {
  const since = now() - INK_ERASURE_KEEP_MS;
  return entries.filter((entry) => entry.t >= since).slice(-INK_ERASURE_MAX);
}

/** This page's presence in the room, to read peers' deletions and publish ours. */
export function registerInkErasureAwareness(
  pageId: string,
  read: Awareness | null,
  write: Awareness | null,
): () => void {
  const session: Session = { read, write };
  const set = sessions.get(pageId) ?? new Set<Session>();
  set.add(session);
  sessions.set(pageId, set);
  return () => {
    set.delete(session);
    if (set.size === 0) sessions.delete(pageId);
    try {
      write?.setLocalStateField(INK_ERASED_FIELD, null);
    } catch {
      /* the session is gone */
    }
  };
}

/** The reader of this view took these strokes out on purpose. */
export function publishInkErasures(pageId: string, fingerprints: readonly string[]): void {
  if (!pageId || fingerprints.length === 0) return;
  const at = now();
  const entries = pruned([
    ...(local.get(pageId) ?? []),
    ...fingerprints.map((f) => ({ i: nextId(), f, t: at })),
  ]);
  local.set(pageId, entries);
  const state: PresenceState = { v: 1, e: entries.map(({ i, f }) => ({ i, f })) };
  for (const session of sessions.get(pageId) ?? []) {
    try {
      session.write?.setLocalStateField(INK_ERASED_FIELD, state);
    } catch {
      /* the session is gone */
    }
  }
}

function presenceEntries(value: unknown): { i: string; f: string }[] {
  if (!value || typeof value !== "object") return [];
  const state = value as Partial<PresenceState>;
  if (state.v !== 1 || !Array.isArray(state.e)) return [];
  return state.e
    .filter((entry) => entry && typeof entry.i === "string" && typeof entry.f === "string")
    .slice(-INK_ERASURE_MAX);
}

/** Every deletion this page's views and members wrote down, by entry key. */
export function inkErasureEntries(pageId: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const entry of pruned(local.get(pageId) ?? [])) out.set(entry.i, entry.f);
  for (const session of sessions.get(pageId) ?? []) {
    const states = session.read?.getStates();
    if (!states) continue;
    for (const state of states.values()) {
      for (const entry of presenceEntries((state as Record<string, unknown>)[INK_ERASED_FIELD])) {
        out.set(entry.i, entry.f);
      }
    }
  }
  return out;
}

/**
 * The fingerprints in `removed` no deletion explains, using up (in `used`) the
 * entries that do: one entry explains one stroke.
 */
export function unexplainedInkRemovals(
  pageId: string | null | undefined,
  removed: readonly string[],
  used: Set<string>,
): string[] {
  if (!pageId || removed.length === 0) return [...removed];
  const available = new Map<string, string[]>();
  for (const [key, fingerprint] of inkErasureEntries(pageId)) {
    if (used.has(key)) continue;
    const keys = available.get(fingerprint) ?? [];
    keys.push(key);
    available.set(fingerprint, keys);
  }
  const left: string[] = [];
  for (const fingerprint of removed) {
    const key = available.get(fingerprint)?.shift();
    if (key) used.add(key);
    else left.push(fingerprint);
  }
  return left;
}

/** Test seam: forget this tab's journal. */
export function resetInkErasureJournalForTests(): void {
  local.clear();
  sessions.clear();
}
