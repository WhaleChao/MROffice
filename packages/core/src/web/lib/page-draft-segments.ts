/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A page's local recovery copy, kept in sessionStorage in parts.
 *
 * The recovery copy of a page is written before every save may be sent
 * (useAutosavedPageDocument persistPageDraft), and it was written whole: the
 * body as JSON and, for a collaborative page, its whole Yjs state as base64.
 * On an inked page of a few hundred strokes that is megabytes of text, and
 * writing it (and serializing it first) took 50-150 ms of the editor's thread
 * at a time under the pen, every save, growing with each stroke.
 *
 * A large copy is now stored as a small manifest under the draft's own key,
 * plus parts under keys of their own: runs of the body's blocks, and pieces
 * of the Yjs state. A part that is already stored, unchanged, is not written
 * again: a block the editor did not touch is the same frozen object from save
 * to save (its text is kept with it), and a piece of the state is an object
 * that stands for the same bytes for as long as it lives (y-state-encoder).
 * So a save writes what changed since the last one, and the manifest.
 *
 * The manifest says how to put the copy together again, character for
 * character the JSON.stringify of the same draft, and how long that text is;
 * a reader that cannot find every part at its length has no copy (it reads
 * nothing, as for any corrupt copy). A part is written under a fresh key and
 * never overwritten, the manifest after the parts, and parts no manifest
 * names any more are removed only after the manifest is written: a write that
 * fails half-way (the tab's quota) leaves the previous copy whole.
 */
import { pageCausalRecoveryPieces } from "./collab/page-causal-checkpoint";
import type { YDocStatePiece } from "./collab/y-state-encoder";

const MARK = "mrnoteDraftSegments";
const MANIFEST_HEAD = /^\{"savedAt":([^,]*),"mrnoteDraftSegments":1,/;
/** Below these a copy is written whole, as ever. */
export const PAGE_DRAFT_SEGMENT_MIN_BLOCKS = 32;
export const PAGE_DRAFT_SEGMENT_MIN_STATE_BYTES = 64 * 1024;
/** A run of blocks ends at the first block past this many characters. */
const BLOCK_RUN_MAX_CHARS = 96 * 1024;
/** A piece of the state this small rides in the manifest itself. */
const INLINE_PIECE_MAX_BYTES = 1024;

type PlanPart =
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "blocks"; readonly texts: readonly string[]; readonly length: number }
  | { readonly kind: "bytes"; readonly pieces: readonly YDocStatePiece[] };

/** What a draft is made of, before anything is stored: see planPageDraftSegments. */
export interface PageDraftSegmentPlan {
  readonly savedAt: number;
  readonly parts: readonly PlanPart[];
  /** The length of the whole copy (JSON.stringify of the draft). */
  readonly length: number;
}

const record = (value: unknown): value is Record<string, unknown> =>
  Boolean(value && typeof value === "object" && !Array.isArray(value));

/** The JSON text of a frozen block, kept with it (it cannot change). */
const frozenBlockTexts = new WeakMap<object, string>();
function blockText(block: object): string {
  if (!Object.isFrozen(block)) return JSON.stringify(block);
  let text = frozenBlockTexts.get(block);
  if (text === undefined) frozenBlockTexts.set(block, (text = JSON.stringify(block)));
  return text;
}

/** Keep a frozen block's JSON text ahead of the draft that needs it
 * (saveTextWarmup): the text the draft would keep. */
export function warmPageDraftBlockText(block: object): void {
  if (Object.isFrozen(block)) blockText(block);
}

/** Whether a run of blocks ends after this block: decided by the block's own
 * text, so a block inserted or changed moves only the run it falls in. */
function endsRun(text: string): boolean {
  const length = text.length;
  const hash = Math.imul(length, 0x9e3779b1) ^ Math.imul(text.charCodeAt(length >> 1), 0x85ebca6b) ^
    Math.imul(text.charCodeAt(Math.max(0, length - 3)), 0xc2b2ae35);
  return ((hash >>> 7) & 7) === 0;
}

function base64Length(bytes: number): number {
  return Math.ceil(bytes / 3) * 4;
}

type NativeBase64 = { toBase64?: () => string };
function encodeBase64(bytes: Uint8Array): string {
  if (typeof (bytes as Uint8Array & NativeBase64).toBase64 === "function") {
    return (bytes as Uint8Array & { toBase64(): string }).toBase64();
  }
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return btoa(binary);
}

function decodeBase64(text: string): Uint8Array | null {
  try {
    const fromBase64 = (Uint8Array as unknown as { fromBase64?: (value: string) => Uint8Array }).fromBase64;
    if (typeof fromBase64 === "function") return fromBase64(text);
    const binary = atob(text);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  } catch {
    return null;
  }
}

/** The pieces of a recovery's state, when they are its very bytes. */
function statePiecesOf(recovery: unknown): readonly YDocStatePiece[] | null {
  if (!record(recovery) || typeof recovery.update_base64 !== "string") return null;
  const text = recovery.update_base64;
  const pieces = pageCausalRecoveryPieces(recovery);
  if (!pieces || pieces.length === 0) {
    // A copy of a captured recovery (the same text in another object: a
    // draft merged or read back): the pieces of the capture with that text.
    for (const known of recentStates) if (known.text === text) return known.pieces;
    return null;
  }
  let size = 0;
  for (const piece of pieces) size += piece.bytes.length;
  if (base64Length(size) !== text.length) return null;
  const first = pieces[0].bytes;
  const head = first.subarray(0, Math.min(first.length, 300) - (Math.min(first.length, 300) % 3));
  if (head.length > 0 && encodeBase64(head) !== text.slice(0, base64Length(head.length))) return null;
  if (recentStates[0]?.pieces !== pieces) recentStates.unshift({ text, pieces });
  recentStates.length = Math.min(recentStates.length, RECENT_STATES_KEPT);
  return pieces;
}
/** The last states stored in pieces, by their base64 text. */
const recentStates: Array<{ readonly text: string; readonly pieces: readonly YDocStatePiece[] }> = [];
const RECENT_STATES_KEPT = 2;

/**
 * How `payload` (the object persistPageDraft stores, `savedAt` its first
 * key) is stored in parts, or null when it is stored whole: a small draft,
 * or one this does not know how to cut exactly.
 */
export function planPageDraftSegments(payload: Record<string, unknown>): PageDraftSegmentPlan | null {
  const savedAt = payload.savedAt;
  if (typeof savedAt !== "number" || !Number.isFinite(savedAt) || Object.keys(payload)[0] !== "savedAt") return null;
  const draft = payload.draft;
  if (!record(draft) || !record(draft.content_json)) return null;
  const content = draft.content_json;
  const blocks = content.content;
  if (!Array.isArray(blocks) || blocks.some((block) => !record(block) || typeof block.toJSON === "function")) return null;
  const recovery = draft.page_causal_recovery;
  const pieces = statePiecesOf(recovery);
  const stateBytes = pieces ? pieces.reduce((sum, piece) => sum + piece.bytes.length, 0) : 0;
  if (blocks.length < PAGE_DRAFT_SEGMENT_MIN_BLOCKS && stateBytes < PAGE_DRAFT_SEGMENT_MIN_STATE_BYTES) return null;

  const nonce = Math.random().toString(36).slice(2, 10);
  const contentMark = `\u0001${nonce}:content`;
  const updateMark = `\u0001${nonce}:update`;
  const skeleton = JSON.stringify({
    ...payload,
    draft: {
      ...draft,
      content_json: { ...content, content: contentMark },
      ...(pieces ? { page_causal_recovery: { ...(recovery as Record<string, unknown>), update_base64: updateMark } } : {}),
    },
  });
  const head = `{"savedAt":${JSON.stringify(savedAt)}`;
  if (!skeleton.startsWith(head + ",")) return null;
  const marks: Array<{ at: number; quoted: string; kind: "content" | "update" }> = [];
  const placed: Array<[string, "content" | "update"]> = [[contentMark, "content"]];
  if (pieces) placed.push([updateMark, "update"]);
  for (const [mark, kind] of placed) {
    const quoted = JSON.stringify(mark);
    const at = skeleton.indexOf(quoted);
    if (at < 0 || skeleton.indexOf(quoted, at + 1) >= 0) return null;
    marks.push({ at, quoted, kind });
  }
  marks.sort((a, b) => a.at - b.at);

  const parts: PlanPart[] = [];
  let length = head.length;
  const text = (value: string) => {
    if (!value) return;
    parts.push({ kind: "text", text: value });
    length += value.length;
  };
  let from = head.length;
  for (const mark of marks) {
    if (mark.kind === "content") {
      text(skeleton.slice(from, mark.at));
      text("[");
      let run: string[] = [], runLength = 0;
      const close = () => {
        if (!run.length) return;
        const joined = runLength + run.length - 1;
        parts.push({ kind: "blocks", texts: run, length: joined });
        length += joined;
        run = [];
        runLength = 0;
      };
      blocks.forEach((block, index) => {
        const value = blockText(block as object);
        if (index > 0 && !run.length) text(",");
        run.push(value);
        runLength += value.length;
        if (runLength >= BLOCK_RUN_MAX_CHARS || endsRun(value)) close();
      });
      close();
      text("]");
      from = mark.at + mark.quoted.length;
    } else {
      // The quotes stay in the text around the state's base64.
      text(skeleton.slice(from, mark.at + 1));
      parts.push({ kind: "bytes", pieces: pieces! });
      length += base64Length(stateBytes);
      from = mark.at + mark.quoted.length - 1;
    }
  }
  text(skeleton.slice(from));
  return { savedAt, parts, length };
}

/** What this window has stored under one draft key. */
interface KeyState {
  /** The ids of the parts this window stored that are still stored (what
   * the stored manifest names, and for a moment after the copy was retired
   * what the last one named: see retirePageDraftSegments). */
  live: Set<string>;
  /** Removes the parts of a retired copy, unless a new copy takes them first. */
  retireTimer: ReturnType<typeof setTimeout> | null;
  /** Runs of blocks stored, by the text of their first block. */
  runs: Map<string, { readonly texts: readonly string[]; readonly id: string }>;
  /** Pieces of a state stored, by the piece. */
  readonly pieces: WeakMap<object, string>;
  /** Parts an earlier page load left under the key were looked for. */
  swept: boolean;
}
const keyStates = new Map<string, KeyState>();
function keyState(storageKey: string): KeyState {
  let state = keyStates.get(storageKey);
  if (!state) keyStates.set(storageKey, (state = { live: new Set(), retireTimer: null, runs: new Map(), pieces: new WeakMap(), swept: false }));
  return state;
}

const loadNonce = Math.random().toString(36).slice(2, 8);
let partSequence = 0;
const partKey = (storageKey: string, id: string) => `${storageKey}#${id}`;

/** A segmented write, decided but not yet made: see prepare/commit. */
export interface PreparedPageDraftSegments {
  readonly storageKey: string;
  /** The manifest to store under the draft key. */
  readonly manifest: string;
  readonly writes: ReadonlyArray<readonly [string, string]>;
  readonly live: ReadonlySet<string>;
  readonly runs: Map<string, { readonly texts: readonly string[]; readonly id: string }>;
  readonly pieces: ReadonlyArray<readonly [object, string]>;
}

/** Which parts of `plan` are stored already (unchanged) and which must be
 * written, and the manifest naming them. Nothing is stored here. */
export function preparePageDraftSegments(storageKey: string, plan: PageDraftSegmentPlan): PreparedPageDraftSegments {
  const state = keyState(storageKey);
  const writes: Array<[string, string]> = [];
  const live = new Set<string>();
  const runs = new Map<string, { readonly texts: readonly string[]; readonly id: string }>();
  const pieces: Array<[object, string]> = [];
  const fresh = () => `${loadNonce}${(partSequence += 1).toString(36)}`;
  const manifestParts: unknown[] = [];
  for (const part of plan.parts) {
    if (part.kind === "text") {
      manifestParts.push(part.text);
      continue;
    }
    if (part.kind === "blocks") {
      const known = state.runs.get(part.texts[0]);
      let id: string;
      if (known && state.live.has(known.id) && known.texts.length === part.texts.length &&
          known.texts.every((value, index) => value === part.texts[index])) id = known.id;
      else {
        id = fresh();
        writes.push([partKey(storageKey, id), part.texts.join(",")]);
      }
      runs.set(part.texts[0], { texts: part.texts, id });
      live.add(id);
      manifestParts.push(["c", id, part.length]);
      continue;
    }
    const entries: unknown[] = [];
    for (const piece of part.pieces) {
      if (piece.id === null || piece.bytes.length <= INLINE_PIECE_MAX_BYTES) {
        entries.push(encodeBase64(piece.bytes));
        continue;
      }
      const known = state.pieces.get(piece.id);
      let id: string;
      if (known !== undefined && state.live.has(known)) id = known;
      else {
        id = fresh();
        writes.push([partKey(storageKey, id), encodeBase64(piece.bytes)]);
      }
      pieces.push([piece.id, id]);
      live.add(id);
      entries.push([id, base64Length(piece.bytes.length)]);
    }
    manifestParts.push(["b", entries]);
  }
  const manifest = `{"savedAt":${JSON.stringify(plan.savedAt)},"${MARK}":1,"length":${plan.length},"parts":${JSON.stringify(manifestParts)}}`;
  return { storageKey, manifest, writes, live, runs, pieces };
}

/** Store a prepared write: its new parts, then its manifest, then remove the
 * parts no longer named. On a failure nothing the previous manifest names is
 * touched (the parts written here are taken back) and the error is thrown. */
export function commitPageDraftSegments(storage: Storage, prepared: PreparedPageDraftSegments): void {
  const attempt = () => {
    const written: string[] = [];
    try {
      for (const [key, value] of prepared.writes) {
        storage.setItem(key, value);
        written.push(key);
      }
      storage.setItem(prepared.storageKey, prepared.manifest);
    } catch (error) {
      for (const key of written) {
        try { storage.removeItem(key); } catch { /* best effort */ }
      }
      throw error;
    }
  };
  try {
    attempt();
  } catch (error) {
    // The tab's storage is full: other pages' retired parts (kept only for
    // their next copy) give way, once, before this copy does.
    if (!purgeRetiredPageDraftSegments(storage, prepared.storageKey)) throw error;
    attempt();
  }
  const state = keyState(prepared.storageKey);
  if (state.retireTimer !== null) clearTimeout(state.retireTimer);
  state.retireTimer = null;
  for (const id of state.live) {
    if (prepared.live.has(id)) continue;
    try { storage.removeItem(partKey(prepared.storageKey, id)); } catch { /* best effort */ }
  }
  state.live = new Set(prepared.live);
  state.runs = prepared.runs;
  for (const [piece, id] of prepared.pieces) state.pieces.set(piece, id);
  if (!state.swept) {
    state.swept = true;
    sweepParts(storage, prepared.storageKey, state.live);
  }
}

/** Remove parts stored under `storageKey` that `keep` does not name (left by
 * an earlier page load of this tab, whose manifest has been replaced). */
function sweepParts(storage: Storage, storageKey: string, keep: ReadonlySet<string>): void {
  try {
    const prefix = `${storageKey}#`;
    const stale: string[] = [];
    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      if (key?.startsWith(prefix) && !keep.has(key.slice(prefix.length))) stale.push(key);
    }
    for (const key of stale) storage.removeItem(key);
  } catch {
    // best effort
  }
}

/** Whether a stored draft value is a manifest (and not a whole copy). */
export function isPageDraftManifest(raw: string | null | undefined): raw is string {
  return typeof raw === "string" && MANIFEST_HEAD.test(raw);
}

/** The ids of the parts a manifest names. */
function manifestIds(raw: string): string[] {
  try {
    const parsed = JSON.parse(raw) as { parts?: unknown };
    const ids: string[] = [];
    for (const part of Array.isArray(parsed.parts) ? parsed.parts : []) {
      if (!Array.isArray(part)) continue;
      if (part[0] === "c" && typeof part[1] === "string") ids.push(part[1]);
      if (part[0] === "b" && Array.isArray(part[1])) {
        for (const entry of part[1]) if (Array.isArray(entry) && typeof entry[0] === "string") ids.push(entry[0]);
      }
    }
    return ids;
  } catch {
    return [];
  }
}

/**
 * Remove the parts of the copy stored under `storageKey` (its manifest is
 * `current`, the value under the key now), before or after the key itself
 * is removed or a whole copy is written there. Whole copies have no parts.
 */
export function removePageDraftSegments(storage: Storage, storageKey: string, current: string | null): void {
  const state = keyStates.get(storageKey);
  const ids = new Set<string>(state ? state.live : []);
  if (isPageDraftManifest(current)) for (const id of manifestIds(current)) ids.add(id);
  if (state) {
    if (state.retireTimer !== null) clearTimeout(state.retireTimer);
    state.retireTimer = null;
    state.live = new Set();
    state.runs = new Map();
  }
  for (const id of ids) {
    try { storage.removeItem(partKey(storageKey, id)); } catch { /* best effort */ }
  }
}

/** How long the parts of a retired copy stay for the next copy to take
 * (a pen writing for a minute is not saved meanwhile). Storage that runs
 * full removes them sooner (purgeRetiredPageDraftSegments). */
export const PAGE_DRAFT_RETIRED_PARTS_MS = 120_000;

/**
 * Remove the parts retired copies left (retirePageDraftSegments), except
 * under `keep`: what a write that found the tab's storage full does before
 * it gives up. Whether anything was removed.
 */
export function purgeRetiredPageDraftSegments(storage: Storage, keep?: string): boolean {
  let removed = false;
  for (const [storageKey, state] of keyStates) {
    if (storageKey === keep || state.retireTimer === null || state.live.size === 0) continue;
    removePageDraftSegments(storage, storageKey, null);
    removed = true;
  }
  return removed;
}

/**
 * The copy under `storageKey` was retired (its manifest, `current` when it
 * was removed, is gone: a save took the draft). Without a manifest its parts
 * are no copy anyone reads; the parts this window wrote stay a moment for
 * the next copy of the page, which is the same blocks and state but for the
 * strokes since (removing them made each save after a stored one write the
 * whole copy again), and are removed then unless a new copy names them.
 * Parts this window did not write go now.
 */
export function retirePageDraftSegments(storage: Storage, storageKey: string, current: string | null): void {
  const state = keyStates.get(storageKey);
  if (isPageDraftManifest(current)) {
    for (const id of manifestIds(current)) {
      if (state?.live.has(id)) continue;
      try { storage.removeItem(partKey(storageKey, id)); } catch { /* best effort */ }
    }
  }
  if (!state || state.live.size === 0 || state.retireTimer !== null) return;
  state.retireTimer = setTimeout(() => {
    state.retireTimer = null;
    removePageDraftSegments(storage, storageKey, null);
  }, PAGE_DRAFT_RETIRED_PARTS_MS);
}

/** Whether this window has parts stored under `storageKey`. */
export function hasPageDraftSegments(storageKey: string): boolean {
  return (keyStates.get(storageKey)?.live.size ?? 0) > 0;
}

/** A stored part, when it is there at the length its manifest says. */
function readPart(storage: Storage, storageKey: string, id: unknown, length: unknown): string | null {
  if (typeof id !== "string" || typeof length !== "number") return null;
  const value = storage.getItem(partKey(storageKey, id));
  return value !== null && value.length === length ? value : null;
}

/**
 * The whole copy a value stored under `storageKey` stands for: the value
 * itself when it is a whole copy, the copy put together from its parts when
 * it is a manifest, or null when a part is missing or not the length the
 * manifest says (or the whole is not).
 */
export function expandPageDraft(storage: Storage, storageKey: string, raw: string | null): string | null {
  if (!isPageDraftManifest(raw)) return raw;
  try {
    const parsed = JSON.parse(raw) as { savedAt?: unknown; length?: unknown; parts?: unknown };
    if (typeof parsed.savedAt !== "number" || typeof parsed.length !== "number" || !Array.isArray(parsed.parts)) return null;
    const out: string[] = [`{"savedAt":${JSON.stringify(parsed.savedAt)}`];
    for (const part of parsed.parts) {
      if (typeof part === "string") { out.push(part); continue; }
      if (!Array.isArray(part)) return null;
      if (part[0] === "c") {
        const value = readPart(storage, storageKey, part[1], part[2]);
        if (value === null) return null;
        out.push(value);
        continue;
      }
      if (part[0] !== "b" || !Array.isArray(part[1])) return null;
      const chunks: Uint8Array[] = [];
      let size = 0;
      for (const entry of part[1]) {
        const text = typeof entry === "string" ? entry
          : Array.isArray(entry) ? readPart(storage, storageKey, entry[0], entry[1]) : null;
        if (text === null) return null;
        const bytes = decodeBase64(text);
        if (!bytes) return null;
        chunks.push(bytes);
        size += bytes.length;
      }
      const whole = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { whole.set(chunk, offset); offset += chunk.length; }
      out.push(encodeBase64(whole));
    }
    const text = out.join("");
    return text.length === parsed.length ? text : null;
  } catch {
    return null;
  }
}
