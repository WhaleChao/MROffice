/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { openPdfRecoveryIdbStore, type PdfRecoveryIdbRun } from "./working-recovery-shadow";

export type PdfNativeTextTypingDraft = Readonly<{
  /** The line as the window named it when the words were kept. */
  id: string;
  /** The page of the file the line is on. */
  pageNumber: number;
  /** Where the line stands on its page (fractions of the page's box). */
  at: Readonly<{ x: number; y: number }> | null;
  /** What the line read when it was opened: the draft is for a line on this
   * page that still reads exactly this. */
  lineText: string;
  /** The words as the editor held them. */
  value: string;
  caret: Readonly<{ start: number; end: number }> | null;
  /** The line was left with these words (they were on their way into the
   * page's queue): brought back, they are finished again, not left open. */
  left: boolean;
  keptAt: number;
  /** How many documents have taken the draft over without placing it. */
  tries: number;
}>;

type Snapshot = Readonly<{ owner: string; seq: number; drafts: readonly PdfNativeTextTypingDraft[] }>;
export type PdfNativeTextTypingDraftShadowRecord = Snapshot & Readonly<{
  name: string;
  scope: string;
  at: number;
  /** The holders whose drafts this one took over (and theirs): a copy of
   * one of those holders' tab stores, met later, is older than this record. */
  took: readonly string[];
}>;
export type PdfNativeTextTypingDraftShadow = {
  put(record: PdfNativeTextTypingDraftShadowRecord): Promise<void>;
  all(): Promise<PdfNativeTextTypingDraftShadowRecord[]>;
  remove(name: string): Promise<void>;
};
type TabLocks = Pick<LockManager, "request">;

const SESSION_PREFIX = "mrai.pdf.native-text-typing-drafts.v1:";
const TAB_LOCK_PREFIX = "mrnote:pdf-typing-drafts:tab:";
const TAKE_LOCK_PREFIX = "mrnote:pdf-typing-drafts:take:";
const MAX_DRAFTS = 32;
const MAX_CHARS = 32 * 1024;



export type PdfNativeTextTypingDraftLimits = Readonly<{
  drafts: number;
  /** Of one record's words (each of its two texts). */
  chars: number;
  /** Of all the scope's records together: the tab's store is shared with
   * the page's queue of unsaved edits, which must not be crowded out. */
  total?: number;
}>;
/** A draft nobody placed is let go after this long. */
const KEPT_MS = 30 * 24 * 60 * 60 * 1000;
/** The disk copy follows the keys by this much, and by no more than
 * `WRITE_WITHIN_MS` while they keep coming. */
const WRITE_AFTER_MS = 250;
const WRITE_WITHIN_MS = 1_000;
/** A gone holder's word that it kept nothing any more is let go after this
 * long: no browser brings a tab back with a copy of its store that old. */
const NOTHING_KEPT_MS = 7 * 24 * 60 * 60 * 1000;
const MAX_TOOK = 16;

const newId = (): string => {
  try { return crypto.randomUUID(); } catch { return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 14)}`; }
};

function isDraft(value: unknown): value is PdfNativeTextTypingDraft {
  if (!value || typeof value !== "object") return false;
  const draft = value as Record<string, unknown>;
  const at = draft.at as Record<string, unknown> | null;
  const caret = draft.caret as Record<string, unknown> | null;
  return typeof draft.id === "string" && Number.isSafeInteger(draft.pageNumber) && (draft.pageNumber as number) > 0 &&
    typeof draft.lineText === "string" && typeof draft.value === "string" && typeof draft.left === "boolean" &&
    typeof draft.keptAt === "number" && Number.isSafeInteger(draft.tries) &&
    (at === null || (typeof at === "object" && Number.isFinite(at.x) && Number.isFinite(at.y))) &&
    (caret === null || (typeof caret === "object" && Number.isSafeInteger(caret.start) && Number.isSafeInteger(caret.end)));
}

function isSnapshot(value: unknown): value is Snapshot {
  if (!value || typeof value !== "object") return false;
  const snapshot = value as Record<string, unknown>;
  return typeof snapshot.owner === "string" && Number.isSafeInteger(snapshot.seq) &&
    Array.isArray(snapshot.drafts) && snapshot.drafts.every(isDraft);
}

/** The drafts' disk copy in this browser's IndexedDB; null where there is none. */
export function createIndexedDbPdfNativeTextTypingDraftShadow(): PdfNativeTextTypingDraftShadow | null {
  const run: PdfRecoveryIdbRun | null = openPdfRecoveryIdbStore("mrnote-pdf-typing-drafts", "tabs");
  if (!run) return null;
  const isRecord = (value: unknown): value is PdfNativeTextTypingDraftShadowRecord => {
    const record = value as { name?: unknown; scope?: unknown; at?: unknown; took?: unknown };
    return isSnapshot(value) && typeof record.name === "string" && typeof record.scope === "string" &&
      typeof record.at === "number" && Array.isArray(record.took) && record.took.every((owner) => typeof owner === "string");
  };
  return {
    async put(record) {
      await run("readwrite", (store) => { store.put(record); return () => undefined; });
    },
    async all() {
      const found = await run("readonly", (store) => {
        const request = store.getAll();
        return () => request.result as unknown[];
      });
      return found.filter(isRecord);
    },
    async remove(name) {
      await run("readwrite", (store) => { store.delete(name); return () => undefined; });
    },
  };
}

export type PdfNativeTextTypingDrafts = {
  /** Keeps the words for the line (by `id`), in place of what was kept for
   * it. False when they are too long to keep or the tab has no room. */
  keep(draft: Omit<PdfNativeTextTypingDraft, "keptAt" | "tries" | "left"> & { left?: boolean; tries?: number }): boolean;
  /** The line was left with its words: they stay kept until `drop`. */
  markLeft(id: string): void;
  drop(id: string): void;
  has(id: string): boolean;
  list(): PdfNativeTextTypingDraft[];
  /** Writes the disk copy now (the document is hidden, or going). */
  flush(): void;
  /**
   * The drafts left by documents that are gone -- this tab's before a
   * reload, a tab closed, discarded or killed -- become this document's and
   * are answered, each counted one more try. They stay kept here until
   * dropped.
   */
  take(): Promise<PdfNativeTextTypingDraft[]>;
};

/** The typing drafts of one document (a tab's page load). */
export function createPdfNativeTextTypingDraftTab(env: {
  session: Storage | null;
  shadow?: PdfNativeTextTypingDraftShadow | null;
  locks?: TabLocks | null;
  now?: () => number;
  instance?: string;
}) {
  const { session } = env;
  const locks = env.locks ?? null;
  // Without the locks nobody can tell a holder that is gone from one that lives.
  const shadow = locks ? env.shadow ?? null : null;
  const now = env.now ?? Date.now;
  const instance = env.instance ?? newId();
  let tabLockAsked = false;
  const holdTabLock = () => {
    if (tabLockAsked || !locks) return;
    tabLockAsked = true;
    void Promise.resolve()
      .then(() => locks.request(TAB_LOCK_PREFIX + instance, () => new Promise<void>(() => undefined)))
      .catch(() => { tabLockAsked = false; });
  };
  const ownerGone = async (owner: string): Promise<boolean> => {
    if (!locks) return true;
    try { return await locks.request(TAB_LOCK_PREFIX + owner, { ifAvailable: true }, async (lock) => lock !== null); }
    catch { return false; }
  };
  const scopes = new Map<string, PdfNativeTextTypingDrafts>();

  const open = (scope: string, limits?: PdfNativeTextTypingDraftLimits): PdfNativeTextTypingDrafts => {
    const maxDrafts = limits?.drafts ?? MAX_DRAFTS;
    const maxChars = limits?.chars ?? MAX_CHARS;
    const maxTotal = limits?.total ?? Number.POSITIVE_INFINITY;
    const sessionKey = SESSION_PREFIX + scope;
    const shadowName = `${scope}|${instance}`;
    let drafts: PdfNativeTextTypingDraft[] = [];
    let seq = 0;
    let timer: ReturnType<typeof setTimeout> | null = null;
    let waitingSince: number | null = null;
    let writing: Promise<void> = Promise.resolve();
    let shadowed = false;
    let took: string[] = [];
    const readSession = (): Snapshot | null => {
      try {
        const value: unknown = JSON.parse(session?.getItem(sessionKey) ?? "null");
        return isSnapshot(value) ? value : null;
      } catch { return null; }
    };
    const writeShadow = (): Promise<void> => {
      if (timer !== null) { clearTimeout(timer); timer = null; }
      waitingSince = null;
      if (!shadow) return writing;
      writing = writing.then(async () => {
        // What is kept now, not what was when the write was asked for. Once
        // something was kept, "nothing any more" is written too, at its
        // count: a browser that brings the tab back with an older copy of
        // the tab's own store finds the later word here.
        if (!drafts.length && !shadowed) return;
        shadowed = true;
        await shadow.put({ name: shadowName, scope, owner: instance, seq, at: now(), drafts, took });
      }).catch(() => undefined);
      return writing;
    };
    const persist = () => {
      seq += 1;
      if (drafts.length) holdTabLock();
      try {
        if (drafts.length) session?.setItem(sessionKey, JSON.stringify({ owner: instance, seq, drafts } satisfies Snapshot));
        else session?.removeItem(sessionKey);
      } catch { /* a full or blocked tab store: the disk copy alone */ }
      if (!shadow) return;
      const at = now();
      waitingSince ??= at;
      if (timer !== null) clearTimeout(timer);
      // Nothing kept any more is said at once: a draft placed or thrown away
      // is not brought back by a browser that goes down the moment after.
      const wait = drafts.length ? Math.max(0, Math.min(WRITE_AFTER_MS, waitingSince + WRITE_WITHIN_MS - at)) : 0;
      timer = setTimeout(() => { timer = null; void writeShadow(); }, wait);
    };
    const store: PdfNativeTextTypingDrafts = {
      keep(input) {
        if (!input.id || input.value.length > maxChars || input.lineText.length > maxChars) return false;
        const existing = drafts.find((draft) => draft.id === input.id);
        if (!existing && drafts.length >= maxDrafts) return false;
        if (Number.isFinite(maxTotal)) {
          let others = 0;
          for (const draft of drafts) if (draft !== existing) others += draft.value.length + draft.lineText.length;
          if (others + input.value.length + input.lineText.length > maxTotal) return false;
        }
        const next: PdfNativeTextTypingDraft = Object.freeze({
          id: input.id, pageNumber: input.pageNumber,
          at: input.at ? Object.freeze({ x: input.at.x, y: input.at.y }) : existing?.at ?? null,
          lineText: input.lineText, value: input.value,
          caret: input.caret ? Object.freeze({ start: input.caret.start, end: input.caret.end }) : null,
          left: input.left ?? false, keptAt: now(), tries: input.tries ?? existing?.tries ?? 0,
        });
        if (existing && existing.value === next.value && existing.left === next.left && existing.lineText === next.lineText &&
          existing.pageNumber === next.pageNumber && existing.caret?.start === next.caret?.start &&
          existing.caret?.end === next.caret?.end && existing.tries === next.tries && existing.at?.x === next.at?.x && existing.at?.y === next.at?.y) return true;
        drafts = [...drafts.filter((draft) => draft.id !== input.id), next];
        persist();
        return true;
      },
      markLeft(id) {
        const existing = drafts.find((draft) => draft.id === id);
        if (!existing || existing.left) return;
        drafts = drafts.map((draft) => draft === existing ? Object.freeze({ ...draft, left: true, keptAt: now() }) : draft);
        persist();
      },
      drop(id) {
        if (!drafts.some((draft) => draft.id === id)) return;
        drafts = drafts.filter((draft) => draft.id !== id);
        persist();
      },
      has: (id) => drafts.some((draft) => draft.id === id),
      list: () => [...drafts],
      flush() { if (timer !== null) void writeShadow(); },
      async take() {
        // By the holder each was kept by, as that holder last had them: the
        // tab's own copy (written as the keys landed, but as old as the
        // browser's last note of the tab when it is the browser that brings
        // the tab back) or the disk's, whichever counts more changes.
        const left = new Map<string, Snapshot>();
        let own = readSession();
        if (own?.owner === instance) own = null;
        const forgetOwn = () => {
          own = null;
          if (!drafts.length) { try { session?.removeItem(sessionKey); } catch { /* read again next time */ } }
        };
        // A duplicated tab's copy of a living tab's drafts is not this one's.
        if (own && !(await ownerGone(own.owner))) forgetOwn();
        const taken: PdfNativeTextTypingDraft[] = [];
        const adopt = async (spent: readonly PdfNativeTextTypingDraftShadowRecord[]) => {
          const mine = new Set(drafts.map((draft) => draft.id));
          const newest = new Map<string, PdfNativeTextTypingDraft>();
          for (const snapshot of left.values()) {
            for (const draft of snapshot.drafts) {
              if (mine.has(draft.id) || now() - draft.keptAt > KEPT_MS) continue;
              const other = newest.get(draft.id);
              if (!other || other.keptAt < draft.keptAt) newest.set(draft.id, draft);
            }
          }
          for (const draft of newest.values()) {
            if (drafts.length >= maxDrafts) break;
            const mineNow = Object.freeze({ ...draft, tries: draft.tries + 1 });
            drafts = [...drafts, mineNow];
            taken.push(mineNow);
          }
          took = [...new Set([...left.keys(), ...spent.flatMap((record) => record.took), ...took])]
            .filter((owner) => owner !== instance).slice(0, MAX_TOOK);
          if (taken.length || own) persist();
          if (!shadow) return;
          // Kept under this holder before the gone holders' copies go.
          await writeShadow();
          await Promise.all(spent.map((record) => shadow.remove(record.name).catch(() => undefined)));
        };
        const ownAlone = async () => {
          if (!own) return;
          left.set(own.owner, own);
          await adopt([]);
        };
        if (!shadow || !locks) {
          await ownAlone();
          return taken;
        }
        let read = false;
        try {
          // One taker at a time: two tabs opened on the page take a gone
          // tab's drafts once between them.
          await locks.request(TAKE_LOCK_PREFIX + scope, async () => {
            const records = await shadow.all();
            read = true;
            // Taken over since by a later document: the copy is an old one.
            if (own && records.some((record) => record.scope === scope && record.took.includes(own!.owner))) forgetOwn();
            if (own) left.set(own.owner, own);
            const spent: PdfNativeTextTypingDraftShadowRecord[] = [];
            for (const record of records) {
              if (record.owner === instance) continue;
              const kept = record.drafts.some((draft) => now() - draft.keptAt <= KEPT_MS);
              if (record.scope !== scope || !kept) {
                // Nobody's any more, here or on another page: swept in passing.
                if (!kept && (record.drafts.length || now() - record.at > NOTHING_KEPT_MS) && await ownerGone(record.owner)) {
                  void shadow.remove(record.name).catch(() => undefined);
                }
                if (record.scope !== scope || record.drafts.length) continue;
              }
              if (record.owner !== own?.owner && !(await ownerGone(record.owner))) continue;
              const other = left.get(record.owner);
              if (!other || other.seq < record.seq) left.set(record.owner, record);
              if (kept) spent.push(record);
            }
            if (left.size) await adopt(spent);
          });
        } catch {
          // The disk did not answer: what the tab itself kept is still taken.
          if (!read) await ownAlone();
        }
        return taken;
      },
    };
    return store;
  };
  return {
    /** `limits` are the scope's from its first opening on. */
    scope(scope: string, limits?: PdfNativeTextTypingDraftLimits): PdfNativeTextTypingDrafts {
      let store = scopes.get(scope);
      if (!store) scopes.set(scope, (store = open(scope, limits)));
      return store;
    },
    flush(): void { for (const store of scopes.values()) store.flush(); },
  };
}

type PdfNativeTextTypingDraftTab = ReturnType<typeof createPdfNativeTextTypingDraftTab>;
let thisTab: PdfNativeTextTypingDraftTab | null = null;








export function pdfNativeTextTypingDrafts(input: {
  principal: string | null | undefined;
  pageId: string;
  dataItemId: string | null | undefined;
  lane?: string;
  limits?: PdfNativeTextTypingDraftLimits;
}): PdfNativeTextTypingDrafts | null {
  if (typeof window === "undefined" || !input.principal || !input.pageId || !input.dataItemId) return null;
  if (!thisTab) {
    let session: Storage | null = null;
    let locks: TabLocks | null = null;
    let shadow: PdfNativeTextTypingDraftShadow | null = null;
    try {
      session = window.sessionStorage;
      locks = typeof navigator !== "undefined" && typeof navigator.locks?.request === "function" ? navigator.locks : null;
      shadow = locks ? createIndexedDbPdfNativeTextTypingDraftShadow() : null;
    } catch { /* storage blocked */ }
    if (!session && !shadow) return null;
    const tab = createPdfNativeTextTypingDraftTab({ session, locks, shadow });
    window.addEventListener("pagehide", () => tab.flush());
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") tab.flush(); });
    thisTab = tab;
  }
  const scope = [input.principal, input.pageId, input.dataItemId].map(encodeURIComponent).join(":");
  return thisTab.scope(input.lane ? `${scope}|${encodeURIComponent(input.lane)}` : scope, input.limits);
}
