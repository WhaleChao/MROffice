/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { captureSavedPageHotCacheScope, isSavedPageHotCacheScopeCurrent } from "@mrnote/core/web/lib/saved-page-hot-cache.ts";
import type { PdfWorkingRecovery } from "./native-working-session";
import { createIndexedDbPdfWorkingRecoveryShadow, type PdfWorkingRecoveryShadow,
  type PdfWorkingRecoveryShadowRecord } from "./working-recovery-shadow";

const PREFIX = "mrnote:pdf-working-recovery:v1:";
const CLAIM_PREFIX = "mrnote:pdf-working-recovery:claim:v1:";
const DATA_PREFIX = "mrnote:pdf-working-recovery:data:v1:";
const WITH_PREFIX = "mrnote:pdf-working-recovery:with:v1:";
/** A shadowed queue's count of changes; `<count>!<when>` once it is finished. */
const SEQ_PREFIX = "mrnote:pdf-working-recovery:seq:v1:";
/** In place of a queue's kept records, when they could not be mirrored. */
const NOT_MIRRORED = "!";
const TAB_LOCK_PREFIX = "mrnote:pdf-working-recovery:tab:";
const TAKE_LOCK_PREFIX = "mrnote:pdf-working-recovery:take:";
const MAX_SERIALIZED_CHARS = 4 * 1024 * 1024;
/** localStorage is a few megabytes for the whole site: a larger queue stays
 * in the tab alone, as every queue did before the mirror. */
const MAX_MIRRORED_CHARS = 1024 * 1024;
/** A claim this young may belong to a document that has not been granted its
 * tab lock yet: its holder is taken as living for that long. */
const CLAIM_SETTLES_MS = 5_000;
/** A queue no tab has held for this long, on a page this browser never
 * opened again, is removed with the next sweep. */
const ORPHAN_KEPT_MS = 90 * 24 * 60 * 60 * 1000;
/** The word that a queue was finished is kept this long after it: it only
 * has to outlast the mirror's own delay in reaching the disk. */
const FINISHED_KEPT_MS = 10 * 60 * 1000;
/** A taking waits this long for the shadow, and goes on by the mirror alone
 * past it: the window that asks gives a queue left behind two seconds. Both
 * are counted as the document runs (`afterPdfRecoveryRunningTime`). */
const SHADOW_WAITED_MS = 1_000;
const RUNNING_STEP_MS = 50;













export function afterPdfRecoveryRunningTime(ms: number, then: () => void): () => void {
  let left = ms;
  let last = Date.now();
  let timer: ReturnType<typeof setTimeout> | undefined;
  const wait = () => {
    const due = Math.max(1, Math.min(RUNNING_STEP_MS, left));
    timer = setTimeout(() => {
      const at = Date.now();
      const hidden = typeof document !== "undefined" && document.visibilityState === "hidden";
      // (A timer is never early: at least what it was set for has passed.)
      const passed = Math.max(due, at - last);
      left -= hidden ? passed : Math.min(passed, 2 * RUNNING_STEP_MS);
      last = at;
      if (left > 0) { wait(); return; }
      timer = undefined;
      then();
    }, due);
  };
  if (left > 0) wait(); else timer = setTimeout(() => { timer = undefined; then(); }, 0);
  return () => { if (timer !== undefined) clearTimeout(timer); timer = undefined; };
}

type Claim = { owner: string | null; at: number };
type TabLocks = Pick<LockManager, "request">;
type StoreInput = { pageId: string; dataItemId: string };
/** Records the window keeps in the tab beside a queue, by name: the tab key each lives under. */
type KeptWith = Readonly<Record<string, string>>;

const newId = (): string => {
  try { return crypto.randomUUID(); } catch { return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 14)}`; }
};
/** The queue a kept record is mirrored as; read from the record's head (the
 * record may be megabytes, and this runs wherever a store is made). */
const QUEUE_OF = /^\{"version":1,"principal":"[^"\\]*","pageId":"[^"\\]*","dataItemId":"[^"\\]*","queue":"([0-9a-z-]{8,64})",/;
const queueOf = (raw: string | null): string | null =>
  raw === null ? null : QUEUE_OF.exec(raw.slice(0, 2048))?.[1] ?? null;

























































export function createPdfWorkingRecoveryTab(env: {
  session: Storage;
  durable?: Storage | null;
  locks?: TabLocks | null;
  shadow?: PdfWorkingRecoveryShadow | null;
  now?: () => number;
  instance?: string;
}) {
  const { session } = env;
  const durable = env.durable && env.locks ? env.durable : null;
  const locks = durable ? env.locks ?? null : null;
  const shadow = durable ? env.shadow ?? null : null;
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
  const readClaim = (key: string): Claim | null => {
    try {
      const value = JSON.parse(durable!.getItem(key) ?? "null");
      return value && (value.owner === null || typeof value.owner === "string") && typeof value.at === "number"
        ? { owner: value.owner, at: value.at } : null;
    } catch { return null; }
  };
  const writeClaim = (key: string, owner: string | null) =>
    durable!.setItem(key, JSON.stringify({ owner, at: now() }));
  const claimsUnder = (prefix: string) => {
    const found: { key: string; claim: Claim }[] = [];
    for (let index = 0; index < durable!.length; index += 1) {
      const key = durable!.key(index);
      const claim = key?.startsWith(prefix) ? readClaim(key) : null;
      if (key && claim) found.push({ key, claim });
    }
    return found;
  };
  const ownerGone = async (owner: string): Promise<boolean> => {
    try { return await locks!.request(TAB_LOCK_PREFIX + owner, { ifAvailable: true }, async (lock) => lock !== null); }
    catch { return false; }
  };
  /** A shadowed queue's count of changes in the mirror, and when it was finished. */
  const readStamp = (name: string): { seq: number; goneAt: number | null } | null => {
    const match = /^(\d{1,15})(?:!(\d{1,15}))?$/.exec(durable!.getItem(SEQ_PREFIX + name) ?? "");
    return match ? { seq: Number(match[1]), goneAt: match[2] === undefined ? null : Number(match[2]) } : null;
  };
  /**
   * Queues (`scope:queue`) this document took again from the tab's own copy
   * (a reload; a tab the browser restored) and has not told the shadow of:
   * the copy, and the mirror's count for it then. Until the shadow has been
   * asked, the copy may be older than what the shadow keeps.
   */
  const resettled = new Map<string, { raw: string; seq: number }>();
  /** Queues the shadow has later than the mirror that could not be brought
   * back yet (their holder may live): not taken as the mirror has them. */
  const heldBack = new Set<string>();
  /** Queues this document held and finished. */
  const finishedHere = new Set<string>();
  /** The names of the records mirrored with a queue (`scope:queue`). */
  const mirroredWith = (name: string): string[] => {
    const prefix = WITH_PREFIX + name + ":";
    const found: string[] = [];
    for (let index = 0; index < durable!.length; index += 1) {
      const key = durable!.key(index);
      if (key?.startsWith(prefix)) found.push(decodeURIComponent(key.slice(prefix.length)));
    }
    return found;
  };
  const dropMirroredWith = (name: string) => {
    for (const kept of mirroredWith(name)) durable!.removeItem(WITH_PREFIX + name + ":" + encodeURIComponent(kept));
  };
  /** Queues (`scope:queue`) being written to the shadow: true when the
   * mirror changed again meanwhile. One write at a time for each, always of
   * what the mirror then holds. */
  const shadowing = new Map<string, boolean>();
  const sendShadow = (name: string): void => {
    if (!shadow) return;
    if (shadowing.has(name)) { shadowing.set(name, true); return; }
    shadowing.set(name, false);
    void (async () => {
      do {
        shadowing.set(name, false);
        let record: PdfWorkingRecoveryShadowRecord | null = null;
        let stamped: string | null = null;
        let goneAt: number | null = null;
        try {
          stamped = durable!.getItem(SEQ_PREFIX + name);
          const stamp = readStamp(name);
          if (stamp) {
            goneAt = stamp.goneAt;
            const data = stamp.goneAt === null ? durable!.getItem(DATA_PREFIX + name) : null;
            const claim = stamp.goneAt === null ? readClaim(CLAIM_PREFIX + name) : null;
            if (data !== null && claim) {
              const withIt: Record<string, string> = {};
              for (const kept of mirroredWith(name)) {
                withIt[kept] = durable!.getItem(WITH_PREFIX + name + ":" + encodeURIComponent(kept)) ?? "";
              }
              record = { name, seq: stamp.seq, at: now(), kept: { claim, data, with: withIt } };
            } else {
              record = { name, seq: stamp.seq, at: stamp.goneAt ?? now(), kept: null };
            }
          }
        } catch { /* the mirror is unreachable */ }
        if (!record) continue;
        try {
          const answer = await shadow.put(record);
          if (durable!.getItem(SEQ_PREFIX + name) !== stamped) continue;
          if (answer.stored) {
            // The shadow says it now: the mirror's own word that the queue
            // was finished has nothing left to hold back.
            if (!record.kept && goneAt !== null) durable!.removeItem(SEQ_PREFIX + name);
          } else if (answer.seq > record.seq && !resettled.has(name) && (record.kept
            ? record.kept.claim.owner === instance && readClaim(CLAIM_PREFIX + name)?.owner === instance
            : goneAt !== null && finishedHere.has(name))) {
            // The shadow counts further than the mirror for a queue this
            // document holds: it went on from an older copy of it (a tab
            // the browser restored, whose working copy was open before the
            // shadow could be asked). What the reader has on the page is
            // the queue now, and is written over the other.
            durable!.setItem(SEQ_PREFIX + name, record.kept ? String(answer.seq + 1) : `${answer.seq + 1}!${goneAt}`);
            shadowing.set(name, true);
          }
        } catch { /* the shadow is unreachable: the mirror alone, as before */ }
      } while (shadowing.get(name));
      shadowing.delete(name);
    })();
  };
  /** This document changed a mirrored queue it holds. */
  const stampKept = (name: string): void => {
    if (!shadow) return;
    try { durable!.setItem(SEQ_PREFIX + name, String((readStamp(name)?.seq ?? 0) + 1)); }
    catch { return; /* no room for the count: this change is not shadowed */ }
    sendShadow(name);
  };
  /** A mirrored queue was finished (saved, discarded, removed as nobody's);
   * `own` when this document held it. */
  const stampGone = (name: string, own = false): void => {
    if (!shadow) return;
    try {
      const stamp = readStamp(name);
      // One the shadow never had needs no word; one that has it keeps it.
      if (!stamp || stamp.goneAt !== null) return;
      if (own) finishedHere.add(name);
      durable!.setItem(SEQ_PREFIX + name, `${stamp.seq + 1}!${now()}`);
    } catch { return; }
    sendShadow(name);
  };
  const keptWithByScope = new Map<string, KeptWith>();
  /** Queues (`scope:queue`) this tab held a copy of that was another
   * document's: the tab's kept records are copies of that queue's too. */
  const copiesDropped = new Set<string>();
  /** Removes the mirror of a queue this document holds. */
  const dropMirror = (scopeKey: string, queue: string) => {
    try {
      durable!.removeItem(DATA_PREFIX + scopeKey + ":" + queue);
      durable!.removeItem(CLAIM_PREFIX + scopeKey + ":" + queue);
      dropMirroredWith(scopeKey + ":" + queue);
    } catch { /* the mirror is unreachable; the tab's record stands alone */ }
    resettled.delete(scopeKey + ":" + queue);
    stampGone(scopeKey + ":" + queue, true);
  };
  /**
   * Mirrors the records kept beside a mirrored queue this document holds,
   * as the tab has them now. When they cannot all be (no room, too large)
   * none is, and the queue is marked as one whose records are in its tab
   * alone, until a later write finds room. True when the mirror changed.
   */
  const mirrorKeptWith = (scopeKey: string, queue: string): boolean => {
    const name = scopeKey + ":" + queue;
    const marker = WITH_PREFIX + name + ":" + NOT_MIRRORED;
    try {
      if (durable!.getItem(DATA_PREFIX + name) === null) return false;
    } catch { return false; }
    let changed = false;
    const remove = (key: string) => {
      if (durable!.getItem(key) === null) return;
      changed = true;
      durable!.removeItem(key);
    };
    try {
      const keptWith = keptWithByScope.get(scopeKey) ?? {};
      let room = MAX_MIRRORED_CHARS;
      for (const [kept, tabKey] of Object.entries(keptWith)) {
        const key = WITH_PREFIX + name + ":" + encodeURIComponent(kept);
        const value = session.getItem(tabKey);
        if (value === null) { remove(key); continue; }
        room -= value.length;
        if (room < 0) throw new Error("The records kept with this queue are past the mirror's budget.");
        if (durable!.getItem(key) !== value) { changed = true; durable!.setItem(key, value); }
      }
      remove(marker);
    } catch {
      changed = true;
      try {
        dropMirroredWith(name);
        durable!.setItem(marker, "1");
      } catch {
        // Not even the mark: the queue is then not to be taken at all.
        try { durable!.removeItem(DATA_PREFIX + name); } catch { /* unreachable */ }
      }
    }
    return changed;
  };
  /** The tab's record for a page, once it is known to be this document's. */
  const settle = (scopeKey: string): string | null => {
    const key = PREFIX + scopeKey;
    const raw = session.getItem(key);
    const queue = durable ? queueOf(raw) : null;
    if (!queue) return raw;
    try {
      const claimKey = CLAIM_PREFIX + scopeKey + ":" + queue;
      const claim = readClaim(claimKey);
      if (claim?.owner === instance) return raw;
      if (claim && claim.owner === null) {
        const name = scopeKey + ":" + queue;
        const mirrored = shadow ? durable!.getItem(DATA_PREFIX + name) : null;
        // (A living browser keeps the two alike. Where they differ, a
        // browser that went down gave the tab back an older copy than the
        // mirror it wrote out, or a later one: the mirror is the one the
        // shadow speaks for, and is taken like any queue left behind.)
        if (mirrored === null || mirrored === raw) {
          writeClaim(claimKey, instance);
          holdTabLock();
          // The shadow is told once it has been asked (`confirmResettled`),
          // or with the window's next write.
          if (shadow) {
            const taken = { raw: raw!, seq: readStamp(name)?.seq ?? -1 };
            resettled.set(name, taken);
            // (A window asks within two seconds of opening the page, or
            // not at all: the tab's copy is the queue then.)
            setTimeout(() => { if (resettled.get(name) === taken) endResettled(scopeKey, queue); }, 3 * SHADOW_WAITED_MS);
          }
          return raw;
        }
      }
    } catch { return raw; }
    // Another document's queue (its holder, if gone, leaves it to `take`),
    // or one that was finished elsewhere: not replayed from this copy.
    session.removeItem(key);
    copiesDropped.add(scopeKey + ":" + queue);
    return null;
  };

  const behind = (record: PdfWorkingRecoveryShadowRecord) => (readStamp(record.name)?.seq ?? -1) < record.seq;
  const lingers = (name: string) =>
    durable!.getItem(DATA_PREFIX + name) !== null || durable!.getItem(CLAIM_PREFIX + name) !== null;
  const dropStale = (name: string) => {
    durable!.removeItem(DATA_PREFIX + name);
    durable!.removeItem(CLAIM_PREFIX + name);
    dropMirroredWith(name);
    durable!.removeItem(SEQ_PREFIX + name);
  };
  /** Whoever holds the queue in the mirror, or held it by the shadow, is
   * gone (a claim too young to tell is taken as a living holder's). */
  const unheld = async (claim: Claim | null): Promise<boolean> =>
    !claim || claim.owner === null ||
    (claim.owner !== instance && now() - claim.at >= CLAIM_SETTLES_MS && await ownerGone(claim.owner));
  /** The mirror of a queue as the shadow keeps it, its last holder's claim
   * with it: left to be taken as that holder left it. */
  const putBack = (name: string, seq: number, kept: NonNullable<PdfWorkingRecoveryShadowRecord["kept"]>): void => {
    try {
      dropMirroredWith(name);
      for (const [record, value] of Object.entries(kept.with)) {
        durable!.setItem(WITH_PREFIX + name + ":" + encodeURIComponent(record), value);
      }
      durable!.setItem(DATA_PREFIX + name, kept.data);
      durable!.setItem(CLAIM_PREFIX + name, JSON.stringify(kept.claim));
      durable!.setItem(SEQ_PREFIX + name, String(seq));
    } catch {
      // No room for all of it: none of it (an older queue is not to be
      // taken for this one); the shadow keeps it for a later opening.
      try { dropStale(name); } catch { /* the mirror is unreachable */ }
    }
  };

  /**
   * Brings the mirror back to what the shadow has, where the mirror is
   * behind it: this browser went down before it wrote the mirror out. Runs
   * before queues are taken, each queue under the lock a taking holds.
   */
  const reconcileOnce = async (): Promise<void> => {
    if (!shadow || !durable || !locks) return;
    let records: PdfWorkingRecoveryShadowRecord[];
    try { records = await shadow.all(); } catch { return; }
    const shadowed = new Map<string, PdfWorkingRecoveryShadowRecord>();
    const aged: string[] = [];
    for (const listed of records) {
      const { name } = listed;
      shadowed.set(name, listed);
      try {
        if (now() - listed.at >= (listed.kept ? ORPHAN_KEPT_MS : FINISHED_KEPT_MS)) aged.push(name);
        if (behind(listed) && (listed.kept ? await unheld(listed.kept.claim) : lingers(name)) &&
          await unheld(readClaim(CLAIM_PREFIX + name))) {
          // (Waited for: another document doing the same holds it for a moment.)
          await locks.request(TAKE_LOCK_PREFIX + name, async () => {
            // Read again now that nobody living writes it: all that its
            // holder wrote is in.
            const record = await shadow.get(name);
            if (!record || !behind(record) || !(await unheld(readClaim(CLAIM_PREFIX + name)))) return;
            // Finished, and the mirror still has it: a removal the browser
            // never wrote out.
            if (!record.kept) dropStale(name);
            else if (await unheld(record.kept.claim)) putBack(name, record.seq, record.kept);
          });
        }
      } catch { /* this queue is left as it is */ }
      // Still behind: what the mirror has of it is not the queue.
      try {
        if (behind(listed) && (listed.kept || lingers(name)) &&
          readClaim(CLAIM_PREFIX + name)?.owner !== instance) heldBack.add(name);
        else heldBack.delete(name);
      } catch { /* the mirror is unreachable */ }
    }
    for (const name of [...heldBack]) if (!shadowed.has(name)) heldBack.delete(name);
    // What is left to clear away waits for nobody.
    void (async () => {
      // The shadow's words that a queue was finished, past their time; and
      // the queues it kept that nothing came back for.
      for (const name of aged) {
        await locks.request(TAKE_LOCK_PREFIX + name, { ifAvailable: true }, async (lock) => {
          const record = lock ? await shadow.get(name) : null;
          if (record && !lingers(name) &&
            now() - record.at >= (record.kept ? ORPHAN_KEPT_MS : FINISHED_KEPT_MS)) await shadow.remove(name);
        }).catch(() => undefined);
      }
      // The mirror's own such words, left by a document that went before the
      // shadow had them; and counts whose queue is nowhere.
      try {
        const stamped: string[] = [];
        for (let index = 0; index < durable.length; index += 1) {
          const key = durable.key(index);
          if (key?.startsWith(SEQ_PREFIX)) stamped.push(key.slice(SEQ_PREFIX.length));
        }
        for (const name of stamped) {
          const stamp = readStamp(name);
          const record = shadowed.get(name);
          if (!stamp) continue;
          if (stamp.goneAt === null) {
            if (!lingers(name) && !record?.kept && !shadowing.has(name)) durable.removeItem(SEQ_PREFIX + name);
          } else if (record ? !record.kept && record.seq >= stamp.seq : now() - stamp.goneAt >= FINISHED_KEPT_MS) {
            durable.removeItem(SEQ_PREFIX + name);
          } else if (record) sendShadow(name);
        }
      } catch { /* the mirror is unreachable */ }
    })().catch(() => undefined);
  };
  /** One at a time in this document: a taking waits for all of it. */
  let reconciling: Promise<void> = Promise.resolve();
  const reconcile = (): Promise<void> => {
    reconciling = reconciling.then(reconcileOnce, reconcileOnce);
    return reconciling;
  };
  /** The shadow was told that this document holds a queue it took again. */
  const endResettled = (scopeKey: string, queue: string): void => {
    const name = scopeKey + ":" + queue;
    if (!resettled.delete(name)) return;
    try {
      if (readClaim(CLAIM_PREFIX + name)?.owner !== instance) return;
      mirrorKeptWith(scopeKey, queue);
    } catch { return; }
    stampKept(name);
  };
  /**
   * The queue this document took again from the tab's copy, asked of the
   * shadow before anything is made of it. A browser restoring its tabs
   * after it went down gives a tab the copy it last wrote out, and may have
   * written the mirror out as long ago: both then are older than the queue
   * the shadow keeps (or keep a queue the shadow knows was finished). The
   * tab's copy is dropped then and the mirror made what the shadow has, for
   * the taking that follows -- unless the window has already opened its
   * working copy on the tab's (`mayTake`), which is then the queue.
   */
  const confirmResettled = async (scopeKey: string, mayTake: () => boolean, wanted: () => boolean): Promise<void> => {
    if (!shadow || !durable || !locks) return;
    const key = PREFIX + scopeKey;
    const queue = queueOf(session.getItem(key));
    const name = queue ? scopeKey + ":" + queue : null;
    const taken = name ? resettled.get(name) : undefined;
    if (!queue || !name || !taken) return;
    const asIs = () => wanted() && resettled.get(name) === taken && session.getItem(key) === taken.raw &&
      readClaim(CLAIM_PREFIX + name)?.owner === instance;
    try {
      const listed = await shadow.get(name);
      if (listed && listed.seq > taken.seq && asIs()) {
        await locks.request(TAKE_LOCK_PREFIX + name, async () => {
          const record = await shadow.get(name);
          if (!record || record.seq <= taken.seq) return;
          const free = !record.kept || await unheld(record.kept.claim);
          if (!asIs() || !mayTake()) return;
          session.removeItem(key);
          copiesDropped.add(name);
          resettled.delete(name);
          if (!record.kept) dropStale(name);
          else if (free) putBack(name, record.seq, record.kept);
          else {
            // Its holder by the shadow may live: for a later opening.
            writeClaim(CLAIM_PREFIX + name, null);
            heldBack.add(name);
          }
        });
      }
    } catch { /* the shadow is unreachable: the tab's copy is the queue */ }
    if (wanted()) endResettled(scopeKey, queue);
  };

  /** False once the shadow kept a taking waiting past its time, until it
   * answers again: takings meanwhile go by the mirror alone. */
  let shadowAnswers = true;
  const scopeKeyOf = (principal: string, input: StoreInput) =>
    [principal, input.pageId, input.dataItemId].map(encodeURIComponent).join(":");

  /** `take`, by what the mirror has now. */
  const takeByMirror = async (scope: ReturnType<typeof captureSavedPageHotCacheScope>, scopeKey: string,
    mayTake: (keptWith?: readonly string[]) => boolean): Promise<boolean> => {
    if (!durable || !locks) return false;
    const key = PREFIX + scopeKey;
    const keptWith = keptWithByScope.get(scopeKey) ?? {};
    // The tab's own copy of a queue whose holder is gone (a browser
    // restoring its tabs after a crash) is that holder's: taken like any.
    try { settle(scopeKey); } catch { return false; }
    // (Unasked of the shadow, where it did not answer: the tab's copy stands.)
    const own = queueOf(session.getItem(key));
    if (own) endResettled(scopeKey, own);
    const free = () => isSavedPageHotCacheScopeCurrent(scope) && session.getItem(key) === null && mayTake();
    const found = claimsUnder(CLAIM_PREFIX + scopeKey + ":").sort((a, b) => b.claim.at - a.claim.at);
    for (const { key: claimKey, claim } of found) {
      if (!free()) return false;
      if (claim.owner === instance) continue;
      const name = claimKey.slice(CLAIM_PREFIX.length);
      if (heldBack.has(name)) continue;
      // A holder is asked for by its lock; one that only just wrote its
      // claim may not have been granted the lock yet and is taken as living.
      if (claim.owner !== null && (now() - claim.at < CLAIM_SETTLES_MS || !(await ownerGone(claim.owner)))) continue;
      const taken = await locks.request(TAKE_LOCK_PREFIX + name, { ifAvailable: true }, async (lock) => {
        if (!lock) return false;
        const again = readClaim(claimKey);
        if (!again || again.owner !== claim.owner || again.at !== claim.at || heldBack.has(name)) return false;
        const data = durable.getItem(DATA_PREFIX + name);
        if (data === null) { durable.removeItem(claimKey); stampGone(name); return false; }
        if (!free()) return false;
        const mirrored = mirroredWith(name);
        const withIt = mirrored.filter((kept) => kept !== NOT_MIRRORED);
        // The records in a tab restored with a copy of this queue are the
        // queue's own, as the browser kept them: the mirrored ones replace
        // them, and where none could be mirrored they stand.
        const ownCopy = copiesDropped.has(name);
        // Otherwise all of it or none. A queue whose records are not here,
        // or have no place in this tab (unnamed, or the tab keeps such a
        // record of its own for the page), stays for a tab that can take
        // them: alone it is, for one, a line's deletion without the
        // line's box.
        if (!ownCopy && (mirrored.length !== withIt.length ||
          !withIt.every((kept) => typeof keptWith[kept] === "string" && session.getItem(keptWith[kept]) === null))) {
          return false;
        }
        if (!mayTake(withIt)) return false;
        writeClaim(claimKey, instance);
        const before = new Map<string, string | null>();
        try {
          for (const kept of withIt) {
            const tabKey = keptWith[kept];
            const value = durable.getItem(WITH_PREFIX + name + ":" + encodeURIComponent(kept));
            if (typeof tabKey !== "string" || value === null) continue;
            before.set(tabKey, session.getItem(tabKey));
            session.setItem(tabKey, value);
          }
          session.setItem(key, data);
        } catch {
          for (const [tabKey, value] of before) {
            try { if (value === null) session.removeItem(tabKey); else session.setItem(tabKey, value); }
            catch { /* the tab's storage is gone */ }
          }
          durable.setItem(claimKey, JSON.stringify(again));
          return false;
        }
        holdTabLock();
        stampKept(name);
        return true;
      }).catch(() => false);
      if (taken) return true;
    }
    return false;
  };

  return {
    /** Uses the app's principal scope and tab-local storage, but a distinct key:
     * ordinary page-save ACKs cannot retire newer unsubmitted PDF operations.
     * This stores replay recipes, never credentials, PDF bytes or writer leases. */
    store(input: StoreInput) {
      const scope = captureSavedPageHotCacheScope();
      if (!scope.principal || !input.pageId || !input.dataItemId) {
        throw new Error("PDF local recovery requires an authenticated document scope.");
      }
      const principal = scope.principal;
      const scopeKey = scopeKeyOf(principal, input);
      const key = PREFIX + scopeKey;
      let expected = settle(scopeKey);
      const heldHere = (raw: string | null) => {
        const queue = durable ? queueOf(raw) : null;
        return !queue || readClaim(CLAIM_PREFIX + scopeKey + ":" + queue)?.owner === instance;
      };
      const check = () => {
        if (!isSavedPageHotCacheScopeCurrent(scope)) throw new Error("PDF local recovery account changed.");
        if (session.getItem(key) !== expected || !heldHere(expected)) {
          throw new Error("Another PDF window changed this recovery queue.");
        }
      };
      const decode = (raw: string | null): PdfWorkingRecovery | null => {
        if (raw === null) return null;
        if (raw.length > MAX_SERIALIZED_CHARS) throw new Error("PDF local recovery exceeds its storage budget.");
        const value = JSON.parse(raw);
        if (!value || value.version !== 1 || value.principal !== principal ||
          value.pageId !== input.pageId || value.dataItemId !== input.dataItemId ||
          value.recovery?.schemaVersion !== "mrai.pdf-working-recovery.v1" ||
          value.recovery.pageId !== input.pageId || !Array.isArray(value.recovery.operations) ||
          value.recovery.operations.length > 32) throw new Error("PDF local recovery identity is invalid.");
        // This shape check is not native authority. Restoration independently
        // proves every recipe against its real saved source and output hash.
        return value.recovery as PdfWorkingRecovery;
      };
      const encode = (queue: string | null, recovery: PdfWorkingRecovery) => JSON.stringify(queue
        ? { version: 1, principal, pageId: input.pageId, dataItemId: input.dataItemId, queue, recovery }
        : { version: 1, principal, pageId: input.pageId, dataItemId: input.dataItemId, recovery });
      const forget = () => {
        const queue = durable ? queueOf(session.getItem(key)) : null;
        session.removeItem(key);
        expected = null;
        if (queue && readClaim(CLAIM_PREFIX + scopeKey + ":" + queue)?.owner === instance) dropMirror(scopeKey, queue);
      };
      return {
        read(): PdfWorkingRecovery | null { check(); return decode(expected); },
        write(recovery: PdfWorkingRecovery): void {
          check();
          const had = durable ? queueOf(expected) : null;
          const queue = durable ? had ?? newId() : null;
          // What the window writes is the queue, whatever the shadow keeps.
          if (had) resettled.delete(scopeKey + ":" + had);
          let next = encode(queue, recovery);
          decode(next);
          const dataKey = DATA_PREFIX + scopeKey + ":" + queue;
          let mirrored = false;
          let before: string | null = null;
          if (durable && queue && next.length <= MAX_MIRRORED_CHARS) {
            try {
              before = had ? durable.getItem(dataKey) : null;
              durable.setItem(dataKey, next);
              if (!had) writeClaim(CLAIM_PREFIX + scopeKey + ":" + queue, instance);
              mirrored = true;
            } catch {
              // No room for the mirror: this queue stays in the tab alone.
              if (!had) dropMirror(scopeKey, queue);
            }
          }
          if (queue && !mirrored) next = encode(null, recovery);
          // setItem is atomic on quota errors. Do not clear the previous valid
          // queue if serialization, quota or scope validation fails.
          try {
            session.setItem(key, next);
          } catch (error) {
            if (durable && queue && mirrored) {
              try {
                if (before === null) dropMirror(scopeKey, queue);
                else durable.setItem(dataKey, before);
              } catch { /* the mirror keeps the newer queue; this document still holds its claim */ }
            }
            throw error;
          }
          expected = next;
          // (Its kept records follow as the window writes them, `keepWith`;
          // here only when the mirror is new.)
          if (mirrored) {
            holdTabLock();
            if (before === null) mirrorKeptWith(scopeKey, queue!);
            stampKept(scopeKey + ":" + queue);
          }
          // A mirror the queue outgrew would be taken over as an older queue.
          else if (queue && had) dropMirror(scopeKey, queue);
        },
        /** The reader chose 不保存 for edits AutoSave off held: the queue goes
         * with the window, as WPS discards an unsaved document. */
        discardUnsaved(): void {
          if (!isSavedPageHotCacheScopeCurrent(scope)) return;
          forget();
        },
        clearConfirmed(): void {
          check();
          const current = decode(expected);
          if (current?.operations.length) throw new Error("Unconfirmed PDF operations cannot be cleared.");
          forget();
        },
      };
    },
    /**
     * Names the records the window keeps in the tab beside the page's queue
     * and mirrors them as they now are; called again after each is written
     * or removed.
     */
    keepWith(input: StoreInput, keptWith: KeptWith): void {
      const scope = captureSavedPageHotCacheScope();
      if (!durable || !scope.principal || !input.pageId || !input.dataItemId) return;
      const scopeKey = scopeKeyOf(scope.principal, input);
      keptWithByScope.set(scopeKey, { ...keptWith });
      try {
        const queue = queueOf(session.getItem(PREFIX + scopeKey));
        // (One taken again and not yet asked of the shadow follows then.)
        if (queue && !resettled.has(scopeKey + ":" + queue) &&
          readClaim(CLAIM_PREFIX + scopeKey + ":" + queue)?.owner === instance &&
          mirrorKeptWith(scopeKey, queue)) stampKept(scopeKey + ":" + queue);
      } catch { /* the mirror is unreachable */ }
    },
    /**
     * Takes over, for this document, a queue of the page whose holder is
     * gone -- the newest one, when the tab holds none for the page -- with
     * the records mirrored beside it, placed under the tab keys this window
     * named (`keepWith`) before anything has read them. `mayTake` is asked
     * in the same turn the queue is placed in the tab, so a caller can
     * refuse once it has opened the page's working copy on an empty queue
     * (that copy would then refuse to write over this one); it is told the
     * names of the records that come with the queue.
     *
     * Where the shadow answers only after its time, what it brought back is
     * taken then, if `mayTake` still allows, and `takenLate` is told: before,
     * a queue put back in the mirror that late was nobody's until the page
     * was next opened.
     */
    async take(input: StoreInput, mayTake: (keptWith?: readonly string[]) => boolean = () => true,
      takenLate?: () => void): Promise<boolean> {
      const scope = captureSavedPageHotCacheScope();
      if (!durable || !locks || !scope.principal || !input.pageId || !input.dataItemId) return false;
      const scopeKey = scopeKeyOf(scope.principal, input);
      // A browser that went down with the mirror unwritten: first what the
      // shadow has of it, for as long as the shadow answers in time.
      /** The shadow's answer, where this taking went on without it. */
      let answersLate: Promise<void> | null = null;
      if (shadow) {
        let wanted = true;
        const asked = (async () => {
          await reconcile();
          if (!wanted) return;
          try { settle(scopeKey); } catch { return; }
          await confirmResettled(scopeKey, () => mayTake(), () => wanted);
        })().then(() => { shadowAnswers = true; }, () => undefined);
        let inTime = false;
        if (shadowAnswers) {
          let callOff = () => undefined as void;
          inTime = await Promise.race([asked.then(() => true),
            new Promise<boolean>((resolve) => { callOff = afterPdfRecoveryRunningTime(SHADOW_WAITED_MS, () => resolve(false)); })]);
          callOff();
          if (!inTime) shadowAnswers = false;
        }
        if (!inTime) answersLate = asked;
        wanted = false;
      }
      const taken = await takeByMirror(scope, scopeKey, mayTake);
      if (!taken && answersLate && takenLate) {
        void answersLate.then(() => takeByMirror(scope, scopeKey, mayTake))
          .then((late) => { if (late) takenLate(); }, () => undefined);
      }
      return taken;
    },
    /** Removes the queues nobody came back for (ORPHAN_KEPT_MS). */
    async sweep(): Promise<void> {
      if (!durable || !locks) return;
      if (shadow) await reconcile();
      for (const { key: claimKey, claim } of claimsUnder(CLAIM_PREFIX)) {
        if (claim.owner === instance || now() - claim.at < ORPHAN_KEPT_MS) continue;
        if (claim.owner !== null && !(await ownerGone(claim.owner))) continue;
        const name = claimKey.slice(CLAIM_PREFIX.length);
        await locks.request(TAKE_LOCK_PREFIX + name, { ifAvailable: true }, async (lock) => {
          const again = lock ? readClaim(claimKey) : null;
          if (!again || again.owner !== claim.owner || again.at !== claim.at) return;
          durable.removeItem(DATA_PREFIX + name);
          durable.removeItem(claimKey);
          dropMirroredWith(name);
          stampGone(name);
        }).catch(() => undefined);
      }
    },
    /** The document is going (pagehide): its queues are free for the reload
     * of this tab to take at once, or for another tab if none follows. */
    release(): void {
      if (!durable) return;
      try {
        for (const { key, claim } of claimsUnder(CLAIM_PREFIX)) if (claim.owner === instance) writeClaim(key, null);
      } catch { /* left to the tab lock, which goes with the document */ }
    },
  };
}

type PdfWorkingRecoveryTab = ReturnType<typeof createPdfWorkingRecoveryTab>;
let thisTab: PdfWorkingRecoveryTab | null = null;
function tabOfThisDocument(): PdfWorkingRecoveryTab {
  if (thisTab) return thisTab;
  let durable: Storage | null = null;
  let locks: TabLocks | null = null;
  let shadow: PdfWorkingRecoveryShadow | null = null;
  try {
    durable = window.localStorage;
    locks = typeof navigator !== "undefined" && typeof navigator.locks?.request === "function" ? navigator.locks : null;
    shadow = durable && locks ? createIndexedDbPdfWorkingRecoveryShadow() : null;
  } catch { /* storage blocked: the tab's record alone */ }
  const tab = createPdfWorkingRecoveryTab({ session: window.sessionStorage, durable, locks, shadow });
  if (durable && locks) {
    window.addEventListener("pagehide", (event) => { if (!event.persisted) tab.release(); });
    void tab.sweep().catch(() => undefined);
  }
  thisTab = tab;
  return tab;
}

export function createPdfWorkingRecoveryStore(input: {
  pageId: string;
  dataItemId: string;
  storage?: Storage;
}) {
  return (input.storage ? createPdfWorkingRecoveryTab({ session: input.storage }) : tabOfThisDocument()).store(input);
}

/** A queue of unsaved edits to this page that another tab of this account
 * left behind (it crashed, was closed or discarded) becomes this tab's, to
 * be restored like one it held across a reload. True when one was taken. */
export function takeOrphanedPdfWorkingRecovery(input: {
  pageId: string;
  dataItemId: string;
  /** Told the names of the records that come with the queue (`keepWith`). */
  mayTake?: (keptWith?: readonly string[]) => boolean;
  /** One was taken after all, after this answered false: the browser's own
   * store of the queues answered late (`mayTake` was asked again then). */
  takenLate?: () => void;
}): Promise<boolean> {
  if (typeof window === "undefined") return Promise.resolve(false);
  return tabOfThisDocument().take(input, input.mayTake, input.takenLate);
}

/** The records a window keeps in the tab beside the page's queue (by name:
 * tab key), mirrored with it so that a queue taken over after its tab went
 * comes with them. Called as the window opens the page, before it asks for
 * a queue left behind, and after each record is written or removed. */
export function keepWithPdfWorkingRecovery(input: {
  pageId: string;
  dataItemId: string;
  keptWith: Readonly<Record<string, string>>;
}): void {
  if (typeof window === "undefined") return;
  try { tabOfThisDocument().keepWith(input, input.keptWith); } catch { /* the tab's records stand alone */ }
}
