/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { advancePdfWorkingRecovery, createPdfNativeWorkingSession, PdfWorkingPlannerUnavailableError, restorePdfNativeWorkingSession,
  type PdfWorkingRecovery, type PdfWorkingRecoveryLostEdit,
  type PdfWorkingSessionOptions, type PdfWorkingSaveAcknowledgement } from "./native-working-session";
import { createPdfWorkingRecoveryStore } from "./working-recovery-store";
import type { PdfEngineSourceIdentity } from "./engine/contracts";

/** One window's local document owner. Stage resolves when native bytes and
 * recovery are local, not when the ordered publication queue finishes. The
 * caller's sink must return a verified receipt after shared hold retirement. */
/** Per page: how often in a row its kept queue could not be planned again
 * because the planner was out of reach (see createPdfWorkingDocument). */
const plannerOutOfReach = new Map<string, number>();
const PLANNER_OUT_OF_REACH_TRIES = 3;

export async function createPdfWorkingDocument(options: Omit<PdfWorkingSessionOptions,"writeRecovery"|"onChange"> & {
  dataItemId: string;
  recoveryStore?: ReturnType<typeof createPdfWorkingRecoveryStore>;
  /** An edit not written yet taken back (Undo) or put back (Redo): its row
   * saved as what it does (pdf_pending_edits) follows it. */
  onHeldUndone?: (operationId: string) => void;
  onHeldRedone?: (operationId: string) => void;
  



  onRecoveryLetGo?: (lost: Readonly<{ operations: number; reason: string;
    /** Each edit let go: a typed line's words are the window's to keep as that line's draft. */
    edits: readonly PdfWorkingRecoveryLostEdit[] }>) => void;
  /** Whether `source` is the version the server holds for the page now (a
   * queue kept for another version is planned again only on that one). */
  isCurrentSource?: () => Promise<boolean>;
}) {
  const store = options.recoveryStore ?? createPdfWorkingRecoveryStore(options);
  let stored = store.read();
  // The queue's first edits reached the server after the record was kept (a
  // save on its way when the page went, sent again by the window as it came
  // back): the page now stands on one of the queue's own outputs, and the
  // edits after it come back on exactly those bytes. The window's restore
  // does this when it opens the copy; a copy opened first by anything else
  // (the first edit after the reload, Save) read the record as another
  // version's and let the rest of the queue go -- two of five moves, with
  
  if (stored?.operations.length && stored.savedSource.sourceSha256 !== options.source.sourceSha256) {
    const advanced = advancePdfWorkingRecovery(stored, options.source);
    if (advanced) {
      store.write(advanced);
      stored = advanced;
    }
  }
  const listeners = new Set<() => void>();
  let closed = false;
  let publishing: Promise<void> | null = null;
  let session: Awaited<ReturnType<typeof createPdfNativeWorkingSession>> | null = null;
  const notify = () => {
    for (const listener of listeners) {
      try { listener(); } catch { /* Display observers cannot reverse durable staging. */ }
    }
  };
  const input = {...options,writeRecovery:store.write,onChange:notify};
  // A confirmed empty queue has nothing to replay. A later note/image save
  // may legitimately advance the same document's source in the meantime.
  // Nonempty recovery still requires its exact pinned base and reconciliation.
  // A queue kept for another saved version of the page (the page restored
  // or saved on since) cannot be planned again on this one: it is let go
  // and the copy opens fresh, instead of the copy not opening at all and
  // every later save of the page refused with "recovery identity is
  
  // The words typed are kept apart (the editor's retained draft).
  const plannerKey = `${options.pageId}:${options.dataItemId}`;
  if (stored?.operations.length) {
    const kept = stored;
    // A queue kept for another version of the page's file (another window
    // saved the page, or it was put back, while this tab's edits waited
    // across a reload) comes back on the version the page now has: each edit
    // is planned again there, as the queue of an open window is when its save
    // meets another window's. An edit that version does not take goes alone
    // and is said, what was typed kept as its line's draft; the rest is
    // unsaved again, as it was. Before, the whole queue was let go (and, until
    
    const otherVersion = kept.savedSource.sourceSha256 !== options.source.sourceSha256 &&
      kept.savedSource.resourceId === options.source.resourceId;
    // (Only on the page's current version: an older one this window happened
    // to open on would turn away edits the current one takes.)
    if (otherVersion && options.isCurrentSource && !(await options.isCurrentSource().catch(() => false))) {
      throw new PdfWorkingPlannerUnavailableError("The page's current version is not known yet.");
    }
    const recovery: PdfWorkingRecovery = otherVersion ? Object.freeze({ ...kept, savedSource: options.source,
      standsOn: Object.freeze({ held: kept.operations[0]!.baseSha256, sameContent: false }) }) : kept;
    const lost: PdfWorkingRecoveryLostEdit[] = [];
    const restore = (tolerate: "verdict" | "any") => {
      lost.length = 0;
      return restorePdfNativeWorkingSession(input, recovery,
        otherVersion ? { tolerate, lost: (edit) => { lost.push(edit); } } : undefined);
    };
    let failure: unknown = null;
    try { session = await restore("verdict"); }
    catch (error) { failure = error; session = null; }
    if (failure !== null) {
      const outOfReach = failure instanceof PdfWorkingPlannerUnavailableError;
      // The planner was out of reach, or (on another version) an edit failed
      // for a reason that is not the planner's verdict: nothing says the
      // queue cannot come back. Kept, and the copy does not open this time
      // (whoever opens it next plans the queue again). The second does not
      // hold the page for good: the third time, that edit is let go, said.
      // (A planner out of reach never lets the queue go, however often: the
      // file is large or the machine busy, and three timeouts in six seconds
      // threw away up to thirty-one edits. The queue waits; nothing can be
      // edited without the planner anyway.)
      if (outOfReach) throw failure;
      const tries = (plannerOutOfReach.get(plannerKey) ?? 0) + 1;
      if (otherVersion && tries < PLANNER_OUT_OF_REACH_TRIES) {
        plannerOutOfReach.set(plannerKey, tries);
        throw new PdfWorkingPlannerUnavailableError(failure instanceof Error ? failure.message : "The kept edits could not be planned again yet.");
      }
      if (otherVersion && !outOfReach) {
        try { session = await restore("any"); failure = null; }
        catch (error) { failure = error; session = null; }
      }
    }
    plannerOutOfReach.delete(plannerKey);
    const textOf = (operation: PdfWorkingRecovery["operations"][number]) => operation.kind === "engine" ? {}
      : { text: Object.freeze({ selection: operation.selection, replacementText: operation.replacementText }) };
    if (failure !== null || !session) {
      const reason = failure instanceof Error ? failure.message.slice(0, 160) : "unknown";
      if (process.env.NODE_ENV !== "production") console.warn("[pdf-working] stale recovery let go", reason);
      store.discardUnsaved();
      session = await createPdfNativeWorkingSession(input);
      try {
        options.onRecoveryLetGo?.(Object.freeze({ operations: kept.operations.length, reason,
          edits: Object.freeze(kept.operations.map((operation) => Object.freeze({ id: operation.id, reason, ...textOf(operation) }))) }));
      } catch { /* the notice is the window's; the copy is open either way */ }
    } else if (lost.length) {
      if (process.env.NODE_ENV !== "production") {
        console.warn("[pdf-working] recovered on another version; let go " + JSON.stringify({ lost: lost.length,
          back: session.snapshot().pendingOperationIds.length, why: lost[0]!.reason.slice(0, 100) }));
      }
      // An edit the version already holds as the edit leaves it (a row the
      // page took with a put-back, a move another window made) is not one
      // that "could not be brought back": it is there. Only the others are said.
      const missing = lost.filter((edit) => edit.alreadyThere !== true);
      try {
        if (missing.length) {
          options.onRecoveryLetGo?.(Object.freeze({ operations: missing.length, reason: missing[0]!.reason, edits: Object.freeze([...missing]) }));
        }
      } catch { /* the notice is the window's; the copy is open either way */ }
    }
  } else session = await createPdfNativeWorkingSession(input);
  const active = session;
  type Snapshot = ReturnType<typeof active.snapshot>;
  let snapshot: Snapshot = active.snapshot();
  const refresh = () => { snapshot = active.snapshot(); };
  // getSnapshot must be referentially stable for useSyncExternalStore. Update
  // it before notifying views; a saved baseline never changes local display.
  listeners.add(refresh);
  const assertOpen = () => { if (closed) throw new Error("PDF working document is closed."); };
  const publish = () => {
    assertOpen();
    if (publishing) return publishing;
    const pending = active.flush();
    publishing = pending.finally(() => {
      publishing = null;refresh();notify();
      if (!closed && !snapshot.failure && snapshot.pendingOperationIds.length) {
        queueMicrotask(() => { if (!closed) void publish().catch(() => undefined); });
      }
    });
    // Background failures are retained by the session and surfaced through
    // its snapshot. Always handle them, even when the UI does not await save.
    void publishing.catch(() => undefined);
    return publishing;
  };
  return {
    canWrite: () => !closed && active.canWrite(),
    getSnapshot: () => snapshot,
    /** A peer notification may retire an idle local display only at the exact
     * source already acknowledged by this owner. A receipt for a newer peer
     * source cannot stand in for the missing acknowledgement of our draft. */
    isFullySavedAt(source: Snapshot["savedSource"]): boolean {
      return !closed && active.canWrite() && !snapshot.disposed && !snapshot.saving && !snapshot.failure &&
        snapshot.pendingOperationIds.length === 0 && snapshot.savedRevision === snapshot.workingRevision &&
        snapshot.workingSha256 === snapshot.savedSource.sourceSha256 &&
        source.resourceId === snapshot.savedSource.resourceId &&
        source.resourceVersionId === snapshot.savedSource.resourceVersionId &&
        source.sourceSha256 === snapshot.savedSource.sourceSha256;
    },
    subscribe(listener: () => void) { assertOpen();listeners.add(listener);return () => {listeners.delete(listener);}; },
    displaySource() { assertOpen();return active.displaySource(options.dataItemId); },
    copyBytes() { assertOpen();return active.copyWorkingBytes(); },
    copyPendingOutput: (operationId: string) => closed ? null : active.copyPendingOutput(operationId),
    pendingOperationOrigin: (operationId: string) => closed ? null : active.pendingOperationOrigin(operationId),
    async stageText(input: Parameters<typeof active.prepareText>[0], publication: { start?: boolean } = {}) {
      assertOpen();
      const operation = await active.prepareText(input);
      try { assertOpen();active.stage(operation); }
      catch (error) { active.discardPrepared(operation);throw error; }
      // Recovered or failed queues need explicit reconciliation before the
      // next send. They may still retain later valid local edits durably.
      if (publication.start !== false && !active.snapshot().failure) void publish().catch(() => undefined);
      return operation;
    },
    /** Stage an edit the engine planned on the working bytes (a move, a
     * box move, a format): shown at once, published in order. */
    stageEngine(input: Parameters<typeof active.stageEngine>[0], publication: { start?: boolean } = {}) {
      assertOpen();
      const operation = active.stageEngine(input);
      if (publication.start !== false && !active.snapshot().failure) void publish().catch(() => undefined);
      return operation;
    },
    flush: publish,
    /** Undo / Redo of an edit this window holds that is not saved yet: the
     * working copy moves, nothing is sent (the edit being written now is
     * not one of them). Null when there is none to move. */
    undoHeld(): string | null {
      assertOpen();
      const id = active.unstageLast();
      if (id) { try { options.onHeldUndone?.(id); } catch { /* the row follows at best */ } }
      return id;
    },
    /** The saved version's own bytes after the server rewrote its chain
     * (see the session's adoptSavedBytes). */
    adoptSavedBytes: (bytes: Uint8Array): Promise<boolean> => closed ? Promise.resolve(false) : active.adoptSavedBytes(bytes),
    /** A saved version this copy did not make, stood on in place (see the
     * session's rebaseSaved). */
    rebaseSaved: (source: PdfEngineSourceIdentity, bytes: Uint8Array,
      options?: Readonly<{ keepPending?: boolean }>): Promise<boolean> =>
      closed ? Promise.resolve(false) : active.rebaseSaved(source, bytes, options),
    /** A failed queue that another window's save overtook (see the
     * session's namePeerVersion). */
    namePeerVersion: (): boolean => !closed && active.namePeerVersion(),
    /** The queue's oldest edits, which the saved version cannot take, go
     * alone (see the session's letGoLeading); what is queued behind them is
     * sent on. */
    letGoLeading(operationIds: readonly string[]): boolean {
      if (closed || !active.letGoLeading(operationIds)) return false;
      // (The send that named them goes on with the rest, as after a reconcile.)
      if (active.snapshot().pendingOperationIds.length) void publish().catch(() => undefined);
      return true;
    },
    redoHeld(): string | null {
      assertOpen();
      const id = active.restageLast();
      if (id) { try { options.onHeldRedone?.(id); } catch { /* the row follows at best */ } }
      return id;
    },
    reconcile(ack: PdfWorkingSaveAcknowledgement) {
      assertOpen();active.reconcile(ack);
      if (active.snapshot().pendingOperationIds.length) void publish().catch(() => undefined);
    },
    reconcileChain(acks: readonly PdfWorkingSaveAcknowledgement[]) {
      assertOpen();active.reconcileChain(acks);
      if (active.snapshot().pendingOperationIds.length) void publish().catch(() => undefined);
    },
    confirmUnchangedBaseline: (...args: Parameters<typeof active.confirmUnchangedBaseline>) => {
      assertOpen();active.confirmUnchangedBaseline(...args);
      // No automatic retry here. A deliberate flush owns that decision.
    },
    /** Close the window without saving what it held (the reader's 不保存):
     * the local queue and its recovery copy are dropped. */
    discardUnsaved() {
      if (closed) return;
      closed=true; active.dispose(); listeners.clear();
      try { store.discardUnsaved(); } catch { /* the tab's storage is gone with it */ }
    },
    dispose() {
      if (closed) return;
      closed=true;
      // Do not clear durable recovery or fabricate success on window close.
      // If HTTP is in flight, its own shared hold keeps reconciliation custody.
      active.dispose();listeners.clear();
    },
  };
}
