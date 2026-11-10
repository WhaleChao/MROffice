/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createPdfNativeWorkingSession, restorePdfNativeWorkingSession,
  type PdfWorkingSessionOptions, type PdfWorkingSaveAcknowledgement } from "./native-working-session";
import { createPdfWorkingRecoveryStore } from "./working-recovery-store";
import type { PdfEngineSourceIdentity } from "./engine/contracts";

/** One window's local document owner. Stage resolves when native bytes and
 * recovery are local, not when the ordered publication queue finishes. The
 * caller's sink must return a verified receipt after shared hold retirement. */
export async function createPdfWorkingDocument(options: Omit<PdfWorkingSessionOptions,"writeRecovery"|"onChange"> & {
  dataItemId: string;
  recoveryStore?: ReturnType<typeof createPdfWorkingRecoveryStore>;
  /** An edit not written yet taken back (Undo) or put back (Redo): its row
   * saved as what it does (pdf_pending_edits) follows it. */
  onHeldUndone?: (operationId: string) => void;
  onHeldRedone?: (operationId: string) => void;
}) {
  const store = options.recoveryStore ?? createPdfWorkingRecoveryStore(options);
  const stored = store.read();
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
  if (stored?.operations.length) {
    try { session = await restorePdfNativeWorkingSession(input, stored); }
    catch (error) {
      if (process.env.NODE_ENV !== "production") {
        console.warn("[pdf-working] stale recovery let go", error instanceof Error ? error.message : error);
      }
      store.discardUnsaved();
      session = await createPdfNativeWorkingSession(input);
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
