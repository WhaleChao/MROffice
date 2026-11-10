/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { jsonEquals, type JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { pageStoredJsonProjection } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import { isLosslessDurableNativeSaveReceipt, readNativeSaveReceiptSnapshot } from "@mrnote/core/web/lib/native-save-receipt.ts";
import { bindMraiNativePreparedTextRecord, type PdfNativeRecordCheckpoint } from "./engine/adapters/mrai-native-record";
import type { PdfEngineSourceIdentity } from "./engine/contracts";
import type { PdfSharedStateRecord } from "./engine/shared-state-bridge";
import type { PdfWorkingDisplaySource } from "./working-display-source";
import { readPdfCompiledTextPreview, type PdfCompiledTextPreview } from "./native-text-compiled-preview";
import type { PdfNativeListTextReplacementPlan, PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";
import type { PdfNativePreparedContentStreamTransaction } from "./engine/native/content-stream-transaction";
import { MAX_MRAI_NATIVE_SOURCE_BYTES, sha256PdfBytes } from "./engine/native/transaction-error";
import { compilePdfNativeTextOffThread } from "./native-text-edit-preparation";
import { notePdfNativeHeadUnheld, type PdfNativeHeadBase, registerPdfNativeWorkingHead } from "./remote/pdf-native-sources";
import { readPdfNativeConsolidatedOutputSha256 } from "./engine/native/source-consolidation";

type Prepared = PdfNativePreparedContentStreamTransaction;
type Plan = PdfNativeListTextReplacementPlan;
/** The head `next` is an incremental update of, when it is one: `previous`
 * with bytes appended (its first upload is then only those bytes). */
function appendedTo(previous: Uint8Array, previousSha256: string, next: Uint8Array): PdfNativeHeadBase | null {
  if (next.length <= previous.length) return null;
  let index = 0;
  // Four bytes a step where both are aligned (a working head is its own
  // buffer): a large file is compared once per staged edit.
  if (previous.byteOffset % 4 === 0 && next.byteOffset % 4 === 0) {
    const words = previous.length >>> 2;
    const left = new Uint32Array(previous.buffer, previous.byteOffset, words);
    const right = new Uint32Array(next.buffer, next.byteOffset, words);
    for (let word = 0; word < words; word += 1) if (left[word] !== right[word]) return null;
    index = words << 2;
  }
  for (; index < previous.length; index += 1) if (next[index] !== previous[index]) return null;
  return { sha256: previousSha256, length: previous.length };
}

export interface PdfLocalTextOperation {
  readonly id: string;
  readonly pageId: string;
  readonly baseWorkingRevision: number;
  readonly baseSha256: string;
  readonly outputSha256: string;
}
interface OwnedTextOperation {
  kind: "text";
  prepared: Prepared;
  plan: Plan;
  authority: string;
  checkpoint?: PdfNativeRecordCheckpoint;
  /** The bytes-publish record last sent for it (publishBytes). */
  sent?: PdfSharedStateRecord;
}
/** Any other edit the engine plans (a move, a box move, a format...): its
 * candidate bytes show at once; its record is planned again against the
 * saved version when the queue publishes it, and must land on the same
 * output (the engine plans deterministically from the same bytes). */
export interface PdfWorkingEngineRecipe {
  readonly kind: "engine";
  /** The engine command, its source fields bound when it is planned. */
  readonly command: JsonValue;
  /** The engine's plan options (authorities it reads). */
  readonly options?: JsonValue;
  /** What the reader sees it as, for history and notices. */
  readonly label?: string;
}
export interface PdfWorkingEngineOperationInput {
  readonly id: string;
  readonly candidateBytes: Uint8Array;
  readonly outputSha256: string;
  readonly recipe: PdfWorkingEngineRecipe;
  /** The record for the saved version the queue publishes it onto. With
   * `rebased`, that version is another file with the same content as the
   * bytes the edit was planned on (the server rewrote the chain as one
   * revision): the edit is planned again on it and its own output taken. */
  readonly bindRecord: (source: PdfEngineSourceIdentity, workingRevision: number,
    rebased?: boolean) => Promise<PdfSharedStateRecord>;
}
interface OwnedEngineOperation {
  kind: "engine";
  candidateBytes: Uint8Array;
  authority: string;
  recipe: PdfWorkingEngineRecipe;
  bindRecord: PdfWorkingEngineOperationInput["bindRecord"];
  /** The record last sent for it (reconciliation compares the server's). */
  sent?: PdfSharedStateRecord;
}
type OwnedOperation = OwnedTextOperation | OwnedEngineOperation;
const candidateOf = (data: OwnedOperation) => data.kind === "text" ? data.prepared.candidateBytes : data.candidateBytes;
export interface PdfWorkingSaveAcknowledgement {
  /** Exact submitted operation and actual server page, never a merged draft projection. */
  readonly record: PdfSharedStateRecord;
  readonly page: { id: string; content_json: Record<string, JsonValue>; editable_page_patchback?: unknown };
}
export interface PdfWorkingRecovery {
  readonly schemaVersion: "mrai.pdf-working-recovery.v1";
  readonly pageId: string;
  readonly savedSource: PdfEngineSourceIdentity;
  readonly savedRevision: number;
  readonly operations: readonly (Readonly<{
    id: string; baseWorkingRevision: number; baseSha256: string; outputSha256: string;
    selection: PdfNativeListTextSelectionEvidence; replacementText: string; kind?: undefined;
  }> | Readonly<{
    id: string; baseWorkingRevision: number; baseSha256: string; outputSha256: string;
    kind: "engine"; recipe: PdfWorkingEngineRecipe;
  }>)[];
}
export interface PdfWorkingSessionOptions {
  readonly pageId: string;
  readonly source: PdfEngineSourceIdentity;
  readonly sourceBytes: Uint8Array;
  readonly originDocumentRevision?: string | number;
  readonly readAuthority: () => { writable: boolean; key: string };
  /** Reading the retained draft is independent of a write-grant generation.
   * Omitting this retains the stricter legacy scope behavior. */
  readonly readDisplayAuthority?: () => { readable: boolean; key: string };
  readonly persist: (input: { record: PdfSharedStateRecord; baseSource: PdfEngineSourceIdentity }) => Promise<PdfWorkingSaveAcknowledgement>;
  readonly onChange?: () => void;
  /** Must finish storing recovery before exposing a new working revision. */
  readonly writeRecovery: (recovery: PdfWorkingRecovery) => void;
  /** Publishes a queued edit as the revision it appends to the saved
   * version (engine unification, stage 2): makes sure the server holds those
   * bytes and answers the record naming them, or null when this edit goes
   * as its planned record instead. Absent, every edit goes as its record. */
  readonly publishBytes?: (input: { operationId: string; source: PdfEngineSourceIdentity;
    sourceBytes: Uint8Array; output: Uint8Array; outputSha256: string; baseWorkingRevision: number;
    label?: string }) => Promise<PdfSharedStateRecord | null>;
  /** Plans a recovered engine edit again on the bytes it was planned on
   * (after a reload). Without it a recovered engine edit is not restored. */
  readonly prepareEngineRecipe?: (input: { recipe: PdfWorkingEngineRecipe; id: string; sourceBytes: Uint8Array;
    sourceSha256: string; savedSource: PdfEngineSourceIdentity }) => Promise<PdfWorkingEngineOperationInput>;
  /** The bytes of a saved version (checked against its digest), or null.
   * With it, an edit queued behind a save the server rewrote (one revision,
   * the same content as another file) is planned again on the rewritten
   * bytes and published as its own bytes, instead of being refused. */
  readonly loadSavedBytes?: (source: PdfEngineSourceIdentity) => Promise<Uint8Array | null>;
  /** Whether a queued edit is planned again on the saved version before it
   * is published, though nothing moved under it (a plan the server keeps
   * only for a while, held long with AutoSave off). */
  readonly replanBeforePublish?: (operationId: string) => boolean;
  /** Whether a recovered engine edit may come back as other bytes than it
   * had (a plan the server makes again, never byte for byte the same):
   * the edits after it are then checked against what it came back as. */
  readonly recipeReplannable?: (recipe: PdfWorkingEngineRecipe) => boolean;
}

/** Verified local edits and durable publication have distinct revisions.
 * The queue never reports an unacknowledged working source as a server version. */
async function createWorkingSession(options: PdfWorkingSessionOptions, initialRevision = 0) {
  const authority = options.readAuthority();
  const displayAuthority = options.readDisplayAuthority?.() ?? { readable: true, key: authority.key };
  if (!authority.writable || !authority.key || !options.pageId ||
    !displayAuthority.readable || !displayAuthority.key || !options.source.resourceId || !options.source.resourceVersionId ||
    !options.sourceBytes.length || options.sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES)
    throw new Error("PDF working session authority is unavailable.");
  let workingBytes = Uint8Array.from(options.sourceBytes);
  if (await sha256PdfBytes(workingBytes) !== options.source.sourceSha256) {
    workingBytes.fill(0); throw new Error("PDF working source hash mismatch.");
  }
  let savedSource = Object.freeze({ ...options.source });
  // The planners run on the server: every head this session holds is
  // registered with the page and the stored version it descends from.
  const registerHead = (sha256: string) => registerPdfNativeWorkingHead(sha256, {
    pageId: options.pageId, resourceId: savedSource.resourceId, resourceVersionId: savedSource.resourceVersionId,
  });
  registerHead(savedSource.sourceSha256);
  let workingSha256 = savedSource.sourceSha256, workingRevision = initialRevision, savedRevision = initialRevision;
  let disposed = false, preparing = false;
  // The saved version's bytes: what the working copy goes back to when every
  // queued edit is undone before it was saved (AutoSave off holds the queue).
  let savedBytes = Uint8Array.from(workingBytes);
  // Display keys name (session, revision); an undo goes back to a revision
  // number with other bytes after it, so it starts a new display session.
  let displaySessionId = crypto.randomUUID();
  let displaySource: PdfWorkingDisplaySource | null = null;
  let failure: string | null = null;
  let flight: Promise<void> | null = null;
  const pending: PdfLocalTextOperation[] = [];
  /** The bytes this copy holds that the saved version has the content of,
   * as another file: the server rewrote its revision chain as one revision
   * when it saved them (near the parser's ceiling, every ~24 saves). The
   * edits queued after them are planned again on the saved version. */
  let rewritten: { held: string; saved: string } | null = null;
  const onSavedBase = (operation: PdfLocalTextOperation) =>
    operation.baseSha256 === savedSource.sourceSha256 ||
    (rewritten !== null && rewritten.held === operation.baseSha256 && rewritten.saved === savedSource.sourceSha256);
  /** Queued edits undone before they were saved, newest last (Redo). */
  const undone: PdfLocalTextOperation[] = [];
  const owned = new Map<PdfLocalTextOperation, OwnedOperation>();
  const ids = new Set<string>();
  const checkAuthority = () => {
    const current = options.readAuthority();
    if (disposed || !current.writable || current.key !== authority.key)
      throw new Error("PDF working session authority changed.");
  };
  const displayIsCurrent = () => {
    const current = options.readDisplayAuthority?.() ?? { readable: true, key: options.readAuthority().key };
    return !disposed && current.readable && current.key === displayAuthority.key;
  };
  const clearOperation = (operation: PdfLocalTextOperation) => {
    const data = owned.get(operation);
    if (data) candidateOf(data).fill(0);
    owned.delete(operation);
  };
  const recovery = (operations: readonly PdfLocalTextOperation[] = pending,
    baseline: PdfEngineSourceIdentity = savedSource, revision = savedRevision): PdfWorkingRecovery => Object.freeze({
      schemaVersion: "mrai.pdf-working-recovery.v1", pageId: options.pageId, savedSource: baseline, savedRevision: revision,
      operations: Object.freeze(operations.map(operation => {
        const data = owned.get(operation)!;
        return data.kind === "engine"
          ? Object.freeze({ ...operation, kind: "engine" as const, recipe: data.recipe })
          : Object.freeze({ ...operation, selection: data.plan.evidence.selection!, replacementText: data.plan.evidence.replacementDisplayText });
      })),
    });
  const acknowledge = (operation: PdfLocalTextOperation, record: PdfSharedStateRecord, ack: PdfWorkingSaveAcknowledgement,
    /** What the record writes (the edit's own output, or a rebased plan's). */
    recordOutput: string = operation.outputSha256,
    /** The bytes it wrote, when planned again on a rewritten save. */
    recordBytes?: Uint8Array) => {
    const data = owned.get(operation);
    const receipt = readNativeSaveReceiptSnapshot(ack.page);
    const journal = ack.page.content_json.pdfNativeTransactions;
    // Near the parser's ceiling the server rewrites the saved output as one
    // revision; its journal then ends in the rewrite naming our output.
    const consolidated = readPdfNativeConsolidatedOutputSha256(ack.page, recordOutput);
    const savedOutput = consolidated ?? recordOutput;
    if (!data || pending[0] !== operation || operation.baseWorkingRevision !== savedRevision ||
      !onSavedBase(operation) || ack.page.id !== options.pageId ||
      !jsonEquals(ack.record, record) || !Array.isArray(journal) ||
      (!consolidated && (
        journal.filter(entry => entry && typeof entry === "object" && !Array.isArray(entry) && entry.id === operation.id).length !== 1 ||
        // The stored record carries U+FFFD where the queued one carried U+0000.
        !jsonEquals(pageStoredJsonProjection(journal.at(-1)!), pageStoredJsonProjection(record)))) ||
      !isLosslessDurableNativeSaveReceipt(receipt, { expectedAdapter: "pdf", expectedResourceId: savedSource.resourceId,
        expectedOutputSha256: savedOutput, previousResourceVersionId: savedSource.resourceVersionId }) ||
      typeof receipt?.resourceVersionId !== "string") throw new Error("PDF save acknowledgement does not match the queued operation.");
    const nextSource = Object.freeze({ resourceId: savedSource.resourceId, resourceVersionId: receipt.resourceVersionId,
      sourceSha256: savedOutput });
    // The copy keeps the bytes it shows; the saved version has their content.
    rewritten = savedOutput !== operation.outputSha256 ? { held: operation.outputSha256, saved: savedOutput } : null;
    const nextRevision = operation.baseWorkingRevision + 1;
    options.writeRecovery(recovery(pending.slice(1), nextSource, nextRevision));
    savedBytes.fill(0); savedBytes = Uint8Array.from(recordBytes ?? candidateOf(data));
    savedSource = nextSource;
    savedRevision = nextRevision;
    registerHead(savedSource.sourceSha256);
    if (workingSha256 !== savedSource.sourceSha256) registerHead(workingSha256);
    pending.shift(); clearOperation(operation);
    // workingBytes/revision and the UI draft deliberately remain untouched.
    failure = null; options.onChange?.();
  };
  /** An edit planned again on the saved version's own bytes (a rewrite of
   * the chain gave them a new digest, the same content): their candidate
   * extends what the server holds, so it publishes as bytes. The saved bytes
   * this session keeps become those bytes. */
  const replanOnSaved = async (operation: PdfLocalTextOperation, data: OwnedOperation):
    Promise<Readonly<{ candidate: Uint8Array; sha256: string }> | null> => {
    // The saved bytes this session holds already, when they are the saved
    // version's (after the first edit planned again); else loaded.
    const held = await sha256PdfBytes(savedBytes) === savedSource.sourceSha256 ? Uint8Array.from(savedBytes) : null;
    const saved = held ?? await options.loadSavedBytes!(savedSource);
    if (!saved) return null;
    try {
      if (await sha256PdfBytes(saved) !== savedSource.sourceSha256) return null;
      let candidate: Uint8Array, sha256: string;
      if (data.kind === "engine") {
        if (!options.prepareEngineRecipe) return null;
        const prepared = await options.prepareEngineRecipe({ recipe: data.recipe, id: operation.id,
          sourceBytes: Uint8Array.from(saved), sourceSha256: savedSource.sourceSha256, savedSource });
        candidate = prepared.candidateBytes; sha256 = prepared.outputSha256;
      } else {
        const selection = data.plan.evidence.selection;
        if (!selection) return null;
        const compiled = await compilePdfNativeTextOffThread({ sourceBytes: Uint8Array.from(saved),
          selection: { ...selection, sourceSha256: savedSource.sourceSha256 },
          replacementText: data.plan.evidence.replacementDisplayText, signal: new AbortController().signal,
          engineSource: { pageId: options.pageId, resourceId: savedSource.resourceId,
            resourceVersionId: savedSource.resourceVersionId } });
        candidate = compiled.prepared.candidateBytes; sha256 = compiled.prepared.candidateSha256;
      }
      if (!appendedTo(saved, savedSource.sourceSha256, candidate)) { candidate.fill(0); return null; }
      savedBytes.fill(0); savedBytes = Uint8Array.from(saved);
      return Object.freeze({ candidate, sha256 });
    } finally { saved.fill(0); }
  };
  const session = {
    canWrite() { try { checkAuthority();return true; } catch { return false; } },
    snapshot() {
      return Object.freeze({ pageId: options.pageId, workingRevision, workingSha256, savedRevision,
        savedSource, pendingOperationIds: Object.freeze(pending.map(item => item.id)),
        /** What the oldest queued edit writes: the server may hold it before
         * this copy confirmed it (a save that landed while the window stopped
         * waiting), and the copy reconciles against it. */
        firstPendingOutputSha256: pending[0]?.outputSha256 ?? null,
        undoneOperationIds: Object.freeze(undone.map(item => item.id)),
        saving: flight !== null, failure, disposed,
        /** The copy shows bytes whose content the saved version holds as
         * another file (the server rewrote the chain). */
        rewritten: rewritten !== null && workingSha256 !== savedSource.sourceSha256 });
    },
    copyWorkingBytes() { checkAuthority(); return Uint8Array.from(workingBytes); },
    /** The bytes a queued edit produces (what the server holds once it is
     * published): the acceptance of its receipt reads them, not a download. */
    copyPendingOutput(operationId: string): Uint8Array | null {
      if (disposed) return null;
      const operation = pending.find(item => item.id === operationId);
      const data = operation && owned.get(operation);
      return data ? Uint8Array.from(candidateOf(data)) : null;
    },
    





    pendingOperationOrigin(operationId: string): Readonly<{
      baseSha256: string; baseLength: number; outputSha256: string; outputLength: number;
      inFlight: boolean; kind: "text" | "engine"; label: string | null;
    }> | null {
      if (disposed) return null;
      const index = pending.findIndex(item => item.id === operationId);
      if (index < 0) return null;
      const operation = pending[index], data = owned.get(operation);
      if (!data) return null;
      const previous = index > 0 ? owned.get(pending[index - 1]) : null;
      const base = index > 0 ? (previous ? candidateOf(previous) : null) : savedBytes;
      if (!base) return null;
      return Object.freeze({
        baseSha256: operation.baseSha256, baseLength: base.length,
        outputSha256: operation.outputSha256, outputLength: candidateOf(data).length,
        inFlight: flight !== null && index === 0,
        kind: data.kind,
        label: data.kind === "engine" ? data.recipe.label ?? null : "text",
      });
    },
    displaySource(dataItemId: string): PdfWorkingDisplaySource {
      if (!displayIsCurrent()) throw new Error("PDF working display authority changed.");
      if (!dataItemId || (displaySource && displaySource.dataItemId !== dataItemId)) {
        throw new Error("PDF working display data item changed.");
      }
      if (displaySource?.workingRevision === workingRevision) return displaySource;
      const revision = workingRevision, sha256 = workingSha256;
      displaySource = Object.freeze({kind:"local-working-pdf", pageId:options.pageId, dataItemId,
        sessionId:displaySessionId, workingRevision:revision, sha256,
        displayKey:`mrai-pdf-working:${displaySessionId}:${revision}`,
        originDocumentRevision: options.originDocumentRevision,
        baseSha256: savedSource.sourceSha256,
        copyBytes: () => {
          // Display is a read. A write-permission loss retains the last valid
          // local bytes for recovery, while session/page disposal retires them.
          if (!displayIsCurrent() ||
            workingRevision !== revision || workingSha256 !== sha256) return null;
          return Uint8Array.from(workingBytes);
        }});
      return displaySource;
    },
    recovery,
    async prepareText(input: { id: string; selection: PdfNativeListTextSelectionEvidence; replacementText: string;
      compiledPreview?: PdfCompiledTextPreview | null }): Promise<PdfLocalTextOperation> {
      checkAuthority();
      if (preparing || !input.id || ids.has(input.id) || input.selection.sourceSha256 !== workingSha256)
        throw new Error("PDF local operation is stale or already being prepared.");
      preparing = true;
      const baseRevision = workingRevision, bytes = Uint8Array.from(workingBytes);
      let prepared: Prepared | null = null;
      try {
        // The preview's own compilation, or a fresh one from the engine: both
        // were planned, prepared and verified there against these bytes.
        const compiled = readPdfCompiledTextPreview(input.compiledPreview,input.selection,input.replacementText) ??
          await compilePdfNativeTextOffThread({ sourceBytes: bytes, selection: input.selection,
            replacementText: input.replacementText, signal: new AbortController().signal,
            engineSource: { pageId: options.pageId, resourceId: savedSource.resourceId,
              resourceVersionId: savedSource.resourceVersionId } });
        const plan = compiled.plan;
        prepared = compiled.prepared;
        if (prepared.request.identity.sourceSha256 !== workingSha256 ||
          prepared.candidateBytes.length < bytes.length) throw new Error("PDF working source changed during preparation.");
        checkAuthority();
        if (workingRevision !== baseRevision) throw new Error("PDF working source changed during preparation.");
        const operation = Object.freeze({ id: input.id, pageId: options.pageId, baseWorkingRevision: baseRevision,
          baseSha256: workingSha256, outputSha256: prepared.candidateSha256 });
        const retainedSize = [...owned.values()].reduce((total, entry) => total + candidateOf(entry).length, workingBytes.length);
        if (retainedSize + prepared.candidateBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES * 2)
          throw new Error("PDF pending operation memory budget is exhausted.");
        const checkpoint = baseRevision === savedRevision && Number.isSafeInteger(prepared.sourceXrefSections)
          ? Object.freeze({ source: savedSource, xrefSections: prepared.sourceXrefSections! }) : undefined;
        checkAuthority();
        if (workingRevision !== baseRevision) throw new Error("PDF working source changed during preparation.");
        owned.set(operation, { kind: "text", prepared, plan, authority: authority.key, checkpoint }); ids.add(input.id);
        prepared = null;
        return operation;
      } finally { bytes.fill(0); prepared?.candidateBytes.fill(0); preparing = false; }
    },
    stage(operation: PdfLocalTextOperation) {
      checkAuthority(); const data = owned.get(operation);
      if (!data || data.authority !== authority.key || pending.includes(operation) || operation.pageId !== options.pageId ||
        operation.baseWorkingRevision !== workingRevision || operation.baseSha256 !== workingSha256)
        throw new Error("PDF prepared operation no longer owns the working source.");
      // Keep only bounded file data; the native compiler retains the source/xref budgets.
      const retainedSize = [...owned.values()].reduce((total, entry) => total + candidateOf(entry).length, 0);
      if (retainedSize + workingBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES * 2)
        throw new Error("PDF pending edits must finish saving before this operation can be staged.");
      options.writeRecovery(recovery([...pending, operation]));
      const base = appendedTo(workingBytes, workingSha256, candidateOf(data));
      workingBytes.fill(0); workingBytes = Uint8Array.from(candidateOf(data));
      workingSha256 = operation.outputSha256; workingRevision++;
      registerHead(workingSha256);
      if (data.kind === "text") notePdfNativeHeadUnheld(workingSha256, base);
      pending.push(operation);
      // A new edit ends what Redo could bring back, as in WPS.
      undone.length = 0;
      for (const prior of owned.keys()) if (prior !== operation && !pending.includes(prior)) clearOperation(prior);
      options.onChange?.();
    },
    /** Undo the newest queued edit that is not saved yet: the working copy
     * goes back to the bytes before it and the edit waits for Redo. Nothing
     * is sent; local recovery holds the shorter queue. While the queue
     * publishes (AutoSave on) the edit being written now is not taken back
     * here: it is the oldest, and any later one can be. */
    unstageLast(): string | null {
      checkAuthority();
      if (preparing || !pending.length || (flight && pending.length < 2)) return null;
      const operation = pending[pending.length - 1];
      const previous = pending.length > 1 ? owned.get(pending[pending.length - 2]) : null;
      const before = pending.length > 1 ? previous && candidateOf(previous) : savedBytes;
      if (!before || !owned.has(operation)) return null;
      options.writeRecovery(recovery(pending.slice(0, -1)));
      pending.pop(); undone.push(operation);
      workingBytes.fill(0); workingBytes = Uint8Array.from(before);
      workingSha256 = operation.baseSha256; workingRevision = operation.baseWorkingRevision;
      displaySessionId = crypto.randomUUID(); displaySource = null;
      if (!pending.length) failure = null;
      options.onChange?.();
      return operation.id;
    },
    /** Redo the edit the last Undo took back, while nothing else changed
     * (it goes back at the end of the queue, after any being written). */
    restageLast(): string | null {
      checkAuthority();
      if (preparing || !undone.length) return null;
      const operation = undone[undone.length - 1], data = owned.get(operation);
      if (!data || operation.baseWorkingRevision !== workingRevision || operation.baseSha256 !== workingSha256) return null;
      options.writeRecovery(recovery([...pending, operation]));
      undone.pop(); pending.push(operation);
      const base = appendedTo(workingBytes, workingSha256, candidateOf(data));
      workingBytes.fill(0); workingBytes = Uint8Array.from(candidateOf(data));
      workingSha256 = operation.outputSha256; workingRevision++;
      registerHead(workingSha256);
      if (data.kind === "text") notePdfNativeHeadUnheld(workingSha256, base);
      displaySessionId = crypto.randomUUID(); displaySource = null;
      options.onChange?.();
      return operation.id;
    },
    






    async adoptSavedBytes(bytes: Uint8Array): Promise<boolean> {
      checkAuthority();
      const idle = () => !preparing && !flight && !pending.length && !failure && workingRevision === savedRevision;
      if (!idle() || await sha256PdfBytes(bytes) !== savedSource.sourceSha256) return false;
      checkAuthority();
      if (!idle()) return false;
      // A Redo of an edit taken back before it was saved was planned on the
      // old bytes: it goes (an edit typed now starts a new history anyway).
      for (const operation of undone.splice(0)) clearOperation(operation);
      workingBytes.fill(0); workingBytes = Uint8Array.from(bytes);
      savedBytes.fill(0); savedBytes = Uint8Array.from(bytes);
      workingSha256 = savedSource.sourceSha256; rewritten = null;
      registerHead(workingSha256);
      displaySessionId = crypto.randomUUID(); displaySource = null;
      options.writeRecovery(recovery([]));
      options.onChange?.();
      return true;
    },
    /** The page saved a version this copy did not make (a peer's write, a
     * save of another kind): an idle copy stands on it from here, as if
     * opened on it (engine unification, stage 6: the copy is not retired
     * and opened again). What was undone on the old version goes. Nothing
     * else of the session changes: the display starts a new session. */
    async rebaseSaved(source: PdfEngineSourceIdentity, bytes: Uint8Array,
      rebaseOptions?: Readonly<{ keepPending?: boolean }>): Promise<boolean> {
      // (The authority reads the copy's saved version against the page's:
      // not writable until the switch, checked after it.)
      if (disposed || options.readAuthority().key !== authority.key) return false;
      // `keepPending`: the queue met a peer's version (PDF_WORKING_PEER_VERSION_FAILURE)
      // and stays; the saved base moves under it and the flush plans each
      // queued edit again on the saved bytes, as behind a rewritten save.
      const ready = rebaseOptions?.keepPending
        ? () => !preparing && !flight && pending.length > 0 && failure === PDF_WORKING_PEER_VERSION_FAILURE
        : () => !preparing && !flight && !pending.length && !failure && workingRevision === savedRevision;
      if (!ready() || source.resourceId !== savedSource.resourceId || !source.resourceVersionId ||
        !/^[a-f0-9]{64}$/.test(source.sourceSha256) || source.sourceSha256 === savedSource.sourceSha256 ||
        !bytes.length || bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES ||
        await sha256PdfBytes(bytes) !== source.sourceSha256) return false;
      if (disposed || !ready()) return false;
      if (rebaseOptions?.keepPending) {
        savedBytes.fill(0); savedBytes = Uint8Array.from(bytes);
        savedSource = Object.freeze({ resourceId: source.resourceId, resourceVersionId: source.resourceVersionId,
          sourceSha256: source.sourceSha256 });
        // The queue's head was planned on the version before: the flush reads
        // it as behind a rewritten save and plans it again on this one.
        rewritten = { held: pending[0].baseSha256, saved: savedSource.sourceSha256 };
        registerHead(savedSource.sourceSha256);
        failure = null;
        options.writeRecovery(recovery(pending, savedSource, savedRevision));
        checkAuthority();
        options.onChange?.();
        return true;
      }
      for (const operation of undone.splice(0)) clearOperation(operation);
      workingBytes.fill(0); workingBytes = Uint8Array.from(bytes);
      savedBytes.fill(0); savedBytes = Uint8Array.from(bytes);
      savedSource = Object.freeze({ resourceId: source.resourceId, resourceVersionId: source.resourceVersionId,
        sourceSha256: source.sourceSha256 });
      workingSha256 = savedSource.sourceSha256; rewritten = null;
      registerHead(workingSha256);
      displaySessionId = crypto.randomUUID(); displaySource = null;
      options.writeRecovery(recovery([]));
      checkAuthority();
      options.onChange?.();
      return true;
    },
    /** Stage an engine-planned edit whose candidate was planned on the
     * current working bytes (see PdfWorkingEngineRecipe). */
    stageEngine(input: PdfWorkingEngineOperationInput): PdfLocalTextOperation {
      checkAuthority();
      if (preparing || !input.id || ids.has(input.id) || !input.candidateBytes.length ||
        input.candidateBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES || !/^[a-f0-9]{64}$/.test(input.outputSha256) ||
        input.outputSha256 === workingSha256)
        throw new Error("PDF engine operation is stale or already staged.");
      const operation = Object.freeze({ id: input.id, pageId: options.pageId, baseWorkingRevision: workingRevision,
        baseSha256: workingSha256, outputSha256: input.outputSha256 });
      owned.set(operation, { kind: "engine", candidateBytes: Uint8Array.from(input.candidateBytes), authority: authority.key,
        recipe: input.recipe, bindRecord: input.bindRecord });
      ids.add(input.id);
      try { session.stage(operation); }
      catch (error) { clearOperation(operation); ids.delete(input.id); throw error; }
      return operation;
    },
    discardPrepared(operation: PdfLocalTextOperation) {
      if (pending.includes(operation)) throw new Error("Queued PDF operations require history or reconciliation.");
      clearOperation(operation);
    },
    flush(): Promise<void> {
      if (flight) return flight;
      checkAuthority();
      if (failure) return Promise.reject(new Error("PDF save requires reconciliation before retry."));
      const run = async () => {
        while (pending.length) {
          checkAuthority();
          const operation = pending[0], data = owned.get(operation)!;
          if (!onSavedBase(operation)) throw new Error("PDF queued source conflicts with the saved baseline.");
          const rebased = operation.baseSha256 !== savedSource.sourceSha256;
          // Planned again: behind a rewritten save, or a plan held too long.
          const replan = rebased || options.replanBeforePublish?.(operation.id) === true;
          let record: PdfSharedStateRecord | null = null;
          // Behind a save the server rewrote: planned again on the rewritten
          // bytes and published as its own (a typed line was refused here,
          // and the copy stopped writing until a reload, 24 saves into any
          
          let replanned: Readonly<{ candidate: Uint8Array; sha256: string }> | null = null;
          if (options.publishBytes && options.loadSavedBytes && replan) {
            if (process.env.NODE_ENV !== "production") {
              console.debug("[pdf-working] replan start " + JSON.stringify({ id: operation.id.slice(0, 8), rebased,
                kind: data.kind, label: data.kind === "engine" ? data.recipe.label : "text" }));
            }
            replanned = await replanOnSaved(operation, data).catch((error) => {
              if (process.env.NODE_ENV !== "production") console.debug("[pdf-working] replan threw", error);
              return null;
            });
            if (process.env.NODE_ENV !== "production") {
              console.debug("[pdf-working] replan " + JSON.stringify({ id: operation.id.slice(0, 8), planned: Boolean(replanned) }));
            }
            checkAuthority();
            if (replanned) {
              record = await options.publishBytes({ operationId: operation.id, source: savedSource,
                sourceBytes: savedBytes, output: replanned.candidate, outputSha256: replanned.sha256,
                baseWorkingRevision: operation.baseWorkingRevision,
                label: data.kind === "engine" ? data.recipe.label : "text" });
              checkAuthority();
              if (record) data.sent = record;
              else { replanned.candidate.fill(0); replanned = null; }
            }
          }
          // The edit's own bytes, when they extend the saved version as the
          // server stores it: published as they are, nothing planned again.
          if (options.publishBytes && !rebased && !replanned &&
            appendedTo(savedBytes, savedSource.sourceSha256, candidateOf(data))) {
            record = await options.publishBytes({ operationId: operation.id, source: savedSource,
              sourceBytes: savedBytes, output: candidateOf(data), outputSha256: operation.outputSha256,
              baseWorkingRevision: operation.baseWorkingRevision,
              label: data.kind === "engine" ? data.recipe.label : "text" });
            checkAuthority();
            if (record) data.sent = record;
          }
          if (!record && process.env.NODE_ENV !== "production" && options.publishBytes) {
            // Stage 7: a record the server plans or replays itself is the
            
            try {
              console.warn("[pdf-working] legacy record " + JSON.stringify({
                kind: data.kind, label: data.kind === "engine" ? data.recipe.label : "text",
                why: !replan ? "not-appended" : rebased ? "rebased-replan-failed" : "replan-failed" }));
            } catch { /* diagnostics never throw */ }
          }
          if (record) {
            // Published as its bytes.
          } else if (data.kind === "engine") {
            record = await data.bindRecord(savedSource, operation.baseWorkingRevision, rebased);
            checkAuthority();
            data.sent = record;
          } else {
            // A typed line's bytes were prepared on the file before the
            // rewrite: it cannot be written on the rewritten one.
            if (rebased) throw new Error("PDF typed text queued behind a rewritten save must be typed again.");
            if (!Number.isSafeInteger(data.prepared.sourceXrefSections)) throw new Error("PDF queued source has no revision depth.");
            data.checkpoint = Object.freeze({ source: savedSource, xrefSections: data.prepared.sourceXrefSections! });
            checkAuthority();
            record = bindMraiNativePreparedTextRecord({ source: savedSource, transactionId: operation.id,
              workingRevision: operation.baseWorkingRevision, selectionRevision: data.plan.evidence.selection!.selectionRevision,
              prepared: data.prepared, fidelity: data.plan.evidence, sourceCheckpoint: data.checkpoint });
          }
          const ack = await options.persist({ record: record!, baseSource: savedSource });
          const derived = (record as { derivedCandidate?: { sha256?: unknown } }).derivedCandidate?.sha256;
          checkAuthority();
          try {
            acknowledge(operation, record!, ack, (rebased || replanned) && typeof derived === "string"
              ? derived : operation.outputSha256, replanned?.candidate);
          } finally { replanned?.candidate.fill(0); }
        }
      };
      flight = Promise.resolve().then(run).catch(error => {
        failure = error instanceof Error ? error.message : "PDF publication failed.";
        throw error;
      }).finally(() => { flight = null; options.onChange?.(); });
      options.onChange?.(); return flight;
    },
    requireReconciliation(reason: string) {
      if (flight) throw new Error("PDF publication is still in flight.");
      failure = reason; options.onChange?.();
    },
    confirmUnchangedBaseline(source: PdfEngineSourceIdentity, page: {id:string;content_json:Record<string,JsonValue>}) {
      checkAuthority();
      const journal = page.content_json.pdfNativeTransactions;
      if (flight || !failure || page.id !== options.pageId || !jsonEquals(source as unknown as JsonValue, savedSource as unknown as JsonValue) ||
        (journal !== undefined && !Array.isArray(journal)) ||
        (Array.isArray(journal) && journal.some(entry => entry && typeof entry === "object" && !Array.isArray(entry) && pending.some(op => op.id === entry.id))))
        throw new Error("PDF recovery baseline or publication state changed.");
      failure = null; options.onChange?.();
    },
    reconcile(ack: PdfWorkingSaveAcknowledgement) {
      checkAuthority();
      if (flight || !failure || !pending[0]) throw new Error("PDF save is not awaiting reconciliation.");
      const operation = pending[0], data = owned.get(operation)!;
      if (data.kind === "text" && data.sent) {
        acknowledge(operation, data.sent, ack);
        return;
      }
      if (data.kind === "engine") {
        if (!data.sent) throw new Error("PDF recovery has no record sent for this edit.");
        const derived = (data.sent as { derivedCandidate?: { sha256?: unknown } }).derivedCandidate?.sha256;
        acknowledge(operation, data.sent, ack, typeof derived === "string" ? derived : operation.outputSha256);
        return;
      }
      if (!data.checkpoint) throw new Error("PDF recovery has no authenticated saved checkpoint.");
      const record = bindMraiNativePreparedTextRecord({ source: savedSource, transactionId: operation.id,
        workingRevision: operation.baseWorkingRevision, selectionRevision: data.plan.evidence.selection!.selectionRevision,
        prepared: data.prepared, fidelity: data.plan.evidence, sourceCheckpoint: data.checkpoint });
      acknowledge(operation, record, ack);
    },
    dispose() {
      disposed = true; workingBytes.fill(0); savedBytes.fill(0); undone.length = 0;
      for (const operation of owned.keys()) clearOperation(operation);
      pending.length = 0; options.onChange?.();
    },
  };
  return session;
}


/** The queue's failure when its save met a version another window saved
 * first (409): no record landed. The window refreshes the page onto that
 * version and rebases the copy with its queue kept (rebaseSaved
 * `keepPending`), which plans the queue again on the saved bytes. */
export const PDF_WORKING_PEER_VERSION_FAILURE = "peer-version";

export function createPdfNativeWorkingSession(options: PdfWorkingSessionOptions) {
  return createWorkingSession(options);
}

/** Recompile saved recipes against their pinned original base. Stored data
 * does not confer writer authority, and a recovered queue never auto-retries. */









export function advancePdfWorkingRecovery(recovery: PdfWorkingRecovery, source: PdfEngineSourceIdentity): PdfWorkingRecovery | null {
  const landed = recovery.operations.findIndex(operation => operation.outputSha256 === source.sourceSha256);
  if (landed < 0) return null;
  const operation = recovery.operations[landed]!;
  return Object.freeze({ ...recovery, savedSource: source,
    savedRevision: operation.baseWorkingRevision + 1, operations: recovery.operations.slice(landed + 1) });
}

export async function restorePdfNativeWorkingSession(options: PdfWorkingSessionOptions, recovery: PdfWorkingRecovery) {
  if (recovery.schemaVersion !== "mrai.pdf-working-recovery.v1" || recovery.pageId !== options.pageId ||
    !Number.isSafeInteger(recovery.savedRevision) || recovery.savedRevision < 0 ||
    !jsonEquals(recovery.savedSource as unknown as JsonValue, options.source as unknown as JsonValue) ||
    !Array.isArray(recovery.operations) || recovery.operations.length > 32)
    throw new Error("PDF working recovery identity is invalid.");
  let opened = false;
  const session = await createWorkingSession({ ...options,
    writeRecovery: state => { if (opened) options.writeRecovery(state); },
    onChange: () => { if (opened) options.onChange?.(); },
  }, recovery.savedRevision);
  try {
    // A replannable edit's output as recorded -> as it came back.
    const cameBackAs = new Map<string, string>();
    for (const recipe of recovery.operations) {
      const snapshot = session.snapshot();
      const base = cameBackAs.get(recipe.baseSha256) ?? recipe.baseSha256;
      if (snapshot.workingRevision !== recipe.baseWorkingRevision || snapshot.workingSha256 !== base)
        throw new Error("PDF recovered operation lineage is invalid.");
      // An edit made on bytes that came back otherwise comes back otherwise too.
      const onOtherBytes = base !== recipe.baseSha256;
      if (recipe.kind === "engine") {
        if (!options.prepareEngineRecipe) throw new Error("PDF recovered engine edit cannot be planned here.");
        const bytes = session.copyWorkingBytes();
        try {
          const prepared = await options.prepareEngineRecipe({ recipe: recipe.recipe, id: recipe.id, sourceBytes: bytes,
            sourceSha256: snapshot.workingSha256, savedSource: snapshot.savedSource });
          if (prepared.outputSha256 !== recipe.outputSha256) {
            if (!onOtherBytes && !options.recipeReplannable?.(recipe.recipe)) {
              throw new Error("PDF recovered operation no longer reproduces its exact candidate.");
            }
            cameBackAs.set(recipe.outputSha256, prepared.outputSha256);
          }
          session.stageEngine(prepared);
        } finally { bytes.fill(0); }
        continue;
      }
      const operation = await session.prepareText({ id: recipe.id, selection: recipe.selection, replacementText: recipe.replacementText });
      if (operation.outputSha256 !== recipe.outputSha256) {
        if (!onOtherBytes) throw new Error("PDF recovered operation no longer reproduces its exact candidate.");
        cameBackAs.set(recipe.outputSha256, operation.outputSha256);
      }
      session.stage(operation);
    }
    if (recovery.operations.length) session.requireReconciliation("Recovered PDF edits require publication reconciliation.");
    options.writeRecovery(session.recovery()); opened = true;
    return session;
  } catch (error) { session.dispose(); throw error; }
}
