/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import type {
  PageContentDocument,
  PageContentSyncController,
} from "../../hooks/usePageContentSync";
import type { PageItem } from "../notebook-sdk";
import type { CollabSession } from "./collab-provider";
import type {
  PageDocumentAdapterReadScope,
  PageDocumentAdapterSource,
  PageDocumentCollaborationAdapter,
} from "@/lib/collab/page-document-adapter";
import { jsonEquals } from "./json-crdt";
import { yDocRevision } from "./y-doc-revision";
import { yieldBetweenSaveSteps } from "./save-yield";
import {
  PageCausalCheckpointError,
  applyPageCausalRecovery,
  capturePageCausalCheckpoint,
  capturePageCausalRecovery,
  markProvenProjection,
  pageCausalVerifiedCopy,
  pageForCausalCheckpoint,
  durablePageSnapshotOffer,
  preparePageCausalCheckpoint,
  preparePageCausalSaveReceipt,
  type PreparedPageCausalCheckpoint,
  type PageCausalModel,
} from "./page-causal-checkpoint";

export interface NativePageCausalHost {
  pageId: string;
  session(): CollabSession;
  canWrite(): boolean;
  hasUnboundDraft(): boolean;
  /** Take the unbound edits out of the way of a new binding: the host keeps
   * its view's body (with `base`, the last body the server confirmed, when
   * known) and merges it into the page once the model is ready again. False
   * when it cannot; the model then fails closed as before. */
  setAsideUnboundDraft?(base: PageContentDocument | null): boolean;
  requiresNativeBinding(): boolean;
  /** Rebind the registry's sole model. Never copy old room bytes to a new epoch. */
  rebind(epoch: number | null, joinLegacyRoom?: boolean): void;
  deliver(content: PageContentDocument): void;
  changed(): void;
  /** Whether the view holding `lease` applies what the room delivers (a
   * document's paged bridge stops applying it when it has an issue). Absent:
   * a connected room's content is the view's. */
  roomContentApplied?(lease: string): boolean;
  /** Every view of the page shows exactly what the room holds (each proves
   * it from its own live state). Absent: unproven. */
  viewMirrorsRoom?(): boolean;
  /** Why the last viewMirrorsRoom() said no, as a content-free code. */
  viewMirrorRefusal?(): string | null;
}

export interface NativePageCausalModel {
  model: PageCausalModel;
  projection(doc: Y.Doc): PageContentDocument;
  canMirror(content: PageContentDocument, baseline: PageContentDocument, doc: Y.Doc): boolean;
  apply(doc: Y.Doc, content: PageContentDocument, baseline: PageContentDocument): boolean;
  contentForSave?(projection: PageContentDocument, requested: PageContentDocument): PageContentDocument | null;
  /** The projection mirror() keeps as its baseline, when the model can read
   * it more cheaply than `projection` (it is never handed to consumers). */
  baselineProjection?(doc: Y.Doc): PageContentDocument;
  /** A non-null contentForSave proves the returned body matches the
   * baseline's projection exactly, as capturePageCausalRecovery checks. */
  saveProvesProjection?: boolean;
  /** canMirror proves a body against the live room itself and apply reads
   * no baseline (Docs): the baseline is only the model's own read of the
   * room, which a mirror asked before any delivery reached a view reads then. */
  mirrorsAgainstRoom?: boolean;
  /** Where the body canMirror last refused differs from the room, as a
   * content-free structural path for save diagnostics. */
  mirrorMiss?(): string | null;
  /** The room's own body, as a save body: what a save sends when the
   * requested body differs from the room only by how it was serialized
   * (every view proves it shows the room, `viewMirrorsRoom`). */
  roomBody?(doc: Y.Doc): PageContentDocument | null;
  shouldDeliver(session: CollabSession, transaction: Y.Transaction): boolean;
  /** Only a model with actual compatible old-room originals can propose
   * epoch1. Docs' physical pageBlock room is deliberately incompatible. */
  seedLegacy?(session: CollabSession, initial: PageContentDocument, first: boolean): void;
}

const IDLE: PageContentSyncController = Object.freeze({
  mode: "legacy",
  presence: null,
  persistContent: async () => null,
});

/** Everything that makes a page's checkpoint the one it is: two answers with
 * the same identity hold the very same bytes at the very same version. */
export function checkpointIdentity(
  checkpoint: NonNullable<PageItem["page_causal_checkpoint"]>,
): string {
  return [checkpoint.page_id, checkpoint.epoch, checkpoint.revision, checkpoint.content_version,
    checkpoint.content_sha256, checkpoint.checkpoint?.model, checkpoint.checkpoint?.update_sha256].join(":");
}

/** One adapter state per buffer, with independent per-view registrations. */
export function createNativePageCausalAdapter(host: NativePageCausalHost, codec: NativePageCausalModel) {
  const INITIAL = Object.freeze({ kind: `${codec.model}-initial` });
  const LOCAL = Object.freeze({ kind: `${codec.model}-local` });
  const RECEIPT = Object.freeze({ kind: `${codec.model}-own-receipt` });
  const REMOTE = Object.freeze({ kind: `${codec.model}-durable` });
  const registrations = new Map<string, PageDocumentAdapterSource>();
  const staleRefreshes = new Map<string, PageDocumentAdapterSource>();
  const leases = new Set<string>();
  const listeners = new Set<() => void>();
  let disposed = false,
    active = false,
    applying = false;
  let scopeKey: string | null = null,
    epoch: number | null = null;
  let status: "preparing" | "ready" | "failed" = "preparing";
  let saved = false,
    receipt: PreparedPageCausalCheckpoint | null = null;
  // Content-free save diagnostics (lib/page-save-diagnostics): why the model
  // last failed and why the last capture returned nothing.
  let failure: string | null = null,
    captureRefusal: string | null = null,
    // Where the room-body fallback of the last refused capture stopped, and
    // where its requested body first differed from the room.
    captureFallback: string | null = null,
    captureMiss: string | null = null;
  // Saves that sent the room's own body (see capture): content-free count.
  let roomBodySaves = 0;
  const diagnostics = () => ({
    failure: status === "failed" ? failure : null,
    captureRefusal,
    mirrorMiss: captureRefusal === "editor_does_not_mirror_original" ? captureMiss : null,
    captureFallback,
    roomBodySaves,
    bound: scopeKey !== null,
    unboundDraft: host.hasUnboundDraft(),
  });
  let baseline: PageContentDocument | null = null;
  // An adoption whose delivery reached no confirmed view (a kept copy's; one
  // whose view was gone when its delivery ran): the page's sources were never
  // handed the room, whatever baseline a mirror has read since.
  let undelivered = false;
  const ownRead = () => codec.baselineProjection
    ? codec.baselineProjection(host.session().ydoc) : projection();
  // The document revision mirror() last read its baseline at, when that
  // baseline came from `baselineProjection`.
  let mirroredRevision: string | null = null;
  let verifiedPage: PageItem | null = null;
  // The page object verifiedPage was copied from.
  let verifiedFrom: PageItem | null = null;
  let observed: CollabSession | null = null,
    unwatch: (() => void) | null = null;
  let snapshot = IDLE;
  let bootstrap: Promise<void> | null = null;
  let legacyAdmission: CollabSession | null = null;
  let retryQueued = false,
    retryAfterBootstrap = false;
  const sourceCurrent = (s: PageDocumentAdapterSource) =>
    !disposed && s.scopeKey === scopeKey && s.isCurrent();
  // A kept copy shown while the server is read (`preview`) builds the model
  // and nothing else: every read, write, save and seed below asks for a
  // source the server answered.
  const confirmedSource = (s: PageDocumentAdapterSource) => !s.preview && sourceCurrent(s);
  const current = () => [...registrations.values()].some(confirmedSource);
  const shown = () => [...registrations.values()].some(sourceCurrent);
  // The model holds a kept copy no server answer has confirmed yet.
  let previewUnconfirmed = false;
  // That copy's checkpoint, when only the very same checkpoint may confirm
  // it (`identicalOnly`: a page shown ahead of its commit). Another answer is
  // never merged into it: its originals may be another seed's.
  let previewIdentity: string | null = null;
  const writable = () =>
    current() &&
    host.canWrite() &&
    [...registrations.values()].some((s) => confirmedSource(s) && s.canWrite());
  const projection = (): PageContentDocument => codec.projection(host.session().ydoc);
  const forSave = (value: PageContentDocument, requested: PageContentDocument) =>
    codec.contentForSave ? codec.contentForSave(value, requested) : value;
  const notify = () => {
    if (disposed) return;
    if (!active) snapshot = IDLE;
    else if (
      snapshot.causal?.status !== status ||
      snapshot.causal.saved !== saved
    ) {
      snapshot = {
        ...IDLE,
        flushContent: (content) => (content ? mirror(content) : read()),
        causal: {
          status,
          saved,
          readContent: read,
          mirror,
          capture,
          adopt: (page, canApply) => adopt(page, undefined, canApply),
          adoptSaveReceipt: (page, submitted) => adopt(page, submitted),
          prepareMaintenance: async () => null,
          diagnostics,
          // A kept copy is drawn and the server's answer has not confirmed it yet.
          confirming: () => active && status === "ready" && previewUnconfirmed,
          closingUpdate,
          // The room is connected and its current socket replayed: a save
          // another live member made reached this document already.
          liveInRoom: () => {
            const network = host.session().network;
            return active && status === "ready" && current() && network?.ready === true &&
              network.connected === true && !network.failed;
          },
        },
      };
    }
    for (const listener of [...listeners]) listener();
    host.changed();
  };
  function read() {
    return active && status === "ready" && current() ? projection() : null;
  }
  /** What the document holds beyond the durable checkpoint it last adopted,
   * for a window that is closing (lib/page-closing-edits): the structs past
   * that checkpoint's state vector and the document's whole delete set. An
   * Undo typed a moment before closing is a deletion in it, whether or not
   * its save left. Synchronous: a closing page has no later task. Null when
   * the checkpoint holds everything, or this view cannot write. */
  function closingUpdate() {
    if (!active || status !== "ready" || epoch === null || !receipt || !writable()) return null;
    const session = host.session();
    // What the room has not been sent yet goes with the socket's last frames.
    try { session.network?.flushLocalUpdates?.(); } catch { /* the request below carries it */ }
    const doc = session.ydoc;
    if (receipt.covers(doc)) return null;
    const beyond = receipt.beyond(doc);
    return beyond ? { update: beyond.update, epoch: receipt.epoch, model: codec.model,
      baseContentVersion: receipt.contentVersion } : null;
  }
  function deliver(notifyRemote = true) {
    // Nothing goes to the page's sources (this author's own receipt): the
    // baseline is the model's own read (a document's is kept from read to
    // read), not a fresh copy of the whole room for nobody.
    const content = !notifyRemote && codec.baselineProjection
      ? codec.baselineProjection(host.session().ydoc) : projection();
    baseline = content;
    if (notifyRemote) undelivered = false;
    host.deliver(content);
    for (const source of [...registrations.values()])
      if (notifyRemote && sourceCurrent(source))
        source.onRemoteContent(content);
  }
  function staleSource(page: PageItem) {
    const incoming = page.page_causal_checkpoint;
    return (
      !!receipt &&
      !!incoming &&
      incoming.epoch === receipt.epoch &&
      (incoming.revision < receipt.revision ||
        incoming.content_version < receipt.contentVersion)
    );
  }
  function refreshStaleSource(
    source: PageDocumentAdapterSource,
    token: string,
    page = source.page,
  ) {
    if (!staleSource(page)) return false;
    if (
      registrations.get(token) === source &&
      sourceCurrent(source) &&
      staleRefreshes.get(token) !== source
    ) {
      staleRefreshes.set(token, source);
      // A stale view's refresh failure is not a failure of the healthy shared
      // originals. Its existing loader owns the error/retry presentation.
      try {
        source.onReconnect();
      } catch {
        /* keep the current model */
      }
    }
    return true;
  }
  function watch() {
    const session = host.session();
    if (observed === session) return;
    unwatch?.();
    observed = session;
    const update = (
      _bytes: Uint8Array,
      _origin: unknown,
      _doc: Y.Doc,
      transaction: Y.Transaction,
    ) => {
      if (
        applying ||
        !active ||
        status !== "ready" ||
        !current() ||
        host.session() !== session
      )
        return;
      saved = false;
      if (codec.shouldDeliver(session, transaction)) deliver();
      notify();
    };
    session.ydoc.on("update", update);
    const seedWhenReady = () => {
      if (
        observed === session &&
        host.session() === session &&
        current() &&
        status === "ready"
      )
        session.provider?.seedDurablePageState?.();
    };
    void session.network?.whenReady
      .then(() => {
        seedWhenReady();
        retryPreparation();
      })
      .catch(() => undefined);
    const unsubscribeWritable = session.provider?.subscribeWritable?.(() =>
      retryPreparation(),
    );
    let connected = session.network?.connected ?? false;
    const unsubscribe = session.network?.subscribeConnection?.((next) => {
      const reconnect = next && !connected;
      connected = next;
      if (next) {
        seedWhenReady();
        retryPreparation();
      }
      if (reconnect && current() && status === "ready")
        for (const source of registrations.values())
          if (sourceCurrent(source)) source.onReconnect();
    });
    unwatch = () => {
      session.ydoc.off("update", update);
      unsubscribe?.();
      unsubscribeWritable?.();
    };
  }
  function mirror(content: PageContentDocument) {
    if (!active || status !== "ready" || !writable()) return null;
    // A model ready before any delivery reached a confirmed view (a kept
    // copy's adoption delivers to none; the server's answer is adopted in
    // steps, its delivery in a task of its own) had no baseline, and every
    
    // exports and Copilot bindings of unedited Docs pages). A model whose
    // mirror proves the body against the live room reads it now.
    if (!baseline && codec.mirrorsAgainstRoom) baseline = ownRead();
    if (!baseline || !codec.canMirror(content, baseline, host.session().ydoc))
      return null;
    const session = host.session();
    let changed = false;
    applying = true;
    try {
      session.ydoc.transact(() => { changed = codec.apply(session.ydoc, content, baseline!); }, LOCAL);
    } finally {
      applying = false;
    }
    const ydoc = host.session().ydoc;
    if (codec.baselineProjection) {
      baseline = codec.baselineProjection(ydoc);
      mirroredRevision = yDocRevision(ydoc);
    } else {
      baseline = projection();
      mirroredRevision = null;
    }
    if (changed) saved = false;
    notify();
    return forSave(baseline, content);
  }
  function capture(
    content: PageContentDocument,
    baseTitle: string,
    allowMirror = true,
  ) {
    if (!active || status !== "ready" || epoch === null || !current()) {
      captureRefusal = !active ? "model_inactive" : status !== "ready" ? `model_${status}`
        : epoch === null ? "epoch_unknown" : "source_not_current";
      return null;
    }
    const mayMirror = allowMirror && writable();
    mirroredRevision = null;
    let value = mayMirror ? mirror(content) : forSave(projection(), content);
    captureMiss = value ? null : codec.mirrorMiss?.() ?? null;
    captureFallback = null;
    // The requested body is the editor's save serialization; the room is what
    // every view shows. When they differ but each view proves it shows the
    
    // its views and room agreed), the difference is the serialization's, not
    // an edit the room missed: the save sends the room's own body. A view
    // that does not prove it keeps the refusal (fail closed).
    if (!value && mayMirror) {
      if (!baseline) captureFallback = "no_baseline";
      else if (!codec.roomBody) captureFallback = "no_room_body";
      else if (!host.viewMirrorsRoom?.()) captureFallback = `view_unproven:${host.viewMirrorRefusal?.() ?? "-"}`;
      else {
        const room = codec.roomBody(host.session().ydoc);
        if (room) value = mirror(room);
        if (value) roomBodySaves += 1;
        else captureFallback = room ? "room_body_not_mirrored" : "room_body_unread";
      }
    }
    // mirror() proved `value` matches the projection of the document at the
    // revision it read; nothing changed the document since, so the recovery
    // capture's own check of that projection would repeat the same proof.
    const projectionProven = Boolean(value) && mayMirror && codec.saveProvesProjection === true &&
      mirroredRevision !== null && mirroredRevision === yDocRevision(host.session().ydoc);
    // A freshness pause permits custody of already-held originals, not edits.
    if (!value || (!mayMirror && !jsonEquals(value, content))) {
      captureRefusal = mayMirror ? "editor_does_not_mirror_original" : "held_originals_differ";
      return null;
    }
    try {
      captureRefusal = null;
      return {
        content: value,
        recovery: capturePageCausalRecovery(host.session().ydoc, value, {
          pageId: host.pageId,
          epoch,
          baseTitle,
          model: codec.model,
        }, { projectionProven, durable: receipt }),
      };
    } catch (error) {
      // Relayed updates still waiting for their predecessors: the room will
      // deliver them (the transport resumes for them), so the save waits.
      if (error instanceof PageCausalCheckpointError && error.code === "incomplete") throw error;
      captureRefusal = error instanceof PageCausalCheckpointError ? `checkpoint_${error.code}` : "checkpoint_error";
      return null;
    }
  }
  async function adopt(
    page: PageItem,
    submitted?: Parameters<
      NonNullable<PageContentSyncController["causal"]>["adoptSaveReceipt"]
    >[1],
    canApply?: () => boolean,
    asPreview = false,
    identicalOnly = false,
  ) {
    const durable = page.page_causal_checkpoint,
      session = host.session(),
      guardScope = scopeKey;
    if (
      !(asPreview ? shown() : current()) ||
      page.id !== host.pageId ||
      !durable ||
      (epoch !== null && durable.epoch !== epoch)
    )
      return false;
    const preparing = submitted
      ? preparePageCausalSaveReceipt(
          durable,
          pageForCausalCheckpoint(page),
          submitted,
        )
      : preparePageCausalCheckpoint(
          durable,
          pageForCausalCheckpoint(page),
        );
    // The room this window joins may hold exactly these bytes as its snapshot.
    session.provider?.offerDurablePageSnapshot?.(durablePageSnapshotOffer(durable, preparing));
    const prepared = await preparing;
    // A Docs save adopts its source and then its own receipt: each merge and
    // delivery in a task of its own (everything below is checked again).
    if (codec.model === "document-canonical-v1") await yieldBetweenSaveSteps();
    if (
      !(asPreview ? shown() : current()) ||
      scopeKey !== guardScope ||
      session !== host.session() ||
      !(canApply?.() ?? true) ||
      (epoch !== null && prepared.epoch !== epoch) ||
      (receipt &&
        (prepared.revision < receipt.revision ||
          prepared.contentVersion < receipt.contentVersion))
    )
      return false;
    applying = true;
    // The very checkpoint this model holds, adopted again into a room that
    // already held all of it (each save adopts its source page first):
    // nothing arrived, so nothing is delivered. Delivering it read the whole
    // room and handed it to the page as if a peer had written it, which then
    // stored a recovery copy of it -- every save, under the pen.
    const roomBefore = yDocRevision(session.ydoc);
    const again = !submitted && status === "ready" && receipt !== null &&
      receipt.snapshot.sha256 === prepared.snapshot.sha256 && receipt.revision === prepared.revision &&
      receipt.contentVersion === prepared.contentVersion;
    let nothingArrived = false;
    try {
      if (submitted && session.provider?.withDurablePageUpdate)
        session.provider.withDurablePageUpdate(RECEIPT, () =>
          prepared.merge(session.ydoc, RECEIPT),
        );
      else if (submitted && session.provider)
        return false; // no untracked resend fallback
      else
        prepared.merge(
          session.ydoc,
          submitted ? RECEIPT : (session.provider ?? REMOTE),
        );
      nothingArrived = again && roomBefore !== null && yDocRevision(session.ydoc) === roomBefore;
      receipt = prepared;
      if (!(nothingArrived && verifiedFrom === page && verifiedPage)) {
        verifiedPage = pageCausalVerifiedCopy(page);
        verifiedFrom = page;
      }
      epoch = prepared.epoch;
      saved = prepared.covers(session.ydoc);
      status = "ready";
      failure = null;
      if (asPreview) {
        previewUnconfirmed = true;
        // In the same step: no answer is ever weighed against this copy
        // without knowing whether only its very checkpoint confirms it.
        previewIdentity = identicalOnly ? checkpointIdentity(durable) : null;
        // The kept body was just found to be what these bytes project to:
        // the server's answer that nothing changed is this same object, and
        // its adoption below need not decode and compare it a second time.
        if (page.content_json) markProvenProjection(page.content_json, prepared.snapshot.sha256);
      } else if (previewUnconfirmed && !submitted) {
        // The first adoption of a server answer after a kept copy: whatever
        // the room delivered meanwhile (the page's source heard none of it)
        // is delivered now; the same checkpoint, nothing beyond it, is not.
        previewUnconfirmed = false;
        previewIdentity = null;
        if (!saved) nothingArrived = false;
      }
    } finally {
      applying = false;
    }
    // Nothing delivered keeps the model's baseline (what its views were last
    // given) -- but a kept copy's adoption gives no confirmed view anything,
    // so after one there was none and every save was refused "no_baseline"
    
    // reopened from its kept copy, unsaved since the day before). A room that
    // is the checkpoint is what the views show; any other room is delivered.
    // (A mirror may have read a baseline since an undelivered adoption: the
    // sources still never had the room.)
    if (nothingArrived && !asPreview && (undelivered || !baseline)) {
      if (!saved) nothingArrived = false;
      else {
        if (!baseline) baseline = ownRead();
        // The source adopted holds this very checkpoint, which the room is.
        undelivered = false;
      }
    }
    if (codec.model === "document-canonical-v1" && !nothingArrived) {
      // The merge, the delivery and what the page builds from it (opening a
      // long document: its editor and the fill) each in a task of their own:
      // in one, the 298-page book's open held the main thread 2.3-2.7 s.
      await yieldBetweenSaveSteps();
      if (session === host.session()) {
        if (current()) deliver(!submitted);
        else if (!submitted) undelivered = true;
      }
      await yieldBetweenSaveSteps();
      notify();
      return true;
    }
    if (!nothingArrived) deliver(!submitted);
    notify();
    return true;
  }
  function prepare(source: PageDocumentAdapterReadScope) {
    if (disposed || !source.isCurrent() || source.page.id !== host.pageId)
      return;
    const checkpoint = source.page.page_causal_checkpoint;
    if (!checkpoint && !source.page.page_causal_activation_available) return;
    // A reader of a page no writer has upgraded has nothing to prepare: only
    // someone who may write makes the upgrade (page_causal_can_activate),
    // and it arrives with the next page read. Preparing anyway met the
    // reader's own display of the saved body as an "unbound draft" and
    
    if (!checkpoint && source.page.page_causal_can_activate === false) return;
    if (scopeKey === source.scopeKey && staleSource(source.page)) return;
    const nextEpoch = checkpoint?.epoch ?? null;
    // A page shown ahead of its commit, answered with another checkpoint
    // than the one shown: opened afresh from the answer (a new document),
    // as a first open without that copy would be. Nothing was written to
    // the copy (a preview is read-only), so nothing is set aside.
    const identityMismatch = previewIdentity !== null && previewUnconfirmed && !source.preview &&
      (!checkpoint || checkpointIdentity(checkpoint) !== previewIdentity);
    if (
      active &&
      scopeKey === source.scopeKey &&
      epoch === nextEpoch &&
      status !== "failed" &&
      !identityMismatch
    )
      return;
    previewIdentity = null;
    active = true;
    status = "preparing";
    saved = false;
    previewUnconfirmed = false;
    // Old unbound edits stay recoverable and fail closed; never transplant
    // reconstructed text or pending old-room updates into migrated originals.
    // A host that can keep them aside (and merge them back as new edits once
    // the new binding is ready) does, instead of leaving the page unsaveable
    
    if (
      host.hasUnboundDraft() &&
      (scopeKey !== source.scopeKey || epoch !== nextEpoch) &&
      !host.setAsideUnboundDraft?.(
        (verifiedPage?.content_json as PageContentDocument | undefined) ?? null)
    ) {
      status = "failed";
      failure = scopeKey === null ? "unbound_draft_at_first_preparation"
        : scopeKey !== source.scopeKey ? "unbound_draft_at_scope_change" : "unbound_draft_at_epoch_change";
      host.session().provider?.suspendLocalDocumentUpdates?.();
      notify();
      return;
    }
    const replace =
      identityMismatch ||
      host.requiresNativeBinding() ||
      (scopeKey !== null && scopeKey !== source.scopeKey) ||
      (nextEpoch !== null && nextEpoch > 1 && nextEpoch !== epoch) ||
      !checkpoint;
    scopeKey = source.scopeKey;
    epoch = nextEpoch;
    receipt = null;
    verifiedPage = null;
    baseline = null;
    undelivered = false;
    if (replace) host.rebind(nextEpoch);
    watch();
    notify();
  }
  async function initialize(source: PageDocumentAdapterSource, token: string) {
    let session = host.session();
    const valid = () =>
      !disposed &&
      registrations.get(token) === source &&
      source.isCurrent() &&
      host.session() === session &&
      source.scopeKey === scopeKey;
    let step = "start";
    try {
      let page = source.page;
      if (refreshStaleSource(source, token)) return;
      // A kept copy is only ever shown as the checkpoint it holds: no
      // upgrade, recovery or room seed is made from it.
      // A copy that cannot be shown (older than the model, unreadable) is
      // simply not: the model keeps preparing for the server's answer.
      if (source.preview) {
        if (!page.page_causal_checkpoint || source.recovered) return;
        try {
          if (await adopt(page, undefined,
            () => registrations.get(token) === source && sourceCurrent(source), true,
            source.identicalOnly === true)) notify();
        } catch { /* the server's answer prepares the model as before */ }
        return;
      }
      if (!page.page_causal_checkpoint) {
        step = "legacy_activation";
        if (source.recovered)
          throw Error("Unbound local changes require recovery");
        if (!codec.seedLegacy) throw Error("Canonical migration requires an inactive legacy room");
        if (!source.canWrite() || !host.canWrite()) return;
        // prepareForSource retired the old admission to allow the server's
        // inactive migration. If that migration could not run, join the real
        // legacy room and wait for its original replay before proposing epoch1.
        if (legacyAdmission !== session) {
          host.rebind(null, true);
          session = host.session();
          legacyAdmission = session;
        }
        watch();
        const network = session.network;
        if (!network) throw Error("Original room election is unavailable");
        const election = await network.whenReady;
        if (
          !valid() ||
          !source.canWrite() ||
          !host.canWrite() ||
          !network.connected ||
          network.failed ||
          !network.writable
        )
          return;
        const initial = pageForCausalCheckpoint(page).content_json;
        if (!session.provider?.withDurablePageUpdate)
          throw Error("Controlled initial-state admission is unavailable");
        applying = true;
        try {
          session.provider.withDurablePageUpdate(INITIAL, () =>
            session.ydoc.transact(() => codec.seedLegacy!(session, initial, election.first), INITIAL),
          );
        } finally {
          applying = false;
        }
        const checkpoint = await capturePageCausalCheckpoint(
          session.ydoc,
          initial,
          codec.model,
        );
        if (
          !valid() ||
          !source.canWrite() ||
          !host.canWrite() ||
          !network.connected ||
          network.failed ||
          !network.writable
        )
          return;
        if (!jsonEquals(projection(), initial))
          throw Error("Room changed during activation preparation");
        page = await source.activate(page, checkpoint);
        if (!valid()) return;
        // Activation can migrate the room. The seed is not personal history.
        if ((page.page_causal_checkpoint?.epoch ?? 0) > 1) {
          host.rebind(page.page_causal_checkpoint!.epoch);
          watch();
        }
      }
      if (!source.isCurrent() || registrations.get(token) !== source) return;
      step = "recovery_identity";
      if (
        source.recovered &&
        (source.recovered.recovery.page_id !== host.pageId ||
          source.recovered.recovery.model !== codec.model ||
          source.recovered.recovery.epoch !==
            page.page_causal_checkpoint?.epoch)
      )
        throw Error("Recovered originals belong to another activation");
      step = "adopt_durable_originals";
      if (
        !(await adopt(
          page,
          undefined,
          () => registrations.get(token) === source && sourceCurrent(source),
        ))
      ) {
        if (
          refreshStaleSource(source, token, page) ||
          !sourceCurrent(source) ||
          registrations.get(token) !== source
        )
          return;
        step = "durable_originals_unavailable";
        throw Error("Durable originals unavailable");
      }
      if (source.recovered) {
        step = "apply_recovery";
        applying = true;
        try {
          applyPageCausalRecovery(
            host.session().ydoc,
            source.recovered.recovery,
            source.recovered.content,
            host.pageId,
            epoch!,
            REMOTE,
            codec.model,
          );
        } finally {
          applying = false;
        }
        saved = receipt!.covers(host.session().ydoc);
        deliver();
      }
      host.session().provider?.seedDurablePageState?.();
      notify();
    } catch (error) {
      if (registrations.get(token) === source && source.isCurrent()) {
        status = "failed";
        failure = error instanceof PageCausalCheckpointError ? `${step}:checkpoint_${error.code}` : step;
        saved = false;
        notify();
      }
    }
  }
  function retryPreparation() {
    if (disposed || !active || status !== "preparing") return;
    if (bootstrap) {
      retryAfterBootstrap = true;
      return;
    }
    if (retryQueued) return;
    retryQueued = true;
    queueMicrotask(() => {
      retryQueued = false;
      if (disposed || !active || status !== "preparing") return;
      if (bootstrap) {
        retryAfterBootstrap = true;
        return;
      }
      if (!host.canWrite()) return;
      const network = host.session().network;
      if (
        legacyAdmission === host.session() &&
        network &&
        (!network.connected || network.failed || !network.writable)
      )
        return;
      const next = [...registrations].find(
        ([, source]) => sourceCurrent(source) && source.canWrite(),
      );
      if (next) start(next[1], next[0]);
    });
  }
  function start(source: PageDocumentAdapterSource, token: string) {
    if (!active || status === "failed" || bootstrap) return;
    const pending = initialize(source, token);
    bootstrap = pending;
    void pending.finally(() => {
      if (bootstrap !== pending) return;
      bootstrap = null;
      const retry = retryAfterBootstrap;
      retryAfterBootstrap = false;
      // A view can close while WebCrypto is pending. Its cleanup must neither
      // publish that stale result nor strand another view's live registration.
      if (status !== "failed" && registrations.get(token) !== source) {
        const next = [...registrations].find(([, s]) => sourceCurrent(s));
        if (next) start(next[1], next[0]);
      }
      if (retry) retryPreparation();
    });
  }
  return {
    active: () => active,
    ready: () => active && status === "ready",
    /** Why the model failed (a content-free code), while it has: nothing
     * opens the page until a new read of it prepares the model again. */
    failure: () => active && status === "failed" ? failure ?? "failed" : null,
    saved: () => saved,
    // A durable original-state witness, not an invocation ACK. New local
    // operations immediately make it unavailable until a receipt covers them.
    // The verified page is this model's own copy (adopt); its body is not
    // copied again on every call: the registry asks on every publish (each
    // room change), and copying a long page each time was 100 ms and more of
    // an inked page's save under the pen. Each caller gets its own record.
    savedPage: () => active && status === "ready" && saved && current() && verifiedPage
      ? { ...verifiedPage } : null,
    applying: () => applying,
    retryPreparation,
    mayWrite: () => !active || (status === "ready" && writable()),
    forLease(token: string): PageDocumentCollaborationAdapter {
      leases.add(token);
      let previous: PageContentSyncController | null = null,
        leaseSnapshot = IDLE;
      const live = () => {
        const source = registrations.get(token);
        return !!source && confirmedSource(source);
      };
      const mayMirror = () => live() && registrations.get(token)!.canWrite();
      const getSnapshot = () => {
        if (previous === snapshot) return leaseSnapshot;
        previous = snapshot;
        const causal = snapshot.causal;
        leaseSnapshot = causal
          ? {
              ...snapshot,
              flushContent: (c) =>
                mayMirror() ? (snapshot.flushContent?.(c) ?? null) : null,
              causal: {
                ...causal,
                readContent: () => (live() ? causal.readContent() : null),
                mirror: (c) => (mayMirror() ? causal.mirror(c) : null),
                capture: (c, t) => {
                  if (live()) return capture(c, t, mayMirror());
                  // Not the last capture's reason: this view's source is not
                  
                  captureRefusal = "view_not_confirmed";
                  captureFallback = null;
                  return null;
                },
                adopt: (p, guard) =>
                  live()
                    ? adopt(p, undefined, () => live() && (guard?.() ?? true))
                    : Promise.resolve(false),
                adoptSaveReceipt: (p, s) =>
                  live() ? adopt(p, s, live) : Promise.resolve(false),
                liveInRoom: () =>
                  live() && (causal.liveInRoom?.() ?? false) && (host.roomContentApplied?.(token) ?? true),
              },
            }
          : snapshot;
        return leaseSnapshot;
      };
      return {
        model: codec.model,
        prepareForSource: (source) => {
          if (leases.has(token)) prepare(source);
        },
        bindSource(source) {
          if (
            !leases.has(token) ||
            source.page.id !== host.pageId ||
            !source.isCurrent()
          )
            return () => {};
          prepare(source);
          registrations.set(token, source);
          watch();
          // A ready model may still have an older checkpoint task pending.
          // Source lifetime changes affect writability immediately, even if
          // start() cannot launch another bootstrap yet.
          notify();
          start(source, token);
          return () => {
            if (registrations.get(token) === source) {
              registrations.delete(token);
              notify();
            }
          };
        },
        getSnapshot,
        subscribe(listener) {
          if (!leases.has(token)) return () => {};
          listeners.add(listener);
          return () => {
            listeners.delete(listener);
          };
        },
      };
    },
    release(token: string) {
      staleRefreshes.delete(token);
      leases.delete(token);
      registrations.delete(token);
    },
    dispose() {
      disposed = true;
      verifiedPage = null;
      unwatch?.();
      registrations.clear();
      staleRefreshes.clear();
      leases.clear();
      listeners.clear();
    },
  };
}
