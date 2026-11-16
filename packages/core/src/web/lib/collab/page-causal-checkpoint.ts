/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

import { isSharedJsonRead, jsonEquals, yToJson, yToJsonShared, type JsonValue } from "./json-crdt";
import { CODE_COLLAB_TEXT_KEY } from "./code-collab-model";
import {
  documentContentSaysProjection, isDocumentSaveCopy, readDocumentCausalProjection, readSharedDocumentCausalProjection,
  retainDocumentCausalSaveProjection,
} from "./document-page-causal-model";
import { appendPdfCausalNativeCommand, pdfSharedContent, withPdfDurableJournal } from "./pdf-page-causal-model";
import { namePageCausalBody, type PageCausalNamedBody } from "./page-causal-named-body";
import { encodeYDocStateAndBeyond, yDocStatePieces, type YDocStatePiece } from "./y-state-encoder";
import { notePageCausalCapture, pageCausalJournalForm } from "./page-causal-journal";
import { yieldBetweenSaveSteps } from "./save-yield";

export type PageCausalModel = "page-json-v1" | "code-text-v1" | "document-canonical-v1" | "pdf-author-v1";

/** The original Yjs history accompanying one page-json save, including deletes.
 * The server must bind these bytes to the actual committed page revision.
 * A state vector alone cannot represent receipt of deletion-only changes. */
export interface PageCausalCheckpoint {
  readonly model: PageCausalModel;
  readonly update_base64: string;
  readonly update_sha256: string;
}

export interface PageCausalCheckpointReceipt {
  readonly page_id: string;
  readonly epoch: number;
  readonly revision: number;
  readonly content_version: number;
  readonly content_sha256: string;
  readonly checkpoint: PageCausalCheckpoint;
}

export interface PageCausalCheckpointPage {
  readonly model?: PageCausalModel;
  readonly id: string;
  readonly content_version: number;
  readonly content_sha256: string;
  readonly content_json: Record<string, JsonValue>;
}

export function pdfPageUsesAuthorCausalModel(page: {
  page_type?: string; page_causal_checkpoint?: PageCausalCheckpointReceipt | null;
} | null | undefined): boolean {
  return page?.page_type === "pdf" && page.page_causal_checkpoint?.checkpoint?.model === "pdf-author-v1" &&
    Number.isSafeInteger(page.page_causal_checkpoint.epoch) && page.page_causal_checkpoint.epoch >= 2;
}

export function pageForCausalCheckpoint(page: {
  id: string; page_type?: string; content_version?: number; content_sha256?: string;
  content_json?: Record<string, unknown>;
  page_causal_checkpoint?: PageCausalCheckpointReceipt | null;
}): PageCausalCheckpointPage {
  if (!page.id || !Number.isSafeInteger(page.content_version) || (page.content_version ?? 0) < 1 ||
      !DIGEST_PATTERN.test(page.content_sha256 ?? "") || !page.content_json || Array.isArray(page.content_json)) {
    throw new PageCausalCheckpointError("identity_mismatch");
  }
  return { id: page.id, content_version: page.content_version!, content_sha256: page.content_sha256!,
    ...(page.page_type === "code" ? { model: "code-text-v1" as const } : {}),
    ...(page.page_type === "document" ? { model: "document-canonical-v1" as const } : {}),
    ...(pdfPageUsesAuthorCausalModel(page) ? { model: "pdf-author-v1" as const } : {}),
    content_json: page.content_json as Record<string, JsonValue> };
}

/**
 * The causal model's own copy of a page it verified, safe from whatever the
 * caller later does to `page`. A save body (documentSaveCopy) is never
 * written into by anyone, so its body is kept as it is rather than copied:
 * on an inked page that copy was megabytes, every save receipt, under the pen.
 */
export function pageCausalVerifiedCopy<T extends { content_json?: unknown }>(page: T): T {
  const body = page.content_json;
  if (!body || !isDocumentSaveCopy(body)) return structuredClone(page);
  const copy = structuredClone({ ...page, content_json: null });
  (copy as { content_json?: unknown }).content_json = body;
  return copy as T;
}

/** Tab-local custody of original operations, not a server acknowledgement.
 * Capture is synchronous so navigation cannot outrun WebCrypto. The bytes are
 * decoded and their complete projection checked again before reuse. */
export interface PageCausalRecovery {
  readonly model: PageCausalModel;
  readonly page_id: string;
  readonly epoch: number;
  readonly base_title: string;
  readonly update_base64: string;
}

export function capturePageCausalRecovery(
  doc: Y.Doc,
  content: Record<string, JsonValue>,
  binding: { pageId: string; epoch: number; baseTitle: string; model?: PageCausalModel },
  /** `projectionProven`: the caller has just proven, at the document's
   * current revision, that `content` matches its projection exactly.
   * `durable`: the checkpoint the document is bound to (its save's delta is
   * taken from it now, with the bytes). */
  options?: { projectionProven?: boolean; durable?: PreparedPageCausalCheckpoint | null },
): PageCausalRecovery {
  requireComplete(doc);
  if (!binding.pageId || !Number.isSafeInteger(binding.epoch) || binding.epoch < 1) invalid();
  const model = binding.model ?? "page-json-v1";
  // What the document projects to, kept with the bytes for the save that
  // sends them (a document's: the shared projection the proof was made on).
  let projection: Record<string, JsonValue> | null = null;
  if (model === "document-canonical-v1") {
    try { projection = readSharedDocumentCausalProjection(doc); } catch { projection = null; }
  }
  if (options?.projectionProven) {
    // Proven by the caller's read; what the read did not check is checked.
    if (model !== "document-canonical-v1") requireProjectableShape(doc, model);
  } else if (!(projection ? matchesRead(projection, model, content) : matchesProjection(doc, model, content))) {
    throw new PageCausalCheckpointError("projection_mismatch");
  }
  // A deck's (page-json) body is its document's projection itself.
  if (model === "page-json-v1") projection = content;
  // The live document's encoding is kept struct by struct: only what changed
  // since the last capture is encoded again (y-state-encoder.ts).
  // What this state holds beyond the durable checkpoint the document was
  // last bound to is taken now with the bytes (the save sends it as its
  // delta): one pass of the encoder, which keeps each struct's encoding and
  // the document's extent, so only what changed is read (y-state-encoder.ts).
  const durable = options?.durable && (model === "document-canonical-v1" || model === "page-json-v1") &&
    options.durable.model === model ? options.durable : null;
  const known = durable ? preparedExtents.get(durable) : undefined;
  const held = known ? (() => { try { return known(); } catch { return null; } })() : null;
  const encoded = encodeYDocStateAndBeyond(doc, held ? held.clocks : null);
  const update = encoded.state;
  if (!update.length || update.length > MAX_UPDATE_BYTES) invalid();
  // The encoder hands back the same bytes for an unchanged document (a save
  // captures the same state more than once): their base64 is kept with them.
  let updateBase64 = capturedBase64.get(update);
  if (updateBase64 === undefined) capturedBase64.set(update, (updateBase64 = encodeBase64(update)));
  const recovery = Object.freeze({ model, page_id: binding.pageId,
    epoch: binding.epoch, base_title: binding.baseTitle, update_base64: updateBase64 });
  const pieces = yDocStatePieces(update);
  if (pieces) recoveryPieces.set(recovery, pieces);
  const extent: CausalExtent = encoded.extent ?? extentOf(doc);
  let beyond: CapturedBeyond | null = null;
  if (durable && held && encoded.beyond && holdsAll(extent, held)) {
    beyond = { durableSha256: durable.snapshot.sha256, durable, update: encoded.beyond,
      stateVector: Y.encodeStateVector(extent.clocks as Map<number, number>) };
  }
  captured.set(content, { updateBase64: recovery.update_base64, extent, update, projection, beyond });
  // A large state's local copy is a base and what changed since
  // (page-causal-journal.ts); a deck's or a code page's body is the state's
  // projection and is derived from it again on reopen.
  if (model === "page-json-v1" || model === "code-text-v1") {
    notePageCausalCapture(doc, recovery, update, updateBase64, content);
  }
  return recovery;
}

/** The field a journal copy's recovery carries its tail in (base64), its
 * `update_base64` then being the journal's base. */
export const PAGE_CAUSAL_JOURNAL_TAIL_FIELD = "update_journal_tail_base64";
/** What a journal copy's body is when it was left out (derived again). */
export const PAGE_CAUSAL_DERIVED_BODY_MARK = "$mraiBodyFromCausalState";

/**
 * A recovery as a journal copy stores it: its base (update_base64) and tail,
 * and whether `body` may be left out. Null when the recovery has no journal
 * whose base the store holds (the copy carries it whole).
 */
export function pageCausalRecoveryJournalCopy(
  recovery: PageCausalRecovery,
  body: unknown,
): { recovery: Record<string, unknown>; bodyDerived: boolean } | null {
  const form = pageCausalJournalForm(recovery);
  if (!form?.baseStored) return null;
  return {
    recovery: { ...recovery, update_base64: form.baseBase64,
      ...(form.tail.length > 0 ? { [PAGE_CAUSAL_JOURNAL_TAIL_FIELD]: encodeBase64(form.tail) } : {}) },
    bodyDerived: form.body !== null && form.body === body,
  };
}

/**
 * A recovery a journal copy stored, whole again: its base and tail merged
 * into one complete state, and (a body left out) the body that state
 * projects to. Throws when they do not make a complete state of the model.
 */
export function expandPageCausalRecoveryJournal(
  recovery: Record<string, unknown>,
  bodyDerived: boolean,
): { recovery: PageCausalRecovery; body: Record<string, JsonValue> | null } {
  const tailText = recovery[PAGE_CAUSAL_JOURNAL_TAIL_FIELD];
  if (typeof recovery.update_base64 !== "string" || (tailText !== undefined && typeof tailText !== "string")) invalid();
  const base = decodeBase64(recovery.update_base64 as string);
  const update = typeof tailText === "string" && tailText.length > 0
    ? (() => { try { return Y.mergeUpdates([base, decodeBase64(tailText)]); } catch { return invalid(); } })()
    : base;
  if (!update.length || update.length > MAX_UPDATE_BYTES) invalid();
  const { [PAGE_CAUSAL_JOURNAL_TAIL_FIELD]: _tail, ...rest } = recovery;
  void _tail;
  const whole = Object.freeze({ ...rest, update_base64: update === base ? recovery.update_base64 as string : encodeBase64(update) }) as
    unknown as PageCausalRecovery;
  if (!bodyDerived) return { recovery: whole, body: null };
  const model = whole.model;
  if (model !== "page-json-v1" && model !== "code-text-v1") invalid();
  const doc = new Y.Doc();
  try {
    try { Y.applyUpdate(doc, update); } catch { invalid(); }
    requireComplete(doc);
    const body = readProjection(doc, model);
    // This body is these bytes' own projection: a save of it does not read
    // the bytes back to compare them again.
    captured.set(body, { updateBase64: whole.update_base64, extent: extentOf(doc) });
    return { recovery: whole, body };
  } finally { doc.destroy(); }
}

const capturedBase64 = new WeakMap<Uint8Array, string>();

/** A recovery this module captured -> its bytes as the encoder's pieces. */
const recoveryPieces = new WeakMap<PageCausalRecovery, readonly YDocStatePiece[]>();

/**
 * The bytes of a recovery's `update_base64`, in pieces whose identity stands
 * for their content (see yDocStatePieces), for a store that keeps each piece
 * once (the local recovery copy, lib/page-draft-segments): null when this
 * module did not capture that very recovery object.
 */
export function pageCausalRecoveryPieces(recovery: unknown): readonly YDocStatePiece[] | null {
  return recovery && typeof recovery === "object" ? recoveryPieces.get(recovery as PageCausalRecovery) ?? null : null;
}

interface CapturedBeyond {
  /** The durable checkpoint `update` is measured from. */
  readonly durableSha256: string;
  /** That checkpoint as the room's binding verified it (its bytes against
   * its page), for the save of this capture to take as it is. */
  readonly durable: PreparedPageCausalCheckpoint;
  readonly update: Uint8Array;
  readonly stateVector: Uint8Array;
}

function recoveryDocument(
  recovery: PageCausalRecovery,
  content: Record<string, JsonValue>,
  pageId: string,
  epoch: number,
  model: PageCausalModel = "page-json-v1",
  /** The body object the caller was handed, when `content` is its copy. */
  source: object = content,
): Y.Doc {
  requireRecoveryBinding(recovery, pageId, epoch, model);
  const doc = new Y.Doc();
  try {
    const update = decodeBase64(recovery.update_base64);
    try { Y.applyUpdate(doc, update); }
    catch { invalid(); }
    requireComplete(doc);
    // Bytes this window encoded from a document it had just found to project
    // to this very body are not read back and compared a second time.
    if (!capturedFrom(source, recovery.update_base64) && !matchesProjection(doc, model, content)) {
      throw new PageCausalCheckpointError("projection_mismatch");
    }
    return doc;
  } catch (error) { doc.destroy(); throw error; }
}

function requireRecoveryBinding(recovery: PageCausalRecovery, pageId: string, epoch: number, model: PageCausalModel): void {
  if (recovery?.model !== model || recovery.page_id !== pageId ||
      !Number.isSafeInteger(epoch) || epoch < 1 || recovery.epoch !== epoch ||
      typeof recovery.base_title !== "string") {
    throw new PageCausalCheckpointError("identity_mismatch");
  }
}

/** Restore original IDs once, before diffing any new input. Never seed JSON
 * and then replay this snapshot: those would be two independent edits. */
export function applyPageCausalRecovery(
  doc: Y.Doc, recovery: PageCausalRecovery, content: Record<string, JsonValue>,
  pageId: string, epoch: number, origin: unknown, model: PageCausalModel = "page-json-v1",
): void {
  const restored = recoveryDocument(recovery, content, pageId, epoch, model);
  try {
    requireComplete(doc);
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(restored), origin);
    requireComplete(doc);
  } finally { restored.destroy(); }
}

/** The request fork and its non-CRDT save projection are captured BEFORE the
 * first await. Later input, layout passes and durable ACKs cannot enter it. */
export async function preparePageCausalSave(
  recovery: PageCausalRecovery,
  content: Record<string, JsonValue>,
  page: PageCausalCheckpointPage,
  receipt: PageCausalCheckpointReceipt,
  /** `nameBody: false`: the caller will send the body (its API refused a
   * named one); the digest is not computed. */
  options: { nameBody?: boolean } = {},
): Promise<{ content: Record<string, JsonValue>; checkpoint: PageCausalCheckpoint; epoch: number;
  delta: PageCausalSaveDelta | null;
  /** This save's body named by digest (page-causal-named-body.ts), or null. */
  named: PageCausalNamedBody | null }> {
  // A document save's body is already its own copy (documentSaveCopy), and
  // the durable page's body is a read-only server body: copying either
  // whole (twice a save, megabytes on an inked page) froze nothing more.
  const pageModel = page.model ?? "page-json-v1";
  const documentModel = pageModel === "document-canonical-v1";
  // A deck's body read from its room is read-only (json-crdt
  // `yToJsonShared`): a copy of its top level freezes it as much as a whole
  // copy did (three megabytes of a 185-row table, every save).
  const frozenContent = documentModel && isDocumentSaveCopy(content) ? content
    : pageModel === "page-json-v1" && sharedJsonBody(content) ? { ...content } : structuredClone(content);
  if (frozenContent !== content && pageModel === "page-json-v1") carryProvenProjection(content, frozenContent);
  const frozenPage = documentModel || pageModel === "page-json-v1" ? { ...page } : structuredClone(page);
  const frozenReceipt = structuredClone(receipt);
  carryProvenProjection(page.content_json, frozenPage.content_json);
  // A document body captured from the live room a moment ago, with the room
  // already holding the durable checkpoint: the request is those very bytes.
  // Decoding them into a second document to read back what the capture had
  // already proven (its projection, its extent, what it holds beyond the
  // checkpoint) was most of a long page's save, on the editor's thread.
  const known = capturedFrom(content, recovery?.update_base64);
  if (known?.update && known.projection && known.beyond && (documentModel || pageModel === "page-json-v1")) {
    requireRecoveryBinding(recovery, frozenPage.id, frozenReceipt.epoch, pageModel);
    // Everything is frozen above: the capture and the stored copy that went
    // before this in the same task are not followed by the rest of the save.
    await yieldBetweenSaveSteps();
    // The checkpoint the room is bound to is this save's own base when the
    // receipt and the page name it exactly: verified when the room bound it
    // (its bytes against its page), it is not decoded and read whole again
    // (a deck's 185-row table: half a second of every save).
    const bound = known.beyond.durable;
    const prepared = bound.model === pageModel && bound.pageId === frozenPage.id && frozenReceipt.page_id === frozenPage.id &&
      bound.epoch === frozenReceipt.epoch &&
      bound.revision === frozenReceipt.revision && bound.contentVersion === frozenPage.content_version &&
      bound.contentVersion === frozenReceipt.content_version && bound.contentSha256 === frozenPage.content_sha256 &&
      bound.contentSha256 === frozenReceipt.content_sha256 && bound.snapshot.sha256 === frozenReceipt.checkpoint?.update_sha256 &&
      frozenReceipt.checkpoint?.model === pageModel
      ? bound : await preparePageCausalCheckpoint(frozenReceipt, frozenPage);
    if (prepared.model === pageModel && known.beyond.durableSha256 === prepared.snapshot.sha256 &&
        holdsAll(known.extent, extentOfPrepared(prepared))) {
      return capturedSave(known as CapturedState & { update: Uint8Array; projection: Record<string, JsonValue>; beyond: CapturedBeyond },
        recovery, prepared, frozenContent, frozenPage, options);
    }
    return requestSave(recovery, content, frozenContent, frozenPage, frozenReceipt, options, prepared);
  }
  return requestSave(recovery, content, frozenContent, frozenPage, frozenReceipt, options);
}

/** The save of a body whose bytes were captured from the live room, which
 * then held everything the durable checkpoint does: the checkpoint is the
 * captured bytes, its projection the one the capture proved the body says,
 * its delta what the capture took beyond that checkpoint. */
async function capturedSave(
  known: CapturedState & { update: Uint8Array; projection: Record<string, JsonValue>; beyond: CapturedBeyond },
  recovery: PageCausalRecovery,
  prepared: PreparedPageCausalCheckpoint,
  frozenContent: Record<string, JsonValue>,
  frozenPage: PageCausalCheckpointPage,
  options: { nameBody?: boolean },
): Promise<PreparedPageCausalSave> {
  const checkpoint = Object.freeze({ model: prepared.model, update_base64: recovery.update_base64,
    update_sha256: await sha256(known.update) });
  rememberExtent(checkpoint.update_sha256, known.extent);
  rememberOwnBytes(checkpoint, known.update);
  captured.set(frozenContent, { ...known, updateBase64: checkpoint.update_base64 });
  let delta: PageCausalSaveDelta | null = null;
  const wholeBytes = Math.floor(checkpoint.update_base64.length / 4) * 3;
  if (wholeBytes >= PAGE_CAUSAL_DELTA_MIN_WHOLE_BYTES && known.beyond.update.length < wholeBytes * 0.75) {
    delta = Object.freeze({
      checkpoint: Object.freeze({ model: prepared.model, update_base64: encodeBase64(known.beyond.update),
        update_sha256: await sha256(known.beyond.update) }),
      state_vector_base64: encodeBase64(known.beyond.stateVector),
    });
  }
  if (options.nameBody !== false) await yieldBetweenSaveSteps();
  const named = options.nameBody === false ? null
    : await namePageCausalBody(prepared.model, known.projection, frozenContent, frozenPage, { mirrors: true });
  return { content: frozenContent, checkpoint, epoch: prepared.epoch, delta, named };
}

type PreparedPageCausalSave = { content: Record<string, JsonValue>; checkpoint: PageCausalCheckpoint; epoch: number;
  delta: PageCausalSaveDelta | null; named: PageCausalNamedBody | null };

/** The save made from a request document decoded from the recovery bytes. */
async function requestSave(
  recovery: PageCausalRecovery,
  content: Record<string, JsonValue>,
  frozenContent: Record<string, JsonValue>,
  frozenPage: PageCausalCheckpointPage,
  frozenReceipt: PageCausalCheckpointReceipt,
  options: { nameBody?: boolean },
  alreadyPrepared?: PreparedPageCausalCheckpoint,
): Promise<PreparedPageCausalSave> {
  const request = recoveryDocument(recovery, frozenContent, frozenPage.id, frozenReceipt.epoch, frozenPage.model ?? "page-json-v1", content);
  try {
    const prepared = alreadyPrepared ?? await preparePageCausalCheckpoint(frozenReceipt, frozenPage);
    // The request document as it was captured for this very body, with
    // nothing of the durable checkpoint left to take in: it still projects to
    // what it was captured for.
    const untouched = Boolean(capturedFrom(content, recovery.update_base64)) &&
      holdsAll(extentOf(request), extentOfPrepared(prepared));
    let merged = prepared.apply(request, "durable-save-basis");
    // What the request document projects to, before any non-shared part of
    // the body is put back: the content both sides digest.
    const projection = merged;
    if (prepared.model === "pdf-author-v1") merged = withPdfDurableJournal(merged, frozenPage.content_json);
    // The authored body is this call's own copy: kept as it is when it says
    // what the request document does (and so is that document's projection).
    const document = prepared.model === "document-canonical-v1";
    if (document) merged = untouched ? frozenContent : retainDocumentCausalSaveProjection(merged, frozenContent, { owned: true });
    // What this request holds beyond the durable checkpoint, taken at the
    // same moment as the whole checkpoint below (no await in between).
    const beyond = prepared.beyond(request);
    const checkpoint = await captureCheckpoint(request, merged, prepared.model, document ? "proven" : projection);
    captured.set(merged, { updateBase64: checkpoint.update_base64, extent: extents.get(checkpoint.update_sha256) ?? extentOf(request) });
    let delta: PageCausalSaveDelta | null = null;
    // Worth sending only for a state large enough to matter, and only when
    // it is clearly the smaller body: a short page keeps the whole one.
    const wholeBytes = Math.floor(checkpoint.update_base64.length / 4) * 3;
    if (beyond && wholeBytes >= PAGE_CAUSAL_DELTA_MIN_WHOLE_BYTES && beyond.update.length < wholeBytes * 0.75) {
      delta = Object.freeze({
        checkpoint: Object.freeze({ model: prepared.model, update_base64: encodeBase64(beyond.update),
          update_sha256: await sha256(beyond.update) }),
        state_vector_base64: encodeBase64(beyond.stateVector),
      });
    }
    const named = options.nameBody === false ? null
      : await namePageCausalBody(prepared.model, projection, frozenContent, frozenPage,
        document ? { mirrors: merged === frozenContent } : {});
    return { content: merged, checkpoint, epoch: prepared.epoch, delta, named };
  } finally { request.destroy(); }
}

/** The same save as its whole checkpoint, as the part the durable checkpoint
 * does not hold: the structs past that checkpoint's state vector, this
 * author's complete delete set, and this author's state vector. A long
 * document's save carried its whole Yjs state every time (1.75 MB of a
 * 3.4 MB request); the server rebuilds it from the checkpoint it stores, or
 * answers `page_causal_delta_base_stale` and is sent the whole one. */
export interface PageCausalSaveDelta {
  readonly checkpoint: PageCausalCheckpoint;
  readonly state_vector_base64: string;
}

/** Separate native commands from pending notes. The caller must still keep
 * those notes in recovery and obtain the existing native byte/history proof.
 * This private model is not selected by public page-type routing yet. */
export async function preparePdfNativeCausalSave(
  page: PageCausalCheckpointPage, receipt: PageCausalCheckpointReceipt,
  requested: Record<string, JsonValue>, operationId: string,
): Promise<{ content: Record<string, JsonValue>; checkpoint: PageCausalCheckpoint; epoch: number; nativeOperationId: string }> {
  if (page.model !== "pdf-author-v1" || receipt.checkpoint.model !== "pdf-author-v1") {
    throw new PageCausalCheckpointError("identity_mismatch");
  }
  const frozenPage = structuredClone(page), frozenReceipt = structuredClone(receipt);
  const content = appendPdfCausalNativeCommand(frozenPage.content_json, requested, operationId);
  const prepared = await preparePageCausalCheckpoint(frozenReceipt, frozenPage);
  // The original checkpoint is already durable. A private native operation
  // does not author or acknowledge any new shared-note structs/deletes.
  return { content, checkpoint: frozenReceipt.checkpoint, epoch: prepared.epoch, nativeOperationId: operationId };
}

/** Use only for the direct response to this caller's exact PATCH. Prove that
 * the returned chain already contains every submitted struct AND deletion.
 * JSON equality or a state vector alone cannot establish that receipt. */
export async function preparePageCausalSaveReceipt(
  receipt: PageCausalCheckpointReceipt,
  page: PageCausalCheckpointPage,
  submitted: { recovery: PageCausalRecovery; content: Record<string, JsonValue> },
): Promise<PreparedPageCausalCheckpoint> {
  const model = page.model ?? "page-json-v1";
  // What was submitted, as the clocks and deletions it holds: known from the
  // save that encoded these bytes for this body, read from the bytes otherwise.
  let held = capturedFrom(submitted.content, submitted.recovery?.update_base64)?.extent ?? null;
  if (held) requireRecoveryBinding(submitted.recovery, page.id, receipt.epoch, model);
  else {
    const request = recoveryDocument(submitted.recovery, submitted.content, page.id, receipt.epoch, model);
    try { held = extentOf(request); } finally { request.destroy(); }
  }
  const prepared = await preparePageCausalCheckpoint(receipt, page);
  if (!holdsAll(extentOfPrepared(prepared), held)) throw new PageCausalCheckpointError("identity_mismatch");
  return prepared;
}

/**
 * The checkpoint a compact save receipt stands for: this author's submitted
 * state with what the stored checkpoint holds beyond it applied
 * (`delta_base64`: someone else's merged change, the block ids the server
 * gave new paragraphs). The receipt carries only that difference instead of
 * the whole checkpoint, every save.
 *
 * The result is this author's own encoding of that state, with its own
 * digest, and what it projects to: the caller names the stored body from the
 * projection and checks it against the receipt's `content_sha256` (the
 * server never said these bytes; the body digest is what binds them).
 */
export async function rebuildPageCausalReceipt(
  submitted: PageCausalCheckpoint,
  deltaBase64: string,
): Promise<{ checkpoint: PageCausalCheckpoint; projection: Record<string, JsonValue> }> {
  const doc = new Y.Doc();
  try {
    const own = decodeBase64(submitted.update_base64), delta = decodeBase64(deltaBase64);
    try { Y.applyUpdate(doc, own); Y.applyUpdate(doc, delta); }
    catch { invalid(); }
    requireComplete(doc);
    const projection = readProjection(doc, submitted.model);
    const update = Y.encodeStateAsUpdate(doc);
    if (!update.length || update.length > MAX_UPDATE_BYTES) invalid();
    const extent = extentOf(doc);
    // Decoded again, these bytes are this same complete state.
    const checkpoint = Object.freeze({ model: submitted.model,
      update_base64: encodeBase64(update), update_sha256: await sha256(update) });
    rememberExtent(checkpoint.update_sha256, extent);
    rememberOwnBytes(checkpoint, update);
    return { projection, checkpoint };
  } finally { doc.destroy(); }
}

/**
 * Bodies this module derived from a checkpoint itself: the body object ->
 * the digest of the update it is the projection of. Preparing that very
 * checkpoint for that very body has nothing left to prove, so it does not
 * decode the update and read its whole projection a second time (what a
 * window did on every open, every save receipt and every read of a peer's
 * save: seconds on a long document). The proof is of an object: a copy of
 * the body carries it only when its maker says so (`carryProvenProjection`).
 */
const provenProjections = new WeakMap<object, string>();

/** `body` is what the update with this digest projects to (plus whatever the
 * model keeps outside the shared state): the caller just derived it. */
export function markProvenProjection(body: object, updateSha256: string): void {
  provenProjections.set(body, updateSha256);
}

/** A faithful copy of a body (structuredClone) is the same projection, and
 * the same body a recovery's bytes were captured for (`captured`). */
export function carryProvenProjection(from: object | null | undefined, to: object | null | undefined): void {
  if (!from || !to) return;
  const proven = provenProjections.get(from);
  if (proven) provenProjections.set(to, proven);
  const bytes = captured.get(from);
  if (bytes) captured.set(to, bytes);
}

/**
 * What a Yjs state holds, as sets: the clock each client has reached and the
 * ranges each client's items were deleted in. One state holds everything
 * another does exactly when it has reached every clock the other has and
 * deleted every range the other has. A save asked that question by merging
 * one whole state into the other and comparing two whole encodings byte for
 * byte (on a long document: seconds of every save, on the editor's thread).
 */
export interface CausalExtent {
  readonly clocks: ReadonlyMap<number, number>;
  /** Per client: [from, to) ranges, ascending, none touching the next. */
  readonly deleted: ReadonlyMap<number, ReadonlyArray<readonly [number, number]>>;
}

function extentOf(doc: Y.Doc): CausalExtent {
  const deleted = new Map<number, Array<[number, number]>>();
  for (const [client, items] of Y.createDeleteSetFromStructStore(doc.store).clients) {
    const ranges: Array<[number, number]> = [];
    for (const item of items) {
      const last = ranges.at(-1);
      if (last && item.clock <= last[1]) last[1] = Math.max(last[1], item.clock + item.len);
      else ranges.push([item.clock, item.clock + item.len]);
    }
    if (ranges.length) deleted.set(client, ranges);
  }
  return { clocks: Y.decodeStateVector(Y.encodeStateVector(doc)), deleted };
}

/** Whether `outer` holds every struct and every deletion `inner` does. */
function holdsAll(outer: CausalExtent, inner: CausalExtent): boolean {
  for (const [client, clock] of inner.clocks) if ((outer.clocks.get(client) ?? 0) < clock) return false;
  for (const [client, ranges] of inner.deleted) {
    const have = outer.deleted.get(client);
    if (!have) return false;
    let index = 0;
    for (const [from, to] of ranges) {
      while (index < have.length && have[index][1] <= from) index += 1;
      if (index === have.length || have[index][0] > from || have[index][1] < to) return false;
    }
  }
  return true;
}

/** The extent of an update, by the digest of its bytes: noted by whoever had
 * the document those bytes are the state of in hand (the one that encoded
 * them, or the one that decoded them to verify them). */
const extents = new Map<string, CausalExtent>();
const EXTENTS_KEPT = 32;
function rememberExtent(updateSha256: string, extent: CausalExtent): CausalExtent {
  extents.delete(updateSha256);
  extents.set(updateSha256, extent);
  if (extents.size > EXTENTS_KEPT) extents.delete(extents.keys().next().value as string);
  return extent;
}

/** The last few checkpoints this module encoded (or rebuilt) and digested
 * itself: digest -> their base64 and bytes. Never written into. */
const ownCheckpointBytes = new Map<string, { readonly base64: string; readonly update: Uint8Array }>();
const OWN_CHECKPOINTS_KEPT = 3;
function rememberOwnBytes(checkpoint: PageCausalCheckpoint, update: Uint8Array): void {
  // A Docs page and a deck (whose saves are the long ones under the pen and
  // in a large table): other models keep verifying a receipt as they always
  // did.
  if (checkpoint.model !== "document-canonical-v1" && checkpoint.model !== "page-json-v1") return;
  ownCheckpointBytes.delete(checkpoint.update_sha256);
  ownCheckpointBytes.set(checkpoint.update_sha256, { base64: checkpoint.update_base64, update });
  if (ownCheckpointBytes.size > OWN_CHECKPOINTS_KEPT) ownCheckpointBytes.delete(ownCheckpointBytes.keys().next().value as string);
}

/** A body object -> the bytes this module encoded from a document it had
 * found (or was told, `projectionProven`) to project to that very object,
 * and that document's extent. */
interface CapturedState {
  readonly updateBase64: string;
  readonly extent: CausalExtent;
  /** The bytes themselves, the projection they were found (or proven) to
   * have, and what they hold beyond the durable checkpoint the document was
   * bound to: kept by a capture from a live document, for its save. */
  readonly update?: Uint8Array;
  readonly projection?: Record<string, JsonValue> | null;
  readonly beyond?: CapturedBeyond | null;
}
const captured = new WeakMap<object, CapturedState>();
function capturedFrom(content: object | null | undefined, updateBase64: string | undefined) {
  const known = content ? captured.get(content) : undefined;
  return known && known.updateBase64 === updateBase64 ? known : null;
}

/**
 * The state vector of a checkpoint's update, for asking a server what it
 * holds beyond it: from what was noted of those bytes when they were decoded
 * or encoded here, read from them otherwise (a pass over a long document's
 * whole state, every time a member of its room was told of a save).
 */
export function pageCausalCheckpointStateVector(checkpoint: PageCausalCheckpoint): Uint8Array {
  const known = extents.get(checkpoint.update_sha256);
  return known
    ? Y.encodeStateVector(known.clocks as Map<number, number>)
    : Y.encodeStateVectorFromUpdate(decodeBase64(checkpoint.update_base64));
}

const preparedExtents = new WeakMap<object, () => CausalExtent>();
function extentOfPrepared(prepared: PreparedPageCausalCheckpoint): CausalExtent {
  const extent = preparedExtents.get(prepared);
  if (!extent) return invalid();
  return extent();
}

/**
 * What a checkpoint projects to, from its own bytes: the body of a page read
 * that sent the checkpoint and no body (`derive_body`). The bytes are proven
 * to be the checkpoint's (digest) and a complete update.
 */
export async function projectPageCausalCheckpoint(
  checkpoint: PageCausalCheckpoint,
): Promise<Record<string, JsonValue>> {
  const { projection, extent } = await projectPageCausalCheckpointAlone(checkpoint);
  rememberExtent(checkpoint.update_sha256, extent);
  return projection;
}

/**
 * The same, leaving this module's memory alone: for a worker, whose module
 * is not the page's. The page keeps the extent the worker sends back
 * (`rememberPageCausalExtent`).
 */
/** `checkpoint` with `updatesBase64` (a deck's room custody beyond it, in
 * order: `room_custody_updates_base64`) merged in. The checkpoint's own
 * bytes are proven to be its digest first; the result names its own. */
export async function checkpointWithUpdates(
  checkpoint: PageCausalCheckpoint, updatesBase64: readonly string[],
): Promise<PageCausalCheckpoint> {
  if (!DIGEST_PATTERN.test(checkpoint?.update_sha256 ?? "")) invalid();
  const update = decodeBase64(checkpoint.update_base64);
  if (await sha256(update) !== checkpoint.update_sha256) invalid();
  let merged: Uint8Array;
  try { merged = Y.mergeUpdates([update, ...updatesBase64.map(decodeBase64)]); }
  catch { invalid(); }
  return { ...checkpoint, update_base64: encodeBase64(merged!), update_sha256: await sha256(merged!) };
}

export async function projectPageCausalCheckpointAlone(
  checkpoint: PageCausalCheckpoint,
): Promise<{ projection: Record<string, JsonValue>; extent: CausalExtent }> {
  if (!DIGEST_PATTERN.test(checkpoint?.update_sha256 ?? "")) invalid();
  const update = decodeBase64(checkpoint.update_base64);
  if (await sha256(update) !== checkpoint.update_sha256) invalid();
  const doc = new Y.Doc();
  try {
    try { Y.applyUpdate(doc, update); }
    catch { invalid(); }
    requireComplete(doc);
    return { projection: readProjection(doc, checkpoint.model), extent: extentOf(doc) };
  } finally { doc.destroy(); }
}

/** Keep the extent a worker read for a checkpoint, as projecting it here would. */
export function rememberPageCausalExtent(updateSha256: string, extent: CausalExtent): void {
  rememberExtent(updateSha256, extent);
}

const ROOT_KEY = "page-json";
const MAX_UPDATE_BYTES = 16 * 1024 * 1024;
/** Below this size of the whole state a save sends it as it always did. */
export const PAGE_CAUSAL_DELTA_MIN_WHOLE_BYTES = 16 * 1024;
const DIGEST_PATTERN = /^[0-9a-f]{64}$/u;

export class PageCausalCheckpointError extends Error {
  constructor(readonly code: "invalid" | "incomplete" | "projection_mismatch" | "identity_mismatch") {
    super(`Page causal checkpoint: ${code}`);
    this.name = "PageCausalCheckpointError";
  }
}

function invalid(): never { throw new PageCausalCheckpointError("invalid"); }

function requireComplete(doc: Y.Doc): void {
  // A syntactically valid update may still depend on absent structs, including
  // a delete set whose referenced items have not arrived. Never store that as
  // a full standalone checkpoint and never project its partial visible JSON.
  if (doc.store.pendingStructs || doc.store.pendingDs) {
    throw new PageCausalCheckpointError("incomplete");
  }
}

/** `Y.encodeStateAsUpdate` of what the document integrated, without the
 * updates it holds pending their predecessors (Yjs merges those into its
 * encoding). Synchronous: nothing runs while they are set aside, and they
 * are put back exactly as they were whether the encoding returns or throws.
 * It reads the store's own fields (Yjs 13.6): a Yjs that names them otherwise
 * throws here rather than encoding the pending updates after all (the read
 * that calls it then gives nothing, as before this was written, and the
 * tests in tests/unit/page-causal-checkpoint.test.ts turn red). */
export function encodeIntegratedStateAsUpdate(doc: Y.Doc, targetStateVector?: Uint8Array): Uint8Array {
  const store = doc.store;
  if (!("pendingStructs" in store) || !("pendingDs" in store)) {
    throw new Error("Yjs StructStore no longer holds pendingStructs/pendingDs");
  }
  const { pendingStructs, pendingDs } = store;
  if (!pendingStructs && !pendingDs) return Y.encodeStateAsUpdate(doc, targetStateVector);
  store.pendingStructs = null;
  store.pendingDs = null;
  try {
    return Y.encodeStateAsUpdate(doc, targetStateVector);
  } finally {
    store.pendingStructs = pendingStructs;
    store.pendingDs = pendingDs;
  }
}

function readProjection(doc: Y.Doc, model: PageCausalModel): Record<string, JsonValue> {
  if (model === "document-canonical-v1") {
    try { return readDocumentCausalProjection(doc); } catch { return invalid(); }
  }
  if (model !== "page-json-v1" && model !== "code-text-v1" && model !== "pdf-author-v1") invalid();
  const roots = model === "code-text-v1" ? [ROOT_KEY, CODE_COLLAB_TEXT_KEY] : [ROOT_KEY];
  if ([...doc.share.keys()].some(key => !roots.includes(key))) invalid();
  const root = doc.getMap(ROOT_KEY);
  const metadata = yToJson(root) as Record<string, JsonValue>;
  if (model === "pdf-author-v1") {
    if (root.has("pdfNativeTransactions")) invalid();
    return metadata;
  }
  if (model === "page-json-v1") return metadata;
  if (root.has("content")) invalid();
  const text = doc.getText(CODE_COLLAB_TEXT_KEY);
  if (text.toDelta().some((run: { insert?: unknown; attributes?: Record<string, unknown> }) => typeof run.insert !== "string" ||
      (run.attributes && Object.keys(run.attributes).length > 0))) invalid();
  return { ...metadata, content: text.toString() };
}

/** What readProjection checks of a document besides reading it: only its
 * own roots, and (a PDF's) no native journal in the shared state. */
function requireProjectableShape(doc: Y.Doc, model: PageCausalModel): void {
  if (model !== "page-json-v1" && model !== "code-text-v1" && model !== "pdf-author-v1") invalid();
  const roots = model === "code-text-v1" ? [ROOT_KEY, CODE_COLLAB_TEXT_KEY] : [ROOT_KEY];
  if ([...doc.share.keys()].some(key => !roots.includes(key))) invalid();
  if (model === "pdf-author-v1" && doc.getMap(ROOT_KEY).has("pdfNativeTransactions")) invalid();
}

/** readProjection of a live document, through its shared read (json-crdt
 * `yToJsonShared`): the cost of what changed since it was last read. */
function readProjectionShared(doc: Y.Doc, model: PageCausalModel): void {
  if (model === "code-text-v1") { readProjection(doc, model); return; }
  requireProjectableShape(doc, model);
  yToJsonShared(doc.getMap(ROOT_KEY));
}

/** A body every value of which is a shared read of a room, or a plain value:
 * read-only below its top level. */
function sharedJsonBody(content: Record<string, JsonValue>): boolean {
  if (isSharedJsonRead(content)) return true;
  for (const key in content) {
    const value = content[key];
    if (value !== null && typeof value === "object" && !isSharedJsonRead(value)) return false;
  }
  return true;
}

function matchesProjection(doc: Y.Doc, model: PageCausalModel, content: Record<string, JsonValue>): boolean {
  return matchesRead(readProjection(doc, model), model, content);
}

function matchesRead(projection: Record<string, JsonValue>, model: PageCausalModel, content: Record<string, JsonValue>): boolean {
  if (model === "document-canonical-v1") return documentContentSaysProjection(content, projection);
  return jsonEquals(projection, model === "pdf-author-v1" ? pdfSharedContent(content) : content);
}

// The platform's own base64 for bytes where it has one: building a binary
// string 8 KiB at a time for btoa took 40 ms of a long document's 2.4 MB
// checkpoint, twice a save; the native encoder takes under 1 ms.
const nativeBase64 = typeof (Uint8Array.prototype as { toBase64?: unknown }).toBase64 === "function" &&
  typeof (Uint8Array as { fromBase64?: unknown }).fromBase64 === "function"
  ? {
      encode: (bytes: Uint8Array) => (bytes as Uint8Array & { toBase64(): string }).toBase64(),
      decode: (value: string) => (Uint8Array as unknown as {
        fromBase64(value: string, options: { lastChunkHandling: "strict" }): Uint8Array;
      }).fromBase64(value, { lastChunkHandling: "strict" }),
    }
  : null;

function encodeBase64(bytes: Uint8Array): string {
  if (nativeBase64) return nativeBase64.encode(bytes);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return btoa(binary);
}

function decodeBase64(value: string): Uint8Array {
  if (typeof value !== "string" || !value.length ||
      value.length > Math.ceil(MAX_UPDATE_BYTES / 3) * 4) invalid();
  if (nativeBase64) {
    let bytes: Uint8Array;
    try { bytes = nativeBase64.decode(value); } catch { return invalid(); }
    // Canonical form: the platform decoder also skips whitespace.
    if (bytes.length > MAX_UPDATE_BYTES || nativeBase64.encode(bytes) !== value) invalid();
    return bytes;
  }
  try {
    const binary = atob(value);
    if (binary.length > MAX_UPDATE_BYTES) invalid();
    // Canonical form, proven on the decoded string itself: re-encoding the
    // bytes (a spread per 8 KiB and a callback per byte) cost as much as the
    // decode on a long document's receipt, every save.
    if (btoa(binary) !== value) invalid();
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  } catch { return invalid(); }
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest("SHA-256", Uint8Array.from(bytes).buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** Capture synchronously before hashing yields. Call after flushing the local
 * JSON delta; the supplied projection must be the exact body being saved.
 * This function provides bytes, never a durable save acknowledgement. */
export function capturePageCausalCheckpoint(
  doc: Y.Doc,
  content: Record<string, JsonValue>,
  model: PageCausalModel = "page-json-v1",
): Promise<PageCausalCheckpoint> {
  return captureCheckpoint(doc, content, model, null);
}

/** `projection`: what the caller has just read `doc` to project to (nothing
 * changed it since), or "proven" when it has also found `content` to say
 * that; null reads and compares here. */
async function captureCheckpoint(
  doc: Y.Doc,
  content: Record<string, JsonValue>,
  model: PageCausalModel,
  projection: Record<string, JsonValue> | "proven" | null,
): Promise<PageCausalCheckpoint> {
  requireComplete(doc);
  if (projection !== "proven" && !(projection ? matchesRead(projection, model, content) : matchesProjection(doc, model, content))) {
    throw new PageCausalCheckpointError("projection_mismatch");
  }
  const update = Y.encodeStateAsUpdate(doc);
  if (!update.length || update.length > MAX_UPDATE_BYTES) invalid();
  const extent = extentOf(doc);
  const checkpoint = Object.freeze({
    model,
    update_base64: encodeBase64(update),
    update_sha256: await sha256(update),
  });
  rememberExtent(checkpoint.update_sha256, extent);
  rememberOwnBytes(checkpoint, update);
  return checkpoint;
}

export interface PreparedPageCausalCheckpoint {
  readonly model: PageCausalModel;
  readonly pageId: string;
  readonly epoch: number;
  readonly revision: number;
  readonly contentVersion: number;
  readonly contentSha256: string;
  /** Full original receipt, including deletes; independent of visible JSON. */
  covers: (doc: Y.Doc) => boolean;
  /** Apply the original causal operations with a remote origin. The caller
   * must recheck its actor, activation, authority and durable baseline after
   * awaiting preparation. This does not advance a REST base or clear drafts. */
  apply: (doc: Y.Doc, remoteOrigin: unknown) => Record<string, JsonValue>;
  /** `apply` for a caller that does not read what the document then
   * projects to (reading it is a pass over the whole document). */
  merge: (doc: Y.Doc, remoteOrigin: unknown) => void;
  /** The checkpoint's own verified bytes and their digest: what the server
   * installs as an empty room's first snapshot (see offerDurablePageSnapshot). */
  readonly snapshot: { readonly sha256: string; readonly update: Uint8Array };
  /** What `doc` holds beyond this checkpoint (see PageCausalSaveDelta), or
   * null when it cannot be told apart from it. `doc` must already hold the
   * checkpoint (`apply`). */
  beyond: (doc: Y.Doc) => { update: Uint8Array; stateVector: Uint8Array } | null;
}

/** Decode and verify away from the live document. The independently fetched
 * page is the authority; a hash or a claimed revision in the envelope alone
 * is insufficient. Changed/canonicalized server bodies require another exact
 * checkpoint, never JSON reconstruction followed by replay of old structs. */
/** What a window tells its room transport about the checkpoint it is still
 * verifying (CollabNetworkProvider.offerDurablePageSnapshot): the digest the
 * server named, and the bytes once they are proven to be that digest's. */
export function durablePageSnapshotOffer(
  receipt: PageCausalCheckpointReceipt,
  preparing: Promise<PreparedPageCausalCheckpoint>,
): { readonly sha256: string; readonly update: Promise<Uint8Array | null> } {
  const sha256 = receipt.checkpoint.update_sha256;
  return {
    sha256,
    update: preparing.then((prepared) => prepared.snapshot.sha256 === sha256 ? prepared.snapshot.update : null, () => null),
  };
}

export async function preparePageCausalCheckpoint(
  receipt: PageCausalCheckpointReceipt,
  page: PageCausalCheckpointPage,
): Promise<PreparedPageCausalCheckpoint> {
  if (!receipt || !page || !page.id || receipt.page_id !== page.id ||
      !Number.isSafeInteger(receipt.epoch) || receipt.epoch < 1 ||
      !Number.isSafeInteger(receipt.revision) || receipt.revision < 1 ||
      !Number.isSafeInteger(page.content_version) || page.content_version < 1 ||
      receipt.content_version !== page.content_version ||
      !DIGEST_PATTERN.test(page.content_sha256) || receipt.content_sha256 !== page.content_sha256) {
    throw new PageCausalCheckpointError("identity_mismatch");
  }
  const { id: pageId, content_version: contentVersion, content_sha256: contentSha256 } = page;
  const { epoch, revision } = receipt;
  const checkpoint = receipt.checkpoint;
  const model = page.model ?? "page-json-v1";
  if (checkpoint?.model !== model) throw new PageCausalCheckpointError("identity_mismatch");
  if (!DIGEST_PATTERN.test(checkpoint.update_sha256)) invalid();
  const expectedDigest = checkpoint.update_sha256;
  // A body this module derived from these very bytes is their projection.
  // Any other is frozen here, before the first await, to be compared.
  const expected = provenProjections.get(page.content_json) === expectedDigest ? null : structuredClone(page.content_json);
  // Bytes this window encoded and digested itself, sent back as they went
  // (a save's own receipt, every save): the same text is the same bytes,
  // and their digest is already known. Decoding and digesting them again was
  // tens of milliseconds of each receipt on an inked page.
  const own = ownCheckpointBytes.get(expectedDigest);
  const known = own !== undefined && own.base64 === checkpoint.update_base64;
  const update = known ? own.update : decodeBase64(checkpoint.update_base64);
  if (!known && await sha256(update) !== expectedDigest) invalid();
  let held: CausalExtent | null = null;
  if (expected) {
    const isolated = new Y.Doc();
    try {
      // Digest verification proves custody, not valid Yjs structure.
      try { Y.applyUpdate(isolated, update); }
      catch { invalid(); }
      requireComplete(isolated);
      if (!matchesProjection(isolated, model, expected)) {
        throw new PageCausalCheckpointError("projection_mismatch");
      }
      held = rememberExtent(expectedDigest, extentOf(isolated));
    } finally { isolated.destroy(); }
  }
  // What these bytes hold: noted when they were decoded or encoded here, read
  // from them (once) by a window that was never shown their document.
  const extent = (): CausalExtent => {
    held ??= extents.get(expectedDigest) ?? null;
    if (held) return held;
    const decoded = new Y.Doc();
    try {
      try { Y.applyUpdate(decoded, update); }
      catch { invalid(); }
      requireComplete(decoded);
      return held = rememberExtent(expectedDigest, extentOf(decoded));
    } finally { decoded.destroy(); }
  };
  const merge = (doc: Y.Doc, remoteOrigin: unknown): void => {
    requireComplete(doc);
    const present = extentOf(doc);
    // The document holds all of it already (this author's own save, read
    // back; a durable source it was opened from): nothing to apply. It must
    // still be a document that projects, as a staged merge would have found.
    if (holdsAll(present, extent())) {
      // Read as the room's own shared read: an unchanged document is not
      // read whole again (a deck's live room, at every save receipt).
      if (model !== "document-canonical-v1") readProjectionShared(doc, model);
      else try { readSharedDocumentCausalProjection(doc); } catch { invalid(); }
      return;
    }
    if (present.clocks.size > 0) {
      // Stage against the current document too. No partial/corrupt projection
      // is allowed to mutate the live editor before validation completes.
      const merged = new Y.Doc();
      try {
        Y.applyUpdate(merged, Y.encodeStateAsUpdate(doc));
        Y.applyUpdate(merged, update);
        requireComplete(merged);
        readProjection(merged, model);
      } finally { merged.destroy(); }
    }
    // An empty document becomes exactly these bytes, which are a complete
    // state that projects (verified above, or by whoever derived the body).
    Y.applyUpdate(doc, update, remoteOrigin);
  };
  const prepared: PreparedPageCausalCheckpoint = Object.freeze({
    model, pageId, epoch, revision, contentVersion, contentSha256,
    snapshot: Object.freeze({ sha256: expectedDigest, update }),
    covers: (doc: Y.Doc) => {
      try {
        requireComplete(doc);
        return holdsAll(extent(), extentOf(doc));
      } catch { return false; }
    },
    merge,
    apply: (doc: Y.Doc, remoteOrigin: unknown) => {
      merge(doc, remoteOrigin);
      return readProjection(doc, model);
    },
    // What the document holds beyond the checkpoint, as one update for the
    // server to check against it (closing edits, the room's custody). Only
    // what the document integrated: a peer's update still waiting here for
    // its predecessors (a relay gap, the very state the room's gap resume
    // reconnects from) is the peer's, and the server already keeps whatever
    // of it reached the room. Refusing the whole read for it left this
    // window's own edits with nothing to carry them: a window closed during
    
    beyond: (doc: Y.Doc) => {
      try {
        const known = extent().clocks;
        const stateVector = Y.encodeStateVector(doc);
        const present = Y.decodeStateVector(stateVector);
        // The checkpoint was applied to this document: it holds every clock.
        for (const [client, clock] of known) if ((present.get(client) ?? 0) < clock) return null;
        return { update: encodeIntegratedStateAsUpdate(doc, Y.encodeStateVector(known as Map<number, number>)),
          stateVector };
      } catch { return null; }
    },
  });
  preparedExtents.set(prepared, extent);
  return prepared;
}

/**
 * Whether a page's saved checkpoint already holds everything a recovered
 * causal draft would replay: applying the draft's update to the checkpoint
 * changes nothing (no insert it lacks, no deletion it has not made). A draft
 * whose body only matches the page is not enough -- it may still carry
 * tombstones the page must receive -- but one the checkpoint holds whole is
 * spent. False whenever either cannot be read.
 */
export function pageCausalRecoveryHeldBy(
  recovery: PageCausalRecovery,
  checkpoint: PageCausalCheckpointReceipt | null | undefined,
): boolean {
  if (!checkpoint || checkpoint.page_id !== recovery.page_id || checkpoint.epoch !== recovery.epoch) return false;
  const doc = new Y.Doc();
  try {
    Y.applyUpdate(doc, decodeBase64(checkpoint.checkpoint.update_base64));
    const before = Y.encodeStateAsUpdate(doc);
    Y.applyUpdate(doc, decodeBase64(recovery.update_base64));
    const after = Y.encodeStateAsUpdate(doc);
    return before.length === after.length && before.every((byte, index) => byte === after[index]);
  } catch {
    return false;
  } finally {
    doc.destroy();
  }
}
