/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_NATIVE_RECEIPT_JOURNAL_LIMIT = 64;

const SHA256_PATTERN = /^[0-9a-f]{64}$/;

/**
 * The native write command an entry's forward ran under: the action the save
 * policy judged that write by. Undo and Redo replay the same writer, so they
 * are judged by the same rule (issueReplayPolicy).
 */
export type PdfNativeReceiptSourceAction =
  // Page-local source writers: each rewrites only its own page content and
  // leaves the package siblings (attachments, layers, Tagged-PDF structure)
  // as exact source-prefix bytes.
  | "pdf.text.edit-content.commit"
  | "pdf.tagged-text.commit"
  | "pdf.vector.geometry.commit"
  | "pdf.vector.style.commit"
  | "pdf.vector.duplicate.commit"
  | "pdf.vector.reorder.commit"
  | "pdf.content.append.commit"
  | "pdf.content.append-image.commit"
  | "pdf.content.append-path.commit"
  | "pdf.text.delete-object"
  | "pdf.text.decoration.commit"
  | "pdf.image.transform-source"
  | "pdf.image.arrange-source"
  | "pdf.image.remove-source"
  | "pdf.image.replace-source"
  // Package writers (annotations, links, outline, forms, page tree and
  // decorations). The save policy judges annotations, links and the outline
  // by the package-sibling rule (PDF_ANNOTATION_WRITER_ACTIONS), the page
  // tree and decorations as page revisions, forms by their own rule
  // (PDF_FORM_AUTHORING_ACTIONS).
  | "pdf.text.free-text.commit"
  | "pdf.page.import-native"
  | "pdf.page.replace-native"
  | "pdf.page.insert-blank"
  | "pdf.page.delete"
  | "pdf.page.reorder"
  | "pdf.page.organize"
  | "pdf.page.rotate"
  | "pdf.page.crop"
  | "pdf.page.split"
  | "pdf.page.background.commit"
  | "pdf.page.resize-media-box"
  | "pdf.page.labels.commit"
  | "pdf.decoration.watermark.commit"
  | "pdf.decoration.page-number.commit"
  | "pdf.decoration.header-footer.commit"
  | "pdf.image.insert.commit"
  | "pdf.image.overlay.commit"
  | "pdf.insert.table.commit"
  | "pdf.insert.chart.commit"
  | "pdf.outline.commit"
  | "pdf.link.commit"
  | "pdf.link.edit.commit"
  | "pdf.signature.visual.commit"
  | "pdf.form.create-detected-fields"
  | "pdf.form.edit-existing-field"
  | "pdf.form.delete-existing-field"
  | "pdf.comment.highlight.commit"
  | "pdf.comment.area-highlight.commit"
  | "pdf.comment.underline.commit"
  | "pdf.comment.strikeout.commit"
  | "pdf.comment.squiggly.commit"
  | "pdf.comment.proofing.commit"
  | "pdf.comment.note.commit"
  | "pdf.comment.batch-delete"
  | "pdf.comment.ink.commit"
  | "pdf.comment.shape.commit"
  | "pdf.comment.stamp.commit"
  | "pdf.comment.file-attachment.commit"
  | "pdf.comment.measure.commit";

/**
 * The forward actions whose history replays under a journal lease: the
 * page-local source writers, which the save policy allows on a PDF that
 * carries attachments, layers or Tagged-PDF structure. Every other named
 * action (and an entry with no name) gets no lease, so its Undo and Redo are
 * judged by the strict rule its forward was judged by -- never wider.
 */
export const PDF_NATIVE_RECEIPT_REPLAY_SOURCE_ACTIONS: ReadonlySet<PdfNativeReceiptSourceAction> =
  new Set<PdfNativeReceiptSourceAction>([
    "pdf.text.edit-content.commit",
    "pdf.tagged-text.commit",
    "pdf.vector.geometry.commit",
    "pdf.vector.style.commit",
    "pdf.vector.duplicate.commit",
    "pdf.vector.reorder.commit",
    "pdf.content.append.commit",
    "pdf.content.append-image.commit",
    "pdf.content.append-path.commit",
    "pdf.page.background.commit",
    "pdf.text.delete-object",
    "pdf.text.decoration.commit",
    "pdf.image.transform-source",
    "pdf.image.arrange-source",
    "pdf.image.remove-source",
    "pdf.image.replace-source",
    // The annotation writers the save policy judges by the same package-
    // sibling rule (PDF_ANNOTATION_WRITER_ACTIONS): their Undo goes as far.
    "pdf.text.free-text.commit",
    "pdf.image.insert.commit",
    "pdf.link.commit",
    "pdf.link.edit.commit",
    "pdf.outline.commit",
    "pdf.signature.visual.commit",
    "pdf.comment.highlight.commit",
    "pdf.comment.area-highlight.commit",
    "pdf.comment.underline.commit",
    "pdf.comment.strikeout.commit",
    "pdf.comment.squiggly.commit",
    "pdf.comment.proofing.commit",
    "pdf.comment.note.commit",
    "pdf.comment.batch-delete",
    "pdf.comment.ink.commit",
    "pdf.comment.shape.commit",
    "pdf.comment.stamp.commit",
    "pdf.comment.file-attachment.commit",
    "pdf.comment.measure.commit",
    // The page-tree and decoration writers (PDF_PAGE_REVISION_ACTIONS): one
    // incremental revision each, judged by the same rule, Undo included.
    "pdf.page.insert-blank",
    "pdf.page.delete",
    "pdf.page.reorder",
    "pdf.page.organize",
    "pdf.page.rotate",
    "pdf.page.crop",
    "pdf.page.split",
    "pdf.page.resize-media-box",
    "pdf.page.labels.commit",
    "pdf.decoration.watermark.commit",
    "pdf.decoration.page-number.commit",
    "pdf.decoration.header-footer.commit",
    // The form authoring writers (PDF_FORM_AUTHORING_ACTIONS): one
    // incremental revision of the field's own objects, Undo included.
    "pdf.form.create-detected-fields",
    "pdf.form.edit-existing-field",
    "pdf.form.delete-existing-field",
  ]);

const replayPolicyBrand: unique symbol = Symbol("PdfNativeReceiptReplayPolicy");
export type PdfNativeReceiptReplayPolicy = Readonly<{ [replayPolicyBrand]: true }>;
type ReplayPolicyView = Readonly<{
  direction: "undo" | "redo";
  /** Every write the step replays: its push's and each extension's. */
  sourceActions: readonly PdfNativeReceiptSourceAction[];
  receipt: PdfNativeReceiptIdentity;
}>;
const replayPolicies = new WeakMap<object, () => ReplayPolicyView | null>();

/** An in-memory journal lease, not a caller-authored action name or JSON flag. */
export function readPdfNativeReceiptReplayPolicy(value: unknown): ReplayPolicyView | null {
  return value && typeof value === "object" ? replayPolicies.get(value)?.() ?? null : null;
}

export type PdfNativeReceiptJournalFamily =
  | "insert"
  | "edit"
  | "page"
  | "comment";

export interface PdfNativeReceiptIdentity {
  readonly resourceId: string;
  readonly resourceVersionId: string;
  readonly outputSha256: string;
}

/** Whether a journal step goes on across a source consolidation, or how to
 * tell from the consolidated journal (PdfNativeReceiptJournalEntry). */
export type PdfNativeSourceConsolidationSurvival =
  | boolean
  | ((consolidatedJournal: unknown) => boolean);

export interface PdfNativeReceiptJournalEntry {
  readonly id: string;
  readonly family: PdfNativeReceiptJournalFamily;
  /** The command the forward write ran under, set by its writer after
   * receipt/reopen. A write that ran under none (a form value fill, an AI
   * document command) names none, and its history keeps the strict rule. */
  readonly sourceAction?: PdfNativeReceiptSourceAction;
  










  readonly survivesSourceConsolidation?: PdfNativeSourceConsolidationSurvival;
  /** The one record whose bytes the step's Undo and Redo write back and
   * forth (a byte restore of it, and nothing else): the window may then
   * stage them on its working copy and move the step when they are saved
   * (recoverReplay). Absent for any other history. */
  readonly restoreOrigin?: string;
  /**
   * One press among several written as one revision (marks pressed while
   * the one before was being written, coalesced into one write of their
   * rows): the step stands on the revision the step before it entered with
   * (`previousReceipt` and `receipt` both name it), and its Undo and Redo
   * write back only its own rows -- planned from rows, never from the
   * revision's bytes -- so each moves the receipt on as any step does. Only
   * pushed right on top of a step that entered with that same receipt; any
   * other step ends the group.
   */
  readonly rowsSubStep?: boolean;
  readonly previousReceipt: PdfNativeReceiptIdentity | null;
  readonly receipt: PdfNativeReceiptIdentity;
  readonly undo: (
    current: PdfNativeReceiptIdentity,
    lease?: PdfNativeReceiptMutationLease,
  ) => Promise<PdfNativeReceiptIdentity | null>;
  readonly redo: (
    current: PdfNativeReceiptIdentity,
    lease?: PdfNativeReceiptMutationLease,
  ) => Promise<PdfNativeReceiptIdentity | null>;
}

export interface PdfNativeReceiptJournalSnapshot {
  readonly version: number;
  readonly length: number;
  readonly cursor: number;
  readonly busy: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly receipt: PdfNativeReceiptIdentity | null;
}

export interface PdfNativeReceiptMutationLease {
  readonly id: symbol;
}

export type PdfNativeMutationDrainResult = "accepted" | "blocked" | "failed";

export class PdfNativeMutationHandoffRegistry<T> {
  readonly #contexts = new Map<string, T>();

  update(pageId: string, context: T): void {
    this.#contexts.set(pageId, context);
  }

  take(pageId: string): T | undefined {
    const context = this.#contexts.get(pageId);
    this.#contexts.delete(pageId);
    return context;
  }

  clear(): void {
    this.#contexts.clear();
  }
}

export class PdfNativeMutationCoalescer<T> {
  readonly #persist: (value: T) => Promise<PdfNativeMutationDrainResult>;
  readonly #flushWaiters = new Set<() => void>();
  readonly #listeners = new Set<() => void>();
  #pending: T | undefined;
  #hasPending = false;
  #latest: T | undefined;
  #hasLatest = false;
  #running = false;
  #scheduled = false;
  #blocked = false;
  #failed = false;
  #failureRetryAvailable = false;
  #generation = 0;
  #revision = 0;

  constructor(persist: (value: T) => Promise<PdfNativeMutationDrainResult>) {
    this.#persist = persist;
  }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  };

  isDirty(): boolean {
    return this.#hasLatest;
  }

  isDraining(): boolean {
    return this.#running || this.#scheduled;
  }

  getRevision(): number {
    return this.#revision;
  }

  acknowledge(revision: number): boolean {
    if (revision !== this.#revision) return false;
    this.takeLatestAndClear();
    return true;
  }

  enqueue(value: T): void {
    this.#revision += 1;
    this.#pending = value;
    this.#hasPending = true;
    this.#latest = value;
    this.#hasLatest = true;
    this.#blocked = false;
    this.#failed = false;
    this.#failureRetryAvailable = true;
    this.#publish();
    this.retry();
  }

  retry(): void {
    if (!this.#hasPending || this.#running || this.#scheduled) return;
    this.#scheduled = true;
    queueMicrotask(() => void this.#drain());
  }

  retryBlocked(options: { retryFailed?: boolean } = {}): void {
    if (this.#blocked) {
      this.#blocked = false;
    } else if (
      options.retryFailed !== false &&
      this.#failed &&
      this.#failureRetryAvailable
    ) {
      this.#failed = false;
      this.#failureRetryAvailable = false;
    } else {
      return;
    }
    this.retry();
  }

  clear(): void {
    this.takeLatestAndClear();
  }

  takeLatestAndClear(): T | undefined {
    const latest = this.#hasLatest ? this.#latest : undefined;
    this.#generation += 1;
    this.#pending = undefined;
    this.#hasPending = false;
    this.#latest = undefined;
    this.#hasLatest = false;
    this.#blocked = false;
    this.#failed = false;
    this.#failureRetryAvailable = false;
    if (!this.#running && !this.#scheduled) this.#settleFlushWaiters();
    this.#publish();
    return latest;
  }

  async flush(): Promise<void> {
    if (!this.#running && !this.#scheduled) return;
    await new Promise<void>((resolve) => this.#flushWaiters.add(resolve));
  }

  async #drain(): Promise<void> {
    this.#scheduled = false;
    if (this.#running || !this.#hasPending) {
      this.#settleFlushWaiters();
      return;
    }
    this.#running = true;
    const generation = this.#generation;
    const value = this.#pending as T;
    this.#pending = undefined;
    this.#hasPending = false;
    let result: PdfNativeMutationDrainResult = "failed";
    try {
      result = await this.#persist(value);
    } catch {
      result = "failed";
    }
    if (generation !== this.#generation) {
      this.#running = false;
      if (this.#hasPending) this.retry();
      this.#settleFlushWaiters();
      return;
    }
    const hadNewerPending = this.#hasPending;
    if (result === "accepted" && !hadNewerPending) {
      this.#latest = undefined;
      this.#hasLatest = false;
      this.#failed = false;
      this.#failureRetryAvailable = false;
    }
    if (result !== "accepted" && !hadNewerPending) {
      this.#pending = value;
      this.#hasPending = true;
    }
    this.#blocked = result === "blocked";
    this.#failed = result === "failed" && !hadNewerPending;
    this.#running = false;
    if (result !== "blocked" && hadNewerPending) this.retry();
    this.#publish();
    this.#settleFlushWaiters();
  }

  #publish(): void {
    for (const listener of this.#listeners) listener();
  }

  #settleFlushWaiters(): void {
    if (this.#running || this.#scheduled) return;
    for (const resolve of this.#flushWaiters) resolve();
    this.#flushWaiters.clear();
  }
}

interface MutableJournalEntry {
  readonly id: string;
  readonly family: PdfNativeReceiptJournalFamily;
  /** The forward action of each write the entry replays; null once any of
   * them was published without one (the strict rule, fail-closed). */
  sourceActions: readonly PdfNativeReceiptSourceAction[] | null;
  readonly survivesSourceConsolidation: PdfNativeSourceConsolidationSurvival;
  /** Null once the step covers more than its one record (extend). */
  restoreOrigin: string | null;
  /** Every record the step restores, in writing order, when each of them
   * restores on its own (a byte restore, or rows written back: an action
   * written as several revisions -- a picture moved, then taken out from
   * under its clips; a box moved with the links on it): Undo restores them
   * newest first, Redo oldest first. Null otherwise. */
  restoreOrigins: readonly string[] | null;
  readonly undo: PdfNativeReceiptJournalEntry["undo"];
  readonly redo: PdfNativeReceiptJournalEntry["redo"];
  appliedReceipt: PdfNativeReceiptIdentity;
  unappliedReceipt: PdfNativeReceiptIdentity | null;
}

function normalizeReceipt(
  value: PdfNativeReceiptIdentity | null | undefined,
): PdfNativeReceiptIdentity | null {
  if (!value || typeof value !== "object") return null;
  const resourceId = value.resourceId.trim();
  const resourceVersionId = value.resourceVersionId.trim();
  const outputSha256 = value.outputSha256.trim().toLowerCase();
  if (!resourceId || !resourceVersionId || !SHA256_PATTERN.test(outputSha256)) {
    return null;
  }
  return Object.freeze({ resourceId, resourceVersionId, outputSha256 });
}

function receiptAdvanced(
  previous: PdfNativeReceiptIdentity,
  next: PdfNativeReceiptIdentity,
): boolean {
  return (
    previous.resourceId === next.resourceId &&
    previous.resourceVersionId !== next.resourceVersionId &&
    previous.outputSha256 !== next.outputSha256
  );
}

function sameReceipt(
  left: PdfNativeReceiptIdentity | null,
  right: PdfNativeReceiptIdentity | null,
): boolean {
  return Boolean(
    left &&
    right &&
    left.resourceId === right.resourceId &&
    left.resourceVersionId === right.resourceVersionId &&
    left.outputSha256 === right.outputSha256,
  );
}

export class PdfNativeReceiptJournal {
  readonly #limit: number;
  readonly #listeners = new Set<() => void>();
  #entries: MutableJournalEntry[] = [];
  #cursor = 0;
  #busy = false;
  #version = 0;
  #epoch = 0;
  #activeReplay: symbol | null = null;
  #activeMutationLease: PdfNativeReceiptMutationLease | null = null;
  #receipt: PdfNativeReceiptIdentity | null = null;
  #snapshot: PdfNativeReceiptJournalSnapshot;

  constructor(limit = PDF_NATIVE_RECEIPT_JOURNAL_LIMIT) {
    if (!Number.isSafeInteger(limit) || limit < 1 || limit > 256) {
      throw new Error("Invalid native PDF receipt journal limit.");
    }
    this.#limit = limit;
    this.#snapshot = this.#buildSnapshot();
  }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  };

  readonly getSnapshot = (): PdfNativeReceiptJournalSnapshot => this.#snapshot;

  getVersion = (): number => this.#version;

  canUndo = (): boolean => this.#snapshot.canUndo;

  canRedo = (): boolean => this.#snapshot.canRedo;

  isBusy = (): boolean => this.#busy;

  issueReplayPolicy(direction: "undo" | "redo"): PdfNativeReceiptReplayPolicy | null {
    if (this.#busy) return null;
    const index = direction === "undo" ? this.#cursor - 1 : this.#cursor;
    const entry = this.#entries[index];
    const receipt = this.#receipt;
    const sourceActions = entry?.sourceActions;
    if (!receipt || !entry || (direction === "redo" && !entry.unappliedReceipt) ||
        !sourceActions?.length ||
        !sourceActions.every((action) => PDF_NATIVE_RECEIPT_REPLAY_SOURCE_ACTIONS.has(action))) return null;
    const epoch = this.#epoch;
    const cursor = this.#cursor;
    const policy = Object.freeze(Object.create(null)) as PdfNativeReceiptReplayPolicy;
    const view = Object.freeze({ direction, sourceActions, receipt });
    replayPolicies.set(policy, () => this.#epoch === epoch && this.#cursor === cursor &&
      this.#entries[index] === entry && sameReceipt(this.#receipt, receipt) ? view : null);
    return policy;
  }

  /** The step at a place in the history (cursor - 1 is the one Undo
   * replays, cursor the one Redo replays): its id and the record it
   * restores, when that is all its history does. */
  stepAt(index: number): Readonly<{ id: string; restoreOrigin: string | null;
    /** Its records, when it restores several (extend). */
    restoreOrigins?: readonly string[] }> | null {
    const entry = Number.isSafeInteger(index) ? this.#entries[index] : undefined;
    return entry ? Object.freeze({ id: entry.id, restoreOrigin: entry.restoreOrigin,
      ...(entry.restoreOrigins && entry.restoreOrigins.length > 1 ? { restoreOrigins: entry.restoreOrigins } : {}) })
      : null;
  }

  ownsMutation(lease: PdfNativeReceiptMutationLease): boolean {
    return this.#activeMutationLease === lease;
  }

  /** Idempotent publication after a durable recovery-hold retirement fails.
   * A moved history cursor or different receipt is never the same publication. */
  hasAppliedEntry(id: string, receipt: PdfNativeReceiptIdentity): boolean {
    const entry = this.#entries[this.#cursor - 1];
    return this.#cursor === this.#entries.length && !!entry && entry.id === id &&
      sameReceipt(this.#receipt, receipt) && sameReceipt(entry.appliedReceipt, receipt);
  }

  /** Publish an already authenticated history save after its original call
   * returned without a receipt. This never invokes the writer a second time. */
  recoverReplay(input: Readonly<{
    entryId: string; direction: "undo" | "redo";
    previousReceipt: PdfNativeReceiptIdentity; receipt: PdfNativeReceiptIdentity;
  }>, lease: PdfNativeReceiptMutationLease): boolean {
    if (this.#activeMutationLease !== lease || this.#activeReplay !== null) return false;
    const previous = normalizeReceipt(input.previousReceipt), next = normalizeReceipt(input.receipt);
    if (!previous || !next || !receiptAdvanced(previous,next)) return false;
    const completed = this.#entries[input.direction === "undo" ? this.#cursor : this.#cursor-1];
    if (completed?.id === input.entryId && sameReceipt(this.#receipt,next) &&
        sameReceipt(input.direction === "undo" ? completed.unappliedReceipt : completed.appliedReceipt,next)) return true;
    const entry = this.#entries[input.direction === "undo" ? this.#cursor-1 : this.#cursor];
    if (!entry || entry.id !== input.entryId || !sameReceipt(this.#receipt,previous) ||
        (input.direction === "redo" && !entry.unappliedReceipt)) return false;
    if (input.direction === "undo") {entry.unappliedReceipt=next;this.#cursor-=1;}
    else {entry.appliedReceipt=next;entry.unappliedReceipt=null;this.#cursor+=1;}
    this.#receipt=next;this.#epoch+=1;this.#publish();
    return true;
  }

  /** One user action written as several revisions (a delete of several
   * objects writes one record per object): the newest entry -- applied, on
   * top, and leaving exactly `previousReceipt` -- takes the next revision as
   * its own, so Undo and Redo replay the whole action as one step. Only the
   * writer holding the mutation lease may extend it; anything else is a new
   * entry or no entry at all. The revision names the action it was written
   * under: the step is then judged by every action it replays, and one
   * revision without a name leaves the whole step to the strict rule. */
  extend(input: Readonly<{
    entryId: string;
    previousReceipt: PdfNativeReceiptIdentity | null;
    receipt: PdfNativeReceiptIdentity;
    sourceAction?: PdfNativeReceiptSourceAction;
    /** The revision's record, when its own history is a byte restore of it.
     * The step's own one record given again is that record widened to hold
     * this revision too (a run of keys typed into a text box: its rows from
     * before the first key to after the last): the step still restores by
     * it alone. */
    restoreOrigin?: string;
  }>, lease: PdfNativeReceiptMutationLease): boolean {
    if (this.#activeMutationLease !== lease || this.#activeReplay !== null) return false;
    const previous = normalizeReceipt(input.previousReceipt), next = normalizeReceipt(input.receipt);
    const entry = this.#entries[this.#cursor - 1];
    if (!previous || !next || !entry || entry.id !== input.entryId ||
        this.#cursor !== this.#entries.length || entry.unappliedReceipt !== null ||
        !sameReceipt(this.#receipt, previous) || !sameReceipt(entry.appliedReceipt, previous) ||
        !receiptAdvanced(previous, next)) return false;
    entry.sourceActions = entry.sourceActions && input.sourceAction
      ? Object.freeze([...new Set([...entry.sourceActions, input.sourceAction])])
      : null;
    const origin = input.restoreOrigin?.trim();
    if (!origin || entry.restoreOrigin !== origin) {
      entry.restoreOrigins = entry.restoreOrigins && origin
        ? Object.freeze([...entry.restoreOrigins, origin]) : null;
      entry.restoreOrigin = null;
    }
    entry.appliedReceipt = next;
    this.#receipt = next;
    this.#epoch += 1;
    this.#publish();
    return true;
  }

  /**
   * A write that is no step of the history moved the document on: Save
   * writing the edits AutoSave off held, which keep their own Undo in the
   * window. The history goes on from the new receipt with every step kept
   * (invalidating it lost the Undo of everything saved before). Only the
   * writer holding the mutation lease may move it, and only on from the
   * receipt the history stands on.
   */
  bridge(input: Readonly<{
    previousReceipt: PdfNativeReceiptIdentity | null;
    receipt: PdfNativeReceiptIdentity;
  }>, lease: PdfNativeReceiptMutationLease): boolean {
    if (this.#activeMutationLease !== lease || this.#activeReplay !== null) return false;
    const previous = normalizeReceipt(input.previousReceipt), next = normalizeReceipt(input.receipt);
    if (!next) return false;
    if (this.#receipt) {
      if (!previous || !sameReceipt(this.#receipt, previous) || !receiptAdvanced(previous, next)) return false;
      const applied = this.#entries[this.#cursor - 1];
      if (applied && sameReceipt(applied.appliedReceipt, previous)) applied.appliedReceipt = next;
      const undone = this.#entries[this.#cursor];
      if (undone && sameReceipt(undone.unappliedReceipt, previous)) undone.unappliedReceipt = next;
    }
    this.#receipt = next;
    this.#epoch += 1;
    this.#publish();
    return true;
  }

  /**
   * The current bytes published again under another version of the same
   * resource (a page restored to its last good state after the server
   * refused a write: the export it verified is re-read and published anew).
   * Nothing the history describes changed, so every step keeps its place;
   * the receipt that names the current bytes, and the steps that stand on
   * it, take the new version. Only the writer holding the mutation lease
   * may move it, and only between two names of the same bytes. A replay in
   * flight is the write that was refused: the moved epoch keeps it from
   * publishing anything, and Undo and Redo go on from the new name.
   */
  rebaseReceipt(input: Readonly<{
    previousReceipt: PdfNativeReceiptIdentity;
    receipt: PdfNativeReceiptIdentity;
  }>, lease: PdfNativeReceiptMutationLease): boolean {
    if (this.#activeMutationLease !== lease) return false;
    const previous = normalizeReceipt(input.previousReceipt), next = normalizeReceipt(input.receipt);
    if (!previous || !next || !sameReceipt(this.#receipt, previous) ||
        previous.resourceId !== next.resourceId || previous.outputSha256 !== next.outputSha256 ||
        previous.resourceVersionId === next.resourceVersionId) return false;
    const applied = this.#entries[this.#cursor - 1];
    if (applied && sameReceipt(applied.appliedReceipt, previous)) applied.appliedReceipt = next;
    const undone = this.#entries[this.#cursor];
    if (undone && sameReceipt(undone.unappliedReceipt, previous)) undone.unappliedReceipt = next;
    this.#receipt = next;
    this.#epoch += 1;
    this.#publish();
    return true;
  }

  beginMutation(): PdfNativeReceiptMutationLease | null {
    if (this.#busy) return null;
    const lease = Object.freeze({ id: Symbol("pdf-native-mutation") });
    this.#activeMutationLease = lease;
    this.#busy = true;
    this.#publish();
    return lease;
  }

  endMutation(lease: PdfNativeReceiptMutationLease): boolean {
    if (this.#activeMutationLease !== lease) return false;
    this.#activeMutationLease = null;
    this.#busy = false;
    this.#publish();
    return true;
  }

  push(
    entry: PdfNativeReceiptJournalEntry,
    lease?: PdfNativeReceiptMutationLease,
  ): boolean {
    const ownsMutationLease = this.#activeMutationLease === lease;
    if (
      !entry.id.trim() ||
      (this.#busy && !ownsMutationLease) ||
      (lease !== undefined && !ownsMutationLease)
    ) {
      return false;
    }
    const previousReceipt = normalizeReceipt(entry.previousReceipt);
    const receipt = normalizeReceipt(entry.receipt);
    if (!receipt) return false;
    // A rows sub-step continues the group its revision opened: on top of
    // the history, right after a step applied with that same receipt.
    const top = this.#entries[this.#cursor - 1];
    const subStep = entry.rowsSubStep === true && previousReceipt !== null &&
      sameReceipt(previousReceipt, receipt) && this.#cursor === this.#entries.length &&
      top !== undefined && top.unappliedReceipt === null && sameReceipt(top.appliedReceipt, receipt);
    if (entry.rowsSubStep === true && !subStep) return false;
    if (
      (this.#receipt && !sameReceipt(this.#receipt, previousReceipt)) ||
      (previousReceipt && !subStep && !receiptAdvanced(previousReceipt, receipt))
    ) {
      return false;
    }

    if (this.#cursor < this.#entries.length) {
      this.#entries = this.#entries.slice(0, this.#cursor);
    }
    this.#entries.push({
      id: entry.id,
      family: entry.family,
      sourceActions: entry.sourceAction ? Object.freeze([entry.sourceAction]) : null,
      survivesSourceConsolidation:
        typeof entry.survivesSourceConsolidation === "function"
          ? entry.survivesSourceConsolidation
          : entry.survivesSourceConsolidation !== false,
      restoreOrigin: entry.restoreOrigin?.trim() || null,
      restoreOrigins: entry.restoreOrigin?.trim() ? Object.freeze([entry.restoreOrigin.trim()]) : null,
      undo: entry.undo,
      redo: entry.redo,
      appliedReceipt: receipt,
      unappliedReceipt: null,
    });
    this.#cursor = this.#entries.length;
    this.#receipt = receipt;
    this.#epoch += 1;

    if (this.#entries.length > this.#limit) {
      const overflow = this.#entries.length - this.#limit;
      this.#entries.splice(0, overflow);
      this.#cursor -= overflow;
    }
    this.#publish();
    return true;
  }

  async undo(lease?: PdfNativeReceiptMutationLease): Promise<boolean> {
    const ownsMutationLease = this.#activeMutationLease === lease;
    if (
      (this.#busy && !ownsMutationLease) ||
      (lease !== undefined && !ownsMutationLease) ||
      this.#cursor === 0 ||
      !this.#receipt
    ) {
      return false;
    }
    const entry = this.#entries[this.#cursor - 1];
    if (!entry) return false;
    return this.#replay("undo", entry, lease);
  }

  async redo(lease?: PdfNativeReceiptMutationLease): Promise<boolean> {
    const ownsMutationLease = this.#activeMutationLease === lease;
    if (
      (this.#busy && !ownsMutationLease) ||
      (lease !== undefined && !ownsMutationLease) ||
      this.#cursor >= this.#entries.length ||
      !this.#receipt
    ) {
      return false;
    }
    const entry = this.#entries[this.#cursor];
    if (!entry || !entry.unappliedReceipt) return false;
    return this.#replay("redo", entry, lease);
  }

  clear(): void {
    if (this.#entries.length === 0 && !this.#receipt && !this.#busy) return;
    this.invalidate();
  }

  /**
   * The history after a save that consolidated the source chain into
   * `consolidatedJournal` (the saved page's native journal). Steps that
   * cannot replay across the rewrite (survivesSourceConsolidation false, or
   * a predicate that finds nothing for them in that journal) leave it with
   * every step Undo or Redo could only reach through them: an applied one
   * takes the older steps with it, an undone one the newer redo steps.
   * Everything else keeps its place, its receipts and the cursor's meaning.
   * Only the writer holding the mutation lease may call it, and never while
   * a replay is in flight.
   */
  retainAcrossSourceConsolidation(
    lease?: PdfNativeReceiptMutationLease,
    consolidatedJournal?: unknown,
  ): boolean {
    if ((lease !== undefined && this.#activeMutationLease !== lease) || this.#activeReplay !== null) {
      return false;
    }
    const survives = this.#entries.map(({ survivesSourceConsolidation: rule }) => {
      if (typeof rule !== "function") return rule;
      try {
        return rule(consolidatedJournal) === true;
      } catch {
        return false;
      }
    });
    let applied = 0;
    for (let index = 0; index < this.#cursor; index += 1) {
      if (!survives[index]) applied = index + 1;
    }
    let undone = this.#entries.length;
    for (let index = this.#entries.length - 1; index >= this.#cursor; index -= 1) {
      if (!survives[index]) undone = index;
    }
    if (applied === 0 && undone === this.#entries.length) return true;
    if (process.env.NODE_ENV !== "production") {
      try {
        console.warn("[pdf-history] consolidation dropped steps", { entries: this.#entries.length, applied, undone });
      } catch { /* diagnostics only */ }
    }
    this.#entries = this.#entries.slice(applied, undone);
    this.#cursor -= applied;
    this.#epoch += 1;
    this.#publish();
    return true;
  }

  invalidate(lease?: PdfNativeReceiptMutationLease): boolean {
    if (lease !== undefined && this.#activeMutationLease !== lease) {
      return false;
    }
    if (process.env.NODE_ENV !== "production" && this.#entries.length) {
      try {
        console.warn("[pdf-history] journal cleared", { entries: this.#entries.length, cursor: this.#cursor,
          at: new Error().stack?.split("\n").slice(2, 8).map((line) => line.trim()).join(" < ") });
      } catch { /* diagnostics only */ }
    }
    this.#epoch += 1;
    this.#activeReplay = null;
    this.#entries = [];
    this.#cursor = 0;
    this.#receipt = null;
    if (lease === undefined) {
      this.#activeMutationLease = null;
      this.#busy = false;
    }
    this.#publish();
    return true;
  }

  async #replay(
    direction: "undo" | "redo",
    entry: MutableJournalEntry,
    lease?: PdfNativeReceiptMutationLease,
  ): Promise<boolean> {
    const current = this.#receipt;
    if (!current) return false;
    const replayEpoch = this.#epoch;
    const replayToken = Symbol(direction);
    const expectedCursor = this.#cursor;
    this.#activeReplay = replayToken;
    const ownsMutationLease = this.#activeMutationLease === lease;
    this.#busy = true;
    this.#publish();
    try {
      const replayed = normalizeReceipt(await entry[direction](current, lease));
      if (
        this.#activeReplay !== replayToken ||
        this.#epoch !== replayEpoch ||
        this.#cursor !== expectedCursor
      ) {
        return false;
      }
      if (!replayed || !receiptAdvanced(current, replayed)) return false;
      if (direction === "undo") {
        entry.unappliedReceipt = replayed;
        this.#cursor -= 1;
      } else {
        entry.appliedReceipt = replayed;
        entry.unappliedReceipt = null;
        this.#cursor += 1;
      }
      this.#receipt = replayed;
      this.#epoch += 1;
      return true;
    } catch {
      return false;
    } finally {
      if (this.#activeReplay === replayToken) {
        this.#activeReplay = null;
        if (!ownsMutationLease) this.#busy = false;
        this.#publish();
      }
    }
  }

  #buildSnapshot(): PdfNativeReceiptJournalSnapshot {
    return Object.freeze({
      version: this.#version,
      length: this.#entries.length,
      cursor: this.#cursor,
      busy: this.#busy,
      canUndo: !this.#busy && this.#cursor > 0,
      canRedo: !this.#busy && this.#cursor < this.#entries.length,
      receipt: this.#receipt,
    });
  }

  #publish(): void {
    this.#version += 1;
    this.#snapshot = this.#buildSnapshot();
    for (const listener of this.#listeners) listener();
  }
}
