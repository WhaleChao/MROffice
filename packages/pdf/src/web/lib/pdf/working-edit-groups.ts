/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfRowStepRunJoined, type PdfRowStepRowsAndCompanions } from "./row-step-target";

export interface PdfWorkingEditGroup<P, A extends string = string, I = unknown> {
  readonly id: string;
  /** The command the action ran under (its history step's source action). */
  readonly sourceAction: A;
  /** The command the edit is written under when the queue publishes it. */
  readonly invocation: I | null;
  /** The action Undo takes back in one press while nothing is saved yet. */
  actionId: string;
  readonly opIds: string[];
   
  publisher: P | null;
   
  historyAtPublication: ((stored: unknown) => P | null) | null;
  /** The step of an earlier action this one continues. */
  extendsEntryId: string | null;
  /** The editor history entry the group's first published edit opened. */
  entryId: string | null;
   
  notTaken: boolean;
  closed: boolean;
  /** Resolves when the action has staged everything it will stage. */
  readonly complete: Promise<void>;
}

/** How one published edit enters the editor history. */
export type PdfWorkingEditPublication<P> =
  | Readonly<{ kind: "step"; publisher: P | null; historyAtPublication: ((stored: unknown) => P | null) | null }>
  | Readonly<{ kind: "extend"; entryId: string }>;

export function createPdfWorkingEditGroups<P extends { readonly id?: string }, A extends string = string, I = unknown>() {
  type Group = PdfWorkingEditGroup<P, A, I>;
  const byLease = new WeakMap<object, Group>();
  const byOp = new Map<string, Group>();
  /** An edit written under another command than its action's (an erase
   * deletes text, removes a picture and draws a cover in one action). */
  const opInvocations = new Map<string, I>();
  

  const publishedHooks = new Map<string, (stored: unknown) => void>();
  const finishers = new WeakMap<Group, () => void>();
  return {
     
    open(sourceAction: A, invocation: I | null = null): Readonly<{ lease: object; group: Group }> {
      let finish: () => void = () => undefined;
      const complete = new Promise<void>((resolve) => { finish = resolve; });
      const id = crypto.randomUUID();
      const group: Group = {
        id, sourceAction, invocation, actionId: id, opIds: [], publisher: null, historyAtPublication: null,
        extendsEntryId: null, entryId: null, notTaken: false, closed: false, complete,
      };
      finishers.set(group, finish);
      const lease = Object.freeze({ id: Symbol("pdf-working-edit") });
      byLease.set(lease, group);
      return Object.freeze({ lease, group });
    },
    forLease(lease: unknown): Group | undefined {
      return lease && typeof lease === "object" ? byLease.get(lease) : undefined;
    },
    forOp(opId: string): Group | undefined {
      return byOp.get(opId);
    },
    addOp(group: Group, opId: string, invocation?: I, onPublished?: (stored: unknown) => void): void {
      if (group.closed) throw new Error("PDF working edit group is closed.");
      group.opIds.push(opId);
      byOp.set(opId, group);
      if (invocation !== undefined) opInvocations.set(opId, invocation);
      if (onPublished) publishedHooks.set(opId, onPublished);
    },
    /** An edit taken back off the copy before its action was complete (a
     * line's delete of several runs that stopped part way): no edit of the
     * action any more, so the first edit that is published opens its step. */
    removeOp(group: Group, opId: string): void {
      const at = group.opIds.indexOf(opId);
      if (at >= 0) group.opIds.splice(at, 1);
      if (byOp.get(opId) === group) byOp.delete(opId);
      opInvocations.delete(opId);
      publishedHooks.delete(opId);
    },
    /** The command an edit is published under. */
    invocationOf(opId: string): I | null {
      return opInvocations.get(opId) ?? byOp.get(opId)?.invocation ?? null;
    },
    

    keepPublisher(group: Group, publisher: P): void {
      if (!group.publisher) group.publisher = publisher;
    },
    /** The action continues the step `entryId` names (a delete batch's
     * later objects): Undo before saving takes them back together. */
    keepExtension(group: Group, entryId: string): void {
      if (group.publisher || group.extendsEntryId) return;
      group.extendsEntryId = entryId;
      for (const other of new Set(byOp.values())) {
        if (other !== group && (other.publisher?.id === entryId || other.entryId === entryId)) {
          group.actionId = other.actionId;
          break;
        }
      }
    },
    close(group: Group): void {
      if (group.closed) return;
      group.closed = true;
      finishers.get(group)?.();
    },
    /** The action an edit belongs to (an edit outside any group is its own). */
    actionOf(opId: string): string {
      return byOp.get(opId)?.actionId ?? opId;
    },
    /** How the publication of `opId` enters the editor history. */
    publication(opId: string): PdfWorkingEditPublication<P> | null {
      const group = byOp.get(opId);
      if (!group) return null;
      if (group.opIds[0] !== opId) {
        return group.entryId ? { kind: "extend", entryId: group.entryId }
          : { kind: "step", publisher: null, historyAtPublication: null };
      }
      if (group.extendsEntryId) return { kind: "extend", entryId: group.extendsEntryId };
      return { kind: "step", publisher: group.publisher, historyAtPublication: group.historyAtPublication };
    },
    /** The entry the edit's publication opened (its group's later edits extend it). */
    published(opId: string, entryId: string): void {
      const group = byOp.get(opId);
      if (group && !group.entryId) group.entryId = entryId;
    },
     
    stored(opId: string, stored: unknown): void {
      const hook = publishedHooks.get(opId);
      publishedHooks.delete(opId);
      try { hook?.(stored); } catch { /* a history seed that cannot be read leaves that step out */ }
    },
    /** Keep only the edits the queue still holds (pending or undone). */
    retain(opIds: Iterable<string>): void {
      const kept = new Set(opIds);
      for (const opId of [...byOp.keys()]) {
        if (kept.has(opId)) continue;
        byOp.delete(opId); opInvocations.delete(opId); publishedHooks.delete(opId);
      }
    },
    clear(): void {
      byOp.clear(); opInvocations.clear(); publishedHooks.clear();
    },
  };
}

export type PdfWorkingEditGroups<P extends { readonly id?: string }, A extends string = string, I = unknown> =
  ReturnType<typeof createPdfWorkingEditGroups<P, A, I>>;

/**
 * A record written after the edit it goes with -- too late to be an edit of
 * that edit's action (links set again after the move of what they stand on;
 * a text box only the text-box writer keeps, written once the delete of the
 * line that became it is saved) -- joins that edit's step when it is saved
 * itself: the step on top of the history, nothing undone above it, which
 * the edit opened or is a record of. One Undo then takes both back.
 */
export function pdfFollowingRecordJoinsStep(input: Readonly<{
  top: Readonly<{ id: string; restoreOrigin?: string | null; restoreOrigins?: readonly string[] | null }> | null | undefined;
  /** Nothing of the history is undone. */
  atEnd: boolean;
  /** The edit it follows. */
  anchorOpId: string;
  /** The entry that edit's action opened, when the action has several edits. */
  anchorEntryId?: string | null;
}>): boolean {
  const { top, anchorOpId, anchorEntryId } = input;
  return Boolean(top && input.atEnd && (top.id === anchorOpId || (anchorEntryId && top.id === anchorEntryId) ||
    top.restoreOrigin === anchorOpId || top.restoreOrigins?.includes(anchorOpId)));
}

/** Why a write of a run does not join the run's step (pdfRunStepJoin). */
export type PdfRunStepJoinRefusal =
  /** No run's step is open: this write opens one. */
  | "no-run-open"
  /** The open step is another run's (other boxes). */
  | "another-run"
  /** Another step is the newest of the history (a move, a format, another box). */
  | "step-not-newest"
  /** A step is undone above it. */
  | "undone-above"
  /** The history stands on another revision than the run's last write left
   * (an Undo and its Redo, a write that is no step). */
  | "another-revision"
  /** One of the two is a row edit of the working copy, the other a page save. */
  | "rows-on-one-side"
  /** A row edit whose rows the window no longer holds. */
  | "rows-unknown"
  /** Rows, or page collections changed with them, that do not join
   * (pdfRowStepRunJoined). */
  | "rows-not-joined";




















export function pdfRunStepJoin<Rows extends PdfRowStepRowsAndCompanions>(input: Readonly<{
  /** The run's step so far: its key, its entry and the revision it left. */
  run: Readonly<{ key: string; entryId: string;
    receipt: Readonly<{ resourceVersionId: string; outputSha256: string }> | null }> | null | undefined;
  /** The run this write is of. */
  key: string;
  top: Readonly<{ id: string; restoreOrigin?: string | null }> | null | undefined;
  /** Nothing of the history is undone. */
  atEnd: boolean;
  /** The revision the history stands on. */
  receipt: Readonly<{ resourceVersionId: string; outputSha256: string }> | null | undefined;
  /** This write's own record (`rows:<id>` for a row edit of the copy). */
  restoreOrigin?: string | null;
  /** The rows a row edit's record holds, by its id. */
  rowsOf: (id: string) => Rows | undefined;
  /** The page collections the run's action changes beside its rows. */
  collections?: readonly string[];
}>): Readonly<{ joined: true; rows: Rows | null }> | Readonly<{ joined: false; reason: PdfRunStepJoinRefusal }> {
  const { run, top, receipt } = input;
  const no = (reason: PdfRunStepJoinRefusal) => ({ joined: false as const, reason });
  if (!run) return no("no-run-open");
  if (run.key !== input.key) return no("another-run");
  if (!top || top.id !== run.entryId) return no("step-not-newest");
  if (!input.atEnd) return no("undone-above");
  if (!run.receipt || !receipt || run.receipt.resourceVersionId !== receipt.resourceVersionId ||
    run.receipt.outputSha256 !== receipt.outputSha256) return no("another-revision");
  const runIsRows = top.restoreOrigin === `rows:${run.entryId}`;
  const ownRowsId = input.restoreOrigin?.startsWith("rows:") ? input.restoreOrigin.slice(5) : null;
  if (!runIsRows && !ownRowsId) return { joined: true, rows: null };
  if (!runIsRows || !ownRowsId) return no("rows-on-one-side");
  const runRows = input.rowsOf(run.entryId);
  const ownRows = input.rowsOf(ownRowsId);
  if (!runRows || !ownRows) return no("rows-unknown");
  const rows = pdfRowStepRunJoined(runRows, ownRows, input.collections);
  return rows ? { joined: true, rows } : no("rows-not-joined");
}

/**
 * Until it is saved, Undo does not take such a record off the working copy
 * alone once the edit it follows is saved: that took back half a step (the
 * new box gone, the line it replaced still deleted). The press waits for
 * its save and takes the whole step. While both are still on the copy they
 * are one press there already.
 */
export function pdfFollowingEditWaitsForItsStep(newestOpId: string | undefined,
  follows: ReadonlyMap<string, string>, pendingOpIds: readonly string[]): boolean {
  const anchor = newestOpId ? follows.get(newestOpId) : undefined;
  return anchor !== undefined && !pendingOpIds.includes(anchor);
}
