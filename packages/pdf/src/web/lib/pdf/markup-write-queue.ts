/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfMarkupWrite {
  /** The mark's id (its reader row's and its PDF annotation's). */
  readonly id: string;
  /** Writes it: the row, then the mark in the file. Rejects when refused. */
  readonly run: () => Promise<void>;
  /** Taken back before it was written: off the page. */
  readonly withdraw: () => void;
  /** Put back by Redo: on the page again (it is queued again). */
  readonly restore: () => void;
  







  readonly runWith?: (writes: readonly PdfMarkupWrite[]) => Promise<void>;
}

type QueuedWrite = PdfMarkupWrite & {
  ticket?: number;
  historyStep?: true;
  onError?: (error: unknown, write: PdfMarkupWrite) => void;
};

type RedoEntry = { readonly kind: "mark"; readonly write: QueuedWrite } | { readonly kind: "history" };

interface Lane {
  readonly waiting: QueuedWrite[];
  running: QueuedWrite | null;
  /** What Redo walks back, newest last: marks withdrawn from the queue and
   * the window's own Undos of written steps, in the order they were taken. */
  readonly redo: RedoEntry[];
  drained: Promise<void>;
  resolveDrained: (() => void) | null;
  listeners: Set<() => void>;
}

const lanes = new Map<string, Lane>();

function laneOf(key: string): Lane {
  let lane = lanes.get(key);
  if (!lane) {
    lane = { waiting: [], running: null, redo: [], drained: Promise.resolve(), resolveDrained: null,
      listeners: new Set() };
    lanes.set(key, lane);
  }
  return lane;
}

function notify(lane: Lane): void {
  for (const listener of [...lane.listeners]) {
    try { listener(); } catch { /* a listener never stops the queue */ }
  }
}

// --- Press-order tickets ---------------------------------------------------

interface TicketBook {
  next: number;
  readonly open: Set<number>;
  readonly waiters: Array<{ ticket: number; resolve: () => void }>;
}
const tickets = new Map<string, TicketBook>();

function bookOf(key: string): TicketBook {
  let book = tickets.get(key);
  if (!book) {
    book = { next: 1, open: new Set(), waiters: [] };
    tickets.set(key, book);
  }
  return book;
}

function wakeWaiters(book: TicketBook): void {
  for (const waiter of [...book.waiters]) {
    if ([...book.open].every((open) => open >= waiter.ticket)) {
      book.waiters.splice(book.waiters.indexOf(waiter), 1);
      waiter.resolve();
    }
  }
}

/** A press that stages an edit on the working copy: its place in line. */
export function takePdfEditTicket(key: string): number {
  const book = bookOf(key);
  const ticket = book.next++;
  book.open.add(ticket);
  return ticket;
}

/** Staged (or given up): the presses after it may stage. */
export function releasePdfEditTicket(key: string, ticket: number | undefined): void {
  if (ticket === undefined) return;
  const book = tickets.get(key);
  if (!book || !book.open.delete(ticket)) return;
  wakeWaiters(book);
}

/** Settles once every edit pressed before `ticket` is staged (or given up).
 * Bounded: a press never waits on one before it longer than `timeoutMs`. */
export function waitPdfEditTurn(key: string, ticket: number, timeoutMs = 20_000): Promise<void> {
  const book = bookOf(key);
  if ([...book.open].every((open) => open >= ticket)) return Promise.resolve();
  return new Promise<void>((resolve) => {
    let settled = false;
    const done = () => { if (!settled) { settled = true; resolve(); } };
    book.waiters.push({ ticket, resolve: done });
    setTimeout(() => {
      const at = book.waiters.findIndex((waiter) => waiter.resolve === done);
      if (at >= 0) book.waiters.splice(at, 1);
      done();
    }, timeoutMs);
  });
}



function pump(key: string, lane: Lane): void {
  if (lane.running) return;
  const next = lane.waiting.shift();
  if (!next) {
    const resolve = lane.resolveDrained;
    lane.resolveDrained = null;
    resolve?.();
    notify(lane);
    if (!lane.redo.length && !lane.listeners.size) lanes.delete(key);
    return;
  }
  // The marks pressed next in line behind it go with it (a press of
  // anything else between them -- a line, an Undo -- ends the group).
  const batch: QueuedWrite[] = [next];
  if (next.runWith && !next.historyStep && next.ticket !== undefined) {
    while (lane.waiting.length) {
      const candidate = lane.waiting[0];
      const last = batch[batch.length - 1];
      if (candidate.historyStep || !candidate.runWith || candidate.ticket === undefined ||
        candidate.ticket !== (last.ticket ?? Number.NaN) + 1) break;
      batch.push(lane.waiting.shift()!);
    }
  }
  lane.running = next;
  notify(lane);
  const turn = next.ticket !== undefined ? waitPdfEditTurn(key, next.ticket) : Promise.resolve();
  void turn
    .then(() => (batch.length > 1 ? next.runWith!(batch) : next.run()))
    .catch((error) => {
      for (const write of batch) {
        try { write.onError?.(error, write); } catch { /* the queue still moves */ }
      }
    })
    .finally(() => {
      for (const write of batch) releasePdfEditTicket(key, write.ticket);
      lane.running = null;
      pump(key, lane);
    });
}

function startDrain(lane: Lane): void {
  if (!lane.running && !lane.waiting.length) {
    lane.drained = new Promise<void>((resolve) => { lane.resolveDrained = resolve; });
  }
}

/** Queue one press's write. A new mark ends what Redo could put back.
 * `onError` is this press's own (a refused write takes its mark back). */
export function enqueuePdfMarkupWrite(key: string, write: PdfMarkupWrite,
  onError: (error: unknown, write: PdfMarkupWrite) => void): void {
  const lane = laneOf(key);
  lane.redo.length = 0;
  startDrain(lane);
  lane.waiting.push({ ...write, ticket: takePdfEditTicket(key), onError });
  pump(key, lane);
}





export function enqueuePdfMarkupHistoryStep<T>(key: string, run: () => Promise<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const write: QueuedWrite = {
      id: `history:${Math.random().toString(36).slice(2)}`,
      run: () => run().then(resolve, (error) => { reject(error); }),
      withdraw: () => undefined,
      restore: () => undefined,
      historyStep: true,
    };
    const lane = laneOf(key);
    startDrain(lane);
    lane.waiting.push(write);
    pump(key, lane);
  });
}

/** Undo of the last mark when it has not been written yet: taken off the
 * queue and the page; true when there was one. */
export function withdrawQueuedPdfMarkup(key: string): boolean {
  const lane = lanes.get(key);
  const last = lane?.waiting.at(-1);
  if (!lane || !last || last.historyStep) return false;
  lane.waiting.pop();
  releasePdfEditTicket(key, last.ticket);
  lane.redo.push({ kind: "mark", write: last });
  try { last.withdraw(); } catch { /* the queue still moves */ }
  notify(lane);
  return true;
}

/** The window's Undo of a written step was taken: Redo walks it back in
 * its turn among the withdrawn marks. */
export function notePdfHistoryUndone(key: string): void {
  laneOf(key).redo.push({ kind: "history" });
}





export function takePdfRedo(key: string): "mark" | "history" | null {
  const lane = lanes.get(key);
  const entry = lane?.redo.pop();
  if (!lane || !entry) return null;
  if (entry.kind === "history") return "history";
  const write = entry.write;
  try { write.restore(); } catch { /* queued all the same */ }
  startDrain(lane);
  lane.waiting.push({ ...write, ticket: takePdfEditTicket(key) });
  pump(key, lane);
  notify(lane);
  return "mark";
}

/** A new edit: nothing is walked back by Redo any more. */
export function clearPdfRedo(key: string): void {
  const lane = lanes.get(key);
  if (lane && lane.redo.length) {
    lane.redo.length = 0;
    notify(lane);
  }
}

/** Whether a write is queued or running. */
export function pdfMarkupWritesPending(key: string): boolean {
  const lane = lanes.get(key);
  return Boolean(lane && (lane.running || lane.waiting.length));
}

/** Whether a mark (not a history step) is queued or being written: an Undo
 * has something to take back. */
export function pdfMarkupMarksPending(key: string): boolean {
  const lane = lanes.get(key);
  if (!lane) return false;
  const isMark = (write: QueuedWrite | null) => Boolean(write && !write.historyStep);
  return isMark(lane.running) || lane.waiting.some(isMark);
}

 
export function pdfMarkupHistoryStepsPending(key: string): boolean {
  const lane = lanes.get(key);
  return Boolean(lane && [lane.running, ...lane.waiting].some((write) => write?.historyStep));
}

/** Marks pressed and not yet being written (an Undo takes the last at once). */
export function pdfMarkupQueuedCount(key: string): number {
  return lanes.get(key)?.waiting.filter((write) => !write.historyStep).length ?? 0;
}

/** Withdrawn marks Redo can bring back. */
export function pdfMarkupWithdrawnCount(key: string): number {
  return lanes.get(key)?.redo.filter((entry) => entry.kind === "mark").length ?? 0;
}

/** Settles when every queued write has run (written or refused). */
export function whenPdfMarkupWritesSettled(key: string): Promise<void> {
  const lane = lanes.get(key);
  return lane && (lane.running || lane.waiting.length) ? lane.drained : Promise.resolve();
}

export function subscribePdfMarkupWrites(key: string, listener: () => void): () => void {
  const lane = laneOf(key);
  lane.listeners.add(listener);
  return () => {
    lane.listeners.delete(listener);
    if (!lane.listeners.size && !lane.running && !lane.waiting.length && !lane.redo.length) lanes.delete(key);
  };
}

/** Tests only. */
export function resetPdfMarkupWriteQueues(): void {
  lanes.clear();
  tickets.clear();
}
