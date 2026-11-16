/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Background saves of PDF text edits, in order.
 *
 * A text edit that has to reach the server before it is page content (a
 * wrapped or resized source line, written as a managed text box) used to keep
 * its editor open for the whole round trip -- seconds, and tens of seconds on
 * a loaded server -- and every other line refused to open meanwhile. As in
 * WPS, the edit now lands at once and is written behind the reader:
 *
 * - `start` runs one save and resolves `true` the moment the save reports
 *   its draft published locally (the editor may close then), or with the
 *   save's own result when it settles before publishing (a refusal the
 *   editor still shows, its draft intact).
 * - Saves run one after another: a second save starts only once the first
 *   has settled, so two edits in a row queue instead of refusing each other.
 * - `enqueue` parks a request (a line clicked while a save is in flight) until
 *   every save in flight has settled, and tells it whether the last one held.
 *
 * Handlers never throw into the queue: a failing handler cannot stall the
 * saves behind it.
 */
export type PdfTextBackgroundSaveJob = (published: () => void) => Promise<boolean>;

export interface PdfTextBackgroundSaveHandlers {
  /** The draft is on the page; the editor may let go of it. */
  readonly onPublished?: () => void;
  /** The save is over. `published` says whether the editor had let go. */
  readonly onSettled?: (saved: boolean, published: boolean) => void;
}

export interface PdfTextBackgroundSaveQueue {
  /** Saves started and not yet settled (queued ones included). */
  readonly pending: number;
  start(job: PdfTextBackgroundSaveJob, handlers?: PdfTextBackgroundSaveHandlers): Promise<boolean>;
  /** Parks `waiter` until no save is in flight; false (not parked) when none is. */
  enqueue(waiter: (saved: boolean) => void): boolean;
}

export function createPdfTextBackgroundSaveQueue(
  /** Told how many saves are started and not settled, each time that
   * changes: what the window's save status reads while one is behind it. */
  onPendingChange?: (pending: number) => void,
): PdfTextBackgroundSaveQueue {
  let tail: Promise<void> = Promise.resolve();
  let pending = 0;
  const waiters: Array<(saved: boolean) => void> = [];
  const call = (handler: (() => void) | undefined) => {
    try { handler?.(); } catch { /* a handler cannot stall the queue */ }
  };
  return {
    get pending() { return pending; },
    start(job, handlers = {}) {
      pending += 1;
      call(() => onPendingChange?.(pending));
      return new Promise<boolean>((resolve) => {
        const run = async () => {
          let published = false;
          let saved = false;
          try {
            saved = await job(() => {
              if (published) return;
              published = true;
              call(handlers.onPublished);
              resolve(true);
            });
          } catch {
            saved = false;
          }
          pending -= 1;
          call(() => onPendingChange?.(pending));
          if (!published) resolve(saved);
          call(() => handlers.onSettled?.(saved, published));
          if (pending === 0) {
            for (const waiter of waiters.splice(0)) call(() => waiter(saved));
          }
        };
        tail = tail.then(run, run);
      });
    },
    enqueue(waiter) {
      if (pending === 0) return false;
      waiters.push(waiter);
      return true;
    },
  };
}
