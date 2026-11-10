/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** What a workbook snapshot refuses while the workbook is still changing:
 * its calculation had not settled, its sheets changed while it waited, or a
 * cell entry was still being composed (an IME). */
export const SPREADSHEET_SETTLING_ERRORS: ReadonlySet<string> = new Set([
  "spreadsheet_calculation_not_ready",
  "spreadsheet_source_changed",
  "spreadsheet_draft_composing",
]);

/** A wait that outlived the export's bound (the take itself never answered). */
const TAKE_TIMED_OUT = "spreadsheet_snapshot_timed_out";

/**
 * Take a workbook snapshot once the workbook holds still.
 *
 * A workbook just opened keeps changing for a moment: a shared one's sheets
 * arrive shard by shard and recalculate. A snapshot that met a change used to
 * fail the export at once ("无法导出 XLSX", nothing sent). It is now taken
 * again, a little later each time (150 ms up to 1 s), until it succeeds or the
 * bound passes; any other failure is the caller's at once.
 *
 * The bound holds for every take too: a take still waiting when it passes (a
 * calculation, a filter or a validation that never answers, a frame a hidden
 * tab never paints) fails the snapshot instead of holding the export -- and
 * every export after it -- forever. `stillSettling` is told the last code it
 * met (null when the take itself timed out).
 */
export async function snapshotWhenWorkbookSettles<T>(
  take: () => Promise<T>,
  options: {
    stillSettling: (lastCode: string | null) => Error;
    timeoutMs?: number;
    now?: () => number;
    sleep?: (ms: number) => Promise<void>;
  },
): Promise<T> {
  const now = options.now ?? Date.now;
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  const deadline = now() + (options.timeoutMs ?? 20_000);
  let lastCode: string | null = null;
  for (let delay = 150; ; delay = Math.min(delay * 2, 1_000)) {
    try {
      return await raceDeadline(take(), Math.max(0, deadline - now()));
    } catch (error) {
      if (error instanceof Error && error.message === TAKE_TIMED_OUT) throw options.stillSettling(lastCode);
      if (!(error instanceof Error) || !SPREADSHEET_SETTLING_ERRORS.has(error.message)) throw error;
      lastCode = error.message;
      if (now() + delay > deadline) throw options.stillSettling(lastCode);
      await sleep(delay);
    }
  }
}

function raceDeadline<T>(work: Promise<T>, remainingMs: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const expired = new Promise<never>((_resolve, reject) => {
    timer = setTimeout(() => reject(new Error(TAKE_TIMED_OUT)), remainingMs);
  });
  // A take that answers after the bound is let go; its refusal is not unhandled.
  work.catch(() => undefined);
  return Promise.race([work, expired]).finally(() => clearTimeout(timer));
}

/**
 * The next frame, or a short timer while the page is hidden: a background
 * tab (or an iPad app sent to the background) paints no frames, so a wait on
 * one never ended and the export it held never ran.
 */
export function nextPaintOrTimer(): Promise<void> {
  return new Promise<void>((resolve) => {
    if (typeof document !== "undefined" && document.hidden) {
      setTimeout(resolve, 16);
      return;
    }
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      resolve();
    };
    requestAnimationFrame(finish);
    // A tab hidden after the request: its frame may never come.
    setTimeout(finish, 250);
  });
}
