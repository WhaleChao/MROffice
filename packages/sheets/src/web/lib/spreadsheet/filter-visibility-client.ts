/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { filterVisibilitySteps, type PreparedFilterVisibility } from "./filter-visibility";
import { yieldCalculationTask, BackgroundCalculationCancelledError } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";

export async function drainFilterVisibility<T>(steps: Generator<void, T>, isCurrent: () => boolean, yieldTask = yieldCalculationTask): Promise<T> {
  for (;;) {
    if (!isCurrent()) throw new BackgroundCalculationCancelledError();
    const next = steps.next();
    if (next.done) return next.value;
    await yieldTask();
  }
}

/** A separate visibility job must not cancel the formula client's pending slot. */
export class FilterVisibilityClient {
  private worker: Worker | null = null;
  private reject: ((error: Error) => void) | null = null;
  private sequence = 0;
  cancel() {
    this.sequence++;
    this.worker?.terminate(); this.worker = null;
    this.reject?.(new BackgroundCalculationCancelledError()); this.reject = null;
  }
  async calculate(input: PreparedFilterVisibility, isCurrent: () => boolean): Promise<number[]> {
    this.cancel(); const sequence = this.sequence;
    const current = () => this.sequence === sequence && isCurrent();
    await yieldCalculationTask();
    if (!current()) throw new BackgroundCalculationCancelledError();
    if (typeof Worker === "undefined") return drainFilterVisibility(filterVisibilitySteps(input), current);
    try {
      return await new Promise<number[]>((resolve, reject) => {
        const worker = new Worker(new URL("./filter-visibility.worker.ts", import.meta.url), {type:"module", name:"mrnote-spreadsheet-filter-visibility"});
        this.worker = worker;
        const timeout = setTimeout(() => finish(new Error("Filter worker timed out.")), 30_000);
        const cancellation = setInterval(() => { if (!current()) finish(new BackgroundCalculationCancelledError()); }, 16);
        const finish = (error: Error | null, rows?: number[]) => {
          clearTimeout(timeout); clearInterval(cancellation); worker.terminate();
          if (this.worker === worker) { this.worker = null; this.reject = null; }
          if (error) reject(error); else resolve(rows!);
        };
        this.reject = error => finish(error);
        worker.onerror = () => finish(new Error("Filter worker unavailable."));
        worker.onmessageerror = () => finish(new Error("Filter worker response unavailable."));
        worker.onmessage = event => {
          if (!current()) return finish(new BackgroundCalculationCancelledError());
          const rows: unknown = event.data?.rows;
          if (!Array.isArray(rows) || new Set(rows).size !== rows.length || rows.some(row => !Number.isSafeInteger(row) || !input.some(group => row >= group.first && row <= group.last))) return finish(new Error("Invalid filter visibility result."));
          finish(null, rows);
        };
        try { worker.postMessage(input); } catch (error) { finish(error instanceof Error ? error : new Error(String(error))); }
      });
    } catch (error) {
      if (!current() || error instanceof BackgroundCalculationCancelledError) throw error;
      return drainFilterVisibility(filterVisibilitySteps(input), current);
    }
  }
}
