/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** How long a structural edit's save waits for its calculation. */
export const STRUCTURAL_SAVE_CALCULATION_WAIT_MS = 10_000;

let timeouts = 0;

/** Structural saves that went ahead of their calculation (this session). */
export function structuralSaveCalculationTimeouts(): number {
  return timeouts;
}

/** The calculation barrier, but no longer than `remainingMs`: false when the
 * wait ran out (counted, and said in development). */
export async function structuralSaveBarrier(barrier: Promise<boolean>, remainingMs: number): Promise<boolean> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const expired = new Promise<"expired">((resolve) => {
    timer = setTimeout(() => resolve("expired"), Math.max(0, remainingMs));
  });
  try {
    const outcome = await Promise.race([barrier, expired]);
    if (outcome !== "expired") return outcome;
    timeouts += 1;
    if (process.env.NODE_ENV !== "production") {
      try {
        console.warn("[spreadsheet-save] structural save went ahead of its calculation", { count: timeouts });
      } catch {
        /* diagnostics never throw */
      }
    }
    return false;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}
