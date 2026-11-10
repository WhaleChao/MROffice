/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type SettledWorkbook =
  /** Nothing is changing now and the calculation answered: run. */
  | "settled"
  /** The person moved on, or the window closed: the command no longer applies. */
  | "moved-on"
  /** The calculation failed on this workbook, or never settled: say so. */
  | "refused";

/** Past this a command is refused rather than run long after it was asked for. */
export const SETTLED_WORKBOOK_LIMIT_MS = 120_000;

export async function waitForSettledWorkbook<T>({
  live,
  barrier,
  stillWanted,
  nextFrame,
  now = () => Date.now(),
  limitMs = SETTLED_WORKBOOK_LIMIT_MS,
}: {
  live: () => T;
  barrier: () => Promise<boolean>;
  stillWanted: () => boolean;
  nextFrame: () => Promise<void>;
  now?: () => number;
  limitMs?: number;
}): Promise<SettledWorkbook> {
  const started = now();
  for (;;) {
    const source = live();
    const ready = await barrier();
    if (!stillWanted()) return "moved-on";
    if (source === live()) {
      if (!ready) return "refused";
      // The answer's projection is published before the command reads it.
      await nextFrame();
      if (!stillWanted()) return "moved-on";
      if (source === live()) return "settled";
    }
    if (now() - started >= limitMs) return "refused";
    await nextFrame();
  }
}
