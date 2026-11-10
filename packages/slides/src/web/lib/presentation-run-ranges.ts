/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Replacing a character range across text runs, apart from the deck's Find and Replace (presentation-find-replace re-exports it): the spreadsheet's spelling walk needs only this, and the deck's command brought masters, themes and layouts with it.
 */
import { type TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";

/** A single styled run (mirrors `lib/presentation-runs.ts` `TextRun`, structurally). */
export interface RunLike {
  text: string;
  style?: unknown;
}

/**
 * Replace the characters `[start, end)` of the runs' concatenated text with
 * `replacement`. The replacement joins the run holding the first replaced
 * character (so it takes that character's formatting, as PowerPoint's
 * `TextRange.Replace` does), every character outside the range keeps its own
 * run, and a run the range empties is dropped. Untouched runs are returned by
 * reference. Pure.
 */
export function replaceRangeInRuns<T extends RunLike>(
  runs: readonly T[],
  start: number,
  end: number,
  replacement: string,
): T[] {
  const out: T[] = [];
  let offset = 0;
  let inserted = false;
  for (const run of runs) {
    const text = typeof run.text === "string" ? run.text : "";
    const runStart = offset;
    const runEnd = offset + text.length;
    offset = runEnd;
    if (runEnd <= start || runStart >= end) {
      out.push(run);
      continue;
    }
    const owns = !inserted && start >= runStart && start < runEnd;
    if (owns) inserted = true;
    const nextText =
      text.slice(0, Math.max(0, start - runStart)) +
      (owns ? replacement : "") +
      text.slice(Math.min(text.length, end - runStart));
    if (nextText) out.push(nextText === text ? run : { ...run, text: nextText });
  }
  return out;
}
