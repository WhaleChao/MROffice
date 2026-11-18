/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type Refusal = { label: string; at: number; shown: boolean };

const refusals = new Map<string, Refusal>();

/** The window: a write of this file was refused by the file's policy. */
export function notePdfSourcePolicyRefusal(dataItemId: string, label: string, now: number = Date.now()): void {
  if (!dataItemId || !label) return;
  refusals.set(dataItemId, { label, at: now, shown: false });
  while (refusals.size > 32) refusals.delete(refusals.keys().next().value!);
}

/**
 * The reason a write of this file was refused within the last `withinMs`,
 * for a notice shown in its place; null when none was. `showing` tells the
 * window the reason is on screen (pdfSourcePolicyRefusalShown).
 */
export function pdfSourcePolicyRefusal(dataItemId: string | null | undefined, withinMs: number,
  options?: Readonly<{ showing?: boolean; now?: number }>): string | null {
  const refusal = dataItemId ? refusals.get(dataItemId) : undefined;
  const now = options?.now ?? Date.now();
  if (!refusal || now - refusal.at > withinMs) return null;
  if (options?.showing) refusal.shown = true;
  return refusal.label;
}

/** Whether a notice of the reader's has shown the reason of the refusal
 * the window noted at `at` (the window then says nothing of its own). */
export function pdfSourcePolicyRefusalShown(dataItemId: string, at: number): boolean {
  const refusal = refusals.get(dataItemId);
  return Boolean(refusal && refusal.at === at && refusal.shown);
}
