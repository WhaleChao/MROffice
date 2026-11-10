/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_NATIVE_DETACH_REFUSAL_MEMORY_MS = 120_000;

/** Whether the line is finished again; records the refusal in `refusals`
 * (keyed by its page, the words before it and the words typed). */
export function pdfNativeDetachFinishesAgain(
  refusals: Map<string, number>,
  input: Readonly<{ pageNumber: number; wordsBefore: string; words: string; unsupported: boolean; now: number }>,
): boolean {
  for (const [key, until] of refusals) if (until < input.now) refusals.delete(key);
  const key = `${input.pageNumber}\u0000${input.wordsBefore}\u0000${input.words}`;
  const again = refusals.has(key);
  refusals.set(key, input.now + PDF_NATIVE_DETACH_REFUSAL_MEMORY_MS);
  while (refusals.size > 64) refusals.delete(refusals.keys().next().value!);
  return !again && !input.unsupported;
}
