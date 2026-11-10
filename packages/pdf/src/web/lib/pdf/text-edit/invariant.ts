/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The text edit model's invariants, watched in development
 * (docs/pdf/text-edit-unified-model.md §2.4): what a box's commit stages is
 * what its last preview showed, in one step. A breach is said in the console
 * as `[pdf-edit-invariant]`, which the sweeps collect; nothing is changed.
 */

type Preview = Readonly<{ text: string; outputSha256: string }>;

/** Each box's last preview (by the editor's key for the box), a few boxes back. */
const previews = new Map<string, Preview>();

/** A preview of `text` in the box `key` was drawn from these candidate bytes. */
export function notePdfNativeTextPreview(key: string, text: string, outputSha256: string): void {
  if (process.env.NODE_ENV === "production") return;
  previews.delete(key);
  previews.set(key, Object.freeze({ text, outputSha256 }));
  while (previews.size > 16) previews.delete(previews.keys().next().value!);
}








export function forgetPdfNativeTextPreview(key: string | null | undefined): void {
  if (process.env.NODE_ENV === "production" || !key) return;
  previews.delete(key);
}

/**
 * A commit of `text` in the box `key` staged `steps` operations and left the
 * working copy at `workingSha256`: it should be the last preview's candidate
 * (when that preview was of the same text). Its steps are one Undo press
 * (pressNativeWorkingOperations); they are said with a breach.
 */
export function checkPdfNativeTextCommit(input: Readonly<{
  key: string | null | undefined;
  text: string;
  workingSha256: string;
  steps: number;
  lane: string;
}>): void {
  if (process.env.NODE_ENV === "production" || !input.key) return;
  const preview = previews.get(input.key);
  const differs = preview !== undefined && preview.text === input.text && preview.outputSha256 !== input.workingSha256;
  if (!differs) return;
  try {
    console.warn("[pdf-edit-invariant] " + JSON.stringify({ lane: input.lane, steps: input.steps,
      previewed: preview?.text === input.text, differs,
      preview: preview?.outputSha256.slice(0, 8) ?? null, staged: input.workingSha256.slice(0, 8) }));
  } catch { /* diagnostics never throw */ }
}
