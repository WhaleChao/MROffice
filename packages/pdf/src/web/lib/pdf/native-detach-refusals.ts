/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfEngineResult } from "./engine/contracts";
import { isMraiNativeListTextMissingGlyphRefusal } from "./engine/native/list-text-model";
import { PdfNativeTransactionError } from "./engine/native/transaction-error";

export const PDF_NATIVE_DETACH_REFUSAL_MEMORY_MS = 120_000;

/** A refusal's detail comes from the compiles' Worker cut to this many
 * characters (native-text-edit-preparation.worker.ts): one as long as this
 * may have lost its end. */
const PDF_NATIVE_COMPILE_REFUSAL_DETAIL_CHARACTERS = 720;

















export function pdfNativePlanAnswerOfCompileRefusal(error: unknown): "missing-glyph" | "refused" | null {
  if (!(error instanceof PdfNativeTransactionError) || error.code !== "unsupported-document") return null;
  if (isMraiNativeListTextMissingGlyphRefusal(error.message)) return "missing-glyph";
  return error.message.length < PDF_NATIVE_COMPILE_REFUSAL_DETAIL_CHARACTERS ? "refused" : null;
}

/** Whether the window's finish takes the line once more (false: the line's
 * own editor finishes it); records the refusal in `refusals` (keyed by its
 * page, the words before it and the words typed). */
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











export function pdfNativeCarriedWordsRefused(input: Readonly<{
  /** The finish was the editor's own (PdfNativeTextQueuedDraft.finishInEditor). */
  inEditor: boolean;
  /** Its last verdict; null when it was not asked a last time. */
  written: boolean | null;
  /** The words it was asked to finish. */
  words: string;
  /** The editor open on the line now (null: none). */
  open: Readonly<{ dirty: boolean; replacementText: string; finishing: boolean }> | null;
}>): boolean {
  return input.inEditor && input.written === false && input.open !== null && input.open.dirty &&
    input.open.replacementText === input.words && !input.open.finishing;
}

/**
 * Whether what the planner answered instead of a plan is its answer about
 * the edit on these bytes, the same however often it is asked: the object is
 * not there what the edit names, a command it does not take, a result it
 * could not verify, a document it does not write. Not one of the moment: a
 * source or a revision that moved under the plan (the next plan stands on
 * the new one), a session not open, or the adapter's own refusals, which
 * include a font it could not load and a file it could not checkpoint.
 */
export function pdfNativePlanRefusalIsForGood(result: PdfEngineResult): boolean {
  if (result.status === "conflict") return result.reason === "object-precondition-mismatch";
  return result.status === "refused" && (result.reason === "invalid-command" ||
    result.reason === "verification-failed" || result.reason === "unsupported-document");
}










export function pdfNativeLeaveRefusalNotice(input: Readonly<{
  /** The planner's answer for good (pdfNativePlanRefusalIsForGood). */
  refusedForGood: boolean;
  /** The draft changes the line's format ... */
  formatChanged: boolean;
  /** ... and its words. */
  wordsChanged: boolean;
}>): "format-not-applied" | "try-again" {
  return input.refusedForGood && input.formatChanged && !input.wordsChanged ? "format-not-applied" : "try-again";
}
