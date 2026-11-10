/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one order a box's Leave tries its words and format in when the writer
 * refuses the format (docs/pdf/text-edit-unified-model.md §2.2): the format
 * asked; the last format the writer took on the box; the words alone, then
 * the format on the words as they now stand. Both Leaves (the editor's and
 * the detached) had an order of their own: the editor's never went back to
 * the format last taken, the detached's lost nothing but differed in which
 * refusals it said.
 */

export type PdfNativeTextFormatLadderStep = "format" | "fallback" | "words" | "format-on-words";

export type PdfNativeTextFormatLadderResult = Readonly<{
  /** The step whose write went in last. */
  landed: PdfNativeTextFormatLadderStep;
  /** The format asked went in (on its own or on the words): else it is said
   * as not applied. */
  formatApplied: boolean;
}>;

export async function stagePdfNativeTextFormatLadder<Format>(input: Readonly<{
  format: Format;
  /** The last format the writer took on this box, or null. */
  fallback: Format | null;
  /** The words differ from the box's: they go in alone when no format does. */
  wordsChanged: boolean;
  /** Stages the words with a format (null: the words alone); throws when
   * the writer refuses it. */
  stage: (format: Format | null) => Promise<void>;
  /** Sets the format on the words already staged; false when refused. */
  stageOnWords: (format: Format) => Promise<boolean>;
  /** The format was already tried (and refused) by the caller. */
  formatTried?: boolean;
  mark?: (step: PdfNativeTextFormatLadderStep, error?: unknown) => void;
}>): Promise<PdfNativeTextFormatLadderResult> {
  let refusal: unknown = new Error("The format was refused.");
  if (!input.formatTried) {
    try {
      await input.stage(input.format);
      return Object.freeze({ landed: "format", formatApplied: true });
    } catch (error) {
      refusal = error;
      input.mark?.("format", error);
    }
    if (input.fallback !== null) {
      try {
        await input.stage(input.fallback);
        return Object.freeze({ landed: "fallback", formatApplied: false });
      } catch (error) {
        refusal = error;
        input.mark?.("fallback", error);
      }
    }
  }
  if (!input.wordsChanged) throw refusal;
  await input.stage(null);
  input.mark?.("words");
  const formatted = await input.stageOnWords(input.format);
  if (formatted) input.mark?.("format-on-words");
  return Object.freeze({ landed: formatted ? "format-on-words" : "words", formatApplied: formatted });
}
