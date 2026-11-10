/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The refusal for an edit, made in the reader's order, of text the reader
 * shows in another order than the page draws it, where the edit cannot be
 * placed in the drawn order: written back as typed it would reverse letters
 * the edit never touched. */
export const MRAI_NATIVE_REORDERED_TEXT_MESSAGE =
  "The selected text is shown in another order than the page draws it; this edit cannot be placed in the drawn order without rewriting text it does not touch.";

// Hebrew through Arabic Extended-A, the right-to-left blocks of the
// supplementary planes, and the Hebrew and Arabic presentation forms: the
// characters the bidi pass reorders a run around.
const RIGHT_TO_LEFT_CHARACTER =
  /[֐-ࣿיִ-﷿ﹰ-﻿\u{10800}-\u{10FFF}\u{1E800}-\u{1EFFF}]/u;

/** Whether the reader reorders some run of `text`. */
export function pdfTextHasRightToLeftRun(text: string): boolean {
  return RIGHT_TO_LEFT_CHARACTER.test(text);
}
