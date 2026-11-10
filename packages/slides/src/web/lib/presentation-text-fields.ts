/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Slide Number and Insert > Date & Time with the cursor in text:
 * PowerPoint puts a field at the cursor -- `a:fld type="slidenum"`, or the
 * Date and Time dialog's format as `a:fld type="datetimeN"` when Update
 * Automatically is on (plain text when it is off) -- the same DrawingML
 * fields the Header and Footer placeholders hold (PPT-A052). With no cursor
 * in text both open Header and Footer instead.
 *
 * The run model already carries a field's identity (`TextRun.field`): the
 * canvas draws a slide-number field as the slide's number and the writers
 * put the `a:fld` back.
 */

import type { PresentationTextBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  formatDeckDate,
  type PresentationDateFieldFormat,
  type PresentationDateFieldLocale,
} from "@mrnote/core/web/lib/presentation-header-footer.ts";
import { textFromParagraphs } from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import {
  // Sizes to the hundredth (DrawingML's): whole px turned 10.5 pt into 11.
  applyRunField,
  normalizeRuns,
  paragraphRunSelectionRanges,
  reconcileTextRunsAfterEdit,
  textEditDelta,
} from "@mrnote/core/web/lib/presentation-point-runs.ts";
import {
  runsFromText,
  textFromRuns,
  type RunField,
} from "@mrnote/core/web/lib/presentation-runs.ts";
import {
  reconcileParagraphsAfterTextEdit,
  reconcileStructuralParagraphsAfterTextEdit,
} from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import {
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";

/** A fresh `a:fld@id`: a braced GUID, as PowerPoint writes it. */
export function newRunFieldId(): string {
  return `{${crypto.randomUUID().toUpperCase()}}`;
}

export function slideNumberField(): RunField {
  return { type: "slidenum", id: newRunFieldId() };
}

/** What Insert > Date & Time puts at the cursor: the text, and the field when it updates. */
export function dateTimeInsertion(
  format: PresentationDateFieldFormat,
  locale: PresentationDateFieldLocale,
  updateAutomatically: boolean,
  now: Date,
): { text: string; field: RunField | null } {
  return {
    text: formatDeckDate({ mode: "auto", format, locale }, now),
    field: updateAutomatically ? { type: format, id: newRunFieldId() } : null,
  };
}

/**
 * The text block with `[start, end)` replaced by `text`, and that text one
 * field run when `field` is given. Null when the text cannot hold it (a line
 * break in the inserted text, an out-of-range selection).
 */
export function textBlockWithInsertedField(
  block: PresentationTextBlock,
  start: number,
  end: number,
  text: string,
  field: RunField | null,
): PresentationTextBlock | null {
  const current = block.text ?? "";
  if (start < 0 || end < start || end > current.length || !text || /[\r\n\v]/.test(text)) {
    return null;
  }
  const nextText = current.slice(0, start) + text + current.slice(end);
  const selection = { start, end };
  const existing = normalizeParagraphsWithRuns(block.paragraphs);
  const range = { from: start, to: start + text.length };
  if (existing && ownsParagraphRichLayout(existing)) {
    const paragraphs = reconcileParagraphsAfterTextEdit(existing, nextText, selection);
    if (!field) return { ...block, text: textFromParagraphs(paragraphs), paragraphs };
    const ranges = paragraphRunSelectionRanges(paragraphs, range.from, range.to);
    if (ranges.length !== 1) return null;
    const [only] = ranges;
    const marked = paragraphs.map((paragraph, index) => {
      if (index !== only.paragraphIndex) return paragraph;
      const runs = normalizeRuns(applyRunField(only.runs, only.from, only.to, field));
      return runs ? { ...paragraph, text: textFromRuns(runs), runs } : paragraph;
    });
    const next = { ...block, text: textFromParagraphs(marked), paragraphs: marked };
    delete (next as { runs?: unknown }).runs;
    return next;
  }
  const delta = textEditDelta(current, nextText, selection);
  const runs = reconcileTextRunsAfterEdit(current, nextText, block.runs, delta);
  const next: PresentationTextBlock = { ...block, text: nextText };
  if (existing) next.paragraphs = reconcileStructuralParagraphsAfterTextEdit(existing, nextText, delta);
  const base = runs ?? runsFromText(nextText);
  const marked = field ? normalizeRuns(applyRunField(base, range.from, range.to, field)) : normalizeRuns(base);
  if (marked) next.runs = marked;
  else delete next.runs;
  return next;
}
