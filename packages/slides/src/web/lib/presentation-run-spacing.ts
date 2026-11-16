/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Character Spacing on the characters of a slide's text box: a run's own
 * `letterSpacing` (points, DrawingML `a:rPr@spc`), as PowerPoint's Home >
 * Character Spacing and the Font dialog's Spacing page act on the selected
 * characters (a caret: the word it is in), and on every character when the
 * box itself is selected.
 *
 * A run's own spacing is what its characters wear; a run without one wears
 * the box's (`style.letterSpacing`). Kept canonical as the WordArt paint is
 * (`presentation-run-text-paint`): a spacing every character wears is the
 * box's, and a box whose characters differ keeps theirs on the runs.
 */

import type {
  PresentationBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  formatTextBlockRunSelection,
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  presentationEditableText,
  selectedTextRange,
  type LiveTextSelection,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import { normalizeRuns, runsFromText, type RunStyle, type TextRun } from "@mrnote/core/web/lib/presentation-point-runs.ts";

const own = (block: PresentationTextBlock): number => {
  const value = (block.style as { letterSpacing?: unknown } | undefined)?.letterSpacing;
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
};

function rows(block: PresentationTextBlock): { runs: TextRun[]; start: number }[] {
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    let start = 0;
    return paragraphs.map((paragraph) => {
      const row = { runs: normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text), start };
      start += paragraph.text.length + 1;
      return row;
    });
  }
  return [{ runs: normalizeRuns(block.runs) ?? runsFromText(block.text ?? ""), start: 0 }];
}

/** The block's runs rewritten by `map`, where the block draws its words. */
function withEveryRun(block: PresentationTextBlock, map: (run: TextRun) => TextRun): PresentationTextBlock {
  const next = { ...block } as PresentationTextBlock & { runs?: TextRun[] };
  const rewrite = (runs: readonly TextRun[]) => normalizeRuns(runs.map(map));
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs) && Array.isArray(block.paragraphs)) {
    next.paragraphs = (block.paragraphs as unknown[]).map((record, index) => {
      const paragraph = paragraphs[index];
      if (!paragraph || !record || typeof record !== "object") return record;
      const runs = rewrite(normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text));
      const row = { ...(record as Record<string, unknown>) };
      if (runs) row.runs = runs;
      else delete row.runs;
      return row;
    }) as PresentationTextBlock["paragraphs"];
    return next;
  }
  const runs = rewrite(normalizeRuns(block.runs) ?? runsFromText(block.text ?? ""));
  if (runs) next.runs = runs;
  else delete next.runs;
  return next;
}

function withRunSpacing(run: TextRun, points: number | undefined): TextRun {
  const style = { ...(run.style ?? {}) } as RunStyle;
  if (points === undefined) delete style.letterSpacing;
  else style.letterSpacing = points;
  return Object.keys(style).length ? { ...run, style } : { ...run, style: undefined };
}

function withBoxSpacing(block: PresentationTextBlock, points: number): PresentationTextBlock {
  const style = { ...((block.style ?? {}) as Record<string, unknown>) };
  if (points === 0) delete style.letterSpacing;
  else style.letterSpacing = points;
  const next = { ...block } as PresentationTextBlock & { style?: unknown };
  if (Object.keys(style).length) next.style = style as PresentationTextBlock["style"];
  else delete next.style;
  return next;
}

/** Whether any run states a spacing of its own. */
export function presentationBlockHasRunSpacing(block: PresentationBlock | null | undefined): boolean {
  if (!block || block.type !== "text") return false;
  return rows(block as PresentationTextBlock).some((row) =>
    row.runs.some((run) => typeof run.style?.letterSpacing === "number"),
  );
}

/**
 * The spacing the characters a command acts on wear (points), or undefined
 * when they differ. `selection` null: every character.
 */
export function presentationLetterSpacingAt(
  block: PresentationBlock | null | undefined,
  selection: LiveTextSelection,
): number | undefined {
  if (!block || block.type !== "text") return undefined;
  const textBlock = block as PresentationTextBlock;
  const text = presentationEditableText(textBlock);
  let [from, to] = selection && selection.blockId === block.id
    ? selectedTextRange(block.id, text, selection)
    : [0, text.length];
  if (from === to && selection && selection.blockId === block.id) {
    from = Math.max(0, from - 1);
    to = Math.max(from, Math.min(text.length, from + 1));
  }
  if (from === to) return own(textBlock);
  let found: number | undefined;
  let seen = false;
  for (const row of rows(textBlock)) {
    let at = row.start;
    for (const run of row.runs) {
      const start = at;
      at += run.text.length;
      if (!run.text.length || at <= from || start >= to) continue;
      const value = run.style?.letterSpacing ?? own(textBlock);
      if (!seen) {
        found = value;
        seen = true;
      } else if (value !== found) return undefined;
    }
  }
  return seen ? found : own(textBlock);
}

/**
 * The block with Character Spacing `points` on the characters a command acts
 * on: the selected ones while the words are typed in, else every character.
 * Null when nothing changes (a bare caret formats no existing character).
 */
export function presentationBlockWithLetterSpacingIn(
  block: PresentationBlock,
  selection: LiveTextSelection,
  points: number,
): PresentationBlock | null {
  // A text box's words, or a shape's (the same text body in PowerPoint).
  if (block.type !== "text" && block.type !== "shape") return null;
  const textBlock = block as PresentationTextBlock;
  const typed = selection !== null && selection.blockId === block.id;
  const text = presentationEditableText(textBlock);
  if (!typed || text === "") {
    // The box selected (or no words yet): the box's own, every run's off.
    const whole = withBoxSpacing(withEveryRun(textBlock, (run) => withRunSpacing(run, undefined)), points);
    return JSON.stringify(whole) === JSON.stringify(textBlock) ? null : (whole as PresentationBlock);
  }
  // The box's own onto every run without one, so the rest keeps it.
  const boxPoints = own(textBlock);
  const spread = withBoxSpacing(
    withEveryRun(textBlock, (run) =>
      typeof run.style?.letterSpacing === "number" || boxPoints === 0 ? run : withRunSpacing(run, boxPoints),
    ),
    0,
  );
  const formatted = formatTextBlockRunSelection(spread, block.id, selection, { letterSpacing: points });
  if (!formatted) return null;
  return presentationBlockWithCanonicalSpacing(formatted) as PresentationBlock;
}

/** A spacing every character wears is the box's; otherwise the runs keep theirs. */
export function presentationBlockWithCanonicalSpacing(block: PresentationTextBlock): PresentationTextBlock {
  const values: number[] = [];
  for (const row of rows(block)) {
    for (const run of row.runs) {
      if (run.text.length) values.push(run.style?.letterSpacing ?? own(block));
    }
  }
  if (!values.length || values.some((value) => value !== values[0])) return block;
  return withBoxSpacing(withEveryRun(block, (run) => withRunSpacing(run, undefined)), values[0]);
}
