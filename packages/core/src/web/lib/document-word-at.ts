/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ResolvedPos } from "@tiptap/pm/model";

export type WordSegment = { index: number; segment: string; isWordLike?: boolean };
type WordSegmenter = new (
  locales?: string | string[],
  options?: { granularity: "word" },
) => { segment: (input: string) => Iterable<WordSegment> };

const OBJECT_REPLACEMENT = "￼";

/**
 * A text's words and the runs between them: `Intl.Segmenter`'s words, so
 * Chinese and Japanese segment into words, else letter/digit runs. Shared by
 * every editor that picks "the word at the caret" (Docs here, the deck's
 * `caretWordRange`).
 */
export function wordSegmentsOf(text: string): WordSegment[] {
  const Segmenter = (Intl as unknown as { Segmenter?: WordSegmenter }).Segmenter;
  if (Segmenter) {
    try {
      return Array.from(new Segmenter(undefined, { granularity: "word" }).segment(text));
    } catch {
      // Fall through to the letter/digit runs below.
    }
  }
  const out: WordSegment[] = [];
  for (const match of text.matchAll(/[\p{L}\p{N}_]+|[^\p{L}\p{N}_]+/gu)) {
    out.push({
      index: match.index ?? 0,
      segment: match[0],
      isWordLike: /[\p{L}\p{N}]/u.test(match[0]),
    });
  }
  return out;
}

/**
 * The word at a caret, as Word picks it for a command given no selection (a
 * new comment anchors to the word at the insertion point): the word the caret
 * is in, else the word it touches on its left, else on its right. Words come
 * from `Intl.Segmenter`, so Chinese and Japanese text segments into words too.
 * Returns document positions, or null when the caret touches no word (an empty
 * paragraph, a caret between two spaces).
 */
export function documentWordRangeAt(
  $pos: ResolvedPos,
): { from: number; to: number } | null {
  const parent = $pos.parent;
  if (!parent.isTextblock || parent.content.size === 0) return null;
  // One character per position: every leaf (hard break, inline atom) stands
  // for one replacement character, so string offsets map 1:1 onto positions.
  const text = parent.textBetween(0, parent.content.size, undefined, OBJECT_REPLACEMENT);
  if (text.length !== parent.content.size) return null;
  const offset = $pos.parentOffset;
  const words = wordSegmentsOf(text).filter(
    (segment) => segment.isWordLike && !segment.segment.includes(OBJECT_REPLACEMENT),
  );
  const pick =
    words.find((w) => w.index < offset && offset < w.index + w.segment.length) ??
    words.find((w) => w.index + w.segment.length === offset) ??
    words.find((w) => w.index === offset);
  if (!pick) return null;
  const start = $pos.start();
  return { from: start + pick.index, to: start + pick.index + pick.segment.length };
}
