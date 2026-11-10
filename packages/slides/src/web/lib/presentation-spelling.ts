/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Review ▸ Proofing ▸ Spelling for presentations (PPT-A098).
 *
 * PowerPoint checks the deck from the current slide on, stops at each word no
 * dictionary accepts and offers Ignore / Ignore All / Add / Change / Change
 * All. The checker, dictionaries, custom dictionary and Ignore All session are
 * the shared `lib/proofing` primitives the spreadsheet Spelling dialog uses;
 * the deck side is Find's: words are found on the lines Find searches and
 * changed through `replaceMatchInDeck` / `replaceAllInDeck`, so Spelling can
 * change exactly what Replace can, and nothing it cannot (locked tables,
 * charts that refuse a title).
 */

import {
  findSpellingIssue,
  type SpellingCheckContext,
} from "@mrnote/core/web/lib/proofing/spelling-checker.ts";
import { proofingDictionaryTag } from "@mrnote/core/web/lib/proofing/spelling-dictionaries.ts";
import {
  findMatchesInDeck,
  presentationFindLineKey,
  presentationSearchableTexts,
  type PresentationFindMatch,
  type PresentationSearchableText,
} from "./presentation-find-replace";

type AnyRecord = Record<string, unknown>;

export interface PresentationSpellingIssue {
  readonly word: string;
  /** The exact, replaceable occurrence (Find's coordinates). */
  readonly match: PresentationFindMatch;
  /** The dictionary that judged it (its language's, else the check's own). */
  readonly tag: string;
}

/** The language the check runs in when the caller names none. */
const DEFAULT_CHECK_TAG = "en-US";

/** Text that no dictionary judges reads as an object, never as a space. */
const OBJECT_REPLACEMENT = "\uFFFC";

/**
 * Per character of a searchable line, the dictionary that judges it: none
 * for text set "Don't check spelling or grammar" (a run's DrawingML
 * `noProof`, Set Proofing Language), none for text in a language MRAI ships
 * no dictionary for, else its proofing language's (the line's own, without
 * one). PowerPoint's Spelling check passes over the first two.
 */
function lineDictionaryTags(
  line: PresentationSearchableText,
  defaultTag: string,
): (string | null)[] {
  if (!line.runs) return Array.from({ length: line.text.length }, () => defaultTag);
  const tags: (string | null)[] = [];
  for (const run of line.runs) {
    const style = (run.style ?? {}) as { noProof?: unknown; proofingLanguage?: unknown };
    const language =
      typeof style.proofingLanguage === "string" && style.proofingLanguage
        ? style.proofingLanguage
        : null;
    const tag = style.noProof === true ? null : proofingDictionaryTag(language, defaultTag);
    const length = typeof run.text === "string" ? run.text.length : 0;
    for (let index = 0; index < length; index += 1) tags.push(tag);
  }
  return tags;
}

/** The dictionaries a deck's text asks for, the check's own first. */
export function presentationSpellingTags(
  slides: readonly AnyRecord[],
  defaultTag: string,
): string[] {
  const tags = new Set<string>([defaultTag]);
  for (const line of presentationSearchableTexts(slides)) {
    for (const tag of lineDictionaryTags(line, defaultTag)) if (tag) tags.add(tag);
  }
  return [...tags];
}

/** Whole-word, case-exact: the occurrence is this spelling of this word. */
export const PRESENTATION_SPELLING_FIND_OPTIONS = Object.freeze({
  matchCase: true,
  wholeWord: true,
});

/**
 * The next misspelled word from `startSlideIndex`, wrapping round the deck
 * once, skipping occurrences already passed over (`skipped` holds their Find
 * keys). Repeated-word findings are not offered: deleting a word is not a
 * Replace edit, so it could not go through the same writer.
 *
 * Each word is judged by the dictionary of its proofing language
 * (`contextFor`; a single context judges the check's own language,
 * `defaultTag`, only). Text set "Don't check spelling or grammar", and text
 * in a language without a dictionary, is passed over and reads as an object:
 * it neither joins the words beside it nor is offered as the occurrence of a
 * word misspelled elsewhere.
 */
export function findNextPresentationSpellingIssue(
  slides: readonly AnyRecord[],
  startSlideIndex: number,
  context: SpellingCheckContext | ((tag: string) => SpellingCheckContext | null),
  skipped: ReadonlySet<string>,
  defaultTag: string = DEFAULT_CHECK_TAG,
): PresentationSpellingIssue | null {
  const count = slides.length;
  if (count === 0) return null;
  const start = Math.min(Math.max(0, Math.floor(startSlideIndex) || 0), count - 1);
  const contextFor =
    typeof context === "function"
      ? context
      : (tag: string) => (tag === defaultTag ? context : null);
  const checkContexts = new Map<string, SpellingCheckContext | null>();
  const checkContextFor = (tag: string): SpellingCheckContext | null => {
    if (!checkContexts.has(tag)) {
      const own = contextFor(tag);
      checkContexts.set(
        tag,
        own ? { ...own, options: { ...own.options, flagRepeatedWords: false } } : null,
      );
    }
    return checkContexts.get(tag) ?? null;
  };
  const lines = presentationSearchableTexts(slides).map((line) => ({
    line,
    tags: lineDictionaryTags(line, defaultTag),
  }));
  // Which dictionary judges each character of each line, for the occurrences
  // Find offers: one in unchecked text, or in another language's, is not the
  // misspelling this dictionary found.
  const lineTags = new Map(lines.map(({ line, tags }) => [line.lineKey, tags]));
  const judgedBy = (candidate: PresentationFindMatch, tag: string): boolean => {
    const tags = lineTags.get(presentationFindLineKey(candidate));
    if (!tags) return tag === defaultTag;
    for (let index = candidate.start; index < candidate.end; index += 1) {
      if (tags[index] !== tag) return false;
    }
    return true;
  };
  const tried = new Set<string>();
  for (let step = 0; step < count; step += 1) {
    const slideIndex = (start + step) % count;
    for (const { line, tags } of lines) {
      if (line.slideIndex !== slideIndex) continue;
      for (const tag of new Set(tags)) {
        if (tag === null) continue;
        const checkContext = checkContextFor(tag);
        if (!checkContext) continue;
        // This dictionary's text, every other character an object.
        let masked = "";
        for (let index = 0; index < line.text.length; index += 1) {
          masked += tags[index] === tag ? line.text[index] : OBJECT_REPLACEMENT;
        }
        let from = 0;
        for (;;) {
          const issue = findSpellingIssue(masked, from, checkContext);
          if (!issue) break;
          from = issue.end;
          const key = `${slideIndex}\u0000${tag}\u0000${issue.word}`;
          if (tried.has(key)) continue;
          tried.add(key);
          const match = findMatchesInDeck(
            slides,
            issue.word,
            PRESENTATION_SPELLING_FIND_OPTIONS,
          ).find(
            (candidate) =>
              candidate.slideIndex === slideIndex &&
              !skipped.has(candidate.key) &&
              judgedBy(candidate, tag),
          );
          if (match) return { word: issue.word, match, tag };
        }
      }
    }
  }
  return null;
}

/**
 * The occurrences passed over with Ignore Once, carried through a Change of
 * another word: a skipped occurrence later in the same text field moves by
 * the length the change added or took away, so it stays skipped. Its key
 * names its offsets, and a longer or shorter word before it gave it a key
 * no skip matched -- the word came back when the slide was checked again
 * (the Docs counterpart of WUA 028). PURE.
 */
export function carrySkippedSpellingKeys(
  skipped: ReadonlySet<string>,
  changed: Pick<PresentationFindMatch, "key" | "start" | "end">,
  delta: number,
): Set<string> {
  const split = (key: string) => {
    const endAt = key.lastIndexOf(":");
    const startAt = endAt > 0 ? key.lastIndexOf(":", endAt - 1) : -1;
    if (startAt < 0) return null;
    const start = Number(key.slice(startAt + 1, endAt));
    const end = Number(key.slice(endAt + 1));
    return Number.isInteger(start) && Number.isInteger(end)
      ? { place: key.slice(0, startAt), start, end }
      : null;
  };
  const at = split(changed.key);
  if (!at || delta === 0) return new Set(skipped);
  const next = new Set<string>();
  for (const key of skipped) {
    const parsed = split(key);
    if (parsed && parsed.place === at.place && parsed.start >= changed.end) {
      next.add(`${parsed.place}:${parsed.start + delta}:${parsed.end + delta}`);
    } else {
      next.add(key);
    }
  }
  return next;
}

