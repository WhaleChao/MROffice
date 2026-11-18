/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's automatic hyphenation for the Docs line engine: the points a
 * paragraph's words may be broken at, read from the page (each run's
 * language and "Do not check spelling or grammar"), for docsEngineLineStarts
 * to choose from by Word's rules (hyphenCut).
 *
 * A paragraph with a word in a language that has no patterns here is left
 * as it was: the browser hyphenates it (`hyphens: auto`).
 */
import type { DocsParagraph } from "./docs-lines";
import { hyphenationLanguage, hyphenationPoints } from "./hyphenation";

/** Layout > Hyphenation > Hyphenation Options, as the document states them. */
export interface DocsHyphenationSettings {
  /** `w:hyphenationZone`, twips; Word's 0.25" when unstated. */
  readonly zoneTwips?: number;
  /** `w:consecutiveHyphenLimit` (0 or unstated: no limit). */
  readonly consecutiveLimit?: number;
  /** false: `w:doNotHyphenateCaps`. */
  readonly hyphenateCaps?: boolean;
}

const WORD_DEFAULT_ZONE_TWIPS = 360;
const TWIPS_PER_CSS_PX = 15;
const LETTER = /^\p{L}$/u;
const SOFT_HYPHEN = "­";
/** Text Word does not proof, it does not hyphenate either. */
const NO_PROOF = '[data-proofing-no-proof="true"]';

/** A run's hyphenation patterns: a language, null (not hyphenated), or
 * undefined (a language without patterns here). */
type RunLanguage = string | null | undefined;

function runLanguage(element: HTMLElement, node: Text): RunLanguage {
  const parent = node.parentElement;
  if (!parent) return null;
  const noProof = parent.closest(NO_PROOF);
  if (noProof && element.contains(noProof)) return null;
  const holder = parent.closest("[lang]");
  const tag = holder && (holder === element || element.contains(holder)) ? holder.getAttribute("lang") : null;
  return hyphenationLanguage(tag) ?? undefined;
}

/**
 * `paragraph` with the automatic hyphenation points of its words
 * (`DocsParagraph.hyphenation`), or `paragraph` itself when the engine does
 * not hyphenate it (a word in a language it has no patterns for, or
 * nothing the engine can read).
 */
export function withDocsAutoHyphenation(
  paragraph: DocsParagraph, element: HTMLElement, settings: DocsHyphenationSettings,
): DocsParagraph {
  if (paragraph.unsupported || paragraph.pending || !paragraph.pieces.length) return paragraph;
  const pieces = paragraph.pieces;
  const languages = new Map<Text, RunLanguage>();
  const languageOf = (node: Text): RunLanguage => {
    if (!languages.has(node)) languages.set(node, runLanguage(element, node));
    return languages.get(node);
  };
  const after = new Set<number>();
  let index = 0;
  while (index < pieces.length) {
    const first = pieces[index]!;
    if (first.kind !== "text" || !LETTER.test(first.text)) {
      index += 1;
      continue;
    }
    // A word: letters in a row (an optional hyphen within it included).
    const members: number[] = [];
    let soft = false;
    let end = index;
    while (end < pieces.length && pieces[end]!.kind === "text" && (LETTER.test(pieces[end]!.text) || pieces[end]!.text === SOFT_HYPHEN)) {
      if (pieces[end]!.text === SOFT_HYPHEN) soft = true;
      else members.push(end);
      end += 1;
    }
    index = end;
    const word = members.map((at) => pieces[at]!.text).join("");
    // Too short to break (two letters before a point, three after).
    if (word.length < 5 || members.some((at) => pieces[at]!.text.length !== 1)) continue;
    // A word with an optional hyphen of its own breaks only there.
    if (soft) continue;
    if (settings.hyphenateCaps === false && word === word.toUpperCase() && word !== word.toLowerCase()) continue;
    let language: RunLanguage = null;
    for (const at of members) {
      const node = pieces[at]!.node;
      const own = node ? languageOf(node) : null;
      // A word in a language without patterns: the browser's paragraph.
      if (own === undefined) return paragraph;
      if (own === null || (language !== null && own !== language)) {
        language = null;
        break;
      }
      language = own;
    }
    if (!language) continue;
    for (const point of hyphenationPoints(word, language)) after.add(members[point - 1]!);
  }
  return {
    ...paragraph,
    hyphenation: {
      after,
      zonePx: (settings.zoneTwips ?? WORD_DEFAULT_ZONE_TWIPS) / TWIPS_PER_CSS_PX,
      consecutiveLimit: settings.consecutiveLimit ?? 0,
    },
  };
}
