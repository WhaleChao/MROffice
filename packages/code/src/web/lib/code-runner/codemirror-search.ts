/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers bridging the code editor's find-bar state to @codemirror/search.
 *
 * The find-bar UI/markup (and its data-testids) is shared with the textarea
 * editor for test parity; this module turns its state into a CM6
 * {@link SearchQuery} and computes the match-count / active-index display, which
 * `@codemirror/search` does not expose directly. Kept framework-free and
 * unit-tested (no EditorView needed).
 */
import { SearchQuery } from "@codemirror/search";

import {
  findMatchPreview,
  findMatches,
  MAX_FIND_QUERY_CHARS,
  nextMatchIndex,
  validateRegex,
  type FindOptions,
  type FindMatchPreview,
  type Match,
} from "./find-replace";

export interface FindBarState {
  query: string;
  replacement: string;
  caseSensitive: boolean;
  /** Whole-word matching (\b boundaries). Mutually exclusive with regex. */
  wholeWord: boolean;
  /** Treat the query as a regular expression ($1 capture-group replacement). */
  useRegex: boolean;
  /** Scope find/replace to a pinned selection range (find-in-selection). */
  inSelection: boolean;
  /** The pinned bounds when {@link inSelection} is on. */
  selRange: { start: number; end: number } | null;
  /**
   * Preserve Case (Aa/AB): rewrite each replacement to the case shape of the
   * text it replaces. Optional so existing callers/tests stay valid; absent ==
   * off. Affects replacement only — never the match set or the highlight.
   */
  preserveCase?: boolean;
}

/**
 * Translate the find-bar state into the shared {@link FindOptions}. Whole-word
 * and regex are mutually exclusive (regex wins) — matching the textarea editor —
 * and find-in-selection only scopes when a range is actually pinned. This is the
 * SINGLE source of truth for the find semantics: the count display, navigation,
 * and replace all build their matches through this so they can never diverge.
 */
export function findOptions(state: FindBarState): FindOptions {
  return {
    caseSensitive: state.caseSensitive,
    wholeWord: state.wholeWord && !state.useRegex,
    regex: state.useRegex,
    selectionRange:
      state.inSelection && state.selRange ? state.selRange : undefined,
    preserveCase: !!state.preserveCase,
  };
}

/**
 * Build a CM6 {@link SearchQuery} from the find-bar state. This drives ONLY the
 * visual match highlighting from `@codemirror/search`; navigation and replace
 * are driven from the shared {@link findMatches}/`expandReplacement` so the
 * stepped match and the replacement (incl. $1 capture groups + find-in-selection
 * scoping) exactly track the displayed count. The regexp/wholeWord flags are
 * mirrored so the highlight matches what the bar reports.
 */
export function buildSearchQuery(
  state: FindBarState,
  suppressHighlights = false,
): SearchQuery {
  const unsafeRegex =
    state.useRegex && validateRegex(state.query, state.caseSensitive) !== null;
  const queryTooLong = state.query.length > MAX_FIND_QUERY_CHARS;
  const wholeWord = state.wholeWord && !state.useRegex;
  return new SearchQuery({
    search: unsafeRegex || queryTooLong || suppressHighlights ? "" : state.query,
    caseSensitive: state.caseSensitive,
    replace: state.replacement,
    literal: !state.useRegex || unsafeRegex || queryTooLong || suppressHighlights,
    regexp:
      state.useRegex && !unsafeRegex && !queryTooLong && !suppressHighlights,
    wholeWord,
  });
}

/**
 * All matches of the find-bar query in `text`, honoring case/whole-word/regex/
 * find-in-selection. Delegates to the shared {@link findMatches} so the count
 * shown in the bar exactly tracks what the editor steps through and replaces.
 */
export function searchMatches(text: string, state: FindBarState): Match[] {
  return findMatches(text, state.query, findOptions(state));
}

/** Bounded match window used by the live CodeMirror find bar. */
export function searchMatchPreview(
  text: string,
  state: FindBarState,
): FindMatchPreview {
  return findMatchPreview(text, state.query, findOptions(state));
}

/**
 * The 1-based "current / total" label for the find bar, e.g. `"2/5"` or
 * `"0/0"` when there is no query or no match. `caret` is the editor's current
 * selection head; the active match is the first at/after the caret (wrapping).
 */
export function matchCountLabel(
  matches: Match[],
  caret: number,
  truncated = false,
): { current: number; total: number; label: string } {
  const total = matches.length;
  if (total === 0) return { current: 0, total: 0, label: "0/0" };
  const idx = nextMatchIndex(matches, caret, 1);
  const current = (idx < 0 ? 0 : idx) + 1;
  return {
    current,
    total,
    label: `${current}/${total}${truncated ? "+" : ""}`,
  };
}
