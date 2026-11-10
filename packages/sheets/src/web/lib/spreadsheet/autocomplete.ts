/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * In-column type-ahead autocomplete (Excel's AutoComplete) — pure matching so
 * the SpreadsheetWindow cell editor can offer an inline ghost completion that is
 * unit-tested in isolation.
 *
 * Excel behaviour modeled:
 *  - As you type into a cell, Excel looks at the OTHER non-empty cells in the
 *    same column and, if exactly one of them begins with what you have typed
 *    (case-insensitively), pre-fills the rest of that value with the suffix
 *    selected (so the next keystroke replaces it, Enter/Arrow accepts it).
 *    Excel keeps the typed prefix verbatim and appends the candidate's remaining
 *    characters: typing "ja" against "January" yields "January" with "ja" kept
 *    and "nuary" highlighted. We return the full value plus `prefixLen` (the
 *    typed length) so the caller highlights only the appended suffix.
 *  - No completion is offered when: the typed text is empty; the text contains
 *    nothing to extend (the candidate equals the typed text); the typed text
 *    looks like a number or formula (Excel never autocompletes numeric entries
 *    or formulas); or more than one distinct candidate matches and they don't
 *    share a longer common prefix to offer.
 *
 * The function is side-effect free and never imports the sheet model — the
 * caller passes the column's existing values.
 */

export interface AutocompleteSuggestion {
  /** The full value to place in the cell (typed prefix + completion). */
  value: string;
  /** Length of the user's typed prefix (the part NOT highlighted). */
  prefixLen: number;
}

/**
 * Whether a typed string should suppress autocomplete because it is a formula or
 * a numeric entry (Excel autocompletes text labels only).
 */
function isNonTextEntry(typed: string): boolean {
  if (typed.startsWith("=")) return true; // formula
  // A leading digit / sign / currency that parses as a number → numeric entry.
  // We only suppress when the WHOLE trimmed string is a finite number, so a
  // label like "3M Company" still autocompletes.
  const trimmed = typed.trim();
  if (trimmed === "") return false;
  return Number.isFinite(Number(trimmed));
}

/**
 * Compute the inline autocomplete suggestion for `typed` against the column's
 * `existingValues`, or null when nothing should be offered.
 *
 * Matching is case-insensitive on the prefix; the returned `value` is the
 * candidate with its ORIGINAL casing for the part beyond the typed prefix, but
 * keeps the user's typed characters verbatim as the prefix (so the cell shows
 * exactly what was typed followed by the candidate's tail).
 */
export function columnAutocomplete(
  typed: string,
  existingValues: readonly string[],
): AutocompleteSuggestion | null {
  if (typeof typed !== "string" || typed === "") return null;
  if (isNonTextEntry(typed)) return null;

  const typedLower = typed.toLowerCase();

  // Distinct candidates (preserve first-seen casing) that START WITH the typed
  // prefix case-insensitively and are STRICTLY LONGER (so there is something to
  // complete). An exact match offers no completion.
  const seen = new Set<string>();
  const candidates: string[] = [];
  for (const raw of existingValues) {
    if (typeof raw !== "string") continue;
    const value = raw;
    if (value.length <= typed.length) continue;
    if (!value.toLowerCase().startsWith(typedLower)) continue;
    if (seen.has(value)) continue;
    seen.add(value);
    candidates.push(value);
  }

  if (candidates.length === 0) return null;

  if (candidates.length === 1) {
    return buildSuggestion(typed, candidates[0]);
  }

  // Multiple candidates: Excel does NOT inline-complete when the next character
  // is ambiguous. Only offer a completion if every candidate shares a common
  // prefix LONGER than what was typed (then complete up to that shared prefix).
  const shared = longestCommonPrefix(candidates);
  if (shared.length > typed.length) {
    return buildSuggestion(typed, shared);
  }
  return null;
}

/**
 * Build the suggestion keeping the user's typed characters as the literal
 * prefix and appending the candidate's tail. `prefixLen` is the typed length so
 * the caller can highlight only the appended suffix.
 */
function buildSuggestion(
  typed: string,
  candidate: string,
): AutocompleteSuggestion {
  // Keep typed verbatim, append the candidate's tail (its original casing).
  const value = typed + candidate.slice(typed.length);
  return { value, prefixLen: typed.length };
}

/** Longest common prefix across a non-empty list of strings. */
function longestCommonPrefix(values: readonly string[]): string {
  if (values.length === 0) return "";
  let prefix = values[0];
  for (let i = 1; i < values.length; i += 1) {
    const value = values[i];
    let j = 0;
    const max = Math.min(prefix.length, value.length);
    while (j < max && prefix[j] === value[j]) j += 1;
    prefix = prefix.slice(0, j);
    if (prefix === "") break;
  }
  return prefix;
}
