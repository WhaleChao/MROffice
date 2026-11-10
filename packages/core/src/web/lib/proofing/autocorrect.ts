/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * AutoCorrect's replace-text-as-you-type list (Microsoft Support, "AutoCorrect
 * features in Excel"): two parallel lists, the word as typed and what replaces
 * it. Entries come from the Spelling dialog's AutoCorrect button and live with
 * the user's other proofing settings (lib/proofing/proofing-preferences).
 *
 * An entry typed in lower case matches the word in any case and hands its
 * capitalization on ("Teh" -> "The", "TEH" -> "THE"); an entry with capitals
 * matches exactly. Only whole words are replaced, never part of a word or of
 * an Internet / file address.
 */

import {
  AUTOCORRECT_ENTRIES_STORAGE_KEY,
  proofingStorage,
  readAutoCorrectEntries,
  type AutoCorrectEntry,
} from "./proofing-preferences";
import { matchSpellingCase } from "./spelling-checker";
import { tokenizeForSpelling } from "./spelling-tokenize";

let cached: { raw: string | null; entries: AutoCorrectEntry[] } | null = null;

/**
 * The user's AutoCorrect entries as stored now, parsed again only when the
 * stored list changed: an entry commit reads them, and must stay cheap.
 */
export function currentAutoCorrectEntries(): AutoCorrectEntry[] {
  const storage = proofingStorage();
  let raw: string | null = null;
  try {
    raw = storage?.getItem(AUTOCORRECT_ENTRIES_STORAGE_KEY) ?? null;
  } catch {
    raw = null;
  }
  if (!cached || cached.raw !== raw) {
    cached = { raw, entries: raw ? readAutoCorrectEntries(storage) : [] };
  }
  return cached.entries;
}

function replacementFor(
  word: string,
  entries: readonly AutoCorrectEntry[],
): string | null {
  let folded: string | null = null;
  for (const entry of entries) {
    if (entry.from === word) return entry.to;
    if (folded === null && entry.from === entry.from.toLowerCase() && entry.from === word.toLowerCase()) {
      folded = matchSpellingCase(word, entry.to);
    }
  }
  return folded;
}

/** `text` with every whole-word AutoCorrect match replaced. */
export function applyAutoCorrect(
  text: string,
  entries: readonly AutoCorrectEntry[],
): string {
  if (!text || entries.length === 0) return text;
  const tokens = tokenizeForSpelling(text, { ignoreInternetAndFileAddresses: true });
  let result = text;
  for (let index = tokens.length - 1; index >= 0; index -= 1) {
    const token = tokens[index];
    const replacement = replacementFor(token.word, entries);
    if (replacement === null) continue;
    result = result.slice(0, token.start) + replacement + result.slice(token.end);
  }
  return result;
}
