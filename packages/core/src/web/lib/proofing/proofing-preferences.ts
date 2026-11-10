/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The user's proofing settings, shared by every editor that checks spelling.
 *
 * In Office these are application settings, never document content (Microsoft
 * Learn, `SpellingOptions`: IgnoreCaps, IgnoreMixedDigits, IgnoreFileNames,
 * SuggestMainOnly, DictLang, UserDict; Microsoft Support, "Add or edit words in
 * a spell check dictionary"): the Proofing options, the dictionary language,
 * the default custom dictionary (Add to Dictionary) and the AutoCorrect
 * replace-as-you-type list follow the user into every workbook. MRAI has no
 * server-side per-user preference store, so they live in this browser's
 * storage, as the Personal Macro Workbook does (lib/spreadsheet/macros.ts).
 * Every reader is total: a missing, blocked or corrupt store reads as the
 * defaults, and a failed write reports false instead of throwing.
 */

export interface SpellingOptions {
  /** "Ignore words in UPPERCASE" — ABC is not flagged. Office default on. */
  ignoreUppercase: boolean;
  /** "Ignore words that contain numbers" — a1b2c3 is not flagged. On. */
  ignoreWordsWithNumbers: boolean;
  /** "Ignore Internet and file addresses". On. */
  ignoreInternetAndFileAddresses: boolean;
  /** "Flag repeated words" — the second "beep" of "beep beep". On. */
  flagRepeatedWords: boolean;
  /** "Enforce accented uppercase in French". Off; French dictionaries only. */
  enforceAccentedUppercaseFrench: boolean;
  /** "Suggest from main dictionary only": custom-dictionary words are still
   *  accepted but never offered as suggestions. Off. */
  suggestFromMainDictionaryOnly: boolean;
  /** The main dictionary chosen in the Spelling dialog; null = the default
   *  for the user's language. */
  dictionaryLanguage: string | null;
}

export const DEFAULT_SPELLING_OPTIONS: Readonly<SpellingOptions> = Object.freeze({
  ignoreUppercase: true,
  ignoreWordsWithNumbers: true,
  ignoreInternetAndFileAddresses: true,
  flagRepeatedWords: true,
  enforceAccentedUppercaseFrench: false,
  suggestFromMainDictionaryOnly: false,
  dictionaryLanguage: null,
});

export const SPELLING_OPTIONS_STORAGE_KEY = "mrai.proofing.spellingOptions.v1";
export const CUSTOM_DICTIONARY_STORAGE_KEY = "mrai.proofing.customDictionary.v1";
export const AUTOCORRECT_ENTRIES_STORAGE_KEY = "mrai.proofing.autoCorrect.v1";

/** Bounds that keep one corrupt or hostile store from costing the page. */
export const MAX_CUSTOM_DICTIONARY_WORDS = 10_000;
export const MAX_CUSTOM_DICTIONARY_WORD_LENGTH = 64;
export const MAX_AUTOCORRECT_ENTRIES = 2_000;
export const MAX_AUTOCORRECT_TEXT_LENGTH = 255;

type ReadableStorage = Pick<Storage, "getItem"> | null | undefined;
type WritableStorage = Pick<Storage, "setItem" | "removeItem"> | null | undefined;

/** The page's localStorage, or null where the browser refuses it. */
export function proofingStorage(): Storage | null {
  try {
    return typeof window === "undefined" ? null : window.localStorage;
  } catch {
    return null;
  }
}

function readJson(storage: ReadableStorage, key: string): unknown {
  try {
    const raw = storage?.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeJson(storage: WritableStorage, key: string, value: unknown): boolean {
  if (!storage) return false;
  try {
    if (value === null) storage.removeItem(key);
    else storage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

const LANGUAGE_TAG = /^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,3}$/;

export function normalizeSpellingOptions(raw: unknown): SpellingOptions {
  const record = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
  const flag = (key: keyof SpellingOptions): boolean =>
    typeof record[key] === "boolean"
      ? (record[key] as boolean)
      : (DEFAULT_SPELLING_OPTIONS[key] as boolean);
  const language = record.dictionaryLanguage;
  return {
    ignoreUppercase: flag("ignoreUppercase"),
    ignoreWordsWithNumbers: flag("ignoreWordsWithNumbers"),
    ignoreInternetAndFileAddresses: flag("ignoreInternetAndFileAddresses"),
    flagRepeatedWords: flag("flagRepeatedWords"),
    enforceAccentedUppercaseFrench: flag("enforceAccentedUppercaseFrench"),
    suggestFromMainDictionaryOnly: flag("suggestFromMainDictionaryOnly"),
    dictionaryLanguage:
      typeof language === "string" && LANGUAGE_TAG.test(language) ? language : null,
  };
}

export function readSpellingOptions(storage: ReadableStorage): SpellingOptions {
  return normalizeSpellingOptions(readJson(storage, SPELLING_OPTIONS_STORAGE_KEY));
}

export function writeSpellingOptions(
  storage: WritableStorage,
  options: SpellingOptions,
): boolean {
  return writeJson(storage, SPELLING_OPTIONS_STORAGE_KEY, normalizeSpellingOptions(options));
}

/**
 * A word the custom dictionary can hold: one token of letters, marks, digits,
 * apostrophes, hyphens or periods, no spaces. Office's custom dictionary is a
 * word list, so a phrase cannot be "added to the dictionary".
 */
export function normalizeCustomDictionaryWord(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const word = raw.normalize("NFC").trim();
  if (!word || word.length > MAX_CUSTOM_DICTIONARY_WORD_LENGTH) return null;
  return /^[\p{L}\p{M}\p{N}'’.-]+$/u.test(word) ? word : null;
}

export function readCustomDictionary(storage: ReadableStorage): string[] {
  const raw = readJson(storage, CUSTOM_DICTIONARY_STORAGE_KEY);
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const words: string[] = [];
  for (const entry of raw) {
    const word = normalizeCustomDictionaryWord(entry);
    if (!word || seen.has(word)) continue;
    seen.add(word);
    words.push(word);
    if (words.length >= MAX_CUSTOM_DICTIONARY_WORDS) break;
  }
  return words;
}

export function writeCustomDictionary(
  storage: WritableStorage,
  words: readonly string[],
): boolean {
  const normalized = [
    ...new Set(
      words
        .map((word) => normalizeCustomDictionaryWord(word))
        .filter((word): word is string => word !== null),
    ),
  ].slice(0, MAX_CUSTOM_DICTIONARY_WORDS);
  return writeJson(
    storage,
    CUSTOM_DICTIONARY_STORAGE_KEY,
    normalized.length > 0 ? normalized : null,
  );
}

/** One AutoCorrect entry: what is typed (`from`) and what replaces it (`to`). */
export interface AutoCorrectEntry {
  from: string;
  to: string;
}

export function normalizeAutoCorrectEntry(raw: unknown): AutoCorrectEntry | null {
  if (!raw || typeof raw !== "object") return null;
  const record = raw as Record<string, unknown>;
  if (typeof record.from !== "string" || typeof record.to !== "string") return null;
  const from = record.from.normalize("NFC").trim();
  const to = record.to.normalize("NFC");
  // The typed side is one word (AutoCorrect fires when a word is finished);
  // the replacement may be a phrase.
  if (
    !from ||
    from.length > MAX_AUTOCORRECT_TEXT_LENGTH ||
    /\s/.test(from) ||
    !to.trim() ||
    to.length > MAX_AUTOCORRECT_TEXT_LENGTH ||
    from === to
  ) {
    return null;
  }
  return { from, to };
}

export function readAutoCorrectEntries(storage: ReadableStorage): AutoCorrectEntry[] {
  const raw = readJson(storage, AUTOCORRECT_ENTRIES_STORAGE_KEY);
  if (!Array.isArray(raw)) return [];
  const byFrom = new Map<string, AutoCorrectEntry>();
  for (const candidate of raw) {
    const entry = normalizeAutoCorrectEntry(candidate);
    if (!entry) continue;
    // One replacement per typed word; the later entry wins, as replacing an
    // existing AutoCorrect entry does.
    byFrom.delete(entry.from.toLowerCase());
    byFrom.set(entry.from.toLowerCase(), entry);
  }
  return [...byFrom.values()].slice(-MAX_AUTOCORRECT_ENTRIES);
}

export function writeAutoCorrectEntries(
  storage: WritableStorage,
  entries: readonly AutoCorrectEntry[],
): boolean {
  const byFrom = new Map<string, AutoCorrectEntry>();
  for (const candidate of entries) {
    const entry = normalizeAutoCorrectEntry(candidate);
    if (!entry) continue;
    byFrom.delete(entry.from.toLowerCase());
    byFrom.set(entry.from.toLowerCase(), entry);
  }
  const normalized = [...byFrom.values()].slice(-MAX_AUTOCORRECT_ENTRIES);
  return writeJson(
    storage,
    AUTOCORRECT_ENTRIES_STORAGE_KEY,
    normalized.length > 0 ? normalized : null,
  );
}
