/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Words of a text, as a spelling checker reads them.
 *
 * A word is a run of letters, marks and digits that may carry apostrophes
 * ("don't", "O’Brien") and hyphens ("well-known") inside it. Scripts written
 * without spaces between words — Chinese, Japanese, Thai and their
 * neighbours — separate words instead of joining them: a dictionary of words
 * cannot check them (Office's proofing tools do not flag them either), and a
 * Latin word inside such text ("Excel表格") is still read on its own.
 *
 * Offsets are UTF-16 indexes into the text, so a caller can replace a word
 * with `text.slice(0, start) + replacement + text.slice(end)`.
 */

export interface SpellingToken {
  /** First UTF-16 index of the word. */
  start: number;
  /** One past its last UTF-16 index. */
  end: number;
  word: string;
}

export interface SpellingTokenizeOptions {
  /** "Ignore Internet and file addresses": skip URLs, e-mail addresses,
   *  file paths and file names whole. */
  ignoreInternetAndFileAddresses: boolean;
}

const NO_SPACE_SCRIPTS =
  "\\p{Script=Han}\\p{Script=Hiragana}\\p{Script=Katakana}\\p{Script=Thai}" +
  "\\p{Script=Lao}\\p{Script=Khmer}\\p{Script=Myanmar}\\p{Script=Tibetan}" +
  "\\p{Script=Bopomofo}\\u3000-\\u303f\\uff00-\\uffef";
const LETTER = `(?:(?![${NO_SPACE_SCRIPTS}])[\\p{L}\\p{M}\\p{N}])`;
const WORD_SOURCE = `${LETTER}+(?:['’]${LETTER}+)*(?:-${LETTER}+(?:['’]${LETTER}+)*)*`;
const WORD = new RegExp(WORD_SOURCE, "gu");
const CHUNK = /\S+/gu;

const URL_SCHEME = /^[a-z][a-z0-9+.-]*:\/\//i;
const URI_PREFIX = /^(?:mailto|file|news|tel|urn):/i;
const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/u;
const WINDOWS_PATH = /^[a-z]:[\\/]/i;
const UNC_PATH = /^\\\\[^\\\s]/;
const POSIX_PATH = /^(?:~|\.{1,2})?\/[^/\s]+\/?/;
const DOMAIN =
  /^(?:[a-z0-9-]+\.)+(?:com|org|net|edu|gov|mil|int|info|biz|io|ai|app|dev|co|me|tv|cn|uk|us|de|fr|nl|jp|kr|es|it|se|no|dk|fi|br|in|ru|au|ca|ch|eu|tw|hk|sg|vn|th|id)(?:[/:?#]\S*)?$/i;
const FILE_NAME =
  /^[\p{L}\p{N}_.\-~()]+\.(?:xlsx|xlsm|xlsb|xls|csv|tsv|docx|doc|pptx|ppt|pdf|txt|rtf|md|png|jpe?g|gif|svg|webp|zip|json|xml|html?|js|ts|tsx|py|java|sql|log)$/iu;

/** Punctuation that wraps an address in running text ("(see www.x.com).") */
function trimAddressPunctuation(chunk: string): string {
  return chunk
    .replace(/^[("'“‘<[{]+/u, "")
    .replace(/[)"'”’>\]}.,;:!?]+$/u, "");
}

/**
 * Whether a whitespace-delimited chunk is an Internet or file address, the
 * class Office's "Ignore Internet and file addresses" passes over (Microsoft
 * Support lists URLs, file paths and mailto addresses as its examples).
 */
export function isInternetOrFileAddress(chunk: string): boolean {
  const core = trimAddressPunctuation(chunk);
  if (core.length < 3) return false;
  return (
    URL_SCHEME.test(core) ||
    URI_PREFIX.test(core) ||
    /^www\./i.test(core) ||
    EMAIL.test(core) ||
    WINDOWS_PATH.test(core) ||
    UNC_PATH.test(core) ||
    (POSIX_PATH.test(core) && core.indexOf("/", 1) > 0) ||
    DOMAIN.test(core) ||
    FILE_NAME.test(core)
  );
}

/**
 * The words of `text` in order. With addresses ignored, a chunk that is an
 * address contributes no word at all.
 */
export function tokenizeForSpelling(
  text: string,
  options: SpellingTokenizeOptions,
): SpellingToken[] {
  const tokens: SpellingToken[] = [];
  CHUNK.lastIndex = 0;
  for (let chunk = CHUNK.exec(text); chunk; chunk = CHUNK.exec(text)) {
    if (options.ignoreInternetAndFileAddresses && isInternetOrFileAddress(chunk[0])) {
      continue;
    }
    WORD.lastIndex = 0;
    for (let match = WORD.exec(chunk[0]); match; match = WORD.exec(chunk[0])) {
      const start = chunk.index + match.index;
      tokens.push({ start, end: start + match[0].length, word: match[0] });
    }
  }
  return tokens;
}

/** True when every letter of the word belongs to the Latin script. */
export function isLatinScriptWord(word: string): boolean {
  return /\p{L}/u.test(word) && /^[\p{Script=Latin}\p{M}\p{N}'’-]+$/u.test(word);
}

/** The scripts MRAI's spelling dictionaries are written in. */
export type SpellingScript = "Latin" | "Greek" | "Cyrillic" | "Hangul";

const SCRIPT_WORD: Record<SpellingScript, RegExp> = {
  Latin: /^[\p{Script=Latin}\p{M}\p{N}'’-]+$/u,
  Greek: /^[\p{Script=Greek}\p{M}\p{N}'’-]+$/u,
  Cyrillic: /^[\p{Script=Cyrillic}\p{M}\p{N}'’-]+$/u,
  Hangul: /^[\p{Script=Hangul}\p{M}\p{N}'’-]+$/u,
};

/** True when every letter of the word belongs to `script`: the words a
 *  dictionary in that script can judge (others are passed over). */
export function isScriptWord(word: string, script: SpellingScript): boolean {
  return /\p{L}/u.test(word) && SCRIPT_WORD[script].test(word);
}
