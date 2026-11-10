/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The main spelling dictionaries MRAI ships, which one judges text set in a
 * given proofing language, and their loader.
 *
 * Each is a Hunspell `.aff` / `.dic` pair served as a static file under
 * `/proofing/dictionaries/<tag>/` (public/proofing/dictionaries/README.md
 * names the source, version and licence of each), most of them stored
 * gzip-compressed. Nothing is fetched until a check needs that language. In
 * the browser a dictionary is read and queried by Hunspell (WebAssembly) in
 * the spelling worker (public/proofing/spelling-worker.js): one per page
 * session, shared by every editor, so neither loading nor judging a long
 * document runs on the page thread. The checker itself is synchronous
 * (spelling-checker.ts): a loaded dictionary answers from the verdicts the
 * worker has sent, asks the worker for any word it has not judged yet
 * (reading it as correct until then) and tells its subscribers when those
 * verdicts arrive, so a caller checks again and converges.
 */

import type { SpellingDictionary, SpellingJudgeOptions } from "./spelling-checker";
import {
  spellingWorkerAvailable,
  spellingWorkerRequest,
} from "./spelling-worker-client";
import type { SpellingScript } from "./spelling-tokenize";

export interface SpellingDictionaryInfo {
  /** BCP 47 tag; also the folder name under /proofing/dictionaries. */
  readonly tag: string;
  /** Script of the words the dictionary judges; other scripts are skipped. */
  readonly script: SpellingScript;
  /** The files are stored as index.aff.gz / index.dic.gz. */
  readonly compressed: boolean;
}

function entry(tag: string, script: SpellingScript = "Latin", compressed = true): SpellingDictionaryInfo {
  return { tag, script, compressed };
}

export const SPELLING_DICTIONARIES: readonly SpellingDictionaryInfo[] = [
  entry("en-US", "Latin", false),
  entry("en-GB", "Latin", false),
  entry("en-AU"),
  entry("en-CA"),
  entry("en-ZA"),
  entry("nl", "Latin", false),
  entry("da"),
  entry("de-DE"),
  entry("de-AT"),
  entry("de-CH"),
  entry("es"),
  entry("fr"),
  entry("it"),
  entry("nb"),
  entry("nn"),
  entry("pt-BR"),
  entry("pt-PT"),
  entry("sv"),
  entry("id"),
  entry("vi"),
  entry("ko", "Hangul"),
  entry("el", "Greek"),
  entry("ru", "Cyrillic"),
  entry("uk", "Cyrillic"),
  entry("bg", "Cyrillic"),
  entry("sr-Cyrl", "Cyrillic"),
  entry("sr-Latn"),
  entry("tr"),
  entry("pl"),
  entry("cs"),
  entry("sk"),
  entry("hu"),
  entry("hr"),
  entry("sl"),
  entry("ro"),
  entry("ca"),
  entry("cy"),
  entry("lt"),
  entry("lv"),
];

export const SPELLING_DICTIONARY_BASE_PATH = "/proofing/dictionaries";

const FALLBACK_LANGUAGE = "en-US";

const BY_TAG = new Map(SPELLING_DICTIONARIES.map((info) => [info.tag.toLowerCase(), info]));

export function spellingDictionaryInfo(tag: string | null | undefined): SpellingDictionaryInfo | null {
  if (!tag) return null;
  return BY_TAG.get(tag.toLowerCase()) ?? null;
}

/** English regions that spell the British way (Word's English variants other
 *  than the US, Australian, Canadian and South African ones MRAI ships). */
const BRITISH_SPELLING_REGIONS = new Set([
  "gb", "uk", "ie", "nz", "in", "sg", "my", "hk", "pk", "bd", "lk", "mt", "cy",
  "gi", "im", "je", "gg", "zw", "zm", "mw", "bw", "na", "ng", "gh", "ke", "ug",
  "tz", "jm", "tt", "bz", "bs", "bb",
]);

/** A primary language whose one dictionary serves every region. */
const ONE_DICTIONARY_LANGUAGES = new Set([
  "nl", "da", "es", "fr", "it", "sv", "id", "vi", "ko", "el", "ru", "uk", "bg",
  "tr", "pl", "cs", "sk", "hu", "hr", "sl", "ro", "ca", "cy", "lt", "lv",
]);

/**
 * The shipped dictionary for text in proofing language `language` (a BCP 47
 * tag as Word writes it in w:lang, "de-DE", "pt-br", "en_GB", "sr-Latn-RS"),
 * or null when MRAI ships none for it: the text is then not checked rather
 * than judged by another language's dictionary.
 */
export function spellingDictionaryForLanguage(language: string | null | undefined): string | null {
  if (!language) return null;
  const tag = language.trim().replace(/_/g, "-");
  const exact = spellingDictionaryInfo(tag);
  if (exact) return exact.tag;
  const [primary = "", ...rest] = tag.toLowerCase().split("-");
  const script = rest.find((part) => part.length === 4) ?? null;
  const region = rest.find((part) => part.length === 2 || /^\d{3}$/.test(part)) ?? null;
  switch (primary) {
    case "en":
      if (region === "au") return "en-AU";
      if (region === "ca") return "en-CA";
      if (region === "za") return "en-ZA";
      return region && BRITISH_SPELLING_REGIONS.has(region) ? "en-GB" : "en-US";
    case "de":
      if (region === "at") return "de-AT";
      if (region === "ch" || region === "li") return "de-CH";
      return "de-DE";
    case "pt":
      // Bare "pt" is Brazilian Portuguese (CLDR likely subtags).
      return region && region !== "br" ? "pt-PT" : "pt-BR";
    case "no":
    case "nb":
      return "nb";
    case "nn":
      return "nn";
    case "in": // the withdrawn code for Indonesian, still in old documents
      return "id";
    case "sr":
      return script === "latn" ? "sr-Latn" : "sr-Cyrl";
    default:
      return ONE_DICTIONARY_LANGUAGES.has(primary) ? primary : null;
  }
}

/**
 * The shipped dictionary that judges text set in proofing `language` (null:
 * the check's own language, `defaultTag`), or null when MRAI ships none for
 * it, so the text is not checked rather than judged by another language's
 * dictionary. Shared by the Docs pane and the presentation Spelling dialog.
 */
export function proofingDictionaryTag(
  language: string | null,
  defaultTag: string,
): string | null {
  return spellingDictionaryForLanguage(language ?? defaultTag);
}

/**
 * Why text in a language is not checked: its script is written without
 * spaces between words, so a word list cannot find the words (Chinese,
 * Japanese, Thai and their neighbours: Office's proofing tools segment them
 * with language-specific tools MRAI does not have), or MRAI ships no
 * dictionary for it (Finnish, whose words a Hunspell list cannot cover, and
 * every other language without one).
 */
export type SpellingUncheckedReason = "noWordSpaces" | "noDictionary";

const NO_WORD_SPACE_LANGUAGES = new Set(["zh", "ja", "th", "lo", "km", "my", "bo", "dz", "yue", "wuu"]);

export function spellingUncheckedReason(language: string): SpellingUncheckedReason | null {
  if (spellingDictionaryForLanguage(language)) return null;
  const primary = language.trim().replace(/_/g, "-").toLowerCase().split("-")[0] ?? "";
  return NO_WORD_SPACE_LANGUAGES.has(primary) ? "noWordSpaces" : "noDictionary";
}

/**
 * The dictionary a check starts in: the language the user last chose in the
 * Spelling dialog; else the one for the UI language when MRAI ships it (en →
 * en-US, en-GB → en-GB, de → de-DE, es-419 → es, no → nb, ...); else English
 * (United States), the language the Latin words in a workbook most often are
 * when the UI language has no dictionary (Chinese, Japanese, Thai, Finnish).
 * Office starts in the editing language the same way and keeps the dialog's
 * choice as an application setting (`SpellingOptions.DictLang`).
 */
export function defaultSpellingLanguage(
  uiLocale: string | null | undefined,
  chosen: string | null | undefined,
): string {
  const remembered = spellingDictionaryInfo(chosen);
  if (remembered) return remembered.tag;
  return spellingDictionaryForLanguage(uiLocale) ?? FALLBACK_LANGUAGE;
}

/**
 * The language a document's text is checked in where neither the run, its
 * paragraph style nor the document's defaults state one: the Spelling
 * dialog's choice, else the UI language. Unlike `defaultSpellingLanguage`,
 * a UI language with words but no dictionary (Finnish) stays itself, so a
 * Finnish user's own writing is reported as not checked instead of being
 * judged as English; a UI language written without word spaces (Chinese,
 * Japanese, Thai) reads its Latin words as US English, as Word's editing
 * language does in those Office versions; Korean's Hangul words are routed
 * by script (documentEditorProofing) and its Latin words read as English.
 */
export function defaultDocumentSpellingLanguage(
  uiLocale: string | null | undefined,
  chosen: string | null | undefined,
): string {
  const remembered = spellingDictionaryInfo(chosen);
  if (remembered?.script === "Latin") return remembered.tag;
  const locale = (uiLocale ?? "").trim();
  if (!locale) return FALLBACK_LANGUAGE;
  const shipped = spellingDictionaryForLanguage(locale);
  if (shipped) {
    return spellingDictionaryInfo(shipped)?.script === "Latin" ? shipped : FALLBACK_LANGUAGE;
  }
  return spellingUncheckedReason(locale) === "noDictionary" ? locale : FALLBACK_LANGUAGE;
}

export interface LoadedSpellingDictionary extends SpellingDictionary {
  readonly tag: string;
  readonly script: SpellingScript;
  /** Words or suggestions asked for are still being judged off the page. */
  pending(): boolean;
  /** Resolves once everything asked for so far has been answered. */
  settled(): Promise<void>;
  /** `listener` runs each time answers asked for by correct()/suggest()
   *  arrive (check again then); returns the unsubscribe. */
  subscribe(listener: () => void): () => void;
  /** Judge these words now, ahead of the check that will ask for them. */
  prepare(words: Iterable<string>, options?: SpellingJudgeOptions): Promise<void>;
}

type FetchText = (url: string) => Promise<string>;

async function fetchDictionaryText(url: string): Promise<string> {
  const response = await fetch(url, { credentials: "same-origin" });
  if (!response.ok) throw new Error(`spelling dictionary ${url}: HTTP ${response.status}`);
  return response.text();
}

/** Words a check asks the worker to judge in one request. */
const CHECK_BATCH = 4000;
/** Verdicts kept per dictionary before the oldest are forgotten. */
const MAX_VERDICTS = 250_000;
const MAX_SUGGESTION_ENTRIES = 2_000;

function remember<V>(map: Map<string, V>, key: string, value: V, limit: number): void {
  if (map.size >= limit) {
    const drop = Math.ceil(limit / 4);
    let index = 0;
    for (const old of map.keys()) {
      if (index++ >= drop) break;
      map.delete(old);
    }
  }
  map.set(key, value);
}

/** The words the checker may ask of `word`: itself, with a straight
 *  apostrophe, and each part of a hyphenated word. */
function judgedForms(word: string): string[] {
  const forms = [word];
  if (word.includes("’")) forms.push(word.replace(/’/g, "'"));
  if (word.includes("-")) forms.push(...word.split("-").filter(Boolean));
  return forms;
}

/** A dictionary read by the spelling worker. */
class WorkerSpellingDictionary implements LoadedSpellingDictionary {
  readonly tag: string;
  readonly script: SpellingScript;
  private readonly french: boolean;
  private readonly strict = new Map<string, boolean>();
  private readonly lenient = new Map<string, boolean>();
  private readonly suggestions = new Map<string, string[]>();
  private readonly queuedStrict = new Set<string>();
  private readonly queuedLenient = new Set<string>();
  private readonly queuedSuggestions = new Set<string>();
  private readonly inflight = new Set<Promise<void>>();
  private readonly listeners = new Set<() => void>();
  private flushScheduled = false;

  constructor(info: SpellingDictionaryInfo) {
    this.tag = info.tag;
    this.script = info.script;
    this.french = info.tag.toLowerCase().startsWith("fr");
  }

  private lenientFor(options: SpellingJudgeOptions | undefined): boolean {
    return this.french && options?.enforceAccentedUppercase === false;
  }

  correct(word: string, options?: SpellingJudgeOptions): boolean {
    const lenient = this.lenientFor(options);
    const known = (lenient ? this.lenient : this.strict).get(word);
    if (known !== undefined) return known;
    (lenient ? this.queuedLenient : this.queuedStrict).add(word);
    this.scheduleFlush();
    return true;
  }

  suggest(word: string): string[] {
    const known = this.suggestions.get(word);
    if (known) return known;
    this.queuedSuggestions.add(word);
    this.scheduleFlush();
    return [];
  }

  pending(): boolean {
    return (
      this.inflight.size > 0 ||
      this.queuedStrict.size > 0 ||
      this.queuedLenient.size > 0 ||
      this.queuedSuggestions.size > 0
    );
  }

  async settled(): Promise<void> {
    while (this.pending()) {
      this.flush();
      await Promise.all([...this.inflight]);
    }
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async prepare(words: Iterable<string>, options?: SpellingJudgeOptions): Promise<void> {
    const lenient = this.lenientFor(options);
    const known = lenient ? this.lenient : this.strict;
    const queue = lenient ? this.queuedLenient : this.queuedStrict;
    for (const word of words) {
      for (const form of judgedForms(word)) if (!known.has(form)) queue.add(form);
    }
    const sent = this.flush(false);
    await Promise.all(sent);
  }

  private scheduleFlush(): void {
    if (this.flushScheduled) return;
    this.flushScheduled = true;
    queueMicrotask(() => {
      this.flushScheduled = false;
      this.flush();
    });
  }

  private track(work: Promise<void>, notify: boolean): Promise<void> {
    const tracked = work.finally(() => {
      this.inflight.delete(tracked);
      if (notify) for (const listener of [...this.listeners]) listener();
    });
    this.inflight.add(tracked);
    return tracked;
  }

  private flush(notify = true): Promise<void>[] {
    const sent: Promise<void>[] = [];
    for (const lenient of [false, true]) {
      const queue = lenient ? this.queuedLenient : this.queuedStrict;
      if (queue.size === 0) continue;
      const words = [...queue];
      queue.clear();
      const into = lenient ? this.lenient : this.strict;
      for (let start = 0; start < words.length; start += CHECK_BATCH) {
        const batch = words.slice(start, start + CHECK_BATCH);
        sent.push(
          this.track(
            spellingWorkerRequest<boolean[]>({
              op: "check",
              tag: this.tag,
              words: batch,
              lenientUppercaseAccents: lenient,
            }).then(
              (verdicts) => {
                batch.forEach((word, index) =>
                  remember(into, word, verdicts[index] !== false, MAX_VERDICTS),
                );
              },
              // A worker that failed mid-check flags nothing it did not judge.
              () => batch.forEach((word) => remember(into, word, true, MAX_VERDICTS)),
            ),
            notify,
          ),
        );
      }
    }
    if (this.queuedSuggestions.size > 0) {
      const words = [...this.queuedSuggestions];
      this.queuedSuggestions.clear();
      for (const word of words) {
        sent.push(
          this.track(
            spellingWorkerRequest<string[]>({ op: "suggest", tag: this.tag, word }).then(
              (list) => remember(this.suggestions, word, list, MAX_SUGGESTION_ENTRIES),
              () => remember(this.suggestions, word, [], MAX_SUGGESTION_ENTRIES),
            ),
            notify,
          ),
        );
      }
    }
    return sent;
  }
}

const NO_PENDING = Promise.resolve();

/** A dictionary read by nspell on the calling thread (tests, and pages that
 *  cannot start a worker). Every answer is immediate. */
function inProcessDictionary(
  info: SpellingDictionaryInfo,
  speller: { correct(word: string): boolean; suggest(word: string): string[] },
): LoadedSpellingDictionary {
  return {
    tag: info.tag,
    script: info.script,
    correct: (word: string) => speller.correct(word),
    suggest: (word: string) => speller.suggest(word),
    pending: () => false,
    settled: () => NO_PENDING,
    subscribe: () => () => undefined,
    prepare: () => NO_PENDING,
  };
}

const cache = new Map<string, Promise<LoadedSpellingDictionary>>();

function absoluteUrl(path: string): string {
  if (typeof window === "undefined" || !window.location) return path;
  return new URL(path, window.location.href).toString();
}

function loadInProcess(
  info: SpellingDictionaryInfo,
  base: string,
  fetchText: FetchText,
): Promise<LoadedSpellingDictionary> {
  if (info.compressed) {
    return Promise.reject(new Error(`spelling dictionary ${info.tag} needs the spelling worker`));
  }
  return Promise.all([
    import("nspell"),
    fetchText(`${base}/index.aff`),
    fetchText(`${base}/index.dic`),
  ]).then(([module, aff, dic]) => {
    // A missing file served as an HTML page would parse as an empty
    // dictionary that flags every word; refuse it instead.
    if (!/^\s*SET\s/m.test(aff) || !/^\d+\s*$/m.test(dic.slice(0, 64))) {
      throw new Error(`spelling dictionary ${info.tag} is not a Hunspell dictionary`);
    }
    const create = (module as { default?: typeof import("nspell") }).default ??
      (module as unknown as typeof import("nspell"));
    return inProcessDictionary(info, create(aff, dic));
  });
}

/**
 * The dictionary for `tag`, loaded on first use and shared after. A failed
 * load is forgotten, so the next attempt loads again. `fetchText` (tests)
 * reads an uncompressed dictionary with nspell on the calling thread.
 */
export function loadSpellingDictionary(
  tag: string,
  fetchText?: FetchText,
): Promise<LoadedSpellingDictionary> {
  const info = spellingDictionaryInfo(tag);
  if (!info) return Promise.reject(new Error(`no spelling dictionary for ${tag}`));
  const cached = cache.get(info.tag);
  if (cached) return cached;
  const base = `${SPELLING_DICTIONARY_BASE_PATH}/${info.tag}`;
  const suffix = info.compressed ? ".gz" : "";
  const loading =
    fetchText || !spellingWorkerAvailable()
      ? loadInProcess(info, base, fetchText ?? fetchDictionaryText)
      : spellingWorkerRequest<true>({
          op: "load",
          tag: info.tag,
          aff: absoluteUrl(`${base}/index.aff${suffix}`),
          dic: absoluteUrl(`${base}/index.dic${suffix}`),
        }).then(() => new WorkerSpellingDictionary(info) as LoadedSpellingDictionary);
  cache.set(info.tag, loading);
  loading.catch(() => {
    if (cache.get(info.tag) === loading) cache.delete(info.tag);
  });
  return loading;
}

/** Whether any of these dictionaries is still judging words off the page. */
export function spellingDictionariesPending(
  dictionaries: Iterable<SpellingDictionary>,
): boolean {
  for (const dictionary of dictionaries) {
    if ((dictionary as Partial<LoadedSpellingDictionary>).pending?.()) return true;
  }
  return false;
}

/** Resolves once none of these dictionaries is judging words any more. */
export async function spellingDictionariesSettled(
  dictionaries: Iterable<SpellingDictionary>,
): Promise<void> {
  await Promise.all(
    [...dictionaries].map((dictionary) =>
      (dictionary as Partial<LoadedSpellingDictionary>).settled?.(),
    ),
  );
}

/**
 * Run a synchronous check until every word it asked of these dictionaries
 * had been judged: a check that met a word the worker has not answered yet
 * waits for the answer and runs again (the answers are kept, so it
 * converges; `rounds` bounds it all the same).
 */
export async function runSpellingCheckJudged<T>(
  dictionaries: () => Iterable<SpellingDictionary>,
  check: () => T,
  rounds = 6,
): Promise<T> {
  let result = check();
  for (let round = 1; round < rounds && spellingDictionariesPending(dictionaries()); round += 1) {
    await spellingDictionariesSettled(dictionaries());
    result = check();
  }
  return result;
}

/** Subscribe to every dictionary in the list; returns the unsubscribe. */
export function subscribeSpellingDictionaries(
  dictionaries: Iterable<SpellingDictionary>,
  listener: () => void,
): () => void {
  const stops = [...dictionaries].map(
    (dictionary) =>
      (dictionary as Partial<LoadedSpellingDictionary>).subscribe?.(listener) ?? (() => undefined),
  );
  return () => stops.forEach((stop) => stop());
}

/** Tests only: forget every loaded dictionary. */
export function resetSpellingDictionaryCacheForTests(): void {
  cache.clear();
}
