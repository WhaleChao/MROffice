"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useMemo, useRef, useState, type ReactElement } from "react";

import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  normalizeCustomDictionaryWord,
  proofingStorage,
  readCustomDictionary,
  readSpellingOptions,
  writeCustomDictionary,
} from "@mrnote/core/web/lib/proofing/proofing-preferences.ts";
import {
  ignoreKey,
  ignoreWordForSession,
  spellingSessionIgnoredWords,
  spellingSuggestions,
  type SpellingCheckContext,
  type SpellingDictionary,
} from "@mrnote/core/web/lib/proofing/spelling-checker.ts";
import {
  defaultSpellingLanguage,
  loadSpellingDictionary,
  runSpellingCheckJudged,
} from "@mrnote/core/web/lib/proofing/spelling-dictionaries.ts";
import type { PresentationFindMatch } from "../../../../lib/presentation-find-replace";
import {
  carrySkippedSpellingKeys,
  findNextPresentationSpellingIssue,
  presentationSpellingTags,
  type PresentationSpellingIssue,
} from "../../../../lib/presentation-spelling";

type AnyRecord = Record<string, unknown>;

/** What the window lends the dialog; called fresh on every action. */
export interface PresentationSpellingHost {
  slides(): readonly AnyRecord[];
  /** Change one occurrence; false when Replace refuses it (locked, stale). */
  replaceOne(match: PresentationFindMatch, word: string, replacement: string): boolean;
  /** Change every whole-word, same-case occurrence; how many changed. */
  replaceAll(word: string, replacement: string): number;
  /** Show the occurrence: go to its slide and select it, as Find does. */
  reveal(match: PresentationFindMatch): void;
}

/**
 * Review ▸ Spelling (PPT-A098): the desktop's Spelling check over the whole
 * deck from the current slide. Not in Dictionary / Suggestions / Ignore Once /
 * Ignore All / Add / Change / Change All, the same controls (and the same
 * shared strings) as the spreadsheet's Spelling dialog.
 */
export function PresentationSpellingDialog({
  host,
  startSlideIndex,
  locale,
  onClose,
  t,
  scope,
}: {
  host: PresentationSpellingHost;
  startSlideIndex: number;
  locale: string;
  onClose: () => void;
  t: (key: string, values?: Record<string, string | number>) => string;
  

  scope?: string;
}): ReactElement {
  const storage = useMemo(() => proofingStorage(), []);
  const options = useMemo(() => readSpellingOptions(storage), [storage]);
  const defaultTag = useMemo(
    () => defaultSpellingLanguage(locale, options.dictionaryLanguage),
    [locale, options.dictionaryLanguage],
  );
  // One dictionary per proofing language the deck's text is set in: each
  // word is judged in its own language, as PowerPoint's Spelling does.
  const [dictionaries, setDictionaries] = useState<ReadonlyMap<
    string,
    SpellingDictionary
  > | null>(null);
  const [loadFailed, setLoadFailed] = useState(false);
  // The deck as the dialog opened on it picks the dictionaries to load.
  const [openingSlides] = useState(() => host.slides());
  const [customWords, setCustomWords] = useState<string[]>(() =>
    readCustomDictionary(storage),
  );
  const [issue, setIssue] = useState<PresentationSpellingIssue | null>(null);
  const [done, setDone] = useState(false);
  const [replacement, setReplacement] = useState("");
  const skippedRef = useRef<Set<string>>(new Set());
  const changeRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    let cancelled = false;
    const tags = presentationSpellingTags(openingSlides, defaultTag);
    void Promise.all(
      tags.map((tag) =>
        loadSpellingDictionary(tag).then(
          (loaded) => [tag, loaded] as const,
          () => [tag, null] as const,
        ),
      ),
    ).then((loaded) => {
      if (cancelled) return;
      const ready = new Map<string, SpellingDictionary>();
      for (const [tag, dictionary] of loaded) if (dictionary) ready.set(tag, dictionary);
      // Another language's dictionary failing leaves its text unchecked; the
      // check's own failing is the dialog's "dictionary unavailable".
      if (ready.has(defaultTag)) setDictionaries(ready);
      else setLoadFailed(true);
    });
    return () => {
      cancelled = true;
    };
  }, [defaultTag, openingSlides]);

  const context = useCallback(
    (tag: string, words: readonly string[]): SpellingCheckContext | null => {
      const dictionary = dictionaries?.get(tag);
      return dictionary
        ? {
            dictionary,
            options,
            customWords: words,
            ignoredWords: spellingSessionIgnoredWords(scope),
          }
        : null;
    },
    [dictionaries, options],
  );

  // The dictionaries judge words in the spelling worker: a stop is shown
  // once every word before it has been judged (runSpellingCheckJudged), and
  // only the latest step shows its stop.
  const stepRef = useRef(0);
  const advance = useCallback(
    async (fromSlide: number, words: readonly string[] = customWords) => {
      if (!dictionaries) return;
      const step = ++stepRef.current;
      const all = () => dictionaries.values();
      const next = await runSpellingCheckJudged(all, () =>
        findNextPresentationSpellingIssue(
          host.slides(),
          fromSlide,
          (tag) => context(tag, words),
          skippedRef.current,
          defaultTag,
        ),
      );
      const checkContext = next ? context(next.tag, words) : null;
      const first = checkContext
        ? (await runSpellingCheckJudged(all, () => spellingSuggestions(next!.word, checkContext)))[0]
        : undefined;
      if (step !== stepRef.current) return;
      setIssue(next);
      setDone(next === null);
      if (next) {
        setReplacement(first ?? next.word);
        host.reveal(next.match);
      }
    },
    [context, customWords, defaultTag, dictionaries, host],
  );

  // The first stop once the dictionary is ready.
  const started = useRef(false);
  useEffect(() => {
    if (!dictionaries || started.current) return;
    started.current = true;
    void advance(startSlideIndex);
  }, [advance, dictionaries, startSlideIndex]);

  const suggestions = useMemo(() => {
    const checkContext = issue ? context(issue.tag, customWords) : null;
    return issue && checkContext ? spellingSuggestions(issue.word, checkContext) : [];
  }, [context, customWords, issue]);

  const here = issue?.match.slideIndex ?? startSlideIndex;
  const ignoreOnce = () => {
    if (!issue) return;
    skippedRef.current.add(issue.match.key);
    void advance(here);
  };
  const ignoreAll = () => {
    if (!issue) return;
    ignoreWordForSession(issue.word, scope);
    void advance(here);
  };
  const add = () => {
    if (!issue) return;
    const word = normalizeCustomDictionaryWord(issue.word);
    if (!word) return;
    const next = customWords.some((existing) => ignoreKey(existing) === ignoreKey(word))
      ? customWords
      : [...customWords, word];
    writeCustomDictionary(storage, next);
    setCustomWords(next);
    void advance(here, next);
  };
  const change = () => {
    if (!issue || !replacement || replacement === issue.word) return;
    if (!host.replaceOne(issue.match, issue.word, replacement)) {
      skippedRef.current.add(issue.match.key);
    } else {
      skippedRef.current = carrySkippedSpellingKeys(
        skippedRef.current,
        issue.match,
        replacement.length - issue.word.length,
      );
    }
    void advance(here);
  };
  const changeAll = () => {
    if (!issue || !replacement || replacement === issue.word) return;
    if (host.replaceAll(issue.word, replacement) === 0) {
      skippedRef.current.add(issue.match.key);
    }
    void advance(here);
  };

  const busy = !dictionaries && !loadFailed;
  return (
    <OfficeDialogShell
      open
      title={t("presentation.review.spelling")}
      description=""
      defaultActionLabel={t("spreadsheet.spelling.change")}
      cancelActionLabel={t("common.close")}
      defaultActionDisabled={!issue || !replacement || replacement === issue.word}
      onDefaultAction={change}
      onCancel={onClose}
      initialFocusRef={changeRef}
      testIdPrefix="presentation-spelling"
      defaultActionTestId="presentation-spelling-change"
    >
      <div role="status" aria-live="polite" data-testid="presentation-spelling-status">
        {loadFailed
          ? t("presentation.review.spellingDictionaryUnavailable")
          : busy
            ? t("spreadsheet.spelling.checking")
            : done
              ? t("spreadsheet.spelling.complete")
              : null}
      </div>
      {issue ? (
        <>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>{t("spreadsheet.spelling.notInDictionary")}</span>
            <output data-testid="presentation-spelling-word">{issue.word}</output>
          </label>
          <label className={dialogParts.field}>
            <span className={dialogParts.label}>{t("spreadsheet.spelling.with")}</span>
            <input
              ref={changeRef}
              type="text"
              value={replacement}
              onChange={(event) => setReplacement(event.currentTarget.value)}
              data-testid="presentation-spelling-replacement"
            />
          </label>
          <div className={dialogParts.field}>
            <span className={dialogParts.label}>{t("spreadsheet.spelling.suggestions")}</span>
            <select
              size={5}
              aria-label={t("spreadsheet.spelling.suggestions")}
              value={suggestions.includes(replacement) ? replacement : ""}
              onChange={(event) => setReplacement(event.currentTarget.value)}
              data-testid="presentation-spelling-suggestions"
            >
              {suggestions.length === 0 ? (
                <option value="" disabled>
                  {t("spreadsheet.spelling.noSuggestions")}
                </option>
              ) : (
                suggestions.map((suggestion) => (
                  <option key={suggestion} value={suggestion}>
                    {suggestion}
                  </option>
                ))
              )}
            </select>
          </div>
          <div className={dialogParts.row} style={{ flexWrap: "wrap", gap: 6 }}>
            <button type="button" className={dialogParts.compactButton} onClick={ignoreOnce} data-testid="presentation-spelling-ignore">
              {t("spreadsheet.spelling.ignoreOnce")}
            </button>
            <button type="button" className={dialogParts.compactButton} onClick={ignoreAll} data-testid="presentation-spelling-ignore-all">
              {t("spreadsheet.spelling.ignoreAll")}
            </button>
            <button type="button" className={dialogParts.compactButton} onClick={add} data-testid="presentation-spelling-add">
              {t("spreadsheet.spelling.add")}
            </button>
            <button
              type="button"
              className={dialogParts.compactButton}
              disabled={!replacement || replacement === issue.word}
              onClick={changeAll}
              data-testid="presentation-spelling-change-all"
            >
              {t("spreadsheet.spelling.changeAll")}
            </button>
          </div>
        </>
      ) : null}
    </OfficeDialogShell>
  );
}
