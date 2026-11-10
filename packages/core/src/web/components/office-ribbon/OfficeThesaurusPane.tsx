"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Review > Thesaurus: the body of Office's Thesaurus task pane, shared by the
 * editors. Microsoft Support ("Look up words in the thesaurus"): the pane
 * opens on the selected word; it lists the word's meanings, each with its
 * synonyms; a synonym is looked up in turn by clicking it (Back returns),
 * and its menu inserts it in place of the selection or copies it. The host
 * supplies the lookup (the model, so every language works) and the insert.
 */

import { useCallback, useEffect, useRef, useState, type FormEvent } from "react";

import { ArrowLeft, Copy, Search } from "./icons/houseIcons";

import styles from "./OfficeThesaurusPane.module.css";

export interface OfficeThesaurusSense {
  readonly meaning: string;
  readonly partOfSpeech: string | null;
  readonly synonyms: readonly string[];
}

export interface OfficeThesaurusEntry {
  readonly word: string;
  readonly language: string | null;
  readonly senses: readonly OfficeThesaurusSense[];
  readonly antonyms: readonly string[];
}

export interface OfficeThesaurusLabels {
  readonly search: string;
  readonly back: string;
  readonly insert: string;
  readonly copy: string;
  readonly loading: string;
  readonly empty: string;
  readonly failed: string;
  readonly antonyms: string;
  readonly prompt: string;
}

type LookupState =
  | { readonly kind: "idle" }
  | { readonly kind: "loading"; readonly word: string }
  | { readonly kind: "done"; readonly entry: OfficeThesaurusEntry }
  | { readonly kind: "failed"; readonly word: string };

export function OfficeThesaurusPane({
  initialWord,
  lookup,
  canInsert,
  onInsert,
  labels,
  testIdPrefix,
}: {
  /** The selected word the pane opens on; empty opens on the search box. */
  initialWord: string;
  lookup: (word: string) => Promise<OfficeThesaurusEntry>;
  /** Insert replaces the word the pane was opened on; false greys it. */
  canInsert: boolean;
  onInsert: (synonym: string) => void;
  labels: OfficeThesaurusLabels;
  testIdPrefix: string;
}) {
  const [query, setQuery] = useState(initialWord);
  const [state, setState] = useState<LookupState>({ kind: "idle" });
  const [history, setHistory] = useState<readonly string[]>([]);
  const requestRef = useRef(0);

  const run = useCallback(
    (word: string, remember: string | null) => {
      const trimmed = word.trim();
      if (!trimmed) return;
      const request = requestRef.current + 1;
      requestRef.current = request;
      setQuery(trimmed);
      if (remember) setHistory((past) => [...past, remember]);
      setState({ kind: "loading", word: trimmed });
      lookup(trimmed).then(
        (entry) => {
          if (requestRef.current === request) setState({ kind: "done", entry });
        },
        () => {
          if (requestRef.current === request) setState({ kind: "failed", word: trimmed });
        },
      );
    },
    [lookup],
  );

  // The pane opens on the selected word, and a new selection's word when
  // the host re-opens it on another one.
  useEffect(() => {
    setHistory([]);
    if (initialWord.trim()) run(initialWord, null);
    else {
      setQuery("");
      setState({ kind: "idle" });
    }
  }, [initialWord, run]);

  const current =
    state.kind === "done" ? state.entry.word : state.kind === "idle" ? null : state.word;
  const submit = (event: FormEvent) => {
    event.preventDefault();
    if (query.trim() && query.trim() !== current) run(query, current);
  };
  const copy = (word: string) => {
    try {
      void navigator.clipboard?.writeText(word).catch(() => undefined);
    } catch {
      // A browser without clipboard access keeps the pane working.
    }
  };
  const wordRow = (word: string, key: string) => (
    <li key={key} className={styles.wordRow}>
      <button
        type="button"
        className={styles.word}
        title={labels.search}
        data-testid={`${testIdPrefix}-word-${word}`}
        onClick={() => run(word, current)}
      >
        {word}
      </button>
      <button
        type="button"
        className={styles.action}
        disabled={!canInsert}
        title={labels.insert}
        aria-label={`${labels.insert} · ${word}`}
        data-testid={`${testIdPrefix}-insert-${word}`}
        onClick={() => onInsert(word)}
      >
        {labels.insert}
      </button>
      <button
        type="button"
        className={styles.iconAction}
        title={labels.copy}
        aria-label={`${labels.copy} · ${word}`}
        data-testid={`${testIdPrefix}-copy-${word}`}
        onClick={() => copy(word)}
      >
        <Copy size={13} aria-hidden />
      </button>
    </li>
  );

  return (
    <div className={styles.pane} data-testid={`${testIdPrefix}-body`}>
      <form className={styles.searchRow} role="search" onSubmit={submit}>
        <button
          type="button"
          className={styles.iconAction}
          disabled={history.length === 0}
          title={labels.back}
          aria-label={labels.back}
          data-testid={`${testIdPrefix}-back`}
          onClick={() => {
            const previous = history[history.length - 1];
            if (!previous) return;
            setHistory((past) => past.slice(0, -1));
            run(previous, null);
          }}
        >
          <ArrowLeft size={14} aria-hidden />
        </button>
        <input
          className={styles.searchInput}
          value={query}
          aria-label={labels.search}
          placeholder={labels.search}
          data-testid={`${testIdPrefix}-search`}
          onChange={(event) => setQuery(event.target.value)}
        />
        <button
          type="submit"
          className={styles.iconAction}
          title={labels.search}
          aria-label={labels.search}
          data-testid={`${testIdPrefix}-search-button`}
        >
          <Search size={14} aria-hidden />
        </button>
      </form>
      <div className={styles.results} aria-live="polite" aria-busy={state.kind === "loading"}>
        {state.kind === "idle" ? <p className={styles.note}>{labels.prompt}</p> : null}
        {state.kind === "loading" ? <p className={styles.note}>{labels.loading}</p> : null}
        {state.kind === "failed" ? (
          <p className={styles.note} data-testid={`${testIdPrefix}-failed`}>
            {labels.failed}
          </p>
        ) : null}
        {state.kind === "done" && state.entry.senses.length === 0 ? (
          <p className={styles.note} data-testid={`${testIdPrefix}-empty`}>
            {labels.empty}
          </p>
        ) : null}
        {state.kind === "done"
          ? state.entry.senses.map((sense, index) => (
              <section key={`${sense.meaning}-${index}`} className={styles.sense}>
                <h3 className={styles.meaning}>
                  {sense.meaning}
                  {sense.partOfSpeech ? (
                    <span className={styles.partOfSpeech}> ({sense.partOfSpeech})</span>
                  ) : null}
                </h3>
                <ul className={styles.words}>
                  {sense.synonyms.map((word) => wordRow(word, `${index}-${word}`))}
                </ul>
              </section>
            ))
          : null}
        {state.kind === "done" && state.entry.antonyms.length > 0 ? (
          <section className={styles.sense}>
            <h3 className={styles.meaning}>{labels.antonyms}</h3>
            <ul className={styles.words}>
              {state.entry.antonyms.map((word) => wordRow(word, `antonym-${word}`))}
            </ul>
          </section>
        ) : null}
      </div>
    </div>
  );
}
