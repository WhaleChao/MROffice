"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Editor } from "@tiptap/react";
import type { Transaction } from "@tiptap/pm/state";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  ChevronDown,
  ChevronLeft,
  ChevronUp,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import {
  OfficeTaskPane,
  officeTaskPaneParts,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import {
  proofingStorage,
  readCustomDictionary,
  readSpellingOptions,
  writeCustomDictionary,
} from "@mrnote/core/web/lib/proofing/proofing-preferences.ts";
import {
  ignoreWordForSession,
  matchSpellingCase,
  spellingSessionIgnoredWords,
  spellingSuggestions,
  type SpellingDictionary,
} from "@mrnote/core/web/lib/proofing/spelling-checker.ts";
import {
  defaultDocumentSpellingLanguage,
  loadSpellingDictionary,
  spellingDictionariesPending,
  subscribeSpellingDictionaries,
} from "@mrnote/core/web/lib/proofing/spelling-dictionaries.ts";

import dockStyles from "@mrnote/core/web/components/console/editor/native-document/DocumentShapeFormatPane.module.css";
import styles from "./DocumentEditorPane.module.css";
import {
  collectDocumentEditorIssues,
  collectDocumentProofingBlocks,
  documentProofingDictionaryTags,
  documentUncheckedProofing,
  mapDocumentEditorIssueKey,
  type DocumentEditorIssue,
  type DocumentEditorIssueCategory,
  type DocumentEditorIssueRule,
} from "./documentEditorProofing";

/**
 * Review > Editor and Review > Spelling & Grammar.
 *
 * Word's Editor pane lists the document's corrections (Spelling, Grammar) and
 * refinements with a count each; choosing one steps through that kind with
 * the arrows at the top of the pane, a suggestion applies when chosen, and
 * every card offers Ignore Once and a way to stop flagging it ("Check
 * grammar, spelling, and more in Word", Microsoft Support). Spelling cards
 * add Ignore All and Add to Dictionary, as the Spelling & Grammar dialog
 * does. Spelling & Grammar opens the pane straight on the corrections, from
 * the insertion point on.
 *
 * Everything is checked in the browser (documentEditorProofing): no document
 * text is sent anywhere.
 */

export type DocumentEditorPaneView =
  | { readonly kind: "overview" }
  | { readonly kind: "step"; readonly filter: DocumentEditorStepFilter };

export type DocumentEditorStepFilter = DocumentEditorIssueCategory | "corrections";

interface DocumentEditorPaneProps {
  readonly editor: Editor;
  readonly onClose: () => void;
  /** Where the pane opens; a new object reopens it there. */
  readonly initialView?: DocumentEditorPaneView;
  /** Tests inject a dictionary; the app loads the shared Hunspell one. */
  readonly loadDictionary?: (tag: string) => Promise<SpellingDictionary>;
}

type RuntimeEditor = Editor & {
  on?: (event: string, callback: (props: unknown) => void) => unknown;
  off?: (event: string, callback: (props: unknown) => void) => unknown;
};

/** Ignore Once lasts as long as the document is open, as in Word. */
const ignoredOnceByEditor = new WeakMap<object, Set<string>>();
/** "Don't check for this issue" lasts for the session. */
const sessionDisabledRules = new Set<DocumentEditorIssueRule>();

function ignoredOnce(editor: Editor): Set<string> {
  let set = ignoredOnceByEditor.get(editor);
  if (!set) {
    const keys = new Set<string>();
    set = keys;
    ignoredOnceByEditor.set(editor, keys);
    // An ignored occurrence stays ignored while other words are corrected:
    // its range follows every edit, and editing its own text ends it.
    (editor as RuntimeEditor).on?.("transaction", (props: unknown) => {
      const transaction = (props as { transaction?: Transaction }).transaction;
      if (!transaction?.docChanged || keys.size === 0) return;
      const textBetween = (from: number, to: number) =>
        transaction.doc.textBetween(from, to, "", "\uFFFC");
      const mapped = Array.from(keys, (key) =>
        mapDocumentEditorIssueKey(key, transaction.mapping, textBetween),
      );
      keys.clear();
      for (const key of mapped) if (key) keys.add(key);
    });
  }
  return set;
}

/** A language or script as the UI language names it ("Finnish", "汉字"). */
function displayName(locale: string, type: "language" | "script", code: string): string {
  try {
    return new Intl.DisplayNames([locale], { type }).of(code) ?? code;
  } catch {
    return code;
  }
}

function displayList(locale: string, items: readonly string[]): string {
  const unique = [...new Set(items)];
  try {
    return new Intl.ListFormat([locale], { style: "short", type: "conjunction" }).format(unique);
  } catch {
    return unique.join(", ");
  }
}

function matchesFilter(
  issue: DocumentEditorIssue,
  filter: DocumentEditorStepFilter,
): boolean {
  return filter === "corrections"
    ? issue.category === "spelling" || issue.category === "grammar"
    : issue.category === filter;
}

export default function DocumentEditorPane({
  editor,
  onClose,
  initialView,
  loadDictionary = loadSpellingDictionary,
}: DocumentEditorPaneProps) {
  const t = useTranslations("console-notebooks");
  const locale = useLocale();
  const [doc, setDoc] = useState(() => editor.state.doc);
  const [canEdit, setCanEdit] = useState(
    () => !editor.isDestroyed && editor.isEditable,
  );
  const [revision, setRevision] = useState(0);
  const [dictionaries, setDictionaries] = useState<
    ReadonlyMap<string, SpellingDictionary>
  >(() => new Map());
  const [loading, setLoading] = useState<readonly string[]>([]);
  const [failed, setFailed] = useState<readonly string[]>([]);
  const [view, setView] = useState<DocumentEditorPaneView>(
    initialView ?? { kind: "overview" },
  );
  const [currentKey, setCurrentKey] = useState<string | null>(null);
  const anchorRef = useRef(editor.state.selection.from);

  useEffect(() => {
    if (!initialView) return;
    anchorRef.current = editor.state.selection.from;
    setView(initialView);
    setCurrentKey(null);
  }, [editor, initialView]);

  useEffect(() => {
    const runtime = editor as RuntimeEditor;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const refresh = () => {
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        timer = null;
        if (editor.isDestroyed) return;
        setDoc(editor.state.doc);
        setCanEdit(editor.isEditable);
      }, 250);
    };
    runtime.on?.("update", refresh);
    return () => {
      if (timer) clearTimeout(timer);
      runtime.off?.("update", refresh);
    };
  }, [editor]);

  const storage = proofingStorage();
  const spellingOptions = useMemo(
    () => readSpellingOptions(storage),
    // Re-read after Add to Dictionary or an options change elsewhere.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [revision],
  );
  const customWords = useMemo(
    () => readCustomDictionary(storage),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [revision],
  );
  const defaultLanguage = defaultDocumentSpellingLanguage(
    locale,
    spellingOptions.dictionaryLanguage,
  );
  const blocks = useMemo(() => collectDocumentProofingBlocks(doc), [doc]);
  const neededTags = useMemo(
    () => documentProofingDictionaryTags(blocks, defaultLanguage),
    [blocks, defaultLanguage],
  );

  const mountedRef = useRef(true);
  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
    };
  }, []);

  useEffect(() => {
    const missing = neededTags.filter(
      (tag) =>
        !dictionaries.has(tag) && !loading.includes(tag) && !failed.includes(tag),
    );
    if (missing.length === 0) return;
    setLoading((current) => [...current, ...missing]);
    // Each load outlives this effect: the first dictionary to arrive re-runs
    // it, and a document in several languages loads several at once. Only
    // the pane closing drops an answer, or the others would stay "loading".
    for (const tag of missing) {
      loadDictionary(tag).then(
        (dictionary) => {
          if (!mountedRef.current) return;
          setDictionaries((current) => new Map(current).set(tag, dictionary));
          setLoading((current) => current.filter((entry) => entry !== tag));
        },
        () => {
          if (!mountedRef.current) return;
          setFailed((current) => [...current, tag]);
          setLoading((current) => current.filter((entry) => entry !== tag));
        },
      );
    }
    // `loading`/`failed` are read as guards only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dictionaries, loadDictionary, neededTags]);

  // The dictionaries judge words in the spelling worker: a word not judged
  // yet reads as correct, and the check runs again when the verdicts arrive.
  // Answers arrive in batches: the check runs again once per burst.
  const [judged, setJudged] = useState(0);
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    const stop = subscribeSpellingDictionaries(dictionaries.values(), () => {
      if (timer) return;
      timer = setTimeout(() => {
        timer = null;
        setJudged((value) => value + 1);
      }, 40);
    });
    return () => {
      stop();
      if (timer) clearTimeout(timer);
    };
  }, [dictionaries]);

  const issues = useMemo(
    () =>
      collectDocumentEditorIssues(blocks, {
        defaultLanguage,
        dictionaries,
        spellingOptions,
        customWords,
        ignoredWords: spellingSessionIgnoredWords(editor),
        disabledRules: sessionDisabledRules,
        ignoredIssues: ignoredOnce(editor),
      }),
    // `revision` re-runs the check after Ignore / Add to Dictionary.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [blocks, customWords, defaultLanguage, dictionaries, editor, judged, revision, spellingOptions],
  );
  const judging = useMemo(
    () => spellingDictionariesPending(dictionaries.values()),
    // `issues` asked the dictionaries; `judged` moves when they answer.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [dictionaries, issues, judged],
  );
  const unchecked = useMemo(
    () => documentUncheckedProofing(blocks, defaultLanguage),
    [blocks, defaultLanguage],
  );

  const counts = useMemo(() => {
    const result: Record<DocumentEditorIssueCategory, number> = {
      spelling: 0,
      grammar: 0,
      refinement: 0,
    };
    for (const issue of issues) result[issue.category] += 1;
    return result;
  }, [issues]);

  const stepIssues = useMemo(
    () =>
      view.kind === "step"
        ? issues.filter((issue) => matchesFilter(issue, view.filter))
        : [],
    [issues, view],
  );
  const current = useMemo(() => {
    if (view.kind !== "step" || stepIssues.length === 0) return null;
    const byKey = currentKey
      ? stepIssues.find((issue) => issue.key === currentKey)
      : undefined;
    if (byKey) return byKey;
    // The next issue from where the check is (the insertion point at first,
    // then the issue just handled), wrapping to the top as Word does.
    return (
      stepIssues.find((issue) => issue.from >= anchorRef.current) ??
      stepIssues[0]
    );
  }, [currentKey, stepIssues, view.kind]);
  const currentIndex = current ? stepIssues.indexOf(current) : -1;

  const select = useCallback(
    (issue: DocumentEditorIssue) => {
      if (editor.isDestroyed) return;
      const size = editor.state.doc.content.size;
      if (issue.to > size) return;
      editor.commands.setTextSelection({ from: issue.from, to: issue.to });
      editor.commands.scrollIntoView();
    },
    [editor],
  );

  useEffect(() => {
    if (current) select(current);
  }, [current, select]);

  const suggestions = useMemo(() => {
    if (!current) return [];
    if (current.suggestions) return current.suggestions;
    const dictionary = current.language
      ? dictionaries.get(current.language)
      : undefined;
    if (!dictionary) return [];
    // Offered in the flagged word's own capitalization, as Word does.
    return spellingSuggestions(current.text, {
      dictionary,
      options: spellingOptions,
      customWords,
      ignoredWords: spellingSessionIgnoredWords(editor),
    })
      .map((suggestion) => matchSpellingCase(current.text, suggestion))
      .filter((suggestion, index, list) => list.indexOf(suggestion) === index)
      .slice(0, 5);
    // `judged`: the suggestions arrive from the spelling worker.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [current, customWords, dictionaries, judged, spellingOptions]);

  const advance = useCallback((issue: DocumentEditorIssue) => {
    anchorRef.current = issue.from;
    setCurrentKey(null);
    setRevision((value) => value + 1);
  }, []);

  const apply = useCallback(
    (issue: DocumentEditorIssue, replacement: string) => {
      if (editor.isDestroyed || !editor.isEditable) return;
      const { state } = editor;
      if (issue.to > state.doc.content.size) return;
      // Only replace the text the card was made for.
      if (state.doc.textBetween(issue.from, issue.to, "", "￼") !== issue.rangeText) {
        return;
      }
      const tr = state.tr.insertText(replacement, issue.from, issue.to);
      editor.view.dispatch(tr);
      setDoc(editor.state.doc);
      advance(issue);
    },
    [advance, editor],
  );

  const ignoreOnce = useCallback(
    (issue: DocumentEditorIssue) => {
      ignoredOnce(editor).add(issue.key);
      advance(issue);
    },
    [advance, editor],
  );

  const ignoreAll = useCallback(
    (issue: DocumentEditorIssue) => {
      ignoreWordForSession(issue.text, editor);
      advance(issue);
    },
    [advance],
  );

  const addToDictionary = useCallback(
    (issue: DocumentEditorIssue) => {
      writeCustomDictionary(proofingStorage(), [...customWords, issue.text]);
      advance(issue);
    },
    [advance, customWords],
  );

  const stopChecking = useCallback(
    (issue: DocumentEditorIssue) => {
      sessionDisabledRules.add(issue.rule);
      advance(issue);
    },
    [advance],
  );

  const move = useCallback(
    (direction: 1 | -1) => {
      if (stepIssues.length === 0) return;
      const index =
        (Math.max(currentIndex, 0) + direction + stepIssues.length) %
        stepIssues.length;
      setCurrentKey(stepIssues[index].key);
    },
    [currentIndex, stepIssues],
  );

  const openStep = useCallback(
    (filter: DocumentEditorStepFilter) => {
      anchorRef.current = editor.state.selection.from;
      setCurrentKey(null);
      setView({ kind: "step", filter });
    },
    [editor],
  );

  const categoryLabel = (filter: DocumentEditorStepFilter) =>
    t(
      `toolbar.documentEditor.category.${filter}` as "toolbar.documentEditor.category.spelling",
    );
  const total = counts.spelling + counts.grammar + counts.refinement;
  const busy = loading.length > 0 || judging;

  const categoryRow = (category: DocumentEditorIssueCategory) => (
    <li key={category}>
      <button
        type="button"
        className={styles.categoryRow}
        data-testid={`document-editor-category-${category}`}
        disabled={counts[category] === 0}
        onClick={() => openStep(category)}
      >
        <span>{categoryLabel(category)}</span>
        <span className={styles.count} data-testid={`document-editor-count-${category}`}>
          {counts[category]}
        </span>
      </button>
    </li>
  );

  return (
    <OfficeTaskPane
      title={t("toolbar.documentEditor.title")}
      closeLabel={t("toolbar.documentEditor.close")}
      onClose={onClose}
      testIdPrefix="document-editor"
      testId="document-editor-pane"
      closeTestId="document-editor-close"
      dockRight={false}
      className={dockStyles.pane}
    >
      {view.kind === "overview" ? (
        <>
          <div className={officeTaskPaneParts.section}>
            <p className={styles.summary} data-testid="document-editor-total" aria-live="polite">
              {busy
                ? t("toolbar.documentEditor.loading")
                : total === 0
                  ? t("toolbar.documentEditor.noSuggestions")
                  : t("toolbar.documentEditor.suggestionCount", { count: total })}
            </p>
            {failed.length > 0 ? (
              <p className={officeTaskPaneParts.hint} role="status">
                {t("toolbar.documentEditor.dictionaryFailed")}
              </p>
            ) : null}
          </div>
          <div className={officeTaskPaneParts.section}>
            <span className={officeTaskPaneParts.sectionLabel}>
              {t("toolbar.documentEditor.corrections")}
            </span>
            <ul className={styles.categoryList}>
              {categoryRow("spelling")}
              {categoryRow("grammar")}
            </ul>
          </div>
          <div className={officeTaskPaneParts.section}>
            <span className={officeTaskPaneParts.sectionLabel}>
              {t("toolbar.documentEditor.refinements")}
            </span>
            <ul className={styles.categoryList}>{categoryRow("refinement")}</ul>
            <p className={officeTaskPaneParts.hint}>
              {t("toolbar.documentEditor.languageNote", {
                language: displayName(locale, "language", defaultLanguage),
              })}
            </p>
            {unchecked.noWordSpaces.length > 0 ? (
              <p
                className={officeTaskPaneParts.hint}
                data-testid="document-editor-unchecked-no-spaces"
              >
                {t("toolbar.documentEditor.uncheckedNoWordSpaces", {
                  scripts: displayList(
                    locale,
                    unchecked.noWordSpaces.map((code) => displayName(locale, "script", code)),
                  ),
                })}
              </p>
            ) : null}
            {unchecked.noDictionary.length > 0 ? (
              <p
                className={officeTaskPaneParts.hint}
                data-testid="document-editor-unchecked-no-dictionary"
              >
                {t("toolbar.documentEditor.uncheckedNoDictionary", {
                  languages: displayList(
                    locale,
                    unchecked.noDictionary.map((tag) => displayName(locale, "language", tag)),
                  ),
                })}
              </p>
            ) : null}
            {unchecked.otherScripts.length > 0 ? (
              <p
                className={officeTaskPaneParts.hint}
                data-testid="document-editor-unchecked-other-scripts"
              >
                {t("toolbar.documentEditor.uncheckedOtherScripts", {
                  scripts: displayList(
                    locale,
                    unchecked.otherScripts.map((code) => displayName(locale, "script", code)),
                  ),
                })}
              </p>
            ) : null}
          </div>
        </>
      ) : (
        <div className={officeTaskPaneParts.section}>
          <div className={styles.stepHeader}>
            <button
              type="button"
              className={officeTaskPaneParts.button}
              data-testid="document-editor-back"
              onClick={() => setView({ kind: "overview" })}
            >
              <ChevronLeft size={14} aria-hidden="true" />
              {t("toolbar.documentEditor.back")}
            </button>
            <span className={officeTaskPaneParts.sectionLabel}>
              {categoryLabel(view.filter)}
            </span>
            <span className={styles.stepHeader}>
              <span
                className={officeTaskPaneParts.hint}
                data-testid="document-editor-position"
                aria-live="polite"
              >
                {current ? `${currentIndex + 1}/${stepIssues.length}` : "0/0"}
              </span>
              <button
                type="button"
                className={officeTaskPaneParts.iconButton}
                aria-label={t("findReplace.previous")}
                title={t("findReplace.previous")}
                data-testid="document-editor-previous"
                disabled={stepIssues.length < 2}
                onClick={() => move(-1)}
              >
                <ChevronUp size={16} aria-hidden="true" />
              </button>
              <button
                type="button"
                className={officeTaskPaneParts.iconButton}
                aria-label={t("findReplace.next")}
                title={t("findReplace.next")}
                data-testid="document-editor-next"
                disabled={stepIssues.length < 2}
                onClick={() => move(1)}
              >
                <ChevronDown size={16} aria-hidden="true" />
              </button>
            </span>
          </div>
          {current ? (
            <div
              className={styles.card}
              data-category={current.category}
              data-rule={current.rule}
              data-testid="document-editor-card"
            >
              <p className={styles.ruleTitle}>
                {t(
                  `toolbar.documentEditor.rule.${current.rule}` as "toolbar.documentEditor.rule.misspelled",
                )}
              </p>
              <p className={styles.context} data-testid="document-editor-context">
                {current.before}
                <span className={styles.flagged}>{current.text}</span>
                {current.after}
              </p>
              {current.rule !== "sentenceLength" ? (
                <div className={styles.suggestions}>
                  {suggestions.length === 0 ? (
                    <span className={officeTaskPaneParts.hint}>
                      {t("toolbar.documentEditor.noReplacement")}
                    </span>
                  ) : (
                    suggestions.map((suggestion, index) => (
                      <button
                        key={`${index}:${suggestion}`}
                        type="button"
                        className={styles.suggestion}
                        data-testid={`document-editor-suggestion-${index}`}
                        disabled={!canEdit}
                        onClick={() => apply(current, suggestion)}
                      >
                        {suggestion === ""
                          ? current.rule === "repeated"
                            ? t("toolbar.documentEditor.deleteRepeated")
                            : t("toolbar.documentEditor.deleteSpace")
                          : suggestion === " "
                            ? t("toolbar.documentEditor.singleSpace")
                            : suggestion}
                      </button>
                    ))
                  )}
                </div>
              ) : null}
              <div className={styles.actions}>
                <button
                  type="button"
                  className={officeTaskPaneParts.button}
                  data-testid="document-editor-ignore-once"
                  onClick={() => ignoreOnce(current)}
                >
                  {t("toolbar.documentEditor.ignoreOnce")}
                </button>
                {current.category === "spelling" && current.rule === "misspelled" ? (
                  <>
                    <button
                      type="button"
                      className={officeTaskPaneParts.button}
                      data-testid="document-editor-ignore-all"
                      onClick={() => ignoreAll(current)}
                    >
                      {t("toolbar.documentEditor.ignoreAll")}
                    </button>
                    <button
                      type="button"
                      className={officeTaskPaneParts.button}
                      data-testid="document-editor-add-to-dictionary"
                      onClick={() => addToDictionary(current)}
                    >
                      {t("toolbar.documentEditor.addToDictionary")}
                    </button>
                  </>
                ) : (
                  <button
                    type="button"
                    className={officeTaskPaneParts.button}
                    data-testid="document-editor-stop-checking"
                    onClick={() => stopChecking(current)}
                  >
                    {t("toolbar.documentEditor.stopChecking")}
                  </button>
                )}
              </div>
            </div>
          ) : (
            <p
              className={officeTaskPaneParts.hint}
              data-testid="document-editor-done"
              role="status"
            >
              {busy
                ? t("toolbar.documentEditor.loading")
                : t("toolbar.documentEditor.done")}
            </p>
          )}
        </div>
      )}
    </OfficeTaskPane>
  );
}
