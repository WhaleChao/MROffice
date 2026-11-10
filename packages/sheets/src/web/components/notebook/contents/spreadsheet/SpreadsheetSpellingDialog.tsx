"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { containDialogTab } from "@mrnote/core/web/lib/dialog-tab-containment.ts";
import {
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import { X } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { SpreadsheetDialogIcon, type SpreadsheetDialogIconName } from "@mrnote/core/web/components/notebook/contents/spreadsheet/SpreadsheetDialogIcon.tsx";

/** The action column's pictures; the footer's answers stay words. */
const SPELLING_ICON: Record<string, SpreadsheetDialogIconName> = {
  "ignore-once": "ignoreOnce",
  "ignore-all": "ignoreAll",
  add: "addToDictionary",
  delete: "delete",
  change: "replace",
  "change-all": "replace",
  autocorrect: "autoCorrect",
  options: "options",
  "undo-last": "undo",
};
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  officeDialogParts as dialogParts,
  spreadsheetDialogClassName,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  normalizeAutoCorrectEntry,
  normalizeCustomDictionaryWord,
  proofingStorage,
  readAutoCorrectEntries,
  readCustomDictionary,
  readSpellingOptions,
  writeAutoCorrectEntries,
  writeCustomDictionary,
  writeSpellingOptions,
  type AutoCorrectEntry,
  type SpellingOptions,
} from "@mrnote/core/web/lib/proofing/proofing-preferences.ts";
import {
  ignoreWordForSession,
  spellingSessionIgnoredWords,
  spellingSuggestions,
  stopIgnoringWordForSession,
  type SpellingCheckContext,
} from "@mrnote/core/web/lib/proofing/spelling-checker.ts";
import {
  SPELLING_DICTIONARIES,
  defaultSpellingLanguage,
  loadSpellingDictionary,
  runSpellingCheckJudged,
  type LoadedSpellingDictionary,
} from "@mrnote/core/web/lib/proofing/spelling-dictionaries.ts";
import { columnToLetters } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { isSpreadsheetImeKeydown } from "../../../../lib/spreadsheet/ime-composition";
import {
  SPELLING_START,
  buildSpellingPlan,
  changeAllInSpellingPlan,
  findNextSpellingIssue,
  readSpellingTarget,
  repeatedWordDeletion,
  replaceSpellingTargetText,
  writeSpellingTarget,
  type SpellingCursor,
  type SpellingFinding,
  type SpellingPlan,
  type SpellingScope,
  type SpellingSheetLike,
  type SpellingTarget,
  type SpellingTextEdit,
} from "../../../../lib/spreadsheet/spelling-walk";

import styles from "./SpreadsheetSpellingDialog.module.css";

export interface SpreadsheetSpellingSheet extends SpellingSheetLike {
  id: string;
}

/**
 * What the dialog needs from the window. The window builds a new host every
 * render; the dialog always calls the latest one, so a commit builds on the
 * sheet as rendered now.
 */
export interface SpreadsheetSpellingHost {
  sheet(): SpreadsheetSpellingSheet;
  /** One corrected sheet, committed as one undoable edit. False if refused
   *  (the host has said why). */
  commit(next: SpreadsheetSpellingSheet, edits: readonly SpellingTextEdit[]): boolean;
  /** True, or the reason the target's text cannot be changed now. */
  writable(target: SpellingTarget): true | string;
  /** Put the selection on the flagged text, as Excel does. */
  reveal(target: SpellingTarget): void;
  notify(message: string): void;
}

export interface SpreadsheetSpellingLaunch {
  /** A new id is a new check (F7 pressed again). */
  id: number;
  sheetId: string;
  scope: SpellingScope;
  /** Cells whose shown value is not their own (spilled arrays). */
  skipCellKeys?: { has(key: string): boolean };
}

export interface SpreadsheetSpellingDialogProps {
  

  scope?: string;
  launch: SpreadsheetSpellingLaunch;
  host: SpreadsheetSpellingHost;
  onClose: () => void;
}

type Phase =
  | { kind: "loading" }
  | { kind: "checking" }
  | { kind: "issue"; finding: SpellingFinding; suggestions: string[] }
  | { kind: "wrap"; cursor: SpellingCursor }
  | { kind: "complete" }
  | { kind: "failed" };

type UndoRecord =
  | { kind: "skip"; finding: SpellingFinding }
  | { kind: "ignore-all"; finding: SpellingFinding; word: string }
  | { kind: "add"; finding: SpellingFinding; word: string }
  | {
      kind: "edit";
      finding: SpellingFinding;
      edits: SpellingTextEdit[];
      autoCorrect?: AutoCorrectEntry;
    };

/** A slice of checking between two yields to the browser. */
const SLICE_MS = 12;

const isPageTarget = (target: SpellingTarget) => target.kind === "header-footer";
const COMMIT_GROUPS: ReadonlyArray<(target: SpellingTarget) => boolean> = [
  (target) => !isPageTarget(target),
  isPageTarget,
];

function nextTask(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function languageName(locale: string, tag: string): string {
  try {
    return (
      new Intl.DisplayNames([locale], { type: "language", languageDisplay: "standard" }).of(
        tag,
      ) ?? tag
    );
  } catch {
    return tag;
  }
}

/**
 * Excel's Spelling dialog: Not in Dictionary (or Repeated Word), Suggestions,
 * Ignore Once / Ignore All / Add to Dictionary, Change (Delete for a repeated
 * word) / Change All / AutoCorrect, Dictionary language, Options, Undo Last
 * and Cancel, which turns into Close once a change has been made (Microsoft
 * Support, "Check spelling in a worksheet"). It walks the scope lazily, one
 * time slice at a time, so a large sheet never freezes the window.
 */
export function SpreadsheetSpellingDialog({
  launch,
  host,
  onClose,
  scope,
}: SpreadsheetSpellingDialogProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const locale = useLocale();
  const titleId = useId();
  const wordId = useId();
  const suggestionsId = useId();
  const languageId = useId();

  const hostRef = useRef(host);
  const onCloseRef = useRef(onClose);
  useLayoutEffect(() => {
    hostRef.current = host;
    onCloseRef.current = onClose;
  });

  const [initial] = useState(() => {
    const storage = proofingStorage();
    const options = readSpellingOptions(storage);
    return {
      options,
      customWords: readCustomDictionary(storage),
      language: defaultSpellingLanguage(locale, options.dictionaryLanguage),
    };
  });
  const [options, setOptionsState] = useState<SpellingOptions>(initial.options);
  const [customWords, setCustomWordsState] = useState<string[]>(initial.customWords);
  const [language, setLanguage] = useState(initial.language);
  const [phase, setPhase] = useState<Phase>({ kind: "loading" });
  const [draft, setDraft] = useState("");
  const [selected, setSelected] = useState(-1);
  const [busy, setBusy] = useState(false);
  const [changed, setChanged] = useState(false);
  const [undoDepth, setUndoDepth] = useState(0);
  const [optionsOpen, setOptionsOpen] = useState(false);

  // The scan runs outside React's render, so it reads the settings from here.
  const live = useRef({
    options: initial.options,
    customWords: initial.customWords,
    dictionary: null as LoadedSpellingDictionary | null,
    plan: null as SpellingPlan | null,
    generation: 0,
    undo: [] as UndoRecord[],
  });

  const rootRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const firstActionRef = useRef<HTMLButtonElement | null>(null);
  const defaultButtonRef = useRef<HTMLButtonElement | null>(null);

  const setOptions = (next: SpellingOptions) => {
    live.current.options = next;
    setOptionsState(next);
    writeSpellingOptions(proofingStorage(), next);
  };
  const setCustomWords = (next: string[]) => {
    live.current.customWords = next;
    setCustomWordsState(next);
    writeCustomDictionary(proofingStorage(), next);
  };

  const context = (): SpellingCheckContext | null => {
    const dictionary = live.current.dictionary;
    if (!dictionary) return null;
    return {
      dictionary,
      options: live.current.options,
      customWords: live.current.customWords,
      ignoredWords: spellingSessionIgnoredWords(scope),
    };
  };

  const showFinding = async (finding: SpellingFinding, generation: number) => {
    const ctx = context();
    // Suggestions come from the spelling worker: the stop shows with them.
    const suggestions =
      finding.kind === "repeated" || !ctx
        ? []
        : await runSpellingCheckJudged(
            () => [ctx.dictionary],
            () => spellingSuggestions(finding.word, ctx),
          );
    if (generation !== live.current.generation) return;
    setPhase({ kind: "issue", finding, suggestions });
    setDraft(finding.word);
    setSelected(suggestions.length > 0 ? 0 : -1);
    hostRef.current.reveal(finding.target);
  };

  const scan = async (from: SpellingCursor) => {
    const generation = ++live.current.generation;
    const plan = live.current.plan;
    if (!plan || !context()) return;
    let cursor = from;
    for (;;) {
      if (generation !== live.current.generation) return;
      const ctx = context();
      if (!ctx) return;
      const sheet = hostRef.current.sheet();
      // Excel's check belongs to one sheet; another sheet ends it.
      if (sheet.id !== launch.sheetId) {
        onCloseRef.current();
        return;
      }
      const sliceEnd = performance.now() + SLICE_MS;
      const step = findNextSpellingIssue({
        sheet,
        plan,
        cursor,
        context: ctx,
        deadline: () => performance.now() >= sliceEnd,
      });
      // A slice that met words the spelling worker has not judged yet is
      // run again once it has: the verdicts are kept, so it gets further.
      if (live.current.dictionary?.pending()) {
        await live.current.dictionary.settled();
        continue;
      }
      if (step.status === "paused") {
        cursor = step.cursor;
        setPhase((current) => (current.kind === "checking" ? current : { kind: "checking" }));
        await nextTask();
        continue;
      }
      if (step.status === "issue") void showFinding(step.finding, generation);
      else if (step.status === "wrap") setPhase({ kind: "wrap", cursor: step.cursor });
      else setPhase({ kind: "complete" });
      return;
    }
  };

  const load = async (tag: string, resume: SpellingCursor) => {
    const generation = ++live.current.generation;
    setPhase({ kind: "loading" });
    try {
      const dictionary = await loadSpellingDictionary(tag);
      if (generation !== live.current.generation) return;
      live.current.dictionary = dictionary;
    } catch {
      if (generation === live.current.generation) setPhase({ kind: "failed" });
      return;
    }
    // Planned once the dictionary is here: by then an entry that was being
    // edited has been committed by the focus leaving it.
    live.current.plan ??= buildSpellingPlan(hostRef.current.sheet(), launch.scope, {
      skipCellKeys: launch.skipCellKeys,
    });
    void scan(resume);
  };

  useEffect(() => {
    // Taking the focus ends a cell entry in progress (it commits on blur),
    // which is the text an entry-mode check reads.
    rootRef.current?.focus({ preventScroll: true });
    void load(initial.language, SPELLING_START);
    const state = live.current;
    return () => {
      state.generation += 1;
    };
    // One check per mount; the window remounts the dialog for a new launch.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const finding = phase.kind === "issue" ? phase.finding : null;
  const suggestions = phase.kind === "issue" ? phase.suggestions : [];
  const repeated = finding?.kind === "repeated";
  const edited = finding !== null && draft !== finding.word;
  const replacement: string | null = !finding
    ? null
    : edited
      ? draft
      : selected >= 0
        ? (suggestions[selected] ?? null)
        : null;

  // A new flag puts the focus where Excel does: on the suggestions, else on
  // the first command.
  useEffect(() => {
    if (phase.kind !== "issue") return;
    (phase.suggestions.length > 0 ? listRef.current : firstActionRef.current)?.focus({
      preventScroll: true,
    });
  }, [phase]);

  const pushUndo = (record: UndoRecord) => {
    live.current.undo.push(record);
    setUndoDepth(live.current.undo.length);
  };

  /** Wait for the window to render the sheet a commit produced. */
  const sheetAfter = async (before: SpreadsheetSpellingSheet) => {
    for (let attempt = 0; attempt < 40; attempt += 1) {
      if (hostRef.current.sheet() !== before) return;
      await new Promise((resolve) => setTimeout(resolve, 15));
    }
  };

  /**
   * Commit `make`'s edits group by group — cell text, notes and shapes first,
   * then headers and footers, which the window writes through Page Setup —
   * each built on the sheet the previous one left.
   */
  const commitInGroups = async (
    make: (
      sheet: SpreadsheetSpellingSheet,
      inGroup: (target: SpellingTarget) => boolean,
    ) => { sheet: SpreadsheetSpellingSheet; edits: SpellingTextEdit[] } | null,
  ): Promise<SpellingTextEdit[]> => {
    const committed: SpellingTextEdit[] = [];
    for (const inGroup of COMMIT_GROUPS) {
      const before = hostRef.current.sheet();
      const result = make(before, inGroup);
      if (!result || result.edits.length === 0) continue;
      if (!hostRef.current.commit(result.sheet, result.edits)) break;
      committed.push(...result.edits);
      await sheetAfter(before);
    }
    return committed;
  };

  const ignoreOnce = () => {
    if (!finding || busy) return;
    pushUndo({ kind: "skip", finding });
    void scan({ ...finding.cursor, offset: finding.end });
  };

  const ignoreAll = () => {
    if (!finding || busy || repeated) return;
    ignoreWordForSession(finding.word, scope);
    pushUndo({ kind: "ignore-all", finding, word: finding.word });
    void scan(finding.cursor);
  };

  const dictionaryWord = finding && !repeated ? normalizeCustomDictionaryWord(finding.word) : null;
  const addToDictionary = () => {
    if (!finding || busy || !dictionaryWord) return;
    const stored = readCustomDictionary(proofingStorage());
    const base = stored.length > 0 ? stored : live.current.customWords;
    if (!base.includes(dictionaryWord)) setCustomWords([...base, dictionaryWord]);
    pushUndo({ kind: "add", finding, word: dictionaryWord });
    void scan(finding.cursor);
  };

  const change = async (all: boolean, autoCorrect: boolean) => {
    if (!finding || busy || replacement === null || repeated) return;
    if (replacement === finding.word) {
      if (all) ignoreAll();
      else ignoreOnce();
      return;
    }
    const writable = hostRef.current.writable(finding.target);
    if (writable !== true) {
      hostRef.current.notify(writable);
      return;
    }
    const plan = live.current.plan;
    if (!plan) return;
    setBusy(true);
    try {
      let resumeOffset = finding.start + replacement.length;
      const edits = await commitInGroups((sheet, inGroup) => {
        if (!all) {
          if (!inGroup(finding.target)) return null;
          const written = writeSpellingTarget(
            sheet,
            finding.target,
            finding.start,
            finding.end,
            replacement,
            finding.word,
          );
          return written ? { sheet: written.sheet, edits: [written.edit] } : null;
        }
        const result = changeAllInSpellingPlan({
          sheet,
          plan,
          word: finding.word,
          replacement,
          options: live.current.options,
          current: finding,
          canWrite: (target) => inGroup(target) && hostRef.current.writable(target) === true,
        });
        if (inGroup(finding.target)) resumeOffset = result.resumeOffset;
        return { sheet: result.sheet, edits: result.edits };
      });
      if (edits.length === 0) {
        // The text moved under the dialog: look at this place again.
        void scan(finding.cursor);
        return;
      }
      let entry: AutoCorrectEntry | undefined;
      if (autoCorrect) {
        entry = normalizeAutoCorrectEntry({ from: finding.word, to: replacement }) ?? undefined;
        if (entry) {
          const storage = proofingStorage();
          writeAutoCorrectEntries(storage, [...readAutoCorrectEntries(storage), entry]);
        }
      }
      pushUndo({ kind: "edit", finding, edits, ...(entry ? { autoCorrect: entry } : {}) });
      setChanged(true);
      void scan({ ...finding.cursor, offset: resumeOffset });
    } finally {
      setBusy(false);
    }
  };

  const deleteRepeated = async () => {
    if (!finding || busy || !repeated) return;
    const writable = hostRef.current.writable(finding.target);
    if (writable !== true) {
      hostRef.current.notify(writable);
      return;
    }
    const text = readSpellingTarget(hostRef.current.sheet(), finding.target);
    if (text === null) {
      void scan(finding.cursor);
      return;
    }
    const range = repeatedWordDeletion(text, finding);
    setBusy(true);
    try {
      const edits = await commitInGroups((sheet, inGroup) => {
        if (!inGroup(finding.target)) return null;
        const written = writeSpellingTarget(
          sheet,
          finding.target,
          range.start,
          range.end,
          "",
          text.slice(range.start, range.end),
        );
        return written ? { sheet: written.sheet, edits: [written.edit] } : null;
      });
      if (edits.length === 0) {
        void scan(finding.cursor);
        return;
      }
      pushUndo({ kind: "edit", finding, edits });
      setChanged(true);
      void scan({ ...finding.cursor, offset: range.start });
    } finally {
      setBusy(false);
    }
  };

  const undoLast = async () => {
    if (busy) return;
    const record = live.current.undo.pop();
    setUndoDepth(live.current.undo.length);
    if (!record) return;
    if (record.kind === "ignore-all") stopIgnoringWordForSession(record.word, scope);
    if (record.kind === "add") {
      setCustomWords(readCustomDictionary(proofingStorage()).filter((word) => word !== record.word));
    }
    if (record.kind === "edit") {
      setBusy(true);
      try {
        const restored = await commitInGroups((sheet, inGroup) => {
          let next = sheet;
          const edits: SpellingTextEdit[] = [];
          for (const edit of [...record.edits].reverse()) {
            if (!inGroup(edit.target)) continue;
            const reverted = replaceSpellingTargetText(next, edit.target, edit.after, edit.before);
            if (!reverted) continue;
            next = reverted;
            edits.push({ ...edit, before: edit.after, after: edit.before });
          }
          return edits.length > 0 ? { sheet: next, edits } : null;
        });
        if (restored.length === 0) {
          hostRef.current.notify(t("spreadsheet.spelling.undoUnavailable"));
        }
        if (record.autoCorrect) {
          const storage = proofingStorage();
          const entry = record.autoCorrect;
          writeAutoCorrectEntries(
            storage,
            readAutoCorrectEntries(storage).filter(
              (candidate) => candidate.from !== entry.from || candidate.to !== entry.to,
            ),
          );
        }
      } finally {
        setBusy(false);
      }
    }
    void scan(record.finding.cursor);
  };

  const chooseLanguage = (tag: string) => {
    setLanguage(tag);
    setOptions({ ...live.current.options, dictionaryLanguage: tag });
    void load(tag, finding?.cursor ?? SPELLING_START);
  };

  const saveOptions = (next: SpellingOptions, words: string[], entries: AutoCorrectEntry[]) => {
    setOptions(next);
    setCustomWords(words);
    writeAutoCorrectEntries(proofingStorage(), entries);
    setOptionsOpen(false);
    if (phase.kind === "issue") void scan(phase.finding.cursor);
  };

  const close = () => {
    live.current.generation += 1;
    onCloseRef.current();
  };

  // Escape closes the dialog from wherever the focus has ended up (a pressed
  // button that became disabled leaves it on the body).
  const closeRef = useRef(close);
  useLayoutEffect(() => {
    closeRef.current = close;
  });
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (
        event.key !== "Escape" ||
        event.defaultPrevented ||
        isSpreadsheetImeKeydown(event) ||
        (document.activeElement !== null && document.activeElement !== document.body)
      ) {
        return;
      }
      event.preventDefault();
      closeRef.current();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.defaultPrevented || isSpreadsheetImeKeydown(event.nativeEvent)) return;
    if (event.key === "Escape") {
      event.preventDefault();
      if (optionsOpen) setOptionsOpen(false);
      else close();
      return;
    }
    if (event.key === "Enter" && !optionsOpen) {
      const target = event.target as HTMLElement;
      if (target.closest("button, select") !== null) return;
      event.preventDefault();
      defaultButtonRef.current?.click();
      return;
    }
    containDialogTab(event, event.currentTarget);
  };

  const onListKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (suggestions.length === 0) return;
    const step =
      event.key === "ArrowDown" ? 1 : event.key === "ArrowUp" ? -1 : 0;
    if (step === 0) return;
    event.preventDefault();
    setDraft(finding?.word ?? "");
    setSelected((current) =>
      Math.max(0, Math.min(suggestions.length - 1, (current < 0 ? -1 : current) + step)),
    );
  };

  const location = (target: SpellingTarget): string => {
    switch (target.kind) {
      case "cell":
        return t("spreadsheet.spelling.inCell", {
          ref: `${columnToLetters(target.col)}${target.row + 1}`,
        });
      case "note":
        return t("spreadsheet.spelling.inNote", {
          ref: `${columnToLetters(target.col)}${target.row + 1}`,
        });
      case "shape":
        return t("spreadsheet.spelling.inShape");
      case "header-footer":
        return target.field.toLowerCase().includes("header")
          ? t("spreadsheet.spelling.inHeader")
          : t("spreadsheet.spelling.inFooter");
    }
  };

  const button = (
    label: string,
    onClick: () => void,
    disabled: boolean,
    testId: string,
    extra: { primary?: boolean; ref?: React.Ref<HTMLButtonElement> } = {},
  ) => (
    <button
      type="button"
      ref={extra.ref}
      className={extra.primary ? dialogParts.primaryButton : dialogParts.button}
      onClick={onClick}
      disabled={disabled}
      data-testid={`spreadsheet-spelling-${testId}`}
    >
      {SPELLING_ICON[testId] ? <SpreadsheetDialogIcon name={SPELLING_ICON[testId]} /> : null}
      {label}
    </button>
  );

  const cancelLabel = changed ? t("common.close") : t("common.cancel");
  const footer = (start: React.ReactNode, end: React.ReactNode) => (
    <div className={dialogParts.footer}>
      {start ? <div className={dialogParts.footerStart}>{start}</div> : null}
      {end}
    </div>
  );

  let body: React.ReactNode;
  if (optionsOpen) {
    body = (
      <SpellingOptionsPanel
        initial={options}
        customWords={customWords}
        onCancel={() => setOptionsOpen(false)}
        onSave={saveOptions}
      />
    );
  } else if (phase.kind === "issue" && finding) {
    const noReplacement = replacement === null;
    body = (
      <>
        <div className={dialogParts.body}>
          <div className={styles.layout}>
            <div className={styles.fields}>
              <label className={dialogParts.label} htmlFor={wordId}>
                {repeated
                  ? t("spreadsheet.spelling.repeatedWord")
                  : t("spreadsheet.spelling.notInDictionary")}
              </label>
              <input
                id={wordId}
                value={draft}
                readOnly={repeated}
                spellCheck={false}
                autoComplete="off"
                onChange={(event) => setDraft(event.target.value)}
                data-testid="spreadsheet-spelling-word"
              />
              <span
                className={`${dialogParts.hint} ${styles.location}`}
                data-testid="spreadsheet-spelling-location"
              >
                {location(finding.target)}
              </span>
              <span className={dialogParts.label} id={suggestionsId}>
                {t("spreadsheet.spelling.suggestions")}
              </span>
              <div
                ref={listRef}
                role="listbox"
                tabIndex={0}
                aria-labelledby={suggestionsId}
                aria-activedescendant={
                  selected >= 0 && !edited ? `${suggestionsId}-${selected}` : undefined
                }
                className={`${dialogParts.list} ${styles.suggestions}`}
                onKeyDown={onListKeyDown}
                data-testid="spreadsheet-spelling-suggestions"
              >
                {suggestions.length > 0 ? (
                  suggestions.map((suggestion, index) => (
                    <div
                      key={suggestion}
                      id={`${suggestionsId}-${index}`}
                      role="option"
                      aria-selected={index === selected && !edited}
                      className={dialogParts.listItem}
                      onClick={() => {
                        setSelected(index);
                        setDraft(finding.word);
                      }}
                      onDoubleClick={() => void change(false, false)}
                      data-testid={`spreadsheet-spelling-suggestion-${index}`}
                    >
                      {suggestion}
                    </div>
                  ))
                ) : (
                  <span className={dialogParts.hint}>
                    {repeated ? "" : t("spreadsheet.spelling.noSuggestions")}
                  </span>
                )}
              </div>
              <label className={dialogParts.label} htmlFor={languageId}>
                {t("spreadsheet.spelling.dictionaryLanguage")}
              </label>
              <select
                id={languageId}
                value={language}
                disabled={busy}
                onChange={(event) => chooseLanguage(event.target.value)}
                data-testid="spreadsheet-spelling-language"
              >
                {SPELLING_DICTIONARIES.map((entry) => ({
                  tag: entry.tag,
                  name: languageName(locale, entry.tag),
                }))
                  .sort((left, right) => left.name.localeCompare(right.name, locale))
                  .map((entry) => (
                    <option key={entry.tag} value={entry.tag}>
                      {entry.name}
                    </option>
                  ))}
              </select>
            </div>
            <div className={styles.actions}>
              {button(t("spreadsheet.spelling.ignoreOnce"), ignoreOnce, busy, "ignore-once", {
                ref: firstActionRef,
              })}
              {button(t("spreadsheet.spelling.ignoreAll"), ignoreAll, busy || repeated, "ignore-all")}
              {button(
                t("spreadsheet.spelling.addToDictionary"),
                addToDictionary,
                busy || !dictionaryWord,
                "add",
              )}
              <span className={styles.actionGap} aria-hidden="true" />
              {repeated
                ? button(
                    t("spreadsheet.spelling.delete"),
                    () => void deleteRepeated(),
                    busy,
                    "delete",
                    { primary: true, ref: defaultButtonRef },
                  )
                : button(
                    t("spreadsheet.spelling.change"),
                    () => void change(false, false),
                    busy || noReplacement,
                    "change",
                    { primary: true, ref: defaultButtonRef },
                  )}
              {button(
                t("spreadsheet.spelling.changeAll"),
                () => void change(true, false),
                busy || repeated || noReplacement,
                "change-all",
              )}
              {button(
                t("spreadsheet.spelling.autoCorrect"),
                () => void change(false, true),
                busy || repeated || noReplacement || !replacement?.trim(),
                "autocorrect",
              )}
            </div>
          </div>
        </div>
        {footer(
          <>
            {button(t("spreadsheet.spelling.options"), () => setOptionsOpen(true), busy, "options")}
            {button(
              t("spreadsheet.spelling.undoLast"),
              () => void undoLast(),
              busy || undoDepth === 0,
              "undo-last",
            )}
          </>,
          button(cancelLabel, close, false, "cancel"),
        )}
      </>
    );
  } else if (phase.kind === "wrap") {
    body = (
      <>
        <div className={dialogParts.body}>
          <p className={`${dialogParts.message} ${styles.status}`} role="status">
            {t("spreadsheet.spelling.continueAtBeginning")}
          </p>
        </div>
        {footer(
          button(t("spreadsheet.spelling.undoLast"), () => void undoLast(), busy || undoDepth === 0, "undo-last"),
          <>
            {button(t("spreadsheet.spelling.no"), close, false, "wrap-no")}
            {button(t("spreadsheet.spelling.yes"), () => void scan(phase.cursor), false, "wrap-yes", {
              primary: true,
              ref: defaultButtonRef,
            })}
          </>,
        )}
      </>
    );
  } else if (phase.kind === "complete") {
    body = (
      <>
        <div className={dialogParts.body}>
          <p
            className={`${dialogParts.message} ${styles.status}`}
            role="status"
            data-testid="spreadsheet-spelling-complete"
          >
            {t("spreadsheet.spelling.complete")}
          </p>
        </div>
        {footer(
          null,
          button(t("spreadsheet.fmtOk"), close, false, "ok", { primary: true, ref: defaultButtonRef }),
        )}
      </>
    );
  } else if (phase.kind === "failed") {
    body = (
      <>
        <div className={dialogParts.body}>
          <p className={`${dialogParts.error} ${styles.status}`} role="alert">
            {t("spreadsheet.spelling.loadFailed", { language: languageName(locale, language) })}
          </p>
        </div>
        {footer(
          null,
          <>
            {button(cancelLabel, close, false, "cancel")}
            {button(
              t("spreadsheet.spelling.retry"),
              () => void load(language, SPELLING_START),
              false,
              "retry",
              { primary: true, ref: defaultButtonRef },
            )}
          </>,
        )}
      </>
    );
  } else {
    body = (
      <>
        <div className={dialogParts.body}>
          <p className={`${dialogParts.message} ${styles.status}`} role="status">
            {t("spreadsheet.spelling.checking")}
          </p>
        </div>
        {footer(null, button(cancelLabel, close, false, "cancel"))}
      </>
    );
  }

  // The completion message and the wrap question take the focus on arrival.
  useEffect(() => {
    if (phase.kind === "complete" || phase.kind === "wrap" || phase.kind === "failed") {
      defaultButtonRef.current?.focus({ preventScroll: true });
    }
  }, [phase.kind]);

  return (
    <div
      ref={rootRef}
      className={`${spreadsheetDialogClassName()} ${styles.dialog}`}
      role="dialog"
      aria-modal="true"
      aria-labelledby={titleId}
      tabIndex={-1}
      data-control-id="XLSCTL.REVIEW.PROOF.SPELL"
      data-testid="spreadsheet-spelling-dialog"
      data-phase={optionsOpen ? "options" : phase.kind}
      onKeyDown={onKeyDown}
    >
      <div className={dialogParts.header}>
        <strong id={titleId} className={dialogParts.title}>
          {optionsOpen
            ? t("spreadsheet.spelling.optionsTitle")
            : t("spreadsheet.spelling.title", { language: languageName(locale, language) })}
        </strong>
        <button
          type="button"
          className={dialogParts.close}
          onClick={close}
          aria-label={t("common.close")}
          data-testid="spreadsheet-spelling-close"
        >
          <X size={14} aria-hidden="true" />
        </button>
      </div>
      {body}
    </div>
  );
}

/**
 * The Proofing options Excel's Options button opens ("When correcting
 * spelling in Microsoft Office programs"), the custom dictionary's word list
 * and the AutoCorrect replace-as-you-type entries, edited as one draft and
 * kept on OK.
 */
function SpellingOptionsPanel({
  initial,
  customWords,
  onSave,
  onCancel,
}: {
  initial: SpellingOptions;
  customWords: readonly string[];
  onSave: (options: SpellingOptions, words: string[], entries: AutoCorrectEntry[]) => void;
  onCancel: () => void;
}): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const [draft, setDraft] = useState(initial);
  const [words, setWords] = useState<string[]>(() => [...customWords]);
  const [entries, setEntries] = useState<AutoCorrectEntry[]>(() =>
    readAutoCorrectEntries(proofingStorage()),
  );
  const [newWord, setNewWord] = useState("");
  const [newFrom, setNewFrom] = useState("");
  const [newTo, setNewTo] = useState("");
  const frenchAvailable = SPELLING_DICTIONARIES.some((entry) =>
    entry.tag.toLowerCase().startsWith("fr"),
  );

  const check = (
    key: Exclude<keyof SpellingOptions, "dictionaryLanguage">,
    label: string,
    disabled = false,
  ) => (
    <label className={styles.check}>
      <input
        type="checkbox"
        checked={draft[key]}
        disabled={disabled}
        onChange={(event) => setDraft({ ...draft, [key]: event.target.checked })}
        data-testid={`spreadsheet-spelling-option-${key}`}
      />
      {label}
    </label>
  );

  const addWord = () => {
    const word = normalizeCustomDictionaryWord(newWord);
    if (!word) return;
    if (!words.includes(word)) setWords([...words, word]);
    setNewWord("");
  };
  const addEntry = () => {
    const entry = normalizeAutoCorrectEntry({ from: newFrom, to: newTo });
    if (!entry) return;
    setEntries([
      ...entries.filter((candidate) => candidate.from.toLowerCase() !== entry.from.toLowerCase()),
      entry,
    ]);
    setNewFrom("");
    setNewTo("");
  };

  return (
    <>
      <div className={`${dialogParts.body} ${styles.optionsBody}`}>
        <div className={dialogParts.group}>
          <span className={dialogParts.sectionTitle}>
            {t("spreadsheet.spelling.whenCorrecting")}
          </span>
          <div className={styles.checks}>
            {check("ignoreUppercase", t("spreadsheet.spelling.ignoreUppercase"))}
            {check("ignoreWordsWithNumbers", t("spreadsheet.spelling.ignoreNumbers"))}
            {check("ignoreInternetAndFileAddresses", t("spreadsheet.spelling.ignoreAddresses"))}
            {check("flagRepeatedWords", t("spreadsheet.spelling.flagRepeated"))}
            {check(
              "enforceAccentedUppercaseFrench",
              t("spreadsheet.spelling.frenchAccents"),
              !frenchAvailable,
            )}
            {frenchAvailable ? null : (
              <span className={dialogParts.hint}>{t("spreadsheet.spelling.frenchUnavailable")}</span>
            )}
            {check("suggestFromMainDictionaryOnly", t("spreadsheet.spelling.suggestMainOnly"))}
          </div>
        </div>
        <div className={dialogParts.group}>
          <span className={dialogParts.sectionTitle}>
            {t("spreadsheet.spelling.customDictionary")}
          </span>
          <span className={dialogParts.hint}>{t("spreadsheet.spelling.customDictionaryHint")}</span>
          <div className={styles.entryEditor}>
            <div className={styles.entryInputs}>
              <input
                value={newWord}
                aria-label={t("spreadsheet.spelling.wordToAdd")}
                onChange={(event) => setNewWord(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    addWord();
                  }
                }}
                data-testid="spreadsheet-spelling-new-word"
              />
              <button
                type="button"
                className={dialogParts.compactButton}
                disabled={!normalizeCustomDictionaryWord(newWord)}
                onClick={addWord}
                data-testid="spreadsheet-spelling-add-word"
              >
                {t("spreadsheet.spelling.add")}
              </button>
            </div>
            <div
              className={`${dialogParts.list} ${styles.entryList}`}
              data-testid="spreadsheet-spelling-word-list"
            >
              {words.length === 0 ? (
                <span className={dialogParts.hint}>{t("spreadsheet.spelling.emptyList")}</span>
              ) : (
                words.map((word) => (
                  <span key={word} className={styles.entryRow}>
                    <span className={styles.entryText}>{word}</span>
                    <button
                      type="button"
                      className={dialogParts.compactButton}
                      onClick={() => setWords(words.filter((candidate) => candidate !== word))}
                    >
                      {t("spreadsheet.spelling.removeEntry")}
                    </button>
                  </span>
                ))
              )}
            </div>
          </div>
        </div>
        <div className={dialogParts.group}>
          <span className={dialogParts.sectionTitle}>
            {t("spreadsheet.spelling.autoCorrectEntries")}
          </span>
          <span className={dialogParts.hint}>{t("spreadsheet.spelling.autoCorrectHint")}</span>
          <div className={styles.entryEditor}>
            <div className={styles.entryInputs}>
              <input
                value={newFrom}
                aria-label={t("spreadsheet.spelling.replace")}
                placeholder={t("spreadsheet.spelling.replace")}
                onChange={(event) => setNewFrom(event.target.value)}
              />
              <input
                value={newTo}
                aria-label={t("spreadsheet.spelling.with")}
                placeholder={t("spreadsheet.spelling.with")}
                onChange={(event) => setNewTo(event.target.value)}
              />
              <button
                type="button"
                className={dialogParts.compactButton}
                disabled={!normalizeAutoCorrectEntry({ from: newFrom, to: newTo })}
                onClick={addEntry}
              >
                {t("spreadsheet.spelling.add")}
              </button>
            </div>
            <div
              className={`${dialogParts.list} ${styles.entryList}`}
              data-testid="spreadsheet-spelling-autocorrect-list"
            >
              {entries.length === 0 ? (
                <span className={dialogParts.hint}>{t("spreadsheet.spelling.emptyList")}</span>
              ) : (
                entries.map((entry) => (
                  <span key={entry.from} className={styles.entryRow}>
                    <span className={styles.entryText}>
                      {entry.from} → {entry.to}
                    </span>
                    <button
                      type="button"
                      className={dialogParts.compactButton}
                      onClick={() =>
                        setEntries(entries.filter((candidate) => candidate !== entry))
                      }
                    >
                      {t("spreadsheet.spelling.removeEntry")}
                    </button>
                  </span>
                ))
              )}
            </div>
          </div>
        </div>
      </div>
      <div className={dialogParts.footer}>
        <button
          type="button"
          className={dialogParts.button}
          onClick={onCancel}
          data-testid="spreadsheet-spelling-options-cancel"
        >
          {t("common.cancel")}
        </button>
        <button
          type="button"
          className={dialogParts.primaryButton}
          onClick={() => onSave(draft, words, entries)}
          data-testid="spreadsheet-spelling-options-ok"
        >
          {t("spreadsheet.fmtOk")}
        </button>
      </div>
    </>
  );
}
