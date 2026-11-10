"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  buildFindSnippet,
  OfficeFindPane,
  OfficeFindPaneToggle,
  officeFindPaneClassNames,
  type OfficeFindPaneResult,
} from "@mrnote/core/web/components/office-ribbon/OfficeFindPane.tsx";
import {
  OfficeTaskPane,
  type OfficeTaskPaneTab,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  PRESENTATION_FIND_SCOPE,
  presentationFindMatchLine,
  type PresentationFindMatch,
} from "../presentation-find-replace";

import { PRESENTATION_TASK_PANE_TITLE_IDS } from "./presentationTaskPaneDock";

/**
 * The deck's Find pane: PowerPoint's Find / Replace (Home > Editing) as the
 * results pane Word's Navigation search is, docked at the stage's right edge
 * with the deck's other task panes. It is the shared OfficeFindPane (the
 * document editor's look) fed by lib/presentation-find-replace: every match
 * in deck order as a row "Slide N · object" with the words around it, the
 * match the editor is on marked Current. What Find Next, a row, Replace and
 * Replace All do stays with the editor, which re-enumerates the live deck
 * before it moves or writes anything.
 */

/** A request to put the keyboard in one of the pane's two fields. */
export interface PresentationFindFocusRequest {
  readonly target: "find" | "replace";
  readonly id: number;
}

export interface PresentationFindPaneProps {
  readonly focusRequest: PresentationFindFocusRequest;
  readonly query: string;
  readonly replacement: string;
  readonly matchCase: boolean;
  readonly wholeWord: boolean;
  /** The live deck's matches, in deck order. */
  readonly matches: readonly PresentationFindMatch[];
  readonly activeMatchKey: string | null;
  readonly slides: readonly Record<string, unknown>[];
  /** The object a match is in, as the Selection Pane names it. */
  readonly objectLabel: (match: PresentationFindMatch) => string;
  readonly status: string;
  /** Whether the current match can be replaced. */
  readonly hasCurrentMatch: boolean;
  readonly readOnly?: boolean;
  readonly onQueryChange: (value: string) => void;
  readonly onReplacementChange: (value: string) => void;
  readonly onMatchCaseChange: (value: boolean) => void;
  readonly onWholeWordChange: (value: boolean) => void;
  readonly onFindNext: () => void;
  readonly onFindPrevious: () => void;
  readonly onActivateMatch: (match: PresentationFindMatch) => boolean;
  readonly onReplaceCurrent: () => void;
  readonly onReplaceAll: () => void;
  readonly onClose: () => void;
  readonly dock?: {
    readonly tabs?: readonly OfficeTaskPaneTab[];
    readonly tabsLabel?: string;
    readonly hidden?: boolean;
    readonly className?: string;
  };
}

/** The pane's result rows: the preview of each hit and where it is. */
export function presentationFindResults(
  slides: readonly Record<string, unknown>[],
  matches: readonly PresentationFindMatch[],
  location: (match: PresentationFindMatch) => string,
): OfficeFindPaneResult[] {
  return matches.map((match) => {
    const line = presentationFindMatchLine(slides, match) ?? "";
    return {
      key: match.key,
      snippet: buildFindSnippet(line, match.start, match.end),
      label: location(match),
    };
  });
}

export default function PresentationFindPane({
  focusRequest,
  query,
  replacement,
  matchCase,
  wholeWord,
  matches,
  activeMatchKey,
  slides,
  objectLabel,
  status,
  hasCurrentMatch,
  readOnly = false,
  onQueryChange,
  onReplacementChange,
  onMatchCaseChange,
  onWholeWordChange,
  onFindNext,
  onFindPrevious,
  onActivateMatch,
  onReplaceCurrent,
  onReplaceAll,
  onClose,
  dock,
}: PresentationFindPaneProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const findInputRef = useRef<HTMLInputElement | null>(null);
  const replaceInputRef = useRef<HTMLInputElement | null>(null);
  // PowerPoint's Find shows Replace under it; Word's pane folds it away. The
  // section starts open (Replace is one click from Find, as in the dialog it
  // replaces) and folds on its toggle.
  const [replaceExpanded, setReplaceExpanded] = useState(true);
  const focusTarget = focusRequest.target;
  const focusId = focusRequest.id;
  useEffect(() => {
    if (focusTarget === "replace" && !replaceExpanded) {
      setReplaceExpanded(true);
      return;
    }
    const input =
      focusTarget === "replace" ? replaceInputRef.current : findInputRef.current;
    if (input?.isConnected) input.focus({ preventScroll: true });
    // Only a new request moves the keyboard, never a fold of the section.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focusId, focusTarget]);
  useEffect(() => {
    if (focusTarget === "replace" && replaceExpanded) {
      replaceInputRef.current?.focus({ preventScroll: true });
    }
    // Opening the section for a Replace request hands it the keyboard.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [replaceExpanded]);

  const results = useMemo(
    () =>
      presentationFindResults(slides, matches, (match) =>
        t("presentation.findReplace.resultLocation", {
          slide: match.slideIndex + 1,
          object: objectLabel(match),
        }),
      ),
    [matches, objectLabel, slides, t],
  );
  const currentIndex = matches.findIndex((match) => match.key === activeMatchKey);
  const hasQuery = query.length > 0;
  const matchCount = matches.length;
  const countText = hasQuery
    ? `${currentIndex >= 0 ? currentIndex + 1 : 0}/${matchCount}`
    : t("findReplace.ready");
  const noNavigation = !hasQuery || matchCount === 0;

  return (
    <OfficeTaskPane
      title={t("presentation.findReplace.label")}
      closeLabel={t("presentation.findReplace.close")}
      onClose={onClose}
      testIdPrefix="presentation-find"
      testId="presentation-find-replace-dialog"
      closeTestId="presentation-find-close"
      controlId="PPTCTL.PANE.FIND"
      titleId={PRESENTATION_TASK_PANE_TITLE_IDS.find}
      tabs={dock?.tabs}
      tabsLabel={dock?.tabsLabel}
      hidden={dock?.hidden}
      dockRight={!dock?.className}
      className={dock?.className}
      dataAttributes={{ "data-search-scope": PRESENTATION_FIND_SCOPE }}
    >
      <OfficeFindPane
        testIdPrefix="presentation-find"
        docked
        query={query}
        onQueryChange={onQueryChange}
        queryPlaceholder={t("findReplace.findPlaceholder")}
        queryLabel={t("presentation.findReplace.findLabel")}
        queryInputRef={findInputRef}
        queryData={{ "data-testid": "presentation-find-input" }}
        onQueryKeyDown={(event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          if (event.shiftKey) onFindPrevious();
          else onFindNext();
        }}
        countText={countText}
        previous={{
          label: t("findReplace.previous"),
          disabled: noNavigation,
          onClick: onFindPrevious,
        }}
        next={{
          label: t("findReplace.next"),
          disabled: noNavigation,
          onClick: onFindNext,
        }}
        optionsLabel={t("presentation.findReplace.searchOptions")}
        options={
          <>
            <OfficeFindPaneToggle
              active={matchCase}
              label={t("presentation.findReplace.matchCase")}
              icon="edit.matchCase"
              testId="presentation-find-match-case"
              controlId="PPTCTL.DIALOG.FIND_REPLACE.MATCH_CASE"
              onClick={() => onMatchCaseChange(!matchCase)}
            />
            <OfficeFindPaneToggle
              active={wholeWord}
              label={t("findReplace.wholeWord")}
              icon="edit.matchWholeWord"
              testId="presentation-find-whole-word"
              controlId="PPTCTL.DIALOG.FIND_REPLACE.WHOLE_WORD"
              onClick={() => onWholeWordChange(!wholeWord)}
            />
          </>
        }
        results={{
          label: t("presentation.findReplace.resultsLabel"),
          heading: t("presentation.findReplace.results"),
          count: hasQuery ? matchCount : 0,
          empty: !hasQuery
            ? t("presentation.findReplace.enterQuery")
            : matchCount === 0
              ? t("presentation.findReplace.noMatches")
              : null,
          items: results,
          currentIndex,
          currentLabel: t("presentation.findReplace.current"),
          onActivate: (index) => {
            const match = matches[index];
            return match ? onActivateMatch(match) : false;
          },
        }}
        replace={{
          label: t("findReplace.replace"),
          expanded: replaceExpanded,
          disabled: readOnly,
          onToggle: () => setReplaceExpanded((expanded) => !expanded),
          children: (
            <>
              <input
                ref={replaceInputRef}
                type="text"
                className={officeFindPaneClassNames.replaceInput}
                data-testid="presentation-replace-input"
                aria-label={t("presentation.findReplace.replaceLabel")}
                placeholder={t("findReplace.replacePlaceholder")}
                value={replacement}
                disabled={readOnly}
                onChange={(event) => onReplacementChange(event.target.value)}
              />
              <div className={officeFindPaneClassNames.replaceActions}>
                <button
                  type="button"
                  className={officeFindPaneClassNames.action}
                  data-testid="presentation-replace-current"
                  disabled={readOnly || !hasCurrentMatch}
                  onClick={onReplaceCurrent}
                >
                  <OfficeSemanticIcon semanticId="edit.replace" size={15} />
                  <span>{t("findReplace.replace")}</span>
                </button>
                <button
                  type="button"
                  className={officeFindPaneClassNames.action}
                  data-testid="presentation-replace-all"
                  disabled={readOnly || noNavigation}
                  onClick={onReplaceAll}
                >
                  <OfficeSemanticIcon semanticId="edit.replace" size={15} />
                  <span>{t("presentation.findReplace.replaceAll")}</span>
                </button>
              </div>
            </>
          ),
        }}
      >
        {status ? (
          <p
            className={officeFindPaneClassNames.status}
            role="status"
            data-testid="presentation-find-status"
          >
            {status}
          </p>
        ) : null}
      </OfficeFindPane>
    </OfficeTaskPane>
  );
}
