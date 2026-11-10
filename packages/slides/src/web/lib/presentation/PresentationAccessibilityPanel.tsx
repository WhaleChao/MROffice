"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useMemo, useState } from "react";
import {
  ChevronDown,
  ChevronUp,
  ListTree,
  LocateFixed,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  OfficeTaskPane,
  officeTaskPaneParts,
  type OfficeTaskPaneTab,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";

import {
  collectPresentationAccessibilityIssues,
  presentationAccessibilityIssueObjectName,
  presentationAccessibilityRemediation,
  presentationAccessibilityIssueSignature,
  type PresentationAccessibilityIssue,
  type PresentationAccessibilitySlideLike,
} from "../presentation-a11y";

interface PresentationAccessibilityPanelProps {
  readonly slides: readonly PresentationAccessibilitySlideLike[];
  readonly onActivate: (issue: PresentationAccessibilityIssue) => boolean;
  /** Puts the caret where the finding's slide title is typed. */
  readonly onEditSlideTitle: (issue: PresentationAccessibilityIssue) => boolean;
  readonly onFixImageAlt: (
    issue: PresentationAccessibilityIssue,
    trigger: HTMLElement | null,
  ) => boolean;
  readonly onFixTableHeader: (issue: PresentationAccessibilityIssue) => boolean;
  readonly onOpenSelectionPane: (
    issue: PresentationAccessibilityIssue,
  ) => boolean;
  readonly onClose: () => void;
  /** The deck's task pane dock (lib/presentation/presentationTaskPaneDock):
   * the open panes' tabs, whether another pane is in front of this one, and
   * the dock's own pane class. Without it the pane docks at the right of
   * its container on its own. */
  readonly dock?: {
    readonly tabs?: readonly OfficeTaskPaneTab[];
    readonly tabsLabel?: string;
    readonly hidden?: boolean;
    readonly className?: string;
  };
}

export default function PresentationAccessibilityPanel({
  slides,
  onActivate,
  onEditSlideTitle,
  onFixImageAlt,
  onFixTableHeader,
  onOpenSelectionPane,
  onClose,
  dock,
}: PresentationAccessibilityPanelProps): React.JSX.Element {
  const t = useTranslations("console-notebooks");
  const issues = useMemo(
    () => collectPresentationAccessibilityIssues(slides),
    [slides],
  );
  const [currentIndex, setCurrentIndex] = useState(0);
  const boundedIndex = issues.length
    ? Math.min(currentIndex, issues.length - 1)
    : 0;
  const currentIssue = issues[boundedIndex];
  const currentRemediation = currentIssue
    ? presentationAccessibilityRemediation(currentIssue)
    : null;

  const severityLabel = useCallback(
    (issue: PresentationAccessibilityIssue) =>
      issue.severity === "error"
        ? t("code.problems.severity.error")
        : issue.severity === "warning"
          ? t("code.problems.severity.warning")
          : t("code.problems.severity.info"),
    [t],
  );

  // The rule and its advice are the product's words and follow the interface
  // language; the object's name is the deck's own and is shown as written.
  const ruleLabel = useCallback(
    (issue: PresentationAccessibilityIssue) =>
      t(
        `presentation.ui.accessibility.rule.${issue.reason}` as "presentation.ui.accessibility.rule.missingAltText",
      ),
    [t],
  );

  const optionLabel = useCallback(
    (issue: PresentationAccessibilityIssue) => {
      const values = {
        severity: severityLabel(issue),
        number: issue.slideIndex + 1,
        rule: ruleLabel(issue),
      };
      const object = presentationAccessibilityIssueObjectName(issue);
      return object
        ? t("presentation.ui.accessibility.issueOptionObject", {
            ...values,
            object,
          })
        : t("presentation.ui.accessibility.issueOption", values);
    },
    [ruleLabel, severityLabel, t],
  );

  const activate = useCallback(
    (issue: PresentationAccessibilityIssue | undefined, index: number) => {
      if (!issue || !onActivate(issue)) return false;
      setCurrentIndex(index);
      return true;
    },
    [onActivate],
  );

  const move = useCallback(
    (direction: 1 | -1) => {
      if (!issues.length) return;
      const nextIndex =
        (boundedIndex + direction + issues.length) % issues.length;
      activate(issues[nextIndex], nextIndex);
    },
    [activate, boundedIndex, issues],
  );

  // PowerPoint's Accessibility pane (Review > Check Accessibility) is a task
  // pane at the right of the window ("Improve accessibility with the
  // Accessibility Checker", Microsoft Support); it is the shared Office task
  // pane here, docked with the deck's other panes.
  return (
    <OfficeTaskPane
      title={t("toolbar.accessibilityCheck")}
      closeLabel={t("officeTaskPane.close")}
      onClose={onClose}
      testIdPrefix="presentation-accessibility"
      testId="presentation-accessibility-panel"
      closeTestId="presentation-accessibility-close"
      titleId="presentation-accessibility-panel-title"
      tabs={dock?.tabs}
      tabsLabel={dock?.tabsLabel}
      hidden={dock?.hidden}
      dockRight={!dock?.className}
      className={dock?.className}
    >
      <div className={officeTaskPaneParts.section}>
      <label className={officeTaskPaneParts.field}>
        {t("code.problems.title")}
      <select
        aria-label={t("toolbar.accessibilityCheck")}
        data-testid="presentation-accessibility-issues"
        value={currentIssue?.id ?? ""}
        disabled={!currentIssue}
        autoFocus
        onChange={(event) => {
          const matches = issues
            .map((issue, index) => ({ issue, index }))
            .filter(({ issue }) => issue.id === event.target.value);
          if (matches.length !== 1) return;
          activate(matches[0].issue, matches[0].index);
        }}
        onKeyDown={(event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          activate(currentIssue, boundedIndex);
        }}
      >
        {currentIssue ? (
          issues.map((issue) => (
            <option
              key={presentationAccessibilityIssueSignature(issue)}
              value={issue.id}
              data-kind={issue.kind}
              data-severity={issue.severity}
              data-reason={issue.reason}
            >
              {optionLabel(issue)}
            </option>
          ))
        ) : (
          <option value="">{t("code.problems.none")}</option>
        )}
      </select>
      </label>

      <div className={officeTaskPaneParts.row}>
        <span
          className={officeTaskPaneParts.hint}
          data-testid="presentation-accessibility-count"
          aria-live="polite"
        >
          {issues.length
            ? t("code.problems.count", { count: issues.length })
            : t("code.problems.none")}
        </span>
        <button
          type="button"
          className={officeTaskPaneParts.close}
          aria-label={t("findReplace.previous")}
          title={t("findReplace.previous")}
          data-testid="presentation-accessibility-previous"
          disabled={!currentIssue}
          onClick={() => move(-1)}
        >
          <ChevronUp size={16} aria-hidden="true" />
        </button>
        <button
          type="button"
          className={officeTaskPaneParts.close}
          aria-label={t("findReplace.next")}
          title={t("findReplace.next")}
          data-testid="presentation-accessibility-next"
          disabled={!currentIssue}
          onClick={() => move(1)}
        >
          <ChevronDown size={16} aria-hidden="true" />
        </button>
        <button
          type="button"
          className={officeTaskPaneParts.button}
          data-testid="presentation-accessibility-apply"
          disabled={!currentIssue}
          onClick={() => activate(currentIssue, boundedIndex)}
        >
          <LocateFixed size={15} aria-hidden="true" />
          <span>{t("goTo.go")}</span>
        </button>
        {currentRemediation?.kind === "locate" ? (
          <button
            type="button"
            className={officeTaskPaneParts.button}
            data-testid="presentation-accessibility-edit-slide-title"
            onClick={() => {
              if (currentIssue) onEditSlideTitle(currentIssue);
            }}
          >
            <LocateFixed size={15} aria-hidden="true" />
            {t("presentation.ui.accessibility.slideTitle")}
          </button>
        ) : currentRemediation?.kind === "fix" ? (
          <button
            type="button"
            className={officeTaskPaneParts.button}
            data-testid="presentation-accessibility-fix-table-header"
            onClick={() => {
              if (currentIssue) onFixTableHeader(currentIssue);
            }}
          >
            {t("presentation.ui.accessibility.tableHeader")}
          </button>
        ) : currentRemediation?.kind === "edit" ? (
          <button
            type="button"
            className={officeTaskPaneParts.button}
            data-testid="presentation-accessibility-fix-image-alt"
            onClick={(event) => {
              if (currentIssue) {
                onFixImageAlt(currentIssue, event.currentTarget);
              }
            }}
          >
            {t("toolbar.imageAlt")}
          </button>
        ) : currentRemediation?.kind === "manage" ? (
          <button
            type="button"
            className={officeTaskPaneParts.button}
            data-testid="presentation-accessibility-open-selection-pane"
            onClick={() => {
              if (currentIssue) onOpenSelectionPane(currentIssue);
            }}
          >
            <ListTree size={15} aria-hidden="true" />
            {t("toolbar.selectionPane")}
          </button>
        ) : null}
      </div>

      </div>
      <div className={officeTaskPaneParts.section}>
        <p
          className={officeTaskPaneParts.hint}
          data-testid="presentation-accessibility-recommendation"
        >
          {currentIssue
            ? t(
                `presentation.ui.accessibility.advice.${currentIssue.reason}` as "presentation.ui.accessibility.advice.missingAltText",
              )
            : t("code.problems.none")}
        </p>
      </div>
    </OfficeTaskPane>
  );
}
