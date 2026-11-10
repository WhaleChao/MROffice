"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import type { Editor } from "@tiptap/react";
import {
  ChevronDown,
  ChevronUp,
  LocateFixed,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  OfficeTaskPane,
  officeTaskPaneParts,
} from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import dockStyles from "@mrnote/core/web/components/console/editor/native-document/DocumentShapeFormatPane.module.css";

import {
  collectDocumentAccessibilityIssues,
  documentAccessibilityIssueSignature,
  selectDocumentAccessibilityIssue,
  type DocumentAccessibilityIssue,
} from "./native-document/documentAccessibility";

interface DocumentAccessibilityPanelProps {
  editor: Editor;
  onFixImageAlt: (
    issue: DocumentAccessibilityIssue,
    expectedDoc: ProseMirrorNode,
  ) => boolean;
  onClose: () => void;
}

type AccessibilitySnapshot = {
  doc: ProseMirrorNode;
  issues: DocumentAccessibilityIssue[];
  canEdit: boolean;
};

type RuntimeEditor = Editor & {
  on?: (event: string, callback: () => void) => unknown;
  off?: (event: string, callback: () => void) => unknown;
};

function issueListSignature(issues: readonly DocumentAccessibilityIssue[]) {
  return issues.map(documentAccessibilityIssueSignature).join("\n");
}

export default function DocumentAccessibilityPanel({
  editor,
  onFixImageAlt,
  onClose,
}: DocumentAccessibilityPanelProps) {
  const t = useTranslations("console-notebooks");
  const [snapshot, setSnapshot] = useState<AccessibilitySnapshot>(() => {
    const doc = editor.state.doc;
    return {
      doc,
      issues: collectDocumentAccessibilityIssues(doc),
      canEdit: !editor.isDestroyed && editor.isEditable,
    };
  });
  const [currentIndex, setCurrentIndex] = useState(0);
  // The current issue is the finding on its target, not a place in the
  // list: fixing a heading above an outer table's issue moved the list's
  // current entry to the nested table after it, and Go To went there
  // (WUA 181). Its target is mapped through every edit and found again.
  const currentTargetRef = useRef<{ kind: string; nodeType: string; nodePos: number } | null>(null);

  useEffect(() => {
    const runtime = editor as RuntimeEditor;
    const refresh = (props?: { transaction?: { mapping?: { map: (pos: number, assoc?: number) => number } } }) => {
      const target = currentTargetRef.current;
      const mapping = props?.transaction?.mapping;
      if (target && mapping) target.nodePos = mapping.map(target.nodePos, 1);
      const doc = editor.state.doc;
      const issues = collectDocumentAccessibilityIssues(doc);
      const canEdit = !editor.isDestroyed && editor.isEditable;
      setSnapshot((previous) => {
        const sameIssues =
          issueListSignature(previous.issues) === issueListSignature(issues);
        if (
          previous.doc === doc &&
          sameIssues &&
          previous.canEdit === canEdit
        ) {
          return previous;
        }
        return {
          doc,
          issues: sameIssues ? previous.issues : issues,
          canEdit,
        };
      });
    };
    refresh();
    runtime.on?.("update", refresh);
    return () => {
      runtime.off?.("update", refresh);
    };
  }, [editor]);

  const issues = snapshot.issues;
  const target = currentTargetRef.current;
  const targetIndex = target
    ? issues.findIndex(
        (issue) =>
          issue.kind === target.kind &&
          issue.nodeType === target.nodeType &&
          issue.nodePos === target.nodePos,
      )
    : -1;
  const boundedIndex =
    targetIndex >= 0
      ? targetIndex
      : issues.length
        ? Math.min(currentIndex, issues.length - 1)
        : 0;
  const currentIssue = issues[boundedIndex];

  const issueLabel = useCallback(
    (issue: DocumentAccessibilityIssue, index: number) => {
      const severity =
        issue.severity === "error"
          ? t("code.problems.severity.error")
          : t("code.problems.severity.warning");
      const kind =
        issue.kind === "imageMissingAlt"
          ? t("toolbar.imageAlt")
          : issue.kind === "headingLevelSkipped"
            ? t("goTo.kind.heading")
            : issue.kind === "tableStructure"
              ? t("goTo.kind.table")
              : t("toolbar.link");
      const label = issue.label ? ` — ${issue.label}` : "";
      return `${severity}: ${kind} ${index + 1}${label}`;
    },
    [t],
  );

  const activate = useCallback(
    (issue: DocumentAccessibilityIssue | undefined, index: number) => {
      if (!issue) return false;
      const selected = selectDocumentAccessibilityIssue(
        editor,
        issue,
        snapshot.doc,
      );
      if (selected) {
        currentTargetRef.current = { kind: issue.kind, nodeType: issue.nodeType, nodePos: issue.nodePos };
        setCurrentIndex(index);
      }
      return selected;
    },
    [editor, snapshot.doc],
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

  const optionLabels = useMemo(
    () => issues.map((issue, index) => issueLabel(issue, index)),
    [issueLabel, issues],
  );

  // Word's Accessibility pane (Review > Check Accessibility, "Improve
  // accessibility with the Accessibility Checker", Microsoft Support) is a
  // task pane at the right of the window: the shared Office task pane, in
  // the Docs dock Format Shape and the Selection Pane use.
  return (
    <OfficeTaskPane
      title={t("toolbar.accessibilityCheck")}
      closeLabel={t("toolbar.accessibilityCheckClose")}
      onClose={onClose}
      testIdPrefix="document-accessibility"
      testId="document-accessibility-panel"
      closeTestId="document-accessibility-close"
      dockRight={false}
      className={dockStyles.pane}
    >
      <div className={officeTaskPaneParts.section}>
      <label className={officeTaskPaneParts.field}>
        {t("code.problems.title")}
      <select
        aria-label={t("toolbar.accessibilityCheck")}
        data-testid="document-accessibility-issues"
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
          issues.map((issue, index) => (
            <option
              key={documentAccessibilityIssueSignature(issue)}
              value={issue.id}
              data-kind={issue.kind}
              data-severity={issue.severity}
              data-reasons={issue.reasons.join(",")}
            >
              {optionLabels[index]}
            </option>
          ))
        ) : (
          <option value="">{t("code.problems.none")}</option>
        )}
      </select>
      </label>
      {currentIssue ? (
        // Why this is an issue and what to do about it, as Word's checker
        
        // else).
        <div
          className={officeTaskPaneParts.hint}
          data-testid="document-accessibility-reasons"
          aria-live="polite"
        >
          <strong>{t("toolbar.accessibilityReasons")}</strong>
          <ul>
            {currentIssue.reasons.map((reason) => (
              <li key={reason}>
                {t(`toolbar.accessibilityReason.${reason}` as "toolbar.accessibilityReason.headingLevelJump")}
              </li>
            ))}
          </ul>
        </div>
      ) : null}
      <div className={officeTaskPaneParts.row}>
      <span
        className={officeTaskPaneParts.hint}
        data-testid="document-accessibility-count"
        aria-live="polite"
      >
        {issues.length
          ? t("code.problems.count", { count: issues.length })
          : t("code.problems.none")}
      </span>
      <button
        type="button"
        className={officeTaskPaneParts.iconButton}
        aria-label={t("findReplace.previous")}
        title={t("findReplace.previous")}
        data-testid="document-accessibility-previous"
        disabled={!currentIssue}
        onClick={() => move(-1)}
      >
        <ChevronUp size={16} aria-hidden="true" />
      </button>
      <button
        type="button"
        className={officeTaskPaneParts.iconButton}
        aria-label={t("findReplace.next")}
        title={t("findReplace.next")}
        data-testid="document-accessibility-next"
        disabled={!currentIssue}
        onClick={() => move(1)}
      >
        <ChevronDown size={16} aria-hidden="true" />
      </button>
      <button
        type="button"
        className={officeTaskPaneParts.button}
        data-testid="document-accessibility-apply"
        disabled={!currentIssue}
        onClick={() => activate(currentIssue, boundedIndex)}
      >
        <LocateFixed size={15} aria-hidden="true" />
        <span>{t("goTo.go")}</span>
      </button>
      {currentIssue?.kind === "imageMissingAlt" ? (
        <button
          type="button"
          className={officeTaskPaneParts.button}
          data-testid="document-accessibility-fix-image-alt"
          disabled={!snapshot.canEdit}
          onClick={() => onFixImageAlt(currentIssue, snapshot.doc)}
        >
          <span>{t("toolbar.imageAlt")}</span>
        </button>
      ) : null}
      </div>
      </div>
    </OfficeTaskPane>
  );
}
