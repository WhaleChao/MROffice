/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isPictureNode } from "@mrnote/core/web/components/console/editor/native-document/documentObjectNodes.ts";
import { TABLE_CONTINUATION_ATTR, TABLE_FRAGMENT_ID_ATTR } from "@mrnote/core/web/components/console/editor/native-document/paged/tablePagination.ts";
import type { Mark, Node as ProseMirrorNode } from "@tiptap/pm/model";
import { NodeSelection, TextSelection } from "@tiptap/pm/state";
import type { Editor } from "@tiptap/react";

export const MAX_DOCUMENT_ACCESSIBILITY_ISSUES = 500;

export type DocumentAccessibilityIssueKind =
  | "imageMissingAlt"
  | "headingLevelSkipped"
  | "tableStructure"
  | "linkTextUnclear";

export type DocumentAccessibilityIssueSeverity = "error" | "warning";

export type DocumentAccessibilityIssueReason =
  | "missingAltText"
  | "firstHeadingAboveLevelOne"
  | "headingLevelJump"
  | "missingHeaderBand"
  | "mergedCells"
  | "nestedTable"
  | "genericLinkText"
  | "exposedUrlText";

export interface DocumentAccessibilityIssue {
  /** Snapshot-local identity. The complete signature below owns authority. */
  id: string;
  kind: DocumentAccessibilityIssueKind;
  severity: DocumentAccessibilityIssueSeverity;
  reasons: readonly DocumentAccessibilityIssueReason[];
  /** PM position before the affected node. */
  nodePos: number;
  nodeType: "image" | "heading" | "table" | "text";
  nodeSize: number;
  /** Exact selection to publish when the issue is activated. */
  from: number;
  to: number;
  selection: "node" | "caret" | "text";
  /** Bounded plain text rendered only as a React child. */
  label: string;
  /** Complete rule context needed to distinguish same-position findings. */
  context: string;
}

function boundedPlainText(value: unknown, max = 120): string {
  if (typeof value !== "string") return "";
  return value.replace(/\s+/g, " ").trim().slice(0, max);
}

function nodeText(node: ProseMirrorNode): string {
  return boundedPlainText(node.textBetween(0, node.content.size, " "));
}

const GENERIC_LINK_TEXT = new Set([
  "click here",
  "see this page",
  "go here",
  "learn more",
  "read more",
  "here",
  "link",
  "点击这里",
  "单击此处",
  "查看此页面",
  "转到此处",
  "了解更多",
  "阅读更多",
  "点此",
]);

function normalizedLinkText(value: string): string {
  return boundedPlainText(value, 2_048)
    .toLowerCase()
    .replace(/[.!?;:。！？；：]+$/u, "")
    .trim();
}

function unclearLinkTextReason(
  value: string,
): "genericLinkText" | "exposedUrlText" | null {
  const normalized = normalizedLinkText(value);
  if (!normalized) return "genericLinkText";
  if (GENERIC_LINK_TEXT.has(normalized)) return "genericLinkText";
  if (/^(?:https?:\/\/|www\.)\S+$/iu.test(normalized)) {
    return "exposedUrlText";
  }
  return null;
}

interface PendingLinkText {
  readonly from: number;
  to: number;
  text: string;
  readonly mark: Mark;
  readonly parent: ProseMirrorNode | null;
}

function imageLabel(node: ProseMirrorNode): string {
  const attrs = node.attrs as Record<string, unknown>;
  return (
    boundedPlainText(attrs.title) ||
    boundedPlainText(attrs.caption) ||
    boundedPlainText(attrs.filename) ||
    boundedPlainText(attrs.name)
  );
}

function firstTextblockCaret(node: ProseMirrorNode, nodePos: number): number {
  let caret: number | null = null;
  node.descendants((child, relativePos) => {
    if (caret !== null) return false;
    if (child.isTextblock) {
      caret = nodePos + relativePos + 2;
      return false;
    }
    return true;
  });
  // A schema-valid table always owns a textblock inside a cell. Keep a closed
  // fallback for malformed test/import nodes; reauthorization rejects it if it
  // cannot produce a legal TextSelection.
  return caret ?? nodePos + 1;
}

function tableStructureReasons(
  table: ProseMirrorNode,
): DocumentAccessibilityIssueReason[] {
  const reasons: DocumentAccessibilityIssueReason[] = [];
  const firstRow = table.firstChild;
  let hasHeaderBand =
    firstRow?.type.name === "tableRow" && firstRow.childCount > 0;
  if (hasHeaderBand && firstRow) {
    firstRow.forEach((cell) => {
      if (cell.type.name !== "tableHeader") hasHeaderBand = false;
    });
  }
  if (!hasHeaderBand) reasons.push("missingHeaderBand");

  let hasMergedCells = false;
  let hasNestedTable = false;
  table.descendants((node) => {
    if (node.type.name === "table") {
      hasNestedTable = true;
      // A nested table is its own finding target. Its cells must not also make
      // the outer table look merged.
      return false;
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const attrs = node.attrs as Record<string, unknown>;
      if (Number(attrs.rowspan) > 1 || Number(attrs.colspan) > 1) {
        hasMergedCells = true;
      }
    }
    return true;
  });
  if (hasMergedCells) reasons.push("mergedCells");
  if (hasNestedTable) reasons.push("nestedTable");
  return reasons;
}

function appendIssue(
  issues: DocumentAccessibilityIssue[],
  issue: DocumentAccessibilityIssue,
): void {
  if (issues.length < MAX_DOCUMENT_ACCESSIBILITY_ISSUES) issues.push(issue);
}

/**
 * Derive the deliberately narrow, session-only Docs accessibility projection.
 * Results are bounded and emitted in document order; no document state is
 * persisted or normalized by the checker.
 */
export function collectDocumentAccessibilityIssues(
  doc: ProseMirrorNode,
): DocumentAccessibilityIssue[] {
  const issues: DocumentAccessibilityIssue[] = [];
  // A table the page layout cut over pages is one table: the reasons its
  // first page reported, by fragment id.
  const tableReasons = new Map<string, Set<DocumentAccessibilityIssueReason>>();
  let previousHeadingLevel: number | null = null;
  let pendingLink: PendingLinkText | null = null;

  const flushPendingLink = () => {
    const current = pendingLink;
    pendingLink = null;
    if (!current) return;
    const reason = unclearLinkTextReason(current.text);
    if (!reason) return;
    appendIssue(issues, {
      id: `linkTextUnclear:${current.from}`,
      kind: "linkTextUnclear",
      severity: "warning",
      reasons: [reason],
      nodePos: current.from,
      nodeType: "text",
      nodeSize: current.to - current.from,
      from: current.from,
      to: current.to,
      selection: "text",
      label: boundedPlainText(current.text),
      context: `${reason};href=${boundedPlainText(current.mark.attrs.href, 512)}`,
    });
  };

  doc.descendants((node, nodePos, parent) => {
    if (node.isText) {
      const link = node.marks.find((mark) => mark.type.name === "link");
      if (!link) {
        flushPendingLink();
        return false;
      }
      const text = node.text ?? "";
      if (
        pendingLink &&
        pendingLink.parent === parent &&
        pendingLink.to === nodePos &&
        pendingLink.mark.eq(link)
      ) {
        pendingLink.to += node.nodeSize;
        pendingLink.text += text;
      } else {
        flushPendingLink();
        pendingLink = {
          from: nodePos,
          to: nodePos + node.nodeSize,
          text,
          mark: link,
          parent,
        };
      }
      return false;
    }

    flushPendingLink();
    if (
      isPictureNode(node) &&
      node.isAtom &&
      NodeSelection.isSelectable(node)
    ) {
      const alt = boundedPlainText((node.attrs as Record<string, unknown>).alt);
      if (!alt) {
        appendIssue(issues, {
          id: `imageMissingAlt:${nodePos}`,
          kind: "imageMissingAlt",
          severity: "error",
          reasons: ["missingAltText"],
          nodePos,
          nodeType: "image",
          nodeSize: node.nodeSize,
          from: nodePos,
          to: nodePos + node.nodeSize,
          selection: "node",
          label: imageLabel(node),
          context: "alt=blank",
        });
      }
      return false;
    }

    if (node.type.name === "heading") {
      const rawLevel = Number((node.attrs as Record<string, unknown>).level);
      const level = Number.isInteger(rawLevel)
        ? Math.min(6, Math.max(1, rawLevel))
        : 1;
      const reason =
        previousHeadingLevel === null && level > 1
          ? "firstHeadingAboveLevelOne"
          : previousHeadingLevel !== null && level > previousHeadingLevel + 1
            ? "headingLevelJump"
            : null;
      if (reason) {
        appendIssue(issues, {
          id: `headingLevelSkipped:${nodePos}`,
          kind: "headingLevelSkipped",
          severity: "warning",
          reasons: [reason],
          nodePos,
          nodeType: "heading",
          nodeSize: node.nodeSize,
          from: nodePos + 1,
          to: nodePos + 1,
          selection: "caret",
          label: nodeText(node),
          context: `previous=${previousHeadingLevel ?? 0};level=${level}`,
        });
      }
      previousHeadingLevel = level;
      // The heading's own text is scanned too: a link inside a heading
      
      return true;
    }

    if (node.type.name === "table") {
      let reasons = tableStructureReasons(node);
      const fragmentId = node.attrs[TABLE_FRAGMENT_ID_ATTR];
      if (typeof fragmentId === "string" && fragmentId) {
        const reported = tableReasons.get(fragmentId) ?? new Set();
        // The rows a later page continues are the same table's: its header
        // row is on the first page, and what that page reported is not
        // reported again (WUA 072: every continued page was a table "with
        // no header row").
        if (node.attrs[TABLE_CONTINUATION_ATTR] === true) {
          reasons = reasons.filter(
            (reason) => reason !== "missingHeaderBand" && !reported.has(reason),
          );
        }
        for (const reason of reasons) reported.add(reason);
        tableReasons.set(fragmentId, reported);
      }
      if (reasons.length) {
        const caret = firstTextblockCaret(node, nodePos);
        appendIssue(issues, {
          id: `tableStructure:${nodePos}`,
          kind: "tableStructure",
          severity: "warning",
          reasons,
          nodePos,
          nodeType: "table",
          nodeSize: node.nodeSize,
          from: caret,
          to: caret,
          selection: "caret",
          label: nodeText(node),
          context: reasons.join(","),
        });
      }
    }
    return true;
  });
  flushPendingLink();

  return issues;
}

/** Complete, order-sensitive authority fingerprint for one rendered finding. */
export function documentAccessibilityIssueSignature(
  issue: DocumentAccessibilityIssue,
): string {
  return JSON.stringify([
    issue.id,
    issue.kind,
    issue.severity,
    issue.reasons,
    issue.nodePos,
    issue.nodeType,
    issue.nodeSize,
    issue.from,
    issue.to,
    issue.selection,
    issue.label,
    issue.context,
  ]);
}

/**
 * Re-derive and uniquely authenticate a finding against the exact live doc.
 * Duplicate snapshot ids fail closed even if one candidate otherwise matches.
 */
export function resolveDocumentAccessibilityIssue(
  doc: ProseMirrorNode,
  expected: DocumentAccessibilityIssue,
): DocumentAccessibilityIssue | null {
  const matches = collectDocumentAccessibilityIssues(doc).filter(
    (candidate) => candidate.id === expected.id,
  );
  if (matches.length !== 1) return null;
  return documentAccessibilityIssueSignature(matches[0]) ===
    documentAccessibilityIssueSignature(expected)
    ? matches[0]
    : null;
}

/** Navigate to one authenticated finding without editing or adding history. */
export function selectDocumentAccessibilityIssue(
  editor: Editor,
  expected: DocumentAccessibilityIssue,
  expectedDoc: ProseMirrorNode,
): boolean {
  if (editor.isDestroyed || editor.state.doc !== expectedDoc) return false;
  const issue = resolveDocumentAccessibilityIssue(expectedDoc, expected);
  if (!issue || editor.state.doc !== expectedDoc) return false;

  let selection: NodeSelection | TextSelection;
  try {
    selection =
      issue.selection === "node"
        ? NodeSelection.create(expectedDoc, issue.from)
        : TextSelection.create(expectedDoc, issue.from, issue.to);
  } catch {
    return false;
  }
  if (editor.state.doc !== expectedDoc) return false;
  const transaction = editor.state.tr
    .setSelection(selection)
    .setMeta("addToHistory", false)
    .scrollIntoView();
  editor.view.dispatch(transaction);
  editor.view.focus();
  return true;
}
