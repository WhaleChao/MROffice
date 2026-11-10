"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/react";
import { logicalTableFragmentsAt } from "@mrnote/core/web/components/console/editor/native-document/paged/tableFragmentGroups.ts";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { closeHistory } from "@tiptap/pm/history";
import type { Selection } from "@tiptap/pm/state";

import {
  collectTextMatchesFromDoc,
  replaceDocumentTextMatch,
} from "../DocumentFindReplace";
import { readDocumentTablePropertiesTarget } from "@mrnote/core/web/components/console/editor/native-document/tableProperties.ts";

export const DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH = 4096;
export const DOCUMENT_TABLE_REPLACE_TEXT_MAX_MATCHES = 10_000;

export interface DocumentTableReplaceTextTarget {
  readonly doc: ProseMirrorNode;
  readonly selection: Selection;
  readonly tablePos: number;
  readonly tableNode: ProseMirrorNode;
}

export interface DocumentTableReplaceTextOptions {
  readonly caseSensitive: boolean;
  readonly wholeWord: boolean;
}

export type DocumentTableReplaceTextResult =
  | "applied"
  | "unchanged"
  | "stale"
  | "invalid";

/**
 * Freeze the exact selected table. Moving focus into the dialog is harmless,
 * but any editor transaction, selection replacement, or table replacement
 * invalidates this authority rather than retargeting a different table.
 */
export function readDocumentTableReplaceTextTarget(
  editor: Editor,
): DocumentTableReplaceTextTarget | null {
  if (editor.isDestroyed || !editor.isEditable) return null;
  const target = readDocumentTablePropertiesTarget(editor);
  if (!target) return null;
  return {
    doc: target.doc,
    selection: target.selection,
    tablePos: target.tablePos,
    tableNode: target.tableNode,
  };
}

/** Replace literal text inside one frozen table in one undoable transaction. */
export function replaceDocumentTableText(
  editor: Editor,
  target: DocumentTableReplaceTextTarget,
  query: string,
  replacement: string,
  options: DocumentTableReplaceTextOptions,
): DocumentTableReplaceTextResult {
  if (
    query.length < 1 ||
    query.length > DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH ||
    replacement.length > DOCUMENT_TABLE_REPLACE_TEXT_MAX_LENGTH
  ) {
    return "invalid";
  }
  if (
    editor.isDestroyed ||
    !editor.isEditable ||
    editor.state.doc !== target.doc ||
    editor.state.selection !== target.selection ||
    editor.state.doc.nodeAt(target.tablePos) !== target.tableNode
  ) {
    return "stale";
  }

  // The whole logical table: every page fragment of it, not only the one
  
  // second page's Alice alone).
  const ranges = logicalTableFragmentsAt(target.doc, target.tablePos).map((fragment) => ({
    from: fragment.pos + 1,
    to: fragment.pos + fragment.node.nodeSize - 1,
  }));
  const matches = collectTextMatchesFromDoc(
    target.doc,
    query,
    options.caseSensitive,
    false,
    options.wholeWord,
  ).filter(
    (match) =>
      // Text an object shows or a control keeps is found, never rewritten
      
      !match.readOnly &&
      ranges.some((range) => match.from >= range.from && match.to <= range.to),
  );

  if (matches.length === 0) return "unchanged";
  if (matches.length > DOCUMENT_TABLE_REPLACE_TEXT_MAX_MATCHES) {
    return "invalid";
  }

  let transaction = closeHistory(editor.state.tr);
  for (const match of [...matches].reverse()) {
    transaction = replaceDocumentTextMatch(transaction, match, replacement);
  }
  if (!transaction.docChanged) return "unchanged";
  const before = editor.state.doc;
  editor.view.dispatch(transaction.scrollIntoView());
  // A guard (a locked control's) may refuse the candidate: nothing changed
  
  return editor.state.doc === before ? "unchanged" : "applied";
}
