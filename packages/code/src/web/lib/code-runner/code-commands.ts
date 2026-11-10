/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Apply AI document-edit commands to a code page's content.
 *
 * The code page is AI-writable through the same command bus the document,
 * spreadsheet and presentation editors use: the assistant's answer becomes an
 * editable-answer command carrying `set_code_content` (replace) or
 * `append_code_content` (insert below) operations, which this pure helper
 * folds into the editor state. Kept framework-free so it is unit-testable.
 */

import type { NativeDocumentCommand } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { MAX_CODE_SOURCE_CHARS } from "./code-file-import";
import { CODE_LANGUAGE_IDS } from "@mrnote/core/web/lib/code-runner/languages.ts";

export interface CodeContentState {
  language: string;
  content: string;
}

export interface CodeCommandResult {
  state: CodeContentState;
  appliedOperations: number;
}

/** Whether a command carries at least one code operation this window handles. */
export function commandHasCodeOperation(command: NativeDocumentCommand): boolean {
  return command.operations.some(
    (operation) =>
      operation.op === "set_code_content" ||
      operation.op === "append_code_content",
  );
}

/** Fold a command's code operations into the current editor state. */
export function applyCodeCommand(
  current: CodeContentState,
  command: NativeDocumentCommand,
): CodeCommandResult {
  let { language, content } = current;
  let appliedOperations = 0;
  for (const operation of command.operations) {
    if (
      (operation.op === "set_code_content" ||
        operation.op === "append_code_content") &&
      operation.language !== undefined &&
      !CODE_LANGUAGE_IDS.includes(operation.language)
    ) {
      return { state: current, appliedOperations: 0 };
    }
    if (operation.op === "set_code_content") {
      if (operation.content.length > MAX_CODE_SOURCE_CHARS) {
        return { state: current, appliedOperations: 0 };
      }
      content = operation.content;
      if (operation.language) language = operation.language;
      appliedOperations += 1;
    } else if (operation.op === "append_code_content") {
      const separator = content && !content.endsWith("\n") ? "\n" : "";
      if (
        content.length + separator.length + operation.content.length >
        MAX_CODE_SOURCE_CHARS
      ) {
        return { state: current, appliedOperations: 0 };
      }
      content = `${content}${separator}${operation.content}`;
      if (operation.language) language = operation.language;
      appliedOperations += 1;
    }
  }
  return { state: { language, content }, appliedOperations };
}
