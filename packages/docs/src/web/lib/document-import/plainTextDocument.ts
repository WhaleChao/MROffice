/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PageImportProjectionBudget } from "./pageImportProjectionBudget";

const UNSUPPORTED_PLAIN_TEXT_CONTROL_RE =
  /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/;

export class PlainTextDocumentImportError extends Error {
  constructor(readonly reason: "unsupported_control") {
    super(
      "Plain text contains control characters that cannot be represented safely.",
    );
    this.name = "PlainTextDocumentImportError";
  }
}

type DocumentNode = Record<string, unknown>;

function textNode(
  text: string,
  budget: PageImportProjectionBudget,
): DocumentNode {
  budget.reserveString(text);
  budget.reserveObjects();
  return { type: "text", text };
}

function hardBreakNode(budget: PageImportProjectionBudget): DocumentNode {
  budget.reserveObjects();
  return { type: "hardBreak" };
}

/**
 * Convert UTF-8 plain text to editable document nodes without trimming it.
 *
 * A single newline remains a soft line break inside the current paragraph.
 * Runs of two or more newlines become paragraph boundaries plus explicit empty
 * paragraphs, which preserves leading, trailing and repeated blank lines when
 * the document is later exported as text. CRLF/CR are canonicalized to LF.
 */
export function plainTextToDocumentNodes(
  text: string,
  budget: PageImportProjectionBudget = new PageImportProjectionBudget(),
): DocumentNode[] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (!normalized) return [];
  if (UNSUPPORTED_PLAIN_TEXT_CONTROL_RE.test(normalized)) {
    throw new PlainTextDocumentImportError("unsupported_control");
  }

  const nodes: DocumentNode[] = [];
  let inline: DocumentNode[] = [];
  let segmentStart = 0;
  let index = 0;
  let endedWithNewline = false;

  const appendText = (end: number) => {
    if (end <= segmentStart) return;
    inline.push(textNode(normalized.slice(segmentStart, end), budget));
  };
  const emitParagraph = () => {
    budget.reserveObjects();
    nodes.push({
      type: "paragraph",
      ...(inline.length > 0 ? { content: inline } : {}),
    });
    inline = [];
  };
  const emitEmptyParagraphs = (count: number) => {
    for (let emitted = 0; emitted < count; emitted += 1) emitParagraph();
  };

  while (index < normalized.length) {
    if (normalized[index] !== "\n") {
      index += 1;
      continue;
    }
    appendText(index);
    const newlineStart = index;
    while (normalized[index] === "\n") index += 1;
    const newlineCount = index - newlineStart;
    const atEnd = index === normalized.length;

    if (newlineCount === 1 && !atEnd) {
      inline.push(hardBreakNode(budget));
    } else {
      emitParagraph();
      emitEmptyParagraphs(atEnd ? newlineCount : newlineCount - 1);
    }
    segmentStart = index;
    endedWithNewline = atEnd;
  }

  if (!endedWithNewline) {
    appendText(normalized.length);
    emitParagraph();
  }
  return nodes;
}
