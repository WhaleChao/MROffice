/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Handwriting: a handwritten note is an MR Note document on handwriting
// paper, opened on the Draw tab (pens, highlighters, erasers, lasso, ink to
// shapes and math). The ink engine lives in this package; the page it writes
// on is the document editor's (@mrnote/docs).
import type { EditorDefinition, EditorProps } from "@mrnote/core";
import { DocumentEditor, mrdocBytes } from "@mrnote/docs";

import {
  HANDWRITING_PAPER_COLOR_IDS,
  HANDWRITING_PAPER_IDS,
  handwritingDocumentContent,
} from "@mrnote/core/web/lib/handwriting-paper.ts";

export function newHandwritingContent(): Record<string, unknown> {
  return handwritingDocumentContent({
    paper: (HANDWRITING_PAPER_IDS.includes("ruled" as never) ? "ruled" : HANDWRITING_PAPER_IDS[0]) as never,
    paperColor: HANDWRITING_PAPER_COLOR_IDS[0],
    pageSize: "a4" as never,
    orientation: "portrait" as never,
    cover: null,
  });
}

export function HandwritingEditor(props: EditorProps) {
  return <DocumentEditor {...props} extension=".mrink" editorProps={{ initialRibbonTab: "draw" as never }} />;
}

export const inkEditor: EditorDefinition = {
  id: "ink",
  label: "Handwriting",
  extensions: [".mrink", ".mrink.json"],
  newFile: () => ({ name: "Handwriting.mrink", bytes: mrdocBytes("Handwriting", newHandwritingContent()) }),
  Component: HandwritingEditor,
};

export default inkEditor;
