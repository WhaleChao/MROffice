/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design tab helpers shared by the document toolbar and the editor-command
 * catalog. The Document Formatting commands (themes, style sets, colours,
 * fonts, paragraph spacing) live in documentDesignThemes.ts.
 */

import type { Editor } from "@tiptap/core";
import { normalizeSectionLayouts, type SectionLayout } from "@mrnote/core/web/components/console/editor/native-document/sectionLayouts.ts";

/**
 * The text area a picture watermark's Auto scale fits: the section's page
 * between its margins, or Word's Letter text area without a section.
 */
export function watermarkTextArea(
  layout: SectionLayout | null,
  editor: Editor,
): { widthPt: number; heightPt: number } {
  const section =
    layout ?? normalizeSectionLayouts(editor.state.doc.attrs.sectionLayouts)?.[0];
  if (!section) return { widthPt: 468, heightPt: 648 };
  const { pageSize, pageMargins } = section;
  return {
    widthPt: Math.max(
      1,
      (pageSize.widthTwips -
        pageMargins.leftTwips -
        pageMargins.rightTwips -
        pageMargins.gutterTwips) /
        20,
    ),
    heightPt: Math.max(
      1,
      (pageSize.heightTwips - pageMargins.topTwips - pageMargins.bottomTwips) /
        20,
    ),
  };
}
