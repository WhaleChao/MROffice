/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfWpsBaseRibbonTab,
  PdfWpsContextualRibbonTab,
} from "../../../../../lib/pdf/ribbon-state";

export type PdfEditorBaseRibbonTab = PdfWpsBaseRibbonTab;
export type PdfEditorContextualRibbonTab = Extract<
  PdfWpsContextualRibbonTab,
  "text-format" | "picture-format" | "comment-tool"
>;
export type PdfEditorRibbonTab = PdfEditorBaseRibbonTab | PdfEditorContextualRibbonTab;

export function isPdfEditorContextualRibbonTab(
  tab: PdfEditorRibbonTab,
): tab is PdfEditorContextualRibbonTab {
  return (
    tab === "text-format" || tab === "picture-format" || tab === "comment-tool"
  );
}
