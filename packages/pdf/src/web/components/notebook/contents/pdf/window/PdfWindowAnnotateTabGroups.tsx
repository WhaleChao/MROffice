"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PdfSurfaceRibbonGroup,
  PdfSurfaceRibbonCommand,
} from "../PdfSurfaceRibbon";
import {
  PdfInsertRibbonIcon,
} from "../PdfInsertRibbonIcons";
import type {
  useAutosavedPageDocument,
} from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowAnnotateTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workspaceWritable: boolean;
  editorRibbonTab: PdfEditorRibbonTab;
  activateTextCommentContext: () => void;
  autosave: ReturnType<typeof useAutosavedPageDocument>;
  workspaceWriteDeniedMessage: string;
}






export function PdfWindowAnnotateTabGroups({
  t,
  workspaceWritable,
  editorRibbonTab,
  activateTextCommentContext,
  autosave,
  workspaceWriteDeniedMessage,
}: PdfWindowAnnotateTabGroupsProps) {
  if (editorRibbonTab !== "annotate") return null;
  const writable = workspaceWritable && !autosave.readOnly;
  const deniedReason = writable ? undefined : workspaceWriteDeniedMessage;
  return (
    <PdfSurfaceRibbonGroup
      label={t("pdfViewer.annotationsPanel")}
      tab="annotate"
      testId="pdf-editor-annotate-surface"
    >
      <PdfSurfaceRibbonCommand
        controlId="PDFCTL.SURFACE.TEXT_COMMENT"
        disabled={!writable}
        disabledReason={deniedReason}
        icon={<PdfInsertRibbonIcon name="comment" size={24} />}
        label={t("pdf.pageWindow.textComment")}
        onClick={() => activateTextCommentContext()}
        testId="pdf-editor-text-comment"
      />
    </PdfSurfaceRibbonGroup>
  );
}
