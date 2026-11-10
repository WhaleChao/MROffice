"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import PdfEditPointerTools from "../PdfEditPointerTools";
import type { Dispatch, SetStateAction } from "react";
import {
  PdfTextCommentToolContextRibbon,
  type PdfTextCommentPointerTool,
} from "../PdfCommentRibbonTab";
import { PDF_TEXT_COMMENT_FONT_FAMILIES } from "./editorDefaults";
import type {
  useAutosavedPageDocument,
} from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type {
  PdfTextBoxContextAction,
  PdfTextBoxSelection,
} from "../PdfTextBoxOverlay";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowCommentToolTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workspaceWritable: boolean;
  hasWorkspaceWriteAuthority: () => boolean;
  editorRibbonTab: PdfEditorRibbonTab;
  textCommentContextActive: boolean;
  textCommentPointerTool: PdfTextCommentPointerTool;
  textCommentFontFamily: string;
  setTextCommentFontFamily: Dispatch<SetStateAction<string>>;
  textCommentFontSizePoints: number;
  setTextCommentFontSizePoints: Dispatch<SetStateAction<number>>;
  textCommentFontColor: string;
  setTextCommentFontColor: Dispatch<SetStateAction<string>>;
  handleTextCommentPointerTool: (tool: PdfTextCommentPointerTool) => void;
  invokeTextContextCommand: (action: PdfTextBoxContextAction) => void;
  autosave: ReturnType<typeof useAutosavedPageDocument>;
  workspaceWriteDeniedMessage: string;
  overlayTextSelection: ({ kind: "text"; } & PdfTextBoxSelection) | null;
}

export function PdfWindowCommentToolTabGroups({
  t,
  workspaceWritable,
  hasWorkspaceWriteAuthority,
  editorRibbonTab,
  textCommentContextActive,
  textCommentPointerTool,
  textCommentFontFamily,
  setTextCommentFontFamily,
  textCommentFontSizePoints,
  setTextCommentFontSizePoints,
  textCommentFontColor,
  setTextCommentFontColor,
  handleTextCommentPointerTool,
  invokeTextContextCommand,
  autosave,
  workspaceWriteDeniedMessage,
  overlayTextSelection,
}: PdfWindowCommentToolTabGroupsProps) {
  return (
    editorRibbonTab === "comment-tool" &&
    textCommentContextActive ? (
      <PdfTextCommentToolContextRibbon
        fontFamilies={
          overlayTextSelection
            ? overlayTextSelection.fontOptions.map((option) => ({
                label: option.label,
                value: option.id,
              }))
            : PDF_TEXT_COMMENT_FONT_FAMILIES
        }
        labels={{
          font: t("pdf.pageWindow.font"),
          fontColor: t("pdf.pageWindow.fontColor"),
          fontFamily: t("pdf.pageWindow.font"),
          fontSize: t("pdf.pageWindow.fontSize"),
          textComment: t("pdf.pageWindow.textComment"),
        }}
        model={{
          activeTool: textCommentPointerTool,
          disabledReason: workspaceWriteDeniedMessage,
          fontColor:
            overlayTextSelection?.color ?? textCommentFontColor,
          fontFamily:
            overlayTextSelection?.fontFamilyId ??
            textCommentFontFamily,
          fontSizePoints:
            overlayTextSelection?.fontSizeDocumentUnit ??
            textCommentFontSizePoints,
          writable: workspaceWritable && !autosave.readOnly,
        }}
        onFontColorChange={(color) => {
          if (!hasWorkspaceWriteAuthority()) return;
          setTextCommentFontColor(color);
          if (overlayTextSelection) {
            invokeTextContextCommand({
              action: "color",
              value: color,
            });
          }
        }}
        onFontFamilyChange={(fontFamily) => {
          if (!hasWorkspaceWriteAuthority()) return;
          setTextCommentFontFamily(fontFamily);
          if (overlayTextSelection) {
            invokeTextContextCommand({
              action: "font-family",
              optionId: fontFamily,
            });
          }
        }}
        onFontSizeChange={(fontSizePoints) => {
          if (!hasWorkspaceWriteAuthority()) return;
          setTextCommentFontSizePoints(fontSizePoints);
          if (overlayTextSelection) {
            invokeTextContextCommand({
              action: "font-size",
              valueDocumentUnit: fontSizePoints,
            });
          }
        }}
        onToolChange={handleTextCommentPointerTool}
        pointerTools={
          <PdfEditPointerTools
            controlScope="COMMENT"
            disabled={false}
            groupLabel={t("pdf.pointerTools")}
            handActive={textCommentPointerTool === "hand"}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={t("pdf.pageWindow.hand")}
            handLabel={t("pdfViewer.handTool")}
            loading={false}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={() =>
              handleTextCommentPointerTool("hand")
            }
            onSelectActivate={() =>
              handleTextCommentPointerTool("select")
            }
            selectActive={textCommentPointerTool === "select"}
            selectLabel={t("toolbar.select")}
            tabs="comment-tool"
            testIdPrefix="pdf-text-comment"
          />
        }
      />
    ) : null
  );
}
