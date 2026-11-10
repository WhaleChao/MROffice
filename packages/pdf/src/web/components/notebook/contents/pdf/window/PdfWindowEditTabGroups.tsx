"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Dispatch, SetStateAction } from "react";
import {
  PdfSurfaceRibbonGroup,
  PdfSurfaceRibbonCommand,
  PdfSurfaceRibbonCaption,
} from "../PdfSurfaceRibbon";
import {
  OfficeSemanticIcon,
} from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type {
  useAutosavedPageDocument,
} from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type { ImageEditSummary } from "../../../../../lib/pdf/embedded-image-session";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowEditTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workspaceWritable: boolean;
  hasWorkspaceWriteAuthority: () => boolean;
  editorRibbonTab: PdfEditorRibbonTab;
  imageEditMode: boolean;
  setImageEditMode: Dispatch<SetStateAction<boolean>>;
  embeddedImageBusy: boolean;
  embeddedImageStatus: string | null;
  autosave: ReturnType<typeof useAutosavedPageDocument>;
  embeddedImagePending: ImageEditSummary;
  discardEmbeddedImageEdits: () => void;
  handleApplyEmbeddedImageEdits: () => Promise<void>;
}

export function PdfWindowEditTabGroups({
  t,
  workspaceWritable,
  hasWorkspaceWriteAuthority,
  editorRibbonTab,
  imageEditMode,
  setImageEditMode,
  embeddedImageBusy,
  embeddedImageStatus,
  autosave,
  embeddedImagePending,
  discardEmbeddedImageEdits,
  handleApplyEmbeddedImageEdits,
}: PdfWindowEditTabGroupsProps) {
  return (
    editorRibbonTab === "edit" ? (
      <PdfSurfaceRibbonGroup
        label={t("toolbar.ribbonEdit")}
        tab="edit"
        testId="pdf-editor-edit-surface"
      >
        <PdfSurfaceRibbonCommand
          active={imageEditMode}
          controlId="PDFCTL.SURFACE.IMAGE_EDIT"
          disabled={!workspaceWritable || autosave.readOnly}
          icon={<OfficeSemanticIcon semanticId="image.edit" size={24} />}
          label={t("pdfViewer.imageEdit")}
          onClick={() => {
            if (!hasWorkspaceWriteAuthority()) return;
            setImageEditMode((value) => !value);
          }}
          testId="pdf-image-edit-toggle"
          toggle
        />
        {imageEditMode && embeddedImagePending.total > 0 ? (
          <>
            <PdfSurfaceRibbonCommand
              controlId="PDFCTL.SURFACE.IMAGE_EDIT_APPLY"
              disabled={embeddedImageBusy}
              icon={
                <OfficeSemanticIcon semanticId="selection.check" size={24} />
              }
              label={t("pdfViewer.imageEditApply")}
              onClick={() => void handleApplyEmbeddedImageEdits()}
              testId="pdf-image-edit-apply"
            />
            <PdfSurfaceRibbonCommand
              controlId="PDFCTL.SURFACE.IMAGE_EDIT_DISCARD"
              disabled={embeddedImageBusy}
              icon={<OfficeSemanticIcon semanticId="image.reset" size={24} />}
              label={t("pdfViewer.imageEditRestore")}
              onClick={discardEmbeddedImageEdits}
              testId="pdf-image-edit-discard"
            />
            <PdfSurfaceRibbonCaption
              role="status"
              testId="pdf-image-edit-status"
              title={
                embeddedImageStatus ??
                t("pdfViewer.imageEditPending", {
                  count: embeddedImagePending.total,
                })
              }
            >
              {embeddedImageStatus ??
                t("pdfViewer.imageEditPending", {
                  count: embeddedImagePending.total,
                })}
            </PdfSurfaceRibbonCaption>
          </>
        ) : embeddedImageStatus ? (
          <PdfSurfaceRibbonCaption
            role="status"
            testId="pdf-image-edit-status"
            title={embeddedImageStatus}
          >
            {embeddedImageStatus}
          </PdfSurfaceRibbonCaption>
        ) : null}
      </PdfSurfaceRibbonGroup>
    ) : null
  );
}
