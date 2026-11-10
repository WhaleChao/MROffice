"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import PdfEditPointerTools from "../PdfEditPointerTools";
import type { Dispatch, SetStateAction, RefObject } from "react";
import {
  PdfPictureFormatContextRibbon,
  type PdfPictureFormatContextRibbonLabels,
} from "../PdfPictureFormatContextRibbon";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type {
  PdfEditorObjectSelection,
} from "../PdfTextEditor";
import type {
  PdfTextCommentPointerTool,
} from "../PdfCommentRibbonTab";
import type { EmbeddedImageEditKind } from "@/lib/pdf/embedded-image-edits";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfWindowPictureFormatTabGroupsProps {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  pictureFormatRibbonLabels: PdfPictureFormatContextRibbonLabels;
  hasWorkspaceWriteAuthority: () => boolean;
  setNativeImageInsertRequestId: Dispatch<SetStateAction<number>>;
  editorRibbonTab: PdfEditorRibbonTab;
  editorObjectSelectionRef: RefObject<PdfEditorObjectSelection | null>;
  picturePointerTool: PdfTextCommentPointerTool;
  setPicturePointerTool: Dispatch<SetStateAction<PdfTextCommentPointerTool>>;
  leaveTextEditModeToReader: (reason: string) => void;
  exitPictureEditing: () => void;
  invokeImageContextCommand: (action: "replace" | "delete" | "restore" | "opacity" | "rotate-left" | "rotate-right" | "flip-horizontal" | "flip-vertical" | "resize", opacity?: number | undefined, size?: Readonly<{ width: number; height: number; }> | undefined) => void;
  invokeImageCropContextCommand: () => void;
  commitNativeEmbeddedImageArrange: (pageNumber: number, imageOrder: number, direction: "front" | "back" | "forward" | "backward") => Promise<boolean>;
  handleExtractSelectedOriginalImage: () => Promise<void>;
  pictureFormatRibbonModel: { status?: string | undefined; page: number; order: number; width: number; height: number; rotated: boolean; opacity: number; editKind: EmbeddedImageEditKind | undefined; cropSupported: boolean | undefined; rotationSupported: boolean | undefined; rotationBlockedByPage: boolean | undefined; flipSupported: boolean | undefined; hasEdit: boolean; busy: boolean; writable: boolean; disabledReason: string; extractOriginalEnabled: boolean; extractOriginalDisabledReason: string; arrangeEnabled: boolean; arrangeDisabledReason: string; } | null;
}

export function PdfWindowPictureFormatTabGroups({
  t,
  pictureFormatRibbonLabels,
  hasWorkspaceWriteAuthority,
  setNativeImageInsertRequestId,
  editorRibbonTab,
  editorObjectSelectionRef,
  picturePointerTool,
  setPicturePointerTool,
  leaveTextEditModeToReader,
  exitPictureEditing,
  invokeImageContextCommand,
  invokeImageCropContextCommand,
  commitNativeEmbeddedImageArrange,
  handleExtractSelectedOriginalImage,
  pictureFormatRibbonModel,
}: PdfWindowPictureFormatTabGroupsProps) {
  return (
    editorRibbonTab === "picture-format" &&
    pictureFormatRibbonModel ? (
      <PdfPictureFormatContextRibbon
        labels={pictureFormatRibbonLabels}
        model={pictureFormatRibbonModel}
        pointerTools={
          <PdfEditPointerTools
            controlScope="EDIT"
            disabled={false}
            groupLabel={t("pdf.pointerTools")}
            handActive={picturePointerTool === "hand"}
            handActiveLabel={t("pdfViewer.handToolActive")}
            handDisplayLabel={t("pdf.pageWindow.hand")}
            handLabel={t("pdfViewer.handTool")}
            loading={false}
            loadingLabel={t("pdfViewer.loading")}
            onHandActivate={() => setPicturePointerTool("hand")}
            onSelectActivate={() => setPicturePointerTool("select")}
            selectActive={picturePointerTool === "select"}
            selectLabel={t("toolbar.select")}
            tabs="picture-format"
            testIdPrefix="pdf-picture"
          />
        }
        onInsertPicture={() => {
          if (!hasWorkspaceWriteAuthority()) return;
          setNativeImageInsertRequestId((value) => value + 1);
          leaveTextEditModeToReader("insert-picture");
        }}
        onChangePicture={() =>
          invokeImageContextCommand("replace")
        }
        onExtractOriginal={() =>
          void handleExtractSelectedOriginalImage()
        }
        onResetPicture={() =>
          invokeImageContextCommand("restore")
        }
        onDeletePicture={() =>
          invokeImageContextCommand("delete")
        }
        onResizePicture={(width, height) =>
          invokeImageContextCommand("resize", undefined, {
            width,
            height,
          })
        }
        onOpacityChange={(opacity) =>
          invokeImageContextCommand("opacity", opacity)
        }
        onCropPicture={invokeImageCropContextCommand}
        onRotatePicture={(direction) =>
          invokeImageContextCommand(
            direction === "right" ? "rotate-right" : "rotate-left",
          )
        }
        onFlipPicture={(axis) =>
          invokeImageContextCommand(
            axis === "horizontal"
              ? "flip-horizontal"
              : "flip-vertical",
          )
        }
        onArrangePicture={(direction) => {
          const selection = editorObjectSelectionRef.current;
          if (
            selection?.kind !== "embedded-image" ||
            !hasWorkspaceWriteAuthority()
          ) {
            return;
          }
          void commitNativeEmbeddedImageArrange(
            selection.page,
            selection.order,
            direction,
          );
        }}
        onExit={exitPictureEditing}
      />
    ) : null
  );
}
