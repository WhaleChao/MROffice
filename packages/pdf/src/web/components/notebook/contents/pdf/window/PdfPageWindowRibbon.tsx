"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { memo } from "react";
import type { Dispatch, SetStateAction, RefObject, JSX } from "react";
import {
  PdfOfficeRibbon,
  type PdfOfficeActionRail,
} from "../PdfOfficeRibbon";
import { PdfWindowHomeTabGroups } from "./PdfWindowHomeTabGroups";
import { PdfWindowInsertTabGroups } from "./PdfWindowInsertTabGroups";
import { PdfWindowEditTabGroups } from "./PdfWindowEditTabGroups";
import {
  PdfTextEditingContextRibbon,
  type PdfTextEditingContextRibbonAction,
  type PdfTextEditingContextRibbonLabels,
  type PdfTextEditingControl,
} from "../PdfTextEditingContextRibbon";
import { OfficeFormatPreviewProvider } from "@mrnote/core/web/components/office-ribbon/officeFormatPreview.tsx";
import { pdfFormatPreviewAdapter } from "../pdfFormatPreview";
import {
  PdfWindowPictureFormatTabGroups,
} from "./PdfWindowPictureFormatTabGroups";
import { PdfWindowCommentToolTabGroups } from "./PdfWindowCommentToolTabGroups";
import { PdfWindowAnnotateTabGroups } from "./PdfWindowAnnotateTabGroups";
import { PdfWindowDrawTabGroups } from "./PdfWindowDrawTabGroups";
import { PdfWindowPageTabGroups } from "./PdfWindowPageTabGroups";
import {
  PdfSurfaceRibbonGroup,
  PdfSurfaceRibbonCommand,
  PdfSurfaceRibbonCaption,
} from "../PdfSurfaceRibbon";
import {
  ChevronLeft,
  ChevronRight,
  ScrollText,
  Square,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type {
  useAutosavedPageDocument,
  PageSaveStatus,
} from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type {
  PdfNativeListTextSelectionEvidence,
} from "../../../../../lib/pdf/engine/native/list-text-fidelity";
import type {
  PdfNativeListTextEditLease,
} from "../../../../../lib/pdf/engine/native/list-text-command-lifecycle";
import type { PdfPendingPatchbackRepair } from "./insertPageAuthority";
import type { CropBoxEntry } from "../../../../../lib/pdf/crop-box";
import type { PdfInkPenCommands } from "./usePdfInkPens";
import type { PdfDrawRibbonModel, PdfInkTool } from "../PdfDrawRibbonTab";
import type { DocumentInkEraser } from "@mrnote/core/web/lib/document-ink.ts";
import type { PdfEditorRibbonTab } from "./editorRibbonTabs";
import type {
  PdfEditorObjectSelection,
} from "../PdfTextEditor";
import type {
  PdfTextCommentPointerTool,
} from "../PdfCommentRibbonTab";
import type {
  PdfTextBoxContextAction,
  PdfTextBoxSelection,
} from "../PdfTextBoxOverlay";
import type { PdfInkStroke } from "../PdfInkOverlay";
import type {
  PdfPictureFormatContextRibbonLabels,
} from "../PdfPictureFormatContextRibbon";
import type { ImageEditSummary } from "../../../../../lib/pdf/embedded-image-session";
import type { EmbeddedImageEditKind } from "@/lib/pdf/embedded-image-edits";
import type {
  PdfInsertRibbonCommand,
} from "../PdfInsertRibbonTab";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";

export interface PdfPageWindowRibbonModel {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  textEditingRibbonLabels: PdfTextEditingContextRibbonLabels;
  pictureFormatRibbonLabels: PdfPictureFormatContextRibbonLabels;
  workspaceWritable: boolean;
  hasWorkspaceWriteAuthority: () => boolean;
  nativeListTextPendingSelection: PdfNativeListTextSelectionEvidence | null;
  nativeListTextDialogLease: PdfNativeListTextEditLease | null;
  nativeListTextDraftSelection: PdfNativeListTextSelectionEvidence | null;
  setInsertPageDialogOpen: Dispatch<SetStateAction<boolean>>;
  insertPageOperationBusy: boolean;
  setInsertPageErrorMessage: Dispatch<SetStateAction<string | null>>;
  pendingPdfPatchbackRepairRef: RefObject<PdfPendingPatchbackRepair | null>;
  insertPageReturnFocusRef: RefObject<HTMLElement | null>;
  setNativeImageInsertRequestId: Dispatch<SetStateAction<number>>;
  nativeFileAttachmentBusy: boolean;
  nativeFileAttachmentBusyRef: RefObject<boolean>;
  nativePageImportBusy: boolean;
  nativePageImportBusyRef: RefObject<boolean>;
  cropBoxes: CropBoxEntry[];
  /** The Draw tab (PdfWindowDrawTabGroups): the armed tool, the eraser, the pens, the pad. */
  inkTool: PdfInkTool | null;
  setInkTool: (tool: PdfInkTool | null) => void;
  inkEraserKind: DocumentInkEraser;
  setInkEraserKind: Dispatch<SetStateAction<DocumentInkEraser>>;
  inkPenCommands: PdfInkPenCommands;
  inkTrackpad: PdfDrawRibbonModel["trackpad"];
  /** Pen-first drawing: the ribbon folds into the pen palette (null while the palette is not available). */
  penFirstHideRibbon: (() => void) | null;
  ribbonCollapse: { onCollapse: () => void } | undefined;
  penCompact: boolean;
  cropOpen: boolean;
  setCropOpen: Dispatch<SetStateAction<boolean>>;
  cropToggleRef: RefObject<HTMLButtonElement | null>;
  cropDialogReturnFocusTargetRef: RefObject<() => HTMLElement | null>;
  editorRibbonTab: PdfEditorRibbonTab;
  editorObjectSelectionRef: RefObject<PdfEditorObjectSelection | null>;
  textEditViewMode: "continuous" | "single" | "spread";
  setTextEditViewMode: Dispatch<SetStateAction<"continuous" | "single" | "spread">>;
  addTextMode: boolean;
  setAddTextMode: Dispatch<SetStateAction<boolean>>;
  textCommentContextActive: boolean;
  textCommentPointerTool: PdfTextCommentPointerTool;
  picturePointerTool: PdfTextCommentPointerTool;
  setPicturePointerTool: Dispatch<SetStateAction<PdfTextCommentPointerTool>>;
  textCommentFontFamily: string;
  setTextCommentFontFamily: Dispatch<SetStateAction<string>>;
  textCommentFontSizePoints: number;
  setTextCommentFontSizePoints: Dispatch<SetStateAction<number>>;
  textCommentFontColor: string;
  setTextCommentFontColor: Dispatch<SetStateAction<string>>;
  imageEditMode: boolean;
  setImageEditMode: Dispatch<SetStateAction<boolean>>;
  embeddedImageBusy: boolean;
  activateTextCommentContext: () => void;
  leaveTextEditModeToReader: (reason: string) => void;
  handleTextCommentPointerTool: (tool: PdfTextCommentPointerTool) => void;
  exitPictureEditing: () => void;
  handleEditorRibbonTabChange: (tab: PdfEditorRibbonTab) => void;
  invokeTextContextCommand: (action: PdfTextBoxContextAction) => void;
  handleOverlayTextRibbonAction: (action: PdfTextEditingContextRibbonAction) => void;
  invokeImageContextCommand: (action: "replace" | "delete" | "restore" | "opacity" | "rotate-left" | "rotate-right" | "flip-horizontal" | "flip-vertical" | "resize", opacity?: number | undefined, size?: Readonly<{ width: number; height: number; }> | undefined) => void;
  invokeImageCropContextCommand: () => void;
  embeddedImageStatus: string | null;
  pdfPageCount: number;
  currentPage: number;
  setCurrentPageWithIntent: (nextPage: number | ((current: number) => number)) => void;
  autosave: ReturnType<typeof useAutosavedPageDocument>;
  windowSaveStatus: PageSaveStatus;
  /** The face's word for it ("Synced" while the page's live room holds an unsent save). */
  windowSaveStatusLabel?: string;
  managedInkStrokes: PdfInkStroke[];
  runAfterPdfFormulaSourceFlush: (action: () => void | Promise<void>) => void;
  embeddedImagePending: ImageEditSummary;
  discardEmbeddedImageEdits: () => void;
  handleOpenBackgroundDialog: (trigger?: HTMLElement | null) => void;
  handleApplyEmbeddedImageEdits: () => Promise<void>;
  commitNativeEmbeddedImageArrange: (pageNumber: number, imageOrder: number, direction: "front" | "back" | "forward" | "backward") => Promise<boolean>;
  canUndoNativeReceipt: boolean;
  canRedoNativeReceipt: boolean;
  executeNativeUndo: () => Promise<boolean>;
  executeNativeRedo: () => Promise<boolean>;
  pdfInsertPageInventoryReady: boolean;
  undoLastInkStroke: () => void;
  workspaceWriteDeniedMessage: string;
  overlayTextSelection: ({ kind: "text"; } & PdfTextBoxSelection) | null;
  overlayTextCapabilities: Readonly<Record<PdfTextEditingControl, Readonly<{ enabled: boolean; reason?: string | undefined; }>>>;
  overlayTextRibbonModel: { wordSpacing: number | null; horizontalScale: number | null; keyboardInput: boolean; kind: "text" | "formula"; saveState: "saved" | "saving" | "ready" | "blocked"; lineSpacing: number; characterSpacing: number; firstLineIndent: number; leftIndent: number; rightIndent: number; target: "overlay-box"; fontSections: { id: string; options: { id: string; label: string; value: string; previewFamily: string; selected: boolean; }[]; }[]; activeFontLabel: string; activeFontFamily: string | undefined; fontSize: number; bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean; subscript: boolean; superscript: boolean; highlight: boolean; color: string; alignment: "justify" | "left" | "center" | "right" | "distributed" | null; } | { wordSpacing: number | null; horizontalScale: number | null; keyboardInput: boolean; kind: "text" | "formula"; saveState: "saved" | "saving" | "ready" | "blocked"; lineSpacing: null; characterSpacing: null; firstLineIndent: null; leftIndent: null; rightIndent: null; target: "overlay-box"; fontSections: { id: string; options: { id: string; label: string; value: string; previewFamily: string; selected: boolean; }[]; }[]; activeFontLabel: string; activeFontFamily: string | undefined; fontSize: number; bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean; subscript: boolean; superscript: boolean; highlight: boolean; color: string; alignment: "justify" | "left" | "center" | "right" | "distributed" | null; } | null;
  pendingNativeTextRibbonReason: string;
  pendingNativeTextCapabilities: Readonly<Record<PdfTextEditingControl, Readonly<{ enabled: boolean; reason?: string | undefined; }>>>;
  pendingNativeTextRibbonModel: { target: "native-list"; fontSections: never[]; activeFontLabel: string; fontSize: null; bold: false; italic: false; underline: false; strikethrough: false; subscript: false; superscript: false; highlight: false; alignment: null; lineSpacing: null; characterSpacing: null; firstLineIndent: null; wordSpacing: null; leftIndent: null; rightIndent: null; horizontalScale: null; keyboardInput: boolean; saveState: "blocked"; };
  handleExtractSelectedOriginalImage: () => Promise<void>;
  pictureFormatRibbonModel: { status?: string | undefined; page: number; order: number; width: number; height: number; rotated: boolean; opacity: number; editKind: EmbeddedImageEditKind | undefined; cropSupported: boolean | undefined; rotationSupported: boolean | undefined; rotationBlockedByPage: boolean | undefined; flipSupported: boolean | undefined; hasEdit: boolean; busy: boolean; writable: boolean; disabledReason: string; extractOriginalEnabled: boolean; extractOriginalDisabledReason: string; arrangeEnabled: boolean; arrangeDisabledReason: string; } | null;
  editorRibbonTabs: ({ id: "home" | "insert" | "draw" | "edit" | "page" | "annotate" | "tools"; label: string; ariaLabel?: string | undefined; controlId?: string | undefined; surface?: "home" | "insert" | undefined; } | { id: "text-format"; label: string; className: string; contextual: string; controlId: string; } | { id: "picture-format"; label: string; className: string; contextual: string; controlId: string; } | { id: "comment-tool"; label: string; className: string; contextual: string; controlId: string; })[];
  pageStructureRibbonActions: PdfInsertRibbonCommand[];
  workspacePrimaryActions: JSX.Element;
  workspaceOverflowOnlyActions: JSX.Element;
  workspaceActionRail: PdfOfficeActionRail;
  workspaceActionTail: JSX.Element;
}

// The tabs whose groups take only the values they show and stable commands
// skip the window's renders that change none of them (every page change while
// scrolling). The others take the autosave object, renewed every render.
const MemoPdfWindowHomeTabGroups = memo(PdfWindowHomeTabGroups);
const MemoPdfWindowPictureFormatTabGroups = memo(PdfWindowPictureFormatTabGroups);
const MemoPdfWindowPageTabGroups = memo(PdfWindowPageTabGroups);

export function PdfPageWindowRibbon({ ribbon }: { ribbon: PdfPageWindowRibbonModel }) {
  const {
    t,
    textEditingRibbonLabels,
    pictureFormatRibbonLabels,
    workspaceWritable,
    hasWorkspaceWriteAuthority,
    nativeListTextPendingSelection,
    nativeListTextDialogLease,
    nativeListTextDraftSelection,
    setInsertPageDialogOpen,
    insertPageOperationBusy,
    setInsertPageErrorMessage,
    pendingPdfPatchbackRepairRef,
    insertPageReturnFocusRef,
    setNativeImageInsertRequestId,
    nativeFileAttachmentBusy,
    nativeFileAttachmentBusyRef,
    nativePageImportBusy,
    nativePageImportBusyRef,
    cropBoxes,
    inkTool,
    setInkTool,
    inkEraserKind,
    setInkEraserKind,
    inkPenCommands,
    inkTrackpad,
    penFirstHideRibbon,
    ribbonCollapse,
    penCompact,
    cropOpen,
    setCropOpen,
    cropToggleRef,
    cropDialogReturnFocusTargetRef,
    editorRibbonTab,
    editorObjectSelectionRef,
    textEditViewMode,
    setTextEditViewMode,
    addTextMode,
    setAddTextMode,
    textCommentContextActive,
    textCommentPointerTool,
    picturePointerTool,
    setPicturePointerTool,
    textCommentFontFamily,
    setTextCommentFontFamily,
    textCommentFontSizePoints,
    setTextCommentFontSizePoints,
    textCommentFontColor,
    setTextCommentFontColor,
    imageEditMode,
    setImageEditMode,
    embeddedImageBusy,
    activateTextCommentContext,
    leaveTextEditModeToReader,
    handleTextCommentPointerTool,
    exitPictureEditing,
    handleEditorRibbonTabChange,
    invokeTextContextCommand,
    handleOverlayTextRibbonAction,
    invokeImageContextCommand,
    invokeImageCropContextCommand,
    embeddedImageStatus,
    pdfPageCount,
    currentPage,
    setCurrentPageWithIntent,
    autosave,
    windowSaveStatus,
    windowSaveStatusLabel,
    managedInkStrokes,
    runAfterPdfFormulaSourceFlush,
    embeddedImagePending,
    discardEmbeddedImageEdits,
    handleOpenBackgroundDialog,
    handleApplyEmbeddedImageEdits,
    commitNativeEmbeddedImageArrange,
    canUndoNativeReceipt,
    canRedoNativeReceipt,
    executeNativeUndo,
    executeNativeRedo,
    pdfInsertPageInventoryReady,
    undoLastInkStroke,
    workspaceWriteDeniedMessage,
    overlayTextSelection,
    overlayTextCapabilities,
    overlayTextRibbonModel,
    pendingNativeTextRibbonReason,
    pendingNativeTextCapabilities,
    pendingNativeTextRibbonModel,
    handleExtractSelectedOriginalImage,
    pictureFormatRibbonModel,
    editorRibbonTabs,
    pageStructureRibbonActions,
    workspacePrimaryActions,
    workspaceOverflowOnlyActions,
    workspaceActionRail,
    workspaceActionTail,
  } = ribbon;
  return (
    <PdfOfficeRibbon
      collapse={ribbonCollapse}
      penCompact={penCompact}
      activeTab={editorRibbonTab}
      ariaLabel={t("pdf.editText")}
      className="pdf-page-window__text-ribbon"
      onActiveTabChange={handleEditorRibbonTabChange}
      actionRail={workspaceActionRail}
      quickOverflowLabel={t("toolbar.quickOverflow")}
      saveStatus={windowSaveStatus}
      saveStatusLabel={windowSaveStatusLabel ?? t(`pages.${windowSaveStatus}` as "pages.saving")}
      scrollLeftLabel={t("toolbar.ribbonScrollLeft")}
      scrollRightLabel={t("toolbar.ribbonScrollRight")}
      tabs={editorRibbonTabs}
      testId="pdf-editor-ribbon"
    >
      <MemoPdfWindowHomeTabGroups
        t={t}
        nativeFileAttachmentBusy={nativeFileAttachmentBusy}
        nativeFileAttachmentBusyRef={nativeFileAttachmentBusyRef}
        nativePageImportBusy={nativePageImportBusy}
        nativePageImportBusyRef={nativePageImportBusyRef}
        editorRibbonTab={editorRibbonTab}
        canUndoNativeReceipt={canUndoNativeReceipt}
        canRedoNativeReceipt={canRedoNativeReceipt}
        executeNativeUndo={executeNativeUndo}
        executeNativeRedo={executeNativeRedo}
      />
      {/* The workspace's own actions (generate a deck, back to reading,
          replace / download / export Markdown, the authority retry) in a
          Home group of their own, as in the reader. The action rail never
          draws them, and since the "more" menu was retired (b60b653bbf) this
          surface showed them nowhere: after adding a text box from Insert,
          only Insert Image led back to reading. */}
      {editorRibbonTab === "home" ? (
        <PdfSurfaceRibbonGroup
          label={t("pdf.viewer.document")}
          tab="home"
          testId="pdf-editor-home-workspace-actions"
        >
          {workspacePrimaryActions}
          {workspaceOverflowOnlyActions}
          {workspaceActionTail}
        </PdfSurfaceRibbonGroup>
      ) : null}

      <PdfWindowInsertTabGroups
        t={t}
        workspaceWritable={workspaceWritable}
        hasWorkspaceWriteAuthority={hasWorkspaceWriteAuthority}
        setInsertPageDialogOpen={setInsertPageDialogOpen}
        insertPageOperationBusy={insertPageOperationBusy}
        setInsertPageErrorMessage={setInsertPageErrorMessage}
        pendingPdfPatchbackRepairRef={pendingPdfPatchbackRepairRef}
        insertPageReturnFocusRef={insertPageReturnFocusRef}
        setNativeImageInsertRequestId={setNativeImageInsertRequestId}
        editorRibbonTab={editorRibbonTab}
        addTextMode={addTextMode}
        setAddTextMode={setAddTextMode}
        leaveTextEditModeToReader={leaveTextEditModeToReader}
        autosave={autosave}
        handleOpenBackgroundDialog={handleOpenBackgroundDialog}
        pdfInsertPageInventoryReady={pdfInsertPageInventoryReady}
      />

      <PdfWindowEditTabGroups
        t={t}
        workspaceWritable={workspaceWritable}
        hasWorkspaceWriteAuthority={hasWorkspaceWriteAuthority}
        editorRibbonTab={editorRibbonTab}
        imageEditMode={imageEditMode}
        setImageEditMode={setImageEditMode}
        embeddedImageBusy={embeddedImageBusy}
        embeddedImageStatus={embeddedImageStatus}
        autosave={autosave}
        embeddedImagePending={embeddedImagePending}
        discardEmbeddedImageEdits={discardEmbeddedImageEdits}
        handleApplyEmbeddedImageEdits={handleApplyEmbeddedImageEdits}
      />

      {editorRibbonTab === "text-format" &&
      (nativeListTextDialogLease !== null || nativeListTextDraftSelection !== null) ? (
        <span
          data-pdf-native-list-text-ribbon-host="true"
          data-pdf-text-target="native-list"
          data-testid="pdf-native-list-text-ribbon-host"
          hidden
        />
      ) : editorRibbonTab === "text-format" &&
        nativeListTextPendingSelection !== null ? (
        <PdfTextEditingContextRibbon
          capabilities={pendingNativeTextCapabilities}
          labels={textEditingRibbonLabels}
          model={pendingNativeTextRibbonModel}
          onAction={() => undefined}
          status={
            <span role="status" data-testid="pdf-native-text-proof-pending">
              {pendingNativeTextRibbonReason}
            </span>
          }
        />
      ) : editorRibbonTab === "text-format" &&
        overlayTextRibbonModel ? (
        <OfficeFormatPreviewProvider
          adapter={pdfFormatPreviewAdapter({
            currentFontSizePt: () => overlayTextRibbonModel.fontSize,
          })}
        >
          <PdfTextEditingContextRibbon
            capabilities={overlayTextCapabilities}
            labels={textEditingRibbonLabels}
            model={overlayTextRibbonModel}
            onAction={handleOverlayTextRibbonAction}
          />
        </OfficeFormatPreviewProvider>
      ) : null}

      <MemoPdfWindowPictureFormatTabGroups
        t={t}
        pictureFormatRibbonLabels={pictureFormatRibbonLabels}
        hasWorkspaceWriteAuthority={hasWorkspaceWriteAuthority}
        setNativeImageInsertRequestId={setNativeImageInsertRequestId}
        editorRibbonTab={editorRibbonTab}
        editorObjectSelectionRef={editorObjectSelectionRef}
        picturePointerTool={picturePointerTool}
        setPicturePointerTool={setPicturePointerTool}
        leaveTextEditModeToReader={leaveTextEditModeToReader}
        exitPictureEditing={exitPictureEditing}
        invokeImageContextCommand={invokeImageContextCommand}
        invokeImageCropContextCommand={invokeImageCropContextCommand}
        commitNativeEmbeddedImageArrange={commitNativeEmbeddedImageArrange}
        handleExtractSelectedOriginalImage={handleExtractSelectedOriginalImage}
        pictureFormatRibbonModel={pictureFormatRibbonModel}
      />

      <PdfWindowCommentToolTabGroups
        t={t}
        workspaceWritable={workspaceWritable}
        hasWorkspaceWriteAuthority={hasWorkspaceWriteAuthority}
        editorRibbonTab={editorRibbonTab}
        textCommentContextActive={textCommentContextActive}
        textCommentPointerTool={textCommentPointerTool}
        textCommentFontFamily={textCommentFontFamily}
        setTextCommentFontFamily={setTextCommentFontFamily}
        textCommentFontSizePoints={textCommentFontSizePoints}
        setTextCommentFontSizePoints={setTextCommentFontSizePoints}
        textCommentFontColor={textCommentFontColor}
        setTextCommentFontColor={setTextCommentFontColor}
        handleTextCommentPointerTool={handleTextCommentPointerTool}
        invokeTextContextCommand={invokeTextContextCommand}
        autosave={autosave}
        workspaceWriteDeniedMessage={workspaceWriteDeniedMessage}
        overlayTextSelection={overlayTextSelection}
      />

      <PdfWindowDrawTabGroups
        t={t}
        workspaceWritable={workspaceWritable}
        hasWorkspaceWriteAuthority={hasWorkspaceWriteAuthority}
        editorRibbonTab={editorRibbonTab}
        autosave={autosave}
        workspaceWriteDeniedMessage={workspaceWriteDeniedMessage}
        inkTool={inkTool}
        setInkTool={setInkTool}
        inkEraserKind={inkEraserKind}
        setInkEraserKind={setInkEraserKind}
        inkPenCommands={inkPenCommands}
        managedInkStrokes={managedInkStrokes}
        undoLastInkStroke={undoLastInkStroke}
        trackpad={inkTrackpad}
        hideRibbon={penFirstHideRibbon}
      />

      <PdfWindowAnnotateTabGroups
        t={t}
        workspaceWritable={workspaceWritable}
        editorRibbonTab={editorRibbonTab}
        activateTextCommentContext={activateTextCommentContext}
        autosave={autosave}
        workspaceWriteDeniedMessage={workspaceWriteDeniedMessage}
      />

      <MemoPdfWindowPageTabGroups
        t={t}
        cropBoxes={cropBoxes}
        cropOpen={cropOpen}
        setCropOpen={setCropOpen}
        cropToggleRef={cropToggleRef}
        cropDialogReturnFocusTargetRef={cropDialogReturnFocusTargetRef}
        editorRibbonTab={editorRibbonTab}
        currentPage={currentPage}
        pageStructureRibbonActions={pageStructureRibbonActions}
      />

      {editorRibbonTab === "tools" ? (
        <PdfSurfaceRibbonGroup
          label={t("toolbar.ribbonPageView")}
          tab="tools"
          testId="pdf-editor-tools-surface"
        >
          <PdfSurfaceRibbonCommand
            controlId="PDFCTL.SURFACE.PREV_PAGE"
            disabled={currentPage <= 1}
            icon={<ChevronLeft size={24} />}
            label={t("pdfViewer.prev")}
            onClick={() =>
              runAfterPdfFormulaSourceFlush(() =>
                setCurrentPageWithIntent((value) =>
                  Math.max(1, value - 1),
                ),
              )
            }
            testId="pdf-text-editor-prev-page"
          />
          <PdfSurfaceRibbonCaption testId="pdf-text-editor-page-of">
            {t("pdfViewer.pageOf", {
              current: currentPage,
              total: pdfPageCount || currentPage,
            })}
          </PdfSurfaceRibbonCaption>
          <PdfSurfaceRibbonCommand
            controlId="PDFCTL.SURFACE.NEXT_PAGE"
            disabled={pdfPageCount > 0 && currentPage >= pdfPageCount}
            icon={<ChevronRight size={24} />}
            label={t("pdfViewer.next")}
            onClick={() =>
              runAfterPdfFormulaSourceFlush(() =>
                setCurrentPageWithIntent((value) =>
                  pdfPageCount
                    ? Math.min(pdfPageCount, value + 1)
                    : value + 1,
                ),
              )
            }
            testId="pdf-text-editor-next-page"
          />
          <PdfSurfaceRibbonCommand
            active={textEditViewMode === "continuous"}
            controlId="PDFCTL.SURFACE.VIEW_MODE"
            icon={
              textEditViewMode === "continuous" ? (
                <ScrollText size={24} />
              ) : (
                <Square size={24} />
              )
            }
            label={
              textEditViewMode === "continuous"
                ? t("pdfViewer.viewModeSingle")
                : t("pdfViewer.viewModeContinuous")
            }
            onClick={() =>
              runAfterPdfFormulaSourceFlush(() =>
                setTextEditViewMode((mode) =>
                  mode === "continuous" ? "single" : "continuous",
                ),
              )
            }
            testId="pdf-text-editor-continuous-toggle"
            toggle
          />
        </PdfSurfaceRibbonGroup>
      ) : null}
    </PdfOfficeRibbon>
  );
}
