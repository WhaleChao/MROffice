"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ImageDown } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useEffect, useRef, useState, type ReactNode } from "react";

import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  OfficeHomeCommand,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeNumberField,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { OfficePictureTransparencyMenu } from "@mrnote/core/web/components/office-ribbon/OfficePictureAdjustMenus.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";

import styles from "./PdfPictureFormatContextRibbon.module.css";
import {
  PdfSurfaceRibbonCaption,
  PdfSurfaceRibbonCommand,
  PdfSurfaceRibbonFieldRow,
  PdfSurfaceRibbonGroup,
} from "./PdfSurfaceRibbon";
import type { PdfRibbonTranslator } from "./pdfRibbonTranslator";
import ribbonLayout from "./pdfRibbonLayout.module.css";

export interface PdfPictureFormatContextRibbonModel {
  page: number;
  order: number;
  width: number;
  height: number;
  rotated: boolean;
  opacity: number;
  editKind?:
    | "delete"
    | "replace"
    | "move"
    | "opacity"
    | "crop"
    | "rotate"
    | "resize"
    | "flip";
  cropSupported?: boolean;
  rotationSupported?: boolean;
  rotationBlockedByPage?: boolean;
  flipSupported?: boolean;
  hasEdit: boolean;
  busy: boolean;
  writable: boolean;
  disabledReason?: string;
  extractOriginalEnabled?: boolean;
  extractOriginalDisabledReason?: string;
  arrangeEnabled?: boolean;
  arrangeDisabledReason?: string;
  /** Last picture-edit message, so a refusal is visible on this surface. */
  status?: string;
  /** The picture's own size in points (the reader's WPS layout shows its
   * size fields in pt, as WPS does, instead of page percentages). */
  widthPoints?: number;
  heightPoints?: number;
  /** Why 清晰化 cannot run now (it is listed only with a host). */
  sharpenDisabledReason?: string;
}

export interface PdfPictureFormatContextRibbonLabels {
  picture: string;
  adjust: string;
  arrange: string;
  size: string;
  actions: string;
  originalPicture: string;
  editedPicture: string;
  rotatedPicture: string;
  pictureWidth: string;
  pictureHeight: string;
  insertPicture: string;
  changePicture: string;
  extractOriginal: string;
  resetPicture: string;
  pictureTransparency: string;
  cropPicture: string;
  cropRequiresReset: string;
  cropUnsupported: string;
  rotateLeft: string;
  rotateRight: string;
  rotateUnsupported: string;
  rotateOutOfPage: string;
  flipHorizontal: string;
  flipVertical: string;
  flipUnsupported: string;
  noEditsToReset: string;
  deletePicture: string;
  bringToFront: string;
  sendToBack: string;
  bringForward: string;
  sendBackward: string;
  backToEditing: string;
  /** WPS 图片编辑 layout: its own command names and groups. */
  flip: string;
  sharpen: string;
  extract: string;
  extractPicture: string;
  wpsRotateLeft: string;
  wpsRotateRight: string;
  wpsDelete: string;
  exitEditing: string;
  points: string;
  pictureNumber: (order: number) => string;
  pageNumber: (page: number) => string;
}

export function createPdfPictureFormatContextRibbonLabels(
  t: PdfRibbonTranslator,
): PdfPictureFormatContextRibbonLabels {
  return {
    picture: t("pdfContextRibbon.picture.picture"),
    adjust: t("pdfContextRibbon.picture.adjust"),
    arrange: t("pdfContextRibbon.picture.arrange"),
    size: t("pdfContextRibbon.picture.size"),
    actions: t("pdfContextRibbon.picture.actions"),
    originalPicture: t("pdfContextRibbon.picture.originalPicture"),
    editedPicture: t("pdfContextRibbon.picture.editedPicture"),
    rotatedPicture: t("pdfContextRibbon.picture.rotatedPicture"),
    pictureWidth: t("pdfContextRibbon.picture.pictureWidth"),
    pictureHeight: t("pdfContextRibbon.picture.pictureHeight"),
    insertPicture: t("pdfContextRibbon.picture.insertPicture"),
    changePicture: t("pdfContextRibbon.picture.changePicture"),
    extractOriginal: t("pdfContextRibbon.picture.extractOriginal"),
    resetPicture: t("pdfContextRibbon.picture.resetPicture"),
    pictureTransparency: t("pdfContextRibbon.picture.pictureTransparency"),
    cropPicture: t("pdfContextRibbon.picture.cropPicture"),
    cropRequiresReset: t("pdfContextRibbon.picture.cropRequiresReset"),
    cropUnsupported: t("pdfContextRibbon.picture.cropUnsupported"),
    rotateLeft: t("pdfContextRibbon.picture.rotateLeft"),
    rotateRight: t("pdfContextRibbon.picture.rotateRight"),
    rotateUnsupported: t("pdfContextRibbon.picture.rotateUnsupported"),
    rotateOutOfPage: t("pdfContextRibbon.picture.rotateOutOfPage"),
    flipHorizontal: t("pdfContextRibbon.picture.flipHorizontal"),
    flipVertical: t("pdfContextRibbon.picture.flipVertical"),
    flipUnsupported: t("pdfContextRibbon.picture.flipUnsupported"),
    noEditsToReset: t("pdfContextRibbon.picture.noEditsToReset"),
    deletePicture: t("pdfContextRibbon.picture.deletePicture"),
    bringToFront: t("pdfContextRibbon.picture.bringToFront"),
    sendToBack: t("pdfContextRibbon.picture.sendToBack"),
    bringForward: t("pdfContextRibbon.picture.bringForward"),
    sendBackward: t("pdfContextRibbon.picture.sendBackward"),
    backToEditing: t("pdfContextRibbon.picture.backToEditing"),
    flip: t("pdfContextRibbon.picture.flip"),
    sharpen: t("pdfContextRibbon.picture.sharpen"),
    extract: t("pdfContextRibbon.picture.extract"),
    extractPicture: t("pdfContextRibbon.picture.extractPicture"),
    wpsRotateLeft: t("pdfContextRibbon.picture.wpsRotateLeft"),
    wpsRotateRight: t("pdfContextRibbon.picture.wpsRotateRight"),
    wpsDelete: t("pdfContextRibbon.picture.wpsDelete"),
    exitEditing: t("pdfContextRibbon.picture.exitEditing"),
    points: t("pdfContextRibbon.picture.points"),
    pictureNumber: (order) => t("pdfContextRibbon.picture.pictureNumber", { order: order + 1 }),
    pageNumber: (page) => t("pdfContextRibbon.picture.pageNumber", { page }),
  };
}

export function PdfPictureFormatContextRibbon({
  labels,
  model,
  layout = "office",
  pointerTools,
  onInsertPicture,
  onExtractOriginal,
  onChangePicture,
  onResetPicture,
  onDeletePicture,
  onOpacityChange,
  onSharpenPicture,
  onCropPicture,
  onRotatePicture,
  onFlipPicture,
  onArrangePicture,
  onResizePicture,
  onResizePicturePoints,
  onExit,
}: {
  labels: PdfPictureFormatContextRibbonLabels;
  model: PdfPictureFormatContextRibbonModel;
  /**
   * `office`: the authoring surface's tab (picture, adjust, arrange, size,
   * actions). `wps`: the reader's 图片编辑 tab in WPS's order -- 插入图片 ·
   * 翻转 · 左/右旋转 · 替换图片 · 删除 | 透明度 · 清晰化 · 裁剪 | 提取图片 |
   * 退出编辑 -- then this editor's own arrange and size (pt) groups.
   */
  layout?: "office" | "wps";
  pointerTools?: ReactNode;
  onInsertPicture?: () => void;
  onExtractOriginal?: () => void;
  onChangePicture: () => void;
  /** Drop a staged edit (the authoring surface only; absent, not listed). */
  onResetPicture?: () => void;
  onDeletePicture: () => void;
  /** 清晰化: enhance the picture's pixels; absent, the command is not listed. */
  onSharpenPicture?: () => void;
  onOpacityChange: (opacity: number) => void;
  onCropPicture: () => void;
  onRotatePicture: (direction: "left" | "right") => void;
  onFlipPicture: (axis: "horizontal" | "vertical") => void;
  /** Native page z-order: move the picture in front of or behind the page. */
  onArrangePicture?: (direction: "front" | "back" | "forward" | "backward") => void;
  /** Numeric size, page-relative percentages; keeps the picture's top-left. */
  onResizePicture?: (width: number, height: number) => void;
  /** Numeric size in points (with `model.widthPoints`/`heightPoints`). */
  onResizePicturePoints?: (width: number, height: number) => void;
  onExit: () => void;
}) {
  const mutationDisabled = model.busy || !model.writable;
  const mutationTitle = model.writable ? undefined : model.disabledReason;
  const resetDisabled = mutationDisabled || !model.hasEdit;
  // A rotated or skewed paint can still take a constant alpha: the native
  // writer only names a graphics state, it does not re-fit pixels.
  const opacityDisabled = mutationDisabled;
  const cropDisabled =
    mutationDisabled ||
    model.rotated ||
    model.cropSupported === false ||
    model.editKind === "replace";
  const cropTitle = model.rotated
    ? labels.rotatedPicture
    : model.editKind === "replace"
      ? labels.cropRequiresReset
      : model.cropSupported === false
        ? labels.cropUnsupported
        : mutationTitle;
  const rotateDisabled =
    mutationDisabled || model.rotated || model.rotationSupported === false;
  const rotateTitle = rotateDisabled
    ? model.rotated
      ? labels.rotatedPicture
      : model.rotationSupported === false
        ? model.rotationBlockedByPage
          ? labels.rotateOutOfPage
          : labels.rotateUnsupported
        : mutationTitle
    : undefined;
  const flipDisabled =
    mutationDisabled || model.rotated || model.flipSupported === false;
  const flipTitle = flipDisabled
    ? model.rotated
      ? labels.rotatedPicture
      : model.flipSupported === false
        ? labels.flipUnsupported
        : mutationTitle
    : undefined;
  const resetTitle = !model.writable
    ? model.disabledReason
    : !model.hasEdit
      ? labels.noEditsToReset
      : undefined;
  const opacityTitle = opacityDisabled
    ? mutationTitle
    : undefined;
  const commandLabel = (label: string, reason?: string): string =>
    reason ? `${label}: ${reason}` : label;
  const opacityAccessibleLabel = commandLabel(
    labels.pictureTransparency,
    opacityTitle,
  );

  const wps = layout === "wps";
  const statusCaption = model.status ? (
    // Office puts no identity text in the band (the selection pane and
    // status bar carry it); only the last message of a command stays, as a
    // short live caption beside the commands it concerns.
    <PdfSurfaceRibbonCaption
      role="status"
      testId="pdf-picture-context-status"
      title={`${labels.pictureNumber(model.order)} · ${labels.pageNumber(model.page)} · ${model.status}`}
    >
      {model.status}
    </PdfSurfaceRibbonCaption>
  ) : null;
  const insertCommand = onInsertPicture ? (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.INSERT"
      disabled={mutationDisabled}
      disabledReason={mutationTitle}
      icon={<OfficeSemanticIcon semanticId="image.insert" size={24} />}
      label={labels.insertPicture}
      onClick={onInsertPicture}
      testId="pdf-context-image-insert"
    />
  ) : null;
  const changeCommand = (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.CHANGE"
      disabled={mutationDisabled}
      disabledReason={mutationTitle}
      icon={<OfficeSemanticIcon semanticId="image.change" size={24} />}
      label={labels.changePicture}
      onClick={onChangePicture}
      testId="pdf-context-image-replace"
    />
  );
  const extractLabel = wps ? labels.extractPicture : labels.extractOriginal;
  const extractCommand = onExtractOriginal ? (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.EXTRACT_ORIGINAL"
      disabled={!model.extractOriginalEnabled}
      // The capability's own sentence is the whole name while the
      // command is unavailable (the established contract of this one
      // command); the caption stays the bare label.
      displayLabel={extractLabel}
      icon={<ImageDown size={24} />}
      label={
        model.extractOriginalEnabled
          ? extractLabel
          : (model.extractOriginalDisabledReason ?? extractLabel)
      }
      disabledReasonAttribute={model.extractOriginalDisabledReason}
      onClick={() => {
        if (!model.extractOriginalEnabled) return;
        onExtractOriginal();
      }}
      testId="pdf-context-image-extract-original"
    />
  ) : null;
  const deleteLabel = wps ? labels.wpsDelete : labels.deletePicture;
  const deleteCommand = (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.DELETE"
      disabled={mutationDisabled}
      disabledReason={mutationTitle}
      icon={<OfficeSemanticIcon semanticId="table.delete" size={24} />}
      label={deleteLabel}
      onClick={onDeletePicture}
      testId="pdf-context-image-delete"
    />
  );
  const cropCommand = (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.CROP"
      disabled={cropDisabled}
      disabledReason={cropTitle}
      icon={<OfficeSemanticIcon semanticId="image.crop" size={24} />}
      label={labels.cropPicture}
      onClick={onCropPicture}
      testId="pdf-context-image-crop"
    />
  );
  const sharpenDisabled =
    mutationDisabled || Boolean(model.sharpenDisabledReason);
  const sharpenCommand = onSharpenPicture ? (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.SHARPEN"
      disabled={sharpenDisabled}
      disabledReason={model.sharpenDisabledReason ?? mutationTitle}
      icon={<OfficeSemanticIcon semanticId="image.enhance" size={24} />}
      label={labels.sharpen}
      onClick={onSharpenPicture}
      testId="pdf-context-image-sharpen"
    />
  ) : null;
  const resetCommand = onResetPicture ? (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.RESET"
      disabled={resetDisabled}
      disabledReason={resetTitle}
      icon={<OfficeSemanticIcon semanticId="image.reset" size={24} />}
      label={labels.resetPicture}
      onClick={onResetPicture}
      testId="pdf-context-image-restore"
    />
  ) : null;
  const rotateLeftLabel = wps ? labels.wpsRotateLeft : labels.rotateLeft;
  const rotateRightLabel = wps ? labels.wpsRotateRight : labels.rotateRight;
  const rotateStack = (
    <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
      <OfficeHomeCommand
        activation="click"
        controlId="PDFCTL.CONTEXT.IMAGE.ROTATE_LEFT"
        disabled={rotateDisabled}
        icon={
          <OfficeSemanticIcon
            className={styles.mirror}
            semanticId="object.rotate"
            size={20}
          />
        }
        label={commandLabel(rotateLeftLabel, rotateTitle)}
        displayLabel={rotateLeftLabel}
        onClick={() => onRotatePicture("left")}
        testId="pdf-context-image-rotate-left"
      />
      <OfficeHomeCommand
        activation="click"
        controlId="PDFCTL.CONTEXT.IMAGE.ROTATE_RIGHT"
        disabled={rotateDisabled}
        icon={<OfficeSemanticIcon semanticId="object.rotate" size={20} />}
        label={commandLabel(rotateRightLabel, rotateTitle)}
        displayLabel={rotateRightLabel}
        onClick={() => onRotatePicture("right")}
        testId="pdf-context-image-rotate-right"
      />
    </OfficeHomeCommandStack>
  );
  const flipStack = (
    <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
      <OfficeHomeCommand
        activation="click"
        controlId="PDFCTL.CONTEXT.IMAGE.FLIP_HORIZONTAL"
        disabled={flipDisabled}
        icon={
          <OfficeSemanticIcon semanticId="object.flipHorizontal" size={20} />
        }
        label={commandLabel(labels.flipHorizontal, flipTitle)}
        displayLabel={labels.flipHorizontal}
        onClick={() => onFlipPicture("horizontal")}
        testId="pdf-context-image-flip-horizontal"
      />
      <OfficeHomeCommand
        activation="click"
        controlId="PDFCTL.CONTEXT.IMAGE.FLIP_VERTICAL"
        disabled={flipDisabled}
        icon={
          <OfficeSemanticIcon semanticId="object.flipVertical" size={20} />
        }
        label={commandLabel(labels.flipVertical, flipTitle)}
        displayLabel={labels.flipVertical}
        onClick={() => onFlipPicture("vertical")}
        testId="pdf-context-image-flip-vertical"
      />
    </OfficeHomeCommandStack>
  );
  // WPS's 翻转▾: one large command whose menu holds the two mirrors.
  const flipMenu = (
    <DropdownMenu modal={false}>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          controlId="PDFCTL.CONTEXT.IMAGE.FLIP"
          disabled={flipDisabled}
          disabledReason={flipTitle}
          displayLabel={labels.flip}
          hasMenuIndicator
          icon={<OfficeSemanticIcon semanticId="object.flipHorizontal" size={24} />}
          label={labels.flip}
          testId="pdf-context-image-flip"
          variant="large"
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent
        align="start"
        aria-label={labels.flip}
        data-testid="pdf-context-image-flip-menu"
      >
        <DropdownMenuItem
          data-testid="pdf-context-image-flip-horizontal"
          onSelect={() => onFlipPicture("horizontal")}
        >
          <OfficeSemanticIcon semanticId="object.flipHorizontal" size={16} />
          <span>{labels.flipHorizontal}</span>
        </DropdownMenuItem>
        <DropdownMenuItem
          data-testid="pdf-context-image-flip-vertical"
          onSelect={() => onFlipPicture("vertical")}
        >
          <OfficeSemanticIcon semanticId="object.flipVertical" size={16} />
          <span>{labels.flipVertical}</span>
        </DropdownMenuItem>
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
  // 透明度: the shared picture transparency menu (the deck's and the sheet's
  // picture tabs), its seven presets the writer's constant-alpha steps; the
  // trigger names the current value. No "Transparency Options..." foot: the
  
  const transparencyPercent = Math.round((1 - model.opacity) * 100);
  const transparencyRow = (
    <PdfSurfaceRibbonFieldRow testId="pdf-picture-context-transparency-row">
      <OfficePictureTransparencyMenu
        trigger={
          <OfficeHomeMenuTrigger
            disabled={opacityDisabled}
            disabledReason={opacityDisabled ? opacityAccessibleLabel : undefined}
            displayLabel={`${labels.pictureTransparency} ${transparencyPercent}%`}
            hasMenuIndicator
            icon={<OfficeSemanticIcon semanticId="image.transparency" size={16} />}
            label={opacityAccessibleLabel}
            testId="pdf-context-image-transparency"
          />
        }
        transparency={transparencyPercent}
        disabled={opacityDisabled}
        disabledReason={opacityDisabled ? opacityAccessibleLabel : undefined}
        onCommit={(transparency) => {
          if (opacityDisabled) return;
          onOpacityChange((100 - transparency) / 100);
        }}
        testIdPrefix="pdf-context-image-transparency"
      />
    </PdfSurfaceRibbonFieldRow>
  );
  const arrangeGroup = onArrangePicture ? (
    // Office's Arrange group: "Bring Forward" and "Send Backward" split
    // buttons whose menus hold the one-step and the all-the-way move. A
    // picture's layer is its place in the page content stream.
    <PdfSurfaceRibbonGroup
      label={labels.arrange}
      tab="picture-format"
      testId="pdf-picture-context-arrange-group"
    >
      <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
        {(
          [
            ["layer", "forward", "front", labels.bringForward, labels.bringToFront, "object.bringForward", "object.bringToFront"],
            ["send", "backward", "back", labels.sendBackward, labels.sendToBack, "object.sendBackward", "object.sendToBack"],
          ] as const
        ).map(([key, stepDirection, endDirection, stepLabel, endLabel, stepIcon, endIcon]) => {
          const stepId = stepDirection === "forward" ? "bring-forward" : "send-backward";
          const endId = endDirection === "front" ? "bring-to-front" : "send-to-back";
          const disabled = mutationDisabled || model.arrangeEnabled === false;
          const reason = model.arrangeDisabledReason ?? mutationTitle;
          const controlId = `PDFCTL.CONTEXT.IMAGE.${stepId.toUpperCase().replace(/-/g, "_")}`;
          return (
            <OfficeHomeSplitLayout
              key={key}
              testId={`pdf-context-image-${key}-split`}
              primary={
                <OfficeHomeCommand
                  activation="click"
                  controlId={controlId}
                  disabled={disabled}
                  disabledReason={reason}
                  displayLabel={stepLabel}
                  icon={<OfficeSemanticIcon semanticId={stepIcon} size={18} />}
                  label={stepLabel}
                  onClick={() => onArrangePicture(stepDirection)}
                  testId={`pdf-context-image-${stepId}`}
                  variant="compact"
                />
              }
              trigger={
                <DropdownMenu modal={false}>
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      controlId={controlId}
                      disabled={disabled}
                      disabledReason={reason}
                      icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                      label={stepLabel}
                      testId={`pdf-context-image-${key}-menu`}
                      variant="icon"
                    />
                  </DropdownMenuTrigger>
                  <RibbonDropdownMenuContent
                    align="start"
                    aria-label={stepLabel}
                    data-testid={`pdf-context-image-${key}-menu-content`}
                  >
                    <DropdownMenuItem
                      data-testid={`pdf-context-image-${stepId}-item`}
                      onSelect={() => onArrangePicture(stepDirection)}
                    >
                      <OfficeSemanticIcon semanticId={stepIcon} size={16} />
                      <span>{stepLabel}</span>
                    </DropdownMenuItem>
                    <DropdownMenuItem
                      data-testid={`pdf-context-image-${endId}`}
                      onSelect={() => onArrangePicture(endDirection)}
                    >
                      <OfficeSemanticIcon semanticId={endIcon} size={16} />
                      <span>{endLabel}</span>
                    </DropdownMenuItem>
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
              }
            />
          );
        })}
      </OfficeHomeCommandStack>
    </PdfSurfaceRibbonGroup>
  ) : null;
  const pointSize =
    onResizePicturePoints &&
    Number.isFinite(model.widthPoints) &&
    Number.isFinite(model.heightPoints);
  const sizeGroup = (
    <PdfSurfaceRibbonGroup
      label={labels.size}
      tab="picture-format"
      testId="pdf-picture-context-size-group"
    >
      {pointSize ? (
        <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
          <PictureSizeInput
            icon={<OfficeSemanticIcon semanticId="object.width" size={16} />}
            label={labels.pictureWidth}
            value={model.widthPoints!}
            unit={labels.points}
            disabled={mutationDisabled}
            testId="pdf-context-image-width"
            onCommit={(width) => onResizePicturePoints!(width, model.heightPoints!)}
          />
          <PictureSizeInput
            icon={<OfficeSemanticIcon semanticId="object.height" size={16} />}
            label={labels.pictureHeight}
            value={model.heightPoints!}
            unit={labels.points}
            disabled={mutationDisabled}
            testId="pdf-context-image-height"
            onCommit={(height) => onResizePicturePoints!(model.widthPoints!, height)}
          />
        </OfficeHomeCommandStack>
      ) : onResizePicture ? (
        <OfficeHomeCommandStack className={ribbonLayout.keepsWidth}>
          <PictureSizeInput
            icon={<OfficeSemanticIcon semanticId="object.width" size={16} />}
            label={labels.pictureWidth}
            value={model.width}
            disabled={mutationDisabled}
            testId="pdf-context-image-width"
            onCommit={(width) => onResizePicture(width, model.height)}
          />
          <PictureSizeInput
            icon={<OfficeSemanticIcon semanticId="object.height" size={16} />}
            label={labels.pictureHeight}
            value={model.height}
            disabled={mutationDisabled}
            testId="pdf-context-image-height"
            onCommit={(height) => onResizePicture(model.width, height)}
          />
        </OfficeHomeCommandStack>
      ) : null}
      {model.rotated && !wps ? (
        <PdfSurfaceRibbonCaption
          testId="pdf-picture-context-rotated"
          title={labels.rotatedPicture}
        >
          {labels.rotatedPicture}
        </PdfSurfaceRibbonCaption>
      ) : null}
    </PdfSurfaceRibbonGroup>
  );
  const exitCommand = (
    <PdfSurfaceRibbonCommand
      controlId="PDFCTL.CONTEXT.IMAGE.EXIT"
      disabled={model.busy}
      icon={<OfficeSemanticIcon semanticId="edit.exit" size={24} />}
      label={wps ? labels.exitEditing : labels.backToEditing}
      onClick={onExit}
      testId="pdf-context-image-exit"
    />
  );
  const adjustDataAttributes =
    process.env.NODE_ENV !== "production"
      ? {
          // Why a transform command is disabled, for the headless
          // walkthrough: the model's own flags, never user content.
          "data-picture-model": JSON.stringify({
            busy: model.busy,
            writable: model.writable,
            disabledReason: model.disabledReason ?? null,
            rotated: model.rotated,
            editKind: model.editKind ?? null,
            hasEdit: model.hasEdit,
            cropSupported: model.cropSupported ?? null,
            rotationSupported: model.rotationSupported ?? null,
            rotationBlockedByPage: model.rotationBlockedByPage ?? null,
          }),
        }
      : undefined;

  if (wps) {
    // WPS 6.8.2 图片编辑, in its order (wps/tabs/image-edit.png); the picture
    // is edited in place on the reading canvas, every command a native write.
    return (
      <>
        <span
          className={styles.surface}
          data-save-disposition="native"
          data-picture-ribbon-layout="wps"
          data-testid="pdf-picture-format-context"
          hidden
        />
        {pointerTools}
        <PdfSurfaceRibbonGroup
          label={labels.picture}
          tab="picture-format"
          testId="pdf-picture-context-picture-group"
        >
          {statusCaption}
          {insertCommand}
          {flipMenu}
          {rotateStack}
          {changeCommand}
          {deleteCommand}
        </PdfSurfaceRibbonGroup>
        <PdfSurfaceRibbonGroup
          dataAttributes={adjustDataAttributes}
          label={labels.adjust}
          tab="picture-format"
          testId="pdf-picture-context-adjust-group"
        >
          {transparencyRow}
          {sharpenCommand}
          {cropCommand}
        </PdfSurfaceRibbonGroup>
        {extractCommand ? (
          <PdfSurfaceRibbonGroup
            label={labels.extract}
            tab="picture-format"
            testId="pdf-picture-context-extract-group"
          >
            {extractCommand}
          </PdfSurfaceRibbonGroup>
        ) : null}
        {arrangeGroup}
        {sizeGroup}
        {/* WPS: every contextual tab ends with 退出编辑. */}
        <PdfSurfaceRibbonGroup
          label={labels.actions}
          tab="picture-format"
          testId="pdf-picture-context-actions-group"
        >
          {exitCommand}
        </PdfSurfaceRibbonGroup>
      </>
    );
  }

  // The groups are the ribbon track's own children, never wrapped: the
  // track's column, divider and caption rules address its direct children,
  // and a `display: contents` wrapper (the previous surface element) took the
  // groups out of their reach, leaving them uncaptioned and laid out as
  // rows. The surface's identity and save disposition live on a marker.
  return (
    <>
      <span
        className={styles.surface}
        data-save-disposition="flatten-with-consent"
        data-testid="pdf-picture-format-context"
        hidden
      />
      {pointerTools}
      {/* The same anatomy as the viewer's Edit and Insert tabs, in the group
          order of the reference suites' picture-editing tab: the picture
          itself, its adjustments, its order on the page, its size, then the
          way out. Frequent commands are large; pairs share a compact stack. */}
      <PdfSurfaceRibbonGroup
        label={labels.picture}
        tab="picture-format"
        testId="pdf-picture-context-picture-group"
      >
        {statusCaption}
        {insertCommand}
        {changeCommand}
        {extractCommand}
        {deleteCommand}
      </PdfSurfaceRibbonGroup>

      <PdfSurfaceRibbonGroup
        dataAttributes={adjustDataAttributes}
        label={labels.adjust}
        tab="picture-format"
        testId="pdf-picture-context-adjust-group"
      >
        {cropCommand}
        {resetCommand}
        {rotateStack}
        {flipStack}
        {transparencyRow}
      </PdfSurfaceRibbonGroup>

      {arrangeGroup}

      {sizeGroup}

      <PdfSurfaceRibbonGroup
        label={labels.actions}
        tab="picture-format"
        testId="pdf-picture-context-actions-group"
      >
        {exitCommand}
      </PdfSurfaceRibbonGroup>
    </>
  );
}

/** The transparency presets WPS lists, as opacities (0% .. 95%). */

/**
 * One size field: a page-relative percentage, or points when `unit` names
 * them (the reader's WPS layout). The value is committed on Enter or blur,
 * never per keystroke, so a native rewrite happens once per entry.
 */
function PictureSizeInput({
  icon,
  label,
  value,
  unit,
  disabled,
  testId,
  onCommit,
}: {
  icon?: ReactNode;
  label: string;
  value: number;
  /** Points: the value is shown and committed as typed, 1..14400 pt. */
  unit?: string;
  disabled: boolean;
  testId: string;
  onCommit: (value: number) => void;
}) {
  const points = unit !== undefined;
  const display = points ? value.toFixed(1) : (value * 100).toFixed(1);
  const [draft, setDraft] = useState(display);
  const [editing, setEditing] = useState(false);
  // Escape blurs the field to leave it; that blur must not commit the draft
  // the user just cancelled. The state reset has not rendered yet when the
  
  const cancelledRef = useRef(false);
  useEffect(() => {
    if (!editing) setDraft(display);
  }, [display, editing]);
  const commit = () => {
    setEditing(false);
    if (cancelledRef.current) {
      cancelledRef.current = false;
      setDraft(display);
      return;
    }
    const typed = Number(draft);
    if (points) {
      if (!Number.isFinite(typed) || typed < 1 || typed > 14_400) {
        setDraft(display);
        return;
      }
      const rounded = Math.round(typed * 10) / 10;
      if (Math.abs(rounded - value) < 0.05) {
        setDraft(display);
        return;
      }
      onCommit(rounded);
      return;
    }
    if (!Number.isFinite(typed) || typed < 0.5 || typed > 100) {
      setDraft(display);
      return;
    }
    const fraction = Math.round(typed * 10) / 1000;
    if (Math.abs(fraction - value) < 0.0005) {
      setDraft(display);
      return;
    }
    onCommit(fraction);
  };
  // Up and Down step the size, as a spin button's keys do (W3C spinbutton;
  // Word steps a picture's size the same way): from what the field shows,
  // a point or a tenth of a percent at a time, within the field's bounds,
  
  const nudge = (direction: 1 | -1) => {
    const shown = Number(editing ? draft : display);
    const base = Number.isFinite(shown) ? shown : Number(display);
    const step = points ? 1 : 0.1;
    const next = Math.round((base + direction * step) * 10) / 10;
    const bounded = points ? Math.min(14_400, Math.max(1, next)) : Math.min(100, Math.max(0.5, next));
    if (Math.abs(bounded - base) < 0.05) return;
    setDraft(bounded.toFixed(1));
    if (points) onCommit(bounded);
    else onCommit(Math.round(bounded * 10) / 1000);
  };
  // The shared Office number field (Docs' font size, the deck's timing
  // fields): a 30px bounded shell; the page-relative percentage is the
  // unit, stated after the field like Word's "cm".
  return (
    <PdfSurfaceRibbonFieldRow className={styles.sizeField}>
      {icon}
      <span>{label}</span>
      <OfficeHomeNumberField
        className={styles.sizeInput}
        disabled={disabled}
        inputMode="decimal"
        label={label}
        max={points ? 14_400 : 100}
        min={points ? 1 : 0.5}
        onBlur={commit}
        onDecrement={() => nudge(-1)}
        onIncrement={() => nudge(1)}
        onChange={(event) => setDraft(event.target.value)}
        onFocus={() => {
          cancelledRef.current = false;
          setEditing(true);
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            (event.currentTarget as HTMLInputElement).blur();
          } else if (event.key === "Escape") {
            cancelledRef.current = true;
            setDraft(display);
            setEditing(false);
            (event.currentTarget as HTMLInputElement).blur();
          }
        }}
        step={points ? 1 : 0.1}
        testId={testId}
        value={draft}
      />
      <span>{points ? unit : "%"}</span>
    </PdfSurfaceRibbonFieldRow>
  );
}
