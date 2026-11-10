"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Picture Format contextual tab of the presentation Ribbon (Adjust,
 * Picture Styles, Arrange, Crop, Size, Properties), drawn from the window
 * bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { lineWeightDesignPxToPoints, lineWeightPointsToDesignPx } from "@mrnote/core/web/lib/presentation/text-units.ts";
import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeSplitLayout,
  OfficeHomeCommandStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { OfficePictureCorrectionsMenu, OfficePictureColorMenu, OfficePictureTransparencyMenu } from "@mrnote/core/web/components/office-ribbon/OfficePictureAdjustMenus.tsx";
import {
  imageCorrectionDeltasFromPersisted,
  isNativeImageRecolorMode,
  imageCorrectionStoredFromDelta,
  imageTransparencyFromOpacity,
  imageOpacityFromTransparency,
  normalizeImageBorderColor,
  normalizeImageBorderWidth,
  type ImageBorderStyle,
} from "@mrnote/core/web/lib/presentation-image-effects.ts";
import {
  DropdownMenuItem,
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuLabel,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficePictureMenuIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficePictureMenuIcon.tsx";
import { OfficeArtisticEffectsMenu } from "@mrnote/core/web/components/office-ribbon/OfficeArtisticEffectsMenu.tsx";
import { normalizeOfficeArtisticEffect, withOfficeArtisticEffect } from "@mrnote/core/web/lib/office-artistic-effects.ts";
import { pictureEffectCanReturnToOriginal } from "../../../../lib/presentation-picture-original";
import { canEditBlockContent, isLockedFor } from "@mrnote/core/web/lib/presentation-text-case.ts";
import type { PresentationImageBlock, PresentationBlock, PresentationEmbedBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { RibbonDropdownMenuContent, RibbonDropdownMenuSubContent as DropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { imageHasResettableFormatting } from "@mrnote/core/web/lib/presentation-conveniences.ts";
import { PresentationPictureStyleGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationPictureStyleGallery.tsx";
import {
  canApplyNativePictureStyle,
  presentationPictureStyleBorder,
} from "@mrnote/core/web/lib/presentation-picture-style-native.ts";
import {
  PRESENTATION_PICTURE_BORDER_PRIMARY_CONTROL_ID,
  PRESENTATION_PICTURE_BORDER_MENU_CONTROL_ID,
  PRESENTATION_PICTURE_SIZE_HEIGHT_FIELD_CONTROL_ID,
  PRESENTATION_PICTURE_SIZE_WIDTH_FIELD_CONTROL_ID,
} from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import { OfficeOutlineMenuContent } from "@mrnote/core/web/components/office-ribbon/OfficeOutlineMenuContent.tsx";
import { sampleEyeDropperColor } from "../../../../lib/presentation/native-eyedropper";
import { IconChangeShape, IconPictureCrop } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { PresentationCameoShapeGallery } from "./PresentationCameo";
import { pictureClipShapeKind, withPictureCropShape } from "../../../../lib/presentation-picture-crop-shape";
import { presentationShapeKindMessageKey } from "../../../../lib/presentation/ribbon-tabs";
import { OfficePictureCropAspectRatioMenu } from "@mrnote/core/web/components/office-ribbon/OfficePictureCropAspectRatioMenu.tsx";
import { isIdentityCrop, type PictureCropAspectRatioId } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import {
  presentationPictureSizeDisplayCm,
  type PresentationPictureBorderCommandPayload,
  type PresentationPictureCropCommandAdapter,
  type PresentationPictureSizeAxis,
} from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { MAX_TABLE_FRAME_CM } from "@mrnote/core/web/lib/presentation-table-size.ts";
import { PresentationAnimatedBackgroundGroup } from "./PresentationAnimatedBackground";
import { blocksWithAnimatedBackground } from "@mrnote/core/web/lib/presentation-animated-background.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { PresentationMeasurementUnit } from "@mrnote/core/web/lib/presentation-measurement.ts";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationPictureAltTextCommandAuthority, PresentationPictureCropFeatureId } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { PresentationPictureFrameStyle } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationPictureStyleGallery.tsx";
import type { JSX } from "react/jsx-runtime";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationPictureFormatRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  sizeUnit: PresentationMeasurementUnit;
  sizeUnitLabel: string;
  sizeStep: number;
  sizeBound: (centimetres: number) => number;
  sizeDisplayFromEmu: (emu: unknown, cm: string) => string;
  pictureReplaceMenuTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureReplaceBusy: boolean;
  pictureCropTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureCropMenuTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureBorderMenuTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureSizeDraft: { blockId: string; width: string; height: string; editing: PresentationPictureSizeAxis | null; } | null;
  pictureBorderRememberedRef: RefObject<{ color: string; width: number; style: Exclude<ImageBorderStyle, "none">; }>;
  pictureBorderRecentColors: string[];
  setPictureBorderRecentColors: Dispatch<SetStateAction<string[]>>;
  pictureCompressionTriggerRef: RefObject<HTMLButtonElement | null>;
  formatPaneOpen: boolean;
  cropEditingId: string | null;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  activeSlideBounds: { width: number; height: number; };
  eyeDropperAvailable: boolean;
  themeAccentColors: string[];
  contextualImageBlock: PresentationImageBlock | null;
  contextualPictureSourcePreserveOnly: boolean;
  contextualPictureSourceReason: string | undefined;
  contextualImageCanResetAndSize: boolean;
  selectedPictureBorderEditable: boolean;
  selectedPictureTransparencyEditable: boolean;
  selectedPictureCorrectionsEditable: boolean;
  selectedPictureColorEditable: boolean;
  selectedPictureCompressionEditable: boolean;
  selectedPictureSizeEditable: boolean;
  pictureAltTextCommandAuthority: PresentationPictureAltTextCommandAuthority;
  pictureCropCommand: PresentationPictureCropCommandAdapter<PresentationImageBlock, SlideModel>;
  isPictureCropCommandEnabled: (featureId: PresentationPictureCropFeatureId) => boolean;
  toggleFormatPane: () => void;
  openFormatPane: () => void;
  openFormatPaneAtPicture: () => void;
  commitLiveActiveBlocks: (transform: (blocks: PresentationBlock[], slide: SlideModel) => PresentationBlock[] | null, coalesceKey?: string) => boolean;
  restorePictureOriginal: (blockId: string, effect: string | null, base?: PresentationImageBlock) => Promise<void>;
  resetSelectedImage: (blockId: string) => void;
  resetSelectedImageAndSize: (blockId: string) => boolean;
  beginReplaceImage: (blockId: string, trigger?: HTMLElement | null) => void;
  resetImageCrop: (blockId: string) => boolean;
  beginPictureAltText: (blockId: string) => void;
  commitPictureSizeDraft: (axis: PresentationPictureSizeAxis, rawOverride?: string, displayUnit?: PresentationMeasurementUnit) => void;
  stepPictureSize: (axis: PresentationPictureSizeAxis, direction: "increment" | "decrement") => void;
  invokePictureBorder: (payload: Omit<PresentationPictureBorderCommandPayload, "blockId">) => void;
  previewPictureStyle: (blockId: string, style: PresentationPictureFrameStyle | null) => void;
  commitPictureStyle: (blockId: string, style: PresentationPictureFrameStyle) => void;
  invokePictureCrop: (blockId: string, mode: "crop" | "fill" | "fit", trigger?: HTMLElement | null) => void;
  invokePictureCropAspectRatio: (blockId: string, ratioId: PictureCropAspectRatioId, trigger?: HTMLElement | null) => void;
  updateImageEffects: (blockId: string, patch: Partial<Pick<PresentationImageBlock, "shadow" | "borderColor" | "borderWidth" | "borderStyle" | "cornerRadius" | "opacity" | "recolor" | "brightness" | "contrast" | "blendMode">>) => void;
  renderPictureArrangeGroup: (kind: "image" | "graphic") => JSX.Element;
  renderPictureEffectsMenu: (block: PresentationImageBlock | PresentationEmbedBlock, options: { label: string; testId: string; controlId: string; }) => JSX.Element;
  autoEnhanceSelectedPicture: () => void;
  pictureEditTarget: PresentationImageBlock | null;
  startPictureEdit: (mode: "removal" | "erase") => Promise<void>;
  openPictureCompression: () => void;
  ribbonImageHasGeometry: boolean;
}

export function PresentationPictureFormatRibbonTab({ ctl }: { ctl: PresentationPictureFormatRibbonTabCtl }) {
  const {
    t,
    sizeUnit,
    sizeUnitLabel,
    sizeStep,
    sizeBound,
    sizeDisplayFromEmu,
    pictureReplaceMenuTriggerRef,
    pictureReplaceBusy,
    pictureCropTriggerRef,
    pictureCropMenuTriggerRef,
    pictureBorderMenuTriggerRef,
    pictureSizeDraft,
    pictureBorderRememberedRef,
    pictureBorderRecentColors,
    setPictureBorderRecentColors,
    formatPaneOpen,
    cropEditingId,
    autosave,
    activeSlideBounds,
    eyeDropperAvailable,
    themeAccentColors,
    contextualImageBlock,
    contextualPictureSourcePreserveOnly,
    contextualPictureSourceReason,
    contextualImageCanResetAndSize,
    selectedPictureBorderEditable,
    selectedPictureTransparencyEditable,
    selectedPictureCorrectionsEditable,
    selectedPictureColorEditable,
    selectedPictureSizeEditable,
    pictureAltTextCommandAuthority,
    pictureCropCommand,
    isPictureCropCommandEnabled,
    toggleFormatPane,
    openFormatPane,
    openFormatPaneAtPicture,
    commitLiveActiveBlocks,
    restorePictureOriginal,
    resetSelectedImage,
    resetSelectedImageAndSize,
    beginReplaceImage,
    resetImageCrop,
    beginPictureAltText,
    commitPictureSizeDraft,
    stepPictureSize,
    invokePictureBorder,
    previewPictureStyle,
    commitPictureStyle,
    invokePictureCrop,
    invokePictureCropAspectRatio,
    updateImageEffects,
    renderPictureArrangeGroup,
    renderPictureEffectsMenu,
    autoEnhanceSelectedPicture,
    pictureEditTarget,
    startPictureEdit,
    ribbonImageHasGeometry,
  } = ctl;
  return (
                  <>
                    {contextualImageBlock ? (
                      <>
                        {/* Picture Format > Adjust, as desktop PowerPoint lays it
                            out (PPT-A125): Remove Background, Corrections,
                            Color, Artistic Effects, Transparency, Compress
                            Pictures, Change Picture, Reset Picture. Erase (the
                            newer Edit Picture panel's) stays beside Remove
                            Background; the Format pane is the tab's own
                            button. */}
                        <OfficeHomeRibbonGroup
                          className={presentationChrome.pictureAdjustGroup}
                          label={t("presentation.image.adjustGroup")}
                          size="wide"
                          testId="presentation-ribbon-picture-adjust-group"
                        >
                          <OfficeHomeCommandRow>
                            {/* Both of these are blocked by the literal below and
                                nothing else: no cutout or mask service is asked
                                about anywhere in the app, so the reason says what
                                is true — we have not built them — rather than
                                sending the reader to look for a service to
                                authorize. `disabledReason` rather than a label
                                baked at the call site, so the reason also lands
                                in `data-disabled-reason` and the button's own
                                tooltip. */}
                            <OfficeHomeCommand
                              variant="large"
                              disabled={!pictureEditTarget}
                              disabledReason={t(
                                "presentation.image.removeBackgroundUnavailable",
                              )}
                              icon={<OfficeSemanticIcon semanticId="image.removeBackground" size={24} />}
                              label={t("presentation.image.removeBackground")}
                              onClick={() => void startPictureEdit("removal")}
                              testId="presentation-ribbon-image-remove-background"
                            />
                            <OfficePictureCorrectionsMenu
                              trigger={
                                <OfficeHomeMenuTrigger
                                  className={presentationChrome.pictureVerticalMenu}
                                  variant="large"
                                  hasMenuIndicator
                                  icon={<OfficeSemanticIcon semanticId="image.corrections" size={24} />}
                                  label={t("presentation.pictureRibbon.correctionsMenu")}
                                  data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.CORRECTIONS"
                                  testId="presentation-ribbon-image-enhance"
                                />
                              }
                              {...imageCorrectionDeltasFromPersisted(
                                contextualImageBlock.brightness,
                                contextualImageBlock.contrast,
                                contextualImageBlock.source,
                              )}
                              recolor={isNativeImageRecolorMode(contextualImageBlock.recolor) ? contextualImageBlock.recolor : "none"}
                              disabled={!selectedPictureCorrectionsEditable}
                              disabledReason={contextualPictureSourceReason}
                              sharpnessUnavailableReason={t("presentation.pictureRibbon.sharpnessUnavailable")}
                              onCommit={(brightness, contrast) =>
                                updateImageEffects(contextualImageBlock.id, {
                                  brightness: imageCorrectionStoredFromDelta(brightness),
                                  contrast: imageCorrectionStoredFromDelta(contrast),
                                })
                              }
                              extraItems={
                                <DropdownMenuItem
                                  disabled={!selectedPictureCorrectionsEditable}
                                  onSelect={autoEnhanceSelectedPicture}
                                  data-testid="presentation-ribbon-image-auto-enhance"
                                >
                                  <OfficePictureMenuIcon icon={{ kind: "auto-enhance" }} />
                                  {t("presentation.image.autoEnhance")}
                                </DropdownMenuItem>
                              }
                              optionsLabel={t("presentation.pictureRibbon.correctionsOptions")}
                              onOptions={openFormatPaneAtPicture}
                              testIdPrefix="presentation-ribbon-image-corrections"
                            />
                            <OfficeHomeCommand
                              variant="large"
                              disabled={!pictureEditTarget}
                              disabledReason={t(
                                "presentation.image.eraseUnavailable",
                              )}
                              icon={<OfficeSemanticIcon semanticId="draw.eraser.stroke" size={24} />}
                              label={t("presentation.image.erase")}
                              onClick={() => void startPictureEdit("erase")}
                              testId="presentation-ribbon-image-erase"
                            />
                            <OfficePictureColorMenu
                              trigger={
                                <OfficeHomeMenuTrigger
                                  className={presentationChrome.pictureVerticalMenu}
                                  variant="large"
                                  hasMenuIndicator
                                  icon={<OfficeSemanticIcon semanticId="image.color" size={24} />}
                                  label={t("toolbar.color")}
                                  data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.COLOR"
                                  testId="presentation-ribbon-image-color"
                                />
                              }
                              recolor={isNativeImageRecolorMode(contextualImageBlock.recolor) ? contextualImageBlock.recolor : "none"}
                              brightness={contextualImageBlock.brightness}
                              contrast={contextualImageBlock.contrast}
                              disabled={!selectedPictureColorEditable}
                              disabledReason={contextualPictureSourceReason}
                              unavailableReason={t("presentation.pictureRibbon.colorUnavailable")}
                              onCommitRecolor={(recolor) =>
                                updateImageEffects(contextualImageBlock.id, { recolor })
                              }
                              optionsLabel={t("presentation.pictureRibbon.colorOptions")}
                              onOptions={openFormatPaneAtPicture}
                              testIdPrefix="presentation-ribbon-image-color"
                            />
                        {/* Picture Format > Adjust > Artistic Effects: the shared
                            gallery; a picture whose effect came from a file is
                            Office's rendered bitmap and is re-effected from the
                            original the package keeps (`a14:imgLayer`), when it
                            keeps one. */}
                        <OfficeArtisticEffectsMenu
                          value={normalizeOfficeArtisticEffect(contextualImageBlock.artisticEffect)}
                          disabled={
                            contextualPictureSourcePreserveOnly ||
                            (contextualImageBlock.artisticEffectRendered === true &&
                              !pictureEffectCanReturnToOriginal(contextualImageBlock)) ||
                            !canEditBlockContent(contextualImageBlock)
                          }
                          disabledReason={
                            contextualImageBlock.artisticEffectRendered === true
                              ? t("presentation.image.artisticEffectFromFile")
                              : contextualPictureSourceReason
                          }
                          onCommit={(id) => {
                            const blockId = contextualImageBlock.id;
                            if (pictureEffectCanReturnToOriginal(contextualImageBlock)) {
                              void restorePictureOriginal(blockId, normalizeOfficeArtisticEffect(id) ?? null);
                              return;
                            }
                            commitLiveActiveBlocks((blocks) => {
                              const target = blocks.find((candidate) => candidate.id === blockId);
                              if (!target || target.type !== "image" || !canEditBlockContent(target)) return null;
                              // The shared gallery pick, which the AI's
                              // presentation.picture.artistic_effect runs too.
                              const next = withOfficeArtisticEffect(
                                target as PresentationImageBlock,
                                normalizeOfficeArtisticEffect(id) ?? "none",
                              );
                              if (!next) return null;
                              return blocks.map((candidate) => (candidate === target ? next : candidate));
                            });
                          }}
                          testIdPrefix="presentation-ribbon-image-artistic-effects"
                          options={{
                            label: t("presentation.pictureRibbon.artisticOptions"),
                            onSelect: openFormatPane,
                          }}
                          trigger={
                            <OfficeHomeMenuTrigger
                              className={presentationChrome.pictureVerticalMenu}
                              variant="large"
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="image.corrections" size={20} />}
                              label={t("pictureFormatPane.artisticEffects")}
                              testId="presentation-ribbon-image-artistic-effects"
                            />
                          }
                        />
                        <OfficePictureTransparencyMenu
                          trigger={
                            <OfficeHomeMenuTrigger
                              className={presentationChrome.pictureVerticalMenu}
                              variant="large"
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="image.transparency" size={24} />}
                              label={t("toolbar.transparency")}
                              data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.TRANSPARENCY"
                              testId="presentation-ribbon-image-transparency"
                            />
                          }
                          transparency={imageTransparencyFromOpacity(contextualImageBlock.opacity)}
                          disabled={!selectedPictureTransparencyEditable}
                          disabledReason={contextualPictureSourceReason}
                          onCommit={(transparency) =>
                            updateImageEffects(contextualImageBlock.id, {
                              opacity: imageOpacityFromTransparency(transparency),
                            })
                          }
                          optionsLabel={t("presentation.pictureRibbon.transparencyOptions")}
                          onOptions={openFormatPaneAtPicture}
                          testIdPrefix="presentation-ribbon-image-transparency"
                        />
                        {/* Compress Pictures is a desktop command: PowerPoint
                            for the web has none (Microsoft Support, "Reduce
                            the file size of a picture": not supported in
                            Microsoft 365 for the web), and the Web scope the
                            owner set keeps the deck's picture tab to the web's
                            commands (PPT-B428). The dialog and its writer stay
                            with the document editor. */}
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              ref={pictureReplaceMenuTriggerRef}
                              className={presentationChrome.pictureVerticalMenu}
                              variant="large"
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="image.change" size={20} />}
                              label={t("presentation.pictureRibbon.changePicture")}
                              disabled={
                                pictureReplaceBusy ||
                                contextualPictureSourcePreserveOnly
                              }
                              disabledReason={contextualPictureSourceReason}
                              aria-busy={pictureReplaceBusy}
                              data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.CHANGE_PICTURE"
                              testId="presentation-ribbon-image-replace"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-image-replace-menu"
                          >
                            <DropdownMenuLabel>
                              {t("presentation.pictureRibbon.changePicture")}
                            </DropdownMenuLabel>
                            <DropdownMenuItem
                              disabled={
                                pictureReplaceBusy ||
                                contextualPictureSourcePreserveOnly
                              }
                              title={contextualPictureSourceReason}
                              onSelect={() =>
                                beginReplaceImage(
                                  contextualImageBlock.id,
                                  pictureReplaceMenuTriggerRef.current,
                                )
                              }
                              data-testid="presentation-ribbon-image-replace-from-device"
                            >
                              <OfficePictureMenuIcon icon={{ kind: "picture-from-device" }} />
                              {t("toolbar.pictureFromDevice")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            className={presentationChrome.pictureVerticalSplit}
                            testId="presentation-ribbon-image-reset-split"
                            primary={
                              <OfficeHomeCommand
                                variant="large"
                                controlId="PPTCTL.CONTEXT.PICTURE.ADJUST.RESET.PRIMARY"
                                icon={<OfficeSemanticIcon semanticId="image.reset" size={20} />}
                                label={
                                  contextualPictureSourceReason
                                    ? `${t("presentation.pictureRibbon.resetPicture")}: ${contextualPictureSourceReason}`
                                    : t("presentation.pictureRibbon.resetPicture")
                                }
                                displayLabel={t("presentation.pictureRibbon.resetPicture")}
                                disabled={
                                  contextualPictureSourcePreserveOnly ||
                                  (!imageHasResettableFormatting(
                                    contextualImageBlock,
                                  ) &&
                                    !pictureEffectCanReturnToOriginal(contextualImageBlock))
                                }
                                onClick={() =>
                                  resetSelectedImage(contextualImageBlock.id)
                                }
                                testId="presentation-ribbon-image-reset"
                              />
                            }
                            trigger={
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  variant="icon"
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId="menu.chevron.down"
                                      size={12}
                                    />
                                  }
                                  label={t("presentation.pictureRibbon.resetPicture")}
                                  disabled={
                                    contextualPictureSourcePreserveOnly ||
                                    (!imageHasResettableFormatting(
                                      contextualImageBlock,
                                    ) &&
                                      !contextualImageCanResetAndSize)
                                  }
                                  disabledReason={contextualPictureSourceReason}
                                  data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.RESET.MENU"
                                  testId="presentation-ribbon-image-reset-menu"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-image-reset-menu-content"
                          >
                            <DropdownMenuItem
                              disabled={
                                contextualPictureSourcePreserveOnly ||
                                !imageHasResettableFormatting(
                                  contextualImageBlock,
                                )
                              }
                              title={contextualPictureSourceReason}
                              onSelect={() =>
                                resetSelectedImage(contextualImageBlock.id)
                              }
                              data-testid="presentation-ribbon-image-reset-menu-picture"
                            >
                              <OfficePictureMenuIcon icon={{ kind: "picture-reset" }} />
                              {t("presentation.pictureRibbon.resetPicture")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                contextualPictureSourcePreserveOnly ||
                                !contextualImageCanResetAndSize
                              }
                              title={contextualPictureSourceReason}
                              onSelect={() =>
                                resetSelectedImageAndSize(
                                  contextualImageBlock.id,
                                )
                              }
                              data-control-id="PPTCTL.CONTEXT.PICTURE.ADJUST.RESET.PICTURE_AND_SIZE"
                              data-testid="presentation-ribbon-image-reset-menu-picture-and-size"
                            >
                              <OfficePictureMenuIcon icon={{ kind: "picture-reset-size" }} />
                              {t("presentation.image.resetAndSize")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.pictureStylesGroup}
                      controlId="PPTCTL.CONTEXT.PICTURE.STYLE.GALLERY"
                      label={t("presentation.image.stylesGroup")}
                      size="gallery"
                      testId="presentation-ribbon-picture-styles-group"
                    >
                      <OfficeHomeCommandRow>
                        <PresentationPictureStyleGallery
                          accentColors={themeAccentColors}
                          disabled={
                            !canEditBlockContent(contextualImageBlock) ||
                            !canApplyNativePictureStyle(contextualImageBlock)
                          }
                          disabledReason={
                            !canApplyNativePictureStyle(contextualImageBlock)
                              ? t("presentation.image.nativeStylePreserved")
                              : undefined
                          }
                          image={contextualImageBlock}
                          labels={{
                            ariaLabel: t("toolbar.ribbonStyles"),
                            more: t("toolbar.quickOverflow"),
                          }}
                          onCommit={(style) =>
                            commitPictureStyle(contextualImageBlock.id, style)
                          }
                          onPreviewChange={(style) =>
                            previewPictureStyle(contextualImageBlock.id, style)
                          }
                        />
                        <div className={presentationChrome.pictureStyleTools}>
                      <div className="note-editor-ribbon-row">
                        {(() => {
                          const remembered = pictureBorderRememberedRef.current;
                          // A styled picture's current line is the style's
                          
                          const styled = presentationPictureStyleBorder(
                            contextualImageBlock,
                          );
                          const color = normalizeImageBorderColor(
                            styled
                              ? styled.color
                              : contextualImageBlock.borderColor,
                            remembered.color,
                          );
                          const currentWidth = styled
                            ? styled.widthPx
                            : normalizeImageBorderWidth(
                                contextualImageBlock.borderWidth,
                              );
                          const width =
                            currentWidth || remembered.width;
                          // A colour or dash picked over a style keeps the
                          // style's own weight, which may be wider than a
                          // Picture Border can be set to.
                          const pickWidth =
                            styled !== null &&
                            styled.widthPx > 0 &&
                            !normalizeImageBorderWidth(
                              contextualImageBlock.borderWidth,
                            )
                              ? undefined
                              : width;
                          const rawStyle = styled
                            ? styled.style
                            : contextualImageBlock.borderStyle;
                          const hasVisibleBorder =
                            currentWidth > 0 && rawStyle !== "none";
                          const style: Exclude<ImageBorderStyle, "none"> =
                            rawStyle === "dashed" || rawStyle === "dotted"
                              ? rawStyle
                              : rawStyle === "solid" || currentWidth > 0
                                ? "solid"
                                : remembered.style;
                          const disabled = !selectedPictureBorderEditable;
                          const label = t(
                            "presentation.pictureRibbon.pictureBorder",
                          );
                          const colorGlyph = (
                            <span className="note-editor-office-toolbar__color-glyph">
                              <OfficeSemanticIcon semanticId="image.border" size={16} />
                              <span
                                className="note-editor-office-toolbar__color-underline"
                                style={{ background: color }}
                              />
                            </span>
                          );
                          return (
                            <span
                              style={{ position: "relative", display: "flex" }}
                            >
                              <OfficeHomeSplitLayout
                                className={presentationChrome.pictureVerticalSplit}
                                testId="presentation-ribbon-picture-border-split"
                                primary={
                                  <OfficeHomeCommand
                                    variant="large"
                                    icon={colorGlyph}
                                    displayLabel={label}
                                    label={label}
                                    disabled={disabled}
                                    controlId={
                                      PRESENTATION_PICTURE_BORDER_PRIMARY_CONTROL_ID
                                    }
                                    onClick={() =>
                                      invokePictureBorder({
                                        color,
                                        width: pickWidth,
                                        style,
                                      })
                                    }
                                    testId="presentation-ribbon-picture-border-primary"
                                  />
                                }
                                trigger={
                                  <DropdownMenu>
                                    <DropdownMenuTrigger asChild>
                                      <OfficeHomeMenuTrigger
                                        ref={pictureBorderMenuTriggerRef}
                                        variant="icon"
                                        icon={
                                          <OfficeSemanticIcon
                                            semanticId="menu.chevron.down"
                                            size={12}
                                          />
                                        }
                                        label={label}
                                        disabled={disabled}
                                        data-control-id={
                                          PRESENTATION_PICTURE_BORDER_MENU_CONTROL_ID
                                        }
                                        testId="presentation-ribbon-picture-border-menu"
                                      />
                                    </DropdownMenuTrigger>
                                    <RibbonDropdownMenuContent
                                      align="start"
                                      data-testid="presentation-ribbon-picture-border-menu-content"
                                    >
                                      {/* PowerPoint's Picture Border: the colour
                                          palette (theme, standard, recent, No
                                          Outline, More Outline Colors), the
                                          Eyedropper, then Weight, Sketched and
                                          Dashes submenus (PPT-A128). */}
                                      <OfficeOutlineMenuContent
                                        palette={{
                                          value: hasVisibleBorder ? color : undefined,
                                          recents: pictureBorderRecentColors,
                                          allowNoFill: true,
                                          noFillSelected: !hasVisibleBorder,
                                          labels: {
                                            themeColors: t("spreadsheet.themeColors"),
                                            standardColors: t("spreadsheet.standardColors"),
                                            recentColors: t("spreadsheet.recentColors"),
                                            moreColors: t("presentation.cameo.otherBorderColors"),
                                            noFill: t("presentation.home.noOutline"),
                                          },
                                          onPick: (nextColor) => {
                                            if (!nextColor) {
                                              if (hasVisibleBorder) {
                                                invokePictureBorder({ style: "none" });
                                              }
                                              return;
                                            }
                                            setPictureBorderRecentColors((current) =>
                                              [
                                                nextColor,
                                                ...current.filter(
                                                  (candidate) => candidate !== nextColor,
                                                ),
                                              ].slice(0, 8),
                                            );
                                            invokePictureBorder({
                                              color: nextColor,
                                              width: pickWidth,
                                              style,
                                            });
                                          },
                                          testIdPrefix: "presentation-picture-border",
                                        }}
                                        eyedropper={
                                          eyeDropperAvailable
                                            ? {
                                                label: t("presentation.ui.eyedropper"),
                                                testId: "presentation-picture-border-eyedropper",
                                                onPick: () => {
                                                  void sampleEyeDropperColor().then((hex) => {
                                                    if (hex) invokePictureBorder({ color: hex, width: pickWidth, style });
                                                  });
                                                },
                                              }
                                            : undefined
                                        }
                                        weight={{
                                          label: t("presentation.home.lineWeight"),
                                          testId: "presentation-picture-border-width",
                                          valuePt: hasVisibleBorder ? lineWeightDesignPxToPoints(width) : null,
                                          onPick: (points) => {
                                            if (hasVisibleBorder && points === lineWeightDesignPxToPoints(width)) return;
                                            invokePictureBorder({
                                              color,
                                              width: lineWeightPointsToDesignPx(points),
                                              style,
                                            });
                                          },
                                        }}
                                        sketched={{
                                          label: t("office.outline.sketched"),
                                          testId: "presentation-picture-border-sketched",
                                          disabledReason: t("office.outline.sketchedUnavailable"),
                                        }}
                                        dashes={{
                                          label: t("presentation.home.dashes"),
                                          testId: "presentation-picture-border-style",
                                          value: hasVisibleBorder ? style : null,
                                          optionLabel: (nextStyle) =>
                                            t(
                                              `presentation.shapeOutlineStyle.${nextStyle}` as "presentation.shapeOutlineStyle.solid",
                                            ),
                                          onPick: (nextStyle) => {
                                            if (hasVisibleBorder && nextStyle === style) return;
                                            invokePictureBorder({ color, width: pickWidth, style: nextStyle });
                                          },
                                        }}
                                      />
                                    </RibbonDropdownMenuContent>
                                  </DropdownMenu>
                                }
                              />
                            </span>
                          );
                        })()}
                      </div>
                      {/* Picture Effects over Format Picture in one stack of
                          two compact faces beside the Border split, as
                          PowerPoint stacks Picture Effects / Picture Layout
                          beside Picture Border. Each in a row of its own,
                          the three ran side by side and the last two left
                          the group over Properties and Arrange. */}
                      <OfficeHomeCommandStack>
                        {renderPictureEffectsMenu(contextualImageBlock, {
                          label: t("office.pictureEffects.button"),
                          testId: "presentation-ribbon-picture-effects",
                          controlId: "PPTCTL.CONTEXT.PICTURE.STYLE.EFFECTS",
                        })}
                        <OfficeHomeCommand
                          active={formatPaneOpen}
                          variant="compact"
                          icon={<OfficeSemanticIcon semanticId="image.format" size={16} />}
                          displayLabel={t(
                            "presentation.image.formatPicture",
                          )}
                          label={t("presentation.image.formatPicture")}
                          onClick={toggleFormatPane}
                          testId="presentation-ribbon-image-format-pane"
                        />
                      </OfficeHomeCommandStack>
                        </div>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.picturePropertiesGroup}
                      label={t("presentation.smartartContext.propertiesGroup")}
                      size="narrow"
                      testId="presentation-ribbon-picture-properties-group"
                    >
                      <OfficeHomeCommand
                        variant="large"
                        icon={<OfficeSemanticIcon semanticId="object.altText" size={24} />}
                        label={t("presentation.pictureRibbon.altText")}
                        disabled={!pictureAltTextCommandAuthority.isEnabled()}
                        onClick={() =>
                          beginPictureAltText(contextualImageBlock.id)
                        }
                        testId="presentation-ribbon-image-alt-text"
                      />
                    </OfficeHomeRibbonGroup>
                      </>
                    ) : null}
                    {renderPictureArrangeGroup("image")}
                    {contextualImageBlock ? (
                      <OfficeHomeRibbonGroup
                        className={presentationChrome.pictureCropGroup}
                        label={t("presentation.crop.toggle")}
                          size="compact"
                          testId="presentation-ribbon-picture-crop-group"
                        >
                      <OfficeHomeCommandRow>
                        <OfficeHomeSplitLayout
                          className={presentationChrome.pictureVerticalSplit}
                          testId="presentation-ribbon-image-crop-split"
                          primary={
                            <OfficeHomeCommand
                              buttonRef={pictureCropTriggerRef}
                              variant="large"
                              active={
                                cropEditingId === contextualImageBlock.id
                              }
                              icon={<OfficeSemanticIcon semanticId="image.crop" size={20} />}
                              label={t("presentation.crop.toggle")}
                              disabled={
                                !ribbonImageHasGeometry ||
                                !isPictureCropCommandEnabled(
                                  pictureCropCommand.featureIds.crop,
                                )
                              }
                              onClick={() =>
                                invokePictureCrop(
                                  contextualImageBlock.id,
                                  "crop",
                                  pictureCropTriggerRef.current,
                                )
                              }
                              testId="presentation-ribbon-image-crop"
                            />
                          }
                          trigger={
                            <DropdownMenu>
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  ref={pictureCropMenuTriggerRef}
                                  variant="icon"
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId="menu.chevron.down"
                                      size={12}
                                    />
                                  }
                                  label={t("presentation.crop.toggle")}
                                  disabled={
                                    !ribbonImageHasGeometry ||
                                    !isPictureCropCommandEnabled(
                                      pictureCropCommand.featureIds.crop,
                                    )
                                  }
                                  testId="presentation-ribbon-image-crop-menu"
                                />
                              </DropdownMenuTrigger>
                              <RibbonDropdownMenuContent
                                align="start"
                                data-testid="presentation-ribbon-image-crop-menu-content"
                              >
                                <DropdownMenuItem
                                  disabled={
                                    !ribbonImageHasGeometry ||
                                    !isPictureCropCommandEnabled(
                                      pictureCropCommand.featureIds.crop,
                                    )
                                  }
                                  onSelect={() =>
                                    invokePictureCrop(
                                      contextualImageBlock.id,
                                      "crop",
                                      pictureCropMenuTriggerRef.current,
                                    )
                                  }
                                  data-testid="presentation-ribbon-image-crop-menu-crop"
                                >
                                  <OfficeSemanticIcon semanticId="image.crop" size={16} />
                                  {t("presentation.crop.toggle")}
                                </DropdownMenuItem>
                                {/* Crop > Crop to Shape (PPT-A132): the deck's
                                    shape gallery, lines left out as Office
                                    does; the picture keeps its frame and crop. */}
                                <DropdownMenuSub>
                                  <DropdownMenuSubTrigger
                                    disabled={
                                      contextualPictureSourcePreserveOnly ||
                                      isLockedFor(contextualImageBlock, "changeShape") ||
                                      Boolean(contextualImageBlock.pictureStylePreset)
                                    }
                                    data-testid="presentation-ribbon-image-crop-to-shape"
                                  >
                                    <IconChangeShape home size={16} />
                                    {t("presentation.cameo.cropToShape")}
                                  </DropdownMenuSubTrigger>
                                  <DropdownMenuSubContent
                                    data-testid="presentation-ribbon-image-crop-to-shape-menu"
                                    onEscapeKeyDown={(event) => event.stopPropagation()}
                                    style={{ maxHeight: "min(560px, var(--radix-dropdown-menu-content-available-height, calc(100vh - 24px)))", maxWidth: "none", overflowY: "auto" }}
                                  >
                                    <PresentationCameoShapeGallery
                                      current={pictureClipShapeKind(contextualImageBlock)}
                                      labelFor={(kind) => t(presentationShapeKindMessageKey(kind))}
                                      onSelect={(kind) => {
                                        const blockId = contextualImageBlock.id;
                                        commitLiveActiveBlocks((blocks) => {
                                          const target = blocks.find((candidate) => candidate.id === blockId);
                                          if (!target || target.type !== "image" || !canEditBlockContent(target)) return null;
                                          const next = withPictureCropShape(target as PresentationImageBlock, kind);
                                          return next ? blocks.map((candidate) => (candidate === target ? next : candidate)) : null;
                                        });
                                      }}
                                      testPrefix="presentation-ribbon-image-crop-shape"
                                      t={t}
                                    />
                                  </DropdownMenuSubContent>
                                </DropdownMenuSub>
                                  <OfficePictureCropAspectRatioMenu
                                    icon={<IconPictureCrop home size={16} />}
                                    testIdPrefix="presentation-ribbon-image-crop-aspect-ratio"
                                    disabled={
                                      !ribbonImageHasGeometry ||
                                      !pictureCropCommand.canApplyAspectRatio()
                                    }
                                    labels={{
                                      aspectRatio: t(
                                        "presentation.crop.aspectRatioLabel",
                                      ),
                                      square: t(
                                        "presentation.crop.aspectRatio.square",
                                      ),
                                      portrait: t(
                                        "presentation.crop.aspectRatio.portrait",
                                      ),
                                      landscape: t(
                                        "presentation.crop.aspectRatio.landscape",
                                      ),
                                    }}
                                    onSelect={(ratioId) =>
                                      invokePictureCropAspectRatio(
                                        contextualImageBlock.id,
                                        ratioId,
                                        pictureCropMenuTriggerRef.current,
                                      )
                                    }
                                  />
                                <DropdownMenuSeparator />
                                <DropdownMenuItem
                                  disabled={
                                    !ribbonImageHasGeometry ||
                                    !isPictureCropCommandEnabled(
                                      pictureCropCommand.featureIds.fill,
                                    )
                                  }
                                  onSelect={() =>
                                    invokePictureCrop(
                                      contextualImageBlock.id,
                                      "fill",
                                      pictureCropMenuTriggerRef.current,
                                    )
                                  }
                                  data-testid="presentation-ribbon-image-crop-menu-fill"
                                >
                                  <OfficePictureMenuIcon icon={{ kind: "crop-fill" }} />
                                  {t("presentation.crop.fill")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={
                                    !ribbonImageHasGeometry ||
                                    !isPictureCropCommandEnabled(
                                      pictureCropCommand.featureIds.fit,
                                    )
                                  }
                                  onSelect={() =>
                                    invokePictureCrop(
                                      contextualImageBlock.id,
                                      "fit",
                                      pictureCropMenuTriggerRef.current,
                                    )
                                  }
                                  data-testid="presentation-ribbon-image-crop-menu-fit"
                                >
                                  <OfficePictureMenuIcon icon={{ kind: "crop-fit" }} />
                                  {t("presentation.crop.fit")}
                                </DropdownMenuItem>
                                <DropdownMenuSeparator />
                                <DropdownMenuItem
                                  disabled={
                                    !contextualImageBlock.cropRect ||
                                    isIdentityCrop(contextualImageBlock.cropRect)
                                  }
                                  onSelect={() =>
                                    resetImageCrop(contextualImageBlock.id)
                                  }
                                  data-testid="presentation-ribbon-image-crop-reset"
                                >
                                  <OfficePictureMenuIcon icon={{ kind: "crop-reset" }} />
                                  {t("presentation.crop.reset")}
                                </DropdownMenuItem>
                              </RibbonDropdownMenuContent>
                            </DropdownMenu>
                          }
                        />
                      </OfficeHomeCommandRow>
                        </OfficeHomeRibbonGroup>
                    ) : null}
                    {contextualImageBlock ? (
                      <OfficeHomeRibbonGroup
                        className={presentationChrome.pictureSizeGroup}
                        label={t("presentation.smartartContext.sizeGroup")}
                        size="standard"
                        testId="presentation-ribbon-picture-size-group"
                      >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommandStack
                          className={presentationChrome.pictureSizeLane}
                        >
                        {(["height", "width"] as const).map((key) => {
                          const fieldControlId =
                            key === "height"
                              ? PRESENTATION_PICTURE_SIZE_HEIGHT_FIELD_CONTROL_ID
                              : PRESENTATION_PICTURE_SIZE_WIDTH_FIELD_CONTROL_ID;
                          const value =
                            pictureSizeDraft?.blockId === contextualImageBlock.id
                              ? pictureSizeDraft[key]
                              : sizeDisplayFromEmu(
                                  contextualImageBlock[key],
                                  presentationPictureSizeDisplayCm(
                                    contextualImageBlock[key],
                                  ),
                                );
                          const label = t(
                            key === "height" ? "presentation.tableRibbon.heightShort" : "presentation.tableRibbon.widthShort",
                          );
                          return (
                            <span
                              key={key}
                              className={presentationChrome.pictureSizeControl}
                              data-testid={`presentation-ribbon-image-${key}-hit-target`}
                              style={{ gap: 4, fontSize: 10 }}
                            >
                                <OfficeRibbonNumberField
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId={
                                        key === "height"
                                          ? "object.height"
                                          : "object.width"
                                      }
                                      size={18}
                                    />
                                  }
                                  visibleLabel={label}
                                  controlId={fieldControlId}
                                  testId={`presentation-ribbon-image-${key}`}
                                  label={label}
                                  unit={sizeUnitLabel}
                                  min={0.01}
                                  max={sizeBound(MAX_TABLE_FRAME_CM)}
                                  step={sizeStep}
                                  value={value}
                                  disabled={!selectedPictureSizeEditable}
                                  onStep={(direction) =>
                                    stepPictureSize(key, direction)
                                  }
                                  onCommit={(rawValue) =>
                                    commitPictureSizeDraft(key, rawValue, sizeUnit)
                                  }
                                />
                            </span>
                          );
                        })}
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                    {contextualImageBlock ? (
                      // Picture Format > Animate as Background, the tab's far
                      // right (PowerPoint for Mac; PPT-A134).
                      <PresentationAnimatedBackgroundGroup
                        block={contextualImageBlock}
                        slide={activeSlideBounds}
                        readOnlyReason={
                          autosave.readOnly || !canEditBlockContent(contextualImageBlock)
                            ? t("ai.readOnlyLock")
                            : undefined
                        }
                        onApply={(blockId, background) =>
                          // The shared writer (lib/presentation-animated-background),
                          // which the AI's presentation.picture.animated_background runs too.
                          commitLiveActiveBlocks((blocks) =>
                            blocksWithAnimatedBackground(blocks, blockId, background, canEditBlockContent),
                          )
                        }
                      />
                    ) : null}
                  </>
  );
}
