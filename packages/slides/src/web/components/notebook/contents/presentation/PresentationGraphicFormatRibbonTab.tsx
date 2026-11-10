"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Graphics Format contextual tab of the presentation Ribbon (Change,
 * Graphics Styles, Arrange, Size, Format Pane, Properties), drawn from the
 * window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeMenuTrigger,
  OfficeHomeCommand,
  OfficeHomeCommandStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuItem } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { canEditBlockContent } from "@mrnote/core/web/lib/presentation-text-case.ts";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficePictureMenuIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficePictureMenuIcon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationGraphicStyleGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationGraphicStyleGallery.tsx";
import { withOfficeGraphicStyle, type OfficeGraphicPaintField } from "@mrnote/core/web/lib/presentation-office-graphic-paint.ts";
import { presentationPictureSizeDisplayCm, type PresentationPictureCropCommandAdapter, type PresentationPictureSizeAxis } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { MAX_TABLE_FRAME_CM } from "@mrnote/core/web/lib/presentation-table-size.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { PresentationMeasurementUnit } from "@mrnote/core/web/lib/presentation-measurement.ts";
import type { RefObject } from "react";
import type { PresentationEmbedBlock, PresentationImageBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { PresentationPictureAltTextCommandAuthority, PresentationPictureCropFeatureId } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { JSX } from "react/jsx-runtime";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationGraphicFormatRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  sizeUnit: PresentationMeasurementUnit;
  sizeUnitLabel: string;
  sizeStep: number;
  sizeBound: (centimetres: number) => number;
  sizeDisplayFromEmu: (emu: unknown, cm: string) => string;
  pictureReplaceBusy: boolean;
  pictureCropTriggerRef: RefObject<HTMLButtonElement | null>;
  pictureSizeDraft: { blockId: string; width: string; height: string; editing: PresentationPictureSizeAxis | null; } | null;
  formatPaneOpen: boolean;
  cropEditingId: string | null;
  deckThemeGridColors: Readonly<Record<"accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "bg1" | "tx1" | "bg2" | "tx2", string>>;
  contextualOfficeGraphicBlock: PresentationImageBlock;
  selectedPictureSizeEditable: boolean;
  pictureAltTextCommandAuthority: PresentationPictureAltTextCommandAuthority;
  pictureCropCommand: PresentationPictureCropCommandAdapter<PresentationImageBlock, SlideModel>;
  isPictureCropCommandEnabled: (featureId: PresentationPictureCropFeatureId) => boolean;
  toggleFormatPane: () => void;
  beginReplaceImage: (blockId: string, trigger?: HTMLElement | null) => void;
  beginPictureAltText: (blockId: string) => void;
  commitPictureSizeDraft: (axis: PresentationPictureSizeAxis, rawOverride?: string, displayUnit?: PresentationMeasurementUnit) => void;
  stepPictureSize: (axis: PresentationPictureSizeAxis, direction: "increment" | "decrement") => void;
  invokePictureCrop: (blockId: string, mode: "crop" | "fill" | "fit", trigger?: HTMLElement | null) => void;
  commitOfficeGraphic: (blockId: string, edit: (block: PresentationImageBlock) => PresentationImageBlock | null) => void;
  renderOfficeGraphicPaintMenu: (block: PresentationImageBlock, field: OfficeGraphicPaintField) => JSX.Element;
  renderPictureArrangeGroup: (kind: "image" | "graphic") => JSX.Element;
  renderPictureEffectsMenu: (block: PresentationImageBlock | PresentationEmbedBlock, options: { label: string; testId: string; controlId: string; }) => JSX.Element;
  ribbonImageHasGeometry: boolean;
}

export function PresentationGraphicFormatRibbonTab({ ctl }: { ctl: PresentationGraphicFormatRibbonTabCtl }) {
  const {
    t,
    sizeUnit,
    sizeUnitLabel,
    sizeStep,
    sizeBound,
    sizeDisplayFromEmu,
    pictureReplaceBusy,
    pictureCropTriggerRef,
    pictureSizeDraft,
    formatPaneOpen,
    cropEditingId,
    deckThemeGridColors,
    contextualOfficeGraphicBlock,
    selectedPictureSizeEditable,
    pictureAltTextCommandAuthority,
    pictureCropCommand,
    isPictureCropCommandEnabled,
    toggleFormatPane,
    beginReplaceImage,
    beginPictureAltText,
    commitPictureSizeDraft,
    stepPictureSize,
    invokePictureCrop,
    commitOfficeGraphic,
    renderOfficeGraphicPaintMenu,
    renderPictureArrangeGroup,
    renderPictureEffectsMenu,
    ribbonImageHasGeometry,
  } = ctl;
  return (
                <>
                  {/* Graphics Format > Change (PPT-A161): Change Graphic and
                      Convert to Shape, PowerPoint's first group. */}
                  <OfficeHomeRibbonGroup
                    label={t("presentation.officeGraphic.changeGroup")}
                    size="narrow"
                    testId="presentation-ribbon-graphic-change-group"
                  >
                    <OfficeHomeCommandRow>
                      <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                          <OfficeHomeMenuTrigger
                            variant="large"
                            hasMenuIndicator
                            icon={<OfficeSemanticIcon semanticId="image.change" size={24} />}
                            label={t("presentation.desktopMenu.changeGraphic")}
                            disabled={pictureReplaceBusy || !canEditBlockContent(contextualOfficeGraphicBlock)}
                            data-control-id="PPTCTL.CONTEXT.GRAPHIC.CHANGE"
                            testId="presentation-ribbon-graphic-change"
                          />
                        </DropdownMenuTrigger>
                        <RibbonDropdownMenuContent align="start" data-testid="presentation-ribbon-graphic-change-menu">
                          <DropdownMenuItem
                            onSelect={() => beginReplaceImage(contextualOfficeGraphicBlock.id, null)}
                            data-testid="presentation-ribbon-graphic-change-from-device"
                          >
                            <OfficePictureMenuIcon icon={{ kind: "picture-from-device" }} />
                            {t("toolbar.pictureFromDevice")}
                          </DropdownMenuItem>
                        </RibbonDropdownMenuContent>
                      </DropdownMenu>
                      <OfficeHomeCommand
                        variant="large"
                        disabled
                        disabledReason={t("presentation.officeGraphic.convertToShapeUnavailable")}
                        icon={<OfficeSemanticIcon semanticId="shape.change" size={24} />}
                        label={t("presentation.desktopMenu.convertToShape")}
                        onClick={() => undefined}
                        testId="presentation-ribbon-graphic-convert-to-shape"
                      />
                    </OfficeHomeCommandRow>
                  </OfficeHomeRibbonGroup>
                  {/* Graphics Styles (PPT-A161): the style gallery, then
                      Graphics Fill / Graphics Outline / Graphics Effects in
                      one column, as PowerPoint stacks them. The group used
                      to borrow Picture Styles' fixed 582px budget and stood
                      mostly empty. */}
                  <OfficeHomeRibbonGroup
                    className={presentationChrome.graphicStylesGroup}
                    label={t("presentation.officeGraphic.stylesGroup")}
                    size="standard"
                    testId="presentation-ribbon-graphic-styles-group"
                  >
                    <OfficeHomeCommandRow className={presentationChrome.galleryWithCommandStack}>
                      <PresentationGraphicStyleGallery
                        block={contextualOfficeGraphicBlock}
                        themeColors={deckThemeGridColors}
                        disabled={!canEditBlockContent(contextualOfficeGraphicBlock)}
                        className={presentationChrome.mediaVideoStylesGallery}
                        onCommit={(style) =>
                          commitOfficeGraphic(contextualOfficeGraphicBlock.id, (target) =>
                            withOfficeGraphicStyle(target, style),
                          )
                        }
                      />
                      <OfficeHomeCommandStack className={presentationChrome.galleryCommandStack} rows={3}>
                        {/* Graphics Fill / Outline rewrite the graphic itself
                            (PPT-A161); Graphics Effects are its own p:spPr. */}
                        <OfficeHomeCommandRow>
                          {renderOfficeGraphicPaintMenu(contextualOfficeGraphicBlock, "graphicFill")}
                        </OfficeHomeCommandRow>
                        <OfficeHomeCommandRow>
                          {renderOfficeGraphicPaintMenu(contextualOfficeGraphicBlock, "graphicOutline")}
                        </OfficeHomeCommandRow>
                        <OfficeHomeCommandRow>
                          {renderPictureEffectsMenu(contextualOfficeGraphicBlock, {
                            label: t("presentation.officeGraphic.effects"),
                            testId: "presentation-ribbon-graphic-effects",
                            controlId: "PPTCTL.CONTEXT.GRAPHIC.STYLE.EFFECTS",
                          })}
                        </OfficeHomeCommandRow>
                      </OfficeHomeCommandStack>
                    </OfficeHomeCommandRow>
                  </OfficeHomeRibbonGroup>
                  <OfficeHomeRibbonGroup
                    className={presentationChrome.picturePropertiesGroup}
                    label={t("presentation.smartartContext.propertiesGroup")}
                    size="narrow"
                    testId="presentation-ribbon-graphic-properties-group"
                  >
                    <OfficeHomeCommand
                      variant="large"
                      icon={
                        <OfficeSemanticIcon
                          semanticId="object.altText"
                          size={24}
                        />
                      }
                      label={t("presentation.altTextButton")}
                      disabled={!pictureAltTextCommandAuthority.isEnabled()}
                      onClick={() =>
                        beginPictureAltText(contextualOfficeGraphicBlock.id)
                      }
                      testId="presentation-ribbon-graphic-alt-text"
                    />
                  </OfficeHomeRibbonGroup>
                  {/* Arrange: PowerPoint's whole group (Bring Forward,
                      Send Backward, Selection Pane, Align, Group, Rotate),
                      shared with Picture Format. */}
                  {renderPictureArrangeGroup("graphic")}
                  <OfficeHomeRibbonGroup
                    className={presentationChrome.graphicSizeGroup}
                    label={t("presentation.smartartContext.sizeGroup")}
                    size="standard"
                    testId="presentation-ribbon-graphic-size-group"
                  >
                    <OfficeHomeCommandRow>
                      {/* Size > Crop (PPT-A161), the picture's own crop
                          session over the graphic. */}
                      <OfficeHomeCommand
                        buttonRef={pictureCropTriggerRef}
                        variant="large"
                        active={cropEditingId === contextualOfficeGraphicBlock.id}
                        icon={<OfficeSemanticIcon semanticId="image.crop" size={24} />}
                        label={t("presentation.crop.toggle")}
                        disabled={
                          !ribbonImageHasGeometry ||
                          !isPictureCropCommandEnabled(pictureCropCommand.featureIds.crop)
                        }
                        onClick={() =>
                          invokePictureCrop(
                            contextualOfficeGraphicBlock.id,
                            "crop",
                            pictureCropTriggerRef.current,
                          )
                        }
                        testId="presentation-ribbon-graphic-crop"
                      />
                      <OfficeHomeCommandStack
                        className={presentationChrome.pictureSizeLane}
                      >
                        {(["height", "width"] as const).map((key) => {
                          const currentDisplay = sizeDisplayFromEmu(
                            contextualOfficeGraphicBlock[key],
                            presentationPictureSizeDisplayCm(
                              contextualOfficeGraphicBlock[key],
                            ),
                          );
                          const value =
                            pictureSizeDraft?.blockId ===
                            contextualOfficeGraphicBlock.id
                              ? pictureSizeDraft[key]
                              : currentDisplay;
                          const label = t(
                            key === "height" ? "presentation.tableRibbon.heightShort" : "presentation.tableRibbon.widthShort",
                          );
                          return (
                            <span
                              key={key}
                              className={presentationChrome.pictureSizeControl}
                              data-testid={`presentation-ribbon-graphic-${key}-hit-target`}
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
                                testId={`presentation-ribbon-graphic-${key}`}
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
                  <OfficeHomeRibbonGroup
                    label={t("presentation.media.formatPaneGroup")}
                    size="narrow"
                    testId="presentation-ribbon-graphic-format-pane-group"
                  >
                    <OfficeHomeCommandRow>
                      <OfficeHomeCommand
                        active={formatPaneOpen}
                        variant="large"
                        icon={<OfficeSemanticIcon semanticId="image.format" size={24} />}
                        label={t("presentation.media.formatPane")}
                        onClick={toggleFormatPane}
                        testId="presentation-ribbon-graphic-format-pane"
                      />
                    </OfficeHomeCommandRow>
                  </OfficeHomeRibbonGroup>
                </>
  );
}
