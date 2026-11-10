"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Camera Format contextual tab of the presentation Ribbon (Camera Styles,
 * Crop, Arrange, Size, Properties), drawn from the window bindings it reads.
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
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeCommand,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { OfficeStyleGallery } from "@mrnote/core/web/components/office-ribbon/OfficeStyleGallery.tsx";
import {
  type PresentationCameoStyleId,
  PRESENTATION_CAMEO_STYLE_PRESETS,
  presentationCameoStyleSignature,
  PresentationCameoStylePreview,
  PresentationCameoShapeGallery,
  PresentationCameraBorderControl,
} from "./PresentationCameo";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuCheckboxItem,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuContent, RibbonDropdownMenuSubContent as DropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { presentationShapeKindMessageKey } from "../../../../lib/presentation/ribbon-tabs";
import {
  IconBringForward,
  IconBringToFront,
  IconSendBackward,
  IconSendToBack,
  IconPictureCrop,
  IconChangeShape,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import {
  AlignHorizontalJustifyStart,
  AlignHorizontalJustifyCenter,
  AlignHorizontalJustifyEnd,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type { AlignEdge } from "@mrnote/core/web/lib/presentation-align.ts";
import { geometryNumber } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { presentationPictureSizeDisplayCm } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import { MAX_TABLE_FRAME_CM } from "@mrnote/core/web/lib/presentation-table-size.ts";
import { presentationMeasurementCmDraft, type PresentationMeasurementUnit } from "@mrnote/core/web/lib/presentation-measurement.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { PresentationShapeBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { PresentationFrameAxis } from "../../../../lib/presentation/window-helpers";
import type { QuarterTurnDirection } from "@mrnote/core/web/lib/presentation-rotate-quarter.ts";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationCameraFormatRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  sizeUnit: PresentationMeasurementUnit;
  sizeUnitLabel: string;
  sizeStep: number;
  sizeBound: (centimetres: number) => number;
  sizeDisplayFromEmu: (emu: unknown, cm: string) => string;
  pictureCropTriggerRef: RefObject<HTMLButtonElement | null>;
  selectionPaneReturnFocusRef: RefObject<HTMLElement | null>;
  openAltTextPane: () => void;
  alignToSlide: boolean;
  writeAlignToSlide: (next: boolean) => void;
  cropEditingId: string | null;
  setCropEditingId: Dispatch<SetStateAction<string | null>>;
  themeAccentColors: string[];
  contextualCameoBlock: PresentationShapeBlock;
  contextualCameoAltTextWritable: boolean;
  contextualShapeFrameSizeEditable: boolean;
  selectionPaneOpen: boolean;
  commitContextualShapeFrameSize: (axis: PresentationFrameAxis, rawValue: string) => boolean;
  alignSelectedBlocks: (edge: AlignEdge) => void;
  autoFixSelectedBlocks: () => void;
  flipSelectedBlocks: (axis: "flipH" | "flipV") => void;
  canAutoFixSelection: boolean;
  canFlipSelection: boolean;
  quarterTurnSelectedBlocks: (direction: QuarterTurnDirection) => void;
  contextualShapeLayerCapability: Readonly<{ canMoveBackward: boolean; canMoveForward: boolean; }>;
  bringBlockToFront: (explicitId?: string, explicitOnly?: boolean) => void;
  sendBlockToBack: (explicitId?: string, explicitOnly?: boolean) => void;
  bringBlockForward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  sendBlockBackward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  updateShapeBlock: (blockId: string, patch: Partial<PresentationShapeBlock>) => boolean;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
}

export function PresentationCameraFormatRibbonTab({ ctl }: { ctl: PresentationCameraFormatRibbonTabCtl }) {
  const {
    t,
    sizeUnit,
    sizeUnitLabel,
    sizeStep,
    sizeBound,
    sizeDisplayFromEmu,
    pictureCropTriggerRef,
    selectionPaneReturnFocusRef,
    openAltTextPane,
    alignToSlide,
    writeAlignToSlide,
    cropEditingId,
    setCropEditingId,
    themeAccentColors,
    contextualCameoBlock,
    contextualCameoAltTextWritable,
    contextualShapeFrameSizeEditable,
    selectionPaneOpen,
    commitContextualShapeFrameSize,
    alignSelectedBlocks,
    autoFixSelectedBlocks,
    flipSelectedBlocks,
    canAutoFixSelection,
    canFlipSelection,
    quarterTurnSelectedBlocks,
    contextualShapeLayerCapability,
    bringBlockToFront,
    sendBlockToBack,
    bringBlockForward,
    sendBlockBackward,
    updateShapeBlock,
    featureDecision,
    executePresentationFeature,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.cameraStylesGroup}
                      label={t("presentation.cameo.stylesGroup")}
                      size="gallery"
                      testId="presentation-ribbon-camera-styles-group"
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.cameraStylesRow}
                      >
                        <OfficeStyleGallery<PresentationCameoStyleId>
                          application="presentation"
                          ariaLabel={t("presentation.cameo.stylesGroup")}
                          className={presentationChrome.cameraStyleGallery}
                          items={PRESENTATION_CAMEO_STYLE_PRESETS.map(
                            (preset) => ({
                              id: preset.id,
                              label: t(
                                `presentation.cameo.style.${preset.id}` as "presentation.cameo.style.none",
                              ),
                            }),
                          )}
                          quickIds={PRESENTATION_CAMEO_STYLE_PRESETS.map(
                            (preset) => preset.id,
                          )}
                          // All nine in the strip, as PowerPoint for the web shows them.
                          visibleItemCount={PRESENTATION_CAMEO_STYLE_PRESETS.length}
                          value={
                            PRESENTATION_CAMEO_STYLE_PRESETS.find((preset) =>
                              presentationCameoStyleSignature(
                                contextualCameoBlock,
                              ) === presentationCameoStyleSignature(preset.patch),
                            )?.id ?? null
                          }
                          labels={{ more: t("toolbar.quickOverflow") }}
                          renderPreview={(item) => {
                            const preset = PRESENTATION_CAMEO_STYLE_PRESETS.find(
                              (candidate) => candidate.id === item.id,
                            );
                            return preset ? (
                              <PresentationCameoStylePreview preset={preset} />
                            ) : null;
                          }}
                          onPreviewChange={() => undefined}
                          onCommit={(id) => {
                            const preset = PRESENTATION_CAMEO_STYLE_PRESETS.find(
                              (candidate) => candidate.id === id,
                            );
                            if (preset) {
                              updateShapeBlock(
                                contextualCameoBlock.id,
                                preset.patch,
                              );
                            }
                          }}
                          testIdPrefix="presentation-camera-style-gallery"
                        />
                        <OfficeHomeCommandStack
                          className={presentationChrome.cameraStyleCommands}
                        >
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="compact"
                                hasMenuIndicator
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="shape.change"
                                    size={20}
                                  />
                                }
                                label={t("presentation.cameo.shape")}
                                testId="presentation-ribbon-camera-shape"
                              />
                            </DropdownMenuTrigger>
                            <RibbonDropdownMenuContent
                              align="start"
                              data-testid="presentation-ribbon-camera-shape-menu"
                              style={{ maxWidth: "none" }}
                            >
                              <PresentationCameoShapeGallery
                                current={normalizeShapeKind(
                                  contextualCameoBlock.shape,
                                )}
                                labelFor={(kind) =>
                                  t(presentationShapeKindMessageKey(kind))
                                }
                                onSelect={(shape) =>
                                  updateShapeBlock(contextualCameoBlock.id, {
                                    shape,
                                  })
                                }
                                testPrefix="presentation-camera-shape"
                                t={t}
                              />
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                          <PresentationCameraBorderControl
                            accentColor={themeAccentColors[0]}
                            block={contextualCameoBlock}
                            onPatch={(patch) =>
                              updateShapeBlock(contextualCameoBlock.id, patch)
                            }
                            t={t}
                          />
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.cameraPropertiesGroup}
                      label={t("presentation.smartartContext.propertiesGroup")}
                      size="narrow"
                      testId="presentation-ribbon-camera-properties-group"
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
                        disabled={!contextualCameoAltTextWritable}
                        onClick={openAltTextPane}
                        testId="presentation-ribbon-camera-alt-text"
                      />
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.cameraArrangeGroup}
                      label={t("presentation.home.arrange")}
                      size="standard"
                      testId="presentation-ribbon-camera-arrange-group"
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.cameraArrangeRow}
                        testId="presentation-ribbon-camera-arrange-row"
                      >
                        <OfficeHomeCommandStack
                          className={presentationChrome.cameraArrangeStack}
                        >
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="compact"
                                hasMenuIndicator
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="object.bringForward"
                                    size={20}
                                  />
                                }
                                label={t(
                                  "presentation.contextMenu.bringForward",
                                )}
                                disabled={
                                  !contextualShapeLayerCapability.canMoveForward
                                }
                                testId="presentation-ribbon-camera-bring-menu"
                              />
                            </DropdownMenuTrigger>
                            <RibbonDropdownMenuContent align="start">
                              <DropdownMenuItem
                                disabled={
                                  !contextualShapeLayerCapability.canMoveForward
                                }
                                onSelect={() =>
                                  bringBlockForward(
                                    contextualCameoBlock.id,
                                    true,
                                  )
                                }
                                data-testid="presentation-ribbon-camera-bring-forward"
                              >
                                <IconBringForward size={16} />
                                {t("presentation.contextMenu.bringForward")}
                              </DropdownMenuItem>
                              <DropdownMenuItem
                                disabled={
                                  !contextualShapeLayerCapability.canMoveForward
                                }
                                onSelect={() =>
                                  bringBlockToFront(
                                    contextualCameoBlock.id,
                                    true,
                                  )
                                }
                                data-testid="presentation-ribbon-camera-bring-front"
                              >
                                <IconBringToFront size={16} />
                                {t("presentation.contextMenu.bringToFront")}
                              </DropdownMenuItem>
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="compact"
                                hasMenuIndicator
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="object.sendBackward"
                                    size={20}
                                  />
                                }
                                label={t(
                                  "presentation.contextMenu.sendBackward",
                                )}
                                disabled={
                                  !contextualShapeLayerCapability.canMoveBackward
                                }
                                testId="presentation-ribbon-camera-send-menu"
                              />
                            </DropdownMenuTrigger>
                            <RibbonDropdownMenuContent align="start">
                              <DropdownMenuItem
                                disabled={
                                  !contextualShapeLayerCapability.canMoveBackward
                                }
                                onSelect={() =>
                                  sendBlockBackward(
                                    contextualCameoBlock.id,
                                    true,
                                  )
                                }
                                data-testid="presentation-ribbon-camera-send-backward"
                              >
                                <IconSendBackward size={16} />
                                {t("presentation.contextMenu.sendBackward")}
                              </DropdownMenuItem>
                              <DropdownMenuItem
                                disabled={
                                  !contextualShapeLayerCapability.canMoveBackward
                                }
                                onSelect={() =>
                                  sendBlockToBack(contextualCameoBlock.id, true)
                                }
                                data-testid="presentation-ribbon-camera-send-back"
                              >
                                <IconSendToBack size={16} />
                                {t("presentation.contextMenu.sendToBack")}
                              </DropdownMenuItem>
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                          <OfficeHomeCommand
                            variant="compact"
                            active={selectionPaneOpen}
                            icon={
                              <OfficeSemanticIcon
                                semanticId="object.selectionPane"
                                size={20}
                              />
                            }
                            label={t("toolbar.selectionPane")}
                            disabled={
                              !featureDecision("editing.selectionPane").enabled
                            }
                            onClick={() => {
                              selectionPaneReturnFocusRef.current =
                                document.activeElement instanceof HTMLButtonElement
                                  ? document.activeElement
                                  : null;
                              executePresentationFeature(
                                "editing.selectionPane",
                              );
                            }}
                            testId="presentation-ribbon-camera-selection-pane"
                          />
                        </OfficeHomeCommandStack>
                        <OfficeHomeCommandStack
                          className={presentationChrome.cameraArrangeStack}
                        >
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="compact"
                                hasMenuIndicator
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="object.align"
                                    size={20}
                                  />
                                }
                                label={t("presentation.ribbon.align")}
                                disabled={!canFlipSelection}
                                testId="presentation-ribbon-camera-align-menu"
                              />
                            </DropdownMenuTrigger>
                            <RibbonDropdownMenuContent align="start">
                              <DropdownMenuCheckboxItem
                                checked={alignToSlide}
                                disabled={!canFlipSelection}
                                onCheckedChange={writeAlignToSlide}
                                data-testid="presentation-camera-align-to-slide"
                              >
                                {t("presentation.align.toSlideLabel")}
                              </DropdownMenuCheckboxItem>
                              <DropdownMenuSeparator />
                              {(
                                [
                                  ["left", AlignHorizontalJustifyStart],
                                  ["hcenter", AlignHorizontalJustifyCenter],
                                  ["right", AlignHorizontalJustifyEnd],
                                  ["top", AlignVerticalJustifyStart],
                                  ["vcenter", AlignVerticalJustifyCenter],
                                  ["bottom", AlignVerticalJustifyEnd],
                                ] as ReadonlyArray<
                                  [
                                    AlignEdge,
                                    typeof AlignHorizontalJustifyStart,
                                  ]
                                >
                              ).map(([edge, Icon]) => (
                                <DropdownMenuItem
                                  key={edge}
                                  disabled={!canFlipSelection}
                                  onSelect={() => alignSelectedBlocks(edge)}
                                  data-testid={`presentation-camera-align-${edge}`}
                                >
                                  <Icon size={16} aria-hidden="true" />
                                  {t(
                                    `presentation.align.${edge}` as "presentation.align.left",
                                  )}
                                </DropdownMenuItem>
                              ))}
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                          <OfficeHomeCommand
                            variant="compact"
                            icon={
                              <OfficeSemanticIcon
                                semanticId="object.group"
                                size={20}
                              />
                            }
                            label={t("presentation.group.group")}
                            disabled
                            disabledReason={t("presentation.cameo.groupUnavailable")}
                            onClick={() => undefined}
                            testId="presentation-ribbon-camera-group"
                          />
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <OfficeHomeMenuTrigger
                                variant="compact"
                                hasMenuIndicator
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="object.rotate"
                                    size={20}
                                  />
                                }
                                label={t("presentation.rotate.right")}
                                disabled={!canFlipSelection}
                                testId="presentation-ribbon-camera-rotate-menu"
                              />
                            </DropdownMenuTrigger>
                            <RibbonDropdownMenuContent align="start">
                              <DropdownMenuItem
                                disabled={!canFlipSelection}
                                onSelect={() =>
                                  quarterTurnSelectedBlocks("left")
                                }
                                data-testid="presentation-camera-rotate-left"
                              >
                                <span
                                  aria-hidden="true"
                                  style={{
                                    display: "inline-flex",
                                    transform: "scaleX(-1)",
                                  }}
                                >
                                  <OfficeSemanticIcon
                                    semanticId="object.rotate"
                                    size={16}
                                  />
                                </span>
                                {t("presentation.rotate.left")}
                              </DropdownMenuItem>
                              <DropdownMenuItem
                                disabled={!canFlipSelection}
                                onSelect={() =>
                                  quarterTurnSelectedBlocks("right")
                                }
                                data-testid="presentation-camera-rotate-right"
                              >
                                <OfficeSemanticIcon
                                  semanticId="object.rotate"
                                  size={16}
                                />
                                {t("presentation.rotate.right")}
                              </DropdownMenuItem>
                              <DropdownMenuSeparator />
                              <DropdownMenuItem
                                disabled={!canFlipSelection}
                                onSelect={() => flipSelectedBlocks("flipH")}
                                data-testid="presentation-camera-flip-horizontal"
                              >
                                <OfficeSemanticIcon
                                  semanticId="object.flipHorizontal"
                                  size={16}
                                />
                                {t("presentation.flip.horizontal")}
                              </DropdownMenuItem>
                              <DropdownMenuItem
                                disabled={!canFlipSelection}
                                onSelect={() => flipSelectedBlocks("flipV")}
                                data-testid="presentation-camera-flip-vertical"
                              >
                                <OfficeSemanticIcon
                                  semanticId="object.flipVertical"
                                  size={16}
                                />
                                {t("presentation.flip.vertical")}
                              </DropdownMenuItem>
                            </RibbonDropdownMenuContent>
                          </DropdownMenu>
                        </OfficeHomeCommandStack>
                        <OfficeHomeCommand
                          variant="large"
                          icon={
                            <OfficeSemanticIcon
                              semanticId="image.enhance"
                              size={24}
                            />
                          }
                          label={t("presentation.cameo.autoFix")}
                          disabled={!canAutoFixSelection}
                          disabledReason={t("presentation.pictureRibbon.autoFixUnavailable")}
                          onClick={autoFixSelectedBlocks}
                          testId="presentation-ribbon-camera-auto-fix"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.cameraCropGroup}
                      label={t("presentation.cameo.crop")}
                      size="compact"
                      testId="presentation-ribbon-camera-crop-group"
                    >
                      <OfficeHomeCommandRow>
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            className={presentationChrome.pictureVerticalSplit}
                            testId="presentation-ribbon-camera-crop-split"
                            primary={
                              <OfficeHomeCommand
                                variant="large"
                                buttonRef={pictureCropTriggerRef}
                                active={
                                  cropEditingId === contextualCameoBlock.id
                                }
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId="image.crop"
                                    size={20}
                                  />
                                }
                                label={t("presentation.cameo.crop")}
                                onClick={() =>
                                  setCropEditingId((current) =>
                                    current === contextualCameoBlock.id
                                      ? null
                                      : contextualCameoBlock.id,
                                  )
                                }
                                testId="presentation-ribbon-camera-crop"
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
                                  label={t("presentation.cameo.crop")}
                                  testId="presentation-ribbon-camera-crop-menu"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-camera-crop-menu-content"
                            style={{ overflow: "visible" }}
                          >
                            <DropdownMenuItem
                              onSelect={() =>
                                setCropEditingId(contextualCameoBlock.id)
                              }
                              data-testid="presentation-camera-crop-free"
                            >
                              <IconPictureCrop home size={16} />
                              {t("presentation.cameo.crop")}
                            </DropdownMenuItem>
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger>
                                <IconChangeShape home size={16} />
                                {t("presentation.cameo.cropToShape")}
                              </DropdownMenuSubTrigger>
                              <DropdownMenuSubContent
                                data-testid="presentation-camera-crop-shape-menu" onEscapeKeyDown={(event) => event.stopPropagation()}
                                style={{ maxHeight: "min(560px, var(--radix-dropdown-menu-content-available-height, calc(100vh - 24px)))", maxWidth: "none", overflowY: "auto" }}
                              >
                                <PresentationCameoShapeGallery
                                  current={normalizeShapeKind(
                                    contextualCameoBlock.shape,
                                  )}
                                  labelFor={(kind) =>
                                    t(presentationShapeKindMessageKey(kind))
                                  }
                                  onSelect={(shape) =>
                                    updateShapeBlock(contextualCameoBlock.id, {
                                      shape,
                                    })
                                  }
                                  testPrefix="presentation-camera-crop-shape"
                                  t={t}
                                />
                              </DropdownMenuSubContent>
                            </DropdownMenuSub>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={`${presentationChrome.shapeSizeGroup} ${presentationChrome.cameraSizeGroup}`}
                      label={t("presentation.cameo.sizeGroup")}
                      size="standard"
                      testId="presentation-ribbon-camera-size-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommandStack>
                          {(["height", "width"] as const).map((axis) => {
                            const currentSize = geometryNumber(
                              contextualCameoBlock,
                              axis,
                            );
                            const label = t(
                              axis === "height"
                                ? "presentation.tableRibbon.heightShort"
                                : "presentation.tableRibbon.widthShort",
                            );
                            return (
                              <span
                                className={`note-editor-office-toolbar__field-label ${presentationChrome.shapeSizeControl}`}
                                key={axis}
                                style={{ flex: "0 0 auto" }}
                              >
                                <OfficeRibbonNumberField
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId={
                                        axis === "height"
                                          ? "object.height"
                                          : "object.width"
                                      }
                                      size={18}
                                    />
                                  }
                                  visibleLabel={label}
                                  testId={`presentation-ribbon-camera-${axis}`}
                                  label={label}
                                  unit={sizeUnitLabel}
                                  value={sizeDisplayFromEmu(
                                    currentSize,
                                    presentationPictureSizeDisplayCm(currentSize),
                                  )}
                                  min={0.01}
                                  max={sizeBound(MAX_TABLE_FRAME_CM)}
                                  step={sizeStep}
                                  disabled={!contextualShapeFrameSizeEditable}
                                  onCommit={(rawValue) =>
                                    commitContextualShapeFrameSize(
                                      axis,
                                      presentationMeasurementCmDraft(rawValue, sizeUnit),
                                    )
                                  }
                                />
                              </span>
                            );
                          })}
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
