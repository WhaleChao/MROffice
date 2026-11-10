"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The media Format contextual tab of the presentation Ribbon (Format Pane,
 * Arrange, Size, Properties), drawn from the window bindings it reads.
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
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
  OfficeHomeCommandStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuCheckboxItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  IconArrangeMenu,
  IconBringForward,
  IconBringToFront,
  IconSendBackward,
  IconSendToBack,
  IconAlignObjects,
  IconSelectionPane,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  AlignHorizontalJustifyStart,
  AlignHorizontalJustifyCenter,
  AlignHorizontalJustifyEnd,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import type { AlignEdge } from "@mrnote/core/web/lib/presentation-align.ts";
import { tableFrameSizeCm, MIN_TABLE_FRAME_CM, MAX_TABLE_FRAME_CM } from "@mrnote/core/web/lib/presentation-table-size.ts";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { presentationMeasurementCmDraft, type PresentationMeasurementUnit } from "@mrnote/core/web/lib/presentation-measurement.ts";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";
import type { PresentationBlock, PresentationEmbedBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { PresentationFrameAxis } from "../../../../lib/presentation/window-helpers";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationMediaFormatRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  sizeUnit: PresentationMeasurementUnit;
  sizeUnitLabel: string;
  sizeStep: number;
  sizeBound: (centimetres: number) => number;
  sizeDisplayFromEmu: (emu: unknown, cm: string) => string;
  selectionPaneReturnFocusRef: RefObject<HTMLElement | null>;
  selectionPaneMediaTriggerRef: RefObject<HTMLButtonElement | null>;
  formatPaneOpen: boolean;
  rootRef: RefObject<HTMLDivElement | null>;
  alignToSlide: boolean;
  writeAlignToSlide: (next: boolean) => void;
  contextualMediaBlock: PresentationEmbedBlock;
  contextualMediaSelectionIsWritable: boolean;
  contextualMediaKind: "iframe" | "video" | "audio" | null;
  selectionPaneOpen: boolean;
  toggleFormatPane: () => void;
  openFormatPane: () => void;
  selectedGeometryBlocks: PresentationBlock[];
  alignSelectedBlocks: (edge: AlignEdge) => void;
  contextualMediaLayerCapability: Readonly<{ canMoveBackward: boolean; canMoveForward: boolean; }>;
  bringBlockToFront: (explicitId?: string, explicitOnly?: boolean) => void;
  sendBlockToBack: (explicitId?: string, explicitOnly?: boolean) => void;
  bringBlockForward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  sendBlockBackward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  commitContextualMediaFrameSize: (axis: PresentationFrameAxis, rawValue: string) => boolean;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
  ribbonMediaHasGeometry: boolean;
}

export function PresentationMediaFormatRibbonTab({ ctl }: { ctl: PresentationMediaFormatRibbonTabCtl }) {
  const {
    t,
    sizeUnit,
    sizeUnitLabel,
    sizeStep,
    sizeBound,
    sizeDisplayFromEmu,
    selectionPaneReturnFocusRef,
    selectionPaneMediaTriggerRef,
    formatPaneOpen,
    rootRef,
    alignToSlide,
    writeAlignToSlide,
    contextualMediaBlock,
    contextualMediaSelectionIsWritable,
    contextualMediaKind,
    selectionPaneOpen,
    toggleFormatPane,
    openFormatPane,
    selectedGeometryBlocks,
    alignSelectedBlocks,
    contextualMediaLayerCapability,
    bringBlockToFront,
    sendBlockToBack,
    bringBlockForward,
    sendBlockBackward,
    commitContextualMediaFrameSize,
    featureDecision,
    executePresentationFeature,
    ribbonMediaHasGeometry,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      className={`${presentationChrome.mediaFormatPropertiesGroup} ${
                        contextualMediaKind === "video"
                            ? presentationChrome.mediaVideoOrderProperties
                            : ""
                      }`}
                      label={t("presentation.smartartContext.propertiesGroup")}
                      size="standard"
                      testId="presentation-ribbon-media-properties-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          icon={<OfficeSemanticIcon semanticId="object.altText" size={24} />}
                          label={t("presentation.altTextButton")}
                          disabled={!contextualMediaSelectionIsWritable}
                          onClick={() => { openFormatPane(); requestAnimationFrame(() => rootRef.current?.querySelector<HTMLInputElement>('[data-testid="presentation-embed-title"]')?.focus({ preventScroll: true })); }}
                          testId="presentation-ribbon-media-properties"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={`${presentationChrome.mediaFormatArrangeGroup} ${
                        contextualMediaKind === "video"
                            ? presentationChrome.mediaVideoOrderArrange
                            : ""
                      }`}
                      label={t("presentation.home.arrange")}
                      size="wide"
                      testId="presentation-ribbon-media-arrange-group"
                    >
                      <OfficeHomeCommandRow>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="large"
                              displayLabel={t("presentation.home.arrange")}
                              hasMenuIndicator
                              icon={<IconArrangeMenu size={24} />}
                              label={t("presentation.home.arrange")}
                              disabled={
                                !contextualMediaSelectionIsWritable ||
                                !ribbonMediaHasGeometry
                              }
                              controlId="PPTCTL.CONTEXT.MEDIA.FORMAT.ARRANGE"
                              testId="presentation-ribbon-media-arrange"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-media-arrange-menu"
                          >
                            <DropdownMenuItem
                              disabled={
                                !contextualMediaLayerCapability.canMoveForward
                              }
                              onSelect={() =>
                                bringBlockForward(contextualMediaBlock.id, true)
                              }
                              data-testid="presentation-ribbon-media-bring-forward"
                            >
                              <IconBringForward size={14} />
                              {t("presentation.contextMenu.bringForward")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                !contextualMediaLayerCapability.canMoveForward
                              }
                              onSelect={() =>
                                bringBlockToFront(contextualMediaBlock.id, true)
                              }
                              data-testid="presentation-ribbon-media-bring-front"
                            >
                              <IconBringToFront size={14} />
                              {t("presentation.contextMenu.bringToFront")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                !contextualMediaLayerCapability.canMoveBackward
                              }
                              onSelect={() =>
                                sendBlockBackward(contextualMediaBlock.id, true)
                              }
                              data-testid="presentation-ribbon-media-send-backward"
                            >
                              <IconSendBackward size={14} />
                              {t("presentation.contextMenu.sendBackward")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                !contextualMediaLayerCapability.canMoveBackward
                              }
                              onSelect={() =>
                                sendBlockToBack(contextualMediaBlock.id, true)
                              }
                              data-testid="presentation-ribbon-media-send-back"
                            >
                              <IconSendToBack size={14} />
                              {t("presentation.contextMenu.sendToBack")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="large"
                              active={alignToSlide}
                              displayLabel={t(
                                "presentation.ribbon.align",
                              )}
                              hasMenuIndicator
                              icon={<IconAlignObjects size={24} />}
                              label={t("presentation.ribbon.align")}
                              disabled={
                                !contextualMediaSelectionIsWritable ||
                                !ribbonMediaHasGeometry
                              }
                              controlId="PPTCTL.CONTEXT.MEDIA.FORMAT.ALIGN"
                              testId="presentation-ribbon-media-align"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-media-align-menu"
                          >
                            <DropdownMenuCheckboxItem
                              checked={alignToSlide}
                              onCheckedChange={(checked) =>
                                writeAlignToSlide(checked === true)
                              }
                              data-testid="presentation-ribbon-media-align-slide"
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
                                [AlignEdge, typeof AlignHorizontalJustifyStart]
                              >
                            ).map(([edge, Icon]) => (
                              <DropdownMenuItem
                                key={edge}
                                disabled={
                                  selectedGeometryBlocks.length <
                                  (alignToSlide ? 1 : 2)
                                }
                                onSelect={() => alignSelectedBlocks(edge)}
                                data-testid={`presentation-ribbon-media-align-${edge}`}
                              >
                                <Icon size={14} aria-hidden="true" />
                                {t(
                                  `presentation.align.${edge}` as "presentation.align.left",
                                )}
                              </DropdownMenuItem>
                            ))}
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <OfficeHomeCommand
                          active={selectionPaneOpen}
                          buttonRef={selectionPaneMediaTriggerRef}
                          controlId="PPTCTL.CONTEXT.MEDIA.FORMAT.SELECTION_PANE"
                          disabled={
                            !contextualMediaSelectionIsWritable ||
                            !featureDecision("editing.selectionPane").enabled
                          }
                          icon={<IconSelectionPane size={20} />}
                          label={t("toolbar.selectionPane")}
                          onClick={() => {
                            selectionPaneReturnFocusRef.current =
                              selectionPaneMediaTriggerRef.current;
                            executePresentationFeature(
                              "editing.selectionPane",
                            );
                          }}
                          testId="presentation-ribbon-media-selection-pane"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={`${presentationChrome.mediaFormatSizeGroup} ${
                        contextualMediaKind === "video"
                            ? presentationChrome.mediaVideoOrderSize
                            : ""
                      }`}
                      label={t("presentation.smartartContext.sizeGroup")}
                      size="standard"
                      testId="presentation-ribbon-media-size-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommandStack>
                          {(["height", "width"] as const).map((axis) => {
                            const currentSize = tableFrameSizeCm(
                              contextualMediaBlock[axis],
                            );
                            const label = t(
                              axis === "height" ? "presentation.tableRibbon.heightShort" : "presentation.tableRibbon.widthShort",
                            );
                            return (
                              <OfficeRibbonNumberField
                                key={axis}
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
                                testId={`presentation-ribbon-media-${axis}`}
                                label={label}
                                unit={sizeUnitLabel}
                                value={
                                  currentSize === null
                                    ? ""
                                    : sizeDisplayFromEmu(
                                        contextualMediaBlock[axis],
                                        String(currentSize),
                                      )
                                }
                                min={sizeBound(MIN_TABLE_FRAME_CM)}
                                max={sizeBound(MAX_TABLE_FRAME_CM)}
                                step={sizeStep}
                                disabled={
                                  !ribbonMediaHasGeometry ||
                                  !contextualMediaSelectionIsWritable
                                }
                                onCommit={(rawValue) =>
                                  commitContextualMediaFrameSize(
                                    axis,
                                    presentationMeasurementCmDraft(rawValue, sizeUnit),
                                  )
                                }
                              />
                            );
                          })}
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    {contextualMediaKind === "video" ? (
                      // Video Format ends in Format Pane, as Picture Format
                      // does (PPT-A155).
                      <OfficeHomeRibbonGroup
                        className={presentationChrome.mediaVideoOrderFormatPane}
                        label={t("presentation.media.formatPaneGroup")}
                        size="narrow"
                        testId="presentation-ribbon-media-format-pane-group"
                      >
                        <OfficeHomeCommandRow>
                          <OfficeHomeCommand
                            active={formatPaneOpen}
                            variant="large"
                            icon={<OfficeSemanticIcon semanticId="image.format" size={24} />}
                            label={t("presentation.media.formatPane")}
                            onClick={toggleFormatPane}
                            testId="presentation-ribbon-media-format-pane"
                          />
                        </OfficeHomeCommandRow>
                      </OfficeHomeRibbonGroup>
                    ) : null}
                  </>
  );
}
