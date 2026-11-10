"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The View tab of the presentation Ribbon (Presentation Views, Master Views, Show, Zoom, Window), drawn from the window bindings it reads.
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
  OfficeHomeCommandStack,
  OfficeHomeCheckbox,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import {
  ListTree,
  LayoutGrid,
  StickyNote,
  LayoutTemplate,
  Maximize2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { ContinuousZoomCrownGlyph } from "@mrnote/core/web/components/ui/ContinuousZoomDialPopover.tsx";
import { presentationZoomPositionFromPercent, type PresentationLiveView, type PresentationViewportZoomMode } from "../../../../lib/presentation/view-zoom";
import { PRESENTATION_VIEW_ZOOM_MIN_PERCENT, type PresentationViewAidKind } from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationViewColorWindowGroups } from "./PresentationViewColorWindowGroups";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { ImmersiveReaderProjection } from "@mrnote/docs/web/components/console/editor/native-document/DocumentImmersiveReader.tsx";
import type { PresentationViewColorMode } from "@mrnote/core/web/lib/presentation/presentation-view-color-mode.ts";
import type { PresentationViewAidCommandAuthority, PresentationViewportZoomCommandAuthority } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";
import type { PresentationBlackWhiteMode } from "@mrnote/core/web/lib/presentation/presentation-black-white-mode.ts";
import type { OfficeWindowChoice } from "@mrnote/core/web/lib/office-window-commands/windowCommands.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationViewRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  slideZoomMode: PresentationViewportZoomMode;
  zoomRibbonTriggerRef: RefObject<HTMLButtonElement | null>;
  presentationZoomDialogOpen: boolean;
  setPresentationZoomDialogOpen: Dispatch<SetStateAction<boolean>>;
  showRuler: boolean;
  showGridlines: boolean;
  showGuides: boolean;
  immersiveReaderProjection: ImmersiveReaderProjection;
  immersiveReaderOpen: boolean;
  setImmersiveReaderOpen: Dispatch<SetStateAction<boolean>>;
  viewColorMode: PresentationViewColorMode;
  setViewColorMode: Dispatch<SetStateAction<PresentationViewColorMode>>;
  immersiveReaderTriggerRef: RefObject<HTMLButtonElement | null>;
  ribbonReadingViewTriggerRef: RefObject<HTMLButtonElement | null>;
  activePresentationLiveView: PresentationLiveView | null;
  dispatchPresentationView: (next: PresentationLiveView, options?: { trigger?: HTMLElement | null; restoreReadingFocus?: boolean; }) => boolean;
  notesVisible: boolean;
  togglePresentationNotesPane: () => void;
  viewAidCommandAuthority: PresentationViewAidCommandAuthority;
  viewportZoomCommandAuthority: PresentationViewportZoomCommandAuthority;
  presentationBlackWhiteObject: { onChange: (mode: PresentationBlackWhiteMode) => void; disabledReason?: string | undefined; label: string; modeLabels: Record<PresentationBlackWhiteMode, string>; mode: "blackWhite" | "auto" | "gray" | "ltGray" | "invGray" | "grayWhite" | "blackGray" | "black" | "white" | "hidden" | null; };
  togglePresentationViewAid: (kind: PresentationViewAidKind) => void;
  zoomFitToWindow: () => void;
  slideZoomPercentValue: number;
  slideZoomPercentLabel: string;
  slideZoomDialMaximumPercent: number;
  presentationViewWindowCommands: { choices: () => OfficeWindowChoice[]; onNewWindow: () => void; onArrangeAll: () => void; onCascade: () => void; onSwitch: (id: string) => void; } | null;
  /** Show group's dialog launcher: Grid Settings…. */
  gridSettingsOpen: boolean;
  openGridSettings: () => void;
}

export function PresentationViewRibbonTab({ ctl }: { ctl: PresentationViewRibbonTabCtl }) {
  const {
    t,
    slideZoomMode,
    zoomRibbonTriggerRef,
    presentationZoomDialogOpen,
    setPresentationZoomDialogOpen,
    showRuler,
    showGridlines,
    showGuides,
    immersiveReaderProjection,
    immersiveReaderOpen,
    setImmersiveReaderOpen,
    viewColorMode,
    setViewColorMode,
    immersiveReaderTriggerRef,
    ribbonReadingViewTriggerRef,
    activePresentationLiveView,
    dispatchPresentationView,
    notesVisible,
    togglePresentationNotesPane,
    viewAidCommandAuthority,
    viewportZoomCommandAuthority,
    presentationBlackWhiteObject,
    togglePresentationViewAid,
    zoomFitToWindow,
    slideZoomPercentValue,
    slideZoomPercentLabel,
    slideZoomDialMaximumPercent,
    presentationViewWindowCommands,
    gridSettingsOpen,
    openGridSettings,
  } = ctl;
  return (
                  <>
                    {/* Desktop PowerPoint's View tab (16.86 accessibility
                        tree; PPT-A104..A107): Presentation Views (Normal,
                        Outline View, Slide Sorter, Notes Page, Reading View)
                        | Master Views (Slide Master, Handout Master, Notes
                        Master) | Show (Ruler, Gridlines, Guides check boxes;
                        Notes) | Zoom. Every command is on the ribbon; the
                        "More" overflows are gone. */}
                    <OfficeHomeRibbonGroup
                      label={t("presentation.view.presentationViewsGroup")}
                      ribbon="view"
                      testId="presentation-ribbon-view-presentation-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "normal"}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.NORMAL"
                          icon={
                            <OfficeSemanticIcon
                              semanticId="view.document"
                              size={24}
                            />
                          }
                          label={t("presentation.view.normal")}
                          onClick={() => dispatchPresentationView("normal")}
                          testId="presentation-ribbon-view-rail"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "outline"}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.OUTLINE"
                          icon={<ListTree size={24} />}
                          label={t("presentation.view.outline")}
                          onClick={() => dispatchPresentationView("outline")}
                          testId="presentation-ribbon-view-outline"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "sorter"}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.SLIDE_SORTER"
                          icon={<LayoutGrid size={24} />}
                          label={t("presentation.view.slideSorter")}
                          onClick={() => dispatchPresentationView("sorter")}
                          testId="presentation-ribbon-view-sorter"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "notes"}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.NOTES_PAGE"
                          icon={<StickyNote size={24} />}
                          label={t("presentation.view.notesPage")}
                          onClick={() => dispatchPresentationView("notes")}
                          testId="presentation-ribbon-view-notes"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "reading"}
                          activation="click"
                          buttonRef={ribbonReadingViewTriggerRef}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.READING"
                          icon={
                            <OfficeSemanticIcon
                              semanticId="view.reading"
                              size={24}
                            />
                          }
                          label={t("presentation.view.reading")}
                          onClick={() =>
                            dispatchPresentationView("reading", {
                              trigger: ribbonReadingViewTriggerRef.current,
                            })
                          }
                          testId="presentation-ribbon-view-reading"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    {





}
                    <OfficeHomeRibbonGroup
                      label={t("presentation.view.masterViewsGroup")}
                      ribbon="view"
                      testId="presentation-ribbon-view-master-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          active={activePresentationLiveView === "master"}
                          controlId="PPTCTL.RIBBON.VIEW.PRESENTATION.SLIDE_MASTER"
                          icon={<LayoutTemplate size={24} />}
                          label={t("presentation.view.slideMaster")}
                          onClick={() => dispatchPresentationView("master")}
                          testId="presentation-ribbon-view-master"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.VIEW.MASTER.HANDOUT"
                          disabled
                          disabledReason={t("presentation.view.masterNotEditable")}
                          icon={<LayoutTemplate size={24} />}
                          label={t("presentation.view.handoutMaster")}
                          onClick={() => undefined}
                          testId="presentation-ribbon-view-handout-master"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.VIEW.MASTER.NOTES"
                          disabled
                          disabledReason={t("presentation.view.masterNotEditable")}
                          icon={<StickyNote size={24} />}
                          label={t("presentation.view.notesMaster")}
                          onClick={() => undefined}
                          testId="presentation-ribbon-view-notes-master"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("presentation.view.showGroup")}
                      ribbon="view"
                      testId="presentation-ribbon-view-show-group"
                      launcher={{
                        // PowerPoint for Windows: the Show group's launcher
                        // opens Grid and Guides (Grid Settings).
                        label: t("presentation.gridSettings.title"),
                        onClick: openGridSettings,
                        disabled: !viewAidCommandAuthority.isEnabled(),
                        active: gridSettingsOpen,
                        testId: "presentation-ribbon-view-show-launcher",
                      }}
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommandStack>
                          <OfficeHomeCheckbox
                            checked={showRuler}
                            disabled={!viewAidCommandAuthority.isEnabled()}
                            onCheckedChange={() => togglePresentationViewAid("ruler")}
                            controlId="PPTCTL.RIBBON.VIEW.SHOW.RULER"
                            label={t("presentation.view.ruler")}
                            testId="presentation-ribbon-ruler"
                          />
                          <OfficeHomeCheckbox
                            checked={showGridlines}
                            disabled={!viewAidCommandAuthority.isEnabled()}
                            onCheckedChange={() =>
                              togglePresentationViewAid("gridlines")
                            }
                            controlId="PPTCTL.RIBBON.VIEW.SHOW.GRIDLINES"
                            label={t("presentation.view.gridlines")}
                            testId="presentation-ribbon-gridlines"
                          />
                          <OfficeHomeCheckbox
                            checked={showGuides}
                            disabled={!viewAidCommandAuthority.isEnabled()}
                            onCheckedChange={() =>
                              togglePresentationViewAid("guides")
                            }
                            controlId="PPTCTL.RIBBON.VIEW.SHOW.GUIDES"
                            label={t("presentation.view.guides")}
                            testId="presentation-ribbon-guides"
                          />
                        </OfficeHomeCommandStack>
                        <OfficeHomeCommand
                          active={notesVisible}
                          controlId="PPTCTL.RIBBON.VIEW.SHOW.NOTES"
                          icon={<StickyNote size={24} />}
                          label={t("presentation.view.notes")}
                          onClick={togglePresentationNotesPane}
                          testId="presentation-ribbon-view-notes-toggle"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("presentation.view.zoomGroup")}
                      ribbon="view"
                      testId="presentation-ribbon-view-zoom-group"
                    >
                      <OfficeHomeCommandRow>
                        {/* View > Zoom opens PowerPoint's Zoom dialog
                            (B032 / A108); the status bar keeps the dial. */}
                        <OfficeHomeCommand
                          buttonRef={zoomRibbonTriggerRef}
                          controlId="PPTCTL.RIBBON.VIEW.ZOOM.DIALOG_ENTRY"
                          disabled={!viewportZoomCommandAuthority.isEnabled()}
                          icon={
                            <ContinuousZoomCrownGlyph
                              position={presentationZoomPositionFromPercent(
                                slideZoomPercentValue,
                                PRESENTATION_VIEW_ZOOM_MIN_PERCENT,
                                slideZoomDialMaximumPercent,
                              )}
                              valueLabel={slideZoomPercentLabel}
                              open={presentationZoomDialogOpen}
                              className={presentationChrome.zoomRibbonCrown}
                            />
                          }
                          label={`${t("presentation.view.zoom")} ${slideZoomPercentLabel}`}
                          displayLabel={t("presentation.view.zoom")}
                          onClick={() => {
                            // Live check: a stale enabled button during a show
                            // must not open the dialog.
                            if (viewportZoomCommandAuthority.isEnabled()) {
                              setPresentationZoomDialogOpen(true);
                            }
                          }}
                          testId="presentation-ribbon-view-zoom"
                          variant="large"
                        />
                        <OfficeHomeCommand
                          active={slideZoomMode === "fit"}
                          controlId="PPTCTL.RIBBON.VIEW.ZOOM.FIT"
                          disabled={!viewportZoomCommandAuthority.isEnabled()}
                          icon={<Maximize2 size={24} />}
                          label={t("presentation.view.fitToWindow")}
                          onClick={zoomFitToWindow}
                          testId="presentation-zoom-fit"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <PresentationViewColorWindowGroups
                      labels={{
                        colorGroup: t("presentation.view.colorGrayscaleGroup"),
                        color: t("presentation.view.color"),
                        grayscale: t("presentation.view.grayscale"),
                        pureBlackAndWhite: t("presentation.view.pureBlackAndWhite"),
                        windowGroup: t("toolbar.view.windowGroup"),
                        newWindow: t("toolbar.newWindow"),
                        arrangeAll: t("presentation.view.arrangeAll"),
                        cascade: t("presentation.view.cascade"),
                        switchWindows: t("toolbar.switchWindows"),
                        windowsUnavailable: t("presentation.view.windowsUnavailable"),
                      }}
                      colorMode={viewColorMode}
                      onColorMode={setViewColorMode}
                      windows={presentationViewWindowCommands}
                      selectedObject={presentationBlackWhiteObject}
                    />

                    {/* Immersive Reader: PowerPoint for Windows and for the
                        web offer it (Mac 16.86 does not); kept, in a group
                        of its own after Zoom, so the Presentation Views
                        group holds only views. */}
                    <OfficeHomeRibbonGroup
                      label={t("toolbar.immersiveReader")}
                      ribbon="view"
                      testId="presentation-ribbon-view-immersive-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          active={immersiveReaderOpen}
                          buttonRef={immersiveReaderTriggerRef}
                          controlId="PPTCTL.RIBBON.VIEW.IMMERSIVE_READER"
                          disabled={immersiveReaderProjection.blocks.length === 0}
                          icon={
                            <OfficeSemanticIcon
                              semanticId="view.immersiveReader"
                              size={24}
                            />
                          }
                          label={t("toolbar.immersiveReader")}
                          onClick={() => setImmersiveReaderOpen(true)}
                          testId="presentation-ribbon-view-immersive-reader"
                          variant="large"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
