"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Slide Show tab of the presentation Ribbon (Start Slide Show, Set Up, Monitors), drawn from the window bindings it reads.
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
  OfficeHomeCheckbox,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  IconSlideShowFromBeginning,
  IconSlideShowFromCurrent,
  IconSlideShowPresenterView,
  IconSlideShowSetUp,
  IconSlideShowHideSlide,
  IconRehearseTimings,
  IconSlideShowRecord,
  IconSlideShowCaptions,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { PRESENTATION_WEB_PRESENTER_VIEW_SUPPORTED, PRESENTATION_SLIDE_VISIBILITY_CONTROL_ID } from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import { ListVideo } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { isSlideHidden } from "@mrnote/core/web/lib/presentation-slide-visibility.ts";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuSub,
  DropdownMenuSubTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent, RibbonDropdownMenuSubContent as DropdownMenuSubContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { RefObject } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { PresentationSlideVisibilityCommandAdapter } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import type { ResolvedPresentationFeatureCommand } from "@mrnote/core/web/components/notebook/contents/ribbon/commandRegistry.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationSlideShowRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  rehearsalReturnFocusRef: RefObject<HTMLElement | null>;
  slideShowMoreTriggerRef: RefObject<HTMLButtonElement | null>;
  slideShowSetupTriggerRef: RefObject<HTMLButtonElement | null>;
  customShowTriggerRef: RefObject<HTMLButtonElement | null>;
  slides: SlideModel[];
  slideShowSetupButtonRef: RefObject<HTMLButtonElement | null>;
  slideShowRehearseButtonRef: RefObject<HTMLButtonElement | null>;
  useRecordedSlideTimings: boolean;
  configuredSlideShowType: "speaker" | "window" | "kiosk";
  showSlideShowMediaControls: boolean;
  playSlideShowNarrations: boolean;
  canPersistSlideShowMediaControls: boolean;
  presenterWindowOpen: boolean;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  activeSlide: SlideModel;
  clearActiveSlideNarration: () => void;
  clearAllSlideNarrations: () => void;
  clearRecordedSlideTimings: (scope: "current" | "all") => void;
  commitUseRecordedSlideTimings: (useTimings: boolean) => void;
  commitPlaySlideShowNarrations: (showNarration: boolean) => void;
  commitShowSlideShowMediaControls: (showMediaControls: boolean) => void;
  openCustomShowManager: () => void;
  openSlideShowSetup: () => void;
  openPresenterWindow: () => void;
  openRecordStudioFromBeginning: () => void;
  openRecordStudioFromCurrent: () => void;
  slideVisibilityCommand: PresentationSlideVisibilityCommandAdapter;
  featureDecision: (featureId: string) => ResolvedPresentationFeatureCommand;
  executePresentationFeature: (featureId: string) => undefined;
  toggleActiveSlideVisibility: () => void;
  activeSlideNarration: string | null;
  deckHasNarration: boolean;
  activeSlideHasRecordedTiming: boolean;
  deckHasRecordedTimings: boolean;
  narrationRecorderBusy: boolean;
}

export function PresentationSlideShowRibbonTab({ ctl }: { ctl: PresentationSlideShowRibbonTabCtl }) {
  const {
    t,
    rehearsalReturnFocusRef,
    slideShowMoreTriggerRef,
    slideShowSetupTriggerRef,
    customShowTriggerRef,
    slides,
    slideShowSetupButtonRef,
    slideShowRehearseButtonRef,
    useRecordedSlideTimings,
    configuredSlideShowType,
    showSlideShowMediaControls,
    playSlideShowNarrations,
    canPersistSlideShowMediaControls,
    presenterWindowOpen,
    autosave,
    activeSlide,
    clearActiveSlideNarration,
    clearAllSlideNarrations,
    clearRecordedSlideTimings,
    commitUseRecordedSlideTimings,
    commitPlaySlideShowNarrations,
    commitShowSlideShowMediaControls,
    openCustomShowManager,
    openSlideShowSetup,
    openPresenterWindow,
    openRecordStudioFromBeginning,
    openRecordStudioFromCurrent,
    slideVisibilityCommand,
    featureDecision,
    executePresentationFeature,
    toggleActiveSlideVisibility,
    activeSlideNarration,
    deckHasNarration,
    activeSlideHasRecordedTiming,
    deckHasRecordedTimings,
    narrationRecorderBusy,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.slideshow.startGroup")}
                      ribbon="slideshow"
                      size="content"
                      testId="presentation-ribbon-slideshow-start-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.SLIDESHOW.START.FROM_BEGINNING"
                          disabled={
                            !featureDecision("slideshow.fromStart").enabled
                          }
                          icon={<IconSlideShowFromBeginning home size={24} />}
                          label={t("presentation.slideshow.fromBeginning")}
                          onClick={() =>
                            executePresentationFeature("slideshow.fromStart")
                          }
                          testId="presentation-ribbon-slideshow-from-start"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.SLIDESHOW.START.FROM_CURRENT"
                          disabled={
                            !featureDecision("slideshow.fromCurrent").enabled
                          }
                          icon={<IconSlideShowFromCurrent home size={24} />}
                          label={t("presentation.slideshow.fromCurrent")}
                          onClick={() =>
                            executePresentationFeature("slideshow.fromCurrent")
                          }
                          testId="presentation-ribbon-slideshow-from-current"
                        />
                        {PRESENTATION_WEB_PRESENTER_VIEW_SUPPORTED ? (
                          <OfficeHomeCommand
                            variant="large"
                            active={presenterWindowOpen}
                            controlId="PPTCTL.RIBBON.SLIDESHOW.START.PRESENTER_VIEW"
                            disabled={configuredSlideShowType === "kiosk"}
                            icon={<IconSlideShowPresenterView home size={24} />}
                            label={t("presentation.openPresenterView")}
                            onClick={openPresenterWindow}
                            testId="presentation-ribbon-presenter-view"
                          />
                        ) : null}
                        <OfficeHomeCommand
                          variant="large"
                          buttonRef={customShowTriggerRef}
                          controlId="PPTCTL.RIBBON.SLIDESHOW.CUSTOM_SHOWS"
                          disabled={slides.length === 0}
                          icon={<ListVideo size={24} />}
                          label={t(
                            "presentation.slideshow.customShowManage",
                          )}
                          onClick={openCustomShowManager}
                          testId="presentation-ribbon-slideshow-custom-shows"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    {/* PPT-A094: desktop PowerPoint lays every Slide Show
                        command out flat -- Set Up Slide Show, Hide Slide,
                        Rehearse Timings, Record ▾ and the Play Narrations /
                        Use Timings / Show Media Controls check boxes (all
                        three checked by default), then Always Use Subtitles
                        and Subtitle Settings (the 16.86 accessibility tree).
                        They were folded into one "More" menu. */}
                    <OfficeHomeRibbonGroup
                      label={t("presentation.slideshow.setupGroup")}
                      ribbon="slideshow"
                      size="content"
                      testId="presentation-ribbon-slideshow-setup-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          activation="click"
                          buttonRef={slideShowSetupButtonRef}
                          controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.DIALOG_ENTRY"
                          icon={<IconSlideShowSetUp home size={24} />}
                          label={t("presentation.slideshow.setUpShow")}
                          onClick={() => {
                            slideShowSetupTriggerRef.current =
                              slideShowSetupButtonRef.current;
                            openSlideShowSetup();
                          }}
                          testId="presentation-ribbon-slideshow-setup"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          active={isSlideHidden(activeSlide)}
                          controlId={PRESENTATION_SLIDE_VISIBILITY_CONTROL_ID}
                          disabled={!slideVisibilityCommand.canExecute()}
                          icon={<IconSlideShowHideSlide home size={24} />}
                          label={t("presentation.hideSlide")}
                          onClick={toggleActiveSlideVisibility}
                          testId="presentation-ribbon-slideshow-hide-slide"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          activation="click"
                          buttonRef={slideShowRehearseButtonRef}
                          controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.REHEARSE"
                          disabled={
                            !featureDecision("slideshow.rehearse").enabled
                          }
                          icon={<IconRehearseTimings home size={24} />}
                          label={t("presentation.rehearseStart")}
                          onClick={() => {
                            rehearsalReturnFocusRef.current =
                              slideShowRehearseButtonRef.current;
                            executePresentationFeature("slideshow.rehearse");
                          }}
                          testId="presentation-ribbon-slideshow-rehearse"
                        />
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              ref={slideShowMoreTriggerRef}
                              variant="large"
                              hasMenuIndicator
                              icon={<IconSlideShowRecord home size={24} />}
                              label={t("presentation.slideshow.record")}
                              controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.RECORD"
                              testId="presentation-ribbon-slideshow-record-menu"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-slideshow-record-menu-content"
                          >
                            <DropdownMenuItem
                              disabled={autosave.readOnly || narrationRecorderBusy}
                              onSelect={openRecordStudioFromBeginning}
                              data-testid="presentation-ribbon-slideshow-record-from-beginning"
                            >
                              <IconSlideShowFromBeginning home size={16} />
                              {t("presentation.record.fromBeginning")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={autosave.readOnly || narrationRecorderBusy}
                              onSelect={openRecordStudioFromCurrent}
                              data-testid="presentation-ribbon-slideshow-record-from-current"
                            >
                              <IconSlideShowFromCurrent home size={16} />
                              {t("presentation.record.fromCurrent")}
                            </DropdownMenuItem>
                            {/* Desktop Slide Show ▸ Record ▸ Clear: the four
                                clears the Record tab's Clear Recording menu
                                offers (16.86 ppt-menu-ss-record). */}
                            <DropdownMenuSub>
                              <DropdownMenuSubTrigger
                                disabled={
                                  autosave.readOnly ||
                                  narrationRecorderBusy ||
                                  (!deckHasNarration && !deckHasRecordedTimings)
                                }
                                data-testid="presentation-ribbon-slideshow-record-clear"
                              >
                                {t("presentation.record.clear")}
                              </DropdownMenuSubTrigger>
                              <DropdownMenuSubContent>
                                <DropdownMenuItem
                                  disabled={!activeSlideHasRecordedTiming}
                                  onSelect={() => clearRecordedSlideTimings("current")}
                                  data-testid="presentation-ribbon-slideshow-record-clear-current-timing"
                                >
                                  {t("presentation.record.clearCurrentTiming")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={!deckHasRecordedTimings}
                                  onSelect={() => clearRecordedSlideTimings("all")}
                                  data-testid="presentation-ribbon-slideshow-record-clear-all-timings"
                                >
                                  {t("presentation.record.clearAllTimings")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={activeSlideNarration === null}
                                  onSelect={clearActiveSlideNarration}
                                  data-testid="presentation-ribbon-slideshow-record-clear-current-narration"
                                >
                                  {t("presentation.record.clearCurrentNarration")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={!deckHasNarration}
                                  onSelect={clearAllSlideNarrations}
                                  data-testid="presentation-ribbon-slideshow-record-clear-all-narrations"
                                >
                                  {t("presentation.record.clearAllNarrations")}
                                </DropdownMenuItem>
                              </DropdownMenuSubContent>
                            </DropdownMenuSub>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <OfficeHomeCommandStack>
                          <OfficeHomeCheckbox
                            checked={playSlideShowNarrations}
                            onCheckedChange={commitPlaySlideShowNarrations}
                            controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.PLAY_NARRATIONS"
                            label={t("presentation.slideshow.playNarrations")}
                            testId="presentation-ribbon-slideshow-play-narrations"
                          />
                          <OfficeHomeCheckbox
                            checked={useRecordedSlideTimings}
                            onCheckedChange={commitUseRecordedSlideTimings}
                            controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.USE_TIMINGS"
                            label={t("presentation.slideshow.useTimings")}
                            testId="presentation-ribbon-slideshow-use-timings"
                          />
                          <OfficeHomeCheckbox
                            checked={showSlideShowMediaControls}
                            disabled={!canPersistSlideShowMediaControls}
                            title={
                              canPersistSlideShowMediaControls
                                ? undefined
                                : t("presentation.slideshow.settingsUnavailable")
                            }
                            onCheckedChange={commitShowSlideShowMediaControls}
                            controlId="PPTCTL.RIBBON.SLIDESHOW.SETUP.MEDIA_CONTROLS"
                            label={t("presentation.slideshow.showMediaControls")}
                            testId="presentation-ribbon-slideshow-media-controls"
                          />
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    {/* PowerPoint for the web does have live subtitles --
                        Microsoft's "Present with real-time, automatic
                        captions or subtitles in PowerPoint" names this very
                        entry point -- so both stay, disabled with their
                        reason: no speech recogniser, provider or policy is
                        consulted anywhere in the deck editor. */}
                    <OfficeHomeRibbonGroup
                      label={t("presentation.slideshow.captionsGroup")}
                      ribbon="slideshow"
                      size="content"
                      testId="presentation-ribbon-slideshow-captions-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommandStack>
                          <OfficeHomeCheckbox
                            checked={false}
                            disabled
                            onCheckedChange={() => undefined}
                            aria-label={`${t(
                              "presentation.slideshow.alwaysUseSubtitles",
                            )}: ${t("presentation.slideshow.captionsUnavailable")}`}
                            title={t("presentation.slideshow.captionsUnavailable")}
                            label={t("presentation.slideshow.alwaysUseSubtitles")}
                            testId="presentation-ribbon-slideshow-always-use-subtitles"
                          />
                          <OfficeHomeCommand
                            variant="compact"
                            disabled
                            disabledReason={t(
                              "presentation.slideshow.captionsUnavailable",
                            )}
                            icon={<IconSlideShowCaptions home size={16} />}
                            label={t("presentation.slideshow.subtitleSettings")}
                            onClick={() => undefined}
                            testId="presentation-ribbon-slideshow-subtitle-settings"
                          />
                        </OfficeHomeCommandStack>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
