"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Record tab of the presentation Ribbon (Record, Narration, Options,
 * Input), drawn from the window bindings it reads.
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
  OfficeHomeSplitLayout,
  OfficeHomeMenuTrigger,
  OfficeHomeCommandStack,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  Circle,
  Square,
  Play,
  Trash2,
  RotateCcw,
  Volume2,
  StickyNote,
  BookOpen,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { IconSlideShowFromBeginning, IconSlideShowFromCurrent, IconInsertCameo } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { isPresentationCameoBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import WindowNotice from "@mrnote/core/web/components/notebook/WindowNotice.tsx";
import { PresentationMicrophoneDeviceMenu } from "./PresentationCameo";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationRecordRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  narrationRecorderStatus: "idle" | "requesting" | "recording" | "paused" | "processing";
  narrationRecorderError: string | null;
  recordStudioOpen: boolean;
  recordStudioEntry: "beginning" | "current" | null;
  recordStudioTriggerRef: RefObject<HTMLButtonElement | null>;
  recordStudioCurrentTriggerRef: RefObject<HTMLButtonElement | null>;
  narrationMicrophones: readonly Readonly<{ deviceId: string; label: string; }>[];
  selectedNarrationMicrophoneId: string;
  setSelectedNarrationMicrophoneId: Dispatch<SetStateAction<string>>;
  narrationPreviewAudioRef: RefObject<HTMLAudioElement | null>;
  setNarrationPreviewPlaying: Dispatch<SetStateAction<boolean>>;
  setNarrationPreviewCurrentSeconds: Dispatch<SetStateAction<number>>;
  setNarrationPreviewDurationSeconds: Dispatch<SetStateAction<number>>;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  activeSlide: SlideModel;
  startNarrationRecording: (target?: "narration" | "audio-clip") => Promise<void>;
  stopNarrationRecording: () => void;
  clearActiveSlideNarration: () => void;
  clearAllSlideNarrations: () => void;
  clearRecordedSlideTimings: (scope: "current" | "all") => void;
  previewActiveSlideNarration: () => void;
  teleprompterOpen: boolean;
  teleprompterTriggerRef: RefObject<HTMLButtonElement | null>;
  closeTeleprompter: () => void;
  openTeleprompter: () => void;
  openRecordStudioFromBeginning: () => void;
  openRecordStudioFromCurrent: () => void;
  openPresentationRecordingHelp: () => void;
  insertCameo: () => boolean;
  resetRecordedSlidesToCameo: (scope: "current" | "all") => boolean;
  insertCapacityReached: boolean;
  narrationRecorderSupported: boolean;
  activeSlideNarration: string | null;
  deckHasNarration: boolean;
  activeSlideCanResetToCameo: boolean;
  deckCanResetAllToCameo: boolean;
  activeSlideHasRecordedTiming: boolean;
  deckHasRecordedTimings: boolean;
  narrationRecorderBusy: boolean;
  selectedNarrationMicrophoneLabel: string;
}

export function PresentationRecordRibbonTab({ ctl }: { ctl: PresentationRecordRibbonTabCtl }) {
  const {
    t,
    narrationRecorderStatus,
    narrationRecorderError,
    recordStudioOpen,
    recordStudioEntry,
    recordStudioTriggerRef,
    recordStudioCurrentTriggerRef,
    narrationMicrophones,
    selectedNarrationMicrophoneId,
    setSelectedNarrationMicrophoneId,
    narrationPreviewAudioRef,
    setNarrationPreviewPlaying,
    setNarrationPreviewCurrentSeconds,
    setNarrationPreviewDurationSeconds,
    autosave,
    activeSlide,
    startNarrationRecording,
    stopNarrationRecording,
    clearActiveSlideNarration,
    clearAllSlideNarrations,
    clearRecordedSlideTimings,
    previewActiveSlideNarration,
    teleprompterOpen,
    teleprompterTriggerRef,
    closeTeleprompter,
    openTeleprompter,
    openRecordStudioFromBeginning,
    openRecordStudioFromCurrent,
    openPresentationRecordingHelp,
    insertCameo,
    resetRecordedSlidesToCameo,
    insertCapacityReached,
    narrationRecorderSupported,
    activeSlideNarration,
    deckHasNarration,
    activeSlideCanResetToCameo,
    deckCanResetAllToCameo,
    activeSlideHasRecordedTiming,
    deckHasRecordedTimings,
    narrationRecorderBusy,
    selectedNarrationMicrophoneLabel,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.record.recordGroup")}
                      ribbon="record"
                      size="content"
                      testId="presentation-ribbon-record-capture-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.RECORD.CURRENT"
                          disabled={
                            autosave.readOnly ||
                            narrationRecorderBusy ||
                            !narrationRecorderSupported
                          }
                          icon={<Circle size={24} fill="currentColor" />}
                          displayLabel={t(
                            "presentation.record.currentSlide",
                          )}
                          label={
                            narrationRecorderSupported
                              ? t("presentation.record.currentSlide")
                              : `${t("presentation.record.currentSlide")} — ${t(
                                  "presentation.record.compatibleFormatUnavailable",
                                )}`
                          }
                          onClick={() => void startNarrationRecording()}
                          testId="presentation-ribbon-record-current"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.RECORD.STOP"
                          disabled={
                            narrationRecorderStatus !== "recording" &&
                            narrationRecorderStatus !== "paused"
                          }
                          icon={<Square size={24} fill="currentColor" />}
                          label={t("presentation.record.stop")}
                          onClick={stopNarrationRecording}
                          testId="presentation-ribbon-record-stop"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          active={recordStudioEntry === "beginning"}
                          ariaExpanded={recordStudioOpen}
                          buttonRef={recordStudioTriggerRef}
                          controlId="PPTCTL.RIBBON.RECORD.START.BEGINNING"
                          disabled={
                            autosave.readOnly || narrationRecorderBusy
                          }
                          icon={<IconSlideShowFromBeginning home size={24} />}
                          displayLabel={t(
                            "presentation.record.fromBeginning",
                          )}
                          label={t("presentation.record.fromBeginning")}
                          onClick={openRecordStudioFromBeginning}
                          testId="presentation-ribbon-record-from-beginning"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          active={recordStudioEntry === "current"}
                          ariaExpanded={recordStudioOpen}
                          buttonRef={recordStudioCurrentTriggerRef}
                          controlId="PPTCTL.RIBBON.RECORD.START.CURRENT"
                          disabled={
                            autosave.readOnly || narrationRecorderBusy
                          }
                          icon={<IconSlideShowFromCurrent home size={24} />}
                          displayLabel={t("presentation.record.fromCurrent")}
                          label={t("presentation.record.fromCurrent")}
                          onClick={openRecordStudioFromCurrent}
                          testId="presentation-ribbon-record-from-current"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.record.narrationGroup")}
                      ribbon="record"
                      size="standard"
                      testId="presentation-ribbon-record-narration-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.RECORD.PREVIEW"
                          disabled={
                            activeSlideNarration === null ||
                            narrationRecorderBusy
                          }
                          icon={<Play size={24} />}
                          label={t("presentation.record.preview")}
                          onClick={previewActiveSlideNarration}
                          testId="presentation-ribbon-record-preview"
                        />
                        <OfficeHomeSplitLayout
                          testId="presentation-ribbon-record-clear-split"
                          primary={
                            <OfficeHomeCommand
                              variant="large"
                              controlId="PPTCTL.RIBBON.RECORD.CLEAR.CURRENT_NARRATION"
                              disabled={
                                autosave.readOnly ||
                                activeSlideNarration === null ||
                                narrationRecorderBusy
                              }
                              icon={<Trash2 size={24} />}
                              label={t("presentation.record.clearCurrent")}
                              onClick={clearActiveSlideNarration}
                              testId="presentation-ribbon-record-clear"
                            />
                          }
                          trigger={
                            <DropdownMenu>
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  variant="icon"
                                  icon={<OfficeSemanticIcon semanticId="menu.chevron.down" size={12} />}
                                  disabled={
                                    autosave.readOnly ||
                                    narrationRecorderBusy ||
                                    (!deckHasNarration &&
                                      !deckHasRecordedTimings)
                                  }
                                  label={t("presentation.record.clearMenu")}
                                  testId="presentation-ribbon-record-clear-menu"
                                />
                              </DropdownMenuTrigger>
                              <RibbonDropdownMenuContent align="start">
                                <DropdownMenuItem
                                  disabled={activeSlideNarration === null}
                                  onSelect={clearActiveSlideNarration}
                                  data-testid="presentation-ribbon-record-clear-current-narration"
                                >
                                  {t("presentation.record.clearCurrent")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={!deckHasNarration}
                                  onSelect={clearAllSlideNarrations}
                                  data-testid="presentation-ribbon-record-clear-all-narrations"
                                >
                                  {t("presentation.record.clearAllNarrations")}
                                </DropdownMenuItem>
                                <DropdownMenuSeparator />
                                <DropdownMenuItem
                                  disabled={!activeSlideHasRecordedTiming}
                                  onSelect={() =>
                                    clearRecordedSlideTimings("current")
                                  }
                                  data-testid="presentation-ribbon-record-clear-current-timing"
                                >
                                  {t("presentation.record.clearCurrentTiming")}
                                </DropdownMenuItem>
                                <DropdownMenuItem
                                  disabled={!deckHasRecordedTimings}
                                  onSelect={() =>
                                    clearRecordedSlideTimings("all")
                                  }
                                  data-testid="presentation-ribbon-record-clear-all-timings"
                                >
                                  {t("presentation.record.clearAllTimings")}
                                </DropdownMenuItem>
                              </RibbonDropdownMenuContent>
                            </DropdownMenu>
                          }
                        />
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              controlId="PPTCTL.RIBBON.RECORD.RESET_CAMEO.MENU"
                              variant="large"
                              hasMenuIndicator
                              icon={<RotateCcw size={24} />}
                              disabled={
                                autosave.readOnly ||
                                narrationRecorderBusy ||
                                (!activeSlideCanResetToCameo &&
                                  !deckCanResetAllToCameo)
                              }
                              label={t("presentation.record.resetToCameo")}
                              testId="presentation-ribbon-record-reset-cameo"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent align="start">
                            <DropdownMenuItem
                              disabled={!activeSlideCanResetToCameo}
                              onSelect={() =>
                                resetRecordedSlidesToCameo("current")
                              }
                              data-testid="presentation-ribbon-record-reset-cameo-current"
                            >
                              {t("presentation.record.resetCurrentToCameo")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={!deckCanResetAllToCameo}
                              onSelect={() =>
                                resetRecordedSlidesToCameo("all")
                              }
                              data-testid="presentation-ribbon-record-reset-cameo-all"
                            >
                              {t("presentation.record.resetAllToCameo")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        {activeSlideNarration ? (
                          <audio
                            ref={narrationPreviewAudioRef}
                            src={activeSlideNarration}
                            preload="metadata"
                            data-testid="presentation-record-preview-audio"
                            onLoadedMetadata={(event) => {
                              const duration = Number.isFinite(
                                event.currentTarget.duration,
                              )
                                ? Math.max(0, event.currentTarget.duration)
                                : 0;
                              setNarrationPreviewDurationSeconds(duration);
                              setNarrationPreviewCurrentSeconds(
                                Math.min(
                                  duration,
                                  Math.max(0, event.currentTarget.currentTime),
                                ),
                              );
                            }}
                            onTimeUpdate={(event) =>
                              setNarrationPreviewCurrentSeconds(
                                Math.max(0, event.currentTarget.currentTime),
                              )
                            }
                            onPlay={() => setNarrationPreviewPlaying(true)}
                            onPause={() => setNarrationPreviewPlaying(false)}
                            onEnded={(event) => {
                              setNarrationPreviewPlaying(false);
                              setNarrationPreviewCurrentSeconds(
                                Number.isFinite(event.currentTarget.duration)
                                  ? Math.max(0, event.currentTarget.duration)
                                  : 0,
                              );
                            }}
                            hidden
                          />
                        ) : null}
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.record.optionsGroup")}
                      ribbon="record"
                      size="content"
                      testId="presentation-ribbon-record-options-group"
                    >
                      <div className="note-editor-ribbon-hstack">
                      <OfficeHomeCommandStack>
                        {/* Narration capture is audio-only by construction:
                            startNarrationRecording asks getUserMedia for
                            `video: false`, and the studio camera is a preview
                            (`presentation.record.cameraPreviewOnly`). The
                            mode is therefore pinned on and cannot be switched,
                            and the button says so instead of offering a toggle
                            that claims a choice and ignores the click. */}
                        <OfficeHomeCommand
                          active
                          disabled
                          disabledReason={t("presentation.record.cameraUnavailable")}
                          icon={<Volume2 size={20} />}
                          label={t("presentation.record.audioOnly")}
                          onClick={() => undefined}
                          testId="presentation-ribbon-record-audio-only"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.RECORD.CAMEO"
                          disabled={
                            autosave.readOnly ||
                            insertCapacityReached ||
                            activeSlide.blocks.some(isPresentationCameoBlock)
                          }
                          icon={<IconInsertCameo home size={20} />}
                          displayLabel={t("presentation.insert.cameo")}
                          label={
                            activeSlide.blocks.some(isPresentationCameoBlock)
                              ? `${t("presentation.insert.cameo")} — ${t("presentation.cameo.alreadyOnSlide")}`
                              : t("presentation.insert.cameo")
                          }
                          onClick={() => insertCameo()}
                          testId="presentation-ribbon-record-camera"
                        />
                      </OfficeHomeCommandStack>
                      {/* Four small commands in one stack were 126px tall in
                          an 84px band: two stacks of two. */}
                      <OfficeHomeCommandStack>
                        <OfficeHomeCommand
                          active={teleprompterOpen}
                          ariaExpanded={teleprompterOpen}
                          buttonRef={teleprompterTriggerRef}
                          controlId="PPTCTL.RIBBON.RECORD.TELEPROMPTER"
                          icon={<StickyNote size={20} />}
                          displayLabel={t("presentation.record.teleprompter")}
                          label={t("presentation.record.teleprompter")}
                          onClick={
                            teleprompterOpen
                              ? closeTeleprompter
                              : openTeleprompter
                          }
                          testId="presentation-ribbon-record-teleprompter"
                        />
                        <OfficeHomeCommand
                          controlId="PPTCTL.RIBBON.RECORD.LEARN_MORE"
                          icon={<BookOpen size={20} />}
                          label={t("presentation.record.learnMore")}
                          onClick={openPresentationRecordingHelp}
                          testId="presentation-ribbon-record-learn-more"
                        />
                      </OfficeHomeCommandStack>
                      </div>
                      {/* The recorder's state is announced, not printed in the
                          ribbon: a sentence under the Options commands pushed
                          the group's name out of the ribbon (PowerPoint shows
                          none there; the Record studio pane shows it). An
                          error keeps its notice, which takes a chip's room. */}
                      <span
                        role="status"
                        aria-live="polite"
                        data-testid="presentation-record-status"
                        className="sr-only"
                      >
                        {narrationRecorderError ??
                          t(
                            narrationRecorderStatus === "recording"
                              ? "presentation.record.recording"
                              : narrationRecorderStatus === "paused"
                                ? "presentation.record.paused"
                              : narrationRecorderStatus === "requesting"
                                ? "presentation.record.requesting"
                                : narrationRecorderStatus === "processing"
                                  ? "presentation.record.processing"
                                  : activeSlideNarration
                                    ? "presentation.record.ready"
                                    : "presentation.record.empty",
                          )}
                      </span>
                      {narrationRecorderError ? (
                        <WindowNotice
                          title={t("presentation.record.optionsGroup")}
                          scopeKey={narrationRecorderError}
                        >
                          <span role="alert">{narrationRecorderError}</span>
                        </WindowNotice>
                      ) : null}
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.record.inputGroup")}
                      ribbon="record"
                      size="standard"
                      testId="presentation-ribbon-record-input-group"
                    >
                      <PresentationMicrophoneDeviceMenu
                        controlId="PPTCTL.RIBBON.RECORD.MIC.OPTIONS"
                        currentLabel={selectedNarrationMicrophoneLabel}
                        defaultLabel={t(
                          "presentation.record.microphoneDefault",
                        )}
                        devices={narrationMicrophones}
                        disabled={narrationRecorderBusy}
                        disabledReason={
                          narrationRecorderBusy
                            ? t("presentation.record.microphoneBusy")
                            : undefined
                        }
                        label={t("presentation.record.microphoneDevice")}
                        onSelect={setSelectedNarrationMicrophoneId}
                        selectedId={selectedNarrationMicrophoneId}
                        testId="presentation-ribbon-record-microphone-device"
                      />
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
