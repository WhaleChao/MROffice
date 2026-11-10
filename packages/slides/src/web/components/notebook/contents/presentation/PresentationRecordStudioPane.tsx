"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The recording studio (record/pause/stop, slide navigation, ink, cameo,
 * teleprompter, notes and the take preview), rendered while it is open.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  X,
  Timer,
  Pause,
  Play,
  RotateCcw,
  Pointer,
  Pen,
  Highlighter,
  Eraser,
  Trash2,
  ChevronLeft,
  Mic,
  LayoutTemplate,
  StickyNote,
  Circle,
  Square,
  ChevronRight,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { formatElapsed } from "../../../../lib/presentation-presenter";
import { PresentationAnnotationLayer } from "./PresentationAnnotationLayer";
import { PresentationCameoSurface, PresentationMicrophoneDeviceMenu } from "./PresentationCameo";
import { RECORD_STUDIO_INK_COLORS } from "../../../../lib/presentation/ink-annotation";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { AnnotationPoint, AnnotationTool, InkStroke } from "@mrnote/core/web/lib/presentation-annotation.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationRecordStudioPaneCtl {
  pageId: string;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  slides: SlideModel[];
  activeIndex: number;
  narrationRecorderStatus: "idle" | "requesting" | "recording" | "paused" | "processing";
  narrationRecorderError: string | null;
  recordStudioLayout: "teleprompter" | "presenter" | "slide";
  setRecordStudioLayout: Dispatch<SetStateAction<"teleprompter" | "presenter" | "slide">>;
  recordStudioElapsedMs: number;
  recordStudioTotalElapsedMs: number;
  recordStudioPanelRef: RefObject<HTMLElement | null>;
  recordStudioStageRef: RefObject<HTMLDivElement | null>;
  recordStudioTeleprompterScrollRef: RefObject<HTMLDivElement | null>;
  recordStudioTeleprompterRunning: boolean;
  setRecordStudioTeleprompterRunning: Dispatch<SetStateAction<boolean>>;
  recordStudioAnnotationColor: string;
  setRecordStudioAnnotationColor: Dispatch<SetStateAction<string>>;
  recordStudioInkSurfaceRect: { readonly left: number; readonly top: number; readonly width: number; readonly height: number; } | null;
  resetRecordStudioElapsed: () => void;
  bankRecordStudioElapsed: () => void;
  narrationMicrophones: readonly Readonly<{ deviceId: string; label: string; }>[];
  selectedNarrationMicrophoneId: string;
  setSelectedNarrationMicrophoneId: Dispatch<SetStateAction<string>>;
  narrationMicrophonePermissionStatus: "idle" | "requesting" | "ready" | "error";
  narrationPreviewPlaying: boolean;
  narrationPreviewCurrentSeconds: number;
  narrationPreviewDurationSeconds: number;
  toggleNarrationPreview: () => void;
  seekNarrationPreview: (seconds: number) => void;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  activeSlide: SlideModel;
  requestNarrationMicrophoneAccess: () => Promise<void>;
  startNarrationRecording: (target?: "narration" | "audio-clip") => Promise<void>;
  stopNarrationRecording: () => void;
  toggleNarrationRecordingPaused: () => void;
  teleprompterOpen: boolean;
  teleprompterSpeed: number;
  setTeleprompterSpeed: Dispatch<SetStateAction<number>>;
  activeTeleprompterNotes: string;
  closeTeleprompter: () => void;
  openTeleprompter: () => void;
  closeRecordStudio: () => void;
  moveRecordStudioSlide: (direction: -1 | 1) => void;
  recordStudioCanDrawInk: boolean;
  recordStudioCanEraseInk: boolean;
  recordStudioActiveAnnotationTool: AnnotationTool | null;
  recordStudioDraftStrokes: readonly InkStroke[];
  selectRecordStudioAnnotationTool: (tool: AnnotationTool) => void;
  startRecordStudioInkStroke: (point: AnnotationPoint) => void;
  appendRecordStudioInkStroke: (point: AnnotationPoint) => void;
  finishRecordStudioInkStroke: () => void;
  eraseRecordStudioInkAt: (point: AnnotationPoint) => void;
  finishRecordStudioInkErase: () => void;
  clearRecordStudioInk: () => void;
  resetActiveSlideRecordingToCameo: () => boolean;
  narrationRecorderSupported: boolean;
  activeSlideNarration: string | null;
  activeSlideCanResetToCameo: boolean;
  narrationRecorderBusy: boolean;
  narrationRecorderCanTogglePause: boolean;
  narrationMicrophonePermissionLabel: string;
  selectedNarrationMicrophoneLabel: string;
}

export function PresentationRecordStudioPane({ ctl }: { ctl: PresentationRecordStudioPaneCtl }) {
  const {
    pageId,
    t,
    slides,
    activeIndex,
    narrationRecorderStatus,
    narrationRecorderError,
    recordStudioLayout,
    setRecordStudioLayout,
    recordStudioElapsedMs,
    recordStudioTotalElapsedMs,
    recordStudioPanelRef,
    recordStudioStageRef,
    recordStudioTeleprompterScrollRef,
    recordStudioTeleprompterRunning,
    setRecordStudioTeleprompterRunning,
    recordStudioAnnotationColor,
    setRecordStudioAnnotationColor,
    recordStudioInkSurfaceRect,
    resetRecordStudioElapsed,
    bankRecordStudioElapsed,
    narrationMicrophones,
    selectedNarrationMicrophoneId,
    setSelectedNarrationMicrophoneId,
    narrationMicrophonePermissionStatus,
    narrationPreviewPlaying,
    narrationPreviewCurrentSeconds,
    narrationPreviewDurationSeconds,
    toggleNarrationPreview,
    seekNarrationPreview,
    autosave,
    activeSlide,
    requestNarrationMicrophoneAccess,
    startNarrationRecording,
    stopNarrationRecording,
    toggleNarrationRecordingPaused,
    teleprompterOpen,
    teleprompterSpeed,
    setTeleprompterSpeed,
    activeTeleprompterNotes,
    closeTeleprompter,
    openTeleprompter,
    closeRecordStudio,
    moveRecordStudioSlide,
    recordStudioCanDrawInk,
    recordStudioCanEraseInk,
    recordStudioActiveAnnotationTool,
    recordStudioDraftStrokes,
    selectRecordStudioAnnotationTool,
    startRecordStudioInkStroke,
    appendRecordStudioInkStroke,
    finishRecordStudioInkStroke,
    eraseRecordStudioInkAt,
    finishRecordStudioInkErase,
    clearRecordStudioInk,
    resetActiveSlideRecordingToCameo,
    narrationRecorderSupported,
    activeSlideNarration,
    activeSlideCanResetToCameo,
    narrationRecorderBusy,
    narrationRecorderCanTogglePause,
    narrationMicrophonePermissionLabel,
    selectedNarrationMicrophoneLabel,
  } = ctl;
  return (
            <aside
              ref={recordStudioPanelRef}
              className={presentationChrome.recordStudioPanel}
              role="dialog"
              aria-modal="false"
              aria-labelledby={`presentation-record-studio-title-${pageId}`}
              tabIndex={-1}
              data-recorder-status={narrationRecorderStatus}
              data-testid="presentation-record-studio"
              onKeyDown={(event) => {
                if (event.key !== "Escape") return;
                event.preventDefault();
                event.stopPropagation();
                closeRecordStudio();
              }}
            >
              <header className={presentationChrome.recordStudioHeader}>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.EXIT"
                  data-testid="presentation-record-studio-exit"
                  disabled={narrationRecorderStatus === "processing"}
                  onClick={closeRecordStudio}
                >
                  <X size={17} aria-hidden="true" />
                  {t("presentation.record.studioExit")}
                </button>
                <strong id={`presentation-record-studio-title-${pageId}`}>
                  {t("presentation.record.studioTitle")}
                </strong>
                <output
                  className={presentationChrome.recordStudioTimer}
                  aria-label={t("presentation.record.studioTimerValue", {
                    current: formatElapsed(recordStudioElapsedMs),
                    total: formatElapsed(recordStudioTotalElapsedMs),
                  })}
                  data-control-id="PPTCTL.RECORD_STUDIO.TIMER"
                  data-current-ms={Math.round(recordStudioElapsedMs)}
                  data-total-ms={Math.round(recordStudioTotalElapsedMs)}
                  data-testid="presentation-record-studio-timer"
                >
                  <Timer size={16} aria-hidden="true" />
                  {formatElapsed(recordStudioElapsedMs)} / {formatElapsed(recordStudioTotalElapsedMs)}
                </output>
              </header>
              <div
                ref={recordStudioStageRef}
                className={presentationChrome.recordStudioStage}
                aria-label={t("presentation.record.studioStage")}
              >
                {recordStudioActiveAnnotationTool ? (
                  <div
                    className={presentationChrome.recordStudioInkSurface}
                    data-tool={recordStudioActiveAnnotationTool}
                    data-testid="presentation-record-studio-ink-surface"
                    style={
                      recordStudioInkSurfaceRect
                        ? {
                            left: recordStudioInkSurfaceRect.left,
                            top: recordStudioInkSurfaceRect.top,
                            width: recordStudioInkSurfaceRect.width,
                            height: recordStudioInkSurfaceRect.height,
                          }
                        : { inset: 0 }
                    }
                  >
                    <PresentationAnnotationLayer
                      tool={recordStudioActiveAnnotationTool}
                      color={recordStudioAnnotationColor}
                      strokes={recordStudioDraftStrokes}
                      onStrokeStart={startRecordStudioInkStroke}
                      onStrokeAppend={appendRecordStudioInkStroke}
                      onStrokeFinish={finishRecordStudioInkStroke}
                      onErase={eraseRecordStudioInkAt}
                      onEraseFinish={finishRecordStudioInkErase}
                      ariaLabel={t("presentation.annotationLayerLabel")}
                    />
                  </div>
                ) : null}
                <span>{activeSlide.title || t("presentation.untitledSlide")}</span>
                {recordStudioLayout !== "slide" ? (
                  <section
                    className={presentationChrome.recordStudioLayoutPane}
                    data-layout={recordStudioLayout}
                    data-testid="presentation-record-studio-layout-pane"
                    tabIndex={0}
                    aria-label={t(
                      recordStudioLayout === "presenter"
                        ? "presentation.record.studioLayoutPresenter"
                        : "presentation.record.studioLayoutTeleprompter",
                    )}
                  >
                    <strong>
                      {t(
                        recordStudioLayout === "presenter"
                          ? "presentation.record.studioLayoutPresenter"
                          : "presentation.record.studioLayoutTeleprompter",
                      )}
                    </strong>
                    <div
                      ref={
                        recordStudioLayout === "teleprompter"
                          ? recordStudioTeleprompterScrollRef
                          : undefined
                      }
                      className={presentationChrome.recordStudioLayoutNotes}
                      data-testid="presentation-record-studio-layout-notes"
                    >
                      <p>
                        {activeTeleprompterNotes.trim() ||
                          t("presentation.record.teleprompterEmpty")}
                      </p>
                      {recordStudioLayout === "presenter" ? (
                        <small>
                          {slides[activeIndex + 1]
                            ? t("presentation.record.studioLayoutNext", {
                                title:
                                  slides[activeIndex + 1].title ||
                                  t("presentation.untitledSlide"),
                              })
                            : t("presentation.record.studioLayoutEnd")}
                        </small>
                      ) : null}
                    </div>
                    {recordStudioLayout === "teleprompter" ? (
                      <div
                        className={
                          presentationChrome.recordStudioTeleprompterControls
                        }
                        data-testid="presentation-record-studio-teleprompter-controls"
                      >
                        <button
                          type="button"
                          data-control-id="PPTCTL.RECORD_STUDIO.TELEPROMPTER.AUTOSCROLL"
                          data-testid="presentation-record-studio-teleprompter-autoscroll"
                          disabled={!activeTeleprompterNotes.trim()}
                          aria-pressed={recordStudioTeleprompterRunning}
                          onClick={() =>
                            setRecordStudioTeleprompterRunning(
                              (running) => !running,
                            )
                          }
                        >
                          {recordStudioTeleprompterRunning ? (
                            <Pause size={15} aria-hidden="true" />
                          ) : (
                            <Play size={15} aria-hidden="true" />
                          )}
                          {t(
                            recordStudioTeleprompterRunning
                              ? "presentation.record.teleprompterPause"
                              : "presentation.record.teleprompterPlay",
                          )}
                        </button>
                        <label>
                          <span>
                            {t("presentation.record.teleprompterSpeed")}
                          </span>
                          <input
                            type="range"
                            min={1}
                            max={5}
                            step={1}
                            value={teleprompterSpeed}
                            data-control-id="PPTCTL.RECORD_STUDIO.TELEPROMPTER.SPEED"
                            data-testid="presentation-record-studio-teleprompter-speed"
                            aria-valuetext={t(
                              "presentation.record.teleprompterSpeedValue",
                              { value: teleprompterSpeed },
                            )}
                            onChange={(event) =>
                              setTeleprompterSpeed(
                                Math.min(
                                  5,
                                  Math.max(1, Number(event.target.value)),
                                ),
                              )
                            }
                          />
                        </label>
                        <button
                          type="button"
                          data-control-id="PPTCTL.RECORD_STUDIO.TELEPROMPTER.RESET"
                          data-testid="presentation-record-studio-teleprompter-reset"
                          onClick={() => {
                            setRecordStudioTeleprompterRunning(false);
                            if (recordStudioTeleprompterScrollRef.current) {
                              recordStudioTeleprompterScrollRef.current.scrollTop =
                                0;
                            }
                          }}
                        >
                          <RotateCcw size={15} aria-hidden="true" />
                          {t("presentation.record.teleprompterReset")}
                        </button>
                      </div>
                    ) : null}
                  </section>
                ) : null}
                <div className={presentationChrome.recordStudioCamera}>
                  <div className={presentationChrome.recordStudioCameraFrame}>
                    <PresentationCameoSurface
                      interactive={
                        !autosave.readOnly && narrationRecorderStatus === "idle"
                      }
                      scope="record-studio"
                      t={t}
                    />
                  </div>
                  <small className={presentationChrome.recordStudioCameraHint}>
                    {t("presentation.record.cameraPreviewOnly")}
                  </small>
                </div>
              </div>
              <div
                className={presentationChrome.recordStudioInkToolbar}
                role="toolbar"
                aria-label={t("presentation.annotationLayerLabel")}
                data-testid="presentation-record-studio-ink-toolbar"
              >
                {(
                  [
                    {
                      tool: "laser",
                      controlId: "PPTCTL.RECORD_STUDIO.POINTER.LASER",
                      label: t("presentation.annotationLaser"),
                      icon: <Pointer size={17} aria-hidden="true" />,
                      disabled: false,
                    },
                    {
                      tool: "pen",
                      controlId: "PPTCTL.RECORD_STUDIO.POINTER.PEN",
                      label: t("presentation.annotationPen"),
                      icon: <Pen size={17} aria-hidden="true" />,
                      disabled: !recordStudioCanDrawInk,
                    },
                    {
                      tool: "highlighter",
                      controlId: "PPTCTL.RECORD_STUDIO.POINTER.HIGHLIGHTER",
                      label: t("presentation.annotationHighlighter"),
                      icon: <Highlighter size={17} aria-hidden="true" />,
                      disabled: !recordStudioCanDrawInk,
                    },
                    {
                      tool: "eraser",
                      controlId: "PPTCTL.RECORD_STUDIO.POINTER.ERASER",
                      label: t("presentation.annotationEraser"),
                      icon: <Eraser size={17} aria-hidden="true" />,
                      disabled: !recordStudioCanEraseInk,
                    },
                  ] as const
                ).map(({ tool, controlId, label, icon, disabled }) => (
                  <button
                    key={tool}
                    type="button"
                    className={presentationChrome.recordStudioInkTool}
                    data-control-id={controlId}
                    data-testid={`presentation-record-studio-ink-${tool}`}
                    disabled={disabled}
                    aria-label={label}
                    title={label}
                    aria-pressed={recordStudioActiveAnnotationTool === tool}
                    onClick={() => selectRecordStudioAnnotationTool(tool)}
                  >
                    {icon}
                  </button>
                ))}
                {recordStudioActiveAnnotationTool !== null &&
                recordStudioActiveAnnotationTool !== "eraser" ? (
                  <div
                    className={presentationChrome.recordStudioInkColors}
                    data-testid="presentation-record-studio-ink-colors"
                  >
                    {RECORD_STUDIO_INK_COLORS.map((color, index) => (
                      <button
                        key={color.id}
                        type="button"
                        className={presentationChrome.recordStudioInkColor}
                        data-control-id={`PPTCTL.RECORD_STUDIO.COLOR.${color.id
                          .replace("-", "_")
                          .toUpperCase()}`}
                        data-testid={`presentation-record-studio-ink-color-${color.id}`}
                        aria-label={`${t("presentation.annotationColor")} ${index + 1}`}
                        title={`${t("presentation.annotationColor")} ${index + 1}`}
                        aria-pressed={
                          recordStudioAnnotationColor === color.value
                        }
                        onClick={() =>
                          setRecordStudioAnnotationColor(color.value)
                        }
                      >
                        <span
                          aria-hidden="true"
                          style={{ backgroundColor: color.value }}
                        />
                      </button>
                    ))}
                  </div>
                ) : null}
                <button
                  type="button"
                  className={presentationChrome.recordStudioInkTool}
                  data-control-id="PPTCTL.RECORD_STUDIO.POINTER.ERASE_ALL"
                  data-testid="presentation-record-studio-ink-clear"
                  disabled={!recordStudioCanEraseInk}
                  aria-label={t("presentation.annotationClear")}
                  title={t("presentation.annotationClear")}
                  onClick={clearRecordStudioInk}
                >
                  <Trash2 size={17} aria-hidden="true" />
                </button>
              </div>
              <div
                className={presentationChrome.recordStudioPreviewBar}
                data-testid="presentation-record-studio-preview-bar"
              >
                <button
                  type="button"
                  className={presentationChrome.recordStudioPreviewButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.PREVIEW"
                  data-testid="presentation-record-studio-preview"
                  disabled={
                    activeSlideNarration === null || narrationRecorderBusy
                  }
                  aria-pressed={narrationPreviewPlaying}
                  onClick={toggleNarrationPreview}
                >
                  {narrationPreviewPlaying ? (
                    <Pause size={16} aria-hidden="true" />
                  ) : (
                    <Play size={16} aria-hidden="true" />
                  )}
                  {t(
                    narrationPreviewPlaying
                      ? "presentation.record.studioPreviewPause"
                      : "presentation.record.studioPreviewPlay",
                  )}
                </button>
                <input
                  type="range"
                  min={0}
                  max={Math.max(0, narrationPreviewDurationSeconds)}
                  step={0.1}
                  value={Math.min(
                    narrationPreviewDurationSeconds,
                    narrationPreviewCurrentSeconds,
                  )}
                  disabled={
                    activeSlideNarration === null ||
                    narrationRecorderBusy ||
                    narrationPreviewDurationSeconds <= 0
                  }
                  aria-label={t("presentation.record.studioPreviewSeek")}
                  aria-valuetext={t(
                    "presentation.record.studioPreviewTime",
                    {
                      current: formatElapsed(
                        narrationPreviewCurrentSeconds * 1_000,
                      ),
                      total: formatElapsed(
                        narrationPreviewDurationSeconds * 1_000,
                      ),
                    },
                  )}
                  data-control-id="PPTCTL.RECORD_STUDIO.PREVIEW.SEEK"
                  data-testid="presentation-record-studio-preview-seek"
                  onChange={(event) =>
                    seekNarrationPreview(Number(event.target.value))
                  }
                />
                <output data-testid="presentation-record-studio-preview-time">
                  {formatElapsed(narrationPreviewCurrentSeconds * 1_000)} /{" "}
                  {formatElapsed(narrationPreviewDurationSeconds * 1_000)}
                </output>
              </div>
              <footer className={presentationChrome.recordStudioControls}>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.PREVIOUS"
                  data-testid="presentation-record-studio-previous"
                  disabled={
                    activeIndex <= 0 || narrationRecorderStatus !== "idle"
                  }
                  onClick={() => moveRecordStudioSlide(-1)}
                >
                  <ChevronLeft size={18} aria-hidden="true" />
                  {t("presentation.record.studioPrevious")}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.RESET_CAMEO"
                  data-testid="presentation-record-studio-reset-cameo"
                  disabled={
                    autosave.readOnly ||
                    narrationRecorderStatus !== "idle" ||
                    !activeSlideCanResetToCameo
                  }
                  onClick={resetActiveSlideRecordingToCameo}
                >
                  <RotateCcw size={17} aria-hidden="true" />
                  {t("presentation.record.resetToCameo")}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.MIC.PERMISSION"
                  data-testid="presentation-record-studio-microphone-permission"
                  disabled={
                    autosave.readOnly ||
                    narrationRecorderStatus !== "idle" ||
                    narrationMicrophonePermissionStatus === "requesting" ||
                    !narrationRecorderSupported
                  }
                  aria-pressed={
                    narrationMicrophonePermissionStatus === "ready"
                  }
                  title={
                    narrationRecorderSupported
                      ? narrationMicrophonePermissionLabel
                      : t(
                          "presentation.record.compatibleFormatUnavailable",
                        )
                  }
                  onClick={() => void requestNarrationMicrophoneAccess()}
                >
                  <Mic size={17} aria-hidden="true" />
                  {narrationMicrophonePermissionLabel}
                </button>
                <PresentationMicrophoneDeviceMenu
                  className={presentationChrome.recordStudioButton}
                  controlId="PPTCTL.RECORD_STUDIO.MIC.OPTIONS"
                  currentLabel={selectedNarrationMicrophoneLabel}
                  defaultLabel={t("presentation.record.microphoneDefault")}
                  devices={narrationMicrophones}
                  disabled={
                    narrationRecorderBusy ||
                    narrationMicrophonePermissionStatus === "requesting"
                  }
                  disabledReason={
                    narrationRecorderBusy ||
                    narrationMicrophonePermissionStatus === "requesting"
                      ? t("presentation.record.microphoneBusy")
                      : undefined
                  }
                  label={t("presentation.record.microphoneDevice")}
                  onSelect={setSelectedNarrationMicrophoneId}
                  selectedId={selectedNarrationMicrophoneId}
                  testId="presentation-record-studio-microphone-options"
                />
                <DropdownMenu modal={false}>
                  <DropdownMenuTrigger asChild>
                    <OfficeHomeMenuTrigger
                      className={presentationChrome.recordStudioButton}
                      controlId="PPTCTL.RECORD_STUDIO.LAYOUT"
                      disabled={narrationRecorderStatus === "processing"}
                      displayLabel={t(
                        recordStudioLayout === "teleprompter"
                          ? "presentation.record.studioLayoutTeleprompter"
                          : recordStudioLayout === "presenter"
                            ? "presentation.record.studioLayoutPresenter"
                            : "presentation.record.studioLayoutSlide",
                      )}
                      hasMenuIndicator
                      icon={<LayoutTemplate size={17} />}
                      label={t("presentation.record.studioLayout")}
                      testId="presentation-record-studio-layout"
                    />
                  </DropdownMenuTrigger>
                  <RibbonDropdownMenuContent align="start">
                    <DropdownMenuLabel>
                      {t("presentation.record.studioLayout")}
                    </DropdownMenuLabel>
                    <DropdownMenuRadioGroup
                      value={recordStudioLayout}
                      onValueChange={(value) =>
                        setRecordStudioLayout(
                          value as "teleprompter" | "presenter" | "slide",
                        )
                      }
                    >
                      {(
                        [
                          "teleprompter",
                          "presenter",
                          "slide",
                        ] as const
                      ).map((layout) => (
                        <DropdownMenuRadioItem
                          key={layout}
                          value={layout}
                          onClick={() => setRecordStudioLayout(layout)}
                          onSelect={() => setRecordStudioLayout(layout)}
                          data-testid={`presentation-record-studio-layout-${layout}`}
                        >
                          {t(
                            layout === "teleprompter"
                              ? "presentation.record.studioLayoutTeleprompter"
                              : layout === "presenter"
                                ? "presentation.record.studioLayoutPresenter"
                                : "presentation.record.studioLayoutSlide",
                          )}
                        </DropdownMenuRadioItem>
                      ))}
                    </DropdownMenuRadioGroup>
                  </RibbonDropdownMenuContent>
                </DropdownMenu>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.TELEPROMPTER"
                  data-testid="presentation-record-studio-teleprompter"
                  aria-pressed={teleprompterOpen}
                  onClick={
                    teleprompterOpen ? closeTeleprompter : openTeleprompter
                  }
                >
                  <StickyNote size={17} aria-hidden="true" />
                  {t("presentation.record.teleprompter")}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioPrimaryButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.RECORD"
                  data-testid="presentation-record-studio-record"
                  disabled={
                    autosave.readOnly ||
                    narrationRecorderStatus !== "idle" ||
                    !narrationRecorderSupported
                  }
                  onClick={() => {
                    resetRecordStudioElapsed();
                    void startNarrationRecording();
                  }}
                >
                  <Circle size={17} fill="currentColor" aria-hidden="true" />
                  {t("presentation.record.studioRecord")}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.PAUSE"
                  data-testid="presentation-record-studio-pause"
                  disabled={!narrationRecorderCanTogglePause}
                  aria-pressed={narrationRecorderStatus === "paused"}
                  onClick={toggleNarrationRecordingPaused}
                >
                  {narrationRecorderStatus === "paused" ? (
                    <Play size={16} aria-hidden="true" />
                  ) : (
                    <Pause size={16} aria-hidden="true" />
                  )}
                  {t(
                    narrationRecorderStatus === "paused"
                      ? "presentation.record.resume"
                      : "presentation.record.pause",
                  )}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.STOP"
                  data-testid="presentation-record-studio-stop"
                  disabled={
                    narrationRecorderStatus !== "recording" &&
                    narrationRecorderStatus !== "paused"
                  }
                  onClick={() => {
                    bankRecordStudioElapsed();
                    stopNarrationRecording();
                  }}
                >
                  <Square size={16} fill="currentColor" aria-hidden="true" />
                  {t("presentation.record.stop")}
                </button>
                <button
                  type="button"
                  className={presentationChrome.recordStudioButton}
                  data-control-id="PPTCTL.RECORD_STUDIO.NEXT"
                  data-testid="presentation-record-studio-next"
                  disabled={
                    activeIndex >= slides.length - 1 ||
                    narrationRecorderStatus !== "idle"
                  }
                  onClick={() => moveRecordStudioSlide(1)}
                >
                  {t("presentation.record.studioNext")}
                  <ChevronRight size={18} aria-hidden="true" />
                </button>
                <span
                  className={presentationChrome.recordStudioCounter}
                  data-current={activeIndex + 1}
                  data-total={slides.length}
                  data-testid="presentation-record-studio-counter"
                >
                  {t("presentation.record.studioCounter", {
                    current: activeIndex + 1,
                    total: slides.length,
                  })}
                </span>
              </footer>
              <span
                className={presentationChrome.recordStudioStatus}
                role="status"
                aria-live="polite"
                data-testid="presentation-record-studio-status"
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
            </aside>
  );
}
