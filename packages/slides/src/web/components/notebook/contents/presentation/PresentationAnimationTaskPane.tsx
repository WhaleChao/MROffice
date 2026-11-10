"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Animation Pane task pane, docked by the window.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the render
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { OfficeTaskPane } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  ANIMATION_PANE_MIN_WIDTH,
  ANIMATION_PANE_MAX_WIDTH,
  animationPaneMaxWidth,
  ANIMATION_PANE_WIDTH_STEP,
  AnimationTimelinePanel,
} from "./PresentationAnimationPane";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type {
  AnimationSequenceEditOperation,
  AnimationTimingValue,
  TimelinePreviewFrame,
  TimelineStep,
} from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { PresentationTaskPaneDock } from "../../../../lib/presentation/presentationTaskPaneDock";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationAnimationTaskPaneCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  animationPaneWidth: number;
  setAnimationPaneWidth: Dispatch<SetStateAction<number>>;
  animationPaneTabRef: RefObject<HTMLButtonElement | null>;
  animationPaneRef: RefObject<HTMLElement | null>;
  animationPaneResizePointerRef: RefObject<number | null>;
  selectedAnimationIndex: number | null;
  animationPreview: { readonly slideModelId: string; readonly token: number; readonly frames: readonly TimelinePreviewFrame[]; readonly frameIndex: number; } | null;
  activeSlide: SlideModel;
  taskPaneDock: PresentationTaskPaneDock;
  taskPaneTabsLabel: string;
  timelineSequence: TimelineStep[];
  timelinePreviewFrames: TimelinePreviewFrame[];
  timelineSequenceEditingEnabled: boolean;
  timelineRowEditingEnabled: boolean;
  animationPanePreserveOnly: boolean;
  timelineTimingEditingEnabled: boolean;
  addTimelineStep: (blockId: string) => void;
  updateTimelineStep: (index: number, patch: Partial<TimelineStep>) => void;
  closeAnimationPane: () => void;
  resizeAnimationPaneFromClientX: (clientX: number) => void;
  stepAnimationPaneWidth: (delta: number) => void;
  selectAnimationStep: (index: number) => void;
  editAnimationSequence: (operation: AnimationSequenceEditOperation) => void;
  editAnimationTiming: (index: number, expectedStep: TimelineStep, value: AnimationTimingValue) => void;
  previewAnimations: () => void;
  /** The block's accessible name, the timeline's label for a wordless block. */
  blockAriaName: (block: PresentationBlock) => string;
}

export function PresentationAnimationTaskPane({ ctl }: { ctl: PresentationAnimationTaskPaneCtl }) {
  const {
    t,
    animationPaneWidth,
    setAnimationPaneWidth,
    animationPaneTabRef,
    animationPaneRef,
    animationPaneResizePointerRef,
    selectedAnimationIndex,
    animationPreview,
    activeSlide,
    taskPaneDock,
    taskPaneTabsLabel,
    timelineSequence,
    timelinePreviewFrames,
    timelineSequenceEditingEnabled,
    timelineRowEditingEnabled,
    animationPanePreserveOnly,
    timelineTimingEditingEnabled,
    addTimelineStep,
    updateTimelineStep,
    closeAnimationPane,
    resizeAnimationPaneFromClientX,
    stepAnimationPaneWidth,
    selectAnimationStep,
    editAnimationSequence,
    editAnimationTiming,
    previewAnimations,
    blockAriaName,
  } = ctl;
  return (
          <OfficeTaskPane
            title={t("presentation.animations.animationPane")}
            closeLabel={t("common.close")}
            onClose={closeAnimationPane}
            testIdPrefix="presentation-animation"
            testId="presentation-animation-pane"
            closeTestId="presentation-animation-pane-close"
            closeControlId="PPTCTL.PANE.ANIMATION.CLOSE"
            titleId="presentation-animation-pane-tab"
            tabs={taskPaneDock.tabs("animation", {
              testId: "presentation-animation-pane-tab",
              controlId: "PPTCTL.PANE.ANIMATION.TAB",
              tabRef: animationPaneTabRef,
            })}
            tabsLabel={taskPaneTabsLabel}
            hidden={!taskPaneDock.isFront("animation")}
            dockRight={false}
            className={`${presentationChrome.taskPane} ${presentationChrome.animationPane}`}
            paneRef={animationPaneRef}
            // The pane scrolls as one column under its sticky title row.
            style={{
              display: taskPaneDock.isFront("animation") ? "block" : "none",
              width: animationPaneWidth,
            }}
          >
            <div
              role="separator"
              tabIndex={0}
              aria-label={t("presentation.animations.animationPane")}
              aria-orientation="vertical"
              aria-valuemin={ANIMATION_PANE_MIN_WIDTH}
              aria-valuemax={
                typeof window === "undefined"
                  ? ANIMATION_PANE_MAX_WIDTH
                  : animationPaneMaxWidth(window.innerWidth)
              }
              aria-valuenow={animationPaneWidth}
              aria-valuetext={`${animationPaneWidth}px`}
              className={presentationChrome.taskPaneSplitter}
              data-testid="presentation-animation-pane-splitter"
              data-control-id="PPTCTL.PANE.ANIMATION.SPLITTER"
              onPointerDown={(event) => {
                if (event.button !== 0) return;
                event.preventDefault();
                animationPaneResizePointerRef.current = event.pointerId;
                event.currentTarget.setPointerCapture?.(event.pointerId);
              }}
              onPointerMove={(event) => {
                if (
                  animationPaneResizePointerRef.current !== event.pointerId
                )
                  return;
                event.preventDefault();
                resizeAnimationPaneFromClientX(event.clientX);
              }}
              onPointerUp={(event) => {
                if (
                  animationPaneResizePointerRef.current !== event.pointerId
                )
                  return;
                resizeAnimationPaneFromClientX(event.clientX);
                animationPaneResizePointerRef.current = null;
                if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
                  event.currentTarget.releasePointerCapture(event.pointerId);
                }
              }}
              onPointerCancel={(event) => {
                if (
                  animationPaneResizePointerRef.current !== event.pointerId
                )
                  return;
                animationPaneResizePointerRef.current = null;
                if (event.currentTarget.hasPointerCapture?.(event.pointerId)) {
                  event.currentTarget.releasePointerCapture(event.pointerId);
                }
              }}
              onLostPointerCapture={() => {
                animationPaneResizePointerRef.current = null;
              }}
              onKeyDown={(event) => {
                const step = event.shiftKey
                  ? ANIMATION_PANE_WIDTH_STEP * 2
                  : ANIMATION_PANE_WIDTH_STEP;
                if (event.key === "ArrowLeft") {
                  event.preventDefault();
                  stepAnimationPaneWidth(step);
                } else if (event.key === "ArrowRight") {
                  event.preventDefault();
                  stepAnimationPaneWidth(-step);
                } else if (event.key === "Home") {
                  event.preventDefault();
                  setAnimationPaneWidth(ANIMATION_PANE_MIN_WIDTH);
                } else if (event.key === "End") {
                  event.preventDefault();
                  setAnimationPaneWidth(
                    typeof window === "undefined"
                      ? ANIMATION_PANE_MAX_WIDTH
                      : animationPaneMaxWidth(window.innerWidth),
                  );
                }
              }}
            />
            {animationPanePreserveOnly ? (
              <p
                role="status"
                aria-live="polite"
                className={presentationChrome.taskPaneNotice}
                data-testid="presentation-animation-pane-warning"
                data-control-id="PPTCTL.PANE.ANIMATION.WARNING"
              >
                {t("presentation.animations.importedPatchbackLimited")}
              </p>
            ) : null}
            <div className={presentationChrome.animationPaneBody}>
              <AnimationTimelinePanel
                sequence={timelineSequence}
                blocks={activeSlide.blocks}
                t={t}
                previewEnabled={timelinePreviewFrames.length > 0}
                previewActive={
                  animationPreview?.slideModelId === activeSlide.id
                }
                selectedIndex={selectedAnimationIndex}
                rowEditingEnabled={timelineRowEditingEnabled}
                timingEditingEnabled={timelineTimingEditingEnabled}
                sequenceEditingEnabled={timelineSequenceEditingEnabled}
                onAddStep={addTimelineStep}
                onPreview={previewAnimations}
                onSelectStep={selectAnimationStep}
                onEditSequence={editAnimationSequence}
                onUpdateStep={updateTimelineStep}
                onUpdateTiming={editAnimationTiming}
                blockFallbackName={blockAriaName}
              />
            </div>
          </OfficeTaskPane>
  );
}
