"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Animation Pane: the timeline list with its timing controls (start,
 * duration, delay, repeat) and the pane width bounds.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import {
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";
import {
  Check,
  ChevronDown,
  ChevronUp,
  Play,
  X,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { IconTimingStart, IconTimingDuration, IconTimingDelay } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  OfficeHomeCommandRow,
  OfficeHomeCommandStack,
  OfficeHomeMenuTrigger,
  OfficeHomeNumberField,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS,
  PRESENTATION_ANIMATION_TIMING_CONTROL_IDS,
} from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import { presentationAnimationEffectTargetEligible } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import {
  SCRATCH_ANIMATION_MOTION_PATHS,
  TIMELINE_TRIGGERS,
  WEB_AUTHORABLE_TIMELINE_ANIMATIONS,
  MAX_TIMELINE_STEPS,
  formatAnimationTimingSeconds,
  parseAnimationTimingSeconds,
  copyTimelineStep,
  type AnimationSequenceEditOperation,
  type AnimationTimingValue,
  type TimelineAnimation,
  type TimelineStep,
} from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import { presetMotionPath, type MotionPathPreset } from "@mrnote/core/web/lib/presentation-motion-path.ts";
import { DropdownMenu, DropdownMenuItem, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import type {
  PresentationBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { presentationBlockLabel } from "../../../../lib/presentation/deck-text";
import { blockOrderButtonStyle } from "../../../../lib/presentation/window-chrome-styles";

function AnimationTimingNumberInput({
  field,
  index,
  step,
  disabled,
  label,
  controlId,
  ribbon = false,
  empty = false,
  testId,
  onCommit,
}: {
  field: "delayMs" | "durationMs";
  /** No effect selected: PowerPoint leaves Duration and Delay blank. */
  empty?: boolean;
  index: number;
  step: TimelineStep;
  disabled: boolean;
  label: string;
  controlId: string;
  ribbon?: boolean;
  testId: string;
  onCommit: (
    index: number,
    expectedStep: TimelineStep,
    value: AnimationTimingValue,
  ) => void;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const editRef = useRef<{
    index: number;
    expectedStep: TimelineStep;
    value: string;
  } | null>(null);
  const skipBlurRef = useRef(false);
  const begin = () => {
    const value = formatAnimationTimingSeconds(step[field]);
    editRef.current = {
      index,
      expectedStep: copyTimelineStep(step),
      value,
    };
    setDraft(value);
  };
  const change = (value: string) => {
    if (!editRef.current) begin();
    if (editRef.current) editRef.current = { ...editRef.current, value };
    setDraft(value);
  };
  const finish = (commit: boolean) => {
    const edit = editRef.current;
    editRef.current = null;
    setDraft(null);
    if (!commit || !edit) return;
    const milliseconds = parseAnimationTimingSeconds(edit.value, field);
    if (milliseconds === undefined) return;
    onCommit(edit.index, edit.expectedStep, {
      trigger: edit.expectedStep.trigger,
      delayMs:
        field === "delayMs" ? milliseconds : edit.expectedStep.delayMs,
      durationMs:
        field === "durationMs"
          ? milliseconds
          : edit.expectedStep.durationMs,
    });
  };
  const blur = () => {
    if (skipBlurRef.current) {
      skipBlurRef.current = false;
      return;
    }
    finish(true);
  };
  const keyDown = (event: ReactKeyboardEvent<HTMLInputElement>) => {
    if (event.key !== "Escape" && event.key !== "Enter") return;
    event.preventDefault();
    skipBlurRef.current = true;
    finish(event.key === "Enter");
    event.currentTarget.blur();
  };
  const value =
    draft ?? (empty ? "" : formatAnimationTimingSeconds(step[field]));
  if (ribbon) {
    return (
      <OfficeHomeNumberField
        appearance="native-spinner"
        controlId={controlId}
        disabled={disabled}
        inputMode="decimal"
        label={label}
        max={60}
        min={0}
        onBlur={blur}
        onChange={(event) => change(event.target.value)}
        onFocus={begin}
        onKeyDown={keyDown}
        step={0.05}
        testId={testId}
        value={value}
        data-unit="seconds"
      />
    );
  }
  return (
    <span className="presentation-window__timeline-num">
      <span className="presentation-window__timeline-sr">{label}</span>
      <input
        type="number"
        disabled={disabled}
        min={0}
        max={60}
        step={0.05}
        inputMode="decimal"
        value={value}
        onFocus={begin}
        onChange={(event) => change(event.target.value)}
        onBlur={blur}
        onKeyDown={keyDown}
        data-unit="seconds"
        data-control-id={controlId}
        data-testid={testId}
        aria-label={label}
        title={label}
      />
    </span>
  );
}

export function AnimationTimingControls({
  scope,
  index,
  step,
  disabled,
  empty = false,
  t,
  onCommit,
}: {
  scope: "ribbon" | "pane";
  index: number;
  step: TimelineStep;
  disabled: boolean;
  empty?: boolean;
  t: (key: string, values?: Record<string, string | number>) => string;
  onCommit: (
    index: number,
    expectedStep: TimelineStep,
    value: AnimationTimingValue,
  ) => void;
}) {
  const startControlId =
    scope === "ribbon"
      ? PRESENTATION_ANIMATION_TIMING_CONTROL_IDS[0]
      : "PPTCTL.PANE.ANIMATION.SETTINGS.START";
  const durationControlId =
    scope === "ribbon"
      ? PRESENTATION_ANIMATION_TIMING_CONTROL_IDS[1]
      : "PPTCTL.PANE.ANIMATION.SETTINGS.DURATION";
  const delayControlId =
    scope === "ribbon"
      ? "PPTCTL.RIBBON.ANIMATIONS.DELAY.FIELD"
      : "PPTCTL.PANE.ANIMATION.SETTINGS.DELAY";
  const startTestId =
    scope === "ribbon"
      ? "presentation-ribbon-animation-start"
      : "presentation-timeline-trigger";
  const durationTestId =
    scope === "ribbon"
      ? "presentation-ribbon-animation-duration"
      : "presentation-timeline-duration";
  const delayTestId =
    scope === "ribbon"
      ? "presentation-ribbon-animation-delay"
      : "presentation-timeline-delay";
  const selectedTriggerLabel = t(
    `presentation.timeline.trigger.${step.trigger}` as "presentation.timeline.trigger.on-click",
  );
  const triggerMenu = (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <OfficeHomeMenuTrigger
          disabled={disabled}
          displayLabel={selectedTriggerLabel}
          hasMenuIndicator
          icon={scope === "ribbon" ? <IconTimingStart /> : <Play size={16} />}
          label={`${t("presentation.timelineTrigger")}: ${selectedTriggerLabel}`}
          controlId={startControlId}
          testId={startTestId}
        />
      </DropdownMenuTrigger>
      <RibbonDropdownMenuContent align="start">
        {TIMELINE_TRIGGERS.map((value) => (
          <DropdownMenuItem
            key={value}
            onSelect={() =>
              onCommit(index, copyTimelineStep(step), {
                trigger: value,
                delayMs: step.delayMs,
                durationMs: step.durationMs,
              })
            }
          >
            {step.trigger === value ? <Check size={15} /> : null}
            {t(
              `presentation.timeline.trigger.${value}` as "presentation.timeline.trigger.on-click",
            )}
          </DropdownMenuItem>
        ))}
      </RibbonDropdownMenuContent>
    </DropdownMenu>
  );
  if (scope === "ribbon") {
    return (
      <OfficeHomeCommandStack>
        <OfficeHomeCommandRow>{triggerMenu}</OfficeHomeCommandRow>
        <OfficeHomeCommandRow className={presentationChrome.effectsTimingRow}>
          <IconTimingDuration />
          <span>{t("presentation.timelineDuration")}</span>
          <AnimationTimingNumberInput
            ribbon
            empty={empty}
            field="durationMs"
            index={index}
            step={step}
            disabled={disabled}
            label={t("presentation.timelineDuration")}
            controlId={durationControlId}
            testId={durationTestId}
            onCommit={onCommit}
          />
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow className={presentationChrome.effectsTimingRow}>
          <IconTimingDelay />
          <span>{t("presentation.timelineDelay")}</span>
          <AnimationTimingNumberInput
            ribbon
            empty={empty}
            field="delayMs"
            index={index}
            step={step}
            disabled={disabled}
            label={t("presentation.timelineDelay")}
            controlId={delayControlId}
            testId={delayTestId}
            onCommit={onCommit}
          />
        </OfficeHomeCommandRow>
      </OfficeHomeCommandStack>
    );
  }
  return (
    <span
      data-control-id={PRESENTATION_ANIMATION_TIMING_CONTROL_IDS[3]}
      style={{ display: "contents" }}
    >
      <select
        value={step.trigger}
        disabled={disabled}
        onChange={(event) =>
          onCommit(index, copyTimelineStep(step), {
            trigger: event.target.value as TimelineStep["trigger"],
            delayMs: step.delayMs,
            durationMs: step.durationMs,
          })
        }
        data-control-id={startControlId}
        data-testid={startTestId}
        aria-label={t("presentation.timelineTrigger")}
      >
        {TIMELINE_TRIGGERS.map((value) => (
          <option key={value} value={value}>
            {t(
              `presentation.timeline.trigger.${value}` as "presentation.timeline.trigger.on-click",
            )}
          </option>
        ))}
      </select>
      <AnimationTimingNumberInput
        field="durationMs"
        index={index}
        step={step}
        disabled={disabled}
        label={t("presentation.timelineDuration")}
        controlId={durationControlId}
        testId={durationTestId}
        onCommit={onCommit}
      />
      <AnimationTimingNumberInput
        field="delayMs"
        index={index}
        step={step}
        disabled={disabled}
        label={t("presentation.timelineDelay")}
        controlId={delayControlId}
        testId={delayTestId}
        onCommit={onCommit}
      />
    </span>
  );
}

function webAnimationOptionsForStep(
  step: TimelineStep,
): TimelineAnimation[] {
  const options = [...WEB_AUTHORABLE_TIMELINE_ANIMATIONS] as TimelineAnimation[];
  if (!options.includes(step.animation)) options.push(step.animation);
  return options;
}

function webMotionPathOptionsForStep(
  step: TimelineStep,
): MotionPathPreset[] {
  const options = [...SCRATCH_ANIMATION_MOTION_PATHS] as MotionPathPreset[];
  const current = step.motionPath?.preset;
  if (current && !options.includes(current)) options.push(current);
  return options;
}

export const ANIMATION_PANE_MIN_WIDTH = 280;

export const ANIMATION_PANE_DEFAULT_WIDTH = 340;

export const ANIMATION_PANE_MAX_WIDTH = 480;

export const ANIMATION_PANE_WIDTH_STEP = 16;

export function animationPaneMaxWidth(viewportWidth: number): number {
  return Math.max(
    ANIMATION_PANE_MIN_WIDTH,
    Math.min(ANIMATION_PANE_MAX_WIDTH, viewportWidth - 48),
  );
}

export function clampAnimationPaneWidth(value: number, viewportWidth: number): number {
  const maximum = animationPaneMaxWidth(viewportWidth);
  return Math.min(maximum, Math.max(ANIMATION_PANE_MIN_WIDTH, value));
}

/**
 * The slide-scoped ANIMATION TIMELINE editor (PowerPoint animation-pane parity).
 * Lists the slide's ordered steps; each row picks the animated block, its
 * entrance animation, its trigger (on-click / with-previous / after-previous),
 * and its delay + duration; rows can be reordered, removed, or appended for any
 * block on the slide. Purely a controlled view over the parent's
 * `animationSequence` state — every change routes back through the parent's
 * timeline callbacks (which persist + clamp). Self-contained + side-effect free.
 */
export function AnimationTimelinePanel({
  sequence,
  blocks,
  t,
  previewEnabled,
  previewActive,
  selectedIndex,
  rowEditingEnabled,
  timingEditingEnabled,
  sequenceEditingEnabled,
  onAddStep,
  onPreview,
  onSelectStep,
  onEditSequence,
  onUpdateStep,
  onUpdateTiming,
  blockFallbackName,
}: {
  sequence: readonly TimelineStep[];
  blocks: readonly PresentationBlock[];
  t: (key: string, values?: Record<string, string | number>) => string;
  /**
   * The block's accessible name (what the Selection Pane shows), used for a
   * block without words or a title so an empty shape reads "Rectangle", not
   * its internal id (PPT audit B442).
   */
  blockFallbackName?: (block: PresentationBlock) => string;
  previewEnabled: boolean;
  previewActive: boolean;
  selectedIndex: number | null;
  rowEditingEnabled: boolean;
  timingEditingEnabled: boolean;
  sequenceEditingEnabled: boolean;
  onAddStep: (blockId: string) => void;
  onPreview: () => void;
  onSelectStep: (index: number) => void;
  onEditSequence: (operation: AnimationSequenceEditOperation) => void;
  onUpdateStep: (index: number, patch: Partial<TimelineStep>) => void;
  onUpdateTiming: (
    index: number,
    expectedStep: TimelineStep,
    value: AnimationTimingValue,
  ) => void;
}) {
  // PERF: precompute id -> label once per blocks change. labelFor is called
  // once per add-menu option AND once per timeline step on every render; a raw
  // blocks.find per call is O(blocks × (blocks + steps)) on a busy slide.
  const blockLabelById = useMemo(() => {
    const map = new Map<string, string>();
    for (const block of blocks)
      map.set(
        block.id,
        presentationBlockLabel(block, blockFallbackName?.(block)),
      );
    return map;
  }, [blockFallbackName, blocks]);
  // Short, stable label for a block id (falls back to the raw id for a block
  // that no longer exists — pruning normally prevents this, defensive only).
  const labelFor = (blockId: string): string =>
    blockLabelById.get(blockId) ?? blockId;
  const atCap = sequence.length >= MAX_TIMELINE_STEPS;
  const canMoveEarlier =
    sequenceEditingEnabled && selectedIndex !== null && selectedIndex > 0;
  const canMoveLater =
    sequenceEditingEnabled &&
    selectedIndex !== null &&
    selectedIndex < sequence.length - 1;
  const canRemove = sequenceEditingEnabled && selectedIndex !== null;
  const paneActionStyle = (enabled: boolean): CSSProperties => ({
    ...blockOrderButtonStyle,
    cursor: enabled ? "pointer" : "default",
    opacity: enabled ? 1 : 0.45,
  });
  return (
    <div
      className="presentation-window__timeline"
      data-testid="presentation-animation-timeline"
    >
      <div className="presentation-window__timeline-header">
        <span className="presentation-window__timeline-title">
          {t("presentation.timelineLabel")}
        </span>
        <div
          role="toolbar"
          aria-label={t("presentation.timelineLabel")}
          style={{ display: "inline-flex", gap: 4 }}
        >
          <button
            type="button"
            onClick={onPreview}
            disabled={!previewEnabled}
            title={t("toolbar.preview")}
            aria-label={t("toolbar.preview")}
            aria-pressed={previewActive}
            data-testid="presentation-timeline-play-all"
            data-control-id="PPTCTL.PANE.ANIMATION.PLAY_ALL"
            style={paneActionStyle(previewEnabled)}
          >
            <Play size={13} />
          </button>
          <button
            type="button"
            onClick={() => onEditSequence("move-earlier")}
            disabled={!canMoveEarlier}
            title={t("presentation.timelineMoveUp")}
            aria-label={t("presentation.timelineMoveUp")}
            data-testid="presentation-timeline-up"
            data-control-id={PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS[1]}
            style={paneActionStyle(canMoveEarlier)}
          >
            <ChevronUp size={13} />
          </button>
          <button
            type="button"
            onClick={() => onEditSequence("move-later")}
            disabled={!canMoveLater}
            title={t("presentation.timelineMoveDown")}
            aria-label={t("presentation.timelineMoveDown")}
            data-testid="presentation-timeline-down"
            data-control-id={PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS[2]}
            style={paneActionStyle(canMoveLater)}
          >
            <ChevronDown size={13} />
          </button>
          <button
            type="button"
            onClick={() => onEditSequence("remove")}
            disabled={!canRemove}
            title={t("presentation.timelineRemove")}
            aria-label={t("presentation.timelineRemove")}
            data-testid="presentation-timeline-remove"
            data-control-id={PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS[3]}
            style={paneActionStyle(canRemove)}
          >
            <X size={13} />
          </button>
        </div>
        {/* Add a block to the timeline. Resets to the placeholder after each
            pick so the same select can append the block again. */}
        <label className="presentation-window__timeline-add">
          <span className="presentation-window__timeline-sr">
            {t("presentation.timelineAdd")}
          </span>
          <select
            value=""
            disabled={!rowEditingEnabled || atCap || blocks.length === 0}
            onChange={(event) => {
              const id = event.target.value;
              if (id) onAddStep(id);
            }}
            data-testid="presentation-timeline-add"
            aria-label={t("presentation.timelineAdd")}
          >
            <option value="">{t("presentation.timelineAdd")}</option>
            {blocks.map((block, blockIndex) => (
              <option key={`${block.id}:${blockIndex}`} value={block.id}>
                {labelFor(block.id)}
              </option>
            ))}
          </select>
        </label>
      </div>
      {sequence.length === 0 ? (
        <p
          className="presentation-window__timeline-empty"
          data-testid="presentation-timeline-empty"
        >
          {t("presentation.timelineEmpty")}
        </p>
      ) : (
        <ol
          className="presentation-window__timeline-list"
          data-testid="presentation-animation-timeline-list"
          data-control-id={PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS[4]}
        >
          {sequence.map((step, index) => (
            <li
              // Key by index, not step.id: the list is fully controlled (every
              // add/remove/reorder replaces the whole sequence), and persisted
              // ids carry no cross-step uniqueness guarantee — imported/hand-
              // edited decks can repeat an id, which would collide React keys.
              key={`${step.blockId}-${index}`}
              className="presentation-window__timeline-step"
              data-testid="presentation-timeline-step"
              data-block-id={step.blockId}
              data-selected={selectedIndex === index ? "true" : "false"}
            >
              <button
                type="button"
                className="presentation-window__timeline-block"
                title={labelFor(step.blockId)}
                aria-pressed={selectedIndex === index}
                onClick={() => onSelectStep(index)}
                data-testid="presentation-timeline-select"
                style={
                  selectedIndex === index
                    ? {
                        borderColor: "var(--console-accent, #0d9488)",
                        background:
                          "color-mix(in srgb, var(--console-accent, #0d9488) 12%, transparent)",
                      }
                    : undefined
                }
              >
                <span className="presentation-window__timeline-order">
                  {index + 1}
                </span>
                {labelFor(step.blockId)}
                {step.paragraphs ? (
                  <span
                    data-testid="presentation-timeline-paragraphs"
                    style={{
                      marginInlineStart: 6,
                      color: "var(--console-text-muted, #64748b)",
                      fontSize: "0.92em",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {step.paragraphs.start === step.paragraphs.end
                      ? t("presentation.animations.paragraphOne", {
                          index: step.paragraphs.start + 1,
                        })
                      : t("presentation.animations.paragraphRange", {
                          start: step.paragraphs.start + 1,
                          end: step.paragraphs.end + 1,
                        })}
                  </span>
                ) : null}
              </button>
              <select
                value={step.animation}
                disabled={!rowEditingEnabled || step.motionPath !== undefined}
                title={
                  step.motionPath !== undefined
                    ? t("presentation.animations.effectOptionsUnavailable")
                    : undefined
                }
                onChange={(event) =>
                  onUpdateStep(index, {
                    animation: event.target.value as TimelineAnimation,
                  })
                }
                data-testid="presentation-timeline-animation"
                aria-label={t("presentation.animationLabel")}
              >
                {webAnimationOptionsForStep(step).map((value) => (
                  <option
                    key={value}
                    value={value}
                    disabled={
                      !WEB_AUTHORABLE_TIMELINE_ANIMATIONS.includes(
                        value as (typeof WEB_AUTHORABLE_TIMELINE_ANIMATIONS)[number],
                      ) ||
                      (value === "none" && step.motionPath === undefined) ||
                      // The Ribbon's own object rule: a colour effect needs
                      // the object that colour lives on (PPT-B033).
                      (value !== "none" &&
                        value !== step.animation &&
                        !presentationAnimationEffectTargetEligible(
                          value as TimelineAnimation,
                          blocks.find((block) => block.id === step.blockId) ?? {
                            id: step.blockId,
                          },
                        ))
                    }
                  >
                    {t(
                      `presentation.animation.${value}` as "presentation.animation.none",
                    )}
                  </option>
                ))}
              </select>
              {






}
              <select
                value={step.motionPath?.preset ?? ""}
                disabled={!rowEditingEnabled || step.motionPath === undefined}
                title={t("presentation.animations.effectOptionsUnavailable")}
                onChange={(event) => {
                  const value = event.target.value;
                  // Empty → clear; a preset → its canonical path. `custom` is not a
                  // selectable option (no canonical geometry), so it never arrives
                  // here; presetMotionPath returns null for it and clears.
                  onUpdateStep(index, {
                    motionPath: value
                      ? (presetMotionPath(value) ?? undefined)
                      : undefined,
                  });
                }}
                data-testid="presentation-timeline-motion-path"
                aria-label={t("presentation.motionPathLabel")}
              >
                <option value="">{t("presentation.motionPath.none")}</option>
                {webMotionPathOptionsForStep(step).map((value) => (
                  <option key={value} value={value} disabled>
                    {t(
                      `presentation.motionPath.${value}` as "presentation.motionPath.line-right",
                    )}
                  </option>
                ))}
              </select>
              <AnimationTimingControls
                scope="pane"
                index={index}
                step={step}
                disabled={!timingEditingEnabled || selectedIndex !== index}
                t={t}
                onCommit={onUpdateTiming}
              />
            </li>
          ))}
        </ol>
      )}
    </div>
  );
}
