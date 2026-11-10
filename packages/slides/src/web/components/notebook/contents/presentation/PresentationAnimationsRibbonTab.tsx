"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Animations tab of the presentation Ribbon (Preview, Animation, Advanced Animation, Timing), drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeMenuTrigger,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  IconRibbonPreview,
  IconAnimationEffectOptions,
  IconAnimationPane,
  IconAnimationPainter,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { PRESENTATION_ANIMATION_EFFECT_CONTROL_IDS, PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS } from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import {
  PresentationAnimationRibbonStrip,
  PRESENTATION_ANIMATION_RIBBON_ENTRANCE,
  PresentationAnimationMoreChevron,
  PRESENTATION_ANIMATION_RIBBON_EMPHASIS,
  PresentationAnimationRibbonMenu,
  PRESENTATION_ANIMATION_RIBBON_EXIT,
  presentationAnimationRibbonMotionPaths,
  presentationAnimationRibbonEmphasisSections,
  PresentationAnimationSectionsMenuContent,
} from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationAnimationRibbonGalleries.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationAnimationIcon, PresentationAnimationGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationAnimationGallery.tsx";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuItem,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  PRESENTATION_WEB_ANIMATION_EFFECTS,
  type PresentationAnimationEffectOption,
  presentationAnimationFontOfOption,
  copyTimelineStep,
  DEFAULT_STEP_DURATION_MS,
  type AnimationSequenceEditOperation,
  type AnimationTimingValue,
  type TimelineAnimation,
  type TimelinePreviewFrame,
  type TimelineStep,
  type TimelineTextSequence,
} from "@mrnote/core/web/lib/presentation-anim-timeline.ts";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { Pointer, ChevronUp, ChevronDown } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import documentHomeStyles from "@mrnote/core/web/components/console/editor/DocumentHomeRibbon.module.css";
import { AnimationTimingControls } from "./PresentationAnimationPane";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { OfficeColorPalettePanelLabels } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePanel.tsx";
import type { PresentationAnimationEffectCommandAdapter } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { MotionPathPreset } from "@mrnote/core/web/lib/presentation-motion-path.ts";

/** The emphasis strip's full list (its chevron), fixed by the catalogue. */
const EMPHASIS_SECTIONS = presentationAnimationRibbonEmphasisSections();

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationAnimationsRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  animationPaneOpen: boolean;
  animationPaneToggleRef: RefObject<HTMLButtonElement | null>;
  selectedAnimationIndex: number | null;
  animationPreview: { readonly slideModelId: string; readonly token: number; readonly frames: readonly TimelinePreviewFrame[]; readonly frameIndex: number; } | null;
  selectedBlockId: string | null;
  animationEffectOptionsMenuOpen: boolean;
  setAnimationEffectOptionsMenuOpen: Dispatch<SetStateAction<boolean>>;
  animationEffectRecentColors: string[];
  activeSlide: SlideModel;
  exclusivelySelectedBlockId: string | null;
  tableCellColorPaletteLabels: OfficeColorPalettePanelLabels;
  animationEffectCommand: PresentationAnimationEffectCommandAdapter;
  blockAriaName: (block: PresentationBlock) => string;
  armAnimationPainter: (sourceBlockId: string) => void;
  animationPainterArmed: boolean;
  timelineSequence: TimelineStep[];
  timelinePreviewFrames: TimelinePreviewFrame[];
  timelineSequenceEditingEnabled: boolean;
  timelineTimingEditingEnabled: boolean;
  selectedAnimationTimingStep: TimelineStep | null;
  selectedAnimationEffectOption: PresentationAnimationEffectOption | null;
  selectedAnimationEffectColor: string | null;
  selectedAnimationEffectOptions: readonly PresentationAnimationEffectOption[];
  animationEffectOptionsEditable: boolean;
  animationTextSequenceOptions: readonly TimelineTextSequence[];
  selectedAnimationTextSequence: TimelineTextSequence | null;
  setAnimationTextSequence: (sequence: TimelineTextSequence) => void;
  animationObjectTriggerTargets: PresentationBlock[];
  animationObjectTriggerEditable: boolean;
  canAuthorTimelineMenuEffect: (animation: TimelineAnimation, motionPathPreset?: MotionPathPreset) => boolean;
  addTimelineAnimationEffect: (kind: TimelineAnimation, motionPathPreset?: MotionPathPreset) => void;
  applyAnimationGalleryEffect: (kind: TimelineAnimation) => void;
  toggleAnimationPane: () => void;
  editAnimationSequence: (operation: AnimationSequenceEditOperation) => void;
  editAnimationTiming: (index: number, expectedStep: TimelineStep, value: AnimationTimingValue) => void;
  previewAnimations: () => void;
  setAnimationEffectOption: (option: PresentationAnimationEffectOption) => void;
  setAnimationObjectTrigger: (triggerBlockId: string | null) => void;
  setAnimationEffectColor: (value: string | null) => void;
}

export function PresentationAnimationsRibbonTab({ ctl }: { ctl: PresentationAnimationsRibbonTabCtl }) {
  const {
    t,
    animationPaneOpen,
    animationPaneToggleRef,
    selectedAnimationIndex,
    animationPreview,
    selectedBlockId,
    animationEffectOptionsMenuOpen,
    setAnimationEffectOptionsMenuOpen,
    animationEffectRecentColors,
    activeSlide,
    exclusivelySelectedBlockId,
    tableCellColorPaletteLabels,
    animationEffectCommand,
    blockAriaName,
    armAnimationPainter,
    animationPainterArmed,
    timelineSequence,
    timelinePreviewFrames,
    timelineSequenceEditingEnabled,
    timelineTimingEditingEnabled,
    selectedAnimationTimingStep,
    selectedAnimationEffectOption,
    selectedAnimationEffectColor,
    selectedAnimationEffectOptions,
    animationEffectOptionsEditable,
    animationTextSequenceOptions,
    selectedAnimationTextSequence,
    setAnimationTextSequence,
    animationObjectTriggerTargets,
    animationObjectTriggerEditable,
    canAuthorTimelineMenuEffect,
    addTimelineAnimationEffect,
    applyAnimationGalleryEffect,
    toggleAnimationPane,
    editAnimationSequence,
    editAnimationTiming,
    previewAnimations,
    setAnimationEffectOption,
    setAnimationObjectTrigger,
    setAnimationEffectColor,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.animations.previewGroup")}
                      ribbon="animations"
                      size="narrow"
                      testId="presentation-ribbon-animation-preview-group"
                    >
                      <OfficeHomeCommand
                        active={
                          animationPreview?.slideModelId === activeSlide.id
                        }
                        controlId="PPTCTL.RIBBON.ANIMATIONS.PREVIEW.PRIMARY"
                        disabled={timelinePreviewFrames.length === 0}
                        disabledReason={
                          timelinePreviewFrames.length === 0
                            ? t("document.capability.unavailable")
                            : undefined
                        }
                        icon={<IconRibbonPreview home size={24} />}
                        label={t("toolbar.preview")}
                        onClick={previewAnimations}
                        testId="presentation-ribbon-animation-preview"
                        variant="large"
                      />
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      controlId={PRESENTATION_ANIMATION_EFFECT_CONTROL_IDS[0]}
                      label={t("presentation.animations.effectsGroup")}
                      ribbon="animations"
                      size="gallery"
                      testId="presentation-animation-strip"
                    >
                      <OfficeHomeCommandRow>
                        <PresentationAnimationRibbonStrip
                          ariaLabel={t("presentation.animations.entranceEffects")}
                          className={presentationChrome.animationGallery}
                          testId="presentation-animation-gallery"
                          tiles={PRESENTATION_ANIMATION_RIBBON_ENTRANCE}
                          canSelect={canAuthorTimelineMenuEffect}
                          onSelect={applyAnimationGalleryEffect}
                          selected={selectedAnimationTimingStep?.animation}
                          t={(key) => t(key as "presentation.animation.fade")}
                          leading={
                            <OfficeHomeCommand
                              variant="large"
                              controlId="PPTCTL.RIBBON.ANIMATIONS.NONE"
                              icon={<PresentationAnimationIcon effect="none" size={24} />}
                              label={t("presentation.animation.none")}
                              onClick={() => editAnimationSequence("remove")}
                              // The remembered step must be the selected object's:
                              // picking another object and pressing None used to
                              // delete the previous object's effect (PPT-B380). The
                              // pane's own buttons keep editing the step they name.
                              disabled={
                                selectedAnimationTimingStep === null ||
                                selectedAnimationTimingStep.blockId !== selectedBlockId ||
                                !timelineSequenceEditingEnabled
                              }
                              testId="presentation-animation-quick-none"
                            />
                          }
                          more={
                            <DropdownMenu>
                              {/* The full catalogue (Add Animation): Entrance,
                                  Emphasis, Exit and Motion Paths in PowerPoint's
                                  Basic / Subtle / Moderate / Exciting groups. */}
                              <PresentationAnimationMoreChevron
                                label={t("presentation.animations.addAnimation")}
                                disabled={
                                  !PRESENTATION_WEB_ANIMATION_EFFECTS.some(({ id }) =>
                                    canAuthorTimelineMenuEffect(id),
                                  )
                                }
                                testId="presentation-animation-more-trigger"
                              />
                              <PresentationAnimationGallery
                                selected={selectedAnimationTimingStep?.animation}
                                canSelect={canAuthorTimelineMenuEffect}
                                onSelect={addTimelineAnimationEffect}
                                canRemove={selectedAnimationTimingStep !== null && timelineSequenceEditingEnabled}
                                onRemove={() => editAnimationSequence("remove")}
                                t={key => t(key as "presentation.animation.fade")}
                              />
                            </DropdownMenu>
                          }
                        />
                        <PresentationAnimationRibbonStrip
                          ariaLabel={t("presentation.animations.emphasisEffects")}
                          className={presentationChrome.animationGallery}
                          testId="presentation-animation-emphasis-gallery"
                          tiles={PRESENTATION_ANIMATION_RIBBON_EMPHASIS}
                          canSelect={canAuthorTimelineMenuEffect}
                          onSelect={applyAnimationGalleryEffect}
                          selected={selectedAnimationTimingStep?.animation}
                          t={(key) => t(key as "presentation.animation.fade")}
                          more={
                            <DropdownMenu>
                              {/* Every emphasis effect, beyond the tiles the
                                  ribbon has room for, in PowerPoint's More
                                  Emphasis Effects groups. */}
                              <PresentationAnimationMoreChevron
                                label={t("presentation.animations.emphasisEffects")}
                                disabled={
                                  !EMPHASIS_SECTIONS.some((section) =>
                                    section.items.some((tile) => canAuthorTimelineMenuEffect(tile.id)),
                                  )
                                }
                                testId="presentation-animation-emphasis-more-trigger"
                              />
                              <PresentationAnimationSectionsMenuContent
                                label={t("presentation.animations.emphasisEffects")}
                                testId="presentation-animation-emphasis-more"
                                sections={EMPHASIS_SECTIONS}
                                sectionLabel={(group) =>
                                  t(`presentation.animations.group.${group}` as "presentation.animations.group.basic")
                                }
                                canSelect={canAuthorTimelineMenuEffect}
                                onSelect={applyAnimationGalleryEffect}
                                selected={selectedAnimationTimingStep?.animation}
                                t={(key) => t(key as "presentation.animation.fade")}
                              />
                            </DropdownMenu>
                          }
                        />
                        <PresentationAnimationRibbonMenu
                          label={t("presentation.animations.exitEffects")}
                          icon={<PresentationAnimationIcon effect="fly-out" size={24} />}
                          testId="presentation-ribbon-animation-exit-effects"
                          controlId="PPTCTL.RIBBON.ANIMATIONS.EXIT_EFFECTS"
                          sections={PRESENTATION_ANIMATION_RIBBON_EXIT}
                          sectionLabel={(group) =>
                            t(`presentation.animations.group.${group}` as "presentation.animations.group.basic")
                          }
                          canSelect={canAuthorTimelineMenuEffect}
                          onSelect={applyAnimationGalleryEffect}
                          selected={selectedAnimationTimingStep?.animation}
                          t={(key) => t(key as "presentation.animation.fade")}
                        />
                        <PresentationAnimationRibbonMenu
                          label={t("presentation.animations.pathAnimation")}
                          icon={<PresentationAnimationIcon effect="path-arc-down" size={24} />}
                          testId="presentation-ribbon-animation-path-effects"
                          controlId="PPTCTL.RIBBON.ANIMATIONS.MOTION_PATHS"
                          sections={presentationAnimationRibbonMotionPaths()}
                          sectionLabel={(group) =>
                            t(`presentation.animations.pathGroup.${group}` as "presentation.animations.pathGroup.basic")
                          }
                          canSelect={canAuthorTimelineMenuEffect}
                          onSelect={applyAnimationGalleryEffect}
                          selected={selectedAnimationTimingStep?.animation}
                          t={(key) => t(key as "presentation.animation.fade")}
                        />
                        <DropdownMenu
                          open={animationEffectOptionsMenuOpen}
                          onOpenChange={setAnimationEffectOptionsMenuOpen}
                        >
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="large"
                              data-control-id="PPTCTL.RIBBON.ANIMATIONS.EFFECT_OPTIONS"
                              data-disabled-reason={
                                animationEffectOptionsEditable
                                  ? undefined
                                  : t(
                                      "presentation.animations.effectOptionsUnavailable",
                                    )
                              }
                              displayLabel={t(
                                "presentation.animations.effectOptions",
                              )}
                              hasMenuIndicator
                              icon={
                                <IconAnimationEffectOptions home size={24} />
                              }
                              label={t(
                                "presentation.animations.effectOptions",
                              )}
                              title={
                                animationEffectOptionsEditable
                                  ? undefined
                                  : t(
                                      "presentation.animations.effectOptionsUnavailable",
                                    )
                              }
                              disabled={!animationEffectOptionsEditable}
                              testId="presentation-ribbon-animation-effect-options"
                            />
                          </DropdownMenuTrigger>
                          {animationEffectOptionsEditable &&
                          (selectedAnimationEffectOption !== null ||
                            selectedAnimationEffectColor !== null ||
                            animationTextSequenceOptions.length > 1) ? (
                            <RibbonDropdownMenuContent
                              align="start"
                              data-testid="presentation-ribbon-animation-effect-options-menu"
                            >
                              <DropdownMenuLabel>
                                {t("presentation.animations.effectOptions")}
                              </DropdownMenuLabel>
                              {selectedAnimationEffectColor !== null ? (
                                <OfficeColorPalettePopover
                                  value={selectedAnimationEffectColor}
                                  recents={animationEffectRecentColors}
                                  labels={tableCellColorPaletteLabels}
                                  onPick={setAnimationEffectColor}
                                  onDismiss={() =>
                                    setAnimationEffectOptionsMenuOpen(false)
                                  }
                                  testIdPrefix="presentation-animation-effect-color"
                                  controlId="PPTCTL.RIBBON.ANIMATIONS.EFFECT_OPTIONS"
                                  controlSurface="presentation.animations.effect-options"
                                />
                              ) : selectedAnimationEffectOption === null ? null : (
                                <DropdownMenuRadioGroup
                                  value={selectedAnimationEffectOption ?? undefined}
                                  onValueChange={(value) =>
                                    setAnimationEffectOption(
                                      value as PresentationAnimationEffectOption,
                                    )
                                  }
                                >
                                  {selectedAnimationEffectOptions.map((option) => (
                                    <DropdownMenuRadioItem
                                      key={option}
                                      value={option}
                                      data-authoring-disposition="native"
                                      data-testid={`presentation-ribbon-animation-effect-option-${option}`}
                                    >
                                      {presentationAnimationFontOfOption(option) ??
                                        t(
                                          `presentation.animation.effectOption.${option}` as "presentation.animation.effectOption.horizontal-in",
                                        )}
                                    </DropdownMenuRadioItem>
                                  ))}
                                </DropdownMenuRadioGroup>
                              )}
                              {animationTextSequenceOptions.length > 1 ? (
                                <>
                                  <DropdownMenuSeparator />
                                  <DropdownMenuLabel>
                                    {t("presentation.animations.sequenceLabel")}
                                  </DropdownMenuLabel>
                                  <DropdownMenuRadioGroup
                                    value={selectedAnimationTextSequence ?? undefined}
                                    onValueChange={(value) =>
                                      setAnimationTextSequence(
                                        value as TimelineTextSequence,
                                      )
                                    }
                                  >
                                    {animationTextSequenceOptions.map((sequence) => (
                                      <DropdownMenuRadioItem
                                        key={sequence}
                                        value={sequence}
                                        data-testid={`presentation-ribbon-animation-sequence-${sequence}`}
                                      >
                                        {t(
                                          `presentation.animations.sequence.${sequence}` as "presentation.animations.sequence.by-paragraph",
                                        )}
                                      </DropdownMenuRadioItem>
                                    ))}
                                  </DropdownMenuRadioGroup>
                                </>
                              ) : null}
                            </RibbonDropdownMenuContent>
                          ) : null}
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.animations.advancedGroup")}
                      ribbon="animations"
                      size="standard"
                      testId="presentation-ribbon-animation-pane-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.RIBBON.ANIMATIONS.PANE"
                          buttonRef={animationPaneToggleRef}
                          icon={<IconAnimationPane home size={24} />}
                          label={t("presentation.animations.animationPane")}
                          onClick={toggleAnimationPane}
                          active={animationPaneOpen}
                          testId="presentation-ribbon-animation-pane"
                        />
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="large"
                              hasMenuIndicator
                              data-control-id="PPTCTL.RIBBON.ANIMATIONS.TRIGGER"
                              disabled={!animationObjectTriggerEditable}
                              disabledReason={
                                animationObjectTriggerEditable
                                  ? undefined
                                  : selectedAnimationTimingStep === null
                                    ? t("presentation.animations.triggerSelectEffect")
                                    : t("presentation.animations.triggerRequiresTarget")
                              }
                              displayLabel={t("presentation.animations.trigger")}
                              icon={<Pointer size={20} />}
                              label={t("presentation.animations.trigger")}
                              testId="presentation-ribbon-animation-trigger"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-animation-trigger-menu"
                          >
                                <DropdownMenuLabel>
                                  {t(
                                    "presentation.animations.triggerOnClick",
                                  )}
                                </DropdownMenuLabel>
                                <DropdownMenuRadioGroup
                                  value={
                                    selectedAnimationTimingStep?.triggerBlockId ??
                                    "none"
                                  }
                                  onValueChange={(value) =>
                                    setAnimationObjectTrigger(
                                      value === "none" ? null : value,
                                    )
                                  }
                                >
                                  <DropdownMenuRadioItem
                                    value="none"
                                    data-testid="presentation-ribbon-animation-trigger-none"
                                  >
                                    {t(
                                      "presentation.animations.triggerNone",
                                    )}
                                  </DropdownMenuRadioItem>
                                  {animationObjectTriggerTargets.map((block) => (
                                    <DropdownMenuRadioItem
                                      key={block.id}
                                      value={block.id}
                                      disabled={
                                        selectedAnimationIndex === null ||
                                        selectedAnimationTimingStep === null ||
                                        !animationEffectCommand.canExecute({
                                          operation: "set-object-trigger",
                                          slideModelId: activeSlide.id,
                                          index: selectedAnimationIndex,
                                          expectedStep: copyTimelineStep(
                                            selectedAnimationTimingStep,
                                          ),
                                          triggerBlockId: block.id,
                                        })
                                      }
                                      data-testid={`presentation-ribbon-animation-trigger-object-${block.id}`}
                                    >
                                      {blockAriaName(block)}
                                    </DropdownMenuRadioItem>
                                  ))}
                                </DropdownMenuRadioGroup>
                                <DropdownMenuSeparator />
                                <DropdownMenuItem
                                  disabled
                                  title={t(
                                    "presentation.animations.triggerBookmarkUnavailable",
                                  )}
                                  data-testid="presentation-ribbon-animation-trigger-bookmark"
                                >
                                  {t(
                                    "presentation.animations.triggerBookmark",
                                  )}
                                </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <OfficeHomeCommand
                          variant="large"
                          controlId={PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS[0]}
                          icon={<IconAnimationPainter home size={24} />}
                          label={t("presentation.animationPainter.copy")}
                          active={animationPainterArmed}
                          disabled={
                                !exclusivelySelectedBlockId ||
                                !timelineSequence.some(
                                  (step) =>
                                    step.blockId === exclusivelySelectedBlockId,
                                )
                              }
                          onClick={() => {
                            if (exclusivelySelectedBlockId) {
                              armAnimationPainter(exclusivelySelectedBlockId);
                            }
                          }}
                          testId="presentation-ribbon-animation-painter"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.animations.timingGroup")}
                      ribbon="animations"
                      size="content"
                      testId="presentation-ribbon-animation-timing"
                      className={`${documentHomeStyles.referencesNotesGroup} ${presentationChrome.effectsTimingGroup}`}
                    >
                      {/* The band's group stacks its direct children, so the
                          timing fields and Reorder Animation sat one above
                          the other and Reorder fell below the caption: one
                          row holds both, as PowerPoint lays them out. */}
                      <div className="note-editor-ribbon-hstack">
                      <AnimationTimingControls
                        scope="ribbon"
                        index={selectedAnimationIndex ?? -1}
                        step={
                          selectedAnimationTimingStep ?? {
                            blockId: "",
                            animation: "appear",
                            trigger: "on-click",
                            delayMs: 0,
                            durationMs: DEFAULT_STEP_DURATION_MS,
                          }
                        }
                        disabled={
                          selectedAnimationTimingStep === null ||
                          !timelineTimingEditingEnabled
                        }
                        empty={selectedAnimationTimingStep === null}
                        t={t}
                        onCommit={editAnimationTiming}
                      />
                      {/* PowerPoint keeps Reorder Animation in the Timing group. */}
                      <div
                        role="group"
                        aria-label={t("presentation.animations.reorderGroup")}
                        data-testid="presentation-ribbon-animation-reorder-group"
                      >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          icon={<ChevronUp size={20} />}
                          label={t("presentation.animations.moveEarlier")}
                          onClick={() => editAnimationSequence("move-earlier")}
                          disabled={
                            !timelineSequenceEditingEnabled ||
                            selectedAnimationIndex === null ||
                            selectedAnimationIndex <= 0
                          }
                          testId="presentation-ribbon-animation-move-earlier"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          icon={<ChevronDown size={20} />}
                          label={t("presentation.animations.moveLater")}
                          onClick={() => editAnimationSequence("move-later")}
                          disabled={
                            !timelineSequenceEditingEnabled ||
                            selectedAnimationIndex === null ||
                            selectedAnimationIndex >= timelineSequence.length - 1
                          }
                          testId="presentation-ribbon-animation-move-later"
                        />
                      </OfficeHomeCommandRow>
                      </div>
                      </div>
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
