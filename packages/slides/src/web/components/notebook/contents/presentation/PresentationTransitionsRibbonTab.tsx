"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Transitions tab of the presentation Ribbon (Preview, Transition to This Slide, Timing), drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { CaseSensitive, Shapes, WholeWord } from "lucide-react";
import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeMenuTrigger,
  OfficeHomeNumberField,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import {
  IconRibbonPreview,
  IconTransitionEffectOptions,
  IconTransitionClockOption,
  IconTransitionP14TypedOption,
  IconTransitionShapeOption,
  IconTransitionZoomOption,
  IconTransitionSplitOption,
  IconTransitionOrientationOption,
  IconTransitionDirection,
  IconTimingDuration,
  IconTimingSound,
  IconTransitionApplyAll,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { PresentationTransitionGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationTransitionGallery.tsx";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import type { PresentationBuiltinTransitionSoundId } from "@mrnote/core/web/lib/presentation-transition-builtin-sounds.ts";
import {
  type SlideTransitionClockOption,
  isSlideTransitionP14TypedOptionsKind,
  type SlideTransitionP14TypedOption,
  isSlideTransitionShapeOption,
  isSlideTransitionZoomOption,
  isSlideTransitionSplitOption,
  isSlideTransitionOrientationOption,
  isSlideTransitionSideDirection,
  isSlideTransitionCornerOption,
  MIN_TRANSITION_DURATION_SECONDS,
  MAX_TRANSITION_DURATION_SECONDS,
  TRANSITION_DURATION_STEP_SECONDS,
  type AuthorableSlideTransition,
  type MorphSlideTransitionPreviewPlayback,
  type PresentationTransitionSound,
  type SlideTransitionDirection,
  type SlideTransitionPreviewPlayback,
} from "@mrnote/core/web/lib/presentation-transitions.ts";
import documentHomeStyles from "@mrnote/core/web/components/console/editor/DocumentHomeRibbon.module.css";
import { PresentationTransitionTimingControls } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationTransitionTiming.tsx";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { EditorTransitionPreviewPlayback } from "./PresentationPresentOverlay";
import type { KeyboardEvent, RefObject } from "react";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import type { PresentationSlideTransitionCommandAdapter } from "@mrnote/core/web/lib/presentation-command-adapters.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationTransitionsRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  transitionPreview: EditorTransitionPreviewPlayback | null;
  transitionDurationSkipBlurRef: RefObject<boolean>;
  activeSlide: SlideModel;
  normalizedActiveTransition: "none" | "fade" | "slide-left" | "slide-up" | "slide-right" | "slide-down" | "zoom" | "dissolve" | "push-left" | "push-right" | "push-up" | "push-down" | "wipe-right" | "wipe-left" | "wipe-down" | "wipe-up" | "morph" | "push" | "wipe" | "cut" | "shape" | "split" | "random-bars" | "checkerboard" | "blinds" | "comb" | "fall-over" | "drape" | "curtains" | "wind" | "prestige" | "fracture" | "crush" | "peel-off" | "page-curl" | "airplane" | "origami" | "switch" | "flip" | "gallery" | "ferris-wheel" | "conveyor" | "uncover" | "cover" | "doors" | "window" | "vortex" | "pan" | "flash" | "honeycomb" | "reveal" | "ripple" | "glitter" | "shred" | "fly-through" | "clock" | "random" | "cube" | "box" | "rotate" | "orbit";
  activeTransitionPreview: SlideTransitionPreviewPlayback | MorphSlideTransitionPreviewPlayback | null;
  transitionPreviewDisabledReason: string | undefined;
  slideTransitionCommand: PresentationSlideTransitionCommandAdapter;
  slideTransitionAdvance: { editable: boolean; setAdvance: (advanceOnClick: boolean, advanceAfterMs: number | null) => void; };
  setSlideTransition: (kind: AuthorableSlideTransition, direction?: SlideTransitionDirection) => void;
  setSlideTransitionDirection: (direction: string) => void;
  beginTransitionSoundPicker: (trigger: HTMLButtonElement | null) => void;
  clearTransitionSound: () => void;
  applyBuiltinTransitionSound: (id: PresentationBuiltinTransitionSoundId) => void;
  applyTransitionToAll: () => void;
  previewSlideTransition: () => void;
  beginTransitionDurationEdit: (slideModelId: string, durationMs: number) => void;
  changeTransitionDurationDraft: (slideModelId: string, value: string) => void;
  commitTransitionDurationDraft: () => void;
  handleTransitionDurationKeyDown: (event: KeyboardEvent<HTMLInputElement>) => void;
  activeTransitionDirection: SlideTransitionDirection | null;
  activeDirectionalTransition: "zoom" | "push" | "wipe" | "shape" | "split" | "random-bars" | "checkerboard" | "blinds" | "comb" | "fall-over" | "drape" | "wind" | "peel-off" | "page-curl" | "airplane" | "origami" | "switch" | "flip" | "gallery" | "ferris-wheel" | "conveyor" | "uncover" | "cover" | "doors" | "window" | "vortex" | "pan" | "reveal" | "ripple" | "glitter" | "shred" | "fly-through" | "clock" | "cube" | "box" | "rotate" | "orbit" | "morph" | null;
  transitionEffectDirectionOptions: { direction: SlideTransitionDirection; enabled: boolean; }[];
  transitionEffectOptionLabel: (option: SlideTransitionDirection) => string;
  transitionEffectOptionsEditable: boolean;
  transitionEffectOptionsDisabledReason: string | undefined;
  activeTransitionSound: PresentationTransitionSound | null;
  transitionSoundEditable: boolean;
  activeTransitionDurationMs: number;
  activeTransitionDurationInputValue: string;
  transitionDurationControlDisabled: boolean;
  stepTransitionDuration: (direction: "increment" | "decrement") => void;
  transitionApplyAllEditable: boolean;
}

export function PresentationTransitionsRibbonTab({ ctl }: { ctl: PresentationTransitionsRibbonTabCtl }) {
  const {
    t,
    transitionPreview,
    transitionDurationSkipBlurRef,
    activeSlide,
    normalizedActiveTransition,
    activeTransitionPreview,
    transitionPreviewDisabledReason,
    slideTransitionCommand,
    slideTransitionAdvance,
    setSlideTransition,
    setSlideTransitionDirection,
    beginTransitionSoundPicker,
    clearTransitionSound,
    applyBuiltinTransitionSound,
    applyTransitionToAll,
    previewSlideTransition,
    beginTransitionDurationEdit,
    changeTransitionDurationDraft,
    commitTransitionDurationDraft,
    handleTransitionDurationKeyDown,
    activeTransitionDirection,
    activeDirectionalTransition,
    transitionEffectDirectionOptions,
    transitionEffectOptionLabel,
    transitionEffectOptionsEditable,
    transitionEffectOptionsDisabledReason,
    activeTransitionSound,
    transitionSoundEditable,
    activeTransitionDurationMs,
    activeTransitionDurationInputValue,
    transitionDurationControlDisabled,
    stepTransitionDuration,
    transitionApplyAllEditable,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      label={t("presentation.transition.previewGroup")}
                      ribbon="transitions"
                      testId="presentation-ribbon-transition-preview-group"
                    >
                      <OfficeHomeCommand
                        active={
                          transitionPreview?.slideModelId === activeSlide.id
                        }
                        controlId="PPTCTL.RIBBON.TRANSITIONS.PREVIEW.PRIMARY"
                        disabled={activeTransitionPreview === null}
                        disabledReason={transitionPreviewDisabledReason}
                        icon={<IconRibbonPreview home size={24} />}
                        label={t("toolbar.preview")}
                        onClick={previewSlideTransition}
                        testId="presentation-ribbon-transition-preview"
                        variant="large"
                      />
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("presentation.transition.galleryGroup")}
                      ribbon="transitions"
                      testId="presentation-ribbon-transition-gallery-group"
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.secondaryGalleryRow}
                      >
                        <div
                          data-control-id="PPTCTL.RIBBON.TRANSITIONS.GALLERY"
                          style={{ display: "contents" }}
                        >
                          <PresentationTransitionGallery
                            groupLabels={{ subtle: t("presentation.transition.group.subtle"), exciting: t("presentation.transition.group.exciting"), dynamic: t("presentation.transition.group.dynamic") }}
                            value={normalizedActiveTransition}
                          labels={{
                            cube: t("presentation.transition.cube"), box: t("presentation.transition.box"), rotate: t("presentation.transition.rotate"), orbit: t("presentation.transition.orbit"),
                            none: t("presentation.transition.none"),
                            morph: t("presentation.transition.morph"),
                            fade: t("presentation.transition.fade"),
                            push: t("presentation.transition.push"),
                            wipe: t("presentation.transition.wipe"),
                            dissolve: t("presentation.transition.dissolve"),
                            cut: t("presentation.transition.cut"),
                            shape: t("presentation.transition.shape"),
                            split: t("presentation.transition.split"),
                            zoom: t("presentation.transition.zoom"),
                            "random-bars": t(
                              "presentation.transition.randomBars",
                            ),
                            checkerboard: t(
                              "presentation.transition.checkerboard",
                            ),
                            blinds: t("presentation.transition.blinds"),
                            clock: t("presentation.transition.clock"),
                            comb: t("presentation.transition.comb"),
                            "fall-over": t(
                              "presentation.transition.fallOver",
                            ),
                            drape: t("presentation.transition.drape"),
                            curtains: t("presentation.transition.curtains"),
                            wind: t("presentation.transition.wind"),
                            prestige: t("presentation.transition.prestige"),
                            fracture: t("presentation.transition.fracture"),
                            crush: t("presentation.transition.crush"),
                            "peel-off": t(
                              "presentation.transition.peelOff",
                            ),
                            "page-curl": t(
                              "presentation.transition.pageCurl",
                            ),
                            airplane: t("presentation.transition.airplane"),
                            origami: t("presentation.transition.origami"),
                            switch: t("presentation.transition.switch"),
                            flip: t("presentation.transition.flip"),
                            gallery: t("presentation.transition.gallery"),
                            "ferris-wheel": t(
                              "presentation.transition.ferrisWheel",
                            ),
                            conveyor: t("presentation.transition.conveyor"),
                            uncover: t("presentation.transition.uncover"),
                            cover: t("presentation.transition.cover"),
                            doors: t("presentation.transition.doors"),
                            window: t("presentation.transition.window"),
                            vortex: t("presentation.transition.vortex"),
                            pan: t("presentation.transition.pan"),
                            flash: t("presentation.transition.flash"),
                            honeycomb: t(
                              "presentation.transition.honeycomb",
                            ),
                            reveal: t("presentation.transition.reveal"),
                            ripple: t("presentation.transition.ripple"),
                            glitter: t("presentation.transition.glitter"),
                            shred: t("presentation.transition.shred"),
                            "fly-through": t(
                              "presentation.transition.flyThrough",
                            ),
                            random: t("presentation.transition.random"),
                          }}
                            unavailableLabel={t(
                              "document.capability.unavailable",
                            )}
                            isItemDisabled={(kind) =>
                              !slideTransitionCommand.canExecute(kind)
                            }
                            onSelect={setSlideTransition}
                            ariaLabel={t(
                              "presentation.transition.galleryGroup",
                            )}
                          />
                        </div>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              data-control-id="PPTCTL.RIBBON.TRANSITIONS.EFFECT_OPTIONS"
                              data-disabled-reason={
                                transitionEffectOptionsDisabledReason
                              }
                              disabled={!transitionEffectOptionsEditable}
                              displayLabel={t(
                                "presentation.transition.effectOptions",
                              )}
                              hasMenuIndicator
                              icon={
                                <IconTransitionEffectOptions home size={24} />
                              }
                              label={`${t("presentation.transition.effectOptions")}${transitionEffectOptionsDisabledReason ? `: ${transitionEffectOptionsDisabledReason}` : ""}`}
                              testId="presentation-ribbon-transition-effect-options"
                              variant="large"
                            />
                          </DropdownMenuTrigger>
                          {transitionEffectOptionsEditable &&
                          activeDirectionalTransition !== null ? (
                            <RibbonDropdownMenuContent
                              align="start"
                              data-testid="presentation-ribbon-transition-effect-options-menu"
                            >
                              <DropdownMenuLabel>
                                {t("presentation.transition.effectOptions")}
                              </DropdownMenuLabel>
                              <DropdownMenuRadioGroup
                                value={activeTransitionDirection ?? ""}
                                onValueChange={setSlideTransitionDirection}
                              >
                                {transitionEffectDirectionOptions.map(
                                  ({ direction, enabled }) => (
                                    <DropdownMenuRadioItem
                                      key={direction}
                                      value={direction}
                                      disabled={!enabled}
                                      data-authoring-disposition={
                                        enabled ? "native" : "blocked"
                                      }
                                      data-testid={`presentation-ribbon-transition-direction-${direction}`}
                                    >
                                      {activeDirectionalTransition ===
                                      "clock" ? (
                                        <IconTransitionClockOption
                                          option={
                                            direction as SlideTransitionClockOption
                                          }
                                          size={16}
                                        />
                                      ) : isSlideTransitionP14TypedOptionsKind(
                                        activeDirectionalTransition,
                                      ) ? (
                                        <IconTransitionP14TypedOption
                                          kind={activeDirectionalTransition}
                                          option={
                                            direction as SlideTransitionP14TypedOption
                                          }
                                          size={16}
                                        />
                                      ) : isSlideTransitionShapeOption(
                                          direction,
                                        ) ? (
                                        <IconTransitionShapeOption
                                          option={direction}
                                          size={16}
                                        />
                                      ) : isSlideTransitionZoomOption(
                                          direction,
                                        ) ? (
                                        <IconTransitionZoomOption
                                          option={direction}
                                          size={16}
                                        />
                                      ) : isSlideTransitionSplitOption(
                                          direction,
                                        ) ? (
                                        <IconTransitionSplitOption
                                          option={direction}
                                          size={16}
                                        />
                                      ) : isSlideTransitionOrientationOption(
                                          direction,
                                        ) ? (
                                        <IconTransitionOrientationOption
                                          option={direction}
                                          size={16}
                                        />
                                      ) : direction === "objects" ? (
                                        <Shapes size={16} aria-hidden="true" />
                                      ) : direction === "words" ? (
                                        <WholeWord size={16} aria-hidden="true" />
                                      ) : direction === "characters" ? (
                                        <CaseSensitive size={16} aria-hidden="true" />
                                      ) : isSlideTransitionSideDirection(
                                          direction,
                                        ) ||
                                        isSlideTransitionCornerOption(
                                          direction,
                                        ) ? (
                                        <IconTransitionDirection
                                          direction={direction}
                                          size={16}
                                        />
                                      ) : null}
                                      {transitionEffectOptionLabel(direction)}
                                    </DropdownMenuRadioItem>
                                  ),
                                )}
                              </DropdownMenuRadioGroup>
                            </RibbonDropdownMenuContent>
                          ) : null}
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>

                    <OfficeHomeRibbonGroup
                      label={t("presentation.transition.timingGroup")}
                      ribbon="transitions"
                      testId="presentation-ribbon-transition-timing-group"
                      className={`${documentHomeStyles.referencesNotesGroup} ${presentationChrome.effectsTimingGroup}`}
                    >
                      <PresentationTransitionTimingControls
                        slideModelId={activeSlide.id}
                        advanceOnClick={activeSlide.transitionAdvanceOnClick}
                        advanceAfterMs={activeSlide.transitionAdvanceAfterMs}
                        advanceEditable={slideTransitionAdvance.editable}
                        onAdvanceChange={slideTransitionAdvance.setAdvance}
                        sound={activeTransitionSound}
                        preservedSoundName={
                          activeTransitionSound === null &&
                          activeSlide.transitionSound &&
                          typeof (activeSlide.transitionSound as { name?: unknown }).name === "string"
                            ? (activeSlide.transitionSound as { name: string }).name
                            : null
                        }
                        soundEditable={transitionSoundEditable}
                        onSoundNone={clearTransitionSound}
                        onSoundBuiltin={applyBuiltinTransitionSound}
                        onSoundOther={beginTransitionSoundPicker}
                        gridClassName={presentationChrome.transitionTimingGrid}
                        soundIcon={<IconTimingSound />}
                        afterFieldClassName={presentationChrome.secondaryDurationField}
                        labels={{
                          duration: t("presentation.transitionDurationLabel"),
                          sound: t("presentation.transition.sound"),
                          soundNone: t("presentation.transition.soundNone"),
                          soundOther: t("presentation.transition.soundOther"),
                          moreSounds: t("presentation.transition.moreSounds"),
                          builtinSound: (id) => t(`presentation.transition.builtinSounds.${id}`),
                          advanceOnClick: t("presentation.transition.advanceOnClick"),
                          advanceAfter: t("presentation.transition.advanceAfter"),
                          advanceAfterSeconds: t("presentation.transition.advanceAfterSeconds"),
                        }}
                        duration={
                          <>
                            <IconTimingDuration />
                            <span>
                              {t("presentation.transitionDurationLabel")}
                            </span>
                            <OfficeHomeNumberField
                                  appearance="native-spinner"
                                  className={
                                    presentationChrome.secondaryDurationField
                                  }
                                  controlId="PPTCTL.RIBBON.TRANSITIONS.DURATION.FIELD"
                                  label={t("presentation.transitionDurationLabel")}
                                  min={MIN_TRANSITION_DURATION_SECONDS}
                                  max={MAX_TRANSITION_DURATION_SECONDS}
                                  step={TRANSITION_DURATION_STEP_SECONDS}
                                  inputMode="decimal"
                                  value={activeTransitionDurationInputValue}
                                  disabled={transitionDurationControlDisabled}
                                  onFocus={() =>
                                    beginTransitionDurationEdit(
                                      activeSlide.id,
                                      activeTransitionDurationMs,
                                    )
                                  }
                                  onChange={(event) =>
                                    changeTransitionDurationDraft(
                                      activeSlide.id,
                                      event.target.value,
                                    )
                                  }
                                  onBlur={() => {
                                    if (transitionDurationSkipBlurRef.current) {
                                      transitionDurationSkipBlurRef.current = false;
                                      return;
                                    }
                                    commitTransitionDurationDraft();
                                  }}
                                  onKeyDown={handleTransitionDurationKeyDown}
                                  data-unit="seconds"
                                  testId="presentation-ribbon-transition-duration"
                                  onDecrement={() => stepTransitionDuration("decrement")}
                                  onIncrement={() => stepTransitionDuration("increment")}
                                />
                          </>
                        }
                        applyAll={
                          <OfficeHomeCommand
                            controlId="PPTCTL.RIBBON.TRANSITIONS.APPLY_ALL"
                            disabled={!transitionApplyAllEditable}
                            icon={<IconTransitionApplyAll home size={24} />}
                            label={t("presentation.transitionApplyAll")}
                            onClick={applyTransitionToAll}
                            testId="presentation-ribbon-transition-apply-all"
                            variant="large"
                          />
                        }
                      />
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
