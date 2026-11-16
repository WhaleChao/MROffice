/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation-local command adapters.
 *
 * The shared Office feature registry owns capability/availability gates. This
 * module owns the editor-specific bridge from one typed domain transform to the
 * Presentation host's transaction + selection effects. Native serialization
 * remains outside the adapter and editor-specific by design.
 */

import { isBlockKeptReadOnly, isLockedFor } from "./presentation-text-case";
import {
  duplicatePresentationSlideAt,
  imageHasResettableFormatting,
  replaceImageSource,
  resetImagePicture,
  type DuplicablePresentationSlide,
} from "./presentation-conveniences";
import type { PreparedPresentationImageFile } from "./presentation-image-file";
import {
  IDENTITY_CROP,
  isIdentityCrop,
  normalizeSignedCropRect,
  pictureAspectRatioCropGeometry,
  type PictureAspectRatioCropGeometry,
  pictureFillCropRect,
  pictureFitCropRect,
  sameCropRect,
  type CropRect,
  type PictureCropAspectRatioId,
} from "./presentation-crop-group";
import { normalizeImageAltText } from "./markdown/policy";
import {
  MAX_IMAGE_BORDER_WIDTH,
  hasEditablePictureColor,
  hasEditablePictureCorrections,
  imageCorrectionDeltasFromPersisted,
  imageCorrectionStoredFromDelta,
  isImageCorrectionDelta,
  isNativeImageRecolorMode,
  normalizeImageBorderColor,
  normalizeImageBorderStyle,
  normalizeImageBorderWidth,
  normalizeImageOpacity,
  pictureCorrectionsSourceWithStoredUnit,
  type ImageBorderStyle,
  type NativeImageRecolorMode,
} from "./presentation-image-effects";
import {
  isZOrderStep,
  moveItemOneStep,
  type ZOrderStep,
} from "./presentation-z-order";
import {
  TABLE_FRAME_EMU_PER_CM,
  tableFrameSizeCm,
} from "./presentation-table-size";
import { EMU_PER_INCH } from "./presentation-slide-size";
import {
  applyResizeThroughZero,
  type BlockGeometry,
  type DragBounds,
  type ResizeHandle,
  type ResizeModifiers,
} from "./presentation-drag";
import { normalizeHiddenFlag } from "./presentation-slide-visibility";
import {
  applyTransitionToAllSlides,
  canAuthorSlideTransitionKind,
  defaultSlideTransitionDirection,
  defaultAuthoringTransitionDuration,
  hasEditableSlideTransitionDirectionSource,
  hasEditableSlideTransitionAdvanceSource,
  hasEditableSlideTransitionSource,
  isEditableTransitionAdvanceAfterMs,
  isEditableTransitionDurationMs,
  isAuthorableSlideTransition,
  isDirectionalSlideTransition,
  SLIDE_TRANSITION_AUTHORED_ZOOM_ELEMENT,
  isSlideTransitionDirectionForKind,
  isTransitionAdvanceAfterMs,
  normalizeSlideTransition,
  normalizeSlideTransitionDirection,
  normalizePresentationTransitionSound,
  samePresentationTransitionSound,
  transitionAdvanceOnClick,
  type AuthorableSlideTransition,
  type PresentationTransitionSound,
  type SlideTransitionDirection,
} from "./presentation-transitions";
import {
  createStep,
  editTimelineSequence,
  editTimelineStepEffectColor,
  editTimelineStepEffectOption,
  editTimelineStepObjectTrigger,
  editTimelineStepTextSequence,
  editTimelineStepTiming,
  hasEditableSlideAnimationSource,
  isLocallyAuthoredSlideAnimationSource,
  isPresentationAnimationEffectId,
  MAX_TIMELINE_STEPS,
  presentationAnimationEffect,
  timelineAnimationAuthoringCapability,
  type AnimationSequenceEditOperation,
  type AnimationTimingValue,
  type PresentationAnimationEffectId,
  type PresentationAnimationEffectOption,
  type TimelineAnimation,
  type TimelineStep,
  timelineTextParagraphIndexes,
  type TimelineTextSequence,
} from "./presentation-anim-timeline";
import { PRESENTATION_OBJECT_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";

export const DUPLICATE_SLIDE_COMMAND_ID = "slides.duplicate" as const;
export const SLIDE_VISIBILITY_FEATURE_ID = "slides.toggleHidden" as const;
export const SLIDE_VISIBILITY_ACTION =
  "presentation.slide.visibility.toggle" as const;
export const SLIDE_VISIBILITY_HANDLER_REF =
  "PresentationWindow#toggleSlideVisibility" as const;
export const PRESENTATION_SLIDE_VISIBILITY_SCHEMA_VERSION =
  "pptx-slide-visibility-v1" as const;
export const SLIDE_TRANSITION_FEATURE_ID = "slides.setTransition" as const;
export const SLIDE_TRANSITION_ACTION =
  "presentation.slide.transition.set" as const;
export const SLIDE_TRANSITION_HANDLER_REF =
  "PresentationWindow#setSlideTransition" as const;
export const SLIDE_TRANSITION_DURATION_FEATURE_ID =
  "slides.setTransitionDuration" as const;
export const SLIDE_TRANSITION_DURATION_ACTION =
  "presentation.slide.transition.duration.set" as const;
export const SLIDE_TRANSITION_DURATION_HANDLER_REF =
  "PresentationWindow#setTransitionDuration" as const;
export const SLIDE_TRANSITION_ADVANCE_FEATURE_ID =
  "slides.setTransitionAdvance" as const;
export const SLIDE_TRANSITION_ADVANCE_ACTION =
  "presentation.slide.transition.advance.set" as const;
export const SLIDE_TRANSITION_ADVANCE_HANDLER_REF =
  "PresentationWindow#setSlideAdvance" as const;
export const SLIDE_TRANSITION_APPLY_ALL_FEATURE_ID =
  "slides.applyTransitionToAll" as const;
export const SLIDE_TRANSITION_APPLY_ALL_ACTION =
  "presentation.slide.transition.apply-to-all" as const;
export const SLIDE_TRANSITION_APPLY_ALL_HANDLER_REF =
  "PresentationWindow#applyTransitionToAll" as const;
export const SLIDE_TRANSITION_SOUND_FEATURE_ID =
  "slides.setTransitionSound" as const;
export const SLIDE_TRANSITION_SOUND_ACTION =
  "presentation.slide.transition.sound.set" as const;
export const SLIDE_TRANSITION_SOUND_HANDLER_REF =
  "PresentationWindow#setTransitionSound" as const;
export const ADD_ANIMATION_EFFECT_FEATURE_ID =
  "animations.addEffect" as const;
export const ADD_ANIMATION_EFFECT_ACTION =
  "presentation.animation.effect.add" as const;
export const ADD_ANIMATION_EFFECT_HANDLER_REF =
  "PresentationWindow#addAnimationEffect" as const;
export const EDIT_ANIMATION_EFFECT_OPTION_FEATURE_ID =
  "animations.editEffectOption" as const;
export const EDIT_ANIMATION_EFFECT_OPTION_ACTION =
  "presentation.animation.effect-option.edit" as const;
export const EDIT_ANIMATION_EFFECT_OPTION_HANDLER_REF =
  "PresentationWindow#editAnimationEffectOption" as const;
export const EDIT_ANIMATION_SEQUENCE_FEATURE_ID =
  "animations.editSequence" as const;
export const EDIT_ANIMATION_SEQUENCE_ACTION =
  "presentation.animation.sequence.edit" as const;
export const EDIT_ANIMATION_SEQUENCE_HANDLER_REF =
  "PresentationWindow#editAnimationSequence" as const;
export const EDIT_ANIMATION_TIMING_FEATURE_ID =
  "animations.editTiming" as const;
export const EDIT_ANIMATION_TIMING_ACTION =
  "presentation.animation.timing.edit" as const;
export const EDIT_ANIMATION_TIMING_HANDLER_REF =
  "PresentationWindow#editAnimationTiming" as const;
export const INSERT_IMAGE_FILE_COMMAND_ID = "drawing.image" as const;
export const INSERT_IMAGE_FILE_ACTION =
  "presentation.image.insert-from-file" as const;
export const INSERT_IMAGE_FILE_HANDLER_REF =
  "PresentationWindow#insertImageFile" as const;
export const REPLACE_PICTURE_FILE_FEATURE_ID =
  "picture.changePicture.file" as const;
export const REPLACE_PICTURE_FILE_ACTION =
  "presentation.image.replace-from-file" as const;
export const REPLACE_PICTURE_FILE_HANDLER_REF =
  "PresentationWindow#replacePictureFromFile" as const;
export const PICTURE_CROP_FEATURE_ID = "picture.crop" as const;
export const PICTURE_CROP_FILL_FEATURE_ID = "picture.crop.fill" as const;
export const PICTURE_CROP_FIT_FEATURE_ID = "picture.crop.fit" as const;
export const PICTURE_CROP_ACTION = "presentation.image.crop" as const;
export const PICTURE_CROP_FILL_ACTION = "presentation.image.crop-fill" as const;
export const PICTURE_CROP_FIT_ACTION = "presentation.image.crop-fit" as const;
export const PICTURE_CROP_HANDLER_REF =
  "PresentationWindow#cropPicture" as const;
export const PICTURE_CROP_FILL_HANDLER_REF =
  "PresentationWindow#cropPictureFill" as const;
export const PICTURE_CROP_FIT_HANDLER_REF =
  "PresentationWindow#cropPictureFit" as const;
export const PICTURE_ALT_TEXT_FEATURE_ID = "picture.altText" as const;
export const PICTURE_ALT_TEXT_ACTION =
  "presentation.image.edit-alt-text" as const;
export const PICTURE_ALT_TEXT_HANDLER_REF =
  "PresentationWindow#editPictureAltText" as const;
export const PICTURE_BORDER_FEATURE_ID = "picture.border" as const;
export const PICTURE_BORDER_ACTION = "presentation.image.edit-border" as const;
export const PICTURE_BORDER_HANDLER_REF =
  "PresentationWindow#editPictureBorder" as const;
export const PICTURE_TRANSPARENCY_FEATURE_ID = "picture.transparency" as const;
export const PICTURE_TRANSPARENCY_ACTION =
  "presentation.image.edit-transparency" as const;
export const PICTURE_TRANSPARENCY_HANDLER_REF =
  "PresentationWindow#editPictureTransparency" as const;
export const PICTURE_CORRECTIONS_FEATURE_ID = "picture.corrections" as const;
export const PICTURE_CORRECTIONS_ACTION =
  "presentation.image.edit-corrections" as const;
export const PICTURE_CORRECTIONS_HANDLER_REF =
  "PresentationWindow#editPictureCorrections" as const;
export const PICTURE_COLOR_FEATURE_ID = "picture.color.recolor" as const;
export const PICTURE_COLOR_ACTION = "presentation.image.edit-color" as const;
export const PICTURE_COLOR_HANDLER_REF =
  "PresentationWindow#editPictureColor" as const;
export const PICTURE_BRING_FORWARD_FEATURE_ID =
  "picture.arrange.bringForward" as const;
export const PICTURE_SEND_BACKWARD_FEATURE_ID =
  "picture.arrange.sendBackward" as const;
export const PICTURE_BRING_FORWARD_ACTION =
  "presentation.image.bring-forward" as const;
export const PICTURE_SEND_BACKWARD_ACTION =
  "presentation.image.send-backward" as const;
export const PICTURE_BRING_FORWARD_HANDLER_REF =
  "PresentationWindow#bringPictureForward" as const;
export const PICTURE_SEND_BACKWARD_HANDLER_REF =
  "PresentationWindow#sendPictureBackward" as const;
export const PICTURE_SIZE_WIDTH_FEATURE_ID = "picture.size.width" as const;
export const PICTURE_SIZE_HEIGHT_FEATURE_ID = "picture.size.height" as const;
export const PICTURE_SIZE_WIDTH_ACTION =
  "presentation.image.resize-width" as const;
export const PICTURE_SIZE_HEIGHT_ACTION =
  "presentation.image.resize-height" as const;
export const PICTURE_SIZE_WIDTH_HANDLER_REF =
  "PresentationWindow#resizePictureWidth" as const;
export const PICTURE_SIZE_HEIGHT_HANDLER_REF =
  "PresentationWindow#resizePictureHeight" as const;
const EMU_PER_96_DPI_PIXEL = 9_525;
export const PRESENTATION_PICTURE_SIZE_MIN_EMU = 1;
export const PRESENTATION_PICTURE_SIZE_MAX_EMU = PRESENTATION_OBJECT_GEOMETRY_MAX_EMU;
export const PRESENTATION_PICTURE_SIZE_STEP_EMU = 9_144;
export const PRESENTATION_PICTURE_SIZE_SCHEMA_VERSION =
  "pptx-picture-size-v1" as const;

export interface PresentationCommandAdapter<Id extends string> {
  readonly featureId: Id;
  /** `false` means the command failed closed without any host mutation. */
  readonly execute: () => boolean;
}

export interface DuplicateSlideCommandAccessors<
  T extends DuplicablePresentationSlide,
> {
  getSlides: () => readonly T[];
  getActiveIndex: () => number;
  commitSlides: (slides: T[]) => void;
  /**
   * Commit a duplicate together with deck state keyed by the source slide's
   * id (its own header/footer). Omitted, the slides alone are committed.
   */
  commitDuplicate?: (slides: T[], sourceId: string, cloneId: string) => void;
  selectSlide: (index: number) => void;
  nextId: () => string;
  /** Native hosts dispatch an authenticated async plan; undefined uses the local transform. */
  duplicateNativeSlide?: (slide: T) => boolean | undefined;
}

/**
 * Bind `slides.duplicate` to live Presentation host state. The adapter commits
 * exactly once and changes selection only after the commit succeeds. It owns no
 * UI source, so Ribbon, shortcut, and context-menu callers can reuse the same
 * `execute` function without acquiring a second mutation or Undo path.
 */
export function createDuplicateSlideCommandAdapter<
  T extends DuplicablePresentationSlide,
>(
  accessors: DuplicateSlideCommandAccessors<T>,
): PresentationCommandAdapter<typeof DUPLICATE_SLIDE_COMMAND_ID> {
  const execute = (): boolean => {
    const source = accessors.getSlides()[accessors.getActiveIndex()];
    if (source && accessors.duplicateNativeSlide) {
      const dispatched = accessors.duplicateNativeSlide(source);
      if (dispatched !== undefined) return dispatched;
    }
    const result = duplicatePresentationSlideAt(
      accessors.getSlides(),
      accessors.getActiveIndex(),
      accessors.nextId,
    );
    if (!result) return false;
    if (source && accessors.commitDuplicate) {
      accessors.commitDuplicate(result.slides, source.id, result.clone.id);
    } else {
      accessors.commitSlides(result.slides);
    }
    accessors.selectSlide(result.insertedIndex);
    return true;
  };

  return Object.freeze({
    featureId: DUPLICATE_SLIDE_COMMAND_ID,
    execute,
  });
}

export interface PresentationSlideVisibilitySlide {
  readonly id: string;
  readonly hidden?: unknown;
  readonly source?: unknown;
}

export interface PresentationSlideVisibilityCommandPayload {
  readonly slideModelId: string;
}

export interface PresentationSlideVisibilityCommandAccessors<
  T extends PresentationSlideVisibilitySlide,
> {
  getSlides: () => readonly T[];
  getActiveIndex: () => number;
  commitSlides: (slides: T[]) => void;
}

export interface PresentationSlideVisibilityCommandAdapter {
  readonly featureId: typeof SLIDE_VISIBILITY_FEATURE_ID;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationSlideVisibilityCommandPayload,
  ) => boolean;
}

function hasEditableSlideVisibility(
  slide: PresentationSlideVisibilitySlide,
): boolean {
  if (slide.hidden !== undefined && typeof slide.hidden !== "boolean") {
    return false;
  }
  if (slide.source === undefined || slide.source === null) {
    return true;
  }
  if (typeof slide.source !== "object" || Array.isArray(slide.source)) {
    return false;
  }
  const source = slide.source as Record<string, unknown>;
  const nativeSlideId = source.slideId;
  if (nativeSlideId === undefined) {
    return ![
      "slideVisibilitySchemaVersion",
      "slideVisibilityDisposition",
      "slideVisibilityOriginalHidden",
    ].some((key) => Object.hasOwn(source, key));
  }
  return (
    typeof nativeSlideId === "string" &&
    /^sld[1-9][0-9]*$/.test(nativeSlideId) &&
    source.slideVisibilitySchemaVersion ===
      PRESENTATION_SLIDE_VISIBILITY_SCHEMA_VERSION &&
    source.slideVisibilityDisposition === "native" &&
    typeof source.slideVisibilityOriginalHidden === "boolean"
  );
}

/** One live-target Hide Slide / Show Slide transaction. */
export function createPresentationSlideVisibilityCommandAdapter<
  T extends PresentationSlideVisibilitySlide,
>(
  accessors: PresentationSlideVisibilityCommandAccessors<T>,
): PresentationSlideVisibilityCommandAdapter {
  const target = () => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    if (!slide || !slide.id || !hasEditableSlideVisibility(slide)) return null;
    return { slides, slide, slideIndex };
  };

  const canExecute = (): boolean => target() !== null;
  const execute = (
    payload: PresentationSlideVisibilityCommandPayload,
  ): boolean => {
    const resolved = target();
    if (!resolved || resolved.slide.id !== payload.slideModelId) return false;
    const nextSlide = { ...resolved.slide } as T;
    if (normalizeHiddenFlag(resolved.slide.hidden)) {
      delete (nextSlide as { hidden?: unknown }).hidden;
    } else {
      (nextSlide as { hidden?: unknown }).hidden = true;
    }
    const nextSlides = resolved.slides.map((slide, index) =>
      index === resolved.slideIndex ? nextSlide : slide,
    ) as T[];
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_VISIBILITY_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationSlideTransitionSlide {
  readonly id: string;
  readonly transition?: unknown;
  readonly transitionDurationMs?: unknown;
  readonly transitionDirection?: unknown;
  readonly transitionZoomElement?: unknown;
  readonly transitionAdvanceOnClick?: unknown;
  readonly transitionAdvanceAfterMs?: unknown;
  readonly transitionSound?: unknown;
  readonly source?: unknown;
}

export interface PresentationSlideTransitionCommandPayload {
  readonly slideModelId: string;
  readonly kind: AuthorableSlideTransition;
  /** Present only for an Effect Options command; gallery selection resets it. */
  readonly direction?: SlideTransitionDirection;
}

export interface PresentationSlideTransitionCommandAccessors<
  T extends PresentationSlideTransitionSlide,
> {
  getSlides: () => readonly T[];
  getActiveIndex: () => number;
  commitSlides: (slides: T[]) => void;
}

export interface PresentationSlideTransitionCommandAdapter {
  readonly featureId: typeof SLIDE_TRANSITION_FEATURE_ID;
  readonly canExecute: (
    kind?: AuthorableSlideTransition,
    direction?: SlideTransitionDirection,
  ) => boolean;
  readonly execute: (
    payload: PresentationSlideTransitionCommandPayload,
  ) => boolean;
}

/** One live-target, one-commit transition-gallery command. */
export function createPresentationSlideTransitionCommandAdapter<
  T extends PresentationSlideTransitionSlide,
>(
  accessors: PresentationSlideTransitionCommandAccessors<T>,
): PresentationSlideTransitionCommandAdapter {
  const target = (
    kind?: AuthorableSlideTransition,
    direction?: SlideTransitionDirection,
  ) => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    if (
      !slide ||
      !slide.id ||
      (kind === undefined
        ? !hasEditableSlideTransitionSource(slide.source) &&
          !canAuthorSlideTransitionKind(slide.source, "morph")
        : !canAuthorSlideTransitionKind(slide.source, kind))
    ) {
      return null;
    }
    if (
      direction !== undefined &&
      (!isDirectionalSlideTransition(kind) ||
        !isSlideTransitionDirectionForKind(kind, direction) ||
        normalizeSlideTransition(slide.transition) !== kind ||
        !hasEditableSlideTransitionDirectionSource(slide.source))
    ) {
      return null;
    }
    return { slide, slideIndex, slides };
  };

  const canExecute = (
    kind?: AuthorableSlideTransition,
    direction?: SlideTransitionDirection,
  ): boolean => target(kind, direction) !== null;
  const execute = (
    payload: PresentationSlideTransitionCommandPayload,
  ): boolean => {
    if (!isAuthorableSlideTransition(payload.kind)) return false;
    if (
      payload.direction !== undefined &&
      !isSlideTransitionDirectionForKind(payload.kind, payload.direction)
    ) {
      return false;
    }
    const resolved = target(payload.kind, payload.direction);
    if (!resolved || resolved.slide.id !== payload.slideModelId) return false;

    const currentKind = normalizeSlideTransition(resolved.slide.transition);
    if (payload.direction !== undefined) {
      if (
        normalizeSlideTransitionDirection(
          currentKind,
          resolved.slide.transitionDirection,
        ) === payload.direction
      ) {
        return false;
      }
      const nextSlide = {
        ...resolved.slide,
        transitionDirection: payload.direction,
      } as T;
      const nextSlides = resolved.slides.map((slide, index) =>
        index === resolved.slideIndex ? nextSlide : slide,
      ) as T[];
      accessors.commitSlides(nextSlides);
      return true;
    }

    const currentDuration = resolved.slide.transitionDurationMs;
    const desiredDuration = defaultAuthoringTransitionDuration(payload.kind);
    const desiredDirection = defaultSlideTransitionDirection(payload.kind);
    const currentDirection = normalizeSlideTransitionDirection(
      currentKind,
      resolved.slide.transitionDirection,
    );
    if (
      payload.kind === "none" &&
      currentKind === "none" &&
      currentDuration === undefined &&
      resolved.slide.transitionDirection === undefined
    ) {
      return false;
    }
    if (
      payload.kind !== "none" &&
      currentKind === payload.kind &&
      currentDuration === desiredDuration &&
      currentDirection === desiredDirection
    ) {
      return false;
    }

    const nextSlide = { ...resolved.slide } as T;
    // Which element spells the Zoom. A new Zoom — picked fresh or switched to
    // from another effect — is written as PowerPoint's own `p14:warp`. A slide
    // already on Zoom keeps whichever element it carries (the 2006 `p:zoom`
    // when it names none), and picking any other effect leaves it behind.
    if (payload.kind !== "zoom") {
      delete (nextSlide as { transitionZoomElement?: unknown })
        .transitionZoomElement;
    } else if (currentKind !== "zoom") {
      (nextSlide as { transitionZoomElement?: unknown }).transitionZoomElement =
        SLIDE_TRANSITION_AUTHORED_ZOOM_ELEMENT;
    }
    if (payload.kind === "none") {
      delete (nextSlide as { transition?: unknown }).transition;
      delete (nextSlide as { transitionDurationMs?: unknown })
        .transitionDurationMs;
      delete (nextSlide as { transitionDirection?: unknown })
        .transitionDirection;
    } else {
      (nextSlide as { transition?: unknown }).transition = payload.kind;
      (nextSlide as { transitionDurationMs?: unknown }).transitionDurationMs =
        desiredDuration;
      if (desiredDirection === null) {
        delete (nextSlide as { transitionDirection?: unknown })
          .transitionDirection;
      } else {
        (nextSlide as { transitionDirection?: unknown }).transitionDirection =
          desiredDirection;
      }
    }
    const nextSlides = resolved.slides.map((slide, index) =>
      index === resolved.slideIndex ? nextSlide : slide,
    ) as T[];
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_TRANSITION_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationSlideTransitionDurationCommandPayload {
  readonly slideModelId: string;
  readonly durationMs: number;
}

export type PresentationSlideTransitionDurationCommandAccessors<
  T extends PresentationSlideTransitionSlide,
> = PresentationSlideTransitionCommandAccessors<T>;

export interface PresentationSlideTransitionDurationCommandAdapter {
  readonly featureId: typeof SLIDE_TRANSITION_DURATION_FEATURE_ID;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationSlideTransitionDurationCommandPayload,
  ) => boolean;
}

/** One live-target, one-commit Duration field/stepper command. */
export function createPresentationSlideTransitionDurationCommandAdapter<
  T extends PresentationSlideTransitionSlide,
>(
  accessors: PresentationSlideTransitionDurationCommandAccessors<T>,
): PresentationSlideTransitionDurationCommandAdapter {
  const target = () => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    const kind = normalizeSlideTransition(slide?.transition);
    if (
      !slide?.id ||
      !isAuthorableSlideTransition(kind) ||
      kind === "none" ||
      !canAuthorSlideTransitionKind(slide.source, kind)
    ) {
      return null;
    }
    return { slide, slideIndex, slides };
  };

  const canExecute = (): boolean => target() !== null;
  const execute = (
    payload: PresentationSlideTransitionDurationCommandPayload,
  ): boolean => {
    if (!isEditableTransitionDurationMs(payload.durationMs)) return false;
    const resolved = target();
    if (
      !resolved ||
      resolved.slide.id !== payload.slideModelId ||
      resolved.slide.transitionDurationMs === payload.durationMs
    ) {
      return false;
    }
    const nextSlide = {
      ...resolved.slide,
      transitionDurationMs: payload.durationMs,
    } as T;
    const nextSlides = resolved.slides.map((slide, index) =>
      index === resolved.slideIndex ? nextSlide : slide,
    ) as T[];
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_TRANSITION_DURATION_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationSlideTransitionAdvanceCommandPayload {
  readonly slideModelId: string;
  readonly advanceOnClick: boolean;
  readonly advanceAfterMs: number | null;
}

export type PresentationSlideTransitionAdvanceCommandAccessors<
  T extends PresentationSlideTransitionSlide,
> = PresentationSlideTransitionCommandAccessors<T>;

export interface PresentationSlideTransitionAdvanceCommandAdapter {
  readonly featureId: typeof SLIDE_TRANSITION_ADVANCE_FEATURE_ID;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationSlideTransitionAdvanceCommandPayload,
  ) => boolean;
}

/** One live-target, one-commit On Click + After transition command. */
export function createPresentationSlideTransitionAdvanceCommandAdapter<
  T extends PresentationSlideTransitionSlide,
>(
  accessors: PresentationSlideTransitionAdvanceCommandAccessors<T>,
): PresentationSlideTransitionAdvanceCommandAdapter {
  const target = () => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    if (
      !slide?.id ||
      !hasEditableSlideTransitionAdvanceSource(slide.source) ||
      (slide.transitionAdvanceOnClick !== undefined &&
        typeof slide.transitionAdvanceOnClick !== "boolean") ||
      (slide.transitionAdvanceAfterMs !== undefined &&
        !isTransitionAdvanceAfterMs(slide.transitionAdvanceAfterMs))
    ) {
      return null;
    }
    return { slide, slideIndex, slides };
  };

  const canExecute = (): boolean => target() !== null;
  const execute = (
    payload: PresentationSlideTransitionAdvanceCommandPayload,
  ): boolean => {
    if (
      typeof payload.advanceOnClick !== "boolean" ||
      (payload.advanceAfterMs !== null &&
        !isEditableTransitionAdvanceAfterMs(payload.advanceAfterMs))
    ) {
      return false;
    }
    const resolved = target();
    if (
      !resolved ||
      resolved.slide.id !== payload.slideModelId ||
      (transitionAdvanceOnClick(resolved.slide.transitionAdvanceOnClick) ===
        payload.advanceOnClick &&
        (resolved.slide.transitionAdvanceAfterMs ?? null) ===
          payload.advanceAfterMs)
    ) {
      return false;
    }

    const nextSlide = { ...resolved.slide } as T;
    if (payload.advanceOnClick) {
      delete (nextSlide as { transitionAdvanceOnClick?: unknown })
        .transitionAdvanceOnClick;
    } else {
      (
        nextSlide as {
          transitionAdvanceOnClick?: unknown;
        }
      ).transitionAdvanceOnClick = false;
    }
    if (payload.advanceAfterMs === null) {
      delete (nextSlide as { transitionAdvanceAfterMs?: unknown })
        .transitionAdvanceAfterMs;
    } else {
      (
        nextSlide as {
          transitionAdvanceAfterMs?: unknown;
        }
      ).transitionAdvanceAfterMs = payload.advanceAfterMs;
    }
    const nextSlides = resolved.slides.map((slide, index) =>
      index === resolved.slideIndex ? nextSlide : slide,
    ) as T[];
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_TRANSITION_ADVANCE_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationSlideTransitionSoundCommandPayload {
  readonly slideModelId: string;
  readonly sound: PresentationTransitionSound | null;
}

export interface PresentationSlideTransitionSoundCommandAdapter {
  readonly featureId: typeof SLIDE_TRANSITION_SOUND_FEATURE_ID;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationSlideTransitionSoundCommandPayload,
  ) => boolean;
}

/** One live-target, one-commit embedded-WAV transition sound command. */
export function createPresentationSlideTransitionSoundCommandAdapter<
  T extends PresentationSlideTransitionSlide,
>(
  accessors: PresentationSlideTransitionCommandAccessors<T>,
): PresentationSlideTransitionSoundCommandAdapter {
  const target = () => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    const source =
      slide?.source &&
      typeof slide.source === "object" &&
      !Array.isArray(slide.source)
        ? (slide.source as Record<string, unknown>)
        : null;
    const importedSoundEditable =
      source?.slideId === undefined ||
      source.transitionSoundDisposition === "absent" ||
      source.transitionSoundDisposition === "native";
    if (
      !slide?.id ||
      !hasEditableSlideTransitionSource(slide.source) ||
      !importedSoundEditable ||
      (slide.transitionSound !== undefined &&
        normalizePresentationTransitionSound(slide.transitionSound) === null)
    ) {
      return null;
    }
    return { slide, slideIndex, slides };
  };
  const canExecute = (): boolean => target() !== null;
  const execute = (
    payload: PresentationSlideTransitionSoundCommandPayload,
  ): boolean => {
    const sound =
      payload.sound === null
        ? null
        : normalizePresentationTransitionSound(payload.sound);
    const resolved = target();
    if (
      !resolved ||
      resolved.slide.id !== payload.slideModelId ||
      (payload.sound !== null && sound === null) ||
      samePresentationTransitionSound(resolved.slide.transitionSound, sound)
    ) {
      return false;
    }
    const nextSlide = { ...resolved.slide } as T;
    if (sound === null) {
      delete (nextSlide as { transitionSound?: unknown }).transitionSound;
    } else {
      (nextSlide as { transitionSound?: unknown }).transitionSound = sound;
    }
    accessors.commitSlides(
      resolved.slides.map((slide, index) =>
        index === resolved.slideIndex ? nextSlide : slide,
      ) as T[],
    );
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_TRANSITION_SOUND_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationSlideTransitionApplyAllCommandPayload {
  readonly sourceSlideModelId: string;
}

export type PresentationSlideTransitionApplyAllCommandAccessors<
  T extends PresentationSlideTransitionSlide,
> = PresentationSlideTransitionCommandAccessors<T>;

export interface PresentationSlideTransitionApplyAllCommandAdapter {
  readonly featureId: typeof SLIDE_TRANSITION_APPLY_ALL_FEATURE_ID;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationSlideTransitionApplyAllCommandPayload,
  ) => boolean;
}

/** One live-source, one atomic deck commit for the complete transition state. */
export function createPresentationSlideTransitionApplyAllCommandAdapter<
  T extends PresentationSlideTransitionSlide,
>(
  accessors: PresentationSlideTransitionApplyAllCommandAccessors<T>,
): PresentationSlideTransitionApplyAllCommandAdapter {
  const resolve = () => {
    const slides = accessors.getSlides();
    const activeIndex = accessors.getActiveIndex();
    if (!Number.isInteger(activeIndex) || activeIndex < 0) return null;
    const active = slides[activeIndex];
    if (!active?.id) return null;
    const nextSlides = applyTransitionToAllSlides(slides, active.id);
    return nextSlides === slides
      ? null
      : { active, nextSlides: nextSlides as T[] };
  };

  const canExecute = (): boolean => resolve() !== null;
  const execute = (
    payload: PresentationSlideTransitionApplyAllCommandPayload,
  ): boolean => {
    if (!payload.sourceSlideModelId) return false;
    const resolved = resolve();
    if (!resolved || resolved.active.id !== payload.sourceSlideModelId) {
      return false;
    }
    accessors.commitSlides(resolved.nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: SLIDE_TRANSITION_APPLY_ALL_FEATURE_ID,
    canExecute,
    execute,
  });
}

export interface PresentationAnimationEffectBlock {
  readonly id: string;
  readonly type?: string;
  readonly text?: unknown;
  readonly source?: unknown;
  readonly locked?: boolean;
}

/** Exact target boundary for native color emphasis effects. */
export function presentationAnimationEffectTargetEligible(
  kind: TimelineAnimation,
  block: PresentationAnimationEffectBlock,
): boolean {
  const effect = presentationAnimationEffect(kind);
  if (!effect) return false;
  if (effect.nativeBehavior !== "color") return true;
  if (
    effect.nativeColorAttribute === "fillcolor" ||
    effect.nativeColorAttribute === "stroke.color"
  ) {
    return block.type === "shape";
  }
  return (
    effect.nativeColorAttribute === "style.color" &&
    (block.type === "text" || block.type === "shape") &&
    typeof block.text === "string" &&
    block.text.trim().length > 0
  );
}

export interface PresentationAnimationEffectSlide {
  readonly id: string;
  readonly blocks: readonly PresentationAnimationEffectBlock[];
  readonly animationSequence?: readonly TimelineStep[];
  readonly source?: unknown;
}

export interface PresentationAnimationEffectAppendCommandPayload {
  readonly operation?: "append";
  readonly slideModelId: string;
  readonly blockModelId: string;
  readonly kind: PresentationAnimationEffectId;
}

export interface PresentationAnimationEffectOptionCommandPayload {
  readonly operation: "set-effect-option";
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
  readonly option: PresentationAnimationEffectOption;
}

export interface PresentationAnimationEffectColorCommandPayload {
  readonly operation: "set-effect-color";
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
  /** Null restores PowerPoint's theme accent2 default. */
  readonly color: string | null;
}

export interface PresentationAnimationObjectTriggerCommandPayload {
  readonly operation: "set-object-trigger";
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
  /** Null restores ordinary sequence timing. */
  readonly triggerBlockId: string | null;
}

/** Effect Options > Sequence on a text box's effect (PowerPoint). */
export interface PresentationAnimationTextSequenceCommandPayload {
  readonly operation: "set-text-sequence";
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
  readonly sequence: TimelineTextSequence;
}

export type PresentationAnimationEffectCommandPayload =
  | PresentationAnimationEffectAppendCommandPayload
  | PresentationAnimationEffectOptionCommandPayload
  | PresentationAnimationEffectColorCommandPayload
  | PresentationAnimationObjectTriggerCommandPayload
  | PresentationAnimationTextSequenceCommandPayload;

/** The next sequence one exact occurrence edit produces, or null. */
function editedAnimationOccurrence(
  sequence: readonly TimelineStep[],
  block: PresentationAnimationEffectBlock,
  blockIds: ReadonlySet<string>,
  payload: Exclude<
    PresentationAnimationEffectCommandPayload,
    PresentationAnimationEffectAppendCommandPayload
  >,
): TimelineStep[] | null {
  switch (payload.operation) {
    case "set-effect-option":
      return editTimelineStepEffectOption(
        sequence,
        payload.index,
        payload.expectedStep,
        payload.option,
      );
    case "set-effect-color":
      return editTimelineStepEffectColor(
        sequence,
        payload.index,
        payload.expectedStep,
        payload.color,
      );
    case "set-object-trigger":
      return editTimelineStepObjectTrigger(
        sequence,
        payload.index,
        payload.expectedStep,
        payload.triggerBlockId,
        blockIds,
      );
    case "set-text-sequence":
      return block.type === "text"
        ? editTimelineStepTextSequence(
            sequence,
            payload.index,
            payload.expectedStep,
            payload.sequence,
            timelineTextParagraphIndexes(block.text),
          )
        : null;
  }
}

export interface PresentationAnimationEffectCommandAccessors<
  T extends PresentationAnimationEffectSlide,
> {
  getSlides: () => readonly T[];
  getActiveIndex: () => number;
  getSelectedBlockId: () => string | null;
  commitSlides: (slides: T[]) => void;
}

export interface PresentationAnimationEffectCommandAdapter {
  readonly featureIds: Readonly<{
    append: typeof ADD_ANIMATION_EFFECT_FEATURE_ID;
    effectOption: typeof EDIT_ANIMATION_EFFECT_OPTION_FEATURE_ID;
  }>;
  readonly canExecute: (
    payload?: PresentationAnimationEffectCommandPayload,
  ) => boolean;
  readonly execute: (
    payload: PresentationAnimationEffectCommandPayload,
  ) => boolean;
}

/** ANI01's single live-target append command. */
export function createPresentationAnimationEffectCommandAdapter<
  T extends PresentationAnimationEffectSlide,
>(
  accessors: PresentationAnimationEffectCommandAccessors<T>,
): PresentationAnimationEffectCommandAdapter {
  const target = (kind: PresentationAnimationEffectId = "appear") => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    const blockModelId = accessors.getSelectedBlockId();
    if (
      !Number.isInteger(slideIndex) ||
      slideIndex < 0 ||
      typeof blockModelId !== "string" ||
      !blockModelId
    ) {
      return null;
    }
    const slide = slides[slideIndex];
    const block = slide?.blocks.find((item) => item.id === blockModelId);
    if (
      !slide?.id ||
      !block ||
      !presentationAnimationEffectTargetEligible(kind, block) ||
      !timelineAnimationAuthoringCapability({
        slideSource: slide.source,
        blockSource: block.source,
        animation: kind,
      }).enabled
    ) {
      return null;
    }
    const sequence = slide.animationSequence ?? [];
    if (!Array.isArray(sequence) || sequence.length >= MAX_TIMELINE_STEPS) {
      return null;
    }
    return { blockModelId, sequence, slide, slideIndex, slides };
  };

  const objectTriggerEligible = (
    slide: T,
    current: TimelineStep,
    triggerBlockId: string | null,
  ): boolean => {
    if (triggerBlockId === null) return current.triggerBlockId !== undefined;
    const candidate = slide.blocks.find((block) => block.id === triggerBlockId);
    if (!candidate || candidate.id === current.blockId) return false;
    if (isLocallyAuthoredSlideAnimationSource(slide.source)) return true;
    const source =
      candidate.source &&
      typeof candidate.source === "object" &&
      !Array.isArray(candidate.source)
        ? (candidate.source as Record<string, unknown>)
        : null;
    return Boolean(
      source &&
        typeof source.slideId === "string" &&
        typeof source.shapeId === "string" &&
        /^shp[1-9][0-9]*$/.test(source.shapeId),
    );
  };

  const execute = (
    payload: PresentationAnimationEffectCommandPayload,
  ): boolean => {
    if (
      payload.operation === "set-effect-option" ||
      payload.operation === "set-effect-color" ||
      payload.operation === "set-object-trigger" ||
      payload.operation === "set-text-sequence"
    ) {
      const slides = accessors.getSlides();
      const slideIndex = accessors.getActiveIndex();
      const slide = slides[slideIndex];
      const sequence = slide?.animationSequence ?? [];
      const current = sequence[payload.index];
      const block = slide?.blocks.find(
        (item) => item.id === current?.blockId,
      );
      if (
        !slide ||
        slide.id !== payload.slideModelId ||
        !Array.isArray(sequence) ||
        !current ||
        !block ||
        (payload.operation === "set-object-trigger" &&
          !objectTriggerEligible(
            slide,
            current,
            payload.triggerBlockId,
          )) ||
        (payload.operation === "set-effect-color" &&
          (presentationAnimationEffect(current.animation)?.nativeBehavior !==
            "color" ||
            !presentationAnimationEffectTargetEligible(
              current.animation,
              block,
            ))) ||
        !timelineAnimationAuthoringCapability({
          slideSource: slide.source,
          blockSource: block.source,
          animation: current.animation,
        }).enabled
      ) {
        return false;
      }
      const nextSequence = editedAnimationOccurrence(
        sequence,
        block,
        new Set(slide.blocks.map((candidate) => candidate.id)),
        payload,
      );
      if (!nextSequence) return false;
      const nextSlide = { ...slide, animationSequence: nextSequence } as T;
      accessors.commitSlides(
        slides.map((candidate, index) =>
          index === slideIndex ? nextSlide : candidate,
        ) as T[],
      );
      return true;
    }
    if (!isPresentationAnimationEffectId(payload.kind)) return false;
    const resolved = target(payload.kind);
    if (
      !resolved ||
      resolved.slide.id !== payload.slideModelId ||
      resolved.blockModelId !== payload.blockModelId
    ) {
      return false;
    }
    const nextSlide = {
      ...resolved.slide,
      animationSequence: [
        ...resolved.sequence,
        createStep(resolved.blockModelId, payload.kind),
      ],
    } as T;
    accessors.commitSlides(
      resolved.slides.map((slide, index) =>
        index === resolved.slideIndex ? nextSlide : slide,
      ) as T[],
    );
    return true;
  };

  return Object.freeze({
    featureIds: Object.freeze({
      append: ADD_ANIMATION_EFFECT_FEATURE_ID,
      effectOption: EDIT_ANIMATION_EFFECT_OPTION_FEATURE_ID,
    }),
    canExecute: (payload?: PresentationAnimationEffectCommandPayload) => {
      if (
        !payload ||
        (payload.operation !== "set-effect-option" &&
          payload.operation !== "set-effect-color" &&
          payload.operation !== "set-object-trigger" &&
          payload.operation !== "set-text-sequence")
      ) {
        return target(payload?.kind) !== null;
      }
      const slides = accessors.getSlides();
      const slideIndex = accessors.getActiveIndex();
      const slide = slides[slideIndex];
      const sequence = slide?.animationSequence ?? [];
      const current = sequence[payload.index];
      const block = slide?.blocks.find(
        (item) => item.id === current?.blockId,
      );
      return Boolean(
        slide &&
          slide.id === payload.slideModelId &&
          Array.isArray(sequence) &&
          current &&
          block &&
          (payload.operation !== "set-object-trigger" ||
            objectTriggerEligible(
              slide,
              current,
              payload.triggerBlockId,
            )) &&
          (payload.operation !== "set-effect-color" ||
            (presentationAnimationEffect(current.animation)?.nativeBehavior ===
              "color" &&
              presentationAnimationEffectTargetEligible(
                current.animation,
                block,
              ))) &&
          timelineAnimationAuthoringCapability({
            slideSource: slide.source,
            blockSource: block.source,
            animation: current.animation,
          }).enabled &&
          editedAnimationOccurrence(
            sequence,
            block,
            new Set(slide.blocks.map((candidate) => candidate.id)),
            payload,
          ) !== null,
      );
    },
    execute,
  });
}

export interface PresentationAnimationSequenceCommandPayload {
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
  readonly operation: AnimationSequenceEditOperation;
}

export interface PresentationAnimationSequenceCommandAdapter {
  readonly featureId: typeof EDIT_ANIMATION_SEQUENCE_FEATURE_ID;
  readonly canExecute: (
    payload?: PresentationAnimationSequenceCommandPayload,
  ) => boolean;
  readonly execute: (
    payload: PresentationAnimationSequenceCommandPayload,
  ) => boolean;
}

/** ANI02's live exact-occurrence move/remove authority. */
export function createPresentationAnimationSequenceCommandAdapter<
  T extends PresentationAnimationEffectSlide,
>(
  accessors: Omit<
    PresentationAnimationEffectCommandAccessors<T>,
    "getSelectedBlockId"
  >,
): PresentationAnimationSequenceCommandAdapter {
  const resolve = (payload?: PresentationAnimationSequenceCommandPayload) => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    const sequence = slide?.animationSequence;
    if (
      !slide?.id ||
      !hasEditableSlideAnimationSource(slide.source) ||
      !Array.isArray(sequence) ||
      sequence.length === 0 ||
      sequence.length > MAX_TIMELINE_STEPS
    ) {
      return null;
    }
    if (!payload) {
      return { nextSequence: null, sequence, slide, slideIndex, slides };
    }
    if (payload.slideModelId !== slide.id) return null;
    const currentStep = sequence[payload.index];
    if (
      !currentStep ||
      !slide.blocks.some((block) => block.id === currentStep.blockId)
    ) {
      return null;
    }
    const nextSequence = editTimelineSequence(
      sequence,
      payload.index,
      payload.expectedStep,
      payload.operation,
    );
    return nextSequence === null
      ? null
      : { nextSequence, sequence, slide, slideIndex, slides };
  };

  const execute = (
    payload: PresentationAnimationSequenceCommandPayload,
  ): boolean => {
    const resolved = resolve(payload);
    if (!resolved?.nextSequence) return false;
    const { animationSequence: _previousSequence, ...slideWithoutSequence } =
      resolved.slide;
    const nextSlide = (resolved.nextSequence.length > 0
      ? {
          ...slideWithoutSequence,
          animationSequence: [...resolved.nextSequence],
        }
      : slideWithoutSequence) as T;
    accessors.commitSlides(
      resolved.slides.map((slide, index) =>
        index === resolved.slideIndex ? nextSlide : slide,
      ) as T[],
    );
    return true;
  };

  return Object.freeze({
    featureId: EDIT_ANIMATION_SEQUENCE_FEATURE_ID,
    canExecute: (payload?: PresentationAnimationSequenceCommandPayload) =>
      resolve(payload) !== null,
    execute,
  });
}

export interface PresentationAnimationTimingCommandPayload
  extends AnimationTimingValue {
  readonly slideModelId: string;
  readonly index: number;
  readonly expectedStep: TimelineStep;
}

export interface PresentationAnimationTimingCommandAdapter {
  readonly featureId: typeof EDIT_ANIMATION_TIMING_FEATURE_ID;
  readonly canExecute: (
    payload?: PresentationAnimationTimingCommandPayload,
  ) => boolean;
  readonly execute: (payload: PresentationAnimationTimingCommandPayload) => boolean;
}

/** ANI03's live exact-occurrence Start/Delay/Duration authority. */
export function createPresentationAnimationTimingCommandAdapter<
  T extends PresentationAnimationEffectSlide,
>(
  accessors: Omit<
    PresentationAnimationEffectCommandAccessors<T>,
    "getSelectedBlockId"
  >,
): PresentationAnimationTimingCommandAdapter {
  const resolve = (payload?: PresentationAnimationTimingCommandPayload) => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    if (!Number.isInteger(slideIndex) || slideIndex < 0) return null;
    const slide = slides[slideIndex];
    const sequence = slide?.animationSequence;
    if (
      !slide?.id ||
      !hasEditableSlideAnimationSource(slide.source) ||
      !Array.isArray(sequence) ||
      sequence.length === 0 ||
      sequence.length > MAX_TIMELINE_STEPS
    ) {
      return null;
    }
    if (!payload) return { nextSequence: null, slide, slideIndex, slides };
    if (payload.slideModelId !== slide.id) return null;
    const current = sequence[payload.index];
    if (
      !current ||
      !slide.blocks.some((block) => block.id === current.blockId)
    ) {
      return null;
    }
    const nextSequence = editTimelineStepTiming(
      sequence,
      payload.index,
      payload.expectedStep,
      payload,
    );
    return nextSequence === null
      ? null
      : { nextSequence, slide, slideIndex, slides };
  };

  const execute = (payload: PresentationAnimationTimingCommandPayload): boolean => {
    const resolved = resolve(payload);
    if (!resolved?.nextSequence) return false;
    const nextSlide = {
      ...resolved.slide,
      animationSequence: resolved.nextSequence,
    } as T;
    accessors.commitSlides(
      resolved.slides.map((slide, index) =>
        index === resolved.slideIndex ? nextSlide : slide,
      ) as T[],
    );
    return true;
  };

  return Object.freeze({
    featureId: EDIT_ANIMATION_TIMING_FEATURE_ID,
    canExecute: (payload?: PresentationAnimationTimingCommandPayload) =>
      resolve(payload) !== null,
    execute,
  });
}

export type PresentationImageInsertFailureCode =
  | "busy"
  | "no-target"
  | "stale-target"
  | "slide-full"
  | "invalid-asset-id"
  | "identity-collision"
  | "commit-failed"
  | "compensation-failed";

export class PresentationImageInsertError extends Error {
  readonly code: PresentationImageInsertFailureCode;
  readonly cause?: unknown;

  constructor(code: PresentationImageInsertFailureCode, cause?: unknown) {
    super(`presentation_image_insert_${code}`);
    this.name = "PresentationImageInsertError";
    this.code = code;
    this.cause = cause;
  }
}

export interface AssetBackedPresentationImageBlock {
  readonly id: string;
  readonly type: "image";
  readonly sourceKind: "asset";
  readonly assetId: string;
  /** Authorized, bounded renderer/patchback cache; DataItem is authoritative. */
  readonly src: string;
  readonly mimeType: "image/png" | "image/jpeg" | "image/svg+xml";
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  readonly sizeBytes: number;
  readonly hasAlpha: boolean;
  readonly exifOrientation: number;
  readonly officeGraphicKind?: "svg";
  readonly graphicFallbackSrc?: string;
  readonly alt: string;
  readonly caption: string;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

export interface PresentationImageInsertSlide {
  readonly id: string;
  readonly blocks: readonly Record<string, unknown>[];
}

export interface PresentationImageInsertMetadata {
  readonly alt?: string;
  readonly caption?: string;
  /** The size the picture is first drawn at, in CSS pixels, when it is not
   * the file's own pixel size: an SVG rasterized above 1x keeps the SVG's
   * size on the slide (PPT-B353). */
  readonly displaySize?: { readonly width: number; readonly height: number };
}

export interface PresentationImageFileCommandAccessors<
  T extends PresentationImageInsertSlide,
> {
  getSlides: () => readonly T[];
  getActiveIndex: () => number;
  getSlideBounds: (slide: T) => { width: number; height: number };
  openFilePicker: () => boolean;
  prepareFile: (file: File) => Promise<PreparedPresentationImageFile>;
  uploadFile: (file: File) => Promise<string>;
  deleteUploadedAsset: (assetId: string) => Promise<void>;
  /** Must be a single synchronous Presentation host transaction. */
  commitSlides: (slides: T[]) => void;
  selectInsertedImage: (blockId: string, slideIndex: number) => void;
  nextId: () => string;
  maxBlocksPerSlide?: number;
}

export interface PresentationImageFileCommandAdapter<
  T extends PresentationImageInsertSlide,
> {
  readonly featureId: typeof INSERT_IMAGE_FILE_COMMAND_ID;
  readonly action: typeof INSERT_IMAGE_FILE_ACTION;
  readonly handlerRef: typeof INSERT_IMAGE_FILE_HANDLER_REF;
  readonly execute: () => boolean;
  readonly cancelPicker: () => void;
  readonly insertFile: (
    file: File,
    metadata?: PresentationImageInsertMetadata,
  ) => Promise<boolean>;
  /** Reserve the live slide and insert a generated/local File without a picker. */
  readonly insertFileAtCurrent: (
    file: File,
    metadata?: PresentationImageInsertMetadata,
  ) => Promise<boolean>;
  /**
   * Insert an SVG as an SVG GRAPHIC on the live slide (Insert > Icons):
   * the vector held inline beside its PNG compatibility preview, no upload
   * (the upload channel takes no SVG).
   */
  readonly insertOfficeGraphicAtCurrent: (
    file: File,
    metadata?: PresentationImageInsertMetadata,
  ) => Promise<boolean>;
  readonly isBusy: () => boolean;
  readonly _slideType?: T;
}

export type PresentationPictureReplaceFailureCode =
  | "busy"
  | "no-target"
  | "stale-target"
  | "invalid-asset-id"
  | "commit-failed"
  | "compensation-failed";

export class PresentationPictureReplaceError extends Error {
  readonly code: PresentationPictureReplaceFailureCode;
  readonly cause?: unknown;

  constructor(code: PresentationPictureReplaceFailureCode, cause?: unknown) {
    super(`presentation_picture_replace_${code}`);
    this.name = "PresentationPictureReplaceError";
    this.code = code;
    this.cause = cause;
  }
}

export interface ReplaceablePresentationPictureBlock extends Record<
  string,
  unknown
> {
  readonly id: string;
  readonly type: "image";
  readonly src: string;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureReplaceSlide<
  TBlock extends ReplaceablePresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureReplaceCommandAccessors<
  TBlock extends ReplaceablePresentationPictureBlock,
  TSlide extends PresentationPictureReplaceSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  /** Returns a block id only while exactly one editable picture is selected. */
  getSelectedPictureId: () => string | null;
  openFilePicker: () => boolean;
  prepareFile: (file: File) => Promise<PreparedPresentationImageFile>;
  uploadFile: (file: File) => Promise<string>;
  deleteUploadedAsset: (assetId: string) => Promise<void>;
  /** Must be a single synchronous Presentation host transaction. */
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureReplaceCommandAdapter<
  TBlock extends ReplaceablePresentationPictureBlock,
  TSlide extends PresentationPictureReplaceSlide<TBlock>,
> {
  readonly featureId: typeof REPLACE_PICTURE_FILE_FEATURE_ID;
  readonly action: typeof REPLACE_PICTURE_FILE_ACTION;
  readonly handlerRef: typeof REPLACE_PICTURE_FILE_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: () => boolean;
  readonly cancelPicker: () => void;
  readonly replaceFile: (file: File) => Promise<boolean>;
  readonly isBusy: () => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export type PresentationPictureCropMode = "crop" | "fill" | "fit";

export interface CroppablePresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly locked?: boolean;
  readonly left?: number;
  readonly top?: number;
  readonly width?: number;
  readonly height?: number;
  readonly pixelWidth?: number;
  readonly pixelHeight?: number;
  readonly cropRect?: CropRect;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureCropSlide<
  TBlock extends CroppablePresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureCropCommandAccessors<
  TBlock extends CroppablePresentationPictureBlock,
  TSlide extends PresentationPictureCropSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  /** Returns a block id only while exactly one editable picture is selected. */
  getSelectedPictureId: () => string | null;
  /** Must be the Presentation host's single synchronous Undo transaction. */
  commitSlides: (slides: TSlide[]) => void;
  setCropMode: (blockId: string, behavior: "toggle" | "enter") => void;
}

export interface PresentationPictureCropCommandAdapter<
  TBlock extends CroppablePresentationPictureBlock,
  TSlide extends PresentationPictureCropSlide<TBlock>,
> {
  readonly featureIds: Readonly<{
    crop: typeof PICTURE_CROP_FEATURE_ID;
    fill: typeof PICTURE_CROP_FILL_FEATURE_ID;
    fit: typeof PICTURE_CROP_FIT_FEATURE_ID;
  }>;
  readonly canExecute: (mode: PresentationPictureCropMode) => boolean;
  readonly execute: (mode: PresentationPictureCropMode) => boolean;
  readonly canApplyAspectRatio: () => boolean;
  readonly commitAspectRatio: (
    blockId: string,
    ratioId: PictureCropAspectRatioId,
  ) => boolean;
  readonly commitFreeCrop: (blockId: string, crop: CropRect | null) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface AltTextPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly alt?: string;
  readonly altTitle?: string;
  readonly decorative?: boolean;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureAltTextSlide<
  TBlock extends AltTextPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureAltTextCommandPayload {
  readonly blockId: string;
  readonly description: string;
  readonly decorative: boolean;
}

export interface PresentationPictureAltTextCommandAccessors<
  TBlock extends AltTextPresentationPictureBlock,
  TSlide extends PresentationPictureAltTextSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  /** Must be the Presentation host's single synchronous Undo transaction. */
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureAltTextCommandAdapter<
  TBlock extends AltTextPresentationPictureBlock,
  TSlide extends PresentationPictureAltTextSlide<TBlock>,
> {
  readonly featureId: typeof PICTURE_ALT_TEXT_FEATURE_ID;
  readonly action: typeof PICTURE_ALT_TEXT_ACTION;
  readonly handlerRef: typeof PICTURE_ALT_TEXT_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationPictureAltTextCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface BorderedPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly borderColor?: string;
  readonly borderWidth?: number;
  readonly borderStyle?: ImageBorderStyle;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureBorderSlide<
  TBlock extends BorderedPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureBorderCommandPayload {
  readonly blockId: string;
  readonly color?: string;
  readonly width?: number;
  readonly style: ImageBorderStyle;
}

export interface PresentationPictureBorderCommandAccessors<
  TBlock extends BorderedPresentationPictureBlock,
  TSlide extends PresentationPictureBorderSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureBorderCommandAdapter<
  TBlock extends BorderedPresentationPictureBlock,
  TSlide extends PresentationPictureBorderSlide<TBlock>,
> {
  readonly featureId: typeof PICTURE_BORDER_FEATURE_ID;
  readonly action: typeof PICTURE_BORDER_ACTION;
  readonly handlerRef: typeof PICTURE_BORDER_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationPictureBorderCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface TransparentPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly opacity?: number;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureTransparencySlide<
  TBlock extends TransparentPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureTransparencyCommandPayload {
  readonly blockId: string;
  /** Persisted visible opacity, not the UI's inverse Transparency value. */
  readonly opacity: number;
}

export interface PresentationPictureTransparencyCommandAccessors<
  TBlock extends TransparentPresentationPictureBlock,
  TSlide extends PresentationPictureTransparencySlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureTransparencyCommandAdapter<
  TBlock extends TransparentPresentationPictureBlock,
  TSlide extends PresentationPictureTransparencySlide<TBlock>,
> {
  readonly featureId: typeof PICTURE_TRANSPARENCY_FEATURE_ID;
  readonly action: typeof PICTURE_TRANSPARENCY_ACTION;
  readonly handlerRef: typeof PICTURE_TRANSPARENCY_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationPictureTransparencyCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface CorrectedPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly locked?: boolean;
  readonly brightness?: number;
  readonly contrast?: number;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureCorrectionsSlide<
  TBlock extends CorrectedPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureCorrectionsCommandPayload {
  readonly blockId: string;
  /** Whole signed native deltas; persisted storage remains delta + 100. */
  readonly brightnessDelta: number;
  readonly contrastDelta: number;
}

export interface PresentationPictureCorrectionsCommandAccessors<
  TBlock extends CorrectedPresentationPictureBlock,
  TSlide extends PresentationPictureCorrectionsSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureCorrectionsCommandAdapter<
  TBlock extends CorrectedPresentationPictureBlock,
  TSlide extends PresentationPictureCorrectionsSlide<TBlock>,
> {
  readonly featureId: typeof PICTURE_CORRECTIONS_FEATURE_ID;
  readonly action: typeof PICTURE_CORRECTIONS_ACTION;
  readonly handlerRef: typeof PICTURE_CORRECTIONS_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationPictureCorrectionsCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface ColoredPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly locked?: boolean;
  readonly recolor?: unknown;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureColorSlide<
  TBlock extends ColoredPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export interface PresentationPictureColorCommandPayload {
  readonly blockId: string;
  readonly recolor: NativeImageRecolorMode;
}

export interface PresentationPictureColorCommandAccessors<
  TBlock extends ColoredPresentationPictureBlock,
  TSlide extends PresentationPictureColorSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureColorCommandAdapter<
  TBlock extends ColoredPresentationPictureBlock,
  TSlide extends PresentationPictureColorSlide<TBlock>,
> {
  readonly featureId: typeof PICTURE_COLOR_FEATURE_ID;
  readonly action: typeof PICTURE_COLOR_ACTION;
  readonly handlerRef: typeof PICTURE_COLOR_HANDLER_REF;
  readonly canExecute: () => boolean;
  readonly execute: (
    payload: PresentationPictureColorCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export type PresentationPictureSizeAxis = "width" | "height";
export type PresentationPictureSizeStepDirection = "increment" | "decrement";

export interface SizedPresentationPictureBlock extends ReplaceablePresentationPictureBlock {
  readonly locked?: boolean;
  readonly left?: number;
  readonly top?: number;
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly width?: number;
  readonly height?: number;
  readonly source?: Record<string, unknown>;
}

export interface PresentationPictureSizeSlide<
  TBlock extends SizedPresentationPictureBlock,
> {
  readonly id: string;
  readonly blocks: readonly (TBlock | Record<string, unknown>)[];
}

export type PresentationPictureSizeCommandPayload = Readonly<{
  blockId: string;
  axis: PresentationPictureSizeAxis;
}> &
  (
    | Readonly<{ kind: "absolute"; targetEmu: number }>
    | Readonly<{
        kind: "step";
        direction: PresentationPictureSizeStepDirection;
        /** The field unit's step in EMU (0.01" or 0.1 cm); 0.01" when absent. */
        stepEmu?: number;
      }>
  );

export interface PresentationPictureSizeCommandAccessors<
  TBlock extends SizedPresentationPictureBlock,
  TSlide extends PresentationPictureSizeSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  commitSlides: (slides: TSlide[]) => void;
  canEdit?: () => boolean;
  getGeneration?: () => unknown;
}

export interface PresentationPictureResizeGesture {
  preview: (
    handle: ResizeHandle,
    dx: number,
    dy: number,
    modifiers?: ResizeModifiers,
  ) => BlockGeometry | null;
  /** The picture's mirror flags the latest preview shows: a handle dragged
   * through the pinned edge mirrors it (PowerPoint), dragging back restores. */
  latestFlip: () => { flipH: boolean; flipV: boolean };
  commit: () => boolean;
  cancel: () => void;
}

export interface PresentationPictureSizeCommandAdapter<
  TBlock extends SizedPresentationPictureBlock,
  TSlide extends PresentationPictureSizeSlide<TBlock>,
> {
  readonly featureIds: Readonly<{
    width: typeof PICTURE_SIZE_WIDTH_FEATURE_ID;
    height: typeof PICTURE_SIZE_HEIGHT_FEATURE_ID;
  }>;
  readonly canExecute: (axis: PresentationPictureSizeAxis) => boolean;
  readonly execute: (payload: PresentationPictureSizeCommandPayload) => boolean;
  readonly beginResize: (
    blockId: string,
    bounds: DragBounds,
  ) => PresentationPictureResizeGesture | null;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}

export interface ArrangeablePresentationBlock {
  readonly id: string;
  readonly type: string;
  readonly locked?: boolean;
  readonly name?: string;
  readonly hidden?: boolean;
  readonly source?: Record<string, unknown>;
  readonly sourceOrder?: number;
  readonly zIndex?: number;
}

export interface PresentationPictureArrangeSlide<
  TBlock extends ArrangeablePresentationBlock,
> {
  readonly id: string;
  readonly source?: Record<string, unknown>;
  readonly blocks: readonly TBlock[];
  readonly nativeShapeOrderTarget?: Readonly<{
    schemaVersion: "presentation-native-shape-order-v2";
    slideId: string;
    steps: readonly Readonly<{
      shapeId: string;
      direction: ZOrderStep;
    }>[];
  }>;
}

/** The steps a slide's native order journal takes one by one, as recorded. */
export const PRESENTATION_PICTURE_ARRANGE_MAX_STEPS = 32;











export function presentationNativeOrderStepBound(objectCount: number): number {
  const count = Math.max(0, Math.floor(objectCount));
  return Math.max(PRESENTATION_PICTURE_ARRANGE_MAX_STEPS, (count * (count - 1)) / 2);
}

export interface PresentationPictureArrangeCommandPayload {
  readonly blockId: string;
  readonly direction: ZOrderStep;
}

export interface PresentationPictureArrangeCommandAccessors<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSelectedPictureId: () => string | null;
  /** Must be the Presentation host's single synchronous Undo transaction. */
  commitSlides: (slides: TSlide[]) => void;
}

export interface PresentationPictureArrangeCommandAdapter<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
> {
  readonly featureIds: Readonly<{
    forward: typeof PICTURE_BRING_FORWARD_FEATURE_ID;
    backward: typeof PICTURE_SEND_BACKWARD_FEATURE_ID;
  }>;
  readonly canExecute: (direction: ZOrderStep) => boolean;
  readonly execute: (
    payload: PresentationPictureArrangeCommandPayload,
  ) => boolean;
  readonly _slideType?: TSlide;
  readonly _blockType?: TBlock;
}








function targetSlideIndex<T extends PresentationImageInsertSlide>(
  slides: readonly T[],
  _activeIndex: number,
  targetSlideId: string,
): number {
  return slides.findIndex((slide) => slide.id === targetSlideId);
}

/**
 * The picture block Insert > Pictures adds for an uploaded workspace image:
 * asset-backed, with the prepared bytes as its bounded render cache, its
 * description and caption, at its natural 96-DPI size (scaled down only to
 * fit the slide) and centred.
 */
export function assetBackedPresentationImageBlock(
  blockId: string,
  assetId: string,
  prepared: PreparedPresentationImageFile,
  metadata: PresentationImageInsertMetadata,
  bounds: { width: number; height: number },
): AssetBackedPresentationImageBlock {
  return Object.freeze({
    id: blockId,
    type: "image",
    sourceKind: "asset",
    assetId,
    src: prepared.src,
    mimeType: prepared.mimeType,
    pixelWidth: prepared.pixelWidth,
    pixelHeight: prepared.pixelHeight,
    sizeBytes: prepared.sizeBytes,
    hasAlpha: prepared.hasAlpha,
    exifOrientation: prepared.exifOrientation,
    ...(prepared.officeGraphicKind
      ? { officeGraphicKind: prepared.officeGraphicKind }
      : {}),
    ...(prepared.graphicFallbackSrc
      ? { graphicFallbackSrc: prepared.graphicFallbackSrc }
      : {}),
    alt:
      typeof metadata.alt === "string"
        ? metadata.alt.trim().slice(0, 1_000)
        : "",
    caption:
      typeof metadata.caption === "string"
        ? metadata.caption.trim().slice(0, 1_000)
        : "",
    ...defaultImageGeometry(bounds, prepared, metadata.displaySize),
  });
}

/**
 * The SVG graphic Insert > Icons adds: PowerPoint inserts an icon as an SVG
 * graphic (Graphics Format, Graphics Fill / Outline), not a picture
 * (Microsoft Support, "Insert icons in Microsoft 365"). The upload channel
 * takes no SVG, so the graphic keeps its vector inline (`src`, the passive
 * Office Graphic profile the save validator and both writers accept) with
 * its PNG compatibility preview, and states only its intrinsic size.
 */
export function inlinePresentationOfficeGraphicBlock(
  blockId: string,
  prepared: PreparedPresentationImageFile,
  metadata: PresentationImageInsertMetadata,
  bounds: { width: number; height: number },
): Record<string, unknown> {
  if (prepared.officeGraphicKind !== "svg" || !prepared.graphicFallbackSrc) {
    throw new PresentationImageInsertError("commit-failed");
  }
  return Object.freeze({
    id: blockId,
    type: "image",
    src: prepared.src,
    officeGraphicKind: "svg",
    graphicFallbackSrc: prepared.graphicFallbackSrc,
    pixelWidth: prepared.pixelWidth,
    pixelHeight: prepared.pixelHeight,
    alt: typeof metadata.alt === "string" ? metadata.alt.trim().slice(0, 1_000) : "",
    caption: typeof metadata.caption === "string" ? metadata.caption.trim().slice(0, 1_000) : "",
    ...defaultImageGeometry(bounds, prepared, metadata.displaySize),
  });
}

function defaultImageGeometry(
  bounds: { width: number; height: number },
  prepared: PreparedPresentationImageFile,
  displaySize?: { readonly width: number; readonly height: number },
): Pick<
  AssetBackedPresentationImageBlock,
  "left" | "top" | "width" | "height"
> {
  // Use the editor's established 96-DPI pixel↔EMU baseline. Preserve a small
  // image's natural size; only scale down when it would exceed the slide.
  const stated =
    displaySize &&
    Number.isFinite(displaySize.width) && displaySize.width > 0 &&
    Number.isFinite(displaySize.height) && displaySize.height > 0
      ? displaySize
      : null;
  const naturalWidth = (stated?.width ?? prepared.pixelWidth) * EMU_PER_96_DPI_PIXEL;
  const naturalHeight = (stated?.height ?? prepared.pixelHeight) * EMU_PER_96_DPI_PIXEL;
  const scale = Math.min(
    1,
    Math.max(1, bounds.width) / naturalWidth,
    Math.max(1, bounds.height) / naturalHeight,
  );
  const width = Math.max(1, Math.round(naturalWidth * scale));
  const height = Math.max(1, Math.round(naturalHeight * scale));
  return {
    left: Math.round((bounds.width - width) / 2),
    top: Math.round((bounds.height - height) / 2),
    width,
    height,
  };
}

async function compensateOrThrow<TError extends Error>(
  deleteUploadedAsset: (assetId: string) => Promise<void>,
  assetId: string,
  original: TError,
  compensationFailure: (cause: {
    original: TError;
    compensation: unknown;
  }) => Error = (cause) =>
    new PresentationImageInsertError("compensation-failed", cause),
): Promise<never> {
  try {
    await deleteUploadedAsset(assetId);
  } catch (error) {
    throw compensationFailure({
      original,
      compensation: error,
    });
  }
  throw original;
}

function selectedPictureTarget<
  TBlock extends ReplaceablePresentationPictureBlock,
  TSlide extends PresentationPictureReplaceSlide<TBlock>,
>(
  slides: readonly TSlide[],
  activeIndex: number,
  selectedPictureId: string | null,
): { slideIndex: number; slide: TSlide; block: TBlock } | null {
  if (!selectedPictureId || !Number.isInteger(activeIndex) || activeIndex < 0) {
    return null;
  }
  const slide = slides[activeIndex];
  if (!slide) return null;
  const block = slide.blocks.find(
    (candidate): candidate is TBlock =>
      candidate.id === selectedPictureId && candidate.type === "image",
  );
  return block ? { slideIndex: activeIndex, slide, block } : null;
}

function positiveFinite(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value > 0;
}

/**
 * Presentation-owned Alt Text mutation. The dialog is editor-only state; Apply
 * resolves the live, single selected picture again and commits one cloned slide.
 * The legacy cNvPr title remains untouched. Missing decorative metadata also
 * remains missing unless the user actually changes a true value back to false.
 */
export function createPresentationPictureAltTextCommandAdapter<
  TBlock extends AltTextPresentationPictureBlock,
  TSlide extends PresentationPictureAltTextSlide<TBlock>,
>(
  accessors: PresentationPictureAltTextCommandAccessors<TBlock, TSlide>,
): PresentationPictureAltTextCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved ? { ...resolved, slides } : null;
  };

  const canExecute = (): boolean => {
    const resolved = target();
    return (
      resolved !== null &&
      resolved.block.source?.pictureAltTextDisposition !== "preserve-only"
    );
  };

  const execute = (
    payload: PresentationPictureAltTextCommandPayload,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId ||
      resolved.block.source?.pictureAltTextDisposition === "preserve-only"
    ) {
      return false;
    }
    const currentDescription =
      typeof resolved.block.alt === "string" ? resolved.block.alt : "";
    const currentDecorative = resolved.block.decorative === true;
    if (
      payload.description === currentDescription &&
      payload.decorative === currentDecorative
    ) {
      return true;
    }

    const description = payload.decorative
      ? ""
      : normalizeImageAltText(payload.description);
    const decorative: boolean | undefined = payload.decorative
      ? true
      : currentDecorative || typeof resolved.block.decorative === "boolean"
        ? false
        : undefined;
    if (
      description === currentDescription &&
      decorative === resolved.block.decorative
    ) {
      return true;
    }

    const replacement = {
      ...resolved.block,
      alt: description,
      ...(decorative === undefined ? {} : { decorative }),
    } as TBlock;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? replacement : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: PICTURE_ALT_TEXT_FEATURE_ID,
    action: PICTURE_ALT_TEXT_ACTION,
    handlerRef: PICTURE_ALT_TEXT_HANDLER_REF,
    canExecute,
    execute,
  });
}

/** One live-target, one-transaction Picture Border mutation. */
export function createPresentationPictureBorderCommandAdapter<
  TBlock extends BorderedPresentationPictureBlock,
  TSlide extends PresentationPictureBorderSlide<TBlock>,
>(
  accessors: PresentationPictureBorderCommandAccessors<TBlock, TSlide>,
): PresentationPictureBorderCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved ? { ...resolved, slides } : null;
  };
  const canExecute = () => {
    const resolved = target();
    return (
      resolved !== null &&
      resolved.block.source?.pictureBorderDisposition !== "preserve-only"
    );
  };
  const execute = (
    payload: PresentationPictureBorderCommandPayload,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId ||
      resolved.block.source?.pictureBorderDisposition === "preserve-only"
    ) {
      return false;
    }
    if (
      normalizeImageBorderStyle(payload.style, "none") !== payload.style ||
      (payload.width !== undefined &&
        (typeof payload.width !== "number" ||
          !Number.isFinite(payload.width) ||
          payload.width <= 0 ||
          payload.width > MAX_IMAGE_BORDER_WIDTH)) ||
      (payload.style === "none" && payload.color !== undefined)
    ) {
      return false;
    }
    const color = normalizeImageBorderColor(payload.color, null);
    const width = normalizeImageBorderWidth(payload.width);
    if (payload.style !== "none" && color === null) return false;
    if (payload.width !== undefined && width <= 0) return false;

    const nextBorder: {
      borderColor?: string;
      borderWidth?: number;
      borderStyle: ImageBorderStyle;
    } =
      payload.style === "none"
        ? {
            borderStyle: "none" as const,
            ...(payload.width === undefined ? {} : { borderWidth: width }),
          }
        : {
            borderColor: color as string,
            borderStyle: payload.style,
            ...(payload.width === undefined ? {} : { borderWidth: width }),
          };
    if (
      resolved.block.borderColor === nextBorder.borderColor &&
      resolved.block.borderWidth === nextBorder.borderWidth &&
      resolved.block.borderStyle === nextBorder.borderStyle
    ) {
      return true;
    }
    const unchanged = { ...resolved.block } as Record<string, unknown>;
    delete unchanged.borderColor;
    delete unchanged.borderWidth;
    delete unchanged.borderStyle;
    const replacement = { ...unchanged, ...nextBorder } as TBlock;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? replacement : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };
  return Object.freeze({
    featureId: PICTURE_BORDER_FEATURE_ID,
    action: PICTURE_BORDER_ACTION,
    handlerRef: PICTURE_BORDER_HANDLER_REF,
    canExecute,
    execute,
  });
}

/** One live-target, one-transaction Picture Transparency mutation. */
export function createPresentationPictureTransparencyCommandAdapter<
  TBlock extends TransparentPresentationPictureBlock,
  TSlide extends PresentationPictureTransparencySlide<TBlock>,
>(
  accessors: PresentationPictureTransparencyCommandAccessors<TBlock, TSlide>,
): PresentationPictureTransparencyCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved ? { ...resolved, slides } : null;
  };
  const canExecute = (): boolean => {
    const resolved = target();
    return (
      resolved !== null &&
      resolved.block.source?.pictureTransparencyDisposition !== "preserve-only"
    );
  };
  const execute = (
    payload: PresentationPictureTransparencyCommandPayload,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId ||
      resolved.block.source?.pictureTransparencyDisposition ===
        "preserve-only" ||
      normalizeImageOpacity(payload.opacity) !== payload.opacity
    ) {
      return false;
    }
    const currentOpacity = normalizeImageOpacity(resolved.block.opacity);
    if (currentOpacity === payload.opacity) return true;

    const replacement = { ...resolved.block } as Record<string, unknown>;
    if (payload.opacity === 100) delete replacement.opacity;
    else replacement.opacity = payload.opacity;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? (replacement as TBlock) : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: PICTURE_TRANSPARENCY_FEATURE_ID,
    action: PICTURE_TRANSPARENCY_ACTION,
    handlerRef: PICTURE_TRANSPARENCY_HANDLER_REF,
    canExecute,
    execute,
  });
}

/** One live-target, one-transaction Picture Corrections mutation. */
export function createPresentationPictureCorrectionsCommandAdapter<
  TBlock extends CorrectedPresentationPictureBlock,
  TSlide extends PresentationPictureCorrectionsSlide<TBlock>,
>(
  accessors: PresentationPictureCorrectionsCommandAccessors<TBlock, TSlide>,
): PresentationPictureCorrectionsCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved ? { ...resolved, slides } : null;
  };
  const canExecute = (): boolean => {
    const resolved = target();
    return (
      resolved !== null &&
      hasEditablePictureCorrections(
        resolved.block.brightness,
        resolved.block.contrast,
        resolved.block.source,
      )
    );
  };
  const execute = (
    payload: PresentationPictureCorrectionsCommandPayload,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId ||
      !hasEditablePictureCorrections(
        resolved.block.brightness,
        resolved.block.contrast,
        resolved.block.source,
      ) ||
      !isImageCorrectionDelta(payload.brightnessDelta) ||
      !isImageCorrectionDelta(payload.contrastDelta)
    ) {
      return false;
    }
    const currentCorrections = imageCorrectionDeltasFromPersisted(
      resolved.block.brightness,
      resolved.block.contrast,
      resolved.block.source,
    );
    if (
      currentCorrections.brightnessDelta === payload.brightnessDelta &&
      currentCorrections.contrastDelta === payload.contrastDelta
    ) {
      return true;
    }

    const replacement = { ...resolved.block } as Record<string, unknown>;
    const canonicalSource = pictureCorrectionsSourceWithStoredUnit(
      resolved.block.source,
    );
    if (canonicalSource !== resolved.block.source) {
      replacement.source = canonicalSource;
    }
    if (payload.brightnessDelta === 0) delete replacement.brightness;
    else {
      replacement.brightness = imageCorrectionStoredFromDelta(
        payload.brightnessDelta,
      );
    }
    if (payload.contrastDelta === 0) delete replacement.contrast;
    else {
      replacement.contrast = imageCorrectionStoredFromDelta(
        payload.contrastDelta,
      );
    }
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? (replacement as TBlock) : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: PICTURE_CORRECTIONS_FEATURE_ID,
    action: PICTURE_CORRECTIONS_ACTION,
    handlerRef: PICTURE_CORRECTIONS_HANDLER_REF,
    canExecute,
    execute,
  });
}

/** One live-target, one-transaction Picture Color mutation. */
export function createPresentationPictureColorCommandAdapter<
  TBlock extends ColoredPresentationPictureBlock,
  TSlide extends PresentationPictureColorSlide<TBlock>,
>(
  accessors: PresentationPictureColorCommandAccessors<TBlock, TSlide>,
): PresentationPictureColorCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved ? { ...resolved, slides } : null;
  };
  const canExecute = (): boolean => {
    const resolved = target();
    return (
      resolved !== null &&
      hasEditablePictureColor(resolved.block.source) &&
      (resolved.block.recolor === undefined ||
        isNativeImageRecolorMode(resolved.block.recolor))
    );
  };
  const execute = (
    payload: PresentationPictureColorCommandPayload,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId ||
      !hasEditablePictureColor(resolved.block.source) ||
      (resolved.block.recolor !== undefined &&
        !isNativeImageRecolorMode(resolved.block.recolor)) ||
      !isNativeImageRecolorMode(payload.recolor)
    ) {
      return false;
    }
    const current = resolved.block.recolor ?? "none";
    if (current === payload.recolor) return true;

    const replacement = { ...resolved.block } as Record<string, unknown>;
    if (payload.recolor === "none") delete replacement.recolor;
    else replacement.recolor = payload.recolor;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? (replacement as TBlock) : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureId: PICTURE_COLOR_FEATURE_ID,
    action: PICTURE_COLOR_ACTION,
    handlerRef: PICTURE_COLOR_HANDLER_REF,
    canExecute,
    execute,
  });
}

interface PresentationPictureSizeReceipt {
  readonly width: number;
  readonly height: number;
  readonly aspectLocked: boolean;
  readonly ratioWidth: number;
  readonly ratioHeight: number;
}

function exactPictureSizeEmu(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isSafeInteger(value) &&
    value >= PRESENTATION_PICTURE_SIZE_MIN_EMU &&
    value <= PRESENTATION_PICTURE_SIZE_MAX_EMU
  );
}

function roundPositiveRatio(numerator: number, denominator: number): number {
  return Math.floor((numerator + Math.floor(denominator / 2)) / denominator);
}

/**
 * Crop > Aspect Ratio reshapes the picture FRAME; from then on that frame is
 * the picture's size baseline (the PPTX writer measures a later locked resize
 * from the frame it wrote), until Reset Picture & Size restores the original.
 * The baseline rides in the block's `source` next to the native size receipt;
 * a reopen re-projects the receipt from the file, where the crop frame IS the
 * current size, so the two agree.
 */
const PICTURE_SIZE_CROP_FRAME_WIDTH_KEY = "pictureSizeCropFrameWidthEmu";
const PICTURE_SIZE_CROP_FRAME_HEIGHT_KEY = "pictureSizeCropFrameHeightEmu";

function presentationPictureCropFrameBaseline(
  source: Record<string, unknown> | undefined,
): Readonly<{ width: number; height: number }> | null {
  const width = source?.[PICTURE_SIZE_CROP_FRAME_WIDTH_KEY];
  const height = source?.[PICTURE_SIZE_CROP_FRAME_HEIGHT_KEY];
  return exactPictureSizeEmu(width) && exactPictureSizeEmu(height)
    ? { width, height }
    : null;
}

function withoutPictureCropFrameBaseline(
  source: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  if (
    !source ||
    (!(PICTURE_SIZE_CROP_FRAME_WIDTH_KEY in source) &&
      !(PICTURE_SIZE_CROP_FRAME_HEIGHT_KEY in source))
  ) {
    return source;
  }
  const next = { ...source };
  delete next[PICTURE_SIZE_CROP_FRAME_WIDTH_KEY];
  delete next[PICTURE_SIZE_CROP_FRAME_HEIGHT_KEY];
  return next;
}

function presentationPictureSizeReceipt(
  block: SizedPresentationPictureBlock,
): PresentationPictureSizeReceipt | null {
  if (
    block.type !== "image" ||
    block.locked === true ||
    !exactPictureSizeEmu(block.width) ||
    !exactPictureSizeEmu(block.height)
  ) {
    return null;
  }
  const source = block.source;
  if (source?.pictureSizeDisposition === "preserve-only") return null;
  // A replaced picture (Change Picture) draws a fresh raster in the frame
  // it found. Its locked ratio is whichever baseline the frame is on: the
  // crop frame Crop > Aspect Ratio shaped, the native original the frame
  // kept through the replacement, or the new raster's own pixels, which is
  // what Reset Picture & Size restores. Measured against the native original
  // alone, the reset frame read as off-ratio and the size boxes went dark
  // (PPT-B349).
  if (
    block.sourceKind === "asset" &&
    Number.isSafeInteger(block.pixelWidth) &&
    Number.isSafeInteger(block.pixelHeight) &&
    (block.pixelWidth as number) >= 1 &&
    (block.pixelHeight as number) >= 1
  ) {
    const aspectLocked =
      typeof source?.pictureSizeAspectLocked === "boolean"
        ? source.pictureSizeAspectLocked
        : true;
    const cropFrame = presentationPictureCropFrameBaseline(source);
    const baselines: Array<{ width: number; height: number }> = [
      ...(cropFrame ? [cropFrame] : []),
      ...(exactPictureSizeEmu(source?.pictureSizeOriginalWidthEmu) &&
      exactPictureSizeEmu(source?.pictureSizeOriginalHeightEmu)
        ? [{ width: source!.pictureSizeOriginalWidthEmu as number, height: source!.pictureSizeOriginalHeightEmu as number }]
        : []),
      { width: block.pixelWidth as number, height: block.pixelHeight as number },
    ];
    const width = block.width as number;
    const height = block.height as number;
    // Within one EMU of the ratio: a frame fitted by integer arithmetic
    // (the AI composer's floor division, a rounded import) lands one EMU
    // off the exact value, and read as off-ratio it lost its resize
    
    const onRatio = (ratio: { width: number; height: number }) =>
      Math.abs(roundPositiveRatio(width * ratio.height, ratio.width) - height) <= 1 ||
      Math.abs(roundPositiveRatio(height * ratio.width, ratio.height) - width) <= 1;
    // A frame on none of them keeps its handles and size fields, its own
    // proportions locked, as PowerPoint keeps a picture's handles whatever
    // its frame: an AI cover picture cropped to fill a 16:9 slide (`cropRect`)
    // is on no raster ratio, and lost every grip and the Size boxes.
    const baseline = aspectLocked
      ? baselines.find(onRatio) ?? { width, height }
      : baselines[baselines.length - 1];
    return {
      width,
      height,
      aspectLocked,
      ratioWidth: baseline.width,
      ratioHeight: baseline.height,
    };
  }
  const hasSlideId = source !== undefined && "slideId" in source;
  const hasShapeId = source !== undefined && "shapeId" in source;
  if (hasSlideId || hasShapeId) {
    if (
      typeof source?.slideId !== "string" ||
      !/^sld[1-9][0-9]*$/.test(source.slideId) ||
      typeof source.shapeId !== "string" ||
      !/^shp[1-9][0-9]*$/.test(source.shapeId) ||
      source.pictureSizeSchemaVersion !==
        PRESENTATION_PICTURE_SIZE_SCHEMA_VERSION ||
      typeof source.pictureSizeAspectLocked !== "boolean" ||
      !exactPictureSizeEmu(source.pictureSizeOriginalWidthEmu) ||
      !exactPictureSizeEmu(source.pictureSizeOriginalHeightEmu)
    ) {
      return null;
    }
    // The locked ratio is measured from the crop frame once Crop > Aspect
    // Ratio has reshaped it, else from the native original.
    const cropFrame = presentationPictureCropFrameBaseline(source);
    const ratioWidth = cropFrame?.width ?? source.pictureSizeOriginalWidthEmu;
    const ratioHeight =
      cropFrame?.height ?? source.pictureSizeOriginalHeightEmu;
    // Within one EMU of the ratio, as for a replaced picture above: an
    // imported frame rounded by its reader is still on its ratio. A frame the
    // file stretched off it keeps its handles too, locked to its own
    // proportions (PowerPoint never drops a picture's grips for its ratio).
    const offRatio =
      source.pictureSizeAspectLocked &&
      Math.abs(roundPositiveRatio(block.width * ratioHeight, ratioWidth) - block.height) > 1 &&
      Math.abs(roundPositiveRatio(block.height * ratioWidth, ratioHeight) - block.width) > 1;
    return {
      width: block.width,
      height: block.height,
      aspectLocked: source.pictureSizeAspectLocked,
      ratioWidth: offRatio ? block.width : ratioWidth,
      ratioHeight: offRatio ? block.height : ratioHeight,
    };
  }
  // Current from-scratch picture serialization emits a:picLocks with
  // noChangeAspect=1. Source-less pictures therefore have one deterministic
  // native lock default; source-backed legacy pages never use this branch.
  return {
    width: block.width,
    height: block.height,
    aspectLocked: true,
    ratioWidth: block.width,
    ratioHeight: block.height,
  };
}

export function hasEditablePresentationPictureSize(
  block: SizedPresentationPictureBlock | null | undefined,
): boolean {
  return block ? presentationPictureSizeReceipt(block) !== null : false;
}

export interface PresentationPictureResetAndSizeBounds {
  readonly width: number;
  readonly height: number;
}

function assetPictureResetExtents(
  block: SizedPresentationPictureBlock,
  bounds: PresentationPictureResetAndSizeBounds,
): Readonly<{ width: number; height: number }> | null {
  if (
    block.sourceKind !== "asset" ||
    typeof block.assetId !== "string" ||
    block.assetId.trim().length === 0 ||
    !Number.isSafeInteger(block.pixelWidth) ||
    !Number.isSafeInteger(block.pixelHeight) ||
    (block.pixelWidth as number) < 1 ||
    (block.pixelHeight as number) < 1 ||
    !exactPictureSizeEmu(bounds.width) ||
    !exactPictureSizeEmu(bounds.height)
  ) {
    return null;
  }
  const naturalWidth = (block.pixelWidth as number) * EMU_PER_96_DPI_PIXEL;
  const naturalHeight = (block.pixelHeight as number) * EMU_PER_96_DPI_PIXEL;
  if (
    !Number.isSafeInteger(naturalWidth) ||
    !Number.isSafeInteger(naturalHeight)
  ) {
    return null;
  }
  // Match insertion: retain the 96-DPI natural size unless it exceeds the
  // slide, then scale both dimensions down together. Resetting size never
  // recentres the picture; left/top remain author-owned.
  const scale = Math.min(
    1,
    bounds.width / naturalWidth,
    bounds.height / naturalHeight,
  );
  const width = Math.max(1, Math.round(naturalWidth * scale));
  const height = Math.max(1, Math.round(naturalHeight * scale));
  return exactPictureSizeEmu(width) && exactPictureSizeEmu(height)
    ? { width, height }
    : null;
}

function presentationPictureResetExtents(
  block: SizedPresentationPictureBlock,
  bounds: PresentationPictureResetAndSizeBounds,
): Readonly<{ width: number; height: number }> | null {
  if (
    block.type !== "image" ||
    block.locked === true ||
    !exactPictureSizeEmu(block.width) ||
    !exactPictureSizeEmu(block.height)
  ) {
    return null;
  }
  // A replacement/local upload owns a fresh, authenticated raster baseline.
  // Prefer it over any native source receipt retained for round-trip custody.
  if ("sourceKind" in block) {
    return assetPictureResetExtents(block, bounds);
  }
  const source = block.source;
  if (
    source?.pictureSizeDisposition === "preserve-only" ||
    typeof source?.slideId !== "string" ||
    !/^sld[1-9][0-9]*$/.test(source.slideId) ||
    typeof source.shapeId !== "string" ||
    !/^shp[1-9][0-9]*$/.test(source.shapeId) ||
    source.pictureSizeSchemaVersion !==
      PRESENTATION_PICTURE_SIZE_SCHEMA_VERSION ||
    typeof source.pictureSizeAspectLocked !== "boolean" ||
    !exactPictureSizeEmu(source.pictureSizeOriginalWidthEmu) ||
    !exactPictureSizeEmu(source.pictureSizeOriginalHeightEmu)
  ) {
    return null;
  }
  return {
    width: source.pictureSizeOriginalWidthEmu,
    height: source.pictureSizeOriginalHeightEmu,
  };
}

/** True only when Reset Picture & Size has a trustworthy original-size receipt. */
export function hasResettablePresentationPictureAndSize(
  block: SizedPresentationPictureBlock | null | undefined,
  bounds: PresentationPictureResetAndSizeBounds,
): boolean {
  if (!block) return false;
  const extents = presentationPictureResetExtents(block, bounds);
  if (!extents) return false;
  return (
    block.width !== extents.width ||
    block.height !== extents.height ||
    imageHasResettableFormatting(
      block as Parameters<typeof imageHasResettableFormatting>[0],
    )
  );
}

/**
 * PowerPoint's distinct Reset Picture & Size leaf: clear picture formatting
 * and crop, restore an authenticated original size, and preserve every other
 * field. Missing/malformed provenance and no-op requests fail closed.
 */
export function resetPresentationPictureAndSize<
  TBlock extends SizedPresentationPictureBlock,
>(
  block: TBlock,
  bounds: PresentationPictureResetAndSizeBounds,
): TBlock | null {
  const extents = presentationPictureResetExtents(block, bounds);
  if (
    !extents ||
    (block.width === extents.width &&
      block.height === extents.height &&
      !imageHasResettableFormatting(
        block as Parameters<typeof imageHasResettableFormatting>[0],
      ))
  ) {
    return null;
  }
  const reset = resetImagePicture(
    block as Parameters<typeof resetImagePicture>[0],
  ) as TBlock;
  // Back at the original size the crop frame is no longer the baseline.
  const source = withoutPictureCropFrameBaseline(reset.source);
  return {
    ...reset,
    ...(source !== reset.source ? { source } : {}),
    width: extents.width,
    height: extents.height,
  };
}

function resizedPictureExtents(
  current: PresentationPictureSizeReceipt,
  axis: PresentationPictureSizeAxis,
  target: number,
): Readonly<{ width: number; height: number }> | null {
  if (!exactPictureSizeEmu(target)) return null;
  if (!current.aspectLocked) {
    return axis === "width"
      ? { width: target, height: current.height }
      : { width: current.width, height: target };
  }
  const width =
    axis === "width"
      ? target
      : roundPositiveRatio(target * current.ratioWidth, current.ratioHeight);
  const height =
    axis === "height"
      ? target
      : roundPositiveRatio(target * current.ratioHeight, current.ratioWidth);
  return exactPictureSizeEmu(width) && exactPictureSizeEmu(height)
    ? { width, height }
    : null;
}

/**
 * A picture's extents once one side is set, by the picture's own size rule
 * (its aspect lock and the ratio it keeps), or null where the picture has no
 * editable size here. Keyboard resizing asks this, so Shift+Arrow keeps the
 * ratio the Size fields keep (PPT-B151).
 */
export function presentationPictureExtentsForSide(
  block: SizedPresentationPictureBlock,
  axis: PresentationPictureSizeAxis,
  target: number,
): Readonly<{ width: number; height: number }> | null {
  const receipt = presentationPictureSizeReceipt(block);
  return receipt ? resizedPictureExtents(receipt, axis, Math.round(target)) : null;
}

/** Reuse the table-frame cm display primitive without sharing resize semantics. */
export function presentationPictureSizeDisplayCm(value: unknown): string {
  const centimetres = tableFrameSizeCm(value);
  return centimetres === null ? "" : centimetres.toFixed(2);
}

/** Parse one local two-decimal centimetre draft into exact integer EMU. */
export function presentationPictureSizeEmuFromCmDraft(
  value: string,
): number | null {
  const normalized = value.trim().replace(",", ".");
  if (!/^(?:[0-9]+(?:\.[0-9]{0,2})?|\.[0-9]{1,2})$/.test(normalized)) {
    return null;
  }
  const centimetres = Number(normalized);
  if (!Number.isFinite(centimetres) || centimetres <= 0) return null;
  const emu = Math.round(centimetres * TABLE_FRAME_EMU_PER_CM);
  return exactPictureSizeEmu(emu) ? emu : null;
}

/** PowerPoint for the web exposes Picture Format dimensions in inches. */
export function presentationPictureSizeDisplayInches(value: unknown): string {
  return exactPictureSizeEmu(value) ? (value / EMU_PER_INCH).toFixed(2) : "";
}

/** Parse one local two-decimal inch draft into exact integer EMU. */
export function presentationPictureSizeEmuFromInchesDraft(
  value: string,
): number | null {
  const normalized = value.trim().replace(",", ".");
  if (!/^(?:[0-9]+(?:\.[0-9]{0,2})?|\.[0-9]{1,2})$/.test(normalized)) {
    return null;
  }
  const inches = Number(normalized);
  if (!Number.isFinite(inches) || inches <= 0) return null;
  const emu = Math.round(inches * EMU_PER_INCH);
  return exactPictureSizeEmu(emu) ? emu : null;
}

export function stepPresentationPictureSizeEmu(
  current: number,
  direction: PresentationPictureSizeStepDirection,
  stepEmu: number = PRESENTATION_PICTURE_SIZE_STEP_EMU,
): number | null {
  if (!exactPictureSizeEmu(current)) return null;
  const step =
    Number.isSafeInteger(stepEmu) && stepEmu > 0
      ? stepEmu
      : PRESENTATION_PICTURE_SIZE_STEP_EMU;
  const target = direction === "increment" ? current + step : current - step;
  return exactPictureSizeEmu(target) ? target : null;
}

/** One live-target, one-transaction Picture Size field/stepper authority. */
export function createPresentationPictureSizeCommandAdapter<
  TBlock extends SizedPresentationPictureBlock,
  TSlide extends PresentationPictureSizeSlide<TBlock>,
>(
  accessors: PresentationPictureSizeCommandAccessors<TBlock, TSlide>,
): PresentationPictureSizeCommandAdapter<TBlock, TSlide> {
  const target = () => {
    if (accessors.canEdit?.() === false) return null;
    const slides = accessors.getSlides();
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      slides,
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    if (!resolved) return null;
    const receipt = presentationPictureSizeReceipt(resolved.block);
    return receipt ? { ...resolved, slides, receipt } : null;
  };
  const canExecute = (axis: PresentationPictureSizeAxis): boolean =>
    (axis === "width" || axis === "height") && target() !== null;

  const execute = (payload: PresentationPictureSizeCommandPayload): boolean => {
    if (payload.axis !== "width" && payload.axis !== "height") return false;
    const resolved = target();
    if (!resolved || resolved.block.id !== payload.blockId) return false;
    const targetEmu =
      payload.kind === "absolute"
        ? payload.targetEmu
        : stepPresentationPictureSizeEmu(
            resolved.receipt[payload.axis],
            payload.direction,
            payload.stepEmu,
          );
    if (targetEmu === null) return false;
    const extents = resizedPictureExtents(
      resolved.receipt,
      payload.axis,
      targetEmu,
    );
    if (
      !extents ||
      (extents.width === resolved.receipt.width &&
        extents.height === resolved.receipt.height)
    ) {
      return false;
    }
    const replacement = {
      ...resolved.block,
      width: extents.width,
      height: extents.height,
    } as TBlock;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? replacement : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  // Same target/receipt and commit path as the Size fields, but both dimensions
  // and the opposite-edge anchor land together in ONE undoable transaction.
  const beginResize = (
    blockId: string,
    bounds: DragBounds,
  ): PresentationPictureResizeGesture | null => {
    const initial = target();
    if (
      !initial ||
      initial.block.id !== blockId ||
      !Number.isSafeInteger(initial.block.left) ||
      !Number.isSafeInteger(initial.block.top) ||
      !Number.isSafeInteger(bounds.width) || bounds.width <= 0 ||
      !Number.isSafeInteger(bounds.height) || bounds.height <= 0
    )
      return null;
    const base: BlockGeometry = {
      left: initial.block.left!,
      top: initial.block.top!,
      width: initial.receipt.width,
      height: initial.receipt.height,
    };
    const generation = accessors.getGeneration?.();
    let latest = base;
    const startFlip = { flipH: initial.block.flipH === true, flipV: initial.block.flipV === true };
    let latestFlip = startFlip;
    let ended = false;
    const live = () => {
      if (ended || accessors.getGeneration?.() !== generation) return null;
      const resolved = target();
      return resolved?.slide.id === initial.slide.id &&
        resolved.block === initial.block
        ? resolved
        : null;
    };
    return {
      preview(handle, dx, dy, modifiers) {
        if (!live() || !Number.isFinite(dx) || !Number.isFinite(dy))
          return null;
        const angle = ((initial.block.rotation ?? 0) * Math.PI) / 180;
        if (!Number.isFinite(angle)) return null;
        const cos = Math.cos(angle),
          sin = Math.sin(angle);
        const flipX = initial.block.flipH ? -1 : 1;
        const flipY = initial.block.flipV ? -1 : 1;
        // A handle dragged through the pinned edge mirrors the picture in the
        // same gesture, as PowerPoint flips it; the crossed axis toggles the
        // picture's own flag and dragging back restores it.
        const through = applyResizeThroughZero(
          base,
          handle,
          (cos * dx + sin * dy) * flipX,
          (-sin * dx + cos * dy) * flipY,
          bounds,
          {
            ...modifiers,
            lockAspect: initial.receipt.aspectLocked || modifiers?.lockAspect,
          },
          PRESENTATION_PICTURE_SIZE_MIN_EMU,
        );
        const raw = through.geom;
        // Generic geometry rounds independently. Canonicalize to the exact native
        // ratio (including one-EMU rounding) before preview AND before saving.
        const axis = handle === "n" || handle === "s" ? "height" : "width";
        const size = initial.receipt.aspectLocked
          ? resizedPictureExtents(initial.receipt, axis, raw[axis])
          : { width: raw.width, height: raw.height };
        if (
          !size ||
          !exactPictureSizeEmu(size.width) ||
          !exactPictureSizeEmu(size.height)
        )
          return null;
        const centered = modifiers?.fromCenter;
        // The pinned edge stays put; a crossed axis puts the picture on the
        // far side of it (the "e" grip past the left edge spans [left - w, left]).
        const localLeft =
          centered || (!handle.includes("e") && !handle.includes("w"))
            ? base.left + (base.width - size.width) / 2
            : handle.includes("w")
              ? through.flipX ? base.left + base.width : base.left + base.width - size.width
              : through.flipX ? base.left - size.width : base.left;
        const localTop =
          centered || (!handle.includes("n") && !handle.includes("s"))
            ? base.top + (base.height - size.height) / 2
            : handle.includes("n")
              ? through.flipY ? base.top + base.height : base.top + base.height - size.height
              : through.flipY ? base.top - size.height : base.top;
        latestFlip = {
          flipH: initial.block.flipH === true ? !through.flipX : through.flipX,
          flipV: initial.block.flipV === true ? !through.flipY : through.flipY,
        };
        const centerDx =
          (localLeft + size.width / 2 - base.left - base.width / 2) * flipX;
        const centerDy =
          (localTop + size.height / 2 - base.top - base.height / 2) * flipY;
        latest = {
          left: Math.round(
            base.left +
              base.width / 2 +
              cos * centerDx -
              sin * centerDy -
              size.width / 2,
          ),
          top: Math.round(
            base.top +
              base.height / 2 +
              sin * centerDx +
              cos * centerDy -
              size.height / 2,
          ),
          ...size,
        };
        return { ...latest };
      },
      latestFlip() {
        return { ...latestFlip };
      },
      commit() {
        const resolved = live();
        ended = true;
        const flipped = latestFlip.flipH !== startFlip.flipH || latestFlip.flipV !== startFlip.flipV;
        if (
          !resolved ||
          (!flipped &&
            Object.keys(base).every(
              (key) =>
                base[key as keyof BlockGeometry] ===
                latest[key as keyof BlockGeometry],
            ))
        )
          return false;
        const mirrored: Record<string, unknown> = { ...resolved.block, ...latest };
        if (flipped) {
          // A false flag is deleted: an un-mirrored picture serializes as before.
          for (const axis of ["flipH", "flipV"] as const) {
            if (latestFlip[axis]) mirrored[axis] = true;
            else delete mirrored[axis];
          }
        }
        const replacement = mirrored as TBlock;
        accessors.commitSlides(
          resolved.slides.map((slide, index) =>
            index === resolved.slideIndex
              ? ({
                  ...slide,
                  blocks: slide.blocks.map((block) =>
                    block === resolved.block ? replacement : block,
                  ),
                } as TSlide)
              : slide,
          ),
        );
        return true;
      },
      cancel() {
        ended = true;
      },
    };
  };

  return Object.freeze({
    featureIds: Object.freeze({
      width: PICTURE_SIZE_WIDTH_FEATURE_ID,
      height: PICTURE_SIZE_HEIGHT_FEATURE_ID,
    }),
    canExecute,
    execute,
    beginResize,
  });
}

function hasExactNativePictureIdentity(
  block: ArrangeablePresentationBlock,
): boolean {
  if (block.type !== "image" || isBlockKeptReadOnly(block)) return false;
  const source = block.source;
  if (!source) return false;
  return (
    typeof source.slideId === "string" &&
    /^sld[1-9][0-9]*$/.test(source.slideId) &&
    typeof source.shapeId === "string" &&
    /^shp[1-9][0-9]*$/.test(source.shapeId)
  );
}

interface NativeShapeProjectionGroup<
  TBlock extends ArrangeablePresentationBlock,
> {
  readonly id: string;
  readonly shapeId: string;
  readonly sourceOrder: number;
  readonly blocks: readonly TBlock[];
}

/**
 * An object the author added to an imported slide: it names no source at
 * all, so it is none of the file's shapes and has no place in the native
 * order. The save writes it before the file shape that follows it in the
 * list (`authored_append_units`).
 */
function isAuthoredSlideObject(block: ArrangeablePresentationBlock): boolean {
  return block.source === undefined || block.source === null;
}










function nativeOrderShapeId(source: Record<string, unknown> | undefined): unknown {
  const shapeId = source?.shapeId;
  if (shapeId !== undefined && shapeId !== null && shapeId !== "") return shapeId;
  const view = source?.nativeSmartArtView;
  if (
    view &&
    typeof view === "object" &&
    !Array.isArray(view) &&
    (view as Record<string, unknown>).schemaVersion === "pptx-smartart-drawing-v1" &&
    typeof (view as Record<string, unknown>).frameShapeId === "string"
  ) {
    return (view as Record<string, unknown>).frameShapeId;
  }
  if (typeof source?.nativeSmartArtBlocker === "string" && typeof source.frameShapeId === "string") {
    return source.frameShapeId;
  }
  return shapeId;
}

function nativeShapeProjectionGroups<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  slide: TSlide,
  /** Pass over the objects the author added (they keep their own places in
   * the list) rather than refuse the slide. */
  withAuthored = false,
): readonly NativeShapeProjectionGroup<TBlock>[] | null {
  const slideId = slide.source?.slideId;
  if (typeof slideId !== "string" || !/^sld[1-9][0-9]*$/.test(slideId)) {
    return null;
  }
  const shapeToOrder = new Map<string, number>();
  const orderToShape = new Map<number, string>();
  const seenBlockIds = new Set<string>();
  const seenGroupIds = new Set<string>();
  const groups: NativeShapeProjectionGroup<TBlock>[] = [];
  for (const block of slide.blocks) {
    const source = block.source;
    if (
      source?.inheritedFrom === "layout" ||
      source?.inheritedFrom === "master" ||
      // A placeholder materialised by a Shape Format edit is a slide object
      // the package does not hold yet: it gains its `cNvPr/@id` only when the
      // writer inserts it. Until then it addresses no native shape, so it sits
      // outside the native z-order the same way its inherited reading did --
      // rather than voiding the whole slide's native ordering.
      source?.materializedPlaceholder !== undefined ||
      (withAuthored && isAuthoredSlideObject(block))
    ) {
      continue;
    }
    const shapeId = nativeOrderShapeId(source);
    const sourceOrder = block.sourceOrder;
    if (
      !block.id ||
      seenBlockIds.has(block.id) ||
      source?.slideId !== slideId ||
      typeof shapeId !== "string" ||
      !/^shp[1-9][0-9]*$/.test(shapeId) ||
      !Number.isInteger(sourceOrder) ||
      (sourceOrder as number) < 0 ||
      (shapeToOrder.has(shapeId) &&
        shapeToOrder.get(shapeId) !== sourceOrder) ||
      (orderToShape.has(sourceOrder as number) &&
        orderToShape.get(sourceOrder as number) !== shapeId)
    ) {
      return null;
    }
    seenBlockIds.add(block.id);
    shapeToOrder.set(shapeId, sourceOrder as number);
    orderToShape.set(sourceOrder as number, shapeId);
    const groupId = `${sourceOrder}:${shapeId}`;
    const current = groups.at(-1);
    if (current?.id === groupId) {
      groups[groups.length - 1] = {
        ...current,
        blocks: [...current.blocks, block],
      };
      continue;
    }
    if (seenGroupIds.has(groupId)) return null;
    seenGroupIds.add(groupId);
    groups.push({
      id: groupId,
      shapeId,
      sourceOrder: sourceOrder as number,
      blocks: [block],
    });
  }
  const sourceOrders = [...orderToShape.keys()].sort((a, b) => a - b);
  if (
    sourceOrders.some(
      (sourceOrder, index) =>
        index > 0 && sourceOrder !== sourceOrders[index - 1] + 1,
    )
  ) {
    return null;
  }
  return groups;
}

function blocksWithNativeGroupsReordered<
  TBlock extends ArrangeablePresentationBlock,
>(
  blocks: readonly TBlock[],
  groups: readonly NativeShapeProjectionGroup<TBlock>[],
  withAuthored = false,
): TBlock[] {
  const reordered = groups.flatMap((group) => group.blocks);
  let cursor = 0;
  return blocks.map((block) => {
    if (
      block.source?.inheritedFrom === "layout" ||
      block.source?.inheritedFrom === "master" ||
      block.source?.materializedPlaceholder !== undefined ||
      (withAuthored && isAuthoredSlideObject(block))
    ) {
      return block;
    }
    const next = reordered[cursor];
    cursor += 1;
    return next ?? block;
  });
}

interface AuthorizedNativeShapeOrder<
  TBlock extends ArrangeablePresentationBlock,
> {
  readonly steps: readonly Readonly<{
    shapeId: string;
    direction: ZOrderStep;
  }>[];
  readonly groups: readonly NativeShapeProjectionGroup<TBlock>[];
}

function sameNativeShapeGroupOrder<TBlock extends ArrangeablePresentationBlock>(
  left: readonly NativeShapeProjectionGroup<TBlock>[],
  right: readonly NativeShapeProjectionGroup<TBlock>[],
): boolean {
  return (
    left.length === right.length &&
    left.every((group, index) => group.shapeId === right[index]?.shapeId)
  );
}

function parseNativeShapeOrderSteps(
  raw: unknown,
  slideId: string,
): readonly Readonly<{ shapeId: string; direction: ZOrderStep }>[] | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const target = raw as Record<string, unknown>;
  if (
    Object.keys(target).length !== 3 ||
    !Object.hasOwn(target, "schemaVersion") ||
    !Object.hasOwn(target, "slideId") ||
    !Object.hasOwn(target, "steps") ||
    target.schemaVersion !== "presentation-native-shape-order-v2" ||
    target.slideId !== slideId ||
    !Array.isArray(target.steps)
  ) {
    return null;
  }
  const steps: Array<Readonly<{ shapeId: string; direction: ZOrderStep }>> = [];
  for (const rawStep of target.steps) {
    if (!rawStep || typeof rawStep !== "object" || Array.isArray(rawStep)) {
      return null;
    }
    const step = rawStep as Record<string, unknown>;
    if (
      Object.keys(step).length !== 2 ||
      !Object.hasOwn(step, "shapeId") ||
      !Object.hasOwn(step, "direction") ||
      typeof step.shapeId !== "string" ||
      !/^shp[1-9][0-9]*$/.test(step.shapeId) ||
      !isZOrderStep(step.direction)
    ) {
      return null;
    }
    steps.push(
      Object.freeze({
        shapeId: step.shapeId,
        direction: step.direction,
      }),
    );
  }
  return Object.freeze(steps);
}

function replayNativeShapeOrderSteps<
  TBlock extends ArrangeablePresentationBlock,
>(
  trustedGroups: readonly NativeShapeProjectionGroup<TBlock>[],
  steps: readonly Readonly<{ shapeId: string; direction: ZOrderStep }>[],
): readonly NativeShapeProjectionGroup<TBlock>[] | null {
  // One swap per step, each group's place kept in a map: a journal of a few
  // thousand steps (a large slide restacked whole) replays in linear time.
  const replayed = [...trustedGroups];
  const placeOf = new Map(replayed.map((group, index) => [group.shapeId, index]));
  for (const step of steps) {
    const groupIndex = placeOf.get(step.shapeId);
    const group = groupIndex === undefined ? undefined : replayed[groupIndex];
    if (groupIndex === undefined || !group || group.blocks.length !== 1) {
      return null;
    }
    const neighborIndex =
      step.direction === "forward" ? groupIndex + 1 : groupIndex - 1;
    const neighbor = replayed[neighborIndex];
    if (!neighbor) return null;
    replayed[neighborIndex] = group;
    replayed[groupIndex] = neighbor;
    placeOf.set(group.shapeId, neighborIndex);
    placeOf.set(neighbor.shapeId, groupIndex);
  }
  return replayed;
}

/**
 * The fewest one-object steps that take the file's order (`trustedGroups`,
 * by their places) to `finalGroups`: each pair the new order swaps is
 * swapped once, by whichever of the two is one direct object (the journal
 * steps only those; an object projected as several blocks -- a SmartArt
 * diagram, a group -- is stepped over). Null when a pair of such objects
 * would have to swap, which no step can say.
 */
function compactNativeShapeOrderSteps<
  TBlock extends ArrangeablePresentationBlock,
>(
  trustedGroups: readonly NativeShapeProjectionGroup<TBlock>[],
  finalGroups: readonly NativeShapeProjectionGroup<TBlock>[],
): Readonly<{ shapeId: string; direction: ZOrderStep }>[] | null {
  const rank = new Map(finalGroups.map((group, index) => [group.id, index]));
  if (rank.size !== trustedGroups.length || trustedGroups.some((group) => !rank.has(group.id))) {
    return null;
  }
  const order = [...trustedGroups];
  const steps: Array<Readonly<{ shapeId: string; direction: ZOrderStep }>> = [];
  for (let swapped = true; swapped; ) {
    swapped = false;
    for (let index = 0; index + 1 < order.length; index += 1) {
      const left = order[index];
      const right = order[index + 1];
      if (!left || !right || (rank.get(left.id) ?? 0) <= (rank.get(right.id) ?? 0)) continue;
      if (left.blocks.length === 1) {
        steps.push(Object.freeze({ shapeId: left.shapeId, direction: "forward" as const }));
      } else if (right.blocks.length === 1) {
        steps.push(Object.freeze({ shapeId: right.shapeId, direction: "backward" as const }));
      } else {
        return null;
      }
      order[index] = right;
      order[index + 1] = left;
      swapped = true;
    }
  }
  return steps;
}

/**
 * The journal a slide records after a restack: `steps` (the earlier ones and
 * the new) while they are few, else the fewest steps from the file's order to
 * `finalGroups` (`presentationNativeOrderStepBound`). Null only when the new
 * order cannot be stepped at all.
 */
function nativeShapeOrderJournal<
  TBlock extends ArrangeablePresentationBlock,
>(
  groups: readonly NativeShapeProjectionGroup<TBlock>[],
  steps: readonly Readonly<{ shapeId: string; direction: ZOrderStep }>[],
  finalGroups: readonly NativeShapeProjectionGroup<TBlock>[],
): readonly Readonly<{ shapeId: string; direction: ZOrderStep }>[] | null {
  if (steps.length <= PRESENTATION_PICTURE_ARRANGE_MAX_STEPS) return steps;
  const trusted = [...groups].sort((left, right) => left.sourceOrder - right.sourceOrder);
  const compact = compactNativeShapeOrderSteps(trusted, finalGroups);
  return compact && compact.length <= presentationNativeOrderStepBound(trusted.length)
    ? compact
    : null;
}

function authorizedNativeShapeOrder<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  slide: TSlide,
  groups: readonly NativeShapeProjectionGroup<TBlock>[],
): AuthorizedNativeShapeOrder<TBlock> | null {
  const slideId = slide.source?.slideId;
  if (typeof slideId !== "string") return null;
  const trustedGroups = [...groups].sort(
    (left, right) => left.sourceOrder - right.sourceOrder,
  );
  if (slide.nativeShapeOrderTarget === undefined) {
    return sameNativeShapeGroupOrder(groups, trustedGroups)
      ? { steps: Object.freeze([]), groups }
      : null;
  }
  const steps = parseNativeShapeOrderSteps(
    slide.nativeShapeOrderTarget,
    slideId,
  );
  if (!steps || steps.length > presentationNativeOrderStepBound(groups.length)) return null;
  const replayed = replayNativeShapeOrderSteps(trustedGroups, steps);
  return replayed && sameNativeShapeGroupOrder(replayed, groups)
    ? { steps, groups }
    : null;
}

interface PresentationNativeSelectionPaneTarget<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
> {
  readonly slide: TSlide;
  readonly block: TBlock;
  readonly groups: readonly NativeShapeProjectionGroup<TBlock>[];
  readonly steps: readonly Readonly<{
    shapeId: string;
    direction: ZOrderStep;
  }>[];
  readonly group: NativeShapeProjectionGroup<TBlock>;
  readonly groupIndex: number;
}

function presentationNativeSelectionPaneTarget<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  slide: TSlide,
  blockId: string,
  withAuthored = false,
): PresentationNativeSelectionPaneTarget<TBlock, TSlide> | null {
  const block = slide.blocks.find((candidate) => candidate.id === blockId);
  if (
    !block ||
    isBlockKeptReadOnly(block) ||
    // A nameless object (LibreOffice writes name="") is still a direct
    // native object the pane controls; a name it has must be a real one.
    (block.name !== undefined &&
      (typeof block.name !== "string" ||
        block.name.trim().length === 0 ||
        block.name.length > 200)) ||
    (block.hidden !== undefined && typeof block.hidden !== "boolean") ||
    block.sourceOrder !== block.zIndex ||
    block.source?.inheritedFrom !== undefined ||
    block.source?.materializedPlaceholder !== undefined
  ) {
    return null;
  }
  const groups = nativeShapeProjectionGroups<TBlock, TSlide>(slide, withAuthored);
  if (!groups) return null;
  const authorized = authorizedNativeShapeOrder<TBlock, TSlide>(slide, groups);
  if (!authorized) return null;
  const groupIndex = authorized.groups.findIndex((candidate) =>
    candidate.blocks.some((member) => member.id === blockId),
  );
  const group = authorized.groups[groupIndex];
  if (!group || group.blocks.length !== 1 || group.blocks[0] !== block) {
    return null;
  }
  return {
    slide,
    block,
    groups: authorized.groups,
    steps: authorized.steps,
    group,
    groupIndex,
  };
}

/**
 * Whether one Selection Pane row is an exact direct native slide object.
 *
 * This deliberately authenticates the complete direct-object projection and
 * the accumulated native order target. A layout/master projection, a flattened
 * group, a partial source identity, or a mixed native/model slide stays
 * read-only rather than borrowing a nearby shape id.
 */
export function presentationNativeSelectionPaneObjectIsWritable<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(slide: TSlide, blockId: string): boolean {
  return presentationNativeSelectionPaneTarget(slide, blockId) !== null;
}

/** One adjacent native Selection Pane order step, on the existing v2 journal. */
export function applyPresentationNativeSelectionPaneOrderStep<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(slide: TSlide, blockId: string, direction: ZOrderStep): TSlide | null {
  if (!isZOrderStep(direction)) return null;
  const resolved = presentationNativeSelectionPaneTarget(slide, blockId);
  if (!resolved) return null;
  const neighborIndex =
    direction === "forward"
      ? resolved.groupIndex + 1
      : resolved.groupIndex - 1;
  if (!resolved.groups[neighborIndex]) return null;
  const nextGroups = moveItemOneStep(
    resolved.groups,
    resolved.group.id,
    direction,
  );
  const journal = nativeShapeOrderJournal(
    resolved.groups,
    [...resolved.steps, { shapeId: resolved.group.shapeId, direction }],
    nextGroups,
  );
  if (!journal) return null;
  const nextBlocks = blocksWithNativeGroupsReordered(
    resolved.slide.blocks,
    nextGroups,
  );
  if (
    nextBlocks.every(
      (block, index) => block.id === resolved.slide.blocks[index]?.id,
    )
  ) {
    return null;
  }
  return {
    ...resolved.slide,
    blocks: nextBlocks,
    nativeShapeOrderTarget: {
      schemaVersion: "presentation-native-shape-order-v2",
      slideId: resolved.slide.source?.slideId as string,
      steps: [...journal],
    },
  } as TSlide;
}

/** The four z-order commands of Home / Shape Format ▸ Arrange and the context menu. */
export type PresentationNativeOrderCommand = "front" | "forward" | "backward" | "back";
















export function applyPresentationNativeOrderCommand<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  slide: TSlide,
  targetIds: readonly string[],
  command: PresentationNativeOrderCommand,
): TSlide | "unchanged" | null {
  if (targetIds.length === 0) return null;
  const targets = new Map<string, NativeShapeProjectionGroup<TBlock>>();
  let groups: readonly NativeShapeProjectionGroup<TBlock>[] | null = null;
  let steps: readonly Readonly<{ shapeId: string; direction: ZOrderStep }>[] = [];
  for (const blockId of targetIds) {
    const resolved = presentationNativeSelectionPaneTarget<TBlock, TSlide>(slide, blockId);
    if (!resolved) return null;
    groups ??= resolved.groups;
    steps = resolved.steps;
    targets.set(resolved.group.id, resolved.group);
  }
  if (!groups) return null;
  let order = [...groups];
  const taken: Array<Readonly<{ shapeId: string; direction: ZOrderStep }>> = [];
  const isTarget = (index: number) => targets.has(order[index]?.id ?? "");
  const stepOnce = (groupId: string, direction: ZOrderStep): boolean => {
    const index = order.findIndex((group) => group.id === groupId);
    const neighbor = direction === "forward" ? index + 1 : index - 1;
    if (index < 0 || neighbor < 0 || neighbor >= order.length || isTarget(neighbor)) {
      return false;
    }
    order = moveItemOneStep(order, groupId, direction) as NativeShapeProjectionGroup<TBlock>[];
    const moved = order[neighbor];
    if (!moved) return false;
    taken.push({ shapeId: moved.shapeId, direction });
    return true;
  };
  const direction: ZOrderStep = command === "front" || command === "forward" ? "forward" : "backward";
  // The target nearest the way it moves goes first, so the others never step
  // over it and the targets keep their relative order.
  const moving = order.filter((group) => targets.has(group.id)).map((group) => group.id);
  if (direction === "forward") moving.reverse();
  for (const groupId of moving) {
    if (command === "forward" || command === "backward") {
      stepOnce(groupId, direction);
      continue;
    }
    while (stepOnce(groupId, direction)) {
      // to the end of the slide: one step per object passed
    }
  }
  if (taken.length === 0) return "unchanged";
  const journal = nativeShapeOrderJournal(groups, [...steps, ...taken], order);
  if (!journal) return null;
  return {
    ...slide,
    blocks: blocksWithNativeGroupsReordered(slide.blocks, order),
    nativeShapeOrderTarget: {
      schemaVersion: "presentation-native-shape-order-v2",
      slideId: slide.source?.slideId as string,
      steps: [...journal],
    },
  } as TSlide;
}










export function applyPresentationNativeObjectOrder<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  slide: TSlide,
  nextBlocks: readonly TBlock[],
  /** The slide also holds objects the author added: the steps order the
   * file's objects among themselves and the slide takes `nextBlocks` whole,
   * so each added object stands where the list puts it. */
  withAuthored = false,
): TSlide | "unchanged" | null {
  const groups = nativeShapeProjectionGroups<TBlock, TSlide>(slide, withAuthored);
  if (!groups) return null;
  const authorized = authorizedNativeShapeOrder<TBlock, TSlide>(slide, groups);
  if (!authorized) return null;
  const groupOf = new Map<string, NativeShapeProjectionGroup<TBlock>>();
  for (const group of authorized.groups) {
    for (const block of group.blocks) groupOf.set(block.id, group);
  }
  // The new order of the native objects, each object's blocks side by side.
  const wanted: NativeShapeProjectionGroup<TBlock>[] = [];
  const seen = new Set<string>();
  for (let index = 0; index < nextBlocks.length; index += 1) {
    const group = groupOf.get(nextBlocks[index].id);
    if (!group) continue;
    if (seen.has(group.id)) return null;
    seen.add(group.id);
    if (group.blocks.some((block, offset) => nextBlocks[index + offset]?.id !== block.id)) {
      return null;
    }
    wanted.push(group);
    index += group.blocks.length - 1;
  }
  if (wanted.length !== authorized.groups.length) return null;
  const rank = new Map(wanted.map((group, index) => [group.id, index]));
  const movable = (group: NativeShapeProjectionGroup<TBlock>) =>
    group.blocks.length === 1 &&
    presentationNativeSelectionPaneTarget<TBlock, TSlide>(slide, group.blocks[0].id, withAuthored) !== null;
  let order = [...authorized.groups];
  const taken: Array<Readonly<{ shapeId: string; direction: ZOrderStep }>> = [];
  for (let swapped = true; swapped; ) {
    swapped = false;
    for (let index = 0; index + 1 < order.length; index += 1) {
      const left = order[index];
      const right = order[index + 1];
      if ((rank.get(left.id) ?? 0) <= (rank.get(right.id) ?? 0)) continue;
      if (movable(left)) {
        taken.push({ shapeId: left.shapeId, direction: "forward" });
      } else if (movable(right)) {
        taken.push({ shapeId: right.shapeId, direction: "backward" });
      } else {
        return null;
      }
      order = [...order.slice(0, index), right, left, ...order.slice(index + 2)];
      swapped = true;
    }
  }
  if (taken.length === 0) return "unchanged";
  const journal = nativeShapeOrderJournal(authorized.groups, [...authorized.steps, ...taken], order);
  if (!journal) return null;
  return {
    ...slide,
    blocks: withAuthored ? [...nextBlocks] : blocksWithNativeGroupsReordered(slide.blocks, order),
    nativeShapeOrderTarget: {
      schemaVersion: "presentation-native-shape-order-v2",
      slideId: slide.source?.slideId as string,
      steps: [...journal],
    },
  } as TSlide;
}

/** One live-target, one-adjacent-layer Picture Arrange transaction. */
export function createPresentationPictureArrangeCommandAdapter<
  TBlock extends ArrangeablePresentationBlock,
  TSlide extends PresentationPictureArrangeSlide<TBlock>,
>(
  accessors: PresentationPictureArrangeCommandAccessors<TBlock, TSlide>,
): PresentationPictureArrangeCommandAdapter<TBlock, TSlide> {
  const target = () => {
    const slides = accessors.getSlides();
    const slideIndex = accessors.getActiveIndex();
    const selectedPictureId = accessors.getSelectedPictureId();
    if (!selectedPictureId || !Number.isInteger(slideIndex) || slideIndex < 0) {
      return null;
    }
    const slide = slides[slideIndex];
    if (!slide) return null;
    const blockIndex = slide.blocks.findIndex(
      (block) => block.id === selectedPictureId,
    );
    const block = slide.blocks[blockIndex];
    if (!block || !hasExactNativePictureIdentity(block)) return null;
    const groups = nativeShapeProjectionGroups(slide);
    if (!groups) return null;
    const authorized = authorizedNativeShapeOrder(slide, groups);
    if (!authorized) return null;
    const groupIndex = authorized.groups.findIndex((group) =>
      group.blocks.some((candidate) => candidate.id === selectedPictureId),
    );
    const group = authorized.groups[groupIndex];
    if (!group || group.blocks.length !== 1 || group.blocks[0] !== block) {
      return null;
    }
    return {
      slides,
      slide,
      slideIndex,
      block,
      groups: authorized.groups,
      steps: authorized.steps,
      group,
      groupIndex,
    };
  };

  const canExecute = (direction: ZOrderStep): boolean => {
    if (!isZOrderStep(direction)) return false;
    const resolved = target();
    if (!resolved) return false;
    const neighborIndex =
      direction === "forward"
        ? resolved.groupIndex + 1
        : resolved.groupIndex - 1;
    return resolved.groups[neighborIndex] !== undefined;
  };

  const execute = (
    payload: PresentationPictureArrangeCommandPayload,
  ): boolean => {
    if (!isZOrderStep(payload.direction)) return false;
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== payload.blockId
    ) {
      return false;
    }
    const neighborIndex =
      payload.direction === "forward"
        ? resolved.groupIndex + 1
        : resolved.groupIndex - 1;
    if (!resolved.groups[neighborIndex]) return false;
    const nextGroups = moveItemOneStep(
      resolved.groups,
      resolved.group.id,
      payload.direction,
    );
    const nextBlocks = blocksWithNativeGroupsReordered(
      resolved.slide.blocks,
      nextGroups,
    );
    if (
      nextBlocks.every(
        (block, index) => block.id === resolved.slide.blocks[index]?.id,
      )
    ) {
      return false;
    }
    const journal = nativeShapeOrderJournal(
      resolved.groups,
      [
        ...resolved.steps,
        { shapeId: resolved.group.shapeId, direction: payload.direction },
      ],
      nextGroups,
    );
    if (!journal) return false;
    const nextSlides = resolved.slides.map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: nextBlocks,
            nativeShapeOrderTarget: {
              schemaVersion: "presentation-native-shape-order-v2",
              slideId: resolved.slide.source?.slideId as string,
              steps: [...journal],
            },
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    return true;
  };

  return Object.freeze({
    featureIds: Object.freeze({
      forward: PICTURE_BRING_FORWARD_FEATURE_ID,
      backward: PICTURE_SEND_BACKWARD_FEATURE_ID,
    }),
    canExecute,
    execute,
  });
}








function pictureCropIsPreserveOnly(block: { source?: Record<string, unknown> | null }): boolean {
  const source = block.source;
  return (
    source?.pictureCropDisposition === "preserve-only" ||
    (typeof source?.oleObjectPreview === "object" && source.oleObjectPreview !== null)
  );
}

/**
 * Presentation-owned adapter for Picture Format → Crop / Fill / Fit.
 *
 * It reads the live selected picture at dispatch time, commits a cloned target
 * slide exactly once, and changes editor-only crop mode only after that commit
 * succeeds. Free-crop gestures reuse the same commit boundary while retaining
 * their stricter 0..1 clamp in `presentation-crop-group`.
 */
export function createPresentationPictureCropCommandAdapter<
  TBlock extends CroppablePresentationPictureBlock,
  TSlide extends PresentationPictureCropSlide<TBlock>,
>(
  accessors: PresentationPictureCropCommandAccessors<TBlock, TSlide>,
): PresentationPictureCropCommandAdapter<TBlock, TSlide> {
  // A crop is a new frame for the picture: a locked one (Lock, or the
  // file's noResize / noCrop) keeps the one it has.
  const target = () => {
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      accessors.getSlides(),
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved && !isLockedFor(resolved.block, "crop") ? resolved : null;
  };

  const canExecute = (mode: PresentationPictureCropMode): boolean => {
    const resolved = target();
    if (
      !resolved ||
      pictureCropIsPreserveOnly(resolved.block) ||
      !positiveFinite(resolved.block.width) ||
      !positiveFinite(resolved.block.height)
    ) {
      return false;
    }
    return (
      mode === "crop" ||
      (positiveFinite(resolved.block.pixelWidth) &&
        positiveFinite(resolved.block.pixelHeight))
    );
  };

  // Crop → Aspect Ratio also moves the picture FRAME (PowerPoint keeps the
  // picture's drawn size and gives the crop frame the ratio), so the frame and
  // the crop are committed in the same replacement and the same undo step.
  const commitCrop = (
    blockId: string,
    rawCrop: CropRect | null,
    enterAfterCommit: boolean,
    frame?: PictureAspectRatioCropGeometry,
  ): boolean => {
    const resolved = target();
    if (
      !resolved ||
      resolved.block.id !== blockId ||
      pictureCropIsPreserveOnly(resolved.block)
    ) {
      return false;
    }
    const normalized =
      rawCrop === null ? IDENTITY_CROP : normalizeSignedCropRect(rawCrop);
    if (!normalized) return false;
    const frameChanged =
      frame !== undefined &&
      (frame.width !== resolved.block.width ||
        frame.height !== resolved.block.height ||
        (frame.left !== undefined && frame.left !== resolved.block.left) ||
        (frame.top !== undefined && frame.top !== resolved.block.top));
    if (sameCropRect(resolved.block.cropRect, normalized) && !frameChanged) {
      if (enterAfterCommit) {
        accessors.setCropMode(blockId, "enter");
      }
      return true;
    }

    let replacement: TBlock;
    if (isIdentityCrop(normalized)) {
      const next = { ...resolved.block } as Record<string, unknown>;
      delete next.cropRect;
      replacement = next as TBlock;
    } else {
      replacement = {
        ...resolved.block,
        cropRect: normalized,
      } as TBlock;
    }
    if (frameChanged && frame) {
      // A native-size-backed picture takes the crop frame as its size
      // baseline so the Size spinner keeps stepping from the new frame; so
      // does an uploaded one, whose size receipt otherwise measured the
      // cropped frame against the raster's own pixels, found it off-ratio and
      // turned the Size boxes and handles off (PPT-B418).
      const source = resolved.block.source;
      const uploaded = (resolved.block as { sourceKind?: unknown }).sourceKind === "asset";
      const baseline =
        source?.pictureSizeSchemaVersion ===
          PRESENTATION_PICTURE_SIZE_SCHEMA_VERSION || uploaded
          ? {
              source: {
                ...source,
                [PICTURE_SIZE_CROP_FRAME_WIDTH_KEY]: frame.width,
                [PICTURE_SIZE_CROP_FRAME_HEIGHT_KEY]: frame.height,
              },
            }
          : {};
      replacement = {
        ...replacement,
        width: frame.width,
        height: frame.height,
        ...(frame.left !== undefined ? { left: frame.left } : {}),
        ...(frame.top !== undefined ? { top: frame.top } : {}),
        ...baseline,
      } as TBlock;
    }
    const nextSlides = accessors.getSlides().map((slide, slideIndex) =>
      slideIndex === resolved.slideIndex
        ? ({
            ...slide,
            blocks: slide.blocks.map((block) =>
              block === resolved.block ? replacement : block,
            ),
          } as TSlide)
        : slide,
    );
    accessors.commitSlides(nextSlides);
    if (enterAfterCommit) {
      accessors.setCropMode(blockId, "enter");
    }
    return true;
  };

  const execute = (mode: PresentationPictureCropMode): boolean => {
    const resolved = target();
    if (!resolved || !canExecute(mode)) return false;
    if (mode === "crop") {
      accessors.setCropMode(resolved.block.id, "toggle");
      return true;
    }
    const crop =
      mode === "fill"
        ? pictureFillCropRect(
            resolved.block.pixelWidth,
            resolved.block.pixelHeight,
            resolved.block.width,
            resolved.block.height,
          )
        : pictureFitCropRect(
            resolved.block.pixelWidth,
            resolved.block.pixelHeight,
            resolved.block.width,
            resolved.block.height,
          );
    return crop ? commitCrop(resolved.block.id, crop, true) : false;
  };

  const canApplyAspectRatio = (): boolean => {
    const resolved = target();
    // The ratio crop resizes the frame, so a picture whose size is locked or
    // whose native size is preserve-only is refused up front.
    return Boolean(
      resolved &&
      resolved.block.locked !== true &&
      !pictureCropIsPreserveOnly(resolved.block) &&
      resolved.block.source?.pictureSizeDisposition !== "preserve-only" &&
      positiveFinite(resolved.block.width) &&
      positiveFinite(resolved.block.height) &&
      positiveFinite(resolved.block.pixelWidth) &&
      positiveFinite(resolved.block.pixelHeight),
    );
  };

  const commitAspectRatio = (
    blockId: string,
    ratioId: PictureCropAspectRatioId,
  ): boolean => {
    const resolved = target();
    if (!resolved || resolved.block.id !== blockId || !canApplyAspectRatio()) {
      return false;
    }
    const geometry = pictureAspectRatioCropGeometry(resolved.block, ratioId);
    return geometry
      ? commitCrop(blockId, geometry.cropRect, false, geometry)
      : false;
  };

  return Object.freeze({
    featureIds: Object.freeze({
      crop: PICTURE_CROP_FEATURE_ID,
      fill: PICTURE_CROP_FILL_FEATURE_ID,
      fit: PICTURE_CROP_FIT_FEATURE_ID,
    }),
    canExecute,
    execute,
    canApplyAspectRatio,
    commitAspectRatio,
    commitFreeCrop: (blockId: string, crop: CropRect | null) => {
      // A free crop cuts the frame down to what is kept, where it was: the
      // kept part used to be stretched back over the old frame (PPT-B210).
      const resolved = target();
      const next = crop === null ? null : normalizeSignedCropRect(crop);
      if (!resolved || resolved.block.id !== blockId || !next) {
        return commitCrop(blockId, crop, false);
      }
      const previous =
        (resolved.block.cropRect && normalizeSignedCropRect(resolved.block.cropRect)) || IDENTITY_CROP;
      const { left, top, width, height } = resolved.block;
      if (
        previous.w <= 0 ||
        previous.h <= 0 ||
        ![left, top, width, height].every((value) => typeof value === "number" && Number.isFinite(value))
      ) {
        return commitCrop(blockId, crop, false);
      }
      const block = resolved.block as unknown as Record<string, unknown>;
      return commitCrop(blockId, crop, false, {
        cropRect: next,
        ...presentationFreeCropFrame(
          {
            left: left as number,
            top: top as number,
            width: width as number,
            height: height as number,
            rotation: typeof block.rotation === "number" ? block.rotation : 0,
            flipH: block.flipH === true,
            flipV: block.flipV === true,
          },
          previous,
          next,
        ),
      });
    },
  });
}










export function presentationFreeCropFrame(
  frame: {
    readonly left: number;
    readonly top: number;
    readonly width: number;
    readonly height: number;
    readonly rotation?: number;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
  },
  previous: CropRect,
  next: CropRect,
): { left: number; top: number; width: number; height: number } {
  const { left, top, width, height } = frame;
  const keptWidth = width * (next.w / previous.w);
  const keptHeight = height * (next.h / previous.h);
  // The kept part in the frame's own box, mirrored as the frame is drawn.
  let x = width * ((next.x - previous.x) / previous.w);
  let y = height * ((next.y - previous.y) / previous.h);
  if (frame.flipH) x = width - (x + keptWidth);
  if (frame.flipV) y = height - (y + keptHeight);
  // Its centre from the frame's centre, turned as the frame is.
  const dx = x + keptWidth / 2 - width / 2;
  const dy = y + keptHeight / 2 - height / 2;
  const angle = (((frame.rotation ?? 0) % 360) * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const centerX = left + width / 2 + dx * cos - dy * sin;
  const centerY = top + height / 2 + dx * sin + dy * cos;
  return {
    width: Math.max(1, Math.round(keptWidth)),
    height: Math.max(1, Math.round(keptHeight)),
    left: Math.round(centerX - keptWidth / 2),
    top: Math.round(centerY - keptHeight / 2),
  };
}

/**
 * Presentation-owned adapter for Picture Format → Change Picture → From File.
 * It deliberately shares IMG01's preflight/upload/compensation primitives but
 * owns a different live-target and transaction adapter: replacing a picture
 * must preserve every non-source field and the native source identity.
 */
export function createPresentationPictureReplaceCommandAdapter<
  TBlock extends ReplaceablePresentationPictureBlock,
  TSlide extends PresentationPictureReplaceSlide<TBlock>,
>(
  accessors: PresentationPictureReplaceCommandAccessors<TBlock, TSlide>,
): PresentationPictureReplaceCommandAdapter<TBlock, TSlide> {
  let pendingTarget: {
    slideId: string;
    blockId: string;
    block: TBlock;
  } | null = null;
  let busy = false;

  const target = () => {
    const resolved = selectedPictureTarget<TBlock, TSlide>(
      accessors.getSlides(),
      accessors.getActiveIndex(),
      accessors.getSelectedPictureId(),
    );
    return resolved?.block.source?.pictureSourceDisposition === "preserve-only"
      ? null
      : resolved;
  };

  // The picture the file was chosen for, unchanged, wherever its slide now
  // stands -- not "the picture selected now": clicking the slide's words
  // while a content placeholder's picture uploaded dropped the chosen file
  
  const resolvePendingTarget = () => {
    if (!pendingTarget) return null;
    const slides = accessors.getSlides();
    const slideIndex = slides.findIndex((slide) => slide.id === pendingTarget!.slideId);
    const slide = slideIndex >= 0 ? slides[slideIndex] : undefined;
    const block = slide?.blocks.find((candidate) => candidate.id === pendingTarget!.blockId);
    if (!slide || block !== pendingTarget.block) return null;
    return { slideIndex, slide, block: block as TBlock };
  };

  const canExecute = (): boolean => !busy && target() !== null;

  const execute = (): boolean => {
    const resolved = target();
    if (busy || !resolved) return false;
    pendingTarget = {
      slideId: resolved.slide.id,
      blockId: resolved.block.id,
      block: resolved.block,
    };
    if (accessors.openFilePicker()) return true;
    pendingTarget = null;
    return false;
  };

  const cancelPicker = (): void => {
    if (!busy) pendingTarget = null;
  };

  const replaceFile = async (file: File): Promise<boolean> => {
    if (busy) throw new PresentationPictureReplaceError("busy");
    if (!pendingTarget) throw new PresentationPictureReplaceError("no-target");
    busy = true;
    let assetId: string | null = null;
    try {
      const prepared = await accessors.prepareFile(file);
      if (!resolvePendingTarget()) {
        throw new PresentationPictureReplaceError("stale-target");
      }

      // A picture converted for the deck (HEIC, WebP...) is stored as the
      // file it was converted into.
      const uploadedId = await accessors.uploadFile(prepared.uploadFile ?? file);
      assetId = typeof uploadedId === "string" ? uploadedId.trim() : "";
      if (!assetId || assetId.length > 200) {
        if (assetId) {
          await compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationPictureReplaceError("invalid-asset-id"),
            (cause) =>
              new PresentationPictureReplaceError("compensation-failed", cause),
          );
        }
        throw new PresentationPictureReplaceError("invalid-asset-id");
      }

      const resolvedTarget = resolvePendingTarget();
      if (!resolvedTarget) {
        return await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationPictureReplaceError("stale-target"),
          (cause) =>
            new PresentationPictureReplaceError("compensation-failed", cause),
        );
      }
      const target = resolvedTarget;
      const sourceReplaced = replaceImageSource(
        target.block as unknown as Parameters<typeof replaceImageSource>[0],
        prepared.src,
      ) as unknown as Record<string, unknown>;
      // Change Graphic on an icon: the new file decides what the object is.
      // An SVG stays an Office Graphic with ITS compatibility preview (not
      // the old icon's); a PNG or JPEG makes it a picture, as Office does --
      // kept as an SVG graphic, the save refused the whole deck ("mimeType
      // must be image/svg+xml") while the slide already showed the PNG
      
      const {
        officeGraphicKind: _graphicKind,
        graphicFallbackSrc: _graphicFallback,
        graphicFill: _graphicFill,
        graphicOutline: _graphicOutline,
        ...pictureFields
      } = sourceReplaced;
      void _graphicKind;
      void _graphicFallback;
      void _graphicFill;
      void _graphicOutline;
      const graphicFields =
        prepared.officeGraphicKind === "svg" && prepared.graphicFallbackSrc
          ? {
              officeGraphicKind: "svg" as const,
              graphicFallbackSrc: prepared.graphicFallbackSrc,
              ...(sourceReplaced.officeGraphicKind === "svg" && sourceReplaced.graphicFill !== undefined
                ? { graphicFill: sourceReplaced.graphicFill }
                : {}),
              ...(sourceReplaced.officeGraphicKind === "svg" && sourceReplaced.graphicOutline !== undefined
                ? { graphicOutline: sourceReplaced.graphicOutline }
                : {}),
            }
          : {};
      const replacement = Object.freeze({
        ...pictureFields,
        ...graphicFields,
        sourceKind: "asset" as const,
        assetId,
        mimeType: prepared.mimeType,
        pixelWidth: prepared.pixelWidth,
        pixelHeight: prepared.pixelHeight,
        sizeBytes: prepared.sizeBytes,
        hasAlpha: prepared.hasAlpha,
        exifOrientation: prepared.exifOrientation,
      }) as unknown as TBlock;
      const nextSlides = accessors.getSlides().map((slide, slideIndex) =>
        slideIndex === target.slideIndex
          ? ({
              ...slide,
              blocks: slide.blocks.map((block) =>
                block === target.block ? replacement : block,
              ),
            } as TSlide)
          : slide,
      );
      try {
        accessors.commitSlides(nextSlides);
      } catch (error) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationPictureReplaceError("commit-failed", error),
          (cause) =>
            new PresentationPictureReplaceError("compensation-failed", cause),
        );
      }
      return true;
    } finally {
      pendingTarget = null;
      busy = false;
    }
  };

  return Object.freeze({
    featureId: REPLACE_PICTURE_FILE_FEATURE_ID,
    action: REPLACE_PICTURE_FILE_ACTION,
    handlerRef: REPLACE_PICTURE_FILE_HANDLER_REF,
    canExecute,
    execute,
    cancelPicker,
    replaceFile,
    isBusy: () => busy,
  });
}

/**
 * Presentation-owned adapter for the complete local-file insertion command.
 * Picker opening is synchronous; validation/upload/commit is single-flight.
 * Every post-upload rejection deletes the authenticated DataItem before the
 * rejection is surfaced, so a failed command never claims success or leaves an
 * orphaned upload.
 */
export function createPresentationImageFileCommandAdapter<
  T extends PresentationImageInsertSlide,
>(
  accessors: PresentationImageFileCommandAccessors<T>,
): PresentationImageFileCommandAdapter<T> {
  const maxBlocks = accessors.maxBlocksPerSlide ?? 200;
  let pendingTargetSlideId: string | null = null;
  let busy = false;

  const execute = (): boolean => {
    if (busy) return false;
    const slides = accessors.getSlides();
    const activeIndex = accessors.getActiveIndex();
    const target = slides[activeIndex];
    if (!target || target.blocks.length >= maxBlocks) return false;
    pendingTargetSlideId = target.id;
    if (accessors.openFilePicker()) return true;
    pendingTargetSlideId = null;
    return false;
  };

  const cancelPicker = (): void => {
    if (!busy) pendingTargetSlideId = null;
  };

  const insertFile = async (
    file: File,
    metadata: PresentationImageInsertMetadata = {},
  ): Promise<boolean> => {
    if (busy) throw new PresentationImageInsertError("busy");
    const targetSlideId = pendingTargetSlideId;
    pendingTargetSlideId = null;
    if (!targetSlideId) throw new PresentationImageInsertError("no-target");
    busy = true;
    let assetId: string | null = null;
    try {
      const prepared = await accessors.prepareFile(file);
      let slides = accessors.getSlides();
      let targetIndex = targetSlideIndex(
        slides,
        accessors.getActiveIndex(),
        targetSlideId,
      );
      if (targetIndex < 0) {
        throw new PresentationImageInsertError("stale-target");
      }
      if (slides[targetIndex].blocks.length >= maxBlocks) {
        throw new PresentationImageInsertError("slide-full");
      }

      // A picture converted for the deck (HEIC, WebP...) is stored as the
      // file it was converted into.
      const uploadedId = await accessors.uploadFile(prepared.uploadFile ?? file);
      assetId = typeof uploadedId === "string" ? uploadedId.trim() : "";
      if (!assetId || assetId.length > 200) {
        if (assetId) {
          await compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationImageInsertError("invalid-asset-id"),
          );
        }
        throw new PresentationImageInsertError("invalid-asset-id");
      }

      slides = accessors.getSlides();
      targetIndex = targetSlideIndex(
        slides,
        accessors.getActiveIndex(),
        targetSlideId,
      );
      if (targetIndex < 0) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationImageInsertError("stale-target"),
        );
      }
      const target = slides[targetIndex];
      if (target.blocks.length >= maxBlocks) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationImageInsertError("slide-full"),
        );
      }

      const blockId = await (async (): Promise<string> => {
        try {
          return accessors.nextId();
        } catch (error) {
          return compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationImageInsertError("commit-failed", error),
          );
        }
      })();
      if (
        !blockId ||
        slides.some(
          (slide) =>
            slide.id === blockId ||
            slide.blocks.some((block) => block.id === blockId),
        )
      ) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationImageInsertError("identity-collision"),
        );
      }
      let block: AssetBackedPresentationImageBlock;
      try {
        block = assetBackedPresentationImageBlock(
          blockId,
          assetId,
          prepared,
          metadata,
          accessors.getSlideBounds(target),
        );
      } catch (error) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationImageInsertError("commit-failed", error),
        );
      }
      const nextSlides = slides.map((slide, index) =>
        index === targetIndex
          ? ({ ...slide, blocks: [...slide.blocks, block] } as T)
          : slide,
      );
      try {
        accessors.commitSlides(nextSlides);
      } catch (error) {
        await compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationImageInsertError("commit-failed", error),
        );
      }
      // Selected where the user still is; on a slide they have left it is
      // added there without taking them back.
      if (accessors.getActiveIndex() === targetIndex) accessors.selectInsertedImage(blockId, targetIndex);
      return true;
    } finally {
      busy = false;
    }
  };

  const insertFileAtCurrent = async (
    file: File,
    metadata: PresentationImageInsertMetadata = {},
  ): Promise<boolean> => {
    if (busy) throw new PresentationImageInsertError("busy");
    const slides = accessors.getSlides();
    const activeIndex = accessors.getActiveIndex();
    const target = slides[activeIndex];
    if (!target) throw new PresentationImageInsertError("no-target");
    if (target.blocks.length >= maxBlocks) {
      throw new PresentationImageInsertError("slide-full");
    }
    pendingTargetSlideId = target.id;
    return insertFile(file, metadata);
  };

  const insertOfficeGraphicAtCurrent = async (
    file: File,
    metadata: PresentationImageInsertMetadata = {},
  ): Promise<boolean> => {
    if (busy) throw new PresentationImageInsertError("busy");
    const startSlides = accessors.getSlides();
    const startIndex = accessors.getActiveIndex();
    const targetSlideId = startSlides[startIndex]?.id;
    if (!targetSlideId) throw new PresentationImageInsertError("no-target");
    busy = true;
    try {
      const prepared = await accessors.prepareFile(file);
      const slides = accessors.getSlides();
      const targetIndex = targetSlideIndex(slides, accessors.getActiveIndex(), targetSlideId);
      if (targetIndex < 0) throw new PresentationImageInsertError("stale-target");
      const target = slides[targetIndex];
      if (target.blocks.length >= maxBlocks) throw new PresentationImageInsertError("slide-full");
      const blockId = accessors.nextId();
      if (
        !blockId ||
        slides.some((slide) => slide.id === blockId || slide.blocks.some((block) => block.id === blockId))
      ) {
        throw new PresentationImageInsertError("identity-collision");
      }
      const block = inlinePresentationOfficeGraphicBlock(
        blockId,
        prepared,
        metadata,
        accessors.getSlideBounds(target),
      );
      try {
        accessors.commitSlides(
          slides.map((slide, index) =>
            index === targetIndex ? ({ ...slide, blocks: [...slide.blocks, block] } as T) : slide,
          ),
        );
      } catch (error) {
        throw new PresentationImageInsertError("commit-failed", error);
      }
      if (accessors.getActiveIndex() === targetIndex) accessors.selectInsertedImage(blockId, targetIndex);
      return true;
    } finally {
      busy = false;
    }
  };

  return Object.freeze({
    featureId: INSERT_IMAGE_FILE_COMMAND_ID,
    action: INSERT_IMAGE_FILE_ACTION,
    handlerRef: INSERT_IMAGE_FILE_HANDLER_REF,
    execute,
    cancelPicker,
    insertFile,
    insertFileAtCurrent,
    insertOfficeGraphicAtCurrent,
    isBusy: () => busy,
  });
}
