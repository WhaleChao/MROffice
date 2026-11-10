/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { blockIsTitle } from "@mrnote/core/web/lib/presentation-block-title.ts";

export type AutofitStep = {
  /** `fontScale` as a share (0.925 is `fontScale="92500"`). */
  readonly fontScale: number;
  /** `lnSpcReduction` as a share (0.2 is `lnSpcReduction="20000"`). */
  readonly lineSpacingReduction: number;
};

const step = (fontScale: number, lineSpacingReduction: number): AutofitStep =>
  Object.freeze({ fontScale, lineSpacingReduction });

/** A body's steps, from fitting as typed to the 25% floor. */
export const POWERPOINT_BODY_AUTOFIT_STEPS: readonly AutofitStep[] = Object.freeze([
  step(1, 0),
  step(1, 0.1),
  step(0.925, 0.1),
  step(0.925, 0.2),
  step(0.85, 0.2),
  step(0.775, 0.2),
  step(0.7, 0.2),
  step(0.625, 0.2),
  step(0.55, 0.2),
  step(0.475, 0.2),
  step(0.4, 0.2),
  step(0.325, 0.2),
  step(0.25, 0.2),
]);

/** A title's (`type="title"` / `"ctrTitle"`): full size, else 90%. */
export const POWERPOINT_TITLE_AUTOFIT_STEPS: readonly AutofitStep[] = Object.freeze([
  step(1, 0),
  step(0.9, 0),
]);

/**
 * Whether the block is a title placeholder, which PowerPoint shrinks one step
 * only: the editor's title keys (`blockIsTitle`: title, section title) and
 * the package's own centred title.
 */
export function autofitBlockIsTitle(block: unknown): boolean {
  if (typeof block !== "object" || block === null) return false;
  return blockIsTitle(block) || (block as { placeholder?: unknown }).placeholder === "ctrTitle";
}

export function autofitStepsFor(block: unknown): readonly AutofitStep[] {
  return autofitBlockIsTitle(block) ? POWERPOINT_TITLE_AUTOFIT_STEPS : POWERPOINT_BODY_AUTOFIT_STEPS;
}

/**
 * The step a free scale stands for: the first whose words are no larger,
 * so an estimate lands where PowerPoint would put the same words.
 */
export function autofitStepAtOrBelow(steps: readonly AutofitStep[], fontScale: number): AutofitStep {
  if (!(fontScale < 1)) return steps[0]!;
  for (const candidate of steps) {
    if (candidate.fontScale < 1 && candidate.fontScale <= fontScale + 1e-9) return candidate;
  }
  return steps[steps.length - 1]!;
}

/** The first step at which `fits` holds, else the last (PowerPoint's floor). */
export function firstFittingAutofitStep(
  steps: readonly AutofitStep[],
  fits: (candidate: AutofitStep, index: number) => boolean,
): AutofitStep {
  for (let index = 0; index < steps.length; index += 1) {
    if (fits(steps[index]!, index)) return steps[index]!;
  }
  return steps[steps.length - 1]!;
}
