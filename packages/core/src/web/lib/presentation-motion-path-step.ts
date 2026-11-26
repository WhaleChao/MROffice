/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Effect Options of a motion path on one animation step (GAP #10):
 * Reverse Path Direction and Path: Unlocked / Locked, as PowerPoint's
 * Animations > Effect Options offers them for a path.
 *
 * A step travels a path in one of two forms: a gallery path (Arcs: Arc
 * Down, Lines: Right ...), written from its measured template, or a path as
 * a file states it (`animation` "none" with `motionPathNative`, row 1078).
 * A gallery path reversed or locked is no longer its template, so the step
 * takes the second form: the same preset, its path as now written, the
 * template's smooth start / end stated as its own. Every geometry rule is in
 * `lib/presentation-motion-path-native`.
 */

import {
  MAX_EFFECT_SHARE,
  canonicalEffectSettings,
  type PresentationAnimationEffectSettings,
} from "./presentation-anim-effect-settings";
import { effectSettingsDefaults } from "./presentation-anim-effect-defaults";
import {
  createStep,
  presentationAnimationEffect,
  resolvePresentationAnimationEffectOption,
  type MotionPathNative,
  type TimelineStep,
} from "./presentation-anim-timeline";
import { effectRecipeMotionPath } from "./presentation-effect-recipes";
import { motionPathOfNativePath, translateNativeMotionPath } from "./presentation-motion-path-native";

/** The keys a path step in the file's form carries (the server's planner takes no others). */
const NATIVE_PATH_STEP_KEYS: ReadonlySet<string> = new Set([
  "id",
  "blockId",
  "animation",
  "effectOption",
  "effectSettings",
  "trigger",
  "delayMs",
  "durationMs",
  "motionPath",
  "motionPathNative",
  "triggerBlockId",
  "targetGroup",
]);

/**
 * The path a step travels as PowerPoint writes it: its own (row 1078), or
 * its gallery path's template; null for a step that travels no path.
 */
export function motionPathNativeOfStep(step: TimelineStep | null | undefined): MotionPathNative | null {
  if (!step) return null;
  if (step.motionPathNative) return step.motionPathNative;
  const effect = presentationAnimationEffect(step.animation);
  if (!effect?.motionPath) return null;
  const path = effectRecipeMotionPath(step.animation, step.effectOption);
  if (!path) return null;
  const option = resolvePresentationAnimationEffectOption(step.animation, step.effectOption);
  return {
    presetId: effect.nativePresetId,
    presetSubtype: option?.nativePresetSubtype ?? effect.nativePresetSubtype,
    path,
    ptsTypes: "",
  };
}

/** Effect Options > Path: Locked. */
export function motionPathStepLocked(step: TimelineStep | null | undefined): boolean {
  return step?.motionPathNative?.pathEditMode === "fixed";
}

/** The step travelling `native` instead, in the file's form; null when it cannot. */
function stepWithNativePath(
  step: TimelineStep,
  native: MotionPathNative,
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep | null {
  if (Object.keys(step).some((key) => !NATIVE_PATH_STEP_KEYS.has(key))) return null;
  const motionPath = motionPathOfNativePath(native.path, slideWidthEmu, slideHeightEmu);
  if (!motionPath) return null;
  let effectSettings: PresentationAnimationEffectSettings | undefined = step.effectSettings;
  if (step.animation !== "none") {
    // A gallery path's smooth start / end are its template's; as a path of
    // its own they are stated.
    const template = effectSettingsDefaults(step.animation, step.effectOption);
    const accel = step.effectSettings?.accel ?? template.accel;
    const decel = step.effectSettings?.decel ?? template.decel;
    if (accel + decel > MAX_EFFECT_SHARE) return null;
    effectSettings = canonicalEffectSettings(
      { ...step.effectSettings, accel, decel },
      effectSettingsDefaults("none", undefined),
    );
  }
  const { effectOption: _option, effectSettings: _settings, ...rest } = step;
  return {
    ...rest,
    animation: "none",
    motionPath,
    motionPathNative: native,
    ...(effectSettings ? { effectSettings } : {}),
  };
}

/**
 * Effect Options > Reverse Path Direction: the step travelling its path the
 * other way, or null. As PowerPoint does it, the path and its point types
 * stay as they are and the motion plays backwards (`reversed`, written
 * `spd="-100000"`, row 1314); a second Reverse plays it forwards again. A
 * path an earlier MRAI reversed by writing its points from end to start is
 * left so: reversing it again only adds the backwards play.
 */
export function reversedMotionPathStep(
  step: TimelineStep,
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep | null {
  const native = motionPathNativeOfStep(step);
  if (!native) return null;
  const { reversed, ...forwards } = native;
  return stepWithNativePath(step, reversed ? forwards : { ...forwards, reversed: true }, slideWidthEmu, slideHeightEmu);
}

/** Edit Points on the step's path: the step travelling the edited path, or null when nothing changed. */
export function editedMotionPathStep(
  step: TimelineStep,
  edited: { path: string; ptsTypes: string },
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep | null {
  const native = motionPathNativeOfStep(step);
  if (!native || (native.path === edited.path && native.ptsTypes === edited.ptsTypes)) return null;
  return stepWithNativePath(step, { ...native, path: edited.path, ptsTypes: edited.ptsTypes }, slideWidthEmu, slideHeightEmu);
}

/**
 * Motion Paths > Custom Path (GAP #10, step 3): a new step on `blockId`
 * travelling a drawn path, as PowerPoint writes a drawn one -- preset 0
 * subtype 0, two seconds, smooth start and end half each (truth deck
 * motion-path-roundtrip) -- or null when the path plays nothing.
 */
export function drawnMotionPathStep(
  blockId: string,
  drawn: { path: string; ptsTypes: string },
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep | null {
  const motionPath = motionPathOfNativePath(drawn.path, slideWidthEmu, slideHeightEmu);
  if (!motionPath) return null;
  const { id } = createStep(blockId, "fade");
  return {
    ...(id ? { id } : {}),
    blockId,
    animation: "none",
    trigger: "on-click",
    delayMs: 0,
    durationMs: 2000,
    motionPath,
    motionPathNative: { presetId: 0, presetSubtype: 0, path: drawn.path, ptsTypes: drawn.ptsTypes },
    effectSettings: { accel: 50_000, decel: 50_000 },
  };
}

/** Effect Options > Path: Locked / Unlocked: the step so, or null when it already is. */
export function lockedMotionPathStep(
  step: TimelineStep,
  locked: boolean,
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep | null {
  const native = motionPathNativeOfStep(step);
  if (!native || motionPathStepLocked(step) === locked) return null;
  const { pathEditMode: _mode, ...unlocked } = native;
  return stepWithNativePath(
    step,
    locked ? { ...unlocked, pathEditMode: "fixed" } : unlocked,
    slideWidthEmu,
    slideHeightEmu,
  );
}

/** A block's place as the path measures from it: its centre, in EMU. */
interface BlockCentre {
  readonly x: number;
  readonly y: number;
}

/**
 * A locked path held where it is on the slide when its object moves from
 * `from` to `to` (centres, EMU): every point moved back by the object's
 * move. The step as it was when it is not locked or nothing moved.
 */
export function heldLockedMotionPathStep(
  step: TimelineStep,
  from: BlockCentre,
  to: BlockCentre,
  slideWidthEmu: number,
  slideHeightEmu: number,
): TimelineStep {
  const native = step.motionPathNative;
  if (native?.pathEditMode !== "fixed" || !(slideWidthEmu > 0) || !(slideHeightEmu > 0)) return step;
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  if (dx === 0 && dy === 0) return step;
  const path = translateNativeMotionPath(native.path, -dx / slideWidthEmu, -dy / slideHeightEmu);
  return path ? { ...step, motionPathNative: { ...native, path } } : step;
}

interface HeldSlide {
  readonly id: string;
  readonly blocks: readonly unknown[];
  readonly animationSequence?: readonly TimelineStep[];
}

function blockCentre(block: unknown): BlockCentre | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  const number = (key: string) =>
    typeof record[key] === "number" && Number.isFinite(record[key]) ? (record[key] as number) : null;
  const left = number("left");
  const top = number("top");
  if (left === null || top === null) return null;
  return { x: left + (number("width") ?? 0) / 2, y: top + (number("height") ?? 0) / 2 };
}

function blockId(block: unknown): string | null {
  const id = block && typeof block === "object" ? (block as Record<string, unknown>).id : null;
  return typeof id === "string" ? id : null;
}

/**
 * The deck an edit makes, its locked paths held where they were on the slide
 * (Effect Options > Path: Locked): for each slide the edit changed, a step
 * whose path is locked, and whose path the edit left as it was, takes back
 * its object's move. The same array when nothing is held.
 */
export function slidesWithLockedMotionPathsHeld<T extends HeldSlide>(
  next: readonly T[],
  previous: readonly T[],
  slideSize: (slide: T) => { width: number; height: number },
): T[] {
  let changed = false;
  const before = new Map(previous.map((slide) => [slide.id, slide]));
  const out = next.map((slide) => {
    const old = before.get(slide.id);
    const sequence = slide.animationSequence;
    if (
      !old ||
      old === slide ||
      !Array.isArray(sequence) ||
      !sequence.some((step) => step.motionPathNative?.pathEditMode === "fixed")
    ) {
      return slide;
    }
    const oldSteps = new Map(
      (old.animationSequence ?? []).flatMap((step) => (step.id ? [[step.id, step] as const] : [])),
    );
    const { width, height } = slideSize(slide);
    let slideChanged = false;
    const held = sequence.map((step, index) => {
      if (step.motionPathNative?.pathEditMode !== "fixed") return step;
      // The path the edit itself set stands as it set it.
      const was = (step.id ? oldSteps.get(step.id) : undefined) ?? old.animationSequence?.[index];
      if (!was || JSON.stringify(was.motionPathNative) !== JSON.stringify(step.motionPathNative)) return step;
      const from = blockCentre(old.blocks.find((block) => blockId(block) === step.blockId));
      const to = blockCentre(slide.blocks.find((block) => blockId(block) === step.blockId));
      if (!from || !to) return step;
      const kept = heldLockedMotionPathStep(step, from, to, width, height);
      if (kept !== step) slideChanged = true;
      return kept;
    });
    if (!slideChanged) return slide;
    changed = true;
    return { ...slide, animationSequence: held };
  });
  return changed ? out : (next as T[]);
}
