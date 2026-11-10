/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Adjustment handles -- the yellow ones -- for every DrawingML preset, shared
 * by Docs, the deck and the worksheet.
 *
 * Office semantics (ECMA-376 §20.1.9.3 `a:ahLst`, §20.1.9.14 `a:ahPolar`,
 * §20.1.9.15 `a:ahXY`; Microsoft Support "Change the shape of a shape using
 * the adjustment handle"): a preset declares its handles next to its guides.
 * A handle is DRAWN at its `a:pos`, two guide references evaluated with the
 * shape's current adjust values. Dragging it writes the guide named by
 * `gdRefX` / `gdRefY` (an XY handle) or `gdRefR` / `gdRefAng` (a polar one),
 * held inside `[min, max]` -- bounds that may themselves be guides (a callout's
 * `maxAdj1` depends on the other handle). An omitted pair means the handle
 * does not move on that axis. The value lands in `a:prstGeom/a:avLst` as
 * `<a:gd name="adj" fmla="val N"/>`, which every host already writes and reads
 * back (`adjustments`), and the canvas redraws through the SAME catalog
 * program with the value substituted.
 *
 * The drag is solved numerically against the program itself rather than by a
 * per-preset inverse, so all 120 presets with handles (243 handles) behave
 * the same way through one code path: for each referenced guide, the value in
 * `[min, max]` whose handle position lands closest to the pointer.
 */

import { evaluateDrawingmlPresetGuides, drawingmlDefaultPresetAdjustNames } from "./drawingml-default-preset-program";
import {
  OFFICE_PRESET_ADJUST_HANDLES,
  OFFICE_PRESET_TEXT_RECTS,
  type OfficePresetAdjustHandleSpec,
} from "./office-preset-adjust-handles.generated";
import { DEFAULT_PRESET_PROGRAMS } from "./presentation-default-preset-programs.generated";

export type { OfficePresetAdjustHandleSpec };

/** One handle, in the shape's own unrotated, unflipped box (CSS pixels). */
export interface OfficeShapeAdjustHandle {
  readonly index: number;
  readonly kind: "xy" | "polar";
  readonly x: number;
  readonly y: number;
}

const EMU_PER_PX = 9525;

function programFor(drawingPreset: string): unknown | null {
  const entry = Object.prototype.hasOwnProperty.call(DEFAULT_PRESET_PROGRAMS, drawingPreset)
    ? (DEFAULT_PRESET_PROGRAMS[drawingPreset] as { program?: unknown } | undefined)
    : undefined;
  return entry?.program ?? null;
}

function handleSpecs(drawingPreset: string): readonly OfficePresetAdjustHandleSpec[] {
  return Object.prototype.hasOwnProperty.call(OFFICE_PRESET_ADJUST_HANDLES, drawingPreset)
    ? OFFICE_PRESET_ADJUST_HANDLES[drawingPreset]
    : [];
}

/** Whether the preset has any adjustment handle at all. */
export function officeShapeHasAdjustHandles(drawingPreset: string | null | undefined): boolean {
  return typeof drawingPreset === "string" && handleSpecs(drawingPreset).length > 0;
}

/** The preset's own `a:avLst` defaults, by guide name. */
export function officeShapeAdjustDefaults(drawingPreset: string): Readonly<Record<string, number>> {
  const program = programFor(drawingPreset);
  const names = program ? drawingmlDefaultPresetAdjustNames(program) : null;
  if (!program || !names || names.length === 0) return {};
  const vars = evaluateDrawingmlPresetGuides(program, 914_400, 914_400);
  if (!vars) return {};
  return Object.fromEntries(names.map((name) => [name, vars[name]]));
}

function reference(token: string | undefined, vars: Readonly<Record<string, number>>): number | null {
  if (token === undefined) return null;
  if (Object.prototype.hasOwnProperty.call(vars, token)) return vars[token];
  return /^-?\d+$/.test(token) ? Number(token) : null;
}

function evaluate(program: unknown, width: number, height: number, adjustments: Readonly<Record<string, number>>) {
  return evaluateDrawingmlPresetGuides(program, width * EMU_PER_PX, height * EMU_PER_PX, adjustments);
}

/**
 * Where each handle of one preset is drawn, for a shape `widthPx` x `heightPx`
 * with its own `adjustments` (absent names take the preset's defaults).
 */
export function officeShapeAdjustHandles(
  drawingPreset: string | null | undefined,
  widthPx: number,
  heightPx: number,
  adjustments?: Readonly<Record<string, number>> | null,
): OfficeShapeAdjustHandle[] {
  if (typeof drawingPreset !== "string" || !(widthPx > 0) || !(heightPx > 0)) return [];
  const program = programFor(drawingPreset);
  const specs = handleSpecs(drawingPreset);
  if (!program || specs.length === 0) return [];
  const vars = evaluate(program, widthPx, heightPx, adjustments ?? {});
  if (!vars) return [];
  const handles: OfficeShapeAdjustHandle[] = [];
  specs.forEach((spec, index) => {
    const x = reference(spec.pos[0], vars);
    const y = reference(spec.pos[1], vars);
    if (x === null || y === null || !Number.isFinite(x) || !Number.isFinite(y)) return;
    handles.push({ index, kind: spec.kind, x: x / EMU_PER_PX, y: y / EMU_PER_PX });
  });
  return handles;
}

/** The guide refs one handle writes, each with its bound attribute names. */
function handleAxes(spec: OfficePresetAdjustHandleSpec): { ref: string; min?: string; max?: string }[] {
  const axes: { ref: string; min?: string; max?: string }[] = [];
  if (spec.kind === "xy") {
    if (spec.gdRefX) axes.push({ ref: spec.gdRefX, min: spec.minX, max: spec.maxX });
    if (spec.gdRefY) axes.push({ ref: spec.gdRefY, min: spec.minY, max: spec.maxY });
  } else {
    if (spec.gdRefAng) axes.push({ ref: spec.gdRefAng, min: spec.minAng, max: spec.maxAng });
    if (spec.gdRefR) axes.push({ ref: spec.gdRefR, min: spec.minR, max: spec.maxR });
  }
  return axes;
}

/**
 * The adjustments after dragging handle `index` to `pointer` (box pixels, in
 * the shape's unrotated, unflipped frame). Returns every adjust name the
 * preset declares, so the written `a:avLst` is complete and two equal shapes
 * compare equal; `null` when the preset or handle does not exist.
 */
export function officeShapeAdjustDrag(
  drawingPreset: string,
  widthPx: number,
  heightPx: number,
  adjustments: Readonly<Record<string, number>> | null | undefined,
  index: number,
  pointer: { readonly x: number; readonly y: number },
): Record<string, number> | null {
  if (!(widthPx > 0) || !(heightPx > 0) || !Number.isFinite(pointer.x) || !Number.isFinite(pointer.y)) return null;
  const program = programFor(drawingPreset);
  const spec = handleSpecs(drawingPreset)[index];
  if (!program || !spec) return null;
  const names = drawingmlDefaultPresetAdjustNames(program) ?? [];
  const defaults = officeShapeAdjustDefaults(drawingPreset);
  const next: Record<string, number> = {};
  for (const name of names) {
    const current = adjustments?.[name];
    next[name] = typeof current === "number" && Number.isFinite(current) ? current : defaults[name];
  }
  const target = { x: pointer.x * EMU_PER_PX, y: pointer.y * EMU_PER_PX };
  const distance = (candidate: Record<string, number>): number => {
    const vars = evaluate(program, widthPx, heightPx, candidate);
    if (!vars) return Number.POSITIVE_INFINITY;
    const x = reference(spec.pos[0], vars);
    const y = reference(spec.pos[1], vars);
    if (x === null || y === null) return Number.POSITIVE_INFINITY;
    return Math.hypot(x - target.x, y - target.y);
  };
  const axes = handleAxes(spec).filter((axis) => names.includes(axis.ref));
  if (axes.length === 0) return null;
  // Two alternating passes settle a two-axis handle (a callout tail); a
  // single-axis handle converges in the first.
  for (let pass = 0; pass < (axes.length > 1 ? 3 : 1); pass += 1) {
    for (const axis of axes) {
      const vars = evaluate(program, widthPx, heightPx, next);
      if (!vars) return null;
      let low = reference(axis.min, vars);
      let high = reference(axis.max, vars);
      if (low === null || high === null) continue;
      if (low > high) [low, high] = [high, low];
      if (low === high) {
        next[axis.ref] = Math.round(low);
        continue;
      }
      const probe = (value: number) => distance({ ...next, [axis.ref]: value });
      const samples = 64;
      let best = low;
      let bestDistance = Number.POSITIVE_INFINITY;
      for (let step = 0; step <= samples; step += 1) {
        const value = low + ((high - low) * step) / samples;
        const d = probe(value);
        if (d < bestDistance) {
          bestDistance = d;
          best = value;
        }
      }
      // Golden-section refinement inside the best sample's neighbourhood.
      const span = (high - low) / samples;
      let a = Math.max(low, best - span);
      let b = Math.min(high, best + span);
      const ratio = (Math.sqrt(5) - 1) / 2;
      for (let iteration = 0; iteration < 96 && b - a > 0.25; iteration += 1) {
        const c = b - ratio * (b - a);
        const d = a + ratio * (b - a);
        if (probe(c) <= probe(d)) b = d;
        else a = c;
      }
      const refined = (a + b) / 2;
      next[axis.ref] = Math.round(probe(refined) <= bestDistance ? refined : best);
    }
  }
  return next;
}

/** The adjust guides a preset's handles write (`avLst` names), in order. */
export function officeShapeAdjustNames(drawingPreset: string | null | undefined): readonly string[] {
  if (typeof drawingPreset !== "string" || handleSpecs(drawingPreset).length === 0) return [];
  const program = programFor(drawingPreset);
  return program ? (drawingmlDefaultPresetAdjustNames(program) ?? []) : [];
}

/**
 * Adjustment values typed by name (the AI's Adjust Shape commands): each named
 * guide set to its value and pinned to the range its handle allows on a shape
 * of this size, as a drag pins it; the other guides kept (or the preset's
 * defaults), so the result is the whole `avLst` a drag would commit. Null when
 * a name is not one of the preset's guides or the preset has no handles.
 */
export function officeShapeAdjustmentsWithValues(
  drawingPreset: string,
  widthPx: number,
  heightPx: number,
  adjustments: Readonly<Record<string, number>> | null | undefined,
  values: Readonly<Record<string, number>>,
): Record<string, number> | null {
  if (!(widthPx > 0) || !(heightPx > 0)) return null;
  const program = programFor(drawingPreset);
  const specs = handleSpecs(drawingPreset);
  const names = officeShapeAdjustNames(drawingPreset);
  if (!program || specs.length === 0 || names.length === 0) return null;
  const defaults = officeShapeAdjustDefaults(drawingPreset);
  const next: Record<string, number> = {};
  for (const name of names) {
    const current = adjustments?.[name];
    next[name] = typeof current === "number" && Number.isFinite(current) ? current : defaults[name];
  }
  for (const [name, value] of Object.entries(values)) {
    if (!names.includes(name) || !Number.isFinite(value)) return null;
    next[name] = Math.round(value);
  }
  for (const spec of specs) {
    for (const axis of handleAxes(spec)) {
      if (!Object.prototype.hasOwnProperty.call(values, axis.ref)) continue;
      const vars = evaluate(program, widthPx, heightPx, next);
      if (!vars) return null;
      let low = reference(axis.min, vars);
      let high = reference(axis.max, vars);
      if (low === null || high === null) continue;
      if (low > high) [low, high] = [high, low];
      next[axis.ref] = Math.round(Math.min(high, Math.max(low, next[axis.ref])));
    }
  }
  return next;
}

/**
 * A client point mapped into a shape's own unrotated, unflipped box. `box` is
 * the shape's untransformed layout box in client pixels (its left/top before
 * rotation, and its size); rotation turns about the box centre, as Office's
 * `a:xfrm/@rot` does, and the flips mirror about it.
 */
export function officeShapeLocalPoint(
  client: { readonly x: number; readonly y: number },
  box: { readonly left: number; readonly top: number; readonly width: number; readonly height: number },
  rotationDeg = 0,
  flipH = false,
  flipV = false,
): { x: number; y: number } {
  const cx = box.left + box.width / 2;
  const cy = box.top + box.height / 2;
  const radians = (-(rotationDeg || 0) * Math.PI) / 180;
  const dx = client.x - cx;
  const dy = client.y - cy;
  let x = dx * Math.cos(radians) - dy * Math.sin(radians);
  let y = dx * Math.sin(radians) + dy * Math.cos(radians);
  if (flipH) x = -x;
  if (flipV) y = -y;
  return { x: x + box.width / 2, y: y + box.height / 2 };
}

/** The handle's box point back in the shape's rotated / flipped frame. */
export function officeShapeHandleDisplayPoint(
  handle: { readonly x: number; readonly y: number },
  width: number,
  height: number,
  flipH = false,
  flipV = false,
): { x: number; y: number } {
  return {
    x: flipH ? width - handle.x : handle.x,
    y: flipV ? height - handle.y : handle.y,
  };
}

/**
 * A preset's TEXT RECTANGLE (DrawingML `a:rect`, ECMA-376 §20.1.9.22) as
 * fractions of the shape's box: Word, Excel and PowerPoint lay a shape's label
 * out inside it, the body insets measured from its edges -- an ellipse's
 * inscribed rectangle, the lower part of a triangle, the box above a callout's
 * tail. Evaluated with the shape's own adjust values through the same guides
 * the outline is drawn from, so a dragged rounded corner moves the text with
 * it. `right`/`bottom` are edge POSITIONS (the deck's
 * `ShapeTextRectFractions` convention). Null for no such preset.
 */
export function officePresetTextRect(
  drawingPreset: string | null | undefined,
  widthPx: number,
  heightPx: number,
  adjustments?: Readonly<Record<string, number>> | null,
): { left: number; top: number; right: number; bottom: number } | null {
  if (typeof drawingPreset !== "string" || !(widthPx > 0) || !(heightPx > 0)) return null;
  const refs = Object.prototype.hasOwnProperty.call(OFFICE_PRESET_TEXT_RECTS, drawingPreset)
    ? OFFICE_PRESET_TEXT_RECTS[drawingPreset]
    : null;
  const program = programFor(drawingPreset);
  if (!refs || !program) return null;
  const vars = evaluate(program, widthPx, heightPx, adjustments ?? {});
  if (!vars) return null;
  const [l, t, r, b] = refs.map((token) => reference(token, vars));
  if (l === null || t === null || r === null || b === null) return null;
  const width = widthPx * EMU_PER_PX;
  const height = heightPx * EMU_PER_PX;
  const rect = { left: l / width, top: t / height, right: r / width, bottom: b / height };
  return Object.values(rect).every(Number.isFinite) && rect.right > rect.left && rect.bottom > rect.top
    ? rect
    : null;
}
