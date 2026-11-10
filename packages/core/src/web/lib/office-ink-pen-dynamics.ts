/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a pen feels, the way a note-taking app's pen settings shape a stroke
 * (Goodnotes: a pen type, Tip Sharpness, Pressure Sensitivity and Stroke
 * Stabilization). Each pen keeps its own settings in its preset, in this
 * browser, beside its colour and width.
 *
 * Nothing new is stored on a stroke: the settings shape the samples a stroke
 * is made of as the pen draws them -- the pressure each sample keeps (which
 * the renderers, the live ink and the exports already turn into width, see
 * `inkPressureWidthFactor`) and, with stabilization, where each sample lands.
 * A stroke drawn with any pen therefore saves, reopens, syncs and exports as
 * every stroke always has.
 */

export const INK_PEN_TYPES = ["fountain", "ballpoint", "brush"] as const;
export type InkPenType = (typeof INK_PEN_TYPES)[number];

/** A pen's feel, every level 0..100 (a slider's percent). */
export interface InkPenFeel {
  readonly penType?: InkPenType;
  /** How sharply a stroke's start tapers (Tip Sharpness). */
  readonly tipSharpness?: number;
  /** How far pressure moves the width; 0 draws one width (Pressure off). */
  readonly pressureSensitivity?: number;
  /** How much the line is smoothed behind the pen (Stroke Stabilization). */
  readonly stabilization?: number;
}

/** The levels a pen that never set them draws with: the pen as it always drew. */
export const INK_PEN_FEEL_DEFAULTS = {
  tipSharpness: 50,
  pressureSensitivity: 50,
  stabilization: 0,
} as const;

function level(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.round(Math.min(100, Math.max(0, value)))
    : undefined;
}

/** A stored pen's feel, kept only where the tool has one (a highlighter has none). */
export function normalizeInkPenFeel(pen: Record<string, unknown>, tool: string): InkPenFeel {
  if (tool === "highlighter") return {};
  const penType = tool === "pen" && INK_PEN_TYPES.includes(pen.penType as InkPenType)
    ? (pen.penType as InkPenType) : undefined;
  const tipSharpness = level(pen.tipSharpness);
  const pressureSensitivity = level(pen.pressureSensitivity);
  const stabilization = level(pen.stabilization);
  return {
    ...(penType ? { penType } : {}),
    ...(tipSharpness !== undefined ? { tipSharpness } : {}),
    ...(pressureSensitivity !== undefined ? { pressureSensitivity } : {}),
    ...(stabilization !== undefined ? { stabilization } : {}),
  };
}

/** A pen menu's change to one pen: its colour, width, effect, Pressure, feel. */
export interface InkPenPatch {
  color?: string;
  widthPt?: number;
  effect?: import("./document-shapes").DocumentInkEffect | null;
  /** Pressure off (true) or back on (false). */
  steady?: boolean;
  /** Pen type and levels; `penType: null` puts the pen back to a plain pen. */
  feel?: { penType?: InkPenType | null; tipSharpness?: number; pressureSensitivity?: number; stabilization?: number };
}

/** A stored pen record with a feel patch applied (normalized afterwards). */
export function withInkPenFeelPatch(pen: Record<string, unknown>, feel: InkPenPatch["feel"]): Record<string, unknown> {
  if (!feel) return pen;
  const next = { ...pen };
  if (feel.penType === null) delete next.penType;
  else if (feel.penType !== undefined) next.penType = feel.penType;
  for (const key of ["tipSharpness", "pressureSensitivity", "stabilization"] as const) {
    if (feel[key] !== undefined) next[key] = feel[key];
  }
  // Pressure Sensitivity at 0 is Pressure off, which the editors and the
  // sample line read as `steady`; any other level turns it back on.
  if (feel.pressureSensitivity !== undefined) {
    if (feel.pressureSensitivity <= 0) next.steady = true;
    else delete next.steady;
  }
  return next;
}

interface FeelPreset extends InkPenFeel {
  readonly tool: string;
  readonly steady?: true;
}

/** The pen's levels, a pen that set none drawing as before. */
export function inkPenFeelOf(preset: FeelPreset): Required<Omit<InkPenFeel, "penType">> & { penType?: InkPenType } {
  return {
    ...(preset.penType ? { penType: preset.penType } : {}),
    tipSharpness: preset.tipSharpness ?? (preset.penType ? INK_PEN_FEEL_DEFAULTS.tipSharpness : 0),
    pressureSensitivity: preset.pressureSensitivity ?? (preset.steady ? 0 : INK_PEN_FEEL_DEFAULTS.pressureSensitivity),
    stabilization: preset.stabilization ?? INK_PEN_FEEL_DEFAULTS.stabilization,
  };
}

/** How each type answers pressure: a ballpoint barely, a brush strongly. */
const TYPE_PRESSURE_GAIN: Record<InkPenType, number> = { fountain: 1, ballpoint: 0.35, brush: 1.6 };
/** A fountain nib held at 45 degrees: thinnest along it, widest across it. */
const NIB_ANGLE = Math.PI / 4;

export interface InkPenDynamics {
  /** Whether the pen draws other than one width (the live ink draws pressure). */
  readonly varies: boolean;
  /**
   * The pressure a sample keeps, its place advancing the stroke (direction,
   * length travelled); undefined where the pen draws one width.
   */
  pressure(raw: number | undefined, at: { x: number; y: number }): number | undefined;
  /** The same without advancing (the browser's predicted samples). */
  peek(raw: number | undefined, at: { x: number; y: number }): number | undefined;
  /** Where a sample lands: behind the pen by the stabilization. */
  place<P extends { x: number; y: number }>(point: P): P;
}

/**
 * One stroke's dynamics: create at touch-down, feed every sample in order.
 * `widthPx` is the pen's drawn width, the unit the tip's taper runs over.
 */
export function createInkPenDynamics(preset: FeelPreset, widthPx: number): InkPenDynamics {
  const feel = inkPenFeelOf(preset);
  const highlighter = preset.tool === "highlighter";
  const sensitivity = highlighter ? 0 : feel.pressureSensitivity / 50; // 1 = as the pen always drew
  const gain = sensitivity * (feel.penType ? TYPE_PRESSURE_GAIN[feel.penType] : 1);
  const fountain = !highlighter && feel.penType === "fountain";
  const taperPx = highlighter ? 0 : (feel.tipSharpness / 100) * Math.max(4, widthPx) * 6;
  const smoothing = Math.min(0.9, (feel.stabilization / 100) * 0.9);
  const varies = gain > 0 || fountain || taperPx > 0;
  let previous: { x: number; y: number } | null = null;
  let travelled = 0;
  let heading: number | null = null;
  let smoothed: { x: number; y: number } | null = null;

  // `raw` is the pen's own pressure (inkSamplePressure), undefined from a
  // mouse or a finger.
  const compute = (raw: number | undefined, at: { x: number; y: number }, advance: boolean): number | undefined => {
    if (!varies || (raw === undefined && !fountain && taperPx === 0)) return gain > 0 ? raw : undefined;
    let distance = travelled;
    let direction = heading;
    if (previous) {
      const dx = at.x - previous.x;
      const dy = at.y - previous.y;
      const step = Math.hypot(dx, dy);
      distance += step;
      if (step > 0.5) direction = Math.atan2(dy, dx);
    }
    if (advance) {
      previous = { x: at.x, y: at.y };
      travelled = distance;
      heading = direction;
    }
    let p = 0.5 + ((raw ?? 0.5) - 0.5) * gain;
    if (fountain && direction !== null) {
      // Across the nib the full width, along it the hairline.
      const across = Math.abs(Math.sin(direction - NIB_ANGLE));
      p = p * (0.25 + 0.75 * across);
    }
    if (taperPx > 0) p *= 0.2 + 0.8 * Math.min(1, distance / taperPx);
    return Math.min(1, Math.max(0, p));
  };

  return {
    varies,
    pressure: (raw, at) => compute(raw, at, true),
    peek: (raw, at) => compute(raw, at, false),
    place(point) {
      if (smoothing <= 0) return point;
      if (!smoothed) {
        smoothed = { x: point.x, y: point.y };
        return point;
      }
      const alpha = 1 - smoothing;
      smoothed = { x: smoothed.x + (point.x - smoothed.x) * alpha, y: smoothed.y + (point.y - smoothed.y) * alpha };
      return { ...point, x: smoothed.x, y: smoothed.y };
    },
  };
}
