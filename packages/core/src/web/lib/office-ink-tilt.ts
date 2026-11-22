/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A pencil laid on its side shades, as Apple Notes' and Procreate's pencils
 * do: tilted past about 45 degrees from upright, its stroke broadens -- up to
 * four times the pencil's width when it lies nearly flat -- and the pencil's
 * grain (lib/document-ink pencil filter) makes it a shading rather than a
 * line. Held as one writes, it draws its own width.
 *
 * The tilt is read once, as the pencil touches down (Safari 18.2's
 * `PointerEvent.altitudeAngle`, or `tiltX`/`tiltY` everywhere a pen reports
 * them), and the stroke keeps that width from its first sample to its lift:
 * the running stroke, the wet one and the landed one are one width, and the
 * stored stroke is an ordinary pencil stroke of that width (no new
 * attribute, nothing a file or another editor must learn).
 */

export interface InkTiltSample {
  readonly pointerType: string;
  readonly altitudeAngle?: number;
  readonly tiltX?: number;
  readonly tiltY?: number;
}

/** The angle from the glass (radians) below which the pencil starts to shade. */
const SHADE_FROM_ALTITUDE = Math.PI / 4;
/** And at or below which it shades at its widest. */
const SHADE_FULL_ALTITUDE = Math.PI / 10;
/** The widest shading, as a multiple of the pencil's width. */
export const INK_TILT_SHADE_MAX = 4;
/** A stroke's outline may be no wider than this (the documents' own bound). */
const MAX_OUTLINE_WIDTH_PT = 24;

/** The pen's angle from the glass (radians, pi/2 upright), or null when it says nothing. */
export function inkPenAltitude(sample: InkTiltSample): number | null {
  if (sample.pointerType !== "pen") return null;
  const altitude = sample.altitudeAngle;
  if (typeof altitude === "number" && Number.isFinite(altitude) && altitude > 0 && altitude <= Math.PI / 2 + 1e-6) {
    // Browsers without a tilting pen report upright (pi/2): no shading.
    return altitude;
  }
  const tx = sample.tiltX;
  const ty = sample.tiltY;
  if (typeof tx !== "number" || typeof ty !== "number" || (!tx && !ty)) return null;
  // tiltX/tiltY are the pen's lean in the XZ and YZ planes (degrees).
  const tanX = Math.tan((Math.max(-89, Math.min(89, tx)) * Math.PI) / 180);
  const tanY = Math.tan((Math.max(-89, Math.min(89, ty)) * Math.PI) / 180);
  return Math.atan(1 / Math.max(1e-6, Math.hypot(tanX, tanY)));
}

/** How many times its width a pencil held at `sample`'s tilt draws (1: upright, writing). */
export function inkTiltShadeFactor(sample: InkTiltSample): number {
  const altitude = inkPenAltitude(sample);
  if (altitude === null || altitude >= SHADE_FROM_ALTITUDE) return 1;
  const t = Math.min(1, (SHADE_FROM_ALTITUDE - altitude) / (SHADE_FROM_ALTITUDE - SHADE_FULL_ALTITUDE));
  // Eased in: a pencil just past the line barely broadens.
  const eased = t * t * (3 - 2 * t);
  return 1 + (INK_TILT_SHADE_MAX - 1) * eased;
}

/**
 * The pen a pencil stroke starting at `sample` draws with: the same pencil,
 * as wide as its tilt makes it (within the outline's bound). Any other tool,
 * or a pencil held upright, is the pen as it is.
 */
export function inkPresetForTilt<P extends { readonly tool: string; readonly widthPt: number }>(
  preset: P,
  sample: InkTiltSample,
  /** The widest outline the editor stores (points); the pen's pressure may widen a stroke by 1.4 on top. */
  maxOutlinePt: number = MAX_OUTLINE_WIDTH_PT,
): P {
  if (preset.tool !== "pencil") return preset;
  const factor = inkTiltShadeFactor(sample);
  if (factor <= 1.02) return preset;
  const widest = Math.max(preset.widthPt, maxOutlinePt / 1.4);
  const widthPt = Math.round(Math.min(widest, preset.widthPt * factor) * 100) / 100;
  return { ...preset, widthPt };
}
