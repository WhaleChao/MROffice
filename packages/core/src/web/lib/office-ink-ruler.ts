/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's Ruler (Office's "Draw straight lines or align things with
 * the ruler"): a straightedge laid over the page. A stroke that starts
 * along either of its long edges is drawn along that edge, straight, for
 * as long as the pen stays near it; a stroke elsewhere is ordinary ink. The
 * ruler is moved by dragging it, turned by the wheel, two fingers or the
 * arrow keys, in one-degree steps (Office: a degree at a time, the angle
 * shown while it turns).
 *
 * Geometry only, host-neutral: the ruler lives in a surface's own pixels.
 */

export interface InkRulerState {
  /** The ruler's centre, in the surface's pixels. */
  readonly x: number;
  readonly y: number;
  /** Degrees, clockwise from horizontal, 0..359. */
  readonly angle: number;
}

export interface InkRulerPoint {
  readonly x: number;
  readonly y: number;
}

/** The ruler's width across (px): Office's is about a finger wide. */
export const INK_RULER_THICKNESS_PX = 56;
/** How near an edge the pen must start for the stroke to follow it (px). */
export const INK_RULER_CAPTURE_PX = 22;
/** The pen may wander this far from the edge before the stroke leaves it (px). */
export const INK_RULER_RELEASE_PX = 48;

export function normalizeInkRulerAngle(angle: number): number {
  if (!Number.isFinite(angle)) return 0;
  const turned = Math.round(angle) % 360;
  return turned < 0 ? turned + 360 : turned;
}

/** The ruler's direction (unit) and normal (unit, pointing to its "far" side). */
export function inkRulerAxes(ruler: InkRulerState): { ux: number; uy: number; nx: number; ny: number } {
  const radians = (ruler.angle * Math.PI) / 180;
  const ux = Math.cos(radians);
  const uy = Math.sin(radians);
  return { ux, uy, nx: -uy, ny: ux };
}

/** Which long edge of the ruler a point is near: the near side (-1), the far
 * side (+1), or none (0), within `within` px of the edge's line. */
export function inkRulerEdgeNear(ruler: InkRulerState, point: InkRulerPoint, within = INK_RULER_CAPTURE_PX): -1 | 0 | 1 {
  const { nx, ny } = inkRulerAxes(ruler);
  const offset = (point.x - ruler.x) * nx + (point.y - ruler.y) * ny;
  const half = INK_RULER_THICKNESS_PX / 2;
  if (Math.abs(offset - half) <= within) return 1;
  if (Math.abs(offset + half) <= within) return -1;
  return 0;
}

/** Whether a point lies on the ruler's body (to drag it). */
export function inkRulerContains(ruler: InkRulerState, point: InkRulerPoint, length: number): boolean {
  const { ux, uy, nx, ny } = inkRulerAxes(ruler);
  const along = (point.x - ruler.x) * ux + (point.y - ruler.y) * uy;
  const across = (point.x - ruler.x) * nx + (point.y - ruler.y) * ny;
  return Math.abs(along) <= length / 2 && Math.abs(across) <= INK_RULER_THICKNESS_PX / 2;
}

/** The point projected onto the ruler's `edge` (-1 near, +1 far). */
export function inkRulerSnap(ruler: InkRulerState, edge: -1 | 1, point: InkRulerPoint): InkRulerPoint {
  const { ux, uy, nx, ny } = inkRulerAxes(ruler);
  const half = (INK_RULER_THICKNESS_PX / 2) * edge;
  const along = (point.x - ruler.x) * ux + (point.y - ruler.y) * uy;
  return { x: ruler.x + ux * along + nx * half, y: ruler.y + uy * along + ny * half };
}

/** How far a point is from the ruler's `edge` line (px). */
export function inkRulerDistanceToEdge(ruler: InkRulerState, edge: -1 | 1, point: InkRulerPoint): number {
  const { nx, ny } = inkRulerAxes(ruler);
  const offset = (point.x - ruler.x) * nx + (point.y - ruler.y) * ny;
  return Math.abs(offset - (INK_RULER_THICKNESS_PX / 2) * edge);
}

/**
 * A stroke's guide while it runs: the edge it started on, and whether the
 * pen is still following it. Once the pen leaves the edge the rest of the
 * stroke is ordinary ink (Office: lift the pen off the ruler and the line
 * ends), so a stroke never jumps back onto the edge.
 */
export class InkRulerGuide {
  private edge: -1 | 1;
  private following = true;
  constructor(private readonly ruler: InkRulerState, first: InkRulerPoint) {
    this.edge = inkRulerEdgeNear(ruler, first) || 1;
  }

  /** Where the sample goes: on the edge while the pen follows it, as is after. */
  place(point: InkRulerPoint): InkRulerPoint {
    if (!this.following) return point;
    if (inkRulerDistanceToEdge(this.ruler, this.edge, point) > INK_RULER_RELEASE_PX) {
      this.following = false;
      return point;
    }
    return inkRulerSnap(this.ruler, this.edge, point);
  }
}

/** The guide for a stroke starting at `first`, or null when the ruler is not there. */
export function inkRulerGuideFor(ruler: InkRulerState | null | undefined, first: InkRulerPoint): InkRulerGuide | null {
  if (!ruler || inkRulerEdgeNear(ruler, first) === 0) return null;
  return new InkRulerGuide(ruler, first);
}

/** The angle a two-finger turn or a wheel asks for, in whole degrees. */
export function inkRulerTurned(ruler: InkRulerState, byDegrees: number): InkRulerState {
  return { ...ruler, angle: normalizeInkRulerAngle(ruler.angle + byDegrees) };
}

/**
 * The ruler turned by `byDegrees` about `pivot` (surface pixels) rather than
 * its centre: the point under the pointer stays under it (Office: "The ruler
 * pivots where the mouse pointer is pointing"). Turned about its centre, a
 * ruler turned by one end swings that end away from the pointer, and the
 * next notch of the wheel lands on the page instead.
 */
export function inkRulerTurnedAbout(ruler: InkRulerState, byDegrees: number, pivot: InkRulerPoint): InkRulerState {
  const turned = inkRulerTurned(ruler, byDegrees);
  // The turn the whole-degree angle actually took.
  const applied = (((turned.angle - ruler.angle) % 360) + 540) % 360 - 180;
  if (applied === 0 || !Number.isFinite(pivot.x) || !Number.isFinite(pivot.y)) return turned;
  const radians = (applied * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = ruler.x - pivot.x;
  const dy = ruler.y - pivot.y;
  return { ...turned, x: pivot.x + dx * cos - dy * sin, y: pivot.y + dx * sin + dy * cos };
}

export function inkRulerMoved(ruler: InkRulerState, dx: number, dy: number): InkRulerState {
  return { ...ruler, x: ruler.x + dx, y: ruler.y + dy };
}
