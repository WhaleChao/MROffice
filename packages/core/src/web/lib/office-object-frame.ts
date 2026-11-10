/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An Office object's frame: its box, its turn about the box's centre and its
 * mirror across the box's own axes, as Word, Excel and PowerPoint draw a
 * shape, a picture or an ink stroke (DrawingML a:xfrm: rot, flipH, flipV;
 * the mirror applies inside the box, then the box turns clockwise about its
 * centre).
 *
 * One answer for every editor and every reader of a placed object — what it
 * paints, where a pointer hits it, what an eraser or a lasso meets, what an
 * export writes — so a turned or mirrored object is never met where it was
 * before the turn. Units are the caller's (EMU, pixels, points): the frame
 * and the points share them.
 */

export interface OfficeFramePoint {
  readonly x: number;
  readonly y: number;
}

export interface OfficeObjectFrame {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** Degrees clockwise; anything not a finite number is no turn. */
  readonly rotation?: unknown;
  readonly flipH?: unknown;
  readonly flipV?: unknown;
}

/** The frame's turn in degrees, (-180, 180]; 0 for a missing or invalid one. */
export function officeFrameRotation(rotation: unknown): number {
  const value = Number(rotation ?? 0);
  if (!Number.isFinite(value)) return 0;
  let turn = value % 360;
  if (turn > 180) turn -= 360;
  if (turn <= -180) turn += 360;
  return turn === 0 ? 0 : turn;
}

/**
 * Where the point at fractions (`fx`, `fy`) of the frame's box (0..1 from
 * its top-left, as stored, before the mirror) shows once mirrored and turned.
 * `mirrored: false` for a point already mirrored by its geometry (an outline
 * drawn with the flips folded in).
 */
export function officeFramePlace(
  frame: OfficeObjectFrame,
  fx: number,
  fy: number,
  options: { readonly mirrored?: boolean } = {},
): OfficeFramePoint {
  const mirror = options.mirrored !== false;
  const x = mirror && frame.flipH === true ? 1 - fx : fx;
  const y = mirror && frame.flipV === true ? 1 - fy : fy;
  const dx = (x - 0.5) * frame.width;
  const dy = (y - 0.5) * frame.height;
  const cx = frame.left + frame.width / 2;
  const cy = frame.top + frame.height / 2;
  const turn = officeFrameRotation(frame.rotation);
  if (turn === 0) return { x: cx + dx, y: cy + dy };
  const radians = (turn * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/** The fractions of the box (as stored, before the mirror) a shown point is at: the inverse of `officeFramePlace`. */
export function officeFrameFractions(frame: OfficeObjectFrame, point: OfficeFramePoint): OfficeFramePoint {
  const cx = frame.left + frame.width / 2;
  const cy = frame.top + frame.height / 2;
  const turn = officeFrameRotation(frame.rotation);
  const radians = (-turn * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const ux = point.x - cx;
  const uy = point.y - cy;
  const dx = turn === 0 ? ux : ux * cos - uy * sin;
  const dy = turn === 0 ? uy : ux * sin + uy * cos;
  const x = frame.width !== 0 ? dx / frame.width + 0.5 : 0.5;
  const y = frame.height !== 0 ? dy / frame.height + 0.5 : 0.5;
  return { x: frame.flipH === true ? 1 - x : x, y: frame.flipV === true ? 1 - y : y };
}

/** Whether a shown point falls on the frame's turned box (its edge included). */
export function officeFrameContains(frame: OfficeObjectFrame, point: OfficeFramePoint): boolean {
  if (!(frame.width > 0) || !(frame.height > 0)) return false;
  const at = officeFrameFractions(frame, point);
  return at.x >= 0 && at.x <= 1 && at.y >= 0 && at.y <= 1;
}

/** The turned box's corners, clockwise from the box's own top-left. */
export function officeFrameCorners(frame: OfficeObjectFrame): OfficeFramePoint[] {
  const box = { ...frame, flipH: false, flipV: false };
  return [
    officeFramePlace(box, 0, 0),
    officeFramePlace(box, 1, 0),
    officeFramePlace(box, 1, 1),
    officeFramePlace(box, 0, 1),
  ];
}

/** The upright rectangle the turned box covers. */
export function officeFrameBounds(frame: OfficeObjectFrame): {
  left: number;
  top: number;
  width: number;
  height: number;
} {
  if (officeFrameRotation(frame.rotation) === 0) {
    return { left: frame.left, top: frame.top, width: frame.width, height: frame.height };
  }
  const corners = officeFrameCorners(frame);
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  return { left, top, width: Math.max(...xs) - left, height: Math.max(...ys) - top };
}

/** `point` turned `degrees` clockwise (y down) about `pivot`: a selection
 * turned as one, a member carried round its group's centre. */
export function officeTurnAround(point: OfficeFramePoint, pivot: OfficeFramePoint, degrees: number): OfficeFramePoint {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = point.x - pivot.x;
  const dy = point.y - pivot.y;
  return { x: pivot.x + dx * cos - dy * sin, y: pivot.y + dx * sin + dy * cos };
}

/** A movement on the page read along the frame's own axes: the turn undone,
 * then the mirror, so a handle moves its own edge however the object is
 * turned or mirrored. */
export function officeFrameAxesDelta(
  frame: Pick<OfficeObjectFrame, "rotation" | "flipH" | "flipV">,
  dx: number,
  dy: number,
): OfficeFramePoint {
  const radians = (officeFrameRotation(frame.rotation) * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return {
    x: (dx * cos + dy * sin) * (frame.flipH === true ? -1 : 1),
    y: (dy * cos - dx * sin) * (frame.flipV === true ? -1 : 1),
  };
}
