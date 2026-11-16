/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How far Office's perspective camera stands from what it looks at -- a
 * shape, a picture, a SmartArt, the words -- whatever its size.
 */

const EMU_PER_PX = 9_525;

/**
 * Office's perspective camera distance at a field of view of 45 degrees, in
 * EMU: an ABSOLUTE length, whatever the size of what it looks at. Measured on
 * SmartArt (lib/office-smartart/canvas-3d: the same group at 0.45x-1.8x its
 * size and on another slide size fits 13.8M EMU within 0.11 %, Heroic
 * Extreme's 80 degrees half of it -- the distance goes as 1 / tan(fov / 2)).
 */
export const OFFICE_CAMERA_DISTANCE_EMU_AT_45 = 13_840_000;












export function officeCameraEyePx(_size: { readonly width: number; readonly height: number }, fovDegrees: number): number {
  const half = (Math.min(179, Math.max(1, fovDegrees)) * Math.PI) / 360;
  return Math.round(((OFFICE_CAMERA_DISTANCE_EMU_AT_45 / EMU_PER_PX) * Math.tan(Math.PI / 8)) / Math.tan(half));
}

