/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's Lock aspect ratio rule for a handle drag, shared by the sheet's
 * Normal view (`resizeObjectBounds`) and Page Layout view
 * (`transformPageLayoutDrawing`): a corner follows the larger relative
 * change, a side handle scales the other axis to match. Which edges stay is
 * each caller's.
 */
export function aspectLockedResizeSize(
  start: { width: number; height: number },
  proposed: { width: number; height: number },
  handle: { horizontal: boolean; vertical: boolean },
  /** Width over height to hold. */
  aspect: number,
  min: { width: number; height: number },
): { width: number; height: number } {
  if (!(aspect > 0)) return proposed;
  const widthRatio = proposed.width / Math.max(1, start.width);
  const heightRatio = proposed.height / Math.max(1, start.height);
  if (handle.horizontal && (!handle.vertical || widthRatio >= heightRatio)) {
    const height = Math.max(min.height, proposed.width / aspect);
    return { width: height * aspect, height };
  }
  const width = Math.max(min.width, proposed.height * aspect);
  return { width, height: width / aspect };
}
