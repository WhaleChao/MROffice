/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Turning page-fraction geometry between the page as stored and the page as
 * shown.
 *
 * A page's content stream is written in one orientation; `/Rotate` (and the
 * reader's own rotation) decide another. pdf.js hands back operator-list
 * transforms in the STORED orientation, and pdf-lib draws in the stored
 * orientation too — but everything a user points at is in the SHOWN one. On an
 * upright page the two are the same, which is why getting this wrong stays
 * invisible until someone opens a scanned document.
 *
 * Rotations are clockwise, matching pdf.js's viewport and `/Rotate`.
 */

export interface FractionRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface FractionPoint {
  x: number;
  y: number;
}

export type QuarterTurn = 0 | 90 | 180 | 270;

export function normalizeQuarterTurn(rotation: number | undefined): QuarterTurn {
  if (!Number.isFinite(rotation ?? NaN)) return 0;
  const turns = (((Math.round((rotation as number) / 90) % 4) + 4) % 4) as
    | 0
    | 1
    | 2
    | 3;
  return (turns * 90) as QuarterTurn;
}

/** Stored orientation to shown orientation. */
export function rotateFractionPoint(
  point: FractionPoint,
  rotation: number,
): FractionPoint {
  switch (normalizeQuarterTurn(rotation)) {
    case 90:
      return { x: 1 - point.y, y: point.x };
    case 180:
      return { x: 1 - point.x, y: 1 - point.y };
    case 270:
      return { x: point.y, y: 1 - point.x };
    default:
      return { x: point.x, y: point.y };
  }
}

/**
 * Stored orientation to shown orientation.
 *
 * A quarter turn swaps width and height: the rect stays a rect, so this maps
 * the two opposite corners and normalises rather than trying to rotate the
 * origin and hoping the size follows.
 */
export function rotateFractionRect(
  rect: FractionRect,
  rotation: number,
): FractionRect {
  const turn = normalizeQuarterTurn(rotation);
  if (turn === 0) return { ...rect };
  const a = rotateFractionPoint({ x: rect.x, y: rect.y }, turn);
  const b = rotateFractionPoint(
    { x: rect.x + rect.width, y: rect.y + rect.height },
    turn,
  );
  return {
    x: Math.min(a.x, b.x),
    y: Math.min(a.y, b.y),
    width: Math.abs(b.x - a.x),
    height: Math.abs(b.y - a.y),
  };
}

/** Shown orientation back to stored orientation. */
export function unrotateFractionPoint(
  point: FractionPoint,
  rotation: number,
): FractionPoint {
  return rotateFractionPoint(point, 360 - normalizeQuarterTurn(rotation));
}

/** Shown orientation back to stored orientation. */
export function unrotateFractionRect(
  rect: FractionRect,
  rotation: number,
): FractionRect {
  return rotateFractionRect(rect, 360 - normalizeQuarterTurn(rotation));
}
