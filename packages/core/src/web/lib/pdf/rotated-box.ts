/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeTurnAround } from "../office-object-frame";
/**
 * Geometry for a text box that is not axis-aligned.
 *
 * Angled text — a watermark across a page, a caption set on a diagram's slope —
 * needs three things that an upright box gets for free: a hit test that knows
 * where the box really is, a pointer position mapped back into the box's own
 * coordinates so a caret lands on the character under the cursor, and a mask
 * big enough to cover what the rotated box actually paints.
 *
 * The rotation is about the box's CENTRE. That is what CSS `transform: rotate`
 * does by default, so the overlay needs no compensating translate — and it is
 * what a user means by "turn this box", since the box stays where they put it
 * instead of swinging away from the pointer. pdf-lib rotates about the text's
 * anchor instead, so the export has to rotate the anchor about the centre; that
 * conversion is `rotatedAnchor` below rather than something each caller works
 * out again.
 *
 * Angles are degrees, clockwise, in the top-left-origin space the editor uses.
 * Everything here is unit-agnostic: page fractions, percent and points all
 * work, as long as one call does not mix them.
 */

export interface BoxPoint {
  x: number;
  y: number;
}

export interface BoxRectLike {
  x: number;
  y: number;
  width: number;
  height: number;
}


/** Wrap to (-180, 180], and treat anything unusable as upright. */
export function normalizeBoxAngle(angle: number | undefined | null): number {
  if (!Number.isFinite(angle ?? NaN)) return 0;
  let value = (angle as number) % 360;
  if (value > 180) value -= 360;
  if (value <= -180) value += 360;
  // -0 compares equal to 0 but prints differently and leaks into snapshots.
  return value === 0 ? 0 : value;
}

/** Clockwise rotation of `point` about `origin`, in a top-left-origin space. */
export function rotatePointAround(
  point: BoxPoint,
  origin: BoxPoint,
  angle: number,
): BoxPoint {
  const a = normalizeBoxAngle(angle);
  if (a === 0) return { x: point.x, y: point.y };
  // y grows downward, so a clockwise turn on screen is the standard positive
  // rotation: the one every editor's placed objects turn by.
  return officeTurnAround(point, origin, a);
}

/**
 * The middle of a rect that spans [y, y + height] — the top-left-origin
 * convention this module works in.
 *
 * PDF user space is the other way up: there a box's stored `y` is often its TOP
 * and it spans [y - height, y], so its middle is BELOW `y`. Passing a user-space
 * box here puts the centre a whole box-height out, and everything rotated about
 * it lands somewhere plausible and wrong.
 */
export function boxCentre(rect: BoxRectLike): BoxPoint {
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}

/** The four corners of the rotated box, clockwise from its top-left. */
export function rotatedBoxCorners(
  rect: BoxRectLike,
  angle: number,
): [BoxPoint, BoxPoint, BoxPoint, BoxPoint] {
  const centre = boxCentre(rect);
  const corners: BoxPoint[] = [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x, y: rect.y + rect.height },
  ];
  return corners.map((corner) => rotatePointAround(corner, centre, angle)) as [
    BoxPoint,
    BoxPoint,
    BoxPoint,
    BoxPoint,
  ];
}

/**
 * The axis-aligned box that contains the rotated one.
 *
 * This is what a mask has to cover: masking the unrotated rectangle leaves the
 * corners of the old text sticking out, which is the same halo the upright
 * cover has to overshoot to avoid, only much larger.
 */
export function rotatedBoundingBox(
  rect: BoxRectLike,
  angle: number,
): BoxRectLike {
  if (normalizeBoxAngle(angle) === 0) return { ...rect };
  const corners = rotatedBoxCorners(rect, angle);
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return {
    x,
    y,
    width: Math.max(...xs) - x,
    height: Math.max(...ys) - y,
  };
}

/**
 * A page point in the box's own unrotated coordinates.
 *
 * The result is relative to the box's top-left, so (0, 0) is its first
 * character's corner whatever the box is turned to — which is what caret and
 * selection maths need.
 */
export function pointInRotatedBox(
  point: BoxPoint,
  rect: BoxRectLike,
  angle: number,
): BoxPoint {
  const unrotated = rotatePointAround(point, boxCentre(rect), -angle);
  return { x: unrotated.x - rect.x, y: unrotated.y - rect.y };
}

/** Whether a page point is inside the rotated box. */
export function isPointInRotatedBox(
  point: BoxPoint,
  rect: BoxRectLike,
  angle: number,
): boolean {
  if (!(rect.width > 0) || !(rect.height > 0)) return false;
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  const local = pointInRotatedBox(point, rect, angle);
  return (
    local.x >= 0 &&
    local.y >= 0 &&
    local.x <= rect.width &&
    local.y <= rect.height
  );
}

/**
 * Where to anchor rotated text for a renderer that turns about the anchor.
 *
 * pdf-lib's `drawText({ rotate })` spins around the point it is given. The
 * model rotates about the centre, so the anchor is the box's own corner taken
 * for a ride around that centre. Handing the unrotated corner instead is the
 * classic version of this bug: the text is at the right angle and in the wrong
 * place, more wrong the further it is from square.
 */
export function rotatedAnchor(
  rect: BoxRectLike,
  angle: number,
  corner: BoxPoint = { x: rect.x, y: rect.y },
): BoxPoint {
  return rotatePointAround(corner, boxCentre(rect), angle);
}

/** A quarter turn, clockwise: the four angles a page (or a line on it) is
 * turned by where its text still runs along the page's own edges. */
export type BoxQuarterTurn = 0 | 90 | 180 | 270;

/** The quarter turn `angle` is (within a thousandth of a degree), or null
 * for any other angle. */
export function boxQuarterTurn(angle: number | undefined | null): BoxQuarterTurn | null {
  const value = normalizeBoxAngle(angle);
  const turns = Math.round(value / 90);
  if (Math.abs(value - turns * 90) > 0.001) return null;
  return ((((turns % 4) + 4) % 4) * 90) as BoxQuarterTurn;
}

/**
 * A box as it is shown on a page -- fractions of the page, its width and
 * height along its own axes (the width a fraction of the page's width, the
 * height of its height), turned by `turn` about its centre -- as it stands on
 * the page turned back by `turn`: there it is upright, its own left edge where
 * its line starts and its top where its first row is, in fractions of the page
 * turned back (`pageHeight` wide for a quarter turn). A line's room on the
 * page, how it grows as it is typed on and where it wraps are read there, as
 * on an upright page. (pdf.js lays a turned page's text out so, and turns the
 * whole layer: `pdfTextLayerLocalRect`.)
 */
export function boxRectInOwnAxes(
  rect: Readonly<BoxRectLike>,
  turn: BoxQuarterTurn,
  pageWidth: number,
  pageHeight: number,
): BoxRectLike {
  if (turn === 0 || !(pageWidth > 0) || !(pageHeight > 0)) {
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  }
  const cx = (rect.x + rect.width / 2) * pageWidth;
  const cy = (rect.y + rect.height / 2) * pageHeight;
  const own = pageSizeInOwnAxes(turn, pageWidth, pageHeight);
  const ownX = turn === 90 ? cy : turn === 180 ? pageWidth - cx : pageHeight - cy;
  const ownY = turn === 90 ? pageWidth - cx : turn === 180 ? pageHeight - cy : cx;
  const width = (rect.width * pageWidth) / own.width;
  const height = (rect.height * pageHeight) / own.height;
  return { x: ownX / own.width - width / 2, y: ownY / own.height - height / 2, width, height };
}

/** `boxRectInOwnAxes` undone: a box upright on the page turned back by
 * `turn` (fractions of that page) as it is shown on the page, turned by
 * `turn` about its centre. */
export function boxRectFromOwnAxes(
  rect: Readonly<BoxRectLike>,
  turn: BoxQuarterTurn,
  pageWidth: number,
  pageHeight: number,
): BoxRectLike {
  if (turn === 0 || !(pageWidth > 0) || !(pageHeight > 0)) {
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  }
  const own = pageSizeInOwnAxes(turn, pageWidth, pageHeight);
  const ownX = (rect.x + rect.width / 2) * own.width;
  const ownY = (rect.y + rect.height / 2) * own.height;
  const cx = turn === 90 ? pageWidth - ownY : turn === 180 ? pageWidth - ownX : ownY;
  const cy = turn === 90 ? ownX : turn === 180 ? pageHeight - ownY : pageHeight - ownX;
  const width = (rect.width * own.width) / pageWidth;
  const height = (rect.height * own.height) / pageHeight;
  return { x: cx / pageWidth - width / 2, y: cy / pageHeight - height / 2, width, height };
}

/** The page's size turned back by `turn` (`boxRectInOwnAxes`'s page). */
export function pageSizeInOwnAxes(
  turn: BoxQuarterTurn,
  pageWidth: number,
  pageHeight: number,
): Readonly<{ width: number; height: number }> {
  return turn % 180 ? { width: pageHeight, height: pageWidth } : { width: pageWidth, height: pageHeight };
}

/**
 * A box shown turned by `turn` given a new own width and/or height (in the
 * box's own convention: the width a fraction of the page's width, the height
 * of its height), its own top-left corner -- where its line starts, the top
 * of its first row -- kept, and its own right and bottom edges kept on the
 * page. A box turned about its centre and widened from its x grew from its
 * end on a page turned half round: the line's start moved by the growth.
 */
export function resizeBoxFromItsStart(
  rect: Readonly<BoxRectLike>,
  turn: BoxQuarterTurn,
  size: Readonly<{ width?: number; height?: number }>,
  pageWidth: number,
  pageHeight: number,
): BoxRectLike {
  const own = boxRectInOwnAxes(rect, turn, pageWidth, pageHeight);
  const ownPage = turn === 0 ? { width: 1, height: 1 } : pageSizeInOwnAxes(turn, pageWidth, pageHeight);
  const toOwnWidth = turn === 0 ? 1 : pageWidth / ownPage.width;
  const toOwnHeight = turn === 0 ? 1 : pageHeight / ownPage.height;
  const resized = {
    ...own,
    width: size.width === undefined ? own.width : Math.min(1 - own.x, size.width * toOwnWidth),
    height: size.height === undefined ? own.height : Math.min(1 - own.y, size.height * toOwnHeight),
  };
  return boxRectFromOwnAxes(resized, turn, pageWidth, pageHeight);
}

/**
 * A 2D affine map `[a, b, c, d, e, f]` (x' = a x + c y + e, y' = b x + d y + f)
 * into a page's pixels (`pageWidth` x `pageHeight`), followed by the page's
 * turn back by `turn`: points then stand in the pixels of the page turned back
 * (`boxRectInOwnAxes`), where a line turned by `turn` reads left to right.
 */
export function turnBackPageTransform(
  matrix: readonly number[],
  turn: BoxQuarterTurn,
  pageWidth: number,
  pageHeight: number,
): [number, number, number, number, number, number] {
  const [a, b, c, d, e, f] = matrix;
  if (turn === 0) return [a, b, c, d, e, f];
  // The turn back: 90 maps (x, y) to (y, W - x); 180 to (W - x, H - y);
  // 270 to (H - y, x).
  const [ra, rb, rc, rd, re, rf] = turn === 90 ? [0, -1, 1, 0, 0, pageWidth]
    : turn === 180 ? [-1, 0, 0, -1, pageWidth, pageHeight] : [0, 1, -1, 0, pageHeight, 0];
  return [
    ra * a + rc * b, rb * a + rd * b,
    ra * c + rc * d, rb * c + rd * d,
    ra * e + rc * f + re, rb * e + rd * f + rf,
  ];
}
