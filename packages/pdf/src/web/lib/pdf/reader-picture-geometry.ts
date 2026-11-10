/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Geometry of a picture selected in place in the PDF reader (WPS 图片编辑).
 *
 * The reader shows a source picture with the frame WPS draws: the picture's
 * own rectangle, turned with it, eight grips and a rotation handle. Every
 * gesture on that frame -- a move, a resize along the picture's own axes, a
 * turn by any angle, a mirror, a quarter turn -- becomes ONE affine in the
 * page's user space, which the native transform writer applies after the
 * paint as the reader walked it (`pageTransform`, see
 * `pdfNativeEmbeddedImagePlacement`). Scaling a turned picture's bounding box
 * instead would shear it; mapping its own rectangle onto the new one never
 * does.
 *
 * Spaces: the paint matrix maps the image's unit square into user space
 * (PDF points, y up); the page's display transform (pdf.js
 * `viewport.transform` at scale 1 and the shown rotation) maps user space to
 * displayed pixels (y down); the frame itself is page PERCENT, top-left
 * origin, plus a clockwise angle, the space `PdfAnnotationFrame` works in.
 * Matrices are PDF arrays `[a b c d e f]` (x' = a x + c y + e,
 * y' = b x + d y + f) and `concatAffine(first, second)` applies `first` then
 * `second`, as `cm` does.
 */

import type { PercentBox } from "./annotation-transform";
import type { ResizeHandle } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";

export type ReaderPictureAffine = readonly [number, number, number, number, number, number];

export const IDENTITY_AFFINE: ReaderPictureAffine = Object.freeze([1, 0, 0, 1, 0, 0]) as ReaderPictureAffine;

export function concatAffine(
  first: ReaderPictureAffine,
  second: ReaderPictureAffine,
): ReaderPictureAffine {
  const [a1, b1, c1, d1, e1, f1] = first;
  const [a2, b2, c2, d2, e2, f2] = second;
  return [
    a1 * a2 + b1 * c2,
    a1 * b2 + b1 * d2,
    c1 * a2 + d1 * c2,
    c1 * b2 + d1 * d2,
    e1 * a2 + f1 * c2 + e2,
    e1 * b2 + f1 * d2 + f2,
  ];
}

/**
 * How a picture painted at a quarter turn (or flipped) shows its pixels:
 * where its pixel rows and columns run on the page, as the canvas transform
 * that draws its W x H pixels as the page shows them, and whether its sides
 * swap. Null for a paint at any other angle, or skewed. `drawn` maps the
 * picture's unit square (rows from its top at t = 1) to displayed pixels,
 * y down, as the reader draws it.
 */
export function readerPictureQuarterOrientation(drawn: ReaderPictureAffine): Readonly<{
  swap: boolean;
  /** The displayed direction a step along a pixel row takes. */
  x: readonly [number, number];
  /** The displayed direction a step down a pixel column takes. */
  y: readonly [number, number];
}> | null {
  const [a, b, c, d] = drawn;
  const tolerance = 1e-6 * Math.max(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
  const sign = (value: number) => (value > 0 ? 1 : -1);
  if (!(tolerance > 0)) return null;
  if (Math.abs(b) <= tolerance && Math.abs(c) <= tolerance) {
    return { swap: false, x: [sign(a), 0], y: [0, sign(-d)] };
  }
  if (Math.abs(a) <= tolerance && Math.abs(d) <= tolerance) {
    return { swap: true, x: [0, sign(b)], y: [sign(-c), 0] };
  }
  return null;
}

export function invertAffine(matrix: ReaderPictureAffine): ReaderPictureAffine | null {
  const [a, b, c, d, e, f] = matrix;
  const det = a * d - b * c;
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return null;
  const ia = d / det;
  const ib = -b / det;
  const ic = -c / det;
  const id = a / det;
  return [ia, ib, ic, id, -(e * ia + f * ic), -(e * ib + f * id)];
}

export function applyAffine(
  matrix: ReaderPictureAffine,
  x: number,
  y: number,
): readonly [number, number] {
  return [
    matrix[0] * x + matrix[2] * y + matrix[4],
    matrix[1] * x + matrix[3] * y + matrix[5],
  ];
}

function translation(x: number, y: number): ReaderPictureAffine {
  return [1, 0, 0, 1, x, y];
}

/** Clockwise on screen (y down) by `degrees`. */
function screenRotation(degrees: number): ReaderPictureAffine {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return [cos, sin, -sin, cos, 0, 0];
}

function scaling(x: number, y: number): ReaderPictureAffine {
  return [x, 0, 0, y, 0, 0];
}

function chain(...matrices: ReaderPictureAffine[]): ReaderPictureAffine {
  return matrices.reduce((acc, next) => concatAffine(acc, next), IDENTITY_AFFINE);
}

/** A page fraction rectangle (0..1, top-left origin), as targets carry it. */
export interface ReaderPictureFractionRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * The picture as the frame shows it, in displayed pixels at scale 1: its
 * centre, its own sides and the clockwise angle they are drawn at. The angle
 * is folded into (-45, 45] (a quarter-turned picture's frame stands upright
 * with its sides swapped: the same rectangle), so a picture turned by the
 * 90-degree buttons keeps an upright frame.
 */
export interface ReaderPictureBox {
  readonly cx: number;
  readonly cy: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  /** The paint skews the picture: the frame is its upright bounding box and
   * resizes it along the page's axes. */
  readonly skewed: boolean;
}

export interface ReaderPictureDisplay {
  /** User space to displayed pixels at scale 1 (pdf.js viewport transform). */
  readonly transform: ReaderPictureAffine;
  /** The displayed page at scale 1, pixels. */
  readonly width: number;
  readonly height: number;
}

const ANGLE_EPSILON = 0.01;

function foldAngle(degrees: number): { rotation: number; swapped: boolean } {
  const turns = Math.round(degrees / 90);
  let rotation = degrees - turns * 90;
  if (Math.abs(rotation) < ANGLE_EPSILON) rotation = 0;
  // (-45, 45]: a picture at exactly -45 shows as +45 with its sides swapped.
  if (rotation <= -45 + 1e-9) {
    return { rotation: rotation + 90, swapped: (turns - 1) % 2 !== 0 };
  }
  return { rotation, swapped: turns % 2 !== 0 };
}

/**
 * The frame of a painted picture. `visible` is the picture's visible box as
 * the page shows it (page fractions): for an upright picture it is the frame
 * exactly, a clip of its own included; a turned picture's frame is its whole
 * paint (its clip, if any, travels with it inside the paint).
 */
export function readerPictureBox(input: {
  readonly matrix: ReaderPictureAffine;
  readonly display: ReaderPictureDisplay;
  readonly visible: ReaderPictureFractionRect;
}): ReaderPictureBox {
  const { display, visible } = input;
  const drawn = concatAffine(input.matrix, display.transform);
  const u = [drawn[0], drawn[1]] as const;
  const v = [drawn[2], drawn[3]] as const;
  const lengthU = Math.hypot(u[0], u[1]);
  const lengthV = Math.hypot(v[0], v[1]);
  const upright = (): ReaderPictureBox => ({
    cx: (visible.x + visible.width / 2) * display.width,
    cy: (visible.y + visible.height / 2) * display.height,
    width: visible.width * display.width,
    height: visible.height * display.height,
    rotation: 0,
    skewed: false,
  });
  if (!(lengthU > 0) || !(lengthV > 0)) return upright();
  const skew = Math.abs(u[0] * v[0] + u[1] * v[1]) / (lengthU * lengthV);
  if (skew > 1e-3) return { ...upright(), skewed: true };
  const angle = (Math.atan2(u[1], u[0]) * 180) / Math.PI;
  const folded = foldAngle(angle);
  if (folded.rotation === 0) return upright();
  const [cx, cy] = applyAffine(drawn, 0.5, 0.5);
  return {
    cx,
    cy,
    width: folded.swapped ? lengthV : lengthU,
    height: folded.swapped ? lengthU : lengthV,
    rotation: folded.rotation,
    skewed: false,
  };
}

/** The frame in page percent (the unturned rectangle about the centre). */
export function readerPicturePercentBox(
  box: ReaderPictureBox,
  display: Pick<ReaderPictureDisplay, "width" | "height">,
): PercentBox {
  return {
    x: ((box.cx - box.width / 2) / display.width) * 100,
    y: ((box.cy - box.height / 2) / display.height) * 100,
    width: (box.width / display.width) * 100,
    height: (box.height / display.height) * 100,
  };
}

export function readerPictureBoxFromPercent(
  percent: PercentBox,
  rotation: number,
  display: Pick<ReaderPictureDisplay, "width" | "height">,
  skewed = false,
): ReaderPictureBox {
  const width = (percent.width / 100) * display.width;
  const height = (percent.height / 100) * display.height;
  return {
    cx: (percent.x / 100) * display.width + width / 2,
    cy: (percent.y / 100) * display.height + height / 2,
    width,
    height,
    rotation,
    skewed,
  };
}

/** The upright bounding box of a frame, displayed pixels. */
export function readerPictureBounds(box: ReaderPictureBox): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  const radians = (box.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(radians));
  const sin = Math.abs(Math.sin(radians));
  const width = box.width * cos + box.height * sin;
  const height = box.width * sin + box.height * cos;
  return { x: box.cx - width / 2, y: box.cy - height / 2, width, height };
}

/**
 * The displayed-pixel affine that carries the picture from one frame to
 * another with the same angle convention: about its centre, along its own
 * axes. A skewed picture's frames are upright boxes and scale on the page's
 * axes.
 */
export function readerPictureDisplayTransform(
  from: ReaderPictureBox,
  to: ReaderPictureBox,
): ReaderPictureAffine | null {
  if (!(from.width > 0) || !(from.height > 0) || !(to.width > 0) || !(to.height > 0)) {
    return null;
  }
  return chain(
    translation(-from.cx, -from.cy),
    screenRotation(-from.rotation),
    scaling(to.width / from.width, to.height / from.height),
    screenRotation(to.rotation),
    translation(to.cx, to.cy),
  );
}

/** A turn by `degrees` clockwise on screen about the frame's centre. */
export function readerPictureTurn(
  box: Pick<ReaderPictureBox, "cx" | "cy">,
  degrees: number,
): ReaderPictureAffine {
  return chain(
    translation(-box.cx, -box.cy),
    screenRotation(degrees),
    translation(box.cx, box.cy),
  );
}






export function readerPictureFlip(
  box: Pick<ReaderPictureBox, "cx" | "cy">,
  axis: "horizontal" | "vertical",
): ReaderPictureAffine {
  return chain(
    translation(-box.cx, -box.cy),
    axis === "horizontal" ? scaling(-1, 1) : scaling(1, -1),
    translation(box.cx, box.cy),
  );
}

/** The same motion as a displayed-pixel affine, in the page's user space. */
export function readerPictureUserTransform(
  displayTransform: ReaderPictureAffine,
  motion: ReaderPictureAffine,
): ReaderPictureAffine | null {
  const inverse = invertAffine(displayTransform);
  if (!inverse) return null;
  return chain(displayTransform, motion, inverse);
}

/** True when an affine moves nothing a reader could see (< 0.01 px). */
export function readerPictureAffineIsIdentity(matrix: ReaderPictureAffine): boolean {
  return (
    Math.abs(matrix[0] - 1) < 1e-6 &&
    Math.abs(matrix[1]) < 1e-6 &&
    Math.abs(matrix[2]) < 1e-6 &&
    Math.abs(matrix[3] - 1) < 1e-6 &&
    Math.abs(matrix[4]) < 1e-2 &&
    Math.abs(matrix[5]) < 1e-2
  );
}

/** Pixels of a picture that must stay on the page when it is dragged off
 * an edge (WPS keeps a part of the object on the page). */
export const READER_PICTURE_KEEP_ON_PAGE = 24;

/**
 * One frame gesture in page percent, as `PdfAnnotationFrame` reports it
 * (`start` is the frame's unturned box, `dx`/`dy` the pointer travel in page
 * percent). A move may carry the picture partly off the page but keeps a
 * part on it; a grip resizes along the picture's own (turned) axes with the
 * opposite side or corner held, and `keepRatio` (Shift) keeps the
 * proportions. A picture never turns inside out (no mirror by dragging).
 */
export function transformReaderPictureFrame(
  start: PercentBox,
  mode: "move" | ResizeHandle,
  dx: number,
  dy: number,
  options: {
    readonly rotation: number;
    readonly pageWidth: number;
    readonly pageHeight: number;
    readonly keepRatio?: boolean;
    /** Smallest side, pixels at `pageWidth`/`pageHeight`. */
    readonly minSide?: number;
  },
): PercentBox {
  const page = { width: options.pageWidth, height: options.pageHeight };
  if (!(page.width > 0) || !(page.height > 0)) return start;
  const box = readerPictureBoxFromPercent(start, options.rotation, page);
  const travelX = (dx / 100) * page.width;
  const travelY = (dy / 100) * page.height;
  if (mode === "move") {
    const bounds = readerPictureBounds(box);
    const keepX = Math.min(bounds.width, READER_PICTURE_KEEP_ON_PAGE);
    const keepY = Math.min(bounds.height, READER_PICTURE_KEEP_ON_PAGE);
    const shiftX = Math.min(
      Math.max(travelX, keepX - (bounds.x + bounds.width)),
      page.width - keepX - bounds.x,
    );
    const shiftY = Math.min(
      Math.max(travelY, keepY - (bounds.y + bounds.height)),
      page.height - keepY - bounds.y,
    );
    return readerPicturePercentBox(
      { ...box, cx: box.cx + shiftX, cy: box.cy + shiftY },
      page,
    );
  }
  const radians = (box.rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  // The pointer's travel along the frame's own axes.
  const localX = travelX * cos + travelY * sin;
  const localY = -travelX * sin + travelY * cos;
  const minSide = Math.max(1, options.minSide ?? 4);
  const east = mode.includes("e");
  const west = mode.includes("w");
  const south = mode.includes("s");
  const north = mode.includes("n");
  let width = box.width;
  let height = box.height;
  if (east) width = box.width + localX;
  if (west) width = box.width - localX;
  if (south) height = box.height + localY;
  if (north) height = box.height - localY;
  width = Math.max(minSide, width);
  height = Math.max(minSide, height);
  const horizontal = east || west;
  const vertical = north || south;
  if (options.keepRatio && box.width > 0 && box.height > 0) {
    const scale =
      horizontal && vertical
        ? Math.max(width / box.width, height / box.height)
        : horizontal
          ? width / box.width
          : height / box.height;
    const floor = Math.max(minSide / box.width, minSide / box.height);
    const applied = Math.max(scale, floor);
    width = box.width * applied;
    height = box.height * applied;
  }
  // The held side or corner stays where it is; an edge grip keeps the other
  // axis centred (as Office does when the proportions are kept).
  const shiftLocalX = east ? (width - box.width) / 2 : west ? -(width - box.width) / 2 : 0;
  const shiftLocalY = south ? (height - box.height) / 2 : north ? -(height - box.height) / 2 : 0;
  const cx = box.cx + shiftLocalX * cos - shiftLocalY * sin;
  const cy = box.cy + shiftLocalX * sin + shiftLocalY * cos;
  return readerPicturePercentBox({ ...box, cx, cy, width, height }, page);
}

/**
 * A crop drawn on the frame, in the frame's own fractions (0..1, top-left
 * origin along the turned axes), as the native writer's crop: fractions of
 * the visible part of the image in the image's own square, measured from its
 * top-left. The picture is drawn with `matrix` (unit square to user space);
 * `visibleUnit` is its visible part in that square (the whole square when it
 * has no clip). Null when the crop keeps everything or cannot be stated.
 */
export function readerPictureCropIntent(input: {
  readonly box: ReaderPictureBox;
  readonly crop: ReaderPictureFractionRect;
  readonly matrix: ReaderPictureAffine;
  readonly display: ReaderPictureDisplay;
  readonly visibleUnit: ReaderPictureFractionRect;
}): ReaderPictureFractionRect | null {
  const { box, crop, visibleUnit } = input;
  if (
    crop.x <= 1e-6 &&
    crop.y <= 1e-6 &&
    crop.width >= 1 - 1e-6 &&
    crop.height >= 1 - 1e-6
  ) {
    return null;
  }
  const drawn = concatAffine(input.matrix, input.display.transform);
  const toUnit = invertAffine(drawn);
  if (!toUnit || !(visibleUnit.width > 0) || !(visibleUnit.height > 0)) return null;
  const radians = (box.rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [fx, fy] of [
    [crop.x, crop.y],
    [crop.x + crop.width, crop.y],
    [crop.x, crop.y + crop.height],
    [crop.x + crop.width, crop.y + crop.height],
  ] as const) {
    const localX = (fx - 0.5) * box.width;
    const localY = (fy - 0.5) * box.height;
    const px = box.cx + localX * cos - localY * sin;
    const py = box.cy + localX * sin + localY * cos;
    const [ux, uy] = applyAffine(toUnit, px, py);
    xs.push(ux);
    ys.push(uy);
  }
  const left = Math.max(visibleUnit.x, Math.min(...xs));
  const right = Math.min(visibleUnit.x + visibleUnit.width, Math.max(...xs));
  const bottom = Math.max(visibleUnit.y, Math.min(...ys));
  const top = Math.min(visibleUnit.y + visibleUnit.height, Math.max(...ys));
  if (!(right - left > 1e-6) || !(top - bottom > 1e-6)) return null;
  const result = {
    x: (left - visibleUnit.x) / visibleUnit.width,
    // The writer measures the crop from the top of the visible part; the
    // unit square's y runs up.
    y: (visibleUnit.y + visibleUnit.height - top) / visibleUnit.height,
    width: (right - left) / visibleUnit.width,
    height: (top - bottom) / visibleUnit.height,
  };
  return {
    x: Math.min(1, Math.max(0, result.x)),
    y: Math.min(1, Math.max(0, result.y)),
    width: Math.min(1, Math.max(1e-4, result.width)),
    height: Math.min(1, Math.max(1e-4, result.height)),
  };
}

/**
 * The visible part of a picture in its own unit square: the frame's corners
 * taken back through the paint. An upright clipped picture's frame is its
 * visible box, so this is its clip; a turned picture's frame is its whole
 * paint (the full square).
 */
export function readerPictureVisibleUnitRect(input: {
  readonly box: ReaderPictureBox;
  readonly matrix: ReaderPictureAffine;
  readonly display: ReaderPictureDisplay;
}): ReaderPictureFractionRect {
  const toUnit = invertAffine(concatAffine(input.matrix, input.display.transform));
  if (!toUnit) return { x: 0, y: 0, width: 1, height: 1 };
  const { box } = input;
  const radians = (box.rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [sx, sy] of [
    [-0.5, -0.5],
    [0.5, -0.5],
    [-0.5, 0.5],
    [0.5, 0.5],
  ] as const) {
    const localX = sx * box.width;
    const localY = sy * box.height;
    const [ux, uy] = applyAffine(
      toUnit,
      box.cx + localX * cos - localY * sin,
      box.cy + localX * sin + localY * cos,
    );
    xs.push(ux);
    ys.push(uy);
  }
  const x = Math.max(0, Math.min(...xs));
  const y = Math.max(0, Math.min(...ys));
  const width = Math.min(1, Math.max(...xs)) - x;
  const height = Math.min(1, Math.max(...ys)) - y;
  return width > 1e-6 && height > 1e-6
    ? { x, y, width, height }
    : { x: 0, y: 0, width: 1, height: 1 };
}

/**
 * The crop that keeps a page-cut picture's visible part (pageClipped,
 * lib/pdf/page-picture-targets.ts) as a clip of its own, in the writer's
 * convention (measured on the whole square, from its top): a geometry write
 * carries it, so the clip moves, scales and turns with the picture, as WPS
 * moves and resizes the part it frames. Null when the whole square shows.
 */
export function readerPicturePageClipCrop(input: {
  readonly box: ReaderPictureBox;
  readonly matrix: ReaderPictureAffine;
  readonly display: ReaderPictureDisplay;
}): ReaderPictureFractionRect | null {
  const unit = readerPictureVisibleUnitRect(input);
  if (unit.x <= 1e-6 && unit.y <= 1e-6 && unit.width >= 1 - 1e-6 && unit.height >= 1 - 1e-6) return null;
  return { x: unit.x, y: 1 - unit.y - unit.height, width: unit.width, height: unit.height };
}

/**
 * A crop measured on a page-cut picture's visible part (readerPictureCropIntent),
 * measured on its whole square instead: the writer measures a crop on the
 * clip it reads in the picture's own level, and the page's cut is none.
 */
export function readerPictureCropOnWholeSquare(
  crop: ReaderPictureFractionRect,
  visibleUnit: ReaderPictureFractionRect,
): ReaderPictureFractionRect {
  const top = 1 - visibleUnit.y - visibleUnit.height;
  return {
    x: visibleUnit.x + crop.x * visibleUnit.width,
    y: top + crop.y * visibleUnit.height,
    width: crop.width * visibleUnit.width,
    height: crop.height * visibleUnit.height,
  };
}

/**
 * Where a pasted copy of a picture is painted: the copy keeps the size,
 * angle and mirror the original is shown with, and its upright bounding box
 * has its top-left corner at `point` (page fractions of the page it lands
 * on), as WPS pastes at the pointer. The copy's pixels are the original's
 * visible part (`visibleUnit` of its square); the copy may run past the
 * page, which then cuts it. Returns the copy's unit-square-to-user-space
 * matrix on the target page.
 */
export function readerPicturePasteMatrix(input: {
  readonly matrix: ReaderPictureAffine;
  readonly visibleUnit: ReaderPictureFractionRect;
  readonly sourceDisplay: ReaderPictureDisplay;
  readonly targetDisplay: ReaderPictureDisplay;
  readonly point: { readonly x: number; readonly y: number };
}): ReaderPictureAffine | null {
  const { visibleUnit } = input;
  const piece: ReaderPictureAffine = [
    visibleUnit.width,
    0,
    0,
    visibleUnit.height,
    visibleUnit.x,
    visibleUnit.y,
  ];
  const drawn = chain(piece, input.matrix, input.sourceDisplay.transform);
  const corners = [
    applyAffine(drawn, 0, 0),
    applyAffine(drawn, 1, 0),
    applyAffine(drawn, 0, 1),
    applyAffine(drawn, 1, 1),
  ];
  const left = Math.min(...corners.map(([x]) => x));
  const top = Math.min(...corners.map(([, y]) => y));
  const moved = concatAffine(
    drawn,
    translation(
      input.point.x * input.targetDisplay.width - left,
      input.point.y * input.targetDisplay.height - top,
    ),
  );
  const toUser = invertAffine(input.targetDisplay.transform);
  if (!toUser) return null;
  const result = concatAffine(moved, toUser);
  return result.every(Number.isFinite) ? result : null;
}
