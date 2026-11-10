/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Ruled lines and small boxes drawn as vector graphics.
 *
 * Real forms rarely use underscore characters: the blank after "Name:" is a
 * DRAWN line, and the tick box is a DRAWN square. Neither exists in the text
 * layer at all, which is why the flat-form detector misses exactly the forms
 * people most want filled. The evidence lives in the operator list, and
 * pdf.js's `constructPath` hands over the one thing needed — each path's
 * bounding box — so classification needs no path decoding, only the same
 * transform tracking the image walker already does.
 */

import { multiplyMatrix, type ImageOperatorIds } from "./embedded-images";

export interface VectorMarkOperatorIds extends
  Pick<ImageOperatorIds, "save" | "restore" | "transform"> {
  constructPath: number | undefined;
  /** Paint op ids that STROKE (stroke, closeStroke, fillStroke, …). */
  strokeOps: ReadonlyArray<number>;
  /** Paint op ids that FILL (fill, eoFill, fillStroke, …). */
  fillOps: ReadonlyArray<number>;
}

export type VectorMarkKind = "rule" | "box";

export interface VectorMark {
  kind: VectorMarkKind;
  /** Page fractions (0..1), top-left origin. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Paint order, for stable tie-breaks downstream. */
  order: number;
}

/** A rule thinner than this many points is a line, not a box. */
const RULE_MAX_THICKNESS_PT = 2.5;
/** A rule shorter than this many points is a dash, not a blank. */
const RULE_MIN_LENGTH_PT = 24;
/** Tick-box side lengths, in points. */
const BOX_MIN_SIDE_PT = 6;
const BOX_MAX_SIDE_PT = 30;
/** Boxes squarer than this aspect band are not tick boxes. */
const BOX_MIN_ASPECT = 0.6;
const BOX_MAX_ASPECT = 1.7;

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

function mapPoint(m: Matrix, x: number, y: number): { x: number; y: number } {
  return { x: m[0] * x + m[2] * y + m[4], y: m[1] * x + m[3] * y + m[5] };
}

/**
 * Collect rules and tick-box-sized rectangles from a page's operator list.
 *
 * `constructPath` args are `[paintOp, pathData, minMax]`; the minMax box is in
 * the CURRENT transform's space, so the walker keeps the same save/restore/
 * transform stack the image walker keeps, and maps the box's corners through
 * it — all four, because a rotated CTM turns an axis-aligned box.
 */
export function vectorMarksFromOperatorList(
  operatorList: { fnArray: ArrayLike<number>; argsArray: ArrayLike<unknown> },
  ids: VectorMarkOperatorIds,
  page: { width: number; height: number },
): VectorMark[] {
  if (!(page.width > 0) || !(page.height > 0)) return [];
  const marks: VectorMark[] = [];
  const stack: Matrix[] = [];
  let ctm: Matrix = [...IDENTITY];
  let order = 0;

  const strokeSet = new Set(ids.strokeOps);
  const fillSet = new Set(ids.fillOps);

  for (let i = 0; i < operatorList.fnArray.length; i += 1) {
    const fn = operatorList.fnArray[i];
    const args = operatorList.argsArray[i];
    if (fn === ids.save) {
      stack.push([...ctm]);
      continue;
    }
    if (fn === ids.restore) {
      ctm = stack.pop() ?? [...IDENTITY];
      continue;
    }
    if (fn === ids.transform) {
      const raw = Array.isArray(args) && Array.isArray(args[0]) ? args[0] : args;
      if (Array.isArray(raw) && raw.length >= 6) {
        ctm = multiplyMatrix(
          ctm,
          raw.slice(0, 6) as unknown as Matrix,
        ) as Matrix;
      }
      continue;
    }
    if (ids.constructPath === undefined || fn !== ids.constructPath) continue;
    if (!Array.isArray(args) || args.length < 3) continue;
    const paintOp = Number(args[0]);
    const painted = strokeSet.has(paintOp) || fillSet.has(paintOp);
    if (!painted) continue;
    const minMax = args[2] as Record<number, number> | number[] | undefined;
    if (!minMax) continue;
    const rawMinX = Number((minMax as Record<number, number>)[0]);
    const rawMinY = Number((minMax as Record<number, number>)[1]);
    const rawMaxX = Number((minMax as Record<number, number>)[2]);
    const rawMaxY = Number((minMax as Record<number, number>)[3]);
    if (
      ![rawMinX, rawMinY, rawMaxX, rawMaxY].every(Number.isFinite) ||
      rawMaxX < rawMinX ||
      rawMaxY < rawMinY
    ) {
      continue;
    }
    const corners = [
      mapPoint(ctm, rawMinX, rawMinY),
      mapPoint(ctm, rawMaxX, rawMinY),
      mapPoint(ctm, rawMaxX, rawMaxY),
      mapPoint(ctm, rawMinX, rawMaxY),
    ];
    const xs = corners.map((corner) => corner.x);
    const ys = corners.map((corner) => corner.y);
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    const widthPt = maxX - minX;
    const heightPt = maxY - minY;

    let kind: VectorMarkKind | null = null;
    if (
      heightPt <= RULE_MAX_THICKNESS_PT &&
      widthPt >= RULE_MIN_LENGTH_PT &&
      strokeSet.has(paintOp) === true
    ) {
      kind = "rule";
    } else if (
      widthPt >= BOX_MIN_SIDE_PT &&
      widthPt <= BOX_MAX_SIDE_PT &&
      heightPt >= BOX_MIN_SIDE_PT &&
      heightPt <= BOX_MAX_SIDE_PT &&
      widthPt / heightPt >= BOX_MIN_ASPECT &&
      widthPt / heightPt <= BOX_MAX_ASPECT
    ) {
      kind = "box";
    }
    if (!kind) continue;

    marks.push({
      kind,
      x: minX / page.width,
      // User space counts from the bottom; the model from the top.
      y: (page.height - maxY) / page.height,
      width: widthPt / page.width,
      height: Math.max(heightPt, 0.5) / page.height,
      order: (order += 1) - 1,
    });
  }
  return marks;
}

/** A filled underline rule also arrives as a thin filled RECT; accept both. */
export function isRuleLike(mark: VectorMark, page: { height: number }): boolean {
  return mark.kind === "rule" || mark.height * page.height <= RULE_MAX_THICKNESS_PT;
}
