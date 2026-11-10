/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A run turned by the rotate handle (WPS turns a text block to any angle):
 * the turn is written into the run's own text matrix, which the property
 * writer turns about the centre of the run's box (list-text-fidelity.ts,
 * propertyRotationPivotOffset). Browser-safe math only.
 */

type Linear = readonly [number, number, number, number];

const round = (value: number): number => {
  const rounded = Math.round(value * 1e6) / 1e6;
  return Object.is(rounded, -0) ? 0 : rounded;
};

/** Row-vector product of two linear parts (`left` applied first). */
function multiply(left: Linear, right: Linear): Linear {
  return [
    left[0] * right[0] + left[1] * right[2],
    left[0] * right[1] + left[1] * right[3],
    left[2] * right[0] + left[3] * right[2],
    left[2] * right[1] + left[3] * right[3],
  ];
}

function invert(linear: Linear): Linear | null {
  const determinant = linear[0] * linear[3] - linear[1] * linear[2];
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) return null;
  return [linear[3] / determinant, -linear[1] / determinant, -linear[2] / determinant, linear[0] / determinant];
}

/**
 * The linear part of a run's text matrix turned `degrees` clockwise as the
 * page is seen. A screen shows the page y-flipped, so the turn in page space
 * runs the other way; the run is drawn under its graphics matrix, so the
 * turn is taken through it (a run under a flipped or scaled `cm` turns the
 * way its frame turns on screen). Null for a matrix that cannot be turned.
 */
export function pdfNativeTextLinearTurned(
  linear: Linear,
  graphicsMatrix: readonly number[] | null | undefined,
  degrees: number,
): Linear | null {
  if (!Number.isFinite(degrees) || linear.some((value) => !Number.isFinite(value))) return null;
  const graphics: Linear = graphicsMatrix && graphicsMatrix.length >= 4
    ? [graphicsMatrix[0], graphicsMatrix[1], graphicsMatrix[2], graphicsMatrix[3]]
    : [1, 0, 0, 1];
  const inverse = invert(graphics);
  if (!inverse || !invert(linear)) return null;
  // Clockwise on screen is counter-clockwise by the negative angle in page
  // space (y up): a row vector (x, y) goes to (x cos + y sin, -x sin + y cos).
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const turn: Linear = [cos, -sin, sin, cos];
  const turned = multiply(multiply(multiply(linear, graphics), turn), inverse);
  const result: Linear = [round(turned[0]), round(turned[1]), round(turned[2]), round(turned[3])];
  return invert(result) ? result : null;
}
