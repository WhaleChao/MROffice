/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's length units, one source for every editor (ECMA-376 20.1.2.1:
 * 914400 EMU to the inch). A copy elsewhere must say the same number;
 * `office-units.test.ts` reads the source tree for every one.
 */
export const EMU_PER_INCH = 914_400;
/** At CSS's 96 pixels to the inch. */
export const EMU_PER_PX = 9_525;
export const EMU_PER_PT = 12_700;
export const EMU_PER_CM = 360_000;
export const EMU_PER_MM = 36_000;
/** A twentieth of a point (Word's twips). */
export const EMU_PER_TWIP = 635;

export const emuToPx = (value: number): number => value / EMU_PER_PX;
export const pxToEmu = (value: number): number => Math.round(value * EMU_PER_PX);
export const emuToPt = (value: number): number => value / EMU_PER_PT;








export const PX_PER_PT = EMU_PER_PT / EMU_PER_PX;
// Written as the exact ratio (12 700 / 9 525 is 4 / 3), multiplied before it
// divides: the same correctly rounded value every hand-written `(x * 4) / 3`
// and `x * 0.75` gave, bit for bit.
export const ptToPx = (points: number): number => (points * 4) / 3;
export const pxToPt = (px: number): number => (px * 3) / 4;
const round = (value: number, digits: number): number => {
  const scale = 10 ** digits;
  return Math.round(value * scale) / scale;
};
/** A type size in points as CSS px, to `digits` decimals (hundredths by default). */
export const ptToPxRounded = (points: number, digits = 2): number => round(ptToPx(points), digits);
/** A type size in CSS px as points, to `digits` decimals (hundredths by default). */
export const pxToPtRounded = (px: number, digits = 2): number => round(pxToPt(px), digits);







export const WORD_LINE_MULTIPLE_MIN = 0.06;
export const WORD_LINE_MULTIPLE_MAX = 132;
export const isWordLineMultiple = (value: number): boolean =>
  Number.isFinite(value) && value >= WORD_LINE_MULTIPLE_MIN && value <= WORD_LINE_MULTIPLE_MAX;
