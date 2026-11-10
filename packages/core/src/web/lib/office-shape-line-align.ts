/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * `a:ln@algn` — Stroke Alignment — for all four hosts, in one place.
 *
 * ECMA-376's `ST_PenAlignment` (§20.1.10.39) has exactly TWO members, and
 * Microsoft's own enumeration names them: `ctr` = "Center Alignment",
 * `in` = "Inset Alignment" (`DocumentFormat.OpenXml.Drawing.PenAlignmentValues`).
 * Office's object model carries the same fact as a boolean —
 * `LineFormat.InsetPen`, "whether lines are drawn inside the specified shape's
 * boundaries" — so `in` is the state in which the WHOLE stroke falls inside
 * the geometry path and the drawn object does not grow by half a line width.
 *
 * The `ln` clause states NO default for `algn` (unlike `cap` → `sq`,
 * `cmpd` → `sng` and `w` → `0`, each of which the same attribute table spells
 * out). So an absent attribute is not "`ctr` stated": it is a file that says
 * nothing, and the page carries no field for it. That is what keeps an
 * untouched save byte-identical.
 *
 * ## Why this is not "shrink the path by half the width"
 *
 * SHAPE-IMPORT-AUDIT.md §59 established the other half of the rule: the
 * outline is drawn ON the geometry path, and the box is not the boundary of
 * the ink — so a centred pen legitimately puts `w/2` of ink outside the frame,
 * and no host's frame may clip it. An inset pen is the exception the FILE asks
 * for, and it is still not a geometry change: the fill boundary stays exactly
 * where it was, and only the ink moves. Insetting the PATH would move the fill
 * too, so two shapes that differ only in line width would stop lining up.
 *
 * SVG has no `stroke-alignment`, so the mechanic here is the standard one:
 * paint the same path with DOUBLE the width and clip it to the path's own fill
 * region. The outer half is removed, the inner half lands exactly where Office
 * puts it, and the fill is untouched.
 */

/** The page's word for `a:ln@algn`. */
export type OfficeStrokeAlign = "center" | "inset";

/** The page word → the `algn` attribute value. Keep in step with the server's
 * `office_shape_paint.STROKE_ALIGNMENT_ATTRIBUTES`. */
const OFFICE_STROKE_ALIGN_ATTRIBUTES = {
  center: "ctr",
  inset: "in",
} as const satisfies Record<OfficeStrokeAlign, string>;

/**
 * How a line with NO stated alignment is drawn. Not a schema default — there
 * is none — but the reading the enumeration and `InsetPen`'s `msoFalse` give.
 */
export const OFFICE_DEFAULT_STROKE_ALIGN: OfficeStrokeAlign = "center";

/** One unknown value is never guessed at: it reads as "states nothing". */
export function officeStrokeAlign(value: unknown): OfficeStrokeAlign | null {
  return value === "center" || value === "inset" ? value : null;
}

/** Whether this line's ink belongs entirely inside the geometry path. */
export function isOfficeInsetStroke(value: unknown): boolean {
  return officeStrokeAlign(value) === "inset";
}

/** The `algn` attribute one page word asks for, or `""` for "write none". */
export function officeStrokeAlignAttribute(value: unknown): string {
  const align = officeStrokeAlign(value);
  return align ? OFFICE_STROKE_ALIGN_ATTRIBUTES[align] : "";
}

/**
 * The width to paint with, so that HALF of it survives the clip and lands
 * inside the path. A centred pen keeps its own width.
 */
export function officeStrokeAlignPaintWidth(
  strokeWidth: number,
  align: unknown,
): number {
  return isOfficeInsetStroke(align) && strokeWidth > 0
    ? strokeWidth * 2
    : strokeWidth;
}

/**
 * The `clipPath` id an inset stroke needs, or `null` for a centred one. One
 * naming rule so a host's `<clipPath>` and the reference to it cannot drift.
 */
export function officeStrokeAlignClipId(
  blockId: string,
  align: unknown,
  suffix?: string | number,
): string | null {
  if (!isOfficeInsetStroke(align)) return null;
  return `office-inset-stroke-${blockId}${suffix === undefined ? "" : `-${suffix}`}`;
}
