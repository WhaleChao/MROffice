/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The exact dash a shape's outline was drawn with in Office, beyond the three
 * every host's `outlineStyle` names.
 *
 * `a:prstDash@val` (ECMA-376 20.1.10.48 ST_PresetLineDashVal) has eleven
 * presets; `outlineStyle` holds the family -- "solid", "dashed" (`dash`),
 * "dotted" (`dot`) -- and `outlineDash` the preset when the file states one of
 * the other eight: Format Shape > Line > Dash type's Round Dot (`sysDot`),
 * Square Dot (`sysDash`), Dash Dot, Long Dash, Long Dash Dot, Long Dash Dot
 * Dot, and the two `sys*` dash-dots. A dash is only read with the family it
 * belongs to: once the style is changed to another family it says nothing.
 * Keep in step with the server's `office_shape_paint.STROKE_DASH_FAMILY`.
 */

export const OFFICE_STROKE_DASH_FAMILY = {
  sysDot: "dotted",
  sysDash: "dashed",
  dashDot: "dashed",
  lgDash: "dashed",
  lgDashDot: "dashed",
  lgDashDotDot: "dashed",
  sysDashDot: "dashed",
  sysDashDotDot: "dashed",
} as const;

export type OfficeStrokeDash = keyof typeof OFFICE_STROKE_DASH_FAMILY;

/** The preset `value` names when it refines `style`; otherwise nothing. */
export function officeStrokeDash(
  style: unknown,
  value: unknown,
): OfficeStrokeDash | null {
  return typeof value === "string" &&
    Object.hasOwn(OFFICE_STROKE_DASH_FAMILY, value) &&
    OFFICE_STROKE_DASH_FAMILY[value as OfficeStrokeDash] === style
    ? (value as OfficeStrokeDash)
    : null;
}

/**
 * Each preset's dash and gap lengths in LINE WIDTHS, as DrawingML defines
 * them (the lengths Office's renderers and `a:custDash` equivalents use:
 * `sysDot` 1-1, `sysDash` 3-1, `dashDot` 4-3-1-3, `lgDash` 8-3, ...).
 */
const OFFICE_STROKE_DASH_UNITS: Record<OfficeStrokeDash, readonly number[]> = {
  sysDot: [1, 1],
  sysDash: [3, 1],
  dashDot: [4, 3, 1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/**
 * Any `a:prstDash@val` preset's dash and gap lengths in line widths (the
 * family names' too: `dash` 4-3, `dot` 1-3); null for `solid` or a name
 * that is not a preset.
 */
export function officeDashUnits(preset: unknown): readonly number[] | null {
  if (preset === "dash") return [4, 3];
  if (preset === "dot") return [1, 3];
  return typeof preset === "string" && Object.hasOwn(OFFICE_STROKE_DASH_UNITS, preset)
    ? OFFICE_STROKE_DASH_UNITS[preset as OfficeStrokeDash]
    : null;
}

/**
 * A dash pattern's lengths (in line widths) at a width, with the line's caps
 * counted inside each dash as Office counts them: a round or square cap
 * reaches half the width past each end, so each dash is drawn a width
 * shorter (a round dot is a dash of no length) and each gap a width longer
 * (Excel for Mac, X-223 oracle: a 2pt round Dash prints 3pt of line between
 * its caps every 7 widths, a round System Dot a dot every two widths).
 */
export function officeDashLengthsWithCaps(
  units: readonly number[],
  width: number,
  capped: boolean,
): number[] {
  return units.map((length, index) =>
    !capped ? length * width : index % 2 === 0 ? Math.max(0.01, (length - 1) * width) : (length + 1) * width,
  );
}

/** The SVG `stroke-dasharray` of one preset at a stroke width (same units). */
export function officeStrokeDashArray(
  dash: OfficeStrokeDash,
  strokeWidth: number,
  /** Drawn with round or square caps, which Office counts inside each dash. */
  capped = false,
): string {
  const width = Number.isFinite(strokeWidth) && strokeWidth > 0 ? strokeWidth : 1;
  return officeDashLengthsWithCaps(OFFICE_STROKE_DASH_UNITS[dash], width, capped)
    .map((length) => Math.round(length * 100) / 100)
    .join(" ");
}










export function officeLineFamilyDash(
  style: unknown,
  strokeWidth: number,
): { readonly dashArray: string; readonly linecap: "round" | "butt" } | null {
  const width = Number.isFinite(strokeWidth) && strokeWidth > 0 ? strokeWidth : 1;
  const length = (units: number) => Math.round(units * width * 100) / 100;
  if (style === "dashed") return { dashArray: `${length(4)} ${length(3)}`, linecap: "butt" };
  if (style === "dotted") return { dashArray: `0 ${length(2)}`, linecap: "round" };
  return null;
}

/**
 * Shape Outline > Dashes: Office's eight, in its order (Solid, Round Dot,
 * Square Dot, Dash, Dash Dot, Long Dash, Long Dash Dot, Long Dash Dot Dot),
 * by DrawingML name (the names the shared dash icons draw).
 */
export const OFFICE_MENU_DASHES = [
  "solid",
  "sysDot",
  "sysDash",
  "dash",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
] as const;
export type OfficeMenuDash = (typeof OFFICE_MENU_DASHES)[number];

/**
 * One Dashes pick in a shape's fields: the family (`strokeStyle` /
 * `outlineStyle`), the exact dash refining it, and the round cap Round Dot
 * is (`sysDot` with `cap="rnd"`, as PowerPoint writes it). Dash is the dashed
 * family's own look, so it states no refinement.
 */
export function officeMenuDashFields(dash: OfficeMenuDash): {
  readonly style: "solid" | "dashed" | "dotted";
  readonly dash?: OfficeStrokeDash;
  readonly cap?: "round";
} {
  switch (dash) {
    case "solid":
      return { style: "solid" };
    case "sysDot":
      return { style: "dotted", dash: "sysDot", cap: "round" };
    case "dash":
      return { style: "dashed" };
    default:
      return { style: "dashed", dash };
  }
}

/** The Dashes entry a shape's line shows checked, or null for a dash the menu does not list. */
export function officeMenuDashOf(style: unknown, dash: unknown): OfficeMenuDash | null {
  if (style === "dotted") return "sysDot";
  if (style !== "dashed") return "solid";
  const exact = officeStrokeDash(style, dash);
  if (!exact) return "dash";
  return (OFFICE_MENU_DASHES as readonly string[]).includes(exact) ? (exact as OfficeMenuDash) : null;
}

/** Every dash a line's `outlineDash` may name: the menu's eight and the two
 * system presets Office reads but does not list. */
const OFFICE_OUTLINE_DASHES: ReadonlySet<string> = new Set<string>([
  ...OFFICE_MENU_DASHES,
  ...Object.keys(OFFICE_STROKE_DASH_FAMILY),
]);

export function normalizeOfficeOutlineDash(value: unknown): OfficeMenuDash | OfficeStrokeDash | undefined {
  return typeof value === "string" && OFFICE_OUTLINE_DASHES.has(value)
    ? (value as OfficeMenuDash | OfficeStrokeDash)
    : undefined;
}

/**
 * How a line stating an exact `outlineDash` (and maybe a cap) is drawn at a
 * width in px: each preset in line widths; Round Dot (sysDot with a round
 * cap) the dotted family's round dots; Dash the dashed family's. Null for a
 * solid line or no dash.
 */
export function officeOutlineDashPaint(
  dash: unknown,
  cap: unknown,
  widthPx: number,
): { readonly dashArray: string; readonly linecap: "round" | "butt" | "square" } | null {
  const value = normalizeOfficeOutlineDash(dash);
  if (!value || value === "solid") return null;
  const linecap = cap === "round" ? "round" : cap === "square" ? "square" : "butt";
  if (value === "dash") return officeLineFamilyDash("dashed", widthPx);
  if (value === "sysDot" && cap === "round") return officeLineFamilyDash("dotted", widthPx);
  return { dashArray: officeStrokeDashArray(value as OfficeStrokeDash, widthPx), linecap };
}
