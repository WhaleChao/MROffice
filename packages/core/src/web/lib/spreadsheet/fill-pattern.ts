/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native Excel/WPS pattern fills.
 *
 * SpreadsheetML stores a patterned fill as an ST_PatternType plus two colors:
 * `fgColor` paints the marks and `bgColor` paints the cell underneath. Keeping
 * the pair in the workbook model lets copy/paste, autosave and XLSX patchback
 * retain the original semantics instead of flattening a hatch to one color.
 */

export const SPREADSHEET_FILL_PATTERN_TYPES = [
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625",
] as const;

/** Format Cells > Fill: the Automatic Pattern Color and the paper a pattern
 *  takes when no Background Color is picked. */
export const AUTOMATIC_PATTERN_COLOR = "#000000";
export const AUTOMATIC_PATTERN_BACKGROUND = "#ffffff";

export type SpreadsheetFillPatternType =
  (typeof SPREADSHEET_FILL_PATTERN_TYPES)[number];

export interface SpreadsheetFillPattern {
  type: SpreadsheetFillPatternType;
  foregroundColor: string;
  backgroundColor: string;
}

export interface SpreadsheetFillPatternCss {
  backgroundColor: string;
  backgroundImage: string;
  backgroundSize?: string;
  backgroundPosition?: string;
}

const PATTERN_TYPE_SET = new Set<string>(SPREADSHEET_FILL_PATTERN_TYPES);
const SAFE_HEX_COLOR = /^#[0-9a-f]{6}$/i;

export function isSpreadsheetFillPatternType(
  value: unknown,
): value is SpreadsheetFillPatternType {
  return typeof value === "string" && PATTERN_TYPE_SET.has(value);
}

export function normalizeSpreadsheetFillPattern(
  raw: unknown,
): SpreadsheetFillPattern | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (!isSpreadsheetFillPatternType(record.type)) return undefined;
  if (
    typeof record.foregroundColor !== "string" ||
    !SAFE_HEX_COLOR.test(record.foregroundColor) ||
    typeof record.backgroundColor !== "string" ||
    !SAFE_HEX_COLOR.test(record.backgroundColor)
  ) {
    return undefined;
  }
  return {
    type: record.type,
    foregroundColor: record.foregroundColor.toLowerCase(),
    backgroundColor: record.backgroundColor.toLowerCase(),
  };
}

/** The style property a cell's pattern carries its printed colour in. */
export const SHEET_PATTERN_PRINT_FILL = "--sheet-pattern-print-fill";

/** The share (of 64) of each pattern's 8 x 8 tile Windows Excel prints in
 * the pattern colour (lookwin p286-lw; the server's
 * _SHEET_FILL_PATTERN_INK): its Save as PDF tiles the GDI bitmap at 600 dpi,
 * 0.96pt a tile, which reads at print size as the two colours mixed. */
const PATTERN_PRINT_INK: Readonly<Record<string, number>> = {
  darkGray: 32, mediumGray: 16, lightGray: 8, gray125: 4, gray0625: 2,
  darkHorizontal: 32, darkVertical: 32, darkDown: 32, darkUp: 32, darkGrid: 32, darkTrellis: 48,
  lightHorizontal: 16, lightVertical: 16, lightDown: 16, lightUp: 16, lightGrid: 28, lightTrellis: 28,
};






export function spreadsheetFillPatternPrintColor(raw: unknown): string | undefined {
  const pattern = normalizeSpreadsheetFillPattern(raw);
  const ink = pattern ? PATTERN_PRINT_INK[pattern.type] : undefined;
  if (!pattern || ink === undefined) return undefined;
  const channel = (hex: string, at: number) => parseInt(hex.slice(1 + at * 2, 3 + at * 2), 16);
  return `#${[0, 1, 2]
    .map((at) => Math.round((channel(pattern.foregroundColor, at) * ink + channel(pattern.backgroundColor, at) * (64 - ink)) / 64).toString(16).padStart(2, "0"))
    .join("")}`;
}

/** The style property a conditional rule's pattern fill carries its marks'
 * colour and share of the tile in ("#rrggbb n", n of 64): the printed page
 * mixes them over the colour the cell shows between the marks -- the rule's
 * own background or the cell's fill -- as Excel prints the tile (R480-X41:
 * 2,582 gray0625 rule cells, each its own page-wide tiling pattern, one page
 * 4.5 GB to draw). */
export const SHEET_CF_PATTERN_PRINT_INK = "--sheet-cf-pattern-print-ink";

export function spreadsheetFillPatternPrintInk(type: unknown, foregroundColor: unknown): string | undefined {
  const ink = typeof type === "string" ? PATTERN_PRINT_INK[type] : undefined;
  if (ink === undefined || typeof foregroundColor !== "string" || !SAFE_HEX_COLOR.test(foregroundColor)) return undefined;
  return `${foregroundColor.toLowerCase()} ${ink}`;
}

/** The colour a printed tile reads as: `printInk` ("#rrggbb n") over `under`
 * (rgb channels), n of 64 in the marks' colour. */
export function spreadsheetFillPatternMixedPrintColor(printInk: string, under: readonly [number, number, number]): string | undefined {
  const match = /^(#[0-9a-f]{6}) (\d{1,2})$/i.exec(printInk.trim());
  if (!match?.[1] || !match[2]) return undefined;
  const marks = match[1];
  const ink = Number(match[2]);
  if (!(ink >= 0 && ink <= 64)) return undefined;
  return `#${[0, 1, 2]
    .map((at) => Math.round((parseInt(marks.slice(1 + at * 2, 3 + at * 2), 16) * ink + (under[at] ?? 255) * (64 - ink)) / 64).toString(16).padStart(2, "0"))
    .join("")}`;
}

function stripes(
  backgroundColor: string,
  foregroundColor: string,
  angle: number,
  lineWidth: number,
  gap: number,
): SpreadsheetFillPatternCss {
  return {
    backgroundColor,
    backgroundImage: `repeating-linear-gradient(${angle}deg, ${foregroundColor} 0 ${lineWidth}px, transparent ${lineWidth}px ${gap}px)`,
  };
}

function grid(
  backgroundColor: string,
  foregroundColor: string,
  firstAngle: number,
  secondAngle: number,
  lineWidth: number,
  gap: number,
): SpreadsheetFillPatternCss {
  const layer = (angle: number) =>
    `repeating-linear-gradient(${angle}deg, ${foregroundColor} 0 ${lineWidth}px, transparent ${lineWidth}px ${gap}px)`;
  return {
    backgroundColor,
    backgroundImage: `${layer(firstAngle)}, ${layer(secondAngle)}`,
  };
}

/**
 * Excel's thick diagonal stripes (darkDown / darkUp): 2 px of the marks, 2 px
 * of the cell, every 4 px along a row, one pixel further each row. A
 * repeating gradient of that period (2.83 px across the stripe) does not land
 * on pixels and drew a fine grey checker (PH122_issue.4248, X-240); one 8 px
 * square tile holding two stripes repeats on whole pixels, so every stripe
 * is drawn alike.
 */
function thickDiagonal(
  backgroundColor: string,
  foregroundColor: string,
  angle: number,
): SpreadsheetFillPatternCss {
  return {
    backgroundColor,
    backgroundImage: `linear-gradient(${angle}deg, ${foregroundColor} 25%, transparent 25% 50%, ${foregroundColor} 50% 75%, transparent 75%)`,
    backgroundSize: "8px 8px",
  };
}

/**
 * Render the native pattern without URLs or arbitrary CSS. Every value is
 * produced from an enum plus validated #rrggbb colors, so imported workbook
 * data cannot inject a CSS function.
 */
export function spreadsheetFillPatternCss(
  raw: unknown,
): SpreadsheetFillPatternCss | undefined {
  const pattern = normalizeSpreadsheetFillPattern(raw);
  if (!pattern) return undefined;
  const { type, foregroundColor: fg, backgroundColor: bg } = pattern;

  // Excel draws every pattern as an 8×8-pixel bitmap tiled over the cell
  // (XLS-A042): stripes repeat every 4 px, the grays are 1-px pixel grids.
  // The diagonal stripes keep the same 4-px horizontal period, which is
  // 4/√2 ≈ 2.83 px along the stripe's normal. "Down" runs from the top left
  // to the bottom right (\\\\) and "Up" from the bottom left to the top right
  // (////) — ST_PatternType; the two were swapped (XLS-A041). A CSS gradient
  // at 45deg points to the top right, so its stripes run "\\".
  const DIAGONAL_PERIOD = 2.828;
  switch (type) {
    case "darkHorizontal":
      return stripes(bg, fg, 0, 2, 4);
    case "lightHorizontal":
      return stripes(bg, fg, 0, 1, 4);
    case "darkVertical":
      return stripes(bg, fg, 90, 2, 4);
    case "lightVertical":
      return stripes(bg, fg, 90, 1, 4);
    case "darkDown":
      return thickDiagonal(bg, fg, 45);
    case "lightDown":
      return stripes(bg, fg, 45, 0.75, DIAGONAL_PERIOD);
    case "darkUp":
      return thickDiagonal(bg, fg, 135);
    case "lightUp":
      return stripes(bg, fg, 135, 0.75, DIAGONAL_PERIOD);
    case "darkGrid":
      // A 2-px checkerboard.
      return {
        backgroundColor: bg,
        backgroundImage: `conic-gradient(${fg} 0 25%, transparent 0 50%, ${fg} 0 75%, transparent 0)`,
        backgroundSize: "4px 4px",
      };
    case "lightGrid":
      return grid(bg, fg, 0, 90, 1, 4);
    case "darkTrellis":
      // Dense crossed diagonals leaving small holes.
      return grid(bg, fg, 45, 135, DIAGONAL_PERIOD / 2, DIAGONAL_PERIOD);
    case "lightTrellis":
      return grid(bg, fg, 45, 135, 0.6, DIAGONAL_PERIOD);
    case "mediumGray":
      // 50%: a 1-px checkerboard.
      return {
        backgroundColor: bg,
        backgroundImage: `conic-gradient(${fg} 0 25%, transparent 0 50%, ${fg} 0 75%, transparent 0)`,
        backgroundSize: "2px 2px",
      };
    case "darkGray":
      // 75%: the marks colour everywhere but one pixel in four.
      return {
        backgroundColor: bg,
        backgroundImage: `conic-gradient(from 270deg at 1px 1px, transparent 0 25%, ${fg} 0)`,
        backgroundSize: "2px 2px",
      };
    case "lightGray":
      // 25%: one pixel in four.
      return {
        backgroundColor: bg,
        backgroundImage: `conic-gradient(from 270deg at 1px 1px, ${fg} 0 25%, transparent 0)`,
        backgroundSize: "2px 2px",
      };
    case "gray125":
      // 12.5%: one pixel in eight.
      return {
        backgroundColor: bg,
        backgroundImage: `conic-gradient(from 270deg at 1px 1px, ${fg} 0 25%, transparent 0)`,
        backgroundSize: "4px 2px",
      };
    case "gray0625": {
      // 6.25%: two pixels in an 8 x 4 tile, staggered -- (4, 0) and (0, 2)
      // -- as Excel draws it (E14_OldIssue11573, X-202: sparse offset rows
      // of dots, not a square grid).
      // Each dot sits inside the tile (no layer offset), so a zoomed grid
      // (E14 opens at 75 %) scales it to a soft dot, not a half-pixel smear.
      const dot = (x: number, y: number) =>
        `radial-gradient(circle at ${x + 0.5}px ${y + 0.5}px, ${fg} 0 0.6px, transparent 0.85px)`;
      return {
        backgroundColor: bg,
        backgroundImage: `${dot(4, 0)}, ${dot(0, 2)}`,
        backgroundSize: "8px 4px, 8px 4px",
      };
    }
  }
}

/**
 * Native Excel/WPS GRADIENT fills.
 *
 * SpreadsheetML stores one as `<gradientFill degree="90">` with positioned
 * `<stop>` colors. The model kept only a single `backgroundColor`, so a header
 * band with a gradient arrived white — the cell looked empty where the workbook
 * had colour. `degree` is measured from left-to-right; CSS measures from
 * bottom-to-top, hence the quarter turn when this is normalized.
 */
export interface SpreadsheetGradientStop {
  /** 0..1 along the gradient axis. */
  position: number;
  color: string;
}

export interface SpreadsheetGradientFill {
  /** CSS degrees, already turned from the file's own convention. Ignored when
   * `path` is present — a path gradient radiates rather than running along an
   * axis. */
  angle: number;
  stops: SpreadsheetGradientStop[];
  /**
   * Excel's PATH gradient: colour radiating from an inner rectangle out to the
   * cell's edges, which is what "from the centre" fills are. Kept as the file's
   * own four insets rather than reduced to a centre point — CSS can only aim a
   * radial gradient at a point, but the workbook is entitled to get its
   * rectangle back unchanged. Absent means a plain linear gradient.
   */
  path?: { left: number; right: number; top: number; bottom: number };
}

export interface SpreadsheetGradientFillCss {
  backgroundColor: string;
  backgroundImage: string;
}

/** Excel allows more, but no real workbook needs a hundred; this bounds the CSS. */
export const MAX_GRADIENT_STOPS = 8;

export function normalizeSpreadsheetGradientFill(
  raw: unknown,
): SpreadsheetGradientFill | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (!Array.isArray(record.stops)) return undefined;
  const stops: SpreadsheetGradientStop[] = [];
  for (const entry of record.stops) {
    if (stops.length >= MAX_GRADIENT_STOPS) break;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const stop = entry as Record<string, unknown>;
    const color = stop.color;
    const position = stop.position;
    if (typeof color !== "string" || !SAFE_HEX_COLOR.test(color)) continue;
    if (typeof position !== "number" || !Number.isFinite(position)) continue;
    stops.push({
      position: Math.min(1, Math.max(0, position)),
      color: color.toLowerCase(),
    });
  }
  // One stop is a solid fill, which `backgroundColor` already says better.
  if (stops.length < 2) return undefined;
  const rawAngle = record.angle;
  const angle =
    typeof rawAngle === "number" && Number.isFinite(rawAngle)
      ? ((Math.round(rawAngle) % 360) + 360) % 360
      : 90;
  const rawPath = record.path;
  if (rawPath && typeof rawPath === "object" && !Array.isArray(rawPath)) {
    const insets = rawPath as Record<string, unknown>;
    const edges = ["left", "right", "top", "bottom"] as const;
    if (
      edges.every((edge) => {
        const value = insets[edge];
        return typeof value === "number" && Number.isFinite(value);
      })
    ) {
      const clamp = (edge: (typeof edges)[number]) =>
        Math.min(1, Math.max(0, insets[edge] as number));
      return {
        angle,
        stops,
        path: {
          left: clamp("left"),
          right: clamp("right"),
          top: clamp("top"),
          bottom: clamp("bottom"),
        },
      };
    }
  }
  return { angle, stops };
}

/**
 * Render the gradient without URLs or arbitrary CSS: every value is a number or
 * a validated #rrggbb, so imported workbook data cannot inject a CSS function.
 */
export function spreadsheetGradientFillCss(
  raw: unknown,
): SpreadsheetGradientFillCss | undefined {
  const gradient = normalizeSpreadsheetGradientFill(raw);
  if (!gradient) return undefined;
  const stops = gradient.stops
    .map((stop) => `${stop.color} ${Math.round(stop.position * 1000) / 10}%`)
    .join(", ");
  const percent = (value: number) => Math.round(value * 1000) / 10;
  // A path gradient radiates out to the corners, which is what
  // `farthest-corner` means; a linear one runs along an axis. CSS can only aim
  // at a point, so the inner rectangle is rendered from its centre. The four
  // values are the rectangle's EDGE POSITIONS from the cell's left / top
  // (ECMA-376 §18.8.24): Excel's From corner writes left=right=1 for the
  // top-right corner and From center 0.5 for all four (XLS-A061).
  const image = gradient.path
    ? `radial-gradient(farthest-corner at ` +
      `${percent((gradient.path.left + gradient.path.right) / 2)}% ` +
      `${percent((gradient.path.top + gradient.path.bottom) / 2)}%, ${stops})`
    : `linear-gradient(${gradient.angle}deg, ${stops})`;
  return {
    // The flat colour underneath is the first stop, so a colour filter and the
    // HTML fallback still see something rather than nothing.
    backgroundColor: gradient.stops[0].color,
    backgroundImage: image,
  };
}

// ── Fill Effects (Format Cells > Fill > Fill Effects…, XLS-A061) ──────────────

export type FillEffectsShading =
  | "horizontal"
  | "vertical"
  | "diagonalUp"
  | "diagonalDown"
  | "fromCorner"
  | "fromCenter";

export const FILL_EFFECTS_SHADINGS: readonly FillEffectsShading[] = [
  "horizontal",
  "vertical",
  "diagonalUp",
  "diagonalDown",
  "fromCorner",
  "fromCenter",
];

/** How many Variants Excel offers for a shading style. */
export function fillEffectsVariantCount(shading: FillEffectsShading): number {
  return shading === "fromCorner" ? 4 : shading === "fromCenter" ? 2 : 3;
}

/** The SpreadsheetML degree of each linear shading (0 = left to right). */
const LINEAR_DEGREES: Record<"horizontal" | "vertical" | "diagonalUp" | "diagonalDown", number> = {
  horizontal: 90,
  vertical: 0,
  diagonalUp: 45,
  diagonalDown: 135,
};

/**
 * The two-colour gradient Excel's Fill Effects writes. Linear styles: variant
 * 0 runs colour 1 → colour 2, 1 the reverse, 2 colour 1 → 2 → 1. From corner:
 * variants 0-3 radiate from the top-left, top-right, bottom-left and
 * bottom-right corner. From center: 0 has colour 1 in the middle, 1 colour 2.
 * `angle` is in the model's CSS degrees (the file's degree + 90).
 */
export function fillEffectsGradient(
  shading: FillEffectsShading,
  variant: number,
  color1: string,
  color2: string,
): SpreadsheetGradientFill {
  const c1 = color1.toLowerCase();
  const c2 = color2.toLowerCase();
  if (shading === "fromCorner") {
    const corners = [
      { left: 0, right: 0, top: 0, bottom: 0 },
      { left: 1, right: 1, top: 0, bottom: 0 },
      { left: 0, right: 0, top: 1, bottom: 1 },
      { left: 1, right: 1, top: 1, bottom: 1 },
    ];
    return {
      angle: 90,
      path: corners[Math.max(0, Math.min(3, variant))]!,
      stops: [{ position: 0, color: c1 }, { position: 1, color: c2 }],
    };
  }
  if (shading === "fromCenter") {
    const [inner, outer] = variant === 1 ? [c2, c1] : [c1, c2];
    return {
      angle: 90,
      path: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 },
      stops: [{ position: 0, color: inner }, { position: 1, color: outer }],
    };
  }
  const angle = (LINEAR_DEGREES[shading] + 90) % 360;
  if (variant === 2) {
    return {
      angle,
      stops: [
        { position: 0, color: c1 },
        { position: 0.5, color: c2 },
        { position: 1, color: c1 },
      ],
    };
  }
  const [from, to] = variant === 1 ? [c2, c1] : [c1, c2];
  return { angle, stops: [{ position: 0, color: from }, { position: 1, color: to }] };
}

/** The Fill Effects choice that writes `raw`, or null when Excel's dialog
 * could not have written it (more stops, another angle or rectangle). */
export function fillEffectsFromGradient(raw: unknown): {
  shading: FillEffectsShading;
  variant: number;
  color1: string;
  color2: string;
} | null {
  const gradient = normalizeSpreadsheetGradientFill(raw);
  if (!gradient) return null;
  const first = gradient.stops[0]!.color;
  const second = gradient.stops[1]!.color;
  const target = JSON.stringify(gradient);
  for (const shading of FILL_EFFECTS_SHADINGS) {
    for (let variant = 0; variant < fillEffectsVariantCount(shading); variant += 1) {
      for (const [color1, color2] of [[first, second], [second, first]] as const) {
        const candidate = normalizeSpreadsheetGradientFill(
          fillEffectsGradient(shading, variant, color1, color2),
        );
        if (JSON.stringify(candidate) === target) return { shading, variant, color1, color2 };
      }
    }
  }
  return null;
}
