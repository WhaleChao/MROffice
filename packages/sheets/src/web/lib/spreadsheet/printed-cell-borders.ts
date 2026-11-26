/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { AUTOMATIC_BORDER_COLOR, normalizeBorderEdges, type BorderLineStyle } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";

export const PRINT_DOT_PT = 72 / 600;

/** The frame printed gridlines close a page's block in: eight printer dots
 * at any scale, centred on the block's edges and over the borders and fills
 * there (Windows Excel, lookwin p91-lw: a red box on the print area's edge
 * prints black). The API's PRINTED_GRID_FRAME_PT is the same. */
export const PRINTED_GRID_FRAME_PT = 8 * PRINT_DOT_PT;

/** Where the printed grid starts past the margin: half the frame, four
 * printer dots, with or without gridlines and headings (lookwin p91-lw). The
 * API's PRINTED_GRID_INSET_PT is the same. */
export const PRINTED_GRID_INSET_PT = 4 * PRINT_DOT_PT;

/** How far a page's printed block (headings and cells, `block` points) starts
 * from its margin in a `room`-point printable width (or height): four printer
 * dots in, or centred four dots short of the centre when Center on Page asks
 * for it and leaves more than the eight dots spare (lookwin p91-lw: every
 * line 0.45-0.5 pt right of and below ours at 50-100%; centred, as far left
 * of and above). The API's writer (`centred_x` / `centred_y`) is the same. */
/**
 * How far right of its left margin a printed sheet starts so as to stand no
 * nearer the paper's left edge than eight printer dots, points: Windows Excel
 * prints a left margin of 0, 0.005 or 0.01 in where one of 8 dots (0.96 pt)
 * does, and 0.02 in (12 dots) where it says. The top margin has no such
 * floor, and the room Fit To fits in is the stated margins' (lookwin p189-lw,
 * p190-lw; R47-X14, left margin 0: every page printed 8 dots left of
 * Excel's). A centred block (`room`, `block` and `centred` as
 * printedGridOffset takes them) is centred in what is left of the width, half
 * as far right (R48-X01 and R77-X23, margins 0 and centred: 4 dots). The
 * API's writer the same (notebooks.py PRINTED_LEFT_EDGE_FLOOR_PT).
 */
export function printedLeftEdgeFloor(marginLeftPt: number, room: number, block: number, centred: boolean): number {
  const floor = Math.max(0, 8 * PRINT_DOT_PT - marginLeftPt);
  return centred && room - block > 2 * PRINTED_GRID_INSET_PT ? floor / 2 : floor;
}

export function printedGridOffset(room: number, block: number, centred: boolean): number {
  return centred && room - block > 2 * PRINTED_GRID_INSET_PT ? (room - block) / 2 - PRINTED_GRID_INSET_PT : PRINTED_GRID_INSET_PT;
}

export interface PrintedBorderLine {
  /** Stroke widths in printer dots before the page's scale (one stroke, or two for double). */
  readonly strokes: readonly number[];
  /** Gap between the two strokes of a double line, dots. */
  readonly gap: number;
  /** Dash runs in dots (on, off, ...), or null for a solid line. */
  readonly dash: readonly number[] | null;
  readonly color: string;
  /** Printed at its width at any scale: a hairline is one dot at 35% and at 200%. */
  readonly fixedWidth?: boolean;
}

const CSS_BORDER = /^\s*([\d.]+)px\s+(solid|double|dashed|dotted)\s+(.+?)\s*$/;

/**
 * The Excel line a cell's CSS border spells (resolveBorderCss / gridBorderCss):
 * 1px solid or double is thin, 2px solid medium, 3px solid thick, 3px double
 * double, 1px dashed dashed, 2px dashed medium dashed, 1px dotted dotted.
 * Undefined for anything else (no border, a transparent edge whose dash-dot
 * pattern is painted as a background strip, a gridline): the page keeps
 * drawing those as it did.
 */
export function printedBorderLine(css: string | undefined): PrintedBorderLine | undefined {
  if (typeof css !== "string") return undefined;
  const match = CSS_BORDER.exec(css);
  if (!match) return undefined;
  const width = Number(match[1]);
  const line = match[2];
  const color = match[3];
  if (!Number.isFinite(width) || width <= 0 || /transparent|var\(/.test(color)) return undefined;
  if (line === "double" && width >= 3) return { strokes: [8, 8], gap: 8, dash: null, color };
  if (line === "solid" || line === "double") {
    return { strokes: [width >= 3 ? 24 : width >= 2 ? 16 : 8], gap: 0, dash: null, color };
  }
  if (line === "dashed") return width >= 2
    ? { strokes: [16], gap: 0, dash: [72, 24], color }
    : { strokes: [8], gap: 0, dash: [24, 8], color };
  if (line === "dotted" && width < 2) return { strokes: [8], gap: 0, dash: [4, 4], color };
  return undefined;
}

/** Each Excel line style as printed, in dots (lookwin probe p10). */
const PRINTED_LINES: Record<BorderLineStyle, Omit<PrintedBorderLine, "color">> = {
  hair: { strokes: [1], gap: 0, dash: null, fixedWidth: true },
  thin: { strokes: [8], gap: 0, dash: null },
  medium: { strokes: [16], gap: 0, dash: null },
  thick: { strokes: [24], gap: 0, dash: null },
  double: { strokes: [8, 8], gap: 8, dash: null },
  dotted: { strokes: [8], gap: 0, dash: [4, 4] },
  dashed: { strokes: [8], gap: 0, dash: [24, 8] },
  dashDot: { strokes: [8], gap: 0, dash: [72, 24, 24, 24] },
  dashDotDot: { strokes: [8], gap: 0, dash: [72, 24, 24, 24, 24, 24] },
  mediumDashed: { strokes: [16], gap: 0, dash: [72, 24] },
  mediumDashDot: { strokes: [16], gap: 0, dash: [72, 24, 24, 24] },
  mediumDashDotDot: { strokes: [16], gap: 0, dash: [72, 24, 24, 24, 24, 24] },
  slantDashDot: { strokes: [16], gap: 0, dash: [72, 24, 24, 24] },
};

/** The line styles the grid paints as a background strip (its dash-dot
 * rhythm, a hairline's dots), not as a CSS border: a printed page draws them
 * as vector lines instead (printedEdgeVariables). */
const STRIP_STYLES: ReadonlySet<BorderLineStyle> = new Set([
  "hair", "dotted", "dashDot", "dashDotDot", "mediumDashDot", "mediumDashDotDot", "slantDashDot",
]);

/** The CSS custom property a printed cell carries for an edge drawn as a vector line. */
export function printedEdgeVariable(side: "top" | "right" | "bottom" | "left"): `--printed-border-${typeof side}` {
  return `--printed-border-${side}`;
}

/**
 * For a printed cell: its border style with the strip-painted edges taken
 * out, and those edges as CSS variables ("<style> <color>") the page draws as
 * vector lines. Undefined when the style has no such edge.
 */
export function printedStripEdges<T extends { borderEdges?: unknown }>(style: T | undefined,
  options?: { diagonal?: boolean }):
  { style: T; variables: Record<string, string> } | undefined {
  const edges = normalizeBorderEdges(style?.borderEdges);
  if (!style || !edges) return undefined;
  const variables: Record<string, string> = {};
  const kept: Record<string, unknown> = { ...(style.borderEdges as Record<string, unknown>) };
  for (const side of ["top", "right", "bottom", "left"] as const) {
    const edge = edges[side];
    if (!edge || !STRIP_STYLES.has(edge.style)) continue;
    variables[printedEdgeVariable(side)] = `${edge.style} ${edge.color ?? AUTOMATIC_BORDER_COLOR}`;
    delete kept[side];
  }
  // A diagonal is a background band on screen, one CSS px at any scale
  // (R121-X08 printed 0.37 pt where Excel prints its 68% thin line):
  // printed, it is a vector line in printer dots too (the API's
  // draw_border_line), corner to corner of the cell or of a whole merge.
  const diagonal = options?.diagonal ? edges.diagonal : undefined;
  if (diagonal && PRINTED_LINES[diagonal.style]) {
    const ways = diagonal.up && diagonal.down ? "cross" : diagonal.up ? "up" : "down";
    variables[PRINTED_DIAGONAL_VARIABLE] = `${diagonal.style} ${diagonal.color ?? AUTOMATIC_BORDER_COLOR} ${ways}`;
    delete kept.diagonal;
  }
  return Object.keys(variables).length ? { style: { ...style, borderEdges: kept }, variables } : undefined;
}

/** The CSS custom property a printed cell carries its diagonal in:
 * "<style> <color> <up|down|cross>". */
export const PRINTED_DIAGONAL_VARIABLE = "--printed-border-diagonal";

/** The diagonal lines a printed cell's diagonal variable asks for over its
 * `box` (page CSS px): down runs top left to bottom right (the split-header
 * form, and what a diagonal stating neither direction reads as), up bottom
 * left to top right; a cross is both. */
export function printedDiagonalEdges(value: unknown,
  box: { left: number; top: number; width: number; height: number }): PrintedEdge[] {
  if (typeof value !== "string") return [];
  const [style, color, ways] = value.trim().split(/\s+/, 3);
  const line = printedBorderLineOfVariable(`${style} ${color}`);
  if (!line || !(box.width > 0) || !(box.height > 0)) return [];
  const { left, top, width, height } = box;
  const out: PrintedEdge[] = [];
  if (ways === "down" || ways === "cross") out.push({ x1: left, y1: top, x2: left + width, y2: top + height, line });
  if (ways === "up" || ways === "cross") out.push({ x1: left, y1: top + height, x2: left + width, y2: top, line });
  return out;
}

/** Whether a printed edge is a diagonal (neither horizontal nor vertical). */
export function isPrintedDiagonal(edge: PrintedEdge): boolean {
  return edge.x1 !== edge.x2 && edge.y1 !== edge.y2;
}

/** A printed diagonal as one stroke along the cell's corners, its width the
 * line's first stroke in dots at the page's scale (a double diagonal prints
 * its two strokes on the one line, as the API's writer does). */
export function printedDiagonalStroke(edge: PrintedEdge, scale: number, pxPerPt: number): {
  readonly width: number; readonly dash: string | null; readonly color: string;
} {
  const dots = edge.line.strokes[0] ?? 8;
  const width = printedDotsPt(dots, edge.line.fixedWidth ? 1 : scale) * pxPerPt;
  const dash = edge.line.dash ? edge.line.dash.map((value) => printedDotsPt(value, scale) * pxPerPt).join(" ") : null;
  return { width, dash, color: edge.line.color };
}

/** The printed line of a "<style> <color>" edge variable. */
export function printedBorderLineOfVariable(value: unknown): PrintedBorderLine | undefined {
  if (typeof value !== "string") return undefined;
  const [style, color] = value.trim().split(/\s+/, 2);
  const line = PRINTED_LINES[style as BorderLineStyle];
  return line && color ? { ...line, color } : undefined;
}

/** A length in printer dots at the page's scale, in points: whole dots, never under one. */
export function printedDotsPt(dots: number, scale: number): number {
  return Math.max(1, Math.round(dots * scale)) * PRINT_DOT_PT;
}

export interface PrintedEdge {
  /** The grid line the edge lies on, in page CSS px: a horizontal edge from (x1, y) to (x2, y). */
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
  readonly line: PrintedBorderLine;
}

/**
 * A printed line's whole printer dots across, and how many of them stand
 * before its grid line (above a horizontal one, left of a vertical one).
 * Windows Excel's printer pens are whole dots on the grid's own dot: an even
 * line is centred on it, an odd one keeps its middle dot after it, as a
 * one-dot pen covers the dot the grid line starts (lookwin p11-win: a
 * column's right border at 35 % prints hair [E, E+1] dots, thin 3 dots
 * [E-1, E+2], medium 6 and thick 8 centred; at 100 % hair [E, E+1], thin 8,
 * medium 16, thick 24 centred). In a right-to-left sheet a vertical line's
 * odd dot goes the other way, as its gridlines do. A double line is its
 * whole width in dots (24 at the page's scale) less its two strokes: at
 * 35 % two 3-dot strokes 2 dots apart, 8 in all.
 */
export function printedLineDots(line: PrintedBorderLine, scale: number): {
  readonly strokes: readonly number[]; readonly gap: number; readonly total: number; readonly before: number;
} {
  const strokeScale = line.fixedWidth ? 1 : scale;
  const strokes = line.strokes.map((dots) => Math.max(1, Math.round(dots * strokeScale)));
  const strokeDots = strokes.reduce((sum, value) => sum + value, 0);
  const gaps = strokes.length - 1;
  const gap = line.gap && gaps > 0
    ? Math.max(1, Math.round((Math.round((line.strokes.reduce((sum, value) => sum + value, 0) + line.gap * gaps) * scale) - strokeDots) / gaps))
    : 0;
  const total = strokeDots + gap * gaps;
  return { strokes, gap, total, before: Math.floor(total / 2) };
}

/**
 * The strokes of one printed edge as rectangles (solid) or dashed lines, in
 * page CSS px: on the grid line in whole printer dots (printedLineDots), each
 * end running on by the line's width before and after it so two edges that
 * meet at a corner close it.
 * `scale` is the page's content scale, `pxPerPt` its CSS px per point.
 */
export function printedEdgeStrokes(edge: PrintedEdge, scale: number, pxPerPt: number, rightToLeft = false): {
  readonly x: number; readonly y: number; readonly width: number; readonly height: number;
  readonly dash: string | null; readonly color: string;
}[] {
  const horizontal = edge.y1 === edge.y2;
  const dot = PRINT_DOT_PT * pxPerPt;
  const dots = printedLineDots(edge.line, scale);
  const widths = dots.strokes.map((value) => value * dot);
  const gap = dots.gap * dot;
  const total = dots.total * dot;
  // Along the line its ends run on by what stands before and after the grid
  // line across; a vertical line in a right-to-left sheet is mirrored.
  const before = (!horizontal && rightToLeft ? dots.total - dots.before : dots.before) * dot;
  const dash = edge.line.dash ? edge.line.dash.map((value) => printedDotsPt(value, scale) * pxPerPt).join(" ") : null;
  const out: { x: number; y: number; width: number; height: number; dash: string | null; color: string }[] = [];
  let offset = -before;
  for (const width of widths) {
    const run = horizontal
      ? { x: Math.min(edge.x1, edge.x2) - (rightToLeft ? total - before : before), y: edge.y1 + offset, width: Math.abs(edge.x2 - edge.x1) + total, height: width }
      : { x: edge.x1 + offset, y: Math.min(edge.y1, edge.y2) - dots.before * dot, width, height: Math.abs(edge.y2 - edge.y1) + total };
    out.push({ ...run, dash, color: edge.line.color });
    offset += width + gap;
  }
  return out;
}
