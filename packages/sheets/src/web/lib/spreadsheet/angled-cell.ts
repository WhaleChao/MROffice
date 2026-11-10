/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PX_PER_PT } from "@mrnote/core/web/lib/office-units.ts";
import { normalizeTextRotation } from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";

/** The cell's angle in degrees (+ rising, - falling) when its text is angled,
 * else null (level, quarter turn, stacked). */
export function spreadsheetAngledCellAngle(
  style: { textRotation?: unknown; verticalText?: unknown } | undefined,
): number | null {
  if (!style || style.verticalText === true) return null;
  const angle = normalizeTextRotation(style.textRotation);
  return angle === 0 || Math.abs(angle) === 90 ? null : angle;
}

export type AngledCellMode = "band" | "rect";

/** How an angled cell is drawn: a band when it carries a border of its own and
 * is not merged, else an upright rectangle; null for level text. */
export function spreadsheetAngledCellMode(
  style: { textRotation?: unknown; verticalText?: unknown } | undefined,
  options: { merged: boolean; ownBorder: boolean },
): AngledCellMode | null {
  if (spreadsheetAngledCellAngle(style) === null) return null;
  return options.ownBorder && !options.merged ? "band" : "rect";
}

/**
 * The horizontal alignment angled text is placed by. A stated one is kept;
 * General puts text that rises at the left and text that falls at the right
 * (measured: a General -45 cell starts where an explicit right one does).
 * `resolved` is what General gives the value level (numbers right).
 */
export function angledCellHorizontalAlign(
  stated: unknown,
  angle: number,
  resolved: "left" | "center" | "right",
): "left" | "center" | "right" {
  if (stated === "left" || stated === "center" || stated === "right") return stated;
  if (typeof stated === "string" && stated && stated !== "general") return resolved;
  return resolved === "left" && angle < 0 ? "right" : resolved;
}

export interface AngledCellBand {
  angle: number;
  /** Horizontal offset of the top edge against the bottom edge, px (+ right). */
  shift: number;
  /** The band's horizontal extent relative to the cell's left edge. */
  left: number;
  width: number;
  height: number;
  /** Corners, cell coordinates: bottom-left, bottom-right, top-right, top-left. */
  corners: readonly [readonly [number, number], readonly [number, number], readonly [number, number], readonly [number, number]];
}

export function spreadsheetAngledCellBand(angle: number, width: number, height: number): AngledCellBand {
  const shift = angledCellShift(angle, height);
  return {
    angle,
    shift,
    left: Math.min(0, shift),
    width: width + Math.abs(shift),
    height,
    corners: [[0, height], [width, height], [width + shift, 0], [shift, 0]],
  };
}

/** How far a band's top edge sits from its bottom edge, px (+ right). */
export function angledCellShift(angle: number, height: number): number {
  const theta = (Math.abs(angle) * Math.PI) / 180;
  return (angle > 0 ? 1 : -1) * (height / Math.tan(theta));
}

/** The band as a CSS clip-path for a box that starts at `band.left`. */
export function angledCellClipPath(band: AngledCellBand): string {
  return `polygon(${band.corners.map(([x, y]) => `${(x - band.left).toFixed(2)}px ${y.toFixed(2)}px`).join(", ")})`;
}

/**
 * The top border segments a band cell draws, in its own coordinates: its own
 * (between the top ends of its left and right edges) and, for a level
 * neighbour with a top border, the stretch that neighbour's top runs out to
 * over this cell. `rightShift` is the right-hand neighbour's shift when that
 * neighbour is a band too (its slant draws the shared edge), else null.
 */
export function angledCellTopSegments(
  band: AngledCellBand,
  width: number,
  rightShift: number | null,
): { own: readonly [number, number]; leftNeighbour: readonly [number, number] | null; rightNeighbour: readonly [number, number] | null } {
  const leftTop = band.shift;
  const rightTop = width + (rightShift ?? band.shift);
  return {
    own: [Math.min(leftTop, rightTop), Math.max(leftTop, rightTop)],
    // A level cell on the left reaches right to this cell's left edge top.
    leftNeighbour: leftTop > 0 ? [0, leftTop] : null,
    // A level cell on the right reaches left to this cell's right edge top
    // (this cell's own slant draws that edge when the neighbour is level).
    rightNeighbour: rightShift === null && rightTop < width ? [rightTop, width] : null,
  };
}

/** px per point at the grid's 96 dpi. */
const PT = PX_PER_PT;
// Line box of the turned text, as a share of the font size (Excel's GDI
// cell metrics, fitted: Calibri 11 ascent 10pt, descent 3pt).
const ASCENT = 10 / 11;
const DESCENT = 3 / 11;
const RECT_PAD = { left: 2.5 * PT, right: 3.7 * PT, bottom: 2 * PT, top: 2.9 * PT };
const BAND_PAD = {
  rising: { left: 1.6 * PT, right: 5 * PT },
  falling: { left: 3 * PT, right: 3.6 * PT },
  bottom: 1.2 * PT,
  top: 2.9 * PT,
};

export interface AngledTextInput {
  mode: AngledCellMode;
  angle: number;
  /** The cell's (spanned) width and height, px. */
  width: number;
  height: number;
  /** Text advance and font size, px. */
  textWidth: number;
  fontSize: number;
  align: "left" | "center" | "right";
  verticalAlign: "top" | "middle" | "bottom";
}

export interface AngledTextBox {
  /** Top-left corner of the line box before turning (cell coordinates, px):
   * draw it with transform-origin 0 0 and `rotate(rotate deg)`. */
  x: number;
  y: number;
  /** The line box's height (ascent + descent) and where its baseline sits. */
  lineHeight: number;
  ascent: number;
  rotate: number;
}

/** Where Excel draws an angled cell's text (see the file comment). */
export function angledCellTextBox(input: AngledTextInput): AngledTextBox {
  const { mode, angle, width: w, height: h, textWidth, fontSize, align, verticalAlign } = input;
  const theta = (Math.abs(angle) * Math.PI) / 180;
  const sin = Math.sin(theta);
  const cos = Math.cos(theta);
  const rising = angle > 0;
  const ascent = fontSize * ASCENT;
  const lineHeight = fontSize * (ASCENT + DESCENT);
  const descentY = fontSize * DESCENT * cos;
  const padTop = mode === "band" ? BAND_PAD.top : RECT_PAD.top;
  const padBottom = mode === "band" ? BAND_PAD.bottom : RECT_PAD.bottom;
  // Text longer than the row height is placed as if cut to what fits.
  const length = Math.min(textWidth, Math.max(0, (h - padTop - padBottom - descentY) / sin));
  // Box of the turned line box, relative to its top-left corner P (screen y
  // down). Rising: u = (cos, -sin), down = (sin, cos); falling: u = (cos,
  // sin), down = (-sin, cos).
  const boxMinX = rising ? 0 : -lineHeight * sin;
  const boxMaxX = rising ? length * cos + lineHeight * sin : length * cos;
  const boxMinY = rising ? -length * sin : 0;
  const boxMaxY = rising ? lineHeight * cos : length * sin + lineHeight * cos;
  const boxHeight = boxMaxY - boxMinY;
  const top =
    verticalAlign === "top"
      ? padTop
      : verticalAlign === "middle"
        ? (padTop + h - padBottom - boxHeight) / 2
        : h - padBottom - boxHeight;
  const y = top - boxMinY;
  let x: number;
  if (mode === "rect") {
    const boxWidth = boxMaxX - boxMinX;
    const left =
      align === "right"
        ? w - RECT_PAD.right - boxWidth
        : align === "center"
          ? (RECT_PAD.left + w - RECT_PAD.right - boxWidth) / 2
          : RECT_PAD.left;
    x = left - boxMinX;
  } else {
    // The long side facing each slanted edge keeps its distance from it,
    // measured across. Edges at screen height Y: left (h - Y) * cot, right
    // w + (h - Y) * cot (signed: falling bands lean left).
    const cot = (rising ? 1 : -1) * (cos / sin);
    const edge = (yy: number) => (h - yy) * cot;
    const pads = rising ? BAND_PAD.rising : BAND_PAD.falling;
    // The long side facing the left edge passes through P (rising: the top
    // side) or P + down*lineHeight (falling: the bottom side), and the one
    // facing the right edge through the other.
    const down = rising ? { x: lineHeight * sin, y: lineHeight * cos } : { x: -lineHeight * sin, y: lineHeight * cos };
    const leftSide = rising ? { x: 0, y: 0 } : down;
    const rightSide = rising ? down : { x: 0, y: 0 };
    const leftX = edge(y + leftSide.y) + pads.left - leftSide.x;
    const rightX = w + edge(y + rightSide.y) - pads.right - rightSide.x;
    x = align === "right" ? rightX : align === "center" ? (leftX + rightX) / 2 : leftX;
  }
  return { x, y, lineHeight, ascent, rotate: rising ? -Math.abs(angle) : Math.abs(angle) };
}

export interface CssBorderLine {
  width: number;
  color: string;
  dash?: string;
  double?: boolean;
}

/** A CSS border declaration ("1px solid #000") as a stroke, or null for none. */
export function cssBorderLine(value: unknown): CssBorderLine | null {
  if (typeof value !== "string" || !value.trim()) return null;
  const width = Number.parseFloat(/(-?\d*\.?\d+)px/.exec(value)?.[1] ?? "1");
  const style = /\b(none|hidden|solid|dashed|dotted|double|groove|ridge|inset|outset)\b/.exec(value)?.[1] ?? "solid";
  if (style === "none" || style === "hidden" || !(width > 0)) return null;
  const color = value.replace(/(-?\d*\.?\d+)px/, "").replace(style, "").trim() || "currentColor";
  return {
    width,
    color,
    ...(style === "dashed" ? { dash: `${3 * width} ${width}` } : style === "dotted" ? { dash: `${width} ${width}` } : {}),
    ...(style === "double" ? { double: true } : {}),
  };
}
