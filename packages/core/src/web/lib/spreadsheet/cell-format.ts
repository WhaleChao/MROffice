/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet cell formatting (alignment, wrap-text, borders).
 *
 * The grid stores formatting on each cell's `style` object and renders it via
 * inline CSS. These helpers keep the formatting logic out of the React
 * component so it can be unit-tested in isolation, and so the toolbar and the
 * cell renderer agree on how style fields map to CSS.
 *
 * Border model: the legacy `border` boolean draws a full box ("all" borders).
 * For the "outline" mode we set the individual `borderTop/Right/Bottom/Left`
 * booleans on only the perimeter cells of a selection so the outer edge is
 * drawn without filling the interior. "none" clears every border field.
 */

export type HorizontalAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";
export type BorderMode = "all" | "outline" | "none";

/**
 * Rich per-edge border line styles (Excel/WPS parity). The legacy boolean
 * borders below render as a 1px solid `#475569` box; a rich edge instead carries
 * its own {style,color} and renders as `<width> <css-style> <color>`.
 */
export type BorderLineStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "dashDot"
  | "dashDotDot"
  | "mediumDashed"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "slantDashDot";

/** Every SpreadsheetML line style, named as the file names it, in the order
 * Excel's Format Cells dialog lists them. */
export const BORDER_LINE_STYLES: readonly BorderLineStyle[] = [
  "hair",
  "dotted",
  "dashDotDot",
  "dashDot",
  "dashed",
  "thin",
  "mediumDashDotDot",
  "slantDashDot",
  "mediumDashDot",
  "mediumDashed",
  "medium",
  "thick",
  "double",
] as const;

/** The four box edges plus an optional diagonal a rich border can target. */
export type BorderEdge = "top" | "right" | "bottom" | "left" | "diagonal";

export const BORDER_EDGES: readonly BorderEdge[] = [
  "top",
  "right",
  "bottom",
  "left",
  "diagonal",
] as const;

/** A single rich edge: a line style plus a CSS color.
 *
 * The colour is OPTIONAL because Excel's is: an edge written with no colour is
 * its "automatic" colour, which is what a plain black grid line is. Storing a
 * concrete black in its place would turn automatic into explicit the first time
 * anyone touched the cell's borders, so the absence is kept and resolved only
 * at render time.
 */
export interface RichBorderEdge {
  style: BorderLineStyle;
  color?: string;
  /** Diagonal edge only: the line runs bottom-left to top-right. */
  up?: boolean;
  /** Diagonal edge only: the line runs top-left to bottom-right. Excel allows
   * both at once, which draws a cross. Neither set reads as `down`, the split
   * -header form. */
  down?: boolean;
}

/** What Excel draws for an edge that names no colour of its own. */
export const AUTOMATIC_BORDER_COLOR = "#000000";

/**
 * Rich per-edge border map. Only edges the user explicitly set are present, so
 * a cell with no rich borders simply omits `borderEdges` entirely and stays
 * byte-stable with the pre-rich-border serialization.
 */
export type CellBorderEdges = Partial<Record<BorderEdge, RichBorderEdge>>;

/** The border-related fields a cell style may carry. */
export interface CellBorderStyle {
  border?: boolean;
  borderTop?: boolean;
  borderRight?: boolean;
  borderBottom?: boolean;
  borderLeft?: boolean;
  /**
   * Rich per-edge borders (color + line style). Layered OVER the legacy boolean
   * borders so the basic "all"/"outline" toolbar still works unchanged. Persist
   * ONLY when at least one edge is set (default = absent = byte-stable).
   */
  borderEdges?: CellBorderEdges;
}

/** The formatting subset this module reasons about. */
export interface CellFormatStyle extends CellBorderStyle {
  textAlign?: HorizontalAlign | string;
  verticalAlign?: VerticalAlign | string;
  wrapText?: boolean;
  /** Text rotation in degrees (Excel "Orientation"); see ROTATION helpers. */
  textRotation?: number;
  /** Excel's Vertical Text (stacked letters, XLSX textRotation 255): an
   *  orientation of its own, exclusive with an angle. */
  verticalText?: boolean;
  /** Left-indent level (Excel "Indent"); see INDENT helpers. */
  indent?: number;
  /** Shrink text to fit the column (Excel "Shrink to fit"). */
  shrinkToFit?: boolean;
  /** Center this cell's text across the empty cells to its right (Excel). */
  centerAcross?: boolean;
  /** Horizontal Justify / Distributed (Indent), beside textAlign. */
  horizontalJustify?: "justify" | "distributed";
  /** Vertical Justify / Distributed, beside verticalAlign. */
  verticalJustify?: "justify" | "distributed";
  /** "Justify distributed": a Distributed cell's last line spreads too. */
  justifyLastLine?: boolean;
}

// ── Text rotation / orientation (Excel "Format Cells > Alignment > Orientation")
//
// Excel stores an angle in degrees from -90 (text reads downward) through 0
// (horizontal) to +90 (text reads upward). We persist the same signed integer
// and render it as a CSS `rotate()` transform. The persisted value is clamped to
// the legal range and snapped to an integer so a forged/garbled stored value can
// never produce a non-finite transform or rotate the text off its baseline.

/** Inclusive bounds on a stored rotation angle (Excel's orientation dial). */
export const MIN_TEXT_ROTATION = -90;
export const MAX_TEXT_ROTATION = 90;

/** The discrete angles the toolbar's orientation control offers (degrees). */
export const TEXT_ROTATION_PRESETS: readonly number[] = [
  -90, -45, 0, 45, 90,
] as const;

/**
 * Coerce an arbitrary value into a legal rotation angle, or `0` (no rotation).
 * Non-finite / non-numeric inputs and out-of-range values are rejected: the
 * result is always an integer in [-90, 90]. `0` means "horizontal" and is the
 * byte-stable default (callers prune it so an unrotated cell stores nothing).
 */
export function normalizeTextRotation(value: unknown): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return 0;
  const clamped = Math.max(MIN_TEXT_ROTATION, Math.min(MAX_TEXT_ROTATION, n));
  return Math.round(clamped);
}

/**
 * Whether a cell's text is turned (a non-zero Orientation angle) or stacked
 * (Vertical Text). Such text is drawn by the cell's display layer alone, the
 * one element that turns: the cell's fill stays on the cell, level, and the
 * editor stays an ordinary horizontal entry over the cell.
 */
export function cellTextIsOriented(
  style: Pick<CellFormatStyle, "textRotation" | "verticalText"> | undefined,
): boolean {
  return style?.verticalText === true || normalizeTextRotation(style?.textRotation) !== 0;
}

/**
 * CSS fragment that renders a cell's text at the given rotation. Excel's dial is
 * counter-clockwise-positive (a +90 reads bottom-to-top), whereas CSS `rotate()`
 * is clockwise-positive, so the sign is negated. Returns `undefined` for an
 * effectively-horizontal angle so an unrotated cell emits no transform (and thus
 * stays byte-for-byte identical to today's render).
 *
 * `transformOrigin` is centered so the rotated text pivots in place rather than
 * drifting out of the cell. `display: inline-block` is required for a transform
 * to take on inline text; the caller wraps the cell's text in such a span.
 */
export function textRotationCss(value: unknown): {
  transform?: string;
  transformOrigin?: string;
  display?: string;
  whiteSpace?: string;
} {
  const angle = normalizeTextRotation(value);
  if (angle === 0) return {};
  return {
    transform: `rotate(${-angle}deg)`,
    transformOrigin: "center",
    display: "inline-block",
    // Rotated text must not wrap (a wrapped rotated run is illegible); Excel
    // also keeps rotated text on a single line within the cell.
    whiteSpace: "nowrap",
  };
}

/**
 * The display layer's orientation CSS (B340, B341). Vertical Text stacks the
 * letters upright, top to bottom. A quarter turn lays the text out in a
 * vertical writing mode rather than turning a horizontal box, so the cell's
 * vertical alignment places the turned text itself: bottom-aligned "ABCD"
 * turned 90° sits inside the cell's bottom edge instead of past it. Other
 * angles keep the turned horizontal box of {@link textRotationCss}.
 */
/**
 * Where a rotated (non-quarter-turn) text block is anchored vertically
 * (XLS-A040). Excel keeps the whole turned text inside the cell: text turned
 * counterclockwise rises from the cell's bottom, text turned clockwise hangs
 * from its top. A Middle cell keeps the block centred. Null when the cell is
 * not at such an angle.
 */
export function rotatedTextVerticalAnchor(
  style: Pick<CellFormatStyle, "textRotation" | "verticalText" | "verticalAlign" | "verticalJustify"> | undefined,
): "flex-start" | "flex-end" | null {
  // Vertical Text set Justify or Distributed down the cell starts at the top,
  // as Excel stacks it (R20-X14's 「満足している」 headers hung from the top in
  // Excel and stood on the bottom here, the default the alignment fell to).
  if (style?.verticalText === true) return style.verticalJustify === "justify" || style.verticalJustify === "distributed" ? "flex-start" : null;
  const angle = normalizeTextRotation(style?.textRotation);
  if (angle === 0 || Math.abs(angle) === 90 || style?.verticalAlign === "middle") return null;
  return angle > 0 ? "flex-end" : "flex-start";
}

/**
 * The horizontal place of turned text under General alignment (XLS-A040),
 * MEASURED in Excel: text turned 90° counterclockwise sits at the cell's
 * right edge and Vertical Text (stacked letters) is centred, while -90°
 * stays at the left. Null when the cell states its own alignment.
 */
export function rotatedTextGeneralJustify(
  style: Pick<CellFormatStyle, "textRotation" | "verticalText" | "textAlign" | "centerAcross"> | undefined,
): "flex-end" | "center" | null {
  if (!style || (style.textAlign !== undefined && style.textAlign !== "") || style.centerAcross === true) return null;
  if (style.verticalText === true) return "center";
  return normalizeTextRotation(style.textRotation) === 90 ? "flex-end" : null;
}

export function textOrientationCss(
  style: Pick<CellFormatStyle, "textRotation" | "verticalText" | "verticalAlign" | "wrapText"> | undefined,
): {
  transform?: string;
  transformOrigin?: string;
  display?: string;
  whiteSpace?: string;
  maxHeight?: string;
  writingMode?: "vertical-rl";
  textOrientation?: "upright" | "sideways";
  letterSpacing?: string;
} {
  if (style?.verticalText === true) {
    // Excel stacks Vertical Text a character to a line: each takes its line's
    // height down the cell, not its em -- 「満足している」 runs about half as
    // long again as the letters packed at 1em, and a header too long for its
    // cell goes on in a second column where Excel puts one (xlsx-rounds
    // R20-X14's survey headers). The spacing makes each character's advance
    // its line box.
    return { writingMode: "vertical-rl", textOrientation: "upright", letterSpacing: "calc(1lh - 1em)" };
  }
  const angle = normalizeTextRotation(style?.textRotation);
  // vertical-rl reads top to bottom (a clockwise quarter turn, -90); turned
  // half way round it reads bottom to top, Excel's counterclockwise 90. The
  // whole line turns, its CJK characters too (sideways): only Vertical Text
  // stands them upright. A vertical line's own CJK orientation is upright,
  // which left them standing at -90 and stood them on their heads at 90.
  // Wrap Text on a quarter-turned cell wraps its line at the cell's height,
  // into columns across its width, as Excel draws it (S23_tdf119292).
  const turnedWrap = style?.wrapText === true
    ? { whiteSpace: "pre-wrap", maxHeight: "100%" }
    : { whiteSpace: "nowrap" };
  if (angle === -90) return { writingMode: "vertical-rl", textOrientation: "sideways", ...turnedWrap };
  if (angle === 90) {
    return {
      writingMode: "vertical-rl",
      textOrientation: "sideways",
      transform: "rotate(180deg)",
      transformOrigin: "center",
      ...turnedWrap,
    };
  }
  const anchor = rotatedTextVerticalAnchor(style);
  if (anchor && angle !== 0) {
    // Turned about its bottom-left (rising) or top-left (falling) corner and
    // moved right by the line box's swing, so no part of the first glyph
    // leaves the cell (XLS-A040: "rot 45" lost its "A" under the cell's
    // bottom edge when the block turned about its centre).
    const swing = (1.25 * Math.sin((Math.abs(angle) * Math.PI) / 180)).toFixed(3);
    return {
      transform: `translateX(${swing}em) rotate(${-angle}deg)`,
      transformOrigin: anchor === "flex-end" ? "left bottom" : "left top",
      display: "inline-block",
      whiteSpace: "nowrap",
    };
  }
  return textRotationCss(angle);
}

// ── Indent (Excel "Format Cells > Alignment > Indent" / toolbar Increase/Decrease)
//
// An indent level shifts the cell's content away from its leading edge by a
// fixed step per level. Excel allows 0..250; we store the same non-negative
// integer level and render it as left padding. Like Excel, the step is a small
// fixed width per level (Excel uses ~3 character-widths; we use a px step that
// reads correctly at the grid's default font size).

/** Max indent level (matches Excel's 0..250). */
export const MAX_INDENT_LEVEL = 250;

/**
 * Base leading padding of a cell (mirrors `.spreadsheet-window__grid input`
 * `padding: 0 9px`); indent levels are added on top of this so a 0-indent
 * cell keeps the byte-stable default.
 */
export const CELL_BASE_PADDING_X_PX = 9;

/** Pixels of leading padding added per indent level on a page's own sheet;
 * an imported grid's level is its Normal font's three spaces
 * (ResolvedNativeGrid.indentStepPx, officeIndentStepPx). */
export const INDENT_STEP_PX = 12;

/**
 * Coerce an arbitrary value into a legal indent LEVEL (a non-negative integer in
 * [0, 250]). Non-finite / negative / non-numeric inputs collapse to `0` (no
 * indent), the byte-stable default. `0` is pruned by callers so an unindented
 * cell stores nothing.
 */
export function normalizeIndentLevel(value: unknown): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return 0;
  const floored = Math.floor(n);
  if (floored <= 0) return 0;
  return Math.min(floored, MAX_INDENT_LEVEL);
}

/**
 * The next indent level after pressing Increase (`+1`) or Decrease (`-1`),
 * clamped to [0, 250]. A pure transition so the toolbar buttons stay thin and
 * the clamp is tested once here. `delta` is rounded so a stray fractional step
 * can't drift the level.
 */
export function stepIndentLevel(current: unknown, delta: number): number {
  const base = normalizeIndentLevel(current);
  const step = Number.isFinite(delta) ? Math.round(delta) : 0;
  return normalizeIndentLevel(base + step);
}

/**
 * The CSS leading-padding (in px) for an indent level. Indent shifts the content
 * away from the cell's LEADING edge — for the grid's default left-to-right text
 * that is the left edge. Returns `0` for level 0 so an unindented cell is
 * unaffected. The padding is additive on TOP of the grid's base cell padding,
 * so callers add it to (not replace) the default.
 */
export function indentPaddingPx(value: unknown, stepPx: number = INDENT_STEP_PX): number {
  return normalizeIndentLevel(value) * stepPx;
}

// ── Justify / Distributed (Format Cells > Alignment, XLS-A037/A038/A058) ─────
//
// Excel's Horizontal list goes on past Left/Center/Right to Fill, Justify and
// Distributed (Indent), and its Vertical list to Justify and Distributed
// (ECMA-376 ST_HorizontalAlignment / ST_VerticalAlignment). The model keeps
// them beside textAlign / verticalAlign as `horizontalJustify` and
// `verticalJustify`, so every reader of the three-value keys stays as it is;
// `justifyLastLine` is the "Justify distributed" box.
//
//  - Justify wraps the text and stretches every line but a paragraph's last
//    to both edges ("Long    text    that").
//  - Distributed stretches every line, the last too, spreading East Asian
//    text character by character ("Ab　　中　　文"); its indent applies at
//    both edges.
//  - Vertically, Justify spreads the lines from the top edge to the bottom
//    one (a single line sits at the top) and Distributed spreads them evenly
//    (a single line sits in the middle).

export type CellSpreadAlignment = "justify" | "distributed";

/** Format Cells > Alignment > Horizontal as one choice ("" is General). */
export type CellHorizontalChoice =
  | ""
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerAcross"
  | "distributed";

/** Every style key the Horizontal list owns, cleared before the chosen one is set. */
const HORIZONTAL_CHOICE_CLEAR = {
  textAlign: undefined,
  centerAcross: undefined,
  horizontalFill: undefined,
  horizontalJustify: undefined,
  justifyLastLine: undefined,
} as const;

/**
 * The style patch one pick of Format Cells' Horizontal list writes: the list
 * is one choice, the model splits it over textAlign, centerAcross,
 * horizontalFill and horizontalJustify, so a pick states all of them (and
 * clears "Justify distributed", which belongs to Distributed alone).
 */
export function cellHorizontalChoicePatch(next: CellHorizontalChoice): {
  textAlign?: "left" | "center" | "right";
  centerAcross?: boolean;
  horizontalFill?: boolean;
  horizontalJustify?: CellSpreadAlignment;
  justifyLastLine?: boolean;
  indent?: number;
} {
  // Only Left, Right and Distributed (Indent) keep an indent: any other pick
  
  // with indent 2 set to Center reads indent 0, and set back to Left stays 0;
  // Left with indent 2 set to Right keeps 2). Kept under Center, the level
  // was invisible and came back the moment the cell was aligned left again.
  const noIndent = { indent: undefined } as const;
  switch (next) {
    case "left":
    case "right":
      return { ...HORIZONTAL_CHOICE_CLEAR, textAlign: next };
    case "center":
      return { ...HORIZONTAL_CHOICE_CLEAR, ...noIndent, textAlign: next };
    case "fill":
      return { ...HORIZONTAL_CHOICE_CLEAR, ...noIndent, horizontalFill: true };
    case "justify":
      return { ...HORIZONTAL_CHOICE_CLEAR, ...noIndent, horizontalJustify: next };
    case "distributed":
      return { ...HORIZONTAL_CHOICE_CLEAR, horizontalJustify: next };
    case "centerAcross":
      return { ...HORIZONTAL_CHOICE_CLEAR, ...noIndent, centerAcross: true };
    default:
      return { ...HORIZONTAL_CHOICE_CLEAR, ...noIndent };
  }
}

/** Whether the cell's horizontal alignment holds an indent: Left (Indent),
 * Right (Indent) and Distributed (Indent), as Excel's Format Cells lists
 * them. General, Center, Fill, Justify and Center Across Selection do not. */
export function cellHorizontalTakesIndent(style: {
  textAlign?: unknown;
  centerAcross?: unknown;
  horizontalFill?: unknown;
  horizontalJustify?: unknown;
} | undefined): boolean {
  const spread = cellHorizontalSpread(style);
  if (spread) return spread === "distributed";
  if (style?.centerAcross === true || style?.horizontalFill === true) return false;
  return style?.textAlign === "left" || style?.textAlign === "right";
}









export function cellIndentPatch(
  style: (NonNullable<Parameters<typeof cellHorizontalTakesIndent>[0]> & { indent?: unknown }) | undefined,
  change: { step: number } | { level: number },
): ReturnType<typeof cellHorizontalChoicePatch> {
  const level = "level" in change ? normalizeIndentLevel(change.level) : stepIndentLevel(style?.indent, change.step);
  const indent = level === 0 ? undefined : level;
  if (level > 0 && !cellHorizontalTakesIndent(style)) {
    return { ...HORIZONTAL_CHOICE_CLEAR, textAlign: "left", indent };
  }
  return { indent };
}

function spreadValue(value: unknown): CellSpreadAlignment | null {
  return value === "justify" || value === "distributed" ? value : null;
}

/** The cell's horizontal Justify / Distributed alignment, or null. */
export function cellHorizontalSpread(style: { horizontalJustify?: unknown } | undefined): CellSpreadAlignment | null {
  return spreadValue(style?.horizontalJustify);
}

/** The cell's vertical Justify / Distributed alignment, or null. */
export function cellVerticalSpread(style: { verticalJustify?: unknown } | undefined): CellSpreadAlignment | null {
  return spreadValue(style?.verticalJustify);
}

/** Han, kana, Hangul and their full-width forms and punctuation. */
const CJK_TEXT = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]|[\u{20000}-\u{3ffff}]/u;










export function distributedLineUnits(line: string): { mode: "spread" | "center"; units: string[] } {
  const text = line.trim();
  const units = CJK_TEXT.test(text) ? Array.from(text) : text.split(" ").filter((word) => word !== "");
  return { mode: units.length < 2 ? "center" : "spread", units };
}

/** CSS that lays a wrapped text block out as Justify / Distributed (with the
 * cell's text, a Distributed block of one Latin word centres). */
export function cellHorizontalSpreadCss(style: {
  horizontalJustify?: unknown;
  justifyLastLine?: unknown;
} | undefined, text?: string): {
  textAlign: "justify" | "center";
  textAlignLast: "justify" | "auto" | "center";
  textJustify: "inter-character" | "inter-word" | "auto";
} | null {
  const spread = cellHorizontalSpread(style);
  if (!spread) return null;
  if (spread === "distributed") {
    if (text === undefined || CJK_TEXT.test(text)) {
      return { textAlign: "justify", textAlignLast: "justify", textJustify: "inter-character" };
    }
    return text.trim().includes(" ")
      ? { textAlign: "justify", textAlignLast: "justify", textJustify: "inter-word" }
      : { textAlign: "center", textAlignLast: "center", textJustify: "auto" };
  }
  return {
    textAlign: "justify",
    textAlignLast: style?.justifyLastLine === true ? "justify" : "auto",
    textJustify: "auto",
  };
}

/** The flex `justify-content` that spreads a cell's lines vertically. */
let wrappedBottomJustify: string | undefined;

/**
 * The flex justification of bottom-aligned wrapped lines. Excel bottom-aligns
 * them only while they fit: lines taller than the row show their FIRST lines,
 * cut at the bottom (E20_Googlesourced: "Special" over a 12.75pt row, where a
 * plain flex-end kept "offer price", X-061). `safe flex-end` is exactly that;
 * an engine without the keyword keeps the plain end.
 */
export function cellWrappedBottomJustifyContent(): string {
  if (wrappedBottomJustify === undefined) {
    wrappedBottomJustify = typeof CSS !== "undefined" && typeof CSS.supports === "function" && CSS.supports("justify-content", "safe flex-end")
      ? "safe flex-end"
      : "flex-end";
  }
  return wrappedBottomJustify;
}

let wrappedMiddleJustify: string | undefined;

/**
 * The flex justification of vertically centred wrapped lines. As at the
 * bottom (X-061), Excel centres them only while they fit: lines taller than
 * the row show their FIRST line, cut below (CR067_GroupTagTests_WithHeader's
 * "{{item.Cu" in a one-line row, where plain centring showed the middle line
 * "omer.Com", X-261). `safe center` is exactly that.
 */
export function cellWrappedMiddleJustifyContent(): string {
  if (wrappedMiddleJustify === undefined) {
    wrappedMiddleJustify = typeof CSS !== "undefined" && typeof CSS.supports === "function" && CSS.supports("justify-content", "safe center")
      ? "safe center"
      : "center";
  }
  return wrappedMiddleJustify;
}

export function cellVerticalSpreadJustifyContent(style: { verticalJustify?: unknown } | undefined): "space-between" | "space-evenly" | null {
  const spread = cellVerticalSpread(style);
  return spread === "justify" ? "space-between" : spread === "distributed" ? "space-evenly" : null;
}

/**
 * The indent's pixels on each side of the text (XLS-A039). Excel measures the
 * indent from the edge the text is aligned to: the left for Left (and
 * General), the RIGHT for Right, both for Distributed. Center takes none.
 */
export function cellIndentInsetsPx(style: {
  indent?: unknown;
  textAlign?: unknown;
  horizontalJustify?: unknown;
  centerAcross?: unknown;
} | undefined, stepPx: number = INDENT_STEP_PX): { left: number; right: number } {
  const pad = indentPaddingPx(style?.indent, stepPx);
  if (pad === 0) return { left: 0, right: 0 };
  const edge = cellIndentEdge(style);
  return {
    left: edge === "left" || edge === "both" ? pad : 0,
    right: edge === "right" || edge === "both" ? pad : 0,
  };
}

/**
 * The edge a cell's indent is measured from (XLS-A039): the right for Right,
 * both for Distributed, none for Center, else the left. The Home ribbon's
 * indent buttons show and name this edge (video audit U02).
 */
export function cellIndentEdge(style: {
  textAlign?: unknown;
  horizontalJustify?: unknown;
  centerAcross?: unknown;
} | undefined): "left" | "right" | "both" | "none" {
  if (cellHorizontalSpread(style) === "distributed") return "both";
  if (style?.textAlign === "right") return "right";
  if (style?.textAlign === "center" || style?.centerAcross === true) return "none";
  return "left";
}

// ── Shrink to fit (Excel "Format Cells > Alignment > Shrink to fit") ──────────
//
// When set, Excel scales the cell's font down so the (unwrapped) text fits the
// column width. We render the closest faithful CSS: keep the text on one line
// and let it scale down within the cell box. Shrink-to-fit and wrap-text are
// mutually exclusive in Excel — wrap wins — so the resolver below encodes that.

/** Whether a style has shrink-to-fit ACTIVE (it is suppressed when wrap is on). */
export function isShrinkToFitActive(style: {
  shrinkToFit?: unknown;
  wrapText?: unknown;
}): boolean {
  return Boolean(style.shrinkToFit) && !style.wrapText;
}

/**
 * CSS fragment for shrink-to-fit. Excel shrinks the FONT to fit the column; the
 * faithful CSS keeps the run on a single line and clips it to the cell, and the
 * renderer pairs this with a fitted font-size (computed from the measured
 * width). Returns `{}` when shrink is not active (wrap on, or flag unset) so the
 * cell renders byte-for-byte as today. The companion `shrinkToFitFontScale`
 * computes the multiplier the renderer applies to the cell's font-size.
 */
export function shrinkToFitCss(style: {
  shrinkToFit?: unknown;
  wrapText?: unknown;
}): { whiteSpace?: string; overflow?: string; textOverflow?: string } {
  if (!isShrinkToFitActive(style)) return {};
  return { whiteSpace: "nowrap", overflow: "hidden", textOverflow: "clip" };
}

/**
 * The font-size SCALE (0,1] that makes `contentWidthPx` of text fit inside
 * `cellWidthPx`. Excel never scales UP (a short string keeps its size), so the
 * result is capped at 1. A bounded floor (`minScale`, default 0.3) keeps the
 * text from collapsing to an unreadable sliver on a very narrow column. Degenerate
 * inputs (non-positive / non-finite widths) yield `1` (no scaling) so a bad
 * measurement never blanks the cell.
 */
export function shrinkToFitFontScale(
  contentWidthPx: number,
  cellWidthPx: number,
  minScale = 0.3,
): number {
  if (
    !Number.isFinite(contentWidthPx) ||
    !Number.isFinite(cellWidthPx) ||
    contentWidthPx <= 0 ||
    cellWidthPx <= 0
  ) {
    return 1;
  }
  if (contentWidthPx <= cellWidthPx) return 1;
  const scale = cellWidthPx / contentWidthPx;
  const floor = Number.isFinite(minScale) && minScale > 0 ? Math.min(minScale, 1) : 0.3;
  return Math.max(floor, Math.min(1, scale));
}

// ── Center Across Selection (Excel "Format Cells > Alignment > Horizontal") ────
//
// "Center Across Selection" centers the leading cell's text over a run of
// cells WITHOUT merging them (so each cell stays independently editable). The
// run extends rightward from the anchor across as many cells as the user
// selected, stopping at the first cell that itself holds a value — mirroring
// Excel, which only centers across the BLANK cells to the right of the text.

/** A non-empty cell predicate for center-across span computation. */
export type CellHasValue = (col: number) => boolean;

/**
 * Compute the inclusive column span `[startCol, endCol]` that a Center-Across
 * anchor at `startCol` should visually cover, bounded by the user's selection
 * (`selectionEndCol`) and the sheet edge (`maxCol`).
 *
 * The span starts at the anchor and extends right while each next column is
 * BLANK (per `hasValue`). It stops BEFORE the first non-blank column, at the end
 * of the selection, or at `maxCol`, whichever comes first. A single-column
 * selection (or an immediately-following non-blank cell) yields just the anchor
 * column — center-across then behaves like an ordinary centered cell.
 *
 * Pure + total: never throws, always returns `endCol >= startCol`, and clamps a
 * reversed / out-of-range selection so the span can't escape the anchor.
 */
export function centerAcrossSpan(
  startCol: number,
  selectionEndCol: number,
  maxCol: number,
  hasValue: CellHasValue,
): { startCol: number; endCol: number } {
  const start = Math.max(0, Math.floor(startCol));
  const hardEnd = Math.min(
    Math.max(start, Math.floor(selectionEndCol)),
    Math.max(start, Math.floor(maxCol)),
  );
  let end = start;
  for (let col = start + 1; col <= hardEnd; col += 1) {
    if (hasValue(col)) break; // stop before a column that holds its own value.
    end = col;
  }
  return { startCol: start, endCol: end };
}

/**
 * The run of columns a Center Across Selection cell (`centerContinuous`) at
 * `col` is centred over, as Excel draws it on screen and in print: the cell
 * and the cells right of it that `joins` (blank, and Center Across Selection
 * too), up to `maxCol`. Null when no neighbour joins: the text is then
 * centred on its own cell like any centred text (X-350).
 */
export function centerAcrossRun(
  col: number,
  maxCol: number,
  joins: (col: number) => boolean,
): { startCol: number; endCol: number } | null {
  const span = centerAcrossSpan(col, maxCol, maxCol, (probe) => !joins(probe));
  return span.endCol > span.startCol ? span : null;
}

/** Drawn over the faint grid lines so a bordered edge reads as a solid box. */
export const CELL_BORDER_CSS = "1px solid #475569";

/** Default color a fresh rich edge picks up (matches the legacy box tone). */
export const DEFAULT_BORDER_COLOR = "#475569";

/**
 * Pixel width + CSS line-style for each rich border style. `thick` and `double`
 * are 3px so the doubled lines are visible; dashed/dotted stay 1px like Excel's
 * hairline variants. DEFER: exact Excel pixel mapping is approximate — these are
 * readable on-screen widths, not OOXML point-for-point.
 */
// CSS has no dash-dot line, so those styles draw dashed at their weight; a
// hairline is the finest dotted line.
const BORDER_STYLE_CSS: Record<BorderLineStyle, { width: number; line: string }> =
  {
    thin: { width: 1, line: "solid" },
    medium: { width: 2, line: "solid" },
    thick: { width: 3, line: "solid" },
    dashed: { width: 1, line: "dashed" },
    dotted: { width: 1, line: "dotted" },
    double: { width: 3, line: "double" },
    hair: { width: 1, line: "dotted" },
    dashDot: { width: 1, line: "dashed" },
    dashDotDot: { width: 1, line: "dashed" },
    mediumDashed: { width: 2, line: "dashed" },
    mediumDashDot: { width: 2, line: "dashed" },
    mediumDashDotDot: { width: 2, line: "dashed" },
    slantDashDot: { width: 2, line: "dashed" },
  };

function isBorderLineStyle(value: unknown): value is BorderLineStyle {
  return (
    typeof value === "string" &&
    (BORDER_LINE_STYLES as readonly string[]).includes(value)
  );
}

function isBorderEdge(value: unknown): value is BorderEdge {
  return (
    typeof value === "string" && (BORDER_EDGES as readonly string[]).includes(value)
  );
}

/** A `#rgb`/`#rrggbb` hex color (the only shape the color <input> emits). */
function isHexColor(value: unknown): value is string {
  return typeof value === "string" && /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(value);
}

/** CSS shorthand (`<width> <line> <color>`) for one rich edge. */
export function richBorderEdgeToCss(edge: RichBorderEdge): string {
  const spec = BORDER_STYLE_CSS[edge.style];
  return `${spec.width}px ${spec.line} ${edge.color ?? AUTOMATIC_BORDER_COLOR}`;
}

/** Dash runs of the dash-dot styles, in CSS pixels: dash, gap, dot, gap(, dot, gap). */
const DASH_DOT_RUNS: Partial<Record<BorderLineStyle, readonly number[]>> = {
  dashDot: [9, 3, 3, 3],
  dashDotDot: [9, 3, 3, 3, 3, 3],
  mediumDashDot: [9, 3, 3, 3],
  mediumDashDotDot: [9, 3, 3, 3, 3, 3],
  // Measured off Excel for Mac's PDF (R07_bootstrap32 B9): 11.25 pt dash,
  // 0.75 pt gap, 5.25 pt dash, 0.75 pt gap, at 2 pt; drawn as two 1 px rows,
  // the upper one a pixel ahead, so each gap leans like a slash.
  slantDashDot: [15, 1, 7, 1],
  // Hair and dotted are drawn the same way, not as a CSS `dotted` border:
  // Chrome paints a 1px dotted edge of a collapsed table cell solid on the
  // vertical sides (CR086's Customer box: dotted top and bottom, solid left
  // and right, where Excel dots every side; X-290). The rhythms are the
  // Format Cells samples' (borderLineStrokeSpec).
  hair: [1, 1],
  dotted: [1, 2],
};

/**
 * How to stroke one line style in a drawn sample (Format Cells > Border's
 * Style list and preview, XLS-A060): the grid's own width, the dash runs the
 * grid paints (dash-dot rhythms, CSS dashed/dotted otherwise) and whether
 * the style is two strokes.
 */
export function borderLineStrokeSpec(style: BorderLineStyle): {
  readonly width: number;
  readonly dash: readonly number[] | null;
  readonly double: boolean;
} {
  const spec = BORDER_STYLE_CSS[style];
  const runs = DASH_DOT_RUNS[style];
  const dash =
    runs ??
    (spec.line === "dashed"
      ? [4 * spec.width, 2 * spec.width]
      : spec.line === "dotted"
        ? style === "hair"
          ? [1, 1]
          : [1, 2]
        : null);
  return {
    width: spec.line === "double" ? 1 : spec.width,
    dash,
    double: spec.line === "double",
  };
}

/** One background layer of a grid cell. */
export interface CellBackgroundLayer {
  readonly image: string;
  readonly size: string;
  readonly position: string;
  /** Where the layer is laid out; the padding box when absent. */
  readonly origin?: "border-box" | "padding-box";
}

/**
 * The grid's drawing of the dash-dot edges CSS borders cannot draw (they
 * have only dashed and dotted): each such edge becomes a transparent border
 * of its width, so the default gridline does not win the shared edge, plus a
 * background strip along it repeating the style's own dash-dot rhythm (live
 * audit B428). Only the on-screen grid uses this; copied HTML and exported
 * pictures keep `resolveBorderCss`'s dashed approximation.
 */
export function dashDotEdgeLayers(style: CellBorderStyle | undefined): {
  readonly borders: Partial<Record<"borderTop" | "borderRight" | "borderBottom" | "borderLeft", string>>;
  readonly layers: readonly CellBackgroundLayer[];
} {
  const edges = normalizeBorderEdges(style?.borderEdges);
  const borders: Partial<Record<"borderTop" | "borderRight" | "borderBottom" | "borderLeft", string>> = {};
  const layers: CellBackgroundLayer[] = [];
  if (!edges) return { borders, layers };
  for (const [side, key, direction, size, position] of [
    ["top", "borderTop", "to right", (w: number) => `100% ${w}px`, "0 0"],
    ["bottom", "borderBottom", "to right", (w: number) => `100% ${w}px`, "0 100%"],
    ["left", "borderLeft", "to bottom", (w: number) => `${w}px 100%`, "0 0"],
    ["right", "borderRight", "to bottom", (w: number) => `${w}px 100%`, "100% 0"],
  ] as const) {
    const edge = edges[side];
    const runs = edge ? DASH_DOT_RUNS[edge.style] : undefined;
    if (!edge || !runs) continue;
    const width = BORDER_STYLE_CSS[edge.style].width;
    const color = edge.color ?? AUTOMATIC_BORDER_COLOR;
    // The style's rhythm, `shift` px later along the edge (its last gap's
    // tail comes first).
    const strip = (shift: number) => {
      const segments = runs.map((run, index) => [index % 2 === 0 ? color : "transparent", run] as [string, number]);
      if (shift > 0) {
        segments[segments.length - 1][1] -= shift;
        segments.unshift([segments[segments.length - 1][0], shift]);
      }
      let at = 0;
      const stops: string[] = [];
      for (const [paint, run] of segments) {
        if (run <= 0) continue;
        stops.push(`${paint} ${at}px ${at + run}px`);
        at += run;
      }
      return `repeating-linear-gradient(${direction}, ${stops.join(", ")})`;
    };
    borders[key] = `${width}px solid transparent`;
    if (edge.style === "slantDashDot") {
      // Two one-pixel rows (columns), the upper (left) one a pixel ahead:
      // the gaps slant.
      const across = side === "top" || side === "bottom";
      const outerFirst = side === "top" || side === "left";
      const row = (offset: number) =>
        across
          ? side === "top" ? `0 ${offset}px` : `0 calc(100% - ${offset}px)`
          : side === "left" ? `${offset}px 0` : `calc(100% - ${offset}px) 0`;
      for (const [offset, shift] of [[0, outerFirst ? 1 : 0], [1, outerFirst ? 0 : 1]] as const) {
        layers.push({ image: strip(shift), size: size(1), position: row(offset) });
      }
      continue;
    }
    layers.push({ image: strip(0), size: size(width), position });
  }
  return { borders, layers };
}

const dashDotOnce = new WeakMap<object, Map<string, CellBorderStyle>>();

/**
 * `style` with the dash-dot left / top edges dropped that its neighbour draws
 * already: the edge strips (dashDotEdgeLayers) sit inside each cell, so a
 * shared line both cells name (A's right, B's left, as Excel files carry
 * it) came out as two dotted lines a pixel apart where Excel draws one. The
 * neighbour's matching right / bottom edge, same style and colour, draws it
 * alone. The same object for the same answer, so the styles' CSS caches
 * keyed by the style keep working.
 */
/** Whether a cell's left or top edge is a dash-dot line, the only edges {@link dashDotEdgesDrawnOnce} can drop. */
export function hasDashDotLeadingEdge(style: CellBorderStyle | undefined): boolean {
  const edges = style?.borderEdges;
  return Boolean(edges && ((edges.left && DASH_DOT_RUNS[edges.left.style]) || (edges.top && DASH_DOT_RUNS[edges.top.style])));
}

export function dashDotEdgesDrawnOnce<S extends CellBorderStyle>(
  style: S | undefined,
  leftNeighbourRight: RichBorderEdge | undefined,
  topNeighbourBottom: RichBorderEdge | undefined,
): S | undefined {
  const edges = style?.borderEdges;
  if (!style || !edges) return style;
  const same = (own: RichBorderEdge | undefined, theirs: RichBorderEdge | undefined) =>
    Boolean(own && theirs && DASH_DOT_RUNS[own.style] && own.style === theirs.style &&
      (own.color ?? AUTOMATIC_BORDER_COLOR).toLowerCase() === (theirs.color ?? AUTOMATIC_BORDER_COLOR).toLowerCase());
  const dropLeft = same(edges.left, leftNeighbourRight);
  const dropTop = same(edges.top, topNeighbourBottom);
  if (!dropLeft && !dropTop) return style;
  const key = `${dropLeft ? 1 : 0}${dropTop ? 1 : 0}`;
  let known = dashDotOnce.get(style);
  if (!known) dashDotOnce.set(style, (known = new Map()));
  let drawn = known.get(key) as S | undefined;
  if (!drawn) {
    const kept: CellBorderEdges = { ...edges };
    if (dropLeft) delete kept.left;
    if (dropTop) delete kept.top;
    drawn = { ...style, borderEdges: kept };
    known.set(key, drawn);
  }
  return drawn;
}

/**
 * Normalize an arbitrary (possibly persisted/untrusted) rich-edge map into one
 * that keeps ONLY well-formed edges with a known style + hex color. Returns
 * `undefined` when nothing survives, so callers prune the key entirely and the
 * cell stays byte-stable. Used both when applying a toolbar edit and when
 * rendering a loaded sheet (defends the inline-style sink against junk colors).
 */
export function normalizeBorderEdges(
  raw: unknown,
): CellBorderEdges | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const out: CellBorderEdges = {};
  let any = false;
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (!isBorderEdge(key)) continue;
    if (!value || typeof value !== "object") continue;
    const candidate = value as { style?: unknown; color?: unknown };
    if (!isBorderLineStyle(candidate.style)) continue;
    // A colourless edge is Excel's automatic colour, not a broken edge.
    // Dropping it here cost an imported workbook the WEIGHT of every rule it
    // had not coloured by hand: a medium black band fell back to the legacy
    // boolean and drew as a thin grey hairline.
    const normalized: RichBorderEdge = isHexColor(candidate.color)
      ? { style: candidate.style, color: candidate.color }
      : { style: candidate.style };
    if (key === "diagonal") {
      const slope = value as { up?: unknown; down?: unknown };
      if (slope.up === true) normalized.up = true;
      if (slope.down === true) normalized.down = true;
    }
    out[key] = normalized;
    any = true;
  }
  return any ? out : undefined;
}

/**
 * A border patch with every side explicitly turned off. Spread this before a
 * narrower patch so applying a new border mode fully clears any previous one
 * (the toolbar passes the cleared keys through `updateCellStyle`, which prunes
 * `false`/`undefined` so they don't bloat the stored style).
 */
export const CLEARED_BORDER_PATCH: Required<CellBorderStyle> = {
  border: false,
  borderTop: false,
  borderRight: false,
  borderBottom: false,
  borderLeft: false,
  // `undefined` is pruned by updateCellStyle, clearing any prior rich edges so a
  // "none" / "all" / "outline" toolbar action fully resets the cell's borders.
  borderEdges: undefined as unknown as CellBorderEdges,
};

function isHorizontalAlign(value: unknown): value is HorizontalAlign {
  return value === "left" || value === "center" || value === "right";
}

function isVerticalAlign(value: unknown): value is VerticalAlign {
  return value === "top" || value === "middle" || value === "bottom";
}

/**
 * CSS `vertical-align` value for a stored alignment, or undefined when unset.
 * NOTE: undefined does NOT mean "middle" — SpreadsheetML's default is bottom,
 * and the renderer resolves it via its own effectiveVerticalAlign. This helper
 * only normalizes what is stored.
 */
export function verticalAlignToCss(
  value: VerticalAlign | string | undefined,
): "top" | "middle" | "bottom" | undefined {
  return isVerticalAlign(value) ? value : undefined;
}

/**
 * The vertical alignment a cell actually renders with. SpreadsheetML's default
 * when `<alignment vertical="…">` is absent is BOTTOM — Excel and WPS both
 * paint an unaligned cell's text at the bottom of the row. This is the ONE
 * resolver every render path shares (the grid's text layers, Page Layout and
 * the pictures of a range); the paths used to disagree (the plain input
 * centred, the wrap overlay went top, the rich-text overlay centred), so two
 * same-height cells in one row could sit on different baselines purely by
 * which path drew them.
 */
export function effectiveVerticalAlign(
  style: { verticalAlign?: VerticalAlign | string } | undefined,
): VerticalAlign {
  const v = style?.verticalAlign;
  return v === "top" || v === "middle" ? v : "bottom";
}

/**
 * The line box a wrapped (or hard-broken) cell lays each of its lines in, as a
 * multiple of its font size: the grid's wrapped-text layer, AutoFit Row
 * Height and the pictures of a range all stack lines at this pitch.
 */
export const SPREADSHEET_WRAP_LINE_HEIGHT = 1.15;

/** CSS `text-align` value for a stored alignment, or undefined when unset. */
export function horizontalAlignToCss(
  value: HorizontalAlign | string | undefined,
): HorizontalAlign | undefined {
  return isHorizontalAlign(value) ? value : undefined;
}

/**
 * Resolve the border CSS for a cell's style.
 *
 * Layering (lowest → highest precedence per edge):
 *   1. legacy `border` full box, then legacy per-side booleans (1px solid box);
 *   2. rich `borderEdges` {style,color} — a set rich edge overrides the legacy
 *      box on that side, so a thick red top sits on an otherwise plain box.
 *
 * The rich diagonal edge is NOT a CSS box edge and is reported separately via
 * `borderImage`-free means; here we surface it as `diagonalCss` so the renderer
 * can decide how to draw it (DEFER: an actual diagonal line uses a background
 * gradient, handled at the call site). Returns an empty object when nothing is
 * configured (byte-stable default for plain cells).
 */
export function resolveBorderCss(
  style: CellBorderStyle | undefined,
): {
  border?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
} {
  if (!style) return {};
  const edges = normalizeBorderEdges(style.borderEdges);

  // Legacy box first.
  let top: string | undefined;
  let right: string | undefined;
  let bottom: string | undefined;
  let left: string | undefined;
  let fullBox: string | undefined;
  if (style.border) {
    fullBox = CELL_BORDER_CSS;
  } else {
    if (style.borderTop) top = CELL_BORDER_CSS;
    if (style.borderRight) right = CELL_BORDER_CSS;
    if (style.borderBottom) bottom = CELL_BORDER_CSS;
    if (style.borderLeft) left = CELL_BORDER_CSS;
  }

  // Rich edges override per side. A rich edge on any side means we can no longer
  // use the `border` shorthand (sides differ), so expand the full box first.
  if (edges) {
    if (fullBox) {
      top = top ?? fullBox;
      right = right ?? fullBox;
      bottom = bottom ?? fullBox;
      left = left ?? fullBox;
      fullBox = undefined;
    }
    if (edges.top) top = richBorderEdgeToCss(edges.top);
    if (edges.right) right = richBorderEdgeToCss(edges.right);
    if (edges.bottom) bottom = richBorderEdgeToCss(edges.bottom);
    if (edges.left) left = richBorderEdgeToCss(edges.left);
  }

  if (fullBox) return { border: fullBox };
  const css: {
    borderTop?: string;
    borderRight?: string;
    borderBottom?: string;
    borderLeft?: string;
  } = {};
  if (top) css.borderTop = top;
  if (right) css.borderRight = right;
  if (bottom) css.borderBottom = bottom;
  if (left) css.borderLeft = left;
  return css;
}

/** Rectangle of a selection in grid coordinates (inclusive bounds). */
export interface SelectionRect {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/**
 * Compute the border-style patch for a single cell at (row, col) given the
 * selected rectangle and the chosen border mode. The patch always clears every
 * border field first, then turns on the relevant ones, so re-applying a mode is
 * idempotent and switching modes never leaves a stale edge behind.
 *
 * - "all": full box on every selected cell.
 * - "outline": draw only the sides of cells that sit on the selection's edge.
 *   A single-cell selection therefore gets a full box (all four edges).
 * - "none": clear borders.
 */
export function borderPatchForCell(
  mode: BorderMode,
  row: number,
  col: number,
  rect: SelectionRect,
): Required<CellBorderStyle> {
  if (mode === "none") {
    return { ...CLEARED_BORDER_PATCH };
  }
  if (mode === "all") {
    return { ...CLEARED_BORDER_PATCH, border: true };
  }
  // outline: only the perimeter cells, and only on their outward-facing edges.
  return {
    ...CLEARED_BORDER_PATCH,
    borderTop: row === rect.r1,
    borderBottom: row === rect.r2,
    borderLeft: col === rect.c1,
    borderRight: col === rect.c2,
  };
}

/**
 * Build the `{ borderEdges }` style patch for setting (or clearing) ONE rich
 * edge on a cell, given that cell's existing edges. `edge` is the target side;
 * `next` is the new {style,color} to set, or `null` to clear just that edge.
 *
 * Because `updateCellStyle` shallow-merges the patch and prunes falsy values,
 * the returned `borderEdges` must be the FULL post-edit map (it replaces the old
 * one). When the edit leaves no edges, we return `borderEdges: undefined` so the
 * key is pruned and the cell falls back to the byte-stable no-rich-border shape.
 */
export function setBorderEdgePatch(
  existing: CellBorderEdges | undefined,
  edge: BorderEdge,
  next: RichBorderEdge | null,
): { borderEdges: CellBorderEdges | undefined } {
  // Normalize the existing map so junk persisted edges don't survive the edit.
  const base = normalizeBorderEdges(existing) ?? {};
  const merged: CellBorderEdges = { ...base };
  if (next === null) {
    delete merged[edge];
  } else {
    merged[edge] = {
      style: next.style,
      ...(next.color ? { color: next.color } : {}),
      ...(next.up ? { up: true } : {}),
      ...(next.down ? { down: true } : {}),
    };
  }
  const cleaned = normalizeBorderEdges(merged);
  return { borderEdges: cleaned };
}

/**
 * The diagonal rule inside a cell, as a background layer.
 *
 * A diagonal is not a CSS box edge, so it was surfaced but never drawn: a split
 * header — the 斜线表头 every Chinese timetable and form uses — showed as an
 * empty cell. It is painted as a hard-stop gradient band along the cell's own
 * diagonal, built only from a number and a validated colour so no imported
 * value reaches the DOM as CSS.
 */
export function resolveDiagonalCss(
  style: CellBorderStyle | undefined,
): { backgroundImage: string } | undefined {
  const edge = normalizeBorderEdges(style?.borderEdges)?.diagonal;
  if (!edge) return undefined;
  const spec = BORDER_STYLE_CSS[edge.style];
  const color = edge.color ?? AUTOMATIC_BORDER_COLOR;
  // Half the line width either side of the diagonal's midline.
  const half = Math.max(spec.width, 1) / 2;
  const band = (corner: "bottom right" | "top right") =>
    `linear-gradient(to ${corner}, transparent calc(50% - ${half}px), ` +
    `${color} calc(50% - ${half}px), ${color} calc(50% + ${half}px), ` +
    `transparent calc(50% + ${half}px))`;
  const layers: string[] = [];
  // A corner-to-corner gradient's 50% line runs through the OTHER two
  // corners: "to bottom right" draws Excel's diagonalUp (bottom-left to
  // top-right), "to top right" its diagonalDown (top-left to bottom-right).
  // Excel allows both at once (a cross); neither set is the split-header
  // form, top-left to bottom-right (B342).
  if (edge.up) layers.push(band("bottom right"));
  if (edge.down || !edge.up) layers.push(band("top right"));
  return { backgroundImage: layers.join(", ") };
}

/**
 * Excel's Borders gallery (Home ▸ Font ▸ 边框).
 *
 * The thirteen presets Excel offers, as the edges each preset wants on ONE
 * cell of the selection. `null` for an edge means "clear it"; a missing edge
 * means "leave whatever is there". Perimeter-only presets ask the cell where it
 * sits in the selection, which is why the rectangle is passed in — an outline
 * on a 3x3 block paints nine different edge sets, not one.
 *
 * Colours are deliberately absent: an edge with no colour is Excel's AUTOMATIC
 * colour, which is what a plain black grid line is (see AUTOMATIC_BORDER_COLOR).
 */
export type SpreadsheetBorderPreset =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  /** Format Cells > Border > Presets > Inside: every interior edge of the
   * selection and none of its perimeter; nothing for a single cell, which
   * is why Excel dims the button then. */
  | "inside"
  | "thickBox"
  | "bottomDouble"
  | "thickBottom"
  | "topAndBottom"
  | "topAndThickBottom"
  | "topAndDoubleBottom";

export function borderPresetEdgesForCell(
  preset: SpreadsheetBorderPreset,
  row: number,
  col: number,
  rect: SelectionRect,
): Partial<Record<BorderEdge, BorderLineStyle | null>> | null {
  const top = Math.min(rect.r1, rect.r2);
  const bottom = Math.max(rect.r1, rect.r2);
  const left = Math.min(rect.c1, rect.c2);
  const right = Math.max(rect.c1, rect.c2);
  const onTop = row === top;
  const onBottom = row === bottom;
  const onLeft = col === left;
  const onRight = col === right;

  switch (preset) {
    case "none":
      // Handled by the caller via the legacy clear, which also drops the
      // boolean box; nothing per-cell to say here.
      return { top: null, right: null, bottom: null, left: null };
    case "bottom":
      return onBottom ? { bottom: "thin" } : null;
    case "top":
      return onTop ? { top: "thin" } : null;
    case "left":
      return onLeft ? { left: "thin" } : null;
    case "right":
      return onRight ? { right: "thin" } : null;
    case "all":
      return { top: "thin", right: "thin", bottom: "thin", left: "thin" };
    case "inside": {
      // Both cells of a shared interior edge paint it, as "all" does, so a
      // later per-cell edit on either side still finds its own edge.
      const edges: Partial<Record<BorderEdge, BorderLineStyle | null>> = {};
      if (!onTop) edges.top = "thin";
      if (!onBottom) edges.bottom = "thin";
      if (!onLeft) edges.left = "thin";
      if (!onRight) edges.right = "thin";
      return Object.keys(edges).length > 0 ? edges : null;
    }
    case "outside":
    case "thickBox": {
      const line: BorderLineStyle = preset === "thickBox" ? "thick" : "thin";
      const edges: Partial<Record<BorderEdge, BorderLineStyle | null>> = {};
      if (onTop) edges.top = line;
      if (onBottom) edges.bottom = line;
      if (onLeft) edges.left = line;
      if (onRight) edges.right = line;
      return Object.keys(edges).length > 0 ? edges : null;
    }
    case "bottomDouble":
      return onBottom ? { bottom: "double" } : null;
    case "thickBottom":
      return onBottom ? { bottom: "thick" } : null;
    case "topAndBottom":
      return onTop || onBottom
        ? {
            ...(onTop ? { top: "thin" as const } : {}),
            ...(onBottom ? { bottom: "thin" as const } : {}),
          }
        : null;
    case "topAndThickBottom":
      return onTop || onBottom
        ? {
            ...(onTop ? { top: "thin" as const } : {}),
            ...(onBottom ? { bottom: "thick" as const } : {}),
          }
        : null;
    case "topAndDoubleBottom":
      return onTop || onBottom
        ? {
            ...(onTop ? { top: "thin" as const } : {}),
            ...(onBottom ? { bottom: "double" as const } : {}),
          }
        : null;
    default:
      return null;
  }
}

/**
 * Every per-cell border patch one gallery preset makes over a selection.
 *
 * The walk lives here rather than in the window component so it can be tested
 * as what it is — pure geometry — and so "outline over a 3x3 block" is one
 * readable function instead of a nested loop inside a render file. Each entry
 * carries the cell it applies to and the edges that cell wants; a cell the
 * preset does not touch is simply absent.
 */
export function borderPresetPatchesForRange(
  preset: SpreadsheetBorderPreset,
  rect: SelectionRect,
): Array<{
  row: number;
  col: number;
  edges: Partial<Record<BorderEdge, BorderLineStyle | null>>;
}> {
  const top = Math.min(rect.r1, rect.r2);
  const bottom = Math.max(rect.r1, rect.r2);
  const left = Math.min(rect.c1, rect.c2);
  const right = Math.max(rect.c1, rect.c2);
  const patches: Array<{
    row: number;
    col: number;
    edges: Partial<Record<BorderEdge, BorderLineStyle | null>>;
  }> = [];
  for (let row = top; row <= bottom; row += 1) {
    for (let col = left; col <= right; col += 1) {
      const edges = borderPresetEdgesForCell(preset, row, col, rect);
      if (edges) patches.push({ row, col, edges });
    }
  }
  return patches;
}
