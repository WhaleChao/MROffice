/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Shared, editor-agnostic FRAME value objects for Word paragraphs and text
 * boxes: the line they are drawn with, and the paint inside a text box.
 * Values are deliberately data-only: callers must pass every imported
 * or persisted value through these normalizers before rendering or exporting.
 *
 * The text box's fill lives beside its outline rather than in a module of its
 * own, because they are the two halves of one thing — a text box IS a shape,
 * drawn with a line and a fill — and the page had an attribute for only the
 * line (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §37.2). The gradient half
 * is not re-modelled here: it is the ONE shared gradient value object from
 * `@/lib/office-shape-paint`, the same one a `documentShape` carries.
 */

import { ptToPx } from "./office-units";
import {
  normalizeOfficeShapeGradient,
  officeGradientCss,
  type OfficeShapeGradient,
} from "./office-shape-paint";

export const DOCUMENT_PARAGRAPH_BORDER_EDGES = [
  "top",
  "left",
  "bottom",
  "right",
  "between",
  "bar",
] as const;

export const DOCUMENT_PARAGRAPH_BORDER_PRESETS = [
  "bottom",
  "top",
  "left",
  "right",
  "none",
  "outside",
  "all",
  "insideH",
] as const;

export const DOCUMENT_BORDER_LINE_STYLES = [
  "single",
  "double",
  "dashed",
  "dotted",
] as const;

export type DocumentParagraphBorderEdge =
  (typeof DOCUMENT_PARAGRAPH_BORDER_EDGES)[number];
export type DocumentParagraphBorderPreset =
  (typeof DOCUMENT_PARAGRAPH_BORDER_PRESETS)[number];
export type DocumentBorderLineStyle =
  (typeof DOCUMENT_BORDER_LINE_STYLES)[number];

export interface DocumentBorderLine {
  readonly style: string;
  readonly sizePt?: number;
  readonly color?: string;
  /** OOXML distance from text, in points. Paragraph borders only. */
  readonly spacePt?: number;
}

export type DocumentParagraphBorders = Partial<
  Record<DocumentParagraphBorderEdge, DocumentBorderLine>
>;

export interface DocumentBorderPen {
  readonly style: DocumentBorderLineStyle;
  readonly sizePt: number;
  readonly color: string;
}

export const DEFAULT_DOCUMENT_PARAGRAPH_BORDER: DocumentBorderPen = {
  style: "single",
  sizePt: 0.5,
  color: "#000000",
};

export const DEFAULT_DOCUMENT_TEXT_BOX_OUTLINE: DocumentBorderPen = {
  style: "single",
  sizePt: 0.75,
  color: "#000000",
};

const BORDER_STYLE_PATTERN = /^[A-Za-z][A-Za-z0-9]{0,31}$/;
const BORDER_COLOR_PATTERN = /^#[0-9a-f]{6}$/;
const PARAGRAPH_EDGE_SET = new Set<string>(DOCUMENT_PARAGRAPH_BORDER_EDGES);

function finiteNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Number.parseFloat(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

export function normalizeDocumentBorderColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const raw = value.trim().toLowerCase();
  const normalized = raw.startsWith("#") ? raw : `#${raw}`;
  return BORDER_COLOR_PATTERN.test(normalized) ? normalized : null;
}

export function normalizeDocumentBorderLine(
  value: unknown,
  options: { readonly allowSpace?: boolean } = {},
): DocumentBorderLine | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const source = value as Record<string, unknown>;
  if (
    typeof source.style !== "string" ||
    !BORDER_STYLE_PATTERN.test(source.style)
  ) {
    return null;
  }
  const line: {
    style: string;
    sizePt?: number;
    color?: string;
    spacePt?: number;
  } = { style: source.style };
  const size = finiteNumber(source.sizePt);
  if (size !== null && size >= 0 && size <= 96) {
    line.sizePt = Math.round(size * 1000) / 1000;
  }
  const color = normalizeDocumentBorderColor(source.color);
  if (color) line.color = color;
  if (options.allowSpace) {
    const space = finiteNumber(source.spacePt);
    if (space !== null && space >= 0 && space <= 31) {
      line.spacePt = Math.round(space * 1000) / 1000;
    }
  }
  return line;
}

export function normalizeDocumentBorderPen(
  value: unknown,
  fallback: DocumentBorderPen = DEFAULT_DOCUMENT_PARAGRAPH_BORDER,
): DocumentBorderPen {
  const line = normalizeDocumentBorderLine(value);
  const style = DOCUMENT_BORDER_LINE_STYLES.includes(
    line?.style as DocumentBorderLineStyle,
  )
    ? (line?.style as DocumentBorderLineStyle)
    : fallback.style;
  return {
    style,
    sizePt:
      typeof line?.sizePt === "number" && line.sizePt > 0
        ? line.sizePt
        : fallback.sizePt,
    color: line?.color ?? fallback.color,
  };
}

export function normalizeDocumentParagraphBorders(
  value: unknown,
): DocumentParagraphBorders | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const source = value as Record<string, unknown>;
  const borders: DocumentParagraphBorders = {};
  for (const edge of DOCUMENT_PARAGRAPH_BORDER_EDGES) {
    if (!PARAGRAPH_EDGE_SET.has(edge)) continue;
    const line = normalizeDocumentBorderLine(source[edge], {
      allowSpace: true,
    });
    if (line) borders[edge] = line;
  }
  return Object.keys(borders).length > 0 ? borders : null;
}

export function normalizeDocumentTextBoxOutline(
  value: unknown,
): DocumentBorderLine | null {
  return normalizeDocumentBorderLine(value);
}

export function documentBorderLineIsVisible(
  line: DocumentBorderLine | null | undefined,
): boolean {
  // A zero width is drawn at Word's narrowest line, not hidden (corpus
  // R83-D20: `w:val="double" w:sz="0"` draws a hairline box in Word).
  return Boolean(line && line.style !== "nil" && line.style !== "none");
}

function stableLine(line: DocumentBorderLine): DocumentBorderLine {
  return {
    style: line.style,
    ...(typeof line.sizePt === "number" ? { sizePt: line.sizePt } : {}),
    ...(line.color ? { color: line.color } : {}),
    ...(typeof line.spacePt === "number" ? { spacePt: line.spacePt } : {}),
  };
}

export function stableDocumentParagraphBordersJson(
  value: unknown,
): string | null {
  const borders = normalizeDocumentParagraphBorders(value);
  if (!borders) return null;
  const ordered: DocumentParagraphBorders = {};
  for (const edge of DOCUMENT_PARAGRAPH_BORDER_EDGES) {
    const line = borders[edge];
    if (line) ordered[edge] = stableLine(line);
  }
  return JSON.stringify(ordered);
}

export function stableDocumentTextBoxOutlineJson(
  value: unknown,
): string | null {
  const line = normalizeDocumentTextBoxOutline(value);
  return line ? JSON.stringify(stableLine(line)) : null;
}

export function parseDocumentParagraphBordersAttribute(
  value: string | null,
): DocumentParagraphBorders | null {
  if (!value) return null;
  try {
    return normalizeDocumentParagraphBorders(JSON.parse(value));
  } catch {
    return null;
  }
}

export function parseDocumentTextBoxOutlineAttribute(
  value: string | null,
): DocumentBorderLine | null {
  if (!value) return null;
  try {
    return normalizeDocumentTextBoxOutline(JSON.parse(value));
  } catch {
    return null;
  }
}

/**
 * ST_Border's compound rules (two or three lines) draw as CSS `double`, as the
 * page border does (sectionLayouts `PAGE_BORDER_LINE_STYLES`); before, a
 * paragraph's `thinThickThinLargeGap` drew one heavy solid line where Word
 * draws three thin and thick ones (corpus R83-D20).
 */
const COMPOUND_BORDER_STYLES = new Set([
  "double",
  "doubleWave",
  "triple",
  "thinThickSmallGap",
  "thickThinSmallGap",
  "thinThickThinSmallGap",
  "thinThickMediumGap",
  "thickThinMediumGap",
  "thinThickThinMediumGap",
  "thinThickLargeGap",
  "thickThinLargeGap",
  "thinThickThinLargeGap",
]);

function cssLineStyle(style: string): string {
  if (style === "nil" || style === "none") return "none";
  if (COMPOUND_BORDER_STYLES.has(style)) return "double";
  if (style === "dotted" || style === "dotDash" || style === "dotDotDash") {
    return "dotted";
  }
  if (
    style === "dashed" ||
    style === "dashSmallGap" ||
    style === "dashDotStroked"
  ) {
    return "dashed";
  }
  if (style === "threeDEmboss") return "groove";
  if (style === "threeDEngrave") return "ridge";
  if (style === "outset" || style === "inset") return style;
  return "solid";
}

/** The CSS width of an edge: a zero width is Word's narrowest line, one pixel, drawn single.
 * `wordStyle` is the ST_Border value, `style` the CSS style it draws as. */
function edgeWidthPx(style: string, sizePt: number, wordStyle?: string): number {
  if (sizePt === 0) return 1;
  const rawPx = Math.min(128, Math.max(1, ptToPx(sizePt)));
  // Word 16.86 (corpus R159-D03, its PDF's strokes): `triple` is three lines
  // of the w:sz width with two gaps of it, five widths in all (sz 24: 14.4 pt);
  // `threeDEmboss` / `threeDEngrave` a half, a whole and a half width, two in
  // all (sz 24: 5.76 pt). Drawn three and one wide, each such paragraph came
  // out 3 to 15 pt short and the page held a paragraph more than Word's.
  if (wordStyle === "triple") return Math.max(5, rawPx * 5);
  if (wordStyle === "threeDEmboss" || wordStyle === "threeDEngrave") return Math.max(2, rawPx * 2);
  return style === "double" ? wordDoubleBorderWidthPx(rawPx) : rawPx;
}

export function documentBorderLineCss(
  value: unknown,
  fallback: DocumentBorderPen = DEFAULT_DOCUMENT_PARAGRAPH_BORDER,
): string | null {
  const line = normalizeDocumentBorderLine(value);
  if (!line) return null;
  const cssStyle = cssLineStyle(line.style);
  if (cssStyle === "none") return "0px solid transparent";
  const rawPt = line.sizePt ?? fallback.sizePt;
  // Keep the authored point value in the model/export path, but make a visible
  // editor edge at least one CSS pixel so the same 0.5pt pen has the same
  // optical authority as an imported table-cell border.
  const widthPx = edgeWidthPx(cssStyle, rawPt, line.style);
  const style = rawPt === 0 ? "solid" : cssStyle;
  return `${Number(widthPx.toFixed(3))}px ${style} ${line.color ?? fallback.color}`;
}

/**
 * The CSS width of a Word double border whose `w:sz` line is `linePx` wide.
 * ST_Border's width is ONE line's: Word 16 draws `double` as two lines of
 * that width with a gap of the same width (sz 6: 0.75pt + 0.75pt + 0.75pt;
 * sz 24: 3pt + 3pt + 3pt), so the edge is three times as wide - CSS
 * `double` splits its width into the same thirds. Never under 3px, the
 * narrowest edge that still shows two lines.
 */
export function wordDoubleBorderWidthPx(linePx: number): number {
  return Math.max(3, linePx * 3);
}

/** The CSS pixel width an edge draws at (as `documentBorderLineCss` writes it). */
export function documentBorderLineWidthPx(value: unknown): number {
  const line = normalizeDocumentBorderLine(value);
  if (!line || !documentBorderLineIsVisible(line)) return 0;
  const style = cssLineStyle(line.style);
  if (style === "none") return 0;
  return Number(edgeWidthPx(style, line.sizePt ?? DEFAULT_DOCUMENT_PARAGRAPH_BORDER.sizePt, line.style).toFixed(3));
}

export function documentParagraphBordersHtmlAttributes(
  value: unknown,
  options: {
    readonly externalLeft?: boolean;
    readonly spaceAfterPt?: number | null;
    /** A right-to-left paragraph keeps its edges where they are drawn today. */
    readonly rtl?: boolean;
    /** How an edge appears on this fragment of a paragraph a page cut
     * (fragmentRoles: Word leaves the box open at the seam, keeping the
     * bottom edge's room before it). */
    readonly edgeOnFragment?: (edge: DocumentParagraphBorderEdge) => "draw" | "reserve" | "omit";
  } = {},
): Record<string, string> {
  const borders = normalizeDocumentParagraphBorders(value);
  const serialized = stableDocumentParagraphBordersJson(borders);
  if (!borders || !serialized) return {};
  const style: string[] = ["box-sizing:border-box"];
  // Native paragraph indentation positions the TEXT edge. A left border is
  // outside that edge and must not consume the paragraph's available width.
  // Keep unmeasured compound borders and other consumers on their existing path.
  const left = borders.left;
  const externalLeft = Boolean(options.externalLeft && left?.style === "single" &&
    documentBorderLineIsVisible(left) &&
    DOCUMENT_PARAGRAPH_BORDER_EDGES.every(edge => edge === "left" || !documentBorderLineIsVisible(borders[edge])));
  if (externalLeft && left) {
    // Independent Word controls: 20 indent/space/width cases and 5 font
    // variations place the inside edge a further 1.5pt outside w:space.
    
    style.push("position:relative", "border-left:none");
    style.push("--doc-native-left-border-clearance:1.5pt");
    style.push(`--doc-native-left-border-width:${left.sizePt ?? DEFAULT_DOCUMENT_PARAGRAPH_BORDER.sizePt}pt`);
    style.push(`--doc-native-left-border-space:${left.spacePt ?? 0}pt`);
    style.push(`--doc-native-left-border-color:${left.color ?? DEFAULT_DOCUMENT_PARAGRAPH_BORDER.color}`);
    const after = options.spaceAfterPt;
    if (typeof after === "number" && Number.isFinite(after) && after >= 0) {
      style.push(`--doc-native-left-border-after:${after}pt`);
    }
  }
  for (const edge of ["top", "right", "bottom", "left"] as const) {
    const line = borders[edge];
    if (!line || (edge === "left" && externalLeft)) continue;
    const shown = options.edgeOnFragment?.(edge) ?? "draw";
    if (shown === "omit") continue;
    if (shown === "reserve") {
      // The edge's width and distance as an undrawn border: the shading
      // stops at the padding box, where the text's own room ends.
      if (documentBorderLineIsVisible(line)) {
        style.push(`border-${edge}:calc(${line.spacePt ?? 0}pt + ${documentBorderLineWidthPx(line)}px) solid transparent`);
        style.push("background-clip:padding-box");
      }
      continue;
    }
    const css = documentBorderLineCss(line);
    if (css) style.push(`border-${edge}:${css}`);
    if (documentBorderLineIsVisible(line) && (line.spacePt ?? 0) > 0) {
      style.push(`padding-${edge}:${line.spacePt}pt`);
    }
  }
  // Word draws the left and right edges OUTSIDE the paragraph's indents, by
  // their distance from text and their own width, into the margin: the text
  // keeps its place and its line width. Inside the indents they pushed the
  // text in and narrowed every line.
  if (!options.rtl) {
    for (const edge of ["left", "right"] as const) {
      const line = borders[edge];
      if (!line || (edge === "left" && externalLeft) || !documentBorderLineIsVisible(line)) continue;
      const out = `${line.spacePt ?? 0}pt - ${documentBorderLineWidthPx(line)}px`;
      style.push(`margin-${edge}:calc(var(--doc-paragraph-indent-${edge}, 0pt) - ${out})`);
      style.push(`--doc-border-${edge}-out:calc(${out.replace(" - ", " + ")})`);
    }
  }
  return {
    "data-paragraph-borders": serialized,
    ...(externalLeft ? { "data-native-paragraph-left-border": "true" } : {}),
    style: style.join(";"),
  };
}

export function documentTextBoxOutlineHtmlAttributes(
  value: unknown,
): Record<string, string> {
  const outline = normalizeDocumentTextBoxOutline(value);
  const serialized = stableDocumentTextBoxOutlineJson(outline);
  if (!outline || !serialized) return {};
  const css = documentBorderLineCss(outline, DEFAULT_DOCUMENT_TEXT_BOX_OUTLINE);
  return {
    "data-text-box-outline": serialized,
    ...(css ? { style: `border:${css};box-sizing:border-box` } : {}),
  };
}

/**
 * The paint inside a text-box frame. `color` is the ONE spelling a shape's
 * fill uses on this page (`none` or `#rrggbb`); `gradient` is the ONE shared
 * gradient value object, whose stops keep their theme identity.
 *
 * `null` from the normalizer is a frame that states NOTHING about its fill,
 * which is not the same as one that states there is none (`{ color: "none" }`)
 * — inventing white for the first would paint over every legacy text box that
 * never had a fill drawn.
 */
export interface DocumentTextBoxFill {
  readonly color: string;
  readonly gradient?: OfficeShapeGradient;
  /** A PICTURE fill (`a:blipFill`): the picture, and how it covers the box. */
  readonly picture?: DocumentTextBoxPictureFill;
}

/**
 * Word's Format Shape > Fill > Picture or texture fill on a text box, as the
 * server reads it (`floating_text_box._picture_fill`): `tile` with the six
 * `a:tile` numbers as the package states them, or `stretch` over the box.
 */
export interface DocumentTextBoxPictureFill {
  readonly src: string;
  readonly tile?: {
    readonly tx?: number;
    readonly ty?: number;
    readonly sx?: number;
    readonly sy?: number;
    readonly flip?: "none" | "x" | "y" | "xy";
    readonly algn?: "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br";
  };
  readonly stretch?: true;
  readonly rotWithShape?: boolean;
}

const TEXT_BOX_FILL_COLOR_PATTERN = /^(?:none|#[0-9a-f]{6})$/;
const TEXT_BOX_PICTURE_SRC_PATTERN = /^data:image\/[a-z0-9.+-]+;base64,[A-Za-z0-9+/=]+$/;
const TEXT_BOX_TILE_FLIPS = new Set(["none", "x", "y", "xy"]);
const TEXT_BOX_TILE_ALIGNMENTS = new Set(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);

function normalizeDocumentTextBoxPictureFill(value: unknown): DocumentTextBoxPictureFill | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const source = value as Record<string, unknown>;
  if (typeof source.src !== "string" || !TEXT_BOX_PICTURE_SRC_PATTERN.test(source.src)) return null;
  const rotWithShape = typeof source.rotWithShape === "boolean" ? { rotWithShape: source.rotWithShape } : {};
  if (source.stretch === true && source.tile === undefined) {
    return { src: source.src, stretch: true, ...rotWithShape };
  }
  if (!source.tile || typeof source.tile !== "object" || Array.isArray(source.tile)) return null;
  const raw = source.tile as Record<string, unknown>;
  const tile: Record<string, number | string> = {};
  for (const key of ["tx", "ty", "sx", "sy"] as const) {
    if (raw[key] === undefined) continue;
    if (typeof raw[key] !== "number" || !Number.isSafeInteger(raw[key])) return null;
    tile[key] = raw[key] as number;
  }
  if (raw.flip !== undefined) {
    if (typeof raw.flip !== "string" || !TEXT_BOX_TILE_FLIPS.has(raw.flip)) return null;
    tile.flip = raw.flip;
  }
  if (raw.algn !== undefined) {
    if (typeof raw.algn !== "string" || !TEXT_BOX_TILE_ALIGNMENTS.has(raw.algn)) return null;
    tile.algn = raw.algn;
  }
  return { src: source.src, tile: tile as DocumentTextBoxPictureFill["tile"], ...rotWithShape };
}

export function normalizeDocumentTextBoxFill(
  value: unknown,
): DocumentTextBoxFill | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const source = value as Record<string, unknown>;
  const raw = typeof source.color === "string" ? source.color.trim().toLowerCase() : "";
  const color = raw === "none" ? "none" : normalizeDocumentBorderColor(raw);
  if (!color || !TEXT_BOX_FILL_COLOR_PATTERN.test(color)) return null;
  // A gradient is only ever paint ON a colour; "no fill" with a gradient is
  // two answers to one question, so the gradient is dropped rather than drawn
  // over a frame the source says is empty.
  if (color === "none") {
    // A picture fill is paint the frame draws instead of a colour: the box
    // states "no colour" beside it, never a guessed one.
    const picture = normalizeDocumentTextBoxPictureFill(source.picture);
    return picture ? { color, picture } : { color };
  }
  const gradient = normalizeOfficeShapeGradient(source.gradient);
  return gradient ? { color, gradient } : { color };
}

export function stableDocumentTextBoxFillJson(value: unknown): string | null {
  const fill = normalizeDocumentTextBoxFill(value);
  if (!fill) return null;
  return JSON.stringify(
    fill.gradient
      ? { color: fill.color, gradient: fill.gradient }
      : fill.picture
        ? { color: fill.color, picture: fill.picture }
        : { color: fill.color },
  );
}

export function parseDocumentTextBoxFillAttribute(
  value: string | null,
): DocumentTextBoxFill | null {
  if (!value) return null;
  try {
    return normalizeDocumentTextBoxFill(JSON.parse(value));
  } catch {
    return null;
  }
}

/** The CSS a frame's fill paints with, or null when it paints nothing. */
export function documentTextBoxFillCss(value: unknown): string | null {
  const fill = normalizeDocumentTextBoxFill(value);
  if (!fill) return null;
  if (fill.gradient) return officeGradientCss(fill.gradient);
  if (fill.picture) {
    // Stretched over the box, or tiled from its offset at the picture's own
    // size. `a:tile`'s scale is a fraction of that natural size, which CSS
    // cannot state without the picture's pixels; the canvas draws it at 100%
    // while the page and the file keep the stated scale exactly.
    const picture = fill.picture;
    if (picture.stretch) return `url("${picture.src}") 0 0 / 100% 100% no-repeat`;
    const tile = picture.tile ?? {};
    return `url("${picture.src}") ${(tile.tx ?? 0) / 9525}px ${(tile.ty ?? 0) / 9525}px / auto repeat`;
  }
  return fill.color === "none" ? "transparent" : fill.color;
}

export function documentTextBoxFillHtmlAttributes(
  value: unknown,
): Record<string, string> {
  const serialized = stableDocumentTextBoxFillJson(value);
  const css = documentTextBoxFillCss(value);
  if (!serialized || !css) return {};
  return {
    "data-text-box-fill": serialized,
    style: `background:${css}`,
  };
}

export function paragraphBordersFromLegacyPreset(
  value: unknown,
  line: DocumentBorderLine = DEFAULT_DOCUMENT_PARAGRAPH_BORDER,
): DocumentParagraphBorders | null {
  const normalizedLine = normalizeDocumentBorderLine(line, {
    allowSpace: true,
  });
  if (!normalizedLine || typeof value !== "string") return null;
  if (value === "box") {
    return {
      top: normalizedLine,
      right: normalizedLine,
      bottom: normalizedLine,
      left: normalizedLine,
    };
  }
  if (value === "bottom") return { bottom: normalizedLine };
  if (value === "left") return { left: normalizedLine };
  return null;
}

function sameVisibleLine(
  left: DocumentBorderLine | undefined,
  right: DocumentBorderLine | undefined,
): boolean {
  if (!left || !right) return left === right;
  return (
    left.style === right.style &&
    (left.sizePt ?? null) === (right.sizePt ?? null) &&
    (left.color ?? null) === (right.color ?? null) &&
    (left.spacePt ?? null) === (right.spacePt ?? null)
  );
}

export function legacyPresetFromParagraphBorders(
  value: unknown,
): "box" | "bottom" | "left" | null {
  const borders = normalizeDocumentParagraphBorders(value);
  if (!borders) return null;
  const visible = DOCUMENT_PARAGRAPH_BORDER_EDGES.filter((edge) =>
    documentBorderLineIsVisible(borders[edge]),
  );
  if (visible.length === 1 && visible[0] === "bottom") return "bottom";
  if (visible.length === 1 && visible[0] === "left") return "left";
  if (
    visible.length === 4 &&
    ["top", "right", "bottom", "left"].every((edge) =>
      visible.includes(edge as DocumentParagraphBorderEdge),
    ) &&
    sameVisibleLine(borders.top, borders.right) &&
    sameVisibleLine(borders.top, borders.bottom) &&
    sameVisibleLine(borders.top, borders.left)
  ) {
    return "box";
  }
  return null;
}


/** Paint native single borders without CSS's one-device-pixel minimum.
 * The existing CSS path supplies other styles; only single-line edges are
 * replaced. Reserve their exact point width as padding so text flow remains
 * independent from the renderer's border snapping.
 */
export function paintPhysicalParagraphBorderEdges(
  paragraph: HTMLElement,
  value: unknown,
): void {
  const borders = normalizeDocumentParagraphBorders(value);
  if (!borders) return;
  const doc = paragraph.ownerDocument;
  for (const edge of ["top", "right", "bottom", "left"] as const) {
    const line = borders[edge];
    if (!line || line.style !== "single" || !documentBorderLineIsVisible(line)) continue;
    const widthPt = line.sizePt ?? DEFAULT_DOCUMENT_PARAGRAPH_BORDER.sizePt;
    const horizontal = edge === "top" || edge === "bottom";
    paragraph.style.position = "relative";
    paragraph.style.setProperty(`border-${edge}`, "none");
    paragraph.style.setProperty(`padding-${edge}`, `${(line.spacePt ?? 0) + widthPt}pt`);
    const svg = doc.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("focusable", "false");
    svg.setAttribute("data-native-border-edge", edge);
    svg.style.position = "absolute";
    svg.style.pointerEvents = "none";
    svg.style.width = horizontal ? "100%" : `${widthPt}pt`;
    svg.style.height = horizontal ? `${widthPt}pt` : "100%";
    svg.style.setProperty(edge, "0");
    svg.style.setProperty(horizontal ? "left" : "top", "0");
    const rect = doc.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("width", "100%");
    rect.setAttribute("height", "100%");
    rect.setAttribute("fill", line.color ?? DEFAULT_DOCUMENT_PARAGRAPH_BORDER.color);
    svg.appendChild(rect);
    paragraph.appendChild(svg);
  }
}
