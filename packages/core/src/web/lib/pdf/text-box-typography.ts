/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeTextPaint, officeTextPaintCss, type OfficeTextPaint, type OfficeTextShadow } from "../office-text-paint";

/**
 * The Word text-box formats an MRAI text box carries beyond font, size,
 * colour, bold/italic/underline/strikethrough and left/center/right/justify:
 * superscript/subscript, text highlight, exact line spacing, character
 * spacing, paragraph indents and distributed alignment. All of them apply to
 * the whole box (the text-box writer has no per-run styling).
 *
 * Lengths are stored on the edit as fractions of the page height, like
 * `fontSize`, so they keep their size on paper when the font changes (Word's
 * indents and "Exactly" line spacing are absolute too). On the wire they are
 * sent relative to the font size, which the writer turns into points with the
 * same page mapping it uses for the font.
 */

export type PdfTextBoxScript = "superscript" | "subscript";
/** Word's Align Text for a text box (Top is the default, never stored). */
export type PdfTextBoxVerticalAlign = "middle" | "bottom";

export interface PdfTextBoxTabStop {
  readonly position: number;
  /** Word's Bar: no stop of the words, a vertical rule at the position
   * through the box's lines (text-box-tabs.ts `pdfTextBoxBarPositions`). */
  readonly align: "left" | "center" | "right" | "decimal" | "bar";
  readonly leader: "none" | "dot" | "hyphen" | "underscore";
}

/** At most this many stops, each at most this many ems from the left edge
 * (mirror of the writer's _TEXT_BOX_MAX_TAB_STOPS / _TEXT_BOX_MAX_TAB_EM). */
export const PDF_TEXT_BOX_MAX_TAB_STOPS = 50;
const MAX_TAB_EM = 400;
const DEFAULT_TAB_EM_LIMITS = [0.05, 60] as const;
const TAB_ALIGN_CODES = { left: "l", center: "c", right: "r", decimal: "d", bar: "b" } as const;
const TAB_LEADER_CODES = { none: "n", dot: "d", hyphen: "h", underscore: "u" } as const;

/** A box's stops as the writer takes them: ascending, within the bound, of
 * known alignments and leaders (null for none). */
export function pdfTextBoxTabStops(value: unknown, fontSize: number): PdfTextBoxTabStop[] | null {
  if (!Array.isArray(value) || value.length === 0 || value.length > PDF_TEXT_BOX_MAX_TAB_STOPS || !(fontSize > 0)) return null;
  const stops: PdfTextBoxTabStop[] = [];
  let previous = 0;
  for (const stop of value) {
    const record = stop as Partial<PdfTextBoxTabStop> | null;
    if (!record || typeof record.position !== "number" || !Number.isFinite(record.position) ||
      !(record.align && record.align in TAB_ALIGN_CODES) || !(record.leader && record.leader in TAB_LEADER_CODES)) return null;
    const em = record.position / fontSize;
    if (!(em > previous + 1e-9) || em > MAX_TAB_EM) return null;
    previous = em;
    stops.push(Object.freeze({ position: record.position, align: record.align, leader: record.leader }));
  }
  return stops;
}

export interface PdfTextBoxTypographyFields {
  textScript?: PdfTextBoxScript;
  /** Text highlight colour (#RRGGBB) behind each drawn line's glyphs. */
  highlight?: string;
  /** Exact line pitch (baseline to baseline), fraction of page height. */
  lineSpacing?: number;
  /** Extra advance after every character, fraction of page height. */
  characterSpacing?: number;
  /** Extra advance after every space, fraction of page height (PDF Tw). */
  wordSpacing?: number;
  /** Horizontal glyph scale in percent (PDF Tz); 100 or absent is none. */
  horizontalScale?: number;
  /** Word's paragraph Spacing Before / After: above each paragraph's first
   * line and below its last, fractions of page height. */
  spaceBefore?: number;
  spaceAfter?: number;
  leftIndent?: number;
  rightIndent?: number;
  /** First line of each paragraph, relative to the left indent (may be < 0). */
  firstLineIndent?: number;
  /** Word's Distributed: every line, the last too, spans the frame width. */
  distributed?: boolean;
  /** Word's Align Text (Text Box > Vertical alignment): the rows sit at the
   * middle or the bottom of a frame taller than they are. */
  verticalAlign?: PdfTextBoxVerticalAlign;
  /** Word's 竖排 (Text Direction > Vertical): the box set in columns
   * (text-box-vertical.ts). Only a box made in MRAI carries it here: an
   * imported line written vertically keeps its own layout. */
  writingMode?: "vertical-rl";
  /** Word's and PowerPoint's text box Columns (2-16): the rows fill each
   * column to the frame's foot before the next (`a:bodyPr@numCol`). */
  columnCount?: number;
  /** Spacing between columns, fraction of page height (`@spcCol`). */
  columnSpacing?: number;
  /** Word's Tabs: the paragraphs' tab stops, ascending, each from the text
   * area's left edge (fraction of page height), with its alignment and
   * leader (lib/document-tabstops.ts, the Docs editor's). */
  tabStops?: readonly PdfTextBoxTabStop[];
  /** Word's Default tab stops: the interval past the last stop, fraction of
   * page height (absent: 0.5 in). */
  defaultTabStop?: number;
  /** Word's and PowerPoint's WordArt (艺术字): the words' fill (flat, see-
   * through, gradient or pattern), outline and effects (shadow, glow,
   * inner shadow, bevel), the Office paint the other editors
   * draw (lib/office-text-paint.ts). Its lengths are 96-dpi pixels (the
   * outline points), as they are there; the writers set them in points. */
  textPaint?: OfficeTextPaint;
}

/** A box's WordArt as the PDF carries it: Office's paint without what only
 * an Office file states (theme colours, exact effect geometry), a PDF cannot
 * draw flat (a 3-D camera turn) or an annotation cannot show (a reflection,
 * below its box). Null where it is no paint. */
export function pdfTextBoxTextPaint(value: unknown): OfficeTextPaint | null {
  let paint: OfficeTextPaint | undefined;
  try {
    paint = normalizeOfficeTextPaint(value);
  } catch {
    return null;
  }
  if (!paint) return null;
  // A reflection stands below the words, outside the box: an annotation's
  // appearance is clipped to its rectangle, so the PDF never shows one.
  const { stated: _stated, rotation3d: _rotation3d, reflection: _reflection, ...portable } = paint;
  return portable;
}

/** WordArt's 96-dpi pixel lengths in the draft at `pixelsPerPoint` (the
 * page's zoom), the outline's points as well: the draft keeps the paint's
 * size on paper at every zoom. */
export function pdfTextBoxTextPaintAtScale(paint: OfficeTextPaint, pixelsPerPoint: number): OfficeTextPaint {
  const px = (value: number) => value * 0.75 * pixelsPerPoint;
  const shadow = (value: OfficeTextShadow) => ({
    ...value, blur: px(value.blur), offsetX: px(value.offsetX), offsetY: px(value.offsetY),
  });
  return {
    ...paint,
    // CSS points are 4/3 px: the outline drawn as many screen pixels as
    // its points take on the page.
    outlineWidthPt: paint.outlineWidthPt * pixelsPerPoint * 0.75,
    ...(paint.shadow ? { shadow: shadow(paint.shadow) } : {}),
    ...(paint.innerShadow ? { innerShadow: shadow(paint.innerShadow) } : {}),
    ...(paint.glow ? { glow: { ...paint.glow, radius: px(paint.glow.radius) } } : {}),
    ...(paint.bevel
      ? { bevel: { ...paint.bevel, width: px(paint.bevel.width), height: px(paint.bevel.height), depth: px(paint.bevel.depth) } }
      : {}),
  };
}

/** A text box's Columns: 2 to 16, as PowerPoint's Columns dialog takes. */
export const PDF_TEXT_BOX_MAX_COLUMNS = 16;

/** The width of each of `count` columns of a frame `width` wide with
 * `spacing` between them (all in one unit). */
export function pdfTextBoxColumnWidth(width: number, count: number | undefined, spacing: number | undefined): number {
  const columns = count && count > 1 ? count : 1;
  return Math.max(1e-6, (width - (columns - 1) * (spacing ?? 0)) / columns);
}

/** How many rows a column holds: as many whole rows as the frame's height
 * takes (one at least). */
export function pdfTextBoxColumnRows(frameHeight: number, pitch: number): number {
  return Math.max(1, Math.floor((frameHeight + 0.01) / Math.max(1e-6, pitch)));
}

/** The formats measured along a row: a vertical box has no place for them
 * (the editor offers none of them on one, the writers refuse them). */
export const PDF_TEXT_BOX_VERTICAL_UNSET = [
  "textScript", "characterSpacing", "wordSpacing", "horizontalScale",
  "leftIndent", "rightIndent", "firstLineIndent", "distributed",
  "columnCount", "columnSpacing", "spaceBefore", "spaceAfter", "tabStops", "defaultTabStop",
] as const satisfies readonly (keyof PdfTextBoxTypographyFields)[];

/**
 * How far the rows move down in a frame taller than they are (Word's Align
 * Text): half the spare height for Middle, all of it for Bottom. The rows
 * that fill or overflow the frame stay at its top, as Word's do once the
 * frame is full. Every writer of the box (browser draft, raster, flattened
 * export, the server's appearance) moves its rows by this same amount.
 */
export function pdfTextBoxVerticalOffset(
  verticalAlign: PdfTextBoxVerticalAlign | undefined,
  frameHeight: number,
  contentHeight: number,
): number {
  const spare = frameHeight - contentHeight;
  if (!verticalAlign || !(spare > 0)) return 0;
  return verticalAlign === "middle" ? spare / 2 : spare;
}

/** Word draws sub/superscript at about two thirds of the size. */
export const PDF_TEXT_SCRIPT_SIZE_RATIO = 0.65;
/** Baseline shift of a superscript / subscript, in the full font size. */
export const PDF_TEXT_SUPERSCRIPT_RISE = 0.35;
export const PDF_TEXT_SUBSCRIPT_DROP = 0.14;
/** Word's default highlight (Yellow). */
export const PDF_TEXT_DEFAULT_HIGHLIGHT = "#FFFF00";

const HEX = /^#[0-9A-Fa-f]{6}$/u;

const LIMITS = {
  lineSpacing: [0.5, 10],
  characterSpacing: [-0.5, 5],
  wordSpacing: [-0.5, 5],
  leftIndent: [0, 60],
  rightIndent: [0, 60],
  firstLineIndent: [-60, 60],
  columnSpacing: [0, 60],
  spaceBefore: [0, 40],
  spaceAfter: [0, 40],
} as const;

type LengthKey = keyof typeof LIMITS;
const LENGTH_KEYS = Object.keys(LIMITS) as LengthKey[];
/** Horizontal scale, percent. */
const HORIZONTAL_SCALE_LIMITS = [10, 500] as const;

/** The edit's own typography, validated against its font size; only the
 * formats that differ from a plain box are returned. */
export function readPdfTextBoxTypography(
  edit: Omit<PdfTextBoxTypographyFields, "writingMode"> & {
    fontSize?: number;
    writingMode?: string;
    created?: boolean;
  },
  fallbackFontSize: number,
): PdfTextBoxTypographyFields {
  const fontSize =
    typeof edit.fontSize === "number" && edit.fontSize > 0
      ? edit.fontSize
      : fallbackFontSize;
  const out: PdfTextBoxTypographyFields = {};
  if (edit.textScript === "superscript" || edit.textScript === "subscript") {
    out.textScript = edit.textScript;
  }
  if (typeof edit.highlight === "string" && HEX.test(edit.highlight)) {
    out.highlight = edit.highlight.toUpperCase();
  }
  for (const key of LENGTH_KEYS) {
    const value = edit[key];
    if (typeof value !== "number" || !Number.isFinite(value) || value === 0) continue;
    const em = value / fontSize;
    const [low, high] = LIMITS[key];
    if (em >= low && em <= high) out[key] = value;
  }
  if (edit.distributed === true) out.distributed = true;
  const tabStops = edit.tabStops === undefined ? null : pdfTextBoxTabStops(edit.tabStops, fontSize);
  if (tabStops) out.tabStops = tabStops;
  if (typeof edit.defaultTabStop === "number" && Number.isFinite(edit.defaultTabStop) &&
    edit.defaultTabStop / fontSize >= DEFAULT_TAB_EM_LIMITS[0] && edit.defaultTabStop / fontSize <= DEFAULT_TAB_EM_LIMITS[1]) {
    out.defaultTabStop = edit.defaultTabStop;
  }
  if (edit.verticalAlign === "middle" || edit.verticalAlign === "bottom") {
    out.verticalAlign = edit.verticalAlign;
  }
  if (
    typeof edit.horizontalScale === "number" &&
    Number.isFinite(edit.horizontalScale) &&
    edit.horizontalScale !== 100 &&
    edit.horizontalScale >= HORIZONTAL_SCALE_LIMITS[0] &&
    edit.horizontalScale <= HORIZONTAL_SCALE_LIMITS[1]
  ) {
    out.horizontalScale = edit.horizontalScale;
  }
  // Columns are a box's own; their spacing means nothing without them.
  if (
    typeof edit.columnCount === "number" && Number.isInteger(edit.columnCount) &&
    edit.columnCount > 1 && edit.columnCount <= PDF_TEXT_BOX_MAX_COLUMNS
  ) {
    out.columnCount = edit.columnCount;
    // Rows that fill columns have no place for space between paragraphs
    // (the writers take neither with the other).
    delete out.spaceBefore;
    delete out.spaceAfter;
  } else {
    delete out.columnSpacing;
  }
  const textPaint = edit.textPaint === undefined ? null : pdfTextBoxTextPaint(edit.textPaint);
  if (textPaint) out.textPaint = textPaint;
  if (edit.created === true && edit.writingMode === "vertical-rl") {
    out.writingMode = "vertical-rl";
    for (const key of PDF_TEXT_BOX_VERTICAL_UNSET) delete out[key];
  }
  return out;
}

export function hasPdfTextBoxTypography(
  typography: PdfTextBoxTypographyFields,
): boolean {
  return Object.keys(typography).length > 0;
}

function round(value: number): number {
  return Math.round(value * 1e6) / 1e6;
}

function base64Url(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return globalThis
    .btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/u, "");
}

/**
 * The directive field (`~` + base64url JSON) for a box's typography, or null
 * for a plain box (whose directive then stays byte-for-byte the legacy one).
 * `~` is outside the base64url alphabet, so the field can never be mistaken
 * for the text token that precedes it.
 */
export function encodePdfTextBoxTypographyField(
  typography: PdfTextBoxTypographyFields,
  fontSize: number,
): string | null {
  if (!hasPdfTextBoxTypography(typography) || !(fontSize > 0)) return null;
  // (`tp`, the WordArt, is an object: the field is JSON throughout.)
  const json: Record<string, unknown> = { v: 1 };
  if (typography.textScript) json.s = typography.textScript === "superscript" ? "sup" : "sub";
  if (typography.highlight) json.h = typography.highlight.slice(1).toUpperCase();
  if (typography.lineSpacing) json.ls = round(typography.lineSpacing / fontSize);
  if (typography.characterSpacing) json.cs = round(typography.characterSpacing / fontSize);
  if (typography.wordSpacing) json.ws = round(typography.wordSpacing / fontSize);
  if (typography.horizontalScale) json.hz = round(typography.horizontalScale);
  if (typography.leftIndent) json.li = round(typography.leftIndent / fontSize);
  if (typography.rightIndent) json.ri = round(typography.rightIndent / fontSize);
  if (typography.firstLineIndent) json.fi = round(typography.firstLineIndent / fontSize);
  if (typography.distributed) json.d = 1;
  if (typography.verticalAlign) json.va = typography.verticalAlign === "middle" ? "m" : "b";
  if (typography.writingMode === "vertical-rl") json.wm = "v";
  if (typography.columnCount) json.cc = typography.columnCount;
  if (typography.columnSpacing) json.cg = round(typography.columnSpacing / fontSize);
  if (typography.spaceBefore) json.sb = round(typography.spaceBefore / fontSize);
  if (typography.spaceAfter) json.sa = round(typography.spaceAfter / fontSize);
  if (typography.tabStops?.length) {
    json.tb = typography.tabStops.map((stop) =>
      [round(stop.position / fontSize), TAB_ALIGN_CODES[stop.align], TAB_LEADER_CODES[stop.leader]]);
  }
  if (typography.defaultTabStop) json.dt = round(typography.defaultTabStop / fontSize);
  if (typography.textPaint) json.tp = typography.textPaint;
  return `~${base64Url(JSON.stringify(json))}`;
}

export function decodePdfTextBoxTypographyField(
  field: string,
  fontSize: number,
): PdfTextBoxTypographyFields | null {
  if (!/^~[A-Za-z0-9_-]{1,2048}$/u.test(field) || !(fontSize > 0)) return null;
  try {
    const padded = field.slice(1).replaceAll("-", "+").replaceAll("_", "/");
    const binary = globalThis.atob(padded + "=".repeat((4 - (padded.length % 4)) % 4));
    const json = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(
        Uint8Array.from(binary, (character) => character.charCodeAt(0)),
      ),
    ) as Record<string, unknown>;
    if (!json || json.v !== 1) return null;
    const out: PdfTextBoxTypographyFields = {};
    if (json.s === "sup") out.textScript = "superscript";
    else if (json.s === "sub") out.textScript = "subscript";
    else if (json.s !== undefined) return null;
    if (json.h !== undefined) {
      if (typeof json.h !== "string" || !/^[0-9A-F]{6}$/u.test(json.h)) return null;
      out.highlight = `#${json.h}`;
    }
    const lengths: Array<[string, LengthKey]> = [
      ["ls", "lineSpacing"], ["cs", "characterSpacing"], ["ws", "wordSpacing"], ["li", "leftIndent"],
      ["ri", "rightIndent"], ["fi", "firstLineIndent"], ["cg", "columnSpacing"],
      ["sb", "spaceBefore"], ["sa", "spaceAfter"],
    ];
    for (const [short, key] of lengths) {
      const value = json[short];
      if (value === undefined) continue;
      const [low, high] = LIMITS[key];
      if (typeof value !== "number" || !Number.isFinite(value) || value < low || value > high) {
        return null;
      }
      out[key] = value * fontSize;
    }
    if (json.d !== undefined) {
      if (json.d !== 1) return null;
      out.distributed = true;
    }
    if (json.va !== undefined) {
      if (json.va === "m") out.verticalAlign = "middle";
      else if (json.va === "b") out.verticalAlign = "bottom";
      else return null;
    }
    if (json.cc !== undefined) {
      if (typeof json.cc !== "number" || !Number.isInteger(json.cc) || json.cc < 2 ||
        json.cc > PDF_TEXT_BOX_MAX_COLUMNS) return null;
      out.columnCount = json.cc;
    }
    if (out.columnSpacing !== undefined && out.columnCount === undefined) return null;
    if (json.tb !== undefined) {
      const aligns = Object.fromEntries(Object.entries(TAB_ALIGN_CODES).map(([name, code]) => [code, name]));
      const leaders = Object.fromEntries(Object.entries(TAB_LEADER_CODES).map(([name, code]) => [code, name]));
      if (!Array.isArray(json.tb)) return null;
      const stops = pdfTextBoxTabStops(json.tb.map((stop) => Array.isArray(stop) && stop.length === 3
        ? { position: Number(stop[0]) * fontSize, align: aligns[String(stop[1])], leader: leaders[String(stop[2])] } : null), fontSize);
      if (!stops) return null;
      out.tabStops = stops;
    }
    if (json.dt !== undefined) {
      if (typeof json.dt !== "number" || !Number.isFinite(json.dt) || json.dt < DEFAULT_TAB_EM_LIMITS[0] ||
        json.dt > DEFAULT_TAB_EM_LIMITS[1]) return null;
      out.defaultTabStop = json.dt * fontSize;
    }
    if (json.tp !== undefined) {
      const paint = pdfTextBoxTextPaint(json.tp);
      if (!paint) return null;
      out.textPaint = paint;
    }
    if (json.wm !== undefined) {
      if (json.wm !== "v" || PDF_TEXT_BOX_VERTICAL_UNSET.some((key) => key in out) ||
        json.hz !== undefined) return null;
      out.writingMode = "vertical-rl";
    }
    if (json.hz !== undefined) {
      const [low, high] = HORIZONTAL_SCALE_LIMITS;
      if (typeof json.hz !== "number" || !Number.isFinite(json.hz) || json.hz < low || json.hz > high || json.hz === 100) {
        return null;
      }
      out.horizontalScale = json.hz;
    }
    return hasPdfTextBoxTypography(out) ? out : null;
  } catch {
    return null;
  }
}

/** The same typography with every length multiplied by `factor` (page
 * fractions -> em when 1/fontSize, em -> points when the point size). */
export function scalePdfTextBoxTypography(
  typography: PdfTextBoxTypographyFields,
  factor: number,
): PdfTextBoxTypographyFields {
  const out: PdfTextBoxTypographyFields = { ...typography };
  for (const key of LENGTH_KEYS) {
    const value = typography[key];
    if (typeof value === "number") out[key] = value * factor;
  }
  if (typography.tabStops) out.tabStops = typography.tabStops.map((stop) => ({ ...stop, position: stop.position * factor }));
  if (typeof typography.defaultTabStop === "number") out.defaultTabStop = typography.defaultTabStop * factor;
  return out;
}

/** A box's paragraphs: the ranges of its text between line breaks (each
 * without its break). */
export function pdfTextBoxParagraphRanges(text: string): { start: number; end: number }[] {
  const ranges: { start: number; end: number }[] = [];
  let start = 0;
  for (let index = 0; index <= text.length; index += 1) {
    if (index === text.length || text[index] === "\n") {
      ranges.push({ start, end: index });
      start = index + 1;
    }
  }
  return ranges;
}

/** Word's Spacing Before / After in pixels, or null when the box has none:
 * the space above each paragraph's first line and below its last. */
export function pdfTextBoxParagraphSpacingPx(
  typography: PdfTextBoxTypographyFields | undefined,
  pageHeightPx: number,
): { before: number; after: number } | null {
  const before = typography?.spaceBefore ?? 0;
  const after = typography?.spaceAfter ?? 0;
  if (!(before > 0) && !(after > 0)) return null;
  return { before: Math.max(0, before) * pageHeightPx, after: Math.max(0, after) * pageHeightPx };
}

/**
 * Browser layout of the same formats for an editable or read-only box: the
 * draft has to wrap and grow exactly as the writer will (padding for the side
 * indents, text-indent for each paragraph's first row, letter-spacing, an exact line
 * pitch, a smaller raised/lowered face for scripts).
 */
export function pdfTextBoxTypographyStyle(
  typography: PdfTextBoxTypographyFields,
  pageHeightPx: number,
  fontSizePx: number,
  /** Screen pixels per point of the page (its zoom): WordArt's sizes. */
  pixelsPerPoint?: number,
): Record<string, string | number> {
  const px = (value: number | undefined) =>
    value ? `${value * pageHeightPx}px` : undefined;
  const style: Record<string, string | number> = {};
  const pitch = typography.lineSpacing
    ? typography.lineSpacing * pageHeightPx
    : null;
  if (typography.textScript) {
    style.fontSize = `${fontSizePx * PDF_TEXT_SCRIPT_SIZE_RATIO}px`;
    // The smaller face in the same line box already sits ~0.1 em higher;
    // the rest of Word's rise (or drop) is a translation.
    style.translate = `0 ${(typography.textScript === "superscript" ? -0.245 : 0.245) * fontSizePx}px`;
  }
  if (pitch !== null || typography.textScript) {
    style.lineHeight = `${pitch ?? fontSizePx * 1.2}px`;
  }
  const letterSpacing = px(typography.characterSpacing);
  if (letterSpacing) style.letterSpacing = letterSpacing;
  const wordSpacing = px(typography.wordSpacing);
  if (wordSpacing) style.wordSpacing = wordSpacing;
  const left = px(typography.leftIndent);
  if (left) style.paddingLeft = left;
  const right = px(typography.rightIndent);
  if (right) style.paddingRight = right;
  // The first row of EACH paragraph (each-line: after every forced break),
  // as the writers indent it -- a box's paragraphs are rows of one block.
  const first = px(typography.firstLineIndent);
  if (first) style.textIndent = `${first} each-line`;
  if (typography.distributed) {
    style.textAlign = "justify";
    style.textAlignLast = "justify";
  }
  if (typography.textPaint && pixelsPerPoint && pixelsPerPoint > 0) {
    // WordArt as the other editors draw it (officeTextPaintCss), its sizes
    // those it has on paper; the caret in the words' own ink, which a
    // gradient or pattern fill leaves transparent.
    Object.assign(style, officeTextPaintCss(pdfTextBoxTextPaintAtScale(typography.textPaint, pixelsPerPoint)));
    style.caretColor = typography.textPaint.fill !== "none" ? typography.textPaint.fill
      : typography.textPaint.outline !== "none" ? typography.textPaint.outline : "currentColor";
  }
  if (typography.columnCount) {
    // The rows fill each column to the frame's foot before the next (the
    // box's height is its frame's: PdfTextBoxOverlay), as the writers set
    // them.
    style.columnCount = typography.columnCount;
    style.columnGap = `${(typography.columnSpacing ?? 0) * pageHeightPx}px`;
    style.columnFill = "auto";
  }
  if (typography.writingMode === "vertical-rl") {
    // The browser's vertical columns are the writers' (text-box-vertical.ts):
    // a Latin word turned, an ideograph upright; under the column is its
    // right side.
    style.writingMode = "vertical-rl";
    style.textOrientation = "mixed";
    style.textUnderlinePosition = "right";
  }
  if (typography.verticalAlign) {
    // The frame is at least the box's height (min-height): its rows stand at
    // the middle or the foot of it, as the writer offsets them.
    style.display = "flex";
    style.flexDirection = "column";
    style.justifyContent = typography.verticalAlign === "middle" ? "center" : "flex-end";
  }
  return style;
}

/**
 * A new text box is committed as plain page text (a native content append)
 * when it is plain: that writer sets only the text, size and colour. A box
 * the reader already formatted in a way it cannot write (bold, italic,
 * underline, strikethrough, alignment, rotation, mixed runs, or any of the
 * formats above) stays a text box, whose writer carries them; committing it
 * as plain text dropped them silently.
 */
export function pdfCreatedTextBoxNeedsTextBoxWriter(
  edit: PdfTextBoxTypographyFields & {
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    justify?: boolean;
    align?: string;
    angle?: number;
    runs?: readonly unknown[];
    /** An Office family: only the text-box writer embeds its program. */
    officeFamily?: string;
    /** A tab in it: only the text-box writers set tabs at their stops. */
    text?: string;
  },
): boolean {
  return Boolean(
    edit.officeFamily ||
    (typeof edit.text === "string" && edit.text.includes("\t")) ||
    edit.bold ||
      edit.italic ||
      edit.underline ||
      edit.strikethrough ||
      edit.justify ||
      (edit.align && edit.align !== "left") ||
      (typeof edit.angle === "number" && edit.angle % 360 !== 0) ||
      (edit.runs && edit.runs.length > 0) ||
      edit.writingMode === "vertical-rl" ||
      hasPdfTextBoxTypography(readPdfTextBoxTypography(edit, edit.fontSize ?? 0.02)),
  );
}

/**
 * Horizontal scale for a browser box: the text is laid out in a frame 1/k as
 * wide and scaled by k about the frame's centre (the same origin the box's
 * rotation turns about), so it wraps where the writer wraps and fills the
 * box's own width on screen.
 */
export function pdfTextBoxHorizontalScaleFrame(
  left: number,
  width: number,
  horizontalScale: number | undefined,
): { left: number; width: number; scale?: string } {
  const k = (horizontalScale ?? 100) / 100;
  if (!(k > 0) || Math.abs(k - 1) < 1e-6) return { left, width };
  const inner = width / k;
  return { left: left + (width - inner) / 2, width: inner, scale: `${k} 1` };
}
