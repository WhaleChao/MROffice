/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeDashLengthsWithCaps, officeDashUnits } from "../office-stroke-dash";

/** The outline's colour (#rrggbb; #rrggbbaa when it is see-through). */
export const OFFICE_TEXT_OUTLINE_COLOR = "--office-text-outline-color";
/** The outline's width in points (a bare number). */
export const OFFICE_TEXT_OUTLINE_WIDTH = "--office-text-outline-width";
/** Rings of drop-shadows that stand in for an outline round clipped (gradient
 * or pattern) words: the overlay's stylesheet turns them off once it draws. */
export const OFFICE_TEXT_OUTLINE_RINGS = "--office-text-outline-rings";
/** The outline's join when it is not round: `miter <limit>` (the limit in
 * widths) or `bevel`. */
export const OFFICE_TEXT_OUTLINE_JOIN = "--office-text-outline-join";
/** A dashed outline: its caps (`butt` / `round` / `square`) and its dash
 * and gap lengths in widths, caps counted inside each dash as Office counts
 * them (`butt 4 3` is Dash). */
export const OFFICE_TEXT_OUTLINE_DASH = "--office-text-outline-dash";
/** Set by the overlay: the drawn band, and its box in the element's own px. */
export const OFFICE_TEXT_OUTLINE_IMAGE = "--office-text-outline-image";
/** Set by the overlay while its picture fills the words too (an installed
 * font draws them, whose outlines the page cannot read): the words' own
 * fill is see-through then. */
export const OFFICE_TEXT_OUTLINE_FILLED = "--office-text-outline-filled";

/**
 * How an outline's line is drawn beyond its colour and width, in Office's
 * words (Format Text Effects > Text Outline: Join type, Cap type, Dash type;
 * `a:ln`'s `a:round` / `a:bevel` / `a:miter@lim`, `@cap`, `a:prstDash`).
 * PowerPoint follows each (text-outline-join-probe); a line that states no
 * join is round.
 */
export interface OfficeTextOutlineLine {
  readonly join?: "round" | "bevel" | "miter";
  /** `a:miter@lim`, in 1000ths of a percent of the width (800000: 8 widths). */
  readonly miterLimit?: number;
  readonly cap?: "flat" | "round" | "square";
  /** `a:prstDash@val`. */
  readonly dash?: string;
}

/** A WordArt statement's line (`stated.line`: cap `flat`/`rnd`/`sq`, join
 * `{ kind, limit }`, dash) as an outline's line. */
export function officeTextOutlineLineOfStated(line: unknown): OfficeTextOutlineLine | undefined {
  if (!line || typeof line !== "object") return undefined;
  const record = line as { cap?: unknown; join?: unknown; dash?: unknown };
  const join = record.join && typeof record.join === "object" ? record.join as { kind?: unknown; limit?: unknown } : null;
  const out: { -readonly [K in keyof OfficeTextOutlineLine]: OfficeTextOutlineLine[K] } = {};
  if (join && (join.kind === "round" || join.kind === "bevel" || join.kind === "miter")) {
    out.join = join.kind;
    if (join.kind === "miter" && typeof join.limit === "number" && join.limit > 0) out.miterLimit = join.limit;
  }
  if (record.cap === "flat" || record.cap === "rnd" || record.cap === "sq") {
    out.cap = record.cap === "rnd" ? "round" : record.cap === "sq" ? "square" : "flat";
  }
  if (typeof record.dash === "string" && record.dash !== "solid") out.dash = record.dash;
  return Object.keys(out).length ? out : undefined;
}

/** A run's or a shape's outline line fields (`outlineJoin`,
 * `outlineMiterLimit`, `outlineCap`, `outlineDash`, as a shape's line
 * states them) as an outline's line. */
export function officeTextOutlineLineOfFields(fields: unknown): OfficeTextOutlineLine | undefined {
  if (!fields || typeof fields !== "object") return undefined;
  const record = fields as { outlineJoin?: unknown; outlineMiterLimit?: unknown; outlineCap?: unknown; outlineDash?: unknown };
  const out: { -readonly [K in keyof OfficeTextOutlineLine]: OfficeTextOutlineLine[K] } = {};
  if (record.outlineJoin === "round" || record.outlineJoin === "bevel" || record.outlineJoin === "miter") {
    out.join = record.outlineJoin;
    if (record.outlineJoin === "miter" && typeof record.outlineMiterLimit === "number" && record.outlineMiterLimit > 0) {
      out.miterLimit = record.outlineMiterLimit;
    }
  }
  if (record.outlineCap === "flat" || record.outlineCap === "round" || record.outlineCap === "square") out.cap = record.outlineCap;
  if (typeof record.outlineDash === "string" && record.outlineDash !== "solid") out.dash = record.outlineDash;
  return Object.keys(out).length ? out : undefined;
}

/** The SVG / canvas stroke an outline's line is drawn with: its join and
 * miter limit (widths), and a dash as lengths in widths with its cap. */
export interface OfficeTextOutlineStroke {
  readonly join: "round" | "bevel" | "miter";
  readonly miterLimit: number;
  readonly cap: "butt" | "round" | "square";
  /** Dash and gap lengths in widths; null for a solid line. */
  readonly dash: readonly number[] | null;
}

/** DrawingML's default miter limit (`a:miter` with no `@lim`): 800%. */
const DEFAULT_MITER_LIMIT = 8;

/** The stroke of an outline's line (null: round, solid). */
export function officeTextOutlineStroke(line: OfficeTextOutlineLine | undefined | null): OfficeTextOutlineStroke {
  const join = line?.join ?? "round";
  const cap = line?.cap === "round" ? "round" : line?.cap === "square" ? "square" : "butt";
  const units = line?.dash ? officeDashUnits(line.dash) : null;
  return {
    join,
    miterLimit: join === "miter" && line?.miterLimit ? line.miterLimit / 100_000 : DEFAULT_MITER_LIMIT,
    cap,
    dash: units ? officeDashLengthsWithCaps(units, 1, cap !== "butt").map((length) => Math.round(length * 1000) / 1000) : null,
  };
}

/** The marks an outline's line states: its join and dash, `initial` for a
 * round, solid one (so words outlined inside other outlined words do not
 * take theirs: custom properties inherit). */
function lineMarks(line: OfficeTextOutlineLine | undefined | null): Record<string, string> {
  const stroke = officeTextOutlineStroke(line);
  return {
    [OFFICE_TEXT_OUTLINE_JOIN]: stroke.join === "round" ? "initial" : stroke.join === "miter" ? `miter ${stroke.miterLimit}` : "bevel",
    [OFFICE_TEXT_OUTLINE_DASH]: stroke.dash ? `${stroke.cap} ${stroke.dash.join(" ")}` : "initial",
  };
}

/** An element's outline stroke as its marks state it (computed values of
 * `--office-text-outline-join` / `-dash`). */
export function officeTextOutlineStrokeOfMarks(join: string, dash: string): OfficeTextOutlineStroke {
  const joinWords = join.trim().split(/\s+/);
  const dashWords = dash.trim().split(/\s+/).filter(Boolean);
  const limit = Number(joinWords[1]);
  const lengths = dashWords.slice(1).map(Number);
  const cap = dashWords[0] === "round" || dashWords[0] === "square" ? dashWords[0] : "butt";
  return {
    join: joinWords[0] === "miter" || joinWords[0] === "bevel" ? joinWords[0] : "round",
    miterLimit: limit > 0 ? limit : DEFAULT_MITER_LIMIT,
    cap,
    dash: lengths.length >= 2 && lengths.every((length) => length >= 0 && Number.isFinite(length)) ? lengths : null,
  };
}

let installing = false;

/** Start the overlay (once per page, lazily loaded). */
export function installOfficeTextOutline(): void {
  if (installing || typeof window === "undefined" || typeof document === "undefined") return;
  installing = true;
  void import("./office-text-outline-overlay").then(
    (overlay) => overlay.startOfficeTextOutlineOverlay(document),
    () => {
      installing = false;
    },
  );
}

/**
 * The custom properties that state an element's text outline, for the
 * overlay to draw: `color` (#rrggbb) and `widthPt` (> 0), and its `line`'s
 * join and dash where they are not round and solid. Ligatures are off where
 * words are outlined, as Office draws them (the band is placed character by
 * character).
 */
export function officeTextOutlineMarks(color: string, widthPt: number, line?: OfficeTextOutlineLine | null): Record<string, string> {
  installOfficeTextOutline();
  return {
    [OFFICE_TEXT_OUTLINE_COLOR]: color,
    [OFFICE_TEXT_OUTLINE_WIDTH]: String(Math.round(widthPt * 1000) / 1000),
    ...lineMarks(line),
    fontVariantLigatures: "none",
  };
}

/** The same marks cleared, for a layer that keeps the keys while another
 * draws the words. */
export const OFFICE_TEXT_OUTLINE_CLEARED: Record<string, string> = {
  [OFFICE_TEXT_OUTLINE_COLOR]: "initial",
  [OFFICE_TEXT_OUTLINE_WIDTH]: "initial",
  [OFFICE_TEXT_OUTLINE_JOIN]: "initial",
  [OFFICE_TEXT_OUTLINE_DASH]: "initial",
  [OFFICE_TEXT_OUTLINE_RINGS]: "initial",
};

/**
 * An outline as CSS: the overlay's marks, with `fallback` (a
 * `-webkit-text-stroke` value) drawing it until the overlay has, and
 * wherever it cannot.
 */
export function officeTextOutlineCss(
  color: string, widthPt: number, fallback: string, line?: OfficeTextOutlineLine | null,
): Record<string, string> {
  return { WebkitTextStroke: fallback, ...officeTextOutlineMarks(color, widthPt, line) };
}

let pendingProbe: (() => boolean) | null = null;

/** The overlay's: whether it has outlines still to draw. */
export function setOfficeTextOutlinePendingProbe(probe: () => boolean): void {
  pendingProbe = probe;
}

function outlinesPending(): boolean {
  if (typeof document === "undefined") return false;
  if (pendingProbe) return pendingProbe();
  // Not loaded yet: pending while the page states an outline it will draw.
  return installing && document.querySelector(`[style*="${OFFICE_TEXT_OUTLINE_WIDTH}:"]`) !== null;
}

/**
 * Settles once every stated outline is drawn (or `timeoutMs` passes): a
 * print or a picture of the page is taken after it, so it shows the
 * outlines Office draws, not the stand-in strokes.
 */
export function officeTextOutlinesSettled(timeoutMs = 5_000): Promise<void> {
  if (!outlinesPending()) return Promise.resolve();
  const deadline = Date.now() + timeoutMs;
  return new Promise((resolve) => {
    const check = () => {
      if (!outlinesPending() || Date.now() >= deadline) resolve();
      else setTimeout(check, 30);
    };
    setTimeout(check, 30);
  });
}

/** Whether outlines are still to be drawn (a print source's readiness). */
export function officeTextOutlinesPending(): boolean {
  return outlinesPending();
}

/** One glyph of a title picture: its outline (font units, y up) placed at
 * its pen on the baseline, `scale` px per unit. */
export interface OfficeTextOutlineSvgGlyph {
  readonly key: string;
  readonly d: string;
  readonly x: number;
  readonly scale: number;
  readonly skew: number;
}

type SvgGlyphs = (
  text: string, family: string, weight: number, italic: boolean, sizePx: number,
) => { glyphs: OfficeTextOutlineSvgGlyph[]; width: number } | null;
let svgGlyphs: SvgGlyphs | null = null;

/** The overlay's: lays a title picture's glyphs out from its kept outlines. */
export function setOfficeTextOutlineSvgGlyphs(layout: SvgGlyphs): void {
  svgGlyphs = layout;
}

/**
 * A title picture's glyphs (`officeTextPaintSvg`), laid out from the faces
 * the page serves, when every one of their outlines is to hand; null while
 * any is not (they are asked for, and the next picture has them).
 */
export function officeTextOutlineSvgGlyphs(
  text: string, family: string, weight: number, italic: boolean, sizePx: number,
): { glyphs: OfficeTextOutlineSvgGlyph[]; width: number } | null {
  installOfficeTextOutline();
  return svgGlyphs ? svgGlyphs(text, family, weight, italic, sizePx) : null;
}

type CanvasBand = (context: CanvasRenderingContext2D, text: string, x: number, baseline: number, width: number) => boolean;
type CanvasFill = (context: CanvasRenderingContext2D, text: string, x: number, baseline: number) => boolean;
let canvasBand: CanvasBand | null = null;
let canvasFill: CanvasFill | null = null;

/** The overlay's: strokes a canvas run's band from the kept outlines. */
export function setOfficeTextOutlineCanvasBand(draw: CanvasBand, fill?: CanvasFill): void {
  canvasBand = draw;
  if (fill) canvasFill = fill;
}

/**
 * An outlined canvas run's fill from the same kept outlines as its band
 * (`officeTextOutlineCanvas`), in the context's fill style, where an
 * installed font would draw the words otherwise (its outlines are not the
 * band's: an installed Calibri under a band of the served Carlito). False
 * where the context's own fillText draws the glyphs the band follows, or
 * the outlines are not all kept yet: the caller fills the text itself.
 */
export function officeTextOutlineCanvasFill(
  context: CanvasRenderingContext2D, text: string, x: number, baseline: number,
): boolean {
  installOfficeTextOutline();
  return canvasFill ? canvasFill(context, text, x, baseline) : false;
}

/**
 * A canvas's words outlined as Office draws them (a 2-D canvas standing in
 * for the page: warped WordArt, a PDF text box's raster): the band of `text`
 * set in the context's font from (`x`, `baseline`), `width` canvas units
 * wide, centred along each glyph's merged outline, in the context's stroke
 * style, join, cap and dash (`applyOfficeTextOutlineStroke`). False when the
 * glyphs' outlines are not all kept yet (they are asked for): the caller
 * strokes the text itself.
 */
export function officeTextOutlineCanvas(
  context: CanvasRenderingContext2D, text: string, x: number, baseline: number, width: number,
): boolean {
  installOfficeTextOutline();
  return canvasBand ? canvasBand(context, text, x, baseline, width) : false;
}

/** A context set to stroke an outline's line `width` units wide: its join
 * and miter limit, and its dash and cap (round and solid where it states
 * none). */
export function applyOfficeTextOutlineStroke(
  context: CanvasRenderingContext2D, line: OfficeTextOutlineLine | undefined | null, width: number,
): void {
  const stroke = officeTextOutlineStroke(line);
  context.lineJoin = stroke.join;
  context.miterLimit = stroke.miterLimit;
  context.lineCap = stroke.dash ? stroke.cap : "round";
  context.setLineDash(stroke.dash ? stroke.dash.map((length) => length * width) : []);
}

/** One glyph's outline for a drawing that places its glyphs itself (the
 * slide text engine's SVG): its outline (font units, y up) and px per unit. */
export interface OfficeTextOutlineGlyph {
  readonly d: string;
  readonly scale: number;
  readonly skew: number;
  /** An installed font draws the character, not this outline's file: the
   * caller fills it from the outline too. */
  readonly installed?: boolean;
}
type GlyphLookup = (text: string, family: string, weight: number, italic: boolean, sizePx: number) => OfficeTextOutlineGlyph | "pending" | null;
let glyphLookup: GlyphLookup | null = null;
let glyphEpoch = 0;
const glyphListeners = new Set<() => void>();

/** The overlay's: looks a grapheme's glyph outline up. */
export function setOfficeTextOutlineGlyphLookup(lookup: GlyphLookup): void {
  glyphLookup = lookup;
  noteOfficeTextOutlineGlyphs();
}

/** The overlay's: glyph outlines asked for have come. */
export function noteOfficeTextOutlineGlyphs(): void {
  glyphEpoch += 1;
  for (const listener of [...glyphListeners]) listener();
}

/** For a drawing that waits on glyph outlines (useSyncExternalStore). */
export function subscribeOfficeTextOutlineGlyphs(listener: () => void): () => void {
  glyphListeners.add(listener);
  return () => glyphListeners.delete(listener);
}
export function officeTextOutlineGlyphsEpoch(): number {
  return glyphEpoch;
}

/**
 * The outline of the glyph `text` (one grapheme) is drawn in, in the face
 * the page serves for that family, weight and slant: "pending" while it is
 * read (it is asked for; subscribers hear when it comes), null when the
 * face is not one the page serves (the caller strokes the text itself).
 */
export function officeTextOutlineGlyph(
  text: string, family: string, weight: number, italic: boolean, sizePx: number,
): OfficeTextOutlineGlyph | "pending" | null {
  installOfficeTextOutline();
  return glyphLookup ? glyphLookup(text, family, weight, italic, sizePx) : "pending";
}
