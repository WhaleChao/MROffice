/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The ONE painter for warped text (WordArt Transform) in the three editors.
 *
 * The browser lays the words out exactly as it lays out any shape's words —
 * same fonts, wrapping, alignment and runs — inside a SOURCE element the host
 * already renders. This module reads that layout back glyph by glyph
 * (`collectTextWarpGlyphs`) and paints it onto a canvas through the shared
 * geometry in `office-text-warp.ts`, whose rules were measured against Word
 * (docs/office-graphics/TEXT-WARP.md):
 *
 * - envelope / fit: the lines are grouped over the preset's bands (earlier
 *   bands first); each group's TIGHT INK rectangle is cut into thin vertical
 *   strips and every strip is mapped onto the band by arc length, two affine
 *   triangles per strip — the standard's point mapping, sampled;
 * - follow: every glyph is drawn rigidly, rotated onto the path and scaled
 *   by `R / (R + h)` (`textWarpFollowGlyphs`).
 *
 * The source stays in the DOM (selection, caret, find, accessibility); the
 * host only makes its glyphs transparent while the canvas shows them.
 */

import {
  OFFICE_TEXT_OUTLINE_COLOR,
  OFFICE_TEXT_OUTLINE_DASH,
  OFFICE_TEXT_OUTLINE_JOIN,
  OFFICE_TEXT_OUTLINE_WIDTH,
  officeTextOutlineCanvas,
  officeTextOutlineCanvasFill,
  officeTextOutlineStrokeOfMarks,
  type OfficeTextOutlineStroke,
} from "./office-text-outline/office-text-outline-css";
import {
  type OfficeTextWarp,
  type TextWarpPath,
  textWarpBands,
  textWarpDrawing,
  textWarpEnvelopePoint,
  textWarpFollowGlyphs,
  textWarpGroupSizes,
} from "./office-text-warp";
import { officeGradientAngleForBox, type OfficeShapeGradient } from "./office-shape-paint";

export type { OfficeTextWarp };

/** One glyph as the browser laid it out, in the text box's CSS pixels. */
export interface TextWarpGlyph {
  readonly text: string;
  /** Left edge of the glyph's advance. */
  readonly x: number;
  /** Baseline. */
  readonly baseline: number;
  readonly advance: number;
  readonly font: string;
  readonly fill: string;
  readonly stroke: string | null;
  readonly strokeWidth: number;
  /** The outline is Office's: centred on the glyph, over its fill. */
  readonly strokeOver?: boolean;
  /** Its join, cap and dash (round and solid where none is stated). */
  readonly strokeLine?: OfficeTextOutlineStroke;
  readonly line: number;
  /** Ascent / descent of the glyph's font at its size (line metrics). */
  readonly ascent: number;
  readonly descent: number;
  readonly fontSize: number;
}

export interface TextWarpPaintInput {
  readonly warp: OfficeTextWarp;
  /** The geometry box (the text box inside the insets), CSS px. */
  readonly width: number;
  readonly height: number;
  readonly glyphs: readonly TextWarpGlyph[];
  /** Device pixels per CSS pixel of the target canvas. */
  readonly scale: number;
  /** Where the geometry box's origin sits on the canvas, CSS px. */
  readonly originX: number;
  readonly originY: number;
  readonly align?: "left" | "center" | "right" | "justify";
  /** The words' outer shadow (the shared text paint's), CSS px. */
  readonly shadow?: TextWarpShadow | null;
  /**
   * The words' gradient fill (the shared text paint's `fillGradient`): one
   * ramp over the geometry box, as Word paints a WordArt's -- not one per
   * glyph. The glyphs' own fill colour then only says where the words are.
   */
  readonly fillGradient?: OfficeShapeGradient | null;
}

/**
 * The shared text shadow (`OfficeTextShadow`), as the painter draws it: a
 * copy of the warped words, transformed about one of the geometry box's nine
 * anchors and then moved by its offset. The transform is DrawingML's
 * (`a:outerShdw` / `w14:shadow`), as Word was measured drawing it
 * (`office_text_paint.SHADOW_TRANSFORM_KEYS`): with `(x, y)` from the anchor
 * and y down, `x' = sx*x + tan(kx)*y`, `y' = tan(ky)*x + sy*y`.
 */
export interface TextWarpShadow {
  readonly color: string;
  readonly blur: number;
  readonly offsetX: number;
  readonly offsetY: number;
  readonly opacity: number;
  /** `@sx`, percent; absent = 100. */
  readonly scale?: number;
  /** `@sy`, percent, when it differs from `scale`; negative = flipped. */
  readonly scaleY?: number;
  /** `@kx` / `@ky`, degrees. */
  readonly skewX?: number;
  readonly skewY?: number;
  /** `@algn`; absent = the centre. */
  readonly align?: "tl" | "t" | "tr" | "l" | "ctr" | "r" | "bl" | "b" | "br";
}

/** The shadow's 2x2 transform `[sx, tan(ky), tan(kx), sy]` (canvas a, b, c, d). */
export function textWarpShadowMatrix(shadow: TextWarpShadow): [number, number, number, number] {
  const sx = (shadow.scale ?? 100) / 100;
  const sy = (shadow.scaleY ?? shadow.scale ?? 100) / 100;
  const radians = (degrees: number | undefined) => ((degrees ?? 0) * Math.PI) / 180;
  return [sx, Math.tan(radians(shadow.skewY)), Math.tan(radians(shadow.skewX)), sy];
}

/** Where on a `width` x `height` box the shadow's anchor sits. */
export function textWarpShadowAnchor(shadow: TextWarpShadow, width: number, height: number): [number, number] {
  const align = shadow.align ?? "ctr";
  const x = align.endsWith("l") ? 0 : align.endsWith("r") && align !== "ctr" ? width : width / 2;
  const y = align.startsWith("t") ? 0 : align.startsWith("b") ? height : height / 2;
  return [x, y];
}

/** How far past the geometry box the shadow can reach, CSS px. */
export function textWarpShadowBleed(shadow: TextWarpShadow, width: number, height: number): number {
  const [a, b, c, d] = textWarpShadowMatrix(shadow);
  const [ax, ay] = textWarpShadowAnchor(shadow, width, height);
  let reach = 0;
  for (const [x, y] of [[0, 0], [width, 0], [0, height], [width, height]]) {
    const dx = x - ax;
    const dy = y - ay;
    const tx = ax + a * dx + c * dy + shadow.offsetX;
    const ty = ay + b * dx + d * dy + shadow.offsetY;
    reach = Math.max(reach, -tx, tx - width, -ty, ty - height);
  }
  return Math.ceil(Math.max(0, reach) + shadow.blur * 2);
}

/** The browser's canvas, narrowed to what the painter needs (testable). */
type Canvas2D = CanvasRenderingContext2D;

const MAX_STRIPS = 720;

/** Lines in order, each the glyphs it holds. */
export function textWarpLines(glyphs: readonly TextWarpGlyph[]): TextWarpGlyph[][] {
  const lines = new Map<number, TextWarpGlyph[]>();
  for (const glyph of glyphs) {
    const line = lines.get(glyph.line) ?? [];
    line.push(glyph);
    lines.set(glyph.line, line);
  }
  return [...lines.entries()].sort((a, b) => a[0] - b[0]).map(([, line]) => line);
}

/**
 * Which part of a glyph a pass draws: all of it, its outline, its fill, or
 * its whole `silhouette` (outline and an opaque fill: where the glyph IS,
 * whatever it is painted with -- a gradient-filled glyph's own fill colour is
 * transparent, and a see-through one is faint).
 */
type GlyphPart = "all" | "stroke" | "fill" | "silhouette";

function drawGlyph(context: Canvas2D, glyph: TextWarpGlyph, x: number, y: number, part: GlyphPart = "all"): void {
  context.font = glyph.font;
  const stroke = () => {
    if (part === "fill" || !glyph.stroke || !(glyph.strokeWidth > 0)) return;
    const line = glyph.strokeLine;
    context.lineJoin = line?.join ?? "round";
    context.miterLimit = line?.miterLimit ?? 10;
    context.lineCap = line?.dash ? line.cap : "round";
    context.setLineDash(line?.dash ? line.dash.map((length) => length * glyph.strokeWidth) : []);
    context.lineWidth = glyph.strokeWidth;
    context.strokeStyle = glyph.stroke;
    // Office's outline follows each glyph's merged outline (no line inside
    // where its contours overlap) once the glyphs' outlines are kept.
    if (!(glyph.strokeOver && officeTextOutlineCanvas(context as CanvasRenderingContext2D, glyph.text, x, y, glyph.strokeWidth))) {
      context.strokeText(glyph.text, x, y);
    }
    context.setLineDash([]);
  };
  if (!glyph.strokeOver) stroke();
  if (part !== "stroke") {
    // A fill-only pass draws where the words ARE (it is painted afterwards).
    context.fillStyle = part === "all" ? glyph.fill : "#000000";
    // An outlined glyph is filled from its band's own outline where an
    // installed font would draw another one under it.
    const outlined = glyph.strokeOver && Boolean(glyph.stroke) && glyph.strokeWidth > 0;
    if (!(outlined && officeTextOutlineCanvasFill(context as CanvasRenderingContext2D, glyph.text, x, y))) context.fillText(glyph.text, x, y);
  }
  if (glyph.strokeOver) stroke();
}

function inkBounds(context: Canvas2D, width: number, height: number): [number, number, number, number] | null {
  if (width <= 0 || height <= 0) return null;
  const data = context.getImageData(0, 0, width, height).data;
  let left = width;
  let right = -1;
  let top = height;
  let bottom = -1;
  for (let y = 0; y < height; y += 1) {
    const row = y * width * 4;
    for (let x = 0; x < width; x += 1) {
      if (data[row + x * 4 + 3] > 8) {
        if (x < left) left = x;
        if (x > right) right = x;
        if (y < top) top = y;
        if (y > bottom) bottom = y;
      }
    }
  }
  return right < 0 ? null : [left, top, right + 1, bottom + 1];
}

/** The 2×3 affine transform taking triangle `s` onto triangle `d`. */
export function triangleTransform(
  s: readonly [number, number][],
  d: readonly [number, number][],
): [number, number, number, number, number, number] | null {
  const [[x0, y0], [x1, y1], [x2, y2]] = s;
  const det = (x1 - x0) * (y2 - y0) - (x2 - x0) * (y1 - y0);
  if (Math.abs(det) < 1e-9) return null;
  const [[u0, v0], [u1, v1], [u2, v2]] = d;
  const a = ((u1 - u0) * (y2 - y0) - (u2 - u0) * (y1 - y0)) / det;
  const c = ((u2 - u0) * (x1 - x0) - (u1 - u0) * (x2 - x0)) / det;
  const b = ((v1 - v0) * (y2 - y0) - (v2 - v0) * (y1 - y0)) / det;
  const dd = ((v2 - v0) * (x1 - x0) - (v1 - v0) * (x2 - x0)) / det;
  return [a, b, c, dd, u0 - a * x0 - c * y0, v0 - b * x0 - dd * y0];
}

function grow(points: [number, number][], by: number): [number, number][] {
  const cx = (points[0][0] + points[1][0] + points[2][0]) / 3;
  const cy = (points[0][1] + points[1][1] + points[2][1]) / 3;
  return points.map(([x, y]) => {
    const dx = x - cx;
    const dy = y - cy;
    const length = Math.hypot(dx, dy) || 1;
    return [x + (dx / length) * by, y + (dy / length) * by];
  });
}

function paintEnvelope(
  context: Canvas2D,
  input: TextWarpPaintInput,
  document: Document,
  part: GlyphPart = "all",
): void {
  const bands = textWarpBands(input.warp.preset, input.width, input.height, input.warp.adjustments, 48);
  if (!bands?.length) return;
  const lines = textWarpLines(input.glyphs);
  const sizes = textWarpGroupSizes(lines.length, bands.length);
  const scale = input.scale;
  let cursor = 0;
  bands.forEach(([top, bottom], bandIndex) => {
    const group = lines.slice(cursor, cursor + sizes[bandIndex]).flat();
    cursor += sizes[bandIndex];
    if (!group.length) return;
    // The group's own layout, unwarped, on a scratch canvas big enough for
    // every glyph's line box. The words are usually STRETCHED onto the band,
    // so they are drawn at the resolution they end up at, not at their size.
    const minX = Math.min(...group.map((g) => g.x)) - 4;
    const maxX = Math.max(...group.map((g) => g.x + g.advance)) + 4;
    const minY = Math.min(...group.map((g) => g.baseline - g.ascent * 1.4)) - 4;
    const maxY = Math.max(...group.map((g) => g.baseline + g.descent * 1.4)) + 4;
    const render = (resolution: number) => {
      const sourceWidth = Math.max(1, Math.ceil((maxX - minX) * resolution));
      const sourceHeight = Math.max(1, Math.ceil((maxY - minY) * resolution));
      if (sourceWidth * sourceHeight > 16_000_000) return null;
      const canvas = document.createElement("canvas");
      canvas.width = sourceWidth;
      canvas.height = sourceHeight;
      const sourceContext = canvas.getContext("2d", { willReadFrequently: true });
      if (!sourceContext) return null;
      sourceContext.setTransform(resolution, 0, 0, resolution, -minX * resolution, -minY * resolution);
      sourceContext.textBaseline = "alphabetic";
      // The band is fitted to the ink of the WHOLE glyphs -- their
      // silhouette -- whichever part this pass draws and whatever paints
      // them: an outline pass and a fill pass must land on the same strips
      // (or the outline would sit beside its own fill), and words whose own
      // fill colour is transparent still have a place.
      for (const glyph of group) drawGlyph(sourceContext, glyph, glyph.x, glyph.baseline, "silhouette");
      const bounds = inkBounds(sourceContext, sourceWidth, sourceHeight);
      if (!bounds) return null;
      sourceContext.save();
      sourceContext.setTransform(1, 0, 0, 1, 0, 0);
      sourceContext.clearRect(0, 0, sourceWidth, sourceHeight);
      sourceContext.restore();
      for (const glyph of group) drawGlyph(sourceContext, glyph, glyph.x, glyph.baseline, part);
      return { canvas, bounds, resolution };
    };
    let drawn = render(scale);
    if (!drawn) return;
    const bandHeight = Math.max(
      ...[0, 0.25, 0.5, 0.75, 1].map((u) => {
        const [tx, ty] = textWarpEnvelopePoint(top as TextWarpPath, bottom as TextWarpPath, u, 0);
        const [bx, by] = textWarpEnvelopePoint(top as TextWarpPath, bottom as TextWarpPath, u, 1);
        return Math.hypot(bx - tx, by - ty);
      }),
    );
    const stretch = Math.max(
      (bandHeight * scale) / Math.max(1, drawn.bounds[3] - drawn.bounds[1]),
      (input.width * scale) / Math.max(1, drawn.bounds[2] - drawn.bounds[0]),
    );
    if (stretch > 1.25) drawn = render(scale * Math.min(8, stretch)) ?? drawn;
    const source = drawn.canvas;
    const [left, inkTop, right, inkBottom] = drawn.bounds;
    const inkWidth = right - left;
    // About one strip per two device pixels of the band, never more strips
    // than the source has columns.
    const strips = Math.max(1, Math.min(MAX_STRIPS, inkWidth, Math.ceil((input.width * scale) / 2)));
    const toCanvas = ([x, y]: readonly [number, number]): [number, number] => [
      (input.originX + x) * scale,
      (input.originY + y) * scale,
    ];
    const pointAt = (u: number, v: number) =>
      toCanvas(textWarpEnvelopePoint(top as TextWarpPath, bottom as TextWarpPath, u, v));
    for (let index = 0; index < strips; index += 1) {
      const u0 = index / strips;
      const u1 = (index + 1) / strips;
      const sx0 = left + u0 * inkWidth;
      const sx1 = left + u1 * inkWidth;
      const tl = pointAt(u0, 0);
      const tr = pointAt(u1, 0);
      const br = pointAt(u1, 1);
      const bl = pointAt(u0, 1);
      for (const [src, dst] of [
        [[[sx0, inkTop], [sx1, inkTop], [sx0, inkBottom]], [tl, tr, bl]],
        [[[sx1, inkTop], [sx1, inkBottom], [sx0, inkBottom]], [tr, br, bl]],
      ] as [[number, number][], [number, number][]][]) {
        const transform = triangleTransform(src, dst);
        if (!transform) continue;
        const clip = grow(dst, 0.6);
        context.save();
        context.beginPath();
        context.moveTo(clip[0][0], clip[0][1]);
        context.lineTo(clip[1][0], clip[1][1]);
        context.lineTo(clip[2][0], clip[2][1]);
        context.closePath();
        context.clip();
        context.setTransform(...transform);
        context.drawImage(source, 0, 0);
        context.restore();
      }
    }
  });
}

function paintFollow(
  context: Canvas2D,
  input: TextWarpPaintInput,
  document: Document,
  part: GlyphPart = "all",
): void {
  const lines = textWarpLines(input.glyphs);
  const measure = document.createElement("canvas").getContext("2d");
  const placements = textWarpFollowGlyphs(
    input.warp.preset,
    input.width,
    input.height,
    lines.map((line) => {
      let inkHeight = 0;
      if (measure) {
        for (const glyph of line) {
          measure.font = glyph.font;
          const metrics = measure.measureText(glyph.text);
          inkHeight = Math.max(inkHeight, metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent);
        }
      }
      const size = Math.max(...line.map((glyph) => glyph.fontSize));
      return {
        advances: line.map((glyph) => glyph.advance),
        size,
        inkHeight: inkHeight || size,
        lineHeight: Math.max(...line.map((glyph) => glyph.ascent + glyph.descent)),
        align: input.align ?? "center",
      };
    }),
    input.warp.adjustments,
  );
  if (!placements) return;
  context.textBaseline = "alphabetic";
  lines.forEach((line, lineIndex) => {
    const placed = placements[lineIndex];
    if (!placed) return;
    line.forEach((glyph, glyphIndex) => {
      const place = placed[glyphIndex];
      if (!place) return;
      context.save();
      context.setTransform(input.scale, 0, 0, input.scale, input.originX * input.scale, input.originY * input.scale);
      context.translate(place.x, place.y);
      context.rotate(place.angle);
      context.scale(place.scale, place.scale);
      drawGlyph(context, glyph, 0, 0, part);
      context.restore();
    });
  });
}

/**
 * Paint warped words onto `canvas` (already sized to the host, device pixels).
 * Returns false when this warp draws nothing different from plain text.
 */
export function paintOfficeTextWarp(canvas: HTMLCanvasElement, input: TextWarpPaintInput): boolean {
  const drawing = textWarpDrawing(input.warp.preset);
  const context = canvas.getContext("2d");
  if (!context || !drawing || drawing === "none") return false;
  context.setTransform(1, 0, 0, 1, 0, 0);
  context.clearRect(0, 0, canvas.width, canvas.height);
  if (!input.glyphs.length || input.width <= 0 || input.height <= 0) return true;
  const shadow = input.shadow;
  // With a shadow the words are drawn once on a layer: the layer tinted and
  // moved is the shadow, the layer itself the words (Office draws a WordArt's
  // shadow as a copy of its warped glyphs).
  const layer = shadow ? canvas.ownerDocument.createElement("canvas") : canvas;
  if (shadow) {
    layer.width = canvas.width;
    layer.height = canvas.height;
  }
  const target = layer.getContext("2d");
  if (!target) return false;
  const paintWords = (onto: Canvas2D, part: GlyphPart) => {
    if (drawing === "follow") paintFollow(onto, input, canvas.ownerDocument, part);
    else paintEnvelope(onto, input, canvas.ownerDocument, part);
  };
  const gradient = input.fillGradient ?? null;
  const filled = gradient ? canvas.ownerDocument.createElement("canvas") : null;
  if (filled) {
    filled.width = layer.width;
    filled.height = layer.height;
  }
  const filledContext = filled ? filled.getContext("2d") : null;
  if (gradient && filled && filledContext) {
    // A gradient fill is ONE ramp over the geometry box. The outlines are
    // drawn as they are; the fills are drawn as a silhouette, the ramp is
    // painted into it, and it goes over the outlines as a glyph's fill does.
    paintWords(target, "stroke");
    paintWords(filledContext, "fill");
    const s = input.scale;
    const left = input.originX * s;
    const top = input.originY * s;
    const width = input.width * s;
    const height = input.height * s;
    // CSS's gradient line: through the centre, long enough to reach the
    // corners (0deg runs to the top, clockwise positive).
    const angle = (officeGradientAngleForBox(gradient, input.width, input.height) * Math.PI) / 180;
    const dx = Math.sin(angle);
    const dy = -Math.cos(angle);
    const half = (Math.abs(width * dx) + Math.abs(height * dy)) / 2;
    const cx = left + width / 2;
    const cy = top + height / 2;
    const ramp = filledContext.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half);
    for (const stop of gradient.stops) ramp.addColorStop(Math.min(1, Math.max(0, stop.position)), stop.color);
    filledContext.save();
    filledContext.setTransform(1, 0, 0, 1, 0, 0);
    filledContext.globalCompositeOperation = "source-in";
    filledContext.fillStyle = ramp;
    filledContext.fillRect(0, 0, filled.width, filled.height);
    filledContext.restore();
    target.save();
    target.setTransform(1, 0, 0, 1, 0, 0);
    target.drawImage(filled, 0, 0);
    target.restore();
  } else {
    paintWords(target, "all");
  }
  if (shadow && layer !== canvas) {
    const tint = canvas.ownerDocument.createElement("canvas");
    tint.width = layer.width;
    tint.height = layer.height;
    const tintContext = tint.getContext("2d");
    if (tintContext) {
      tintContext.drawImage(layer, 0, 0);
      tintContext.globalCompositeOperation = "source-in";
      tintContext.fillStyle = shadow.color;
      tintContext.fillRect(0, 0, tint.width, tint.height);
      const s = input.scale;
      const [a, b, c, d] = textWarpShadowMatrix(shadow);
      const [boxX, boxY] = textWarpShadowAnchor(shadow, input.width, input.height);
      const anchorX = (input.originX + boxX) * s;
      const anchorY = (input.originY + boxY) * s;
      context.save();
      context.globalAlpha = shadow.opacity;
      if (shadow.blur > 0) context.filter = `blur(${shadow.blur * s}px)`;
      // About the anchor: p' = anchor + M (p - anchor) + offset.
      context.setTransform(
        a,
        b,
        c,
        d,
        anchorX - a * anchorX - c * anchorY + shadow.offsetX * s,
        anchorY - b * anchorX - d * anchorY + shadow.offsetY * s,
      );
      context.drawImage(tint, 0, 0);
      context.restore();
    }
    context.drawImage(layer, 0, 0);
  }
  return true;
}

function segments(text: string): string[] {
  const Segmenter = (Intl as unknown as { Segmenter?: new (locale?: string, options?: object) => { segment(input: string): Iterable<{ segment: string }> } }).Segmenter;
  if (Segmenter) return [...new Segmenter(undefined, { granularity: "grapheme" }).segment(text)].map((part) => part.segment);
  return Array.from(text);
}

/**
 * Every visible glyph under `root`, as laid out, in `root`'s own CSS pixels
 * (a host scaled by a CSS transform is measured in its untransformed space).
 */
export function collectTextWarpGlyphs(root: HTMLElement): TextWarpGlyph[] {
  const document = root.ownerDocument;
  const view = document.defaultView;
  const bounds = root.getBoundingClientRect();
  const zoom = root.offsetWidth > 0 ? bounds.width / root.offsetWidth : 1;
  if (!view || !(zoom > 0)) return [];
  const measure = document.createElement("canvas").getContext("2d");
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const glyphs: TextWarpGlyph[] = [];
  const lineTops: number[] = [];
  const range = document.createRange();
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    const parent = node.parentElement;
    if (!parent) continue;
    const style = view.getComputedStyle(parent);
    if (style.display === "none") continue;
    const fontSize = Number.parseFloat(style.fontSize) || 16;
    const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    let ascent = fontSize * 0.9;
    let descent = fontSize * 0.25;
    if (measure) {
      measure.font = font;
      const metrics = measure.measureText("Hg");
      ascent = metrics.fontBoundingBoxAscent || ascent;
      descent = metrics.fontBoundingBoxDescent || descent;
    }
    // An outline stated as Office draws it (office-text-outline-css.ts): a
    // centred, round stroke of its width over the fill. Otherwise whatever
    // the words' CSS stroke is, under it.
    const officeWidthPt = Number.parseFloat(style.getPropertyValue(OFFICE_TEXT_OUTLINE_WIDTH));
    const officeColor = style.getPropertyValue(OFFICE_TEXT_OUTLINE_COLOR).trim();
    const office = officeWidthPt > 0 && /^#[0-9a-f]{6}$/i.test(officeColor);
    const strokeLine = office
      ? officeTextOutlineStrokeOfMarks(style.getPropertyValue(OFFICE_TEXT_OUTLINE_JOIN), style.getPropertyValue(OFFICE_TEXT_OUTLINE_DASH))
      : undefined;
    const strokeWidth = office ? (officeWidthPt * 96) / 72 : Number.parseFloat(style.getPropertyValue("-webkit-text-stroke-width")) || 0;
    const stroke = strokeWidth > 0 ? (office ? officeColor : style.getPropertyValue("-webkit-text-stroke-color") || null) : null;
    const fill = style.getPropertyValue("--office-text-warp-fill") || style.color;
    let offset = 0;
    for (const text of segments(node.data)) {
      const start = offset;
      offset += text.length;
      // A space still advances the line (it is measured) and draws nothing.
      range.setStart(node, start);
      range.setEnd(node, offset);
      const rect = range.getClientRects()[0];
      if (!rect || rect.width <= 0) continue;
      const x = (rect.left - bounds.left) / zoom;
      const top = (rect.top - bounds.top) / zoom;
      const height = rect.height / zoom;
      // The line box: glyphs whose boxes overlap one already seen share it.
      let line = lineTops.findIndex((seen) => Math.abs(seen - top) < height * 0.5);
      if (line < 0) {
        line = lineTops.length;
        lineTops.push(top);
      }
      const baseline = top + (height - (ascent + descent)) / 2 + ascent;
      glyphs.push({
        text,
        x,
        baseline,
        advance: rect.width / zoom,
        font,
        fill,
        stroke,
        strokeWidth,
        ...(office ? { strokeOver: true, strokeLine } : {}),
        line,
        ascent,
        descent,
        fontSize,
      });
    }
  }
  range.detach?.();
  // Lines in reading order: renumber by their top.
  const order = lineTops.map((top, index) => [top, index] as const).sort((a, b) => a[0] - b[0]);
  const renumber = new Map(order.map(([, index], rank) => [index, rank]));
  return glyphs.map((glyph) => ({ ...glyph, line: renumber.get(glyph.line) ?? glyph.line }));
}

export interface TextWarpSurfaceOptions {
  /** The element whose words are warped (it keeps its layout). */
  readonly source: HTMLElement;
  /**
   * The element the canvas is appended to. By default the canvas covers it
   * and it IS the geometry box; with `overSource` the canvas is laid over the
   * source's own box instead (a ProseMirror frame whose content element
   * cannot hold a foreign child).
   */
  readonly host: HTMLElement;
  readonly overSource?: boolean;
  /**
   * How the source's own glyphs give way: `opacity` (the whole element; for a
   * label that is not edited in place) or `fill` (only the glyph fill, so a
   * frame's border, background and caret stay).
   */
  readonly hide?: "opacity" | "fill";
  /** Read the insets from the source's computed padding. */
  readonly insetsFromPadding?: boolean;
  readonly warp: OfficeTextWarp;
  /** Text-box insets, CSS px: the geometry box is the host minus these. */
  readonly insets?: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };
  /** `wps:bodyPr@wrap="none"`: the box is as wide as the longest line. */
  readonly noWrap?: boolean;
  readonly align?: "left" | "center" | "right" | "justify";
  /** The words' outer shadow; the canvas grows to hold it. */
  readonly shadow?: TextWarpShadow | null;
  /** The words' gradient fill: one ramp over the geometry box. */
  readonly fillGradient?: OfficeShapeGradient | null;
}

export const TEXT_WARP_SOURCE_ATTRIBUTE = "data-office-text-warp-source";
export const TEXT_WARP_CANVAS_ATTRIBUTE = "data-office-text-warp";

/**
 * Keep a canvas over `host` painted with `source`'s words warped. Repaints on
 * resize, on font load and when the words change. Returns the disposer.
 */
export function attachOfficeTextWarp(options: TextWarpSurfaceOptions): () => void {
  const { source, host, warp } = options;
  const document = host.ownerDocument;
  const view = document.defaultView;
  if (!view || textWarpDrawing(warp.preset) === null || textWarpDrawing(warp.preset) === "none") {
    return () => {};
  }
  const canvas = document.createElement("canvas");
  canvas.setAttribute(TEXT_WARP_CANVAS_ATTRIBUTE, warp.preset);
  canvas.setAttribute("aria-hidden", "true");
  Object.assign(canvas.style, {
    position: "absolute",
    pointerEvents: "none",
  });
  // The source keeps its layout (it is what is measured, selected and read
  // aloud); only its own paint gives way to the canvas.
  const hide = options.hide ?? "opacity";
  const previousOpacity = source.style.opacity;
  const previousFill = source.style.getPropertyValue("-webkit-text-fill-color");
  source.setAttribute(TEXT_WARP_SOURCE_ATTRIBUTE, "");
  if (hide === "opacity") source.style.opacity = "0";
  else source.style.setProperty("-webkit-text-fill-color", "transparent");
  host.append(canvas);
  let frame = 0;
  const paint = () => {
    frame = 0;
    const box = options.overSource ? source : host;
    const width = box.offsetWidth;
    const height = box.offsetHeight;
    if (!(width > 0 && height > 0)) return;
    const bounds = box.getBoundingClientRect();
    const zoom = bounds.width / width || 1;
    const shadow = options.shadow ?? null;
    // Room past the box for a shadow that is offset or larger than the words.
    const bleed = shadow ? textWarpShadowBleed(shadow, width, height) : 0;
    if (options.overSource) {
      // Wherever the canvas's containing block is: measure where (0, 0)
      // lands and move by the difference.
      Object.assign(canvas.style, { left: "0px", top: "0px" });
      const origin = canvas.getBoundingClientRect();
      canvas.style.left = `${(bounds.left - origin.left) / zoom - bleed}px`;
      canvas.style.top = `${(bounds.top - origin.top) / zoom - bleed}px`;
    } else {
      canvas.style.left = `${-bleed}px`;
      canvas.style.top = `${-bleed}px`;
    }
    canvas.style.width = `${width + bleed * 2}px`;
    canvas.style.height = `${height + bleed * 2}px`;
    const scale = Math.min(4, Math.max(1, (view.devicePixelRatio || 1) * zoom));
    const pixelWidth = Math.ceil((width + bleed * 2) * scale);
    const pixelHeight = Math.ceil((height + bleed * 2) * scale);
    if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
    if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
    let insets = options.insets ?? { left: 0, top: 0, right: 0, bottom: 0 };
    if (options.insetsFromPadding) {
      // The geometry box's own padding: the source's when the words are
      // drawn over it, the host's when the host is the frame.
      const style = view.getComputedStyle(box);
      const border = (side: string) => Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0;
      const padding = (side: string) => Number.parseFloat(style.getPropertyValue(`padding-${side}`)) || 0;
      insets = {
        left: padding("left") + border("left"),
        top: padding("top") + border("top"),
        right: padding("right") + border("right"),
        bottom: padding("bottom") + border("bottom"),
      };
    }
    const sourceBounds = source.getBoundingClientRect();
    const glyphs = collectTextWarpGlyphs(source).map((glyph) => ({
      ...glyph,
      // Into the geometry box's coordinates.
      x: glyph.x + (sourceBounds.left - bounds.left) / zoom - insets.left,
      baseline: glyph.baseline + (sourceBounds.top - bounds.top) / zoom - insets.top,
    }));
    let boxWidth = width - insets.left - insets.right;
    if (options.noWrap && glyphs.length) {
      const lines = textWarpLines(glyphs);
      boxWidth = Math.max(
        ...lines.map((line) => line.reduce((sum, glyph) => sum + glyph.advance, 0)),
      );
    }
    paintOfficeTextWarp(canvas, {
      warp,
      width: boxWidth,
      height: height - insets.top - insets.bottom,
      glyphs,
      scale,
      originX: insets.left + bleed,
      originY: insets.top + bleed,
      align: options.align,
      shadow,
      fillGradient: options.fillGradient ?? null,
    });
  };
  const schedule = () => {
    if (!frame) frame = view.requestAnimationFrame(paint);
  };
  const resize = typeof view.ResizeObserver === "function" ? new view.ResizeObserver(schedule) : null;
  resize?.observe(host);
  resize?.observe(source);
  const mutation = typeof view.MutationObserver === "function" ? new view.MutationObserver(schedule) : null;
  // Text and child changes only: the source's own `style` is ours to write.
  mutation?.observe(source, { subtree: true, childList: true, characterData: true });
  const fonts = (document as Document & { fonts?: FontFaceSet }).fonts;
  fonts?.addEventListener?.("loadingdone", schedule);
  void fonts?.ready?.then(schedule);
  schedule();
  return () => {
    if (frame) view.cancelAnimationFrame(frame);
    resize?.disconnect();
    mutation?.disconnect();
    fonts?.removeEventListener?.("loadingdone", schedule);
    canvas.remove();
    source.removeAttribute(TEXT_WARP_SOURCE_ATTRIBUTE);
    if (hide === "opacity") source.style.opacity = previousOpacity;
    else source.style.setProperty("-webkit-text-fill-color", previousFill);
  };
}
