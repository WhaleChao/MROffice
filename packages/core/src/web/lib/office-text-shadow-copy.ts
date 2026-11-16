/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A perspective text shadow on words that are not warped: Text Effects >
 * Shadow > Perspective (`a:outerShdw` with `@sx`/`@sy`/`@kx`/`@ky`, Word's
 * `w14:shadow`) is a copy of the words, scaled and skewed about one of the
 * nine anchors of the text box, moved by its offset, blurred and tinted
 * (`office_text_paint.SHADOW_TRANSFORM_KEYS`). CSS's `drop-shadow` can only
 * move a copy, so such a shadow was drawn as a halo round the words.
 *
 * Shared by the deck (`PresentationTextVisual`, a text box or placeholder's
 * `textPaint`) and the document (`documentShapeExtension`, a WordArt shape's
 * `textPaint`): the words stay the page's own DOM (measured, selected,
 * edited, read aloud); a canvas BEHIND them paints only the shadow, from the
 * glyphs as laid out (`collectTextWarpGlyphs`) and the warp painter's own
 * shadow geometry (`textWarpShadowMatrix` / `textWarpShadowAnchor`); its anchor is on the words' own bounds, as Windows
 * PowerPoint casts it. A shadow with no transform stays the paint's `drop-shadow`
 * (`officeTextPaintCss`), pixel for pixel as before.
 */

import { officeTextShadowTransformed } from "./office-text-paint";
import {
  collectTextWarpGlyphs,
  textWarpShadowAnchor,
  textWarpShadowMatrix,
  type TextWarpGlyph,
  type TextWarpShadow,
} from "./office-text-warp-render";

export { officeTextShadowTransformed };

export interface TextShadowCopyInput {
  /** The text box (inside its insets), CSS px: the anchor's box. */
  readonly width: number;
  readonly height: number;
  readonly glyphs: readonly TextWarpGlyph[];
  /** Device pixels per CSS pixel of the target canvas. */
  readonly scale: number;
  /** Where the text box's origin sits on the canvas, CSS px. */
  readonly originX: number;
  readonly originY: number;
  readonly shadow: TextWarpShadow;
}

/**
 * Paint the shadow alone: the glyphs' silhouettes in the shadow's colour,
 * transformed about the anchor (`p' = anchor + M (p - anchor) + offset`),
 * blurred, at the shadow's opacity. Returns false without a 2-D context.
 */
export function paintOfficeTextShadowCopy(canvas: HTMLCanvasElement, input: TextShadowCopyInput): boolean {
  const context = canvas.getContext("2d");
  if (!context) return false;
  context.setTransform(1, 0, 0, 1, 0, 0);
  context.clearRect(0, 0, canvas.width, canvas.height);
  if (!input.glyphs.length || input.width <= 0 || input.height <= 0) return true;
  const s = input.scale;
  const layer = canvas.ownerDocument.createElement("canvas");
  layer.width = canvas.width;
  layer.height = canvas.height;
  const layerContext = layer.getContext("2d");
  if (!layerContext) return false;
  layerContext.setTransform(s, 0, 0, s, input.originX * s, input.originY * s);
  layerContext.fillStyle = input.shadow.color;
  for (const glyph of input.glyphs) {
    layerContext.font = glyph.font;
    // The silhouette: where the glyph IS, whatever it is painted with.
    layerContext.fillText(glyph.text, glyph.x, glyph.baseline);
    if (glyph.stroke && glyph.strokeWidth > 0) {
      layerContext.lineWidth = glyph.strokeWidth;
      layerContext.lineJoin = "round";
      layerContext.strokeStyle = input.shadow.color;
      layerContext.strokeText(glyph.text, glyph.x, glyph.baseline);
    }
  }
  // The anchor is on the words' own bounds, not the text box: Windows
  // PowerPoint casts a placeholder's (342 pt tall) and an auto-fit text box's
  // (61 pt) perspective shadow alike, against the words' bottom-left
  
  // Upper Right / Lower Left at 192 dpi).
  const [a, b, c, d, e, f] = officeTextShadowCopyTransform(input.shadow, officeTextShadowWordBounds(input.glyphs));
  context.save();
  context.globalAlpha = input.shadow.opacity;
  if (input.shadow.blur > 0) context.filter = `blur(${input.shadow.blur * s}px)`;
  // The box-space transform, on the canvas: device px from the box's origin.
  const ox = input.originX * s;
  const oy = input.originY * s;
  context.setTransform(a, b, c, d, ox + e * s - a * ox - c * oy, oy + f * s - b * ox - d * oy);
  context.drawImage(layer, 0, 0);
  context.restore();
  return true;
}

/** The words' bounds in the text box: their advances across, their lines' ascent to descent down. */
export function officeTextShadowWordBounds(glyphs: readonly TextWarpGlyph[]): [number, number, number, number] {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const glyph of glyphs) {
    left = Math.min(left, glyph.x);
    right = Math.max(right, glyph.x + glyph.advance);
    top = Math.min(top, glyph.baseline - glyph.ascent);
    bottom = Math.max(bottom, glyph.baseline + glyph.descent);
  }
  return Number.isFinite(left) ? [left, top, right, bottom] : [0, 0, 0, 0];
}

/**
 * The shadow copy's affine transform in the text box's own space (CSS px, y
 * down), as canvas `[a, b, c, d, e, f]`: `p' = A + M (p - A) + offset`, `M`
 * the shadow's scale and skew (`textWarpShadowMatrix`) and `A` its anchor
 * (`textWarpShadowAnchor`) on the words' `bounds`. The print draws the same
 * copy through its twin, `office_text_paint.text_shadow_copy_transform`
 * (cases shared: tests/fixtures/office-text-shadow-copy-cases.json).
 */
export function officeTextShadowCopyTransform(
  shadow: TextWarpShadow,
  bounds: readonly [number, number, number, number],
): [number, number, number, number, number, number] {
  const [left, top, right, bottom] = bounds;
  const [a, b, c, d] = textWarpShadowMatrix(shadow);
  const [boxX, boxY] = textWarpShadowAnchor(shadow, right - left, bottom - top);
  const ax = left + boxX;
  const ay = top + boxY;
  return [a, b, c, d, ax - a * ax - c * ay + shadow.offsetX, ay - b * ax - d * ay + shadow.offsetY];
}

/** How far past the text box the shadow of these words can reach, CSS px. */
export function officeTextShadowCopyBleed(
  shadow: TextWarpShadow,
  glyphs: readonly TextWarpGlyph[],
  width: number,
  height: number,
): number {
  const [left, top, right, bottom] = officeTextShadowWordBounds(glyphs);
  const [a, b, c, d, e, f] = officeTextShadowCopyTransform(shadow, [left, top, right, bottom]);
  let reach = 0;
  for (const [x, y] of [[left, top], [right, top], [left, bottom], [right, bottom]]) {
    const tx = a * x + c * y + e;
    const ty = b * x + d * y + f;
    reach = Math.max(reach, -tx, tx - width, -ty, ty - height);
  }
  return Math.ceil(Math.max(0, reach) + shadow.blur * 2);
}

export const TEXT_SHADOW_COPY_ATTRIBUTE = "data-office-text-shadow-copy";

export interface TextShadowCopyOptions {
  /** The element whose words cast the shadow (untouched). */
  readonly source: HTMLElement;
  /**
   * The element the canvas goes into, under its content (it is made its own
   * stacking context while the canvas is there). The canvas is laid over the
   * source's box.
   */
  readonly host: HTMLElement;
  readonly shadow: TextWarpShadow;
  /**
   * The text box the anchor is on: the source by default, or the frame the
   * source sits in when the source is shrink-wrapped to its lines (a deck's
   * show and thumbnails draw a padded frame round a one-line body).
   */
  readonly frame?: HTMLElement;
  /** Read the text box's insets from the frame's computed padding. */
  readonly insetsFromPadding?: boolean;
}

/**
 * Keep a canvas behind `source`'s words painted with their perspective
 * shadow. Repaints on resize, on font load and when the words change.
 * Returns the disposer.
 */
export function attachOfficeTextShadowCopy(options: TextShadowCopyOptions): () => void {
  const { source, host, shadow } = options;
  const document = host.ownerDocument;
  const view = document.defaultView;
  if (!view) return () => {};
  const canvas = document.createElement("canvas");
  canvas.setAttribute(TEXT_SHADOW_COPY_ATTRIBUTE, "");
  canvas.setAttribute("aria-hidden", "true");
  Object.assign(canvas.style, { position: "absolute", pointerEvents: "none", zIndex: "-1" });
  const previousIsolation = host.style.isolation;
  host.style.isolation = "isolate";
  host.prepend(canvas);
  let frame = 0;
  const paint = () => {
    frame = 0;
    const box = options.frame ?? source;
    const width = box.offsetWidth;
    const height = box.offsetHeight;
    if (!(width > 0 && height > 0)) return;
    const bounds = box.getBoundingClientRect();
    const zoom = bounds.width / width || 1;
    let insets = { left: 0, top: 0, right: 0, bottom: 0 };
    if (options.insetsFromPadding) {
      const style = view.getComputedStyle(box);
      const side = (name: string) =>
        (Number.parseFloat(style.getPropertyValue(`padding-${name}`)) || 0) +
        (Number.parseFloat(style.getPropertyValue(`border-${name}-width`)) || 0);
      insets = { left: side("left"), top: side("top"), right: side("right"), bottom: side("bottom") };
    }
    const boxWidth = width - insets.left - insets.right;
    const boxHeight = height - insets.top - insets.bottom;
    const sourceBounds = source.getBoundingClientRect();
    const glyphs = collectTextWarpGlyphs(source).map((glyph) => ({
      ...glyph,
      // Into the text box's coordinates.
      x: glyph.x + (sourceBounds.left - bounds.left) / zoom - insets.left,
      baseline: glyph.baseline + (sourceBounds.top - bounds.top) / zoom - insets.top,
    }));
    const bleed = officeTextShadowCopyBleed(shadow, glyphs, boxWidth, boxHeight);
    // Wherever the canvas's containing block is: measure where (0, 0) lands
    // and move by the difference.
    Object.assign(canvas.style, { left: "0px", top: "0px" });
    const origin = canvas.getBoundingClientRect();
    canvas.style.left = `${(bounds.left - origin.left) / zoom - bleed}px`;
    canvas.style.top = `${(bounds.top - origin.top) / zoom - bleed}px`;
    canvas.style.width = `${width + bleed * 2}px`;
    canvas.style.height = `${height + bleed * 2}px`;
    const scale = Math.min(4, Math.max(1, (view.devicePixelRatio || 1) * zoom));
    const pixelWidth = Math.ceil((width + bleed * 2) * scale);
    const pixelHeight = Math.ceil((height + bleed * 2) * scale);
    if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
    if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
    paintOfficeTextShadowCopy(canvas, {
      width: boxWidth,
      height: boxHeight,
      glyphs,
      scale,
      originX: insets.left + bleed,
      originY: insets.top + bleed,
      shadow,
    });
  };
  const schedule = () => {
    if (!frame) frame = view.requestAnimationFrame(paint);
  };
  const resize = typeof view.ResizeObserver === "function" ? new view.ResizeObserver(schedule) : null;
  resize?.observe(source);
  if (options.frame && options.frame !== source) resize?.observe(options.frame);
  const mutation = typeof view.MutationObserver === "function" ? new view.MutationObserver(schedule) : null;
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
    host.style.isolation = previousIsolation;
  };
}
