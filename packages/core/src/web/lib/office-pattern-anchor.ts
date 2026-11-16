/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where an Office pattern fill's tiles sit on screen, as PowerPoint draws
 * them in its window (pattern-presets-probe, PowerPoint for Mac at 100% and
 * 200%, Retina: every preset's period is 16 device pixels at both zooms):
 *
 * - one bit a screen point -- one CSS pixel -- whatever the zoom: the bitmap
 *   never grows or shrinks with the view, the slide or the shape;
 * - the tiles on the slide's (the page's) own grid from its top-left corner,
 *   not from the shape's: two shapes' patterns line up across them.
 *
 * Printing is not this: a PDF prints one bit a point on the page's 8 pt grid
 * (``office_chart_shape_paint.paint_office_fill_pattern(page_grid=True)``).
 *
 * One primitive for every host -- a slide on the canvas, in the show, in a
 * thumbnail, a Word page, WordArt's clipped words, a table cell -- that draws
 * its pattern either as an SVG `<pattern data-office-pattern>` or as a CSS
 * background naming `--office-pattern-size` (``officeFillPatternBackground``):
 * the host's root (the slide, the page) is the ANCHOR, and
 * {@link syncOfficePatterns} sets each pattern's scale and phase from the
 * drawn geometry, so no host has to know its zoom.
 */

/** Office's tile: 8 x 8 bits. */
export const OFFICE_PATTERN_TILE_BITS = 8;

/** The attribute an SVG `<pattern>` drawing an Office pattern carries. */
export const OFFICE_PATTERN_SVG_ATTRIBUTE = "data-office-pattern";

/** The attribute a host's root carries to anchor its patterns' grid. */
export const OFFICE_PATTERN_ANCHOR_ATTRIBUTE = "data-office-pattern-anchor";

/**
 * A CSS-background pattern's size and phase, in the element's own (layout)
 * pixels: the tile `size` that is `bitCssPx` CSS pixels a bit on screen, and
 * the `x`/`y` background position that puts a tile corner on the anchor's
 * top-left corner.
 *
 * `scale`: the element's on-screen CSS pixels per layout pixel (its zoom,
 * the slide's fit, a thumbnail's shrink). `offsetX`/`offsetY`: where the
 * element's top-left corner is from the anchor's, in on-screen CSS pixels.
 */
export function officePatternPlacement({
  scale,
  offsetX,
  offsetY,
  bitCssPx = 1,
}: {
  scale: number;
  offsetX: number;
  offsetY: number;
  bitCssPx?: number;
}): { size: number; x: number; y: number } | null {
  if (!(scale > 0) || !Number.isFinite(scale) || !(bitCssPx > 0)) return null;
  const screenTile = OFFICE_PATTERN_TILE_BITS * bitCssPx;
  const size = screenTile / scale;
  // The anchor's corner, in the element's own pixels, moved into [−size, 0].
  const phase = (offset: number) => {
    const local = -offset / scale;
    const wrapped = local - Math.ceil(local / size) * size;
    return Math.abs(wrapped) < 1e-9 ? 0 : wrapped;
  };
  return { size, x: phase(offsetX), y: phase(offsetY) };
}

/**
 * An SVG pattern's `patternTransform` (a `matrix(...)`) so that one pattern
 * unit is `bitCssPx` CSS pixels on screen and the tiles start at the anchor's
 * top-left corner, whatever the SVG's own user space (a shape's box at the
 * slide's scale, a rotated shape): the inverse of the user space's screen
 * matrix, then the anchor's origin. Null for a degenerate matrix.
 */
export function officePatternSvgTransform(
  screen: { a: number; b: number; c: number; d: number; e: number; f: number },
  anchor: { left: number; top: number },
  bitCssPx = 1,
): string | null {
  const { a, b, c, d, e, f } = screen;
  const det = a * d - b * c;
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return null;
  // inverse(screen) . translate(anchor) . scale(bitCssPx)
  const ia = d / det;
  const ib = -b / det;
  const ic = -c / det;
  const id = a / det;
  const ie = (c * f - d * e) / det;
  const iff = (b * e - a * f) / det;
  const tx = anchor.left;
  const ty = anchor.top;
  const s = bitCssPx;
  const m = [ia * s, ib * s, ic * s, id * s, ia * tx + ic * ty + ie, ib * tx + id * ty + iff];
  return `matrix(${m.map((value) => Number(value.toFixed(6))).join(" ")})`;
}

/** The CSS-background hosts' marker: their inline style names this variable. */
const CSS_PATTERN_SELECTOR = '[style*="--office-pattern-size"]';

/**
 * Lays every Office pattern under `anchor` on the anchor's grid, one bit
 * `bitCssPx` CSS pixels on screen: an SVG `<pattern data-office-pattern>`
 * by its `patternTransform`, a CSS background by the element's
 * `--office-pattern-size` / `--office-pattern-x` / `--office-pattern-y`.
 * A turned or sheared HTML element (its box not its layout box scaled
 * alike both ways) keeps its own defaults.
 */
export function syncOfficePatterns(anchor: Element, bitCssPx = 1, originSelector?: string): void {
  const svgPatterns = Array.from(anchor.querySelectorAll(`pattern[${OFFICE_PATTERN_SVG_ATTRIBUTE}]`));
  const cssPatterns = Array.from(anchor.querySelectorAll(CSS_PATTERN_SELECTOR));
  // No pattern: no measuring (a frame's layout is never forced for nothing,
  // the slide being typed into).
  if (!svgPatterns.length && !cssPatterns.length) return;
  const rootBox = anchor.getBoundingClientRect();
  // A paged host's patterns sit on their own page's grid: the page (an
  // `originSelector` match under the anchor) whose box holds the pattern's
  // corner, else the anchor's.
  const origins = originSelector ? Array.from(anchor.querySelectorAll(originSelector)).map((page) => page.getBoundingClientRect()) : [];
  const originOf = (left: number, top: number) =>
    origins.find((page) => left >= page.left - 0.5 && left < page.right && top >= page.top - 0.5 && top < page.bottom) ?? rootBox;
  for (const pattern of svgPatterns) {
    const owner = (pattern as SVGPatternElement).ownerSVGElement;
    const screen = owner?.getScreenCTM?.();
    if (!screen || !owner) continue;
    const ownerBox = owner.getBoundingClientRect();
    const anchorBox = originOf(ownerBox.left, ownerBox.top);
    const transform = officePatternSvgTransform(screen, { left: anchorBox.left, top: anchorBox.top }, bitCssPx);
    if (transform && pattern.getAttribute("patternTransform") !== transform) pattern.setAttribute("patternTransform", transform);
  }
  for (const element of cssPatterns) {
    const html = element as HTMLElement;
    const box = html.getBoundingClientRect();
    const anchorBox = originOf(box.left, box.top);
    const width = html.offsetWidth;
    const height = html.offsetHeight;
    const set = (name: string, value: string) => {
      if (html.style.getPropertyValue(name) !== value) html.style.setProperty(name, value);
    };
    const scaleX = width > 0 ? box.width / width : 0;
    const scaleY = height > 0 ? box.height / height : 0;
    const placement =
      scaleX > 0 && Math.abs(scaleX - scaleY) <= 0.02 * scaleX
        ? officePatternPlacement({ scale: scaleX, offsetX: box.left - anchorBox.left, offsetY: box.top - anchorBox.top, bitCssPx })
        : null;
    if (!placement) {
      // Its own tile from its own corner (an inherited placement would be
      // another element's): a turned box, an inline run's broken box.
      set("--office-pattern-size", `${OFFICE_PATTERN_TILE_BITS}px`);
      set("--office-pattern-x", "0px");
      set("--office-pattern-y", "0px");
      continue;
    }
    set("--office-pattern-size", `${Number(placement.size.toFixed(4))}px`);
    set("--office-pattern-x", `${Number(placement.x.toFixed(4))}px`);
    set("--office-pattern-y", `${Number(placement.y.toFixed(4))}px`);
  }
}

/**
 * Keeps `anchor`'s patterns on its grid: now, after every change under it
 * (a shape added, moved, re-sized, its fill changed) and whenever its drawn
 * size changes, once a frame at most. The caller re-runs it when the zoom
 * changes outside the anchor (`refresh`). Returns the disconnect.
 */
export function watchOfficePatterns(
  anchor: HTMLElement,
  bitCssPx = 1,
  originSelector?: string,
): { refresh(): void; disconnect(): void } {
  let frame = 0;
  const run = () => {
    frame = 0;
    syncOfficePatterns(anchor, bitCssPx, originSelector);
  };
  const schedule = () => {
    if (frame || typeof requestAnimationFrame === "undefined") return;
    frame = requestAnimationFrame(run);
  };
  // Our own writes (patternTransform, the three variables) never re-schedule.
  const mutations =
    typeof MutationObserver === "undefined"
      ? null
      : new MutationObserver((records) => {
          for (const record of records) {
            if (record.type !== "attributes") return schedule();
            if (record.attributeName === "patternTransform") continue;
            if (record.attributeName === "style" && record.target !== anchor) {
              const style = (record.target as HTMLElement).style;
              const old = record.oldValue ?? "";
              // Only our variables changed: the style text minus them is the same.
              const strip = (text: string) => text.replace(/--office-pattern-[xy]?(?:size)?:[^;]*;?\s*/g, "");
              if (strip(old) === strip(style?.cssText ?? "")) continue;
            }
            return schedule();
          }
        });
  mutations?.observe(anchor, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeOldValue: true,
    attributeFilter: ["style", "transform", "width", "height", "x", "y", "viewBox", "class"],
  });
  const resize = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(schedule);
  resize?.observe(anchor);
  run();
  return {
    refresh: schedule,
    disconnect() {
      mutations?.disconnect();
      resize?.disconnect();
      if (frame && typeof cancelAnimationFrame !== "undefined") cancelAnimationFrame(frame);
      frame = 0;
    },
  };
}
