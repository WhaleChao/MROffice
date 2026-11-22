/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The stroke under the pen, drawn the moment the pen moves.
 *
 * Every ink surface previewed the running stroke by putting its samples in
 * React state on each pointer move and letting React rebuild one SVG path
 * of the whole stroke: a render of the surface (on the worksheet and the
 * deck, of the whole window) per sample, the path data of every sample
 * recomputed, and the stroke's pixels a frame or two behind the pen tip.
 * Office and every note app draw the running stroke on a canvas instead,
 * appending only the newest piece, and draw a little AHEAD of the tip from
 * the pointer's own prediction, so the ink seems to come out of the pen.
 *
 * This is that canvas, framework-free: a host mounts it over its drawing
 * surface, calls {@link InkLiveCanvas.begin} at pointerdown with the box the
 * stroke is confined to, pushes every sample (and the browser's predicted
 * ones) from pointermove, and clears it once the stroke is on screen by
 * other means (the wet-ink SVG, the document's own node).
 *
 * - Two canvases: the INK canvas takes a segment once its curve is settled
 *   (a Catmull-Rom segment needs the sample after its end), never redrawn;
 *   the TAIL canvas is cleared and redrawn each move with the last settled
 *   point to the tip and on through the predicted samples.
 * - The eraser's and the lasso's trail (`trace`) is no ink: it is drawn
 *   straight through the real samples only, where the eraser cuts.
 * - The pen's width follows the pressure when the host asks for it; the
 *   curve and the corner rule are the smoother's own (lib/office-ink-smooth),
 *   so the running stroke matches the stroke once it lands.
 * - A translucent tool (the highlighter) is drawn opaque and the CANVAS is
 *   made translucent, so a stroke crossing itself does not darken where it
 *   overlaps, as a single SVG path does not.
 * - Ordinary composition, never `desynchronized`: a low-latency canvas
 *   skips the compositor, and a TRANSPARENT one laid over the page is where
 *   that path breaks: on a Windows Chrome/Edge the whole overlay showed
 *   black while the pen was down (the slide reappeared on release), and
 *   Chromium has the same alpha + desynchronized blackout on record. The
 *   frame it saves is not worth a page the writer cannot see.
 * - The overlay is shown only once both canvases have a live 2-D context;
 *   a host checks {@link InkLiveCanvas.usable} and previews the stroke its
 *   own way when they do not, never through a page-sized layer that cannot
 *   draw.
 * - Sized to the stroke's box at the device's pixel ratio and the host's
 *   zoom, capped so a whole page at 400 % stays within a texture; the
 *   canvas is never larger than what one stroke can reach.
 */

import {
  inkCurveControls,
  inkPressuresVary,
  inkPressureWidthFactor,
  inkSteadyPressureWidthFactor,
  inkStrokePathData,
  inkTurnIsCorner,
  inkVariableWidthOutlinePathData,
  INK_CORNER_DEGREES,
  type InkPressurePoint,
} from "./office-ink-smooth";
import { documentInkPencilFilter, documentInkPencilRegionFor, type DocumentInkFilter } from "./document-ink";
import { pencilScribblePossible } from "./ink-scribble-guard";
import {
  InkSparkleWalker,
  inkEffectCanvasGradient,
  inkEffectGradient,
  inkEffectPaint,
  isInkEffect,
  inkEffectReach,
  paintInkEffectOnCanvas,
  type InkEffectPaint,
} from "./office-ink-effects";

export interface InkLiveRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

export interface InkLiveStyle {
  readonly color: string;
  /** The pen's width in the host's own pixels (before zoom). */
  readonly widthPx: number;
  /** The tool's opacity (the highlighter's wash); applied to the canvas. */
  readonly opacity: number;
  /** Multiply over the page, as a highlighter does. */
  readonly multiply?: boolean;
  /** Lighten the page instead, as a highlighter does on a dark page. */
  readonly screen?: boolean;
  /** Vary the width with each sample's pressure. */
  readonly pressure?: boolean;
  /** The corner rule (degrees); default the smoother's. */
  readonly corner?: number;
  /** A dashed line (an eraser's trail), in host pixels. */
  readonly dash?: readonly number[];
  /**
   * A trace of where the pointer has been (the eraser's reach, the lasso's
   * loop) rather than ink: drawn straight from sample to sample, as the
   * eraser's cut and the lasso's polygon measure it, and never ahead of the
   * real pointer. A predicted front is taken back by the next real sample;
   * on ink that hides a frame of lag, on the eraser it is a front that jumps
   * ahead of what is erased and snaps back with every move.
   */
  readonly trace?: boolean;
  /**
   * A pen effect's colours: the stroke wears them as a gradient from its
   * box's top-left corner to its bottom-right one, as the landed stroke's
   * SVG paints it (`objectBoundingBox`, 0,0 to 1,1). Without it the running
   * stroke showed the effect's one solid stand-in and changed look on the
   * lift.
   */
  readonly gradient?: readonly string[];
  /**
   * The pen's effect whole (lib/office-ink-effects): the stroke is drawn in
   * its layers -- a neon's halo and core, a metal's sheen, a shimmer's
   * glints -- and its base in the very gradient the landed stroke wears
   * (`gradient` is then only the tail's tint).
   */
  readonly effect?: InkEffectPaint;
  







  readonly pencil?: boolean;
}

/**
 * The part of a live style a pen's effect makes (none for a plain pen, a
 * highlighter or a pencil): its layers and its gradient.
 */
export function inkLiveEffectStyle(pen: {
  readonly tool: string;
  readonly effect?: string | null;
}): Pick<InkLiveStyle, "gradient" | "effect"> {
  if (pen.tool !== "pen" || !pen.effect || !isInkEffect(pen.effect)) return {};
  const effect = inkEffectPaint(pen.effect);
  return { gradient: effect.stops, effect };
}

/**
 * Where a floating canvas sits on screen: the client position of the host's
 * origin and how many screen pixels one host pixel is.
 */
export interface InkLiveScreen {
  readonly left: number;
  readonly top: number;
  readonly scale: number;
}

const FLOAT_LAYER_ATTRIBUTE = "data-ink-live-float-layer";

/**
 * The one layer floating canvases live in: fixed over the viewport, over
 * everything, taking no pointer, outside every page's transforms.
 */
export function inkFloatLayer(ownerDocument: Document): HTMLElement {
  const existing = ownerDocument.querySelector<HTMLElement>(`[${FLOAT_LAYER_ATTRIBUTE}]`);
  if (existing) return existing;
  const layer = ownerDocument.createElement("div");
  layer.setAttribute(FLOAT_LAYER_ATTRIBUTE, "");
  layer.setAttribute("aria-hidden", "true");
  Object.assign(layer.style, {
    position: "fixed",
    left: "0px",
    top: "0px",
    width: "0px",
    height: "0px",
    overflow: "visible",
    pointerEvents: "none",
    zIndex: "2147483000",
  } satisfies Partial<CSSStyleDeclaration>);
  ownerDocument.body.append(layer);
  return layer;
}

/** A running stroke up to this many samples is redrawn whole each move ({@link InkLiveCanvas.extend}). */
const WHOLE_PATH_MAX_POINTS = 1500;

type InkBox = { minX: number; minY: number; maxX: number; maxY: number };

/** The bounds of `points` (host px), or null for none. */
function boxOf(points: readonly { x: number; y: number }[]): InkBox | null {
  let box: InkBox | null = null;
  for (const point of points) {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
    if (!box) box = { minX: point.x, minY: point.y, maxX: point.x, maxY: point.y };
    else {
      box.minX = Math.min(box.minX, point.x);
      box.minY = Math.min(box.minY, point.y);
      box.maxX = Math.max(box.maxX, point.x);
      box.maxY = Math.max(box.maxY, point.y);
    }
  }
  return box;
}

function unionBox(a: InkBox, b: InkBox | null): InkBox {
  if (!b) return a;
  return {
    minX: Math.min(a.minX, b.minX),
    minY: Math.min(a.minY, b.minY),
    maxX: Math.max(a.maxX, b.maxX),
    maxY: Math.max(a.maxY, b.maxY),
  };
}

/** The largest backing texture a side may take, in device pixels. */
const MAX_BACKING_PX = 4096;
/** Predicted samples further than this from the tip (host px) are not drawn: a
 * wild guess past a corner reads as a wobble when the next frame takes it back. */
const MAX_PREDICTION_PX = 14;
const SVG_NS = "http://www.w3.org/2000/svg";
/** The live pencil's grain seed: the wet layer's (DocumentInkSurface) too. */
export const INK_LIVE_PENCIL_SEED = 7;
const PENCIL_LIVE_SEED = INK_LIVE_PENCIL_SEED;
let pencilFilterIds = 0;
/** Samples in one piece of a pencil stroke being drawn (see `InkLiveCanvas.drawPencil`). */
const PENCIL_PIECE_POINTS = 8;











const PEN_PIECE_POINTS = 32;
const PEN_PIECE_OVERLAP = 2;













const PENCIL_WHOLE_MAX_AREA_PX = 220 * 220;
/**
 * How far the texture filter looks around a pixel (its blur's three sigmas
 * and its reach, at most 2.2 and 2 px), beyond the stroke's half width: a
 * piece drawn this far past its ends is, within them, the whole stroke.
 */
const PENCIL_FILTER_REACH_PX = 10;

/**
 * Paints the stroke `run` (host pixels) into `box` as the page paints a
 * landed stroke, placing `box` at the stroke's canvas in host pixels; returns
 * the SVG it drew and where that SVG's user space starts, in host pixels.
 */
export type InkLivePainter = (
  box: HTMLDivElement,
  run: readonly InkPressurePoint[],
) => { readonly svg: SVGSVGElement; readonly originX: number; readonly originY: number } | null;

interface PaintedPiece {
  readonly box: HTMLDivElement;
  readonly cutId: string;
}

interface PencilPiece {
  readonly group: SVGSVGElement;
  readonly filter: SVGFilterElement;
  readonly cut: SVGPathElement;
  readonly path: SVGPathElement;
}

/** A line across the stroke at a piece boundary: through (x, y), (dx, dy) along the stroke. */
interface PencilCut {
  readonly x: number;
  readonly y: number;
  readonly dx: number;
  readonly dy: number;
}

/** The half disc of `radius` about the cut's point on its side `side` (1: ahead along the stroke, -1: behind). */
function halfDisc(cut: PencilCut, radius: number, side: 1 | -1): string {
  const steps = 16;
  const along = Math.atan2(cut.dy * side, cut.dx * side);
  const points: string[] = [];
  for (let step = 0; step <= steps; step += 1) {
    const angle = along - Math.PI / 2 + (Math.PI * step) / steps;
    points.push(`${cut.x + radius * Math.cos(angle)} ${cut.y + radius * Math.sin(angle)}`);
  }
  return `M ${points.join(" L ")} Z`;
}

export class InkLiveCanvas {
  /** The wrapper the host mounts: absolutely positioned, takes no pointer. */
  readonly element: HTMLDivElement;
  private readonly ink: HTMLCanvasElement;
  private readonly tail: HTMLCanvasElement;
  private inkContext: CanvasRenderingContext2D | null = null;
  private tailContext: CanvasRenderingContext2D | null = null;
  private points: InkPressurePoint[] = [];
  /** Segments `0 .. settled-1` are on the ink canvas. */
  private settled = 0;
  private style: InkLiveStyle | null = null;
  private rect: InkLiveRect = { left: 0, top: 0, width: 0, height: 0 };
  /** Device pixels per host pixel on the backing store. */
  private density = 1;
  private readonly layer: HTMLDivElement;
  /** The pencil texture on the canvases: the landed stroke's own filter (see `pencilTexture`). */
  private pencilBox: HTMLDivElement | null = null;
  private pencilFilterSpec: DocumentInkFilter | null = null;
  /** The piece being drawn, and the sample its own part starts at (see `drawPencil`). */
  private pencilLive: PencilPiece | null = null;
  private pencilStart = 0;
  /** Screen pixels per host pixel, and where the rect's corner is from the element's (snapped) corner. */
  private screenScale = 1;
  private cornerOffset = { x: 0, y: 0 };
  private supported = true;
  /** The samples' box (host px), for a gradient that follows the stroke's box. */
  private box: { minX: number; minY: number; maxX: number; maxY: number } | null = null;
  






  private tailDrawn: { minX: number; minY: number; maxX: number; maxY: number } | null = null;
  /** Both canvases hold nothing (cleared since the last stroke drew). */
  private canvasesClean = false;
  







  private readonly drawOncePerFrame = pencilScribblePossible();
  private drawFrame: number | null = null;
  private drawPredicted: readonly InkPressurePoint[] = [];
  /**
   * A long canvas stroke in pieces (see `drawPieces`): samples before
   * `piecedStart` are on the ink canvas in finished pieces, drawn once; the
   * rest is the live piece on the tail canvas. `piecedOutline` is the look
   * the pieces were drawn in (a pressure outline or a line); null, none is.
   */
  private piecedStart = 0;
  private piecedOutline: boolean | null = null;
  /** A shimmer's glints so far, and how many samples they were found in. */
  private sparkles: InkSparkleWalker | null = null;
  private sparkleSamples = 0;

  








  private readonly floating: boolean;

  /** The element whose pixels the strokes are in (a floating canvas's parent is the float layer). */
  readonly host: HTMLElement | null;

  /**
   * Where a pencil stroke's SVG goes when the page offers a place for it: an
   * element inside the page's zoom whose pixels are the samples' (the ink
   * surface). See `pencilTexture`.
   */
  private readonly pageHost: HTMLElement | null;
  







  private painter: InkLivePainter | null = null;
  private readonly paintedPieces: PaintedPiece[] = [];
  private paintedLive: PaintedPiece | null = null;
  private paintedStart = 0;
  /** The page-zoomed boxes the painted strokes go in: clipped to the rect, and inside it. */
  private paintClip: HTMLDivElement | null = null;
  private paintBox: HTMLDivElement | null = null;

  constructor(
    ownerDocument: Document = document,
    options: { floating?: boolean; host?: HTMLElement | null; pageHost?: HTMLElement | null } = {},
  ) {
    this.floating = options.floating === true;
    this.host = options.host ?? null;
    this.pageHost = options.pageHost ?? null;
    this.element = ownerDocument.createElement("div");
    this.element.setAttribute("data-testid", "ink-live-canvas");
    this.element.setAttribute("aria-hidden", "true");
    Object.assign(this.element.style, {
      position: "absolute",
      left: "0px",
      top: "0px",
      width: "0px",
      height: "0px",
      pointerEvents: "none",
      overflow: "hidden",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    // The canvases' layer: in a pencil stroke it is laid out in host pixels
    // and zoomed like the page, so the texture filter runs where the landed
    // stroke's does (see `writePlace`).
    this.layer = ownerDocument.createElement("div");
    Object.assign(this.layer.style, {
      position: "absolute",
      left: "0px",
      top: "0px",
      transformOrigin: "0 0",
    } satisfies Partial<CSSStyleDeclaration>);
    this.element.append(this.layer);
    this.ink = ownerDocument.createElement("canvas");
    this.tail = ownerDocument.createElement("canvas");
    for (const canvas of [this.ink, this.tail]) {
      Object.assign(canvas.style, {
        position: "absolute",
        left: "0px",
        top: "0px",
      } satisfies Partial<CSSStyleDeclaration>);
      this.layer.append(canvas);
    }
    this.ink.setAttribute("data-ink-live", "ink");
    this.tail.setAttribute("data-ink-live", "tail");
    if (this.floating) inkFloatLayer(ownerDocument).append(this.element);
  }

  /** Whether the running stroke is on these canvases. */
  get active(): boolean {
    return this.style !== null;
  }

  /** The samples pushed so far (the host keeps its own copy; this is for tests). */
  get sampleCount(): number {
    return this.points.length;
  }

  private context(canvas: HTMLCanvasElement): CanvasRenderingContext2D | null {
    if (!this.supported) return null;
    try {
      // A canvas keeps the options of its FIRST context: these are the only
      // ones this canvas is ever asked for.
      const context = canvas.getContext("2d", { alpha: true });
      if (!context) this.supported = false;
      return context;
    } catch {
      this.supported = false;
      return null;
    }
  }

  /**
   * Whether the running stroke can be drawn here: both canvases give a 2-D
   * context that is not lost. A host asks before choosing this canvas over
   * its own preview.
   */
  get usable(): boolean {
    this.inkContext ??= this.context(this.ink);
    this.tailContext ??= this.context(this.tail);
    return [this.inkContext, this.tailContext].every(
      (context) => context !== null && !(context.isContextLost?.() ?? false),
    );
  }

  /**
   * Start a stroke confined to `rect` (host pixels, from the mounted
   * element's origin), drawn in `style`; `zoom` is how many screen pixels one
   * host pixel is (the document zoom), so the backing store is crisp.
   * False (and nothing shown) when the canvases cannot draw.
   */
  begin(rect: InkLiveRect, style: InkLiveStyle, zoom = 1, screen?: InkLiveScreen, paint?: InkLivePainter): boolean {
    this.clear();
    if (!this.usable) return false;
    this.style = style;
    this.painter = paint && this.pageHost && !style.trace ? paint : null;
    this.sparkles = style.effect?.sparkle?.length ? new InkSparkleWalker(style.widthPx) : null;
    this.sparkleSamples = 0;
    this.rect = rect;
    const view = this.element.ownerDocument.defaultView;
    const dpr = view && Number.isFinite(view.devicePixelRatio) && view.devicePixelRatio > 0 ? view.devicePixelRatio : 1;
    const ideal = Math.max(0.25, zoom) * dpr;
    const cap = Math.min(
      1,
      MAX_BACKING_PX / Math.max(1, rect.width * ideal),
      MAX_BACKING_PX / Math.max(1, rect.height * ideal),
    );
    this.density = ideal * cap;
    // The backing store: the rect, and two pixels over for the fraction the
    // device-pixel-snapped corner moves the drawing in by (see writePlace).
    const backingWidth = Math.max(1, Math.ceil(rect.width * this.density) + 2);
    const backingHeight = Math.max(1, Math.ceil(rect.height * this.density) + 2);
    let resized = false;
    for (const canvas of [this.ink, this.tail]) {
      if (canvas.width !== backingWidth) {
        canvas.width = backingWidth;
        resized = true;
      }
      if (canvas.height !== backingHeight) {
        canvas.height = backingHeight;
        resized = true;
      }
    }
    // Floating, the canvas covers the same rect on screen: host pixels
    // placed by `screen` (the backing store already counts the zoom).
    const place = this.floating && screen ? screen : { left: 0, top: 0, scale: 1 };
    this.placed = null;
    this.writePlace(place);
    if (this.painter) this.paintedBegin();
    else this.pencilTexture(style);
    Object.assign(this.element.style, {
      opacity: String(Math.min(1, Math.max(0, style.opacity))),
      mixBlendMode: style.multiply ? "multiply" : style.screen ? "screen" : "normal",
    });
    this.inkContext = this.context(this.ink);
    this.tailContext = this.context(this.tail);
    // A resized canvas is blank, and one the last lift cleared is too: the
    // full clear (a page-sized backing store) only when neither holds.
    const fullClear = !resized && !this.canvasesClean;
    this.canvasesClean = false;
    for (const context of [this.inkContext, this.tailContext]) {
      if (!context) continue;
      context.setTransform(1, 0, 0, 1, 0, 0);
      if (fullClear) context.clearRect(0, 0, backingWidth, backingHeight);
      context.setTransform(
        this.density,
        0,
        0,
        this.density,
        -rect.left * this.density + this.subpixel.x,
        -rect.top * this.density + this.subpixel.y,
      );
      // The whole line style, every stroke: a context keeps its state while
      // the canvas keeps its size, so whatever is not set here is the last
      // tool's (a pen after the stroke eraser drew dashed).
      context.globalAlpha = 1;
      context.globalCompositeOperation = "source-over";
      context.lineCap = "round";
      context.lineJoin = "round";
      context.strokeStyle = style.color;
      context.fillStyle = style.color;
      context.setLineDash(style.dash && style.dash.length > 0 ? [...style.dash] : []);
      context.lineDashOffset = 0;
    }
    // Shown only now: cleared, sized and styled.
    this.element.style.display = "block";
    return true;
  }

  /** Where a floating canvas was last put on screen. */
  private placed: InkLiveScreen | null = null;

  /**
   * What the element's device-pixel-snapped corner leaves over, in backing
   * pixels: the strokes are drawn that much further in, so they sit at the
   * exact sub-pixel position the page draws them at once landed.
   */
  private subpixel = { x: 0, y: 0 };

  private writePlace(place: InkLiveScreen): void {
    const last = this.placed;
    if (last && last.left === place.left && last.top === place.top && last.scale === place.scale) return;
    this.placed = place;
    const rect = this.rect;
    const view = this.element.ownerDocument.defaultView;
    const dpr = view && Number.isFinite(view.devicePixelRatio) && view.devicePixelRatio > 0 ? view.devicePixelRatio : 1;
    // A canvas is a layer of its own, and WebKit puts a layer on whole
    // device pixels: a corner left at a fraction moved every running stroke
    // by up to half a device pixel from where it landed (plain to see at
    // 200 % on a Retina iPad). The corner goes on a device pixel and the
    // fraction into the drawing.
    const left = place.left + rect.left * place.scale;
    const top = place.top + rect.top * place.scale;
    const snappedLeft = Math.floor(left * dpr) / dpr;
    const snappedTop = Math.floor(top * dpr) / dpr;
    const ratio = dpr > 0 ? this.density / Math.max(1e-6, place.scale) / dpr : 1;
    const subpixel = { x: (left - snappedLeft) * dpr * ratio, y: (top - snappedTop) * dpr * ratio };
    this.screenScale = place.scale;
    this.cornerOffset = { x: left - snappedLeft, y: top - snappedTop };
    this.placePencil();
    this.placePainted();
    Object.assign(this.element.style, {
      left: `${snappedLeft}px`,
      top: `${snappedTop}px`,
      width: `${Math.max(0, Math.ceil((left + rect.width * place.scale) * dpr) / dpr - snappedLeft)}px`,
      height: `${Math.max(0, Math.ceil((top + rect.height * place.scale) * dpr) / dpr - snappedTop)}px`,
    });
    // The canvases at exactly their backing store's size on screen, from the
    // snapped corner: the snap only widens the box, its new edge transparent
    // (the element clips at it). Stretched to fill the widened box, a running
    // stroke drew up to 1.3px further out than it landed, the more the
    // further from the corner.
    // A pencil's layer is in host pixels, zoomed by the page's zoom: its
    // texture filter then works in the very units the landed stroke's does
    // and is zoomed with it (Safari draws a filter zoomed with its element
    // a shade apart from one run at screen size). Any other tool's layer is
    // in screen pixels, as it always was.
    const layerScale = this.style?.pencil && !this.pageHost ? place.scale : 1;
    this.layer.style.transform = layerScale !== 1 ? `scale(${layerScale})` : "";
    const backingPerLayerPx = (this.density / Math.max(1e-6, place.scale)) * layerScale;
    for (const canvas of [this.ink, this.tail]) {
      canvas.style.width = `${canvas.width / backingPerLayerPx}px`;
      canvas.style.height = `${canvas.height / backingPerLayerPx}px`;
    }
    if (subpixel.x !== this.subpixel.x || subpixel.y !== this.subpixel.y) {
      this.subpixel = subpixel;
      this.applyTransform();
    }
  }

  /** Host pixels onto the backing store: the rect's corner at the snapped corner, plus the fraction. */
  private applyTransform(): void {
    const rect = this.rect;
    const density = this.density;
    for (const context of [this.inkContext, this.tailContext]) {
      context?.setTransform(
        density,
        0,
        0,
        density,
        -rect.left * density + this.subpixel.x,
        -rect.top * density + this.subpixel.y,
      );
    }
  }

  /**
   * A floating canvas follows its host on screen: the host's pixels moved
   * under the pen (the page scrolled, the ribbon folded into the pen
   * palette, a batch of ink landed above) and the samples are read against
   * where the host is now. Left where the stroke began, the running stroke
   * drew away from the pen and jumped back to it on the lift.
   */
  place(screen: InkLiveScreen): void {
    if (!this.floating || !this.style) return;
    this.writePlace(screen);
  }

  /** Whether the pen turned by more than the corner rule at sample `index`. */
  private isCorner(index: number): boolean {
    const points = this.points;
    if (index <= 0 || index >= points.length - 1) return false;
    return inkTurnIsCorner(
      points[index - 1]!,
      points[index]!,
      points[index + 1]!,
      this.style?.corner ?? INK_CORNER_DEGREES,
    );
  }

  private widthAt(a: InkPressurePoint, b: InkPressurePoint): number {
    const style = this.style;
    if (!style) return 1;
    if (!style.pressure) return style.widthPx;
    const pa = a.p;
    const pb = b.p;
    const pressure =
      pa === undefined ? pb : pb === undefined ? pa : (pa + pb) / 2;
    return style.widthPx * inkPressureWidthFactor(pressure);
  }

  private segment(
    context: CanvasRenderingContext2D,
    p0: InkPressurePoint | null,
    p1: InkPressurePoint,
    p2: InkPressurePoint,
    p3: InkPressurePoint | null,
  ): void {
    context.lineWidth = Math.max(0.5, this.widthAt(p1, p2));
    context.beginPath();
    context.moveTo(p1.x, p1.y);
    if (this.style?.trace) {
      // The eraser cuts along the straight path between its samples.
      context.lineTo(p2.x, p2.y);
    } else {
      const { c1x, c1y, c2x, c2y } = inkCurveControls(p0, p1, p2, p3);
      context.bezierCurveTo(c1x, c1y, c2x, c2y, p2.x, p2.y);
    }
    context.stroke();
  }

  private dot(context: CanvasRenderingContext2D, at: InkPressurePoint): void {
    const width = Math.max(0.5, this.widthAt(at, at));
    context.beginPath();
    context.arc(at.x, at.y, width / 2, 0, Math.PI * 2);
    context.fill();
  }

  /**
   * Take the samples a pointer move brought (host pixels) and redraw the
   * tail through `predicted` (the browser's guess at the next samples; a
   * trace ignores it). Samples on top of the previous one are ignored.
   */
  extend(samples: readonly InkPressurePoint[], predicted: readonly InkPressurePoint[] = []): void {
    if (!this.style) return;
    for (const sample of samples) {
      if (!Number.isFinite(sample.x) || !Number.isFinite(sample.y)) continue;
      const previous = this.points[this.points.length - 1];
      if (previous && previous.x === sample.x && previous.y === sample.y) continue;
      this.points.push(sample);
      const box = this.box;
      if (!box) this.box = { minX: sample.x, minY: sample.y, maxX: sample.x, maxY: sample.y };
      else {
        box.minX = Math.min(box.minX, sample.x);
        box.minY = Math.min(box.minY, sample.y);
        box.maxX = Math.max(box.maxX, sample.x);
        box.maxY = Math.max(box.maxY, sample.y);
      }
    }
    const view = this.element.ownerDocument.defaultView;
    if (this.drawOncePerFrame && view && typeof view.requestAnimationFrame === "function") {
      this.drawPredicted = predicted;
      if (this.drawFrame === null) {
        this.drawFrame = view.requestAnimationFrame(() => {
          this.drawFrame = null;
          if (this.style) this.draw(this.drawPredicted);
        });
      }
      return;
    }
    this.draw(predicted);
  }

  /** Draw the stroke as its samples stand, on through `predicted`. */
  private draw(predicted: readonly InkPressurePoint[]): void {
    if (!this.style) return;
    if (this.painter && this.paintBox && this.paintClip?.style.display !== "none") {
      this.drawPainted(predicted);
      return;
    }
    if (this.style.pencil && this.pencilBox) {
      this.drawPencil(predicted);
      return;
    }
    const points = this.points;
    const ink = this.inkContext;
    // The stroke as it will land, drawn whole: the very path the landed
    // stroke's SVG draws (its pressure outline, or its curve at its width),
    // not segment by segment -- segments stroked one by one overlap their
    // round ends and soft edges at every joint, and a tight curve drawn
    // that way thinned on the lift. A stroke past WHOLE_PATH_MAX_POINTS goes
    // on segment by segment from where it is (a frame's work stays bounded).
    const pieces = ink && !this.style.trace && typeof Path2D !== "undefined" ? this.pieceLook(points) : null;
    if (ink && pieces !== null) {
      // Long enough to go in pieces: the finished ones stay on the ink
      // canvas, the last one is drawn on the tail with the prediction.
      const live = this.drawPieces(ink, pieces);
      this.settled = Math.max(0, points.length - 1);
      this.drawTail(predicted, live, pieces);
      return;
    }
    if (ink && this.piecedOutline !== null) {
      // The pieces' look no longer holds (the pressure began to vary): the
      // stroke is drawn whole again below, the pieces cleared with it, or
      // past WHOLE_PATH_MAX_POINTS it goes on segment by segment from the
      // live piece, put down first.
      if (points.length > WHOLE_PATH_MAX_POINTS) {
        this.drawRun(ink, points.slice(Math.max(0, this.piecedStart - PEN_PIECE_OVERLAP)), this.piecedOutline);
        this.settled = Math.max(0, points.length - 1);
      }
      this.piecedOutline = null;
      this.piecedStart = 0;
    }
    if (ink && !this.style.trace && points.length <= WHOLE_PATH_MAX_POINTS && typeof Path2D !== "undefined") {
      this.drawWhole(ink, points);
      this.settled = Math.max(0, points.length - 1);
      if (!this.style.effect) this.tint(ink);
    } else if (ink) {
      // Settle every segment whose curve is known: segment i needs sample i+2
      // (the controls past its end) and the corner verdicts at i and i+1.
      while (this.settled + 2 < points.length) {
        const index = this.settled;
        const p0 = index > 0 && !this.isCorner(index) ? points[index - 1]! : null;
        const p3 = !this.isCorner(index + 1) ? points[index + 2]! : null;
        this.segment(ink, p0, points[index]!, points[index + 1]!, p3);
        this.settled += 1;
      }
      this.tint(ink);
    }
    this.drawTail(this.style.trace ? [] : predicted);
    if (this.tailContext) this.tint(this.tailContext);
  }

  /**
   * Repaint what is drawn on `context` in the effect's gradient over the
   * stroke's box as it stands (the box grows as the pen goes, so the whole
   * stroke is repainted, not the new segment alone): what the canvas shows
   * at the lift is what the landed stroke shows. Only where there is ink:
   * `source-in` keeps each pixel's coverage and takes the gradient's colour.
   */
  private tint(context: CanvasRenderingContext2D): void {
    const stops = this.style?.gradient;
    const box = this.box;
    if (!stops || stops.length === 0 || !box) return;
    const effect = this.style?.effect;
    if (effect) {
      // The landed stroke's own gradient, in host pixels (the context's units).
      const reachPx = (this.style?.widthPx ?? 1) * inkEffectReach(effect) + MAX_PREDICTION_PX + 2;
      context.save();
      context.globalCompositeOperation = "source-in";
      context.globalAlpha = 1;
      context.fillStyle = inkEffectCanvasGradient(
        context,
        inkEffectGradient(
          effect,
          { x: box.minX, y: box.minY, width: box.maxX - box.minX, height: box.maxY - box.minY },
          this.style?.widthPx ?? 1,
          this.points[0] ?? null,
        ),
      );
      context.fillRect(box.minX - reachPx, box.minY - reachPx, box.maxX - box.minX + 2 * reachPx, box.maxY - box.minY + 2 * reachPx);
      context.restore();
      return;
    }
    const width = Math.max(1e-3, box.maxX - box.minX);
    const height = Math.max(1e-3, box.maxY - box.minY);
    // The ink reaches half a width past the samples (and the prediction a
    // little further); the gradient pads with its end colours out there.
    const reach = (this.style?.widthPx ?? 1) + MAX_PREDICTION_PX + 2;
    const density = this.density;
    context.save();
    // The unit square onto the box, in device pixels: the gradient's vector
    // is then the box's own diagonal, as objectBoundingBox maps it.
    context.setTransform(
      density * width,
      0,
      0,
      density * height,
      (box.minX - this.rect.left) * density + this.subpixel.x,
      (box.minY - this.rect.top) * density + this.subpixel.y,
    );
    const gradient = context.createLinearGradient(0, 0, 1, 1);
    stops.forEach((stop, index) => gradient.addColorStop(index / Math.max(1, stops.length - 1), stop));
    context.globalCompositeOperation = "source-in";
    context.fillStyle = gradient;
    context.fillRect(-reach / width, -reach / height, 1 + (2 * reach) / width, 1 + (2 * reach) / height);
    context.restore();
  }

  




















  private pencilTexture(style: InkLiveStyle): void {
    if (!style.pencil) {
      if (this.pencilBox) this.pencilBox.style.display = "none";
      return;
    }
    const doc = this.element.ownerDocument;
    if (!this.pencilBox) {
      const box = doc.createElement("div");
      box.setAttribute("aria-hidden", "true");
      Object.assign(box.style, {
        position: "absolute",
        left: "0px",
        top: "0px",
        width: "0px",
        height: "0px",
        overflow: "visible",
        pointerEvents: "none",
      } satisfies Partial<CSSStyleDeclaration>);
      (this.pageHost ?? this.layer).append(box);
      this.pencilBox = box;
    }
    this.pencilBox.replaceChildren();
    this.pencilLive = null;
    this.pencilStart = 0;
    this.pencilFilterSpec = documentInkPencilFilter(style.widthPx, { x: 0, y: 0, width: 1, height: 1 }, PENCIL_LIVE_SEED);
    this.pencilBox.style.display = "";
    this.placePencil();
  }

  /**
   * The painted strokes' boxes: clipped to the rect, in host pixels. The
   * clip and the box under it stand on whole host pixels (the rect widened
   * out to them): each is a layer, and a layer at a fraction of a pixel has
   * its origin snapped -- to a device pixel in WebKit, a CSS pixel in
   * Chromium -- with the remainder rounded away by the SVG inside, which
   * moved the stroke being written off where it lands. The painter places
   * each stroke's own box on a whole pixel too, its fraction in a translate
   * (DocumentInkSurface).
   */
  private placePainted(): void {
    const clip = this.paintClip;
    const box = this.paintBox;
    if (!clip || !box) return;
    const rect = this.rect;
    const left = Math.floor(rect.left);
    const top = Math.floor(rect.top);
    Object.assign(clip.style, {
      left: `${left}px`,
      top: `${top}px`,
      width: `${Math.ceil(rect.left + rect.width) - left}px`,
      height: `${Math.ceil(rect.top + rect.height) - top}px`,
    } satisfies Partial<CSSStyleDeclaration>);
    box.style.transform = `translate(${-left}px, ${-top}px)`;
  }

  private paintedBegin(): void {
    const doc = this.element.ownerDocument;
    if (!this.paintClip || !this.paintBox) {
      const clip = doc.createElement("div");
      clip.setAttribute("aria-hidden", "true");
      clip.setAttribute("data-ink-live", "painted");
      Object.assign(clip.style, {
        position: "absolute",
        overflow: "hidden",
        pointerEvents: "none",
        // No compositing layer of its own. Promoted (will-change, contained
        // paint) on iPadOS WebKit, Safari and Chrome alike, the strokes on
        // the surface beside it stopped showing between the lift and the
        
        
      } satisfies Partial<CSSStyleDeclaration>);
      const box = doc.createElement("div");
      Object.assign(box.style, {
        position: "absolute",
        left: "0px",
        top: "0px",
        width: "0px",
        height: "0px",
        overflow: "visible",
      } satisfies Partial<CSSStyleDeclaration>);
      clip.append(box);
      this.pageHost?.append(clip);
      this.paintClip = clip;
      this.paintBox = box;
    }
    this.paintBox.replaceChildren();
    this.paintedPieces.length = 0;
    this.paintedLive = null;
    this.paintedStart = 0;
    this.paintClip.style.display = "";
    this.placePainted();
  }

  /** A new box for one painted piece. */
  private paintedPiece(): PaintedPiece | null {
    const box = this.paintBox;
    if (!box) return null;
    const element = this.element.ownerDocument.createElement("div");
    Object.assign(element.style, { position: "absolute" } satisfies Partial<CSSStyleDeclaration>);
    box.append(element);
    const piece: PaintedPiece = { box: element, cutId: `ink-live-cut-${(pencilFilterIds += 1)}` };
    this.paintedPieces.push(piece);
    return piece;
  }

  /**
   * Paint `run` into a piece with the page's painter and, for a pencil piece,
   * cut it back to its own samples (see `drawPencil` for why it goes in pieces).
   */
  private paintPiece(
    piece: PaintedPiece,
    run: readonly InkPressurePoint[],
    cuts: { start: PencilCut | null; end: PencilCut | null } | null,
  ): void {
    const style = this.style;
    const painter = this.painter;
    if (!style || !painter || run.length === 0) return;
    const painted = painter(piece.box, run);
    if (!painted || !cuts || (!cuts.start && !cuts.end)) return;
    const { svg, originX, originY } = painted;
    const doc = svg.ownerDocument;
    const radius = PENCIL_FILTER_REACH_PX * 2 + style.widthPx;
    const local = (cut: PencilCut): PencilCut => ({ ...cut, x: cut.x - originX, y: cut.y - originY });
    const reach = radius + 8;
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    for (const point of run) {
      minX = Math.min(minX, point.x - originX);
      minY = Math.min(minY, point.y - originY);
      maxX = Math.max(maxX, point.x - originX);
      maxY = Math.max(maxY, point.y - originY);
    }
    const parts = [
      `M ${minX - reach} ${minY - reach} H ${maxX + reach} V ${maxY + reach} H ${minX - reach} Z`,
    ];
    if (cuts.start) parts.push(halfDisc(local(cuts.start), radius, -1));
    if (cuts.end) parts.push(halfDisc(local(cuts.end), radius, 1));
    const clip = doc.createElementNS(SVG_NS, "clipPath");
    clip.id = piece.cutId;
    clip.setAttribute("clipPathUnits", "userSpaceOnUse");
    const shape = doc.createElementNS(SVG_NS, "path");
    shape.setAttribute("d", parts.join(" "));
    shape.setAttribute("clip-rule", "evenodd");
    // Whole pixels: two neighbours' cuts share every pixel out, none twice.
    shape.setAttribute("shape-rendering", "crispEdges");
    clip.append(shape);
    const group = doc.createElementNS(SVG_NS, "g");
    group.setAttribute("clip-path", `url(#${clip.id})`);
    for (const child of Array.from(svg.childNodes)) {
      if (child.nodeName.toLowerCase() === "defs") continue;
      group.append(child);
    }
    svg.append(clip, group);
  }

  /** Every stroke, painted as it lands; a pencil in pieces (see `drawPencil`). */
  private drawPainted(predicted: readonly InkPressurePoint[]): void {
    const style = this.style;
    if (!style) return;
    const points = this.points;
    if (points.length === 0) return;
    const tip = points[points.length - 1]!;
    const withPrediction = (run: InkPressurePoint[]) => {
      for (const guess of predicted) {
        if (!Number.isFinite(guess.x) || !Number.isFinite(guess.y)) continue;
        if (Math.hypot(guess.x - tip.x, guess.y - tip.y) > MAX_PREDICTION_PX) break;
        run.push(guess);
      }
      return run;
    };
    if (!style.pencil && style.opacity >= 1 && !style.multiply && !style.screen && !style.gradient?.length) {
      // An opaque pen: finished pieces stay, the last one is repainted.
      while (points.length - this.paintedStart > PEN_PIECE_POINTS * 2) {
        const end = this.paintedStart + PEN_PIECE_POINTS;
        const piece = this.paintedLive ?? this.paintedPiece();
        if (!piece) return;
        this.paintPiece(piece, points.slice(Math.max(0, this.paintedStart - PEN_PIECE_OVERLAP), end + 1), null);
        this.paintedLive = null;
        this.paintedStart = end;
      }
      const piece = this.paintedLive ?? this.paintedPiece();
      if (!piece) return;
      this.paintedLive = piece;
      this.paintPiece(piece, withPrediction(points.slice(Math.max(0, this.paintedStart - PEN_PIECE_OVERLAP))), null);
      return;
    }
    // A pencil stroke that is still small is drawn whole too (PENCIL_WHOLE_MAX_AREA_PX).
    if (!style.pencil || (this.paintedStart === 0 && this.pencilDrawnWhole(points, style.widthPx))) {
      const piece = this.paintedLive ?? this.paintedPiece();
      if (!piece) return;
      this.paintedLive = piece;
      this.paintPiece(piece, withPrediction([...points]), null);
      return;
    }
    const reach = PENCIL_FILTER_REACH_PX + style.widthPx;
    for (;;) {
      const end = this.paintedStart + PENCIL_PIECE_POINTS;
      if (end >= points.length) break;
      const leadOut = this.pencilLeadOut(end, reach);
      if (leadOut === null) break;
      const piece = this.paintedLive ?? this.paintedPiece();
      if (!piece) return;
      const from = this.paintedStart > 0 ? this.pencilLeadIn(this.paintedStart, reach) : 0;
      this.paintPiece(piece, points.slice(from, leadOut + 1), {
        start: this.paintedStart > 0 ? this.pencilCut(this.paintedStart) : null,
        end: this.pencilCut(end),
      });
      this.paintedLive = null;
      this.paintedStart = end;
    }
    const piece = this.paintedLive ?? this.paintedPiece();
    if (!piece) return;
    this.paintedLive = piece;
    const from = this.paintedStart > 0 ? this.pencilLeadIn(this.paintedStart, reach) : 0;
    this.paintPiece(piece, withPrediction(points.slice(from)), {
      start: this.paintedStart > 0 ? this.pencilCut(this.paintedStart) : null,
      end: null,
    });
  }

  /** Whether a pencil stroke of these samples is still drawn whole (see PENCIL_WHOLE_MAX_AREA_PX). */
  private pencilDrawnWhole(points: readonly InkPressurePoint[], widthPx: number): boolean {
    const region = documentInkPencilRegionFor(points, widthPx);
    return region.width * region.height <= PENCIL_WHOLE_MAX_AREA_PX;
  }

  /** The pencil paths' host pixels from the layer's corner (the rect's corner, snapped). */
  private placePencil(): void {
    if (!this.pencilBox || !this.style?.pencil) return;
    // In the page's own pixels already: nothing to place.
    if (this.pageHost) {
      this.pencilBox.style.transform = "";
      return;
    }
    const scale = this.screenScale || 1;
    this.pencilBox.style.transform =
      `translate(${-this.rect.left + this.cornerOffset.x / scale}px, ${-this.rect.top + this.cornerOffset.y / scale}px)`;
  }

  /** One piece's elements: its path, wearing its own filter region and cut. */
  private pencilPiece(): PencilPiece | null {
    const doc = this.element.ownerDocument;
    const spec = this.pencilFilterSpec;
    if (!spec || !this.pencilBox) return null;
    const id = `ink-live-pencil-${(pencilFilterIds += 1)}`;
    // Each piece an SVG of its own: in one SVG, Safari filtered every finished
    // piece again on every frame, and the frames grew with the stroke.
    const group = doc.createElementNS(SVG_NS, "svg");
    group.setAttribute("aria-hidden", "true");
    group.setAttribute("width", "1");
    group.setAttribute("height", "1");
    Object.assign(group.style, {
      position: "absolute",
      left: "0px",
      top: "0px",
      overflow: "visible",
    } satisfies Partial<CSSStyleDeclaration>);
    const filter = doc.createElementNS(SVG_NS, "filter");
    filter.id = id;
    for (const [name, value] of Object.entries(spec.attrs)) filter.setAttribute(name, value);
    for (const primitive of spec.primitives) {
      const element = doc.createElementNS(SVG_NS, primitive.tag);
      for (const [name, value] of Object.entries(primitive.attrs)) element.setAttribute(name, value);
      filter.append(element);
    }
    const clip = doc.createElementNS(SVG_NS, "clipPath");
    clip.id = `${id}-cut`;
    clip.setAttribute("clipPathUnits", "userSpaceOnUse");
    const cut = doc.createElementNS(SVG_NS, "path");
    // Whole pixels: two neighbours' cuts share every pixel out, none twice.
    cut.setAttribute("shape-rendering", "crispEdges");
    cut.setAttribute("clip-rule", "evenodd");
    clip.append(cut);
    const path = doc.createElementNS(SVG_NS, "path");
    path.setAttribute("filter", `url(#${id})`);
    path.setAttribute("clip-path", `url(#${clip.id})`);
    path.setAttribute("stroke-linecap", "round");
    path.setAttribute("stroke-linejoin", "round");
    group.append(filter, clip, path);
    this.pencilBox.append(group);
    return { group, filter, cut, path };
  }

  /**
   * Draw one piece: samples `from`..`to` (and the prediction when it is the
   * piece being drawn), shown between the cut at `start` and the one at
   * `end` (null: no cut on that side).
   */
  private drawPencilPiece(
    piece: PencilPiece,
    run: readonly InkPressurePoint[],
    cuts: { start: PencilCut | null; end: PencilCut | null },
  ): void {
    const style = this.style;
    if (!style || run.length === 0) return;
    const { path, filter, cut } = piece;
    const corner = style.corner;
    const whole = this.points.map((point) => point.p);
    if (run.length === 1) {
      const tip = run[0]!;
      const r = Math.max(0.25, this.widthAt(tip, tip) / 2);
      path.setAttribute("d", `M ${tip.x - r} ${tip.y} a ${r} ${r} 0 1 0 ${2 * r} 0 a ${r} ${r} 0 1 0 ${-2 * r} 0 Z`);
      path.setAttribute("fill", style.color);
      path.setAttribute("stroke", "none");
    } else if (style.pressure && inkPressuresVary(whole)) {
      path.setAttribute("d", inkVariableWidthOutlinePathData(run, style.widthPx, { corner }));
      path.setAttribute("fill", style.color);
      path.setAttribute("stroke", "none");
    } else {
      path.setAttribute("d", inkStrokePathData(run, { corner }));
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", style.color);
      path.setAttribute(
        "stroke-width",
        String(Math.max(0.5, style.widthPx * (style.pressure ? inkSteadyPressureWidthFactor(whole) : 1))),
      );
    }
    const region = documentInkPencilRegionFor(run, style.widthPx);
    filter.setAttribute("x", String(region.x));
    filter.setAttribute("y", String(region.y));
    filter.setAttribute("width", String(region.width));
    filter.setAttribute("height", String(region.height));
    // The cut: the region's box with, at each cut, the half disc on its far
    // side taken out (even-odd). Only near the boundary: the lead-in and
    // lead-out go no further than the disc, and the piece's own samples that
    // curl back past the cut line further away stay drawn.
    const radius = PENCIL_FILTER_REACH_PX * 2 + style.widthPx;
    const parts = [
      `M ${region.x} ${region.y} H ${region.x + region.width} V ${region.y + region.height} H ${region.x} Z`,
    ];
    if (cuts.start) parts.push(halfDisc(cuts.start, radius, -1));
    if (cuts.end) parts.push(halfDisc(cuts.end, radius, 1));
    cut.setAttribute("d", parts.join(" "));
  }

  /** The line a piece boundary at sample `index` is cut along: across the stroke there. */
  private pencilCut(index: number): PencilCut {
    const points = this.points;
    const at = points[index]!;
    const before = points[Math.max(0, index - 1)]!;
    const after = points[Math.min(points.length - 1, index + 1)]!;
    let dx = after.x - before.x;
    let dy = after.y - before.y;
    const length = Math.hypot(dx, dy);
    if (length < 1e-6) {
      dx = 1;
      dy = 0;
    } else {
      dx /= length;
      dy /= length;
    }
    return { x: at.x, y: at.y, dx, dy };
  }

  /** The first sample of a piece's lead-in: far enough back for the filter to see what it would. */
  private pencilLeadIn(index: number, reach: number): number {
    const points = this.points;
    let travelled = 0;
    let at = index;
    while (at > 0 && (index - at < 2 || travelled < reach)) {
      travelled += Math.hypot(points[at]!.x - points[at - 1]!.x, points[at]!.y - points[at - 1]!.y);
      at -= 1;
    }
    return at;
  }

  /** The last sample of a piece's lead-out, or null while the stroke has not gone that far yet. */
  private pencilLeadOut(index: number, reach: number): number | null {
    const points = this.points;
    let travelled = 0;
    let at = index;
    while (at < points.length - 1 && (at - index < 2 || travelled < reach)) {
      travelled += Math.hypot(points[at + 1]!.x - points[at]!.x, points[at + 1]!.y - points[at]!.y);
      at += 1;
    }
    return at - index >= 2 && travelled >= reach ? at : null;
  }

  /** The pencil stroke so far and on through the prediction. */
  private drawPencil(predicted: readonly InkPressurePoint[]): void {
    const style = this.style;
    if (!style || !this.pencilBox) return;
    const points = this.points;
    if (points.length === 0) return;
    const reach = PENCIL_FILTER_REACH_PX + style.widthPx;
    // Small still: the whole stroke in one piece, as it lands (PENCIL_WHOLE_MAX_AREA_PX).
    if (this.pencilStart === 0 && this.pencilDrawnWhole(points, style.widthPx)) {
      const piece = this.pencilLive ?? this.pencilPiece();
      if (!piece) return;
      this.pencilLive = piece;
      const tip = points[points.length - 1]!;
      const run: InkPressurePoint[] = [...points];
      for (const guess of predicted) {
        if (!Number.isFinite(guess.x) || !Number.isFinite(guess.y)) continue;
        if (Math.hypot(guess.x - tip.x, guess.y - tip.y) > MAX_PREDICTION_PX) break;
        run.push(guess);
      }
      this.drawPencilPiece(piece, run, { start: null, end: null });
      return;
    }
    // Finish the pieces the stroke has gone far enough past.
    for (;;) {
      const end = this.pencilStart + PENCIL_PIECE_POINTS;
      if (end >= points.length) break;
      const leadOut = this.pencilLeadOut(end, reach);
      if (leadOut === null) break;
      const piece = this.pencilLive ?? this.pencilPiece();
      if (!piece) return;
      const from = this.pencilStart > 0 ? this.pencilLeadIn(this.pencilStart, reach) : 0;
      this.drawPencilPiece(piece, points.slice(from, leadOut + 1), {
        start: this.pencilStart > 0 ? this.pencilCut(this.pencilStart) : null,
        end: this.pencilCut(end),
      });
      this.pencilLive = null;
      this.pencilStart = end;
    }
    // The piece being drawn, through the prediction.
    const piece = this.pencilLive ?? this.pencilPiece();
    if (!piece) return;
    this.pencilLive = piece;
    const tip = points[points.length - 1]!;
    const from = this.pencilStart > 0 ? this.pencilLeadIn(this.pencilStart, reach) : 0;
    const run: InkPressurePoint[] = points.slice(from);
    for (const guess of predicted) {
      if (!Number.isFinite(guess.x) || !Number.isFinite(guess.y)) continue;
      if (Math.hypot(guess.x - tip.x, guess.y - tip.y) > MAX_PREDICTION_PX) break;
      run.push(guess);
    }
    this.drawPencilPiece(piece, run, {
      start: this.pencilStart > 0 ? this.pencilCut(this.pencilStart) : null,
      end: null,
    });
  }

  /**
   * Whether this canvas stroke goes in pieces of PEN_PIECE_POINTS samples,
   * and in which look: true a pressure outline, false a line; null to draw
   * it whole. Drawn whole on every frame, a stroke's frame grew with it --
   * the outline of every sample again, each frame (Excalidraw does the same
   * with its stroke being drawn, and its frames grow too). In pieces only
   * the last one is drawn again: an opaque piece overlapping its neighbour
   * by PEN_PIECE_OVERLAP samples draws the pixels of the stroke drawn whole
   * (the canvas is opaque ink; a translucent tool's wash is the element's).
   * Not for a gradient (repainted over the growing box each frame anyway),
   * nor for a line whose steady width still moves with its samples.
   */
  private pieceLook(points: readonly InkPressurePoint[]): boolean | null {
    const style = this.style;
    if (!style || style.gradient?.length) return null;
    if (points.length <= PEN_PIECE_POINTS * 2 && this.piecedOutline === null) return null;
    const pressures = points.map((point) => point.p);
    const outline = Boolean(style.pressure) && inkPressuresVary(pressures);
    if (!outline && style.pressure && pressures.some((pressure) => pressure !== pressures[0])) return null;
    if (this.piecedOutline !== null && this.piecedOutline !== outline) return null;
    return outline;
  }

  /**
   * Put the pieces the stroke has gone past on the ink canvas, once each;
   * the samples of the live piece (from the last finished one, overlapping
   * it) are returned for the tail.
   */
  private drawPieces(ink: CanvasRenderingContext2D, outline: boolean): InkPressurePoint[] {
    const points = this.points;
    if (this.piecedOutline === null) {
      // Until now the stroke was drawn whole on the ink canvas.
      this.clearAround(ink, this.box);
      this.piecedOutline = outline;
      this.piecedStart = 0;
    }
    while (points.length - this.piecedStart > PEN_PIECE_POINTS * 2) {
      const end = this.piecedStart + PEN_PIECE_POINTS;
      this.drawRun(ink, points.slice(Math.max(0, this.piecedStart - PEN_PIECE_OVERLAP), end + 1), outline);
      this.piecedStart = end;
    }
    return points.slice(Math.max(0, this.piecedStart - PEN_PIECE_OVERLAP));
  }

  /** `run` on `context` in the stroke's look: its pressure outline, or its curve at its width. */
  private drawRun(context: CanvasRenderingContext2D, run: readonly InkPressurePoint[], outline: boolean): void {
    const style = this.style;
    if (!style || run.length === 0) return;
    const corner = style.corner;
    if (outline) {
      context.fill(new Path2D(inkVariableWidthOutlinePathData(run, style.widthPx, { corner })));
      return;
    }
    const pressures = run.map((point) => point.p);
    context.lineWidth = Math.max(0.5, style.widthPx * (style.pressure ? inkSteadyPressureWidthFactor(pressures) : 1));
    context.stroke(new Path2D(inkStrokePathData(run, { corner })));
  }

  /** The whole stroke on `context`, as the landed stroke draws it. */
  private drawWhole(context: CanvasRenderingContext2D, points: readonly InkPressurePoint[]): void {
    const style = this.style;
    if (!style) return;
    // Only where the stroke is (its box only grows, so last frame's ink is
    // inside it), not the whole page-sized canvas (see `tailDrawn`).
    this.clearAround(context, this.box);
    if (points.length === 0) return;
    const corner = style.corner;
    const pressures = points.map((point) => point.p);
    if (style.effect) {
      // The effect's layers, the whole stroke at once, as the landed stroke.
      const pressured = Boolean(style.pressure) && inkPressuresVary(pressures);
      const sparkles = this.sparkles;
      if (sparkles && this.sparkleSamples < points.length) {
        // Only the new samples: the walker carries on from the last one.
        sparkles.push(points.slice(this.sparkleSamples));
        this.sparkleSamples = points.length;
      }
      const box = this.box;
      paintInkEffectOnCanvas(
        context,
        style.effect,
        {
          points,
          widthPx: pressured
            ? style.widthPx
            : style.widthPx * (style.pressure ? inkSteadyPressureWidthFactor(pressures) : 1),
          pressured,
          corner,
          ...(box ? { box: { x: box.minX, y: box.minY, width: box.maxX - box.minX, height: box.maxY - box.minY } } : {}),
        },
        sparkles?.dots ?? [],
      );
      return;
    }
    if (style.pressure && inkPressuresVary(pressures)) {
      context.fill(new Path2D(inkVariableWidthOutlinePathData(points, style.widthPx, { corner })));
      return;
    }
    context.lineWidth = Math.max(0.5, style.widthPx * (style.pressure ? inkSteadyPressureWidthFactor(pressures) : 1));
    context.stroke(new Path2D(inkStrokePathData(points, { corner })));
  }

  /**
   * Clear `box` (host px) on `context` with the room the ink reaches past
   * its samples: half the widest pressure width, the round cap, the
   * prediction and a pixel of anti-aliasing. Nothing when there is no box.
   */
  private clearAround(
    context: CanvasRenderingContext2D,
    box: { minX: number; minY: number; maxX: number; maxY: number } | null,
  ): void {
    if (!box) return;
    const effectReach = this.style?.effect ? inkEffectReach(this.style.effect) : 0;
    const reach = (this.style?.widthPx ?? 8) * Math.max(2, effectReach) + MAX_PREDICTION_PX + 4;
    context.clearRect(box.minX - reach, box.minY - reach, box.maxX - box.minX + 2 * reach, box.maxY - box.minY + 2 * reach);
  }

  /**
   * The tail: from the last settled sample to the tip and on through the
   * prediction; first, for a stroke in pieces, the live piece `live` in the
   * pieces' look (`outline`).
   */
  private drawTail(
    predicted: readonly InkPressurePoint[],
    live: readonly InkPressurePoint[] | null = null,
    outline = false,
  ): void {
    const context = this.tailContext;
    if (!context || !this.style) return;
    this.clearAround(context, this.tailDrawn);
    this.tailDrawn = null;
    const points = this.points;
    if (points.length === 0) return;
    const liveBox = live && live.length > 0 ? boxOf(live) : null;
    if (live && liveBox) this.drawRun(context, live, outline);
    if (liveBox && points.length === 1 && predicted.length === 0) {
      this.tailDrawn = liveBox;
      return;
    }
    if (points.length === 1 && predicted.length === 0) {
      this.dot(context, points[0]!);
      this.tailDrawn = boxOf([points[0]!]);
      return;
    }
    const tip = points[points.length - 1]!;
    const ahead: InkPressurePoint[] = [];
    for (const guess of predicted) {
      if (!Number.isFinite(guess.x) || !Number.isFinite(guess.y)) continue;
      if (Math.hypot(guess.x - tip.x, guess.y - tip.y) > MAX_PREDICTION_PX) break;
      ahead.push(guess);
    }
    const run: InkPressurePoint[] = [...points.slice(this.settled), ...ahead];
    // The segment before the run reaches back one sample (its curve's control).
    const segmentsBox = boxOf(this.settled > 0 ? [points[this.settled - 1]!, ...run] : run);
    this.tailDrawn = liveBox && segmentsBox ? unionBox(liveBox, segmentsBox) : liveBox ?? segmentsBox;
    const base = this.settled;
    for (let index = 0; index + 1 < run.length; index += 1) {
      const absolute = base + index;
      const p0 =
        index === 0
          ? absolute > 0 && !this.isCorner(absolute)
            ? points[absolute - 1]!
            : null
          : run[index - 1]!;
      const p3 = index + 2 < run.length ? run[index + 2]! : null;
      this.segment(context, p0, run[index]!, run[index + 1]!, p3);
    }
  }

  /** Take the stroke off the canvases (it is on screen elsewhere now). */
  clear(): void {
    if (this.drawFrame !== null) {
      this.element.ownerDocument.defaultView?.cancelAnimationFrame(this.drawFrame);
      this.drawFrame = null;
    }
    const box = this.box;
    const tail = this.tailDrawn;
    this.points = [];
    this.box = null;
    this.tailDrawn = null;
    this.settled = 0;
    this.piecedStart = 0;
    this.piecedOutline = null;
    this.sparkles = null;
    this.sparkleSamples = 0;
    for (const [canvas, context, drawn] of [
      [this.ink, this.inkContext, box],
      [this.tail, this.tailContext, tail ? unionBox(tail, box) : box],
    ] as const) {
      if (!context) continue;
      // The stroke's own box (the lift lands it; the next stroke follows at
      // once): a full clear of the page-sized canvas only when it is unknown
      // and the canvas may hold something.
      if (drawn) {
        this.clearAround(context, drawn);
        continue;
      }
      if (this.canvasesClean) continue;
      context.setTransform(1, 0, 0, 1, 0, 0);
      context.clearRect(0, 0, canvas.width, canvas.height);
    }
    this.canvasesClean = true;
    this.style = null;
    this.element.style.display = "none";
    this.pencilBox?.replaceChildren();
    this.pencilLive = null;
    this.pencilStart = 0;
    if (this.pencilBox) this.pencilBox.style.display = "none";
    this.paintBox?.replaceChildren();
    this.paintedPieces.length = 0;
    this.paintedLive = null;
    this.paintedStart = 0;
    if (this.paintClip) this.paintClip.style.display = "none";
  }

  /** Drop the canvases (the surface unmounted). */
  dispose(): void {
    this.clear();
    this.element.remove();
    // Free the backing stores at once rather than when the GC gets to them.
    for (const canvas of [this.ink, this.tail]) {
      canvas.width = 0;
      canvas.height = 0;
    }
    this.pencilBox?.remove();
    this.pencilBox = null;
    this.pencilLive = null;
    this.paintClip?.remove();
    this.paintClip = null;
    this.paintBox = null;
  }
}
