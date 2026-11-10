/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { inkStrokePathData } from "./office-ink-smooth";
import { REDUCE_MOTION_ATTRIBUTE } from "./reduce-motion";

/** One stroke as drawn on screen: client pixels, its colour and width. */
export interface OfficeInkFlightStroke {
  readonly points: readonly { readonly x: number; readonly y: number }[];
  readonly color: string;
  readonly width: number;
}

/** Where a reading landed on screen, read once it rendered. */
export interface OfficeInkLanding {
  readonly rect: DOMRect | { left: number; top: number; width: number; height: number };
  /** The new text box, equation or block, which surfaces from a blur. */
  readonly element?: Element | null;
}

export interface OfficeInkFlight {
  /** The writing moved (zoom, the document reflowed) before the reading came: draw it there. */
  readonly redraw: (strokes: readonly OfficeInkFlightStroke[]) => void;
  /** The reading is in: fly to where it landed (read after two frames). */
  readonly land: (landing: () => OfficeInkLanding | null) => void;
  /** The reading did not land (failed, dropped): the copy goes, nothing moves. */
  readonly cancel: () => void;
}

const SVG = "http://www.w3.org/2000/svg";
const FLIGHT_MS = 560;
const SURFACE_MS = 420;
const GLINT_MS = 1100;

export function officeInkMotionAllowed(): boolean {
  if (typeof window === "undefined" || typeof document === "undefined") return false;
  if (document.documentElement.hasAttribute(REDUCE_MOTION_ATTRIBUTE)) return false;
  try {
    return !window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
  } catch {
    return true;
  }
}

interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Where the copy is drawn: the editor's layer (client pixels mapped into its
 * own, scaled and scrolled), or the viewport.
 */
interface FlightSpace {
  readonly host: Element;
  readonly position: "absolute" | "fixed";
  /** Client pixels to the host's own. */
  readonly toLocal: (x: number, y: number) => { x: number; y: number };
  /** Client lengths to the host's own. */
  readonly scale: number;
}

const VIEWPORT_SPACE: Omit<FlightSpace, "host"> = {
  position: "fixed",
  toLocal: (x, y) => ({ x, y }),
  scale: 1,
};

/**
 * The layer's coordinates as they stand now: its box on screen over its own
 * size is its zoom (a CSS transform or `zoom`), its border and scroll offset
 * the rest. A layer that is no containing block for an absolute child (not
 * positioned, not transformed) cannot carry the copy: the viewport does.
 */
function spaceOf(layer: Element | null | undefined): FlightSpace {
  const body = document.body;
  if (!(layer instanceof HTMLElement) || !layer.isConnected) return { host: body, ...VIEWPORT_SPACE };
  let style: CSSStyleDeclaration | null = null;
  try {
    style = getComputedStyle(layer);
  } catch {
    style = null;
  }
  if (!style || (style.position === "static" && (!style.transform || style.transform === "none"))) {
    return { host: body, ...VIEWPORT_SPACE };
  }
  const rect = layer.getBoundingClientRect();
  const scaleX = layer.offsetWidth > 0 && rect.width > 0 ? rect.width / layer.offsetWidth : 1;
  const scaleY = layer.offsetHeight > 0 && rect.height > 0 ? rect.height / layer.offsetHeight : scaleX;
  const sx = Number.isFinite(scaleX) && scaleX > 0 ? scaleX : 1;
  const sy = Number.isFinite(scaleY) && scaleY > 0 ? scaleY : sx;
  return {
    host: layer,
    position: "absolute",
    toLocal: (x, y) => ({
      x: (x - rect.left) / sx - layer.clientLeft + layer.scrollLeft,
      y: (y - rect.top) / sy - layer.clientTop + layer.scrollTop,
    }),
    scale: sx,
  };
}

/** The strokes in the space's own pixels. */
function localStrokes(strokes: readonly OfficeInkFlightStroke[], space: FlightSpace): OfficeInkFlightStroke[] {
  return strokes.map((stroke) => ({
    ...stroke,
    width: stroke.width / space.scale,
    points: stroke.points.map((point) => space.toLocal(point.x, point.y)),
  }));
}

function localBox(rect: Box, space: FlightSpace): Box {
  const start = space.toLocal(rect.left, rect.top);
  const end = space.toLocal(rect.left + rect.width, rect.top + rect.height);
  return { left: start.x, top: start.y, width: end.x - start.x, height: end.y - start.y };
}

function boundsOf(strokes: readonly OfficeInkFlightStroke[]): Box | null {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const stroke of strokes) {
    const pad = Math.max(1, stroke.width) / 2 + 2;
    for (const point of stroke.points) {
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
      left = Math.min(left, point.x - pad);
      top = Math.min(top, point.y - pad);
      right = Math.max(right, point.x + pad);
      bottom = Math.max(bottom, point.y + pad);
    }
  }
  if (!Number.isFinite(left) || right - left < 1 || bottom - top < 1) return null;
  return { left, top, width: right - left, height: bottom - top };
}

function pathData(points: OfficeInkFlightStroke["points"], box: Box): string {
  if (points.length === 1) {
    const { x, y } = points[0]!;
    return `M${x - box.left} ${y - box.top}h0.01`;
  }
  // the curve the stroke itself is drawn as, so the copy that flies is the ink
  return inkStrokePathData(points.map((point) => ({ x: point.x - box.left, y: point.y - box.top })));
}

function paint(svg: SVGSVGElement, strokes: readonly OfficeInkFlightStroke[], glint: boolean): Box | null {
  const box = boundsOf(strokes);
  svg.replaceChildren();
  if (!box) return null;
  Object.assign(svg.style, {
    left: `${box.left}px`,
    top: `${box.top}px`,
    width: `${box.width}px`,
    height: `${box.height}px`,
  });
  svg.setAttribute("viewBox", `0 0 ${box.width} ${box.height}`);
  for (const stroke of strokes) {
    const d = pathData(stroke.points, box);
    const ink = document.createElementNS(SVG, "path");
    ink.setAttribute("d", d);
    ink.setAttribute("fill", "none");
    ink.setAttribute("stroke", stroke.color);
    ink.setAttribute("stroke-width", String(Math.max(1, stroke.width)));
    ink.setAttribute("stroke-linecap", "round");
    ink.setAttribute("stroke-linejoin", "round");
    svg.append(ink);
    if (!glint) continue;
    // A short bright run travelling along the stroke while it is read.
    const run = document.createElementNS(SVG, "path");
    run.setAttribute("d", d);
    run.setAttribute("fill", "none");
    run.setAttribute("pathLength", "100");
    run.setAttribute("stroke", "rgba(255, 255, 255, 0.78)");
    run.setAttribute("stroke-width", String(Math.max(1, stroke.width * 0.62)));
    run.setAttribute("stroke-linecap", "round");
    run.setAttribute("stroke-dasharray", "16 84");
    run.setAttribute("data-office-ink-glint", "");
    svg.append(run);
    run.animate?.(
      [{ strokeDashoffset: "100" }, { strokeDashoffset: "0" }],
      { duration: GLINT_MS, iterations: Infinity, easing: "linear" },
    );
  }
  return box;
}

function nextFrames(count: number, then: () => void): void {
  if (count <= 0) {
    then();
    return;
  }
  requestAnimationFrame(() => nextFrames(count - 1, then));
}

/**
 * Begin the motion for a piece of writing being read. Null when motion is
 * off or the strokes cannot be drawn (the conversion then just happens).
 * `layer` is the editor's element the ink scrolls and zooms with; the copy
 * is drawn inside it (see the module note). Without one it is pinned to the
 * viewport.
 */
export function startOfficeInkFlight(
  strokes: readonly OfficeInkFlightStroke[] | null,
  layer?: Element | null,
): OfficeInkFlight | null {
  if (!strokes || strokes.length === 0 || !officeInkMotionAllowed()) return null;
  const svg = document.createElementNS(SVG, "svg");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("data-office-ink-flight", "reading");
  const space = spaceOf(layer);
  svg.setAttribute("data-office-ink-flight-space", space.position === "absolute" ? "layer" : "viewport");
  Object.assign(svg.style, {
    position: space.position,
    pointerEvents: "none",
    zIndex: "2147482000",
    overflow: "visible",
    transformOrigin: "0 0",
  });
  let box = paint(svg, localStrokes(strokes, space), true);
  if (!box) return null;
  space.host.append(svg);
  let finished = false;
  const remove = () => {
    finished = true;
    svg.remove();
  };
  return {
    redraw(next) {
      if (finished) return;
      const moved = paint(svg, localStrokes(next, spaceOf(space.position === "absolute" ? space.host : null)), true);
      if (moved) box = moved;
    },
    cancel: remove,
    land(landing) {
      if (finished) return;
      // The ink copy stops glinting and stays where the writing was while
      // the editor renders the reading in its place.
      for (const run of svg.querySelectorAll("[data-office-ink-glint]")) run.remove();
      svg.setAttribute("data-office-ink-flight", "landing");
      nextFrames(2, () => {
        if (finished || !box) return;
        let landed: OfficeInkLanding | null = null;
        try {
          landed = landing();
        } catch {
          landed = null;
        }
        const source = box;
        // The reading's place in the copy's own pixels (the layer may have
        // scrolled or zoomed since the copy was drawn).
        const now = spaceOf(space.position === "absolute" ? space.host : null);
        const target =
          landed?.rect && landed.rect.width > 0 && landed.rect.height > 0 ? localBox(landed.rect, now) : source;
        // Keep the handwriting's proportions: it shrinks (or grows) to the
        // height of what it became and settles on its centre.
        const scale = Math.max(0.12, Math.min(1.6, Math.min(target.height / source.height, target.width / source.width) || 1));
        const dx = target.left + target.width / 2 - (source.left + (source.width * scale) / 2);
        const dy = target.top + target.height / 2 - (source.top + (source.height * scale) / 2);
        if (typeof svg.animate !== "function") {
          remove();
          return;
        }
        const flight = svg.animate(
          [
            { transform: "translate(0px, 0px) scale(1)", opacity: 1, filter: "blur(0px)" },
            {
              transform: `translate(${dx * 0.55}px, ${dy * 0.55}px) scale(${1 - (1 - scale) * 0.6})`,
              opacity: 0.85,
              filter: "blur(0.3px)",
              offset: 0.5,
            },
            { transform: `translate(${dx}px, ${dy}px) scale(${scale})`, opacity: 0, filter: "blur(1.6px)" },
          ],
          { duration: FLIGHT_MS, easing: "cubic-bezier(0.32, 0, 0.18, 1)", fill: "forwards" },
        );
        flight.onfinish = remove;
        flight.oncancel = remove;
        window.setTimeout(remove, FLIGHT_MS + 400);
        const element = landed?.element;
        if (element instanceof HTMLElement || element instanceof SVGElement) {
          element.animate?.(
            [
              { opacity: 0, filter: "blur(5px)", transform: "scale(0.96)" },
              { opacity: 1, filter: "blur(0px)", transform: "scale(1)" },
            ],
            { duration: SURFACE_MS, delay: FLIGHT_MS * 0.35, easing: "cubic-bezier(0.2, 0, 0, 1)", fill: "backwards" },
          );
        } else if (landed) {
          surfaceGlow(target, now);
        }
      });
    },
  };
}

/** Text typed into a line has no element of its own: a soft light where it landed (`target` in `space`'s pixels). */
function surfaceGlow(target: Box, space: FlightSpace): void {
  const glow = document.createElement("div");
  glow.setAttribute("aria-hidden", "true");
  glow.setAttribute("data-office-ink-flight", "surface");
  Object.assign(glow.style, {
    position: space.position,
    left: `${target.left - 3}px`,
    top: `${target.top - 2}px`,
    width: `${target.width + 6}px`,
    height: `${target.height + 4}px`,
    borderRadius: "4px",
    pointerEvents: "none",
    zIndex: "2147481999",
    background: "color-mix(in srgb, var(--console-accent, #0a7d72) 22%, transparent)",
    opacity: "0",
  });
  space.host.append(glow);
  if (typeof glow.animate !== "function") {
    glow.remove();
    return;
  }
  const fade = glow.animate(
    [{ opacity: 0 }, { opacity: 1, offset: 0.35 }, { opacity: 0 }],
    { duration: SURFACE_MS + 260, delay: FLIGHT_MS * 0.4, easing: "ease-out", fill: "both" },
  );
  fade.onfinish = () => glow.remove();
  window.setTimeout(() => glow.remove(), FLIGHT_MS + SURFACE_MS + 900);
}
