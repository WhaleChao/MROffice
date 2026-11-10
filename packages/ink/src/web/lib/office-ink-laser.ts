/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Laser Pointer (a note-taking app's laser, as in Goodnotes): the pen
 * draws a bright line that fades away behind it and is never saved -- for
 * pointing at the page while presenting or sharing it. The trail keeps the
 * deck's present-mode laser rules (lib/presentation-annotation: samples age
 * out, the newest is brightest) over a longer life, since it is drawn, not
 * merely followed. Each editor paints it into an SVG group in its own
 * surface's pixels.
 */

import {
  pushTrailSample,
  pruneTrail,
  trailRenderPoints,
  type TrailSample,
} from "@mrnote/core/web/lib/presentation-annotation.ts";

/** How long a sample of the laser's line stays before it has faded. */
export const INK_LASER_LIFETIME_MS = 900;
/** The most samples a trail holds (oldest dropped first). */
export const INK_LASER_MAX_SAMPLES = 180;
/** The laser's colour: the bright red of a pointer's dot. */
export const INK_LASER_COLOR = "#ff2d2d";

export interface InkLaserTrail {
  /** In the unit square of the surface (the shared trail rules clamp to it). */
  samples: TrailSample[];
  /** The pen is down: its tip is drawn as the dot. */
  down: boolean;
}

export function createInkLaserTrail(): InkLaserTrail {
  return { samples: [], down: false };
}

export interface InkLaserSpace {
  readonly width: number;
  readonly height: number;
}

/** A sample at (x, y) of a surface `space` pixels large. */
export function addInkLaserSample(trail: InkLaserTrail, x: number, y: number, space: InkLaserSpace, now: number): void {
  const width = space.width > 0 ? space.width : 1;
  const height = space.height > 0 ? space.height : 1;
  trail.samples = pushTrailSample(trail.samples, x / width, y / height, now, INK_LASER_LIFETIME_MS, INK_LASER_MAX_SAMPLES);
}

/** Whether there is anything left to draw (the fade loop stops when not). */
export function inkLaserTrailLive(trail: InkLaserTrail, now: number): boolean {
  trail.samples = pruneTrail(trail.samples, now, INK_LASER_LIFETIME_MS);
  return trail.down || trail.samples.length > 0;
}

const SVG_NS = "http://www.w3.org/2000/svg";

/**
 * Paints the trail into `group` (replacing what it held): segments thinning
 * and fading with age, under a soft glow, and the tip's dot while the pen is
 * down. `widthPx` is the line's width at its newest.
 */
export function paintInkLaserTrail(
  group: SVGGElement,
  trail: InkLaserTrail,
  space: InkLaserSpace,
  now: number,
  widthPx = 4,
): void {
  const ownerDocument = group.ownerDocument;
  group.replaceChildren();
  const points = trailRenderPoints(trail.samples, now, INK_LASER_LIFETIME_MS).map((point) => ({
    ...point,
    x: point.x * space.width,
    y: point.y * space.height,
  }));
  group.setAttribute("fill", "none");
  group.setAttribute("stroke", INK_LASER_COLOR);
  group.setAttribute("stroke-linecap", "round");
  group.setAttribute("stroke-linejoin", "round");
  group.style.filter = `drop-shadow(0 0 ${Math.max(2, widthPx)}px ${INK_LASER_COLOR})`;
  for (let index = 1; index < points.length; index += 1) {
    const from = points[index - 1]!;
    const to = points[index]!;
    if (to.opacity <= 0) continue;
    const line = ownerDocument.createElementNS(SVG_NS, "line");
    line.setAttribute("x1", String(from.x));
    line.setAttribute("y1", String(from.y));
    line.setAttribute("x2", String(to.x));
    line.setAttribute("y2", String(to.y));
    line.setAttribute("stroke-opacity", String(to.opacity));
    line.setAttribute("stroke-width", String(widthPx * (0.35 + 0.65 * to.opacity)));
    group.append(line);
  }
  const last = trail.samples[trail.samples.length - 1];
  const tip = last ? { x: last.x * space.width, y: last.y * space.height } : null;
  if (trail.down && tip) {
    const dot = ownerDocument.createElementNS(SVG_NS, "circle");
    dot.setAttribute("cx", String(tip.x));
    dot.setAttribute("cy", String(tip.y));
    dot.setAttribute("r", String(widthPx * 1.1));
    dot.setAttribute("fill", INK_LASER_COLOR);
    dot.setAttribute("stroke", "#ffffff");
    dot.setAttribute("stroke-width", String(Math.max(1, widthPx * 0.35)));
    group.append(dot);
  }
}
