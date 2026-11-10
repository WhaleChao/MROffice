/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { canvasCoveredLeft } from "../canvas-covered-edge";

/**
 * The visible part of the free canvas around a window, in canvas (world)
 * units: what Arrange All and View Side by Side lay windows out in.
 */
export interface OfficeWindowViewport {
  height: number;
  width: number;
  x: number;
  y: number;
}

export interface OfficeWindowRect {
  height: number;
  width: number;
  x: number;
  y: number;
}

function readCanvasPlacement(windowId: string): {
  shell: HTMLElement;
  viewport: HTMLElement;
  zoom: number;
  /** The viewport's top-left corner, in canvas (world) units. */
  originX: number;
  originY: number;
  visible: OfficeWindowViewport;
} | null {
  if (typeof document === "undefined" || !windowId) return null;
  const shell = Array.from(
    document.querySelectorAll<HTMLElement>("[data-window-id]"),
  ).find((candidate) => candidate.dataset.windowId === windowId);
  const viewport = shell?.closest<HTMLElement>(".wm-canvas-viewport") ?? null;
  if (!shell || !viewport) return null;
  const zoom = Number.parseFloat(viewport.dataset.canvasZoom || "1");
  if (!Number.isFinite(zoom) || zoom <= 0) return null;
  const worldX = Number.parseFloat(viewport.dataset.worldOriginX || "0");
  const worldY = Number.parseFloat(viewport.dataset.worldOriginY || "0");
  const originX =
    (Number.isFinite(worldX) ? worldX : 0) + viewport.scrollLeft / zoom;
  const originY =
    (Number.isFinite(worldY) ? worldY : 0) + viewport.scrollTop / zoom;
  // The strip under the notebook's floating sidebar is not in view: windows
  // arranged there would sit under its glass.
  const covered = canvasCoveredLeft(viewport) / zoom;
  return {
    shell,
    viewport,
    zoom,
    originX,
    originY,
    visible: {
      x: originX + covered,
      y: originY,
      width: viewport.clientWidth / zoom - covered,
      height: viewport.clientHeight / zoom,
    },
  };
}

/** The canvas viewport a window sits in, or null outside the canvas. */
export function readOfficeWindowViewport(
  windowId: string,
): OfficeWindowViewport | null {
  return readCanvasPlacement(windowId)?.visible ?? null;
}

/**
 * A window's frame on the canvas as drawn, in the same canvas (world) units
 * as its viewport, or null outside the canvas.
 */
export function readOfficeWindowRect(
  windowId: string,
): OfficeWindowRect | null {
  const placement = readCanvasPlacement(windowId);
  if (!placement) return null;
  const { shell, viewport, zoom, originX, originY } = placement;
  const frame = viewport.getBoundingClientRect();
  const drawn = shell.getBoundingClientRect();
  if (!(drawn.width > 0) || !(drawn.height > 0)) return null;
  return {
    x: originX + (drawn.left - frame.left - viewport.clientLeft) / zoom,
    y: originY + (drawn.top - frame.top - viewport.clientTop) / zoom,
    width: drawn.width / zoom,
    height: drawn.height / zoom,
  };
}
