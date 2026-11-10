/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Rendering density is independent of PDF coordinates and editor geometry.
 * Use the same canvas zoom published by WindowCanvas for world projection,
 * scoped to this viewer so a second notebook cannot supply its zoom. */
export function readPdfDisplayDensity(root: HTMLElement | null): number {
  const rawZoom = Number(root?.closest("[data-canvas-zoom]")?.getAttribute("data-canvas-zoom"));
  const zoom = Number.isFinite(rawZoom) && rawZoom > 0 ? rawZoom : 1;
  const rawDpr = root?.ownerDocument.defaultView?.devicePixelRatio ?? 1;
  const dpr = Number.isFinite(rawDpr) && rawDpr > 0 ? Math.min(2, rawDpr) : 1;
  return dpr * zoom;
}

/** CSS transforms do not trigger ResizeObserver. Observe the canvas's public
 * zoom attribute, plus device-scale changes when moving between screens.
 * Coalesce wheel/dial updates without rebuilding the text or annotation layer. */
export function observePdfDisplayDensity(root: HTMLElement, changed: () => void): () => void {
  const view = root.ownerDocument.defaultView;
  if (!view) return () => undefined;
  let density = readPdfDisplayDensity(root);
  let frame: number | null = null;
  let disposed = false;
  let media: MediaQueryList | null = null;
  const measure = () => {
    frame = null;
    if (disposed) return;
    const next = readPdfDisplayDensity(root);
    if (next !== density) { density = next; changed(); }
  };
  const schedule = () => {
    if (frame !== null) return;
    frame = 0;
    const id = view.requestAnimationFrame(measure);
    if (frame === 0) frame = id;
  };
  const deviceChanged = () => { watchDevice(); schedule(); };
  const watchDevice = () => {
    media?.removeEventListener?.("change", deviceChanged);
    media = view.matchMedia?.(`(resolution: ${view.devicePixelRatio || 1}dppx)`) ?? null;
    media?.addEventListener?.("change", deviceChanged);
  };
  const canvas = root.closest("[data-canvas-zoom]");
  const observer = canvas ? new MutationObserver(schedule) : null;
  observer?.observe(canvas!, { attributes: true, attributeFilter: ["data-canvas-zoom"] });
  view.addEventListener("resize", schedule);
  watchDevice();
  // A load/render may have started before this effect mounted.
  changed();
  return () => {
    disposed = true;
    observer?.disconnect();
    view.removeEventListener("resize", schedule);
    media?.removeEventListener?.("change", deviceChanged);
    if (frame !== null) view.cancelAnimationFrame(frame);
  };
}
