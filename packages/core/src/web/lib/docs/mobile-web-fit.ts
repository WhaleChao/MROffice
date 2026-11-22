/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A phone's Web Layout (the reading view, Word's Mobile View) stands each
 * object placed on the page -- a shape, a text box, a chart, a picture -- in
 * the text's flow at its anchor (styles/note-editor.css). Its wrapper is
 * already as wide as the screen, but what it holds is still drawn at its
 * page geometry: a 550 px shape parked 140 px left of its anchor, a warped
 * word-art at its page offset. This measures each wrapper's drawing and
 * writes the custom properties the stylesheet turns into a translate (the
 * drawing's left edge to the wrapper's), a scale (down to the wrapper's
 * width, never up) and the height it then takes in the flow.
 *
 * The scale is visual only (CSS `scale` does not lay the drawing out again),
 * so the file and Print Layout never see it; leaving the view clears it.
 */

const FIT_PROPS = ["--mobile-web-fit-shift", "--mobile-web-fit-scale", "--mobile-web-fit-height"] as const;

function clearFit(host: HTMLElement): void {
  for (const prop of FIT_PROPS) host.style.removeProperty(prop);
}

/** The box (relative to `host`) every painted descendant of `host` covers. */
function drawingExtent(host: HTMLElement): { left: number; top: number; right: number; bottom: number } | null {
  const origin = host.getBoundingClientRect();
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const element of host.querySelectorAll<HTMLElement | SVGElement>("*")) {
    const box = element.getBoundingClientRect();
    if (box.width < 1 || box.height < 1) continue;
    left = Math.min(left, box.left - origin.left);
    top = Math.min(top, box.top - origin.top);
    right = Math.max(right, box.right - origin.left);
    bottom = Math.max(bottom, box.bottom - origin.top);
  }
  return Number.isFinite(left) ? { left, top, right, bottom } : null;
}

/** Fit every placed object of `content` (the editor's content root). */
export function fitMobileWebObjects(content: HTMLElement): void {
  const hosts = [...content.querySelectorAll<HTMLElement>("[data-float-mode]")].filter(
    // the outermost wrapper only: a group's members ride their group's fit
    (host) => !host.parentElement?.closest("[data-float-mode]"),
  );
  // measure every drawing with no fit applied, then write: one layout pass
  for (const host of hosts) clearFit(host);
  const plans = hosts.map((host) => {
    const extent = drawingExtent(host);
    const width = host.clientWidth;
    return { host, extent, width, overlay: host.getAttribute("data-float-mode") === "overlay" };
  });
  for (const { host, extent, width, overlay } of plans) {
    if (!extent || width <= 0) continue;
    const drawn = extent.right - extent.left;
    const scale = drawn > width ? width / drawn : 1;
    const shiftX = -extent.left * scale + (drawn * scale < width ? (width - drawn * scale) / 2 : 0);
    const shiftY = -Math.min(0, extent.top) * scale;
    const needsShift = Math.abs(shiftX) > 0.5 || Math.abs(shiftY) > 0.5;
    if (scale < 1 || needsShift) {
      host.style.setProperty("--mobile-web-fit-shift", `${shiftX.toFixed(1)}px ${shiftY.toFixed(1)}px`);
      if (scale < 1) host.style.setProperty("--mobile-web-fit-scale", scale.toFixed(4));
    }
    // An overlay's wrapper takes no height of its own; a scaled drawing's
    // layout box keeps its unscaled height: both take the drawn height.
    if (overlay || scale < 1) {
      const height = (extent.bottom - Math.min(0, extent.top)) * scale;
      host.style.setProperty("--mobile-web-fit-height", `${Math.max(0, Math.ceil(height))}px`);
    }
  }
}

/** Remove every fit written under `content`. */
export function clearMobileWebObjects(content: HTMLElement): void {
  for (const host of content.querySelectorAll<HTMLElement>("[data-float-mode]")) clearFit(host);
}

/** Whether the phone reading view's flow rules apply (the stylesheet's two conditions). */
export function mobileWebFlowActive(): boolean {
  if (typeof window === "undefined") return false;
  return (
    window.matchMedia("(max-width: 600px)").matches ||
    document.documentElement.getAttribute("data-mobile-layout") === "on"
  );
}

/**
 * Keep `content`'s placed objects fitted while the phone reading view lasts:
 * on start, when the window or the content resizes and when objects are
 * added, removed or re-placed (debounced to one frame). Returns the stop function, which clears every fit.
 */
export function watchMobileWebObjects(content: HTMLElement): () => void {
  let frame = 0;
  const schedule = () => {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      if (mobileWebFlowActive()) fitMobileWebObjects(content);
      else clearMobileWebObjects(content);
    });
  };
  const resize = new ResizeObserver(schedule);
  resize.observe(content);
  // Objects added, removed or re-placed; typing that grows a line resizes
  // the content, which the ResizeObserver above already reports. Not every
  // keystroke: measuring each drawing is a layout read.
  const mutations = new MutationObserver(schedule);
  mutations.observe(content, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ["data-float-mode", "data-float-anchor", "data-document-view-mode"],
  });
  window.addEventListener("resize", schedule);
  schedule();
  return () => {
    if (frame) cancelAnimationFrame(frame);
    resize.disconnect();
    mutations.disconnect();
    window.removeEventListener("resize", schedule);
    clearMobileWebObjects(content);
  };
}
