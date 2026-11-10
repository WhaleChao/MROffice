/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draws `position: fixed` elements at viewport coordinates even when an
 * ancestor is their containing block.
 *
 * Every workspace window hints a transform (`will-change: transform`), and a
 * transform, filter, backdrop filter or containment on any ancestor makes that
 * ancestor, not the viewport, the box a fixed child is placed in. A popup
 * placed at a pointer's or a selection's client coordinates inside a window
 * was therefore drawn shifted by the window's own offset (and scaled with the
 * canvas zoom). A zero-size probe in the same parent reads that box's origin
 * and scale, so the element lands exactly on the requested point whatever
 * transform the element itself carries.
 */
export function fixedContainingBlock(parent: Element): {
  x: number;
  y: number;
  scaleX: number;
  scaleY: number;
} {
  const probe = parent.ownerDocument.createElement("div");
  probe.setAttribute("aria-hidden", "true");
  probe.style.cssText =
    "position:fixed;left:0;top:0;width:0;height:0;margin:0;visibility:hidden;pointer-events:none;";
  parent.appendChild(probe);
  const origin = probe.getBoundingClientRect();
  probe.style.left = "100px";
  probe.style.top = "100px";
  const unit = probe.getBoundingClientRect();
  probe.remove();
  return {
    x: origin.left,
    y: origin.top,
    scaleX: (unit.left - origin.left) / 100 || 1,
    scaleY: (unit.top - origin.top) / 100 || 1,
  };
}

/** The fixed `left`/`top` that put a child of `parent` at a viewport point. */
export function viewportPointInFixedParent(
  parent: Element,
  left: number,
  top: number,
): { left: number; top: number } {
  const block = fixedContainingBlock(parent);
  return {
    left: (left - block.x) / block.scaleX,
    top: (top - block.y) / block.scaleY,
  };
}

/** Moves a fixed element so its (untransformed) corner is at the point. */
export function pinFixedToViewport(
  element: HTMLElement,
  left: number,
  top: number,
): void {
  const parent = element.parentElement;
  if (!parent) return;
  const placed = viewportPointInFixedParent(parent, left, top);
  element.style.left = `${placed.left}px`;
  element.style.top = `${placed.top}px`;
}
