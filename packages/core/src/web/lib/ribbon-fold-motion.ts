/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cssZoomOf } from "./css-zoom";

/**
 * The ribbon folding to its tab row and back, without laying the window out
 * on every frame.
 *
 * Animating the strip's height made the page below it (a worksheet's grid, a
 * document's pages) re-measure and re-render at every frame: 25-50 ms frames
 * for a 220 ms glide. Here the window is laid out once, in its final state,
 * and what moved is slid back from where it was on the compositor (FLIP):
 *
 * - Everything below the ribbon that moved is translated from its old place
 *   to its new one.
 * - The part that grew or shrank (the grid, the pages) keeps its bottom edge
 *   still by a clip that follows the slide. Shrinking (the ribbon unfolding),
 *   it keeps its old height for the glide, the extra hanging past its slot
 *   under a negative margin, so no gap opens above what stands below it.
 * - Folding, the strip stays painted where it was (out of the flow); the
 *   ribbon is clipped from its bottom edge in step with the page sliding up
 *   over it, and uncovered the same way as the page slides down when it
 *   unfolds.
 */

interface Mover {
  element: HTMLElement;
  /** Its height before the fold. */
  height: number;
  dy: number;
  /** The box's height change: the part that grows or shrinks. */
  dh: number;
  styles: { transform: string; clipPath: string; zIndex: string; flex: string; height: string; marginBottom: string; willChange: string };
}

function below(root: HTMLElement, stop: HTMLElement | null): HTMLElement[] {
  const top = root.getBoundingClientRect().top;
  const found: HTMLElement[] = [];
  let node: HTMLElement = root;
  while (node.parentElement) {
    const parent = node.parentElement;
    for (const sibling of Array.from(parent.children)) {
      if (sibling === node || !(sibling instanceof HTMLElement)) continue;
      const position = getComputedStyle(sibling).position;
      if (position === "absolute" || position === "fixed") continue;
      const rect = sibling.getBoundingClientRect();
      if (rect.height > 0 && rect.top > top + 0.5) found.push(sibling);
    }
    if (parent === stop || parent === document.body) break;
    node = parent;
  }
  return found;
}

export interface RibbonFoldGlide {
  /** Jump to the end and put every borrowed style back. */
  finish(): void;
}

/**
 * Fold (`collapsed`) or unfold the ribbon rooted at `root`, whose command
 * strip is `strip`; `apply` states the new fold on the DOM (the attribute the
 * stylesheet folds by). Returns the running glide.
 */
export function glideRibbonFold(
  root: HTMLElement,
  strip: HTMLElement,
  collapsed: boolean,
  apply: () => void,
  timing: { duration: number; easing: string },
): RibbonFoldGlide {
  const stop = root.closest<HTMLElement>(".native-page-window");
  const movers = below(root, stop);
  const before = movers.map((element) => element.getBoundingClientRect());
  const rootRect = root.getBoundingClientRect();
  const stripRect = strip.getBoundingClientRect();
  const stripStyles = {
    position: strip.style.position,
    left: strip.style.left,
    top: strip.style.top,
    width: strip.style.width,
    display: strip.style.getPropertyValue("display"),
    displayPriority: strip.style.getPropertyPriority("display"),
    zIndex: strip.style.zIndex,
  };
  const stripDisplay = getComputedStyle(strip).display;
  const rootPosition = root.style.position;
  const host = strip.parentElement;
  const hostOverflow = host?.style.overflow ?? "";
  const rootClip = root.style.clipPath;
  const rootOverflow = root.style.overflow;

  apply();
  if (collapsed && stripRect.height > 0) {
    // Out of the flow, where it was: the page slides up over it.
    if (getComputedStyle(root).position === "static") root.style.position = "relative";
    // Neither the folded band nor the ribbon's own box clips it (the clip
    // below draws the ribbon's edge for the glide).
    if (host && host !== root) host.style.overflow = "visible";
    root.style.overflow = "visible";
    strip.style.position = "absolute";
    strip.style.width = `${stripRect.width}px`;
    strip.style.zIndex = "0";
    strip.style.setProperty("display", stripDisplay === "none" ? "block" : stripDisplay, "important");
    // Placed against the box it is now positioned in, wherever that is.
    const frame = (strip.offsetParent as HTMLElement | null) ?? root;
    const frameRect = frame.getBoundingClientRect();
    strip.style.left = `${stripRect.left - frameRect.left - frame.clientLeft}px`;
    strip.style.top = `${stripRect.top - frameRect.top - frame.clientTop}px`;
    // Then measured where it actually landed and moved by the difference.
    // Under a touch-scaled region the strip is drawn at a CSS zoom: its own
    // lengths are multiplied by it, and the box it is offset against need not
    // be the one measured (it was placed over the tab row, hiding the tabs for
    // the first frames of the fold). Screen pixels converted by the zoom it is
    // drawn at land it exactly where it was, in either engine's zoom model.
    for (let pass = 0; pass < 2; pass += 1) {
      const placed = strip.getBoundingClientRect();
      const dx = stripRect.left - placed.left;
      const dy = stripRect.top - placed.top;
      const dw = stripRect.width - placed.width;
      if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5 && Math.abs(dw) < 0.5) break;
      const zoom = cssZoomOf(strip) || 1;
      strip.style.left = `${(parseFloat(strip.style.left) || 0) + dx / zoom}px`;
      strip.style.top = `${(parseFloat(strip.style.top) || 0) + dy / zoom}px`;
      if (placed.width > 0) strip.style.width = `${((parseFloat(strip.style.width) || 0) * stripRect.width) / placed.width}px`;
    }
  }

  const after = movers.map((element) => element.getBoundingClientRect());
  const moved: Mover[] = [];
  movers.forEach((element, index) => {
    const dy = before[index].top - after[index].top;
    const dh = after[index].height - before[index].height;
    if (Math.abs(dy) < 0.5 && Math.abs(dh) < 0.5) return;
    moved.push({
      element,
      height: before[index].height,
      dy,
      dh,
      styles: {
        transform: element.style.transform,
        clipPath: element.style.clipPath,
        zIndex: element.style.zIndex,
        flex: element.style.flex,
        height: element.style.height,
        marginBottom: element.style.marginBottom,
        willChange: element.style.willChange,
      },
    });
  });

  const animations: Animation[] = [];
  const options: KeyframeAnimationOptions = { duration: timing.duration, easing: timing.easing, fill: "both" };
  for (const mover of moved) {
    const { element, dy, dh, height } = mover;
    // Above the strip that stays painted under it.
    element.style.zIndex = "1";
    element.style.willChange = "transform";
    let extra = 0;
    if (dh < -0.5) {
      // Shrinking: it keeps its old height for the glide, hanging past its
      // slot, so nothing below it opens a gap.
      extra = -dh;
      element.style.flex = "none";
      element.style.height = `${height}px`;
      element.style.marginBottom = `${dh}px`;
    }
    const grows = Math.abs(dh) >= 0.5;
    // The clip that keeps a growing or shrinking part's bottom edge still:
    // the box's bottom past its final bottom edge, at each end of the slide.
    const clipAt = (shift: number) => `inset(0 0 ${Math.max(0, shift + extra)}px 0)`;
    animations.push(
      element.animate(
        grows
          ? [
              { transform: `translateY(${dy}px)`, clipPath: clipAt(dy) },
              { transform: "translateY(0px)", clipPath: clipAt(0) },
            ]
          : [{ transform: `translateY(${dy}px)` }, { transform: "translateY(0px)" }],
        options,
      ),
    );
  }
  // The ribbon shows only where the page has left it: its box is clipped
  // from the bottom edge in step with the slide (past its folded box while
  // the strip still hangs below it), so nothing of it is ever drawn over the
  // page, whatever stacks above what.
  const travel = collapsed ? stripRect.height : Math.max(0, root.getBoundingClientRect().height - rootRect.height);
  if (travel > 0.5) {
    const clip = (px: number) => `inset(0 0 ${px}px 0)`;
    animations.push(
      root.animate(
        collapsed ? [{ clipPath: clip(-travel) }, { clipPath: clip(0) }] : [{ clipPath: clip(travel) }, { clipPath: clip(0) }],
        options,
      ),
    );
  }

  let done = false;
  const restore = () => {
    if (done) return;
    done = true;
    for (const animation of animations) animation.cancel();
    for (const { element, styles } of moved) {
      element.style.transform = styles.transform;
      element.style.clipPath = styles.clipPath;
      element.style.zIndex = styles.zIndex;
      element.style.flex = styles.flex;
      element.style.height = styles.height;
      element.style.marginBottom = styles.marginBottom;
      element.style.willChange = styles.willChange;
    }
    root.style.position = rootPosition;
    root.style.clipPath = rootClip;
    root.style.overflow = rootOverflow;
    if (host && host !== root) host.style.overflow = hostOverflow;
    strip.style.position = stripStyles.position;
    strip.style.left = stripStyles.left;
    strip.style.top = stripStyles.top;
    strip.style.width = stripStyles.width;
    strip.style.zIndex = stripStyles.zIndex;
    if (stripStyles.display) strip.style.setProperty("display", stripStyles.display, stripStyles.displayPriority);
    else strip.style.removeProperty("display");
  };
  void Promise.all(animations.map((animation) => animation.finished)).then(restore, () => {});
  return { finish: restore };
}
