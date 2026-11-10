/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Keeps `:hover` on what the pointer is over after a scroller's content moved
 * under a pointer that did not move.
 *
 * A browser decides `:hover` when the pointer moves. Desktop engines also
 * decide it again a moment after a scroll (a synthetic mouse move); WebKit on
 * iPadOS does not, so with a trackpad or mouse there a swipe over a ribbon's
 * tab list left the highlight on the tab that slid away: the pointer rested
 * on "References" and "Review" stayed lit until the pointer moved again.
 *
 * No script can ask a browser to decide `:hover` again, but an element taken
 * out of the document stops being hovered at once, in every engine. So when a
 * scroll has settled, the hovered element the pointer is no longer inside is
 * taken out and put back where it was: the same node, in the same place, no
 * longer lit. The right element lights up with the browser's next decision
 * (at once on desktop, with the next pointer move on iPadOS); nothing is ever
 * lit that the pointer is not on.
 */

export interface HoverPointer {
  readonly x: number;
  readonly y: number;
}

function contains(element: Element, pointer: HoverPointer): boolean {
  const box = element.getBoundingClientRect();
  return (
    pointer.x >= box.left &&
    pointer.x < box.right &&
    pointer.y >= box.top &&
    pointer.y < box.bottom
  );
}

/**
 * Of the elements a browser reports as hovered (ancestors first), the
 * outermost one the pointer is not inside: the one whose highlight is stale.
 * Null when every hovered element still holds the pointer.
 */
export function staleHoverTarget(
  hovered: readonly Element[],
  pointer: HoverPointer,
): HTMLElement | null {
  for (const element of hovered) {
    if (!(element instanceof HTMLElement)) continue;
    if (!contains(element, pointer)) return element;
  }
  return null;
}

/** Content that restarts or reloads when its element leaves the document. */
const UNSAFE_TO_REINSERT = "iframe, video, audio, canvas, object, embed";

/**
 * Take `element` out of the document and put it back where it was. The node,
 * its place, the scroller's scroll offsets and a focus held inside it are the
 * same afterwards; only its hover is gone.
 */
export function reinsertInPlace(element: HTMLElement, scroller: HTMLElement): boolean {
  const parent = element.parentNode;
  if (!parent || element === scroller || !scroller.contains(element)) return false;
  if (element.matches(UNSAFE_TO_REINSERT) || element.querySelector(UNSAFE_TO_REINSERT)) {
    return false;
  }
  const next = element.nextSibling;
  const active = element.ownerDocument.activeElement;
  const focused =
    active instanceof HTMLElement && element.contains(active) ? active : null;
  const left = scroller.scrollLeft;
  const top = scroller.scrollTop;
  // A scrolled list inside the element (a gallery) starts from its top again
  // when its element comes back into the document: keep where each one was.
  const scrolled: Array<[Element, number, number]> = [];
  for (const node of [element, ...Array.from(element.querySelectorAll("*"))]) {
    if (node.scrollLeft !== 0 || node.scrollTop !== 0) {
      scrolled.push([node, node.scrollLeft, node.scrollTop]);
    }
  }
  element.remove();
  parent.insertBefore(element, next);
  for (const [node, nodeLeft, nodeTop] of scrolled) {
    node.scrollLeft = nodeLeft;
    node.scrollTop = nodeTop;
  }
  if (scroller.scrollLeft !== left) scroller.scrollLeft = left;
  if (scroller.scrollTop !== top) scroller.scrollTop = top;
  focused?.focus({ preventScroll: true });
  return true;
}

/** The elements inside `scroller` the browser reports as hovered, ancestors first. */
export type HoveredReader = (scroller: HTMLElement) => readonly Element[];

const browserHovered: HoveredReader = (scroller) => {
  try {
    return Array.from(scroller.querySelectorAll(":hover"));
  } catch {
    return [];
  }
};

/** Clear the hover the pointer has left inside `scroller`; the element it was cleared from. */
export function clearStaleHover(
  scroller: HTMLElement,
  pointer: HoverPointer | null,
  hoveredIn: HoveredReader = browserHovered,
): HTMLElement | null {
  if (!pointer || !scroller.isConnected) return null;
  const stale = staleHoverTarget(hoveredIn(scroller), pointer);
  return stale && reinsertInPlace(stale, scroller) ? stale : null;
}

/** How long after the last scroll event the content counts as at rest. */
const SCROLL_SETTLE_MS = 90;

/**
 * Watch one scroller: once a scroll has settled with a hovering pointer over
 * it, clear the hover that pointer has left. Returns the unsubscribe.
 */
export function resyncHoverAfterScroll(
  scroller: HTMLElement,
  hoveredIn?: HoveredReader,
): () => void {
  let pointer: HoverPointer | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;
  const onPointer = (event: PointerEvent) => {
    // A finger does not hover; a mouse, a trackpad pointer and a pen do.
    pointer = event.pointerType === "touch" ? null : { x: event.clientX, y: event.clientY };
  };
  const onLeave = () => {
    pointer = null;
  };
  const settle = () => {
    timer = null;
    clearStaleHover(scroller, pointer, hoveredIn);
  };
  const onScroll = () => {
    if (!pointer) return;
    if (timer) clearTimeout(timer);
    timer = setTimeout(settle, SCROLL_SETTLE_MS);
  };
  scroller.addEventListener("pointermove", onPointer, { passive: true });
  scroller.addEventListener("pointerenter", onPointer, { passive: true });
  scroller.addEventListener("pointerdown", onPointer, { passive: true });
  scroller.addEventListener("pointerleave", onLeave, { passive: true });
  scroller.addEventListener("pointercancel", onLeave, { passive: true });
  scroller.addEventListener("scroll", onScroll, { passive: true });
  return () => {
    if (timer) clearTimeout(timer);
    scroller.removeEventListener("pointermove", onPointer);
    scroller.removeEventListener("pointerenter", onPointer);
    scroller.removeEventListener("pointerdown", onPointer);
    scroller.removeEventListener("pointerleave", onLeave);
    scroller.removeEventListener("pointercancel", onLeave);
    scroller.removeEventListener("scroll", onScroll);
  };
}
