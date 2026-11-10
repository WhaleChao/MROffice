/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Hand tool (grab-to-pan) decision + delta math for the PDF reader.
 *
 * Adapted from pdf.js `web/grab_to_pan.js` (Copyright 2013 Rob Wu,
 * https://github.com/Rob--W/grab-to-pan.js, Apache License 2.0): the
 * interactive-target ignore list and the "scroll = start − pointer delta"
 * mapping mirror that reference; the pointer-capture wiring, Space-hold
 * temporary pan and middle-button pan live in the component. Kept pure (the
 * only DOM type touched is a read-only `target`) so every branch is
 * unit-testable without a mounted tree, mirroring `keyboard-nav.ts`.
 */

/** What the viewer knows at pointerdown time when deciding whether to pan. */
export interface PanIntent {
  /** `PointerEvent.button`: 0 = primary, 1 = middle/auxiliary. */
  button: number;
  /** `PointerEvent.pointerType`; undefined tolerated (jsdom MouseEvent). */
  pointerType?: string;
  /** The toolbar hand-tool toggle is armed. */
  panMode: boolean;
  /** Space is held for a temporary pan (Figma / Preview convention). */
  spaceHeld: boolean;
  /**
   * A drag/click page tool owns the primary pointer (snapshot marquee,
   * measure, draw, shape, link, redact, sign, note placement). Arming those
   * tools disarms the pan toggle, but this guard also covers any state skew.
   */
  toolArmed: boolean;
}

/**
 * Whether a pointerdown should START a pan drag.
 *
 *  - Touch never pans here: single-finger touch scrolls natively and the
 *    two-finger pinch handler owns multi-touch.
 *  - Middle-button drag ALWAYS pans (reader convention), independent of the
 *    toggle — it conflicts with nothing (tools use the primary button).
 *  - Space-hold pans the primary button even while a tool is armed (the
 *    temporary-pan override every canvas app ships).
 *  - The toolbar toggle pans the primary button unless a page tool is armed.
 */
export function shouldStartPan(intent: PanIntent): boolean {
  if (!intent || typeof intent.button !== "number") return false;
  if (intent.pointerType === "touch") return false;
  if (intent.button === 1) return true;
  if (intent.button !== 0) return false;
  if (intent.spaceHeld) return true;
  return intent.panMode && !intent.toolArmed;
}

/**
 * Selector of elements a pan press must LEAVE ALONE so pan mode never breaks
 * click behaviors: links, buttons, form fields (AcroForm overlay inputs),
 * editable text, and the annotation selection frame (its drag handles own the
 * pointer). Mirrors pdf.js `GrabToPan.ignoreTarget`'s list, extended with this
 * viewer's frame + contenteditable surfaces.
 */
export const PAN_IGNORE_SELECTOR = [
  "a[href]",
  "button",
  "input",
  "textarea",
  "select",
  "option",
  '[contenteditable="true"]',
  "[data-pdf-annotation-frame]",
].join(", ");

/**
 * True when `target` (or an ancestor) is an interactive element the hand tool
 * must not hijack. Defensive narrowing like `isTypingTarget`: anything without
 * a usable `closest` is treated as non-interactive (pan proceeds — the presses
 * this guards arrive on real elements).
 */
export function isPanIgnoredTarget(target: EventTarget | null): boolean {
  if (!target || typeof target !== "object") return false;
  const node = target as {
    closest?: (selector: string) => unknown;
    isContentEditable?: unknown;
  };
  if (node.isContentEditable === true) return true;
  if (typeof node.closest !== "function") return false;
  try {
    return node.closest(PAN_IGNORE_SELECTOR) != null;
  } catch {
    return false;
  }
}

/** The scroll + pointer position captured when a pan drag starts. */
export interface PanOrigin {
  scrollLeft: number;
  scrollTop: number;
  clientX: number;
  clientY: number;
}

/** A later pointer position within the same drag. */
export interface PanPoint {
  clientX: number;
  clientY: number;
}

/**
 * Map a pointer move to the scroll position that keeps the grabbed content
 * under the cursor: `scroll = originScroll − (pointer − originPointer)`
 * (pdf.js `GrabToPan.#onMouseMove`). Non-finite inputs degrade to the origin
 * scroll so a malformed event never teleports the viewport. Clamping is left
 * to the DOM (scrollLeft/Top assignments clamp natively).
 */
export function panScrollPosition(
  origin: PanOrigin,
  point: PanPoint,
): { scrollLeft: number; scrollTop: number } {
  const dx = Number(point?.clientX) - Number(origin?.clientX);
  const dy = Number(point?.clientY) - Number(origin?.clientY);
  const left = Number(origin?.scrollLeft);
  const top = Number(origin?.scrollTop);
  return {
    scrollLeft: Number.isFinite(dx) && Number.isFinite(left) ? left - dx : left || 0,
    scrollTop: Number.isFinite(dy) && Number.isFinite(top) ? top - dy : top || 0,
  };
}

/**
 * Movement (px, Chebyshev) beyond which a pan drag counts as a REAL pan and
 * the trailing click event must be swallowed (so a pan that ends over the
 * presentation-mode click-to-advance surface, a link, etc. doesn't also
 * "click"). Small jitters below the threshold stay clicks.
 */
export const PAN_CLICK_SUPPRESS_PX = 4;

/** Whether the drag from `origin` to `point` exceeds the click threshold. */
export function exceedsPanThreshold(
  origin: PanPoint,
  point: PanPoint,
  threshold: number = PAN_CLICK_SUPPRESS_PX,
): boolean {
  const dx = Math.abs(Number(point?.clientX) - Number(origin?.clientX));
  const dy = Math.abs(Number(point?.clientY) - Number(origin?.clientY));
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return false;
  return Math.max(dx, dy) > threshold;
}
