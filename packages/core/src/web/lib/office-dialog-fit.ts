/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Keeps every Office command dialog inside the part of its window the user
 * can see (live audit S-030).
 *
 * A card is centred in the window that opened it: `.app-modal-overlay`
 * (Format Cells, Fill Effects...) and `.native-window__modal-dialog` are
 * `position: fixed; inset: 0`, and the window's transformed shell makes that
 * the window, not the page. A window taller than the canvas that shows it,
 * or partly scrolled out of it, then centres the card where nobody can see
 * it: in a 1280x540 page Format Cells' title sat above the canvas and OK /
 * Cancel below it. CSS cannot intersect the window with the canvas, so this
 * observer does: while a dialog's frame reaches past the visible region (the
 * page, cut by every clipping ancestor), the frame is inset to that region
 * and the card is told its height (`--office-dialog-fit-height`), which the
 * chrome bounds the card by (OfficeDialogChrome.css). A frame that fits is
 * left exactly as it was.
 */

/** The visible height a card may use, set on the card while its frame is inset. */
export const OFFICE_DIALOG_FIT_HEIGHT = "--office-dialog-fit-height";

const FITTED_ATTRIBUTE = "data-office-dialog-fitted";

let installed = false;

type Box = { top: number; left: number; bottom: number; right: number };

/** The frame a card is centred in: its overlay, or the card itself. */
function frameOf(card: HTMLElement): HTMLElement | null {
  const parent = card.parentElement;
  if (parent?.classList.contains("app-modal-overlay")) return parent;
  if (card.classList.contains("native-window__modal-dialog")) return card;
  return null;
}

function establishesFixedContainingBlock(style: CSSStyleDeclaration): boolean {
  return (
    style.transform !== "none" ||
    style.perspective !== "none" ||
    style.filter !== "none" ||
    /\b(transform|perspective|filter)\b/.test(style.willChange) ||
    /\b(paint|layout|strict|content)\b/.test(style.contain)
  );
}

/** The region of the page the frame can show (null: nothing to adjust). */
function visibleRegion(frame: HTMLElement): { visible: Box; container: Box } | null {
  let visible: Box = { top: 0, left: 0, bottom: window.innerHeight, right: window.innerWidth };
  let container: Box | null = null;
  for (let node = frame.parentElement; node && node !== document.documentElement; node = node.parentElement) {
    const style = getComputedStyle(node);
    const rect = node.getBoundingClientRect();
    if (!container && establishesFixedContainingBlock(style)) container = rect;
    if (style.overflowX !== "visible" || style.overflowY !== "visible") {
      visible = {
        top: Math.max(visible.top, rect.top),
        left: Math.max(visible.left, rect.left),
        bottom: Math.min(visible.bottom, rect.bottom),
        right: Math.min(visible.right, rect.right),
      };
    }
  }
  if (visible.bottom - visible.top < 48 || visible.right - visible.left < 48) return null;
  return { visible, container: container ?? { top: 0, left: 0, bottom: window.innerHeight, right: window.innerWidth } };
}

function clearFit(frame: HTMLElement, card: HTMLElement): void {
  if (!frame.hasAttribute(FITTED_ATTRIBUTE)) return;
  frame.removeAttribute(FITTED_ATTRIBUTE);
  for (const side of ["top", "left", "bottom", "right"] as const) frame.style.removeProperty(side);
  card.style.removeProperty(OFFICE_DIALOG_FIT_HEIGHT);
}

function fit(card: HTMLElement): void {
  const frame = frameOf(card);
  if (!frame || !card.isConnected) return;
  const region = visibleRegion(frame);
  if (!region) return;
  const { visible, container } = region;
  // The frame's own box is its containing block (inset: 0).
  const inside =
    container.top >= visible.top - 0.5 &&
    container.left >= visible.left - 0.5 &&
    container.bottom <= visible.bottom + 0.5 &&
    container.right <= visible.right + 0.5;
  if (inside) {
    clearFit(frame, card);
    return;
  }
  frame.setAttribute(FITTED_ATTRIBUTE, "");
  frame.style.top = `${Math.max(0, visible.top - container.top)}px`;
  frame.style.left = `${Math.max(0, visible.left - container.left)}px`;
  frame.style.bottom = `${Math.max(0, container.bottom - visible.bottom)}px`;
  frame.style.right = `${Math.max(0, container.right - visible.right)}px`;
  card.style.setProperty(OFFICE_DIALOG_FIT_HEIGHT, `${Math.round(visible.bottom - visible.top)}px`);
}

/**
 * Start fitting dialogs in this document. Idempotent; does nothing on the
 * server or in a DOM without layout (a test DOM's boxes are all empty).
 */
export function installOfficeDialogFit(): void {
  if (installed || typeof window === "undefined" || typeof MutationObserver === "undefined") return;
  installed = true;
  const cards = new Set<HTMLElement>();
  let frame = 0;
  const schedule = () => {
    if (frame || cards.size === 0) return;
    frame = window.requestAnimationFrame(() => {
      frame = 0;
      for (const card of [...cards]) {
        if (!card.isConnected) cards.delete(card);
        else fit(card);
      }
    });
  };
  const resizes = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(schedule);
  // Only the card itself, or the overlay that holds it, is looked at: an
  // added grid row or editor node costs one class test, never a subtree walk.
  const track = (node: Node): boolean => {
    if (!(node instanceof HTMLElement)) return false;
    const card = node.classList.contains("office-dialog")
      ? node
      : node.classList.contains("app-modal-overlay")
        ? Array.from(node.children).find((child): child is HTMLElement => child instanceof HTMLElement && child.classList.contains("office-dialog"))
        : undefined;
    if (!card || card.hasAttribute("data-office-dialog-ghost") || !frameOf(card) || cards.has(card)) return false;
    cards.add(card);
    resizes?.observe(card);
    return true;
  };
  new MutationObserver((records) => {
    let added = false;
    for (const record of records) {
      for (const node of record.addedNodes) added = track(node) || added;
    }
    if (added) schedule();
  }).observe(document.documentElement, { childList: true, subtree: true });
  window.addEventListener("resize", schedule, { passive: true });
  // The canvas scrolling a window under its open dialog.
  window.addEventListener("scroll", schedule, { capture: true, passive: true });
  for (const card of document.querySelectorAll<HTMLElement>(".office-dialog")) track(card);
  schedule();
}
