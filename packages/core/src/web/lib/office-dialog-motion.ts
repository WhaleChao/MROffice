/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The way an Office command dialog leaves the screen, for every editor.
 *
 * A dialog's entrance is plain CSS (OfficeDialogChrome.css): the card and its
 * backdrop animate when they mount. Its exit cannot be: the editors unmount a
 * card the moment its command ends, and an unmounted element has nothing left
 * to animate. Giving each of the ~80 dialogs a delayed unmount would put a
 * timer and a "closing" state into every one of them.
 *
 * Instead one observer watches for a dialog card or backdrop leaving the
 * document and puts a lifeless copy of it back where it was, just long enough
 * to fade out. The copy is inert and carries no ids, test ids, roles or form
 * names, so focus, the keyboard, assistive technology and anything that looks
 * a dialog up by selector see the dialog as gone the instant its command says
 * so; only the picture lingers. Dialogs that already own their exit (Radix
 * content, which carries `data-state`) are left to it.
 */

import { exitMotionAllowed, inertGhostOf } from "./inert-ghost";

/** Marks the fading copy; live-dialog lookups exclude it. */
export const OFFICE_DIALOG_GHOST_ATTRIBUTE = "data-office-dialog-ghost";

/** What leaves with an exit: a dialog card, or the dimmer behind an in-window one. */
const LEAVING_SELECTOR = ".office-dialog, .native-window__modal-dimmer";

/** Longer than the CSS exit (OfficeDialogChrome.css), so the copy is never cut short. */
const GHOST_LIFETIME_MS = 220;

let installed = false;

function leavesWithMotion(node: Node): node is HTMLElement {
  if (!(node instanceof HTMLElement)) return false;
  if (node.hasAttribute(OFFICE_DIALOG_GHOST_ATTRIBUTE)) return false;
  // The card itself, or the layer that holds it (an app-modal backdrop with
  // its card). A wrapper with a dialog somewhere deeper (a whole window
  // closing) takes everything with it and does not fade.
  const card = node.matches(LEAVING_SELECTOR)
    ? node
    : Array.from(node.children).find((child): child is HTMLElement => child instanceof HTMLElement && child.matches(".office-dialog"));
  if (!card) return false;
  // A task pane wears the card's chrome but leaves with its own motion
  // (components/office-ribbon/officeTaskPaneMotion).
  if (card.matches(".office-dialog--pane")) return false;
  // A Radix shell that was closed (`open` set to false) has already played
  // its exit by the time it is removed. One whose caller unmounted it while
  // open has not, and leaves like any other card.
  return node.getAttribute("data-state") !== "closed" && card.getAttribute("data-state") !== "closed";
}

/** The dim a Radix shell paints as the card's sibling, removed open with it. */
function isOpenRadixLayer(node: Node): node is HTMLElement {
  return (
    node instanceof HTMLElement &&
    node.getAttribute("data-state") === "open" &&
    !node.hasAttribute(OFFICE_DIALOG_GHOST_ATTRIBUTE) &&
    !node.matches(".office-dialog") &&
    node.childElementCount === 0
  );
}

function ghostOf(source: HTMLElement): HTMLElement {
  const copy = inertGhostOf(source, OFFICE_DIALOG_GHOST_ATTRIBUTE);
  // A Radix card or dim carries its own exit motion for the closed state.
  if (copy.hasAttribute("data-state")) copy.setAttribute("data-state", "closed");
  return copy;
}

/**
 * Start playing dialog exits in this document. Idempotent; does nothing on
 * the server or in a test DOM. Whether the reader allows motion is asked at
 * each exit, so Settings > Reduce motion takes effect at once.
 */
export function installOfficeDialogMotion(): void {
  if (installed || typeof window === "undefined" || typeof MutationObserver === "undefined") return;
  if (typeof window.matchMedia !== "function" || /jsdom/i.test(window.navigator.userAgent)) return;
  installed = true;
  const leave = (removed: HTMLElement, record: MutationRecord, parent: Element) => {
    const ghost = ghostOf(removed);
    const before = record.nextSibling && record.nextSibling.parentNode === parent ? record.nextSibling : null;
    parent.insertBefore(ghost, before);
    window.setTimeout(() => ghost.remove(), GHOST_LIFETIME_MS);
  };
  const observer = new MutationObserver((records) => {
    // Asked only when a dialog is actually leaving, not on every edit.
    let allowed: boolean | null = null;
    const motionAllowed = () => (allowed ??= exitMotionAllowed(window));
    // Parents a Radix card left while open: its dim leaves with it.
    const radixParents = new Set<Node>();
    for (const record of records) {
      for (const removed of record.removedNodes) {
        if (removed instanceof HTMLElement && removed.matches(".office-dialog[data-state='open']")) radixParents.add(record.target);
      }
    }
    for (const record of records) {
      if (record.removedNodes.length === 0) continue;
      const parent = record.target;
      if (!(parent instanceof Element) || !parent.isConnected) continue;
      for (const removed of record.removedNodes) {
        // The same dialog put back in the same turn (a re-keyed render) has
        // not left.
        if (removed.isConnected) continue;
        if (leavesWithMotion(removed)) {
          if (motionAllowed()) leave(removed, record, parent);
        } else if (radixParents.has(parent) && isOpenRadixLayer(removed)) {
          if (motionAllowed()) leave(removed, record, parent);
        }
      }
    }
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
}
