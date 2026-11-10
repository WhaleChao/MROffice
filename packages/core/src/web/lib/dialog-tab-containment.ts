/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Tab and Shift+Tab kept inside a hand-built modal dialog (one not on the
 * shared OfficeDialogShell, whose Radix layer already contains focus): from
 * its last control Tab goes back to its first, and Shift+Tab from its first
 * to its last, so the keyboard never reaches the sheet behind the dimmer.
 */

const TABBABLE =
  'a[href], input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), ' +
  'button:not([disabled]), [tabindex]:not([tabindex="-1"]), [contenteditable="true"]';

function shown(element: HTMLElement): boolean {
  if (!element.isConnected || element.closest("[inert], [hidden], [aria-hidden='true']")) return false;
  return element.getClientRects().length > 0;
}

/** The controls Tab visits in `dialog`, in order. */
export function dialogTabStops(dialog: HTMLElement): HTMLElement[] {
  return Array.from(dialog.querySelectorAll<HTMLElement>(TABBABLE)).filter(
    (candidate) => candidate.tabIndex >= 0 && shown(candidate),
  );
}

/**
 * Call from the dialog's keydown. Returns true when it moved focus (the event
 * is then default-prevented, so an outer dialog's handler leaves it alone).
 */
export function containDialogTab(
  event: { key: string; shiftKey: boolean; defaultPrevented: boolean; preventDefault(): void },
  dialog: HTMLElement | null,
): boolean {
  if (event.key !== "Tab" || event.defaultPrevented || !dialog) return false;
  const stops = dialogTabStops(dialog);
  const active = dialog.ownerDocument.activeElement as HTMLElement | null;
  if (stops.length === 0) {
    event.preventDefault();
    dialog.focus({ preventScroll: true });
    return true;
  }
  const index = active ? stops.indexOf(active) : -1;
  const inside = Boolean(active && dialog.contains(active));
  const wrapBack = event.shiftKey && (index === 0 || (index < 0 && !inside) || active === dialog);
  const wrapForward = !event.shiftKey && (index === stops.length - 1 || (index < 0 && !inside));
  if (!wrapBack && !wrapForward) return false;
  event.preventDefault();
  stops[wrapBack ? stops.length - 1 : 0].focus({ preventScroll: true });
  return true;
}
