/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Native Popover escapes clipping, not DOM ancestry. A private native editor
 * nested in a Copilot preview owns its shortcuts; the source editor must not
 * save, undo or open another command merely because it contains that focus. */
export function nativeOfficeOwnsFocus(
  root: HTMLElement | null,
  active: Element | null,
): boolean {
  const interaction = active?.closest("[data-copilot-interaction]");
  return Boolean(
    root &&
    active &&
    root.contains(active) &&
    active.closest(".presentation-window, .spreadsheet-window") === root &&
    (!interaction || interaction.contains(root)),
  );
}
