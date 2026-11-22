/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Plugins' context-menu items (manifest `contextMenu`), added to the
 * editors' own right-click menu (components/ui/PointerContextMenu, which
 * every editor's menu is built on).
 *
 * The editor's plugin dock claims the next menu when a right-click happens
 * inside its editor (a capture listener, before the editor opens its menu);
 * the menu takes the claim once and appends the items below its own. A menu
 * opened anywhere else (a notebook's page list, a menu opened later from the
 * keyboard) finds no fresh claim and shows only its own items.
 */

export interface PluginContextMenuItem {
  readonly id: string;
  readonly label: string;
  readonly onSelect: () => void;
}

const FRESH_MS = 1_500;
let claim: { at: number; items: () => readonly PluginContextMenuItem[] } | null = null;

/** The dock's claim: ``items`` is asked when the menu opens. */
export function claimPluginContextMenu(items: () => readonly PluginContextMenuItem[]): void {
  claim = { at: Date.now(), items };
}

/** The items of a fresh claim (once), or none. */
export function takePluginContextMenu(): readonly PluginContextMenuItem[] {
  const taken = claim;
  claim = null;
  if (!taken || Date.now() - taken.at > FRESH_MS) return [];
  try {
    return taken.items().slice(0, 20);
  } catch {
    return [];
  }
}
