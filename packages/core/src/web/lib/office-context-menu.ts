/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The right-click menu grammar the three Office editors share, aligned with
 * Office on the web (Word, Excel and PowerPoint for the web, 2025-2026):
 *
 *  - every row is an icon column, a label and, at the right edge, either the
 *    command's keyboard shortcut or a submenu chevron;
 *  - the clipboard group (Cut, Copy, Paste[, Paste Special]) always comes
 *    first; groups are separated by a thin rule;
 *  - a command that does not apply stays listed and greyed, never hidden;
 *  - shortcuts are written the way the web apps write them: "⌘+X" on a Mac,
 *    "Ctrl+X" elsewhere.
 *
 * `formatOfficeShortcut` turns a platform-neutral spec such as "Mod+Shift+M"
 * into that text; the OFFICE_SHORTCUTS table holds the ones Office on the web
 * shows in its context menus.
 */

import { isMacPlatform } from "./mac-platform";

const MAC_MODIFIERS: Record<string, string> = {
  Mod: "⌘",
  Cmd: "⌘",
  Ctrl: "⌃",
  Alt: "⌥",
  Shift: "⇧",
};

const OTHER_MODIFIERS: Record<string, string> = {
  Mod: "Ctrl",
  Cmd: "Ctrl",
  Ctrl: "Ctrl",
  Alt: "Alt",
  Shift: "Shift",
};

/** "Mod+Shift+M" -> "⌘+⇧+M" (Mac) or "Ctrl+Shift+M". Order as Office writes
 *  it: Ctrl/⌘ first, then Alt/⌥, then Shift/⇧, then the key. */
export function formatOfficeShortcut(spec: string, mac = isMacPlatform()): string {
  const parts = spec.split("+").map((part) => part.trim()).filter(Boolean);
  if (parts.length === 0) return "";
  const key = parts[parts.length - 1];
  const modifiers = parts.slice(0, -1);
  const table = mac ? MAC_MODIFIERS : OTHER_MODIFIERS;
  const order = ["Mod", "Cmd", "Ctrl", "Alt", "Shift"];
  const sorted = [...modifiers].sort(
    (a, b) => order.indexOf(a) - order.indexOf(b),
  );
  const shownKey = key.length === 1 ? key.toUpperCase() : key;
  return [...sorted.map((m) => table[m] ?? m), shownKey].join("+");
}

/** The shortcuts Office on the web prints beside its context-menu commands. */
export const OFFICE_SHORTCUTS = {
  cut: "Mod+X",
  copy: "Mod+C",
  paste: "Mod+V",
  formatCells: "Mod+1",
  link: "Mod+K",
  newComment: "Mod+Alt+M",
  newNote: "Shift+F2",
  find: "Mod+F",
  bold: "Mod+B",
  italic: "Mod+I",
  underline: "Mod+U",
} as const;

export type OfficeShortcutName = keyof typeof OFFICE_SHORTCUTS;

export function officeShortcut(name: OfficeShortcutName): string {
  return formatOfficeShortcut(OFFICE_SHORTCUTS[name]);
}

/**
 * The class every right-click menu panel (and each of its submenus) wears,
 * whatever builds it: the shared PointerContextMenu, or a Radix menu anchored
 * to its row (the slide rail, the sheet tabs). It carries the one size
 * specification, written as tokens in styles/office-context-menu.css (where
 * the specification and its sources are set out).
 */
export const OFFICE_CONTEXT_MENU_CLASS = "office-context-menu";
/** Added beside OFFICE_CONTEXT_MENU_CLASS on a submenu panel. */
export const OFFICE_CONTEXT_MENU_SUBMENU_CLASS = "office-context-menu--submenu";

/** The part classes a hand-built row uses (a Radix row's icon, shortcut). */
export const OFFICE_CONTEXT_MENU_PART = {
  icon: "office-context-menu__icon",
  check: "office-context-menu__check",
  label: "office-context-menu__label",
  shortcut: "office-context-menu__shortcut",
  detail: "office-context-menu__detail",
  chevron: "office-context-menu__chevron",
  heading: "office-context-menu__heading",
  iconButton: "office-context-menu__icon-button",
} as const;

/**
 * The specification in numbers (CSS pixels): the values the tokens in
 * styles/office-context-menu.css hold, kept here so a test can hold every
 * editor's menu to them. `touch` is Touch mode (`data-touch-mode` on <html>),
 * which swaps the whole set on the menu at once.
 */
export const OFFICE_CONTEXT_MENU_METRICS = {
  mouse: {
    fontSize: 13,
    lineHeight: 18,
    rowMinHeight: 28,
    rowPaddingBlock: 5,
    rowPaddingInline: 10,
    rowGap: 8,
    rowRadius: 6,
    iconSize: 16,
    iconColumn: 20,
    shortcutFontSize: 12,
    shortcutGap: 24,
    chevronSize: 14,
    separatorMarginBlock: 4,
    panelPadding: 4,
    panelRadius: 10,
    minWidth: 220,
    maxWidth: 320,
    submenuMinWidth: 180,
  },
  touch: {
    fontSize: 15,
    lineHeight: 20,
    rowMinHeight: 44,
    rowPaddingBlock: 8,
    rowPaddingInline: 12,
    rowGap: 10,
    rowRadius: 8,
    iconSize: 20,
    iconColumn: 24,
    shortcutFontSize: 13,
    shortcutGap: 24,
    chevronSize: 16,
    separatorMarginBlock: 5,
    panelPadding: 5,
    panelRadius: 12,
    minWidth: 260,
    maxWidth: 360,
    submenuMinWidth: 220,
  },
} as const;

/** The panel class list for a right-click menu (or one of its submenus). */
export function officeContextMenuClassName(
  options: { readonly submenu?: boolean; readonly className?: string } = {},
): string {
  return [
    OFFICE_CONTEXT_MENU_CLASS,
    options.submenu ? OFFICE_CONTEXT_MENU_SUBMENU_CLASS : null,
    options.className || null,
  ]
    .filter(Boolean)
    .join(" ");
}
