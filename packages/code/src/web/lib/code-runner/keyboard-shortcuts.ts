/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editor's keyboard map, as a readable table (VS Code's Keyboard Shortcuts).
 *
 * The editor binds enough chords that discovering them by accident stopped
 * being realistic. This is the one place that names them, grouped the way a
 * person looks for them rather than the way they happen to be registered.
 *
 * Labels reuse the command palette's existing message keys wherever the same
 * action appears there, so the two surfaces cannot drift into describing the
 * same command differently.
 */

export type ShortcutGroup =
  | "navigate"
  | "edit"
  | "intelligence"
  | "run"
  | "view";

export interface EditorShortcut {
  /** Stable id (also the React key). */
  id: string;
  group: ShortcutGroup;
  /**
   * Bound by the HOST, not by the editor's own keymap — a code window has a
   * console, an output panel and a debugger; a notebook cell has none of them.
   * Rows marked this way are listed only when the host says it binds them, so
   * the table never offers a chord that does nothing where you are reading it.
   */
  hostBound?: true;
  /**
   * Chord tokens. `Mod` renders as ⌘ on macOS and Ctrl elsewhere; a chord with
   * two steps (VS Code's Ctrl+K prefix) is written as two token arrays.
   */
  keys: readonly (readonly string[])[];
  /** Message key for the action's name. */
  labelKey: string;
}

/**
 * Every chord the editor actually binds. Adding a row here without binding the
 * chord would be a lie in the UI, so the list stays hand-checked against the
 * keymap rather than generated from a guess.
 */
export const EDITOR_SHORTCUTS: readonly EditorShortcut[] = [
  // --- Navigate -----------------------------------------------------------
  { id: "find", group: "navigate", keys: [["Mod", "F"]], labelKey: "code.palette.cmd.find" },
  { id: "findNext", group: "navigate", keys: [["F3"]], labelKey: "code.find.nextMatch" },
  { id: "findPrevious", group: "navigate", keys: [["Shift", "F3"]], labelKey: "code.find.previousMatch" },
  { id: "goToLine", group: "navigate", keys: [["Ctrl", "G"]], labelKey: "code.palette.cmd.goToLine" },
  { id: "goToSymbol", group: "navigate", keys: [["Mod", "Shift", "O"]], labelKey: "code.palette.cmd.goToSymbol" },
  { id: "commandPalette", group: "navigate", keys: [["Mod", "Shift", "P"]], labelKey: "code.palette.title" },
  { id: "problems", group: "navigate", keys: [["Mod", "Shift", "M"]], labelKey: "code.palette.cmd.toggleProblems" },
  { id: "shortcuts", group: "navigate", keys: [["Mod", "K"], ["Mod", "S"]], labelKey: "code.shortcuts.title" },

  // --- Edit ---------------------------------------------------------------
  { id: "save", group: "edit", keys: [["Mod", "S"]], labelKey: "pages.save" },
  { id: "toggleComment", group: "edit", keys: [["Mod", "/"]], labelKey: "code.contextMenu.toggleComment" },
  { id: "formatDocument", group: "edit", keys: [["Shift", "Alt", "F"]], labelKey: "code.palette.cmd.formatDocument" },
  { id: "selectNextOccurrence", group: "edit", keys: [["Mod", "D"]], labelKey: "code.shortcuts.selectNextOccurrence" },
  { id: "changeAllOccurrences", group: "edit", keys: [["Mod", "Shift", "L"]], labelKey: "code.palette.cmd.changeAllOccurrences" },
  { id: "expandSelection", group: "edit", keys: [["Mod", "I"]], labelKey: "code.palette.cmd.expandSelection" },
  { id: "moveLineUp", group: "edit", keys: [["Alt", "↑"]], labelKey: "code.shortcuts.moveLine" },
  { id: "copyLineDown", group: "edit", keys: [["Shift", "Alt", "↓"]], labelKey: "code.shortcuts.copyLine" },

  // --- Code intelligence --------------------------------------------------
  { id: "goToDefinition", group: "intelligence", keys: [["F12"]], labelKey: "code.palette.cmd.goToDefinition" },
  { id: "peekDefinition", group: "intelligence", keys: [["Alt", "F12"]], labelKey: "code.palette.cmd.peekDefinition" },
  { id: "findReferences", group: "intelligence", keys: [["Shift", "F12"]], labelKey: "code.palette.cmd.findReferences" },
  { id: "peekReferences", group: "intelligence", keys: [["Mod", "Alt", "F12"]], labelKey: "code.palette.cmd.peekReferences" },
  { id: "rename", group: "intelligence", keys: [["F2"]], labelKey: "code.palette.cmd.rename" },
  { id: "quickFix", group: "intelligence", keys: [["Mod", "."]], labelKey: "code.palette.cmd.quickFix" },
  { id: "refactor", group: "intelligence", keys: [["Mod", "Shift", "R"]], labelKey: "code.palette.cmd.refactor" },

  // --- Run and debug ------------------------------------------------------
  { id: "run", group: "run", keys: [["Mod", "Enter"]], labelKey: "code.run" },
  { id: "console", group: "run", hostBound: true, keys: [["Mod", "`"]], labelKey: "code.console.title" },
  { id: "output", group: "run", hostBound: true, keys: [["Mod", "J"]], labelKey: "codeOutput.showOutput" },
  { id: "debugStart", group: "run", hostBound: true, keys: [["F5"]], labelKey: "code.debugRun" },
  { id: "debugRestart", group: "run", hostBound: true, keys: [["Mod", "Shift", "F5"]], labelKey: "code.debugRestart" },
  { id: "debugStop", group: "run", hostBound: true, keys: [["Shift", "F5"]], labelKey: "code.debugStop" },
  { id: "debugStepOver", group: "run", hostBound: true, keys: [["F10"]], labelKey: "code.debugStepOver" },
  { id: "debugStepInto", group: "run", hostBound: true, keys: [["F11"]], labelKey: "code.debugStepInto" },

  // --- View ---------------------------------------------------------------
  { id: "compare", group: "view", keys: [["Mod", "Shift", "\\"]], labelKey: "code.compare.title" },
  { id: "wordWrap", group: "view", keys: [["Alt", "Z"]], labelKey: "code.palette.cmd.toggleWordWrap" },
  { id: "zoomIn", group: "view", keys: [["Mod", "="]], labelKey: "code.palette.cmd.zoomIn" },
  { id: "zoomOut", group: "view", keys: [["Mod", "-"]], labelKey: "code.palette.cmd.zoomOut" },
  { id: "zoomReset", group: "view", keys: [["Mod", "0"]], labelKey: "code.palette.cmd.zoomReset" },
];

/** Group order for display; the picker lists them in this order. */
export const SHORTCUT_GROUP_ORDER: readonly ShortcutGroup[] = [
  "navigate",
  "edit",
  "intelligence",
  "run",
  "view",
];

export const SHORTCUT_GROUP_LABEL_KEYS: Readonly<
  Record<ShortcutGroup, string>
> = {
  navigate: "code.shortcuts.groupNavigate",
  edit: "code.shortcuts.groupEdit",
  intelligence: "code.shortcuts.groupIntelligence",
  run: "code.shortcuts.groupRun",
  view: "code.shortcuts.groupView",
};

/** Render one chord for display, using the platform's own modifier names. */
export function formatShortcutKeys(
  keys: readonly (readonly string[])[],
  isMac: boolean,
): string {
  const token = (key: string) => {
    if (key === "Mod") return isMac ? "⌘" : "Ctrl";
    // The Control key itself, also on a Mac (Go to Line is ⌃G there).
    if (key === "Ctrl") return isMac ? "⌃" : "Ctrl";
    if (key === "Alt") return isMac ? "⌥" : "Alt";
    if (key === "Shift") return isMac ? "⇧" : "Shift";
    return key;
  };
  // macOS writes modifiers as adjacent glyphs; elsewhere they are joined by +.
  return keys
    .map((step) => step.map(token).join(isMac ? "" : "+"))
    .join(" ");
}

/**
 * Filter the table by a free-text query against the chord and the LABEL the
 * caller resolved (labels are localized, so matching happens on rendered text).
 */
/**
 * The rows worth showing in this host: everything the editor binds itself, plus
 * the host-bound ids the caller says it actually binds.
 */
export function shortcutsForHost(
  shortcuts: readonly EditorShortcut[],
  hostBoundIds: readonly string[] = [],
): EditorShortcut[] {
  const available = new Set(hostBoundIds);
  return shortcuts.filter(
    (shortcut) => !shortcut.hostBound || available.has(shortcut.id),
  );
}

export function filterShortcuts(
  shortcuts: readonly EditorShortcut[],
  query: string,
  labelFor: (shortcut: EditorShortcut) => string,
  isMac: boolean,
): EditorShortcut[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [...shortcuts];
  return shortcuts.filter((shortcut) => {
    const chord = formatShortcutKeys(shortcut.keys, isMac).toLowerCase();
    return (
      chord.includes(needle) ||
      labelFor(shortcut).toLowerCase().includes(needle) ||
      shortcut.id.toLowerCase().includes(needle)
    );
  });
}
