/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The slide editor's desktop PowerPoint chords that act on the deck rather
 * than on the words being typed: New Slide, Duplicate Slide, Group / Ungroup
 * and Page Down / Page Up through the slides (audit PPT-E001 E002 E004 E007,
 * C015).
 *
 * Microsoft Support, "Keyboard shortcuts in PowerPoint"
 * (support.microsoft.com/office/ebb3d20e-dcd4-444f-a38e-bb5c5ed180f4):
 *
 *   | Command          | Windows        | macOS           |
 *   | New slide        | Ctrl+M         | Shift+Cmd+N     |
 *   | Duplicate slide  | Ctrl+Shift+D   | Shift+Cmd+D     |
 *   | Group            | Ctrl+G         | Option+Cmd+G    |
 *   | Ungroup          | Ctrl+Shift+G   | Option+Shift+Cmd+G |
 *   | Regroup          | Ctrl+Shift+J   | Option+Cmd+J    |
 *   | Hyperlink        | Ctrl+K         | Cmd+K           |
 *   | Next / previous slide (Normal view) | Page Down / Page Up |
 *
 * Browser trade-offs (PowerPoint for the web's own list says the web app's
 * keys differ "because PowerPoint for the web runs in your web browser"):
 *  - Shift+Cmd+N opens a private window in Chrome and Safari before the page
 *    sees it; PowerPoint for the web lists Ctrl+M for a new slide on Windows
 *    and none on a Mac. The deck takes Shift+Cmd+N whenever the browser does
 *    hand it over, and Control+M on a Mac as the key that always arrives; the
 *    menus print the key that works (Ctrl+M, or ⌃+M on a Mac).
 *  - Regroup's Ctrl+Shift+J / Option+Cmd+J open the browser's developer
 *    console, so Regroup has no chord here.
 *  - Hyperlink on a Mac is Cmd+K only: Microsoft's macOS list and PowerPoint
 *    for the web's Mac column give no Control+K, and in a Mac text field
 *    Control+K is the system's "delete to the end of the paragraph" (Cocoa
 *    text key bindings), which the slide's text boxes keep.
 *
 * Pure: no DOM, no React, so the keymap is tested on its own.
 */

import { isMacPlatform } from "@mrnote/code/web/lib/code-runner/mac-option-chords.ts";
import { formatOfficeShortcut } from "@mrnote/core/web/lib/office-context-menu.ts";

export type PresentationDesktopShortcut =
  | "newSlide"
  | "duplicateSlide"
  | "group"
  | "ungroup"
  | "nextSlide"
  | "previousSlide"
  | "hyperlink";

export interface PresentationShortcutKeyEvent {
  readonly key: string;
  readonly code?: string;
  readonly shiftKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
  readonly altKey?: boolean;
}

/** The chord each command prints in a menu, per platform. */
const LABEL_SPECS: Readonly<
  Record<"newSlide" | "duplicateSlide" | "group" | "ungroup", { mac: string; other: string }>
> = {
  newSlide: { mac: "Ctrl+M", other: "Ctrl+M" },
  duplicateSlide: { mac: "Mod+Shift+D", other: "Ctrl+Shift+D" },
  group: { mac: "Mod+Alt+G", other: "Ctrl+G" },
  ungroup: { mac: "Mod+Alt+Shift+G", other: "Ctrl+Shift+G" },
};

/** "⇧⌘D" style text on a Mac ("⌘+⇧+D" in the Office web grammar), "Ctrl+Shift+D" elsewhere. */
export function presentationDesktopShortcutLabel(
  name: keyof typeof LABEL_SPECS,
  mac = isMacPlatform(),
): string {
  const spec = LABEL_SPECS[name];
  return formatOfficeShortcut(mac ? spec.mac : spec.other, mac);
}

/** The letter a chord names, read from the physical key when Option rewrote `key` (⌥G types "©"). */
function letter(event: PresentationShortcutKeyEvent): string {
  if (typeof event.code === "string" && /^Key[A-Z]$/.test(event.code)) {
    return event.code.slice(3).toLowerCase();
  }
  return event.key.length === 1 ? event.key.toLowerCase() : "";
}

/**
 * Which deck command a key press is, or null. `mac` picks the platform's
 * modifiers (Cmd on a Mac, Ctrl elsewhere).
 */
export function presentationDesktopShortcutFor(
  event: PresentationShortcutKeyEvent,
  mac = isMacPlatform(),
): PresentationDesktopShortcut | null {
  const ctrl = Boolean(event.ctrlKey);
  const meta = Boolean(event.metaKey);
  const shift = Boolean(event.shiftKey);
  const alt = Boolean(event.altKey);
  if (!ctrl && !meta && !alt && !shift) {
    if (event.key === "PageDown") return "nextSlide";
    if (event.key === "PageUp") return "previousSlide";
    return null;
  }
  const key = letter(event);
  if (mac) {
    if (meta && !ctrl) {
      if (!shift && !alt && key === "k") return "hyperlink";
      if (shift && !alt && key === "n") return "newSlide";
      if (shift && !alt && key === "d") return "duplicateSlide";
      if (alt && !shift && key === "g") return "group";
      if (alt && shift && key === "g") return "ungroup";
      return null;
    }
    if (ctrl && !meta && !alt && !shift && key === "m") return "newSlide";
    return null;
  }
  if (!ctrl || meta || alt) return null;
  if (!shift && key === "m") return "newSlide";
  if (!shift && key === "k") return "hyperlink";
  if (shift && key === "d") return "duplicateSlide";
  if (!shift && key === "g") return "group";
  if (shift && key === "g") return "ungroup";
  return null;
}
