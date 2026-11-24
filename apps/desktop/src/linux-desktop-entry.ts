/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The desktop entry an AppImage installs for the person (freedesktop.org
 * Desktop Entry Specification 1.5): what makes the app appear among the
 * applications, in a file manager's Open With for its documents, and the
 * handler of its sign-in links (x-scheme-handler/<protocol>). A .deb brings
 * its own (/usr/share/applications, electron-builder); an AppImage is a file
 * the person put somewhere, so the app writes one under
 * ~/.local/share/applications pointing at that file, and refreshes it when
 * the file moves or is replaced by an update (./classic/linux-desktop.ts).
 *
 * Pure text, no Electron: test/linux-desktop.test.mjs.
 */

import { fileURLToPath } from "node:url";

export interface LinuxDesktopEntry {
  /** The name people read (Name=). */
  readonly name: string;
  readonly comment: string;
  /** The AppImage's absolute path. */
  readonly executable: string;
  /** The icon's name in the icon theme (the package's name). */
  readonly icon: string;
  /** Electron's window class (package.json desktopName without .desktop). */
  readonly wmClass: string;
  readonly mimeTypes: readonly string[];
  /** The sign-in link scheme, without "://". */
  readonly protocol: string;
  readonly version: string;
}

/** Marks an entry this app wrote (never one the person or a package wrote). */
export const APPIMAGE_ENTRY_MARK = "X-MRAI-AppImage=true";

/**
 * One argument of Exec=, quoted as the specification asks: inside double
 * quotes, a backslash before each of " ` $ \, and every % doubled (a field
 * code otherwise). A desktop entry's string value then escapes each
 * backslash once more.
 */
export function quoteExecArgument(argument: string): string {
  const quoted = `"${argument.replace(/["`$\\]/g, (character) => `\\${character}`).replace(/%/g, "%%")}"`;
  return quoted.replace(/\\/g, "\\\\");
}

/** A string value: no line breaks (a newline would end the entry's key). */
function value(text: string): string {
  return text.replace(/\\/g, "\\\\").replace(/\r?\n|\r/g, " ").replace(/\t/g, " ");
}

export function desktopEntryText(entry: LinuxDesktopEntry): string {
  const types = [...new Set([...entry.mimeTypes, `x-scheme-handler/${entry.protocol}`])];
  return [
    "[Desktop Entry]",
    "Type=Application",
    `Name=${value(entry.name)}`,
    `Comment=${value(entry.comment)}`,
    `Exec=${quoteExecArgument(entry.executable)} %U`,
    `TryExec=${value(entry.executable)}`,
    `Icon=${value(entry.icon)}`,
    "Terminal=false",
    "Categories=Office;",
    `MimeType=${types.join(";")};`,
    `StartupWMClass=${value(entry.wmClass)}`,
    `X-AppImage-Version=${value(entry.version)}`,
    APPIMAGE_ENTRY_MARK,
    "",
  ].join("\n");
}

/**
 * The documents among a process's arguments as a file manager hands them
 * through %U: a path, or a file:// URL (GNOME and KDE pass URLs). Anything
 * else (a link, an option) is left out here.
 */
export function pathFromFileArgument(argument: string): string | null {
  if (!argument.startsWith("file://")) return null;
  try {
    // Throws for a file on another host (file://server/...).
    return fileURLToPath(argument);
  } catch {
    return null;
  }
}
