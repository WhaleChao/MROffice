/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The desktop app's own settings, kept in its data folder
 * (desktop-settings.json): what the person chose in the app's menu.
 */
import fs from "node:fs";
import path from "node:path";
import { app } from "electron";

export interface DesktopSettings {
  /** Windows: closing the last window leaves the app in the tray (./tray.ts). */
  keepInTrayOnClose?: boolean;
  /** "Ready in the background at login", as the person set it in the menu (./prelaunch.ts); absent: never set. */
  prelaunch?: boolean;
  /** A page that came from a file on this computer is written back to that
   * file once its edits settle, not only on ⌘S (Settings › General);
   * absent: on. */
  autoSaveToOriginal?: boolean;
}

const settingsFile = () => path.join(app.getPath("userData"), "desktop-settings.json");

export function readSettings(): DesktopSettings {
  try {
    const value = JSON.parse(fs.readFileSync(settingsFile(), "utf8")) as unknown;
    return value && typeof value === "object" ? (value as DesktopSettings) : {};
  } catch {
    return {};
  }
}

export function writeSettings(change: DesktopSettings): void {
  const partial = `${settingsFile()}.partial`;
  fs.writeFileSync(partial, JSON.stringify({ ...readSettings(), ...change }));
  fs.renameSync(partial, settingsFile());
}

/** Autosave to the original file (on unless the person turned it off). */
export function autoSaveToOriginalEnabled(): boolean {
  return readSettings().autoSaveToOriginal !== false;
}
