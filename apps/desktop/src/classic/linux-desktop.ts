/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Linux: the app among the person's applications, in Open With for its
 * documents, and the handler of its sign-in links (../linux-desktop-entry.ts).
 *
 * - The .deb installs /usr/share/applications/<package>.desktop itself
 *   (electron-builder, desktop entry from electron-builder.classic.config.cjs),
 *   with its icons; the package's scripts refresh the system's caches.
 * - An AppImage is a file the person downloaded and made executable: at each
 *   launch the app writes ~/.local/share/applications/<package>.desktop for
 *   it (Exec= the AppImage's own path, so moving or updating the file is
 *   followed at the next launch) and copies the AppImage's icons into
 *   ~/.local/share/icons/hicolor. A file it wrote is rewritten only when it
 *   changed; one that is not its own (the person's, a launcher's such as
 *   AppImageLauncher) is left alone.
 *
 * Either way the sign-in link scheme (<protocol>://) is made this app's when
 * nothing else claims it, so the browser's "Open MR Office?" reaches it.
 * No document type is made the app's default here (src/default-app.ts asks
 * the person, as on the other systems).
 *
 * MRAI_LINUX_DESKTOP_INTEGRATION=0 leaves everything as it is.
 */

import { execFile } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { app } from "electron";

import { APPIMAGE_ENTRY_MARK, desktopEntryText } from "../linux-desktop-entry";
import { PRODUCT } from "./product";

/** The app's desktop entry, as the package names it (package.json desktopName). */
export function linuxDesktopId(): string {
  return `${PRODUCT.packageName}.desktop`;
}

/** The documents Open With offers the app for (document-types.cjs, the same list). */
function documentMimeTypes(): string[] {
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const types = require(path.join(__dirname, "..", "..", "document-types.cjs")) as { LINUX_DOCUMENT_MIME_TYPES?: readonly string[] };
    return [...(types.LINUX_DOCUMENT_MIME_TYPES ?? [])];
  } catch {
    return [];
  }
}

function run(command: string, args: string[]): Promise<string> {
  return new Promise((resolve) => {
    execFile(command, args, { timeout: 10_000 }, (error, stdout) => resolve(error ? "" : String(stdout).trim()));
  });
}

function dataHome(): string {
  const configured = process.env.XDG_DATA_HOME;
  return configured && path.isAbsolute(configured) ? configured : path.join(os.homedir(), ".local", "share");
}

/** Write `text` to `file` unless it already says that; true when written. */
function writeIfChanged(file: string, text: string | Buffer): boolean {
  try {
    const current = fs.readFileSync(file);
    if (current.equals(Buffer.isBuffer(text) ? text : Buffer.from(text))) return false;
  } catch {
    // Not there yet.
  }
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const temporary = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(temporary, text, { mode: 0o644 });
  fs.renameSync(temporary, file);
  return true;
}

/** The AppImage's icons (usr/share/icons/hicolor/<size>/apps/<package>.png) into the person's icon theme. */
function copyIcons(appDir: string, iconName: string): boolean {
  const from = path.join(appDir, "usr", "share", "icons", "hicolor");
  let changed = false;
  let sizes: string[] = [];
  try {
    sizes = fs.readdirSync(from);
  } catch {
    return false;
  }
  for (const size of sizes) {
    const source = path.join(from, size, "apps", `${iconName}.png`);
    if (!fs.existsSync(source)) continue;
    try {
      changed = writeIfChanged(path.join(dataHome(), "icons", "hicolor", size, "apps", `${iconName}.png`), fs.readFileSync(source)) || changed;
    } catch {
      // An icon the theme cannot take is not worth failing the launch for.
    }
  }
  return changed;
}

async function integrateAppImage(appImage: string): Promise<void> {
  const id = linuxDesktopId();
  const file = path.join(dataHome(), "applications", id);
  try {
    const existing = fs.readFileSync(file, "utf8");
    if (!existing.includes(APPIMAGE_ENTRY_MARK)) {
      console.log(`[linux] ${file} is not this app's own; left as it is`);
      return;
    }
  } catch {
    // Not there yet: written below.
  }
  const text = desktopEntryText({
    name: PRODUCT.productName,
    comment: `${PRODUCT.displayName.en}: Word, Excel, PowerPoint and PDF documents on this computer.`,
    executable: appImage,
    icon: PRODUCT.packageName,
    wmClass: PRODUCT.packageName,
    mimeTypes: documentMimeTypes(),
    protocol: PRODUCT.protocol,
    version: app.getVersion(),
  });
  // The AppImage's runtime names its mount (APPDIR); the app's resources
  // folder is in it either way ($APPDIR/resources).
  const appDir = process.env.APPDIR && path.isAbsolute(process.env.APPDIR) ? process.env.APPDIR : path.dirname(process.resourcesPath);
  const icons = copyIcons(appDir, PRODUCT.packageName);
  const written = writeIfChanged(file, text);
  if (written) {
    console.log(`[linux] desktop entry ${file} → ${appImage}`);
    await run("update-desktop-database", [path.dirname(file)]);
  }
  if (icons) await run("gtk-update-icon-cache", ["--ignore-theme-index", "--quiet", path.join(dataHome(), "icons", "hicolor")]);
}

/** The sign-in links open this app, unless the person chose another for them. */
async function claimLinkScheme(): Promise<void> {
  const type = `x-scheme-handler/${PRODUCT.protocol}`;
  const current = await run("xdg-mime", ["query", "default", type]);
  if (current === linuxDesktopId()) return;
  // Another app named for our own scheme (a copy of an older build under
  // another name, say) is replaced; nothing else ever claims it.
  await run("xdg-mime", ["default", linuxDesktopId(), type]);
  console.log(`[linux] ${type} → ${linuxDesktopId()}${current ? ` (was ${current})` : ""}`);
}

export async function integrateLinuxDesktop(): Promise<void> {
  if (process.platform !== "linux" || !app.isPackaged || process.env.MRAI_LINUX_DESKTOP_INTEGRATION === "0") return;
  const appImage = process.env.APPIMAGE;
  try {
    if (appImage && path.isAbsolute(appImage) && fs.existsSync(appImage)) await integrateAppImage(appImage);
    await claimLinkScheme();
  } catch (error) {
    console.warn(`[linux] desktop integration: ${error instanceof Error ? error.message : error}`);
  }
}
