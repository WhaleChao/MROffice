/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The stored original of an imported file: download it as it is, or hand it to
 * an app on this computer through the desktop app.
 *
 * The original is a DataItem; its signed download link streams the stored
 * bytes (with ranges), so a file of any size downloads without passing
 * through the page or the export pipeline.
 */

import { apiGet } from "../api";
import { downloadUrlFile } from "../download-file";
import { canDownloadInParallel, downloadInParallel } from "../parallel-download";
import { desktopBridge } from "../open-files/desktop";

export interface StoredOriginal {
  readonly dataItemId: string;
  readonly filename: string;
  readonly sizeBytes: number;
}

async function signedOriginalUrl(original: StoredOriginal): Promise<string> {
  const item = await apiGet<{ download_url?: string | null }>(
    `/api/v1/data-items/${encodeURIComponent(original.dataItemId)}`,
  );
  if (!item.download_url) throw new Error("original_download_unavailable");
  return new URL(item.download_url, globalThis.location?.href).href;
}

export async function downloadStoredOriginal(original: StoredOriginal): Promise<void> {
  // A large original where the browser can write a picked file: parallel
  // ranges, several times faster over a long path, with progress shown.
  if (canDownloadInParallel(original.sizeBytes)) {
    await downloadInParallel({
      filename: original.filename,
      sizeBytes: original.sizeBytes,
      resolveUrl: () => signedOriginalUrl(original),
    });
    return;
  }
  const url = await signedOriginalUrl(original);
  if (!downloadUrlFile(original.filename, url)) {
    throw new Error("original_download_unavailable");
  }
}

/** Whether this window can open a file in an app on this computer. */
export function canOpenOriginalInDesktopApp(): boolean {
  return typeof desktopBridge()?.openFileInApp === "function";
}

/**
 * Download the original into the desktop app's own folder and open it with an
 * app the person picks (the desktop app is itself the default for these
 * files, so the system's default would open it here again).
 */
export async function openStoredOriginalInDesktopApp(
  original: StoredOriginal,
): Promise<"opened" | "cancelled" | "revealed"> {
  const bridge = desktopBridge();
  if (!bridge?.openFileInApp) throw new Error("desktop_open_unavailable");
  const url = await signedOriginalUrl(original);
  return bridge.openFileInApp({
    url,
    filename: original.filename,
    sizeBytes: original.sizeBytes,
  });
}
