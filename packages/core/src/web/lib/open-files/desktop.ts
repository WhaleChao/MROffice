/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The desktop app's bridge (apps/desktop/src/preload.ts), when the site runs
 * inside MRNote for macOS or Windows: the documents a window was
 * opened for, and making the app the default for documents. Absent in a
 * browser; every caller treats it as optional.
 */

import { lazyFile } from "../document-import/lazy-file";
import { markDesktopExistingPage, markDesktopLocalFile } from "../desktop-local-files";

export interface DesktopDefaultAppState {
  /** "direct": the app sets itself, macOS confirming each type; "settings": the system's page opens (Windows). */
  readonly supported: "direct" | "settings" | "none";
  /** Whether the system answered which app opens each type (absent in older builds). */
  readonly checked?: boolean;
  /** Extensions (without the dot) that already open with the app. */
  readonly defaultFor: readonly string[];
  readonly extensions: readonly string[];
}

export interface DesktopDefaultAppOffer {
  readonly supported: "direct" | "settings";
  /** The offered types that open with another app now. */
  readonly missing: readonly string[];
}

interface DesktopOpenedFile {
  readonly name: string;
  readonly type: string;
  readonly lastModified: number;
  /** Absent from desktop builds older than part reads. */
  readonly size?: number;
  /** The bytes of a file read whole. */
  readonly data?: Uint8Array;
  /** A large file's token, read part by part with `readOpenFile`. */
  readonly token?: string;
  /** MRNote Classic: the token its API reads the file by, where it is. */
  readonly localToken?: string;
  /** MRNote Classic: the page this file already became, unchanged since. */
  readonly existingPage?: { readonly pageId: string; readonly notebookId: string };
}

export interface DesktopBridge {
  readonly platform: string;
  /** ``keep`` (MRNote Classic): look without taking them. */
  takeOpenFiles(keep?: boolean): Promise<DesktopOpenedFile[]>;
  /** Bytes start..end of a large opened file (absent in older builds). */
  readOpenFile?(token: string, start: number, end: number): Promise<Uint8Array | null>;
  /**
   * Download a stored file into the app's folder and open it with an app the
   * person chooses. Absent in desktop builds older than this bridge.
   */
  openFileInApp?(request: {
    url: string;
    filename: string;
    sizeBytes: number;
  }): Promise<"opened" | "cancelled" | "revealed">;
  defaultAppState(): Promise<DesktopDefaultAppState | null>;
  /** Make the app the default for these extensions (all offered ones when none). */
  makeDefaultApp(extensions?: string[]): Promise<DesktopDefaultAppState | null>;
  /**
   * MRNote Classic, at launch: what to ask about when one of the offered
   * types opens with another app and the person has not said "No thanks";
   * at most once per run of the app, else null. Absent elsewhere.
   */
  defaultAppOffer?(): Promise<DesktopDefaultAppOffer | null>;
  /** MRNote Classic: "No thanks", remembered; launches no longer ask. */
  declineDefaultApp?(): Promise<boolean>;
}

/** MRNote Classic: a notebook opened with this query imports the documents
 * its window was opened for (`takeDesktopOpenFiles`). */
export const DESKTOP_OPEN_FILES_QUERY = "openFiles=desktop";

/** Dispatched with those documents (``detail: File[]``) once the notebook's
 * windows are restored; its page list imports them as a drop. */
export const DESKTOP_OPEN_FILES_EVENT = "mrnote:desktop-open-files";

export function desktopBridge(): DesktopBridge | null {
  if (typeof window === "undefined") return null;
  const bridge = (window as unknown as { mraiDesktop?: DesktopBridge }).mraiDesktop;
  return bridge && typeof bridge.takeOpenFiles === "function" ? bridge : null;
}

/** The documents this desktop window was opened for, as files; once
 * (``keep``: MRNote Classic looks, and a later call takes them). */
export async function takeDesktopOpenFiles(options: { keep?: boolean } = {}): Promise<File[]> {
  const bridge = desktopBridge();
  if (!bridge) return [];
  try {
    const opened = await bridge.takeOpenFiles(options.keep === true);
    return opened.flatMap((file): File[] => {
      if (file.data) {
        return [
          markDesktopExistingPage(
            markDesktopLocalFile(
              new File([new Uint8Array(file.data)], file.name, {
                type: file.type,
                lastModified: file.lastModified,
              }),
              file.localToken,
            ),
            file.existingPage,
          ),
        ];
      }
      const token = file.token;
      const read = bridge.readOpenFile;
      if (!token || !read || typeof file.size !== "number") return [];
      return [
        markDesktopLocalFile(
          lazyFile(
            { name: file.name, type: file.type, size: file.size, lastModified: file.lastModified },
            async (start, end) => {
              const bytes = await read(token, start, end);
              if (!bytes) throw new Error("desktop_file_unreadable");
              return new Blob([new Uint8Array(bytes)]);
            },
          ),
          file.localToken,
        ),
      ].map((opened) => markDesktopExistingPage(opened, file.existingPage));
    });
  } catch {
    return [];
  }
}
