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
import { lazyFileReader } from "../document-import/lazy-file";
import { desktopExistingPage, markDesktopExistingPage, markDesktopLocalFile, markedDesktopLocalFileToken } from "../desktop-local-files";

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

/** A page a document opened from the system already became (unchanged since). */
export interface DesktopReopenedPage {
  readonly pageId: string;
  readonly notebookId: string;
}

/** What ``DesktopBridge.onDocumentsOpened``'s listener is told. */
export interface DesktopDocumentsOpened {
  /** Documents opened again, each unchanged since it became a page: the page
   * shows these pages where it is, and takes no file (``reopens``). Absent:
   * new documents wait, taken with ``takeOpenFiles``. */
  readonly reopened?: readonly DesktopReopenedPage[];
}

export interface DesktopDocumentsOpenedOptions {
  /** The page shows a document opened again where it is (the pages it
   * became, ``DesktopDocumentsOpened.reopened``) instead of the window going
   * to the open route and loading the notebook again. */
  readonly reopens?: boolean;
  /** The notebook it shows: only pages of it are handed over (``reopens``). */
  readonly notebookId?: string;
}

export interface DesktopBridge {
  readonly platform: string;
  /** ``keep`` (MRNote Classic): look without taking them. */
  takeOpenFiles(keep?: boolean): Promise<DesktopOpenedFile[]>;
  /**
   * MR Office / Classic: a notebook's page takes the documents opened from
   * the system while it is shown, where it is (``listener`` hears when some
   * wait; take them with ``takeOpenFiles``). Returns the unsubscribe.
   * Absent in older builds, where the window loads the open route again.
   */
  onDocumentsOpened?(listener: (opened?: DesktopDocumentsOpened) => void, options?: DesktopDocumentsOpenedOptions): () => void;
  /**
   * MR Office / Classic: pages handed to ``onDocumentsOpened``'s listener as
   * documents opened again (``DesktopDocumentsOpened.reopened``) that are
   * gone: the app hands their documents over again as new ones, imported as
   * any other. Absent in builds that never hand pages over.
   */
  reopenedPagesMissing?(pageIds: readonly string[]): Promise<boolean>;
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
/** Documents opened again from the system are shown as the pages they became,
 * in the notebook shown (lib/open-files/desktop-in-place
 * `showReopenedPagesInPlace`): as for files handed over, MR Office's home
 * steps back so their windows come forward. */
export const DESKTOP_PAGES_REOPENED_EVENT = "mrnote:desktop-pages-reopened";

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

/** MR Office / Classic: hear documents opened from the system while this
 * page is shown (``DesktopBridge.onDocumentsOpened``); the unsubscribe, or
 * null where the app cannot hand them to the page. */
/**
 * The documents a launch hands over, as first looked at (``takeDesktopOpenFiles``
 * with ``keep``, before the API answered) when they are the same files the
 * notebook then takes: the looked-at copies, given what the take adds (the
 * API's token for each, the page an unchanged one already became). Work the
 * import does once per File (its container check) started on them while the
 * API was starting is then not done again. Any difference -- another count,
 * name, size, type or date, or a large file read in parts -- keeps the taken
 * files as they are.
 */
export function adoptLookedAtOpenFiles(lookedAt: readonly File[], taken: File[]): File[] {
  if (lookedAt.length === 0 || lookedAt.length !== taken.length) return taken;
  for (let index = 0; index < taken.length; index += 1) {
    const before = lookedAt[index];
    const after = taken[index];
    if (before.name !== after.name || before.size !== after.size || before.type !== after.type ||
        before.lastModified !== after.lastModified || lazyFileReader(before) || lazyFileReader(after)) return taken;
  }
  return lookedAt.map((file, index) => {
    const after = taken[index];
    markDesktopLocalFile(file, markedDesktopLocalFileToken(after) ?? undefined);
    markDesktopExistingPage(file, desktopExistingPage(after) ?? undefined);
    return file;
  });
}

export function onDesktopDocumentsOpened(
  listener: (opened?: DesktopDocumentsOpened) => void,
  options?: DesktopDocumentsOpenedOptions,
): (() => void) | null {
  const subscribe = desktopBridge()?.onDocumentsOpened;
  return typeof subscribe === "function" ? subscribe(listener, options) : null;
}

/** MR Office / Classic: these pages, handed over as documents opened again,
 * are gone; the app hands their documents over again as new ones. False
 * where it cannot (an older build). */
export async function desktopReopenedPagesMissing(pageIds: readonly string[]): Promise<boolean> {
  const missing = desktopBridge()?.reopenedPagesMissing;
  if (typeof missing !== "function") return false;
  try {
    return (await missing(pageIds)) === true;
  } catch {
    return false;
  }
}
