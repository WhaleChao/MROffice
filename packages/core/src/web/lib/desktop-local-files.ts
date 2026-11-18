/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Files on this computer, in MRNote Classic (docs/desktop-classic.md).
 *
 * The app's API runs on the computer the files are on, so an import sends a
 * token for the file instead of its bytes and the API reads it where it is
 * (apps/api app/desktop/local_files.py): a part of the import batch whose
 * type is {@link DESKTOP_LOCAL_FILE_TYPE}, or a large file's parts filled
 * from disk. The page a file became is linked to it, so Save writes the page
 * back to that file (apps/desktop src/classic/local-files.ts), and opening
 * the file again, unchanged, opens that page instead of importing a copy.
 *
 * Everywhere but the desktop app every function here answers "no local
 * file", and the bytes travel as they always have.
 */

import { IS_DESKTOP_CLASSIC } from "./desktop-classic";

export const DESKTOP_LOCAL_FILE_TYPE = "application/x-mrai-local-file";

interface LocalFilesBridge {
  localFileToken?(file: File): Promise<string | null>;
  linkOriginal?(pageId: string, localToken: string, notebookId?: string): Promise<boolean>;
  originalOf?(pageId: string): Promise<DesktopOriginal | null>;
  saveToOriginal?(request: {
    pageId: string;
    bytes: ArrayBuffer;
    filename: string;
    state?: string;
  }): Promise<{ path: string; name: string } | { error: string } | null>;
}

function bridge(): LocalFilesBridge | null {
  if (!IS_DESKTOP_CLASSIC || typeof window === "undefined") return null;
  return (window as unknown as { mraiDesktop?: LocalFilesBridge }).mraiDesktop ?? null;
}

/** The file a page came from: its name, and the page body the file holds
 * while it is unchanged on disk (its content version and causal revision;
 * null when unknown). */
export interface DesktopOriginal {
  readonly name: string;
  readonly savedState?: string | null;
}

const tokens = new WeakMap<File, string>();
const existingPages = new WeakMap<File, DesktopExistingPage>();

/** The page a file the app opened already became (the file unchanged since). */
export interface DesktopExistingPage {
  readonly pageId: string;
  readonly notebookId: string;
}

export function markDesktopExistingPage(file: File, page: DesktopExistingPage | undefined): File {
  if (page && IS_DESKTOP_CLASSIC) existingPages.set(file, page);
  return file;
}

/** Opening this file again opens this page instead of importing a copy. */
export function desktopExistingPage(file: File): DesktopExistingPage | null {
  return existingPages.get(file) ?? null;
}

/** ``file`` is the file the app registered as ``token`` (opened from the system). */
export function markDesktopLocalFile(file: File, token: string | undefined): File {
  if (token && IS_DESKTOP_CLASSIC) tokens.set(file, token);
  return file;
}

/** The token ``file`` was marked with (``markDesktopLocalFile``), or null. */
export function markedDesktopLocalFileToken(file: File): string | null {
  return tokens.get(file) ?? null;
}

/** Whether ``file`` was opened from the system in the desktop app (its API
 * reads it where it is), known at once. */
export function isDesktopLocalFile(file: File): boolean {
  return tokens.has(file);
}

/** The token of a file on this computer, or null (not the desktop app, or a
 * file made in the page). */
export async function desktopLocalFileToken(file: File): Promise<string | null> {
  const known = tokens.get(file);
  if (known) return known;
  const ask = bridge()?.localFileToken;
  if (!ask) return null;
  try {
    const token = await ask(file);
    if (typeof token === "string" && token) {
      tokens.set(file, token);
      return token;
    }
  } catch {
    // Sent as bytes.
  }
  return null;
}

/** What an import batch sends for ``file``: its token when the API can read
 * it where it is, else the file. */
/**
 * Whether an upload of ``file`` sends its bytes from the page: false for a
 * file on this computer in MR Note Classic, which the request only names and
 * the app's API reads where it is ({@link desktopLocalUploadPart}). An
 * upload of it never shows progress; what is waited for is the server's work.
 */
export async function uploadSendsFileBytes(file: File): Promise<boolean> {
  if (!IS_DESKTOP_CLASSIC) return true;
  return (await desktopLocalFileToken(file)) === null;
}

export async function desktopLocalUploadPart(file: File): Promise<File> {
  const token = await desktopLocalFileToken(file);
  if (!token) return file;
  return new File([JSON.stringify({ token, type: file.type })], file.name, {
    type: DESKTOP_LOCAL_FILE_TYPE,
    lastModified: file.lastModified,
  });
}

/** Link the page ``file`` became to it: Save writes back there, and opening
 * the file again, unchanged, opens this page. */
export function linkDesktopOriginal(pageId: string, file: File, notebookId: string): void {
  const token = tokens.get(file);
  const link = bridge()?.linkOriginal;
  if (!token || !link || !pageId) return;
  void link(pageId, token, notebookId).catch(() => undefined);
}

export function desktopSaveToOriginalAvailable(): boolean {
  return Boolean(bridge()?.saveToOriginal);
}

export async function desktopOriginalOf(pageId: string): Promise<DesktopOriginal | null> {
  const ask = bridge()?.originalOf;
  if (!ask) return null;
  try {
    return await ask(pageId);
  } catch {
    return null;
  }
}

/** Write ``file`` over the file the page came from (asking where once when
 * there is none); the saved file's name, or null when the person cancelled. */
export async function saveDesktopOriginal(
  pageId: string,
  file: Blob,
  filename: string,
  state?: string,
): Promise<{ name: string } | null> {
  const save = bridge()?.saveToOriginal;
  if (!save) return null;
  const answer = await save({ pageId, bytes: await file.arrayBuffer(), filename, state });
  if (!answer) return null;
  if ("error" in answer) throw new Error(answer.error);
  return { name: answer.name };
}
