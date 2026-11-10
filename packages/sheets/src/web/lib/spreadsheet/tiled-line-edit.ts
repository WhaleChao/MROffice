/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whole rows or columns of a tiled sheet too large to hold in the browser. The
 * cells are moved where they are stored (`POST .../workbook-tiles/line-edit`:
 * tiles, the body, every formula and the file together), so the window saves
 * what it holds first and reads the workbook again afterwards.
 *
 * A reload clears the window's Undo history, so the step that takes the edit
 * back is kept for this tab beside it: Undo takes it once nothing newer is left
 * to undo. The server refuses it if the page has moved on since.
 */

import { apiGet, apiPost, isApiRequestError } from "@mrnote/core/web/lib/api.ts";

export interface TiledLineEditRequest {
  sheet: string;
  axis: "row" | "col";
  /** The first line, 0-based: new lines go in before it, deleted ones start at it. */
  at: number;
  /** > 0 inserted; < 0 deleted. */
  delta: number;
  base_content_version: number;
}

export interface TiledLineEditUndoStep {
  revision: number;
  contentVersion: number;
}

const undoKey = (pageId: string) => `mrai.tiled-line-edit-undo.${pageId}`;

/** The version the page is stored at now; a tiled page's body is its small part. */
export async function storedPageVersion(pageId: string): Promise<number> {
  const page = await apiGet<{ content_version: number }>(`/api/v1/pages/${encodeURIComponent(pageId)}`);
  return page.content_version;
}

export async function sendTiledLineEdit(pageId: string, request: TiledLineEditRequest): Promise<TiledLineEditUndoStep> {
  const answer = await apiPost<{ revision: number; content_version: number }>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/workbook-tiles/line-edit`,
    request,
  );
  return { revision: answer.revision, contentVersion: answer.content_version };
}

/**
 * Whether a line edit that failed may be sent again as it is, at the page's
 * version now: only when the server is known to have written nothing.
 * - The page moved on between the read of its version and the edit (a save
 *   of this tab's own landing, a collaborator): `moved_on`, or the body's save
 *   refused at `page_version_conflict`. Both are refused before any commit.
 * - The database was busy (a row lock waited out, `database_busy`): the
 *   transaction was rolled back, but where it stood is not said. The page's
 *   version tells: every line edit that lands moves it, so an unmoved version
 *   means nothing landed; a moved one may be this edit, and sending it again
 *   would move the lines twice.
 * Anything else (the editors' own refusal, a file that could not take the
 * edit, an unreachable server) is said to the user, not retried.
 */
export function tiledLineEditMayRunAgain(error: unknown, sentAt: number, versionNow: number | null): boolean {
  if (!isApiRequestError(error)) return false;
  if (error.code === "page_version_conflict") return true;
  if (error.code === "workbook_line_edit_refused") return error.details?.reason === "moved_on";
  if (error.code === "database_busy") return versionNow === sentAt;
  return false;
}

export async function undoTiledLineEdit(pageId: string, step: TiledLineEditUndoStep): Promise<void> {
  await apiPost(`/api/v1/pages/${encodeURIComponent(pageId)}/workbook-tiles/line-edit/undo`, {
    revision: step.revision,
    base_content_version: step.contentVersion,
  });
}

export function rememberTiledLineEditUndo(pageId: string, step: TiledLineEditUndoStep): void {
  try {
    sessionStorage.setItem(undoKey(pageId), JSON.stringify(step));
  } catch {
    // Without storage the edit stands; the version history still brings it back.
  }
}

export function tiledLineEditUndo(pageId: string): TiledLineEditUndoStep | null {
  try {
    const raw = sessionStorage.getItem(undoKey(pageId));
    const step = raw ? (JSON.parse(raw) as Partial<TiledLineEditUndoStep>) : null;
    return step && Number.isInteger(step.revision) && Number.isInteger(step.contentVersion)
      ? (step as TiledLineEditUndoStep)
      : null;
  } catch {
    return null;
  }
}

export function forgetTiledLineEditUndo(pageId: string): void {
  try {
    sessionStorage.removeItem(undoKey(pageId));
  } catch {
    // Nothing kept, nothing to forget.
  }
}

const frontKey = (pageId: string) => `mrai.tiled-line-edit-front.${pageId}`;

/** The reload that shows the edit restores the canvas: this page's window is
 * the one the person was working in, whichever one the canvas had saved on top. */
export function keepTiledLineEditWindowInFront(pageId: string): void {
  try {
    sessionStorage.setItem(frontKey(pageId), "1");
  } catch {
    // The window just opens where the canvas had it.
  }
}

export function takeTiledLineEditWindowInFront(pageId: string): boolean {
  try {
    const kept = sessionStorage.getItem(frontKey(pageId)) !== null;
    sessionStorage.removeItem(frontKey(pageId));
    return kept;
  } catch {
    return false;
  }
}





export interface TiledLineEditFailureKept {
  sheetId: string;
  axis: "row" | "col";
  at: number;
  delta: number;
  message: string;
}

const failureKey = (pageId: string) => `mrai.tiled-line-edit-failure.${pageId}`;

export function keepTiledLineEditFailure(pageId: string, failure: TiledLineEditFailureKept): void {
  try {
    sessionStorage.setItem(failureKey(pageId), JSON.stringify(failure));
  } catch {
    // Without storage the reload shows the page as the server holds it.
  }
}

export function takeTiledLineEditFailure(pageId: string): TiledLineEditFailureKept | null {
  try {
    const raw = sessionStorage.getItem(failureKey(pageId));
    sessionStorage.removeItem(failureKey(pageId));
    const kept = raw ? (JSON.parse(raw) as Partial<TiledLineEditFailureKept>) : null;
    return kept
      && typeof kept.sheetId === "string"
      && (kept.axis === "row" || kept.axis === "col")
      && Number.isInteger(kept.at)
      && Number.isInteger(kept.delta)
      && typeof kept.message === "string"
      ? (kept as TiledLineEditFailureKept)
      : null;
  } catch {
    return null;
  }
}
