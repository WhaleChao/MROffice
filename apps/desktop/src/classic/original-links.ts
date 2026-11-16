/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which page each file on this computer became, kept in `originals.json`
 * (the app's data folder; ./local-files.ts names it). No Electron here: the
 * shell's windows (Save, opening a file again) and the AI's computer tools
 * (`/v1/originals` on the agent control port, ./agent-control.ts) read and
 * write the same links through it.
 *
 * The file is read again whenever it changed on disk since this process last
 * read or wrote it (its size, modification time and inode): a link written by
 * the other side, or by another copy of the app, is seen at once instead of
 * the copy read at start (the AI's import linked a page the window's Save
 * never saw).
 */

import fs from "node:fs";
import path from "node:path";

export interface OriginalLink {
  path: string;
  /** The notebook the page is in: opening the file again opens it there. */
  notebookId?: string;
  /** The file as it was when linked or last saved: unchanged, it opens as
   * that page again instead of being imported again. */
  size?: number;
  mtimeMs?: number;
  /** The page body the file holds (last saved: its content version and
   * causal revision): Save of a page still at it writes nothing. */
  savedState?: string;
}

const PAGE_ID = /^[A-Za-z0-9-]{8,64}$/;
const SAVED_STATE = /^[A-Za-z0-9_=&:.%+-]{1,400}$/;

interface Stamp {
  size: number;
  mtimeMs: number;
  ino: number;
}

function sameStamp(a: Stamp | null, b: Stamp | null): boolean {
  return !!a && !!b && a.size === b.size && a.mtimeMs === b.mtimeMs && a.ino === b.ino;
}

function stampOf(file: string): Stamp | null {
  try {
    const stat = fs.statSync(file);
    return { size: stat.size, mtimeMs: stat.mtimeMs, ino: stat.ino };
  } catch {
    return null;
  }
}

/** One path or the other through a symbolic link (macOS's /var and
 * /private/var): the same file. */
function samePath(a: string, b: string): boolean {
  if (a === b) return true;
  try {
    return fs.realpathSync(a) === fs.realpathSync(b);
  } catch {
    return false;
  }
}

/** Write ``bytes`` over ``target`` in one step: a reader never sees half a file. */
export async function writeFileAtomically(target: string, bytes: Uint8Array): Promise<void> {
  const temp = path.join(path.dirname(target), `.${path.basename(target)}.mrnote-saving`);
  await fs.promises.writeFile(temp, bytes);
  try {
    await fs.promises.rename(temp, target);
  } catch (error) {
    await fs.promises.rm(temp, { force: true });
    throw error;
  }
}

export interface OriginalLinks {
  linkOriginal(pageId: string, filePath: string, notebookId?: string, savedState?: string): boolean;
  originalOf(pageId: string): string | null;
  savedStateOf(pageId: string): string | null;
  linkOf(pageId: string): OriginalLink | null;
  /** The file is as it was when it was linked or last saved. */
  unchangedSinceLinked(pageId: string): boolean;
  pageOfUnchangedFile(filePath: string, stat: { size: number; mtimeMs: number }): { pageId: string; notebookId: string } | null;
}

export function originalLinks(storeFile: () => string): OriginalLinks {
  let links: Record<string, OriginalLink> | null = null;
  let stamp: Stamp | null = null;

  const read = (): Record<string, OriginalLink> => {
    const now = stampOf(storeFile());
    if (links && (now === null || sameStamp(now, stamp))) {
      // Unchanged, or not on disk (a write that failed keeps its links until
      // the app quits).
      return links;
    }
    try {
      const value = JSON.parse(fs.readFileSync(storeFile(), "utf8")) as Record<string, OriginalLink>;
      links = value && typeof value === "object" && !Array.isArray(value) ? value : {};
      stamp = now;
    } catch {
      // Half written by another writer, or never written: what was read last.
      links = links ?? {};
    }
    return links;
  };

  const write = (): void => {
    try {
      const file = storeFile();
      fs.mkdirSync(path.dirname(file), { recursive: true });
      const temp = `${file}.tmp`;
      fs.writeFileSync(temp, JSON.stringify(links ?? {}));
      fs.renameSync(temp, file);
      stamp = stampOf(file);
    } catch {
      // The link lasts until the app quits.
    }
  };

  const unchanged = (link: OriginalLink): boolean => {
    try {
      const stat = fs.statSync(link.path);
      return (
        typeof link.size === "number" &&
        typeof link.mtimeMs === "number" &&
        stat.size === link.size &&
        Math.abs(stat.mtimeMs - link.mtimeMs) < 1
      );
    } catch {
      return false;
    }
  };

  const linkOf = (pageId: string): OriginalLink | null => {
    if (!PAGE_ID.test(pageId)) return null;
    const link = read()[pageId];
    return link && typeof link.path === "string" ? link : null;
  };

  return {
    /** Link a page to the file it came from (or was saved to), as the file is now. */
    linkOriginal(pageId, filePath, notebookId, savedState) {
      if (!PAGE_ID.test(pageId) || !path.isAbsolute(filePath)) return false;
      // Read again first: a link written since by the other side is kept.
      const current = read();
      const known = current[pageId];
      let stat: fs.Stats | null = null;
      try {
        stat = fs.statSync(filePath);
      } catch {
        stat = null;
      }
      current[pageId] = {
        path: filePath,
        notebookId: notebookId && PAGE_ID.test(notebookId) ? notebookId : known?.notebookId,
        ...(stat ? { size: stat.size, mtimeMs: stat.mtimeMs } : {}),
        ...(savedState && SAVED_STATE.test(savedState) ? { savedState } : {}),
      };
      links = current;
      write();
      return true;
    },

    /** The file a page came from, while it is still there. */
    originalOf(pageId) {
      const linked = linkOf(pageId)?.path;
      return linked && fs.existsSync(linked) ? linked : null;
    },

    /** The page body that file holds, while it is unchanged on disk. */
    savedStateOf(pageId) {
      const link = linkOf(pageId);
      return link?.savedState && unchanged(link) ? link.savedState : null;
    },

    linkOf,

    unchangedSinceLinked(pageId) {
      const link = linkOf(pageId);
      return !!link && unchanged(link);
    },

    /** The page a file became, while the file is as it was then (opening it
     * again opens that page, as Word opens a document, instead of importing
     * a copy); null when it was never opened or changed since. */
    pageOfUnchangedFile(filePath, stat) {
      let found: { pageId: string; notebookId: string; mtimeMs: number } | null = null;
      for (const [pageId, link] of Object.entries(read())) {
        if (
          link &&
          typeof link.path === "string" &&
          link.notebookId &&
          link.size === stat.size &&
          typeof link.mtimeMs === "number" &&
          Math.abs(link.mtimeMs - stat.mtimeMs) < 1 &&
          (!found || link.mtimeMs > found.mtimeMs) &&
          samePath(link.path, filePath)
        ) {
          found = { pageId, notebookId: link.notebookId, mtimeMs: link.mtimeMs };
        }
      }
      return found ? { pageId: found.pageId, notebookId: found.notebookId } : null;
    },
  };
}

// --- the AI's computer tools (`POST /v1/originals`) ---------------------------------

export type OriginalsAnswer = Record<string, unknown> & { ok: boolean };

function refuse(error: string, message: string): OriginalsAnswer {
  return { ok: false, error, message };
}

/**
 * One request of the API's computer tools (app/desktop/computer/control.py):
 *
 * - `lookup` {path}: the page an unchanged file already is.
 * - `link` {page_id, path, notebook_id?, saved_state?}: link a page to a file.
 * - `of` {page_id}: the file a page is linked to, whether it is unchanged
 *   since, and the page body it holds.
 * - `write` {path, bytes_base64, page_id?, saved_state?}: write a file in one
 *   step. With page_id it is that page's own file, written only while it is
 *   still the file linked and unchanged since (someone else's edit of it is
 *   never written over), and the link then records what it holds.
 *
 * The person approved the step in the API (its cards); the shell only does it.
 */
export async function originalsAct(store: OriginalLinks, request: Record<string, unknown>): Promise<OriginalsAnswer> {
  const action = String(request.action ?? "");
  const pageId = typeof request.page_id === "string" ? request.page_id : "";
  const filePath = typeof request.path === "string" ? request.path : "";
  if (action === "lookup") {
    if (!path.isAbsolute(filePath)) return refuse("bad_request", "path must be absolute");
    let stat: fs.Stats;
    try {
      stat = fs.statSync(filePath);
    } catch {
      return refuse("not_found", "The file does not exist.");
    }
    if (!stat.isFile()) return refuse("not_a_file", "The path is not a file.");
    const page = store.pageOfUnchangedFile(filePath, stat);
    return { ok: true, page_id: page?.pageId ?? null, notebook_id: page?.notebookId ?? null };
  }
  if (action === "link") {
    const notebookId = typeof request.notebook_id === "string" ? request.notebook_id : undefined;
    const savedState = typeof request.saved_state === "string" ? request.saved_state : undefined;
    if (!fs.existsSync(filePath)) return refuse("not_found", "The file does not exist.");
    return store.linkOriginal(pageId, filePath, notebookId, savedState)
      ? { ok: true }
      : refuse("bad_request", "page_id or path is invalid");
  }
  if (action === "of") {
    const link = store.linkOf(pageId);
    if (!link || !fs.existsSync(link.path)) return { ok: true, path: null };
    return {
      ok: true,
      path: link.path,
      name: path.basename(link.path),
      notebook_id: link.notebookId ?? null,
      unchanged: store.unchangedSinceLinked(pageId),
      saved_state: store.savedStateOf(pageId),
    };
  }
  if (action === "write") {
    if (!path.isAbsolute(filePath)) return refuse("bad_request", "path must be absolute");
    if (typeof request.bytes_base64 !== "string") return refuse("bad_request", "bytes_base64 is required");
    const bytes = Buffer.from(request.bytes_base64, "base64");
    if (pageId) {
      const link = store.linkOf(pageId);
      if (!link || !samePath(link.path, filePath)) {
        return refuse("not_linked", "This page is not linked to that file.");
      }
      if (!store.unchangedSinceLinked(pageId)) {
        return refuse("original_changed", "The file was changed outside MR Note since it was opened or last saved.");
      }
    }
    try {
      if (fs.statSync(filePath, { throwIfNoEntry: false })?.isDirectory()) {
        return refuse("is_a_folder", "The path is a folder.");
      }
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      await writeFileAtomically(filePath, bytes);
    } catch (error) {
      return refuse("write_failed", error instanceof Error ? error.message : String(error));
    }
    if (pageId) {
      const savedState = typeof request.saved_state === "string" ? request.saved_state : undefined;
      store.linkOriginal(pageId, store.linkOf(pageId)?.path ?? filePath, undefined, savedState);
    }
    return { ok: true, path: filePath, size: bytes.length };
  }
  return refuse("bad_request", `unknown action ${action}`);
}
