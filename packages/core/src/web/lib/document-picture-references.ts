/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pictures an imported Word document names by digest.
 *
 * The server's twin is `apps/api/app/services/file_adapters/docx/picture_references.py`.
 * A document imported from Word names each of its pictures by the digest of
 * its bytes, `mrai-picture:<sha256>`, and the bytes stay in the stored file,
 * read from `GET /api/v1/pages/{page}/pictures/{sha256}` -- the file's own
 * picture at its own fidelity, never a downscaled copy.
 *
 * The token is what the document holds: every save, the collaboration
 * document, history and Undo carry it as an ordinary string. Only the picture
 * the editor draws turns it into that URL (`ImageResizeNodeView`). Which page
 * holds a picture is learned when a document page is read
 * (`registerDocumentPictureReferences`); a token no page read so far names is
 * drawn from the document opened last.
 */

import { assertPageKey } from "./page-key-guard";
import { pagePictureUrl, pictureObjectTokenSha } from "./page-picture-objects";
import { activeMountWorkspaceId, mountWorkspaceFor } from "./workspace-mounts";

interface PictureHome {
  readonly pageId: string;
  readonly workspaceId: string | null;
}

/** Bounded like the server's per-page bound; the oldest entries go first. */
const REGISTRY_LIMIT = 20_000;
const homes = new Map<string, PictureHome>();
let lastHome: PictureHome | null = null;
/** Drawers of a token no page read so far named, told when a page is read. */
const waiting = new Set<() => void>();

function visit(node: unknown, found: Set<string>, depth: number): void {
  if (depth > 200 || !node || typeof node !== "object") return;
  if (Array.isArray(node)) {
    for (const child of node) visit(child, found, depth + 1);
    return;
  }
  // Any record naming a picture by token: a picture node's attrs, and the
  // pictures a header/footer band holds in the document's own attrs
  // (`headerRich[].images[]`, `sectionBands[]`), which are no node's attrs.
  const record = node as Record<string, unknown>;
  const sha = pictureObjectTokenSha(record.src);
  if (sha) found.add(sha);
  for (const child of Object.values(record)) {
    if (child && typeof child === "object") visit(child, found, depth + 1);
  }
}

export function documentPictureReferenceShas(content: unknown): Set<string> {
  const found = new Set<string>();
  visit(content, found, 0);
  return found;
}

/**
 * Remember that `pageId` holds every picture its stored `content` names, so
 * the editor can draw them. Call it with a document page as the API answered it.
 */
export function registerDocumentPictureReferences(
  pageId: string,
  content: unknown,
  notebookId?: string | null,
): void {
  assertPageKey(pageId, "document-picture-references registerDocumentPictureReferences");
  if (!pageId) return;
  const workspaceId =
    (notebookId ? mountWorkspaceFor(notebookId) : null) ?? activeMountWorkspaceId();
  const home: PictureHome = { pageId, workspaceId };
  lastHome = home;
  for (const sha of documentPictureReferenceShas(content)) {
    homes.delete(sha);
    if (homes.size >= REGISTRY_LIMIT) {
      const oldest = homes.keys().next().value;
      if (oldest !== undefined) homes.delete(oldest);
    }
    homes.set(sha, home);
  }
  for (const listener of [...waiting]) listener();
}

/**
 * Call `listener` after each page read from now on, until the returned
 * function is called: a picture drawn before its page was read (the editor
 * can draw a document's room before the page itself is read) resolves its
 * token again then, rather than staying blank (corpus R281-D10: seven cell
 * pictures, drawn once and never again).
 */
export function onDocumentPictureReferencesRegistered(listener: () => void): () => void {
  waiting.add(listener);
  return () => {
    waiting.delete(listener);
  };
}

/**
 * Whether `src` is drawn from the page that holds it: anything but a token,
 * or a token a page read so far names. A token no page named yet is drawn
 * from the document opened last ({@link documentPictureDisplaySrc}), or not
 * at all: that is a guess until its page is read.
 */
export function documentPictureHomeKnown(src: unknown): boolean {
  const sha = pictureObjectTokenSha(typeof src === "string" ? src : "");
  return !sha || homes.has(sha);
}

/**
 * Draw `img` from `src` ({@link documentPictureDisplaySrc}); a token no page
 * read so far names is drawn again from its own page once that page is read.
 * Returns a function that stops waiting (call it when the picture is redrawn
 * or goes away).
 */
export function drawDocumentPicture(img: HTMLImageElement, src: unknown): () => void {
  img.src = documentPictureDisplaySrc(src);
  if (documentPictureHomeKnown(src)) return () => {};
  const stop = onDocumentPictureReferencesRegistered(() => {
    if (!documentPictureHomeKnown(src)) return;
    stop();
    const later = documentPictureDisplaySrc(src);
    if (img.getAttribute("src") !== later) img.src = later;
  });
  return stop;
}

/**
 * The address a picture is drawn from: the stored picture's URL for a token,
 * `src` itself for anything else (a `data:` picture, a link). A token no
 * document read so far holds is drawn from the document opened last.
 */
export function documentPictureDisplaySrc(src: unknown): string {
  const value = typeof src === "string" ? src : "";
  const sha = pictureObjectTokenSha(value);
  if (!sha) return value;
  const home = homes.get(sha) ?? lastHome;
  return home ? pagePictureUrl(home.pageId, sha, home.workspaceId) : "";
}

/** Test seam: forget every registered page. */
export function resetDocumentPictureReferencesForTests(): void {
  homes.clear();
  lastHome = null;
  waiting.clear();
}
