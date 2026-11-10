/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A page's pictures stored beside it (picture objects).
 *
 * The twin of `apps/api/app/services/page_picture_objects.py`. An imported
 * deck's (or workbook's) page names each picture by the digest of its bytes,
 * `mrai-picture:<sha256>`, and the bytes are the page's own stored file, read
 * from `GET /api/v1/pages/{page}/pictures/{sha256}` -- the file's original,
 * never a display copy.
 *
 * The editor draws a picture from that URL: the loader turns a token into it
 * (`presentation-page-media.expandPresentationMedia`) and the serializer turns
 * it back (`compactPresentationMedia`), so the saved JSON never carries the
 * bytes and the browser loads each picture lazily, once (the answer is
 * content-addressed and cached).
 *
 * Which page holds a picture is learned when a page is read
 * (`registerPagePictureObjects`): a deck's own pictures when it loads, another
 * deck's when Reuse Slides reads it, a thumbnail's deck. A token no page read
 * so far names is drawn from the deck opened last; before any deck is opened,
 * it stays a token. A URL naming another page's picture (a paste from another deck) is
 * saved as its token and the server copies the picture into this page.
 */

import { activeMountWorkspaceId, mountWorkspaceFor } from "./workspace-mounts";

export const PICTURE_OBJECT_TOKEN_PREFIX = "mrai-picture:";
const TOKEN_RE = /^mrai-picture:([0-9a-f]{64})$/;
const URL_RE = /^\/api\/v1\/pages\/[^/?#]+\/pictures\/([0-9a-f]{64})(?:\?[^#]*)?$/;
/** Bounded like the server's per-page bound; the oldest entries go first. */
const REGISTRY_LIMIT = 20_000;

interface PictureHome {
  readonly pageId: string;
  readonly workspaceId: string | null;
}

const homes = new Map<string, PictureHome>();
/**
 * The deck opened last: a token no page has been seen to hold (a picture a
 * save stored for the first time, such as a reused slide's) is drawn from it.
 */
let lastHome: PictureHome | null = null;

type Rec = Record<string, unknown>;

function asRec(value: unknown): Rec | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Rec)
    : null;
}

/** The digest a token names, or null. */
export function pictureObjectTokenSha(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const match = TOKEN_RE.exec(value);
  return match ? match[1] : null;
}

/** The digest a picture URL of this module names, or null. */
export function pictureObjectUrlSha(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const match = URL_RE.exec(value);
  return match ? match[1] : null;
}

export function isPictureObjectUrl(value: unknown): boolean {
  return pictureObjectUrlSha(value) !== null;
}

/** The URL a picture is drawn from, for a page known to hold it. */
export function pagePictureUrl(
  pageId: string,
  sha: string,
  workspaceId: string | null = null,
): string {
  const base = `/api/v1/pages/${encodeURIComponent(pageId)}/pictures/${sha}`;
  // A mounted shared notebook lives in its host's workspace; an <img> sends
  // no header, so the host is named in the query (the API reads either).
  return workspaceId ? `${base}?workspace_id=${encodeURIComponent(workspaceId)}` : base;
}

/**
 * Every image block list a page holds: a deck's slides, masters and layouts;
 * a workbook's sheets' drawing objects.
 */
function blockLists(content: Rec): unknown[][] {
  const lists: unknown[][] = [];
  for (const sheet of Array.isArray(content.sheets) ? content.sheets : []) {
    const record = asRec(sheet);
    if (record && Array.isArray(record.objects)) lists.push(record.objects);
  }
  for (const slide of Array.isArray(content.slides) ? content.slides : []) {
    const record = asRec(slide);
    if (record && Array.isArray(record.blocks)) lists.push(record.blocks);
  }
  for (const master of Array.isArray(content.masters) ? content.masters : []) {
    const record = asRec(master);
    if (!record) continue;
    if (Array.isArray(record.blocks)) lists.push(record.blocks);
    for (const layout of Array.isArray(record.layouts) ? record.layouts : []) {
      const layoutRecord = asRec(layout);
      if (!layoutRecord) continue;
      for (const key of ["blocks", "masterGraphics"] as const) {
        if (Array.isArray(layoutRecord[key])) lists.push(layoutRecord[key] as unknown[]);
      }
    }
  }
  return lists;
}

/** How deep a picture record may sit under a slide or a master. */
const PICTURE_RECORD_MAX_DEPTH = 12;

/**
 * A picture drawn as a fill rather than as an image block -- a slide's,
 * layout's or master's background, a shape's picture fill: `{ url, fit }`.
 */
function isPictureRecord(record: Rec): boolean {
  return typeof record.url === "string" && "fit" in record;
}

/** `node` with every picture record's `url` passed through `rewrite`, copy-on-write. */
function rewritePictureRecordNode(
  node: unknown,
  rewrite: (url: string) => string,
  depth: number,
): unknown {
  if (depth > PICTURE_RECORD_MAX_DEPTH) return node;
  if (Array.isArray(node)) {
    let copy: unknown[] | null = null;
    node.forEach((item, index) => {
      const next = rewritePictureRecordNode(item, rewrite, depth + 1);
      if (next !== item) {
        copy ??= node.slice();
        copy[index] = next;
      }
    });
    return copy ?? node;
  }
  const record = asRec(node);
  if (!record) return node;
  let copy: Rec | null = null;
  if (isPictureRecord(record)) {
    const url = rewrite(record.url as string);
    if (url !== record.url) copy = { ...record, url };
  }
  for (const [key, item] of Object.entries(record)) {
    if (typeof item !== "object" || item === null) continue;
    const next = rewritePictureRecordNode(item, rewrite, depth + 1);
    if (next !== item) {
      copy ??= { ...record };
      copy[key] = next;
    }
  }
  return copy ?? record;
}

/**
 * A slide or master (or a whole deck page) with each picture record that
 * names a stored picture drawn from its URL; the same object when none does.
 */
export function expandPictureObjectRecords<T>(node: T): T {
  return rewritePictureRecordNode(
    node,
    (url) => pictureObjectUrlForToken(url) ?? url,
    0,
  ) as T;
}

/** The stored form of {@link expandPictureObjectRecords}: each URL its token. */
export function compactPictureObjectRecords<T>(node: T): T {
  return rewritePictureRecordNode(node, pictureObjectTokenForUrl, 0) as T;
}

/**
 * Remember that `pageId` holds every picture its stored `content` names, so
 * the editor can draw them. Call it with a page as the API answered it.
 */
export function registerPagePictureObjects(
  pageId: string,
  content: unknown,
  notebookId?: string | null,
  options: { readonly open?: boolean } = {},
): void {
  const record = asRec(content);
  if (!pageId || !record) return;
  const workspaceId =
    (notebookId ? mountWorkspaceFor(notebookId) : null) ?? activeMountWorkspaceId();
  // Only a deck opened for editing is where a picture nobody named yet is
  // looked for: a thumbnail or a Reuse Slides source is read, not opened.
  if (options.open) lastHome = { pageId, workspaceId };
  const remember = (sha: string) => {
    homes.delete(sha);
    if (homes.size >= REGISTRY_LIMIT) {
      const oldest = homes.keys().next().value;
      if (oldest !== undefined) homes.delete(oldest);
    }
    homes.set(sha, { pageId, workspaceId });
  };
  for (const blocks of blockLists(record)) {
    for (const block of blocks) {
      const image = asRec(block);
      if (!image || image.type !== "image") continue;
      const sha = pictureObjectTokenSha(image.src);
      if (sha) remember(sha);
    }
  }
  // A slide's recorded narration is the file's own clip, named the same way
  
  for (const slide of Array.isArray(record.slides) ? record.slides : []) {
    const sha = pictureObjectTokenSha(asRec(slide)?.narration);
    if (sha) remember(sha);
  }
  // Backgrounds and picture fills name stored pictures too.
  for (const key of ["slides", "masters"] as const) {
    rewritePictureRecordNode(
      record[key],
      (url) => {
        const sha = pictureObjectTokenSha(url);
        if (sha) remember(sha);
        return url;
      },
      0,
    );
  }
}

/** The URL a token is drawn from, or null when no page read so far holds it. */
export function pictureObjectUrlForToken(value: unknown): string | null {
  const sha = pictureObjectTokenSha(value);
  if (!sha) return null;
  const home = homes.get(sha) ?? lastHome;
  return home ? pagePictureUrl(home.pageId, sha, home.workspaceId) : null;
}

/**
 * What an `<img>` draws for a picture source: a stored picture's URL for its
 * token (`pictureObjectUrlForToken`), anything else as it is. A workbook keeps
 * the token in its model and resolves it only where a picture is drawn.
 */
export function pictureObjectDisplaySrc(value: string): string;
export function pictureObjectDisplaySrc(value: string | undefined): string | undefined;
export function pictureObjectDisplaySrc(value: string | undefined): string | undefined {
  if (value === undefined || !value.startsWith(PICTURE_OBJECT_TOKEN_PREFIX)) return value;
  return pictureObjectUrlForToken(value) ?? value;
}






export function slideNarrationPlaybackSrc(value: unknown): string | null {
  if (typeof value !== "string") return null;
  if (value.startsWith("data:audio/")) return value;
  return pictureObjectUrlForToken(value);
}

/** The token a picture URL stands for (any page's), else `value` itself. */
export function pictureObjectTokenForUrl(value: string): string {
  const sha = pictureObjectUrlSha(value);
  return sha ? `${PICTURE_OBJECT_TOKEN_PREFIX}${sha}` : value;
}

/** Test seam: forget every registered page. */
export function resetPagePictureObjectsForTests(): void {
  homes.clear();
  lastHome = null;
}
