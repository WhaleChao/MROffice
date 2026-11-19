/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiPost } from "./api";
import {
  captureSavedPageHotCacheScope,
  isSavedPageHotCacheScopeCurrent,
} from "./saved-page-hot-cache";
import { runCollaborationRequestWithTimeout } from "./collaboration-windows/request-timeout";

export const NOTEBOOK_PAGE_CATALOGUE_EVENT = "mrnote:notebook-page-catalogue";
const LIMIT = 5000;
export interface CataloguePage {
  id: string;
  title: string;
  page_type: string;
  content_version: number;
  created_by?: string;
  sort_order?: number;
  created_at?: string;
}
export interface PageCatalogueDelta {
  notebookId: string;
  reset: boolean;
  upserts: CataloguePage[];
  removed: string[];
}
interface Reply {
  protocol: "page-catalogue-v1";
  scope: string;
  reset: boolean;
  upserts: Array<{ revision: string; page: CataloguePage }>;
  removed: string[];
}
type RequestCatalogue = (
  notebookId: string,
  body: { scope: string | null; known: Record<string, string> },
  signal: AbortSignal,
) => Promise<unknown>;
const requestCatalogue: RequestCatalogue = (notebookId, body, signal) =>
  apiPost(
    `/api/v1/notebooks/${encodeURIComponent(notebookId)}/page-changes`,
    body,
    { signal },
  );

function parseReply(value: unknown): Reply {
  if (!value || typeof value !== "object")
    throw new Error("Invalid page catalogue");
  const reply = value as Reply;
  if (
    reply.protocol !== "page-catalogue-v1" ||
    typeof reply.scope !== "string" ||
    !/^[0-9a-f]{64}$/u.test(reply.scope) ||
    typeof reply.reset !== "boolean" ||
    !Array.isArray(reply.upserts) ||
    reply.upserts.length > LIMIT ||
    !Array.isArray(reply.removed) ||
    reply.removed.length > LIMIT
  )
    throw new Error("Invalid page catalogue");
  const seen = new Set<string>();
  for (const entry of reply.upserts) {
    const page = entry?.page;
    if (
      !page ||
      typeof page.id !== "string" ||
      !page.id ||
      page.id.length > 255 ||
      typeof page.title !== "string" ||
      typeof page.page_type !== "string" ||
      !Number.isSafeInteger(page.content_version) ||
      page.content_version < 1 ||
      typeof entry.revision !== "string" ||
      !/^[0-9a-f]{64}$/u.test(entry.revision) ||
      seen.has(page.id)
    )
      throw new Error("Invalid page catalogue item");
    seen.add(page.id);
  }
  for (const id of reply.removed) {
    if (typeof id !== "string" || !id || id.length > 255 || seen.has(id)) {
      throw new Error("Invalid page catalogue removal");
    }
    seen.add(id);
  }
  return reply;
}

/** Preserve existing row identities and any newer local save receipt. */
export function applyPageCatalogueDelta<T extends CataloguePage>(
  current: T[],
  delta: PageCatalogueDelta,
): T[] {
  const prior = new Map(current.map((page) => [page.id, page]));
  const next = delta.reset ? new Map<string, T>() : new Map(prior);
  for (const id of delta.removed) next.delete(id);
  for (const page of delta.upserts) {
    const existing = delta.reset ? undefined : prior.get(page.id);
    next.set(
      page.id,
      existing && existing.content_version > page.content_version
        ? existing
        : ({ ...existing, ...page } as T),
    );
  }
  const pages = [...next.values()].sort(
    (a, b) =>
      (a.sort_order ?? 0) - (b.sort_order ?? 0) ||
      (b.created_at ?? "").localeCompare(a.created_at ?? "") ||
      a.id.localeCompare(b.id),
  );
  return pages.length === current.length &&
    pages.every((page, index) => page === current[index])
    ? current
    : pages;
}

export function createNotebookPageCatalogue(
  notebookId: string,
  onDelta: (delta: PageCatalogueDelta, changedPageIds: string[]) => void,
  request: RequestCatalogue = requestCatalogue,
) {
  let scope: string | null = null;
  let known = new Map<string, string>();
  let controller: AbortController | null = null;
  let active: Promise<void> | null = null;
  let requested = false;
  let closed = false;
  let account = captureSavedPageHotCacheScope();
  const refresh = (): Promise<void> => {
    if (closed) return Promise.resolve();
    requested = true;
    if (active) return active;
    const work = (async () => {
      while (requested && !closed) {
        requested = false;
        if (!isSavedPageHotCacheScopeCurrent(account)) {
          account = captureSavedPageHotCacheScope();
          scope = null;
          known = new Map();
        }
        const requestAccount = account;
        controller = new AbortController();
        try {
          const reply = parseReply(
            await runCollaborationRequestWithTimeout(
              controller,
              (signal) =>
                request(
                  notebookId,
                  { scope, known: Object.fromEntries(known) },
                  signal,
                ),
              "Notebook catalogue refresh timed out",
            ),
          );
          if (closed || !isSavedPageHotCacheScopeCurrent(requestAccount))
            return;
          if (!reply.reset && reply.scope !== scope)
            throw new Error("Catalogue scope changed");
          const wasInitialized = scope !== null;
          const next = reply.reset ? new Map<string, string>() : new Map(known);
          for (const id of reply.removed) next.delete(id);
          for (const entry of reply.upserts)
            next.set(entry.page.id, entry.revision);
          if (next.size > LIMIT) throw new Error("Catalogue is too large");
          const changed = wasInitialized
            ? [
                ...new Set([
                  ...reply.upserts.map((entry) => entry.page.id),
                  ...reply.removed,
                  ...[...known.keys()].filter((id) => !next.has(id)),
                ]),
              ]
            : [];
          scope = reply.scope;
          known = next;
          if (reply.reset || reply.upserts.length || reply.removed.length) {
            onDelta(
              {
                notebookId,
                reset: reply.reset,
                upserts: reply.upserts.map((entry) => entry.page),
                removed: reply.removed,
              },
              changed,
            );
          }
        } catch {
          // Keep the last acknowledged catalogue. The next heartbeat/open
          // rechecks durable rows; an unavailable request cannot delete rows.
          requested = false;
        } finally {
          controller = null;
        }
      }
    })();
    active = work.finally(() => {
      active = null;
    });
    return active;
  };
  return {
    refresh,
    close() {
      closed = true;
      requested = false;
      controller?.abort();
    },
  };
}
