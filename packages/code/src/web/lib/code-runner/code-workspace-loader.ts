/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded detail loading and search projection for notebook code workspaces.
 *
 * The notebook page-list endpoint intentionally returns metadata only. Code
 * tools hydrate authorized code pages through the existing per-page API, with
 * explicit concurrency and memory budgets, before projecting their content
 * into the synchronous Find-in-Files engine.
 */

import type { SearchBuffer } from "./find-in-files";
import type {
  CodeWorkspaceFile,
  CodeWorkspacePageLike,
} from "./code-workspace";

export const CODE_WORKSPACE_DETAIL_LIMIT = 500;
export const CODE_WORKSPACE_DETAIL_CONCURRENCY = 4;
/** Pages per request when the caller can read several at once. */
export const CODE_WORKSPACE_DETAIL_BATCH = 25;
export const CODE_WORKSPACE_SEARCH_FILE_CODE_UNIT_LIMIT = 8 * 1024 * 1024;
export const CODE_WORKSPACE_SEARCH_TOTAL_CODE_UNIT_LIMIT = 32 * 1024 * 1024;

const CODE_WORKSPACE_METADATA_KEYS = [
  "type",
  "language",
  "sourceFileName",
  "codeWorkspace",
  // A binary file's attachment (workspace-binary-file.ts): small, and what a
  // LaTeX compile sends for it.
  "binaryAsset",
] as const;

export interface CodeWorkspaceDetailLoadOptions<
  TPage extends CodeWorkspacePageLike = CodeWorkspacePageLike,
> {
  readonly limit?: number;
  readonly concurrency?: number;
  readonly maxFileCodeUnits?: number;
  readonly maxTotalCodeUnits?: number;
  /**
   * Reads several pages in one request, each as `fetchPage` would (the
   * notebook's batch read). A page it does not return fails like a failed
   * single read; a batch that fails as a whole is read page by page.
   */
  readonly fetchPages?: (pageIds: readonly string[]) => Promise<readonly TPage[]>;
}

export interface CodeWorkspaceDetailLoadResult<
  TPage extends CodeWorkspacePageLike = CodeWorkspacePageLike,
> {
  readonly pages: TPage[];
  readonly attemptedPageIds: string[];
  readonly failedPageIds: string[];
  readonly oversizedPageIds: string[];
  readonly searchableCodeUnits: number;
  readonly truncated: boolean;
}

export interface CodeWorkspaceLiveBuffer {
  readonly pageId: string;
  readonly text: string;
}

export interface CodeWorkspaceSearchProjection {
  readonly buffers: SearchBuffer[];
  readonly totalCodeUnits: number;
  readonly unloadedFileCount: number;
  readonly oversizedFileCount: number;
  readonly truncated: boolean;
}

function boundedInteger(
  value: number | undefined,
  fallback: number,
  maximum: number,
): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(0, Math.min(maximum, Math.trunc(value)));
}

export function codeWorkspacePageSearchText(
  page: CodeWorkspacePageLike | undefined,
): string | null {
  if (page?.page_type !== "code") return null;
  const content = page.content_json?.content;
  return typeof content === "string" ? content : null;
}

function metadataOnlyCodePage<TPage extends CodeWorkspacePageLike>(
  page: TPage,
): TPage {
  const contentJson: Record<string, unknown> = {};
  for (const key of CODE_WORKSPACE_METADATA_KEYS) {
    const value = page.content_json?.[key];
    if (value !== undefined) contentJson[key] = value;
  }
  return {
    ...page,
    content_json: contentJson,
  } as TPage;
}

/**
 * `fetchPage` for one candidate, answered from a read of its batch of
 * candidates (and the next batch is asked for meanwhile): a notebook of 500
 * code files was 500 requests, four at a time.
 */
function batchedPageReader<TPage extends CodeWorkspacePageLike>(
  candidates: readonly TPage[],
  fetchPage: (pageId: string) => Promise<TPage>,
  fetchPages: (pageIds: readonly string[]) => Promise<readonly TPage[]>,
): (pageId: string) => Promise<TPage> {
  const position = new Map(
    candidates.map((page, index) => [page.id, index] as const),
  );
  const batches = new Map<number, Promise<Map<string, TPage> | null>>();
  const batch = (number: number) => {
    let read = batches.get(number);
    if (!read) {
      const ids = candidates
        .slice(
          number * CODE_WORKSPACE_DETAIL_BATCH,
          (number + 1) * CODE_WORKSPACE_DETAIL_BATCH,
        )
        .map((page) => page.id);
      read =
        ids.length === 0
          ? Promise.resolve(new Map<string, TPage>())
          : fetchPages(ids).then(
              (pages) => new Map(pages.map((page) => [page.id, page] as const)),
              () => null,
            );
      batches.set(number, read);
    }
    return read;
  };
  return async (pageId) => {
    const index = position.get(pageId);
    if (index === undefined) return fetchPage(pageId);
    const number = Math.floor(index / CODE_WORKSPACE_DETAIL_BATCH);
    const read = batch(number);
    void batch(number + 1);
    const pages = await read;
    if (pages === null) return fetchPage(pageId);
    const page = pages.get(pageId);
    if (!page) throw new Error("The page could not be read.");
    return page;
  };
}

/**
 * Hydrate a bounded set of shallow code-page rows through an authorized
 * per-page fetcher (or batch reader, `fetchPages`). Oversized pages retain only path/language metadata so one
 * hostile source cannot pin an unbounded search corpus in the renderer.
 */
export async function loadCodeWorkspacePageDetails<
  TPage extends CodeWorkspacePageLike,
>(
  pages: readonly TPage[],
  fetchPage: (pageId: string) => Promise<TPage>,
  options: CodeWorkspaceDetailLoadOptions<TPage> = {},
): Promise<CodeWorkspaceDetailLoadResult<TPage>> {
  const limit = boundedInteger(
    options.limit,
    CODE_WORKSPACE_DETAIL_LIMIT,
    CODE_WORKSPACE_DETAIL_LIMIT,
  );
  const concurrency = Math.max(
    1,
    boundedInteger(
      options.concurrency,
      CODE_WORKSPACE_DETAIL_CONCURRENCY,
      16,
    ),
  );
  const maxFileCodeUnits = boundedInteger(
    options.maxFileCodeUnits,
    CODE_WORKSPACE_SEARCH_FILE_CODE_UNIT_LIMIT,
    CODE_WORKSPACE_SEARCH_FILE_CODE_UNIT_LIMIT,
  );
  const maxTotalCodeUnits = boundedInteger(
    options.maxTotalCodeUnits,
    CODE_WORKSPACE_SEARCH_TOTAL_CODE_UNIT_LIMIT,
    CODE_WORKSPACE_SEARCH_TOTAL_CODE_UNIT_LIMIT,
  );
  const codePages = pages.filter((page) => page.page_type === "code");
  const candidates = codePages.slice(0, limit);
  const details: TPage[] = [];
  const attemptedPageIds: string[] = [];
  const failedPageIds: string[] = [];
  const oversizedPageIds: string[] = [];
  let searchableCodeUnits = 0;
  let cursor = 0;
  let totalBudgetExhausted = false;
  const readPage = options.fetchPages
    ? batchedPageReader(candidates, fetchPage, options.fetchPages)
    : fetchPage;

  const worker = async () => {
    while (!totalBudgetExhausted) {
      const index = cursor;
      cursor += 1;
      const candidate = candidates[index];
      if (!candidate) return;
      attemptedPageIds.push(candidate.id);
      let detail: TPage;
      try {
        detail = await readPage(candidate.id);
      } catch {
        failedPageIds.push(candidate.id);
        continue;
      }
      if (detail.id !== candidate.id || detail.page_type !== "code") {
        failedPageIds.push(candidate.id);
        continue;
      }

      const text = codeWorkspacePageSearchText(detail);
      if (text === null) {
        oversizedPageIds.push(candidate.id);
        details.push(metadataOnlyCodePage(detail));
        continue;
      }
      const codeUnits = text?.length ?? 0;
      if (
        totalBudgetExhausted ||
        codeUnits > maxFileCodeUnits ||
        searchableCodeUnits + codeUnits > maxTotalCodeUnits
      ) {
        oversizedPageIds.push(candidate.id);
        details.push(metadataOnlyCodePage(detail));
        if (searchableCodeUnits + codeUnits > maxTotalCodeUnits) {
          totalBudgetExhausted = true;
        }
        continue;
      }
      searchableCodeUnits += codeUnits;
      details.push(detail);
    }
  };

  await Promise.all(
    Array.from(
      { length: Math.min(concurrency, Math.max(1, candidates.length)) },
      () => worker(),
    ),
  );

  return {
    pages: details,
    attemptedPageIds,
    failedPageIds,
    oversizedPageIds,
    searchableCodeUnits,
    truncated:
      candidates.length < codePages.length ||
      totalBudgetExhausted ||
      failedPageIds.length > 0 ||
      oversizedPageIds.length > 0,
  };
}

/**
 * Prefer unsaved live buffers, then hydrated saved pages, while projecting one
 * stable search buffer per workspace file. Missing and oversized pages are
 * reported instead of silently presenting a complete-looking result set.
 */
export function buildCodeWorkspaceSearchProjection(
  files: readonly CodeWorkspaceFile[],
  pages: readonly CodeWorkspacePageLike[],
  liveBuffers: readonly CodeWorkspaceLiveBuffer[],
  options: Pick<
    CodeWorkspaceDetailLoadOptions,
    "limit" | "maxFileCodeUnits" | "maxTotalCodeUnits"
  > = {},
): CodeWorkspaceSearchProjection {
  const limit = boundedInteger(
    options.limit,
    CODE_WORKSPACE_DETAIL_LIMIT,
    CODE_WORKSPACE_DETAIL_LIMIT,
  );
  const maxFileCodeUnits = boundedInteger(
    options.maxFileCodeUnits,
    CODE_WORKSPACE_SEARCH_FILE_CODE_UNIT_LIMIT,
    CODE_WORKSPACE_SEARCH_FILE_CODE_UNIT_LIMIT,
  );
  const maxTotalCodeUnits = boundedInteger(
    options.maxTotalCodeUnits,
    CODE_WORKSPACE_SEARCH_TOTAL_CODE_UNIT_LIMIT,
    CODE_WORKSPACE_SEARCH_TOTAL_CODE_UNIT_LIMIT,
  );
  const pageById = new Map(pages.map((page) => [page.id, page] as const));
  const liveByPageId = new Map<string, string>();
  const conflictedLivePageIds = new Set<string>();
  for (const buffer of liveBuffers) {
    const existing = liveByPageId.get(buffer.pageId);
    if (existing === undefined) {
      liveByPageId.set(buffer.pageId, buffer.text);
    } else if (existing !== buffer.text) {
      conflictedLivePageIds.add(buffer.pageId);
    }
  }

  const buffers: SearchBuffer[] = [];
  let totalCodeUnits = 0;
  let unloadedFileCount = 0;
  let oversizedFileCount = 0;
  let inspectedFiles = 0;
  let truncated = false;

  for (const file of files) {
    if (file.pageType !== "code") continue;
    if (inspectedFiles >= limit) {
      truncated = true;
      break;
    }
    inspectedFiles += 1;
    if (conflictedLivePageIds.has(file.pageId)) {
      // Duplicate views disagree about the same page. There is no single live
      // source to search or navigate against, so omit it until they converge.
      unloadedFileCount += 1;
      truncated = true;
      continue;
    }
    const liveText = liveByPageId.get(file.pageId);
    const text =
      liveText ?? codeWorkspacePageSearchText(pageById.get(file.pageId));
    if (text === null || text === undefined) {
      unloadedFileCount += 1;
      truncated = true;
      continue;
    }
    if (
      text.length > maxFileCodeUnits ||
      totalCodeUnits + text.length > maxTotalCodeUnits
    ) {
      oversizedFileCount += 1;
      truncated = true;
      continue;
    }
    buffers.push({
      id: `page:${file.pageId}`,
      label: file.path,
      text,
    });
    totalCodeUnits += text.length;
  }

  return {
    buffers,
    totalCodeUnits,
    unloadedFileCount,
    oversizedFileCount,
    truncated,
  };
}
