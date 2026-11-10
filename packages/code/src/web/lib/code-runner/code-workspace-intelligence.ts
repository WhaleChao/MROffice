/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Project authorized notebook code pages into the standalone CodeWindow
 * intelligence protocol.
 *
 * Search hydration already gives the renderer a bounded, notebook-scoped
 * source corpus. This adapter reuses that corpus for compiler-backed
 * definition/reference queries while keeping page ids as stable navigation
 * targets. Open buffers override saved text, path and language so unsaved edits
 * remain authoritative.
 */

import type {
  CodeWorkspaceFile,
  CodeWorkspacePageLike,
} from "./code-workspace";
import {
  codeWorkspacePathIdentity,
  normalizeCodeWorkspacePath,
} from "./code-workspace";
import {
  buildCodeWorkspaceSearchProjection,
  type CodeWorkspaceSearchProjection,
} from "./code-workspace-loader";
import type { CodeWindowBufferText } from "./code-window-intelligence";

const WORKSPACE_PAGE_BUFFER_PREFIX = "page:";

export interface CodeWorkspaceIntelligenceLiveBuffer {
  readonly pageId: string;
  readonly path: string;
  readonly language: string;
  readonly text: string;
}

export interface CodeWorkspaceIntelligenceTarget {
  readonly bufferId: string;
  readonly pageId: string;
  readonly file: CodeWorkspaceFile;
  readonly expectedSource: string;
}

export interface CodeWorkspaceIntelligenceProjection
  extends Pick<
    CodeWorkspaceSearchProjection,
    | "totalCodeUnits"
    | "unloadedFileCount"
    | "oversizedFileCount"
    | "truncated"
  > {
  readonly activeBufferId: string;
  readonly buffers: CodeWindowBufferText[];
  readonly targets: CodeWorkspaceIntelligenceTarget[];
}

export function codeWorkspaceIntelligenceBufferId(pageId: string): string {
  return `${WORKSPACE_PAGE_BUFFER_PREFIX}${pageId}`;
}

export function codeWorkspaceIntelligencePageId(
  bufferId: string,
): string | null {
  return bufferId.startsWith(WORKSPACE_PAGE_BUFFER_PREFIX)
    ? bufferId.slice(WORKSPACE_PAGE_BUFFER_PREFIX.length)
    : null;
}

/**
 * Build an active-first compiler input from the same bounded projection used
 * by Find in Files. Duplicate case-insensitive virtual paths are ambiguous, so
 * only the active file is retained from a conflicting group; unique siblings
 * remain available.
 */
export function buildCodeWorkspaceIntelligenceProjection(
  files: readonly CodeWorkspaceFile[],
  pages: readonly CodeWorkspacePageLike[],
  liveBuffers: readonly CodeWorkspaceIntelligenceLiveBuffer[],
  activePageId: string,
): CodeWorkspaceIntelligenceProjection {
  const liveByPageId = new Map<string, CodeWorkspaceIntelligenceLiveBuffer>();
  const conflictedLivePageIds = new Set<string>();
  for (const live of liveBuffers) {
    const existing = liveByPageId.get(live.pageId);
    if (!existing) {
      liveByPageId.set(live.pageId, live);
    } else if (
      existing.text !== live.text ||
      existing.path !== live.path ||
      existing.language !== live.language
    ) {
      conflictedLivePageIds.add(live.pageId);
    }
  }

  const effectiveFiles = files.map((file) => {
    const live = liveByPageId.get(file.pageId);
    const livePath = normalizeCodeWorkspacePath(live?.path);
    return live
      ? {
          ...file,
          path: livePath ?? file.path,
          language: live.language,
        }
      : file;
  });
  const codeFiles = effectiveFiles.filter((file) => file.pageType === "code");
  const pathCounts = new Map<string, number>();
  for (const file of codeFiles) {
    const key = codeWorkspacePathIdentity(file.path);
    pathCounts.set(key, (pathCounts.get(key) ?? 0) + 1);
  }
  const eligibleFiles = codeFiles.filter(
    (file) =>
      !conflictedLivePageIds.has(file.pageId) &&
      ((pathCounts.get(codeWorkspacePathIdentity(file.path)) ?? 0) === 1 ||
        file.pageId === activePageId),
  );
  const search = buildCodeWorkspaceSearchProjection(
    eligibleFiles,
    pages,
    liveBuffers.map((live) => ({
      pageId: live.pageId,
      text: live.text,
    })),
  );
  const fileByPageId = new Map(
    eligibleFiles.map((file) => [file.pageId, file] as const),
  );
  const targets: CodeWorkspaceIntelligenceTarget[] = [];
  const buffers: CodeWindowBufferText[] = [];

  for (const searchBuffer of search.buffers) {
    const pageId = codeWorkspaceIntelligencePageId(searchBuffer.id);
    const file = pageId ? fileByPageId.get(pageId) : undefined;
    if (!pageId || !file) continue;
    const bufferId = codeWorkspaceIntelligenceBufferId(pageId);
    buffers.push({
      bufferId,
      label: file.path,
      language: file.language,
      source: searchBuffer.text,
    });
    targets.push({
      bufferId,
      pageId,
      file,
      expectedSource: searchBuffer.text,
    });
  }

  const activeBufferId = codeWorkspaceIntelligenceBufferId(activePageId);
  const activeIndex = buffers.findIndex(
    (buffer) => buffer.bufferId === activeBufferId,
  );
  if (activeIndex > 0) {
    buffers.unshift(buffers.splice(activeIndex, 1)[0]);
    targets.unshift(targets.splice(activeIndex, 1)[0]);
  }

  return {
    activeBufferId,
    buffers,
    targets,
    totalCodeUnits: search.totalCodeUnits,
    unloadedFileCount: search.unloadedFileCount,
    oversizedFileCount: search.oversizedFileCount,
    truncated:
      search.truncated ||
      conflictedLivePageIds.size > 0 ||
      eligibleFiles.length < codeFiles.length,
  };
}
