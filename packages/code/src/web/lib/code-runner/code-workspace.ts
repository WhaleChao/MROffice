/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Notebook-scoped virtual workspace helpers for native code pages.
 *
 * The workspace never reads an ambient host directory. Its catalog is derived
 * from pages the current notebook API already authorizes, and a code page may
 * persist one bounded relative path inside its existing content_json.
 */

import {
  codeExportFilename,
  extensionFromCodeFilename,
  normalizeCodeSourceFilename,
} from "./code-file-import";
import { languageForExtension } from "@mrnote/core/web/lib/code-runner/languages.ts";

export const CODE_WORKSPACE_REFERENCE_KEY = "codeWorkspace";
export const CODE_WORKSPACE_REFERENCE_VERSION = 1;
export const CODE_WORKSPACE_CATALOG_LIMIT = 10_000;
export const CODE_WORKSPACE_TREE_RENDER_LIMIT = 500;
export const CODE_WORKSPACE_TREE_BUILD_NODE_LIMIT = 20_000;
export const MAX_CODE_WORKSPACE_PATH_CHARS = 1_024;
export const MAX_CODE_WORKSPACE_SEGMENT_CHARS = 255;
export const MAX_CODE_WORKSPACE_DEPTH = 100;

/** Locale-independent case-folded identity for virtual workspace paths. */
export function codeWorkspacePathIdentity(path: string): string {
  return path.toLowerCase();
}

const INVALID_PATH_SEGMENT =
  /[\u0000-\u001f\u007f-\u009f\u200e\u200f\u2028\u2029\u202a-\u202e\u2066-\u2069<>:"|?*]/;

export interface PersistedCodeWorkspaceReference {
  readonly version: typeof CODE_WORKSPACE_REFERENCE_VERSION;
  readonly path: string;
}

export interface CodeWorkspacePageLike {
  readonly id: string;
  readonly notebook_id?: string;
  readonly title?: string;
  readonly page_type?: string;
  readonly content_version?: number;
  readonly content_json?: Record<string, unknown>;
}

export type CodeWorkspacePageType = "code";

export interface CodeWorkspaceFile {
  readonly pageId: string;
  readonly notebookId: string;
  readonly pageType: CodeWorkspacePageType;
  readonly title: string;
  readonly path: string;
  readonly language: string;
  readonly contentVersion: number;
  readonly current: boolean;
  readonly pathConflict: boolean;
}

export interface CodeWorkspaceCatalog {
  readonly files: CodeWorkspaceFile[];
  readonly total: number;
  readonly inspected: number;
  readonly truncated: boolean;
  readonly conflictCount: number;
}

export interface CodeWorkspaceTreeRow {
  readonly id: string;
  readonly kind: "directory" | "file";
  readonly name: string;
  readonly path: string;
  readonly depth: number;
  readonly file?: CodeWorkspaceFile;
}

export interface CodeWorkspaceTree {
  readonly rows: CodeWorkspaceTreeRow[];
  readonly totalRows: number;
  /** False when the build ceiling makes totalRows a safe lower bound. */
  readonly totalRowsExact: boolean;
  readonly truncated: boolean;
}

export interface CodeWorkspaceLocation {
  readonly line: number;
  readonly column: number;
  readonly from?: number;
  readonly to?: number;
  readonly contentVersion?: number;
}

interface MutableTreeDirectory {
  readonly name: string;
  readonly path: string;
  readonly directories: Map<string, MutableTreeDirectory>;
  readonly files: CodeWorkspaceFile[];
}

/** Normalize an untrusted path without silently resolving aliases. */
export function normalizeCodeWorkspacePath(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const normalizedSeparators = value
    .trim()
    .normalize("NFC")
    .replaceAll("\\", "/");
  if (
    normalizedSeparators.length === 0 ||
    normalizedSeparators.length > MAX_CODE_WORKSPACE_PATH_CHARS ||
    normalizedSeparators.startsWith("/") ||
    normalizedSeparators.endsWith("/") ||
    /^[A-Za-z]:/.test(normalizedSeparators)
  ) {
    return null;
  }

  const segments = normalizedSeparators.split("/");
  if (
    segments.length === 0 ||
    segments.length > MAX_CODE_WORKSPACE_DEPTH
  ) {
    return null;
  }
  for (const segment of segments) {
    if (
      segment.length === 0 ||
      segment.length > MAX_CODE_WORKSPACE_SEGMENT_CHARS ||
      segment === "." ||
      segment === ".." ||
      segment !== segment.trim() ||
      segment.endsWith(".") ||
      normalizeCodeSourceFilename(segment) !== segment ||
      INVALID_PATH_SEGMENT.test(segment)
    ) {
      return null;
    }
  }
  return segments.join("/");
}

const NAME_BRACKETS: ReadonlyMap<string, string> = new Map([
  ["(", ")"],
  ["（", "）"],
  ["[", "]"],
  ["【", "】"],
  ["{", "}"],
  ["「", "」"],
  ["《", "》"],
]);
const NAME_CLOSING_BRACKETS: ReadonlySet<string> = new Set(
  NAME_BRACKETS.values(),
);

function separatorInsideBrackets(value: string): boolean {
  let depth = 0;
  for (const character of value) {
    if (NAME_BRACKETS.has(character)) depth += 1;
    else if (NAME_CLOSING_BRACKETS.has(character)) depth = Math.max(0, depth - 1);
    else if ((character === "/" || character === "\\") && depth > 0) return true;
  }
  return false;
}

/**
 * Read the text typed into the Code ribbon's file-name box as a name.
 *
 * The box still accepts a workspace path ("src/app.ts" moves the page into
 * src/). Text reads as a page title instead when its last segment carries no
 * file extension or a separator sits inside brackets, as in
 * "番茄钟（25/5 交替）" or "Timer (25/5).html". A title's "/" and "\" are part
 * of the name and become "-", exactly as codeExportFilename derives a file
 * name from a page title, so the title and the workspace path agree.
 *
 * When the page already lives in a folder and the text still starts with that
 * folder, the folder is kept and only the rest is read as the name.
 *
 * Returns null for ordinary path input (the caller keeps its path rules).
 */
export function codeRenameInputAsTitle(
  raw: string,
  currentPath = "",
): { readonly path: string; readonly title: string } | null {
  const typed = raw.trim().normalize("NFC");
  const slash = currentPath.lastIndexOf("/");
  const folder = slash > 0 ? currentPath.slice(0, slash + 1) : "";
  const keepsFolder =
    folder.length > 0 && typed.replaceAll("\\", "/").startsWith(folder);
  const name = keepsFolder ? typed.slice(folder.length).trim() : typed;
  if (!/[\\/]/.test(name)) return null;
  const lastSegment = name.slice(
    Math.max(name.lastIndexOf("/"), name.lastIndexOf("\\")) + 1,
  );
  const extension = extensionFromCodeFilename(lastSegment);
  if (extension && !separatorInsideBrackets(name)) return null;
  const title =
    extension && languageForExtension(extension)
      ? name.slice(0, name.length - extension.length).trim() || name
      : name;
  return {
    path: `${keepsFolder ? folder : ""}${name.replace(/[\\/]+/g, "-")}`,
    title,
  };
}








export function codePageTitleForPath(path: string): string | null {
  const normalized = normalizeCodeWorkspacePath(path);
  if (!normalized) return null;
  const name = normalized.slice(normalized.lastIndexOf("/") + 1);
  const extension = extensionFromCodeFilename(name);
  const title =
    extension && languageForExtension(extension)
      ? name.slice(0, name.length - extension.length).trim()
      : name.trim();
  return title || null;
}

export function codeWorkspaceBasename(path: string): string {
  return path.slice(path.lastIndexOf("/") + 1);
}

export function replaceCodeWorkspaceBasename(
  path: string,
  basename: string,
): string | null {
  const normalizedPath = normalizeCodeWorkspacePath(path);
  const normalizedBasename = normalizeCodeSourceFilename(basename);
  if (!normalizedPath || !normalizedBasename) return null;
  const slash = normalizedPath.lastIndexOf("/");
  return normalizeCodeWorkspacePath(
    slash < 0
      ? normalizedBasename
      : `${normalizedPath.slice(0, slash + 1)}${normalizedBasename}`,
  );
}

export function persistedCodeWorkspaceReference(
  path: unknown,
): PersistedCodeWorkspaceReference | null {
  const normalized = normalizeCodeWorkspacePath(path);
  return normalized
    ? {
        version: CODE_WORKSPACE_REFERENCE_VERSION,
        path: normalized,
      }
    : null;
}

export function restoreCodeWorkspacePath(
  contentJson: unknown,
): string | null {
  if (
    !contentJson ||
    typeof contentJson !== "object" ||
    Array.isArray(contentJson)
  ) {
    return null;
  }
  const raw = (contentJson as Record<string, unknown>)[
    CODE_WORKSPACE_REFERENCE_KEY
  ];
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const reference = raw as Record<string, unknown>;
  if (reference.version !== CODE_WORKSPACE_REFERENCE_VERSION) return null;
  return normalizeCodeWorkspacePath(reference.path);
}

function pageTypeOf(page: CodeWorkspacePageLike): CodeWorkspacePageType | null {
  return page.page_type === "code" ? page.page_type : null;
}

function codeLanguage(page: CodeWorkspacePageLike): string {
  const raw = page.content_json?.language;
  return typeof raw === "string" && raw.trim() ? raw.slice(0, 64) : "javascript";
}

/** Resolve a durable or legacy page into one safe workspace-relative path. */
export function codeWorkspacePathForPage(
  page: CodeWorkspacePageLike,
): string | null {
  const pageType = pageTypeOf(page);
  if (!pageType) return null;
  const persisted = restoreCodeWorkspacePath(page.content_json);
  if (persisted) return persisted;

  return normalizeCodeWorkspacePath(
    codeExportFilename(
      page.content_json?.sourceFileName,
      page.title || page.id,
      codeLanguage(page),
    ),
  );
}

/**
 * Build a bounded deterministic workspace catalog.
 *
 * Duplicate normalized paths remain visible and are marked as conflicts rather
 * than silently renamed; page ids stay the stable navigation identity.
 */
export function buildCodeWorkspaceCatalog(
  pages: readonly CodeWorkspacePageLike[],
  currentPageId?: string,
): CodeWorkspaceCatalog {
  const inspected = Math.min(pages.length, CODE_WORKSPACE_CATALOG_LIMIT);
  const candidates: Omit<CodeWorkspaceFile, "pathConflict">[] = [];
  for (let index = 0; index < inspected; index += 1) {
    const page = pages[index];
    const pageType = pageTypeOf(page);
    const path = pageType ? codeWorkspacePathForPage(page) : null;
    if (!pageType || !path) continue;
    candidates.push({
      pageId: page.id,
      notebookId: page.notebook_id ?? "",
      pageType,
      title: (page.title || path).slice(0, 512),
      path,
      language: codeLanguage(page),
      contentVersion:
        typeof page.content_version === "number" &&
        Number.isFinite(page.content_version)
          ? Math.max(0, Math.trunc(page.content_version))
          : 0,
      current: page.id === currentPageId,
    });
  }

  candidates.sort(
    (left, right) =>
      Number(right.current) - Number(left.current) ||
      left.path.localeCompare(right.path, undefined, {
        numeric: true,
        sensitivity: "base",
      }) ||
      left.pageId.localeCompare(right.pageId),
  );

  const pathCounts = new Map<string, number>();
  for (const file of candidates) {
    // Conflict identity must not depend on the device locale (notably Turkish
    // I/i rules), otherwise collaborators can disagree about the same path.
    const key = codeWorkspacePathIdentity(file.path);
    pathCounts.set(key, (pathCounts.get(key) ?? 0) + 1);
  }
  const files = candidates.map((file) => ({
    ...file,
    pathConflict:
      (pathCounts.get(codeWorkspacePathIdentity(file.path)) ?? 0) > 1,
  }));
  const conflictCount = files.reduce(
    (count, file) => count + Number(file.pathConflict),
    0,
  );
  return {
    files,
    total: pages.length,
    inspected,
    truncated: pages.length > inspected,
    conflictCount,
  };
}

function makeDirectory(name: string, path: string): MutableTreeDirectory {
  return {
    name,
    path,
    directories: new Map(),
    files: [],
  };
}

/**
 * Flatten a path tree for an accessible, bounded Explorer list.
 *
 * Both construction and rendering are capped. A file is admitted atomically
 * with every missing ancestor, so reaching the build ceiling never leaves an
 * empty partial path behind. Collapsed paths omit descendants before the
 * render cap is applied.
 */
export function buildCodeWorkspaceTree(
  files: readonly CodeWorkspaceFile[],
  collapsedPaths: ReadonlySet<string> = new Set(),
): CodeWorkspaceTree {
  const root = makeDirectory("", "");
  const sourceFiles = files.slice(0, CODE_WORKSPACE_CATALOG_LIMIT);
  let builtNodes = 0;
  let buildTruncated = files.length > sourceFiles.length;
  for (const file of sourceFiles) {
    const segments = file.path.split("/");
    let probe = root;
    let missingDirectories = 0;
    for (let index = 0; index < segments.length - 1; index += 1) {
      const child = probe.directories.get(segments[index]);
      if (!child) {
        missingDirectories = segments.length - 1 - index;
        break;
      }
      probe = child;
    }
    if (
      builtNodes + missingDirectories + 1 >
      CODE_WORKSPACE_TREE_BUILD_NODE_LIMIT
    ) {
      buildTruncated = true;
      continue;
    }
    let directory = root;
    for (let index = 0; index < segments.length - 1; index += 1) {
      const name = segments[index];
      const path = directory.path ? `${directory.path}/${name}` : name;
      let child = directory.directories.get(name);
      if (!child) {
        child = makeDirectory(name, path);
        directory.directories.set(name, child);
        builtNodes += 1;
      }
      directory = child;
    }
    directory.files.push(file);
    builtNodes += 1;
  }

  const allRows: CodeWorkspaceTreeRow[] = [];
  const visit = (directory: MutableTreeDirectory, depth: number) => {
    const childDirectories = [...directory.directories.values()].sort(
      (left, right) =>
        left.name.localeCompare(right.name, undefined, {
          numeric: true,
          sensitivity: "base",
        }),
    );
    for (const child of childDirectories) {
      allRows.push({
        id: `directory:${child.path}`,
        kind: "directory",
        name: child.name,
        path: child.path,
        depth,
      });
      if (!collapsedPaths.has(child.path)) visit(child, depth + 1);
    }
    const childFiles = [...directory.files].sort(
      (left, right) =>
        codeWorkspaceBasename(left.path).localeCompare(
          codeWorkspaceBasename(right.path),
          undefined,
          { numeric: true, sensitivity: "base" },
        ) || left.pageId.localeCompare(right.pageId),
    );
    for (const file of childFiles) {
      allRows.push({
        id: `file:${file.pageId}`,
        kind: "file",
        name: codeWorkspaceBasename(file.path),
        path: file.path,
        depth,
        file,
      });
    }
  };
  visit(root, 0);
  return {
    rows: allRows.slice(0, CODE_WORKSPACE_TREE_RENDER_LIMIT),
    totalRows: allRows.length,
    totalRowsExact: !buildTruncated,
    truncated:
      buildTruncated || allRows.length > CODE_WORKSPACE_TREE_RENDER_LIMIT,
  };
}

const pendingLocations = new Map<string, CodeWorkspaceLocation>();

export function rememberCodeWorkspaceLocation(
  pageId: string,
  location: CodeWorkspaceLocation,
): void {
  if (!pageId) return;
  const from =
    typeof location.from === "number" && Number.isFinite(location.from)
      ? Math.max(0, Math.trunc(location.from))
      : null;
  const to =
    from !== null &&
    typeof location.to === "number" &&
    Number.isFinite(location.to)
      ? Math.max(from, Math.trunc(location.to))
      : null;
  const contentVersion =
    typeof location.contentVersion === "number" &&
    Number.isFinite(location.contentVersion)
      ? Math.max(0, Math.trunc(location.contentVersion))
      : null;
  pendingLocations.set(pageId, {
    line: Math.max(1, Math.trunc(location.line) || 1),
    column: Math.max(1, Math.trunc(location.column) || 1),
    ...(from === null ? {} : { from }),
    ...(to === null ? {} : { to }),
    ...(contentVersion === null ? {} : { contentVersion }),
  });
}

export function takeCodeWorkspaceLocation(
  pageId: string,
): CodeWorkspaceLocation | null {
  const location = pendingLocations.get(pageId) ?? null;
  pendingLocations.delete(pageId);
  return location;
}

export function _clearCodeWorkspaceLocationsForTest(): void {
  pendingLocations.clear();
}
