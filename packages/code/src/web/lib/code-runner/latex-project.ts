/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A LaTeX project, as Overleaf compiles one: the main file and every other
 * file of the Code workspace under the main file's folder (chapters,
 * bibliographies, classes and styles, data), sent with the compile at their
 * paths relative to that folder. The sandbox writes them beside main.tex.
 * Text files travel as typed; a binary file (a figure, a PDF) travels by its
 * page attachment's id, and the API reads its bytes after checking the member
 * may (workspace-binary-file.ts).
 *
 * Paths follow the executor's rules (internal/protocol/project_files.go and
 * the API's contracts.py): letters, digits and . _ + - per segment, at most 8
 * deep, never a file the compile itself owns.
 */

export const MAX_LATEX_PROJECT_FILES = 200;
/** Text files travel as typed; the API refuses more than this. */
export const MAX_LATEX_PROJECT_TEXT_BYTES = 2_000_000;
/** Text and binary files together (the API's MAX_PROJECT_FILE_BYTES). */
export const MAX_LATEX_PROJECT_BYTES = 8_000_000;

const SEGMENT = /^[A-Za-z0-9][A-Za-z0-9._+-]{0,127}$/;
const COMPILE_OWNED =
  /^main\.(aux|bbl|bcf|blg|toc|lof|lot|loa|lol|out|fdb_latexmk|run\.xml|nav|snm|vrb|idx|ind|ilg|glo|gls|glg|acn|acr|alg|ist)$/;
/** The API's attachment ids (BrowserProjectAsset.attachment_id is a UUID). */
const ATTACHMENT_ID =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export interface LatexProjectSource {
  readonly pageId: string;
  /** The file's Code workspace path ("paper/chapters/intro.tex"). */
  readonly path: string;
  readonly source: string;
  /** Present for a binary file: its attachment, sent in place of `source`. */
  readonly binaryAsset?: {
    readonly attachmentId: string;
    readonly sizeBytes: number;
  } | null;
}

export interface LatexProjectFile {
  /** Relative to the main file's folder ("chapters/intro.tex"). */
  readonly path: string;
  readonly content: string;
}

export interface LatexProjectAsset {
  /** Relative to the main file's folder ("figures/plot.png"). */
  readonly path: string;
  readonly attachmentId: string;
}

export interface LatexProject {
  readonly files: readonly LatexProjectFile[];
  readonly assets: readonly LatexProjectAsset[];
  /** Files under the folder whose names a compile cannot carry. */
  readonly skipped: readonly string[];
  /** The files exceed what one compile may send. */
  readonly tooLarge: boolean;
}

/** Whether a language compiles a multi-file project. */
export function compilesLatexProject(language: string): boolean {
  return language === "latex";
}

export function isLatexProjectPath(path: string): boolean {
  if (!path || new TextEncoder().encode(path).length > 255) return false;
  const segments = path.split("/");
  if (segments.length > 8 || !segments.every((segment) => SEGMENT.test(segment))) {
    return false;
  }
  if (path === "main.tex" || path === "main.pdf") return false;
  return !(segments.length === 1 && COMPILE_OWNED.test(path));
}

export function isLatexProjectAssetId(attachmentId: string): boolean {
  return ATTACHMENT_ID.test(attachmentId);
}

function folderOf(path: string): string {
  const index = path.lastIndexOf("/");
  return index < 0 ? "" : path.slice(0, index);
}

/** The sandbox folder the compile runs in, where the project is written. */
const COMPILE_ROOT = "/workspace/";

/**
 * A path the compile itself names (TeX's log: `./chapters/intro.tex`;
 * SyncTeX's inputs: `/workspace/./chapters/intro.tex`) as the project path
 * it is ("chapters/intro.tex", "main.tex"), or null for a file outside the
 * project (a TeX Live class or package).
 */
export function compiledProjectPath(path: string): string | null {
  let rest = path.trim();
  if (rest.startsWith(COMPILE_ROOT)) rest = rest.slice(COMPILE_ROOT.length);
  else if (rest.startsWith("/")) return null;
  const segments: string[] = [];
  for (const segment of rest.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (segments.length === 0) return null;
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.length > 0 ? segments.join("/") : null;
}

/**
 * The Code workspace path of a project file of the main file at `mainPath`.
 * The compile writes the main file as main.tex, whatever it is called here.
 */
export function latexProjectWorkspacePath(mainPath: string, projectPath: string): string {
  if (projectPath === "main.tex") return mainPath;
  const root = folderOf(mainPath);
  return root ? `${root}/${projectPath}` : projectPath;
}

/**
 * The project path a workspace file has in the project of the main file at
 * `mainPath` ("main.tex" for the main file itself), or null when the compile
 * does not send it.
 */
export function latexWorkspaceProjectPath(
  mainPath: string,
  workspacePath: string,
): string | null {
  if (workspacePath === mainPath) return "main.tex";
  const root = folderOf(mainPath);
  const prefix = root ? `${root}/` : "";
  if (!workspacePath.startsWith(prefix)) return null;
  const path = workspacePath.slice(prefix.length);
  return isLatexProjectPath(path) ? path : null;
}

function byPath(left: { path: string }, right: { path: string }): number {
  return left.path < right.path ? -1 : left.path > right.path ? 1 : 0;
}

/**
 * The project of the main file at `mainPath`: every other workspace file under
 * its folder, text files and binary files apart, each sorted by path.
 */
export function assembleLatexProject(
  mainPageId: string,
  mainPath: string,
  sources: readonly LatexProjectSource[],
): LatexProject {
  const root = folderOf(mainPath);
  const prefix = root ? `${root}/` : "";
  const files: LatexProjectFile[] = [];
  const assets: LatexProjectAsset[] = [];
  const skipped: string[] = [];
  let textBytes = 0;
  let assetBytes = 0;
  const encoder = new TextEncoder();
  for (const source of sources) {
    if (source.pageId === mainPageId || !source.path.startsWith(prefix)) continue;
    const path = source.path.slice(prefix.length);
    if (!isLatexProjectPath(path)) {
      skipped.push(path);
      continue;
    }
    if (source.binaryAsset) {
      assetBytes += source.binaryAsset.sizeBytes;
      assets.push({ path, attachmentId: source.binaryAsset.attachmentId });
      continue;
    }
    textBytes += encoder.encode(source.source).length;
    files.push({ path, content: source.source });
  }
  files.sort(byPath);
  assets.sort(byPath);
  skipped.sort();
  return {
    files,
    assets,
    skipped,
    tooLarge:
      files.length + assets.length > MAX_LATEX_PROJECT_FILES ||
      textBytes > MAX_LATEX_PROJECT_TEXT_BYTES ||
      textBytes + assetBytes > MAX_LATEX_PROJECT_BYTES,
  };
}
