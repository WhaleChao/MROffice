/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Overleaf's two jumps between a LaTeX source and its PDF, over the documents
 * the compiled-PDF registry keeps (lib/code-runner/compiled-document-window)
 * and the SyncTeX map each compile returned (lib/code-runner/synctex).
 *
 * A chapter's own Code window finds the document its main file compiled: any
 * kept document of the same notebook whose map read that file answers, the
 * window's own compile first.
 */

import {
  listCompiledDocuments,
  type CompiledDocumentSource,
  type CompiledDocumentWindowEntry,
} from "./compiled-document-window";
import {
  latexProjectWorkspacePath,
  latexWorkspaceProjectPath,
} from "./latex-project";
import { OUTPUT_SYNCTEX_NAME, outputFile } from "./output-files";
import {
  loadSynctexMap,
  synctexForward,
  synctexHasInput,
  synctexInverse,
  type SynctexMap,
  type SynctexRect,
} from "./synctex";

export type TypesetPlaceResult =
  | {
      readonly status: "found";
      /** The registry key of the document it is in. */
      readonly key: string;
      readonly page: number;
      readonly rects: readonly SynctexRect[];
    }
  /** No kept document of the notebook came with a SyncTeX map. */
  | { readonly status: "no-document" }
  /** Documents were compiled, but none typeset this file or line. */
  | { readonly status: "not-typeset" };

function entryMap(entry: CompiledDocumentWindowEntry): Promise<SynctexMap | null> | null {
  const synctex = outputFile(entry.project?.outputFiles, OUTPUT_SYNCTEX_NAME);
  return synctex ? loadSynctexMap(synctex.bytes) : null;
}

/** Where line `line` of the workspace file `workspacePath` is typeset. */
export async function findTypesetPlace({
  workspacePath,
  line,
  notebookId,
  preferKey,
}: {
  workspacePath: string;
  line: number;
  notebookId: string | null;
  /** The asking Code window's own document, tried first. */
  preferKey: string | null;
}): Promise<TypesetPlaceResult> {
  const documents = listCompiledDocuments()
    .filter(
      ([key, entry]) =>
        entry.project !== null &&
        (key === preferKey || (notebookId !== null && entry.notebookId === notebookId)),
    )
    .sort(([left], [right]) => Number(right === preferKey) - Number(left === preferKey));
  let anyMap = false;
  for (const [key, entry] of documents) {
    const project = entry.project!;
    const projectPath = latexWorkspaceProjectPath(project.mainPath, workspacePath);
    const pending = entryMap(entry);
    if (!projectPath || !pending) continue;
    const map = await pending;
    if (!map) continue;
    anyMap = true;
    if (!synctexHasInput(map, projectPath)) continue;
    const place = synctexForward(map, projectPath, line);
    if (place) return { status: "found", key, page: place.page, rects: place.rects };
  }
  return { status: anyMap ? "not-typeset" : "no-document" };
}

export interface SourceLine {
  readonly workspacePath: string;
  readonly line: number;
  /** The workspace file to open, when the compile read it from one. */
  readonly source: CompiledDocumentSource | null;
}

/** The source line under a point on page `page` of `entry`'s document. */
export async function findSourceLine(
  entry: CompiledDocumentWindowEntry,
  page: number,
  x: number,
  y: number,
): Promise<SourceLine | null> {
  const project = entry.project;
  const pending = entryMap(entry);
  if (!project || !pending) return null;
  const map = await pending;
  const found = map ? synctexInverse(map, page, x, y) : null;
  if (!found) return null;
  const workspacePath = latexProjectWorkspacePath(project.mainPath, found.file);
  return {
    workspacePath,
    line: found.line,
    source: project.sources.find((source) => source.path === workspacePath) ?? null,
  };
}

/** Whether `entry`'s document came with a SyncTeX map to jump through. */
export function compiledDocumentHasSynctex(
  entry: CompiledDocumentWindowEntry | null,
): boolean {
  return Boolean(outputFile(entry?.project?.outputFiles, OUTPUT_SYNCTEX_NAME));
}
