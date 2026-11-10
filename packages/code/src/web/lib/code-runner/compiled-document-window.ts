/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The PDF a LaTeX compile produced, held in memory for the canvas window that
 * shows it beside its Code window.
 *
 * Entries are keyed by the Code window's id. The Code window and the PDF
 * window each retain the key while they are mounted, so closing either one
 * keeps the document for the other, and the bytes are dropped once neither
 * is open. Nothing here is persisted: a reload starts without a document,
 * which is why the PDF window is an ephemeral window type.
 *
 * Beside the document an entry keeps what the compile said about its project
 * (the log and SyncTeX map, the main file's path, the files it read), so the
 * PDF window can go from a point on a page to its source line, and a Code
 * window can ask it to show where a line is typeset.
 */

import type { LanguageExecutionArtifact } from "./execution-artifact";
import type { LanguageExecutionOutputFiles } from "./output-files";
import type { SynctexRect } from "./synctex";

/** A project file the compile read, to go to from the PDF (SyncTeX). */
export interface CompiledDocumentSource {
  /** Its Code workspace path ("paper/chapters/intro.tex"). */
  readonly path: string;
  readonly pageId: string;
  readonly title: string;
}

/** What the compile knew about its project beside the document. */
export interface CompiledDocumentProject {
  /** The log and SyncTeX map the compile returned with this document. */
  readonly outputFiles: LanguageExecutionOutputFiles | null;
  /** The main file's Code workspace path; SyncTeX's paths are relative to
   * its folder. */
  readonly mainPath: string;
  readonly sources: readonly CompiledDocumentSource[];
}

/** A place in the document to show (a SyncTeX forward search). */
export interface CompiledDocumentReveal {
  /** New for every request, so asking for the same place again shows it again. */
  readonly id: number;
  /** When it was asked for (Date.now()): a window opened much later does not
   * replay it. */
  readonly requestedAt: number;
  readonly page: number;
  /** PDF points from the page's top left. */
  readonly rects: readonly SynctexRect[];
}

export interface CompiledDocumentWindowEntry {
  readonly artifact: LanguageExecutionArtifact;
  /** Its name in the window title and its download ("paper.pdf"). */
  readonly filename: string;
  /** The notebook of the Code page that compiled it, where Open in PDF
   * editor saves it; null when the Code window has none. */
  readonly notebookId: string | null;
  /** The PDF page this document was saved as, so opening it in the PDF
   * editor again opens that page instead of saving a second copy. */
  readonly savedPage: { readonly pageId: string; readonly title: string } | null;
  /**
   * A later compile of the same Code window failed: the window still shows
   * this document, the last one that compiled, and says so.
   */
  readonly stale: boolean;
  readonly project: CompiledDocumentProject | null;
  /** The last place a member asked to see; null after a new compile. */
  readonly reveal: CompiledDocumentReveal | null;
  /** When (Date.now()) the Code window's next compile started; null while
   * none runs. The window says so and keeps this document up meanwhile. */
  readonly compilingSince: number | null;
  /** How long the compile that produced this document took, in ms. */
  readonly compileMs: number | null;
}

const entries = new Map<string, CompiledDocumentWindowEntry>();
const retainers = new Map<string, number>();
const listeners = new Map<string, Set<() => void>>();
let revealSequence = 0;

function notify(key: string): void {
  listeners.get(key)?.forEach((listener) => listener());
}

export function publishCompiledDocument(
  key: string,
  artifact: LanguageExecutionArtifact,
  filename: string,
  notebookId: string | null,
  project: CompiledDocumentProject | null = null,
  compileMs: number | null = null,
): void {
  entries.set(key, {
    artifact,
    filename,
    notebookId,
    savedPage: null,
    stale: false,
    project,
    reveal: null,
    compilingSince: null,
    compileMs,
  });
  notify(key);
}

/**
 * The Code window behind `key` started compiling at `since` (Date.now()), or
 * stopped (null). Nothing changes when no document is kept: there is no
 * window to tell.
 */
export function markCompiledDocumentCompiling(
  key: string,
  since: number | null,
): void {
  const entry = entries.get(key);
  if (!entry || entry.compilingSince === since) return;
  entries.set(key, { ...entry, compilingSince: since });
  notify(key);
}

/**
 * Ask the window showing `key`'s document to bring `page` into view and mark
 * `rects` on it. False when no document is kept for `key`.
 */
export function revealInCompiledDocument(
  key: string,
  page: number,
  rects: readonly SynctexRect[],
): boolean {
  const entry = entries.get(key);
  if (!entry) return false;
  revealSequence += 1;
  entries.set(key, {
    ...entry,
    reveal: { id: revealSequence, requestedAt: Date.now(), page, rects },
  });
  notify(key);
  return true;
}

/** Every document kept now, by key: where a source line may be typeset. */
export function listCompiledDocuments(): Array<
  readonly [string, CompiledDocumentWindowEntry]
> {
  return [...entries.entries()];
}

/** Record the PDF page `artifact` was saved as (Open in PDF editor). */
export function markCompiledDocumentSaved(
  key: string,
  artifact: LanguageExecutionArtifact,
  savedPage: { readonly pageId: string; readonly title: string },
): void {
  const entry = entries.get(key);
  if (!entry || entry.artifact !== artifact) return;
  entries.set(key, { ...entry, savedPage });
  notify(key);
}

export function markCompiledDocumentStale(key: string): void {
  const entry = entries.get(key);
  if (!entry || (entry.stale && entry.compilingSince === null)) return;
  entries.set(key, { ...entry, stale: true, compilingSince: null });
  notify(key);
}

export function readCompiledDocument(
  key: string,
): CompiledDocumentWindowEntry | null {
  return entries.get(key) ?? null;
}

export function subscribeCompiledDocument(
  key: string,
  listener: () => void,
): () => void {
  const set = listeners.get(key) ?? new Set<() => void>();
  set.add(listener);
  listeners.set(key, set);
  return () => {
    set.delete(listener);
    if (!set.size) listeners.delete(key);
  };
}

/** Keep `key`'s document while the caller is mounted; returns the release. */
export function retainCompiledDocument(key: string): () => void {
  retainers.set(key, (retainers.get(key) ?? 0) + 1);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    const remaining = (retainers.get(key) ?? 1) - 1;
    if (remaining > 0) {
      retainers.set(key, remaining);
      return;
    }
    retainers.delete(key);
    if (entries.delete(key)) notify(key);
  };
}
