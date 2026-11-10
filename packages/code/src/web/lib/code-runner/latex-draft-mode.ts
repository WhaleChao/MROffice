/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The LaTeX compile options a page keeps in this browser, each of which
 * makes a compile faster by asking TeX for less work:
 *
 * - Fast [draft] mode: images are drawn as boxes of their size instead of
 *   being read and placed, so a document with many or large images compiles
 *   in a fraction of the time (Overleaf's compile mode of the same name).
 * - Cache TikZ figures: with pdfLaTeX, every TikZ or pgfplots picture is
 *   drawn once into a figure of its own (TikZ's external library) and reused
 *   while its code is the same, so an edit to the text no longer redraws
 *   every plot. A figure that reads a data file is not redrawn when only the
 *   file changes; Recompile from scratch draws everything again.
 *
 * Both are written at the start of the \documentclass line itself so that no
 * line moves: TeX's error lines and the SyncTeX map still point at the page's
 * own lines, and a `% !TEX program` line above stays where the compile looks
 * for it.
 */

const DRAFT_OPTIONS =
  "\\PassOptionsToPackage{draft}{graphicx}\\PassOptionsToPackage{draft}{graphics}";
// Once the preamble is read, so after the document loads TikZ (pgfplots
// does too), and only under pdfLaTeX: the external library compiles each
// figure with pdflatex. An explicit empty option keeps \tikzexternalize from
// looking past \fi for one; calling it again in a document that already
// externalizes is a no-op.
const FIGURE_CACHE =
  "\\AddToHook{begindocument/before}{\\IfPackageLoadedTF{tikz}{\\ifx\\pdftexversion\\undefined" +
  "\\else\\usetikzlibrary{external}\\tikzexternalize[]\\fi}{}}";
/** A \documentclass (or LaTeX 2.09's \documentstyle) opening its own line. */
const DOCUMENT_CLASS_LINE = /^([ \t]*)\\document(?:class|style)\b/m;
const DRAFT_PREFIX = "mrnote.code-editor.latexDraft:";
const FIGURES_PREFIX = "mrnote.code-editor.latexFigureCache:";

/** The source a compile with these options reads; unchanged when it has no class line. */
export function latexCompileSource(
  source: string,
  options: { readonly draft: boolean; readonly figures: boolean },
): string {
  const prefix = (options.draft ? DRAFT_OPTIONS : "") + (options.figures ? FIGURE_CACHE : "");
  if (!prefix) return source;
  const match = DOCUMENT_CLASS_LINE.exec(source);
  if (!match) return source;
  const at = match.index + match[1].length;
  return source.slice(0, at) + prefix + source.slice(at);
}

/** The source a draft compile reads; unchanged when it has no class line. */
export function latexDraftSource(source: string): string {
  return latexCompileSource(source, { draft: true, figures: false });
}

const listeners = new Set<() => void>();

function readOption(prefix: string, pageId: string): boolean {
  try {
    return globalThis.localStorage?.getItem(`${prefix}${pageId}`) === "true";
  } catch {
    return false;
  }
}

function writeOption(prefix: string, pageId: string, enabled: boolean): void {
  try {
    const key = `${prefix}${pageId}`;
    if (enabled) globalThis.localStorage?.setItem(key, "true");
    else globalThis.localStorage?.removeItem(key);
  } catch {
    // Storage refused: the choice lasts until this tab reloads.
  }
  for (const listener of [...listeners]) listener();
}

/** Whether this page compiles in Fast [draft] mode (default Normal). */
export function latexDraftModeEnabled(pageId: string): boolean {
  return readOption(DRAFT_PREFIX, pageId);
}

export function setLatexDraftMode(pageId: string, enabled: boolean): void {
  writeOption(DRAFT_PREFIX, pageId, enabled);
}

/** Whether this page caches its TikZ figures (default off). */
export function latexFigureCacheEnabled(pageId: string): boolean {
  return readOption(FIGURES_PREFIX, pageId);
}

export function setLatexFigureCache(pageId: string, enabled: boolean): void {
  writeOption(FIGURES_PREFIX, pageId, enabled);
}

/** Every Code window of the page follows them, in this tab and the member's others. */
export function subscribeLatexDraftMode(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key?.startsWith(DRAFT_PREFIX) || event.key?.startsWith(FIGURES_PREFIX)) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}
