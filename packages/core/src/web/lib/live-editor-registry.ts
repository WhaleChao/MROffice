/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editors open on this page, by page id: what a caller outside an
 * editor's own component (the embedded editor's Automation API,
 * lib/embed/automation) reaches the live editor through.
 *
 * Each editor registers a handle while it is mounted and writable to by its
 * own commit path: the Docs editor its Tiptap editor (a command dispatches a
 * transaction, so Undo, the collaboration room and the autosave take it as
 * they take a keystroke); a workbook, a deck and a PDF a small adapter over
 * the same commit their ribbon's buttons use. Nothing here keeps state of its
 * own beyond the handle, and an unmounted editor is forgotten.
 */

import type { Editor } from "@tiptap/core";

export interface LiveDocumentEditor {
  readonly kind: "document";
  readonly editor: Editor;
}

/** A workbook's content as stored (the page's content_json, every cell). */
export type LiveContent = Record<string, unknown>;

/** A cell as the workbook stores it: what is typed (a formula's "=…"). */
export interface LiveCell {
  readonly value?: string;
  readonly formula?: string;
}

export interface LiveSheet {
  readonly id: string;
  readonly name: string;
  readonly hidden: boolean;
  readonly rowCount: number;
  readonly columnCount: number;
  /** "row:col" (0-based) -> the cell. */
  readonly cells: Readonly<Record<string, LiveCell>>;
}

export type LiveWriteOutcome = "written" | "read_only" | "protected" | "not_found";

export interface LiveSpreadsheetEditor {
  readonly kind: "spreadsheet";
  /** The sheets as the window holds them now. */
  sheets(): readonly LiveSheet[];
  /** The open sheet's id. */
  activeSheetId(): string | null;
  /** Whether the window takes edits now (not read only, loaded). */
  writable(): boolean;
  /** Whether the window has rendered its last committed change: a change is
   * the base of the next one, and of what is read back, only then. */
  settled(): boolean;
  /** A cell's computed value, as the grid shows it (the window's value store). */
  computed(sheetId: string, row: number, col: number): unknown;
  /** Values typed into cells of one sheet, as the formula bar enters them
   * (``asTyped``: "50%", dates and numbers read as Excel reads them typed;
   * "=…" is a formula), committed as one undoable step of the window's own
   * commit (history, collaboration room, autosave). */
  writeCells(sheetId: string, writes: readonly { row: number; col: number; value: string }[], asTyped: boolean): LiveWriteOutcome;
  /** A new sheet after the last (the window's Insert Sheet), named ``name``
   * when given; null when the workbook refused it. */
  addSheet(name?: string): { id: string; name: string } | null;
}

export interface LivePresentationEditor {
  readonly kind: "presentation";
  /** The deck as stored (the page's content_json). */
  read(): LiveContent;
  activeSlideIndex(): number;
  writable(): boolean;
  /** The deck with a new slide from ``layout`` after slide ``afterIndex``
   * (the ribbon's New Slide, the window's theme and masters), not committed. */
  withNewSlide(layout: string | undefined, afterIndex: number): { content: LiveContent; index: number } | null;
  /** The deck after a change of its slides, as one step of the window's
   * own commit; slides not in ``changedSlideIds`` keep the window's own. */
  commit(next: LiveContent, changedSlideIds: readonly string[]): boolean;
  goToSlide(index: number): boolean;
}

export interface PdfSearchHit {
  /** 1-based page number. */
  readonly page: number;
  /** The words around the match, as the page prints them. */
  readonly snippet: string;
  /** 0-based index among all matches of the search. */
  readonly index: number;
}

export interface LivePdfEditor {
  readonly kind: "pdf";
  pageCount(): number;
  currentPage(): number;
  goToPage(page: number): boolean;
  search(text: string, options: { limit?: number }): Promise<PdfSearchHit[]>;
  /** Highlight (or underline, strike out) the `occurrence`-th match of
   * `text` (on `page` when given) as the ribbon's Highlight does: the
   * page's own markup, saved with it. */
  markText(
    text: string,
    options: { page?: number; occurrence?: number; style?: "highlight" | "underline" | "strikethrough" | "squiggly" },
  ): Promise<{ page: number } | { error: "not_found" | "not_measurable" | "read_only" }>;
}

export type LiveEditorHandle = LiveDocumentEditor | LiveSpreadsheetEditor | LivePresentationEditor | LivePdfEditor;

const handles = new Map<string, LiveEditorHandle[]>();
const listeners = new Set<() => void>();

/** Register ``handle`` for ``pageId`` while the editor is mounted; returns
 * the unregister. The newest registration of a page answers first. */
export function registerLiveEditor(pageId: string, handle: LiveEditorHandle): () => void {
  if (!pageId) return () => undefined;
  const list = handles.get(pageId) ?? [];
  list.unshift(handle);
  handles.set(pageId, list);
  for (const listener of [...listeners]) listener();
  return () => {
    const current = handles.get(pageId);
    if (!current) return;
    const at = current.indexOf(handle);
    if (at >= 0) current.splice(at, 1);
    if (current.length === 0) handles.delete(pageId);
  };
}

/** The live editor of ``pageId`` (of ``kind`` when given), or null. */
export function liveEditorFor<K extends LiveEditorHandle["kind"]>(
  pageId: string,
  kind?: K,
): Extract<LiveEditorHandle, { kind: K }> | null {
  const list = handles.get(pageId) ?? [];
  const found = kind ? list.find((handle) => handle.kind === kind) : list[0];
  return (found as Extract<LiveEditorHandle, { kind: K }> | undefined) ?? null;
}

/** The live editor once it registers (within ``timeoutMs``), or null. */
export function waitForLiveEditor<K extends LiveEditorHandle["kind"]>(
  pageId: string,
  kind: K,
  timeoutMs: number,
): Promise<Extract<LiveEditorHandle, { kind: K }> | null> {
  const now = liveEditorFor(pageId, kind);
  if (now || timeoutMs <= 0) return Promise.resolve(now);
  return new Promise((resolve) => {
    const done = (value: Extract<LiveEditorHandle, { kind: K }> | null) => {
      listeners.delete(listener);
      clearTimeout(timer);
      resolve(value);
    };
    const listener = () => {
      const found = liveEditorFor(pageId, kind);
      if (found) done(found);
    };
    const timer = setTimeout(() => done(liveEditorFor(pageId, kind)), timeoutMs);
    listeners.add(listener);
  });
}
