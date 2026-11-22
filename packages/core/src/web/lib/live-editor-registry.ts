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
  /** Resolves once the calculation of the workbook as it is now has answered
   * and the values store reads it (the window's backgroundCalculation
   * settledView barrier); false when the source moved on or it failed. */
  calculated?(): Promise<boolean>;
  /** Values typed into cells of one sheet, as the formula bar enters them
   * (``asTyped``: "50%", dates and numbers read as Excel reads them typed;
   * "=…" is a formula), committed as one undoable step of the window's own
   * commit (history, collaboration room, autosave). */
  writeCells(sheetId: string, writes: readonly { row: number; col: number; value: string }[], asTyped: boolean): LiveWriteOutcome;
  /** A new sheet after the last (the window's Insert Sheet), named ``name``
   * when given; null when the workbook refused it. */
  addSheet(name?: string): { id: string; name: string } | null;
  /** The person's selection on the open sheet (0-based): its active cell and
   * the selected rectangle. */
  selection?(): { sheetId: string; row: number; col: number; r1: number; c1: number; r2: number; c2: number } | null;
  /** The workbook as the spreadsheet commands read it
   * (lib/editor-commands/spreadsheet): every sheet's whole model and the
   * workbook's settings; null while it is still opening. */
  commandWorkbook?(): LiveCommandWorkbook | null;
  /** The workbook's sheets after one command (the same sheets, the changed
   * ones replaced), committed as one undoable step of the window's own commit
   * (commitSheets: history, collaboration room, autosave). A whole-line edit
   * also travels to collaborators as the ribbon's does. */
  commitWorkbook?(sheets: readonly LiveContent[], lineEdit?: LiveLineEdit): LiveWriteOutcome;
  /** The sheet tab commands (Rename, Delete, Hide / Unhide), as the tab menu
   * runs them: Protect Workbook and the window's own notices included. */
  renameSheet?(sheetId: string, name: string): LiveWriteOutcome;
  deleteSheet?(sheetId: string): LiveWriteOutcome;
  setSheetHidden?(sheetId: string, hidden: boolean): LiveWriteOutcome;
  /** Open a sheet's tab, as clicking it does; false for a hidden sheet. */
  activateSheet?(sheetId: string): boolean;
  /** Select a range (0-based, its sheet opened), as Go To selects one. */
  select?(sheetId: string, rect: { r1: number; c1: number; r2: number; c2: number }): boolean;
  /** Insert > Pictures > This Device with this picture, over the selected
   * cell of the open sheet as that command places it; false when refused. */
  insertPicture?(file: File): boolean;
}

/** Whole rows or columns inserted or deleted (lib/collab/workbook-line-edits). */
export interface LiveLineEdit {
  readonly sheetId: string;
  readonly axis: "row" | "col";
  readonly op: "insert" | "delete";
  readonly at: number;
  readonly count: number;
}

export interface LiveCommandWorkbook {
  /** Each sheet's model as the window holds it (the ribbon's sheet transactions read these). */
  readonly sheets: readonly LiveContent[];
  readonly activeSheetId: string | null;
  /** The workbook's view options (theme, custom lists ...), as the window holds them. */
  readonly viewOptions: LiveContent;
  readonly calculationMode: "automatic" | "manual";
  readonly dateSystem?: string;
  /** Review > Protect Workbook > Structure. */
  readonly workbookStructureProtection?: LiveContent;
  /** The page saves back into the xlsx it was imported from (patchback). */
  readonly imported: boolean;
  /** Sheets the window holds only in part: a large sheet read in tiles, a
   * sheet not filled in yet. A command never runs on one. */
  readonly partialSheetIds: readonly string[];
  /** Sheets read in tiles, whose rows and columns only the server moves. */
  readonly tiledSheetIds: readonly string[];
  /** Who a comment is written by: the person in this window. */
  readonly commentAuthor: { readonly name: string; readonly id: string | null };
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
  /** The shapes (block ids) selected on the slide shown. */
  selectedShapeIds?(): readonly string[];
  /** A number that changes whenever the deck's content does (anyone's edit, Undo). */
  generation?(): number;
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
  /** The pages in reading order: 1-based number, size in points, rotation as shown. */
  pages?(): readonly { page: number; width: number; height: number; rotation: number }[];
  /** A page's text, as its text layer (or its OCR) reads; null for no such page. */
  pageText?(page: number): Promise<string | null>;
  /** The annotations the window holds: markups, notes, ink, shapes, stamps, signatures, redactions. */
  annotations?(): readonly LivePdfAnnotation[];
  /** The form's fields, by full name, with the values shown now. */
  formFields?(): Promise<readonly LivePdfFormField[]>;
  /** Insert > Image with this picture, centred on the page shown at the
   * command's default size; false when refused (read-only, not a picture). */
  insertPicture?(file: File): Promise<boolean>;
}

export interface LivePdfAnnotation {
  readonly id: string;
  readonly type: string;
  readonly page: number;
  readonly text: string;
  readonly color: string | null;
  readonly author: string | null;
  readonly createdAt: string | null;
}

export interface LivePdfFormField {
  readonly name: string;
  readonly type: string;
  readonly value: string | boolean | readonly string[] | null;
  readonly page: number | null;
  readonly readOnly: boolean;
  readonly required: boolean;
  readonly options?: readonly string[];
}

export type LiveEditorHandle = LiveDocumentEditor | LiveSpreadsheetEditor | LivePresentationEditor | LivePdfEditor;

const handles = new Map<string, LiveEditorHandle[]>();
const listeners = new Set<() => void>();
/** Hear every registration and unregistration (lib/desktop-ai-screen). */
const changeListeners = new Set<() => void>();

function changed(): void {
  for (const listener of [...listeners]) listener();
  for (const listener of [...changeListeners]) listener();
}

/** The pages with a live editor now, each with the kind of its newest. */
export function liveEditorPages(): { pageId: string; kind: LiveEditorHandle["kind"] }[] {
  return [...handles.entries()].filter(([, list]) => list.length > 0).map(([pageId, list]) => ({ pageId, kind: list[0]!.kind }));
}

/** ``listener`` runs whenever an editor registers or goes; returns the unsubscribe. */
export function onLiveEditorsChange(listener: () => void): () => void {
  changeListeners.add(listener);
  return () => changeListeners.delete(listener);
}

/** Register ``handle`` for ``pageId`` while the editor is mounted; returns
 * the unregister. The newest registration of a page answers first. */
export function registerLiveEditor(pageId: string, handle: LiveEditorHandle): () => void {
  if (!pageId) return () => undefined;
  const list = handles.get(pageId) ?? [];
  list.unshift(handle);
  handles.set(pageId, list);
  changed();
  return () => {
    const current = handles.get(pageId);
    if (!current) return;
    const at = current.indexOf(handle);
    if (at < 0) return;
    current.splice(at, 1);
    if (current.length === 0) handles.delete(pageId);
    for (const listener of [...changeListeners]) listener();
  };
}

/** The page ``handle`` is registered for, or null once it is gone. */
export function liveEditorPageId(handle: LiveEditorHandle): string | null {
  for (const [pageId, list] of handles) if (list.includes(handle)) return pageId;
  return null;
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
