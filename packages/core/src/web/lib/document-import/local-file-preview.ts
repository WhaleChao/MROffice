/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A file shown from the reader's own copy while the server makes its page.
 *
 * An Office or PDF file becomes a page on the server: its upload takes
 * seconds, the page minutes for a large workbook. The browser holds the file
 * all along, so the window it will open in opens at once and shows what the
 * file itself can show without being read as a document: a PDF's pages
 * (pdf.js, as the PDF editor draws them) or the picture of its first page an
 * Office file carries (`docProps/thumbnail.*`). When the page exists the
 * window becomes its editor in place, and the editor's loading card keeps
 * showing the same preview (components/ui/LocalFilePreviewLayer) until the
 * editor has drawn the page.
 *
 * Nothing here reads an Office document's content: the server reads the
 * package (feedback: Office import is server side). Previews live in memory
 * only, for this tab, and are dropped a while after nothing shows them.
 */

import { editableDocumentExtension } from "./editableDocumentFile";
import { randomUUID } from "../random-uuid";
import {
  isLegacyDocumentImportFile,
  isLegacyPresentationImportFile,
  isLegacySpreadsheetImportFile,
} from "../page-import-capabilities";

export type LocalFilePreviewKind = "pdf" | "document" | "spreadsheet" | "presentation";

export interface LocalFilePreviewEntry {
  readonly id: string;
  readonly file: File;
  readonly kind: LocalFilePreviewKind;
  readonly title: string;
  readonly notebookId: string;
  /** The step the import is at, as its sidebar placeholder says it. */
  readonly status: string | null;
  /** The page the file became, once it exists. */
  readonly pageId: string | null;
  readonly pageType: string | null;
  /** Why no page came of it (the import failed). */
  readonly failure: string | null;
  /** The server import that is making the page, once it has said it read the
   * deck's first slides beside it (lib/document-import/pptx-first-slides). */
  readonly firstSlidesJob?: { readonly notebookId: string; readonly jobId: string } | null;
}

type Listener = () => void;

const entries = new Map<string, LocalFilePreviewEntry>();
const byPage = new Map<string, string>();
const listeners = new Set<Listener>();
const holds = new Map<string, number>();
const dropTimers = new Map<string, number>();
/** Per-entry resources a view made (a PDF document), freed with the entry. */
const resources = new Map<string, Map<string, { value: unknown; free: () => void }>>();

/** A preview nobody shows any more is dropped after this long. */
const DROP_AFTER_MS = 30_000;

function emit(): void {
  for (const listener of Array.from(listeners)) listener();
}

export function subscribeLocalFilePreviews(listener: Listener): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

export function readLocalFilePreview(id: string | null | undefined): LocalFilePreviewEntry | null {
  return id ? entries.get(id) ?? null : null;
}

/** The preview of the file a page was made from, while it is kept. */
export function readLocalFilePreviewForPage(
  pageId: string | null | undefined,
): LocalFilePreviewEntry | null {
  if (!pageId) return null;
  return readLocalFilePreview(byPage.get(pageId));
}

/**
 * The kind of page an Office/PDF file becomes, from its name and type alone
 * (the importer's own rule, lib/page-file-import `predictedImportedPage`),
 * so its window can open before the importer's code has loaded. Null for
 * any other file.
 */
export function localFilePreviewKindOfFile(file: File): LocalFilePreviewKind | null {
  const extension = editableDocumentExtension(file);
  if (extension === ".pdf") return "pdf";
  if (extension === ".pptx" || isLegacyPresentationImportFile(file)) return "presentation";
  if (extension === ".xlsx" || isLegacySpreadsheetImportFile(file)) return "spreadsheet";
  if (extension === ".docx" || isLegacyDocumentImportFile(file)) return "document";
  return null;
}

export function createLocalFilePreview(input: {
  readonly file: File;
  readonly kind: LocalFilePreviewKind;
  readonly title: string;
  readonly notebookId: string;
  readonly status?: string | null;
}): LocalFilePreviewEntry {
  const entry: LocalFilePreviewEntry = {
    id: randomUUID(),
    file: input.file,
    kind: input.kind,
    title: input.title,
    notebookId: input.notebookId,
    status: input.status ?? null,
    pageId: null,
    pageType: null,
    failure: null,
  };
  entries.set(entry.id, entry);
  // Shown by the window that opens next; dropped if none ever does.
  scheduleDrop(entry.id);
  emit();
  return entry;
}

export function updateLocalFilePreview(
  id: string,
  patch: Partial<Pick<LocalFilePreviewEntry, "status" | "pageId" | "pageType" | "failure" | "firstSlidesJob">>,
): void {
  const current = entries.get(id);
  if (!current) return;
  const next = { ...current, ...patch };
  if (
    next.status === current.status &&
    next.pageId === current.pageId &&
    next.pageType === current.pageType &&
    next.failure === current.failure &&
    next.firstSlidesJob?.jobId === current.firstSlidesJob?.jobId
  ) {
    return;
  }
  entries.set(id, next);
  if (next.pageId) byPage.set(next.pageId, id);
  emit();
}

function scheduleDrop(id: string): void {
  window.clearTimeout(dropTimers.get(id));
  dropTimers.set(
    id,
    window.setTimeout(() => {
      dropTimers.delete(id);
      if ((holds.get(id) ?? 0) > 0) return;
      dropLocalFilePreview(id);
    }, DROP_AFTER_MS),
  );
}

export function dropLocalFilePreview(id: string): void {
  const entry = entries.get(id);
  if (!entry) return;
  entries.delete(id);
  if (entry.pageId && byPage.get(entry.pageId) === id) byPage.delete(entry.pageId);
  window.clearTimeout(dropTimers.get(id));
  dropTimers.delete(id);
  holds.delete(id);
  const made = resources.get(id);
  resources.delete(id);
  if (made) for (const { free } of made.values()) free();
  emit();
}

/** Keep `id` while a view shows it; returns the release. */
export function retainLocalFilePreview(id: string): () => void {
  holds.set(id, (holds.get(id) ?? 0) + 1);
  window.clearTimeout(dropTimers.get(id));
  dropTimers.delete(id);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    const left = (holds.get(id) ?? 1) - 1;
    if (left > 0) {
      holds.set(id, left);
      return;
    }
    holds.delete(id);
    if (entries.has(id)) scheduleDrop(id);
  };
}

/**
 * The resource `key` of preview `id`, made once and shared by every view of
 * it (a window's preview and the editor's loading card after it), so the
 * hand-over from one to the next does not read the file again.
 */
export function localFilePreviewResource<T>(
  id: string,
  key: string,
  make: () => { value: T; free: () => void },
): T | null {
  if (!entries.has(id)) return null;
  let made = resources.get(id);
  if (!made) {
    made = new Map();
    resources.set(id, made);
  }
  const existing = made.get(key);
  if (existing) return existing.value as T;
  const created = make();
  made.set(key, created);
  return created.value;
}
