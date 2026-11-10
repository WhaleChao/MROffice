/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Save a copy to my notebook": a workbook's edits the window can no longer
 * save where they were made (Excel's Save a Copy) become a new page the
 * person can write to. The page's own notebook first (only the page may have
 * been locked); otherwise one of the person's notebooks; otherwise a new
 * notebook of theirs for recovered copies. The copy is an ordinary new page:
 * nothing of the original is changed.
 */
import { isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import { createNotebook, createPage, listNotebooks, type NotebookInfo, type PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";

function refusedHere(error: unknown): boolean {
  return isApiRequestError(error) && (error.status === 403 || error.status === 404 || error.status === 409);
}

export async function saveWorkbookCopyToNotebook(input: {
  sourceNotebookId: string | null | undefined;
  title: string;
  content: Record<string, unknown>;
  recoveredNotebookTitle: string;
}): Promise<{ page: PageItem; notebookTitle: string | null }> {
  const body = { title: input.title, page_type: "spreadsheet" as const, content_json: input.content };
  if (input.sourceNotebookId) {
    try {
      return { page: await createPage(input.sourceNotebookId, body), notebookTitle: null };
    } catch (error) {
      if (!refusedHere(error)) throw error;
    }
  }
  let notebooks: NotebookInfo[] = [];
  try {
    notebooks = (await listNotebooks()).items ?? [];
  } catch {
    notebooks = [];
  }
  const candidates = notebooks
    .filter((notebook) => notebook.id !== input.sourceNotebookId)
    .sort((left, right) => String(right.updated_at ?? "").localeCompare(String(left.updated_at ?? "")));
  const recovered = candidates.find((notebook) => notebook.title === input.recoveredNotebookTitle);
  for (const notebook of recovered ? [recovered] : candidates.slice(0, 1)) {
    try {
      return { page: await createPage(notebook.id, body), notebookTitle: notebook.title };
    } catch (error) {
      if (!refusedHere(error)) throw error;
    }
  }
  const notebook = await createNotebook({ title: input.recoveredNotebookTitle, notebook_type: "work" });
  return { page: await createPage(notebook.id, body), notebookTitle: notebook.title };
}
