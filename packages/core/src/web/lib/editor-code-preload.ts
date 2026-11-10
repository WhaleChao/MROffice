/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editor an imported file opens in, loaded while the import is still on
 * its way. A first opening after an import waited for the editor's code
 * (the workbook editor alone is 4 MB) after the import's seconds, when the
 * browser had been idle through them; started with the upload, it is in
 * place by the time the page opens. Best effort, once per editor: a failed
 * load leaves the window to load it as it always did.
 */

export type PreloadedEditorKind = "pdf" | "docx" | "xlsx" | "pptx";

const started = new Set<PreloadedEditorKind>();

export function preloadPageEditorCode(kind: PreloadedEditorKind): void {
  // Not under a test runner: a unit test of an import would compile the editor.
  if (typeof window === "undefined" || process.env.NODE_ENV === "test" || started.has(kind)) return;
  started.add(kind);
  const load = kind === "xlsx"
    ? () => import("@mrnote/sheets/web/components/notebook/contents/SpreadsheetWindow.tsx")
    : kind === "pptx"
      ? () => import("@mrnote/slides/web/components/notebook/contents/PresentationWindow.tsx")
      : kind === "pdf"
        ? () => import("@mrnote/pdf/web/components/notebook/contents/PdfPageWindow.tsx")
        : () => import("../components/notebook/contents/NoteWindow");
  void load().catch(() => {
    started.delete(kind);
  });
}
