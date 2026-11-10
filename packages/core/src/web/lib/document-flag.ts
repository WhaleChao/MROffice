/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const DOCUMENT_FLAGS = {
  /** A window is maximized (InstallAppPrompt.module.css). */
  windowMaximized: "data-window-maximized-open",
  /** A modal dialog is open: floating cards (a worksheet's comment card)
   * step under it rather than over its buttons (live audit S-084). */
  modalDialogOpen: "data-modal-dialog-open",
  /** The notebook shelf is showing: what covers it answers for it while a
   * modal is open (globals.css). */
  notebookShelf: "data-notebook-shelf",
  /** An ink tool is picked (lib/ink-touch-shield): the page's text cannot be
   * selected or called out by a pen or hand (globals.css). */
  inkArmed: "data-ink-armed",
} as const;

type DocumentFlag = (typeof DOCUMENT_FLAGS)[keyof typeof DOCUMENT_FLAGS];

const holders = new Map<DocumentFlag, number>();

/** Raise `flag` on <html> for one holder; the returned call lowers it. */
export function holdDocumentFlag(flag: DocumentFlag): () => void {
  if (typeof document === "undefined") return () => undefined;
  holders.set(flag, (holders.get(flag) ?? 0) + 1);
  document.documentElement.setAttribute(flag, "");
  let held = true;
  return () => {
    if (!held) return;
    held = false;
    const left = Math.max(0, (holders.get(flag) ?? 1) - 1);
    holders.set(flag, left);
    if (left === 0) document.documentElement.removeAttribute(flag);
  };
}
