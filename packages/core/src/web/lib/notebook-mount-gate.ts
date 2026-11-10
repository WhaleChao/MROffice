/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The way a notebook-scoped API request waits to learn which workspace its
 * notebook lives in (see settleNotebookMount in lib/api.ts).
 *
 * The shared-notebook module provides it and the API client uses it. It sits
 * here, with no imports, so neither has to import the other and a test that
 * replaces either module still loads the other.
 */
export type NotebookMountGate = (notebookId: string) => Promise<unknown>;

let gate: NotebookMountGate | null = null;

export function registerNotebookMountGate(next: NotebookMountGate): void {
  gate = next;
}

export function notebookMountGate(): NotebookMountGate | null {
  return gate;
}
