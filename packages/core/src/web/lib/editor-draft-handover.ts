/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Editors whose newest typing lives only on screen (a PDF line still being
 * typed) hand it to their save queue before the page reloads. The "reload to
 * update" notices promise that unsaved changes are kept; a line being typed
 * was not, and the window came back on the old text (user audit PDF 0051).
 *
 * Each open editor registers a handover; a reload the app asks for runs them
 * all first, never waiting longer than the bound.
 */

import { noteAppReload } from "./reload-reason";

type Handover = () => Promise<void>;

const handovers = new Set<Handover>();

/** Longest a reload waits for the editors to hand their drafts over. */
export const EDITOR_DRAFT_HANDOVER_TIMEOUT_MS = 10_000;

/** Register an editor's handover; the returned function removes it. */
export function registerEditorDraftHandover(handover: Handover): () => void {
  handovers.add(handover);
  return () => {
    handovers.delete(handover);
  };
}

/** Run every registered handover, bounded; failures never stop the others. */
export async function handOverEditorDrafts(
  timeoutMs: number = EDITOR_DRAFT_HANDOVER_TIMEOUT_MS,
): Promise<void> {
  if (handovers.size === 0) return;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const bound = new Promise<void>((resolve) => {
    timer = setTimeout(resolve, timeoutMs);
  });
  const all = Promise.allSettled(
    [...handovers].map((handover) => Promise.resolve().then(handover)),
  ).then(() => undefined);
  try {
    await Promise.race([all, bound]);
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Reload the page once the open editors have handed their drafts over. The
 * reason is kept for the next load's session_start (lib/reload-reason); used
 * as a click handler it is the notice's button.
 */
export function reloadAfterEditorHandover(reason?: unknown): void {
  noteAppReload(typeof reason === "string" ? reason : "notice_button");
  void handOverEditorDrafts().finally(() => window.location.reload());
}
