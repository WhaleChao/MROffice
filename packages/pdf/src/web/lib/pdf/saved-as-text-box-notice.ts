/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Takes the notice for the editor at `from` (an element of its window):
 * true when this window says it (or keeps it unsaid) itself. */
export type PdfSavedAsTextBoxNoticeWindow = (from: Element | null, say: () => void) => boolean;

const windows = new Set<PdfSavedAsTextBoxNoticeWindow>();

/** A window answers for the notices of its editors until the returned
 * function is called. */
export function answerForPdfSavedAsTextBoxNotice(window: PdfSavedAsTextBoxNoticeWindow): () => void {
  windows.add(window);
  return () => { windows.delete(window); };
}

/** Says the notice (`say`) through the window of the editor at `from`. */
export function sayPdfSavedAsTextBox(from: Element | null, say: () => void): void {
  for (const window of windows) {
    try {
      if (window(from, say)) return;
    } catch { /* another window's, or none's: said below */ }
  }
  say();
}

/**
 * How a window answers. `busy` says whether a write of the window's is
 * still on its way (the background save behind an editor, a box's own
 * write, the working copy's queue); `refusedSince` whether a text box write
 * was refused since that moment. The notice is said once the window has
 * been quiet for two looks in a row -- a write that follows the editor's
 * letting go starts a moment after it -- and not at all when a write was
 * refused meanwhile: a line of Urdu became a box whose own write the server
 * refused four seconds after 「已按文本框保存」 had been said (the box's
 * write is not the editor's background save: waited for alone, that one
 * was over before the box's began). After a minute and a half of writes it
 * is not said either: nothing is known to be saved.
 */
export function pdfSavedAsTextBoxNoticeWhenWritten(input: Readonly<{
  busy: () => boolean;
  refusedSince: (at: number) => boolean;
  alive: () => boolean;
  now?: () => number;
  wait?: (ms: number) => Promise<void>;
}>): (say: () => void) => Promise<void> {
  const now = input.now ?? (() => Date.now());
  const wait = input.wait ?? ((ms: number) => new Promise<void>((resolve) => { setTimeout(resolve, ms); }));
  return async (say) => {
    const asked = now();
    await wait(600);
    let quiet = 0;
    while (quiet < 2) {
      if (!input.alive() || now() - asked > 90_000) return;
      quiet = input.busy() ? 0 : quiet + 1;
      if (quiet < 2) await wait(300);
    }
    if (input.alive() && !input.refusedSince(asked - 5_000)) say();
  };
}
