/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The edits /open/start leaves for a page's own editor to make (each editor
 * saves through its own model, which only the open editor holds): what the
 * editor reads when the page opens, and where it says how that went. A
 * refusal keeps the item (the notebook's notice shows the edits); success
 * removes it.
 */

import {
  readClassicStart,
  updateClassicItem,
  type ClassicGuestEdits,
  type ClassicItem,
} from "./stash";

export const CLASSIC_START_CHANGED_EVENT = "mrnote:classic-start-changed";

const APPLY_KEY = (pageId: string) => `mrnote.classic.apply:${pageId}`;

export function markClassicEditsPending(pageId: string): void {
  try {
    window.sessionStorage.setItem(APPLY_KEY(pageId), "1");
  } catch {
    // the edits then wait in the notice instead
  }
}

/** Read once by the page's editor window when it opens. */
export function takeClassicEditsPending(pageId: string): boolean {
  try {
    if (window.sessionStorage.getItem(APPLY_KEY(pageId)) !== "1") return false;
    window.sessionStorage.removeItem(APPLY_KEY(pageId));
    return true;
  } catch {
    return false;
  }
}

export interface ClassicPendingEdits {
  readonly load: () => Promise<ClassicGuestEdits | null>;
  readonly settle: (outcome: "applied" | string) => Promise<void>;
}

async function itemFor(pageId: string): Promise<ClassicItem | null> {
  const record = await readClassicStart();
  return record?.items.find((item) => item.placed?.pageId === pageId) ?? null;
}

export function classicPendingEditsFor(pageId: string): ClassicPendingEdits {
  return {
    load: async () => {
      const item = await itemFor(pageId);
      if (!item || item.kept) return null;
      return item.edits ?? null;
    },
    settle: async (outcome) => {
      if (process.env.NODE_ENV !== "production") console.info("[classic-start] settled", outcome);
      const item = await itemFor(pageId);
      if (!item?.placed) return;
      const placed = item.placed;
      await updateClassicItem(item.type, (current) =>
        outcome === "applied"
          ? null
          : { ...current, kept: { notebookId: placed.notebookId, pageId, reason: outcome } },
      );
      window.dispatchEvent(new CustomEvent(CLASSIC_START_CHANGED_EVENT));
    },
  };
}
