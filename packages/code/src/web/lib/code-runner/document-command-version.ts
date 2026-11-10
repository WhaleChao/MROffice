/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";

export type DocumentCommandVersionFailure = "save_failed" | "stale";

export type DocumentCommandVersionResult =
  | {
      ok: true;
      sourceContentVersion: number;
    }
  | {
      ok: false;
      reason: DocumentCommandVersionFailure;
    };

interface VerifyDocumentCommandVersionOptions {
  pageId: string;
  /** Runtime input can be forged even though the event type says number. */
  baseContentVersion: unknown;
  saveCurrent(): Promise<PageItem | null>;
  /**
   * The editor generation captured before saveCurrent started must still own
   * the buffer after that await. This closes the local/peer-edit window which
   * a server content version alone cannot observe.
   */
  isCurrentGeneration(): boolean;
}

function validContentVersion(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 1;
}

/**
 * Flush the current draft, then prove that a versioned AI command still owns
 * the exact editor generation it was derived from. No editor mutation may run
 * before this resolves successfully.
 */
export async function verifyDocumentCommandVersion({
  pageId,
  baseContentVersion,
  saveCurrent,
  isCurrentGeneration,
}: VerifyDocumentCommandVersionOptions): Promise<DocumentCommandVersionResult> {
  if (!validContentVersion(baseContentVersion)) {
    return { ok: false, reason: "stale" };
  }

  let sourcePage: PageItem | null;
  try {
    sourcePage = await saveCurrent();
  } catch {
    return { ok: false, reason: "save_failed" };
  }
  if (
    sourcePage?.id !== pageId ||
    !validContentVersion(sourcePage.content_version)
  ) {
    return { ok: false, reason: "save_failed" };
  }
  if (
    sourcePage.content_version !== baseContentVersion ||
    !isCurrentGeneration()
  ) {
    return { ok: false, reason: "stale" };
  }
  return {
    ok: true,
    sourceContentVersion: sourcePage.content_version,
  };
}
