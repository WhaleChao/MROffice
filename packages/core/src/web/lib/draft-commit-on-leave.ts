/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";
import { flushSync } from "react-dom";

import { registerUnsavedWorkFlush } from "./unsaved-work-hold";

function now(write: () => void): void {
  try {
    flushSync(write);
  } catch {
    try {
      write();
    } catch {
      // a box that cannot write its words leaves the others to theirs
    }
  }
}

/** The focused element blurred, its blur's writing done. */
export function commitFocusedDraft(): void {
  if (typeof document === "undefined") return;
  const active = document.activeElement;
  if (!(active instanceof HTMLElement) || active === document.body) return;
  now(() => active.blur());
}

export function installFocusedDraftCommit(): () => void {
  if (typeof window === "undefined") return () => undefined;
  const unregister = registerUnsavedWorkFlush("focused-draft", commitFocusedDraft);
  // Capture: before any page's own beforeunload/pagehide reads what is unsaved.
  window.addEventListener("beforeunload", commitFocusedDraft, true);
  window.addEventListener("pagehide", commitFocusedDraft, true);
  return () => {
    unregister();
    window.removeEventListener("beforeunload", commitFocusedDraft, true);
    window.removeEventListener("pagehide", commitFocusedDraft, true);
  };
}

/** `commit` writes a draft its blur does not; called before the page
 * unloads or hides and when the app settles its work. */
export function useDraftCommitOnLeave(id: string, commit: () => void, enabled = true): void {
  const latest = useRef(commit);
  latest.current = commit;
  useEffect(() => {
    if (!enabled || typeof window === "undefined") return undefined;
    const write = () => now(() => latest.current());
    const unregister = registerUnsavedWorkFlush(id, write);
    window.addEventListener("beforeunload", write, true);
    window.addEventListener("pagehide", write, true);
    return () => {
      unregister();
      window.removeEventListener("beforeunload", write, true);
      window.removeEventListener("pagehide", write, true);
    };
  }, [enabled, id]);
}
