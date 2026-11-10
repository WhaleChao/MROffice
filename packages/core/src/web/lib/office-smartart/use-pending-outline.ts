"use client";


















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useRef, useState } from "react";

import type { SmartArtOutlineItem } from "./outline";

export interface SmartArtPendingOutline {
  /** The diagram the commands are for. */
  readonly groupId: string;
  readonly items: readonly SmartArtOutlineItem[];
}

export interface SmartArtPendingOutlineState {
  /** What the commands still on their way ask for, or `null`. */
  readonly pending: SmartArtPendingOutline | null;
  /** The outline a new command for `groupId` starts from. */
  readonly base: (
    groupId: string,
    model: readonly SmartArtOutlineItem[],
  ) => readonly SmartArtOutlineItem[];
  /**
   * Show `items` for `groupId` from now on; the returned callback says this
   * command settled (call it once, whatever came of the command). A row's
   * words committed on blur are drawn `"after-event"`; a structural command
   * (Enter's new row, a reorder) is drawn at once, so the pane can put the
   * caret in the row it selects.
   */
  readonly begin: (
    groupId: string,
    items: readonly SmartArtOutlineItem[],
    draw?: "now" | "after-event",
  ) => () => void;
}

/** The outline the pane shows for `groupId`: the pending one, else the members'. */
export function smartArtShownOutline<T extends readonly SmartArtOutlineItem[] | null>(
  pending: SmartArtPendingOutline | null,
  groupId: string | null | undefined,
  model: T,
): T | readonly SmartArtOutlineItem[] {
  return groupId && pending?.groupId === groupId ? pending.items : model;
}

export function useSmartArtPendingOutline(): SmartArtPendingOutlineState {
  const [pending, setPending] = useState<SmartArtPendingOutline | null>(null);
  const current = useRef<SmartArtPendingOutline | null>(null);
  const open = useRef(0);
  const base = useCallback(
    (groupId: string, model: readonly SmartArtOutlineItem[]) =>
      current.current?.groupId === groupId ? current.current.items : model,
    [],
  );
  // A row's words are committed on blur, while the browser is moving the
  // focus to the row the author pressed, and WebKit drops that focus when the
  // list under it is re-rendered mid-move (the next keys then went to the
  // slide): that change is drawn a moment later, after the event. The next
  // command reads `current` at once either way. The settling draw is late
  // too, so it can never land before a late draw it follows.
  const later = useRef(0);
  const draw = useCallback((when: "now" | "after-event") => {
    if (when === "now" && later.current === 0) {
      setPending(current.current);
      return;
    }
    later.current += 1;
    setTimeout(() => {
      later.current -= 1;
      setPending(current.current);
    }, 0);
  }, []);
  const begin = useCallback(
    (
      groupId: string,
      items: readonly SmartArtOutlineItem[],
      when: "now" | "after-event" = "now",
    ) => {
      open.current += 1;
      current.current = { groupId, items };
      draw(when);
      let settled = false;
      return () => {
        if (settled) return;
        settled = true;
        open.current -= 1;
        if (open.current > 0) return;
        current.current = null;
        draw("after-event");
      };
    },
    [draw],
  );
  return { pending, base, begin };
}
