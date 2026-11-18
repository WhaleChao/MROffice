/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { FontReplacement } from "@mrnote/core/web/lib/presentation-replace-font.ts";

/** How long the log must stay as it is before its warm-up is asked for. */
export const REPLACE_FONTS_PREWARM_DEBOUNCE_MS = 2_000;

export interface ReplaceFontsPrewarm {
  /** The deck's log changed to `log`: (re)start the wait for it. */
  schedule(log: readonly FontReplacement[]): void;
  /** Drop a pending warm-up (the window closes, the page changes). */
  cancel(): void;
}

export function createReplaceFontsPrewarm(
  send: (log: readonly FontReplacement[]) => void,
  timers: { set: (callback: () => void, ms: number) => unknown; clear: (handle: unknown) => void } = {
    set: (callback, ms) => setTimeout(callback, ms),
    clear: (handle) => clearTimeout(handle as ReturnType<typeof setTimeout>),
  },
): ReplaceFontsPrewarm {
  let pending: unknown = null;
  let lastSent: string | null = null;
  return {
    schedule(log) {
      if (pending !== null) timers.clear(pending);
      const key = JSON.stringify(log.map(({ id, from, to }) => [id, from, to]));
      pending = timers.set(() => {
        pending = null;
        if (key === lastSent) return;
        lastSent = key;
        send(log);
      }, REPLACE_FONTS_PREWARM_DEBOUNCE_MS);
    },
    cancel() {
      if (pending !== null) timers.clear(pending);
      pending = null;
    },
  };
}
