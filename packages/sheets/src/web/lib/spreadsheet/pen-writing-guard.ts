/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pencilScribblePossible } from "@mrnote/core/web/lib/ink-scribble-guard.ts";

let override: boolean | null = null;

/**
 * True when Ready-mode cell entries are read-only so a pen stroke selects:
 * where Scribble exists (lib/ink-scribble-guard, the detection the ink tools
 * share).
 */
export function spreadsheetReadyEntriesResistPenWriting(): boolean {
  if (override !== null) return override;
  return pencilScribblePossible();
}

export function setSpreadsheetReadyEntriesResistPenWritingForTests(
  value: boolean | null,
): void {
  override = value;
}

/**
 * A key pressed on a Ready-mode entry that is read-only for the pen. A
 * character (or the first key of an IME composition) takes the read-only
 * state off before the browser's default action, so the character is
 * inserted and the entry opens on it as on any other device. If no entry
 * opened by the next frame (the key was consumed: Space on a checkbox cell,
 * a shortcut), the input is locked again. Returns true when it lifted.
 */
export function liftReadyEntryForKeystroke(
  input: HTMLInputElement | HTMLTextAreaElement,
  event: { key: string; metaKey: boolean; ctrlKey: boolean; altKey: boolean; keyCode?: number },
): boolean {
  if (!input.readOnly || event.metaKey || event.ctrlKey || event.altKey) return false;
  if (!(event.key.length === 1 || event.key === "Process" || event.keyCode === 229)) return false;
  input.readOnly = false;
  const relock = () => {
    if (!input.isConnected || input.getAttribute("data-editing") === "true") return;
    input.readOnly = true;
  };
  if (typeof requestAnimationFrame === "function") requestAnimationFrame(relock);
  else setTimeout(relock, 0);
  return true;
}
