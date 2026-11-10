/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Return true while a spreadsheet keydown belongs to an active IME session.
 *
 * Chromium exposes the standards-based `isComposing` flag. WebKit can report
 * the candidate-confirming Enter after compositionend with `isComposing=false`
 * while still tagging the native key as PROCESS (`keyCode === 229`). Treat
 * either signal as IME-owned so grid navigation and formula autocomplete do not
 * consume the candidate-confirmation key.
 */
export function isSpreadsheetImeKeydown(event: {
  isComposing?: boolean;
  keyCode?: number;
}): boolean {
  return event.isComposing === true || event.keyCode === 229;
}
