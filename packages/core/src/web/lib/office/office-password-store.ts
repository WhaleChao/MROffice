/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The open documents' passwords: this tab's memory only, never a URL, a
 * header, a log or any storage, and forgotten when the person signs out.
 * Zero-dependency on purpose: the sign-out in `lib/api` clears it, and
 * `lib/office/office-encryption` depends on `lib/api`.
 */

const pagePasswords = new Map<string, string>();

export function heldPagePassword(pageId: string): string | null {
  return pagePasswords.get(pageId) ?? null;
}

export function holdPagePassword(pageId: string, password: string | null): void {
  if (password) pagePasswords.set(pageId, password);
  else pagePasswords.delete(pageId);
}

/** Forget every password this tab holds (sign-out). */
export function forgetPagePasswords(): void {
  pagePasswords.clear();
}
