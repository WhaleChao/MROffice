/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const MAX_ENTRIES = 2;
const MAX_ENTRY_BYTES = 48 * 1024 * 1024;
const SHA256 = /^[a-f\d]{64}$/;

const entries: Array<{ sha256: string; bytes: Uint8Array }> = [];

export function rememberPdfSourceBytes(sha256: string, bytes: Uint8Array): void {
  const digest = sha256.trim().toLowerCase();
  if (!SHA256.test(digest) || bytes.length === 0 || bytes.length > MAX_ENTRY_BYTES) {
    return;
  }
  const existing = entries.findIndex((entry) => entry.sha256 === digest);
  if (existing >= 0) {
    // Most recently shown last; the bytes are the same by definition.
    entries.push(...entries.splice(existing, 1));
    return;
  }
  entries.push({ sha256: digest, bytes: Uint8Array.from(bytes) });
  while (entries.length > MAX_ENTRIES) entries.shift()?.bytes.fill(0);
}

/** A private copy of the remembered bytes for `sha256`, or null. */
export function copyRememberedPdfSourceBytes(sha256: string): Uint8Array | null {
  const digest = sha256.trim().toLowerCase();
  const entry = entries.find((candidate) => candidate.sha256 === digest);
  return entry ? Uint8Array.from(entry.bytes) : null;
}

export function forgetPdfSourceBytes(): void {
  for (const entry of entries.splice(0)) entry.bytes.fill(0);
}
