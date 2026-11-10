/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const loaded = new WeakMap<object, unknown>();

export function rememberLoadedSpelling<T>(normalized: T, stored: unknown): T {
  if (normalized && typeof normalized === "object" && stored !== undefined) loaded.set(normalized as object, stored);
  return normalized;
}

/** The value as read when it is the very object that reading made; else itself. */
export function asLoadedSpelling<T>(value: T): T {
  if (!value || typeof value !== "object") return value;
  return loaded.has(value as object) ? (loaded.get(value as object) as T) : value;
}
