/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A file whose bytes are read on request (a large document the desktop app
 * opened from the system: it hands the page a token, not the bytes).
 *
 * It is a `File` of the right name, type and size so every import decision
 * reads it as the file it is; its own bytes are empty, so it can only travel
 * the road that reads it part by part (the source import), which asks
 * {@link lazyFileReader} for each part.
 */

const READER = Symbol.for("mrai.lazyFileReader");

export type LazyFileReader = (start: number, end: number) => Promise<Blob>;

export function lazyFile(
  meta: { readonly name: string; readonly type: string; readonly size: number; readonly lastModified: number },
  read: LazyFileReader,
): File {
  const file = new File([], meta.name, { type: meta.type, lastModified: meta.lastModified });
  Object.defineProperty(file, "size", { value: meta.size });
  Object.defineProperty(file, READER, { value: read });
  return file;
}

export function lazyFileReader(file: File): LazyFileReader | null {
  const reader = (file as unknown as Record<symbol, unknown>)[READER];
  return typeof reader === "function" ? (reader as LazyFileReader) : null;
}
