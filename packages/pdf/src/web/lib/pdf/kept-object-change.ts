/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfKeptObjectKind = "shape" | "ink" | "signature" | "stamp" | "link";

type Kept = { readonly id: string };

const files = new Map<string, Map<string, Kept>>();

/** The reader-held-work key of one object's kept change. */
export function pdfKeptObjectKey(kind: PdfKeptObjectKind, id: string): string {
  return `${kind}:${id}`;
}

/** Keep `wanted` as this object's change; a later one for it replaces it. */
export function keepPdfObjectChange<T extends Kept>(
  file: string,
  kind: PdfKeptObjectKind,
  wanted: T,
): void {
  if (!file) return;
  const kept = files.get(file) ?? new Map<string, Kept>();
  files.set(file, kept);
  kept.set(pdfKeptObjectKey(kind, wanted.id), wanted);
}

/** The object as the user left it, while the file does not hold it. */
export function pdfKeptObjectChange<T extends Kept>(
  file: string,
  kind: PdfKeptObjectKind,
  id: string,
): T | null {
  return (files.get(file)?.get(pdfKeptObjectKey(kind, id)) as T | undefined) ?? null;
}

/** Every change kept for this file, for a reader that opens over them. */
export function pdfKeptObjectChanges(
  file: string,
): { kind: PdfKeptObjectKind; id: string }[] {
  return [...(files.get(file)?.keys() ?? [])].map((key) => {
    const at = key.indexOf(":");
    return { kind: key.slice(0, at) as PdfKeptObjectKind, id: key.slice(at + 1) };
  });
}

/** The change landed, or its object is gone: nothing is kept for it. */
export function settlePdfObjectChange(file: string, kind: PdfKeptObjectKind, id: string): void {
  const kept = files.get(file);
  if (!kept?.delete(pdfKeptObjectKey(kind, id))) return;
  if (kept.size === 0) files.delete(file);
}

/**
 * Rows as read, with each object that has a kept change drawn as kept. An
 * object the rows no longer hold is not brought back: it was deleted.
 */
export function withKeptPdfObjectChanges<T extends Kept>(
  file: string | null | undefined,
  kind: PdfKeptObjectKind,
  items: T[],
): T[] {
  const kept = file ? files.get(file) : undefined;
  if (!kept?.size) return items;
  let changed = false;
  const next = items.map((item) => {
    const wanted = kept.get(pdfKeptObjectKey(kind, item.id)) as T | undefined;
    if (!wanted) return item;
    changed = true;
    return wanted;
  });
  return changed ? next : items;
}

/** Tests only. */
export function resetPdfKeptObjectChanges(): void {
  files.clear();
}
