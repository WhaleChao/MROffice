/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The page edits AutoSave off holds until Save (comments, ink, stamps, form
 * values, page operations: PdfPageWindow heldContentRef), as this tab keeps
 * them through a reload.
 *
 * Kept as what the edits changed, not as the page they were made on: per
 * collection, the rows they changed or added and the ids of the rows they
 * took away (a collection whose entries are not all told apart by id: its
 * value before and after). A reload lays that onto whatever the page holds
 * then -- the same version or one another window saved meanwhile -- so a
 * peer's rows stay and the held ones come back. The record used to be the
 * whole collections on one version of the page: dropped, with the rows the
 * edits had made deleted, when the page had moved (this window's own Save of
 * its text queue was enough), and dropped past 2 MB.
 */
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";

export type PdfHeldKeyDelta =
  | Readonly<{
      kind: "rows";
      /** Rows the held edits changed or added, as they left them. */
      changed: readonly unknown[];
      /** Ids of the rows they took away. */
      removed: readonly string[];
      /** The collection's ids in the order the held edits left them. */
      order: readonly string[];
    }>
  | Readonly<{
      kind: "value";
      /** What the collection held when the edits were made (null: nothing). */
      before: unknown;
      after: unknown;
      /** The edits left the page without the key. */
      absent?: boolean;
    }>;

export interface PdfHeldContentRecord {
  readonly v: 2;
  /** The page version the edits last stood on (diagnostics only). */
  readonly baseVersion: number;
  readonly keys: Readonly<Record<string, PdfHeldKeyDelta>>;
  readonly actions: readonly string[];
  readonly createdRows: readonly string[];
}

function rowId(row: unknown): string | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return null;
  const id = (row as { id?: unknown }).id;
  return typeof id === "string" && id ? id : null;
}

/** A collection whose entries are told apart by id (one the page does not hold: empty). */
function rowsById(value: unknown): Map<string, unknown> | null {
  if (value === undefined || value === null) return new Map();
  if (!Array.isArray(value)) return null;
  const map = new Map<string, unknown>();
  for (const row of value) {
    const id = rowId(row);
    if (!id || map.has(id)) return null;
    map.set(id, row);
  }
  return map;
}

/** What the held edits changed in `keys`, between the page they stand on and the page as shown. */
export function pdfHeldContentDelta(
  base: Readonly<Record<string, unknown>>,
  content: Readonly<Record<string, unknown>>,
  keys: Iterable<string>,
): Record<string, PdfHeldKeyDelta> {
  const out: Record<string, PdfHeldKeyDelta> = {};
  for (const key of keys) {
    const before = rowsById(base[key]);
    const after = rowsById(content[key]);
    if (before && after && (before.size || after.size)) {
      const changed: unknown[] = [];
      for (const [id, row] of after) {
        const saved = before.get(id);
        if (saved === undefined || !pageDocumentJsonEquals(saved as never, row as never)) changed.push(row);
      }
      const removed = [...before.keys()].filter((id) => !after.has(id));
      out[key] = { kind: "rows", changed, removed, order: [...after.keys()] };
      continue;
    }
    out[key] = {
      kind: "value",
      before: base[key] ?? null,
      after: content[key] ?? null,
      ...(Object.prototype.hasOwnProperty.call(content, key) ? {} : { absent: true }),
    };
  }
  return out;
}

/**
 * The held edits laid onto `loaded` (the page as it is now). `moved` names
 * the collections kept as one value that the page changed meanwhile: the
 * held value stands (as it does when the page is read in under held edits
 * without a reload), and Save writes it over what the page holds there.
 */
export function layPdfHeldContent(
  loaded: Readonly<Record<string, unknown>>,
  deltas: Readonly<Record<string, PdfHeldKeyDelta>>,
): { content: Record<string, unknown>; moved: string[] } {
  const content: Record<string, unknown> = { ...loaded };
  const moved: string[] = [];
  for (const [key, delta] of Object.entries(deltas)) {
    const now = loaded[key];
    if (delta.kind === "rows" && (now === undefined || now === null || Array.isArray(now))) {
      const removed = new Set(delta.removed);
      const changed = new Map<string, unknown>();
      for (const row of delta.changed) {
        const id = rowId(row);
        if (id) changed.set(id, row);
      }
      const rows: unknown[] = [];
      const seen = new Set<string>();
      let told = true;
      for (const row of (now ?? []) as unknown[]) {
        const id = rowId(row);
        if (!id) { told = false; rows.push(row); continue; }
        if (removed.has(id) || seen.has(id)) continue;
        seen.add(id);
        rows.push(changed.has(id) ? changed.get(id) : row);
      }
      for (const [id, row] of changed) {
        if (!seen.has(id)) { seen.add(id); rows.push(row); }
      }
      // On the rows the edits were made over: in the order they left them.
      if (told && seen.size === delta.order.length && delta.order.every((id) => seen.has(id))) {
        const place = new Map(delta.order.map((id, index) => [id, index]));
        rows.sort((left, right) => place.get(rowId(left)!)! - place.get(rowId(right)!)!);
      }
      content[key] = rows;
      continue;
    }
    if (delta.kind === "rows") {
      // The page holds something else than rows there now: the held rows
      // cannot be told from it.
      moved.push(key);
      continue;
    }
    if (!pageDocumentJsonEquals((now ?? null) as never, (delta.before ?? null) as never)) moved.push(key);
    if (delta.absent) delete content[key];
    else content[key] = delta.after;
  }
  return { content, moved };
}

function deltaOf(value: unknown): PdfHeldKeyDelta | null {
  if (!value || typeof value !== "object") return null;
  const delta = value as { kind?: unknown; changed?: unknown; removed?: unknown; order?: unknown };
  if (delta.kind === "value") return value as PdfHeldKeyDelta;
  if (delta.kind !== "rows" || !Array.isArray(delta.changed) || !Array.isArray(delta.removed) ||
    !Array.isArray(delta.order) || !delta.changed.every((row) => rowId(row) !== null) ||
    ![...delta.removed, ...delta.order].every((id) => typeof id === "string")) return null;
  return value as PdfHeldKeyDelta;
}

/**
 * The record a tab kept, read back. `null`: nothing was kept. "unreadable":
 * something was, that this build cannot lay back (the reader is told).
 * A record of the earlier form (whole collections on one page version) is
 * read as values: it comes back on that version only, as it did.
 */
export function readPdfHeldContentRecord(
  raw: string | null,
  loaded: Readonly<{ version: number; content: Readonly<Record<string, unknown>> }>,
): PdfHeldContentRecord | "unreadable" | null {
  if (!raw) return null;
  let parsed: unknown;
  try { parsed = JSON.parse(raw); } catch { return "unreadable"; }
  if (!parsed || typeof parsed !== "object") return null;
  const record = parsed as { v?: unknown; baseVersion?: unknown; keys?: unknown; content?: unknown;
    actions?: unknown; createdRows?: unknown };
  const strings = (value: unknown) => Array.isArray(value)
    ? value.filter((entry): entry is string => typeof entry === "string") : [];
  const rest = { baseVersion: Number(record.baseVersion ?? 0),
    actions: Array.isArray(record.actions) ? strings(record.actions) : ["save-current"],
    createdRows: strings(record.createdRows) };
  if (record.v === 1) {
    const names = strings(record.keys);
    if (!names.length) return null;
    if (!record.content || typeof record.content !== "object" || rest.baseVersion !== loaded.version) return "unreadable";
    const held = record.content as Record<string, unknown>;
    // Laid whole, on the version it was made on: "before" is what that version holds.
    return { v: 2, ...rest, keys: Object.fromEntries(names.map((name) =>
      [name, { kind: "value", before: loaded.content[name] ?? null, after: held[name] ?? null } as PdfHeldKeyDelta])) };
  }
  if (record.v !== 2 || !record.keys || typeof record.keys !== "object" || Array.isArray(record.keys)) return "unreadable";
  const keys: Record<string, PdfHeldKeyDelta> = {};
  for (const [name, value] of Object.entries(record.keys as Record<string, unknown>)) {
    const delta = deltaOf(value);
    if (!delta) return "unreadable";
    keys[name] = delta;
  }
  return Object.keys(keys).length ? { v: 2, ...rest, keys } : null;
}
