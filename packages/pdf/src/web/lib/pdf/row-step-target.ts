/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";

/** The page's annotation rows before and after one row edit (every managed
 * object -- comment, shape, ink, stamp, link, text box -- is a row). */
export interface PdfRowStepRows {
  readonly before: readonly JsonValue[];
  readonly after: readonly JsonValue[];
}

function rowId(row: JsonValue): string | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return null;
  const id = (row as { id?: unknown }).id;
  return typeof id === "string" && id.trim() ? id : null;
}

/** A row as a step compares it: as JSON holds it (a key the window left
 * undefined is none, as the page saves it), and without what the reader
 * reads of the PDF's own object it stands for (`nativeSource`: its flags,
 * whether it has an appearance), read again with every saved revision, not
 * changed by an edit. */
function comparable(row: JsonValue | null): JsonValue | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return row;
  const rest: Record<string, JsonValue> = { ...(row as Record<string, JsonValue>) };
  delete rest.nativeSource;
  return JSON.parse(JSON.stringify(rest)) as JsonValue;
}

function sameRow(left: JsonValue | null, right: JsonValue | null): boolean {
  return pageDocumentJsonEquals(comparable(left), comparable(right));
}

function byId(rows: readonly JsonValue[]): Map<string, JsonValue> | null {
  const map = new Map<string, JsonValue>();
  for (const row of rows) {
    const id = rowId(row);
    if (!id || map.has(id)) return null;
    map.set(id, row);
  }
  return map;
}










export function pdfRowStepTarget(
  current: readonly JsonValue[],
  step: PdfRowStepRows,
  direction: "undo" | "redo",
): JsonValue[] | null {
  // A step that changed none of these rows (a form fill's annotations):
  // they stay as they are now.
  if (pageDocumentJsonEquals(step.before as JsonValue[], step.after as JsonValue[])) return [...current];
  const from = byId(direction === "undo" ? step.after : step.before);
  const toRows = direction === "undo" ? step.before : step.after;
  const to = byId(toRows);
  const now = byId(current);
  if (!from || !to || !now) return null;
  const changed = new Set<string>();
  for (const id of new Set([...from.keys(), ...to.keys()])) {
    if (!sameRow(from.get(id) ?? null, to.get(id) ?? null)) changed.add(id);
  }
  // (Rows that differ only in what the reader read of their PDF object:
  // the step changed none of them.)
  if (changed.size === 0) return [...current];
  for (const id of changed) {
    if (!sameRow(now.get(id) ?? null, from.get(id) ?? null)) return null;
  }
  // Rows in place; a changed row replaced, or dropped when the step's other
  // side has none.
  const result: JsonValue[] = [];
  for (const row of current) {
    const id = rowId(row)!;
    if (!changed.has(id)) result.push(row);
    else if (to.has(id)) result.push(to.get(id)!);
  }
  // A row only the other side has comes back after the row it followed there.
  toRows.forEach((row, index) => {
    const id = rowId(row)!;
    if (!changed.has(id) || now.has(id)) return;
    let at = 0;
    for (let previous = index - 1; previous >= 0; previous -= 1) {
      const found = result.findIndex((candidate) => rowId(candidate) === rowId(toRows[previous]!));
      if (found >= 0) { at = found + 1; break; }
    }
    result.splice(at, 0, row);
  });
  return result;
}
