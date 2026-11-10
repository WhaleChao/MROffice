/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type Content = Record<string, unknown>;

const same = (left: unknown, right: unknown) => JSON.stringify(left ?? null) === JSON.stringify(right ?? null);

function rowId(row: unknown): string | null {
  const id = row && typeof row === "object" && !Array.isArray(row) ? (row as { id?: unknown }).id : null;
  return typeof id === "string" && id ? id : null;
}

/**
 * The page a step's Undo or Redo writes: the page as it is now, with only
 * what the step itself changed set to the step's other side.
 *
 * A step keeps the page as it was before and after it (a page turned, moved,
 * deleted or added). Written whole, that snapshot also took back everything
 * saved since -- the edits published on the file after it (its native journal
 * lost them, and the server refused the save as a second writer), notes,
 * boxes. So only the collections that differ between the step's two sides are
 * replaced, and in the annotation rows only the rows that differ (a page's
 * decorations re-drawn for its new place), each by its id.
 */
export function pageStepContent(current: Content, target: Content, other: Content): Content {
  const next: Content = { ...current };
  for (const key of new Set([...Object.keys(target), ...Object.keys(other)])) {
    if (same(target[key], other[key])) continue;
    if (key === "annotations" && Array.isArray(current.annotations)) {
      const before = new Map<string, unknown>(), after = new Map<string, unknown>();
      for (const row of Array.isArray(other[key]) ? other[key] as unknown[] : []) {
        const id = rowId(row); if (id) before.set(id, row);
      }
      for (const row of Array.isArray(target[key]) ? target[key] as unknown[] : []) {
        const id = rowId(row); if (id) after.set(id, row);
      }
      const touched = new Set([...before.keys(), ...after.keys()].filter((id) => !same(before.get(id), after.get(id))));
      const rows: unknown[] = [];
      const placed = new Set<string>();
      for (const row of current.annotations as unknown[]) {
        const id = rowId(row);
        if (!id || !touched.has(id)) { rows.push(row); continue; }
        placed.add(id);
        if (after.has(id)) rows.push(after.get(id));
      }
      for (const [id, row] of after) if (touched.has(id) && !placed.has(id)) rows.push(row);
      next.annotations = rows;
      continue;
    }
    if (key in target) next[key] = target[key];
    else delete next[key];
  }
  return next;
}
