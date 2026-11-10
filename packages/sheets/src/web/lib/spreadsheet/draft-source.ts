/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const sourceKeys = new WeakMap<object, { value: unknown; formula: unknown; key: string }>();

/** A normalized remote snapshot may replace every cell object. Unrelated
 * remote edits must not discard private typing in an unchanged cell. */
export function spreadsheetDraftSource(cell: { value?: string; formula?: string } | undefined, referenceStyle: string): string {
  if (!cell) return `${referenceStyle}:["",""]`;
  let cached = sourceKeys.get(cell);
  if (!cached || cached.value !== cell.value || cached.formula !== cell.formula) {
    cached = { value: cell.value, formula: cell.formula, key: JSON.stringify([cell.value ?? "", cell.formula ?? ""]) };
    sourceKeys.set(cell, cached);
  }
  return `${referenceStyle}:${cached.key}`;
}
