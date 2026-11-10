/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function upsertAnnotationRowById<T extends { id: string }>(
  rows: readonly T[],
  created: T,
): T[] {
  const index = rows.findIndex((row) => row.id === created.id);
  if (index < 0) return [...rows, created];
  return rows.map((row, rowIndex) => (rowIndex === index ? created : row));
}
