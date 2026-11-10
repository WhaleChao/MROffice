/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function pdfNativeBoxRowStarts(boxText: string, rows: readonly Readonly<{ str: string }>[]): number[] | null {
  const starts: number[] = [];
  let cursor = 0;
  for (const row of rows) {
    const at = boxText.indexOf(row.str, cursor);
    if (at < 0) return null;
    starts.push(at);
    cursor = at + row.str.length;
  }
  return starts;
}
