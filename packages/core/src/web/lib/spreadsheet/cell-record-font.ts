/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** What the record says, drawn: for a cell with no runs left to say otherwise.
 * Text typed over a rich cell takes the cell's font, and so does a new cell that
 * copies the record (R102-X24 E8): it is bold, as the file it is saved to says. */
export function recordFontDrawn<T>(style: T): T {
  if (style === null || typeof style !== "object") return style;
  const { cellFontBold, cellFontItalic, cellFontUnderline, ...rest } = style as Record<string, unknown>;
  if (cellFontBold === undefined && cellFontItalic === undefined && cellFontUnderline === undefined) return style;
  return {
    ...rest,
    ...(cellFontBold === true ? { bold: true } : {}),
    ...(cellFontItalic === true ? { italic: true } : {}),
    ...(cellFontUnderline === true ? { underline: true } : {}),
  } as T;
}

/** A font command on the whole cell says what the record is: what the page said
 * of the record's switch before no longer holds. Off stays off in the file. */
export function dropRecordFontOverridden(style: Record<string, unknown>, patch: Record<string, unknown>): void {
  if ("bold" in patch) delete style.cellFontBold;
  if ("italic" in patch) delete style.cellFontItalic;
  if ("underline" in patch) delete style.cellFontUnderline;
}
