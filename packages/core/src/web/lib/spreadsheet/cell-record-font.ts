/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { numberFormatRewritesText } from "./number-format";










/** What the record says, drawn: for a cell with no runs left to say otherwise.
 * Text typed over a rich cell takes the cell's font, and so does a new cell that
 * copies the record (R102-X24 E8): it is bold, as the file it is saved to says. */
export function recordFontDrawn<T>(style: T): T {
  if (style === null || typeof style !== "object") return style;
  const { cellFontBold, cellFontItalic, cellFontUnderline, cellFontSuperscript, cellFontSubscript, ...rest } =
    style as Record<string, unknown>;
  if (
    cellFontBold === undefined &&
    cellFontItalic === undefined &&
    cellFontUnderline === undefined &&
    cellFontSuperscript === undefined &&
    cellFontSubscript === undefined
  ) {
    return style;
  }
  return {
    ...rest,
    ...(cellFontBold === true ? { bold: true } : {}),
    ...(cellFontItalic === true ? { italic: true } : {}),
    ...(cellFontUnderline === true ? { underline: true } : {}),
    ...(cellFontSuperscript === true && rest.subscript !== true ? { superscript: true } : {}),
    ...(cellFontSubscript === true && rest.superscript !== true ? { subscript: true } : {}),
  } as T;
}

/** A font command on the whole cell says what the record is: what the page said
 * of the record's switch before no longer holds. Off stays off in the file. */
export function dropRecordFontOverridden(
  style: Record<string, unknown>,
  patch: Record<string, unknown>,
  previous?: Record<string, unknown>,
): void {
  // Another face: the class the record stated of its old one (`fontRecord`, by
  // which Windows draws a face it lacks) is not the new face's (xlsx-rounds
  // R426-X90 WALK1: Times New Roman over a charset-1 record read back without it).
  if ("fontFamily" in patch && !("fontRecord" in patch) && patch.fontFamily !== previous?.fontFamily) {
    delete style.fontRecord;
  }
  if ("bold" in patch) delete style.cellFontBold;
  if ("italic" in patch) delete style.cellFontItalic;
  if ("underline" in patch) delete style.cellFontUnderline;
  if ("superscript" in patch || "subscript" in patch) {
    delete style.cellFontSuperscript;
    delete style.cellFontSubscript;
  }
}

/** The cell with its record's font drawn once no run is left to say
 * otherwise (recordFontDrawn): a write that takes a rich cell's runs away --
 * Clear Contents, a fill over it -- leaves the record's bold the cell's, as the
 * file it is saved to and Excel read it (xlsx-rounds R325-X07 WALK1: a page
 * still saying cellFontBold over an emptied cell read back bold). */
export function withRecordFontIfPlain<T extends { runs?: unknown; style?: unknown }>(cell: T): T {
  if (Array.isArray(cell.runs) && cell.runs.length > 0) return cell;
  if (!cell.style || typeof cell.style !== "object") return cell;
  const drawn = recordFontDrawn(cell.style);
  return drawn === cell.style ? cell : { ...cell, style: drawn };
}


/** Whether a rich cell's runs are void as shown: a number format whose text
 * section writes more than `@` (`_(@_)`) shows the text as one string in the
 * cell's own font (numberFormatRewritesText; lookwin p166-lw), so the runs do
 * not draw and the record's switches (recordFontDrawn) are the cell's again. */
export function cellRunsVoidedByFormat(cell: { runs?: unknown; style?: unknown } | undefined): boolean {
  if (!cell || !Array.isArray(cell.runs) || cell.runs.length < 2) return false;
  const style = cell.style;
  return Boolean(style && typeof style === "object" && numberFormatRewritesText((style as { numberFormat?: unknown }).numberFormat));
}
