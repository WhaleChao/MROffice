/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where the complete PDF a file starts with ends, when other bytes follow it.
 *
 * A file is read from its end (ISO 32000-1 7.5.5): the last line is `%%EOF`,
 * the lines before it say where the newest cross-reference section is.
 * Readers allow the marker anywhere in the last 1,024 bytes. A file that
 * carries more than that after its marker (a transfer that appended to it, a
 * container's own trailer) has no end a reader finds there, and each opens it
 * its own way: by the last `startxref` anywhere in the file (pdf.js), or by
 * rebuilding the table from the objects. The planner required the file to end
 * with its marker and refused every such file: no line of it opened for
 * editing, and nothing typed, inserted or marked on it was ever saved.
 *
 * Nothing in such a file's chain is misplaced, and nothing has to be searched
 * for: every offset it records is exact, and the PDF simply ends before the
 * file does. So the chain is the one that ends at the file's last `%%EOF`,
 * provided the bytes after it hold nothing a reader could take for PDF: no
 * header, no object, no table or trailer, no second marker. Then there is one
 * chain to find however a reader looks for it, every offset of the complete
 * file is the same offset in the whole one, and a revision appended after
 * those bytes ends the file the way a file has to end, so afterwards every
 * reader follows that same chain.
 *
 * One rule with the server: `_complete_pdf_length_before_trailing_bytes`
 * (native_pdf_primitives.py), which its parser and its incremental writer
 * read such a file by. The two are held to the same answers by the cases of
 * apps/api/tests/file_adapters/pdf_roundtrip/native_complete_pdf_before_trailing_bytes.json.
 */

/**
 * What a reader that rebuilds a file from its objects, or looks for where its
 * chain starts, could take for part of the PDF in bytes that follow it: a
 * header, a table or its `startxref` (both spell `xref`), a trailer, the end
 * of an object, an object's header. The header's white space is PDF's own
 * and every byte the planner's and the server's own patterns read as white
 * space (`\s` of JavaScript and of Python over Latin-1), so nothing either of
 * them could take for an object stands there. One pattern with the server's
 * `_PDF_STRUCTURE_AFTER_THE_END`.
 */
const PDF_STRUCTURE_AFTER_THE_END =
  /%PDF-|xref|trailer|endobj|[0-9][\x00\t\n\x0b\x0c\r\x1c-\x1f \x85\xa0]+[0-9]+[\x00\t\n\x0b\x0c\r\x1c-\x1f \x85\xa0]+obj/g;

/**
 * The length of the complete PDF `source` (a file's bytes read as Latin-1)
 * starts with, when bytes that are not PDF follow it; null for every other
 * file: one that ends with its marker, one that has none, and one whose
 * bytes after the marker a reader could take for PDF.
 */
export function completePdfLengthBeforeTrailingBytes(source: string): number | null {
  const marker = source.lastIndexOf("%%EOF");
  if (marker < 0) return null;
  let end = marker + 5;
  if (source.startsWith("\r\n", end)) end += 2;
  else if (source[end] === "\r" || source[end] === "\n") end += 1;
  if (end === source.length) return null;
  PDF_STRUCTURE_AFTER_THE_END.lastIndex = end;
  return PDF_STRUCTURE_AFTER_THE_END.test(source) ? null : end;
}
