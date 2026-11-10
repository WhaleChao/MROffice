/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spelling a native text line is edited in.
 *
 * A typesetter (TeX, and the producers that write like it) draws a line as
 * one show and spaces its words with kerning: the source spells no space,
 * while the reader reads one into every gap wide enough (pdf.js). Edited in
 * the source's spelling the line showed its words run together, and a word
 * replaced in it was saved with every gap of the line closed. Given the
 * reader's spelling the planner keeps the gaps and writes a new space as a
 * gap (the kerned writer), so the line is edited as the reader reads it and
 * an untouched draft answers in the source's spelling again, as every
 * "unchanged" rule downstream expects.
 */
export interface PdfNativeReaderSpelling {
  /** The text as the reader reads it: what the box shows and is edited in. */
  readonly reader: string;
  /** The text as the source spells it (the proof's display text). */
  readonly source: string;
  /** The source offset a reader offset stands at (UTF-16 units). */
  readonly sourceOffset: (readerOffset: number) => number;
}

/**
 * The reader's spelling of `source` when it is the source with spaces the
 * reader reads in the gaps added (characters compared as NFKC, as the reader
 * folds a ligature); null when the two differ in anything else, so text the
 * reader leaves out (a joiner, a combining mark) keeps the source spelling.
 */
export function pdfNativeReaderSpelling(
  source: string,
  reader: string | null | undefined,
): PdfNativeReaderSpelling | null {
  if (typeof reader !== "string" || reader === source || !/\s/u.test(reader)) return null;
  const map = new Array<number>(reader.length + 1);
  let i = 0;
  let j = 0;
  let readerOnlySpaces = 0;
  while (j < source.length) {
    const point = source.codePointAt(j)!;
    const unit = point > 0xffff ? 2 : 1;
    const raw = source.slice(j, j + unit);
    const folded = raw.normalize("NFKC");
    if (reader.startsWith(raw, i) || (folded !== raw && reader.startsWith(folded, i))) {
      const length = reader.startsWith(raw, i) ? raw.length : folded.length;
      for (let k = 0; k < length; k += 1) map[i + k] = j;
      i += length;
      j += unit;
      continue;
    }
    if (i < reader.length && /\s/u.test(reader[i]) && !/\s/u.test(raw)) {
      map[i] = j;
      i += 1;
      readerOnlySpaces += 1;
      continue;
    }
    return null;
  }
  while (i < reader.length) {
    if (!/\s/u.test(reader[i])) return null;
    map[i] = j;
    i += 1;
    readerOnlySpaces += 1;
  }
  map[reader.length] = source.length;
  if (readerOnlySpaces === 0) return null;
  // A line set with character spacing (or a TJ number after every glyph)
  // reads a space after nearly every letter, "L I N E 0 9": those are the
  // reader's tracking spaces, not word gaps. Edited in that spelling, the box
  // showed the letter spacing twice, wrapped, and an appended "!" saved 39
  // spaces into the line (user audit PDF 0050). Such a line is edited in the
  // source's own spelling; its letter spacing is the box's.
  const sourceCharacters = Array.from(source);
  let letterPairs = 0;
  for (let index = 1; index < sourceCharacters.length; index += 1) {
    if (!/\s/u.test(sourceCharacters[index - 1]) && !/\s/u.test(sourceCharacters[index])) letterPairs += 1;
  }
  if (readerOnlySpaces >= 3 && readerOnlySpaces * 2 >= letterPairs) return null;
  return Object.freeze({
    reader,
    source,
    sourceOffset: (offset: number) => map[Math.max(0, Math.min(reader.length, offset))] ?? source.length,
  });
}

/**
 * The reader's spelling of one piece of a show: the reader cut the show at a
 * gap wider than a few spaces (a code listing's line number and its code,
 * `[(5)-2625(primes[k])-525(=)...]TJ`), the item selected is the code, and
 * the planner edits that piece alone (its reader chunk) and keeps the codes
 * around it. The box is edited in the piece as the reader reads it; the
 * proof's display text spells the whole show, so an untouched draft answers
 * in it, and a reader offset stands where the piece does in it. Edited in
 * the whole show's spelling, the line number came back as part of the text
 * and was written again in front of the code (user audit PDF 0136). Null
 * when the piece is not found in the show exactly once.
 */
export function pdfNativeReaderChunkSpelling(
  source: string,
  reader: string | null | undefined,
): PdfNativeReaderSpelling | null {
  if (typeof reader !== "string" || !reader.trim() || reader === source) return null;
  // Its letters, the reader's spaces aside, as the reader folds them: a piece
  // of the show is a run of the source's characters that reads as exactly
  // these, tried from each start only as far as it keeps reading so.
  const letters = (text: string) => text.normalize("NFKC").replace(/\s+/gu, "");
  const wanted = letters(reader);
  if (!wanted) return null;
  const pieces: Array<Readonly<{ start: number; spelling: PdfNativeReaderSpelling | null }>> = [];
  for (let start = 0; start < source.length; start += 1) {
    let read = "";
    for (let end = start + 1; end <= source.length; end += 1) {
      read += letters(source.slice(end - 1, end));
      if (!wanted.startsWith(read)) break;
      if (read.length < wanted.length || (start === 0 && end === source.length)) continue;
      const piece = source.slice(start, end);
      const spelling = piece === reader ? null : pdfNativeReaderSpelling(piece, reader);
      if (piece === reader || spelling) pieces.push({ start, spelling });
    }
    if (pieces.length > 1) return null;
  }
  if (pieces.length !== 1) return null;
  const [{ start, spelling }] = pieces;
  return Object.freeze({
    reader,
    source,
    sourceOffset: (offset: number) =>
      start + (spelling ? spelling.sourceOffset(offset) : Math.max(0, Math.min(reader.length, offset))),
  });
}

/** The text a draft answers with: the source's spelling while it still reads
 * as the line did, else the reader-spelled text as typed. */
export function pdfNativeTextAsCommitted(
  spelling: PdfNativeReaderSpelling | null,
  text: string,
): string {
  return spelling && text === spelling.reader ? spelling.source : text;
}

/** Ranges of an untouched draft's text, moved onto the source's spelling. */
export function pdfNativeRangesAsCommitted<T extends Readonly<{ start: number; end: number }>>(
  spelling: PdfNativeReaderSpelling | null,
  text: string,
  ranges: readonly T[] | undefined,
): readonly T[] | undefined {
  if (!spelling || text !== spelling.reader || !ranges) return ranges;
  return ranges.map((range) => ({
    ...range,
    start: spelling.sourceOffset(range.start),
    end: spelling.sourceOffset(range.end),
  }));
}

/** Whether a record that spells `recorded` holds the reader-spelled text a
 * box saved (the planner records the source's spelling of what it wrote). */
export function pdfNativeRecordSpellsReaderText(recorded: string, typed: string): boolean {
  return recorded === typed || pdfNativeReaderSpelling(recorded, typed) !== null;
}
