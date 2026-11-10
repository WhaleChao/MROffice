/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type PdfNativeTextPropertyProof, type PdfNativeSingleLineFitHint } from "./engine/native/list-text-fidelity";
import { readMraiNativeStandard14Authority } from "./engine/native/standard14-metrics";

const measuredAdvances = new WeakMap<PdfNativeSingleLineFitHint, ReadonlyMap<string, number>>();
export function pdfNativeSourceTextFitsMeasuredFrame(text: string, hint: PdfNativeSingleLineFitHint | null | undefined): boolean {
  if (!hint || !text || text.length > 2_000 || /[\r\n]/.test(text) ||
    !Number.isFinite(hint.capacity) || hint.capacity <= 0 || !Array.isArray(hint.advances) || hint.advances.length > 65_536) return false;
  let widths = measuredAdvances.get(hint);
  if (!widths) {
    const table = new Map<string, number>();
    for (const [character, width] of hint.advances) {
      if (typeof character !== "string" || Array.from(character).length !== 1 || table.has(character) || !Number.isFinite(width) || width < 0) return false;
      table.set(character, width);
    }
    widths = table; measuredAdvances.set(hint, widths);
  }
  if (hint.positionedRuns) {
    if (!Array.isArray(hint.positionedRuns) || hint.positionedRuns.length < 2 || hint.positionedRuns.length > 64 ||
      (hint.sourceSuffix !== undefined && !/^ *$/.test(hint.sourceSuffix))) return false;
    const characters = Array.from(text + (hint.sourceSuffix ?? ""));
    let offset = 0;
    for (const run of hint.positionedRuns) {
      if (!Number.isSafeInteger(run.characters) || run.characters < 1 || !Number.isFinite(run.capacity) || run.capacity <= 0 ||
        !Number.isFinite(run.kerningAdvance) || Math.abs(run.kerningAdvance) > 1_000_000 || offset + run.characters > characters.length) return false;
      if (!Number.isFinite(run.characterSpacingDelta ?? 0) || !Number.isFinite(run.wordSpacingDelta ?? 0)) return false;
      let width = run.kerningAdvance;
      for (const character of characters.slice(offset, offset + run.characters)) {
        const advance = widths.get(character);
        if (advance === undefined) return false;
        width += advance + (run.characterSpacingDelta ?? 0) + (character === hint.wordSpacingCharacter ? run.wordSpacingDelta ?? 0 : 0);
      }
      if (width <= 0 || width > run.capacity + 0.001) return false;
      offset += run.characters;
    }
    return offset === characters.length;
  }
  let total = 0;
  let characterCount = 0;
  // A reader chunk's capacity is the whole show's: the codes around the
  // item are measured with it.
  const measured = (hint.readerChunkPrefix ?? "") + text + (hint.readerChunkSuffix ?? "");
  for (const character of measured) {
    const width = widths.get(character);
    if (width === undefined) return false;
    total += width;
    characterCount += 1;
  }
  // A kept TJ array retains its kerning when the glyph count is unchanged.
  if (
    typeof hint.kerningAdvance === "number" && Number.isFinite(hint.kerningAdvance) &&
    hint.kernedCharacterCount === characterCount
  ) {
    total += hint.kerningAdvance;
  }
  // A layout hint only; the original source-bound compiler still owns save.
  return total > 0 && total <= hint.capacity + 0.001;
}

/** A layout hint only: the source planner/preview still authenticates every
 * write. Browser font substitution and CSS rounding must not force a native
 * single-line run into a reflowed box when its PDF glyph advances still fit.
 * Unknown encodings retain the existing measured-layout path. */
export function pdfNativeSourceTextFitsPropertyFrame(
  text: string,
  proof: PdfNativeTextPropertyProof | null | undefined,
): boolean {
  if (!proof || !text || text.length > 2_000 || /[\r\n]/.test(text)) return false;
  const state = proof.original;
  const encoding = state.font.encoding;
  if (!("baseFont" in encoding)) return false;
  const glyphs = readMraiNativeStandard14Authority(encoding.baseFont, encoding.kind);
  if (!glyphs) return false;
  const byCharacter = new Map(glyphs.map(glyph => [glyph.character, glyph]));
  let units = 0, count = 0, spaces = 0;
  for (const character of text) {
    const glyph = byCharacter.get(character);
    if (!glyph) return false;
    units += glyph.width;
    count++;
    if (glyph.code === 0x20) spaces++;
  }
  const advance = (units * state.fontSize / 1_000 +
    count * state.characterSpacing + spaces * state.wordSpacing) * state.horizontalScale / 100;
  const capacity = proof.frame.right - state.rightIndent -
    (proof.frame.left + state.leftIndent + state.firstLineIndent);
  // This is the existing native property compiler's tolerance, not a new
  // browser overflow allowance. The compiler remains the final authority.
  return Number.isFinite(advance) && Number.isFinite(capacity) &&
    advance > 0 && capacity > 0 && advance <= capacity + 0.001;
}

/** Whether the source line's own metrics cannot say if `text` fits: the
 * line has a measured capacity, but the text holds a character the source
 * font never drew (a letter the font extension will set). Only the exact
 * preview can answer then; the browser's measure of a substitute face is not
 * evidence either way. */
export function pdfNativeSourceTextFitUndecided(
  text: string,
  hint: PdfNativeSingleLineFitHint | null | undefined,
): boolean {
  if (!hint || !text || /[\r\n]/.test(text) || !Array.isArray(hint.advances)) return false;
  const known = new Set(hint.advances.map(([character]) => character));
  return Array.from(text).some((character) => !known.has(character));
}

/** How much wider than the source line's own advance its frame is (>= 1):
 * the room the page proves free beside a one-line run, as a factor of the
 * line's width. Null when neither the property frame nor the measured hint
 * can say. A layout hint only: it sizes the edited line's box, the writer
 * still owns what is saved. */
export function pdfNativeSourceLineRoomFactor(
  originalText: string,
  proof: PdfNativeTextPropertyProof | null | undefined,
  hint: PdfNativeSingleLineFitHint | null | undefined,
): number | null {
  if (!originalText || /[\r\n]/.test(originalText)) return null;
  if (proof && Number.isFinite(proof.extent) && proof.extent > 0) {
    const state = proof.original;
    const capacity = proof.frame.right - state.rightIndent -
      (proof.frame.left + state.leftIndent + state.firstLineIndent);
    const factor = capacity / proof.extent;
    if (Number.isFinite(factor) && factor > 0) return Math.max(1, factor);
  }
  if (hint?.positionedRuns) {
    // A line drawn as several positioned shows (Skia, Chrome and Google
    // Docs set each word and space apart): its capacity is the group's, from
    // the first show's origin to the free space the page proves past the
    // last, and the line is its shows' own advances end to end. Without it
    // such a line had only the margin mirrored from its left edge: none at
    
    // #7b, an author line).
    const runs = hint.positionedRuns;
    if (!Array.isArray(runs) || runs.length < 2 || runs.length > 64 ||
      !Number.isFinite(hint.capacity) || hint.capacity <= 0) return null;
    let extent = 0;
    for (const run of runs) {
      if (!Number.isFinite(run.capacity) || run.capacity <= 0) return null;
      extent += run.capacity;
    }
    return Math.max(1, hint.capacity / extent);
  }
  if (!hint || !Number.isFinite(hint.capacity) || hint.capacity <= 0 ||
    !Array.isArray(hint.advances) || hint.advances.length > 65_536) return null;
  const widths = new Map<string, number>();
  for (const [character, width] of hint.advances) {
    if (typeof character !== "string" || !Number.isFinite(width) || width < 0) return null;
    widths.set(character, width);
  }
  let total = 0, count = 0;
  for (const character of (hint.readerChunkPrefix ?? "") + originalText + (hint.readerChunkSuffix ?? "")) {
    const width = widths.get(character);
    if (width === undefined) return null;
    total += width;
    count += 1;
  }
  if (typeof hint.kerningAdvance === "number" && Number.isFinite(hint.kerningAdvance) &&
    hint.kernedCharacterCount === count) total += hint.kerningAdvance;
  return total > 0 ? Math.max(1, hint.capacity / total) : null;
}
