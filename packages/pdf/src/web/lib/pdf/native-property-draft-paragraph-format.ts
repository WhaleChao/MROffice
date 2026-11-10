/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A run's property draft (its whole state and the parts set on their own) as
 * the paragraph format the paragraph rewrite sets: the same formats, read as
 * factors of each character's own. The run's own write cannot set words its
 * embedded subset has no glyph for (a LibreOffice slide's title typed into);
 * the font extension's paragraph route sets them in the substitute face with
 * these formats. Null when the draft asks what a paragraph format does not
 * say (a move, a turn, spacing, indents).
 */
import type {
  PdfNativeParagraphFormat,
  PdfNativeTextPropertyDraft,
  PdfNativeTextPropertyState,
} from "./engine/native/list-text-fidelity";
import { pdfNativeParagraphFormatOfPropertyDraft as paragraphFormatOfPropertyDraft } from "./engine/native/list-text-fidelity";
import { pdfNativeEmbeddedFontKey, readPdfNativeEmbeddedFontKey } from "./native-embedded-font-key";

export function pdfNativeParagraphFormatOfPropertyDraft(
  draft: PdfNativeTextPropertyDraft,
  original: PdfNativeTextPropertyState,
  _textLength?: number,
): PdfNativeParagraphFormat | null {
  // One reading of a draft as a paragraph format: the engine's, which its
  // own paragraph fallback takes too. Two of them disagreed -- this one
  
  return paragraphFormatOfPropertyDraft(original, draft);
}

/** The format with no slant anywhere: the substitute face has no italic, and
 * a slant set from inside a word reads as a line of its own (the rewrite
 * refuses it). Null when the format sets no slant. */
export function pdfNativeParagraphFormatWithoutSlant(format: PdfNativeParagraphFormat): PdfNativeParagraphFormat | null {
  if (format.slanted === undefined && !format.parts?.some((part) => part.slanted !== undefined)) return null;
  const { slanted: _slanted, parts, ...rest } = format;
  void _slanted;
  const kept = (parts ?? []).flatMap((part) => {
    const { slanted: _partSlanted, ...fields } = part;
    void _partSlanted;
    return Object.keys(fields).some((key) => key !== "start" && key !== "end") ? [fields] : [];
  });
  return Object.freeze({ ...rest, ...(kept.length ? { parts: kept } : {}) });
}








export function pdfNativeParagraphFormatWithoutMidWordLean(
  format: PdfNativeParagraphFormat,
  text: string,
): PdfNativeParagraphFormat | null {
  const inWord = (at: number) => at > 0 && at < text.length && /\S/u.test(text[at - 1]) && /\S/u.test(text[at]);
  let changed = false;
  const parts = (format.parts ?? []).flatMap((part) => {
    if (!inWord(part.start) && !inWord(part.end)) return [part];
    const asked = readPdfNativeEmbeddedFontKey(part.fontResourceName);
    if (part.slanted === undefined && !asked?.italic) return [part];
    changed = true;
    const { slanted: _slanted, fontResourceName, ...fields } = part;
    void _slanted;
    const upright = asked?.italic ? pdfNativeEmbeddedFontKey({ ...asked, italic: false }) : fontResourceName;
    const next = { ...fields, ...(upright ? { fontResourceName: upright } : {}) };
    return Object.keys(next).some((key) => key !== "start" && key !== "end") ? [Object.freeze(next)] : [];
  });
  if (!changed) return null;
  const { parts: _parts, ...rest } = format;
  void _parts;
  return Object.freeze({ ...rest, ...(parts.length ? { parts: Object.freeze(parts) } : {}) });
}







export function pdfNativeParagraphFormatWithoutMidWordRise(
  format: PdfNativeParagraphFormat,
  text: string,
): PdfNativeParagraphFormat | null {
  const inWord = (at: number) => at > 0 && at < text.length && /\S/u.test(text[at - 1]) && /\S/u.test(text[at]);
  let changed = false;
  const parts = (format.parts ?? []).flatMap((part) => {
    if (part.riseEm === undefined || (!inWord(part.start) && !inWord(part.end))) return [part];
    changed = true;
    const { riseEm: _rise, sizeScale: _size, sizePoints: _points, ...fields } = part;
    void _rise; void _size; void _points;
    return Object.keys(fields).some((key) => key !== "start" && key !== "end") ? [Object.freeze(fields)] : [];
  });
  if (!changed) return null;
  const { parts: _parts, ...rest } = format;
  void _parts;
  return Object.freeze({ ...rest, ...(parts.length ? { parts: Object.freeze(parts) } : {}) });
}
