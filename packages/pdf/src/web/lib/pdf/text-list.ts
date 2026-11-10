/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bullets and numbering for PDF text (ONLYOFFICE's and Word's 项目符号 /
 * 编号): a PDF has no list structure, so a list is its markers written as
 * text at the start of each paragraph -- the glyph or label and a space --
 * with a hanging indent that sets the paragraph's other rows under its
 * words. The markers are the presentation editor's own library (its bullet
 * glyphs, its numbering formats and their spellings), so a list looks the
 * same in every editor.
 */
import {
  POWERPOINT_BULLET_LIBRARY,
  powerPointNumberLibrary,
  type PresentationListLibraryItem,
} from "@mrnote/core/web/lib/presentation-bullet-library.ts";
import { nativeAutoNumberLabel, numberLabelForFormat } from "@mrnote/core/web/lib/office-paragraph-render.ts";
import {
  PDF_TEXT_LIST_BULLET_GLYPHS as BULLET_GLYPHS,
  PDF_TEXT_LIST_NUMBER_LABEL as NUMBER_LABEL,
  pdfTextListMarkerOf,
} from "./text-list-marker";

export { pdfTextListMarkerOf };

export type PdfTextListStyle = Readonly<
  | { kind: "bullet"; id: string; glyph: string }
  | { kind: "number"; id: string }
>;

/** What separates a marker from its paragraph's words. */
const MARKER_GAP = " ";

function libraryItem(style: PdfTextListStyle, locale: string): PresentationListLibraryItem | undefined {
  return style.kind === "bullet"
    ? POWERPOINT_BULLET_LIBRARY.find((item) => item.id === style.id)
    : powerPointNumberLibrary(locale).find((item) => item.id === style.id);
}

/** The label of a numbered paragraph (1-based), as the library spells it. */
export function pdfTextListNumberLabel(id: string, ordinal: number, locale = "zh-CN"): string {
  const item = powerPointNumberLibrary(locale).find((entry) => entry.id === id);
  if (item?.kind !== "number") return numberLabelForFormat("decimal", ordinal);
  if (item.autoNum) return nativeAutoNumberLabel(item.autoNum, ordinal) ?? `${ordinal}.`;
  return numberLabelForFormat(item.format ?? "decimal", ordinal);
}

/** The list the text is: every paragraph with words starts with a bullet
 * (its glyph), or every one with a label; else none. */
export function pdfTextListKindOf(text: string): Readonly<{ kind: "bullet"; glyph: string } | { kind: "number" }> | null {
  const paragraphs = text.split("\n").filter((paragraph) => paragraph.trim());
  if (paragraphs.length === 0) return null;
  const markers = paragraphs.map(pdfTextListMarkerOf);
  if (markers.some((marker) => marker === null)) return null;
  const glyphs = new Set(markers.map((marker) => [...marker!][0]));
  if (glyphs.size === 1) {
    const glyph = [...glyphs][0]!;
    if (BULLET_GLYPHS.has(glyph)) return { kind: "bullet", glyph };
  }
  return markers.every((marker) => NUMBER_LABEL.test(marker!)) ? { kind: "number" } : null;
}

/**
 * The text with each paragraph that has words set as an item of the list
 * (its old marker, if any, replaced), or with every marker taken off when
 * `style` is null. Empty paragraphs are left as they are (Word numbers
 * only the paragraphs it lists).
 */
export function applyPdfTextList(text: string, style: PdfTextListStyle | null, locale = "zh-CN"): string {
  const item = style ? libraryItem(style, locale) : undefined;
  let ordinal = 0;
  return text.split("\n").map((paragraph) => {
    if (!paragraph.trim()) return paragraph;
    const marker = pdfTextListMarkerOf(paragraph);
    const words = marker ? paragraph.slice(marker.length) : paragraph;
    if (!style) return words;
    if (style.kind === "bullet") return `${item?.preview ?? style.glyph}${MARKER_GAP}${words}`;
    ordinal += 1;
    return `${pdfTextListNumberLabel(style.id, ordinal, locale)}${MARKER_GAP}${words}`;
  }).join("\n");
}

/** The width a marker takes (its widest, with its gap), for the hanging
 * indent: measured in the text's own face, in the face's size units. */
export function pdfTextListMarkerWidth(text: string, measure: (value: string) => number): number {
  let widest = 0;
  for (const paragraph of text.split("\n")) {
    const marker = pdfTextListMarkerOf(paragraph);
    if (marker) widest = Math.max(widest, measure(marker));
  }
  return widest;
}
