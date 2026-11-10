/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a page's text layer is built from, as one string: two revisions of a
 * file whose page reads the same (a reader mark added, an annotation moved)
 * give the same string, and the page keeps the text layer it shows instead
 * of building it again (PDFViewer prepareNativeTextOfPages; in WebKit the
 * rebuild of the pages on screen after each saved mark was the bulk of a
 * second of long frames).
 *
 * pdf.js names a page's fonts per document (`g_d3_f1`), so a font is named
 * here by where it first appears on the page, and its style read under that
 * name. Everything else the text items carry is kept exactly: a letter, a
 * position, a width or a font that differs gives another string.
 */
type TextStyleLike = Readonly<{ fontFamily?: unknown; ascent?: unknown; descent?: unknown; vertical?: unknown }>;

export function pdfNativeTextContentSignature(input: Readonly<{
  items: readonly unknown[];
  styles: Readonly<Record<string, TextStyleLike | undefined>>;
  hardGroups: readonly (string | null)[] | null;
  artifacts: readonly boolean[] | null;
  /** What else the layer is built from, as the caller reads it (the
   * viewport, the session's kept blocks, the page's saved boxes). */
  context: unknown;
}>): string {
  const fonts = new Map<string, number>();
  const fontOf = (name: unknown) => {
    if (typeof name !== "string") return null;
    let ordinal = fonts.get(name);
    if (ordinal === undefined) { ordinal = fonts.size; fonts.set(name, ordinal); }
    return ordinal;
  };
  const items = input.items.map((raw) => {
    if (!raw || typeof raw !== "object") return null;
    const item = raw as Record<string, unknown>;
    if (typeof item.str !== "string") return ["m", item.type ?? null, item.id ?? null, item.tag ?? null];
    return [item.str, item.dir ?? null, item.transform ?? null, item.width ?? null, item.height ?? null,
      fontOf(item.fontName), item.hasEOL ?? null];
  });
  const styles = [...fonts].map(([name]) => {
    const style = input.styles[name];
    return style ? [style.fontFamily ?? null, style.ascent ?? null, style.descent ?? null, style.vertical ?? null] : null;
  });
  return JSON.stringify([items, styles, input.hardGroups, input.artifacts, input.context ?? null]);
}
