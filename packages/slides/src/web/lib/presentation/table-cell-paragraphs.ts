/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationCellParagraphStyle } from "@mrnote/core/web/components/notebook/contents/PresentationCellRichTextEditor.tsx";
import {
  normalizeOfficeTextLayout,
  officeTextLayoutFollowingText,
  officeTextLayoutLines,
  officeTextLineIndentCss,
} from "@mrnote/core/web/lib/office-text-layout.ts";

function textOfRuns(value: unknown): string | null {
  if (!value || typeof value !== "object") return null;
  const runs = (value as { runs?: unknown }).runs;
  if (!Array.isArray(runs)) return null;
  let text = "";
  for (const run of runs) {
    if (!run || typeof run !== "object" || typeof (run as { text?: unknown }).text !== "string") return null;
    text += (run as { text: string }).text;
  }
  return text;
}

export function presentationTableCellParagraphStyles(
  cell: { readonly text?: unknown; readonly textLayout?: unknown; readonly nativeText?: unknown },
  text: string,
): PresentationCellParagraphStyle[] | undefined {
  const layout = normalizeOfficeTextLayout(cell.textLayout);
  if (!layout) return undefined;
  const readWith = textOfRuns(cell.nativeText) ?? (typeof cell.text === "string" ? cell.text : text);
  const lines = officeTextLayoutLines(text, officeTextLayoutFollowingText(layout, readWith, text));
  if (!lines.some((line) => line.marker !== null || line.marginPx || line.indentPx || line.spaceBeforePx || line.spaceAfterPx)) {
    return undefined;
  }
  return lines.map((line, index) => {
    const css = officeTextLineIndentCss(line);
    // The space between paragraphs PowerPoint adds (`a:spcBef` / `a:spcAft`,
    
    return {
      block: {
        ...css.block,
        ...(index > 0 && line.spaceBeforePx ? { paddingTop: `${line.spaceBeforePx}px` } : {}),
        ...(index < lines.length - 1 && line.spaceAfterPx ? { paddingBottom: `${line.spaceAfterPx}px` } : {}),
      },
      ...(line.marker !== null ? { marker: { text: line.marker, css: css.marker } } : {}),
    };
  });
}

/** A cell's runs split into its paragraphs (at each "\n", which no paragraph keeps). */
export function presentationTableCellRunParagraphs<R extends { text: string }>(runs: readonly R[]): R[][] {
  const paragraphs: R[][] = [[]];
  for (const run of runs) {
    const pieces = run.text.split("\n");
    pieces.forEach((piece, index) => {
      if (index) paragraphs.push([]);
      if (piece) paragraphs[paragraphs.length - 1].push({ ...run, text: piece });
    });
  }
  return paragraphs;
}
