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
import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import { exactLineSpacingCss } from "@mrnote/core/web/lib/presentation-line-spacing.ts";
import type { RunStyle } from "@mrnote/core/web/lib/presentation-runs.ts";

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
  cell: {
    readonly text?: unknown;
    readonly textLayout?: unknown;
    readonly nativeText?: unknown;
    readonly align?: unknown;
    readonly fontSize?: unknown;
  },
  text: string,
): PresentationCellParagraphStyle[] | undefined {
  const layout = normalizeOfficeTextLayout(cell.textLayout);
  if (!layout) return undefined;
  const readWith = textOfRuns(cell.nativeText) ?? (typeof cell.text === "string" ? cell.text : text);
  const lines = officeTextLayoutLines(text, officeTextLayoutFollowingText(layout, readWith, text));
  if (
    !lines.some(
      (line) =>
        line.marker !== null || line.marginPx || line.indentPx || line.spaceBeforePx || line.spaceAfterPx ||
        line.lineSpacingPoints !== undefined || line.dir === "rtl",
    )
  ) {
    return undefined;
  }
  const typeSizePt = typeof cell.fontSize === "number" && cell.fontSize > 0 ? cell.fontSize : undefined;
  const leftAligned = cell.align === undefined || cell.align === null || cell.align === "left";
  return lines.map((line, index) => {
    const css = officeTextLineIndentCss(line);
    // The space between paragraphs PowerPoint adds (`a:spcBef` / `a:spcAft`,
    
    return {
      block: {
        ...css.block,
        ...(index > 0 && line.spaceBeforePx ? { paddingTop: `${line.spaceBeforePx}px` } : {}),
        ...(index < lines.length - 1 && line.spaceAfterPx ? { paddingBottom: `${line.spaceAfterPx}px` } : {}),
        // Line Spacing > Exactly on this paragraph (readings row 832): its
        // lines that tall, as a text box's row (`exactLineSpacingCss`).
        ...(line.lineSpacingPoints !== undefined ? exactLineSpacingCss(line.lineSpacingPoints, typeSizePt) : {}),
        // A Right-to-Left paragraph (readings row 831): its base direction,
        // which orders its words and puts its marker and margin on the right.
        // Its alignment is a physical side, Left where it states none
        // (PowerPoint's print of truth-probes/lookfit-rtl-align-probe.pptx).
        ...(line.dir === "rtl" ? { direction: "rtl" as const, ...(leftAligned ? { textAlign: "left" as const } : {}) } : {}),
      },
      ...(line.marker !== null ? { marker: { text: line.marker, css: css.marker } } : {}),
    };
  });
}










export function presentationTableCellParagraphSpacing(
  cell: { readonly text?: unknown; readonly textLayout?: unknown; readonly nativeText?: unknown },
  text: string,
  /** Paragraph margins and first-line indents taken too (else they are the browser's to draw). */
  withIndents = false,
): { spaceBeforePt: number; spaceAfterPt: number; marginLeftPt: number; indentPt: number; lineSpacing?: number }[] | "drawn" | null {
  const layout = normalizeOfficeTextLayout(cell.textLayout);
  if (!layout) return null;
  const readWith = textOfRuns(cell.nativeText) ?? (typeof cell.text === "string" ? cell.text : text);
  const lines = officeTextLayoutLines(text, officeTextLayoutFollowingText(layout, readWith, text));
  // Exactly-spaced or Right-to-Left paragraphs (readings rows 831 / 832) are
  // the browser's to draw too.
  if (lines.some((line) => line.marker !== null || line.lineSpacingPoints !== undefined || line.dir === "rtl" || (!withIndents && (line.marginPx || line.indentPx)))) return "drawn";
  if (!lines.some((line) => line.spaceBeforePx || line.spaceAfterPx || line.lineSpacing !== undefined || line.marginPx || line.indentPx)) return null;
  return lines.map((line) => ({
    spaceBeforePt: (line.spaceBeforePx ?? 0) * 0.75,
    spaceAfterPt: (line.spaceAfterPx ?? 0) * 0.75,
    marginLeftPt: (line.marginPx ?? 0) * 0.75,
    indentPt: (line.indentPx ?? 0) * 0.75,
    ...(line.lineSpacing !== undefined ? { lineSpacing: line.lineSpacing } : {}),
  }));
}







export function presentationTableCellParagraphBreakStyles<S>(runs: readonly { text: string; style?: S }[]): (S | undefined)[] {
  const styles: (S | undefined)[] = [];
  for (const run of runs) {
    for (let index = run.text.indexOf("\n"); index !== -1; index = run.text.indexOf("\n", index + 1)) styles.push(run.style);
  }
  styles.push(undefined);
  return styles;
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

















export function presentationTableCellTrailingParagraphStyle(
  table: { source?: unknown },
  cell: { text?: unknown; nativeEndMark?: unknown; fontSize?: unknown; fontFamily?: unknown },
): RunStyle | null {
  // Only a table read under readings row 1004 says so: a page stored before
  
  const source = table.source && typeof table.source === "object" ? (table.source as Record<string, unknown>) : null;
  if (source?.cellTrailingParagraphLines !== true) return null;
  if (typeof cell.text !== "string" || !cell.text.endsWith("\n")) return null;
  const mark = cell.nativeEndMark;
  if (!mark || typeof mark !== "object" || Array.isArray(mark)) return {};
  const record = mark as Record<string, unknown>;
  const read = record.cell && typeof record.cell === "object" && !Array.isArray(record.cell)
    ? (record.cell as Record<string, unknown>)
    : {};
  const style: RunStyle = {};
  if (
    typeof record.fontSize === "number" && Number.isFinite(record.fontSize) && record.fontSize > 0 &&
    cell.fontSize === read.fontSize
  ) {
    style.fontSize = ptToPx(record.fontSize);
  }
  if (typeof record.fontFamily === "string" && record.fontFamily.trim() && cell.fontFamily === read.fontFamily) {
    style.fontFamily = record.fontFamily;
  }
  if (typeof record.baseline === "number" && record.baseline > 0) style.superscript = true;
  else if (typeof record.baseline === "number" && record.baseline < 0) style.subscript = true;
  return style;
}

