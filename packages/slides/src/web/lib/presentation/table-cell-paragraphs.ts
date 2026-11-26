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
  type OfficeTextLine,
  type OfficeTextParagraph,
} from "@mrnote/core/web/lib/office-text-layout.ts";
import { POWERPOINT_AUTONUM_BY_NUMBER_FORMAT } from "@mrnote/core/web/lib/office-list-catalog.ts";
import { nativeAutoNumberLabel } from "@mrnote/core/web/lib/office-paragraph-render.ts";
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

interface CellParagraphIndent {
  readonly marginLeftPx: number;
  readonly marginRightPx: number;
  readonly indentPx: number;
}

/**
 * A cell's paragraph indents when one of its paragraphs states a right indent
 * (`paragraphIndents`, readings row 1028, TRIAGE 111): one per paragraph of
 * `text` (its "\n"-separated paragraphs), a paragraph typed past the last read
 * taking the last's; null when the cell states none. PowerPoint for Windows
 * takes them from the cell's room on top of the cell's own margins
 * (office-truth/probes/tp-cellmar): centred words between them, right-aligned
 * and justified ones ending at both.
 */
export function presentationTableCellParagraphIndents(
  cell: { readonly paragraphIndents?: unknown },
  text: string,
): CellParagraphIndent[] | null {
  const raw = cell.paragraphIndents;
  if (!Array.isArray(raw) || raw.length === 0) return null;
  const read: CellParagraphIndent[] = [];
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") return null;
    const { marginLeftPx, marginRightPx, indentPx } = entry as Record<string, unknown>;
    const number = (value: unknown) => (typeof value === "number" && Number.isFinite(value) ? value : 0);
    read.push({ marginLeftPx: Math.max(0, number(marginLeftPx)), marginRightPx: Math.max(0, number(marginRightPx)), indentPx: number(indentPx) });
  }
  const count = text.split("\n").length;
  return Array.from({ length: count }, (_, index) => read[Math.min(index, read.length - 1)]);
}

export function presentationTableCellParagraphStyles(
  cell: {
    readonly text?: unknown;
    readonly textLayout?: unknown;
    readonly nativeText?: unknown;
    readonly align?: unknown;
    readonly fontSize?: unknown;
    readonly paragraphIndents?: unknown;
  },
  text: string,
): PresentationCellParagraphStyle[] | undefined {
  const indents = presentationTableCellParagraphIndents(cell, text);
  if (indents) {
    // A cell whose paragraphs state a right indent (readings row 1028): its
    // indents as the bulleted cells above take theirs (margins, a first-line indent).
    return indents.map((indent) => ({
      block: {
        ...(indent.marginLeftPx ? { marginInlineStart: `${indent.marginLeftPx}px` } : {}),
        ...(indent.marginRightPx ? { marginInlineEnd: `${indent.marginRightPx}px` } : {}),
        ...(indent.indentPx ? { textIndent: `${indent.indentPx}px` } : {}),
      },
    }));
  }
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
  cell: { readonly text?: unknown; readonly textLayout?: unknown; readonly nativeText?: unknown; readonly paragraphIndents?: unknown },
  text: string,
  /** Paragraph margins and first-line indents taken too (else they are the browser's to draw). */
  withIndents = false,
  /** Bullets and numbers taken too (readings row 1036), else a marked paragraph is "drawn". */
  withMarkers = false,
): { spaceBeforePt: number; spaceAfterPt: number; marginLeftPt: number; marginRightPt?: number; indentPt: number; lineSpacing?: number; marker?: CellParagraphMarker }[] | "drawn" | null {
  const indents = presentationTableCellParagraphIndents(cell, text);
  if (indents) {
    // Readings row 1028: the indents the cell states (the text layout is never
    // read for such a cell: the importer's reader refuses its right indents).
    if (!withIndents) return "drawn";
    return indents.map((indent) => ({
      spaceBeforePt: 0,
      spaceAfterPt: 0,
      marginLeftPt: indent.marginLeftPx * 0.75,
      marginRightPt: indent.marginRightPx * 0.75,
      indentPt: indent.indentPx * 0.75,
    }));
  }
  const layout = normalizeOfficeTextLayout(cell.textLayout);
  if (!layout) return null;
  const readWith = textOfRuns(cell.nativeText) ?? (typeof cell.text === "string" ? cell.text : text);
  const following = officeTextLayoutFollowingText(layout, readWith, text);
  const lines = officeTextLayoutLines(text, following);
  // Exactly-spaced or Right-to-Left paragraphs (readings rows 831 / 832) are
  // the browser's to draw too.
  if (lines.some((line) => (line.marker !== null && !withMarkers) || line.lineSpacingPoints !== undefined || line.dir === "rtl" || (!withIndents && (line.marginPx || line.indentPx)))) return "drawn";
  const rows = following?.paragraphs;
  const markers = lines.map((line, index) => (line.marker !== null ? cellParagraphMarker(line, rows?.[index]) : undefined));
  // A marker the cell's paragraph names but this reader cannot spell as
  // PowerPoint does (a numbering scheme outside the catalogue, a bullet
  // stating no character, a picture): the browser's, never a guess.
  if (markers.some((marker) => marker === null)) return "drawn";
  if (!lines.some((line, index) => line.spaceBeforePx || line.spaceAfterPx || line.lineSpacing !== undefined || line.marginPx || line.indentPx || markers[index])) return null;
  return lines.map((line, index) => {
    const marker = markers[index];
    return {
      spaceBeforePt: (line.spaceBeforePx ?? 0) * 0.75,
      spaceAfterPt: (line.spaceAfterPx ?? 0) * 0.75,
      marginLeftPt: (line.marginPx ?? 0) * 0.75,
      indentPt: (line.indentPx ?? 0) * 0.75,
      ...(line.lineSpacing !== undefined ? { lineSpacing: line.lineSpacing } : {}),
      ...(marker ? { marker } : {}),
    };
  });
}

/**
 * A cell paragraph's bullet or number as the text engine draws it (readings
 * row 1036): its characters (`officeTextLayoutLines`: the package's
 * `a:buChar`, else the numbering its `a:buAutoNum` counts to), the bullet's
 * own `a:buFont` and character where it states one, and `a:buClr` /
 * `a:buSzPct` / `a:buSzPts`. A marker stating none of those takes its first
 * run's (`a:buFontTx`, `a:buClrTx`, `a:buSzTx`), as in a text box.
 */
export interface CellParagraphMarker {
  readonly text: string;
  /** `a:buChar@char` and `a:buFont@typeface`, as the package states them. */
  readonly bullet?: { readonly char: string; readonly typeface?: string };
  readonly color?: string;
  /** `a:buSzPct`: a multiple of the first run's size. */
  readonly scale?: number;
  /** `a:buSzPts`, points: not scaled by autofit. */
  readonly sizePt?: number;
}

/** A line's marker for the engine, or null when it cannot be spelled as PowerPoint does (see the caller). */
function cellParagraphMarker(line: OfficeTextLine, row: OfficeTextParagraph | undefined): CellParagraphMarker | null {
  if (!row || line.marker === null) return null;
  const native = row.nativeBullet;
  if (native?.picture !== undefined) return null;
  let bullet: CellParagraphMarker["bullet"];
  if (row.bullet === "number") {
    // The package's scheme where the catalogue spells it (`a:buAutoNum@type`),
    // else a gallery format the catalogue maps to one; a numbered paragraph
    // stating neither is numbered by nothing PowerPoint wrote.
    const spelled =
      (native?.autoNum !== undefined && nativeAutoNumberLabel(native.autoNum, 1) !== undefined) ||
      (native?.autoNum === undefined && row.numberFormat !== undefined && Object.hasOwn(POWERPOINT_AUTONUM_BY_NUMBER_FORMAT, row.numberFormat));
    if (!spelled) return null;
  } else if (row.bullet === "bullet") {
    const char = native?.char ?? row.bulletChar;
    if (char === undefined) return null;
    bullet = native?.char !== undefined ? { char: native.char, ...(native.typeface !== undefined ? { typeface: native.typeface } : {}) } : { char };
  } else return null;
  return {
    text: line.marker,
    ...(bullet ? { bullet } : {}),
    ...(line.markerColor !== undefined ? { color: line.markerColor } : {}),
    ...(line.markerScale !== undefined ? { scale: line.markerScale } : {}),
    ...(line.markerSizePt !== undefined ? { sizePt: line.markerSizePt } : {}),
  };
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

