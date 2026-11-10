"use client";



















/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useContext, useEffect, useMemo, useState, type CSSProperties, type ReactElement } from "react";
import type { PresentationTableBlock, PresentationTableCell } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { layoutOfficeTextBody, type OfficeParagraphInput, type OfficeTextBodyLayout, type OfficeTextFont } from "@mrnote/core/web/lib/text-engine/office-text-body.ts";
import { pageOfficeTextMeasurer } from "@mrnote/core/web/lib/text-engine/office-text-measurer.ts";
import { onTextEngineChange, textEngineEpoch } from "@mrnote/core/web/lib/text-engine/measure.ts";
import { textEngineOn } from "@mrnote/core/web/lib/text-engine/page-adapter.ts";
import { cellCharacterRuns, cellTextRunCss } from "@mrnote/core/web/lib/presentation-cell-rich-text.ts";
import { normalizeCellStyle } from "@mrnote/core/web/lib/presentation-table-style.ts";
import { presentationTableCellParagraphStyles, presentationTableCellRunParagraphs } from "../../../../lib/presentation/table-cell-paragraphs";
import {
  TABLE_CELL_NORMAL_MARGIN_PX,
  tableCellVisualTextStyle,
  tableStoredRowHeights,
} from "../../../../lib/presentation/table-canvas-layout";
import { logicalCellPositions, tableLogicalWidth } from "@mrnote/core/web/lib/presentation-table-merge.ts";
import { asRecord } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import { isImportedPptxBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { PRESENTATION_TEXT_ENGINE_ATTRIBUTE, PresentationTextEngineSurfaceContext } from "./PresentationEngineTextBody";

const PT_PER_PX = 0.75;
const EMU_PER_PX = 9_525;
const PX_PER_PT = 4 / 3;

class Unsupported extends Error {}

/** A CSS length in px: a number, `Npx`, `Npt`, or `Nem` of `emPx`. */
function lengthPx(value: unknown, emPx: number): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = typeof value === "string" ? /^\s*(-?[0-9]*\.?[0-9]+)(px|pt|em)?\s*$/i.exec(value) : null;
  if (!match) throw new Unsupported(`length ${String(value)}`);
  const number = Number(match[1]);
  const unit = (match[2] ?? "px").toLowerCase();
  return unit === "pt" ? number * PX_PER_PT : unit === "em" ? number * emPx : number;
}

function weight(value: unknown): number {
  if (value === undefined || value === null || value === "normal") return 400;
  if (value === "bold") return 700;
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Unsupported("weight");
  return number;
}

export interface EngineTableCell {
  readonly layout: Extract<OfficeTextBodyLayout, { status: "ok" }>;
  /** The cell's runs per paragraph, as drawn. */
  readonly paragraphs: readonly (readonly TextRun[])[];
  /** Each run's CSS, per paragraph, as the browser-flowed cell draws it. */
  readonly runCss: readonly (readonly CSSProperties[])[];
  readonly fonts: readonly (readonly OfficeTextFont[])[];
}

export type EngineTableAnswer =
  | { readonly status: "ok"; readonly rowHeightsPx: readonly number[]; readonly cells: readonly (readonly (EngineTableCell | null)[])[] }
  | { readonly status: "pending" }
  | { readonly status: "off" };

function cellInput(
  table: PresentationTableBlock,
  cell: PresentationTableCell,
  rowIndex: number,
  kernFrom: number | null,
  tableFace: string | null,
): { paragraphs: OfficeParagraphInput[]; drawn: TextRun[][]; css: CSSProperties[][]; fonts: OfficeTextFont[][] } {
  const style = normalizeCellStyle(cell);
  if (style.textDirection !== "horizontal") throw new Unsupported("turned words");
  if (style.textPaint) throw new Unsupported("cell WordArt");
  if (presentationTableCellParagraphStyles(cell, cell.text)) throw new Unsupported("cell markers or indents");
  const header = table.headerRow === true && rowIndex === 0;
  const base = tableCellVisualTextStyle(table, { ...style, superscript: false, subscript: false }, header, false, cell);
  const baseSizePx = lengthPx(base.fontSize ?? 24, 16);
  const runs = cellCharacterRuns(String(cell.text ?? ""), (cell as { runs?: unknown }).runs);
  const split = presentationTableCellRunParagraphs(runs);
  const lineHeight = typeof base.lineHeight === "number" ? base.lineHeight : 1.2;
  const align = base.textAlignLast === "justify" ? "dist" : base.textAlign === "center" ? "ctr" : base.textAlign === "right" ? "r" : base.textAlign === "justify" ? "just" : "l";
  const fontOf = (css: CSSProperties): OfficeTextFont => {
    // A cell that states no face (an empty one, words inheriting the
    // table's) takes the face the table's words are in.
    const stack = css.fontFamily ?? base.fontFamily ?? tableFace;
    if (typeof stack !== "string" || !stack.trim()) throw new Unsupported("no face");
    if (css.fontVariantCaps === "small-caps") throw new Unsupported("small capitals");
    const sizePt = Math.round(lengthPx(css.fontSize ?? base.fontSize ?? 24, baseSizePx) * PT_PER_PX * 100) / 100;
    return {
      stack,
      weight: weight(css.fontWeight ?? base.fontWeight),
      italic: (css.fontStyle ?? base.fontStyle) === "italic",
      sizePt,
      ...(kernFrom !== null && sizePt >= kernFrom - 1e-6 ? { kerning: true } : {}),
    };
  };
  const endFont = fontOf({});
  const drawn: TextRun[][] = [];
  const cssRows: CSSProperties[][] = [];
  const fontRows: OfficeTextFont[][] = [];
  const paragraphs = split.map((paragraph): OfficeParagraphInput => {
    const css = paragraph.map((run) => {
      if ((run as { math?: unknown }).math) throw new Unsupported("equation");
      if (run.style?.textPaint) throw new Unsupported("run WordArt");
      if (/[-]/.test(run.text)) throw new Unsupported("symbol-font characters");
      return { ...base, ...cellTextRunCss(run.style, base, { superscript: style.superscript, subscript: style.subscript }) };
    });
    const fonts = css.map(fontOf);
    drawn.push(paragraph);
    cssRows.push(css);
    fontRows.push(fonts);
    return {
      runs: paragraph.map((run, index) => ({
        text: run.text,
        font: fonts[index],
        ...(css[index].textTransform === "uppercase" ? { allCaps: true } : {}),
        ...(typeof css[index].letterSpacing === "string"
          ? { letterSpacingPt: Math.round(lengthPx(css[index].letterSpacing, 16) * PT_PER_PX * 100) / 100 }
          : {}),
      })),
      endFont: fonts[fonts.length - 1] ?? endFont,
      align,
      marginLeftPt: 0,
      marginRightPt: 0,
      indentPt: 0,
      lineSpacing: { kind: "percent", value: lineHeight / 1.2 },
      spaceBeforePt: 0,
      spaceAfterPt: 0,
      kinsoku: "zh-CN",
    };
  });
  return { paragraphs, drawn, css: cssRows, fonts: fontRows };
}

/** Where a cell's last whole line ends (its last line's full pitch, not its descent), pt. */
function cellLinesBottomPt(layout: Extract<OfficeTextBodyLayout, { status: "ok" }>): number {
  const lastLines = layout.paragraphs.map((paragraph) => paragraph.lines[paragraph.lines.length - 1]).filter(Boolean);
  return lastLines.length ? Math.max(...lastLines.map((line) => line.bottom)) : layout.height;
}

/** Lay every cell of `table` out and work out its rows' heights, in px. */
export function layoutPresentationEngineTable(table: PresentationTableBlock): EngineTableAnswer {
  try {
    const rowCount = table.rows.length;
    const columnCount = tableLogicalWidth(table.rows);
    const positions = logicalCellPositions(table.rows);
    const colWidths = Array.isArray(table.colWidths) ? table.colWidths : null;
    const rowHeights = Array.isArray(table.rowHeights) ? table.rowHeights : null;
    if (!colWidths || colWidths.length !== columnCount || !rowHeights || rowHeights.length !== rowCount) {
      return tableOff("grid sizes");
    }
    const tableFace = (() => {
      for (const row of table.rows) {
        for (const cell of row) {
          const family = normalizeCellStyle(cell).fontFamily ??
            cellCharacterRuns(String(cell.text ?? ""), (cell as { runs?: unknown }).runs).find((run) => run.style?.fontFamily)?.style?.fontFamily;
          if (typeof family === "string" && family.trim()) return family;
        }
      }
      return null;
    })();
    const kerning = asRecord(table.source)?.nativeKerning;
    const kernFrom = typeof kerning === "number" && kerning > 0 ? kerning : null;
    // Each row from its stored height (`tableStoredRowHeights`: the page's, before
    // the frame fit), a minimum its words grow (ECMA-376 21.1.3.18 `a:tr@h`),
    // not the frame-fitted size, which spreads the frame's cached extent over
    
    // 7.2 pt apart in PowerPoint, drawn at the frame's average 6.9 pt). A row
    // still as imported takes its exact `a:tr@h` from the source grid.
    const stored = tableStoredRowHeights(table);
    const storedRows = stored && stored.length === rowCount ? stored : rowHeights;
    const sourceRows = asRecord(asRecord(table.source)?.slideSizeOriginalTableGrid)?.rows;
    const heights = storedRows.map((height, index) => {
      if (typeof height !== "number" || !(height > 0)) return 0;
      const emu = Array.isArray(sourceRows) ? sourceRows[index] : undefined;
      return typeof emu === "number" && Math.abs(emu / EMU_PER_PX - height) <= 0.06 ? emu / EMU_PER_PX : height;
    });
    const cells: (EngineTableCell | null)[][] = table.rows.map((row) => row.map(() => null));
    const spanned: { row: number; span: number; need: number }[] = [];
    let pending = false;
    table.rows.forEach((row, rowIndex) => {
      row.forEach((cell, cellIndex) => {
        const position = positions[rowIndex]?.[cellIndex];
        if (!position) return;
        const style = normalizeCellStyle(cell);
        // Turned or vertical words (Text Direction, `a:tcPr@vert`): that cell
        // alone is the browser's, drawn in its writing mode in the row the
        // engine sizes (a row is a minimum, so it still grows to them); the
        
        // eaVert row label was the browser's whole, 74 slides of cells.
        if (style.textDirection !== "horizontal") return;
        const colspan = Math.max(1, Math.floor(Number(cell.colspan) || 1));
        let widthPx = 0;
        for (let offset = 0; offset < colspan && position.col + offset < columnCount; offset += 1) {
          widthPx += Number(colWidths[position.col + offset]) || 0;
        }
        const marginLeft = style.marginLeft ?? TABLE_CELL_NORMAL_MARGIN_PX.horizontal;
        const marginRight = style.marginRight ?? TABLE_CELL_NORMAL_MARGIN_PX.horizontal;
        const marginTop = style.marginTop ?? TABLE_CELL_NORMAL_MARGIN_PX.vertical;
        const marginBottom = style.marginBottom ?? TABLE_CELL_NORMAL_MARGIN_PX.vertical;
        const input = cellInput(table, cell, rowIndex, kernFrom, tableFace);
        const layout = layoutOfficeTextBody(
          { widthPt: Math.max(0, widthPx - marginLeft - marginRight) * PT_PER_PX, wrap: true, keepLongWords: true, cellLines: true, paragraphs: input.paragraphs },
          pageOfficeTextMeasurer,
        );
        if (layout.status === "pending") {
          pending = true;
          return;
        }
        if (layout.status === "unsupported") throw new Unsupported(`cell: ${layout.reason}`);
        cells[rowIndex][cellIndex] = { layout, paragraphs: input.drawn, runCss: input.css, fonts: input.fonts };
        // A row is as tall as its tallest cell's lines, each a whole line
        // (the last one too, not cut at its descent as a text box's height
        
        
        // 8 pt header row 67.2 pt and its margin.
        const linesHeight = cellLinesBottomPt(layout);
        const need = marginTop + linesHeight * PX_PER_PT + marginBottom;
        const span = Math.max(1, Math.floor(Number(cell.rowspan) || 1));
        if (span === 1) heights[position.row] = Math.max(heights[position.row], need);
        else spanned.push({ row: position.row, span, need });
      });
    });
    if (pending) return { status: "pending" };
    // A merged cell taller than the rows it spans: the last of them grows.
    for (const { row, span, need } of spanned) {
      const last = Math.min(rowCount - 1, row + span - 1);
      let total = 0;
      for (let index = row; index <= last; index += 1) total += heights[index];
      if (need > total) heights[last] += need - total;
    }
    return { status: "ok", rowHeightsPx: heights, cells };
  } catch (error) {
    if (error instanceof Unsupported) return tableOff(error.message);
    throw error;
  }
}

/** The table's answer tallied with the bodies' (`window.__mraiSlidesEngine`, the probes' view). */
function tableOff(reason: string): EngineTableAnswer {
  if (typeof window !== "undefined") {
    const tally = ((window as unknown as { __mraiSlidesEngine?: Record<string, number> }).__mraiSlidesEngine ??= {});
    const key = `table off: ${reason.slice(0, 80)}`;
    tally[key] = (tally[key] ?? 0) + 1;
  }
  return { status: "off" };
}

/** The engine's table for this surface: `ok` to place its rows and draw its cells. */
export function usePresentationEngineTable(table: PresentationTableBlock | null): EngineTableAnswer {
  const surface = useContext(PresentationTextEngineSurfaceContext);
  const active =
    surface &&
    table !== null &&
    isImportedPptxBlock(table) &&
    // Read under the text engine's reading row (a page stored before keeps its look).
    asRecord(table.source)?.nativeKerning !== undefined &&
    textEngineOn("slides");
  const [epoch, setEpoch] = useState("");
  useEffect(() => {
    if (!active) return;
    const update = () => setEpoch(textEngineEpoch());
    update();
    return onTextEngineChange(update);
  }, [active]);
  return useMemo(() => {
    void epoch;
    if (surface && table && !active) return tableOff("not a candidate");
    return active && table ? layoutPresentationEngineTable(table) : { status: "off" as const };
  }, [surface, active, table, epoch]);
}

/** One cell's words, drawn where the engine laid them, filling the cell's text box. */
export function PresentationEngineTableCellText({ cell }: { cell: EngineTableCell }): ReactElement {
  const { layout } = cell;
  const nodes: ReactElement[] = [];
  layout.paragraphs.forEach((paragraph, paragraphIndex) => {
    paragraph.lines.forEach((line, lineIndex) => {
      const spans: ReactElement[] = [];
      line.segments.forEach((segment, segmentIndex) => {
        if (segment.tab) return;
        const run = cell.paragraphs[paragraphIndex]?.[segment.run];
        const css = cell.runCss[paragraphIndex]?.[segment.run];
        const font = cell.fonts[paragraphIndex]?.[segment.run];
        if (!run || !css || !font) return;
        const caps = css.textTransform === "uppercase";
        const paint: CSSProperties = {
          fontFamily: font.stack,
          fontSize: Math.round(font.sizePt * PX_PER_PT * 1000) / 1000,
          fontWeight: font.weight,
          fontStyle: font.italic ? "italic" : "normal",
          textDecorationLine: css.textDecorationLine,
          textDecorationStyle: css.textDecorationStyle,
          ...(css.color !== undefined ? { fill: css.color as CSSProperties["fill"] } : {}),
          ...(css.verticalAlign === "super" || css.verticalAlign === "sub" ? ({ baselineShift: css.verticalAlign } as CSSProperties) : {}),
        };
        const glyphs = segment.glyphs;
        if (typeof css.backgroundColor === "string" && css.backgroundColor) {
          spans.push(
            <rect key={`h${segmentIndex}`} x={segment.x * PX_PER_PT} y={line.top * PX_PER_PT} width={segment.width * PX_PER_PT} height={(line.bottom - line.top) * PX_PER_PT} fill={css.backgroundColor} />,
          );
        }
        const shown = (text: string) => (caps ? text.toUpperCase() : text);
        if (glyphs.every((glyph) => shown(glyph.text).length === 1)) {
          spans.push(
            <tspan key={segmentIndex} x={glyphs.map((glyph) => Math.round(glyph.x * PX_PER_PT * 1000) / 1000).join(" ")} style={paint}>
              {glyphs.map((glyph) => shown(glyph.text)).join("")}
            </tspan>,
          );
        } else {
          glyphs.forEach((glyph, glyphIndex) => {
            spans.push(
              <tspan key={`${segmentIndex}-${glyphIndex}`} x={Math.round(glyph.x * PX_PER_PT * 1000) / 1000} style={paint}>
                {shown(glyph.text)}
              </tspan>,
            );
          });
        }
      });
      if (!spans.length) return;
      nodes.push(
        <g key={`${paragraphIndex}-${lineIndex}`}>
          {spans.filter((span) => span.type === "rect")}
          <text y={Math.round(line.baseline * PX_PER_PT * 1000) / 1000} style={{ whiteSpace: "pre", fontKerning: "none", fontVariantLigatures: "none" }}>
            {spans.filter((span) => span.type !== "rect")}
          </text>
        </g>,
      );
    });
  });
  return (
    <div
      data-testid="presentation-engine-table-cell"
      {...{ [PRESENTATION_TEXT_ENGINE_ATTRIBUTE]: "laid" }}
      // As tall as its whole lines: PowerPoint centres a cell's words on them
      // (P2-cell: Arial 10 pt in a 40 pt row 14 pt down, not 14.06 by its descent).
      style={{ position: "relative", width: "100%", height: Math.round(cellLinesBottomPt(layout) * PX_PER_PT * 1000) / 1000 }}
    >
      <svg style={{ position: "absolute", left: 0, top: 0, width: "100%", height: "100%", overflow: "visible" }}>{nodes}</svg>
    </div>
  );
}
