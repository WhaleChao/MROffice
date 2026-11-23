/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a table block is laid out on the slide canvas: cell addressing and the
 * focused span, cell text and <td> styles, the grid (and its <colgroup>), row
 * minimum heights from measured text, direction, and the frame-aligned and
 * gesture-scaled table geometry.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { CSS_PX_PER_POINT, designPxToPoints } from "@mrnote/core/web/lib/presentation/text-units.ts";
import { exactLineSpacingCss, normalizeExactLineSpacingPoints } from "@mrnote/core/web/lib/presentation-line-spacing.ts";
import { officePercentLeadingEm } from "./text-layout-css";
import type { CSSProperties, ReactElement } from "react";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { cellCharacterRuns } from "@mrnote/core/web/lib/presentation-cell-rich-text.ts";
import { EMPTY_PRESENTATION_FONT_ALIASES, presentationFontFamilyWithAssets, presentationFontStacksJoined } from "@mrnote/core/web/lib/presentation-font-face-reads.ts";
import {
  logicalCellPositions,
  logicalPositionOfCell,
  tableLogicalWidth,
} from "@mrnote/core/web/lib/presentation-table-merge.ts";
import {
  DEFAULT_DISTRIBUTE_AXIS_PX,
  MIN_TABLE_AXIS_PX,
  TABLE_FRAME_EMU_PER_CM,
  TABLE_FRAME_EMU_PER_CSS_PX,
  axisSizeAt,
  materializedAxisSizeAt,
  normalizeAxisSizes,
  resizeTableFrame,
} from "@mrnote/core/web/lib/presentation-table-size.ts";
import {
  tableCellAlignmentCss,
  cellStyleCss,
  tableCellTextPaintCss,
  nativeTableStyleCellCss,
  nativeTableThemeColorAuthority,
  normalizeCellStyle,
  normalizeTableStyleLook,
  tableCellFontFamilyCss,
  tableCellTextDecoration,
  tableCellTextDecorationKindCss,
  tableBorderCss,
  tableEdgeCompoundVar,
  tableColEmphasis,
  tableRowEmphasis,
  tableStyleLookCellCss,
  tableStyleLookTableCss,
  TOTAL_ROW_BORDER_TOP,
} from "@mrnote/core/web/lib/presentation-table-style.ts";
import { galleryStyleById } from "@mrnote/core/web/lib/presentation-table-gallery.ts";
import { nativeWebTableStyleById } from "@mrnote/core/web/lib/presentation-table-web-gallery.ts";
import { officeDataVersion } from "@mrnote/core/web/lib/office-data/store.ts";
import type { TableBorderCellBox } from "../presentation-table-border-pen";
import {
  geometryNumber,
  hasLayoutGeometry,
  isImportedPptxBlock,
} from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { frameSize, presentationTableFrameCss } from "@mrnote/core/web/lib/presentation-table-effects.ts";
import type { BlockGeometry } from "@mrnote/core/web/lib/presentation-drag.ts";
import { estimateWrappedLineCount } from "@mrnote/core/web/lib/presentation-autofit.ts";
import type {
  PresentationTableBlock,
  PresentationTableCell,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "@mrnote/core/web/lib/presentation-geometry-limit.ts";
import { clampNumber } from "./selection-geometry";
import { presentationTextShapingCss } from "./presentation-text-shaping";

export const MAX_PRESENTATION_GEOMETRY_EMU = PRESENTATION_GEOMETRY_MAX_EMU;

/**
 * Resolve the inline CSS for a table `<td>`: per-cell fill / align / bold /
 * colour (plus first-row header emphasis when the table-level `headerRow` flag
 * is set, and first/last/banded COLUMN emphasis from the column-axis flags)
 * layered over the table-level thin borders. Shared by the editor canvas and
 * the present overlay so both render paths stay in sync. Returns `undefined`
 * (no `style` attr) when nothing applies, so an un-styled table's cells render
 * byte-for-byte as before.
 */
/** The logical rectangle a focused cell covers, its whole merged span. */
export function focusedTableCellSpan(
  block: PresentationTableBlock,
  focus: { row: number; col: number } | null,
): { row1: number; col1: number; row2: number; col2: number } | null {
  if (!focus) return null;
  const position = logicalPositionOfCell(block.rows, focus.row, focus.col);
  const cell = block.rows[focus.row]?.[focus.col] as
    | { rowspan?: unknown; colspan?: unknown }
    | undefined;
  if (!position || !cell) return null;
  const span = (value: unknown) =>
    typeof value === "number" && Number.isInteger(value) && value > 1 ? value : 1;
  return {
    row1: position.row,
    col1: position.col,
    row2: position.row + span(cell.rowspan) - 1,
    col2: position.col + span(cell.colspan) - 1,
  };
}

/**
 * The array address of the cell at a logical grid position, or of the first
 * cell of that logical row when a merged cell covers the position.
 */
export function tableCellAddressAt(
  rows: readonly (readonly unknown[])[],
  logicalRow: number,
  logicalCol: number,
): { row: number; col: number } | null {
  const positions = logicalCellPositions(rows as PresentationTableCell[][]);
  let rowStart: { row: number; col: number } | null = null;
  for (let row = 0; row < positions.length; row += 1) {
    for (let col = 0; col < positions[row].length; col += 1) {
      const position = positions[row][col];
      if (position?.row !== logicalRow) continue;
      if (position.col === logicalCol) return { row, col };
      if (!rowStart) rowStart = { row, col };
    }
  }
  return rowStart;
}

/** A positioned table cell's own padding (`.presentation-window__table-block td`,
 * `.presentation-present__table td`: 6 px 10 px). */
export const TABLE_CELL_PADDING_PX = { vertical: 6, horizontal: 10 } as const;
/** Layout > Cell Margins > Normal (`tableCellMarginPresetPatch("normal")`). */
export const TABLE_CELL_NORMAL_MARGIN_PX = { vertical: 4.8, horizontal: 9.6 } as const;

/** PowerPoint's single line in a table cell (see `tableCellVisualTextStyle`). */
export const IMPORTED_TABLE_CELL_LINE_HEIGHT = 1.2;

/** A table read from a deck (its source names the slide it came from). */
function importedPresentationTable(block: PresentationTableBlock): boolean {
  const source = (block as { source?: unknown }).source;
  return Boolean(
    source && typeof source === "object" && typeof (source as { slideId?: unknown }).slideId === "string",
  );
}

/**
 * Whether an imported table is drawn by a style of its own (the package's
 * resolved `tableStyleLook`, or a native style the app knows by its id) whose
 * First Row is not bold: its header is then drawn regular, as the style says.
 */
function importedTableDrawnByItsStyle(block: PresentationTableBlock): boolean {
  if (!importedPresentationTable(block)) return false;
  const columns = Math.max(1, ...block.rows.map((row) => row.length));
  const look = normalizeTableStyleLook(block.tableStyleLook, block.tableStyleId);
  const header = look
    ? tableStyleLookCellCss(look, block, 0, block.rows.length, 0, columns)
    : nativeTableStyleCellCss(block.tableStyleId, block, 0, block.rows.length, 0, columns);
  if (!header) return false;
  const weight = (header as { fontWeight?: unknown }).fontWeight;
  return weight !== 700 && weight !== "bold" && weight !== "700";
}

/** PowerPoint's text size in a table cell that states none: 18 pt (24 design px). */
export const PRESENTATION_TABLE_CELL_DEFAULT_FONT_PX = 24;









export function tableCellStrutFontSizePt(cell: PresentationTableCell | undefined): number | null {
  return tableCellStrut(cell)?.sizePt ?? null;
}









export function tableCellStrut(
  cell: PresentationTableCell | undefined,
): { sizePt: number | null; fontFamily: string | null } | null {
  if (!cell) return null;
  const runs = cellCharacterRuns(String(cell.text ?? ""), (cell as { runs?: unknown }).runs).filter((run) =>
    /\S/.test(run.text),
  );
  const sizes = runs.map((run) => run.style?.fontSize);
  const families = new Set(runs.map((run) => run.style?.fontFamily));
  const [family] = families;
  const fontFamily = runs.length && families.size === 1 && typeof family === "string" && family.trim() ? family : null;
  const sized = sizes.length > 0 && sizes.every((size) => typeof size === "number" && Number.isFinite(size) && size > 0);
  if (!sized && fontFamily === null) return null;
  return {
    sizePt: sized ? designPxToPoints(Math.min(...(sizes as number[]))) : null,
    fontFamily,
  };
}

/** A cell's own proportional line spacing (`lineSpacing`), else 1. */
/** Line Spacing > Exactly every paragraph of the cell states, in points
 * (`lineSpacingPoints`, readings row 832), else undefined. */
export function tableCellExactLineSpacingPt(cell: PresentationTableCell | undefined): number | undefined {
  return normalizeExactLineSpacingPoints((cell as { lineSpacingPoints?: unknown } | undefined)?.lineSpacingPoints) || undefined;
}

function tableCellLineSpacing(cell: PresentationTableCell | undefined): number {
  const value = (cell as { lineSpacing?: unknown } | undefined)?.lineSpacing;
  return typeof value === "number" && Number.isFinite(value) && value >= 0.25 && value <= 5 ? value : 1;
}

/**
 * The height a cell's row -- or the rows it spans -- state, in px, or
 * undefined when any of them states none. Turned words wrap at it
 * (`tableCellVisualTextStyle`).
 */
export function tableCellStatedHeightPx(
  block: PresentationTableBlock,
  rowIndex: number,
  cell: { rowspan?: unknown } | undefined,
): number | undefined {
  const count = block.rows.length;
  const heights = normalizeAxisSizes(block.rowHeights, count);
  if (!heights || rowIndex < 0 || rowIndex >= count) return undefined;
  const span = Math.max(1, Math.floor(Number(cell?.rowspan) || 1));
  let total = 0;
  for (let row = rowIndex; row < Math.min(count, rowIndex + span); row += 1) {
    const height = axisSizeAt(heights, row);
    if (height === null || !(height > 0)) return undefined;
    total += height;
  }
  return total;
}

/**
 * The faces a cell's complex-script and East Asian characters are drawn in
 * (`scriptFaces`, readings row 833: the run's own `a:cs` / `a:ea`, else the
 * theme's body font for the slot, else its `a:font` of the run's script).
 * Drawn only.
 */
export function tableCellScriptFaces(cell: unknown): { cs?: string; ea?: string } | undefined {
  const raw = (cell as { scriptFaces?: unknown } | undefined)?.scriptFaces;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const out: { cs?: string; ea?: string } = {};
  for (const slot of ["cs", "ea"] as const) {
    const face = (raw as Record<string, unknown>)[slot];
    if (typeof face === "string" && face.trim() && face.length <= 200) out[slot] = face.trim();
  }
  return out.cs || out.ea ? out : undefined;
}









export function servedTableCellFontFamily(value: unknown, scriptFaces?: { cs?: string; ea?: string }): string | undefined {
  const css = tableCellFontFamilyCss(value);
  let served = css ? (presentationFontFamilyWithAssets(css, EMPTY_PRESENTATION_FONT_ALIASES) ?? css) : css;
  if (!served) return served;
  for (const face of [scriptFaces?.cs, scriptFaces?.ea]) {
    const own = face ? tableCellFontFamilyCss(face) : undefined;
    if (own) served = presentationFontStacksJoined(served, presentationFontFamilyWithAssets(own, EMPTY_PRESENTATION_FONT_ALIASES) ?? own);
  }
  return served;
}

/** The face PowerPoint sets the words of a cell naming none in (readings row 1033, `source.nativeTableFace`), or null. */
export function tableInheritedFace(block: PresentationTableBlock): string | null {
  const face = (block.source as { nativeTableFace?: unknown } | null | undefined)?.nativeTableFace;
  return typeof face === "string" && face.trim() && face.length <= 200 ? face.trim() : null;
}

export function tableCellVisualTextStyle(
  block: PresentationTableBlock,
  cellStyle: ReturnType<typeof normalizeCellStyle>,
  headerEmphasis: boolean,
  /** The cell's editing field (a textarea) rather than its drawn words. */
  field = false,
  /** The cell itself: a cell stating no size stands on its words' smallest. */
  cell?: PresentationTableCell,
  /** The height its row(s) state, in px (`tableCellStatedHeightPx`). */
  statedHeightPx?: number,
): CSSProperties {
  const strut = tableCellStrut(cell);
  const strutFontSizePt = cellStyle.fontSize === null ? (strut?.sizePt ?? null) : null;
  const strutFontFamily = cellStyle.fontFamily === null ? (strut?.fontFamily ?? null) : null;
  // A positioned cell's text box reaches over the `<td>`'s default 6 px / 8 px
  // padding to fill the cell. A side whose margin the cell STATES (Cell
  // Margins > None, Narrow, Wide or Custom) is padded by exactly that margin,
  // so the box must not reach over it too: with None it pulled the text 8 px
  // into the next cell (PPT-B102).
  // A side whose margin the cell does NOT state takes PowerPoint's Normal
  // margins (0.1 in left and right, 0.05 in top and bottom: 9.6 / 4.8 design
  // px), the preset Cell Margins shows checked: the box reaches over only the
  // rest of the cell's 6 px / 10 px padding. It reached over nearly all of
  // it, so a new table's words sat 2 px from the edge while Normal was
  
  const reach = (stated: number | null, padding: number, normal: number) =>
    stated === null ? Math.round((padding - normal) * 100) / 100 : 0;
  const reachTop = reach(cellStyle.marginTop, TABLE_CELL_PADDING_PX.vertical, TABLE_CELL_NORMAL_MARGIN_PX.vertical);
  const reachBottom = reach(cellStyle.marginBottom, TABLE_CELL_PADDING_PX.vertical, TABLE_CELL_NORMAL_MARGIN_PX.vertical);
  const reachLeft = reach(cellStyle.marginLeft, TABLE_CELL_PADDING_PX.horizontal, TABLE_CELL_NORMAL_MARGIN_PX.horizontal);
  const reachRight = reach(cellStyle.marginRight, TABLE_CELL_PADDING_PX.horizontal, TABLE_CELL_NORMAL_MARGIN_PX.horizontal);
  // A cell's PERCENTAGE line spacing lays its lines out as a text box's
  // (`officePercentLeadingEm`): the extra above each line, the first
  // baseline at the larger of single's and three quarters of the line, and
  // nothing under the last line's descent. Pitched at 1.2 x the spacing in a
  // CSS line box, a 115 % cell stood a sixth of a line too tall and its words
  
  // PowerPoint 12.96 pt apart, the editor 13.8; PowerPoint's print of
  // truth-probes/lookfit-cell-spacing-probe.pptx, 80-200 %, one and three
  // lines, Arial and Calibri). The glyphs move down by `shift`; the box gives
  // back `tail` under its last line (a negative bottom margin, so a row its
  // words decide is that much lower), and an anchored cell's words move down
  // by what the anchor shared of it.
  const exactSpacing = importedPresentationTable(block) ? tableCellExactLineSpacingPt(cell) : undefined;
  const percentLeading =
    importedPresentationTable(block) && exactSpacing === undefined && tableCellLineSpacing(cell) !== 1
      ? officePercentLeadingEm(
          Math.round(IMPORTED_TABLE_CELL_LINE_HEIGHT * tableCellLineSpacing(cell) * 10_000) / 10_000,
          cellStyle.fontFamily ?? strutFontFamily ?? tableInheritedFace(block) ?? undefined,
        )
      : { shift: 0, tail: 0 };
  const leadingDrop =
    percentLeading.shift +
    (cellStyle.verticalAlign === "bottom" ? percentLeading.tail : cellStyle.verticalAlign === "middle" ? percentLeading.tail / 2 : 0);
  const tailCss = percentLeading.tail ? ` - ${percentLeading.tail}em` : "";
  return {
    border: 0,
    borderRadius: 0,
    boxSizing: "border-box",
    background: "transparent",
    display: "block",
    minHeight: 34,
    // A cell stating no size is PowerPoint's new-table text, 18 pt (24
    // design px), in the editor, the show, the PDF and the file alike; it
    // was 17 px here and PowerPoint drew the exported table a third larger.
    fontSize: PRESENTATION_TABLE_CELL_DEFAULT_FONT_PX,
    overflow: hasLayoutGeometry(block) ? "hidden" : "auto",
    padding: hasLayoutGeometry(block) ? 0 : "6px 9px",
    // An imported table's words sit on PowerPoint's single line, 1.2 of their
    // size (its PDF: a 15 pt Calibri row is 18.0 pt tall with no margins);
    // they inherited the editor's 1.5 and every row of a dense table stood a
    // quarter taller than PowerPoint's, pushing the table off the slide
    
    // A cell whose paragraphs state their own proportional spacing
    // (`lineSpacing`, `a:lnSpc/a:spcPct`) pitches its lines at that multiple
    
    lineHeight: importedPresentationTable(block)
      ? Math.round(IMPORTED_TABLE_CELL_LINE_HEIGHT * tableCellLineSpacing(cell) * 10_000) / 10_000
      : "inherit",
    resize: "none",
    whiteSpace: "pre-wrap",
    // A word wider than its cell breaks at the cell's edge, as PowerPoint breaks
    
    // rest of it out of sight), as the text engine's lines already do.
    overflowWrap: "anywhere",
    width: "100%",
    ...(hasLayoutGeometry(block)
      ? {
          // An imported cell's text box also takes back the halves of its
          // lines its `<td>` added (`tableCellTdStyle`): top and bottom ...
          ...(importedPresentationTable(block)
            ? {
                height: `calc(100% + ${reachTop + reachBottom}px + var(${CELL_BORDER_TOP_HALF_VAR}, 0px) + var(${CELL_BORDER_BOTTOM_HALF_VAR}, 0px))`,
                margin: `calc(${-reachTop}px - var(${CELL_BORDER_TOP_HALF_VAR}, 0px)) calc(${-reachRight}px - var(${CELL_BORDER_RIGHT_HALF_VAR}, 0px)) calc(${-reachBottom}px - var(${CELL_BORDER_BOTTOM_HALF_VAR}, 0px)${tailCss}) calc(${-reachLeft}px - var(${CELL_BORDER_LEFT_HALF_VAR}, 0px))`,
                ...(leadingDrop ? { top: `${Math.round(leadingDrop * 10_000) / 10_000}em` } : {}),
                maxHeight: `calc(100% + ${reachTop + reachBottom}px + var(${CELL_BORDER_TOP_HALF_VAR}, 0px) + var(${CELL_BORDER_BOTTOM_HALF_VAR}, 0px))`,
              }
            : {
                height: `calc(100% + ${reachTop + reachBottom}px)`,
                margin: `${-reachTop}px ${-reachRight}px ${-reachBottom}px ${-reachLeft}px`,
                maxHeight: `calc(100% + ${reachTop + reachBottom}px)`,
              }),
          minHeight: 0,
          position: "relative",
          // ... and of its left and right lines (`tableCellTdStyle`).
          width: importedPresentationTable(block)
            ? `calc(100% + ${reachLeft + reachRight}px + var(${CELL_BORDER_LEFT_HALF_VAR}, 0px) + var(${CELL_BORDER_RIGHT_HALF_VAR}, 0px))`
            : `calc(100% + ${reachLeft + reachRight}px)`,
          // The text box fills its cell, so the `<td>`'s vertical-align has
          // nothing to move: the cell's Top / Center / Bottom anchor aligns
          // the runs inside the box itself (block `align-content`), in the
          // editor and the slide show alike.
          ...(cellStyle.verticalAlign !== "top"
            ? {
                alignContent:
                  cellStyle.verticalAlign === "middle" ? "center" : "end",
              }
            : {}),
        }
      : {}),
    // Turned words (Text Direction other than Horizontal) wrap at the height
    // their row states, as PowerPoint wraps them, and a word longer than it
    // still makes the row taller. Filling the cell, the box took the length
    
    // header of 'Rotate all text 270' cells stood twice its 2.5 in).
    ...(hasLayoutGeometry(block) &&
    importedPresentationTable(block) &&
    cellStyle.textDirection !== "horizontal" &&
    typeof statedHeightPx === "number" &&
    statedHeightPx > 0
      ? {
          height: "auto",
          maxHeight: Math.max(
            1,
            Math.round(
              (statedHeightPx -
                (cellStyle.marginTop ?? TABLE_CELL_NORMAL_MARGIN_PX.vertical) -
                (cellStyle.marginBottom ?? TABLE_CELL_NORMAL_MARGIN_PX.vertical)) *
                100,
            ) / 100,
          ),
          minHeight: "min-content",
        }
      : {}),
    ...(cellStyle.align !== "left" ? tableCellAlignmentCss(cellStyle.align) : {}),
    // Right-to-Left paragraphs (`dir`, readings row 831): the lines' base
    
    // label of "A1:2017 - הזרקת קוד" on the right, as PowerPoint sets it).
    // The alignment stays the cell's own, a physical side.
    ...((cell as { dir?: unknown } | undefined)?.dir === "rtl"
      ? { direction: "rtl" as const, ...(cellStyle.align === "left" ? { textAlign: "left" as const } : {}) }
      : {}),
    // A header cell whose runs all say bold off stays regular (CDC p60). An
    // imported table drawn by its own style takes the header's weight from
    
    // package style states no bold there, and PowerPoint draws the header
    // regular.
    ...(cellStyle.bold ||
    (headerEmphasis && !(cell && nativeCellRunsStateNotBold(cell)) && !importedTableDrawnByItsStyle(block))
      ? { fontWeight: 700 }
      : {}),
    ...(cellStyle.italic ? { fontStyle: "italic" } : {}),
    ...(tableCellTextDecoration(cellStyle)
      ? {
          textDecoration: tableCellTextDecoration(cellStyle),
          ...tableCellTextDecorationKindCss(cellStyle),
        }
      : {}),
    ...(cellStyle.color !== null ? { color: cellStyle.color } : {}),
    ...(cellStyle.highlight !== null
      ? {
          backgroundColor: cellStyle.highlight,
        }
      : {}),
    // Served as a run's face is (`presentationFontFamilyWithAssets`): the
    // Office name's static alias rule carries only its stand-in's Latin
    // slices, so Cyrillic, Greek and Vietnamese words need the stand-in's own
    
    ...(cellStyle.fontFamily !== null
      ? {
          fontFamily: servedTableCellFontFamily(cellStyle.fontFamily, tableCellScriptFaces(cell)),
        }
      : strutFontFamily !== null
        ? { fontFamily: servedTableCellFontFamily(strutFontFamily, tableCellScriptFaces(cell)) }
        : tableInheritedFace(block) !== null
          // A cell naming no face is set in the one PowerPoint sets its
          // table's words in (readings row 1033, `nativeTableFace`): on the
          // canvas, in the text engine and in print alike.
          ? { fontFamily: servedTableCellFontFamily(tableInheritedFace(block), tableCellScriptFaces(cell)) }
          : {}),
    ...(cellStyle.fontSize !== null
      ? {
          fontSize: `${cellStyle.fontSize}pt`,
        }
      : strutFontSizePt !== null
        ? { fontSize: `${strutFontSizePt}pt` }
        : {}),
    // Line Spacing > Exactly on every paragraph (readings row 832): each line
    // that tall, its glyphs where PowerPoint sets them (`exactLineSpacingCss`,
    
    ...(importedPresentationTable(block) && tableCellExactLineSpacingPt(cell) !== undefined
      ? exactLineSpacingCss(
          tableCellExactLineSpacingPt(cell)!,
          cellStyle.fontSize ?? strutFontSizePt ?? undefined,
        )
      : {}),
    ...(cellStyle.superscript || cellStyle.subscript
      ? {
          position: "relative",
          top: cellStyle.superscript ? "-0.25em" : "0.25em",
          ...(cellStyle.fontSize === null ? { fontSize: "0.7em" } : {}),
        }
      : {}),
    // Table Design > WordArt Styles (PPT-A138): the paint owns the glyphs.
    ...tableCellTextPaintCss(cellStyle.textPaint, { field }),
  };
}

/**
 * The strips a cell draws its table's lines with: an imported table's runs of
 * line it owns (`TABLE_CELL_STROKES_VAR`, `tableEdgeStrokePlan`), layered in
 * PowerPoint's stroke order; else the rules of its compound edges (Thick
 * Thin, Thin Thick, Triple: the `<td>` style's `tableEdgeCss` custom
 * properties), drawn over the transparent width its border keeps: each a
 * strip centred on the edge's line and reaching its corners. Null for a cell
 * with none.
 */
export function TableCellCompoundEdges({ tdStyle }: { tdStyle: CSSProperties | undefined }): ReactElement | null {
  if (!tdStyle) return null;
  const record = tdStyle as Record<string, unknown>;
  const strips: ReactElement[] = [];
  const planned = record[TABLE_CELL_STROKES_VAR];
  if (typeof planned === "string" && planned) {
    planned.split(";").forEach((strip, index) => {
      const [z, left, top, right, bottom, width, height, ...paint] = strip.split("|");
      const length = (value: string | undefined) => (value === undefined || value === "auto" ? undefined : value);
      strips.push(
        <span
          key={`stroke-${index}`}
          aria-hidden="true"
          data-testid="presentation-table-stroke"
          style={{
            position: "absolute",
            pointerEvents: "none",
            zIndex: Number(z) || undefined,
            left: length(left),
            top: length(top),
            right: length(right),
            bottom: length(bottom),
            width: length(width),
            height: length(height),
            background: paint.join("|"),
          }}
        />,
      );
    });
    return <>{strips}</>;
  }
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const value = record[tableEdgeCompoundVar(side)];
    if (typeof value !== "string") continue;
    const border = record[`border${side[0].toUpperCase()}${side.slice(1)}`];
    if (typeof border !== "string" || !border.endsWith("solid transparent")) continue;
    const at = value.indexOf("|");
    const width = Number.parseFloat(value.slice(0, at));
    const image = value.slice(at + 1);
    if (!(width > 0) || !image) continue;
    // How far it runs past the cell's box at each end (`compoundEdgeReaches`),
    // else to the middle of the line across it.
    const ends = record[`${tableEdgeCompoundVar(side)}-ends`];
    const [from, to] = typeof ends === "string" ? ends.split("|").map(Number) : [Number.NaN, Number.NaN];
    const reach = (edge: "top" | "bottom" | "left" | "right") => {
      const stated = edge === "top" || edge === "left" ? from : to;
      return Number.isFinite(stated)
        ? `${-stated}px`
        : `calc(-1 * var(--presentation-cell-border-${edge}-half, 0px))`;
    };
    strips.push(
      <span
        key={side}
        aria-hidden="true"
        data-testid={`presentation-table-compound-${side}`}
        style={{
          position: "absolute",
          pointerEvents: "none",
          backgroundImage: image,
          ...(side === "top"
            ? { left: reach("left"), right: reach("right"), height: width, top: -width }
            : side === "bottom"
              ? { left: reach("left"), right: reach("right"), height: width, bottom: -width }
              : side === "left"
                ? { top: reach("top"), bottom: reach("bottom"), width, left: -width }
                : { top: reach("top"), bottom: reach("bottom"), width, right: -width }),
        }}
      />,
    );
  }
  return strips.length ? <>{strips}</> : null;
}

/**
 * A cell's start edge (`a:lnL`, or its style's left line) taken off its
 * `<td>`: `none`, which loses every collapsed-border conflict, so the line
 * the column before states on its end side (`a:lnR`) draws the edge, or
 * nothing does, as PowerPoint draws a table (the first column's start edge
 * is its own). The edge's compound rules go with it.
 */
export function dropTableCellStartEdge(css: CSSProperties): void {
  dropTableCellEdge(css, "left");
}

/** One side's line taken off a cell's `<td>` (`dropTableCellStartEdge`). */
function dropTableCellEdge(css: CSSProperties, side: "top" | "left"): void {
  const record = css as Record<string, unknown>;
  const name = side === "top" ? "borderTop" : "borderLeft";
  delete record[`${name}Width`];
  delete record[`${name}Style`];
  delete record[`${name}Color`];
  delete record[tableEdgeCompoundVar(side)];
  record[name] = "none";
}

/**
 * The cells across a cell's top (the row above it) and across its start side
 * (the logical column before it), on the table's grid; and across its bottom
 * (the row under its last row), which share its bottom edge.
 */
function cellsBeforeEdges(
  block: PresentationTableBlock,
  facts: TableTdFacts,
  rowIndex: number,
  colIndex: number,
): { above: PresentationTableCell[]; before: PresentationTableCell[]; below: PresentationTableCell[] } {
  const cell = block.rows[rowIndex]?.[colIndex];
  const col = facts.logicalCol(rowIndex, colIndex);
  const colspan = Math.max(1, cell?.colspan ?? 1);
  const rowspan = Math.max(1, cell?.rowspan ?? 1);
  const collect = (cells: (PresentationTableCell | undefined)[]) =>
    cells.filter((value, index, all): value is PresentationTableCell => value !== undefined && all.indexOf(value) === index);
  return {
    above: rowIndex > 0 ? collect(Array.from({ length: colspan }, (_v, at) => facts.cellAt(rowIndex - 1, col + at))) : [],
    before: col > 0 ? collect(Array.from({ length: rowspan }, (_v, at) => facts.cellAt(rowIndex + at, col - 1))) : [],
    below: collect(Array.from({ length: colspan }, (_v, at) => facts.cellAt(rowIndex + rowspan, col + at))),
  };
}

/** Half of an imported cell's top / bottom line, set on its `<td>`. */
const CELL_BORDER_TOP_HALF_VAR = "--presentation-cell-border-top-half";
const CELL_BORDER_BOTTOM_HALF_VAR = "--presentation-cell-border-bottom-half";
/** Half of an imported cell's left / right line, set on its `<td>`. */
const CELL_BORDER_LEFT_HALF_VAR = "--presentation-cell-border-left-half";
const CELL_BORDER_RIGHT_HALF_VAR = "--presentation-cell-border-right-half";







function drawnBorderWidthPx(width: number): number {
  if (!(width > 0)) return 0;
  const ratio = typeof window !== "undefined" && window.devicePixelRatio > 0 ? window.devicePixelRatio : 1;
  return Math.max(1, Math.floor(width * ratio + 1e-6)) / ratio;
}

/** The width in px a CSS border (shorthand or width) states, or 0. */
function cssBorderWidthPx(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) && value > 0 ? value : 0;
  if (typeof value !== "string" || /\b(none|hidden)\b/.test(value)) return 0;
  const width = /(?:^|\s)(\d*\.?\d+)px(?:\s|$)/.exec(value);
  if (width) return Number(width[1]);
  if (/\bthin\b/.test(value)) return 1;
  if (/\bmedium\b/.test(value)) return 3;
  if (/\bthick\b/.test(value)) return 5;
  // A shorthand with a style and no width draws the medium default.
  return /\b(solid|dashed|dotted|double|groove|ridge|inset|outset)\b/.test(value) ? 3 : 0;
}

/** The width in px of one side's line of a cell's `<td>` style. */
function borderOf(css: CSSProperties, side: "Top" | "Bottom" | "Left" | "Right"): number {
  const record = css as Record<string, unknown>;
  return cssBorderWidthPx(record[`border${side}Width`] ?? record[`border${side}`] ?? css.borderWidth ?? css.border);
}













function cellEdgeNeighbours(
  block: PresentationTableBlock,
  rowIndex: number,
  colIndex: number,
): { above: CSSProperties[]; below: CSSProperties[]; left: CSSProperties[]; right: CSSProperties[] } {
  const facts = tableTdEntry(block).facts;
  const styleOf = (neighbour: PresentationTableCell | undefined, r: number, c: number): CSSProperties[] => {
    const css = neighbour ? tableCellTdStyle(block, neighbour, r, c, false) : undefined;
    return css ? [css] : [];
  };
  if (facts.hasMerge) {
    const across = cellsBeforeEdges(block, facts, rowIndex, colIndex);
    const styles = (cells: readonly PresentationTableCell[]) =>
      cells.flatMap((neighbour) => {
        const place = facts.placeOf(neighbour);
        return place ? styleOf(neighbour, place.row, place.col) : [];
      });
    return { above: styles(across.above), below: styles(across.below), left: [], right: [] };
  }
  const at = (r: number, c: number) => styleOf(block.rows[r]?.[c], r, c);
  const rtl = block.rightToLeft === true;
  return {
    above: at(rowIndex - 1, colIndex),
    below: at(rowIndex + 1, colIndex),
    left: at(rowIndex, rtl ? colIndex + 1 : colIndex - 1),
    right: at(rowIndex, rtl ? colIndex - 1 : colIndex + 1),
  };
}

/** Whether an imported cell's words all say bold off: its native runs
 * (`nativeText.runs`), each with text, each stating `bold: false`. */
function nativeCellRunsStateNotBold(cell: PresentationTableCell): boolean {
  const native = (cell as { nativeText?: unknown }).nativeText;
  const runs =
    native && typeof native === "object" && Array.isArray((native as { runs?: unknown }).runs)
      ? ((native as { runs: unknown[] }).runs as { text?: unknown; style?: { bold?: unknown } }[])
      : [];
  const worded = runs.filter((run) => typeof run?.text === "string" && run.text.trim() !== "");
  return worded.length > 0 && worded.every((run) => run.style?.bold === false);
}











type TableTdFacts = {
  hasMerge: boolean;
  colCount: number;
  logicalCol: (rowIndex: number, colIndex: number) => number;
  /** The cell covering a place on the table's grid (row, logical column). */
  cellAt: (rowIndex: number, logicalCol: number) => PresentationTableCell | undefined;
  /** A cell's place in the table's rows (row, index in its row). */
  placeOf: (cell: PresentationTableCell) => { row: number; col: number } | undefined;
  sourceThemeColors: ReturnType<typeof nativeTableThemeColorAuthority>;
  blockStyleLook: ReturnType<typeof normalizeTableStyleLook>;
};
type TableTdEntry = {
  stamp: number;
  facts: TableTdFacts;
  frame: readonly unknown[];
  styles: Map<string, CSSProperties | undefined>;
};
const tableTdEntries = new WeakMap<PresentationTableBlock, TableTdEntry>();

/**
 * A cell's style carried from the table it was last drawn in: an edit
 * replaces the table but keeps every cell it did not touch, so a keystroke
 * re-derives the cell it typed in (and its neighbours' lines) and takes every
 * other cell's style as it was. Kept while the table's frame (what every
 * cell's style reads of the table: its style, flags, size, merge grid,
 * source) is the same, the cell is at the same place, and -- for the halves
 * of its lines -- its four neighbours are the same cells.
 */
type CellTdMemo = {
  frame: readonly unknown[];
  rowIndex: number;
  colIndex: number;
  logicalCol: number;
  neighbours: readonly unknown[] | null;
  /** The cells across its top and start side, whose own lines decide those edges (`cellsBeforeEdges`). */
  deciders: readonly unknown[];
  /** The strips it draws of its table's lines (`tableEdgeStrokePlan`): they read the whole table's lines. */
  strokes: unknown;
  css: CSSProperties | undefined;
};
const cellTdMemos = new WeakMap<object, [CellTdMemo | undefined, CellTdMemo | undefined]>();

function sameFrame(a: readonly unknown[], b: readonly unknown[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

function cellNeighbourIdentities(block: PresentationTableBlock, rowIndex: number, colIndex: number): readonly unknown[] {
  return [
    block.rows[rowIndex - 1]?.[colIndex],
    block.rows[rowIndex + 1]?.[colIndex],
    block.rows[rowIndex]?.[colIndex - 1],
    block.rows[rowIndex]?.[colIndex + 1],
    // Its strips' ends read the cells across its corners too (`tableEdgeStrokePlan`).
    block.rows[rowIndex - 1]?.[colIndex - 1],
    block.rows[rowIndex - 1]?.[colIndex + 1],
    block.rows[rowIndex + 1]?.[colIndex - 1],
    block.rows[rowIndex + 1]?.[colIndex + 1],
  ];
}

/**
 * The facts of a table its cells' `<td>` styles read besides each cell and
 * its neighbours (style, look, flags, direction, borders, row count, logical
 * width, merges, the served style data): equal frames, element by element,
 * draw the same cells the same.
 */
export function tableCellRenderFrame(block: PresentationTableBlock): readonly unknown[] {
  return tableTdEntry(block).frame;
}

function tableTdEntry(block: PresentationTableBlock): TableTdEntry {
  const stamp = officeDataVersion();
  const known = tableTdEntries.get(block);
  if (known && known.stamp === stamp) return known;
  const hasMerge = block.rows.some((r) =>
    r.some((c) => (c.colspan ?? 1) > 1 || (c.rowspan ?? 1) > 1),
  );
  const positions = hasMerge ? logicalCellPositions(block.rows) : null;
  let covering: Map<string, PresentationTableCell> | null = null;
  if (positions) {
    covering = new Map();
    block.rows.forEach((row, r) =>
      row.forEach((cell, c) => {
        const at = positions[r]?.[c];
        if (!at) return;
        for (let dr = 0; dr < Math.max(1, cell.rowspan ?? 1); dr += 1)
          for (let dc = 0; dc < Math.max(1, cell.colspan ?? 1); dc += 1) covering!.set(`${at.row + dr}:${at.col + dc}`, cell);
      }),
    );
  }
  const sourceThemeColors = nativeTableThemeColorAuthority(block.source);
  let places: Map<PresentationTableCell, { row: number; col: number }> | null = null;
  const facts: TableTdFacts = {
    hasMerge,
    // Column count mirrors the `<colgroup>` derivation (widest row → logical
    // width when merged).
    colCount: hasMerge
      ? tableLogicalWidth(block.rows)
      : block.rows.reduce((widest, row) => Math.max(widest, row.length), 0),
    logicalCol: (rowIndex, colIndex) => positions?.[rowIndex]?.[colIndex]?.col ?? colIndex,
    cellAt: (rowIndex, logicalCol) =>
      covering ? covering.get(`${rowIndex}:${logicalCol}`) : block.rows[rowIndex]?.[logicalCol],
    placeOf: (cell) => {
      if (!places) {
        places = new Map();
        block.rows.forEach((row, r) => row.forEach((each, c) => places!.set(each, { row: r, col: c })));
      }
      return places.get(cell);
    },
    sourceThemeColors,
    blockStyleLook:
      normalizeTableStyleLook(block.tableStyleLook, block.tableStyleId) ??
      (sourceThemeColors
        ? null
        : normalizeTableStyleLook(
            nativeWebTableStyleById(block.tableStyleId)?.look ??
              galleryStyleById(block.tableStyleId)?.look ??
              null,
            block.tableStyleId,
          )),
  };
  const flags = block as {
    headerRow?: unknown; totalRow?: unknown; bandedRows?: unknown;
    firstColumn?: unknown; lastColumn?: unknown; bandedColumns?: unknown;
  };
  const frame = [
    stamp,
    block.tableStyleLook,
    block.tableStyleId,
    block.source,
    flags.headerRow,
    flags.totalRow,
    flags.bandedRows,
    flags.firstColumn,
    flags.lastColumn,
    flags.bandedColumns,
    block.rightToLeft,
    block.borders,
    block.rows.length,
    facts.colCount,
    hasMerge,
  ] as const;
  const entry = { stamp, facts, frame, styles: new Map<string, CSSProperties | undefined>() };
  tableTdEntries.set(block, entry);
  return entry;
}

export function tableCellTdStyle(
  block: PresentationTableBlock,
  cell: PresentationTableCell,
  rowIndex: number,
  colIndex: number,
  /** Internal: the neighbours' own lines are read without their halves. */
  withBorderHalves = true,
): CSSProperties | undefined {
  // The display model may contain a gallery preview; don't repaint its cells
  // with the original block's direct shading supplied by the editing loop.
  const own = block.rows[rowIndex]?.[colIndex];
  if (!own) return tableCellTdStyleOf(block, cell, rowIndex, colIndex, withBorderHalves, tableTdEntry(block).facts);
  const entry = tableTdEntry(block);
  const key = `${rowIndex}:${colIndex}:${withBorderHalves ? 1 : 0}`;
  if (entry.styles.has(key)) return entry.styles.get(key);
  const slot = withBorderHalves ? 1 : 0;
  const logicalCol = entry.facts.logicalCol(rowIndex, colIndex);
  // The halves of a cell's lines read its neighbours (`cellEdgeNeighbours`):
  // in a table with merged cells, the cells across its top and bottom on the grid.
  const across = cellsBeforeEdges(block, entry.facts, rowIndex, colIndex);
  const neighbours = !withBorderHalves
    ? null
    : entry.facts.hasMerge
      ? [...across.above, null, ...across.below]
      : cellNeighbourIdentities(block, rowIndex, colIndex);
  const deciders = [...across.above, null, ...across.before];
  const strokes = withBorderHalves ? tableEdgeStrokeIdentity(block, rowIndex, colIndex) : null;
  const memo = cellTdMemos.get(own)?.[slot];
  if (
    memo &&
    memo.rowIndex === rowIndex &&
    memo.colIndex === colIndex &&
    memo.logicalCol === logicalCol &&
    sameFrame(memo.frame, entry.frame) &&
    (memo.neighbours === null) === (neighbours === null) &&
    (neighbours === null || sameFrame(memo.neighbours!, neighbours)) &&
    sameFrame(memo.deciders, deciders) &&
    memo.strokes === strokes
  ) {
    entry.styles.set(key, memo.css);
    return memo.css;
  }
  const css = tableCellTdStyleOf(block, own, rowIndex, colIndex, withBorderHalves, entry.facts);
  entry.styles.set(key, css);
  const memos = cellTdMemos.get(own) ?? [undefined, undefined];
  memos[slot] = { frame: entry.frame, rowIndex, colIndex, logicalCol, neighbours, deciders, strokes, css };
  cellTdMemos.set(own, memos);
  return css;
}

function tableCellTdStyleOf(
  block: PresentationTableBlock,
  cell: PresentationTableCell,
  rowIndex: number,
  colIndex: number,
  withBorderHalves: boolean,
  facts: TableTdFacts,
): CSSProperties | undefined {
  const { bandFill, totalEmphasis } = tableRowEmphasis(
    block,
    rowIndex,
    block.rows.length,
  );
  // Column banding / first-last emphasis must address the LOGICAL column: a
  // merged cell (colspan), or a rowspan reaching down from a row above, shifts a
  // cell's logical position away from its array index, so striping by the array
  // index alone would emphasise the wrong physical columns. An un-merged
  // table's array index already IS its logical column.
  const logicalColIndex = facts.logicalCol(rowIndex, colIndex);
  const colCount = facts.colCount;
  const { colEmphasis } = tableColEmphasis(block, logicalColIndex, colCount);
  // An imported deck carries the RESOLVED definition of its native style
  // (`tableStyleLook`, read from the package's own `tableStyles.xml`), which
  // beats the local gallery preview: it is the look the author actually saved.
  // Falling back to the preset keeps a deck authored here rendering as before.
  //
  // A style this app AUTHORS needs no definition in the package: the gallery
  // holds it. That is what recovers a deck whose `tableStyles.xml` was rewritten
  // by a tool that dropped unknown entries — the GUID survives the round trip and
  // still renders the look it was built with, instead of a plain grid.
  //
  // The block's OWN look only wins once it has been validated. A stale or
  // hand-edited `tableStyleLook` that the validator refuses must degrade to the
  // same native definition a block with NO look resolves to — otherwise the one
  // deck whose look failed validation paints a different colour from the
  // identical style beside it. `tableLayoutStyle` has always resolved in this
  // order; this is the cell-level copy of that same rule (`tableTdEntry`).
  const sourceThemeColors = facts.sourceThemeColors;
  const blockStyleLook = facts.blockStyleLook;
  const nativeStyleCss =
    tableStyleLookCellCss(
      blockStyleLook,
      block,
      rowIndex,
      block.rows.length,
      logicalColIndex,
      colCount,
    ) ??
    (sourceThemeColors
      ? null
      : nativeTableStyleCellCss(
          block.tableStyleId,
          block,
          rowIndex,
          block.rows.length,
          logicalColIndex,
          colCount,
        ));
  // A line between two rows is the upper cell's `a:lnB` and one between two
  // columns the left cell's `a:lnR`: PowerPoint draws a cell's own `a:lnT`
  // only in the first row and its own `a:lnL` only in the first column, and
  // a later cell's top / start side is its style's line alone, which a line
  // the cell before states of its own beats (PowerPoint's prints:
  // census29-probes/compound-border-probe B2-D2 stating a 12 pt left line
  // beside cells stating none drew none; census30-probes probe 2 A3's right
  // Triple over B3's left Double drew the Triple; census31-probes probe 3 B2
  // stating only a top line drew none, C1's 2 pt bottom beat C2's 8 pt top,
  // and in Medium Style 2 a cell's own 8 pt top / left drew the style's 1 pt
  // inside line). The PDF export decides the same (`_paint_collapsed_edges`).
  const ownStyle = normalizeCellStyle(cell);
  const ownDrawn = {
    ...ownStyle,
    ...(rowIndex > 0 ? { borderTop: null } : {}),
    ...(logicalColIndex > 0 ? { borderLeft: null } : {}),
  };
  const css: CSSProperties = {
    ...(nativeStyleCss ?? {}),
    ...tableBorderCss(block.borders === true),
    ...cellStyleCss(ownDrawn, {
      headerEmphasis:
        nativeStyleCss === null && block.headerRow === true && rowIndex === 0,
      bandFill: nativeStyleCss === null ? bandFill : null,
      totalEmphasis: nativeStyleCss === null && totalEmphasis,
      colEmphasis: nativeStyleCss === null && colEmphasis,
      rightToLeft: block.rightToLeft === true,
    }),
    ...(nativeStyleCss === null && totalEmphasis
      ? { borderTop: TOTAL_ROW_BORDER_TOP }
      : {}),
    ...(cell.text.includes("\n") ? { whiteSpace: "pre-wrap" as const } : {}),
    // A cell anchored Top (DrawingML's default, `a:tcPr@anchor` absent) sits
    // at the top of a row its neighbour made taller. The cell's text box
    // cannot fill a `<td>` whose height is the row's, so it took the cell's
    // own default `vertical-align: middle` and every short cell beside a
    
    ...(importedPresentationTable(block) && ownStyle.verticalAlign === "top"
      ? { verticalAlign: "top" as const }
      : {}),
  };
  // A cell whose every run states bold off (`a:rPr b="0"`) is drawn regular
  // inside a style's bold row or column, as PowerPoint draws it: the run's
  // own formatting is over the table style's (CDC p60: "64" under a First
  // Row style drew bold).
  if (css.fontWeight !== undefined && cell.bold !== true && nativeCellRunsStateNotBold(cell)) {
    css.fontWeight = 400;
  }
  if (importedPresentationTable(block)) {
    // A side the file draws no line on draws none, as PowerPoint draws it -- not the canvas's faint default
    // cell rule (`.presentation-window__table-block td`), which also took half a pixel of every such edge
    // into the table's layout (census31-probes/compound-border-probe3: No Style, No Grid).
    const record = css as Record<string, unknown>;
    const shorthand = ["border", "borderWidth", "borderStyle", "borderColor"].some((key) => record[key] !== undefined);
    if (!shorthand) {
      for (const side of ["Top", "Bottom", "Left", "Right"] as const) {
        const name = `border${side}`;
        if (record[name] === undefined && record[`${name}Width`] === undefined && record[`${name}Style`] === undefined) {
          record[name] = "none";
        }
      }
    }
  }
  {
    const { above, before } = cellsBeforeEdges(block, facts, rowIndex, colIndex);
    if (above.length > 0 && above.every((upper) => normalizeCellStyle(upper).borderBottom !== null)) {
      dropTableCellEdge(css, "top");
    }
    if (logicalColIndex > 0) {
      // `none` loses every collapsed-border conflict, so the cell before it
      // decides an edge its style draws no line on, its line or none.
      if (
        (before.length > 0 && before.every((left) => normalizeCellStyle(left).borderRight !== null)) ||
        borderOf(css, "Left") === 0
      ) {
        dropTableCellEdge(css, "left");
      }
    }
  }
  if (importedPresentationTable(block)) {
    // PowerPoint centres a cell's top and bottom lines on the row's edges
    // and keeps them out of its height: a row is its text and margins
    
    // collapsed `<td>` adds half of each line to its row, so the cell's text
    // box takes those halves back (`tableCellVisualTextStyle`).
    // A collapsed edge draws the wider of the two cells' lines on it, half on
    // each side: the half inside this cell is the edge's line, its own or its
    
    // one stating a 0.5 px left line lost 0.25 px of its text width, and
    // "alternativa o disyuntiva", 121.63 pt on PowerPoint's 1/8 pt grid in a
    // 121.74 pt column, wrapped).
    const edge = withBorderHalves ? cellEdgeNeighbours(block, rowIndex, colIndex) : null;
    const top = Math.max(
      cssBorderWidthPx(css.borderTopWidth ?? css.borderTop ?? css.borderWidth ?? css.border),
      ...(edge?.above ?? []).map((neighbour) => borderOf(neighbour, "Bottom")),
    );
    const bottom = Math.max(
      cssBorderWidthPx(css.borderBottomWidth ?? css.borderBottom ?? css.borderWidth ?? css.border),
      ...(edge?.below ?? []).map((neighbour) => borderOf(neighbour, "Top")),
    );
    const halves = css as Record<string, unknown>;
    if (top > 0) halves[CELL_BORDER_TOP_HALF_VAR] = `${Math.round(drawnBorderWidthPx(top) * 50) / 100}px`;
    if (bottom > 0) halves[CELL_BORDER_BOTTOM_HALF_VAR] = `${Math.round(drawnBorderWidthPx(bottom) * 50) / 100}px`;
  }
  if (block.rightToLeft === true) {
    // `a:lnL` / `a:lnR` (and a style's left / right edges) are the cell's
    // START and END: PowerPoint draws `a:lnL` on the physical right of a
    // right-to-left table's cell (measured in PowerPoint for Mac).
    const { borderLeft, borderRight } = css;
    delete css.borderLeft;
    delete css.borderRight;
    if (borderRight !== undefined) css.borderLeft = borderRight;
    if (borderLeft !== undefined) css.borderRight = borderLeft;
    // A compound edge's rules go with its side (`TableCellCompoundEdges`
    // reads a side's rules beside that side's transparent border).
    const record = css as Record<string, unknown>;
    const leftRules = record[tableEdgeCompoundVar("left")];
    const rightRules = record[tableEdgeCompoundVar("right")];
    delete record[tableEdgeCompoundVar("left")];
    delete record[tableEdgeCompoundVar("right")];
    if (rightRules !== undefined) record[tableEdgeCompoundVar("left")] = rightRules;
    if (leftRules !== undefined) record[tableEdgeCompoundVar("right")] = leftRules;
  }
  if (importedPresentationTable(block)) {
    // Its left and right lines are centred on the column's edges too: the
    // cell's text runs the column's width less its margins, as PowerPoint
    // lays it out, not less the halves of its lines a collapsed `<td>` keeps
    
    // value PowerPoint keeps on one line).
    const edge = withBorderHalves ? cellEdgeNeighbours(block, rowIndex, colIndex) : null;
    const left = Math.max(
      cssBorderWidthPx(css.borderLeftWidth ?? css.borderLeft ?? css.borderWidth ?? css.border),
      ...(edge?.left ?? []).map((neighbour) => borderOf(neighbour, "Right")),
    );
    const right = Math.max(
      cssBorderWidthPx(css.borderRightWidth ?? css.borderRight ?? css.borderWidth ?? css.border),
      ...(edge?.right ?? []).map((neighbour) => borderOf(neighbour, "Left")),
    );
    const halves = css as Record<string, unknown>;
    if (left > 0) halves[CELL_BORDER_LEFT_HALF_VAR] = `${Math.round(drawnBorderWidthPx(left) * 50) / 100}px`;
    if (right > 0) halves[CELL_BORDER_RIGHT_HALF_VAR] = `${Math.round(drawnBorderWidthPx(right) * 50) / 100}px`;
  }
  if (withBorderHalves && importedPresentationTable(block)) {
    applyTableEdgeStrokes(block, rowIndex, colIndex, css);
  }
  return Object.keys(css).length > 0 ? css : undefined;
}

/** The strips a cell draws its table's lines with (`TableCellCompoundEdges`), one per run: see `tableEdgeStrokePlan`. */
export const TABLE_CELL_STROKES_VAR = "--presentation-cell-strokes";

type EdgeSide = "top" | "bottom" | "left" | "right";
const EDGE_SIDES: readonly EdgeSide[] = ["top", "bottom", "left", "right"];
const capitalSide = (side: EdgeSide) => `${side[0].toUpperCase()}${side.slice(1)}` as "Top";

/** One side's line of a cell's `<td>` style, as the stroke plan reads it. */
type EdgeLine = {
  /** The width the `<td>` keeps (a double / compound line's at least 3 px): the stroke's width. */
  width: number;
  style: "solid" | "dashed" | "dotted" | "double";
  color: string;
  /** A compound edge's rules (`tableEdgeCss`), across its width. */
  rules: string | null;
  /** The token written back on the `<td>`, the line made transparent. */
  transparent: string;
};

const COLLAPSE_STYLE_RANK: Record<EdgeLine["style"], number> = { double: 4, solid: 3, dashed: 2, dotted: 1 };

/** A side's line, `null` for none, `undefined` for one the plan cannot read (the table keeps CSS borders). */
function edgeLineOf(css: CSSProperties | undefined, side: EdgeSide): EdgeLine | null | undefined {
  if (!css) return null;
  const record = css as Record<string, unknown>;
  const name = `border${capitalSide(side)}`;
  if (["border", "borderWidth", "borderStyle", "borderColor", `${name}Width`, `${name}Style`, `${name}Color`].some(
    (key) => record[key] !== undefined,
  )) return undefined;
  const value = record[name];
  if (value === undefined || value === "none" || value === 0) return null;
  if (typeof value !== "string") return undefined;
  const parsed = /^(\d*\.?\d+)px (solid|dashed|dotted|double) (.+)$/.exec(value.trim());
  if (!parsed) return undefined;
  const width = Number(parsed[1]);
  const color = parsed[3].trim();
  const rulesValue = record[tableEdgeCompoundVar(side)];
  let rules: string | null = null;
  if (color === "transparent") {
    if (typeof rulesValue !== "string") return null;
    rules = rulesValue.slice(rulesValue.indexOf("|") + 1) || null;
    if (!rules) return null;
  }
  if (!(width > 0)) return null;
  return {
    width,
    style: parsed[2] as EdgeLine["style"],
    color,
    rules,
    transparent: `${parsed[1]}px solid transparent`,
  };
}

function edgeLineRank(line: EdgeLine): [number, number] {
  return [line.width, COLLAPSE_STYLE_RANK[line.style]];
}

/** Whether `a` beats `b` on a shared stretch of line (wider, then double, solid, dashed, dotted). */
function edgeLineBeats(a: EdgeLine, b: EdgeLine): boolean {
  const [aw, as] = edgeLineRank(a);
  const [bw, bs] = edgeLineRank(b);
  return aw > bw + 1e-9 || (Math.abs(aw - bw) <= 1e-9 && as > bs);
}

function translucentColor(color: string): boolean {
  return /rgba\(|hsla\(|color-mix|transparent|#[0-9a-f]{8}\b|#[0-9a-f]{4}\b/i.test(color);
}

/** The paint of one strip (`background`), across (`rules`) or along (dashes) its line. */
function edgeStripPaint(line: EdgeLine, horizontal: boolean): string {
  if (line.rules) return line.rules;
  const c = line.color;
  const w = line.width;
  if (line.style === "double") {
    return `linear-gradient(${horizontal ? "to bottom" : "to right"}, ${c} 0%, ${c} 33.3333%, transparent 33.3333%, transparent 66.6667%, ${c} 66.6667%, ${c} 100%)`;
  }
  if (line.style === "dashed" || line.style === "dotted") {
    // The PDF export's dashes (`_edge_stroke`): a dash 3 widths on, 2 off; a dot 1 on, 2 off; from the line's
    // start -- the left of a row line, the bottom of a column line (the PDF strokes a column line upwards).
    const on = (line.style === "dashed" ? 3 : 1) * w;
    const period = on + 2 * w;
    return `repeating-linear-gradient(${horizontal ? "to right" : "to top"}, ${c} 0px, ${c} ${round2(on)}px, transparent ${round2(on)}px, transparent ${round2(period)}px)`;
  }
  return c;
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

/** `calc(f * 100% + px)`, short where it can be. */
function fractionCalc(fraction: number, px: number): string {
  const f = round2(fraction * 100 * 100) / 100;
  const p = round2(px);
  if (f === 0) return `${p}px`;
  return `calc(${f}% + ${p}px)`;
}

type StrokePlan = {
  /** Per anchor cell (`row:col` of the rows array): the strips it draws, `TABLE_CELL_STROKES_VAR`'s value. */
  strokes: Map<string, string>;
  /** Per anchor cell: its lines made transparent (`border<Side>` tokens). */
  transparent: Map<string, Partial<Record<EdgeSide, string>>>;
};

const strokePlans = new WeakMap<PresentationTableBlock, { stamp: number; plan: StrokePlan | null }>();





















function tableEdgeStrokePlan(block: PresentationTableBlock): StrokePlan | null {
  const stamp = officeDataVersion();
  const known = strokePlans.get(block);
  if (known && known.stamp === stamp) return known.plan;
  const plan = buildTableEdgeStrokePlan(block);
  strokePlans.set(block, { stamp, plan });
  return plan;
}

function buildTableEdgeStrokePlan(block: PresentationTableBlock): StrokePlan | null {
  if (!importedPresentationTable(block) || !block.rows.length) return null;
  const facts = tableTdEntry(block).facts;
  const R = block.rows.length;
  const C = facts.colCount;
  if (!(C > 0)) return null;
  const rtl = block.rightToLeft === true;
  type Anchor = { key: string; r0: number; r1: number; p0: number; p1: number; lines: Record<EdgeSide, EdgeLine | null> };
  const anchors: Anchor[] = [];
  const cover = new Int32Array(R * C).fill(-1);
  const paints = new Set<string>();
  let forced = false;
  for (let r = 0; r < R; r += 1) {
    const row = block.rows[r];
    for (let c = 0; c < row.length; c += 1) {
      const cell = row[c];
      const col = facts.logicalCol(r, c);
      const rowspan = Math.max(1, Math.min(cell.rowspan ?? 1, R - r));
      const colspan = Math.max(1, Math.min(cell.colspan ?? 1, C - col));
      if (col < 0 || col >= C) continue;
      const css = tableCellTdStyle(block, cell, r, c, false);
      const lines = {} as Record<EdgeSide, EdgeLine | null>;
      for (const side of EDGE_SIDES) {
        const line = edgeLineOf(css, side);
        if (line === undefined) return null;
        lines[side] = line;
        if (line) {
          paints.add(`${line.style}|${line.color}|${line.width}`);
          if (line.rules || line.style !== "solid" || translucentColor(line.color)) forced = true;
        }
      }
      const p0 = rtl ? C - col - colspan : col;
      const index = anchors.length;
      anchors.push({ key: `${r}:${c}`, r0: r, r1: r + rowspan, p0, p1: p0 + colspan, lines });
      for (let y = r; y < r + rowspan; y += 1) for (let x = p0; x < p0 + colspan; x += 1) cover[y * C + x] = index;
    }
  }
  if (!forced && paints.size < 2) return null;

  type Segment = { owner: number; side: EdgeSide; line: EdgeLine };
  // Row lines H[y * C + x] (y 0..R) and column lines V[x * R + y] (x 0..C).
  const H: (Segment | null)[] = new Array((R + 1) * C).fill(null);
  const V: (Segment | null)[] = new Array((C + 1) * R).fill(null);
  // What Chromium lays each `<td>` out with: half of the widest collapsed line along each side.
  const halves = anchors.map(() => ({ top: 0, bottom: 0, left: 0, right: 0 }));
  const drawn = (line: EdgeLine | null) => (line ? drawnBorderWidthPx(line.width) : 0);
  const pick = (first: Segment | null, second: Segment | null): Segment | null => {
    if (!first) return second;
    if (!second) return first;
    return edgeLineBeats(second.line, first.line) ? second : first;
  };
  for (let y = 0; y <= R; y += 1) {
    for (let x = 0; x < C; x += 1) {
      const up = y > 0 ? cover[(y - 1) * C + x] : -1;
      const down = y < R ? cover[y * C + x] : -1;
      if (up === down) continue;
      const upper = up >= 0 ? anchors[up].lines.bottom : null;
      const lower = down >= 0 ? anchors[down].lines.top : null;
      const collapsed = Math.max(drawn(upper), drawn(lower)) / 2;
      if (up >= 0) halves[up].bottom = Math.max(halves[up].bottom, collapsed);
      if (down >= 0) halves[down].top = Math.max(halves[down].top, collapsed);
      H[y * C + x] = pick(
        upper ? { owner: up, side: "bottom", line: upper } : null,
        lower ? { owner: down, side: "top", line: lower } : null,
      );
    }
  }
  for (let x = 0; x <= C; x += 1) {
    for (let y = 0; y < R; y += 1) {
      const left = x > 0 ? cover[y * C + x - 1] : -1;
      const right = x < C ? cover[y * C + x] : -1;
      if (left === right) continue;
      const leftLine = left >= 0 ? anchors[left].lines.right : null;
      const rightLine = right >= 0 ? anchors[right].lines.left : null;
      const collapsed = Math.max(drawn(leftLine), drawn(rightLine)) / 2;
      if (left >= 0) halves[left].right = Math.max(halves[left].right, collapsed);
      if (right >= 0) halves[right].left = Math.max(halves[right].left, collapsed);
      const a = leftLine ? { owner: left, side: "right" as const, line: leftLine } : null;
      const b = rightLine ? { owner: right, side: "left" as const, line: rightLine } : null;
      // A tie goes to the cell before: the left one, the right one in a right-to-left table.
      V[x * R + y] = rtl ? pick(b, a) : pick(a, b);
    }
  }
  const at = <T,>(list: (T | null)[], index: number, inRange: boolean) => (inRange ? list[index] : null);
  const hAt = (y: number, x: number) => at(H, y * C + x, y >= 0 && y <= R && x >= 0 && x < C);
  const vAt = (x: number, y: number) => at(V, x * R + y, x >= 0 && x <= C && y >= 0 && y < R);
  const width = (segment: Segment | null) => (segment ? segment.line.width : 0);
  // The widest line crossing a row line `y` at column line `x`, and a column line `x` at row line `y`.
  const crossRow = (y: number, x: number) => Math.max(width(vAt(x, y - 1)), width(vAt(x, y)));
  const crossColumn = (x: number, y: number) => Math.max(width(hAt(y, x - 1)), width(hAt(y, x)));

  const colSizes = normalizeAxisSizes(block.colWidths, C);
  const rowSizes = normalizeAxisSizes(block.rowHeights, R);
  const cumulative = (sizes: Array<number | null> | undefined, count: number, reverse: boolean) => {
    const out = [0];
    for (let i = 0; i < count; i += 1) {
      const size = sizes?.[reverse ? count - 1 - i : i];
      out.push(out[i] + (typeof size === "number" && size > 0 ? size : 1));
    }
    return out;
  };
  const xs = cumulative(colSizes, C, rtl);
  const ys = cumulative(rowSizes, R, false);
  const layer = (horizontal: boolean, line: number) =>
    horizontal ? (line === 0 ? 5 : line === R ? 6 : 2) : line === 0 ? 3 : line === C ? 4 : 1;

  const strokes = new Map<string, string>();
  const transparent = new Map<string, Partial<Record<EdgeSide, string>>>();
  anchors.forEach((anchor, index) => {
    const own: Partial<Record<EdgeSide, string>> = {};
    for (const side of EDGE_SIDES) if (anchor.lines[side]) own[side] = anchor.lines[side]!.transparent;
    transparent.set(anchor.key, own);
    const h = halves[index];
    const strips: string[] = [];
    for (const side of EDGE_SIDES) {
      const horizontal = side === "top" || side === "bottom";
      const lineIndex = side === "top" ? anchor.r0 : side === "bottom" ? anchor.r1 : side === "left" ? anchor.p0 : anchor.p1;
      const from = horizontal ? anchor.p0 : anchor.r0;
      const to = horizontal ? anchor.p1 : anchor.r1;
      const segmentAt = (k: number) => (horizontal ? hAt(lineIndex, k) : vAt(lineIndex, k));
      let k = from;
      while (k < to) {
        const segment = segmentAt(k);
        if (!segment || segment.owner !== index || segment.side !== side) {
          k += 1;
          continue;
        }
        const a = k;
        while (k < to && segmentAt(k)?.owner === index && segmentAt(k)?.side === side) k += 1;
        const b = k;
        const line = segment.line;
        const before = (horizontal ? crossRow(lineIndex, a) : crossColumn(lineIndex, a)) / 2;
        const after = (horizontal ? crossRow(lineIndex, b) : crossColumn(lineIndex, b)) / 2;
        const goesOn = !!(horizontal ? hAt(lineIndex, b) : vAt(lineIndex, b));
        const endShift = goesOn ? -after : after;
        const grid = horizontal ? xs : ys;
        const span = grid[to] - grid[from];
        const fa = span > 0 ? (grid[a] - grid[from]) / span : 0;
        const fb = span > 0 ? (grid[b] - grid[from]) / span : 1;
        // The padding box the strip is placed in runs from the cell's edge plus its half line at each end.
        const [startHalf, endHalf] = horizontal ? [h.left, h.right] : [h.top, h.bottom];
        const startAlong = fractionCalc(fa, fa * (startHalf + endHalf) - startHalf - before);
        const endAlong = fractionCalc(1 - fb, (1 - fb) * (startHalf + endHalf) - endHalf - endShift);
        const across = `${round2(-h[side] - line.width / 2)}px`;
        const thickness = `${round2(line.width)}px`;
        const rect = horizontal
          ? { left: startAlong, right: endAlong, top: side === "top" ? across : "auto", bottom: side === "bottom" ? across : "auto", width: "auto", height: thickness }
          : { top: startAlong, bottom: endAlong, left: side === "left" ? across : "auto", right: side === "right" ? across : "auto", width: thickness, height: "auto" };
        strips.push(
          [layer(horizontal, lineIndex), rect.left, rect.top, rect.right, rect.bottom, rect.width, rect.height, edgeStripPaint(line, horizontal)].join("|"),
        );
      }
    }
    if (strips.length) strokes.set(anchor.key, strips.join(";"));
  });
  return { strokes, transparent };
}

/** A cell's lines handed to its strips (`tableEdgeStrokePlan`): transparent on the `<td>`, drawn by the strips. */
function applyTableEdgeStrokes(block: PresentationTableBlock, rowIndex: number, colIndex: number, css: CSSProperties): void {
  const plan = tableEdgeStrokePlan(block);
  if (!plan) return;
  const key = `${rowIndex}:${colIndex}`;
  const record = css as Record<string, unknown>;
  for (const side of EDGE_SIDES) {
    delete record[tableEdgeCompoundVar(side)];
    delete record[`${tableEdgeCompoundVar(side)}-ends`];
  }
  const own = plan.transparent.get(key);
  if (own) for (const side of EDGE_SIDES) if (own[side]) record[`border${capitalSide(side)}`] = own[side];
  const strokes = plan.strokes.get(key);
  if (strokes) {
    record[TABLE_CELL_STROKES_VAR] = strokes;
    record.position = "relative";
  }
}

/** What a cell's style takes from its table's stroke plan, for the cell's style memo. */
function tableEdgeStrokeIdentity(block: PresentationTableBlock, rowIndex: number, colIndex: number): unknown {
  if (!importedPresentationTable(block)) return null;
  const plan = tableEdgeStrokePlan(block);
  if (!plan) return null;
  return `${plan.strokes.get(`${rowIndex}:${colIndex}`) ?? ""}`;
}

/**
 * The `<colgroup>` that applies a table block's optional per-column widths, or
 * `null` when the table carries no `colWidths` (so an un-sized table renders the
 * native auto-layout `<table>` exactly as before). One `<col>` per column (the
 * widest row's length); a slot with no size omits its `width` so that column
 * stays auto. Shared by the editor canvas + present overlay so both size the
 * same. `<colgroup>` widths only take effect with a fixed table layout, so the
 * caller pairs this with `tableLayoutStyle`.
 */
export function drawTableGrid(
  block: PresentationTableBlock,
  boxes?: TableBorderCellBox[],
) {
  const positions = logicalCellPositions(block.rows);
  const axis = (
    values: unknown,
    count: number,
    extent: number,
    horizontal: boolean,
  ): number[] => {
    if (
      Array.isArray(values) &&
      values.length === count &&
      values.every(
        (value) =>
          typeof value === "number" && Number.isFinite(value) && value > 0,
      )
    )
      return values;
    // Auto heights are measured from this mounted table, never evenly guessed
    // from its frame. A missing/zero/ambiguous boundary rejects the draw.
    if (
      !boxes?.length ||
      boxes.some((box) => !(box.width > 0 && box.height > 0))
    )
      return Array(count).fill(0);
    const first = Math.min(
      ...boxes.map((box) => (horizontal ? box.left : box.top)),
    );
    const last = Math.max(
      ...boxes.map((box) =>
        horizontal ? box.left + box.width : box.top + box.height,
      ),
    );
    // (`last - first` is the drawn extent, the same along a mirrored axis.)
    if (!(last > first) || !(extent > 0)) return Array(count).fill(0);
    const boundaries: Array<number | undefined> = Array(count + 1).fill(
      undefined,
    );
    // A right-to-left table lays logical column 0 out on the right: read its
    // columns along a mirrored x so the boundaries come out in logical order.
    const mirrored = horizontal && block.rightToLeft === true;
    for (const box of boxes) {
      const at = positions[box.row]?.[box.col];
      const cell = block.rows[box.row]?.[box.col];
      if (!at || !cell) return Array(count).fill(0);
      const index = horizontal ? at.col : at.row;
      const span = horizontal ? (cell.colspan ?? 1) : (cell.rowspan ?? 1);
      const start = mirrored
        ? -(box.left + box.width)
        : horizontal
          ? box.left
          : box.top;
      const end = start + (horizontal ? box.width : box.height);
      for (const [slot, point] of [
        [index, start],
        [index + span, end],
      ]) {
        if (
          boundaries[slot] !== undefined &&
          Math.abs(boundaries[slot]! - point) > 1
        )
          return Array(count).fill(0);
        boundaries[slot] = point;
      }
    }
    return Array.from({ length: count }, (_, i) =>
      boundaries[i] === undefined || boundaries[i + 1] === undefined
        ? 0
        : Math.round(
            ((boundaries[i + 1]! - boundaries[i]!) / (last - first)) * extent,
          ) / 9525,
    );
  };
  return {
    rows: block.rows,
    colWidths: axis(
      block.colWidths,
      tableLogicalWidth(block.rows),
      geometryNumber(block, "width") ?? 0,
      true,
    ),
    rowHeights: axis(
      block.rowHeights,
      block.rows.length,
      geometryNumber(block, "height") ?? 0,
      false,
    ),
  };
}

export function TableColgroup({
  block,
}: {
  block: PresentationTableBlock;
}): ReactElement | null {
  const columnCount = tableLogicalWidth(block.rows);
  const widths = block.rows.some(row => row.some(cell => cell.drawCellId))
    ? drawTableGrid(block).colWidths
    : normalizeAxisSizes(block.colWidths, columnCount);
  if (!widths) return null;
  return (
    <colgroup>
      {Array.from({ length: columnCount }, (_unused, col) => {
        const width = widths[col] ?? null;
        return (
          <col
            key={`col-${col}`}
            style={width !== null ? { width } : undefined}
          />
        );
      })}
    </colgroup>
  );
}

/**
 * The `<table>` inline style that switches to FIXED layout when ANY axis is
 * sized (so the `<colgroup>` widths + row heights are honoured), else `undefined`
 * (native auto layout, byte-stable). A fixed layout is required for explicit
 * column widths to take effect.
 */










export function tableOuterLineHalvesCss(block: PresentationTableBlock): CSSProperties | undefined {
  if (!importedPresentationTable(block) || !block.rows.length || block.rightToLeft === true) return undefined;
  const positions = logicalCellPositions(block.rows);
  const width = (css: CSSProperties | undefined, side: "Left" | "Top") => (css ? drawnBorderWidthPx(borderOf(css, side)) : 0);
  let left = 0;
  block.rows.forEach((row, rowIndex) => {
    if (row.length && positions[rowIndex]?.[0]?.col === 0) left = Math.max(left, width(tableCellTdStyle(block, row[0], rowIndex, 0), "Left"));
  });
  let top = 0;
  block.rows[0].forEach((cell, colIndex) => {
    top = Math.max(top, width(tableCellTdStyle(block, cell, 0, colIndex), "Top"));
  });
  if (left <= 0 && top <= 0) return undefined;
  return {
    ...(left > 0 ? { marginLeft: -Math.round(left * 50) / 100 } : {}),
    ...(top > 0 ? { marginTop: -Math.round(top * 50) / 100 } : {}),
  };
}

/** The table style look a table is drawn with: its own, else its built-in style's. */
export function tableDrawnStyleLook(block: PresentationTableBlock): unknown {
  return normalizeTableStyleLook(block.tableStyleLook, block.tableStyleId) ??
    (nativeTableThemeColorAuthority(block.source)
      ? null
      : nativeWebTableStyleById(block.tableStyleId)?.look);
}

export function tableLayoutStyle(
  block: PresentationTableBlock,
): CSSProperties | undefined {
  const columnCount = tableLogicalWidth(block.rows);
  const hasCols = !!normalizeAxisSizes(block.colWidths, columnCount);
  const css = {
    // A cell naming no typeface draws in the theme's body font, as PowerPoint
    // draws it (`presentationBodyFontStyle` states it on the slide surface).
    fontFamily: "var(--presentation-body-font)",
    // PowerPoint's shaping of the cells' words where the browser draws them
    // (`presentationTextShapingCss`): no common ligatures; unkerned when the
    // table's kerning receipt says PowerPoint kerns none of them.
    ...presentationTextShapingCss(((block as { source?: unknown }).source as { nativeKerning?: unknown } | null | undefined)?.nativeKerning, null),
    ...(hasCols ? { tableLayout: "fixed" as const, width: "100%" } : {}),
    // An imported table is as tall as its rows, as PowerPoint draws it: the
    // frame's own extent is not its size (a Google Slides export states
    // 3,000,000 EMU square frames over 1.3 M EMU of rows, and the rows were
    
    ...(hasLayoutGeometry(block) && !importedPresentationTable(block) ? { height: "100%" } : {}),
    // `a:tblPr@rtl` (ECMA-376): logical column 0 is laid out on the RIGHT.
    // A table's column progression follows its own `direction`; the cells
    // are reset to LTR by `tableRtl` so only their placement mirrors.
    ...(block.rightToLeft === true ? { direction: "rtl" as const } : {}),
    ...tableStyleLookTableCss(
      tableDrawnStyleLook(block),
      // The frame in EMU (`presentationTableFrameCss` reads the same).
      frameSize(block as unknown as { readonly width?: unknown; readonly height?: unknown }),
      // Its background shadow is cast by the cells' paint (`PresentationTableShadowFrame`).
      { shadow: false },
    ),
    // Table Background (`a:tblPr` fill), imported or picked in Table Design >
    // Shading, under every cell's own fill; the table's Shadow and
    // Reflection on its frame (PPT-A137).
    ...presentationTableFrameCss(block),
  };
  return Object.keys(css).length > 0 ? css : undefined;
}

/**
 * A right-to-left table's marker: the class that keeps its cells' own text
 * LTR and mirrors the cell-edge handles, and the attribute the pointer tools
 * read to know the table's x axis runs from the right.
 */
export function tableDirectionProps(block: PresentationTableBlock): {
  className?: string;
  "data-table-rtl"?: "true";
} {
  return block.rightToLeft === true
    ? { className: presentationChrome.tableRtl, "data-table-rtl": "true" }
    : {};
}

/**
 * The `<tr>` inline style of each row of a table block: its optional row
 * height, or `undefined` when the row has no explicit height (auto,
 * byte-stable). `height` on a `<tr>` acts as a MINIMUM (content can still
 * grow it), matching PowerPoint. Worked out for all rows at once: asked row
 * by row, the visual minima (each cell's wrapped lines, over the table's
 * whole merge grid) were worked out again for every row, a table's whole
 * layout once per row on every render of the slide.
 */
export function tableRowHeightStyles(
  block: PresentationTableBlock,
): (CSSProperties | undefined)[] {
  const count = block.rows.length;
  if (block.rows.some(row => row.some(cell => cell.drawCellId))) {
    const { rowHeights } = drawTableGrid(block);
    return Array.from({ length: count }, (_, rowIndex) => ({ height: rowHeights[rowIndex] }));
  }
  const rowHeights = normalizeAxisSizes(block.rowHeights, count);
  if (rowHeights) {
    // The rows as stored (`normalizeAxisSizes`): the 16 px sizing minimum
    // (`axisSizeAt`) is for what a person sets, and an imported row PowerPoint
    
    return Array.from({ length: count }, (_, rowIndex) => {
      const height = rowHeights[rowIndex] ?? null;
      return height !== null ? { height } : undefined;
    });
  }
  if (!hasLayoutGeometry(block) || count === 0) return Array.from({ length: count }, () => undefined);
  // A geometry table must occupy its authored frame without letting textarea
  // intrinsic rows grow the HTML table through the frame. Keep the model
  // untouched and distribute only otherwise-auto rows by their visual minima.
  const minima = tableRowMinimumHeightsPx(block);
  const total = minima.reduce((sum, rowHeight) => sum + rowHeight, 0);
  return minima.map((rowHeight) => total > 0
    ? { height: `${(rowHeight / total) * 100}%` }
    : undefined);
}

/** The paragraph-alignment icon for a table cell's alignment: Distributed,
 * Thai Distributed and Justify Low (imported cells) draw as Justify. */
export function tableCellAlignIcon(
  value: string | null | undefined,
): "left" | "center" | "right" | "justify" {
  if (value === "center" || value === "right" || value === "justify") return value;
  if (value === "distributed" || value === "thaiDistributed" || value === "justifyLow") {
    return "justify";
  }
  return "left";
}

const MIN_TABLE_CANVAS_ROW_PX = 46;

const TABLE_CELL_FALLBACK_VERTICAL_MARGIN_PX = 6;

const TABLE_CELL_FALLBACK_HORIZONTAL_MARGIN_PX = 10;

const TABLE_CELL_LINE_HEIGHT = 1.5;

const TABLE_FRAME_EMU_PER_CENTIMETRE_HUNDREDTH =
  TABLE_FRAME_EMU_PER_CM / 100;

function tableCellAverageGlyphAdvanceEm(text: string): number {
  const characters = Array.from(text);
  if (characters.length === 0) return 0.52;
  const total = characters.reduce((width, character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    return width + (codePoint > 0xff ? 1 : 0.52);
  }, 0);
  return total / characters.length;
}

/**
 * The largest size a cell's text is drawn at, in CSS px: the cell's own
 * size, or a larger one some of its characters state. A row can be no lower
 * than its tallest characters, whichever way their size was set; measured by
 * the cell's size alone, a table shrank past a 48 pt word set on the
 * characters and its bottom grips sat inside the cell (PPT-B324).
 */
function tableCellDrawnFontSizePx(
  cell: PresentationTableCell,
  cellFontSizePt: number | null,
): number {
  let largest = cellFontSizePt === null ? 16 : cellFontSizePt * CSS_PX_PER_POINT;
  for (const run of cellCharacterRuns(String(cell.text ?? ""), (cell as { runs?: unknown }).runs)) {
    const size = run.style?.fontSize;
    // A paragraph break's own size (an empty paragraph's end mark, readings
    // row 807) is that one line's, not every line's: the drawn cell stands it.
    if (run.text && !/^\n+$/.test(run.text) && typeof size === "number" && Number.isFinite(size) && size > largest) largest = size;
  }
  return largest;
}

/** Every line the cell's text takes, however many: the table's frame cannot
 * be shorter than its content (PPT-B416; an editor's eight-row cap here let
 * a ninth line hang below the frame). */
function tableCellVisualLineCount(
  block: PresentationTableBlock,
  rowIndex: number,
  cellIndex: number,
  /** The table's merge grid, worked out once for all its cells. */
  grid: { columnCount: number; positions: ReturnType<typeof logicalCellPositions> } = {
    columnCount: tableLogicalWidth(block.rows),
    positions: logicalCellPositions(block.rows),
  },
): number {
  const cell = block.rows[rowIndex]?.[cellIndex];
  if (!cell) return 1;
  const style = normalizeCellStyle(cell);
  const explicitLines = Math.max(
    1,
    // A paragraph (Enter) or a line break inside one (Shift+Enter, `\v`).
    String(cell.text ?? "").split(/\r?\n|\v/).length,
  );
  if (style.textDirection !== "horizontal") {
    return explicitLines;
  }

  const { columnCount } = grid;
  const position = grid.positions[rowIndex]?.[cellIndex];
  if (!position || columnCount <= 0) {
    return explicitLines;
  }
  const frameWidthPx =
    (geometryNumber(block, "width") ??
      columnCount * DEFAULT_DISTRIBUTE_AXIS_PX * TABLE_FRAME_EMU_PER_CSS_PX) /
    TABLE_FRAME_EMU_PER_CSS_PX;
  const columnSpan = Math.max(1, Math.floor(Number(cell.colspan) || 1));
  let cellWidthPx = 0;
  for (let offset = 0; offset < columnSpan; offset += 1) {
    const columnIndex = position.col + offset;
    if (columnIndex >= columnCount) break;
    cellWidthPx +=
      materializedAxisSizeAt(
        block.colWidths,
        columnCount,
        columnIndex,
        frameWidthPx,
      ) ?? frameWidthPx / columnCount;
  }
  const marginLeft =
    style.marginLeft ?? TABLE_CELL_FALLBACK_HORIZONTAL_MARGIN_PX;
  const marginRight =
    style.marginRight ?? TABLE_CELL_FALLBACK_HORIZONTAL_MARGIN_PX;
  const contentWidthPx = Math.max(1, cellWidthPx - marginLeft - marginRight);
  const fontSizePx = tableCellDrawnFontSizePx(cell, style.fontSize);
  return Math.max(
    explicitLines,
    estimateWrappedLineCount(
      String(cell.text ?? ""),
      contentWidthPx,
      fontSizePx,
      tableCellAverageGlyphAdvanceEm(String(cell.text ?? "")),
    ),
  );
}

function tableRowMinimumHeightsPx(
  block: PresentationTableBlock,
): number[] {
  const minima = Array.from(
    { length: block.rows.length },
    () => MIN_TABLE_CANVAS_ROW_PX,
  );
  const positions = logicalCellPositions(block.rows);
  const grid = { columnCount: tableLogicalWidth(block.rows), positions };
  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex += 1) {
    for (
      let cellIndex = 0;
      cellIndex < block.rows[rowIndex].length;
      cellIndex += 1
    ) {
      const cell = block.rows[rowIndex][cellIndex];
      const position = positions[rowIndex]?.[cellIndex];
      if (!position) continue;
      const style = normalizeCellStyle(cell);
      const fontSizePx = tableCellDrawnFontSizePx(cell, style.fontSize);
      const lineHeightPx = fontSizePx * TABLE_CELL_LINE_HEIGHT;
      const lineCount = tableCellVisualLineCount(
        block,
        rowIndex,
        cellIndex,
        grid,
      );
      const marginTop =
        style.marginTop ?? TABLE_CELL_FALLBACK_VERTICAL_MARGIN_PX;
      const marginBottom =
        style.marginBottom ?? TABLE_CELL_FALLBACK_VERTICAL_MARGIN_PX;
      const rowSpan = Math.max(1, Math.floor(Number(cell.rowspan) || 1));
      const requiredPerRow =
        (lineHeightPx * lineCount + marginTop + marginBottom) / rowSpan;
      for (let offset = 0; offset < rowSpan; offset += 1) {
        const logicalRow = position.row + offset;
        if (logicalRow >= minima.length) break;
        minima[logicalRow] = Math.max(minima[logicalRow], requiredPerRow);
      }
    }
  }
  return minima.map((height) => Math.ceil(height * 10) / 10);
}

/**
 * A whole-table handle changes the native frame AND its DrawingML grid. The
 * generic block gesture previously changed only width/height, leaving the old
 * row/column sizes behind a smaller clipped frame. Preserve the dragged edge,
 * scale only axes that changed, and keep the frame large enough for the fixed
 * cell editor baseline so content cannot be covered by its own selection box.
 */
export function tableWithScaledGestureGeometry(
  block: PresentationTableBlock,
  requested: BlockGeometry,
): PresentationTableBlock {
  const currentLeft = geometryNumber(block, "left") ?? requested.left;
  const currentTop = geometryNumber(block, "top") ?? requested.top;
  const currentWidth = geometryNumber(block, "width") ?? requested.width;
  const currentHeight = geometryNumber(block, "height") ?? requested.height;
  const widthChanged = Math.round(requested.width) !== Math.round(currentWidth);
  const heightChanged =
    Math.round(requested.height) !== Math.round(currentHeight);
  const minimumWidth = Math.max(
    1,
    tableLogicalWidth(block.rows) *
      MIN_TABLE_AXIS_PX *
      TABLE_FRAME_EMU_PER_CSS_PX,
  );
  const width = clampNumber(
    Math.round(
      widthChanged
        ? Math.max(requested.width, minimumWidth)
        : requested.width,
    ),
    1,
    MAX_PRESENTATION_GEOMETRY_EMU,
  );
  let resized = block;
  if (widthChanged) {
    resized = resizeTableFrame(
      resized,
      "width",
      width / TABLE_FRAME_EMU_PER_CM,
    );
  }
  const minimumHeight = Math.max(
    1,
    Math.ceil(
      (tableRowMinimumHeightsPx(resized).reduce(
        (total, rowHeight) => total + rowHeight,
        0,
      ) *
        TABLE_FRAME_EMU_PER_CSS_PX) /
        TABLE_FRAME_EMU_PER_CENTIMETRE_HUNDREDTH,
    ) * TABLE_FRAME_EMU_PER_CENTIMETRE_HUNDREDTH,
  );
  const height = clampNumber(
    Math.round(
      heightChanged || widthChanged
        ? Math.max(requested.height, minimumHeight)
        : requested.height,
    ),
    1,
    MAX_PRESENTATION_GEOMETRY_EMU,
  );
  const requestedRight = requested.left + requested.width;
  const requestedBottom = requested.top + requested.height;
  const left = clampNumber(
    Math.round(
      widthChanged && requested.left !== currentLeft
        ? requestedRight - width
        : requested.left,
    ),
    -MAX_PRESENTATION_GEOMETRY_EMU,
    MAX_PRESENTATION_GEOMETRY_EMU,
  );
  const top = clampNumber(
    Math.round(
      heightChanged && requested.top !== currentTop
        ? requestedBottom - height
        : requested.top,
    ),
    -MAX_PRESENTATION_GEOMETRY_EMU,
    MAX_PRESENTATION_GEOMETRY_EMU,
  );
  if (heightChanged) {
    resized = resizeTableFrame(
      resized,
      "height",
      height / TABLE_FRAME_EMU_PER_CM,
    );
  }
  return { ...resized, left, top, width, height };
}

/**
 * Imported DrawingML can carry explicit grid sizes whose sum differs slightly
 * (or substantially) from the authored frame after conversion. Keep that
 * source model untouched, but render its explicit grid proportionally inside
 * the frame so selection never reveals a wider/taller table beneath it.
 */










export function importedTableRowMinimums(block: PresentationTableBlock): readonly unknown[] | null {
  return importedPresentationTable(block) && Array.isArray(block.rowHeights) ? block.rowHeights : null;
}

export function tableWithFrameAlignedGrid(
  block: PresentationTableBlock,
): PresentationTableBlock {
  if (block.rows.some(row => row.some(cell => cell.drawCellId))) return block;
  let aligned = block;
  let width = geometryNumber(block, "width");
  const height = geometryNumber(block, "height");
  const colWidths = normalizeAxisSizes(block.colWidths, tableLogicalWidth(block.rows));
  // An imported table's grid is exact (`a:gridCol`); PowerPoint draws the
  // grid and lets the frame follow it, so a graphicFrame extent smaller than
  // its columns (R18: 3 000 000 EMU over 8 487 975 of columns) must not
  // squeeze the columns to a third and push the rows off the page (PPT-B372).
  if (width !== null && colWidths && isImportedPptxBlock(block)) {
    const frameWidth: number = width;
    const gridEmu =
      colWidths.reduce<number>((sum, size) => sum + (typeof size === "number" ? size : 0), 0) *
      TABLE_FRAME_EMU_PER_CSS_PX;
    if (gridEmu > frameWidth * 1.02) {
      width = Math.round(gridEmu);
      aligned = { ...aligned, width };
    }
  }
  if (width !== null && colWidths) {
    aligned = resizeTableFrame(
      aligned,
      "width",
      width / TABLE_FRAME_EMU_PER_CM,
    );
  }
  const rowHeights = normalizeAxisSizes(block.rowHeights, block.rows.length);
  // An imported row stated `h="0"` (no height) is as tall as its words, as
  // PowerPoint draws it, and the frame's extent is only the cache of what the
  // rows came to. Fitting the rows to the frame handed such a row the whole
  
  
  // rows). Rows that state a height keep it as their minimum and grow with
  // their words; the table is as tall as its rows (`tableLayoutStyle`).
  const importedAutoRow =
    rowHeights !== undefined &&
    importedPresentationTable(block) &&
    block.rows.some((_, rowIndex) => axisSizeAt(rowHeights, rowIndex) === null);
  // An imported table whose stated rows already fill its frame keeps them as
  // stated: PowerPoint's row is as low as its `a:tr@h` and its words allow,
  
  // 29 rows of h="115200", 9.07 pt, drawn 12 pt apart and run 85 pt past
  // PowerPoint's last one).
  const statedRowsPx = importedPresentationTable(block) && Array.isArray(block.rowHeights) &&
      block.rowHeights.length === block.rows.length &&
      block.rowHeights.every((size) => typeof size === "number" && Number.isFinite(size) && size > 0)
    ? (block.rowHeights as number[])
    : null;
  if (
    height !== null &&
    statedRowsPx &&
    Math.abs(statedRowsPx.reduce((sum, size) => sum + size, 0) - height / TABLE_FRAME_EMU_PER_CSS_PX) <= 1 &&
    statedRowsPx.some((size) => size < MIN_TABLE_AXIS_PX)
  ) {
    return { ...aligned, rowHeights: statedRowsPx };
  }
  // An imported table's rows are its stated heights, each a minimum its words
  // grow, and the frame follows the rows -- never the rows the frame
  // (`importedTableRowMinimums`): a frame lower than its rows (Google's stale
  
  
  
  
  // p19, 525.6) handed each row a share in proportion to its stated height,
  // not what its words take.
  if (
    height !== null &&
    rowHeights &&
    !importedAutoRow &&
    importedTableRowMinimums(block) === null
  ) {
    aligned = resizeTableFrame(
      aligned,
      "height",
      height / TABLE_FRAME_EMU_PER_CM,
    );
  }
  return aligned;
}
