"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PX_PER_PT, pxToPt } from "@mrnote/core/web/lib/office-units.ts";
import { officeObjectKeyAction } from "@mrnote/core/web/lib/office-object-nudge.ts";
import { isMacPlatform } from "@mrnote/core/web/lib/mac-platform.ts";
import { officeLocaleUsesInches } from "@mrnote/core/web/lib/office-measurement-units.ts";
import type { TranslationFn } from "@mrnote/core/web/lib/errors/user-message.ts";

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type InputHTMLAttributes,
  type FocusEvent as ReactFocusEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
  type ReactElement,
  type ReactNode,
  type RefObject,
} from "react";
import { flushSync } from "react-dom";

import {
  resolveHeaderFooterDocumentScale,
  resolveHeaderFooterSegment,
  resolvePageHeaderFooter,
  resolvePageHeaderFooterVariant,
  resolveSpreadsheetPageLayoutGeometry,
  resolveSpreadsheetPageLayoutRenderPlan,
  resolveSpreadsheetPrintBatchPlan,
  resolveSpreadsheetPageLayoutTileGrid,
  resolveSpreadsheetPageLayoutTilePosition,
  pageLayoutDrawingBands,
  pageLayoutAxisSourceAt,
  pageLayoutDrawingHandlePoint,
  pageLayoutEdgeScrollStep,
  transformPageLayoutDrawing,
  PAGE_LAYOUT_DRAWING_HANDLES,
  type PageLayoutDrawingFrame,
  type PageLayoutDrawingHandle,
  type ResolvedHeaderFooterSegment,
} from "../../../../lib/spreadsheet/page-layout-view";
import { extendPageLayoutForBrowsing, resolvePageLayoutViewportPlan } from "../../../../lib/spreadsheet/page-layout-navigation";
import { transformChartDrawingGroup } from "../../../../lib/spreadsheet/chart-groups";
import { officeDrawingSelectionStyles } from "@mrnote/core/web/components/office-ribbon/OfficeDrawingSelection.tsx";
import { SHEET_DRAWING_HANDLE, transformSpreadsheetDrawingSelection } from "../../../../lib/spreadsheet/drawing-selection";
import {
  clampColumnWidth,
  clampRowHeight,
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  MAX_COLUMN_WIDTH_PX,
  MAX_ROW_HEIGHT_PX,
  MIN_COLUMN_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
  resolvedColumnWidth,
  resolvedRowHeight,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import type {
  HeaderFooterContext,
  PrintHeaderFooter,
  PrintHeaderFooterRun,
  PrintMargins,
  PrintSettings,
} from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import {
  MAX_HEADER_FOOTER_SECTION_TEXT,
  normalizePrintSettings,
} from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import {
  headerFooterRunCss,
  headerFooterRunStyle,
} from "../../../../lib/spreadsheet/header-footer-rich-text";
import { clampNativeRowHeight, nativeRowCssClamp } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";

import type { SpreadsheetCellControlFrame } from "./SpreadsheetCellCheckbox";
import { registerSpreadsheetCellMaterializer, retargetSpreadsheetCellEvent } from "./spreadsheetCellLookup";
import { withHeaderFooterPicture } from "./SpreadsheetHeaderFooterPicture";

export interface SpreadsheetPageLayoutViewProps {
  sheetId: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  hiddenRows?: ReadonlySet<number>;
  hiddenCols?: ReadonlySet<number>;
  sourceRowAt?: (visualRow: number) => number;
  rowOrder?: readonly number[];
  printSettings?: PrintSettings;
  fileName: string;
  filePath?: string;
  locale: string;
  translate: TranslationFn;
  zoom: number;
  showGridlines: boolean;
  rowHeights?: Readonly<Record<string, number>>;
  columnWidths?: Readonly<Record<string, number>>;
  /**
   * Where a cell's text may run over its empty neighbours (the grid's text
   * spill), as a column range, or null when it stays in its cell. Asked for
   * unmerged cells with the cell's drawn width; the run is painted across
   * those columns and clipped at the page's edge.
   */
  getTextSpill?: (
    row: number,
    col: number,
    cellWidthPx: number,
    fontScale: number,
  ) => { startCol: number; endCol: number } | null;
  defaultRowHeight?: number;
  nativeRowHeights?: boolean;
  nativeTextDpi?: number;
  /** The grid font's maximum digit width in device pixels at nativeTextDpi:
   * printed row numbers are counted in it. */
  nativeDigitWidth?: number;
  /** The workbook's Normal font: printed headings are drawn in it. */
  headingFont?: { family: string; sizePt: number };
  /** Page Setup > Sheet > Comments and notes: As displayed on sheet. The
   * shown notes, by "sourceRow:col", printed where they show (print only). */
  printedNotes?: ReadonlyMap<string, string>;
  hasNativeTextLayout?: (row: number, col: number, display: string) => boolean;
  defaultColumnWidth?: number;
  merges: readonly {
    row: number;
    col: number;
    rowSpan: number;
    colSpan: number;
  }[];
  drawings: readonly SpreadsheetPageLayoutDrawing[];
  /**
   * The sheet's background picture (Page Layout > Background), tiled behind
   * each page's cells on screen. Never part of a printout: the print path
   * does not draw it.
   */
  backgroundPicture?: string;
  /**
   * The sheet runs right to left (sheetView rightToLeft): each page lays its
   * columns out from the right, column A against the right margin, as the
   * Normal view and the printout do.
   */
  rightToLeft?: boolean;
  onSelectDrawing?: (id: string, additive?: boolean, preserve?: boolean) => void;
  selectedDrawingIds?: readonly string[];
  onClearDrawingSelection?: () => void;
  selectedDrawingId?: string | null;
  onBeginDrawingTransform?: (id: string) => PageLayoutDrawingTransaction | null;
  /** Alt + Left / Right on a selected drawing: turn it by `degrees` through
   * Normal view's rotation writer (a shape or a picture, as there). */
  onRotateDrawing?: (id: string, degrees: number) => boolean;
  activeCell: Readonly<{ row: number; col: number }>;
  /** The cell being edited, when it is on this sheet: it keeps its editor
   * wherever the active cell is. */
  editingCell?: Readonly<{ row: number; col: number }> | null;
  selection: Readonly<{ r1: number; c1: number; r2: number; c2: number }>;
  viewportRef: RefObject<HTMLDivElement | null>;
  /** Interactive worksheet browsing; print-only previews stay content-bounded. */
  browseBlankPages?: boolean;
  /**
   * Printing (/render/spreadsheet-print): every page mounted at once, no
   * page chrome, no selection, gridlines as Page Setup prints them, empty
   * header and footer sections empty; mounted a batch at a time from
   * printPageStart (resolveSpreadsheetPrintBatchPlan).
   */
  printMode?: boolean;
  printPageStart?: number;
  getCellDisplay: (row: number, col: number) => string;
  getCellStyle?: (row: number, col: number) => CSSProperties | undefined;
  renderCellOverlay?: (row: number, col: number, occurrence: string, frame: SpreadsheetCellControlFrame) => ReactNode;
  renderCellContent?: (row: number, col: number, fontScale: number, nativeTextWidth?: number, frame?: { width: number; height: number }) => ReactNode;
  /** A cell whose painting leans past its box (an angled cell's band). */
  cellPaintsOutside?: (row: number, col: number) => boolean;
  getCellInputProps: (
    row: number,
    col: number,
    occurrence: string,
  ) => InputHTMLAttributes<HTMLInputElement> & {
    "data-cell-text"?: "layered";
    "data-editing"?: "true" | "false";
  };
  onColumnResizeStart: (
    col: number,
    event: ReactPointerEvent<HTMLElement>,
  ) => void;
  onRowResizeStart: (
    row: number,
    event: ReactPointerEvent<HTMLElement>,
  ) => void;
  onCommitColumnWidth: (col: number, width: number) => void;
  onCommitRowHeight: (row: number, height: number) => void;
  onBeginLineResize?: (axis: "row" | "col", sourceIndex: number) => PageLayoutLineTransaction | null;
  /** Commits the margins a ruler's handle moved (left/right or top/bottom). */
  onCommitHorizontalMargins: (margins: PrintMargins) => boolean;
  /** View > Show > Ruler (XLS-A030); the rulers are shown unless false. */
  showRuler?: boolean;
  lineResizeDisabledReason?: string;
  marginResizeDisabledReason?: string;
  headerFooterDisabledReason?: string;
  onCommitHeaderFooter: (field: HeaderFooterField, align: "left" | "center" | "right", text: string) => boolean;
}

export type HeaderFooterField = "header" | "footer" | "firstHeader" | "firstFooter" | "evenHeader" | "evenFooter";

export type SpreadsheetPageLayoutDrawing = readonly [
  string,
  readonly [number, number, number, number],
  string,
  ReactNode,
  { frame: PageLayoutDrawingFrame; rotation: number; selectionType?: "shape" | "chart"; groupId?: string; disabledReason?: string; constrainSize?: (width: number, height: number) => { width: number; height: number };
    /** Lock aspect ratio (a picture's by default, a shape's when locked): a handle drag keeps the proportions, as Shift does. */
    aspectLocked?: boolean }?,
];

export interface PageLayoutDrawingTransaction {
  isCurrent: () => boolean;
  commit: (frame: PageLayoutDrawingFrame) => boolean;
}

export interface PageLayoutLineTransaction {
  isCurrent: () => boolean;
  commit: (size: number) => boolean;
}

const PAPER_POINT_TO_CSS_PIXEL = PX_PER_PT;
const PAGE_LAYOUT_PAGE_TOP_CHROME = 32;
const PAGE_LAYOUT_MIN_PRINTABLE_WIDTH_PT = 72;

type PageLayoutMarginSide = "left" | "right" | "top" | "bottom";

const OPPOSITE_MARGIN: Record<PageLayoutMarginSide, PageLayoutMarginSide> = {
  left: "right",
  right: "left",
  top: "bottom",
  bottom: "top",
};

/** A margin dragged on a ruler, kept inside the paper with a printable band. */
function clampPageMargins(
  side: PageLayoutMarginSide,
  value: number,
  source: PrintMargins,
  paperSize: number,
): PrintMargins {
  if (value === source[side]) return source;
  const opposite = source[OPPOSITE_MARGIN[side]];
  const maximum = Math.max(
    0,
    paperSize - opposite - PAGE_LAYOUT_MIN_PRINTABLE_WIDTH_PT,
  );
  return {
    ...source,
    [side]: Math.max(0, Math.min(maximum, source[side] + Math.round(value - source[side]))),
  };
}

/**
 * XLS-A030: the rulers of Page Layout view count in the region's unit, as
 * Excel's do (inches for US English, centimetres elsewhere), from the margin;
 * the minor ticks are quarter inches or half centimetres.
 */
function pageLayoutRulerUnit(locale: string): { points: number; minor: number } {
  return officeLocaleUsesInches(locale)
    ? { points: 72, minor: 18 }
    : { points: 72 / 2.54, minor: 36 / 2.54 };
}

function safeFormattedDate(date: Date, locale: string): string {
  try {
    return new Intl.DateTimeFormat(locale, { dateStyle: "short" }).format(date);
  } catch {
    return date.toISOString().slice(0, 10);
  }
}

function safeFormattedTime(date: Date, locale: string): string {
  try {
    return new Intl.DateTimeFormat(locale, {
      hour: "2-digit",
      minute: "2-digit",
    }).format(date);
  } catch {
    return date.toISOString().slice(11, 16);
  }
}

function runStyle(
  run: PrintHeaderFooterRun,
  paperScale: number,
  documentScale: number,
): CSSProperties {
  return headerFooterRunCss(headerFooterRunStyle(run), {
    pointToPx: (points) => points * paperScale * documentScale,
  });
}

function HeaderFooterSegment({
  segment,
  align,
  paperScale,
  documentScale,
}: {
  segment: ResolvedHeaderFooterSegment | null;
  align: "left" | "center" | "right";
  paperScale: number;
  documentScale: number;
}) {
  // A section's picture prints at its own size on the paper; Scale with
  // Document scales it with the text.
  const pictureToPx = (points: number) => points * paperScale * documentScale;
  const pictureTestId = `spreadsheet-page-layout-hf-picture-${align}`;
  return (
    <span
      data-page-layout-segment={align}
      style={{
        minWidth: 0,
        flex: "1 1 0",
        // Spaces at either end are text the section is laid out by (B386):
        // break-spaces keeps a trailing run of them from hanging past the
        // right margin, so "Trail    " sits four spaces in, as Excel prints it.
        whiteSpace: "break-spaces",
        overflowWrap: "anywhere",
        textAlign: align,
        lineHeight: 1.15,
      }}
    >
      {segment?.runs?.length
        ? segment.runs.map((run, index) => (
            <span
              // Imported runs are positional and may legitimately repeat text.
              key={index}
              data-page-layout-run={index}
              style={runStyle(run, paperScale, documentScale)}
            >
              {withHeaderFooterPicture(run.text, segment.picture, pictureToPx, pictureTestId)}
            </span>
          ))
        : withHeaderFooterPicture(segment?.text ?? "", segment?.picture, pictureToPx, pictureTestId)}
    </span>
  );
}

function HeaderFooterBand({
  kind,
  value,
  context,
  paperScale,
  documentScale,
  style,
  onCommit,
  disabledReason,
  translate: t,
  placeholder,
  printMode = false,
  fontFamily,
}: {
  kind: "header" | "footer";
  value: PrintHeaderFooter | undefined;
  context: HeaderFooterContext;
  paperScale: number;
  documentScale: number;
  style: CSSProperties;
  onCommit: (align: "left" | "center" | "right", text: string) => boolean;
  disabledReason?: string;
  translate: TranslationFn;
  placeholder: string;
  printMode?: boolean;
  /** The workbook's Normal font, as its cells are drawn in. */
  fontFamily?: string;
}) {
  const [editing, setEditing] = useState<"left" | "center" | "right" | null>(null);
  const [draft, setDraft] = useState("");
  const [error, setError] = useState(false);
  const commitRef = useRef(onCommit);
  const cancelRef = useRef(false);
  const composingRef = useRef(false);
  const originalRef = useRef("");
  const finish = () => {
    if (!editing || cancelRef.current) return;
    if (draft === originalRef.current || commitRef.current(editing, draft)) {
      cancelRef.current = true;
      setEditing(null);
      setError(false);
    } else setError(true);
  };
  const segments = (["left", "center", "right"] as const).map((align) =>
    resolveHeaderFooterSegment(
      value?.[align],
      value?.[`${align}Runs` as const],
      context,
      value?.[`${align}Picture` as const],
    ),
  );
  if (!printMode && !segments.some((segment) => segment?.text))
    segments[1] = { text: placeholder };
  return (
    <div
      data-testid={`spreadsheet-page-layout-${kind}`}
      data-page-layout-band={kind}
      role="group"
      aria-label={t(kind === "header" ? "spreadsheet.helperCopy.pageLayout.header" : "spreadsheet.helperCopy.pageLayout.footer")}
      style={{
        position: "absolute",
        display: "flex",
        alignItems: kind === "header" ? "flex-start" : "flex-end",
        gap: 0,
        color: "var(--page-text, #111827)",
        // A section without a font of its own prints in the workbook's
        // Normal font, as the cells do: in a Chinese Office that is the
        // theme's Hans face for every character, its spaces and digits too
        // ("第 1 页" 37.7 pt wide in Excel; 31.1 with a Latin space, P15).
        fontFamily: fontFamily ?? "Aptos, Calibri, sans-serif",
        fontSize: 11 * paperScale * documentScale,
        border: 0,
        background: "transparent",
        ...style,
      }}
    >
      {segments.map((segment, index) => {
        const align = (["left", "center", "right"] as const)[index];
        // A formatted section edits like a plain one: the window keeps the
        // formatting of the characters the edit leaves alone.
        const reason = disabledReason;
        const label = t(`spreadsheet.helperCopy.pageLayout.segment.${kind}.${align}`);
        const cellStyle: CSSProperties = {
          boxSizing: "border-box", flex: "1 1 0", width: 0, minWidth: 0,
          minHeight: 28 * paperScale, height: 32 * paperScale,
          padding: printMode ? 0 : 4 * paperScale, margin: 0, border: printMode ? 0 : "1px solid #b8bec3",
          // Where Excel sets the words against the margin, in the text's own
          // size (its PDF: a header's glyphs start 0.2 em below the header
          // margin, a footer's end 0.07 em above the footer margin, at 11 pt
          // and at 5.5 pt alike, P03 / P06).
          ...(kind === "header"
            ? { paddingTop: printMode ? "0.2em" : `calc(${4 * paperScale}px + 0.2em)` }
            : { paddingBottom: printMode ? "0.07em" : `calc(${4 * paperScale}px + 0.07em)` }),
          borderRadius: 0, background: "transparent", boxShadow: "none",
          color: "inherit", font: "inherit", textAlign: align, transform: "none",
          // Excel sets a header's first line at the header margin and a
          // footer's last line on the footer margin, growing away from the
          // edge; a button centres its content, which put the words 7.5 pt
          // below the header margin and 11 pt above the footer's (P03).
          display: "flex", alignItems: kind === "header" ? "flex-start" : "flex-end",
        };
        return editing === align ? (
          <input key={align} autoFocus aria-label={label}
            data-testid={`spreadsheet-page-layout-${kind}-${align}-input`}
            value={draft} maxLength={MAX_HEADER_FOOTER_SECTION_TEXT} aria-invalid={error || undefined}
            title={error ? (t("spreadsheet.helperCopy.pageLayout.saveStateChanged")) : undefined}
            onChange={(event) => { setDraft(event.target.value); setError(false); }}
            onCompositionStart={() => { composingRef.current = true; }}
            onCompositionEnd={() => { composingRef.current = false; }}
            onBlur={finish}
            onKeyDown={(event) => {
              event.stopPropagation();
              if (event.nativeEvent.isComposing || composingRef.current || event.keyCode === 229) return;
              if (event.key === "Escape") { event.preventDefault(); cancelRef.current = true; setEditing(null); setError(false); }
              if (event.key === "Enter") { event.preventDefault(); finish(); }
            }}
            onPointerDown={(event) => event.stopPropagation()}
            style={{ ...cellStyle, background: "#fff", outline: `2px solid ${error ? "#b91c1c" : "#15803d"}`, outlineOffset: -2 }}
          />
        ) : (
          <button key={align} type="button" aria-label={label} disabled={Boolean(reason)} title={reason}
            data-testid={`spreadsheet-page-layout-${kind}-${align}`}
            onPointerDown={(event) => event.stopPropagation()}
            onClick={() => {
              if (reason) return;
              commitRef.current = onCommit;
              originalRef.current = value?.[align] ?? "";
              cancelRef.current = false;
              setDraft(originalRef.current); setError(false); setEditing(align);
            }}
            style={cellStyle}
          >
            <HeaderFooterSegment segment={segment} align={align} paperScale={paperScale} documentScale={documentScale} />
          </button>
        );
      })}
      {error ? <span role="alert" style={{ position: "absolute", top: "100%", color: "#b91c1c", fontSize: 12 }}>{t("spreadsheet.helperCopy.pageLayout.notSaved")}</span> : null}
    </div>
  );
}

/** A shown note as Excel prints it "As displayed on sheet": its default box
 * (144 x 79 px) beside the cell's top-right corner, pale yellow with a thin
 * black edge, the text in the note font (Tahoma 9), clipped to the box. */
const NOTE_WIDTH_PT = 108;
const NOTE_HEIGHT_PT = 59.25;
const NOTE_GAP_PT = 5.25;
const NOTE_RISE_PT = 12.75;
function PrintedNote({ text, left, top, scale, lineWidth }: {
  text: string; left: number; top: number; scale: number; lineWidth: number;
}) {
  return (
    <div
      data-testid="spreadsheet-page-layout-printed-note"
      style={{
        position: "absolute", left, top, width: NOTE_WIDTH_PT * scale, height: NOTE_HEIGHT_PT * scale,
        boxSizing: "border-box", border: `${lineWidth}px solid #000`, background: "#ffffe1",
        padding: `${2 * scale}px ${3 * scale}px`, overflow: "hidden", zIndex: 5,
        fontFamily: "Tahoma, \"Segoe UI\", Arial, sans-serif", fontSize: 9 * scale, lineHeight: 1.2,
        color: "#000", whiteSpace: "pre-wrap", overflowWrap: "anywhere",
      }}
    >
      {text}
    </div>
  );
}

/** Page Setup > Sheet > Row and column headings, printed: as Excel prints
 * them (its own PDF, measured), a frame around headings and cells, one line
 * between the headings and the cells, no lines between headings, each label
 * centred in the Normal font. Lines are 1 pt and lie inside the edge they
 * mark, as Excel's do. */
function PrintedHeadings({
  left, top, gridLeft, headingWidth, headingHeight, gridWidth, gridHeight, lineWidth, font, fontSize, columns, rows,
}: {
  left: number; top: number; gridLeft: number; headingWidth: number; headingHeight: number;
  gridWidth: number; gridHeight: number; lineWidth: number;
  font?: { family: string; sizePt: number }; fontSize: number;
  columns: readonly { key: number; label: string; start: number; size: number }[];
  rows: readonly { key: number; label: string; start: number; size: number }[];
}) {
  const numbersLeft = gridLeft === 0 ? gridWidth : 0;
  const label: CSSProperties = {
    position: "absolute", display: "flex", alignItems: "center", justifyContent: "center",
    overflow: "hidden", whiteSpace: "nowrap", lineHeight: 1,
  };
  return (
    <div
      aria-hidden="true"
      data-testid="spreadsheet-page-layout-printed-headings"
      style={{
        position: "absolute", left, top, width: gridWidth + headingWidth, height: gridHeight + headingHeight,
        fontFamily: font?.family ? `"${font.family}", Aptos, Calibri, sans-serif` : undefined,
        fontSize, color: "#000", pointerEvents: "none",
      }}
    >
      {/* Drawn as vector rectangles: a CSS border or box of 1 pt (1.333 px)
          is snapped to 1 px and printed 0.75 pt. The dividers lie on the
          grid's own edge, over its frame (gridlines on) or a cell's border:
          one 1 pt line, as Excel prints it (P05 and P09, measured between
          the row lines; beside the grid's edge they printed 2 pt). */}
      <svg aria-hidden="true" width={gridWidth + headingWidth} height={gridHeight + headingHeight}
        style={{ position: "absolute", left: 0, top: 0, overflow: "visible", pointerEvents: "none" }}>
        <path fill="#000" fillRule="evenodd" d={`M0 0H${gridWidth + headingWidth}V${gridHeight + headingHeight}H0Z M${lineWidth} ${lineWidth}V${gridHeight + headingHeight - lineWidth}H${gridWidth + headingWidth - lineWidth}V${lineWidth}Z`} />
        <rect fill="#000" x={gridLeft} y={headingHeight} width={gridWidth} height={lineWidth} />
        <rect fill="#000" x={gridLeft === 0 ? gridWidth - lineWidth : gridLeft} y={headingHeight} width={lineWidth} height={gridHeight} />
      </svg>
      {columns.map((column) => (
        <span key={column.key} style={{ ...label, left: gridLeft + column.start, top: 0, width: column.size, height: headingHeight }}>
          {column.label}
        </span>
      ))}
      {rows.map((row) => (
        <span key={row.key} style={{ ...label, left: numbersLeft, top: headingHeight + row.start, width: headingWidth, height: row.size }}>
          {row.label}
        </span>
      ))}
    </div>
  );
}

/** How far past its frame a drawing paints (half its outline, a small
 * shadow), in points before scaling. */
const DRAWING_OUTLINE_BLEED_PT = 6;

/** Page Setup > Sheet > Gridlines, printed: 1 pt lines (Excel's own PDF),
 * each inside the edge it marks (a cell's right and bottom, the frame's
 * left and top), under the cells so a fill hides them; Black and white
 * prints the inner ones dotted, 1 pt on and 1 pt off, in a solid frame.
 * Vector rectangles and dashes: a 1.333 px CSS border is snapped to 1 px. */
export function PrintedGridlines({ width, height, lineWidth, dotted, columns, rows, rightToLeft, gaps }: {
  width: number; height: number; lineWidth: number; dotted: boolean;
  /** Each printed column's box, px from the grid's left, in page order. */
  columns: readonly { left: number; right: number }[];
  /** Each printed row's box, px from the grid's top, in page order. */
  rows: readonly { top: number; bottom: number }[];
  rightToLeft: boolean;
  /** The inner edges no gridline is printed on: `v:<col>:<row>` the edge
   * after page column <col> in row <row>, `h:<row>:<col>` the edge under
   * page row <row> in column <col> (a merge's inside, a run's columns). */
  gaps: ReadonlySet<string>;
}) {
  const inner = (edge: number, extent: number) => edge > lineWidth && Math.abs(edge - extent) >= 0.01;
  // Consecutive pieces of one line are drawn as one.
  const runs = (pieces: readonly (readonly [number, number])[]) => {
    const out: [number, number][] = [];
    for (const [from, to] of [...pieces].sort((a, b) => a[0] - b[0])) {
      const last = out.at(-1);
      if (last && from <= last[1] + 0.01) last[1] = Math.max(last[1], to);
      else out.push([from, to]);
    }
    return out;
  };
  const vertical: { x: number; from: number; to: number }[] = [];
  const seenX = new Set<number>();
  columns.forEach((column, index) => {
    if (index === columns.length - 1) return;
    // The edge this column shares with the next one in page order.
    const x = rightToLeft ? column.left : column.right;
    const key = Math.round(x * 100);
    if (!inner(x, width) || seenX.has(key)) return;
    seenX.add(key);
    for (const [from, to] of runs(rows.flatMap((row, rowIndex) => gaps.has(`v:${index}:${rowIndex}`) ? [] : [[row.top, row.bottom] as const]))) {
      vertical.push({ x, from, to });
    }
  });
  const horizontal: { y: number; from: number; to: number }[] = [];
  const seenY = new Set<number>();
  rows.forEach((row, index) => {
    if (index === rows.length - 1) return;
    const y = row.bottom;
    const key = Math.round(y * 100);
    if (!inner(y, height) || seenY.has(key)) return;
    seenY.add(key);
    for (const [from, to] of runs(columns.flatMap((column, colIndex) => gaps.has(`h:${index}:${colIndex}`) ? [] : [[column.left, column.right] as const]))) {
      horizontal.push({ y, from, to });
    }
  });
  const half = lineWidth / 2;
  return (
    <svg aria-hidden="true" data-testid="spreadsheet-page-layout-printed-gridlines" width={width} height={height}
      style={{ position: "absolute", left: 0, top: 0, pointerEvents: "none" }}>
      <path style={{ fill: "var(--page-gridline, #000)" }} fillRule="evenodd"
        d={`M0 0H${width}V${height}H0Z M${lineWidth} ${lineWidth}V${height - lineWidth}H${width - lineWidth}V${lineWidth}Z`} />
      {dotted ? (
        <g style={{ stroke: "var(--page-gridline, #000)" }} strokeWidth={lineWidth} strokeDasharray={`${lineWidth} ${lineWidth}`}>
          {vertical.map(({ x, from, to }) => <line key={`x${x}:${from}`} x1={x - half} x2={x - half} y1={Math.max(lineWidth, from)} y2={Math.min(height - lineWidth, to)} />)}
          {horizontal.map(({ y, from, to }) => <line key={`y${y}:${from}`} x1={Math.max(lineWidth, from)} x2={Math.min(width - lineWidth, to)} y1={y - half} y2={y - half} />)}
        </g>
      ) : (
        <g style={{ fill: "var(--page-gridline, #000)" }}>
          {vertical.map(({ x, from, to }) => <rect key={`x${x}:${from}`} x={x - lineWidth} y={from} width={lineWidth} height={to - from} />)}
          {horizontal.map(({ y, from, to }) => <rect key={`y${y}:${from}`} x={from} y={y - lineWidth} width={to - from} height={lineWidth} />)}
        </g>
      )}
    </svg>
  );
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function GridCell({
  row,
  col,
  display,
  cellStyle,
  width,
  height,
  left,
  top,
  fontScale,
  nativeTextDpi,
  nativeGridDpi,
  showGridlines,
  active,
  selected,
  interactive,
  editor,
  occurrence,
  paintsDisplay,
  cellHandlers,
  inputProps,
  mergeContent,
  overlay,
  content,
  spillBox,
  paintsOutside = false,
  sharedLeftEdge = false,
  sharedTopEdge = false,
}: {
  row: number;
  col: number;
  display: string;
  cellStyle: CSSProperties | undefined;
  width: number;
  height: number;
  left: number;
  top: number;
  fontScale: number;
  nativeTextDpi?: number;
  nativeGridDpi?: number;
  showGridlines: boolean;
  active: boolean;
  selected: boolean;
  interactive: boolean;
  /**
   * The cell carries its editor (the active or edited cell, or one something
   * asked for by address). Every other interactive cell is a painted box: no
   * input, its presses handed to the editor they mount (cellHandlers).
   */
  editor: boolean;
  /** Which of the cell's appearances this is (a print title repeats a cell on every page). */
  occurrence: string;
  /** No content renderer: a cell at rest paints its display text itself. */
  paintsDisplay: boolean;
  cellHandlers: PageLayoutCellHandlers;
  overlay?: ReactNode;
  content?: ReactNode;
  /** The text run's box, relative to the cell, when its text spills (B048). */
  spillBox?: { left: number; width: number };
  /** Its painting leans past its box (an angled cell's band). */
  paintsOutside?: boolean;
  /** The neighbour drawn on its left / above states the shared edge itself
   * (its right / bottom border, inside its own box): this cell leaves its
   * own left / top out, so the edge is one line, as Excel draws it, not two
   * side by side (TG020: every bordered grid line printed doubled). */
  sharedLeftEdge?: boolean;
  sharedTopEdge?: boolean;
  mergeContent?: {
    key: string;
    left: number;
    top: number;
    width: number;
    height: number;
  };
  inputProps?: InputHTMLAttributes<HTMLInputElement> & {
    "data-cell-text"?: "layered";
    "data-editing"?: "true" | "false";
  };
}) {
  const secondaryMergeFragment = Boolean(mergeContent && !interactive);
  const resting = interactive && !editor;
  const scaledCellStyle =
    typeof cellStyle?.fontSize === "number" &&
    Number.isFinite(cellStyle.fontSize)
      ? { ...cellStyle, fontSize: cellStyle.fontSize * fontScale }
      : cellStyle;
  const {border,borderTop,borderRight,borderBottom,borderLeft,...textBoxStyle}=scaledCellStyle??{};
  const nativeInset=nativeTextDpi?96/nativeTextDpi*fontScale:0;
  // An entry with manual line breaks is edited in a text area: a single-line
  // input silently drops every line break (the normal grid does the same).
  const editorValue = inputProps?.value ?? display;
  const multilineEntry = typeof editorValue === "string" && editorValue.includes("\n");
  const Editor = (multilineEntry ? "textarea" : "input") as "input";
  const key = `${row}:${col}`;
  // The merge's whole box, from this fragment's corner.
  const mergeBox = mergeContent ? {
    left: mergeContent.left,
    top: mergeContent.top,
    width: `calc(100% + ${mergeContent.width - width}px)`,
    height: `calc(100% + ${mergeContent.height - height}px)`,
  } : null;
  // The cell's text painted without an editor (a merge's secondary page
  // fragment, or a cell at rest when there is no content renderer).
  const displayText = () => (
    <span
      aria-hidden="true"
      data-page-layout-merge-content={mergeContent?.key}
      data-page-layout-merge-origin={mergeContent?.key}
      data-page-layout-painted-content={mergeContent?.key ?? key}
      className="box-border block min-w-0 overflow-hidden border-0 bg-transparent px-1 py-0.5 font-[inherit] leading-[inherit] text-ellipsis"
      style={{
        // The origin fragment is a full-height input, whose line box is
        // vertically centred by the browser. Secondary page fragments are
        // read-only spans, so they must reproduce that same baseline
        // explicitly instead of painting the split value at the top.
        alignItems: "center",
        display: "flex",
        justifyContent:
          scaledCellStyle?.textAlign === "center"
            ? "center"
            : scaledCellStyle?.textAlign === "right" ||
                scaledCellStyle?.textAlign === "end"
              ? "flex-end"
              : "flex-start",
        position: "absolute",
        left: mergeBox?.left ?? 0,
        top: mergeBox?.top ?? 0,
        width: mergeBox?.width ?? "100%",
        height: mergeBox?.height ?? "100%",
        color: "inherit",
        font: "inherit",
        textAlign: scaledCellStyle?.textAlign,
        whiteSpace: scaledCellStyle?.whiteSpace ?? "nowrap",
        pointerEvents: "none",
      }}
    >
      {content ?? display}
    </span>
  );
  // The painted content's box: the cell, its spill run, or the whole merge.
  const paintedBox = (attributes: Record<string, string | undefined>) => (
    <span aria-hidden="true" data-page-layout-painted-content={mergeContent?.key ?? key} {...attributes}
      style={{position:"absolute",left:spillBox?.left ?? mergeBox?.left ?? 0,top:mergeBox?.top ?? 0,width:spillBox ? spillBox.width : mergeBox?.width ?? "100%",height:mergeBox?.height ?? "100%",pointerEvents:"none"}}>{content}</span>
  );
  // A cell at rest is its painting only. Content that fills the cell needs
  // no box of its own (it lays itself out over the cell); a spill run or a
  // merge does. The merge's box keeps the attributes its editor carried.
  const restingPaint = !resting
    ? null
    : paintsDisplay
      ? (display ? displayText() : null)
      : content == null
        ? null
        : spillBox || mergeBox
          ? paintedBox(mergeContent ? { "data-page-layout-merge-content": mergeContent.key, "data-page-layout-merge-origin": mergeContent.key } : {})
          : content;
  const nativeBorders = nativeGridDpi ? {
    borderTop: sharedTopEdge ? undefined : borderTop ?? border,
    borderLeft: sharedLeftEdge ? undefined : borderLeft ?? border,
    borderRight: borderRight ?? border ?? (showGridlines ? "1px var(--page-gridline-style, solid) var(--page-gridline, #d9dee7)" : undefined),
    borderBottom: borderBottom ?? border ?? (showGridlines ? "1px var(--page-gridline-style, solid) var(--page-gridline, #d9dee7)" : undefined),
  } : null;
  // A cell that states no border of its own wears only the gridlines on its
  // right and bottom: its box paints them (.spreadsheet-page-layout-gridline,
  // the same lines over the same content) instead of a border layer each.
  const gridlineOnly = Boolean(nativeGridDpi) && showGridlines &&
    (borderTop ?? border) === undefined && (borderLeft ?? border) === undefined &&
    (borderRight ?? border) === undefined && (borderBottom ?? border) === undefined;
  return (
    <div
      role="gridcell"
      aria-label={
        secondaryMergeFragment ? undefined : `${columnLabel(col)}${row + 1}`
      }
      aria-hidden={secondaryMergeFragment || undefined}
      data-page-layout-cell={key}
      data-page-layout-merge-fragment={mergeContent?.key}
      // The cell's address for a press or a drag over its box (a gridline,
      // the edge around its editor), and, at rest, for the lookup that gives
      // it its editor (spreadsheetCellLookup).
      data-range-cell={interactive ? key : undefined}
      data-page-layout-ghost={resting ? "" : undefined}
      data-page-layout-occurrence={interactive ? occurrence : undefined}
      className={`absolute box-border min-w-0 overflow-hidden p-0 font-[Aptos,Calibri,sans-serif] leading-[1.2]${gridlineOnly ? " spreadsheet-page-layout-gridline" : ""}`}
      style={{
        left,
        top,
        width,
        height,
        borderRight: !nativeGridDpi && showGridlines ? "1px var(--page-gridline-style, solid) var(--page-gridline, #d9dee7)" : undefined,
        borderBottom: !nativeGridDpi && showGridlines ? "1px var(--page-gridline-style, solid) var(--page-gridline, #d9dee7)" : undefined,
        color: "var(--page-text, #111827)",
        backgroundColor: "transparent",
        fontSize: 12 * fontScale,
        ...(nativeGridDpi?textBoxStyle:scaledCellStyle),
        outline: active
          ? "2px solid #217346"
          : selected
            ? "1px solid rgba(33, 115, 70, 0.72)"
            : undefined,
        outlineOffset: active ? -2 : -1,
        pointerEvents: secondaryMergeFragment ? "none" : undefined,
        // A cell at rest shows the editor's pointer over its box.
        cursor: resting ? "cell" : undefined,
        // A spilling cell paints its run over the neighbours it runs into.
        ...(spillBox || paintsOutside ? { overflow: "visible", zIndex: 1 } : {}),
      }}
      onPointerDown={resting ? cellHandlers.pointerDown : undefined}
      onDoubleClick={resting ? cellHandlers.doubleClick : undefined}
      onMouseDown={interactive ? cellHandlers.mouseDown : undefined}
    >
      {resting ? (
        nativeTextDpi && restingPaint ? (
          <div data-native-text-content-box="" style={{position:"absolute",inset:0,left:nativeInset}}>{restingPaint}</div>
        ) : restingPaint
      ) : (
      <div data-native-text-content-box={nativeTextDpi?"":undefined} style={nativeTextDpi?{position:"absolute",inset:0,left:nativeInset}:{display:"contents"}}>
      {secondaryMergeFragment ? displayText() : (
        <Editor
          {...inputProps}
          {...((multilineEntry ? { rows: 1 } : {}) as object)}
          value={editorValue}
          aria-label={`${columnLabel(col)}${row + 1}`}
          data-cell={interactive ? key : undefined}
          data-page-layout-merge-content={mergeContent?.key}
          data-page-layout-merge-origin={mergeContent?.key}
          tabIndex={active && interactive ? 0 : -1}
          className="box-border h-full w-full min-w-0 border-0 bg-transparent px-1 py-0.5 font-[inherit] leading-[inherit] text-ellipsis outline-0"
          style={{
            position: mergeBox ? "absolute" : undefined,
            left: mergeBox?.left,
            top: mergeBox?.top,
            width: mergeBox?.width,
            height: mergeBox?.height,
            color: "inherit",
            font: "inherit",
            textAlign: scaledCellStyle?.textAlign,
            whiteSpace: multilineEntry ? "pre-wrap" : scaledCellStyle?.whiteSpace ?? "nowrap",
            ...(multilineEntry ? { resize: "none" as const, overflow: "hidden" } : {}),
            cursor: inputProps?.["data-editing"] === "true" ? "text" : "cell",
            caretColor:
              inputProps?.["data-editing"] === "true"
                ? undefined
                : "transparent",
            ...inputProps?.style,
          }}
        />
      )}
      {!secondaryMergeFragment && content ? paintedBox({}) : null}
      </div>
      )}
      {overlay}
      {nativeBorders && !gridlineOnly && (nativeBorders.borderTop || nativeBorders.borderLeft || nativeBorders.borderRight || nativeBorders.borderBottom)
        ? <span aria-hidden="true" data-page-layout-cell-borders="" style={{position:"absolute",inset:0,pointerEvents:"none",...nativeBorders}}/>
        : null}
    </div>
  );
}

/** A cell at rest's handlers, one set shared by every cell of the view. */
interface PageLayoutCellHandlers {
  pointerDown: (event: ReactPointerEvent<HTMLElement>) => void;
  doubleClick: (event: ReactMouseEvent<HTMLElement>) => void;
  mouseDown: (event: ReactMouseEvent<HTMLElement>) => void;
}

/** A cell's own controls (a picker, a filter button, a check box, a link
 * mark): a press on them is theirs, not the cell's. */
const PAGE_LAYOUT_CELL_CONTROL = 'button, a, input, textarea, select, label, [role="button"], [role="checkbox"], [role="menuitem"]';

/** The press landed on the cell's own box or painting, not on one of its controls. */
function pressesPageLayoutCell(target: EventTarget | null, cell: HTMLElement): boolean {
  if (!(target instanceof Element) || !cell.contains(target)) return false;
  const control = target.closest(PAGE_LAYOUT_CELL_CONTROL);
  return !control || !cell.contains(control);
}

/** Past the page stack's entrance (seven 40 ms steps and 300 ms): a page
 * mounted later is a scroll fill. */
const PAGE_LAYOUT_ENTRANCE_SETTLE_MS = 600;

/**
 * One page tile of the grid. Pages on screen when the view opens rise in with
 * the Docs page stack's entrance (styles/office-motion.css, staggered by
 * their place in the grid); a page the scroll mounts later only fades up from
 * 55%, as the Docs stack's virtualized fills do. Which one a tile is, is fixed
 * when it mounts, so its animation never restarts.
 */
function PageLayoutPageTile({
  entranceSettled,
  style,
  children,
}: {
  entranceSettled: boolean;
  style: CSSProperties;
  children: ReactNode;
}) {
  const [filled] = useState(entranceSettled);
  return (
    <div
      className={`relative overflow-visible office-page-enter${
        filled ? " office-page-enter--filled" : ""
      }`}
      style={style}
    >
      {children}
    </div>
  );
}

/** The extent assumed for the first paint, before the viewport is measured. */
const PAGE_LAYOUT_FIRST_PAINT_VIEWPORT = { left: 0, top: 0, width: 1600, height: 1000 };

/**
 * Excel-style Page Layout canvas. Cell edits delegate to SpreadsheetWindow's
 * existing selection, validation, protection and Undo authority.
 */
export default function SpreadsheetPageLayoutView(pageLayoutProps: SpreadsheetPageLayoutViewProps) {
  const {
  sheetId,
  sheetName,
  rowCount,
  columnCount,
  printSettings,
  fileName,
  filePath,
  locale,
  translate: t,
  zoom,
  showGridlines,
  rowHeights,
  columnWidths,
  getTextSpill,
  defaultRowHeight,
  nativeRowHeights,
  nativeTextDpi,
  nativeDigitWidth,
  headingFont,
  printedNotes,
  hasNativeTextLayout,
  defaultColumnWidth,
  merges,
  drawings,
  backgroundPicture,
  rightToLeft = false,
  onSelectDrawing,
  selectedDrawingId,
  selectedDrawingIds,
  onClearDrawingSelection,
  onBeginDrawingTransform,
  onRotateDrawing,
  hiddenRows,
  hiddenCols,
  sourceRowAt = (row) => row,
  rowOrder,
  activeCell,
  editingCell,
  selection,
  viewportRef,
  browseBlankPages = false,
  printMode = false,
  printPageStart = 0,
  getCellDisplay,
  getCellStyle,
  renderCellOverlay,
  renderCellContent,
  cellPaintsOutside,
  getCellInputProps,
  onColumnResizeStart,
  onRowResizeStart,
  onCommitColumnWidth,
  onCommitRowHeight,
  onBeginLineResize,
  onCommitHorizontalMargins,
  lineResizeDisabledReason,
  marginResizeDisabledReason,
  showRuler = true,
  onCommitHeaderFooter,
  headerFooterDisabledReason,
} = pageLayoutProps;
  const [renderedAt] = useState(() => new Date());
  // A cell carries an editor only while it is the active or the edited cell,
  // or something asked for it by address (spreadsheetCellLookup: a focus, a
  // read of its entry, a press on it): these, as "row:col". Every other cell
  // is its painting alone, so a page mounts without 400 inputs. The set is
  // emptied once the active cell moves on.
  const [materializedCells, setMaterializedCells] = useState<ReadonlySet<string>>(() => new Set());
  const materializeCell = useCallback((key: string) => {
    // Mounted now, in the caller's turn: it focuses or reads the editor next.
    flushSync(() => setMaterializedCells((current) => (current.has(key) ? current : new Set(current).add(key))));
  }, []);
  useEffect(() => {
    setMaterializedCells((current) => (current.size ? new Set() : current));
  }, [activeCell.row, activeCell.col, editingCell?.row, editingCell?.col]);
  // The cell whose editor holds the focus keeps it until the focus moves on:
  // a key that moves the active cell renders before the next cell is
  // focused (or finds no next cell, one on a page not built yet), and the
  // editor taken out from under the focus left the keyboard nowhere.
  const [focusedCell, setFocusedCell] = useState<string | null>(null);
  const trackCellFocus = useCallback((event: ReactFocusEvent<HTMLElement>) => {
    const key = event.target.closest("[data-page-layout-cell]") ? event.target.getAttribute("data-cell") : null;
    if (key) setFocusedCell(key);
  }, []);
  const trackCellBlur = useCallback((event: ReactFocusEvent<HTMLElement>) => {
    if (!event.target.hasAttribute("data-cell")) return;
    const next = event.relatedTarget instanceof Element && event.currentTarget.contains(event.relatedTarget)
      ? event.relatedTarget.getAttribute("data-cell") : null;
    setFocusedCell(next);
  }, []);
  useLayoutEffect(() => registerSpreadsheetCellMaterializer(viewportRef.current, materializeCell));
  const latestPropsRef = useRef(pageLayoutProps);
  latestPropsRef.current = pageLayoutProps;
  // The cell a press on a cell at rest just focused: the mouse-down that
  // follows the press must not move the focus off it again.
  const restingPressFocusRef = useRef<HTMLElement | null>(null);
  const [cellHandlers] = useState<PageLayoutCellHandlers>(() => {
    // The editor a press on a cell at rest is handed to: mounted for the
    // press, in this appearance of the cell.
    const restingEditor = (cell: HTMLElement, target: EventTarget | null) => {
      if (!pressesPageLayoutCell(target, cell)) return null;
      const key = cell.getAttribute("data-range-cell");
      const occurrence = cell.getAttribute("data-page-layout-occurrence");
      if (!key || occurrence === null) return null;
      if (cell.hasAttribute("data-page-layout-ghost")) materializeCell(key);
      const editor = cell.querySelector<HTMLElement>(`[data-cell="${key}"]`);
      if (!editor) return null;
      const split = key.indexOf(":");
      const handlers = latestPropsRef.current.getCellInputProps(Number(key.slice(0, split)), Number(key.slice(split + 1)), occurrence);
      return { editor, handlers };
    };
    return {
      pointerDown: (event) => {
        const cell = event.currentTarget;
        const found = restingEditor(cell, event.target);
        if (!found) return;
        found.handlers.onPointerDown?.(retargetSpreadsheetCellEvent(event, found.editor) as unknown as ReactPointerEvent<HTMLInputElement>);
        // A press the editor's handler left to the browser (a right or
        // middle button) focuses the cell under it, as it focused the
        // input that used to lie there. A touch focuses on its tap.
        if (!event.defaultPrevented && event.pointerType !== "touch" && document.activeElement !== found.editor) {
          found.editor.focus({ preventScroll: true });
          restingPressFocusRef.current = cell;
          // The mouse-down comes in the same task; a press without one
          // (a pen that never sends it) must not hold the next one.
          window.setTimeout(() => {
            if (restingPressFocusRef.current === cell) restingPressFocusRef.current = null;
          }, 0);
        }
      },
      doubleClick: (event) => {
        const found = restingEditor(event.currentTarget, event.target);
        found?.handlers.onDoubleClick?.(retargetSpreadsheetCellEvent(event, found.editor) as unknown as ReactMouseEvent<HTMLInputElement>);
      },
      mouseDown: (event) => {
        if (restingPressFocusRef.current !== event.currentTarget) return;
        restingPressFocusRef.current = null;
        event.preventDefault();
      },
    };
  });
  const [keyboardResize, setKeyboardResize] = useState<{
    axis: "col" | "row";
    index: number;
    original: number;
    size: number;
  } | null>(null);
  const lineKeyboardReceiptRef = useRef<PageLayoutLineTransaction | null>(null);
  const linePointerRef = useRef<{
    axis: "row" | "col"; index: number; pointerId: number; position: number;
    original: number; size: number; ratio: number; transaction: PageLayoutLineTransaction;
  } | null>(null);
  const [linePointerActive, setLinePointerActive] = useState(false);
  const [marginPreview, setMarginPreview] = useState<PrintMargins | null>(null);
  const marginDragRef = useRef<{
    side: PageLayoutMarginSide;
    startX: number;
    startY: number;
    original: PrintMargins;
    /** The paper's extent along the dragged axis, in points. */
    paperWidth: number;
    pointerId: number;
    pixelsPerPoint: number;
    commit: (margins: PrintMargins) => boolean;
  } | null>(null);
  const marginKeyboardOriginalRef = useRef<PrintMargins | null>(null);
  const marginKeyboardCommitRef = useRef<((margins: PrintMargins) => boolean) | null>(null);
  const previewColumnWidths =
    keyboardResize?.axis === "col"
      ? { ...(columnWidths ?? {}), [keyboardResize.index]: keyboardResize.size }
      : columnWidths;
  const previewRowHeights =
    keyboardResize?.axis === "row"
      ? { ...(rowHeights ?? {}), [keyboardResize.index]: keyboardResize.size }
      : rowHeights;
  const contentGeometry = resolveSpreadsheetPageLayoutGeometry({
    drawingBounds: drawings.map((drawing) => drawing[1]),
    totalRows: rowCount,
    nativeGridDpi: nativeTextDpi,
    nativeDigitWidth,
    rowOrder,
    totalCols: columnCount,
    hiddenRows,
    hiddenCols,
    printSettings: marginPreview
      ? {
          ...(normalizePrintSettings(printSettings) ?? {}),
          margins: marginPreview,
        }
      : printSettings,
    rowHeights: previewRowHeights,
    columnWidths: previewColumnWidths,
    defaultRowHeight,
    defaultColumnWidth,
  });
  const [browseExtent, setBrowseExtent] = useState({ sheetId, rowPages: 2, colPages: 2 });
  const [viewport, setViewport] = useState({ left: 0, top: 0, width: 0, height: 0 });
  // A host that measures to nothing (a test document) will never report a
  // viewport, so it keeps the full budgeted plan of every page.
  const [viewportUnmeasurable, setViewportUnmeasurable] = useState(false);
  // Advances once a frame while pages new to the view wait to be built.
  const [pageBuildFrame, setPageBuildFrame] = useState(0);
  const browsing = browseBlankPages && !printMode && !contentGeometry.settings.printArea;
  const geometry = browsing ? extendPageLayoutForBrowsing(contentGeometry,
    browseExtent.sheetId === sheetId ? browseExtent : { rowPages: 2, colPages: 2 }) : contentGeometry;
  const authoringMargins = geometry.sourceMargins ?? geometry.margins;
  const paperScale = PAPER_POINT_TO_CSS_PIXEL * Math.max(0.1, Math.min(4, zoom));
  const tileWidth = geometry.paperWidth * paperScale;
  const tileHeight = geometry.paperHeight * paperScale + PAGE_LAYOUT_PAGE_TOP_CHROME;
  // What readViewport last stored. A scroll event that changes nothing calls
  // no setter at all: a setter handed an unchanged value still re-ran this
  // whole view (every page's cells) whenever React could not bail out of it
  // eagerly, which during a scroll was most events.
  const viewportStateRef = useRef({ viewport, unmeasurable: viewportUnmeasurable, browseExtent });
  viewportStateRef.current = { viewport, unmeasurable: viewportUnmeasurable, browseExtent };
  const readViewport = useCallback(() => {
    const element = viewportRef.current;
    if (!element) return;
    const next = { left: element.scrollLeft, top: element.scrollTop,
      width: element.clientWidth, height: element.clientHeight };
    const stored = viewportStateRef.current;
    const unmeasurable = next.width === 0 && next.height === 0;
    if (unmeasurable !== stored.unmeasurable) setViewportUnmeasurable(unmeasurable);
    // The viewport plan reads the scroll position only as the band of page
    // tiles in view (resolvePageLayoutViewportPlan: its default 80/40 offset
    // and 24px gap). Storing every scrolled pixel re-rendered every mounted
    // page and cell on each scroll event (Safari: seconds of jank per
    // flick); a scroll within the same band keeps the state as it is.
    const band = (start: number, extent: number, tile: number) =>
      `${Math.floor(start / (tile + 24))}:${Math.floor((start + extent) / (tile + 24))}`;
    const old = stored.viewport;
    if (!(old.width === next.width && old.height === next.height &&
      band(old.top - 40, old.height, tileHeight) === band(next.top - 40, next.height, tileHeight) &&
      band(old.left - 80, old.width, tileWidth) === band(next.left - 80, next.width, tileWidth))) {
      // A band change builds only the pages coming into view (the others
      // are kept, see the pages below), so it renders at once: as a
      // transition it was restarted by every later scroll event and a long
      // scroll built nothing until it paused. The pages' paper is drawn
      // ahead of them either way.
      setViewport(next);
    }
    if (!browsing) return;
    // Three pages past the view, so browsing on grows the sheet (a whole
    // re-render) once every few pages rather than at every page.
    const rowPages = Math.max(2, Math.ceil((next.top + next.height - 40) / (tileHeight + 24)) + 3);
    const colPages = Math.max(2, Math.ceil((next.left + next.width - 80) / (tileWidth + 24)) + 3);
    const extent = stored.browseExtent;
    if (extent.sheetId === sheetId && extent.rowPages >= rowPages && extent.colPages >= colPages) return;
    setBrowseExtent((current) => current.sheetId === sheetId && current.rowPages >= rowPages && current.colPages >= colPages
      ? current : { sheetId, rowPages: Math.max(current.sheetId === sheetId ? current.rowPages : 2, rowPages),
        colPages: Math.max(current.sheetId === sheetId ? current.colPages : 2, colPages) });
  }, [browsing, sheetId, tileHeight, tileWidth, viewportRef]);
  useLayoutEffect(() => {
    const element = viewportRef.current;
    readViewport();
    if (!element || typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(readViewport);
    observer.observe(element);
    return () => observer.disconnect();
  }, [readViewport, browsing, viewportRef]);
  // Pages mount around the viewport in every mode. A first open of a large
  // sheet used to mount every page (up to the cell budget) before the first
  // paint; until the viewport has been measured, a typical window's extent
  // stands in for it and the layout effect re-plans from the real one.
  // One page of overscan applies before measurement too: an unmeasured host
  // (a test document, or a paper moved past the assumed extent before the
  // layout effect runs) still finds its neighbouring pages mounted.
  // A host without ResizeObserver never measures its viewport (test documents
  // in particular), so it keeps the full budgeted plan of every page, exactly
  // as before the viewport plan served every mode.
  const printBatch = printMode ? resolveSpreadsheetPrintBatchPlan(geometry, printPageStart) : null;
  const renderPlan = printBatch ?? (typeof ResizeObserver === "undefined" || viewportUnmeasurable
    ? resolveSpreadsheetPageLayoutRenderPlan(geometry)
    : resolvePageLayoutViewportPlan(geometry, {
        ...(viewport.width > 0 && viewport.height > 0 ? viewport : PAGE_LAYOUT_FIRST_PAINT_VIEWPORT),
        paperScale, tileWidth, tileHeight,
      }));
  const sourceRowSizes = Array.from({ length: rowCount }, () => 0);
  geometry.rowSizes.forEach((size, visual) => { sourceRowSizes[sourceRowAt(visual)] = size; });
  const sourceRowOffsets = [0];
  for (const size of sourceRowSizes) sourceRowOffsets.push(sourceRowOffsets[sourceRowOffsets.length - 1] + size);
  const sourceRowHeight = (start: number, end: number) => (sourceRowOffsets[end] ?? sourceRowOffsets.at(-1)!) - (sourceRowOffsets[start] ?? 0);
  const pages = renderPlan.pages;
  const [entranceSettledSheetId, setEntranceSettledSheetId] = useState<string | null>(null);
  useEffect(() => {
    const timer = window.setTimeout(
      () => setEntranceSettledSheetId(sheetId),
      PAGE_LAYOUT_ENTRANCE_SETTLE_MS,
    );
    return () => window.clearTimeout(timer);
  }, [sheetId]);
  const entranceSettled = printMode || entranceSettledSheetId === sheetId;
  // Printed, the gridlines are Page Setup's (Sheet > Gridlines), not the view's.
  const pageGridlines = printMode ? printSettings?.showGridlines === true : showGridlines;
  const headerFooterFontFamily = headingFont?.family ? `"${headingFont.family}", Aptos, Calibri, sans-serif` : undefined;
  const renderedPageCount = pages.length;
  const documentScale = resolveHeaderFooterDocumentScale(
    geometry.settings,
    geometry.contentScale,
  );
  const alignWithMargins =
    geometry.settings.headerFooterAlignWithMargins !== false;
  const date = safeFormattedDate(renderedAt, locale);
  const time = safeFormattedTime(renderedAt, locale);
  const [drawingPreview, setDrawingPreview] = useState<{ id: string; frame: PageLayoutDrawingFrame } | null>(null);
  const selectedGroup = drawings.find(drawing => drawing[0] === selectedDrawingId)?.[4]?.groupId;
  const groupDrawings=drawings.flatMap(drawing=>drawing[4]?[{id:drawing[0],frame:drawing[4].frame,groupId:drawing[4].groupId}]:[]);
  const groupPreview = new Map(drawingPreview
    ? ((selectedDrawingIds?.length ?? 0)>1 && selectedDrawingIds?.includes(drawingPreview.id)
      ? transformSpreadsheetDrawingSelection(groupDrawings,selectedDrawingIds,drawingPreview.id,drawingPreview.frame)
      : selectedGroup ? transformChartDrawingGroup(groupDrawings,drawingPreview.id,drawingPreview.frame) : [])?.map(drawing=>[drawing.id,drawing.frame]) ?? [] : []);
  const [drawingError, setDrawingError] = useState(false);
  const drawingClickMoved = useRef(false);
  const drawingDrag = useRef<{
    id: string; pointerId: number; x: number; y: number; clientX: number; clientY: number;
    startClientX: number; startClientY: number; frame: PageLayoutDrawingFrame; rotation: number;
    handle?: PageLayoutDrawingHandle; transaction: PageLayoutDrawingTransaction;
    preview: PageLayoutDrawingFrame; moved: boolean;
    constrainSize?: (width: number, height: number) => { width: number; height: number };
    aspectLocked?: boolean; shiftKey?: boolean;
  } | null>(null);
  // Keyed by the page's tile, which stays put while browsing renumbers the pages.
  const drawingGridRefs = useRef(new Map<string, HTMLDivElement>());
  const pageElementCacheRef = useRef(new Map<string, { deps: readonly unknown[]; element: ReactElement }>());
  const drawingBandsForPage = (page: (typeof pages)[number]) => {
    const overThenDown = geometry.settings.pageOrder === "overThenDown";
    const rowIndex = overThenDown ? Math.floor(page.index / pageColumnCount) : page.index % pageRowCount;
    const colIndex = overThenDown ? page.index % pageColumnCount : Math.floor(page.index / pageRowCount);
    const width = page.cols.reduce((sum, col) => sum + (geometry.colSizes[col] ?? geometry.colWidth), 0);
    const height = page.rows.reduce((sum, row) => sum + (geometry.rowSizes[row] ?? geometry.rowHeight), 0);
    return {
      cols: pageLayoutDrawingBands(geometry.colSizes, page.cols,
        geometry.dataAreaWidth - (geometry.settings.centerHorizontally ? Math.max(0, (geometry.dataAreaWidth - width) / 2) : 0),
        !geometry.settings.printArea && colIndex === pageColumnCount - 1),
      rows: pageLayoutDrawingBands(geometry.rowSizes, page.rows,
        geometry.dataAreaHeight - (geometry.settings.centerVertically ? Math.max(0, (geometry.dataAreaHeight - height) / 2) : 0),
        !geometry.settings.printArea && rowIndex === pageRowCount - 1),
    };
  };
  const sourcePoint = (clientX: number, clientY: number) => {
    let closest: { distance: number; x: number; y: number } | null = null;
    for (const page of pages) {
      const gridTile = resolveSpreadsheetPageLayoutTilePosition(geometry, page.index);
      const rect = gridTile ? drawingGridRefs.current.get(`${gridTile.row}:${gridTile.col}`)?.getBoundingClientRect() : undefined;
      if (!rect || !rect.width || !rect.height) continue;
      const {cols, rows} = drawingBandsForPage(page);
      const extent = (bands: typeof cols) => { const last = bands.at(-1); return last ? last.offset + last.end - last.start : 0; };
      // The notebook canvas can scale independently of this sheet's zoom.
      // Derive each physical ratio from the unchanged cell grid, not from
      // the expanded drawing bands (which include otherwise unused paper).
      const gridWidthPoints = page.cols.reduce((sum, col) => sum + (geometry.colSizes[col] ?? geometry.colWidth), 0);
      const gridHeightPoints = page.rows.reduce((sum, row) => sum + (geometry.rowSizes[row] ?? geometry.rowHeight), 0);
      const xPixelsPerPoint = rect.width / gridWidthPoints;
      const yPixelsPerPoint = rect.height / gridHeightPoints;
      if (!(xPixelsPerPoint > 0) || !(yPixelsPerPoint > 0)) continue;
      const x = Math.max(rect.left, Math.min(clientX, rect.left + extent(cols) * xPixelsPerPoint));
      const y = Math.max(rect.top, Math.min(clientY, rect.top + extent(rows) * yPixelsPerPoint));
      const distance = (x - clientX) ** 2 + (y - clientY) ** 2;
      if (closest && closest.distance <= distance) continue;
      const scale = pxToPt(geometry.contentScale);
      closest = { distance,
        x: pageLayoutAxisSourceAt((x - rect.left) / xPixelsPerPoint, cols) / scale,
        y: pageLayoutAxisSourceAt((y - rect.top) / yPixelsPerPoint, rows) / scale };
    }
    return closest;
  };
  const sourcePointRef = useRef(sourcePoint);
  useLayoutEffect(() => { sourcePointRef.current = sourcePoint; });
  const beginDrawing = (event: ReactPointerEvent<HTMLElement>, drawing: SpreadsheetPageLayoutDrawing, handle?: PageLayoutDrawingHandle) => {
    if (!handle && (event.shiftKey||event.ctrlKey||event.metaKey)) {event.preventDefault();return;}
    onSelectDrawing?.(drawing[0],false,true);
    if (event.button !== 0 || drawing[4]?.disabledReason || !drawing[4]) return;
    if (!handle && (event.target as HTMLElement).closest("button,input,textarea,select,[contenteditable='true']")) return;
    const point = sourcePoint(event.clientX, event.clientY);
    const transaction = onBeginDrawingTransform?.(drawing[0]);
    if (!point || !transaction) return;
    event.stopPropagation();
    event.preventDefault();
    event.currentTarget.focus();
    const { frame, rotation } = drawing[4];
    drawingClickMoved.current=false;
    drawingDrag.current = { id: drawing[0], pointerId: event.pointerId, x: point.x, y: point.y,
      clientX: event.clientX, clientY: event.clientY, startClientX: event.clientX, startClientY: event.clientY,
      frame, rotation, handle, transaction, preview: frame, moved: false, constrainSize: drawing[4].constrainSize,
      aspectLocked: drawing[4].aspectLocked, shiftKey: event.shiftKey };
    setDrawingError(false);
    setDrawingPreview({ id: drawing[0], frame });
  };
  const drawingDragging = drawingPreview !== null;
  useEffect(() => {
    if (!drawingDragging) return;
    const cancel = () => { drawingDrag.current = null; setDrawingPreview(null); };
    const update = (clientX: number, clientY: number) => {
      const drag = drawingDrag.current;
      if (!drag) return;
      if(clientX!==drag.startClientX||clientY!==drag.startClientY)drawingClickMoved.current=true;
      if (!drag.transaction.isCurrent()) { cancel(); return; }
      drag.clientX = clientX; drag.clientY = clientY;
      drag.moved ||= Math.hypot(clientX - drag.startClientX, clientY - drag.startClientY) >= 3;
      if (!drag.moved) return;
      const point = sourcePointRef.current(clientX, clientY);
      if (!point) return;
      // Lock aspect ratio, or Shift while dragging (Excel's Normal view rule).
      const aspect = drag.handle && (drag.aspectLocked || drag.shiftKey) && drag.frame.height > 0 ? drag.frame.width / drag.frame.height : null;
      drag.preview = transformPageLayoutDrawing(drag.frame, drag.rotation, drag.handle, point.x - drag.x, point.y - drag.y, drag.constrainSize, aspect);
      setDrawingPreview({ id: drag.id, frame: drag.preview });
    };
    const move = (event: PointerEvent) => {
      if (event.pointerId !== drawingDrag.current?.pointerId) return;
      if (drawingDrag.current) drawingDrag.current.shiftKey = event.shiftKey;
      event.preventDefault(); update(event.clientX, event.clientY);
    };
    const finish = (event: PointerEvent) => {
      if (event.pointerId !== drawingDrag.current?.pointerId) return;
      update(event.clientX, event.clientY);
      const drag = drawingDrag.current;
      cancel();
      if (drag?.moved && !drag.transaction.commit(drag.preview)) setDrawingError(true);
    };
    const pointerCancel = (event: PointerEvent) => { if (event.pointerId === drawingDrag.current?.pointerId) cancel(); };
    const key = (event: KeyboardEvent) => { if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); cancel(); } };
    let raf = 0;
    const scroll = () => {
      const drag = drawingDrag.current, viewport = viewportRef.current;
      if (!drag) return;
      if (viewport) {
        const rect = viewport.getBoundingClientRect();
        const left = viewport.scrollLeft, top = viewport.scrollTop;
        if (rect.width > 0 && rect.height > 28) {
          viewport.scrollLeft += pageLayoutEdgeScrollStep(drag.clientX, rect.left, rect.right);
          viewport.scrollTop += pageLayoutEdgeScrollStep(drag.clientY, rect.top + 28, rect.bottom);
        }
        if (left !== viewport.scrollLeft || top !== viewport.scrollTop) update(drag.clientX, drag.clientY);
      }
      raf = requestAnimationFrame(scroll);
    };
    raf = requestAnimationFrame(scroll);
    window.addEventListener("pointermove", move, { passive: false });
    window.addEventListener("pointerup", finish);
    window.addEventListener("pointercancel", pointerCancel);
    window.addEventListener("keydown", key, true);
    window.addEventListener("blur", cancel);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", finish);
      window.removeEventListener("pointercancel", pointerCancel);
      window.removeEventListener("keydown", key, true);
      window.removeEventListener("blur", cancel);
    };
  // The gesture owns its immutable receipt; preview updates must not reinstall listeners.
  }, [drawingDragging, viewportRef]);
  const drawingGeometryKey = JSON.stringify([sheetId, zoom, printSettings, rowHeights, columnWidths, defaultRowHeight, defaultColumnWidth, [...(hiddenRows ?? [])], [...(hiddenCols ?? [])], rowOrder]);
  useEffect(() => {
    drawingDrag.current = null;
    // Changing the coordinate system invalidates the in-flight user gesture.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setDrawingPreview(null);
    setKeyboardResize(null);
    linePointerRef.current = null;
    lineKeyboardReceiptRef.current = null;
    setLinePointerActive(false);
    marginDragRef.current = null;
    marginKeyboardOriginalRef.current = null;
    marginKeyboardCommitRef.current = null;
    setMarginPreview(null);
  }, [drawingGeometryKey]);
  const pageLabel = (pageNumber: number) =>
    t("spreadsheet.helperCopy.pageLayout.pageCount", { pageNumber, totalPages: contentGeometry.totalPages });
  const lineSize = (axis: "col" | "row", index: number) => {
    if (keyboardResize?.axis === axis && keyboardResize.index === index)
      return keyboardResize.size;
    if (axis === "col")
      return resolvedColumnWidth(
        index,
        columnWidths,
        clampColumnWidth(
          defaultColumnWidth,
          DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
        ),
      );
    if (nativeRowHeights) return nativeRowCssClamp(rowHeights?.[index],defaultRowHeight ?? DEFAULT_SPREADSHEET_ROW_HEIGHT_PX);
    return resolvedRowHeight(
      index,
      rowHeights,
      clampRowHeight(defaultRowHeight, DEFAULT_SPREADSHEET_ROW_HEIGHT_PX),
    );
  };
  const onResizeKeyDown = (
    axis: "col" | "row",
    index: number,
    event: ReactKeyboardEvent<HTMLElement>,
  ) => {
    if (lineResizeDisabledReason) return;
    if (event.key === "Escape") {
      if (keyboardResize?.axis === axis && keyboardResize.index === index) {
        event.preventDefault();
        setKeyboardResize(null);
        lineKeyboardReceiptRef.current = null;
      }
      return;
    }
    if (event.key === "Enter") {
      if (keyboardResize?.axis === axis && keyboardResize.index === index) {
        event.preventDefault();
        if (keyboardResize.size !== keyboardResize.original) {
          const receipt = lineKeyboardReceiptRef.current;
          if (receipt) { if (!receipt.isCurrent() || !receipt.commit(keyboardResize.size)) setDrawingError(true); }
          else if (onBeginLineResize) setDrawingError(true);
          else if (axis === "col") onCommitColumnWidth(index, keyboardResize.size);
          else onCommitRowHeight(sourceRowAt(index), keyboardResize.size);
        }
        lineKeyboardReceiptRef.current = null;
        setKeyboardResize(null);
      }
      return;
    }
    // The arrows move the column's divider the way they point, and that
    // divider is the column's left edge on a right-to-left sheet.
    const negative = axis === "col" ? (rightToLeft ? "ArrowRight" : "ArrowLeft") : "ArrowUp";
    const positive = axis === "col" ? (rightToLeft ? "ArrowLeft" : "ArrowRight") : "ArrowDown";
    if (event.key !== negative && event.key !== positive) return;
    event.preventDefault();
    if (keyboardResize?.axis !== axis || keyboardResize.index !== index) {
      lineKeyboardReceiptRef.current = onBeginLineResize?.(axis, axis === "row" ? sourceRowAt(index) : index) ?? null;
      if (onBeginLineResize && !lineKeyboardReceiptRef.current) return;
    }
    const current = lineSize(axis, index);
    const original =
      keyboardResize?.axis === axis && keyboardResize.index === index
        ? keyboardResize.original
        : current;
    const delta = (event.shiftKey ? 10 : 1) * (event.key === positive ? 1 : -1);
    const size =
      axis === "col"
        ? clampColumnWidth(current + delta, current)
        : (nativeRowHeights ? clampNativeRowHeight : clampRowHeight)(current + delta, current);
    setKeyboardResize({ axis, index, original, size });
  };
  const startLineResize = (axis: "row" | "col", index: number, event: ReactPointerEvent<HTMLElement>) => {
    if (lineResizeDisabledReason || event.button !== 0) return;
    if (!onBeginLineResize) { if (axis === "row") onRowResizeStart(sourceRowAt(index), event); else onColumnResizeStart(index, event); return; }
    const transaction = onBeginLineResize(axis, axis === "row" ? sourceRowAt(index) : index);
    if (!transaction) return;
    event.preventDefault(); event.stopPropagation();
    const original = lineSize(axis, index);
    const rect = event.currentTarget.parentElement?.getBoundingClientRect();
    const pixels = axis === "row" ? rect?.height : rect?.width;
    // A right-to-left column's divider is its left edge: dragging it left widens.
    const grow = axis === "col" && rightToLeft ? -1 : 1;
    linePointerRef.current = { axis, index, pointerId: event.pointerId, position: axis === "row" ? event.clientY : event.clientX,
      original, size: original, ratio: grow * (pixels ? original / pixels : 1 / (zoom * geometry.contentScale)), transaction };
    setKeyboardResize({ axis, index, original, size: original });
    setLinePointerActive(true);
  };
  useEffect(() => {
    if (!linePointerActive) return;
    const cancel = () => { linePointerRef.current = null; setLinePointerActive(false); setKeyboardResize(null); };
    const update = (event: PointerEvent) => {
      const drag = linePointerRef.current;
      if (!drag || event.pointerId !== drag.pointerId) return;
      if (!drag.transaction.isCurrent()) { cancel(); return; }
      const point = drag.axis === "row" ? event.clientY : event.clientX;
      drag.size = point === drag.position ? drag.original : (drag.axis === "row" ? (nativeRowHeights ? clampNativeRowHeight : clampRowHeight) : clampColumnWidth)(drag.original + (point - drag.position) * drag.ratio, drag.original);
      setKeyboardResize({ axis: drag.axis, index: drag.index, original: drag.original, size: drag.size });
    };
    const finish = (event: PointerEvent) => {
      if (event.pointerId !== linePointerRef.current?.pointerId) return;
      update(event);
      const drag = linePointerRef.current;
      cancel();
      if (drag && drag.size !== drag.original && !drag.transaction.commit(drag.size)) setDrawingError(true);
    };
    const pointerCancel = (event: PointerEvent) => { if (event.pointerId === linePointerRef.current?.pointerId) cancel(); };
    const key = (event: KeyboardEvent) => { if (event.key === "Escape") { event.preventDefault(); cancel(); } };
    window.addEventListener("pointermove", update);
    window.addEventListener("pointerup", finish);
    window.addEventListener("pointercancel", pointerCancel);
    window.addEventListener("keydown", key);
    window.addEventListener("blur", cancel);
    return () => { window.removeEventListener("pointermove", update); window.removeEventListener("pointerup", finish); window.removeEventListener("pointercancel", pointerCancel); window.removeEventListener("keydown", key); window.removeEventListener("blur", cancel); };
  }, [linePointerActive,nativeRowHeights]);
  const commitMarginPreview = useCallback((latest?: PrintMargins) => {
    const preview = latest ?? marginPreview;
    const drag = marginDragRef.current;
    const original =
      drag?.original ?? marginKeyboardOriginalRef.current ?? authoringMargins;
    const commit = drag?.commit ?? marginKeyboardCommitRef.current ?? onCommitHorizontalMargins;
    marginDragRef.current = null;
    marginKeyboardOriginalRef.current = null;
    marginKeyboardCommitRef.current = null;
    setMarginPreview(null);
    if (
      preview &&
      (preview.left !== original.left || preview.right !== original.right)
    ) {
      if (!commit(preview)) setDrawingError(true);
    }
  }, [authoringMargins, marginPreview, onCommitHorizontalMargins]);
  const cancelMarginPreview = useCallback(() => {
    marginDragRef.current = null;
    marginKeyboardOriginalRef.current = null;
    marginKeyboardCommitRef.current = null;
    setMarginPreview(null);
  }, []);
  useEffect(() => {
    if (!marginDragRef.current) return;
    const position = (event: PointerEvent) => {
      const drag = marginDragRef.current;
      if (!drag || event.pointerId !== drag.pointerId) return;
      const vertical = drag.side === "top" || drag.side === "bottom";
      const deltaPoints =
        (vertical ? event.clientY - drag.startY : event.clientX - drag.startX) /
        drag.pixelsPerPoint;
      // Left and top grow as the handle moves away from the paper's start.
      const value =
        drag.side === "left" || drag.side === "top"
          ? drag.original[drag.side] + deltaPoints
          : drag.original[drag.side] - deltaPoints;
      const next = clampPageMargins(
        drag.side,
        value,
        drag.original,
        drag.paperWidth,
      );
      return next;
    };
    const onMove = (event: PointerEvent) => { const next = position(event); if (next) setMarginPreview(next); };
    const onUp = (event: PointerEvent) => { const next = position(event); if (next) commitMarginPreview(next); };
    const onCancel = (event: PointerEvent) => { if (event.pointerId === marginDragRef.current?.pointerId) cancelMarginPreview(); };
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      cancelMarginPreview();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onCancel);
    window.addEventListener("blur", cancelMarginPreview);
    window.addEventListener("keydown", onKey);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onCancel);
      window.removeEventListener("blur", cancelMarginPreview);
      window.removeEventListener("keydown", onKey);
    };
  }, [cancelMarginPreview, commitMarginPreview, paperScale]);
  const startMarginResize = (
    side: PageLayoutMarginSide,
    event: ReactPointerEvent<HTMLElement>,
  ) => {
    if (marginResizeDisabledReason) return;
    event.preventDefault();
    event.currentTarget.setPointerCapture?.(event.pointerId);
    const original = { ...authoringMargins };
    const vertical = side === "top" || side === "bottom";
    const paperSize = vertical ? geometry.paperHeight : geometry.paperWidth;
    const ruler = event.currentTarget.parentElement?.getBoundingClientRect();
    marginKeyboardOriginalRef.current = null;
    marginDragRef.current = {
      side,
      startX: event.clientX,
      startY: event.clientY,
      original,
      paperWidth: paperSize,
      pointerId: event.pointerId,
      pixelsPerPoint: ((vertical ? ruler?.height : ruler?.width) || paperSize * paperScale) / paperSize,
      commit: onCommitHorizontalMargins,
    };
    setMarginPreview(original);
  };
  const onMarginKeyDown = (
    side: PageLayoutMarginSide,
    event: ReactKeyboardEvent<HTMLElement>,
  ) => {
    if (marginResizeDisabledReason) return;
    if (event.key === "Escape") {
      event.preventDefault();
      cancelMarginPreview();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      commitMarginPreview();
      return;
    }
    const vertical = side === "top" || side === "bottom";
    const forward = vertical ? "ArrowDown" : "ArrowRight";
    const backward = vertical ? "ArrowUp" : "ArrowLeft";
    if (event.key !== forward && event.key !== backward) return;
    event.preventDefault();
    const source = marginPreview ?? authoringMargins;
    if (!marginPreview) { marginKeyboardOriginalRef.current = authoringMargins; marginKeyboardCommitRef.current = onCommitHorizontalMargins; }
    const delta =
      (event.shiftKey ? 10 : 1) * (event.key === forward ? 1 : -1);
    const value =
      side === "left" || side === "top" ? source[side] + delta : source[side] - delta;
    setMarginPreview(
      clampPageMargins(
        side,
        value,
        source,
        vertical ? geometry.paperHeight : geometry.paperWidth,
      ),
    );
  };
  // XLS-A030: the rulers count from the margin in the region's unit.
  const rulerUnit = pageLayoutRulerUnit(locale);
  const rulerMarks = (extent: number) =>
    Array.from(
      { length: Math.max(0, Math.floor(extent / rulerUnit.points + 1e-6)) },
      (_, index) => index + 1,
    );
  const rulerTickPx = Math.max(2, rulerUnit.minor * paperScale);
  const displayedMargins = geometry.margins;
  // Paper tiles: one grid for the sheet or its print area, and one grid per
  // area stacked below the previous for a multi-area print area.
  const { rows: pageRowCount, cols: pageColumnCount } = resolveSpreadsheetPageLayoutTileGrid(geometry);
  // The tiles whose page is not built: all of them on a sheet of up to 1,200
  // pages, else those within eight pages of the view.
  // Each page's element from the last render, by tile (see the pages below).
  const previousPageElements = pageElementCacheRef.current;
  const nextPageElements = new Map<string, { deps: readonly unknown[]; element: ReactElement }>();
  pageElementCacheRef.current = nextPageElements;
  // Pages new to the view are built one per frame, nearest the view first
  // (the plan lists them so): a band of three new pages built in one render
  // held the scroll for its whole length. The rest show their paper until
  // their frame. The first build of the view, and print, take every page.
  const builtPages = printMode || previousPageElements.size === 0
    ? pages
    : (() => {
        let newPagesLeft = 1;
        return pages.filter((page) => {
          const tile = resolveSpreadsheetPageLayoutTilePosition(geometry, page.index) ?? { row: 0, col: 0 };
          if (previousPageElements.has(`${tile.row}:${tile.col}`)) return true;
          if (newPagesLeft <= 0) return false;
          newPagesLeft -= 1;
          return true;
        });
      })();
  const deferredPageCount = pages.length - builtPages.length;
  useEffect(() => {
    if (deferredPageCount === 0) return;
    const frame = requestAnimationFrame(() => setPageBuildFrame((value) => value + 1));
    return () => cancelAnimationFrame(frame);
  }, [deferredPageCount, pageBuildFrame]);
  const pagePaperPlaceholders = (built: readonly { index: number }[]) => {
    const mounted = new Set(built.map((page) => {
      const tile = resolveSpreadsheetPageLayoutTilePosition(geometry, page.index);
      return tile ? `${tile.row}:${tile.col}` : "";
    }));
    const near = pageRowCount * pageColumnCount > 1200;
    const rowStep = tileHeight + 24, colStep = tileWidth + 24;
    const firstRow = near ? Math.max(0, Math.floor((viewport.top - 40) / rowStep) - 8) : 0;
    const lastRow = near ? Math.min(pageRowCount - 1, Math.floor((viewport.top + viewport.height) / rowStep) + 8) : pageRowCount - 1;
    const firstCol = near ? Math.max(0, Math.floor((viewport.left - 80) / colStep) - 8) : 0;
    const lastCol = near ? Math.min(pageColumnCount - 1, Math.floor((viewport.left + viewport.width) / colStep) + 8) : pageColumnCount - 1;
    const out: { row: number; col: number }[] = [];
    for (let row = firstRow; row <= lastRow; row++) {
      for (let col = firstCol; col <= lastCol; col++) {
        if (!mounted.has(`${row}:${col}`)) out.push({ row, col });
      }
    }
    return out;
  };
  const lineTotal = (sizes: readonly number[], start: number, end: number) =>
    sizes.slice(start, end).reduce((sum, size) => sum + size, 0);
  const mergeAt = (row: number, col: number) =>
    merges.find(
      (merge) =>
        row >= merge.row &&
        row < merge.row + merge.rowSpan &&
        col >= merge.col &&
        col < merge.col + merge.colSpan,
    );

  return (
    <div
      ref={viewportRef}
      className="order-1 min-h-0 flex-1 overflow-auto bg-slate-200/80"
      data-testid="spreadsheet-page-layout-view"
      data-spreadsheet-cell-host=""
      data-sheet-id={sheetId}
      data-page-count={contentGeometry.totalPages}
      data-browsing-page-count={geometry.totalPages}
      data-rendered-page-count={renderedPageCount}
      data-rendered-cell-count={renderPlan.renderedCellCount}
      data-print-page-start={printBatch ? printPageStart : undefined}
      data-print-next-start={printBatch?.nextStart ?? undefined}
      role="region"
      aria-label={t("spreadsheet.helperCopy.pageLayout.view")}
      tabIndex={0}
      onScroll={readViewport}
      onFocus={trackCellFocus}
      onBlur={trackCellBlur}
      style={{ outlineOffset: -2 }}
      onPointerDown={event=>{if(event.target===event.currentTarget)onClearDrawingSelection?.();}}
    >
      {drawingError ? <span role="alert" className="sticky left-0 top-0 z-30 bg-white px-3 text-sm">{t("spreadsheet.helperCopy.pageLayout.drawingChanged")}</span> : null}
      {showRuler ? (
      <div
        data-testid="spreadsheet-page-layout-horizontal-ruler"
        data-page-layout-chrome="outside-paper"
        data-ruler-unit={rulerUnit.points === 72 ? "in" : "cm"}
        className="sticky top-0 z-20 border-b border-slate-400 bg-slate-100 shadow-sm"
        style={{
          left: 80,
          marginLeft: 80,
          width: geometry.paperWidth * paperScale,
          height: 28,
          backgroundImage:
            `repeating-linear-gradient(90deg, #9aa4b2 0 1px, transparent 1px ${rulerTickPx}px)`,
          backgroundSize: `100% 7px`,
          backgroundRepeat: "repeat-x",
          backgroundPosition: `${displayedMargins.left * paperScale}px 100%`,
        }}
      >
        <span
          aria-hidden="true"
          className="absolute inset-y-0 left-0 bg-emerald-100/70"
          style={{ width: displayedMargins.left * paperScale }}
        />
        <span
          aria-hidden="true"
          className="absolute inset-y-0 right-0 bg-slate-50/90"
          style={{ width: displayedMargins.right * paperScale }}
        />
        {rulerMarks(geometry.dataAreaWidth).map((mark) => (
          <span
            key={mark}
            aria-hidden="true"
            data-page-layout-ruler-mark={mark}
            className="absolute top-0 -translate-x-1/2 text-[9px] text-slate-600"
            style={{ left: (displayedMargins.left + mark * rulerUnit.points) * paperScale }}
          >
            {mark}
          </span>
        ))}
        {(["left", "right"] as const).map((side) => {
          const value = displayedMargins[side];
          const opposite = displayedMargins[side === "left" ? "right" : "left"];
          const maximum = Math.max(
            0,
            geometry.paperWidth - opposite - PAGE_LAYOUT_MIN_PRINTABLE_WIDTH_PT,
          );
          let label: string;
          if (side === "left") label = t("spreadsheet.helperCopy.pageLayout.leftMargin");
          else label = t("spreadsheet.helperCopy.pageLayout.rightMargin");
          const accessibleLabel = marginResizeDisabledReason
            ? `${label}: ${marginResizeDisabledReason}`
            : label;
          return (
            <span
              key={side}
              role="separator"
              aria-orientation="vertical"
              aria-label={accessibleLabel}
              aria-valuemin={0}
              aria-valuemax={Math.round(maximum)}
              aria-valuenow={Math.round(value)}
              aria-disabled={marginResizeDisabledReason ? "true" : undefined}
              data-disabled-reason={marginResizeDisabledReason}
              data-testid={`spreadsheet-page-layout-margin-${side}`}
              tabIndex={0}
              title={accessibleLabel}
              className="absolute inset-y-0 z-30 w-3 -translate-x-1/2 cursor-col-resize touch-none border-l-2 border-emerald-700 outline-none focus-visible:ring-2 focus-visible:ring-emerald-600"
              style={{
                left:
                  (side === "left" ? value : geometry.paperWidth - value) *
                  paperScale,
              }}
              onPointerDown={(event) => startMarginResize(side, event)}
              onKeyDown={(event) => onMarginKeyDown(side, event)}
            />
          );
        })}
      </div>
      ) : null}
      <div
        className="grid w-max min-w-full items-start gap-6 pb-12 pl-20 pr-6 pt-3"
        style={{
          gridTemplateColumns: `repeat(${pageColumnCount}, ${geometry.paperWidth * paperScale}px)`,
          gridTemplateRows: `repeat(${pageRowCount}, ${geometry.paperHeight * paperScale + PAGE_LAYOUT_PAGE_TOP_CHROME}px)`,
        }}
      >
        {/* Every page not built yet shows its paper at once (a scroll used
            to reach an empty grey band and wait for the page to appear);
            its cells come when the page is built. Never in print. */}
        {printMode ? null : pagePaperPlaceholders(builtPages).map(({ row, col }) => (
          <div
            key={`paper:${row}:${col}`}
            aria-hidden="true"
            data-testid="spreadsheet-page-layout-paper-placeholder"
            className="relative"
            style={{
              gridColumn: col + 1,
              gridRow: row + 1,
              width: geometry.paperWidth * paperScale,
              height: geometry.paperHeight * paperScale + PAGE_LAYOUT_PAGE_TOP_CHROME,
            }}
          >
            <div
              className="absolute left-0 border border-slate-300 bg-white shadow-[0_3px_12px_rgba(15,23,42,0.18)]"
              style={{
                top: PAGE_LAYOUT_PAGE_TOP_CHROME,
                width: geometry.paperWidth * paperScale,
                height: geometry.paperHeight * paperScale,
              }}
            />
          </div>
        ))}
        {builtPages.map((page) => {
          if (!page.rows.length || !page.cols.length) return null;
          // A page whose inputs are as they were is handed back as the same
          // element, so React skips its whole subtree: a scroll that brings
          // new pages into view builds those pages only, not every page and
          // cell already on screen (each band used to re-render them all).
          const cacheTile = resolveSpreadsheetPageLayoutTilePosition(geometry, page.index) ?? { row: 0, col: 0 };
          const cacheKey = `${cacheTile.row}:${cacheTile.col}`;
          // The cells of this page that were asked for by address: only the
          // page holding a cell that mounts its editor is built again.
          let pageMaterialized = "";
          for (const key of focusedCell ? [...materializedCells, focusedCell] : materializedCells) {
            const split = key.indexOf(":");
            const row = Number(key.slice(0, split)), col = Number(key.slice(split + 1));
            if (page.cols.includes(col) && page.rows.some((visual) => sourceRowAt(visual) === row)) pageMaterialized += `${key};`;
          }
          const cacheDeps: readonly unknown[] = [
            pageLayoutProps, keyboardResize, linePointerActive, marginPreview, drawingPreview, drawingError, pageMaterialized,
            entranceSettled, cacheTile.row === pageRowCount - 1, cacheTile.col === pageColumnCount - 1,
            page.pageNumber, page.rows.length, page.rows[0], page.rows[page.rows.length - 1],
            page.cols.length, page.cols[0], page.cols[page.cols.length - 1],
          ];
          const cachedPage = previousPageElements.get(cacheKey);
          if (cachedPage && cachedPage.deps.length === cacheDeps.length && cachedPage.deps.every((value, at) => Object.is(value, cacheDeps[at]))) {
            nextPageElements.set(cacheKey, cachedPage);
            return cachedPage.element;
          }
          const pageElement = (() => {
          const overThenDown = geometry.settings.pageOrder === "overThenDown";
          const tile = resolveSpreadsheetPageLayoutTilePosition(geometry, page.index) ?? { row: 0, col: 0 };
          const rowPageIndex = tile.row;
          const colPageIndex = tile.col;
          const contentRows = contentGeometry.pagination.rowRanges.length;
          const contentCols = contentGeometry.pagination.colRanges.length;
          const blankPage = !geometry.segments && (rowPageIndex >= contentRows || colPageIndex >= contentCols);
          // Pages of a multi-area print area are numbered straight through in
          // area order; browsing blanks never exist beside a print area.
          const printPageNumber = geometry.segments ? page.pageNumber : (geometry.settings.firstPageNumber ?? 1) + (overThenDown
            ? rowPageIndex * contentCols + colPageIndex : colPageIndex * contentRows + rowPageIndex);
          const label = blankPage ? (t("spreadsheet.helperCopy.pageLayout.blankPage")) : pageLabel(printPageNumber);
          const variant = resolvePageHeaderFooter(geometry.settings, { pageNumber: printPageNumber });
          // Editing a band edits the variant this page prints.
          const variantKey = resolvePageHeaderFooterVariant(geometry.settings, { pageNumber: printPageNumber });
          const context: HeaderFooterContext = {
            pageNumber: printPageNumber,
            totalPages: contentGeometry.totalPages,
            date,
            time,
            sheetName,
            fileName,
            ...(filePath ? { filePath } : {}),
          };
          const colOffsets = new Map<number, number>();
          const rowOffsets = new Map<number, number>();
          let gridWidthPoints = 0;
          let gridHeightPoints = 0;
          for (const col of page.cols) {
            colOffsets.set(col, gridWidthPoints);
            gridWidthPoints += geometry.colSizes[col] ?? geometry.colWidth;
          }
          for (const row of page.rows) {
            rowOffsets.set(row, gridHeightPoints);
            gridHeightPoints += geometry.rowSizes[row] ?? geometry.rowHeight;
          }
          const gridWidth = gridWidthPoints * paperScale;
          const gridHeight = gridHeightPoints * paperScale;
          // Printed headings (Page Setup > Sheet) sit inside the margins
          // before the cells: the letters above, the numbers on the leading
          // side. Centre on page centres them with the cells.
          const headingWidth = printMode ? geometry.headingWidth * paperScale : 0;
          const headingHeight = printMode ? geometry.headingHeight * paperScale : 0;
          const centerOffset = geometry.settings.centerHorizontally
            ? Math.max(
                0,
                (geometry.dataAreaWidth * paperScale - gridWidth - headingWidth) / 2,
              )
            : 0;
          // A right-to-left sheet starts at the right margin.
          const dataLeft = rightToLeft
            ? geometry.paperWidth * paperScale -
              geometry.margins.right * paperScale -
              centerOffset -
              headingWidth -
              gridWidth
            : geometry.margins.left * paperScale + centerOffset + headingWidth;
          // Where a run of columns starting at `first` and `widthPoints` wide
          // sits from the grid's left edge, in points: mirrored right to left.
          const columnsLeft = (first: number, widthPoints: number) => {
            const offset = colOffsets.get(first) ?? 0;
            return rightToLeft ? gridWidthPoints - offset - widthPoints : offset;
          };
          const dataTop =
            geometry.margins.top * paperScale +
            headingHeight +
            (geometry.settings.centerVertically
              ? Math.max(
                  0,
                  (geometry.dataAreaHeight * paperScale - gridHeight - headingHeight) / 2,
                )
              : 0);
          // Off the page margins, Excel insets the side sections a fixed
          // 0.75 in from the paper edge (its own PDF, whatever the margins).
          const bandLeft = alignWithMargins
            ? geometry.margins.left * paperScale
            : Math.min(54, geometry.paperWidth / 2) * paperScale;
          const bandRight = alignWithMargins
            ? geometry.margins.right * paperScale
            : Math.min(54, geometry.paperWidth / 2) * paperScale;
          const seen = new Set<string>();
          const sizeOf = (
            indices: readonly number[],
            sizes: readonly number[],
            fallback: number,
          ) =>
            indices.reduce((sum, index) => sum + (sizes[index] ?? fallback), 0);
          const fontScale =
            Math.max(0.001, geometry.contentScale) *
            Math.max(0.1, Math.min(4, zoom));
          // One style read per cell of the page: a cell's shared-edge check
          // reads its left and upper neighbours' styles too.
          const pageCellStyles = new Map<string, CSSProperties | undefined>();
          const pageCellStyle = (row: number, col: number) => {
            const styleKey = `${row}:${col}`;
            if (!pageCellStyles.has(styleKey)) pageCellStyles.set(styleKey, getCellStyle?.(row, col));
            return pageCellStyles.get(styleKey);
          };
          const statesEdge = (style: CSSProperties | undefined, side: "Right" | "Bottom") => {
            const value = style?.[`border${side}`] ?? style?.border;
            return typeof value === "string" && value !== "" && !/^\s*(none|0(px)?)\b/.test(value);
          };
          const pageColIndex = new Map(page.cols.map((value, index) => [value, index]));
          const pageRowIndex = new Map(page.rows.map((value, index) => [value, index]));
          // Whether the cell drawn next to (row, col) on this page -- on its
          // left, or above -- states the edge they share. A merge states it
          // with its anchor's style; a neighbour on another page draws
          // nothing here.
          const neighbourStatesEdge = (row: number, col: number, side: "Left" | "Top") => {
            // Right to left, the row above is still above: only a side
            // neighbour lies elsewhere (P09's bordered rows printed every
            // horizontal edge doubled).
            if (!nativeTextDpi || (rightToLeft && side === "Left")) return false;
            const colIndex = pageColIndex.get(col);
            const rowIndex = pageRowIndex.get(row);
            if (colIndex === undefined || rowIndex === undefined) return false;
            const neighbourRow = side === "Top" ? page.rows[rowIndex - 1] : row;
            const neighbourCol = side === "Left" ? page.cols[colIndex - 1] : col;
            if (neighbourRow === undefined || neighbourCol === undefined) return false;
            const sourceRow = sourceRowAt(neighbourRow);
            const merge = mergeAt(sourceRow, neighbourCol);
            return statesEdge(pageCellStyle(merge?.row ?? sourceRow, merge?.col ?? neighbourCol), side === "Left" ? "Right" : "Bottom");
          };
          // The inner edges a merge or a text run covers: no gridline is
          // printed across them (Excel prints a run's words, and a merge's
          // inside, clear of its gridlines; probe E1, P12).
          const gridlineGaps = new Set<string>();
          const coverEdges = (rowsCovered: readonly number[], colsCovered: readonly number[]) => {
            const rowIndexes = rowsCovered.map((value) => pageRowIndex.get(value)).filter((value): value is number => value !== undefined);
            const colIndexes = colsCovered.map((value) => pageColIndex.get(value)).filter((value): value is number => value !== undefined);
            for (const rowIndex of rowIndexes) {
              for (const colIndex of colIndexes) {
                if (colIndexes.includes(colIndex + 1)) gridlineGaps.add(`v:${colIndex}:${rowIndex}`);
                if (rowIndexes.includes(rowIndex + 1)) gridlineGaps.add(`h:${rowIndex}:${colIndex}`);
              }
            }
          };
          const cellNodes = page.rows.flatMap((row) =>
            page.cols.flatMap((col) => {
              const sourceRow = sourceRowAt(row);
              const merge = mergeAt(sourceRow, col);
              const candidates = merge
                ? page.rows.filter(
                    (value) =>
                      sourceRowAt(value) >= merge.row && sourceRowAt(value) < merge.row + merge.rowSpan,
                  )
                : [row];
              // A sorted view may split a merge into separate display bands.
              // Never paint over unrelated rows between those bands.
              const groups: number[][] = [];
              for (const value of candidates) {
                const previous = groups.at(-1)?.at(-1);
                if (previous !== undefined && value === previous + 1 && sourceRowAt(value) > sourceRowAt(previous)) groups[groups.length - 1].push(value);
                else groups.push([value]);
              }
              const rows = groups.find((group) => group.includes(row)) ?? [row];
              const key = merge ? `${merge.row}:${merge.col}` : `${sourceRow}:${col}`;
              const fragmentKey = `${key}:${rows[0]}`;
              if (seen.has(fragmentKey)) return [];
              seen.add(fragmentKey);
              const cols = merge
                ? page.cols.filter(
                    (value) =>
                      value >= merge.col && value < merge.col + merge.colSpan,
                  )
                : [col];
              if (!rows.length || !cols.length) return [];
              const anchorRow = merge?.row ?? sourceRow;
              const anchorCol = merge?.col ?? col;
              const cellWidthPx = sizeOf(cols, geometry.colSizes, geometry.colWidth) * paperScale;
              const spill = !merge ? getTextSpill?.(anchorRow, anchorCol, cellWidthPx, fontScale) ?? null : null;
              const spillBox = (() => {
                if (!spill) return undefined;
                const first = page.cols[0];
                const last = page.cols[page.cols.length - 1];
                const startCol = Math.max(spill.startCol, first);
                const endCol = Math.min(spill.endCol, last);
                if (startCol === col && endCol === col) return undefined;
                const span = (from: number, to: number) => Array.from({ length: Math.max(0, to - from) }, (_, index) => from + index);
                return {
                  // Right to left, the columns after this one lie on its left.
                  left: -sizeOf(rightToLeft ? span(col + 1, endCol + 1) : span(startCol, col), geometry.colSizes, geometry.colWidth) * paperScale,
                  width: sizeOf(span(startCol, endCol + 1), geometry.colSizes, geometry.colWidth) * paperScale,
                };
              })();
              if (printMode && merge) coverEdges(rows, cols);
              if (printMode && spill && spillBox) {
                coverEdges([row], page.cols.filter((value) => value >= spill.startCol && value <= spill.endCol));
              }
              const cellStyle = pageCellStyle(anchorRow, anchorCol);
              const display = getCellDisplay(anchorRow, anchorCol);
              const nativeText = nativeTextDpi && hasNativeTextLayout?.(anchorRow,anchorCol,display);
              // A text that runs on is laid out over the columns it runs into
              // (its spill box), so its alignment holds over the whole run: laid
              // out in its own column, a right-aligned run was set from the
              // run's far end (P12's RTL rows printed broken and overlapping).
              const nativeTextWidth = nativeText ? Math.max(1,
                (merge?lineTotal(geometry.colSizes,merge.col,merge.col+merge.colSpan):spillBox?spillBox.width/paperScale:sizeOf(cols,geometry.colSizes,geometry.colWidth))
                * PAPER_POINT_TO_CSS_PIXEL / Math.max(.001,geometry.contentScale) - 5*96/nativeTextDpi
              ) : undefined;
              const active = merge
                ? activeCell.row >= merge.row &&
                  activeCell.row < merge.row + merge.rowSpan &&
                  activeCell.col >= merge.col &&
                  activeCell.col < merge.col + merge.colSpan
                : activeCell.row === sourceRow && activeCell.col === col;
              const selected =
                anchorRow >= selection.r1 &&
                anchorRow <= selection.r2 &&
                anchorCol >= selection.c1 &&
                anchorCol <= selection.c2;
              const interactive =
                !merge || (sourceRowAt(rows[0]) === merge.row && cols[0] === merge.col);
              const cellKey = `${anchorRow}:${anchorCol}`;
              const occurrence = `${sheetId}-${rowPageIndex}.${colPageIndex}-${fragmentKey}`;
              // Only the active cell, the edited one and those asked for
              // carry an editor; print carries none.
              const editor = interactive && !printMode && (active ||
                (editingCell?.row === anchorRow && editingCell.col === anchorCol) ||
                materializedCells.has(cellKey) || focusedCell === cellKey);
              return [
                <GridCell
                  // Keyed within its page's tile, never by the page's
                  // number: browsing past the last page adds blank pages,
                  // which renumbers the pages, and a key with the number in
                  // it remounted every cell of every page on each scroll.
                  key={fragmentKey}
                  row={anchorRow}
                  col={anchorCol}
                  display={display}
                  cellStyle={cellStyle}
                  left={columnsLeft(cols[0], sizeOf(cols, geometry.colSizes, geometry.colWidth)) * paperScale}
                  top={(rowOffsets.get(rows[0]) ?? 0) * paperScale}
                  width={
                    sizeOf(cols, geometry.colSizes, geometry.colWidth) *
                    paperScale
                  }
                  height={
                    sizeOf(rows, geometry.rowSizes, geometry.rowHeight) *
                    paperScale
                  }
                  fontScale={fontScale}
                  nativeTextDpi={nativeText?nativeTextDpi:undefined}
                  nativeGridDpi={nativeTextDpi}
                  // Printed gridlines are drawn by PrintedGridlines (exact 1 pt).
                  showGridlines={pageGridlines && !printMode}
                  active={!printMode && active}
                  selected={!printMode && selected}
                  interactive={interactive}
                  editor={editor}
                  occurrence={occurrence}
                  paintsDisplay={!renderCellContent}
                  cellHandlers={cellHandlers}
                  overlay={interactive ? renderCellOverlay?.(anchorRow, anchorCol, occurrence, {
                    width: sizeOf(cols, geometry.colSizes, geometry.colWidth) * paperScale,
                    height: sizeOf(rows, geometry.rowSizes, geometry.rowHeight) * paperScale,
                    fontScale,
                  }) : undefined}
                  content={renderCellContent?.(anchorRow, anchorCol, fontScale, nativeTextWidth, {
                    // The box the content is laid out in: the whole merge.
                    width: (merge ? lineTotal(geometry.colSizes, merge.col, merge.col + merge.colSpan) : sizeOf(cols, geometry.colSizes, geometry.colWidth)) * paperScale,
                    height: (merge ? sourceRowHeight(merge.row, merge.row + merge.rowSpan) : sizeOf(rows, geometry.rowSizes, geometry.rowHeight)) * paperScale,
                  })}
                  paintsOutside={!merge && Boolean(cellPaintsOutside?.(anchorRow, anchorCol))}
                  sharedLeftEdge={!merge && neighbourStatesEdge(row, col, "Left")}
                  sharedTopEdge={!merge && neighbourStatesEdge(row, col, "Top")}
                  spillBox={spillBox}
                  mergeContent={
                    merge
                      ? {
                          key,
                          // The merge's columns before this fragment lie to
                          // its left, or to its right on a right-to-left sheet.
                          left: rightToLeft
                            ? (sizeOf(cols, geometry.colSizes, geometry.colWidth) +
                                lineTotal(geometry.colSizes, merge.col, cols[0]) -
                                lineTotal(geometry.colSizes, merge.col, merge.col + merge.colSpan)) *
                              paperScale
                            : -lineTotal(geometry.colSizes, merge.col, cols[0]) *
                              paperScale,
                          top:
                            -sourceRowHeight(merge.row, sourceRowAt(rows[0])) *
                            paperScale,
                          width:
                            lineTotal(
                              geometry.colSizes,
                              merge.col,
                              merge.col + merge.colSpan,
                            ) * paperScale,
                          height:
                            sourceRowHeight(
                              merge.row,
                              merge.row + merge.rowSpan,
                            ) * paperScale,
                        }
                      : undefined
                  }
                  inputProps={
                    editor
                      ? getCellInputProps(anchorRow, anchorCol, occurrence)
                      : undefined
                  }
                />,
              ];
            }),
          );
          const {cols: drawingColumnBands, rows: drawingRowBands} = drawingBandsForPage(page);
          const renderAxis = (column: boolean) => {
            const indices = column ? page.cols : page.rows;
            const offsets = column ? colOffsets : rowOffsets;
            const sizes = column ? geometry.colSizes : geometry.rowSizes;
            const fallback = column ? geometry.colWidth : geometry.rowHeight;
            const active = column ? activeCell?.col : activeCell?.row;
            const axis = column ? "col" : "row";
            let label: string;
            if (column) label = t("spreadsheet.helperCopy.pageLayout.resizeColumn");
            else label = t("spreadsheet.helperCopy.pageLayout.resizeRow");
            const accessibleLabel = lineResizeDisabledReason
              ? `${label}: ${lineResizeDisabledReason}`
              : label;
            return (
              <div className="contents">
                {!column && showRuler ? (
                  // XLS-A030: the vertical ruler beside the page, in the
                  // region's unit from the top margin, with the top and
                  // bottom margin handles.
                  <div
                    data-testid="spreadsheet-page-layout-vertical-ruler"
                    data-page-layout-chrome="outside-paper"
                    className="absolute border-r border-slate-400 bg-slate-100"
                    style={{
                      left: -58,
                      top: PAGE_LAYOUT_PAGE_TOP_CHROME,
                      width: 18,
                      height: geometry.paperHeight * paperScale,
                      backgroundImage:
                        `repeating-linear-gradient(180deg, #9aa4b2 0 1px, transparent 1px ${rulerTickPx}px)`,
                      backgroundSize: "6px 100%",
                      backgroundRepeat: "repeat-y",
                      backgroundPosition: `100% ${displayedMargins.top * paperScale}px`,
                    }}
                  >
                    <span
                      aria-hidden="true"
                      className="absolute inset-x-0 top-0 bg-emerald-100/70"
                      style={{ height: displayedMargins.top * paperScale }}
                    />
                    <span
                      aria-hidden="true"
                      className="absolute inset-x-0 bottom-0 bg-slate-50/90"
                      style={{ height: displayedMargins.bottom * paperScale }}
                    />
                    {rulerMarks(geometry.dataAreaHeight).map((mark) => (
                      <span
                        key={mark}
                        aria-hidden="true"
                        data-page-layout-ruler-mark={mark}
                        className="absolute left-0 -translate-y-1/2 text-[9px] leading-none text-slate-600"
                        style={{ top: (displayedMargins.top + mark * rulerUnit.points) * paperScale }}
                      >
                        {mark}
                      </span>
                    ))}
                    {(["top", "bottom"] as const).map((side) => {
                      const value = displayedMargins[side];
                      const label = t(
                        side === "top"
                          ? "spreadsheet.helperCopy.pageLayout.topMargin"
                          : "spreadsheet.helperCopy.pageLayout.bottomMargin",
                      );
                      const accessibleLabel = marginResizeDisabledReason
                        ? `${label}: ${marginResizeDisabledReason}`
                        : label;
                      return (
                        <span
                          key={side}
                          role="separator"
                          aria-orientation="horizontal"
                          aria-label={accessibleLabel}
                          aria-valuemin={0}
                          aria-valuemax={Math.round(
                            Math.max(
                              0,
                              geometry.paperHeight -
                                displayedMargins[OPPOSITE_MARGIN[side]] -
                                PAGE_LAYOUT_MIN_PRINTABLE_WIDTH_PT,
                            ),
                          )}
                          aria-valuenow={Math.round(value)}
                          aria-disabled={marginResizeDisabledReason ? "true" : undefined}
                          data-disabled-reason={marginResizeDisabledReason}
                          data-testid={`spreadsheet-page-layout-margin-${side}`}
                          tabIndex={0}
                          title={accessibleLabel}
                          className="absolute inset-x-0 z-30 h-3 -translate-y-1/2 cursor-row-resize touch-none border-t-2 border-emerald-700 outline-none focus-visible:ring-2 focus-visible:ring-emerald-600"
                          style={{
                            top:
                              (side === "top" ? value : geometry.paperHeight - value) *
                              paperScale,
                          }}
                          onPointerDown={(event) => startMarginResize(side, event)}
                          onKeyDown={(event) => onMarginKeyDown(side, event)}
                        />
                      );
                    })}
                  </div>
                ) : null}
                {indices.map((index) => (
                  <div
                    key={index}
                    {...(column
                      ? { "data-page-layout-column": index }
                      : { "data-page-layout-row": sourceRowAt(index), "data-page-layout-visual-row": index })}
                    data-page-layout-chrome="outside-paper"
                    role={column ? "columnheader" : "rowheader"}
                    className="absolute box-border flex items-center justify-center border border-slate-300 text-[10px]"
                    style={{
                      left: column
                        ? dataLeft + columnsLeft(index, sizes[index] ?? fallback) * paperScale
                        : -36,
                      top: column
                        ? 0
                        : PAGE_LAYOUT_PAGE_TOP_CHROME +
                          dataTop +
                          (offsets.get(index) ?? 0) * paperScale,
                      width: column
                        ? (sizes[index] ?? fallback) * paperScale
                        : 34,
                      height: column
                        ? 24
                        : (sizes[index] ?? fallback) * paperScale,
                      background: active === (column ? index : sourceRowAt(index)) ? "#d9eee2" : "#f5f6f8",
                      color: active === (column ? index : sourceRowAt(index)) ? "#17633a" : "#4b5563",
                    }}
                  >
                    {column ? columnLabel(index) : sourceRowAt(index) + 1}
                    <span
                      role="separator"
                      aria-orientation={column ? "vertical" : "horizontal"}
                      aria-label={accessibleLabel}
                      aria-valuemin={
                        column ? MIN_COLUMN_WIDTH_PX : MIN_ROW_HEIGHT_PX
                      }
                      aria-valuemax={
                        column ? MAX_COLUMN_WIDTH_PX : MAX_ROW_HEIGHT_PX
                      }
                      aria-valuenow={lineSize(axis, index)}
                      aria-disabled={
                        lineResizeDisabledReason ? "true" : undefined
                      }
                      data-disabled-reason={lineResizeDisabledReason}
                      data-testid={`spreadsheet-page-layout-${axis}-resize-${index}`}
                      tabIndex={0}
                      title={accessibleLabel}
                      className="absolute z-10 block touch-none outline-none focus-visible:ring-2 focus-visible:ring-emerald-600"
                      style={
                        column
                          ? {
                              top: 0,
                              ...(rightToLeft ? { left: -4 } : { right: -4 }),
                              width: 8,
                              height: "100%",
                              cursor: "col-resize",
                              // A finger drags the divider; the page view may
                              // not take it for a scroll.
                              touchAction: "none",
                            }
                          : {
                              left: 0,
                              bottom: -4,
                              width: "100%",
                              height: 8,
                              cursor: "row-resize",
                              touchAction: "none",
                            }
                      }
                      onPointerDown={(event) => {
                        if (lineResizeDisabledReason) return;
                        event.currentTarget.setPointerCapture?.(
                          event.pointerId,
                        );
                        startLineResize(axis, index, event);
                      }}
                      onKeyDown={(event) => onResizeKeyDown(axis, index, event)}
                    />
                  </div>
                ))}
              </div>
            );
          };
          return (
            <PageLayoutPageTile
              key={`${rowPageIndex}:${colPageIndex}`}
              entranceSettled={entranceSettled}
              style={{
                gridColumn: colPageIndex + 1,
                gridRow: rowPageIndex + 1,
                width: geometry.paperWidth * paperScale,
                height:
                  geometry.paperHeight * paperScale +
                  PAGE_LAYOUT_PAGE_TOP_CHROME,
              }}
            >
              {renderAxis(true)}
              {colPageIndex === 0 ? renderAxis(false) : null}
              <article
                role="group"
                aria-label={label}
                data-testid="spreadsheet-page-layout-page"
                data-page-tile={`${rowPageIndex}:${colPageIndex}`}
                data-print-page-index={blankPage ? undefined : printPageNumber - (geometry.settings.firstPageNumber ?? 1)}
                data-page-number={blankPage ? undefined : printPageNumber}
                data-blank-page={blankPage || undefined}
                className="relative overflow-hidden border border-slate-300 bg-white shadow-[0_3px_12px_rgba(15,23,42,0.18)]"
                style={{
                  position: "absolute",
                  left: 0,
                  top: PAGE_LAYOUT_PAGE_TOP_CHROME,
                  width: geometry.paperWidth * paperScale,
                  height: geometry.paperHeight * paperScale,
                  // Printed gridlines are black, as Excel prints them (its
                  // own PDF); on screen they keep the view's light grey.
                  ...(printMode ? { "--page-gridline": "#000" } : {}),
                  // Automatic text prints black (Excel's window text), not
                  // the view's slate (#111827 printed on every page).
                  ...(printMode ? { "--page-text": "#000" } : {}),
                  // Black and white prints the gridlines dotted (1 pt on,
                  // 1 pt off) inside a solid frame: Excel's own PDF of the
                  // same sheet with and without the option (P05).
                  ...(printMode && geometry.settings.blackAndWhite === true ? { "--page-gridline-style": "dotted" } : {}),
                } as CSSProperties}
              >
                <HeaderFooterBand
                  printMode={printMode}
                  fontFamily={headerFooterFontFamily}
                  key={`${sheetId}-header-${rowPageIndex}.${colPageIndex}`}
                  kind="header"
                  value={variant.header}
                  context={context}
                  paperScale={paperScale}
                  documentScale={documentScale}
                  translate={t}
                  disabledReason={headerFooterDisabledReason}
                  onCommit={(align, text) => onCommitHeaderFooter(
                    variantKey === "first" ? "firstHeader" : variantKey === "even" ? "evenHeader" : "header", align, text,
                  )}
                  placeholder={t("spreadsheet.helperCopy.pageLayout.addHeader")}
                  style={{
                    top: geometry.headerMargin * paperScale,
                    left: bandLeft,
                    right: bandRight,
                    minHeight:
                      Math.max(
                        0,
                        geometry.margins.top - geometry.headerMargin,
                      ) * paperScale,
                  }}
                />
                <div
                  role="grid"
                  aria-label={`${sheetName} ${label}`}
                  data-testid="spreadsheet-page-layout-grid"
                  ref={(node) => { const tileKey = `${rowPageIndex}:${colPageIndex}`; if (node) drawingGridRefs.current.set(tileKey, node); else drawingGridRefs.current.delete(tileKey); }}
                  className="absolute overflow-hidden"
                  style={{
                    left: dataLeft,
                    top: dataTop,
                    width: gridWidth,
                    height: gridHeight,
                    // Grid chrome must not move cells away from drawing/axis
                    // coordinates. Cell borders already own their box insets.
                    boxShadow: pageGridlines && !printMode ? "inset 1px 1px var(--page-gridline, #d9dee7)" : undefined,
                    ...(backgroundPicture
                      ? {
                          backgroundImage: `url("${backgroundPicture}")`,
                          backgroundRepeat: "repeat",
                          backgroundPosition: "0 0",
                        }
                      : {}),
                  }}
                  data-sheet-background={backgroundPicture ? "true" : undefined}
                >
                  {printMode && pageGridlines ? (
                    <PrintedGridlines
                      width={gridWidth}
                      height={gridHeight}
                      lineWidth={paperScale}
                      dotted={geometry.settings.blackAndWhite === true}
                      columns={page.cols.map((col) => {
                        const size = geometry.colSizes[col] ?? geometry.colWidth;
                        const offset = (colOffsets.get(col) ?? 0) + size;
                        const right = (rightToLeft ? gridWidthPoints - offset + size : offset) * paperScale;
                        return { left: right - size * paperScale, right };
                      })}
                      rows={page.rows.map((row) => {
                        const top = (rowOffsets.get(row) ?? 0) * paperScale;
                        return { top, bottom: top + (geometry.rowSizes[row] ?? geometry.rowHeight) * paperScale };
                      })}
                      rightToLeft={rightToLeft}
                      gaps={gridlineGaps}
                    />
                  ) : null}
                  {cellNodes}
                </div>
                {printMode && printedNotes?.size
                  ? page.rows.flatMap((row) => page.cols.flatMap((col) => {
                      const text = printedNotes.get(`${sourceRowAt(row)}:${col}`);
                      if (!text) return [];
                      const width = geometry.colSizes[col] ?? geometry.colWidth;
                      const scale = geometry.contentScale * paperScale;
                      const cellLeft = dataLeft + columnsLeft(col, width) * paperScale;
                      const cellRight = cellLeft + width * paperScale;
                      return [
                        <PrintedNote
                          key={`note-${row}:${col}`}
                          text={text}
                          left={rightToLeft ? cellLeft - (NOTE_GAP_PT + NOTE_WIDTH_PT) * scale : cellRight + NOTE_GAP_PT * scale}
                          top={dataTop + (rowOffsets.get(row) ?? 0) * paperScale - NOTE_RISE_PT * scale}
                          scale={scale}
                          lineWidth={0.5 * paperScale}
                        />,
                      ];
                    }))
                  : null}
                {headingWidth > 0 && headingHeight > 0 ? (
                  <PrintedHeadings
                    left={rightToLeft ? dataLeft : dataLeft - headingWidth}
                    top={dataTop - headingHeight}
                    gridLeft={rightToLeft ? 0 : headingWidth}
                    headingWidth={headingWidth}
                    headingHeight={headingHeight}
                    gridWidth={gridWidth}
                    gridHeight={gridHeight}
                    lineWidth={paperScale}
                    font={headingFont}
                    fontSize={(headingFont?.sizePt ?? 11) * geometry.contentScale * paperScale}
                    columns={page.cols.map((col) => {
                      const width = geometry.colSizes[col] ?? geometry.colWidth;
                      return { key: col, label: columnLabel(col), start: columnsLeft(col, width) * paperScale, size: width * paperScale };
                    })}
                    rows={page.rows.map((row) => ({
                      key: row,
                      label: String(sourceRowAt(row) + 1),
                      start: (rowOffsets.get(row) ?? 0) * paperScale,
                      size: (geometry.rowSizes[row] ?? geometry.rowHeight) * paperScale,
                    }))}
                  />
                ) : null}
                {drawingRowBands.flatMap((rowBand, rowBandIndex) => drawingColumnBands.flatMap((colBand, colBandIndex) => drawings.map((drawing) => {
                  const [
                    id,
                    [drawingLeft, drawingTop, width, height],
                    label,
                    content,
                  ] = drawing;
                  const scale = pxToPt(geometry.contentScale);
                  const left = drawingLeft * scale;
                  const top = drawingTop * scale;
                  // An outline is centred on the frame's edge: its outer half
                  // lies past the frame, and clipped to the frame a 1.5 pt
                  // line printed 0.75 pt (P20's text box). The page's own
                  // band still clips it.
                  const bleed = DRAWING_OUTLINE_BLEED_PT * geometry.contentScale;
                  const clip = [
                    Math.max(left - bleed, colBand.start),
                    Math.max(top - bleed, rowBand.start),
                    Math.min(left + width * scale + bleed, colBand.end),
                    Math.min(top + height * scale + bleed, rowBand.end),
                  ] as const;
                  if (clip[0] >= clip[2] || clip[1] >= clip[3]) return null;
                  return (
                    <div
                      key={`${rowPageIndex}.${colPageIndex}:${rowBandIndex}:${colBandIndex}:${id}`}
                      aria-label={label}
                      data-page-layout-drawing={id}
                      tabIndex={0}
                      aria-disabled={drawing[4]?.disabledReason ? true : undefined}
                      data-disabled-reason={drawing[4]?.disabledReason}
                      title={drawing[4]?.disabledReason ?? label}
                      className="absolute z-[6]"
                      onPointerDownCapture={(event) => beginDrawing(event, drawing)}
                      onClick={(event) => { event.stopPropagation(); const moved=drawingClickMoved.current;drawingClickMoved.current=false;onSelectDrawing?.(id,event.shiftKey||event.ctrlKey||event.metaKey,moved); }}
                      onFocus={() => onSelectDrawing?.(id,false,true)}
                      onKeyDown={(event) => {
                        if (event.target===event.currentTarget && event.key === "Escape") {event.preventDefault();event.stopPropagation();onClearDrawingSelection?.();return;}
                        if (event.target !== event.currentTarget || !drawing[4] || drawing[4].disabledReason) return;
                        // Excel's arrows on a selected object (the shared
                        // reading, `lib/office-object-nudge`): a set step, Ctrl
                        // one SCREEN pixel of the paper, Shift resize.
                        const action = officeObjectKeyAction(event, paperScale, { plainStep: "screen-pixel", shiftArrow: isMacPlatform() ? "move" : "resize" });
                        if (!action) return;
                        event.preventDefault(); event.stopPropagation();
                        // Alt + Left / Right: 15 degrees, as in Normal view.
                        if (action.kind === "rotate") { onRotateDrawing?.(id, action.degrees); return; }
                        const transaction = onBeginDrawingTransform?.(id);
                        const next = action.kind === "move"
                          ? transformPageLayoutDrawing(drawing[4].frame, drawing[4].rotation, undefined, action.dx, action.dy)
                          : transformPageLayoutDrawing(drawing[4].frame, drawing[4].rotation, action.dw !== 0 ? "east" : "south", action.dw, action.dh);
                        if (transaction && !transaction.commit(next)) setDrawingError(true);
                      }}
                      style={{
                        cursor: drawing[4] && !drawing[4].disabledReason ? "move" : undefined,
                        touchAction: drawing[4] && !drawing[4].disabledReason ? "none" : undefined,
                        left:
                          dataLeft + (colBand.offset + clip[0] - colBand.start) * paperScale,
                        top: dataTop + (rowBand.offset + clip[1] - rowBand.start) * paperScale,
                        width: (clip[2] - clip[0]) * paperScale,
                        height: (clip[3] - clip[1]) * paperScale,
                        // The box (frame plus bleed) is what a press hits;
                        // what the drawing paints past it -- words a text box
                        // lets overflow it (`vertOverflow` "overflow", P20's
                        // second line under its frameless box, as Excel
                        // prints it) -- is cut only by the page's own band.
                        clipPath: `inset(${-(clip[1] - rowBand.start) * paperScale}px ${-(colBand.end - clip[2]) * paperScale}px ${-(rowBand.end - clip[3]) * paperScale}px ${-(clip[0] - colBand.start) * paperScale}px)`,
                      }}
                    >
                      <div
                        className="absolute origin-top-left"
                        style={{
                          left: -(clip[0] - left) * paperScale,
                          top: -(clip[1] - top) * paperScale,
                          width,
                          height,
                          transform: `scale(${scale * paperScale})`,
                        }}
                      >
                        {content}
                      </div>
                    </div>
                  );
                })))}
                {drawings.flatMap((drawing) => {
                  const metadata = drawing[4];
                  if (!metadata || (drawing[0] !== selectedDrawingId && !selectedDrawingIds?.includes(drawing[0]) && drawing[0] !== drawingPreview?.id && (!selectedGroup || metadata.groupId !== selectedGroup))) return [];
                  const frame = groupPreview.get(drawing[0]) ?? (drawingPreview?.id === drawing[0] ? drawingPreview.frame : metadata.frame);
                  const scale = pxToPt(geometry.contentScale);
                  const pixels = scale * paperScale;
                  const angle = metadata.rotation * Math.PI / 180;
                  const width = Math.abs(frame.width * Math.cos(angle)) + Math.abs(frame.height * Math.sin(angle));
                  const height = Math.abs(frame.width * Math.sin(angle)) + Math.abs(frame.height * Math.cos(angle));
                  const left = (frame.left + (frame.width - width) / 2) * scale;
                  const top = (frame.top + (frame.height - height) / 2) * scale;
                  return drawingRowBands.flatMap((rowBand, ri) => drawingColumnBands.flatMap((colBand, ci) => {
                    const clipLeft = Math.max(left, colBand.start), clipTop = Math.max(top, rowBand.start);
                    const clipRight = Math.min(left + width * scale, colBand.end), clipBottom = Math.min(top + height * scale, rowBand.end);
                    if (clipLeft >= clipRight || clipTop >= clipBottom) return [];
                    return [<div key={`outline:${drawing[0]}:${ri}:${ci}`} aria-hidden data-page-layout-drawing-outline={drawing[0]} className="pointer-events-none absolute z-[7] overflow-hidden"
                      style={{ left: dataLeft + (colBand.offset + clipLeft - colBand.start) * paperScale, top: dataTop + (rowBand.offset + clipTop - rowBand.start) * paperScale, width: (clipRight - clipLeft) * paperScale, height: (clipBottom - clipTop) * paperScale }}>
                      <div
                        className={metadata.selectionType ? officeDrawingSelectionStyles.frame : undefined}
                        data-object-type={metadata.selectionType}
                        style={{ position: "absolute", left: (frame.left * scale - clipLeft) * paperScale, top: (frame.top * scale - clipTop) * paperScale, width: frame.width * pixels, height: frame.height * pixels, transform: `rotate(${metadata.rotation}deg)`, boxShadow: metadata.selectionType ? undefined : "inset 0 0 0 2px #107c41", background: !metadata.selectionType && drawingPreview ? "rgb(16 124 65 / 8%)" : undefined }} />
                    </div>, ...(!metadata.disabledReason && drawing[0] === (drawingPreview?.id ?? selectedDrawingId) ? PAGE_LAYOUT_DRAWING_HANDLES.flatMap((handle) => {
                      const point = pageLayoutDrawingHandlePoint(frame, handle, metadata.rotation);
                      const x = point.x * scale, y = point.y * scale;
                      if (x < colBand.start || x > colBand.end || y < rowBand.start || y > rowBand.end) return [];
                      return [<button key={`handle:${drawing[0]}:${ri}:${ci}:${handle}`} type="button"
                        className={metadata.selectionType ? officeDrawingSelectionStyles.resize : undefined}
                        data-office-drawing-handle={metadata.selectionType ? SHEET_DRAWING_HANDLE[handle] : undefined}
                        data-object-type={metadata.selectionType}
                        data-page-layout-resize={handle} data-drawing-id={drawing[0]}
                        aria-label={t("spreadsheet.helperCopy.pageLayout.resizeDrawing", { label: drawing[2], handle })}
                        onPointerDown={(event) => beginDrawing(event, drawing, handle)}
                        onClick={(event) => event.stopPropagation()}
                        onKeyDown={(event) => {
                          const delta = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[event.key];
                          if (!delta) return;
                          event.preventDefault(); event.stopPropagation();
                          const step = event.shiftKey ? 10 : 1;
                          const transaction = onBeginDrawingTransform?.(drawing[0]);
                          const aspect = metadata.aspectLocked && metadata.frame.height > 0 ? metadata.frame.width / metadata.frame.height : null;
                          if (transaction && !transaction.commit(transformPageLayoutDrawing(metadata.frame, metadata.rotation, handle, delta[0] * step, delta[1] * step, metadata.constrainSize, aspect))) setDrawingError(true);
                        }}
                        style={{ position: "absolute", zIndex: 8, left: dataLeft + (colBand.offset + x - colBand.start) * paperScale - (metadata.selectionType ? 10 : 6), top: dataTop + (rowBand.offset + y - rowBand.start) * paperScale - (metadata.selectionType ? 10 : 6),
                          ...(metadata.selectionType ? { width: 20, height: 20, "--office-drawing-grip-size":"20px", right:"auto", bottom:"auto" } as CSSProperties : { width: 12, height: 12, minWidth: 12, minHeight: 12, padding: 0, border: "1px solid #107c41", borderRadius: 2, background: "white" }),
                          touchAction: "none", cursor: `${handle === "north" || handle === "south" ? "ns" : handle === "east" || handle === "west" ? "ew" : handle === "north-west" || handle === "south-east" ? "nwse" : "nesw"}-resize` }} />];
                    }) : [])];
                  }));
                })}
                <HeaderFooterBand
                  printMode={printMode}
                  fontFamily={headerFooterFontFamily}
                  key={`${sheetId}-footer-${rowPageIndex}.${colPageIndex}`}
                  kind="footer"
                  value={variant.footer}
                  context={context}
                  paperScale={paperScale}
                  documentScale={documentScale}
                  translate={t}
                  disabledReason={headerFooterDisabledReason}
                  onCommit={(align, text) => onCommitHeaderFooter(
                    variantKey === "first" ? "firstFooter" : variantKey === "even" ? "evenFooter" : "footer", align, text,
                  )}
                  placeholder={t("spreadsheet.helperCopy.pageLayout.addFooter")}
                  style={{
                    bottom: geometry.footerMargin * paperScale,
                    left: bandLeft,
                    right: bandRight,
                    minHeight:
                      Math.max(
                        0,
                        geometry.margins.bottom - geometry.footerMargin,
                      ) * paperScale,
                  }}
                />
                <span
                  aria-hidden="true"
                  style={{
                    position: "absolute",
                    right: 8,
                    bottom: 5,
                    color: "#94a3b8",
                    fontSize: 9,
                    fontVariantNumeric: "tabular-nums",
                  }}
                >
                  {blankPage || printMode ? "" : printPageNumber}
                </span>
              </article>
            </PageLayoutPageTile>
          );
          })();
          nextPageElements.set(cacheKey, { deps: cacheDeps, element: pageElement });
          return pageElement;
        })}
        {renderPlan.pageLimitReached || renderPlan.cellLimitReached ? (
          <p
            data-page-layout-limit=""
            role="status"
            data-testid="spreadsheet-page-layout-page-limit"
            className="m-0 rounded-md bg-white px-3 py-2 text-xs text-slate-700 shadow-sm"
          >
            {t("spreadsheet.helperCopy.pageLayout.displayLimit", { renderedPageCount, renderedCellCount: renderPlan.renderedCellCount, totalPages: contentGeometry.totalPages })}
          </p>
        ) : null}
      </div>
    </div>
  );
}
