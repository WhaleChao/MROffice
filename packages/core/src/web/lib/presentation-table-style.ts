/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation slide TABLE cell formatting (PowerPoint / WPS parity subset):
 * per-cell background fill, horizontal text alignment, text formatting, and
 * text colour, plus table-level thin cell borders and a header-row emphasis
 * flag.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the present-mode overlay so the validated bounds + default
 * resolution + render-ready CSS live in exactly one place. Keep this module
 * dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * imported from the zod schema.
 *
 * Mirrors `presentation-shapes.ts` / `presentation-image-effects.ts`: a
 * six-digit-hex colour validator (a `url()` / `;}selector{` payload collapses to
 * the fallback and can never reach a style attribute raw), a whitelist for the
 * alignment enum, a normalizer that drops unknown / stale / hand-edited values,
 * and a CSS resolver. A cell with NO formatting resolves to an empty descriptor
 * and `persistedCellStyle` returns `{}`, so an un-styled table serializes and
 * renders byte-for-byte as before.
 */

import type { CSSProperties } from "react";
import { OFFICE_SHADOW_ALIGNMENTS, officeShadowCss, type OfficeShadowAlignment } from "./office-shadow-effects";

import {
  normalizeOfficeChartFillPattern,
  officeFillPatternCss,
  type OfficeChartFillPattern,
} from "./office-chart-shape-paint";
import {
  normalizeOfficeShapeGradient,
  officeGradientCss,
  type OfficeShapeGradient,
} from "./office-shape-paint";
import { normalizeOfficeShapeFontFamily } from "./office-shape-text-font";
import { isDrawableImageSource } from "./inline-image-source";
import { bgPictureLayer, normalizeBgPicture, type BgPicture } from "./presentation-slide-bg";
import {
  normalizeOfficeTextPaint,
  officeTextPaintCss,
  type OfficeTextPaint,
} from "./office-text-paint";
import {
  normalizePresentationUnderlineStyle,
  presentationTextDecorationCss,
  type PresentationUnderlineStyle,
} from "./presentation-text-decoration";
import {
  DEFAULT_VERTICAL_ANCHOR,
  normalizeVerticalAnchor,
  persistedVerticalAnchor,
  type VerticalAnchor,
} from "./presentation-vertical-anchor";

/** Six-digit hex colour, e.g. `#1a2b3c`. Matches the shape / image colour regex. */
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;







export const CELL_ALIGNMENTS = ["left", "center", "right", "justify"] as const;
/**
 * The horizontal alignments PowerPoint's Layout > Alignment group offers: Align
 * Left, Center, Align Right. Justify is not there (it is on Home > Paragraph),
 * PowerPoint for Mac 16.86 (PPT-A145).
 */
export const TABLE_LAYOUT_CELL_ALIGNMENTS = ["left", "center", "right"] as const;
/**
 * The rest of `ST_TextAlignType` an imported cell may state: Distributed
 * (`dist`), Thai Distributed (`thaiDist`) and Justify Low (`justLow`) --
 * PowerPoint's Paragraph dialog alignments. Kept and drawn (justified, the
 * last line too where distributed), never offered by the four ribbon buttons.
 */
export const NATIVE_CELL_ALIGNMENTS = [
  ...CELL_ALIGNMENTS,
  "distributed",
  "thaiDistributed",
  "justifyLow",
] as const;
export type CellAlignment = (typeof NATIVE_CELL_ALIGNMENTS)[number];

/** A cell alignment as CSS `text-align` (+ `text-align-last` when distributed). */
export function tableCellAlignmentCss(
  align: CellAlignment,
): Pick<CSSProperties, "textAlign" | "textAlignLast"> {
  if (align === "distributed" || align === "thaiDistributed") {
    return { textAlign: "justify", textAlignLast: "justify" };
  }
  return { textAlign: align === "justifyLow" ? "justify" : align };
}

/** The alignment used when a cell omits (or mis-specifies) a value. */
export const DEFAULT_CELL_ALIGNMENT: CellAlignment = "left";

/**
 * Common PowerPoint / WPS table-cell typefaces. The persisted value is the
 * actual DrawingML typeface (not a CSS stack), while the renderer resolves it
 * through the safe, complete stack below. Keeping the editable set bounded
 * prevents hand-edited page JSON from injecting arbitrary CSS.
 */
export const TABLE_CELL_FONT_FAMILIES = [
  {
    typeface: "Aptos",
    label: "Aptos",
    css: "Aptos, Calibri, 'Segoe UI', sans-serif",
  },
  {
    typeface: "Calibri",
    label: "Calibri",
    css: "Calibri, 'Segoe UI', Candara, sans-serif",
  },
  {
    typeface: "Arial",
    label: "Arial",
    css: "Arial, 'Helvetica Neue', Helvetica, sans-serif",
  },
  {
    typeface: "Cambria",
    label: "Cambria",
    css: "Cambria, Georgia, 'Times New Roman', serif",
  },
  {
    typeface: "Times New Roman",
    label: "Times New Roman",
    css: "'Times New Roman', Times, serif",
  },
  {
    typeface: "Georgia",
    label: "Georgia",
    css: "Georgia, Cambria, 'Times New Roman', serif",
  },
  {
    typeface: "Courier New",
    label: "Courier New",
    css: "'Courier New', Courier, monospace",
  },
  {
    typeface: "Consolas",
    label: "Consolas",
    css: "Consolas, 'SFMono-Regular', Menlo, monospace",
  },
  {
    typeface: "DengXian",
    label: "DengXian (等线)",
    css: "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif",
  },
  {
    typeface: "Microsoft YaHei",
    label: "Microsoft YaHei (微软雅黑)",
    css: "'Microsoft YaHei', 'PingFang SC', sans-serif",
  },
  {
    typeface: "SimSun",
    label: "SimSun (宋体)",
    css: "SimSun, 'Songti SC', serif",
  },
  {
    typeface: "PingFang SC",
    label: "PingFang SC (苹方)",
    css: "'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif",
  },
  {
    typeface: "Microsoft JhengHei",
    label: "Microsoft JhengHei (微軟正黑體)",
    css: "'Microsoft JhengHei', 'PingFang TC', sans-serif",
  },
  {
    typeface: "Yu Gothic",
    label: "Yu Gothic (游ゴシック)",
    css: "'Yu Gothic', 'Hiragino Kaku Gothic ProN', sans-serif",
  },
  {
    typeface: "Malgun Gothic",
    label: "Malgun Gothic (맑은 고딕)",
    css: "'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif",
  },
] as const;
/**
 * A cell's DrawingML typeface. The listed faces carry a curated browser stack;
 * any other face an imported deck names (Tahoma, Cordia New, B Zar, David ...,
 * which PowerPoint draws as written) is kept by name, cleaned the way the
 * shared Office font pickers clean one (`normalizeOfficeShapeFontFamily`).
 */
export type TableCellFontFamily =
  | (typeof TABLE_CELL_FONT_FAMILIES)[number]["typeface"]
  | "Franklin Gothic Medium"
  | "Franklin Gothic Book"
  | (string & {});
const TABLE_CELL_FONT_FAMILY_BY_TYPEFACE = new Map<
  string,
  (typeof TABLE_CELL_FONT_FAMILIES)[number]
>(TABLE_CELL_FONT_FAMILIES.map((option) => [option.typeface, option]));
const TABLE_SOURCE_FONT_FAMILY_CSS = new Map<string, string>([
  [
    "Franklin Gothic Medium",
    "'Franklin Gothic Medium', 'Arial Narrow', Arial, sans-serif",
  ],
  [
    "Franklin Gothic Book",
    "'Franklin Gothic Book', 'Arial Narrow', Arial, sans-serif",
  ],
]);

/** Font size is persisted in points, matching PowerPoint/WPS and DrawingML. */
export const MIN_TABLE_CELL_FONT_SIZE = 1;
export const MAX_TABLE_CELL_FONT_SIZE = 4000;

export function normalizeTableCellFontFamily(
  value: unknown,
): TableCellFontFamily | null {
  const family = normalizeOfficeShapeFontFamily(value);
  // A theme reference (`+mn-lt`) is not a face, and no face is named with
  // CSS or string punctuation.
  if (!family || family.startsWith("+") || /[;{}"\\]/.test(family)) return null;
  return family;
}

/** CSS's generic families: a stack's fallback, never a face a cell names. */
const GENERIC_CSS_FAMILIES = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui",
  "ui-serif", "ui-sans-serif", "ui-monospace", "ui-rounded", "math", "emoji",
  "fangsong", "-apple-system",
]);









export function tableCellTypefaceFromChoice(value: unknown): TableCellFontFamily | null {
  if (typeof value !== "string") return null;
  const first = (value.split(",")[0] ?? value).replace(/["']/g, "").trim();
  if (GENERIC_CSS_FAMILIES.has(first.toLowerCase())) return null;
  return normalizeTableCellFontFamily(first);
}

export function tableCellFontFamilyCss(value: unknown): string | undefined {
  // A browser stack stored as the cell's face (an older Replace Fonts) names
  
  const family =
    typeof value === "string" && value.includes(",")
      ? tableCellTypefaceFromChoice(value)
      : normalizeTableCellFontFamily(value);
  return family
    ? (TABLE_CELL_FONT_FAMILY_BY_TYPEFACE.get(family)?.css ??
        TABLE_SOURCE_FONT_FAMILY_CSS.get(family) ??
        // Quoted, so a name is never read as CSS; the page's Office font
        // layer stands a substitute in for a face the browser lacks.
        `"${family.replace(/["\\]/g, "")}", sans-serif`)
    : undefined;
}

export function normalizeTableCellFontSize(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  if (value < MIN_TABLE_CELL_FONT_SIZE || value > MAX_TABLE_CELL_FONT_SIZE) {
    return null;
  }
  // Hundredths: the file's own unit (`a:rPr@sz`, ST_TextFontSize, 100ths of
  // a point). A size the deck states keeps its precision -- a master's 16.33
  // pt cells read back as 16.3 on every open and the untouched save could not
  
  // What a person types goes through `tableCellFontSizeFromInput`.
  return Math.round(value * 100) / 100;
}

/**
 * A size typed in a cell's size box: tenths, as PowerPoint's size box takes
 * them for any text (half points turned 10.3 into 10.5 in a cell only).
 */
export function tableCellFontSizeFromInput(value: unknown): number | null {
  const size = normalizeTableCellFontSize(value);
  return size === null ? null : normalizeTableCellFontSize(Math.round(size * 10) / 10);
}

export const TABLE_CELL_TEXT_DIRECTIONS = [
  "horizontal",
  "rotate-90",
  "rotate-270",
  "stacked",
] as const;
/**
 * The other vertical kinds of `ST_TextVerticalType` an imported cell may
 * state -- East Asian vertical (`eaVert`), Mongolian (`mongolianVert`),
 * stacked right to left (`wordArtVertRtl`) -- kept and drawn, never offered
 * by Text Direction's four.
 */
export const NATIVE_TABLE_CELL_TEXT_DIRECTIONS = [
  ...TABLE_CELL_TEXT_DIRECTIONS,
  "eastAsianVertical",
  "mongolianVertical",
  "stackedRightToLeft",
] as const;
export type TableCellTextDirection =
  (typeof NATIVE_TABLE_CELL_TEXT_DIRECTIONS)[number];
const TABLE_CELL_TEXT_DIRECTION_SET: ReadonlySet<string> = new Set(
  NATIVE_TABLE_CELL_TEXT_DIRECTIONS,
);
export const DEFAULT_TABLE_CELL_TEXT_DIRECTION: TableCellTextDirection =
  "horizontal";

export const TABLE_CELL_BORDER_STYLES = [
  "none",
  "solid",
  "dash",
  "dot",
  "double",
] as const;
export type TableCellBorderStyle = (typeof TABLE_CELL_BORDER_STYLES)[number];
const TABLE_CELL_BORDER_STYLE_SET: ReadonlySet<string> = new Set(
  TABLE_CELL_BORDER_STYLES,
);
export const DEFAULT_TABLE_CELL_BORDER_COLOR = "#cbd5e1";
export const DEFAULT_TABLE_CELL_BORDER_WIDTH = 1;
/** The Web quarter-point option expressed in the model's CSS-pixel unit. */
export const MIN_TABLE_CELL_BORDER_WIDTH = 1 / 3;
export const MAX_TABLE_CELL_BORDER_WIDTH = 12;
/** The widest line a file can state (`ST_LineWidth` 20 116 800 EMU, 1584 pt), in CSS px. */
const MAX_FILE_TABLE_CELL_BORDER_WIDTH = 2112;
export const MAX_TABLE_CELL_MARGIN = 96;
export const TABLE_CELL_MARGIN_PRESETS = [
  "normal",
  "none",
  "narrow",
  "wide",
] as const;
export type TableCellMarginPreset = (typeof TABLE_CELL_MARGIN_PRESETS)[number];

/** A table edge of several rules (`a:ln@cmpd`, readings row 476). */
export const TABLE_CELL_BORDER_COMPOUNDS = ["thickThin", "thinThick", "tri"] as const;
export type TableCellBorderCompound = (typeof TABLE_CELL_BORDER_COMPOUNDS)[number];

export interface TableCellBorder {
  style: TableCellBorderStyle;
  color: string;
  width: number;
   
  compound?: TableCellBorderCompound;
}
















export const TABLE_EDGE_COMPOUND_RULES: Record<TableCellBorderCompound, readonly (readonly [number, number])[]> = {
  thickThin: [[0, 3 / 5], [4 / 5, 1]],
  thinThick: [[0, 1 / 5], [2 / 5, 1]],
  tri: [[0, 1 / 6], [2 / 6, 4 / 6], [5 / 6, 1]],
};

type TableEdgeSide = "top" | "bottom" | "left" | "right";

/** The custom properties a `<td>` style states a compound edge with
 * (`tableCellCompoundEdges` draws them). */
export function tableEdgeCompoundVar(side: TableEdgeSide): string {
  return `--presentation-cell-compound-${side}`;
}

/**
 * A side's line as `<td>` CSS: the border itself, and, for a compound edge,
 * the rules `tableCellCompoundEdges` draws over its transparent width (a CSS
 * border draws no rules of unequal widths). They are drawn only while the
 * side's border is still that transparent one: a cell's own plain line over
 * its style's compound one replaces the border and so the rules.
 */
export function tableEdgeCss(side: TableEdgeSide, border: TableCellBorder): CSSProperties {
  const name = `border${side[0].toUpperCase()}${side.slice(1)}` as "borderTop";
  const rules = border.compound && border.style !== "none" ? TABLE_EDGE_COMPOUND_RULES[border.compound] : null;
  if (!rules) {
    return { [name]: tableStyleBorderCss(border) } as CSSProperties;
  }
  const width = Math.max(border.width, 3);
  const direction = side === "top" || side === "bottom" ? "to bottom" : "to left";
  const percent = (fraction: number) => `${Number((fraction * 100).toFixed(4))}%`;
  const stops = rules.flatMap(([from, to]) => [
    `transparent ${percent(from)}`,
    `${border.color} ${percent(from)}`,
    `${border.color} ${percent(to)}`,
    `transparent ${percent(to)}`,
  ]);
  return {
    [name]: `${width}px solid transparent`,
    position: "relative",
    [tableEdgeCompoundVar(side)]: `${width}px|linear-gradient(${direction}, ${stops.join(", ")})`,
  } as CSSProperties;
}

/**
 * WPS 表格样式 → 效果 → 单元格凹凸效果, PowerPoint "Cell Bevel". The vocabulary is
 * DrawingML's own `ST_BevelPresetType`, so a projected cell keeps the exact
 * preset the deck declared instead of being flattened to "some bevel".
 */
export const TABLE_CELL_BEVELS = [
  "relaxedInset",
  "circle",
  "slope",
  "cross",
  "angle",
  "softRound",
  "convex",
  "coolSlant",
  "divot",
  "riblet",
  "hardEdge",
  "artDeco",
] as const;
export type TableCellBevelPreset = (typeof TABLE_CELL_BEVELS)[number];
const TABLE_CELL_BEVEL_SET: ReadonlySet<string> = new Set(TABLE_CELL_BEVELS);

/** DrawingML's own default bevel size, in points (76200 EMU). */
export const DEFAULT_TABLE_CELL_BEVEL_PT = 6;
export const MAX_TABLE_CELL_BEVEL_PT = 72;

export interface TableCellBevel {
  preset: TableCellBevelPreset;
  /** Bevel width in points; absent ⇒ DrawingML's default. */
  width?: number;
  /** Bevel height in points; absent ⇒ DrawingML's default. */
  height?: number;
}

function normalizeBevelSize(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return undefined;
  }
  return Math.min(MAX_TABLE_CELL_BEVEL_PT, Math.round(value * 10) / 10);
}

/** Validate a cell bevel, or `null` when it names no known preset. */
export function normalizeTableCellBevel(value: unknown): TableCellBevel | null {
  if (typeof value === "string") {
    return TABLE_CELL_BEVEL_SET.has(value)
      ? { preset: value as TableCellBevelPreset }
      : null;
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const preset = record.preset;
  if (typeof preset !== "string" || !TABLE_CELL_BEVEL_SET.has(preset)) {
    return null;
  }
  const bevel: TableCellBevel = { preset: preset as TableCellBevelPreset };
  const width = normalizeBevelSize(record.width);
  const height = normalizeBevelSize(record.height);
  if (width !== undefined) bevel.width = width;
  if (height !== undefined) bevel.height = height;
  return bevel;
}

/**
 * A bevel as CSS. DrawingML lights a cell in 3-D; the browser has no equivalent,
 * so this paints the SHADING a bevel produces — a light inner edge top-left and
 * a dark one bottom-right — scaled by the declared size. Concave presets invert
 * the pair so a divot reads as pressed in rather than raised.
 *
 * This is openly an approximation of a 3-D material, not a reproduction. The
 * export keeps the native `a:cell3D`, so PowerPoint and WPS still render the
 * real thing; only the editor's preview is stylized.
 */
export function tableCellBevelCss(bevel: TableCellBevel | null): string | null {
  if (!bevel) return null;
  // Guard the arithmetic rather than trusting the caller: this is exported, and
  // a non-finite size would produce `inset NaNpx …`, which is not CSS at all.
  const finite = (value: unknown): number =>
    typeof value === "number" && Number.isFinite(value) && value > 0
      ? value
      : DEFAULT_TABLE_CELL_BEVEL_PT;
  const size = Math.max(
    1,
    Math.min(
      8,
      Math.round((finite(bevel.width) + finite(bevel.height)) / 2 / 1.5),
    ),
  );
  const concave = bevel.preset === "divot" || bevel.preset === "coolSlant";
  const light = "rgba(255, 255, 255, 0.72)";
  const dark = "rgba(15, 23, 42, 0.26)";
  const first = concave ? dark : light;
  const second = concave ? light : dark;
  return `inset ${size}px ${size}px ${size * 1.5}px -${size}px ${first}, inset -${size}px -${size}px ${size * 1.5}px -${size}px ${second}`;
}

export const TABLE_BORDER_TARGETS = [
  "all",
  "outside",
  "inside",
  "inside-horizontal",
  "inside-vertical",
  "top",
  "bottom",
  "left",
  "right",
  "diagonal-down",
  "diagonal-up",
  "none",
] as const;
export type TableBorderTarget = (typeof TABLE_BORDER_TARGETS)[number];

export interface TableBorderSelectionPosition {
  row: number;
  col: number;
  row1: number;
  row2: number;
  col1: number;
  col2: number;
  /** A merged cell's span; its far edges are where it ends, not where it starts. */
  rowSpan?: number;
  colSpan?: number;
}

/**
 * Resolve the native per-edge patch for one cell inside a rectangular
 * PowerPoint/WPS border command. Internal borders are written on both adjacent
 * cells so the OOXML remains editable even when PowerPoint does not collapse
 * neighbouring DrawingML edges.
 */
export function tableBorderPatchForSelection(
  target: TableBorderTarget,
  border: TableCellBorder,
  position: TableBorderSelectionPosition,
): Partial<CellStyleInput> {
  const { row, col, row1, row2, col1, col2 } = position;
  const rowEnd = row + Math.max(1, Math.floor(position.rowSpan ?? 1)) - 1;
  const colEnd = col + Math.max(1, Math.floor(position.colSpan ?? 1)) - 1;
  // A merged cell touches the selection's right / bottom edge when it ENDS
  // there (PPT-B062: a merged top row lost its right outside border).
  const top = row <= row1;
  const bottom = rowEnd >= row2;
  const left = col <= col1;
  const right = colEnd >= col2;
  const result: Partial<CellStyleInput> = {};
  const setEdge = (
    edge:
      | "borderLeft"
      | "borderRight"
      | "borderTop"
      | "borderBottom"
      | "borderDiagonalDown"
      | "borderDiagonalUp",
  ) => {
    result[edge] = border;
  };
  if (target === "none") {
    const none = { ...border, style: "none" as const };
    return {
      borderLeft: none,
      borderRight: none,
      borderTop: none,
      borderBottom: none,
      borderDiagonalDown: none,
      borderDiagonalUp: none,
    };
  }
  if (target === "all") {
    return {
      borderLeft: border,
      borderRight: border,
      borderTop: border,
      borderBottom: border,
    };
  }
  if (target === "outside") {
    if (left) setEdge("borderLeft");
    if (right) setEdge("borderRight");
    if (top) setEdge("borderTop");
    if (bottom) setEdge("borderBottom");
    return result;
  }
  if (target === "inside" || target === "inside-horizontal") {
    if (!top) setEdge("borderTop");
    if (!bottom) setEdge("borderBottom");
  }
  if (target === "inside" || target === "inside-vertical") {
    if (!left) setEdge("borderLeft");
    if (!right) setEdge("borderRight");
  }
  if (target === "top" && top) setEdge("borderTop");
  if (target === "bottom" && bottom) setEdge("borderBottom");
  if (target === "left" && left) setEdge("borderLeft");
  if (target === "right" && right) setEdge("borderRight");
  if (target === "diagonal-down") setEdge("borderDiagonalDown");
  if (target === "diagonal-up") setEdge("borderDiagonalUp");
  return result;
}

/**
 * Header-row default emphasis applied by the table-level `headerRow` flag: a
 * light fill and bold text on the first row. Resolved at render time, never
 * persisted onto individual cells, so toggling the flag stays a one-bit change.
 */
export const HEADER_ROW_FILL = "#e2e8f0";

/**
 * Subtle fill for an emphasised column — the first / last column, or an
 * alternate banded column (PowerPoint First Column / Last Column / Banded
 * Columns). A single fill mirrors the single banded-row fill; resolved at
 * render time and layered UNDER any explicit per-cell colour by `cellStyleCss`,
 * never persisted onto individual cells.
 */
export const COLUMN_EMPHASIS_FILL = "#f1f5f9";

/** The thin cell-border style used when the table-level `borders` flag is set. */
export const TABLE_BORDER_CSS = "1px solid #cbd5e1";

/** The persisted per-cell style field names handled by the normalizers. */
export const CELL_STYLE_KEYS = [
  "fill",
  "align",
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
  "highlight",
  "color",
  "fontFamily",
  "fontSize",
  "verticalAlign",
  "textDirection",
  "marginLeft",
  "marginRight",
  "marginTop",
  "marginBottom",
  "borderLeft",
  "borderRight",
  "borderTop",
  "borderBottom",
  "borderDiagonalDown",
  "borderDiagonalUp",
  "bevel",
  "fillGradient",
  "fillPattern",
  // Table Design > Shading > Picture (`a:tcPr/a:blipFill`): the slide
  // background's picture model, drawn while the cell states no colour.
  "fillImage",
  // The kind of an imported underline beyond the single line (double, heavy,
  // dotted, dashed, wavy: `a:rPr@u`), kept while `underline` is on.
  "underlineStyle",
  // Font > Double Strikethrough (a:rPr@strike dblStrike), kept while `strike`.
  "strikeStyle",
  // Table Design > WordArt Styles (PPT-A138): the shared WordArt paint on
  // every run of the cell (`lib/office-text-paint`).
  "textPaint",
] as const;

/** Native DrawingML table style GUID (`a:tableStyleId`). */
const TABLE_STYLE_ID_RE =
  /^\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}$/;

/**
 * Built-in PowerPoint / WPS table-style identities. The first two and the
 * Accent 1 / Accent 2 values are captured from real WPS-authored PPTX fixtures;
 * the remaining Accent variants are stable Microsoft Office built-ins.
 */
export const NATIVE_TABLE_STYLE_IDS = {
  noStyleNoGrid: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
  noStyleTableGrid: "{5940675A-B579-460E-94D1-54222C63F5DA}",
  mediumAccent1: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
  mediumAccent2: "{284E427A-3D55-4303-BF80-6455036E1DE7}",
  mediumAccent3: "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
  mediumAccent6: "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
} as const;

export type NativeTableStyleId =
  (typeof NATIVE_TABLE_STYLE_IDS)[keyof typeof NATIVE_TABLE_STYLE_IDS];

export function normalizeTableStyleId(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return TABLE_STYLE_ID_RE.test(trimmed) ? trimmed.toUpperCase() : null;
}

/**
 * WPS / PowerPoint-style quick table looks shown in the contextual Table Tools
 * ribbon. Choosing a look writes the native `a:tableStyleId` plus editable
 * table-style option flags. No table is flattened and no style is rasterised.
 */
export const TABLE_STYLE_PRESET_IDS = [
  "plain",
  "grid-grey",
  "banded-grey",
  "banded-blue",
  "banded-green",
  "header-blue",
] as const;
export type TableStylePresetId = (typeof TABLE_STYLE_PRESET_IDS)[number];
const TABLE_STYLE_PRESET_ID_SET: ReadonlySet<string> = new Set(
  TABLE_STYLE_PRESET_IDS,
);

export interface TableStylePreset {
  id: TableStylePresetId;
  label: string;
  tableStyleId: NativeTableStyleId;
  borders?: true;
  headerRow?: true;
  bandedRows?: true;
  headerFill?: string;
  headerColor?: string;
  headerBold?: true;
  bandFill?: string;
  bodyFill?: string;
  borderColor?: string;
}

export const TABLE_STYLE_PRESETS: readonly TableStylePreset[] = [
  {
    id: "plain",
    label: "无样式 / No style",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.noStyleNoGrid,
    bodyFill: "#ffffff",
  },
  {
    id: "grid-grey",
    label: "清除样式·网格 / Clear style",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.noStyleTableGrid,
    borders: true,
    headerFill: "#ffffff",
    headerColor: "#0f172a",
    bodyFill: "#ffffff",
    borderColor: "#64748b",
  },
  {
    id: "banded-grey",
    label: "强调 3 / Accent 3",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.mediumAccent3,
    borders: true,
    headerRow: true,
    bandedRows: true,
    headerFill: "#475569",
    headerColor: "#ffffff",
    headerBold: true,
    bandFill: "#e2e8f0",
    borderColor: "#ffffff",
  },
  {
    id: "banded-blue",
    label: "强调 1 / Accent 1",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.mediumAccent1,
    borders: true,
    headerRow: true,
    bandedRows: true,
    headerFill: "#0f6cbd",
    headerColor: "#ffffff",
    headerBold: true,
    bandFill: "#dbeafe",
    borderColor: "#ffffff",
  },
  {
    id: "banded-green",
    label: "强调 6 / Accent 6",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.mediumAccent6,
    borders: true,
    headerRow: true,
    bandedRows: true,
    headerFill: "#0f766e",
    headerColor: "#ffffff",
    headerBold: true,
    bandFill: "#d1fae5",
    borderColor: "#ffffff",
  },
  {
    id: "header-blue",
    label: "强调 2 / Accent 2",
    tableStyleId: NATIVE_TABLE_STYLE_IDS.mediumAccent2,
    borders: true,
    headerRow: true,
    bandedRows: true,
    headerFill: "#ed7d31",
    headerColor: "#ffffff",
    headerBold: true,
    bandFill: "#fee2d5",
    borderColor: "#ffffff",
  },
] as const;

export function normalizeTableStylePresetId(
  value: unknown,
): TableStylePresetId {
  return typeof value === "string" && TABLE_STYLE_PRESET_ID_SET.has(value)
    ? (value as TableStylePresetId)
    : "plain";
}

export interface TableStylePresetCell {
  text: string;
  [key: string]: unknown;
}

export interface TableStylePresetTable {
  rows: TableStylePresetCell[][];
  [key: string]: unknown;
}

/**
 * Apply one gallery look as a native style identity plus table-style option
 * flags. Existing cells, direct formatting, merges and dimensions are untouched:
 * Office renders the style beneath direct formatting and keeps both editable.
 */
export function applyTableStylePreset<T extends TableStylePresetTable>(
  table: T,
  presetId: unknown,
): T & {
  tableStyleId: NativeTableStyleId;
  headerRow?: true;
  bandedRows?: true;
} {
  const preset =
    TABLE_STYLE_PRESETS.find(
      (candidate) => candidate.id === normalizeTableStylePresetId(presetId),
    ) ?? TABLE_STYLE_PRESETS[0];
  const next = { ...table } as Record<string, unknown>;
  next.tableStyleId = preset.tableStyleId;
  // The preset's grid is its own: one without (Plain, "No Style, No Grid")
  
  if (preset.borders) next.borders = true;
  else delete next.borders;
  if (preset.headerRow) next.headerRow = true;
  if (preset.bandedRows) next.bandedRows = true;
  return next as T & {
    tableStyleId: NativeTableStyleId;
    headerRow?: true;
    bandedRows?: true;
  };
}

/**
 * WPS "Clear Table Style" keeps the table native and switches it to the built-in
 * editable grid style; table-style options and direct cell formatting survive.
 */
export function clearNativeTableStyle<T extends TableStylePresetTable>(
  table: T,
): T & {
  tableStyleId: typeof NATIVE_TABLE_STYLE_IDS.noStyleTableGrid;
} {
  return {
    ...table,
    tableStyleId: NATIVE_TABLE_STYLE_IDS.noStyleTableGrid,
  };
}

export function isTableStylePresetActive(
  tableStyleId: unknown,
  presetId: unknown,
): boolean {
  const preset =
    TABLE_STYLE_PRESETS.find(
      (candidate) => candidate.id === normalizeTableStylePresetId(presetId),
    ) ?? TABLE_STYLE_PRESETS[0];
  return normalizeTableStyleId(tableStyleId) === preset.tableStyleId;
}

/* ---------------------------------------------------------------------- *
 * RESOLVED native table styles.
 *
 * A gallery-styled PowerPoint / WPS table carries no direct cell formatting at
 * all — the look lives in the `<a:tblStyle>` definition the authoring app copies
 * into `ppt/tableStyles.xml`. The importer resolves that definition against the
 * deck theme (see `apps/api/.../pptx/table_style_lookup.py`) and attaches the
 * result to the block as `tableStyleLook`. Rendering it here is what makes an
 * imported deck look like the deck the author saved, instead of a plain grid.
 *
 * This payload comes from a FILE, so every colour and width is re-validated
 * before it can reach a style attribute — the same contract the per-cell colour
 * validators follow. It is also render-only: nothing here is persisted onto
 * cells, so the export keeps the native GUID and the package's own
 * `tableStyles.xml` untouched.
 * ---------------------------------------------------------------------- */

export const TABLE_STYLE_LOOK_PARTS = [
  "wholeTbl",
  "band1H",
  "band2H",
  "band1V",
  "band2V",
  "firstCol",
  "lastCol",
  "firstRow",
  "lastRow",
  "nwCell",
  "neCell",
  "swCell",
  "seCell",
] as const;
export type TableStyleLookPartName = (typeof TABLE_STYLE_LOOK_PARTS)[number];

export const TABLE_STYLE_LOOK_EDGES = [
  "left",
  "right",
  "top",
  "bottom",
  "insideH",
  "insideV",
] as const;
export type TableStyleLookEdge = (typeof TABLE_STYLE_LOOK_EDGES)[number];

export interface TableStyleLookPart {
  /** A resolved fill, or `"none"` for an explicit clear. */
  fill?: string | "none";
  fillOpacity?: number;
  color?: string;
  bold?: true;
  italic?: true;
  /** The part's own face (`a:tcTxStyle/a:font`): the words of its cells
   * whose runs name none are set in it, over the theme's. */
  fontFamily?: string;
  borders?: Partial<Record<TableStyleLookEdge, TableCellBorder>>;
}

export interface TableStyleLook {
  styleId: string;
  styleName?: string;
  background?: string;
  backgroundGradient?: {
    angle: number;
    stops: Array<{ offset: number; color: string }>;
  };
  shadow?: {
    color: string;
    opacity: number;
    blur: number;
    distance: number;
    angle: number;
    /**
     * A scaled / skewed shadow (`a:outerShdw@sx/@sy/@kx/@ky`, 1/1000 % and
     * 1/60000 degree) about its alignment corner (`@algn`), as the import
     * reads a style's `a:tblBg` effect (readings row 500).
     */
    sx?: number;
    sy?: number;
    kx?: number;
    ky?: number;
    algn?: OfficeShadowAlignment;
  };
  parts: Partial<Record<TableStyleLookPartName, TableStyleLookPart>>;
}

const TABLE_STYLE_LOOK_PART_SET: ReadonlySet<string> = new Set(
  TABLE_STYLE_LOOK_PARTS,
);
const TABLE_STYLE_LOOK_EDGE_SET: ReadonlySet<string> = new Set(
  TABLE_STYLE_LOOK_EDGES,
);

/** A plain object, or `null` for anything else (arrays included). */
function asStyleRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export function nativeTableThemeColorAuthority(source: unknown): boolean {
  const receipt = asStyleRecord(
    asStyleRecord(source)?.nativeTableThemeColorReceipt,
  );
  const digestFields = [
    "themeFingerprint",
    "masterFingerprint",
    "colorMapFingerprint",
    "tableFingerprint",
    "resolvedDigest",
  ];
  return (
    receipt?.schemaVersion === "pptx-table-theme-color-v1" &&
    receipt.disposition === "native" &&
    typeof receipt.slideId === "string" &&
    /^sld[1-9][0-9]*$/.test(receipt.slideId) &&
    typeof receipt.shapeId === "string" &&
    /^shp[1-9][0-9]*$/.test(receipt.shapeId) &&
    Number.isInteger(receipt.atomCount) &&
    Number(receipt.atomCount) >= 1 &&
    Number(receipt.atomCount) <= 256 &&
    digestFields.every(
      (field) =>
        typeof receipt[field] === "string" &&
        /^[0-9a-f]{64}$/.test(receipt[field] as string),
    )
  );
}
export function nativeTableThemeTableCss(
  background: unknown,
  source: unknown,
): CSSProperties | undefined {
  if (!nativeTableThemeColorAuthority(source)) return undefined;
  const color = normalizeCellColor(background, null);
  return color ? { background: color } : undefined;
}

function normalizeLookBorder(value: unknown): TableCellBorder | null {
  const record = asStyleRecord(value);
  if (!record) return null;
  return normalizeCellStyle({ borderTop: record }).borderTop;
}

function normalizeLookPart(value: unknown): TableStyleLookPart | null {
  const record = asStyleRecord(value);
  if (!record) return null;
  const part: TableStyleLookPart = {};
  if (record.fill === "none") {
    part.fill = "none";
  } else {
    const fill = normalizeCellColor(record.fill, null);
    if (fill !== null) part.fill = fill;
  }
  const color = normalizeCellColor(record.color, null);
  if (
    typeof record.fillOpacity === "number" &&
    Number.isFinite(record.fillOpacity) &&
    record.fillOpacity >= 0 &&
    record.fillOpacity <= 1
  )
    part.fillOpacity = record.fillOpacity;
  if (color !== null) part.color = color;
  if (record.bold === true) part.bold = true;
  if (record.italic === true) part.italic = true;
  const fontFamily = normalizeTableCellFontFamily(record.fontFamily);
  if (fontFamily !== null) part.fontFamily = fontFamily;
  const borders = asStyleRecord(record.borders);
  if (borders) {
    const resolved: Partial<Record<TableStyleLookEdge, TableCellBorder>> = {};
    for (const [edge, raw] of Object.entries(borders)) {
      if (!TABLE_STYLE_LOOK_EDGE_SET.has(edge)) continue;
      const border = normalizeLookBorder(raw);
      if (border) resolved[edge as TableStyleLookEdge] = border;
    }
    if (Object.keys(resolved).length > 0) part.borders = resolved;
  }
  return Object.keys(part).length > 0 ? part : null;
}

/**
 * Validate an imported `tableStyleLook` payload. Returns `null` when it carries
 * nothing usable, so the renderer falls back to the preset gallery path and an
 * un-styled table renders exactly as before.
 *
 * Pass `expectedStyleId` to bind the look to the table's CURRENT identity: a
 * resolved look describes exactly one GUID, so once the user picks a different
 * style the old payload must stop painting. Keeping it in the deck is harmless —
 * it becomes live again if that identity comes back.
 */
export function normalizeTableStyleLook(
  value: unknown,
  expectedStyleId?: unknown,
): TableStyleLook | null {
  const record = asStyleRecord(value);
  if (!record) return null;
  if (expectedStyleId !== undefined) {
    const expected = normalizeTableStyleId(expectedStyleId);
    if (!expected || normalizeTableStyleId(record.styleId) !== expected) {
      return null;
    }
  }
  const styleId = normalizeTableStyleId(record.styleId);
  if (!styleId) return null;
  const parts: Partial<Record<TableStyleLookPartName, TableStyleLookPart>> = {};
  const rawParts = asStyleRecord(record.parts);
  if (rawParts) {
    for (const [name, raw] of Object.entries(rawParts)) {
      if (!TABLE_STYLE_LOOK_PART_SET.has(name)) continue;
      const part = normalizeLookPart(raw);
      if (part) parts[name as TableStyleLookPartName] = part;
    }
  }
  const background = normalizeCellColor(record.background, null);
  if (Object.keys(parts).length === 0 && background === null) return null;
  const look: TableStyleLook = { styleId, parts };
  if (typeof record.styleName === "string" && record.styleName.trim()) {
    look.styleName = record.styleName.trim().slice(0, 120);
  }
  if (background !== null) look.background = background;
  const gradient = asStyleRecord(record.backgroundGradient);
  if (
    gradient &&
    typeof gradient.angle === "number" &&
    Number.isFinite(gradient.angle) &&
    Math.abs(gradient.angle) <= 360 &&
    Array.isArray(gradient.stops) &&
    gradient.stops.length >= 2 &&
    gradient.stops.length <= 16
  ) {
    const stops = gradient.stops.map(asStyleRecord);
    if (
      stops.every(
        (stop) =>
          stop &&
          typeof stop.offset === "number" &&
          Number.isFinite(stop.offset) &&
          stop.offset >= 0 &&
          stop.offset <= 1 &&
          normalizeCellColor(stop.color, null),
      )
    ) {
      look.backgroundGradient = {
        angle: gradient.angle,
        stops: stops.map((stop) => ({
          offset: stop!.offset as number,
          color: normalizeCellColor(stop!.color, null)!,
        })),
      };
    }
  }
  const shadow = asStyleRecord(record.shadow);
  if (
    shadow &&
    normalizeCellColor(shadow.color, null) &&
    ["opacity", "blur", "distance", "angle"].every(
      (key) => typeof shadow[key] === "number" && Number.isFinite(shadow[key]),
    ) &&
    (shadow.opacity as number) >= 0 &&
    (shadow.opacity as number) <= 1 &&
    (shadow.blur as number) >= 0 &&
    (shadow.blur as number) <= 100 &&
    (shadow.distance as number) >= 0 &&
    (shadow.distance as number) <= 100 &&
    Math.abs(shadow.angle as number) <= 360
  ) {
    const transform: Partial<Record<"sx" | "sy" | "kx" | "ky", number>> = {};
    for (const key of ["sx", "sy", "kx", "ky"] as const) {
      const value = shadow[key];
      const limit = key === "sx" || key === "sy" ? 10_000_000 : 5_400_000;
      if (typeof value === "number" && Number.isInteger(value) && Math.abs(value) <= limit) transform[key] = value;
    }
    const algn = OFFICE_SHADOW_ALIGNMENTS.find((value) => value === shadow.algn);
    look.shadow = {
      color: normalizeCellColor(shadow.color, null)!,
      opacity: shadow.opacity as number,
      blur: shadow.blur as number,
      distance: shadow.distance as number,
      angle: shadow.angle as number,
      ...transform,
      ...(algn ? { algn } : {}),
    };
  }
  return look;
}

/**
 * A style look's background shadow in the shared shadow model
 * (`lib/office-shadow-effects`): px and degrees back to DrawingML's units.
 */
export function tableStyleLookShadowEffect(value: unknown) {
  const shadow = normalizeTableStyleLook(value)?.shadow;
  return shadow ? lookShadowEffect(shadow) : null;
}

function lookShadowEffect(shadow: NonNullable<TableStyleLook["shadow"]>) {
  return {
    blurRad: Math.round(shadow.blur * 9525),
    dist: Math.round(shadow.distance * 9525),
    dir: Math.round(shadow.angle * 60_000),
    algn: shadow.algn ?? ("b" as const),
    color: shadow.color,
    alpha: Math.round(shadow.opacity * 100_000),
    ...(shadow.sx !== undefined ? { sx: shadow.sx } : {}),
    ...(shadow.sy !== undefined ? { sy: shadow.sy } : {}),
    ...(shadow.kx !== undefined ? { kx: shadow.kx } : {}),
    ...(shadow.ky !== undefined ? { ky: shadow.ky } : {}),
  };
}

export function tableStyleLookTableCss(
  value: unknown,
  /** The table's frame in px: a scaled shadow grows with it (`officeShadowCss`). */
  size: { readonly width: number; readonly height: number } = { width: 0, height: 0 },
): CSSProperties {
  const look = normalizeTableStyleLook(value);
  if (!look) return {};
  // A scaled or skewed shadow is drawn as the shared model draws a shape's:
  // an even scale as the box shadow's spread, any other on its own layer
  // under the table (`presentationTableShadowLayers`).
  const scaled = look.shadow && (look.shadow.sx !== undefined || look.shadow.sy !== undefined ||
    look.shadow.kx !== undefined || look.shadow.ky !== undefined);
  const modelled = scaled ? officeShadowCss(lookShadowEffect(look.shadow!), size) : null;
  return {
    ...(look.background ? { backgroundColor: look.background } : {}),
    ...(look.backgroundGradient
      ? {
          backgroundImage: `linear-gradient(${look.backgroundGradient.angle + 90}deg, ${look.backgroundGradient.stops.map((stop) => `${stop.color} ${stop.offset * 100}%`).join(", ")})`,
        }
      : {}),
    ...(scaled
      ? modelled?.boxShadow ? { boxShadow: modelled.boxShadow } : {}
      : look.shadow
      ? {
          boxShadow: `${Math.round(Math.cos((look.shadow.angle * Math.PI) / 180) * look.shadow.distance * 100) / 100}px ${Math.round(Math.sin((look.shadow.angle * Math.PI) / 180) * look.shadow.distance * 100) / 100}px ${look.shadow.blur}px rgba(${[1, 3, 5].map((i) => parseInt(look.shadow!.color.slice(i, i + 2), 16)).join(", ")}, ${look.shadow.opacity})`,
        }
      : {}),
  };
}

/** Which logical region of the table a cell sits in, per style part. */
interface LookRegion {
  atTop: boolean;
  atBottom: boolean;
  atLeft: boolean;
  atRight: boolean;
}

/**
 * Which parts apply to one cell, in Office's own painting order: the whole
 * table, then row banding, then column banding, then the first / last column,
 * then the first / last row. A later part overrides the earlier one, which is
 * why a header row runs unbroken across an emphasised first column.
 */
function lookPartsForCell(
  look: TableStyleLook,
  flags: TableRowEmphasisFlags,
  rowIndex: number,
  rowCount: number,
  colIndex: number,
  colCount: number,
): Array<{ part: TableStyleLookPart; region: LookRegion }> {
  const isHeader = flags.headerRow === true && rowIndex === 0;
  const isTotal =
    flags.totalRow === true && rowCount > 0 && rowIndex === rowCount - 1;
  const isFirstCol = flags.firstColumn === true && colIndex === 0;
  const isLastCol =
    flags.lastColumn === true && colCount > 0 && colIndex === colCount - 1;
  const wholeRegion: LookRegion = {
    atTop: rowIndex === 0,
    atBottom: rowIndex === rowCount - 1,
    atLeft: colIndex === 0,
    atRight: colIndex === colCount - 1,
  };
  const applied: Array<{ part: TableStyleLookPart; region: LookRegion }> = [];
  const push = (name: TableStyleLookPartName, region: LookRegion): void => {
    const part = look.parts[name];
    if (part) applied.push({ part, region });
  };
  push("wholeTbl", wholeRegion);
  if (flags.bandedRows === true && !isHeader && !isTotal) {
    // One band is one ROW here: the model has no multi-row band size. The
    // FIRST body row is band 1 (ECMA-376 `band1H`, "the first row in
    // horizontal banding"): PowerPoint shades the row under the header with
    // Medium Style 2's darker tint. The deck PDF (`presentation_pdf._look_cell`)
    // and Word's table styles (`document-table-design-authority`) count the
    // same way.
    const bodyIndex = rowIndex - (flags.headerRow === true ? 1 : 0);
    if (bodyIndex >= 0) {
      push(bodyIndex % 2 === 0 ? "band1H" : "band2H", {
        ...wholeRegion,
        atTop: true,
        atBottom: true,
      });
    }
  }
  if (flags.bandedColumns === true && !isFirstCol && !isLastCol) {
    // Column bands start after an emphasised first column, as row bands
    // start after the header row.
    const bodyCol = colIndex - (flags.firstColumn === true ? 1 : 0);
    push(bodyCol % 2 === 0 ? "band1V" : "band2V", {
      ...wholeRegion,
      atLeft: true,
      atRight: true,
    });
  }
  if (isFirstCol) {
    push("firstCol", { ...wholeRegion, atLeft: true, atRight: true });
  }
  if (isLastCol) {
    push("lastCol", { ...wholeRegion, atLeft: true, atRight: true });
  }
  if (isHeader) {
    push("firstRow", { ...wholeRegion, atTop: true, atBottom: true });
  }
  if (isTotal) {
    push("lastRow", { ...wholeRegion, atTop: true, atBottom: true });
  }
  const corner = { atTop: true, atBottom: true, atLeft: true, atRight: true };
  if (isHeader && isFirstCol) push("nwCell", corner);
  if (isHeader && isLastCol) push("neCell", corner);
  if (isTotal && isFirstCol) push("swCell", corner);
  if (isTotal && isLastCol) push("seCell", corner);
  return applied;
}

/**
 * The edge a part contributes to one side of a cell: the region's own outer
 * line when the cell sits on that side of the region, else the region's
 * interior line. An edge the part does not declare contributes nothing, so a
 * style with only interior rules never paints an outer frame.
 */
function lookEdge(
  part: TableStyleLookPart,
  region: LookRegion,
  side: "top" | "bottom" | "left" | "right",
): TableCellBorder | undefined {
  const borders = part.borders;
  if (!borders) return undefined;
  if (side === "top") return region.atTop ? borders.top : borders.insideH;
  if (side === "bottom") {
    return region.atBottom ? borders.bottom : borders.insideH;
  }
  if (side === "left") return region.atLeft ? borders.left : borders.insideV;
  return region.atRight ? borders.right : borders.insideV;
}

/**
 * CSS for one `<td>` from a RESOLVED native table style. Returns `null` when
 * the look contributes nothing to this cell, so the caller can fall through to
 * the preset path untouched.
 */
export function tableStyleLookCellCss(
  look: TableStyleLook | null,
  flags: TableRowEmphasisFlags,
  rowIndex: number,
  rowCount: number,
  colIndex: number,
  colCount: number,
): CSSProperties | null {
  if (!look) return null;
  const applied = lookPartsForCell(
    look,
    flags,
    rowIndex,
    rowCount,
    colIndex,
    colCount,
  );
  if (applied.length === 0) return null;
  let fill: string | null = look.background ?? null;
  let fillOpacity = 1;
  let color: string | null = null;
  let bold = false;
  let italic = false;
  let fontFamily: string | null = null;
  const edges: Partial<
    Record<"top" | "bottom" | "left" | "right", TableCellBorder>
  > = {};
  for (const { part, region } of applied) {
    // An explicit `noFill` makes the cell transparent, which in OOXML reveals
    // the TABLE background beneath it — not the slide. Falling back to the
    // table's own fill keeps a themed style's backdrop visible instead of
    // punching a hole through it.
    if (part.fill === "none") {
      fill = look.backgroundGradient
        ? "transparent"
        : (look.background ?? "transparent");
      fillOpacity = 1;
    } else if (part.fill !== undefined) {
      fill = part.fill;
      fillOpacity = part.fillOpacity ?? 1;
    }
    if (part.color !== undefined) color = part.color;
    if (part.bold) bold = true;
    if (part.italic) italic = true;
    if (part.fontFamily !== undefined) fontFamily = part.fontFamily;
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const border = lookEdge(part, region, side);
      if (border) edges[side] = border;
    }
  }
  const css: CSSProperties = {};
  if (fill !== null) {
    css.background =
      fillOpacity < 1 && HEX_COLOR_RE.test(fill)
        ? `rgba(${[1, 3, 5].map((i) => parseInt(fill!.slice(i, i + 2), 16)).join(", ")}, ${fillOpacity})`
        : fill;
  }
  if (color !== null) css.color = color;
  if (bold) css.fontWeight = 700;
  if (italic) css.fontStyle = "italic";
  // A cell's or a run's own face is set over it (`cellStyleCss`, the runs).
  const face = fontFamily === null ? undefined : tableCellFontFamilyCss(fontFamily);
  if (face) css.fontFamily = face;
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const edge = edges[side];
    if (edge) Object.assign(css, tableEdgeCss(side, edge));
  }
  return Object.keys(css).length > 0 ? css : null;
}

/**
 * The four colours a gallery swatch needs to PREVIEW a resolved style: header
 * band, banded body row, plain body, and the gridline. Derived from the same
 * resolution the canvas paints, so the swatch cannot drift from the table —
 * which is exactly what a hand-tuned preview gradient does.
 */
export function tableStyleLookSwatch(look: TableStyleLook): {
  headerFill: string;
  bandFill: string;
  bodyFill: string;
  borderColor: string;
} {
  const bodyFill =
    look.parts.wholeTbl?.fill && look.parts.wholeTbl.fill !== "none"
      ? look.parts.wholeTbl.fill
      : (look.background ?? "#ffffff");
  const headerFill =
    look.parts.firstRow?.fill && look.parts.firstRow.fill !== "none"
      ? look.parts.firstRow.fill
      : bodyFill;
  const bandFill =
    look.parts.band1H?.fill && look.parts.band1H.fill !== "none"
      ? look.parts.band1H.fill
      : bodyFill;
  const borders = look.parts.wholeTbl?.borders;
  const gridline =
    borders?.insideH ?? borders?.insideV ?? borders?.top ?? borders?.left;
  return {
    headerFill,
    bandFill,
    bodyFill,
    borderColor:
      gridline && gridline.style !== "none"
        ? gridline.color
        : DEFAULT_TABLE_CELL_BORDER_COLOR,
  };
}

/**
 * One resolved border as a CSS shorthand (shared with `cellStyleCss`). A
 * browser draws a `double` border thinner than 3px as one solid line, so a
 * double line is given the 3px it needs to show as two (PowerPoint draws its
 * compound double line at any weight).
 */
export function tableStyleBorderCss(border: TableCellBorder): string {
  // Where a compound edge's rules cannot be drawn (`tableEdgeCss` draws them
  // on a `<td>`), it draws as the even double line.
  const style = border.compound && border.style !== "none" ? "double" : border.style;
  return style === "none"
    ? "none"
    : `${style === "double" ? Math.max(border.width, 3) : border.width}px ${
        {
          solid: "solid",
          dash: "dashed",
          dot: "dotted",
          double: "double",
        }[style]
      } ${border.color}`;
}

/**
 * Resolve a lightweight web preview of a known native style. The PPTX retains
 * only the GUID and flags; these CSS values are never persisted into cells.
 */
export function nativeTableStyleCellCss(
  tableStyleId: unknown,
  flags: TableRowEmphasisFlags,
  rowIndex: number,
  rowCount: number,
  colIndex: number,
  colCount: number,
): CSSProperties | null {
  const styleId = normalizeTableStyleId(tableStyleId);
  const preset = TABLE_STYLE_PRESETS.find(
    (candidate) => candidate.tableStyleId === styleId,
  );
  if (!preset) return null;

  const isHeader = flags.headerRow === true && rowIndex === 0;
  const isTotal =
    flags.totalRow === true && rowCount > 0 && rowIndex === rowCount - 1;
  const { bandFill } = tableRowEmphasis(flags, rowIndex, rowCount);
  const { colEmphasis } = tableColEmphasis(flags, colIndex, colCount);
  const css: CSSProperties = {};
  const fill = isHeader
    ? preset.headerFill
    : colEmphasis || bandFill
      ? (preset.bandFill ?? COLUMN_EMPHASIS_FILL)
      : preset.bodyFill;
  if (fill) css.background = fill;
  if (isHeader && preset.headerColor) css.color = preset.headerColor;
  if ((isHeader && preset.headerBold) || isTotal) css.fontWeight = 700;
  if (preset.borders) {
    const border = `1px solid ${preset.borderColor ?? DEFAULT_TABLE_CELL_BORDER_COLOR}`;
    css.borderTop = border;
    css.borderRight = border;
    css.borderBottom = border;
    css.borderLeft = border;
  }
  if (isTotal) css.borderTop = TOTAL_ROW_BORDER_TOP;
  return css;
}

/**
 * Normalize an arbitrary colour value into a six-digit hex string, or return
 * `fallback` when the value is missing / malformed. Always lower-cases so a
 * persisted deck stays byte-stable. Identical contract to the shape / image
 * colour validators so a stale or hand-edited deck (or an injected
 * `url(...)` / `red;}body{...` payload) can never reach a style attribute raw.
 */
export function normalizeCellColor<F extends string | null | undefined>(
  value: unknown,
  fallback: F,
): string | F {
  if (typeof value === "string" && HEX_COLOR_RE.test(value.trim())) {
    return value.trim().toLowerCase();
  }
  return fallback;
}

/** A cell gradient while `fill` is its first stop, else `null`. */
function pairedCellGradient(value: unknown, fill: string | null): OfficeShapeGradient | null {
  if (value == null || typeof fill !== "string") return null;
  try {
    const gradient = normalizeOfficeShapeGradient(value);
    return gradient && gradient.stops[0]?.color === fill ? gradient : null;
  } catch {
    return null;
  }
}

/** A cell picture while the cell states no colour, else `null`. */
function pairedCellPicture(value: unknown, fill: string | null): BgPicture | null {
  if (value == null || fill !== null) return null;
  return normalizeBgPicture(value, isDrawableImageSource) ?? null;
}

/** A cell pattern while `fill` is its foreground, else `null`. */
function pairedCellPattern(value: unknown, fill: string | null): OfficeChartFillPattern | null {
  if (value == null || typeof fill !== "string") return null;
  try {
    const pattern = normalizeOfficeChartFillPattern(value);
    return pattern && pattern.foreground === fill ? pattern : null;
  } catch {
    return null;
  }
}

/** Normalize a persisted alignment into the whitelist, defaulting to left. */
export function normalizeCellAlignment(value: unknown): CellAlignment {
  if (typeof value !== "string") return DEFAULT_CELL_ALIGNMENT;
  const lowered = value.trim().toLowerCase();
  return (
    NATIVE_CELL_ALIGNMENTS.find((align) => align.toLowerCase() === lowered) ??
    DEFAULT_CELL_ALIGNMENT
  );
}

/**
 * The raw, optional per-cell style fields a cell may persist. All optional /
 * additive — a cell without any of them resolves to "no style".
 */
export interface CellStyleInput {
  fill?: unknown;
  align?: unknown;
  bold?: unknown;
  italic?: unknown;
  underline?: unknown;
  strike?: unknown;
  superscript?: unknown;
  subscript?: unknown;
  highlight?: unknown;
  color?: unknown;
  fontFamily?: unknown;
  fontSize?: unknown;
  verticalAlign?: unknown;
  textDirection?: unknown;
  marginLeft?: unknown;
  marginRight?: unknown;
  marginTop?: unknown;
  marginBottom?: unknown;
  borderLeft?: unknown;
  borderRight?: unknown;
  borderTop?: unknown;
  borderBottom?: unknown;
  borderDiagonalDown?: unknown;
  borderDiagonalUp?: unknown;
  bevel?: unknown;
  fillGradient?: unknown;
  fillPattern?: unknown;
  fillImage?: unknown;
  underlineStyle?: unknown;
  strikeStyle?: unknown;
  textPaint?: unknown;
}

/** A resolved, render-ready per-cell style descriptor. */
export interface ResolvedCellStyle {
  /** Background fill hex, or null when the cell uses the table default. */
  fill: string | null;
  /** Horizontal text alignment. */
  align: CellAlignment;
  /** Whether the cell text is bold. */
  bold: boolean;
  /** Whether the complete cell text is italic. */
  italic: boolean;
  /** Whether the complete cell text has a single underline. */
  underline: boolean;
  /** Whether the complete cell text has a single strikethrough. */
  strike: boolean;
  /** Mutually exclusive whole-cell script position. */
  superscript: boolean;
  subscript: boolean;
  /** Whole-cell DrawingML text highlight, distinct from the cell fill. */
  highlight: string | null;
  /** Text colour hex, or null when the cell uses the table default. */
  color: string | null;
  /** Explicit DrawingML typeface, or null when the cell inherits the theme. */
  fontFamily: TableCellFontFamily | null;
  /** Explicit point size, or null when the cell inherits the theme. */
  fontSize: number | null;
  /** Vertical text anchor within the cell box (top is the default). */
  verticalAlign: VerticalAnchor;
  textDirection: TableCellTextDirection;
  marginLeft: number | null;
  marginRight: number | null;
  marginTop: number | null;
  marginBottom: number | null;
  borderLeft: TableCellBorder | null;
  borderRight: TableCellBorder | null;
  borderTop: TableCellBorder | null;
  borderBottom: TableCellBorder | null;
  borderDiagonalDown: TableCellBorder | null;
  borderDiagonalUp: TableCellBorder | null;
  bevel: TableCellBevel | null;
  /**
   * Shading > Gradient (`a:gradFill`), drawn while `fill` is still its first
   * stop: picking another Shading colour replaces it, as PowerPoint does.
   */
  fillGradient: OfficeShapeGradient | null;
  /** A pattern fill (`a:pattFill`), drawn while `fill` is its foreground. */
  fillPattern: OfficeChartFillPattern | null;
  /**
   * Shading > Picture (`a:blipFill`), the slide background's picture model
   * (`{url, fit, fillRect?, transparency?}`), drawn while the cell states no
   * colour: picking a Shading colour or No Fill replaces it, as PowerPoint
   * does. The package keeps the cell's own picture.
   */
  fillImage: BgPicture | null;
  /** The underline's kind beyond the single line, only while `underline`. */
  underlineStyle: PresentationUnderlineStyle | null;
  /** The double strikethrough, only while `strike`. */
  strikeStyle: "double" | null;
  /**
   * Table Design > WordArt Styles (PPT-A138): the shared WordArt paint the
   * cell's words wear (fill, outline, effects); it owns the glyph colour, so
   * `color` stays what the cell states underneath.
   */
  textPaint: OfficeTextPaint | null;
  /** True iff at least one style is non-default (so the caller can branch). */
  hasStyle: boolean;
}

/** PowerPoint/WPS Layout → Cell Margins presets, expressed in CSS pixels. */
export function tableCellMarginPresetPatch(
  preset: TableCellMarginPreset,
): Pick<
  CellStyleInput,
  "marginLeft" | "marginRight" | "marginTop" | "marginBottom"
> {
  switch (preset) {
    case "none":
      return {
        marginLeft: 0,
        marginRight: 0,
        marginTop: 0,
        marginBottom: 0,
      };
    case "narrow":
      return {
        marginLeft: 4.8,
        marginRight: 4.8,
        marginTop: 4.8,
        marginBottom: 4.8,
      };
    case "wide":
      return {
        marginLeft: 14.4,
        marginRight: 14.4,
        marginTop: 14.4,
        marginBottom: 14.4,
      };
    case "normal":
      return {
        marginLeft: 9.6,
        marginRight: 9.6,
        marginTop: 4.8,
        marginBottom: 4.8,
      };
  }
}

export function tableCellMarginPresetForStyle(
  style:
    | Pick<
        ResolvedCellStyle,
        "marginLeft" | "marginRight" | "marginTop" | "marginBottom"
      >
    | null
    | undefined,
): TableCellMarginPreset | "custom" {
  if (!style) return "normal";
  const margins = [
    style.marginLeft,
    style.marginRight,
    style.marginTop,
    style.marginBottom,
  ] as const;
  if (margins.every((margin) => margin === null)) return "normal";
  for (const preset of TABLE_CELL_MARGIN_PRESETS) {
    const patch = tableCellMarginPresetPatch(preset);
    if (
      style.marginLeft === patch.marginLeft &&
      style.marginRight === patch.marginRight &&
      style.marginTop === patch.marginTop &&
      style.marginBottom === patch.marginBottom
    ) {
      return preset;
    }
  }
  return "custom";
}

/**
 * Resolve a cell's persisted style fields into a normalized descriptor. A
 * malformed / injected colour collapses to `null` (the table default), an
 * unknown alignment collapses to `left`, and `bold` is a strict boolean. The
 * descriptor is `hasStyle: false` exactly when no fill, default alignment, no
 * bold, and no colour apply — letting the renderer keep the plain-cell path.
 */
export function normalizeCellStyle(
  input: CellStyleInput | undefined,
): ResolvedCellStyle {
  // "none" is an explicit No Fill (DrawingML `a:noFill`): the cell is clear
  // even where the table style would shade it (PPT-B068). An absent fill
  // takes the style's.
  // A fill may carry Transparency as `#rrggbbaa` (an imported `a:alpha`);
  // the Shading palette writes six digits, which replaces it.
  const fill =
    input?.fill === "none"
      ? "none"
      : typeof input?.fill === "string" && /^#[0-9a-fA-F]{8}$/.test(input.fill.trim())
        ? input.fill.trim().toLowerCase()
        : normalizeCellColor(input?.fill, null);
  const fillGradient = pairedCellGradient(input?.fillGradient, fill);
  const fillPattern = pairedCellPattern(input?.fillPattern, fill);
  const fillImage = pairedCellPicture(input?.fillImage, fill);
  const align = normalizeCellAlignment(input?.align);
  const bold = input?.bold === true;
  const italic = input?.italic === true;
  const underline = input?.underline === true;
  const underlineStyle = underline
    ? (normalizePresentationUnderlineStyle(input?.underlineStyle) ?? null)
    : null;
  const strike = input?.strike === true;
  const strikeStyle = strike && input?.strikeStyle === "double" ? "double" : null;
  const superscript = input?.superscript === true;
  const subscript = !superscript && input?.subscript === true;
  const highlight = normalizeCellColor(input?.highlight, null);
  const color = normalizeCellColor(input?.color, null);
  const fontFamily = normalizeTableCellFontFamily(input?.fontFamily);
  const fontSize = normalizeTableCellFontSize(input?.fontSize);
  const verticalAlign = normalizeVerticalAnchor(input?.verticalAlign);
  const textDirection =
    typeof input?.textDirection === "string" &&
    TABLE_CELL_TEXT_DIRECTION_SET.has(input.textDirection)
      ? (input.textDirection as TableCellTextDirection)
      : DEFAULT_TABLE_CELL_TEXT_DIRECTION;
  const normalizeMargin = (value: unknown): number | null =>
    typeof value === "number" && Number.isFinite(value) && value >= 0
      ? Math.min(MAX_TABLE_CELL_MARGIN, Math.round(value * 10) / 10)
      : null;
  const normalizeBorder = (value: unknown): TableCellBorder | null => {
    if (!value || typeof value !== "object" || Array.isArray(value))
      return null;
    const record = value as Record<string, unknown>;
    if (
      typeof record.style !== "string" ||
      !TABLE_CELL_BORDER_STYLE_SET.has(record.style)
    ) {
      return null;
    }
    const style = record.style as TableCellBorderStyle;
    // An imported edge with Transparency keeps its alpha, `#rrggbbaa`
    
    const color =
      typeof record.color === "string" &&
      /^#[0-9a-fA-F]{8}$/.test(record.color.trim())
        ? record.color.trim().toLowerCase()
        : normalizeCellColor(record.color, DEFAULT_TABLE_CELL_BORDER_COLOR);
    // A positive width is kept as authored, below the menus' quarter point
    // too: a file's hairline (LibreOffice's `a:lnL w="720"`, 0.06 pt) raised
    // to 0.25 pt on open was a change to every cell the user never made, and
    
    // and commands still offer and accept nothing thinner. Nor wider: a
    // file's 12 pt rule (16 px, census30-probes compound-border-probe2) is
    // kept up to DrawingML's own largest line (`ST_LineWidth`, 1584 pt),
    // where it was cut to the menus' 9 pt and drawn narrower than the PDF
    // export and PowerPoint print it.
    const width =
      typeof record.width === "number" && Number.isFinite(record.width)
        ? Math.min(
            MAX_FILE_TABLE_CELL_BORDER_WIDTH,
            record.width > 0 ? record.width : MIN_TABLE_CELL_BORDER_WIDTH,
          )
        : DEFAULT_TABLE_CELL_BORDER_WIDTH;
    const compound =
      style !== "none" &&
      style !== "double" &&
      typeof record.compound === "string" &&
      (TABLE_CELL_BORDER_COMPOUNDS as readonly string[]).includes(record.compound)
        ? (record.compound as TableCellBorderCompound)
        : undefined;
    return compound ? { style, color, width, compound } : { style, color, width };
  };
  const marginLeft = normalizeMargin(input?.marginLeft);
  const marginRight = normalizeMargin(input?.marginRight);
  const marginTop = normalizeMargin(input?.marginTop);
  const marginBottom = normalizeMargin(input?.marginBottom);
  const borderLeft = normalizeBorder(input?.borderLeft);
  const borderRight = normalizeBorder(input?.borderRight);
  const borderTop = normalizeBorder(input?.borderTop);
  const borderBottom = normalizeBorder(input?.borderBottom);
  const borderDiagonalDown = normalizeBorder(input?.borderDiagonalDown);
  const borderDiagonalUp = normalizeBorder(input?.borderDiagonalUp);
  const bevel = normalizeTableCellBevel(input?.bevel);
  let textPaint: OfficeTextPaint | null = null;
  try {
    textPaint = normalizeOfficeTextPaint(input?.textPaint) ?? null;
  } catch {
    // A malformed paint draws nothing rather than breaking the table.
    textPaint = null;
  }
  return {
    fill,
    align,
    bold,
    italic,
    underline,
    strike,
    superscript,
    subscript,
    highlight,
    color,
    fontFamily,
    fontSize,
    verticalAlign,
    textDirection,
    marginLeft,
    marginRight,
    marginTop,
    marginBottom,
    borderLeft,
    borderRight,
    borderTop,
    borderBottom,
    borderDiagonalDown,
    borderDiagonalUp,
    bevel,
    fillGradient,
    fillPattern,
    fillImage,
    underlineStyle,
    strikeStyle,
    textPaint,
    hasStyle:
      textPaint !== null ||
      fill !== null ||
      fillImage !== null ||
      align !== DEFAULT_CELL_ALIGNMENT ||
      bold ||
      italic ||
      underline ||
      strike ||
      superscript ||
      subscript ||
      highlight !== null ||
      color !== null ||
      fontFamily !== null ||
      fontSize !== null ||
      verticalAlign !== DEFAULT_VERTICAL_ANCHOR ||
      textDirection !== DEFAULT_TABLE_CELL_TEXT_DIRECTION ||
      marginLeft !== null ||
      marginRight !== null ||
      marginTop !== null ||
      marginBottom !== null ||
      borderLeft !== null ||
      borderRight !== null ||
      borderTop !== null ||
      borderBottom !== null ||
      borderDiagonalDown !== null ||
      borderDiagonalUp !== null ||
      bevel !== null,
  };
}

/**
 * The CSS to layer onto a `<td>` for the resolved cell style, optionally folding
 * in a table-level header-row emphasis (light fill + bold) for the first row.
 * Returns an empty object when nothing applies (so spreading it onto a plain
 * cell is a no-op and the historical render stays byte-stable). Alignment is
 * only emitted when non-default so an un-aligned cell adds no `text-align` /
 * `vertical-align`.
 */
/**
 * A cell's diagonal border as a background layer, in the line style the pen
 * chose, as the cell's other borders draw it (PPT-B164: a dashed diagonal
 * drew solid). Down runs from the top-left corner to the bottom-right, Up
 * from the bottom-left to the top-right. A solid or double line is a band
 * across the cell's diagonal; a dashed or dotted one is a line whose dashes
 * are measured in the line's own width, drawn at that width whatever the
 * cell's shape.
 */
export function tableCellDiagonalImage(
  border: TableCellBorder | null,
  side: "down" | "up",
): string | null {
  if (!border || border.style === "none") return null;
  // The band's direction runs across the line.
  const across = side === "down" ? "to top right" : "to bottom right";
  const band = (stops: readonly [number, number][]) =>
    `linear-gradient(${across}, ${[
      "transparent 0",
      ...stops.flatMap(([from, to]) => [
        `transparent calc(50% + ${from}px)`,
        `${border.color} calc(50% + ${from}px)`,
        `${border.color} calc(50% + ${to}px)`,
        `transparent calc(50% + ${to}px)`,
      ]),
    ].join(", ")})`;
  if (border.style === "solid" && !border.compound) {
    const half = Math.max(0.5, border.width / 2);
    return band([[-half, half]]);
  }
  if (border.compound) {
    // A diagonal of several rules: `TABLE_EDGE_COMPOUND_RULES` across it.
    const width = Math.max(border.width, 3);
    return band(
      TABLE_EDGE_COMPOUND_RULES[border.compound].map(
        ([from, to]) => [(from - 0.5) * width, (to - 0.5) * width] as [number, number],
      ),
    );
  }
  if (border.style === "double") {
    // Two rules and the gap between them, each a third of the line, as the
    // cell's straight double borders draw it.
    const width = Math.max(border.width, 3);
    return band([
      [-width / 2, -width / 6],
      [width / 6, width / 2],
    ]);
  }
  const width = Math.max(1, border.width);
  const dashes = border.style === "dash" ? `${4 * width} ${3 * width}` : `${width} ${width}`;
  const [y1, y2] = side === "down" ? [0, 100] : [100, 0];
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 100 100" preserveAspectRatio="none">` +
    `<line x1="0" y1="${y1}" x2="100" y2="${y2}" stroke="${border.color}" stroke-width="${width}" ` +
    `stroke-dasharray="${dashes}" vector-effect="non-scaling-stroke"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`;
}

export function cellStyleCss(
  style: ResolvedCellStyle,
  options?: {
    headerEmphasis?: boolean;
    /** Banded-row fallback fill (PowerPoint Banded Rows), under any cell fill. */
    bandFill?: string | null;
    /** Total-row emphasis (PowerPoint Total Row): bold like the header. */
    totalEmphasis?: boolean;
    /**
     * Column emphasis (PowerPoint First/Last/Banded Columns): a light fill on
     * the emphasised column, layered UNDER any explicit cell fill but OVER the
     * banded-row fill (the column emphasis wins the cross-banding tie, matching
     * PowerPoint, where the column styling takes the foreground).
     */
    colEmphasis?: boolean;
    /**
     * A right-to-left table (`a:tblPr@rtl`): PowerPoint draws a cell's
     * diagonals mirrored (`a:lnTlToBr` from the bottom-left corner to the
     * top-right), as it draws `a:lnL` on the cell's right. Measured in
     * PowerPoint for Mac.
     */
    rightToLeft?: boolean;
  },
): CSSProperties {
  const headerEmphasis = options?.headerEmphasis === true;
  const totalEmphasis = options?.totalEmphasis === true;
  const bandFill = options?.bandFill ?? null;
  const colEmphasis = options?.colEmphasis === true;
  const css: CSSProperties = {};
  // Fill precedence: an explicit per-cell colour wins over the header default,
  // which wins over the column-emphasis fill, which wins over the banded-row
  // fill — so a hand-coloured cell is honoured inside an emphasised header,
  // column, or banded body row.
  const fill =
    style.fill === "none"
      ? "transparent"
      : (style.fill ??
        (headerEmphasis
          ? HEADER_ROW_FILL
          : colEmphasis
            ? COLUMN_EMPHASIS_FILL
            : bandFill));
  if (fill !== null) css.background = fill;
  if (style.fillPattern) css.background = officeFillPatternCss(style.fillPattern);
  const gradientImage = style.fillGradient ? officeGradientCss(style.fillGradient) : null;
  if (style.align !== DEFAULT_CELL_ALIGNMENT) {
    Object.assign(css, tableCellAlignmentCss(style.align));
  }
  if (style.bold || headerEmphasis || totalEmphasis) css.fontWeight = 700;
  if (style.italic) css.fontStyle = "italic";
  const textDecoration = tableCellTextDecoration(style);
  if (textDecoration) {
    css.textDecoration = textDecoration;
    Object.assign(css, tableCellTextDecorationKindCss(style));
  }
  if (style.color !== null) css.color = style.color;
  const fontFamily = tableCellFontFamilyCss(style.fontFamily);
  if (fontFamily) css.fontFamily = fontFamily;
  if (style.fontSize !== null) css.fontSize = `${style.fontSize}pt`;
  // VERTICAL anchor (top/middle/bottom). The anchor enum values are themselves
  // valid CSS `vertical-align` keywords, so they map 1:1. Emitted only when
  // non-default `top` so an un-anchored cell adds no `vertical-align`.
  if (style.verticalAlign !== DEFAULT_VERTICAL_ANCHOR) {
    css.verticalAlign = style.verticalAlign;
  }
  if (style.textDirection === "rotate-90") {
    // "Rotate all text 90°": every glyph turns, CJK included, as the same
    
    // upright (its `mixed` orientation).
    css.writingMode = "vertical-rl";
    css.textOrientation = "sideways";
  } else if (style.textDirection === "rotate-270") {
    // Bottom to top with the glyphs turned anticlockwise (PowerPoint's
    // "Rotate all text 270°"), not vertical-lr's clockwise text.
    css.writingMode = "sideways-lr";
  } else if (style.textDirection === "eastAsianVertical") {
    css.writingMode = "vertical-rl";
  } else if (style.textDirection === "mongolianVertical") {
    css.writingMode = "vertical-lr";
  } else if (style.textDirection === "stackedRightToLeft") {
    css.writingMode = "vertical-rl";
    css.textOrientation = "upright";
  } else if (style.textDirection === "stacked") {
    css.writingMode = "vertical-rl";
    css.textOrientation = "upright";
  }
  if (style.marginLeft !== null) css.paddingLeft = style.marginLeft;
  if (style.marginRight !== null) css.paddingRight = style.marginRight;
  if (style.marginTop !== null) css.paddingTop = style.marginTop;
  if (style.marginBottom !== null) css.paddingBottom = style.marginBottom;
  if (style.borderLeft) Object.assign(css, tableEdgeCss("left", style.borderLeft));
  if (style.borderRight) Object.assign(css, tableEdgeCss("right", style.borderRight));
  if (style.borderTop) Object.assign(css, tableEdgeCss("top", style.borderTop));
  if (style.borderBottom) Object.assign(css, tableEdgeCss("bottom", style.borderBottom));
  const mirrored = options?.rightToLeft === true;
  const diagonals = [
    tableCellDiagonalImage(style.borderDiagonalDown, mirrored ? "up" : "down"),
    tableCellDiagonalImage(style.borderDiagonalUp, mirrored ? "down" : "up"),
  ].filter((value): value is string => value !== null);
  const layers = gradientImage ? [...diagonals, gradientImage] : diagonals;
  if (style.fillImage) {
    // Shading > Picture: the cell's own fill, over no style colour, drawn
    // beneath its diagonal borders as the slide background's picture is.
    const picture = bgPictureLayer(style.fillImage);
    const pictureLayers = [...diagonals, ...picture.images];
    css.background = picture.color ?? "transparent";
    css.backgroundImage = pictureLayers.join(", ");
    css.backgroundPosition = [...diagonals.map(() => "0 0"), ...picture.positions].join(", ");
    css.backgroundSize = [...diagonals.map(() => "100% 100%"), ...picture.sizes].join(", ");
    css.backgroundRepeat = [...diagonals.map(() => "no-repeat"), ...picture.repeats].join(", ");
  } else if (layers.length > 0) css.backgroundImage = layers.join(", ");
  // The 3-D cell bevel reads as inner shading. Appended to any existing
  // box-shadow (the focused-cell ring sets its own inline) so neither wins.
  const bevelShadow = tableCellBevelCss(style.bevel);
  if (bevelShadow) css.boxShadow = bevelShadow;
  return css;
}

/** Compose underline and strikethrough without either format masking the other. */
export function tableCellTextDecoration(
  style: Pick<ResolvedCellStyle, "underline" | "strike">,
): string | undefined {
  const decorations: string[] = [];
  if (style.underline) decorations.push("underline");
  if (style.strike) decorations.push("line-through");
  return decorations.length > 0 ? decorations.join(" ") : undefined;
}

/**
 * The decoration's kind (double, heavy, dotted, dashed, wavy) as CSS, for an
 * imported cell's `underlineStyle`; `{}` for the single line.
 */
export function tableCellTextDecorationKindCss(
  style: Pick<
    ResolvedCellStyle,
    "underline" | "underlineStyle" | "strike" | "strikeStyle"
  >,
): CSSProperties {
  return presentationTextDecorationCss(
    style.underline ? (style.underlineStyle ?? undefined) : undefined,
    style.strike ? (style.strikeStyle ?? undefined) : undefined,
  );
}

/** Inline text styling for PowerPoint/WPS whole-cell super/subscript. */
export function tableCellBaselineCss(
  style: Pick<ResolvedCellStyle, "superscript" | "subscript" | "fontSize">,
): CSSProperties {
  if (!style.superscript && !style.subscript) return {};
  return {
    verticalAlign: style.superscript ? "super" : "sub",
    ...(style.fontSize === null ? { fontSize: "0.7em" } : {}),
  };
}

/**
 * Inline glyph styling for whole-cell script and DrawingML text highlight.
 * The highlight stays on the text span rather than the `<td>`, so it never
 * silently turns into a cell fill.
 */
export function tableCellTextCss(
  style: Pick<
    ResolvedCellStyle,
    "superscript" | "subscript" | "fontSize" | "highlight"
  >,
): CSSProperties {
  const css = tableCellBaselineCss(style);
  if (style.highlight !== null) {
    css.backgroundColor = style.highlight;
    css.boxDecorationBreak = "clone";
    css.WebkitBoxDecorationBreak = "clone";
    css.padding = "0 1px";
  }
  return css;
}

/** Subtle fill for banded (alternate) body rows (PowerPoint Banded Rows). */
export const BANDED_ROW_FILL = "#f1f5f9";

/** Top border separating an emphasised total row (PowerPoint Total Row). */
export const TOTAL_ROW_BORDER_TOP = "2px solid #475569";

/** The table-level style-option flags that drive per-row emphasis. */
export interface TableRowEmphasisFlags {
  headerRow?: boolean;
  bandedRows?: boolean;
  totalRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  bandedColumns?: boolean;
}

/**
 * Per-row table-style emphasis from the table-level flags (PowerPoint "Table
 * Style Options"). Pure + render-agnostic: returns the banded-row fallback fill
 * (alternate BODY rows, skipping the header + total rows) and whether the row is
 * the emphasised total row. The render layers `bandFill` UNDER each cell's own
 * fill and adds the total-row top border. Banding starts on the SECOND body row
 * so the first data row stays unbanded (matching PowerPoint).
 */
export function tableRowEmphasis(
  flags: TableRowEmphasisFlags,
  rowIndex: number,
  rowCount: number,
): { bandFill: string | null; totalEmphasis: boolean } {
  const isHeader = flags.headerRow === true && rowIndex === 0;
  const isTotal =
    flags.totalRow === true && rowCount > 0 && rowIndex === rowCount - 1;
  let bandFill: string | null = null;
  if (flags.bandedRows === true && !isHeader && !isTotal) {
    const bodyIndex = rowIndex - (flags.headerRow === true ? 1 : 0);
    if (bodyIndex >= 0 && bodyIndex % 2 === 1) bandFill = BANDED_ROW_FILL;
  }
  return { bandFill, totalEmphasis: isTotal };
}

/**
 * Per-column table-style emphasis from the table-level flags (PowerPoint "Table
 * Style Options"): First Column / Last Column / Banded Columns. Pure +
 * render-agnostic, mirroring {@link tableRowEmphasis}: returns whether the
 * column should carry the emphasis fill. The first / last columns take
 * precedence over banding (and are never themselves banded), and banding starts
 * on the SECOND ordinary column so the first data column stays unbanded
 * (matching the row banding, which starts on the second body row).
 */
export function tableColEmphasis(
  flags: TableRowEmphasisFlags,
  colIndex: number,
  colCount: number,
): { colEmphasis: boolean } {
  const isFirst = flags.firstColumn === true && colIndex === 0;
  const isLast =
    flags.lastColumn === true && colCount > 0 && colIndex === colCount - 1;
  const colBand =
    flags.bandedColumns === true &&
    !isFirst &&
    !isLast &&
    colIndex >= 1 &&
    (colIndex - 1) % 2 === 1;
  return { colEmphasis: isFirst || isLast || colBand };
}

/** CSS for a `<td>` when the table-level `borders` flag is set, else `{}`. */
export function tableBorderCss(borders: boolean): CSSProperties {
  return borders
    ? {
        borderTop: TABLE_BORDER_CSS,
        borderRight: TABLE_BORDER_CSS,
        borderBottom: TABLE_BORDER_CSS,
        borderLeft: TABLE_BORDER_CSS,
      }
    : {};
}

/**
 * The style fields to PERSIST for a cell, omitting every field at its default
 * (no fill / left align / not bold / no colour) so a cell without meaningful
 * formatting serializes byte-for-byte as before. Returns an empty object in
 * that case — callers spread it so an un-styled cell keeps `{ text }` only.
 */
export function persistedCellStyle(input: CellStyleInput | undefined): Partial<{
  fill: string;
  align: CellAlignment;
  bold: true;
  italic: true;
  underline: true;
  strike: true;
  superscript: true;
  subscript: true;
  highlight: string;
  color: string;
  fontFamily: TableCellFontFamily;
  fontSize: number;
  verticalAlign: VerticalAnchor;
  textDirection: TableCellTextDirection;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  borderLeft: TableCellBorder;
  borderRight: TableCellBorder;
  borderTop: TableCellBorder;
  borderBottom: TableCellBorder;
  borderDiagonalDown: TableCellBorder;
  borderDiagonalUp: TableCellBorder;
  bevel: TableCellBevel;
  fillGradient: OfficeShapeGradient;
  fillPattern: OfficeChartFillPattern;
  fillImage: BgPicture;
  underlineStyle: PresentationUnderlineStyle;
  strikeStyle: "double";
  textPaint: OfficeTextPaint;
}> {
  const style = normalizeCellStyle(input);
  const out: ReturnType<typeof persistedCellStyle> = {};
  if (style.fill !== null) out.fill = style.fill;
  if (style.align !== DEFAULT_CELL_ALIGNMENT) out.align = style.align;
  if (style.bold) out.bold = true;
  if (style.italic) out.italic = true;
  if (style.underline) out.underline = true;
  if (style.strike) out.strike = true;
  if (style.superscript) out.superscript = true;
  if (style.subscript) out.subscript = true;
  if (style.highlight !== null) out.highlight = style.highlight;
  if (style.color !== null) out.color = style.color;
  if (style.fontFamily !== null) out.fontFamily = style.fontFamily;
  if (style.fontSize !== null) out.fontSize = style.fontSize;
  const verticalAlign = persistedVerticalAnchor(style.verticalAlign);
  if (verticalAlign) out.verticalAlign = verticalAlign;
  if (style.textDirection !== DEFAULT_TABLE_CELL_TEXT_DIRECTION) {
    out.textDirection = style.textDirection;
  }
  if (style.marginLeft !== null) out.marginLeft = style.marginLeft;
  if (style.marginRight !== null) out.marginRight = style.marginRight;
  if (style.marginTop !== null) out.marginTop = style.marginTop;
  if (style.marginBottom !== null) out.marginBottom = style.marginBottom;
  if (style.borderLeft) out.borderLeft = style.borderLeft;
  if (style.borderRight) out.borderRight = style.borderRight;
  if (style.borderTop) out.borderTop = style.borderTop;
  if (style.borderBottom) out.borderBottom = style.borderBottom;
  if (style.borderDiagonalDown) {
    out.borderDiagonalDown = style.borderDiagonalDown;
  }
  if (style.borderDiagonalUp) out.borderDiagonalUp = style.borderDiagonalUp;
  if (style.bevel) out.bevel = style.bevel;
  if (style.fillGradient) out.fillGradient = style.fillGradient;
  if (style.fillPattern) out.fillPattern = style.fillPattern;
  if (style.fillImage) out.fillImage = style.fillImage;
  if (style.underlineStyle) out.underlineStyle = style.underlineStyle;
  if (style.strikeStyle) out.strikeStyle = style.strikeStyle;
  if (style.textPaint) out.textPaint = style.textPaint;
  return out;
}

/**
 * Table Design > WordArt Styles on a cell's words (PPT-A138): the shared
 * WordArt paint as the cell text container's CSS. A gradient or pattern
 * fill clips to the glyphs where the words are drawn as text; inside the
 * editing field (`field`), which cannot clip its own background to its
 * glyphs, the words take the fill's first colour instead.
 */
export function tableCellTextPaintCss(
  paint: OfficeTextPaint | null,
  options?: { field?: boolean },
): CSSProperties {
  if (!paint) return {};
  const css = officeTextPaintCss(paint);
  if (options?.field && (paint.fillGradient || paint.fillPattern)) {
    const rest: Record<string, unknown> = { ...css };
    for (const key of [
      "backgroundImage",
      "background",
      "backgroundColor",
      "backgroundSize",
      "backgroundPosition",
      "backgroundClip",
      "WebkitBackgroundClip",
    ]) {
      delete rest[key];
    }
    return { ...(rest as CSSProperties), color: paint.fill === "none" ? "transparent" : paint.fill };
  }
  return css;
}
