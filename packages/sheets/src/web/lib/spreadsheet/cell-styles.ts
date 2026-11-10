/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { SOURCE_XF_STYLE_KEY } from "./style-source-xf";
import {
  BORDER_EDGES,
  BORDER_LINE_STYLES,
  type BorderLineStyle,
} from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { paletteById, themeGrid } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import { spreadsheetNameFold } from "@mrnote/core/web/lib/spreadsheet/name-fold.ts";

/**
 * Excel's built-in Cell Styles gallery.
 *
 * Each entry is a NAMED look the user can stamp onto the selected rectangle
 * with a single click. The entries live here (pure data + a pure helper) so the
 * registry can be unit-tested in isolation and so the gallery UI in
 * SpreadsheetWindow stays thin.
 *
 * ── Style shape ──────────────────────────────────────────────────────────────
 * A patch is a partial of the cell's stored `style` object (see
 * SpreadsheetWindow's `SpreadsheetCell["style"]`). The keys these named styles
 * touch all render in `cellStyleCss`:
 *   - backgroundColor / color   → fill + font color
 *   - bold / italic             → font weight / style
 *   - fontSize                  → on-screen font size (px)
 *   - border / borderTop / borderBottom → box / top / bottom edges
 *   - numberFormat                     → the Number Format gallery group
 * `updateCellStyle` prunes falsy keys (`false`/""/null/undefined), so a patch
 * that sets a key to `false` CLEARS it on the target rather than storing junk.
 *
 * ── MERGE-vs-REPLACE rule (documented choice) ────────────────────────────────
 * Applying a named style MERGES its patch over the cell's existing style, NOT a
 * wholesale replace. This mirrors Excel: a named cell style defines a SUBSET of
 * attributes and leaves the rest of the cell intact. Concretely:
 *
 *   1. KEYS THE NAMED STYLE OWNS are set DETERMINISTICALLY. A style that wants a
 *      look without bold/italic/border sets those keys to `false` in its patch
 *      so the merge clears any prior bold/italic/border — re-stamping or
 *      switching between named styles is idempotent and never leaves a stale
 *      edge or weight behind. (`updateCellStyle` prunes the `false` keys.)
 *   2. KEYS OUTSIDE THE NAMED STYLE'S CONCERN are left UNTOUCHED. Visual styles
 *      leave number format / alignment / wrap alone, while Number Format styles
 *      own only `numberFormat`. The Normal style is the deliberate exception:
 *      it resets the complete built-in visual/number/alignment axes to workbook
 *      defaults while editor-specific control flags remain untouched.
 *
 * So: the named style is authoritative over the visual attributes it lists and
 * a no-op over everything else. This is the "set a defined subset, leave the
 * rest" rule, made deterministic by listing the cleared keys explicitly.
 *
 * ── Determinism / immutability ───────────────────────────────────────────────
 * The registry and every patch object are deeply frozen. `applyCellStyle`
 * returns a FRESH shallow copy of the patch each call, so a caller that mutates
 * the returned object cannot poison the shared registry.
 */

/** The subset of a cell's `style` a named style may carry. */
export interface NamedCellStylePatch {
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  underlineStyle?: "single" | "double" | "singleAccounting" | "doubleAccounting" | false;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontSize?: number | false;
  fontFamily?: string;
  numberFormat?: string;
  border?: boolean;
  borderTop?: boolean;
  borderRight?: boolean;
  borderBottom?: boolean;
  borderLeft?: boolean;
  borderEdges?: false;
  fillPattern?: false;
  gradientFill?: false;
  textAlign?: "left" | "center" | "right" | false;
  verticalAlign?: "top" | "middle" | "bottom" | false;
  wrapText?: boolean;
  textRotation?: number | false;
  verticalText?: boolean;
  indent?: number | false;
  shrinkToFit?: boolean;
  centerAcross?: boolean;
  horizontalFill?: boolean;
  horizontalJustify?: "justify" | "distributed" | false;
  verticalJustify?: "justify" | "distributed" | false;
  justifyLastLine?: boolean;
  readingOrder?: "ltr" | "rtl" | false;
}

export type NamedCellStyleOwnedAxis = "fill" | "border";

/** A single gallery entry: a stable id, a display label, and its style patch. */
export interface NamedCellStyle {
  /** Stable machine id (used for data-testid + lookup); never localized. */
  id: string;
  /** Human label shown in the gallery (Excel's English name). */
  label: string;
  /** Built-in Chinese label; avoids coupling the immutable gallery to catalogs. */
  labelZh?: string;
  /** Short sample rendered inside the visual tile. */
  preview?: string;
  /** Complete renderer axes whose legacy/rich representations this look owns. */
  ownedAxes?: readonly NamedCellStyleOwnedAxis[];
  /** The partial cell style this named style stamps. */
  patch: NamedCellStylePatch;
}

export interface NamedCellStyleGroup {
  id: string;
  label: string;
  labelZh: string;
  styles: readonly NamedCellStyle[];
}

export interface ResolvedCellStyleGallery {
  groups: readonly NamedCellStyleGroup[];
  styles: readonly NamedCellStyle[];
}

const FILL_AND_BORDER_AXES = Object.freeze([
  "fill",
  "border",
] as const satisfies readonly NamedCellStyleOwnedAxis[]);
const BORDER_AXIS = Object.freeze([
  "border",
] as const satisfies readonly NamedCellStyleOwnedAxis[]);

function freezeStyle(style: NamedCellStyle): NamedCellStyle {
  const ownedAxes = Object.freeze([...(style.ownedAxes ?? [])]);
  const patch: NamedCellStylePatch = { ...style.patch };
  // A named style owns the rendered FILL, not merely the flat fallback color.
  // Pattern and gradient layers win in cellStyleCss, so both must be removed
  // whenever the style claims a solid/no-fill result.
  if (ownedAxes.includes("fill")) {
    patch.backgroundColor ??= "";
    patch.fillPattern = false;
    patch.gradientFill = false;
  }
  // Rich borderEdges render over every legacy boolean. A border-owning style
  // therefore sets/clears the complete equivalent axis as one authority.
  if (ownedAxes.includes("border")) {
    patch.border ??= false;
    patch.borderTop ??= false;
    patch.borderRight ??= false;
    patch.borderBottom ??= false;
    patch.borderLeft ??= false;
    patch.borderEdges = false;
  }
  return Object.freeze({
    ...style,
    ownedAxes,
    patch: Object.freeze(patch),
  });
}

// Excel's status-style palette (the fills/fonts Excel uses for Good/Bad/Neutral)
// and the heading greys, kept as literal constants so the look is deterministic
// and there is no computed-color injection surface.
const GOOD_FILL = /* own data: GOOD_FILL */ "#d2f4e0";
const GOOD_TEXT = /* own data: GOOD_TEXT */ "#13723a";
const BAD_FILL = /* own data: BAD_FILL */ "#f4d2d6";
const BAD_TEXT = /* own data: BAD_TEXT */ "#72131f";
const NEUTRAL_FILL = /* own data: NEUTRAL_FILL */ "#f4ecd2";
const NEUTRAL_TEXT = /* own data: NEUTRAL_TEXT */ "#724513";
const HEADING_TEXT = /* own data: HEADING_TEXT */ "#1f2a37"; // Excel's heading blue-grey.
const TITLE_TEXT = /* own data: TITLE_TEXT */ "#1f2a37";

/**
 * The frozen registry. Order here is the gallery's display order: the three
 * status styles first (Excel's "Good, Bad and Neutral" group), then the
 * titles/headings group (Title, Heading 1-4), then Total.
 *
 * Every patch sets the keys it owns — including clearing (`false`) the
 * bold/italic/border keys a given look does NOT want — so switching between
 * named styles is deterministic (see the MERGE-vs-REPLACE note above).
 */
export const NAMED_CELL_STYLES: readonly NamedCellStyle[] = Object.freeze(
  (
    [
      {
        id: "good",
        label: "Good",
        labelZh: "好",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          backgroundColor: GOOD_FILL,
          color: GOOD_TEXT,
          bold: false,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fillPattern: false,
          gradientFill: false,
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "bad",
        label: "Bad",
        labelZh: "差",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          backgroundColor: BAD_FILL,
          color: BAD_TEXT,
          bold: false,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fillPattern: false,
          gradientFill: false,
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "neutral",
        label: "Neutral",
        labelZh: "适中",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          backgroundColor: NEUTRAL_FILL,
          color: NEUTRAL_TEXT,
          bold: false,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fillPattern: false,
          gradientFill: false,
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "title",
        label: "Title",
        labelZh: "标题",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          color: TITLE_TEXT,
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fontSize: 28,
          // Title has no fill; clear any prior fill so it reads as a banner heading.
          backgroundColor: "",
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "heading1",
        label: "Heading 1",
        labelZh: "标题 1",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          color: HEADING_TEXT,
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fontSize: 20,
          backgroundColor: "",
          // Heading 1 carries a bottom rule (Excel underlines headings with a
          // bottom border, not text-underline).
          border: false,
          borderTop: false,
          borderBottom: true,
        }),
      },
      {
        id: "heading2",
        label: "Heading 2",
        labelZh: "标题 2",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          color: HEADING_TEXT,
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fontSize: 16,
          backgroundColor: "",
          border: false,
          borderTop: false,
          borderBottom: true,
        }),
      },
      {
        id: "heading3",
        label: "Heading 3",
        labelZh: "标题 3",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          color: HEADING_TEXT,
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fontSize: 14,
          backgroundColor: "",
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "heading4",
        label: "Heading 4",
        labelZh: "标题 4",
        ownedAxes: FILL_AND_BORDER_AXES,
        patch: Object.freeze({
          color: HEADING_TEXT,
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          fontSize: 13,
          backgroundColor: "",
          border: false,
          borderTop: false,
          borderBottom: false,
        }),
      },
      {
        id: "total",
        label: "Total",
        labelZh: "汇总",
        ownedAxes: BORDER_AXIS,
        patch: Object.freeze({
          bold: true,
          italic: false,
          underline: false,
          underlineStyle: false,
          strikethrough: false,
          // Total rules a line above AND below the row (Excel's double-rule look),
          // and leaves any number format intact (see MERGE-vs-REPLACE note).
          border: false,
          borderTop: true,
          borderBottom: true,
        }),
      },
    ] satisfies NamedCellStyle[]
  ).map(freezeStyle),
);

function coreStyle(id: string): NamedCellStyle {
  const style = NAMED_CELL_STYLES.find((candidate) => candidate.id === id);
  if (!style) throw new Error(`Missing core cell style: ${id}`);
  return freezeStyle(style);
}

const NORMAL_STYLE = freezeStyle({
  id: "normal",
  label: "Normal",
  labelZh: "常规",
  ownedAxes: FILL_AND_BORDER_AXES,
  patch: {
    backgroundColor: "",
    color: "",
    bold: false,
    italic: false,
    underline: false,
    underlineStyle: false,
    strikethrough: false,
    fontSize: false,
    fontFamily: "",
    numberFormat: "",
    border: false,
    borderTop: false,
    borderRight: false,
    borderBottom: false,
    borderLeft: false,
    borderEdges: false,
    fillPattern: false,
    gradientFill: false,
    textAlign: false,
    verticalAlign: false,
    wrapText: false,
    textRotation: false,
    verticalText: false,
    indent: false,
    shrinkToFit: false,
    centerAcross: false,
    // Normal's font sits on the baseline and its alignment is General in
    // the Context reading order.
    superscript: false,
    subscript: false,
    horizontalFill: false,
    horizontalJustify: false,
    verticalJustify: false,
    justifyLastLine: false,
    readingOrder: false,
  },
});

const DATA_MODEL_STYLES: readonly NamedCellStyle[] = /* own data: DATA_MODEL_STYLES */ [{"id":"calculation","label":"Calculation","labelZh":"计算","ownedAxes":["fill","border"],"patch":{"backgroundColor":"","color":"#214e96","bold":true,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":true,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"check-cell","label":"Check Cell","labelZh":"检查单元格","ownedAxes":["fill","border"],"patch":{"backgroundColor":"#d1e9ec","color":"#164d54","bold":true,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":true,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"explanatory-text","label":"Explanatory Text","labelZh":"解释性文本","ownedAxes":["fill","border"],"patch":{"backgroundColor":"","color":"#848a91","bold":false,"italic":true,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"input","label":"Input","labelZh":"输入","ownedAxes":["fill","border"],"patch":{"backgroundColor":"#e2ebf9","color":"#3a444f","bold":false,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":true,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"linked-cell","label":"Linked Cell","labelZh":"链接单元格","ownedAxes":["fill","border"],"patch":{"backgroundColor":"","color":"#2f6fd6","bold":false,"italic":false,"underline":true,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"note","label":"Note","labelZh":"注释","ownedAxes":["fill","border"],"patch":{"backgroundColor":"#f4ecd2","color":"#3a444f","bold":false,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"output","label":"Output","labelZh":"输出","ownedAxes":["fill","border"],"patch":{"backgroundColor":"#eef1f4","color":"#3a444f","bold":true,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":false,"borderBottom":true,"borderRight":false,"borderLeft":false,"borderEdges":false}},{"id":"warning-text","label":"Warning Text","labelZh":"警告文本","ownedAxes":["fill","border"],"patch":{"backgroundColor":"","color":"#954e57","bold":true,"italic":false,"underline":false,"underlineStyle":false,"strikethrough":false,"fillPattern":false,"gradientFill":false,"border":false,"borderTop":false,"borderBottom":false,"borderRight":false,"borderLeft":false,"borderEdges":false}}];

const THEMED_STYLE_STRENGTHS = [20, 40, 60, 100] as const;

function themedStyles(themeId: unknown): readonly NamedCellStyle[] {
  const palette = paletteById(themeId);
  // themeGrid is also the palette popover's color authority. Its last six
  // columns are the live Accent 1..6 slots for both the curated Office palette
  // and generated non-default palettes. Cell Styles uses the 80/60/40% tint
  // rows followed by the base accent.
  const accentColumns = themeGrid(palette).slice(-palette.accents.length);
  return Object.freeze(
    accentColumns.flatMap((column, accentIndex) => {
      const fills = [column[1], column[2], column[3], column[0]];
      return THEMED_STYLE_STRENGTHS.map((strength, strengthIndex) =>
        freezeStyle({
          id: `accent-${accentIndex + 1}-${strength}`,
          label:
            strength === 100
              ? `Accent ${accentIndex + 1}`
              : `${strength}% - Accent ${accentIndex + 1}`,
          labelZh:
            strength === 100
              ? `着色 ${accentIndex + 1}`
              : `${strength}% - 着色 ${accentIndex + 1}`,
          ownedAxes: FILL_AND_BORDER_AXES,
          patch: {
            backgroundColor:
              fills[strengthIndex] ?? palette.accents[accentIndex],
            color: strength === 100 ? "#ffffff" : "#000000",
            bold: false,
            italic: false,
            underline: false,
            underlineStyle: false,
            strikethrough: false,
            border: false,
          },
        }),
      );
    }),
  );
}

const NUMBER_FORMAT_STYLES: readonly NamedCellStyle[] = Object.freeze([
  freezeStyle({
    id: "comma",
    label: "Comma",
    labelZh: "千位分隔",
    preview: "1,234.00",
    // Excel's built-in Comma style (en-US): symbol-less Accounting.
    patch: { numberFormat: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)' },
  }),
  freezeStyle({
    id: "comma-0",
    label: "Comma [0]",
    labelZh: "千位分隔 [0]",
    preview: "1,234",
    patch: { numberFormat: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)' },
  }),
  freezeStyle({
    id: "currency",
    label: "Currency",
    labelZh: "货币",
    preview: "$1,234.00",
    patch: { numberFormat: "$#,##0.00" },
  }),
  freezeStyle({
    id: "currency-0",
    label: "Currency [0]",
    labelZh: "货币 [0]",
    preview: "$1,234",
    patch: { numberFormat: "$#,##0" },
  }),
  freezeStyle({
    id: "percent",
    label: "Percent",
    labelZh: "百分比",
    preview: "12%",
    patch: { numberFormat: "0%" },
  }),
]);

function buildCellStyleGallery(themeId: unknown): ResolvedCellStyleGallery {
  const groups: readonly NamedCellStyleGroup[] = Object.freeze([
    // XLS-A012: Excel's first section is "Good, Bad and Neutral", led by
    // Normal and in Excel's order Normal, Bad, Good, Neutral.
    Object.freeze({
      id: "good-bad-neutral",
      label: "Good, Bad and Neutral",
      labelZh: "好、差和适中",
      styles: Object.freeze([
        NORMAL_STYLE,
        coreStyle("bad"),
        coreStyle("good"),
        coreStyle("neutral"),
      ]),
    }),
    Object.freeze({
      id: "data-model",
      label: "Data and Model",
      labelZh: "数据和模型",
      styles: DATA_MODEL_STYLES,
    }),
    Object.freeze({
      id: "titles-headings",
      label: "Titles and Headings",
      labelZh: "标题",
      styles: Object.freeze([
        coreStyle("title"),
        coreStyle("heading1"),
        coreStyle("heading2"),
        coreStyle("heading3"),
        coreStyle("heading4"),
        coreStyle("total"),
      ]),
    }),
    Object.freeze({
      id: "themed",
      label: "Themed Cell Styles",
      labelZh: "主题单元格样式",
      styles: themedStyles(themeId),
    }),
    Object.freeze({
      id: "number-format",
      label: "Number Format",
      labelZh: "数字格式",
      styles: NUMBER_FORMAT_STYLES,
    }),
  ]);
  return Object.freeze({
    groups,
    styles: Object.freeze(groups.flatMap((group) => group.styles)),
  });
}

const CELL_STYLE_GALLERIES = new Map<string, ResolvedCellStyleGallery>();

/** Resolve one immutable gallery from the workbook's exact live theme. */
export function cellStyleGalleryForTheme(
  themeId: unknown,
): ResolvedCellStyleGallery {
  const normalizedThemeId = paletteById(themeId).id;
  const cached = CELL_STYLE_GALLERIES.get(normalizedThemeId);
  if (cached) return cached;
  const gallery = buildCellStyleGallery(normalizedThemeId);
  CELL_STYLE_GALLERIES.set(normalizedThemeId, gallery);
  return gallery;
}

/** Backward-compatible Office-theme registry for pure/non-React callers. */
const DEFAULT_CELL_STYLE_GALLERY = cellStyleGalleryForTheme("office");
export const CELL_STYLE_GALLERY_GROUPS = DEFAULT_CELL_STYLE_GALLERY.groups;
export const CELL_STYLE_GALLERY_STYLES = DEFAULT_CELL_STYLE_GALLERY.styles;

const NAMED_STYLE_BOOLEAN_FIELDS = new Set([
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "border",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
  "wrapText",
  "shrinkToFit",
  "centerAcross",
  "verticalText",
  "horizontalFill",
  "justifyLastLine",
  "superscript",
  "subscript",
]);

function isClearedStyleValue(value: unknown): boolean {
  return (
    value === undefined || value === null || value === false || value === ""
  );
}

function isHexColor(value: unknown): value is string {
  return (
    typeof value === "string" && /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(value)
  );
}

function validBorderEdgesAuthority(value: unknown): boolean {
  if (
    value === null ||
    typeof value !== "object" ||
    Array.isArray(value) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(value))
  ) {
    return false;
  }
  const edgeMap = value as Record<string, unknown>;
  for (const edge of BORDER_EDGES) {
    if (!Object.prototype.hasOwnProperty.call(edgeMap, edge)) continue;
    const rawEdge = edgeMap[edge];
    if (
      rawEdge === null ||
      typeof rawEdge !== "object" ||
      Array.isArray(rawEdge) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(rawEdge))
    ) {
      return false;
    }
    const candidate = rawEdge as Record<string, unknown>;
    if (
      typeof candidate.style !== "string" ||
      !BORDER_LINE_STYLES.includes(candidate.style as BorderLineStyle) ||
      (candidate.color !== undefined && !isHexColor(candidate.color)) ||
      (candidate.up !== undefined && typeof candidate.up !== "boolean") ||
      (candidate.down !== undefined && typeof candidate.down !== "boolean")
    ) {
      return false;
    }
  }
  return true;
}

function validNamedStyleAuthorityField(key: string, value: unknown): boolean {
  if (key === SOURCE_XF_STYLE_KEY) return true;
  if (isClearedStyleValue(value)) return true;
  if (NAMED_STYLE_BOOLEAN_FIELDS.has(key)) return typeof value === "boolean";
  if (key === "fontSize") {
    return typeof value === "number" && Number.isFinite(value) && value > 0;
  }
  if (key === "indent") {
    return (
      typeof value === "number" && Number.isSafeInteger(value) && value >= 0
    );
  }
  if (key === "textRotation") {
    return (
      typeof value === "number" &&
      Number.isFinite(value) &&
      value >= -90 &&
      value <= 90
    );
  }
  if (
    key === "backgroundColor" ||
    key === "color" ||
    key === "fontFamily" ||
    key === "numberFormat"
  ) {
    return key === "backgroundColor" || key === "color"
      ? isHexColor(value)
      : typeof value === "string";
  }
  if (key === "underlineStyle") {
    return (
      value === "single" ||
      value === "double" ||
      value === "singleAccounting" ||
      value === "doubleAccounting"
    );
  }
  if (key === "readingOrder") return value === "ltr" || value === "rtl";
  if (key === "horizontalJustify" || key === "verticalJustify") {
    return value === "justify" || value === "distributed";
  }
  if (key === "textAlign") {
    return value === "left" || value === "center" || value === "right";
  }
  if (key === "verticalAlign") {
    return value === "top" || value === "middle" || value === "bottom";
  }
  if (key === "borderEdges") return validBorderEdgesAuthority(value);
  if (key === "fillPattern" || key === "gradientFill") {
    return (
      value !== null &&
      typeof value === "object" &&
      !Array.isArray(value) &&
      [Object.prototype, null].includes(Object.getPrototypeOf(value))
    );
  }
  return true;
}

function comparableNamedStyleValue(key: string, value: unknown): unknown {
  if (isClearedStyleValue(value)) return undefined;
  if (
    key === "numberFormat" &&
    typeof value === "string" &&
    value.trim().toLowerCase() === "general"
  ) {
    return undefined;
  }
  if (
    (key === "backgroundColor" || key === "color") &&
    typeof value === "string"
  ) {
    return value.toLowerCase();
  }
  // A custom style's rich border/fill is an object: compare it by content.
  if (value !== null && typeof value === "object") return JSON.stringify(value);
  return value;
}

/**
 * Resolve every gallery tile's semantic state for one live cell style.
 * `null` is a fail-closed malformed authority; unknown style extensions remain
 * preserve-only and do not prevent a built-in style from matching its fields.
 */
export function namedCellStyleMatchStates(
  rawStyle: unknown,
  galleryStyles: readonly NamedCellStyle[] = CELL_STYLE_GALLERY_STYLES,
): Readonly<Record<string, boolean>> | null {
  if (
    rawStyle !== undefined &&
    (rawStyle === null ||
      typeof rawStyle !== "object" ||
      Array.isArray(rawStyle) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(rawStyle)))
  ) {
    return null;
  }
  const style = (rawStyle ?? {}) as Record<string, unknown>;
  for (const [key, value] of Object.entries(style)) {
    if (!validNamedStyleAuthorityField(key, value)) return null;
  }

  const matches: Record<string, boolean> = {};
  for (const named of galleryStyles) {
    matches[named.id] = Object.entries(named.patch).every(
      ([key, expected]) =>
        comparableNamedStyleValue(key, style[key]) ===
        comparableNamedStyleValue(key, expected),
    );
  }
  return matches;
}

/**
 * Resolve a named cell style by id and return a FRESH copy of its patch.
 *
 * - Unknown name → `null` (caller treats as a no-op; never throws).
 * - Known name → a brand-new object each call, so a caller mutating the result
 *   cannot poison the frozen registry (determinism guarantee).
 */
export function applyCellStyle(
  name: string,
  galleryStyles: readonly NamedCellStyle[] = CELL_STYLE_GALLERY_STYLES,
): NamedCellStylePatch | null {
  const style = galleryStyles.find((candidate) => candidate.id === name);
  if (!style) return null;
  return { ...style.patch };
}

/**
 * XLS-A012 — custom cell styles (Home ▸ Cell Styles ▸ New Cell Style… and
 * Merge Styles…). Excel's Style dialog builds a style "by example" from the
 * active cell, one formatting category per check box; a checked category is
 * owned entirely by the style (keys the example leaves unset are cleared when
 * the style is applied), an unchecked one is not touched.
 */
export type CustomCellStyleCategory =
  | "number"
  | "alignment"
  | "font"
  | "border"
  | "fill"
  | "protection";

/** The Style dialog's six "Style Includes" boxes, in Excel's order. */
export const CUSTOM_CELL_STYLE_CATEGORIES: readonly CustomCellStyleCategory[] =
  Object.freeze(["number", "alignment", "font", "border", "fill", "protection"]);

/**
 * Format Cells ▸ Protection as a style carries it. In the cell model these
 * are the cell's own `locked` / `formulaHidden` flags, not `style` keys, so a
 * style that includes Protection holds them beside its patch.
 */
export interface CustomCellStyleProtection {
  locked: boolean;
  hidden: boolean;
}

export const CUSTOM_CELL_STYLE_CATEGORY_KEYS: Readonly<
  Record<CustomCellStyleCategory, readonly string[]>
> = Object.freeze({
  number: Object.freeze(["numberFormat"]),
  alignment: Object.freeze([
    "textAlign",
    "verticalAlign",
    "wrapText",
    "textRotation",
    "verticalText",
    "indent",
    "shrinkToFit",
    "centerAcross",
    "horizontalFill",
    "horizontalJustify",
    "verticalJustify",
    "justifyLastLine",
    "readingOrder",
  ]),
  font: Object.freeze([
    "fontFamily",
    "fontSize",
    "bold",
    "italic",
    "underline",
    "underlineStyle",
    "strikethrough",
    "superscript",
    "subscript",
    "color",
  ]),
  border: Object.freeze([
    "border",
    "borderTop",
    "borderRight",
    "borderBottom",
    "borderLeft",
    "borderEdges",
  ]),
  fill: Object.freeze(["backgroundColor", "fillPattern", "gradientFill"]),
  // Locked / Hidden live on the cell, not in its style (see
  // CustomCellStyleProtection).
  protection: Object.freeze([]),
});

/** The value a key takes when a style owns its category but sets nothing. */
function clearedCustomStyleValue(key: string): unknown {
  return key === "backgroundColor" ||
    key === "color" ||
    key === "fontFamily" ||
    key === "numberFormat"
    ? ""
    : false;
}

export interface CustomCellStyle {
  /** Stable id, `custom-…`; never shown. */
  id: string;
  /** The user's style name (data, never translated). */
  name: string;
  /** The categories the style includes, in Excel's dialog order. */
  categories: CustomCellStyleCategory[];
  patch: NamedCellStylePatch;
  /** Present exactly when `categories` includes "protection". */
  protection?: CustomCellStyleProtection;
}

export const MAX_CUSTOM_CELL_STYLES = 64;
const MAX_CUSTOM_CELL_STYLE_NAME = 255;

/** Excel's style names compare case-insensitively. */
export function customCellStyleNameKey(name: string): string {
  // The server's `name_key` (`casefold`), so a duplicate is one on both sides.
  return spreadsheetNameFold(name.trim());
}

/** The style a New Cell Style… dialog makes from the active cell's format. */
export function customCellStyleFromExample(
  id: string,
  name: string,
  example: Record<string, unknown> | undefined,
  categories: readonly CustomCellStyleCategory[],
  /** The example's Format Cells ▸ Protection (absent: Excel's default, locked). */
  protection?: Partial<CustomCellStyleProtection>,
): CustomCellStyle | null {
  const trimmed = name.trim();
  if (!trimmed || trimmed.length > MAX_CUSTOM_CELL_STYLE_NAME) return null;
  const patch: Record<string, unknown> = {};
  const included = CUSTOM_CELL_STYLE_CATEGORIES.filter((category) =>
    categories.includes(category),
  );
  for (const category of included) {
    for (const key of CUSTOM_CELL_STYLE_CATEGORY_KEYS[category]) {
      const value = example?.[key];
      patch[key] =
        value === undefined || !validNamedStyleAuthorityField(key, value)
          ? clearedCustomStyleValue(key)
          : value;
    }
  }
  return {
    id,
    name: trimmed,
    categories: [...included],
    patch: patch as NamedCellStylePatch,
    ...(included.includes("protection")
      ? {
          protection: {
            locked: protection?.locked !== false,
            hidden: protection?.hidden === true,
          },
        }
      : {}),
  };
}

/**
 * The formatting a Style dialog edits: every category's keys read off a
 * cell style (an example cell, or an existing style's patch), so the
 * check boxes can be turned on and off without losing what Format… set.
 */
export function customCellStyleFormat(
  source: Record<string, unknown> | undefined,
): Record<string, unknown> {
  const format: Record<string, unknown> = {};
  for (const category of CUSTOM_CELL_STYLE_CATEGORIES) {
    for (const key of CUSTOM_CELL_STYLE_CATEGORY_KEYS[category]) {
      const value = source?.[key];
      if (value !== undefined && !isClearedStyleValue(value) && validNamedStyleAuthorityField(key, value)) {
        format[key] = value;
      }
    }
  }
  return format;
}

/** The categories whose keys differ between two Style dialog formats. */
export function customCellStyleChangedCategories(
  before: Record<string, unknown>,
  after: Record<string, unknown>,
): CustomCellStyleCategory[] {
  return CUSTOM_CELL_STYLE_CATEGORIES.filter((category) =>
    CUSTOM_CELL_STYLE_CATEGORY_KEYS[category].some(
      (key) =>
        comparableNamedStyleValue(key, before[key]) !==
        comparableNamedStyleValue(key, after[key]),
    ),
  );
}

/**
 * Excel's Style dialog summary beside a check box: the facts a renderer can
 * spell without translation (format code, face and size); words are the
 * caller's.
 */
export function customCellStyleDescribes(
  format: Record<string, unknown>,
): {
  numberFormat: string | null;
  font: { family: string | null; sizePx: number | null };
  borders: boolean;
  fill: boolean;
} {
  const numberFormat =
    typeof format.numberFormat === "string" && format.numberFormat.trim() &&
    format.numberFormat.trim().toLowerCase() !== "general"
      ? format.numberFormat
      : null;
  return {
    numberFormat,
    font: {
      // The face's own name: a theme font is stored as its fallback stack,
      // and the dialog names the font as Excel's Style dialog does.
      family:
        typeof format.fontFamily === "string" && format.fontFamily
          ? format.fontFamily.split(",")[0]!.trim().replace(/^["']|["']$/g, "") || null
          : null,
      sizePx: typeof format.fontSize === "number" ? format.fontSize : null,
    },
    borders: CUSTOM_CELL_STYLE_CATEGORY_KEYS.border.some((key) => !isClearedStyleValue(format[key])),
    fill: CUSTOM_CELL_STYLE_CATEGORY_KEYS.fill.some((key) => !isClearedStyleValue(format[key])),
  };
}

/**
 * Whether a cell wears a custom style: every key the style owns reads the
 * same on the cell (the rule the gallery's pressed state already uses), and
 * its Locked / Hidden match when the style includes Protection.
 */
export function cellWearsCustomCellStyle(
  cell: { style?: object; locked?: boolean; formulaHidden?: boolean } | undefined,
  style: CustomCellStyle,
): boolean {
  const own = (cell?.style ?? {}) as Record<string, unknown>;
  // A style that says nothing beyond Normal is worn by every plain cell;
  // only a style with a look of its own can be told apart.
  const distinct =
    Object.values(style.patch).some((value) => !isClearedStyleValue(value)) ||
    (style.protection !== undefined &&
      (!style.protection.locked || style.protection.hidden));
  if (!distinct) return false;
  if (
    !Object.entries(style.patch).every(
      ([key, expected]) =>
        comparableNamedStyleValue(key, own[key]) ===
        comparableNamedStyleValue(key, expected),
    )
  ) {
    return false;
  }
  if (style.protection) {
    if ((cell?.locked !== false) !== style.protection.locked) return false;
    if ((cell?.formulaHidden === true) !== style.protection.hidden) return false;
  }
  return style.categories.length > 0;
}

/**
 * Modify… / Delete on a custom style reach the cells that wear it, as in
 * Excel: a modified style restyles them, a deleted one leaves them at Normal
 * (its keys cleared, protection back to locked). `next === null` deletes.
 * Returns the same array when nothing changed.
 */
export function customCellStyleRestyledSheets<
  S extends { cells: Record<string, { style?: object; locked?: boolean; formulaHidden?: boolean }> },
>(sheets: readonly S[], previous: CustomCellStyle, next: CustomCellStyle | null): S[] {
  let changed = false;
  const out = sheets.map((sheet) => {
    let cells: S["cells"] | null = null;
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!cellWearsCustomCellStyle(cell, previous)) continue;
      const style: Record<string, unknown> = { ...((cell.style ?? {}) as Record<string, unknown>) };
      for (const owned of Object.keys(previous.patch)) delete style[owned];
      if (next) {
        for (const [owned, value] of Object.entries(next.patch)) {
          if (!isClearedStyleValue(value)) style[owned] = value;
        }
      }
      const updated: typeof cell = { ...cell };
      if (Object.keys(style).length > 0) updated.style = style;
      else delete updated.style;
      const protection = next ? next.protection : previous.protection ? { locked: true, hidden: false } : undefined;
      if (protection) {
        if (protection.locked) delete updated.locked;
        else updated.locked = false;
        if (protection.hidden) updated.formulaHidden = true;
        else delete updated.formulaHidden;
      }
      cells ??= { ...sheet.cells };
      (cells as Record<string, typeof cell>)[key] = updated;
    }
    if (!cells) return sheet;
    changed = true;
    return { ...sheet, cells };
  });
  return changed ? out : (sheets as S[]);
}

/** Fail-closed reader for the workbook's `customCellStyles`. */
export function normalizeCustomCellStyles(raw: unknown): CustomCellStyle[] {
  if (!Array.isArray(raw)) return [];
  const out: CustomCellStyle[] = [];
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_CUSTOM_CELL_STYLES) break;
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) continue;
    const candidate = entry as Record<string, unknown>;
    const id = candidate.id;
    const name =
      typeof candidate.name === "string" ? candidate.name.trim() : "";
    if (
      typeof id !== "string" ||
      !/^custom-[a-z0-9-]{1,60}$/.test(id) ||
      ids.has(id) ||
      !name ||
      name.length > MAX_CUSTOM_CELL_STYLE_NAME ||
      names.has(customCellStyleNameKey(name))
    ) {
      continue;
    }
    const categories = Array.isArray(candidate.categories)
      ? CUSTOM_CELL_STYLE_CATEGORIES.filter((category) =>
          (candidate.categories as unknown[]).includes(category),
        )
      : [];
    const rawPatch = candidate.patch;
    if (!rawPatch || typeof rawPatch !== "object" || Array.isArray(rawPatch)) {
      continue;
    }
    const allowed = new Set(
      categories.flatMap((category) => CUSTOM_CELL_STYLE_CATEGORY_KEYS[category]),
    );
    const patch: Record<string, unknown> = {};
    let valid = true;
    for (const [key, value] of Object.entries(
      rawPatch as Record<string, unknown>,
    )) {
      if (!allowed.has(key) || !validNamedStyleAuthorityField(key, value)) {
        valid = false;
        break;
      }
      patch[key] = value;
    }
    if (!valid) continue;
    let protection: CustomCellStyleProtection | undefined;
    if (categories.includes("protection")) {
      const raw = candidate.protection as Record<string, unknown> | undefined;
      if (!raw || typeof raw.locked !== "boolean" || typeof raw.hidden !== "boolean") continue;
      protection = { locked: raw.locked, hidden: raw.hidden };
    }
    ids.add(id);
    names.add(customCellStyleNameKey(name));
    out.push({
      id,
      name,
      categories,
      patch: patch as NamedCellStylePatch,
      ...(protection ? { protection } : {}),
    });
  }
  return out;
}

/** The gallery with Excel's "Custom" section first when the workbook has any. */
export function cellStyleGalleryWithCustom(
  gallery: ResolvedCellStyleGallery,
  custom: readonly CustomCellStyle[],
): ResolvedCellStyleGallery {
  if (custom.length === 0) return gallery;
  const styles: NamedCellStyle[] = custom.map((style) =>
    Object.freeze({
      id: style.id,
      label: style.name,
      patch: Object.freeze({ ...style.patch }),
    }),
  );
  const groups: readonly NamedCellStyleGroup[] = Object.freeze([
    Object.freeze({
      id: "custom",
      label: "Custom",
      labelZh: "自定义",
      styles: Object.freeze(styles),
    }),
    ...gallery.groups,
  ]);
  return Object.freeze({
    groups,
    styles: Object.freeze(groups.flatMap((group) => group.styles)),
  });
}

/**
 * Merge Styles…: the source workbook's custom styles into this one. A name
 * both have is replaced only when `replaceSameNames` (Excel's "Merge styles
 * that have same names?" Yes); a replaced style keeps this workbook's id.
 */
export function mergeCustomCellStyles(
  current: readonly CustomCellStyle[],
  incoming: readonly CustomCellStyle[],
  replaceSameNames: boolean,
  makeId: () => string,
): CustomCellStyle[] {
  const out = current.map((style) => ({ ...style }));
  for (const style of incoming) {
    const key = customCellStyleNameKey(style.name);
    const existing = out.findIndex(
      (candidate) => customCellStyleNameKey(candidate.name) === key,
    );
    if (existing >= 0) {
      if (replaceSameNames) out[existing] = { ...style, id: out[existing]!.id };
      continue;
    }
    if (out.length >= MAX_CUSTOM_CELL_STYLES) break;
    out.push({ ...style, id: makeId() });
  }
  return out;
}
