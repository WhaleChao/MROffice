/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Build a styled `text/html` `<table>` for the system clipboard from a copied
 * range, so a paste into Excel / Word / Sheets yields a STYLED table (not just
 * the plain TSV).
 *
 * This is the rich flavor that sits ALONGSIDE the existing `text/plain` TSV
 * (which stays byte-stable — see SpreadsheetWindow `copySelection`). The grid
 * passed in holds, per cell, the on-screen DISPLAY value (formula result /
 * number-format-resolved text), optional canonical A1 formula, and stored
 * `style`, so the HTML carries what the user sees without silently baking a
 * formula into its cached value.
 *
 * Formatting mirrored onto each `<td>` (the attributes the model actually
 * stores and renders in SpreadsheetWindow's `cellStyleCss`):
 *   - bold / italic / underline   → font-weight / font-style / text-decoration
 *   - color / backgroundColor     → font color / fill
 *   - textAlign / verticalAlign   → horizontal / vertical alignment
 *   - borders                     → legacy box + per-side + rich edges, reusing
 *                                   the shared `resolveBorderCss` so the look
 *                                   matches the grid.
 *
 * SECURITY: every cell's display text is HTML-escaped, and every style VALUE
 * the model carries (colors, etc.) is emitted only through a numeric/keyword
 * allowlist or an escaped declaration, so a hostile cell value or stored color
 * string cannot inject markup or break out of the inline `style` attribute.
 *
 * A fully-contained merged range is represented with HTML rowspan/colspan; its
 * covered cells are omitted. A partial merge is intentionally flattened by the
 * caller because a clipboard rectangle cannot carry half of a merged object.
 */

import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import {
  type CellBorderStyle,
  horizontalAlignToCss,
  resolveBorderCss,
  verticalAlignToCss,
} from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import {
  normalizeCellRuns,
  type TextRun,
} from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import { sanitizeHyperlink } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { formulaA1ToR1C1 } from "@mrnote/core/web/lib/spreadsheet/r1c1.ts";
import { cssFontFamilyList } from "@mrnote/core/web/lib/css-font-family.ts";

const MAX_FORMULA_CHARS = 8_192;

/** The subset of a cell's stored style the HTML flavor reads. */
export interface ClipboardCellStyle extends CellBorderStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  fontFamily?: string;
  numberFormat?: string;
  wrapText?: boolean;
  textAlign?: "left" | "center" | "right" | string;
  verticalAlign?: "top" | "middle" | "bottom" | string;
}

/** One copied cell: its display string plus (optionally) its stored style. */
export interface ClipboardCell {
  /** The on-screen value (formula result / number-format-resolved text). */
  display: string;
  /** Canonical A1 formula; emitted only through explicit formula attributes. */
  formula?: string;
  /** HTML-table span metadata for a fully selected merged-cell anchor. */
  rowSpan?: number;
  colSpan?: number;
  /** Covered coordinate of a fully selected merged range; omitted from HTML. */
  covered?: boolean;
  /** Character-level formatting for a literal rich-text cell. */
  runs?: TextRun[];
  /** Safe external hyperlink metadata carried by the copied cell. */
  hyperlink?: string;
  hyperlinkTooltip?: string;
  style?: ClipboardCellStyle;
  /**
   * The cell's exact number when it is one (the display may be 123.46 of
   * 123.456789). Written as Excel's own `x:num` attribute, so a paste that no
   * longer has this page's copy (after a reload, in another tab, in Excel)
   * keeps the value and not the rounded text.
   */
  number?: number;
  /**
   * The cell holds text whatever its characters look like (a quote-prefixed
   * 0007). Written as Excel's `x:str`, so the paste keeps the text.
   */
  text?: boolean;
}

export interface ClipboardHtmlBuildOptions {
  /** Zero-based copied-range anchor used to encode Excel/WPS R1C1 formulas. */
  baseRow?: number;
  baseCol?: number;
}

/** HTML-escape a text node / attribute value (no markup injection). */
function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Whether a stored color string is a safe CSS color to emit inside an inline
 * `style`. Restricts to `#rgb`/`#rrggbb`/`#rrggbbaa` hex and a plain CSS
 * identifier (e.g. `red`, `transparent`) — anything with a quote, semicolon,
 * paren, or whitespace is rejected so it can neither break out of the
 * declaration nor smuggle `url(...)` / `expression(...)`.
 */
function isSafeCssColor(value: string | undefined): value is string {
  if (!value) return false;
  if (/^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.test(value)) {
    return true;
  }
  return /^[a-zA-Z]+$/.test(value);
}

function isSafeCssFontFamily(value: string | undefined): value is string {
  if (!value || value.length > 256) return false;
  return /^[\p{L}\p{N} _,'"-]+$/u.test(value);
}

function safeCssNumberFormat(value: string | undefined): string | undefined {
  if (
    !value ||
    value.length > 128 ||
    /[\u0000-\u001f\u007f]/.test(value)
  ) {
    return undefined;
  }
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

/**
 * A `resolveBorderCss` output value (e.g. "1px solid #475569") is built from
 * literal constants + the model's stored edge {style,color}; allow it through
 * only when it matches the expected `<width> <line-style> <color>` shape (or a
 * bare keyword) so a poisoned stored value cannot inject into the declaration.
 */
function isSafeBorderCss(value: string | undefined): value is string {
  if (!value) return false;
  // width (px) + line-style word + hex/keyword color, in that canonical order.
  return /^[0-9.]+px [a-z]+ (?:#[0-9a-fA-F]{3,8}|[a-zA-Z]+)$/.test(value);
}

function safeFormula(value: string | undefined): string | undefined {
  if (
    !value?.startsWith("=") ||
    value.length > MAX_FORMULA_CHARS ||
    /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value)
  ) {
    return undefined;
  }
  return value;
}

function safeSpan(value: number | undefined, maximum: number): number {
  if (!Number.isInteger(value) || !value || value < 2) return 1;
  return Math.min(maximum, value);
}

/** Build the inline `style` declarations for one copied cell. */
function cellInlineStyle(style: ClipboardCellStyle | undefined): string {
  const decls: string[] = [];
  if (!style) return "";
  if (style.bold) decls.push("font-weight:700");
  if (style.italic) decls.push("font-style:italic");
  const decoration: string[] = [];
  if (style.underline) decoration.push("underline");
  if (style.strikethrough) decoration.push("line-through");
  if (decoration.length > 0) {
    decls.push(`text-decoration:${decoration.join(" ")}`);
  }
  if (isSafeCssColor(style.color)) decls.push(`color:${style.color}`);
  if (isSafeCssColor(style.backgroundColor)) {
    decls.push(`background-color:${style.backgroundColor}`);
  }
  if (
    typeof style.fontSize === "number" &&
    Number.isFinite(style.fontSize) &&
    // Excel's 1-409 pt, in the pixels the model stores (live audit B430).
    style.fontSize >= ptToPx(1) &&
    style.fontSize <= ptToPx(409)
  ) {
    decls.push(`font-size:${style.fontSize}px`);
  }
  if (isSafeCssFontFamily(style.fontFamily)) {
    decls.push(`font-family:${cssFontFamilyList(style.fontFamily)}`);
  }
  if (style.wrapText) decls.push("white-space:normal");
  const numberFormat = safeCssNumberFormat(style.numberFormat);
  if (numberFormat) decls.push(`mso-number-format:"${numberFormat}"`);
  const hAlign = horizontalAlignToCss(style.textAlign);
  if (hAlign) decls.push(`text-align:${hAlign}`);
  const vAlign = verticalAlignToCss(style.verticalAlign);
  if (vAlign) decls.push(`vertical-align:${vAlign}`);
  // Borders: reuse the grid's resolver, then emit only declarations whose value
  // passes the canonical-shape allowlist.
  const border = resolveBorderCss(style as CellBorderStyle);
  if (isSafeBorderCss(border.border)) decls.push(`border:${border.border}`);
  if (isSafeBorderCss(border.borderTop)) {
    decls.push(`border-top:${border.borderTop}`);
  }
  if (isSafeBorderCss(border.borderRight)) {
    decls.push(`border-right:${border.borderRight}`);
  }
  if (isSafeBorderCss(border.borderBottom)) {
    decls.push(`border-bottom:${border.borderBottom}`);
  }
  if (isSafeBorderCss(border.borderLeft)) {
    decls.push(`border-left:${border.borderLeft}`);
  }
  return decls.join(";");
}

function runInlineStyle(run: TextRun): string {
  const style = run.style;
  if (!style) return "";
  const decls: string[] = [];
  if (style.bold) decls.push("font-weight:700");
  if (style.italic) decls.push("font-style:italic");
  const decoration: string[] = [];
  if (style.underline) decoration.push("underline");
  if (style.strike) decoration.push("line-through");
  if (decoration.length > 0) {
    decls.push(`text-decoration:${decoration.join(" ")}`);
  }
  if (isSafeCssColor(style.color)) decls.push(`color:${style.color}`);
  if (
    typeof style.fontSize === "number" &&
    Number.isFinite(style.fontSize) &&
    style.fontSize >= 1 &&
    style.fontSize <= 400
  ) {
    decls.push(`font-size:${style.fontSize}px`);
  }
  if (isSafeCssFontFamily(style.fontFamily)) {
    decls.push(`font-family:${cssFontFamilyList(style.fontFamily)}`);
  }
  return decls.join(";");
}

function escapeTextWithBreaks(value: string): string {
  return escapeHtml(value).replace(/\r\n|\r|\n/g, "<br>");
}

function cellBody(cell: ClipboardCell): string {
  const runs = cell.formula
    ? undefined
    : normalizeCellRuns(cell.runs, cell.display);
  const text = runs
    ? runs
        .map((run) => {
          const content = escapeTextWithBreaks(run.text);
          const inline = runInlineStyle(run);
          return inline
            ? `<span style="${escapeHtml(inline)}">${content}</span>`
            : content;
        })
        .join("")
    : escapeTextWithBreaks(cell.display);
  const hyperlink = sanitizeHyperlink(cell.hyperlink);
  if (!hyperlink) return text;
  const tooltip =
    typeof cell.hyperlinkTooltip === "string" &&
    cell.hyperlinkTooltip.length <= 1_024 &&
    !/[\u0000-\u001f\u007f]/.test(cell.hyperlinkTooltip)
      ? ` title="${escapeHtml(cell.hyperlinkTooltip)}"`
      : "";
  return `<a href="${escapeHtml(hyperlink)}"${tooltip}>${text}</a>`;
}

/**
 * Build a `text/html` `<table>` string for the copied range.
 *
 * Each row is a `<tr>`; each cell a `<td>` carrying its escaped display text
 * and (when present) an inline `style` attribute. The wrapping `<table>` sets
 * `border-collapse:collapse` so per-cell borders render as a grid. An empty
 * grid yields an empty `<table>`.
 */
export function buildClipboardHtml(
  grid: ClipboardCell[][],
  options: ClipboardHtmlBuildOptions = {},
): string {
  const baseRow = Math.max(0, Math.trunc(options.baseRow ?? 0));
  const baseCol = Math.max(0, Math.trunc(options.baseCol ?? 0));
  const carriesFormula = grid.some((row) =>
    row.some(
      (cell) =>
        Boolean(safeFormula(cell.formula)) ||
        Number.isFinite(cell.number) ||
        cell.text === true,
    ),
  );
  const rows = grid
    .map((row, rowIndex) => {
      const cells = row
        .map((cell, colIndex) => {
          if (cell.covered) return "";
          const inline = cellInlineStyle(cell.style);
          const styleAttr = inline ? ` style="${escapeHtml(inline)}"` : "";
          const rowSpan = safeSpan(cell.rowSpan, 1_000);
          const colSpan = safeSpan(cell.colSpan, 200);
          const spanAttr = `${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ""}${
            colSpan > 1 ? ` colspan="${colSpan}"` : ""
          }`;
          const formula = safeFormula(cell.formula);
          const formulaAttr = formula
            ? ` data-formula="${escapeHtml(formula)}" x:fmla="${escapeHtml(
                formulaA1ToR1C1(formula, {
                  row: baseRow + rowIndex,
                  col: baseCol + colIndex,
                }),
              )}"`
            : "";
          // Preserve an in-cell line break (Alt+Enter / wrap-text): a bare "\n"
          // inside a <td> is whitespace-collapsed by Excel/Word/Sheets (which
          // consume the text/html flavor), merging a two-line cell into one. Emit
          // <br> AFTER escaping, exactly as Excel's own HTML clipboard does. The
          // text/plain TSV flavor already quotes embedded newlines (csvEscapeField).
          const body = cellBody(cell);
          const valueAttr =
            typeof cell.number === "number" && Number.isFinite(cell.number)
              ? ` x:num="${Object.is(cell.number, -0) ? "0" : String(cell.number)}"`
              : cell.text === true
                ? " x:str"
                : "";
          return `<td${styleAttr}${spanAttr}${formulaAttr}${valueAttr}>${body}</td>`;
        })
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
  const namespace = carriesFormula
    ? ' xmlns:x="urn:schemas-microsoft-com:office:excel"'
    : "";
  return `<table${namespace} style="border-collapse:collapse">${rows}</table>`;
}
