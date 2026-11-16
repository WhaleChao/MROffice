/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The native cell text effects an imported workbook carries beyond the
 * toolbar's own vocabulary, and how the grid draws them (live audit
 * B397-B400). Each is SpreadsheetML cell format the projection reads and the
 * export writes back; none of them is authored here.
 *
 *  - Reading order (alignment/@readingOrder, Microsoft Learn "Alignment
 *    Class"): 0 Context — the first strong character decides — 1
 *    Left-to-Right, 2 Right-to-Left. Only the explicit two are stored
 *    (`readingOrder: "ltr" | "rtl"`); absence is Context, Excel's default
 *    for every cell. It sets the base direction mixed-script text is laid
 *    out in, never the sheet's column order.
 *  - Fill (Format Cells > Alignment > Horizontal > Fill, `horizontalFill`):
 *    "repeats the contents of the selected cell until the cell is full". The
 *    value is untouched; only the drawing repeats it, and it never runs over
 *    the neighbours.
 *  - Superscript / Subscript (Format Cells > Font > Effects, font/vertAlign):
 *    the whole cell's text raised or lowered off the baseline, smaller. A
 *    single run carries the same pair in its run style.
 *  - Accounting underlines (Format Cells > Font > Underline): Single and
 *    Double Accounting run "the entire width of the cell", below the
 *    descenders, where the plain kinds underline only the characters.
 */
import type { CSSProperties } from "react";

export type SpreadsheetUnderlineKind =
  | "single"
  | "double"
  | "singleAccounting"
  | "doubleAccounting";

type EffectStyle =
  | {
      underline?: unknown;
      underlineStyle?: unknown;
      superscript?: unknown;
      subscript?: unknown;
      readingOrder?: unknown;
      horizontalFill?: unknown;
      centerAcross?: unknown;
      wrapText?: unknown;
    }
  | undefined;

/** The size a raised or lowered script is drawn at, relative to its line —
 * the same scale the header/footer runs use (header-footer-rich-text). */
export const SPREADSHEET_SCRIPT_FONT_SCALE = 0.75;

/** Printed, Windows Excel sets a superscript or subscript at two thirds of
 * its size and raises a superscript's baseline half the size above the
 * line's (lookwin p108-lw: Arial 9 / 12, Calibri 11 / 20 at 6.0 / 8.04 /
 * 7.32 / 13.32 pt, raised 0.49-0.51 em; the API's run_look). */
const PRINTED_SCRIPT_FONT_SCALE = 2 / 3;
const PRINTED_SUPERSCRIPT_RISE = 0.5;
let printedScripts = false;
/** Whether scripts are set as Excel prints them (the print page, with the
 * printer's resolution: native-typography setNativeTextPrinterDpi). */
export function setPrintedCellScripts(on: boolean): void {
  printedScripts = on;
}

/** The underline a cell wears, or null when it has none. */
export function cellUnderlineKind(style: EffectStyle): SpreadsheetUnderlineKind | null {
  if (style?.underline !== true) return null;
  const kind = style.underlineStyle;
  return kind === "double" || kind === "singleAccounting" || kind === "doubleAccounting"
    ? kind
    : "single";
}

/** A two-line underline (Double or Double Accounting). */
export function cellUnderlineIsDouble(style: EffectStyle): boolean {
  const kind = cellUnderlineKind(style);
  return kind === "double" || kind === "doubleAccounting";
}

/**
 * The weight of each line of a Double Underline under the characters. Left
 * to the font ("auto"), a small size drew two hairlines less than a device
 * pixel apart that read as one line at 8.5 pt (video audit S02); Excel's
 * two lines stay a whole pixel each with a pixel between them, and grow with
 * the type.
 */
export const DOUBLE_UNDERLINE_THICKNESS = "max(1px, 0.07em)";

/** The accounting rule the cell's text block carries across the cell, or
 * null for a cell whose underline (if any) is drawn under its characters. */
export function cellAccountingUnderline(style: EffectStyle): "single" | "double" | null {
  const kind = cellUnderlineKind(style);
  return kind === "singleAccounting" ? "single" : kind === "doubleAccounting" ? "double" : null;
}

/**
 * The CSS that turns a single-line text block into one ruled across its
 * whole cell. `padLeft`/`padRight` are the cell's own text insets: the block
 * reaches into them so the rule spans the cell, and pads them back so the
 * characters stay exactly where they were. The rule is the block's bottom
 * edge, under the line box and so under the descenders.
 */
export function accountingUnderlineBlockCss(
  kind: "single" | "double",
  padLeft: number,
  padRight: number,
): CSSProperties {
  return {
    boxSizing: "border-box",
    width: `calc(100% + ${padLeft + padRight}px)`,
    maxWidth: "none",
    marginLeft: -padLeft,
    marginRight: -padRight,
    paddingLeft: padLeft,
    paddingRight: padRight,
    borderBottom: kind === "double" ? "3px double currentColor" : "1px solid currentColor",
  };
}

/**
 * A cell's text box clips its words at the CELL's edge, as Excel's grid
 * does, not at the text inset inside it: the box reaches into the inset by
 * as much as it pads itself back in, so the words stay where they are. An
 * italic last character leans past its own advance (the top right of a
 * right-aligned italic 2 or 7 was shaved off at the inset).
 */
export function cellTextClipAtCellEdgeCss(padLeft: number, padRight: number): CSSProperties {
  return {
    boxSizing: "content-box",
    marginLeft: -padLeft,
    marginRight: -padRight,
    paddingLeft: padLeft,
    paddingRight: padRight,
  };
}

/**
 * The accounting rule of a NUMBER: as wide as the number itself, set below
 * the descenders (XLS-A049). MEASURED in Excel: 1234 with Double Accounting
 * is ruled under its four digits, not across the cell, while text and a
 * format whose `*` fill spans the cell take the whole width (B400).
 */
export function accountingUnderlineInlineCss(kind: "single" | "double"): CSSProperties {
  return {
    boxSizing: "border-box",
    paddingBottom: 1,
    borderBottom: kind === "double" ? "3px double currentColor" : "1px solid currentColor",
  };
}

/** `text-decoration-line` without the underline an accounting rule replaces. */
export function withoutUnderlineDecoration(
  line: CSSProperties["textDecorationLine"],
): CSSProperties["textDecorationLine"] {
  if (typeof line !== "string") return line;
  const rest = line.split(/\s+/).filter((part) => part && part !== "underline");
  return rest.length > 0 ? rest.join(" ") : undefined;
}

/** Whole-cell Superscript / Subscript, or null on the baseline. */
export function cellVerticalScript(style: EffectStyle): "superscript" | "subscript" | null {
  if (style?.superscript === true) return "superscript";
  if (style?.subscript === true) return "subscript";
  return null;
}

/**
 * The CSS for text raised or lowered off its line. `fontSize` is the size the
 * text would have on the baseline; a number scales exactly, anything else
 * (absent, a keyword) scales relative to the line.
 */
export function scriptBaselineCss(
  script: "superscript" | "subscript",
  fontSize?: CSSProperties["fontSize"],
): CSSProperties {
  const numeric = typeof fontSize === "number" && Number.isFinite(fontSize);
  if (printedScripts) {
    return {
      // A length raises the text off its parent's baseline: half the line's
      // size, or three quarters of its own (two thirds of the line's).
      verticalAlign: script === "superscript"
        ? numeric ? `${fontSize * PRINTED_SUPERSCRIPT_RISE}px` : `${PRINTED_SUPERSCRIPT_RISE / PRINTED_SCRIPT_FONT_SCALE}em`
        : "sub",
      fontSize: numeric ? fontSize * PRINTED_SCRIPT_FONT_SCALE : `${PRINTED_SCRIPT_FONT_SCALE * 100}%`,
    };
  }
  return {
    verticalAlign: script === "superscript" ? "super" : "sub",
    fontSize: numeric ? fontSize * SPREADSHEET_SCRIPT_FONT_SCALE : `${SPREADSHEET_SCRIPT_FONT_SCALE * 100}%`,
  };
}

/**
 * The wrapper of a whole-cell script's text (Format Cells > Font >
 * Superscript / Subscript on the cell): smaller, and moved off the line by
 * the canvas offsets ({@link scriptBaselineShiftPx}). `vertical-align` only
 * moves text against the rest of its line; with the whole line scripted the
 * cell's line box grows by the shift and its bottom alignment takes it back,
 * so the text sat on the baseline (X-296).
 */
export function wholeCellScriptCss(
  script: "superscript" | "subscript",
  fontSize?: CSSProperties["fontSize"],
): CSSProperties {
  const px = typeof fontSize === "number" ? fontSize : Number.parseFloat(String(fontSize ?? ""));
  const line = Number.isFinite(px) && px > 0 ? px : 14.6667;
  return {
    ...scriptBaselineCss(script, fontSize),
    verticalAlign: "baseline",
    position: "relative",
    top: scriptBaselineShiftPx(script, line),
  };
}

/**
 * A cell's text parts (its runs, or its one plain part) with the whole-cell
 * script applied, for the surfaces that lay out parts: each part raised or
 * lowered as the plain text is, except a run with a script of its own, whose
 * font says where it sits. The same array comes back on the baseline.
 */
export function partsWithCellScript<P extends { text: string; style?: CSSProperties }>(
  parts: readonly P[],
  script: "superscript" | "subscript" | null,
): readonly P[] {
  if (!script) return parts;
  return parts.map((part) =>
    part.style?.verticalAlign
      ? part
      : { ...part, style: { ...part.style, ...scriptBaselineCss(script, part.style?.fontSize) } },
  );
}

/**
 * How far a raised or lowered script's baseline moves on a canvas, in px of
 * its line's font (`lineFontPx`, the size before the script scale): up for a
 * superscript (negative), down for a subscript — the offsets the browser
 * gives `vertical-align: super / sub`, so a picture puts it where the grid
 * does.
 */
export function scriptBaselineShiftPx(
  script: "superscript" | "subscript",
  lineFontPx: number,
): number {
  if (printedScripts && script === "superscript") return -lineFontPx * PRINTED_SUPERSCRIPT_RISE;
  return script === "superscript" ? -lineFontPx / 3 : lineFontPx / 5;
}

// Strong right-to-left characters: Hebrew, Arabic, Syriac, Thaana, NKo,
// Samaritan, Mandaic and their presentation forms, plus the RLM/ALM marks.
const RTL_STRONG = /[֐-ࣿיִ-﷿ﹰ-ﻼ‏؜]|[\u{10800}-\u{10fff}\u{1e800}-\u{1efff}]/u;

/**
 * The `dir` a cell's text is laid out in: its explicit reading order, or for
 * Context the first strong character (`auto`). Context text with nothing
 * right-to-left in it is left to the grid's own direction, so an ordinary
 * cell's markup does not change.
 */
export function cellTextDirection(
  style: EffectStyle,
  text: string,
): "ltr" | "rtl" | "auto" | undefined {
  if (style?.readingOrder === "ltr" || style?.readingOrder === "rtl") return style.readingOrder;
  return RTL_STRONG.test(text) ? "auto" : undefined;
}

/**
 * The base direction {@link cellTextDirection} resolves to, for a surface
 * that cannot say `auto` (a canvas): Context takes the first strong
 * character's direction, left to right when there is none.
 */
export function cellResolvedTextDirection(style: EffectStyle, text: string): "ltr" | "rtl" {
  const dir = cellTextDirection(style, text);
  if (dir !== "auto") return dir ?? "ltr";
  for (const char of text) {
    if (RTL_STRONG.test(char)) return "rtl";
    if (/\p{L}/u.test(char)) return "ltr";
  }
  return "ltr";
}

/**
 * Where General alignment sets text: at the right edge when its first strong
 * character reads right to left, else the left, whatever the cell's reading
 * order (Windows Excel, lookwin p82-lw: General "كابل" and "كابل Kabul 12"
 * sit right in Context, Left-to-Right and Right-to-Left alike, "Kabul كابل 12"
 * left in all three). The reading order sets only the line's base direction.
 */
export function generalTextAlignment(text: string): "left" | "right" {
  return cellResolvedTextDirection(undefined, text) === "rtl" ? "right" : "left";
}

/** Fill alignment applies: set, and not overridden by Center Across Selection.
 * Fill is a one-line repetition; a cell that also wraps keeps the grid's
 * wrapped-text drawing. */
export function cellHorizontalFill(style: EffectStyle): boolean {
  return style?.horizontalFill === true && style.centerAcross !== true && style.wrapText !== true;
}

/**
 * The text a Fill-aligned cell draws: its contents repeated as many whole
 * times as fit `availablePx` (at least once, clipped by the cell like any
 * other text). `unitWidthPx` is the width of one copy.
 */
export function horizontalFillText(
  text: string,
  unitWidthPx: number,
  availablePx: number,
): string {
  if (!text || !(unitWidthPx > 0) || !Number.isFinite(availablePx)) return text;
  // A pathological width ratio must not build a huge string: a cell is never
  // more than a few thousand pixels wide.
  const copies = Math.min(4096, Math.max(1, Math.floor((availablePx + 0.01) / unitWidthPx)));
  return text.repeat(copies);
}

/**
 * A wrapped line's spaces for `white-space: pre-line`, which collapses them:
 * every space of a run but its last becomes a no-break space, so the run
 * keeps its width and its last space stays the one place the line may break
 * (and, at a break, the one space neither shown nor underlined, as Excel
 * draws it). A run that opens a line is kept whole.
 */
export function spacesKeptForPreLine(text: string): string {
  return text
    .split("\n")
    .map((line) => line.replace(/ +/g, (run: string, offset: number) =>
      offset === 0 ? "\u00a0".repeat(run.length) : `${"\u00a0".repeat(run.length - 1)} `))
    .join("\n");
}
