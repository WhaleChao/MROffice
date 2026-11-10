/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The look a cell position shows, and the style an edit there starts from,
 * with the workbook's default format (cellXfs[0], "xf 0") as a look of its own.
 *
 * A workbook whose xf 0 is centred vertically (common in Japanese and Chinese
 * workbooks) draws every cell that has no format of its own centred; a cell
 * with ANY other format draws what that format says, Excel's built-in default
 * for what it leaves out (xlsx-rounds WALK R269-X13: xf 0 vertical="center").
 * So the fallback is per cell, never per attribute:
 *
 * - a cell holding a style of its own shows that style alone;
 * - a cell that holds nothing as its style (the explicit default `{}`, or no
 *   style at all) is xf 0: the sheet's `defaultStyle`;
 * - a position with no cell shows its row's format, else its column's
 *   (Excel's line formats), else xf 0.
 *
 * An edit adds to the same style the position shows: bold on an xf 0 cell of
 * a centred workbook is centred and bold, as Excel's new xf (xf 0 + bold) is.
 * The font is not part of `defaultStyle` (nativeGrid.font says it).
 */

import { lineStyleAt, type LineStyleMap } from "./line-styles";

/** What a sheet carries for its default format: alignment, fill, borders and
 *  number format of cellXfs[0], read by the importer (reading 496). */
export type DefaultStyle = object;

export interface DefaultLookSheet {
  cells?: Readonly<Record<string, { readonly style?: object | null } | undefined>>;
  rowStyles?: LineStyleMap;
  columnStyles?: LineStyleMap;
  defaultStyle?: DefaultStyle;
}

export interface CellLook {
  /** The style the position is drawn with; undefined draws Excel's built-in default. */
  display: object | undefined;
  /** The style an edit at the position starts from (the same look, as a style to add to). */
  editBase: object | undefined;
  /** Where the look comes from. */
  source: "cell" | "line" | "default" | "none";
}

function nonEmpty(style: unknown): style is object {
  return typeof style === "object" && style !== null && !Array.isArray(style) && Object.keys(style).length > 0;
}

/**
 * The look at (row, col). `useDefault` is the editor's xlsxDefaultLook switch:
 * off, a cell with no style of its own and a position with no line format draw
 * the built-in default as before.
 */
export function cellLookAt(sheet: DefaultLookSheet, row: number, col: number, useDefault = true): CellLook {
  const fallback = useDefault && nonEmpty(sheet.defaultStyle) ? sheet.defaultStyle : undefined;
  const cell = sheet.cells?.[`${row}:${col}`];
  if (cell) {
    if (nonEmpty(cell.style)) return { display: cell.style, editBase: cell.style, source: "cell" };
    return fallback ? { display: fallback, editBase: fallback, source: "default" } : { display: undefined, editBase: undefined, source: "none" };
  }
  const line = lineStyleAt(sheet.rowStyles, sheet.columnStyles, row, col);
  if (nonEmpty(line)) return { display: line, editBase: line, source: "line" };
  return fallback ? { display: fallback, editBase: fallback, source: "default" } : { display: undefined, editBase: undefined, source: "none" };
}

let defaultLookSwitch: boolean | undefined;






export function xlsxDefaultLookEnabled(): boolean {
  if (typeof window === "undefined") return false;
  if (defaultLookSwitch !== undefined) return defaultLookSwitch;
  let asked = new RegExp("[?&]xlsxDefaultLook=([01])").exec(window.location.search)?.[1] ?? null;
  if (asked === null) {
    try { asked = window.localStorage.getItem("mrai.xlsxDefaultLook"); } catch { asked = null; }
  }
  // On unless asked off (`?xlsxDefaultLook=0`, or the stored "0"): the server
  // reads the default format from the same release (XLSX_DEFAULT_LOOK).
  defaultLookSwitch = asked !== "0";
  return defaultLookSwitch;
}

/** A default look a position with no cell shows at all: a fill or a border.
 *  An alignment or a number format shows nothing on an empty position. */
function showsOnEmpty(style: object): boolean {
  const look = style as Record<string, unknown>;
  return Boolean(look.backgroundColor || look.fillPattern || look.fillGradient || look.borderEdges ||
    look.borderTop || look.borderBottom || look.borderLeft || look.borderRight || look.borderDiagonal);
}

const dressedByDefault = new WeakMap<object, WeakMap<object, object>>();
const blankOfDefault = new WeakMap<object, object>();

/**
 * The cell the grid draws at (row, col) given what it would draw without xf 0
 * (`cell`: the sheet's own cell, else its line blank), with xf 0's look on a
 * cell holding no style of its own and, where that look shows on an empty
 * position, on a position with no cell inside the sheet's extent. For drawing
 * only: the objects are shared and never written back.
 */
export function gridCellWithDefaultLook<C extends object>(
  sheet: DefaultLookSheet & { rowCount?: number; columnCount?: number },
  row: number,
  col: number,
  cell: C | undefined,
): C | undefined {
  const look = sheet.defaultStyle;
  if (!nonEmpty(look)) return cell;
  if (cell) {
    if (nonEmpty((cell as { style?: unknown }).style)) return cell;
    let dressed = dressedByDefault.get(look);
    if (!dressed) { dressed = new WeakMap(); dressedByDefault.set(look, dressed); }
    let worn = dressed.get(cell);
    if (!worn) { worn = { ...cell, style: look }; dressed.set(cell, worn); }
    return worn as C;
  }
  if (!showsOnEmpty(look)) return undefined;
  if (row >= (sheet.rowCount ?? 0) || col >= (sheet.columnCount ?? 0)) return undefined;
  let blank = blankOfDefault.get(look);
  if (!blank) { blank = { value: "", kind: "text", style: look }; blankOfDefault.set(look, blank); }
  return blank as C;
}
