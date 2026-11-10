/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture placed IN a cell (Excel's Insert > Pictures > Place in Cell).
 *
 * Excel keeps such a picture as the cell's VALUE: it fits the cell keeping its
 * aspect ratio, moves with the cell through sort, filter, insert, delete, copy
 * and fill, and a formula that references the cell returns the picture
 * (Microsoft Support, "Insert a picture with the Place in Cell feature in
 * Excel"; "IMAGE function"). In the file the cell is the error `#VALUE!`
 * (`<c t="e" vm="n"><v>#VALUE!</v></c>`) whose value metadata points at a
 * rich value of type `_localImage`, so an application that does not know
 * in-cell pictures reads `#VALUE!` and one that does reads the picture.
 *
 * The page model keeps exactly that shape: `value` stays the `#VALUE!`
 * fallback and `picture` carries the image. The picture is the cell's value
 * only while that fallback stands -- a cell whose value or formula was
 * replaced (typing, Delete, a paste) no longer holds its picture, whichever
 * writer replaced it, and every reader goes through {@link readCellPicture}.
 * A reader that knows nothing of pictures (CSV, plain text, the server's
 * display cache) therefore says `#VALUE!`, as Excel's own fallback does.
 */

import type { CSSProperties } from "react";

import { printSettingsHeaderFooterPictures } from "./print-settings";
import { PICTURE_OBJECT_TOKEN_PREFIX } from "../page-picture-objects";
import { sanitizeImageSource, sanitizeSheetPictureSource } from "./sheet-image";

/** The characters an in-cell picture's cell stores for readers that cannot draw it. */
export const CELL_PICTURE_FALLBACK_VALUE = "#VALUE!";

/** Alternative text bound, the same 2,000 characters a floating picture's description takes. */
export const MAX_CELL_PICTURE_ALT_LENGTH = 2_000;

export interface SpreadsheetCellPicture {
  /**
   * An embedded raster data URL (see lib/spreadsheet/sheet-image), or, for a
   * picture read from a file past the page's embedded bounds, its token
   * (`mrai-picture:<sha256>`, lib/page-picture-objects).
   */
  src: string;
  /** Excel's Alt Text for the picture in the cell. Absent means none. */
  alt?: string;
}

type CellRecord = Record<string, unknown> & { value?: unknown };

/** The persisted picture shape, validated; null for anything else. */
export function normalizeCellPicture(raw: unknown): SpreadsheetCellPicture | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const record = raw as Record<string, unknown>;
  const src = sanitizeSheetPictureSource(record.src);
  if (src === null) return null;
  const alt =
    typeof record.alt === "string"
      ? record.alt.slice(0, MAX_CELL_PICTURE_ALT_LENGTH)
      : "";
  return alt ? { src, alt } : { src };
}

/**
 * The picture a cell holds as its value, or null. A `picture` field beside a
 * value other than the fallback, or beside a formula, is a stale leftover of
 * a replaced picture and holds nothing.
 *
 * This is the hot-path read (every calculation, every drawn cell), so it
 * checks shape only; the bytes were authenticated when the picture was
 * written ({@link normalizeCellPicture}) and when the page was parsed (the
 * page schema's image-source check).
 */
export function readCellPicture(cell: unknown): SpreadsheetCellPicture | null {
  if (!cell || typeof cell !== "object") return null;
  const record = cell as CellRecord;
  const raw = record.picture;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  if (record.value !== CELL_PICTURE_FALLBACK_VALUE) return null;
  if (typeof record.formula === "string" && record.formula.length > 0) return null;
  const { src, alt } = raw as { src?: unknown; alt?: unknown };
  if (typeof src !== "string" || !(src.startsWith("data:image/") || src.startsWith(PICTURE_OBJECT_TOKEN_PREFIX))) return null;
  return typeof alt === "string" && alt ? { src, alt } : { src };
}

/** Whether the cell's value is a picture. */
export function cellHoldsPicture(cell: unknown): boolean {
  return readCellPicture(cell) !== null;
}

/**
 * The cell with `picture` as its value. Placing a picture overwrites what the
 * cell held (Excel: the pictures "overwrite any values that are currently in
 * those cells") and keeps the cell's own properties: its format, lock, note
 * or comment and hyperlink. The value-shaped fields go: the formula, rich
 * text runs and the caches a value writer derives.
 */
export function placePictureInCell<C extends CellRecord>(
  cell: C | undefined,
  picture: SpreadsheetCellPicture,
): C & { value: string; picture: SpreadsheetCellPicture } {
  const next: CellRecord = { ...(cell ?? {}) };
  for (const key of [
    "formula",
    "runs",
    "display",
    "chartValueCache",
    "hyperlinkDisplay",
  ]) {
    delete next[key];
  }
  if (next.style && typeof next.style === "object" && !Array.isArray(next.style)) {
    const style = { ...(next.style as Record<string, unknown>) };
    // A quote prefix types an entry as text; a picture is not text.
    delete style.quotePrefix;
    next.style = style;
  }
  next.value = CELL_PICTURE_FALLBACK_VALUE;
  // The fallback is stored text, as both importers read an error cell.
  next.kind = "text";
  next.picture = picture.alt
    ? { src: picture.src, alt: picture.alt.slice(0, MAX_CELL_PICTURE_ALT_LENGTH) }
    : { src: picture.src };
  return next as C & { value: string; picture: SpreadsheetCellPicture };
}

/**
 * The cell without any picture field, valid or stale. A writer that replaces
 * a cell's value calls it so a replaced picture does not linger in the page.
 */
export function withoutCellPicture<C>(cell: C): C {
  if (!cell || typeof cell !== "object" || !("picture" in (cell as object))) {
    return cell;
  }
  const next = { ...(cell as Record<string, unknown>) };
  delete next.picture;
  return next as C;
}

/**
 * A picture cell with new Alt Text (empty clears it). The cell is returned
 * unchanged when it holds no picture.
 */
export function withCellPictureAltText<C extends CellRecord>(cell: C, alt: string): C {
  const picture = readCellPicture(cell);
  if (!picture) return cell;
  const text = alt.slice(0, MAX_CELL_PICTURE_ALT_LENGTH);
  return {
    ...cell,
    picture: text ? { src: picture.src, alt: text } : { src: picture.src },
  };
}

/**
 * Where a picture sits in its cell: fitted keeping its aspect ratio (IMAGE's
 * default sizing 0, the one Place in Cell uses), along the cell's own
 * alignment. A General horizontal alignment sets it at the left, as Excel
 * prints and draws it (XlsxWriter XW101: a 32px picture at the left edge of
 * A1, 13pt in a 15pt row; X-309); the vertical default is the cell's
 * Bottom. Returned as the CSS `object-position` a contained image takes.
 */
export function cellPictureObjectPosition(style: unknown): string {
  const record =
    style && typeof style === "object" && !Array.isArray(style)
      ? (style as Record<string, unknown>)
      : {};
  const horizontal =
    record.textAlign === "center" || record.textAlign === "centerContinuous"
      ? "center"
      : record.textAlign === "right"
        ? "right"
        : "left";
  const vertical =
    record.verticalAlign === "top"
      ? "top"
      : record.verticalAlign === "middle"
        ? "center"
        : "bottom";
  return `${horizontal} ${vertical}`;
}

/**
 * A picture a cell DRAWS: its own (Place in Cell), the one a reference brings,
 * or the one an IMAGE formula returns with its sizing (Microsoft Support,
 * "IMAGE function": 0 fit keeping the aspect ratio, 1 fill the cell ignoring
 * it, 2 the picture's own size, which may exceed the cell, 3 the given height
 * and width in pixels, one of them alone keeping the aspect ratio). A placed
 * picture draws as sizing 0.
 */
export interface SpreadsheetShownCellPicture extends SpreadsheetCellPicture {
  sizing?: 0 | 1 | 2 | 3;
  /** IMAGE's height / width arguments, pixels (sizing 3). */
  height?: number;
  width?: number;
  /** The picture's own pixel size (sizing 2 and 3 need it). */
  naturalWidth?: number;
  naturalHeight?: number;
}

/**
 * How the `<img>` of a shown picture is laid in its cell box (the box clips
 * what passes its edges, as a cell does): fitted or filled, or at a pixel
 * size placed along the cell's alignment. `scale` is the drawing's pixels per
 * sheet pixel (a page layout drawn smaller passes its paper scale).
 */
export function cellPictureImageStyle(
  picture: SpreadsheetShownCellPicture,
  cellStyle: unknown,
  scale = 1,
): CSSProperties {
  const position = cellPictureObjectPosition(cellStyle);
  const sizing = picture.sizing ?? 0;
  if (sizing === 0 || sizing === 1) {
    return {
      display: "block",
      width: "100%",
      height: "100%",
      objectFit: sizing === 1 ? "fill" : "contain",
      objectPosition: position,
    };
  }
  const naturalWidth = picture.naturalWidth ?? 0;
  const naturalHeight = picture.naturalHeight ?? 0;
  let width = naturalWidth;
  let height = naturalHeight;
  if (sizing === 3) {
    const ratio = naturalWidth > 0 && naturalHeight > 0 ? naturalWidth / naturalHeight : 1;
    width = picture.width ?? (picture.height !== undefined ? picture.height * ratio : naturalWidth);
    height = picture.height ?? (picture.width !== undefined ? picture.width / ratio : naturalHeight);
  }
  const [horizontal, vertical] = position.split(" ");
  return {
    display: "block",
    position: "absolute",
    width: Math.max(1, width * scale),
    height: Math.max(1, height * scale),
    maxWidth: "none",
    maxHeight: "none",
    objectFit: "fill",
    ...(horizontal === "left"
      ? { left: 0 }
      : horizontal === "right"
        ? { right: 0 }
        : { left: "50%", transform: "translateX(-50%)" }),
    ...(vertical === "top"
      ? { top: 0 }
      : vertical === "center"
        ? { top: "50%", transform: horizontal === "center" ? "translate(-50%, -50%)" : "translateY(-50%)" }
        : { bottom: 0 }),
  };
}

/** Every picture cell of a cell map, as "row:col" key and Alt Text. */
export function cellPictureEntries(
  cells: Readonly<Record<string, unknown>>,
): Array<[string, string]> {
  const entries: Array<[string, string]> = [];
  for (const [key, cell] of Object.entries(cells)) {
    const picture = readCellPicture(cell);
    if (picture) entries.push([key, picture.alt ?? ""]);
  }
  return entries;
}

/**
 * Several pictures placed at once land "starting from the active cell and
 * down the column", one a cell, each overwriting what that cell held
 * (Microsoft Support, Place in Cell). Pictures that would fall past the last
 * row are not placed. Returns the new cell map and the keys written.
 */
export function placePicturesDownColumn<C extends CellRecord>(
  cells: Readonly<Record<string, C>>,
  start: { row: number; col: number },
  pictures: readonly SpreadsheetCellPicture[],
  rowCount: number,
): { cells: Record<string, C>; keys: string[] } {
  const next: Record<string, C> = { ...cells };
  const keys: string[] = [];
  pictures.forEach((picture, index) => {
    const row = start.row + index;
    if (row < 0 || row >= rowCount) return;
    const key = `${row}:${start.col}`;
    next[key] = placePictureInCell(next[key], picture) as unknown as C;
    keys.push(key);
  });
  return { cells: next, keys };
}

/**
 * How many characters of embedded pictures a workbook may carry across its
 * floating pictures, in-cell pictures, sheet backgrounds and header/footer
 * pictures before a new one
 * is refused at insert: the drawing objects' own share of the page
 * (lib/spreadsheet/object-budget), so a picture the editor accepts is one the
 * page can still save.
 */
export const SPREADSHEET_EMBEDDED_PICTURE_CHAR_BUDGET = 8 * 1024 * 1024;

/** The embedded picture characters a workbook already holds. */
export function workbookEmbeddedPictureChars(
  sheets: ReadonlyArray<{
    cells?: Readonly<Record<string, unknown>>;
    objects?: readonly unknown[];
    backgroundPicture?: unknown;
    printSettings?: unknown;
  }>,
): number {
  let total = 0;
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells ?? {})) {
      const picture = readCellPicture(cell);
      if (picture) total += picture.src.length;
    }
    for (const object of sheet.objects ?? []) {
      const src = (object as { src?: unknown } | null)?.src;
      if (typeof src === "string") total += src.length;
    }
    const background = readSheetBackgroundPicture(sheet);
    if (background) total += background.src.length;
    for (const picture of printSettingsHeaderFooterPictures(sheet.printSettings)) {
      total += picture.src.length;
    }
  }
  return total;
}

/**
 * A sheet's background (Excel's Page Layout > Background): one picture the
 * sheet tiles behind its cells on screen and never prints (Microsoft Support,
 * "Add or remove a sheet background").
 */
export interface SpreadsheetSheetBackgroundPicture {
  src: string;
}

/** The sheet's background picture, or null (hot-path read, shape only). */
export function readSheetBackgroundPicture(
  sheet: unknown,
): SpreadsheetSheetBackgroundPicture | null {
  if (!sheet || typeof sheet !== "object") return null;
  const raw = (sheet as { backgroundPicture?: unknown }).backgroundPicture;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const src = (raw as { src?: unknown }).src;
  return typeof src === "string" && src.startsWith("data:image/") ? { src } : null;
}

/** A validated background picture for a sheet write, or null. */
export function normalizeSheetBackgroundPicture(
  raw: unknown,
): SpreadsheetSheetBackgroundPicture | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const src = sanitizeImageSource((raw as { src?: unknown }).src);
  return src ? { src } : null;
}
