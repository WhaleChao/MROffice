/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Bullets and Numbering libraries and dialog, over a slide's rows
 * (PPT-A039, PPT-A040). Pure: the window commits what these return.
 *
 * Home > Bullets ▾ and Numbering ▾ in PowerPoint for Mac 16.86 (the audit's
 * reference, `dumps/ppt-menu-bullets.txt` / `ppt-menu-numbering.txt`) show
 * None plus seven items each and a "Bullets and Numbering..." entry. The
 * dialog (`ppt-dlg-bullets.txt`) has Bulleted / Numbered pages, Size (25-400 %
 * of text), Color, and on the Numbered page Start at (Microsoft Support, "Add
 * bullets or numbers to text": "Start at" sets the first number).
 *
 * The desktop bullets are Wingdings glyphs, written as `a:buChar` with
 * `a:buFont typeface="Wingdings"` -- the pairs the 1 500-deck corpus scan
 * found as the most frequent non-gallery markers (`ppt-audit/L10/bu_scan.py`:
 * Wingdings ü 1 144, Ø 1 087, § 760, q 303, l 257, n 101). The canvas draws
 * each with Wingdings where the font exists, else its Unicode stand-in
 * (`nativeBulletGlyph`). The filled round bullet is the plain disc, which the
 * gallery model already names.
 */

import type { LiveTextSelection, SlideParagraphWithRuns } from "./presentation-text-buttons";
import {
  materializeNativeBulletReceipt,
  normalizeParagraphBulletFormat,
  type ParagraphBulletFormat,
} from "./presentation-paragraphs";
import type { BulletChar, NumberFormat } from "./office-paragraph-render";

export type PresentationListLibraryItem =
  | {
      readonly id: string;
      readonly kind: "bullet";
      /** A gallery glyph the row model names. */
      readonly glyph?: BulletChar;
      /** Or a symbol-font glyph: the row's own `bulletFormat`. */
      readonly char?: string;
      readonly typeface?: string;
      /** What the tile shows. */
      readonly preview: string;
    }
  | {
      readonly id: string;
      readonly kind: "number";
      readonly format?: NumberFormat;
      readonly autoNum?: string;
      readonly preview: string;
    };

/** Home > Bullets ▾, in PowerPoint's order, after None. */
export const POWERPOINT_BULLET_LIBRARY: readonly PresentationListLibraryItem[] = [
  { id: "large-filled-round", kind: "bullet", char: "l", typeface: "Wingdings", preview: "●" },
  { id: "large-filled-square", kind: "bullet", char: "n", typeface: "Wingdings", preview: "■" },
  { id: "filled-diamond", kind: "bullet", char: "u", typeface: "Wingdings", preview: "◆" },
  { id: "bold-hollow-square", kind: "bullet", char: "q", typeface: "Wingdings", preview: "❑" },
  { id: "checkmark", kind: "bullet", char: "ü", typeface: "Wingdings", preview: "✓" },
  { id: "arrow", kind: "bullet", char: "Ø", typeface: "Wingdings", preview: "➢" },
  { id: "filled-round", kind: "bullet", glyph: "•", preview: "•" },
];

/**
 * Home > Numbering ▾, in PowerPoint's order, after None. The eighth item is
 * the East Asian one a Chinese or Japanese PowerPoint shows ("Ideographic
 * Numbering, wide period"); elsewhere the gallery's i. ii. iii.
 */
export function powerPointNumberLibrary(locale: string): readonly PresentationListLibraryItem[] {
  const eastAsian = /^(zh|ja)\b/i.test(locale);
  return [
    { id: "decimal", kind: "number", format: "decimal", preview: "1. 2. 3." },
    { id: "circled", kind: "number", autoNum: "circleNumDbPlain", preview: "① ② ③" },
    { id: "upper-roman", kind: "number", format: "upper-roman", preview: "I. II. III." },
    { id: "upper-alpha", kind: "number", format: "upper-alpha", preview: "A. B. C." },
    { id: "lower-alpha-paren", kind: "number", format: "lower-alpha-paren", preview: "a) b) c)" },
    { id: "lower-alpha", kind: "number", format: "lower-alpha", preview: "a. b. c." },
    eastAsian
      ? { id: "ideographic", kind: "number", autoNum: "ea1JpnChsDbPeriod", preview: "一． 二． 三．" }
      : { id: "lower-roman", kind: "number", format: "lower-roman", preview: "i. ii. iii." },
  ];
}

const IDENTITY_KEYS = ["char", "typeface", "autoNum"] as const;

function withoutIdentity(format: ParagraphBulletFormat | undefined): ParagraphBulletFormat {
  const kept: ParagraphBulletFormat = { ...(format ?? {}) };
  for (const key of IDENTITY_KEYS) delete kept[key];
  return kept;
}

/** The library item a row draws, or undefined when it is none of them. */
export function listLibraryItemOfRow(
  row: Pick<SlideParagraphWithRuns, "bullet" | "bulletChar" | "numberFormat" | "bulletFormat">,
  library: readonly PresentationListLibraryItem[],
  block: { bulletChar?: string; numberFormat?: string },
): string | undefined {
  const format = row.bulletFormat;
  return library.find((item) => {
    if (item.kind !== row.bullet) return false;
    if (item.kind === "bullet") {
      if (item.char !== undefined) {
        return format?.char === item.char && format.typeface === item.typeface;
      }
      const glyph = row.bulletChar ?? block.bulletChar ?? "•";
      return format?.char === undefined && glyph === item.glyph;
    }
    if (item.autoNum !== undefined) return format?.autoNum === item.autoNum;
    const numberFormat = row.numberFormat ?? block.numberFormat ?? "decimal";
    return format?.autoNum === undefined && numberFormat === item.format;
  })?.id;
}

/** The rows a ribbon command addresses, as the paragraph commands find them. */
export type RowSpan = readonly [first: number, last: number];

/**
 * A library item on the selected rows: they become that list kind, drawing
 * that marker. A symbol-font glyph or scheme is the row's own
 * `bulletFormat`; a gallery one its `bulletChar` / `numberFormat`. The row's
 * colour, size and Start at stay, as PowerPoint keeps them.
 */
export function rowsWithListLibraryItem(
  rows: readonly SlideParagraphWithRuns[],
  [first, last]: RowSpan,
  item: PresentationListLibraryItem,
): SlideParagraphWithRuns[] {
  return rows.map((row, index) => {
    if (index < first || index > last) return row;
    const next: SlideParagraphWithRuns = { ...row, bullet: item.kind };
    delete next.bulletChar;
    delete next.numberFormat;
    const kept = withoutIdentity(normalizeParagraphBulletFormat(row.bulletFormat, item.kind));
    if (item.kind === "bullet") {
      if (item.char !== undefined) {
        next.bulletFormat = { ...kept, char: item.char, typeface: item.typeface };
        return next;
      }
      if (item.glyph !== undefined && item.glyph !== "•") next.bulletChar = item.glyph;
    } else if (item.autoNum !== undefined) {
      next.bulletFormat = { ...kept, autoNum: item.autoNum };
      return next;
    } else if (item.format !== undefined && item.format !== "decimal") {
      next.numberFormat = item.format;
    }
    // `{}` is a statement: this row draws the gallery's marker now, not the
    // symbol-font one its receipt or format named.
    if (row.bulletFormat !== undefined || Object.keys(kept).length > 0) next.bulletFormat = kept;
    else delete next.bulletFormat;
    return next;
  });
}

// --- Bullets and Numbering dialog -------------------------------------------

export const BULLET_SIZE_PERCENT_BOUNDS = [25, 400] as const;
export const BULLET_START_AT_BOUNDS = [1, 32_767] as const;

/** What the dialog opens with: the first selected row's marker. */
export interface BulletsNumberingSource {
  readonly kind: "bullet" | "number";
  readonly itemId?: string;
  readonly sizePercent: number;
  readonly color?: string;
  readonly startAt: number;
}

export function bulletsNumberingSource(
  row: SlideParagraphWithRuns | undefined,
  bulletLibrary: readonly PresentationListLibraryItem[],
  numberLibrary: readonly PresentationListLibraryItem[],
  block: { bulletChar?: string; numberFormat?: string },
): BulletsNumberingSource {
  const kind = row?.bullet === "number" ? "number" : "bullet";
  const format = row?.bulletFormat;
  return {
    kind,
    itemId:
      row && row.bullet !== "none"
        ? listLibraryItemOfRow(row, kind === "number" ? numberLibrary : bulletLibrary, block)
        : undefined,
    sizePercent: format?.sizePct !== undefined ? Math.round(format.sizePct / 1000) : 100,
    color: format?.color,
    startAt: format?.startAt ?? 1,
  };
}

export interface BulletsNumberingResult {
  readonly item: PresentationListLibraryItem | null;
  readonly sizePercent: number;
  /** `null` is Automatic: the text's colour. */
  readonly color: string | null;
  readonly startAt: number;
}

export type BulletsNumberingProblem = "size" | "startAt";

/** Validate the dialog's typed values. */
export function bulletsNumberingResult(draft: {
  item: PresentationListLibraryItem | null;
  size: string;
  color: string | null;
  startAt: string;
}): BulletsNumberingResult | { problem: BulletsNumberingProblem } {
  const size = Number(draft.size.trim());
  if (
    !Number.isInteger(size) ||
    size < BULLET_SIZE_PERCENT_BOUNDS[0] ||
    size > BULLET_SIZE_PERCENT_BOUNDS[1]
  ) {
    return { problem: "size" };
  }
  const startAt = Number(draft.startAt.trim() || "1");
  if (
    !Number.isInteger(startAt) ||
    startAt < BULLET_START_AT_BOUNDS[0] ||
    startAt > BULLET_START_AT_BOUNDS[1]
  ) {
    return { problem: "startAt" };
  }
  return { item: draft.item, sizePercent: size, color: draft.color, startAt };
}

/**
 * OK on the selected rows: the picked item (None clears the list), then the
 * marker's size, colour and, for numbers, Start at -- written as the row's
 * own `bulletFormat`, 100 % / Automatic / 1 stated by omission.
 */
export function rowsWithBulletsNumbering(
  rows: readonly SlideParagraphWithRuns[],
  span: RowSpan,
  result: BulletsNumberingResult,
): SlideParagraphWithRuns[] {
  if (result.item === null) {
    return rows.map((row, index) => {
      if (index < span[0] || index > span[1] || row.bullet === "none") return row;
      const next: SlideParagraphWithRuns = { ...row, bullet: "none" };
      delete next.bulletChar;
      delete next.numberFormat;
      delete next.bulletFormat;
      return next;
    });
  }
  const listed = rowsWithListLibraryItem(rows, span, result.item);
  return listed.map((row, index) => {
    if (index < span[0] || index > span[1]) return row;
    const format: ParagraphBulletFormat = { ...(row.bulletFormat ?? {}) };
    delete format.sizePts;
    if (result.sizePercent === 100) delete format.sizePct;
    else format.sizePct = result.sizePercent * 1000;
    if (result.color === null) delete format.color;
    else format.color = result.color.toLowerCase();
    if (row.bullet === "number" && result.startAt !== 1) format.startAt = result.startAt;
    else delete format.startAt;
    const next: SlideParagraphWithRuns = { ...row };
    if (Object.keys(format).length > 0 || row.bulletFormat !== undefined) next.bulletFormat = format;
    return next;
  });
}

/** Hand a block's receipt to its rows before a list command edits them. */
export { materializeNativeBulletReceipt };

export type { LiveTextSelection };
