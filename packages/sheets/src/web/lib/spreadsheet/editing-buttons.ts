/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Editing and Insert buttons that write cell content: Insert Link,
 * AutoSum, Find & Replace in a selection, Checkbox. Moved out of
 * SpreadsheetWindow so the window and the spreadsheet editor commands run the
 * same code; behaviour is unchanged.
 */

import { recordFontDrawn } from "@mrnote/core/web/lib/spreadsheet/cell-record-font.ts";
import {
  cellKind,
  copyCellMap,
  hasCellMetadata,
  updateCellStyle,
  updateCellValue,
  type CellButtonsSheet,
  type CellMap,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { phoneticAfterValueWrite } from "@mrnote/core/web/lib/spreadsheet/cell-phonetic.ts";
import { canBecomeCheckboxCell, checkboxSeedValue } from "./checkbox";
import { spreadsheetCellIsFormula } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import { sanitizeHyperlink } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { replaceAll as replaceAllCells, type FindReplaceOptions } from "./find-replace";
import { firstReferencedCell, inheritedNumberFormat } from "@mrnote/core/web/lib/spreadsheet/formula-result-format.ts";
import type { AutoSumFn, AutoSumTarget } from "./grid-ops";
import {
  formatInternalLink,
  parseInternalLink,
  resolveInternalLinkTarget,
} from "@mrnote/core/web/lib/spreadsheet/sheet-interactions.ts";

type Rect = { r1: number; c1: number; r2: number; c2: number };

export interface CellHyperlinkDetails {
  displayText?: string;
  tooltip?: string;
}

export function normalizeHyperlinkDetail(raw: string | undefined): string {
  return String(raw ?? "")
    .replace(/[\u0000-\u001f\u007f-\u009f]/g, "")
    .trim()
    .slice(0, 255);
}

/** The Hyperlink slot of a workbook theme (SpreadsheetML `<color theme="10"/>`). */
export const SPREADSHEET_THEME_HYPERLINK_INDEX = 10;

/**
 * Excel's built-in Hyperlink cell style: the theme's Hyperlink colour and a
 * single underline. Inserting a link stamps it the way Excel does, but only
 * the facets the cell has not set itself: a font colour or underline already
 * on the cell stays. The colour stays bound to the theme slot, so a theme
 * switch recolours the link as it recolours Excel's.
 */
export function withHyperlinkCellStyle(
  style: SpreadsheetCell["style"] | undefined,
  hyperlinkColor: string,
): NonNullable<SpreadsheetCell["style"]> {
  const next: NonNullable<SpreadsheetCell["style"]> = { ...(style ?? {}) };
  if (!next.color) {
    next.color = hyperlinkColor;
    next.themeColors = {
      ...(next.themeColors ?? {}),
      color: { theme: SPREADSHEET_THEME_HYPERLINK_INDEX, rgb: hyperlinkColor },
    };
  }
  if (!next.underline) next.underline = true;
  return next;
}

/**
 * Set or clear a cell's hyperlink (Excel/WPS parity). A non-null href is stored
 * verbatim after the caller validates an external URL or canonicalizes an
 * internal workbook target. Optional display/ScreenTip details share the same
 * atomic cell rewrite: plain cells use their value as the visible link text,
 * while a formula remains intact and may carry the XLSX display attribute.
 * A cell that had no link takes the Hyperlink cell style when the caller
 * passes the theme's `hyperlinkColor`; editing an existing link leaves the
 * cell's look alone. Clearing prunes a now-metadata-less, valueless cell so
 * the no-link path stays byte-identical.
 */
export function updateCellHyperlink<S extends CellButtonsSheet>(
  sheet: S,
  row: number,
  col: number,
  href: string | null,
  details?: CellHyperlinkDetails,
  hyperlinkColor?: string,
): S {
  const key = `${row}:${col}`;
  const nextCells = copyCellMap(sheet.cells);
  const previous = nextCells[key] || { value: "", kind: "text" as const };
  const nextCell: SpreadsheetCell = {
    ...previous,
    value: String(previous.value || ""),
    kind: previous.kind || cellKind(String(previous.value || "")),
  };
  if (href) {
    if (!previous.hyperlink && hyperlinkColor) {
      nextCell.style = withHyperlinkCellStyle(previous.style, hyperlinkColor);
    }
    // Changing the target invalidates the imported ScreenTip/display extras
    // (they described the OLD link); an unchanged target keeps them.
    if (nextCell.hyperlink !== href) {
      delete nextCell.hyperlinkTooltip;
      delete nextCell.hyperlinkDisplay;
      
      delete (nextCell as { hyperlinkInArea?: boolean }).hyperlinkInArea;
    }
    nextCell.hyperlink = href;
    if (details) {
      const tooltip = normalizeHyperlinkDetail(details.tooltip);
      if (tooltip) nextCell.hyperlinkTooltip = tooltip;
      else delete nextCell.hyperlinkTooltip;

      // The display text is the cell's own content: only a text the user
      // actually changed is written (the ScreenTip cleaning, its trim and
      // 255-character cut, never touches an unedited value; live audit B136).
      const displayEdited =
        details.displayText !== undefined && details.displayText !== nextCell.value;
      const requestedDisplay = displayEdited
        ? normalizeHyperlinkDetail(details.displayText)
        : nextCell.value;
      // A formula's text is its result; a text entry that starts with "="
      // (quote prefix, Text format) is plain text (B345).
      const hasFormula = spreadsheetCellIsFormula(nextCell);
      if (hasFormula) {
        if (requestedDisplay && requestedDisplay !== nextCell.value) {
          nextCell.hyperlinkDisplay = requestedDisplay;
        } else {
          delete nextCell.hyperlinkDisplay;
        }
      } else {
        const visibleText = requestedDisplay || nextCell.value.trim() || href;
        if (visibleText !== nextCell.value) {
          const phonetic = phoneticAfterValueWrite(
            nextCell.phonetic,
            nextCell.value,
            visibleText,
            { textual: cellKind(visibleText) !== "number" },
          );
          if (phonetic) nextCell.phonetic = phonetic;
          else delete nextCell.phonetic;
          nextCell.value = visibleText;
          nextCell.kind = cellKind(visibleText);
          delete nextCell.display;
          delete nextCell.chartValueCache;
          delete nextCell.runs;
          // The new text takes the cell's own font (cell-record-font).
          if (nextCell.style) nextCell.style = recordFontDrawn(nextCell.style);
        }
        delete nextCell.hyperlinkDisplay;
      }
    }
    nextCells[key] = nextCell;
  } else {
    delete nextCell.hyperlink;
    delete nextCell.hyperlinkTooltip;
    delete nextCell.hyperlinkDisplay;
    delete (nextCell as { hyperlinkInArea?: boolean }).hyperlinkInArea;
    if (nextCell.value || hasCellMetadata(nextCell)) {
      nextCells[key] = nextCell;
    } else {
      delete nextCells[key];
    }
  }
  return { ...sheet, cells: nextCells };
}

/**
 * The href Insert Link stores for what a person typed: an internal
 * "#Sheet!A1" link to a sheet that exists, canonicalized, or a safe external
 * URL (http(s) / mailto). Null when it is neither.
 */
export function hyperlinkTargetFor(trimmed: string, sheetNames: readonly string[]): string | null {
  // Internal link first (so a "#…" value is never mis-sent to the URL
  // sanitizer). Resolve the target so a link to a missing sheet is rejected
  // with the same error as a bad URL; store the canonical form on success.
  if (trimmed.startsWith("#")) {
    const parsed = parseInternalLink(trimmed);
    const resolved = parsed
      ? resolveInternalLinkTarget(trimmed, sheetNames, () => ({ rowCount: 1, columnCount: 1 }))
      : null;
    if (!parsed || !resolved) return null;
    return formatInternalLink(sheetNames[resolved.sheetIndex], parsed.ref);
  }
  return sanitizeHyperlink(trimmed);
}

/**
 * AutoSum's write: each planned target gets its formula, and a General result
 * shows the total the way its first summed cell is formatted (currency stays
 * currency), as in Excel. A count is a count, so Count Numbers stays General.
 */
export function sheetAfterAutoSum<S extends CellButtonsSheet>(
  sheet: S,
  targets: readonly AutoSumTarget[],
  fn: AutoSumFn,
): S {
  let nextSheet = sheet;
  for (const target of targets) {
    nextSheet = updateCellValue(
      nextSheet,
      target.writeRow,
      target.writeCol,
      target.formula,
    );
    const first = fn === "count" ? null : firstReferencedCell(target.formula);
    const inherited = first
      ? inheritedNumberFormat(
          sheet.cells[`${first.row}:${first.col}`]?.style?.numberFormat,
          nextSheet.cells[`${target.writeRow}:${target.writeCol}`]?.style
            ?.numberFormat,
        )
      : null;
    if (inherited) {
      nextSheet = updateCellStyle(
        nextSheet,
        target.writeRow,
        target.writeCol,
        { numberFormat: inherited },
      );
    }
  }
  return nextSheet;
}

/**
 * Replace All over the listed cells only (a selection, or the rows a filter
 * leaves visible): the cells keep everything but the replaced text, and a
 * cell left with nothing is pruned.
 */
export function replaceAllInKeys(
  cells: CellMap,
  keys: readonly string[],
  query: string,
  replacement: string,
  options: FindReplaceOptions,
): { cells: CellMap; count: number } {
  const selectedCells = Object.fromEntries(
    keys.flatMap((key) => {
      const cell = cells[key];
      return cell ? [[key, cell] as const] : [];
    }),
  );
  const result = replaceAllCells(selectedCells, query, replacement, options);
  if (result.count === 0) return { cells, count: 0 };
  const nextCells = copyCellMap(cells);
  for (const key of keys) {
    const nextCell = result.cells[key];
    if (nextCell && (nextCell.value || hasCellMetadata(nextCell as SpreadsheetCell))) {
      nextCells[key] = nextCell as SpreadsheetCell;
    } else {
      delete nextCells[key];
    }
  }
  return { cells: nextCells, count: result.count };
}

/**
 * Insert > Checkbox over a rectangle: a check box IS the values TRUE and
 * FALSE with checkbox formatting, so one is added only where the cell is
 * empty or already holds TRUE/FALSE (Microsoft Support, "Using check boxes in
 * Excel"); a cell with other content keeps it and gets no box.
 */
export function sheetWithCheckboxes<S extends CellButtonsSheet>(
  sheet: S,
  selRect: Rect,
  /** A formula cell's current result. A formula whose result is TRUE or
   * FALSE takes the box as a format and keeps its formula -- the box shows
   * the result, as the same formula with the format pasted already does
   * (live audit S-189); without it a formula cell takes no box. */
  formulaResultAt?: (row: number, col: number) => unknown,
): S {
  let nextSheet = sheet;
  for (let r = selRect.r1; r <= selRect.r2; r += 1) {
    for (let c = selRect.c1; c <= selRect.c2; c += 1) {
      const existing = nextSheet.cells[`${r}:${c}`];
      const formula = typeof existing?.formula === "string" && existing.formula.trim() !== ""
        ? existing.formula
        : typeof existing?.value === "string" && existing.value.trim().startsWith("=") ? existing.value : null;
      if (formula !== null) {
        const result = formulaResultAt?.(r, c);
        const logical = typeof result === "boolean" ||
          (typeof result === "string" && /^(TRUE|FALSE)$/i.test(result));
        if (!logical) continue;
      } else if (!canBecomeCheckboxCell(existing)) continue;
      else nextSheet = updateCellValue(
        nextSheet,
        r,
        c,
        checkboxSeedValue(existing?.value),
      );
      // Set checkbox AND clear the other control flags (mutual exclusivity —
      // updateCellStyle strips the `false` keys), so re-typing a radio/spinner/
      // slider cell as a checkbox swaps the control instead of double-flagging.
      nextSheet = updateCellStyle(nextSheet, r, c, {
        checkbox: true,
        radio: false,
        spinner: false,
        slider: false,
      });
    }
  }
  return nextSheet;
}

/** Remove the checkbox flag (the cells keep their TRUE/FALSE as values). */
export function sheetWithoutCheckboxes<S extends CellButtonsSheet>(sheet: S, selRect: Rect): S {
  let nextSheet = sheet;
  for (let r = selRect.r1; r <= selRect.r2; r += 1) {
    for (let c = selRect.c1; c <= selRect.c2; c += 1) {
      nextSheet = updateCellStyle(nextSheet, r, c, { checkbox: false });
    }
  }
  return nextSheet;
}
