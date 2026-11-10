/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The print-settings and workbook-theme writes behind the Page Layout tab
 * (Themes, Margins, Orientation, Size, Print Area, Breaks, Scale to Fit,
 * Sheet Options' Print boxes) and the Page Setup dialog's OK and Clear.
 * Moved out of SpreadsheetWindow so the window and the spreadsheet editor
 * commands (lib/editor-commands/spreadsheet) run the same code; behaviour is
 * unchanged. Each function takes the sheet's stored `printSettings` and
 * answers the settings to store, or null when the button would change
 * nothing; protection, receipts and the dialog itself stay in the window.
 */

import {
  isValidPrintAreaString,
  isValidTitleColRangeString,
  isValidTitleRowRangeString,
  stripRangeAnchors,
  MARGIN_PRESETS,
  MARGIN_PRESET_HEADER_FOOTER,
  normalizeHeaderFooter,
  normalizePrintSettings,
  PAPER_SIZE_NATIVE_IDS,
  resolveMargins,
  resolvePaperSizePoints,
  type MarginPreset,
  type PrintScaling,
  type PrintSettings,
} from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { viewOptionsWithWorkbookTheme, workbookThemeFromViewOptions } from "./theme-buttons";
import type { SpreadsheetViewOptions } from "./view-options";
import {
  normalizeSpreadsheetWorkbookTheme,
  spreadsheetWorkbookThemesEqual,
  type SpreadsheetWorkbookThemeSelection,
} from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

/** The settings a page-layout write stores (`undefined` = none at all). */
export interface PrintSettingsChange {
  settings: PrintSettings | undefined;
}

function changeTo(next: PrintSettings): PrintSettingsChange {
  return { settings: Object.keys(next).length > 0 ? next : undefined };
}

export type PageMarginSelection = MarginPreset | "custom";

export const PAGE_MARGIN_SIDES = ["top", "right", "bottom", "left"] as const;

export const PAGE_SETUP_HEADER_FOOTER_BANDS = [
  "header",
  "footer",
  "evenHeader",
  "evenFooter",
  "firstHeader",
  "firstFooter",
] as const;

/**
 * Read the persisted margin choice without normalizing malformed native data.
 * A malformed authority disables Page Layout margins instead of silently
 * replacing it with a preset on the next click.
 */
export function pageMarginSelectionFromPrintSettings(
  rawSettings: unknown,
): PageMarginSelection | null {
  if (rawSettings === undefined) return "normal";
  if (
    rawSettings === null ||
    typeof rawSettings !== "object" ||
    Array.isArray(rawSettings)
  ) {
    return null;
  }
  const margins = (rawSettings as Record<string, unknown>).margins;
  if (margins === undefined || margins === "normal") return "normal";
  if (margins === "narrow" || margins === "wide") return margins;
  if (
    margins === null ||
    typeof margins !== "object" ||
    Array.isArray(margins)
  ) {
    return null;
  }
  const record = margins as Record<string, unknown>;
  for (const side of PAGE_MARGIN_SIDES) {
    if (!Object.prototype.hasOwnProperty.call(record, side)) return null;
    const value = record[side];
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      value < 0 ||
      value > 288
    ) {
      return null;
    }
  }
  // Excel keeps only numbers, so margins that equal a preset show as it.
  for (const preset of ["normal", "wide", "narrow"] as const) {
    if (
      PAGE_MARGIN_SIDES.every(
        (side) => Math.abs((record[side] as number) - MARGIN_PRESETS[preset][side]) < 0.01,
      )
    ) {
      return preset;
    }
  }
  return "custom";
}

/** Set Print Area / Clear Print Area (and Add to Print Area's result). */
export function printSettingsWithPrintArea(
  rawSettings: unknown,
  printArea: string | null,
): PrintSettingsChange | null {
  const current = normalizePrintSettings(rawSettings) ?? {};
  const next: PrintSettings = { ...current };
  if (printArea) next.printArea = printArea;
  else delete next.printArea;
  const normalized = normalizePrintSettings(next) ?? {};
  if (JSON.stringify(normalized) === JSON.stringify(current)) return null;
  return changeTo(normalized);
}

/** Manual page breaks as one print-settings write, both axes at once. */
export function printSettingsWithManualBreaks(
  rawSettings: unknown,
  rowBreaks: readonly number[],
  colBreaks: readonly number[],
): PrintSettingsChange | null {
  const current = normalizePrintSettings(rawSettings) ?? {};
  const next: PrintSettings = { ...current };
  const normalizedRows = [...new Set(rowBreaks)].sort(
    (left, right) => left - right,
  );
  const normalizedCols = [...new Set(colBreaks)].sort(
    (left, right) => left - right,
  );
  if (normalizedRows.length > 0) next.manualRowBreaks = normalizedRows;
  else delete next.manualRowBreaks;
  if (normalizedCols.length > 0) next.manualColBreaks = normalizedCols;
  else delete next.manualColBreaks;
  const normalized = normalizePrintSettings(next) ?? {};
  if (JSON.stringify(normalized) === JSON.stringify(current)) return null;
  return changeTo(normalized);
}

/**
 * Breaks > Insert Page Break at a cell: a break above its row and one left
 * of its column (none on the first row or column), as Excel does.
 */
export function manualBreaksInsertedAt(
  rawSettings: unknown,
  row: number,
  col: number,
): { rows: number[]; cols: number[] } {
  const current = normalizePrintSettings(rawSettings) ?? {};
  const rows = [...(current.manualRowBreaks ?? [])];
  const cols = [...(current.manualColBreaks ?? [])];
  if (row > 0 && !rows.includes(row)) rows.push(row);
  if (col > 0 && !cols.includes(col)) cols.push(col);
  return { rows, cols };
}

/** Breaks > Remove Page Break at a cell: the breaks on its row and column. */
export function manualBreaksRemovedAt(
  rawSettings: unknown,
  row: number,
  col: number,
): { rows: number[]; cols: number[] } {
  const current = normalizePrintSettings(rawSettings) ?? {};
  return {
    rows: (current.manualRowBreaks ?? []).filter((value) => value !== row),
    cols: (current.manualColBreaks ?? []).filter((value) => value !== col),
  };
}

/** Margins gallery: Normal, Wide or Narrow, with its header/footer distances. */
export function printSettingsWithMarginPreset(
  rawSettings: PrintSettings | undefined,
  preset: MarginPreset,
): PrintSettingsChange | null {
  const current = pageMarginSelectionFromPrintSettings(rawSettings);
  if (current === null) return null;
  // A preset also sets the header and footer distances (Excel); an absent
  // distance already is Normal's 0.3".
  const headerFooter = MARGIN_PRESET_HEADER_FOOTER[preset];
  const nextPrintSettings: PrintSettings = { ...(rawSettings ?? {}) };
  const headerFooterChanges = (["headerMargin", "footerMargin"] as const).filter(
    (key) =>
      (nextPrintSettings[key] ?? MARGIN_PRESET_HEADER_FOOTER.normal) !==
      headerFooter,
  );
  if (current === preset && headerFooterChanges.length === 0) return null;
  for (const key of headerFooterChanges) nextPrintSettings[key] = headerFooter;
  if (preset === "normal") delete nextPrintSettings.margins;
  else nextPrintSettings.margins = preset;
  return changeTo(nextPrintSettings);
}

export type PageSetupQuickSetting = Partial<
  Pick<PrintSettings, "orientation" | "paperSize" | "paperSizeNativeId" | "margins" | "scaling">
>;

/** The ribbon's quick menus (Orientation, Size, Scale): one field each. */
export function printSettingsWithQuickSetting(
  rawSettings: unknown,
  update: PageSetupQuickSetting,
): PrintSettingsChange | null {
  const current = normalizePrintSettings(rawSettings) ?? {};
  const next = normalizePrintSettings({ ...current, ...update }) ?? {};
  if (JSON.stringify(next) === JSON.stringify(current)) return null;
  return changeTo(next);
}

/**
 * Scale to Fit > Width or Height: 0 (Automatic) to 9 pages on one axis; the
 * other axis keeps its value. Null when that axis already reads so.
 */
export function fitPagesScaling(
  rawSettings: unknown,
  axis: "width" | "height",
  pages: number,
): PrintScaling | null {
  if (!Number.isInteger(pages) || pages < 0 || pages > 9) return null;
  const scaling = normalizePrintSettings(rawSettings)?.scaling;
  const widthPages = scaling?.kind === "fitToWidth"
    ? scaling.pages
    : scaling?.kind === "fitToPages"
      ? ("pages" in scaling ? scaling.pages : scaling.widthPages)
      : 0;
  const heightPages = scaling?.kind === "fitToPages"
    ? ("pages" in scaling ? 1 : scaling.heightPages)
    : 0;
  // Automatic is already true in percent mode. Reselecting it must not erase
  // a native scale percentage or silently switch the workbook's print mode.
  if (pages === (axis === "width" ? widthPages : heightPages)) return null;
  return {
    kind: "fitToPages",
    widthPages: axis === "width" ? pages : widthPages,
    heightPages: axis === "height" ? pages : heightPages,
  };
}

/** Scale to Fit > Width or Height as one print-settings write. */
export function printSettingsWithFitPages(
  rawSettings: unknown,
  axis: "width" | "height",
  pages: number,
): PrintSettingsChange | null {
  const scaling = fitPagesScaling(rawSettings, axis, pages);
  return scaling ? printSettingsWithQuickSetting(rawSettings, { scaling }) : null;
}

/** Sheet Options > Gridlines > Print and Headings > Print. */
export function printSettingsWithPrintOptionToggled(
  rawSettings: unknown,
  field: "showGridlines" | "showHeadings",
): PrintSettingsChange {
  const printSettings: PrintSettings = {
    ...(normalizePrintSettings(rawSettings) ?? {}),
  };
  if (printSettings[field] === true) delete printSettings[field];
  else printSettings[field] = true;
  const normalized = normalizePrintSettings(printSettings);
  return { settings: normalized ?? undefined };
}

/** Why the Page Setup dialog refuses its OK (the message the dialog shows). */
export type PageSetupDraftErrorKey =
  | "spreadsheet.pageSetupErrorRange"
  | "spreadsheet.pageSetupErrorRepeatRows"
  | "spreadsheet.pageSetupErrorRepeatCols"
  | "spreadsheet.pageSetupErrorScaling"
  | "spreadsheet.pageSetupDialog.errorMargins"
  | "spreadsheet.pageSetupDialog.errorMarginsDoNotFit"
  | "spreadsheet.pageSetupDialog.errorFirstPageNumber";

/**
 * The Page Setup dialog's OK: the draft written back over the sheet's stored
 * settings as ONE write. A field the person never touched is written back
 * exactly as the sheet holds it; a touched field is validated as Excel's
 * dialog validates it, then stored in its non-default form. `changed` is
 * false when the stored settings would read the same.
 */
export function printSettingsFromPageSetupDraft(
  rawSettings: PrintSettings | undefined,
  draft: PrintSettings,
): { errorKey: PageSetupDraftErrorKey } | (PrintSettingsChange & { changed: boolean }) {
  const sourcePrintSettings = normalizePrintSettings(rawSettings);
  const touched = (field: keyof PrintSettings) =>
    JSON.stringify(draft[field] ?? null) !==
    JSON.stringify(sourcePrintSettings?.[field] ?? null);
  const printArea = stripRangeAnchors(draft.printArea ?? "");
  if (printArea !== "" && !isValidPrintAreaString(printArea)) {
    return { errorKey: "spreadsheet.pageSetupErrorRange" };
  }
  const repeatTitleRows = stripRangeAnchors(draft.repeatTitleRows ?? "");
  if (
    repeatTitleRows !== "" &&
    !isValidTitleRowRangeString(repeatTitleRows)
  ) {
    return { errorKey: "spreadsheet.pageSetupErrorRepeatRows" };
  }
  const repeatTitleCols = stripRangeAnchors(draft.repeatTitleCols ?? "");
  if (
    repeatTitleCols !== "" &&
    !isValidTitleColRangeString(repeatTitleCols)
  ) {
    return { errorKey: "spreadsheet.pageSetupErrorRepeatCols" };
  }
  // Scaling: 10–400 % (Adjust to), or Fit to 0–100 pages per axis where 0
  // is Automatic and at least one axis is bounded.
  const scaling = draft.scaling;
  const scalingInvalid =
    scaling?.kind === "percent"
      ? !Number.isInteger(scaling.percent) ||
        scaling.percent < 10 ||
        scaling.percent > 400
      : scaling?.kind === "fitToWidth"
        ? !Number.isInteger(scaling.pages) ||
          scaling.pages < 1 ||
          scaling.pages > 100
        : scaling?.kind === "fitToPages"
          ? "pages" in scaling
            ? !Number.isInteger(scaling.pages) ||
              scaling.pages < 1 ||
              scaling.pages > 100
            : !Number.isInteger(scaling.widthPages) ||
              !Number.isInteger(scaling.heightPages) ||
              scaling.widthPages < 0 ||
              scaling.widthPages > 100 ||
              scaling.heightPages < 0 ||
              scaling.heightPages > 100 ||
              (scaling.widthPages === 0 && scaling.heightPages === 0)
          : false;
  if (scalingInvalid) {
    return { errorKey: "spreadsheet.pageSetupErrorScaling" };
  }
  // Margins and the header/footer distances: 0–288 points (4 inches).
  const marginValid = (value: unknown) =>
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 288;
  const draftMargins = draft.margins;
  if (
    (typeof draftMargins === "object" &&
      draftMargins !== null &&
      !PAGE_MARGIN_SIDES.every((side) => marginValid(draftMargins[side]))) ||
    (draft.headerMargin !== undefined && !marginValid(draft.headerMargin)) ||
    (draft.footerMargin !== undefined && !marginValid(draft.footerMargin))
  ) {
    return { errorKey: "spreadsheet.pageSetupDialog.errorMargins" };
  }
  // Each margin may be in range while the two of a side leave no paper
  // between them (A6 with 10 cm left and right): Excel refuses that OK with
  // "Margins do not fit page size" rather than printing a sliver (user
  // audit S-017).
  const paper = resolvePaperSizePoints(draft.paperSize);
  const landscape = draft.orientation === "landscape";
  const pageWidth = landscape ? paper.height : paper.width;
  const pageHeight = landscape ? paper.width : paper.height;
  const margins = resolveMargins(draft.margins);
  if (margins.left + margins.right >= pageWidth || margins.top + margins.bottom >= pageHeight) {
    return { errorKey: "spreadsheet.pageSetupDialog.errorMarginsDoNotFit" };
  }
  if (
    draft.firstPageNumber !== undefined &&
    (!Number.isInteger(draft.firstPageNumber) ||
      draft.firstPageNumber < 0 ||
      draft.firstPageNumber > 32_767)
  ) {
    return { errorKey: "spreadsheet.pageSetupDialog.errorFirstPageNumber" };
  }

  // Build the next settings over the RAW stored object: fields the dialog
  // does not own (manual breaks, printer custody, copies) and untouched
  // fields stay byte-identical.
  const next: PrintSettings = { ...(rawSettings ?? {}) };
  const assign = <K extends keyof PrintSettings>(
    field: K,
    value: PrintSettings[K] | undefined,
  ) => {
    if (value === undefined) delete next[field];
    else next[field] = value;
  };
  if (touched("printArea")) {
    assign("printArea", printArea ? printArea.toUpperCase() : undefined);
  }
  if (touched("repeatTitleRows")) {
    assign(
      "repeatTitleRows",
      repeatTitleRows ? repeatTitleRows.toUpperCase() : undefined,
    );
  }
  if (touched("repeatTitleCols")) {
    assign(
      "repeatTitleCols",
      repeatTitleCols ? repeatTitleCols.toUpperCase() : undefined,
    );
  }
  if (touched("orientation")) assign("orientation", draft.orientation);
  if (
    draft.paperSize !== undefined &&
    (touched("paperSize") || touched("paperSizeNativeId"))
  ) {
    const sourcePaper = sourcePrintSettings?.paperSize;
    if (sourcePaper !== undefined && draft.paperSize === sourcePaper) {
      // Back on the stock the sheet came with: keep its exact native
      // identity (a Note/11x17 alias is not the canonical id).
      assign("paperSize", sourcePaper);
      assign("paperSizeNativeId", sourcePrintSettings?.paperSizeNativeId);
    } else if (sourcePaper === undefined && draft.paperSize === "letter") {
      // Picking the cached default again invents no paper identity.
    } else {
      assign("paperSize", draft.paperSize);
      assign("paperSizeNativeId", PAPER_SIZE_NATIVE_IDS[draft.paperSize]);
    }
  }
  if (touched("margins")) {
    if (typeof draftMargins === "string") {
      assign("margins", draftMargins === "normal" ? undefined : draftMargins);
    } else if (draftMargins) {
      // Values that spell one of the gallery presets are stored as the
      // preset, so the Margins gallery shows them as checked again.
      const preset = (Object.keys(MARGIN_PRESETS) as MarginPreset[]).find(
        (name) =>
          PAGE_MARGIN_SIDES.every(
            (side) => MARGIN_PRESETS[name][side] === draftMargins[side],
          ),
      );
      assign(
        "margins",
        preset
          ? preset === "normal"
            ? undefined
            : preset
          : {
              top: draftMargins.top,
              right: draftMargins.right,
              bottom: draftMargins.bottom,
              left: draftMargins.left,
            },
      );
    } else {
      assign("margins", undefined);
    }
  }
  if (touched("scaling")) assign("scaling", scaling);
  for (const band of PAGE_SETUP_HEADER_FOOTER_BANDS) {
    if (!touched(band)) continue;
    // The band as the page keeps it: only non-blank sections, each trimmed
    // with its formatted runs trimmed alongside (both XLSX routes write the
    // runs as Excel's formatting codes), and nothing for an all-empty band.
    assign(band, normalizeHeaderFooter(draft[band]));
  }
  for (const field of [
    "differentOddEven",
    "differentFirstPage",
    "showGridlines",
    "showHeadings",
    "centerHorizontally",
    "centerVertically",
    "blackAndWhite",
    "draft",
  ] as const) {
    if (touched(field)) assign(field, draft[field] === true ? true : undefined);
  }
  for (const field of [
    "headerFooterScaleWithDoc",
    "headerFooterAlignWithMargins",
  ] as const) {
    if (touched(field)) {
      assign(field, draft[field] === false ? false : undefined);
    }
  }
  if (touched("headerMargin")) assign("headerMargin", draft.headerMargin);
  if (touched("footerMargin")) assign("footerMargin", draft.footerMargin);
  if (touched("firstPageNumber")) {
    assign("firstPageNumber", draft.firstPageNumber);
  }
  if (touched("pageOrder")) {
    assign(
      "pageOrder",
      draft.pageOrder === "overThenDown" ? "overThenDown" : undefined,
    );
  }
  if (touched("printErrorsAs")) assign("printErrorsAs", draft.printErrorsAs);
  if (touched("printCellComments")) {
    assign("printCellComments", draft.printCellComments);
  }
  const nextPrintSettings = Object.keys(next).length > 0 ? next : undefined;
  return {
    settings: nextPrintSettings,
    changed:
      JSON.stringify(normalizePrintSettings(rawSettings)) !==
      JSON.stringify(normalizePrintSettings(nextPrintSettings)),
  };
}

/**
 * The Page Setup dialog's Clear: the fields this dialog owns go back to their
 * defaults. Manual breaks and advanced/native print metadata stay; every
 * header/footer band, a formatted one included, is the dialog's and is
 * cleared.
 */
export function printSettingsClearedByPageSetup(
  rawSettings: unknown,
): PrintSettingsChange | null {
  if (!rawSettings) return null;
  const current = normalizePrintSettings(rawSettings) ?? {};
  const next: PrintSettings = { ...current };
  delete next.printArea;
  delete next.repeatTitleRows;
  delete next.repeatTitleCols;
  delete next.orientation;
  delete next.paperSize;
  delete next.paperSizeNativeId;
  delete next.margins;
  delete next.scaling;
  delete next.header;
  delete next.footer;
  // The first-page and even-page bands and their switches are this
  // dialog's too: Clear empties them with the normal pair (B284).
  delete next.firstHeader;
  delete next.firstFooter;
  delete next.evenHeader;
  delete next.evenFooter;
  delete next.differentFirstPage;
  delete next.differentOddEven;
  delete next.showHeadings;
  if (JSON.stringify(next) === JSON.stringify(current)) return null;
  return changeTo(next);
}

/**
 * Themes, Colors, Fonts and Effects: the workbook view options with the
 * requested theme, or null when the workbook already has this one or keeps a
 * theme it cannot change. A theme part MRAI cannot read (preserve-only) keeps
 * its single schemes, but a whole Office theme (Page Layout > Themes,
 * `request.wholeTheme`) replaces it as Excel does, when the workbook names
 * that part unambiguously (`nativeThemeReplaceable`).
 */
export function workbookThemeChange(
  options: SpreadsheetViewOptions,
  requested: SpreadsheetWorkbookThemeSelection,
  request?: { wholeTheme?: boolean },
): {
  previous: SpreadsheetWorkbookThemeSelection;
  next: SpreadsheetWorkbookThemeSelection;
  viewOptions: SpreadsheetViewOptions;
} | null {
  if (
    options.nativeThemeDisposition === "preserve-only" &&
    !(
      request?.wholeTheme === true &&
      options.nativeThemeReplaceable === true &&
      !requested.nativeColors &&
      !requested.nativeFonts &&
      !requested.nativeEffects
    )
  ) {
    return null;
  }
  const next = normalizeSpreadsheetWorkbookTheme(requested);
  const previous = workbookThemeFromViewOptions(options);
  if (spreadsheetWorkbookThemesEqual(previous, next)) return null;
  return { previous, next, viewOptions: viewOptionsWithWorkbookTheme(options, next) };
}
